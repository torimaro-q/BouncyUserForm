VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "glExplosion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Implements ICFormPhysicsEf
Private Const PARTICLECOUNT As Long = 10000
Private pD(PARTICLECOUNT) As Vector2d
Private pV(PARTICLECOUNT) As Vector2d
Private pC(PARTICLECOUNT) As Color4
Private LF(PARTICLECOUNT) As Double
Private InvLF(PARTICLECOUNT) As Double
Private effectCX As Double
Private effectCY As Double
Private myGL As OpenGL, Life As Double, delay As Double, dd As Double
Private Sub ICFormPhysicsEf_Render(x As Double, y As Double, dt As Long, v As Double)
    delay = delay - dt
    If delay < 0 Then
        Life = Life - dt
        If Life < 0 Then Exit Sub
        Call Update(dt)
        Dim i As Long
        With myGL
            .Disable GL_LIGHTING
            .Enable GL_BLEND
            .BlendFunc GL_SRC_ALPHA, GL_ONE_MINUS_SRC_ALPHA
            .PointSize 5#
            .EnableClientState GL_VERTEX_ARRAY
            .EnableClientState GL_COLOR_ARRAY
            .ColorPointer 4, GL_FLOAT, 0, VarPtr(pC(0))
            .VertexPointer 2, GL_DOUBLE, 0, VarPtr(pD(0))
            .DrawArrays GL_POINTS, 0, UBound(pD) + 1
            .DisableClientState GL_VERTEX_ARRAY
            .DisableClientState GL_COLOR_ARRAY
            .Disable GL_BLEND
        End With
    End If
End Sub
Private Sub Update(ByRef dt)
    Dim i As Long, alpha As Single, t As Single
    For i = 0 To PARTICLECOUNT
        With pV(i)
            .x = .x * (1 - dt * 0.005)
            .y = .y * (1 - dt * 0.005)
        End With
        With pD(i)
            .x = .x + pV(i).x * dt
            .y = .y + pV(i).y * dt
        End With
        LF(i) = LF(i) - dt
        alpha = LF(i) * InvLF(i)
        t = 1 - alpha
        If alpha >= 0 Then
            Dim R As Single, G As Single, B As Single
            If t < 0.2 Then
                R = 1
                G = 0.95
                B = 0.8
            ElseIf t < 0.5 Then
                R = 1
                G = 0.7 - (t - 0.2) * 1#
                B = 0.2
            ElseIf t < 0.8 Then
                R = 0.8 - (t - 0.5) * 1#
                G = 0.2
                B = 0.1
            Else
                R = 0.2
                G = 0.2
                B = 0.2
            End If
            With pC(i)
                .R = R
                .G = G
                .B = B
                .a = alpha
            End With
        End If
    Next i
End Sub
Private Sub ICFormPhysicsEf_Reset(cx As Double, cy As Double, ddmg As Double, Optional hw As Double = 0#, Optional hh As Double = 0#)
    Dim i As Long, rv As Double, th As Double
    dd = ddmg
    effectCX = cx
    effectCY = cy
    Life = 0
    delay = 200
    For i = 0 To PARTICLECOUNT
        th = Rnd * 6.28
        rv = 3 * ddmg * (Rnd) ^ 3
        LF(i) = Rnd * 500 + 500
        If Life < LF(i) Then Life = LF(i)
        InvLF(i) = 1 / LF(i)
        With pD(i)
            .x = cx
            .y = cy
        End With
        With pV(i)
            .x = rv * FastCos(th)
            .y = rv * FastSin(th)
        End With
        With pC(i)
            .R = Rnd
            .G = Rnd
            .B = Rnd
            .a = 1
        End With
    Next i
End Sub
Private Property Get ICFormPhysicsEf_CreateInstance() As ICFormPhysicsEf
    Set ICFormPhysicsEf_CreateInstance = New glExplosion
End Property
Private Sub ICFormPhysicsEf_Init(targetGL As OpenGL)
    Set myGL = targetGL
End Sub

