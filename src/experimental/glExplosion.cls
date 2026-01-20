VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "glExplosion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Const PARTICLECOUNT As Long = 3000
Private pD(PARTICLECOUNT) As Vector2d
Private pV(PARTICLECOUNT) As Vector2d
Private pC(PARTICLECOUNT) As Color4
Private LF(PARTICLECOUNT) As Double
Private InvLF(PARTICLECOUNT) As Double
Private effectCX As Double
Private effectCY As Double
Private myGL As OpenGL, Life As Double
Public Sub Render(ByRef x, ByRef y, ByRef dt)
    Life = Life - dt
    If Life < 0 Then Exit Sub
    Call update(dt)
    Dim i As Long
    With myGL
        .Disable GL_LIGHTING
        .Enable GL_BLEND
        .BlendFunc GL_SRC_ALPHA, GL_ONE_MINUS_SRC_ALPHA
        .PointSize 6#
        .EnableClientState GL_VERTEX_ARRAY
        .EnableClientState GL_COLOR_ARRAY
        .ColorPointer 4, GL_FLOAT, 0, VarPtr(pC(0))
        .VertexPointer 2, GL_DOUBLE, 0, VarPtr(pD(0))
        .DrawArrays GL_POINTS, 0, UBound(pD) + 1
        .DisableClientState GL_VERTEX_ARRAY
        .DisableClientState GL_COLOR_ARRAY
        .Disable GL_BLEND
    End With
End Sub
Private Sub update(ByRef dt)
    Dim i As Long, alpha As Single, t As Single
    For i = 0 To PARTICLECOUNT
        With pV(i)
            .x = .x * (1 - dt * 0.002)
            .y = .y * (1 - dt * 0.002)
        End With
        With pD(i)
            .x = .x + pV(i).x * dt
            .y = .y + pV(i).y * dt
        End With
        LF(i) = LF(i) - dt
        alpha = LF(i) * InvLF(i)
        t = 1 - alpha
        If alpha >= 0 Then
        With pC(i)
            .r = myGL.FastSin(t * 6.28) * 0.5 + 0.5
            .G = myGL.FastSin(t * 6.28 + 2.09) * 0.5 + 0.5
            .B = myGL.FastSin(t * 6.28 + 4.18) * 0.5 + 0.5
            .a = alpha
        End With
        End If
    Next i
End Sub
Public Sub Init(ByRef targetGL As OpenGL)
    Set myGL = targetGL
End Sub
Public Sub Reset(ByRef cx As Double, cy As Double)
    Dim i As Long, rv As Double, th As Double
    effectCX = cx
    effectCY = cy
    Life = 0
    For i = 0 To PARTICLECOUNT
        th = Rnd * 6.28
        rv = Rnd * 1.5
        LF(i) = Rnd * 500 + 500
        If Life < LF(i) Then Life = LF(i)
        InvLF(i) = 1 / LF(i)
        With pD(i)
            .x = cx
            .y = cy
        End With
        With pV(i)
            .x = rv * myGL.FastCos(th)
            .y = rv * myGL.FastSin(th)
        End With
        With pC(i)
            .r = Rnd
            .G = Rnd
            .B = Rnd
            .a = 1
        End With
    Next i
End Sub
Public Sub Terminate()
    Call Class_Terminate
End Sub
Private Sub Class_Terminate()
    Set myGL = Nothing
End Sub
