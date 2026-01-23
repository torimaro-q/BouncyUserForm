VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "glShockWave"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const PARTICLECOUNT As Long = 5000
Private pD(PARTICLECOUNT) As Vector2d
Private pV(PARTICLECOUNT) As Vector2d
Private pC(PARTICLECOUNT) As Color4
Private LF(PARTICLECOUNT) As Double
Private InvLF(PARTICLECOUNT) As Double
Private effectCX As Double
Private effectCY As Double
Private myGL As OpenGL, Life As Double, dd As Double
Public Sub Render(ByRef x, ByRef y, ByRef dt)
    Life = Life - dt
    If Life < 0 Then Exit Sub
    Call Update(dt)
    Dim i As Long
    With myGL
        .Disable GL_LIGHTING
        .Enable GL_BLEND
        .LineWidth 5 + dd * 3
        .BlendFunc GL_SRC_ALPHA, GL_ONE_MINUS_SRC_ALPHA
        .EnableClientState GL_VERTEX_ARRAY
        .EnableClientState GL_COLOR_ARRAY
        
        .ColorPointer 4, GL_FLOAT, 0, VarPtr(pC(0))
        .VertexPointer 2, GL_DOUBLE, 0, VarPtr(pD(0))
        
        .DrawArrays GL_LINE_LOOP, 0, UBound(pD) + 1
        
        .DisableClientState GL_VERTEX_ARRAY
        .DisableClientState GL_COLOR_ARRAY
        .Disable GL_BLEND
    End With
End Sub
Private Sub Update(ByRef dt)
    Dim i As Long, alpha As Single, t As Single
    For i = 0 To PARTICLECOUNT
        With pV(i)
            .x = .x * (1 - dt * 0.0002)
            .y = .y * (1 - dt * 0.0002)
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
            .a = alpha
        End With
        End If
    Next i
End Sub
Public Sub Reset(ByRef cx As Double, cy As Double, ByRef ddmg As Double)
    Dim i As Long, rv As Double, th As Double, dth As Double
    effectCX = cx
    effectCY = cy
    Life = 400
    th = 0
    dth = 6.28 / PARTICLECOUNT
    For i = 0 To PARTICLECOUNT
        th = th + dth
        rv = ddmg + ddmg * Rnd * 0.3
        LF(i) = Life
        InvLF(i) = 1 / Life
        With pD(i)
            .x = cx
            .y = cy
        End With
        With pV(i)
            .x = rv * myGL.FastCos(th)
            .y = rv * myGL.FastSin(th)
        End With
        With pC(i)
            .r = 0.9 + Rnd * 0.1
            .g = 0.9 + Rnd * 0.1
            .b = 0.9 + Rnd * 0.1
            .a = 1
        End With
    Next i
End Sub
Public Sub Init(ByRef targetGL As OpenGL)
    Set myGL = targetGL
End Sub
Public Property Get CreateInstance() As glShockWave
    Set CreateInstance = New glShockWave
End Property

