VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "glMoveTrail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const TRAIL_COUNT As Long = 10
Private pos(TRAIL_COUNT) As Vector2d
Private col(TRAIL_COUNT) As Color4
Private Const TInv As Double = 1 / TRAIL_COUNT
Private idx As Long
Private myGL As OpenGL
Public Sub Render(ByRef x, ByRef y, ByRef dt)
    Call Update(x, y, dt)
    With myGL
        .Disable GL_LIGHTING
        .Enable GL_BLEND
        .BlendFunc GL_SRC_ALPHA, GL_ONE_MINUS_SRC_ALPHA
        .EnableClientState GL_VERTEX_ARRAY
        .EnableClientState GL_COLOR_ARRAY
        
        .PointSize 200
        .ColorPointer 4, GL_FLOAT, 0, VarPtr(col(0))
        .VertexPointer 2, GL_DOUBLE, 0, VarPtr(pos(0))
        .DrawArrays GL_POINTS, 0, TRAIL_COUNT + 1
        
        .DisableClientState GL_VERTEX_ARRAY
        .DisableClientState GL_COLOR_ARRAY
        .Disable GL_BLEND
    End With
End Sub
Public Sub Update(ByVal x As Double, ByVal y As Double, ByVal dt As Long)
    idx = idx + 1
    If idx > TRAIL_COUNT Then idx = 0
    pos(idx).x = x + (Rnd - 0.5) * 8
    pos(idx).y = y + (Rnd - 0.5) * 8
    Dim i As Long, t As Double
    For i = 0 To TRAIL_COUNT
        t = i * TInv
        With col((idx - i + TRAIL_COUNT) Mod TRAIL_COUNT)
            .r = (1 - t) * 1# + t * 0#
            .g = (1 - t) * 1# + t * 0.3
            .b = 1#
            .a = Exp(-i * 0.02)
        End With
    Next i
End Sub
Public Sub Init(ByRef targetGL As OpenGL)
    Set myGL = targetGL
End Sub
Public Property Get CreateInstance() As glMoveTrail
    Set CreateInstance = New glMoveTrail
End Property

