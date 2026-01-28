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
Implements ICFormPhysicsEf
Private Const TRAIL_COUNT As Long = 5
Private Const ELM_COUNT As Long = TRAIL_COUNT * 4
Private Const VMAX As Double = 200
Private Const OFS As Double = 20
Private Const IVMAX As Double = 1 / VMAX
Private pos(ELM_COUNT - 1) As Vector2d
Private col(ELM_COUNT - 1) As Color4
Private mask(3) As Vector2d
Private maskcl(3) As Color4
Private Const TInv As Double = 1 / ELM_COUNT
Private Const Pinv As Double = 6.28 / ELM_COUNT
Private idx As Long
Private myGL As OpenGL
Private ofsx As Double, ofsy As Double
Private sdt As Long, px As Double, py As Double, cp4(3) As Vector4d
Private Sub ICFormPhysicsEf_Render(x As Double, y As Double, dt As Long, v As Double)
    sdt = sdt + dt
    If sdt > 100 Then
        Call Update(x, y, dt, v)
        px = x: py = y: sdt = 0
    End If
    If Abs(x - px) > 50 Or Abs(y - py) > 50 Then
        Call Update(x, y, dt, v)
        px = x: py = y: sdt = 0
    End If
    With mask(0)
        .x = x - ofsx
        .y = y - ofsy - OFS
    End With
    With mask(1)
        .x = x + ofsx
        .y = y - ofsy - OFS
    End With
    With mask(2)
        .x = x + ofsx
        .y = y + ofsy - OFS
    End With
    With mask(3)
        .x = x - ofsx
        .y = y + ofsy - OFS
    End With
    With myGL
        .Disable GL_LIGHTING
        .EnableClientState GL_VERTEX_ARRAY
            .EnableClientState GL_COLOR_ARRAY
            .VertexPointer 2, GL_DOUBLE, 0, VarPtr(mask(0))
            .ColorPointer 4, GL_FLOAT, 0, VarPtr(maskcl(0))
            .DrawArrays GL_QUADS, 0, 4
                .Enable GL_BLEND
                .BlendFunc GL_SRC_ALPHA, GL_ONE_MINUS_SRC_ALPHA
                    .VertexPointer 2, GL_DOUBLE, 0, VarPtr(pos(0))
                    .ColorPointer 4, GL_FLOAT, 0, VarPtr(col(0))
                    .DrawArrays GL_QUADS, idx, ELM_COUNT - idx
                    If idx >= 4 Then .DrawArrays GL_QUADS, 0, idx
                .Disable GL_BLEND
            .DisableClientState GL_COLOR_ARRAY
        .DisableClientState GL_VERTEX_ARRAY
    End With
End Sub
Private Sub Update(ByVal x As Double, ByVal y As Double, ByVal dt As Long, ByVal v As Double)
    If idx <= 0 Then: idx = ELM_COUNT
    idx = (idx - 4)
    Dim i As Long, j As Long, k As Long, t As Double
    With pos(idx)
        .x = x - ofsx * 0.9
        .y = y - ofsy * 0.9
    End With
    With pos(idx + 1)
        .x = x + ofsx * 0.9
        .y = y - ofsy * 0.9
    End With
    With pos(idx + 2)
        .x = x + ofsx * 0.9
        .y = y + ofsy * 0.9
    End With
    With pos(idx + 3)
        .x = x - ofsx * 0.9
        .y = y + ofsy * 0.9
    End With
    For i = 0 To UBound(col) Step 4
        k = (idx + i) Mod ELM_COUNT
        For j = 0 To 3
            With col(k + j)
                .R = 0
                .G = 0
                .B = 100
                .a = Exp(-i * 0.1)
            End With
        Next j
    Next i
End Sub
Private Sub ICFormPhysicsEf_Reset(cx As Double, cy As Double, ddmg As Double, Optional hw As Double = 0#, Optional hh As Double = 0#)
    ofsy = hh
    ofsx = hw
    idx = ELM_COUNT
End Sub
Private Property Get ICFormPhysicsEf_CreateInstance() As ICFormPhysicsEf
    Set ICFormPhysicsEf_CreateInstance = New glMoveTrail
End Property
Private Sub ICFormPhysicsEf_Init(targetGL As OpenGL)
    Set myGL = targetGL
    Dim i As Long
    For i = 0 To 3
        With maskcl(i)
            .R = 254 / 255
            .G = 254 / 255
            .B = 254 / 255
            .a = 1
        End With
    Next i
End Sub

