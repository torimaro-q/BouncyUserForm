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
Private Const DMAX As Double = 100
Private Const InvDMAX As Double = 1 / DMAX
Private Const OFS As Double = 50
Private Const IVMAX As Double = 1 / VMAX
Private pos(ELM_COUNT - 1) As Vector2d
Private col(ELM_COUNT - 1) As Color4
Private Mask(3) As Vector2d
Private maskcl(3) As Color4
Private Const TInv As Double = 1 / ELM_COUNT
Private Const Pinv As Double = 6.28 / ELM_COUNT
Private myGL As OpenGL
Private ofsx As Double, ofsy As Double, idx As Long
Private sdt As Long, PX As Double, py As Double, cp4(3) As Vector4d, dmg As Double, dR As Double, dB As Double
Private Sub ICFormPhysicsEf_Render(X As Double, Y As Double, dt As Long, v As Double)
    sdt = sdt + dt
    If sdt > 100 Then
        Call Update(X, Y, dt, v)
        PX = X: py = Y: sdt = 0
    End If
    If Abs(X - PX) > 50 Or Abs(Y - py) > 50 Then
        Call Update(X, Y, dt, v)
        PX = X: py = Y: sdt = 0
    End If
    With Mask(0)
        .X = X - ofsx: .Y = Y - ofsy - OFS
    End With
    With Mask(1)
        .X = X + ofsx: .Y = Y - ofsy - OFS
    End With
    With Mask(2)
        .X = X + ofsx: .Y = Y + ofsy - OFS
    End With
    With Mask(3)
        .X = X - ofsx: .Y = Y + ofsy - OFS
    End With
    With myGL
        .PushMatrix
        .Disable GL_LIGHTING
        .EnableClientState GL_VERTEX_ARRAY
            .EnableClientState GL_COLOR_ARRAY
            .VertexPointer 2, GL_DOUBLE, 0, VarPtr(Mask(0))
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
        .PopMatrix
    End With
End Sub
Private Sub Update(ByVal X As Double, ByVal Y As Double, ByVal dt As Long, ByVal v As Double)
    If idx <= 0 Then: idx = ELM_COUNT
    idx = (idx - 4)
    Dim i As Long, j As Long, k As Long, t As Double
    With pos(idx)
        .X = X - ofsx: .Y = Y - ofsy - OFS
    End With
    With pos(idx + 1)
        .X = X + ofsx: .Y = Y - ofsy - OFS
    End With
    With pos(idx + 2)
        .X = X + ofsx: .Y = Y + ofsy - OFS
    End With
    With pos(idx + 3)
        .X = X - ofsx: .Y = Y + ofsy - OFS
    End With
    For i = 0 To UBound(col) Step 4
        k = (idx + i) Mod ELM_COUNT
        For j = 0 To 3
            With col(k + j)
                .R = dR
                .G = 0
                .B = dB
                .A = Exp(-i * 0.1)
            End With
        Next j
    Next i
End Sub
Private Sub ICFormPhysicsEf_Reset(cx As Double, cy As Double, ddmg As Double, Optional hw As Double = 0#, Optional hh As Double = 0#)
    dmg = dmg + ddmg
    ofsy = hh
    ofsx = hw
    Dim td As Single
    td = dmg * InvDMAX
    If td > 1 Then td = 1
    dR = td
    dB = 1 - td
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
            .A = 1
        End With
    Next i
    idx = ELM_COUNT
End Sub

