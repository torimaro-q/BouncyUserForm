VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "glControlShatter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Implements ICFormPhysicsEf
Private Const PARTICLECOUNT As Long = 100
Private pD(PARTICLECOUNT - 1) As Vector2d
Private pV(PARTICLECOUNT - 1) As Vector2d
Private pC(PARTICLECOUNT * 3 - 1) As Color4
Private pTri(PARTICLECOUNT - 1, 2) As Vector2d
Private verts(PARTICLECOUNT * 3 - 1) As Vector2d
Private lineIdx(PARTICLECOUNT * 6 - 1) As Long
Private pAngle(PARTICLECOUNT - 1) As Double
Private pAngVel(PARTICLECOUNT - 1) As Double
Private myGL As OpenGL
Private life As Double
Private delay As Double
Private Sub ICFormPhysicsEf_Render(X As Double, Y As Double, dt As Long, v As Double)
    life = life - dt
    If life < 0 Then Exit Sub
    Call Update(dt)
    With myGL
        .Disable GL_LIGHTING
        .Enable GL_BLEND
            .BlendFunc GL_SRC_ALPHA, GL_ONE_MINUS_SRC_ALPHA
            .EnableClientState GL_VERTEX_ARRAY
            .VertexPointer 2, GL_DOUBLE, 0, VarPtr(verts(0))
            .Color4f 1, 1, 1, 1
            .LineWidth 1
            .DrawElements GL_LINES, PARTICLECOUNT * 6, GL_UNSIGNED_INT, VarPtr(lineIdx(0))
            .EnableClientState GL_COLOR_ARRAY
            .ColorPointer 4, GL_FLOAT, 0, VarPtr(pC(0))
            .DrawArrays GL_TRIANGLES, 0, PARTICLECOUNT * 3
            .DisableClientState GL_COLOR_ARRAY
            .DisableClientState GL_VERTEX_ARRAY
        .Disable GL_BLEND
    End With
End Sub
Private Sub Update(ByRef dt As Long)
    Dim i As Long, v As Long
    Dim ca As Double, SA As Double
    Dim lx As Double, ly As Double
    Dim rx As Double, ry As Double
    For i = 0 To PARTICLECOUNT - 1
        With pV(i)
            .X = .X * (1 - dt * 0.001)
            .Y = .Y * (1 - dt * 0.001)
        End With
        With pD(i)
            .X = .X + pV(i).X * dt
            .Y = .Y + pV(i).Y * dt
            pAngle(i) = pAngle(i) + pAngVel(i) * dt
            If pAngle(i) > 6.28 Then
                pAngle(i) = pAngle(i) - 6.28
            ElseIf pAngle(i) < 0 Then
                pAngle(i) = pAngle(i) + 6.28
            End If
            ca = FastCos(pAngle(i))
            SA = FastSin(pAngle(i))
            For v = 0 To 2
                lx = pTri(i, v).X
                ly = pTri(i, v).Y
                rx = lx * ca - ly * SA
                ry = lx * SA + ly * ca
                verts(3 * i + v).X = .X + rx
                verts(3 * i + v).Y = .Y + ry
            Next v
        End With
    Next i
End Sub
Private Sub ICFormPhysicsEf_Reset(cx As Double, cy As Double, ddmg As Double, Optional hw As Double = 0#, Optional hh As Double = 0#)
    Dim i As Long
    Dim th As Double, rv As Double
    Dim a0 As Double, a1 As Double, a2 As Double
    Dim S As Double, base As Long
    Dim size As Double
    size = 2 + hh * hw * 0.01
    life = 300 + hh * hw
    For i = 0 To PARTICLECOUNT - 1
        With pD(i)
            .X = cx
            .Y = cy
        End With
        th = Rnd * 6.28
        rv = 0.1 + (Rnd) ^ 2
        With pV(i)
            .X = rv * FastCos(th)
            .Y = rv * FastSin(th)
        End With
        S = size + Rnd * size
        a0 = Rnd * 6.28
        a1 = a0 + 1 + Rnd
        a2 = a1 + 1 + Rnd
        With pTri(i, 0)
            .X = S * FastCos(a0)
            .Y = S * FastSin(a0)
        End With
        With pTri(i, 1)
            .X = S * FastCos(a1)
            .Y = S * FastSin(a1)
        End With
        With pTri(i, 2)
            .X = S * FastCos(a2)
            .Y = S * FastSin(a2)
        End With
        pAngle(i) = Rnd * 6.28
        pAngVel(i) = (Rnd - 0.5) * 0.01
        base = i * 3
        With pC(base)
            .R = 0.7
            .G = 0.7
            .B = 1
            .A = 1
        End With
        pC(base + 1) = pC(base)
        pC(base + 2) = pC(base)
    Next i
End Sub
Private Property Get ICFormPhysicsEf_CreateInstance() As ICFormPhysicsEf
    Set ICFormPhysicsEf_CreateInstance = New glControlShatter
End Property
Private Sub ICFormPhysicsEf_Init(targetGL As OpenGL)
    Set myGL = targetGL
    Dim t As Long
    Dim i As Long
    Dim v As Long
    For t = 0 To PARTICLECOUNT - 1
        i = t * 6
        v = t * 3
        lineIdx(i + 0) = v + 0
        lineIdx(i + 1) = v + 1
        lineIdx(i + 2) = v + 1
        lineIdx(i + 3) = v + 2
        lineIdx(i + 4) = v + 2
        lineIdx(i + 5) = v + 0
    Next t
End Sub
