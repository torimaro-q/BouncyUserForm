VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "glStatusVisualizer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Implements ICFormPhysicsEf
Private chr1() As Byte, ccnt1 As Long, hp As Double, dmg As Double
Private chr2() As Byte, ccnt2 As Long, v2 As Long
Private Const HPCF As Double = 5
Private Const LBOFSX As Double = 180
Private Const LBOFSY As Double = 30
Private myGL As OpenGL
Private Sub ICFormPhysicsEf_Render(x As Double, y As Double, dt As Long, v As Double)
    With myGL
        .PushMatrix
            .Disable GL_LIGHTING
            .listbase 2000
            .Color4f 1, 1, 1, 1
            .RasterPos2d 30, 55
            .CallLists ccnt1, GL_UNSIGNED_BYTE, VarPtr(chr1(0))
            .Color4f 1#, 0#, 0#, 1
            .Begin GL_QUADS
                .Vertex2d LBOFSX, LBOFSY
                .Vertex2d LBOFSX + HPCF * hp, LBOFSY
                .Vertex2d LBOFSX + HPCF * hp, LBOFSY + 5
                .Vertex2d LBOFSX, LBOFSY + 5
            .End1
            .Color4f 1, 1, 1, 1
            .Begin GL_QUADS
                .Vertex2d LBOFSX, LBOFSY
                .Vertex2d LBOFSX + HPCF * 100, LBOFSY
                .Vertex2d LBOFSX + HPCF * 100, LBOFSY + 5
                .Vertex2d LBOFSX, LBOFSY + 5
            .End1
            v2 = (v \ 10) * 10
            chr2 = StrConv("Velocity : " & format(v2, "0"), vbFromUnicode)
            ccnt2 = UBound(chr2) + 1
            
            .listbase 1000
            .Color4f 1, 1, 1, 1
            .RasterPos2d 30, 95
            .CallLists ccnt2, GL_UNSIGNED_BYTE, VarPtr(chr2(0))
            
            .Color4f 0, 0, 0, 1
            .Begin GL_QUADS
                .Vertex2d 15, 0
                .Vertex2d 700, 0
                .Vertex2d 700, 100
                .Vertex2d 15, 100
            .End1
        .PopMatrix
    End With
End Sub
Private Sub Update(ByVal x As Double, ByVal y As Double, ByVal dt As Long, ByVal v As Double)
End Sub
Private Sub ICFormPhysicsEf_Reset(cx As Double, cy As Double, ddmg As Double, Optional hw As Double = 0#, Optional hh As Double = 0#)
    dmg = dmg + ddmg
    hp = 100 - dmg
    If hp < 0 Then hp = 0
End Sub
Private Property Get ICFormPhysicsEf_CreateInstance() As ICFormPhysicsEf
    Set ICFormPhysicsEf_CreateInstance = New glStatusVisualizer
End Property
Private Sub ICFormPhysicsEf_Init(targetGL As OpenGL)
    Set myGL = targetGL
    hp = 100
    chr1 = StrConv("Form", vbFromUnicode)
    ccnt1 = UBound(chr1) + 1
End Sub

