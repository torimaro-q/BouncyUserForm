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
Private LBOFSX As Double
Private Const LBOFSY As Double = 65
Private Const MAX_HP As Double = 150
Private centerx As Double
Private myGL As OpenGL
Private Sub ICFormPhysicsEf_Render(X As Double, Y As Double, dt As Long, v As Double)
    With myGL
        .PushMatrix
            .listbase 0
            .Disable GL_LIGHTING
            .Color4f 1#, 0#, 0#, 1
            .Begin GL_QUADS
                .Vertex2d LBOFSX, LBOFSY
                .Vertex2d LBOFSX + HPCF * hp, LBOFSY
                .Vertex2d LBOFSX + HPCF * hp, LBOFSY + 7
                .Vertex2d LBOFSX, LBOFSY + 7
            .End1
            .CallList 10
        .PopMatrix
    End With
End Sub
Private Sub Update(ByVal X As Double, ByVal Y As Double, ByVal dt As Long, ByVal v As Double)
End Sub
Private Sub ICFormPhysicsEf_Reset(cx As Double, cy As Double, ddmg As Double, Optional hw As Double = 0#, Optional hh As Double = 0#)
    dmg = dmg + ddmg
    hp = MAX_HP - dmg
    If hp < 0 Then hp = 0
End Sub
Private Property Get ICFormPhysicsEf_CreateInstance() As ICFormPhysicsEf
    Set ICFormPhysicsEf_CreateInstance = New glStatusVisualizer
End Property
Private Sub ICFormPhysicsEf_Init(targetGL As OpenGL)
    Set myGL = targetGL
    hp = MAX_HP
    chr1 = StrConv(myGL.Param("FormName"), vbFromUnicode)
    ccnt1 = UBound(chr1) + 1
    Dim CapX As Double
    With myGL
        centerx = .Param("width") * 0.5
        LBOFSX = centerx - MAX_HP * HPCF * 0.5
        CapX = .Param("width") * 0.5 - ccnt1 * 16.5
        .listbase 0
        .NewList 10, GL_COMPILE
        .listbase 2000
            .PushMatrix
                .Enable GL_BLEND
                .BlendFunc GL_SRC_ALPHA, GL_ONE_MINUS_SRC_ALPHA
                    .Color4f 1, 1, 1, 1
                    .RasterPos2d CapX, 55
                    .CallLists ccnt1, GL_UNSIGNED_BYTE, VarPtr(chr1(0))
                    .Color4f 0#, 0#, 0#, 0.9
                    .RasterPos2d CapX + 1, 56#
                    .CallLists ccnt1, GL_UNSIGNED_BYTE, VarPtr(chr1(0))
                    .RasterPos2d CapX - 1, 56#
                    .CallLists ccnt1, GL_UNSIGNED_BYTE, VarPtr(chr1(0))
                    .RasterPos2d CapX + 1, 54#
                    .CallLists ccnt1, GL_UNSIGNED_BYTE, VarPtr(chr1(0))
                    .RasterPos2d CapX - 1, 54#
                    .CallLists ccnt1, GL_UNSIGNED_BYTE, VarPtr(chr1(0))
                    .Color4f 0.5, 0.5, 0.5, 1
                    .Color4f 0, 0, 0, 1
                    .Begin GL_QUADS
                        .Vertex2d LBOFSX, LBOFSY
                        .Vertex2d LBOFSX + HPCF * MAX_HP, LBOFSY
                        .Vertex2d LBOFSX + HPCF * MAX_HP, LBOFSY + 7
                        .Vertex2d LBOFSX, LBOFSY + 7
                    .End1
                .Disable GL_BLEND
            .PopMatrix
        .EndList
    End With
    
End Sub


