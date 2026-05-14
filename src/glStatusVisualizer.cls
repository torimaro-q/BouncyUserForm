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
Private Const LBOFSY As Double = 80
Private Const MAX_HP As Double = 150
Private centerx As Double
Private myGL As OpenGL
Private lastHp As Double
Private WLen As Double, wLf As Double
Private Const WHITE_LIFE_MAX As Double = 1000
Private Const WHITE_LIFE_MAX_INV As Double = 1 / WHITE_LIFE_MAX
Private Sub ICFormPhysicsEf_Render(X As Double, Y As Double, dt As Long, v As Double)
    Call Update(X, Y, dt, v)
    Dim tHP As Double: tHP = hp
    If tHP < 0 Then tHP = 0
    If tHP > MAX_HP Then tHP = MAX_HP
    Dim RR As Double
    RR = LBOFSX + HPCF * tHP
    With myGL
        .PushMatrix
            .listbase 0
            .Disable GL_LIGHTING
            If WLen > 0 Then
                Dim XR As Double: XR = RR + HPCF * WLen
                If XR > LBOFSX + HPCF * MAX_HP Then XR = LBOFSX + HPCF * MAX_HP
                .Color4f 1#, 1#, 1#, 1#
                .Begin GL_QUADS
                    .Vertex2d RR, LBOFSY
                    .Vertex2d XR, LBOFSY
                    .Vertex2d XR, LBOFSY + 7
                    .Vertex2d RR, LBOFSY + 7
                .End1
            End If
            .Color4f 1#, 0#, 0#, 1
            .Begin GL_QUADS
                .Vertex2d LBOFSX, LBOFSY
                .Vertex2d LBOFSX + HPCF * tHP, LBOFSY
                .Vertex2d LBOFSX + HPCF * tHP, LBOFSY + 7
                .Vertex2d LBOFSX, LBOFSY + 7
            .End1
            .CallList 10
        .PopMatrix
    End With
End Sub
Private Sub Update(ByVal X As Double, ByVal Y As Double, ByVal dt As Long, ByVal v As Double)
    If WLen <= 0 Then GoTo continue
    If wLf <= 0 Then GoTo continue
    wLf = wLf - dt
    If wLf <= 0 Then
        WLen = 0
    Else
        WLen = WLen * wLf * WHITE_LIFE_MAX_INV
    End If
continue:
    lastHp = hp
End Sub
Private Sub ICFormPhysicsEf_Reset(cx As Double, cy As Double, ddmg As Double, Optional hw As Double = 0#, Optional hh As Double = 0#)
    Dim pHP As Double: pHP = hp
    dmg = dmg + ddmg
    hp = MAX_HP - dmg
    Dim lost As Double
    lost = pHP - hp
    If lost > 0 Then
        WLen = WLen + lost
        wLf = WHITE_LIFE_MAX
    End If
    lastHp = hp
End Sub
Private Property Get ICFormPhysicsEf_CreateInstance() As ICFormPhysicsEf
    Set ICFormPhysicsEf_CreateInstance = New glStatusVisualizer
End Property
Private Sub ICFormPhysicsEf_Init(targetGL As OpenGL)
    Dim CapX1 As Double, CapX2 As Double
    Dim CapY1 As Double, CapY2 As Double
    Dim h As Long, S As Long, v As Long
    Set myGL = targetGL
    hp = MAX_HP
    lastHp = hp
    WLen = 0
    wLf = 0
    chr1 = StrConv(myGL.Param("FormName"), vbFromUnicode)
    Call RGB2HSV(CLng(myGL.Param("color")), h, S, v)
    chr2 = StrConv("- " & GetSubName(h, S, v) & " -", vbFromUnicode)
    ccnt1 = UBound(chr1) + 1
    ccnt2 = UBound(chr2) + 1
    CapY1 = 70
    CapY2 = 1
    With myGL
        centerx = .Param("width") * 0.5
        LBOFSX = centerx - MAX_HP * HPCF * 0.5
        CapX1 = .Param("width") * 0.5 - ccnt1 * 17.6
        CapX2 = .Param("width") * 0.5 - ccnt2 * 8.2
        .listbase 0
        .NewList 10, GL_COMPILE
            .PushMatrix
                .Enable GL_BLEND
                .BlendFunc GL_SRC_ALPHA, GL_ONE_MINUS_SRC_ALPHA
                    .Color4f 1, 1, 1, 1
                    .listbase 1000
                    .RasterPos2d CapX2, CapY2
                    .CallLists ccnt2, GL_UNSIGNED_BYTE, VarPtr(chr2(0))
                    .Color4f 0#, 0#, 0#, 0.9
                    .RasterPos2d CapX2 + 1, CapY2 + 1
                    .CallLists ccnt2, GL_UNSIGNED_BYTE, VarPtr(chr2(0))
                    .RasterPos2d CapX2 - 1, CapY2 + 1
                    .CallLists ccnt2, GL_UNSIGNED_BYTE, VarPtr(chr2(0))
                    .RasterPos2d CapX2 + 1, CapY2 - 1
                    .CallLists ccnt2, GL_UNSIGNED_BYTE, VarPtr(chr2(0))
                    .RasterPos2d CapX2 - 1, CapY2 - 1
                    .CallLists ccnt2, GL_UNSIGNED_BYTE, VarPtr(chr2(0))
                    .listbase 2500
                    .Color4f 1, 1, 1, 1
                    .RasterPos2d CapX1, CapY1
                    .CallLists ccnt1, GL_UNSIGNED_BYTE, VarPtr(chr1(0))
                    .Color4f 0#, 0#, 0#, 0.9
                    .RasterPos2d CapX1 + 1, CapY1 + 1
                    .CallLists ccnt1, GL_UNSIGNED_BYTE, VarPtr(chr1(0))
                    .RasterPos2d CapX1 - 1, CapY1 + 1
                    .CallLists ccnt1, GL_UNSIGNED_BYTE, VarPtr(chr1(0))
                    .RasterPos2d CapX1 + 1, CapY1 - 1
                    .CallLists ccnt1, GL_UNSIGNED_BYTE, VarPtr(chr1(0))
                    .RasterPos2d CapX1 - 1, CapY1 - 1
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
Private Function GetSubName(ByVal h As Long, ByVal S As Long, ByVal v As Long)
    Dim rindex As Long, clarr As Variant
    If S < 50 Then
        If v > 240 Then 'white
            clarr = Array("Glacial-pale", "Frozen-star", "Snowbound")
        ElseIf v < 80 Then 'black
            clarr = Array("Umbral", "Nether-shadow", "Void-dark")
        Else
            clarr = Array("Steel-white", "Metallic-howl", "Clockwork-silver")
        End If
    Else
        If 50 >= h And h > 35 Then 'Yellow
            clarr = Array("Aurora-blessed", "Dawn-lit", "Divine-radiant")
        ElseIf 120 >= h And h > 50 Then 'Green
            clarr = Array("Verdant-gloom", "Emerald-curse", "Eldritch-verdure")
        ElseIf 135 >= h And h > 120 Then 'Sky
            clarr = Array("Skyfall-whisper", "Celestial-wind", "Aerial-mirage")
        ElseIf 205 >= h And h > 135 Then 'Blue
            clarr = Array("Abyss-blue", "Ethereal-azure", "Cosmic-glow")
        ElseIf 220 >= h And h > 205 Then 'violet
            clarr = Array("Arcane-violet", "Eclipse-veil", "Nightshade-curse")
        Else 'red
            clarr = Array("Crimson-stained", "Scarlet-wrath", "Blood-soaked")
        End If
    End If
    rindex = Int((UBound(clarr) + 1) * Rnd())
    GetSubName = clarr(rindex)
End Function
Private Sub RGB2HSV(ByVal C As Long, ByRef h As Long, ByRef S As Long, ByRef v As Long)
    Dim R As Long, G As Long, B As Long
    Dim max As Long, min As Long, dif As Long
    B = (C And &HFF0000) \ &H10000
    G = (C And &HFF00&) \ &H100
    R = C And &HFF
    max = R: min = R
    If G > max Then max = G
    If B > max Then max = B
    If G < min Then min = G
    If B < min Then min = B
    v = max
    dif = max - min
    If dif = 0 Then
        h = 0: S = 0
        Exit Sub
    End If
    S = (255 * dif) \ max
    If max = R Then
        h = (43 * (G - B)) \ dif
    ElseIf max = G Then
        h = 85 + (43 * (B - R)) \ dif
    Else
        h = 170 + (43 * (R - G)) \ dif
    End If
    If h < 0 Then h = h + 255
End Sub

