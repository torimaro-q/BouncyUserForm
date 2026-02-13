VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "glHitNumber"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Implements ICFormPhysicsEf
Private effectCX As Double
Private effectCY As Double
Private chr() As Byte, ccnt As Long
Private myGL As OpenGL, life As Double, InvLF As Double, delay As Double, dd As Double, hp As Double, myListBase As Long
Private R As Single, G As Single, B As Single, alpha As Double, t As Double
Private Sub ICFormPhysicsEf_Render(X As Double, Y As Double, dt As Long, v As Double)
    delay = delay - dt
    If delay >= 0 Then Exit Sub
    life = life - dt
    If life < 0 Then Exit Sub
    t = 1 - (life * InvLF)
    effectCY = effectCY - dt * 0.05
    With myGL
        .PushMatrix
            .Disable GL_LIGHTING
            .Enable GL_BLEND
            .BlendFunc GL_SRC_ALPHA, GL_ONE_MINUS_SRC_ALPHA
            If t > 0.6 Then
                alpha = (1# - t) * 2.5
                If alpha < 0 Then alpha = 0
            Else
                alpha = 1
            End If
            .listbase myListBase
            .Color4f R, G, B, alpha
            .RasterPos2d effectCX, effectCY
            .CallLists ccnt, GL_UNSIGNED_BYTE, VarPtr(chr(0))
            .listbase 0
            .Disable GL_BLEND
        .PopMatrix
    End With
End Sub
Private Sub ICFormPhysicsEf_Reset(cx As Double, cy As Double, ddmg As Double, Optional hw As Double = 0#, Optional hh As Double = 0#)
    Dim i As Long, rv As Double, th As Double
    dd = ddmg
    effectCX = cx
    effectCY = cy
    life = 500
    InvLF = 1 / life
    delay = 5
    Select Case dd
        Case Is > 5
            myListBase = 3000
            R = 1
            G = 0
            B = 0
        Case Is > 1
            myListBase = 2000
            R = 1
            G = 0.5
            B = 0
        Case Else
            myListBase = 1000
            R = 0.5
            G = 0.5
            B = 1
    End Select
    chr = StrConv(format(dd, "0.0"), vbFromUnicode)
    ccnt = UBound(chr) + 1
End Sub
Private Property Get ICFormPhysicsEf_CreateInstance() As ICFormPhysicsEf
    Set ICFormPhysicsEf_CreateInstance = New glHitNumber
End Property
Private Sub ICFormPhysicsEf_Init(targetGL As OpenGL)
    Set myGL = targetGL
    hp = 300
End Sub
