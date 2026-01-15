VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CFormPhysics"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Option Explicit 'To determine the presence of an extension class
Private Type POINT
    x As Double
    y As Double
End Type
Private Type RGBA
    r As Byte
    g As Byte
    b As Byte
    A As Byte
End Type
Private Type LONGTYPE
    Value As Long
End Type
Private Const SM_XVIRTUALSCREEN = 76
Private Const SM_YVIRTUALSCREEN = 77
Private Const SM_CXVIRTUALSCREEN = 78
Private Const SM_CYVIRTUALSCREEN = 79
Private Const LOGPIXELSX As Long = 88
Private Const GWL_EXSTYLE = -20
Private Const WS_EX_LAYERED = &H80000
Private Const LWA_COLORKEY = &H1
Private Declare PtrSafe Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Private Declare PtrSafe Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Private Declare PtrSafe Function GetDC Lib "user32" (ByVal hWnd As LongPtr) As LongPtr
Private Declare PtrSafe Function GetDeviceCaps Lib "gdi32" (ByVal hdc As LongPtr, ByVal nIndex As Long) As Long
Private Declare PtrSafe Function ReleaseDC Lib "user32" (ByVal hWnd As LongPtr, ByVal hdc As LongPtr) As Long
Private Declare PtrSafe Function GetTickCount Lib "kernel32" () As Long
'##############################################################################################
Private Const SAFETYOFFSET As Double = 0.00001
Private Const ALPHA As Double = 0.9604
Private Const BETA As Double = 0.3978
Private Const ZETA As Double = 0.8
Private Const GRAV As Double = 300
Private Const DMGCOEFF As Double = 0.000000002
Private Const SPDCOEFF As Double = 0.75
Private Const PRES2VIS As Double = 0.001
Private Const THICKNESS As Double = 1
Private Const FORMATCAP As String = "0.000"
Private WithEvents myForm As MSForms.UserForm
Attribute myForm.VB_VarHelpID = -1
Private myFormRaw As Object 'Workaround for UserForm TypeLib lacking Left and Top properties
Private DragCoeff As POINT
Private Dst As POINT, VVec As POINT, AVec As POINT
Private Density As Double
Private isDragging As Boolean, isAnimation As Boolean
Private pTime As Long, cTime As Long, frameCnt As Long
Private damage As Double
Private Caps(3) As String
Private Mass As Double, InvMass As Double, hw As Double, hh As Double
Private RChk As Double, TChk As Double
Private dCaption As String
Private cV As Double, pV As Double
Private CrackPictArr() As Long
Private sR As Long, sG As Long, sB As Long, tCl As Long
Public ScrLeft As Double, ScrTop As Double, ScrWidth As Double, ScrHeight As Double
Public Px2Tw  As Double
Public Event Crash(ByRef x As Double, ByRef y As Double, dmg As Double)
Public Event Move(ByRef x As Double, ByRef y As Double, ByRef veloc As Double)
Public Event Stopped(ByRef x As Double, ByRef y As Double)
Private ExColl As New Collection
Private Sub AnimationLoop()
    Dim dt As Double, CrashFlg As Boolean
    With myFormRaw
        Dst.x = .Left
        Dst.y = .Top
    End With
    Do
        With VVec
            cTime = GetTickCount
            If pTime <> 0 Then
                dt = SAFETYOFFSET + (cTime * 0.001 - pTime * 0.001)
                AVec.x = -DragCoeff.x * InvMass * (.x + PRES2VIS * (.x * Abs(.x)))
                AVec.y = GRAV - DragCoeff.y * InvMass * (.y + PRES2VIS * (.y * Abs(.y)))
                .x = .x + AVec.x * dt
                .y = .y + AVec.y * dt
                Dst.x = Dst.x + .x * dt
                Dst.y = Dst.y + .y * dt
            End If
            pTime = cTime
            CrashFlg = ApplyCollision
            Call UpdateVeloc
            If isDragging Then Exit Do
            If isAnimation = False Then Exit Do
            If CrashFlg = True Then ApplyDamage
            If cV < 10 Then Exit Do
        End With
        On Error GoTo err 'Access when the form is closed to avoid automation errors
            myFormRaw.Move Dst.x, Dst.y
        On Error GoTo 0
        Render
    Loop
    RaiseEvent Stopped(Dst.x + hw, Dst.y + hh)
    ResetPhysics
    Render
    Exit Sub
err:
    isAnimation = False
End Sub
Private Function ApplyCollision() As Boolean  'Wall collision detection
    ApplyCollision = False
    With VVec
        If Dst.x < ScrLeft Then
            Dst.x = ScrLeft
            .x = -.x * ZETA
            .y = .y * ZETA
            ApplyCollision = True
        ElseIf Dst.x > RChk Then
            Dst.x = RChk
            .x = -.x * ZETA
            .y = .y * ZETA
            ApplyCollision = True
        ElseIf Dst.y < ScrTop Then
            Dst.y = ScrTop
            .x = .x * ZETA
            .y = -.y * ZETA
            ApplyCollision = True
        ElseIf Dst.y > TChk Then
            Dst.y = TChk
            .x = .x * ZETA
            .y = -.y * ZETA
            ApplyCollision = True
        End If
    End With
End Function
Private Sub ApplyDamage()
    Dim cntr As MSForms.Control, th As Double
    damage = damage + DMGCOEFF * 0.5 * Mass * Abs(pV * pV - cV * cV) 'Kinetic energy lost upon colliding with the wall
    RaiseEvent Crash(Dst.x + hw, Dst.y + hh, damage)
    With WorksheetFunction
        tCl = RGB(.Min(sR + damage, 255), .Max(sG - damage * 0.1, sG * 0.9), .Max(sB - damage * 0.1, sB * 0.9))
    End With
    If damage > 10 Then
        th = 0.995
        If damage > 100 Then th = 0.85
        On Error GoTo err 'Access when the form is closed to avoid automation errors
            For Each cntr In myFormRaw.Controls
                If Rnd() > th Then cntr.Visible = False
            Next cntr
        On Error GoTo 0
    End If
err:
End Sub
Private Sub myForm_Layout()
    If isAnimation Then
        'Forced termination conditions
        With myFormRaw
            If Abs(Dst.x - .Left) > 30 Or Abs(Dst.y - .Top) > 30 Then
                isDragging = False
                Call ResetPhysics
                Call Render
            End If
        End With
    Else
        Mass = getMass
        InvMass = 1 / Mass
        With myFormRaw
            RChk = ScrLeft + ScrWidth * Px2Tw - .Width
            TChk = ScrTop + ScrHeight * Px2Tw - .Height
            DragCoeff.x = .Height * THICKNESS
            DragCoeff.y = .Width * THICKNESS
            hh = .Height * 0.5
            hw = .Width * 0.5
            Call Render
        End With
    End If
End Sub
Private Function getVeloc(ByRef vx As Double, ByRef vy As Double) As Double
    Dim avx As Double: avx = Abs(vx)
    Dim avy As Double: avy = Abs(vy)
    'Linear Approximation of Pythagoras' Theorem
    If avx > avy Then getVeloc = ALPHA * avx + BETA * avy Else getVeloc = ALPHA * avy + BETA * avx
End Function
Private Sub Render()
    On Error GoTo err 'Access when the form is closed to avoid automation errors
        With myFormRaw
            If isAnimation = True Then
                If frameCnt Mod 30 = 0 Then
                    Caps(1) = Format(cV, FORMATCAP)
                    Caps(3) = Format(damage, FORMATCAP)
                    .Caption = Join(Caps, "") 'High-Speed String Concatenation
                    .BackColor = tCl
                    RaiseEvent Move(Dst.x + hw, Dst.y + hh, cV)
                End If
            Else
                .Caption = dCaption
                If damage > 30 Then .Caption = .Caption & "?"
                .BackColor = tCl
            End If
        End With
    On Error GoTo 0
    frameCnt = frameCnt + 1
    If frameCnt >= 50 Then
        frameCnt = 0
        DoEvents
    End If
err:
End Sub
Private Sub UpdateVeloc()
    With VVec
        pV = cV
        cV = getVeloc(.x, .y)
    End With
End Sub
Private Function getMass() As Double
    With myFormRaw
        getMass = .Width * .Height * THICKNESS * Density
    End With
End Function
Private Sub Class_Terminate()
    isDragging = True
    Set ExColl = Nothing
End Sub
Private Sub Restore()
    Dim cntr As MSForms.Control
    damage = 0
    For Each cntr In myFormRaw.Controls
        cntr.Visible = True
    Next cntr
    tCl = RGB(sR, sG, sB)
    Render
End Sub
Private Sub myForm_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    If Button = 1 Then
        isDragging = True
        Call ResetPhysics
    Else
        Call Restore
    End If
End Sub
Private Sub myForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    If isDragging Then 'When the title bar is selected directly, disable physics calculations (for user consideration).
        With myFormRaw
            Dst.x = .Left
            Dst.y = .Top
            .Left = .Left + x - 0.5 * .Width
            .Top = .Top + y + 10 'The offset to automatically center the selection in the title bar when the user form is clicked.
            cTime = GetTickCount
            If pTime <> 0 Then
                Dim dt As Double, cff As Double
                dt = SAFETYOFFSET + (cTime * 0.001 - pTime * 0.001)
                cff = 1 / (dt + 0.005) 'Vibration Countermeasures
                VVec.x = 0.5 * VVec.x + (.Left - Dst.x) * cff
                VVec.y = 0.5 * VVec.y + (.Top - Dst.y) * cff
                With VVec
                    If Abs(.x) > 10000 Then .x = Sgn(.x) * 10000
                    If Abs(.y) > 10000 Then .y = Sgn(.y) * 10000
                End With
            End If
            pTime = cTime
        End With
    Else
        'If you want to add an action to avoid the user when taking damage, put it here.
    End If
End Sub
Private Sub myForm_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    If isDragging Then
        isDragging = False
        isAnimation = True
        pV = 0
        pTime = 0
        Call AnimationLoop
    End If
End Sub
Private Function ResetPhysics()
    With VVec
        .x = 0: .y = 0
    End With
    pV = 0: cV = 0: pTime = 0
    isAnimation = False
End Function
Private Function CL2RGBA(ByRef cl As Long) As RGBA
    Dim l As LONGTYPE: l.Value = cl
    LSet CL2RGBA = l 'Pseudo-Move Memory
End Function
Public Sub Init(ByRef form As Object)
    If myFormRaw Is Nothing Then
        Set myFormRaw = form
        Set myForm = form
        Px2Tw = 72 / GetDPI()
        With myFormRaw
            .DrawBuffer = 320000 'Set the DrawBuffer to prevent flickering
            dCaption = .Caption
            With CL2RGBA(SystemColorToRGB(.BackColor))
                sR = .r: sG = .g: sB = .b
                tCl = RGB(sR, sG, sB)
                Density = (300 - WorksheetFunction.Max(sR, sG, sB)) / 5000
            End With
        End With
        isDragging = False
        Call ResetPhysics
        On Error GoTo err
            Set ExColl = New Collection
            Dim cfex As Object
            Set cfex = CFormPhysicsEx.CreateInstance(Me)
            Call ExColl.Add(cfex) 'Load if an extension class exists
    End If
err:
End Sub
Private Sub Class_Initialize()
    ScrLeft = GetSystemMetrics(SM_XVIRTUALSCREEN)
    ScrTop = GetSystemMetrics(SM_YVIRTUALSCREEN)
    ScrWidth = GetSystemMetrics(SM_CXVIRTUALSCREEN)
    ScrHeight = GetSystemMetrics(SM_CYVIRTUALSCREEN)
    Caps(0) = "speed : "
    Caps(2) = "   damage : "
End Sub
Private Function GetDPI() As Long
    Dim hdc As LongPtr: hdc = GetDC(0)
    GetDPI = GetDeviceCaps(hdc, LOGPIXELSX)
    Call ReleaseDC(0, hdc)
End Function
Public Function SystemColorToRGB(ByVal sysColor As Long) As Long
    If (sysColor And &H80000000) <> 0 Then SystemColorToRGB = GetSysColor(CLng(sysColor And &HFF&)) Else SystemColorToRGB = sysColor
End Function
