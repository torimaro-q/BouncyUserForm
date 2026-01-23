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
    a As Byte
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
Private Const CLMASK As Long = &H80000000
Private Declare PtrSafe Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Private Declare PtrSafe Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Private Declare PtrSafe Function GetDC Lib "user32" (ByVal hWnd As LongPtr) As LongPtr
Private Declare PtrSafe Function GetDeviceCaps Lib "gdi32" (ByVal hdc As LongPtr, ByVal nIndex As Long) As Long
Private Declare PtrSafe Function ReleaseDC Lib "user32" (ByVal hWnd As LongPtr, ByVal hdc As LongPtr) As Long
Private Declare PtrSafe Function GetTickCount Lib "kernel32" () As Long
'##############################################################################################
Private Const SAFEOFS As Double = 0.00001
Private Const PA As Double = 0.9604
Private Const PB As Double = 0.3978
Private Const FMT As String = "0.000"
Private Const THICKNESS As Double = 1
Private Const ZETA As Double = 0.8
Private Const GRAV As Double = 300
Private Const DMGCF As Double = 0.00000001
Private Const SPDCF As Double = 0.75
Private Const PRES2VIS As Double = 0.001
Private DragCf As POINT, dst As POINT, VV As POINT, VA As POINT
Private InvMass As Double, hw As Double, hh As Double
Private isDragging As Boolean, isAnimation As Boolean
Private Caps(3) As String, dCap As String, frameCnt As Long
Private RChk As Double, TChk As Double
Private cV As Double, pV As Double, sR As Long, sG As Long, sB As Long, tCl As Long
Private Ex As Object
Private WithEvents mFrm As MSForms.UserForm
Attribute mFrm.VB_VarHelpID = -1
Public mFrmRaw As Object
Public Damage As Double, pTime As Long, cTime As Long, Dens As Double, Mass As Double
Public ScrLeft As Double, ScrTop As Double, ScrWidth As Double, ScrHeight As Double, Px2Tw As Double
Public Event Crash(ByRef x As Double, ByRef y As Double, ByRef dmg As Double, ByRef time As Long)
Public Event Broke(ByRef x As Double, ByRef y As Double, ByRef ofsx As Double, ByRef ofsy As Double, ByRef hw As Double, ByRef hh As Long)
Public Event Move(ByRef x As Double, ByRef y As Double, ByRef veloc As Double, ByRef time As Long)
Public Event Started(ByRef x As Double, ByRef y As Double, ByRef time As Long)
Public Event Stopped(ByRef x As Double, ByRef y As Double, ByRef time As Long)
Private Sub AnimationLoop()
    Dim dt As Double, isCrashed As Boolean
    With mFrmRaw
        dst.x = .left
        dst.y = .top
    End With
    RaiseEvent Started(dst.x + hw, dst.y + hh, cTime)
    Do
        With VV
            cTime = GetTickCount
            If pTime <> 0 Then
                dt = SAFEOFS + (cTime - pTime) * 0.001
                VA.x = -DragCf.x * InvMass * (.x + PRES2VIS * (.x * Abs(.x)))
                VA.y = GRAV - DragCf.y * InvMass * (.y + PRES2VIS * (.y * Abs(.y)))
                .x = .x + VA.x * dt
                .y = .y + VA.y * dt
                dst.x = dst.x + .x * dt
                dst.y = dst.y + .y * dt
            End If
            pTime = cTime
            isCrashed = ApplyCollision
            Call UpdateVeloc
            If isDragging Then Exit Do
            If isAnimation = False Then Exit Do
            If isCrashed = True Then ApplyDamage
            If cV < 10 Then Exit Do
        End With
        On Error GoTo err 'avoid automation errors
            mFrmRaw.Move dst.x, dst.y
        On Error GoTo 0
        Render
    Loop
    RaiseEvent Stopped(dst.x + hw, dst.y + hh, cTime)
    ResetPhysics
    Render
    Exit Sub
err:
    isAnimation = False
End Sub
Private Function ApplyCollision() As Boolean
    With VV
        ApplyCollision = True
        If dst.x < ScrLeft Then
            dst.x = ScrLeft
            .x = -.x * ZETA
            .y = .y * ZETA
        ElseIf dst.x > RChk Then
            dst.x = RChk
            .x = -.x * ZETA
            .y = .y * ZETA
        ElseIf dst.y < ScrTop Then
            dst.y = ScrTop
            .x = .x * ZETA
            .y = -.y * ZETA
        ElseIf dst.y > TChk Then
            dst.y = TChk
            .x = .x * ZETA
            .y = -.y * ZETA
        Else
            ApplyCollision = False
        End If
    End With
End Function
Private Sub ApplyDamage()
    Dim cntr As MSForms.Control, th As Double, xn As Long, yn As Long
    Damage = Damage + DMGCF * 0.5 * Mass * Abs(pV * pV - cV * cV)
    RaiseEvent Crash(dst.x + hw, dst.y + hh, Damage, cTime)
    With WorksheetFunction
        tCl = RGB(.Min(sR + Damage, 255), .Max(sG - Damage * 0.1, sG * 0.9), .Max(sB - Damage * 0.1, sB * 0.9))
    End With
    If Damage > 10 Then
        th = 1 - Damage * 0.002
        On Error GoTo err 'avoid automation errors
            For Each cntr In mFrmRaw.Controls
                If Rnd() > th Then
                    With cntr
                        RaiseEvent Broke(dst.x + hw, dst.y + hh, dst.x + .left, dst.y + .top, 0.5 * .width, 0.5 * .height)
                        .Visible = False
                    End With
                End If
            Next cntr
        On Error GoTo 0
    End If
err:
End Sub
Private Sub mFrm_Layout()
    If isAnimation Then
        'Forced termination conditions
        With mFrmRaw
            If Abs(dst.x - .left) > 30 Or Abs(dst.y - .top) > 30 Then
                isDragging = False
                Call ResetPhysics
                Call Render
            End If
        End With
    Else
        With mFrmRaw
            RChk = ScrLeft + ScrWidth * Px2Tw - .width
            TChk = ScrTop + ScrHeight * Px2Tw - .height
            DragCf.x = .height * THICKNESS
            DragCf.y = .width * THICKNESS
            hh = .height * 0.5
            hw = .width * 0.5
            Call Render
        End With
        Mass = getMass
        InvMass = 1 / Mass
    End If
End Sub
Private Function getVeloc(ByRef vx As Double, ByRef vy As Double) As Double
    Dim avx As Double: avx = Abs(vx)
    Dim avy As Double: avy = Abs(vy)
    'Linear Approximation of Pythagoras' Theorem
    If avx > avy Then getVeloc = PA * avx + PB * avy Else getVeloc = PA * avy + PB * avx
End Function
Private Sub Render()
    On Error GoTo err 'avoid automation errors
        With mFrmRaw
            If isAnimation = True Then
                If frameCnt Mod 10 = 0 Then
                    RaiseEvent Move(dst.x + hw, dst.y + hh, cV, cTime)
                    If frameCnt Mod 30 = 0 Then
                        Caps(1) = format(cV, FMT)
                        Caps(3) = format(Damage, FMT)
                        .Caption = Join(Caps, "") 'High-Speed String Concatenation
                        .BackColor = tCl
                    End If
                End If
            Else
                .Caption = dCap
                If Damage > 30 Then .Caption = .Caption & "?"
                .BackColor = tCl
            End If
        End With
    On Error GoTo 0
    frameCnt = frameCnt + 1
    If frameCnt >= 50 Then frameCnt = 0: DoEvents
err:
End Sub
Private Sub UpdateVeloc()
    With VV
        pV = cV
        cV = getVeloc(.x, .y)
    End With
End Sub
Private Function getMass() As Double
    With mFrmRaw
        getMass = .width * .height * THICKNESS * Dens
    End With
End Function
Private Sub Restore()
    Dim c As MSForms.Control
    Damage = 0
    For Each c In mFrmRaw.Controls
        c.Visible = True
    Next c
    tCl = RGB(sR, sG, sB)
    Render
End Sub
Private Sub Class_Terminate()
    With Application
        .ScreenUpdating = True
        .Calculation = xlCalculationAutomatic
    End With
    isDragging = True
    If Not (Ex Is Nothing) Then
        On Error GoTo err
        Dim tmp As Variant
        For Each tmp In Ex.keys()
            Ex.Item(tmp).Terminate
        Next tmp
err:
        Set Ex = Nothing
    End If
End Sub
Private Sub mFrm_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    If Button = 1 Then
        isDragging = True
        Call ResetPhysics
    Else
        Call Restore
    End If
End Sub
Private Sub mFrm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    If isDragging Then 'When the title bar is selected directly, disable physics calculations (for user consideration).
        With mFrmRaw
            dst.x = .left
            dst.y = .top
            .left = .left + x - hw
            .top = .top + y + 10 'The offset to automatically center the selection in the title bar when the user form is clicked.
            cTime = GetTickCount
            If pTime <> 0 Then
                Dim dt As Double, cff As Double
                dt = SAFEOFS + (cTime - pTime) * 0.001
                cff = 1 / (dt + 0.005) 'Vibration Countermeasures
                VV.x = 0.5 * VV.x + (.left - dst.x) * cff
                VV.y = 0.5 * VV.y + (.top - dst.y) * cff
                With VV
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
Private Sub mFrm_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    If isDragging Then
        isDragging = False
        isAnimation = True
        pV = 0
        pTime = 0
        Call AnimationLoop
    End If
End Sub
Private Function ResetPhysics()
    With VV
        .x = 0: .y = 0
    End With
    pV = 0: cV = 0: pTime = 0
    isAnimation = False
End Function
Private Function CL2RGBA(ByRef cl As Long) As RGBA
    Dim l As LONGTYPE: l.Value = cl
    LSet CL2RGBA = l
End Function
Public Sub Init(ByRef form As Object, Optional ByRef extenions As Variant = Empty)
    If mFrmRaw Is Nothing Then
        Set mFrmRaw = form
        Set mFrm = form
        Px2Tw = 72 / GetDPI()
        With mFrmRaw
            .DrawBuffer = 320000 'Set the DrawBuffer to prevent flickering
            dCap = .Caption
            With CL2RGBA(SysCL2RGB(.BackColor))
                sR = .r: sG = .g: sB = .b
                tCl = RGB(sR, sG, sB)
                Dens = (300 - WorksheetFunction.Max(sR, sG, sB)) / 5000
            End With
        End With
        isDragging = False
        Call ResetPhysics
        On Error GoTo err
            Dim tmp, obj
            If IsEmpty(extenions) Then extenions = Array(CFormPhysicsLogger, CFormPhysicsWsRenderer)
            Set Ex = CreateObject("Scripting.Dictionary")
            For Each tmp In extenions
                If Ex.exists(TypeName(tmp)) = False Then
                    Set obj = CallByName(tmp, "CreateInstance", VbGet)
                    obj.Init Me
                    Ex.Add TypeName(tmp), obj
                End If
            Next tmp
    End If
err:
End Sub
Private Sub Class_Initialize()
    ScrLeft = GetSystemMetrics(SM_XVIRTUALSCREEN)
    ScrTop = GetSystemMetrics(SM_YVIRTUALSCREEN)
    ScrWidth = GetSystemMetrics(SM_CXVIRTUALSCREEN)
    ScrHeight = GetSystemMetrics(SM_CYVIRTUALSCREEN)
    Caps(0) = "speed : ": Caps(2) = "   damage : "
End Sub
Private Function GetDPI() As Long
    Dim hdc As LongPtr: hdc = GetDC(0)
    GetDPI = GetDeviceCaps(hdc, LOGPIXELSX)
    Call ReleaseDC(0, hdc)
End Function
Public Function SysCL2RGB(ByVal sysCL As Long) As Long
    If (sysCL And CLMASK) <> 0 Then SysCL2RGB = GetSysColor(CLng(sysCL And &HFF&)) Else SysCL2RGB = sysCL
End Function
Public Sub Terminate()
    Call Class_Terminate
End Sub
