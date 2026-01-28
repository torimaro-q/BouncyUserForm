VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CFormPhysics"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit 'To determine the presence of an extension class
Private Type RGBA
    R As Byte
    G As Byte
    B As Byte
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
Private Declare PtrSafe Function GetDC Lib "user32" (ByVal hwnd As LongPtr) As LongPtr
Private Declare PtrSafe Function GetDeviceCaps Lib "gdi32" (ByVal hdc As LongPtr, ByVal nIndex As Long) As Long
Private Declare PtrSafe Function ReleaseDC Lib "user32" (ByVal hwnd As LongPtr, ByVal hdc As LongPtr) As Long
Private Declare PtrSafe Function GetTickCount Lib "kernel32" () As Long
'##############################################################################################
Private Const SPDLMT As Double = 10000
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
Private DragX As Double, DragY As Double
Private Caps(3) As String, dCap As String, frameCnt As Long
Private RChk As Double, TChk As Double, cV As Double, pV As Double, tCl As Long
Private WithEvents mFrm As MSForms.UserForm
Attribute mFrm.VB_VarHelpID = -1
Public mFrmRaw As Object, ex As Object
Public isDragging As Boolean, isAnimation As Boolean
Public lx As Double, ty As Double, VX As Double, VY As Double, AX As Double, AY As Double
Public Damage As Double, ptime As Long, cTime As Long, Dens As Double, Mass As Double, InvMass As Double, hw As Double, hh As Double
Public ScrLeft As Double, ScrTop As Double, ScrWidth As Double, ScrHeight As Double, Px2Tw As Double
Public Event Crash(ByRef x As Double, ByRef y As Double, ByRef dmg As Double, ByRef time As Long)
Public Event Break(ByRef x As Double, ByRef y As Double, ByRef ofsx As Double, ByRef ofsy As Double, ByRef hw As Double, ByRef hh As Double)
Public Event Move(ByRef x As Double, ByRef y As Double, ByRef veloc As Double, ByRef time As Long)
Public Event Started(ByRef x As Double, ByRef y As Double, ByRef time As Long)
Public Event Stopped(ByRef x As Double, ByRef y As Double, ByRef time As Long)
Private Sub AnimationLoop()
    Dim dt As Double, isCrashed As Boolean
    With mFrmRaw
        lx = .left
        ty = .top
    End With
    RaiseEvent Started(lx + hw, ty + hh, cTime)
    Do
        cTime = GetTickCount
        If ptime <> 0 Then
            dt = SAFEOFS + (cTime - ptime) * 0.001
            AX = -DragX * InvMass * (VX + PRES2VIS * (VX * Abs(VX)))
            AY = GRAV - DragY * InvMass * (VY + PRES2VIS * (VY * Abs(VY)))
            VX = VX + AX * dt
            VY = VY + AY * dt
            lx = lx + VX * dt
            ty = ty + VY * dt
        End If
        ptime = cTime
        isCrashed = ApplyCollision
        Call UpdateVeloc
        If isDragging Then Exit Do
        If isAnimation = False Then Exit Do
        If isCrashed = True Then ApplyDamage
        If cV < 10 Then Exit Do
        On Error GoTo err 'avoid automation errors
            mFrmRaw.Move lx, ty
        On Error GoTo 0
        Render
    Loop
    RaiseEvent Stopped(lx + hw, ty + hh, cTime)
    ResetPhysics
    Render
    Exit Sub
err:
    isAnimation = False
End Sub
Private Function ApplyCollision() As Boolean
    ApplyCollision = True
    If lx < ScrLeft Then
        lx = ScrLeft
        VX = -VX * ZETA
        VY = VY * ZETA
    ElseIf lx > RChk Then
        lx = RChk
        VX = -VX * ZETA
        VY = VY * ZETA
    ElseIf ty < ScrTop Then
        ty = ScrTop
        VX = VX * ZETA
        VY = -VY * ZETA
    ElseIf ty > TChk Then
        ty = TChk
        VX = VX * ZETA
        VY = -VY * ZETA
    Else
        ApplyCollision = False
    End If
End Function
Private Sub ApplyDamage()
    Damage = Damage + DMGCF * 0.5 * Mass * Abs(pV * pV - cV * cV)
    RaiseEvent Crash(lx + hw, ty + hh, Damage, cTime)
    If Damage > 10 Then
        Dim th As Double, cntr As MSForms.Control
        th = 1 - Damage * 0.002
        On Error GoTo err 'avoid automation errors
            For Each cntr In mFrmRaw.Controls
                If Rnd() > th Then
                    With cntr
                        If .Visible Then
                            RaiseEvent Break(lx + hw, ty + hh, lx + .left, ty + .top, 0.5 * .width, 0.5 * .height)
                            .Visible = False
                        End If
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
            If Abs(lx - .left) > 30 Or Abs(ty - .top) > 30 Then
                isDragging = False
                Call ResetPhysics
                Call Render
            End If
        End With
    Else
        With mFrmRaw
            RChk = ScrLeft + ScrWidth * Px2Tw - .width
            TChk = ScrTop + ScrHeight * Px2Tw - .height
            DragX = .height * THICKNESS
            DragY = .width * THICKNESS
            hh = .height * 0.5
            hw = .width * 0.5
            Call Render
        End With
        Mass = getMass
        InvMass = 1 / Mass
    End If
End Sub
Private Function getVeloc(ByRef VX As Double, ByRef VY As Double) As Double
    'Linear Approximation of Pythagoras' Theorem
    If Abs(VX) > Abs(VY) Then getVeloc = PA * Abs(VX) + PB * Abs(VY) Else getVeloc = PA * Abs(VY) + PB * Abs(VX)
End Function
Private Sub Render()
    On Error GoTo err 'avoid automation errors
        With mFrmRaw
            If isAnimation = True Then
                If frameCnt Mod 5 = 0 Then
                    RaiseEvent Move(lx + hw, ty + hh, cV, cTime)
                    If frameCnt Mod 30 = 0 Then
                        Caps(1) = format(cV, FMT)
                        Caps(3) = format(Damage, FMT)
                        .Caption = Join(Caps, "") 'High-Speed String Concatenation
                    End If
                End If
            Else
                .Caption = dCap
                If Damage > 30 Then .Caption = .Caption & "?"
            End If
        End With
    On Error GoTo 0
    frameCnt = frameCnt + 1
    If frameCnt >= 50 Then frameCnt = 0: DoEvents
err:
End Sub
Private Sub UpdateVeloc()
    pV = cV
    cV = getVeloc(VX, VY)
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
    Render
End Sub
Private Sub Class_Terminate()
    With Application
        .ScreenUpdating = True
        .Calculation = xlCalculationAutomatic
    End With
    isDragging = True
    If Not (ex Is Nothing) Then
        On Error GoTo err
        Dim tmp As Variant
        For Each tmp In ex.keys()
            ex.Item(tmp).Terminate
        Next tmp
err:
        Set ex = Nothing
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
            lx = .left
            ty = .top
            .left = .left + x - hw
            .top = .top + y + 10 'The offset to automatically center the selection in the title bar when the user form is clicked.
            cTime = GetTickCount
            If ptime <> 0 Then
                Dim dt As Double, cff As Double
                dt = SAFEOFS + (cTime - ptime) * 0.001
                cff = 1 / (dt + 0.005) 'Vibration Countermeasures
                VX = 0.5 * VX + (.left - lx) * cff
                VY = 0.5 * VY + (.top - ty) * cff
                If Abs(VX) > SPDLMT Then VX = Sgn(VX) * SPDLMT
                If Abs(VY) > SPDLMT Then VY = Sgn(VY) * SPDLMT
            End If
            ptime = cTime
        End With
    Else
        'If you want to add an action to avoid the user when taking damage, put it here.
    End If
End Sub
Private Sub mFrm_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    If isDragging Then Launch
End Sub
Public Sub Launch()
    If isAnimation Then Exit Sub
    isDragging = False
    isAnimation = True
    pV = 0
    ptime = 0
    Call AnimationLoop
End Sub
Private Function ResetPhysics()
    VX = 0: VY = 0
    pV = 0: cV = 0: ptime = 0
    isAnimation = False
End Function
Private Function CL2RGBA(ByRef cl As Long) As RGBA
    Dim L As LONGTYPE: L.Value = cl
    LSet CL2RGBA = L
End Function
Public Sub init(ByRef form As Object, Optional ByRef extenions As Variant = Empty, _
                        Optional ByRef cEffectors As Variant = Empty, Optional ByRef mEffectors As Variant = Empty, Optional ByRef bEffectors As Variant = Empty)
    If mFrmRaw Is Nothing Then
        Set mFrmRaw = form
        Set mFrm = form
        Px2Tw = 72 / GetDPI()
        With mFrmRaw
            .DrawBuffer = 640000 'prevent flickering
            dCap = .Caption
            tCl = SysCL2RGB(.BackColor)
            With CL2RGBA(tCl)
                Dens = (300 - (0.299 * .R + 0.587 * .G + 0.114 * .B)) / 5000
            End With
        End With
        isDragging = False
        Call ResetPhysics
        On Error GoTo err
            Dim nm, t, fc As ICFormPhysicsEx, ci As ICFormPhysicsEx
            Set ex = CreateObject("Scripting.Dictionary")
            For Each t In extenions
                nm = TypeName(t)
                If ex.exists(nm) = False Then
                    Set fc = t
                    Set ci = fc.CreateInstance
                    ci.init Me, Array(cEffectors, mEffectors, bEffectors)
                    ex.Add nm, ci
                End If
            Next t
    End If
err:
End Sub
Private Sub Class_Initialize()
    ScrLeft = GetSystemMetrics(SM_XVIRTUALSCREEN): ScrTop = GetSystemMetrics(SM_YVIRTUALSCREEN)
    ScrWidth = GetSystemMetrics(SM_CXVIRTUALSCREEN): ScrHeight = GetSystemMetrics(SM_CYVIRTUALSCREEN)
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
