VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} CFormPhysicsGLEffector 
   Caption         =   "CFormPhysicsEffecter"
   ClientHeight    =   9570.001
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   11910
   OleObjectBlob   =   "CFormPhysicsGLEffector.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "CFormPhysicsGLEffector"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Implements ICFormPhysicsEx
#If Win64 Then
    Private Declare PtrSafe Function GetWindowLongPtr Lib "user32" Alias "GetWindowLongPtrA" (ByVal hwnd As LongPtr, ByVal nIndex As Long) As LongPtr
    Private Declare PtrSafe Function SetWindowLongPtr Lib "user32" Alias "SetWindowLongPtrA" (ByVal hwnd As LongPtr, ByVal nIndex As Long, ByVal dwNewLong As LongPtr) As LongPtr
#Else
    Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
    Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
#End If
Private Declare PtrSafe Function WindowFromAccessibleObject Lib "oleacc.dll" (ByVal IAccessible As Object, ByRef hwnd As LongPtr) As LongPtr
Private Declare PtrSafe Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As LongPtr, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Private Declare PtrSafe Function ShowWindow Lib "user32" (ByVal hwnd As LongPtr, ByVal nCmdShow As Long) As Long
Private Declare PtrSafe Function SetWindowPos Lib "user32" (ByVal hwnd As LongPtr, ByVal hWndInsertAfter As LongPtr, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Const SWP_NOSIZE As Long = &H1, SWP_NOMOVE As Long = &H2, SW_SHOWMAXIMIZED = 3
Private Const HWND_BOTTOM As Long = 1, HWND_TOPMOST As Long = -1
Private Const GWL_STYLE As Long = -16, GWL_EXSTYLE As Long = -20
Private Const WS_EX_LAYERED = &H80000, WS_THICKFRAME = &H40000, WS_CAPTION = &HC00000, WS_EX_DLGMODALFRAME = &H1&
Private Const LWA_COLORKEY = &H1
Private Const EF_MAX_CNT = 3
Private Const DICT As String = "Scripting.Dictionary"
Private GL As OpenGL
Private hwnd As LongPtr, exStyle As LongPtr, style As LongPtr, rhWnd As LongPtr
Private width As Double, height As Double, hh As Double, hw As Double, ptime As Long, pdmg As Double, Tw2Px As Double
Private WithEvents myCore As CFormPhysics
Attribute myCore.VB_VarHelpID = -1
Private ChFct As Variant, ccnt As Long, BkFct As Variant, bcnt As Long, MvFct As Variant
Public ChDict As Object, BkDict As Object, MvDict As Object
Public Sub Render(Optional time = 0, Optional x As Double = -1, Optional y As Double = -1, Optional v As Double = -1)
    Dim dt As Long, i As Long
    Dim effName As Variant
    dt = time - ptime
    ptime = time
    With GL
        .Clear GL_COLOR_BUFFER_BIT Or GL_DEPTH_BUFFER_BIT
        .MatrixMode GL_PROJECTION
        .LoadIdentity
        .Ortho2D 0, width, height, 0
        .MatrixMode GL_MODELVIEW
        .LoadIdentity
        If ChDict.count > 0 Then
            .PushMatrix
                With ChDict
                    For Each effName In .keys()
                        With .Item(effName)
                            For i = 1 To EF_MAX_CNT
                                Call .Item(i).Render(x, y, dt, v)
                            Next i
                        End With
                    Next effName
                End With
            .PopMatrix
        End If
        If BkDict.count > 0 Then
            .PushMatrix
                With BkDict
                    For Each effName In .keys()
                        With .Item(effName)
                            For i = 1 To EF_MAX_CNT
                                Call .Item(i).Render(x, y, dt, v)
                            Next i
                        End With
                    Next effName
                End With
            .PopMatrix
        End If
        If MvDict.count > 0 Then
            .PushMatrix
                With MvDict
                    For Each effName In .keys()
                        Call .Item(effName).Render(x, y, dt, v)
                    Next effName
                End With
            .PopMatrix
        End If
        .SwapBuffers
    End With
End Sub
Private Sub myCore_Move(x As Double, y As Double, veloc As Double, time As Long)
    Render time, x * Tw2Px, y * Tw2Px, veloc
End Sub
Private Sub myCore_Crash(x As Double, y As Double, dmg As Double, time As Long)
    Dim ddmg As Double, tx As Double, ty As Double
    ddmg = dmg - pdmg
    tx = Tw2Px * x
    ty = Tw2Px * y
    If ddmg > 0.05 Then
        Dim effName As Variant
        With ChDict
            For Each effName In .keys()
                Call .Item(effName).Item(ccnt).Reset(tx, ty, ddmg, hw, hh)
            Next effName
        End With
        ccnt = ccnt + 1
        If ccnt > EF_MAX_CNT Then ccnt = 1
    End If
    With MvDict
        For Each effName In .keys()
            Call .Item(effName).Reset(tx, ty, ddmg, hw, hh)
        Next effName
    End With
    pdmg = dmg
End Sub
Private Sub myCore_Break(x As Double, y As Double, ofsx As Double, ofsy As Double, hw As Double, hh As Double)
    Dim effName As Variant
    With BkDict
        For Each effName In .keys()
            Call .Item(effName).Item(ccnt).Reset(Tw2Px * ofsx, Tw2Px * ofsy, 1#, hw, hh)
        Next effName
    End With
    bcnt = bcnt + 1
    If bcnt > EF_MAX_CNT Then bcnt = 1
End Sub
Private Sub myCore_Started(x As Double, y As Double, time As Long)
    Dim effName As Variant
    ptime = Timer
    With MvDict
        For Each effName In .keys()
            Call .Item(effName).Reset(x * Tw2Px, y * Tw2Px, 0, hw, hh)
        Next effName
    End With
End Sub
Private Sub myCore_Stopped(x As Double, y As Double, time As Long)
    With GL
        .Clear GL_COLOR_BUFFER_BIT Or GL_DEPTH_BUFFER_BIT
        .SwapBuffers
    End With
End Sub

Private Sub UserForm_Activate()
    If GL Is Nothing Then
        Me.RenderFrame.BackColor = RGB(254, 254, 254)
        WindowFromAccessibleObject Me, hwnd
        WindowFromAccessibleObject Me.RenderFrame, rhWnd
        If True Then 'debug_flg
            #If Win64 Then
                style = GetWindowLongPtr(hwnd, GWL_STYLE)
            #Else
                style = GetWindowLong(hwnd, GWL_STYLE)
            #End If
            style = (style Or WS_THICKFRAME Or &H30000) And Not WS_CAPTION
            #If Win64 Then
                SetWindowLongPtr hwnd, GWL_STYLE, style
            #Else
                SetWindowLong hwnd, GWL_STYLE, style
            #End If
            #If Win64 Then
                exStyle = GetWindowLongPtr(hwnd, GWL_EXSTYLE)
            #Else
                exStyle = GetWindowLong(hwnd, GWL_EXSTYLE)
            #End If
            #If Win64 Then
                SetWindowLongPtr hwnd, GWL_EXSTYLE, exStyle Or WS_EX_LAYERED
            #Else
                SetWindowLong hwnd, GWL_EXSTYLE, exStyle Or WS_EX_LAYERED
            #End If
            'Force-enable SetLayeredWindowAttributes by rendering OpenGL
            'through GDI to make the background color transparent.”
            SetLayeredWindowAttributes hwnd, RGB(254, 254, 254), 0, LWA_COLORKEY
            ShowWindow hwnd, SW_SHOWMAXIMIZED
            SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
        End If
        Call GLInit
    End If
End Sub
Public Sub GLInit()
    If width <= 0 Then width = 1920: height = 1080
    Set GL = New OpenGL
    With GL
        .hwnd = rhWnd
        .PaintStart
        .ClearColor 254 / 255, 254 / 255, 254 / 255, 1
        .Enable GL_DEPTH_TEST
        .Viewport 0, 0, width, height
        .Disable GL_LIGHTING
    End With
    DoEvents
    Set ChDict = CreateObject(DICT)
    Set MvDict = CreateObject(DICT)
    Set BkDict = CreateObject(DICT)
    On Error GoTo err:
        Dim efs As Variant, ek As Variant, ef As ICFormPhysicsEf, ci As ICFormPhysicsEf, t, i As Long
        If IsEmpty(ChFct) = False Then
            For Each t In ChFct
                ek = TypeName(t)
                Set ef = t
                If Not ChDict.exists(ek) Then
                    Dim coll As Collection: Set coll = New Collection
                    For i = 1 To EF_MAX_CNT
                        Set ci = ef.CreateInstance
                        Call ci.init(GL)
                        coll.Add ci
                    Next i
                    ChDict.Add ek, coll
                End If
            Next t
        End If
        If IsEmpty(MvFct) = False Then
            For Each t In MvFct
                ek = TypeName(t)
                Set ef = t
                If Not MvDict.exists(ek) Then
                    Set ci = ef.CreateInstance
                    Call ci.init(GL)
                    MvDict.Add ek, ci
                End If
            Next t
        End If
        If IsEmpty(BkFct) = False Then
            For Each t In BkFct
                ek = TypeName(t)
                Set ef = t
                If Not BkDict.exists(ek) Then
                    Dim coll2 As Collection: Set coll2 = New Collection
                    For i = 1 To EF_MAX_CNT
                        Set ci = ef.CreateInstance
                        Call ci.init(GL)
                        coll2.Add ci
                    Next i
                    BkDict.Add ek, coll2
                End If
            Next t
        End If
err:
    With myCore
        hw = .hw * Tw2Px
        hh = .hh * Tw2Px
    End With
    Call Render
End Sub
Private Sub ICFormPhysicsEx_init(core As CFormPhysics, Optional params As Variant = Empty)
    Set myCore = core
    With core
        width = .ScrWidth
        height = .ScrHeight
        Tw2Px = 1 / .Px2Tw
    End With
    ccnt = 1
    bcnt = 1
    On Error GoTo err
        'params = Array(Array(glExplosion, glShockWave), Array(glMoveTrail))
        If IsEmpty(params) Then
            Debug.Print "NoEffects"
        Else
            Dim acnt As Long: acnt = UBound(params)
            If IsEmpty(params(0)) = False Then ChFct = params(0)
            If acnt > 0 Then If IsEmpty(params(1)) = False Then MvFct = params(1)
            If acnt > 1 Then If IsEmpty(params(2)) = False Then BkFct = params(2)
        End If
        Me.Show vbModeless
        Exit Sub
err:
    Me.Label1.Caption = Me.Label1.Caption & err.Description
    Me.Show vbModeless
End Sub
Private Property Get ICFormPhysicsEx_CreateInstance() As ICFormPhysicsEx
    Set ICFormPhysicsEx_CreateInstance = New CFormPhysicsGLEffector
End Property
Private Sub ICFormPhysicsEx_Terminate()
On Error GoTo err
    Set MvDict = Nothing
    Set ChDict = Nothing
    Set BkDict = Nothing
    Set myCore = Nothing
    GL.PaintEnd
err:
    Set GL = Nothing
    Unload Me
End Sub
Private Sub UserForm_Layout()
    With Me.RenderFrame
        Me.Label1.width = Me.width
        .width = Me.width
        .height = Me.height - Me.top
    End With
End Sub
Private Sub Label1_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Call ICFormPhysicsEx_Terminate
End Sub
Private Sub UserForm_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Call ICFormPhysicsEx_Terminate
End Sub
