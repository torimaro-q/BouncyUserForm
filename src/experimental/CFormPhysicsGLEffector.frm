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
#If Win64 Then
    Private Declare PtrSafe Function GetWindowLongPtr Lib "user32" Alias "GetWindowLongPtrA" (ByVal hWnd As LongPtr, ByVal nIndex As Long) As LongPtr
    Private Declare PtrSafe Function SetWindowLongPtr Lib "user32" Alias "SetWindowLongPtrA" (ByVal hWnd As LongPtr, ByVal nIndex As Long, ByVal dwNewLong As LongPtr) As LongPtr
#Else
    Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
    Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
#End If
Private Declare PtrSafe Function WindowFromAccessibleObject Lib "oleacc.dll" (ByVal IAccessible As Object, ByRef hWnd As LongPtr) As LongPtr
Private Declare PtrSafe Function SetLayeredWindowAttributes Lib "user32" (ByVal hWnd As LongPtr, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Private Declare PtrSafe Function ShowWindow Lib "user32" (ByVal hWnd As LongPtr, ByVal nCmdShow As Long) As Long
Private Const GWL_STYLE = -16
Private Const GWL_EXSTYLE = -20
Private Const WS_EX_LAYERED = &H80000
Private Const WS_THICKFRAME = &H40000
Private Const WS_CAPTION = &HC00000
Private Const WS_EX_DLGMODALFRAME = &H1&
Private Const LWA_COLORKEY = &H1
Private Const SW_SHOWMAXIMIZED = 3
Private Const EFFECT_MAX_COUNT = 4
Private GL As OpenGL
Private hWnd As LongPtr, exStyle As LongPtr, style As LongPtr
Private rndhWnd As LongPtr
Private width As Double, height As Double
Private WithEvents myCore As CFormPhysics, pTime As Long, pdmg As Double, Tw2Px As Double, ecnt2 As Long
Attribute myCore.VB_VarHelpID = -1
Private CrashDict As Object, ccnt As Long
Private MoveDict As Object, mcnt As Long
Private Sub myCore_Crash(x As Double, y As Double, dmg As Double, time As Long)
    If (dmg - pdmg) > 0.5 Then
        Dim effName As Variant
        For Each effName In CrashDict.keys()
            Call CrashDict.Item(effName).Item(ccnt).Reset(Tw2Px * x, Tw2Px * y, (dmg - pdmg))
        Next effName
        ccnt = ccnt + 1
        If ccnt > EFFECT_MAX_COUNT Then ccnt = 1
    End If
    pdmg = dmg
End Sub
Private Sub myCore_Move(x As Double, y As Double, veloc As Double, time As Long)
    Render time, x * Tw2Px, y * Tw2Px
End Sub
Public Sub Render(Optional time = 0, Optional x = -1, Optional y = -1)
    Dim dt As Long, r, i As Long
    Dim effName As Variant
    r = Rnd * 30
    dt = time - pTime
    pTime = time
    With GL
        .Clear GL_COLOR_BUFFER_BIT Or GL_DEPTH_BUFFER_BIT
        .MatrixMode GL_PROJECTION
        .LoadIdentity
        .Ortho2D 0, width, height, 0
        .MatrixMode GL_MODELVIEW
        .LoadIdentity
        .PushMatrix
            For Each effName In CrashDict.keys()
                For i = 1 To EFFECT_MAX_COUNT
                    Call CrashDict.Item(effName).Item(i).Render(x, y, dt)
                Next i
            Next effName
        .PopMatrix
        .PushMatrix
            For Each effName In MoveDict.keys()
                Call MoveDict.Item(effName).Render(x, y, dt)
            Next effName
        .PopMatrix
        .SwapBuffers
    End With
End Sub
Private Sub myCore_Stopped(x As Double, y As Double, time As Long)
    With GL
        .Clear GL_COLOR_BUFFER_BIT Or GL_DEPTH_BUFFER_BIT
        .SwapBuffers
    End With
End Sub
Private Sub Label1_Click()
    Unload Me
End Sub


Private Sub RenderFrame_Click()

End Sub

Private Sub UserForm_Activate()
    If GL Is Nothing Then
        Me.RenderFrame.BackColor = RGB(254, 254, 254)
        WindowFromAccessibleObject Me, hWnd
        WindowFromAccessibleObject Me.RenderFrame, rndhWnd
        #If Win64 Then
            style = GetWindowLongPtr(hWnd, GWL_STYLE)
        #Else
            style = GetWindowLong(hWnd, GWL_STYLE)
        #End If
        style = (style Or WS_THICKFRAME Or &H30000) And Not WS_CAPTION
        #If Win64 Then
            SetWindowLongPtr hWnd, GWL_STYLE, style
        #Else
            SetWindowLong hWnd, GWL_STYLE, style
        #End If
        #If Win64 Then
            exStyle = GetWindowLongPtr(hWnd, GWL_EXSTYLE)
        #Else
            exStyle = GetWindowLong(hWnd, GWL_EXSTYLE)
        #End If
        #If Win64 Then
            SetWindowLongPtr hWnd, GWL_EXSTYLE, exStyle Or WS_EX_LAYERED
        #Else
            SetWindowLong hWnd, GWL_EXSTYLE, exStyle Or WS_EX_LAYERED
        #End If
        'Force-enable SetLayeredWindowAttributes by rendering OpenGL
        'through GDI to make the background color transparent.”
        SetLayeredWindowAttributes hWnd, RGB(254, 254, 254), 0, LWA_COLORKEY
        ShowWindow hWnd, SW_SHOWMAXIMIZED
        Call GLInit
    End If
End Sub
Public Sub GLInit()
    If width <= 0 Then width = 1920: height = 1080
    Set GL = New OpenGL
    With GL
        .hWnd = rndhWnd
        .PaintStart
        .ClearColor 254 / 255, 254 / 255, 254 / 255, 0.5
        .Enable GL_DEPTH_TEST
        .Viewport 0, 0, width, height
        .Disable GL_LIGHTING
    End With
    DoEvents
    Set CrashDict = CreateObject("Scripting.Dictionary")
    Set MoveDict = CreateObject("Scripting.Dictionary")
    Dim efs As Variant, ek As Variant, ef As Variant, i As Long, ci
    On Error GoTo err:
    For Each ef In Array(glExplosion, glShockWave)
        ek = TypeName(ef)
        If Not CrashDict.exists(ek) Then
            Dim coll As Collection: Set coll = New Collection
            For i = 1 To EFFECT_MAX_COUNT
                Set ci = CallByName(ef, "CreateInstance", VbGet)
                Call ci.Init(GL)
                coll.Add ci
            Next i
            CrashDict.Add ek, coll
        End If
    Next ef
    For Each ef In Array(glMoveTrail)
        ek = TypeName(ef)
        If Not MoveDict.exists(ek) Then
            Set ci = CallByName(ef, "CreateInstance", VbGet)
            Call ci.Init(GL)
            MoveDict.Add ek, ci
        End If
    Next ef
err:
    Call Render
End Sub
Private Sub UserForm_Layout()
    Me.Label1.width = Me.width
    Me.RenderFrame.width = Me.width
    Me.RenderFrame.height = Me.height - Me.RenderFrame.top
End Sub
Public Property Get CreateInstance() As CFormPhysicsGLEffector
    Set CreateInstance = New CFormPhysicsGLEffector
End Property
Public Sub Init(ByRef core As CFormPhysics)
    Set myCore = core
    width = core.ScrWidth
    height = core.ScrHeight
    Tw2Px = 1 / core.Px2Tw
    ccnt = 1
    Me.Show vbModeless
End Sub
Public Sub Terminate()
On Error GoTo err
    Set MoveDict = Nothing
    Set CrashDict = Nothing
    Set myCore = Nothing
    GL.PaintEnd
err:
    Set GL = Nothing
    Unload Me
End Sub
