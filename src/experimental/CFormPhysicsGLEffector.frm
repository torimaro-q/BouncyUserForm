VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} CFormPhysicsGLEffector 
   Caption         =   "CFormPhysicsEffecter"
   ClientHeight    =   9570
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
Private GL As OpenGL
Private hWnd As LongPtr, exStyle As LongPtr, style As LongPtr
Private rndhWnd As LongPtr
Private Width As Double, Height As Double
Private WithEvents myCore As CFormPhysics, ptime As Long, pdmg As Double, Tw2Px As Double, ecnt As Long
Attribute myCore.VB_VarHelpID = -1
Private effects As Collection
Private Sub myCore_Crash(x As Double, y As Double, dmg As Double, time As Long)
    If (dmg - pdmg) > 1 Then
        Debug.Print effects.Count
        If effects.Count <= 3 Then
            Dim ef As glExplosion
            Set ef = New glExplosion
            ef.Init GL
            effects.Add ef
        End If
        Call effects.Item(ecnt).Reset(Tw2Px * x, Tw2Px * y)
        ecnt = ecnt + 1
        If ecnt > 3 Then ecnt = 1
    End If
    pdmg = dmg
End Sub
Private Sub myCore_Move(x As Double, y As Double, veloc As Double, time As Long)
    Render time, x * Tw2Px, y * Tw2Px
End Sub
Public Sub Render(Optional time = 0, Optional x = -1, Optional y = -1)
    Dim dt As Long, r, i
    r = Rnd * 30
    dt = time - ptime
    ptime = time
    With GL
        .Clear GL_COLOR_BUFFER_BIT Or GL_DEPTH_BUFFER_BIT
        .MatrixMode GL_PROJECTION
        .LoadIdentity
        .Ortho2D 0, Width, Height, 0
        .MatrixMode GL_MODELVIEW
        .LoadIdentity
        .PushMatrix
        For i = 1 To effects.Count
            effects.Item(i).Render x, y, dt
        Next i
        .PopMatrix
        .SwapBuffers
    End With
End Sub
Private Sub myCore_Started(x As Double, y As Double, time As Long)
    ecnt = 1
End Sub
Private Sub myCore_Stopped(x As Double, y As Double, time As Long)
    With GL
        .Clear GL_COLOR_BUFFER_BIT Or GL_DEPTH_BUFFER_BIT
        .SwapBuffers
    End With
End Sub
Public Property Get CreateInstance(Optional ByRef core As CFormPhysics = Nothing) As CFormPhysicsGLEffector
    Dim ch As CFormPhysicsGLEffector: Set ch = New CFormPhysicsGLEffector
    If Not core Is Nothing Then ch.Init core
    Set CreateInstance = ch
End Property
Public Sub Init(ByRef core As CFormPhysics)
    Set myCore = core
    Width = core.ScrWidth
    Height = core.ScrHeight
    Tw2Px = 1 / core.Px2Tw
    Me.Show vbModeless
End Sub
Private Sub Label1_Click()
    Unload Me
End Sub
Private Sub UserForm_Activate()
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
        exStyle = GetWindowLongPtr(hWnd, GWL_STYLE)
    #Else
        exStyle = GetWindowLong(hWnd, GWL_STYLE)
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
End Sub
Public Sub GLInit()
    If Width <= 0 Then
        Width = 1920: Height = 1080
    End If
    Set GL = New OpenGL
    DoEvents
    With GL
        .hWnd = rndhWnd
        .PaintStart
        .ClearColor 254 / 255, 254 / 255, 254 / 255, 0.5
        .Enable GL_DEPTH_TEST
        .Viewport 0, 0, Width, Height
        .Disable GL_LIGHTING
    End With
    DoEvents
    Set effects = New Collection
    Call Render
End Sub
Private Sub UserForm_Layout()
    Me.RenderFrame.Width = Me.Width
    Me.Label1.Width = Me.Width
    Me.RenderFrame.Height = Me.Height - Me.RenderFrame.top
End Sub
Private Sub UserForm_Terminate()
    Dim effect
    For Each effect In effects
        effect.Terminate
    Next effect
    Set myCore = Nothing
    GL.PaintEnd
    Unload Me
End Sub
Public Sub Terminate()
    Call UserForm_Terminate
End Sub
