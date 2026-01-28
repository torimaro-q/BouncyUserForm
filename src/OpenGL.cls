VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "OpenGL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Declare PtrSafe Function GetDC Lib "user32" (ByVal hwnd As LongPtr) As LongPtr
Private Declare PtrSafe Function ReleaseDC Lib "user32" (ByVal hwnd As LongPtr, ByVal hdc As LongPtr) As Long
Private Declare PtrSafe Function SelectObject Lib "gdi32" (ByVal hdc As LongPtr, ByVal hObject As LongPtr) As LongPtr
Private Declare PtrSafe Function DeleteObject Lib "gdi32" (ByVal hObject As LongPtr) As Long
Private Declare PtrSafe Function ChoosePixelFormat Lib "gdi32" (ByVal hdc As LongPtr, ByRef pfd As PIXELFORMATDESCRIPTOR) As Long
Private Declare PtrSafe Function SetPixelFormat Lib "gdi32" (ByVal hdc As LongPtr, ByVal format As Long, ByRef ppfd As PIXELFORMATDESCRIPTOR) As Long
Private Declare PtrSafe Function SwapBuffersA Lib "gdi32" Alias "SwapBuffers" (ByVal hdc As LongPtr) As Long
Private Declare PtrSafe Sub gluPerspective Lib "glu32.dll" (ByVal fovy As Double, ByVal aspect As Double, ByVal zNear As Double, ByVal zFar As Double)
Private Declare PtrSafe Sub gluLookAt Lib "glu32.dll" (ByVal eyex As Double, ByVal eyey As Double, ByVal eyez As Double, ByVal centerx As Double, ByVal centery As Double, ByVal centerz As Double, ByVal upx As Double, ByVal upy As Double, ByVal upz As Double)
Private Declare PtrSafe Sub gluOrtho2D Lib "glu32.dll" (ByVal left As Double, ByVal right As Double, ByVal bottom As Double, ByVal top As Double)
Private Declare PtrSafe Sub glBlendFunc Lib "opengl32.dll" (ByVal GLenum1 As Long, ByVal GLenum2 As Long)
Private Declare PtrSafe Function wglCreateContext Lib "opengl32.dll" (ByVal hdc As LongPtr) As LongPtr
Private Declare PtrSafe Function wglMakeCurrent Lib "opengl32.dll" (ByVal hdc As LongPtr, ByVal hGLRC As LongPtr) As Long
Private Declare PtrSafe Function wglDeleteContext Lib "opengl32.dll" (ByVal hdc As LongPtr) As Long
Private Declare PtrSafe Sub glColor4f Lib "opengl32.dll" (ByVal red As Single, ByVal green As Single, ByVal blue As Single, ByVal alpha As Single)
Private Declare PtrSafe Sub glColor3d Lib "opengl32.dll" (ByVal red As Double, ByVal green As Double, ByVal blue As Double)
Private Declare PtrSafe Sub glVertex2d Lib "opengl32.dll" (ByVal a As Double, ByVal B As Double)
Private Declare PtrSafe Sub glVertex3d Lib "opengl32.dll" (ByVal x As Double, ByVal y As Double, ByVal z As Double)
Private Declare PtrSafe Sub glNormal3d Lib "opengl32.dll" (ByVal x As Double, ByVal y As Double, ByVal z As Double)
Private Declare PtrSafe Sub glVertex3dv Lib "opengl32.dll" (ByVal dPtr As LongPtr)
Private Declare PtrSafe Sub glNormal3dv Lib "opengl32.dll" (ByVal dPtr As LongPtr)
Private Declare PtrSafe Sub glBegin Lib "opengl32.dll" (ByVal Glenum As Long)
Private Declare PtrSafe Sub glEnd Lib "opengl32.dll" ()
Private Declare PtrSafe Sub glMultMatrixd Lib "opengl32.dll" (ByVal Ptr As LongPtr)
Private Declare PtrSafe Sub glMatrixMode Lib "opengl32.dll" (ByVal Glenum As Long)
Private Declare PtrSafe Sub glPushMatrix Lib "opengl32.dll" ()
Private Declare PtrSafe Sub glPopMatrix Lib "opengl32.dll" ()
Private Declare PtrSafe Sub glLoadIdentity Lib "opengl32.dll" ()
Private Declare PtrSafe Sub glViewport Lib "opengl32.dll" (ByVal x As Long, ByVal y As Long, ByVal width As Long, ByVal height As Long)
Private Declare PtrSafe Sub glEnable Lib "opengl32.dll" (ByVal Glenum As Long)
Private Declare PtrSafe Sub glDisable Lib "opengl32.dll" (ByVal Glenum As Long)
Private Declare PtrSafe Sub glPointSize Lib "opengl32.dll" (ByVal size As Single)
Private Declare PtrSafe Sub glLineWidth Lib "opengl32.dll" (ByVal width As Single)
Private Declare PtrSafe Sub glEnableClientState Lib "opengl32.dll" (ByVal Glenum As Long)
Private Declare PtrSafe Sub glDisableClientState Lib "opengl32.dll" (ByVal Glenum As Long)
Private Declare PtrSafe Sub glMaterialfv Lib "opengl32.dll" (ByVal Face As Long, ByVal pname As Long, ByVal param As LongPtr)
Private Declare PtrSafe Sub glClear Lib "opengl32.dll" (ByVal mask As Long)
Private Declare PtrSafe Sub glClearColor Lib "opengl32.dll" (ByVal red As Single, ByVal green As Single, ByVal blue As Single, ByVal alpha As Single)
Private Declare PtrSafe Sub glColorPointer Lib "opengl32.dll" (ByVal size As Long, ByVal Glenum As Long, ByVal Stride As Long, ByVal Ptr As LongPtr)
Private Declare PtrSafe Sub glNormalPointer Lib "opengl32.dll" (ByVal Glenum As Long, ByVal Stride As Long, ByVal Ptr As LongPtr)
Private Declare PtrSafe Sub glVertexPointer Lib "opengl32.dll" (ByVal size As Long, ByVal Glenum As Long, ByVal Stride As Long, ByVal Ptr As LongPtr)
Private Declare PtrSafe Sub glDrawArrays Lib "opengl32.dll" (ByVal Glenum As Long, ByVal First As Long, ByVal size As Long)
Private Declare PtrSafe Sub glClipPlane Lib "opengl32.dll" (ByVal Glenum As Long, ByVal Ptr As LongPtr)

Private pfd As PIXELFORMATDESCRIPTOR
Public hformat As Long, hGLRC As LongPtr, hdc As LongPtr
Private myhwnd As LongPtr
Public Sub PointSize(ByVal size As Single)
    Call glPointSize(size)
End Sub
Public Sub LineWidth(ByVal size As Single)
    Call glLineWidth(size)
End Sub
Public Sub EnableClientState(ByVal Gle As Glenum)
    Call glEnableClientState(Gle)
End Sub
Public Sub DisableClientState(ByVal Gle As Glenum)
    Call glDisableClientState(Gle)
End Sub
Public Sub ColorPointer(ByVal size As Long, ByVal Gle As Glenum, ByVal Stride As Long, ByVal Ptr As LongPtr)
    Call glColorPointer(size, Gle, Stride, Ptr)
End Sub
Public Sub VertexPointer(ByVal size As Long, ByVal Gle As Glenum, ByVal Stride As Long, ByVal Ptr As LongPtr)
    Call glVertexPointer(size, Gle, Stride, Ptr)
End Sub
Public Sub NormalPointer(ByVal Gle As Glenum, ByVal Stride As Long, ByVal Ptr As LongPtr)
    Call glNormalPointer(Gle, Stride, Ptr)
End Sub
Public Sub DrawArrays(ByVal Mode As Glenum, ByVal First As Long, ByVal size As Long)
    Call glDrawArrays(Mode, First, size)
End Sub
Public Sub MultMatrixd(ByVal Ptr As LongPtr)
    Call glMultMatrixd(Ptr)
End Sub
Public Sub PaintStart()
    Call wglMakeCurrent(hdc, hGLRC)
End Sub
Public Sub PaintEnd()
    Call SwapBuffersA(hdc)
    Call wglMakeCurrent(0, 0)
End Sub
Public Sub MakeCurrent(ByVal hdc As LongPtr, ByVal hGLRC As LongPtr)
    Call wglMakeCurrent(hdc, hGLRC)
End Sub
Public Sub SwapBuffers()
    Call SwapBuffersA(hdc)
End Sub
Public Sub Viewport(ByVal x As Long, ByVal y As Long, ByVal width As Long, ByVal height As Long)
    Call glViewport(x, y, width, height)
End Sub
Public Sub Enable(ByVal Gle As Glenum)
    Call glEnable(Gle)
End Sub
Public Sub Disable(ByVal Gle As Glenum)
    Call glDisable(Gle)
End Sub
Public Sub MatrixMode(ByVal Gle As Glenum)
    Call glMatrixMode(Gle)
End Sub
Public Sub Materialfv(ByVal GLe1 As Glenum, ByVal GLe2 As Glenum, ByVal param As LongPtr)
    Call glMaterialfv(GLe1, GLe2, param)
End Sub
Public Sub LoadIdentity()
    Call glLoadIdentity
End Sub
Public Sub PushMatrix()
    Call glPushMatrix
End Sub
Public Sub PopMatrix()
    Call glPopMatrix
End Sub
Public Sub Begin(ByVal Gle As Glenum)
    Call glBegin(Gle)
End Sub
Public Sub End1()
    Call glEnd
End Sub
Public Sub Color4f(ByVal red As Single, ByVal green As Single, ByVal blue As Single, ByVal alpha As Single)
    Call glColor4f(red, green, blue, alpha)
End Sub
Public Sub Color3d(ByVal red As Double, ByVal green As Double, ByVal blue As Double)
    Call glColor3d(red, green, blue)
End Sub
Public Sub Vertex2d(ByVal a As Double, ByVal B As Double)
    Call glVertex2d(a, B)
End Sub
Public Sub Vertex3d(ByVal x As Double, ByVal y As Double, ByVal z As Double)
    Call glVertex3d(x, y, z)
End Sub
Public Sub Normal3d(ByVal x As Double, ByVal y As Double, ByVal z As Double)
    Call glNormal3d(x, y, z)
End Sub
Public Sub Vertex3dv(ByVal Vec4Ptr As LongPtr)
    Call glVertex3dv(Vec4Ptr)
End Sub
Public Sub Normal3dv(ByVal Vec3Ptr As LongPtr)
    Call glNormal3dv(Vec3Ptr)
End Sub
Public Sub BlendFunc(ByVal GLe1 As Glenum, ByVal GLe2 As Glenum)
    Call glBlendFunc(GLe1, GLe2)
End Sub
Public Sub Clear(Optional ByVal mask As Long = 16384)
    Call glClear(mask)
End Sub
Public Sub ClipPlane(ByVal plane As Glenum, ByVal Vector4dPtr As LongPtr)
     Call glClipPlane(plane, Vector4dPtr)
End Sub
Public Sub ClearColor(ByVal red As Single, ByVal green As Single, ByVal blue As Single, ByVal alpha As Single)
    Call glClearColor(red, green, blue, alpha)
End Sub
Public Sub Perspective(ByVal fovy As Double, ByVal aspect As Double, ByVal zNear As Double, ByVal zFar As Double)
    Call gluPerspective(fovy, aspect, zNear, zFar)
End Sub
Public Sub Ortho2D(ByVal left As Double, ByVal right As Double, ByVal bottom As Double, ByVal top As Double)
    Call gluOrtho2D(left, right, bottom, top)
End Sub
Public Sub LookAt(ByVal eyex As Double, ByVal eyey As Double, ByVal eyez As Double, ByVal centerx As Double, ByVal centery As Double, ByVal centerz As Double, ByVal upx As Double, ByVal upy As Double, ByVal upz As Double)
    Call gluLookAt(eyex, eyey, eyez, centerx, centery, centerz, upx, upy, upz)
End Sub
'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Public Property Let hwnd(newPtr As LongPtr)
    myhwnd = newPtr
    Call Inithwnd
End Property
Public Property Get hwnd() As LongPtr
    hwnd = myhwnd
End Property
Private Sub Inithwnd()
    If hdc <> 0 Then Releasehwnd
    hdc = GetDC(myhwnd)
    hformat = ChoosePixelFormat(hdc, pfd)
    Call SetPixelFormat(hdc, hformat, pfd)
    hGLRC = wglCreateContext(hdc)
End Sub
Private Sub Releasehwnd()
    Call wglMakeCurrent(0, 0)
    Call wglDeleteContext(hGLRC)
    Call ReleaseDC(myhwnd, hdc)
End Sub
Private Sub Class_Initialize()
    With pfd
        .nSize = 40 ' LenB(pfd)
        .nVersion = 1
        .dwFlags = PFD_DRAW_TO_WINDOW Or PFD_SUPPORT_OPENGL Or PFD_SUPPORT_GDI 'GDI
        .cColorBits = 32
        .cRedBits = 0
        .cRedShift = 0
        .cGreenBits = 0
        .cGreenShift = 0
        .cBlueBits = 0
        .cBlueShift = 0
        .cAlphaBits = 0
        .cAlphaShift = 0
        .cAccumBits = 0
        .cAccumRedBits = 0
        .cAccumGreenBits = 0
        .cAccumBlueBits = 0
        .cAccumAlphaBits = 0
        .cDepthBits = 32
        .cStencilBits = 8
        .cAuxBuffers = 0
        .iLayerType = 0 'PFD_MAIN_PLANE
        .bReserved = 0
        .dwLayerMask = 0
        .dwVisibleMask = 0
        .dwDamageMask = 0
    End With
End Sub
Private Sub Class_Terminate()
    Call Releasehwnd
End Sub


