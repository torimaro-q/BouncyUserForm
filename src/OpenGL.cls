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
Private Declare PtrSafe Function GetDC Lib "user32" (ByVal hWnd As LongPtr) As LongPtr
Private Declare PtrSafe Function ReleaseDC Lib "user32" (ByVal hWnd As LongPtr, ByVal hdc As LongPtr) As Long
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
Private Declare PtrSafe Sub glVertex2d Lib "opengl32.dll" (ByVal A As Double, ByVal B As Double)
Private Declare PtrSafe Sub glVertex3d Lib "opengl32.dll" (ByVal X As Double, ByVal Y As Double, ByVal Z As Double)
Private Declare PtrSafe Sub glNormal3d Lib "opengl32.dll" (ByVal X As Double, ByVal Y As Double, ByVal Z As Double)
Private Declare PtrSafe Sub glVertex3dv Lib "opengl32.dll" (ByVal dPtr As LongPtr)
Private Declare PtrSafe Sub glNormal3dv Lib "opengl32.dll" (ByVal dPtr As LongPtr)
Private Declare PtrSafe Sub glBegin Lib "opengl32.dll" (ByVal Glenum As Long)
Private Declare PtrSafe Sub glEnd Lib "opengl32.dll" ()
Private Declare PtrSafe Sub glMultMatrixd Lib "opengl32.dll" (ByVal ptR As LongPtr)
Private Declare PtrSafe Sub glMatrixMode Lib "opengl32.dll" (ByVal Glenum As Long)
Private Declare PtrSafe Sub glPushMatrix Lib "opengl32.dll" ()
Private Declare PtrSafe Sub glPopMatrix Lib "opengl32.dll" ()
Private Declare PtrSafe Sub glLoadIdentity Lib "opengl32.dll" ()
Private Declare PtrSafe Sub glViewport Lib "opengl32.dll" (ByVal X As Long, ByVal Y As Long, ByVal width As Long, ByVal height As Long)
Private Declare PtrSafe Sub glEnable Lib "opengl32.dll" (ByVal Glenum As Long)
Private Declare PtrSafe Sub glDisable Lib "opengl32.dll" (ByVal Glenum As Long)
Private Declare PtrSafe Sub glPointSize Lib "opengl32.dll" (ByVal size As Single)
Private Declare PtrSafe Sub glLineWidth Lib "opengl32.dll" (ByVal width As Single)
Private Declare PtrSafe Sub glEnableClientState Lib "opengl32.dll" (ByVal Glenum As Long)
Private Declare PtrSafe Sub glDisableClientState Lib "opengl32.dll" (ByVal Glenum As Long)
Private Declare PtrSafe Sub glMaterialfv Lib "opengl32.dll" (ByVal Face As Long, ByVal pname As Long, ByVal Param As LongPtr)
Private Declare PtrSafe Sub glClear Lib "opengl32.dll" (ByVal Mask As Long)
Private Declare PtrSafe Sub glClearColor Lib "opengl32.dll" (ByVal red As Single, ByVal green As Single, ByVal blue As Single, ByVal alpha As Single)
Private Declare PtrSafe Sub glColorPointer Lib "opengl32.dll" (ByVal size As Long, ByVal Glenum As Long, ByVal Stride As Long, ByVal ptR As LongPtr)
Private Declare PtrSafe Sub glNormalPointer Lib "opengl32.dll" (ByVal Glenum As Long, ByVal Stride As Long, ByVal ptR As LongPtr)
Private Declare PtrSafe Sub glVertexPointer Lib "opengl32.dll" (ByVal size As Long, ByVal Glenum As Long, ByVal Stride As Long, ByVal ptR As LongPtr)
Private Declare PtrSafe Sub glDrawArrays Lib "opengl32.dll" (ByVal Glenum As Long, ByVal first As Long, ByVal size As Long)
Private Declare PtrSafe Sub glDrawElements Lib "opengl32.dll" (ByVal Glenum As Long, ByVal count As Long, ByVal GLenumT As Long, ByVal ptR As LongPtr)
Private Declare PtrSafe Sub glClipPlane Lib "opengl32.dll" (ByVal Glenum As Long, ByVal ptR As LongPtr)

Private Declare PtrSafe Function CreateFont Lib "gdi32" Alias "CreateFontA" (ByVal nHeight As Long, ByVal nWidth As Long, ByVal nEscapement As Long, ByVal nOrientation As Long, ByVal fnWeight As Long, ByVal IfdwItalic As Long, ByVal fdwUnderline As Long, ByVal fdwStrikeOut As Long, ByVal fdwCharSet As Long, ByVal fdwOutputPrecision As Long, ByVal fdwClipPrecision As Long, ByVal fdwQuality As Long, ByVal fdwPitchAndFamily As Long, ByVal lpszFace As String) As LongPtr
Private Declare PtrSafe Function wglUseFontBitmaps Lib "opengl32.dll" Alias "wglUseFontBitmapsA" (ByVal hdc As LongPtr, ByVal dfirst As Long, ByVal dcount As Long, ByVal dlistbase As Long) As Long
Private Declare PtrSafe Function wglUseFontBitmapsW Lib "opengl32.dll" (ByVal hdc As LongPtr, ByVal dfirst As Long, ByVal dcount As Long, ByVal dlistbase As Long) As Long

Private Declare PtrSafe Sub glRasterPos3d Lib "opengl32.dll" (ByVal X As Double, ByVal Y As Double, ByVal Z As Double)
Private Declare PtrSafe Sub glRasterPos2d Lib "opengl32.dll" (ByVal X As Double, ByVal Y As Double)
Private Declare PtrSafe Sub glTranslated Lib "opengl32.dll" (ByVal X As Double, ByVal Y As Double, ByVal Z As Double)
Private Declare PtrSafe Sub glScaled Lib "opengl32.dll" (ByVal X As Double, ByVal Y As Double, ByVal Z As Double)

Private Declare PtrSafe Function glIsList Lib "opengl32.dll" (ByVal Index As Long) As Byte
Private Declare PtrSafe Sub glNewList Lib "opengl32.dll" (ByVal Index As Long, ByVal Glenum As Long)
Private Declare PtrSafe Sub glEndList Lib "opengl32.dll" ()
Private Declare PtrSafe Sub glCallList Lib "opengl32.dll" (ByVal Index As Long)
Private Declare PtrSafe Sub glDeleteLists Lib "opengl32.dll" (ByVal Index As Long, ByVal size As Long)
Private Declare PtrSafe Sub glDeleteList Lib "opengl32.dll" (ByVal Index As Long)
Private Declare PtrSafe Sub glListBase Lib "opengl32.dll" (ByVal base As Long)
Private Declare PtrSafe Sub glCallLists Lib "opengl32.dll" (ByVal size As Long, ByVal vType As Long, ByVal ptR As LongPtr)

Private Declare PtrSafe Function wglGetProcAddress Lib "opengl32.dll" (ByVal lpszProc As String) As LongPtr
Private Declare PtrSafe Function wglSwapIntervalEXT Lib "opengl32.dll" (ByVal interval As Long) As LongPtr
Private Declare PtrSafe Function CallWindowProc Lib "user32.dll" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As LongPtr, ByVal hWnd As LongPtr, ByVal Msg As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As LongPtr
Private Declare PtrSafe Function DispCallFunc Lib "oleaut32" (ByVal pvInstance As LongPtr, ByVal oVft As LongPtr, ByVal cc As Long, ByVal vtReturn As Integer, ByVal cActuals As Long, ByRef prgvt As Integer, ByRef prgpvarg As LongPtr, ByRef pvargResult As Variant) As Long

Private Const CC_STDCALL As Long = 4
#If Win64 Then
    Private Const NULLPTR As LongPtr = 0
#Else
    Private Const NULLPTR As Long = 0
#End If
Private Enum VARENUM
    VT_I8 = 20
    VT_I4 = 3
    VT_EMPTY = 0
    VT_BYREF = &H4000
    #If Win64 Then
        VT_PTR2 = VT_I8
    #Else
        VT_PTR2 = VT_I4
    #End If
End Enum
Private attribs(6) As Long
Private ShaderBuf() As Byte
Private pfd As PIXELFORMATDESCRIPTOR
Public hformat As Long, hGLRC As LongPtr, hdc As LongPtr
Public Param As Object
Private pDict As Object, pArr(99) As LongPtr
Private myhwnd As LongPtr
Public Sub CreateFontList(Optional ByVal FontSize As Long = 30, Optional ByVal FontName As String = "Arial", Optional ByVal weight As Long = 200, Optional listbase As Long = 1000)
    Dim hFont As LongPtr, hOldFont As LongPtr
    If glIsList(listbase) = 1 Then glDeleteLists listbase, 255
    hFont = CreateFont(FontSize, 0, 0, 0, weight, 0, 0, 0, DEFAULT_CHARSET, OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS, ANTIALIASED_QUALITY, DEFAULT_PITCH Or FF_SCRIPT, FontName)
    If hFont = 0 Then Exit Sub
    hOldFont = SelectObject(hdc, hFont)
    wglUseFontBitmaps hdc, 0, 255, listbase
    Call SelectObject(hdc, hOldFont)
    Call DeleteObject(hFont)
    DoEvents
End Sub
Public Sub DisposeFontList()
    If glIsList(1000) = 1 Then glDeleteLists 1000, 10000
End Sub
Public Sub NewList(ByVal Index As Long, CompileMode As Glenum)
    Call glNewList(Index, CompileMode)
End Sub
Public Sub EndList()
    Call glEndList
End Sub
Public Sub CallList(ByVal Index As Long)
    Call glCallList(Index)
End Sub
Public Sub CallLists(ByVal Index As Long, ByVal vType As Glenum, ByVal ptR As LongPtr)
    Call glCallLists(Index, vType, ptR)
End Sub
Public Sub DeleteLists(ByVal Index As Long, size As Long)
    Call glDeleteLists(Index, size)
End Sub
Public Sub listbase(ByVal base As Long)
    Call glListBase(base)
End Sub
Public Function IsList(ByVal Index As Long) As Boolean
    IsList = glIsList(Index)
End Function
Public Sub RasterPos3d(ByVal X As Double, ByVal Y As Double, ByVal Z As Double)
    Call glRasterPos3d(X, Y, Z)
End Sub
Public Sub RasterPos2d(ByVal X As Double, ByVal Y As Double)
    Call glRasterPos2d(X, Y)
End Sub
Public Sub Scaled(ByVal X As Double, ByVal Y As Double, ByVal Z As Double)
    Call glScaled(X, Y, Z)
End Sub
Public Sub Translated(ByVal X As Double, ByVal Y As Double, ByVal Z As Double)
    Call glTranslated(X, Y, Z)
End Sub
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
Public Sub ColorPointer(ByVal size As Long, ByVal Gle As Glenum, ByVal Stride As Long, ByVal ptR As LongPtr)
    Call glColorPointer(size, Gle, Stride, ptR)
End Sub
Public Sub VertexPointer(ByVal size As Long, ByVal Gle As Glenum, ByVal Stride As Long, ByVal ptR As LongPtr)
    Call glVertexPointer(size, Gle, Stride, ptR)
End Sub
Public Sub NormalPointer(ByVal Gle As Glenum, ByVal Stride As Long, ByVal ptR As LongPtr)
    Call glNormalPointer(Gle, Stride, ptR)
End Sub
Public Sub DrawArrays(ByVal mode As Glenum, ByVal first As Long, ByVal size As Long)
    Call glDrawArrays(mode, first, size)
End Sub
Public Sub DrawElements(ByVal mode As Glenum, ByVal count As Long, ByVal Type1 As Glenum, ByVal ptR As LongPtr)
    Call glDrawElements(mode, count, Type1, ptR)
End Sub
Public Sub MultMatrixd(ByVal ptR As LongPtr)
    Call glMultMatrixd(ptR)
End Sub
Public Sub MakeCurrent(ByVal hdc As LongPtr, ByVal hGLRC As LongPtr)
    Call wglMakeCurrent(hdc, hGLRC)
End Sub
Public Sub SwapBuffers()
    Call SwapBuffersA(hdc)
End Sub
Public Sub Viewport(ByVal X As Long, ByVal Y As Long, ByVal width As Long, ByVal height As Long)
    Call glViewport(X, Y, width, height)
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
Public Sub Materialfv(ByVal GLe1 As Glenum, ByVal GLe2 As Glenum, ByVal Param As LongPtr)
    Call glMaterialfv(GLe1, GLe2, Param)
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
Public Sub Vertex2d(ByVal A As Double, ByVal B As Double)
    Call glVertex2d(A, B)
End Sub
Public Sub Vertex3d(ByVal X As Double, ByVal Y As Double, ByVal Z As Double)
    Call glVertex3d(X, Y, Z)
End Sub
Public Sub Normal3d(ByVal X As Double, ByVal Y As Double, ByVal Z As Double)
    Call glNormal3d(X, Y, Z)
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
Public Sub Clear(Optional ByVal Mask As Long = 16384)
    Call glClear(Mask)
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
'Ex
Private Function GetPointer(ByVal fname As String) As LongPtr
    With pDict
        If .exists(fname) = False Then
            Dim adr As LongPtr, idx As Long
            adr = wglGetProcAddress(fname)
            idx = .count
            .Add fname, idx
            pArr(idx) = adr
        End If
        GetPointer = pArr(.Item(fname))
    End With
End Function
Private Function CallStd(ByVal pFunc As LongPtr, ByVal rType As VARENUM, ParamArray args()) As Variant
    If pFunc = 0 Then Exit Function
    Dim n As Long, hr As Long, i As Long, ret As Variant: ret = vbEmpty
    If UBound(args) >= 0 Then n = UBound(args) Else n = -1
    If n > -1 Then
        Dim vars() As Variant: ReDim vars(n)
        Dim typs() As Integer: ReDim typs(n)
        Dim ptrs() As LongPtr: ReDim ptrs(n)
        For i = 0 To n
            vars(i) = args(i)
            typs(i) = VarType(args(i))
            #If Win64 Then
                ptrs(i) = VarPtr(vars(i))
            #Else
                ptrs(i) = VarPtr(vars(i))
            #End If
        Next i
        hr = DispCallFunc(NULLPTR, pFunc, CC_STDCALL, CInt(rType), CLng(n + 1), typs(0), ptrs(0), ret)
    Else
        hr = DispCallFunc(NULLPTR, pFunc, CC_STDCALL, CInt(rType), CLng(0), CInt(0), NULLPTR, ret)
    End If
    If hr = 0 Then CallStd = ret
End Function
Public Function SwapIntervalEXT(ByVal interval As Long) As Long
    SwapIntervalEXT = CallStd(GetPointer("wglSwapIntervalEXT"), VT_I4, interval)
End Function
Public Function CreateContextAttribsARB(ByVal hdc As LongPtr, ByVal hShare As LongPtr, ByVal attribList As LongPtr) As LongPtr
    CreateContextAttribsARB = CallStd(GetPointer("wglCreateContextAttribsARB"), VT_PTR2, hdc, hShare, attribList)
End Function
Public Function CreateShader(ByVal shaderType As Glenum) As Long
    CreateShader = CallStd(GetPointer("glCreateShader"), VT_I4, shaderType)
End Function
Public Sub ShaderSource(ByVal shader As Long, ByVal src As String)
    Dim srcLen As Long, pText As LongPtr, P As LongPtr
    ShaderBuf = StrConv(src & vbNullChar, vbFromUnicode)
    srcLen = UBound(ShaderBuf) + 1
    pText = VarPtr(ShaderBuf(0))
    P = GetPointer("glShaderSource")
    CallStd P, VT_EMPTY, shader, CLng(1&), VarPtr(pText), VarPtr(srcLen)
End Sub
Public Sub CompileShader(ByVal shader As Long)
    CallStd GetPointer("glCompileShader"), VT_EMPTY, CLng(shader)
End Sub
Public Function GetShaderiv(ByVal shader As Long, ByVal pname As Glenum) As Long
    Dim outVal As Long
    CallStd GetPointer("glGetShaderiv"), VT_EMPTY, CLng(shader), CLng(pname), VarPtr(outVal)
    GetShaderiv = outVal
End Function
Public Function GetProgramiv(ByVal program As Long, ByVal pname As Glenum) As Long
    Dim Value As Long
    CallStd GetPointer("glGetProgramiv"), VT_EMPTY, CLng(program), CLng(pname), VarPtr(Value)
    GetProgramiv = Value
End Function
Public Function CreateProgram() As Long
    CreateProgram = CallStd(GetPointer("glCreateProgram"), VT_I4)
End Function
Public Sub AttachShader(ByVal program As Long, ByVal shader As Long)
    CallStd GetPointer("glAttachShader"), VT_EMPTY, CLng(program), CLng(shader)
End Sub
Public Sub LinkProgram(ByVal program As Long)
    CallStd GetPointer("glLinkProgram"), VT_EMPTY, program
End Sub
Public Function GetShaderInfoLog(ByVal shader As Long) As String
    Dim bufSize As Long: bufSize = 4096
    Dim utf8() As Byte: ReDim utf8(bufSize - 1)
    Dim Length As Long
    CallStd GetPointer("glGetShaderInfoLog"), VT_EMPTY, shader, bufSize, VarPtr(Length), VarPtr(utf8(0))
    If Length <= 0 Then Exit Function
    GetShaderInfoLog = StrConv(LeftB$(utf8, Length), vbUnicode)
End Function
Public Function GetProgramInfoLog(ByVal program As Long) As String
    Dim bufSize As Long: bufSize = 4096
    Dim utf8() As Byte: ReDim utf8(bufSize - 1)
    Dim Length As Long
    CallStd GetPointer("glGetProgramInfoLog"), VT_EMPTY, program, bufSize, VarPtr(Length), VarPtr(utf8(0))
    If Length <= 0 Then Exit Function
    GetProgramInfoLog = StrConv(LeftB$(utf8, Length), vbUnicode)
End Function
'Init/Term
Public Sub PaintStart()
    Call CreateFontList(40, "Impact", 600, 1000)
    Call CreateFontList(80, "Impact", 600, 2000)
    Call CreateFontList(90, "Impact", 600, 2500)
    Call CreateFontList(160, "Impact", 600, 3000)
End Sub
Public Sub PaintEnd()
    Call SwapBuffersA(hdc)
    Call wglMakeCurrent(0, 0)
End Sub
Public Property Let hWnd(newPtr As LongPtr)
    myhwnd = newPtr
    Call Inithwnd
End Property
Public Property Get hWnd() As LongPtr
    hWnd = myhwnd
End Property
Private Sub Inithwnd()
    If hdc <> 0 Then Releasehwnd
    hdc = GetDC(myhwnd)
    hformat = ChoosePixelFormat(hdc, pfd)
    Call SetPixelFormat(hdc, hformat, pfd)
    Dim ohGLRC As LongPtr: ohGLRC = wglCreateContext(hdc)
    Call wglMakeCurrent(hdc, ohGLRC)
    Dim nhGLRC As LongPtr: nhGLRC = CreateContextAttribsARB(hdc, 0, VarPtr(attribs(0)))
    If nhGLRC <> 0 Then
        Call wglMakeCurrent(hdc, nhGLRC)
        Call wglDeleteContext(ohGLRC)
        hGLRC = nhGLRC
    Else
        hGLRC = ohGLRC
    End If
End Sub
Private Sub Releasehwnd()
    Call DisposeFontList
    Call wglMakeCurrent(0, 0)
    Call wglDeleteContext(hGLRC)
    Call ReleaseDC(myhwnd, hdc)
End Sub
Private Sub Class_Initialize()
    With pfd
        .nSize = 40
        .nVersion = 1
        .dwFlags = PFD_DRAW_TO_WINDOW Or PFD_SUPPORT_OPENGL Or PFD_SUPPORT_GDI 'GDI
        .cColorBits = 32
        .cDepthBits = 32
        .cStencilBits = 8
    End With
    attribs(0) = WGL_CONTEXT_MAJOR_VERSION_ARB: attribs(1) = 2
    attribs(2) = WGL_CONTEXT_MINOR_VERSION_ARB: attribs(3) = 1
    attribs(4) = WGL_CONTEXT_FLAGS_ARB: attribs(5) = 0
    attribs(6) = 0
    Set pDict = CreateObject("Scripting.Dictionary")
End Sub
Private Sub Class_Terminate()
    Call Releasehwnd
End Sub
