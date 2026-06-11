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
Private Declare PtrSafe Sub glVertex2f Lib "opengl32.dll" (ByVal A As Single, ByVal B As Single)
Private Declare PtrSafe Sub glVertex3d Lib "opengl32.dll" (ByVal X As Double, ByVal Y As Double, ByVal Z As Double)
Private Declare PtrSafe Sub glNormal3d Lib "opengl32.dll" (ByVal X As Double, ByVal Y As Double, ByVal Z As Double)
Private Declare PtrSafe Sub glVertex3dv Lib "opengl32.dll" (ByVal dPtr As LongPtr)
Private Declare PtrSafe Sub glNormal3dv Lib "opengl32.dll" (ByVal dPtr As LongPtr)
Private Declare PtrSafe Sub glBegin Lib "opengl32.dll" (ByVal GLenum As Long)
Private Declare PtrSafe Sub glEnd Lib "opengl32.dll" ()
Private Declare PtrSafe Sub glMultMatrixd Lib "opengl32.dll" (ByVal ptR As LongPtr)
Private Declare PtrSafe Sub glMatrixMode Lib "opengl32.dll" (ByVal GLenum As Long)
Private Declare PtrSafe Sub glPushMatrix Lib "opengl32.dll" ()
Private Declare PtrSafe Sub glPopMatrix Lib "opengl32.dll" ()
Private Declare PtrSafe Sub glLoadIdentity Lib "opengl32.dll" ()
Private Declare PtrSafe Sub glViewport Lib "opengl32.dll" (ByVal X As Long, ByVal Y As Long, ByVal width As Long, ByVal height As Long)
Private Declare PtrSafe Sub glEnable Lib "opengl32.dll" (ByVal GLenum As Long)
Private Declare PtrSafe Sub glDisable Lib "opengl32.dll" (ByVal GLenum As Long)
Private Declare PtrSafe Sub glPointSize Lib "opengl32.dll" (ByVal size As Single)
Private Declare PtrSafe Sub glLineWidth Lib "opengl32.dll" (ByVal width As Single)
Private Declare PtrSafe Sub glEnableClientState Lib "opengl32.dll" (ByVal GLenum As Long)
Private Declare PtrSafe Sub glDisableClientState Lib "opengl32.dll" (ByVal GLenum As Long)
Private Declare PtrSafe Sub glMaterialfv Lib "opengl32.dll" (ByVal Face As Long, ByVal pname As Long, ByVal Param As LongPtr)
Private Declare PtrSafe Sub glClear Lib "opengl32.dll" (ByVal Mask As Long)
Private Declare PtrSafe Sub glClearColor Lib "opengl32.dll" (ByVal red As Single, ByVal green As Single, ByVal blue As Single, ByVal alpha As Single)
Private Declare PtrSafe Sub glColorPointer Lib "opengl32.dll" (ByVal size As Long, ByVal GLenum As Long, ByVal Stride As Long, ByVal ptR As LongPtr)
Private Declare PtrSafe Sub glNormalPointer Lib "opengl32.dll" (ByVal GLenum As Long, ByVal Stride As Long, ByVal ptR As LongPtr)
Private Declare PtrSafe Sub glVertexPointer Lib "opengl32.dll" (ByVal size As Long, ByVal GLenum As Long, ByVal Stride As Long, ByVal ptR As LongPtr)
Private Declare PtrSafe Sub glDrawArrays Lib "opengl32.dll" (ByVal GLenum As Long, ByVal first As Long, ByVal size As Long)
Private Declare PtrSafe Sub glDrawElements Lib "opengl32.dll" (ByVal GLenum As Long, ByVal count As Long, ByVal GlenumT As Long, ByVal ptR As LongPtr)
Private Declare PtrSafe Sub glClipPlane Lib "opengl32.dll" (ByVal GLenum As Long, ByVal ptR As LongPtr)
Private Declare PtrSafe Function CreateFont Lib "gdi32" Alias "CreateFontA" (ByVal nHeight As Long, ByVal nWidth As Long, ByVal nEscapement As Long, ByVal nOrientation As Long, ByVal fnWeight As Long, ByVal IfdwItalic As Long, ByVal fdwUnderline As Long, ByVal fdwStrikeOut As Long, ByVal fdwCharSet As Long, ByVal fdwOutputPrecision As Long, ByVal fdwClipPrecision As Long, ByVal fdwQuality As Long, ByVal fdwPitchAndFamily As Long, ByVal lpszFace As String) As LongPtr
Private Declare PtrSafe Function wglUseFontBitmaps Lib "opengl32.dll" Alias "wglUseFontBitmapsA" (ByVal hdc As LongPtr, ByVal dfirst As Long, ByVal dcount As Long, ByVal dlistbase As Long) As Long
Private Declare PtrSafe Function wglUseFontBitmapsW Lib "opengl32.dll" (ByVal hdc As LongPtr, ByVal dfirst As Long, ByVal dcount As Long, ByVal dlistbase As Long) As Long
Private Declare PtrSafe Sub glRasterPos3d Lib "opengl32.dll" (ByVal X As Double, ByVal Y As Double, ByVal Z As Double)
Private Declare PtrSafe Sub glRasterPos2d Lib "opengl32.dll" (ByVal X As Double, ByVal Y As Double)
Private Declare PtrSafe Sub glTranslated Lib "opengl32.dll" (ByVal X As Double, ByVal Y As Double, ByVal Z As Double)
Private Declare PtrSafe Sub glScaled Lib "opengl32.dll" (ByVal X As Double, ByVal Y As Double, ByVal Z As Double)
Private Declare PtrSafe Function glIsList Lib "opengl32.dll" (ByVal Index As Long) As Byte
Private Declare PtrSafe Sub glNewList Lib "opengl32.dll" (ByVal Index As Long, ByVal GLenum As Long)
Private Declare PtrSafe Sub glEndList Lib "opengl32.dll" ()
Private Declare PtrSafe Sub glCallList Lib "opengl32.dll" (ByVal Index As Long)
Private Declare PtrSafe Sub glDeleteLists Lib "opengl32.dll" (ByVal Index As Long, ByVal size As Long)
Private Declare PtrSafe Sub glDeleteList Lib "opengl32.dll" (ByVal Index As Long)
Private Declare PtrSafe Sub glListBase Lib "opengl32.dll" (ByVal base As Long)
Private Declare PtrSafe Sub glCallLists Lib "opengl32.dll" (ByVal size As Long, ByVal vType As Long, ByVal ptR As LongPtr)
Private Declare PtrSafe Sub glPixelStorei Lib "opengl32.dll" (ByVal GLenum As Long, ByVal Param As Long)
Private Declare PtrSafe Sub glTexParameteri Lib "opengl32.dll" (ByVal GLenum1 As Long, ByVal GLenum2 As Long, ByVal Param As Long)
Private Declare PtrSafe Sub glTexCoordPointer Lib "opengl32.dll" (ByVal size As Long, ByVal GLenum As Long, ByVal Stride As Long, ByVal ptR As LongPtr)
Private Declare PtrSafe Sub glTexImage2D Lib "opengl32.dll" (ByVal GLenum As Long, ByVal level As Long, ByVal internalformat As Long, _
                                                             ByVal width As Long, ByVal height As Long, ByVal border As Long, _
                                                             ByVal GlenumF As Long, ByVal GlenumT As Long, ByVal PixelPtr As LongPtr)
Private Declare PtrSafe Sub gluBuild2DMipmaps Lib "glu32.dll" (ByVal GLenum1 As Long, ByVal GLenum2 As Long, _
                                                              ByVal width As Long, ByVal height As Long, _
                                                              ByVal GLenum3 As Long, ByVal GLenum4 As Long, ByVal PixelPtr As LongPtr)
Private Declare PtrSafe Sub glTexCoord2d Lib "opengl32.dll" (ByVal S As Double, ByVal t As Double)
Private Declare PtrSafe Sub glTexCoord2f Lib "opengl32.dll" (ByVal S As Single, ByVal t As Single)
Private Declare PtrSafe Sub glGenTextures Lib "opengl32.dll" (ByVal n As Long, ByVal textures As LongPtr)
Private Declare PtrSafe Sub glBindTexture Lib "opengl32.dll" (ByVal target As Long, ByVal texture As Long)
Private Declare PtrSafe Sub glDeleteTextures Lib "opengl32.dll" (ByVal n As Long, ByVal textures As LongPtr)



Private Declare PtrSafe Function wglGetProcAddress Lib "opengl32.dll" (ByVal lpszProc As String) As LongPtr
Private Declare PtrSafe Function wglSwapIntervalEXT Lib "opengl32.dll" (ByVal interval As Long) As LongPtr
Private Declare PtrSafe Function CallWindowProc Lib "user32.dll" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As LongPtr, ByVal hWnd As LongPtr, ByVal msg As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As LongPtr
Private Declare PtrSafe Function DispCallFunc Lib "oleaut32" (ByVal pvInstance As LongPtr, ByVal oVft As LongPtr, ByVal cc As Long, ByVal vtReturn As Integer, ByVal cActuals As Long, ByRef prgvt As Integer, ByRef prgpvarg As LongPtr, ByRef pvargResult As Variant) As Long
Private Const CC_STDCALL As Long = 4
#If Win64 Then
    Private Const NULLPTR As LongPtr = 0
#Else
    Private Const NULLPTR As Long = 0
#End If
Private Enum VARENUM
    VT_EMPTY = 0
    VT_NULL = 1
    VT_I2 = 2
    VT_I4 = 3
    VT_I8 = 20
    VT_R4 = 4
    VT_R8 = 5
    VT_BOOL = 11
    VT_BYREF = &H4000
    #If Win64 Then
        VT_PTR2 = VT_I8
    #Else
        VT_PTR2 = VT_I4
    #End If
End Enum
Private attribs(6) As Long, pfd As PIXELFORMATDESCRIPTOR
Public hformat As Long, hGLRC As LongPtr, hdc As LongPtr, Param As Object
Private pDict As Object, ShaderBuf() As Byte, myhwnd As LongPtr
Private Type CallCache
    pfunc As LongPtr
    rType As VARENUM
    n As Long
    typs() As Integer
    vars() As Variant
    ptrs() As LongPtr
End Type
Private cCache(300) As CallCache
Private Function CallFp(ByRef rtId As Long, ByRef fname As String, ByVal rType As VARENUM, ParamArray args()) As Variant
    With pDict
        Dim i As Long
        If .exists(fname) Then
            rtId = .Item(fname)
        Else
            rtId = .count
            .Add fname, rtId
            With cCache(rtId)
                .pfunc = wglGetProcAddress(fname)
                .rType = CInt(rType)
                .n = UBound(args)
                If .n >= 0 Then
                    ReDim .vars(.n): ReDim .typs(.n): ReDim .ptrs(.n)
                    For i = 0 To .n
                        .typs(i) = VarType(args(i))
                    Next i
                End If
            End With
        End If
    End With
    With cCache(rtId)
        If .pfunc = 0 Then Exit Function
        Dim hr As Long, ret As Variant: ret = vbEmpty
        If .n >= 0 Then
            For i = 0 To .n
                .vars(i) = args(i)
                .ptrs(i) = VarPtr(.vars(i))
            Next i
            hr = DispCallFunc(NULLPTR, .pfunc, CC_STDCALL, .rType, CLng(.n + 1), .typs(0), .ptrs(0), ret)
        Else
            hr = DispCallFunc(NULLPTR, .pfunc, CC_STDCALL, .rType, CLng(0), CInt(0), NULLPTR, ret)
        End If
        If hr = 0 Then CallFp = ret
    End With
End Function
Public Sub BindBuffer(ByVal target As Long, ByVal buf As Long)
    Static fid As Long
    If fid > 0 Then
        With cCache(fid)
            .vars(0) = target: .vars(1) = buf
            .ptrs(0) = VarPtr(.vars(0)): .ptrs(1) = VarPtr(.vars(1))
            Call DispCallFunc(NULLPTR, .pfunc, CC_STDCALL, .rType, 2, .typs(0), .ptrs(0), vbEmpty)
        End With
    Else
        CallFp fid, "glBindBuffer", VT_EMPTY, target, buf
    End If
End Sub
Public Sub BufferSubData(ByVal target As Long, ByVal offset As Long, ByVal size As Long, ByVal data As LongPtr)
    Static fid As Long
    If fid > 0 Then
        With cCache(fid)
            .vars(0) = target: .vars(1) = offset: .vars(2) = size: .vars(3) = data
            .ptrs(0) = VarPtr(.vars(0)): .ptrs(1) = VarPtr(.vars(1)): .ptrs(2) = VarPtr(.vars(2)): .ptrs(3) = VarPtr(.vars(3))
            Call DispCallFunc(NULLPTR, .pfunc, CC_STDCALL, .rType, 4, .typs(0), .ptrs(0), vbEmpty)
        End With
    Else
        CallFp fid, "glBufferSubData", VT_EMPTY, target, offset, size, data
    End If
End Sub
Public Sub EnableVertexAttribArray(ByVal idx As Long)
    Static fid As Long
    If fid > 0 Then
        With cCache(fid)
            .vars(0) = idx
            .ptrs(0) = VarPtr(.vars(0))
            Call DispCallFunc(NULLPTR, .pfunc, CC_STDCALL, .rType, 1, .typs(0), .ptrs(0), vbEmpty)
        End With
    Else
        CallFp fid, "glEnableVertexAttribArray", VT_EMPTY, idx
    End If
End Sub
Public Sub VertexAttribPointer(ByVal idx As Long, ByVal size As Long, ByVal typ As GLenum, ByVal normalized As Long, ByVal Stride As Long, ByVal ptR As LongPtr)
    Static fid As Long
    If fid > 0 Then
        With cCache(fid)
            .vars(0) = idx: .vars(1) = size: .vars(2) = typ: .vars(3) = normalized: .vars(4) = Stride: .vars(5) = ptR
            .ptrs(0) = VarPtr(.vars(0)): .ptrs(1) = VarPtr(.vars(1)): .ptrs(2) = VarPtr(.vars(2))
            .ptrs(3) = VarPtr(.vars(3)): .ptrs(4) = VarPtr(.vars(4)): .ptrs(5) = VarPtr(.vars(5))
            Call DispCallFunc(NULLPTR, .pfunc, CC_STDCALL, .rType, 6, .typs(0), .ptrs(0), vbEmpty)
        End With
    Else
        CallFp fid, "glVertexAttribPointer", VT_EMPTY, idx, size, typ, normalized, Stride, ptR
    End If
End Sub
Public Sub DisableVertexAttribArray(ByVal idx As Long)
    Static fid As Long
    If fid > 0 Then
        With cCache(fid)
            .vars(0) = idx
            .ptrs(0) = VarPtr(.vars(0))
            Call DispCallFunc(NULLPTR, .pfunc, CC_STDCALL, .rType, 1, .typs(0), .ptrs(0), vbEmpty)
        End With
    Else
        CallFp fid, "glDisableVertexAttribArray", VT_EMPTY, idx
    End If
End Sub
Public Sub UseProgram(ByVal program As Long)
    Static fid As Long
    If fid > 0 Then
        With cCache(fid)
            .vars(0) = program
            .ptrs(0) = VarPtr(.vars(0))
            Call DispCallFunc(NULLPTR, .pfunc, CC_STDCALL, .rType, 1, .typs(0), .ptrs(0), vbEmpty)
        End With
    Else
        CallFp fid, "glUseProgram", VT_EMPTY, program
    End If
End Sub
Public Sub Uniform1f(ByVal loc As Long, ByVal v0 As Single)
    Static fid As Long
    If fid > 0 Then
        With cCache(fid)
            .vars(0) = loc
            .vars(1) = v0
            .ptrs(0) = VarPtr(.vars(0))
            .ptrs(1) = VarPtr(.vars(1))
            Call DispCallFunc(NULLPTR, .pfunc, CC_STDCALL, .rType, 2, .typs(0), .ptrs(0), vbEmpty)
        End With
    Else
        CallFp 0, "glUniform1f", VT_EMPTY, CLng(loc), CSng(v0)
    End If
End Sub
Public Sub Uniform1i(ByVal loc As Long, ByVal v0 As Long)
    Static fid As Long
    If fid > 0 Then
        With cCache(fid)
            .vars(0) = loc
            .vars(1) = v0
            .ptrs(0) = VarPtr(.vars(0))
            .ptrs(1) = VarPtr(.vars(1))
            Call DispCallFunc(NULLPTR, .pfunc, CC_STDCALL, .rType, 2, .typs(0), .ptrs(0), vbEmpty)
        End With
    Else
        CallFp 0, "glUniform1i", VT_EMPTY, CLng(loc), CSng(v0)
    End If
End Sub
Public Function GetUniformLocation(ByVal program As Long, ByVal Name As String) As Long
    Dim B() As Byte: B = StrConv(Name & vbNullChar, vbFromUnicode)
    GetUniformLocation = CallFp(0, "glGetUniformLocation", VT_I4, CLng(program), VarPtr(B(0)))
End Function
Public Sub PushMatrix()
    Call glPushMatrix
End Sub
Public Sub PopMatrix()
    Call glPopMatrix
End Sub
Public Sub Enable(ByVal Gle As GLenum)
    Call glEnable(Gle)
End Sub
Public Sub Disable(ByVal Gle As GLenum)
    Call glDisable(Gle)
End Sub
Public Sub EnableClientState(ByVal Gle As GLenum)
    Call glEnableClientState(Gle)
End Sub
Public Sub DisableClientState(ByVal Gle As GLenum)
    Call glDisableClientState(Gle)
End Sub
Public Sub NewList(ByVal Index As Long, CompileMode As GLenum)
    Call glNewList(Index, CompileMode)
End Sub
Public Sub EndList()
    Call glEndList
End Sub
Public Sub CallList(ByVal Index As Long)
    Call glCallList(Index)
End Sub
Public Sub CallLists(ByVal Index As Long, ByVal vType As GLenum, ByVal ptR As LongPtr)
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
Public Sub PixelStorei(ByVal pname As GLenum, ByVal Param As Long)
    Call glPixelStorei(pname, Param)
End Sub
Public Sub TexParameteri(ByVal target As GLenum, ByVal pname As GLenum, ByVal Param As Long)
    Call glTexParameteri(target, pname, Param)
End Sub
Public Sub TexImage2D(ByVal target As GLenum, ByVal level As Long, ByVal internalformat As Long, ByVal width As Long, ByVal height As Long, ByVal border As Long, ByVal format As GLenum, ByVal pType As GLenum, ByVal PixelPtr As LongPtr)
    Call glTexImage2D(target, level, internalformat, width, height, border, format, pType, PixelPtr)
End Sub
Public Sub Build2DMipmaps(ByVal target As GLenum, ByVal internalformat As GLenum, ByVal width As Long, ByVal height As Long, ByVal format As GLenum, ByVal pType As GLenum, ByVal PixelPtr As LongPtr)
    Call gluBuild2DMipmaps(target, internalformat, width, height, format, pType, PixelPtr)
End Sub
Public Sub TexCoordPointer(ByVal size As Long, ByVal Gle As GLenum, ByVal Stride As Long, ByVal ptR As LongPtr)
    Call glTexCoordPointer(size, Gle, Stride, ptR)
End Sub
Public Sub TexCoord2d(ByVal S As Double, ByVal t As Double)
    Call glTexCoord2d(S, t)
End Sub
Public Sub TexCoord2f(ByVal S As Single, ByVal t As Single)
    Call glTexCoord2f(S, t)
End Sub
Public Sub GenTextures(ByVal n As Long, ByVal textures As LongPtr)
    glGenTextures n, textures
End Sub
Public Sub BindTexture(ByVal target As GLenum, ByRef textures As Long)
    glBindTexture target, textures
End Sub
Public Sub DeleteTextures(ByVal n As Long, ByVal textures As LongPtr)
    glDeleteTextures n, textures
End Sub
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
Public Sub ColorPointer(ByVal size As Long, ByVal Gle As GLenum, ByVal Stride As Long, ByVal ptR As LongPtr)
    Call glColorPointer(size, Gle, Stride, ptR)
End Sub
Public Sub VertexPointer(ByVal size As Long, ByVal Gle As GLenum, ByVal Stride As Long, ByVal ptR As LongPtr)
    Call glVertexPointer(size, Gle, Stride, ptR)
End Sub
Public Sub NormalPointer(ByVal Gle As GLenum, ByVal Stride As Long, ByVal ptR As LongPtr)
    Call glNormalPointer(Gle, Stride, ptR)
End Sub
Public Sub DrawArrays(ByVal mode As GLenum, ByVal first As Long, ByVal size As Long)
    Call glDrawArrays(mode, first, size)
End Sub
Public Sub DrawElements(ByVal mode As GLenum, ByVal count As Long, ByVal Type1 As GLenum, ByVal ptR As LongPtr)
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
Public Sub MatrixMode(ByVal Gle As GLenum)
    Call glMatrixMode(Gle)
End Sub
Public Sub Materialfv(ByVal GLe1 As GLenum, ByVal GLe2 As GLenum, ByVal Param As LongPtr)
    Call glMaterialfv(GLe1, GLe2, Param)
End Sub
Public Sub LoadIdentity()
    Call glLoadIdentity
End Sub
Public Sub Begin(ByVal Gle As GLenum)
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
Public Sub Vertex2f(ByVal A As Single, ByVal B As Single)
    Call glVertex2f(A, B)
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
Public Sub BlendFunc(ByVal GLe1 As GLenum, ByVal GLe2 As GLenum)
    Call glBlendFunc(GLe1, GLe2)
End Sub
Public Sub Clear(Optional ByVal Mask As Long = 16384)
    Call glClear(Mask)
End Sub
Public Sub ClipPlane(ByVal plane As GLenum, ByVal Vector4dPtr As LongPtr)
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
Public Function CheckFramebufferStatus(ByVal target As GLenum) As Long
    CheckFramebufferStatus = CallFp(0, "glCheckFramebufferStatus", VT_I4, target)
End Function
Public Function SwapIntervalEXT(ByVal interval As Long) As Long
    SwapIntervalEXT = CallFp(0, "wglSwapIntervalEXT", VT_I4, interval)
End Function
Public Function CreateContextAttribsARB(ByVal hdc As LongPtr, ByVal hShare As LongPtr, ByVal attribList As LongPtr) As LongPtr
    CreateContextAttribsARB = CallFp(0, "wglCreateContextAttribsARB", VT_PTR2, hdc, hShare, attribList)
End Function
Public Function CreateShader(ByVal shaderType As GLenum) As Long
    CreateShader = CallFp(0, "glCreateShader", VT_I4, shaderType)
End Function
Public Sub ShaderSource(ByVal shader As Long, ByVal src As String)
    Dim srcLen As Long, pText As LongPtr, P As LongPtr
    ShaderBuf = StrConv(src & vbNullChar, vbFromUnicode)
    srcLen = UBound(ShaderBuf) + 1
    pText = VarPtr(ShaderBuf(0))
    CallFp 0, "glShaderSource", VT_EMPTY, shader, CLng(1&), VarPtr(pText), VarPtr(srcLen)
End Sub
Public Sub CompileShader(ByVal shader As Long)
    CallFp 0, "glCompileShader", VT_EMPTY, CLng(shader)
End Sub
Public Function GetShaderiv(ByVal shader As Long, ByVal pname As GLenum) As Long
    Dim outVal As Long
    CallFp 0, "glGetShaderiv", VT_EMPTY, CLng(shader), CLng(pname), VarPtr(outVal)
    GetShaderiv = outVal
End Function
Public Function GetProgramiv(ByVal program As Long, ByVal pname As GLenum) As Long
    Dim Value As Long
    CallFp 0, "glGetProgramiv", VT_EMPTY, CLng(program), CLng(pname), VarPtr(Value)
    GetProgramiv = Value
End Function
Public Function CreateProgram() As Long
    CreateProgram = CallFp(0, "glCreateProgram", VT_I4)
End Function
Public Sub DeleteProgram(ByVal program As Long)
    CallFp 0, "glDeleteProgram", VT_EMPTY, program
End Sub
Public Sub AttachShader(ByVal program As Long, ByVal shader As Long)
    CallFp 0, "glAttachShader", VT_EMPTY, CLng(program), CLng(shader)
End Sub
Public Sub LinkProgram(ByVal program As Long)
    CallFp 0, "glLinkProgram", VT_EMPTY, program
End Sub
Public Sub GenFramebuffersEXT(ByVal n As Long, ByVal pIds As LongPtr)
    CallFp 0, "glGenFramebuffersEXT", VT_EMPTY, n, pIds
End Sub
Public Sub BindFramebufferEXT(ByVal target As Long, ByVal fbo As Long)
    CallFp 0, "glBindFramebufferEXT", VT_EMPTY, target, fbo
End Sub
Public Sub FramebufferTexture2DEXT(ByVal target As Long, ByVal attachment As Long, ByVal textarget As Long, ByVal tex As Long, ByVal level As Long)
    CallFp 0, "glFramebufferTexture2DEXT", VT_EMPTY, target, attachment, textarget, tex, level
End Sub
Public Function CheckFramebufferStatusEXT(ByVal target As Long) As Long
    CheckFramebufferStatusEXT = CallFp(0, "glCheckFramebufferStatusEXT", VT_I4, target)
End Function



Public Function GetShaderInfoLog(ByVal shader As Long) As String
    Dim bufSize As Long: bufSize = 4096
    Dim utf8() As Byte: ReDim utf8(bufSize - 1)
    Dim Length As Long
    CallFp 0, "glGetShaderInfoLog", VT_EMPTY, shader, bufSize, VarPtr(Length), VarPtr(utf8(0))
    If Length <= 0 Then Exit Function
    GetShaderInfoLog = StrConv(LeftB$(utf8, Length), vbUnicode)
End Function
Public Function GetProgramInfoLog(ByVal program As Long) As String
    Dim bufSize As Long: bufSize = 4096
    Dim utf8() As Byte: ReDim utf8(bufSize - 1)
    Dim Length As Long
    CallFp 0, "glGetProgramInfoLog", VT_EMPTY, program, bufSize, VarPtr(Length), VarPtr(utf8(0))
    If Length <= 0 Then Exit Function
    GetProgramInfoLog = StrConv(LeftB$(utf8, Length), vbUnicode)
End Function



Public Sub UniformMatrix4fv(ByVal loc As Long, ByVal count As Long, ByVal transpose As Long, ByVal ptrMat As LongPtr)
    CallFp 0, "glUniformMatrix4fv", VT_EMPTY, CLng(loc), CLng(count), CLng(transpose), ptrMat
End Sub
Public Function GetAttribLocation(ByVal program As Long, ByVal Name As String) As Long
    Dim B() As Byte: B = StrConv(Name & vbNullChar, vbFromUnicode)
    GetAttribLocation = CallFp(0, "glGetAttribLocation", VT_I4, CLng(program), VarPtr(B(0)))
End Function
Public Sub GenBuffers(ByVal n As Long, ByVal ptrBuf As LongPtr)
    CallFp 0, "glGenBuffers", VT_EMPTY, CLng(n), ptrBuf
End Sub
Public Sub BufferData(ByVal target As Long, ByVal size As Long, ByVal data As LongPtr, ByVal usage As GLenum)
    CallFp 0, "glBufferData", VT_EMPTY, CLng(target), CLng(size), data, CLng(usage)
End Sub
Public Sub ActiveTexture(ByVal tex As GLenum)
    CallFp 0, "glActiveTexture", VT_EMPTY, CLng(tex)
End Sub
Public Sub DeleteBuffers(ByVal n As Long, ByVal ptrBuf As LongPtr)
    CallFp 0, "glDeleteBuffers", VT_EMPTY, CLng(n), ptrBuf
End Sub
Public Sub DeleteFramebuffersEXT(ByVal n As Long, ByVal ptrBuf As LongPtr)
    CallFp 0, "glDeleteFramebuffersEXT", VT_EMPTY, CLng(n), ptrBuf
End Sub
Public Sub DeleteShader(ByVal shader As Long)
    CallFp 0, "glDeleteShader", VT_EMPTY, CLng(shader)
End Sub
Public Sub GenVertexArrays(ByVal n As Long, ByVal ptrArr As LongPtr)
    CallFp 0, "glGenVertexArrays", VT_EMPTY, CLng(n), ptrArr
End Sub
Public Sub BindVertexArray(ByVal arr As Long)
    CallFp 0, "glBindVertexArray", VT_EMPTY, CLng(arr)
End Sub
Public Sub DeleteVertexArrays(ByVal n As Long, ByVal ptrArr As LongPtr)
    CallFp 0, "glDeleteVertexArrays", VT_EMPTY, CLng(n), ptrArr
End Sub
Public Sub DrawArraysShd(ByVal mode As GLenum, ByVal first As Long, ByVal count As Long)
    CallFp 0, "glDrawArrays", VT_EMPTY, CLng(mode), CLng(first), CLng(count)
End Sub
Public Sub ValidateProgram(ByVal program As Long)
    CallFp 0, "glValidateProgram", VT_EMPTY, program
End Sub
'Init/Term
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
    If glIsList(1000) = 1 Then glDeleteLists 1, 10000
End Sub
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
        .dwFlags = PFD_DRAW_TO_WINDOW Or PFD_SUPPORT_OPENGL Or PFD_DOUBLEBUFFER Or PFD_SUPPORT_GDI  'GDI
        .cColorBits = 32
        .cDepthBits = 24
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

