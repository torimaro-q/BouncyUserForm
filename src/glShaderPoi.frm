VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} glShaderPoi 
   Caption         =   "UserForm"
   ClientHeight    =   12360
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   19110
   OleObjectBlob   =   "glShaderPoi.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "glShaderPoi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const LOGPIXELSX As Long = 88
Private Declare PtrSafe Function GetDC Lib "user32" (ByVal hWnd As LongPtr) As LongPtr
Private Declare PtrSafe Function GetDeviceCaps Lib "gdi32" (ByVal hdc As LongPtr, ByVal nIndex As Long) As Long
Private Declare PtrSafe Function ReleaseDC Lib "user32" (ByVal hWnd As LongPtr, ByVal hdc As LongPtr) As Long
Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal ms As Long)
Private Declare PtrSafe Function GetTickCount Lib "kernel32" () As Long

Implements ICFormPhysicsEf 'tentative program
Public canRender As Boolean
Private Const PCOUNT As Long = 3000
Private Const TXSIZE As Long = 64
Private Const DATA_LENGTH As Long = TXSIZE * TXSIZE * 4 - 1
Private Const lf As String = vbLf
Private Const WD As Long = 1920, HT As Long = 1080
Private myGL As OpenGL
Private posTex(1) As Long, cSrc As Long
Private fbo As Long
Private prog As Long, sProg As Long
Private vboID As Long, locPos As Long, locColor
Private life As Single
Private IsEffectMode As Boolean, sw As Long, sh As Long
Private vsh As Long, fsh As Long, vshSim As Long, fshSim As Long

Private Sub ICFormPhysicsEf_Render(ByRef X As Double, ByRef Y As Double, ByRef dt As Long, ByRef v As Double)
    If canRender = False Then Exit Sub
    life = life - dt
    If life < 0 Then Exit Sub
    Call SimulateOnGPU(dt)
    Call DrawParticles
End Sub
Private Sub SimulateOnGPU(ByVal dt As Long)
    If canRender = False Then Exit Sub
    Dim dst As Long
    dst = 1 - cSrc
    With myGL
        .BindFramebufferEXT GL_FRAMEBUFFER, fbo
        .FramebufferTexture2DEXT GL_FRAMEBUFFER, GL_COLOR_ATTACHMENT0, GL_TEXTURE_2D, posTex(dst), 0
        .Viewport 0, 0, TXSIZE, TXSIZE
        .UseProgram sProg
        .Uniform1i .GetUniformLocation(sProg, "posTex"), 0
        .Uniform1f .GetUniformLocation(sProg, "dt"), CSng(dt)
        .Uniform1f .GetUniformLocation(sProg, "damp"), 0.999
        .ActiveTexture GL_TEXTURE0
        .BindTexture GL_TEXTURE_2D, posTex(cSrc)
        DrawFullScreenQuad
        .BindTexture GL_TEXTURE_2D, 0
        .UseProgram 0
        .BindFramebufferEXT GL_FRAMEBUFFER, 0
        .Viewport 0, 0, sw, sh
    End With
    cSrc = dst
End Sub
Private Sub DrawFullScreenQuad()
    If canRender = False Then Exit Sub
    With myGL
        .Disable GL_DEPTH_TEST
        .Begin GL_TRIANGLE_STRIP
            .TexCoord2f 0, 0
            .Vertex2f -1, -1
            .TexCoord2f 1, 0
            .Vertex2f 1, -1
            .TexCoord2f 0, 1
            .Vertex2f -1, 1
            .TexCoord2f 1, 1
            .Vertex2f 1, 1
        .End1
        .Enable GL_DEPTH_TEST
    End With
End Sub
Private Sub DrawParticles()
    If canRender = False Then Exit Sub
    With myGL
        .PointSize 5
        .UseProgram prog
        .Uniform1i .GetUniformLocation(prog, "posTex"), 0
        .Uniform1f .GetUniformLocation(prog, "texSize"), CSng(TXSIZE)
        .ActiveTexture GL_TEXTURE0
        .BindTexture GL_TEXTURE_2D, posTex(cSrc)
        .BindBuffer GL_ARRAY_BUFFER, vboID
        .EnableVertexAttribArray locPos
        .VertexAttribPointer locPos, 1, GL_FLOAT, GL_FALSE, 4, 0
        .DrawArrays GL_POINTS, 0, PCOUNT
        .DisableVertexAttribArray locPos
        .BindBuffer GL_ARRAY_BUFFER, 0
        .BindTexture GL_TEXTURE_2D, 0
        .UseProgram 0
    End With
End Sub
Private Sub ICFormPhysicsEf_Reset(ByRef cx As Double, ByRef cy As Double, ByRef ddmg As Double, Optional ByRef hw As Double = 0#, Optional ByRef hh As Double = 0#)
    If canRender = False Then Exit Sub
    If ddmg < 1 Then Exit Sub
    Dim data() As Single: ReDim data(0 To DATA_LENGTH)
    Dim i As Long, i4 As Long, t As Long, th As Double, rv As Double, dth As Double
    th = 0
    life = 400
    dth = 6.28 / PCOUNT
    For i = 0 To PCOUNT - 1
        th = th + dth
        rv = ddmg + ddmg * Rnd * 0.3
        i4 = i * 4
        data(i4 + 0) = CSng(cx)
        data(i4 + 1) = CSng(cy)
        data(i4 + 2) = CSng(rv * FastCos(th))
        data(i4 + 3) = CSng(rv * FastSin(th))
    Next i
    With myGL
        For t = 0 To 1
            .BindTexture GL_TEXTURE_2D, posTex(t)
            .TexImage2D GL_TEXTURE_2D, 0, GL_RGBA32F, TXSIZE, TXSIZE, 0, GL_RGBA, GL_FLOAT, VarPtr(data(0))
        Next t
        .BindTexture GL_TEXTURE_2D, 0
    End With
    cSrc = 0
End Sub
Private Sub ICFormPhysicsEf_Init(ByRef targetGL As OpenGL)
    Set myGL = targetGL
    With myGL
        Dim i As Long, fbStatus As Long, mg As String

        vsh = .CreateShader(GL_VERTEX_SHADER)
        .ShaderSource vsh, VertexSource
        .CompileShader vsh
        
        fsh = .CreateShader(GL_FRAGMENT_SHADER)
        .ShaderSource fsh, FragmentSource
        .CompileShader fsh
        
        prog = .CreateProgram
        .AttachShader prog, vsh
        .AttachShader prog, fsh
        .LinkProgram prog

        locPos = .GetAttribLocation(prog, "particleID")
        locColor = .GetAttribLocation(prog, "aColor")
        
        vshSim = .CreateShader(GL_VERTEX_SHADER)
        .ShaderSource vshSim, SimVertexSource
        .CompileShader vshSim
        
        fshSim = .CreateShader(GL_FRAGMENT_SHADER)
        .ShaderSource fshSim, SimFragmentSource
        .CompileShader fshSim
        
        sProg = .CreateProgram
        .AttachShader sProg, vshSim
        .AttachShader sProg, fshSim
        .LinkProgram sProg
        
        mg = mg & GetLog(vsh)
        mg = mg & GetLog(fsh, prog)
        mg = mg & GetLog(vshSim)
        mg = mg & GetLog(fshSim, sProg)
        .GenTextures 2, VarPtr(posTex(0))
        
        For i = 0 To 1
            .BindTexture GL_TEXTURE_2D, posTex(i)
            .TexParameteri GL_TEXTURE_2D, GL_TEXTURE_MIN_FILTER, GL_NEAREST
            .TexParameteri GL_TEXTURE_2D, GL_TEXTURE_MAG_FILTER, GL_NEAREST
            .TexParameteri GL_TEXTURE_2D, GL_TEXTURE_WRAP_S, GL_CLAMP_TO_EDGE
            .TexParameteri GL_TEXTURE_2D, GL_TEXTURE_WRAP_T, GL_CLAMP_TO_EDGE
            .TexImage2D GL_TEXTURE_2D, 0, GL_RGBA32F, TXSIZE, TXSIZE, 0, GL_RGBA, GL_FLOAT, ByVal 0&
        Next i
        
        .BindTexture GL_TEXTURE_2D, 0
        .GenFramebuffersEXT 1, VarPtr(fbo)
        Dim ids() As Single: ReDim ids(0 To PCOUNT - 1)
        For i = 0 To PCOUNT - 1
            ids(i) = CSng(i)
        Next i
        
        .GenBuffers 1, VarPtr(vboID)
        .BindBuffer GL_ARRAY_BUFFER, vboID
        .BufferData GL_ARRAY_BUFFER, 4 * PCOUNT, VarPtr(ids(0)), GL_STATIC_DRAW
        .BindBuffer GL_ARRAY_BUFFER, 0
        
        IsEffectMode = False
        sw = Frame1.width * Tw2Px
        sh = Frame1.height * Tw2Px
        If Not .Param Is Nothing Then
            With .Param
                If .Item("width") > 0 Then
                    IsEffectMode = True
                    sw = .Item("width")
                    sh = .Item("height")
                End If
            End With
        End If
        
        fbStatus = myGL.CheckFramebufferStatusEXT(GL_FRAMEBUFFER)
        If fbStatus <> 36053 Then
            canRender = False
        Else
            If mg <> "" Then
                canRender = False
            Else
                canRender = True
                'warmup
                Call ICFormPhysicsEf_Render(100, 100, 10, 10)
                Call ICFormPhysicsEf_Render(100, 100, 10, 10)
                Call ICFormPhysicsEf_Render(100, 100, 10, 10)
            End If
        End If
        Me.TextBox0.Value = Me.TextBox0.Value & lf & mg
    End With
End Sub
Private Property Get ICFormPhysicsEf_CreateInstance() As ICFormPhysicsEf
    Set ICFormPhysicsEf_CreateInstance = New glShaderPoi
End Property
Private Sub UserForm_Terminate()
    If Not myGL Is Nothing Then
        Call CleanupGL
        myGL.PaintEnd
        Set myGL = Nothing
    End If
End Sub
Private Function GetLog(Optional ByVal shader As Long = 0, Optional ByVal program As Long = 0) As String
    Dim msg As String
    With myGL
        If shader > 0 Then If CLng(.GetShaderiv(shader, GL_COMPILE_STATUS)) = 0 Then msg = msg & "Vertex Shader Error:" & lf & .GetShaderInfoLog(shader) & lf
        If program > 0 Then If CLng(.GetProgramiv(program, GL_LINK_STATUS)) = 0 Then msg = msg & "Program Link Error:" & lf & .GetProgramInfoLog(program) & lf
    End With
    GetLog = msg
End Function
Private Sub UserForm_Initialize()
    Me.CommandButton1.Picture = Application.CommandBars.GetImageMso("MacroPlay", 32, 32)
End Sub
Private Sub CleanupGL()
    With myGL
        On Error GoTo err
            .BindFramebufferEXT GL_FRAMEBUFFER, fbo
            .FramebufferTexture2DEXT GL_FRAMEBUFFER, GL_COLOR_ATTACHMENT0, GL_TEXTURE_2D, 0, 0
            .BindFramebufferEXT GL_FRAMEBUFFER, 0
            .BindTexture GL_TEXTURE_2D, 0
err:
        If vboID <> 0 Then .DeleteBuffers 1, VarPtr(vboID)
        If posTex(0) <> 0 Then .DeleteTextures 2, VarPtr(posTex(0))
        If fbo <> 0 Then .DeleteFramebuffersEXT 1, VarPtr(fbo)
        If prog <> 0 Then .DeleteProgram prog
        If sProg <> 0 Then .DeleteProgram sProg
        If vsh <> 0 Then .DeleteShader vsh
        If fsh <> 0 Then .DeleteShader fsh
        If vshSim <> 0 Then .DeleteShader fshSim
        If fshSim <> 0 Then .DeleteShader fshSim
    End With
End Sub
Public Property Get VertexSource() As String
    VertexSource = CStr(TextBox1.Value)
End Property
Public Property Get FragmentSource() As String
    FragmentSource = CStr(TextBox2.Value)
End Property
Public Property Get SimVertexSource() As String
    SimVertexSource = CStr(TextBox3.Value)
End Property
Public Property Get SimFragmentSource() As String
    SimFragmentSource = CStr(TextBox4.Value)
End Property
'Chack Mode
Private Sub CommandButton1_Click()
    If myGL Is Nothing Then
        Dim hWnd As LongPtr
        hWnd = Me.Frame1.[_GethWnd]
        Set myGL = New OpenGL
        myGL.hWnd = hWnd
        myGL.PaintStart
    End If
    Call CleanupGL
    Call ICFormPhysicsEf_Init(myGL)
    Call TestRender
End Sub
Private Sub TestRender()
    If canRender = False Then Exit Sub
    Dim i As Long, fw As Long, fh As Long, tw As Double, th As Double, ct As Long, pt As Long
    fw = Frame1.width * Tw2Px
    fh = Frame1.height * Tw2Px
    tw = WD 'FHD
    th = HT
    With myGL
        Call ICFormPhysicsEf_Reset(WD * 0.5, HT * 0.5, 3, 0, 0)
        .ClearColor 254 / 255, 254 / 255, 254 / 255, 1
        .Enable GL_DEPTH_TEST
        .Viewport 0, 0, fw, fh
        .Disable GL_LIGHTING
        .SwapIntervalEXT 0
        ct = GetTickCount
        For i = 0 To 50
            pt = ct
            ct = GetTickCount
            Debug.Print ct - pt
            .Clear GL_COLOR_BUFFER_BIT Or GL_DEPTH_BUFFER_BIT
            .MatrixMode GL_PROJECTION
                .LoadIdentity
                .Ortho2D 0, tw, th, 0
            .MatrixMode GL_MODELVIEW
                Call ICFormPhysicsEf_Render(tw * 0.5 - 5 * i, th * 0.5 - 5 * i, 10, 1.414 * 5 * i)
                .LoadIdentity
            .SwapBuffers
        Next i
    End With
End Sub
Private Function GetDPI() As Long
    Dim hdc As LongPtr: hdc = GetDC(0)
    GetDPI = GetDeviceCaps(hdc, LOGPIXELSX)
    Call ReleaseDC(0, hdc)
End Function
Private Function Tw2Px() As Double
    Tw2Px = GetDPI() / 72
End Function
