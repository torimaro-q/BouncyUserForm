VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} glShaderpoi_old 
   Caption         =   "ShaderForm"
   ClientHeight    =   13710
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   19905
   OleObjectBlob   =   "glShaderpoi_old.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "glShaderpoi_old"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit 'tentative program
Private Const LOGPIXELSX As Long = 88
Private Declare PtrSafe Function GetDC Lib "user32" (ByVal hWnd As LongPtr) As LongPtr
Private Declare PtrSafe Function GetDeviceCaps Lib "gdi32" (ByVal hdc As LongPtr, ByVal nIndex As Long) As Long
Private Declare PtrSafe Function ReleaseDC Lib "user32" (ByVal hWnd As LongPtr, ByVal hdc As LongPtr) As Long
Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal ms As Long)
Private Declare PtrSafe Function GetTickCount Lib "kernel32" () As Long

Private Const pcnt As Long = 3000
Private Const TXSZ As Long = 64
Private Const DLENGTH As Long = TXSZ * TXSZ * 4 - 1
Private Const lf As String = vbLf
Private Const WD As Long = 1920, HT As Long = 1080
Private Const BCL As Single = 254 / 255
Private Const dth As Double = 6.28 / pcnt
Private Const EMSG_SH As String = "Shader Error id:"
Private Const EMSG_PG As String = "Program Link Error id:"
Private Const OMSG As String = " Compile OK : vsh-fsh-prg x2 : "

Implements ICFormPhysicsEf
Private myGL As OpenGL, canRender As Boolean, cSrc As Long, fbo As Long, vbo As Long, lPos As Long, lColor As Long
Private tex(1) As Long, vsh(1) As Long, fsh(1) As Long, prg(1) As Long
Private life As Single, IsEffect As Boolean, sw As Long, sh As Long, cx As Double, cy As Double

Private Sub ICFormPhysicsEf_Render(ByRef X As Double, ByRef Y As Double, ByRef dt As Long, ByRef v As Double)
    If canRender = False Then Exit Sub
    life = life - dt
    If life < 0 Then Exit Sub
    Call Update(dt)
    Call DrawParticles
End Sub
Private Sub Update(ByVal dt As Long)
    Dim dst As Long: dst = 1 - cSrc
    With myGL
        .BindFramebufferEXT GL_FRAMEBUFFER, fbo
        .FramebufferTexture2DEXT GL_FRAMEBUFFER, GL_COLOR_ATTACHMENT0, GL_TEXTURE_2D, tex(dst), 0
        .Viewport 0, 0, TXSZ, TXSZ
        .UseProgram prg(1)
        .Uniform1i .GetUniformLocation(prg(1), "posTex"), 0
        .Uniform1f .GetUniformLocation(prg(1), "dt"), CSng(dt)
        .Uniform1f .GetUniformLocation(prg(1), "damp"), 0.999
        .ActiveTexture GL_TEXTURE0
        .BindTexture GL_TEXTURE_2D, tex(cSrc)
        DrawFullScreenQuad
        .BindTexture GL_TEXTURE_2D, 0
        .UseProgram 0
        .BindFramebufferEXT GL_FRAMEBUFFER, 0
        .Viewport 0, 0, sw, sh
    End With
    cSrc = dst
End Sub
Private Sub DrawFullScreenQuad()
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
    With myGL
        .PointSize 5
        .UseProgram prg(0)
        .Uniform1i .GetUniformLocation(prg(0), "posTex"), 0
        .Uniform1f .GetUniformLocation(prg(0), "texSize"), CSng(TXSZ)
        .ActiveTexture GL_TEXTURE0
        .BindTexture GL_TEXTURE_2D, tex(cSrc)
        .BindBuffer GL_ARRAY_BUFFER, vbo
        .EnableVertexAttribArray lPos
        .VertexAttribPointer lPos, 1, GL_FLOAT, GL_FALSE, 4, 0
        .DrawArrays GL_POINTS, 0, pcnt
        .DisableVertexAttribArray lPos
        .BindBuffer GL_ARRAY_BUFFER, 0
        .BindTexture GL_TEXTURE_2D, 0
        .UseProgram 0
    End With
End Sub
Private Sub ICFormPhysicsEf_Reset(ByRef cx As Double, ByRef cy As Double, ByRef ddmg As Double, Optional ByRef hw As Double = 0#, Optional ByRef hh As Double = 0#)
    If canRender = False Then Exit Sub
    If ddmg < 1 Then Exit Sub
    Dim i As Long, i4 As Long, th As Double, rv As Double, d() As Single: ReDim d(0 To DLENGTH)
    th = 0
    life = 400
    For i = 0 To pcnt - 1
        th = Rnd * 6.28
        rv = 3 * ddmg * Rnd
        i4 = i * 4
        d(i4) = CSng(cx)
        d(i4 + 1) = CSng(cy)
        d(i4 + 2) = CSng(rv * FastCos(th))
        d(i4 + 3) = CSng(rv * FastSin(th))
    Next i
    With myGL
        For i = 0 To 1
            .BindTexture GL_TEXTURE_2D, tex(i)
            .TexImage2D GL_TEXTURE_2D, 0, GL_RGBA32F, TXSZ, TXSZ, 0, GL_RGBA, GL_FLOAT, VarPtr(d(0))
        Next i
        .BindTexture GL_TEXTURE_2D, 0
    End With
    cSrc = 0
End Sub
Private Sub ICFormPhysicsEf_Init(ByRef targetGL As OpenGL)
    Set myGL = targetGL
    With myGL
        Dim i As Long, mg As String, vs(1) As String, fs(1) As String
        vs(0) = VertexSrc
        vs(1) = SimVertexSrc
        fs(0) = FragmentSrc
        fs(1) = SimFragmentSrc
        For i = 0 To 1
            vsh(i) = .CreateShader(GL_VERTEX_SHADER)
            .ShaderSource vsh(i), vs(i)
            .CompileShader vsh(i)
            
            fsh(i) = .CreateShader(GL_FRAGMENT_SHADER)
            .ShaderSource fsh(i), fs(i)
            .CompileShader fsh(i)
            
            prg(i) = .CreateProgram
            .AttachShader prg(i), vsh(i)
            .AttachShader prg(i), fsh(i)
            .LinkProgram prg(i)
        Next i
        lPos = .GetAttribLocation(prg(0), "particleID")
        lColor = .GetAttribLocation(prg(0), "aColor")
        For i = 0 To 1
            mg = mg & GetLog(vsh(i))
            mg = mg & GetLog(fsh(i), prg(i))
        Next i
        .GenTextures 2, VarPtr(tex(0))
        For i = 0 To 1
            .BindTexture GL_TEXTURE_2D, tex(i)
            .TexParameteri GL_TEXTURE_2D, GL_TEXTURE_MIN_FILTER, GL_NEAREST
            .TexParameteri GL_TEXTURE_2D, GL_TEXTURE_MAG_FILTER, GL_NEAREST
            .TexParameteri GL_TEXTURE_2D, GL_TEXTURE_WRAP_S, GL_CLAMP_TO_EDGE
            .TexParameteri GL_TEXTURE_2D, GL_TEXTURE_WRAP_T, GL_CLAMP_TO_EDGE
            .TexImage2D GL_TEXTURE_2D, 0, GL_RGBA32F, TXSZ, TXSZ, 0, GL_RGBA, GL_FLOAT, ByVal 0&
        Next i
        .BindTexture GL_TEXTURE_2D, 0
        .GenFramebuffersEXT 1, VarPtr(fbo)
        Dim ids() As Single: ReDim ids(0 To pcnt - 1)
        For i = 0 To pcnt - 1
            ids(i) = CSng(i)
        Next i
        .GenBuffers 1, VarPtr(vbo)
        .BindBuffer GL_ARRAY_BUFFER, vbo
        .BufferData GL_ARRAY_BUFFER, 4 * pcnt, VarPtr(ids(0)), GL_STATIC_DRAW
        .BindBuffer GL_ARRAY_BUFFER, 0
        IsEffect = False
        sw = Frame1.width * Tw2Px
        sh = Frame1.height * Tw2Px
        If Not .Param Is Nothing Then
            With .Param
                If .Item("width") > 0 Then
                    IsEffect = True
                    sw = .Item("width")
                    sh = .Item("height")
                End If
            End With
        End If
        If .CheckFramebufferStatusEXT(GL_FRAMEBUFFER) <> 36053 Then
            canRender = False
        Else
            If mg <> "" Then
                canRender = False
            Else
                canRender = True
                For i = 0 To 9 'warm up
                    Call ICFormPhysicsEf_Render(9, 9, i, 9)
                Next i
                mg = OMSG & vsh(0) & "-" & fsh(0) & "-" & prg(0) & "," & vsh(1) & "-" & fsh(1) & "-" & prg(1)
            End If
        End If
    End With
    With TextBox0
        .Value = .Value & lf & mg
    End With
End Sub
Private Property Get ICFormPhysicsEf_CreateInstance() As ICFormPhysicsEf
    Set ICFormPhysicsEf_CreateInstance = New glShaderpoi_old
End Property
Private Sub UserForm_Terminate()
    If Not myGL Is Nothing Then
        myGL.PaintEnd
        Call CleanupGL
        Set myGL = Nothing
    End If
End Sub
Private Sub UserForm_Initialize()
    CommandButton1.Picture = Application.CommandBars.GetImageMso("MacroPlay", 16, 16)
    cx = WD * 0.5
    cy = HT * 0.5
End Sub
Private Sub CleanupGL()
    Dim i As Long
    With myGL
        On Error GoTo err
            .BindFramebufferEXT GL_FRAMEBUFFER, fbo
            .FramebufferTexture2DEXT GL_FRAMEBUFFER, GL_COLOR_ATTACHMENT0, GL_TEXTURE_2D, 0, 0
            .BindFramebufferEXT GL_FRAMEBUFFER, 0
            .BindTexture GL_TEXTURE_2D, 0
err:
        If vbo <> 0 Then .DeleteBuffers 1, VarPtr(vbo)
        If tex(0) <> 0 Then .DeleteTextures 2, VarPtr(tex(0))
        If fbo <> 0 Then .DeleteFramebuffersEXT 1, VarPtr(fbo)
        For i = 0 To 1
            If prg(i) <> 0 Then .DeleteProgram prg(i)
            If vsh(i) <> 0 Then .DeleteShader vsh(i)
            If fsh(i) <> 0 Then .DeleteShader fsh(i)
        Next i
    End With
End Sub
Public Property Get VertexSrc() As String
    VertexSrc = CStr(TextBox1.Value)
End Property
Public Property Get FragmentSrc() As String
    FragmentSrc = CStr(TextBox2.Value)
End Property
Public Property Get SimVertexSrc() As String
    SimVertexSrc = CStr(TextBox3.Value)
End Property
Public Property Get SimFragmentSrc() As String
    SimFragmentSrc = CStr(TextBox4.Value)
End Property
'Chack Mode
Private Sub InitTestGL()
    Dim hWnd As LongPtr: hWnd = Frame1.[_GethWnd]
    Set myGL = New OpenGL
    With myGL
        .hWnd = hWnd
        .PaintStart
        .ClearColor BCL, BCL, BCL, 1
        .Enable GL_DEPTH_TEST
        .SwapIntervalEXT 0
        .Disable GL_LIGHTING
    End With
End Sub
Private Sub CommandButton1_Click()
    TextBox0.Value = ""
    If myGL Is Nothing Then Call InitTestGL
    Call CleanupGL
    Call ICFormPhysicsEf_Init(myGL)
    Call TestRender
End Sub
Private Sub LabelP_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    With LabelP
        If Button = 1 Then
            .left = .left + X - .width * 0.5
            .top = .top + Y - .height * 0.5
            cx = WD * ((.left + .width * 0.5) / Frame3.width)
            cy = HT * ((.top + .height * 0.5) / Frame3.height)
            Label4.Caption = "X = " & format(cx, "0") & "/ Y = " & format(cy, "0")
        End If
    End With
End Sub
Private Function GetLog(Optional ByVal shd As Long = 0, Optional ByVal prg As Long = 0) As String
    Dim msg As String
    With myGL
        If shd > 0 Then If CLng(.GetShaderiv(shd, GL_COMPILE_STATUS)) = 0 Then msg = msg & EMSG_SH & shd & lf & .GetShaderInfoLog(shd)
        If prg > 0 Then If CLng(.GetProgramiv(prg, GL_LINK_STATUS)) = 0 Then msg = msg & EMSG_PG & prg & lf & .GetProgramInfoLog(prg)
    End With
    GetLog = msg
End Function
Private Sub TestRender()
    Dim i As Long, fw As Long, fh As Long, ct As Long, pt As Long, t0 As Long, fc As Long
    fw = Frame1.width * Tw2Px
    fh = Frame1.height * Tw2Px
    fc = 1500
    With myGL
        Call ICFormPhysicsEf_Reset(cx, cy, 3, 0, 0)
        .Viewport 0, 0, fw, fh
        t0 = GetTickCount
        ct = t0
        For i = 0 To fc
            pt = ct
            ct = GetTickCount
            .Clear GL_COLOR_BUFFER_BIT Or GL_DEPTH_BUFFER_BIT
            .MatrixMode GL_PROJECTION
                .LoadIdentity
                .Ortho2D 0, WD, HT, 0
            .MatrixMode GL_MODELVIEW
                .LoadIdentity
                Call ICFormPhysicsEf_Render(cx, cy, ct - pt, 5)
            .SwapBuffers
        Next i
        Label3.Caption = format((1000# * fc) / (GetTickCount - t0), "0.00") & " fps"
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
