VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} glShaderTest 
   Caption         =   "UserForm"
   ClientHeight    =   12360
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   19110
   OleObjectBlob   =   "glShaderTest.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "glShaderTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Private Const LOGPIXELSX As Long = 88
Private Declare PtrSafe Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Private Declare PtrSafe Function GetDC Lib "user32" (ByVal hWnd As LongPtr) As LongPtr
Private Declare PtrSafe Function GetDeviceCaps Lib "gdi32" (ByVal hdc As LongPtr, ByVal nIndex As Long) As Long
Private Declare PtrSafe Function ReleaseDC Lib "user32" (ByVal hWnd As LongPtr, ByVal hdc As LongPtr) As Long
Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal ms As Long)
Implements ICFormPhysicsEf
Private Type ParticleVertex
    X As Single '4
    Y As Single '8
    VX As Single '12
    VY As Single '16
    lf As Single '20
    R As Byte '21
    G As Byte '22
    B As Byte '23
    A As Byte '24
End Type
Private Const stride As Long = 24
Private Const PARTICLECOUNT As Long = 3000
Private Const VBO_SIZE As Long = stride * (PARTICLECOUNT + 1)
Private Const dth As Double = 6.283185307 / PARTICLECOUNT
Private pVB(PARTICLECOUNT) As ParticleVertex
Private InvLF(PARTICLECOUNT) As Single
Private prog As Long, locPos As Long, locColor As Long, vbo(0) As Long, life As Single
Private myGL As OpenGL

Private Sub ICFormPhysicsEf_Render(ByRef X As Double, ByRef Y As Double, ByRef dt As Long, ByRef v As Double)
    life = life - dt
    If life < 0 Then Exit Sub
    Call Update(dt)
    With myGL
        .PointSize 3
        .UseProgram prog
            .BindBuffer GL_ARRAY_BUFFER, vbo(0)
            .BufferSubData GL_ARRAY_BUFFER, 0, VBO_SIZE, VarPtr(pVB(0))
                .EnableVertexAttribArray locPos
                .VertexAttribPointer locPos, 2, GL_FLOAT, GL_FALSE, stride, 0
                
                .EnableVertexAttribArray locColor
                .VertexAttribPointer locColor, 4, GL_UNSIGNED_BYTE, GL_TRUE, stride, 20
                
                .DrawArrays GL_POINTS, 0, PARTICLECOUNT + 1
                
                .DisableVertexAttribArray locColor
                .DisableVertexAttribArray locPos
            .BindBuffer GL_ARRAY_BUFFER, 0
        .UseProgram 0
    End With
End Sub
Private Sub Update(ByVal dt As Long)
    Dim i As Long
    Dim alpha As Single
    Dim damp As Single
    damp = 1 - dt * 0.0002
    For i = 0 To PARTICLECOUNT
        With pVB(i)
            .VX = .VX * damp
            .VY = .VY * damp
            .X = .X + .VX * dt
            .Y = .Y + .VY * dt
            .lf = .lf - dt
            alpha = .lf * InvLF(i)
            If alpha < 0 Then
                .A = 0
            Else
                .A = alpha * 255
            End If
        End With
    Next i
End Sub
Private Sub ICFormPhysicsEf_Reset(ByRef cx As Double, ByRef cy As Double, ByRef ddmg As Double, Optional ByRef hw As Double = 0#, Optional ByRef hh As Double = 0#)
    If ddmg < 1 Then Exit Sub
    Dim i As Long
    Dim rv As Double
    Dim th As Double
    life = 400
    th = 0
    For i = 0 To PARTICLECOUNT
        With pVB(i)
            th = th + dth
            rv = ddmg + ddmg * Rnd * 0.3
            .lf = life
            InvLF(i) = 1 / life
            .X = cx
            .Y = cy
            .VX = rv * FastCos(th)
            .VY = rv * FastSin(th)
            .R = 230 + Int(Rnd * 25)
            .G = 204 + Int(Rnd * 25)
            .B = 179 + Int(Rnd * 25)
            .A = 255
        End With
    Next
End Sub
Private Sub ICFormPhysicsEf_Init(ByRef targetGL As OpenGL)
    Set myGL = targetGL
    Dim vsh As Long
    Dim fsh As Long
    With myGL
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
        locPos = .GetAttribLocation(prog, "aPos")
        locColor = .GetAttribLocation(prog, "aColor")
        .GenBuffers 1, VarPtr(vbo(0))
        .BindBuffer GL_ARRAY_BUFFER, vbo(0)
        .BufferData GL_ARRAY_BUFFER, VBO_SIZE, ByVal 0&, GL_DYNAMIC_DRAW
        .BindBuffer GL_ARRAY_BUFFER, 0
        Debug.Print "vsh:" & vsh & "/fsh:" & fsh & "/prog:" & prog
    End With
End Sub
Private Property Get ICFormPhysicsEf_CreateInstance() As ICFormPhysicsEf
    Set ICFormPhysicsEf_CreateInstance = New glShaderTest
End Property
'test mode
Private Sub CommandButton1_Click()
    CompileCheck
End Sub
Public Property Get VertexSource() As String
    VertexSource = CStr(TextBox1.Value)
End Property
Public Property Get FragmentSource() As String
    FragmentSource = CStr(TextBox2.Value)
End Property
Private Sub CompileCheck()
    If myGL Is Nothing Then
        Dim vsh As Long, fsh As Long, prog As Long, hWnd As LongPtr
        hWnd = Me.Frame1.[_GethWnd]
        Set myGL = New OpenGL
        myGL.hWnd = hWnd
        myGL.PaintStart
        
        With myGL
            vsh = .CreateShader(GL_VERTEX_SHADER)
            Call .ShaderSource(vsh, VertexSource)
            Call .CompileShader(vsh)
            
            fsh = .CreateShader(GL_FRAGMENT_SHADER)
            Call .ShaderSource(fsh, FragmentSource)
            Call .CompileShader(fsh)
            
            prog = .CreateProgram()
            Call .AttachShader(prog, vsh)
            Call .AttachShader(prog, fsh)
            Call .LinkProgram(prog)
            Call .UseProgram(prog)
            
            Dim err As String, okV As Long, okF As Long, okP As Long
            
            okV = .GetShaderiv(vsh, GL_COMPILE_STATUS)
            okF = .GetShaderiv(fsh, GL_COMPILE_STATUS)
            okP = .GetProgramiv(prog, GL_LINK_STATUS)
            
            If okV = 0 Then err = err & "Vertex Shader Error:" & vbCrLf & .GetShaderInfoLog(vsh) & vbCrLf
            If okF = 0 Then err = err & "Fragment Shader Error:" & vbCrLf & .GetShaderInfoLog(fsh) & vbCrLf
            If okP = 0 Then err = err & "Program Link Error:" & vbCrLf & .GetProgramInfoLog(prog) & vbCrLf
            
            If err = "" Then
                Me.TextBox3.Value = "suceess!!"
                If vbo(0) <> 0 Then .DeleteBuffers 1, VarPtr(vbo(0))
                If prog <> 0 Then .DeleteProgram prog
                Call ICFormPhysicsEf_Init(myGL)
                Call TestRender
            Else
                Me.TextBox3.Value = err & vbNewLine & "vsh:" & vsh & vbNewLine & "fsh:" & fsh
            End If
        End With
    End If
    If Not myGL Is Nothing Then
        If vbo(0) <> 0 Then myGL.DeleteBuffers 1, VarPtr(vbo(0))
        If prog <> 0 Then myGL.DeleteProgram prog
        myGL.PaintEnd
        Set myGL = Nothing
    End If
End Sub
Private Sub TestRender()
    Dim i As Long, Tw2Px As Double, fw As Long, fh As Long, tw As Double, th As Double
    Tw2Px = GetDPI() / 72
    fw = Frame1.width * Tw2Px
    fh = Frame1.height * Tw2Px
    tw = 1920 'FHD
    th = 1080
    With myGL
        Call ICFormPhysicsEf_Reset(1920 * 0.5, 1080 * 0.5, 3, 0, 0)
        For i = 0 To 100
            .ClearColor 1, 1, 1, 1
            .Enable GL_DEPTH_TEST
            .Viewport 0, 0, fw, fh
            .Clear GL_COLOR_BUFFER_BIT Or GL_DEPTH_BUFFER_BIT
            .MatrixMode GL_PROJECTION
            .LoadIdentity
            .Ortho2D 0, tw, th, 0
            .MatrixMode GL_MODELVIEW
            Call ICFormPhysicsEf_Render(tw * 0.5 - 5 * i, th * 0.5 - 5 * i, 10, 1.414 * 5 * i)
            .LoadIdentity
            .SwapBuffers
            Sleep 10
        Next i
    End With
End Sub
Private Function GetDPI() As Long
    Dim hdc As LongPtr: hdc = GetDC(0)
    GetDPI = GetDeviceCaps(hdc, LOGPIXELSX)
    Call ReleaseDC(0, hdc)
End Function
Private Sub UserForm_Initialize()
    Me.CommandButton1.Picture = Application.CommandBars.GetImageMso("MacroPlay", 32, 32)
End Sub
Private Sub UserForm_Terminate()
    If Not myGL Is Nothing Then
        If vbo(0) <> 0 Then myGL.DeleteBuffers 1, VarPtr(vbo(0))
        If prog <> 0 Then myGL.DeleteProgram prog
        myGL.PaintEnd
        Set myGL = Nothing
    End If
End Sub

