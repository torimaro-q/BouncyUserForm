VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} glShaderTest 
   Caption         =   "ShaderTest"
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
Implements ICFormPhysicsEf
Private prog As Long
Private locPos As Long
Private locColor As Long
Private Const PARTICLECOUNT As Long = 1000
Private Const PARTICLECOUNT_AB_POS As Long = (PARTICLECOUNT + 1) * 8
Private Const PARTICLECOUNT_AB_COL As Long = (PARTICLECOUNT + 1) * 16
Private pD(PARTICLECOUNT) As Vector2f
Private pV(PARTICLECOUNT) As Vector2f
Private pC(PARTICLECOUNT) As Color4
Private lf(PARTICLECOUNT) As Double
Private InvLF(PARTICLECOUNT) As Double
Private effectCX As Double, effectCY As Double
Private vboPos(0) As Long, vboCol(0) As Long
Private myGL As OpenGL, life As Double, dd As Double
Private Sub ICFormPhysicsEf_Render(ByRef X As Double, ByRef Y As Double, ByRef dt As Long, ByRef v As Double)
    life = life - dt
    If life < 0 Then Exit Sub
    Call Update(dt, v)
    With myGL
        .LineWidth 5 + dd * 3
        .Enable GL_BLEND
        .BlendFunc GL_SRC_ALPHA, GL_ONE_MINUS_SRC_ALPHA
            .UseProgram prog
            .BindBuffer GL_ARRAY_BUFFER, vboPos(0)
            .BufferSubData GL_ARRAY_BUFFER, 0, PARTICLECOUNT_AB_POS, VarPtr(pD(0))
            .EnableVertexAttribArray locPos
            .VertexAttribPointer locPos, 2, GL_FLOAT, GL_FALSE, 0, 0
            
                .BindBuffer GL_ARRAY_BUFFER, vboCol(0)
                .BufferSubData GL_ARRAY_BUFFER, 0, PARTICLECOUNT_AB_COL, VarPtr(pC(0))
                
                .EnableVertexAttribArray locColor
                .VertexAttribPointer locColor, 4, GL_FLOAT, GL_FALSE, 0, 0
                
                .DrawArraysShd GL_LINE_STRIP, 0, PARTICLECOUNT + 1
                .DisableVertexAttribArray locColor
                
            .DisableVertexAttribArray locPos
        .Disable GL_BLEND
    End With
End Sub
Private Sub Update(ByRef dt, ByRef v)
    Dim i As Long, alpha As Single, t As Single
    For i = 0 To PARTICLECOUNT
        With pV(i)
            .X = .X * (1 - dt * 0.0002)
            .Y = .Y * (1 - dt * 0.0002)
        End With
        With pD(i)
            .X = .X + pV(i).X * dt
            .Y = .Y + pV(i).Y * dt
        End With
        lf(i) = lf(i) - dt
        alpha = lf(i) * InvLF(i)
        If alpha >= 0 Then pC(i).A = alpha
    Next i
End Sub
Private Sub ICFormPhysicsEf_Reset(ByRef cx As Double, ByRef cy As Double, ByRef ddmg As Double, Optional ByRef hw As Double = 0#, Optional ByRef hh As Double = 0#)
    If ddmg < 1 Then Exit Sub
    Dim i As Long, rv As Double, th As Double, dth As Double
    effectCX = cx
    effectCY = cy
    life = 400
    th = 0
    dth = 6.28 / PARTICLECOUNT
    For i = 0 To PARTICLECOUNT
        th = th + dth
        rv = ddmg + ddmg * Rnd * 0.3
        lf(i) = life
        InvLF(i) = 1 / life
        With pD(i)
            .X = cx
            .Y = cy
        End With
        With pV(i)
            .X = rv * FastCos(th)
            .Y = rv * FastSin(th)
        End With
        With pC(i)
            .R = 0.9 + Rnd * 0.1
            .G = 0.8 + Rnd * 0.1
            .B = 0.7 + Rnd * 0.1
            .A = 1
        End With
    Next i
End Sub
Private Sub ICFormPhysicsEf_Init(ByRef targetGL As OpenGL)
    Set myGL = targetGL
    Dim vsh As Long, fsh As Long
    With myGL
        vsh = .CreateShader(GL_VERTEX_SHADER)
        .ShaderSource vsh, VertexSource
        .CompileShader vsh
        fsh = .CreateShader(GL_FRAGMENT_SHADER)
        .ShaderSource fsh, FragmentSource
        .CompileShader fsh
        prog = .CreateProgram()
        .AttachShader prog, vsh
        .AttachShader prog, fsh
        .LinkProgram prog
        locPos = .GetAttribLocation(prog, "aPos")
        locColor = .GetAttribLocation(prog, "aColor")
        .GenBuffers 1, VarPtr(vboPos(0))
        .BindBuffer GL_ARRAY_BUFFER, vboPos(0)
        .BufferData GL_ARRAY_BUFFER, PARTICLECOUNT_AB_POS, ByVal 0&, GL_STREAM_DRAW
        .GenBuffers 1, VarPtr(vboCol(0))
        .BindBuffer GL_ARRAY_BUFFER, vboCol(0)
        .BufferData GL_ARRAY_BUFFER, PARTICLECOUNT_AB_COL, ByVal 0&, GL_STREAM_DRAW
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
        If myGL Is Nothing Then
            hWnd = Me.Frame1.[_GethWnd]
            Set myGL = New OpenGL
            myGL.hWnd = hWnd
            myGL.PaintStart
        End If
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
    
            Dim err As String
            Dim okV As Long, okF As Long, okP As Long
            
            okV = .GetShaderiv(vsh, GL_COMPILE_STATUS)
            okF = .GetShaderiv(fsh, GL_COMPILE_STATUS)
            okP = .GetProgramiv(prog, GL_LINK_STATUS)
            
            If okV = 0 Then err = err & "Vertex Shader Error:" & vbCrLf & .GetShaderInfoLog(vsh) & vbCrLf
            If okF = 0 Then err = err & "Fragment Shader Error:" & vbCrLf & .GetShaderInfoLog(fsh) & vbCrLf
            If okP = 0 Then err = err & "Program Link Error:" & vbCrLf & .GetProgramInfoLog(prog) & vbCrLf
            
            If err = "" Then
                Me.TextBox3.Value = "suceess!!"
                .ClearColor 0, 1, 0, 1
                .Enable GL_DEPTH_TEST
                .Viewport 0, 0, Frame1.width, Frame1.height
                .Clear GL_COLOR_BUFFER_BIT Or GL_DEPTH_BUFFER_BIT
                .MatrixMode GL_PROJECTION
                .LoadIdentity
                .Ortho2D 0, Frame1.width, Frame1.height, 0
                .MatrixMode GL_MODELVIEW
                .LoadIdentity
                .SwapBuffers
            Else
                Me.TextBox3.Value = err & vbNewLine & "vsh:" & vsh & vbNewLine & "fsh:" & fsh
            End If
        End With
    End If
End Sub

Private Sub UserForm_Terminate()
    If Not myGL Is Nothing Then
        If vboPos(0) <> 0 Then myGL.DeleteBuffers 1, VarPtr(vboPos(0))
        If vboCol(0) <> 0 Then myGL.DeleteBuffers 1, VarPtr(vboCol(0))
        If prog <> 0 Then myGL.DeleteProgram prog
        myGL.PaintEnd
        Set myGL = Nothing
    End If
End Sub

