VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} CFormPhysicsExtBreakable 
   ClientHeight    =   495
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   2385
   OleObjectBlob   =   "CFormPhysicsExtBreakable.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "CFormPhysicsExtBreakable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit
Private Type RECT
    left As Long
    top As Long
    right As Long
    bottom As Long
End Type
Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal ms As Long)
Private Declare PtrSafe Function GetTickCount Lib "kernel32" () As Long
Private Declare PtrSafe Function GetWindowRect Lib "user32" (ByVal hWnd As LongPtr, lpRect As RECT) As Long
Private Declare PtrSafe Function MoveWindow Lib "user32" (ByVal hWnd As LongPtr, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Private Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As LongPtr
Private Declare PtrSafe Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWndParent As LongPtr, ByVal hWndChildAfter As LongPtr, ByVal lpszClass As String, ByVal lpszWindow As String) As LongPtr
Private Const MIN_RIBBON As String = "MinimizeRibbon"
Private Const LOGPIXELSX As Long = 88
Private ox As Double, oy As Double, moji As String
Private GL As Object, tHwnd As LongPtr, rc As RECT
Private ww As Double, wh As Double
Private apHwnd As LongPtr, dkHwnd As LongPtr, e2Hwnd As LongPtr
Private busy As Boolean, isSkip As Boolean
'###########################################################################################################
Implements ICFormPhysicsEx
Private ws As Worksheet
Private WithEvents myCore As CFormPhysics
Attribute myCore.VB_VarHelpID = -1
Private Type PhEx
    state As Long
    lx As Double
    rx As Double
    ty As Double
    by As Double
    time As Long
    obj As ICFormPhysicsEx
    shp As Shape
End Type
Private mylock As Boolean
Private Const H_OFS As Double = 200, E_SIZE As Double = 45
Private exKeys As Variant, exCnt As Long, myName As String, exts() As PhEx, i As Long, ptime As Long, flg As Boolean, d As Double
Private Sub myCore_Move(ByRef X As Double, ByRef Y As Double, ByRef veloc As Double, time As Long)
    If flg Then DoEvents
    flg = False
    For i = 0 To exCnt
        With exts(i)
            If .state = 2 Then
                .state = 3
                .obj.Terminate
                flg = True
            ElseIf .state = 1 Then
                d = time - .time
                If d > 300 Then
                    .state = 2
                    .shp.Visible = msoFalse
                Else
                    With .shp
                        .Fill.Transparency = d * 0.0033
                        .top = .top + 1
                    End With
                End If
                flg = True
            ElseIf .state <= 0 Then
                If X < .lx Then GoTo continue
                If X > .rx Then GoTo continue
                If Y < .ty Then GoTo continue
                If Y > .by Then GoTo continue
                .state = .state + 1
                If .state = 1 Then
                    .shp.Fill.ForeColor.RGB = &HDD
                Else
                    .shp.Fill.ForeColor.RGB = &H1099DD
                End If
                .time = time
                With myCore
                    .VX = .VX * -1.5
                    .VY = .VY * -1.5
                    .UpdateVeloc
                    .ApplyDamage
                End With
                flg = True
            End If
        End With
continue:
    Next i
    If flg Then DoEvents
End Sub
Private Sub myCore_Crash(X As Double, Y As Double, dmg As Double, time As Long)
    Dim i As Long, flg As Boolean
    flg = True
    For i = 0 To exCnt
        If exts(i).state < 3 Then flg = False
    Next i
    If flg Then
        If mylock = False Then
            mylock = True
            ThisWorkbook.isAddin = True
            Me.Show vbModal
        End If
    End If
End Sub
Private Sub ClearShapes()
On Error GoTo err
    With ws
        Dim n As Variant
        For Each n In .Shapes
            If n.Name Like "*ex_shp_*" Then n.Delete
        Next n
    End With
err:
End Sub
Private Sub CreateShapes()
    With ws
        Dim n As Variant
        With myCore
            Dim baseL As Double: baseL = .Px2Tw * (0.9 * .ScrWidth - 2 * E_SIZE)
            Dim baseH As Double: baseH = .Px2Tw * (0.9 * .ScrHeight - 2 * E_SIZE - H_OFS)
        End With
        With .Shapes
            For i = 0 To exCnt
                n = exKeys(i)
                With exts(i)
                    .state = -1
                    If myName = n Then .state = 10
                    If .state <= 0 Then
                        Set .obj = myCore.ex.Item(n)
                        Set .shp = ws.Shapes.AddShape(msoShapeHexagon, E_SIZE + baseL * Rnd(), baseH * Rnd(), E_SIZE * 2, E_SIZE * 2)
                        .lx = -E_SIZE + .shp.left + .shp.width * 0.5
                        .ty = -E_SIZE + .shp.top + .shp.height * 0.5 + (H_OFS * myCore.Px2Tw)
                        .rx = E_SIZE + .shp.left + .shp.width * 0.5
                        .by = E_SIZE + .shp.top + .shp.height * 0.5 + (H_OFS * myCore.Px2Tw)
                        With .shp
                            .Name = "ex_shp_" & n
                            .TextFrame2.TextRange.Text = Replace(n, "CFormPhysics", "")
                        End With
                        ApplyCoolTheme .shp
                    End If
                End With
            Next i
        End With
    End With
End Sub
Private Sub myCore_Initialized()
    exKeys = myCore.ex.keys()
    exCnt = UBound(exKeys)
    ReDim exts(exCnt) As PhEx
    Call ClearShapes
    Call CreateShapes
    flg = False
End Sub
Private Sub ApplyCoolTheme(ByRef target As Shape)
    With target.Fill
        .Visible = msoTrue
        .ForeColor.RGB = RGB(128, 128, 128)
        .BackColor.RGB = RGB(0, 0, 0)
    End With
    With target.Line
        .Visible = msoTrue
        .ForeColor.RGB = RGB(255, 255, 255)
        .Transparency = 0.2
        .weight = 2
    End With
    With target.Shadow
        .Visible = msoTrue
        .style = msoShadowStyleOuterShadow
        .Blur = 8
        .OffsetX = 3
        .OffsetY = 3
        .ForeColor.RGB = RGB(0, 0, 0)
        .Transparency = 0.5
    End With
    With target.ThreeD
        .Visible = msoTrue
        .Depth = 6
        .BevelTopType = msoBevelCircle
        .BevelTopDepth = 4
    End With
    With target.TextFrame2
        .WordWrap = msoFalse
        .VerticalAnchor = msoAnchorMiddle
        With .TextRange
            With .Font
                .Bold = msoTrue
                .size = 12
                .Fill.ForeColor.RGB = RGB(255, 255, 255)
            End With
            .ParagraphFormat.Alignment = msoAlignCenter
        End With
    End With
End Sub
Private Property Get ICFormPhysicsEx_CreateInstance() As ICFormPhysicsEx
    Set ICFormPhysicsEx_CreateInstance = New CFormPhysicsExtBreakable
End Property
Private Sub ICFormPhysicsEx_init(core As CFormPhysics, Optional params As Variant = Empty)
    Set myCore = core
    Set ws = ActiveSheet
    ActiveWindow.Zoom = 100
    myName = TypeName(Me)
    mylock = False
End Sub
Private Sub ICFormPhysicsEx_Terminate()
    ThisWorkbook.isAddin = False
    Set myCore = Nothing
    Call ClearShapes
End Sub
'###########################################################################################################
Private Sub Label1_Click() 'SKIP
    isSkip = True
End Sub
Private Sub StafRoll()
    Dim i As Long, arr, tmp, Name, tt, pt
    arr = Array("Producer", "Director", "Main Programmer", "Art Director", "Motion Director", "Thank you for playing BouncyUserform")
    Name = "torimaro-q"
    For Each tmp In arr
        oy = 0
        Do Until oy < -wh
            If isSkip Then Exit For
            If i = UBound(arr) Then Exit For
            pt = tt: tt = GetTickCount
            oy = oy - Abs(pt - tt) * 0.3
            If pt > 0 Then Paint (CStr(tmp & "  " & Name))
            If Rnd() > 0.9 Then DoEvents
        Loop
        i = i + 1
    Next tmp
    oy = -wh * 0.5
    Call Paint(CStr(arr(UBound(arr))))
    Sleep 2000
    DoEvents
End Sub
Private Sub Paint(ByRef tstr As String)
    If busy Then Exit Sub
    busy = True
    With GL
        .Clear GL_COLOR_BUFFER_BIT Or GL_DEPTH_BUFFER_BIT
        .MatrixMode GL_PROJECTION
            .LoadIdentity
            .Ortho2D 0, ww, wh, 0
        .MatrixMode GL_MODELVIEW
            .LoadIdentity
            .PushMatrix
                If moji <> tstr Then
                    moji = tstr
                    Dim chr() As Byte: chr = StrConv(tstr, vbFromUnicode)
                    If UBound(chr) > 0 Then
                        .listbase 0
                        If .IsList(10) Then .DeleteLists 10, 1
                        .listbase 0
                        .NewList 10, GL_COMPILE
                        .listbase 2000
                        .Color4f 1, 1, 1, 1
                        .CallLists UBound(chr) + 1, GL_UNSIGNED_BYTE, VarPtr(chr(0))
                        .EndList
                    End If
                End If
                .RasterPos2d ww * 0.2 + ox, wh * 1 + oy
                .CallList 10
            .PopMatrix
        .SwapBuffers
    End With
    busy = False
End Sub
Private Sub UserForm_Activate()
    If GL Is Nothing Then
        Me.left = 30
        Me.top = 80
        isSkip = False
        busy = True
        SwMode True
        With Application
            apHwnd = .hWnd
            If myCore Is Nothing Then
                ww = .width * 1.33
                wh = .height * 1.33
            Else
                ww = .width / myCore.Px2Tw
                wh = .height / myCore.Px2Tw
            End If
        End With
        If apHwnd = 0 Then GoTo err
        dkHwnd = FindWindowEx(apHwnd, 0&, "XLDESK", vbNullString)
        e2Hwnd = FindWindowEx(apHwnd, 0&, "EXCEL2", vbNullString)
        If e2Hwnd = 0 Then GoTo err
        Call GetWindowRect(e2Hwnd, rc)
        Call MoveWindow(e2Hwnd, 0, 0, ww, wh, 1)
        tHwnd = e2Hwnd
        If tHwnd = 0 Then GoTo err
        On Error GoTo err
        Set GL = Application.Run("GenOpenGL")
        With GL
            .hWnd = tHwnd
            .PaintStart
            .Viewport 0, 0, ww, wh
            .ClearColor 0.1, 0.1, 0.1, 1
            .Enable &HB71& 'GL_DEPTH_TEST
        End With
        busy = False
        Call StafRoll
    End If
err:
    Unload Me
End Sub
Private Sub UserForm_Terminate()
    busy = True
    With rc
        If e2Hwnd <> 0 Then Call MoveWindow(e2Hwnd, .left, .top, .right - .left, .bottom - .top, 1)
    End With
    If Not GL Is Nothing Then GL.PaintEnd
    Set GL = Nothing
    SwMode False
    Call ICFormPhysicsEx_Terminate
End Sub
Private Sub SwMode(isAddin As Boolean) 'to disable workbook rendering
    With Application
        With .CommandBars
            If .GetPressedMso(MIN_RIBBON) = Not isAddin Then .ExecuteMso MIN_RIBBON
        End With
        DoEvents
        .ScreenUpdating = Not isAddin
        .EnableEvents = Not isAddin
        .DisplayAlerts = Not isAddin
        .DisplayStatusBar = Not isAddin
        .PrintCommunication = Not isAddin
    End With
    ThisWorkbook.isAddin = isAddin
End Sub


