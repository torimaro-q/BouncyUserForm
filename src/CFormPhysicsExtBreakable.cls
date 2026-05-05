VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CFormPhysicsExtBreakable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
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
                    .VX = .VX * -3
                    .VY = .VY * -3
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
Private Sub ClearShapes()
    With ws
        Dim n As Variant
        For Each n In .Shapes
            If n.Name Like "*ex_shp_*" Then n.Delete
        Next n
    End With
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
End Sub
Private Sub ICFormPhysicsEx_Terminate()
    Call ClearShapes
    Set myCore = Nothing
End Sub
