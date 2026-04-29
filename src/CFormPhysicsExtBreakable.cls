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
    cx As Double
    cy As Double
    time As Long
    Obj As ICFormPhysicsEx
    shp As Shape
End Type
Private Const H_OFS As Double = 200, E_SIZE As Double = 45
Private exKeys As Variant, exCnt As Long, myName As String, exts() As PhEx, i As Long, d As Double, ptime As Long
Private Sub myCore_Move(ByRef X As Double, ByRef Y As Double, ByRef veloc As Double, time As Long)
    For i = 0 To exCnt
        With exts(i)
            If .state = 0 Then
                If Abs(X - .cx) < E_SIZE Then
                    If Abs(Y - .cy) < E_SIZE Then
                        Application.ScreenUpdating = False
                        .Obj.Terminate
                        .state = 1
                        .time = time
                        With myCore
                            .VX = .VX * -0.9
                            .VY = .VY * -0.9
                            .UpdateVeloc
                            .ApplyDamage
                        End With
                        .shp.Fill.ForeColor.RGB = &HFF
                        Application.ScreenUpdating = True
                    End If
                End If
            ElseIf .state = 1 Then
                If Abs(ptime - time) > 100 Then
                    Application.ScreenUpdating = False
                    d = Abs(.time - time)
                    If d > 300 Then
                        .state = 2
                        .shp.Delete
                    Else
                        With .shp
                            .Fill.Transparency = d * 0.0033
                            .top = .top + 1
                        End With
                    End If
                    Application.ScreenUpdating = True
                    ptime = time
                End If
            End If
        End With
    Next i
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
    Set myCore = Nothing
End Sub
Private Sub myCore_Initialized()
    exKeys = myCore.ex.keys()
    exCnt = UBound(exKeys)
    ReDim exts(exCnt) As PhEx
    Dim n As Variant
    With myCore
        Dim baseL As Double: baseL = .Px2Tw * (0.9 * .ScrWidth - 2 * E_SIZE)
        Dim baseH As Double: baseH = .Px2Tw * (0.9 * .ScrHeight - 2 * E_SIZE - H_OFS)
    End With
    With ws
        For Each n In .Shapes
            If n.Name Like "*ex_shp_*" Then n.Delete
        Next n
        With .Shapes
            For i = 0 To exCnt
                n = exKeys(i)
                With exts(i)
                    .state = 0
                    If myName = n Then .state = 2
                    If .state = 0 Then
                        Set .Obj = myCore.ex.Item(n)
                        Set .shp = ws.Shapes.AddShape(msoShapeHexagon, E_SIZE + baseL * Rnd(), baseH * Rnd(), E_SIZE * 2, E_SIZE * 2)
                        .cx = .shp.left + .shp.width * 0.5
                        .cy = .shp.top + .shp.height * 0.5 + (H_OFS * myCore.Px2Tw)
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
            .Font.Bold = msoTrue
            .Font.size = 12
            .Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
            .ParagraphFormat.Alignment = msoAlignCenter
        End With
    End With
End Sub

