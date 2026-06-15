Attribute VB_Name = "Sample"
Option Explicit
Public Sub ShowForm()
    With VBA.UserForms.Add("Calculator")
        .BackColor = GetFormColor
        .init SelectedLibs("Ex"), SelectedLibs("Crash"), SelectedLibs("Move"), SelectedLibs("Break")
        .Show
    End With
End Sub
Public Sub ShowShaderForm()
    glShaderpoi.Show
End Sub
Private Function GetFormColor() As Long
    GetFormColor = ActiveSheet.Range("B2").Interior.Color
End Function
Private Function AllExLibs() As Collection
    Set AllExLibs = New Collection
    With AllExLibs
        .Add CFormPhysicsLogger
        .Add CFormPhysicsWsRenderer
        .Add CFormPhysicsFmRenderer
        .Add CFormPhysicsController
        .Add CFormPhysicsGLEffector
        .Add CFormPhysicsExtBreakable
    End With
End Function
Private Function AllEfLibs() As Collection
    Set AllEfLibs = New Collection
    With AllEfLibs
        .Add glShockWave
        .Add glExplosion
        .Add glHitNumber
        .Add glMoveTrail
        .Add glStatusVisualizer
        .Add glControlShatter
        .Add glShaderpoi
    End With
End Function
Private Function toDict(ByRef libs As Collection) As Object
    Dim tmp As Object, DICT: Set DICT = CreateObject("Scripting.Dictionary")
    For Each tmp In libs
        DICT.Add TypeName(tmp), tmp
    Next tmp
    Set toDict = DICT
End Function
Private Function AllExDict() As Object
    Set AllExDict = toDict(AllExLibs)
End Function
Private Function AllEfDict() As Object
    Set AllEfDict = toDict(AllEfLibs)
End Function
Public Function SelectedLibs(ByVal Tag As String)
    Dim i As Long, j As Long, shp As Shape, dc As Object, SelList(): ReDim SelList(30)
    If Tag = "Ex" Then Set dc = AllExDict Else Set dc = AllEfDict
    SelectedLibs = Empty
    For Each shp In ActiveSheet.Shapes
        With shp
            If .Name Like Tag & "-*" Then
                If .ControlFormat.Value = 1 Then
                    With .TextFrame.Characters
                        If dc.exists(.Text) Then Set SelList(j) = dc.Item(.Text): j = j + 1
                    End With
                End If
            End If
        End With
    Next shp
    If j > 0 Then
        ReDim Preserve SelList(j - 1)
        SelectedLibs = SelList
    End If
End Function
Public Sub ResetButtons()
    Dim shp As Shape, LibBuf As Object, lb As Variant
    Dim i As Long, j As Long, tofs As Double, ty As Double
    Dim b2x As Double, b2y As Double, b2w As Double, b2h As Double
    
    With ActiveSheet
        ActiveWindow.Zoom = 100
        .Columns("A:CZ").ColumnWidth = 2.5
        For Each shp In .Shapes
            shp.Delete
        Next shp
        
        With .Range("B2")
            b2x = .left
            b2y = .top
            b2w = .width
            b2h = .height
            .Interior.Color = RGB(230, 230, 230)
            With .Borders
                .LineStyle = xlContinuous
                .ColorIndex = xlAutomatic
                .TintAndShade = 0
                .weight = xlThin
            End With
        End With
        
        Set LibBuf = AllExDict
        
        With .Shapes
            
            With .AddFormControl(xlButtonControl, b2x + b2w + 5, b2y - 3, 90, 25)
                .OnAction = "ShowForm"
                .TextFrame.Characters.Text = "ShowForm"
                .DrawingObject.Font.size = 14
            End With
            
            tofs = b2y + b2h + 15
            
            i = 0
            For Each lb In LibBuf.keys()
                If lb Like "*GLEff*" Then
                    ty = tofs + (LibBuf.count) * 20
                Else
                    ty = tofs + i * 20
                    i = i + 1
                End If
                With .AddFormControl(xlCheckBox, b2x, ty, 150, 30)
                    .TextFrame.Characters.Text = lb
                    .Name = "Ex-" & lb
                End With
            Next lb
            
            With .AddFormControl(xlGroupBox, b2x - 5, tofs - 5, 150, 15 + (LibBuf.count + 1) * 20)
                .TextFrame.Characters.Text = "ICFormPhysicsEx"
            End With
            
            tofs = tofs + (LibBuf.count + 2) * 20
            
            Dim ct: ct = Array("Crash-", "Move-", "Break-")
            Set LibBuf = AllEfDict
            For j = LBound(ct) To UBound(ct)
                i = 0
                For Each lb In LibBuf.keys()
                    If (Not lb Like "*Shader*") Or (ct(j) Like "Move*") Then
                        With .AddFormControl(xlCheckBox, b2x + j * 170, tofs + i * 20, 150, 30)
                            .TextFrame.Characters.Text = lb
                            .Name = ct(j) & lb
                        End With
                        If lb Like "*Shader*" Then
                        
                            With .AddFormControl(xlButtonControl, 100 + b2x + j * 170, tofs + (i + 0.5) * 20 - 2, 30, 15)
                                .OnAction = "ShowShaderForm"
                                .TextFrame.Characters.Text = "test"
                                .DrawingObject.Font.size = 10.5
                            End With
                        
                        End If
                        
                        i = i + 1
                    End If
                Next lb
                With .AddFormControl(xlGroupBox, b2x - 5 + j * 170, tofs - 5, 150, (LibBuf.count + 2) * 20)
                    .TextFrame.Characters.Text = ct(j) & "ICFormPhysicsEf"
                End With
            Next j
            
            With .AddFormControl(xlGroupBox, b2x - 10, tofs - 45, 170 * 3 - 10, (LibBuf.count + 2) * 20 + 50)
                .TextFrame.Characters.Text = "GLExtensions"
            End With
        
        End With
    End With
End Sub



