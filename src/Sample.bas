Attribute VB_Name = "Sample"
Option Explicit
Public Sub ShowForm()
    Dim cf As Object
    Dim ExLibs As Variant
    Dim CrashLibs As Variant
    Dim MoveLibs As Variant
    Dim BreakLibs As Variant
    
    Set cf = VBA.UserForms.Add("Calculator")
    
    ExLibs = SelectedLibs("Ex")
    CrashLibs = SelectedLibs("Crash")
    MoveLibs = SelectedLibs("Move")
    BreakLibs = SelectedLibs("Break")
    
    cf.init ExLibs, CrashLibs, MoveLibs, BreakLibs
    
    cf.Show
End Sub
Private Function AllExArray() As Variant
    AllExArray = Array(CFormPhysicsLogger, _
                       CFormPhysicsWsRenderer, _
                       CFormPhysicsFmRenderer, _
                       CFormPhysicsController, _
                       CFormPhysicsGLEffector, _
                       CFormPhysicsExtBreakable)
End Function
Private Function AllEfArray() As Variant
    AllEfArray = Array(glShockWave, _
                       glExplosion, _
                       glHitNumber, _
                       glMoveTrail, _
                       glStatusVisualizer, _
                       glControlShatter, _
                       glShaderTest)
End Function
Private Function toDict(arr) As Object
    Dim i As Long, DICT: Set DICT = CreateObject("Scripting.Dictionary")
    For i = LBound(arr) To UBound(arr)
        DICT.Add TypeName(arr(i)), arr(i)
    Next i
    Set toDict = DICT
End Function
Private Function AllExDict() As Object
    Set AllExDict = toDict(AllExArray)
End Function
Private Function AllEfDict() As Object
    Set AllEfDict = toDict(AllEfArray)
End Function
Public Function SelectedLibs(ByVal Tag As String)
    Dim ws As Worksheet: Set ws = ActiveSheet
    Dim i As Long, j As Long, shp As Shape, tkey As String, DICT
    Dim SelList(): ReDim SelList(20)
    If Tag = "Ex" Then Set DICT = AllExDict Else Set DICT = AllEfDict
    SelectedLibs = Empty
    For Each shp In ws.Shapes
        With shp
            If .Name Like Tag & "-*" Then
                If .ControlFormat.Value = 1 Then
                    tkey = .TextFrame.Characters.Text
                    If DICT.exists(tkey) Then
                        Set SelList(j) = DICT.Item(tkey)
                        j = j + 1
                    End If
                End If
            End If
        End With
    Next shp
    If j > 0 Then
        ReDim Preserve SelList(j - 1)
        SelectedLibs = SelList
    End If
End Function
Private Sub ResetButtons()
    Dim ws As Worksheet: Set ws = ActiveSheet
    Dim shp As Shape, LibBuf, lb
    Dim i As Long, j As Long, tofs As Double, ty As Double
    For Each shp In ws.Shapes
        shp.Delete
    Next shp
    Set LibBuf = AllExDict
    With ws.Shapes
        With .AddFormControl(xlButtonControl, 200, 55, 90, 25)
            .OnAction = "ShowForm"
            .TextFrame.Characters.Text = "ShowForm"
            .DrawingObject.Font.size = 16
        End With
        tofs = 60
        i = 0
        For Each lb In LibBuf.keys()
            If lb Like "*GLEff*" Then
                ty = tofs + (LibBuf.count) * 20
            Else
                ty = tofs + i * 20
                i = i + 1
            End If
            With .AddFormControl(xlCheckBox, 30, ty, 150, 30)
                .TextFrame.Characters.Text = lb
                .Name = "Ex-" & lb
            End With
        Next lb
        With .AddFormControl(xlGroupBox, 25, 55, 150, tofs - 30 + (LibBuf.count) * 20)
            .TextFrame.Characters.Text = "ICFormPhysicsEx"
        End With
        tofs = tofs + (LibBuf.count + 2) * 20
        Dim ct: ct = Array("Crash-", "Move-", "Break-")
        Set LibBuf = AllEfDict
        For j = LBound(ct) To UBound(ct)
            i = 0
            For Each lb In LibBuf.keys()
                With .AddFormControl(xlCheckBox, 30 + j * 170, tofs + i * 20, 150, 30)
                    .TextFrame.Characters.Text = lb
                    .Name = ct(j) & lb
                End With
                i = i + 1
            Next lb
            With .AddFormControl(xlGroupBox, 25 + j * 170, tofs - 5, 150, (LibBuf.count + 1) * 20)
                .TextFrame.Characters.Text = ct(j) & "ICFormPhysicsEf"
            End With
        Next j
        With .AddFormControl(xlGroupBox, 20, tofs - 45, 170 * 3 - 10, (LibBuf.count + 1) * 20 + 50)
            .TextFrame.Characters.Text = "GLExtensions"
        End With
    End With
End Sub



