VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CFormPhysicsLogger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Implements ICFormPhysicsEx
Private Const BUF_SIZE As Long = 10000
Private WithEvents myCore As CFormPhysics
Attribute myCore.VB_VarHelpID = -1
Private ws As Worksheet
Private myChart As Chart
Private chartObj As ChartObject
Private Type POINT
    x As Double
    y As Double
    dmg As Double
End Type
Private Type SeriesType
    cnt As Long
    xvalues() As Double
    yvalues() As Double
End Type
Private buf() As POINT
Private mySeries() As SeriesType
Private pcnt As Long
Private cnt As Long
Private dmg As Double
Private pdmg As Double
Private interp() As Double, interpCount As Long
Private Sub UpdatePoint(ByRef x As Double, ByRef y As Double, Optional ByRef Damage As Double = 0)
    With buf(cnt)
        .x = x
        .y = y
        .dmg = Damage
    End With
    If Damage > 0 Then HandleDamageUpdate Damage
    cnt = cnt + 1
    If cnt > UBound(buf) Then cnt = 0
End Sub
Private Sub myCore_Move(ByRef x As Double, ByRef y As Double, ByRef veloc As Double, ByRef time As Long)
    UpdatePoint x, y
End Sub
Private Sub myCore_Crash(ByRef x As Double, ByRef y As Double, ByRef dmg As Double, ByRef time As Long)
    UpdatePoint x, y, dmg
End Sub
Private Sub myCore_Started(ByRef x As Double, ByRef y As Double, ByRef time As Long)
    Reset
    UpdatePoint x, y
End Sub
Private Sub myCore_Stopped(ByRef x As Double, ByRef y As Double, ByRef time As Long)
    UpdatePoint x, y
End Sub
Private Sub HandleDamageUpdate(ByRef Damage As Double)
    Dim i As Long, HCnt As Long, delta As Long
    Dim dc As Long, OFS As Long
    delta = cnt - pcnt
    If delta < 0 Then delta = delta + UBound(buf)
    For i = LBound(interp) To UBound(interp)
        dc = CLng(delta * interp(i))
        If dc > 1 Then
            HCnt = pcnt + dc
            If HCnt > UBound(buf) Then HCnt = HCnt - UBound(buf)
            AppendToSeries 0, buf(HCnt).x, buf(HCnt).y
        End If
    Next i
    dmg = buf(cnt).dmg
    OFS = 1 + Fix((dmg - pdmg) * 0.2)
    If OFS > 4 Then OFS = 4
    AppendToSeries 0, buf(cnt).x, buf(cnt).y
    AppendToSeries OFS, buf(cnt).x, buf(cnt).y
    UpdateChartSeries 0
    UpdateChartSeries OFS
    pdmg = dmg
    pcnt = cnt
    myChart.ChartArea.Select 'for excel 2024
    Application.Calculate 'In excel 2024, chart updates are no longer triggered.
End Sub
Private Sub AppendToSeries(ByVal index As Long, ByVal x As Double, ByVal y As Double)
    With mySeries(index)
        ReDim Preserve .xvalues(0 To .cnt)
        ReDim Preserve .yvalues(0 To .cnt)
        .xvalues(.cnt) = x
        .yvalues(.cnt) = y
        .cnt = .cnt + 1
    End With
End Sub
Private Sub UpdateChartSeries(ByVal seriesIndex As Long)
    With myChart.SeriesCollection.Item(seriesIndex + 1)
        .xvalues = mySeries(seriesIndex).xvalues
        .values = mySeries(seriesIndex).yvalues
    End With
End Sub
Public Sub Reset()
    cnt = 0
    pcnt = 0
    pdmg = 0
    ReDim interp(1 To interpCount) As Double
    Dim stp As Double, i As Long
    stp = 1 / (interpCount + 1)
    For i = 1 To UBound(interp)
        interp(i) = stp * i
    Next i
    ReDim buf(0 To BUF_SIZE) As POINT
End Sub
Private Sub ClearCharts()
    Dim tcht As ChartObject
    If ws Is Nothing Then Exit Sub
    For Each tcht In ws.ChartObjects
        If tcht.name Like "*FormPhysic*" Then
            tcht.Delete
        End If
    Next tcht
End Sub
Private Sub CreateChart()
    Dim i As Long, nw As Series, AX As Axes, labels As Variant
    labels = Array("Y", "dmg<1", "dmg<5", "dmg<10", "10<dmg")
    ReDim mySeries(LBound(labels) To UBound(labels))
    For i = LBound(mySeries) To UBound(mySeries)
        ReDim mySeries(i).xvalues(0 To 0)
        ReDim mySeries(i).yvalues(0 To 0)
        mySeries(i).cnt = 0
    Next i
    With myCore
        Set chartObj = ws.ChartObjects.Add(left:=5, top:=5, width:=(.ScrWidth * .Px2Tw) * 0.3, height:=(.ScrHeight * .Px2Tw) * 0.3)
    End With
    With chartObj
        .name = "CFormPhysicLogScatter"
        Set myChart = .Chart
        With myChart
            .ChartType = xlXYScatter
            .HasLegend = True
            For i = LBound(mySeries) To UBound(mySeries)
                Set nw = .SeriesCollection.NewSeries
                nw.name = labels(i)
                If i = 0 Then
                    nw.ChartType = xlXYScatterLinesNoMarkers
                    nw.format.Line.Weight = 0.01
                Else
                    nw.ChartType = xlXYScatter
                    nw.MarkerSize = 2 + (i - 1) * 3
                    nw.MarkerStyle = xlMarkerStyleCircle
                End If
            Next i
            Set AX = .Axes
            ConfigureAxes AX
        End With
    End With
    ApplyBrilliantTemplate myChart
End Sub
Private Sub ConfigureAxes(ByRef AX As Axes)
    With AX.Item(1) ' X axis
        .MaximumScale = (myCore.ScrLeft + myCore.ScrWidth) * myCore.Px2Tw
        .MinimumScale = myCore.ScrLeft * myCore.Px2Tw
        .MaximumScaleIsAuto = False
        .MinimumScaleIsAuto = False
    End With
    With AX.Item(2) ' Y axis
        .ReversePlotOrder = True
        .MaximumScale = (myCore.ScrTop + myCore.ScrHeight) * myCore.Px2Tw
        .MinimumScale = myCore.ScrTop * myCore.Px2Tw
        .MaximumScaleIsAuto = False
        .MinimumScaleIsAuto = False
    End With
End Sub
Public Sub ApplyBrilliantTemplate(ByVal ch As Chart)
    Dim colors As Variant
    Dim i As Long
    colors = Array(RGB(0, 200, 255), RGB(255, 80, 80), RGB(255, 150, 0), RGB(255, 255, 0), RGB(0, 255, 120))
    With ch
        .ChartArea.format.Fill.ForeColor.RGB = RGB(20, 20, 20)
        .PlotArea.format.Fill.ForeColor.RGB = RGB(20, 20, 20)
    End With
    With ch.Axes(xlValue).MajorGridlines.format.Line
        .Visible = msoTrue
        .ForeColor.RGB = RGB(60, 60, 60)
        .Weight = 0.25
    End With
    With ch.Axes(xlCategory).MajorGridlines.format.Line
        .Visible = msoTrue
        .ForeColor.RGB = RGB(60, 60, 60)
        .Weight = 0.25
    End With
    With ch.Axes(xlCategory).format.Line
        .ForeColor.RGB = RGB(200, 200, 200)
        .Weight = 0.75
    End With
    With ch.Axes(xlValue).format.Line
        .ForeColor.RGB = RGB(200, 200, 200)
        .Weight = 0.75
    End With
    For i = 1 To ch.SeriesCollection.Count
        With ch.SeriesCollection(i)
            With .format.Glow
                .Color.RGB = colors(i - 1)
                .Radius = 5
                .Transparency = 0.8
            End With
            .format.Fill.ForeColor.RGB = colors(i - 1)
        End With
    Next i
    With ch.Legend
        .Position = xlLegendPositionTop
        .format.Fill.Visible = msoFalse
        .format.Line.Visible = msoFalse
        .Font.Color = RGB(230, 230, 230)
    End With
    With ch
        .HasTitle = True
        .ChartTitle.Text = "CFormPhysics Trajectory Log"
        .ChartTitle.Font.Color = RGB(240, 240, 240)
        .ChartTitle.Font.size = 14
        .ChartTitle.Font.name = "Segoe UI Semibold"
    End With
    With ch.Axes(xlCategory).format.Line
        .ForeColor.RGB = RGB(200, 200, 200)
        .Weight = 0.75
    End With
    With ch.Axes(xlValue).format.Line
        .ForeColor.RGB = RGB(200, 200, 200)
        .Weight = 0.75
    End With
    With ch.Axes(xlCategory).TickLabels
        .Font.Color = RGB(220, 220, 220)
        .Font.size = 9
    End With
    With ch.Axes(xlValue).TickLabels
        .Font.Color = RGB(220, 220, 220)
        .Font.size = 9
    End With
End Sub
Private Property Get ICFormPhysicsEx_CreateInstance() As ICFormPhysicsEx
    Set ICFormPhysicsEx_CreateInstance = New CFormPhysicsLogger
End Property
Private Sub ICFormPhysicsEx_init(core As CFormPhysics, Optional params As Variant = Empty)
    Application.Calculation = xlCalculationManual
    Set myCore = core
    Set ws = ActiveSheet
    interpCount = 7
    Reset
    ClearCharts
    CreateChart
End Sub
Private Sub ICFormPhysicsEx_Terminate()
    Set myCore = Nothing
End Sub

