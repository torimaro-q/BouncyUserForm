VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CFormPhysicsWsRenderer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Implements ICFormPhysicsEx
Private WithEvents myCore As CFormPhysics
Attribute myCore.VB_VarHelpID = -1
Private ws As Worksheet
Private Type rc
    R As Long
    C As Long
    isUsed As Boolean
End Type
Private Const OFS = 50
Private Const MGN = OFS * 2
Private Const BUF_CNT = 80
Private cnt As Long
Private symbols(3) As String
Private buf() As rc
Private SheetLut() As rc
Private ScrL As Long, ScrT As Long, ScrW As Long, ScrH As Long, tx As Long, ty As Long, rd As Double
Private pt As Long
Private Sub myCore_Move(ByRef X As Double, ByRef Y As Double, ByRef veloc As Double, time As Long)
    If (time - pt) < 10 Then Exit Sub
    pt = time
    rd = Rnd()
    tx = CLng(X + rd * MGN)
    ty = CLng(Y - rd * MGN)
    UpdateSheet SheetLut(tx, ty)
End Sub
Private Sub UpdateSheet(ByRef cellInfo As rc)
    With cellInfo
        If .isUsed Then GoTo continue
        If .R < 1 Then GoTo continue
        If .C < 1 Then GoTo continue
        With ws.Cells(.R, .C)
            .Value = symbols(CLng(rd * 3))
            .Font.ColorIndex = (1 + CLng(rd * 25))
        End With
    End With
    buf(cnt) = cellInfo
    cnt = cnt + 1
    If cnt > BUF_CNT Then cnt = 0
    With buf(cnt)
        If .R < 1 Then GoTo continue
        If .C < 1 Then GoTo continue
        ws.Cells(.R, .C).Value = ""
    End With
continue:
End Sub
Private Sub myCore_Stopped(X As Double, Y As Double, time As Long)
    Call ClearCells
End Sub
Private Sub ClearCells()
Application.ScreenUpdating = False
    Dim i As Long
    On Error Resume Next
    For i = 0 To BUF_CNT
        With buf(i)
            If .R < 1 Then GoTo continue
            If .C < 1 Then GoTo continue
            With ws.Cells(.R, .C)
                .Value = ""
                .Font.ColorIndex = 1
            End With
        End With
continue:
    Next i
    cnt = 0
Application.ScreenUpdating = True
End Sub
Private Sub BuildSheetLut()
    Dim tmp As Range, Area As Range
    Dim R As Long, C As Long
    Dim ll As Long, LT As Long, ul As Long, ut As Long
    Dim v As Variant, i As Long, j As Long
    Set Area = ws.Range("A1:DA80")
    For Each tmp In Area
        With tmp
            R = .Row
            C = .Column
            v = .Value
            ll = CLng(.left) + OFS
            LT = CLng(.top) + OFS
            ul = CLng(.left + .width) + OFS
            ut = CLng(.top + .height) + OFS
            If ul <= ScrW And ut <= ScrH Then
                For i = ll To ul
                    For j = LT To ut
                        With SheetLut(i, j)
                            .R = R
                            .C = C
                            .isUsed = (v <> Empty)
                        End With
                    Next j
                Next i
            End If
        End With
    Next tmp
End Sub
Private Property Get ICFormPhysicsEx_CreateInstance() As ICFormPhysicsEx
    Set ICFormPhysicsEx_CreateInstance = New CFormPhysicsWsRenderer
End Property
Private Sub ICFormPhysicsEx_init(core As CFormPhysics, Optional params As Variant = Empty)
    Application.Calculation = xlManual
    Set myCore = core
    Set ws = ActiveSheet
    ActiveWindow.Zoom = 100
    With core
        ScrL = .ScrLeft * .Px2Tw
        ScrT = .ScrTop * .Px2Tw
        ScrW = .ScrWidth * .Px2Tw
        ScrH = .ScrHeight * .Px2Tw
    End With
    symbols(0) = "'Åú"
    symbols(1) = "'Åü"
    symbols(2) = "'Å°"
    symbols(3) = "'Å£"
    ReDim SheetLut(ScrL - MGN To ScrL + ScrW + MGN, ScrT - MGN To ScrT + ScrH + MGN)
    ReDim buf(BUF_CNT)
    cnt = 1
    BuildSheetLut
End Sub
Private Sub ICFormPhysicsEx_Terminate()
    Application.Calculation = xlAutomatic
    Call ClearCells
    Set myCore = Nothing
End Sub
