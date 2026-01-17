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
Private WithEvents myCore As CFormPhysics
Attribute myCore.VB_VarHelpID = -1
Private ws As Worksheet
Private Type rc
    r As Long
    c As Long
    isUsed As Boolean
End Type
Private cnt As Long
Private symbols(3) As String
Private buf() As rc
Private SheetLut() As rc
Private ScrL As Long, ScrT As Long, ScrW As Long, ScrH As Long, tx As Long, ty As Long
Private pt As Long
Private Function Clamp(ByRef v As Long, ByRef minV As Long, ByRef maxV As Long) As Long
    If v < minV Then
        Clamp = minV
    ElseIf v > maxV Then
        Clamp = maxV
    Else
        Clamp = v
    End If
End Function
Private Sub myCore_Move(ByRef x As Double, ByRef y As Double, ByRef veloc As Double, time As Long)
    If (time - pt) < 5 Then Exit Sub
    pt = time
    tx = Clamp(CLng(x) + Rnd * 100 - 50, ScrL, ScrW)
    ty = Clamp(CLng(y) + Rnd * 100 - 50, ScrT, ScrH)
    UpdateSheet SheetLut(tx, ty)
End Sub
Private Sub UpdateSheet(ByRef cellInfo As rc)
    With cellInfo
        If .isUsed Then Exit Sub
        If .r > 0 And .c > 0 Then
            With ws.Cells(.r, .c)
                .Value = symbols(CLng(Rnd() * 3))
                .Font.ColorIndex = (1 + CLng(Rnd() * 25))
            End With
        End If
    End With
    With buf(cnt)
        If .r > 0 And .c > 0 Then ws.Cells(.r, .c).Value = ""
    End With
    buf(cnt) = cellInfo
    cnt = cnt + 1
    If cnt > UBound(buf) Then cnt = 0
End Sub
Private Sub myCore_Stopped(x As Double, y As Double, time As Long)
    Dim i As Long
    On Error Resume Next
    For i = LBound(buf) To UBound(buf)
        With buf(i)
            If .r > 0 And .c > 0 Then
                With ws.Cells(.r, .c)
                    .Value = ""
                    .Font.ColorIndex = 1
                End With
            End If
        End With
    Next i
    cnt = 0
End Sub
Private Sub ClearCell(ByRef rc As rc)
End Sub
Public Property Get CreateInstance(Optional ByRef core As CFormPhysics = Nothing) As CFormPhysicsWsRenderer
    'Constructor using the default instance
    Dim ch As CFormPhysicsWsRenderer: Set ch = New CFormPhysicsWsRenderer
    If Not core Is Nothing Then ch.Init core
    Set CreateInstance = ch
End Property
Public Sub Init(ByRef core As CFormPhysics)
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
    ReDim SheetLut(ScrL To ScrL + ScrW, ScrT To ScrT + ScrH)
    ReDim buf(80)
    cnt = 1
    BuildSheetLut
End Sub
Private Sub BuildSheetLut()
    Dim tmp As Range, Area As Range
    Dim r As Long, c As Long
    Dim ll As Long, lt As Long, ul As Long, ut As Long
    Dim v As Variant
    Dim i As Long, j As Long
    Set Area = ws.Range("A1:DA80")
    For Each tmp In Area
        With tmp
            r = .Row
            c = .Column
            v = .Value
            ll = CLng(.Left)
            lt = CLng(.Top)
            ul = CLng(.Left + .Width)
            ut = CLng(.Top + .Height)
            If ul <= ScrW And ut <= ScrH Then
                For i = ll To ul
                    For j = lt To ut
                        With SheetLut(i, j)
                            .r = r
                            .c = c
                            .isUsed = (v <> Empty)
                        End With
                    Next j
                Next i
            End If
        End With
    Next tmp
End Sub
