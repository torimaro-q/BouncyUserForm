VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CFormPhysicsEx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private WithEvents myCore As CFormPhysics
Attribute myCore.VB_VarHelpID = -1
Private ws As Worksheet
Private Type RC
    r As Long
    c As Long
End Type
Private cnt As Long, symbols(3) As String
Private buf() As RC, SheetLut() As RC
Private ScrL As Long, ScrT As Long, ScrW As Long, ScrH As Long
Private Sub myCore_Move(ByRef x As Double, ByRef y As Double, ByRef veloc As Double)
    Dim tx As Long, ty As Long
    Dim r As Long, c As Long
    tx = CLng(x) + Rnd * 250 - 125
    ty = CLng(y) + Rnd * 250 - 125
    If tx < ScrL Then
        tx = ScrL
    ElseIf tx > ScrW Then
        tx = ScrW
    End If
    If ty < ScrT Then
        ty = ScrT
    ElseIf ty > ScrH Then
        ty = ScrH
    End If
    With SheetLut(tx, ty)
        r = .r
        c = .c
    End With
    With ws
        If r > 0 And c > 0 Then .Cells(r, c).Value = symbols(CLng(Rnd() * 3))
        If buf(cnt).r > 0 And buf(cnt).c > 0 Then .Cells(buf(cnt).r, buf(cnt).c).Value = ""
    End With
    buf(cnt).r = r
    buf(cnt).c = c
    cnt = cnt + 1
    If cnt > UBound(buf) Then cnt = 0
End Sub
Private Sub myCore_Stopped(x As Double, y As Double)
    Dim i As Long
    On Error Resume Next
    For i = LBound(buf) To UBound(buf)
        With buf(i)
            If .r > 0 And .c > 0 Then
                ws.Cells(.r, .c).Value = ""
            End If
        End With
    Next i
    cnt = 0
End Sub
Public Property Get CreateInstance(Optional ByRef core As CFormPhysics = Nothing) As CFormPhysicsEx
    Dim ch As CFormPhysicsEx: Set ch = New CFormPhysicsEx
    If Not core Is Nothing Then ch.Init core
    Set CreateInstance = ch
End Property
Public Sub Init(ByRef core As CFormPhysics)
    Set myCore = core
    Set ws = ActiveSheet
    With core
        ScrL = .ScrLeft * .Px2Tw
        ScrT = .ScrTop * .Px2Tw
        ScrW = .ScrWidth * .Px2Tw
        ScrH = .ScrHeight * .Px2Tw
    End With
    symbols(0) = "'+"
    symbols(1) = "'-"
    symbols(2) = "'*"
    symbols(3) = "'#"
    ReDim SheetLut(ScrL To ScrL + ScrW, ScrT To ScrT + ScrH) As RC
    ReDim buf(150) As RC
    cnt = 1
    Call BuildSheetLut
End Sub
Private Sub BuildSheetLut()
    Dim tmp As Range, Area As Range
    Dim r As Long, c As Long
    Dim ll As Long, lt As Long, ul As Long, ut As Long
    Dim i As Long, j As Long
    Dim xlLeft As Long, xlTop As Long
    Set Area = ws.Range("A1:DA80")
    For Each tmp In Area
        With tmp
            r = .Row
            c = .Column
            ll = CLng(.Left)
            lt = CLng(.Top)
            ul = CLng((.Left + .Width))
            ut = CLng((.Top + .Height))
            If ul <= ScrW And ut <= ScrH Then
                For i = ll To ul
                    For j = lt To ut
                        SheetLut(i, j).r = r
                        SheetLut(i, j).c = c
                    Next j
                Next i
            End If
        End With
    Next tmp
End Sub

