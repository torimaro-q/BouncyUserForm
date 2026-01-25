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
Private Sub myCore_Move(ByRef X As Double, ByRef Y As Double, ByRef veloc As Double, time As Long)
    If (time - pt) < 5 Then Exit Sub
    pt = time
    tx = Clamp(CLng(X) + Rnd * 100 - 50, ScrL, ScrW)
    ty = Clamp(CLng(Y) + Rnd * 100 - 50, ScrT, ScrH)
    UpdateSheet SheetLut(tx, ty)
End Sub
Private Sub UpdateSheet(ByRef cellInfo As rc)
    With cellInfo
        If .isUsed Then Exit Sub
        If .R > 0 And .c > 0 Then
            With ws.Cells(.R, .c)
                .Value = symbols(CLng(Rnd() * 3))
                .Font.ColorIndex = (1 + CLng(Rnd() * 25))
            End With
        End If
    End With
    With buf(cnt)
        If .R > 0 And .c > 0 Then ws.Cells(.R, .c).Value = ""
    End With
    buf(cnt) = cellInfo
    cnt = cnt + 1
    If cnt > UBound(buf) Then cnt = 0
End Sub
Private Sub myCore_Stopped(X As Double, Y As Double, time As Long)
    Dim i As Long
    On Error Resume Next
    For i = LBound(buf) To UBound(buf)
        With buf(i)
            If .R > 0 And .c > 0 Then
                With ws.Cells(.R, .c)
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
Private Sub BuildSheetLut()
    Dim tmp As Range, Area As Range
    Dim R As Long, c As Long
    Dim ll As Long, LT As Long, ul As Long, ut As Long
    Dim v As Variant
    Dim i As Long, j As Long
    Set Area = ws.Range("A1:DA80")
    For Each tmp In Area
        With tmp
            R = .Row
            c = .Column
            v = .Value
            ll = CLng(.left)
            LT = CLng(.top)
            ul = CLng(.left + .width)
            ut = CLng(.top + .height)
            If ul <= ScrW And ut <= ScrH Then
                For i = ll To ul
                    For j = LT To ut
                        With SheetLut(i, j)
                            .R = R
                            .c = c
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
Private Sub ICFormPhysicsEx_Terminate()
    Set myCore = Nothing
End Sub
