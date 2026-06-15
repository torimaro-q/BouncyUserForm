VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Calculator 
   Caption         =   "Calculator"
   ClientHeight    =   4710
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4455
   OleObjectBlob   =   "Calculator.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "Calculator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private engine As CFormPhysics
Private Sub UserForm_Error(ByVal Number As Integer, ByVal Description As MSForms.ReturnString, ByVal SCode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, ByVal CancelDisplay As MSForms.ReturnBoolean)
    Unload Me
End Sub
Public Sub init(Optional ByVal ExLibs = Empty, Optional ByVal ChLibs = Empty, Optional ByVal MvLibs = Empty, Optional ByVal BkLibs = Empty)
    Set engine = New CFormPhysics
    Call engine.init(Me, ExLibs, ChLibs, MvLibs, BkLibs)
End Sub
Private Sub UserForm_Terminate()
On Error GoTo err
    engine.Terminate
err:
End Sub
'############################################################################################################################
Private Sub btnInput(ByVal inp As String)
    With Me.TextBox1
        .Value = .Value & inp
    End With
End Sub
Private Sub CommandButton15_Click()
On Error GoTo err
    Me.TextBox1 = Application.Evaluate(Me.TextBox1.Value)
err:
End Sub
Private Sub CommandButton16_Click()
    Me.TextBox1 = ""
End Sub
Private Sub CommandButton0_Click()
    btnInput "0"
End Sub
Private Sub CommandButton1_Click()
    btnInput "1"
End Sub
Private Sub CommandButton2_Click()
    btnInput "2"
End Sub
Private Sub CommandButton3_Click()
    btnInput "3"
End Sub
Private Sub CommandButton4_Click()
    btnInput "4"
End Sub
Private Sub CommandButton5_Click()
    btnInput "5"
End Sub
Private Sub CommandButton6_Click()
    btnInput "6"
End Sub
Private Sub CommandButton7_Click()
    btnInput "7"
End Sub
Private Sub CommandButton8_Click()
    btnInput "8"
End Sub
Private Sub CommandButton9_Click()
    btnInput "9"
End Sub
Private Sub CommandButton10_Click()
    btnInput "."
End Sub
Private Sub CommandButton11_Click()
    btnInput "+"
End Sub
Private Sub CommandButton12_Click()
    btnInput "-"
End Sub
Private Sub CommandButton13_Click()
    btnInput "*"
End Sub
Private Sub CommandButton14_Click()
    btnInput "/"
End Sub
