VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} CFormPhysicsController 
   Caption         =   "Controller"
   ClientHeight    =   5520
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   2745
   OleObjectBlob   =   "CFormPhysicsController.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "CFormPhysicsController"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Implements ICFormPhysicsEx
Private WithEvents myCore As CFormPhysics
Attribute myCore.VB_VarHelpID = -1
Private strArr(50) As String
Private Initialized As Boolean
Private Sub CommandButtonB_Click()
    With myCore
        .VY = .VY + 2000
        .Launch
    End With
End Sub
Private Sub CommandButtonL_Click()
    With myCore
        .VY = .VY - 100
        .VX = .VX - 1500
        .Launch
    End With
End Sub
Private Sub CommandButtonR_Click()
    With myCore
        .VY = .VY - 100
        .VX = .VX + 1500
        .Launch
    End With
End Sub
Private Sub CommandButtonT_Click()
    With myCore
        .VY = .VY - 200
        .Launch
    End With
End Sub
Private Sub LabelA_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim tx, ty
    With LabelA
        If Button = 1 Then
            .left = .left + X - .width * 0.5
            .top = .top + Y - .height * 0.5
            Me.Repaint
            tx = .left - 34
            ty = .top - 35
            With myCore
                .VY = .VY + ty * 2
                .VX = .VX + tx * 2
                .Launch
            End With
        End If
    End With
End Sub
Private Sub LabelA_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    With LabelA
        .left = 34
        .top = 35
        Me.Repaint
    End With
End Sub
Private Sub Frame1_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim tx, ty
    With Frame1
        tx = X - .width * 0.5
        ty = Y - .height * 0.5
    End With
    With myCore
        .VY = .VY + ty * 2
        .VX = .VX + tx * 2
        .Launch
    End With
End Sub
Private Property Get ICFormPhysicsEx_CreateInstance() As ICFormPhysicsEx
    Set ICFormPhysicsEx_CreateInstance = New CFormPhysicsController
End Property
Private Sub ICFormPhysicsEx_init(core As CFormPhysics, Optional params As Variant = Empty)
    Initialized = False
    Set myCore = core
    Me.Show
End Sub
Private Sub ICFormPhysicsEx_Terminate()
    Set myCore = Nothing
    Unload Me
End Sub


Private Sub UserForm_Activate()
On Error GoTo err
    If Initialized = True Then Exit Sub
    Dim baseColor As Long, panelColor As Long, keyColor As Long, cntrs As Variant, tmp As Variant
    If Rnd >= 0.5 Then
        baseColor = &H444986
        panelColor = &HCAFFFF
        keyColor = &H111111
    Else
        baseColor = &HDDDDDD
        panelColor = &H303030
        keyColor = &H111111
    End If
    With Me
        .BackColor = baseColor
        .Label1.BackColor = baseColor
        .Label2.BackColor = baseColor
        .Label3.BackColor = baseColor
        .Label5.BackColor = keyColor
        .CommandButtonB.BackColor = keyColor
        .CommandButtonL.BackColor = keyColor
        .CommandButtonR.BackColor = keyColor
        .CommandButtonT.BackColor = keyColor
        .Frame1.BackColor = panelColor
        .Frame2.BackColor = panelColor
        .Label6.BackColor = baseColor
        .LabelA.BackColor = keyColor
        .left = myCore.mFrmRaw.left - .width
    End With
err:
    Initialized = True
End Sub
