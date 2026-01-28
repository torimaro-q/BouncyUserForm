VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} CFormPhysicsController 
   Caption         =   "Controller"
   ClientHeight    =   2895
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5340
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
Private Initialized As Boolean, ptime As Long
Private Sub LabelA_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    With LabelA
        If Button = 1 Then
            .left = .left + x - .width * 0.5
            .top = .top + y - .height * 0.5
            Me.Repaint
            If myCore.isAnimation Then Exit Sub
            ptime = 0
            Call ApplyAnalogVelocity(100)
            myCore.Launch
        End If
    End With
End Sub
Private Sub ApplyAnalogVelocity(ByRef time)
    Dim tx As Double, ty As Double
    With LabelA
            tx = .left - 34
            ty = .top - 35
            With myCore
                .VY = .VY + ty * time * 0.02
                .VX = .VX + tx * time * 0.02
            End With
    End With
End Sub
Private Sub myCore_Move(x As Double, y As Double, veloc As Double, time As Long)
    If ptime > 0 Then Call ApplyAnalogVelocity(time - ptime)
    ptime = time
End Sub
Private Sub CommandButtonB_Click()
    With myCore
        .VY = .VY + 2000
        ptime = 0
        .Launch
    End With
End Sub
Private Sub CommandButtonL_Click()
    With myCore
        .VY = .VY - 100
        .VX = .VX - 1500
        ptime = 0
        .Launch
    End With
End Sub
Private Sub CommandButtonR_Click()
    With myCore
        .VY = .VY - 100
        .VX = .VX + 1500
        ptime = 0
        .Launch
    End With
End Sub
Private Sub CommandButtonT_Click()
    With myCore
        .VY = .VY - 200
        ptime = 0
        .Launch
    End With
End Sub
Private Sub LabelA_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    With LabelA
        .left = 34
        .top = 35
        Me.Repaint
    End With
End Sub
Private Sub Frame1_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    Dim tx, ty
    With Frame1
        tx = x - .width * 0.5
        ty = y - .height * 0.5
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
