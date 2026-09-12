VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm4 
   Caption         =   "Seleção de Gráfico"
   ClientHeight    =   2040
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4215
   OleObjectBlob   =   "UserForm4.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub UserForm_Initialize()
CommandButton1.SetFocus
End Sub

Private Sub CommandButton1_Click()

Select Case True
    Case OptionButton1 = True
        PLANREF = "SAINTE"
    Case OptionButton2 = True
        PLANREF = "ROAMING"
End Select

Unload Me

End Sub

Private Sub OptionButton1_Click()
CommandButton1.SetFocus
End Sub

Private Sub OptionButton2_Click()
CommandButton1.SetFocus
End Sub


