VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm5 
   Caption         =   "Seleção de Região para FCN7 LDN"
   ClientHeight    =   3120
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4710
   OleObjectBlob   =   "UserForm5.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
If CX1.Value = True Then Cells(72, 1) = 1
If CX2.Value = True Then Cells(72, 1) = 2
Unload Me
End Sub

Private Sub CommandButton2_Click()
Cells(72, 1) = 3
Unload Me
End Sub


