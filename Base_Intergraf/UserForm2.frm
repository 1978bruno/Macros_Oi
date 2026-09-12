VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm2 
   Caption         =   "Seleção de TOP's"
   ClientHeight    =   2400
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3525
   OleObjectBlob   =   "UserForm2.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()

Select Case True
    Case OptionButton1 = True
        TOPREF = 20
    
    Case OptionButton2 = True
        TOPREF = 30
    
    Case OptionButton3 = True
        TOPREF = 40
    
    Case OptionButton4 = True
        TOPREF = 50
    
    Case OptionButton5 = True
        TOPREF = 60
        
    Case OptionButton6 = True
        TOPREF = 70
        
End Select

Unload Me

End Sub

Private Sub OptionButton1_Click()
CommandButton1.SetFocus
End Sub

Private Sub OptionButton2_Click()
CommandButton1.SetFocus
End Sub

Private Sub OptionButton3_Click()
CommandButton1.SetFocus
End Sub

Private Sub OptionButton4_Click()
CommandButton1.SetFocus
End Sub

Private Sub OptionButton5_Click()
CommandButton1.SetFocus
End Sub

Private Sub OptionButton6_Click()
CommandButton1.SetFocus
End Sub
