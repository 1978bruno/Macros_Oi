VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm3 
   Caption         =   "Data inicial"
   ClientHeight    =   3915
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   2730
   OleObjectBlob   =   "UserForm3.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub UserForm_Initialize()

Label2.Caption = "Intervalo: " & DT1 & " à " & DT2
NREF = DateDiff("d", DTI, DTF)
If NREF < 0 Then
    Label10.Caption = "Data fim inferior à Data inicio."
End If

End Sub


Private Sub CommandButton1_Click()

DTI = TextBox1.Value & "/" & TextBox2.Value & "/" & TextBox3.Value
DTF = TextBox4.Value & "/" & TextBox5.Value & "/" & TextBox6.Value

Unload Me

End Sub
