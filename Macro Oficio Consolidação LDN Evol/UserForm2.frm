VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm2 
   Caption         =   "UserForm2"
   ClientHeight    =   3120
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5670
   OleObjectBlob   =   "UserForm2.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
    If NR1 = "" Then NR1 = 0
    If NR2 = "" Then NR2 = 0
    If NR3 = "" Then NR3 = 0
    If LO1 = "" Then LO1 = 0
    If LO2 = "" Then LO2 = 0
    If LO3 = "" Then LO3 = 0
    Sheets("Calculo").Cells(64, 2) = NR1
    Sheets("Calculo").Cells(65, 2) = LO1
    Sheets("Calculo").Cells(64, 4) = NR2
    Sheets("Calculo").Cells(65, 4) = LO2
    Sheets("Calculo").Cells(64, 6) = NR3
    Sheets("Calculo").Cells(65, 6) = LO3
    
    Unload Me

End Sub

Private Sub CommandButton2_Click()
    If NR1 = "" Then NR1 = 0
    If NR2 = "" Then NR2 = 0
    If NR3 = "" Then NR3 = 0
    If LO1 = "" Then LO1 = 0
    If LO2 = "" Then LO2 = 0
    If LO3 = "" Then LO3 = 0
    Sheets("Calculo").Cells(64, 2) = NR1
    Sheets("Calculo").Cells(65, 2) = LO1
    Sheets("Calculo").Cells(64, 4) = NR2
    Sheets("Calculo").Cells(65, 4) = LO2
    Sheets("Calculo").Cells(64, 6) = NR3
    Sheets("Calculo").Cells(65, 6) = LO3
    OK1 = 0
    OK2 = 0
    OK3 = 0
    OK = 79
    Do Until OK = 1000
        If (Cells(OK, 1) = "PMM1" Or Cells(OK, 2) = "PMM1") And Cells(OK, 24) < "70" And Cells(OK, 23) <> "" And Cells(OK, 23) <> "0" Then
            OK1 = OK1 + 1
            ElseIf (Cells(OK, 1) = "PMM1" Or Cells(OK, 2) = "PMM2") And Cells(OK, 24) < "70" And Cells(OK, 23) <> "" And Cells(OK, 23) <> "0" Then
                OK2 = OK2 + 1
                ElseIf (Cells(OK, 1) = "PMM1" Or Cells(OK, 2) = "PMM3") And Cells(OK, 24) < "70" And Cells(OK, 23) <> "" And Cells(OK, 23) <> "0" Then
                    OK3 = OK3 + 1
        End If
        OK = OK + 1
    Loop
    
    CO1 = 0
    CO2 = 0
    CO3 = 0
    CO = 79
    Do Until CO = 1000
        If (Cells(CO, 1) = "PMM1" Or Cells(CO, 2) = "PMM1") And Cells(CO, 24) < "70" And Cells(CO, 23) <> "" And Cells(CO, 23) <> "0" Then
            CO1 = CO1 + 1
            ElseIf (Cells(CO, 1) = "PMM2" Or Cells(CO, 2) = "PMM2") And Cells(CO, 24) < "70" And Cells(CO, 23) <> "" And Cells(CO, 23) <> "0" Then
                CO2 = CO2 + 1
                ElseIf (Cells(CO, 1) = "PMM3" Or Cells(CO, 2) = "PMM3") And Cells(CO, 24) < "70" And Cells(CO, 23) <> "" And Cells(CO, 23) <> "0" Then
                    CO3 = CO3 + 1
        End If
        CO = CO + 1
    Loop
    OK1.Value = OK1
    OK2.Value = OK2
    OK3.Value = OK3
    CO1.Value = CO1
    CO2.Value = CO2
    CO3.Value = CO3
    
End Sub
