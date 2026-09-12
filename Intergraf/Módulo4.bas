Attribute VB_Name = "Módulo4"
Sub FillArrayMulti()

Dim MyVar
MyVar = "Venha me ver no painel Imediato."
Debug.Print MyVar


    Dim intI As Integer, intJ As Integer
    Dim sngMulti(1 To 5, 1 To 10) As Single
    
    ' Preenche a matriz com valores.
    For intI = 1 To 5
        For intJ = 1 To 10
            Teste = intI * intJ
            Debug.Print Teste
        Next intJ
    Next intI
End Sub


