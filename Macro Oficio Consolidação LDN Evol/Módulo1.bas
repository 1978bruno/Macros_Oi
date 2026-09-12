Attribute VB_Name = "Módulo1"
Sub auto_open()
Static pula
' Abrir_arquivo Macro
' Macro gravada em 6/9/2005 por 92033
If pula <> 1 Then
UserForm1.Show
End If
pula = 1

' Seleciona o tipo de consolidação que deseja executar
inicio1:
    'Marc = InputBox("No campo abaixo selecione o tipo de tratamento que deseja que a macro de aos valores:" & Chr(13) & "1 - DDD-X FCN7 LDN" & Chr(13) & "2 - FCN7 LDN (Oficio 745)" & Chr(13) & "3 - Sair", "Operação a ser executada")
    UserForm3.Show
    Marc = Cells(72, 1)
    If Marc = 1 Then
        GoTo inicio3:
            ElseIf Marc = 2 Then
                GoTo inicio2:
                ElseIf Marc = 3 Then
                    GoTo fim:
    End If
    GoTo inirot:
inicio2:
' Seleção da operadora a ser tratada
    'marc1 = InputBox("No campo abaixo selecione a operadora a qual deseja fazer o tratamento:" & Chr(13) & "1 - Brasil Telecom" & Chr(13) & "2 - Sercomtel" & Chr(13) & "3 - Telefonica" & Chr(13) & "4 - CTBC" & Chr(13) & "5 - Sair", "Seleção de Operadora")
    UserForm4.Show
    marc1 = Cells(72, 1)
' Montando a Base
    If marc1 = 1 Then
        LFM = Cells(4, 72).End(xlDown).Row
        Range(Cells(4, 72), Cells(4, 72).End(xlDown).End(xlToRight)).Select
        Selection.Copy
    ElseIf marc1 = 2 Then
        LFM = Cells(4, 84).End(xlDown).Row
        Range(Cells(4, 84), Cells(4, 84).End(xlDown).End(xlToRight)).Select
        Selection.Copy
    ElseIf marc1 = 3 Then
        LFM = Cells(4, 80).End(xlDown).Row
        Range(Cells(4, 80), Cells(4, 80).End(xlDown).End(xlToRight)).Select
        Selection.Copy
    ElseIf marc1 = 4 Then
        IG = MsgBox("No momento, esta operadora não esta sendo contemplada por esta macro. Aguarde, em breve, nova alteração.", vbOK, "Indisponibilidade do sistema")
        GoTo inicio2:
    ElseIf marc1 = 5 Then
        GoTo fim:
    End If
        
        Cells(80, 1).Select
        ActiveSheet.Paste
    GoTo inirot:
    
inicio3:
' Seleção da operadora a ser tratada
    'marc2 = InputBox("No campo abaixo selecione a Região a qual deseja fazer o tratamento:" & Chr(13) & "1 - Região II" & Chr(13) & "2 - Região III" & Chr(13) & "3 - Sair", "Seleção de Operadora")
    UserForm5.Show
    marc2 = Cells(72, 1)

' Montando a Base
    If marc2 = 1 Then
        LFM = Cells(4, 89).End(xlDown).Row
        Range(Cells(4, 89), Cells(4, 89).End(xlDown).End(xlToRight)).Select
        Selection.Copy
    ElseIf marc2 = 2 Then
        LFM = Cells(4, 92).End(xlDown).Row
        Range(Cells(4, 92), Cells(4, 92).End(xlDown).End(xlToRight)).Select
        Selection.Copy
    ElseIf marc2 = 3 Then
        GoTo fim:
    End If
        
        Cells(80, 1).Select
        ActiveSheet.Paste
    
    
inirot:
    
    
    
    
    If Marc = 1 Then
    col = 2
    Else: col = 1
    End If
    
    LIN = 80
    Do Until LIN = LFM + 80 - 3
        ' totalização do PMM
        If Cells(LIN, col) = "PMM1" And Cells(LIN, col + 1) <> "[TOTAL]" And Cells(LIN, col + 2) <> "[TOTAL]" Then
            Cells(63, 9).Select
            Range(Selection, Selection.End(xlToRight)).Select
            Selection.Copy
            Cells(LIN, 5).Select
            ActiveSheet.Paste
            LIN = LIN + 1
        ElseIf Cells(LIN, col) = "PMM2" And Cells(LIN, col + 1) <> "[TOTAL]" And Cells(LIN, col + 2) <> "[TOTAL]" Then
            Cells(64, 9).Select
            Range(Selection, Selection.End(xlToRight)).Select
            Selection.Copy
            Cells(LIN, 5).Select
            ActiveSheet.Paste
            LIN = LIN + 1
        ElseIf Cells(LIN, col) = "PMM3" And Cells(LIN, col + 1) <> "[TOTAL]" And Cells(LIN, col + 2) <> "[TOTAL]" Then
            Cells(65, 9).Select
            Range(Selection, Selection.End(xlToRight)).Select
            Selection.Copy
            Cells(LIN, 5).Select
            ActiveSheet.Paste
            LIN = LIN + 1
        ' totalização somatoria nivel1
        ElseIf Cells(LIN, col + 1) <> "[TOTAL]" And Cells(LIN, col + 2) = "[TOTAL]" Then
            Cells(66, 9).Select
            Range(Selection, Selection.End(xlToRight)).Select
            Selection.Copy
            Cells(LIN, 5).Select
            ActiveSheet.Paste
            LIN = LIN + 1
        ' totalização somatoria nivel2
        ElseIf Cells(LIN, col + 1) = "[TOTAL]" And Cells(LIN, col + 2) <> "[TOTAL]" Then
            Cells(67, 9).Select
            Range(Selection, Selection.End(xlToRight)).Select
            Selection.Copy
            Cells(LIN, 5).Select
            ActiveSheet.Paste
            LIN = LIN + 1
        ' totalização somatoria nivel3
        ElseIf Cells(LIN, col + 1) = "[TOTAL]" And Cells(LIN, col + 2) = "[TOTAL]" And Cells(LIN, col + 2) = "[TOTAL]" Then
            Cells(68, 9).Select
            Range(Selection, Selection.End(xlToRight)).Select
            Selection.Copy
            Cells(LIN, 5).Select
            ActiveSheet.Paste
            LIN = LIN + 1
        End If
    Loop

    Cells(1, 1).Select

' Liberação de ocultamento
    Columns("BR:CG").Select
    Selection.EntireColumn.Hidden = False
    ActiveWindow.LargeScroll ToRight:=-3
    ActiveWindow.SmallScroll Down:=22
    Rows("59:3850").Select
    Selection.EntireRow.Hidden = False
    Range("F57").Select
    Cells(1, 1).Select
' Disponibilizar a Planilha Base
    Sheets("Base").Visible = True
    
    MACRO = ActiveWorkbook.Name
' Copiar o CDR
' Escolher o arquivo a ser aberto , do tipo  *.xls
    SGD = Application.GetOpenFilename("Arquivos do Microsoft Excel,*.XLS", , "SGD")

' Abrir o arquivo escolhido
    Workbooks.Open Filename:=SGD
    PLANCDR = ActiveWorkbook.Name
    
' Selecionar a ultima linha
    LFM2 = Cells(13, 1).End(xlDown).Row
    
' Seleciona a Coluna "A"
    Columns("A:A").Select
' Insere uma coluna "A" em branco.
    Selection.Insert Shift:=xlToRight

' Busca da linha Cinza para referencia
    LIN = 1
    Do Until Cells(LIN, 1).Interior.ColorIndex = 15
    LIN = LIN + 1
    Loop
    
' Concatenação dos valores (Juntar DDD, CSP, IND e PMM)
    Do Until LIN = LFM2
    LIN = LIN + 1
    DDD = Mid(Cells(LIN, 3), 1, 2)
    CSP = Cells(LIN, 5)
    IND = Cells(LIN, 12)
    PMM = Mid(Cells(LIN, 8), 12, 2)
    Cells(LIN, 1) = DDD & CSP & IND & "'" & PMM
    Loop
    
' Transpor Planilha
    Cells.Select
    Selection.Copy
    
    Windows(MACRO).Activate
    Sheets("Base").Select
    Cells(1, 1).Select
    ActiveSheet.Paste
    
    Application.DisplayAlerts = False
    Windows(PLANCDR).Close
    
    Application.DisplayAlerts = True
   
    
    Sheets("Calculo").Activate
    Cells(1, 1).Select
    


    
' Fazer Loop
    REP = 1
    Do Until REP = 7
        
'Indicar qual a porcentagem de NR deve ser excluida
    Sheets("Calculo").Select
    'NR = InputBox("O valor atual é: " & Cells(64, 2) & "%, para modificá-lo , digite o valor desejado ", "Exclusão de NR")
    'Cells(64, 2) = NR
    'REP = MsgBox("Para " & Cells(64, 2) & "% de exclusão do NR, existem:" & Chr(13) & OKFORA & " indicadores de OK fora da meta" & Chr(13) & COFORA & " indicadores de CO fora da meta" & Chr(13) & "Gostaria de fazer uma nova tentativa?", vbYesNo, "Situação do Exercicio")
    UserForm2.Show
    
' Exportar planilha pronta
' Reexibir Planilha SGD
    Sheets("SGD").Visible = True
    Sheets("SGD").Select
    
    If Marc = 1 Then
        Cells(5, 1) = "Origem.: BRASIL TELECOM"
    ElseIf Marc = 2 Then
        Cells(5, 1) = "Origem.: SERCOMTEL"
    ElseIf Marc = 3 Then
        Cells(5, 1) = "Origem.: TELEFONICA"
    ElseIf Marc = 4 Then
        Cells(5, 1) = "Origem.: CTBC"
    End If
    
    Range("A8:K8").Select
    Selection.Copy
    LIN = 8
    
    Do Until LIN = LFM + 7 - 2
        LIN = LIN + 1
        Cells(LIN, 1).Select
        ActiveSheet.Paste
    Loop
    
    Cells(9, 1).Select
    Range(Selection, Selection.End(xlToRight).End(xlDown)).Select
    Selection.Interior.ColorIndex = 2

    
    Sheets("SGD").Select
    Sheets("SGD").Copy
    Cells.Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:= _
        False, Transpose:=False
    Consolidacao = Application.GetSaveAsFilename(SGD, "Arquivos do Microsoft Excel,*.XLS", , "Salvar como:")
    ActiveWorkbook.SaveAs Filename:= _
    Consolidacao
    
    ActiveWorkbook.Close
        

    Loop

    Windows(MACRO).Activate
    Cells(10, 1).Select
    Range(Selection, Selection.End(xlToRight).End(xlDown)).Select
    Selection.Clear
    Sheets("SGD").Visible = False
    
    Cells(69, 1).Select
    Range(Selection, Selection.End(xlToRight).End(xlDown)).Select
    Selection.Clear

    Sheets("Base").Activate
    Cells.Select
    Selection.Clear
    Sheets("Base").Visible = False

    Rows("61:3848").Select
    Selection.EntireRow.Hidden = True
    ActiveWindow.SmallScroll Down:=-9
    ActiveWindow.ScrollRow = 1
    ActiveWindow.LargeScroll ToRight:=3
    Columns("BS:CF").Select
    Selection.EntireColumn.Hidden = True
    
    Cells(1, 1).Select
fim:
End Sub




