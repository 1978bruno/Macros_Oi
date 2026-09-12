Attribute VB_Name = "Módulo1"
Sub RegiaoII()
'
' Abrir_arquivo Macro
' Macro gravada em 6/9/2005 por 92033

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
    LFM = Cells(13, 1).End(xlDown).Row
    
' Seleciona a Coluna "A"
    Columns("A:A").Select
' Insere uma coluna "A" em branco.
    Selection.Insert Shift:=xlToRight

' Busca da linha Cinza para referencia
    lin = 1
    Do Until Cells(lin, 2).Interior.ColorIndex = 48
    lin = lin + 1
    Loop
    
' Concatenação dos valores (Juntar PRE, CSP, IND e PMM)
    Do Until lin = LFM
    lin = lin + 1
    PRE = Mid(Cells(lin, 3), 1, 6)
    CSP = Cells(lin, 4)
    IND = Cells(lin, 12)
    PMM = Mid(Right(Cells(lin, 8), 8), 1, 2)
    Cells(lin, 1) = PRE & CSP & IND & "'" & PMM
    
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
   
    
' Ocultar a planilha Base
    Sheets("Base").Visible = False

' Fazer Loop
    REP = 1
    Do Until REP = 2

'Indicar qual a porcentagem de NR deve ser excluida
    NR = InputBox("O valor atual é: " & Cells(65, 2) & "%, para modificá-lo , digite o valor desejado ", "Exclusão de NR")
    Cells(65, 2) = NR
    
' Exportar planilha pronta
' Reexibir Planilha SGD
    Sheets("SGD").Visible = True
    
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
        
    Sheets("Calculo").Activate
    REP = InputBox("Deseja fazer uma nova tentativa?                                 1-SIM  2-NÃO " & XYZ & "", "Novo exercicio", XYZ)

    Loop

    Windows(MACRO).Activate
    Sheets("SGD").Visible = False

End Sub




