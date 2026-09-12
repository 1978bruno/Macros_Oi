Attribute VB_Name = "Módulo1"
Sub Import()

macro = ActiveWorkbook.Name       ' Planilha principal, base da Macro

Sheets("Grafico").Select

' Abre o Arquivo a ser tratado
RelCOI = Application.GetOpenFilename("Macro - Internacional.xls,*.xls,Todos os arquivos,*.*", , "Relatório de Localidades Ofensoras")


UserForm2.Show 0
Application.Wait (Now + TimeValue("0:00:01"))

Application.DisplayAlerts = False
Application.ScreenUpdating = False
Sheets("ASR").Visible = True
Sheets("Base").Select
If ActiveSheet.AutoFilterMode = True Then Selection.AutoFilter
Range("6:65000").Delete

Workbooks.Open Filename:=RelCOI
RELCOO = ActiveWorkbook.Name

Sheets("SAINTE").Select
Cells(2, 4).Select
Selection.End(xlToRight).Select
CLFM = ActiveCell.Column

LIN3 = 6
NREF = 0
SMN21 = 0
SMN22 = 0
Do Until NREF = 20
    Workbooks(RELCOO).Activate
    Sheets("FACE").Select
    PAIS = Cells(10 + NREF, 18)
    TELE = Cells(10 + NREF, 19)
    
    Sheets("SAINTE").Select
    LIN = 2
    
    
    Do Until PAIS = Cells(LIN, 1) And TELE = Cells(LIN, 3)
        LIN = LIN + 1
    Loop
    
    LIN2 = LIN
    Do Until PAIS <> Cells(LIN2, 1) Or TELE <> Cells(LIN2, 3)
        LIN2 = LIN2 + 1
    Loop
    
    LIN2 = LIN2 - 1
    DIA_ANTE = 0
    BARRA = 8 'Para definir a barra que sera colorida
    
    Do Until DIA_ANTE = 24
        Workbooks(RELCOO).Activate
        Sheets("SAINTE").Select
        
        Range(Cells(LIN, 1), Cells(LIN2, 3)).Copy
        Workbooks(macro).Activate
        Sheets("Base").Select
        Cells(LIN3, 2).Select
        ActiveSheet.Paste
        
        Workbooks(RELCOO).Activate
        Sheets("SAINTE").Select
        
        DT = Cells(1, CLFM - 2 - DIA_ANTE)
        Range(Cells(LIN, CLFM - DIA_ANTE), Cells(LIN2, CLFM - 2 - DIA_ANTE)).Copy
        Workbooks(macro).Activate
        Sheets("Base").Select
        Cells(LIN3, 5).Select
        ActiveSheet.Paste
        LDT = LIN3
        
        'Implementar a tabela de ASR
        BUSCA = PAIS & TELE
        Sheets("ASR").Select
        Range("A:A").Select
        Selection.Find(What:=BUSCA).Activate
        LINR = ActiveCell.Row
        ASR = Cells(LINR, 4)
        Sheets("Base").Select
        
        Do Until LDT = LIN3 + (LIN2 - LIN) + 1
            Cells(LDT, 1) = DT
            Cells(LDT, 8) = ASR
            LDT = LDT + 1
            
        Loop
        
        ' Definição de Final de semana
        'Stop
        SMN = Weekday(DT, vbMonday)
        If SMN = 6 And SMN21 = 0 Then
            BL1 = BARRA
            SMN21 = SMN21 + 1
            ElseIf SMN = 6 And SMN21 = 1 Then
                BL3 = BARRA
                SMN21 = SMN21 + 1
                SMN22 = SMN22 + 1
        End If
        If SMN = 7 And SMN22 = 0 Then
            BL2 = BARRA
            SMN22 = SMN22 + 1
            ElseIf SMN = 7 And SMN22 = 1 Then
                BL3 = BARRA
                SMN22 = SMN22 + 1
                SMN21 = SMN21 + 1
        End If
        BARRA = BARRA - 1
        ' Fim da rotina de definição do final de semana
                
        DIA_ANTE = DIA_ANTE + 3
        LIN3 = LIN3 + (LIN2 - LIN) + 1
    
    Loop
    NREF = NREF + 1
Loop

'Classificar
Cells(6, 1).Select
Selection.Sort Key1:=Range("A6"), Order1:=xlAscending, Key2:=Range("B6") _
    , Order2:=xlAscending, Key3:=Range("C6"), Order3:=xlAscending, Header:= _
    xlGuess, OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom

'Ajustar Planilha
For LINHA = 6 To LDT - 1
    On Error Resume Next
    t = Cells(LINHA, 5) * 1
    If Err.Number = 13 Then
        Err.Clear
        Else
            Cells(LINHA, 5) = t
    End If
Next


Range("G:G").Select
Selection.Replace What:="%", Replacement:="", LookAt:=xlPart, _
    SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
    ReplaceFormat:=False

For LINHA = 6 To LDT - 1
    On Error Resume Next
    t = Left(Cells(LINHA, 7), 5) * 1
    If Err.Number = 13 Then
        Err.Clear
        Else
            Cells(LINHA, 7) = t
    End If
Next


Cells(6, 1).Select
' Excluir carrier sem trafego
Selection.Sort Key1:=Range("D6"), Order1:=xlAscending, Key2:=Range("C6") _
    , Order2:=xlAscending, Key3:=Range("B6"), Order3:=xlAscending, Header:= _
    xlGuess, OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom

LIN = 6
Do Until Cells(LIN, 1) = ""
    cont = 0
    LIN2 = LIN
    Do Until Cells(LIN, 2) <> Cells(LIN + 1, 2) Or Cells(LIN, 3) <> Cells(LIN + 1, 3) Or Cells(LIN, 4) <> Cells(LIN + 1, 4)
        If Cells(LIN, 6) = "" Then cont = cont + 1
        LIN = LIN + 1
    Loop
    If cont = 7 Then
        Range(Rows(LIN2), Rows(LIN)).Delete
        LIN = LIN2
        Else
            LIN = LIN + 1
    End If
    Cells(LIN, 1).Select
Loop

'Classificar
Cells(6, 1).Select
Selection.Sort Key1:=Range("A6"), Order1:=xlAscending, Key2:=Range("B6") _
    , Order2:=xlAscending, Key3:=Range("C6"), Order3:=xlAscending, Header:= _
    xlGuess, OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom

'Cells(6, 1).End(xlToRight).Select

'LDT = Selection.Row

LDT = LIN

Workbooks(RELCOO).Activate
Sheets("FACE").Select
Range("Q7:T30").Copy
Workbooks(macro).Activate
Sheets("Grafico").Select
ActiveWindow.Visible = False
Windows(macro).Activate
Cells(1, 1).Select
ActiveSheet.Paste

Workbooks(RELCOO).Activate
Sheets("FACE").Select
Range("B7:G30").Copy
Workbooks(macro).Activate
Sheets("Grafico").Select
ActiveWindow.Visible = False
Windows(macro).Activate
Cells(36, 1).Select
ActiveSheet.Paste

Workbooks(RELCOO).Close

'Preparação para o 1º nivel do Drop Down
'Sheets("Base").Select                          '| Buscar da tabela top 20
'Cells(6, 2).Select                             '| Com isso irei conseguir um ganho de
'Range(Selection, Selection.End(xlDown)).Copy   '| performance.

Sheets("Grafico").Select
Range(Cells(4, 2), Cells(23, 2)).Copy
Cells(2, 22).Select
ActiveSheet.Paste

Selection.Sort Key1:=Range("V2"), Order1:=xlAscending, Header:=xlGuess, _
    OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom

Cells(1, 1).Select

LIN = 2
Do Until Cells(LIN, 22) = ""
    If Cells(LIN, 22) = Cells(LIN + 1, 22) Then
        Cells(LIN + 1, 22).Delete
        GoTo FMLP
    End If
    LIN = LIN + 1
FMLP:
Cells(LIN, 22).Select
Loop

Cells(2, 22).Select
Selection.End(xlDown).Select
ULTLN = ActiveCell.Row

ActiveSheet.ChartObjects("Gráfico 1").Activate
ActiveChart.Shapes("Drop Down 5").Select
With Selection
    .ListFillRange = "Grafico!$V$2:$V$" & ULTLN
    .LinkedCell = "Grafico!$V$1"
    .DropDownLines = 8
    .Display3DShading = False
End With

ActiveChart.Shapes("Drop Down 6").Select
With Selection
    .ListFillRange = "Grafico!$X$2:$X$10"
    .LinkedCell = "Grafico!$X$1"
    .DropDownLines = 8
    .Display3DShading = False
End With

ActiveChart.Shapes("Drop Down 7").Select
With Selection
    .ListFillRange = "Grafico!$Z$2:$Z$10"
    .LinkedCell = "Grafico!$Z$1"
    .DropDownLines = 8
    .Display3DShading = False
End With

ActiveChart.ChartArea.Select
ActiveChart.SeriesCollection(1).XValues = "=Base!R6C1:R" & LDT - 1 & "C1"
ActiveChart.SeriesCollection(1).Name = "=Base!R5C6"
ActiveChart.SeriesCollection(2).Name = "=Base!R5C5"
ActiveChart.SeriesCollection(3).Name = "=Base!R5C7"
ActiveChart.SeriesCollection(4).Name = "=Base!R5C8"
ActiveChart.SeriesCollection(1).Values = "=Base!R6C6:R" & LDT - 1 & "C6"
ActiveChart.SeriesCollection(2).Values = "=Base!R6C5:R" & LDT - 1 & "C5"
ActiveChart.SeriesCollection(3).Values = "=Base!R6C7:R" & LDT - 1 & "C7"
ActiveChart.SeriesCollection(4).Values = "=Base!R6C8:R" & LDT - 1 & "C8"

'Ajuste de cores do grafico
'Preparação Base
'Sheets("Base").Select
'Cells(6, 1).Select
'Range(Selection, Selection.End(xlDown)).Copy
'
'Sheets("Grafico").Select
'ActiveWindow.Visible = False
'Windows(macro).Activate
'Cells(2, 28).Select
'ActiveSheet.Paste
'
'Selection.Sort Key1:=Range("AB2"), Order1:=xlAscending, Header:=xlGuess, _
'    OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom
'
'Cells(1, 1).Select
'
'SMN21 = 0
'SMN22 = 0
'LIN = 2
'Do Until Cells(LIN, 28) = ""
'    If Cells(LIN, 28) = Cells(LIN + 1, 28) Then
'        Cells(LIN + 1, 28).Delete
'        GoTo FMLP2
'    End If
'    SMN = Weekday(Cells(LIN, 28), vbMonday)
'    Cells(LIN, 29) = SMN
'    If SMN = 6 And SMN21 = 0 Then
'        BL1 = LIN - 1
'        SMN21 = 1
'        ElseIf SMN = 6 And SMN21 = 1 Then
'            BL3 = LIN - 1
'            SMN21 = 0
'    End If
'    If SMN = 7 And SMN22 = 0 Then
'        BL2 = LIN - 1
'        SMN22 = 1
'        ElseIf SMN = 7 And SMN22 = 1 Then
'            BL3 = LIN - 1
'            SMN22 = 0
'    End If
'    LIN = LIN + 1
'FMLP2:
'Cells(LIN, 28).Select
'Loop

'Range(Cells(2, 28), Selection.End(xlDown)).Delete
'Range(Cells(2, 28), Selection.End(xlDown)).Delete

ActiveSheet.ChartObjects("Gráfico 1").Activate
ActiveChart.SeriesCollection(1).Select
With Selection.Border
    .Weight = xlThin
    .LineStyle = xlAutomatic
End With
Selection.Shadow = False
Selection.InvertIfNegative = False
With Selection.Interior
    .ColorIndex = 46
    .Pattern = xlSolid
End With
ActiveChart.SeriesCollection(1).Points(BL1).Select
With Selection.Border
    .Weight = xlThin
    .LineStyle = xlAutomatic
End With
Selection.Shadow = False
Selection.InvertIfNegative = False
With Selection.Interior
    .ColorIndex = 44
    .Pattern = xlSolid
End With
ActiveChart.SeriesCollection(1).Points(BL2).Select
With Selection.Border
    .Weight = xlThin
    .LineStyle = xlAutomatic
End With
Selection.Shadow = False
Selection.InvertIfNegative = False
With Selection.Interior
    .ColorIndex = 44
    .Pattern = xlSolid
End With
ActiveChart.SeriesCollection(1).Points(BL3).Select
With Selection.Border
    .Weight = xlThin
    .LineStyle = xlAutomatic
End With
Selection.Shadow = False
Selection.InvertIfNegative = False
With Selection.Interior
    .ColorIndex = 44
    .Pattern = xlSolid
End With

ActiveWindow.Visible = False
Windows(macro).Activate

Cells(1, 22) = "1"
Cells(1, 24) = "1"
Cells(1, 26) = "1"

Cells(1, 1).Select
Sheets("ASR").Visible = False

Application.DisplayAlerts = True
Application.ScreenUpdating = True

Unload UserForm2

UserForm1.Show 0
Application.Wait (Now + TimeValue("0:00:03"))

Unload UserForm1

End Sub
