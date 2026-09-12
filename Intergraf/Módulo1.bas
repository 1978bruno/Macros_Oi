Attribute VB_Name = "Módulo1"
Private Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" _
(ByVal lpBuffer As String, nSize As Long) As Long

Public DTI As Date
Public DTF As Date
Public DTT As Date
Public DT1 As Date
Public DT2 As Date
Public PLANREF As String

Sub Import()

Dim usuario As String * 255
GetUserName usuario, Len(usuario)
USU_Name = Left(usuario, InStr(usuario, Chr(0)) - 1)

If USU_Name <> "92033" And USU_Name <> "84078" And USU_Name <> "josearf" And USU_Name <> "T87529" And USU_Name <> "0928" Then
    Response = MsgBox("Este usuário não tem autorização para executar esta função.", vbYes, "AUTORIZAÇÃO")
    GoTo ENDOFPLAN
End If

Dim NREFF As Integer

macro = ActiveWorkbook.Name       ' Planilha principal, base da Macro

Sheets("Grafico").Select

' Abre o Arquivo a ser tratado
RelCOI = Application.GetOpenFilename("Planilha Excell,*.xls,Todos os arquivos,*.*", , "Relatório de Localidades Ofensoras")


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

Sheets("FACE").Select
NREFF = Mid(Cells(7, 13), 4, 2)

UserForm4.Show

Sheets(PLANREF).Select
Cells(2, 4).Select
Selection.End(xlToRight).Select
CLFM = ActiveCell.Column

DT1 = Cells(1, 4)
DT2 = Cells(1, CLFM - 3)

pula:
UserForm3.Show
NREF = DateDiff("d", DTI, DTF)
NREF1 = DateDiff("d", DT1, DTI)
NREF2 = DateDiff("d", DTF, DT2)

If NREF < 0 Or NREF1 < 0 Or NREF2 < 0 Then
    GoTo pula
End If

CLF = 4
DT1 = Cells(1, CLF)
Do Until DTI = DT1
    CLF = CLF + 1
    DT1 = Cells(1, CLF)
Loop
DDFFT = DateDiff("d", DTI, DTF)
DDFF = Format((DDFFT / 7) + 1, "0") * 2

Dim varData() As Integer ' Definindo como matriz
ReDim varData(DDFF) ' REDefinindo como matriz
DDFF1 = 1 'Definindo inicio da matriz

LIN3 = 6
NREF = 0
SMN21 = 0
SMN22 = 0

Select Case PLANREF
    Case "SAINTE"
        COLBUSDA = 13
    Case "ROAMING"
        COLBUSDA = 2
End Select

Do Until NREF = NREFF
    Workbooks(RELCOO).Activate
    Sheets("FACE").Select
    PAIS = Cells(10 + NREF, COLBUSDA + 1)
    TELE = Cells(10 + NREF, COLBUSDA + 2)
    
    Sheets(PLANREF).Select
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
    BARRA = 1 'Para definir a barra que sera colorida
    DT = Cells(1, CLF)
    'Importar os Dados da Base_InterGraf.XLS
    Do Until DT = DTF
        Workbooks(RELCOO).Activate
        Sheets(PLANREF).Select
        
        Range(Cells(LIN, 1), Cells(LIN2, 3)).Copy
        Workbooks(macro).Activate
        Sheets("Base").Select
        Cells(LIN3, 2).Select
        ActiveSheet.Paste
        
        Workbooks(RELCOO).Activate
        Sheets(PLANREF).Select
        
        DT = Cells(1, CLF + DIA_ANTE)
        Range(Cells(LIN, CLF + DIA_ANTE), Cells(LIN2, CLF + 3 + DIA_ANTE)).Copy
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
            Cells(LDT, 9) = ASR
            LDT = LDT + 1
        Loop

        ' Definição de Final de semana
        SMN = Weekday(DT, vbMonday)
        If (SMN = 6 Or SMN = 7) And DDFF1 < DDFF + 1 Then
            varData(DDFF1) = BARRA
            DDFF1 = DDFF1 + 1
        End If
        
        BARRA = BARRA + 1
        ' Fim da rotina de definição do final de semana
        
                
        DIA_ANTE = DIA_ANTE + 4
        LIN3 = LIN3 + (LIN2 - LIN) + 1
    
    Loop
    NREF = NREF + 1
    DDFF1 = DDFF + 3
Loop

'Classificar
Cells(6, 1).Select
Selection.Sort Key1:=Range("A6"), Order1:=xlAscending, Key2:=Range("B6") _
    , Order2:=xlAscending, Key3:=Range("C6"), Order3:=xlAscending, Header:= _
    xlGuess, OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom

'Ajustar Planilha
For LINHA = 6 To LDT - 1
    On Error Resume Next
    T = Cells(LINHA, 5) * 1
    If Err.Number = 13 Then
        Err.Clear
        Else
            Cells(LINHA, 5) = T
    End If
Next


Range("G:G").Select
Selection.Replace What:="%", Replacement:="", LookAt:=xlPart, _
    SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
    ReplaceFormat:=False

For LINHA = 6 To LDT - 1
    On Error Resume Next
    T = Left(Cells(LINHA, 7), 5) * 1
    If Err.Number = 13 Then
        Err.Clear
        Else
            Cells(LINHA, 7) = T
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
    Do Until Cells(LIN, 2) <> Cells(LIN + 1, 2) Or Cells(LIN, 3) <> Cells(LIN + 1, 3) Or Cells(LIN, 4) <> Cells(LIN + 1, 4) Or Cells(LIN, 1) = ""
        If Cells(LIN, 6) = "" Then cont = cont + 1
        LIN = LIN + 1
    Loop
    If Cells(LIN, 6) = "" Then cont = cont + 1
    If cont = DDFFT + 1 Then
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


LDT = LIN
' Trazer Tabela de TOP
Workbooks(RELCOO).Activate
Sheets("FACE").Select
Range(Cells(7, COLBUSDA), Cells(NREFF + 9, COLBUSDA + 4)).Copy
Workbooks(macro).Activate
Sheets("Grafico").Select
ActiveWindow.Visible = False
Windows(macro).Activate
Cells(1, 1).Select
ActiveSheet.Paste

Range(Cells(NREFF + 4, 1), Cells(100, 5)).Delete Shift:=xlUp
Range("I29:I33").Clear

Workbooks(RELCOO).Close

Sheets("Grafico").Select
Range(Cells(4, 2), Cells(NREFF + 3, 2)).Copy
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
ActiveChart.SeriesCollection(5).Name = "=Base!R5C9"
ActiveChart.SeriesCollection(1).Values = "=Base!R6C6:R" & LDT - 1 & "C6"
ActiveChart.SeriesCollection(2).Values = "=Base!R6C5:R" & LDT - 1 & "C5"
ActiveChart.SeriesCollection(3).Values = "=Base!R6C7:R" & LDT - 1 & "C7"
ActiveChart.SeriesCollection(4).Values = "=Base!R6C8:R" & LDT - 1 & "C8"
ActiveChart.SeriesCollection(5).Values = "=Base!R6C9:R" & LDT - 1 & "C9"


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
DDFF1 = 1
Do Until DDFF1 = DDFF + 1
    ActiveChart.SeriesCollection(1).Points(varData(DDFF1)).Select
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
    DDFF1 = DDFF1 + 1
Loop

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

ENDOFPLAN:

End Sub
