Attribute VB_Name = "Módulo2"
Public PLANREF

Sub sainte()

PLANREF = "SAINTE"
internacional

End Sub

Sub entrante()

PLANREF = "ENTRANTE"
internacional

End Sub

Sub pan()

PLANREF = "ROOMING"
internacional

End Sub

Sub internacional()

macro = ActiveWorkbook.Name       ' Planilha principal, base da Macro
Application.DisplayAlerts = False
Ref = 0

'Mantem a tela congelada
Cells(35, 8) = "Aguarde.... Executando...."

' Abre o Arquivo a ser tratado
RelCOI = Application.GetOpenFilename("Planilha Excel,*.XLS,Todos os arquivos,*.*", , "Relatório de Localidades Ofensoras")
Workbooks.Open Filename:=RelCOI

RelCOO = ActiveWorkbook.Name

CONSULT = ActiveSheet.Name
Sheets(CONSULT).Copy After:=Workbooks(macro).Sheets(1)

Workbooks(RelCOO).Close

REFCOMP = ActiveSheet.Name
LIN = 1
'Achar a linha de cabeçalho
Do Until Cells(LIN, 1).Interior.ColorIndex = 48
    If Left(Cells(LIN, 1), 4) = "Data" Then DT = Right(RTrim(Cells(LIN, 1)), 10)
    LIN = LIN + 1
Loop

Sheets(PLANREF).Select
If ActiveSheet.AutoFilterMode = True Then Selection.AutoFilter

Cells(2, 1).Select
Selection.End(xlToRight).Select

If ActiveCell = "Telefonia" Then
    CLFM = ActiveCell.Column + 1
    FLAG = 0
    Else
        Cells(2, 4).Select
        Selection.End(xlToRight).Select
        CLFM = ActiveCell.Column + 1
        For CLED = 4 To CLFM
            If Trim(Cells(1, CLED)) = DT Then
                Cells(1, CLED).Select
                CLFM = ActiveCell.Column
            End If
        Next
        
        
        Cells(2, 4).Select
        Selection.End(xlToRight).Select
        CLFM1 = ActiveCell.Column - 2
        
        Data1 = Cells(1, CLFM1)
        If (Day(Data1) + 1 <> Day(DT) And Month(Data1) = Month(DT)) Or (Day(DT) = "01" Or Day(DT) = "1" And Month(Data1) + 1 = Month(DT)) Then
            erro = MsgBox("Houve um erro ao tentar tratar o arquivo. Deveria estar sendo tratado o dia " & Day(Data1) + 1 & " e não o dia " & Day(DT) & ". Ok para continuar assim mesmo. Cancel para tratar o arquivo correto.", vbOKCancel, "Erro de Execução")
            If erro = 2 Then GoTo Fim2
        End If
        FLAG = 0
End If

Pula:

'Preparação da planilha
Sheets(REFCOMP).Select
If ActiveSheet.AutoFilterMode = True Then Selection.AutoFilter

'Localização de coluna
Do Until COL = 6
Select Case COL
    Case 1
        If PLANREF = "ENTRANTE" Then
            BUSCAVAL = "Localidade"
            Else
                BUSCAVAL = "País"
        End If
    Case 2
        If PLANREF = "ENTRANTE" Then
            BUSCAVAL = "Operadora"
            Else
                BUSCAVAL = "Rota de saída"
        End If
    Case 3
        If PLANREF = "ENTRANTE" Then
            BUSCAVAL = "PRD"
            Else
                BUSCAVAL = "Telefonia"
        End If
    Case 4
        BUSCAVAL = "OK%"
    Case 5
        BUSCAVAL = "TCH"
End Select


If BUSCAVAL <> "" Then
Rows(LIN).Select
Selection.Find(What:=BUSCAVAL, After:=ActiveCell, LookIn:=xlFormulas, _
    LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
    MatchCase:=False).Activate
End If
Select Case COL
    Case 1
        COL_1 = ActiveCell.Column
    Case 2
        COL_2 = ActiveCell.Column
    Case 3
        COL_3 = ActiveCell.Column
    Case 4
        COL_4 = ActiveCell.Column
    Case 5
        COL_5 = ActiveCell.Column
End Select
COL = COL + 1
Loop



If FLAG <> 1 Then
    Sheets(PLANREF).Select

    'Fortmatar linha
    Range(Cells(1, CLFM), Cells(1, CLFM + 2)).Select
    With Selection
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = True
    End With
    With Selection.Font
        .Name = "Arial"
        .FontStyle = "Negrito"
        .Size = 10
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ColorIndex = xlAutomatic
    End With
    With Selection.Interior
        .ColorIndex = 48
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
    End With
    
    Cells(1, CLFM) = Format(DT, "MM/DD/YYYY")
    
    Range(Cells(2, CLFM - 3), Cells(2, CLFM - 1)).Select
    Selection.Copy
    Cells(2, CLFM).Select
    ActiveSheet.Paste
    
End If
    


LIN = LIN + 1

Sheets(REFCOMP).Select
Do Until Cells(LIN, 1) = ""
    
    Sheets(REFCOMP).Select
    
    PAI = Cells(LIN, COL_1)
    RSA = Cells(LIN, COL_2)
    TEL = Cells(LIN, COL_3)
    OKP = Cells(LIN, COL_4)
    TCH = Cells(LIN, COL_5)
    
    If PAI = "[TOTAL]" Or PAI = "NI-NI-NÃO IDENTIFICADA" Or PAI = "EX-EXT-EXTERIOR" Or RSA = "[TOTAL]" _
        Or RSA = "NI" Or RSA = "ADMI" Or TEL = "[TOTAL]" Then GoTo FIM
    
    Sheets(PLANREF).Select
    
    PARALOOP = 0
    LIN2 = 3

    Do Until PARALOOP = 1
    
        If (PLANREF = "ENTRANTE" And PAI = Cells(LIN2, 1) And RSA = Cells(LIN2, 2)) Or _
            (PLANREF = "SAINTE" And TEL = Cells(LIN2, 3) And PAI = Cells(LIN2, 1) And RSA = Cells(LIN2, 2)) Or _
            (PLANREF = "PAN" And TEL = Cells(LIN2, 3) And PAI = Cells(LIN2, 1) And RSA = Cells(LIN2, 2)) Then
            Cells(LIN2, CLFM) = OKP
            Cells(LIN2, CLFM + 1) = TCH
            If PLANREF = "ENTRANTE" Then Cells(LIN2, CLFM + 2) = TEL
            Cells(LIN2, CLFM).Select 'SO PARA ACOMPANHAR

            PARALOOP = 1
            Sheets(REFCOMP).Select
            ElseIf Cells(LIN2, 1) = "" And Cells(LIN2, 2) = "" And Cells(LIN2, 3) = "" Then
                Cells(LIN2, 1) = PAI
                Cells(LIN2, 2) = RSA
                If PLANREF = "ENTRANTE" Then
                    Cells(LIN2, CLFM + 2) = TEL
                    Else
                        Cells(LIN2, 3) = TEL
                End If
                Cells(LIN2, CLFM) = OKP
                Cells(LIN2, CLFM + 1) = TCH
                Cells(LIN2 + 1, CLFM).Select 'SO PARA ACOMPANHAR
                PARALOOP = 1
                Sheets(REFCOMP).Select
        End If
        LIN2 = LIN2 + 1
    Loop
    
FIM:
    
    LIN = LIN + 1

Loop

UTM = LIN2

Sheets(PLANREF).Select
Range("3:65000").Select
Selection.Sort Key1:=Range("A3"), Order1:=xlAscending, Key2:=Range("C3") _
    , Order2:=xlAscending, Key3:=Range("B3"), Order3:=xlAscending, Header:= _
    xlGuess, OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom

LFIM = 3
If PLANREF = "SAINTE" Or PLANREF = "PAN" Then
    Do Until LFIM > 65000
        LFIN = LFIM
        FLG0 = 0
        Do Until FLG0 = 1
            TTCH = TTCH + Cells(LFIM, CLFM + 1)
            If Cells(LFIM, 1) <> Cells(LFIM + 1, 1) Or Cells(LFIM, 3) <> Cells(LFIM + 1, 3) Or Cells(LFIM, 1) = "" Then FLG0 = 1
            LFIM = LFIM + 1
        Loop
        
        Do Until LFIN = LFIM
            If Cells(LFIN, CLFM + 1) <> "" Then
                Cells(LFIN, CLFM + 2) = Format(Cells(LFIN, CLFM + 1) / TTCH, "0.00%")
                LFIN = LFIN + 1
                Else
                    LFIN = LFIN + 1
            End If
        Loop
        TTCH = 0
    Loop
End If
    
Cells(1, 1).Select
Fim2:
Sheets(REFCOMP).Delete
Sheets("Face").Select
Cells(35, 8) = ""

End Sub
