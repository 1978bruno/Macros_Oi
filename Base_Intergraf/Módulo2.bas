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

PLANREF = "ROAMING"
internacional

End Sub

Sub internacional()

macro = ActiveWorkbook.Name       ' Planilha principal, base da Macro
Application.DisplayAlerts = False
Ref = 0

' Abre o Arquivo a ser tratado (Continua...)
RelCOI = Application.GetOpenFilename("Planilha Excel,*.XLS,Todos os arquivos,*.*", , "Relatório de Localidades Ofensoras")

'Mantem a tela congelada
UserForm3.Show 0
Application.Wait (Now + TimeValue("0:00:01"))

Application.DisplayAlerts = False
Application.ScreenUpdating = False

INICIO:
' (...Continua) Abre o Arquivo a ser tratado
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
        CLFM1 = ActiveCell.Column - 3
        
        Data1 = Cells(1, CLFM1)
        If DateDiff("d", Data1, DT) <> 1 Then
            erro = MsgBox("Houve um erro ao tentar tratar o arquivo. Deveria estar sendo tratado o dia " & Day(Data1) + 1 & " e não o dia " & Day(DT) & " para o " & PLANREF & " . Ok para continuar assim mesmo. Cancel para tratar o arquivo correto.", vbOKCancel, "Erro de Execução")
            If erro = 2 Then GoTo Fim2

            If DateDiff("d", DT, Data1) >= 0 Then
                For COLDT = 4 To CLFM1 + 4

                    If DateDiff("d", Cells(1, COLDT + 4), DT) = 0 Then
                        CLFM = COLDT + 4
                        COLDT = CLFM1 + 4
                        ElseIf DateDiff("d", Cells(1, COLDT), DT) = 1 Then
                            Range(Columns(COLDT + 4), Columns(COLDT + 7)).Insert
                            CLFM = COLDT + 4
                            COLDT = CLFM1 + 4
                    End If
                Next
            End If
            ElseIf DT = Data1 Then CLFM = CLFM - 4
        End If
        FLAG = 0
End If

Pula:


'Eliminar a data mais antiga da base
NREF = DateDiff("d", Cells(1, 4), Cells(1, CLFM - 4))
If NREF > 60 Then
    Range(Columns(4), Columns(7)).Delete
    CLFM = CLFM - 4
End If

'Preparação da planilha
Sheets(REFCOMP).Select
If ActiveSheet.AutoFilterMode = True Then Selection.AutoFilter

'Localização de coluna
Do Until COL = 7
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
    Case 6
        BUSCAVAL = "TTC(Min)"
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
    Case 6
        COL_6 = ActiveCell.Column
End Select
COL = COL + 1
Loop

If FLAG <> 1 Then
    Sheets(PLANREF).Select

    'Fortmatar linha
    Range(Cells(1, CLFM), Cells(1, CLFM + 3)).Select
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
    
    Range(Cells(2, CLFM - 4), Cells(2, CLFM - 1)).Select
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
    TTC = Cells(LIN, COL_6)
    
    
    Select Case PLANREF
        Case "ROAMING"
            If RSA <> "KPNO" Or TEL <> "MOVEL" Then GoTo FIM
        Case "SAINTE"
            If PAI = "[TOTAL]" Or PAI = "NI-NI-NÃO IDENTIFICADA" Or PAI = "EX-EXT-EXTERIOR" Or RSA = _
                "[TOTAL]" Or RSA = "NI" Or RSA = "KPNO" Or RSA = "ADMI" Or TEL = "[TOTAL]" Then GoTo FIM
    End Select
        
    
    Sheets(PLANREF).Select
    
    PARALOOP = 0
    LIN2 = 3

    Do Until PARALOOP = 1
    
        If (PLANREF = "ENTRANTE" And PAI = Cells(LIN2, 1) And RSA = Cells(LIN2, 2)) Or _
            (PLANREF = "SAINTE" And TEL = Cells(LIN2, 3) And PAI = Cells(LIN2, 1) And _
            RSA = Cells(LIN2, 2)) Or (PLANREF = "ROAMING" And TEL = Cells(LIN2, 3) And _
            PAI = Cells(LIN2, 1) And RSA = Cells(LIN2, 2)) Then
            Cells(LIN2, CLFM) = OKP
            Cells(LIN2, CLFM + 1) = TCH
            Cells(LIN2, CLFM + 2) = TTC
            If PLANREF = "ENTRANTE" Then Cells(LIN2, CLFM + 3) = TEL
            PARALOOP = 1
            Sheets(REFCOMP).Select
        
            ElseIf Cells(LIN2, 1) = "" And Cells(LIN2, 2) = "" And Cells(LIN2, 3) = "" Then
                Cells(LIN2, 1) = PAI
                Cells(LIN2, 2) = RSA
                If PLANREF = "ENTRANTE" Then
                    Cells(LIN2, CLFM + 3) = TEL
                    Else
                        Cells(LIN2, 3) = TEL
                End If
                Cells(LIN2, CLFM) = OKP
                Cells(LIN2, CLFM + 1) = TCH
                Cells(LIN2, CLFM + 2) = TTC
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
If PLANREF = "SAINTE" Or PLANREF = "ROAMING" Then
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
                Cells(LFIN, CLFM + 3) = Format(Cells(LFIN, CLFM + 1) * 100 / TTCH, "0.00") * 1
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

If PLANREF = "SAINTE" Then
    PLANREF = "ROAMING"
    GoTo INICIO
End If

Application.DisplayAlerts = True
Application.ScreenUpdating = True

Unload UserForm3

UserForm4.Show 0
Application.Wait (Now + TimeValue("0:00:03"))

Unload UserForm4

End Sub
