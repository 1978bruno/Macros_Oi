Attribute VB_Name = "Módulo3"
Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" _
(ByVal lpBuffer As String, nSize As Long) As Long

Private Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" _
(ByVal lpBuffer As String, nSize As Long) As Long

Public PLANREF
Public TOPREF
Public DTI As Date
Public DTF As Date
Public DTT As Date
Public DT1 As Date
Public DT2 As Date

Sub atu_sainte_geral()

PLANREF = "SAINTE"
atualiza_sainte

End Sub

Sub atu_sainte_roaming()

PLANREF = "ROAMING"
atualiza_sainte

End Sub

Sub atualiza_sainte()

UserForm2.Show

Sheets.Add
ATUA = ActiveSheet.Name



Sheets(PLANREF).Select

If ActiveSheet.AutoFilterMode = True Then Selection.AutoFilter

Range("A:C").Select
Selection.Copy

Sheets(ATUA).Select
ActiveSheet.Paste
Range("B:B").Delete

Range("3:6500").Select
Selection.Sort Key1:=Range("A3"), Order1:=xlDescending, Key2:=Range("B3"), Order2:=xlDescending, _
    Header:=xlGuess, OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom

LIN = 3

'Excluir repetições
Do Until Cells(LIN, 1) = ""
    If Cells(LIN, 1) = Cells(LIN + 1, 1) And Cells(LIN, 2) = Cells(LIN + 1, 2) Then
        Rows(LIN + 1).Delete
        Else
            LIN = LIN + 1
    End If
Loop

Sheets(PLANREF).Select
LIN = 3
LIN2 = 3



Cells(2, 4).Select
Selection.End(xlToRight).Select
UTCOL = ActiveCell.Column - 2
DT = Cells(1, UTCOL - 1)

Sheets(ATUA).Select

Do Until Cells(LIN, 1) = ""
    PAI = Cells(LIN, 1)
    TEL = Cells(LIN, 2)
    Sheets(PLANREF).Select
    LIN2 = 3
    Do Until Cells(LIN2, 1) = ""
        If PAI = Cells(LIN2, 1) And TEL = Cells(LIN2, 3) Then TTCH = TTCH + Cells(LIN2, UTCOL)
        LIN2 = LIN2 + 1
    Loop
    Sheets(ATUA).Select
    Cells(LIN, 3) = TTCH
    TCHF = TCHF + TTCH
    TTCH = 0
    LIN = LIN + 1
Loop
Sheets(ATUA).Select

Range("3:6500").Select
Selection.Sort Key1:=Range("C3"), Order1:=xlDescending, Header:= _
    xlGuess, OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom

Range(Rows(TOPREF + 4), Selection.End(xlDown)).Delete

LIN3 = 3
Do Until Cells(LIN3, 1) = ""
    Cells(LIN3, 4) = Format(Cells(LIN3, 3) / TCHF * 100, "0.00") * 1
    LIN3 = LIN3 + 1
Loop

Range(Cells(3, 1), Cells(TOPREF + 3, 4)).Copy

Sheets("FACE").Select

If PLANREF = "SAINTE" Then
    REFCOL = 13
    CABÇ = "TOP" & TOPREF & " - PAISES"
    'Cells(10, 14).Select
    'Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
    '    :=False, Transpose:=False
    'Cells(7, 14) = "TOP" & TOPREF & " - PAISES"
    'Cells(8, 14) = "DATA: " & DT
    'Range(Cells(TOPREF + 10, 13), Cells(79, 17)).Clear
    'Cells(TOPREF + 10, 13) = "Ultima Atualização: " & Now
    'Cells(TOPREF + 11, 20).Copy
    'Range(Cells(TOPREF + 11, 13), Cells(79, 17)).Select
    'Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
    '    SkipBlanks:=False, Transpose:=False
    'Range(Cells(10, 13), Cells(10, 17)).Copy
    'Range(Cells(10, 13), Cells(TOPREF + 10, 17)).Select
    'Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
    '    SkipBlanks:=False, Transpose:=False
    Else
        REFCOL = 2
        CABÇ = "TOP" & TOPREF & " ROAMING - PAISES"
     '   Cells(10, 3).Select
     '   Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
     '       :=False, Transpose:=False
     '   Cells(7, 2) = "TOP" & TOPREF & " - PAISES - ROAMING"
     '   Cells(8, 2) = "DATA: " & DT
     '   Range(Cells(TOPREF + 11, 2), Cells(79, 6)).Clear
     '   Cells(TOPREF + 10, 2) = "Ultima Atualização: " & Now
     '   Cells(TOPREF + 11, 20).Copy
     '   Range(Cells(TOPREF + 11, 2), Cells(79, 6)).Select
     '   Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
     '       SkipBlanks:=False, Transpose:=False
     '   Range(Cells(10, 2), Selection(xlToight)).Copy
     '   Range(Cells(10, 2), Cells(TOPREF + 10, 6)).Select
     '   Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
     '    SkipBlanks:=False, Transpose:=False
End If
Cells(10, REFCOL + 1).Select
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
    :=False, Transpose:=False
Cells(7, REFCOL) = CABÇ
Cells(8, REFCOL) = "DATA: " & DT
Range(Cells(TOPREF + 10, REFCOL), Cells(80, REFCOL + 4)).Clear
Cells(TOPREF + 10, REFCOL) = "Ultima Atualização: " & Now
Cells(TOPREF + 11, REFCOL + 6).Copy
Range(Cells(TOPREF + 11, REFCOL), Cells(79, REFCOL + 4)).Select
Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
    SkipBlanks:=False, Transpose:=False
Range(Cells(10, REFCOL), Cells(10, REFCOL + 4)).Copy
Range(Cells(10, REFCOL), Cells(TOPREF + 9, REFCOL + 4)).Select
Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
    SkipBlanks:=False, Transpose:=False
Range(Cells(10, REFCOL), Cells(11, REFCOL)).Select
Selection.AutoFill Destination:=Range(Cells(10, REFCOL), Cells(TOPREF + 9, REFCOL))
Range(Cells(TOPREF + 9, REFCOL), Cells(TOPREF + 9, REFCOL + 4)).Select
With Selection.Borders(xlEdgeBottom)
    .Weight = xlThick
    .ColorIndex = xlAutomatic
End With
Cells(1, 1).Select

Sheets(ATUA).Delete

Ult_Atua
    
End Sub

Sub atualiza_periodo()

UserForm2.Show

For ST = 1 To 2
    If ST = 1 Then
        PLANREF = "SAINTE"
        Else
            PLANREF = "ROAMING"
            TCHF = 0
    End If


    Sheets(PLANREF).Select
    Cells(2, 4).Select
    Selection.End(xlToRight).Select
    CLFM = ActiveCell.Column
    
    DT1 = Cells(1, 4)
    DT2 = Cells(1, CLFM - 3)
    
    UserForm5.Show
    CLFI = 4
    CLFF = 4
    'Localizar data Inicio
    Do Until DTI = DT1
        CLFI = CLFI + 1
        DT1 = Cells(1, CLFI)
    Loop
    'localizar data Fim
    If DTF = DT2 Then
        CLFF = CLFM
        Else
            Do Until DTF = DT2
                CLFF = CLFF + 1
                DT2 = Cells(1, CLFF)
            Loop
    End If


    Sheets.Add
    ATUA = ActiveSheet.Name
    Sheets(PLANREF).Select
    
    If ActiveSheet.AutoFilterMode = True Then Selection.AutoFilter
    
    Range("A:C").Select
    Selection.Copy
    
    Sheets(ATUA).Select
    ActiveSheet.Paste
    Range("B:B").Delete
    
    Range("3:6500").Select
    Selection.Sort Key1:=Range("A3"), Order1:=xlDescending, Key2:=Range("B3"), Order2:=xlDescending, _
        Header:=xlGuess, OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom
    
    LIN = 3
    
    'Excluir repetições
    Do Until Cells(LIN, 1) = ""
        If Cells(LIN, 1) = Cells(LIN + 1, 1) And Cells(LIN, 2) = Cells(LIN + 1, 2) Then
            Rows(LIN + 1).Delete
            Else
                LIN = LIN + 1
        End If
    Loop
    
    Sheets(PLANREF).Select
    LIN = 3
    LIN2 = 3
    
    
    
    Cells(2, 4).Select
    Selection.End(xlToRight).Select
    UTCOL = ActiveCell.Column - 2
    DT = Cells(1, UTCOL - 1)
    
    Sheets(ATUA).Select
    
    Do Until Cells(LIN, 1) = ""
        PAI = Cells(LIN, 1)
        TEL = Cells(LIN, 2)
        Sheets(PLANREF).Select
        LIN2 = 3
        COLU = 1
        Do Until Cells(LIN2, 1) = ""
            COLU = 1
            If PAI = Cells(LIN2, 1) And TEL = Cells(LIN2, 3) Then
                Do Until CLFF + 1 < CLFI + COLU
                    TTCH = TTCH + Cells(LIN2, CLFI + COLU)
                    COLU = COLU + 4
                Loop
            End If
            LIN2 = LIN2 + 1
        Loop
        Sheets(ATUA).Select
        Cells(LIN, 3) = TTCH
        TCHF = TCHF + TTCH
        TTCH = 0
        LIN = LIN + 1
    Loop
    Sheets(ATUA).Select
    
    Range("3:6500").Select
    Selection.Sort Key1:=Range("C3"), Order1:=xlDescending, Header:= _
        xlGuess, OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom

    Range(Rows(TOPREF + 4), Selection.End(xlDown)).Delete
    
    LIN3 = 3
    Do Until Cells(LIN3, 1) = ""
        Cells(LIN3, 4) = Format(Cells(LIN3, 3) / TCHF * 100, "0.00") * 1
        LIN3 = LIN3 + 1
    Loop
    
    Range(Cells(3, 1), Cells(TOPREF + 3, 4)).Copy
    
    Sheets("FACE").Select
    
    If PLANREF = "SAINTE" Then
        REFCOL = 13
        CABÇ = "TOP" & TOPREF & " - PAISES"
        Else
            REFCOL = 2
            CABÇ = "TOP" & TOPREF & " ROAMING - PAISES"
    End If
    Cells(10, REFCOL + 1).Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Cells(7, REFCOL) = CABÇ
    Cells(8, REFCOL) = "Periodo: " & DTI & " à " & DTF
    Range(Cells(TOPREF + 10, REFCOL), Cells(80, REFCOL + 4)).Clear
    Cells(TOPREF + 10, REFCOL) = "Ultima Atualização: " & Now
    Cells(TOPREF + 11, REFCOL + 6).Copy
    Range(Cells(TOPREF + 11, REFCOL), Cells(79, REFCOL + 4)).Select
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Range(Cells(10, REFCOL), Cells(10, REFCOL + 4)).Copy
    Range(Cells(10, REFCOL), Cells(TOPREF + 9, REFCOL + 4)).Select
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Range(Cells(10, REFCOL), Cells(11, REFCOL)).Select
    Selection.AutoFill Destination:=Range(Cells(10, REFCOL), Cells(TOPREF + 9, REFCOL))
    Range(Cells(TOPREF + 9, REFCOL), Cells(TOPREF + 9, REFCOL + 4)).Select
    With Selection.Borders(xlEdgeBottom)
        .Weight = xlThick
        .ColorIndex = xlAutomatic
    End With
    Cells(1, 1).Select
    
    Sheets(ATUA).Delete
Next

Ult_Atua
    
End Sub

Sub atualiza_entrante()

Sheets.Add
ATUA = ActiveSheet.Name

Sheets("ENTRANTE").Select

If ActiveSheet.AutoFilterMode = True Then Selection.AutoFilter

Range("A:A").Select
Selection.Copy

Sheets(ATUA).Select
ActiveSheet.Paste

LIN = 3

Do Until Cells(LIN, 1) = ""
    If Cells(LIN, 1) = Cells(LIN + 1, 1) Then
        Rows(LIN + 1).Delete
        Else
            LIN = LIN + 1
    End If

Loop

Sheets("ENTRANTE").Select
LIN = 3
LIN2 = 3



Cells(2, 4).Select
Selection.End(xlToRight).Select
UTCOL = ActiveCell.Column - 1
DT = Cells(1, UTCOL - 1)

Sheets(ATUA).Select

Do Until Cells(LIN, 1) = ""
    PAI = Cells(LIN, 1)
    Sheets("ENTRANTE").Select
    LIN2 = 3
    Do Until Cells(LIN2, 1) = ""
        If PAI = Cells(LIN2, 1) Then TTCH = TTCH + Cells(LIN2, UTCOL)
        LIN2 = LIN2 + 1
    Loop
    Sheets(ATUA).Select
    Cells(LIN, 2) = TTCH
    TTCH = 0
    LIN = LIN + 1
Loop

Sheets(ATUA).Select
Range("3:6500").Select
Selection.Sort Key1:=Range("B3"), Order1:=xlDescending, Header:= _
    xlGuess, OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom
    
Range(Cells(3, 1), Cells(12, 2)).Copy

Sheets("FACE").Select
Range("R35").Select
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
    :=False, Transpose:=False
Range("Q33") = "DATA: " & DT
Range("Q45") = "Ultima Atualização: " & Now

Sheets(ATUA).Delete

Ult_Atua

End Sub

Sub Ult_Atua()

Dim Comp_Name_B As String * 255
Dim usuario As String * 255

GetComputerName Comp_Name_B, Len(Comp_Name_B)
GetUserName usuario, Len(usuario)


Comp_Name = Left(Comp_Name_B, InStr(Comp_Name_B, Chr(0)) - 1)
USU_Name = Left(usuario, InStr(usuario, Chr(0)) - 1)


Cells(4, 1) = USU_Name & " usando a máquina " & Comp_Name
End Sub
