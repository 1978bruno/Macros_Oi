Attribute VB_Name = "Módulo3"
Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" _
(ByVal lpBuffer As String, nSize As Long) As Long

Private Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" _
(ByVal lpBuffer As String, nSize As Long) As Long

Public PLANREF

Sub atu_sainte_geral()

PLANREF = "SAINTE"
atualiza_sainte

End Sub

Sub atu_sainte_pan()

PLANREF = "ROOMING"
atualiza_sainte

End Sub

Sub atualiza_sainte()

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
UTCOL = ActiveCell.Column - 1
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
    TTCH = 0
    LIN = LIN + 1
Loop

Sheets(ATUA).Select
Range("3:6500").Select
Selection.Sort Key1:=Range("C3"), Order1:=xlDescending, Header:= _
    xlGuess, OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom
    
Range(Cells(3, 1), Cells(22, 3)).Copy

Sheets("FACE").Select

If PLANREF = "SAINTE" Then
    Range("R10").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("Q8") = "DATA: " & DT
    Range("Q30") = "Ultima Atualização: " & Now
    Else
        Range("C10").Select
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
        
        Range("D10:E29").Copy
        Range("F10").Select
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
        Range("D10:E29").ClearContents
        Range("B8") = "DATA: " & DT
        Range("B30") = "Ultima Atualização: " & Now
End If

Sheets(ATUA).Delete

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
