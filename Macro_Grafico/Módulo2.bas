Attribute VB_Name = "Módulo2"
Public PAIS
Public RSAI
Public TELE

Sub filtro_1()

Application.ScreenUpdating = False
ActiveWindow.Visible = False
macro = ActiveWorkbook.Name
Windows(macro).Activate

Cells(1, 24) = "1"
Cells(1, 26) = "1"

LIN1 = Cells(1, 22) + 1
LIN2 = Cells(1, 24) + 1
LIN3 = Cells(1, 26) + 1
PAIS = Cells(LIN1, 22)
TELE = Cells(LIN2, 24)
RSAI = Cells(LIN3, 26)

Sheets("Base").Select

If ActiveSheet.AutoFilterMode = True Then Selection.AutoFilter
Rows("5:5").Select
Selection.AutoFilter
Selection.AutoFilter Field:=2, Criteria1:=PAIS

Cells(6, 4).Select
Range(Selection, Selection.End(xlDown)).Copy

Sheets("Grafico").Select
Cells(3, 24).Select
ActiveSheet.Paste

Selection.Sort Key1:=Range("X3"), Order1:=xlAscending, Header:=xlGuess, _
    OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom

LIN = 3
Do Until Cells(LIN, 24) = ""
    If Cells(LIN, 24) = Cells(LIN + 1, 24) Then
        Cells(LIN + 1, 24).Delete
         GoTo FMLP
    End If
    LIN = LIN + 1
FMLP:
Loop

Range(Cells(LIN, 25), Cells(LIN, 26)).Delete

Cells(3, 24).Select
Selection.End(xlDown).Select
ULTLN = ActiveCell.Row

If ULTLN = "65536" Then ULTLN = "3"

ActiveSheet.ChartObjects("Gráfico 1").Activate
ActiveChart.Shapes("Drop Down 6").Select
With Selection
    .ListFillRange = "Grafico!$X$2:$X$" & ULTLN
    .LinkedCell = "Grafico!$X$1"
    .DropDownLines = 8
    .Display3DShading = False
End With
ActiveChart.Shapes("Drop Down 8").Select
ActiveSheet.ChartObjects("Gráfico 1").Select
ActiveWindow.Visible = False
Windows(macro).Activate



Cells(1, 1).Select

LCR

Application.ScreenUpdating = True

End Sub

Sub filtro_2()

Application.ScreenUpdating = False
ActiveWindow.Visible = True
macro = ActiveWorkbook.Name
Windows(macro).Activate

Cells(1, 26) = "1"

LIN1 = Cells(1, 22) + 1
LIN2 = Cells(1, 24) + 1
LIN3 = Cells(1, 26) + 1
PAIS = Cells(LIN1, 22)
TELE = Cells(LIN2, 24)
RSAI = Cells(LIN3, 26)

Sheets("Base").Select

If ActiveSheet.AutoFilterMode = True Then Selection.AutoFilter
Rows("5:5").Select
Selection.AutoFilter
Selection.AutoFilter Field:=2, Criteria1:=PAIS
Selection.AutoFilter Field:=4, Criteria1:=TELE

Cells(6, 3).Select
Range(Selection, Selection.End(xlDown)).Copy

Sheets("Grafico").Select
Cells(3, 26).Select
ActiveSheet.Paste

Selection.Sort Key1:=Range("Z3"), Order1:=xlAscending, Header:=xlGuess, _
    OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom

LIN = 3
Do Until Cells(LIN, 26) = ""
    If Cells(LIN, 26) = Cells(LIN + 1, 26) Then
        Cells(LIN + 1, 26).Delete
        GoTo FMLP
    End If
    LIN = LIN + 1
FMLP:
Loop

Cells(3, 26).Select
Selection.End(xlDown).Select
ULTLN = ActiveCell.Row

If ULTLN = "65536" Then ULTLN = "3"

ActiveSheet.ChartObjects("Gráfico 1").Activate
ActiveChart.Shapes("Drop Down 7").Select
With Selection
    .ListFillRange = "Grafico!$z$2:$z$" & ULTLN
    .LinkedCell = "Grafico!$z$1"
    .DropDownLines = 8
    .Display3DShading = False
End With
ActiveChart.Shapes("Drop Down 8").Select
ActiveSheet.ChartObjects("Gráfico 1").Select
ActiveWindow.Visible = False
Windows(macro).Activate



'=x=x=x=x=x=x=x=x=x=x=Colorir a linha atual=x=x=x=x=x=x=x=x=x=x=x=x
For COR_ATI = 4 To 23
If Cells(COR_ATI, 2).Interior.ColorIndex = 15 Then
    Cells(COR_ATI, 2).Interior.ColorIndex = 2
    Cells(COR_ATI, 3).Interior.ColorIndex = 2
    Cells(COR_ATI, 4).Interior.ColorIndex = 2
End If
If PAIS = Cells(COR_ATI, 2) And TELE = Cells(COR_ATI, 3) Then
    Cells(COR_ATI, 2).Interior.ColorIndex = 15
    Cells(COR_ATI, 3).Interior.ColorIndex = 15
    Cells(COR_ATI, 4).Interior.ColorIndex = 15
End If

Next

'=x=x=x=x=x=x=x=x=x=x=x=x=x=x=x=x=FIM=x=x=x=x=x=x=x=x=x=x=x=xx=x=x

Cells(1, 1).Select

LCR

Application.ScreenUpdating = True

End Sub

Sub filtro_3()

Application.ScreenUpdating = False
ActiveWindow.Visible = False
macro = ActiveWorkbook.Name
Windows(macro).Activate

LIN1 = Cells(1, 22) + 1
LIN2 = Cells(1, 24) + 1
LIN3 = Cells(1, 26) + 1
PAIS = Cells(LIN1, 22)
TELE = Cells(LIN2, 24)
RSAI = Cells(LIN3, 26)

Sheets("Base").Select

If ActiveSheet.AutoFilterMode = True Then Selection.AutoFilter
Rows("5:5").Select
Selection.AutoFilter
Selection.AutoFilter Field:=2, Criteria1:=PAIS
Selection.AutoFilter Field:=3, Criteria1:=RSAI
Selection.AutoFilter Field:=4, Criteria1:=TELE

Sheets("Grafico").Select
ActiveWindow.Visible = False
Windows(macro).Activate

Cells(1, 1).Select

Application.ScreenUpdating = True


End Sub

Sub LCR()


Sheets("LCR").Visible = True
Sheets("LCR").Select

BUSCA = PAIS & TELE

Range("C:C").Select

Selection.Find(What:=BUSCA).Activate
LINR = ActiveCell.Row
On Error Resume Next    ' Ative a rotina de tratamento de erro.

OP1 = Cells(LINR, 6)
COP1 = Cells(LINR, 6).Comment.Text
OP2 = Cells(LINR, 7)
COP2 = Cells(LINR, 7).Comment.Text
OP3 = Cells(LINR, 8)
COP3 = Cells(LINR, 8).Comment.Text
OP4 = Cells(LINR, 9)
COP4 = Cells(LINR, 9).Comment.Text
OP5 = Cells(LINR, 10)
COP5 = Cells(LINR, 10).Comment.Text

Sheets("Grafico").Select

Cells(29, 7) = OP1
Cells(30, 7) = OP2
Cells(31, 7) = OP3
Cells(32, 7) = OP4
Cells(33, 7) = OP5
Cells(29, 9) = COP1
Cells(30, 9) = COP2
Cells(31, 9) = COP3
Cells(32, 9) = COP4
Cells(33, 9) = COP5

Err.Clear

Sheets("LCR").Visible = False

End Sub

