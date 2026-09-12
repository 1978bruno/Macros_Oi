Attribute VB_Name = "Módulo4"
Dim Busca
Dim PLANREF

Sub Filtro_Sainte()

PLANREF = "SAINTE"
Filtro

End Sub

Sub Filtro_Entrante()

PLANREF = "ENTRANTE"
Filtro

End Sub

Sub Filtro()


Busca = Selection

Sheets(PLANREF).Select

If ActiveSheet.AutoFilterMode = True Then Selection.AutoFilter

Rows("2:2").Select
Selection.AutoFilter
Selection.AutoFilter Field:=1, Criteria1:=Busca

End Sub
