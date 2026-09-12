Attribute VB_Name = "Módulo1"
Sub auto_open()
' Abrir_arquivo Macro
' Macro gravada em 07/03/2007 por 92033
' Atualizada em 07/03/2007 por 92033

UserForm1.Show 0
Application.Wait (Now + TimeValue("0:00:03"))

Unload UserForm1

End Sub


