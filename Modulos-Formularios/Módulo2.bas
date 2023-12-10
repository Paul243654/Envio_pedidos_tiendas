Attribute VB_Name = "Módulo2"
Option Explicit

Sub Auto_Open()
Dim msgvalue As Integer
Dim inicio As String

inicio = MsgBox(" ¿Estan cerrados todos los libros excel en este ordenador? ", vbYesNo + vbExclamation, " ADVERTENCIA ")

Select Case inicio
    Case vbYes
        formulariollamada.Show
    Case vbNo
        Application.Quit
 End Select
End Sub
