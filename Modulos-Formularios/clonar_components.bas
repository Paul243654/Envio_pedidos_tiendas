Attribute VB_Name = "clonar_components"
Sub Exportar_Modulos_Codigos()
Dim ruta, nomb, nombre, datos As String, cantidad As Byte


nomb = "Formulario versión final 30-08-19 .xlsm"
ruta = ThisWorkbook.Path & "\"

Set Proyecto = Workbooks(nomb).VBProject
For Each Module In Proyecto.VBComponents

If Module.Type = 1 Then
nombre = "C:\Users\Paul\Desktop\Paul 18_19\ANALISIS_DATOS\Proyecto\8_FORMULARIO_TIENDA_VBA\Modulos_bas\" & Module.Name & ".bas"
Module.Export nombre
End If

If Module.Type = 100 Then
nombre = "C:\Users\Paul\Desktop\Paul 18_19\ANALISIS_DATOS\Proyecto\8_FORMULARIO_TIENDA_VBA\Modulos_bas\" & Module.Name & ".txt"
Module.Export nombre
End If

If Module.Type = 3 Then
nombre = "C:\Users\Paul\Desktop\Paul 18_19\ANALISIS_DATOS\Proyecto\8_FORMULARIO_TIENDA_VBA\Modulos_bas\" & Module.Name & ".frx "
Module.Export nombre
End If

Next
End Sub
