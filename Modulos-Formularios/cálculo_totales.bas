Attribute VB_Name = "cálculo_totales"
Option Explicit

Dim ultimafila As Long
Dim rango_a_buscar As Range
Dim valor_encontrado As Variant
Dim m As Integer
Dim suma As Integer


Sub calcular_suma_tienda1()

Worksheets("Hoja2").Select

ultimafila = Hoja2.Range("A" & Rows.Count).End(xlUp).Row
suma = 0
For m = 3 To ultimafila

    Set rango_a_buscar = Range(Cells(m, 3), Cells(m, 7))

    valor_encontrado = Application.VLookup("San_Quirze", rango_a_buscar, 5, False)
    
    
    If IsError(valor_encontrado) Then
        valor_encontrado = 0
    Else
        
         suma = suma + valor_encontrado

     
    End If
    
Next m

Worksheets("Hoja3").Select
Range("B2").Value = suma

End Sub

Sub calcular_suma_tienda2()

Worksheets("Hoja2").Select

ultimafila = Hoja2.Range("A" & Rows.Count).End(xlUp).Row
suma = 0

For m = 3 To ultimafila

    Set rango_a_buscar = Range(Cells(m, 3), Cells(m, 7))

    valor_encontrado = Application.VLookup("San_Boi", rango_a_buscar, 5, False)
    
    
    If IsError(valor_encontrado) Then
        valor_encontrado = 0
    Else
         suma = suma + valor_encontrado
    End If
    
Next m

Worksheets("Hoja3").Select
Range("B3").Value = suma

End Sub

Sub calcular_suma_tienda3()

Worksheets("Hoja2").Select

ultimafila = Hoja2.Range("A" & Rows.Count).End(xlUp).Row
suma = 0

For m = 3 To ultimafila

    Set rango_a_buscar = Range(Cells(m, 3), Cells(m, 7))

    valor_encontrado = Application.VLookup("Mataró", rango_a_buscar, 5, False)
    
    
    If IsError(valor_encontrado) Then
        valor_encontrado = 0
    Else
        
         suma = suma + valor_encontrado

     
    End If
    
Next m

Worksheets("Hoja3").Select
Range("B4").Value = suma

End Sub

Sub calcular_suma_tienda4()

Worksheets("Hoja2").Select

ultimafila = Hoja2.Range("A" & Rows.Count).End(xlUp).Row
suma = 0

For m = 3 To ultimafila

    Set rango_a_buscar = Range(Cells(m, 3), Cells(m, 7))

    valor_encontrado = Application.VLookup("Diagonal", rango_a_buscar, 5, False)
    
    
    If IsError(valor_encontrado) Then
        valor_encontrado = 0
    Else
        
         suma = suma + valor_encontrado

    End If
    
Next m

Worksheets("Hoja3").Select
Range("B5").Value = suma

End Sub

Sub calcular_suma_tienda5()

Worksheets("Hoja2").Select

ultimafila = Hoja2.Range("A" & Rows.Count).End(xlUp).Row
suma = 0

For m = 3 To ultimafila

    Set rango_a_buscar = Range(Cells(m, 3), Cells(m, 7))

    valor_encontrado = Application.VLookup("San_Adria", rango_a_buscar, 5, False)
    
    
    If IsError(valor_encontrado) Then
        valor_encontrado = 0
    Else
        
         suma = suma + valor_encontrado

     
    End If
    
Next m

Worksheets("Hoja3").Select
Range("B6").Value = suma

End Sub

Sub calcular_suma_tienda6()

Worksheets("Hoja2").Select

ultimafila = Hoja2.Range("A" & Rows.Count).End(xlUp).Row
suma = 0

For m = 3 To ultimafila

    Set rango_a_buscar = Range(Cells(m, 3), Cells(m, 7))

    valor_encontrado = Application.VLookup("Palma", rango_a_buscar, 5, False)
    
    
    If IsError(valor_encontrado) Then
        valor_encontrado = 0
    Else
        
         suma = suma + valor_encontrado

     
    End If
    
Next m

Worksheets("Hoja3").Select
Range("B7").Value = suma

End Sub

Sub calcular_suma_tienda7()

Worksheets("Hoja2").Select

ultimafila = Hoja2.Range("A" & Rows.Count).End(xlUp).Row
suma = 0

For m = 3 To ultimafila

    Set rango_a_buscar = Range(Cells(m, 3), Cells(m, 7))

    valor_encontrado = Application.VLookup("Vilanova", rango_a_buscar, 5, False)
    
    
    If IsError(valor_encontrado) Then
        valor_encontrado = 0
    Else
        
         suma = suma + valor_encontrado

     
    End If
    
Next m

Worksheets("Hoja3").Select
Range("B8").Value = suma

End Sub

Sub calcular_suma_tienda8()

Worksheets("Hoja2").Select

ultimafila = Hoja2.Range("A" & Rows.Count).End(xlUp).Row
suma = 0

For m = 3 To ultimafila

    Set rango_a_buscar = Range(Cells(m, 3), Cells(m, 7))

    valor_encontrado = Application.VLookup("Esplugues", rango_a_buscar, 5, False)
    
    
    If IsError(valor_encontrado) Then
        valor_encontrado = 0
    Else
        
         suma = suma + valor_encontrado

     
    End If
    
Next m

Worksheets("Hoja3").Select
Range("B9").Value = suma

End Sub


