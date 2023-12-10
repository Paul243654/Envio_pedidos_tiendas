Attribute VB_Name = "cálculo_pequeñas"
Option Explicit

Dim ultimafila As Long
Dim rango_a_buscarpequeñas As Range
Dim valor_encontrado As Variant
Dim m As Integer
Dim suma As Integer



Sub calcular_pequeñas_tienda1()

Worksheets("Hoja2").Select

ultimafila = Hoja2.Range("A" & Rows.Count).End(xlUp).Row
suma = 0
For m = 3 To ultimafila

    Set rango_a_buscarpequeñas = Range(Cells(m, 3), Cells(m, 5))

    valor_encontrado = Application.VLookup("San_Quirze", rango_a_buscarpequeñas, 3, False)
    
    
    If IsError(valor_encontrado) Then
        valor_encontrado = 0
    Else
        
         suma = suma + valor_encontrado

     
    End If
    
Next m

Worksheets("Hoja3").Select
Range("D2").Value = suma

End Sub

Sub calcular_pequeñas_tienda2()

Worksheets("Hoja2").Select

ultimafila = Hoja2.Range("A" & Rows.Count).End(xlUp).Row
suma = 0

For m = 3 To ultimafila

    Set rango_a_buscarpequeñas = Range(Cells(m, 3), Cells(m, 5))

    valor_encontrado = Application.VLookup("San_Boi", rango_a_buscarpequeñas, 3, False)
    
    
    If IsError(valor_encontrado) Then
        valor_encontrado = 0
    Else
         suma = suma + valor_encontrado
    End If
    
Next m

Worksheets("Hoja3").Select
Range("D3").Value = suma

End Sub

Sub calcular_pequeñas_tienda3()

Worksheets("Hoja2").Select

ultimafila = Hoja2.Range("A" & Rows.Count).End(xlUp).Row
suma = 0

For m = 3 To ultimafila

    Set rango_a_buscarpequeñas = Range(Cells(m, 3), Cells(m, 5))

    valor_encontrado = Application.VLookup("Mataró", rango_a_buscarpequeñas, 3, False)
    
    
    If IsError(valor_encontrado) Then
        valor_encontrado = 0
    Else
        
         suma = suma + valor_encontrado

     
    End If
    
Next m

Worksheets("Hoja3").Select
Range("D4").Value = suma

End Sub

Sub calcular_pequeñas_tienda4()

Worksheets("Hoja2").Select

ultimafila = Hoja2.Range("A" & Rows.Count).End(xlUp).Row
suma = 0

For m = 3 To ultimafila

    Set rango_a_buscarpequeñas = Range(Cells(m, 3), Cells(m, 5))

    valor_encontrado = Application.VLookup("Diagonal", rango_a_buscarpequeñas, 3, False)
    
    
    If IsError(valor_encontrado) Then
        valor_encontrado = 0
    Else
        
         suma = suma + valor_encontrado

    End If
    
Next m

Worksheets("Hoja3").Select
Range("D5").Value = suma

End Sub

Sub calcular_pequeñas_tienda5()

Worksheets("Hoja2").Select

ultimafila = Hoja2.Range("A" & Rows.Count).End(xlUp).Row
suma = 0

For m = 3 To ultimafila

    Set rango_a_buscarpequeñas = Range(Cells(m, 3), Cells(m, 5))

    valor_encontrado = Application.VLookup("San_Adria", rango_a_buscarpequeñas, 3, False)
    
    
    If IsError(valor_encontrado) Then
        valor_encontrado = 0
    Else
        
         suma = suma + valor_encontrado

     
    End If
    
Next m

Worksheets("Hoja3").Select
Range("D6").Value = suma

End Sub

Sub calcular_pequeñas_tienda6()

Worksheets("Hoja2").Select

ultimafila = Hoja2.Range("A" & Rows.Count).End(xlUp).Row
suma = 0

For m = 3 To ultimafila

    Set rango_a_buscarpequeñas = Range(Cells(m, 3), Cells(m, 5))

    valor_encontrado = Application.VLookup("Palma", rango_a_buscarpequeñas, 3, False)
    
    
    If IsError(valor_encontrado) Then
        valor_encontrado = 0
    Else
        
         suma = suma + valor_encontrado

     
    End If
    
Next m

Worksheets("Hoja3").Select
Range("D7").Value = suma

End Sub

Sub calcular_pequeñas_tienda7()

Worksheets("Hoja2").Select

ultimafila = Hoja2.Range("A" & Rows.Count).End(xlUp).Row
suma = 0

For m = 3 To ultimafila

    Set rango_a_buscarpequeñas = Range(Cells(m, 3), Cells(m, 5))

    valor_encontrado = Application.VLookup("Vilanova", rango_a_buscarpequeñas, 3, False)
    
    
    If IsError(valor_encontrado) Then
        valor_encontrado = 0
    Else
        
         suma = suma + valor_encontrado

     
    End If
    
Next m

Worksheets("Hoja3").Select
Range("D8").Value = suma

End Sub

Sub calcular_pequeñas_tienda8()

Worksheets("Hoja2").Select

ultimafila = Hoja2.Range("A" & Rows.Count).End(xlUp).Row
suma = 0

For m = 3 To ultimafila

    Set rango_a_buscarpequeñas = Range(Cells(m, 3), Cells(m, 5))

    valor_encontrado = Application.VLookup("Esplugues", rango_a_buscarpequeñas, 3, False)
    
    
    If IsError(valor_encontrado) Then
        valor_encontrado = 0
    Else
        
         suma = suma + valor_encontrado

     
    End If
    
Next m

Worksheets("Hoja3").Select
Range("D9").Value = suma

End Sub

