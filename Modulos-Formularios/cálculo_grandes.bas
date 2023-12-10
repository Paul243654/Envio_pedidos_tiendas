Attribute VB_Name = "cálculo_grandes"
Option Explicit

Dim ultimafila As Long
Dim rango_a_buscargrandes As Range
Dim valor_encontrado As Variant
Dim m As Integer
Dim suma As Integer
Public San_Quirze As String
Public San_Boi As String
Public Mataró As String
Public Diagonal As String
Public San_Adria As String
Public Palma As String
Public Vilanova As String
Public Esplugues As String


Sub calcular_grandes_tienda1()

Worksheets("Hoja2").Select

ultimafila = Hoja2.Range("A" & Rows.Count).End(xlUp).Row
suma = 0
For m = 3 To ultimafila

    Set rango_a_buscargrandes = Range(Cells(m, 3), Cells(m, 4))

    valor_encontrado = Application.VLookup("San_Quirze", rango_a_buscargrandes, 2, False)
    
    
    If IsError(valor_encontrado) Then
        valor_encontrado = 0
    Else
        
         suma = suma + valor_encontrado

     
    End If
    
Next m

Worksheets("Hoja3").Select
Range("C2").Value = suma

End Sub

Sub calcular_grandes_tienda2()

Worksheets("Hoja2").Select

ultimafila = Hoja2.Range("A" & Rows.Count).End(xlUp).Row
suma = 0

For m = 3 To ultimafila

    Set rango_a_buscargrandes = Range(Cells(m, 3), Cells(m, 4))

    valor_encontrado = Application.VLookup("San_Boi", rango_a_buscargrandes, 2, False)
    
    
    If IsError(valor_encontrado) Then
        valor_encontrado = 0
    Else
         suma = suma + valor_encontrado
    End If
    
Next m

Worksheets("Hoja3").Select
Range("C3").Value = suma

End Sub

Sub calcular_grandes_tienda3()

Worksheets("Hoja2").Select

ultimafila = Hoja2.Range("A" & Rows.Count).End(xlUp).Row
suma = 0

For m = 3 To ultimafila

    Set rango_a_buscargrandes = Range(Cells(m, 3), Cells(m, 4))

    valor_encontrado = Application.VLookup("Mataró", rango_a_buscargrandes, 2, False)
    
    
    If IsError(valor_encontrado) Then
        valor_encontrado = 0
    Else
        
         suma = suma + valor_encontrado

     
    End If
    
Next m

Worksheets("Hoja3").Select
Range("C4").Value = suma

End Sub

Sub calcular_grandes_tienda4()

Worksheets("Hoja2").Select

ultimafila = Hoja2.Range("A" & Rows.Count).End(xlUp).Row
suma = 0

For m = 3 To ultimafila

    Set rango_a_buscargrandes = Range(Cells(m, 3), Cells(m, 4))

    valor_encontrado = Application.VLookup("Diagonal", rango_a_buscargrandes, 2, False)
    
    
    If IsError(valor_encontrado) Then
        valor_encontrado = 0
    Else
        
         suma = suma + valor_encontrado

    End If
    
Next m

Worksheets("Hoja3").Select
Range("C5").Value = suma

End Sub

Sub calcular_grandes_tienda5()

Worksheets("Hoja2").Select

ultimafila = Hoja2.Range("A" & Rows.Count).End(xlUp).Row
suma = 0

For m = 3 To ultimafila

    Set rango_a_buscargrandes = Range(Cells(m, 3), Cells(m, 4))

    valor_encontrado = Application.VLookup("San_Adria", rango_a_buscargrandes, 2, False)
    
    
    If IsError(valor_encontrado) Then
        valor_encontrado = 0
    Else
        
         suma = suma + valor_encontrado

     
    End If
    
Next m

Worksheets("Hoja3").Select
Range("C6").Value = suma

End Sub

Sub calcular_grandes_tienda6()

Worksheets("Hoja2").Select

ultimafila = Hoja2.Range("A" & Rows.Count).End(xlUp).Row
suma = 0

For m = 3 To ultimafila

    Set rango_a_buscargrandes = Range(Cells(m, 3), Cells(m, 4))

    valor_encontrado = Application.VLookup("Palma", rango_a_buscargrandes, 2, False)
    
    
    If IsError(valor_encontrado) Then
        valor_encontrado = 0
    Else
        
         suma = suma + valor_encontrado

     
    End If
    
Next m

Worksheets("Hoja3").Select
Range("C7").Value = suma

End Sub

Sub calcular_grandes_tienda7()

Worksheets("Hoja2").Select

ultimafila = Hoja2.Range("A" & Rows.Count).End(xlUp).Row
suma = 0

For m = 3 To ultimafila

    Set rango_a_buscargrandes = Range(Cells(m, 3), Cells(m, 4))

    valor_encontrado = Application.VLookup("Vilanova", rango_a_buscargrandes, 2, False)
    
    
    If IsError(valor_encontrado) Then
        valor_encontrado = 0
    Else
        
         suma = suma + valor_encontrado

     
    End If
    
Next m

Worksheets("Hoja3").Select
Range("C8").Value = suma

End Sub

Sub calcular_grandes_tienda8()

Worksheets("Hoja2").Select

ultimafila = Hoja2.Range("A" & Rows.Count).End(xlUp).Row
suma = 0

For m = 3 To ultimafila

    Set rango_a_buscargrandes = Range(Cells(m, 3), Cells(m, 4))

    valor_encontrado = Application.VLookup("Esplugues", rango_a_buscargrandes, 2, False)
    
    
    If IsError(valor_encontrado) Then
        valor_encontrado = 0
    Else
        
         suma = suma + valor_encontrado

     
    End If
    
Next m

Worksheets("Hoja3").Select
Range("C9").Value = suma

End Sub

