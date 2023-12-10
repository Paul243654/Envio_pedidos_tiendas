Attribute VB_Name = "totales_tipo"
Option Explicit

Public Fruta As String
Public Verdura As String
Dim ultimafila As Long
Dim rango_a_buscar As Range
Dim valor_encontrado As Variant
Dim m As Integer
Dim suma As Double


Sub calcular_tipo_fruta()

Worksheets("Hoja2").Select

ultimafila = Hoja2.Range("A" & Rows.Count).End(xlUp).Row
suma = 0
For m = 3 To ultimafila

    Set rango_a_buscar = Range(Cells(m, 2), Cells(m, 7))

    valor_encontrado = Application.VLookup("Fruta", rango_a_buscar, 6, False)
    
    
    If IsError(valor_encontrado) Then
        valor_encontrado = 0
    Else
        
         suma = suma + valor_encontrado

     
    End If
    
Next m

Worksheets("Hoja4").Select
Range("B2").Value = suma

End Sub

Sub calcular_tipo_verdura()

Worksheets("Hoja2").Select

ultimafila = Hoja2.Range("A" & Rows.Count).End(xlUp).Row
suma = 0

For m = 3 To ultimafila

    Set rango_a_buscar = Range(Cells(m, 2), Cells(m, 7))

    valor_encontrado = Application.VLookup("Verdura", rango_a_buscar, 6, False)
    
    
    If IsError(valor_encontrado) Then
        valor_encontrado = 0
    Else
         suma = suma + valor_encontrado
    End If
    
Next m

Worksheets("Hoja4").Select
Range("B3").Value = suma

End Sub

Sub calcular_totales_tipo()

Dim total_tipos As Double

Worksheets("Hoja4").Select

Range("A1").Value = "Tipo"
Range("A2").Value = "Fruta"
Range("A3").Value = "Verdura"
Range("A4").Value = "Total"
Range("B1").Value = "Cantidad"

total_tipos = Range("B2").Value + Range("B3").Value
Range("B4").Value = total_tipos
Range("B5").Select

End Sub
