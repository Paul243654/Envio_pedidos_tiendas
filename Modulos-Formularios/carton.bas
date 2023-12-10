Attribute VB_Name = "carton"
Option Explicit

Sub calcular_carton_tiendas()

Dim total_carton As Integer
Dim z As Integer

Worksheets("Hoja3").Select

For z = 2 To 9

        total_carton = (Cells(z, 2).Value) - (Cells(z, 3).Value) - (Cells(z, 4).Value)
        Cells(z, 5).Value = total_carton
Next z


End Sub
