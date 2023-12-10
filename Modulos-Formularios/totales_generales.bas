Attribute VB_Name = "totales_generales"
Option Explicit

Sub calcular_totales_generales()

Dim total_general As Double
Dim s As Integer
Dim r As Integer
Dim sumatoria As Double

Worksheets("Hoja3").Select
sumatoria = 0


For r = 2 To 6


        For s = 2 To 9
        
                total_general = Cells(s, r).Value
                sumatoria = sumatoria + total_general
                
        Next s
        
        Cells(10, r).Value = sumatoria
        sumatoria = 0
Next r

End Sub
