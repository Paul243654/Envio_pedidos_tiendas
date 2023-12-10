Attribute VB_Name = "Copiar_columnas_ahoja1"
Option Explicit

Sub copiar_columnas()

Worksheets("Hoja1").Select
Worksheets("Hoja1").Cells.clear

Hoja2.Columns(1).Copy Hoja1.Columns(1)
Hoja2.Columns(2).Copy Hoja1.Columns(2)
Hoja2.Columns(3).Copy Hoja1.Columns(3)
Hoja2.Columns(4).Copy Hoja1.Columns(4)
Hoja2.Columns(5).Copy Hoja1.Columns(5)
Hoja2.Columns(6).Copy Hoja1.Columns(6)
Hoja2.Columns(7).Copy Hoja1.Columns(7)


End Sub
