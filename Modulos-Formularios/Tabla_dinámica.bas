Attribute VB_Name = "Tabla_dinámica"
Option Explicit

Sub crear_tabla_dinamica()

Rem datos es la hoja donde se encuentra nuestra base de datos
Dim datos As Worksheet
Rem tdp es la hoja donde se ejecuta la tabla pivotante
Dim tdp As Worksheet
Rem PTcache es la memoria destinada a almacenar datos para que en solicitudes futuras estos datos puedan atenderse con mayor rápidez.
Dim PTcache As PivotCache
Rem tabladinámica es la tabla pivotamte que crearemos
Dim Tabladinámica As PivotTable
Rem Rangodatos es el rango de nuestra base de datos en la hoja datos
Dim Rangodatos As Range
Rem últimafila es la ultima linea ingresada en nuestra base de datos
Dim últimafila As Long


Rem 1ro borramos la tabla dinámica que se encuentra en la hoja dinámica (en forma de actualización)

For Each Tabladinámica In Worksheets("Hoja5").PivotTables
        Tabladinámica.TableRange2.clear 'porque usa en nº2?
Next Tabladinámica

Rem 2do definimos cual sera nuestro rango a utilizar d ela base de datos y establecemos el cache dinámico

últimafila = Worksheets("Hoja2").Cells(Rows.Count, 1).End(xlUp).Row 'aquí contamos las filas de la columna 1 hasta el final que tenga algun dato

Set Rangodatos = Worksheets("Hoja2").Cells(2, 1).Resize(últimafila, 7) 'aqui seleccionamos desde A2 hasta (ultimafila,7)

Rem 3ro nos situamos en la hoja que tiene nuestra base de datos y
'definimos la variable PTcache como valor intermedio necesario para la creación de la tabla dinámica

Sheets("Hoja2").Select

Set PTcache = ActiveWorkbook.PivotCaches.Add(SourceType:=xlDatabase, SourceData:=Rangodatos.Address)

Rem 4to creamos una tabla dinámica en blanco y especificamos la hoja donde se ejecutara nuestra tabla y definimos su nombre

Set Tabladinámica = PTcache.CreatePivotTable(tabledestination:=Worksheets("Hoja5").Range("A1"), tablename:="pivottable1")

Rem 5to aplicamos el formato predefinido para tablas dinámicas

Tabladinámica.Format xlReport6

Rem 6to actualización automática definiendo campos

Tabladinámica.ManualUpdate = True

Tabladinámica.AddFields RowFields:=Array("Tienda") 'aqui definimos que filtros se utilizaran en la tabla dinámica

Rem 7mo introducimos los campos que queremos que aparezcan en nuestra tabla dinámica

With Tabladinámica.PivotFields("Tienda")
.Orientation = xlDataField
.Function = xlCount
.Position = 1
.Caption = "Palets"
End With

With Tabladinámica.PivotFields("Gran.")
.Orientation = xlDataField
.Function = xlSum
.Position = 2
.Caption = "cajas grandes"
End With

With Tabladinámica.PivotFields("Peq.")
.Orientation = xlDataField
.Function = xlSum
.Position = 3
.Caption = "Cajas pequeñas"
End With

With Tabladinámica.PivotFields("Cart.")
.Orientation = xlDataField
.Function = xlSum
.Position = 4
.Caption = "Cajas carton"
End With


With Tabladinámica.PivotFields("Total")
.Orientation = xlDataField
.Function = xlSum
.Position = 5
.Caption = "Total cajas"
End With

Tabladinámica.ManualUpdate = False ' con esto quiere decir que la tabla no se actualiza sola cada vez que se introduce un dato para no ralentiza,
' por ello se ejecuta su actualización en el boton calcular de este formulario.
'tambien hay opciones que si se trabaja con hojas, cada vez que se abre una hoja ó se eesta en ella se actualizae la tabla dinamica,
' y eso se hace haciendo clik en hoja y escribiendo el codigo de refresh en esa hoja en la parte del modulo.


End Sub

