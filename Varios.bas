Attribute VB_Name = "Varios"
Option Explicit
'
'Public Sub comparar_columnas()
'
'Call preparacionInicio
'
'Dim i, j, k As Long
'Dim columnaHoja1, filaHoja1 As Long
'Dim auxHora As Date
'
'auxHora = Now
'
'columnaHoja1 = Sheets("Enviados").Cells(1, Columns.Count).End(xlToLeft).Column
'filaHoja1 = Sheets("Enviados").Cells(Rows.Count, 1).End(xlUp).Row
'
'For i = 1 To columnaHoja1
'
'    For j = 2 To filaHoja1
'
'        If (Sheets("Enviados").Cells(j, i).Value <> Sheets("Recibido").Cells(j, i).Value) Then
'
'            Sheets("Recibido").Cells(j, i).Interior.Color = RGB(255, 255, 0)
'            Sheets("Enviados").Cells(j, i).Interior.Color = RGB(255, 255, 0)
'
'        End If
'
'    Next j
'
'Next i
'
'Call finalizarTestVelocidad(auxHora)
'Call preparacionFinal
'
'End Sub

'compara la cantidad de registros entre las ultimas 2 solapas de un archivo
Public Sub comparar_cantidad_registros()

Dim i, j, columnaHoja1, filas As Long
Dim registrosHoja1, registrosHoja2 As Double
Dim rango1, rango2 As Range
Dim hoja1, hoja2 As Worksheet
Dim resumen() As String
Dim auxHora As Date

'preparo el archivo
Call preparacionInicio

'otorgo valores a varias varibales
auxHora = Now
Set hoja1 = ActiveWorkbook.Sheets(Sheets.Count - 1)
Set hoja2 = ActiveWorkbook.Sheets(Sheets.Count)
columnaHoja1 = hoja1.Cells(1, Columns.Count).End(xlToLeft).Column
ReDim resumen(1 To columnaHoja1, 1 To 4)

'convierto el formato de las hojas en general porque si no uno de los if con la condicion
'"(hoja2.Cells(j, i).Value = "")" no funciona
hoja1.Cells.NumberFormat = "General"
hoja2.Cells.NumberFormat = "General"

'recorro las columnas para obtener la cantidad de filas maximas
For i = 1 To columnaHoja1
    
    If (filas < hoja1.Cells(Rows.Count, i).End(xlUp).Row) Then
        
        filas = hoja1.Cells(Rows.Count, i).End(xlUp).Row
        
    End If

Next i

'recorro nuevamente las columnas
For i = 1 To columnaHoja1

    'recorro la totalidad de las filas de la columna i
    For j = 1 To filas
        
        'este if sirve para limpiar las casillas en blanco que devuelve IDEA que excel no considera vacias
        If (hoja2.Cells(j, i).Value = "") Then
            
            hoja2.Cells(j, i).Value = ""
            
        End If
        
    Next j
    
    'otorgo valores a distintas variables
    Set rango1 = hoja1.Range(hoja1.Cells(1, i), hoja1.Cells(filas, i))
    Set rango2 = hoja2.Range(hoja2.Cells(1, i), hoja2.Cells(filas, i))
    registrosHoja1 = rango1.Cells.SpecialCells(xlCellTypeConstants).Count
    registrosHoja2 = rango2.Cells.SpecialCells(xlCellTypeConstants).Count
    
    'completo los valores de la variable "resumen" para asi armar la solapa "Resumen"
    resumen(i, 1) = hoja1.Cells(1, i).Value
    resumen(i, 2) = registrosHoja1
    resumen(i, 3) = registrosHoja2
    resumen(i, 4) = registrosHoja1 - registrosHoja2
    
Next i

'llamo a la funcion cuadrosComparacionCantidadRegistros y vuelvo a preparar el archivo
Call cuadrosComparacionCantidadRegistros(hoja1, hoja2, resumen, columnaHoja1)
Call preparacionFinal

End Sub

