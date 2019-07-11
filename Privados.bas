Attribute VB_Name = "Privados"
Option Explicit

'resumen: preparacion previa al incio de funcion
'parametros: void
'retorno: void
Public Function preparacionInicio()

Application.EnableEvents = False
Application.ScreenUpdating = False
Application.DisplayStatusBar = False
Application.Calculation = xlCalculationManual

End Function

'resumen: preparacion para el cierre de la funcion
'parametros: void
'retorno: void
Public Function preparacionFinal()

Application.EnableEvents = True
Application.ScreenUpdating = True
Application.DisplayStatusBar = True
Application.Calculation = xlCalculationAutomatic

End Function

'resumen: imprime por pantalla el tiempo que tardo en ejecutarse un progama
'parametros: variable que contiene en que momento se emepezo la ejecucion
'retorno: void
Public Function finalizarTestVelocidad(ByVal auxAhora As Date)

Dim tiempoSegundos, tiempoMinutos As Long

tiempoSegundos = DateDiff("s", auxAhora, Now)
tiempoMinutos = tiempoSegundos / 60

If (tiempoSegundos <= 60) Then
    
    Debug.Print ("Termino en " & tiempoSegundos & " segundos")
    
Else

    Debug.Print ("Termino en " & tiempoMinutos & " minutos")

End If

End Function

'resumen: imprime por pantalla el nombre de todos los archivos abiertos
'parametros: void
'retorno: void
Public Function listarArchivos()

Dim auxArch As Workbook

For Each auxArch In Workbooks
    
    Debug.Print auxArch.name

Next

End Function

'resumen: determina la cantidad de columnas de la fila 1 de la hoja activa
'parametros: void
'retorno: la cantidad de columnas en un entero
Public Function columnaMax(ByVal hoja As Worksheet) As Integer

columnaMax = hoja.Cells(1, Columns.Count).End(xlToLeft).Column

End Function

'resumen: determina la cantidad de filas maxi para la hoja activa
'parametros: la cantidad de columnas de la hoja activa
'retorno: la cantidad de filas maxima en un entero
Public Function filaMax(ByVal cantidadColumnas As Long, ByVal hoja As Worksheet) As Double

Dim i, filas As Double

    For i = 1 To cantidadColumnas
    
    If (filas < hoja.Cells(Rows.Count, i).End(xlUp).Row) Then
        
        filas = hoja.Cells(Rows.Count, i).End(xlUp).Row
        
    End If

Next i

filaMax = filas
    
End Function

'resumen: le hace un clear a la hoja y la pinta de blanco
'parametros: la hoja a limpiar
'retorno: void
Function limpiarHoja(ByVal wsHoja As Worksheet)

'limpieza
With wsHoja.Cells
    
    .Clear
    .Interior.Color = RGB(255, 255, 255)
    
End With

End Function

'resumen: crea una hoja nueva
'parametros: nombres es el nombre nuevo de la hoja
'retorno: void
Function crearHoja(ByVal nombre As String)
    
'creo una sola nueva y la nombro
On Error GoTo erroralcrearhoja

    Sheets.Add After:=Sheets(Sheets.Count)
    Sheets(Sheets.Count).name = nombre
    
    Exit Function
    
erroralcrearhoja:
Sheets(Sheets.Count).name = nombre & Sheets.Count - 1
Resume Next
    
End Function

'resumen: elimina una hoja
'parametros: nombre de la hoja a eliminar
'retorno: void
Public Function eliminarHoja(ByVal nombre As String)

ActiveWorkbook.Sheets(nombre).Delete

End Function

'resumen: descombina las celdas y ajusta el texto
'parametros: wsHoja es la hoja donde se ejecuta el metodo
'retorno: void
Public Function desagruparCeldasHoja(ByVal wsHoja As Worksheet)

wsHoja.Cells.Select

With Selection
    .HorizontalAlignment = xlGeneral
    .VerticalAlignment = xlBottom
    .WrapText = True
    .Orientation = 0
    .AddIndent = False
    .IndentLevel = 0
    .ShrinkToFit = False
    .ReadingOrder = xlContext
End With

Selection.UnMerge

With Selection
    .HorizontalAlignment = xlGeneral
    .VerticalAlignment = xlBottom
    .WrapText = False
    .Orientation = 0
    .AddIndent = False
    .IndentLevel = 0
    .ShrinkToFit = False
    .ReadingOrder = xlContext
    .MergeCells = False
End With

With wsHoja
    .Cells.Select
    .Cells.EntireRow.AutoFit
End With
    
End Function


'resumen: descombina las celdas y ajusta el texto
'parametros: wsHoja es la hoja donde se ejecuta el metodo
'retorno: void
Public Function guardarComoCsv()

Dim provincia As String, periodo As String

provincia = Application.InputBox("Ingrese la provincia", ActiveWorkbook.name, Type:=2)

periodo = Application.InputBox("Ingrese el periodo", ActiveWorkbook.name, Type:=2)

ActiveWorkbook.SaveAs fileName:=ActiveWorkbook.Path & "\Resumen de detalle para analisis de tendencia de " & provincia & " - " & periodo, FileFormat:=xlCSV

End Function

'resumen: borro el contenido de un rango
'parametros: ws hoja donde se encuentra el rango de datos, inicioFila es la fila del primer elemento,
'            inicioColumna es la columna del primer elemento, finalFila es la fila del ultimo elemento,
'            finalColumna es la columna del ultimo elemento
'retorno: void
Public Function eliminarDatosRango(ws As Worksheet, inicioFila As Integer, inicioColumna As Integer, _
                                   finalFila As Integer, finalColumna As Integer)

 ws.Range(ws.Cells(inicioFila, inicioColumna), .ws.Cells(finalFila, finalColumna)).ClearContents

End Function

Public Function populateUniqueArray(ws As Worksheet, rango As Range)

Dim celda As Variant
Dim temporal As String, valoresArray() As String

For Each celda In rango
    
    If (celda <> "" And InStr(1, temporal, celda) = 0) Then
        
        temporal = temporal & celda & ";"
        
    End If
    
Next celda

If (Len(temporal) > 0) Then
    
    tmp = Left(tmp, Len(tmp) - 1)
    
End If

valoresArray = Split(tmp, "|")

End Function
