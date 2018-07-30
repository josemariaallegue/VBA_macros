Attribute VB_Name = "Privados"
Option Explicit
Option Private Module

Public Sub preparacionInicio()

'preparacion previa al incio de funcion
Application.EnableEvents = False
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual

End Sub

Public Sub preparacionFinal()

'preparacion para el cierre de la funcion
Application.EnableEvents = True
Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic

End Sub

Public Sub finalizarTestVelocidad(ByVal auxAhora As Date)

Dim tiempoSegundos, tiempoMinutos As Long

tiempoSegundos = DateDiff("s", auxAhora, Now)
tiempoMinutos = tiempoSegundos / 60

If (tiempoSegundos < 60) Then
    
    Debug.Print ("Termino en " & tiempoSegundos & " segundos")
    
Else

    Debug.Print ("Termino en " & tiempoMinutos & " minutos")

End If

End Sub


Public Sub listarArchivos()

Dim auxArch As Workbook

For Each auxArch In Workbooks
    
    Debug.Print auxArch.Name

Next

End Sub

'crea la solapa "Resumen" para el metodo "compararCantidadRegistros" con los cuadros correspondientes
Sub cuadrosComparacionCantidadRegistros(ByVal hoja1 As Worksheet, ByVal hoja2 As Worksheet, ByVal resumen As Variant, ByVal columnaHoja1 As Long)

Dim i, j As Integer

'creo una sola nueva y la nombro
On Error GoTo erroralcrearhoja

    Sheets.Add After:=Sheets(Sheets.Count)
    Sheets(Sheets.Count).Name = "Resumen"
    
    'pongo los valores a las columnas de los datos
    Sheets(Sheets.Count).Cells(1, 1).Value = "Columna"
    Sheets(Sheets.Count).Cells(1, 2).Value = "Registros en " & hoja1.Name
    Sheets(Sheets.Count).Cells(1, 3).Value = "Registros en " & hoja2.Name
    Sheets(Sheets.Count).Cells(1, 4).Value = "Diferencia entre las hojas"
    
    'recorro la variable "resumen" para darle valor al cuadro
    For i = 1 To columnaHoja1
        
        For j = 1 To 4
            
            Sheets(Sheets.Count).Cells(i + 1, 1).Value = resumen(i, 1)
            Sheets(Sheets.Count).Cells(i + 1, 2).Value = resumen(i, 2)
            Sheets(Sheets.Count).Cells(i + 1, 3).Value = resumen(i, 3)
            Sheets(Sheets.Count).Cells(i + 1, 4).Value = resumen(i, 4)
            
            
        Next j
        
    Next i

    Exit Sub
    
'excepcion por si hay un error al crear la hoja
erroralcrearhoja:
Sheets(Sheets.Count).Name = "Resumen" & Sheets.Count - 1
Resume Next

End Sub

Sub cuadrosMuestraPadron(ByVal cuieArray As Variant, ByVal cantidadMuestraArray As Variant, ByVal provinciaArray As Variant, ByVal muestraArray As Variant, _
ByVal n As Variant, ByVal noElegiblesArray As Variant, ByVal validosXcuie As Variant, ByVal contador As Integer)


Dim i, j, l As Integer

i = 2

'creo una sola nueva y la nombro
On Error GoTo erroralcrearhoja

    Sheets.Add After:=Sheets(Sheets.Count)
    Sheets(Sheets.Count).Name = "Resumen"
    
    'pongo los valores a las columnas de los datos
    Sheets(Sheets.Count).Cells(1, 1).Value = "Provincia ID"
    Sheets(Sheets.Count).Cells(1, 2).Value = "N"
    Sheets(Sheets.Count).Cells(1, 3).Value = "Cuie"
    Sheets(Sheets.Count).Cells(1, 4).Value = "Casos validos por efector"
    Sheets(Sheets.Count).Cells(1, 5).Value = "Cantidades determinadas por calculo"
    Sheets(Sheets.Count).Cells(1, 6).Value = "Cantidades tomadas"
    Sheets(Sheets.Count).Cells(1, 7).Value = "Diferencias"
    Sheets(Sheets.Count).Cells(1, 8).Value = "Codigos no elegibles por efector"
    Sheets(Sheets.Count).Cells(1, 10).Value = "Codigos no elegibles tomados"
    
    'coloco los valores obtenidos del analisis
    Sheets(Sheets.Count).Cells(2, 10).Value = contador
    
    'recorro los arrays para ver sus contenidos
    For j = 1 To largoArray(cuieArray)
        
        If (cuieArray(j) <> "") Then
        
            Sheets(Sheets.Count).Cells(j + 1, 1).Value = provinciaArray(j)
            Sheets(Sheets.Count).Cells(j + 1, 2).Value = n(j)
            Sheets(Sheets.Count).Cells(j + 1, 3).Value = cuieArray(j)
            Sheets(Sheets.Count).Cells(j + 1, 4).Value = validosXcuie(j)
            Sheets(Sheets.Count).Cells(j + 1, 5).Value = cantidadMuestraArray(j)
            Sheets(Sheets.Count).Cells(j + 1, 6).Value = muestraArray(j)
            Sheets(Sheets.Count).Cells(j + 1, 7).Value = muestraArray(j) - cantidadMuestraArray(j)
            Sheets(Sheets.Count).Cells(j + 1, 8).Value = noElegiblesArray(j)
        
        End If
        
        
    Next j
    
    'verifico que ninguna de las muestras sea menor a 5 y si lo es pinto la celda
    Do Until Sheets(Sheets.Count).Cells(i, 5).Value = ""
        
        If (Sheets(Sheets.Count).Cells(i, 5).Value < 5 And Sheets(Sheets.Count).Cells(i, 5).Value <> "") Then
            
            Sheets(Sheets.Count).Cells(i, 5).Interior.Color = RGB(255, 255, 0)
            
        End If
        
        i = i + 1
    
    Loop
    
    'le doy formato a la solapa
    Cells.Select
        With Selection
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .WrapText = True
            .Orientation = 0
            .AddIndent = False
            .IndentLevel = 0
            .ShrinkToFit = False
            .ReadingOrder = xlContext
            .MergeCells = False
        End With
        
    Rows("2:200").Select
    Selection.NumberFormat = "#,##0"
    
    Exit Sub
    
erroralcrearhoja:
Sheets(Sheets.Count).Name = "Resumen" & Sheets.Count - 1
Resume Next

End Sub

Sub cuadrosMuestraPagos(ByVal cuieArray As Variant, ByVal n As Integer, ByVal contador1, ByVal contador2 As Integer, ByVal CONTADOR3 As Integer)

Dim i, j, l As Integer
Dim diferenciaAbs As Integer

i = 2

'creo una sola nueva y la nombro
On Error GoTo erroralcrearhoja

    Sheets.Add After:=Sheets(Sheets.Count)
    Sheets(Sheets.Count).Name = "Resumen"

    'pongo los valores a las columnas de los datos
    Sheets(Sheets.Count).Cells(1, 1).Value = "Efectores"
    Sheets(Sheets.Count).Cells(1, 2).Value = "Casos validos por efector"
    Sheets(Sheets.Count).Cells(1, 3).Value = "Cantidades determinadas por calculo"
    Sheets(Sheets.Count).Cells(1, 4).Value = "Cantidades tomadas"
    Sheets(Sheets.Count).Cells(1, 5).Value = "Codigos no elegibles por efector"
    Sheets(Sheets.Count).Cells(1, 6).Value = "Diferencias"
    Sheets(Sheets.Count).Cells(1, 8).Value = "Cantidad de efectores"
    Sheets(Sheets.Count).Cells(1, 9).Value = "Sumatoria cantidad determinada por calculo"
    Sheets(Sheets.Count).Cells(1, 10).Value = "Casos realmente tomados (totalidad)"
    Sheets(Sheets.Count).Cells(1, 11).Value = "Diferencia (totalidad)"
    Sheets(Sheets.Count).Cells(1, 12).Value = "Codigos no elegibles tomados"

    
    
    'recorro el array cuieArray para ver sus contenidos
    For j = 1 To 12
        
        If (cuieArray(j, 1) <> "") Then
        
            For l = 1 To 5
        
                Sheets(Sheets.Count).Cells(j + 1, l).Value = cuieArray(j, l)
            
            Next l
            
            'Calculo la diferencia entre las cantidades determinadas por calculo y las cantidades tomadas
            Sheets(Sheets.Count).Cells(j + 1, 6).Value = cuieArray(j, 4) - cuieArray(j, 3)
            
        End If
        
    Next j
    
    'verifico que ninguna de las muestras sea menor a 5 y si lo es pinto la celda
    Do Until i = 14
        
        If (Sheets(Sheets.Count).Cells(i, 4).Value < 5 And Sheets(Sheets.Count).Cells(i, 4).Value <> "") Then
            
            Sheets(Sheets.Count).Cells(i, 4).Interior.Color = RGB(255, 255, 0)
            
        End If
        
        diferenciaAbs = Abs(Sheets(Sheets.Count).Cells(i, 6).Value) + diferenciaAbs
        
        i = i + 1
    
    Loop

    'coloco los valores obtenidos del analisis
    Sheets(Sheets.Count).Cells(2, 8).Value = contador1
    Sheets(Sheets.Count).Cells(2, 9).Value = Application.Sum(Range(Cells(2, 3), Cells(15, 3)))
    Sheets(Sheets.Count).Cells(2, 10).Value = Application.Sum(Range(Cells(2, 4), Cells(15, 4)))
    Sheets(Sheets.Count).Cells(2, 11).Value = diferenciaAbs
    Sheets(Sheets.Count).Cells(2, 12).Value = contador2
    
    
    'le doy formato a la solapa
    Cells.Select
        With Selection
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .WrapText = True
            .Orientation = 0
            .AddIndent = False
            .IndentLevel = 0
            .ShrinkToFit = False
            .ReadingOrder = xlContext
            .MergeCells = False
        End With
        
    Rows("2:20").Select
    Selection.NumberFormat = "#,##0"
    
    Exit Sub
    
erroralcrearhoja:
Sheets(Sheets.Count).Name = "Resumen" & Sheets.Count - 1
Resume Next
    
End Sub

