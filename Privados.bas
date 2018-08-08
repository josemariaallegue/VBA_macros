Attribute VB_Name = "Privados"
Option Explicit
Option Private Module

'resumen: preparacion previa al incio de funcion
'parametros: void
'retorno: void
Public Sub preparacionInicio()

Application.EnableEvents = False
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual

End Sub

'resumen: preparacion para el cierre de la funcion
'parametros: void
'retorno: void
Public Sub preparacionFinal()

Application.EnableEvents = True
Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic

End Sub

'resumen: imprime por pantalla el tiempo que tardo en ejecutarse un progama
'parametros: variable que contiene en que momento se emepezo la ejecucion
'retorno: void
Public Sub finalizarTestVelocidad(ByVal auxAhora As Date)

Dim tiempoSegundos, tiempoMinutos As Long

tiempoSegundos = DateDiff("s", auxAhora, Now)
tiempoMinutos = tiempoSegundos / 60

If (tiempoSegundos <= 60) Then
    
    Debug.Print ("Termino en " & tiempoSegundos & " segundos")
    
Else

    Debug.Print ("Termino en " & tiempoMinutos & " minutos")

End If

End Sub

'resumen: imprime por pantalla el nombre de todos los archivos abiertos
'parametros: void
'retorno: void
Public Sub listarArchivos()

Dim auxArch As Workbook

For Each auxArch In Workbooks
    
    Debug.Print auxArch.name

Next

End Sub

'resumen: crea la solapa "Resumen" para el metodo "comparar_cantidad_registros" con los cuadros correspondientes
'parametros: las 2 hojas que se compararon, un array con el resumen con los resultados y la cantidad de columnas de la hoja1
'retorno: void
Sub cuadrosComparacionCantidadRegistros(ByVal hoja1 As Worksheet, ByVal hoja2 As Worksheet, ByVal resumen As Variant, ByVal columnaHoja1 As Long)

Dim i, j, k As Integer

k = 2

'creo una sola nueva y la nombro
On Error GoTo erroralcrearhoja

    Sheets.Add After:=Sheets(Sheets.Count)
    Sheets(Sheets.Count).name = "Resumen"
    
    'pongo los valores a las columnas de los datos
    Sheets(Sheets.Count).Cells(1, 1).Value = "Columna"
    Sheets(Sheets.Count).Cells(1, 2).Value = "Registros en " & hoja1.name
    Sheets(Sheets.Count).Cells(1, 3).Value = "Registros en " & hoja2.name
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
    
    'le doy formato a la hoja ahora porque sino el siguiente loop es cancelado
    Call formatosCompararCantidadRegistros
    
    'si se encuentra una diferencia entre las cantidades se pinta la celda de rojo
    Do Until Sheets(Sheets.Count).Cells(k, 4).Value = ""
        
        If (Sheets(Sheets.Count).Cells(k, 4).Value > 0 Or Sheets(Sheets.Count).Cells(k, 4) < 0) Then
            
            Sheets(Sheets.Count).Cells(k, 4).Interior.Color = RGB(255, 255, 0)
            
        End If
        
        k = k + 1
    
    Loop

    Exit Sub
    
'excepcion por si hay un error al crear la hoja
erroralcrearhoja:
Sheets(Sheets.Count).name = "Resumen" & Sheets.Count - 1
Resume Next

End Sub

'resumen: crea la hoja "Resumen" para el metodo "revision_muestra_padron" con los cuadros correspondientes
'parametros: un array con los cuie, un array con la cantidad de muestra por calculo para cada cuie, un array con el id de las provincias
'            un array con la muestra realmente tomada para cada cuie, la n para cada provincia, la cantidad de codigos no elegibles para cada cuie,
'            la cantidad de casos validos para cada cuie y la totalidad de codigos no elegibles
'retorno: void
Sub cuadrosMuestraPadron(ByVal cuieArray As Variant, ByVal cantidadMuestraArray As Variant, ByVal provinciaArray As Variant, ByVal muestraArray As Variant, _
ByVal n As Variant, ByVal noElegiblesArray As Variant, ByVal validosXcuie As Variant, ByVal contador As Integer)


Dim i, j, l As Integer

i = 2

'creo una sola nueva y la nombro
On Error GoTo erroralcrearhoja

    Sheets.Add After:=Sheets(Sheets.Count)
    Sheets(Sheets.Count).name = "Resumen"
    
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
    
    'le doy formato a la hoja
    Call formatosRevisionMuestraPadron
    
    'verifico que ninguna de las muestras sea menor a 5 y si lo es pinto la celda
    Do Until Sheets(Sheets.Count).Cells(i, 5).Value = ""
        
        If (Sheets(Sheets.Count).Cells(i, 5).Value < 5 And Sheets(Sheets.Count).Cells(i, 5).Value <> "") Then
            
            Sheets(Sheets.Count).Cells(i, 5).Interior.Color = RGB(255, 255, 0)
            
        End If
        
        i = i + 1
    
    Loop

    Exit Sub
    
erroralcrearhoja:
Sheets(Sheets.Count).name = "Resumen" & Sheets.Count - 1
Resume Next

End Sub

'resumen: crea la hoja "Resumen" para el metodo "revision_muestra_pagos" con los cuadros correspondientes
'parametros: un array con las datos de la muestra, la n para la provincia, la cantidad de efectores, la cantidad de codigos no elegibles tomadas
'retorno: void
Sub cuadrosMuestraPagos(ByVal cuieArray As Variant, ByVal n As Integer, ByVal contador1, ByVal contador2 As Integer)

Dim i, j, l As Integer
Dim diferenciaAbs As Integer

i = 2

'creo una sola nueva y la nombro
On Error GoTo erroralcrearhoja

    Sheets.Add After:=Sheets(Sheets.Count)
    Sheets(Sheets.Count).name = "Resumen"

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
    
    'coloco los valores obtenidos del analisis
    Sheets(Sheets.Count).Cells(2, 8).Value = contador1
    Sheets(Sheets.Count).Cells(2, 9).Value = Application.Sum(Range(Cells(2, 3), Cells(15, 3)))
    Sheets(Sheets.Count).Cells(2, 10).Value = Application.Sum(Range(Cells(2, 4), Cells(15, 4)))
    Sheets(Sheets.Count).Cells(2, 12).Value = contador2
    
    'le doy formato a la solapa
    Call formatosRevisionMuestraPagos
    
    'verifico que ninguna de las muestras sea menor a 5 y si lo es pinto la celda
    Do Until i = 14
        
        If (Sheets(Sheets.Count).Cells(i, 4).Value < 5 And Sheets(Sheets.Count).Cells(i, 4).Value <> "") Then
            
            Sheets(Sheets.Count).Cells(i, 4).Interior.Color = RGB(255, 255, 0)
            
        End If
        
        diferenciaAbs = Abs(Sheets(Sheets.Count).Cells(i, 6).Value) + diferenciaAbs
        
        i = i + 1
    
    Loop
    
    'este parte tiene que estar aca porque si se pone con las demas el valor de la variable esta en 0'
    'y si quiero darle valor antes tendria que hacer una iteracion de mas
    Sheets(Sheets.Count).Cells(2, 11).Value = diferenciaAbs
    Exit Sub
    
erroralcrearhoja:
Sheets(Sheets.Count).name = "Resumen" & Sheets.Count - 1
Resume Next
    
End Sub

'resumen: da formato al cuadro de resumen del metodo "Comparar_cantidad_registros"
'parametros: void
'retorno: void
Sub formatosCompararCantidadRegistros()

    Cells.Select
    With Selection.Interior
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    Cells.EntireColumn.AutoFit
    With Selection
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Range("A1:D1").Select
    With Selection.Font
        .Color = -1003520
        .TintAndShade = 0
    End With
    With Selection.Font
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
    End With
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 5287936
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 15773696
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    Selection.Font.Bold = True
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Range("A1:D1").Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    
    Rows("2:50").Select
    Selection.NumberFormat = "#,##0"
    Range("A1").Select
    
End Sub

'resumen: da formato al cuadro de resumen del metodo "revision_muestra_pagos"
'parametros: void
'retorno: void'
Sub formatosRevisionMuestraPagos()

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
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    Range("A1").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Range("A1:F1").Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    With Selection.Font
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
    End With
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 15773696
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    Selection.Font.Bold = True
    Range("H1:L2").Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Range("H1:L1").Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    With Selection.Font
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
    End With
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 15773696
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    Selection.Font.Bold = True
    Rows("2:44").Select
    Selection.NumberFormat = "#,##0"
    Range("A1").Select
End Sub

'resumen: da formato al cuadro de resumen del metodo "revision_muestra_padron"
'parametros: void
'retorno: void'
Sub formatosRevisionMuestraPadron()

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
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    Range("A1").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Range("A1").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    With Selection.Font
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
    End With
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 15773696
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    Selection.Font.Bold = True
    Range("J1:J2").Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    Range("J1").Select
    With Selection.Font
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
    End With
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 15773696
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    Selection.Font.Bold = True
    Rows("2:200").Select
    Selection.NumberFormat = "#,##0"
    Range("A1").Select
End Sub


'resumen: determina la cantidad de columnas de la fila 1 de la hoja activa
'parametros: void
'retorno: la cantidad de columnas en un entero
Function columnaMax(ByVal hoja As Worksheet) As Integer

columnaMax = hoja.Cells(1, Columns.Count).End(xlToLeft).Column

End Function

'resumen: determina la cantidad de filas maxi para la hoja activa
'parametros: la cantidad de columnas de la hoja activa
'retorno: la cantidad de filas maxima en un entero
Function filaMax(ByVal cantidadColumnas As Long, ByVal hoja As Worksheet) As Integer

Dim i, filas As Integer

    For i = 1 To cantidadColumnas
    
    If (filas < hoja.Cells(Rows.Count, i).End(xlUp).Row) Then
        
        filas = hoja.Cells(Rows.Count, i).End(xlUp).Row
        
    End If

Next i

filaMax = filas
    
End Function
