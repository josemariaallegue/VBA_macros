Attribute VB_Name = "Tendencias"
'resumen: le doy formato al archivo de resultado de tendencias
'parametros: void
'retorno: void
Public Sub tendencias_formato()

Dim columnas As Double, filas As Double
Dim hojas1 As String, hojas2 As String, hojas3 As String
Dim ws As Worksheet

Call preparacionInicio

hojas1 = "Universo 4;Interes 1;Interes 2"
hojas2 = "Universo 1;Universo 2"
hojas3 = "Universo 3;Interes 3"


For Each ws In ActiveWorkbook.Sheets
        
    columnas = columnaMax(ws)
    filas = filaMax(columnas, ws)
    
    'elimino la celdas y filas que estan en blanco
    If (InStr(1, hojas1, ws.name)) Then

        ws.Rows("2:2").Delete SHIFT:=xlUp
        ws.Columns("A:A").Delete SHIFT:=xlUp

    ElseIf (InStr(1, hojas2, ws.name)) Then

        ws.Range("A1").Delete SHIFT:=xlUp
        ws.Range("B2:Z2").Delete SHIFT:=xlUp
        
    ElseIf (InStr(1, hojas3, ws.name)) Then

        ws.Range("A2:Z2").Delete SHIFT:=xlUp

    End If
    
    'le doy formato a cada hoja
    If (ws.name <> "Log") Then
        
        Call cabiarNombreColumnas(ws, filas)
        Call formatos(ws)
        
    End If

Next ws

Call preparacionFinal

End Sub
'resumen: pinta las hojas en blanco, da formato de numero a las columnas y enmarca los dataframe
'parametros: ws es la hoja a la que se le va a dar formato
'retorno: void
Private Function formatos(ws As Worksheet)

    ws.Activate
    
    Cells.Select
    
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    
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
    Selection.Font.Bold = True
    
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
    
    Cells.Select
    Cells.EntireColumn.AutoFit
    Range("A1").Select
    
    Columns("A:F").Select
    Selection.Style = "Comma"
    Selection.NumberFormat = "#,##0"
    
    'da formato a las columnas de promedio
    If (ws.name = "Universo 2") Then
    
        Columns("D:D").Select
        Selection.Style = "Comma"
        Selection.NumberFormat = "#,##0.00"
        
    ElseIf (ws.name = "Interes 1" Or ws.name = "Interes 2") Then
        
        Columns("E:E").Select
        Selection.Style = "Comma"
        Selection.NumberFormat = "#,##0.00"
        
    End If
    
    
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
    
    Range("A1").Select
    
    If (ws.name = "Universo 1" Or ws.name = "Universo 2" Or ws.name = "Universo 3" Or _
    ws.name = "Universo 4" Or ws.name = "Interes 3") Then
        
        Selection.End(xlDown).Select
        Range(Selection, Selection.End(xlToRight)).Select
        
        With Selection.Font
            .ThemeColor = xlThemeColorDark1
            .TintAndShade = 0
        End With
        
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
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlMedium
        End With
        
        Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
        Selection.Font.Bold = True
        
        With Selection.Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .Color = 15773696
            .TintAndShade = 0
            .PatternTintAndShade = 0
        End With
        
    End If
    
    Range("A1").Select
    
End Function

Private Function cabiarNombreColumnas(ws As Worksheet, filas As Double)

With ws
    
    If (.name = "Universo 1") Then
        
        .Cells(1, 1).Value = "Categoria de población"
        .Cells(1, 2).Value = "Cantidad de beneficiarios"
        .Cells(filas, 1).Value = "Totales"
        ws.Cells(filas, 2).Formula = "=SUM(B2:B" & filas & ")"
        
    ElseIf (.name = "Universo 2") Then
    
        .Cells(1, 1).Value = "Categoria de población"
        .Cells(1, 2).Value = "Cantidad de beneficiarios"
        .Cells(1, 3).Value = "Cantidad de prestaciones"
        .Cells(1, 4).Value = "Promedio"
        .Cells(filas, 1).Value = "Totales"
        ws.Cells(filas, 2).Formula = "=SUM(B2:B" & filas & ")"
        ws.Cells(filas, 3).Formula = "=SUM(C2:C" & filas & ")"
        ws.Cells(filas, 4).Value = "-"
        
    ElseIf (.name = "Universo 3" Or .name = "Interes 3") Then
    
        .Cells(1, 1).Value = "Categoria de población"
        .Cells(1, 2).Value = "Cantidad de prestaciones consumidas por usuario"
        .Cells(1, 3).Value = "Cantidad de prestaciones"
        .Cells(1, 4).Value = "Total de prestaciones"
        .Cells(filas, 1).Value = "Totales"
        ws.Cells(filas, 2).Formula = "=SUM(B2:B" & filas & ")"
        ws.Cells(filas, 3).Formula = "=SUM(C2:C" & filas & ")"
        ws.Cells(filas, 4).Formula = "=SUM(D2:D" & filas & ")"
        
    ElseIf (.name = "Universo 4") Then
    
        .Cells(1, 1).Value = "Categoria de población"
        .Cells(1, 2).Value = "Codigo de prestación"
        .Cells(1, 3).Value = "Cantidad de prestaciones"
        .Cells(1, 4).Value = "Cantidad de beneficiarios"
        .Cells(filas, 1).Value = "Totales"
        ws.Cells(filas, 2).Value = "-"
        ws.Cells(filas, 3).Formula = "=SUM(C2:C" & filas & ")"
        ws.Cells(filas, 4).Formula = "=SUM(D2:D" & filas & ")"
        
    ElseIf (.name = "Universo 4") Then
    
        .Cells(1, 1).Value = "Categoria de población"
        .Cells(1, 2).Value = "Codigo de prestación"
        .Cells(1, 3).Value = "Cantidad de prestaciones"
        .Cells(1, 4).Value = "Cantidad de beneficiarios"
        .Cells(filas, 1).Value = "Totales"
        ws.Cells(filas, 2).Value = "-"
        ws.Cells(filas, 3).Formula = "=SUM(C2:C" & filas & ")"
        ws.Cells(filas, 4).Formula = "=SUM(D2:D" & filas & ")"
        
'    ElseIf (.name = "Interes 1") Then
'
'        .Cells(1, 1).Value = "Codigo de prestación"
'        .Cells(1, 2).Value = "Categoria de población"
'        .Cells(1, 4).Value = "Cantidad de beneficiarios"
'        .Cells(1, 3).Value = "Cantidad de prestaciones"
'        .Cells(filas, 1).Value = "Totales"
'        ws.Cells(filas, 2).Value = "-"
'        ws.Cells(filas, 3).Formula = "=SUM(C2:C" & filas & ")"
'        ws.Cells(filas, 4).Formula = "=SUM(D2:D" & filas & ")"
        
        End If

End With

End Function

