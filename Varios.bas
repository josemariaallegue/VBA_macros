Attribute VB_Name = "Varios"
Option Explicit

'resumen: compara la cantidad de registros entre 2 solapas seleccionadas
'parametros: void
'retorno: void
Public Sub comparar_cantidad_registros()

usfmCompararRegistros.Show

End Sub

'resumen: compara los valores de los registros entre 2 solapas seleccionadas
'parametros: void
'retorno: void
Public Sub comparar_valores_registros()

usfmCompararValores.Show

End Sub

'resumen: unifica un rango de celdas seleccionadas (separadas por ;) y las colaca en una celda seleccionada
'parametros: void
'retorno: void
Public Sub unificar_celdas()

Dim rango1, rango2 As Range
Dim celda As Variant
Dim texto As String
Dim flag As Boolean

flag = False
Set rango1 = Application.InputBox("Seleccione un rango", "Unificar celdas", Type:=8)
Set rango2 = Application.InputBox("Seleccione el destino", "Unificar celdas", Type:=8)

For Each celda In rango1
    
    If (flag = False) Then
        
        texto = celda.Value
        flag = True
    
    Else
    
        texto = texto & ";" & celda.Value
    
    End If

Next celda

rango2.Value = texto

End Sub

'resumen: crea la solapa "Resumen" para el metodo "comparar_cantidad_registros" con los cuadros correspondientes
'parametros: las 2 hojas que se compararon, un array con el resumen con los resultados y la cantidad de columnas de la hoja1
'retorno: void
Public Function cuadrosComparacionCantidadRegistros(ByVal hoja1 As Worksheet, ByVal hoja2 As Worksheet, ByVal resumen As Variant, ByVal columnaHoja1 As Long)

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

    Exit Function
    
'excepcion por si hay un error al crear la hoja
erroralcrearhoja:
Sheets(Sheets.Count).name = "Resumen" & Sheets.Count - 1
Resume Next

End Function

'resumen: da formato al cuadro de resumen del metodo "Comparar_cantidad_registros"
'parametros: void
'retorno: void
Public Function formatosCompararCantidadRegistros()

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
    
End Function

'resumen:pide que se seleccione un rango y se para el contenido por "-" y luego copia el contenido en la fila 15
'        MODIFICAR!!!!!!!!!!!!!!!!
'parametros: void
'retorno: void
Public Sub separar_codigos()
Attribute separar_codigos.VB_ProcData.VB_Invoke_Func = "r\n14"

Dim i As Integer
Dim rango As Range
Dim texto As Variant

Set rango = Application.InputBox("Seleccione un rango", "Separacion de codigos", Type:=8)

texto = Split(rango.Value, "-")

For i = 0 To UBound(texto)

    If (i > 0) Then
        
        ActiveSheet.Cells(rango.Row + i, 15).Value = Left(texto(0), Len(texto(0)) - 3) & texto(i)
        
    Else
    
        ActiveSheet.Cells(rango.Row + i, 15).Value = texto(i)
        
    End If

Next i

End Sub

'resumen: desprotege la hoja activa a travez de un ataque de fuerza bruta
'parametros: void
'retorno: void
Public Sub sacar_claves()

Dim Contraseña As String
Dim i01 As Integer, i02 As Integer, i03 As Integer
Dim i04 As Integer, i05 As Integer, i06 As Integer
Dim i07 As Integer, i08 As Integer, i09 As Integer
Dim i10 As Integer, i11 As Integer, i12 As Integer

On Error Resume Next
For i01 = 65 To 66: For i02 = 65 To 66: For i03 = 65 To 66
For i04 = 65 To 66: For i05 = 65 To 66: For i06 = 65 To 66
For i07 = 65 To 66: For i08 = 65 To 66: For i09 = 65 To 66
For i10 = 65 To 66: For i11 = 65 To 66: For i12 = 32 To 126

    Contraseña = Chr(i01) & Chr(i02) & Chr(i03) & Chr(i04) & Chr(i05) & Chr(i06) & Chr(i07) & Chr(i08) & Chr(i09) & Chr(i10) & Chr(i11) & Chr(i12)
    ActiveSheet.Unprotect Contraseña
    
    If ActiveSheet.ProtectContents = False Then
        Debug.Print Contraseña
        MsgBox "La contraseña encontrada es " & Contraseña
        Exit Sub
    End If
    
Next: Next: Next
Next: Next: Next
Next: Next: Next
Next: Next: Next

End Sub

Public Sub limpiar()

Dim i As Double, j As Double, filas As Double, columnas As Double


columnas = columnaMax(ActiveSheet)
filas = filaMax(columnas, ActiveSheet)

For i = 1 To filas

    For j = 1 To columnas
        
        If (ActiveSheet.Cells(i, j).Value = "") Then
            
            ActiveSheet.Cells(i, j).Value = ""
            
        End If
        
    Next j
    
Next i

End Sub
