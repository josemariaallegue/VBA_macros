Attribute VB_Name = "csv"
'resumen: prepara un detalle de prestaciones, separa columnas claves y guarda todo en un .csv
'parametros: void
'retorno: void
Public Sub csv()

Call preparacionInicio

Call preparArchivo

Call copiarColumnas

'elimino las hojas de mas
Call eliminarHoja("Hoja1")
Call eliminarHoja("Database")

Call corregirPoblaciones

Call guardarComoCsv

Call preparacionFinal

ActiveWorkbook.Close (False)

End Sub

Private Function preparArchivo()

Dim wsDatabase As Worksheet
Dim filas As Double, columnas As Double
Dim tabla1 As ListObject


Set wsDatabase = ActiveWorkbook.Sheets("Database")
columnas = columnaMax(wsDatabase)
filas = filaMax(columnaMax(wsDatabase), wsDatabase)

Call crearHoja("CSV")

ActiveWorkbook.Sheets("CSV").Columns("A:J").NumberFormat = "@"
ActiveWorkbook.Sheets("CSV").Columns("F:G").NumberFormat = "d/m/yyyy"

'intento crear una tabal con los datos en la hoja, si ya esta creada solamente la asigno a una variable
On Error GoTo tablaexistente
Set tabla1 = wsDatabase.ListObjects.Add(xlSrcRange, Range(wsDatabase.Cells(1, 1), wsDatabase.Cells(filas, columnas)))

tabla1.name = "tabla1"

'le doy formato a la tabla
wsDatabase.Range(wsDatabase.Cells(1, 1), wsDatabase.Cells(1, columnas)).Interior.Color = RGB(0, 0, 0)
wsDatabase.Rows(1).Font.Color = RGB(255, 255, 255)

tablaexistente:

    Set tabla1 = wsDatabase.ListObjects("tabla1")
    Resume Next

End Function
'resumen: agrego la poblacion segun fertilidad y corrigo la poblacion segun sistemas
'parametros: void
'retorno: void
Private Function copiarColumnas()

Dim wsDatabase As Worksheet, wsCsv As Worksheet
Dim filasDatabase As Double, columnasDatabase As Double, filasCsv As Double
Dim i As Double
Dim codigoColumna As Integer, cuieColumna As Integer, fPrestacionColumnas As Integer
Dim claveColumna As Integer, sexoColumna As Integer, fNacimientoColumna As Integer
Dim precioColumna As Integer, edadColumna As Integer, categoriaColumna As Integer
Dim documentoColumna As Integer


Set wsDatabase = ActiveWorkbook.Sheets("Database")
columnasDatabase = columnaMax(wsDatabase)
filasDatabase = filaMax(columnasDatabase, wsDatabase)
Set wsCsv = ActiveWorkbook.Sheets("CSV")


For i = 1 To columnasDatabase

    With wsDatabase
        
        If (.Cells(1, i).Value = "CODIGO_PRESTACION") Then
        
            codigoColumna = i
            
        ElseIf (.Cells(1, i).Value = "CUIE_EFECTOR") Then
        
            cuieColumna = i
            
        ElseIf (.Cells(1, i).Value = "FECHA_PRESTACION") Then
        
            fPrestacionColumna = i
            
        ElseIf (.Cells(1, i).Value = "CLAVE_BENEFICIARIO") Then
        
            claveColumna = i
            
        ElseIf (.Cells(1, i).Value = "SEXO") Then
        
            sexoColumna = i
            
        ElseIf (.Cells(1, i).Value = "FECHA_DE_NACIMIENTO") Then
        
            fNacimientoColumna = i
            
        ElseIf (LCase(.Cells(1, i).Value) = "precio" Or LCase(.Cells(1, i).Value) = "valor_unitario" Or _
        .Cells(1, i).Value = "PRECIO_UNITARIO" Or .Cells(1, i).Value = "MONTO") Then
        
            precioColumna = i
            
        ElseIf (.Cells(1, i).Value = "AÑOS_EN_DIA_PRESTACION") Then
        
            edadColumna = i
            
        ElseIf (.Cells(1, i).Value = "CATEGORIA_LIQ") Then
        
            categoriaColumna = i
            
        ElseIf (.Cells(1, i).Value = "BENEF_NRO_DOCUMENTO") Then
            
            documentoColumna = i
            
        End If
        
    End With

Next i


With wsDatabase.ListObjects("Tabla1")

    .ListColumns(claveColumna).DataBodyRange.Copy
    wsCsv.Cells(1, 1).PasteSpecial Paste:=xlPasteValues
    
    .ListColumns(documentoColumna).DataBodyRange.Copy
    wsCsv.Cells(1, 2).PasteSpecial Paste:=xlPasteValues
    
    .ListColumns(cuieColumna).DataBodyRange.Copy
    wsCsv.Cells(1, 3).PasteSpecial Paste:=xlPasteValues
    
    .ListColumns(codigoColumna).DataBodyRange.Copy
    wsCsv.Cells(1, 4).PasteSpecial Paste:=xlPasteValues
    
    .ListColumns(sexoColumna).DataBodyRange.Copy
    wsCsv.Cells(1, 5).PasteSpecial Paste:=xlPasteValues
    
    .ListColumns(fNacimientoColumna).DataBodyRange.Copy
    wsCsv.Cells(1, 6).PasteSpecial Paste:=xlPasteValues
    
    .ListColumns(fPrestacionColumna).DataBodyRange.Copy
    wsCsv.Cells(1, 7).PasteSpecial Paste:=xlPasteValues
    
    .ListColumns(edadColumna).DataBodyRange.Copy
    wsCsv.Cells(1, 8).PasteSpecial Paste:=xlPasteValues
    
    .ListColumns(precioColumna).DataBodyRange.Copy
    wsCsv.Cells(1, 9).PasteSpecial Paste:=xlPasteValues

    .ListColumns(categoriaColumna).DataBodyRange.Copy
    wsCsv.Cells(1, 10).PasteSpecial Paste:=xlPasteValues
    
End With

End Function
'resumen: agrego la poblacion segun fertilidad y corrigo la poblacion segun sistemas
'parametros: void
'retorno: void
Private Function corregirPoblaciones()

Dim i As Double
Dim filas As Double, columnas As Double
Dim categoriaColumna As Integer, sexoColumna As Integer, edadColumna As Integer, fertilColumna As Integer

columnas = columnaMax(ActiveSheet)
filas = filaMax(columnas, ActiveSheet)
categoriaColumna = 9
fertilColumna = 10
sexoColumna = 4
edadColumna = 7


For i = 1 To filas

    With ActiveSheet
        
        'corrigo la columna de poblacion segun sistemas
        With .Cells(i, categoriaColumna)
            
            Select Case .Value
            
                Case "Adolecentes 10-19"
                    .Value = "10-19"
            
                Case "Hombres 20-64", "Mujeres 20-64"
                    .Value = "20-64"
                    
                Case "Niños 0-5"
                    .Value = "0-5"
    
                Case "Niños 6-9"
                    .Value = "6-9"
                    
                Case ""
                    .Value = "No categorizable"
                    
                Case "NO CATEGORIZABLE"
                    .Value = "No categorizable"
                                
            End Select
            
        End With
        
        'agrego la poblacion segun fertilidad
        If (LCase(.Cells(i, categoriaColumna).Value) = "no categorizable") Then
            
            .Cells(i, fertilColumna).Value = "No categorizable"
        
        ElseIf ((LCase(.Cells(i, sexoColumna).Value) = "f" Or LCase(.Cells(i, sexoColumna).Value) = "femenino") And _
        (.Cells(i, edadColumna).Value >= 12 And .Cells(i, edadColumna).Value <= 55)) Then
        
            .Cells(i, fertilColumna).Value = "Mujer 12-55"
            
        ElseIf ((LCase(.Cells(i, sexoColumna).Value) = "m" Or LCase(.Cells(i, sexoColumna).Value) = "masculino") And _
        (.Cells(i, edadColumna).Value >= 12 And .Cells(i, edadColumna).Value <= 55)) Then
        
            .Cells(i, fertilColumna).Value = "Hombre 12-55"
            
        ElseIf (.Cells(i, edadColumna).Value <= 11) Then
        
            .Cells(i, fertilColumna).Value = "0-11"
            
        ElseIf (.Cells(i, edadColumna).Value >= 56 And .Cells(i, edadColumna).Value <= 64) Then
        
            .Cells(i, fertilColumna).Value = "56-64"
            
        Else
            
            .Cells(i, fertilColumna).Value = "No categorizable"
            
        End If

    End With
       
Next i

End Function
