Attribute VB_Name = "Nomencladores"
Option Explicit

'resumen: arma un resumen indidividual para cada hoja de poblacion de los nomencladores
'parametros: void
'retorno: void
Public Sub nomencladores()

Dim hoja As Worksheet
Dim nombreAuxiliar As String

Call preparacionInicio

For Each hoja In ActiveWorkbook.Worksheets
    
    nombreAuxiliar = Replace(LCase(hoja.name), " ", "")
    
    hoja.Activate
    
    If (nombreAuxiliar = "emb.part.puerp" Or nombreAuxiliar = "emb.parto.puerp" Or nombreAuxiliar = "0a5años" Or _
    nombreAuxiliar = "niños6a9años" Or nombreAuxiliar = "adolescente" Or nombreAuxiliar = "adolescentes" Or _
    nombreAuxiliar = "adultos" Or nombreAuxiliar = "adulto" Or nombreAuxiliar = "catastroficas") Then
        
        Call resumenGeneral(hoja)
    
    End If
    
Next hoja

Call resumenAnexo

Call preparacionFinal

End Sub

'resumen: arma un resumen indidividual para la hoja activa del nomencladores
'parametros: void
'retorno: void
Public Sub nomenclador_hoja_activa()

Call preparacionInicio

If (ActiveSheet.name = "Anexos") Then
    
    Call resumenAnexo
    
Else

    Call resumenGeneral(ActiveSheet)
    
End If

Call preparacionFinal

End Sub

'resumen: arma un cuadro con todos los codigos de los resumens del nomenclador
'parametros: void
'retorno: void
Public Sub nomenclador_unificado()

Dim filasResumen As Double, filasUnificado As Double, i As Double
Dim columnas As Integer
Dim wsHoja As Worksheet, wsUnificado As Worksheet
Dim nombre As String, provincia As String, año As String, poblacionAuxiliar As String

Call preparacionInicio

columnas = 4
nombre = "Unificado"
provincia = Application.InputBox("Ingrese la provincia", "Nomencladores")
año = Application.InputBox("Ingrese el año del nomenclador", "Nomencladores")

Call crearHoja(nombre)

Set wsUnificado = ActiveWorkbook.Sheets(nombre)

Call limpiarHoja(wsUnificado)

'nombre y formato a columnas
With wsUnificado
    
    .Columns("C:C").NumberFormat = "@"
    .Columns("D:D").Style = "Currency"
    .Columns("E:E").NumberFormat = "@"
    .Cells(1, 1).Value = "Codigos"
    .Cells(1, 2).Value = "Nombres"
    .Cells(1, 3).Value = "Poblacion"
    .Cells(1, 4).Value = "Precio"
    .Cells(1, 5).Value = "Provincia"
    .Cells(1, 6).Value = "Año"
    
End With

'recorro las hojas de resumen y copio los datos a la hoja de unificado
For Each wsHoja In ActiveWorkbook.Worksheets
    
    If (InStr(1, wsHoja.name, "Resumen") > 0) Then
        
        wsHoja.Activate
        
        filasResumen = filaMax(columnas, wsHoja)
        filasUnificado = filaMax(columnas, wsUnificado)

        wsHoja.Range(wsHoja.Cells(2, 1), wsHoja.Cells(filasResumen, 4)).Select
        Selection.Copy
        
        wsUnificado.Activate
        wsUnificado.Range("A" & filasUnificado + 1).Select
        wsUnificado.Paste
        Application.CutCopyMode = False
        
    End If
    
Next wsHoja

'doy formato al cuadro

filasUnificado = filaMax(columnas, wsUnificado)

For i = 2 To filasUnificado
    
    With wsUnificado
        
        poblacionAuxiliar = Replace(LCase(.Cells(i, 3).Value), " ", "")
        
        Select Case poblacionAuxiliar
        
            Case Is = "0a5años", "0-5años", "niñosde0a5"
                .Cells(i, 3).Value = "0 - 5 años"
        
            Case Is = "6a9años", "niños6a9años", "6-9años", "niñosde6a9"
                .Cells(i, 3).Value = "6 - 9 años"
                
            Case Is = "adolescente", "10-19años"
                .Cells(i, 3).Value = "Adolescentes"
                
            Case Is = "20-64años", "adultos"
                .Cells(i, 3).Value = "Adultos"
                            
            Case Is = "emb.parto.puerp", "embarazonormal", "embarazoriesgoso"
                .Cells(i, 3).Value = "Embarazos"
    
        End Select
        
        'pinto los codigos que pueden tener cualquier diagnostico
        If ((Len(.Cells(i, 1).Value) = 6 And LCase(Left(.Cells(i, 1).Value, 2)) <> "it") Or .Cells(i, 3).Value = "Todas") Then
            
            .Cells(i, 1).Interior.Color = RGB(255, 255, 0)
            
        End If
        
        .Cells(i, 5).Value = provincia
        .Cells(i, 6).Value = año
   
    End With
    
Next i

Call formatos(wsUnificado)

'ordeno por los codigos pintados por amarillo
wsUnificado.AutoFilter.Sort.SortFields.Clear
wsUnificado.AutoFilter.Sort.SortFields.Add(Range( _
    "A1:A" & filasUnificado), xlSortOnCellColor, xlAscending, , xlSortNormal).SortOnValue.Color _
    = RGB(255, 255, 0)
    
With wsUnificado.AutoFilter.Sort
    .Header = xlYes
    .MatchCase = False
    .Orientation = xlTopToBottom
    .SortMethod = xlPinYin
    .Apply
End With

Call preparacionFinal

End Sub

'resumen: recorre la solapa de anexo del nomenclador y toma los codigos, descripciones y precios, se los pasa a otra
'         funcion y arma los cuadros depurados y resumidos
'parametros: void
'retorno: void
Public Function resumenAnexo()

Dim i As Double, filas As Double, columnas As Double, m As Double
Dim j As Integer, k As Integer, codigoColumna As Integer, nombreColumna As Integer, precioColumna As Integer, poblacionFila As Integer
Dim codigosFila As Integer
Dim nombreHoja As String, codigoAuxiliar As String
Dim codigo() As Variant, nombre() As Variant, precio() As Variant, poblacion() As Variant
Dim wsAnexos As Worksheet
Dim flag As Boolean

ReDim codigo(1 To 750)
ReDim nombre(1 To 750)
ReDim precio(1 To 750)
ReDim poblacion(1 To 750)

Set wsAnexos = ActiveWorkbook.Sheets("Anexos")
columnas = 15
filas = filaMax(columnas, wsAnexos)
nombreHoja = "Resumen " & wsAnexos.name
k = 1

With wsAnexos

    'recorro las filas hasta encontrar el primer cuadro
    For i = 1 To filas
        
        If (LCase(.Cells(i, 1).Value) = "tipo de prestación" Or LCase(.Cells(i, 1).Value) = "tipo de prestacion") Then
        
            'recorro las filas hasta encontra una celda en blanco
            'es para poder guardar la fila de poblacion
            For m = i - 1 To filas
                
                If (.Cells(m, 2).Interior.Color = RGB(255, 255, 255) And .Cells(m, 2).Value <> "") Then
                    
                    i = m
                    Exit For
                    
                Else
                    
                    'guardo las columnas importantes y la fila de poblacion
                    For j = 1 To columnas
                    
                        If (wsAnexos.Cells(m, j).Value = "Nombre de la Prestación") Then
                
                            nombreColumna = j
                            
                        ElseIf (wsAnexos.Cells(m, j).Value = "Código" Or wsAnexos.Cells(m, j).Value = "Codigo SUMAR") Then
                        
                            codigoColumna = j
                        
                        ElseIf (wsAnexos.Cells(m, j).Value = "Precio") Then
                        
                            precioColumna = j
                            
                        ElseIf (LCase(.Cells(m, j).Value) = "normal") Then
                        
                            poblacionFila = m
                            flag = True
                
                        End If
                    
                    Next j
                    
                End If
                
            Next m
            
        End If
        
        If (flag = True) Then
        
            If (.Cells(i, precioColumna).Value <> "") Then
                    
                If (LCase(.Cells(i, codigoColumna).Value) = "ro" Or LCase(.Cells(i, codigoColumna).Value) = "ds") Then
                    
                    codigoAuxiliar = wsAnexos.Cells(i, codigoColumna).Value & wsAnexos.Cells(i, codigoColumna + 1).Value & wsAnexos.Cells(i, codigoColumna + 2).Value
                    codigoAuxiliar = Replace(codigoAuxiliar, " ", "")
                    codigoAuxiliar = Replace(codigoAuxiliar, "VMD(*)", "")
                    
                    codigo(k) = codigoAuxiliar
                    nombre(k) = wsAnexos.Cells(i, nombreColumna).Value
                    precio(k) = wsAnexos.Cells(i, precioColumna).Value
                    poblacion(k) = "Todas"
                    
                    k = k + 1
                    
                Else
                    
                    For j = codigoColumna + 3 To precioColumna - 1
                    
                        If (wsAnexos.Cells(i, j).Value <> "") Then
                            
                            codigoAuxiliar = wsAnexos.Cells(i, codigoColumna).Value & wsAnexos.Cells(i, codigoColumna + 1).Value & wsAnexos.Cells(i, codigoColumna + 2).Value
                            codigoAuxiliar = Replace(codigoAuxiliar, " ", "")
                            codigoAuxiliar = Replace(codigoAuxiliar, "VMD(*)", "")
                            
                            codigo(k) = codigoAuxiliar
                            nombre(k) = wsAnexos.Cells(i, nombreColumna).Value
                            precio(k) = wsAnexos.Cells(i, precioColumna).Value
                            poblacion(k) = wsAnexos.Cells(poblacionFila, j).Value
                            
                            k = k + 1
                            
                        End If
                
                    Next j
                    
                End If
                               
            End If
        
        End If
    
    Next i
    
End With

Call armarCuadrosNomenclador(nombreHoja, codigo, nombre, precio, poblacion)

End Function

'resumen: recorre una solapa del nomenclador especifica y toma los codigos, descripciones y precios, los pasa a otra
'         funcion y arma los cuadros depurados y resumidos
'parametros: void
'retorno: void
Public Function resumenGeneral(ByVal wsHoja As Worksheet)

Dim i As Double, x As Double, filas As Double, columnas As Double
Dim j As Integer, k As Integer, codigoColumna As Integer, nombreColumna As Integer
Dim precioColumna As Integer, poblacionFila As Integer, comienzoCuadro As Integer
Dim nombreHoja As String, codigoCorregido As String, columnaMiniscula As String
Dim codigo() As Variant, nombre() As Variant, precio() As Variant, poblacion() As Variant, codigoAuxiliar As Variant
Dim flag As Boolean, esInternacion As Boolean

ReDim codigo(1 To 1500)
ReDim nombre(1 To 1500)
ReDim precio(1 To 1500)
ReDim poblacion(1 To 1500)

columnas = 20
filas = filaMax(columnas, wsHoja)
nombreHoja = "Resumen " & wsHoja.name
k = 1

Call desagruparCeldasHoja(wsHoja)
Call eliminarFilasNomenclador(wsHoja, filas, columnas)

'calculo la nueva cantidad de filas luego de eliminar las que estaban en blanco
filas = filaMax(columnas, wsHoja)

With wsHoja
    
    'recorro todas las filas de la hoja
    For i = 1 To filas
        
        'entra a cada cuadro y guarda la ubicacion de las columnas importantes
        If (LCase(.Cells(i, 1).Value) = "linea de cuidado" Or LCase(.Cells(i, 1).Value) = "línea de cuidado") Then
            
            comienzoCuadro = i
            
            For j = 1 To columnas
                
                'convierto el nombre de la columna actual a minuscula
                columnaMiniscula = Replace(LCase(.Cells(i, j).Value), " ", "")
                
                If (columnaMiniscula = "nombredelaprestacion" Or columnaMiniscula = "nombredelaprestación" Or _
                columnaMiniscula = "modulo" Or columnaMiniscula = "módulo" Or columnaMiniscula = "cirugía" Or _
                columnaMiniscula = "cirugia" Or columnaMiniscula = "conceptosincluidos") Then
                
                    nombreColumna = j
                    
                ElseIf (columnaMiniscula = "código" Or columnaMiniscula = "códigosumar" Or _
                columnaMiniscula = "codigo" Or columnaMiniscula = "codigosumar") Then
                
                    codigoColumna = j
                
                ElseIf (columnaMiniscula = "precio" Or columnaMiniscula = "precioxdía" Or _
                LCase(columnaMiniscula) = "valor") Then
                
                    precioColumna = j
                    esInternacion = False
                    
                ElseIf (columnaMiniscula = "díapostquirúrgicoencuidadosintermedios" Or _
                columnaMiniscula = "díaestadapostquirúrgicaensalacomún" Or columnaMiniscula = "valorcubierto" Or _
                columnaMiniscula = "preciodíaestadapostquirúrgicaensalacomún") Then
                
                    precioColumna = j
                    
                    'sirve para saver si el cuado actual es o no de codigos de internacion
                    'y si lo es tomar correctmante la fila de precio con ayuda de la variable
                    'comienzoCuadro
                    esInternacion = True
                    
                    'esta suma es para que no se guarden las cabezeras en la array
                    i = i + 1
                
                End If
                
            Next j
            
            'para que no arranque con las filas en blanco
            'revisar si sirve
            flag = True
            
        End If
        
        If (flag = True) Then
        
            If ((.Cells(i, precioColumna).Value <> "" And LCase(.Cells(i, precioColumna).Value) <> "precio" And _
            LCase(.Cells(i, precioColumna).Value) <> "precio x día" And LCase(.Cells(i, precioColumna).Value) <> "valor" And _
            .Cells(i, precioColumna).Value <> "Día post quirúrgico en cuidados intermedios" And _
            .Cells(i, precioColumna).Value <> "Día Estada Post Quirúrgica en Sala Común" And _
            .Cells(i, precioColumna).Value <> "Valor Cubierto" And .Cells(i, codigoColumna).Value <> "") Or esInternacion = True) Then
            
'                If (largoArray(codigo) = k) Then
'
'                    'necesito arrays individuales porque el "ReDim Preserve" solo sirve para arrays unidimensionales
'                    ReDim Preserve codigo(1 To k + 12)
'                    ReDim Preserve nombre(1 To k + 12)
'                    ReDim Preserve precio(1 To k + 12)
'                    ReDim Preserve poblacion(i To k + 12)
'
'                End If
                
                'elimino espacios, salto de linea, (**) y (*). Convierto ",", ";" y "/" en "-"
                'siempre que la fila la fila no sea la cabecera
                'revisar porque no funciona correctamente
                If (.Cells(i, codigoColumna).Value <> "CÓDIGO" Or .Cells(i, codigoColumna).Value <> "CODIGO" Or _
                .Cells(i, codigoColumna).Value <> "CÓDIGO SUMAR" Or .Cells(i, codigoColumna).Value <> "CODIGO SUMAR") Then
                    
                    .Cells(i, codigoColumna).Value = Replace(.Cells(i, codigoColumna).Value, " ", "")
                    .Cells(i, codigoColumna).Value = Replace(.Cells(i, codigoColumna).Value, vbLf, "")
                    .Cells(i, codigoColumna).Value = Replace(.Cells(i, codigoColumna).Value, "(**)", "")
                    .Cells(i, codigoColumna).Value = Replace(.Cells(i, codigoColumna).Value, "(*)", "")
                    .Cells(i, codigoColumna).Value = Replace(.Cells(i, codigoColumna).Value, ",", "-")
                    .Cells(i, codigoColumna).Value = Replace(.Cells(i, codigoColumna).Value, ";", "-")
                    .Cells(i, codigoColumna).Value = Replace(.Cells(i, codigoColumna).Value, "/", "-")
                    .Cells(i, codigoColumna).Value = Replace(.Cells(i, codigoColumna).Value, "VMD", "")
                    .Cells(i, codigoColumna).Value = Replace(.Cells(i, codigoColumna).Value, "Vermatrizdiagnóstica", "")
                    .Cells(i, codigoColumna).Interior.Color = RGB(144, 238, 144)
                    
                End If
                
                'parte para guardar codigos que se dividen por "-"
                If (InStr(.Cells(i, codigoColumna).Value, "-") <> 0) Then
                    
                    codigoAuxiliar = Split(.Cells(i, codigoColumna).Value, "-")
                    
                    For j = 0 To UBound(codigoAuxiliar)
                        
                        If (codigoAuxiliar(j) <> "") Then
                        
                            If (j > 0) Then
                                
                                codigo(k) = Left(codigoAuxiliar(0), 6) & codigoAuxiliar(j)
                                nombre(k) = .Cells(i, nombreColumna).Value
                                poblacion(k) = wsHoja.name
                                
                                If (esInternacion = True) Then
                                    
                                    precio(k) = .Cells(comienzoCuadro + 1, precioColumna).Value
                                    
                                Else
                                
                                    precio(k) = .Cells(i, precioColumna).Value
                                    
                                End If
                                
                                
                                k = k + 1
                                
                            Else
                                
                                codigo(k) = codigoAuxiliar(j)
                                nombre(k) = .Cells(i, nombreColumna).Value
                                poblacion(k) = wsHoja.name
                                
                                If (esInternacion = True) Then
                                    
                                    precio(k) = .Cells(comienzoCuadro + 1, precioColumna).Value
                                    
                                Else
                                
                                    precio(k) = .Cells(i, precioColumna).Value
                                    
                                End If
                                
                                k = k + 1
                                
                            End If
                        
                        End If
                    
                    Next j
                
'                'coma
'                ElseIf (InStr(.Cells(i, codigoColumna).Value, ",") <> 0) Then
'
'                    codigoAuxiliar = Split(.Cells(i, codigoColumna).Value, ",")
'
'                    For j = 0 To UBound(codigoAuxiliar)
'
'                        If (j > 0) Then
'
'                            codigo(k) = Left(codigoAuxiliar(0), 6) & codigoAuxiliar(j)
'                            nombre(k) = .Cells(i, nombreColumna).Value
'                            poblacion(k) = wsHoja.name
'
'                            If (esInternacion = True) Then
'
'                                precio(k) = .Cells(comienzoCuadro + 1, precioColumna).Value
'
'                            Else
'
'                                precio(k) = .Cells(i, precioColumna).Value
'
'                            End If
'
'                            k = k + 1
'
'                        Else
'
'                            codigo(k) = codigoAuxiliar(j)
'                            nombre(k) = .Cells(i, nombreColumna).Value
'                            poblacion(k) = wsHoja.name
'
'                            If (esInternacion = True) Then
'
'                                precio(k) = .Cells(comienzoCuadro + 1, precioColumna).Value
'
'                            Else
'
'                                precio(k) = .Cells(i, precioColumna).Value
'
'                            End If
'
'                            k = k + 1
'
'                        End If
'
'                    Next j
'
'                'punto y coma
'                ElseIf (InStr(.Cells(i, codigoColumna).Value, ";") <> 0) Then
'
'                    codigoAuxiliar = Split(.Cells(i, codigoColumna).Value, ";")
'
'                    For j = 0 To UBound(codigoAuxiliar)
'
'                        If (j > 0) Then
'
'                            codigo(k) = Left(codigoAuxiliar(0), 6) & codigoAuxiliar(j)
'                            nombre(k) = .Cells(i, nombreColumna).Value
'                            poblacion(k) = wsHoja.name
'
'                            If (esInternacion = True) Then
'
'                                precio(k) = .Cells(comienzoCuadro + 1, precioColumna).Value
'
'                            Else
'
'                                precio(k) = .Cells(i, precioColumna).Value
'
'                            End If
'
'                            k = k + 1
'
'                        Else
'
'                            codigo(k) = codigoAuxiliar(j)
'                            nombre(k) = .Cells(i, nombreColumna).Value
'                            poblacion(k) = wsHoja.name
'
'                            If (esInternacion = True) Then
'
'                                precio(k) = .Cells(comienzoCuadro + 1, precioColumna).Value
'
'                            Else
'
'                                precio(k) = .Cells(i, precioColumna).Value
'
'                            End If
'
'                            k = k + 1
'
'                        End If
'
'                    Next j
                
'                'salto de linea
'                ElseIf (InStr(.Cells(i, codigoColumna).Value, vbLf) <> 0) Then
'
'                    codigoAuxiliar = Split(.Cells(i, codigoColumna).Value, vbLf)
'
'                    For j = 0 To UBound(codigoAuxiliar)
'
'                        codigo(k) = codigoAuxiliar(j)
'                        nombre(k) = .Cells(i, nombreColumna).Value
'                        poblacion(k) = wsHoja.name
'
'                        If (esInternacion = True) Then
'
'                                precio(k) = .Cells(comienzoCuadro + 1, precioColumna).Value
'
'                        Else
'
'                            precio(k) = .Cells(i, precioColumna).Value
'
'                        End If
'
'                        k = k + 1
'
'                    Next j
'
                'para todos los codigos individuales
                Else

                codigo(k) = .Cells(i, codigoColumna).Value
                nombre(k) = .Cells(i, nombreColumna).Value
                poblacion(k) = wsHoja.name

                If (esInternacion = True) Then

                    precio(k) = .Cells(comienzoCuadro + 1, precioColumna).Value

                Else

                    precio(k) = .Cells(i, precioColumna).Value

                End If

                k = k + 1

                End If
                
            End If
            
             If ((.Cells(i + 1, 2).Value = "" And .Cells(i + 2, 3).Interior.Color = RGB(0, 176, 239)) Or _
             (.Cells(i + 1, 2).Value = "" And .Cells(i + 2, 3).Interior.Color = RGB(0, 176, 240))) Then
        
                esInternacion = False
                
            End If
        
        End If
    
    Next i

End With

Call armarCuadrosNomenclador(nombreHoja, codigo, nombre, precio, poblacion)

End Function
'resumen: arma los cuadros resumidos del nomenclador
'parametros: nombreHoja es el nombre de la hoja nueva donde se guardan los datos de la matriz, codigos es un array con los codigos,
'            nombres, es un array con las descripciones de los codigos, precios es un array con los precios de los codigos segun
'            poblacion y poblaciones es un array con las poblaciones posibles para cada codigo
'retorno: void
Private Function armarCuadrosNomenclador(ByVal nombreHoja As String, ByVal codigos As Variant, ByVal nombres As Variant, _
ByVal precios As Variant, poblaciones As Variant)

Dim i As Double, filas As Double
Dim columnas As Integer
Dim wsHoja As Worksheet

Call crearHoja(nombreHoja)

Set wsHoja = ActiveWorkbook.Sheets(nombreHoja)

Call limpiarHoja(wsHoja)

With wsHoja
    
    .Columns("C:C").NumberFormat = "@"
    .Columns("D:D").Style = "Currency"
    .Cells(1, 1).Value = "Codigos"
    .Cells(1, 2).Value = "Nombres"
    .Cells(1, 3).Value = "Poblacion"
    .Cells(1, 4).Value = "Precio"
    
End With

For i = 2 To largoArray(codigos)
    
    With wsHoja
        
        .Cells(i, 1).Value = codigos(i - 1)
        .Cells(i, 2).Value = nombres(i - 1)
        .Cells(i, 3).Value = poblaciones(i - 1)
        .Cells(i, 4).Value = precios(i - 1)

    End With
    
Next i

columnas = columnaMax(wsHoja)
filas = filaMax(columnas, wsHoja)

For i = 2 To filas
    
    With wsHoja.Cells(i, 3)
        
        Select Case .Value
            
            Case Is = "Normal"
                .Value = "Embarazo normal"
                
            Case Is = "Riesgo"
                .Value = "Embarazo riesgoso"
                
            Case Is = "0- 5"
                .Value = "Niños de 0 a 5"
            
            Case Is = "6-9"
                .Value = "Niños de 6 a 9"
                
            Case Is = "10- 19"
                .Value = "Adolescentes"
                
            Case Is = "20- 64"
                .Value = "Adultos"
                
        End Select
        
    End With
    
Next i

Call formatos(wsHoja)

End Function

'resumen: le da formato a la solapa de resumen de codigos
'parametros: wsHoja es la hoja donde se aplica el formato
'retorno: void
Private Function formatos(ByVal wsHoja As Worksheet)

Dim columnas As Integer
Dim filas As Double

columnas = columnaMax(wsHoja)
filas = filaMax(columnas, wsHoja)


wsHoja.Range("A1").Select
wsHoja.Range(Selection, Selection.End(xlToRight)).Select

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

wsHoja.Range(Selection, Selection.End(xlDown)).Select
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

wsHoja.Range("A1").Select
wsHoja.Range(Selection, Selection.End(xlToRight)).Select

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

Columns("A:A").Select

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

Columns("D:D").Select

Selection.Style = "Currency"

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

Columns("F:F").Select

With Selection
    .HorizontalAlignment = xlCenter
    .WrapText = False
    .Orientation = 0
    .AddIndent = False
    .IndentLevel = 0
    .ShrinkToFit = False
    .ReadingOrder = xlContext
    .MergeCells = False
End With

Columns("E:E").Select

With Selection
    .HorizontalAlignment = xlCenter
    .WrapText = False
    .Orientation = 0
    .AddIndent = False
    .IndentLevel = 0
    .ShrinkToFit = False
    .ReadingOrder = xlContext
    .MergeCells = False
End With

wsHoja.Range("A1").Select
wsHoja.Range(Selection, Selection.End(xlToRight)).Select

With Selection
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

Cells.Select
With Selection
    .WrapText = False
    .Orientation = 0
    .AddIndent = False
    .IndentLevel = 0
    .ShrinkToFit = False
    .ReadingOrder = xlContext
    .MergeCells = False
End With

wsHoja.Range("A1").Select
wsHoja.Range(Selection, Selection.End(xlToRight)).Select
wsHoja.Range(Selection, Selection.End(xlDown)).Select
Selection.AutoFilter

wsHoja.AutoFilter.Sort.SortFields.Clear

wsHoja.AutoFilter.Sort.SortFields.Add Key:=wsHoja.Range _
    ("A2:A" & filas), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
    xlSortNormal

wsHoja.AutoFilter.Sort.SortFields.Add Key:=wsHoja.Range _
    ("C2:C" & filas), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
    xlSortNormal

wsHoja.AutoFilter.Sort.SortFields.Add Key:=wsHoja.Range _
    ("D2:D" & filas), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
    xlSortNormal

With wsHoja.AutoFilter.Sort
    .Header = xlYes
    .MatchCase = False
    .Orientation = xlTopToBottom
    .SortMethod = xlPinYin
    .Apply
End With

Columns("A:F").Select
Columns("A:F").EntireColumn.AutoFit

Columns("B:B").Select
Selection.ColumnWidth = 53

wsHoja.Range("A1").Select
    
End Function

'resumen: elimina filas en blanco y que comiencen por "(*)" para que qude el siguiente formato, cuadro, fila en blanco y cuadro
'parametros: wshoja es la hoja en donde se eliminan las filas, filas es la cantidad de filas que tiene wsHoja, columnas es
'            es la cantidad de columnas que tienen wsHoja
'retorno: void
Public Function eliminarFilasNomenclador(ByVal wsHoja As Worksheet, ByVal filas As Double, ByVal columnas As Double)

Dim i As Double, contador As Double

For i = 1 To filas
    
    With wsHoja
    
        If ((.Cells(i, 1).Value = "" And .Cells(i + 1, 1).Value = "" And (.Cells(i, 1).Interior.Color <> RGB(0, 204, 255) And .Cells(i, 1).Interior.Color <> RGB(0, 176, 239) And .Cells(i, 1).Interior.Color <> RGB(0, 176, 240))) Or _
        Left(.Cells(i, 1).Value, 3) = "(*)" Or LCase(.Cells(i, 1).Value) = "total") Then
            
            .Rows(i).Select
            Application.CutCopyMode = False
            Selection.Delete SHIFT:=xlUp
            i = i - 2
            contador = contador + 1
            
        End If
        
        If (contador = filas) Then
            
            Exit For
            
        End If
    
    End With
    
Next i

End Function

