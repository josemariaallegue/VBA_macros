Attribute VB_Name = "Padron"
Option Explicit

'resumen: revisa la muestra de padron enviada de sistemas
'parametros: void
'retorno: void
Sub revision_muestra_padron()

Call preparacionInicio

'declaro las varibles
Dim i As Integer, j As Integer, k As Integer, m As Integer, contador As Integer
Dim cuieColumna As Integer, nColumna As Integer, codigoPrestacionColumna As Integer, cantidadMuestraColumna As Integer
Dim poblacionColumna As Integer, beneficiariosValidosColumna As Integer, muestraColumna As Integer
Dim fechaNacimientoColumna As Integer, fechaPrestacionColumna As Integer, provinciaColumna As Integer
Dim edad As Double, columnas As Double, filas As Double
Dim codigos0a1 As String, codigos1a2 As String, codigos2a6 As String, codigos6a9 As String, codigosAdolescentes1 As String
Dim codigosAdolescentes2 As String, codigosHombres As String, codigosMujeres As String, diagnosticosNoPermitidos As String
Dim cuieArray(), cantidadMuestraArray(), provinciaArray(), muestraArray(), n(), noElegiblesArray(), validosXcuie(), codigo, codigoIzquierda, diagnostico As String
Dim auxiliar As Variant
Dim flag As Boolean
Dim wsDatabase As Worksheet

'asigno a las variables que contienen los codigos sus valores
codigos0a1 = "CAW003A98;CTC001A97;CTC001H86;CTC020P07.0;CTC020P07.2;CTC021P07.0;CTC021P07.2;CTC021Q03;CTC021Q05;CTC021Q39.0;CTC021Q39.1;CTC021Q39.2;CTC021Q41;CTC021Q42;CTC021Q42.0;CTC021Q42.1;CTC021Q42.2;CTC021Q42.3;CTC021Q43.3;CTC021Q43.4;CTC021Q79.3;IGR005A98;IGR005L30;IGR025A98;IGR025L30;ITE013P07.0;ITE013P07.2;ITE014P07.0;ITE014P07.2;LBL001A98;LBL013A98;LBL035A98;LBL043A98;LBL115A98;LBL116A98;PRP017A46;PRP017A97;PRP021A97;PRP021H86;PRP022H86;TAT002A98;TAT003A98"
codigos1a2 = "CAW003A98;CTC001A97;CTC001R78;CTC002R78;CTC010A97;IMV001A98;IMV005A98;IMV015A98;ITE001R78;TAT002A98;TAT003A98"
codigos2a6 = "CAW003A98;CTC001A97;CTC001R78;CTC002R78;CTC010A97;CTC011A97;ITE001R78;TAT002A98;TAT003A98"
codigos6a9 = "CAW003A98;CAW006A75;CAW006A97;CAW006B72;CAW006B73;CAW006B78;CAW006B80;CAW006B81;CAW006B82;CAW006B90;CAW006D61;CAW006D62;CAW006D72;CAW006D96;CAW006K73;CAW006K83;CAW006K86;CAW006T79;CAW006T82;CAW006T83;CAW006T89;CAW006T90;CTC001A97;CTC001B80;CTC001R96;CTC001T83;CTC002B80;CTC002R96;CTC002T83;CTC009A97;CTC010A97;CTC011A97;IMV001A98;IMV002A98;IMV008A98;IMV011A98;PRP006R96;PRP026D60"
codigosAdolescentes1 = "CAW004A98;CAW005A98;CAW006A75;CAW006B72;CAW006B73;CAW006B78;CAW006B80;CAW006B81;CAW006B82;CAW006B90;CAW006D61;CAW006D62;CAW006D72;CAW006D96;CAW006K73;CAW006K83;CAW006K86;CAW006K96;CAW006T79;CAW006T82;CAW006T83;CAW006T89;CAW006T90;CAW006Y70;COT015A98;COT016A98;COT018A98;CTC001A97;CTC001B72;CTC001B73;CTC001B80;CTC001P20;CTC001P23;CTC001P24;CTC001P98;CTC001R96;CTC001T79;CTC001T82;CTC001T83;CTC002B72;CTC002B73;CTC002B80;CTC002P20;CTC002P23;CTC002P24;CTC002R96;CTC002T79;CTC002T82;CTC002T83;CTC005W78;CTC006W78;CTC007O10.0;CTC007O10.1;CTC007O10.2;CTC007O10.3;CTC007O10.4;CTC007O16;CTC007O24.4;CTC008A97;CTC009A97;CTC010A97;CTC011A97;CTC022O10.0;CTC022O10.1;CTC022O10.2;CTC022O10.3;CTC022O10.4;CTC022O16;CTC022O24.4;IGR003;IGR004;IGR007;IGR008;IGR017;IGR019;IGR020;IGR021;IGR022;IGR026;IGR030;IGR031W78"
codigosAdolescentes2 = "IGR032;IGR038;IMV008A98;IMV009A98;IMV010A98;IMV011A98;IMV013A98;IMV014A98;LBL002;LBL004;LBL005;LBL006;LBL008;LBL009;LBL010;LBL011;LBL014;LBL015;LBL016;LBL018;LBL021;LBL022;LBL025;LBL026;LBL027;LBL028;LBL029;LBL030;LBL031;LBL032;LBL033;LBL037;LBL038;LBL040;LBL041;LBL042;LBL044;LBL045;LBL047;LBL048;LBL050;LBL051;LBL052;LBL053;LBL055;LBL057;LBL058;LBL059;LBL060;LBL061;LBL062;LBL065;LBL066;LBL067;LBL068;LBL069;LBL070;LBL072;LBL073;LBL075;LBL076;LBL078;LBL079;LBL081;LBL083;LBL084;LBL086;LBL087;LBL088;LBL089;LBL091;LBL092;LBL094;LBL095;LBL097;LBL099;LBL100;LBL101;LBL103;LBL104;LBL105;LBL108;LBL112;LBL113;LBL114;LBL116;LBL117;LBL118;LBL119;LBL121;LBL122;LBL123;LBL124;LBL127;LBL128;LBL129;LBL130;LBL131;LBL132;LBL133;LBL134;PRP001;PRP003;PRP004;PRP005;PRP006;PRP007;PRP008;PRP009;PRP010;PRP011;PRP014;PRP017;PRP028;PRP029"
codigosMujeres = "APA001A98;APA001X75;APA001X86;CTC001A97;CTC005W78;CTC006W78;CTC007O10.0;CTC007O10.4;CTC007O16;CTC007O24.4;CTC008A97;CTC009A97;CTC022O10.0;CTC022O10.4;CTC022O16;CTC022O24.4;IGR014A98;IGR031W78;LBL065W78;LBL090O16;LBL099W78;LBL110W78;LBL111W78;LBL119W78;LBL121W78;LBL122W78;LBL128W78;PRP018A98;PRP018W78"
codigosHombres = "CTC001A97;CTC001A98;CTC009A97;CTC010A97;CTC011A97;CTC047A98;CTC047U89;CTC048K22;CTC050A98;CTC050T89;CTC050T90;IGR048A98;IGR048D04;IGR048D16;IGR048D18;IGR049A98;IGR049D04;IGR049D16;IGR049D18;LBL098A98;NTN007K22;NTN008K22;NTN009K22;NTN010K22"
diagnosticosNoPermitidos = "P20;P23;P24;P98;Z31"

i = 1
j = 2
k = 1
m = 0
flag = False

'redimenciono arrays
ReDim cuieArray(1 To 12)
ReDim cantidadMuestraArray(1 To 12)
ReDim provinciaArray(1 To 12)
ReDim muestraArray(1 To 12)
ReDim n(1 To 12)
ReDim noElegiblesArray(1 To 12)
ReDim validosXcuie(1 To 12)

Set wsDatabase = ActiveWorkbook.Sheets("Database")
columnas = columnaMax(wsDatabase)
filas = filaMax(columnas, wsDatabase)

With wsDatabase

    'guardo la ubicacion de las columnas importantes
    For i = 1 To columnas
        
        If (.Cells(1, i).Value = "CUIE") Then
            
            cuieColumna = i
            
        ElseIf (.Cells(1, i).Value = "CODIGO_PRESTACION") Then
        
            codigoPrestacionColumna = i
            
        ElseIf (.Cells(1, i).Value = "N") Then
        
            nColumna = i
            
        ElseIf (.Cells(1, i).Value = "CANTIDAD_MUESTRA") Then
            
            cantidadMuestraColumna = i
            
        ElseIf (.Cells(1, i).Value = "CATEGORIA_LIQUIDACION") Then
            
            poblacionColumna = i
        
        ElseIf (.Cells(1, i).Value = "BENEF_FECHA_NACIMIENTO") Then
            
            fechaNacimientoColumna = i
            
        ElseIf (.Cells(1, i).Value = "FECHA_ULTIMA_PRESTACION") Then
            
            fechaPrestacionColumna = i
            
        ElseIf (.Cells(1, i).Value = "PROVINCIA") Then
            
            provinciaColumna = i
            
        ElseIf (.Cells(1, i).Value = "CUIE_X_BENEF_VALIDOS") Then
        
            beneficiariosValidosColumna = i
            
        ElseIf (.Cells(1, i).Value = "MUESTRA" Or .Cells(1, i).Value = "MUESTRAS" _
        Or .Cells(1, i).Value = "SELECCION" Or .Cells(1, i).Value = "MUESTRA_VALIDO") Then
        
            muestraColumna = i
            
        End If
        
    Next i
        
    'recorrido vertical
    For i = 2 To filas
        
        'otorgo valor del codigo, primeros 6 caracteres del codigo, el diagnostico y la edad para cada caso
        codigo = UCase(.Cells(i, codigoPrestacionColumna).Value)
        codigoIzquierda = Left(codigo, 6)
        diagnostico = Right(codigo, 3)
        edad = (.Cells(i, fechaPrestacionColumna).Value - .Cells(i, fechaNacimientoColumna).Value) / 365
        
        'entra al if cuando entre dos filas cambia el cuie
        'guardo los cuie en cuieArray, la cantidad de muestra determinada por los calculos en cantidadMuestraArray
        'el id de la provincia en provinciaArray, la n para cada una de las provincias en n y la cantidad de casos
        'validos por cuie en validosXcuie
        If (.Cells(i, cuieColumna).Value <> .Cells(i - 1, cuieColumna).Value) Then
            
            If (largoArray(cuieArray) = k) Then
                
                'necesito arrays individuales porque el "ReDim Preserve" solo sirve para arrays unidimensionales
                ReDim Preserve cuieArray(1 To k + 12)
                ReDim Preserve cantidadMuestraArray(1 To k + 12)
                ReDim Preserve provinciaArray(1 To k + 12)
                ReDim Preserve n(1 To k + 12)
                ReDim Preserve noElegiblesArray(1 To k + 12)
                ReDim Preserve validosXcuie(1 To k + 12)
            
            End If
            
            cuieArray(k) = .Cells(i, cuieColumna).Value
            cantidadMuestraArray(k) = .Cells(i, cantidadMuestraColumna).Value
            provinciaArray(k) = .Cells(i, provinciaColumna).Value
            n(k) = .Cells(i, nColumna).Value
            validosXcuie(k) = .Cells(i, beneficiariosValidosColumna).Value
            
            'el flag esta para poder cambiar la posicion en el array muestraArray
            k = k + 1
            flag = False
    
        End If
        
        'entra cuando la columan de MUESTRA es x y cuando el cuie de la fila actual es igual al del anterior
        'cuento la cantidad de casos que tienen x
        'si no esta la columna "MUESTRA", "MUESTRAS" o "SELECCION" el On Error obliga a que se cumpla la condicion
        'y cuenta todos los casos del archivo
        On Error GoTo sinColumnaMuestra
        If (muestraColumna <> 0 And LCase(.Cells(i, muestraColumna).Value) = "x" And _
        LCase(.Cells(i, cuieColumna).Value) = LCase(.Cells(i - 1, cuieColumna).Value)) Then
            
            If (largoArray(muestraArray) = m) Then
                
                ReDim Preserve muestraArray(1 To m + 12)
                
            End If
            
            If (flag = False) Then
    
                m = m + 1
    '            muestraArray(m) = muestraArray(m) + 1
                flag = True
    
            End If
    
            muestraArray(m) = muestraArray(m) + 1
            
        
                If (edad >= 20 And edad < 65 And .Cells(i, poblacionColumna).Value = "Mujeres 20-64") Then
    
                    If (InStr(1, codigosMujeres, codigo) = 0) Then
    
                        .Cells(i, codigoPrestacionColumna).Interior.Color = RGB(255, 255, 0)
                        contador = contador + 1
                        noElegiblesArray(k) = noElegiblesArray(k) + 1
    
                    End If
    
                ElseIf (edad >= 20 And edad < 65 And .Cells(i, poblacionColumna).Value = "Hombres 20-64") Then
    
                    If (InStr(1, codigosHombres, codigo) = 0) Then
    
                        .Cells(i, codigoPrestacionColumna).Interior.Color = RGB(255, 255, 0)
                        contador = contador + 1
                        noElegiblesArray(k) = noElegiblesArray(k) + 1
    
                     End If
    
                ElseIf (edad >= 6 And edad < 10) Then
    
                    If (InStr(1, codigos6a9, codigo) = 0) Then
    
                        .Cells(i, codigoPrestacionColumna).Interior.Color = RGB(255, 255, 0)
                        contador = contador + 1
                        noElegiblesArray(k) = noElegiblesArray(k) + 1
    
                    End If
    
                ElseIf (edad >= 10 And edad < 20) Then
    
                    If ((InStr(1, codigosAdolescentes1, codigo) = 0) _
                    And (InStr(1, codigosAdolescentes2, codigo) = 0)) Then
        
                        If ((InStr(1, codigosAdolescentes1, codigoIzquierda) <> 0) _
                        Or (InStr(1, codigosAdolescentes2, codigoIzquierda) <> 0)) Then
        
                            If (InStr(1, diagnosticosNoPermitidos, diagnostico) <> 0) Then
        
                                .Cells(i, codigoPrestacionColumna).Interior.Color = RGB(255, 255, 0)
                                contador = contador + 1
                                noElegiblesArray(k) = noElegiblesArray(k) + 1
        
                            End If
        
                        Else
        
                            .Cells(i, codigoPrestacionColumna).Interior.Color = RGB(255, 255, 0)
                            contador = contador + 1
                            noElegiblesArray(k) = noElegiblesArray(k) + 1
        
                        End If
        
                    End If
        
                ElseIf (edad >= 0 And edad < 6) Then
        
                    
        
                    If (edad >= 0 And edad < 1) Then
        
                        If (InStr(1, codigos0a1, codigo) = 0) Then
        
                            .Cells(i, codigoPrestacionColumna).Interior.Color = RGB(255, 255, 0)
                            contador = contador + 1
                            noElegiblesArray(k) = noElegiblesArray(k) + 1
        
                        End If
        
                    ElseIf (edad >= 1 And edad < 2) Then
        
                        If (InStr(1, codigos1a2, codigo) = 0) Then
        
                            .Cells(i, codigoPrestacionColumna).Interior.Color = RGB(255, 255, 0)
                            contador = contador + 1
                            noElegiblesArray(k) = noElegiblesArray(k) + 1
        
                        End If
        
                    ElseIf (edad >= 2) Then
        
                        If (InStr(1, codigos2a6, codigo) = 0) Then
        
                            .Cells(i, codigoPrestacionColumna).Interior.Color = RGB(255, 255, 0)
                            contador = contador + 1
                            noElegiblesArray(k) = noElegiblesArray(k) + 1
        
                        End If
        
                    End If
        
            End If
    
        End If
        
    Next i

End With



'excepcion por si no esta la columna "Muestra"
sinColumnaMuestra:

muestraColumna = 0

Resume Next

'se arma el resumen con los datos obtenidos
Call cuadrosMuestraPadron(cuieArray, cantidadMuestraArray, provinciaArray, muestraArray, n, noElegiblesArray, validosXcuie, contador)
Call preparacionFinal

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



