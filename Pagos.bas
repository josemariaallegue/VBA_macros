Attribute VB_Name = "Pagos"
Option Explicit

'resumen: crea la hoja "Resumen" para el metodo "revision_muestra_pagos" con los cuadros correspondientes
'parametros: un array con las datos de la muestra, la n para la provincia, la cantidad de efectores, la cantidad de codigos no elegibles tomadas
'retorno: void
Private Function cuadrosMuestraPagos(ByVal cuieArray As Variant, ByVal n As Integer, ByVal contador1, ByVal contador2 As Integer)

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
    Exit Function
    
erroralcrearhoja:
Sheets(Sheets.Count).name = "Resumen" & Sheets.Count - 1
Resume Next
    
End Function

'resumen: da formato al cuadro de resumen del metodo "revision_muestra_pagos"
'parametros: void
'retorno: void'
Private Function formatosRevisionMuestraPagos()

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
End Function

'resumen: revisa la muestra de pagos enviada de sistemas
'parametros: void
'retorno: void
Public Sub revision_muestra_pagos()

Call preparacionInicio

'declaro las variables
Dim i, j, k, m, n As Integer
Dim cuieColumna, prestacionColumna, cantidadMuestraColumna, muestraColumna, beneficiariosValidosColumna As Integer
Dim cuieArray(1 To 15, 1 To 5), mensaje, codigos As String, codigos2 As String, codigos3 As String
Dim codigos4 As String, codigos5 As String, codigos6 As String, codigos7 As String
Dim contador1, contador2 As Integer
Dim auxiliar As Variant

'asigno a la variables que contiene los codigos su valor
codigos = "APA001A98;APA001W78;APA001X75;APA001X86;APA002X75;APA002X76;APA002X80;APA002A98;CTC001A97;CTC001B80;CTC001D11;CTC001R74;CTC001R78;CTC001T91;CTC002A97;CTC002R96;CTC002T82;CTC002T91;CTC005B80;CTC005W78;CTC006W78;CTC007W84;CTC008A97;CTC010A97;CTC010W78;CTC011A97;CTC012A03;CTC012A81;CTC012A92;CTC012D01;CTC012D10;CTC012H71;CTC012H72;CTC012H76;CTC012L72;CTC012L73;CTC012L74;CTC012L77;CTC012L78;CTC012L80;CTC012R72;CTC012R77;CTC012R80;CTC012S13;CTC012T11;IMV001A98;IMV002A98;IMV003A98;IMV004A98;IMV005A98;IMV006A98;IMV007A98;IMV013A98;IMV014A98;ITE002A40;ITE002A41;ITE002A42;ITE002A44;ITE002R78;ITQ001W90;ITQ001W91;ITQ002W88;ITQ002W89;ITQ005W06;ITQ006W07;ITQ007W08;TAT001A98;TAT002A98;TAT003A98;TAT007A98;TAT008A98;TAT009A98;TAT010A98;TAT013A98;TAT014A98;CTC001T79;CTC001R96;CTC001R81"
codigos2 = "CAW003A98;CAW004A98;CAW005A98;CAW006A75;CAW006A97;CAW006B72;CAW006B73;CAW006B78;CAW006B80;CAW006B81;CAW006B82;CAW006B90;CAW006D61;CAW006D62;CAW006D72;CAW006D96;CAW006K73;CAW006K83;CAW006K86;CAW006K96;CAW006T79;CAW006T82;CAW006T83;CAW006T89;CAW006T90;CAW006Y70;COT016A98;COT018A98;CTC001B72;CTC001B73;CTC001H86;CTC001P20;CTC001P23;CTC001P24;CTC001P98;CTC002B72;CTC002B73;CTC002P20;CTC002P23;CTC002P24;CTC005W78;CTC006W78;CTC007O10.0;CTC007O10.4;CTC007O16;CTC009A97;CTC020P07.0;CTC020P07.2;CTC021P07.0;CTC021P07.2;CTC021Q03;CTC021Q05;CTC021Q39.0;CTC021Q39.1;CTC021Q39.2;CTC021Q41;CTC021Q42;CTC021Q42.0;CTC021Q42.1;CTC021Q42.2;CTC021Q42.3;CTC021Q43.3;CTC021Q43.4;CTC021Q79.3;CTC022O10.0;CTC022O10.4;CTC022O16;CTC022O24.4;CTC047A98;CTC047U89;CTC050A98;CTC050T89;CTC050T90;ITE001R78;ITE013P07.0;ITE013P07.2;ITE014P07.0;ITE014P07.2;NTN007K22;NTN008K22;NTN009K22;NTN010K22"
codigos3 = "IGR002;IGR003;IGR004;IGR005;IGR006;IGR007;IGR008;IGR009;IGR010;IGR011;IGR012;IGR013;IGR014;IGR015;IGR017;IGR018;IGR019;IGR020;IGR021;IGR022;IGR023;IGR024;IGR025;IGR026;IGR028;IGR029;IGR030;IGR031;IGR032;IGR037;IGR038;IGR039;IGR040;IGR041;IGR042;IGR043;IGR044;IGR045;IGR046;IGR047;IGR048;IGR049;IMV001;IMV002;IMV003;IMV004;IMV005;IMV006;IMV007;IMV008"
codigos4 = "IMV009;IMV010;IMV011;IMV012;IMV013;IMV014;IMV015;IMV016;IMV017;IMV018;IMV019;LBL001;LBL002;LBL003;LBL004;LBL005;LBL006;LBL008;LBL009;LBL010;LBL011;LBL012;LBL013;LBL014;LBL015;LBL016;LBL017;LBL018;LBL019;LBL020;LBL021;LBL022;LBL023;LBL024;LBL025;LBL026;LBL027;LBL028;LBL029;LBL030;LBL031;LBL032;LBL033;LBL034;LBL035;LBL036;LBL037;LBL038;LBL040;LBL041"
codigos5 = "LBL042;LBL043;LBL044;LBL045;LBL046;LBL047;LBL048;LBL049;LBL050;LBL051;LBL052;LBL053;LBL054;LBL055;LBL056;LBL057;LBL058;LBL059;LBL060;LBL061;LBL062;LBL063;LBL064;LBL065;LBL066;LBL067;LBL068;LBL069;LBL070;LBL071;LBL072;LBL073;LBL074;LBL075;LBL076;LBL078;LBL079;LBL080;LBL081;LBL082;LBL083;LBL084;LBL085;LBL086;LBL087;LBL088;LBL089;LBL090;LBL091;LBL092"
codigos6 = "LBL093;LBL094;LBL095;LBL096;LBL097;LBL098;LBL099;LBL100;LBL101;LBL102;LBL103;LBL104;LBL105;LBL106;LBL107;LBL108;LBL109;LBL110;LBL111;LBL112;LBL113;LBL114;LBL115;LBL116;LBL117;LBL118;LBL119;LBL120;LBL121;LBL122;LBL123;LBL124;LBL125;LBL126;LBL127;LBL128;LBL129;LBL130;LBL131;LBL132;LBL133;LBL134;LBL135;LBL136;LBL137;LBL138;LBL139;LBL140;PRP001;PRP002"
codigos7 = "PRP003;PRP004;PRP005;PRP006;PRP008;PRP009;PRP010;PRP011;PRP014;PRP016;PRP018;PRP019;PRP020;PRP021;PRP022;PRP024;PRP025;PRP026;PRP028;PRP029;PRP030;PRP031;PRP033;PRP034;PRP035;PRP036;PRP037;PRP038;PRP039;PRP040;PRP041;PRP042;PRP044;PRP045;PRP046;PRP047;PRP048"

i = 1
j = 2
k = 1
muestraColumna = 0


Worksheets("Database").Activate

'guardo la ubicacion de las columnas
Do Until ActiveSheet.Cells(1, i).Value = ""

    If (ActiveSheet.Cells(1, i).Value = "CUIE_EFECTOR" Or ActiveSheet.Cells(1, i).Value = "CUIE") Then
        
        cuieColumna = i
        
    ElseIf (ActiveSheet.Cells(1, i).Value = "CODIGO_PRESTACION") Then
    
        prestacionColumna = i
    
    ElseIf (ActiveSheet.Cells(1, i).Value = "N") Then
    
        n = ActiveSheet.Cells(2, i).Value
    
    ElseIf (ActiveSheet.Cells(1, i).Value = "MUESTRA" Or ActiveSheet.Cells(1, i).Value = "MUESTRAS" _
    Or ActiveSheet.Cells(1, i).Value = "SELECCION" Or ActiveSheet.Cells(1, i).Value = "MUESTRA_VALIDO") Then
        
        muestraColumna = i
    
    ElseIf (ActiveSheet.Cells(1, i).Value = "CANTIDAD_MUESTRA") Then
        
        cantidadMuestraColumna = i
        
    ElseIf (ActiveSheet.Cells(1, i).Value = "CUIE_X_BENEF_VALIDOS") Then
        
        beneficiariosValidosColumna = i
        
    End If
  
    i = i + 1
    
Loop


Do Until ActiveSheet.Cells(j, 1).Value = ""
    
    'entra al if cuando entre dos filas cambia el cuie
    'en la primer posicion de cuieArray se guarda el CUIE, en la segunda la cantidad de casos validos para el efector
    'en la tercera la cantidad de muestra por formula
    'en el contador se guardan la cantidad de efectores tomados
    If (ActiveSheet.Cells(j, cuieColumna).Value <> ActiveSheet.Cells(j - 1, cuieColumna).Value) Then
    
        cuieArray(k, 1) = ActiveSheet.Cells(j, cuieColumna).Value
        cuieArray(k, 2) = ActiveSheet.Cells(j, beneficiariosValidosColumna).Value
        cuieArray(k, 3) = ActiveSheet.Cells(j, cantidadMuestraColumna).Value
        
        contador1 = contador1 + 1

        k = k + 1

    End If
    
    'cuento la cantidad de casos que tienen x
    'si no esta la columna "Muestra" el On Error obliga a que se cumpla la condicion
    On Error GoTo sinColumnaMuestra
    If (muestraColumna <> 0 And LCase(ActiveSheet.Cells(j, muestraColumna).Value) = "x") Then
        
        For m = 1 To 12
            
            'cuento la cantidad de casos seleccionados por efector
            If (cuieArray(m, 1) = ActiveSheet.Cells(j, cuieColumna).Value) Then
                
                    cuieArray(m, 4) = cuieArray(m, 4) + 1
                
            End If
            
        Next m
        
        'reviso que los codigos seleccionados sean los del listado de codigos elegibles
        'pinto la celda del codigo invalido
        'cuento la cantidad de codigos invalidos en contador2
        If (InStr(1, codigos, ActiveSheet.Cells(j, prestacionColumna).Value) = 0 And _
        InStr(1, codigos2, ActiveSheet.Cells(j, prestacionColumna).Value) = 0 And _
        InStr(1, codigos3, Left(ActiveSheet.Cells(j, prestacionColumna).Value, 6)) = 0 And _
        InStr(1, codigos4, Left(ActiveSheet.Cells(j, prestacionColumna).Value, 6)) = 0 And _
        InStr(1, codigos5, Left(ActiveSheet.Cells(j, prestacionColumna).Value, 6)) = 0 And _
        InStr(1, codigos6, Left(ActiveSheet.Cells(j, prestacionColumna).Value, 6)) = 0 And _
        InStr(1, codigos7, Left(ActiveSheet.Cells(j, prestacionColumna).Value, 6)) = 0) Then
    
            ActiveSheet.Cells(j, prestacionColumna).Interior.Color = RGB(255, 255, 0)
            contador2 = contador2 + 1
            cuieArray(k - 1, 5) = cuieArray(k - 1, 5) + 1
        
        End If

    End If
    
    j = j + 1
    
Loop

'armo la solapa con el resumen con los datos obtenidos
Call cuadrosMuestraPagos(cuieArray, n, contador1, contador2)
Call preparacionFinal

Exit Sub

'excepcion por si no esta la columna "Muestra"
sinColumnaMuestra:

muestraColumna = 0

Resume Next



End Sub
