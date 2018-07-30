Attribute VB_Name = "Padron"
Option Explicit

Sub revision_muestra_padron()

Call preparacionInicio

Dim i, j, k, m As Integer
Dim contador As Integer
Dim codigos0a1, codigos1a2, codigos2a6, codigos6a9, codigosAdolescentes1, codigosAdolescentes2, codigosHombres, codigosMujeres As String
Dim diagnosticosNoPermitidos As String
Dim cuieColumna, nColumna, codigoPrestacionColumna, cantidadMuestraColumna, poblacionColumna, beneficiariosValidosColumna As Integer
Dim fechaNacimientoColumna, fechaPrestacionColumna, provinciaColumna As Integer
Dim cuieArray(), cantidadMuestraArray(), provinciaArray(), muestraArray(), n(), noElegiblesArray(), validosXcuie(), codigo, codigoIzquierda, diagnostico As String
Dim edad As Double
Dim auxiliar As Variant
Dim flag As Boolean

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

ReDim cuieArray(1 To 12)
ReDim cantidadMuestraArray(1 To 12)
ReDim provinciaArray(1 To 12)
ReDim muestraArray(1 To 12)
ReDim n(1 To 12)
ReDim noElegiblesArray(1 To 12)
ReDim validosXcuie(1 To 12)

'selecciona la solapa "Database" para que no hay errores de ejecucion
Worksheets("Database").Activate

'recorrido horizontal de las columnas
Do Until ActiveSheet.Cells(1, i).Value = ""

    'guardo el valor de la columna de "CUIE"
    If (ActiveSheet.Cells(1, i).Value = "CUIE") Then
        
        cuieColumna = i
        
    End If
    
    'guardo el valor de la columna de "CODIGO_PRESTACION"
    If (ActiveSheet.Cells(1, i).Value = "CODIGO_PRESTACION") Then
    
        codigoPrestacionColumna = i
        
    End If
    
    'guardo el valor de "N"
    If (ActiveSheet.Cells(1, i).Value = "N") Then
    
        nColumna = i
        
    End If
    
    'guardo el valor de la columna de "CANTIDAD_MUESTRA"
    If (ActiveSheet.Cells(1, i).Value = "CANTIDAD_MUESTRA") Then
        
        cantidadMuestraColumna = i
        
    End If
    
    'guardo el valor de la columna de "CATEGORIA_LIQUIDACION"
    If (ActiveSheet.Cells(1, i).Value = "CATEGORIA_LIQUIDACION") Then
        
        poblacionColumna = i
    
    End If
    
    'guardo el valor de la columna de "BENEF_FECHA_NACIMIENTO"
    If (ActiveSheet.Cells(1, i).Value = "BENEF_FECHA_NACIMIENTO") Then
        
        fechaNacimientoColumna = i
        
    End If
    
    'guardo el valor de la columna de "FECHA_ULTIMA_PRESTACION"
    If (ActiveSheet.Cells(1, i).Value = "FECHA_ULTIMA_PRESTACION") Then
        
        fechaPrestacionColumna = i
        
    End If
    
    'guardo el valor de la columna de "PROVINCIA"
    If (ActiveSheet.Cells(1, i).Value = "PROVINCIA") Then
        
        provinciaColumna = i
        
    End If
    
    'guardo el valor de la columna de "CUIE_X_BENEF_VALIDOS"
    If (ActiveSheet.Cells(1, i).Value = "CUIE_X_BENEF_VALIDOS") Then
    
        beneficiariosValidosColumna = i
        
    End If
    
    
    
    i = i + 1
    
Loop

'recorrido vertical
Do Until ActiveSheet.Cells(j, 1).Value = ""
    
    'otorgo valor del codigo, primeros 6 caracteres del codigo y el diagnostico a 3 variables
    codigo = UCase(ActiveSheet.Cells(j, codigoPrestacionColumna).Value)
    codigoIzquierda = Left(codigo, 6)
    diagnostico = Right(codigo, 3)
    
    'guardo la muestraArray para cada cuie
    'en el contador se guardan la cantidad de efectores tomados
    If (ActiveSheet.Cells(j, cuieColumna).Value <> ActiveSheet.Cells(j - 1, cuieColumna).Value) Then
        
        If (largoArray(cuieArray) = k) Then
            
            'necesito arrays individuales porque el "ReDim Preserve" solo sirve para arrays unidimensionales
            ReDim Preserve cuieArray(1 To k + 12)
            ReDim Preserve cantidadMuestraArray(1 To k + 12)
            ReDim Preserve provinciaArray(1 To k + 12)
            ReDim Preserve n(1 To k + 12)
            ReDim Preserve noElegiblesArray(1 To k + 12)
            ReDim Preserve validosXcuie(1 To k + 12)
        
        End If
        
        cuieArray(k) = ActiveSheet.Cells(j, cuieColumna).Value
        cantidadMuestraArray(k) = ActiveSheet.Cells(j, cantidadMuestraColumna).Value
        provinciaArray(k) = ActiveSheet.Cells(j, provinciaColumna).Value
        n(k) = ActiveSheet.Cells(j, nColumna).Value
        validosXcuie(k) = ActiveSheet.Cells(j, beneficiariosValidosColumna).Value

        k = k + 1
        flag = False

    End If
    
    If (LCase(ActiveSheet.Cells(j, cuieColumna).Value) = LCase(ActiveSheet.Cells(j - 1, cuieColumna).Value)) Then
        
        If (largoArray(muestraArray) = m) Then
            
            ReDim Preserve muestraArray(1 To m + 12)
            
        End If
        
        If (flag = False) Then
            
            m = m + 1
            muestraArray(m) = muestraArray(m) + 1
            flag = True
            
        End If
        
        muestraArray(m) = muestraArray(m) + 1
        
    End If
    
    'reviso que los codigos seleccionados sean los del listado de codigos elegibles
    'pinto la celda del codigo invalido
    'cuento la cantidad de codigos invalidos en contador
    Select Case ActiveSheet.Cells(j, poblacionColumna).Value

        Case "Mujeres 20-64"

            If (InStr(1, codigosMujeres, codigo) = 0) Then

                ActiveSheet.Cells(j, codigoPrestacionColumna).Interior.Color = RGB(255, 255, 0)
                contador = contador + 1
                noElegiblesArray(k) = noElegiblesArray(k) + 1

            End If

        Case "Hombres 20-64"

            If (InStr(1, codigosHombres, codigo) = 0) Then

                ActiveSheet.Cells(j, codigoPrestacionColumna).Interior.Color = RGB(255, 255, 0)
                contador = contador + 1
                noElegiblesArray(k) = noElegiblesArray(k) + 1

            End If

        Case "Niños 6-9"

            If (InStr(1, codigos6a9, codigo) = 0) Then

                ActiveSheet.Cells(j, codigoPrestacionColumna).Interior.Color = RGB(255, 255, 0)
                contador = contador + 1
                noElegiblesArray(k) = noElegiblesArray(k) + 1

            End If

        Case "Adolecentes 10-19"

            If ((InStr(1, codigosAdolescentes1, codigo) = 0) _
            And (InStr(1, codigosAdolescentes2, codigo) = 0)) Then

                If ((InStr(1, codigosAdolescentes1, codigoIzquierda) <> 0) _
                Or (InStr(1, codigosAdolescentes2, codigoIzquierda) <> 0)) Then

                    If (InStr(1, diagnosticosNoPermitidos, diagnostico) <> 0) Then

                        ActiveSheet.Cells(j, codigoPrestacionColumna).Interior.Color = RGB(255, 255, 0)
                        contador = contador + 1
                        noElegiblesArray(k) = noElegiblesArray(k) + 1

                    End If

                Else

                    ActiveSheet.Cells(j, codigoPrestacionColumna).Interior.Color = RGB(255, 255, 0)
                    contador = contador + 1
                    noElegiblesArray(k) = noElegiblesArray(k) + 1

                End If

            End If

        Case "Niños 0-5"

            'cargo la siguiente variable para reducir escritura
            edad = (ActiveSheet.Cells(j, fechaPrestacionColumna).Value - ActiveSheet.Cells(j, fechaNacimientoColumna).Value) / 365

            If (edad >= 0 And edad < 1) Then

                If (InStr(1, codigos0a1, codigo) = 0) Then

                    ActiveSheet.Cells(j, codigoPrestacionColumna).Interior.Color = RGB(255, 255, 0)
                    contador = contador + 1
                    noElegiblesArray(k) = noElegiblesArray(k) + 1

                End If

            ElseIf (edad >= 1 And edad < 2) Then

                If (InStr(1, codigos1a2, codigo) = 0) Then

                    ActiveSheet.Cells(j, codigoPrestacionColumna).Interior.Color = RGB(255, 255, 0)
                    contador = contador + 1
                    noElegiblesArray(k) = noElegiblesArray(k) + 1

                End If

            Else

                If (InStr(1, codigos2a6, codigo) = 0) Then

                    ActiveSheet.Cells(j, codigoPrestacionColumna).Interior.Color = RGB(255, 255, 0)
                    contador = contador + 1
                    noElegiblesArray(k) = noElegiblesArray(k) + 1

                End If

            End If

    End Select
    
    j = j + 1
    
Loop

Call cuadrosMuestraPadron(cuieArray, cantidadMuestraArray, provinciaArray, muestraArray, n, noElegiblesArray, validosXcuie, contador)

Call preparacionFinal

End Sub
