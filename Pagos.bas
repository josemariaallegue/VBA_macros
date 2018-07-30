Attribute VB_Name = "Pagos"
Option Explicit

Sub revision_muestra_pagos()

Call preparacionInicio

Dim i, j, k, m, n As Integer
Dim cuieColumna, codigoPrestacionColumna, cantidadMuestraColumna, muestraColumna, beneficiariosValidosColumna As Integer
Dim cuieArray(1 To 15, 1 To 5), mensaje, codigos As String
Dim contador1, contador2, CONTADOR3 As Integer
Dim auxiliar As Variant

i = 1
j = 2
k = 1

codigos = "APA001A98;APA001W78;APA001X75;APA001X86;APA002X75;APA002X76;APA002X80;APA002A98;CTC001A97;CTC001B80;CTC001D11;CTC001R74;CTC001R78;CTC001T91;CTC002A97;CTC002R96;CTC002T82;CTC002T91;CTC005B80;CTC005W78;CTC006W78;CTC007W84;CTC008A97;CTC010A97;CTC010W78;CTC011A97;CTC012A03;CTC012A81;CTC012A92;CTC012D01;CTC012D10;CTC012H71;CTC012H72;CTC012H76;CTC012L72;CTC012L73;CTC012L74;CTC012L77;CTC012L78;CTC012L80;CTC012R72;CTC012R77;CTC012R80;CTC012S13;CTC012T11;IMV001A98;IMV002A98;IMV003A98;IMV004A98;IMV005A98;IMV006A98;IMV007A98;IMV013A98;IMV014A98;ITE002A40;ITE002A41;ITE002A42;ITE002A44;ITE002R78;ITQ001W90;ITQ001W91;ITQ002W88;ITQ002W89;ITQ005W06;ITQ006W07;ITQ007W08;TAT001A98;TAT002A98;TAT003A98;TAT007A98;TAT008A98;TAT009A98;TAT010A98;TAT013A98;TAT014A98;CTC001T79;CTC001R96;CTC001R81"

'selecciona la solapa "Database" para que no hay errores de ejecucion
Worksheets("Database").Activate

'recorrido horizontal de las columnas
Do Until ActiveSheet.Cells(1, i).Value = ""

    'guardo el valor de la columna de "CUIE_EFECTOR"
    If (ActiveSheet.Cells(1, i).Value = "CUIE_EFECTOR") Then
        
        cuieColumna = i
        
    End If
    
    'guardo el valor de la columna de "CODIGO_PRESTACION"
    If (ActiveSheet.Cells(1, i).Value = "CODIGO_PRESTACION") Then
    
        codigoPrestacionColumna = i
        
    End If
    
    'guardo el valor de "N"
    If (ActiveSheet.Cells(1, i).Value = "N") Then
    
        n = ActiveSheet.Cells(2, i).Value
        
    End If
    
    'guardo el valor de la columna de "MUESTRA"
    If (ActiveSheet.Cells(1, i).Value = "MUESTRA") Then
        
        muestraColumna = i
    
    End If
    
    'guardo el valor de la columna de "CANTIDAD_MUESTRA"
    If (ActiveSheet.Cells(1, i).Value = "CANTIDAD_MUESTRA") Then
        
        cantidadMuestraColumna = i
        
    End If
    
    'guardo el valor de la columna de "CUIE_X_BENEF_VALIDOS"
    If (ActiveSheet.Cells(1, i).Value = "CUIE_X_BENEF_VALIDOS") Then
        
        beneficiariosValidosColumna = i
        
    End If
  
    i = i + 1
    
Loop

    
'recorrido vertical
Do Until ActiveSheet.Cells(j, 1).Value = ""
    
    'guardo la muestraArray para cada cuie
    'en el contador se guardan la cantidad de efectores tomados
    If (ActiveSheet.Cells(j, cuieColumna).Value <> ActiveSheet.Cells(j - 1, cuieColumna).Value) Then

        cuieArray(k, 1) = ActiveSheet.Cells(j, cuieColumna).Value
        cuieArray(k, 2) = ActiveSheet.Cells(j, beneficiariosValidosColumna).Value
        cuieArray(k, 3) = ActiveSheet.Cells(j, cantidadMuestraColumna).Value
        
        
        
        contador1 = contador1 + 1

        k = k + 1

    End If
    
    'reviso que los codigos seleccionados sean los del listado de codigos elegibles
    'pinto la celda del codigo invalido
    'cuento la cantidad de codigos invalidos en contador2
    If (InStr(1, codigos, ActiveSheet.Cells(j, codigoPrestacionColumna).Value) = 0) Then
    
        ActiveSheet.Cells(j, codigoPrestacionColumna).Interior.Color = RGB(255, 255, 0)
        contador2 = contador2 + 1
        cuieArray(k - 1, 5) = cuieArray(k - 1, 5) + 1
        
    End If
    
    'cuento la cantidad de casos que tienen x en CONTADOR3
    On Error GoTo sinColumnaMuestra
    If (muestraColumna <> 0 And LCase(ActiveSheet.Cells(j, muestraColumna).Value) = "x") Then

        CONTADOR3 = CONTADOR3 + 1
        
        For m = 1 To 12
            
            'cuento la cantidad de casos seleccionados por efector
            If (cuieArray(m, 1) = ActiveSheet.Cells(j, cuieColumna).Value) Then
                
                    cuieArray(m, 4) = cuieArray(m, 4) + 1
                
            End If
            
        Next m

    End If
    
    j = j + 1
    
Loop

Call cuadrosMuestraPagos(cuieArray, n, contador1, contador2, CONTADOR3)

Exit Sub

'excepcion por si no esta la columna "Muestra
sinColumnaMuestra:

muestraColumna = 0

Resume Next

Call preparacionFinal

End Sub
