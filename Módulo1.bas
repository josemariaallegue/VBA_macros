Attribute VB_Name = "Módulo1"
Option Explicit

Sub revision_muestra()

Dim i, j, k, l As Integer
Dim CUIE_COLUMNA, N, CODIGO_PRESTACION, CANTIDAD_MUESTRA(1 To 12) As Integer
Dim CUIE(1 To 12), mensaje, codigos As String
Dim contador1, contador2, contador3 As Integer

i = 1
j = 2
k = 1
l = 2
codigos = "APA001A98;APA001W78;APA001X75;APA001X86;APA002X75;APA002X76;APA002X80;APA002A98;CTC001A97;CTC001B80;CTC001D11;CTC001R74;CTC001R78;CTC001T91;CTC002A97;CTC002R96;CTC002T82;CTC002T91;CTC005B80;CTC005W78;CTC006W78;CTC007W84;CTC008A97;CTC010A97;CTC010W78;CTC011A97;CTC012A03;CTC012A81;CTC012A92;CTC012D01;CTC012D10;CTC012H71;CTC012H72;CTC012H76;CTC012L72;CTC012L73;CTC012L74;CTC012L77;CTC012L78;CTC012L80;CTC012R72;CTC012R77;CTC012R80;CTC012S13;CTC012T11;IMV001A98;IMV002A98;IMV003A98;IMV004A98;IMV005A98;IMV006A98;IMV007A98;IMV013A98;IMV014A98;ITE002A40;ITE002A41;ITE002A42;ITE002A44;ITE002R78;ITQ001W90;ITQ001W91;ITQ002W88;ITQ002W89;ITQ005W06;ITQ006W07;ITQ007W08;TAT001A98;TAT002A98;TAT003A98;TAT007A98;TAT008A98;TAT009A98;TAT010A98;TAT013A98;TAT014A98;CTC001T79;CTC001R96;CTC001R81"
contador2 = 0


'recorrido horizontal de las columnas
Do Until ActiveSheet.Cells(1, i).Value = ""

    'guardo el valor de la columna de "CUIE"
    If (ActiveSheet.Cells(1, i).Value = "CUIE_EFECTOR") Then
        
        CUIE_COLUMNA = i
        
    End If
    
    'guardo el valor de la columna de "CODIGO_PRESTACION"
    If (ActiveSheet.Cells(1, i).Value = "CODIGO_PRESTACION") Then
    
        CODIGO_PRESTACION = i
        
    End If
    
    'guardo el valor de "N"
    If (ActiveSheet.Cells(1, i).Value = "N") Then
    
        N = ActiveSheet.Cells(2, i).Value
        
    End If
    
    'guardo la muestra para cada CUIE
    If (ActiveSheet.Cells(1, i).Value = "CANTIDAD_MUESTRA") Then
    
        Do Until ActiveSheet.Cells(j, i).Value = ""
            
            If (ActiveSheet.Cells(j, i).Value <> ActiveSheet.Cells(j - 1, i).Value) Then

                CUIE(k) = ActiveSheet.Cells(j, CUIE_COLUMNA).Value
                CANTIDAD_MUESTRA(k) = ActiveSheet.Cells(j, i).Value

                k = k + 1

            End If
            
            'reviso que los codigos seleccionados sean los del listado de codigos elegibles
            If (InStr(1, codigos, ActiveSheet.Cells(j, CODIGO_PRESTACION).Value) = 0) Then
            
                ActiveSheet.Cells(j, CODIGO_PRESTACION).Interior.Color = RGB(255, 255, 0)
                contador2 = contador2 + 1
                
            End If
            
            j = j + 1
            
        Loop
        
    End If

    i = i + 1
    
Loop

Do Until ActiveSheet.Cells(l, 1).Value = ""

    contador1 = contador1 + 1
    l = l + 1

Loop

If (contador1 <> N) Then

    mensaje = "El valor de la N es: " & N & " pero hay: " & contador1 & " por lo que hay " & (N - contador1) & " casos de menos." & vbCrLf
    mensaje = mensaje & "De estos " & contador1 & " casos " & contador2 & " es/son codigos no elegibles." & vbCrLf

End If

MsgBox mensaje
End Sub

Sub concatenar()

Dim i As Integer
Dim j As Integer
Dim texto As String

i = 1
j = 1

Do Until ActiveSheet.Cells(j, i).Value = ""

    texto = texto & ";" & ActiveSheet.Cells(j, i).Value
    j = j & 1

Loop

ActiveSheet.Cells(2, 2).Value = texto

End Sub

Sub instr()

Dim texto As String

texto = "APA001A98;APA001W78;APA001X75;APA001X86;APA002X75;APA002X76;APA002X80;APA002A98;CTC001A97;CTC001B80;CTC001D11;CTC001R74;CTC001R78;CTC001T91;CTC002A97;CTC002R96;CTC002T82;CTC002T91;CTC005B80;CTC005W78;CTC006W78;CTC007W84;CTC008A97;CTC010A97;CTC010W78;CTC011A97;CTC012A03;CTC012A81;CTC012A92;CTC012D01;CTC012D10;CTC012H71;CTC012H72;CTC012H76;CTC012L72;CTC012L73;CTC012L74;CTC012L77;CTC012L78;CTC012L80;CTC012R72;CTC012R77;CTC012R80;CTC012S13;CTC012T11;IMV001A98;IMV002A98;IMV003A98;IMV004A98;IMV005A98;IMV006A98;IMV007A98;IMV013A98;IMV014A98;ITE002A40;ITE002A41;ITE002A42;ITE002A44;ITE002R78;ITQ001W90;ITQ001W91;ITQ002W88;ITQ002W89;ITQ005W06;ITQ006W07;ITQ007W08;TAT001A98;TAT002A98;TAT003A98;TAT007A98;TAT008A98;TAT009A98;TAT010A98;TAT013A98;TAT014A98;CTC001T79;CTC001R96;CTC001R81"

If (InStr(1, texto, "CTC001A99") = 0) Then

    MsgBox "entro"

End If

MsgBox "final"
End Sub


