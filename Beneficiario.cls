VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Beneficiario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'declaracion de atributos
Private pNombre, pApellido As String
Private pEdad As Double
Private pFechaPrestacion, pFechaNacimiento As Date
Private pClaveBeneficiario, pTipoDoc, pDocumento, pSexo, pCodigoPrestacion As String
Private pPeso, pTalla, pTa, pPerimetroCefalico, pSemanasEmbarazo, pIndiceOdonto, pResultadoOto As String
Private pResultadoRetino, pBiopsiaMamaria, pBiopsiaCervico, pLecturaPap, pMamografia, pVdrl, pTratamientoInstaurado As String

'reusmen: metodo de inicializacion de variables
'parametros: una coleccion con el numero de la columna de donde tiene que tomar el dato
'            la fila donde se encuentra el beneficiario
'retorno: void
Public Sub Init(ByRef coll As Collection, ByVal fila As Double)

    pNombre = ActiveSheet.Cells(fila, coll("NOMBRE_BENEFICIARIO"))
    pApellido = ActiveSheet.Cells(fila, coll("APELLIDO_BENEFICIARIO"))
    pFechaPrestacion = ActiveSheet.Cells(fila, coll("FECHA_PRESTACION"))
    pFechaNacimiento = ActiveSheet.Cells(fila, coll("FECHA_DE_NACIMIENTO"))
    pClaveBeneficiario = ActiveSheet.Cells(fila, coll("CLAVE_BENEFICIARIO"))
    pTipoDoc = ActiveSheet.Cells(fila, coll("TIPO_DOC"))
    pDocumento = ActiveSheet.Cells(fila, coll("BENEF_NRO_DOCUMENTO"))
    pSexo = UCase(Left(ActiveSheet.Cells(fila, coll("SEXO")), 1))
    pCodigoPrestacion = ActiveSheet.Cells(fila, coll("CODIGO_PRESTACION"))
    pPeso = ActiveSheet.Cells(fila, coll("PESO"))
    pTalla = ActiveSheet.Cells(fila, coll("TALLA"))
    pTa = ActiveSheet.Cells(fila, coll("TENSION_ARTERIAL"))
    pPerimetroCefalico = ActiveSheet.Cells(fila, coll("PERIMETRO_CEFALICO"))
    pSemanasEmbarazo = ActiveSheet.Cells(fila, coll("SEMANAS_EMBARAZO"))
    pIndiceOdonto = ActiveSheet.Cells(fila, coll("INDICE_ODONTO"))
    pResultadoOto = ActiveSheet.Cells(fila, coll("RESULTADO_OTO"))
    pResultadoRetino = ActiveSheet.Cells(fila, coll("RESULTADO_RETINO"))
    pBiopsiaMamaria = ActiveSheet.Cells(fila, coll("BIOPSIA_MAMA"))
    pBiopsiaCervico = ActiveSheet.Cells(fila, coll("BIOPSIA_CERVICO"))
    pLecturaPap = ActiveSheet.Cells(fila, coll("LECTURA_PAP"))
    pMamografia = ActiveSheet.Cells(fila, coll("MAMOGRAFIA"))
    pVdrl = ActiveSheet.Cells(fila, coll("VDRL"))
    pTratamientoInstaurado = ActiveSheet.Cells(fila, coll("TRAT_INSTAURADO"))
    pEdad = (pFechaPrestacion - pFechaNacimiento) / 365
    
End Sub

'resumen: le da a una variable los valores mas importantes del beneficiario
'parametros: void
'retorno: una cadena con los datos del beneficiario
Public Function Mostrar() As String
    
    Dim texto As String
    
    texto = texto & "Nombre: " & pNombre & vbNewLine
    texto = texto & "Apellido: " & pApellido & vbNewLine
    texto = texto & "Clave de beneficiario: " & pClaveBeneficiario & vbNewLine
    texto = texto & "Documento: " & pDocumento & vbNewLine
    texto = texto & "Sexo: " & pSexo & vbNewLine
    texto = texto & "Edad a la fecha de prestacion: " & pEdad & vbNewLine
    texto = texto & "Fecha de nacimiento: " & pFechaNacimiento & vbNewLine
    texto = texto & "Fecha de prestacion: " & pFechaPrestacion & vbNewLine
    texto = texto & "Codigo de prestación: " & pCodigoPrestacion & vbNewLine
    
    Mostrar = texto
    
End Function

'resumen: analiza si requiere datos
'parametros: void
'retorno: 0 si no requiere datos reportables
'         1x si requiere para niños de 0 a 5
'         2x si requiere para niños de 6 a9
'         3x si requiere para adolescentes 10 a 19
'         4x si requiere para mujeres 20 a 64
Private Function requiereDatosReportables() As Integer

Dim codigosAdolescentes1, codigosAdolescentes2, codigosAdolescentes3 As String
Dim codigosMujeres1, codigosMujeres2, codigosMujeres3, codigosMujeres4, codigosMujeres5 As String
Dim codigosMujeres6, codigosMujeres7, codigosMujeres8, codigosMujeres9 As String
Dim codigosNiños0a51, codigosNiños0a52, codigosNiños0a53, codigosNiños0a54 As String
Dim codigosNiños6a91, codigosNiños6a92 As String


codigosNiños0a51 = "PRP017A46;PRP017A97"
codigosNiños0a52 = "PRP021A97;PRP021H86"
codigosNiños0a53 = "CTC001A97"
codigosNiños0a54 = "CTC010A97"
codigosNiños6a91 = "CTC001A97;CTC001T79;CTC001T82;CTC001T83;CTC002T79;CTC002T82;CTC002T83;CTC009A97"
codigosNiños6a92 = "CTC010A97"
codigosAdolescentes1 = "CTC001A97;CTC001T79;CTC001T82;CTC001T83;CTC002T79;CTC002T82;CTC002T83;CTC009A97"
codigosAdolescentes2 = "CTC005W78;CTC006W78;CTC007O10;CTC007O10.4;CTC007O16;CTC007O24.4;CTC017P05;CTC022O10;CTC022O10.4;CTC022O16;CTC022O24.4"
codigosAdolescentes3 = "CTC010A97;CTC010W78"
codigosMujeres1 = "CTC005W78;CTC006W78;CTC007O10;CTC007O10.4;CTC007O16;CTC007O24.4;CTC017P05;CTC022O10;CTC022O10.4;CTC022O16;CTC022O24.4"
codigosMujeres2 = "APA001A98;APA001X75;APA001X86"
codigosMujeres3 = "APA002A98;APA002X75;APA002X80"
codigosMujeres4 = "CTC001A97;CTC009A97"
codigosMujeres5 = "CTC010A97;CTC010W78"
codigosMujeres6 = "LBL119A97;LBL119W78"
codigosMujeres7 = "APA002X76;APA002X79"
codigosMujeres8 = "IGR014A98"
codigosMujeres9 = "NTN002X75"


'con estos if si el codigo requiere datos reportables y para que poblacion
If (pEdad >= 0 And pEdad < 6) Then
    
    If (InStr(1, codigosNiños0a51, pCodigoPrestacion) <> 0) Then
        
        requiereDatosReportables = 11
        Exit Function
        
    ElseIf (InStr(1, codigosNiños0a52, pCodigoPrestacion) <> 0) Then
        
        requiereDatosReportables = 12
        Exit Function
        
    ElseIf (InStr(1, codigosNiños0a53, pCodigoPrestacion) <> 0) Then
        
        requiereDatosReportables = 13
        Exit Function
    
    ElseIf (InStr(1, codigosNiños0a54, pCodigoPrestacion) <> 0) Then
        
        requiereDatosReportables = 14
        Exit Function
        
    End If

ElseIf (pEdad >= 6 And pEdad < 10) Then

    If (InStr(1, codigosNiños6a91, pCodigoPrestacion) <> 0) Then
        
        requiereDatosReportables = 21
        Exit Function
        
    ElseIf (InStr(1, codigosNiños6a92, pCodigoPrestacion) <> 0) Then
        
        requiereDatosReportables = 22
        Exit Function
        
    End If

ElseIf (pEdad >= 10 And pEdad < 20) Then

    If (InStr(1, codigosAdolescentes1, pCodigoPrestacion) <> 0) Then
            
        requiereDatosReportables = 31
        Exit Function
            
    ElseIf (InStr(1, codigosAdolescentes2, pCodigoPrestacion) <> 0) Then
    
        requiereDatosReportables = 32
        Exit Function
        
    ElseIf (InStr(1, codigosAdolescentes3, pCodigoPrestacion) <> 0) Then
    
        requiereDatosReportables = 33
        Exit Function
    
    End If
    
ElseIf (pEdad >= 20 And pSexo = "F") Then

    If (InStr(1, codigosMujeres1, pCodigoPrestacion) <> 0) Then
        
        requiereDatosReportables = 41
        Exit Function
        
    ElseIf (InStr(1, codigosMujeres2, pCodigoPrestacion) <> 0) Then
        
        requiereDatosReportables = 42
        Exit Function
        
    ElseIf (InStr(1, codigosMujeres3, pCodigoPrestacion) <> 0) Then
        
        requiereDatosReportables = 43
        Exit Function
        
    ElseIf (InStr(1, codigosMujeres4, pCodigoPrestacion) <> 0) Then
        
        requiereDatosReportables = 44
        Exit Function
        
    ElseIf (InStr(1, codigosMujeres5, pCodigoPrestacion) <> 0) Then
        
        requiereDatosReportables = 45
        Exit Function
        
    ElseIf (InStr(1, codigosMujeres6, pCodigoPrestacion) <> 0) Then
        
        requiereDatosReportables = 46
        Exit Function
        
    ElseIf (InStr(1, codigosMujeres7, pCodigoPrestacion) <> 0) Then
        
        requiereDatosReportables = 47
        Exit Function
        
    ElseIf (InStr(1, codigosMujeres8, pCodigoPrestacion) <> 0) Then
        
        requiereDatosReportables = 48
        Exit Function
        
    ElseIf (InStr(1, codigosMujeres9, pCodigoPrestacion) <> 0) Then
        
        requiereDatosReportables = 49
        Exit Function
        
    End If
    
Else

    requiereDatosReportables = 0

End If
    
End Function

Public Function analisisDatosReportablesMarzo() As String

Dim fundamento As String

Select Case (requiereDatosReportables)

    Case 11
        
        
        
    
    Case 12
    
    Case 13
    
    Case 14
    
    Case 21
    
    Case 22
    
    Case 31
    
    Case 32
    
    Case 33
    
    Case 41
    
    Case 42
    
    Case 43
    
    Case 44
    
    Case 45
    
    Case 46
    
    Case 47
    
    Case 48
    
    Case 49

End Select

End Function


