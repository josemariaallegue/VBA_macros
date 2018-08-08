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
Public Sub Init(ByRef coll As Collection, ByVal fila As Integer)

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
Private Function requiereDatosReportables()

Dim flag As Boolean
Dim codigosAdolescentes1, codigosAdolescentes2, codigosAdolescentes3 As String


codigosAdolescentes1 = "CTC001A97;CTC001T79;CTC001T82;CTC001T83;CTC002T79;CTC002T82;CTC002T83;CTC009A97"
codigosAdolescentes2 = "CTC005W78;CTC006W78;CTC007O10;CTC007O10.4;CTC007O16;CTC007O24.4;CTC017P05;CTC022O10;CTC022O10.4;CTC022O16;CTC022O24.4"
codigosAdolescentes3 = "CTC010A97;CTC010W78"


If (pEdad >= 10 And pEdad < 20) Then

    If (InStr(1, codigosAdolescentes1, pCodigoPrestacion) <> 0 Or InStr(1, codigosAdolescentes2, pCodigoPrestacion) <> 0 _
    Or InStr(1, codigosAdolescentes3, pCodigoPrestacion) <> 0) Then
            
            flag = True
    End If
ElseIf (pEdad >= 20 And pSexo = "F") Then

ElseIf (pEdad >= 0 And pEdad < 6) Then

ElseIf (pEdad >= 6 And pEdad < 10) Then

Else

End If
    
        
        


End Function

