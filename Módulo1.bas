Attribute VB_Name = "Módulo1"
Option Explicit

Private Sub analisisDatosReportables()

Call preparacionInicio

Dim i, j, k, fila As Integer
Dim edad As Double
Dim beneficiario1 As New Beneficiario
Dim coll As New Collection


Call columnas(coll)

For fila = 2 To filaMaxActiva(columnaMaxActiva)
    
    beneficiario1 = Nothing
    beneficiario1.Init coll, fila
    
    
Next fila

End Sub

Sub columnas(ByRef coll As Collection)

Dim i, j, k As Integer
Dim nombreColumna As String
Dim celda As Variant

i = 1

Worksheets("Database").Activate

For Each celda In ActiveSheet.Range("A1:CA1")

    nombreColumna = celda.Value
    
    If (nombreColumna = "NOMBRE_BENEFICIARIO") Then
        
        coll.Add celda.Column, "NOMBRE_BENEFICIARIO"
    
    ElseIf (nombreColumna = "APELLIDO_BENEFICIARIO") Then
        
        coll.Add celda.Column, "APELLIDO_BENEFICIARIO"
    
    ElseIf (nombreColumna = "FECHA_PRESTACION") Then
        
        coll.Add celda.Column, "FECHA_PRESTACION"
    
    ElseIf (nombreColumna = "FECHA_DE_NACIMIENTO") Then
        
        coll.Add celda.Column, "FECHA_DE_NACIMIENTO"

    ElseIf (nombreColumna = "CLAVE_BENEFICIARIO") Then
        
        coll.Add celda.Column, "CLAVE_BENEFICIARIO"

    ElseIf (nombreColumna = "TIPO_DOC") Then
        
        coll.Add celda.Column, "TIPO_DOC"
        
    ElseIf (nombreColumna = "BENEF_NRO_DOCUMENTO") Then
        
        coll.Add celda.Column, "BENEF_NRO_DOCUMENTO"
    
    ElseIf (nombreColumna = "SEXO") Then
        
        coll.Add celda.Column, "SEXO"
    
    ElseIf (nombreColumna = "CODIGO_PRESTACION") Then
        
        coll.Add celda.Column, "CODIGO_PRESTACION"
    
    ElseIf (nombreColumna = "PESO") Then
        
        coll.Add celda.Column, "PESO"
        
    ElseIf (nombreColumna = "TALLA") Then
    
        coll.Add celda.Column, "TALLA"
    
    ElseIf (nombreColumna = "TENSION_ARTERIAL") Then

        coll.Add celda.Column, "TENSION_ARTERIAL"
    
    ElseIf (nombreColumna = "PERIMETRO_CEFALICO") Then

        coll.Add celda.Column, "PERIMETRO_CEFALICO"
    
    ElseIf (nombreColumna = "SEMANAS_EMBARAZO") Then

        coll.Add celda.Column, "SEMANAS_EMBARAZO"
    
    ElseIf (nombreColumna = "INDICE_ODONTO") Then

        coll.Add celda.Column, "INDICE_ODONTO"
    
    ElseIf (nombreColumna = "RESULTADO_OTO") Then

        coll.Add celda.Column, "RESULTADO_OTO"
    
    ElseIf (nombreColumna = "RESULTADO_RETINO") Then

        coll.Add celda.Column, "RESULTADO_RETINO"
    
    ElseIf (nombreColumna = "BIOPSIA_MAMA") Then

        coll.Add celda.Column, "BIOPSIA_MAMA"
    
    ElseIf (nombreColumna = "BIOPSIA_CERVICO") Then

        coll.Add celda.Column, "BIOPSIA_CERVICO"
    
    ElseIf (nombreColumna = "LECTURA_PAP") Then

        coll.Add celda.Column, "LECTURA_PAP"
        
    ElseIf (nombreColumna = "MAMOGRAFIA") Then

        coll.Add celda.Column, "MAMOGRAFIA"
        
    ElseIf (nombreColumna = "VDRL") Then

        coll.Add celda.Column, "VDRL"
        
    ElseIf (nombreColumna = "TRAT_INSTAURADO") Then

        coll.Add celda.Column, "TRAT_INSTAURADO"
        
    End If

Next celda


'Do Until (ActiveSheet.Cells(1, i).Value = "")
'
'    nombreColumna = ActiveSheet.Cells(1, i).Value
'
'
'    If (nombreColumna = "APELLIDO_BENEFICIARIO") Then
'
'        coll.Add i, "APELLIDO_BENEFICIARIO"
'
'    ElseIf (nombreColumna = "FECHA_PRESTACION") Then
'
'        coll.Add i, "FECHA_PRESTACION"
'
'    ElseIf (nombreColumna = "FECHA_DE_NACIMIENTO") Then
'
'        coll.Add i, "FECHA_DE_NACIMIENTO"
'
'    ElseIf (nombreColumna = "CLAVE_BENEFICIARIO") Then
'
'        coll.Add i, "CLAVE_BENEFICIARIO"
'
'    ElseIf (nombreColumna = "TIPO_DOC") Then
'
'        coll.Add i, "TIPO_DOC"
'
'    ElseIf (nombreColumna = "SEXO") Then
'
'        coll.Add i, "SEXO"
'
'    ElseIf (nombreColumna = "CODIGO_PRESTACION") Then
'
'        coll.Add i, "CODIGO_PRESTACION"
'
'    ElseIf (nombreColumna = "PESO") Then
'
'        coll.Add i, "PESO"
'
'    ElseIf (nombreColumna = "TALLA") Then
'
'        coll.Add i, "TALLA"
'
'    ElseIf (nombreColumna = "TENSION_ARTERIAL") Then
'
'        coll.Add i, "TENSION_ARTERIAL"
'
'    ElseIf (nombreColumna = "PERIMETRO_CEFALICO") Then
'
'        coll.Add i, "PERIMETRO_CEFALICO"
'
'    ElseIf (nombreColumna = "SEMANAS_EMBARAZO") Then
'
'        coll.Add i, "SEMANAS_EMBARAZO"
'
'    ElseIf (nombreColumna = "INDICE_ODONTO") Then
'
'        coll.Add i, "INDICE_ODONTO"
'
'    ElseIf (nombreColumna = "RESULTADO_OTO") Then
'
'        coll.Add i, "RESULTADO_OTO"
'
'    ElseIf (nombreColumna = "RESULTADO_RETINO") Then
'
'        coll.Add i, "RESULTADO_RETINO"
'
'    ElseIf (nombreColumna = "BIOPSIA_MAMA") Then
'
'        coll.Add i, "BIOPSIA_MAMA"
'
'    ElseIf (nombreColumna = "BIOPSIA_CERVICO") Then
'
'        coll.Add i, "BIOPSIA_CERVICO"
'
'    ElseIf (nombreColumna = "LECTURA_PAP") Then
'
'        coll.Add i, "LECTURA_PAP"
'
'    ElseIf (nombreColumna = "MAMOGRAFIA") Then
'
'        coll.Add i, "MAMOGRAFIA"
'
'    ElseIf (nombreColumna = "VDRL") Then
'
'        coll.Add i, "VDRL"
'
'    ElseIf (nombreColumna = "TRAT_INSTAURADO") Then
'
'        coll.Add i, "TRAT_INSTAURADO"
'
'    End If
'
'
'    i = i + 1
'
'Loop

End Sub
