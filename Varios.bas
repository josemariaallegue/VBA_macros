Attribute VB_Name = "Varios"
Option Explicit

'resumen: compara la cantidad de registros entre 2 solapas seleccionadas
'parametros: void
'retorno: void
Public Sub comparar_cantidad_registros()

usfmCompararRegistros.Show

End Sub

'resumen: unifica un rango de celdas seleccionadas (separadas por ;) y las colaca en una celda seleccionada
'parametros: void
'retorno: void
Public Sub unificar_celdas()

Dim rango1, rango2 As Range
Dim celda As Variant
Dim texto As String
Dim flag As Boolean

flag = False
Set rango1 = Application.InputBox("Seleccione un rango", "Unificar celdas", Type:=8)
Set rango2 = Application.InputBox("Seleccione el destino", "Unificar celdas", Type:=8)

For Each celda In rango1
    
    If (flag = False) Then
        
        texto = celda.Value
        flag = True
    
    Else
    
        texto = texto & ";" & celda.Value
    
    End If

Next celda

rango2.Value = texto

End Sub
