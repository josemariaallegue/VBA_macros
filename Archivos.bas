Attribute VB_Name = "Archivos"
Option Explicit
Option Private Module

'resumen: devuelvo un archivo seleccionado a travez de una ventana
'parametros: void
'retorno: el archivo seleccionado
Public Function obtenerArchivo() As Workbook

Dim fd As Office.FileDialog

Set fd = Application.FileDialog(msoFileDialogOpen)

   With fd
   
      .AllowMultiSelect = False

      'pongo un titulo a la venta
      .Title = "Seleccione un archivo."

      'quito los filtros
      .Filters.Clear

      ' muestro la venta. Si .Show devuel verdadero, el
      ' usuario eligio un archivo. Si devuelve
      ' falso, el usuario selecciono cancelar
      If (.Show = True) Then
        
        'le otorgo a la variable de retorno el archivo que abrio
        Set obtenerArchivo = Workbooks.Open(nombreArchivo(fd.SelectedItems(1)))

      End If
      
   End With

End Function

'resumen: recive una cadena con la ruta completa del archivo quita lo que esta de mas
'parametros: una cade con la ruta del archivo
'retorno: el nombre y la extencion del archivo
Public Function nombreArchivo(ByVal ruta As String) As String

Dim largoRuta, nombre As Long

largoRuta = Len(ruta)

Do Until largoRuta = 0

    'se revisa de derecha a izquierda la ruta hasta encontrar el primer "\"
    'que es donde empieza el nombre
    If (InStr(largoRuta, ruta, "\") <> 0) Then
            
        'copio el nombre a la variable de retorno
        nombreArchivo = Right(ruta, Len(ruta) - largoRuta)
        
        'corto el loop
        largoRuta = 1
        
    End If
    
    largoRuta = largoRuta - 1

Loop

End Function
