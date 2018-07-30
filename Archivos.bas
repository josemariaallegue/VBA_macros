Attribute VB_Name = "Archivos"
Option Explicit
Option Private Module

'abro un archivo y retorno el nombre y extension de mismo
Public Function obtenerArchivo() As Workbook

Dim fd As Office.FileDialog

Set fd = Application.FileDialog(msoFileDialogOpen)

   With fd

      .AllowMultiSelect = False

      ' Set the title of the dialog box
      .Title = "Please select the file."

      ' Clear out the current filters
      .Filters.Clear

      ' Show the dialog box. If the .Show method returns True, the
      ' user picked at least one file. If the .Show method returns
      ' False, the user clicked Cancel
      If (.Show = True) Then
        
        'le otorgo a la variable de retorno el archivo que abrio
        Set obtenerArchivo = Workbooks.Open(nombreArchivo(fd.SelectedItems(1)))

      End If
      
   End With

End Function

'retorna el nombre y extension de un archivo
'recibe la ruta del archivo
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
