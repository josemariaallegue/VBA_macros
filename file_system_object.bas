Attribute VB_Name = "file_system_object"
Option Explicit

Public Function copiarArchivo(rutaOrigen, rutaDestino, nombreArch, extension)

Dim fso As Object
Dim archivoAux As File
Dim carpetaAux As Folder

Set fso = CreateObject("Scripting.FileSystemObject")
Set carpetaAux = fso.GetFolder(rutaOrigen)

For Each archivoAux In carpetaAux.Files
    
    If (InStr(1, archivoAux.name, nombreArch) > 0) Then
    
        fso.CopyFile Source:=archivoAux.Path, Destination:=rutaDestino & "\" & nombreArch & extension
        Exit For
    
    End If
    
Next archivoAux


End Function

Public Function chequearEspacioDisco()

Dim fso As FileSystemObject
Dim disco As Drive
Dim space As Double

Set fso = New FileSystemObject
Set disco = fso.GetDrive("C:")

space = disco.AvailableSpace
space = space / 1073741824
space = WorksheetFunction.Round(space, 2)
MsgBox "C: has free space = " & space & " GB"

End Function

Public Function existenciaCarpeta()

Dim fso As FileSystemObject
Dim carpetaNombre As String

Set fso = New FileSystemObject

carpetaNombre = InputBox("Ingrese el nombre de la carpete a checkear: ")

If (Len(carpetaNombre) > 0) Then
    
    If (fso.FolderExists(carpetaNombre) = True) Then
        
        MsgBox ("La carpte existe")
        
    Else
    
        fso.CreateFolder (carpetaNombre)
        MsgBox ("Carpeta creada")
        
    End If
    
Else

    MsgBox ("Valores incorrectos")
    
End If

End Function

Public Function copiarCarpeta()

Dim fso As FileSystemObject

Set fso = New FileSystemObject

fso.CopyFolder "C:\Users\jmallegue\Documents\Certificados\Banco", _
"C:\Users\jmallegue\Documents\Certificados", True

End Function

Public Function getCarpetasEspeciales()

Dim fso As FileSystemObject
Dim windowsCarpeta As String, systemCarpeta As String, tempCarpeta As String

Set fso = New FileSystemObject

windowsCarpeta = fso.GetSpecialFolder(0)
systemCarpeta = fso.GetSpecialFolder(1)
tempCarpeta = fso.GetSpecialFolder(2)

MsgBox ("Carpeta windows ruta: " & windowsCarpeta & vbNewLine & _
"Carpeta system ruta: " & systemCarpeta & vbNewLine & _
"Carpeta temp ruta: " & tempCarpeta)

End Function

Public Function crearArchivo()

Dim fso As FileSystemObject
Dim txtStr As TextStream
Dim nombre As String, contenido As String
Dim archivo As File
Dim i As Double

Set fso = New FileSystemObject

nombre = "C:\Users\jmallegue\Documents\Office\Sumar\File.txt"

contenido = InputBox("Ingrese el contenido del archivo")

If (Len(contenido) > 0) Then
    
    Set txtStr = fso.CreateTextFile(nombre, True, True)
    txtStr.Write (contenido)
    txtStr.Close
    
End If

If (fso.FileExists(nombre)) Then
    
    Set archivo = fso.GetFile(nombre)
    Set txtStr = archivo.OpenAsTextStream(ForReading, TristateUseDefault)
    MsgBox (txtStr.ReadAll)
    txtStr.Close
    archivo.Delete
    
End If

End Function

Public Function listarArchivos()

Dim fso As FileSystemObject
Dim carpeta As Folder
Dim archivo As File
Dim ruta As String
Dim siguienteFila As Long

Set fso = New FileSystemObject
ruta = "C:\Users\jmallegue\Documents\Certificados\Office"

Set carpeta = fso.GetFolder(ruta)

If (carpeta.Files.Count = 0) Then
    
    MsgBox ("Sin archivos")
    Exit Function
    
End If

Cells(1, "A").Value = "Nombre del archivo"
Cells(1, "B").Value = "Tamaño"
Cells(1, "C").Value = "Fecha de modificacion"

siguienteFila = ActiveSheet.Cells(Rows.Count, "A").End(xlUp).Row + 1

For Each archivo In carpeta.Files
    
    Cells(siguienteFila, 1).Value = archivo.name
    Cells(siguienteFila, 2).Value = archivo.Size
    Cells(siguienteFila, 3).Value = archivo.DateLastModified
    
    siguienteFila = siguienteFila + 1
    
Next archivo

End Function





