VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} usfmCompararRegistros 
   Caption         =   "Comparar cantidad registros"
   ClientHeight    =   1980
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6495
   OleObjectBlob   =   "usfmCompararRegistros.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "usfmCompararRegistros"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnContinuar_Click()

Dim i, j, columnaHoja1, filas As Long
Dim registrosHoja1, registrosHoja2 As Double
Dim rango1, rango2 As Range
Dim hoja1, hoja2 As Worksheet
Dim resumen() As String
Dim auxHora As Date

'preparo el archivo
Call preparacionInicio

'otorgo valores a varias varibales
auxHora = Now
Set hoja1 = ActiveWorkbook.Sheets(usfmCompararRegistros.cmbxHoja1.Text)
Set hoja2 = ActiveWorkbook.Sheets(usfmCompararRegistros.cmbxHoja2.Text)
columnaHoja1 = columnaMax(hoja1)
ReDim resumen(1 To columnaHoja1, 1 To 4)

'convierto el formato de las hojas en general porque si no uno de los if con la condicion
'"(hoja2.Cells(j, i).Value = "")" no funciona
hoja1.Cells.NumberFormat = "General"
hoja2.Cells.NumberFormat = "General"

'recorro las columnas para obtener la cantidad de filas maximas
filas = filaMax(columnaMax(hoja1), hoja1)

'recorro nuevamente las columnas
For i = 1 To columnaHoja1

    'recorro la totalidad de las filas de la columna i
    For j = 1 To filas

        'este if sirve para limpiar las casillas en blanco que devuelve IDEA que excel no considera vacias
        If (hoja2.Cells(j, i).Value = "") Then

            hoja2.Cells(j, i).Value = ""

        End If

    Next j

    'otorgo valores a distintas variables
    'a registrosHoja1 y registrosHoja2 se les resta 1 porque si no cuentra el encabezado de la columna
    On Error Resume Next
    Set rango1 = hoja1.Range(hoja1.Cells(1, i), hoja1.Cells(filas, i))
    Set rango2 = hoja2.Range(hoja2.Cells(1, i), hoja2.Cells(filas, i))
    registrosHoja1 = rango1.Cells.SpecialCells(xlCellTypeConstants).Count - 1
    registrosHoja2 = rango2.Cells.SpecialCells(xlCellTypeConstants).Count - 1

    'completo los valores de la variable "resumen" para asi armar la solapa "Resumen"
    resumen(i, 1) = hoja1.Cells(1, i).Value
    resumen(i, 2) = registrosHoja1
    resumen(i, 3) = registrosHoja2
    resumen(i, 4) = registrosHoja1 - registrosHoja2

Next i

'llamo a la funcion cuadrosComparacionCantidadRegistros y vuelvo a preparar el archivo
Call cuadrosComparacionCantidadRegistros(hoja1, hoja2, resumen, columnaHoja1)
Call preparacionFinal

End Sub

Private Sub cmbxHoja1_Change()

End Sub

Private Sub UserForm_Initialize()

Dim hoja As Worksheet
Dim i As Integer

Me.cmbxHoja1.Clear
Me.cmbxHoja2.Clear

For i = 1 To Sheets.Count
    
    Me.cmbxHoja1.AddItem Sheets(i).name
    Me.cmbxHoja2.AddItem Sheets(i).name
    
Next i

End Sub

