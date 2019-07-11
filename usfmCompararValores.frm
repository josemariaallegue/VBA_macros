VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} usfmCompararValores 
   Caption         =   "UserForm1"
   ClientHeight    =   1980
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6525
   OleObjectBlob   =   "usfmCompararValores.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "usfmCompararValores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnContinuar_Click()

Dim i, j, columnasMax, filasMax, auxFila, auxColumna As Long
Dim rango1, rango2 As Range
Dim hoja1, hoja2 As Worksheet
Dim resumen() As String
Dim auxHora As Date

Call preparacionInicio

'otorgo valores a varias variables
auxHora = Now
Set hoja1 = ActiveWorkbook.Sheets(usfmCompararValores.cmbxHoja1.Text)
Set hoja2 = ActiveWorkbook.Sheets(usfmCompararValores.cmbxHoja2.Text)
columnasMax = columnaMax(hoja1)
filasMax = filaMax(columnaMax(hoja1), hoja1)
ReDim resumen(1 To columnasMax, 1 To 4)

Set rango1 = hoja1.Range(hoja1.Cells(1, 1), hoja1.Cells(filasMax, columnasMax))
Set rango2 = hoja2.Range(hoja2.Cells(1, 1), hoja2.Cells(filasMax, columnasMax))

'recorro el rango de la primera hoja
For auxFila = 1 To filasMax

    For auxColumna = 1 To columnasMax
        
        'si las celdas son distintas pinto ambas de amarillo
        If (rango1(auxFila, auxColumna) <> rango2(auxFila, auxColumna)) Then
            
            rango1(auxFila, auxColumna).Interior.Color = RGB(255, 255, 0)
            rango2(auxFila, auxColumna).Interior.Color = RGB(255, 255, 0)

        End If
        
    Next auxColumna
    
Next auxFila

Call preparacionFinal

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


