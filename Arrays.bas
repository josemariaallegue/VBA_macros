Attribute VB_Name = "Arrays"
Option Explicit
Option Private Module


Public Function largoArray(ByVal arr As Variant) As Integer

    largoArray = UBound(arr) - LBound(arr) + 1
    
End Function

Public Sub MostrarArray(ByVal arr As Variant)
    
    Dim i As Integer
    Dim auxiliar As Variant
    
    For i = 1 To UBound(arr)
        
        Debug.Print arr(i)
        
    Next i
    
    Exit Sub
    
End Sub

Sub MostrarMatriz2(ByVal arr As Variant)
    
    Dim i, j As Integer
    Dim auxiliar As Variant

    For i = LBound(arr, 1) To UBound(arr, 1)
    
        For j = LBound(arr, 2) To UBound(arr, 2)
        
            Debug.Print arr(i, j)
        
        Next j
        
    Next i
    
End Sub

