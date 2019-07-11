Attribute VB_Name = "Arrays_colecciones"
Option Explicit
Option Private Module

'resumen: cuenta la cantidad de elementos de un array
'parametro: un array
'retorno: devuelve un entero con la cantidad de elementos del array
Public Function largoArray(ByVal arr As Variant) As Integer

    largoArray = UBound(arr) - LBound(arr) + 1
    
End Function

'resumen: muestra los elementos de un array unidimensional por consola
'parametro: un array
'retorno: void
Public Function MostrarArray(ByVal arr As Variant)
    
    Dim i As Integer
    Dim auxiliar As Variant
    
    For i = 1 To UBound(arr)
        
        Debug.Print arr(i)
        
    Next i
    
    Exit Function
    
End Function

'resumen: muestra los elementos de un matriz de 2 dimensiones por consola
'parametro: una matriz
'retorno: void
Public Function MostrarMatriz(ByVal arr As Variant)
    
    Dim i, j As Integer
    Dim auxiliar As Variant

    For i = LBound(arr, 1) To UBound(arr, 1)
    
        For j = LBound(arr, 2) To UBound(arr, 2)
        
            Debug.Print arr(i, j)
        
        Next j
        
    Next i
    
End Function

'resumen: ordena una coleccion
'parametro: la coleccion a ordenar por referencia, zaraza1, zaraza2
'retorno: void
Public Function QuickSort(coll As Collection, first As Long, last As Long)
  
  Dim vCentreVal As Variant, vTemp As Variant
  
  Dim lTempLow As Long
  Dim lTempHi As Long
  lTempLow = first
  lTempHi = last
  
  vCentreVal = coll((first + last) \ 2)
  Do While lTempLow <= lTempHi
  
    Do While coll(lTempLow) < vCentreVal And lTempLow < last
      lTempLow = lTempLow + 1
    Loop
    
    Do While vCentreVal < coll(lTempHi) And lTempHi > first
      lTempHi = lTempHi - 1
    Loop
    
    If lTempLow <= lTempHi Then
    
      ' Swap values
      vTemp = coll(lTempLow)
      
      coll.Add coll(lTempHi), After:=lTempLow
      coll.Remove lTempLow
      
      coll.Add vTemp, Before:=lTempHi
      coll.Remove lTempHi + 1
      
      ' Move to next positions
      lTempLow = lTempLow + 1
      lTempHi = lTempHi - 1
      
    End If
    
  Loop
  
  If first < lTempHi Then QuickSort coll, first, lTempHi
  If lTempLow < last Then QuickSort coll, lTempLow, last
  
End Function
