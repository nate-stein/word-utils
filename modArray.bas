Attribute VB_Name = "modArray"
Option Explicit

Public Sub Arr_DeleteElement(ByVal index As Integer, ByRef arr As Variant)
'********************************************************
' Deletes array element @ index.
'********************************************************
   
   Dim i As Integer
   ' Move all element back one position
   For i = index + 1 To UBound(arr)
      arr(i - 1) = arr(i)
   Next
      
   ' Shrink the array by one, removing the last one
   ReDim Preserve arr(0 To UBound(arr) - 1)

End Sub

Public Sub Array_DisplayArray(ByVal arr As Variant)
'*********************************************************
' Displays contents of an array with messagebox. Format of output:
'        (0,0) = ??
'        (0,1) = ??
'*********************************************************

   Dim dimensions As Integer
   dimensions = Array_GetNumberOfDimensions(arr)
   
   Dim msg As String, q As Long
   For q = 0 To UBound(arr)
      If dimensions > 1 Then
         Dim dimension As Integer
         For dimension = 0 To (dimensions - 1) Step 1
            msg = msg & vbLf & _
            "[" & q & ", " & dimension & "] = " & arr(q, dimension)
         Next dimension
      Else:
         msg = msg & vbLf & _
         "[" & q & "] = " & arr(q)
      End If
   Next q
   
   MsgBox msg, , "Array_DisplayArray"

End Sub

Public Function Array_ContainsValue( _
   ByRef arr As Variant, _
   ByVal value As Variant) As Boolean
'*********************************************************
' Returns True if value is one of the elements in arr.
'*********************************************************

   Dim q As Integer
   For q = 0 To UBound(arr) Step 1
      If arr(q) = value Then
         Array_ContainsValue = True
         Exit Function
      End If
   Next q
   Array_ContainsValue = False

End Function

Public Sub Array_AddValue(ByRef arr As Variant, ByVal values As Variant)
'*********************************************************
' Adds values to arr (assumed to be an array). values can be a single value or an array.
'*********************************************************

   On Error GoTo RAISE_ERR
   
   '''''''''''''''''''''''''''''''''''''''
   ' If passed array wasn't yet initialized (i.e. we need to use ReDim only instead of
   ' ReDim Preserve.
   '''''''''''''''''''''''''''''''''''''''
   If Not Array_IsAllocated(arr) Then
      If IsArray(values) Then
         ReDim arr(0 To UBound(values))
         Dim q As Integer
         For q = 0 To UBound(values)
            arr(q) = values(q)
         Next q
      Else:
         ReDim arr(0 To 0)
         arr(0) = values
      End If
      Exit Sub
   End If
   
   '''''''''''''''''''''''''''''''''''''''
   ' If passed array was already initialized.
   '''''''''''''''''''''''''''''''''''''''
   Dim nextPositionInArray As Integer
   If IsArray(values) Then
      Dim elementsToAdd As Integer: elementsToAdd = UBound(values) + 1
      Dim j As Integer
      For j = 1 To elementsToAdd Step 1
         nextPositionInArray = UBound(arr) + 1
         ReDim Preserve arr(0 To nextPositionInArray)
         arr(nextPositionInArray) = values(j - 1)
      Next j
   Else:
      nextPositionInArray = UBound(arr) + 1
      ReDim Preserve arr(0 To nextPositionInArray)
      arr(nextPositionInArray) = values
   End If
   Exit Sub
   
RAISE_ERR:
   Err.Raise g_ERROR_ARRAYS, "Array_AddValue"
   
End Sub

Public Function Array_GetNumberOfDimensions(ByVal arr As Variant) As Integer
'*********************************************************
' Returns number of dimensions in an array.
'*********************************************************

   On Error GoTo RETURN_RESULT
   Dim i As Integer
   Dim tmp As Integer
   i = 0
   Do While True:
      i = i + 1
      tmp = UBound(arr, i)
   Loop
RETURN_RESULT:
   Array_GetNumberOfDimensions = i - 1
   
End Function

Public Function Array_IsAllocated(ByVal arr As Variant) As Boolean
'*********************************************************
' Returns True if the array is allocated (either a static array or a dynamic array that has been
' sized with Redim) or False if the array has not been allocated (a dynamic that has not yet been
' sized with Redim, or a dynamic array that has been Erased).
'*********************************************************

   Dim n As Long
   
   ' If arr is not an array, return FALSE and get out.
   If Not IsArray(arr) Then
      Array_IsAllocated = False
      Exit Function
   End If
   
   ' Try to get the UBound of the array. If the array has not been allocated, an error will occur.
   On Error Resume Next
   n = UBound(arr, 1)
   If Err.number = 0 Then
      Array_IsAllocated = True
   Else: Array_IsAllocated = False
   End If

End Function
