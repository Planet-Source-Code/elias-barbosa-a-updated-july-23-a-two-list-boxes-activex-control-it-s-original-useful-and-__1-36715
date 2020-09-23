Attribute VB_Name = "modMyArray"
'======================================================================================================
'============ WHAT THIS FUNCTION DOES =================================================================
'======================================================================================================
'
'This function will sort an array with 2
'dimensions in alphabetical order. This
'function should not be used to sort
'numbers because it will sort, for example,
'the number 2 after the number 10. It
'happens because it compares character
'by character and not the whole number.
'
'======================================================================================================
'============ ELEMENTS OF THE FUNCTION ================================================================
'======================================================================================================
'
'ELEMENT 1 ==> TempArray = The Array that you want to sort.
'
'------------------------------------------------------------------------------------------------------
'
'ELEMENT 2 ==> iElement = The "Column" that will be used
'                         as reference while sorting the array.
'
'------------------------------------------------------------------------------------------------------
'
'ELEMENT 3 ==> iDimension = This variant will determine if
'                           the dimension that has the number
'                           of "Columns" on the array is the first
'                           dimension or the second dimension of
'                           the array. This is a little complicated.
'                           An example may explain it a little better:
'
'   EXAMPLE1:
'
'      In this example the number of "Columns" is on the first
'      dimension of the array. In this case, the iDimension value
'      would be 2. You have to count from right to left.
'
'      MyArray(2, 5)
'
'      Column = 2
'      Row = 5
'
'      |-------------|---------------|-----------------|
'      |   MyArray   |    Column 1   |     Column 2    |
'      |-------------|---------------|-----------------|
'      |    Row 1    |       1       |    Pineapple    |
'      |    Row 2    |       2       |    Orange       |
'      |    Row 3    |       3       |    Mango        |
'      |    Row 4    |       4       |    Apple        |
'      |    Row 5    |       5       |    Grape        |
'      |-------------|---------------|-----------------|
'
'
'   EXAMPLE2:
'
'      In this example the number of "Columns" is on the last
'      dimension of the array. In this case, the iDimension value
'      would be 1. Remember that you have to count from right to left.
'      When dealing with an array constructed this way, you can omit
'      the iDimension value because 1 is the default value.
'
'      MyArray(5, 2)
'
'      Row = 5
'      Column = 2
'
'      |-------------|---------------|-----------------|
'      |   MyArray   |    Column 1   |     Column 2    |
'      |-------------|---------------|-----------------|
'      |    Row 1    |       1       |    Pineapple    |
'      |    Row 2    |       2       |    Orange       |
'      |    Row 3    |       3       |    Mango        |
'      |    Row 4    |       4       |    Apple        |
'      |    Row 5    |       5       |    Grape        |
'      |-------------|---------------|-----------------|
'
'------------------------------------------------------------------------------------------------------
'
'ELEMENT 4 ==> bAscOrder = This property will determine
'                          whether the sorting is in ascending or
'                          descending order.
'
'
'======================================================================================================
'============ HOW TO USE THE FUNCTION =================================================================
'======================================================================================================
'
'On the array of EXAMPLE1
'
'   MyArray(2, 5)
'
'   Call Sort_2D_Bubble(MyArray, 1, 2)
'
'   The result will be:
'   |-------------|---------------|-----------------|
'   |   MyArray   |    Column 1   |     Column 2    |
'   |-------------|---------------|-----------------|
'   |    Row 1    |       1       |    Pineapple    |
'   |    Row 2    |       2       |    Orange       |
'   |    Row 3    |       3       |    Mango        |
'   |    Row 4    |       4       |    Apple        |
'   |    Row 5    |       5       |    Grape        |
'   |-------------|---------------|-----------------|
'
'   Call Sort_2D_Bubble(MyArray, 2, 2)
'
'   The result will be:
'   |-------------|---------------|-----------------|
'   |   MyArray   |    Column 1   |     Column 2    |
'   |-------------|---------------|-----------------|
'   |    Row 1    |       4       |    Apple        |
'   |    Row 2    |       5       |    Grape        |
'   |    Row 3    |       3       |    Mango        |
'   |    Row 4    |       2       |    Orange       |
'   |    Row 5    |       1       |    Pineapple    |
'   |-------------|---------------|-----------------|
'
'
'On the array of EXAMPLE2
'
'   MyArray(5, 2)
'
'   Call Sort_2D_Bubble(MyArray, 1)
'
'   The result will be:
'   |-------------|---------------|-----------------|
'   |   MyArray   |    Column 1   |     Column 2    |
'   |-------------|---------------|-----------------|
'   |    Row 1    |       1       |    Pineapple    |
'   |    Row 2    |       2       |    Orange       |
'   |    Row 3    |       3       |    Mango        |
'   |    Row 4    |       4       |    Apple        |
'   |    Row 5    |       5       |    Grape        |
'   |-------------|---------------|-----------------|
'
'   Call Sort_2D_Bubble(MyArray, 2)
'
'   The result will be:
'   |-------------|---------------|-----------------|
'   |   MyArray   |    Column 1   |     Column 2    |
'   |-------------|---------------|-----------------|
'   |    Row 1    |       4       |    Apple        |
'   |    Row 2    |       5       |    Grape        |
'   |    Row 3    |       3       |    Mango        |
'   |    Row 4    |       2       |    Orange       |
'   |    Row 5    |       1       |    Pineapple    |
'   |-------------|---------------|-----------------|
'
Public Function Sort_2D_Bubble( _
    ByRef TempArray As Variant, _
    Optional iElement As Integer = 1, _
    Optional iDimension As Integer = 1, _
    Optional bAscOrder As Boolean = True) As Boolean
    
    Dim NoExchanges As Boolean
    Dim arrTemp() As Variant
    Dim i As Integer
    Dim j As Integer
    
    On Error GoTo Error_BubbleSort
    
    If (iDimension = 1) Then
        ReDim arrTemp(1, UBound(TempArray, 2))
        
    Else
        ReDim arrTemp(UBound(TempArray, 1), 1)
        
    End If

    'Loop until no more "exchanges" are made.
    Do While Not (NoExchanges)
        NoExchanges = True
        
        'First, check if array has 1 or 2 dimensions.
        If (iDimension = 1) Then
            
            'Loop through each element in the array.
            For i = LBound(TempArray, iDimension) To UBound(TempArray, iDimension) - 1
                
                'If the element is greater than the element
                'following it, exchange the two elements.
                If (bAscOrder And (TempArray(i, iElement) > TempArray(i + 1, iElement))) _
                Or (Not bAscOrder And (TempArray(i, iElement) < TempArray(i + 1, iElement))) Then
                    
                    NoExchanges = False
                    
                    For j = LBound(TempArray, 2) To UBound(TempArray, 2)
                        arrTemp(1, j) = TempArray(i, j)
                        
                    Next j
                    
                    For j = LBound(TempArray, 2) To UBound(TempArray, 2)
                        TempArray(i, j) = TempArray(i + 1, j)
                        
                    Next j
                    
                    For j = LBound(TempArray, 2) To UBound(TempArray, 2)
                        TempArray(i + 1, j) = arrTemp(1, j)
                        
                    Next j
                    
                End If
            Next i
            
        Else

            For i = LBound(TempArray, iDimension) To UBound(TempArray, iDimension) - 1
                
                'If the element is greater than the element
                ' following it, exchange the two elements.
                If (bAscOrder And (TempArray(iElement, i) > TempArray(iElement, i + 1))) _
                Or (Not bAscOrder And (TempArray(iElement, i) < TempArray(iElement, i + 1))) Then
                    NoExchanges = False
                    
                    For j = LBound(TempArray, 1) To UBound(TempArray, 1)
                        arrTemp(j, 1) = TempArray(j, i)
                        
                    Next j
                    
                    For j = LBound(TempArray, 1) To UBound(TempArray, 1)
                        TempArray(j, i) = TempArray(j, i + 1)
                        
                    Next j
                    
                    For j = LBound(TempArray, 1) To UBound(TempArray, 1)
                        TempArray(j, i + 1) = arrTemp(j, 1)
                        
                    Next j
                    
                End If
                
            Next i
            
        End If
        
    Loop
    
    Sort_2D_Bubble = True
    
    Exit Function

Error_BubbleSort:
    Sort_2D_Bubble = False
    
End Function

