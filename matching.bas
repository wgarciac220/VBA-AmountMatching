Sub test()
    Dim item As Variant
    Dim data As Variant
    
'        data = bestSum(7, Array(5, 3, 4, 7))
'        data = bestSum(8, Array(2, 3, 5))
'        data = bestSum(8, Array(1, 4, 5, 3))
'        data = bestSum(9935018, Array(2628223, 3751854, 7306795))
        
        For Each item In data
            Debug.Print item
        Next item

End Sub

Public Function bestSum(targetSum, numbers, Optional memo As Scripting.Dictionary) As Variant

    ' initialize dictionary
    If memo Is Nothing Then Set memo = New Scripting.Dictionary
    
    ' retrieve key value if exists
    If memo.Exists(targetSum) Then
        bestSum = memo(targetSum)
        GoTo EXIT_HERE
    End If

    ' return 0 if is perfect match
    If targetSum = 0 Then
        bestSum = Array(0)
        GoTo EXIT_HERE
    End If
    
    ' return -1 if targetSum is negative
    If targetSum < 0 Then
        bestSum = Array(-1)
        Exit Function
    End If
    
    Dim shortestCombination() As Variant
    shortestCombination = Array(-1)     ' initialize array
    
    Dim num As Variant
    
    For Each num In numbers     ' loop through given candidates
    
        Dim j As Integer
                
        Dim remainder As Double
        remainder = targetSum - num     ' get remainder
        
        Dim remainderCombination() As Variant
        remainderCombination = bestSum(remainder, removeItem(numbers, j), memo)       ' remainder combination
        'remainderCombination = bestSum(remainder, numbers, memo)        ' remainder combination
        
        If Not remainderCombination(0) = -1 Then
            Dim combination() As Variant
            combination = expand(remainderCombination, num)     ' add current number to combination
            
            If (shortestCombination(0) = -1) _
            Or (UBound(combination) < UBound(shortestCombination) + 1) Then
                shortestCombination = combination       ' update combination
            End If
        End If
        
        j = j + 1
        
    Next num
    
    memo(targetSum) = shortestCombination       ' add sum to dictionary
    
    bestSum = memo(targetSum)
    
EXIT_HERE:
End Function

Private Function expand(remainderCombination, num) As Variant
        Dim temp As Variant
        ReDim temp(UBound(remainderCombination) + 1)
        
        Dim item As Variant
        
        For Each item In remainderCombination
        
            Dim i As Integer
            temp(i) = item
            i = i + 1
            
        Next item
        
        temp(i) = num
        expand = temp
    
End Function

Private Function removeItem(source, index) As Variant

    If UBound(source) = 0 Then
        removeItem = Array()
        Exit Function
    End If

    Dim temp() As Variant
    ReDim temp(UBound(source) - 1)

    Dim i As Integer
    
    For i = LBound(source) To UBound(source) - 1
        temp(i) = source(IIf(i >= index, i + 1, i))
    Next i
    
    removeItem = temp
    
End Function

