# VBA-AmountMatching

This is an amount matching recursive algorithm that matches one to many and returns the shortest combination


```VBA
Sub test()
    Dim item As Variant
    Dim data As Variant
    
        data = bestSum(7, Array(5, 3, 4, 7)) ' returns 7
        data = bestSum(8, Array(1, 3, 4, 5)) ' returns 3,5
        
        For Each item In data
            Debug.Print item
        Next item

End Sub
```
