Attribute VB_Name = "RandomCode"
Public Sub open_form(formname As Form, studno As String, Name As String)
    formname.Visible = True
    formname.Caption = formname.Caption & " :: " & studno & " - " & Name
End Sub

Function Mean(k As Integer, Arr() As Single)
     Dim Sum As Single
     Dim i As Integer

     Sum = 0
     For i = 1 To k
         Sum = Sum + Arr(i)
     Next i
 
     Mean = Sum / k

End Function

Function StdDev(k As Integer, Arr() As Single)
    Dim i As Integer
    Dim avg As Single, SumSq As Single
 
    avg = Mean(k, Arr)
    For i = 1 To k
        SumSq = SumSq + (Arr(i) - avg) ^ 2
    Next i
 
    StdDev = Sqr(SumSq / (k - 1))
    
End Function

Function biggest_number(k As Integer, Arr() As Single)
    Dim i As Integer
    Dim max As Single
    
    For i = 0 To k
        If Arr(i) > max Then
            max = Arr(i)
        End If
    Next
    biggest_number = max
End Function

Function smallest_number(k As Integer, Arr() As Single)
    Dim i As Integer
    Dim min As Single
    
    For i = 0 To k
        If Arr(i) < min Then
            min = Arr(i)
        End If
    Next
    smallest_number = min
End Function
