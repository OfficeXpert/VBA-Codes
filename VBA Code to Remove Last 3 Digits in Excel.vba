Function RemoveLastThreeDigits(inputText As String) As String
    If Len(inputText) <= 3 Then
        RemoveLastThreeDigits = ""
    Else
        RemoveLastThreeDigits = Left(inputText, Len(inputText) - 3)
    End If
End Function
