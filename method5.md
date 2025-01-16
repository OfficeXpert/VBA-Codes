# VBA-Codes

Sub RemoveLastCharacter()

Dim rng As Range
Dim cell As Range
Dim newValue As String

' Prompt the user to select a range
On Error Resume Next
Set rng = Application.InputBox("Select a range of cells:", Type:=8)
On Error GoTo 0

' Check if a range was selected
If rng Is Nothing Then
MsgBox "No range selected. Operation canceled.", vbExclamation

Exit Sub
End If

' Loop through each cell in the selected range
For Each cell In rng
If Len(cell.Value) > 0 Then
' Check if the cell is not empty
newValue = Left(cell.Value, Len(cell.Value) - 1)
cell.Value = newValue
End If
Next cell

MsgBox "Last character removed from the selected range.", vbInformation

End Sub
