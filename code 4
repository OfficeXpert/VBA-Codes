Sub Unhide_Sheets_By_Name()
    Dim wks As Worksheet
    Dim count As Integer
    Dim searchName As String

    ' Prompt the user to enter the keyword
    searchName = InputBox("Enter the name or keyword to search for in worksheet names:", "Unhide Sheets")
    
    ' Exit if the user cancels or leaves the input blank
    If searchName = "" Then
        MsgBox "No input provided. Operation canceled.", vbExclamation, "Unhiding Worksheets"
        Exit Sub
    End If

    count = 0

    ' Loop through all worksheets in the active workbook
    For Each wks In ActiveWorkbook.Worksheets
        If (wks.Visible <> xlSheetVisible) And (InStr(1, wks.Name, searchName, vbTextCompare) > 0) Then
            wks.Visible = xlSheetVisible
            count = count + 1
        End If
    Next wks

    ' Display the result
    If count > 0 Then
        MsgBox count & " worksheet(s) containing '" & searchName & "' have been unhidden.", vbInformation, "Unhiding Worksheets"
    Else
        MsgBox "No hidden worksheets containing '" & searchName & "' were found.", vbInformation, "Unhiding Worksheets"
    End If
End Sub
