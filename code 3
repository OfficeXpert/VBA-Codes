Sub Unhide_Selected_Sheets()
    Dim wks As Worksheet
    Dim MsgResult As VbMsgBoxResult

    For Each wks In ActiveWorkbook.Worksheets
        If wks.Visible = xlSheetHidden Then
            MsgResult = MsgBox("Unhide sheet " & wks.Name & "?", vbYesNo, "Unhiding Worksheets")
            If MsgResult = vbYes Then wks.Visible = xlSheetVisible
        End If
    Next
End Sub
