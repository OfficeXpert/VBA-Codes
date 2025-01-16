Sub Unhide_All_Sheets_Count()
  Dim wks As Worksheet
  Dim count As Integer

  count = 0

  For Each wks In ActiveWorkbook.Worksheets
    If wks.Visible <> xlSheetVisible Then
      wks.Visible = xlSheetVisible
      count = count + 1
    End If
  Next wks

  If count > 0 Then
    MsgBox count &amp; " worksheets have been unhidden.", vbOKOnly, "Unhiding worksheets"
  Else
    MsgBox "No hidden worksheets have been found.", vbOKOnly, "Unhiding worksheets"
  End If
End Sub  
