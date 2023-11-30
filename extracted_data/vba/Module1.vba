Attribute VB_Name = "Module1"
Sub DeleteOutputWithConfirmation()
    Dim ws As Worksheet
    Dim LastRow As Long
    Dim StartRow As Long
    Dim EndRow As Long
    Dim response As VbMsgBoxResult

    ' Define the worksheet you want to work with
    Set ws = ThisWorkbook.Worksheets("Tool") ' Change "Tool" to your sheet name

    ' Ask for confirmation
    response = MsgBox("Are you sure you want to clear the output data?", vbQuestion + vbYesNo, "Confirmation")

    ' Check the user's response
    If response = vbNo Then
        Exit Sub ' Exit the macro if the user chooses not to proceed
    End If

    ' Find the last row with data in column S
    LastRow = ws.Cells(ws.Rows.Count, "S").End(xlUp).Row

    ' Determine the start row and end row for clearing
    StartRow = 2 ' Start from the second row (assuming headers are in the first row)
    EndRow = LastRow

    ' Clear the entire block from columns S to AB for all rows in the determined range
    ws.Range("S" & StartRow & ":AB" & EndRow).ClearContents
End Sub
