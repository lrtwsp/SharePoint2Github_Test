Attribute VB_Name = "Module3"
Sub DeleteInputWithConfirmation()
    Dim ws As Worksheet
    Dim LastRow As Long
    Dim i As Long
    Dim response As VbMsgBoxResult

    ' Define the worksheet you want to work with
    Set ws = ThisWorkbook.Worksheets("Tool") ' Change "Sheet1" to your sheet name

    ' Ask for confirmation
    response = MsgBox("Are you sure you want to clear the input data?", vbQuestion + vbYesNo, "Confirmation")

    ' Check the user's response
    If response = vbNo Then
        Exit Sub ' Exit the macro if the user chooses not to proceed
    End If

    ' Find the last row with data in column A
    LastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    ' Loop through the rows, starting from the second row (assuming headers are in the first row)
    For i = 2 To LastRow
        ' Clear the cell contents in columns A to P for the current row
        ws.Range("A" & i & ":P" & i).ClearContents
    Next i
End Sub

