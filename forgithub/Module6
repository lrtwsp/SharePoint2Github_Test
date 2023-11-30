Attribute VB_Name = "Module6"
Sub ExportColumnsToFileExplorer()
    On Error GoTo ErrorHandler
    Dim SourceWorkbook As Workbook
    Dim NewWorkbook As Workbook
    Dim SourceWorksheet As Worksheet
    Dim NewWorksheet As Worksheet
    Dim LastRow As Long
    Dim LastColumn As Long
    Dim ExportRange As Range
    Dim SavePath As String

    ' Set the source workbook and worksheet
    Set SourceWorkbook = ThisWorkbook
    Set SourceWorksheet = SourceWorkbook.Sheets("Tool") ' Change sheet name to "Tool"

    ' Find the last row and last column with data
    LastRow = SourceWorksheet.Cells(SourceWorksheet.Rows.Count, "S").End(xlUp).Row
    LastColumn = SourceWorksheet.Cells(1, SourceWorksheet.Columns.Count).End(xlToLeft).Column

    ' Define the range to be exported, including formatting
    Set ExportRange = SourceWorksheet.Range(SourceWorksheet.Cells(1, 19), SourceWorksheet.Cells(LastRow, LastColumn)) ' Columns S to AB (columns 19 to 28)

    ' Get the name of the source workbook
    Dim SourceWorkbookName As String
    SourceWorkbookName = Left(SourceWorkbook.Name, InStrRev(SourceWorkbook.Name, ".") - 1)

    ' Create a new workbook
    Set NewWorkbook = Workbooks.Add
    Set NewWorksheet = NewWorkbook.Sheets(1)

    ' Copy both values and formatting
    ExportRange.Copy
    NewWorksheet.Cells(1, 1).PasteSpecial Paste:=xlPasteValues
    NewWorksheet.Cells(1, 1).PasteSpecial Paste:=xlPasteFormats

    ' Set the default save name as the name of the source file plus 'Exported Results'
    Dim DefaultSaveName As String
    DefaultSaveName = SourceWorkbookName & " Exported Results.xlsx"

    ' Prompt the user to select the location and name of the export file
    With Application.FileDialog(msoFileDialogSaveAs)
        .Title = "Save As"
        .InitialFileName = DefaultSaveName
        If .Show = -1 Then
            SavePath = .SelectedItems(1)
            ' Save the new workbook with the selected path and name
            NewWorkbook.SaveAs SavePath
        Else
            ' User canceled the save dialog
            Exit Sub
        End If
    End With

    ' Close the new workbook without saving changes
    NewWorkbook.Close SaveChanges:=False

    ' Clean up
    Set NewWorksheet = Nothing
    Set NewWorkbook = Nothing
    Set ExportRange = Nothing
    Set SourceWorksheet = Nothing
    Set SourceWorkbook = Nothing

    Exit Sub

ErrorHandler:
    MsgBox "An error occurred: " & Err.Description
End Sub

