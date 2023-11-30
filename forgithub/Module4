Attribute VB_Name = "Module4"
Sub OpenUserForm()
    ' Show the user form
    UF1_ToolInput.Show vbModal
    
    ' Once the user form is closed, retrieve the selected values
    Dim TimeBuffer As Double
    Dim AnalysisType As String
    
    ' Check if the user form was submitted
    If UF1_ToolInput.ClosedByOK Then
        ' Get the selected Time Buffer value from the user form
        TimeBuffer = CDbl(UF1_ToolInput.TextBoxTimeBuffer.Value)
        
        ' Get the selected Analysis Type from the user form
        If UF1_ToolInput.OptionButtonSameDirection.Value Then
            AnalysisType = "Same Direction"
        ElseIf UF1_ToolInput.OptionButtonHeadOn.Value Then
            AnalysisType = "Head On"
        End If
        
        ' Call the main code with the selected Time Buffer and Analysis Type
        FlagMatchingRowsLowMemoryUsage TimeBuffer, AnalysisType
    End If
End Sub

