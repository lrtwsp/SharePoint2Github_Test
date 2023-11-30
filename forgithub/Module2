Attribute VB_Name = "Module2"
Sub FlagMatchingRowsDynamicTimeBuffer(AnalysisType As String, SameDirectionSelected As Boolean, TimeBufferValue As Double)
    Dim ws As Worksheet
    Dim LastRow As Long
    Dim outputData As Variant
    Dim matchStorage() As Variant
    Dim matchCount As Long
    Dim i As Long, k As Long
    Dim condFormulaRed As String
    Dim condFormulaYellow As String
    Dim originalRowNumbers() As Long
    Dim iOriginal As Long, kOriginal As Long
    Dim matchType As String

    Set ws = ThisWorkbook.Worksheets("Tool")

    ' Calculate the time buffer in minutes
    Dim TimeBuffer As Double
    TimeBuffer = TimeBufferValue / (24 * 60)

    ' Find the last row in column D
    LastRow = ws.Cells(ws.Rows.Count, "D").End(xlUp).Row

    ' Read the entire dataset into an array and store original row numbers
    outputData = ws.Range("A2:K" & LastRow).Value
    ReDim originalRowNumbers(1 To LastRow - 1)
    For i = 1 To LastRow - 1
        originalRowNumbers(i) = i + 1 ' Modify the original row numbers to match the row numbers in the file
    Next i

    ' Custom sorting function to sort the array based on time in column K
    QuickSort outputData, 1, UBound(outputData, 1) - 1, 11

    ' Initialize matchCount
    matchCount = 0

    ' Remove existing conditional formatting in column AB
    ws.Range("AB2:AB" & LastRow).FormatConditions.Delete

    ' Apply new conditional formatting rules
    condFormulaRed = "=($AB2=""Head On"")"
    condFormulaYellow = "=($AB2=""Same Direction"")"

    With ws.Range("AB2:AB" & LastRow)
        .FormatConditions.Add Type:=xlExpression, Formula1:=condFormulaRed
        .FormatConditions(1).Interior.Color = RGB(255, 0, 0) ' Red for "Head On"

        .FormatConditions.Add Type:=xlExpression, Formula1:=condFormulaYellow
        .FormatConditions(2).Interior.Color = RGB(255, 255, 0) ' Yellow for "Same Direction"
    End With

    ' Process the data and store matches using original row numbers
    ReDim matchStorage(1 To LastRow, 1 To 12) ' Extend the array to accommodate the row numbers
    For i = 1 To LastRow - 2
        For k = i + 1 To LastRow - 1
            ' Retrieve original row numbers from sorted array
            iOriginal = originalRowNumbers(i)
            kOriginal = originalRowNumbers(k)

            ' Check if column I is empty, and if so, use the value from the row above
            If Not IsEmpty(outputData(i, 9)) And Not IsEmpty(outputData(k, 9)) Then
                If outputData(i, 9) = outputData(k, 9) Then
                    matchType = "Same Direction"
                Else
                    matchType = "Head On"
                End If
            End If

            ' Check if column K is empty, and if so, use the value from column J as a replacement
            If IsEmpty(outputData(i, 11)) Then
                outputData(i, 11) = outputData(i, 10)
            End If
            If IsEmpty(outputData(k, 11)) Then
                outputData(k, 11) = outputData(k, 10)
            End If

            If Not (IsEmpty(outputData(i, 4)) Or IsEmpty(outputData(k, 4)) Or _
                    IsEmpty(outputData(i, 6)) Or IsEmpty(outputData(k, 6)) Or _
                    IsEmpty(outputData(i, 11)) Or IsEmpty(outputData(k, 11))) Then
                If outputData(i, 4) = outputData(k, 4) And outputData(i, 6) = outputData(k, 6) Then
                    Dim startTime As Date
                    Dim endTime As Date
                    Dim compareTime As Date

                    startTime = CDate(outputData(i, 11)) - TimeBuffer
                    endTime = CDate(outputData(i, 11)) + TimeBuffer
                    compareTime = CDate(outputData(k, 11))

                    If startTime <= compareTime And compareTime <= endTime Then
                        If matchType = AnalysisType Or AnalysisType = "Both" Then
                            matchCount = matchCount + 1

                            ' Store matching data including the row numbers in the matchStorage array
                            matchStorage(matchCount, 1) = outputData(i, 2)
                            matchStorage(matchCount, 2) = outputData(k, 2)
                            matchStorage(matchCount, 3) = outputData(i, 4)
                            matchStorage(matchCount, 4) = outputData(k, 4)
                            matchStorage(matchCount, 5) = outputData(i, 6)
                            matchStorage(matchCount, 6) = outputData(i, 11) ' Value from column K for lower time match
                            matchStorage(matchCount, 7) = outputData(k, 11) ' Value from column K for higher time match
                            matchStorage(matchCount, 8) = iOriginal  ' Store row number of the row with the lower time
                            matchStorage(matchCount, 9) = kOriginal ' Store row number of the row with the higher time
                            matchStorage(matchCount, 10) = matchType ' Store the match type in column AB
                            matchStorage(matchCount, 11) = i ' Store the row number of the data source in column Z
                            matchStorage(matchCount, 12) = k ' Store the row number of the data source in column AA
                        End If
                    End If
                End If
            End If
        Next k
    Next i

    ' Write the matched data back to the worksheet
    ws.Range("S2").Resize(matchCount, 12).Value = matchStorage

    ' Clean up
    Set ws = Nothing
    Erase outputData
    Erase matchStorage
End Sub


Sub QuickSort(vArray As Variant, inLow As Long, inHi As Long, SortIndex As Long)
    Dim pivot   As Variant
    Dim tmpSwap As Variant
    Dim tmpLow   As Long
    Dim tmpHi    As Long

    tmpLow = inLow
    tmpHi = inHi

    pivot = vArray((inLow + inHi) \ 2, SortIndex)

    While (tmpLow <= tmpHi)
        While (vArray(tmpLow, SortIndex) < pivot And tmpLow < inHi)
            tmpLow = tmpLow + 1
        Wend

        While (pivot < vArray(tmpHi, SortIndex) And tmpHi > inLow)
            tmpHi = tmpHi - 1
        Wend

        If (tmpLow <= tmpHi) Then
            For i = 1 To UBound(vArray, 2)
                tmpSwap = vArray(tmpLow, i)
                vArray(tmpLow, i) = vArray(tmpHi, i)
                vArray(tmpHi, i) = tmpSwap
            Next i

            tmpLow = tmpLow + 1
            tmpHi = tmpHi - 1
        End If
    Wend

    If (inLow < tmpHi) Then QuickSort vArray, inLow, tmpHi, SortIndex
    If (tmpLow < inHi) Then QuickSort vArray, tmpLow, inHi, SortIndex
End Sub



