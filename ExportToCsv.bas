Private Sub ExportToCsvButton_Click()

    Dim CurrentWB As Workbook, TempWB As Workbook
    Dim CsvFileName As String
    Dim lastRow As Long, lastCol As Long
    Dim TempWS As Worksheet
    Dim Cell As Range
    Dim DataRange As Range

    ' Set the current workbook and worksheet
    Set CurrentWB = ActiveWorkbook
    Dim CurrentWS As Worksheet
    Set CurrentWS = CurrentWB.ActiveSheet

    ' Find the last used row and column with visible data
    With CurrentWS
        lastRow = 0
        lastCol = 0

        ' Find the last row with actual data (ignoring empty cells with formulas)
        For Each Cell In .Columns("A:A").Cells
            If Cell.Value <> "" Then
                lastRow = Cell.Row
            ElseIf lastRow <> 0 Then
                Exit For ' Exit loop once the last non-empty cell is found
            End If
        Next Cell

        ' Find the last column with actual data (ignoring empty cells with formulas)
        For Each Cell In .Rows("1:1").Cells
            If Cell.Value <> "" Then
                lastCol = Cell.Column
            ElseIf lastCol <> 0 Then
                Exit For ' Exit loop once the last non-empty cell is found
            End If
        Next Cell

        ' Define the range to copy
        If lastRow > 0 And lastCol > 0 Then
            Set DataRange = .Range(.Cells(1, 1), .Cells(lastRow, lastCol))
        Else
            Set DataRange = .Range("A1") ' Default to a single cell if no data
        End If
    End With

    ' Create a new workbook and copy the data
    Set TempWB = Application.Workbooks.Add(1)
    Set TempWS = TempWB.Sheets(1)
    
    ' Copy data to the new workbook
    DataRange.Copy
    With TempWS
        .Range("A1").PasteSpecial xlPasteValues

        ' Determine the actual used range in the temporary sheet
        lastRow = .Cells(.Rows.Count, "A").End(xlUp).Row
        lastCol = .Cells(1, .Columns.Count).End(xlToLeft).Column
        
        ' Remove excess rows and columns
        On Error Resume Next
        If lastRow < .Rows.Count Then
            .Rows(lastRow + 1 & ":" & .Rows.Count).Delete
        End If
        If lastCol < .Columns.Count Then
            .Columns(lastCol + 1 & ":" & .Columns.Count).Delete
        End If
        On Error GoTo 0
    End With

    ' Generate the CSV filename using the current date and time
    CsvFileName = CurrentWB.Path & "/" & CurrentWS.Name & "_" & Format(Now(), "DD-MMM-YYYY_HH-MM-SS") & ".csv"

    ' Save the temporary workbook as a CSV file
    Application.DisplayAlerts = False
    TempWB.SaveAs Filename:=CsvFileName, FileFormat:=xlCSV, CreateBackup:=False, Local:=True
    TempWB.Close SaveChanges:=False
    Application.DisplayAlerts = True

    ' Inform the user that the process is complete
    MsgBox "Export complete. File saved as: " & CsvFileName, vbInformation

End Sub
