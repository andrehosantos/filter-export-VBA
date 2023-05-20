Sub FilterAndExport()
    Dim arr() As Variant
    Dim i As Integer
    
    'Read values from Sheet1 into array
    arr = Sheets("Sheet2").Range("A2:A100").Value
    
    'Loop through array and apply each non-blank value as filter to Sheet2
    For i = LBound(arr) To UBound(arr)
        If Not IsEmpty(arr(i, 1)) Then ' Check if value is not blank
            Dim criteria As String
            If IsNumeric(arr(i, 1)) Then
                'Convert numerical value to text value
                criteria = CStr(arr(i, 1))
            Else
                'Use text value as-is
                criteria = arr(i, 1)
            End If
            Sheets("Sheet1").Range("E1").AutoFilter Field:=5, Criteria1:=criteria
            Sheets("Sheet1").UsedRange.SpecialCells(xlCellTypeVisible).Copy
            
            If Dir("X:\folder\path" & criteria & ".csv") <> "" Then
                Kill "X:\folder\path" & criteria & ".csv"
            End If
            
            'Create new workbook to export filtered data as csv file
            Workbooks.Add
            Range("A1").Select
            Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
            Columns("F:F").Select
            Application.CutCopyMode = False
            Application.CutCopyMode = False
            Selection.NumberFormat = "000"
           ActiveWorkbook.SaveAs Filename:="X:\folder\path" & criteria & ".csv", FileFormat:=xlCSV, CreateBackup:=False
            ActiveWorkbook.Close False
            ActiveSheet.Range("$A$1:$K$1369").AutoFilter Field:=5
        End If
    Next i
    
    'Clear filters on Sheet2
    Sheets("Sheet1").AutoFilterMode = False
End Sub
