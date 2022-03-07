Function GetSheetWithCodename(ByVal worksheetCodename As String, Optional wb As Workbook) As Worksheet
    Dim iSheet As Long
    If wb Is Nothing Then Set wb = ThisWorkbook ' mimics the default behaviour
    For iSheet = 1 To wb.Worksheets.Count
        If wb.Worksheets(iSheet).CodeName = worksheetCodename Then
            Set GetSheetWithCodename = wb.Worksheets(iSheet)
            Exit Function
        End If
    Next iSheet
End Function