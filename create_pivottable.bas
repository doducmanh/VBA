Attribute VB_Name = "Code_CreatePivotTable"
Sub createPivotTableExistingSheet()
 
    'Source: https://powerspreadsheets.com/
    'For further information: https://powerspreadsheets.com/vba-create-pivot-table/
 
    'declare variables to hold row and column numbers that define source data cell range
    Dim myFirstRow As Long
    Dim myLastRow As Long
    Dim myFirstColumn As Long
    Dim myLastColumn As Long
 
    'declare variables to hold source and destination cell range address
    Dim mySourceData As String
    Dim myDestinationRange As String
 
    'declare object variables to hold references to source and destination worksheets, and new Pivot Table
    Dim mySourceWorksheet As Worksheet
    Dim myDestinationWorksheet As Worksheet
    Dim myPivotTable As PivotTable
 
    'identify source and destination worksheets
    With ThisWorkbook
        Set mySourceWorksheet = .Worksheets("CHI TIET HSUV 2021")
        Set myDestinationWorksheet = .Worksheets("CHI TIET HSUV 2021")
        'Set myDestinationWorksheet = .Worksheets("Sheet4")
    End With
 
    'obtain address of destination cell range
    myDestinationRange = myDestinationWorksheet.Range("AF8").Address(ReferenceStyle:=xlR1C1)
 
    'identify row and column numbers that define source data cell range
    'Dim lrWBaftercombine As Integer
    'lrWBaftercombine = Sheet4.Range("C" & Rows.Count).End(xlUp).Row
    myFirstRow = 8
    myLastRow = 2000
    myFirstColumn = 27
    myLastColumn = 28
 
    'obtain address of source data cell range
    With mySourceWorksheet.Cells
        mySourceData = .Range(.Cells(myFirstRow, myFirstColumn), .Cells(myLastRow, myLastColumn)).Address(ReferenceStyle:=xlR1C1)

    End With
 
    'create Pivot Table cache and create Pivot Table report based on that cache
    Set myPivotTable = ThisWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:="'" & mySourceWorksheet.Name & "'!" & mySourceData).CreatePivotTable(TableDestination:="'" & myDestinationWorksheet.Name & "'!" & myDestinationRange, TableName:="PivotTableExistingSheet")
 
    'add, organize and format Pivot Table fields
    With myPivotTable
        .PivotFields("ADCODEMIENVP").Orientation = xlRowField
        With .PivotFields("SLHS")
            .Orientation = xlDataField
            .Position = 1
            .Function = xlSum
            '.NumberFormat = "#,##0.00"
        End With

    End With
 
End Sub

Sub DeleteAllPivotTablesInWorkbook()
'Updateby20140618
Dim xWs As Worksheet
Dim xPT As PivotTable
For Each xWs In Application.ActiveWorkbook.Worksheets
    For Each xPT In xWs.PivotTables
        xWs.Range(xPT.TableRange2.Address).Delete Shift:=xlUp
    Next
Next
End Sub

Sub createPivotTableNewSheet()
 
    'Source: https://powerspreadsheets.com/
    'For further information: https://powerspreadsheets.com/vba-create-pivot-table/
 
    'declare variables to hold row and column numbers that define source data cell range
    Dim myFirstRow As Long
    Dim myLastRow As Long
    Dim myFirstColumn As Long
    Dim myLastColumn As Long
 
    'declare variables to hold source and destination cell range address
    Dim mySourceData As String
    Dim myDestinationRange As String
 
    'declare object variables to hold references to source and destination worksheets, and new Pivot Table
    Dim mySourceWorksheet As Worksheet
    Dim myDestinationWorksheet As Worksheet
    Dim myPivotTable As PivotTable
 
    'identify source and destination worksheets. Add destination worksheet
    With ThisWorkbook
        Set mySourceWorksheet = .Worksheets("Data")
        Set myDestinationWorksheet = .Worksheets.Add
    End With
 
    'obtain address of destination cell range
    myDestinationRange = myDestinationWorksheet.Range("A5").Address(ReferenceStyle:=xlR1C1)
 
    'identify row and column numbers that define source data cell range
    myFirstRow = 5
    myLastRow = 20005
    myFirstColumn = 1
    myLastColumn = 6
 
    'obtain address of source data cell range
    With mySourceWorksheet.Cells
        mySourceData = .Range(.Cells(myFirstRow, myFirstColumn), .Cells(myLastRow, myLastColumn)).Address(ReferenceStyle:=xlR1C1)
    End With
 
    'create Pivot Table cache and create Pivot Table report based on that cache
    Set myPivotTable = ThisWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=mySourceWorksheet.Name & "!" & mySourceData).CreatePivotTable(TableDestination:=myDestinationWorksheet.Name & "!" & myDestinationRange, TableName:="PivotTableNewSheet")
 
    'add, organize and format Pivot Table fields
    With myPivotTable
        .PivotFields("Item").Orientation = xlRowField
        With .PivotFields("Units Sold")
            .Orientation = xlDataField
            .Position = 1
            .Function = xlSum
            .NumberFormat = "#,##0.00"
        End With
        With .PivotFields("Sales Amount")
            .Orientation = xlDataField
            .Position = 2
            .Function = xlSum
            .NumberFormat = "#,##0.00"
        End With
    End With
 
End Sub

Sub createPivotTableNewWorkbook()
 
    'Source: https://powerspreadsheets.com/
    'For further information: https://powerspreadsheets.com/vba-create-pivot-table/
 
    'declare variables to hold row and column numbers that define source data cell range
    Dim myFirstRow As Long
    Dim myLastRow As Long
    Dim myFirstColumn As Long
    Dim myLastColumn As Long
 
    'declare variables to hold source and destination cell range address
    Dim mySourceData As String
    Dim myDestinationRange As String
 
    'declare object variables to hold references to destination workbook, source and destination worksheets, and new Pivot Table
    Dim myDestinationWorkbook As Workbook
    Dim mySourceWorksheet As Worksheet
    Dim myDestinationWorksheet As Worksheet
    Dim myPivotTable As PivotTable
 
    'add and identify destination worksheet
    Set myDestinationWorkbook = Workbooks.Add
 
    'identify source and destination worksheets
    Set mySourceWorksheet = ThisWorkbook.Worksheets("Data")
    Set myDestinationWorksheet = myDestinationWorkbook.Worksheets(1)
 
    'obtain address of destination cell range
    myDestinationRange = myDestinationWorksheet.Range("A5").Address(ReferenceStyle:=xlR1C1)
 
    'identify row and column numbers that define source data cell range
    myFirstRow = 5
    myLastRow = 20005
    myFirstColumn = 1
    myLastColumn = 6
 
    'obtain address of source data cell range
    With mySourceWorksheet.Cells
        mySourceData = .Range(.Cells(myFirstRow, myFirstColumn), .Cells(myLastRow, myLastColumn)).Address(ReferenceStyle:=xlR1C1)
    End With
 
    'create Pivot Table cache and create Pivot Table report based on that cache
    Set myPivotTable = myDestinationWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:="[" & ThisWorkbook.Name & "]" & mySourceWorksheet.Name & "!" & mySourceData).CreatePivotTable(TableDestination:="[" & myDestinationWorkbook.Name & "]" & myDestinationWorksheet.Name & "!" & myDestinationRange, TableName:="PivotTableNewWorkbook")
 
    'add, organize and format Pivot Table fields
    With myPivotTable
        .PivotFields("Item").Orientation = xlRowField
        With .PivotFields("Units Sold")
            .Orientation = xlDataField
            .Position = 1
            .Function = xlSum
            .NumberFormat = "#,##0.00"
        End With
        With .PivotFields("Sales Amount")
            .Orientation = xlDataField
            .Position = 2
            .Function = xlSum
            .NumberFormat = "#,##0.00"
        End With
    End With
 
End Sub

Sub createPivotTableDynamicRange()
 
    'Source: https://powerspreadsheets.com/
    'For further information: https://powerspreadsheets.com/vba-create-pivot-table/
 
    'declare variables to hold row and column numbers that define source data cell range
    Dim myFirstRow As Long
    Dim myLastRow As Long
    Dim myFirstColumn As Long
    Dim myLastColumn As Long
 
    'declare variables to hold source and destination cell range address
    Dim mySourceData As String
    Dim myDestinationRange As String
 
    'declare object variables to hold references to source and destination worksheets, and new Pivot Table
    Dim mySourceWorksheet As Worksheet
    Dim myDestinationWorksheet As Worksheet
    Dim myPivotTable As PivotTable
 
    'identify source and destination worksheets
    With ThisWorkbook
        Set mySourceWorksheet = .Worksheets("Data")
        Set myDestinationWorksheet = .Worksheets("DynamicRange")
    End With
 
    'obtain address of destination cell range
    myDestinationRange = myDestinationWorksheet.Range("A5").Address(ReferenceStyle:=xlR1C1)
 
    'identify first row and first column of source data cell range
    myFirstRow = 5
    myFirstColumn = 1
 
    With mySourceWorksheet.Cells
 
        'find last row and last column of source data cell range
        myLastRow = .Find(What:="*", LookIn:=xlFormulas, LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
        myLastColumn = .Find(What:="*", LookIn:=xlFormulas, LookAt:=xlPart, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Column
 
        'obtain address of source data cell range
        mySourceData = .Range(.Cells(myFirstRow, myFirstColumn), .Cells(myLastRow, myLastColumn)).Address(ReferenceStyle:=xlR1C1)
 
    End With
 
    'create Pivot Table cache and create Pivot Table report based on that cache
    Set myPivotTable = ThisWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=mySourceWorksheet.Name & "!" & mySourceData).CreatePivotTable(TableDestination:=myDestinationWorksheet.Name & "!" & myDestinationRange, TableName:="PivotTableExistingSheet")
 
    'add, organize and format Pivot Table fields
    With myPivotTable
        .PivotFields("Item").Orientation = xlRowField
        With .PivotFields("Units Sold")
            .Orientation = xlDataField
            .Position = 1
            .Function = xlSum
            .NumberFormat = "#,##0.00"
        End With
        With .PivotFields("Sales Amount")
            .Orientation = xlDataField
            .Position = 2
            .Function = xlSum
            .NumberFormat = "#,##0.00"
        End With
    End With
 
End Sub
