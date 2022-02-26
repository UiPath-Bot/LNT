Sub Insert_PivotTable_Critical_Repeated()
'Declare Variables
Dim PSheet As Worksheet
Dim DSheet As Worksheet
Dim PCache As PivotCache
Dim PField As PivotField
Dim PTable As PivotTable
Dim DRange As Range
Dim LastRow As Long
Dim LastCol As Long

'Insert a New Blank Worksheet
On Error Resume Next
Application.DisplayAlerts = False
Worksheets("Critical_Repeated_Analysis").Delete
Sheets.Add After:=ActiveSheet
ActiveSheet.Name = "Critical_Repeated_Analysis"
Application.DisplayAlerts = False
Set PSheet = Worksheets("Critical_Repeated_Analysis")
Set DSheet = Worksheets("Critical_Repeated")

'Define Data Range
LastRow = DSheet.Cells(Rows.Count, 1).End(xlUp).Row
LastCol = DSheet.Cells(1, Columns.Count).End(xlToLeft).Column
Set DRange = DSheet.Cells(1, 1).Resize(LastRow, LastCol)

'Define Pivot Cache
Set PCache = ActiveWorkbook.PivotCaches.Create _
(SourceType:=xlDatabase, SourceData:=DRange). _
CreatePivotTable(TableDestination:=PSheet.Cells(1, 1), _
TableName:="CriticalRepeatedPivotTable")

'Insert Blank Pivot Table
Set PTable = PCache.CreatePivotTable _
(TableDestination:=PSheet.Cells(1, 1), TableName:="CriticalRepeatedPivotTable")

'Insert Row Fields
With ActiveSheet.PivotTables("CriticalRepeatedPivotTable").PivotFields("Name")
.Orientation = xlRowField
.Position = 1
End With

With ActiveSheet.PivotTables("CriticalRepeatedPivotTable").PivotFields("Ownership")
.Orientation = xlRowField
.Position = 2
End With

'Insert Column Fields
With ActiveSheet.PivotTables("CriticalRepeatedPivotTable").PivotFields("Month")
.Orientation = xlColumnField
.Position = 1
End With

'Insert Data Field
With ActiveSheet.PivotTables("CriticalRepeatedPivotTable").PivotFields("Host")
.Orientation = xlDataField
.Function = xlCount
.NumberFormat = "#,##0"
.Name = "Count of Host"
End With

On Error Resume Next
Application.ScreenUpdating = False
 
Set PTable = ActiveSheet.PivotTables("CriticalRepeatedPivotTable")
PTable.ManualUpdate = True
  
'This section applies Classic PivotTable settings
'and turns off the Contextual Tooltips and the Expand/Collapse buttons
'With ActiveSheet.PivotTables("HighAnalysisPivotTable")
With PTable
    .InGridDropZones = True
    .RowAxisLayout xlTabularRow
    .DisplayContextTooltips = False
    .ShowDrillIndicators = False
End With


PTable.ManualUpdate = False
Application.ScreenUpdating = True

'This ensures that only data that still exists in the data
'     that drives the PivotTable
'     will appear in the PivotTable dropdown lists
PTable.PivotCache.MissingItemsLimit = xlMissingItemsNone

'ActiveSheet.PivotTables("HighAnalysisPivotTable").ShowTableStyleRowStripes = True
'ActiveSheet.PivotTables("HighAnalysisPivotTable").ShowTableStyleColumnStripes = True
'ActiveSheet.PivotTables("HighAnalysisPivotTable").TableStyle2 = "PivotStyleMedium9"

Application.CutCopyMode = False
    ActiveSheet.PivotTables("CriticalRepeatedPivotTable").PivotSelect "", xlDataAndLabel, True
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    With Selection
        .HorizontalAlignment = xlCenter
    End With
    
'first, ensure that no panes are frozen
With ActiveSheet.PivotTables("CriticalRepeatedPivotTable").PivotSelect
ActiveWindow.FreezePanes = False
'select the row that you want to freeze based on
ActiveSheet.Range("C3").Select

'freeze panes
ActiveWindow.FreezePanes = True
End With
'For selecting None
'turns off subtotals in pivot table

On Error Resume Next
With ActiveSheet.PivotTables("CriticalRepeatedPivotTable").PivotSelect
For Each PTable In ActiveSheet.PivotTables
  PTable.ManualUpdate = True
  For Each PField In PTable.PivotFields
    'First, set index 1 (Automatic) to True,
    'so all other values are set to False
    PField.Subtotals(1) = True
    PField.Subtotals(1) = False
  Next PField
  PTable.ManualUpdate = False
Next PTable
End With

End Sub


