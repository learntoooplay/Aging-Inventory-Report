' Approx date August 2019
' John Theodorakis
' File will take raw sales data and output cost analysis of stored models. No sample i/o for data security reasons.



Sub inventoryAgingReport()

sheetToData
dataToDivvy
ongoingSKU
EOLInventory
BTVRefurb
dataToSummary

End Sub


Sub sheetToData()
' Copy needed data to new sheet called "Data" that we can pull for pivot tables

' Filter data and copy needed data to new sheet called "Data"
    Range("A1").Select
    ActiveSheet.Range("$A$1:$N$19122").AutoFilter Field:=5, Criteria1:=Array( _
        "V102", "V109", "V111", "V113"), Operator:=xlFilterValues
    ActiveSheet.Range("$A$1:$N$19122").AutoFilter Field:=6, Criteria1:=Array( _
        "1101", "2001", "2002", "2101", "2102", "2103", "2104", "2105", "2106", "2108", "2110", "2112", "2115", "2117", "2118", "2119", "2122", "2123", "2124", "2125", "2126", "2127", "2128", "2129", "3001", "4001"), Operator:=xlFilterValues
    Columns("A:N").Select
    Selection.Copy
    Sheets.Add.Name = "Data"
    Sheets("Data").Select
    Columns("A:N").Select
    ActiveSheet.Paste
' Name necessary columns on "Data"
    Range("O1").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "Total"
    Range("P1").Select
    ActiveCell.FormulaR1C1 = "Data"
    Range("Q1").Select
    ActiveCell.FormulaR1C1 = "Group"
    Range("R1").Select
    ActiveCell.FormulaR1C1 = "ODM"
' Calculate total and create line data that can be used to pull information from previous documents
    Range("O2").Select
    ActiveCell.FormulaR1C1 = "=SUM(RC[-7],RC[-6],RC[-5],RC[-4],RC[-3],RC[-2])"
    Range("P2").Select
    ActiveCell.FormulaR1C1 = "=RC[-15]&RC[-11]&RC[-10]"
' Pull group info based on the above line data, then determine how many units are ODM units based on storage locatoin
    Range("Q2").Select
    ActiveCell.FormulaR1C1 = _
        "=VLOOKUP(RC[-1],'[2019-09-18-Inventory-Aging-Report.xlsm]Data'!C16:C17,2,FALSE)"
    Range("R2").Select
    ActiveCell.FormulaR1C1 = "=IF(OR(RC[-12]=""2101"",RC[-12]=""2119"",RC[-12]=""2123"",RC[-12]=""2128"",AND(RC[-13]=""V113"",RC[-12]=""4001"")),RC[-3],0)"
' Auto fill columns O, P, Q, and R based on the length of the original data
    Dim startCell As Range, lastRow As Long
    Set startCell = Range("B4")
    lastRow = Cells(Rows.Count, startCell.Column).End(xlUp).Row
    Range("O2").Select
    Selection.AutoFill Destination:=Range(Cells(2, 15), Cells(lastRow, 15))
    Range("P2").Select
    Selection.AutoFill Destination:=Range(Cells(2, 16), Cells(lastRow, 16))
    Range("Q2").Select
    Selection.AutoFill Destination:=Range(Cells(2, 17), Cells(lastRow, 17))
    Range("R2").Select
    Selection.AutoFill Destination:=Range(Cells(2, 18), Cells(lastRow, 18))
    Range("A1").Select

End Sub


Sub dataToDivvy()
' Create new sheet, "Divvy", with information on older models stored in a specific plant

' Create new sheet called "Divvy"
    Sheets.Add.Name = "Divvy"
' Create pivot table in "Divvy" based on data from "Data" and format accordingly
    Sheets("Data").Select
    Cells.Select
    ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
        "Data!R1C1:R1048576C18", Version:=6).CreatePivotTable TableDestination:= _
        "Divvy!R1C1", TableName:="PivotTable28", DefaultVersion:=6
    Sheets("Divvy").Select
    Cells(1, 1).Select
    With ActiveSheet.PivotTables("PivotTable28")
        .ColumnGrand = True
        .HasAutoFormat = True
        .DisplayErrorString = False
        .DisplayNullString = True
        .EnableDrilldown = True
        .ErrorString = ""
        .MergeLabels = False
        .NullString = ""
        .PageFieldOrder = 2
        .PageFieldWrapCount = 0
        .PreserveFormatting = True
        .RowGrand = True
        .SaveData = True
        .PrintTitles = False
        .RepeatItemsOnEachPrintedPage = True
        .TotalsAnnotation = False
        .CompactRowIndent = 1
        .InGridDropZones = False
        .DisplayFieldCaptions = True
        .DisplayMemberPropertyTooltips = False
        .DisplayContextTooltips = True
        .ShowDrillIndicators = True
        .PrintDrillIndicators = False
        .AllowMultipleFilters = False
        .SortUsingCustomLists = True
        .FieldListSortAscending = False
        .ShowValuesRow = False
        .CalculatedMembersInFilters = False
        .RowAxisLayout xlCompactRow
    End With
    With ActiveSheet.PivotTables("PivotTable28").PivotCache
        .RefreshOnFileOpen = False
        .MissingItemsLimit = xlMissingItemsDefault
    End With
    ActiveSheet.PivotTables("PivotTable28").RepeatAllLabels xlRepeatLabels
' Create filters for group, plant (V113), and storage location
    With ActiveSheet.PivotTables("PivotTable28").PivotFields("Group")
        .Orientation = xlPageField
        .Position = 1
    End With
    With ActiveSheet.PivotTables("PivotTable28").PivotFields("Plant")
        .Orientation = xlPageField
        .Position = 1
    End With
    ActiveSheet.PivotTables("PivotTable28").PivotFields("Plant").ClearAllFilters
    ActiveSheet.PivotTables("PivotTable28").PivotFields("Plant").CurrentPage = _
        "V113"
    With ActiveSheet.PivotTables("PivotTable28").PivotFields("Storage Loc.")
        .Orientation = xlPageField
        .Position = 1
    End With
' Set rows to "old material numbers"
    With ActiveSheet.PivotTables("PivotTable28").PivotFields("Old matl number")
        .Orientation = xlRowField
        .Position = 1
' Set columns of pivot table and filter
    End With
    ActiveSheet.PivotTables("PivotTable28").AddDataField ActiveSheet.PivotTables( _
        "PivotTable28").PivotFields("0-30 days"), "Sum of 0-30 days", xlSum
    ActiveSheet.PivotTables("PivotTable28").AddDataField ActiveSheet.PivotTables( _
        "PivotTable28").PivotFields(">30-60 days"), "Sum of >30-60 days", xlSum
    ActiveSheet.PivotTables("PivotTable28").AddDataField ActiveSheet.PivotTables( _
        "PivotTable28").PivotFields(">60-90 days"), "Sum of >60-90 days", xlSum
    ActiveSheet.PivotTables("PivotTable28").AddDataField ActiveSheet.PivotTables( _
        "PivotTable28").PivotFields(">90-120 days"), "Sum of >90-120 days", xlSum
    ActiveSheet.PivotTables("PivotTable28").AddDataField ActiveSheet.PivotTables( _
        "PivotTable28").PivotFields(">120-150 days"), "Sum of >120-150 days", xlSum
    ActiveSheet.PivotTables("PivotTable28").AddDataField ActiveSheet.PivotTables( _
        "PivotTable28").PivotFields(">150 days"), "Sum of >150 days", xlSum
    ActiveSheet.PivotTables("PivotTable28").AddDataField ActiveSheet.PivotTables( _
        "PivotTable28").PivotFields("Total"), "Sum of Total", xlSum
    Range("A5").Select
    ActiveSheet.PivotTables("PivotTable28").PivotFields("Old matl number"). _
        PivotFilters.Add2 Type:=xlValueDoesNotEqual, DataField:=ActiveSheet. _
        PivotTables("PivotTable28").PivotFields("Sum of Total"), Value1:=0
' Add column for sum of ODM data and copy to column M, then hide sum of ODM (otherwise error will occur when attempting to populate column I)
    ActiveSheet.PivotTables("PivotTable28").AddDataField ActiveSheet.PivotTables( _
        "PivotTable28").PivotFields("ODM"), "Sum of ODM", xlSum
    Columns("I:I").Select
    Selection.Copy
    Columns("M:M").Select
    ActiveSheet.Paste
    ActiveSheet.PivotTables("PivotTable28").PivotFields("Sum of ODM").Orientation _
    = xlHidden
' Label Columns
    Range("M3").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "ODM"
    'Range("H11").Select
    Range("I3").Select
    ActiveCell.FormulaR1C1 = "Orders"
    Range("J3").Select
    ActiveCell.FormulaR1C1 = "Balance"
    Range("L3").Select
    ActiveCell.FormulaR1C1 = "VIZIO"
' Change cell colors
    Range("I3").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 5296274
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    Range("L3:M3").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 13311
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
' Enter formulas for how many units below to VIZIO and how many belong to ODMs
    Range("L4").Select
    ActiveCell.FormulaR1C1 = "=RC[-4]-RC[1]"
    Range("L4").Select
    Range("M4").Select
    ActiveCell.FormulaR1C1 = "=IFERROR(VLOOKUP(RC1,C16:C23,8,FALSE),0)"
' Auto fill columns L and M and add sum for the column at the last row
    Range("L4").Select
    Dim startCell1 As Range, lastRow1 As Long
    Set startCell1 = Range("H4")
    lastRow1 = Cells(Rows.Count, startCell1.Column).End(xlUp).Row
    Selection.AutoFill Destination:=Range(Cells(4, 12), Cells(lastRow1, 12))
    Range("M4").Select
    Selection.AutoFill Destination:=Range(Cells(4, 13), Cells(lastRow1 - 1, 13))
    Range(Cells(lastRow1, 13), Cells(lastRow1, 13)).Select
    Range(Cells(lastRow1, 13), Cells(lastRow1, 13)).Formula = "=SUM(" & Range(Cells(4, 13), Cells(lastRow1 - 1, 13)).Address(False, False) & ")"
' Enter formula for balance, auto fill column, and add sum for column at the last row
    Range("J4").Select
    ActiveCell.FormulaR1C1 = "=RC[-2]-RC[-1]"
    Range("J4").Select
    Selection.AutoFill Destination:=Range(Cells(4, 10), Cells(lastRow1 - 1, 10)), Type:=xlFillDefault
    Range(Cells(lastRow1, 10), Cells(lastRow1, 10)).Select
    Range(Cells(lastRow1, 10), Cells(lastRow1, 10)).Formula = "=SUM(" & Range(Cells(4, 10), Cells(lastRow1 - 1, 10)).Address(False, False) & ")"
    Sheets("Data").Select
    Range("A1").Select
    
' Select all cells in "Data" and create a new pivot table in "Divvy"
    Cells.Select
    ActiveWorkbook.Worksheets("Divvy").PivotTables("PivotTable28").PivotCache. _
        CreatePivotTable TableDestination:="Divvy!R1C16", TableName:="PivotTable34" _
        , DefaultVersion:=6
    Sheets("Divvy").Select
    Cells(1, 16).Select
    With ActiveSheet.PivotTables("PivotTable34")
        .ColumnGrand = True
        .HasAutoFormat = True
        .DisplayErrorString = False
        .DisplayNullString = True
        .EnableDrilldown = True
        .ErrorString = ""
        .MergeLabels = False
        .NullString = ""
        .PageFieldOrder = 2
        .PageFieldWrapCount = 0
        .PreserveFormatting = True
        .RowGrand = True
        .SaveData = True
        .PrintTitles = False
        .RepeatItemsOnEachPrintedPage = True
        .TotalsAnnotation = False
        .CompactRowIndent = 1
        .InGridDropZones = False
        .DisplayFieldCaptions = True
        .DisplayMemberPropertyTooltips = False
        .DisplayContextTooltips = True
        .ShowDrillIndicators = True
        .PrintDrillIndicators = False
        .AllowMultipleFilters = False
        .SortUsingCustomLists = True
        .FieldListSortAscending = False
        .ShowValuesRow = False
        .CalculatedMembersInFilters = False
        .RowAxisLayout xlCompactRow
    End With
    With ActiveSheet.PivotTables("PivotTable34").PivotCache
        .RefreshOnFileOpen = False
        .MissingItemsLimit = xlMissingItemsDefault
    End With
    ActiveSheet.PivotTables("PivotTable34").RepeatAllLabels xlRepeatLabels
'   Create filters for group, plant (V113), and storage location (not 1101, 2001, 2002)
    With ActiveSheet.PivotTables("PivotTable34").PivotFields("Group")
        .Orientation = xlPageField
        .Position = 1
    End With
    With ActiveSheet.PivotTables("PivotTable34").PivotFields("Plant")
        .Orientation = xlPageField
        .Position = 1
    End With
    ActiveSheet.PivotTables("PivotTable34").PivotFields("Plant").ClearAllFilters
    ActiveSheet.PivotTables("PivotTable34").PivotFields("Plant").CurrentPage = _
        "V113"
    With ActiveSheet.PivotTables("PivotTable34").PivotFields("Storage Loc.")
        .Orientation = xlPageField
        .Position = 1
    End With
     ActiveSheet.PivotTables("PivotTable34").PivotFields("Storage Loc."). _
        CurrentPage = "(All)"
    With ActiveSheet.PivotTables("PivotTable34").PivotFields("Storage Loc.")
        .PivotItems("1101").Visible = False
        .PivotItems("2001").Visible = False
        .PivotItems("2002").Visible = False
    End With
' Set rows to "old material numbers"
    Range("Q3").Select
    With ActiveSheet.PivotTables("PivotTable34").PivotFields("Old matl number")
        .Orientation = xlRowField
        .Position = 1
' Set columns of pivot table and filter
    End With
    ActiveSheet.PivotTables("PivotTable34").AddDataField ActiveSheet.PivotTables( _
        "PivotTable34").PivotFields("0-30 days"), "Sum of 0-30 days", xlSum
    ActiveSheet.PivotTables("PivotTable34").AddDataField ActiveSheet.PivotTables( _
        "PivotTable34").PivotFields(">30-60 days"), "Sum of >30-60 days", xlSum
    ActiveSheet.PivotTables("PivotTable34").AddDataField ActiveSheet.PivotTables( _
        "PivotTable34").PivotFields(">60-90 days"), "Sum of >60-90 days", xlSum
    ActiveSheet.PivotTables("PivotTable34").AddDataField ActiveSheet.PivotTables( _
        "PivotTable34").PivotFields(">90-120 days"), "Sum of >90-120 days", xlSum
    ActiveSheet.PivotTables("PivotTable34").AddDataField ActiveSheet.PivotTables( _
        "PivotTable34").PivotFields(">120-150 days"), "Sum of >120-150 days", xlSum
    ActiveSheet.PivotTables("PivotTable34").AddDataField ActiveSheet.PivotTables( _
        "PivotTable34").PivotFields(">150 days"), "Sum of >150 days", xlSum
    ActiveSheet.PivotTables("PivotTable34").AddDataField ActiveSheet.PivotTables( _
        "PivotTable34").PivotFields("Total"), "Sum of Total", xlSum
    ActiveSheet.PivotTables("PivotTable34").PivotFields("Old matl number"). _
        PivotFilters.Add2 Type:=xlValueDoesNotEqual, DataField:=ActiveSheet. _
        PivotTables("PivotTable34").PivotFields("Sum of Total"), Value1:=0
    Sheets("Data").Select
    Range("A1").Select

End Sub


Sub ongoingSKU()
' Add new sheet, "Summary", that includes information on models that are still in production

' Add new sheet called "Summary"
    Sheets.Add.Name = "Summary"
' Create new pivot table in "Summary" based on "Data"
    Sheets("Data").Select
    Cells.Select
    ActiveWorkbook.Worksheets("Divvy").PivotTables("PivotTable34").PivotCache. _
        CreatePivotTable TableDestination:="Summary!R1C1", TableName:= _
        "PivotTable38", DefaultVersion:=6
    Sheets("Summary").Select
    Cells(1, 1).Select
    With ActiveSheet.PivotTables("PivotTable38")
        .ColumnGrand = True
        .HasAutoFormat = True
        .DisplayErrorString = False
        .DisplayNullString = True
        .EnableDrilldown = True
        .ErrorString = ""
        .MergeLabels = False
        .NullString = ""
        .PageFieldOrder = 2
        .PageFieldWrapCount = 0
        .PreserveFormatting = True
        .RowGrand = True
        .SaveData = True
        .PrintTitles = False
        .RepeatItemsOnEachPrintedPage = True
        .TotalsAnnotation = False
        .CompactRowIndent = 1
        .InGridDropZones = False
        .DisplayFieldCaptions = True
        .DisplayMemberPropertyTooltips = False
        .DisplayContextTooltips = True
        .ShowDrillIndicators = True
        .PrintDrillIndicators = False
        .AllowMultipleFilters = False
        .SortUsingCustomLists = True
        .FieldListSortAscending = False
        .ShowValuesRow = False
        .CalculatedMembersInFilters = False
        .RowAxisLayout xlCompactRow
    End With
    With ActiveSheet.PivotTables("PivotTable38").PivotCache
        .RefreshOnFileOpen = False
        .MissingItemsLimit = xlMissingItemsDefault
    End With
' Create filters for group (BTV), plant (not V113 or blank), and storage location (not 3001 or 4001)
    ActiveSheet.PivotTables("PivotTable38").RepeatAllLabels xlRepeatLabels
    With ActiveSheet.PivotTables("PivotTable38").PivotFields("Group")
        .Orientation = xlPageField
        .Position = 1
    End With
    ActiveSheet.PivotTables("PivotTable38").PivotFields("Group").ClearAllFilters
    ActiveSheet.PivotTables("PivotTable38").PivotFields("Group").CurrentPage = _
        "BTV"
    With ActiveSheet.PivotTables("PivotTable38").PivotFields("Plant")
        .Orientation = xlPageField
        .Position = 1
    End With
    ActiveSheet.PivotTables("PivotTable38").PivotFields("Plant").CurrentPage = _
        "(All)"
    With ActiveSheet.PivotTables("PivotTable38").PivotFields("Plant")
        .PivotItems("V113").Visible = False
        .PivotItems("(blank)").Visible = False
    End With
    ActiveSheet.PivotTables("PivotTable38").PivotFields("Plant"). _
        EnableMultiplePageItems = True
    With ActiveSheet.PivotTables("PivotTable38").PivotFields("Storage Loc.")
        .Orientation = xlPageField
        .Position = 1
    End With
    ActiveSheet.PivotTables("PivotTable38").PivotFields("Storage Loc."). _
        CurrentPage = "(All)"
    With ActiveSheet.PivotTables("PivotTable38").PivotFields("Storage Loc.")
        .PivotItems("3001").Visible = False
        .PivotItems("4001").Visible = False
    End With
    ActiveSheet.PivotTables("PivotTable38").PivotFields("Storage Loc."). _
        EnableMultiplePageItems = True
' Set rows to "old material numbers"
    With ActiveSheet.PivotTables("PivotTable38").PivotFields("Old matl number")
        .Orientation = xlRowField
        .Position = 1
' Set columns of pivot table and filter
    End With
    ActiveSheet.PivotTables("PivotTable38").AddDataField ActiveSheet.PivotTables( _
        "PivotTable38").PivotFields("0-30 days"), "Sum of 0-30 days", xlSum
    ActiveSheet.PivotTables("PivotTable38").AddDataField ActiveSheet.PivotTables( _
        "PivotTable38").PivotFields(">30-60 days"), "Sum of >30-60 days", xlSum
    ActiveSheet.PivotTables("PivotTable38").AddDataField ActiveSheet.PivotTables( _
        "PivotTable38").PivotFields(">60-90 days"), "Sum of >60-90 days", xlSum
    ActiveSheet.PivotTables("PivotTable38").AddDataField ActiveSheet.PivotTables( _
        "PivotTable38").PivotFields(">90-120 days"), "Sum of >90-120 days", xlSum
    ActiveSheet.PivotTables("PivotTable38").AddDataField ActiveSheet.PivotTables( _
        "PivotTable38").PivotFields(">120-150 days"), "Sum of >120-150 days", xlSum
    ActiveSheet.PivotTables("PivotTable38").AddDataField ActiveSheet.PivotTables( _
        "PivotTable38").PivotFields(">150 days"), "Sum of >150 days", xlSum
    ActiveSheet.PivotTables("PivotTable38").AddDataField ActiveSheet.PivotTables( _
        "PivotTable38").PivotFields("Total"), "Sum of Total", xlSum
    ActiveSheet.PivotTables("PivotTable38").AddDataField ActiveSheet.PivotTables( _
        "PivotTable38").PivotFields("ODM"), "Sum of ODM", xlSum
    ActiveSheet.PivotTables("PivotTable38").PivotFields("Old matl number"). _
        PivotFilters.Add2 Type:=xlValueDoesNotEqual, DataField:=ActiveSheet. _
        PivotTables("PivotTable38").PivotFields("Sum of Total"), Value1:=0
' Start new table "ongoing SKUs" and copy pivot table column headers to "Ongoing SKUs" table
    Range("N2").Select
    ActiveCell.FormulaR1C1 = "Ongoing SKUs"
    Range("A3:H3").Select
    Selection.Copy
    Range("N3:U3").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    
' Copy material numbers from pivot table on
    Range("A4:A25").Select
    Selection.Copy
    Range("N4:N25").Select
    ActiveSheet.Paste
' Use VLOOKUP to add values from the pivot table and sum formula to the "Ongoing SKUs" table
    Range("O4").Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP(RC14,C1:C8,2,FALSE)"
    Range("O4").Select
    Selection.AutoFill Destination:=Range("O4:U4"), Type:=xlFillDefault
    Range("O4:U4").Select
    Range("P4").Select
    ActiveCell.FormulaR1C1 = "=IFERROR(VLOOKUP(RC14,C1:C8,3,FALSE),0)"
    Range("Q4").Select
    ActiveCell.FormulaR1C1 = "=IFERROR(VLOOKUP(RC14,C1:C8,4,FALSE),0)"
    Range("R4").Select
    ActiveCell.FormulaR1C1 = "=IFERROR(VLOOKUP(RC14,C1:C8,5,FALSE),0)"
    Range("S4").Select
    ActiveCell.FormulaR1C1 = "=IFERROR(VLOOKUP(RC14,C1:C8,6,FALSE),0)"
    Range("T4").Select
    ActiveCell.FormulaR1C1 = "=IFERROR(VLOOKUP(RC14,C1:C8,7,FALSE),0)"
    Range("U4").Select
    ActiveCell.FormulaR1C1 = "=SUM(RC[-6]:RC[-1])"
    Range("O4").Select
    ActiveCell.FormulaR1C1 = "=IFERROR(VLOOKUP(RC[-1],C1:C8,2,FALSE),0)"
' Auto fill the cells in "Ongoing SKUs" based on the number of material numbers (column N)
    Dim startCell1 As Range, lastRow1 As Long
    Set startCell1 = Range("N4")
    lastRow1 = Cells(Rows.Count, startCell1.Column).End(xlUp).Row
    Range("O4").Select
    Selection.AutoFill Destination:=Range(Cells(4, 15), Cells(lastRow1, 15)), Type:=xlFillDefault
    Range("P4").Select
    ActiveCell.FormulaR1C1 = "=IFERROR(VLOOKUP(RC[-2],C1:C8,3,FALSE),0)"
    Range("P4").Select
    Selection.AutoFill Destination:=Range(Cells(4, 16), Cells(lastRow1, 16)), Type:=xlFillDefault
    Range("Q4").Select
    ActiveCell.FormulaR1C1 = "=IFERROR(VLOOKUP(RC[-3],C1:C8,4,FALSE),0)"
    Range("Q4").Select
    Selection.AutoFill Destination:=Range(Cells(4, 17), Cells(lastRow1, 17)), Type:=xlFillDefault
    Range("R4").Select
    ActiveCell.FormulaR1C1 = "=IFERROR(VLOOKUP(RC[-4],C1:C8,5,FALSE),0)"
    Range("R4").Select
    Selection.AutoFill Destination:=Range(Cells(4, 18), Cells(lastRow1, 18)), Type:=xlFillDefault
    Range("S4").Select
    ActiveCell.FormulaR1C1 = "=IFERROR(VLOOKUP(RC[-5],C1:C8,6,FALSE),0)"
    Range("S4").Select
    Selection.AutoFill Destination:=Range(Cells(4, 19), Cells(lastRow1, 19)), Type:=xlFillDefault
    Range("T4").Select
    ActiveCell.FormulaR1C1 = "=IFERROR(VLOOKUP(RC[-6],C1:C8,7,FALSE),0)"
    Range("T4").Select
    Selection.AutoFill Destination:=Range(Cells(4, 20), Cells(lastRow1, 20)), Type:=xlFillDefault
    Range("U4").Select
    Selection.AutoFill Destination:=Range(Cells(4, 21), Cells(lastRow1, 21)), Type:=xlFillDefault
' Copy last row of the pivot table to the last row of "Ongoing SKUs" table
    Dim startCell2 As Range, lastRow2 As Long
    Set startCell2 = Range("A4")
    lastRow2 = Cells(Rows.Count, startCell2.Column).End(xlUp).Row
    Range(Cells(lastRow2, 1), Cells(lastRow2, 8)).Select
    Selection.Copy
    Range(Cells(lastRow1 + 1, 14), Cells(lastRow1 + 1, 20)).Select
    ActiveSheet.Paste
    Range(Cells(lastRow1 + 1, 15), Cells(lastRow1 + 1, 15)).Select
    Application.CutCopyMode = False
' Create formula for totals and autofill to the cells in the last row of "Ongoing SKUs"
    ActiveCell.FormulaR1C1 = "=SUM(R[-23]C:R[-1]C)"
    Range(Cells(lastRow1 + 1, 15), Cells(lastRow1 + 1, 15)).Select
    Selection.AutoFill Destination:=Range(Cells(lastRow1 + 1, 15), Cells(lastRow1 + 1, 21)), Type:=xlFillDefault
' Create new columns for "Ongoing SKUs" table and format cells accordingly
    Range("V3").Select
    ActiveCell.FormulaR1C1 = "Storage cost"
    Range("W3").Select
    ActiveCell.FormulaR1C1 = "Monthly cost"
    Range("X3").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternThemeColor = xlThemeColorAccent1
        .ThemeColor = xlThemeColorAccent1
        .TintAndShade = 0.799981688894314
        .PatternTintAndShade = 0.799981688894314
    End With
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    Selection.Borders(xlEdgeLeft).LineStyle = xlNone
    Selection.Borders(xlEdgeTop).LineStyle = xlNone
    Selection.Borders(xlEdgeBottom).LineStyle = xlNone
    Selection.Borders(xlEdgeRight).LineStyle = xlNone
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    ActiveCell.FormulaR1C1 = "VIZIO"
    Range("Y3").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternThemeColor = xlThemeColorAccent1
        .ThemeColor = xlThemeColorAccent1
        .TintAndShade = 0.799981688894314
        .PatternTintAndShade = 0.799981688894314
    End With
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    Selection.Borders(xlEdgeLeft).LineStyle = xlNone
    Selection.Borders(xlEdgeTop).LineStyle = xlNone
    Selection.Borders(xlEdgeBottom).LineStyle = xlNone
    Selection.Borders(xlEdgeRight).LineStyle = xlNone
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    ActiveCell.FormulaR1C1 = "VIZIO cost"
    Range("Z3").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternThemeColor = xlThemeColorAccent1
        .ThemeColor = xlThemeColorAccent1
        .TintAndShade = 0.799981688894314
        .PatternTintAndShade = 0.799981688894314
    End With
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    Selection.Borders(xlEdgeLeft).LineStyle = xlNone
    Selection.Borders(xlEdgeTop).LineStyle = xlNone
    Selection.Borders(xlEdgeBottom).LineStyle = xlNone
    Selection.Borders(xlEdgeRight).LineStyle = xlNone
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    ActiveCell.FormulaR1C1 = "ODM"
    Range("AA3").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternThemeColor = xlThemeColorAccent1
        .ThemeColor = xlThemeColorAccent1
        .TintAndShade = 0.799981688894314
        .PatternTintAndShade = 0.799981688894314
    End With
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    Selection.Borders(xlEdgeLeft).LineStyle = xlNone
    Selection.Borders(xlEdgeTop).LineStyle = xlNone
    Selection.Borders(xlEdgeBottom).LineStyle = xlNone
    Selection.Borders(xlEdgeRight).LineStyle = xlNone
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    ActiveCell.FormulaR1C1 = "ODM cost"
' Auto fill column for ODM based on pivot table values
    Range("Z4").Select
    ActiveCell.FormulaR1C1 = "=IFERROR(VLOOKUP(RC[-12],R4C1:R29C9,9,FALSE),0)"
    Range("Z4").Select
    Selection.AutoFill Destination:=Range(Cells(4, 26), Cells(lastRow1, 26)), Type:=xlFillDefault
    Range(Cells(4, 26), Cells(lastRow1, 26)).Select
' Auto fill VIZIO column with necessary formula
    Range("X4").Select
    Application.CutCopyMode = False
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "=RC[-3]-RC[2]"
    Range("X4").Select
    Selection.AutoFill Destination:=Range(Cells(4, 24), Cells(lastRow1, 24)), Type:=xlFillDefault
    Range(Cells(4, 24), Cells(lastRow1, 24)).Select
    Range(Cells(lastRow1 + 1, 24), Cells(lastRow1 + 1, 27)).Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternThemeColor = xlThemeColorAccent1
        .ThemeColor = xlThemeColorAccent1
        .TintAndShade = 0.799981688894314
        .PatternTintAndShade = 0.799981688894314
    End With
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    Selection.Borders(xlEdgeLeft).LineStyle = xlNone
    Selection.Borders(xlEdgeTop).LineStyle = xlNone
    Selection.Borders(xlEdgeBottom).LineStyle = xlNone
    Selection.Borders(xlEdgeRight).LineStyle = xlNone
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    Range(Cells(lastRow1 + 1, 24), Cells(lastRow1 + 1, 24)).Select
    ActiveCell.FormulaR1C1 = "=SUM(R[-23]C:R[-1]C)"
    Range(Cells(lastRow1 + 1, 26), Cells(lastRow1 + 1, 26)).Select
    ActiveCell.FormulaR1C1 = "=SUM(R[-23]C:R[-1]C)"
    
' Add formula to VIZIO cost and auto fill column in table
    Range("Y4").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "=RC[-2]*RC[-1]"
    Dim startCell3 As Range, lastRow3 As Long
    Set startCell3 = Range("N4")
    lastRow3 = Cells(Rows.Count, startCell3.Column).End(xlUp).Row
    Range("Y4").Select
    Selection.AutoFill Destination:=Range(Cells(4, 25), Cells(lastRow3 - 1, 25)), Type:=xlFillDefault
' Add formula to ODM cost and auto fill column in table
    Range("AA4").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "=RC[-4]*RC[-1]"
    Range("AA4").Select
    Selection.AutoFill Destination:=Range(Cells(4, 27), Cells(lastRow3 - 1, 27)), Type:=xlFillDefault
' Add total formulas to last row cells
    Range(Cells(lastRow3, 25), Cells(lastRow3, 25)).Select
    ActiveCell.FormulaR1C1 = "=SUM(R[-23]C:R[-1]C)"
    Range(Cells(lastRow3, 27), Cells(lastRow3, 27)).Select
    ActiveCell.FormulaR1C1 = "=SUM(R[-23]C:R[-1]C)"
' Change format of cells
    Range(Cells(4, 25), Cells(lastRow3, 25)).Select
    Selection.Style = "Currency"
    Range(Cells(4, 27), Cells(lastRow3, 27)).Select
    Selection.Style = "Currency"
    Range("X3:AA3").Select
    Selection.Font.Bold = True
    Range(Cells(lastRow3, 24), Cells(lastRow3, 27)).Select
    Selection.Font.Bold = True

End Sub


Sub EOLInventory()
' Create pivot table and another table in "Summary" that includes information on EOL models

' Create new pivot table in "Summary" starting at row 36
    Sheets("Data").Select
    ActiveWorkbook.Worksheets("Divvy").PivotTables("PivotTable34").PivotCache. _
        CreatePivotTable TableDestination:="Summary!R36C1", TableName:= _
        "PivotTable39", DefaultVersion:=6
    Sheets("Summary").Select
    Cells(36, 1).Select
    With ActiveSheet.PivotTables("PivotTable39")
        .ColumnGrand = True
        .HasAutoFormat = True
        .DisplayErrorString = False
        .DisplayNullString = True
        .EnableDrilldown = True
        .ErrorString = ""
        .MergeLabels = False
        .NullString = ""
        .PageFieldOrder = 2
        .PageFieldWrapCount = 0
        .PreserveFormatting = True
        .RowGrand = True
        .SaveData = True
        .PrintTitles = False
        .RepeatItemsOnEachPrintedPage = True
        .TotalsAnnotation = False
        .CompactRowIndent = 1
        .InGridDropZones = False
        .DisplayFieldCaptions = True
        .DisplayMemberPropertyTooltips = False
        .DisplayContextTooltips = True
        .ShowDrillIndicators = True
        .PrintDrillIndicators = False
        .AllowMultipleFilters = False
        .SortUsingCustomLists = True
        .FieldListSortAscending = False
        .ShowValuesRow = False
        .CalculatedMembersInFilters = False
        .RowAxisLayout xlCompactRow
    End With
    With ActiveSheet.PivotTables("PivotTable39").PivotCache
        .RefreshOnFileOpen = False
        .MissingItemsLimit = xlMissingItemsDefault
    End With
' Create filters for group (BTV-EOL), plant (not V113 or blank), and storage location (not 3001 or 4001)
    ActiveSheet.PivotTables("PivotTable39").RepeatAllLabels xlRepeatLabels
    With ActiveSheet.PivotTables("PivotTable39").PivotFields("Group")
        .Orientation = xlPageField
        .Position = 1
    End With
    With ActiveSheet.PivotTables("PivotTable39").PivotFields("Plant")
        .Orientation = xlPageField
        .Position = 1
    End With
    With ActiveSheet.PivotTables("PivotTable39").PivotFields("Storage Loc.")
        .Orientation = xlPageField
        .Position = 1
    End With
    ActiveSheet.PivotTables("PivotTable39").PivotFields("Storage Loc."). _
        CurrentPage = "(All)"
    With ActiveSheet.PivotTables("PivotTable39").PivotFields("Storage Loc.")
        .PivotItems("3001").Visible = False
        .PivotItems("4001").Visible = False
    End With
    ActiveSheet.PivotTables("PivotTable39").PivotFields("Storage Loc."). _
        EnableMultiplePageItems = True
    ActiveSheet.PivotTables("PivotTable39").PivotFields("Group").ClearAllFilters
    ActiveSheet.PivotTables("PivotTable39").PivotFields("Group").CurrentPage = _
        "BTV-EOL"
    ActiveSheet.PivotTables("PivotTable39").PivotFields("Plant").CurrentPage = _
        "(All)"
    With ActiveSheet.PivotTables("PivotTable39").PivotFields("Plant")
        .PivotItems("V113").Visible = False
        .PivotItems("(blank)").Visible = False
    End With
    ActiveSheet.PivotTables("PivotTable39").PivotFields("Plant"). _
        EnableMultiplePageItems = True
' Set rows to old material numbers
    With ActiveSheet.PivotTables("PivotTable39").PivotFields("Old matl number")
        .Orientation = xlRowField
        .Position = 1
' Set columns of pivot table and filter
    End With
    ActiveSheet.PivotTables("PivotTable39").AddDataField ActiveSheet.PivotTables( _
        "PivotTable39").PivotFields("0-30 days"), "Sum of 0-30 days", xlSum
    ActiveSheet.PivotTables("PivotTable39").AddDataField ActiveSheet.PivotTables( _
        "PivotTable39").PivotFields(">30-60 days"), "Sum of >30-60 days", xlSum
    ActiveSheet.PivotTables("PivotTable39").AddDataField ActiveSheet.PivotTables( _
        "PivotTable39").PivotFields(">60-90 days"), "Sum of >60-90 days", xlSum
    ActiveSheet.PivotTables("PivotTable39").AddDataField ActiveSheet.PivotTables( _
        "PivotTable39").PivotFields(">90-120 days"), "Sum of >90-120 days", xlSum
    ActiveSheet.PivotTables("PivotTable39").AddDataField ActiveSheet.PivotTables( _
        "PivotTable39").PivotFields(">120-150 days"), "Sum of >120-150 days", xlSum
    ActiveSheet.PivotTables("PivotTable39").AddDataField ActiveSheet.PivotTables( _
        "PivotTable39").PivotFields(">150 days"), "Sum of >150 days", xlSum
    ActiveSheet.PivotTables("PivotTable39").AddDataField ActiveSheet.PivotTables( _
        "PivotTable39").PivotFields("Total"), "Sum of Total", xlSum
    ActiveSheet.PivotTables("PivotTable39").AddDataField ActiveSheet.PivotTables( _
        "PivotTable39").PivotFields("ODM"), "Sum of ODM", xlSum
    ActiveSheet.PivotTables("PivotTable39").PivotFields("Old matl number"). _
        PivotFilters.Add2 Type:=xlValueDoesNotEqual, DataField:=ActiveSheet. _
        PivotTables("PivotTable39").PivotFields("Sum of Total"), Value1:=0
' Prepare new "EOL Inventory" table on N35 (copy column headers from most recent pivot table)
    Range("N35").Select
    ActiveCell.FormulaR1C1 = "EOL Inventory"
    Range("A36:H36").Select
    Selection.Copy
    Range("N36:U36").Select
    ActiveSheet.Paste
' Add column headers for "storage cost" and "monthly cost"
    Range("V36").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "Storage cost"
    Range("W36").Select
    ActiveCell.FormulaR1C1 = "Monthly cost"
' Format header cell and name "VIZIO cost"
    Range("X36").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternThemeColor = xlThemeColorAccent1
        .ThemeColor = xlThemeColorAccent1
        .TintAndShade = 0.799981688894314
        .PatternTintAndShade = 0.799981688894314
    End With
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    Selection.Borders(xlEdgeLeft).LineStyle = xlNone
    Selection.Borders(xlEdgeTop).LineStyle = xlNone
    Selection.Borders(xlEdgeBottom).LineStyle = xlNone
    Selection.Borders(xlEdgeRight).LineStyle = xlNone
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    ActiveCell.FormulaR1C1 = "VIZIO cost"
' Copy last row of most recent pivot table to "EOL Inventory" (material numbers change week to week so fixed last row for this table)
    Dim startCell4 As Range, lastRow4 As Long
    Set startCell4 = Range("A37")
    lastRow4 = Cells(Rows.Count, startCell4.Column).End(xlUp).Row
    Range(Cells(lastRow4, 1), Cells(lastRow4, 8)).Select
    Selection.Copy
    Range("N49:U49").Select
    ActiveSheet.Paste
' Fill in cells with VLOOKUP pointing to most recent pivot table and sum formula
    Range("O37").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "=IFERROR(VLOOKUP(RC14,C1:C8,2,FALSE),0)"
    Range("O37").Select
    Selection.AutoFill Destination:=Range("O37:T37"), Type:=xlFillDefault
    Range("O37:T37").Select
    Range("P37").Select
    ActiveCell.FormulaR1C1 = "=IFERROR(VLOOKUP(RC14,C1:C8,3,FALSE),0)"
    Range("Q37").Select
    ActiveCell.FormulaR1C1 = "=IFERROR(VLOOKUP(RC14,C1:C8,4,FALSE),0)"
    Range("R37").Select
    ActiveCell.FormulaR1C1 = "=IFERROR(VLOOKUP(RC14,C1:C8,5,FALSE),0)"
    Range("S37").Select
    ActiveCell.FormulaR1C1 = "=IFERROR(VLOOKUP(RC14,C1:C8,6,FALSE),0)"
    Range("T37").Select
    ActiveCell.FormulaR1C1 = "=IFERROR(VLOOKUP(RC14,C1:C8,FALSE),0)"
    Selection.AutoFill Destination:=Range("T37:U37"), Type:=xlFillDefault
    Range("T37:U37").Select
    Range("U37").Select
    ActiveCell.FormulaR1C1 = "=SUM(RC[-6]:RC[-1])"
' Auto fill formulas to each column
    Range("O37").Select
    ActiveCell.FormulaR1C1 = "=IFERROR(VLOOKUP(RC[-1],C1:C8,2,FALSE),0)"
    Range("O37").Select
    Selection.AutoFill Destination:=Range("O37:O48"), Type:=xlFillDefault
    Range("O37:O48").Select
    Range("P37").Select
    ActiveCell.FormulaR1C1 = "=IFERROR(VLOOKUP(RC[-2],C1:C8,3,FALSE),0)"
    Range("P37").Select
    Selection.AutoFill Destination:=Range("P37:P48"), Type:=xlFillDefault
    Range("P37:P48").Select
    Range("Q37").Select
    ActiveCell.FormulaR1C1 = "=IFERROR(VLOOKUP(RC[-3],C1:C8,4,FALSE),0)"
    Range("Q37").Select
    Selection.AutoFill Destination:=Range("Q37:Q48"), Type:=xlFillDefault
    Range("Q37:Q48").Select
    Range("R37").Select
    ActiveCell.FormulaR1C1 = "=IFERROR(VLOOKUP(RC[-4],C1:C8,5,FALSE),0)"
    Range("R37").Select
    Selection.AutoFill Destination:=Range("R37:R48"), Type:=xlFillDefault
    Range("R37:R48").Select
    Range("S37").Select
    ActiveCell.FormulaR1C1 = "=IFERROR(VLOOKUP(RC[-5],C1:C8,6,FALSE),0)"
    Range("S37").Select
    Selection.AutoFill Destination:=Range("S37:S48"), Type:=xlFillDefault
    Range("S37:S48").Select
    Range("T37").Select
    ActiveCell.FormulaR1C1 = "=IFERROR(VLOOKUP(RC[-6],C1:C8,7,FALSE),0)"
    Range("T37").Select
    Selection.AutoFill Destination:=Range("T37:T48"), Type:=xlFillDefault
    Range("T37:T48").Select
    Range("U37").Select
    Selection.AutoFill Destination:=Range("U37:U48"), Type:=xlFillDefault
    Range("U37:U48").Select
' Add total formula and auto fill to cells in row
    Range("O49").Select
    ActiveCell.FormulaR1C1 = "=SUM(R[-12]C:R[-1]C)"
    Range("O49").Select
    Selection.AutoFill Destination:=Range("O49:U49"), Type:=xlFillDefault
    Range("O49:U49").Select
' Add formula for "VIZIO cost" column and auto fill cells
    Range("X37").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "=RC[-3]*RC[-1]"
    Range("X37").Select
    Selection.AutoFill Destination:=Range("X37:X48"), Type:=xlFillDefault
    Range("X37:X48").Select
' Format "VIZIO cost" cells (change cell color, vold, and change to curreny)
    Range("X49").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternThemeColor = xlThemeColorAccent1
        .ThemeColor = xlThemeColorAccent1
        .TintAndShade = 0.799981688894314
        .PatternTintAndShade = 0.799981688894314
    End With
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    Selection.Borders(xlEdgeLeft).LineStyle = xlNone
    Selection.Borders(xlEdgeTop).LineStyle = xlNone
    Selection.Borders(xlEdgeBottom).LineStyle = xlNone
    Selection.Borders(xlEdgeRight).LineStyle = xlNone
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    ActiveCell.FormulaR1C1 = "=SUM(R[-12]C:R[-1]C)"
    Range("X36").Select
    Selection.Font.Bold = True
    Range("X49").Select
    Selection.Font.Bold = True
    Range(Cells(37, 24), Cells(49, 24)).Select
    Selection.Style = "Currency"
    
End Sub


Sub BTVRefurb()
' Create pivot table and another table in "Summary" that includes information on refurbished models

' Create new pivot table in "Summary" starting on row 68
    Sheets("Data").Select
    ActiveWorkbook.Worksheets("Divvy").PivotTables("PivotTable34").PivotCache. _
        CreatePivotTable TableDestination:="Summary!R68C1", TableName:= _
        "PivotTable40", DefaultVersion:=6
    Sheets("Summary").Select
    Cells(68, 1).Select
    With ActiveSheet.PivotTables("PivotTable40")
        .ColumnGrand = True
        .HasAutoFormat = True
        .DisplayErrorString = False
        .DisplayNullString = True
        .EnableDrilldown = True
        .ErrorString = ""
        .MergeLabels = False
        .NullString = ""
        .PageFieldOrder = 2
        .PageFieldWrapCount = 0
        .PreserveFormatting = True
        .RowGrand = True
        .SaveData = True
        .PrintTitles = False
        .RepeatItemsOnEachPrintedPage = True
        .TotalsAnnotation = False
        .CompactRowIndent = 1
        .InGridDropZones = False
        .DisplayFieldCaptions = True
        .DisplayMemberPropertyTooltips = False
        .DisplayContextTooltips = True
        .ShowDrillIndicators = True
        .PrintDrillIndicators = False
        .AllowMultipleFilters = False
        .SortUsingCustomLists = True
        .FieldListSortAscending = False
        .ShowValuesRow = False
        .CalculatedMembersInFilters = False
        .RowAxisLayout xlCompactRow
    End With
    With ActiveSheet.PivotTables("PivotTable40").PivotCache
        .RefreshOnFileOpen = False
        .MissingItemsLimit = xlMissingItemsDefault
' Create filters for group (BTV-refurb), plant (not V113 or blank), and storage location (not 3001 or 4001)
    End With
    ActiveSheet.PivotTables("PivotTable40").RepeatAllLabels xlRepeatLabels
    With ActiveSheet.PivotTables("PivotTable40").PivotFields("Group")
        .Orientation = xlPageField
        .Position = 1
    End With
    ActiveSheet.PivotTables("PivotTable40").PivotFields("Group").ClearAllFilters
    ActiveSheet.PivotTables("PivotTable40").PivotFields("Group").CurrentPage = _
        "BTV-refurb"
    With ActiveSheet.PivotTables("PivotTable40").PivotFields("Plant")
        .Orientation = xlPageField
        .Position = 1
    End With
    With ActiveSheet.PivotTables("PivotTable40").PivotFields("Storage Loc.")
        .Orientation = xlPageField
        .Position = 1
    End With
    ActiveSheet.PivotTables("PivotTable40").PivotFields("Storage Loc."). _
        CurrentPage = "(All)"
    With ActiveSheet.PivotTables("PivotTable40").PivotFields("Storage Loc.")
        .PivotItems("3001").Visible = False
        .PivotItems("4001").Visible = False
    End With
    ActiveSheet.PivotTables("PivotTable40").PivotFields("Storage Loc."). _
        EnableMultiplePageItems = True
    ActiveSheet.PivotTables("PivotTable40").PivotFields("Plant").CurrentPage = _
        "(All)"
    With ActiveSheet.PivotTables("PivotTable40").PivotFields("Plant")
        .PivotItems("V113").Visible = False
        .PivotItems("(blank)").Visible = False
    End With
    ActiveSheet.PivotTables("PivotTable40").PivotFields("Plant"). _
        EnableMultiplePageItems = True
' Set rows to old material numbers
    With ActiveSheet.PivotTables("PivotTable40").PivotFields("Old matl number")
        .Orientation = xlRowField
        .Position = 1
' Set columns of pivot table and filter
    End With
    ActiveSheet.PivotTables("PivotTable40").AddDataField ActiveSheet.PivotTables( _
        "PivotTable40").PivotFields("0-30 days"), "Sum of 0-30 days", xlSum
    ActiveSheet.PivotTables("PivotTable40").AddDataField ActiveSheet.PivotTables( _
        "PivotTable40").PivotFields(">30-60 days"), "Sum of >30-60 days", xlSum
    ActiveSheet.PivotTables("PivotTable40").AddDataField ActiveSheet.PivotTables( _
        "PivotTable40").PivotFields(">60-90 days"), "Sum of >60-90 days", xlSum
    ActiveSheet.PivotTables("PivotTable40").AddDataField ActiveSheet.PivotTables( _
        "PivotTable40").PivotFields(">90-120 days"), "Sum of >90-120 days", xlSum
    ActiveSheet.PivotTables("PivotTable40").AddDataField ActiveSheet.PivotTables( _
        "PivotTable40").PivotFields(">120-150 days"), "Sum of >120-150 days", xlSum
    ActiveSheet.PivotTables("PivotTable40").AddDataField ActiveSheet.PivotTables( _
        "PivotTable40").PivotFields(">150 days"), "Sum of >150 days", xlSum
    ActiveSheet.PivotTables("PivotTable40").AddDataField ActiveSheet.PivotTables( _
        "PivotTable40").PivotFields("Total"), "Sum of Total", xlSum
    ActiveSheet.PivotTables("PivotTable40").PivotFields("Old matl number"). _
        PivotFilters.Add2 Type:=xlValueDoesNotEqual, DataField:=ActiveSheet. _
        PivotTables("PivotTable40").PivotFields("Sum of Total"), Value1:=0
' Prepare new "BTV Refurb" table on N57 (copy column headers from most recent pivot table)
    Range("N57").Select
    ActiveCell.FormulaR1C1 = "BTV Refurb"
    Range("A68:H68").Select
    Selection.Copy
    Range("N58:U58").Select
    ActiveSheet.Paste
' Add column headers for "storage cost" and "monthly cost"
    Range("V58").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "Storage cost"
    Range("W58").Select
    ActiveCell.FormulaR1C1 = "Monthly cost"
' Format header cell and name "VIZIO cost"
    Range("X58").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternThemeColor = xlThemeColorAccent1
        .ThemeColor = xlThemeColorAccent1
        .TintAndShade = 0.799981688894314
        .PatternTintAndShade = 0.799981688894314
    End With
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    Selection.Borders(xlEdgeLeft).LineStyle = xlNone
    Selection.Borders(xlEdgeTop).LineStyle = xlNone
    Selection.Borders(xlEdgeBottom).LineStyle = xlNone
    Selection.Borders(xlEdgeRight).LineStyle = xlNone
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    ActiveCell.FormulaR1C1 = "VIZIO cost"
' Copy last row of most recent pivot table to "BTV Refurb" (material numbers change week to week so fixed last row for this table)
    Dim startCell5 As Range, lastRow5 As Long
    Set startCell5 = Range("A69")
    lastRow5 = Cells(Rows.Count, startCell5.Column).End(xlUp).Row
    Range(Cells(lastRow5, 1), Cells(lastRow5, 8)).Select
    Selection.Copy
    Range("N77:U77").Select
    ActiveSheet.Paste
    Range("O59").Select
    Application.CutCopyMode = False
' Fill in cells with VLOOKUP pointing to most recent pivot table and sum formula
    ActiveCell.FormulaR1C1 = "=IFERROR(VLOOKUP(RC14,C1:C8,2,FALSE),0)"
    Range("O59").Select
    Selection.AutoFill Destination:=Range("O59:T59"), Type:=xlFillDefault
    Range("O59:T59").Select
    Range("P59").Select
    ActiveCell.FormulaR1C1 = "=IFERROR(VLOOKUP(RC14,C1:C8,3,FALSE),0)"
    Range("Q59").Select
    ActiveCell.FormulaR1C1 = "=IFERROR(VLOOKUP(RC14,C1:C8,4,FALSE),0)"
    Range("R59").Select
    ActiveCell.FormulaR1C1 = "=IFERROR(VLOOKUP(RC14,C1:C8,5,FALSE),0)"
    Range("S59").Select
    ActiveCell.FormulaR1C1 = "=IFERROR(VLOOKUP(RC14,C1:C8,6,FALSE),0)"
    Range("T59").Select
    ActiveCell.FormulaR1C1 = "=IFERROR(VLOOKUP(RC14,C1:C8,7,FALSE),0)"
    Selection.AutoFill Destination:=Range("T59:U59"), Type:=xlFillDefault
    Range("T59:U59").Select
    Range("U59").Select
    ActiveCell.FormulaR1C1 = "=SUM(RC[-6]:RC[-1])"
' Auto fill formulas to each column
    Range("O59").Select
    ActiveCell.FormulaR1C1 = "=IFERROR(VLOOKUP(RC[-1],C1:C8,2,FALSE),0)"
    Range("O59").Select
    Selection.AutoFill Destination:=Range("O59:O76"), Type:=xlFillDefault
    Range("O59:O76").Select
    Range("P59").Select
    ActiveCell.FormulaR1C1 = "=IFERROR(VLOOKUP(RC[-2],C1:C8,3,FALSE),0)"
    Range("P59").Select
    Selection.AutoFill Destination:=Range("P59:P76"), Type:=xlFillDefault
    Range("P59:P76").Select
    Range("Q59").Select
    ActiveCell.FormulaR1C1 = "=IFERROR(VLOOKUP(RC[-3],C1:C8,4,FALSE),0)"
    Range("Q59").Select
    Selection.AutoFill Destination:=Range("Q59:Q76"), Type:=xlFillDefault
    Range("Q59:Q76").Select
    Range("R59").Select
    ActiveCell.FormulaR1C1 = "=IFERROR(VLOOKUP(RC[-4],C1:C8,5,FALSE),0)"
    Range("R59").Select
    Selection.AutoFill Destination:=Range("R59:R76"), Type:=xlFillDefault
    Range("R59:R76").Select
    Range("S59").Select
    ActiveCell.FormulaR1C1 = "=IFERROR(VLOOKUP(RC[-5],C1:C8,6,FALSE),0)"
    Range("S59").Select
    Selection.AutoFill Destination:=Range("S59:S76"), Type:=xlFillDefault
    Range("S59:S76").Select
    Range("T59").Select
    ActiveCell.FormulaR1C1 = "=IFERROR(VLOOKUP(RC[-6],C1:C8,7,FALSE),0)"
    Range("T59").Select
    Selection.AutoFill Destination:=Range("T59:T76"), Type:=xlFillDefault
    Range("T59:T76").Select
    Range("U59").Select
    Selection.AutoFill Destination:=Range("U59:U76"), Type:=xlFillDefault
    Range("U59:U76").Select
    Range("O77").Select
    ActiveCell.FormulaR1C1 = "=SUM(R[-18]C:R[-1]C)"
    Range("O77").Select
    Selection.AutoFill Destination:=Range("O77:U77"), Type:=xlFillDefault
    Range("O77:U77").Select
' Add formula for "VIZIO cost" and autofill to cells in column
    Range("X59").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "=RC[-3]*RC[-1]"
    Range("X59").Select
    Selection.AutoFill Destination:=Range("X59:X76"), Type:=xlFillDefault
    Range("X59:X76").Select
' Add total formula for the column and format cells
    Range("X77").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternThemeColor = xlThemeColorAccent1
        .ThemeColor = xlThemeColorAccent1
        .TintAndShade = 0.799981688894314
        .PatternTintAndShade = 0.799981688894314
    End With
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    Selection.Borders(xlEdgeLeft).LineStyle = xlNone
    Selection.Borders(xlEdgeTop).LineStyle = xlNone
    Selection.Borders(xlEdgeBottom).LineStyle = xlNone
    Selection.Borders(xlEdgeRight).LineStyle = xlNone
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    ActiveCell.FormulaR1C1 = "=SUM(R[-18]C:R[-1]C)"
    Range("X58").Select
    Selection.Font.Bold = True
    Range("X77").Select
    Selection.Font.Bold = True
    Range(Cells(59, 24), Cells(77, 24)).Select
    Selection.Style = "Currency"
    
End Sub

Sub dataToSummary()
' Create pivot table in "Summary" that includes information on ongoing models that aren't in certain storage locations

' Add new pivot table to "Summary" starting on column 32
    Cells.Select
    ActiveWorkbook.Worksheets("Divvy").PivotTables("PivotTable34").PivotCache. _
        CreatePivotTable TableDestination:="Summary!R1C32", TableName:= _
        "PivotTable41", DefaultVersion:=6
    Sheets("Summary").Select
    Cells(1, 32).Select
    With ActiveSheet.PivotTables("PivotTable41")
        .ColumnGrand = True
        .HasAutoFormat = True
        .DisplayErrorString = False
        .DisplayNullString = True
        .EnableDrilldown = True
        .ErrorString = ""
        .MergeLabels = False
        .NullString = ""
        .PageFieldOrder = 2
        .PageFieldWrapCount = 0
        .PreserveFormatting = True
        .RowGrand = True
        .SaveData = True
        .PrintTitles = False
        .RepeatItemsOnEachPrintedPage = True
        .TotalsAnnotation = False
        .CompactRowIndent = 1
        .InGridDropZones = False
        .DisplayFieldCaptions = True
        .DisplayMemberPropertyTooltips = False
        .DisplayContextTooltips = True
        .ShowDrillIndicators = True
        .PrintDrillIndicators = False
        .AllowMultipleFilters = False
        .SortUsingCustomLists = True
        .FieldListSortAscending = False
        .ShowValuesRow = False
        .CalculatedMembersInFilters = False
        .RowAxisLayout xlCompactRow
    End With
    With ActiveSheet.PivotTables("PivotTable41").PivotCache
        .RefreshOnFileOpen = False
        .MissingItemsLimit = xlMissingItemsDefault
' Create filters for group (BTV), plant (not V113 or blank), and storage location (not 1101, 2001, or 2002)
    End With
    ActiveSheet.PivotTables("PivotTable41").RepeatAllLabels xlRepeatLabels
    With ActiveSheet.PivotTables("PivotTable41").PivotFields("Group")
        .Orientation = xlPageField
        .Position = 1
    End With
    ActiveSheet.PivotTables("PivotTable41").PivotFields("Group").ClearAllFilters
    ActiveSheet.PivotTables("PivotTable41").PivotFields("Group").CurrentPage = _
        "BTV"
    With ActiveSheet.PivotTables("PivotTable41").PivotFields("Plant")
        .Orientation = xlPageField
        .Position = 1
    End With
    With ActiveSheet.PivotTables("PivotTable41").PivotFields("Storage Loc.")
        .Orientation = xlPageField
        .Position = 1
    End With
    ActiveSheet.PivotTables("PivotTable41").PivotFields("Plant").CurrentPage = _
        "(All)"
    With ActiveSheet.PivotTables("PivotTable41").PivotFields("Plant")
        .PivotItems("V113").Visible = False
        .PivotItems("(blank)").Visible = False
    End With
    ActiveSheet.PivotTables("PivotTable41").PivotFields("Plant"). _
        EnableMultiplePageItems = True
    ActiveSheet.PivotTables("PivotTable41").PivotFields("Storage Loc."). _
        CurrentPage = "(All)"
    With ActiveSheet.PivotTables("PivotTable41").PivotFields("Storage Loc.")
        .PivotItems("1101").Visible = False
        .PivotItems("2001").Visible = False
        .PivotItems("2002").Visible = False
' Set rows to old material numbers
    End With
    ActiveSheet.PivotTables("PivotTable41").PivotFields("Storage Loc."). _
        EnableMultiplePageItems = True
    With ActiveSheet.PivotTables("PivotTable41").PivotFields("Old matl number")
        .Orientation = xlRowField
        .Position = 1
' Set columns of pivot table and filter
    End With
    ActiveSheet.PivotTables("PivotTable41").AddDataField ActiveSheet.PivotTables( _
        "PivotTable41").PivotFields("0-30 days"), "Sum of 0-30 days", xlSum
    ActiveSheet.PivotTables("PivotTable41").AddDataField ActiveSheet.PivotTables( _
        "PivotTable41").PivotFields(">30-60 days"), "Sum of >30-60 days", xlSum
    ActiveSheet.PivotTables("PivotTable41").AddDataField ActiveSheet.PivotTables( _
        "PivotTable41").PivotFields(">60-90 days"), "Sum of >60-90 days", xlSum
    ActiveSheet.PivotTables("PivotTable41").AddDataField ActiveSheet.PivotTables( _
        "PivotTable41").PivotFields(">90-120 days"), "Sum of >90-120 days", xlSum
    ActiveSheet.PivotTables("PivotTable41").AddDataField ActiveSheet.PivotTables( _
        "PivotTable41").PivotFields(">120-150 days"), "Sum of >120-150 days", xlSum
    ActiveSheet.PivotTables("PivotTable41").AddDataField ActiveSheet.PivotTables( _
        "PivotTable41").PivotFields(">150 days"), "Sum of >150 days", xlSum
    ActiveSheet.PivotTables("PivotTable41").AddDataField ActiveSheet.PivotTables( _
        "PivotTable41").PivotFields("Total"), "Sum of Total", xlSum
    ActiveSheet.PivotTables("PivotTable41").PivotFields("Old matl number"). _
        PivotFilters.Add2 Type:=xlValueDoesNotEqual, DataField:=ActiveSheet. _
        PivotTables("PivotTable41").PivotFields("Sum of Total"), Value1:=0
    Range("A2").Select
    
End Sub
