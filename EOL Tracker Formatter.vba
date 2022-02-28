Sub Temp()
'
' Temp Macro
'

'
    ActiveSheet.PivotTables("PivotTable38").PivotFields("Storage Loc."). _
        CurrentPage = "(All)"
    With ActiveSheet.PivotTables("PivotTable38").PivotFields("Storage Loc.")
        .PivotItems("3001").Visible = False
        .PivotItems("4001").Visible = False
    End With
    ActiveSheet.PivotTables("PivotTable38").PivotFields("Storage Loc."). _
        EnableMultiplePageItems = True
End Sub
Sub Temp1()
'
' Temp1 Macro
'

'
End Sub
Sub temp4()
'
' temp4 Macro
'

'
    Range("J4").Select
    ActiveCell.FormulaR1C1 = "=RC[-2]-RC[-1]"
    Range("J4").Select
    Selection.AutoFill Destination:=Range("J4:J60"), Type:=xlFillDefault
    Range("J4:J60").Select
    ActiveWindow.SmallScroll Down:=6
End Sub
Sub temp5()
'
' temp5 Macro
'

'
    Range("J60").Select
    ActiveCell.FormulaR1C1 = "=SUM(R[-56]C:R[-1]C)"
    Range("J61").Select
End Sub








Sub Temp()
'
' Temp Macro
'

'
    ActiveSheet.PivotTables("PivotTable38").PivotFields("Storage Loc."). _
        CurrentPage = "(All)"
    With ActiveSheet.PivotTables("PivotTable38").PivotFields("Storage Loc.")
        .PivotItems("3001").Visible = False
        .PivotItems("4001").Visible = False
    End With
    ActiveSheet.PivotTables("PivotTable38").PivotFields("Storage Loc."). _
        EnableMultiplePageItems = True
End Sub
Sub Temp1()
'
' Temp1 Macro
'

'
End Sub
Sub temp4()
'
' temp4 Macro
'

'
    Range("J4").Select
    ActiveCell.FormulaR1C1 = "=RC[-2]-RC[-1]"
    Range("J4").Select
    Selection.AutoFill Destination:=Range("J4:J60"), Type:=xlFillDefault
    Range("J4:J60").Select
    ActiveWindow.SmallScroll Down:=6
End Sub
Sub temp5()
'
' temp5 Macro
'

'
    Range("J60").Select
    ActiveCell.FormulaR1C1 = "=SUM(R[-56]C:R[-1]C)"
    Range("J61").Select
End Sub






Sub AlltoData()
'
' AlltoData Macro
'
' Keyboard Shortcut: Ctrl+a
'
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
    Range("O1").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "Total"
    Range("P1").Select
    ActiveCell.FormulaR1C1 = "Data"
    Range("Q1").Select
    ActiveCell.FormulaR1C1 = "Group"
    Range("R1").Select
    ActiveCell.FormulaR1C1 = "ODM"
    Range("O2").Select
    ActiveCell.FormulaR1C1 = "=SUM(RC[-7],RC[-6],RC[-5],RC[-4],RC[-3],RC[-2])"
    Range("P2").Select
    ActiveCell.FormulaR1C1 = "=RC[-15]&RC[-11]&RC[-10]"
    Range("Q2").Select
    ActiveCell.FormulaR1C1 = _
        "=VLOOKUP(RC[-1],'[copy-2019-07-17-Inventory-Aging-Report - Audio.xlsx]Data'!C16:C17,2,FALSE)"
    Range("R2").Select
    ActiveCell.FormulaR1C1 = "=IF(OR(RC[-12]=""2101"",RC[-12]=""2119"",RC[-12]=""2123"",RC[-12]=""2128"",AND(RC[-13]=""V113"",RC[-12]=""4001"")),RC[-3],0)"
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
End Sub
Sub DatatoDivvy1()
'
' DatatoDivvy1 Macro
'

'
    Sheets.Add.Name = "Divvy"
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
    With ActiveSheet.PivotTables("PivotTable28").PivotFields("Old matl number")
        .Orientation = xlRowField
        .Position = 1
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
    ActiveSheet.PivotTables("PivotTable28").AddDataField ActiveSheet.PivotTables( _
        "PivotTable28").PivotFields("ODM"), "Sum of ODM", xlSum
    Columns("I:I").Select
    Selection.Copy
    Columns("M:M").Select
    ActiveSheet.Paste
    Range("M3").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "ODM"
    Range("H11").Select
    ActiveSheet.PivotTables("PivotTable28").PivotFields("Sum of ODM").Orientation _
        = xlHidden
    Range("I3").Select
    ActiveCell.FormulaR1C1 = "Orders"
    Range("J3").Select
    ActiveCell.FormulaR1C1 = "Balance"
    Range("L3").Select
    ActiveCell.FormulaR1C1 = "VIZIO"
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
    Range("L4").Select
    ActiveCell.FormulaR1C1 = "=RC[-4]-RC[1]"
    Range("L4").Select
    Dim startCell As Range, lastRow As Long
    Set startCell = Range("M4")
    lastRow = Cells(Rows.Count, startCell.Column).End(xlUp).Row
    Selection.AutoFill Destination:=Range(Cells(4, 12), Cells(lastRow, 12))
End Sub







Sub DatatoDivvy2()
'
' DatatoDivvy2 Macro
'

'
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
    Range("Q3").Select
    With ActiveSheet.PivotTables("PivotTable34").PivotFields("Old matl number")
        .Orientation = xlRowField
        .Position = 1
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
End Sub
Sub DatatoDivvy3()
'
' DatatoDivvy3 Macro
'

'
    Cells.Select
    ActiveWorkbook.Worksheets("Divvy").PivotTables("PivotTable34").PivotCache. _
        CreatePivotTable TableDestination:="Divvy!R20C16", TableName:= _
        "PivotTable35", DefaultVersion:=6
    Sheets("Divvy").Select
    Cells(20, 16).Select
    With ActiveSheet.PivotTables("PivotTable35")
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
    With ActiveSheet.PivotTables("PivotTable35").PivotCache
        .RefreshOnFileOpen = False
        .MissingItemsLimit = xlMissingItemsDefault
    End With
    ActiveSheet.PivotTables("PivotTable35").RepeatAllLabels xlRepeatLabels
    ActiveWorkbook.ShowPivotTableFieldList = True
    With ActiveSheet.PivotTables("PivotTable35").PivotFields("Group")
        .Orientation = xlPageField
        .Position = 1
    End With
    With ActiveSheet.PivotTables("PivotTable35").PivotFields("Plant")
        .Orientation = xlPageField
        .Position = 1
    End With
    ActiveSheet.PivotTables("PivotTable35").PivotFields("Plant").ClearAllFilters
    ActiveSheet.PivotTables("PivotTable35").PivotFields("Plant").CurrentPage = _
        "V113"
    With ActiveSheet.PivotTables("PivotTable35").PivotFields("Storage Loc.")
        .Orientation = xlPageField
        .Position = 1
    End With
    ActiveSheet.PivotTables("PivotTable35").PivotFields("Storage Loc."). _
        ClearAllFilters
    ActiveSheet.PivotTables("PivotTable35").PivotFields("Storage Loc."). _
        CurrentPage = "2119"
    With ActiveSheet.PivotTables("PivotTable35").PivotFields("Old matl number")
        .Orientation = xlRowField
        .Position = 1
    End With
    ActiveSheet.PivotTables("PivotTable35").AddDataField ActiveSheet.PivotTables( _
        "PivotTable35").PivotFields("0-30 days"), "Sum of 0-30 days", xlSum
    ActiveSheet.PivotTables("PivotTable35").AddDataField ActiveSheet.PivotTables( _
        "PivotTable35").PivotFields(">30-60 days"), "Sum of >30-60 days", xlSum
    ActiveSheet.PivotTables("PivotTable35").AddDataField ActiveSheet.PivotTables( _
        "PivotTable35").PivotFields(">60-90 days"), "Sum of >60-90 days", xlSum
    ActiveSheet.PivotTables("PivotTable35").AddDataField ActiveSheet.PivotTables( _
        "PivotTable35").PivotFields(">90-120 days"), "Sum of >90-120 days", xlSum
    ActiveSheet.PivotTables("PivotTable35").AddDataField ActiveSheet.PivotTables( _
        "PivotTable35").PivotFields(">120-150 days"), "Sum of >120-150 days", xlSum
    ActiveSheet.PivotTables("PivotTable35").AddDataField ActiveSheet.PivotTables( _
        "PivotTable35").PivotFields(">150 days"), "Sum of >150 days", xlSum
    ActiveSheet.PivotTables("PivotTable35").AddDataField ActiveSheet.PivotTables( _
        "PivotTable35").PivotFields("Total"), "Sum of Total", xlSum
    ActiveSheet.PivotTables("PivotTable35").PivotFields("Old matl number"). _
        PivotFilters.Add2 Type:=xlValueDoesNotEqual, DataField:=ActiveSheet. _
        PivotTables("PivotTable35").PivotFields("Sum of Total"), Value1:=0
End Sub
Sub DatatoDivvy4()
'
' DatatoDivvy4 Macro
'

'
    Cells.Select
    ActiveWorkbook.Worksheets("Divvy").PivotTables("PivotTable35").PivotCache. _
        CreatePivotTable TableDestination:="Divvy!R30C16", TableName:= _
        "PivotTable36", DefaultVersion:=6
    Sheets("Divvy").Select
    Cells(30, 16).Select
    With ActiveSheet.PivotTables("PivotTable36")
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
    With ActiveSheet.PivotTables("PivotTable36").PivotCache
        .RefreshOnFileOpen = False
        .MissingItemsLimit = xlMissingItemsDefault
    End With
    ActiveSheet.PivotTables("PivotTable36").RepeatAllLabels xlRepeatLabels
    ActiveWorkbook.ShowPivotTableFieldList = True
    With ActiveSheet.PivotTables("PivotTable36").PivotFields("Group")
        .Orientation = xlPageField
        .Position = 1
    End With
    With ActiveSheet.PivotTables("PivotTable36").PivotFields("Plant")
        .Orientation = xlPageField
        .Position = 1
    End With
    ActiveSheet.PivotTables("PivotTable36").PivotFields("Plant").ClearAllFilters
    ActiveSheet.PivotTables("PivotTable36").PivotFields("Plant").CurrentPage = _
        "V113"
    With ActiveSheet.PivotTables("PivotTable36").PivotFields("Storage Loc.")
        .Orientation = xlPageField
        .Position = 1
    End With
    ActiveSheet.PivotTables("PivotTable36").PivotFields("Storage Loc."). _
        ClearAllFilters
    ActiveSheet.PivotTables("PivotTable36").PivotFields("Storage Loc."). _
        CurrentPage = "2128"
    With ActiveSheet.PivotTables("PivotTable36").PivotFields("Old matl number")
        .Orientation = xlRowField
        .Position = 1
    End With
    ActiveSheet.PivotTables("PivotTable36").AddDataField ActiveSheet.PivotTables( _
        "PivotTable36").PivotFields("0-30 days"), "Sum of 0-30 days", xlSum
    ActiveSheet.PivotTables("PivotTable36").AddDataField ActiveSheet.PivotTables( _
        "PivotTable36").PivotFields(">30-60 days"), "Sum of >30-60 days", xlSum
    ActiveSheet.PivotTables("PivotTable36").AddDataField ActiveSheet.PivotTables( _
        "PivotTable36").PivotFields(">60-90 days"), "Sum of >60-90 days", xlSum
    ActiveSheet.PivotTables("PivotTable36").AddDataField ActiveSheet.PivotTables( _
        "PivotTable36").PivotFields(">90-120 days"), "Sum of >90-120 days", xlSum
    ActiveSheet.PivotTables("PivotTable36").AddDataField ActiveSheet.PivotTables( _
        "PivotTable36").PivotFields(">120-150 days"), "Sum of >120-150 days", xlSum
    ActiveSheet.PivotTables("PivotTable36").AddDataField ActiveSheet.PivotTables( _
        "PivotTable36").PivotFields(">150 days"), "Sum of >150 days", xlSum
    ActiveSheet.PivotTables("PivotTable36").AddDataField ActiveSheet.PivotTables( _
        "PivotTable36").PivotFields("Total"), "Sum of Total", xlSum
    ActiveSheet.PivotTables("PivotTable36").PivotFields("Old matl number"). _
        PivotFilters.Add2 Type:=xlValueDoesNotEqual, DataField:=ActiveSheet. _
        PivotTables("PivotTable36").PivotFields("Sum of Total"), Value1:=0
End Sub

Sub DatatoDivvy5()
'
' DatatoDivvy5 Macro
'

'
    Cells.Select
    ActiveWorkbook.Worksheets("Divvy").PivotTables("PivotTable36").PivotCache. _
        CreatePivotTable TableDestination:="Divvy!R46C16", TableName:= _
        "PivotTable37", DefaultVersion:=6
    Sheets("Divvy").Select
    Cells(46, 16).Select
    With ActiveSheet.PivotTables("PivotTable37")
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
    With ActiveSheet.PivotTables("PivotTable37").PivotCache
        .RefreshOnFileOpen = False
        .MissingItemsLimit = xlMissingItemsDefault
    End With
    ActiveSheet.PivotTables("PivotTable37").RepeatAllLabels xlRepeatLabels
    With ActiveSheet.PivotTables("PivotTable37").PivotFields("Group")
        .Orientation = xlPageField
        .Position = 1
    End With
    With ActiveSheet.PivotTables("PivotTable37").PivotFields("Plant")
        .Orientation = xlPageField
        .Position = 1
    End With
    ActiveSheet.PivotTables("PivotTable37").PivotFields("Plant").ClearAllFilters
    ActiveSheet.PivotTables("PivotTable37").PivotFields("Plant").CurrentPage = _
        "V113"
    With ActiveSheet.PivotTables("PivotTable37").PivotFields("Storage Loc.")
        .Orientation = xlPageField
        .Position = 1
    End With
    ActiveSheet.PivotTables("PivotTable37").PivotFields("Storage Loc."). _
        ClearAllFilters
    ActiveSheet.PivotTables("PivotTable37").PivotFields("Storage Loc."). _
        CurrentPage = "2101"
    With ActiveSheet.PivotTables("PivotTable37").PivotFields("Old matl number")
        .Orientation = xlRowField
        .Position = 1
    End With
    ActiveSheet.PivotTables("PivotTable37").AddDataField ActiveSheet.PivotTables( _
        "PivotTable37").PivotFields("0-30 days"), "Sum of 0-30 days", xlSum
    ActiveSheet.PivotTables("PivotTable37").AddDataField ActiveSheet.PivotTables( _
        "PivotTable37").PivotFields(">30-60 days"), "Sum of >30-60 days", xlSum
    ActiveSheet.PivotTables("PivotTable37").AddDataField ActiveSheet.PivotTables( _
        "PivotTable37").PivotFields(">60-90 days"), "Sum of >60-90 days", xlSum
    ActiveSheet.PivotTables("PivotTable37").AddDataField ActiveSheet.PivotTables( _
        "PivotTable37").PivotFields(">90-120 days"), "Sum of >90-120 days", xlSum
    ActiveSheet.PivotTables("PivotTable37").AddDataField ActiveSheet.PivotTables( _
        "PivotTable37").PivotFields(">120-150 days"), "Sum of >120-150 days", xlSum
    ActiveSheet.PivotTables("PivotTable37").AddDataField ActiveSheet.PivotTables( _
        "PivotTable37").PivotFields(">150 days"), "Sum of >150 days", xlSum
    ActiveSheet.PivotTables("PivotTable37").AddDataField ActiveSheet.PivotTables( _
        "PivotTable37").PivotFields("Total"), "Sum of Total", xlSum
    ActiveSheet.PivotTables("PivotTable37").PivotFields("Old matl number"). _
        PivotFilters.Add2 Type:=xlValueDoesNotEqual, DataField:=ActiveSheet. _
        PivotTables("PivotTable37").PivotFields("Sum of Total"), Value1:=0
End Sub
Sub dataToSummary1()
'
' dataToSummary1 Macro
'

'
    Sheets.Add.Name = "Summary"
    Sheets("Data").Select
    Cells.Select
    ActiveWorkbook.Worksheets("Divvy").PivotTables("PivotTable37").PivotCache. _
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
    With ActiveSheet.PivotTables("PivotTable38").PivotFields("Old matl number")
        .Orientation = xlRowField
        .Position = 1
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
    Sheets("Data").Select
    ActiveWorkbook.Worksheets("Divvy").PivotTables("PivotTable37").PivotCache. _
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
    With ActiveSheet.PivotTables("PivotTable39").PivotFields("Old matl number")
        .Orientation = xlRowField
        .Position = 1
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
    Range("A94").Select
    Sheets("Data").Select
    ActiveWorkbook.Worksheets("Divvy").PivotTables("PivotTable37").PivotCache. _
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
    With ActiveSheet.PivotTables("PivotTable40").PivotFields("Old matl number")
        .Orientation = xlRowField
        .Position = 1
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
End Sub
Sub dataToSummary2()
'
' dataToSummary2 Macro
'

'
    Cells.Select
    ActiveWorkbook.Worksheets("Divvy").PivotTables("PivotTable37").PivotCache. _
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
    End With
    ActiveSheet.PivotTables("PivotTable41").PivotFields("Storage Loc."). _
        EnableMultiplePageItems = True
    With ActiveSheet.PivotTables("PivotTable41").PivotFields("Old matl number")
        .Orientation = xlRowField
        .Position = 1
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
    Range("AF34").Select
    Sheets("Data").Select
    ActiveWorkbook.Worksheets("Divvy").PivotTables("PivotTable37").PivotCache. _
        CreatePivotTable TableDestination:="Summary!R32C32", TableName:= _
        "PivotTable42", DefaultVersion:=6
    Sheets("Summary").Select
    Cells(32, 32).Select
    With ActiveSheet.PivotTables("PivotTable42")
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
    With ActiveSheet.PivotTables("PivotTable42").PivotCache
        .RefreshOnFileOpen = False
        .MissingItemsLimit = xlMissingItemsDefault
    End With
    ActiveSheet.PivotTables("PivotTable42").RepeatAllLabels xlRepeatLabels
    With ActiveSheet.PivotTables("PivotTable42").PivotFields("Group")
        .Orientation = xlPageField
        .Position = 1
    End With
    ActiveSheet.PivotTables("PivotTable42").PivotFields("Group").ClearAllFilters
    ActiveSheet.PivotTables("PivotTable42").PivotFields("Group").CurrentPage = _
        "BTV"
    With ActiveSheet.PivotTables("PivotTable42").PivotFields("Plant")
        .Orientation = xlPageField
        .Position = 1
    End With
    With ActiveSheet.PivotTables("PivotTable42").PivotFields("Storage Loc.")
        .Orientation = xlPageField
        .Position = 1
    End With
    ActiveSheet.PivotTables("PivotTable42").PivotFields("Plant").CurrentPage = _
        "(All)"
    With ActiveSheet.PivotTables("PivotTable42").PivotFields("Plant")
        .PivotItems("V113").Visible = False
        .PivotItems("(blank)").Visible = False
    End With
    ActiveSheet.PivotTables("PivotTable42").PivotFields("Plant"). _
        EnableMultiplePageItems = True
    ActiveSheet.PivotTables("PivotTable42").PivotFields("Storage Loc."). _
        ClearAllFilters
    ActiveSheet.PivotTables("PivotTable42").PivotFields("Storage Loc."). _
        CurrentPage = "2123"
    With ActiveSheet.PivotTables("PivotTable42").PivotFields("Old matl number")
        .Orientation = xlRowField
        .Position = 1
    End With
    ActiveSheet.PivotTables("PivotTable42").AddDataField ActiveSheet.PivotTables( _
        "PivotTable42").PivotFields("0-30 days"), "Sum of 0-30 days", xlSum
    ActiveSheet.PivotTables("PivotTable42").AddDataField ActiveSheet.PivotTables( _
        "PivotTable42").PivotFields(">30-60 days"), "Sum of >30-60 days", xlSum
    ActiveSheet.PivotTables("PivotTable42").AddDataField ActiveSheet.PivotTables( _
        "PivotTable42").PivotFields(">60-90 days"), "Sum of >60-90 days", xlSum
    ActiveSheet.PivotTables("PivotTable42").AddDataField ActiveSheet.PivotTables( _
        "PivotTable42").PivotFields(">90-120 days"), "Sum of >90-120 days", xlSum
    ActiveSheet.PivotTables("PivotTable42").AddDataField ActiveSheet.PivotTables( _
        "PivotTable42").PivotFields(">120-150 days"), "Sum of >120-150 days", xlSum
    ActiveSheet.PivotTables("PivotTable42").AddDataField ActiveSheet.PivotTables( _
        "PivotTable42").PivotFields(">150 days"), "Sum of >150 days", xlSum
    ActiveSheet.PivotTables("PivotTable42").AddDataField ActiveSheet.PivotTables( _
        "PivotTable42").PivotFields("Total"), "Sum of Total", xlSum
    ActiveSheet.PivotTables("PivotTable42").PivotFields("Old matl number"). _
        PivotFilters.Add2 Type:=xlValueDoesNotEqual, DataField:=ActiveSheet. _
        PivotTables("PivotTable42").PivotFields("Sum of Total"), Value1:=0
    Range("AF55").Select
    Sheets("Data").Select
    ActiveWorkbook.Worksheets("Divvy").PivotTables("PivotTable37").PivotCache. _
        CreatePivotTable TableDestination:="Summary!R52C32", TableName:= _
        "PivotTable43", DefaultVersion:=6
    Sheets("Summary").Select
    Cells(52, 32).Select
    With ActiveSheet.PivotTables("PivotTable43")
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
    With ActiveSheet.PivotTables("PivotTable43").PivotCache
        .RefreshOnFileOpen = False
        .MissingItemsLimit = xlMissingItemsDefault
    End With
    ActiveSheet.PivotTables("PivotTable43").RepeatAllLabels xlRepeatLabels
    With ActiveSheet.PivotTables("PivotTable43").PivotFields("Group")
        .Orientation = xlRowField
        .Position = 1
    End With
    With ActiveSheet.PivotTables("PivotTable43").PivotFields("Group")
        .Orientation = xlPageField
        .Position = 1
    End With
    ActiveSheet.PivotTables("PivotTable43").PivotFields("Group").ClearAllFilters
    ActiveSheet.PivotTables("PivotTable43").PivotFields("Group").CurrentPage = _
        "BTV"
    With ActiveSheet.PivotTables("PivotTable43").PivotFields("Plant")
        .Orientation = xlPageField
        .Position = 1
    End With
    ActiveSheet.PivotTables("PivotTable43").PivotFields("Plant").CurrentPage = _
        "(All)"
    With ActiveSheet.PivotTables("PivotTable43").PivotFields("Plant")
        .PivotItems("V113").Visible = False
        .PivotItems("(blank)").Visible = False
    End With
    ActiveSheet.PivotTables("PivotTable43").PivotFields("Plant"). _
        EnableMultiplePageItems = True
    With ActiveSheet.PivotTables("PivotTable43").PivotFields("Storage Loc.")
        .Orientation = xlPageField
        .Position = 1
    End With
    ActiveSheet.PivotTables("PivotTable43").PivotFields("Storage Loc."). _
        ClearAllFilters
    ActiveSheet.PivotTables("PivotTable43").PivotFields("Storage Loc."). _
        CurrentPage = "2119"
    With ActiveSheet.PivotTables("PivotTable43").PivotFields("Old matl number")
        .Orientation = xlRowField
        .Position = 1
    End With
    ActiveSheet.PivotTables("PivotTable43").AddDataField ActiveSheet.PivotTables( _
        "PivotTable43").PivotFields("0-30 days"), "Sum of 0-30 days", xlSum
    ActiveSheet.PivotTables("PivotTable43").AddDataField ActiveSheet.PivotTables( _
        "PivotTable43").PivotFields(">30-60 days"), "Sum of >30-60 days", xlSum
    ActiveSheet.PivotTables("PivotTable43").AddDataField ActiveSheet.PivotTables( _
        "PivotTable43").PivotFields(">60-90 days"), "Sum of >60-90 days", xlSum
    ActiveSheet.PivotTables("PivotTable43").AddDataField ActiveSheet.PivotTables( _
        "PivotTable43").PivotFields(">90-120 days"), "Sum of >90-120 days", xlSum
    ActiveSheet.PivotTables("PivotTable43").AddDataField ActiveSheet.PivotTables( _
        "PivotTable43").PivotFields(">120-150 days"), "Sum of >120-150 days", xlSum
    ActiveSheet.PivotTables("PivotTable43").AddDataField ActiveSheet.PivotTables( _
        "PivotTable43").PivotFields(">150 days"), "Sum of >150 days", xlSum
    ActiveSheet.PivotTables("PivotTable43").AddDataField ActiveSheet.PivotTables( _
        "PivotTable43").PivotFields("Total"), "Sum of Total", xlSum
    Range("AF67").Select
    Sheets("Data").Select
    ActiveWorkbook.Worksheets("Divvy").PivotTables("PivotTable37").PivotCache. _
        CreatePivotTable TableDestination:="Summary!R63C32", TableName:= _
        "PivotTable44", DefaultVersion:=6
    Sheets("Summary").Select
    Cells(63, 32).Select
    With ActiveSheet.PivotTables("PivotTable44")
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
    With ActiveSheet.PivotTables("PivotTable44").PivotCache
        .RefreshOnFileOpen = False
        .MissingItemsLimit = xlMissingItemsDefault
    End With
    ActiveSheet.PivotTables("PivotTable44").RepeatAllLabels xlRepeatLabels
    With ActiveSheet.PivotTables("PivotTable44").PivotFields("Group")
        .Orientation = xlPageField
        .Position = 1
    End With
    ActiveSheet.PivotTables("PivotTable44").PivotFields("Group").ClearAllFilters
    ActiveSheet.PivotTables("PivotTable44").PivotFields("Group").CurrentPage = _
        "BTV"
    With ActiveSheet.PivotTables("PivotTable44").PivotFields("Plant")
        .Orientation = xlPageField
        .Position = 1
    End With
    ActiveSheet.PivotTables("PivotTable44").PivotFields("Plant").CurrentPage = _
        "(All)"
    With ActiveSheet.PivotTables("PivotTable44").PivotFields("Plant")
        .PivotItems("V113").Visible = False
        .PivotItems("(blank)").Visible = False
    End With
    ActiveSheet.PivotTables("PivotTable44").PivotFields("Plant"). _
        EnableMultiplePageItems = True
    With ActiveSheet.PivotTables("PivotTable44").PivotFields("Storage Loc.")
        .Orientation = xlPageField
        .Position = 1
    End With
    ActiveSheet.PivotTables("PivotTable44").PivotFields("Storage Loc."). _
        ClearAllFilters
    ActiveSheet.PivotTables("PivotTable44").PivotFields("Storage Loc."). _
        CurrentPage = "2101"
    With ActiveSheet.PivotTables("PivotTable44").PivotFields("Old matl number")
        .Orientation = xlRowField
        .Position = 1
    End With
    ActiveSheet.PivotTables("PivotTable44").AddDataField ActiveSheet.PivotTables( _
        "PivotTable44").PivotFields("0-30 days"), "Sum of 0-30 days", xlSum
    ActiveSheet.PivotTables("PivotTable44").AddDataField ActiveSheet.PivotTables( _
        "PivotTable44").PivotFields(">30-60 days"), "Sum of >30-60 days", xlSum
    ActiveSheet.PivotTables("PivotTable44").AddDataField ActiveSheet.PivotTables( _
        "PivotTable44").PivotFields(">60-90 days"), "Sum of >60-90 days", xlSum
    ActiveSheet.PivotTables("PivotTable44").AddDataField ActiveSheet.PivotTables( _
        "PivotTable44").PivotFields(">90-120 days"), "Sum of >90-120 days", xlSum
    ActiveSheet.PivotTables("PivotTable44").AddDataField ActiveSheet.PivotTables( _
        "PivotTable44").PivotFields(">120-150 days"), "Sum of >120-150 days", xlSum
    ActiveSheet.PivotTables("PivotTable44").AddDataField ActiveSheet.PivotTables( _
        "PivotTable44").PivotFields(">150 days"), "Sum of >150 days", xlSum
    ActiveSheet.PivotTables("PivotTable44").AddDataField ActiveSheet.PivotTables( _
        "PivotTable44").PivotFields("Total"), "Sum of Total", xlSum
    ActiveSheet.PivotTables("PivotTable44").PivotFields("Old matl number"). _
        PivotFilters.Add2 Type:=xlValueDoesNotEqual, DataField:=ActiveSheet. _
        PivotTables("PivotTable44").PivotFields("Sum of Total"), Value1:=0
    ActiveSheet.PivotTables("PivotTable43").PivotFields("Old matl number"). _
        PivotFilters.Add2 Type:=xlValueDoesNotEqual, DataField:=ActiveSheet. _
        PivotTables("PivotTable43").PivotFields("Sum of Total"), Value1:=0
End Sub
Sub ongoingSKUs()
'
' ongoingSKUs Macro
'

'
    Sheets.Add.Name = "Summary"
    Sheets("Data").Select
    Cells.Select
    ActiveWorkbook.Worksheets("Divvy").PivotTables("PivotTable37").PivotCache. _
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
    With ActiveSheet.PivotTables("PivotTable38").PivotFields("Old matl number")
        .Orientation = xlRowField
        .Position = 1
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
    Range("N2").Select
    ActiveCell.FormulaR1C1 = "Ongoing SKUs"
    Range("A3:H3").Select
    Selection.Copy
    Range("N3:U3").Select
    ActiveSheet.Paste
End Sub
Sub ongoingSKUs2()
'
' ongoingSKUs2 Macro
'

'
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
    Dim startCell2 As Range, lastRow2 As Long
    Set startCell2 = Range("A4")
    lastRow2 = Cells(Rows.Count, startCell2.Column).End(xlUp).Row
    Range(Cells(lastRow2, 1), Cells(lastRow2, 8)).Select
    Selection.Copy
    Range(Cells(lastRow1 + 1, 14), Cells(lastRow1 + 1, 20)).Select
    ActiveSheet.Paste
    Range(Cells(lastRow1 + 1, 15), Cells(lastRow1 + 1, 15)).Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "=SUM(R[-23]C:R[-1]C)"
    Range(Cells(lastRow1 + 1, 15), Cells(lastRow1 + 1, 15)).Select
    Selection.AutoFill Destination:=Range(Cells(lastRow1 + 1, 15), Cells(lastRow1 + 1, 21)), Type:=xlFillDefault
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
    Range("Z4").Select
    ActiveCell.FormulaR1C1 = "=IFERROR(VLOOKUP(RC[-12],R4C1:R29C9,9,FALSE),0)"
    Range("Z4").Select
    Selection.AutoFill Destination:=Range(Cells(4, 26), Cells(lastRow1, 26)), Type:=xlFillDefault
    Range(Cells(4, 26), Cells(lastRow1, 26)).Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
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
End Sub
Sub ongoingSKUs3()
'
' ongoingSKUs3 Macro
'

'
    Range("Y4").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "=RC[-2]*RC[-1]"
    Dim startCell3 As Range, lastRow3 As Long
    Set startCell3 = Range("N4")
    lastRow3 = Cells(Rows.Count, startCell3.Column).End(xlUp).Row
    Range("Y4").Select
    Selection.AutoFill Destination:=Range(Cells(4, 25), Cells(lastRow3 - 1, 25)), Type:=xlFillDefault
    Range("AA4").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "=RC[-4]*RC[-1]"
    Range("AA4").Select
    Selection.AutoFill Destination:=Range(Cells(4, 27), Cells(lastRow3 - 1, 27)), Type:=xlFillDefault
    Range(Cells(lastRow3, 25), Cells(lastRow3, 25)).Select
    ActiveCell.FormulaR1C1 = "=SUM(R[-23]C:R[-1]C)"
    Range(Cells(lastRow3, 27), Cells(lastRow3, 27)).Select
    ActiveCell.FormulaR1C1 = "=SUM(R[-23]C:R[-1]C)"
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
'
' EOLInventory Macro
'

'
    Sheets("Data").Select
    ActiveWorkbook.Worksheets("Divvy").PivotTables("PivotTable37").PivotCache. _
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
    With ActiveSheet.PivotTables("PivotTable39").PivotFields("Old matl number")
        .Orientation = xlRowField
        .Position = 1
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
    Range("N35").Select
    ActiveCell.FormulaR1C1 = "EOL Inventory"
    Range("A36:H36").Select
    Selection.Copy
    Range("N36:U36").Select
    ActiveSheet.Paste
    Range("V36").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "Storage cost"
    Range("W36").Select
    ActiveCell.FormulaR1C1 = "Monthly cost"
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
    Dim startCell As Range, lastRow As Long
    Set startCell = Range("A37")
    lastRow = Cells(Rows.Count, startCell.Column).End(xlUp).Row
    Range(Cells(lastRow, 1), Cells(lastRow, 8)).Select
    Selection.Copy
    Range("N49:U49").Select
    ActiveSheet.Paste
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
    Range("O49").Select
    ActiveCell.FormulaR1C1 = "=SUM(R[-12]C:R[-1]C)"
    Range("O49").Select
    Selection.AutoFill Destination:=Range("O49:U49"), Type:=xlFillDefault
    Range("O49:U49").Select
    Range("X37").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "=RC[-3]*RC[-1]"
    Range("X37").Select
    Selection.AutoFill Destination:=Range("X37:X48"), Type:=xlFillDefault
    Range("X37:X48").Select
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
'
' BTVRefurb Macro
'

'
    Sheets("Data").Select
    ActiveWorkbook.Worksheets("Divvy").PivotTables("PivotTable37").PivotCache. _
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
    With ActiveSheet.PivotTables("PivotTable40").PivotFields("Old matl number")
        .Orientation = xlRowField
        .Position = 1
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
    Range("N57").Select
    ActiveCell.FormulaR1C1 = "BTV Refurb"
    Range("A68:H68").Select
    Selection.Copy
    Range("N58:U58").Select
    ActiveSheet.Paste
    Range("V58").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "Storage cost"
    Range("W58").Select
    ActiveCell.FormulaR1C1 = "Monthly cost"
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
    Dim startCell As Range, lastRow As Long
    Set startCell = Range("A69")
    lastRow = Cells(Rows.Count, startCell.Column).End(xlUp).Row
    Range(Cells(lastRow, 1), Cells(lastRow, 8)).Select
    Selection.Copy
    Range("N77:U77").Select
    ActiveSheet.Paste
    Range("O59").Select
    Application.CutCopyMode = False
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
    Range("X59").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "=RC[-3]*RC[-1]"
    Range("X59").Select
    Selection.AutoFill Destination:=Range("X59:X76"), Type:=xlFillDefault
    Range("X59:X76").Select
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
Sub cleanUp()
'
' cleanUp Macro
'

'
    Range("X3:AA3").Select
    Selection.Font.Bold = True
    Range("X27:AA27").Select
    Selection.Font.Bold = True
    Range("X36").Select
    Selection.Font.Bold = True
    Range("X49").Select
    Selection.Font.Bold = True
    Range("X58").Select
    Selection.Font.Bold = True
    Range("X77").Select
    Selection.Font.Bold = True
End Sub
Sub inventoryAgingReport1()
'
' inventoryAgingReport1 Macro
'
' Keyboard Shortcut: Ctrl+Shift+Q
'
    Range("A1").Select
        ' AlltoData Macro
'
' Keyboard Shortcut: Ctrl+a
'
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
    Range("O1").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "Total"
    Range("P1").Select
    ActiveCell.FormulaR1C1 = "Data"
    Range("Q1").Select
    ActiveCell.FormulaR1C1 = "Group"
    Range("R1").Select
    ActiveCell.FormulaR1C1 = "ODM"
    Range("O2").Select
    ActiveCell.FormulaR1C1 = "=SUM(RC[-7],RC[-6],RC[-5],RC[-4],RC[-3],RC[-2])"
    Range("P2").Select
    ActiveCell.FormulaR1C1 = "=RC[-15]&RC[-11]&RC[-10]"
    Range("Q2").Select
    ActiveCell.FormulaR1C1 = _
        "=VLOOKUP(RC[-1],'[2019-09-18-Inventory-Aging-Report.xlsm]Data'!C16:C17,2,FALSE)"
    Range("R2").Select
    ActiveCell.FormulaR1C1 = "=IF(OR(RC[-12]=""2101"",RC[-12]=""2119"",RC[-12]=""2123"",RC[-12]=""2128"",AND(RC[-13]=""V113"",RC[-12]=""4001"")),RC[-3],0)"
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
        '
' DatatoDivvy1 Macro
'

'
    Sheets.Add.Name = "Divvy"
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
    With ActiveSheet.PivotTables("PivotTable28").PivotFields("Old matl number")
        .Orientation = xlRowField
        .Position = 1
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
    ActiveSheet.PivotTables("PivotTable28").AddDataField ActiveSheet.PivotTables( _
        "PivotTable28").PivotFields("ODM"), "Sum of ODM", xlSum
    'Columns("I:I").Select
    'Selection.Copy
    'Columns("M:M").Select
    'ActiveSheet.Paste
    Range("M3").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "ODM"
    Range("H11").Select
    ActiveSheet.PivotTables("PivotTable28").PivotFields("Sum of ODM").Orientation _
        = xlHidden
    Range("I3").Select
    ActiveCell.FormulaR1C1 = "Orders"
    Range("J3").Select
    ActiveCell.FormulaR1C1 = "Balance"
    Range("L3").Select
    ActiveCell.FormulaR1C1 = "VIZIO"
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
    Range("L4").Select
    ActiveCell.FormulaR1C1 = "=RC[-4]-RC[1]"
    Range("L4").Select
    Range("M4").Select
    ActiveCell.FormulaR1C1 = "=IFERROR(VLOOKUP(RC1,C16:C23,8,FALSE),0)"
    Range("L4").Select
    Dim startCell1 As Range, lastRow1 As Long
    Set startCell1 = Range("H4")
    lastRow1 = Cells(Rows.Count, startCell1.Column).End(xlUp).Row
    Selection.AutoFill Destination:=Range(Cells(4, 12), Cells(lastRow1, 12))
    Range("M4").Select
    Selection.AutoFill Destination:=Range(Cells(4, 13), Cells(lastRow1 - 1, 13))
    Range(Cells(lastRow1, 13), Cells(lastRow1, 13)).Select
    Range(Cells(lastRow1, 13), Cells(lastRow1, 13)).Formula = "=SUM(" & Range(Cells(4, 13), Cells(lastRow1 - 1, 13)).Address(False, False) & ")"
    Range("J4").Select
    ActiveCell.FormulaR1C1 = "=RC[-2]-RC[-1]"
    Range("J4").Select
    Selection.AutoFill Destination:=Range(Cells(4, 10), Cells(lastRow1 - 1, 10)), Type:=xlFillDefault
    Range(Cells(lastRow1, 10), Cells(lastRow1, 10)).Select
    Range(Cells(lastRow1, 10), Cells(lastRow1, 10)).Formula = "=SUM(" & Range(Cells(4, 10), Cells(lastRow1 - 1, 10)).Address(False, False) & ")"
    Sheets("Data").Select
    Range("A1").Select
        '
' DatatoDivvy2 Macro
'

'
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
    Range("Q3").Select
    With ActiveSheet.PivotTables("PivotTable34").PivotFields("Old matl number")
        .Orientation = xlRowField
        .Position = 1
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
        '
' ongoingSKUs Macro
'

'
    Sheets.Add.Name = "Summary"
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
    With ActiveSheet.PivotTables("PivotTable38").PivotFields("Old matl number")
        .Orientation = xlRowField
        .Position = 1
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
    Range("N2").Select
    ActiveCell.FormulaR1C1 = "Ongoing SKUs"
    Range("A3:H3").Select
    Selection.Copy
    Range("N3:U3").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
End Sub
Sub inventoryAgingReport2()
'
' inventoryAgingReport2 Macro
'

'
        '
' ongoingSKUs2 Macro
'

'
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
    Dim startCell2 As Range, lastRow2 As Long
    Set startCell2 = Range("A4")
    lastRow2 = Cells(Rows.Count, startCell2.Column).End(xlUp).Row
    Range(Cells(lastRow2, 1), Cells(lastRow2, 8)).Select
    Selection.Copy
    Range(Cells(lastRow1 + 1, 14), Cells(lastRow1 + 1, 20)).Select
    ActiveSheet.Paste
    Range(Cells(lastRow1 + 1, 15), Cells(lastRow1 + 1, 15)).Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "=SUM(R[-23]C:R[-1]C)"
    Range(Cells(lastRow1 + 1, 15), Cells(lastRow1 + 1, 15)).Select
    Selection.AutoFill Destination:=Range(Cells(lastRow1 + 1, 15), Cells(lastRow1 + 1, 21)), Type:=xlFillDefault
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
    Range("Z4").Select
    ActiveCell.FormulaR1C1 = "=IFERROR(VLOOKUP(RC[-12],R4C1:R29C9,9,FALSE),0)"
    Range("Z4").Select
    Selection.AutoFill Destination:=Range(Cells(4, 26), Cells(lastRow1, 26)), Type:=xlFillDefault
    Range(Cells(4, 26), Cells(lastRow1, 26)).Select
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
        '
' ongoingSKUs3 Macro
'

'
    Range("Y4").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "=RC[-2]*RC[-1]"
    Dim startCell3 As Range, lastRow3 As Long
    Set startCell3 = Range("N4")
    lastRow3 = Cells(Rows.Count, startCell3.Column).End(xlUp).Row
    Range("Y4").Select
    Selection.AutoFill Destination:=Range(Cells(4, 25), Cells(lastRow3 - 1, 25)), Type:=xlFillDefault
    Range("AA4").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "=RC[-4]*RC[-1]"
    Range("AA4").Select
    Selection.AutoFill Destination:=Range(Cells(4, 27), Cells(lastRow3 - 1, 27)), Type:=xlFillDefault
    Range(Cells(lastRow3, 25), Cells(lastRow3, 25)).Select
    ActiveCell.FormulaR1C1 = "=SUM(R[-23]C:R[-1]C)"
    Range(Cells(lastRow3, 27), Cells(lastRow3, 27)).Select
    ActiveCell.FormulaR1C1 = "=SUM(R[-23]C:R[-1]C)"
    Range(Cells(4, 25), Cells(lastRow3, 25)).Select
    Selection.Style = "Currency"
    Range(Cells(4, 27), Cells(lastRow3, 27)).Select
    Selection.Style = "Currency"
    Range("X3:AA3").Select
    Selection.Font.Bold = True
    Range(Cells(lastRow3, 24), Cells(lastRow3, 27)).Select
    Selection.Font.Bold = True
        '
' EOLInventory Macro
'

'
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
    With ActiveSheet.PivotTables("PivotTable39").PivotFields("Old matl number")
        .Orientation = xlRowField
        .Position = 1
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
    Range("N35").Select
    ActiveCell.FormulaR1C1 = "EOL Inventory"
    Range("A36:H36").Select
    Selection.Copy
    Range("N36:U36").Select
    ActiveSheet.Paste
    Range("V36").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "Storage cost"
    Range("W36").Select
    ActiveCell.FormulaR1C1 = "Monthly cost"
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
    Dim startCell4 As Range, lastRow4 As Long
    Set startCell4 = Range("A37")
    lastRow4 = Cells(Rows.Count, startCell4.Column).End(xlUp).Row
    Range(Cells(lastRow4, 1), Cells(lastRow4, 8)).Select
    Selection.Copy
    Range("N49:U49").Select
    ActiveSheet.Paste
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
    Range("O49").Select
    ActiveCell.FormulaR1C1 = "=SUM(R[-12]C:R[-1]C)"
    Range("O49").Select
    Selection.AutoFill Destination:=Range("O49:U49"), Type:=xlFillDefault
    Range("O49:U49").Select
    Range("X37").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "=RC[-3]*RC[-1]"
    Range("X37").Select
    Selection.AutoFill Destination:=Range("X37:X48"), Type:=xlFillDefault
    Range("X37:X48").Select
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
'
' BTVRefurb Macro
'

'
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
    With ActiveSheet.PivotTables("PivotTable40").PivotFields("Old matl number")
        .Orientation = xlRowField
        .Position = 1
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
    Range("N57").Select
    ActiveCell.FormulaR1C1 = "BTV Refurb"
    Range("A68:H68").Select
    Selection.Copy
    Range("N58:U58").Select
    ActiveSheet.Paste
    Range("V58").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "Storage cost"
    Range("W58").Select
    ActiveCell.FormulaR1C1 = "Monthly cost"
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
    Dim startCell5 As Range, lastRow5 As Long
    Set startCell5 = Range("A69")
    lastRow5 = Cells(Rows.Count, startCell5.Column).End(xlUp).Row
    Range(Cells(lastRow5, 1), Cells(lastRow5, 8)).Select
    Selection.Copy
    Range("N77:U77").Select
    ActiveSheet.Paste
    Range("O59").Select
    Application.CutCopyMode = False
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
    Range("X59").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "=RC[-3]*RC[-1]"
    Range("X59").Select
    Selection.AutoFill Destination:=Range("X59:X76"), Type:=xlFillDefault
    Range("X59:X76").Select
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
    '
' dataToSummary2 Macro
'

'
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
    End With
    ActiveSheet.PivotTables("PivotTable41").PivotFields("Storage Loc."). _
        EnableMultiplePageItems = True
    With ActiveSheet.PivotTables("PivotTable41").PivotFields("Old matl number")
        .Orientation = xlRowField
        .Position = 1
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
Sub cleanUp2()
'
' cleanUp2 Macro
'

'
    Range("Y27").Select
    Selection.Style = "Currency"
    Range("AA27").Select
    Selection.Style = "Currency"
    Range("X49").Select
    Selection.Style = "Currency"
    Range("X77").Select
    Selection.Style = "Currency"
    
End Sub
