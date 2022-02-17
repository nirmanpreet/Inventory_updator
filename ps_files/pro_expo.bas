Attribute VB_Name = "Module1"
Sub Prep()
Attribute Prep.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro1 Macro
'

'
    Columns("A:N").Select
    Selection.Delete Shift:=xlToLeft
    Columns("B:E").Select
    Selection.Delete Shift:=xlToLeft
    Columns("D:AD").Select
    Selection.Delete Shift:=xlToLeft
    Cells.Select
    ActiveWorkbook.Worksheets("products_export_1").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("products_export_1").Sort.SortFields.Add2 Key:= _
        Range("A1:A17601"), SortOn:=xlSortOnValues, Order:=xlDescending, _
        DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("products_export_1").Sort
        .SetRange Range("A1:AV17601")
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Range("D2").Select
    ActiveCell.FormulaR1C1 = "=IF(RC[-1]=0,RC[-2],RC[-1])"
    Range("D2").Select
    Selection.AutoFill Destination:=Range("D2:D2605")
    Range("D2:D2605").Select
    Selection.Copy
    Range("C2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Columns("D:D").Select
    Application.CutCopyMode = False
    Selection.Delete Shift:=xlToLeft
End Sub
