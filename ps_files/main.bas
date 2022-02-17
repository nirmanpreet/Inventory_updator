Attribute VB_Name = "Module1"
Sub Macro1()
Attribute Macro1.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro1 Macro
'

'
    Application.DisplayAlerts = False
    Cells.Select
    ActiveWorkbook.Names.Add Name:="Price", RefersToR1C1:= _
        "=products_export_1!R1:R1048576"
    ActiveWorkbook.Names("Price").Comment = ""
    Sheets("inventory_export_1").Select
    Cells.Select
    ActiveWorkbook.Names.Add Name:="Stock", RefersToR1C1:= _
        "=inventory_export_1!R1:R1048576"
    ActiveWorkbook.Names("Stock").Comment = ""
    Sheets("Products").Select
    Range("G4").Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-5],Price,3,)"
    Range("G4").Select
    Selection.AutoFill Destination:=Range("G4:G878")
    Range("G4:G878").Select
    Range("H4").Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-6],Price,2,)"
    Range("H4").Select
    Selection.AutoFill Destination:=Range("H4:H878")
    Range("H4:H878").Select
    Range("J4").Select
    ActiveCell.FormulaR1C1 = "0"
    Sheets("Products").Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-8],Stock,2,)"
    Range("J4").Select
    Selection.AutoFill Destination:=Range("J4:J878")
    Range("J4:J878").Select
    Range("G4").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    
    Range("G4").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("H4").Select
    Range(Selection, Selection.End(xlDown)).Select
    Application.CutCopyMode = False
    Selection.Copy
   
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("J4").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range("J4").Select
    Range(Selection, Selection.End(xlDown)).Select
    Application.CutCopyMode = False
    Selection.Copy
    
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("G4").Select
    Sheets("inventory_export_1").Select
    Application.CutCopyMode = False
    ActiveWindow.SelectedSheets.Delete
    Sheets("products_export_1").Select
    ActiveWindow.SelectedSheets.Delete
End Sub
