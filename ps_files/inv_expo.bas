Attribute VB_Name = "Module11"
Sub Prep()
    'Author Nirmanpreet Singh
    'Ver 1
    'Description : Used to Prep data to used for blockshop update
    '
    
    Dim lRow        As Long
    Dim lCol        As Long
    
    lRow = Cells.Find(What:="*", _
           After:=Range("A1"), _
           LookAt:=xlPart, _
           LookIn:=xlFormulas, _
           SearchOrder:=xlByRows, _
           SearchDirection:=xlPrevious, _
           MatchCase:=False).Row
    Range("O2").Select
    ActiveCell.FormulaR1C1 = "=MAX(0,SUM(RC[-3]:RC[-1]))"
    Range("O2").Select
    Selection.AutoFill Destination:=Range("O2:O" & lRow)
    Range("O2:O" & lRow).Select
    Range("P1").Select
    ActiveCell.FormulaR1C1 = "Total"
    Range("O2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Range("P1").Select
    Application.CutCopyMode = False
    Range("O2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Range(Selection, Selection.End(xlUp)).Select
    Range("P2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                           :=False, Transpose:=False
    Columns("A:H").Select
    Application.CutCopyMode = False
    Selection.Delete Shift:=xlToLeft
    Columns("B:G").Select
    Selection.Delete Shift:=xlToLeft
    Range("J13").Select
End Sub

