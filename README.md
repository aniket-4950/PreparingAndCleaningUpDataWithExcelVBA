# PreparingAndCleaningUpDataWithExcelVBA

The VBA Code:

Sub LoopingThroughSheets()
    Dim ws As Worksheet
    
    For Each ws In Worksheets
        Worksheets(ws.Name).Select
        If Range("A1") <> "Divison" Then    'adding logic for formatting a newly entered worksheet so that all the other worksheets doesn't get formatted again
            InsertHeaders
            FormattingHeaders
        End If
    Next ws
        
End Sub

Sub InsertHeaders()
'
' InsertHeaders Macro
' Inserts header by appending a column at the beginning.
'

'
    Rows("1:1").Select
    Selection.Insert Shift:=xlDown
    Range("A1").Select
    ActiveCell.FormulaR1C1 = "Divison"
    Range("B1").Select
    ActiveCell.FormulaR1C1 = "Category"
    Range("C1").Select
    ActiveCell.FormulaR1C1 = "Jan"
    Range("D1").Select
    ActiveCell.FormulaR1C1 = "Feb"
    Range("E1").Select
    ActiveCell.FormulaR1C1 = "March"
    Range("F1").Select
    ActiveCell.FormulaR1C1 = "Total"
    Range("G7").Select
End Sub
Sub FormattingHeaders()
'
' FormattingHeaders Macro
' Formats the headers and adds the currency symbols.
'

'
    Range("A1:F1").Select
    With Selection.Font
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
    End With
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    Selection.Font.Bold = True
    Range("C2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Style = "Currency"
    Range("A1:F1").Select
    Columns("C:C").EntireColumn.AutoFit
    Range("A2").Select
End Sub
