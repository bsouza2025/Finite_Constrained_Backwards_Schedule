Attribute VB_Name = "Module2"
Sub WipeOutAndReboot()
Attribute WipeOutAndReboot.VB_ProcData.VB_Invoke_Func = " \n14"
'
' WipeOutAndReboot Macro
'

'
    Application.Calculation = xlCalculationManual
    Sheets("Pivot Templates").Activate
    Range("U4").Select
        Selection.FormulaR1C1 = _
        "=IF(RC20=""SHIP DATE"",SUMIFS('NEW Projected Revenue 2024'!C9,'NEW Projected Revenue 2024'!C4,'Pivot Templates'!R[1]C3,'NEW Projected Revenue 2024'!C1,'Pivot Templates'!R[1]C2,'NEW Projected Revenue 2024'!C6,'Pivot Templates'!R3C),"""")"

    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    ActiveWindow.SmallScroll Down:=15
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveWindow.SmallScroll Down:=-24
    With Selection.Interior
        .Pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    Range("U4").Select
    Selection.Copy
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    ActiveWindow.SmallScroll Down:=6
    Selection.PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Application.Calculate
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlWhole, SkipBlanks _
        :=False, Transpose:=False
    Range("J5").Select
    ReloadedInitialLoad
    Application.Calculation = xlCalculationAutomatic
    Sheets("Pivot Templates").Activate
End Sub
