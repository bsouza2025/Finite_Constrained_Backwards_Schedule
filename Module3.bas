Attribute VB_Name = "Module3"
Sub ReloadedInitialLoad()
Attribute ReloadedInitialLoad.VB_ProcData.VB_Invoke_Func = " \n14"
'
' ReloadedInitialLoad Macro
'

'
    Sheets("WC Load").Activate
    Range("J5").Select
    Selection.FormulaR1C1 = _
        "=IF(SUMIFS('WC Pre-Load'!C26,'WC Pre-Load'!C24,'WC Load'!RC3,'WC Pre-Load'!C25,'WC Load'!R3C)>0,SUMIFS('WC Pre-Load'!C26,'WC Pre-Load'!C24,'WC Load'!RC3,'WC Pre-Load'!C25,'WC Load'!R3C),"""")"
    Selection.Copy
    Range("J5:IU100").Select
    Selection.PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False

    Application.CutCopyMode = False
    Application.Calculate
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlWhole, SkipBlanks _
        :=False, Transpose:=False
    Range("J5").Select
    Application.CutCopyMode = False
End Sub
