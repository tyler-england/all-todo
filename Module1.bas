Attribute VB_Name = "Module1"
Sub NewEntry()
Attribute NewEntry.VB_ProcData.VB_Invoke_Func = " \n14"
'
' NewEntry Macro
'
    Range("B4").Select
    Selection.End(xlDown).Select
    Selection.Offset(0, -1).Select
    ActiveCell.Value = Date

End Sub

Function NewMonthly(strTask As String, contName As String, strDesc As String)
'
' NewMonthly Macro
'
    Dim rowNum As Integer

    Range("B4").Select
    Selection.End(xlDown).Select
    rowNum = ActiveCell.row + 1
    Selection.Offset(1, -1).Select
    ActiveCell.Value = Date
    Range("C" & rowNum).Value = contName
    Range("D" & rowNum).Value = strDesc
    Range("E" & rowNum).Value = "Ready"
    Range("F" & rowNum).Value = "Low"
    Range("B" & rowNum).Value = strTask

End Function
