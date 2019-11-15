Attribute VB_Name = "Module1"
Sub NewEntry()
Attribute NewEntry.VB_ProcData.VB_Invoke_Func = " \n14"
'
' NewEntry Macro
'
    Range("A1").Select
    Selection.Copy
    Range("B4").Select
    Selection.End(xlDown).Select
    Selection.Offset(0, -1).Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
        
    
End Sub

Sub NewMonthly(strTask As String, contName As String, strDesc As String)
'
' NewMonthly Macro
'

    Dim row As Integer
    
    If Day(Date) > 5 Then 'past first 5 days of month
        Range("C1").Value = 5 'marker that it's ready for a new month
    End If
    
    If Day(Date) > 25 And Range("C1").Value = 5 Then 'first few days of month & ready for new month
        Range("C1").Value = 0 'new month is here
    End If

    If Range("C1").Value = 0 Then 'new month is here

        Range("A1").Select
        Selection.Copy
        Range("B4").Select
        Selection.End(xlDown).Select
        row = ActiveCell.row
        Selection.Offset(1, -1).Select
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
        Application.CutCopyMode = False
        Range("D1").Value = Range("B" & row).Value 'putting in "B" would trigger macro
        Range("C" & row + 1).Value = Range("C" & row).Value
        Range("D" & row + 1).Value = Range("D" & row).Value
        Range("E" & row + 1).Value = Range("E" & row).Value
        Range("F" & row + 1).Value = Range("F" & row).Value
        Range("G" & row + 1).Value = Range("G" & row).Value
        Range("C" & row).Select
        Range("C" & row).Value = contName
        Range("D" & row).Value = strDesc
        Range("E" & row).Value = "Ready"
        Range("F" & row).Value = "Low"
        Range("C1").Value = 1 'C1=1 -> first 5 days of month, but program has run (don't run again, even in first 5 days, until C1=5)
        Range("B" & row).Value = strTask
        Range("B" & row + 1).Value = Range("D1").Value
        Range("D1").Value = ""
    
    End If
    
        Range("A4").Select
        Selection.End(xlDown).Select
        Selection.Offset(0, 2).Select
        Range("A4", ActiveCell).Locked = True

End Sub
