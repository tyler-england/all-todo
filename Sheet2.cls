VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Worksheet_Change(ByVal Target As Range)
    Dim task As Range, row As Integer, row2 As Integer
    
    ' The variable task contains the cells that will
    ' cause an alert when they are changed.
    Set task = Range("B3:B5000")
    
    If Not Application.Intersect(task, Range(Target.Address)) _
           Is Nothing Then
            
            ActiveSheet.Unprotect
            
            Application.ScreenUpdating = False
            
            Rows.EntireRow.Hidden = False
            
            ActiveSheet.Protect UserInterfaceOnly:=True
            Call NewEntry
            Range("B4").Select
            Selection.End(xlDown).Select
            row = ActiveCell.row
            Selection.Offset(1, -1).Select
            Range(Selection, Selection.End(xlDown)).Select
            Selection.ClearContents
            Range("B" & row).Locked = False
            Range("B" & row + 1).Locked = False
            Range("C" & row).Locked = False
            Range("C" & row + 1).Locked = False
            Range("C" & row).Select
            
            If ActiveSheet.CodeName = "Sheet1" Then
            
            ActiveSheet.Range("$E:$E").AutoFilter Field:=1, Criteria1:="<>Done", _
        Operator:=xlAnd
            
            ElseIf ActiveSheet.CodeName = "Sheet2" Then
            
            ActiveSheet.Range("$F:$F").AutoFilter Field:=1, Criteria1:="<>Discarded", _
        Operator:=xlAnd, Criteria2:="<>Returned"
        
            End If
        
            Application.ScreenUpdating = True
            
           End If
End Sub
