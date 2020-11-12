VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Public newMonth As Boolean, inProgress As Boolean
Private Sub Worksheet_Change(ByVal Target As Range)
    Dim task As Range, row As Integer, row2 As Integer
    Dim caseNo As Integer, strTemp(6) As String

    ' The variable task contains the cells that will
    ' cause an alert when they are changed.
    Set task = Range("B3:B10000")
    
    If Not Application.Intersect(task, Range(Target.Address)) Is Nothing Then
        If Not inProgress Then
            inProgress = True
            
            ActiveSheet.Unprotect
            Application.ScreenUpdating = False
            Rows.EntireRow.Hidden = False
            ActiveSheet.Protect UserInterfaceOnly:=True
            
            newMonth = False
            If Day(Date) > 25 And Range("C1").Value = 0 Then
                newMonth = True 'marker that it's ready for a new month
            ElseIf Day(Date) < 25 And Range("C1").Value = 5 Then 'time for a new month
                Range("C1").Value = 0
            End If
            
            If newMonth Then
                Range("B4").Select
                Selection.End(xlDown).Select
                strTemp(1) = ActiveCell.Value
                ActiveCell.ClearContents
                If ActiveCell.Value > 0 Then
                    strTemp(0) = ActiveCell.Offset(0, 1).Value
                    strTemp(2) = ActiveCell.Offset(0, 2).Value
                    strTemp(3) = ActiveCell.Offset(0, 3).Value
                    strTemp(4) = ActiveCell.Offset(0, 4).Value
                    strTemp(5) = ActiveCell.Offset(0, 5).Value
                    strTemp(6) = ActiveCell.Offset(0, 6).Formula
                End If
                Debug.Print "row:" & Range("B4").End(xlDown).row & "    val:" & Range("B4").End(xlDown).Value
                Call NewMonthly("Lab Audit", "Chad", "Safety audit for Mateer test lab, due by the first week of " & Format(Date + 14, "mmm"))
                Debug.Print "row:" & Range("B4").End(xlDown).row & "    val:" & Range("B4").End(xlDown).Value
                Call NewMonthly("Backup", "-", "Update backup repo for all programs on Git")
                Debug.Print "row:" & Range("B4").End(xlDown).row & "    val:" & Range("B4").End(xlDown).Value
                Range("C1").Value = 5 'indicates program has run for this month
            End If
            
            Range("B4").Select
            Selection.End(xlDown).Select
            row = ActiveCell.row

            If newMonth Then
                row = row + 1
                Range("B" & row).Value = strTemp(1)
                If strTemp(0) <> "" Then
                    Range("A" & row).Value = strTemp(0)
                    Range("C" & row).Value = strTemp(2)
                    Range("D" & row).Value = strTemp(3)
                    Range("E" & row).Value = strTemp(4)
                    Range("F" & row).Value = strTemp(5)
                    Range("G" & row).Formula = strTemp(6)
                End If
            End If
            
            row2 = Range("A4").End(xlDown).row
            If row > row2 Then 'A needs date inserted
                Range("A" & row).Value = Date
            ElseIf row2 > row Then
                Range("A" & row2).ClearContents
            End If
            
            Range("B" & row).Locked = False
            Range("B" & row + 1).Locked = False
            Range("C" & row).Locked = False
            Range("C" & row + 1).Locked = False
            Range("C" & row).Select
            
             ActiveSheet.Range("$E:$E").AutoFilter Field:=1, Criteria1:="<>Done", _
        Operator:=xlAnd, Criteria2:="<>Abandoned"
        
            Application.ScreenUpdating = True
            inProgress = False
        End If 'if not inprogress
    End If 'if column B
End Sub
