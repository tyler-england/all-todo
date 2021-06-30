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

Function ExportModules() As Boolean
    Dim wbMacro As Workbook, varVar As Variant, bOpen As Boolean
    For Each varVar In Application.Workbooks
        If UCase(varVar.Name) = "MACROBOOK.XLSM" Then
            bOpen = True
            Set wbMacro = varVar
            Exit For
        End If
    Next
    If Not bOpen Then Set wbMacro = Workbooks.Open("\\PSACLW02\HOME\SHARED\MacroBook.xlsm")
    Application.Run "'" & wbMacro.Name & "'!ExportModules", ThisWorkbook
    If Not bOpen Then wbMacro.Close savechanges:=False
    ExportModules = True
End Function

Public Sub ErrorRep(rouName, rouType, curVal, errNum, errDesc, miscInfo)
    Dim wbMacro As Workbook, varVar As Variant, bOpen As Boolean
    bNewMsg = True 'default value
    If iNumMsgs > 0 Then 'at least one email has been generated already
        For Each varVar In arrErrorEmails 'see if there were any matches
            If UCase(varVar) = UCase(ThisWorkbook.Name & "-" & errNum) Then Exit Sub 'repeat message (this was already sent this session)
        Next
    End If
    For Each varVar In Application.Workbooks
        If UCase(varVar.Name) = "MACROBOOK.XLSM" Then
            bOpen = True
            Set wbMacro = varVar
            Exit For
        End If
    Next
    If Not bOpen Then Set wbMacro = Workbooks.Open("\\PSACLW02\HOME\SHARED\MacroBook.xlsm")
    Application.Run "'MacroBook.xlsm'!ErrorReport", rouName, rouType, curVal, errNum, errDesc, miscInfo
    If Not bOpen Then wbMacro.Close savechanges:=False
    iNumMsgs = iNumMsgs + 1
    ReDim Preserve arrErrorEmails(iNumMsgs)
    arrErrorEmails(iNumMsgs) = ThisWorkbook.Name & "-" & errNum
End Sub
