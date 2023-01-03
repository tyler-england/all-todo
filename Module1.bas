Option Explicit
Dim bGoing As Boolean

Function SheetUpdate(ByVal Target) 'when a cell is edited
    Dim i As Integer, iRow As Integer, iLastRow As Integer
    Dim iColImpt As Integer, iColRec As Integer, iColDue As Integer, iColDone As Integer, iColLast As Integer
    Dim lInterior As Long, lFont As Long
    Dim sColLastLet As String, sColDueLet As String
    Dim dFri As Date, dGen As Date
    Dim rngCell As Range
    Dim bHide As Boolean, bBold As Boolean, bEmpty As Boolean, bGo As Boolean, bColA As Boolean, bColDue As Boolean
    
    If Target.Address = "$K$1" Then 'keyword highlighting -> exit
        bGo = KeywordFunc(Target) 'true if keyword was entered, false if "keyword"
        Application.EnableEvents = True
        If bGo Then Exit Function
    End If
    
    If bGoing Then Exit Function
    
    With ThisWorkbook.Worksheets(1)
        iColDone = WorksheetFunction.Match("*Done*", .Range("1:1"), 0)
        iColImpt = WorksheetFunction.Match("*Priority*", .Range("1:1"), 0)
        iColRec = WorksheetFunction.Match("*Receive*", .Range("1:1"), 0)
        iColDue = WorksheetFunction.Match("*Due*", .Range("1:1"), 0)
        sColDueLet = Split(Cells(1, iColDue).Address, "$")(1)
        iColLast = .Range("A1").End(xlToRight).Column
        sColLastLet = Split(Cells(1, iColLast).Address, "$")(1)
        iLastRow = .Range("B10000").End(xlUp).Row

        For Each rngCell In Target 'do this for every cell in selection
            If rngCell.Row > iLastRow Then iLastRow = rngCell.Row
            If Not bColA Then 'see whether to trigger colors, and whether col A is being edited
                If rngCell.Column = 1 Then
                    bGo = True
                    bColA = True
                ElseIf rngCell.Column = iColDue Then 'make sure day is a week/work day
                    bGo = True
                    bGoing = True 'adjusting a trigger column cell's value
                    If VarType(rngCell.Value) = vbInteger Or VarType(rngCell.Value) = vbDate Then
                        dGen = rngCell.Value
                        If dGen > 0 Then
                            Do While Weekday(dGen) < 2 Or Weekday(dGen) > 6
                                dGen = dGen - 1
                            Loop
                            rngCell.Value = dGen
                        End If
                    End If
                    bGoing = False
                ElseIf Target.Address = "$K$1" Or rngCell.Column = iColDone Then ' update when changing search or marking done
                    bGo = True
                End If
            End If
        Next
        Call RecurringItems
        If bGo Then
            bGo = False
            iRow = 2
            Do While .Cells(iRow, 1).Value > 0 Or iRow < iLastRow 'go through entire list
                If Not .Rows(iRow).Hidden Then 'ignore hidden rows
                    If .Cells(iRow, 1).Value > 0 And .Cells(iRow, iColRec).Value = 0 Then .Cells(iRow, iColRec).Value = Date 'anything missing a received date should have current date
                    bHide = False 'default
                    If InStr(1, .Cells(iRow, iColImpt).Value, "high", 1) > 0 Then
                        bBold = True
                    Else
                        bBold = False
                    End If
                    lInterior = xlNone 'default
                    lFont = 0 'default
                    If .Cells(iRow, iColDone).Value > 0 Then 'highlight complete/abandoned items & hide them
                        If InStr(1, .Cells(iRow, iColDone).Value, "abandon", 1) > 0 Then 'abandoned
                            lInterior = 13551615
                            lFont = 393372
                        Else 'done/completed
                            lInterior = 13561798
                            lFont = 24832
                        End If
                        bBold = False
                        bHide = True
                    ElseIf IsDate(.Cells(iRow, iColDue).Value) Then
                        If .Cells(iRow, iColDue).Value - Date < 4 Or (Weekday(Date) > 4 And .Cells(iRow, iColDue).Value - Date < 6) Then 'highlight/bold impending (within a week) items
                            If .Cells(iRow, iColDue).Value - Date < 1 Then 'immediately due -> RED
                                lInterior = 255
                                lFont = 16777215
                            ElseIf .Cells(iRow, iColDue).Value - Date < 2 Or _
                            (Weekday(.Cells(iRow, iColDue).Value) < 3 And Weekday(Date) > 5) Then 'due soon -> orange
                                lInterior = 4434687
                            Else 'due soon-ish -> yellow
                                lInterior = 10284031
                            End If
                            bBold = True
                        End If
                    End If

                    .Range("A" & iRow & ":" & sColLastLet & iRow).Interior.Color = lInterior 'set colors
                    .Range("A" & iRow & ":" & sColLastLet & iRow).Font.Color = lFont 'set colors
                    .Range("A" & iRow & ":" & sColLastLet & iRow).Font.Bold = bBold 'bold
                    
                    If bColA And bHide Then .Rows(iRow).Hidden = True 'hide completed/abandoned when adding new items
                End If
                iRow = iRow + 1
            Loop
            Call ColorCategories
            If .Range("A" & iLastRow).Value = 0 Then 'value was deleted
                If bColA Then .Range("A" & iLastRow).Interior.Color = xlNone
                bEmpty = True 'check if row is empty
                For i = 2 To iColLast
                    If .Cells(iLastRow, i).Value > 0 And i <> iColRec Then
                        bEmpty = False
                        Exit For
                    End If
                Next
                If bEmpty Then
                    .Cells(iLastRow, iColRec).ClearContents 'get rid of autofilled receive date
                    With .Rows(iLastRow)
                        .Interior.Color = xlNone
                        .Font.Color = 0
                        .Font.Bold = False
                    End With
                End If
            End If
            
        End If
        
        '.Range("K1").Value = "Keyword"
    End With
    
End Function

Function RecurringItems() 'add any recurring items

    If Sheet2.Cells(2, WorksheetFunction.Match("*update*", Sheet2.Rows(1), 0)).Value >= Date Then Exit Function 'updated today already
    Application.EnableEvents = False
    
    Dim rngVals As Range
    Dim vVals As Variant, vDatesUsed As Variant, vDueDates As Variant
    Dim sTypes() As String, sDescs() As String, sConts() As String, sPrtis() As String, sEases() As String
    Dim sCol As String, sErrs() As String, sError As String, sTyp As String, sMonth As String
    Dim dDues() As Date, dUpdated As Date, dDayInc As Date
    Dim i As Integer, iInd As Integer, iDay As Integer, iErr As Integer
    Dim iFirstRow As Integer, iLastRow As Integer, iMonth As Integer
    Dim iColUpdate As Integer, iColUsed As Integer, iColDates As Integer
    Dim wsRecur As Worksheet, wsToDo As Worksheet
    Dim bAdd As Boolean, bNewPer As Boolean
    
    Set wsToDo = ThisWorkbook.Worksheets(1)
    Set wsRecur = ThisWorkbook.Worksheets(2)
    iColUpdate = WorksheetFunction.Match("*update*", wsRecur.Rows(1), 0)
    iColUsed = WorksheetFunction.Match("*used*", wsRecur.Rows(1), 0)
    iLastRow = wsRecur.Range("A5000").End(xlUp).Row
    ReDim dUseds(iLastRow) 'necessary variable?
    
    sCol = Split(Cells(1, iColUpdate).Address, "$")(1)
    Set rngVals = wsRecur.Range("A1:" & sCol & iLastRow) 'range with values (incl. headers)
    vVals = rngVals.Value2 'get values as range variant
    sCol = Split(Cells(1, iColUsed).Address, "$")(1)
    vDatesUsed = wsRecur.Range(sCol & "1:" & sCol & iLastRow).Value 'date used column values
    On Error GoTo erradd  'for notification of failed rows
    For i = 2 To UBound(vVals) 'for each row
        bNewPer = False 'default
        bAdd = False 'default
        If LCase(vVals(i, 1)) Like "*week*" Then 'weekly
            sTyp = "week"
        ElseIf LCase(vVals(i, 1)) Like "*month*" Then 'monthly
            sTyp = "month"
        ElseIf LCase(vVals(i, 1)) Like "*quarter*" Then 'quarterly
            sTyp = "quarter"
        ElseIf LCase(vVals(i, 1)) Like "*day*" Or LCase(vVals(i, 1)) Like "*daily*" Then
            sTyp = "daily"
        End If
        If CDate(vDatesUsed(i, 1)) < Date Then 'last used sometime before today
            For dDayInc = CDate(vDatesUsed(i, 1)) + 1 To Date
                If Not bNewPer Then 'see if this dDayInc was the cross into a new period
                    If sTyp = "daily" Then
                        bNewPer = True
                    ElseIf sTyp = "week" And WeekdayName(Weekday(dDayInc)) = "Sunday" Then 'bnewper is true
                        bNewPer = True
                    ElseIf sTyp = "month" And Day(dDayInc) = vVals(i, 2) Then 'bnewper is true
                        bNewPer = True
                    ElseIf sTyp = "quarter" And Day(dDayInc) = 1 Then 'see if month is a quarter-start
                        sMonth = StrConv(vVals(i, 2), vbProperCase)
                        If LCase(sMonth) Like "*-*" Then sMonth = "January"
                        iMonth = Month(sMonth & " 1, 2000")
                        Do While iMonth > 3 'go to beginning of year
                            iMonth = iMonth - 3
                        Loop
                        If (Month(dDayInc) = iMonth Or Month(dDayInc) = iMonth + 3 Or Month(dDayInc) = iMonth + 6 Or Month(dDayInc) = iMonth + 9) Then
                            bNewPer = True
                            bAdd = True
                        End If
                    End If
                End If
                If bNewPer And (StrConv(vVals(i, 2), vbProperCase) = WeekdayName(Weekday(dDayInc)) Or vVals(i, 2) = Day(dDayInc) Or vVals(i, 2) = "-") Then 'if bNewPer and (Day name match OR day # match)
                    bAdd = True 'needs to be added
                    Exit For
                End If
            Next
            If bAdd Then 'add to output arrays
                ReDim Preserve sTypes(iInd)
                ReDim Preserve sDescs(iInd)
                ReDim Preserve sConts(iInd)
                ReDim Preserve sPrtis(iInd)
                ReDim Preserve sEases(iInd)
                ReDim Preserve dDues(iInd)
                sTypes(iInd) = vVals(i, 3)
                sDescs(iInd) = vVals(i, 4)
                sConts(iInd) = vVals(i, 5)
                sPrtis(iInd) = vVals(i, 6)
                sEases(iInd) = vVals(i, 7)
                If IsNumeric(vVals(i, 8)) Then 'due date exists
                    dDues(iInd) = Date + vVals(i, 8)
                Else 'no due date
                    dDues(iInd) = CDate(0)
                End If
                iInd = iInd + 1
                vDatesUsed(i, 1) = Date 'update dused
            End If
        End If
erradd:
        If Err.Number <> 0 Then
            ReDim Preserve sErrs(iErr)
            sErrs(iErr) = Chr(149) & " Row " & i & ": Err #" & Err.Number & ", " & Err.Description
            iErr = iErr + 1
            Err.Clear
        End If
    Next
    
    If iInd > 0 Then 'iterate & add to main list
        ReDim vVals(UBound(sTypes), 5)
        For i = 0 To UBound(sTypes)
            vVals(i, 0) = sTypes(i)
            vVals(i, 1) = sDescs(i)
            vVals(i, 2) = sConts(i)
            vVals(i, 3) = sPrtis(i)
            vVals(i, 4) = sEases(i)
            vVals(i, 5) = "-"
        Next
        With wsToDo
            iFirstRow = .Range("A2").End(xlDown).Row + 1 'find location on sheet1
            Do While wsToDo.Range("A" & iFirstRow).Value > 0
                iFirstRow = iFirstRow + 1
            Loop
            .Range("A" & iFirstRow & ":F" & iFirstRow + UBound(sTypes)).Value = vVals 'put in vvals
            iColDates = WorksheetFunction.Match("*due*", .Range("1:1"), 0)
            .Range(.Cells(iFirstRow, iColDates), .Cells(iFirstRow + UBound(sTypes), iColDates)).Value = Application.Transpose(dDues) 'put in dates
            For i = iFirstRow To iFirstRow + UBound(sTypes) 'change 0 dates
                If .Cells(i, iColDates).Value = 0 Then .Cells(i, iColDates).Value = "-"
            Next
            .Range(.Cells(iFirstRow, iColDates - 1), .Cells(iFirstRow + UBound(sTypes), iColDates - 1)).Value = Date 'put in dates
        End With
    End If
    
    If iErr > 0 Then 'at least one error occurred
        sError = "The following rows (on 'Recurring') resulted in errors:" & vbCrLf & vbCrLf & Join(sErrs, vbCrLf)
        MsgBox sError
    End If
    
    wsRecur.Range(sCol & "1:" & sCol & iLastRow).Value = vDatesUsed 'update dates used column
    wsRecur.Cells(2, iColUpdate).Value = Date 'update date
    Application.EnableEvents = True
End Function

Function ColorCategories()  'highlights column A of cell's row, by category
    Dim i As Integer, j As Integer, iRow As Integer, iColor As Integer, iColDone As Integer
    Dim oDictGood As New Dictionary, oDictCols As New Dictionary, oDictBlank As New Dictionary
    Dim lColors(15) As Long, lColor As Long
    Dim sCategory As String, sColDone As String
    Dim bDone As Boolean
    Dim rngValsA As Range
    Dim vValsA As Variant, vVar As Variant, vRows As Variant
    
    lColors(0) = 15773696 'blue
    lColors(1) = 16751103 'magenta
    lColors(2) = 5287936 'green
    lColors(3) = 12566463 'grey
    lColors(4) = 8487423 'red
    lColors(5) = 65535 'yellow
    lColors(6) = 16764108 'purple
    lColors(7) = 16762245 'aqua
    lColors(8) = 13434777 'seafoam
    lColors(9) = 9359529 'olive
    lColors(10) = 8761047 'brown
    lColors(11) = 16764159 'pink
    lColors(12) = 6750156 'lime
    lColors(13) = 8696052 'rust
    lColors(14) = 16777164 'baby
    lColors(15) = 49407 'orange

    'Set oDictGood = CreateObject("Scripting.Dictionary")
    oDictGood.CompareMode = vbTextCompare 'case insensitive
    'Set oDictBlank = CreateObject("Scripting.Dictionary")
    oDictBlank.CompareMode = vbTextCompare 'case insensitive
    'Set oDictCols = CreateObject("Scripting.Dictionary")
    
    With Worksheets(1)

'        iRow = 2
'        bDone = False
'        Do While .Cells(iRow, 1).Value > 0 'for each cell in A
'            If .Cells(iRow, 1).Interior.Color = xlNone Then 'needs color
'                bDone = False
'                Exit Do
'            End If
'            iRow = iRow + 1
'        Loop
'
'        If bDone Then Exit Function 'no blanks

        iRow = 2
        Do While .Cells(iRow, 1).Value > 0 'for each cell in A
            If Not Rows(iRow).Hidden Then
                sCategory = .Cells(iRow, 1).Value 'find value of task category
                If .Cells(iRow, 1).Interior.Color = 16777215 Then  'blank background
                    If oDictBlank.Exists(sCategory) Then 'append the row #
                        oDictBlank(sCategory) = oDictBlank(sCategory) & "," & Str(iRow)
                    Else 'create key
                        oDictBlank.Add sCategory, Str(iRow)
                    End If
                ElseIf Not oDictGood.Exists(sCategory) Then 'not in good dict yet --> add w/ color
                    lColor = .Cells(iRow, 1).Interior.Color
                    If lColor <> 13551615 And lColor <> 13561798 And lColor <> 10284031 And _
                    lColor <> 4434687 And lColor <> 255 Then 'not done/aband/impend
                        oDictGood.Add sCategory, .Cells(iRow, 1).Interior.Color
                        If Not oDictCols.Exists(.Cells(iRow, 1).Interior.Color) Then oDictCols.Add .Cells(iRow, 1).Interior.Color, 0
                    End If
                End If
            End If
            iRow = iRow + 1
        Loop
        
        'dict checking
'        For i = 0 To oDictGood.Count - 1
'            Debug.Print oDictGood.Keys()(i), oDictGood.Items()(i)
'        Next i
        
        i = 0
        For Each vVar In oDictBlank.Keys
            vRows = Split(oDictBlank(vVar), ",") 'rows with the same blank category
            lColor = 0
            If oDictGood.Exists(vVar) Then 'color all item row vals
                lColor = oDictGood(vVar)
            Else 'go through lcols(i) to find next col -> use that
                For j = 0 To UBound(lColors)
                    If Not oDictCols.Exists(lColors(j)) Then
                        lColor = lColors(j)
                        Exit For
                    End If
                Next
            End If
            For j = 0 To UBound(vRows)
                .Cells(vRows(j), 1).Interior.Color = lColor
            Next
            If Not oDictCols.Exists(lColor) Then oDictCols.Add lColor, 0
        Next
        
    End With
End Function

Function AddItem(sType As String, sDesc As String, sCont As String, sPri As String, sEase As String, sWait As String, dDue As Date)
    Application.EnableEvents = False
    Dim iRow As Integer
    iRow = ThisWorkbook.Worksheets("To-do").Range("A10000").End(xlUp).Row + 1
    With ThisWorkbook.Worksheets("To-do")
        .Range("A" & iRow).Value = sType
        .Range("B" & iRow).Value = sDesc
        .Range("C" & iRow).Value = sCont
        .Range("D" & iRow).Value = sPri
        .Range("E" & iRow).Value = sEase
        .Range("F" & iRow).Value = sWait
        .Range("G" & iRow).Value = Date
        .Range("H" & iRow).Value = dDue
    End With
    Application.EnableEvents = True
End Function

Function KeywordFunc(ByVal Target) As Boolean
'highlights all cells in column B of sheet 1 with Target cell's value + wildcards
    Dim sKeyword As String
    
    sKeyword = LCase(Target.Value)
    
    If sKeyword = "keyword" Or sKeyword = "" Then
        Target.Font.Italic = True
        Target.Font.Color = 8421504
        Application.EnableEvents = False
        Target.Value = "Keyword"
        KeywordFunc = False
        Exit Function
    End If
    
    KeywordFunc = True
    
    Dim iRow As Integer
    Dim lColorBG As Long
    
    lColorBG = 65535
    
    iRow = 2
    With ThisWorkbook.Worksheets(1)
        Do While .Range("A" & iRow).Value > 0
            If Not .Rows(iRow).Hidden Then 'only visible rows
                If LCase(.Range("B" & iRow).Value) Like "*" & sKeyword & "*" Then 'highlight
                    .Range("B" & iRow).Interior.Color = lColorBG
                    .Range("B" & iRow).Font.Color = 0
                End If
            End If
            iRow = iRow + 1
        Loop
    End With
    
End Function
