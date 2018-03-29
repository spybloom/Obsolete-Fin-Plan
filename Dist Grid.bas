Attribute VB_Name = "Module1"
Sub DistGrid()
Attribute DistGrid.VB_ProcData.VB_Invoke_Func = "A\n14"
'
' DistGrid Macro
'
' Keyboard Shortcut: Ctrl+Shift+A
'
' This macro combines the dist and gridsort macros to run everything needed for the third sheet and grid
'
' This macro doesn't work with the Radeke's kids - all their portfolios are on one sheet, so it will always
'   set the portfolio total to be their parents' total, A1 on the grid would always be "Kirk & Susan Radeke"
'   and the total equities for the kids would always be compared with their parents.
'
' Most of the dist part doesn't work if the dist page is separated by years
'   Would need to alter CellStart formulas (CellStart.Offset(-x, ...), formulas in YTD lines,
'       and total return cells, if applicable
'   Peplinski - No space between years, years are labeled "Overall", actual overall line needs most cells to change each update
'       Can change each yearly "Overall" to be "Overall ", but the added line would need to be CellStart.Offset(-4, ...
'       They also have total and annual return cells at the bottom. These would need to change with each update.
'   Zehren - Spaces between years, years are labeled "Total". Added line would need to be CellStart.Offset(-3, ...
'   Osburn - Same as Zehren. Also has annual yield per year below overall line.
'
' Hutter has two profile pages for Robert and Alice, this macro doesn't work when pulling portfolio present value or setting the dates
'
' To be implemented:
' For first updates after the initial setup, check if the number of dates between "Overall" and "Date" is one.
'   Add formulas to add Overall Previous Value
' For the grid, compare the number of funds in a sector range (SortStart to SortEnd) and how many rows are in the range.
'   If the number of funds is >4 and the difference is >1 delete the extra rows.
'   If the number of funds is <4 delete extra rows until the difference is 4
'
' 11/22/17 - Switched TurnOn/Off with StateToggle
' 11/14/17 - Cleaned up Goto's, added TurnOn/Off and AddError
' 5/12/17 - Created macro
'
    Dim TodayDate As Date
    Dim CellStart As Range
    Dim djia As Range
    Dim Contribution As Variant
    Dim Withdrawal As Variant
    Dim Distribution As Variant
    Dim DJInput As Variant
    Dim SPInput As Variant
    Dim s1 As Range
    Dim s2 As Variant
    Dim s3 As Variant
    Dim i As Integer
    Dim j As Integer
    Dim ChangeBox As Range
    Dim k As Integer
    Dim Portfolio As Worksheet
    Dim a As Integer
    Dim Grid As Worksheet
    Dim GridParts() As String
    Dim b As Integer
    Dim DayPart As String
    Dim MonthPart As String
    Dim YearPart As String
    Dim ClientName As Range
    Dim GridTotal As Long
    Dim PortTotal As Variant
    Dim Dist As Worksheet
    Dim PortEquityTotal As Long
    
    On Error GoTo BackOn
    
    Set Dist = ActiveSheet
    
    StateToggle "Off"
    
    'Check for Overall and DJIA
    If Dist.Range("A:A").Find("Overall", After:=Range("A1"), LookAt:=xlPart) Is Nothing Then
        AddError "Macro has been halted. Please put ""Overall"" on last line", True
    End If
    If Dist.Cells.Find("DJIA", After:=Cells.Range("A1"), LookIn:=xlValues) Is Nothing Then
        AddError "Macro has been halted. Please put ""DJIA"" next to the Dow Jones Index number.", True
    Else
        Set djia = Dist.Cells.Find("DJIA", After:=Cells.Range("A1"), LookIn:=xlValues)
    End If
    
    'Set Grid to be grid tab
    Set Grid = SetSheet("Grid")
    
    If Grid Is Nothing Then
        AddError "Macro has been halted; grid tab does not contain ""Grid"". Please revise and rerun macro.", True
    End If
    
    'Set D2 to equal portfolio total from portfolio tab
    Set Portfolio = SetSheet("Portfolio")
    
    If Portfolio Is Nothing Then
        AddError "Macro has been halted; portfolio tab does not contain ""Portfolio"". Please revise and" _
            & "rerun macro.", True
    End If
    
    'Set dates
    TodayDate = Date - 1
    DayPart = format(Day(TodayDate), "00")
    MonthPart = format(Month(TodayDate), "00")
    YearPart = Year(TodayDate)
    Dist.PageSetup.RightHeader = TodayDate
    
    If Dist.Rows(1).Find("Date", After:=Range("A1")) Is Nothing Then
        AddError " Date wasn't found on first row. This date hasn't been updated.", False
    Else
        Dist.Rows(1).Find("Date").Offset(1, 0).Formula = TodayDate
    End If
    
    If Portfolio.UsedRange.Find("Total Investments:") Is Nothing Then
        AddError " ""Total Investments:"" not found on portfolio tab. Present value will need" _
            & "to be checked manually.", False
    Else
        PortTotal = Portfolio.UsedRange.Find("Total Investments:").Offset(0, 2).Address
        Dist.Range("D2").Formula = "='" & Portfolio.Name & "'!" & PortTotal
    End If
    
    'Set column widths
    With Portfolio
        .Columns("A").ColumnWidth = 32
        .Columns("B").ColumnWidth = 12
        .Columns("C").ColumnWidth = 8
        .Columns("D").ColumnWidth = 12
        .Columns("E").ColumnWidth = 8.71
        .Columns("F").ColumnWidth = 8
        .Columns("G").ColumnWidth = 12
        .Columns("H").ColumnWidth = 8.71
        .Columns("I").ColumnWidth = 7
        .Columns("J").ColumnWidth = 8.43
        .Columns("K").AutoFit
    End With
    
    'Change date of portfolio and grid tabs using client's name from A1 of portfolio tab
        Dim IncPercent As Range
    Dim EqPercent As Range
    Set ClientName = Portfolio.Range("A1")
        
    'Check if A1 is yellow or red and for specific client conditions
    If ClientName.Interior.ColorIndex = 6 Or ClientName.Interior.ColorIndex = 3 Then
        Set ClientName = ClientName.Offset(1, 0)
        Set IncPercent = Portfolio.Range("E5")
        Set EqPercent = Portfolio.Range("H5")
    Else
        Set IncPercent = Portfolio.Range("E4")
        Set EqPercent = Portfolio.Range("H4")
    End If
    
    IncPercent = "%"
    EqPercent = "%"
    IncPercent.HorizontalAlignment = xlCenter
    EqPercent.HorizontalAlignment = xlCenter
        
    If ClientName = "Dan Bucholtz Trust" Then
        Grid.Range("A1").Formula = ClientName.Value & " - " & MonthPart & "/" & DayPart & "/" & YearPart
        ClientName.Offset(2, 0).Formula = "Portfolio Analysis - " & MonthPart & "/" & DayPart & "/" & YearPart
    ElseIf ClientName = "Tad (Chip) & Karen Bircher" Then
        Grid.Range("A1").Formula = "Chip & Karen Bircher - " & MonthPart & "/" & DayPart & "/" & YearPart
        ClientName.Offset(1, 0).Formula = "Portfolio Analysis - " & MonthPart & "/" & DayPart & "/" & YearPart
    Else
        Grid.Range("A1").Formula = ClientName.Value & " - " & MonthPart & "/" & DayPart & "/" & YearPart
        ClientName.Offset(1, 0).Formula = "Portfolio Analysis - " & MonthPart & "/" & DayPart & "/" & YearPart
    End If
    
    'Check fixed income and equity boxes are summing correctly on portfolio tab
    Dim Inputs() As String
    Dim TextFile As Integer
    Dim Markets As String
    Dim MarketsInput() As String
    Dim Fixed As Range
    Dim Equity As Range
    Dim SumRange As Range
    
    Set Fixed = Portfolio.UsedRange.Find("Category Totals:").Offset(0, 2)
    Set Equity = Portfolio.UsedRange.Find("Category Totals:").Offset(0, 5)
    Set SumRange = Portfolio.Range(Portfolio.Range("A" & ClientName.Offset(4, 0).Row).Offset(0, 3), Fixed.Offset(-1, 0))
    Fixed.Formula = "=SUM(" & SumRange.Address & ")"
    
    Set SumRange = Portfolio.Range(Portfolio.Range("A" & ClientName.Offset(4, 0).Row).Offset(0, 6), Equity.Offset(-1, 0))
    Equity.Formula = "=SUM(" & SumRange.Address & ")"
    
    'Input Dow Jones and S&P from Markets file
    TextFile = FreeFile
    Markets = "Z:\YungwirthSteve\Macros\Documents\Markets.txt"
    
    Open Markets For Input As TextFile
        MarketsInput = Split(Input(LOF(TextFile), TextFile), " ")
    Close TextFile
    
    DJInput = MarketsInput(0)
    SPInput = MarketsInput(1)
    
    djia.Offset(0, 1) = DJInput
    djia.Offset(1, 1) = SPInput
        
    Set CellStart = Dist.UsedRange.Find("Overall", After:=Range("A1"))
    CellStart.Offset(-1, 0).EntireRow.Insert
    
    'Take inputs
    Inputs() = Split(InputBox("Enter contributions, withdrawals, and distributions, separated by spaces"))
    ReDim Preserve Inputs(0 To 2)
    
    For m = 0 To UBound(Inputs)
        If Inputs(m) = vbNullString Then 'If text box is completely empty
            Inputs(m) = 0
        End If
    Next m

    Contribution = Inputs(0)
    Withdrawal = Inputs(1)
    Distribution = Inputs(2)
    
    'Third page columns - Date and Present Value
    Dim PresentValue As Single
    Dim n As Integer
    
    For n = 0 To 12
        Select Case n
            Case 0
                CellStart.Offset(-2, n).Formula = TodayDate
            Case 1
                CellStart.Offset(-2, n).FormulaR1C1 = "=R[-1]C[5]" 'This cell reference doesn't work if dist page is separated by years
            Case 2
                CellStart.Offset(-2, n).Value = Contribution
            Case 3
                CellStart.Offset(-2, n).Value = Withdrawal
            Case 4
                CellStart.Offset(-2, n).Value = Distribution
            Case 5
                CellStart.Offset(-2, n).FormulaR1C1 = "=R[0]C[-4]+R[0]C[-3]-R[0]C[-2]-R[0]C[-1]"
            Case 6
                Calculate
                PresentValue = CellStart.Offset(0, n).Value2
                CellStart.Offset(-2, n).Value2 = PresentValue
                CellStart.Offset(-2, n).NumberFormat = "_($* #,##0_);_($* (#,##0);_($* ""-""_);_(@_)"
            Case 7
                CellStart.Offset(-2, n).FormulaR1C1 = "=R[0]C[-1]-R[0]C[-2]"
            Case 8
                CellStart.Offset(-2, n).FormulaR1C1 = "=R[0]C[-1]/R[0]C[-3]"
            Case 9
                CellStart.Offset(-2, n).FormulaR1C1 = "=R[0]C[2]/R[-1]C[2]-1" 'Cell reference doesn't work if dist page is separated by years
            Case 10
                CellStart.Offset(-2, n).FormulaR1C1 = "=R[0]C[-2]-R[0]C[-1]"
            Case 11
                CellStart.Offset(-2, n).Formula = djia.Offset(1, 1).Formula
            Case 12
                CellStart.Offset(-2, n).FormulaR1C1 = "=R[-1]C[0]*(1+R[0]C[-4])" 'Cell reference doesn't work if dist page is separated by years
        End Select
    Next n
    
    'Overall Change & S&P 500
    For i = 1 To 1000
        If CellStart.Offset(0, 8).Formula = "=M" & i & "-1" Then
            CellStart.Offset(0, 8).FormulaR1C1 = "=R[-2]C[4]-1"
        End If
    Next i
    
    Set s1 = Dist.Columns(1).Find("Date")
    If Not s1 Is Nothing Then
        s2 = s1.Offset(1, 11).Row
        s3 = s1.Offset(1, 11).Column
        CellStart.Offset(0, 9).FormulaR1C1 = "=R[-2]C[2]/R" & s2 & "C" & s3 & "-1"
    Else
        AddError " ""Date"" wasn't found on the top row of the client's performance numbers. The " _
            & "S&P 500 performance in the ""Overall"" row will need to be changed manually.", False
    End If
    
    'Diagnostics
    Calculate
    Set ChangeBox = Range("A1:K20").Find("Net")
    If ChangeBox.Offset(1, 0).Value <> "Change" Then
        ChangeBox.Offset(-1, 0).Value = ChangeBox.Value
        ChangeBox = "Change"
        Set ChangeBox = ChangeBox.Offset(1, 0)
    Else: Set ChangeBox = ChangeBox.Offset(2, 0)
    End If
    
    If ChangeBox.Value > CellStart.Offset(-2, 6).Value * 0.1 Then
        AddError " Macro completed. Net change box may be too high. Check numbers and re-run " _
            & "macro if necessary.", False
    End If

    If ChangeBox.Value < CellStart.Offset(-2, 6).Value * -0.1 Then
        AddError " Macro completed. Net change box may be too low. Check numbers and re-run " _
            & "macro if necessary.", False
    End If
    
    For j = 0 To 12
        If IsEmpty(CellStart.Offset(-2, j)) = True Or CellStart.Offset(-2, j).Value = "" Then
            AddError " Macro completed with empty cells, manually check output.", False
            Exit For
        End If
    Next j

    'Net change box
    If CellStart.Offset(-2, 10).Value > 0.005 And CellStart.Offset(-2, 9) > 0 Then
        AddError " Macro completed. Portfolio performed higher than S&P 500. Check numbers and re-run " _
            & "macro if necessary.", False
    ElseIf CellStart.Offset(-2, 10).Value < 0 And CellStart.Offset(-2, 9) < 0 Then
        AddError " Macro completed. Portfolio performed lower than S&P 500. Check numbers and re-run " _
            & "macro if necessary.", False
    End If
    
    'Set print area
    Dim LastRow As Integer
    Dim TotalRows As Integer
    
    TotalRows = Range("A1", Range("A" & Dist.UsedRange.Rows(Dist.UsedRange.Rows.Count).Row).End(xlUp)).Rows.Count
    
    If TotalRows > 80 Then
        With Dist.PageSetup
            .Orientation = xlPortrait
            .BottomMargin = Application.InchesToPoints(1)
            .Zoom = False
            .FitToPagesTall = False
            .FitToPagesWide = 1
        End With
    Else
        With Dist.PageSetup
            .Orientation = xlPortrait
            .BottomMargin = Application.InchesToPoints(1)
            .Zoom = False
            .FitToPagesTall = 1
            .FitToPagesWide = 1
        End With
    End If
    
    'Sort Grid
    Dim SortStart As Range
    Dim SortEnd As Range
    Dim SectTotal As Range
    Dim GridRowSize As Integer
    Dim GridArea As Range
    
    'Set print area to be 1 page wide and tall
    With Grid.PageSetup
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = 1
        If .TopMargin <> Application.InchesToPoints(0.25) And .CenterVertically = False Then
            .LeftMargin = Application.InchesToPoints(0.25)
            .RightMargin = Application.InchesToPoints(0.25)
            .TopMargin = Application.InchesToPoints(0.25)
            .BottomMargin = Application.InchesToPoints(0.25)
            .HeaderMargin = Application.InchesToPoints(0.25)
            .FooterMargin = Application.InchesToPoints(0.25)
            .CenterVertically = True
            .CenterHorizontally = True
        End If
    End With
    
    With Grid
        .Columns("A").ColumnWidth = 8.43
        If .Columns("B").ColumnWidth < 22 Then
            .Columns("B").ColumnWidth = 22
        End If
        .Columns("C").ColumnWidth = 11.29
        .Columns("D").ColumnWidth = 5
        
        .Columns("E").ColumnWidth = 8.43
        If .Columns("F").ColumnWidth < 27 Then
            .Columns("F").ColumnWidth = 27
        End If
        .Columns("G").ColumnWidth = 11.29
        .Columns("H").ColumnWidth = 5
        
        .Columns("I").ColumnWidth = 8.43
        If .Columns("J").ColumnWidth < 23 Then
            .Columns("J").ColumnWidth = 23
        End If
        .Columns("K").ColumnWidth = 11.29
        .Columns("L").ColumnWidth = 5
    End With
    
    'Sort the grid alphabetically
    GridParts = Split("Large Value,Large Blend,Large Growth,Medium Value,Medium Blend,Medium Growth," _
        & "Small Value,Small Blend,Small Growth,Specialty Holdings", ",")
    
    For b = 0 To UBound(GridParts)
        Grid.Activate
        If Grid.UsedRange.Find(GridParts(b), LookAt:=xlPart) Is Nothing _
            And (b = 1 And Grid.UsedRange.Find("Foreign") Is Nothing Or b <> 1) Then
                AddError """" & GridParts(b) & """ wasn't found. This category wasn't sorted.", False
        Else
            If Not Grid.UsedRange.Find("Foreign") Is Nothing And GridParts(b) = "Large Blend" Then
                Set SortStart = Grid.UsedRange.Find("Foreign").Offset(1, 0)
            Else
                Set SortStart = Grid.UsedRange.Find(GridParts(b)).Offset(1, 0)
            End If
            
            If SortStart.Offset(-1, 0) = "Large Blend" Then
                SortStart.Offset(-1, 0) = "Foreign"
            End If
            
            If Range(SortStart, SortStart.Offset(100, 0)).Find("Sector Total", After:=SortStart, _
                LookAt:=xlPart) Is Nothing Then
                    Set SectTotal = Range(SortStart, SortStart.Offset(100, 0)).Find("Total", After:=SortStart, _
                        LookAt:=xlWhole)
            Else
                Set SectTotal = Range(SortStart, SortStart.Offset(100, 0)).Find("Sector Total", _
                    After:=SortStart, LookAt:=xlPart)
            End If
            Set SortEnd = SectTotal.Offset(-1, 3)
            SectTotal.Value = "Sector Total"
            SectTotal.IndentLevel = 1
            Set GridArea = Range(SortStart.Offset(-1, 0), SortEnd.Offset(1, 0))
            
            GridArea.BorderAround LineStyle:=xlContinuous, Weight:=xlThin, ColorIndex:=xlColorIndexAutomatic
            
            GridRowSize = GridArea.Rows.Count - 2
            SectTotal.Offset(0, 2).FormulaR1C1 = "=SUM(R[" & -1 * GridRowSize & "]C:R[-1]C)"
            SectTotal.Offset(0, 3).FormulaR1C1 = "=SUM(R[" & -1 * GridRowSize & "]C:R[-1]C)"

            With Grid.Sort
                .SortFields.Clear
                .SortFields.Add Key:=SortStart, SortOn:=xlSortOnValues, Order:=xlAscending, _
                    DataOption:=xlSortNormal
                .SetRange Range(SortStart, SortEnd)
                .Header = xlNo
                .MatchCase = False
                .Orientation = xlTopToBottom
                .SortMethod = xlPinYin
                .Apply
            End With
'            If b = 2 Or b = 5 Or b = 8 Then
'                Delete rows at the end if all 3 sectors in the row are empty
'            End If
        End If
    Next b
        
    'Check to make sure grid matches portfolio tab
    Calculate
    
    If Not Grid.UsedRange.Find("Total", LookIn:=xlValues, LookAt:=xlWhole) Is Nothing Then
        If Application.IsNA(Grid.UsedRange.Find("Total", LookIn:=xlValues, LookAt:=xlWhole).Offset(0, 2).Value) _
            Or IsError(Grid.UsedRange.Find("Total", LookIn:=xlValues, LookAt:=xlWhole).Offset(0, 2).Value) Then
            AddError "One or more of the grid sector sums results in an error. Please check and rerun macro.", False
        Else
            GridTotal = Grid.UsedRange.Find("Total", LookIn:=xlValues, LookAt:=xlWhole).Offset(0, 2).Value
            PortTotal = Portfolio.UsedRange.Find("Category Totals:", LookIn:=xlValues, _
                LookAt:=xlWhole).Offset(0, 5).Value
            If GridTotal > 0 And PortTotal > 0 And GridTotal <> PortTotal Then
                AddError "Total of equities doesn't match between grid and portfolio tabs. Grid has $" _
                    & GridTotal & ", portfolio has $" & PortTotal & " - a difference of $" & _
                    GridTotal - PortTotal & ".", False
            End If
        End If
    Else
        AddError "Equity totals between grid and portfolio tabs could not be compared - Grid total not " _
            & "named ""Total"" and/or portfolio total not named ""Category Totals:"".", False
    End If
    
    Dist.Activate
    StateToggle "On"
    
    AddError "", True
    
    Portfolio.PrintOut Preview:=True
    Dist.PrintOut Preview:=True
    Grid.PrintOut Preview:=True
    
    Exit Sub
    
BackOn:
    StateToggle ("On")
    MsgBox ("Macro ended prematurely due to error in execution.")
End Sub
Sub StateToggle(OnOrOff As String)
    Static OriginalScreen As Boolean
    Static OriginalEvents As Boolean
    Static OriginalStatus As Boolean
    Static OriginalCalc As XlCalculation
    Dim Reset As Long
    
    If OnOrOff = "Off" Then
        OriginalScreen = Application.ScreenUpdating
        OriginalEvents = Application.EnableEvents
        OriginalStatus = Application.DisplayStatusBar
        OriginalCalc = Application.Calculation
        
        Application.ScreenUpdating = False
        Application.EnableEvents = False
        Application.DisplayStatusBar = False
        Application.Calculation = xlCalculationManual
    ElseIf OnOrOff = "On" Then
        Application.ScreenUpdating = OriginalScreen
        Application.EnableEvents = OriginalEvents
        Application.DisplayStatusBar = OriginalStatus
        Application.Calculation = OriginalCalc
        Reset = ActiveSheet.UsedRange.Rows.Count
    End If
End Sub
Function AddError(Error As String, Display As Boolean) As Integer
    Static ErrorMessage As String
    
    If Error <> "" Then
        ErrorMessage = ErrorMessage & Chr(149) & Error & vbNewLine
        If Display = True Then
            MsgBox (ErrorMessage)
            ErrorMessage = ""
            StateToggle "On"
            End
        End If
    ElseIf Error = "" And ErrorMessage <> "" Then
        MsgBox (ErrorMessage)
        ErrorMessage = ""
    End If
End Function
Function SetSheet(TargetSheet As String) As Worksheet
    Dim i As Integer
    
    For i = 1 To Worksheets.Count
        If InStr(UCase(Worksheets(i).Name), UCase(TargetSheet)) > 0 Then
            Set SetSheet = Worksheets(i)
            Exit Function
        End If
    Next i
End Function
Public Sub UserForm_Initialize()

End Sub
