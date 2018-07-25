Attribute VB_Name = "Dist"
Option Explicit
Private Portfolio As Worksheet
Private Grid As Worksheet
Private Dist As Worksheet
Private PortTotal As Long
Sub DistGrid()
Attribute DistGrid.VB_ProcData.VB_Invoke_Func = "A\n14"
'
' Keyboard Shortcut: Ctrl+Shift+A
'
' This macro combines the dist and gridsort macros to run everything needed for the third sheet and grid
'
' This macro doesn't work with with files where the portfolios/grids/dists are split
'   - The Radeke's kids - all their portfolios are on one sheet, so it will always set the portfolio total
'       to be their parents' total, A1 on the grid would always be "Kirk & Susan Radeke" and the total
'       equities for the kids would always be compared with their parents.
'   - Hutter has two portfolio pages for Robert and Alice, this macro doesn't work when pulling portfolio
'       present value or setting the dates
'
' Most of the dist part doesn't work if the dist page is separated by years
'   Would need to alter CellStart formulas (CellStart.Offset(-x, ...), formulas in YTD lines,
'       and total return cells, if applicable
'   Peplinski - No space between years, years are labeled "Total", actual overall line needs most cells to
'       change each update. Can change each yearly "Overall" to be "Overall ", but the added line would need
'       to be CellStart.Offset(-4, ... They also have total and annual return cells at the bottom. These
'       would need to change with each update.
'   Zehren - Spaces between years, years are labeled "Total". Added line would need to be CellStart.Offset(-3, ...
'   Osburn - Same as Zehren. Also has annual yield per year below overall line.
'   Peplinski = Same as Osburn
'
' To be implemented:
' For the grid, compare the number of funds in a sector range (SortStart to SortEnd) and how many rows are in the range.
'   If the number of funds is >4 and the difference is >1 delete the extra rows.
'   If the number of funds is <4 delete extra rows until the difference is 4
'
' 5/12/17 - Created macro
' 11/14/17 - Cleaned up Goto's, added TurnOn/Off and AddError
' 11/22/17 - Switched TurnOn/Off with StateToggle
' 5/25/18 - Started cleaning up, added link to StateToggle.UpdateScreen
' 6/22/18 - Still cleaning up. Deleted GridSort section and added a link to the GridSort macro itself
    
    On Error GoTo BackOn
    StateToggle.UpdateScreen "Off"
    
    'Set PA sheets
    Set Portfolio = SetSheet("Portfolio")
    Set Grid = SetSheet("Grid")
    Set Dist = SetSheet("Dist")
    
    'Check for Overall and DJIA
    Dim djia As Range
    
    Set djia = Dist.UsedRange.Find("DJIA", After:=Dist.Range("A1"), LookIn:=xlValues)
        
    If Dist.Range("A:A").Find("Overall", After:=Dist.Range("A1"), LookAt:=xlPart) Is Nothing Then
        AddError "Macro has been halted. Please put ""Overall"" on last line", True
    ElseIf djia Is Nothing Then
        AddError "Macro has been halted. Please put ""DJIA"" next to the Dow Jones Index number.", True
    End If
    
    'Take inputs
    Dim Contribution As Variant
    Dim Withdrawal As Variant
    Dim Distribution As Variant
    Dim LineOnly As Variant
    
    DistInputs_RowOption.Show vbModal
    
    Contribution = DistInputs_RowOption.TransIn
    Withdrawal = DistInputs_RowOption.TransOut
    Distribution = DistInputs_RowOption.Distributions
    LineOnly = DistInputs_RowOption.DistRow
    
    If Contribution = "" Then
        Contribution = 0
    End If
    
    If Withdrawal = "" Then
        Withdrawal = 0
    End If
    
    If Distribution = "" Then
        Distribution = 0
    End If
    
    If LineOnly = "" Then
        LineOnly = 0
    End If
    
    Unload DistInputs_RowOption
    
    'Set dates
    Dim TodayDate As Date
    Dim DayPart As String
    Dim MonthPart As String
    Dim YearPart As String
    Dim FirstDate As Range
    Dim NextDates As Range
    
    TodayDate = Date - 1
    DayPart = format(Day(TodayDate), "00")
    MonthPart = format(Month(TodayDate), "00")
    YearPart = Year(TodayDate)
    Dist.PageSetup.RightHeader = TodayDate
    
    If Dist.Rows(1).Find("Date", After:=Range("A1")) Is Nothing Then
        AddError " Date wasn't found on first row. This date hasn't been updated.", False
    Else
        Set FirstDate = Dist.Rows(1).Find("Date").Offset(1, 0)
        Set NextDates = FirstDate.Offset(1, 0)
        
        FirstDate.Formula = TodayDate
        
        Do While NextDates <> ""
            NextDates.Formula = TodayDate
            Set NextDates = NextDates.Offset(1, 0)
        Loop
    End If
    
    'Set D2 to equal portfolio total from portfolio tab
    Dim PortTotalLoc As String
    
    If Portfolio.UsedRange.Find("Total Investments:") Is Nothing Then
        AddError " ""Total Investments:"" not found on portfolio tab. Present value will need " _
            & "to be checked manually.", False
    Else
        PortTotalLoc = Portfolio.UsedRange.Find("Total Investments:").Offset(0, 2).Address
        PortTotal = Portfolio.UsedRange.Find("Total Investments:").Offset(0, 2).Value2
        Dist.Range("D2").Formula = "='" & Portfolio.Name & "'!" & PortTotalLoc
    End If
    
    'Check fixed income and equity boxes are summing correctly on portfolio tab
    Dim Fixed As Range
    Dim Equity As Range
    Dim FixedRange As Range
    Dim EqRange As Range
    Dim InvOrFund As Range
    Dim ClientName As Range
    
    Set InvOrFund = Portfolio.UsedRange.Find("Investment or Fund", After:=Portfolio.Range("A1"))
    Set ClientName = InvOrFund.Offset(-3, 0)
    Set Fixed = Portfolio.UsedRange.Find("Category Totals:").Offset(0, 2)
    Set Equity = Portfolio.UsedRange.Find("Category Totals:").Offset(0, 5)
    Set FixedRange = Portfolio.Range(Portfolio.Range("A" & ClientName.Offset(4, 0).Row).Offset(0, 3), Fixed.Offset(-1, 0))
    Set EqRange = Portfolio.Range(Portfolio.Range("A" & ClientName.Offset(4, 0).Row).Offset(0, 6), Equity.Offset(-1, 0))
    
    Fixed.Formula = "=SUM(" & FixedRange.Address & ")"
    Equity.Formula = "=SUM(" & EqRange.Address & ")"
    
    'Input Dow Jones and S&P from Markets file
    Dim TextFile As Integer
    Dim Markets As String
    Dim MarketsInput() As String
    Dim DJInput As Variant
    Dim SPInput As Variant
    Dim CellStart As Range
        
    TextFile = FreeFile
    Markets = "Z:\YungwirthSteve\Macros\Documents\Markets.txt"
    
    Open Markets For Input As TextFile
        MarketsInput = Split(Input(LOF(TextFile), TextFile), " ")
    Close TextFile
    
    DJInput = MarketsInput(0)
    SPInput = MarketsInput(1)
    
    djia.Offset(0, 1) = DJInput
    djia.Offset(1, 1) = SPInput
    
    'Third page columns - Date and Present Value
    Dim RowOffset As Integer
    
    If LineOnly = 1 Then
        Set CellStart = ActiveCell
        RowOffset = 1
        CellStart.Offset(RowOffset, 0).EntireRow.Insert
    Else
        Set CellStart = Dist.UsedRange.Find("Overall", After:=Range("A1"))
        RowOffset = -2
        
        If CellStart.Offset(RowOffset, 0).Value = vbNullString Then 'For first updates - Portfolio template
                                                                    'has two spaces between first date and Overall
            CellStart.Offset(RowOffset, 0).EntireRow.ClearContents
            CellStart.Offset(RowOffset + 2, 1).FormulaR1C1 = "=R[" & RowOffset & "]C[0]"
        Else
            CellStart.Offset(RowOffset + 1, 0).EntireRow.Insert
        End If
    End If
    
    Dim n As Integer
    Dim PresentValue As Double
    
    For n = 0 To 12
        Select Case n
            Case 0
                CellStart.Offset(RowOffset, n).Formula = TodayDate
                CellStart.Offset(RowOffset - 1, n).Copy
                CellStart.Offset(RowOffset, n).PasteSpecial xlPasteFormats
                Application.CutCopyMode = False
            Case 1
                CellStart.Offset(RowOffset, n).FormulaR1C1 = "=R[-1]C[5]" 'This cell reference doesn't work if dist page is separated by years
            Case 2
                CellStart.Offset(RowOffset, n).Value = Contribution
            Case 3
                CellStart.Offset(RowOffset, n).Value = Withdrawal
            Case 4
                CellStart.Offset(RowOffset, n).Value = Distribution
            Case 5
                CellStart.Offset(RowOffset, n).FormulaR1C1 = "=R[0]C[-4]+R[0]C[-3]-R[0]C[-2]-R[0]C[-1]"
            Case 6
                Calculate
                PresentValue = CellStart.Offset(RowOffset + 2, n).Value2
                CellStart.Offset(RowOffset, n).Value2 = PresentValue
                CellStart.Offset(RowOffset, n).NumberFormat = "_($* #,##0_);_($* (#,##0);_($* ""-""_);_(@_)"
            Case 7
                CellStart.Offset(RowOffset, n).FormulaR1C1 = "=R[0]C[-1]-R[0]C[-2]"
            Case 8
                CellStart.Offset(RowOffset, n).FormulaR1C1 = "=R[0]C[-1]/R[0]C[-3]"
            Case 9
                CellStart.Offset(RowOffset, n).FormulaR1C1 = "=R[0]C[2]/R[-1]C[2]-1" 'Cell reference doesn't work if dist page is separated by years
            Case 10
                CellStart.Offset(RowOffset, n).FormulaR1C1 = "=R[0]C[-2]-R[0]C[-1]"
            Case 11
                CellStart.Offset(RowOffset, n).Formula = djia.Offset(1, 1).Formula
            Case 12
                CellStart.Offset(RowOffset, n).FormulaR1C1 = "=R[-1]C[0]*(1+R[0]C[-4])" 'Cell reference doesn't work if dist page is separated by years
        End Select
    Next n
    
    'Overall Change & S&P 500
    Dim i As Integer
    Dim BottomDate As Range
    Dim FirstSPRow As Integer
    Dim FirstSPCol As Integer
    
    For i = 1 To 1000
        If CellStart.Offset(RowOffset + 2, 8).Formula = "=M" & i & "-1" Then
            CellStart.Offset(RowOffset + 2, 8).FormulaR1C1 = "=R[-2]C[4]-1"
        End If
    Next i
    
    Set BottomDate = Dist.Columns(1).Find("Date")
    If Not BottomDate Is Nothing Then
        FirstSPRow = BottomDate.Offset(1, 11).Row
        FirstSPCol = BottomDate.Offset(1, 11).Column
        CellStart.Offset(RowOffset + 2, 9).FormulaR1C1 = "=R[-2]C[2]/R" & FirstSPRow & "C" & FirstSPCol & "-1"
    Else
        AddError " ""Date"" wasn't found on the top row of the client's performance numbers. The " _
            & "S&P 500 performance in the ""Overall"" row will need to be changed manually.", False
    End If
    
    'Net change box
    Dim TopRange As Range
    Dim ChangeBox As Range
    Dim UpperBound As Variant
    Dim LowerBound As Variant
    
    Calculate
    Set TopRange = Dist.Range("A1:K20")
    Set ChangeBox = TopRange.Find("Net")
    
    UpperBound = PresentValue * 0.1 + Contribution
    LowerBound = PresentValue * -0.1 - Withdrawal - Distribution
    
    If ChangeBox Is Nothing Then
        Set ChangeBox = TopRange.Find("Change")
        
        If ChangeBox Is Nothing Then
            AddError " Macro completed. Net change box couldn't be found. Check numbers and " _
                & "re-run macro if necessary", False
        End If
    End If
    
    If Not ChangeBox Is Nothing Then
        If ChangeBox.Offset(1, 0).Value <> "Change" Then
            ChangeBox.Offset(-1, 0).Value = "Net"
            ChangeBox = "Change"
            Set ChangeBox = ChangeBox.Offset(1, 0)
        Else
            Set ChangeBox = ChangeBox.Offset(2, 0)
        End If
        
        If ChangeBox.Value > UpperBound Then
            AddError " Macro completed. Net change box may be too high. Check numbers and re-run " _
                & "macro if necessary.", False
        End If
    
        If ChangeBox.Value < LowerBound Then
            AddError " Macro completed. Net change box may be too low. Check numbers and re-run " _
                & "macro if necessary.", False
        End If
    End If
    
    'Diagnostics
    Dim j As Integer
    Dim Error As Boolean
    
    Error = False
    
    For j = 0 To 12
        If IsError(CellStart.Offset(RowOffset, j)) Then
            AddError " Macro completed with an error on the new line.", False
            Error = True
            Exit For
        End If
    Next j
    
    If Not Error Then
        If CellStart.Offset(RowOffset, 10).Value > 0.005 And CellStart.Offset(RowOffset, 9) > 0 Then
            AddError " Macro completed. Portfolio performed higher than S&P 500. Check numbers and re-run " _
                & "macro if necessary.", False
        ElseIf CellStart.Offset(RowOffset, 10).Value < 0 And CellStart.Offset(RowOffset, 9) < 0 Then
            AddError " Macro completed. Portfolio performed lower than S&P 500. Check numbers and re-run " _
                & "macro if necessary.", False
        End If
    End If
    
    'Set dist column widths - Portfolio and grid column widths are set later in Grid_Sort
    Dim DistCols() As Variant
    Dim Count As Integer
    
    DistCols = Array(9.29, 11.57, 11.57, 11.57, 11.57, 11.57, 11.57, 11.71, 11.43, 9.43, 8.43)
    
    With Dist
        For Count = 0 To UBound(DistCols)
            .Columns(Count + 1).ColumnWidth = DistCols(Count)
        Next Count
    End With
    
    If LineOnly <> 1 Then
        'Set print area
        Dim DistRows As Integer
        
        DistRows = Dist.Range("A1", Dist.Range("A" & Dist.UsedRange.Rows(Dist.UsedRange.Rows.Count).Row).End(xlUp)).Rows.Count
        
        With Dist.PageSetup
            .Orientation = xlPortrait
            .BottomMargin = Application.InchesToPoints(1)
            .Zoom = False
            .FitToPagesWide = 1
            
            If DistRows > 80 Then
                .FitToPagesTall = False
            Else
                .FitToPagesTall = 1
            End If
        End With
        
        If Dist.PageSetup.LeftHeader <> "Brad Weyers" Then
            Grid_Sort.GridSort
        Else
            AddError " Brad Weyers' portfolio and grid are abnormal, so they weren't checked.", False
        End If
        
        Portfolio.Activate
        
        StateToggle.UpdateScreen "On"
        
        AddError "", True
        
        Portfolio.PrintOut Preview:=True
        Dist.PrintOut Preview:=True
        Grid.PrintOut Preview:=True
    Else
        StateToggle.UpdateScreen "On"
        AddError "", True
    End If
    
    Exit Sub
    
BackOn:
    StateToggle.UpdateScreen ("On")
    MsgBox ("Macro ended prematurely due to error in execution.")
End Sub
Function AddError(Error As String, Display As Boolean) As Integer
    Static ErrorMessage As String
    
    If Error <> "" Then
        ErrorMessage = ErrorMessage & Chr(149) & Error & vbNewLine
        If Display = True Then
            MsgBox (ErrorMessage)
            ErrorMessage = ""
            StateToggle.UpdateScreen "On"
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
    
    If SetSheet Is Nothing Then
        AddError "Macro has been halted; " & TargetSheet & " tab does not contain """ & TargetSheet _
            & """. Please revise and rerun macro, or sort manually.", True
    End If
End Function
