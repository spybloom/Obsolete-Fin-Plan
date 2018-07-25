Attribute VB_Name = "Grid_Sort"
Option Explicit
Private Grid As Worksheet
Private Portfolio As Worksheet
Sub GridSort()
Attribute GridSort.VB_ProcData.VB_Invoke_Func = "F\n14"
'
' GridSort Macro
'
' This macro alphabetically sorts each section of the grid, sets print area to be 1 page, and changes
'   the date of the grid and portfolio pages
'
' Need to add a condition for Radeke - all their portfolios are on the one sheet, so A1 on the grid would
'   always be "Kirk & Susan Radeke" and the total equities for the kids would always be compared with
'   their parents.
'
' 12/8/17:  Added AddError and SetSheet.
'           Added code to add borders to each grid section and to sum up each section (Total and percent).
' 6/22/18:  Cleaned up
'
' Keyboard Shortcut: Ctrl+Shift+F
'
    On Error GoTo BackOn
    StateToggle.UpdateScreen "Off"
    
    'Set sheets
    Set Portfolio = SetSheet("Portfolio")
    Set Grid = SetSheet("Grid")
    
    FormatSheets 'Format column widths, font, print size
    
    SortSectors 'Sort each sector of the grid alphabetically
    
    NameAndDate 'Change date of portfolio and grid tabs using client's name from A1 of portfolio tab
    
    TotalCheck 'Check to make sure grid matches portfolio tab
    
    AddError vbNullString
    StateToggle.UpdateScreen "On"
    
    Exit Sub
    
BackOn:
    StateToggle.UpdateScreen ("On")
    MsgBox ("Macro ended prematurely due to error in execution.")
End Sub
Function AddError(Error As String, Optional Display As Boolean) As Integer
    Static ErrorMessage As String
    
    If Error <> vbNullString Then
        ErrorMessage = ErrorMessage & Chr(149) & " " & Error & vbNewLine
        If Display = True Then
            MsgBox (ErrorMessage)
            ErrorMessage = vbNullString
            StateToggle.UpdateScreen "On"
            End
        End If
    ElseIf Error = vbNullString And ErrorMessage <> vbNullString Then
        MsgBox (ErrorMessage)
        ErrorMessage = vbNullString
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
Sub SetCols()
    Dim PortCols() As Variant
    Dim GridCols() As Variant
    
    PortCols = Array(32, 12, 8, 12, 8.71, 8, 12, 8.71, 7, 8.43)
    GridCols = Array(8.43, 24, 11.29, 5, 8.43, 27, 11.29, 5, 8.43, 23, 11.29, 5)
    
    ColSize Portfolio, PortCols
    Portfolio.Columns("K").AutoFit
    ColSize Grid, GridCols
End Sub
Sub ColSize(TargetSheet As Worksheet, ColArr As Variant)
    Dim Count As Integer
    
    With TargetSheet
        For Count = 0 To UBound(ColArr)
            .Columns(Count + 1).ColumnWidth = ColArr(Count)
        Next Count
    End With
End Sub
Sub FormatSheets()
    SetCols
    With Grid
        With .UsedRange.Font
            .Name = "Arial"
            .FontStyle = "Regular"
            .Size = 10
        End With
        
        If .Range("A2") = "David Bucholtz, Trustee" Then
            .Range("A4").Font.Bold = True
        Else
            .Range("A3").Font.Bold = True
        End If
        
        With .PageSetup
            .Zoom = False
            .FitToPagesWide = 1
            .FitToPagesTall = 1
            .PrintErrors = xlPrintErrorsDisplayed
            .ScaleWithDocHeaderFooter = True
            .AlignMarginsHeaderFooter = True
            If .TopMargin <> Application.InchesToPoints(0.25) Or .CenterVertically = False Then
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
    End With
End Sub
Sub SortSectors()
    Dim GridParts() As String
    Dim i As Integer
    Dim SortStart As Range
    Dim SortEnd As Range
    Dim SectionTotal As Range
    Dim GridRowSize As Integer
    Dim GridArea As Range
    Dim SectTotal As Range
    
    GridParts = Split("Large Value,Large Blend,Large Growth,Medium Value,Medium Blend,Medium Growth," _
          & "Small Value,Small Blend,Small Growth,Specialty Holdings", ",")
    
    For i = 0 To UBound(GridParts)
        If Grid.UsedRange.Find(GridParts(i), LookAt:=xlPart) Is Nothing _
            And (i = 1 And Grid.UsedRange.Find("Foreign") Is Nothing Or i <> 1) Then
                AddError """" & GridParts(i) & """ wasn't found. This category wasn't sorted.", False
        Else
            If Not Grid.Range(Grid.Range("A1"), Grid.Range("Z5")).Find("Foreign") Is Nothing And GridParts(i) = "Large Blend" Then
                Set SortStart = Grid.UsedRange.Find("Foreign").Offset(1, 0)
            Else
                Set SortStart = Grid.UsedRange.Find(GridParts(i)).Offset(1, 0)
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
            Set GridArea = Range(SortStart.Offset(-1, 0), SortEnd.Offset(1, 0))
            SectTotal.Value = "Sector Total"
            SectTotal.IndentLevel = 1
            
            GridArea.BorderAround LineStyle:=xlContinuous, Weight:=xlThin, ColorIndex:=xlColorIndexAutomatic
 
            GridRowSize = GridArea.Rows.Count - 2
            SectTotal.Offset(0, 2).FormulaR1C1 = "=SUM(R[" & -1 * GridRowSize & "]C:R[-1]C)"
            SectTotal.Offset(0, 3).FormulaR1C1 = "=SUM(R[" & -1 * GridRowSize & "]C:R[-1]C)"
            
            Range(SortStart, SortEnd).Sort Key1:=SortStart, Order1:=xlAscending, Header:=xlNo
        End If
    Next i
End Sub
Sub NameAndDate()
    Dim DayPart As String
    Dim MonthPart As String
    Dim YearPart As String
    Dim DateString As String
    Dim InvOrFund As Range
    Dim ClientName As Range
    Dim DateRange As Range
    Dim IncPercent As Range
    Dim EqPercent As Range
    
    DayPart = format(Day(Date - 1), "00")
    MonthPart = format(Month(Date - 1), "00")
    YearPart = Year(Date - 1)
    DateString = " - " & MonthPart & "/" & DayPart & "/" & YearPart
    
    Set InvOrFund = Portfolio.UsedRange.Find("Investment or Fund", After:=Portfolio.Range("A1"))
    Set ClientName = InvOrFund.Offset(-3, 0)
    Set DateRange = ClientName.Offset(1, 0)
    
    DateRange.Formula = "Portfolio Analysis" & DateString
    
    If ClientName = "David Bucholtz, Trustee" Then
        Grid.Range("A1").Formula = ClientName.Value & DateString
    ElseIf ClientName = "Tad (Chip) & Karen Bircher" Then
        Grid.Range("A1").Formula = "Chip & Karen Bircher" & DateString
    Else
        Grid.Range("A1").Formula = ClientName.Value & DateString
    End If
    
    Set IncPercent = InvOrFund.Offset(0, 4)
    Set EqPercent = InvOrFund.Offset(0, 7)
    
    IncPercent = "%"
    EqPercent = "%"
    IncPercent.HorizontalAlignment = xlCenter
    EqPercent.HorizontalAlignment = xlCenter
End Sub
Sub TotalCheck()
    Dim GridTotalLoc As Range
    Dim PortTotalLoc As Range
    Dim GridTotal As Variant
    Dim PortTotal As Variant
    
    Set GridTotalLoc = Grid.UsedRange.Find("Total", LookIn:=xlValues, LookAt:=xlWhole)
    Set PortTotalLoc = Portfolio.UsedRange.Find("Category Totals:", LookIn:=xlValues, LookAt:=xlWhole)
    
    Calculate
    If Not GridTotalLoc Is Nothing Then
        If Application.IsNA(GridTotalLoc.Offset(0, 2).Value) Or IsError(GridTotalLoc.Offset(0, 2).Value) Then
            AddError "One or more of the grid sector sums results in an error. Please check and rerun macro.", False
        Else
            GridTotal = GridTotalLoc.Offset(0, 2).Value
            PortTotal = PortTotalLoc.Offset(0, 5).Value
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
End Sub
