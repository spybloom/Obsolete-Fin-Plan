Attribute VB_Name = "AllTDSheets"
Option Explicit
Private TDBook As Workbook
Private ClientBook As Workbook
Private TDSheet As Worksheet
Private Portfolio As Worksheet
Private TransType As Range
Private TransTypeEnd As Range
Private MSArray() As String
Private NameArray() As String
Private AllErrors As String
Private AllTDTotal As Variant
Sub AllTD()
Attribute AllTD.VB_ProcData.VB_Invoke_Func = "H\n14"

'Keyboard Shortcut: Ctrl+Shift+H
'
'When all .csv files for a portfolio are exported from TD Ameritrade, this macro can be run on those
'   sheets to clean them up and copy/paste position values to the clients' portfolio sheet, as well as
'   show which funds should be on the Morningstar sheet. The TD sheets are then all printed, and if
'   no funds are bought or sold the DistGrid macro is run to complete the file.
'
'
'5/9/18     Created macro to run FinalTD macro on all open TD Ameritrade exports
'5/11/18    Tweaked bought/sold error message to specify which funds have been bought/sold, rather than
'           just a number. This fixes the problem of if an equal number are bought and sold.
'5/22/18    General cleaning up of entire macro. Added Morningstar text box. Added RemoveFund Sub
'5/23/18    Added link to run DistGrid if no funds are bought or sold (Since turned off until bugs are
'           all fixed)
'5/29/18    1. Fixed bug with funds being added to an empty account - PastePointEnd/AddSum
'           2. Changed error message to be a compilation of errors for all TD sheets
'           3. Fixed macro deleting money market from portfolio if MMDA12 isn't on the TD sheet and
'           there's a balance in cash alternatives. Money market still won't paste below the yellow cell,
'           but there should be an empty cell below pasted cells that the sum accounts for
'           4. Removed error for if there's no "% Mkt Val" on TD sheet. Turn back on if its absence causes
'           a problem
'6/25/18    1. Added function to sort arrays, so final MS list doesn't populate on the portfolio. Each
'           account's MS list still populates on the portfolio, though.
'           2. Changed condition for going straight into DistGrid to if the TD totals match the portfolio
'           total. Leaving this commented out for now, though, since it hasn't been fully tested.
'6/26/18    1. Changed Morningstar sub so it doesn't output onto the portfolio; everything is handled
'           with arrays now
'           2. If the TD sheet has a fee, money in, and/or money out, boxes should populate indicating
'           these so the sheets won't have to be written on. This doesn't account for uncommon cases where
'           (1) you have an equal amount in and out on the same date, (2) an influx or outflux of funds
'           via TRNs, (3) cash coming in from rollovers, or (4) VTRs in/out from outside accounts
'6/27/18    Added AddFund sub - if a new fund is in the normal fund list, then it will be added to the
'           other funds in the printable area of the account. This doesn't add in any known stocks yet,
'           only mutual funds and ETFs normally used by us. Also doesn't add new funds to the grid
'
'
'If funds are removed, need to make sure sum under yellow cell is deleted, too
'Need to make sure vlookup ranges are set for each fund in each account
'Need to print RMD and Bene sheets after IRAs if it's the appropriate time of year
'The sum of money in/out isn't right 100% of the time yet. Macro doesn't account for if
'   1. There's an equal amount of money in and out on the same date (offsetting)
'   2. An influx or outflux of funds via TRNs
'   3. Cash coming in from rollovers
'   4. VTRs in/out from outside accounts
'   When these are all fixed, this macro can be run into DistGrid without the need for DistGrid's input
'       box. This wouldn't account for heldaway accounts, however. Maybe compare the totals of the TD
'       accounts against the portfolio total (doesn't work with below the line TD accounts).
'Have a textbox with all the funds that were bought/sold for Morningstar.
'Need to add check for number of stock shares
'Because funds on the grid are most likely abbreviated, need to figure out how to integrate the grid with
'   which funds are bought/sold on the portfolio.
'   Consider bringing up a new grid style with columns for the ticker and account type.

    Dim NumberOfWindows As Integer
    Dim i As Integer
    Dim WindowName As String
    Dim TDArraySize As Integer
    Dim TDArray() As String
    Dim TDCaption As Integer
    
    ReDim MSArray(0)
    ReDim NameArray(0)
    
    'On Error GoTo BackOn
    StateToggle.UpdateScreen ("Off") 'Turns off screen updating to make macro run quicker
    
    Set ClientBook = Workbooks(BookCheck)
    ClientBook.Activate
    Set Portfolio = SetSheet("Portfolio")   'Want to set this asap because it's a top-level variable
                                            'Workbook needs to be activated first
    NumberOfWindows = Windows.Count
    TDArraySize = 0
    AllErrors = ""
    AllTDTotal = 0
    
    For i = 1 To NumberOfWindows 'Setup array to contain only csv files.
        WindowName = Windows(i).Caption
        If InStr(UCase(WindowName), "BALANCESPOSITIONHISTORY") > 0 Then
            ReDim Preserve TDArray(0 To TDArraySize)
            TDArray(TDArraySize) = Windows(i).Caption
            FinalTD TDArray(TDArraySize)
            TDArraySize = TDArraySize + 1
        End If
    Next i
    
    'If csv files are exported in the portfolio's order, sheets are printed in reverse order to follow this order.
    Dim k As Integer
    
    StateToggle.UpdateScreen ("On")
    Portfolio.Range("K:K").NumberFormat = "#,##0.00"
    For k = UBound(TDArray) To 0 Step -1
        Workbooks(TDArray(k)).Activate
        Worksheets(1).PrintOut From:=1, To:=1, Preview:=True
    Next k
    
    Portfolio.Activate
    
    If AllErrors <> "" Then
        MsgBox AllErrors
    End If
    
    If MSArray(0) <> vbNullString Then
        MsgBox MSList
    End If
    
'    'If there aren't any other accounts
'    'DistGrid still has the problem of if it's broken out into years, so this is still a problem here.
'    Dim PortTotal As Variant
'
'    PortTotal = Portfolio.UsedRange.Find("Total Investments:", After:=Portfolio.Range("A1")).Offset(0, 2).Value2
'
'    If AllTDTotal = PortTotal Then
'        ClientBook.Activate
'        Dist.DistGrid
'    End If
    
    Exit Sub
    
BackOn:
    StateToggle.UpdateScreen ("On")
    MsgBox ("Fatal error - Macro ended prematurely due to error in code.")
End Sub
Sub FinalTD(TDCaption As String)
    Set TDBook = Workbooks(TDCaption)
    TDBook.Activate
    Set TDSheet = Worksheets(1)
    TDSheet.Activate
    
    FormatSheets
    
    CleanTrans  'Sorts the transactions alphabetically and deletes extra transactions if there is money in or out
    
    If FindCheck Then 'If Symbol, Mkt Val, Client Account are on TD sheet, and the portfolio name contains Portfolio
        CleanMiddle 'Moves Mkt Val next to Symbol and deletes Mkt Val %
        CopyPaste   'Copies the position symbols and values and pastes them onto the portfolio sheet under
                    'the correct account
    End If
    
    AddError ""
End Sub
Sub CleanTrans()
    Dim TransLength As Long
    
    If TDSheet.UsedRange.Find("Trans Type", After:=Range("A1")) Is Nothing Then
        AddError " ""Trans Type"" not found. Transactions have not been sorted.", False
        Exit Sub
    End If

    Set TransType = TDSheet.UsedRange.Find("Trans Type", After:=Range("A1"))
    Set TransTypeEnd = TDSheet.Range(TransType.End(xlDown).Address)
    
    TransLength = TDSheet.Range(TransType.Offset(1, 0), TransTypeEnd).Rows.Count
    
    SortTrans TransLength
    DeleteLines
End Sub
Function FindCheck() As Boolean
    FindCheck = True
    
    If TDSheet.UsedRange.Find("Symbol") Is Nothing Then
        AddError " ""Symbol"" wasn't found. Account values were not copied over to the portfolio.", False
        FindCheck = False
    End If
    
    If TDSheet.UsedRange.Find("Mkt Val", After:=Range("A1"), LookAt:=xlWhole) Is Nothing Then
        AddError " ""Mkt Val"" wasn't found. Account values were not copied over to the portfolio.", False
        FindCheck = False
    End If
    
    If TDSheet.UsedRange.Find("Client Account", Range("A1"), xlValues, xlPart) Is Nothing Then
        AddError " ""Client Account"" wasn't found. Accout values were not copied over to the portfolio", False
        FindCheck = False
    End If
End Function
Sub CleanMiddle()
    'Move Market Value next to Symbol
    Dim MktValStart As Range
    Dim MktValEnd As Range
    Dim PastePoint As Range
    Dim i As Integer
    
    Set MktValStart = TDSheet.UsedRange.Find("Mkt Val", After:=Range("A1"), LookAt:=xlWhole)
    Set MktValEnd = TDSheet.Range(MktValStart.End(xlDown).Address)
    Set PastePoint = TDSheet.UsedRange.Find("Symbol").Offset(0, 1)
    
    If MktValStart <> PastePoint Then
        For i = 0 To Range(MktValStart, MktValEnd).Rows.Count
            PastePoint.Offset(i, 0).Value = MktValStart.Offset(i, 0).Value
            MktValStart.Offset(i, 0).ClearContents
        Next i
    End If
    
    'Delete Market Value %
    Dim PerMktValStart As Range
    Dim PerMktValEnd As Range
    
    If TDSheet.UsedRange.Find("% Mkt Val") Is Nothing Then
'        AddError " ""% Mkt Val"" wasn't found. This column may need to be deleted manually.", False
    Else
        Set PerMktValStart = TDSheet.UsedRange.Find("% Mkt Val")
        Set PerMktValEnd = TDSheet.Range(PerMktValStart.End(xlDown).Address)
        Range(PerMktValStart, PerMktValEnd).ClearContents
    End If
End Sub
Sub CopyPaste()
    'The Wajdas have a joint account above the line with a dozen or so bonds below the line.
    'If it's their joint account, skip it for now
    'Maybe add other problem accounts to skip here
    Dim ClientName() As String
    
    ClientName = Split(Portfolio.Range("A1").Value, " ")
    If ClientName(UBound(ClientName)) = "Wajda" And LastThree = 469 Then
        AddError " The Wajdas' joint account was not copied/pasted due to bonds below the line.", False
        Exit Sub
    End If

    'Set copy range
    Dim CopyPointStart As Range
    Dim CopyPointEnd As Range
    Dim CopySize As Integer
    
    Set CopyPointStart = TDSheet.UsedRange.Find("Symbol").Offset(1, 0)
    CopySize = 0
    
    Do 'Do loop instead of End(xlDown) prevents problems when there's only 1 fund
        Set CopyPointEnd = CopyPointStart.Offset(CopySize, 0)
        CopySize = CopySize + 1
    Loop While CopyPointEnd.Offset(1, 0) <> ""
    
    Set CopyPointEnd = CopyPointEnd.Offset(0, 1)
    
    If Portfolio Is Nothing Then
        AddError " Portfolio tab not named ""Portfolio"". Account values have not been pasted to " _
            & "client's portfolio.", False
        Exit Sub
    End If
    
    'Find yellow cell
    Dim YellowCell As Range
    
    If Portfolio.UsedRange.Find("Acct # xxx-xxx" & LastThree, After:=Range("A1")) Is Nothing Then
        AddError " Client's account wasn't found on portfolio sheet. Account values were not pasted.", False
        AllTDTotal = AllTDTotal + 1
        Exit Sub
    End If
    
    Set YellowCell = ClientAcctLoc
    Do While YellowCell.Interior.ColorIndex <> 6 And YellowCell.Column < 50
        Set YellowCell = YellowCell.Offset(0, 1)
    Loop
    
    If YellowCell.Column = 50 Then
        AddError " Cell with yellow background not found. Yellow cells must be on same row " _
            & "as account number. Account values have not been pasted.", False
        Exit Sub
    End If
    
    'Set paste range
    Dim PastePointStart As Range
    Dim PastePointEnd As Range
    Dim PasteSize As Integer
    
    Set PastePointStart = YellowCell.Offset(1, 0)
    PasteSize = 0
    
    Do 'Do loop instead of End(xlDown) prevents problems when there's only 1 fund
        Set PastePointEnd = PastePointStart.Offset(PasteSize, 0)
        PasteSize = PasteSize + 1
    Loop While PastePointEnd.Offset(1, 0) <> ""
    
    Set PastePointEnd = PastePointEnd.Offset(0, 1)
    
    'If funds are added or removed
    Dim TDFunds()
    Dim PortFunds()
    Dim i As Integer
    Dim j As Integer
    
    ReDim TDFunds(CopySize - 1)
    ReDim PortFunds(PasteSize - 1)
    
    For i = 0 To UBound(TDFunds)
        TDFunds(i) = CopyPointStart.Offset(i, 0).Value
    Next i
    
    For j = 0 To PasteSize - 1
        PortFunds(j) = PastePointStart.Offset(j, 0).Value
    Next j
    
    'Check if there's an element in either array that's not in the other
    Dim FundsBought()
    Dim FundsSold()
    
    ReDim FundsBought(Application.Max(UBound(TDFunds), UBound(PortFunds))) As Variant
    ReDim FundsSold(Application.Max(UBound(TDFunds), UBound(PortFunds))) As Variant
    
    FundsBought = ArrayCompare(TDFunds, PortFunds)
    FundsSold = ArrayCompare(PortFunds, TDFunds)
    
    AddRemove PastePointStart, PastePointEnd, YellowCell, FundsBought, FundsSold
    
    'Paste
    TDSheet.Range(CopyPointStart, CopyPointEnd).Copy
    Portfolio.Paste Destination:=PastePointStart, Link:=False
    Application.CutCopyMode = False
     
    'Make sure money market matches
    MMCheck PastePointStart, PastePointEnd
    
    'Check account value of TD sheet matches Portfolio
    Dim AcctTotalLoc As Range
    Dim AcctTotal As Variant
    Dim FixedRange As String
    Dim EqRange As String
    Dim SizeDifference As Integer
    
    SizeDifference = CopySize - PasteSize
    
    Calculate
    Set AcctTotalLoc = ClientAcctLoc.Offset(0, 1)
    FixedRange = Range(AcctTotalLoc.Offset(1, 2), AcctTotalLoc.Offset(PasteSize, 2)).Address
    EqRange = Range(AcctTotalLoc.Offset(1, 5), AcctTotalLoc.Offset(PasteSize, 5)).Address
    AcctTotalLoc.Formula = "=SUM(" & FixedRange & "," & EqRange & ")"
    AcctTotal = AcctTotalLoc.Value2
    
    If Not IsError(AcctTotal) Then
        If CopySize = PasteSize And _
            Round(PastePointEnd.Offset(SizeDifference + 1, 0).Value2, 0) <> Round(AcctTotal, 0) Then
                AddError " Account total doesn't match the sum of the exported TD values. " _
                    & "TD sheet has $" & Round(AcctTotal, 0) & ", portfolio has $" _
                    & Round(PastePointEnd.Offset(SizeDifference + 1, 0).Value2, 0) & ".", False
        End If
        
        AllTDTotal = AllTDTotal + AcctTotal
    End If
              
    'Add positions to Morningstar list
    Morningstar CopyPointStart, CopyPointEnd
End Sub
Sub FormatSheets()
    Application.DisplayStatusBar = True
    
    'Set column widths
    'Columns A and B contain more text, so these sizes should be fine rather than AutoFitting them
    Columns("A:A").ColumnWidth = 10
    Columns("B:B").ColumnWidth = 13.57
    Columns("C:H").AutoFit
    
    'Set printer to all columns on 1 page
    With TDSheet.PageSetup
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = False
        .PrintErrors = xlPrintErrorsDisplayed
        .ScaleWithDocHeaderFooter = True
        .AlignMarginsHeaderFooter = True
    End With
    
    'Change window size
    'This should move the windows over so everything on both sheets can be seen
    With ActiveWindow
        .WindowState = xlNormal
        .Height = Application.UsableHeight
        .Width = Application.UsableWidth * 0.4
        .Top = 0
        .Left = Application.UsableWidth * 0.6
    End With
    
    ClientBook.Activate
    With ActiveWindow
        .WindowState = xlNormal
        .Height = Application.UsableHeight
        .Width = Application.UsableWidth * 0.6
        .Top = 0
        .Left = 0
    End With
    
    Application.DisplayStatusBar = False
End Sub
Function AddError(Error As Variant, Optional Display As Boolean) As Integer
    Static ErrorMessage As String
    
    If Error <> "" Then
        If ErrorMessage = "" Then
            If LastThree <> "NA" Then
                ErrorMessage = "Acct # xxx-xxx" & LastThree & " (" & AcctType(TDSheet.Range("B1")) & ")" & vbNewLine
            End If
        End If
        
        ErrorMessage = ErrorMessage & Chr(149) & Error & vbNewLine
        If Display = True Then
            MsgBox (ErrorMessage)
            ErrorMessage = ""
            StateToggle.UpdateScreen "On"
            End
        End If
    ElseIf Error = "" And ErrorMessage <> "" Then
        If AllErrors = "" Then
            AllErrors = ErrorMessage
        Else
            AllErrors = ErrorMessage & vbNewLine & AllErrors
        End If
        
        ErrorMessage = ""
    End If
End Function
Sub SortTrans(TransLength As Long)
    Dim SortStart As Range
    Dim SortEnd As Range
    
    If TDSheet.UsedRange.Find("Date") Is Nothing Then
        AddError " ""Date"" not found under Account History on TD sheet. Transactions have not " _
            & "been sorted.", False
    ElseIf TransType.Offset(1, 0).Value <> "" Then
        Set SortStart = TDSheet.UsedRange.Find("Date").Offset(1, 0)
        Set SortEnd = TDSheet.Range(SortStart.Offset(-1, 0).End(xlToRight).Offset(TransLength, 0).Address)
        
        TDSheet.Range(SortStart, SortEnd).Sort Key1:=Range(TransType.Offset(1, 0), TransTypeEnd), _
            Order1:=xlAscending, Header:=xlNo
    End If
End Sub
Sub DeleteLines()
    Dim FindJNL As Range
    Dim FindTRN As Range
    Dim FindVTR As Range
    Dim Trans() As String
    Dim i As Integer
    Dim j As Integer
    Dim Symbol As String
        
    With TDSheet.Range(TransType, TransTypeEnd)
        Set FindJNL = .Find("JNL", After:=TransType)
        Set FindTRN = .Find("TRN", After:=TransType)
        Set FindVTR = .Find("VTR", After:=TransType)
    End With
    
    If Not FindJNL Is Nothing Or Not FindTRN Is Nothing Or Not FindVTR Is Nothing Then
        Trans = Split("BUY DIV TRD SELL DVIO", " ")
        
        For i = Range(TransType, TransTypeEnd).Rows.Count To 0 Step -1
            For j = 0 To UBound(Trans)
                If TransType.Offset(i, 0).Value = Trans(j) Then
                    TransType.Offset(i, 0).EntireRow.Delete
                End If
            Next j
        Next i
        
        Set TransTypeEnd = TDSheet.Range(TransType.End(xlDown).Address)
        
        TDSheet.PageSetup.FitToPagesTall = 1
        
        MoveVTR
        MoveFee
        PosNeg
        AddSum TransType.Offset(1, 3), TransTypeEnd.Offset(0, 3)
        
        If TransTypeEnd.Offset(1, 3) < 0 Then
            Symbol = ""
        Else
            Symbol = "+"
        End If
        
        If TransTypeEnd.Offset(1, 3) <> 0 Then
            TDSheet.Range("G8").Formula = "=CONCATENATE(""" & Symbol & """, TEXT(" _
                & TransTypeEnd.Offset(1, 3).Address & ",""#,###""))"
            Boxes TDSheet.Range("G8")
        End If
        
        TDSheet.Range(TransType.Offset(1, 1), TransTypeEnd.Offset(0, 1)).NumberFormat = "#,##0.00"
        TDSheet.Range(TransType.Offset(1, 2), TransTypeEnd.Offset(1, 4)).NumberFormat = "$#,##0.00_);[Red]($#,##0.00)"
    End If
End Sub
Function TDCount() As Integer
    Dim NumberOfWindows As Integer
    Dim i As Integer
    Dim WindowName As String
    
    NumberOfWindows = Windows.Count
    TDCount = 0
    
    For i = 1 To NumberOfWindows
        WindowName = Windows(i).Caption
        If InStr(UCase(WindowName), "BALANCESPOSITIONHISTORY") > 0 Then
            TDCount = TDCount + 1
        End If
    Next i
End Function
Function BookCheck() As String 'Used to make sure the .csv file is selected and the portfolio is named correctly
    Dim NumberOfWindows As Integer
    Dim i As Integer
    Dim WindowName As String
    
    NumberOfWindows = Windows.Count
    BookCheck = ""

    For i = 1 To NumberOfWindows
        WindowName = Windows(i).Caption
        If InStr(UCase(WindowName), "PORTFOLIO") > 0 Or InStr(UCase(WindowName), "PA") > 0 Then
            BookCheck = Windows(i).Caption
            If Windows(BookCheck).Caption = ActiveWindow.Caption Then
                AddError "Please select the exported .csv file from Veo.", True
            Else: Exit Function
            End If
        End If
    Next i
End Function
Function SetSheet(TargetSheet As String) As Worksheet
    Dim i As Integer
    
    For i = 1 To Worksheets.Count
        If InStr(UCase(Worksheets(i).Name), UCase(TargetSheet)) > 0 Then
            Set SetSheet = Worksheets(i)
            Exit For
        End If
    Next i
End Function
Function LastThree() As String
    Dim TDAccount As String
    
    If Not TDSheet Is Nothing Then
        If Not TDSheet.UsedRange.Find("Client Account", Range("A1"), xlValues, xlPart).Offset(0, 1) Is Nothing Then
            TDAccount = TDSheet.UsedRange.Find("Client Account", Range("A1"), xlValues, xlPart).Offset(0, 1).Value
            LastThree = Mid(TDAccount, Len(TDAccount) - 2, 3)
        Else: GoTo ElseCase
        End If
    Else: GoTo ElseCase
    End If
    Exit Function
    
ElseCase:
    LastThree = "NA"
End Function
Function AcctType(TargetRange As Range) As String
    Dim DiffTypes() As String
        DiffTypes = Split("IRA,SEP,BENE,SIMPLE,IND,INDIVIDUAL,REG,TOD,UTMA,UGMA,ROTH,72T,JT,JOINT,TRUST,401K,ROTH 401K", ",")
    Dim i As Integer
    
    For i = UBound(DiffTypes) To 0 Step -1
        If InStr(UCase(TargetRange), DiffTypes(i)) > 0 Then
            AcctType = DiffTypes(i)
            Exit For
        End If
    Next i
End Function
Sub AddRemove(PastePointStart As Range, PastePointEnd As Range, YellowCell As Range, FundsBought() As Variant, FundsSold() As Variant)
    Dim SizeDifference As Integer
    Dim RemoveLineStart As Range
    Dim RemoveLineEnd As Range
    Dim FundsBoughtSize As Integer
    Dim FundsSoldSize As Integer
    Dim i As Integer
    Dim j As Integer
    Dim Fund As String
    
    FundsBoughtSize = ArraySize(FundsBought)
    FundsSoldSize = ArraySize(FundsSold)
    
    If FundsBoughtSize > 0 Then
        AddError " " & ArraySplit(FundsBought) & PluralString(FundsBoughtSize) & "been added.", False
        
        For j = 0 To UBound(FundsBought)
            Fund = FundsBought(j)
            AddFund YellowCell, Fund
            'FundsBoughtSize = FundsBoughtSize - AddFund(YellowCell, Fund)
        Next j
    End If
    
    If FundsSoldSize > 0 Then
        AddError " " & ArraySplit(FundsSold) & PluralString(FundsSoldSize) & "been sold.", False
        
        For i = 0 To UBound(FundsSold)
            Fund = FundsSold(i)
            RemoveFund YellowCell, Fund
        Next i
    End If
    
    SizeDifference = FundsBoughtSize - FundsSoldSize
    If SizeDifference > 0 Then
        Rows(PastePointEnd.Offset(1, -1).Row & ":" & PastePointEnd.Offset(SizeDifference, 0).Row).Insert xlDown
        Set PastePointEnd = PastePointEnd.Offset(SizeDifference, 0)
    ElseIf SizeDifference < 0 Then
        Set RemoveLineStart = PastePointEnd.Offset(SizeDifference + 1, -1)
        Set RemoveLineEnd = PastePointEnd
        Range(RemoveLineStart, RemoveLineEnd).Delete Shift:=xlUp
        If PastePointStart.Offset(1, 0) = "" Then
            Set PastePointEnd = PastePointStart.Offset(0, 1)
        Else
            Set PastePointEnd = Range(PastePointStart.End(xlDown).Offset(0, 1).Address)
        End If
    End If
    
    'Add sum
    Dim SumStart As Range
    Dim SumEnd As Range
    
    Set SumStart = PastePointStart.Offset(0, 1)
    Set SumEnd = PastePointEnd
    AddSum SumStart, SumEnd
End Sub
Sub MMCheck(PastePointStart As Range, PastePointEnd As Range)
    Dim CashAlt As Single
    Dim AcctTotal As Variant
    Dim MoneyMarket As Range
    
    If TDSheet.UsedRange.Find("Cash Alternatives", Range("A1"), xlValues, xlPart) Is Nothing Then
        AddError " ""Cash Alternatives"" not found. Money market value may not be correct.", False
    Else
        CashAlt = TDSheet.UsedRange.Find("Cash Alternatives", Range("A1"), xlValues, xlPart).Offset(0, 1).Value2
    End If
    
    If Range(PastePointStart, PastePointEnd).Find("MMDA12") Is Nothing And _
        Range(PastePointStart, PastePointEnd).Find("ZFD90") Is Nothing Then
        AddError " ""MMDA12"" wasn't found on portfolio. Money Market amount may not be correct.", False
    Else
        If Not Range(PastePointStart, PastePointEnd).Find("MMDA12") Is Nothing Then
            Set MoneyMarket = Range(PastePointStart, PastePointEnd).Find("MMDA12").Offset(0, 1)
        Else
            Set MoneyMarket = Range(PastePointStart, PastePointEnd).Find("ZFD90").Offset(0, 1)
        End If
        
        If MoneyMarket.Value <> CashAlt Then
            MoneyMarket.Value = CashAlt
            MoneyMarket.Value = Application.WorksheetFunction.RoundDown(MoneyMarket.Value, 2)
        End If
    End If
End Sub
Sub Morningstar(CopyPointStart As Range, CopyPointEnd As Range)
'Adds any funds not in the master MSArray into it, excepting money market or stock symbols
    
    'Put names and tickers from TDSheet into array
    Dim NamesAndTickers() As Variant
    
    NamesAndTickers = TDSheet.Range(CopyPointStart.Offset(0, -1), CopyPointEnd.Offset(0, -1)).Value2
    
    'Extract text from stock files
    Dim TextFile As Integer
    Dim StockPath As String
    Dim StockTickers() As String
    Dim StockCount As Integer
    
    TextFile = FreeFile
    StockPath = "Z:\YungwirthSteve\Macros\Documents\StockTickers.txt"
    
    Open StockPath For Input As TextFile
        StockTickers = Split(Input(LOF(TextFile), TextFile), ",")
        
        'Remove spaces from stock array (Note: Need to remove spaces from the StockTickers file)
        For StockCount = 0 To UBound(StockTickers)
            StockTickers(StockCount) = Left(StockTickers(StockCount), Len(StockTickers(StockCount)) - 2)
        Next StockCount
    Close TextFile
    
    'Delete stocks and money market from Morningstar list
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    Dim Removed As Integer
    
    Removed = 0
    
    For i = 1 To UBound(NamesAndTickers)
        For j = 0 To UBound(StockTickers)
            If NamesAndTickers(i, 2) = StockTickers(j) _
                Or NamesAndTickers(i, 2) = "MMDA12" _
                Or NamesAndTickers(i, 2) = "ZFD90" Then
                
                If UBound(NamesAndTickers) > 1 Then
                    For k = i To UBound(NamesAndTickers) - 1
                        NamesAndTickers(k, 1) = NamesAndTickers(k + 1, 1)
                        NamesAndTickers(k, 2) = NamesAndTickers(k + 1, 2)
                    Next k
                    
                    NamesAndTickers(UBound(NamesAndTickers), 1) = vbNullString
                    NamesAndTickers(UBound(NamesAndTickers), 2) = vbNullString
                    Removed = Removed + 1
                Else
                    NamesAndTickers(i, 1) = vbNullString
                    NamesAndTickers(i, 2) = vbNullString
                End If
            End If
        Next j
    Next i
    
    'for each fund in morningstar list, put it in MSArray if it's not there already
    Dim m As Integer
    Dim n As Integer
    Dim InArr As Boolean
    
    For n = 1 To UBound(NamesAndTickers)
        If NamesAndTickers(n, 1) <> vbNullString Then
            For m = 0 To UBound(MSArray)
                If MSArray(m) = NamesAndTickers(n, 2) Then
                    InArr = True
                End If
            Next m
            
            If Not InArr Then
                MSArray(UBound(MSArray)) = NamesAndTickers(n, 2)
                NameArray(UBound(MSArray)) = NamesAndTickers(n, 1)
                
                If n <> UBound(NamesAndTickers) - Removed Then
                    ReDim Preserve MSArray(0 To UBound(MSArray) + 1) As String
                    ReDim Preserve NameArray(0 To UBound(NameArray) + 1) As String
                End If
            End If
            
            InArr = False
        End If
    Next n
End Sub
Sub MoveFee()
    Dim Descriptor As Variant
    Dim TargetLoc As Range
    Dim AmountLoc As Range
    
    For Each Descriptor In Range(TransType.Offset(0, -2), TransTypeEnd.Offset(0, -2))
        If InStr(UCase(Descriptor), "FEE") > 0 And Descriptor.Offset(0, 5) <> vbNullString Then
            Set TargetLoc = Descriptor
            TargetLoc.Offset(0, 4).Value = TargetLoc.Offset(0, 5).Value
            TargetLoc.Offset(0, 5).ClearContents
            TDSheet.Range("G4") = "Fee"
            Boxes TDSheet.Range("G4")
        End If
    Next Descriptor
End Sub
Sub MoveVTR()
    Dim Descriptor As Variant
    Dim TargetLoc As Range
    Dim AmountLoc As Range
    
    For Each Descriptor In Range(TransType, TransTypeEnd)
        If InStr(UCase(Descriptor), "VTR") > 0 Then
            Set TargetLoc = Descriptor
            TargetLoc.Offset(0, 2).Value = TargetLoc.Offset(0, 3).Value
            TargetLoc.Offset(0, 3).ClearContents
        End If
    Next Descriptor
End Sub
Sub PosNeg()
    Dim AmountStart As Range
    Dim AmountEnd As Range
    Dim PosCount As Integer
    Dim NegCount As Integer
    Dim k As Integer
    
    Set AmountStart = TransType.Offset(1, 3)
    Set AmountEnd = TransTypeEnd.Offset(0, 3)
    
    PosCount = Application.WorksheetFunction.CountIf(Range(AmountStart, AmountEnd), ">0")
    NegCount = Application.WorksheetFunction.CountIf(Range(AmountStart, AmountEnd), "<0")
    
    For k = 0 To Range(AmountStart, AmountEnd).Rows.Count
        If PosCount > 0 And NegCount > 0 And AmountStart.Offset(k, 0).Value > 0 Then
            AmountStart.Offset(k, 1) = AmountStart.Offset(k, 0).Value
            AmountStart.Offset(k, 0).ClearContents
        End If
        
        'Reinvested dividends come up as TRN, but there are values in Quantity and Price columns
        If AmountStart.Offset(k, -1) <> vbNullString And AmountStart.Offset(k, -2) <> vbNullString Then
            Range(AmountStart.Offset(k, 0), AmountStart.Offset(k, -2)).ClearContents
        End If
    Next k
    
    If PosCount > 0 And NegCount > 0 Then
        AddSum AmountStart.Offset(0, 1), AmountEnd.Offset(0, 1)
        TDSheet.Range("G12").Formula = "=CONCATENATE(""+""," & AmountEnd.Offset(1, 1).Address & ")"
        Boxes TDSheet.Range("G12")
    End If
End Sub
Sub AddSum(FirstAmount As Range, LastAmount As Range)
    LastAmount.Offset(1, 0).Formula = "=SUM(" & FirstAmount.Address & ":" & LastAmount.Address & ")"
    With LastAmount.Offset(1, 0).Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
End Sub
Function ArrayCompare(Array1 As Variant, Array2 As Variant) As Variant
'Returns elements in Array1 that aren't in Array2
    Dim InArr2 As Boolean
    Dim Result() As Variant
    Dim Array1Ele As Integer
    Dim Array2Ele As Integer
    Dim ResultSize As Integer
    Dim i As Integer
    Dim CashAlt As Variant
    
    InArr2 = False
    i = 0
    CashAlt = TDSheet.UsedRange.Find("Cash Alternatives", Range("A1"), xlValues, xlPart).Offset(0, 1).Value2
    ResultSize = Application.Max(UBound(Array1), UBound(Array2))
    ReDim Result(ResultSize)
    
    For Array1Ele = 0 To UBound(Array1)
        For Array2Ele = 0 To UBound(Array2)
            If Array1(Array1Ele) = Array2(Array2Ele) Then
                InArr2 = True
            End If
        Next Array2Ele
        
        If Not InArr2 Then  'And (Array1(Array1Ele) <> "MMDA12" Or CashAlt = 0) 'Uncommented causes MMDA12
                            'to not be added if it's not in arr1. Not sure if this is an issue yet
            Result(i) = Array1(Array1Ele)
            i = i + 1
        Else
            InArr2 = False
        End If
    Next Array1Ele
    
    If i > 0 Then
        ReDim Preserve Result(i - 1)
    Else
        ReDim Result(0)
    End If
    
    ArrayCompare = Result
End Function
Function ArraySplit(MyArr As Variant) As String
    Dim i As Integer

    ArraySplit = ""

    For i = 0 To UBound(MyArr)
        If i = 0 Then
            ArraySplit = MyArr(i)
        Else
            ArraySplit = ArraySplit & ", " & MyArr(i)
        End If
    Next i
End Function
Function ArraySize(MyArr As Variant) As Integer
    If MyArr(0) = "" Then
        ArraySize = 0
    Else
        ArraySize = UBound(MyArr) - LBound(MyArr) + 1
    End If
End Function
Function PluralString(StringSize As Integer) As String
    If StringSize = 1 Then
        PluralString = " has "
    Else
        PluralString = " have "
    End If
End Function
Function ClientAcctLoc() As Range
    'Search for cell with "Acct # xxx-xxxyyy", where yyy is these last 3 numbers
    Dim ClientAcct As Range
    Dim ClientAcctType As String
    Dim TDAcctType As String
    Dim FirstAddress As String
        
    Set ClientAcct = Portfolio.UsedRange.Find("Acct # xxx-xxx" & LastThree, After:=Range("A1"))
    FirstAddress = ClientAcct.Address
    ClientAcctType = AcctType(ClientAcct.Offset(-1, 0))
    TDAcctType = AcctType(TDSheet.Range("B1"))
    
    Do
        Set ClientAcct = Portfolio.UsedRange.Find("Acct # xxx-xxx" & LastThree, After:=ClientAcct)
        ClientAcctType = AcctType(ClientAcct.Offset(-1, 0))
    Loop While ClientAcctType <> TDAcctType And ClientAcct.Address <> FirstAddress
    
    Set ClientAcctLoc = Portfolio.Range(ClientAcct.Address)
End Function
Sub RemoveFund(YellowCell As Range, Fund As String)
    Dim FirstCell As Range
    Dim LastRow As Range
    Dim FundCount As Integer
    Dim LastCell As Range
    Dim AcctRange As Range
    Dim FundRow As Integer
    Dim CashAlt As Variant
    
    CashAlt = TDSheet.UsedRange.Find("Cash Alternatives", Range("A1"), xlValues, xlPart).Offset(0, 1).Value2
    
    If Fund <> "MMDA12" Or CashAlt = 0 Then
        Set FirstCell = ClientAcctLoc.Offset(1, 0)
        Set LastRow = Portfolio.Range(FirstCell.End(xlDown).Address)
        
        FundCount = Range(FirstCell, LastRow).Rows.Count
        
        Set LastCell = YellowCell.Offset(FundCount, -1)
        Set AcctRange = Range(FirstCell, LastCell)
        
        FundRow = Portfolio.Range(FirstCell, LastCell).Find(Fund, After:=FirstCell, LookAt:=xlWhole, _
            MatchCase:=False).Row
        
        Portfolio.Range("A" & FundRow, "I" & FundRow).Delete Shift:=xlUp
    End If
End Sub
Sub AddFund(YellowCell As Range, Fund As String)
    'Extract text from sector files
    Dim GridSectors As Variant
    Dim i As Integer
    Dim TextFile As Integer
    Dim FundPath As String
    Dim TickerPath As String
    Dim FundNames() As String
    Dim FundTickers() As String
    Dim Sector As Variant
    Dim AllSectors(0 To 10) As Variant
    
    GridSectors = Array("Bond", "LV", "LB", "LG", "MV", "MB", "MG", "SV", "SB", "SG", "Spec")
    
    For i = 0 To UBound(GridSectors)
        TextFile = FreeFile
        FundPath = "Z:\YungwirthSteve\Macros\Documents\Funds\" & GridSectors(i) & "Funds.txt"
        TickerPath = "Z:\YungwirthSteve\Macros\Documents\Funds\" & GridSectors(i) & "Tickers.txt"
        
        Open FundPath For Input As TextFile
            FundNames = Split(Input(LOF(TextFile), TextFile), ",")
        Close TextFile
        
        Open TickerPath For Input As TextFile
            FundTickers = Split(Input(LOF(TextFile), TextFile), ",")
        Close TextFile
        
        Sector = Array(GridSectors(i), FundNames, FundTickers)
        AllSectors(i) = Sector
    Next i
    
    'AllSectors(0 To 10)(0 To 2)(0 To Sector size)
    'First range is the sector
    'Second range holds the sector name abbreviation (0), fund name array (1), ticker array (2)
    '   If the second range is 0, there is no third range ((3)(0) is LG, (3)(0)(0) doesn't exist)
    'Third range is element in fund name/ticker array
    
    'Client's account range
    Dim FirstCell As Range
    Dim LastRow As Range
    Dim FundCount As Integer
    Dim LastCell As Range
    Dim AcctRange As Range
    
    Set FirstCell = ClientAcctLoc.Offset(1, 0)
    Set LastRow = Portfolio.Range(FirstCell.End(xlDown).Address)
    
    FundCount = Range(FirstCell, LastRow).Rows.Count
    
    Set LastCell = YellowCell.Offset(FundCount, -1)
    Set AcctRange = Range(FirstCell, LastCell)
        
    If Fund <> "MMDA12" Then
        'Search for new fund ticker in array
        Dim First As Integer
        Dim Third As Integer
        Dim FundInfo As Variant
        Dim Found As Boolean
        
        Found = False
        
        For First = 0 To 10
            For Third = 0 To UBound(AllSectors(First)(2))
                If Fund = AllSectors(First)(2)(Third) Then
                    FundInfo = Array(AllSectors(First)(0), AllSectors(First)(1)(Third), AllSectors(First)(2)(Third))
                    Found = True
                End If
            Next Third
        Next First
        
        If Not Found Then 'If it's not in the regular mutual funds/ETFs, search the stocks
            'Extract text from stock files
            Dim StockPath As String
            Dim NamePath As String
            Dim GridPath As String
            Dim StockTickers() As String
            Dim StockNames() As String
            Dim StockGrid() As String
            Dim StockCount As Integer
            Dim StockArray() As String
            
            StockPath = "Z:\YungwirthSteve\Macros\Documents\StockTickers.txt"
            NamePath = "Z:\YungwirthSteve\Macros\Documents\StockNames.txt"
            GridPath = "Z:\YungwirthSteve\Macros\Documents\StockGrid.txt"
            
            Open StockPath For Input As TextFile
                StockTickers = Split(Input(LOF(TextFile), TextFile), ",")
                
                'Remove spaces from stock array (Note: Need to remove spaces from the StockTickers file)
                For StockCount = 0 To UBound(StockTickers)
                    StockTickers(StockCount) = Left(StockTickers(StockCount), Len(StockTickers(StockCount)) - 2)
                Next StockCount
            Close TextFile
            
            Open NamePath For Input As TextFile
                StockNames = Split(Input(LOF(TextFile), TextFile), ",")
                
                'Remove spaces from stock array
                For StockCount = 0 To UBound(StockNames)
                    StockNames(StockCount) = Left(StockNames(StockCount), Len(StockNames(StockCount)) - 2)
                Next StockCount
            Close TextFile
            
            Open GridPath For Input As TextFile
                StockGrid = Split(Input(LOF(TextFile), TextFile), ",")
                
                'Remove spaces from stock array
                For StockCount = 0 To UBound(StockGrid)
                    StockGrid(StockCount) = Left(StockGrid(StockCount), Len(StockGrid(StockCount)) - 2)
                Next StockCount
            Close TextFile
    
            For i = 0 To UBound(StockTickers)
                If Fund = StockTickers(i) Then
                    FundInfo = Array(StockGrid(i), StockNames(i), StockTickers(i))
                    Found = True
                End If
            Next i
            
            If Not Found Then 'If the fund isn't in the regular used funds or stocks
                Range("A" & LastCell.Row + 1, LastCell.Offset(1, 0)).Insert xlDown
                AddError " " & Fund & " is being added, but it is not in the master fund list. Please add it to " _
                    & "the list and/or add it to the portfolio manually", False
                'AddFund = 0
                Exit Sub
            End If
        Else
            'AddFund = 1
        End If
         
        'Find where fund goes
        Dim FundRow As Range
        Dim BondCount As Integer
        Dim Upper As Integer
        Dim Lower As Integer
        Dim ValueLocOffset As Integer
        Dim j As Integer
        
        BondCount = Range(FirstCell.Offset(0, 3), Portfolio.Range(FirstCell.Offset(0, 3).End(xlDown).Address)).Rows.Count
        
        If FundInfo(0) = "Bond" Then
            Lower = 0
            Upper = BondCount - 1
            ValueLocOffset = 3
        Else
            Lower = BondCount
            Upper = FundCount - 1
            ValueLocOffset = 6
        End If
        
        For j = Lower To Upper
            If UCase(FundInfo(1)) < UCase(FirstCell.Offset(j, 0).Value2) _
            And FirstCell.Offset(j, 0).Value2 <> "TD Bank FDIC Money Market" Then '_
            'And FirstCell.Offset(j, ValueLocOffset).Value2 > 0 Then    This condition doesn't work if there
            '                                                           are multiple sold funds, and the value
            '                                                           one of them is #N/A. Don't remember why
            '                                                           it was necessary, though
                Set FundRow = FirstCell.Offset(j, 0)
                Exit For
            ElseIf FundInfo(1) > FirstCell.Offset(j, 0).Value2 And j = Upper Then
                Set FundRow = FirstCell.Offset(j + 1, 0)
            End If
        Next j
    Else
        Set FundRow = FirstCell
        FundInfo = Array("Bond", "TD Bank FDIC Money Market", "MMDA12")
        ValueLocOffset = 3
        'AddFund = 1
    End If
    
    'Insert row with fund info
    Dim AcctTotal As String
    Dim YellowRange As Range
    
    Range(FundRow, FundRow.Offset(0, AcctRange.Columns.Count - 1)).Insert xlDown
    Set YellowRange = Range(YellowCell.Offset(1, 0), Portfolio.Range(YellowCell.Offset(1, 0).End(xlDown).Address).Offset(1, 1))
    AcctTotal = Portfolio.UsedRange.Find("Total Investments:").Offset(0, 2).Address
    Set FundRow = FundRow.Offset(-1, 0)
    
    FundRow.Value2 = FundInfo(1)
    FundRow.Offset(0, 2).Value2 = FundInfo(2)
    FundRow.Offset(0, ValueLocOffset).Formula = "=VLOOKUP(" & FundRow.Offset(0, 2).Address(0, 0) & "," & YellowRange.Address & ",2,FALSE)"
    FundRow.Offset(0, ValueLocOffset + 1).Formula = "=" & FundRow.Offset(0, ValueLocOffset).Address(0, 0) & "/" & AcctTotal
End Sub
Function MSList() As String
    Dim NewArr As Variant
    Dim k As Integer
    
    MSList = ""
    
    NewArr = TwoDimArraySort(NameArray, MSArray)
    
    For k = 0 To UBound(NewArr)
        MSList = MSList & NewArr(k, 1) & vbNewLine
    Next k
    
    If MSList <> "" Then
        MSList = "Morningstar funds (not including held away accounts):" & vbNewLine & MSList
    End If
End Function
Function TwoDimArraySort(Arr1 As Variant, Arr2 As Variant) As Variant
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    Dim MaxSize As Integer
    Dim Result() As Variant
    
    MaxSize = Application.Max(UBound(Arr1), UBound(Arr2))
    ReDim Result(MaxSize, 1)
    
    Result(0, 0) = Arr1(0)
    Result(0, 1) = Arr2(0)
    
    For i = 0 To UBound(Arr1)
        For j = 0 To UBound(Result)
            If Arr1(i) < Result(j, 0) Then
                For k = UBound(Result) To j + 1 Step -1
                    Result(k, 0) = Result(k - 1, 0)
                    Result(k, 1) = Result(k - 1, 1)
                Next k
                
                Result(j, 0) = Arr1(i)
                Result(j, 1) = Arr2(i)
                
                Exit For
            ElseIf j = i Then
                Result(j, 0) = Arr1(j)
                Result(j, 1) = Arr2(j)
                
                Exit For
            End If
        Next j
    Next i
    
    TwoDimArraySort = Result
End Function
Sub Boxes(Target As Range)
    If Target.Font.Bold = False Then
        Set Target = Range(Target, Target.Offset(1, 1))
        Target.MergeCells = True
        Target.Font.Bold = True
        Target.Font.Size = 24
        Target.HorizontalAlignment = xlCenter
        Target.VerticalAlignment = xlCenter
        Target.BorderAround xlContinuous, xlThick
    End If
End Sub
