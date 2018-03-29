Attribute VB_Name = "TDFinal"
Option Explicit
Private TDBook As Workbook
Private TDSheet As Worksheet
Private ClientBook As Workbook
Private TransType As Range
Private TransTypeEnd As Range
Private Portfolio As Worksheet
Sub FinalTD()
'
'When a .csv file is exported from TD Ameritrade, this macro can be run on that sheet to clean it up
'   and copy/paste position values to the clients' portfolio sheet. It will also copy the position names
'   and symbols in order to easily verify it matches the Morningstar portfolio.
'
'Keyboard Shortcut: Ctrl+Shift+E
'
'3/29/18    Created macro to clean up the previous TD macro by adding and arranging subs/functions to
'           make macro more readable. Still need to clean up CopyPaste and Morningstar
'
'
    'On Error GoTo BackOn
    Set TDBook = ActiveWorkbook
    Set TDSheet = ActiveSheet
    
    StateToggle "Off" 'Turns off screen updating to make macro run quicker
    
    CleanTrans  'Sorts the transactions alphabetically and deletes extra transactions if there is money in or out
    
    If FindCheck Then
        CleanMiddle 'Moves Mkt Val next to Symbol and deletes Mkt Val %
        CopyPaste   'Copies the position symbols and values and pastes them onto the portfolio sheet under
                    'the correct account
    End If
    
    FormatSheets
    
    AddError ""
    StateToggle "On"
    TDSheet.PrintOut From:=1, To:=1, Preview:=True
    
    Exit Sub
    
BackOn:
    StateToggle ("On")
    MsgBox ("Macro ended prematurely due to execution error.")
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
Sub CleanTrans()
    Dim TransLength As Long
    
    If TDSheet.Cells.Find("Trans Type", After:=Range("A1")) Is Nothing Then
        AddError " ""Trans Type"" not found. Transactions have not been sorted.", False
        Exit Sub
    End If

    Set TransType = TDSheet.Cells.Find("Trans Type", After:=Range("A1"))
    Set TransTypeEnd = TDSheet.Range(TransType.End(xlDown).Address)
    TransLength = TDSheet.Range(TransType.Offset(1, 0), TransTypeEnd).Rows.Count
    
    SortTrans TransLength
    DeleteLines TransLength
End Sub
Function FindCheck() As Boolean
    FindCheck = True
    
    If TDSheet.Cells.Find("Symbol") Is Nothing Then
        AddError " ""Symbol"" wasn't found. Account values were not copied over to the portfolio.", False
        FindCheck = False
    End If
    
    If TDSheet.Cells.Find("Mkt Val", After:=Range("A1"), LookAt:=xlWhole) Is Nothing Then
        AddError " ""Mkt Val"" wasn't found. Account values were not copied over to the portfolio.", False
        FindCheck = False
    End If
    
    If TDSheet.Cells.Find("Client Account", Range("A1"), xlValues, xlPart) Is Nothing Then
        AddError " ""Client Account"" wasn't found. Accout values were not copied over to the portfolio", False
        FindCheck = False
    End If
    
    If BookCheck = "" Then
        AddError " Client's file name does not contain ""Portfolio"" or ""PA"". Account values " _
            & "were not copied over to the portfolio.", False
        FindCheck = False
    End If
End Function
Sub CleanMiddle()
    'Move Market Value next to Symbol
    Dim MktValStart As Range
    Dim MktValEnd As Range
    Dim PastePoint As Range
    
    Set MktValStart = TDSheet.Cells.Find("Mkt Val", After:=Range("A1"), LookAt:=xlWhole)
    Set MktValEnd = Range(MktValStart.End(xlDown).Address)
    
    Set PastePoint = TDSheet.Cells.Find("Symbol").Offset(0, 1)
    Range(MktValStart, MktValEnd).Cut
    TDSheet.Paste Destination:=PastePoint, Link:=False
    Application.CutCopyMode = False
    
    'Delete Market Value %
    Dim PerMktValStart As Range
    Dim PerMktValEnd As Range
    
    If TDSheet.Cells.Find("% Mkt Val") Is Nothing Then
        AddError " ""% Mkt Val"" wasn't found. This column may need to be deleted manually.", False
    Else
        Set PerMktValStart = TDSheet.Cells.Find("% Mkt Val")
        Set PerMktValEnd = Range(PerMktValStart.End(xlDown).Address)
        Range(PerMktValStart, PerMktValEnd).ClearContents
    End If
End Sub
Sub CopyPaste()
    'Switch window
    Set ClientBook = Workbooks(BookCheck)
    ClientBook.Activate
    
    If SetPort = False Then
        AddError " Portfolio tab not named ""Portfolio"". Account values have not been pasted to " _
            & "client's portfolio.", False
        Exit Sub
    End If
    
    'Search for cell with "Acct # xxx-xxxyyy", where yyy is these last 3 numbers
    Dim ClientAcct As Range
    Dim ClientAcctType As String
    Dim TDAcctType As String
    Dim FirstAddress As String
    
    If Portfolio.UsedRange.Find("Acct # xxx-xxx" & LastThree, Range("A1")) Is Nothing Then
        AddError " Client's account wasn't found on portfolio sheet. Account values were not pasted.", False
        Exit Sub
    End If
        
    Set ClientAcct = Portfolio.UsedRange.Find("Acct # xxx-xxx" & LastThree, After:=Range("A1"))
    FirstAddress = ClientAcct.Address
    ClientAcctType = AcctType(ClientAcct.Offset(-1, 0))
    TDAcctType = AcctType(TDSheet.Range("B1"))
    
    Do
        Set ClientAcct = Portfolio.UsedRange.Find("Acct # xxx-xxx" & LastThree, After:=ClientAcct)
        ClientAcctType = AcctType(ClientAcct.Offset(-1, 0))
    Loop While ClientAcctType <> TDAcctType And ClientAcct.Address <> FirstAddress
    
    'Set copy range
    Dim CopyPointStart As Range
    Dim CopyPointEnd As Range
    Dim CopySize As Integer
    
    TDSheet.Activate
    Set CopyPointStart = TDSheet.Cells.Find("Symbol").Offset(1, 0)
    
    If CopyPointStart.Offset(1, 0) = "" Then
        Set CopyPointEnd = CopyPointStart.Offset(0, 1)
    Else
        Set CopyPointEnd = Range(CopyPointStart.Offset(0, 1).End(xlDown).Address)
    End If

    CopySize = Range(CopyPointStart, CopyPointEnd).Rows.Count
    
    'Find paste range
    ClientBook.Activate
    Dim YellowCell As Range
    Dim PastePointStart As Range
    Dim PastePointEnd As Range
    
    Set YellowCell = ClientAcct
    Do While YellowCell.Interior.ColorIndex <> 6
        Set YellowCell = YellowCell.Offset(0, 1)
        If YellowCell.Column = 50 Then
            Exit Do
        End If
    Loop
    
    If YellowCell.Column = 50 Then
        AddError " Cell with yellow background not found. Yellow cells must be on same row " _
            & "as account number. Account values have not been pasted.", False
        Exit Sub
    End If
    
    Set PastePointStart = YellowCell.Offset(1, 0)
    
    If PastePointStart.Offset(1, 0) = "" Then
        Set PastePointEnd = PastePointStart.Offset(0, 1)
    Else
        Set PastePointEnd = Range(PastePointStart.End(xlDown).Offset(0, 1).Address)
    End If

    'If funds are added or removed (Overall; if the same number of funds are added and removed
    'the macro won't pick it up)
    AddRemove PastePointStart, PastePointEnd
    
    'Paste
    TDSheet.Range(CopyPointStart, CopyPointEnd).Copy
    Portfolio.Paste Destination:=PastePointStart, Link:=False
    Application.CutCopyMode = False
    
    'Add sum
    Dim SumStart As Variant
    Dim SumEnd As Variant
    
    SumStart = PastePointStart.Offset(0, 1).Address
    SumEnd = PastePointEnd.Offset(PasteDifference, 0).Address
    PastePointEnd.Offset(PasteDifference + 1, 0).Formula = "=SUM(" & SumStart & ":" & SumEnd & ")"
    Range(PastePointStart, PastePointEnd).NumberFormat = "#,##0.00"
     
    'Make sure money market matches
    MMCheck PastePointStart, PastePointEnd
    
    'Check account value of TD sheet matches Portfolio
    Dim AcctTotal As Variant
    
    If PasteDifference = 0 Then
        Calculate
        AcctTotal = ClientAcct.Offset(0, 1).Value2
        If IsError(AcctTotal) Then
        ElseIf Round(PastePointEnd.Offset(PasteDifference + 1, 0).Value2, 0) <> Round(AcctTotal, 0) Then
            AddError " Account total doesn't match the sum of the exported TD values. " _
                & "TD sheet has $" & Round(AcctTotal, 0) & ", portfolio has $" _
                & Round(PastePointEnd.Offset(PasteDifference + 1, 0).Value2, 0) & ".", False
        End If
    End If
                       
    'Move positions near values for Morningstar
    If InStr(Range("A1"), "Trees") = 0 And InStr(Range("A1"), "FBI") = 0 Then
        Morningstar CopyPointStart, CopyPointEnd
    End If
End Sub
Sub FormatSheets()
    Application.DisplayStatusBar = True
    TDSheet.Activate
    
    'Set column widths
    'Columns A and B contain more text, so these sizes should be fine rather than AutoFitting them
    Dim ColNum As Integer
    
    Columns("A:A").ColumnWidth = 10
    Columns("B:B").ColumnWidth = 13.57
    For ColNum = 3 To 8
        Columns(ColNum).EntireColumn.AutoFit
    Next ColNum
    
    'Set printer to all columns on 1 page
    With TDSheet.PageSetup
        .Zoom = False
        .FitToPagesWide = 1
        If .FitToPagesTall <> 1 Then
            .FitToPagesTall = False
        End If
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
End Sub
Function AddError(Error As String, Optional Display As Boolean) As Integer
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
Sub SortTrans(TransLength As Long)
    Dim SortStart As Range
    Dim SortEnd As Range
    
    If TDSheet.UsedRange.Find("Date") Is Nothing Then
        AddError " ""Date"" not found under Account History on TD sheet. Transactions have not " _
            & "been sorted.", False
    ElseIf TransType.Offset(1, 0).Value <> "" Then
        Set SortStart = TDSheet.UsedRange.Find("Date").Offset(1, 0)
        Set SortEnd = TDSheet.Range(SortStart.Offset(-1, 0).End(xlToRight).Offset(TransLength, 0).Address)
        
        TDSheet.Sort.SortFields.Clear
        TDSheet.Sort.SortFields.Add Key:=Range(TransType.Offset(1, 0), TransTypeEnd), _
            SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        With TDSheet.Sort
            .SetRange Range(SortStart, SortEnd)
            .Header = xlNo
            .MatchCase = False
            .Orientation = xlTopToBottom
            .SortMethod = xlPinYin
            .Apply
        End With
    End If
End Sub
Sub DeleteLines(TransLength As Long)
    Dim FindJNL As Range
    Dim FindTRN As Range
    Dim FindVTR As Range
    Dim Filter As Range
    Dim DeletedLines As Integer
        DeletedLines = 0
        
    With TDSheet.Range(TransType, TransTypeEnd)
        Set FindJNL = .Find("JNL", After:=TransType)
        Set FindTRN = .Find("TRN", After:=TransType)
        Set FindVTR = .Find("VTR", After:=TransType)
    End With
    
    If Not FindJNL Is Nothing Or Not FindTRN Is Nothing Or Not FindVTR Is Nothing Then
        Set Filter = TransType.Offset(2, 0)
        If Not Filter Is Nothing Then
            Do While Filter.Offset(-1, 0) <> ""
                If Filter.Offset(-1, 0).Value = "BUY" Or _
                    Filter.Offset(-1, 0).Value = "DIV" Or _
                    Filter.Offset(-1, 0).Value = "TRD" Or _
                    Filter.Offset(-1, 0).Value = "SELL" Or _
                    Filter.Offset(-1, 0).Value = "DVIO" Then
                        Filter.Offset(-1, 0).EntireRow.Delete
                        Set Filter = Filter.Offset(1, 0)
                        DeletedLines = DeletedLines + 1
                Else
                    Set Filter = Filter.Offset(1, 0)
                End If
            Loop
        End If
    End If
    
    If DeletedLines > 0 Then
        ActiveSheet.PageSetup.FitToPagesTall = 1
        If TransLength > 1 Then
            MoveVTR
            MoveFee
            PosNeg
            AddSum
            
            TDSheet.Range(TransType.Offset(1, 1), TransTypeEnd.Offset(0, 1)).NumberFormat = "#,##0.00"
            TDSheet.Range(TransType.Offset(1, 2), TransTypeEnd.Offset(1, 4)).NumberFormat = "$#,##0.00_);[Red]($#,##0.00)"
        End If
    End If
End Sub
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
Function SetPort() As Boolean
    Dim i As Integer
    
    SetPort = False
    For i = 1 To Worksheets.Count
        If InStr(UCase(Worksheets(i).Name), "PORTFOLIO") > 0 Then
            Set Portfolio = Worksheets(i)
            SetPort = True
            Portfolio.Activate
            Exit For
        End If
    Next i
End Function
Function LastThree() As Long
    Dim TDAccount As String
    
    TDAccount = TDSheet.Cells.Find("Client Account", Range("A1"), xlValues, xlPart).Offset(0, 1).Value
    LastThree = Mid(TDAccount, Len(TDAccount) - 2, 3)
End Function
Function AcctType(TargetRange As Range) As String
    Dim DiffTypes() As String
        DiffTypes = Split("IRA,SEP,BENE,SIMPLE,IND,INDIVIDUAL,REG,TOD,UTMA,ROTH,72T,JOINT,TRUST,401K,ROTH 401K", ",")
    Dim i As Integer
    
    For i = 0 To UBound(DiffTypes)
        If InStr(UCase(TargetRange), DiffTypes(i)) > 0 Then
            AcctType = DiffTypes(i)
            Exit For
        End If
    Next i
End Function
Sub AddRemove(PastePointStart As Range, PastePointEnd As Range)
    Dim PasteSize As Integer
    Dim PasteDifference As Integer
    Dim PluralString As String
    Dim RemoveLineStart As Range
    Dim RemoveLineEnd As Range
    
    PasteSize = Range(PastePointStart, PastePointEnd).Rows.Count
    PasteDifference = CopySize - PasteSize
    If Abs(PasteDifference) = 1 Then
        PluralString = " position has "
    Else
        PluralString = " positions have "
    End If
    
    If PasteDifference > 0 Then
        AddError " Macro has been completed and account values were pasted. " & PasteDifference _
            & PluralString & "been added.", False
        Rows(PastePointEnd.Offset(1, -1).Row & ":" & PastePointEnd.Offset(PasteDifference, 0).Row).Insert Shift:=xlShiftDown
    ElseIf PasteDifference < 0 Then
        AddError " Macro has been completed and account values were pasted. " & PasteDifference * -1 _
            & PluralString & "been sold from this account.", False
        Set RemoveLineStart = PastePointEnd.Offset(PasteDifference + 1, -1)
        Set RemoveLineEnd = PastePointEnd
        Range(RemoveLineStart, RemoveLineEnd).Delete Shift:=xlUp
        If PastePointStart.Offset(1, 0) = "" Then
            Set PastePointEnd = PastePointStart.Offset(0, 1)
        Else
            Set PastePointEnd = Range(PastePointStart.End(xlDown).Offset(PasteDifference * -1, 1).Address)
        End If
    End If
End Sub
Sub MMCheck(PastePointStart As Range, PastePointEnd As Range)
    Dim CashAlt As Long
    Dim AcctTotal As Variant
    Dim MoneyMarket As Range
    
    If TDSheet.Cells.Find("Cash Alternatives", Range("A1"), xlValues, xlPart) Is Nothing Then
        AddError " ""Cash Alternatives"" not found. Money market value may not be correct.", False
    Else
        CashAlt = TDSheet.Cells.Find("Cash Alternatives", Range("A1"), xlValues, xlPart).Offset(0, 1).Value
    End If
    
    If PasteDifference > 0 Then
        Set PastePointEnd = Range(PastePointStart.End(xlDown).Offset(0, 1).Address)
    End If
    If Range(PastePointStart, PastePointEnd).Find("MMDA12") Is Nothing And _
        Range(PastePointStart, PastePointEnd).Find("ZFD90") Is Nothing Then
        AddError " ""MMDA12"" wasn't found. Money Market amount may not be correct.", False
    Else
        If Not Range(PastePointStart, PastePointEnd).Find("MMDA12") Is Nothing Then
            Set MoneyMarket = Range(PastePointStart, PastePointEnd).Find("MMDA12").Offset(0, 1)
        Else
            Set MoneyMarket = Range(PastePointStart, PastePointEnd).Find("ZFD90").Offset(0, 1)
        End If
        
        If MoneyMarket <> CashAlt Then
            MoneyMarket = CashAlt
            If Not MoneyMarket.Offset(-1, 0) = "" Then
                MoneyMarket.Offset(-1, 0).Copy
                MoneyMarket.PasteSpecial (xlPasteFormats)
                Application.CutCopyMode = False
            Else
                MoneyMarket.Offset(1, 0).Copy
                MoneyMarket.PasteSpecial (xlPasteFormats)
                Application.CutCopyMode = False
            End If
        End If
    End If
End Sub
Sub Morningstar(CopyPointStart As Range, CopyPointEnd As Range)
'Copies the position names and symbols and pastes them near the position values on the portfolio
'sheet. When this macro is run for all the accounts in the portfolio, the column with the names
'can be deleted, and the column of symbols can easily be compared to the clients' Morningstar portfolio.
        
    Dim DupeStart As Range
    Dim DupeCheck As String
    Dim FirstStar As Range
    Dim DupeEnd As Range
    Dim TextFile As Integer
    Dim StockPath As String
    Dim StockTickers As String
    Dim StockArray() As String
    Dim Stock As Integer
    Dim FirstStarAddress As String
    Dim DupeEndAddress As String

    'Find first empty cell
    Set DupeStart = Portfolio.Range("M15")
    DupeCheck = "M100"

    Do
        Set DupeStart = DupeStart.Offset(1, 0)
    Loop While DupeStart.Value2 <> vbNullString And DupeStart.Address <> DupeCheck

    'Reset Morningstar files if this macro hasn't been run yet on the portfolio
    TextFile = FreeFile
    StockPath = "Z:\YungwirthSteve\Macros\Documents\StockTickers.txt"

    'Keeping this section uncommented because FirstStar needs to be set
    If DupeStart.Offset(-1, 0).Value = vbNullString Then 'Reset
        Set FirstStar = DupeStart
    Else
        If DupeStart.Offset(-2, 0).Value = vbNullString Then
            Set FirstStar = DupeStart.Offset(-1, 0)
        Else
            Set FirstStar = Range(DupeStart.Offset(-1, 0).End(xlUp).Address)
        End If
    End If
    FirstStarAddress = FirstStar.Address
    
    'Extract text from files
    Open StockPath For Input As TextFile
        StockTickers = Input(LOF(TextFile), TextFile)
        StockArray = Split(StockTickers, ",")
    Close TextFile

    'Add and sort Morningstar tickers on portfolio
    Dim DupeMoney As Range

    TDSheet.Range(CopyPointStart.Offset(0, -1), CopyPointEnd(1, 0)).Copy
    Portfolio.Paste Destination:=DupeStart, Link:=False
    
    Set DupeEnd = Range(FirstStar.End(xlDown).Offset(0, 1).Address)
    DupeEndAddress = DupeEnd.Address

    Portfolio.Range(FirstStar.Address, DupeEnd.Address).RemoveDuplicates Columns:=2, Header:=xlNo
    
    'Delete stocks from Morningstar list
    Dim n As Integer
    
    If FirstStar.Offset(1, 0).Value <> "" Then
        For n = Range(FirstStar, DupeEnd).Rows.Count To 0 Step -1
            If InStr(StockTickers, FirstStar.Offset(n, 1)) > 0 And FirstStar.Offset(n, 1) <> "VO" Then
                Range(FirstStar.Offset(n, 0), FirstStar.Offset(n, 1)).Delete Shift:=xlShiftUp
            End If
        Next n
        Set FirstStar = Range(FirstStarAddress)
        Set DupeEnd = Range(DupeEndAddress)
    End If
    
    With Portfolio.Sort
        .SortFields.Clear
        .SortFields.Add Key:=FirstStar, SortOn:=xlSortOnValues, Order:=xlAscending, _
            DataOption:=xlSortNormal
        .SetRange Range(FirstStar.Address, DupeEnd.Address)
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    If Not Range(FirstStar.Address, DupeEnd.Address).Find("MMDA12", After:=FirstStar) Is Nothing Then
        'Delete the money market from the Morningstar list
        Set DupeMoney = Range(FirstStar.Address, DupeEnd.Address).Find("MMDA12", After:=FirstStar)
        Range(DupeMoney.Offset(0, -1), DupeMoney).Delete Shift:=xlShiftUp
    ElseIf Not Range(FirstStar.Address, DupeEnd.Address).Find("ZFD90", After:=FirstStar) Is Nothing Then
        'Deletes Cheri Salmon's money market from the Morningstar list
        Set DupeMoney = Range(FirstStar.Address, DupeEnd.Address).Find("ZFD90", After:=FirstStar)
        Range(DupeMoney.Offset(0, -1), DupeMoney).Delete Shift:=xlShiftUp
    End If
End Sub
Sub MoveVTR()
    Dim VTRLoc As Range
    Dim FirstVTR As String
    
    If Not TDSheet.Range(TransType, TransTypeEnd).Find("VTR", After:=TransType) Is Nothing Then
        Set VTRLoc = TDSheet.Range(TransType, TransTypeEnd).Find("VTR", After:=TransType)
        FirstVTR = VTRLoc.Address
        Do
            Set VTRLoc = TDSheet.Range(TransType, TransTypeEnd).Find("VTR", After:=VTRLoc)
            VTRLoc.Offset(0, 2) = VTRLoc.Offset(0, 3).Value
            VTRLoc.Offset(0, 3).ClearContents
        Loop While VTRLoc.Address <> FirstVTR
    End If
End Sub
Sub MoveFee()
    Dim Descriptor As Variant
    Dim FeeLoc As Range
    
    For Each Descriptor In Range(TransType.Offset(0, -2), TransTypeEnd.Offset(0, -2))
        If InStr(Descriptor, "FEE") Or InStr(Descriptor, "Fee") > 0 Then
            Set FeeLoc = Descriptor
            FeeLoc.Offset(0, 4) = FeeLoc.Offset(0, 5).Value
            FeeLoc.Offset(0, 5).ClearContents
        End If
    Next Descriptor
End Sub
Sub PosNeg()
    Dim AmountStart As Range
    Dim AmountEnd As Range
    Dim PosCount As Integer
    Dim NegCount As Integer
    
    Set AmountStart = TransType.Offset(1, 3)
    Set AmountEnd = TransTypeEnd.Offset(0, 3)
    
    PosCount = Application.WorksheetFunction.CountIf(Range(AmountStart, AmountEnd), ">0")
    NegCount = Application.WorksheetFunction.CountIf(Range(AmountStart, AmountEnd), "<0")
    
    Dim k As Integer
    For k = 0 To Range(AmountStart, AmountEnd).Rows.Count
        If PosCount > 0 And NegCount > 0 Then
            If AmountStart.Offset(k, 0).Value > 0 Then
                AmountStart.Offset(k, 1) = AmountStart.Offset(k, 0).Value
                AmountStart.Offset(k, 0).ClearContents
            End If
        End If
        
        If AmountStart.Offset(k, -1) <> vbNullString And AmountStart.Offset(k, -2) <> vbNullString Then
            AmountStart.Offset(k, 0).ClearContents
        End If
    Next k
    
    If PosCount > 0 And NegCount > 0 Then
        AmountEnd.Offset(1, 1).Formula = "=SUM(" & AmountStart.Offset(0, 1).Address & ":" & AmountEnd.Offset(0, 1).Address & ")"
        With AmountEnd.Offset(1, 1).Borders(xlEdgeTop)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .ColorIndex = xlAutomatic
        End With
    End If
End Sub
Sub AddSum()
    Dim AmountStart As Range
    Dim AmountEnd As Range
    Dim AmountCount As Integer
    
    Set AmountStart = TransType.Offset(1, 3)
    Set AmountEnd = TransTypeEnd.Offset(0, 3)
    AmountCount = Range(AmountStart, AmountEnd).Rows.Count
    
    If AmountCount > 1 Then
        AmountEnd.Offset(1, 0).Formula = "=SUM(" & AmountStart.Address & ":" & AmountEnd.Address & ")"
        With AmountEnd.Offset(1, 0).Borders(xlEdgeTop)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .ColorIndex = xlAutomatic
        End With
    End If
End Sub
