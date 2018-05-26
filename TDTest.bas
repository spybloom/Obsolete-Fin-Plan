Attribute VB_Name = "Module5"
Sub Test_TD()
Attribute Test_TD.VB_ProcData.VB_Invoke_Func = "E\n14"
'
' TestTD Macro
'
' Used for cleaning up exported TD file, copying values, and pasting them into portfolio
' As of 5/18/17, this macro adjusts column size, resizes window, sets to print one column wide (or one page total),
'   deletes and sorts transaction lines, moves Mkt Val column over, deletes % Mkt Val column, copies values,
'   activates portfolio sheet, pastes values if the areas are the same size, changes money market if applicable,
'   and reactivates exported TD file for totaling transactions and printing.
'   Error cases have been added.
'   Conditions for where the accounts end in the same three numbers have been added
'   Conditions for different paste sizes have been added
'
' 11/21/17 Fixed transaction sorting so fees, VTRs, and both money in & out move as they should. Fee still doesn't
'   move every time
' 11/16/17 Going back to Veo since Veo One doesn't export quantities to 3 digits after the decimal. Cleaning up macro.
'   Among other things, need to change WirtzType to how it is in the Veo One macro, and need to add Morningstar section
'
' Keyboard Shortcut: Ctrl+Shift+E
'
    
    Dim TDBook As Workbook
    Dim ClientBook As Workbook
    Dim TDSheet As Worksheet
    Dim BookName As String
    
    On Error GoTo BackOn
    
    StateToggle "Off"
    
    Set TDBook = ActiveWorkbook
    Set TDSheet = ActiveSheet
    BookName = BookCheck
    TDSheet.PageSetup.FitToPagesTall = False
    
    If BookName = "" Then
        AddError " Client's file name does not contain ""Portfolio"" or ""PA"". Account values " _
            & "will need to be pasted manually.", False
        CleanAndCopy TDBook, TDSheet
    Else
        Set ClientBook = Workbooks(BookName)
        CleanAndCopy TDBook, TDSheet, ClientBook
    End If
    
    'Select Market Value
    Dim MktValStart As Range
    Dim MktValEnd As Range
    Dim PastePoint As Range
    
    If Cells.Find("Mkt Val", After:=Range("A1"), LookAt:=xlWhole) Is Nothing Then
        AddError " ""Mkt Val"" wasn't found. Macro has been halted.", True
    End If
    
    Set MktValStart = Cells.Find("Mkt Val", After:=Range("A1"), LookAt:=xlWhole)
    Set MktValEnd = Range(MktValStart.End(xlDown).Address)
    
    'Move Market Value next to Symbol
    If Cells.Find("Symbol") Is Nothing Then
        AddError " ""Symbol"" wasn't found. Macro has been halted." & vbNewLine, True
    End If
    
    Set PastePoint = Cells.Find("Symbol").Offset(0, 1)
    Range(MktValStart, MktValEnd).Cut
    ActiveSheet.Paste Destination:=PastePoint, Link:=False
    Application.CutCopyMode = False
    
    'Set copy range
    Dim CopyPointStart As Range
    Dim CopyPointEnd As Range
    Dim CopySize As Integer
    
    Set CopyPointStart = Cells.Find("Symbol").Offset(1, 0)
    
    If CopyPointStart.Offset(1, 0) = "" Then
        Set CopyPointEnd = CopyPointStart.Offset(0, 1)
    Else
        Set CopyPointEnd = Range(CopyPointStart.Offset(0, 1).End(xlDown).Address)
    End If

    CopySize = Range(CopyPointStart, CopyPointEnd).Rows.Count
    
    'Delete Market Value %
    Dim PerMktValStart As Range
    Dim PerMktValEnd As Range
    
    If Cells.Find("% Mkt Val") Is Nothing Then
        AddError " ""% Mkt Val"" wasn't found. This column may need to be deleted manually.", False
    Else
        Set PerMktValStart = Cells.Find("% Mkt Val")
        Set PerMktValEnd = Range(PerMktValStart.End(xlDown).Address)
        Range(PerMktValStart, PerMktValEnd).ClearContents
    End If
    
    'Take the last three numbers of the account number from the TD sheet
    Dim TDAccount As String
    Dim LastThree As String
    Dim CashAlt As Variant
    Dim WirtzType As String
    
    If Cells.Find("Client Account", Range("A1"), xlValues, xlPart) Is Nothing Then
        AddError " ""Client Account"" wasn't found. Macro has been halted", True
    End If
    
    TDAccount = Cells.Find("Client Account", Range("A1"), xlValues, xlPart).Offset(0, 1).Value
    LastThree = Mid(TDAccount, Len(TDAccount) - 2, 3)
    
    If Cells.Find("Cash Alternatives", Range("A1"), xlValues, xlPart) Is Nothing Then
        AddError " ""Cash Alternatives"" not found. Money market value may not be correct.", False
    Else
        CashAlt = Cells.Find("Cash Alternatives", Range("A1"), xlValues, xlPart).Offset(0, 1).Value
    End If
    
    If InStr(UCase(Range("B1")), "ROTH") > 0 Then
        WirtzType = "Roth"
    ElseIf InStr(UCase(Range("B1")), "IRA") > 0 Then
        WirtzType = "IRA"
    ElseIf InStr(UCase(Range("B1")), "REG") > 0 Then
        WirtzType = "Reg"
    End If
    
    'Switch window
    Dim NumberOfWindows As Integer
    Dim Count1 As Integer
    Dim WindowName As String
    Dim MatchedWindowNumber As Integer
    Dim PortCount As Integer
    Dim j As Integer
    Dim Portfolio As Worksheet
    Dim ColNum As Integer
    
    NumberOfWindows = Windows.Count

    For Count1 = 1 To NumberOfWindows
        WindowName = Windows(Count1).Caption
        If InStr(UCase(WindowName), "PORTFOLIO") > 0 Then
            MatchedWindowNumber = Count1
        End If
    Next Count1

    If MatchedWindowNumber > 0 Then
        Windows(MatchedWindowNumber).Activate
        Set ClientBook = ActiveWorkbook
        With ActiveWindow
            .WindowState = xlNormal
            .Height = Application.UsableHeight
            .Width = Application.UsableWidth * 0.6
            .Top = 0
            .Left = 0
        End With
        
        PortCount = 0
        For j = 1 To Worksheets.Count
            If InStr(UCase(Worksheets(j).Name), "PORTFOLIO") > 0 Then
                Set Portfolio = Worksheets(j)
                PortCount = 1
                Exit For
            End If
        Next j
    End If
    
    If PortCount = 0 Then
        AddError " Portfolio tab not named ""Portfolio"". Account values have not been pasted to " _
            & "client's portfolio.", False
    Else
        Portfolio.Activate
        'Search for cell with "Acct # xxx-xxxyyy", where yyy is these last 3 numbers
        Dim ClientAccount As Range
        
        If Portfolio.UsedRange.Find("Acct # xxx-xxx" & LastThree, Range("A1")) Is Nothing Then
            AddError " Client's account wasn't found on portfolio sheet. Account values were not pasted.", False
        Else
            Set ClientAccount = Portfolio.UsedRange.Find("Acct # xxx-xxx" & LastThree, After:=Range("A1"))
            
            If InStr(ClientAccount.Offset(-1, 0), WirtzType) < 1 Or _
                (InStr(ClientAccount.Offset(-1, 0), "Roth IRA") > 0 And WirtzType = "IRA") Then
                Set ClientAccount = Portfolio.UsedRange.Find("Acct # xxx-xxx" & LastThree, After:=ClientAccount)
            End If
        
            'Find paste point
            Dim YellowCell As Range
            
            Set YellowCell = ClientAccount
            Do While YellowCell.Interior.ColorIndex <> 6
                Set YellowCell = YellowCell.Offset(0, 1)
                If YellowCell.Column = 50 Then
                    Exit Do
                End If
            Loop
            
            If YellowCell.Column = 50 Then
                    AddError " Cell with yellow background not found. Yellow cells must be on same row " _
                        & "as account number. Account values have not been pasted.", False
            Else
                'Paste
                Dim PastePointStart As Range
                Dim PastePointEnd As Range
                Dim PasteSize As Integer
                Dim PasteDifference As Integer
                Dim PluralString As String
                Dim RemoveLineStart As Range
                Dim RemoveLineEnd As Range
                Dim SumStart As Variant
                Dim SumEnd As Variant
                Dim AccountTotal As Variant
                
                Set PastePointStart = YellowCell.Offset(1, 0)
                
                If PastePointStart.Offset(1, 0) = "" Then
                    Set PastePointEnd = PastePointStart.Offset(0, 1)
                Else
                    Set PastePointEnd = Range(PastePointStart.End(xlDown).Offset(0, 1).Address)
                End If
            
                PasteSize = Range(PastePointStart, PastePointEnd).Rows.Count
                PasteDifference = CopySize - PasteSize
                If PasteDifference = 1 Or PasteDifference = -1 Then
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
                
                TDSheet.Range(CopyPointStart, CopyPointEnd).Copy
                Portfolio.Paste Destination:=PastePointStart, Link:=False
                Application.CutCopyMode = False
                
                SumStart = PastePointStart.Offset(0, 1).Address
                SumEnd = PastePointEnd.Offset(PasteDifference, 0).Address
                PastePointEnd.Offset(PasteDifference + 1, 0).Formula = "=SUM(" & SumStart & ":" & SumEnd & ")"
                Range(PastePointStart, PastePointEnd).NumberFormat = "#,##0.00"
                 
                'Make sure money market matches
                Dim MoneyMarket As Range
                
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
                
                'Check account value of TD sheet matches Portfolio
                If PasteDifference = 0 Then
                    Calculate
                    AccountTotal = ClientAccount.Offset(0, 1).Value2
                    If IsError(AccountTotal) Then
                    ElseIf Round(PastePointEnd.Offset(PasteDifference + 1, 0).Value2, 0) <> Round(AccountTotal, 0) Then
                        AddError " Account total doesn't match the sum of the exported TD values. " _
                            & "TD sheet has $" & Round(AccountTotal, 0) & ", portfolio has $" _
                            & Round(PastePointEnd.Offset(PasteDifference + 1, 0).Value2, 0) & ".", False
                    End If
                End If
                       
                'Move positions next to values for Morningstar
                
                If InStr(Range("A1"), "Trees") = 0 And InStr(Range("A1"), "FBI") = 0 Then
                    Morningstar TDSheet, Portfolio, CopyPointStart, CopyPointEnd
                End If
            End If
        End If
    End If
    
    TDSheet.Activate
    
    'Set column widths
    'Columns A and B contain more text, so these sizes should be fine rather than AutoFitting them
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
    'This should move the window over so everything on both sheets can be seen
    With ActiveWindow
            .WindowState = xlNormal
            .Height = Application.UsableHeight
            .Width = Application.UsableWidth * 0.4
            .Top = 0
            .Left = Application.UsableWidth * 0.6
    End With
    
    AddError ""
    StateToggle "On"
    TDSheet.PrintOut From:=1, To:=1, Preview:=True
    
    Exit Sub
    
BackOn:
    StateToggle ("On")
    MsgBox ("Macro ended prematurely due to error in execution.")
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
        'Application.DisplayStatusBar = False
        Application.Calculation = xlCalculationManual
    ElseIf OnOrOff = "On" Then
        Application.ScreenUpdating = OriginalScreen
        Application.EnableEvents = OriginalEvents
        Application.DisplayStatusBar = OriginalStatus
        Application.Calculation = OriginalCalc
        Reset = ActiveSheet.UsedRange.Rows.Count
    End If
End Sub
Sub CleanAndCopy(TDBook As Workbook, TDSheet As Worksheet, Optional ClientBook As Workbook)
    'Delete extra transactions
    If TDSheet.Cells.Find("Trans Type", After:=Range("A1")) Is Nothing Then
        AddError " ""Trans Type"" not found. Transactions have not been sorted.", False
    Else ': DeleteLines
        Dim TransType As Range
        Dim TransTypeEnd As Range
        Dim FindJNL As Range
        Dim FindTRN As Range
        Dim FindVTR As Range
        Dim Filter As Range
        Dim DeletedLines As Integer
            DeletedLines = 0
            
        Set TransType = TDSheet.Cells.Find("Trans Type", After:=Range("A1"))
        Set TransTypeEnd = TDSheet.Range(TransType.End(xlDown).Address)
        Set FindJNL = TDSheet.Range(TransType, TransTypeEnd).Find("JNL", After:=TransType)
        Set FindTRN = TDSheet.Range(TransType, TransTypeEnd).Find("TRN", After:=TransType)
        Set FindVTR = TDSheet.Range(TransType, TransTypeEnd).Find("VTR", After:=TransType)
        
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
        
        'Sort transactions
        Dim SortStart As Range
        Dim SortEnd As Range
        Dim AmountStart As Range
        Dim AmountEnd As Range
        Dim AmountCount As Integer
        Dim FirstVTR As String
        Dim VTRLoc As Range
        Dim FeeLoc As Range
        Dim PosCount As Integer
        Dim NegCount As Integer
        
        If TransType.Offset(1, 0).Value = "" Then
        ElseIf TDSheet.UsedRange.Find("Date") Is Nothing Then
            AddError " ""Date"" not found under Account History on TD sheet. Transactions have not " _
                & "been sorted.", False
        Else
            Set TransTypeEnd = TDSheet.Range(TransType.End(xlDown).Address)
            TransLength = TDSheet.Range(TransType.Offset(1, 0), TransTypeEnd).Rows.Count
            Set SortStart = TDSheet.UsedRange.Find("Date").Offset(1, 0)
            Set SortEnd = TDSheet.Range(SortStart.Offset(-1, 0).End(xlToRight).Offset(TransLength, 0).Address)
            
            ActiveSheet.Sort.SortFields.Clear
            ActiveSheet.Sort.SortFields.Add _
                Key:=Range(TransType.Offset(1, 0), TransTypeEnd), SortOn:=xlSortOnValues, Order:=xlAscending, _
                DataOption:=xlSortNormal
            With ActiveWorkbook.ActiveSheet.Sort
                .SetRange Range(SortStart, SortEnd)
                .Header = xlNo
                .MatchCase = False
                .Orientation = xlTopToBottom
                .SortMethod = xlPinYin
                .Apply
            End With
            
            If DeletedLines > 0 And TransLength > 1 Then
                'Move VTRs over
                If Not FindVTR Is Nothing Then
                    Set VTRLoc = TDSheet.Range(TransType, TransTypeEnd).Find("VTR", After:=TransType)
                    FirstVTR = VTRLoc.Address
                    Do
                        Set VTRLoc = TDSheet.Range(TransType, TransTypeEnd).Find("VTR", After:=VTRLoc)
                        VTRLoc.Offset(0, 2) = VTRLoc.Offset(0, 3).Value
                        VTRLoc.Offset(0, 3).ClearContents
                    Loop While VTRLoc.Address <> FirstVTR
                End If
                
                'Move fee over
                For Each Descriptor In Range(TransType.Offset(0, -2), TransTypeEnd.Offset(0, -2))
                    If InStr(Descriptor, "FEE") Or InStr(Descriptor, "Fee") > 0 Then
                        Set FeeLoc = Descriptor
                        FeeLoc.Offset(0, 4) = FeeLoc.Offset(0, 5).Value
                        FeeLoc.Offset(0, 5).ClearContents
                    End If
                Next Descriptor
                
                'Separate positive and negative transactions
                Set AmountStart = TransType.Offset(1, 3)
                Set AmountEnd = TransTypeEnd.Offset(0, 3)
                
                PosCount = Application.WorksheetFunction.CountIf(Range(AmountStart, AmountEnd), ">0")
                NegCount = Application.WorksheetFunction.CountIf(Range(AmountStart, AmountEnd), "<0")
                
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
                
                'Add sum to the bottom
                AmountCount = Application.WorksheetFunction.Count(Range(AmountStart, AmountEnd))
                
                If AmountCount > 1 Then
                    AmountEnd.Offset(1, 0).Formula = "=SUM(" & AmountStart.Address & ":" & AmountEnd.Address & ")"
                    With AmountEnd.Offset(1, 0).Borders(xlEdgeTop)
                        .LineStyle = xlContinuous
                        .Weight = xlThin
                        .ColorIndex = xlAutomatic
                    End With
                End If
                
                TDSheet.Range(TransType.Offset(1, 1), TransTypeEnd.Offset(0, 1)).NumberFormat = "#,##0.00"
                TDSheet.Range(TransType.Offset(1, 2), TransTypeEnd.Offset(1, 4)).NumberFormat = "$#,##0.00_);[Red]($#,##0.00)"
            End If
        End If
    End If
    
    'Clean up position values
    'CleanMiddle
    
    'Copy values from TD sheet to portfolio
    If Not ClientBook Is Nothing Then
        'SetupPages TDBook, ClientBook
        
        If ActiveSheet.Cells.Find("Client Account", Range("A1"), xlValues, xlPart) Is Nothing Then
            AddError " ""Client Account"" wasn't found. Macro has been halted", True
        End If
        
        'CopyValues TDBook, ClientBook
    Else: 'SetupPages TDBook
    End If
    
    If DeletedLines > 0 Then
        ActiveSheet.PageSetup.FitToPagesTall = 1
    End If
End Sub
Sub Morningstar(TDSheet As Worksheet, Portfolio As Worksheet, CopyPointStart As Range, CopyPointEnd As Range)
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
'    Dim MTickerPath As String
'    Dim MTickerList As String
'    Dim MTickerArray() As String
'    Dim MNamePath As String
'    Dim MNameList As String
'    Dim MNameArray() As String
'    Dim k As Integer
'    Dim Symbol As String
'    Dim m As Integer
'    Dim EOList As Integer
'    Dim FirstRun As Boolean
'    Dim Ticker As Integer

    'Find first empty cell
    Set DupeStart = Portfolio.Range("M15")
    DupeCheck = "M100"

    Do
        Set DupeStart = DupeStart.Offset(1, 0)
    Loop While DupeStart.Value2 <> vbNullString And DupeStart.Address <> DupeCheck

'    'Reset Morningstar files if this macro hasn't been run yet on the portfolio
    TextFile = FreeFile
    StockPath = "Z:\YungwirthSteve\Macros\Documents\StockTickers.txt"
'    MTickerPath = "Z:\YungwirthSteve\Macros\Documents\MorningstarTickers.txt"
'    MNamePath = "Z:\YungwirthSteve\Macros\Documents\MorningstarNames.txt"

    'Keeping this section uncommented because FirstStar needs to be set
    If DupeStart.Offset(-1, 0).Value = vbNullString Then 'Reset
        Set FirstStar = DupeStart
'        Open MTickerPath For Output As TextFile: Close TextFile
'        Open MNamePath For Output As TextFile: Close TextFile
'        FirstRun = True
    Else
        If DupeStart.Offset(-2, 0).Value = vbNullString Then
            Set FirstStar = DupeStart.Offset(-1, 0)
        Else
            Set FirstStar = Range(DupeStart.Offset(-1, 0).End(xlUp).Address)
        End If
'        FirstRun = False
    End If
    FirstStarAddress = FirstStar.Address
    
'This section is all good
    'From Reddit:
    'Dim arrSomeRange As Variant
    '
    'arrSomeRange = Range("A1:Z1000")
    '
    'Therefore:
    'MNameArray = Range(DupeStart.Offset(m, 0),DupeStart.Offset(UBound(MTickerArray), 0))
    'MTickerArray = Range(DupeStart.Offset(m, 0),DupeStart.Offset(UBound(MTickerArray), 0))
    'If the tickers are put on Excel first, this can be used to create the array used to fill the files
    'Possibly put all the tickers onto Excel, put them all in an array, remove them if they don't belong in the array
    'clear the tickers on Excel, and then put the array into Excel
'
    'Extract text from files
    Open StockPath For Input As TextFile
        StockTickers = Input(LOF(TextFile), TextFile)
        StockArray = Split(StockTickers, ",")
    Close TextFile
'
'    Open MTickerPath For Input As TextFile
'        MTickerList = Input(LOF(TextFile), TextFile)
'    Close TextFile
'
'    Open MNamePath For Input As TextFile
'        MNameList = Input(LOF(TextFile), TextFile)
'    Close TextFile
'
'    If FirstRun = True Then
'        EOList = 0
'    Else
'        EOList = UBound(Split(MTickerList, ","))
'    End If
'
'    'Add tickers to Morningstar files
'    Ticker = 0
'    Open MTickerPath For Append As TextFile
'    Open MNamePath For Append As TextFile + 1
'
'    For k = 0 To Range(CopyPointStart, CopyPointEnd).Rows.Count - 1
'        Symbol = CopyPointStart.Offset(k, 0).Value
'
'        If MTickerList = vbNullString _
'        And InStr(StockTickers, Symbol) = 0 _
'        And Symbol <> "MMDA12" Then 'If the Morningstar list is empty and if it's not on the stock list or MMDA12
'            MTickerList = MTickerList & Symbol & ","
'            MNameList = MNameList & CopyPointStart.Offset(Ticker, -1) & ","
'
'            MTickerArray = Split(MTickerList, ",")
'            MNameArray = Split(MNameList, ",")
'
'            Print #TextFile + 1, MNameArray(Ticker + EOList)
'            Print #TextFile + 1, ","
'            Print #TextFile, MTickerArray(Ticker + EOList)
'            Print #TextFile, ","
'
'            Ticker = Ticker + 1
'        ElseIf InStr(MTickerList, Symbol) = 0 _
'        And InStr(StockTickers, Symbol) = 0 _
'        And Symbol <> "MMDA12" Then 'If it's not already on the list, not on the stock list, and not MMDA12
'            MTickerList = MTickerList & Symbol & "," 'Add ticker to the ticker list
'            MNameList = MNameList & CopyPointStart.Offset(k, -1) & "," 'Add name to the name list
'
'            MTickerArray = Split(MTickerList, ",")
'            MNameArray = Split(MNameList, ",")
'
'            Print #TextFile + 1, MNameArray(Ticker + EOList)
'            Print #TextFile + 1, ","
'            Print #TextFile, MTickerArray(Ticker + EOList)
'            Print #TextFile, ","
'
'            Ticker = Ticker + 1
'        End If
'    Next k
'
'    Close TextFile
'    Close TextFile + 1

''''''''''
    ''This will iterate past the end of the array - need to figure out the right combination of upper bound,
    ''offset value, and array position.
    ''Once this is figured out, delete copy/paste, and possibly remove duplicates/MMDA lines below

'    For m = 0 To UBound(MTickerArray)
'        DupeStart.Offset(m, 0).Value = MNameArray(m + EOList)
'        DupeStart.Offset(m, 1).Value = MTickerArray(m + EOList)
'    Next m

    'Add and sort Morningstar tickers on portfolio
    Dim DupeMoney As Range

    TDSheet.Range(CopyPointStart.Offset(0, -1), CopyPointEnd(1, 0)).Copy
    Portfolio.Paste Destination:=DupeStart, Link:=False
    
    Set DupeEnd = Range(FirstStar.End(xlDown).Offset(0, 1).Address)
    DupeEndAddress = DupeEnd.Address

    Portfolio.Range(FirstStar.Address, DupeEnd.Address).RemoveDuplicates Columns:=2, Header:=xlNo
    
    'Delete stocks from Morningstar list
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
