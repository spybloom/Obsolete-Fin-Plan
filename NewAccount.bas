Attribute VB_Name = "Module2"
Sub New_Account()
Attribute New_Account.VB_ProcData.VB_Invoke_Func = "N\n14"
'
' Used to add all the cells and formatting needed for a new account on the portfolio page.
'
' Keyboard Shortcut: Ctrl+Shift+N
'
    Dim AcctType As String
    Dim AcctNumber As String
    Dim AcctName As String
    Dim StartCell As Range
    Dim InvestTotal As Range
    Dim InvestRow As Integer
    Dim InvestCol As Integer
    Dim ColNumberOne As Integer
    Dim ColNumberTwo As Integer
    Dim RowNumber As Integer
    
    On Error GoTo BackOn
    
    Application.CutCopyMode = False
    
    'Take inputs
    AcctType = ""
    AcctNumber = ""
    AcctName = ""
    
    AcctType = InputBox("What type of account is this?")
    AcctNumber = InputBox("What are the last three numbers of the account?")
    AcctName = InputBox("Who is the owner of the account? If there is only one person, put ""N/A"".")
    
    If AcctType = "" Or AcctNumber = "" Or AcctName = "" Then
        MsgBox ("Macro halted. One or more of the inputs was blank.")
        End
    End If
    
    'Insert rows for new account
    Set StartCell = ActiveCell
    
    Range(StartCell.EntireRow, StartCell.Offset(3, 0).EntireRow).Insert _
        Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    Set StartCell = StartCell.Offset(-4, 0)
    Set InvestTotal = Cells.Find("Total Investments:").Offset(0, 2)
    InvestRow = InvestTotal.Row
    InvestCol = InvestTotal.Column
    
    'Fill out necessary information
    With StartCell
        .Font.Bold = True
        
        If UCase(AcctName) <> "N/A" And UCase(AcctName) <> "NA" Then
            .FormulaR1C1 = "TD Ameritrade " & AcctType & " - " & AcctName
        Else
            .FormulaR1C1 = "TD Ameritrade - " & AcctType
        End If
        
        .Offset(1, 0).FormulaR1C1 = "Acct # xxx-xxx" & AcctNumber
        .Offset(2, 0).FormulaR1C1 = "TD Bank FDIC Money Market"
        .Offset(1, 1).FormulaR1C1 = "=SUM(R[1]C[2]:R[2]C[2],R[1]C[5]:R[2]C[5])"
        .Offset(1, 1).Font.Bold = True
        .Offset(1, 1).NumberFormat = "_($* #,##0_);_($* (#,##0);_($* ""-""??_);_(@_)"
        .Offset(2, 2).Formula = "mmda12"
        .Offset(2, 2).Font.ColorIndex = 2
        ColNumberOne = .Offset(2, 9).Column
        ColNumberTwo = .Offset(2, 10).Column
        RowNumber = .Offset(2, 0).Row
        .Offset(2, 3).FormulaR1C1 = "=VLOOKUP(RC[-1],R" & RowNumber & "C" & ColNumberOne & ":R" & RowNumber + 1 & "C" & ColNumberTwo & ",2,FALSE)"
        .Offset(2, 3).NumberFormat = "_($* #,##0_);_($* (#,##0);_($* ""-""??_);_(@_)"
        .Offset(3, 3).NumberFormat = "_($* #,##0_);_($* (#,##0);_($* ""-""??_);_(@_)"
        .Offset(2, 4).FormulaR1C1 = "=RC[-1]/R" & InvestRow & "C" & InvestCol
        .Offset(2, 4).NumberFormat = "0.00%"
        .Offset(2, 6).NumberFormat = "_($* #,##0_);_($* (#,##0);_($* ""-""??_);_(@_)"
        .Offset(3, 6).NumberFormat = "_($* #,##0_);_($* (#,##0);_($* ""-""??_);_(@_)"
        .Offset(2, 7).NumberFormat = "0.00%"
        .Offset(2, 9).Formula = "MMDA12"
        .Offset(2, 10).Formula = "0.00"
        .Offset(2, 10).NumberFormat = "#,##0.00"
        .Offset(3, 10).FormulaR1C1 = "=SUM(R[-1]C)"
        .Offset(3, 10).NumberFormat = "#,##0.00"
    End With
    
    'Make yellow cells
    With Range(StartCell.Offset(1, 9), StartCell.Offset(1, 10)).Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ColorIndex = 6
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    
    Exit Sub
    
BackOn:
    MsgBox ("Macro ended prematurely due to error in execution.")
End Sub
