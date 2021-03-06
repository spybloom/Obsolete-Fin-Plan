Attribute VB_Name = "Module4"
Sub Dist_Row()
Attribute Dist_Row.VB_ProcData.VB_Invoke_Func = "Q\n14"
'
' Dist_Row Macro
'
' Keyboard Shortcut: Ctrl+Shift+Q
'
' This macro is the same as the one below, except it adds a line directly below the selected cell instead of searching
'   for a starting point.
'
    Dim CellStart As Range
    Dim TodayDate As Date
    Dim djia As Range
    Dim Inputs() As String
    Dim Contribution As Variant
    Dim Withdrawal As Variant
    Dim Distribution As Variant
    Dim DJInput As Variant
    Dim SPInput As Variant
    Dim i As Integer
    Dim s1 As Range
    Dim s2 As Variant
    Dim s3 As Variant
    Dim j As Integer
    Dim ChangeBox As Range
    Dim Portfolio As Worksheet
    Dim PortTotal As Range
    Dim k As Integer
    
    StateToggle ("Off")
    
    Set CellStart = ActiveCell
    
    TodayDate = Date - 1
    ActiveSheet.PageSetup.RightHeader = TodayDate
    If Rows(1).Find("Date") Is Nothing Then
        MsgBox ("Date wasn't found on first row. This date hasn't been updated.")
    Else
        Rows(1).Find("Date").Offset(1, 0).Formula = TodayDate
    End If
    
    Set djia = Cells.Find("DJIA", After:=Cells.Range("A1"), LookIn:=xlValues)
    If djia Is Nothing Then
        MsgBox ("Macro has been halted. Please enter ""DJIA"" next to Dow Jones value")
        StateToggle ("On")
        End
    End If
    
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
    
    CellStart.Offset(1, 0).EntireRow.Insert
    
    'Date
    CellStart.Offset(1, 0).FormulaR1C1 = TodayDate
        
    'Present Value
    CellStart.Offset(1, 1).FormulaR1C1 = "=R[-1]C[5]"
    
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
    CellStart.Offset(1, 2) = Contribution
    CellStart.Offset(1, 3) = Withdrawal
    CellStart.Offset(1, 4) = Distribution
    
    'Adjusted
    CellStart.Offset(1, 5).FormulaR1C1 = "=R[0]C[-4]+R[0]C[-3]-R[0]C[-2]-R[0]C[-1]"
    
    'Present Value
    CellStart.Offset(1, 6).Value = CellStart.Offset(3, 6).Value
    CellStart.Offset(1, 6).NumberFormat = "_($* #,##0_);_($* (#,##0);_($* ""-""_);_(@_)"
    
    'Gains & Losses
    CellStart.Offset(1, 7).FormulaR1C1 = "=R[0]C[-1]-R[0]C[-2]"
    
    'Change
    CellStart.Offset(1, 8).FormulaR1C1 = "=R[0]C[-1]/R[0]C[-3]"
    
    'S&P500
    CellStart.Offset(1, 9).FormulaR1C1 = "=R[0]C[2]/R[-1]C[2]-1"
    
    'Better than S&P
    CellStart.Offset(1, 10).FormulaR1C1 = "=R[0]C[-2]-R[0]C[-1]"
    
    'S&P 500
    CellStart.Offset(1, 11).Formula = djia.Offset(1, 1).Formula
    
    'Performance
    CellStart.Offset(1, 12).FormulaR1C1 = "=R[-1]C[0]*(1+R[0]C[-4])"
    'Cell reference doesn't work if dist page is separated by years
    
    'Overall Change
    For i = 1 To 1000
        If CellStart.Offset(3, 8).Formula = "=M" & i & "-1" Then
            CellStart.Offset(3, 8).FormulaR1C1 = "=R[-2]C[4]-1"
        End If
    Next i
    
    'Overall S&P500
    Set s1 = Columns(1).Find("Date")
    If Not s1 Is Nothing Then
        s2 = s1.Offset(1, 11).Row
        s3 = s1.Offset(1, 11).Column
    End If

    CellStart.Offset(3, 9).FormulaR1C1 = "=R[-2]C[2]/R" & s2 & "C" & s3 & "-1"
    
    'Diagnostics
    For j = 0 To 12
        On Error GoTo ErrMsg
        If CellStart.Offset(1, j).Value = vbNullString Then
            MsgBox ("Macro completed with empty cells. Manually check output.")
        End If
    Next j

    If CellStart.Offset(1, 10).Value > 0.005 And CellStart.Offset(1, 9) > 0 Then
        MsgBox " Macro completed. Portfolio performed higher than S&P 500. Check numbers and re-run " _
            & "macro if necessary."
    ElseIf CellStart.Offset(1, 10).Value < 0 And CellStart.Offset(1, 9) < 0 Then
        MsgBox " Macro completed. Portfolio performed lower than S&P 500. Check numbers and re-run " _
            & "macro if necessary."
    End If
    
    Set ChangeBox = Range("A1:Z20").Find("Change")
    If ChangeBox Is Nothing Then
        MsgBox ("Net Change box should be labeled as ""Change"" one cell above it. Please add this or check numbers manually.")
    Else
        Set ChangeBox = ChangeBox.Offset(1, 0)
        If ChangeBox.Value > CellStart.Offset(1, 6).Value * 0.1 Then
            MsgBox ("Macro completed. Net change box may be too high. Check numbers and re-run macro if necessary")
        End If
    
        If ChangeBox.Value < CellStart.Offset(1, 6).Value * -0.1 Then
            MsgBox ("Macro completed. Net change box may be too low. Check numbers and re-run macro if necessary")
        End If
    End If
    
    StateToggle ("On")
    Exit Sub
    
ErrMsg:
    MsgBox ("Macro completed with error in output.")
        
    StateToggle ("On")
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

