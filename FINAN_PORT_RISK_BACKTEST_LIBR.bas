Attribute VB_Name = "FINAN_PORT_RISK_BACKTEST_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.


'************************************************************************************
'************************************************************************************
'FUNCTION      : RNG_PORT_BACK_TESTING_FUNC
'DESCRIPTION   : Portfolio Back Testing
'LIBRARY       : PORTFOLIO
'GROUP         : BACK_TESTING
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/08/2008
'REFERENCE:
'http://www.bogleheads.org/forum/viewtopic.php?p=403843
'************************************************************************************
'************************************************************************************

Function RNG_PORT_BACK_TESTING_FUNC(ByRef DST_RNG As Excel.Range, _
Optional ByVal NO_PERIODS As Long = 37, _
Optional ByVal NO_ASSETS As Long = 28, _
Optional ByVal NO_PORTFOLIOS As Long = 5)

'NO_ASSETS = INCLUDING INFLATION

If NO_PERIODS < 3 Then: GoTo ERROR_LABEL
If NO_ASSETS < 2 Then: GoTo ERROR_LABEL
If NO_PORTFOLIOS < 1 Then: GoTo ERROR_LABEL

Dim h As Long
Dim i As Long
Dim j As Long
Dim l(1 To 7) As Long

Dim POS_RNG As Excel.Range
Dim IND_RNG As Excel.Range

Dim START_RNG As Excel.Range
Dim END_RNG As Excel.Range
Dim TMP_RNG As Excel.Range
Dim SRC_RNG As Excel.Range

On Error GoTo ERROR_LABEL

RNG_PORT_BACK_TESTING_FUNC = False

'-----------------------------------------------------------------------------------------
'FIRST PASS: HISTORICAL RETURNS
'-----------------------------------------------------------------------------------------
l(1) = ((((((0 + 17 + 2) + 6 + NO_PORTFOLIOS + 2) + 18 + 2) + _
    2 + NO_PERIODS + 2) + 17 + 2) + 2 + NO_ASSETS + 2)

Set SRC_RNG = DST_RNG.Cells(l(1), 1)
Set SRC_RNG = Range(SRC_RNG.Cells(1), SRC_RNG.Cells(5 + NO_PERIODS, 1 + NO_ASSETS))
With SRC_RNG
    Set TMP_RNG = SRC_RNG
    GoSub FONT_LINE
    .Cells(1, 1) = "LONG_DESCRIPTION"
    .Cells(2, 1) = "INDEX"
    .Cells(3, 1) = "SHORT_DESCRIPTION"
    .Cells(4, 1) = "SYMBOL"

    j = Year(Now)
    For i = 1 To NO_PERIODS
        .Cells(5 + i, 1).value = j - NO_PERIODS + i
    Next i

    For j = 1 To NO_ASSETS
        .Cells(1, 1 + j).value = "LONG_DESCR_" & j
        .Cells(2, 1 + j).value = j
        .Cells(3, 1 + j).value = "SHORT_DESCR_" & j
        .Cells(4, 1 + j).value = "SYMBOL_" & j
    
        For i = 1 To NO_PERIODS
            .Cells(5 + i, 1 + j).value = 0
        Next i
    Next j
    .Cells(1, NO_ASSETS + 1).value = "INFLATION_RATE"
    .Cells(3, NO_ASSETS + 1).value = "INFLATION_DATA_LABEL"
    .Cells(4, NO_ASSETS + 1).value = "INFLATION_SYMBOL"
    
    h = 35
    Set TMP_RNG = Range(SRC_RNG.Rows(1), SRC_RNG.Rows(4))
    GoSub BORDER_LINE
    h = -16776961
    Set TMP_RNG = Range(SRC_RNG.Rows(1).Cells(1, 2), SRC_RNG.Rows(1))
    TMP_RNG.Font.Color = h
    TMP_RNG.Cells(1, 1).Font.Color = 0
    h = -4165632
    Set TMP_RNG = Range(SRC_RNG.Rows(3), SRC_RNG.Rows(4))
    TMP_RNG.Font.Color = h
    TMP_RNG.Cells(1, 1).Font.Color = 0
    TMP_RNG.Cells(2, 1).Font.Color = 0
    Set TMP_RNG = Range(SRC_RNG.Rows(6), SRC_RNG.Rows(6 + NO_PERIODS - 1))
    TMP_RNG.Font.Color = h
    TMP_RNG.NumberFormat = "0.00%"
    TMP_RNG.Columns(1).NumberFormat = "0"


End With
'-----------------------------------------------------------------------------------------
'SECOND PASS: CONTROL PANEL
'-----------------------------------------------------------------------------------------
l(2) = 1

Set SRC_RNG = DST_RNG.Cells(l(2), 1)
Set SRC_RNG = Range(SRC_RNG.Cells(1, 1), SRC_RNG.Cells(17, 2))

With SRC_RNG
    Set TMP_RNG = SRC_RNG
    GoSub FONT_LINE

'-----------------------------------------------------------------------------------------
    .Cells(1, 1) = "INITIAL INVESTMENT"
    .Cells(1, 2) = 10000
    
    Set TMP_RNG = .Cells(1, 1)
    h = 35
    GoSub BORDER_LINE
    
    Set TMP_RNG = .Cells(1, 2)
    h = 36
    GoSub BORDER_LINE
    TMP_RNG.Font.Color = -4165632
    
'-----------------------------------------------------------------------------------------
    .Cells(3, 1) = "STARTING PERIOD FOR BACKTESTING"
    Set TMP_RNG = DST_RNG.Cells(l(1), 1)
    Set TMP_RNG = Range(TMP_RNG.Cells(6, 1), TMP_RNG.Cells(5 + NO_PERIODS, 1))
    .Cells(4, 1).Validation.Delete
    .Cells(4, 1).Validation.Add Type:=xlValidateList, _
        AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:="=" & TMP_RNG.Address
    .Cells(4, 2).formula = "=MATCH(" & .Cells(4, 1).Address & _
                            "," & TMP_RNG.Address & ",FALSE)"
    
    Set TMP_RNG = .Cells(3, 1)
    h = 35
    GoSub BORDER_LINE
    
    Set TMP_RNG = .Cells(4, 1)
    h = 36
    GoSub BORDER_LINE
    TMP_RNG.Font.Color = -16776961

    Set TMP_RNG = .Cells(4, 2)
    h = 37
    GoSub BORDER_LINE


'-----------------------------------------------------------------------------------------
    .Cells(5, 1) = "ENDING PERIOD FOR BACKTESTING"
    Set TMP_RNG = DST_RNG.Cells(l(1), 1)
    Set TMP_RNG = Range(TMP_RNG.Cells(6, 1), TMP_RNG.Cells(5 + NO_PERIODS, 1))
    .Cells(6, 1).Validation.Delete
    .Cells(6, 1).Validation.Add Type:=xlValidateList, _
        AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:="=" & TMP_RNG.Address
    .Cells(6, 2).formula = "=MATCH(" & .Cells(6, 1).Address & _
                            "," & TMP_RNG.Address & ",FALSE)"
    

    Set TMP_RNG = .Cells(5, 1)
    h = 35
    GoSub BORDER_LINE
    
    Set TMP_RNG = .Cells(6, 1)
    h = 36
    GoSub BORDER_LINE
    TMP_RNG.Font.Color = -16776961

    Set TMP_RNG = .Cells(6, 2)
    h = 37
    GoSub BORDER_LINE

'-----------------------------------------------------------------------------------------
    .Cells(8, 1) = "CASH RATE"
    Set TMP_RNG = DST_RNG.Cells(l(1), 1)
    Set TMP_RNG = Range(TMP_RNG.Cells(4, 2), TMP_RNG.Cells(4, 1 + NO_ASSETS))
    .Cells(9, 1).Validation.Delete
    .Cells(9, 1).Validation.Add Type:=xlValidateList, _
        AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:="=" & TMP_RNG.Address
    .Cells(9, 2).formula = "=MATCH(" & .Cells(9, 1).Address & _
                            "," & TMP_RNG.Address & ",FALSE)"
    
    Set TMP_RNG = .Cells(8, 1)
    h = 35
    GoSub BORDER_LINE
    
    Set TMP_RNG = .Cells(9, 1)
    h = 36
    GoSub BORDER_LINE
    TMP_RNG.Font.Color = -16776961

    Set TMP_RNG = .Cells(9, 2)
    h = 37
    GoSub BORDER_LINE
'-----------------------------------------------------------------------------------------
    .Cells(10, 1) = "INFLATION RATE"
    Set TMP_RNG = DST_RNG.Cells(l(1), 1)
    Set TMP_RNG = Range(TMP_RNG.Cells(4, 2), TMP_RNG.Cells(4, 1 + NO_ASSETS))
    .Cells(11, 1).Validation.Delete
    .Cells(11, 1).Validation.Add Type:=xlValidateList, _
        AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:="=" & TMP_RNG.Address
    .Cells(11, 2).formula = "=MATCH(" & .Cells(11, 1).Address & _
                            "," & TMP_RNG.Address & ",FALSE)"
    .Cells(11, 1).value = "INFLATION_SYMBOL"
    
    Set TMP_RNG = .Cells(10, 1)
    h = 35
    GoSub BORDER_LINE
    
    Set TMP_RNG = .Cells(11, 1)
    h = 36
    GoSub BORDER_LINE
    TMP_RNG.Font.Color = -16776961

    Set TMP_RNG = .Cells(11, 2)
    h = 37
    GoSub BORDER_LINE
'-----------------------------------------------------------------------------------------
    .Cells(12, 1) = "HOME BENCHMARK"
    Set TMP_RNG = DST_RNG.Cells(l(1), 1)
    Set TMP_RNG = Range(TMP_RNG.Cells(4, 2), TMP_RNG.Cells(4, 1 + NO_ASSETS))
    .Cells(13, 1).Validation.Delete
    .Cells(13, 1).Validation.Add Type:=xlValidateList, _
        AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:="=" & TMP_RNG.Address
    .Cells(13, 2).formula = "=MATCH(" & .Cells(13, 1).Address & _
                            "," & TMP_RNG.Address & ",FALSE)"
    
    Set TMP_RNG = .Cells(12, 1)
    h = 35
    GoSub BORDER_LINE
    
    Set TMP_RNG = .Cells(13, 1)
    h = 36
    GoSub BORDER_LINE
    TMP_RNG.Font.Color = -16776961

    Set TMP_RNG = .Cells(13, 2)
    h = 37
    GoSub BORDER_LINE
'-----------------------------------------------------------------------------------------
    .Cells(14, 1) = "FOREIGNER BENCHMARK"
    Set TMP_RNG = DST_RNG.Cells(l(1), 1)
    Set TMP_RNG = Range(TMP_RNG.Cells(4, 2), TMP_RNG.Cells(4, 1 + NO_ASSETS))
    .Cells(15, 1).Validation.Delete
    .Cells(15, 1).Validation.Add Type:=xlValidateList, _
        AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:="=" & TMP_RNG.Address
    .Cells(15, 2).formula = "=MATCH(" & .Cells(15, 1).Address & _
                            "," & TMP_RNG.Address & ",FALSE)"
    
    Set TMP_RNG = .Cells(14, 1)
    h = 35
    GoSub BORDER_LINE
    
    Set TMP_RNG = .Cells(15, 1)
    h = 36
    GoSub BORDER_LINE
    TMP_RNG.Font.Color = -16776961

    Set TMP_RNG = .Cells(15, 2)
    h = 37
    GoSub BORDER_LINE
'-----------------------------------------------------------------------------------------
    .Cells(17, 1) = "FACTOR"
    .Cells(17, 2) = 1

    Set TMP_RNG = .Cells(17, 1)
    h = 35
    GoSub BORDER_LINE
    
    Set TMP_RNG = .Cells(17, 2)
    h = 36
    GoSub BORDER_LINE
    TMP_RNG.Font.Color = -4165632

End With

'-----------------------------------------------------------------------------------------
'THIRD PASS: PORTFOLIO WEIGHTS
'-----------------------------------------------------------------------------------------
l(3) = l(2) + 17 + 2

Set SRC_RNG = DST_RNG.Cells(l(3), 1)
Set SRC_RNG = Range(SRC_RNG.Cells(1), SRC_RNG.Cells(6 + NO_PORTFOLIOS, NO_ASSETS))
With SRC_RNG
    Set TMP_RNG = SRC_RNG
    GoSub FONT_LINE
    .Cells(1, 1) = "LONG_DESCRIPTION"
    .Cells(2, 1) = "SHORT_DESCRIPTION"
    .Cells(3, 1) = "SYMBOL"
    .Cells(5, 1) = "ALLOCATION"
    
    .Cells(6, 1) = "BEST/OPTIMAL"
    .Cells(6, 1 + NO_ASSETS).formula = "=SUM(" & Range(.Cells(6, 2), _
                                                       .Cells(6, NO_ASSETS)).Address & ")"

    For i = 1 To NO_PORTFOLIOS
        .Cells(6 + i, 1).value = "PORT_" & i
        .Cells(6 + i, 1 + NO_ASSETS).formula = "=SUM(" & Range(.Cells(6 + i, 2), _
                                                       .Cells(6 + i, NO_ASSETS)).Address & ")"
    Next i

    For j = 1 To NO_ASSETS - 1
        .Cells(1, 1 + j).formula = "=" & DST_RNG.Cells(l(1) + 0, 1 + j).Address
        .Cells(2, 1 + j).formula = "=" & DST_RNG.Cells(l(1) + 2, 1 + j).Address
        .Cells(3, 1 + j).formula = "=" & DST_RNG.Cells(l(1) + 3, 1 + j).Address
    
        .Cells(6, 1 + j).value = 0
        For i = 1 To NO_PORTFOLIOS
            .Cells(6 + i, 1 + j).value = 0
        Next i
    Next j
    h = 35
    Set TMP_RNG = Range(SRC_RNG.Rows(1), SRC_RNG.Rows(3))
    GoSub BORDER_LINE
    Set TMP_RNG = Range(SRC_RNG.Rows(6), SRC_RNG.Rows(6))
    GoSub BORDER_LINE
    
    h = 36
    Set TMP_RNG = Range(SRC_RNG.Rows(7), SRC_RNG.Rows(6 + NO_PORTFOLIOS))
    GoSub BORDER_LINE
    
    Set TMP_RNG = Range(SRC_RNG.Cells(6, 2), SRC_RNG.Cells(6 + NO_PORTFOLIOS, NO_ASSETS))
    TMP_RNG.Font.Color = -4165632
    TMP_RNG.NumberFormat = "0.00%"
    
    Set TMP_RNG = Range(SRC_RNG.Cells(6, 1), SRC_RNG.Cells(6 + NO_PORTFOLIOS, 1))
    TMP_RNG.Font.Color = -16776961
    
    
    Set TMP_RNG = .Cells(5, 1)
    h = 37
    GoSub BORDER_LINE
    
    Set TMP_RNG = Range(.Cells(6, NO_ASSETS + 1), .Cells(6 + NO_PORTFOLIOS, NO_ASSETS + 1))
    h = 37
    GoSub BORDER_LINE
    TMP_RNG.Font.Color = -16776961
    TMP_RNG.NumberFormat = "0.00%"

End With


'-----------------------------------------------------------------------------------------
'FORTH PASS: PEARSON TABLE
'NO_ASSETS - 1 ; Exclude Inflation in the Last Column of the Return Data
'-----------------------------------------------------------------------------------------
l(4) = l(1) - (2 + (NO_ASSETS - 1) + 2)

Set SRC_RNG = DST_RNG.Cells(l(4), 1)
Set SRC_RNG = Range(SRC_RNG.Cells(1, 1), SRC_RNG.Cells(2 + NO_ASSETS - 1, 1 + NO_ASSETS - 1))

With SRC_RNG
    Set TMP_RNG = SRC_RNG
    GoSub FONT_LINE
    .Cells(1, 1) = "PEARSON CORRELATIONS"

    For i = 1 To NO_ASSETS - 1
        .Cells(1, 1 + i).formula = "=" & DST_RNG.Cells(l(1) + 3, 1 + i).Address
        .Cells(2 + i, 1).formula = "=" & .Cells(1, 1 + i).Address
    Next i

    h = 35
    Set TMP_RNG = Range(SRC_RNG.Rows(1), SRC_RNG.Rows(1))
    GoSub BORDER_LINE

End With

Set POS_RNG = DST_RNG.Cells(l(1) + 5, 1)
Set TMP_RNG = SRC_RNG.Cells(3, 2)
Set IND_RNG = DST_RNG.Cells(l(1) + 1, 2)
Set START_RNG = DST_RNG.Cells(l(2) + 3, 2)
Set END_RNG = DST_RNG.Cells(l(2) + 5, 2)

For j = 1 To NO_ASSETS - 1
    For i = j To NO_ASSETS - 1
        TMP_RNG.Cells(i, j).formula = "=PEARSON(" & _
                    "OFFSET(" & POS_RNG.Address & "," & _
                          START_RNG.Address & "-1," & _
                          IND_RNG.Offset(0, i - 1).Address & "," & _
                          END_RNG.Address & "+1-" & _
                          START_RNG.Address & ",1)" & "," & _
                    "OFFSET(" & POS_RNG.Address & "," & _
                          START_RNG.Address & "-1," & _
                          IND_RNG.Offset(0, j - 1).Address & "," & _
                          END_RNG.Address & "+1-" & _
                          START_RNG.Address & ",1))"
    Next i
Next j

Set TMP_RNG = Range(TMP_RNG.Cells(1, 1), TMP_RNG.Cells(NO_ASSETS - 1, NO_ASSETS - 1))
TMP_RNG.NumberFormat = "0.0000"

'-----------------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------------
'FIFTH PASS: ASSETS STATISTICS
'-----------------------------------------------------------------------------------------
l(5) = l(4) - (2 + 15 + 2)

Set SRC_RNG = DST_RNG.Cells(l(5), 1)
Set SRC_RNG = Range(SRC_RNG.Cells(1, 1), SRC_RNG.Cells(2 + 15, 1 + NO_ASSETS))

With SRC_RNG
    Set TMP_RNG = SRC_RNG
    GoSub FONT_LINE
    .Cells(1, 1) = "ASSETS STATISTICS"

    For i = 1 To NO_ASSETS
        .Cells(1, 1 + i).formula = "=" & DST_RNG.Cells(l(1) + 3, 1 + i).Address
    Next i

    h = 35
    Set TMP_RNG = Range(SRC_RNG.Rows(1), SRC_RNG.Rows(1))
    GoSub BORDER_LINE

    .Cells(3, 1) = "Average"
    .Cells(4, 1) = "Std. Dev."
    .Cells(5, 1) = "Down Sigma"
    .Cells(6, 1) = "Up Sigma"
    .Cells(7, 1) = "CAGR"
    .Cells(8, 1) = "Variance"
    .Cells(9, 1).formula = _
        "=" & """" & "Pearson w/" & """" & "&" & DST_RNG.Cells(l(2) + 12, 1).Address
    .Cells(10, 1).formula = _
        "=" & """" & "Pearson w/" & """" & "&" & DST_RNG.Cells(l(2) + 14, 1).Address
    .Cells(11, 1) = "Sharpe Ratio"
    .Cells(12, 1) = "Sortino"
    .Cells(13, 1) = "Skew"
    .Cells(14, 1) = "Kurtosis"
    .Cells(15, 1) = "$1 Asset Growth - Nominal"
    .Cells(16, 1) = "$1 Asset Growth - Real"
    .Cells(17, 1) = "Best Performance Asset"

    .Cells(17, 2).formula = "=INDEX(" & _
                    Range(.Cells(1, 2), .Cells(1, 1 + NO_ASSETS)).Address & ",MATCH(MAX(" & _
                    Range(.Cells(7, 2), .Cells(7, 1 + NO_ASSETS)).Address & ")," & _
                    Range(.Cells(7, 2), .Cells(7, 1 + NO_ASSETS)).Address & ",0))"
    
    Set TMP_RNG = .Cells(17, 2)
    h = 36
    GoSub BORDER_LINE
    
    .Cells(17, 3).formula = "=INDEX(" & _
                            Range(DST_RNG.Cells(l(3), 2), _
                                  DST_RNG.Cells(l(3), 1 + NO_ASSETS)).Address & ",1,MATCH(" & _
                            .Cells(17, 2).Address & _
                            "," & Range(DST_RNG.Cells(l(3) + 2, 2), _
                                  DST_RNG.Cells(l(3) + 2, 1 + NO_ASSETS)).Address & ",FALSE))"
    
    Set TMP_RNG = Range(.Cells(17, 3), .Cells(17, 4))
    h = 36
    GoSub BORDER_LINE
    
    TMP_RNG.HorizontalAlignment = xlCenterAcrossSelection

    For j = 1 To NO_ASSETS
        .Cells(3, 1 + j).formula = "=AVERAGE(OFFSET(" & _
                    DST_RNG.Cells(l(1) + 5, 1).Address & "," & _
                    DST_RNG.Cells(l(2) + 3, 2).Address & "-1," & _
                    DST_RNG.Cells(l(1) + 1, 1).Offset(0, j).Address & "," & _
                    DST_RNG.Cells(l(2) + 5, 2).Address & "+1-" & _
                    DST_RNG.Cells(l(2) + 3, 2).Address & ",1))"
        
        .Cells(4, 1 + j).formula = "=STDEV(OFFSET(" & _
                            DST_RNG.Cells(l(1) + 5, 1).Address & "," & _
                            DST_RNG.Cells(l(2) + 3, 2).Address & "-1," & _
                            DST_RNG.Cells(l(1) + 1, 1).Offset(0, j).Address & "," & _
                            DST_RNG.Cells(l(2) + 5, 2).Address & "+1-" & _
                            DST_RNG.Cells(l(2) + 3, 2).Address & ",1))"
        
        .Cells(5, 1 + j).FormulaArray = "=STDEVP(IF(OFFSET(" & _
                            DST_RNG.Cells(l(1) + 5, 1).Address & "," & _
                            DST_RNG.Cells(l(2) + 3, 2).Address & "-1," & _
                            DST_RNG.Cells(l(1) + 1, 1).Offset(0, j).Address & "," & _
                            DST_RNG.Cells(l(2) + 5, 2).Address & "+1-" & _
                            DST_RNG.Cells(l(2) + 3, 2).Address & ",1)<OFFSET(" & _
                            .Cells(3, 1).Address & ",0," & _
                            DST_RNG.Cells(l(2) + 8, 2).Address & "),OFFSET(" & _
                            DST_RNG.Cells(l(1) + 5, 1).Address & "," & _
                            DST_RNG.Cells(l(2) + 3, 2).Address & "-1," & _
                            DST_RNG.Cells(l(1) + 1, 1).Offset(0, j).Address & "," & _
                            DST_RNG.Cells(l(2) + 5, 2).Address & "+1-" & _
                            DST_RNG.Cells(l(2) + 3, 2).Address & ",1)))"
        
        .Cells(6, 1 + j).FormulaArray = "=STDEVP(IF(OFFSET(" & _
                            DST_RNG.Cells(l(1) + 5, 1).Address & "," & _
                            DST_RNG.Cells(l(2) + 3, 2).Address & "-1," & _
                            DST_RNG.Cells(l(1) + 1, 1).Offset(0, j).Address & "," & _
                            DST_RNG.Cells(l(2) + 5, 2).Address & "+1-" & _
                            DST_RNG.Cells(l(2) + 3, 2).Address & ",1)>OFFSET(" & _
                            .Cells(3, 1).Address & ",0," & _
                            DST_RNG.Cells(l(2) + 8, 2).Address & "),OFFSET(" & _
                            DST_RNG.Cells(l(1) + 5, 1).Address & "," & _
                            DST_RNG.Cells(l(2) + 3, 2).Address & "-1," & _
                            DST_RNG.Cells(l(1) + 1, 1).Offset(0, j).Address & "," & _
                            DST_RNG.Cells(l(2) + 5, 2).Address & "+1-" & _
                            DST_RNG.Cells(l(2) + 3, 2).Address & ",1)))"
        
        .Cells(7, 1 + j).FormulaArray = "=(PRODUCT(1+OFFSET(" & _
                            DST_RNG.Cells(l(1) + 5, 1).Address & "," & _
                            DST_RNG.Cells(l(2) + 3, 2).Address & "-1," & _
                            DST_RNG.Cells(l(1) + 1, 1).Offset(0, j).Address & "," & _
                            DST_RNG.Cells(l(2) + 5, 2).Address & "+1-" & _
                            DST_RNG.Cells(l(2) + 3, 2).Address & ",1)/" & _
                            DST_RNG.Cells(l(2) + 16, 2).Address & ")^(1/COUNT(OFFSET(" & _
                            DST_RNG.Cells(l(1) + 5, 1).Address & "," & _
                            DST_RNG.Cells(l(2) + 3, 2).Address & "-1," & _
                            DST_RNG.Cells(l(1) + 1, 1).Offset(0, j).Address & "," & _
                            DST_RNG.Cells(l(2) + 5, 2).Address & "+1-" & _
                            DST_RNG.Cells(l(2) + 3, 2).Address & ",1)))-1)*" & _
                            DST_RNG.Cells(l(2) + 16, 2).Address & ""
        
        .Cells(8, 1 + j).formula = "=VARA(OFFSET(" & _
                            DST_RNG.Cells(l(1) + 5, 1).Address & "," & _
                            DST_RNG.Cells(l(2) + 3, 2).Address & "-1," & _
                            DST_RNG.Cells(l(1) + 1, 1).Offset(0, j).Address & "," & _
                            DST_RNG.Cells(l(2) + 5, 2).Address & "+1-" & _
                            DST_RNG.Cells(l(2) + 3, 2).Address & ",1))"
        
        .Cells(9, 1 + j).formula = "=PEARSON(OFFSET(" & _
                            DST_RNG.Cells(l(1) + 5, 1).Address & "," & _
                            DST_RNG.Cells(l(2) + 3, 2).Address & "-1," & _
                            DST_RNG.Cells(l(1) + 1, 1).Offset(0, j).Address & "," & _
                            DST_RNG.Cells(l(2) + 5, 2).Address & "+1-" & _
                            DST_RNG.Cells(l(2) + 3, 2).Address & ",1),OFFSET(" & _
                            DST_RNG.Cells(l(1) + 5, 1).Address & "," & _
                            DST_RNG.Cells(l(2) + 3, 2).Address & "-1," & _
                            DST_RNG.Cells(l(2) + 12, 2).Address & "," & _
                            DST_RNG.Cells(l(2) + 5, 2).Address & "+1-" & _
                            DST_RNG.Cells(l(2) + 3, 2).Address & ",1))"
        
        .Cells(10, 1 + j).formula = "=PEARSON(OFFSET(" & _
                            DST_RNG.Cells(l(1) + 5, 1).Address & "," & _
                            DST_RNG.Cells(l(2) + 3, 2).Address & "-1," & _
                            DST_RNG.Cells(l(1) + 1, 1).Offset(0, j).Address & "," & _
                            DST_RNG.Cells(l(2) + 5, 2).Address & "+1-" & _
                            DST_RNG.Cells(l(2) + 3, 2).Address & ",1),OFFSET(" & _
                            DST_RNG.Cells(l(1) + 5, 1).Address & "," & _
                            DST_RNG.Cells(l(2) + 3, 2).Address & "-1," & _
                            DST_RNG.Cells(l(2) + 14, 2).Address & "," & _
                            DST_RNG.Cells(l(2) + 5, 2).Address & "+1-" & _
                            DST_RNG.Cells(l(2) + 3, 2).Address & ",1))"
        
        .Cells(11, 1 + j).formula = "=(" & _
                            .Cells(3, 1).Offset(0, j).Address & "-(OFFSET(" & _
                            .Cells(3, 1).Address & ",0," & _
                            DST_RNG.Cells(l(2) + 8, 2).Address & ")))/" & _
                            .Cells(4, 1).Offset(0, j).Address & ""
        .Cells(12, 1 + j).formula = "=(" & _
                            .Cells(3, 1).Offset(0, j).Address & "-(OFFSET(" & _
                            .Cells(3, 1).Address & ",0," & _
                            DST_RNG.Cells(l(2) + 8, 2).Address & ")))/" & _
                            .Cells(5, 1).Offset(0, j).Address & ""
        
        .Cells(13, 1 + j).formula = "=SKEW(OFFSET(" & _
                            DST_RNG.Cells(l(1) + 5, 1).Address & "," & _
                            DST_RNG.Cells(l(2) + 3, 2).Address & "-1," & _
                            DST_RNG.Cells(l(1) + 1, 1).Offset(0, j).Address & "," & _
                            DST_RNG.Cells(l(2) + 5, 2).Address & "+1-" & _
                            DST_RNG.Cells(l(2) + 3, 2).Address & ",1))"
        .Cells(14, 1 + j).formula = "=KURT(OFFSET(" & _
                            DST_RNG.Cells(l(1) + 5, 1).Address & "," & _
                            DST_RNG.Cells(l(2) + 3, 2).Address & "-1," & _
                            DST_RNG.Cells(l(1) + 1, 1).Offset(0, j).Address & "," & _
                            DST_RNG.Cells(l(2) + 5, 2).Address & "+1-" & _
                            DST_RNG.Cells(l(2) + 3, 2).Address & ",1))"
        
        .Cells(15, 1 + j).FormulaArray = "=PRODUCT(1+OFFSET(" & _
                            DST_RNG.Cells(l(1) + 5, 1).Address & "," & _
                            DST_RNG.Cells(l(2) + 3, 2).Address & "-1," & _
                            DST_RNG.Cells(l(1) + 1, 1).Offset(0, j).Address & "," & _
                            DST_RNG.Cells(l(2) + 5, 2).Address & "+1-" & _
                            DST_RNG.Cells(l(2) + 3, 2).Address & ",1)/" & _
                            DST_RNG.Cells(l(2) + 16, 2).Address & ")"
        
        .Cells(16, 1 + j).FormulaArray = "=PRODUCT(1+((OFFSET(" & _
                            DST_RNG.Cells(l(1) + 5, 1).Address & "," & _
                            DST_RNG.Cells(l(2) + 3, 2).Address & "-1," & _
                            DST_RNG.Cells(l(1) + 1, 1).Offset(0, j).Address & "," & _
                            DST_RNG.Cells(l(2) + 5, 2).Address & "+1-" & _
                            DST_RNG.Cells(l(2) + 3, 2).Address & ",1)/" & _
                            DST_RNG.Cells(l(2) + 16, 2).Address & "-OFFSET(" & _
                            DST_RNG.Cells(l(1) + 5, 1).Address & "," & _
                            DST_RNG.Cells(l(2) + 3, 2).Address & "-1," & _
                            DST_RNG.Cells(l(2) + 10, 2).Address & "," & _
                            DST_RNG.Cells(l(2) + 5, 2).Address & "+1-" & _
                            DST_RNG.Cells(l(2) + 3, 2).Address & ",1)/" & _
                            DST_RNG.Cells(l(2) + 16, 2).Address & ")/(1+OFFSET(" & _
                            DST_RNG.Cells(l(1) + 5, 1).Address & "," & _
                            DST_RNG.Cells(l(2) + 3, 2).Address & "-1," & _
                            DST_RNG.Cells(l(2) + 10, 2).Address & "," & _
                            DST_RNG.Cells(l(2) + 5, 2).Address & "+1-" & _
                            DST_RNG.Cells(l(2) + 3, 2).Address & ",1)/" & _
                            DST_RNG.Cells(l(2) + 16, 2).Address & ")))"
    Next j

End With


Set TMP_RNG = Range(SRC_RNG.Cells(3, 2), SRC_RNG.Cells(16, 1 + NO_ASSETS))
TMP_RNG.NumberFormat = "0.0000"

'-----------------------------------------------------------------------------------------
'SIXTH PASS: WEIGHTED PORTFOLIOS RETURNS
'-----------------------------------------------------------------------------------------
l(6) = l(5) - (2 + NO_PERIODS + 2)

Set SRC_RNG = DST_RNG.Cells(l(6), 1)
Set SRC_RNG = Range(SRC_RNG.Cells(1, 1), SRC_RNG.Cells(2 + NO_PERIODS, 2 + NO_PORTFOLIOS))

With SRC_RNG
    Set TMP_RNG = SRC_RNG
    GoSub FONT_LINE
    .Cells(1, 1) = "WEIGHTED PORTFOLIOS RETURNS"
    .Cells(1, 2).formula = "=" & DST_RNG.Cells(l(3) + 5, 1).Address

    For i = 1 To NO_PORTFOLIOS
        .Cells(1, 2 + i).formula = "=" & DST_RNG.Cells(l(3) + 5 + i, 1).Address
    Next i
    For i = 1 To NO_PERIODS
        .Cells(2 + i, 1).formula = "=" & DST_RNG.Cells(l(1) + 4 + i, 1).Address
    Next i

    h = 35
    Set TMP_RNG = Range(SRC_RNG.Rows(1), SRC_RNG.Rows(1))
    GoSub BORDER_LINE

End With

Set TMP_RNG = SRC_RNG.Cells(3, 2)
For j = 1 To NO_PORTFOLIOS + 1
    For i = 1 To NO_PERIODS
        'Remember to Exclude Inflation from Weights
        TMP_RNG.Cells(i, j).formula = "=SUMPRODUCT(" & _
                            Range(DST_RNG.Cells(l(3) + 4 + j, 2), _
                                DST_RNG.Cells(l(3) + 4 + j, NO_ASSETS)).Address & "," & _
                            Range(DST_RNG.Cells(l(1) + 4 + i, 2), _
                                DST_RNG.Cells(l(1) + 4 + i, NO_ASSETS)).Address & ")/" & _
                                DST_RNG.Cells(l(2) + 16, 2).Address & ""
    Next i
Next j

Set TMP_RNG = Range(TMP_RNG.Cells(1, 1), TMP_RNG.Cells(NO_PERIODS, NO_PORTFOLIOS + 1))
TMP_RNG.NumberFormat = "0.0000%"

'-----------------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------------
'SEVENTH STEP: PORTFOLIO STATISTICS
'-----------------------------------------------------------------------------------------
l(7) = l(6) - (2 + 16 + 2)

Set SRC_RNG = DST_RNG.Cells(l(7), 1)
Set SRC_RNG = Range(SRC_RNG.Cells(1, 1), SRC_RNG.Cells(18, 2 + NO_PORTFOLIOS))

With SRC_RNG
    Set TMP_RNG = SRC_RNG
    GoSub FONT_LINE
    .Cells(1, 1) = "PORTFOLIOS STATISTICS"

    For i = 1 To NO_PORTFOLIOS + 1
        .Cells(1, 1 + i).formula = "=" & DST_RNG.Cells(l(3) + 4 + i, 1).Address
    Next i

    h = 35
    Set TMP_RNG = Range(SRC_RNG.Rows(1), SRC_RNG.Rows(1))
    GoSub BORDER_LINE

    .Cells(3, 1) = "Average"
    .Cells(4, 1) = "Std. Dev."
    .Cells(5, 1) = "Down Sigma"
    .Cells(6, 1) = "Up Sigma"
    .Cells(7, 1) = "CAGR"
    .Cells(8, 1) = "Variance"
    .Cells(9, 1).formula = _
        "=" & """" & "Pearson w/" & """" & "&" & DST_RNG.Cells(l(2) + 12, 1).Address
    .Cells(10, 1).formula = _
        "=" & """" & "Pearson w/" & """" & "&" & DST_RNG.Cells(l(2) + 14, 1).Address
    .Cells(11, 1) = "Sharpe Ratio"
    .Cells(12, 1) = "Sortino"
    .Cells(13, 1) = "Skew"
    .Cells(14, 1) = "Kurtosis"
    
    .Cells(15, 1) = "Total: Rebalanced (Nominal)"
    .Cells(16, 1) = "Total: Un -Rebalanced(Nominal)"
    .Cells(17, 1) = "Total: Rebalanced (Real)"
    .Cells(18, 1) = "Total: Un -Rebalanced(Real)"

    For j = 1 To NO_PORTFOLIOS + 1
    
        .Cells(3, 1 + j).formula = "=AVERAGE(OFFSET(" & _
                    DST_RNG.Cells(l(6) + 2, 1).Offset(0, j).Address & "," & _
                    DST_RNG.Cells(l(2) + 3, 2).Address & "-1," & _
                    0 & "," & _
                    DST_RNG.Cells(l(2) + 5, 2).Address & "+1-" & _
                    DST_RNG.Cells(l(2) + 3, 2).Address & ",1))"
        
        .Cells(4, 1 + j).formula = "=STDEVP(OFFSET(" & _
                            DST_RNG.Cells(l(6) + 2, 1).Offset(0, j).Address & "," & _
                            DST_RNG.Cells(l(2) + 3, 2).Address & "-1," & _
                            0 & "," & _
                            DST_RNG.Cells(l(2) + 5, 2).Address & "+1-" & _
                            DST_RNG.Cells(l(2) + 3, 2).Address & ",1))"
        
        .Cells(5, 1 + j).FormulaArray = "=STDEVP(IF(OFFSET(" & _
                            DST_RNG.Cells(l(6) + 2, 1).Offset(0, j).Address & "," & _
                            DST_RNG.Cells(l(2) + 3, 2).Address & "-1," & _
                            0 & "," & _
                            DST_RNG.Cells(l(2) + 5, 2).Address & "+1-" & _
                            DST_RNG.Cells(l(2) + 3, 2).Address & ",1)<OFFSET(" & _
                            DST_RNG.Cells(l(5) + 2, 1).Address & ",0," & _
                            DST_RNG.Cells(l(2) + 8, 2).Address & ")/" & _
                            DST_RNG.Cells(l(2) + 16, 2).Address & ",OFFSET(" & _
                            DST_RNG.Cells(l(6) + 2, 1).Offset(0, j).Address & "," & _
                            DST_RNG.Cells(l(2) + 3, 2).Address & "-1," & _
                            0 & "," & _
                            DST_RNG.Cells(l(2) + 5, 2).Address & "+1-" & _
                            DST_RNG.Cells(l(2) + 3, 2).Address & ",1)))"
        
        .Cells(6, 1 + j).FormulaArray = "=STDEVP(IF(OFFSET(" & _
                            DST_RNG.Cells(l(6) + 2, 1).Offset(0, j).Address & "," & _
                            DST_RNG.Cells(l(2) + 3, 2).Address & "-1," & _
                            0 & "," & _
                            DST_RNG.Cells(l(2) + 5, 2).Address & "+1-" & _
                            DST_RNG.Cells(l(2) + 3, 2).Address & ",1)>OFFSET(" & _
                            DST_RNG.Cells(l(5) + 2, 1).Address & ",0," & _
                            DST_RNG.Cells(l(2) + 8, 2).Address & ")/" & _
                            DST_RNG.Cells(l(2) + 16, 2).Address & ",OFFSET(" & _
                            DST_RNG.Cells(l(6) + 2, 1).Offset(0, j).Address & "," & _
                            DST_RNG.Cells(l(2) + 3, 2).Address & "-1," & _
                            0 & "," & _
                            DST_RNG.Cells(l(2) + 5, 2).Address & "+1-" & _
                            DST_RNG.Cells(l(2) + 3, 2).Address & ",1)))"
        
        .Cells(7, 1 + j).FormulaArray = "=(PRODUCT(1+OFFSET(" & _
                            DST_RNG.Cells(l(6) + 2, 1).Offset(0, j).Address & "," & _
                            DST_RNG.Cells(l(2) + 3, 2).Address & "-1," & _
                            0 & "," & _
                            DST_RNG.Cells(l(2) + 5, 2).Address & "+1-" & _
                            DST_RNG.Cells(l(2) + 3, 2).Address & ",1))^(1/COUNT(OFFSET(" & _
                            DST_RNG.Cells(l(6) + 2, 1).Offset(0, j).Address & "," & _
                            DST_RNG.Cells(l(2) + 3, 2).Address & "-1," & _
                            0 & "," & _
                            DST_RNG.Cells(l(2) + 5, 2).Address & "+1-" & _
                            DST_RNG.Cells(l(2) + 3, 2).Address & ",1)))-1)"
        
        .Cells(8, 1 + j).formula = "=VARA(OFFSET(" & _
                            DST_RNG.Cells(l(6) + 2, 1).Offset(0, j).Address & "," & _
                            DST_RNG.Cells(l(2) + 3, 2).Address & "-1," & _
                            0 & "," & _
                            DST_RNG.Cells(l(2) + 5, 2).Address & "+1-" & _
                            DST_RNG.Cells(l(2) + 3, 2).Address & ",1))"
        
        .Cells(9, 1 + j).formula = "=PEARSON(OFFSET(" & _
                            DST_RNG.Cells(l(6) + 2, 1).Offset(0, j).Address & "," & _
                            DST_RNG.Cells(l(2) + 3, 2).Address & "-1," & _
                            0 & "," & _
                            DST_RNG.Cells(l(2) + 5, 2).Address & "+1-" & _
                            DST_RNG.Cells(l(2) + 3, 2).Address & ",1),OFFSET(" & _
                            DST_RNG.Cells(l(1) + 5, 1).Address & "," & _
                            DST_RNG.Cells(l(2) + 3, 2).Address & "-1," & _
                            DST_RNG.Cells(l(2) + 12, 2).Address & "," & _
                            DST_RNG.Cells(l(2) + 5, 2).Address & "+1-" & _
                            DST_RNG.Cells(l(2) + 3, 2).Address & ",1)/" & _
                            DST_RNG.Cells(l(2) + 16, 2).Address & ")"
        
        .Cells(10, 1 + j).formula = "=PEARSON(OFFSET(" & _
                            DST_RNG.Cells(l(6) + 2, 1).Offset(0, j).Address & "," & _
                            DST_RNG.Cells(l(2) + 3, 2).Address & "-1," & _
                            0 & "," & _
                            DST_RNG.Cells(l(2) + 5, 2).Address & "+1-" & _
                            DST_RNG.Cells(l(2) + 3, 2).Address & ",1),OFFSET(" & _
                            DST_RNG.Cells(l(1) + 5, 1).Address & "," & _
                            DST_RNG.Cells(l(2) + 3, 2).Address & "-1," & _
                            DST_RNG.Cells(l(2) + 14, 2).Address & "," & _
                            DST_RNG.Cells(l(2) + 5, 2).Address & "+1-" & _
                            DST_RNG.Cells(l(2) + 3, 2).Address & ",1)/" & _
                            DST_RNG.Cells(l(2) + 16, 2).Address & ")"
        
        .Cells(11, 1 + j).formula = "=(" & _
                            .Cells(3, 1).Offset(0, j).Address & "-(OFFSET(" & _
                            DST_RNG.Cells(l(5) + 2, 1).Address & ",0," & _
                            DST_RNG.Cells(l(2) + 8, 2).Address & ")/" & _
                            DST_RNG.Cells(l(2) + 16, 2).Address & "))/" & _
                            .Cells(4, 1).Offset(0, j).Address & ""
        
        .Cells(12, 1 + j).formula = "=(" & _
                            .Cells(3, 1).Offset(0, j).Address & "-(OFFSET(" & _
                            DST_RNG.Cells(l(5) + 2, 1).Address & ",0," & _
                            DST_RNG.Cells(l(2) + 8, 2).Address & ")/" & _
                            DST_RNG.Cells(l(2) + 16, 2).Address & "))/" & _
                            .Cells(5, 1).Offset(0, j).Address & ""
        
        .Cells(13, 1 + j).formula = "=SKEW(OFFSET(" & _
                            DST_RNG.Cells(l(6) + 2, 1).Offset(0, j).Address & "," & _
                            DST_RNG.Cells(l(2) + 3, 2).Address & "-1," & _
                            0 & "," & _
                            DST_RNG.Cells(l(2) + 5, 2).Address & "+1-" & _
                            DST_RNG.Cells(l(2) + 3, 2).Address & ",1))"
        .Cells(14, 1 + j).formula = "=KURT(OFFSET(" & _
                            DST_RNG.Cells(l(6) + 2, 1).Offset(0, j).Address & "," & _
                            DST_RNG.Cells(l(2) + 3, 2).Address & "-1," & _
                            0 & "," & _
                            DST_RNG.Cells(l(2) + 5, 2).Address & "+1-" & _
                            DST_RNG.Cells(l(2) + 3, 2).Address & ",1))"
        
        .Cells(15, 1 + j).FormulaArray = "=" & DST_RNG.Cells(l(2), 2).Address & _
                            "*PRODUCT(1+OFFSET(" & _
                            DST_RNG.Cells(l(6) + 2, 1).Offset(0, j).Address & "," & _
                            DST_RNG.Cells(l(2) + 3, 2).Address & "-1,0," & _
                            DST_RNG.Cells(l(2) + 5, 2).Address & "+1-" & _
                            DST_RNG.Cells(l(2) + 3, 2).Address & ",1))"
        
        .Cells(16, 1 + j).formula = "=" & DST_RNG.Cells(l(2), 2).Address & _
                            "*SUMPRODUCT(" & Range(DST_RNG.Cells(l(3) + 5, 2), _
                            DST_RNG.Cells(l(3) + 5, 2 + (NO_ASSETS - 2))).Offset(j - 1, 0).Address & _
                            "," & Range(DST_RNG.Cells(l(5) + 14, 2), _
                                        DST_RNG.Cells(l(5) + 14, 2 + _
                                        (NO_ASSETS - 2))).Address & ")"
        
        .Cells(17, 1 + j).FormulaArray = "=" & DST_RNG.Cells(l(2), 2).Address & _
                            "*PRODUCT(1+((OFFSET(" & _
                            DST_RNG.Cells(l(6) + 2, 1).Offset(0, j).Address & "," & _
                            DST_RNG.Cells(l(2) + 3, 2).Address & "-1,0," & _
                            DST_RNG.Cells(l(2) + 5, 2).Address & "+1-" & _
                            DST_RNG.Cells(l(2) + 3, 2).Address & ",1)-OFFSET(" & _
                            DST_RNG.Cells(l(1) + 5, 1).Address & "," & _
                            DST_RNG.Cells(l(2) + 3, 2).Address & "-1," & _
                            DST_RNG.Cells(l(2) + 10, 2).Address & "," & _
                            DST_RNG.Cells(l(2) + 5, 2).Address & "+1-" & _
                            DST_RNG.Cells(l(2) + 3, 2).Address & ",1)/" & _
                            DST_RNG.Cells(l(2) + 16, 2).Address & ")/(1+OFFSET(" & _
                            DST_RNG.Cells(l(1) + 5, 1).Address & "," & _
                            DST_RNG.Cells(l(2) + 3, 2).Address & "-1," & _
                            DST_RNG.Cells(l(2) + 10, 2).Address & "," & _
                            DST_RNG.Cells(l(2) + 5, 2).Address & "+1-" & _
                            DST_RNG.Cells(l(2) + 3, 2).Address & ",1)/" & _
                            DST_RNG.Cells(l(2) + 16, 2).Address & ")))"
        
        .Cells(18, 1 + j).formula = "=" & DST_RNG.Cells(l(2), 2).Address & _
                            "*SUMPRODUCT(" & Range(DST_RNG.Cells(l(3) + 5, 2), _
                            DST_RNG.Cells(l(3) + 5, 2 + (NO_ASSETS - 2))).Offset(j - 1, 0).Address & _
                            "," & Range(DST_RNG.Cells(l(5) + 15, 2), _
                                        DST_RNG.Cells(l(5) + 15, 2 + _
                                        (NO_ASSETS - 2))).Address & ")"
    Next j

    Set TMP_RNG = Range(.Cells(3, 2), .Cells(18, 2 + NO_PORTFOLIOS))
    TMP_RNG.NumberFormat = "#,##0.00"

End With



'-----------------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------------
RNG_PORT_BACK_TESTING_FUNC = True


Exit Function

'-----------------------------------------------------------------------------------------
FONT_LINE:
'-----------------------------------------------------------------------------------------
    With TMP_RNG
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
        .ColumnWidth = 15
        .Columns(1).ColumnWidth = 40
        .RowHeight = 15
        
        With .Font
            .name = "Courier New"
            .Size = 10
            .Strikethrough = False
            .Superscript = False
            .Subscript = False
            .OutlineFont = False
            .Shadow = False
            .Underline = xlUnderlineStyleNone
            .ColorIndex = xlAutomatic
            .TintAndShade = 0
            .ThemeFont = xlThemeFontNone
        End With
    End With

'-----------------------------------------------------------------------------------------
Return
'-----------------------------------------------------------------------------------------

'-----------------------------------------------------------------------------------------
BORDER_LINE:
'-----------------------------------------------------------------------------------------
    With TMP_RNG
        .Interior.ColorIndex = h
        .Borders(xlDiagonalDown).LineStyle = xlNone
        .Borders(xlDiagonalUp).LineStyle = xlNone
        With .Borders(xlEdgeLeft)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .WEIGHT = xlThin
        End With
        With .Borders(xlEdgeTop)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .WEIGHT = xlThin
        End With
        With .Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .WEIGHT = xlThin
        End With
        With .Borders(xlEdgeRight)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .WEIGHT = xlThin
        End With
        .Borders(xlInsideVertical).LineStyle = xlNone
        .Borders(xlInsideHorizontal).LineStyle = xlNone
    End With

'-----------------------------------------------------------------------------------------
Return
'-----------------------------------------------------------------------------------------

ERROR_LABEL:
RNG_PORT_BACK_TESTING_FUNC = False
End Function
