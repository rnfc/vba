Attribute VB_Name = "FINAN_ASSET_MOMENTS_DRAWD_LIBR"


Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.


'************************************************************************************
'************************************************************************************
'FUNCTION      : ASSET_DRAWDOWN_RUN_UP_FUNC
'DESCRIPTION   : Maximum drawdown, time under water, under water,
'drawdowns, non-overlapping nth drawdowns, run ups (rallies), and Avg. Drawdown

'LIBRARY       : FINAN_ASSET
'GROUP         : DRAWDOWN
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 06/04/2010
'************************************************************************************
'************************************************************************************

Function ASSET_DRAWDOWN_RUN_UP_FUNC(ByVal TICKER_STR As Variant, _
Optional ByVal START_DATE As Date, _
Optional ByVal END_DATE As Date, _
Optional ByVal NSIZE As Long = 16, _
Optional ByVal OUTPUT As Integer = 3)

'NSIZE = NO DRAWDOWNS / RUN UPS

'OUTPUT 0 = Calcs
'OUTPUT 1 = Maximum Drawdown & Maximum Run Up
'OUTPUT 2/4 = n-th Non-Overlapping Drawdown
'OUTPUT 3/5 = n-th Non-Overlapping Run Ups

Dim h As Long
Dim i As Long 'Start index of maximum drawdown phase
Dim j As Long 'End index of maximum drawdown phase (lowest point)
Dim k As Long 'End index of recovery phase
Dim l As Long

Dim hh As Long
Dim ii As Long 'Start index of maximum run up phase
Dim jj As Long 'End index of maximum run up phase (lowest point)
Dim kk As Long 'End index of down phase
Dim ll As Long

Dim NROWS As Long
Dim NCOLUMNS As Long

Dim MAX1_VAL As Double
Dim MAX2_VAL As Double

Dim MIN_VAL As Double
Dim MAX_VAL As Double

Dim MIN_DD_VAL As Double
Dim MAX_DD_VAL As Double

Dim MIN_TEMP_VAL As Double
Dim MAX_TEMP_VAL As Double

Dim MIN_LAST_VAL As Double
Dim MAX_LAST_VAL As Double

Dim MAX_RUNUP_VAL As Double 'Maximum run up (relative, positive sign)
Dim MAX_DRAWDOWN_VAL As Double 'Maximum drawdown (relative, positive sign)

Dim TEMP_SUM As Double
Dim RETURN_VAL As Double
Dim TEMP_VECTOR As Variant
Dim TEMP_MATRIX As Variant
Dim DATA_MATRIX As Variant

On Error GoTo ERROR_LABEL

If IsArray(TICKER_STR) = False Then
    DATA_MATRIX = YAHOO_HISTORICAL_DATA_SERIE_FUNC(TICKER_STR, START_DATE, END_DATE, "d", "DA", False, True, True)
Else
    DATA_MATRIX = TICKER_STR
End If

'time-series vector (format: indexed (start=1) % figures)
NROWS = UBound(DATA_MATRIX, 1)
If OUTPUT = 0 Then NCOLUMNS = 12 Else NCOLUMNS = 9
'------------------------------------------------------------
ReDim TEMP_MATRIX(0 To NROWS, 1 To NCOLUMNS)
'------------------------------------------------------------
TEMP_MATRIX(0, 1) = "DATE"
TEMP_MATRIX(0, 2) = "PRICE"
TEMP_MATRIX(0, 3) = "RETURN"
TEMP_MATRIX(0, 4) = "CUMUL RETURN"

TEMP_MATRIX(0, 5) = "UNDER WATER" 'DRAWDOWNS
'unsorted vector of all drawdown values (with negative sign)
TEMP_MATRIX(0, 6) = "OVER WATER" 'RUN UPS
'unsorted vector of all drawdown values (with negative sign)
TEMP_MATRIX(0, 7) = "IS IN DRAWDOWN"
TEMP_MATRIX(0, 8) = "IS IN RECOVERY"
TEMP_MATRIX(0, 9) = "IS IN UP RUN"
'------------------------------------------------------------------------
h = 1
For j = 3 To NCOLUMNS: TEMP_MATRIX(h, j) = "": Next j
TEMP_MATRIX(h, 1) = DATA_MATRIX(h, 1)
TEMP_MATRIX(h, 2) = DATA_MATRIX(h, 2)
'------------------------------------------------------------------------
h = 2
RETURN_VAL = DATA_MATRIX(h, 2) / DATA_MATRIX(h - 1, 2) - 1
'------------------------------------------------------------------------
MIN_VAL = 1 + MINIMUM_FUNC(RETURN_VAL, 0)
MAX_RUNUP_VAL = 0
MIN_LAST_VAL = 1
ii = 2
kk = 1
If RETURN_VAL > 0 Then ll = 1 Else ll = 2
'------------------------------------------------------------------------
MAX_VAL = 1 + MAXIMUM_FUNC(RETURN_VAL, 0)
MAX_DRAWDOWN_VAL = 0
MAX_LAST_VAL = 1
i = 2
k = 1
If RETURN_VAL < 0 Then l = 1 Else l = 2
'-----------------------------------------------------------------------------
For h = 2 To NROWS
'------------------------------------------------------------------------------
    TEMP_MATRIX(h, 1) = DATA_MATRIX(h, 1)
    TEMP_MATRIX(h, 2) = DATA_MATRIX(h, 2)
    If TEMP_MATRIX(h, 1) <> 0 Then
        TEMP_MATRIX(h, 3) = DATA_MATRIX(h, 2) / DATA_MATRIX(h - 1, 2) - 1
    Else 'For OUPUT 2 or 3
        TEMP_MATRIX(h, 3) = 0
    End If
    If h <> 2 Then
        TEMP_MATRIX(h, 4) = (1 + TEMP_MATRIX(h - 1, 4)) * (1 + TEMP_MATRIX(h, 3)) - 1
    Else
        TEMP_MATRIX(h, 4) = TEMP_MATRIX(h, 3)
    End If
'-----------------------------------------------------------------------------
    MIN_TEMP_VAL = MIN_LAST_VAL * (1 + TEMP_MATRIX(h, 3))
    If MIN_TEMP_VAL < MIN_VAL Then
        MIN_VAL = MIN_TEMP_VAL
        ll = h
    End If
    
    TEMP_MATRIX(h, 6) = (MIN_TEMP_VAL / MIN_VAL) - 1 'Overwater
    MIN_DD_VAL = 1 - (MIN_TEMP_VAL / MIN_VAL)
    If MIN_DD_VAL < MAX_RUNUP_VAL Then
        MAX_RUNUP_VAL = MIN_DD_VAL
        ii = ll
        jj = h
    End If
    If ii = ll Then: kk = h + 1
    MIN_LAST_VAL = MIN_TEMP_VAL
'-----------------------------------------------------------------------------
    MAX_TEMP_VAL = MAX_LAST_VAL * (1 + TEMP_MATRIX(h, 3))
    If MAX_TEMP_VAL > MAX_VAL Then
        MAX_VAL = MAX_TEMP_VAL
        l = h
    End If
    
    TEMP_MATRIX(h, 5) = (MAX_TEMP_VAL / MAX_VAL) - 1 'Underwater
    MAX_DD_VAL = 1 - (MAX_TEMP_VAL / MAX_VAL)
    If MAX_DD_VAL > MAX_DRAWDOWN_VAL Then
        MAX_DRAWDOWN_VAL = MAX_DD_VAL
        i = l
        j = h
    End If
    If i = l Then: k = h + 1
    MAX_LAST_VAL = MAX_TEMP_VAL
Next h
'-----------------------------------------------------------------------------
For h = 2 To NROWS
    TEMP_MATRIX(h, 7) = (h >= i And h <= j)
    TEMP_MATRIX(h, 8) = (h >= j And h <= k)
    TEMP_MATRIX(h, 9) = (h >= ii And h <= jj)
Next h
'-----------------------------------------------------------------------------
Select Case OUTPUT
'-----------------------------------------------------------------------------
Case 0 'Perfect
'-----------------------------------------------------------------------------
    h = 1
    TEMP_MATRIX(h, 10) = 1
    MAX1_VAL = TEMP_MATRIX(h, 10)
    hh = h
    TEMP_MATRIX(h, 11) = MAX1_VAL
    TEMP_MATRIX(h, 12) = 0
    
    MAX2_VAL = TEMP_MATRIX(h, 12)
    TEMP_SUM = TEMP_MATRIX(h, 12)
    
    For h = 2 To NROWS
        TEMP_MATRIX(h, 10) = TEMP_MATRIX(h - 1, 10) * (1 + (DATA_MATRIX(h, 2) / DATA_MATRIX(h - 1, 2) - 1))
        If TEMP_MATRIX(h, 10) > MAX1_VAL Then: MAX1_VAL = TEMP_MATRIX(h, 10)
        TEMP_MATRIX(h, 11) = MAX1_VAL
        TEMP_MATRIX(h, 12) = 1 - TEMP_MATRIX(h, 10) / TEMP_MATRIX(h, 11)
        
        If TEMP_MATRIX(h, 12) > MAX2_VAL Then
            MAX2_VAL = TEMP_MATRIX(h, 12)
            hh = h
        End If
        
        TEMP_SUM = TEMP_SUM + TEMP_MATRIX(h, 12)
    Next h
    TEMP_MATRIX(0, 10) = "GROWTH OF $1.00"
    TEMP_MATRIX(0, 11) = "PREVIOUS MAX / MAX DRAWDOWN = " & Format(MAX2_VAL, "0.0%") & _
                         " on " & Format(DATA_MATRIX(hh, 1), "mmm d/yy")
    TEMP_MATRIX(0, 12) = "AVG DRAWDOWN = " & Format(TEMP_SUM / NROWS, "0.0%")
    
    ASSET_DRAWDOWN_RUN_UP_FUNC = TEMP_MATRIX 'Exit
'-----------------------------------------------------------------------------
Case 1, 2, 3 'Perfect
'-----------------------------------------------------------------------------
    ReDim TEMP_VECTOR(0 To 4, 0 To 2)
    TEMP_VECTOR(0, 0) = "UNDER/OVER WATER"
    TEMP_VECTOR(0, 1) = "MAX DRAWDOWN"
    TEMP_VECTOR(0, 2) = "MAX RUN UP"
    
    TEMP_VECTOR(1, 0) = "% CHANGE"
    TEMP_VECTOR(2, 0) = "START DATE" 'Beginning of period
    TEMP_VECTOR(3, 0) = "END DATE" 'End of period
    TEMP_VECTOR(4, 0) = "RECOVERY/DETERIORATION DATE" 'Growth --> X<=0 Between Beg and Ending Periods

    If MAX_RUNUP_VAL <> 0 Then
        TEMP_VECTOR(1, 2) = -MAX_RUNUP_VAL
        TEMP_VECTOR(2, 2) = DATA_MATRIX(ii, 1) '(ii + 1, 1)
        TEMP_VECTOR(3, 2) = DATA_MATRIX(jj, 1)
        If kk > NROWS Then
            TEMP_VECTOR(4, 2) = CVErr(xlErrNA)
        Else
            TEMP_VECTOR(4, 2) = DATA_MATRIX(kk, 1)
        End If
    End If
    
    If MAX_DRAWDOWN_VAL <> 0 Then
        TEMP_VECTOR(1, 1) = -MAX_DRAWDOWN_VAL
        TEMP_VECTOR(2, 1) = DATA_MATRIX(i, 1) '(i + 1, 1)
        TEMP_VECTOR(3, 1) = DATA_MATRIX(j, 1)
        If k > NROWS Then
            TEMP_VECTOR(4, 1) = CVErr(xlErrNA)
        Else
            TEMP_VECTOR(4, 1) = DATA_MATRIX(k, 1)
        End If
    End If

    If OUTPUT = 1 Then 'Perfect
        ASSET_DRAWDOWN_RUN_UP_FUNC = TEMP_VECTOR 'Exit
    Else
        If OUTPUT = 2 Or OUTPUT = 3 Then
            If OUTPUT = 2 Then
                l = 1 'n-th Non-Overlapping Draw Downs
            Else
                l = 2 'n-th Non-Overlapping Run Ups
            End If
            ReDim TEMP_MATRIX(0 To NSIZE, 1 To 5)
            TEMP_MATRIX(0, 1) = IIf(OUTPUT = 2, "UNDER WATER", "OVER WATER")
            TEMP_MATRIX(0, 2) = "% CHANGE"
            TEMP_MATRIX(0, 3) = "START DATE"
            TEMP_MATRIX(0, 4) = "END DATE"
            TEMP_MATRIX(0, 5) = IIf(OUTPUT = 2, "RECOVERY DATE", "DETERIORATION DATE")
            For h = 1 To NSIZE
                If IsArray(TEMP_VECTOR) = False Then: Exit For
                TEMP_MATRIX(h, 1) = IIf(l = 1, "DRAW DOWNS: ", "RUN UPS: ") & CStr(h)
                TEMP_MATRIX(h, 2) = TEMP_VECTOR(1, l)
                TEMP_MATRIX(h, 3) = TEMP_VECTOR(2, l)
                TEMP_MATRIX(h, 4) = TEMP_VECTOR(3, l)
                TEMP_MATRIX(h, 5) = TEMP_VECTOR(4, l)
                i = 0: j = 0
                For k = LBound(DATA_MATRIX, 1) To UBound(DATA_MATRIX, 1)
                    If DATA_MATRIX(k, 1) = TEMP_MATRIX(h, 3) Then: i = k
                    If DATA_MATRIX(k, 1) = TEMP_MATRIX(h, 4) Then: j = k
                    If i <> 0 And j <> 0 Then: Exit For
                Next k
                If i = 0 Or j = 0 Then: Exit For
                For k = i To j: DATA_MATRIX(k, 1) = 0: Next k
                TEMP_VECTOR = ASSET_DRAWDOWN_RUN_UP_FUNC(DATA_MATRIX, , , NSIZE, 1)
            Next h
            ASSET_DRAWDOWN_RUN_UP_FUNC = TEMP_MATRIX 'Exit
        End If
    End If
'-----------------------------------------------------------------------------
Case Is >= 4 'Perfect
'-----------------------------------------------------------------------------
    If OUTPUT = 4 Then 'n-th Drawdown
        DATA_MATRIX = MATRIX_QUICK_SORT_FUNC(TEMP_MATRIX, 5, 1)
        ReDim TEMP_MATRIX(0 To NROWS - 1, 1 To 2)
        TEMP_MATRIX(0, 1) = "PERIOD"
        TEMP_MATRIX(0, 2) = "NTH DRAWDOWN"
        For i = 0 To NROWS - 2
            TEMP_MATRIX(i + 1, 1) = DATA_MATRIX(i, 1)
            TEMP_MATRIX(i + 1, 2) = DATA_MATRIX(i, 5)
        Next i
    Else 'n-th Run Up
        DATA_MATRIX = MATRIX_QUICK_SORT_FUNC(TEMP_MATRIX, 6, 0)
        ReDim TEMP_MATRIX(0 To NROWS - 1, 1 To 2)
        TEMP_MATRIX(0, 1) = "PERIOD"
        TEMP_MATRIX(0, 2) = "NTH RUNUP"
        For i = 2 To NROWS
            TEMP_MATRIX(i - 1, 1) = DATA_MATRIX(i, 1)
            TEMP_MATRIX(i - 1, 2) = DATA_MATRIX(i, 6)
        Next i
    End If
    ASSET_DRAWDOWN_RUN_UP_FUNC = TEMP_MATRIX 'Exit
'-----------------------------------------------------------------------------
End Select
'-----------------------------------------------------------------------------

Exit Function
ERROR_LABEL:
ASSET_DRAWDOWN_RUN_UP_FUNC = Err.number
End Function


Function EXPECTED_MAXIMUM_DRAWDOWN_FUNC(ByVal MU_VAL As Double, _
ByVal SIGMA_VAL As Double, _
Optional ByVal PERIODS_RNG As Variant = 10, _
Optional ByVal nLOOPS As Long = 1000, _
Optional ByVal CONFIDENCE_VAL As Double = 0.95)

'This function calculates the expected maximum drawdown ans its confidence bands based
'on a given expected return (mean) and volatility (standard deviation). It is assumed that returns
'are IID and follow a normal distribution. Note: The results are derived with a resampling procedure.
'Excel's internal functionality is used, resulting in extremly slow calculations


'Calulations expected maximum drawdown for a normal distribution

'We calculate expected drawdown and its confidence bands with the help
'of a resampling procedure assuming a normal distribution.

Dim i As Long
Dim j As Long
Dim k As Long

Dim NROWS As Long
Dim NCOLUMNS As Long

Dim P_VAL As Double 'initialize chain-linked returns
Dim R_VAL As Double 'Simulate a Normal distributed random variable

Dim DD_VAL As Double
Dim MDD_VAL As Double
Dim HWM_VAL As Double 'initialize high-water mark
Dim AMDD_VAL As Double

Dim TEMP_MATRIX As Variant
Dim PERIODS_VECTOR As Variant

On Error GoTo ERROR_LABEL

If IsArray(PERIODS_RNG) Then
    PERIODS_VECTOR = PERIODS_RNG
    If UBound(PERIODS_VECTOR, 2) = 1 Then
        PERIODS_VECTOR = MATRIX_TRANSPOSE_FUNC(PERIODS_VECTOR)
    End If
Else
    ReDim PERIODS_VECTOR(1 To 1, 1 To 1)
    PERIODS_VECTOR(1, 1) = PERIODS_RNG
End If
NCOLUMNS = UBound(PERIODS_VECTOR, 2)
ReDim TEMP_MATRIX(1 To nLOOPS, 1 To NCOLUMNS)
For j = 1 To NCOLUMNS
    NROWS = PERIODS_VECTOR(1, j)
    For i = 1 To nLOOPS
        Do
            GoSub DRAWDOWN_LINE
        Loop Until AMDD_VAL < 0
        TEMP_MATRIX(i, j) = AMDD_VAL
    Next i
Next j

If NCOLUMNS = 1 Then 'Expected Maximum Drawdown / Lower Confidence / Upper Confidence
    EXPECTED_MAXIMUM_DRAWDOWN_FUNC = Array(MATRIX_MEAN_FUNC(TEMP_MATRIX)(1, 1), _
                                        HISTOGRAM_PERCENTILE_FUNC(TEMP_MATRIX, 1 - CONFIDENCE_VAL, 1), _
                                        HISTOGRAM_PERCENTILE_FUNC(TEMP_MATRIX, CONFIDENCE_VAL, 1))
Else
    EXPECTED_MAXIMUM_DRAWDOWN_FUNC = TEMP_MATRIX
End If

'------------------------------------------------------------------------------------------------------------------------------------------
Exit Function
'------------------------------------------------------------------------------------------------------------------------------------------
DRAWDOWN_LINE: 'Maximum Drawdown
'------------------------------------------------------------------------------------------------------------------------------------------
    P_VAL = 1: HWM_VAL = 1: MDD_VAL = 0
    For k = 1 To NROWS
        R_VAL = RANDOM_NORMAL_FUNC(MU_VAL, SIGMA_VAL, 0)
        P_VAL = P_VAL * (1 + R_VAL)
        If P_VAL > HWM_VAL Then: HWM_VAL = P_VAL
        DD_VAL = P_VAL / HWM_VAL - 1
        If DD_VAL < MDD_VAL Then MDD_VAL = DD_VAL
    Next k
    AMDD_VAL = DD_VAL
'------------------------------------------------------------------------------------------------------------------------------------------
Return
'------------------------------------------------------------------------------------------------------------------------------------------
ERROR_LABEL:
EXPECTED_MAXIMUM_DRAWDOWN_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : ASSET_DD_LOG_FUNC
'DESCRIPTION   : Calculates drawdown vector from a vector of log returns
'LIBRARY       : FINAN_ASSET
'GROUP         : DRAWDOWN
'ID            : 002
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 06/04/2010
'************************************************************************************
'************************************************************************************
'PERFECT
Function ASSET_DD_LOG_FUNC(ByRef DATA_RNG As Variant, _
Optional ByVal DATA_TYPE As Integer = 0)

Dim i As Long
Dim NROWS As Long

Dim MAX_VAL As Double
Dim TEMP_VECTOR As Variant

Dim DATA_VECTOR As Variant

On Error GoTo ERROR_LABEL

DATA_VECTOR = DATA_RNG
If DATA_TYPE <> 0 Then: DATA_VECTOR = MATRIX_PERCENT_FUNC(DATA_VECTOR, 1)
NROWS = UBound(DATA_VECTOR, 1)

ReDim TEMP_VECTOR(1 To NROWS, 1 To 2)

i = 1
TEMP_VECTOR(i, 2) = DATA_VECTOR(i, 1)
MAX_VAL = TEMP_VECTOR(i, 2)
TEMP_VECTOR(i, 1) = MAX_VAL - TEMP_VECTOR(i, 2)
        
For i = 2 To NROWS
    TEMP_VECTOR(i, 2) = TEMP_VECTOR(i - 1, 2) + DATA_VECTOR(i, 1)
    MAX_VAL = MAXIMUM_FUNC(TEMP_VECTOR(i, 2), MAX_VAL)
    TEMP_VECTOR(i, 1) = MAX_VAL - TEMP_VECTOR(i, 2)
Next i

TEMP_VECTOR = MATRIX_QUICK_SORT_FUNC(TEMP_VECTOR, 1, 1)
For i = 1 To NROWS
    TEMP_VECTOR(i, 1) = -1 * TEMP_VECTOR(i, 1)
Next i
        
ASSET_DD_LOG_FUNC = TEMP_VECTOR

Exit Function
ERROR_LABEL:
ASSET_DD_LOG_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : ASSET_UW_LOG_FUNC
'DESCRIPTION   : Calculates Under Water vector from a vector of log returns
'LIBRARY       : FINAN_ASSET
'GROUP         : DRAWDOWN
'ID            : 003
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 06/04/2010
'************************************************************************************
'************************************************************************************

Function ASSET_UW_LOG_FUNC(ByRef DATA_RNG As Variant, _
Optional ByVal DATA_TYPE As Integer = 0)

Dim i As Long
Dim NROWS As Long

Dim MAX_VAL As Double
Dim TEMP_VECTOR As Variant

Dim DATA_VECTOR As Variant

On Error GoTo ERROR_LABEL

DATA_VECTOR = DATA_RNG
If DATA_TYPE <> 0 Then: DATA_VECTOR = MATRIX_PERCENT_FUNC(DATA_VECTOR, 1)
NROWS = UBound(DATA_VECTOR, 1)

ReDim TEMP_VECTOR(1 To NROWS, 1 To 2)

i = 1
TEMP_VECTOR(i, 2) = DATA_VECTOR(i, 1)
MAX_VAL = TEMP_VECTOR(i, 2)
TEMP_VECTOR(i, 1) = MAX_VAL - TEMP_VECTOR(i, 2)

For i = 2 To NROWS
    TEMP_VECTOR(i, 2) = TEMP_VECTOR(i - 1, 2) + DATA_VECTOR(i, 1)
    MAX_VAL = MAXIMUM_FUNC(TEMP_VECTOR(i, 2), MAX_VAL)
    TEMP_VECTOR(i, 1) = -(MAX_VAL - TEMP_VECTOR(i, 2))
Next i
    
ASSET_UW_LOG_FUNC = TEMP_VECTOR
    
Exit Function
ERROR_LABEL:
ASSET_UW_LOG_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : ASSET_DD_MAX_LOG_FUNC
'DESCRIPTION   : Calculates Maximum Drawdown from a vector of log returns
'LIBRARY       : FINAN_ASSET
'GROUP         : DRAWDOWN
'ID            : 004
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 06/04/2010
'************************************************************************************
'************************************************************************************

Function ASSET_DD_MAX_LOG_FUNC(ByRef DATA_RNG As Variant, _
Optional ByVal DATA_TYPE As Integer = 0)

Dim i As Long
Dim j As Long
Dim NROWS As Long

Dim MAX1_VAL As Double
Dim MAX2_VAL As Double
Dim TEMP_VECTOR As Variant

Dim DATA_VECTOR As Variant

On Error GoTo ERROR_LABEL

DATA_VECTOR = DATA_RNG
If DATA_TYPE <> 0 Then: DATA_VECTOR = MATRIX_PERCENT_FUNC(DATA_VECTOR, 1)
NROWS = UBound(DATA_VECTOR, 1)

ReDim TEMP_VECTOR(1 To NROWS, 1 To 2)


i = 1
TEMP_VECTOR(i, 2) = DATA_VECTOR(i, 1)
MAX1_VAL = TEMP_VECTOR(i, 2)
TEMP_VECTOR(i, 1) = MAX1_VAL - TEMP_VECTOR(i, 2)
GoSub MAX_LINE
For i = 2 To NROWS
    TEMP_VECTOR(i, 2) = TEMP_VECTOR(i - 1, 2) + DATA_VECTOR(i, 1)
    MAX1_VAL = MAXIMUM_FUNC(TEMP_VECTOR(i, 2), MAX1_VAL)
    TEMP_VECTOR(i, 1) = MAX1_VAL - TEMP_VECTOR(i, 2)
    GoSub MAX_LINE
Next i
    
ASSET_DD_MAX_LOG_FUNC = MAX2_VAL
    
'-----------------------------------------------------------------------------------------------------
Exit Function
'-----------------------------------------------------------------------------------------------------
MAX_LINE:
'-----------------------------------------------------------------------------------------------------
    If i = 1 Then: MAX2_VAL = -2 ^ 52
    If TEMP_VECTOR(i, 1) > MAX2_VAL Then
        MAX2_VAL = TEMP_VECTOR(i, 1)
        j = i
    End If
'-----------------------------------------------------------------------------------------------------
Return
'-----------------------------------------------------------------------------------------------------
ERROR_LABEL:
ASSET_DD_MAX_LOG_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : ASSET_DD_MAX_LOG_START_PERIOD_FUNC
'DESCRIPTION   : Calculates starting point of max drawdown period of a vector of
'log returns
'LIBRARY       : FINAN_ASSET
'GROUP         : DRAWDOWN
'ID            : 005
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 06/04/2010
'************************************************************************************
'************************************************************************************
'//PERFECT
Function ASSET_DD_MAX_LOG_START_PERIOD_FUNC(ByRef DATA_RNG As Variant, _
Optional ByVal DATA_TYPE As Integer = 0, _
Optional ByVal nDIGITS As Integer = 8)

Dim i As Long
Dim j As Long
Dim NROWS As Long

Dim MAX1_VAL As Double
Dim MAX2_VAL As Double
Dim TEMP_VECTOR As Variant

Dim DATA_VECTOR As Variant

On Error GoTo ERROR_LABEL

DATA_VECTOR = DATA_RNG
If DATA_TYPE <> 0 Then: DATA_VECTOR = MATRIX_PERCENT_FUNC(DATA_VECTOR, 1)
NROWS = UBound(DATA_VECTOR, 1)

ReDim TEMP_VECTOR(1 To NROWS, 1 To 2)


i = 1
TEMP_VECTOR(i, 2) = DATA_VECTOR(i, 1)
MAX1_VAL = TEMP_VECTOR(i, 2)
TEMP_VECTOR(i, 1) = MAX1_VAL - TEMP_VECTOR(i, 2)
GoSub MAX_LINE
For i = 2 To NROWS
    TEMP_VECTOR(i, 2) = TEMP_VECTOR(i - 1, 2) + DATA_VECTOR(i, 1)
    MAX1_VAL = MAXIMUM_FUNC(TEMP_VECTOR(i, 2), MAX1_VAL)
    TEMP_VECTOR(i, 1) = MAX1_VAL - TEMP_VECTOR(i, 2)
    GoSub MAX_LINE
Next i

i = j
Do
i = i - 1
Loop Until Round(TEMP_VECTOR(j, 2) + MAX2_VAL, nDIGITS) = Round(TEMP_VECTOR(i, 2), nDIGITS)

ASSET_DD_MAX_LOG_START_PERIOD_FUNC = i
    
'-----------------------------------------------------------------------------------------------------
Exit Function
'-----------------------------------------------------------------------------------------------------
MAX_LINE:
'-----------------------------------------------------------------------------------------------------
    If i = 1 Then: MAX2_VAL = -2 ^ 52
    If TEMP_VECTOR(i, 1) > MAX2_VAL Then
        MAX2_VAL = TEMP_VECTOR(i, 1)
        j = i
    End If
'-----------------------------------------------------------------------------------------------------
Return
'-----------------------------------------------------------------------------------------------------
ERROR_LABEL:
ASSET_DD_MAX_LOG_START_PERIOD_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : ASSET_DD_MAX_LOG_END_PERIOD_FUNC
'DESCRIPTION   : Calculates ending point of max drawdown period of a vector of
'log returns
'LIBRARY       : FINAN_ASSET
'GROUP         : DRAWDOWN
'ID            : 006
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 06/04/2010
'************************************************************************************
'************************************************************************************
'//PERFECT
Function ASSET_DD_MAX_LOG_END_PERIOD_FUNC(ByRef DATA_RNG As Variant, _
Optional ByVal DATA_TYPE As Integer = 0)

Dim i As Long
Dim j As Long
Dim NROWS As Long

Dim MAX1_VAL As Double
Dim MAX2_VAL As Double
Dim TEMP_VECTOR As Variant

Dim DATA_VECTOR As Variant

On Error GoTo ERROR_LABEL

DATA_VECTOR = DATA_RNG
If DATA_TYPE <> 0 Then: DATA_VECTOR = MATRIX_PERCENT_FUNC(DATA_VECTOR, 1)
NROWS = UBound(DATA_VECTOR, 1)

ReDim TEMP_VECTOR(1 To NROWS, 1 To 2)


i = 1
TEMP_VECTOR(i, 2) = DATA_VECTOR(i, 1)
MAX1_VAL = TEMP_VECTOR(i, 2)
TEMP_VECTOR(i, 1) = MAX1_VAL - TEMP_VECTOR(i, 2)
GoSub MAX_LINE
For i = 2 To NROWS
    TEMP_VECTOR(i, 2) = TEMP_VECTOR(i - 1, 2) + DATA_VECTOR(i, 1)
    MAX1_VAL = MAXIMUM_FUNC(TEMP_VECTOR(i, 2), MAX1_VAL)
    TEMP_VECTOR(i, 1) = MAX1_VAL - TEMP_VECTOR(i, 2)
    GoSub MAX_LINE
Next i
ASSET_DD_MAX_LOG_END_PERIOD_FUNC = j
    
'-----------------------------------------------------------------------------------------------------
Exit Function
'-----------------------------------------------------------------------------------------------------
MAX_LINE:
'-----------------------------------------------------------------------------------------------------
    If i = 1 Then: MAX2_VAL = -2 ^ 52
    If TEMP_VECTOR(i, 1) > MAX2_VAL Then
        MAX2_VAL = TEMP_VECTOR(i, 1)
        j = i
    End If
'-----------------------------------------------------------------------------------------------------
Return
'-----------------------------------------------------------------------------------------------------
ERROR_LABEL:
ASSET_DD_MAX_LOG_END_PERIOD_FUNC = Err.number
End Function
  

'************************************************************************************
'************************************************************************************
'FUNCTION      : ASSET_DD_MAX_LOG_RECOVERY_PERIOD_FUNC
'DESCRIPTION   : Calculates point of recovery after max drawdown period of a
'vector of log returns
'LIBRARY       : FINAN_ASSET
'GROUP         : DRAWDOWN
'ID            : 007
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 06/04/2010
'************************************************************************************
'************************************************************************************
'// PERFECT
Function ASSET_DD_MAX_LOG_RECOVERY_PERIOD_FUNC(ByRef DATA_RNG As Variant, _
Optional ByVal DATA_TYPE As Integer = 0, _
Optional ByVal nDIGITS As Integer = 8)

Dim i As Long
Dim j As Long
Dim NROWS As Long

Dim MAX1_VAL As Double
Dim MAX2_VAL As Double
Dim TEMP_VECTOR As Variant

Dim DATA_VECTOR As Variant

On Error GoTo ERROR_LABEL

DATA_VECTOR = DATA_RNG
If DATA_TYPE <> 0 Then: DATA_VECTOR = MATRIX_PERCENT_FUNC(DATA_VECTOR, 1)
NROWS = UBound(DATA_VECTOR, 1)

ReDim TEMP_VECTOR(1 To NROWS, 1 To 2)


i = 1
TEMP_VECTOR(i, 2) = DATA_VECTOR(i, 1)
MAX1_VAL = TEMP_VECTOR(i, 2)
TEMP_VECTOR(i, 1) = MAX1_VAL - TEMP_VECTOR(i, 2)
GoSub MAX_LINE
For i = 2 To NROWS
    TEMP_VECTOR(i, 2) = TEMP_VECTOR(i - 1, 2) + DATA_VECTOR(i, 1)
    MAX1_VAL = MAXIMUM_FUNC(TEMP_VECTOR(i, 2), MAX1_VAL)
    TEMP_VECTOR(i, 1) = MAX1_VAL - TEMP_VECTOR(i, 2)
    GoSub MAX_LINE
Next i

i = j - 1
Do
i = i + 1
Loop Until Round(TEMP_VECTOR(j, 2) + MAX2_VAL, nDIGITS) <= Round(TEMP_VECTOR(i, 2), nDIGITS) Or i = NROWS

If Round(TEMP_VECTOR(j, 2) + MAX2_VAL, nDIGITS) > Round(TEMP_VECTOR(i, 2), nDIGITS) Then
    ASSET_DD_MAX_LOG_RECOVERY_PERIOD_FUNC = "--"
Else
    ASSET_DD_MAX_LOG_RECOVERY_PERIOD_FUNC = i
End If

'-----------------------------------------------------------------------------------------------------
Exit Function
'-----------------------------------------------------------------------------------------------------
MAX_LINE:
'-----------------------------------------------------------------------------------------------------
    If i = 1 Then: MAX2_VAL = -2 ^ 52
    If TEMP_VECTOR(i, 1) > MAX2_VAL Then
        MAX2_VAL = TEMP_VECTOR(i, 1)
        j = i
    End If
'-----------------------------------------------------------------------------------------------------
Return
'-----------------------------------------------------------------------------------------------------
ERROR_LABEL:
ASSET_DD_MAX_LOG_RECOVERY_PERIOD_FUNC = Err.number
End Function
