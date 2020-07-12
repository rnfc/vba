Attribute VB_Name = "FINAN_DERIV_BS_DIVIDEND_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.


'************************************************************************************
'************************************************************************************
'FUNCTION      : OPTION_DIVIDEND_YIELD_FUNC
'DESCRIPTION   : DIVD YIELD ESTIMATION
'LIBRARY       : DERIV_BS
'GROUP         : DIVIDEND
'ID            : 001
'LAST UPDATE   : 27/02/2008
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'************************************************************************************
'************************************************************************************

Function OPTION_DIVIDEND_YIELD_FUNC(ByRef DATES_RNG As Variant, _
ByRef DIVIDENDS_RNG As Variant, _
ByVal S_VAL As Double, _
ByVal RF_VAL As Double, _
ByVal EXPIRY_DATE As Date, _
Optional ByVal TDAYS_PER_YEAR As Double = 365.25, _
Optional ByVal CURRENT_DATE As Date = 0)
'S_VAL: Underlying Price (e.g. KLSE CI)
'RF_VAL: Interest Rate
Dim i As Long
Dim NROWS As Long
Dim TEMP_SUM As Double
Dim DATES_VECTOR As Variant
Dim DIVIDENDS_VECTOR As Variant

On Error GoTo ERROR_LABEL

If CURRENT_DATE = 0 Then
    CURRENT_DATE = Now
    CURRENT_DATE = DateSerial(Year(CURRENT_DATE), Month(CURRENT_DATE), Day(CURRENT_DATE))
End If

If IsArray(DATES_RNG) = True Then 'Period Ending
    DATES_VECTOR = DATES_RNG
    If UBound(DATES_VECTOR, 1) = 1 Then
        DATES_VECTOR = MATRIX_TRANSPOSE_FUNC(DATES_VECTOR)
    End If
Else
    ReDim DATES_VECTOR(1 To 1, 1 To 1)
    DATES_VECTOR(1, 1) = DATES_RNG
End If
NROWS = UBound(DATES_VECTOR, 1)
    
If IsArray(DIVIDENDS_RNG) = True Then 'Dividend Amount
    DIVIDENDS_VECTOR = DIVIDENDS_RNG
    If UBound(DIVIDENDS_VECTOR, 1) = 1 Then
        DIVIDENDS_VECTOR = MATRIX_TRANSPOSE_FUNC(DIVIDENDS_VECTOR)
    End If
Else
    ReDim DIVIDENDS_VECTOR(1 To 1, 1 To 1)
    DIVIDENDS_VECTOR(1, 1) = DIVIDENDS_RNG
End If
If NROWS <> UBound(DIVIDENDS_VECTOR, 1) Then: GoTo ERROR_LABEL
    
TEMP_SUM = 0
For i = 1 To NROWS
    If (DATES_VECTOR(i, 1) > CURRENT_DATE And DATES_VECTOR(i, 1) < EXPIRY_DATE) Then
        TEMP_SUM = TEMP_SUM + DIVIDENDS_VECTOR(i, 1) * Exp(-RF_VAL * (DATES_VECTOR(i, 1) - CURRENT_DATE) / TDAYS_PER_YEAR)
    End If
Next i

OPTION_DIVIDEND_YIELD_FUNC = TEMP_SUM / S_VAL 'Dividend Yield

Exit Function
ERROR_LABEL:
OPTION_DIVIDEND_YIELD_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : OPTION_DIVIDENDS_ADJUSTED_SPOT_PRICE_FUNC
'DESCRIPTION   : Adjusted Spot Price for the European option Model (stock with
'cash dividends)
'LIBRARY       : DERIV_BS
'GROUP         : DIVIDEND
'ID            : 002
'LAST UPDATE   : 21/10/2013
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'************************************************************************************
'************************************************************************************

Function OPTION_DIVIDENDS_ADJUSTED_SPOT_PRICE_FUNC(ByVal SETTLEMENT As Date, _
ByVal START_DATE As Date, _
ByVal END_DATE As Date, _
ByVal SPOT As Double, _
ByVal RISK_FREE As Double, _
ByRef DATES_RNG As Variant, _
ByRef DIVIDENDS_RNG As Variant, _
Optional ByVal COUNT_BASIS As Integer = 0)

Dim i As Long
Dim j As Long

Dim SROW As Long
Dim NROWS As Long
Dim NSIZE As Long
Dim TEMP_SUM As Double
Dim CAGR_VAL As Double
Dim COST_VAL As Double

Dim TEMP_MATRIX As Variant
Dim DATES_VECTOR As Variant
Dim DIVIDENDS_VECTOR As Variant

On Error GoTo ERROR_LABEL

DATES_VECTOR = DATES_RNG
If UBound(DATES_VECTOR, 1) = 1 Then
    DATES_VECTOR = MATRIX_TRANSPOSE_FUNC(DATES_VECTOR)
End If

DIVIDENDS_VECTOR = DIVIDENDS_RNG
If UBound(DIVIDENDS_VECTOR, 1) = 1 Then
    DIVIDENDS_VECTOR = MATRIX_TRANSPOSE_FUNC(DIVIDENDS_VECTOR)
End If
If UBound(DATES_VECTOR, 1) <> UBound(DIVIDENDS_VECTOR, 1) Then: GoTo ERROR_LABEL

SROW = 0: NROWS = 0
For i = 1 To UBound(DIVIDENDS_VECTOR, 1)
    If DATES_VECTOR(i, 1) = START_DATE Then: SROW = i
    If DATES_VECTOR(i, 1) = END_DATE Then: NROWS = i
    If SROW <> 0 And NROWS <> 0 Then: Exit For
Next i

NSIZE = (NROWS - SROW + 1)
ReDim TEMP_MATRIX(1 To NSIZE, 1 To 2)

CAGR_VAL = (DIVIDENDS_VECTOR(NROWS, 1) / DIVIDENDS_VECTOR(SROW, 1)) ^ (1 / (YEARFRAC_FUNC(DATES_VECTOR(SROW, 1), DATES_VECTOR(NROWS, 1), COUNT_BASIS))) - 1
COST_VAL = DIVIDENDS_VECTOR(NROWS, 1) * (1 + CAGR_VAL) / SPOT + CAGR_VAL

j = 1
For i = SROW To NROWS
    TEMP_MATRIX(j, 1) = YEARFRAC_FUNC(DATES_VECTOR(i, 1), SETTLEMENT, COUNT_BASIS)
    TEMP_MATRIX(j, 2) = DIVIDENDS_VECTOR(i, 1)
    j = j + 1
Next i

TEMP_SUM = 0
For i = 1 To NSIZE
    TEMP_SUM = TEMP_SUM + (TEMP_MATRIX(i, 2) * Exp(-RISK_FREE * TEMP_MATRIX(i, 1)))
Next i
OPTION_DIVIDENDS_ADJUSTED_SPOT_PRICE_FUNC = SPOT - TEMP_SUM

Exit Function
ERROR_LABEL:
OPTION_DIVIDENDS_ADJUSTED_SPOT_PRICE_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : OPTION_DIVIDEND_YIELD_ADJUSTED_SPOT_PRICE_FUNC
'DESCRIPTION   : Adjust Stock Price for dividends
'LIBRARY       : DERIV_BS
'GROUP         : DIVIDEND
'ID            : 003
'LAST UPDATE   : 27/02/2008
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'************************************************************************************
'************************************************************************************

Function OPTION_DIVIDEND_YIELD_ADJUSTED_SPOT_PRICE_FUNC(ByVal S_VAL As Double, _
ByVal DY_VAL As Double, _
ByVal T_VAL As Double)
'DY_VAL --> Annualized
'T_VAL --> in Years --> # Periods / TDAYS PER YEAR
'If used, then yield on the model = 0.
On Error GoTo ERROR_LABEL
OPTION_DIVIDEND_YIELD_ADJUSTED_SPOT_PRICE_FUNC = S_VAL * Exp(-DY_VAL * T_VAL)
'Also uses it for Theta Calc =S*EXP(-Div_Yield*Theta_Time)
Exit Function
ERROR_LABEL:
OPTION_DIVIDEND_YIELD_ADJUSTED_SPOT_PRICE_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : DIVIDEND_VALUE_RATE_FUNC
'DESCRIPTION   : Dividend Dollar to Dividend Rate Converter
'LIBRARY       : DERIV_BS
'GROUP         : DIVIDEND
'ID            : 004
'LAST UPDATE   : 27/02/2008
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'************************************************************************************
'************************************************************************************

Function DIVIDEND_VALUE_RATE_FUNC(ByVal Q_VAL As Double, _
ByVal START_DATE As Date, _
ByVal END_DATE As Date, _
ByVal SPOT_VAL As Double, _
Optional ByVal COUNT_BASIS As Integer = 0) As Double
'Q_VAL -> Cash Dividend
Dim TAU_VAL As Double
On Error GoTo ERROR_LABEL
TAU_VAL = YEARFRAC_FUNC(START_DATE, END_DATE, COUNT_BASIS)
DIVIDEND_VALUE_RATE_FUNC = -1 * (1 / TAU_VAL) * Log((SPOT_VAL - Q_VAL) / SPOT_VAL)
'-365 / (END_DATE - START_DATE) * Log((SPOT_VAL - Q_VAL) / SPOT_VAL)

Exit Function
ERROR_LABEL:
DIVIDEND_VALUE_RATE_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : DIVIDEND_RATE_VALUE_FUNC
'DESCRIPTION   : Dividend Rate to Dividend Value Converter
'LIBRARY       : DERIV_BS
'GROUP         : DIVIDEND
'ID            : 005
'LAST UPDATE   : 27/02/2008
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'************************************************************************************
'************************************************************************************

Function DIVIDEND_RATE_VALUE_FUNC(ByVal Q_VAL As Double, _
ByVal START_DATE As Date, _
ByVal END_DATE As Date, _
ByVal SPOT_VAL As Double, _
Optional ByVal COUNT_BASIS As Integer = 0) As Double
'Q_VAL -> DIVIDEND_RATE
Dim TAU_VAL As Double
On Error GoTo ERROR_LABEL
TAU_VAL = YEARFRAC_FUNC(START_DATE, END_DATE, COUNT_BASIS)
'DIVIDEND_RATE_VALUE_FUNC = SPOT_VAL * (1 - DISCOUNT_FUNC(END_DATE, START_DATE, Q_VAL, 1, , COUNT_BASIS))
DIVIDEND_RATE_VALUE_FUNC = SPOT_VAL * (1 - (1 * Exp(-TAU_VAL * Q_VAL)))

Exit Function
ERROR_LABEL:
DIVIDEND_RATE_VALUE_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : PERTUBED_DIVIDEND_FUNC
'DESCRIPTION   : PERTUBED DIVD ESTIMATION
'LIBRARY       : DERIV_BS
'GROUP         : DIVIDEND
'ID            : 006
'LAST UPDATE   : 27/02/2008
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'************************************************************************************
'************************************************************************************

Function PERTUBED_DIVIDEND_FUNC(ByVal OLD_DIVIDEND As Double, _
ByVal TENOR As Double, _
ByVal OLD_SPOT As Double, _
ByVal NEW_SPOT As Double)

On Error GoTo ERROR_LABEL

If TENOR = 0 Then
    PERTUBED_DIVIDEND_FUNC = OLD_DIVIDEND
ElseIf NEW_SPOT < OLD_SPOT * (1 - Exp(-OLD_DIVIDEND * TENOR)) Then
    PERTUBED_DIVIDEND_FUNC = 1
Else
    PERTUBED_DIVIDEND_FUNC = -1 / TENOR * Log(1 - OLD_SPOT / NEW_SPOT * (1 - Exp(-OLD_DIVIDEND * TENOR)))
End If

Exit Function
ERROR_LABEL:
PERTUBED_DIVIDEND_FUNC = Err.number
End Function


