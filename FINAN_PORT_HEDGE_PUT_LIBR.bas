Attribute VB_Name = "FINAN_PORT_HEDGE_PUT_LIBR"
'----------------------------------------------------------------------------------------------------------------
'----------------------------------------------------------------------------------------------------------------
Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.
'----------------------------------------------------------------------------------------------------------------
'----------------------------------------------------------------------------------------------------------------

Function PORT_HEDGE_INDEX_PUT_OPTION_FUNC( _
ByRef INDEX_TICKER_RNG As Variant, _
ByRef PORT_BETA_RNG As Variant, _
ByRef CASH_RATE_RNG As Variant, _
ByRef DIVIDEND_YIELD_RNG As Variant, _
ByRef INDEX_ANNUALIZED_RETURN_RNG As Variant, _
ByRef INDEX_CURRENT_VALUE_RNG As Variant, _
ByRef INDEX_STRIKE_PRICE_RNG As Variant, _
ByRef PORT_CURRENT_VALUE_RNG As Variant, _
ByRef PERCENT_DECLINE_HEDGE_RNG As Variant, _
ByRef INDEX_CLOSING_PRICE_RNG As Variant, _
ByRef PERIODS_RNG As Variant, _
Optional ByVal NO_PERIODS_PER_YEAR_RNG As Double = 12)

Dim i As Long
Dim j As Long
Dim k As Double
Dim NROWS As Long
Dim NCOLUMNS As Long

Dim INDEX_TICKER_VECTOR As Variant
Dim PORT_BETA_VECTOR As Variant
Dim CASH_RATE_VECTOR As Variant
Dim DIVIDEND_YIELD_VECTOR As Variant
Dim INDEX_ANNUALIZED_RETURN_VECTOR As Variant
Dim INDEX_CURRENT_VALUE_VECTOR As Variant
Dim INDEX_STRIKE_PRICE_VECTOR As Variant
Dim PORT_CURRENT_VALUE_VECTOR As Variant
Dim PERCENT_DECLINE_HEDGE_VECTOR As Variant
Dim INDEX_CLOSING_PRICE_VECTOR As Variant
Dim PERIODS_VECTOR As Variant
Dim NO_PERIODS_PER_YEAR_VECTOR

Dim HEADINGS_STR As String

On Error GoTo ERROR_LABEL

If IsArray(INDEX_TICKER_RNG) = True Then
    INDEX_TICKER_VECTOR = INDEX_TICKER_RNG
    If UBound(INDEX_TICKER_VECTOR, 1) Then
        INDEX_TICKER_VECTOR = MATRIX_TRANSPOSE_FUNC(INDEX_TICKER_VECTOR)
    End If
Else
    ReDim INDEX_TICKER_VECTOR(1 To 1, 1 To 1)
    INDEX_TICKER_VECTOR(1, 1) = INDEX_TICKER_RNG
End If
NROWS = UBound(INDEX_TICKER_VECTOR, 1)

If IsArray(PORT_BETA_RNG) = True Then
    PORT_BETA_VECTOR = PORT_BETA_RNG
    If UBound(PORT_BETA_VECTOR, 1) Then
        PORT_BETA_VECTOR = MATRIX_TRANSPOSE_FUNC(PORT_BETA_VECTOR)
    End If
Else
    ReDim PORT_BETA_VECTOR(1 To 1, 1 To 1)
    PORT_BETA_VECTOR(1, 1) = PORT_BETA_RNG
End If
If NROWS <> UBound(PORT_BETA_VECTOR, 1) Then: GoTo ERROR_LABEL

If IsArray(CASH_RATE_RNG) = True Then
    CASH_RATE_VECTOR = CASH_RATE_RNG
    If UBound(CASH_RATE_VECTOR, 1) Then
        CASH_RATE_VECTOR = MATRIX_TRANSPOSE_FUNC(CASH_RATE_VECTOR)
    End If
Else
    ReDim CASH_RATE_VECTOR(1 To 1, 1 To 1)
    CASH_RATE_VECTOR(1, 1) = CASH_RATE_RNG
End If
If NROWS <> UBound(CASH_RATE_VECTOR, 1) Then: GoTo ERROR_LABEL

If IsArray(DIVIDEND_YIELD_RNG) = True Then
    DIVIDEND_YIELD_VECTOR = DIVIDEND_YIELD_RNG
    If UBound(DIVIDEND_YIELD_VECTOR, 1) Then
        DIVIDEND_YIELD_VECTOR = MATRIX_TRANSPOSE_FUNC(DIVIDEND_YIELD_VECTOR)
    End If
Else
    ReDim DIVIDEND_YIELD_VECTOR(1 To 1, 1 To 1)
    DIVIDEND_YIELD_VECTOR(1, 1) = DIVIDEND_YIELD_RNG
End If
If NROWS <> UBound(DIVIDEND_YIELD_VECTOR, 1) Then: GoTo ERROR_LABEL

If IsArray(INDEX_ANNUALIZED_RETURN_RNG) = True Then
    INDEX_ANNUALIZED_RETURN_VECTOR = INDEX_ANNUALIZED_RETURN_RNG
    If UBound(INDEX_ANNUALIZED_RETURN_VECTOR, 1) Then
        INDEX_ANNUALIZED_RETURN_VECTOR = MATRIX_TRANSPOSE_FUNC(INDEX_ANNUALIZED_RETURN_VECTOR)
    End If
Else
    ReDim INDEX_ANNUALIZED_RETURN_VECTOR(1 To 1, 1 To 1)
    INDEX_ANNUALIZED_RETURN_VECTOR(1, 1) = INDEX_ANNUALIZED_RETURN_RNG
End If
If NROWS <> UBound(INDEX_ANNUALIZED_RETURN_VECTOR, 1) Then: GoTo ERROR_LABEL

If IsArray(INDEX_CURRENT_VALUE_RNG) = True Then
    INDEX_CURRENT_VALUE_VECTOR = INDEX_CURRENT_VALUE_RNG
    If UBound(INDEX_CURRENT_VALUE_VECTOR, 1) Then
        INDEX_CURRENT_VALUE_VECTOR = MATRIX_TRANSPOSE_FUNC(INDEX_CURRENT_VALUE_VECTOR)
    End If
Else
    ReDim INDEX_CURRENT_VALUE_VECTOR(1 To 1, 1 To 1)
    INDEX_CURRENT_VALUE_VECTOR(1, 1) = INDEX_CURRENT_VALUE_RNG
End If
If NROWS <> UBound(INDEX_CURRENT_VALUE_VECTOR, 1) Then: GoTo ERROR_LABEL

If IsArray(INDEX_STRIKE_PRICE_RNG) = True Then
    INDEX_STRIKE_PRICE_VECTOR = INDEX_STRIKE_PRICE_RNG
    If UBound(INDEX_STRIKE_PRICE_VECTOR, 1) Then
        INDEX_STRIKE_PRICE_VECTOR = MATRIX_TRANSPOSE_FUNC(INDEX_STRIKE_PRICE_VECTOR)
    End If
Else
    ReDim INDEX_STRIKE_PRICE_VECTOR(1 To 1, 1 To 1)
    INDEX_STRIKE_PRICE_VECTOR(1, 1) = INDEX_STRIKE_PRICE_RNG
End If
If NROWS <> UBound(INDEX_STRIKE_PRICE_VECTOR, 1) Then: GoTo ERROR_LABEL

If IsArray(PORT_CURRENT_VALUE_RNG) = True Then
    PORT_CURRENT_VALUE_VECTOR = PORT_CURRENT_VALUE_RNG
    If UBound(PORT_CURRENT_VALUE_VECTOR, 1) Then
        PORT_CURRENT_VALUE_VECTOR = MATRIX_TRANSPOSE_FUNC(PORT_CURRENT_VALUE_VECTOR)
    End If
Else
    ReDim PORT_CURRENT_VALUE_VECTOR(1 To 1, 1 To 1)
    PORT_CURRENT_VALUE_VECTOR(1, 1) = PORT_CURRENT_VALUE_RNG
End If
If NROWS <> UBound(PORT_CURRENT_VALUE_VECTOR, 1) Then: GoTo ERROR_LABEL

If IsArray(PERCENT_DECLINE_HEDGE_RNG) = True Then
    PERCENT_DECLINE_HEDGE_VECTOR = PERCENT_DECLINE_HEDGE_RNG
    If UBound(PERCENT_DECLINE_HEDGE_VECTOR, 1) Then
        PERCENT_DECLINE_HEDGE_VECTOR = MATRIX_TRANSPOSE_FUNC(PERCENT_DECLINE_HEDGE_VECTOR)
    End If
Else
    ReDim PERCENT_DECLINE_HEDGE_VECTOR(1 To 1, 1 To 1)
    PERCENT_DECLINE_HEDGE_VECTOR(1, 1) = PERCENT_DECLINE_HEDGE_RNG
End If
If NROWS <> UBound(PERCENT_DECLINE_HEDGE_VECTOR, 1) Then: GoTo ERROR_LABEL

If IsArray(INDEX_CLOSING_PRICE_RNG) = True Then
    INDEX_CLOSING_PRICE_VECTOR = INDEX_CLOSING_PRICE_RNG
    If UBound(INDEX_CLOSING_PRICE_VECTOR, 1) Then
        INDEX_CLOSING_PRICE_VECTOR = MATRIX_TRANSPOSE_FUNC(INDEX_CLOSING_PRICE_VECTOR)
    End If
Else
    ReDim INDEX_CLOSING_PRICE_VECTOR(1 To 1, 1 To 1)
    INDEX_CLOSING_PRICE_VECTOR(1, 1) = INDEX_CLOSING_PRICE_RNG
End If
If NROWS <> UBound(INDEX_CLOSING_PRICE_VECTOR, 1) Then: GoTo ERROR_LABEL

PERIODS_VECTOR = PERIODS_RNG
If IsArray(PERIODS_RNG) = True Then
    PERIODS_VECTOR = PERIODS_RNG
    If UBound(PERIODS_VECTOR, 1) Then
        PERIODS_VECTOR = MATRIX_TRANSPOSE_FUNC(PERIODS_VECTOR)
    End If
Else
    ReDim PERIODS_VECTOR(1 To 1, 1 To 1)
    PERIODS_VECTOR(1, 1) = PERIODS_RNG
End If
If NROWS <> UBound(PERIODS_VECTOR, 1) Then: GoTo ERROR_LABEL

NO_PERIODS_PER_YEAR_VECTOR = NO_PERIODS_PER_YEAR_RNG
If IsArray(NO_PERIODS_PER_YEAR_RNG) = True Then
    NO_PERIODS_PER_YEAR_VECTOR = NO_PERIODS_PER_YEAR_RNG
    If UBound(NO_PERIODS_PER_YEAR_VECTOR, 1) Then
        NO_PERIODS_PER_YEAR_VECTOR = MATRIX_TRANSPOSE_FUNC(NO_PERIODS_PER_YEAR_VECTOR)
    End If
Else
    ReDim NO_PERIODS_PER_YEAR_VECTOR(1 To NROWS, 1 To 1)
    For i = 1 To NROWS
        NO_PERIODS_PER_YEAR_VECTOR(i, 1) = NO_PERIODS_PER_YEAR_RNG
    Next i
End If
If NROWS <> UBound(NO_PERIODS_PER_YEAR_VECTOR, 1) Then: GoTo ERROR_LABEL

HEADINGS_STR = "No,Index Name,Portfolio Beta,Risk-Free Rate (X periods Treasury),Dividend Yield (Annualized),Annual Return of Index ,Total Return of SPY in X periods,Excess Return of Index,Annual Portfolio Return,Total Return of Portfolio in X periods,Excess Return of Portfolio,Current Value of Index ,Current Portfolio Value,Percent Decline to Hedge,Portfolio Value After Decline,Current Strike Price of Index Put Option,Strike Price of Index Put Option to Buy,Number of Index Put Options to Buy,Index Closing Price,Strike Price of Index Put Option,Value of Put Option,Portfolio Value Before Hedge,Portfolio Value After Hedge,No Periods,"
NCOLUMNS = 24
ReDim TEMP_MATRIX(0 To NROWS, 1 To NCOLUMNS)
i = 1
For k = 1 To NCOLUMNS
    j = InStr(i, HEADINGS_STR, ",")
    TEMP_MATRIX(0, k) = Mid(HEADINGS_STR, i, j - i)
    i = j + 1
Next k
'--------------------------------------------------------------------------------------------------------------------
For i = 1 To NROWS
'--------------------------------------------------------------------------------------------------------------------
    k = (NO_PERIODS_PER_YEAR_VECTOR(i, 1) / PERIODS_VECTOR(i, 1))
    TEMP_MATRIX(i, 1) = i
    TEMP_MATRIX(i, 2) = INDEX_TICKER_VECTOR(i, 1)
    TEMP_MATRIX(i, 3) = PORT_BETA_VECTOR(i, 1)
    TEMP_MATRIX(i, 4) = CASH_RATE_VECTOR(i, 1)
    TEMP_MATRIX(i, 5) = DIVIDEND_YIELD_VECTOR(i, 1)
    TEMP_MATRIX(i, 6) = INDEX_ANNUALIZED_RETURN_VECTOR(i, 1)
    TEMP_MATRIX(i, 7) = (TEMP_MATRIX(i, 5) + TEMP_MATRIX(i, 6)) / k
    TEMP_MATRIX(i, 8) = TEMP_MATRIX(i, 7) - TEMP_MATRIX(i, 4)
    TEMP_MATRIX(i, 9) = TEMP_MATRIX(i, 6) * TEMP_MATRIX(i, 3)
    TEMP_MATRIX(i, 10) = (TEMP_MATRIX(i, 9) + TEMP_MATRIX(i, 5)) / k
    TEMP_MATRIX(i, 11) = TEMP_MATRIX(i, 10) - TEMP_MATRIX(i, 4)
    TEMP_MATRIX(i, 12) = INDEX_CURRENT_VALUE_VECTOR(i, 1)
    TEMP_MATRIX(i, 13) = PORT_CURRENT_VALUE_VECTOR(i, 1)
    TEMP_MATRIX(i, 14) = PERCENT_DECLINE_HEDGE_VECTOR(i, 1)
    TEMP_MATRIX(i, 15) = (1 + TEMP_MATRIX(i, 14)) * TEMP_MATRIX(i, 13)
    TEMP_MATRIX(i, 16) = INDEX_STRIKE_PRICE_VECTOR(i, 1)
    TEMP_MATRIX(i, 17) = (1 + (TEMP_MATRIX(i, 14) / TEMP_MATRIX(i, 3))) * TEMP_MATRIX(i, 16)
    TEMP_MATRIX(i, 18) = TEMP_MATRIX(i, 13) / ((100 * TEMP_MATRIX(i, 17)) / TEMP_MATRIX(i, 3))
    TEMP_MATRIX(i, 18) = ASYM_DOWN_FUNC(TEMP_MATRIX(i, 18), 1)
    TEMP_MATRIX(i, 19) = INDEX_CLOSING_PRICE_VECTOR(i, 1) 'Value of Index in X Months
    TEMP_MATRIX(i, 20) = TEMP_MATRIX(i, 17)
    TEMP_MATRIX(i, 21) = ((TEMP_MATRIX(i, 20) - TEMP_MATRIX(i, 19)) * 100) * TEMP_MATRIX(i, 18)
    TEMP_MATRIX(i, 22) = (1 + (TEMP_MATRIX(i, 14) * TEMP_MATRIX(i, 3))) * TEMP_MATRIX(i, 13) 'Value of Portfolio in X Months
    TEMP_MATRIX(i, 23) = TEMP_MATRIX(i, 21) + TEMP_MATRIX(i, 22)
    TEMP_MATRIX(i, 24) = PERIODS_VECTOR(i, 1)
'--------------------------------------------------------------------------------------------------------------------
Next i
'--------------------------------------------------------------------------------------------------------------------
PORT_HEDGE_INDEX_PUT_OPTION_FUNC = TEMP_MATRIX
'--------------------------------------------------------------------------------------------------------------------

Exit Function
ERROR_LABEL:
PORT_HEDGE_INDEX_PUT_OPTION_FUNC = Err.number
End Function

