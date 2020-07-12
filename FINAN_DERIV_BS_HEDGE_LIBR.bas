Attribute VB_Name = "FINAN_DERIV_BS_HEDGE_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.

'Volatility and Drift per period
'Output in Descending Order

Function ASSET_OPTION_DELTA_HEDGING_FUNC(ByVal STARTING_PRICE As Double, _
ByVal STRIKE_PRICE As Double, _
ByVal VOLATILITY As Double, _
ByVal DRIFT As Double, _
ByVal EXPIRATION As Double, _
Optional ByVal QUANTITY As Double = 100, _
Optional ByRef OPTION_TYPE As Integer = 1, _
Optional ByVal SHORT_FLAG As Boolean = False, _
Optional ByVal COUNT_BASIS As Double = 252, _
Optional ByVal RND_TYPE As Integer = 0, _
Optional ByVal CND_TYPE As Integer = 0, _
Optional ByVal OUTPUT As Integer = 0)

'QUANTITY --> Shares
'EXPIRATION --> In Years
'COUNT_BASIS --> Steps Per Year
'SHORT_FLAG = TRUE Want to hedge a Short Position otherwise Long Position

Dim i As Long
Dim j As Long

Dim PI_VAL As Double

Dim TEMP_SUM As Double
Dim DELTA_VAL As Double

Dim TEMP_MATRIX As Variant

On Error GoTo ERROR_LABEL

PI_VAL = 3.14159265358979
DELTA_VAL = 1 / COUNT_BASIS
j = EXPIRATION * COUNT_BASIS
If OPTION_TYPE <> 1 Then: OPTION_TYPE = -1 'Put Option

ReDim TEMP_MATRIX(0 To j, 1 To 8)

TEMP_MATRIX(0, 1) = "PERIODS"
TEMP_MATRIX(0, 2) = "D1_VAL"
TEMP_MATRIX(0, 3) = "D2_VAL"
TEMP_MATRIX(0, 4) = "STOCK_PRICE"
TEMP_MATRIX(0, 5) = "OPTION_PRICE"
TEMP_MATRIX(0, 6) = "DELTA"
TEMP_MATRIX(0, 7) = "SHARES"
TEMP_MATRIX(0, 8) = "DAILY_PL"

TEMP_MATRIX(1, 1) = EXPIRATION
TEMP_MATRIX(1, 4) = STARTING_PRICE

TEMP_MATRIX(1, 2) = (Log(TEMP_MATRIX(1, 4) / STRIKE_PRICE) + (DRIFT + VOLATILITY ^ 2 / 2) * TEMP_MATRIX(1, 1)) / (VOLATILITY * Sqr(TEMP_MATRIX(1, 1)))
TEMP_MATRIX(1, 3) = (Log(TEMP_MATRIX(1, 4) / STRIKE_PRICE) + (DRIFT - VOLATILITY ^ 2 / 2) * TEMP_MATRIX(1, 1)) / (VOLATILITY * Sqr(TEMP_MATRIX(1, 1)))

'-----------------------------------------------------------------------------------
If OPTION_TYPE = 1 Then 'Call Option
'-----------------------------------------------------------------------------------
   TEMP_MATRIX(1, 5) = TEMP_MATRIX(1, 4) * CND_FUNC(TEMP_MATRIX(1, 2), CND_TYPE) - STRIKE_PRICE * Exp(-DRIFT * TEMP_MATRIX(1, 1)) * CND_FUNC(TEMP_MATRIX(1, 3), CND_TYPE)
   TEMP_MATRIX(1, 6) = CND_FUNC(TEMP_MATRIX(1, 2), CND_TYPE)
   If SHORT_FLAG = True Then
       TEMP_MATRIX(1, 7) = Round(QUANTITY * TEMP_MATRIX(1, 6), 0)
   Else
       TEMP_MATRIX(1, 7) = Round(QUANTITY * TEMP_MATRIX(1, 6), 0) * -1
   End If
'-----------------------------------------------------------------------------------
Else 'Put Option
'-----------------------------------------------------------------------------------
   TEMP_MATRIX(1, 5) = STRIKE_PRICE * Exp(-DRIFT * TEMP_MATRIX(1, 1)) * CND_FUNC(-TEMP_MATRIX(1, 3), CND_TYPE) - TEMP_MATRIX(1, 4) * CND_FUNC(-TEMP_MATRIX(1, 2), CND_TYPE)
   TEMP_MATRIX(1, 6) = -1 * CND_FUNC(-1 * TEMP_MATRIX(1, 2), CND_TYPE)
   If SHORT_FLAG = True Then
       TEMP_MATRIX(1, 7) = Round(QUANTITY * TEMP_MATRIX(1, 6), 0)
   Else
       TEMP_MATRIX(1, 7) = Round(QUANTITY * TEMP_MATRIX(1, 6), 0) * -1
   End If
'-----------------------------------------------------------------------------------
End If
TEMP_MATRIX(1, 8) = 0
TEMP_SUM = TEMP_MATRIX(1, 8)
'-----------------------------------------------------------------------------------
For i = 2 To j 'Dynamic evolution with respect to simulation days
       'These are hidden data in column M to O, which represent time(T) and Black-Scholes parameters(d1,d2)
        TEMP_MATRIX(i, 1) = TEMP_MATRIX(i - 1, 1) - DELTA_VAL
       'These are output data in column B to F, whick represent stock price,
       'option price, delta, shares and daily P/L
        TEMP_MATRIX(i, 4) = TEMP_MATRIX(i - 1, 4) * Exp((DRIFT - VOLATILITY ^ 2 / 2) * DELTA_VAL + VOLATILITY * Sqr(DELTA_VAL) * (Sqr(-2 * Log(PSEUDO_RANDOM_FUNC(RND_TYPE))) * Sin(2 * PI_VAL * PSEUDO_RANDOM_FUNC(RND_TYPE))))
        TEMP_MATRIX(i, 2) = (Log(TEMP_MATRIX(i, 4) / STRIKE_PRICE) + (DRIFT + VOLATILITY ^ 2 / 2) * TEMP_MATRIX(i, 1)) / (VOLATILITY * Sqr(TEMP_MATRIX(i, 1)))
        TEMP_MATRIX(i, 3) = (Log(TEMP_MATRIX(i, 4) / STRIKE_PRICE) + (DRIFT - VOLATILITY ^ 2 / 2) * TEMP_MATRIX(i, 1)) / (VOLATILITY * Sqr(TEMP_MATRIX(i, 1)))
'-----------------------------------------------------------------------------------
    If OPTION_TYPE = 1 Then 'Call Option
'-----------------------------------------------------------------------------------
        TEMP_MATRIX(i, 5) = TEMP_MATRIX(i - 1, 4) * CND_FUNC(TEMP_MATRIX(i, 2), CND_TYPE) - STRIKE_PRICE * Exp(-DRIFT * TEMP_MATRIX(i, 1)) * CND_FUNC(TEMP_MATRIX(i, 3), CND_TYPE)
        TEMP_MATRIX(i, 6) = CND_FUNC(TEMP_MATRIX(i, 2), CND_TYPE)
        If SHORT_FLAG = True Then
            TEMP_MATRIX(i, 7) = Round(QUANTITY * TEMP_MATRIX(i, 6), 0)
        Else
            TEMP_MATRIX(i, 7) = Round(QUANTITY * TEMP_MATRIX(i, 6), 0) * -1
        End If
'-----------------------------------------------------------------------------------
    Else 'Put Option
'-----------------------------------------------------------------------------------
        TEMP_MATRIX(i, 5) = STRIKE_PRICE * Exp(-DRIFT * TEMP_MATRIX(i, 1)) * CND_FUNC(-TEMP_MATRIX(i, 3), CND_TYPE) - TEMP_MATRIX(i, 4) * CND_FUNC(-TEMP_MATRIX(i, 2), CND_TYPE)
        TEMP_MATRIX(i, 6) = -1 * CND_FUNC(-1 * TEMP_MATRIX(i, 2), CND_TYPE)
        If SHORT_FLAG = True Then
            TEMP_MATRIX(i, 7) = Round(QUANTITY * TEMP_MATRIX(i, 6), 0)
        Else
            TEMP_MATRIX(i, 7) = Round(QUANTITY * TEMP_MATRIX(i, 6), 0) * -1
        End If
'-----------------------------------------------------------------------------------
    End If
'-----------------------------------------------------------------------------------
    TEMP_MATRIX(i, 8) = TEMP_MATRIX(i - 1, 7) * (TEMP_MATRIX(i, 4) - TEMP_MATRIX(i - 1, 4))
    TEMP_SUM = TEMP_SUM + TEMP_MATRIX(i, 8)
Next i
        
'---------------------------------------------------------------------------------
'---------------------------------------------------------------------------------
Select Case OUTPUT
'---------------------------------------------------------------------------------
Case 0
'---------------------------------------------------------------------------------
    ASSET_OPTION_DELTA_HEDGING_FUNC = TEMP_SUM
'---------------------------------------------------------------------------------
Case Else
'---------------------------------------------------------------------------------
    ASSET_OPTION_DELTA_HEDGING_FUNC = TEMP_MATRIX
'---------------------------------------------------------------------------------
'---------------------------------------------------------------------------------
End Select
'---------------------------------------------------------------------------------
'---------------------------------------------------------------------------------

Exit Function
ERROR_LABEL:
ASSET_OPTION_DELTA_HEDGING_FUNC = Err.number
End Function
