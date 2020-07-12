Attribute VB_Name = "FINAN_ASSET_SIMUL_BMV_END_LIBR"

'--------------------------------------------------------------------------------------------
'--------------------------------------------------------------------------------------------
Option Explicit
Option Base 1
'--------------------------------------------------------------------------------------------
'--------------------------------------------------------------------------------------------

'Simulated portfolio with given expected return and volatility

Function ASSET_DAILY_BMV_EMV_FUNC(Optional ByVal INITIAL_VAL As Double = 100, _
Optional ByVal EXPECTED_RETURN_VAL As Double = 15, _
Optional ByVal VOLATILITY_VAL As Double = 20, _
Optional ByVal NO_PERIODS As Long = 10, _
Optional ByVal COUNT_BASIS As Long = 366, _
Optional ByVal nLOOPS As Long = 10, _
Optional ByVal FACTOR_VAL As Double = 100)

'INITIAL_VAL --> Initial Portfolio Value
'Exp. Return [% p.a.]    15
'Volatility [% p.a.] 20
'Days per Year

Dim i As Long
Dim j As Long

Dim RETURN_VAL As Double
Dim BMV_VAL As Double
Dim EMV_VAL As Double
Dim TEMP_MATRIX As Variant

On Error GoTo ERROR_LABEL

ReDim TEMP_MATRIX(0 To nLOOPS, 1 To 5)
TEMP_MATRIX(0, 1) = "ITERATION"
TEMP_MATRIX(0, 2) = "PERIOD"
TEMP_MATRIX(0, 3) = "RETURN"
TEMP_MATRIX(0, 4) = "BMV"
TEMP_MATRIX(0, 5) = "EMV"

For j = 1 To nLOOPS
    GoSub RETURN_LINE
    BMV_VAL = INITIAL_VAL
    EMV_VAL = BMV_VAL * (1 + RETURN_VAL / FACTOR_VAL)
    For i = 2 To NO_PERIODS
        GoSub RETURN_LINE
        BMV_VAL = EMV_VAL '+ CONTRIBUTION
        EMV_VAL = BMV_VAL * (1 + RETURN_VAL / FACTOR_VAL)
    Next i
    TEMP_MATRIX(j, 1) = j
    TEMP_MATRIX(j, 2) = NO_PERIODS
    TEMP_MATRIX(j, 3) = RETURN_VAL / FACTOR_VAL
    TEMP_MATRIX(j, 4) = BMV_VAL
    TEMP_MATRIX(j, 5) = EMV_VAL
Next j
ASSET_DAILY_BMV_EMV_FUNC = TEMP_MATRIX
Exit Function
RETURN_LINE:
    RETURN_VAL = EXPECTED_RETURN_VAL * (1 / COUNT_BASIS) + _
                 NORMSINV_FUNC(Rnd(), 0, 1, 0) * _
                 VOLATILITY_VAL * Sqr(1 / COUNT_BASIS)
Return
ERROR_LABEL:
ASSET_DAILY_BMV_EMV_FUNC = Err.number
End Function


