Attribute VB_Name = "FINAN_PORT_WEIGHTS_CAGR_LIBR"


Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.

'************************************************************************************
'************************************************************************************
'FUNCTION      : PORT_WEIGHTS_CAGR_FUNC
'DESCRIPTION   : CAGR versus Stock Allocation
'I asked, on a couple of investment forums, for other "magic formulas" that
'suggest stock allocations ... and got these:
'S = D * 2     where D = maximum allowable decline to portfolio
'and S = stock allocation
'and
'18 x dividend yield
'and
'2p-1   where p is the probability of the asset achieving your
'investment objective

'LIBRARY       : PORTFOLIO
'GROUP         : WEIGHTS_CAGR
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 21/01/2009
'************************************************************************************
'************************************************************************************

Function PORT_WEIGHTS_CAGR_FUNC(ByVal CASH_RATE As Double, _
ByVal MEAN_VAL As Double, _
ByVal VOLATILITY As Double, _
Optional ByVal NBINS As Long = 21, _
Optional ByVal MIN_WEIGHT As Double = 0, _
Optional ByVal DELTA_WEIGHT As Double = 0.05, _
Optional ByVal COUNT_BASIS As Double = 12)

'CASH_RATE --> Bank Rate Annualized
'MEAN --> Period
'Volatility --> Period

Dim k As Long

Dim ATEMP_VAL As Double
Dim BTEMP_VAL As Double

Dim ANNUAL_MEAN_VAL As Double
Dim ANNUAL_VOLATILITY_VAL As Double

Dim TEMP_VECTOR As Variant

On Error GoTo ERROR_LABEL

ANNUAL_MEAN_VAL = MEAN_VAL * COUNT_BASIS
ANNUAL_VOLATILITY_VAL = Sqr(COUNT_BASIS) * VOLATILITY

ATEMP_VAL = (ANNUAL_MEAN_VAL - CASH_RATE) / (1 + CASH_RATE)
BTEMP_VAL = ANNUAL_VOLATILITY_VAL / (1 + CASH_RATE)

ReDim TEMP_VECTOR(0 To NBINS, 1 To 2)

TEMP_VECTOR(0, 1) = "WEIGHT"
TEMP_VECTOR(0, 2) = "CAGR"

For k = 1 To NBINS
    
    If k = 1 Then
        TEMP_VECTOR(k, 1) = MIN_WEIGHT
    Else
        TEMP_VECTOR(k, 1) = TEMP_VECTOR(k - 1, 1) + DELTA_WEIGHT
    End If
    TEMP_VECTOR(k, 2) = Exp(TEMP_VECTOR(k, 1) * ATEMP_VAL - 0.5 * TEMP_VECTOR(k, 1) ^ 2 * (BTEMP_VAL ^ 2 + ATEMP_VAL ^ 2))
Next k

PORT_WEIGHTS_CAGR_FUNC = TEMP_VECTOR

Exit Function
ERROR_LABEL:
PORT_WEIGHTS_CAGR_FUNC = Err.number
End Function

