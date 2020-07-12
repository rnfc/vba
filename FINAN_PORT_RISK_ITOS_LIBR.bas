Attribute VB_Name = "FINAN_PORT_RISK_ITOS_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.

'************************************************************************************
'************************************************************************************
'FUNCTION      : PORT_ITO_PROBABILITY_LESS_FUNC

'DESCRIPTION   : Probability that your portfolio will be less than or equal to
'$Px1000 in Z periods"

'LIBRARY       : PORTFOLIO
'GROUP         : RISK_ITOS
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/08/2008
'************************************************************************************
'************************************************************************************

Function PORT_ITO_PROBABILITY_LESS_FUNC(ByVal INITIAL_VAL As Double, _
ByVal PORT_MIN_VAL As Double, _
ByVal PORT_MAX_VAL As Double, _
ByVal PERIOD_MEAN_VAL As Double, _
ByVal PERIOD_SIGMA_VAL As Double, _
ByVal NO_PERIODS As Integer)

Dim i As Long

Dim PI_VAL As Double
Dim F_VAL As Double
Dim DPORT_VAL As Double

Dim TEMP_STR As String
Dim TEMP_VAL As Double
Dim TEMP_SUM As Double

Dim TEMP_MATRIX As Variant

On Error GoTo ERROR_LABEL

F_VAL = 1000

ReDim TEMP_MATRIX(0 To NO_PERIODS, 1 To 4)

PI_VAL = 3.14159265358979

TEMP_VAL = INITIAL_VAL / F_VAL
TEMP_MATRIX(0, 1) = "PORT VALUE"
TEMP_MATRIX(0, 2) = "MASS PROB"
TEMP_MATRIX(0, 3) = "CUMUL PROB"
TEMP_MATRIX(0, 4) = "SUMMARY"

DPORT_VAL = PORT_MIN_VAL / F_VAL
TEMP_MATRIX(1, 1) = DPORT_VAL
DPORT_VAL = (PORT_MAX_VAL / F_VAL - PORT_MIN_VAL / F_VAL) / NO_PERIODS
TEMP_MATRIX(1, 2) = DPORT_VAL / (PERIOD_SIGMA_VAL * Sqr(2 * PI_VAL * NO_PERIODS)) / _
                    TEMP_MATRIX(1, 1) * _
                    Exp(-(1 / (2 * NO_PERIODS * PERIOD_SIGMA_VAL ^ 2)) * _
                    (Log(TEMP_MATRIX(1, 1) / TEMP_VAL) - _
                    (PERIOD_MEAN_VAL - 0.5 * _
                    PERIOD_SIGMA_VAL ^ 2) * NO_PERIODS) ^ 2)

TEMP_SUM = TEMP_MATRIX(1, 2)
TEMP_MATRIX(1, 3) = TEMP_SUM

TEMP_STR = "There's an " & Format(TEMP_MATRIX(1, 3), "0.00%") & _
                    " probability that your " & Format(INITIAL_VAL, "0,000") & _
                    " portfolio will be less than " & _
                    Format(TEMP_MATRIX(1, 1) * F_VAL, "0,000") & _
                    " in " & Format(NO_PERIODS, "0") & " periods."

TEMP_MATRIX(1, 4) = TEMP_STR


TEMP_STR = ""

For i = 2 To NO_PERIODS
    TEMP_MATRIX(i, 1) = TEMP_MATRIX(i - 1, 1) + DPORT_VAL
    TEMP_MATRIX(i, 2) = DPORT_VAL / (PERIOD_SIGMA_VAL * Sqr(2 * PI_VAL * NO_PERIODS)) / _
                    TEMP_MATRIX(i, 1) * _
                    Exp(-(1 / (2 * NO_PERIODS * PERIOD_SIGMA_VAL ^ 2)) * _
                    (Log(TEMP_MATRIX(i, 1) / TEMP_VAL) - _
                    (PERIOD_MEAN_VAL - 0.5 * _
                    PERIOD_SIGMA_VAL ^ 2) * NO_PERIODS) ^ 2)
                    
    TEMP_SUM = TEMP_SUM + TEMP_MATRIX(i, 2)
    TEMP_MATRIX(i, 3) = TEMP_SUM

    TEMP_STR = "There's an " & Format(TEMP_MATRIX(i, 3), "0.00%") & _
                    " probability that your " & Format(INITIAL_VAL, "0,000") & _
                    " portfolio will be less than " & _
                    Format(TEMP_MATRIX(i, 1) * F_VAL, "0,000") & _
                    " in " & Format(NO_PERIODS, "0") & " periods."

    TEMP_MATRIX(i, 4) = TEMP_STR
Next i

PORT_ITO_PROBABILITY_LESS_FUNC = TEMP_MATRIX

Exit Function
ERROR_LABEL:
PORT_ITO_PROBABILITY_LESS_FUNC = Err.number
End Function
