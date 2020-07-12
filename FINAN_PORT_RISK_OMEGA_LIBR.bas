Attribute VB_Name = "FINAN_PORT_RISK_OMEGA_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.



'************************************************************************************
'************************************************************************************
'FUNCTION      : PORT_RISK_REWARD_OMEGA_FUNC

'Note that the "Risk / Reward" ratio whereas Omega is a ratio which reflects "Gains / Losses".
'Their relationship is : (1/F(r) - 1) / Omega   where "r" is the threshold return (which may
'be taken as a Risk-free Rate).

'LIBRARY       : PORTFOLIO
'GROUP         : MOMENTS_OMEGA
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 21/01/2009
'************************************************************************************
'************************************************************************************


Function PORT_RISK_REWARD_OMEGA_FUNC(ByRef RETURNS_RNG As Variant, _
Optional ByVal CASH_RATE As Double = 0.04, _
Optional ByVal OUTPUT As Integer = 1)

Dim i As Long
Dim j As Long
Dim NROWS As Long

Dim MIN_VAL As Double
Dim MAX_VAL As Double

Dim NBINS As Long
Dim BIN_MIN As Double
Dim BIN_WIDTH As Double

Dim MEAN_VAL As Double
Dim SIGMA_VAL As Double

Dim TEMP1_SUM As Double
Dim TEMP2_SUM As Double

Dim AREA_VAL As Double ' at cash rate

Dim I1_VAL As Double
Dim I2_VAL As Double

Dim OMEGA_VAL As Double
Dim RISK_REWARD_VAL As Double

Dim RETURNS_VECTOR As Variant
Dim FREQUENCY_VECTOR As Variant

On Error GoTo ERROR_LABEL

RETURNS_VECTOR = RETURNS_RNG
If UBound(RETURNS_VECTOR, 1) = 1 Then
    RETURNS_VECTOR = MATRIX_TRANSPOSE_FUNC(RETURNS_VECTOR)
End If
NROWS = UBound(RETURNS_VECTOR, 1)

FREQUENCY_VECTOR = DATA_BASIC_MOMENTS_FUNC(RETURNS_VECTOR, 0, 0, 0.05, 0)
MIN_VAL = FREQUENCY_VECTOR(1, 2)
MAX_VAL = FREQUENCY_VECTOR(1, 3)
MEAN_VAL = FREQUENCY_VECTOR(1, 4)
SIGMA_VAL = FREQUENCY_VECTOR(1, 7)
FREQUENCY_VECTOR = HISTOGRAM_BIN_LIMITS_FUNC(MIN_VAL, MAX_VAL, NROWS, 3)

BIN_WIDTH = FREQUENCY_VECTOR(LBound(FREQUENCY_VECTOR))
BIN_MIN = FREQUENCY_VECTOR(LBound(FREQUENCY_VECTOR) + 1)
NBINS = FREQUENCY_VECTOR(LBound(FREQUENCY_VECTOR) + 2)
FREQUENCY_VECTOR = HISTOGRAM_FREQUENCY_FUNC(RETURNS_VECTOR, NBINS, BIN_MIN, BIN_WIDTH, 1)

'FREQUENCY_VECTOR = Range("NICO")
'BIN_WIDTH = (MAX_VAL - MIN_VAL) / 10

NBINS = UBound(FREQUENCY_VECTOR, 1)
ReDim Preserve FREQUENCY_VECTOR(1 To NBINS, 1 To 4)
'Bins
'f(r)
'F(r)
'F(r)


TEMP1_SUM = 0: TEMP2_SUM = 0
For i = 1 To NBINS
    If i > 1 Then
        If FREQUENCY_VECTOR(i, 1) > CASH_RATE And FREQUENCY_VECTOR(i - 1, 1) <= CASH_RATE Then
            j = i - 1
        End If
    End If
    TEMP1_SUM = TEMP1_SUM + FREQUENCY_VECTOR(i, 2)
    FREQUENCY_VECTOR(i, 3) = TEMP1_SUM / NROWS
    TEMP2_SUM = TEMP2_SUM + FREQUENCY_VECTOR(i, 3)
    FREQUENCY_VECTOR(i, 4) = TEMP2_SUM * BIN_WIDTH
Next i

If j = 0 Then: GoTo ERROR_LABEL ' Change Cash/Threshold Rate
AREA_VAL = FREQUENCY_VECTOR(j, 4) + (FREQUENCY_VECTOR(j + 1, 4) - FREQUENCY_VECTOR(j, 4)) / (FREQUENCY_VECTOR(j + 1, 1) - FREQUENCY_VECTOR(j, 1)) * (CASH_RATE - FREQUENCY_VECTOR(j, 1))
I1_VAL = AREA_VAL
I2_VAL = (MAX_VAL - MIN_VAL) - (FREQUENCY_VECTOR(NBINS, 4) - AREA_VAL)

OMEGA_VAL = I2_VAL / I1_VAL 'Maximize by changing weight allocation
RISK_REWARD_VAL = (1 / AREA_VAL - 1) / OMEGA_VAL 'Minimize

Select Case OUTPUT
Case 0
    PORT_RISK_REWARD_OMEGA_FUNC = Array(OMEGA_VAL, RISK_REWARD_VAL)
Case Else
    PORT_RISK_REWARD_OMEGA_FUNC = FREQUENCY_VECTOR
End Select

Exit Function
ERROR_LABEL:
PORT_RISK_REWARD_OMEGA_FUNC = Err.number
End Function
