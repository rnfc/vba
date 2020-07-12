Attribute VB_Name = "FINAN_ASSET_SIMUL_BOOTS_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.

Function ASSET_BOOSTRAP_PRICES_DISTRIBUTION_FUNC(ByRef PRICES_RNG As Variant, _
Optional ByVal NO_PERIODS As Long = 30)

'NO_PERIODS --> FORWARD!!!!!
'boostrap Future Price Distribution (with the same simulated data)

Dim i As Long
Dim j As Long
Dim k As Long

Dim NBINS As Long
Dim NROWS As Long

Dim P_VAL As Double
Dim BIN_MIN As Double
Dim BIN_MAX As Double
Dim BIN_WIDTH As Double

Dim FREQUENCY_VECTOR As Variant
Dim PRICES_VECTOR As Variant
Dim TEMP_VECTOR As Variant

On Error GoTo ERROR_LABEL

PRICES_VECTOR = PRICES_RNG
If UBound(PRICES_VECTOR, 1) = 1 Then
    PRICES_VECTOR = MATRIX_TRANSPOSE_FUNC(PRICES_VECTOR)
End If
NROWS = UBound(PRICES_VECTOR, 1)
NROWS = NROWS - 1 'Exclude Po

ReDim TEMP_VECTOR(1 To NROWS, 1 To 1)

Randomize
BIN_MIN = 2 ^ 52: BIN_MAX = -2 ^ 52
For j = 1 To NROWS
    P_VAL = PRICES_VECTOR(1, 1)        ' start with Po
    For i = 1 To NO_PERIODS         ' apply T returns forward
        k = (NROWS - 1) * Rnd + 2
        P_VAL = P_VAL * PRICES_VECTOR(k, 1) / PRICES_VECTOR(k - 1, 1)
    Next i
    TEMP_VECTOR(j, 1) = P_VAL         ' save price
    If TEMP_VECTOR(j, 1) < BIN_MIN Then: BIN_MIN = TEMP_VECTOR(j, 1)
    If TEMP_VECTOR(j, 1) > BIN_MAX Then: BIN_MAX = TEMP_VECTOR(j, 1)
Next j

FREQUENCY_VECTOR = HISTOGRAM_BIN_LIMITS_FUNC(BIN_MIN, BIN_MAX, NROWS, 3)
BIN_WIDTH = FREQUENCY_VECTOR(LBound(FREQUENCY_VECTOR))
BIN_MIN = FREQUENCY_VECTOR(LBound(FREQUENCY_VECTOR) + 1)
NBINS = FREQUENCY_VECTOR(LBound(FREQUENCY_VECTOR) + 2)

FREQUENCY_VECTOR = HISTOGRAM_FREQUENCY_FUNC(TEMP_VECTOR, NBINS, BIN_MIN, BIN_WIDTH, 1)
NBINS = UBound(FREQUENCY_VECTOR, 1)
ReDim Preserve FREQUENCY_VECTOR(1 To NBINS, 1 To 3)
For j = 1 To NBINS
    FREQUENCY_VECTOR(j, 3) = FREQUENCY_VECTOR(j, 2) / NROWS
Next j

ASSET_BOOSTRAP_PRICES_DISTRIBUTION_FUNC = FREQUENCY_VECTOR

Exit Function
ERROR_LABEL:
ASSET_BOOSTRAP_PRICES_DISTRIBUTION_FUNC = Err.number
End Function
