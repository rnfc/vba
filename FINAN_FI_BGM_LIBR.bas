Attribute VB_Name = "FINAN_FI_BGM_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.


'************************************************************************************
'************************************************************************************
'FUNCTION      : BGM_CAPLET_SIMULATION_FUNC

'DESCRIPTION   : This BGM Function computes the value of a caplet given
'the forward rate scenario, using Monte Carlo.
'The function calculates each draw's value of the caplet, and processes
'these draws and computes the value of the caplet.
'-------------------------------------------------------------------------------
'-------------------------------------------------------------------------------
'BGM or LIBOR Market Model
'Similar to Black, Derman and Toy in that we fit a set of market prices.
'Also consistent with Black's model (for caplets).

'The market model approach is currently very popular on Wall Street, despite
'being rather convoluted.  Its major advantages are that
'it works directly with observed forward rates (so that there
'are no unobserved factors) and that it is consistent with
'Black's model.

'The key is that each forward rate moves to the next period as a distinct
'process.  We designate a future-dated zero-coupon bond as numeraire.
'This means that the drift of each forward rate must be
'adjusted to reflect the (arbitrary) numeraire.

'As an example, if we just had one date then the numeraire would be 1,
'and the drift would be 0, as the expected future spot rate
'is the forward rate--as in Black's Model.
'-------------------------------------------------------------------------------
'-------------------------------------------------------------------------------

'LIBRARY       : FIXED_INCOME
'GROUP         : BGM
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 10-06-2008
'************************************************************************************
'************************************************************************************

Function BGM_CAPLET_SIMULATION_FUNC(ByVal SIGMA As Variant, _
ByVal STRIKE As Variant, _
ByVal TENOR As Variant, _
ByVal RESET_TENOR As Variant, _
ByRef SPOT_RNG As Variant, _
ByRef TENOR_RNG As Variant, _
Optional ByVal DELTA_TENOR As Variant = 0.25, _
Optional ByVal NOTIONAL As Variant = 10000, _
Optional ByVal nLOOPS As Variant = 1000, _
Optional ByVal OUTPUT As Variant = 0)
    
'SIGMA: Read in SIGMA
'DELTA_TENOR: Tenor (Also Delta-TENOR for the Monte Carlo)
'TENOR: Expiration of the caplet (number of years)
'STRIKE: Strike Rate of the Caplet
'nLOOPS: Loops in the simulation
'RESET_TENOR: This is how far ahead we go (reset date on caplet)

Dim h As Long
Dim i As Long
Dim j As Long
Dim k As Long

Dim SROW As Long
Dim NROWS As Long
Dim NSIZE As Long

Dim TEMP_SD As Double
Dim TEMP_VAR As Double
Dim TEMP_SUM As Double
Dim TEMP_MULT As Double

Dim NUMER_VAL As Double
Dim DENOM_VAL As Double

Dim NORM_RND_VAL As Double

Dim TEMP_ARR As Variant
Dim ATEMP_MATRIX As Variant
Dim BTEMP_MATRIX As Variant

Dim SPOT_VECTOR As Variant
Dim TENOR_VECTOR As Variant
Dim FORWARD_VECTOR As Variant

On Error GoTo ERROR_LABEL

SPOT_VECTOR = SPOT_RNG
If UBound(SPOT_VECTOR, 1) = 1 Then
    SPOT_VECTOR = MATRIX_TRANSPOSE_FUNC(SPOT_VECTOR)
End If

TENOR_VECTOR = TENOR_RNG
If UBound(TENOR_VECTOR, 1) = 1 Then
    TENOR_VECTOR = MATRIX_TRANSPOSE_FUNC(TENOR_VECTOR)
End If

SROW = LBound(TENOR_VECTOR, 1)
NROWS = UBound(TENOR_VECTOR, 1)

If SROW = 0 Then
    NSIZE = NROWS
Else
    NSIZE = NROWS - SROW
End If

ReDim TEMP_ARR(0 To NSIZE)
ReDim FORWARD_VECTOR(0 To NSIZE, 1 To 1)
ReDim ATEMP_MATRIX(0 To NSIZE, 0 To NSIZE)
ReDim BTEMP_MATRIX(1 To nLOOPS, 1 To 2)

For i = 0 To NSIZE
    If i = 0 Then
        FORWARD_VECTOR(i, 1) = SPOT_VECTOR(SROW, 1)
            SPOT_VECTOR(SROW, 1) = 1 / (1 + SPOT_VECTOR(SROW, 1) _
            / (1 / DELTA_TENOR)) ^ (i + 1)
    Else
        SPOT_VECTOR(SROW + i, 1) = 1 / (1 + SPOT_VECTOR(SROW + i, 1) _
            / (1 / DELTA_TENOR)) ^ (i + 1)
            FORWARD_VECTOR(i, 1) = (SPOT_VECTOR(SROW + i - 1, 1) / _
             SPOT_VECTOR(SROW + i, 1) - 1) / (1 / (1 / DELTA_TENOR))
             'Price (Discount)
            'The assumption of BGM is that the forward rate follows a log-normal
            'process (In the equivalent risk-neutral world associated with the
            'forward-dated numeraire)
    End If
    ATEMP_MATRIX(0, i) = FORWARD_VECTOR(i, 1)
    'Load LIBOR forward rates
Next i
    
TEMP_SD = SIGMA * Sqr(DELTA_TENOR)
TEMP_VAR = SIGMA * SIGMA
TEMP_SUM = 0

For i = 1 To nLOOPS
      For j = 1 To RESET_TENOR 'This is how far ahead we go (reset date on caplet)
        TEMP_MULT = 1
        For k = 1 To NSIZE  'This is the total number of nodes.
            TEMP_ARR(k) = 0
            For h = k + 1 To NSIZE
                NUMER_VAL = DELTA_TENOR * TEMP_VAR * ATEMP_MATRIX(j - 1, h)
                DENOM_VAL = 1 + DELTA_TENOR * ATEMP_MATRIX(j - 1, h)
                TEMP_ARR(k) = TEMP_ARR(k) - NUMER_VAL / DENOM_VAL
            Next h
            NORM_RND_VAL = RANDOM_POLAR_MARSAGLIA_FUNC() '--> Marsaglia's Normal Rand Draw
            ATEMP_MATRIX(j, k) = ATEMP_MATRIX(j - 1, k) * _
                                Exp((TEMP_ARR(k) - 0.5 * TEMP_VAR) * _
                                DELTA_TENOR + NORM_RND_VAL * TEMP_SD)
        Next k
        For h = j + 1 To NSIZE
            TEMP_MULT = TEMP_MULT / (1 + ATEMP_MATRIX(j, h) * DELTA_TENOR)
        Next h
    Next j
    
    'This is the value of the option today from this particular path.
    BTEMP_MATRIX(i, 1) = i
    BTEMP_MATRIX(i, 2) = MAXIMUM_FUNC(ATEMP_MATRIX(1, TENOR / DELTA_TENOR) - STRIKE, 0) _
                    / TEMP_MULT 'Caplet Value
    TEMP_SUM = TEMP_SUM + BTEMP_MATRIX(i, 2) 'compute the Average of the Monte Carlo Draws
Next i

Select Case OUTPUT
    Case 0 'Multiply the average times the Tenor and Notional Factor
        BGM_CAPLET_SIMULATION_FUNC = (TEMP_SUM / nLOOPS) * _
                      (DELTA_TENOR) * NOTIONAL * _
                       SPOT_VECTOR(RESET_TENOR + SROW, 1) 'Price (Discount)
        'Bring the Expected future value of the caplet to today by discounting
        'at the spot rate
    Case Else 'Draws Array
        BGM_CAPLET_SIMULATION_FUNC = BTEMP_MATRIX
End Select

Exit Function
ERROR_LABEL:
BGM_CAPLET_SIMULATION_FUNC = Err.number
End Function
