Attribute VB_Name = "FINAN_DERIV_BS_IMPLIED_LIBR"

'---------------------------------------------------------------------------------
'---------------------------------------------------------------------------------
Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.
'---------------------------------------------------------------------------------

Private Const PUB_DBL_EPSILON = 2.22044604925031E-16
Private Const PUB_DBL_MIN = 2.2250738585072E-308
Private Const PUB_DBL_MIN_EXP = (-1021)
Private Const PUB_PI_VAL = 3.14159265358979
Private Const PUB_BY_LN2_VAL = 1.44269504088896    ' = 1/ln(2)
Private Const INV_SQRT_2PI_VAL = 0.398942280401433 ' = 1/sqrt(2*PUB_PI_VAL)

'---------------------------------------------------------------------------------
Private PUB_PARAM_ARR(0 To 7) As Variant
Private Const PUB_EPSILON As Double = 2 ^ 52 '0

'************************************************************************************
'************************************************************************************
'FUNCTION      : BLACK_SCHOLES_IMPLIED_VOLATILITY_FUNC
'DESCRIPTION   : BS Implied Sigma Function with Bisec Method
'LIBRARY       : DERIVATIVES
'GROUP         : BS_IMPLIED
'ID            :
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   :
'************************************************************************************
'************************************************************************************

Function BLACK_SCHOLES_IMPLIED_VOLATILITY_FUNC( _
ByVal PREMIUM As Double, _
ByVal SPOT As Double, _
ByVal STRIKE As Double, _
ByVal EXPIRATION As Double, _
ByVal RATE As Double, _
ByVal CARRY As Double, _
Optional ByVal OPTION_FLAG As Integer = 1, _
Optional ByVal LOWER_VAL As Double = 0#, _
Optional ByVal UPPER_VAL As Double = 1#, _
Optional ByVal CND_TYPE As Integer = 0)

'LOWER_VAL min implied volatility estimates
'UPPER_VAL max implied volatility estimates

'OPTION_FLAG = 1 --> CALL_OPTION
'OPTION_FLAG = -1 --> PUT_OPTION

'CARRY = COST_OF_CARRY = (RISK_FREE - DIVD_YIELD)

Dim CONVERG_VAL As Integer
Dim COUNTER As Long
Dim nLOOPS As Long
Dim Y_VAL As Double
Dim tolerance As Double

On Error GoTo ERROR_LABEL
PUB_PARAM_ARR(0) = PREMIUM
PUB_PARAM_ARR(1) = SPOT
PUB_PARAM_ARR(2) = STRIKE
PUB_PARAM_ARR(3) = EXPIRATION
PUB_PARAM_ARR(4) = RATE
PUB_PARAM_ARR(5) = CARRY
PUB_PARAM_ARR(6) = OPTION_FLAG
PUB_PARAM_ARR(7) = CND_TYPE

nLOOPS = 500
tolerance = 0.00000000000001

Y_VAL = MULLER_ZERO_FUNC(LOWER_VAL, UPPER_VAL, "BLACK_SCHOLES_IMPLIED_VOLATILITY1_OBJ_FUNC", _
        CONVERG_VAL, COUNTER, nLOOPS, tolerance)
If Y_VAL = PUB_EPSILON Or CONVERG_VAL <> 0 Then: GoTo ERROR_LABEL

BLACK_SCHOLES_IMPLIED_VOLATILITY_FUNC = Y_VAL
Exit Function
ERROR_LABEL:
BLACK_SCHOLES_IMPLIED_VOLATILITY_FUNC = PUB_EPSILON
End Function

Function BLACK_SCHOLES_IMPLIED_VOLATILITY1_OBJ_FUNC(ByVal X_VAL As Variant)
Dim Y_VAL As Double
On Error GoTo ERROR_LABEL
Y_VAL = GENERALIZED_BLACK_SCHOLES_FUNC(PUB_PARAM_ARR(1), PUB_PARAM_ARR(2), PUB_PARAM_ARR(3), PUB_PARAM_ARR(4), _
                                       PUB_PARAM_ARR(5), X_VAL, PUB_PARAM_ARR(6), PUB_PARAM_ARR(7))
BLACK_SCHOLES_IMPLIED_VOLATILITY1_OBJ_FUNC = Y_VAL - PUB_PARAM_ARR(0) 'Abs(Y_VAL - PUB_PARAM_ARR(0)) ^ 2
Exit Function
ERROR_LABEL:
BLACK_SCHOLES_IMPLIED_VOLATILITY1_OBJ_FUNC = PUB_EPSILON
End Function

Function BLACK_SCHOLES_IMPLIED_VOLATILITY_TEST_FUNC(ByVal PREMIUM As Double, _
ByVal SPOT As Double, _
ByVal STRIKE As Double, _
ByVal EXPIRATION As Double, _
ByVal RATE As Double, _
Optional ByVal OPTION_FLAG As Integer = 1, _
Optional ByVal CND_TYPE As Integer = 1, _
Optional ByVal LOWER_VAL As Double = 0.0000001, _
Optional ByVal UPPER_VAL As Double = 10, _
Optional ByVal tolerance As Double = 0.0000001, _
Optional ByVal nLOOPS As Long = 100)

On Error GoTo ERROR_LABEL

PUB_PARAM_ARR(0) = PREMIUM
PUB_PARAM_ARR(1) = SPOT
PUB_PARAM_ARR(2) = STRIKE
PUB_PARAM_ARR(3) = EXPIRATION
PUB_PARAM_ARR(4) = RATE
PUB_PARAM_ARR(5) = RATE 'CARRY
PUB_PARAM_ARR(6) = OPTION_FLAG
PUB_PARAM_ARR(7) = CND_TYPE

BLACK_SCHOLES_IMPLIED_VOLATILITY_TEST_FUNC = CALL_TEST_ZERO_FRAME_FUNC(LOWER_VAL, UPPER_VAL, _
"BLACK_SCHOLES_IMPLIED_VOLATILITY1_OBJ_FUNC", nLOOPS, tolerance)

Exit Function
ERROR_LABEL:
BLACK_SCHOLES_IMPLIED_VOLATILITY_TEST_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : BLACK_SCHOLES_IMPLIED_VOLATILITY_TABLE_FUNC
'DESCRIPTION   : BS IMPLIED_SIGMA_TABLE
'LIBRARY       : DERIVATIVES
'GROUP         : BS_IMPLIED
'ID            :
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   :
'************************************************************************************
'************************************************************************************

Function BLACK_SCHOLES_IMPLIED_VOLATILITY_TABLE_FUNC(ByVal PREMIUM As Double, _
ByVal SPOT As Double, _
ByVal STRIKE As Double, _
ByVal EXPIRATION As Double, _
ByVal RATE As Double, _
ByVal LOWER_VAL As Double, _
ByVal UPPER_VAL As Double, _
ByVal DELTA_IMP_VAL As Double, _
Optional ByVal OPTION_FLAG As Integer = 1, _
Optional ByVal CND_TYPE As Integer = 0)

'GUESS_SIGMA = (2 * Abs(Log(SPOT / STRIKE) + RATE * EXPIRATION) / EXPIRATION) ^ (1 / 2)

Dim i As Long
Dim nSTEPS As Long

Dim TEMP_SIGMA As Double
Dim TEMP_MATRIX As Variant

On Error GoTo ERROR_LABEL

nSTEPS = (UPPER_VAL - LOWER_VAL) / DELTA_IMP_VAL + 1

ReDim TEMP_MATRIX(0 To nSTEPS, 1 To 3)

TEMP_MATRIX(0, 1) = "SIGMA"
TEMP_MATRIX(0, 2) = "ERROR"
TEMP_MATRIX(0, 3) = "PREMIUM"

TEMP_SIGMA = UPPER_VAL

For i = nSTEPS To 1 Step -1
    TEMP_MATRIX(i, 1) = TEMP_SIGMA
    TEMP_MATRIX(i, 2) = (BLACK_SCHOLES_OPTION_FUNC(SPOT, STRIKE, EXPIRATION, RATE, TEMP_SIGMA, OPTION_FLAG, CND_TYPE) / SPOT) - (PREMIUM / SPOT)
    TEMP_MATRIX(i, 3) = TEMP_MATRIX(i, 2) * SPOT + PREMIUM
    TEMP_SIGMA = TEMP_SIGMA - DELTA_IMP_VAL
Next i

BLACK_SCHOLES_IMPLIED_VOLATILITY_TABLE_FUNC = TEMP_MATRIX
    
Exit Function
ERROR_LABEL:
BLACK_SCHOLES_IMPLIED_VOLATILITY_TABLE_FUNC = PUB_EPSILON
End Function


'The following routines compute Black-Scholes prices and retrieve volatility.
'They are robust (working with the option premium and switch to 'normed'
'situations) and the volatility is computed in the spirit of a fairly good
'initial guess (similar to Jaeckel).

'Further, they also show how one can increase the usual solution for vol
'and still gets given prices. This (partially) solves the problem, that vol
'numerical is not well-defined as the inverse of a price. It even works in
'extreme situations (like vol ~ 10% and small time or very far off the money).

'A more sound solution has to use C code (or similar), but it is just a stripped
'down version of that. Of course that depends on the quality of the pricing
'function and to judge it one can not use Excel.

Function BS_VOLATILITY_INCREASED_FUNC(ByRef SPOT_RNG As Variant, _
ByRef STRIKE_RNG As Variant, _
ByRef EXPIRATION_RNG As Variant, _
ByRef RATE_RNG As Variant, _
ByRef VOLATILITY_RNG As Variant, _
Optional ByVal OUTPUT As Integer = 0)

Dim i As Long
Dim j As Long
Dim k As Long
Dim NROWS As Long
Dim NCOLUMNS As Long

Dim HEADINGS_STR As String
Dim SPOT_VECTOR As Variant
Dim STRIKE_VECTOR As Variant
Dim EXPIRATION_VECTOR As Variant
Dim RATE_VECTOR As Variant
Dim VOLATILITY_VECTOR As Variant

Dim TEMP_MATRIX As Variant

On Error GoTo ERROR_LABEL

If IsArray(SPOT_RNG) = True Then
    SPOT_VECTOR = SPOT_RNG
    If UBound(SPOT_VECTOR, 1) = 1 Then
        SPOT_VECTOR = MATRIX_TRANSPOSE_FUNC(SPOT_VECTOR)
    End If
Else
    ReDim SPOT_VECTOR(1 To 1, 1 To 1)
    SPOT_VECTOR(1, 1) = SPOT_RNG
End If
NROWS = UBound(SPOT_VECTOR, 1)

If IsArray(STRIKE_RNG) = True Then
    STRIKE_VECTOR = STRIKE_RNG
    If UBound(STRIKE_VECTOR, 1) = 1 Then
        STRIKE_VECTOR = MATRIX_TRANSPOSE_FUNC(STRIKE_VECTOR)
    End If
Else
    ReDim STRIKE_VECTOR(1 To 1, 1 To 1)
    STRIKE_VECTOR(1, 1) = STRIKE_RNG
End If
If UBound(STRIKE_VECTOR, 1) <> NROWS Then: GoTo ERROR_LABEL

If IsArray(EXPIRATION_RNG) = True Then
    EXPIRATION_VECTOR = EXPIRATION_RNG
    If UBound(EXPIRATION_VECTOR, 1) = 1 Then
        EXPIRATION_VECTOR = MATRIX_TRANSPOSE_FUNC(EXPIRATION_VECTOR)
    End If
Else
    ReDim EXPIRATION_VECTOR(1 To 1, 1 To 1)
    EXPIRATION_VECTOR(1, 1) = EXPIRATION_RNG
End If
If UBound(EXPIRATION_VECTOR, 1) <> NROWS Then: GoTo ERROR_LABEL

If IsArray(RATE_RNG) = True Then
    RATE_VECTOR = RATE_RNG
    If UBound(RATE_VECTOR, 1) = 1 Then
        RATE_VECTOR = MATRIX_TRANSPOSE_FUNC(RATE_VECTOR)
    End If
Else
    ReDim RATE_VECTOR(1 To 1, 1 To 1)
    RATE_VECTOR(1, 1) = RATE_RNG
End If
If UBound(RATE_VECTOR, 1) <> NROWS Then: GoTo ERROR_LABEL

If IsArray(VOLATILITY_RNG) = True Then
    VOLATILITY_VECTOR = VOLATILITY_RNG
    If UBound(VOLATILITY_VECTOR, 1) = 1 Then
        VOLATILITY_VECTOR = MATRIX_TRANSPOSE_FUNC(VOLATILITY_VECTOR)
    End If
Else
    ReDim VOLATILITY_VECTOR(1 To 1, 1 To 1)
    VOLATILITY_VECTOR(1, 1) = VOLATILITY_RNG
End If
If UBound(VOLATILITY_VECTOR, 1) <> NROWS Then: GoTo ERROR_LABEL


'-------------------------------------------------------------------------------------
Select Case OUTPUT
'-------------------------------------------------------------------------------------
Case 0
'-------------------------------------------------------------------------------------
    NCOLUMNS = 22
    HEADINGS_STR = "SPOT,STRIKE,TIME,RATES,VOLATILITY,PARITY,VI: CALL PRICES," & _
    "VI: CALL INITIAL GUESS,VI: CALL IMPLVOL,VI: CALL VOLATILITY ERROR,VI: PUT PRICES," & _
    "VI: PUT IMPLVOL,VI: PUT VOLATILITY ERROR,VI: PUT PARITY CHECK,VI: PUT ERROR CHECK," & _
    "VI: CHOOSE IMPLVOL,VII: CALL IMPLVOL,VII: CALL PRICING ERROR,VII: PUT IMPLVOL," & _
    "VII: PUT PRICING ERROR,VII: PUT PARITY CHECK,VII: PUT ERROR CHECK,"
    GoSub LOAD_LINE
    For i = 1 To NROWS
        GoSub INPUTS_LINE
        'spot - strike*exp(-r*t)
        TEMP_MATRIX(i, 6) = TEMP_MATRIX(i, 1) - TEMP_MATRIX(i, 2) * Exp(-TEMP_MATRIX(i, 4) * TEMP_MATRIX(i, 3))
        'Compute prices  (and check parity for them), then retrieve volatility from them (and show errors against given vol)
        TEMP_MATRIX(i, 7) = BS_INTRINSIC_FUNC(TEMP_MATRIX(i, 1), TEMP_MATRIX(i, 2), TEMP_MATRIX(i, 3), TEMP_MATRIX(i, 4), TEMP_MATRIX(i, 5), 3)
        TEMP_MATRIX(i, 8) = BS_CALL_INITIAL_VOLATILITY_FUNC(TEMP_MATRIX(i, 1), TEMP_MATRIX(i, 2), TEMP_MATRIX(i, 3), TEMP_MATRIX(i, 4), TEMP_MATRIX(i, 7), False)
        TEMP_MATRIX(i, 9) = BS_VOLATILITY_FUNC(TEMP_MATRIX(i, 1), TEMP_MATRIX(i, 2), TEMP_MATRIX(i, 3), TEMP_MATRIX(i, 4), TEMP_MATRIX(i, 7), 0)
        
        'given vol - impl vol
        TEMP_MATRIX(i, 10) = TEMP_MATRIX(i, 5) - TEMP_MATRIX(i, 9)
        TEMP_MATRIX(i, 11) = BS_INTRINSIC_FUNC(TEMP_MATRIX(i, 1), TEMP_MATRIX(i, 2), TEMP_MATRIX(i, 3), TEMP_MATRIX(i, 4), TEMP_MATRIX(i, 5), 4)
        
        TEMP_MATRIX(i, 12) = BS_VOLATILITY_FUNC(TEMP_MATRIX(i, 1), TEMP_MATRIX(i, 2), TEMP_MATRIX(i, 3), TEMP_MATRIX(i, 4), TEMP_MATRIX(i, 11), 1)
        
        'given vol - impl vol
        TEMP_MATRIX(i, 13) = TEMP_MATRIX(i, 5) - TEMP_MATRIX(i, 12)
        'call - put
        TEMP_MATRIX(i, 14) = TEMP_MATRIX(i, 7) - TEMP_MATRIX(i, 11)
        'spot - strike*exp(-r*t) - (call -put)
        TEMP_MATRIX(i, 15) = TEMP_MATRIX(i, 6) - TEMP_MATRIX(i, 14)
        '  take the one, which is off the money (consisting of premium only)
        TEMP_MATRIX(i, 16) = IIf(TEMP_MATRIX(i, 1) <= TEMP_MATRIX(i, 2), TEMP_MATRIX(i, 9), TEMP_MATRIX(i, 12))
        
        'If having a pair of Call and Put combine their vol and see what happens (even in extreme cases)
        TEMP_MATRIX(i, 17) = BS_INTRINSIC_FUNC(TEMP_MATRIX(i, 1), TEMP_MATRIX(i, 2), TEMP_MATRIX(i, 3), TEMP_MATRIX(i, 4), TEMP_MATRIX(i, 16), 3)
        'Call(v) - Call(implVol)
        TEMP_MATRIX(i, 18) = TEMP_MATRIX(i, 7) - TEMP_MATRIX(i, 17)
        TEMP_MATRIX(i, 19) = BS_INTRINSIC_FUNC(TEMP_MATRIX(i, 1), TEMP_MATRIX(i, 2), TEMP_MATRIX(i, 3), TEMP_MATRIX(i, 4), TEMP_MATRIX(i, 16), 4)
        
        'Put(v) - Put(implVol)
        TEMP_MATRIX(i, 20) = TEMP_MATRIX(i, 11) - TEMP_MATRIX(i, 19)
        'call - put
        TEMP_MATRIX(i, 21) = TEMP_MATRIX(i, 17) - TEMP_MATRIX(i, 19)
        'spot - strike*exp(-r*t) - (call -put)
        TEMP_MATRIX(i, 22) = TEMP_MATRIX(i, 6) - TEMP_MATRIX(i, 21)
    Next i
'-------------------------------------------------------------------------------------
Case Else 'Increasing Volatility
'-------------------------------------------------------------------------------------
    NCOLUMNS = 21
    HEADINGS_STR = "SPOT,STRIKE,TIME,RATES,VOLATILITY,PARITY,VI: CALL PRICES," & _
    "VI: CALL IMPLVOL,VI: CALL VOLATILITY ERROR,VI: PUT PRICES,VI: PUT IMPLVOL," & _
    "VI: PUT VOLATILITY ERROR,VI: PUT PARITY CHECK,VI: PUT ERROR CHECK,VI: CHOOSE IMPLVOL," & _
    "VII: CALL IMPLVOL,VII: CALL PRICING ERROR,VII: PUT IMPLVOL,VII: PUT PRICING ERROR," & _
    "VII: PUT PARITY CHECK,VII: PUT ERROR CHECK,"
    GoSub LOAD_LINE
    For i = 1 To NROWS
        GoSub INPUTS_LINE
        TEMP_MATRIX(i, 6) = TEMP_MATRIX(i, 1) - TEMP_MATRIX(i, 2) * Exp(-TEMP_MATRIX(i, 4) * TEMP_MATRIX(i, 3))
        
        'Compute prices  (and check parity for them), then retrieve volatility from them (and show errors against given vol)
        TEMP_MATRIX(i, 7) = BS_INTRINSIC_FUNC(TEMP_MATRIX(i, 1), TEMP_MATRIX(i, 2), TEMP_MATRIX(i, 3), TEMP_MATRIX(i, 4), TEMP_MATRIX(i, 5), 3)
        TEMP_MATRIX(i, 8) = BS_VOLATILITY_FUNC(TEMP_MATRIX(i, 1), TEMP_MATRIX(i, 2), TEMP_MATRIX(i, 3), TEMP_MATRIX(i, 4), TEMP_MATRIX(i, 7), 2)
        
        'given vol - impl vol
        TEMP_MATRIX(i, 9) = TEMP_MATRIX(i, 5) - TEMP_MATRIX(i, 8)
        TEMP_MATRIX(i, 10) = BS_INTRINSIC_FUNC(TEMP_MATRIX(i, 1), TEMP_MATRIX(i, 2), TEMP_MATRIX(i, 3), TEMP_MATRIX(i, 4), TEMP_MATRIX(i, 5), 4)
        
        TEMP_MATRIX(i, 11) = BS_VOLATILITY_FUNC(TEMP_MATRIX(i, 1), TEMP_MATRIX(i, 2), TEMP_MATRIX(i, 3), TEMP_MATRIX(i, 4), TEMP_MATRIX(i, 10), 3)
        
        'given vol - impl vol
        TEMP_MATRIX(i, 12) = TEMP_MATRIX(i, 5) - TEMP_MATRIX(i, 11)
        TEMP_MATRIX(i, 13) = TEMP_MATRIX(i, 7) - TEMP_MATRIX(i, 10)
        'spot - strike*exp(-r*t) - (call -put)
        TEMP_MATRIX(i, 14) = TEMP_MATRIX(i, 6) - TEMP_MATRIX(i, 13)
        'If having a pair of Call and Put combine their vol and see what happens
        '(even in extreme cases)
        
        'Thus volatility can be increased (possibly a lot) and still the prices are
        'the same. Within this routine the vol however is not maximal for that (a
        'mess). Of course all that can not be better than the pricing function.
        If TEMP_MATRIX(i, 8) >= TEMP_MATRIX(i, 11) Then 'take the maximum
            TEMP_MATRIX(i, 15) = TEMP_MATRIX(i, 8)
        Else
            TEMP_MATRIX(i, 15) = TEMP_MATRIX(i, 11)
        End If
        TEMP_MATRIX(i, 16) = BS_INTRINSIC_FUNC(TEMP_MATRIX(i, 1), TEMP_MATRIX(i, 2), TEMP_MATRIX(i, 3), TEMP_MATRIX(i, 4), TEMP_MATRIX(i, 15), 3)
        
        'Call(v) - Call(implVol)
        TEMP_MATRIX(i, 17) = TEMP_MATRIX(i, 7) - TEMP_MATRIX(i, 16)
        
        
        TEMP_MATRIX(i, 18) = BS_INTRINSIC_FUNC(TEMP_MATRIX(i, 1), TEMP_MATRIX(i, 2), TEMP_MATRIX(i, 3), TEMP_MATRIX(i, 4), TEMP_MATRIX(i, 15), 4)
        'Put(v) - Put(implVol)
        TEMP_MATRIX(i, 19) = TEMP_MATRIX(i, 10) - TEMP_MATRIX(i, 18)
        'call - put
        TEMP_MATRIX(i, 20) = TEMP_MATRIX(i, 16) - TEMP_MATRIX(i, 18)
        'spot - strike*exp(-r*t) - (call -put)
        TEMP_MATRIX(i, 21) = TEMP_MATRIX(i, 6) - TEMP_MATRIX(i, 20)
    Next i
'-------------------------------------------------------------------------------------
End Select
'-------------------------------------------------------------------------------------

BS_VOLATILITY_INCREASED_FUNC = TEMP_MATRIX

'-------------------------------------------------------------------------------------
Exit Function
'-------------------------------------------------------------------------------------
INPUTS_LINE:
'-------------------------------------------------------------------------------------
    TEMP_MATRIX(i, 1) = SPOT_VECTOR(i, 1)
    TEMP_MATRIX(i, 2) = STRIKE_VECTOR(i, 1)
    TEMP_MATRIX(i, 3) = EXPIRATION_VECTOR(i, 1)
    TEMP_MATRIX(i, 4) = RATE_VECTOR(i, 1)
    TEMP_MATRIX(i, 5) = VOLATILITY_VECTOR(i, 1)
'-------------------------------------------------------------------------------------
Return
'-------------------------------------------------------------------------------------
LOAD_LINE:
'-------------------------------------------------------------------------------------
    ReDim TEMP_MATRIX(0 To NROWS, 1 To NCOLUMNS)
    i = 1
    For k = 1 To NCOLUMNS
        j = InStr(i, HEADINGS_STR, ",")
        TEMP_MATRIX(0, k) = Mid(HEADINGS_STR, i, j - i)
        i = j + 1
    Next k
'-------------------------------------------------------------------------------------
Return
'-------------------------------------------------------------------------------------
ERROR_LABEL:
BS_VOLATILITY_INCREASED_FUNC = Err.number
End Function


Function BS_INTRINSIC_FUNC(ByVal SPOT As Double, _
ByVal STRIKE As Double, _
ByVal EXPIRATION As Double, _
ByVal RATE As Double, _
Optional ByVal VOLATILITY As Double, _
Optional ByVal OUTPUT As Integer = 0) 'As Double

On Error GoTo ERROR_LABEL

'Parity --> ' = spot - strike*exp(-r*t)
'----------------------------------------------------------------------------
Select Case OUTPUT
'----------------------------------------------------------------------------
Case 0 'BSCall_intrinsic
'----------------------------------------------------------------------------
    If SPOT <= STRIKE * Exp(-RATE * EXPIRATION) Then
      BS_INTRINSIC_FUNC = 0#
    Else
      BS_INTRINSIC_FUNC = SPOT - STRIKE * Exp(-RATE * EXPIRATION)
    End If
'----------------------------------------------------------------------------
Case 1 'BSPut_intrinsic
'----------------------------------------------------------------------------
    If STRIKE * Exp(-RATE * EXPIRATION) <= SPOT Then
      BS_INTRINSIC_FUNC = 0#
    Else
      BS_INTRINSIC_FUNC = STRIKE * Exp(-RATE * EXPIRATION) - SPOT
    End If
'----------------------------------------------------------------------------
Case 2 'BS_Premium --> Call & Put
'----------------------------------------------------------------------------
    Dim ABSM_VAL As Double
    Dim ETA_VAL As Double
    Dim H_VAL As Double
    Dim SIGMA_VAL As Double
       
    SIGMA_VAL = VOLATILITY * Sqr(EXPIRATION)
    ABSM_VAL = Abs((RATE * EXPIRATION + Log(SPOT / STRIKE)))
    
    If (SIGMA_VAL < PUB_DBL_EPSILON) Then
      BS_INTRINSIC_FUNC = 0# ' no premium for vol ~ 0
      Exit Function
    End If
      
    ETA_VAL = ABSM_VAL / SIGMA_VAL
    H_VAL = Exp(ABSM_VAL / 2#)
    SIGMA_VAL = SIGMA_VAL / 2#
    
    BS_INTRINSIC_FUNC = Exp(-RATE * EXPIRATION / 2) * Sqr(STRIKE * SPOT) * _
                (CN_FUNC(ETA_VAL - SIGMA_VAL) / H_VAL - CN_FUNC(ETA_VAL + SIGMA_VAL) * H_VAL)

'----------------------------------------------------------------------------
Case 3 'BSCall
'----------------------------------------------------------------------------
    BS_INTRINSIC_FUNC = BS_INTRINSIC_FUNC(SPOT, STRIKE, EXPIRATION, RATE, 0, 0) + _
             BS_INTRINSIC_FUNC(SPOT, STRIKE, EXPIRATION, RATE, VOLATILITY, 2)
'----------------------------------------------------------------------------
Case Else 'BSPut
'----------------------------------------------------------------------------
    BS_INTRINSIC_FUNC = BS_INTRINSIC_FUNC(SPOT, STRIKE, EXPIRATION, RATE, 0, 1) + _
            BS_INTRINSIC_FUNC(SPOT, STRIKE, EXPIRATION, RATE, VOLATILITY, 2)
End Select

Exit Function
ERROR_LABEL:
BS_INTRINSIC_FUNC = Err.number
End Function

Function BS_VOLATILITY_FUNC(ByVal SPOT As Double, _
ByVal STRIKE As Double, _
ByVal EXPIRATION As Double, _
ByVal RATE As Double, _
ByVal PRICE As Double, _
Optional ByVal OUTPUT As Integer = 0) 'As Double

On Error GoTo ERROR_LABEL

'----------------------------------------------------------------------------
Select Case OUTPUT
'----------------------------------------------------------------------------
Case 0 'BSCallVola
'----------------------------------------------------------------------------
    BS_VOLATILITY_FUNC = BS_CALL_INCREASING_VOLATILITY_FUNC(SPOT, STRIKE, _
                         EXPIRATION, RATE, PRICE, False)
'----------------------------------------------------------------------------
Case 1 'BSPutVola
'----------------------------------------------------------------------------
    BS_VOLATILITY_FUNC = BS_VOLATILITY_FUNC(Exp(-RATE * EXPIRATION) * STRIKE, _
                         Exp(RATE * EXPIRATION) * SPOT, EXPIRATION, RATE, _
                         PRICE, 0)
'----------------------------------------------------------------------------
Case 2 'BSCallVola_increased
'----------------------------------------------------------------------------
    BS_VOLATILITY_FUNC = BS_CALL_INCREASING_VOLATILITY_FUNC(SPOT, STRIKE, _
                         EXPIRATION, RATE, PRICE, True)
'----------------------------------------------------------------------------
Case Else 'BSPutVola_increased
'----------------------------------------------------------------------------
    BS_VOLATILITY_FUNC = BS_VOLATILITY_FUNC(Exp(-RATE * EXPIRATION) * STRIKE, _
                         Exp(RATE * EXPIRATION) * SPOT, EXPIRATION, _
                         RATE, PRICE, 0)
'----------------------------------------------------------------------------
End Select
'----------------------------------------------------------------------------

Exit Function
ERROR_LABEL:
BS_VOLATILITY_FUNC = Err.number
End Function


' only to show the initial guess for the vol
Function BS_CALL_INITIAL_VOLATILITY_FUNC( _
ByVal SPOT As Double, _
ByVal STRIKE As Double, _
ByVal EXPIRATION As Double, _
ByVal RATE As Double, _
ByVal PRICE As Double, _
Optional ByVal INCREASE_VOLAT_FLAG As Integer = False) As Double
  
Dim MU_VAL As Double
Dim PREMIUM As Double
Dim SIGMA_VAL As Double
Dim INTRINSIC_VAL As Double
Dim INF_PRICE_VAL As Variant
Dim SUP_PRICE_VAL As Variant
Dim ADJUSTED_PRICE_VAL As Double

On Error GoTo ERROR_LABEL

INTRINSIC_VAL = BS_INTRINSIC_FUNC(SPOT, STRIKE, EXPIRATION, RATE, 0, 0)

' ensure, that PRICE is within admissible bounds and adjust if needed
INF_PRICE_VAL = INTRINSIC_VAL
SUP_PRICE_VAL = SPOT * (1# - 8# * PUB_DBL_EPSILON)

ADJUSTED_PRICE_VAL = PRICE
If (SUP_PRICE_VAL <= ADJUSTED_PRICE_VAL) Then ADJUSTED_PRICE_VAL = SUP_PRICE_VAL
If (ADJUSTED_PRICE_VAL <= INF_PRICE_VAL) Then ADJUSTED_PRICE_VAL = INF_PRICE_VAL

PREMIUM = ADJUSTED_PRICE_VAL - INTRINSIC_VAL

If (INCREASE_VOLAT_FLAG = True) Then
  'MsgBox ("premium1: " & premium)
  If (0# < INTRINSIC_VAL) Then PREMIUM = LINSOLVE_MAX_FUNC(INTRINSIC_VAL, ADJUSTED_PRICE_VAL)
  'MsgBox ("premium2: " & premium)
End If

PREMIUM = PREMIUM * Exp(RATE * EXPIRATION / 2) / Sqr(STRIKE * SPOT)
MU_VAL = (RATE * EXPIRATION + Log(SPOT / STRIKE)) / 2#
SIGMA_VAL = BS_INITIAL_SIGMA_FUNC(Abs(MU_VAL), PREMIUM)
'Debug.Print "MU_VAL = " & MU_VAL, "premium = " & premium, "BS_INITIAL_SIGMA_FUNC = " & SIGMA_VAL

BS_CALL_INITIAL_VOLATILITY_FUNC = 2# * SIGMA_VAL / Sqr(EXPIRATION)

Exit Function
ERROR_LABEL:
BS_CALL_INITIAL_VOLATILITY_FUNC = Err.number
End Function


Private Function BS_CALL_INCREASING_VOLATILITY_FUNC( _
ByVal SPOT As Double, _
ByVal STRIKE As Double, _
ByVal EXPIRATION As Double, _
ByVal RATE As Double, _
ByVal PRICE As Double, _
Optional ByVal INCREASE_VOLAT_FLAG As Boolean = False) As Double

Dim MU_VAL As Double
Dim PREMIUM As Double
Dim SIGMA_VAL As Double
Dim INTRINSIC_VAL As Double
Dim INF_PRICE_VAL As Variant
Dim SUP_PRICE_VAL As Variant
Dim ADJUSTED_PRICE_VAL As Double

On Error GoTo ERROR_LABEL

INTRINSIC_VAL = BS_INTRINSIC_FUNC(SPOT, STRIKE, EXPIRATION, RATE, 0, 0)

' ensure, that PRICE is within admissible bounds and adjust if needed
INF_PRICE_VAL = INTRINSIC_VAL
SUP_PRICE_VAL = SPOT * (1# - 8# * PUB_DBL_EPSILON)

ADJUSTED_PRICE_VAL = PRICE
If (SUP_PRICE_VAL <= ADJUSTED_PRICE_VAL) Then ADJUSTED_PRICE_VAL = SUP_PRICE_VAL
If (ADJUSTED_PRICE_VAL <= INF_PRICE_VAL) Then ADJUSTED_PRICE_VAL = INF_PRICE_VAL

PREMIUM = ADJUSTED_PRICE_VAL - INTRINSIC_VAL

If (INCREASE_VOLAT_FLAG = True) Then
  'MsgBox ("premium1: " & PREMIUM)
  If (0# < INTRINSIC_VAL) Then PREMIUM = LINSOLVE_MAX_FUNC(INTRINSIC_VAL, ADJUSTED_PRICE_VAL)
  'MsgBox ("premium2: " & PREMIUM)
End If

PREMIUM = PREMIUM * Exp(RATE * EXPIRATION / 2) / Sqr(STRIKE * SPOT)
MU_VAL = (RATE * EXPIRATION + Log(SPOT / STRIKE)) / 2#
SIGMA_VAL = BS_IMPLIED_VOLATILITY_PREMIUM_FUNC(Abs(MU_VAL), PREMIUM)

BS_CALL_INCREASING_VOLATILITY_FUNC = 2# * SIGMA_VAL / Sqr(EXPIRATION)

Exit Function
ERROR_LABEL:
BS_CALL_INCREASING_VOLATILITY_FUNC = Err.number
End Function

' solve P_VAL = NORMED_PREMIUM_FUNC(M_VAL,SPOT) for SPOT
Private Function BS_IMPLIED_VOLATILITY_PREMIUM_FUNC(ByVal MU_VAL As Double, _
ByVal PREMIUM As Double)

Dim M_VAL As Double
Dim N_VAL As Integer
Dim X_VAL As Double
Dim FX_VAL As Double
Dim DFX_VAL As Double
Dim XN_VAL As Double
Dim TERM_VAL As Double

On Error GoTo ERROR_LABEL

M_VAL = Abs(MU_VAL)
X_VAL = BS_INITIAL_SIGMA_FUNC(M_VAL, PREMIUM)
'MsgBox ("initial: " & X_VAL)

If (NORMED_PREMIUM_FUNC(M_VAL, X_VAL) = PREMIUM) Then
  BS_IMPLIED_VOLATILITY_PREMIUM_FUNC = X_VAL
  Exit Function
End If
'BS_IMPLIED_VOLATILITY_PREMIUM_FUNC = X_VAL ' test only

TERM_VAL = 1# * PUB_DBL_EPSILON
N_VAL = 0
For N_VAL = 0 To 8
  FX_VAL = NORMED_PREMIUM_FUNC(M_VAL, X_VAL) - PREMIUM
  If (Abs(FX_VAL) < TERM_VAL) Then
    BS_IMPLIED_VOLATILITY_PREMIUM_FUNC = X_VAL
    Exit Function
  End If
  
  DFX_VAL = VEGA_NORMED_PREMIUM_FUNC(M_VAL, X_VAL)
  If (Abs(DFX_VAL) < 16# * PUB_DBL_MIN) Then
    BS_IMPLIED_VOLATILITY_PREMIUM_FUNC = X_VAL
    Exit Function
  End If
  
  XN_VAL = X_VAL - FX_VAL / DFX_VAL ' Newton step
  If (Abs(1# - X_VAL / XN_VAL) <= PUB_DBL_EPSILON * 16#) Then Exit For
  
  X_VAL = XN_VAL
Next N_VAL
If 8 <= N_VAL Then
  'MsgBox ("Warning about convergence problem, steps used: " & N_VAL) ' test
  Debug.Print "Warning about convergence problem, steps used: " & N_VAL, "FX_VAL = " & FX_VAL, "X_VAL = " & X_VAL
End If

If X_VAL < 0# Then X_VAL = 0#

BS_IMPLIED_VOLATILITY_PREMIUM_FUNC = X_VAL

Exit Function
ERROR_LABEL:
BS_IMPLIED_VOLATILITY_PREMIUM_FUNC = Err.number
End Function


Private Function BS_INITIAL_SIGMA_FUNC(ByVal MU_VAL As Double, _
ByVal PREMIUM As Double) As Double

Dim M_VAL As Double
Dim V0_VAL As Double
Dim V1_VAL As Double
Dim P0_VAL As Double
Dim P1_VAL As Double
Dim TEMP_VAL As Double

On Error GoTo ERROR_LABEL

M_VAL = Abs(MU_VAL)
'MsgBox (M_VAL)
If (M_VAL <= 1E-20) Then M_VAL = 1E-20

If (Abs(PREMIUM) <= 1E-280) Then
  BS_INITIAL_SIGMA_FUNC = 2.70552394623868E-02 * M_VAL
  Exit Function
End If

TEMP_VAL = (M_VAL * PUB_PI_VAL * Sqr(2#) / PREMIUM) ^ (2# / 3#)
TEMP_VAL = TEMP_VAL / PUB_PI_VAL / 3#
V0_VAL = M_VAL / Sqr(3# * LAMBERT_W_FUNC(TEMP_VAL))
 
V1_VAL = Sqr(M_VAL) + Sqr(PUB_PI_VAL / 2#) * Exp(M_VAL) * (PREMIUM + _
     Exp(M_VAL) * CN_FUNC(2# * Sqr(M_VAL)) - Exp(-M_VAL) / 2#)

P0_VAL = NORMED_PREMIUM_FUNC(MU_VAL, V0_VAL)
P1_VAL = NORMED_PREMIUM_FUNC(MU_VAL, V1_VAL)

If (Abs(P0_VAL - PREMIUM) <= Abs(P1_VAL - PREMIUM)) Then
  BS_INITIAL_SIGMA_FUNC = V0_VAL
Else
  BS_INITIAL_SIGMA_FUNC = V1_VAL
End If

'Debug.Print "V0_VAL = " & V0_VAL, "V1_VAL = " & V1_VAL,
'"in LAMBERT_W_FUNC = " & TEMP_VAL / PUB_PI_VAL / 3#, "M_VAL = " & M_VAL

Exit Function
ERROR_LABEL:
BS_INITIAL_SIGMA_FUNC = Err.number
End Function



' Find P for A given linear equation
' Y = A + P with 0 <= A, 0 <= P and 0 < Y,
' where P numerical maximal (which only makes sense if A <> 0)
' this is A bit messy in Excel, it seems not to meet IEEE 754 ...

Private Function LINSOLVE_MAX_FUNC(ByVal AA_VAL As Double, _
ByVal YY_VAL As Double)

Dim j As Integer
Dim k As Integer

Dim A_VAL As Double
Dim Y_VAL As Double
Dim X_VAL As Double
Dim XI_VAL As Double
Dim P_VAL As Double
Dim P0_VAL As Double
Dim P1_VAL As Double
Dim H_VAL As Double

On Error GoTo ERROR_LABEL

A_VAL = AA_VAL
Y_VAL = YY_VAL
H_VAL = PUB_DBL_EPSILON / 2#

If ((Y_VAL <= 0#) Or (A_VAL <= 0#) Or (Y_VAL < A_VAL)) Then
  LINSOLVE_MAX_FUNC = YY_VAL - AA_VAL
  Exit Function
End If

XI_VAL = Log(Abs(Y_VAL)) * PUB_BY_LN2_VAL ' = log2(|y|)

' do not use cint: cint(1.5) = 2
If (0 <= XI_VAL) Then
  k = FLOOR_FUNC(XI_VAL) + 1
Else
  k = -(FLOOR_FUNC(-XI_VAL) + 1)
End If

X_VAL = Y_VAL / (2 ^ k) ' hence: Y_VAL = X_VAL*2^k, 0.5 <= X_VAL < 1

If Not (0.5 <= Abs(X_VAL) And Abs(X_VAL) < 1) Then
  ' something went wrong: display message or raise an error:
  'MsgBox ("false mantissa X_VAL = " & X_VAL & ", XI_VAL = " & XI_VAL)
  Debug.Print "false mantissa X_VAL = " & X_VAL & _
  ", y= " & Y_VAL & _
  ", XI_VAL = " & XI_VAL & _
  ", k = " & k
End If
'Y_VAL = YY_VAL * 2 ^ (-k) ' = X_VAL

Y_VAL = X_VAL
' hence care only for mantissa, i.e. 0.5 <= Y_VAL < 1, scale A_VAL for that

If (0 <= k And A_VAL <= 2# ^ (PUB_DBL_MIN_EXP + k)) Then
  LINSOLVE_MAX_FUNC = YY_VAL - AA_VAL
  Exit Function
End If

A_VAL = AA_VAL * CDbl(2 ^ (-k))

P0_VAL = Y_VAL - A_VAL

For j = 1 To 56 ' 52
  P_VAL = P0_VAL + H_VAL
  If (Y_VAL < A_VAL + P_VAL) Then
    Exit For
  Else
    P0_VAL = P_VAL
    H_VAL = 2# * H_VAL
  End If
  If (0.5 < H_VAL) Then Exit For
Next j
P_VAL = P0_VAL
'MsgBox ("P0_VAL = " & P0_VAL & ", j = " & j)

For j = 1 To 80
  H_VAL = H_VAL / 2#
  P1_VAL = P_VAL + H_VAL
  If (Y_VAL < CDbl(A_VAL + P1_VAL) Or CDbl(P1_VAL - P_VAL) = 0#) Then Exit For
    P_VAL = P1_VAL
Next j
'MsgBox ("P_VAL = " & P_VAL & ", j = " & j)
'MsgBox ("Y_VAL = " & Y_VAL & ", a+p = " & A_VAL + P_VAL & ", a+p+h = " & A_VAL + P_VAL + H_VAL)

If (Not (IsNumeric(P_VAL))) Then
  LINSOLVE_MAX_FUNC = YY_VAL - AA_VAL
  Exit Function
End If

LINSOLVE_MAX_FUNC = P_VAL * CDbl(2 ^ k)

Exit Function
ERROR_LABEL:
LINSOLVE_MAX_FUNC = Err.number
End Function




' LambertW: quick, but not too dirty, taken from DESY

Private Function LAMBERT_W_FUNC(ByVal X_VAL As Double) As Double
On Error GoTo ERROR_LABEL
If (376# < X_VAL) Then
  LAMBERT_W_FUNC = Log(X_VAL - 4#) - (1# - 1# / Log(X_VAL)) * Log(Log(X_VAL))
Else
  LAMBERT_W_FUNC = 0.665 * (1# + 0.0195 * Log(X_VAL + 1#)) * Log(X_VAL + 1#) + 0.04
End If
Exit Function
ERROR_LABEL:
LAMBERT_W_FUNC = Err.number
End Function

' premium for normed price (for Call and Put)
' = C(M_VAL,S_VAL) = premium for normed price (for Call and Put)
' = exp(-abs(M_VAL))*CN_FUNC(abs(M_VAL)/S_VAL - S_VAL) -
'exp(abs(M_VAL))*CN_FUNC(abs(M_VAL)/S_VAL + S_VAL)

Private Function NORMED_PREMIUM_FUNC(ByVal M_VAL As Double, _
ByVal S_VAL As Double)

Dim H_VAL As Double
Dim ETA_VAL As Double
Dim ABSM_VAL As Double

On Error GoTo ERROR_LABEL

If (S_VAL <= PUB_DBL_EPSILON) Then
  NORMED_PREMIUM_FUNC = 0#
  Exit Function
End If

ABSM_VAL = Abs(M_VAL)
ETA_VAL = ABSM_VAL / S_VAL
H_VAL = Exp(ABSM_VAL)

NORMED_PREMIUM_FUNC = CN_FUNC(ETA_VAL - S_VAL) / H_VAL - CN_FUNC(ETA_VAL + S_VAL) * H_VAL

Exit Function
ERROR_LABEL:
NORMED_PREMIUM_FUNC = Err.number
End Function

' = diff( C(M_VAL,S_VAL), S_VAL) = exp(-M_VAL^2/sigma^2/2 - sigma^2/2)*sqrt(2/Pi)
Private Function VEGA_NORMED_PREMIUM_FUNC(ByVal M_VAL As Double, _
ByVal S_VAL As Double)

Dim ETA_VAL As Double
Dim XI_VAL  As Double

On Error GoTo ERROR_LABEL

ETA_VAL = M_VAL / S_VAL
XI_VAL = (-ETA_VAL * ETA_VAL - S_VAL * S_VAL) / 2#
XI_VAL = XI_VAL - 0.225791352644727    ' ln( sqrt(2/Pi) )

VEGA_NORMED_PREMIUM_FUNC = Exp(XI_VAL)

Exit Function
ERROR_LABEL:
VEGA_NORMED_PREMIUM_FUNC = Err.number
End Function

'greatest integer less than or equal to A_VAL number
Private Function FLOOR_FUNC(ByVal X_VAL As Double) As Integer
On Error GoTo ERROR_LABEL
FLOOR_FUNC = CInt(X_VAL)
If X_VAL < FLOOR_FUNC Then FLOOR_FUNC = FLOOR_FUNC - 1
Exit Function
ERROR_LABEL:
FLOOR_FUNC = Err.number
End Function

Private Function CDFN_NEGATIVE_AXIS_FUNC(ByVal X_VAL As Double) As Double
'accurate to 1.e-15, according to J. Hart, cf. Graeme West

Dim ABSX_VAL As Double
Dim FRAC_VAL As Double

On Error GoTo ERROR_LABEL

ABSX_VAL = Abs(X_VAL)

If ABSX_VAL < 7.07106781186547 Then
  CDFN_NEGATIVE_AXIS_FUNC = Exp(-ABSX_VAL * ABSX_VAL / 2) * _
    ((((((3.52624965998911E-02 * ABSX_VAL _
          + 0.700383064443688) * ABSX_VAL _
        + 6.37396220353165) * ABSX_VAL _
        + 33.912866078383) * ABSX_VAL _
        + 112.079291497871) * ABSX_VAL _
        + 221.213596169931) * ABSX_VAL _
        + 220.206867912376) _
    / _
    (((((((8.83883476483184E-02 * ABSX_VAL _
        + 1.75566716318264) * ABSX_VAL _
        + 16.064177579207) * ABSX_VAL _
        + 86.7807322029461) * ABSX_VAL _
        + 296.564248779674) * ABSX_VAL _
        + 637.333633378831) * ABSX_VAL _
        + 793.826512519948) * ABSX_VAL _
        + 440.413735824752)

ElseIf 37# <= ABSX_VAL Then   ' cut off
  CDFN_NEGATIVE_AXIS_FUNC = 0#

Else                      ' asymptotic series
  FRAC_VAL = 4# / (ABSX_VAL + 0.65)
  FRAC_VAL = 3# / (ABSX_VAL + FRAC_VAL)
  FRAC_VAL = 2# / (ABSX_VAL + FRAC_VAL)
  FRAC_VAL = 1# / (ABSX_VAL + FRAC_VAL)
  CDFN_NEGATIVE_AXIS_FUNC = Exp(-ABSX_VAL * ABSX_VAL * 0.5) * INV_SQRT_2PI_VAL / (ABSX_VAL + FRAC_VAL)
End If

'If 0# < X_VAL Then CDFN_FUNC = 1# - CDFN_FUNC

Exit Function
ERROR_LABEL:
CDFN_NEGATIVE_AXIS_FUNC = Err.number
End Function


' cumulative normal distribution
Private Function CDFN_FUNC(X_VAL As Double) As Double

On Error GoTo ERROR_LABEL

CDFN_FUNC = CDFN_NEGATIVE_AXIS_FUNC(Abs(X_VAL))
If 0# < X_VAL Then CDFN_FUNC = 1# - CDFN_FUNC

Exit Function
ERROR_LABEL:
CDFN_FUNC = Err.number
End Function

' complementary cumulative normal distribution, CN_FUNC(X_VAL) = 1 - CDFN_FUNC(X_VAL)
Private Function CN_FUNC(X_VAL As Double) As Double

On Error GoTo ERROR_LABEL

CN_FUNC = CDFN_NEGATIVE_AXIS_FUNC(Abs(X_VAL))
If X_VAL < 0# Then CN_FUNC = 1# - CN_FUNC

Exit Function
ERROR_LABEL:
CN_FUNC = Err.number
End Function


Private Sub TEST_TIMING_FUNC()

Dim i As Long
Dim nLOOPS As Long

Dim K_VAL As Double
Dim ST_VAL As Double
Dim PRICE As Double

Dim ERROR_VAL As Double
Dim MAX_ERROR_VAL As Double
Dim RESULT_VAL As Double
Dim STEP_SIZE_VAL As Double

nLOOPS = 10000

'-----------------------------------------------------------------------------------
'tst_call_timing
'-----------------------------------------------------------------------------------
STEP_SIZE_VAL = (160 - 60) * 1# / nLOOPS ' from 60 to 160 in nLOOPS steps

K_VAL = 60#
ST_VAL = Timer
For i = 0 To nLOOPS
  K_VAL = K_VAL + STEP_SIZE_VAL
  RESULT_VAL = BS_INTRINSIC_FUNC(100#, K_VAL, 1, 0.03, 0.25, 3)
  ' activate to check some results
  'If 0 = i Mod 1000 Then
  '  Debug.Print i, K_VAL, RESULT_VAL
  'End If
Next i
Debug.Print ("------------------")
Debug.Print "1) tst_call_timing"
Debug.Print "number evaluations: " & nLOOPS
Debug.Print "msecs needed: " & Format((Timer - ST_VAL) * 1000, "#.#")

'-----------------------------------------------------------------------------------
'tst_vola_timing
'-----------------------------------------------------------------------------------
STEP_SIZE_VAL = (100 - 60) * 1# / nLOOPS ' from 60 to 100 in nLOOPS steps

K_VAL = 60#
ST_VAL = Timer
For i = 0 To nLOOPS
  K_VAL = K_VAL + STEP_SIZE_VAL
  PRICE = 3# + (100# - K_VAL)
  RESULT_VAL = BS_VOLATILITY_FUNC(100#, K_VAL, 1#, 0.03, PRICE, 0)
  ' activate to check some results
  'If 0 = i Mod 1000 Then
  '  Debug.Print i, K_VAL, RESULT_VAL
  'End If
Next i

Debug.Print ("------------------")
Debug.Print "2) tst_vola_timing"
Debug.Print "number evaluations: " & nLOOPS
Debug.Print "msecs needed: " & Format((Timer - ST_VAL) * 1000, "#.#")

'-----------------------------------------------------------------------------------
'tst_vola_increased_timing
'-----------------------------------------------------------------------------------
STEP_SIZE_VAL = (100 - 60) * 1# / nLOOPS ' from 60 to 100 in nLOOPS steps

K_VAL = 60#
ST_VAL = Timer
For i = 0 To nLOOPS
  K_VAL = K_VAL + STEP_SIZE_VAL
  PRICE = 3# + (100# - K_VAL)
  RESULT_VAL = BS_VOLATILITY_FUNC(100#, K_VAL, 1#, 0.03, PRICE, 0)
  ' activate to check some results
  'If 0 = i Mod 1000 Then
  '  Debug.Print i, K_VAL, RESULT_VAL
  'End If
Next i

Debug.Print ("------------------")
Debug.Print "3) tst_vola_increased_timing"
Debug.Print "number evaluations: " & nLOOPS
Debug.Print "msecs needed: " & Format((Timer - ST_VAL) * 1000, "#.#")

'-----------------------------------------------------------------------------------
'tst_vola_exactness
'-----------------------------------------------------------------------------------
STEP_SIZE_VAL = (100 - 60) * 1# / nLOOPS ' from 60 to 100 in nLOOPS steps

K_VAL = 60#
ST_VAL = Timer
For i = 0 To nLOOPS
  K_VAL = K_VAL + STEP_SIZE_VAL
  PRICE = 3# + (100# - K_VAL)
  RESULT_VAL = BS_VOLATILITY_FUNC(100#, K_VAL, 1#, 0.03, PRICE, 0)
  RESULT_VAL = BS_INTRINSIC_FUNC(100#, K_VAL, 1, 0.03, RESULT_VAL, 3)
  ERROR_VAL = Abs(PRICE - RESULT_VAL)
  If MAX_ERROR_VAL <= ERROR_VAL Then MAX_ERROR_VAL = ERROR_VAL
Next i

Debug.Print ("------------------")
Debug.Print "4) tst_vola_exactness"
Debug.Print "number evaluations: " & nLOOPS
Debug.Print "maximal error: " & MAX_ERROR_VAL

'-----------------------------------------------------------------------------------

End Sub



