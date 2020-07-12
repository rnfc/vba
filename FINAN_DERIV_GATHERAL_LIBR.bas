Attribute VB_Name = "FINAN_DERIV_GATHERAL_LIBR"


'-------------------------------------------------------------------------------------
Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.
'-------------------------------------------------------------------------------------

Private PUB_FORWARD_VAL As Double
Private PUB_EXPIRATION_VAL As Double


'************************************************************************************
'************************************************************************************
'FUNCTION      : GATHERAL_SVI_FUNC
'DESCRIPTION   : Ensure that globals are set before calling this function
' only this makes it param_a function of STRIKE (and parameters),
' while FORWARD and time are constants

' The following is the LVM method (provide gradient etc)
' using matrix manipulation and solving linear equations

' volatility function depending on strike only, p=vector of parameters

'LIBRARY       : DERIVATIVES
'GROUP         : GATHERAL
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/31/2009
'************************************************************************************
'************************************************************************************

Function GATHERAL_SVI_FUNC(ByVal FORWARD As Double, _
ByVal EXPIRATION As Double, _
ByRef STRIKE_RNG As Variant, _
ByRef SIGMA_RNG As Variant, _
ByRef PARAM_RNG As Variant, _
Optional ByVal OUTPUT As Integer = 1)

'EXPIRATOIN --> Maturity Per period
'REFERENCE: http://www.math.nyu.edu/fellows_fin_math/gatheral/

Dim i As Long
Dim NROWS As Long
Dim TEMP_SUM As Double

Dim SIGMA_VECTOR As Variant
Dim STRIKE_VECTOR As Variant

Dim TEMP_MATRIX As Variant
Dim APARAM_VECTOR As Variant
Dim BPARAM_VECTOR As Variant

Dim MONEY_VECTOR As Variant 'This vector represents a measure of
'the degree to which a derivative is likely to have positive monetary
'value at its expiration, in the risk-neutral measure. It can be
'measured in percentage probability, or
'in standard deviations.
Dim VARIANCE_VECTOR As Variant

On Error GoTo ERROR_LABEL

STRIKE_VECTOR = STRIKE_RNG
If UBound(STRIKE_VECTOR, 1) = 1 Then
    STRIKE_VECTOR = MATRIX_TRANSPOSE_FUNC(STRIKE_VECTOR)
End If

SIGMA_VECTOR = SIGMA_RNG
If UBound(SIGMA_VECTOR, 1) = 1 Then
    SIGMA_VECTOR = MATRIX_TRANSPOSE_FUNC(SIGMA_VECTOR)
End If

If UBound(STRIKE_VECTOR, 1) <> UBound(SIGMA_VECTOR, 1) Then: GoTo ERROR_LABEL

NROWS = UBound(STRIKE_VECTOR, 1)

ReDim MONEY_VECTOR(1 To NROWS, 1 To 1)
ReDim VARIANCE_VECTOR(1 To NROWS, 1 To 1)

For i = 1 To NROWS
    MONEY_VECTOR(i, 1) = Log(STRIKE_VECTOR(i, 1) / FORWARD)
    VARIANCE_VECTOR(i, 1) = SIGMA_VECTOR(i, 1) ^ 2 * EXPIRATION
Next i

APARAM_VECTOR = PARAM_RNG
If UBound(APARAM_VECTOR, 1) = 1 Then
    APARAM_VECTOR = MATRIX_TRANSPOSE_FUNC(APARAM_VECTOR)
End If
BPARAM_VECTOR = GATHERAL_PARAMETERS_APPROXIMATION_FUNC(FORWARD, EXPIRATION, _
                STRIKE_VECTOR, SIGMA_VECTOR)
    
ReDim TEMP_MATRIX(0 To NROWS, 1 To 8)

TEMP_MATRIX(0, 1) = "STRIKE"
TEMP_MATRIX(0, 2) = "SIGMA"
TEMP_MATRIX(0, 3) = "MONEYNESS"
TEMP_MATRIX(0, 4) = "VARIANCE"
TEMP_MATRIX(0, 5) = "VAR_LS"
TEMP_MATRIX(0, 6) = "SIGMA_LS"
TEMP_MATRIX(0, 7) = "ERRORS"
TEMP_MATRIX(0, 8) = "APPROX_GUESS"

For i = 1 To NROWS

    TEMP_MATRIX(i, 1) = STRIKE_VECTOR(i, 1)
    TEMP_MATRIX(i, 2) = SIGMA_VECTOR(i, 1)
    TEMP_MATRIX(i, 3) = MONEY_VECTOR(i, 1)
    TEMP_MATRIX(i, 4) = VARIANCE_VECTOR(i, 1)
        
    TEMP_MATRIX(i, 5) = GATHERAL_VARIANCE_FUNC(MONEY_VECTOR(i, 1), _
            APARAM_VECTOR(1, 1), APARAM_VECTOR(2, 1), _
            APARAM_VECTOR(3, 1), APARAM_VECTOR(4, 1), _
            APARAM_VECTOR(5, 1))

    TEMP_MATRIX(i, 6) = GATHERAL_VOLATILITY_FUNC(STRIKE_VECTOR(i, 1), _
                FORWARD, EXPIRATION, APARAM_VECTOR(1, 1), _
                APARAM_VECTOR(2, 1), APARAM_VECTOR(3, 1), _
                APARAM_VECTOR(4, 1), APARAM_VECTOR(5, 1))

    TEMP_MATRIX(i, 7) = TEMP_MATRIX(i, 2) - TEMP_MATRIX(i, 6)
    TEMP_MATRIX(i, 8) = GATHERAL_VOLATILITY_FUNC(STRIKE_VECTOR(i, 1), _
                FORWARD, EXPIRATION, BPARAM_VECTOR(1, 1), BPARAM_VECTOR(2, 1), _
                BPARAM_VECTOR(3, 1), BPARAM_VECTOR(4, 1), BPARAM_VECTOR(5, 1))
        
    TEMP_SUM = TEMP_SUM + ((TEMP_MATRIX(i, 2) - TEMP_MATRIX(i, 6)) * _
                               (TEMP_MATRIX(i, 2) - TEMP_MATRIX(i, 6)))
Next i
        
Select Case OUTPUT
Case 0
    GATHERAL_SVI_FUNC = TEMP_MATRIX
Case 1
    GATHERAL_SVI_FUNC = APARAM_VECTOR
Case 2
    GATHERAL_SVI_FUNC = BPARAM_VECTOR
Case Else
    GATHERAL_SVI_FUNC = TEMP_SUM ^ 0.5
End Select

Exit Function
ERROR_LABEL:
GATHERAL_SVI_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : GATHERAL_PARAMETERS_SOLVER_FUNC
'DESCRIPTION   : This is a least-square fitting through LVM interface for
'an objective functions f: IR^1 x IR^n -> IR^1 with n parameters. This is
'fitting of 1-dim curves against data by estimating parameters.
'LIBRARY       : DERIVATIVES
'GROUP         : GATHERAL
'ID            : 002
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/31/2009
'************************************************************************************
'************************************************************************************

Function GATHERAL_PARAMETERS_SOLVER_FUNC(ByVal FORWARD As Double, _
ByVal EXPIRATION As Double, _
ByRef STRIKE_RNG As Variant, _
ByRef SIGMA_RNG As Variant, _
Optional ByVal nLOOPS As Long = 1000, _
Optional ByVal tolerance As Double = 10 ^ -15)

Dim SIGMA_VECTOR As Variant
Dim STRIKE_VECTOR As Variant
Dim PARAMETERS_VECTOR As Variant

On Error GoTo ERROR_LABEL

PUB_FORWARD_VAL = FORWARD
PUB_EXPIRATION_VAL = EXPIRATION

STRIKE_VECTOR = STRIKE_RNG
    If UBound(STRIKE_VECTOR, 1) = 1 Then: _
        STRIKE_VECTOR = MATRIX_TRANSPOSE_FUNC(STRIKE_VECTOR)

SIGMA_VECTOR = SIGMA_RNG
    If UBound(SIGMA_VECTOR, 1) = 1 Then: _
        SIGMA_VECTOR = MATRIX_TRANSPOSE_FUNC(SIGMA_VECTOR)

If UBound(STRIKE_VECTOR, 1) <> UBound(SIGMA_VECTOR, 1) Then: _
GoTo ERROR_LABEL

PARAMETERS_VECTOR = GATHERAL_PARAMETERS_APPROXIMATION_FUNC(FORWARD, EXPIRATION, _
               STRIKE_VECTOR, SIGMA_VECTOR)


GATHERAL_PARAMETERS_SOLVER_FUNC = LEVENBERG_MARQUARDT_OPTIMIZATION3_FUNC(STRIKE_VECTOR, SIGMA_VECTOR, _
                    PARAMETERS_VECTOR, "GATHERAL_OBJECTIVE_FUNC", _
                    "GATHERAL_JACOBI_FUNC", nLOOPS, 10 ^ 3, 10 ^ 9, tolerance)
  
Exit Function
ERROR_LABEL:
GATHERAL_PARAMETERS_SOLVER_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : GATHERAL_OBJECTIVE_FUNC
'DESCRIPTION   : Gatheral Objective Function
'LIBRARY       : DERIVATIVES
'GROUP         : GATHERAL
'ID            : 003
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/31/2009
'************************************************************************************
'************************************************************************************

Function GATHERAL_OBJECTIVE_FUNC(ByRef STRIKE_RNG As Variant, _
ByRef PARAM_RNG As Variant)

Dim i As Long
Dim NROWS As Long

Dim ATEMP_VAL As Double
Dim BTEMP_VAL As Double
Dim CTEMP_VAL As Double
Dim DTEMP_VAL As Double
Dim TEMP_STRIKE As Double

Dim TEMP_VECTOR As Variant
Dim STRIKE_VECTOR As Variant
Dim PARAMETERS_VECTOR As Variant

On Error GoTo ERROR_LABEL

PARAMETERS_VECTOR = PARAM_RNG
If UBound(PARAMETERS_VECTOR, 1) = 1 Then
    PARAMETERS_VECTOR = MATRIX_TRANSPOSE_FUNC(PARAMETERS_VECTOR)
End If

STRIKE_VECTOR = STRIKE_RNG
If UBound(STRIKE_VECTOR, 1) = 1 Then
    STRIKE_VECTOR = MATRIX_TRANSPOSE_FUNC(STRIKE_VECTOR)
End If

NROWS = UBound(STRIKE_VECTOR, 1)

ReDim TEMP_VECTOR(1 To NROWS, 1 To 1)

'PUB_FORWARD_VAL = 5066.5
'PUB_EXPIRATION_VAL = (77 - 7 / 24) / 365

For i = 1 To NROWS

    TEMP_STRIKE = STRIKE_VECTOR(i, 1)

    ATEMP_VAL = Log(TEMP_STRIKE / PUB_FORWARD_VAL)
    BTEMP_VAL = ATEMP_VAL - PARAMETERS_VECTOR(5, 1)

    CTEMP_VAL = PARAMETERS_VECTOR(1, 1) + PARAMETERS_VECTOR(2, 1) * _
            (PARAMETERS_VECTOR(4, 1) * BTEMP_VAL + (BTEMP_VAL * _
            BTEMP_VAL + PARAMETERS_VECTOR(3, 1) * PARAMETERS_VECTOR(3, 1)) ^ 0.5)

    DTEMP_VAL = (Abs(CTEMP_VAL / PUB_EXPIRATION_VAL)) ^ 0.5

    TEMP_VECTOR(i, 1) = DTEMP_VAL

Next i

GATHERAL_OBJECTIVE_FUNC = TEMP_VECTOR

Exit Function
ERROR_LABEL:
GATHERAL_OBJECTIVE_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : GATHERAL_GRADIENT_FUNC
'DESCRIPTION   : Gatheral Gradient Function
'LIBRARY       : DERIVATIVES
'GROUP         : GATHERAL
'ID            : 004
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/31/2009
'************************************************************************************
'************************************************************************************

Function GATHERAL_GRADIENT_FUNC(ByVal FORWARD As Double, _
ByVal EXPIRATION As Double, _
ByRef STRIKE_RNG As Variant, _
ByRef PARAM_RNG As Variant)

Dim i As Long

Dim NROWS As Long
Dim NSIZE As Long

Dim A_VAL As Double
Dim B_VAL As Double
Dim SIGMA As Double
Dim RHO As Double
Dim M_VAL As Double

Dim TEMP_STRIKE As Double

Dim ATEMP_VAL As Double
Dim BTEMP_VAL As Double
Dim CTEMP_VAL As Double
Dim DTEMP_VAL As Double
Dim ETEMP_VAL As Double

Dim TEMP_VECTOR As Variant
Dim TEMP_MATRIX As Variant
Dim STRIKE_VECTOR As Variant
Dim PARAMETERS_VECTOR As Variant

On Error GoTo ERROR_LABEL

STRIKE_VECTOR = STRIKE_RNG
If UBound(STRIKE_VECTOR, 1) = 1 Then
    STRIKE_VECTOR = MATRIX_TRANSPOSE_FUNC(STRIKE_VECTOR)
End If

PARAMETERS_VECTOR = PARAM_RNG
If UBound(PARAMETERS_VECTOR, 1) = 1 Then
    PARAMETERS_VECTOR = MATRIX_TRANSPOSE_FUNC(PARAMETERS_VECTOR)
End If

NSIZE = UBound(PARAMETERS_VECTOR, 1)
NROWS = UBound(STRIKE_VECTOR, 1)

ReDim TEMP_MATRIX(1 To NROWS, 1 To NSIZE)

A_VAL = PARAMETERS_VECTOR(1, 1)
B_VAL = PARAMETERS_VECTOR(2, 1)
SIGMA = PARAMETERS_VECTOR(3, 1)
RHO = PARAMETERS_VECTOR(4, 1)
M_VAL = PARAMETERS_VECTOR(5, 1)

PUB_FORWARD_VAL = FORWARD '5066.5
PUB_EXPIRATION_VAL = EXPIRATION '(77 - 7 / 24) / 365

TEMP_VECTOR = GATHERAL_OBJECTIVE_FUNC(STRIKE_VECTOR, PARAMETERS_VECTOR)

For i = 1 To NROWS

    TEMP_STRIKE = STRIKE_VECTOR(i, 1)

    ATEMP_VAL = Log(TEMP_STRIKE / PUB_FORWARD_VAL)
    BTEMP_VAL = 1 / (2 * TEMP_VECTOR(i, 1) * PUB_EXPIRATION_VAL)

    CTEMP_VAL = ATEMP_VAL - M_VAL

    DTEMP_VAL = GATHERAL_VARIANCE_FUNC(ATEMP_VAL, 0, 1, SIGMA, RHO, M_VAL)
    ETEMP_VAL = GATHERAL_VARIANCE_FUNC(ATEMP_VAL, 0, 1, SIGMA, 0, M_VAL)

    TEMP_MATRIX(i, 1) = BTEMP_VAL
    TEMP_MATRIX(i, 1) = TEMP_MATRIX(i, 1)
    
    TEMP_MATRIX(i, 2) = BTEMP_VAL * ETEMP_VAL
    TEMP_MATRIX(i, 2) = TEMP_MATRIX(i, 2)
    
    TEMP_MATRIX(i, 3) = BTEMP_VAL * B_VAL * SIGMA / DTEMP_VAL
    TEMP_MATRIX(i, 3) = TEMP_MATRIX(i, 3)
    
    TEMP_MATRIX(i, 4) = BTEMP_VAL * B_VAL * CTEMP_VAL
    TEMP_MATRIX(i, 4) = TEMP_MATRIX(i, 4)
    
    TEMP_MATRIX(i, 5) = -BTEMP_VAL * B_VAL * (RHO + _
    CTEMP_VAL / DTEMP_VAL)
    TEMP_MATRIX(i, 5) = TEMP_MATRIX(i, 5)
Next i

GATHERAL_GRADIENT_FUNC = TEMP_MATRIX
  
Exit Function
ERROR_LABEL:
GATHERAL_GRADIENT_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : GATHERAL_JACOBI_FUNC
'DESCRIPTION   : Gatheral Jacobi Function
'LIBRARY       : DERIVATIVES
'GROUP         : GATHERAL
'ID            : 005
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/31/2009
'************************************************************************************
'************************************************************************************

Function GATHERAL_JACOBI_FUNC(ByRef STRIKE_RNG As Variant, _
ByRef PARAM_RNG As Variant)

Dim STRIKE_VECTOR As Variant
Dim PARAMETERS_VECTOR As Variant

Dim tolerance As Double

On Error GoTo ERROR_LABEL

tolerance = 0.001
STRIKE_VECTOR = STRIKE_RNG
If UBound(STRIKE_VECTOR, 1) = 1 Then
    STRIKE_VECTOR = MATRIX_TRANSPOSE_FUNC(STRIKE_VECTOR)
End If

PARAMETERS_VECTOR = PARAM_RNG
If UBound(PARAMETERS_VECTOR, 1) = 1 Then
    PARAMETERS_VECTOR = MATRIX_TRANSPOSE_FUNC(PARAMETERS_VECTOR)
End If

GATHERAL_JACOBI_FUNC = JACOBI_MATRIX_FUNC("GATHERAL_OBJECTIVE_FUNC", _
                       STRIKE_VECTOR, PARAMETERS_VECTOR, tolerance)
Exit Function
ERROR_LABEL:
GATHERAL_JACOBI_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : GATHERAL_VOLATILITY_FUNC
'DESCRIPTION   : Volatility function, parameters given individually
'LIBRARY       : DERIVATIVES
'GROUP         : GATHERAL
'ID            : 006
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/31/2009
'************************************************************************************
'************************************************************************************

Function GATHERAL_VOLATILITY_FUNC(ByVal STRIKE As Double, _
ByVal FORWARD As Double, _
ByVal EXPIRATION As Double, _
ByVal A_VAL As Double, _
ByVal B_VAL As Double, _
ByVal SIGMA As Double, _
ByVal RHO As Double, _
ByVal M_VAL As Double)

Dim TEMP_VAL As Double

On Error GoTo ERROR_LABEL

TEMP_VAL = Log(STRIKE / FORWARD)

GATHERAL_VOLATILITY_FUNC = (Abs(GATHERAL_VARIANCE_FUNC(TEMP_VAL, A_VAL, _
                B_VAL, SIGMA, RHO, M_VAL) / EXPIRATION)) ^ 0.5
Exit Function
ERROR_LABEL:
GATHERAL_VOLATILITY_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : GATHERAL_VOLATILITY_PARAMETERS_FUNC
'DESCRIPTION   : Volatility function, using a vector of model parameters
'LIBRARY       : DERIVATIVES
'GROUP         : GATHERAL
'ID            : 007
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/31/2009
'************************************************************************************
'************************************************************************************

Function GATHERAL_VOLATILITY_PARAMETERS_FUNC(ByVal STRIKE As Double, _
ByVal FORWARD As Double, _
ByVal EXPIRATION As Double, _
ByRef PARAM_RNG As Variant)

Dim A_VAL As Double
Dim B_VAL As Double
Dim SIGMA As Double
Dim RHO As Double
Dim M_VAL As Double

Dim PARAMETERS_VECTOR As Variant

On Error GoTo ERROR_LABEL

PARAMETERS_VECTOR = PARAM_RNG
If UBound(PARAMETERS_VECTOR, 1) = 1 Then
    PARAMETERS_VECTOR = MATRIX_TRANSPOSE_FUNC(PARAMETERS_VECTOR)
End If

A_VAL = PARAMETERS_VECTOR(1, 1)
B_VAL = PARAMETERS_VECTOR(2, 1)
SIGMA = PARAMETERS_VECTOR(3, 1)
RHO = PARAMETERS_VECTOR(4, 1)
M_VAL = PARAMETERS_VECTOR(5, 1)

GATHERAL_VOLATILITY_PARAMETERS_FUNC = GATHERAL_VOLATILITY_FUNC(STRIKE, FORWARD, EXPIRATION, _
                           A_VAL, B_VAL, SIGMA, RHO, M_VAL)
  
Exit Function
ERROR_LABEL:
GATHERAL_VOLATILITY_PARAMETERS_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : GATHERAL_VARIANCE_FUNC
'DESCRIPTION   : Variance function from the paper, parameters given individually
'LIBRARY       : DERIVATIVES
'GROUP         : GATHERAL
'ID            : 008
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/31/2009
'************************************************************************************
'************************************************************************************


Function GATHERAL_VARIANCE_FUNC(ByVal MONEYNESS As Double, _
ByVal A_VAL As Double, _
ByVal B_VAL As Double, _
ByVal SIGMA As Double, _
ByVal RHO As Double, _
ByVal M_VAL As Double)

Dim TEMP_DELTA As Double

On Error GoTo ERROR_LABEL

TEMP_DELTA = MONEYNESS - M_VAL

GATHERAL_VARIANCE_FUNC = A_VAL + B_VAL * (RHO * TEMP_DELTA + _
                (TEMP_DELTA * TEMP_DELTA + SIGMA * SIGMA) ^ 0.5)
  
  
Exit Function
ERROR_LABEL:
GATHERAL_VARIANCE_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : GATHERAL_PARAMETERS_APPROXIMATION_FUNC
'DESCRIPTION   : Estimate initial parameters (and refine them by Levenberg-Marquardt)
' for a volatility smile given through Gatheral's variance model
'LIBRARY       : DERIVATIVES
'GROUP         : GATHERAL
'ID            : 009
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/31/2009
'************************************************************************************
'************************************************************************************

Function GATHERAL_PARAMETERS_APPROXIMATION_FUNC(ByVal FORWARD As Double, _
ByVal EXPIRATION As Double, _
ByRef STRIKE_RNG As Variant, _
ByRef SIGMA_RNG As Variant)

Dim i As Long
Dim NROWS As Long

Dim X1_VAL As Double
Dim X2_VAL As Double

Dim Y1_VAL As Double
Dim Y2_VAL As Double

Dim MIN_VAR As Double

Dim A_INIT_VAL As Double
Dim B_INIT_VAL As Double
Dim M_INIT_VAL As Double
Dim R_INIT_VAL As Double
Dim S_INIT_VAL As Double

Dim ATEMP_VAL As Double
Dim BTEMP_VAL As Double
Dim CTEMP_VAL As Double
Dim DTEMP_VAL As Double

Dim MONEY_VECTOR As Variant
Dim VARIANCE_VECTOR As Variant
Dim PARAMETERS_VECTOR As Variant

Dim SIGMA_VECTOR As Variant
Dim STRIKE_VECTOR As Variant

On Error GoTo ERROR_LABEL

SIGMA_VECTOR = SIGMA_RNG
If UBound(SIGMA_VECTOR, 1) = 1 Then
    SIGMA_VECTOR = MATRIX_TRANSPOSE_FUNC(SIGMA_VECTOR)
End If

STRIKE_VECTOR = STRIKE_RNG
If UBound(STRIKE_VECTOR, 1) = 1 Then
    STRIKE_VECTOR = MATRIX_TRANSPOSE_FUNC(STRIKE_VECTOR)
End If

If UBound(SIGMA_VECTOR, 1) <> UBound(STRIKE_VECTOR, 1) Then: GoTo ERROR_LABEL

NROWS = UBound(STRIKE_VECTOR, 1)

ReDim MONEY_VECTOR(1 To NROWS, 1 To 1)
ReDim VARIANCE_VECTOR(1 To NROWS, 1 To 1)

For i = 1 To NROWS ' transform to variance over moneyness
  MONEY_VECTOR(i, 1) = Log(STRIKE_VECTOR(i, 1) / FORWARD)
  VARIANCE_VECTOR(i, 1) = SIGMA_VECTOR(i, 1) ^ 2 * EXPIRATION
Next i

'---------------------------------------------------------------------------------

ReDim PARAMETERS_VECTOR(1 To 5, 1 To 1)

' coefficients left asymptotics
X1_VAL = MONEY_VECTOR(1, 1)
X2_VAL = MONEY_VECTOR(2, 1)
Y1_VAL = VARIANCE_VECTOR(1, 1)
Y2_VAL = VARIANCE_VECTOR(2, 1)

ATEMP_VAL = (X1_VAL * Y2_VAL - Y1_VAL * X2_VAL) / (X1_VAL - X2_VAL)  ' Line(0)
BTEMP_VAL = (Y2_VAL - Y1_VAL) / (X2_VAL - X1_VAL)            ' Line(1) - Line(0)
BTEMP_VAL = -Abs(BTEMP_VAL)                      ' enforce descending line

' coefficients right asymptotics
X1_VAL = MONEY_VECTOR(NROWS - 1, 1)
X2_VAL = MONEY_VECTOR(NROWS, 1)
Y1_VAL = VARIANCE_VECTOR(NROWS - 1, 1)
Y2_VAL = VARIANCE_VECTOR(NROWS, 1)

CTEMP_VAL = (X1_VAL * Y2_VAL - Y1_VAL * X2_VAL) / (X1_VAL - X2_VAL)  ' Line(0)
DTEMP_VAL = (Y2_VAL - Y1_VAL) / (X2_VAL - X1_VAL)            ' Line(1) - Line(0)
DTEMP_VAL = Abs(DTEMP_VAL)                       ' enforce ascending line

B_INIT_VAL = -1 / 2 * BTEMP_VAL + 1 / 2 * DTEMP_VAL ' 5 through asymptotics
R_INIT_VAL = (BTEMP_VAL + DTEMP_VAL) / (-BTEMP_VAL + DTEMP_VAL)

A_INIT_VAL = ATEMP_VAL + BTEMP_VAL * (-ATEMP_VAL + CTEMP_VAL) / (BTEMP_VAL - DTEMP_VAL)
M_INIT_VAL = BTEMP_VAL * (-ATEMP_VAL + CTEMP_VAL) / (BTEMP_VAL - DTEMP_VAL) / B_INIT_VAL / _
(R_INIT_VAL - 1)

MIN_VAR = VARIANCE_VECTOR(NROWS, 1) ' sigma = smoothing the vertex at _
the minimum
For i = NROWS - 1 To 1 Step -1
    If VARIANCE_VECTOR(i, 1) < MIN_VAR Then: MIN_VAR = VARIANCE_VECTOR(i, 1)
Next i

S_INIT_VAL = -(-MIN_VAR + ATEMP_VAL + BTEMP_VAL * (-ATEMP_VAL + CTEMP_VAL) / _
(BTEMP_VAL - DTEMP_VAL))
S_INIT_VAL = S_INIT_VAL / B_INIT_VAL / Sqr(Abs(1 - R_INIT_VAL ^ 2))

'-----------------------------------------------------------------------
PARAMETERS_VECTOR(1, 1) = A_INIT_VAL
PARAMETERS_VECTOR(2, 1) = B_INIT_VAL
PARAMETERS_VECTOR(3, 1) = S_INIT_VAL
PARAMETERS_VECTOR(4, 1) = R_INIT_VAL
PARAMETERS_VECTOR(5, 1) = M_INIT_VAL
'-----------------------------------------------------------------------

GATHERAL_PARAMETERS_APPROXIMATION_FUNC = PARAMETERS_VECTOR

Exit Function
ERROR_LABEL:
GATHERAL_PARAMETERS_APPROXIMATION_FUNC = Err.number
End Function
