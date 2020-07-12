Attribute VB_Name = "FINAN_DERIV_BS_FINITE_LIBR"


Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.


'--------------------------------------------------------------
' use globals as the integrator accepts no parameters
Private PUB_FORWARD_VAL As Double   ' Exp(RATE_VAL * time) * SPOT_VAL
Private PUB_MU_VAL As Double     ' Log(STRIKE_VAL / PUB_FORWARD_VAL)
Private PUB_SIGMA_VAL_VAL As Double  ' SIGMA_VAL * Sqr(time)
'--------------------------------------------------------------

'************************************************************************************
'************************************************************************************
'FUNCTION      : PDF_CALL_OPTION_FUNC
'DESCRIPTION   : BS pricing through integrating the pay off against the
'risk neutral density, both over spots and
'logarithmic moneyness (i.e. Breeden-Litzenberger).

'In finance, moneyness is a measure of the degree to which
'a derivative is likely to have positive monetary value at
'its expiration, in the risk-neutral measure. It can be
'measured in percentage probability, or in standard deviations.

'USEFUL REFERENCE: http://www.math.nyu.edu/research/carrp/

'LIBRARY       : DERIVATIVES
'GROUP         : BS_FINITE
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA

'************************************************************************************
'************************************************************************************


Function PDF_CALL_OPTION_FUNC(ByVal SPOT_VAL As Double, _
ByVal STRIKE_VAL As Double, _
ByVal EXPIRATION As Double, _
ByVal RATE_VAL As Double, _
ByVal SIGMA_VAL As Double)

Dim TEMP_VAL As Double
Dim FUNC_STR As String

Dim nLOOPS As Long
Dim epsilon As Double

On Error GoTo ERROR_LABEL

nLOOPS = 400
epsilon = 0.000000000000001

FUNC_STR = "PDF_CALL_OPTION_INTEGRAND_FUNC"

'MU --> RATE_VAL

PUB_FORWARD_VAL = Exp(RATE_VAL * EXPIRATION) * SPOT_VAL
PUB_MU_VAL = Log(STRIKE_VAL / PUB_FORWARD_VAL)
PUB_SIGMA_VAL_VAL = SIGMA_VAL * Sqr(EXPIRATION)

TEMP_VAL = PUB_SIGMA_VAL_VAL * PUB_SIGMA_VAL_VAL / 2 + 9 * PUB_SIGMA_VAL_VAL
' REMEMBER: TEMP_VAL (cut off) as infinite integral
' gives an overflow int(f, a, oo)
'
If PUB_MU_VAL < 0 Then ' if needed take care for a peak
    PDF_CALL_OPTION_FUNC = SPOT_VAL * (GAUSS_KRONROD_INTEGRATION_FUNC(FUNC_STR, PUB_MU_VAL, -PUB_MU_VAL, nLOOPS, epsilon) + _
                                       GAUSS_KRONROD_INTEGRATION_FUNC(FUNC_STR, -PUB_MU_VAL, TEMP_VAL, nLOOPS, epsilon))
ElseIf 0 <= PUB_MU_VAL Then
    PDF_CALL_OPTION_FUNC = SPOT_VAL * GAUSS_KRONROD_INTEGRATION_FUNC(FUNC_STR, PUB_MU_VAL, TEMP_VAL, nLOOPS, epsilon)
Else
    PDF_CALL_OPTION_FUNC = 0
End If ' or split at the peak
'Here i use adaptive Gauss-Kronrod point rules for finite or
'semi-infinite intervalls. That is 'known' to be a good working
'tool in Financial Mathematics.
Exit Function
ERROR_LABEL:
PDF_CALL_OPTION_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : PDF_CALL_OPTION_INTEGRAND_FUNC
'DESCRIPTION   : PDF Call Integral
'LIBRARY       : DERIVATIVES
'GROUP         : BS_FINITE
'ID            : 002
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA

'************************************************************************************
'************************************************************************************

Private Function PDF_CALL_OPTION_INTEGRAND_FUNC(ByVal MU_VAL As Double)

Dim PI_VAL As Double

On Error GoTo ERROR_LABEL
PI_VAL = 3.14159265358979

PDF_CALL_OPTION_INTEGRAND_FUNC = (Exp(MU_VAL) - Exp(PUB_MU_VAL)) * Exp(-0.5 * ((MU_VAL * (1 / PUB_SIGMA_VAL_VAL) + 0.5 * PUB_SIGMA_VAL_VAL)) ^ 2) * (1 / PUB_SIGMA_VAL_VAL) * (1 / (2 * PI_VAL) ^ 0.5) '--> ' the pdf for Black-Scholes: exp(-1/2*(mu/sigma+sigma/2)^2)/sigma/sqrt(2*PI_VAL)

Exit Function
ERROR_LABEL:
PDF_CALL_OPTION_INTEGRAND_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : EXPLICIT_FINITE_DIFFERENCE_CALL_OPTION_FUNC
'DESCRIPTION   : Explicit finite difference method as applied to the
'Black-Scholes PDE for Call Options
'LIBRARY       : DERIVATIVES
'GROUP         : BS_FINITE
'ID            : 003
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA

'************************************************************************************
'************************************************************************************

Function EXPLICIT_FINITE_DIFFERENCE_CALL_OPTION_FUNC(ByVal OUTPUT As Integer, _
ByVal SPOT_VAL As Double, _
ByVal STRIKE_VAL As Double, _
ByVal EXPIRATION_VAL As Double, _
ByVal RATE_VAL As Double, _
ByVal SIGMA_VAL As Double, _
Optional ByVal DT_VAL As Double = 0.04, _
Optional ByVal SPOT0_VAL As Double = 0, _
Optional ByVal DSPOT_VAL As Double = 5, _
Optional ByVal NSPOT_VAL As Long = 10, _
Optional ByVal CND_TYPE As Integer = 0, _
Optional ByVal tolerance As Double = 0.000001, _
Optional ByVal epsilon As Double = 0.000000000000001)

Dim i As Long
Dim j As Long

Dim SROW As Long
Dim NROWS As Long

Dim SCOLUMN As Long
Dim NCOLUMNS As Long

Dim TEMP_MATRIX() As Variant
Dim TEMP_VECTOR() As Variant

On Error GoTo ERROR_LABEL

NROWS = Int((EXPIRATION_VAL - 0) / DT_VAL) + 1
NCOLUMNS = Int(NSPOT_VAL)

ReDim TEMP_VECTOR(1 To 4, 1 To NCOLUMNS)

For j = 1 To NCOLUMNS
    TEMP_VECTOR(1, j) = 0.5 * DT_VAL * (-RATE_VAL * (j - 1) + (SIGMA_VAL * (j - 1)) ^ 2) / (1 + RATE_VAL * DT_VAL)
    TEMP_VECTOR(2, j) = (1 - (SIGMA_VAL * (j - 1)) ^ 2 * DT_VAL) / (1 + RATE_VAL * DT_VAL)
    TEMP_VECTOR(3, j) = 0.5 * DT_VAL * (RATE_VAL * (j - 1) + (SIGMA_VAL * (j - 1)) ^ 2) / (1 + RATE_VAL * DT_VAL)
    TEMP_VECTOR(4, j) = TEMP_VECTOR(1, j) + TEMP_VECTOR(2, j) + TEMP_VECTOR(3, j)
Next j

ReDim TEMP_MATRIX(0 To NROWS, 0 To NCOLUMNS)
TEMP_MATRIX(0, 0) = "TN / Sp"
        
If SPOT0_VAL < tolerance Then
    SCOLUMN = (NCOLUMNS - 1) * (SPOT_VAL / DSPOT_VAL)
Else
    SCOLUMN = SPOT0_VAL + NCOLUMNS * (SPOT_VAL / DSPOT_VAL)
End If
        
For j = NCOLUMNS To 1 Step -1
    TEMP_MATRIX(0, j) = SCOLUMN
    SCOLUMN = SCOLUMN - (SPOT_VAL / DSPOT_VAL)
    TEMP_MATRIX(NROWS, j) = MAXIMUM_FUNC(TEMP_MATRIX(0, j) - STRIKE_VAL, 0)
Next j
                
SROW = EXPIRATION_VAL
For i = 1 To NROWS
    TEMP_MATRIX(i, 0) = SROW
    If i < NROWS Then
        TEMP_MATRIX(i, 1) = epsilon
        TEMP_MATRIX(i, NCOLUMNS) = TEMP_MATRIX(0, NCOLUMNS) - Exp(-RATE_VAL * (TEMP_MATRIX(i, 0))) * STRIKE_VAL
    End If
    SROW = SROW - DT_VAL
Next i

Select Case OUTPUT
'-----------------------------------------------------------------------------------------
Case 0
'-----------------------------------------------------------------------------------------
    EXPLICIT_FINITE_DIFFERENCE_CALL_OPTION_FUNC = TEMP_VECTOR
'-----------------------------------------------------------------------------------------
Case 1 'Explicit Finite Difference
'-----------------------------------------------------------------------------------------
    For i = (NROWS - 1) To 1 Step -1
        For j = (NCOLUMNS - 1) To 2 Step -1
            TEMP_MATRIX(i, j) = TEMP_MATRIX(i + 1, j - 1) * TEMP_VECTOR(1, j) + TEMP_MATRIX(i + 1, j) * TEMP_VECTOR(2, j) + TEMP_MATRIX(i + 1, j + 1) * TEMP_VECTOR(3, j)
        Next j
    Next i
    EXPLICIT_FINITE_DIFFERENCE_CALL_OPTION_FUNC = TEMP_MATRIX
'-----------------------------------------------------------------------------------------
Case Else 'Black-Scholes PDE
'-----------------------------------------------------------------------------------------
    For i = (NROWS) To 1 Step -1
        For j = (NCOLUMNS) To 1 Step -1
            If TEMP_MATRIX(0, j) < tolerance Then: TEMP_MATRIX(0, j) = epsilon
            TEMP_MATRIX(i, j) = PDF_CALL_PRICE_BS_FUNC(STRIKE_VAL, TEMP_MATRIX(i, 0), RATE_VAL, FINITE_DIFFERENCE_ALPHA_VAL_FUNC(TEMP_MATRIX(0, j), RATE_VAL, SIGMA_VAL, TEMP_MATRIX(i, 0)), FINITE_DIFFERENCE_BETA_VAL_FUNC(SIGMA_VAL, TEMP_MATRIX(i, 0)), CND_TYPE)
        Next j
    Next i
    EXPLICIT_FINITE_DIFFERENCE_CALL_OPTION_FUNC = TEMP_MATRIX
'-----------------------------------------------------------------------------------------
End Select
'-----------------------------------------------------------------------------------------

Exit Function
ERROR_LABEL:
EXPLICIT_FINITE_DIFFERENCE_CALL_OPTION_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : PDF_CALL_PRICE_BS_FUNC
'DESCRIPTION   : BS Call Price PDF Function
'LIBRARY       : DERIVATIVES
'GROUP         : BS_FINITE
'ID            : 004
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA

'************************************************************************************
'************************************************************************************

Function PDF_CALL_PRICE_BS_FUNC(ByVal STRIKE_VAL As Double, _
ByVal EXPIRATION As Double, _
ByVal RATE_VAL As Double, _
ByVal ALPHA_VAL As Double, _
ByVal BETA_VAL As Double, _
Optional ByVal CND_TYPE As Integer = 0)

Dim D1_VAL As Double
Dim D2_VAL As Double
Dim TEMP_VAL As Double

On Error GoTo ERROR_LABEL

D1_VAL = (-1 * Log(STRIKE_VAL) / Log(Exp(1)) + ALPHA_VAL + BETA_VAL ^ 2) / BETA_VAL
D2_VAL = D1_VAL - BETA_VAL
TEMP_VAL = Exp(ALPHA_VAL + 0.5 * BETA_VAL ^ 2) * CND_FUNC(D1_VAL, CND_TYPE) - STRIKE_VAL * CND_FUNC(D2_VAL, CND_TYPE)
PDF_CALL_PRICE_BS_FUNC = TEMP_VAL * Exp(-RATE_VAL * EXPIRATION)

Exit Function
ERROR_LABEL:
PDF_CALL_PRICE_BS_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : IMPLICIT_FINITE_DIFFERENCE_PUT_OPTION_FUNC

'DESCRIPTION   : This calculates price for European/ American Put
'Options using implicit finite difference method as described
'in Hull's book "Options, Futures and other Derivatives" in
'section 18.8

'LIBRARY       : DERIVATIVES
'GROUP         : BS_FINITE
'ID            : 005
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA

'************************************************************************************
'************************************************************************************

Function IMPLICIT_FINITE_DIFFERENCE_PUT_OPTION_FUNC(ByVal SPOT_VAL As Double, _
ByVal STRIKE_VAL As Double, _
ByVal RATE_VAL As Double, _
ByVal SIGMA_VAL As Double, _
ByVal EXPIRATION_VAL As Double, _
Optional ByVal DSPOT_VAL As Long = 20, _
Optional ByVal nSTEPS As Long = 10, _
Optional ByVal EXERCISE_TYPE As Integer = 0)
'--------------------------------------------------------------------------
'--------------------------------------------------------------------------
'SPOT_VAL: Spot Price
'STRIKE_VAL: Strike
'RATE_VAL: Risk free rate
'SIGMA_VAL: Volatility
'EXPIRATION_VAL: Time to Maturity
'DSPOT_VAL: No. of steps in Stock Price
'nSTEPS: No. of Time steps
'EXERCISE_TYPE: Exercise Type (0 for American,  else for European)
'--------------------------------------------------------------------------
'--------------------------------------------------------------------------

Dim i As Long '
Dim j As Long
Dim k As Long
Dim l As Long

Dim ii As Long
Dim jj As Long
Dim kk As Long

Dim A1_ARR As Variant
Dim A2_ARR As Variant
Dim A3_ARR As Variant

Dim A4_ARR As Variant
Dim A5_ARR As Variant
Dim A6_ARR As Variant

Dim A7_ARR As Variant
Dim A8_ARR As Variant
Dim A9_ARR As Variant
Dim A10_ARR As Variant

Dim A11_ARR As Variant
Dim A12_ARR As Variant
Dim A13_ARR As Variant

Dim A14_ARR As Variant
Dim A15_ARR As Variant
Dim A16_ARR As Variant

Dim TEMP_MATRIX As Variant

Dim MULT_VAL As Double
Dim DT_VAL As Double

On Error GoTo ERROR_LABEL

DT_VAL = EXPIRATION_VAL / nSTEPS 'time step size
MULT_VAL = (SPOT_VAL * 2) / DSPOT_VAL 'stock price step size
l = DSPOT_VAL + 1
ReDim A11_ARR(1 To l) 'vector to store option price grid
For i = 1 To l
  A11_ARR(i) = 0
Next i

ReDim A16_ARR(1 To l)
For i = 0 To DSPOT_VAL
  A11_ARR(i + 1) = MAXIMUM_FUNC(STRIKE_VAL - i * MULT_VAL, 0)
  A16_ARR(i + 1) = i * MULT_VAL
Next i

A12_ARR = A11_ARR

ReDim A1_ARR(1 To DSPOT_VAL - 1)
ReDim A2_ARR(1 To DSPOT_VAL - 1)
ReDim A3_ARR(1 To DSPOT_VAL - 1)
ReDim A13_ARR(1 To DSPOT_VAL - 1)

For j = 1 To DSPOT_VAL - 1
  A1_ARR(j) = 0.5 * RATE_VAL * j * DT_VAL - 0.5 * SIGMA_VAL * SIGMA_VAL * j * j * DT_VAL
  A2_ARR(j) = 1 + SIGMA_VAL * SIGMA_VAL * j * j * DT_VAL + RATE_VAL * DT_VAL
  A3_ARR(j) = -0.5 * RATE_VAL * j * DT_VAL - 0.5 * SIGMA_VAL * SIGMA_VAL * j * j * DT_VAL
  A13_ARR(j) = STRIKE_VAL - j * MULT_VAL
Next j

ReDim A11_ARR(1 To l)

'Apply boundary conditions for Put
A11_ARR(1) = STRIKE_VAL
A11_ARR(l) = 0

'matrix of coffecients for set of equations in 18.29
ReDim TEMP_MATRIX(1 To DSPOT_VAL - 1, 1 To DSPOT_VAL - 1)
TEMP_MATRIX(1, 1) = A2_ARR(1)
TEMP_MATRIX(1, 2) = A3_ARR(1)

TEMP_MATRIX(DSPOT_VAL - 1, DSPOT_VAL - 2) = A1_ARR(DSPOT_VAL - 1)
TEMP_MATRIX(DSPOT_VAL - 1, DSPOT_VAL - 1) = A2_ARR(DSPOT_VAL - 1)

For j = 2 To DSPOT_VAL - 2
    TEMP_MATRIX(j, j - 1) = A1_ARR(j)
    TEMP_MATRIX(j, j) = A2_ARR(j)
    TEMP_MATRIX(j, j + 1) = A3_ARR(j)
Next j

For i = 1 To nSTEPS 'Start rollback of time steps
    ReDim A14_ARR(1 To DSPOT_VAL - 1)
    A14_ARR(1) = A12_ARR(2) - A1_ARR(1) * A11_ARR(1)
    A14_ARR(DSPOT_VAL - 1) = A12_ARR(DSPOT_VAL) - A3_ARR(DSPOT_VAL - 1) * A11_ARR(l)
    For j = 2 To DSPOT_VAL - 2
        A14_ARR(j) = A12_ARR(j + 1)
    Next j
'-------------------------------------------------------------------------------
'Solve Tridiagonal System
'-------------------------------------------------------------------------------
    jj = UBound(TEMP_MATRIX, 1)
    ii = LBound(TEMP_MATRIX, 1)
    kk = jj - ii + 1

'Extract the vectors A4_ARR,A5_ARR,A6_ARR of the tridiagonal matrix

    ReDim A4_ARR(1 To kk)
    ReDim A5_ARR(1 To kk)
    ReDim A6_ARR(1 To kk)
    ReDim A10_ARR(1 To kk)
    
    A4_ARR(1) = 0
    A6_ARR(kk) = 0
    For j = 2 To kk
        A4_ARR(j) = TEMP_MATRIX(ii + j - 1, ii + j - 2)
    Next j
    For j = 1 To kk
        A5_ARR(j) = TEMP_MATRIX(ii + j - 1, ii + j - 1)
    Next j
    For j = 1 To kk - 1
        A6_ARR(j) = TEMP_MATRIX(ii + j - 1, ii + j)
    Next j

    A10_ARR = A14_ARR
    
    ReDim A7_ARR(1 To kk)
    ReDim A8_ARR(1 To kk)
    
    A7_ARR(1) = A5_ARR(1)
    A8_ARR(1) = A10_ARR(1)
    
    For k = 2 To kk
        MULT_VAL = A4_ARR(k) / A7_ARR(k - 1)
        A7_ARR(k) = A5_ARR(k) - MULT_VAL * A6_ARR(k - 1)
        A8_ARR(k) = A10_ARR(k) - MULT_VAL * A8_ARR(k - 1)
    Next k
    
    ReDim A9_ARR(1 To kk)
    A9_ARR(kk) = A8_ARR(kk) / A7_ARR(kk)
    For k = kk - 1 To 1 Step -1
        A9_ARR(k) = (A8_ARR(k) - A6_ARR(k) * A9_ARR(k + 1)) / A7_ARR(k)
    Next k
    A15_ARR = A9_ARR
    For j = 2 To DSPOT_VAL
        Select Case EXERCISE_TYPE
        Case 0 ', "a", "amer"
            A12_ARR(j) = MAXIMUM_FUNC(A15_ARR(j - 1), A13_ARR(j - 1))
        Case Else 'European
            A12_ARR(j) = A15_ARR(j - 1)
        End Select
    Next j
    A12_ARR(1) = A11_ARR(1)
    A12_ARR(l) = A11_ARR(l)
Next i

'-----------------------------------------------------------------------------
'-----------------------------------------------------------------------------
'Returns an interpolated value of A9_ARR doing A4_ARR lookup of xarr->yarr
'-----------------------------------------------------------------------------
'-----------------------------------------------------------------------------

If ((SPOT_VAL < A16_ARR(LBound(A16_ARR))) Or (SPOT_VAL > A16_ARR(UBound(A16_ARR)))) Then
    IMPLICIT_FINITE_DIFFERENCE_PUT_OPTION_FUNC = "Interp: SPOT_VAL is out of bound"
    Exit Function
End If
If A16_ARR(LBound(A16_ARR)) = SPOT_VAL Then
    IMPLICIT_FINITE_DIFFERENCE_PUT_OPTION_FUNC = A12_ARR(LBound(A12_ARR))
    Exit Function
End If
For i = LBound(A16_ARR) To UBound(A16_ARR)
    If A16_ARR(i) >= SPOT_VAL Then
        IMPLICIT_FINITE_DIFFERENCE_PUT_OPTION_FUNC = A12_ARR(i - 1) + (SPOT_VAL - A16_ARR(i - 1)) / (A16_ARR(i) - A16_ARR(i - 1)) * (A12_ARR(i) - A12_ARR(i - 1))
        Exit Function
    End If
Next i

Exit Function
ERROR_LABEL:
IMPLICIT_FINITE_DIFFERENCE_PUT_OPTION_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : FINITE_DIFFERENCE_ALPHA_VAL_FUNC
'DESCRIPTION   : FD Alpha Function
'LIBRARY       : DERIVATIVES
'GROUP         : BS_FINITE
'ID            : 006
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA

'************************************************************************************
'************************************************************************************

Function FINITE_DIFFERENCE_ALPHA_VAL_FUNC(ByVal SPOT_VAL As Double, _
ByVal RATE_VAL As Double, _
ByVal SIGMA_VAL As Double, _
ByVal EXPIRATION As Double)
On Error GoTo ERROR_LABEL
FINITE_DIFFERENCE_ALPHA_VAL_FUNC = Log(SPOT_VAL) / Log(Exp(1)) + (RATE_VAL - 0.5 * SIGMA_VAL ^ 2) * EXPIRATION
Exit Function
ERROR_LABEL:
FINITE_DIFFERENCE_ALPHA_VAL_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : FINITE_DIFFERENCE_BETA_VAL_FUNC
'DESCRIPTION   : FD Beta Function
'LIBRARY       : DERIVATIVES
'GROUP         : BS_FINITE
'ID            : 007
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA

'************************************************************************************
'************************************************************************************

Function FINITE_DIFFERENCE_BETA_VAL_FUNC(ByVal SIGMA_VAL As Double, _
ByVal EXPIRATION As Double)
On Error GoTo ERROR_LABEL
FINITE_DIFFERENCE_BETA_VAL_FUNC = SIGMA_VAL * EXPIRATION ^ 0.5
Exit Function
ERROR_LABEL:
FINITE_DIFFERENCE_BETA_VAL_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : FINITE_DIFFERENCE_LOG_RETURNS_DENSITY_FUNC
'DESCRIPTION   : Log Returns Density Function
'LIBRARY       : DERIVATIVES
'GROUP         : BS_FINITE
'ID            : 008
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA

'************************************************************************************
'************************************************************************************

Function FINITE_DIFFERENCE_LOG_RETURNS_DENSITY_FUNC(ByVal LOG_RETURN As Double, _
ByVal ALPHA_VAL As Double, _
ByVal BETA_VAL As Double, _
ByVal SPOT_VAL As Double)

Dim PI_VAL As Double
Dim TEMP_VAL As Double
On Error GoTo ERROR_LABEL
PI_VAL = 3.14159265358979
TEMP_VAL = (LOG_RETURN - ALPHA_VAL + (Log(SPOT_VAL) / Log(Exp(1)))) ^ 2 / (2 * BETA_VAL ^ 2)
FINITE_DIFFERENCE_LOG_RETURNS_DENSITY_FUNC = 1 / (BETA_VAL * (2 * PI_VAL) ^ 0.5) * Exp(-TEMP_VAL)
Exit Function
ERROR_LABEL:
FINITE_DIFFERENCE_LOG_RETURNS_DENSITY_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : FINITE_DIFFERENCE_LOG_NORMAL_DENSITY_FUNC
'DESCRIPTION   : Log Normal Density Function
'LIBRARY       : DERIVATIVES
'GROUP         : BS_FINITE
'ID            : 009
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA

'************************************************************************************
'************************************************************************************

Function FINITE_DIFFERENCE_LOG_NORMAL_DENSITY_FUNC(ByVal ALPHA_VAL As Double, _
ByVal BETA_VAL As Double, _
ByVal SPOT_VAL As Double)
Dim PI_VAL As Double
Dim TEMP_VAL As Double
On Error GoTo ERROR_LABEL
PI_VAL = 3.14159265358979
TEMP_VAL = (-(Log(SPOT_VAL) / Log(Exp(1)) - ALPHA_VAL) ^ 2) / (2 * BETA_VAL ^ 2)
FINITE_DIFFERENCE_LOG_NORMAL_DENSITY_FUNC = 1 / (SPOT_VAL * BETA_VAL * (2 * PI_VAL) ^ 0.5) * Exp(TEMP_VAL)
Exit Function
ERROR_LABEL:
FINITE_DIFFERENCE_LOG_NORMAL_DENSITY_FUNC = Err.number
End Function
