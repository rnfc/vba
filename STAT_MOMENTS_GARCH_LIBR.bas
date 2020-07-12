Attribute VB_Name = "STAT_MOMENTS_GARCH_LIBR"

'----------------------------------------------------------------------------------
'----------------------------------------------------------------------------------
Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.
'----------------------------------------------------------------------------------
'----------------------------------------------------------------------------------
Private PUB_FACTOR_VAL As Integer
Private PUB_GARCH_DATA As Variant
Private PUB_GARCH_OBJ As Integer
Private Const PUB_EPSILON As Double = 2 ^ 52
'----------------------------------------------------------------------------------
'----------------------------------------------------------------------------------


'************************************************************************************
'************************************************************************************
'FUNCTION      : GARCH_MODEL1_FUNC
'DESCRIPTION   : Generalized Autoregressive Conditional Heteroscedasticity Function
'LIBRARY       : STATISTICS
'GROUP         : GARCH
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/08/2008
'************************************************************************************
'************************************************************************************

Function GARCH_MODEL1_FUNC(ByRef DATE_RNG As Variant, _
ByRef DATA_RNG As Variant, _
ByVal OMEGA_VAL As Double, _
ByVal ALPHA_VAL As Double, _
ByVal BETA_VAL As Double, _
Optional ByVal COUNT_BASIS As Long = 254, _
Optional ByVal OUTPUT As Integer = 2)

Dim i As Long
Dim j As Long
Dim NROWS As Long

Dim TEMP_SUM As Double
Dim DELTA_VAL As Double
Dim MEAN_VAL As Double
Dim SIGMA_VAL As Double
Dim VAR_VAL As Double
Dim KURT_VAL As Double
Dim FACTOR_VAL As Double

Dim DATE_VECTOR As Variant
Dim DATA_VECTOR As Variant
Dim TEMP_MATRIX As Variant

Dim PI_VAL As Double

On Error GoTo ERROR_LABEL

'One important departure from independency over time is when returns follow
'GARCH(1,1) processes: r(t) = e(t) * sig(t)^2, with e(t) = N(0,1) and
'sig(t)^2 = w + a*r(t-1)^2 + b*sig(t-1)^2, 0<w<inf, a>=0, b>=0, a+b<1 and
'sig(0)^2=w/(1-a-b). The key feature of a GARCH(1,1) process is time-varying
'conditional volatility when a and/or b are not equal to zero. There are many
'financial time series which can be described very well with GARCH(1,1) models.

'Note that volatility is underestimated if the return time series is subject
'to positive autocorrelation. Simple adjustments for autocorrelation exist.

PI_VAL = 3.14159265358979

DATE_VECTOR = DATE_RNG
If UBound(DATE_VECTOR, 1) = 1 Then: DATE_VECTOR = MATRIX_TRANSPOSE_FUNC(DATE_VECTOR)
DATA_VECTOR = DATA_RNG
If UBound(DATA_VECTOR, 1) = 1 Then: DATA_VECTOR = MATRIX_TRANSPOSE_FUNC(DATA_VECTOR)

If UBound(DATE_VECTOR, 1) <> UBound(DATA_VECTOR, 1) Then: GoTo ERROR_LABEL
NROWS = UBound(DATA_VECTOR, 1)

'----------------------------------------------------------------------------------
TEMP_SUM = 0
For i = 2 To NROWS
    DELTA_VAL = Log(DATA_VECTOR(i, 1) / DATA_VECTOR(i - 1, 1))
    TEMP_SUM = TEMP_SUM + DELTA_VAL
Next i
MEAN_VAL = TEMP_SUM / (NROWS - 1)

TEMP_SUM = 0
For i = 2 To NROWS
    DELTA_VAL = Log(DATA_VECTOR(i, 1) / DATA_VECTOR(i - 1, 1))
    TEMP_SUM = TEMP_SUM + (DELTA_VAL - MEAN_VAL) ^ 2
Next i
SIGMA_VAL = (TEMP_SUM / (NROWS - 2)) ^ 0.5
VAR_VAL = SIGMA_VAL ^ 2
'----------------------------------------------------------------------------------

ReDim TEMP_MATRIX(0 To NROWS, 1 To 10)
'---------------------------------------------------------------------------------------------
TEMP_MATRIX(0, 1) = "DATE"
TEMP_MATRIX(0, 2) = "PRICE"
TEMP_MATRIX(0, 3) = "LOG(PRICE)"
TEMP_MATRIX(0, 4) = "LOG(RETURN)"
TEMP_MATRIX(0, 5) = "ZERO_MEAN" 'X(T) = X(T) - M
TEMP_MATRIX(0, 6) = "COND VAR"
TEMP_MATRIX(0, 7) = "LOG LIKELIHOOD"
TEMP_MATRIX(0, 8) = "COND SIGMA"
TEMP_MATRIX(0, 9) = "GARCH SIGMA"
TEMP_MATRIX(0, 10) = "RESIDUES"
'---------------------------------------------------------------------------------------------
i = 1
TEMP_MATRIX(i, 1) = DATE_VECTOR(i, 1)
TEMP_MATRIX(i, 2) = DATA_VECTOR(i, 1)
TEMP_MATRIX(i, 3) = Log(DATA_VECTOR(i, 1))
For j = 4 To 10: TEMP_MATRIX(i, j) = "": Next j
'---------------------------------------------------------------------------------------------
i = 2
TEMP_MATRIX(i, 1) = DATE_VECTOR(i, 1)
TEMP_MATRIX(i, 2) = DATA_VECTOR(i, 1)
TEMP_MATRIX(i, 3) = Log(DATA_VECTOR(i, 1))
TEMP_MATRIX(i, 4) = TEMP_MATRIX(i, 3) - TEMP_MATRIX(i - 1, 3)
TEMP_MATRIX(i, 5) = TEMP_MATRIX(i, 4) - MEAN_VAL
TEMP_MATRIX(i, 6) = VAR_VAL
TEMP_MATRIX(i, 7) = -Log(TEMP_MATRIX(i, 6)) - (TEMP_MATRIX(i, 5)) ^ 2 / TEMP_MATRIX(i, 6)
TEMP_SUM = TEMP_MATRIX(i, 7)
TEMP_MATRIX(i, 8) = TEMP_MATRIX(i, 6) ^ 0.5
TEMP_MATRIX(i, 9) = TEMP_MATRIX(i, 8) * (COUNT_BASIS) ^ 0.5
TEMP_MATRIX(i, 10) = TEMP_MATRIX(i, 5) / (TEMP_MATRIX(i, 6)) ^ 0.5

'---------------------------------------------------------------------------------------------
For i = 3 To NROWS
'---------------------------------------------------------------------------------------------
    TEMP_MATRIX(i, 1) = DATE_VECTOR(i, 1)
    TEMP_MATRIX(i, 2) = DATA_VECTOR(i, 1)
    TEMP_MATRIX(i, 3) = Log(DATA_VECTOR(i, 1))
    TEMP_MATRIX(i, 4) = TEMP_MATRIX(i, 3) - TEMP_MATRIX(i - 1, 3)
    TEMP_MATRIX(i, 5) = TEMP_MATRIX(i, 4) - MEAN_VAL
    TEMP_MATRIX(i, 6) = OMEGA_VAL + ALPHA_VAL * (TEMP_MATRIX(i - 1, 5)) ^ 2 + BETA_VAL * TEMP_MATRIX(i - 1, 6)
    TEMP_MATRIX(i, 7) = -Log(TEMP_MATRIX(i, 6)) - (TEMP_MATRIX(i, 5)) ^ 2 / TEMP_MATRIX(i, 6)
    TEMP_SUM = TEMP_SUM + TEMP_MATRIX(i, 7)
    
    TEMP_MATRIX(i, 8) = TEMP_MATRIX(i, 6) ^ 0.5
    TEMP_MATRIX(i, 9) = TEMP_MATRIX(i, 8) * (COUNT_BASIS) ^ 0.5
    TEMP_MATRIX(i, 10) = TEMP_MATRIX(i, 5) / TEMP_MATRIX(i, 6) ^ 0.5
'---------------------------------------------------------------------------------------------
Next i
'---------------------------------------------------------------------------------------------
Select Case OUTPUT
'---------------------------------------------------------------------------------------------
Case 0
    GARCH_MODEL1_FUNC = TEMP_MATRIX
'---------------------------------------------------------------------------------------------
Case 1
    GARCH_MODEL1_FUNC = 0.5 * (-Log(2 * PI_VAL) * (NROWS - 1) + TEMP_SUM) * 1000
'---------------------------------------------------------------------------------------------
Case Else
    For i = 2 To NROWS
        DELTA_VAL = Log(DATA_VECTOR(i, 1) / DATA_VECTOR(i - 1, 1))
        KURT_VAL = KURT_VAL + ((DELTA_VAL - MEAN_VAL) / SIGMA_VAL) ^ 4
    Next i
    'KURT_VAL = KURT_VAL / (NROWS - 1)
    KURT_VAL = (KURT_VAL * ((NROWS - 1) * ((NROWS - 1) + 1) / (((NROWS - 1) - 1) * ((NROWS - 1) - 2) * ((NROWS - 1) - 3)))) - ((3 * ((NROWS - 1) - 1) ^ 2 / (((NROWS - 1) - 2) * ((NROWS - 1) - 3)))) 'Excel Definition
    KURT_VAL = KURT_VAL + 3
    FACTOR_VAL = 1 / 3 * (-KURT_VAL * BETA_VAL + (-2 * KURT_VAL ^ 2 * BETA_VAL ^ 2 + 3 * KURT_VAL ^ 2 + 6 * KURT_VAL - 6 * KURT_VAL * BETA_VAL ^ 2) ^ (1 / 2)) / (KURT_VAL + 2)
    
    ReDim TEMP_MATRIX(1 To 30, 1 To 1)
    
    'Enter a first guess for OMEGA, ALPHA, BETA (may be 0.0, 0.1, 0.7).
    'Insert "2nd guess" for OMEGA and ALPHA (by copy / paste special / values),
    'which sets parameters according to Kurtosis and Variance and
    'if data are very skewed ... Garch(p,q) assumes that to be zero :-)
    
    'Now run Solver with tolerance 1 - 3 % (then run it again with tolerance
    '0.1% if you feel the need)
    
    'If one wants Garch Kurtosis to exist (there is no need for that):
    'include the constraint kurtConstraint < 1 to be observed by Solver
    'but i prefer to do it not (bad volatility results)
    
    'It might be a question whether the model matches with data ...
    
    'If so it should converge and i use it in the appended dirty Excel
    'solution as follows to calibrate and estimate: starting with a
    'and b at ~ 0.4 i set OMEGA such that OMEGA / (a + b) is close to the
    'usual, unconditioned variance (simply using 'goal seek').
    
    'Then use 'solver' in Excel to find a first maximum for LLH with
    '_mild exactness demands_ (solver trapps into local extrema, so
    'give it a chance). In a 2nd or 3rd step refine your demands
    '(as i said: dirty) to pick the maximum.
    
    'Adding some data you hope the extremum to depend continously on
    'them and +-10% difference in the two variances should not have
    'to much consequences for the estimated vol (below 1 vol pt).
    
    TEMP_MATRIX(1, 1) = "Likelihood"
    TEMP_MATRIX(2, 1) = 0.5 * (-Log(2 * PI_VAL) * (NROWS - 1) + TEMP_SUM) * 1000
    TEMP_MATRIX(3, 1) = "log Likelihood function * 1000, to be maximized"
    TEMP_MATRIX(4, 1) = ""
    
    TEMP_MATRIX(5, 1) = "OMEGA (a0) >=0"
    TEMP_MATRIX(6, 1) = OMEGA_VAL
    TEMP_MATRIX(7, 1) = "Omega Guess = " & Format(VAR_VAL * (1 - FACTOR_VAL - BETA_VAL), "0.000000")
    TEMP_MATRIX(8, 1) = ""
    
    TEMP_MATRIX(9, 1) = "ALPHA (a1) >=0"
    TEMP_MATRIX(10, 1) = ALPHA_VAL
    TEMP_MATRIX(11, 1) = "Alpha Guess = " & Format(FACTOR_VAL, "0.000000")
    TEMP_MATRIX(12, 1) = ""
                
    TEMP_MATRIX(13, 1) = "BETA (b1) >=0"
    TEMP_MATRIX(14, 1) = BETA_VAL
    TEMP_MATRIX(15, 1) = "After entering: BETA , ALPHA, OMEGA - new values are computed which match statistical mean and kurtosis. Use this values as ALPHA and BETA for a start (BETA remains unchanged)."
    TEMP_MATRIX(16, 1) = ""
                
    TEMP_MATRIX(17, 1) = "Var condition <= 0.999999" ' < 1 (~0,99999)"
    TEMP_MATRIX(18, 1) = ALPHA_VAL + BETA_VAL
    TEMP_MATRIX(19, 1) = ""
    
    TEMP_MATRIX(20, 1) = "Kurt condition < 1 (~0,99)"
    TEMP_MATRIX(21, 1) = BETA_VAL ^ 2 + 2 * ALPHA_VAL * BETA_VAL + 3 * ALPHA_VAL '< 1 (~0,99)
    TEMP_MATRIX(22, 1) = "If you want that a conditional kurtosis exists, then set constraint to be < 0.99 (or similar, needs some experimentations) otherwise delete that constraint: I prefer not to use it"
    TEMP_MATRIX(23, 1) = ""

    TEMP_MATRIX(24, 1) = "GARCH Var"
    TEMP_MATRIX(25, 1) = OMEGA_VAL / (1 - ALPHA_VAL - BETA_VAL)
    TEMP_MATRIX(26, 1) = "estimated conditional variance = Garch variance = OMEGA / (1-ALPHA-BETA)"
    TEMP_MATRIX(27, 1) = ""
                
    TEMP_MATRIX(28, 1) = "Est. Kurt"
    TEMP_MATRIX(29, 1) = 6 * ALPHA_VAL / (1 - (BETA_VAL ^ 2 + 2 * ALPHA_VAL * BETA_VAL + 3 * ALPHA_VAL))
    TEMP_MATRIX(30, 1) = "estimated Kurtosis (if it exists, check whether kurt constraint < 1)"
    
    GARCH_MODEL1_FUNC = TEMP_MATRIX

'---------------------------------------------------------------------------------------------
End Select
'---------------------------------------------------------------------------------------------

Exit Function
ERROR_LABEL:
GARCH_MODEL1_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : GARCH_MODEL2_FUNC
'DESCRIPTION   : Generalized Autoregressive Conditional Heteroscedasticity Function B
'LIBRARY       : STATISTICS
'GROUP         : GARCH
'ID            : 002
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/08/2008
'************************************************************************************
'************************************************************************************

Function GARCH_MODEL2_FUNC(ByRef DATE_RNG As Variant, _
ByRef DATA_RNG As Variant, _
ByVal OMEGA_VAL As Double, _
ByVal ALPHA_VAL As Double, _
ByVal BETA_VAL As Double, _
Optional ByVal CONFIDENCE_VAL As Double = 0.95, _
Optional ByVal OUTPUT As Integer = 0)

Dim i As Long
Dim j As Long
Dim k As Long

Dim NROWS As Long

Dim PI_VAL As Double

Dim TEMP1_SUM As Double
Dim TEMP2_SUM As Double

Dim G_VAL As Double 'GARCH(1,1) Unconditional Variance
Dim MEAN_VAL As Double 'Unconditional Standard Deviation
Dim VAR_VAL As Double 'Sample Unconditional Variance
Dim SIGMA_VAL As Double
Dim FACTOR_VAL As Double

Dim TEMP_VECTOR As Variant
Dim TEMP_MATRIX As Variant
Dim DATA_VECTOR As Variant
Dim DATE_VECTOR As Variant

On Error GoTo ERROR_LABEL

DATE_VECTOR = DATE_RNG
If UBound(DATE_VECTOR, 1) = 1 Then: DATE_VECTOR = MATRIX_TRANSPOSE_FUNC(DATE_VECTOR)
DATA_VECTOR = DATA_RNG
If UBound(DATA_VECTOR, 1) = 1 Then: DATA_VECTOR = MATRIX_TRANSPOSE_FUNC(DATA_VECTOR)

If UBound(DATE_VECTOR, 1) <> UBound(DATA_VECTOR, 1) Then: GoTo ERROR_LABEL
NROWS = UBound(DATA_VECTOR, 1)

PI_VAL = 3.14159265358979
FACTOR_VAL = NORMSINV_FUNC(1 - CONFIDENCE_VAL, 0, 1, 0)
'----------------------------------------------------------------------------------
TEMP1_SUM = 0
For i = 2 To NROWS: TEMP1_SUM = TEMP1_SUM + Log(DATA_VECTOR(i, 1) / DATA_VECTOR(i - 1, 1)): Next i
MEAN_VAL = TEMP1_SUM / (NROWS - 1)
TEMP1_SUM = 0
For i = 2 To NROWS: TEMP1_SUM = TEMP1_SUM + ((Log(DATA_VECTOR(i, 1) / DATA_VECTOR(i - 1, 1))) - MEAN_VAL) ^ 2: Next i
SIGMA_VAL = (TEMP1_SUM / (NROWS - 2)) ^ 0.5
'----------------------------------------------------------------------------------

VAR_VAL = SIGMA_VAL ^ 2
G_VAL = OMEGA_VAL / (1 - ALPHA_VAL - BETA_VAL)

TEMP1_SUM = 0: TEMP2_SUM = 0
ReDim TEMP_MATRIX(0 To NROWS, 1 To 13)

TEMP_MATRIX(0, 1) = "DATE"
TEMP_MATRIX(0, 2) = "PRICE"
TEMP_MATRIX(0, 3) = "LOG RETURN"
TEMP_MATRIX(0, 4) = "CUMULATIVE"
TEMP_MATRIX(0, 5) = "DMEAN"
TEMP_MATRIX(0, 6) = "DEMEAN^2"
TEMP_MATRIX(0, 7) = "CONDITIONAL_VARIANCE"
TEMP_MATRIX(0, 8) = "LOG LIKE"
TEMP_MATRIX(0, 9) = "CONDITIONAL SIGMA"
TEMP_MATRIX(0, 10) = "UNCONDITIONAL SIGMA"
TEMP_MATRIX(0, 11) = "CONDITIONAL VAR"
TEMP_MATRIX(0, 12) = "1: r < VAR else r > VAR"
TEMP_MATRIX(0, 13) = "TAIL LOSSES"

i = 1
TEMP_MATRIX(i, 1) = DATE_VECTOR(i, 1)
TEMP_MATRIX(i, 2) = DATA_VECTOR(i, 1)
For j = 3 To 13: TEMP_MATRIX(i, j) = "": Next j

i = 2
TEMP_MATRIX(i, 1) = DATE_VECTOR(i, 1)
TEMP_MATRIX(i, 2) = DATA_VECTOR(i, 1)
TEMP_MATRIX(i, 3) = Log(DATA_VECTOR(i, 1) / DATA_VECTOR(i - 1, 1))
TEMP_MATRIX(i, 4) = TEMP_MATRIX(i, 3)
TEMP_MATRIX(i, 5) = TEMP_MATRIX(i, 3) - MEAN_VAL
TEMP_MATRIX(i, 6) = TEMP_MATRIX(i, 5) ^ 2
TEMP_MATRIX(i, 7) = VAR_VAL 'Starting point can be either
For j = 8 To 13: TEMP_MATRIX(i, j) = "": Next j         'unconditional variance or zero or...

k = 0
For i = 3 To NROWS
    TEMP_MATRIX(i, 1) = DATE_VECTOR(i, 1)
    TEMP_MATRIX(i, 2) = DATA_VECTOR(i, 1)
    TEMP_MATRIX(i, 3) = Log(DATA_VECTOR(i, 1) / DATA_VECTOR(i - 1, 1))
    TEMP_MATRIX(i, 4) = TEMP_MATRIX(i - 1, 4) + TEMP_MATRIX(i, 3)
    TEMP_MATRIX(i, 5) = TEMP_MATRIX(i, 3) - MEAN_VAL
    TEMP_MATRIX(i, 6) = TEMP_MATRIX(i, 5) ^ 2
    TEMP_MATRIX(i, 7) = OMEGA_VAL + ALPHA_VAL * TEMP_MATRIX(i - 1, 6) + BETA_VAL * TEMP_MATRIX(i - 1, 7)
    TEMP_MATRIX(i, 8) = Log((1 / Sqr(2 * PI_VAL * TEMP_MATRIX(i, 7))) * Exp(-0.5 * TEMP_MATRIX(i, 6) / TEMP_MATRIX(i, 7)))
    TEMP1_SUM = TEMP_MATRIX(i, 8) + TEMP1_SUM
    TEMP_MATRIX(i, 9) = TEMP_MATRIX(i, 7) ^ 0.5
    TEMP_MATRIX(i, 10) = SIGMA_VAL
    
    TEMP_MATRIX(i, 11) = MEAN_VAL + FACTOR_VAL * TEMP_MATRIX(i, 9)
    
    If TEMP_MATRIX(i - 1, 3) < TEMP_MATRIX(i, 11) Then
        TEMP_MATRIX(i, 12) = 1
        k = k + TEMP_MATRIX(i, 12)
        TEMP_MATRIX(i, 13) = TEMP_MATRIX(i - 1, 3) * TEMP_MATRIX(i, 12)
        TEMP2_SUM = TEMP2_SUM + TEMP_MATRIX(i, 13)
    Else
        TEMP_MATRIX(i, 12) = ""
        TEMP_MATRIX(i, 13) = ""
    End If
Next i

'-------------------------------------------------------------------
Select Case OUTPUT
'-------------------------------------------------------------------
Case 0
'-------------------------------------------------------------------
    GARCH_MODEL2_FUNC = TEMP_MATRIX
'-------------------------------------------------------------------
Case Else 'VaR BACKTESTING
'-------------------------------------------------------------------
    ReDim TEMP_VECTOR(1 To 5, 1 To 2) '
    TEMP_VECTOR(1, 1) = "Likelihoods"
    TEMP_VECTOR(1, 2) = TEMP1_SUM
    
    TEMP_VECTOR(2, 1) = "#Obs"
    TEMP_VECTOR(2, 2) = NROWS - 2
    
    TEMP_VECTOR(3, 1) = "#VaR Events"
    TEMP_VECTOR(3, 2) = k
    
    TEMP_VECTOR(4, 1) = "%VaR Events"
    If TEMP_VECTOR(2, 2) <> 0 Then
        TEMP_VECTOR(4, 2) = 100 * TEMP_VECTOR(3, 2) / TEMP_VECTOR(2, 2)
    Else
        TEMP_VECTOR(4, 2) = "N/A"
    End If
    TEMP_VECTOR(5, 1) = "Expected Tail Loss"
    If TEMP_VECTOR(3, 2) <> 0 Then
        TEMP_VECTOR(5, 2) = 100 * TEMP2_SUM / TEMP_VECTOR(3, 2)
    Else
        TEMP_VECTOR(5, 2) = "N/A"
    End If
        
    GARCH_MODEL2_FUNC = TEMP_VECTOR
'-------------------------------------------------------------------
End Select
'-------------------------------------------------------------------

Exit Function
ERROR_LABEL:
GARCH_MODEL2_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : GARCH_MODEL3_FUNC
'DESCRIPTION   : Generalized Autoregressive Conditional Heteroscedasticity
'with Variance Targeting Function
'LIBRARY       : STATISTICS
'GROUP         : GARCH
'ID            : 003
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/08/2008
'************************************************************************************
'************************************************************************************

Function GARCH_MODEL3_FUNC(ByRef DATE_RNG As Variant, _
ByRef DATA_RNG As Variant, _
ByVal OMEGA_VAL As Double, _
ByVal ALPHA_VAL As Double, _
ByVal BETA_VAL As Double, _
ByVal THETA_VAL As Double, _
Optional ByVal OUTPUT As Integer = 0)

'OMEGA WITH TARGET VARIANCE = GARCH_MODEL3_FUNC(1,1) * (1 - ALPHA_VAL - BETA)

Dim i As Long
Dim j As Long
Dim NROWS As Long

Dim PI_VAL As Double

Dim TEMP_SUM As Double
Dim MEAN_VAL As Double
Dim SIGMA_VAL As Double

Dim TEMP_MATRIX As Variant
Dim DATA_VECTOR As Variant
Dim DATE_VECTOR As Variant

On Error GoTo ERROR_LABEL

DATE_VECTOR = DATE_RNG
If UBound(DATE_VECTOR, 1) = 1 Then: DATE_VECTOR = MATRIX_TRANSPOSE_FUNC(DATE_VECTOR)
DATA_VECTOR = DATA_RNG
If UBound(DATA_VECTOR, 1) = 1 Then: DATA_VECTOR = MATRIX_TRANSPOSE_FUNC(DATA_VECTOR)

If UBound(DATE_VECTOR, 1) <> UBound(DATA_VECTOR, 1) Then: GoTo ERROR_LABEL
NROWS = UBound(DATA_VECTOR, 1)

PI_VAL = 3.14159265358979
'----------------------------------------------------------------------------------
TEMP_SUM = 0
For i = 2 To NROWS: TEMP_SUM = TEMP_SUM + Log(DATA_VECTOR(i, 1) / DATA_VECTOR(i - 1, 1)): Next i
MEAN_VAL = TEMP_SUM / (NROWS - 1)
TEMP_SUM = 0
For i = 2 To NROWS: TEMP_SUM = TEMP_SUM + ((Log(DATA_VECTOR(i, 1) / DATA_VECTOR(i - 1, 1))) - MEAN_VAL) ^ 2: Next i
SIGMA_VAL = (TEMP_SUM / (NROWS - 2)) ^ 0.5
'----------------------------------------------------------------------------------

ReDim TEMP_MATRIX(0 To NROWS, 1 To 6)
    
TEMP_MATRIX(0, 1) = "DATE"
TEMP_MATRIX(0, 2) = "PRICE"
TEMP_MATRIX(0, 3) = "LOG RETURN"
TEMP_MATRIX(0, 4) = "LOG RETURN ^ 2"
TEMP_MATRIX(0, 5) = "GARCH(1,1)"
TEMP_MATRIX(0, 6) = "MLE"

i = 1
TEMP_MATRIX(i, 1) = DATE_VECTOR(i, 1)
TEMP_MATRIX(i, 2) = DATA_VECTOR(i, 1)
For j = 3 To 6: TEMP_MATRIX(i, j) = "": Next j

i = 2
TEMP_MATRIX(i, 1) = DATE_VECTOR(i, 1)
TEMP_MATRIX(i, 2) = DATA_VECTOR(i, 1)
TEMP_MATRIX(i, 3) = Log(DATA_VECTOR(i, 3) / DATA_VECTOR(i - 1, 3)) 'LOG RETURNS
TEMP_MATRIX(i, 4) = TEMP_MATRIX(i, 3) ^ 2 'SQUARED-RETURNS
TEMP_MATRIX(i, 5) = (SIGMA_VAL) ^ 2 'GARCH_MODEL3_FUNC
TEMP_MATRIX(i, 6) = (-0.5 * Log(2 * PI_VAL) - 0.5 * Log(TEMP_MATRIX(i, 5)) - 0.5 * (TEMP_MATRIX(i, 4) / TEMP_MATRIX(i, 5))) 'MLE

TEMP_SUM = TEMP_MATRIX(i, 6)
For i = 3 To NROWS
    TEMP_MATRIX(i, 1) = DATE_VECTOR(i, 1)
    TEMP_MATRIX(i, 2) = DATA_VECTOR(i, 1)
    TEMP_MATRIX(i, 3) = Log(DATA_VECTOR(i, 3) / DATA_VECTOR(i - 1, 3)) 'LOG RETURNS
    TEMP_MATRIX(i, 4) = TEMP_MATRIX(i, 3) ^ 2 'SQUARED-RETURNS
    
    TEMP_MATRIX(i, 5) = OMEGA_VAL + BETA_VAL * TEMP_MATRIX(i - 1, 5) + ALPHA_VAL * ((TEMP_MATRIX(i - 1, 3) - THETA_VAL * (TEMP_MATRIX(i - 1, 5) ^ 0.5)) ^ 2)
    TEMP_MATRIX(i, 6) = (-0.5 * Log(2 * PI_VAL) - 0.5 * Log(TEMP_MATRIX(i, 5)) - 0.5 * (TEMP_MATRIX(i, 4) / TEMP_MATRIX(i, 5))) 'MLE
    
    TEMP_SUM = TEMP_SUM + TEMP_MATRIX(i, 6)
Next i

Select Case OUTPUT
Case 0
    GARCH_MODEL3_FUNC = TEMP_MATRIX
Case 1
    GARCH_MODEL3_FUNC = Array(TEMP_SUM, ALPHA_VAL * (1 + THETA_VAL ^ 2) + BETA_VAL)
End Select

Exit Function
ERROR_LABEL:
GARCH_MODEL3_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : GARCH_VOLATILITY_STRUCTURE_FUNC
'DESCRIPTION   : Volatility Term Structure
'LIBRARY       : STATISTICS
'GROUP         : GARCH
'ID            : 004
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/08/2008
'************************************************************************************
'************************************************************************************

Function GARCH_VOLATILITY_STRUCTURE_FUNC(ByVal NSIZE As Long, _
ByVal ALPHA_VAL As Double, _
ByVal BETA_VAL As Double, _
ByVal LT_VAR As Double, _
ByVal CURR_VAR As Double, _
Optional ByVal DELTA_TIME As Double = 1, _
Optional ByVal COUNT_BASIS As Long = 252)

'CURR_VAR = Day 0 Est. Variance
'LT_VAR = Long Term Variance

Dim i As Long

Dim TEMP1_SUM As Double
Dim TEMP2_SUM As Double
Dim TEMP_VALUE As Double

Dim TEMP_MATRIX As Variant

On Error GoTo ERROR_LABEL

ReDim TEMP_MATRIX(0 To NSIZE, 1 To 4)

TEMP_MATRIX(0, 1) = "PERIOD"
TEMP_MATRIX(0, 2) = "EST. VARIANCE"
TEMP_MATRIX(0, 3) = "CUM AVERAGE"
TEMP_MATRIX(0, 4) = "EST VOL"

TEMP_VALUE = BETA_VAL + ALPHA_VAL

TEMP1_SUM = 0: TEMP2_SUM = 0
For i = 1 To NSIZE
    TEMP1_SUM = TEMP1_SUM + DELTA_TIME
    TEMP_MATRIX(i, 1) = TEMP1_SUM
    TEMP_MATRIX(i, 2) = LT_VAR + (TEMP_VALUE ^ TEMP_MATRIX(i, 1)) * (CURR_VAR - LT_VAR)
    TEMP2_SUM = TEMP2_SUM + TEMP_MATRIX(i, 2)
    TEMP_MATRIX(i, 3) = 1 / (TEMP1_SUM + 1) * (CURR_VAR + TEMP2_SUM)
    TEMP_MATRIX(i, 4) = Sqr(TEMP_MATRIX(i, 3)) * Sqr(COUNT_BASIS)
Next i

GARCH_VOLATILITY_STRUCTURE_FUNC = TEMP_MATRIX

Exit Function
ERROR_LABEL:
GARCH_VOLATILITY_STRUCTURE_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : GARCH_SIMULATION_FUNC

'DESCRIPTION   : GARCH(1,1) simulation function to generate independent, but not
' identically distributed return series resulting in fat tails

'LIBRARY       : STATISTICS
'GROUP         : GARCH
'ID            : 005
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/08/2008
'************************************************************************************
'************************************************************************************

Function GARCH_SIMULATION_FUNC(ByVal NSIZE As Long, _
ByVal THRESD_RETURN As Variant, _
ByVal OMEGA_VAL As Double, _
ByVal ALPHA_VAL As Double, _
ByVal BETA_VAL As Double, _
Optional ByVal nLOOPS As Long = 100)

Dim i As Long
Dim j As Long

Dim RANDOM_VAL As Double
Dim FACTOR_VAL As Double

Dim TEMP_MATRIX As Variant
Dim RANDOM_MATRIX As Variant

On Error GoTo ERROR_LABEL

ReDim TEMP_MATRIX(1 To NSIZE, 1 To nLOOPS)

RANDOM_MATRIX = MATRIX_RANDOM_NORMAL_FUNC(NSIZE, nLOOPS, 0, 0, 1, 0)

For j = 1 To nLOOPS
    FACTOR_VAL = OMEGA_VAL / (1 - ALPHA_VAL - BETA_VAL)
    For i = 1 To NSIZE
        RANDOM_VAL = RANDOM_MATRIX(i, j) * Sqr(FACTOR_VAL)
        TEMP_MATRIX(i, j) = THRESD_RETURN + RANDOM_VAL
        FACTOR_VAL = OMEGA_VAL + ALPHA_VAL * (RANDOM_VAL ^ 2) + BETA_VAL * FACTOR_VAL
    Next i
Next j

GARCH_SIMULATION_FUNC = TEMP_MATRIX

Exit Function
ERROR_LABEL:
GARCH_SIMULATION_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : GARCH_LOGLIKE_3D_PLOT_FUNC
'DESCRIPTION   : Loglikehood Garch(1,1) 3D Plot function
'LIBRARY       : STATISTICS
'GROUP         : GARCH
'ID            : 006
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/08/2008
'************************************************************************************
'************************************************************************************

Function GARCH_LOGLIKE_3D_PLOT_FUNC(ByRef DATA_RNG As Variant, _
Optional ByRef OMEGA0_VAL As Double = 3.18002190070703E-06, _
Optional ByRef ALPHA0_VAL As Double = 0.08, _
Optional ByRef BETA0_VAL As Double = 0.9, _
Optional ByRef DELTA_ALPHA As Double = 0.01, _
Optional ByRef DELTA_BETA As Double = 0.05, _
Optional ByVal DATA_TYPE As Integer = 1, _
Optional ByVal NBINS As Long = 10)

Dim i As Long
Dim j As Long
Dim k As Long
Dim l As Long

Dim BETA_VAL As Double
Dim OMEGA_VAL As Double
Dim ALPHA_VAL As Double

Dim A0_VAL As Double
Dim A1_VAL As Double

Dim B0_VAL As Double
Dim B1_VAL As Double

Dim X_VAL As Double
Dim Y_VAL As Double
Dim MAX_VAL As Double

Dim TEMP_ALPHA As Double
Dim TEMP_BETA As Double

Dim XTEMP_VAL As Double
Dim YTEMP_VAL As Double

Dim LOKLIKE_VAL As Double

Dim DATA_VECTOR As Variant
Dim TEMP_MATRIX As Variant

On Error GoTo ERROR_LABEL

'-----------------------------------------------------------------------------
'-----------------------------------------------------------------------------
OMEGA_VAL = OMEGA0_VAL
ALPHA_VAL = ALPHA0_VAL
BETA_VAL = BETA0_VAL
'-----------------------------------------------------------------------------
'-----------------------------------------------------------------------------

'PI_VAL = 3.14159265358979
DATA_VECTOR = DATA_RNG
If UBound(DATA_VECTOR, 1) = 1 Then: DATA_VECTOR = MATRIX_TRANSPOSE_FUNC(DATA_VECTOR)
If DATA_TYPE <> 0 Then: DATA_VECTOR = MATRIX_PERCENT_FUNC(DATA_VECTOR, 1)

TEMP_ALPHA = ALPHA_VAL
TEMP_BETA = BETA_VAL
MAX_VAL = -2 ^ 52

A0_VAL = ALPHA_VAL - DELTA_ALPHA
A1_VAL = ALPHA_VAL + DELTA_ALPHA

B0_VAL = BETA_VAL - DELTA_BETA
B1_VAL = BETA_VAL + DELTA_BETA

DELTA_ALPHA = DELTA_ALPHA / 4
DELTA_BETA = DELTA_BETA / 4

ReDim TEMP_MATRIX(1 To (NBINS + 1) * 2 + 1, 1 To NBINS + 1)

j = 0
For k = 0 To NBINS
    X_VAL = A0_VAL + k * DELTA_ALPHA
    ALPHA_VAL = X_VAL
    TEMP_MATRIX(1, 1 + j) = X_VAL
    j = j + 1
Next k
i = 1: j = 0
For k = 0 To NBINS
    Y_VAL = B0_VAL + k * DELTA_BETA
    BETA_VAL = Y_VAL
    TEMP_MATRIX(1 + i, 1 + j) = Y_VAL
    For l = 0 To NBINS
        X_VAL = A0_VAL + l * DELTA_ALPHA
        ALPHA_VAL = X_VAL
        LOKLIKE_VAL = GARCH_LOGLIKE2_FUNC(DATA_VECTOR, OMEGA_VAL, ALPHA_VAL, BETA_VAL, 0)
        If j > 0 Then: TEMP_MATRIX(1 + i, 1 + j) = ""
        TEMP_MATRIX(2 + i, 1 + j) = LOKLIKE_VAL
        j = j + 1
        If MAX_VAL < LOKLIKE_VAL Then
            MAX_VAL = LOKLIKE_VAL
            XTEMP_VAL = X_VAL
            YTEMP_VAL = Y_VAL
        End If
    Next l
    i = i + 2
    j = 0
Next k

GARCH_LOGLIKE_3D_PLOT_FUNC = TEMP_MATRIX

Exit Function
ERROR_LABEL:
GARCH_LOGLIKE_3D_PLOT_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : GARCH_LOGLIKE2_REFINE_PARAM_FUNC
'DESCRIPTION   : Generalized Autoregressive Conditional Heteroscedasticity
'initial paremeters for the parameterrization
'LIBRARY       : STATISTICS
'GROUP         : GARCH
'ID            : 007
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/08/2008
'************************************************************************************
'************************************************************************************

Function GARCH_LOGLIKE2_REFINE_PARAM_FUNC(ByRef DATA_RNG As Variant, _
Optional ByRef OMEGA0_VAL As Double = 3.18002190070703E-06, _
Optional ByRef ALPHA0_VAL As Double = 0.08, _
Optional ByRef BETA0_VAL As Double = 0.9, _
Optional ByVal DATA_TYPE As Integer = 1, _
Optional ByVal nLOOPS As Long = 10, _
Optional ByVal epsilon As Double = 0.00001)

Dim i As Long
Dim NROWS As Long

'Dim PI_VAL As Double

Dim OMEGA_VAL As Double
Dim ALPHA_VAL As Double
Dim BETA_VAL As Double

Dim B0_VAL As Double
Dim B1_VAL As Double
Dim B2_VAL As Double
Dim Z_VAL As Double

Dim C1_VAL As Double
Dim C2_VAL As Double
Dim C3_VAL As Double

Dim MAX_VAL As Double
'Dim NORM_VAL As Double
'Dim FACTOR_VAL As Double

Dim GRAD1_VAL As Double
Dim GRAD2_VAL As Double
Dim GRAD3_VAL As Double

Dim LOKLIKE_VAL As Double

Dim DATA_VECTOR As Variant

On Error GoTo ERROR_LABEL

'-----------------------------------------------------------------------------
OMEGA_VAL = OMEGA0_VAL
ALPHA_VAL = ALPHA0_VAL
BETA_VAL = BETA0_VAL
'-----------------------------------------------------------------------------

'PI_VAL = 3.14159265358979
DATA_VECTOR = DATA_RNG
If UBound(DATA_VECTOR, 1) = 1 Then: DATA_VECTOR = MATRIX_TRANSPOSE_FUNC(DATA_VECTOR)
If DATA_TYPE <> 0 Then: DATA_VECTOR = MATRIX_PERCENT_FUNC(DATA_VECTOR, 1)
NROWS = UBound(DATA_VECTOR, 1)

'FACTOR_VAL = -0.5 * Log(2 * PI_VAL) + Log(0.05)
LOKLIKE_VAL = GARCH_LOGLIKE2_FUNC(DATA_VECTOR, OMEGA_VAL, ALPHA_VAL, BETA_VAL, 0)

B0_VAL = OMEGA_VAL
B1_VAL = ALPHA_VAL
B2_VAL = BETA_VAL

C1_VAL = 0.05 * B0_VAL
C2_VAL = 0.05 * B1_VAL
C3_VAL = 0.05 * B2_VAL

MAX_VAL = LOKLIKE_VAL
'get current MAX_VAL

For i = 0 To nLOOPS

    Z_VAL = LOKLIKE_VAL
    
    OMEGA_VAL = B0_VAL + C1_VAL        ' change B0_VAL
    LOKLIKE_VAL = GARCH_LOGLIKE2_FUNC(DATA_VECTOR, OMEGA_VAL, ALPHA_VAL, BETA_VAL, 0)
    GRAD1_VAL = LOKLIKE_VAL - Z_VAL ' calculate dz
    Z_VAL = LOKLIKE_VAL  'get new z-value
    If Z_VAL > MAX_VAL Then 'increase?
        MAX_VAL = Z_VAL
        B0_VAL = B0_VAL + C1_VAL 'update B0_VAL
    End If
    
    ALPHA_VAL = B1_VAL + C2_VAL
    LOKLIKE_VAL = GARCH_LOGLIKE2_FUNC(DATA_VECTOR, OMEGA_VAL, ALPHA_VAL, BETA_VAL, 0)
    GRAD2_VAL = LOKLIKE_VAL - Z_VAL
    Z_VAL = LOKLIKE_VAL
    If Z_VAL > MAX_VAL Then
        MAX_VAL = Z_VAL
        B1_VAL = B1_VAL + C2_VAL
    End If
    
    BETA_VAL = B2_VAL + C3_VAL
    LOKLIKE_VAL = GARCH_LOGLIKE2_FUNC(DATA_VECTOR, OMEGA_VAL, ALPHA_VAL, BETA_VAL, 0)
    
    GRAD3_VAL = LOKLIKE_VAL - Z_VAL
    Z_VAL = LOKLIKE_VAL
    If Z_VAL > MAX_VAL Then
        MAX_VAL = Z_VAL
        B2_VAL = B2_VAL + C3_VAL
    End If
    
    If (B1_VAL + B2_VAL) >= 1 Then B1_VAL = (1 - epsilon) - B2_VAL
    If B0_VAL < 0 Then B0_VAL = epsilon
    LOKLIKE_VAL = GARCH_LOGLIKE2_FUNC(DATA_VECTOR, OMEGA_VAL, ALPHA_VAL, BETA_VAL, 0)

Next i

GARCH_LOGLIKE2_REFINE_PARAM_FUNC = Array(B0_VAL, B1_VAL, B2_VAL, MAX_VAL)

Exit Function
ERROR_LABEL:
GARCH_LOGLIKE2_REFINE_PARAM_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : GARCH_FITTING_FUNC
'DESCRIPTION   : This Function implements fitting of the Generalized
'Autoregressive Conditional Heteroscedasticity
'LIBRARY       : STATISTICS
'GROUP         : GARCH
'ID            : 008
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/08/2008
'************************************************************************************
'************************************************************************************

Function GARCH_FITTING_FUNC(ByRef DATA_RNG As Variant, _
ByRef PARAM_RNG As Variant, _
Optional ByRef CONST_RNG As Variant, _
Optional ByVal DATA_TYPE As Integer = 1, _
Optional ByVal VERSION As Integer = 0)

Dim i As Long
Dim j As Long
Dim CONST_BOX As Variant
Dim DATA_VECTOR As Variant
Dim PARAM_VECTOR As Variant

On Error GoTo ERROR_LABEL

PUB_GARCH_DATA = 0
PUB_GARCH_OBJ = 0

DATA_VECTOR = DATA_RNG
If UBound(DATA_VECTOR, 1) = 1 Then: DATA_VECTOR = MATRIX_TRANSPOSE_FUNC(DATA_VECTOR)
If DATA_TYPE <> 0 Then: DATA_VECTOR = MATRIX_PERCENT_FUNC(DATA_VECTOR, 1)

PARAM_VECTOR = PARAM_RNG
If UBound(PARAM_VECTOR, 1) = 1 Then: PARAM_VECTOR = MATRIX_TRANSPOSE_FUNC(PARAM_VECTOR)
PUB_GARCH_DATA = DATA_VECTOR
PUB_GARCH_OBJ = VERSION

If IsArray(CONST_RNG) = True Then
    CONST_BOX = MATRIX_TRANSPOSE_FUNC(CONST_RNG)
Else
    If VERSION > 2 Then j = 4 Else j = 3
    ReDim CONST_BOX(1 To 2, 1 To j) 'Initial
    For i = 1 To j
        CONST_BOX(1, i) = 0
        CONST_BOX(2, i) = 1
    Next i
End If


PUB_FACTOR_VAL = -1
PARAM_VECTOR = NELDER_MEAD_OPTIMIZATION2_FUNC("GARCH_FITTING_OBJ_FUNC", PUB_GARCH_DATA, PARAM_VECTOR)
GARCH_FITTING_FUNC = PARAM_VECTOR

Exit Function
ERROR_LABEL:
GARCH_FITTING_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : GARCH_INITIAL_PARAMETERS_FUNC
'DESCRIPTION   : This Function implements fitting of the Generalized
'Autoregressive Conditional Heteroscedasticity
'LIBRARY       : STATISTICS
'GROUP         : GARCH
'ID            : 009
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/08/2008
'************************************************************************************
'************************************************************************************

Function GARCH_INITIAL_PARAMETERS_FUNC(ByRef DATA_RNG As Variant, _
Optional ByRef PARAM_RNG As Variant, _
Optional ByRef CONST_RNG As Variant, _
Optional ByVal DATA_TYPE As Integer = 1, _
Optional ByVal VERSION As Integer = 0)

Dim i As Long
Dim j As Long
Dim CONST_BOX As Variant
Dim PARAM_VECTOR As Variant
Dim DATA_VECTOR As Variant

On Error GoTo ERROR_LABEL

PUB_GARCH_DATA = 0: PUB_GARCH_OBJ = 0
DATA_VECTOR = DATA_RNG
If UBound(DATA_VECTOR, 1) = 1 Then: DATA_VECTOR = MATRIX_TRANSPOSE_FUNC(DATA_VECTOR)
If DATA_TYPE <> 0 Then: DATA_VECTOR = MATRIX_PERCENT_FUNC(DATA_VECTOR, 1)

PUB_GARCH_DATA = DATA_VECTOR
PUB_GARCH_OBJ = VERSION

If IsArray(CONST_RNG) = True Then
    CONST_BOX = MATRIX_TRANSPOSE_FUNC(CONST_RNG)
Else
    If VERSION > 2 Then j = 4 Else j = 3
    ReDim CONST_BOX(1 To 2, 1 To j) 'Initial
    For i = 1 To j
        CONST_BOX(1, i) = 0
        CONST_BOX(2, i) = 1
    Next i
End If
PUB_FACTOR_VAL = 1

'---------------------------------------------------------------------------------------------------------------
If IsArray(PARAM_RNG) = True Then
'---------------------------------------------------------------------------------------------------------------
    PARAM_VECTOR = PARAM_RNG
    If UBound(PARAM_VECTOR, 1) = 1 Then: PARAM_VECTOR = MATRIX_TRANSPOSE_FUNC(PARAM_VECTOR)
    PARAM_VECTOR = NELDER_MEAD_OPTIMIZATION_FRAME_FUNC("GARCH_FITTING_OBJ_FUNC", PARAM_VECTOR, CONST_BOX, False, 0, 10000, 10 ^ -10)
'---------------------------------------------------------------------------------------------------------------
Else
'---------------------------------------------------------------------------------------------------------------
    PARAM_VECTOR = PIKAIA_OPTIMIZATION_FUNC("GARCH_FITTING_OBJ_FUNC", CONST_BOX, False, , , , , , , , , , , , , , 0)
'---------------------------------------------------------------------------------------------------------------
End If
'---------------------------------------------------------------------------------------------------------------

GARCH_INITIAL_PARAMETERS_FUNC = PARAM_VECTOR

Exit Function
ERROR_LABEL:
GARCH_INITIAL_PARAMETERS_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : GARCH_FITTING_OBJ_FUNC
'DESCRIPTION   : Objective Function of the Generalized Autoregressive
'Conditional Heteroscedasticity
'LIBRARY       : STATISTICS
'GROUP         : GARCH
'ID            : 010
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/08/2008
'************************************************************************************
'************************************************************************************

Function GARCH_FITTING_OBJ_FUNC(Optional ByRef DATA_RNG As Variant, _
Optional ByRef PARAM_RNG As Variant)

Dim OMEGA_VAL As Double
Dim ALPHA_VAL As Double
Dim BETA_VAL As Double
Dim THETA_VAL As Double

Dim YTEMP_VAL As Double
Dim XTEMP_VAL As Double

Dim PARAM_VECTOR As Variant
Dim DATA_VECTOR As Variant

On Error GoTo ERROR_LABEL

If IsArray(PARAM_RNG) = False Then
    DATA_VECTOR = PUB_GARCH_DATA
    PARAM_VECTOR = DATA_RNG
Else
    DATA_VECTOR = DATA_RNG
    PARAM_VECTOR = PARAM_RNG
End If
If UBound(DATA_VECTOR, 1) = 1 Then
    DATA_VECTOR = MATRIX_TRANSPOSE_FUNC(DATA_VECTOR)
End If
If UBound(PARAM_VECTOR, 1) = 1 Then
    PARAM_VECTOR = MATRIX_TRANSPOSE_FUNC(PARAM_VECTOR)
End If

'-----------------------------------------------------------------------------------
Select Case PUB_GARCH_OBJ
'-----------------------------------------------------------------------------------
Case 0
'-----------------------------------------------------------------------------------
    OMEGA_VAL = PARAM_VECTOR(1, 1)
    ALPHA_VAL = PARAM_VECTOR(2, 1)
    BETA_VAL = PARAM_VECTOR(3, 1)
    YTEMP_VAL = GARCH_LOGLIKE1_FUNC(DATA_VECTOR, OMEGA_VAL, ALPHA_VAL, BETA_VAL, 0)
'-----------------------------------------------------------------------------------
Case 1
'-----------------------------------------------------------------------------------
    OMEGA_VAL = PARAM_VECTOR(1, 1)
    ALPHA_VAL = PARAM_VECTOR(2, 1)
    BETA_VAL = PARAM_VECTOR(3, 1)
    YTEMP_VAL = GARCH_LOGLIKE2_FUNC(DATA_VECTOR, OMEGA_VAL, ALPHA_VAL, BETA_VAL, 0)
'-----------------------------------------------------------------------------------
Case Else
'-----------------------------------------------------------------------------------
    OMEGA_VAL = PARAM_VECTOR(1, 1)
    ALPHA_VAL = PARAM_VECTOR(2, 1)
    BETA_VAL = PARAM_VECTOR(3, 1)
    THETA_VAL = PARAM_VECTOR(4, 1)
    YTEMP_VAL = GARCH_LOGLIKE3_FUNC(DATA_VECTOR, OMEGA_VAL, ALPHA_VAL, BETA_VAL, THETA_VAL, 0)
'-----------------------------------------------------------------------------------
End Select
'-----------------------------------------------------------------------------------

'If YTEMP_VAL = -PUB_EPSILON Then: GoTo ERROR_LABEL 'Error Flag
If GARCH_FITTING_CONST_FUNC(PARAM_VECTOR) = False Then
    XTEMP_VAL = PUB_EPSILON
Else
    XTEMP_VAL = 1
End If
GARCH_FITTING_OBJ_FUNC = Abs(YTEMP_VAL) ^ 2 / XTEMP_VAL * PUB_FACTOR_VAL 'Maximize
Exit Function
ERROR_LABEL:
GARCH_FITTING_OBJ_FUNC = 1 / PUB_EPSILON
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : GARCH_FITTING_CONST_FUNC
'DESCRIPTION   : Constraint Function of the Generalized Autoregressive
'Conditional Heteroscedasticity
'LIBRARY       : STATISTICS
'GROUP         : GARCH
'ID            : 011
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/08/2008
'************************************************************************************
'************************************************************************************

Private Function GARCH_FITTING_CONST_FUNC(ByRef PARAM_VECTOR As Variant, _
Optional ByVal VERSION As Variant = "")

Dim OMEGA_VAL As Double
Dim ALPHA_VAL As Double
Dim BETA_VAL As Double
Dim THETA_VAL As Double

On Error GoTo ERROR_LABEL

GARCH_FITTING_CONST_FUNC = True

If VERSION = "" Then: VERSION = PUB_GARCH_OBJ
'-----------------------------------------------------------------------------------
Select Case PUB_GARCH_OBJ
'-----------------------------------------------------------------------------------
Case 0
'-----------------------------------------------------------------------------------
    OMEGA_VAL = PARAM_VECTOR(1, 1)
    ALPHA_VAL = PARAM_VECTOR(2, 1)
    BETA_VAL = PARAM_VECTOR(3, 1)

    If OMEGA_VAL < 0 Or ALPHA_VAL < 0 Or BETA_VAL < 0 Then
        GARCH_FITTING_CONST_FUNC = False
        Exit Function
    End If
    If (ALPHA_VAL + BETA_VAL) > 1 Then 'Var condition
        GARCH_FITTING_CONST_FUNC = False
        Exit Function
    End If
    If (BETA_VAL ^ 2 + 2 * ALPHA_VAL * BETA_VAL + 3 * ALPHA_VAL) > 1 Then '0.99 Then
        GARCH_FITTING_CONST_FUNC = False 'conditional kurtosis
        Exit Function
    End If
'-----------------------------------------------------------------------------------
Case 1
'-----------------------------------------------------------------------------------
    OMEGA_VAL = PARAM_VECTOR(1, 1)
    ALPHA_VAL = PARAM_VECTOR(2, 1)
    BETA_VAL = PARAM_VECTOR(3, 1)

    If OMEGA_VAL < 0 Or ALPHA_VAL < 0 Or BETA_VAL < 0 Then
        GARCH_FITTING_CONST_FUNC = False
        Exit Function
    End If
    If (ALPHA_VAL + BETA_VAL) > 1 Then 'Var condition
        GARCH_FITTING_CONST_FUNC = False
        Exit Function
    End If
'-----------------------------------------------------------------------------------
Case Else 'Need to add more constraint here
'-----------------------------------------------------------------------------------
    OMEGA_VAL = PARAM_VECTOR(1, 1)
    ALPHA_VAL = PARAM_VECTOR(2, 1)
    BETA_VAL = PARAM_VECTOR(3, 1)
    THETA_VAL = PARAM_VECTOR(4, 1)

    If OMEGA_VAL < 0 Or ALPHA_VAL < 0 Or BETA_VAL < 0 Or THETA_VAL < 0 Then
        GARCH_FITTING_CONST_FUNC = False
        Exit Function
    End If
    If (ALPHA_VAL + BETA_VAL) > 1 Then
        GARCH_FITTING_CONST_FUNC = False
        Exit Function
    End If
'-----------------------------------------------------------------------------------
End Select
'-----------------------------------------------------------------------------------

Exit Function
ERROR_LABEL:
GARCH_FITTING_CONST_FUNC = False
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : GARCH_LOGLIKE1_FUNC
'DESCRIPTION   : Garch(1,1) Frame A - Loglikehood
'LIBRARY       : STATISTICS
'GROUP         : GARCH
'ID            : 012
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/08/2008
'************************************************************************************
'************************************************************************************

Function GARCH_LOGLIKE1_FUNC(ByRef DATA_RNG As Variant, _
ByVal OMEGA_VAL As Double, _
ByVal ALPHA_VAL As Double, _
ByVal BETA_VAL As Double, _
Optional ByVal DATA_TYPE As Integer = 0)

Dim i As Long
Dim NROWS As Long

Dim VAR_VAL As Double
Dim TEMP_SUM As Double
Dim MEAN_VAL As Double

Dim TEMP1_VAL As Double
Dim TEMP2_VAL As Double
Dim TEMP3_VAL As Double
Dim TEMP4_VAL As Double

Dim DATA_VECTOR As Variant

Dim PI_VAL As Double

On Error GoTo ERROR_LABEL

'One important departure from independency over time is when returns follow
'GARCH(1,1) processes: r(t) = e(t) * sig(t)^2, with e(t) = N(0,1) and
'sig(t)^2 = w + a*r(t-1)^2 + b*sig(t-1)^2, 0<w<inf, a>=0, b>=0, a+b<1 and
'sig(0)^2=w/(1-a-b). The key feature of a GARCH(1,1) process is time-varying
'conditional volatility when a and/or b are not equal to zero. There are many
'financial time series which can be described very well with GARCH(1,1) models.

'Note that volatility is underestimated if the return time series is subject
'to positive autocorrelation. Simple adjustments for autocorrelation exist.


PI_VAL = 3.14159265358979

DATA_VECTOR = DATA_RNG
If UBound(DATA_VECTOR, 1) = 1 Then: DATA_VECTOR = MATRIX_TRANSPOSE_FUNC(DATA_VECTOR)
If DATA_TYPE <> 0 Then: DATA_VECTOR = MATRIX_PERCENT_FUNC(DATA_VECTOR, 1)
NROWS = UBound(DATA_VECTOR, 1)

'----------------------------------------------------------------------------------
TEMP_SUM = 0
For i = 1 To NROWS: TEMP_SUM = TEMP_SUM + DATA_VECTOR(i, 1): Next i
MEAN_VAL = TEMP_SUM / NROWS
'----------------------------------------------------------------------------------
TEMP_SUM = 0
For i = 1 To NROWS: TEMP_SUM = TEMP_SUM + (DATA_VECTOR(i, 1) - MEAN_VAL) ^ 2: Next i
VAR_VAL = (TEMP_SUM / (NROWS - 1))
'----------------------------------------------------------------------------------

TEMP1_VAL = DATA_VECTOR(1, 1) - MEAN_VAL
TEMP2_VAL = VAR_VAL
TEMP_SUM = -Log(TEMP2_VAL) - (TEMP1_VAL) ^ 2 / TEMP2_VAL

For i = 2 To NROWS
    TEMP3_VAL = DATA_VECTOR(i, 1) - MEAN_VAL
    TEMP4_VAL = OMEGA_VAL + ALPHA_VAL * (TEMP1_VAL) ^ 2 + BETA_VAL * TEMP2_VAL
    TEMP_SUM = TEMP_SUM + -Log(TEMP4_VAL) - (TEMP3_VAL) ^ 2 / TEMP4_VAL
    TEMP1_VAL = TEMP3_VAL
    TEMP2_VAL = TEMP4_VAL
Next i
    
GARCH_LOGLIKE1_FUNC = 0.5 * (-Log(2 * PI_VAL) * (NROWS) + TEMP_SUM)

Exit Function
ERROR_LABEL:
GARCH_LOGLIKE1_FUNC = Err.number '-PUB_EPSILON
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : GARCH_LOGLIKE2_FUNC
'DESCRIPTION   : Garch(1,1) Frame B - Loglikehood
'LIBRARY       : STATISTICS
'GROUP         : GARCH
'ID            : 013
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/08/2008
'************************************************************************************
'************************************************************************************

Function GARCH_LOGLIKE2_FUNC(ByRef DATA_RNG As Variant, _
ByVal OMEGA_VAL As Double, _
ByVal ALPHA_VAL As Double, _
ByVal BETA_VAL As Double, _
Optional ByVal DATA_TYPE As Integer = 0)

Dim i As Long
Dim NROWS As Long

Dim PI_VAL As Double

Dim TEMP_SUM As Double
Dim FACTOR_VAL As Double
Dim MEAN_VAL As Double 'Unconditional Standard Deviation
Dim VAR_VAL As Double 'Sample Unconditional Variance
Dim SIGMA_VAL As Double

Dim DATA_VECTOR As Variant

On Error GoTo ERROR_LABEL

DATA_VECTOR = DATA_RNG
If UBound(DATA_VECTOR, 1) = 1 Then: DATA_VECTOR = MATRIX_TRANSPOSE_FUNC(DATA_VECTOR)
If DATA_TYPE <> 0 Then: DATA_VECTOR = MATRIX_PERCENT_FUNC(DATA_VECTOR, 1)
NROWS = UBound(DATA_VECTOR, 1)

PI_VAL = 3.14159265358979

'----------------------------------------------------------------------------------
TEMP_SUM = 0
For i = 1 To NROWS: TEMP_SUM = TEMP_SUM + DATA_VECTOR(i, 1): Next i
MEAN_VAL = TEMP_SUM / NROWS
TEMP_SUM = 0
For i = 1 To NROWS: TEMP_SUM = TEMP_SUM + (DATA_VECTOR(i, 1) - MEAN_VAL) ^ 2: Next i
SIGMA_VAL = (TEMP_SUM / (NROWS - 1)) ^ 0.5
'----------------------------------------------------------------------------------

VAR_VAL = SIGMA_VAL ^ 2

TEMP_SUM = 0
FACTOR_VAL = VAR_VAL 'Starting point can be either
'unconditional variance or zero or...

For i = 2 To NROWS
    FACTOR_VAL = OMEGA_VAL + ALPHA_VAL * (DATA_VECTOR(i - 1, 1) - MEAN_VAL) ^ 2 + BETA_VAL * FACTOR_VAL
    TEMP_SUM = Log((1 / Sqr(2 * PI_VAL * FACTOR_VAL)) * Exp(-0.5 * (DATA_VECTOR(i, 1) - MEAN_VAL) ^ 2 / FACTOR_VAL)) + TEMP_SUM
Next i

GARCH_LOGLIKE2_FUNC = TEMP_SUM

Exit Function
ERROR_LABEL:
GARCH_LOGLIKE2_FUNC = Err.number '-PUB_EPSILON
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : GARCH_LOGLIKE3_FUNC
'DESCRIPTION   : Garch(1,1) Frame C - Loglikehood; employing variance targeting
'techniques
'LIBRARY       : STATISTICS
'GROUP         : GARCH
'ID            : 014
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/08/2008
'************************************************************************************
'************************************************************************************

Function GARCH_LOGLIKE3_FUNC(ByRef DATA_RNG As Variant, _
ByVal OMEGA_VAL As Double, _
ByVal ALPHA_VAL As Double, _
ByVal BETA_VAL As Double, _
ByVal THETA_VAL As Double, _
Optional ByVal DATA_TYPE As Integer = 0)

Dim i As Long
Dim NROWS As Long

Dim PI_VAL As Double

Dim TEMP_SUM As Double
Dim MEAN_VAL As Double
Dim SIGMA_VAL As sDouble

Dim TEMP1_VAL As Double
Dim TEMP2_VAL As Double

Dim DATA_VECTOR As Variant

On Error GoTo ERROR_LABEL

PI_VAL = 3.14159265358979

DATA_VECTOR = DATA_RNG
If UBound(DATA_VECTOR, 1) = 1 Then: DATA_VECTOR = MATRIX_TRANSPOSE_FUNC(DATA_VECTOR)
If DATA_TYPE <> 0 Then: DATA_VECTOR = MATRIX_PERCENT_FUNC(DATA_VECTOR, 1)
NROWS = UBound(DATA_VECTOR, 1)

'----------------------------------------------------------------------------------
TEMP_SUM = 0
For i = 1 To NROWS: TEMP_SUM = TEMP_SUM + DATA_VECTOR(i, 1): Next i
MEAN_VAL = TEMP_SUM / NROWS
TEMP_SUM = 0
For i = 1 To NROWS: TEMP_SUM = TEMP_SUM + (DATA_VECTOR(i, 1) - MEAN_VAL) ^ 2: Next i
SIGMA_VAL = (TEMP_SUM / (NROWS - 1)) ^ 0.5
'----------------------------------------------------------------------------------
TEMP1_VAL = (SIGMA_VAL) ^ 2
TEMP2_VAL = (-0.5 * Log(2 * PI_VAL) - 0.5 * Log(TEMP1_VAL) - 0.5 * (DATA_VECTOR(1, 1) ^ 2 / TEMP1_VAL)) 'MLE
'----------------------------------------------------------------------------------
TEMP_SUM = TEMP2_VAL
For i = 2 To NROWS
    TEMP2_VAL = OMEGA_VAL + BETA_VAL * TEMP1_VAL + ALPHA_VAL * ((DATA_VECTOR(i - 1, 1) - THETA_VAL * (TEMP1_VAL ^ 0.5)) ^ 2)
    TEMP_SUM = TEMP_SUM + (-0.5 * Log(2 * PI_VAL) - 0.5 * Log(TEMP2_VAL) - 0.5 * (DATA_VECTOR(i, 1) ^ 2 / TEMP2_VAL)) 'MLE
    TEMP1_VAL = TEMP2_VAL
Next i
'----------------------------------------------------------------------------------
GARCH_LOGLIKE3_FUNC = TEMP_SUM
'----------------------------------------------------------------------------------

Exit Function
ERROR_LABEL:
GARCH_LOGLIKE3_FUNC = Err.number ' -PUB_EPSILON
End Function


