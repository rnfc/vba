Attribute VB_Name = "FINAN_FI_HW_2F_LIBR"


Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.

Private PUB_CAP_VAL As Double
Private PUB_NOM_VAL As Double
Private PUB_EXPIRAT_VAL As Double
Private PUB_INIT_STEP_VAL As Double

Private PUB_FIX_ARR As Variant
Private PUB_START_ARR As Variant
Private PUB_END_ARR As Variant
Private PUB_ACCR_ARR As Variant
Private PUB_FLOOR_ARR As Variant
Private PUB_FORW_ARR As Variant

Private PUB_TENOR_ARR As Variant
Private PUB_RATE_ARR As Variant

Private PUB_RF_TENOR_ARR As Variant
Private PUB_RF_RATE_ARR As Variant
Private PUB_RF_DISC_ARR As Variant

Private PUB_MEAN_VAL As Double
Private PUB_SIGMA_VAL As Double
Private PUB_TARGET_VAL As Double

Private PUB_START_TENOR_VAL As Double
Private PUB_FIXED_RATE_VAL As Double

Private PUB_LOWER_BOUND_VAL As Double
Private PUB_UPPER_BOUND_VAL As Double

Private PUB_GUESS_VAL As Double
Private PUB_nLOOPS_VAL As Double
Private PUB_TOL_VAL As Double
Private PUB_EPS_VAL As Double

Private PUB_BK_MIN As Double
Private PUB_BK_DELTA As Double
Private PUB_BK_PRICE As Double
Private PUB_BK_GRAD As Double
Private PUB_BK_SIZE As Double
Private PUB_BK_ARR As Variant

Private PUB_VERSION_VAL As Integer

'************************************************************************************
'************************************************************************************



'************************************************************************************
'************************************************************************************
'FUNCTION      : HW_BK_CAP_SIGMA_FUNC

'DESCRIPTION   : Implementations of the 2 Factor HW model (mathematically equivalent
'to the G2++ model)

'LIBRARY       : FIXED_INCOME
'GROUP         : HW_2F
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 10-06-2008
'************************************************************************************
'************************************************************************************

Function HW_BK_CAP_SIGMA_FUNC(ByRef FORWARD_TENOR_RNG As Variant, _
ByRef FORWARD_RATE_RNG As Variant, _
ByRef CAP_TENOR_RNG As Variant, _
ByRef PARAM_RNG As Variant, _
Optional ByVal FIXED_RATE As Double = 0.04, _
Optional ByVal NOMINAL As Double = 1, _
Optional ByVal VERSION As Integer = 1)

Dim i As Long
Dim j As Long
Dim NSIZE As Long
    
Dim XDATA_VECTOR As Variant
Dim YDATA_VECTOR As Variant

Dim PARAM_VECTOR As Variant

Dim FORWARD_RATE_ARR As Variant
Dim FORWARD_TENOR_ARR As Variant

Dim INIT_STEP_VAL As Double
Dim BRENT_LOWER_BOUND As Double
Dim BRENT_UPPER_BOUND As Double
Dim BRENT_GUESS_SIGMA As Double

Dim ZERO_RATE_ARR As Variant
Dim START_TENOR_VAL As Double

Dim nLOOPS As Long
Dim epsilon As Double
Dim tolerance As Double

'---------------------------------------------------------------------------------
On Error GoTo ERROR_LABEL
Call HW_BK_RESET_VAR_FUNC
'---------------------------------------------------------------------------------
'---------------------------------------------------------------------------------
BRENT_LOWER_BOUND = 0.001
BRENT_UPPER_BOUND = 5
BRENT_GUESS_SIGMA = 0.05
'---------------------------------------------------------------------------------
nLOOPS = 100
epsilon = 2 ^ 52
tolerance = 0.00000001
'---------------------------------------------------------------------------------
INIT_STEP_VAL = 0
START_TENOR_VAL = 0.5
'---------------------------------------------------------------------------------
PUB_LOWER_BOUND_VAL = BRENT_LOWER_BOUND
PUB_UPPER_BOUND_VAL = BRENT_UPPER_BOUND
PUB_INIT_STEP_VAL = INIT_STEP_VAL
PUB_GUESS_VAL = BRENT_GUESS_SIGMA
PUB_nLOOPS_VAL = nLOOPS
PUB_TOL_VAL = tolerance
PUB_EPS_VAL = epsilon
'---------------------------------------------------------------------------------

XDATA_VECTOR = FORWARD_TENOR_RNG
If UBound(XDATA_VECTOR, 1) = 1 Then: XDATA_VECTOR = MATRIX_TRANSPOSE_FUNC(XDATA_VECTOR)
YDATA_VECTOR = FORWARD_RATE_RNG
If UBound(YDATA_VECTOR, 1) = 1 Then: YDATA_VECTOR = MATRIX_TRANSPOSE_FUNC(YDATA_VECTOR)

If UBound(XDATA_VECTOR, 1) <> UBound(YDATA_VECTOR, 1) Then: GoTo ERROR_LABEL
NSIZE = UBound(XDATA_VECTOR, 1)
ReDim FORWARD_TENOR_ARR(0 To NSIZE - 1)
ReDim FORWARD_RATE_ARR(0 To NSIZE - 1)

j = 0
For i = 1 To NSIZE
  FORWARD_TENOR_ARR(j) = XDATA_VECTOR(i, 1)
  FORWARD_RATE_ARR(j) = YDATA_VECTOR(i, 1)
  j = j + 1
Next i

'-------------------------------------------------------------------------------
'---------------------------SET TERM STRUCTURE----------------------------------
'-------------------------------------------------------------------------------
PUB_RF_TENOR_ARR = FORWARD_TENOR_ARR
PUB_RF_RATE_ARR = FORWARD_RATE_ARR

ReDim PUB_RF_DISC_ARR(0 To NSIZE - 1)
ReDim ZERO_RATE_ARR(0 To NSIZE - 1)

PUB_RF_DISC_ARR(0) = 1
ZERO_RATE_ARR(0) = FORWARD_RATE_ARR(0)
For i = 1 To NSIZE - 1
  ZERO_RATE_ARR(i) = (FORWARD_RATE_ARR(i) * (FORWARD_TENOR_ARR(i) - _
                      FORWARD_TENOR_ARR(i - 1)) + ZERO_RATE_ARR(i - 1) * _
                      FORWARD_TENOR_ARR(i - 1)) / FORWARD_TENOR_ARR(i)
  PUB_RF_DISC_ARR(i) = Exp(-ZERO_RATE_ARR(i) * PUB_RF_TENOR_ARR(i))
Next i
'-------------------------------------------------------------------------------
'-------------------------------------------------------------------------------
'-------------------------------------------------------------------------------
XDATA_VECTOR = CAP_TENOR_RNG
If UBound(XDATA_VECTOR, 1) = 1 Then: XDATA_VECTOR = MATRIX_TRANSPOSE_FUNC(XDATA_VECTOR)
NSIZE = UBound(XDATA_VECTOR, 1)
  
'-------------------------------------------------------------------------------
'-------------------------------------------------------------------------------
'-------------------------------------------------------------------------------

PARAM_VECTOR = PARAM_RNG
If UBound(PARAM_VECTOR, 1) = 1 Then: PARAM_VECTOR = MATRIX_TRANSPOSE_FUNC(PARAM_VECTOR)
'------------------------------------------------------------------------------------
PUB_TENOR_ARR = PUB_RF_TENOR_ARR
PUB_START_TENOR_VAL = START_TENOR_VAL
PUB_NOM_VAL = NOMINAL 'Nominal Value
PUB_FIXED_RATE_VAL = FIXED_RATE 'Flat Rate
'------------------------------------------------------------------------------------

PUB_VERSION_VAL = VERSION
YDATA_VECTOR = HW_BK_SIGMA_OBJ_FUNC(XDATA_VECTOR, PARAM_VECTOR)
HW_BK_CAP_SIGMA_FUNC = YDATA_VECTOR

Exit Function
ERROR_LABEL:
HW_BK_CAP_SIGMA_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************



'************************************************************************************
'************************************************************************************
'FUNCTION      : HW_BK_CAP_SIGMA_OPTIM_FUNC

'DESCRIPTION   : These functions contain implementation of Hull White/Black
'Karasinki model calibration using Levenberg Marquardt optimization.

'IF VERSION = 0 Then: calibrating Black Karasinki Term Structure
'on a Trinomial tree using Levenberg Marquardt optimization
  
'IF VERSION = 1 Then: calibrating Hull White Term Structure
'on a Trinomial tree using Levenberg Marquardt optimization

'LIBRARY       : FIXED_INCOME
'GROUP         : HW_2F
'ID            : 002
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 10-06-2008
'************************************************************************************
'************************************************************************************
 
Function HW_BK_CAP_SIGMA_OPTIM_FUNC(ByRef FORWARD_TENOR_RNG As Variant, _
ByRef FORWARD_RATE_RNG As Variant, _
ByRef CAP_TENOR_RNG As Variant, _
ByRef CAP_SIGMA_RNG As Variant, _
Optional ByVal FIXED_RATE As Double = 0.04, _
Optional ByVal NOMINAL As Double = 1, _
Optional ByVal INIT_MEAN_VAL As Double = 0.05, _
Optional ByVal INIT_VOLAT_VAL As Double = 0.001, _
Optional ByVal VERSION As Integer = 1, _
Optional ByVal OUTPUT As Integer = 1)

Dim i As Long
Dim j As Long
Dim NSIZE As Long
    
Dim XDATA_VECTOR As Variant
Dim YDATA_VECTOR As Variant

Dim TEMP_GROUP As Variant

Dim SCALE_VECTOR As Variant
Dim PARAM_VECTOR As Variant
Dim EPSILON_VECTOR As Variant

Dim FORWARD_RATE_ARR As Variant
Dim FORWARD_TENOR_ARR As Variant

Dim LOWER_BOUND_VECTOR As Variant
Dim UPPER_BOUND_VECTOR As Variant

Dim OPTIONAL_VECTOR As Variant

Dim INIT_STEP_VAL As Double
Dim BRENT_LOWER_BOUND As Double
Dim BRENT_UPPER_BOUND As Double
Dim BRENT_GUESS_SIGMA As Double

Dim ZERO_RATE_ARR As Variant
Dim START_TENOR_VAL As Double

Dim LEVENBERG_MARQUARDT_EPS_VAL As Double
Dim LEVENBERG_MARQUARDT_LOOPS_VAL As Double
Dim LEVENBERG_MARQUARDT_LOWER_BOUND As Double
Dim LEVENBERG_MARQUARDT_UPPER_BOUND As Double

Dim nLOOPS As Long
Dim epsilon As Double
Dim tolerance As Double

'---------------------------------------------------------------------------------
On Error GoTo ERROR_LABEL
Call HW_BK_RESET_VAR_FUNC
'---------------------------------------------------------------------------------
LEVENBERG_MARQUARDT_EPS_VAL = 0.001
LEVENBERG_MARQUARDT_LOWER_BOUND = 0.01
LEVENBERG_MARQUARDT_UPPER_BOUND = 0.8
LEVENBERG_MARQUARDT_LOOPS_VAL = 50
'---------------------------------------------------------------------------------
BRENT_LOWER_BOUND = 0.001
BRENT_UPPER_BOUND = 5
BRENT_GUESS_SIGMA = 0.05
'---------------------------------------------------------------------------------
nLOOPS = 100
epsilon = 2 ^ 52
tolerance = 0.00000001
'---------------------------------------------------------------------------------
INIT_STEP_VAL = 0
START_TENOR_VAL = 0.5
'---------------------------------------------------------------------------------
PUB_LOWER_BOUND_VAL = BRENT_LOWER_BOUND
PUB_UPPER_BOUND_VAL = BRENT_UPPER_BOUND
PUB_INIT_STEP_VAL = INIT_STEP_VAL
PUB_GUESS_VAL = BRENT_GUESS_SIGMA
PUB_nLOOPS_VAL = nLOOPS
PUB_TOL_VAL = tolerance
PUB_EPS_VAL = epsilon
'---------------------------------------------------------------------------------

XDATA_VECTOR = FORWARD_TENOR_RNG
If UBound(XDATA_VECTOR, 1) = 1 Then: XDATA_VECTOR = MATRIX_TRANSPOSE_FUNC(XDATA_VECTOR)
YDATA_VECTOR = FORWARD_RATE_RNG
If UBound(YDATA_VECTOR, 1) = 1 Then: YDATA_VECTOR = MATRIX_TRANSPOSE_FUNC(YDATA_VECTOR)

If UBound(XDATA_VECTOR, 1) <> UBound(YDATA_VECTOR, 1) Then: GoTo ERROR_LABEL
NSIZE = UBound(XDATA_VECTOR, 1)
ReDim FORWARD_TENOR_ARR(0 To NSIZE - 1)
ReDim FORWARD_RATE_ARR(0 To NSIZE - 1)

j = 0
For i = 1 To NSIZE
  FORWARD_TENOR_ARR(j) = XDATA_VECTOR(i, 1)
  FORWARD_RATE_ARR(j) = YDATA_VECTOR(i, 1)
  j = j + 1
Next i

'-------------------------------------------------------------------------------
'---------------------------SET TERM STRUCTURE----------------------------------
'-------------------------------------------------------------------------------
PUB_RF_TENOR_ARR = FORWARD_TENOR_ARR
PUB_RF_RATE_ARR = FORWARD_RATE_ARR

ReDim PUB_RF_DISC_ARR(0 To NSIZE - 1)
ReDim ZERO_RATE_ARR(0 To NSIZE - 1)

PUB_RF_DISC_ARR(0) = 1
ZERO_RATE_ARR(0) = FORWARD_RATE_ARR(0)
For i = 1 To NSIZE - 1
  ZERO_RATE_ARR(i) = (FORWARD_RATE_ARR(i) * (FORWARD_TENOR_ARR(i) - _
                      FORWARD_TENOR_ARR(i - 1)) + ZERO_RATE_ARR(i - 1) * _
                      FORWARD_TENOR_ARR(i - 1)) / FORWARD_TENOR_ARR(i)
  PUB_RF_DISC_ARR(i) = Exp(-ZERO_RATE_ARR(i) * PUB_RF_TENOR_ARR(i))
Next i
'-------------------------------------------------------------------------------
'-------------------------------------------------------------------------------
'-------------------------------------------------------------------------------
XDATA_VECTOR = CAP_TENOR_RNG
If UBound(XDATA_VECTOR, 1) = 1 Then: XDATA_VECTOR = MATRIX_TRANSPOSE_FUNC(XDATA_VECTOR)
YDATA_VECTOR = CAP_SIGMA_RNG
If UBound(YDATA_VECTOR, 1) = 1 Then: YDATA_VECTOR = MATRIX_TRANSPOSE_FUNC(YDATA_VECTOR)

If UBound(XDATA_VECTOR, 1) <> UBound(YDATA_VECTOR, 1) Then: GoTo ERROR_LABEL
NSIZE = UBound(XDATA_VECTOR, 1)
  
'-------------------------------------------------------------------------------
'-------------------------------------------------------------------------------
'-------------------------------------------------------------------------------

ReDim SCALE_VECTOR(1 To NSIZE, 1 To 1)
For i = 1 To NSIZE
  SCALE_VECTOR(i, 1) = 1
Next i

ReDim EPSILON_VECTOR(1 To 2, 1 To 1)
EPSILON_VECTOR(1, 1) = LEVENBERG_MARQUARDT_EPS_VAL
EPSILON_VECTOR(2, 1) = LEVENBERG_MARQUARDT_EPS_VAL

ReDim PARAM_VECTOR(1 To 2, 1 To 1)

PARAM_VECTOR(1, 1) = INIT_MEAN_VAL
PARAM_VECTOR(2, 1) = INIT_VOLAT_VAL

ReDim LOWER_BOUND_VECTOR(1 To 2, 1 To 1)
LOWER_BOUND_VECTOR(1, 1) = LEVENBERG_MARQUARDT_LOWER_BOUND
LOWER_BOUND_VECTOR(2, 1) = LEVENBERG_MARQUARDT_LOWER_BOUND
   
ReDim UPPER_BOUND_VECTOR(1 To 2, 1 To 1)
UPPER_BOUND_VECTOR(1, 1) = LEVENBERG_MARQUARDT_UPPER_BOUND
UPPER_BOUND_VECTOR(2, 1) = LEVENBERG_MARQUARDT_UPPER_BOUND
 
ReDim OPTIONAL_VECTOR(1 To 2, 1 To 2)

For i = 1 To 2 'Set Matrix Column
    OPTIONAL_VECTOR(i, 1) = LOWER_BOUND_VECTOR(i, 1)
Next i

For i = 1 To 2 'Set Matrix Column
    OPTIONAL_VECTOR(i, 2) = UPPER_BOUND_VECTOR(i, 1)
Next i
  
'------------------------------------------------------------------------------------
PUB_TENOR_ARR = PUB_RF_TENOR_ARR
PUB_START_TENOR_VAL = START_TENOR_VAL
PUB_NOM_VAL = NOMINAL 'Nominal Value
PUB_FIXED_RATE_VAL = FIXED_RATE 'Flat Rate
'------------------------------------------------------------------------------------

PUB_VERSION_VAL = VERSION

TEMP_GROUP = LEVENBERG_MARQUARDT_OPTIMIZATION4_FUNC("HW_BK_SIGMA_OBJ_FUNC", XDATA_VECTOR, _
              YDATA_VECTOR, PARAM_VECTOR, SCALE_VECTOR, _
              OPTIONAL_VECTOR, EPSILON_VECTOR, LEVENBERG_MARQUARDT_LOOPS_VAL, _
              PUB_TOL_VAL, PUB_EPS_VAL)

'------------------------------------------------------------------------------------
'------------Hull white model parameters for risk free rate process------------------
'------------------------------------------------------------------------------------
If OUTPUT = 0 Then
    HW_BK_CAP_SIGMA_OPTIM_FUNC = TEMP_GROUP(1)
    Exit Function
ElseIf OUTPUT = 1 Then
    HW_BK_CAP_SIGMA_OPTIM_FUNC = TEMP_GROUP(2)
    Exit Function
ElseIf OUTPUT = 2 Then
    HW_BK_CAP_SIGMA_OPTIM_FUNC = TEMP_GROUP
    Exit Function
End If
'------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------

Exit Function
ERROR_LABEL:
HW_BK_CAP_SIGMA_OPTIM_FUNC = Err.number
End Function

'// PERFECT

'************************************************************************************
'************************************************************************************



'************************************************************************************
'************************************************************************************
'FUNCTION      : EUROPEAN_EXCHANGE_OPTION_FUNC

'DESCRIPTION   : Calculates European exchange option for risky bond vs. Risk
'free bond Exchange option price is calculated using 2 factor Trinomial
'tree for default intensity.

'2D tree is based on Schonbucher's book "Credit derivative pricing models"
'section 7.4. (http://www.math.ethz.ch/~schonbuc/)

'1D tree is based on method as described in appendix for Brigo's
'"Interest rate models" (http://www.damianobrigo.it/book.html)

'Following inputs are needed:

'1.Term structure of risk free rates as piecewise flat forward indexed
'  by 6 months (0.5)

'2.Market Cap volatilties indexed by 6 months

'3.Term structure of piecewise flat forward indexed by 6 months which is
'  obtained for risky bonds

'4.Model parameters for default risk hull white process.

'5.Correlation of default risk vs. risk free rate.

'6.Exchange option data

'LIBRARY       : FIXED_INCOME
'GROUP         : HW_2F
'ID            : 003
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 10-06-2008
'************************************************************************************
'************************************************************************************

Function EUROPEAN_EXCHANGE_OPTION_FUNC(ByRef FORWARD_TENOR_RNG As Variant, _
ByRef FORWARD_RATE_RNG As Variant, _
ByRef CAP_TENOR_RNG As Variant, _
ByRef CAP_SIGMA_RNG As Variant, _
ByRef RISKY_FORWARD_TENOR_RNG As Variant, _
ByRef RISKY_FORWARD_RATE_RNG As Variant, _
ByRef POS_ADJ_RHO_RNG As Variant, _
ByRef NEG_ADJ_RHO_RNG As Variant, _
ByVal DRIFT_VAL As Double, _
ByVal SIGMA_VAL As Double, _
ByVal RHO_VAL As Double, _
ByVal EXCH_EXPIRAT_VAL As Double, _
ByVal EXCH_RATIO_VAL As Double, _
ByVal EXCH_NOMIN_VAL As Double, _
Optional ByRef PARAM_RNG As Variant, _
Optional ByVal FIXED_RATE As Double = 0.04, _
Optional ByVal NOMINAL As Double = 1, _
Optional ByVal INIT_MEAN_VAL As Double = 0.05, _
Optional ByVal INIT_VOLAT_VAL As Double = 0.001, _
Optional ByVal RHO_FACTOR As Double = 36, _
Optional ByVal OUTPUT As Integer = 0)
  
Dim i As Long
Dim j As Long
Dim k As Long

Dim ii As Long
Dim jj As Long
Dim kk As Long

Dim NSIZE As Long
Dim nSTEPS As Long
Dim nLOOPS As Long

Dim FORWARD_RATE_ARR As Variant
Dim FORWARD_TENOR_ARR As Variant
  
Dim XDATA_VECTOR As Variant
Dim YDATA_VECTOR As Variant

Dim SCALE_VECTOR As Variant
Dim PARAM_VECTOR As Variant
Dim EPSILON_VECTOR As Variant

Dim LOWER_BOUND_VECTOR As Variant
Dim UPPER_BOUND_VECTOR As Variant

Dim TEMP_SUM As Double
Dim TEMP_GROUP As Variant
Dim OPTIONAL_VECTOR As Variant

Dim ATEMP_ARR As Variant
Dim BTEMP_ARR As Variant
Dim CTEMP_ARR As Variant
Dim DTEMP_ARR As Variant
  
Dim FIRST_GROUP As Variant
Dim FIRST_FIT_ARR As Variant
Dim FIRST_PRICE_ARR As Variant

Dim SECOND_GROUP As Variant
Dim SECOND_FIT_ARR As Variant
Dim SECOND_PRICE_ARR As Variant

Dim INIT_STEP_VAL As Double
Dim START_TENOR_VAL As Double

Dim RF_PRICE_ARR As Variant 'State Prices Risk Free
Dim RISKY_TENOR_ARR As Variant
Dim RISKY_RATE_ARR As Variant
Dim RISKY_PRICE_ARR As Variant

Dim POS_ADJ_RHO_ARR As Variant
Dim NEG_ADJ_RHO_ARR As Variant

Dim ZERO_RATE_ARR As Variant

Dim LEVENBERG_MARQUARDT_EPS_VAL As Double
Dim LEVENBERG_MARQUARDT_LOOPS_VAL As Double
Dim LEVENBERG_MARQUARDT_LOWER_BOUND As Double
Dim LEVENBERG_MARQUARDT_UPPER_BOUND As Double

Dim BRENT_LOWER_BOUND As Double
Dim BRENT_UPPER_BOUND As Double
Dim BRENT_GUESS_SIGMA As Double

Dim epsilon As Double
Dim tolerance As Double

'---------------------------------------------------------------------------------
On Error GoTo ERROR_LABEL
Call HW_BK_RESET_VAR_FUNC
'---------------------------------------------------------------------------------
LEVENBERG_MARQUARDT_EPS_VAL = 0.001
LEVENBERG_MARQUARDT_LOWER_BOUND = 0.01
LEVENBERG_MARQUARDT_UPPER_BOUND = 0.8
LEVENBERG_MARQUARDT_LOOPS_VAL = 50
'---------------------------------------------------------------------------------
BRENT_LOWER_BOUND = 0.001
BRENT_UPPER_BOUND = 5
BRENT_GUESS_SIGMA = 0.05
'---------------------------------------------------------------------------------
nLOOPS = 100
epsilon = 2 ^ 52
tolerance = 0.00000001
'---------------------------------------------------------------------------------
INIT_STEP_VAL = 0
kk = 0
START_TENOR_VAL = 0.5
'---------------------------------------------------------------------------------
PUB_LOWER_BOUND_VAL = BRENT_LOWER_BOUND
PUB_UPPER_BOUND_VAL = BRENT_UPPER_BOUND
PUB_INIT_STEP_VAL = INIT_STEP_VAL
PUB_GUESS_VAL = BRENT_GUESS_SIGMA
PUB_nLOOPS_VAL = nLOOPS
PUB_TOL_VAL = tolerance
PUB_EPS_VAL = epsilon
'---------------------------------------------------------------------------------

XDATA_VECTOR = FORWARD_TENOR_RNG
If UBound(XDATA_VECTOR, 1) = 1 Then: XDATA_VECTOR = MATRIX_TRANSPOSE_FUNC(XDATA_VECTOR)
YDATA_VECTOR = FORWARD_RATE_RNG
If UBound(YDATA_VECTOR, 1) = 1 Then: YDATA_VECTOR = MATRIX_TRANSPOSE_FUNC(YDATA_VECTOR)

If UBound(XDATA_VECTOR, 1) <> UBound(YDATA_VECTOR, 1) Then: GoTo ERROR_LABEL
NSIZE = UBound(XDATA_VECTOR, 1)
ReDim FORWARD_TENOR_ARR(0 To NSIZE - 1)
ReDim FORWARD_RATE_ARR(0 To NSIZE - 1)

j = 0
For i = 1 To NSIZE
  FORWARD_TENOR_ARR(j) = XDATA_VECTOR(i, 1)
  FORWARD_RATE_ARR(j) = YDATA_VECTOR(i, 1)
  j = j + 1
Next i

'---------------------------SET TERM STRUCTURE----------------------------------
PUB_RF_TENOR_ARR = FORWARD_TENOR_ARR
PUB_RF_RATE_ARR = FORWARD_RATE_ARR

ReDim PUB_RF_DISC_ARR(0 To NSIZE - 1)
ReDim ZERO_RATE_ARR(0 To NSIZE - 1)

PUB_RF_DISC_ARR(0) = 1
ZERO_RATE_ARR(0) = FORWARD_RATE_ARR(0)
For i = 1 To NSIZE - 1
  ZERO_RATE_ARR(i) = (FORWARD_RATE_ARR(i) * (FORWARD_TENOR_ARR(i) - _
                      FORWARD_TENOR_ARR(i - 1)) + ZERO_RATE_ARR(i - 1) * _
                      FORWARD_TENOR_ARR(i - 1)) / FORWARD_TENOR_ARR(i)
  PUB_RF_DISC_ARR(i) = Exp(-ZERO_RATE_ARR(i) * PUB_RF_TENOR_ARR(i))
Next i
'---------------------------------------------------------------------------------
XDATA_VECTOR = RISKY_FORWARD_TENOR_RNG
If UBound(XDATA_VECTOR, 1) = 1 Then: XDATA_VECTOR = MATRIX_TRANSPOSE_FUNC(XDATA_VECTOR)
YDATA_VECTOR = RISKY_FORWARD_RATE_RNG
If UBound(YDATA_VECTOR, 1) = 1 Then: YDATA_VECTOR = MATRIX_TRANSPOSE_FUNC(YDATA_VECTOR)

If UBound(XDATA_VECTOR, 1) <> UBound(YDATA_VECTOR, 1) Then: GoTo ERROR_LABEL
NSIZE = UBound(XDATA_VECTOR, 1)
ReDim FORWARD_TENOR_ARR(0 To NSIZE - 1)
ReDim FORWARD_RATE_ARR(0 To NSIZE - 1)

j = 0
For i = 1 To NSIZE
  FORWARD_TENOR_ARR(j) = XDATA_VECTOR(i, 1)
  FORWARD_RATE_ARR(j) = YDATA_VECTOR(i, 1)
  j = j + 1
Next i
RISKY_TENOR_ARR = FORWARD_TENOR_ARR
RISKY_RATE_ARR = FORWARD_RATE_ARR
PUB_FORW_ARR = RISKY_RATE_ARR
'-------------------------------------------------------------------------------
XDATA_VECTOR = CAP_TENOR_RNG
If UBound(XDATA_VECTOR, 1) = 1 Then: XDATA_VECTOR = MATRIX_TRANSPOSE_FUNC(XDATA_VECTOR)
YDATA_VECTOR = CAP_SIGMA_RNG
If UBound(YDATA_VECTOR, 1) = 1 Then: YDATA_VECTOR = MATRIX_TRANSPOSE_FUNC(YDATA_VECTOR)

If UBound(XDATA_VECTOR, 1) <> UBound(YDATA_VECTOR, 1) Then: GoTo ERROR_LABEL
NSIZE = UBound(XDATA_VECTOR, 1)
 
'-------------------------------------------------------------------------------

ReDim SCALE_VECTOR(1 To NSIZE, 1 To 1)
For i = 1 To NSIZE
  SCALE_VECTOR(i, 1) = 1
Next i

ReDim EPSILON_VECTOR(1 To 2, 1 To 1)
EPSILON_VECTOR(1, 1) = LEVENBERG_MARQUARDT_EPS_VAL
EPSILON_VECTOR(2, 1) = LEVENBERG_MARQUARDT_EPS_VAL

ReDim PARAM_VECTOR(1 To 2, 1 To 1)

PARAM_VECTOR(1, 1) = INIT_MEAN_VAL
PARAM_VECTOR(2, 1) = INIT_VOLAT_VAL

ReDim LOWER_BOUND_VECTOR(1 To 2, 1 To 1)
LOWER_BOUND_VECTOR(1, 1) = LEVENBERG_MARQUARDT_LOWER_BOUND
LOWER_BOUND_VECTOR(2, 1) = LEVENBERG_MARQUARDT_LOWER_BOUND
   
ReDim UPPER_BOUND_VECTOR(1 To 2, 1 To 1)
UPPER_BOUND_VECTOR(1, 1) = LEVENBERG_MARQUARDT_UPPER_BOUND
UPPER_BOUND_VECTOR(2, 1) = LEVENBERG_MARQUARDT_UPPER_BOUND
 
ReDim OPTIONAL_VECTOR(1 To 2, 1 To 2)

For i = 1 To 2 'SetMatrixColumn
  OPTIONAL_VECTOR(i, 1) = LOWER_BOUND_VECTOR(i, 1)
Next i

For i = 1 To 2 'SetMatrixColumn
  OPTIONAL_VECTOR(i, 2) = UPPER_BOUND_VECTOR(i, 1)
Next i
  
'------------------------------------------------------------------------------------
PUB_TENOR_ARR = PUB_RF_TENOR_ARR
PUB_START_TENOR_VAL = START_TENOR_VAL
PUB_NOM_VAL = NOMINAL 'Nominal Value
PUB_FIXED_RATE_VAL = FIXED_RATE 'Flat Rate

'------------------------------------------------------------------------------------
If IsArray(PARAM_RNG) = False Then
'------------------------------------------------------------------------------------
    TEMP_GROUP = LEVENBERG_MARQUARDT_OPTIMIZATION4_FUNC("HW_BK_SIGMA_OBJ_FUNC", XDATA_VECTOR, _
              YDATA_VECTOR, PARAM_VECTOR, SCALE_VECTOR, _
              OPTIONAL_VECTOR, EPSILON_VECTOR, LEVENBERG_MARQUARDT_LOOPS_VAL, _
              PUB_TOL_VAL, PUB_EPS_VAL)
    If OUTPUT = 1 Then
        EUROPEAN_EXCHANGE_OPTION_FUNC = TEMP_GROUP(1)
        Exit Function
    ElseIf OUTPUT = 2 Then
        EUROPEAN_EXCHANGE_OPTION_FUNC = TEMP_GROUP(2)
        Exit Function
    ElseIf OUTPUT = 3 Then
        EUROPEAN_EXCHANGE_OPTION_FUNC = TEMP_GROUP
        Exit Function
    End If
'------------------------------------------------------------------------------------
Else
'------------------------------------------------------------------------------------
    PARAM_VECTOR = PARAM_RNG
    If UBound(PARAM_VECTOR, 1) = 1 Then: PARAM_VECTOR = MATRIX_TRANSPOSE_FUNC(PARAM_VECTOR)
    YDATA_VECTOR = HW_BK_SIGMA_OBJ_FUNC(XDATA_VECTOR, PARAM_VECTOR)
'------------------------------------------------------------------------------------
End If
'------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------

FIRST_GROUP = HW_TRINOM_TREE_FUNC(PUB_INIT_STEP_VAL, PUB_SIGMA_VAL, _
              PUB_MEAN_VAL, PUB_TENOR_ARR, PUB_EPS_VAL)

Call HW_FIT_RATE_FUNC(PUB_INIT_STEP_VAL, PUB_RF_TENOR_ARR, PUB_RF_DISC_ARR, _
              PUB_RF_RATE_ARR, FIRST_FIT_ARR, FIRST_PRICE_ARR, _
              FIRST_GROUP(1), PUB_TENOR_ARR, _
              FIRST_GROUP(2), FIRST_GROUP(3), _
              FIRST_GROUP(4), FIRST_GROUP(5), PUB_EPS_VAL, 0)

'-----------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------
'----------------------------Default Intensity Procedures---------------------------
'-----------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------

PUB_MEAN_VAL = DRIFT_VAL
PUB_SIGMA_VAL = SIGMA_VAL

SECOND_GROUP = HW_TRINOM_TREE_FUNC(PUB_INIT_STEP_VAL, PUB_SIGMA_VAL, _
               PUB_MEAN_VAL, PUB_TENOR_ARR, PUB_EPS_VAL)

nSTEPS = UBound(PUB_TENOR_ARR, 1)
ReDim SECOND_FIT_ARR(0 To nSTEPS)
ReDim SECOND_PRICE_ARR(0 To nSTEPS - 1)
ReDim TEMP_ARR(0 To 0)
TEMP_ARR(0) = 1
SECOND_PRICE_ARR(0) = TEMP_ARR

'-------------------------------------------------------------------------------
XDATA_VECTOR = POS_ADJ_RHO_RNG
If UBound(XDATA_VECTOR, 1) = 1 Then: XDATA_VECTOR = MATRIX_TRANSPOSE_FUNC(XDATA_VECTOR)
YDATA_VECTOR = NEG_ADJ_RHO_RNG
If UBound(YDATA_VECTOR, 1) = 1 Then: YDATA_VECTOR = MATRIX_TRANSPOSE_FUNC(YDATA_VECTOR)

If UBound(XDATA_VECTOR, 1) <> UBound(YDATA_VECTOR, 1) Then: GoTo ERROR_LABEL
NSIZE = UBound(XDATA_VECTOR, 1)
  
ReDim POS_ADJ_RHO_ARR(0 To NSIZE - 1)
ReDim NEG_ADJ_RHO_ARR(0 To NSIZE - 1)

j = 0
For i = 1 To NSIZE
  POS_ADJ_RHO_ARR(j) = XDATA_VECTOR(i, 1)
  NEG_ADJ_RHO_ARR(j) = YDATA_VECTOR(i, 1)
  j = j + 1
Next i

'---------------------------------------------------------------------------------------

nSTEPS = UBound(PUB_TENOR_ARR, 1)

ReDim RISKY_PRICE_ARR(0 To nSTEPS - 1)
ReDim RF_PRICE_ARR(0 To nSTEPS - 1)
ReDim ATEMP_ARR(0 To 0, 0 To 0)
ATEMP_ARR(0, 0) = 1
RISKY_PRICE_ARR(0) = ATEMP_ARR
RF_PRICE_ARR(0) = ATEMP_ARR

kk = HW_SURVIVAL_TREE_FUNC(PUB_INIT_STEP_VAL, NSIZE - 1, _
                kk, RHO_VAL, PUB_TENOR_ARR, RF_PRICE_ARR, _
                PUB_RF_DISC_ARR, PUB_RF_RATE_ARR, _
                PUB_RF_TENOR_ARR, RISKY_PRICE_ARR, RISKY_RATE_ARR, _
                RISKY_TENOR_ARR, FIRST_GROUP, FIRST_FIT_ARR, FIRST_PRICE_ARR, _
                SECOND_GROUP, SECOND_FIT_ARR, SECOND_PRICE_ARR, _
                POS_ADJ_RHO_ARR, NEG_ADJ_RHO_ARR, RHO_FACTOR, PUB_EPS_VAL)

k = HW_FIND_INDEX_FUNC(EXCH_EXPIRAT_VAL, PUB_TENOR_ARR, 0)

PUB_EXPIRAT_VAL = EXCH_EXPIRAT_VAL
ii = HW_NODE_TREE_FUNC(k, FIRST_GROUP(2), PUB_EPS_VAL)
jj = HW_NODE_TREE_FUNC(k, SECOND_GROUP(2), PUB_EPS_VAL)

ReDim PUB_RATE_ARR(0 To ii - 1, 0 To jj - 1)

k = HW_FIND_INDEX_FUNC(PUB_EXPIRAT_VAL, PUB_TENOR_ARR, 0)
ReDim BTEMP_ARR(0 To ii - 1)
ATEMP_ARR = RISKY_PRICE_ARR(k) 'risky State Prices

'BTEMP_ARR --> Risky State Bond Prices
'CTEMP_ARR --> Risk Free State Bond Prices

For i = 0 To ii - 1
  BTEMP_ARR(i) = 0
  For j = 0 To jj - 1
    BTEMP_ARR(i) = BTEMP_ARR(i) + ATEMP_ARR(i, j)
  Next j
Next i

kk = 0
Call HW_STATE_PRICE_FUNC(PUB_INIT_STEP_VAL, k, kk, FIRST_FIT_ARR, _
                   FIRST_PRICE_ARR, FIRST_GROUP(1), PUB_TENOR_ARR, _
                   FIRST_GROUP(2), FIRST_GROUP(3), FIRST_GROUP(4), _
                   FIRST_GROUP(5), PUB_EPS_VAL, 0, 0)

CTEMP_ARR = FIRST_PRICE_ARR(k)

ReDim DTEMP_ARR(0 To ii - 1) 'pay off Vector
For i = 0 To ii - 1
  DTEMP_ARR(i) = EXCH_NOMIN_VAL * MAXIMUM_FUNC(CTEMP_ARR(i) _
      - EXCH_RATIO_VAL * BTEMP_ARR(i), 0)
Next i
PUB_RATE_ARR = DTEMP_ARR

'------------------------------------------------------------------------------

If (PUB_EXPIRAT_VAL < 0) Then
  EUROPEAN_EXCHANGE_OPTION_FUNC = "Lattice: cannot roll the asset back to" & _
          0 & " (it is already at t = " & PUB_EXPIRAT_VAL
  Exit Function
End If

If (PUB_EXPIRAT_VAL > 0) Then
  jj = HW_FIND_INDEX_FUNC(PUB_EXPIRAT_VAL, PUB_TENOR_ARR, 0)
  ii = HW_FIND_INDEX_FUNC(0, PUB_TENOR_ARR, 0)
  For i = jj - 1 To ii Step -1
    ReDim ATEMP_ARR(0 To HW_NODE_TREE_FUNC(i, FIRST_GROUP(2), PUB_EPS_VAL) - 1)
    
    Call HW_STEP_BACK_FUNC(i, PUB_INIT_STEP_VAL, FIRST_GROUP(1), PUB_TENOR_ARR, _
          FIRST_FIT_ARR, PUB_RATE_ARR, ATEMP_ARR, _
              FIRST_GROUP(2), FIRST_GROUP(3), FIRST_GROUP(4), _
              FIRST_GROUP(5), PUB_EPS_VAL, 0, 0)
    
    
    PUB_EXPIRAT_VAL = PUB_TENOR_ARR(i)
    PUB_RATE_ARR = ATEMP_ARR
  Next i
End If


BTEMP_ARR = PUB_RATE_ARR
CTEMP_ARR = FIRST_PRICE_ARR(HW_FIND_INDEX_FUNC(PUB_EXPIRAT_VAL, PUB_TENOR_ARR, 0))

TEMP_SUM = 0
For i = 0 To UBound(BTEMP_ARR, 1)
  TEMP_SUM = TEMP_SUM + BTEMP_ARR(i) * CTEMP_ARR(i)
Next i

EUROPEAN_EXCHANGE_OPTION_FUNC = TEMP_SUM

Exit Function
ERROR_LABEL:
EUROPEAN_EXCHANGE_OPTION_FUNC = Err.number
End Function

'// PERFECT

'************************************************************************************
'************************************************************************************



'************************************************************************************
'************************************************************************************
'FUNCTION      : BERMUDAN_SWAPTION_FUNC

'DESCRIPTION   : Calculates price of European or Bermudean Swaption in Hull White
'Model using Trinomial Tree

'-----------------------------------------------------------------------------------
'If VERSION = 0 Then
    'The function calculates price of Bermudan swaption using the assumptions:
    '1.Interest rate model is Hull White
    '2.Term structure is flat with continuous compunding
    '3.Stopping times for bermudan swaption are the reset times of fixed leg of swap
'Else
'   The function calculates price of European swaption in Hull White model using
'   Jamshidian trick of decomposing option on a coupon bearing bond into a portfolio
'   of put options on bond options which are based on a spot rate r* which
'   satisfies the equation as described in the paper:
'   http://www.ma.ic.ac.uk/~mdavis/course_material/SDEIRM/HW_SWAPTION_FORMULA.pdf
'End If
'-----------------------------------------------------------------------------------

'LIBRARY       : FIXED_INCOME
'GROUP         : HW_2F
'ID            : 004
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 10-06-2008
'************************************************************************************
'************************************************************************************

Function BERMUDAN_SWAPTION_FUNC(ByVal MEAN_VAL As Double, _
ByVal SIGMA_VAL As Double, _
ByVal FIXED_RATE As Double, _
ByVal FLAT_RATE As Double, _
ByVal OPTION_TENOR As Double, _
ByVal SWAP_LENGTH_TENOR As Double, _
Optional ByVal nSTEPS As Long = 75, _
Optional ByVal NOMINAL As Double = 1000, _
Optional ByVal DELTA_TENOR As Double = 0.5, _
Optional ByVal epsilon As Double = 2 ^ 52, _
Optional ByVal tolerance As Double = 0.01, _
Optional ByVal VERSION As Integer = 0)

'OPTION TENOR: Maturity of Option
'FLAT_RATE: Specify flat rate assuming continous compounding

Dim i As Long
Dim j As Long

Dim ii As Long
Dim jj As Long

Dim iii As Long
Dim jjj As Long

Dim NSIZE As Long

Dim TEMP_SUM As Double
Dim END_TENOR_VAL As Double
Dim START_TENOR_VAL As Double

Dim ATEMP_DELTA As Double
Dim BTEMP_DELTA As Double
Dim CTEMP_DELTA As Double
Dim DTEMP_DELTA As Double

Dim COUPON_VAL As Double
Dim INIT_STEP_VAL As Double

Dim GRAD_ARR As Variant
Dim TEMP_GROUP As Variant

Dim FIT_ARR As Variant
Dim BERM_ARR As Variant

Dim ATEMP_ARR As Variant
Dim BTEMP_ARR As Variant
Dim CTEMP_ARR As Variant
Dim DTEMP_ARR As Variant

Dim PRICE_ARR As Variant

Dim ATENOR_ARR As Variant
Dim BTENOR_ARR As Variant

Dim COUPON_ARR As Variant
Dim SWAPTION_ARR As Variant

Dim FIX_RES_TENOR As Variant
Dim FIX_PAY_TENOR As Variant
Dim FLOAT_RES_TENOR As Variant
Dim FLOAT_PAY_TENOR As Variant
  
Dim PROB_UP_ARR As Variant
Dim PROB_MED_ARR As Variant
Dim PROB_DOWN_ARR As Variant
Dim PROB_INDEX_ARR As Variant
    
On Error GoTo ERROR_LABEL

  INIT_STEP_VAL = 0
  ATEMP_DELTA = (OPTION_TENOR + SWAP_LENGTH_TENOR) / nSTEPS 'Swap Length
  NSIZE = CInt(0.5 / ATEMP_DELTA)
  NSIZE = NSIZE * 2 * (OPTION_TENOR + SWAP_LENGTH_TENOR)
  
  'NSIZE --> Time steps used
  '(END_TIME / NSIZE) should be such that the swap pegs
  'are also touched otherwise it is possible that
  'distance between the swap mandatory times and the time
  'grid division poisnt is so small that variance is 0 and
  'causes error while building the trinomial tree
  
  ReDim ATENOR_ARR(0 To NSIZE)
  ATEMP_DELTA = (OPTION_TENOR + SWAP_LENGTH_TENOR - 0) / NSIZE
  ATENOR_ARR(0) = 0
  For i = 1 To NSIZE
    ATENOR_ARR(i) = ATENOR_ARR(i - 1) + ATEMP_DELTA
  Next i
  
  NSIZE = SWAP_LENGTH_TENOR / DELTA_TENOR
  
  ReDim FIX_RES_TENOR(0 To NSIZE - 1)
  ReDim FIX_PAY_TENOR(0 To NSIZE - 1)
  ReDim FLOAT_RES_TENOR(0 To NSIZE - 1)
  ReDim FLOAT_PAY_TENOR(0 To NSIZE - 1)
  ReDim COUPON_ARR(0 To NSIZE - 1)
  ReDim BTENOR_ARR(0 To NSIZE - 1)
    
  TEMP_SUM = OPTION_TENOR

  For i = 0 To NSIZE - 1
    FIX_RES_TENOR(i) = TEMP_SUM
    FIX_PAY_TENOR(i) = TEMP_SUM + DELTA_TENOR
    FLOAT_RES_TENOR(i) = TEMP_SUM
    FLOAT_PAY_TENOR(i) = TEMP_SUM + DELTA_TENOR
    COUPON_ARR(i) = NOMINAL * FIXED_RATE * DELTA_TENOR
    TEMP_SUM = TEMP_SUM + DELTA_TENOR
  Next i

  BTENOR_ARR = FIX_RES_TENOR
  
  ATENOR_ARR = ARRAY_MERGE_FUNC(ATENOR_ARR, FIX_RES_TENOR, tolerance)
  ATENOR_ARR = ARRAY_MERGE_FUNC(ATENOR_ARR, FIX_PAY_TENOR, tolerance)
  
  TEMP_GROUP = HW_TRINOM_TREE_FUNC(INIT_STEP_VAL, SIGMA_VAL, _
              MEAN_VAL, ATENOR_ARR, epsilon)
  
  GRAD_ARR = TEMP_GROUP(1)
  PROB_INDEX_ARR = TEMP_GROUP(2)
  PROB_DOWN_ARR = TEMP_GROUP(3)
  PROB_UP_ARR = TEMP_GROUP(4)
  PROB_MED_ARR = TEMP_GROUP(5)

  TEMP_GROUP = HW_FIT_SWAP_FUNC(INIT_STEP_VAL, FLAT_RATE, FIT_ARR, PRICE_ARR, _
                GRAD_ARR, ATENOR_ARR, PROB_INDEX_ARR, PROB_DOWN_ARR, _
                PROB_UP_ARR, PROB_MED_ARR, epsilon, tolerance)
   
  BTEMP_DELTA = UBound(FIX_PAY_TENOR, 1)
  BTEMP_DELTA = FIX_PAY_TENOR(BTEMP_DELTA)
  
  ATENOR_ARR = ARRAY_MERGE_FUNC(ATENOR_ARR, FIX_RES_TENOR, tolerance)
  ATENOR_ARR = ARRAY_MERGE_FUNC(ATENOR_ARR, FIX_PAY_TENOR, tolerance)
  
  CTEMP_DELTA = BTEMP_DELTA
  '-----------------------RESET_SWAPTION-----------------------
  i = HW_NODE_TREE_FUNC(HW_FIND_INDEX_FUNC(BTEMP_DELTA, ATENOR_ARR, tolerance), _
                    PROB_INDEX_ARR, epsilon)
  ReDim SWAPTION_ARR(0 To i - 1)
  For j = 0 To i - 1
    SWAPTION_ARR(j) = 0
  Next j
  BERM_ARR = SWAPTION_ARR
  GoSub 1983
    
  END_TENOR_VAL = FIX_RES_TENOR(0)
  START_TENOR_VAL = CTEMP_DELTA
  If (START_TENOR_VAL > END_TENOR_VAL) Then
    ii = HW_FIND_INDEX_FUNC(START_TENOR_VAL, ATENOR_ARR, tolerance)
    jj = HW_FIND_INDEX_FUNC(END_TENOR_VAL, ATENOR_ARR, tolerance)
    For i = ii - 1 To jj Step -1
      ReDim BTEMP_ARR(0 To HW_NODE_TREE_FUNC(i, PROB_INDEX_ARR, epsilon) - 1)
      ATEMP_ARR = SWAPTION_ARR
      Call HW_STEP_BACK_FUNC(i, INIT_STEP_VAL, GRAD_ARR, ATENOR_ARR, FIT_ARR, _
            ATEMP_ARR, BTEMP_ARR, _
            PROB_INDEX_ARR, PROB_DOWN_ARR, PROB_UP_ARR, _
            PROB_MED_ARR, epsilon, tolerance, 0)
      
      CTEMP_DELTA = ATENOR_ARR(i)
      SWAPTION_ARR = BTEMP_ARR
      GoSub 1983
    Next i
  End If
  
  Select Case VERSION
    Case 0 'Calculates price of Bermudan Swaption in Hull White
    'Model using Trinomial Tree
          
      TEMP_SUM = 0
      ATEMP_ARR = PRICE_ARR(HW_FIND_INDEX_FUNC(CTEMP_DELTA, ATENOR_ARR, tolerance))
      For i = 0 To UBound(BERM_ARR, 1)
        TEMP_SUM = TEMP_SUM + BERM_ARR(i) * ATEMP_ARR(i)
      Next i
      BERMUDAN_SWAPTION_FUNC = TEMP_SUM
      Exit Function
    Case Else 'Calculates price of European Swaption in Hull White Model
    'using Trinomial Tree
      TEMP_SUM = 0
      ATEMP_ARR = PRICE_ARR(HW_FIND_INDEX_FUNC(CTEMP_DELTA, ATENOR_ARR, tolerance))
      For i = 0 To UBound(SWAPTION_ARR, 1)
        SWAPTION_ARR(i) = MAXIMUM_FUNC(0, SWAPTION_ARR(i))
        TEMP_SUM = TEMP_SUM + SWAPTION_ARR(i) * ATEMP_ARR(i)
      Next i
      BERMUDAN_SWAPTION_FUNC = TEMP_SUM
      Exit Function
  End Select
  
Exit Function
1983:
'Updates the member variable SWAPTION_ARR while the
'trinomial tree is being rolled back on the time grid
'and when the timepeg for swap reset is approached
    
'------------------------------------------------------------------------------------
Select Case VERSION
'------------------------------------------------------------------------------------
Case 0 'For Bermudan Swaption in Hull White Model
'------------------------------------------------------------------------------------
  For iii = 0 To UBound(FLOAT_RES_TENOR, 1)
    DTEMP_DELTA = FLOAT_RES_TENOR(iii)
    If (HW_ON_TIME_FUNC(DTEMP_DELTA, CTEMP_DELTA, ATENOR_ARR, tolerance)) Then
      CTEMP_ARR = HW_ROLL_BACK_OBJ_FUNC(FLOAT_PAY_TENOR(iii), CTEMP_DELTA, _
            INIT_STEP_VAL, GRAD_ARR, ATENOR_ARR, FIT_ARR, _
            PROB_INDEX_ARR, PROB_DOWN_ARR, PROB_UP_ARR, _
            PROB_MED_ARR, , epsilon, tolerance, 0, 0)
      For jjj = 0 To UBound(SWAPTION_ARR, 1)
        COUPON_VAL = NOMINAL * (1 - CTEMP_ARR(jjj))
        SWAPTION_ARR(jjj) = SWAPTION_ARR(jjj) + COUPON_VAL
      Next jjj
    End If
  Next iii
  
  For iii = 0 To UBound(FIX_RES_TENOR, 1)
    DTEMP_DELTA = FIX_RES_TENOR(iii)
    If (HW_ON_TIME_FUNC(DTEMP_DELTA, CTEMP_DELTA, ATENOR_ARR, tolerance)) Then
      CTEMP_ARR = HW_ROLL_BACK_OBJ_FUNC(FIX_PAY_TENOR(iii), CTEMP_DELTA, _
            INIT_STEP_VAL, GRAD_ARR, ATENOR_ARR, FIT_ARR, _
            PROB_INDEX_ARR, PROB_DOWN_ARR, PROB_UP_ARR, _
            PROB_MED_ARR, , epsilon, tolerance, 0, 0)
      For jjj = 0 To UBound(SWAPTION_ARR, 1)
        COUPON_VAL = COUPON_ARR(iii) * CTEMP_ARR(jjj)
        SWAPTION_ARR(jjj) = SWAPTION_ARR(jjj) - COUPON_VAL
      Next jjj
    End If
  Next iii
  
  For iii = 0 To UBound(BTENOR_ARR, 1)
    DTEMP_DELTA = BTENOR_ARR(iii)
    If (HW_ON_TIME_FUNC(DTEMP_DELTA, CTEMP_DELTA, ATENOR_ARR, tolerance)) Then
      CTEMP_DELTA = DTEMP_DELTA
      
      DTEMP_ARR = HW_ROLL_BACK_OBJ_FUNC(FLOAT_PAY_TENOR(iii), DTEMP_DELTA, INIT_STEP_VAL, GRAD_ARR, _
                      ATENOR_ARR, FIT_ARR, PROB_INDEX_ARR, _
                      PROB_DOWN_ARR, PROB_UP_ARR, _
                      PROB_MED_ARR, BERM_ARR, epsilon, tolerance, 0, 0)
      For jjj = 0 To UBound(DTEMP_ARR, 1)
        DTEMP_ARR(jjj) = MAXIMUM_FUNC(DTEMP_ARR(jjj), SWAPTION_ARR(jjj))
      Next jjj
      BERM_ARR = DTEMP_ARR
    End If
  Next iii
'------------------------------------------------------------------------------------
Case Else 'For European Swaption in Hull White Model
'------------------------------------------------------------------------------------
  For iii = 0 To UBound(FLOAT_RES_TENOR, 1)
    DTEMP_DELTA = FLOAT_RES_TENOR(iii)
    If (HW_ON_TIME_FUNC(DTEMP_DELTA, CTEMP_DELTA, ATENOR_ARR, tolerance)) Then
      
      CTEMP_ARR = HW_ROLL_BACK_OBJ_FUNC(FLOAT_PAY_TENOR(iii), CTEMP_DELTA, _
            INIT_STEP_VAL, GRAD_ARR, ATENOR_ARR, FIT_ARR, _
            PROB_INDEX_ARR, PROB_DOWN_ARR, PROB_UP_ARR, _
            PROB_MED_ARR, , epsilon, tolerance, 0, 0)
      For jjj = 0 To UBound(SWAPTION_ARR, 1)
        COUPON_VAL = NOMINAL * (1 - CTEMP_ARR(jjj))
        SWAPTION_ARR(jjj) = SWAPTION_ARR(jjj) + COUPON_VAL
      Next jjj
    End If
  Next iii
  
  For iii = 0 To UBound(FIX_RES_TENOR, 1)
    DTEMP_DELTA = FIX_RES_TENOR(iii)
    If (HW_ON_TIME_FUNC(DTEMP_DELTA, CTEMP_DELTA, ATENOR_ARR, tolerance)) Then
      
      CTEMP_ARR = HW_ROLL_BACK_OBJ_FUNC(FIX_PAY_TENOR(iii), CTEMP_DELTA, _
            INIT_STEP_VAL, GRAD_ARR, ATENOR_ARR, FIT_ARR, _
            PROB_INDEX_ARR, PROB_DOWN_ARR, PROB_UP_ARR, _
            PROB_MED_ARR, , epsilon, tolerance, 0, 0)
      For jjj = 0 To UBound(SWAPTION_ARR, 1)
        COUPON_VAL = COUPON_ARR(iii) * CTEMP_ARR(jjj)
        SWAPTION_ARR(jjj) = SWAPTION_ARR(jjj) - COUPON_VAL
      Next jjj
    End If
  Next iii
End Select
Return
'-------------------------------------------------------------------------------------
ERROR_LABEL:
BERMUDAN_SWAPTION_FUNC = Err.number
End Function

'// PERFECT

'************************************************************************************
'************************************************************************************



'************************************************************************************
'************************************************************************************
'FUNCTION      : HW_FIT_SWAP_FUNC
'DESCRIPTION   :
'LIBRARY       : FIXED_INCOME
'GROUP         : HW_2F
'ID            : 005
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 10-06-2008
'************************************************************************************
'************************************************************************************

Private Function HW_FIT_SWAP_FUNC(ByVal INIT_STEP_VAL As Double, _
ByVal FLAT_RATE As Double, _
ByRef FIT_ARR As Variant, _
ByRef PRICE_ARR As Variant, _
ByRef GRAD_ARR As Variant, _
ByRef TENOR_ARR As Variant, _
ByRef PROB_INDEX_ARR As Variant, _
ByRef PROB_DOWN_ARR As Variant, _
ByRef PROB_UP_ARR As Variant, _
ByRef PROB_MED_ARR As Variant, _
Optional ByVal epsilon As Double = 2 ^ 52, _
Optional ByVal tolerance As Double = 0.01)

Dim i As Long
Dim j As Long
Dim k As Long
Dim l As Long

Dim nSTEPS As Long

Dim TEMP_SUM As Double
Dim TEMP_DELTA As Double
Dim TEMP_GRAD As Double

Dim X_TEMP_VAL As Double
Dim Y_TEMP_VAL As Double

Dim TEMP_ARR As Variant

On Error GoTo ERROR_LABEL

  nSTEPS = UBound(TENOR_ARR, 1)
  ReDim PRICE_ARR(0 To nSTEPS - 1)
  
  ReDim TEMP_ARR(0 To 0)
  TEMP_ARR(0) = 1
  PRICE_ARR(0) = TEMP_ARR
  
  ReDim FIT_ARR(0 To nSTEPS - 1) 'Fix This One
  l = 0
  For i = 0 To nSTEPS - 1

    Call HW_STATE_PRICE_FUNC(INIT_STEP_VAL, i, l, FIT_ARR, PRICE_ARR, _
        GRAD_ARR, TENOR_ARR, PROB_INDEX_ARR, PROB_DOWN_ARR, PROB_UP_ARR, _
        PROB_MED_ARR, epsilon, tolerance, 0)
    
    TEMP_ARR = PRICE_ARR(i)
    k = HW_NODE_TREE_FUNC(i, PROB_INDEX_ARR, epsilon)
    
    TEMP_DELTA = TENOR_ARR(i + 1) - TENOR_ARR(i)
    TEMP_GRAD = GRAD_ARR(i)
    X_TEMP_VAL = HW_NODE_RATE_FUNC(i, 0, INIT_STEP_VAL, GRAD_ARR, _
                                PROB_INDEX_ARR, epsilon)
    
    TEMP_SUM = 0
    For j = 0 To k - 1
      TEMP_SUM = TEMP_SUM + TEMP_ARR(j) * Exp(-X_TEMP_VAL * TEMP_DELTA)
      X_TEMP_VAL = X_TEMP_VAL + TEMP_GRAD
    Next j
    
    Y_TEMP_VAL = Exp(-FLAT_RATE * TENOR_ARR(i + 1)) 'Discount Bond
    FIT_ARR(i) = (Log(TEMP_SUM / Y_TEMP_VAL) / Log(Exp(1))) / TEMP_DELTA
  Next i

Exit Function
ERROR_LABEL:
    HW_FIT_SWAP_FUNC = Err.number
End Function

'// PERFECT


'************************************************************************************
'************************************************************************************



'************************************************************************************
'************************************************************************************
'FUNCTION      : HW_BK_SIGMA_OBJ_FUNC
'DESCRIPTION   : Call back function for calibration
'LIBRARY       : FIXED_INCOME
'GROUP         : HW_2F
'ID            : 006
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 10-06-2008
'************************************************************************************
'************************************************************************************

Private Function HW_BK_SIGMA_OBJ_FUNC(ByRef XDATA_RNG As Variant, _
ByRef PARAM_RNG As Variant)
  
  Dim i As Long
  Dim j As Long
  Dim k As Long
  
  Dim NSIZE As Long
  
  Dim DRIFT_VAL As Double
  Dim SIGMA_VAL As Double
  
  Dim TO_VAL As Double
  Dim FROM_VAL As Double
  Dim NPV_VAL As Double
  
  Dim XTEMP_ARR As Variant
  Dim YTEMP_ARR As Variant
  
  Dim FIT_ARR As Variant
  Dim PRICE_ARR As Variant
  Dim GRAD_ARR As Variant
  
  Dim PROB_INDEX_ARR As Variant
  Dim PROB_DOWN_ARR As Variant
  Dim PROB_UP_ARR As Variant
  Dim PROB_MED_ARR As Variant
  Dim TEMP_GROUP As Variant
  
  Dim XDATA_VECTOR As Variant
  Dim PARAM_VECTOR As Variant
  Dim YDATA_VECTOR As Variant
  
  On Error GoTo ERROR_LABEL
  XDATA_VECTOR = XDATA_RNG
  If UBound(XDATA_VECTOR, 1) = 1 Then: XDATA_VECTOR = MATRIX_TRANSPOSE_FUNC(XDATA_VECTOR)
  NSIZE = UBound(XDATA_VECTOR, 1)
  
  PARAM_VECTOR = PARAM_RNG
  If UBound(PARAM_VECTOR, 1) = 1 Then: PARAM_VECTOR = MATRIX_TRANSPOSE_FUNC(PARAM_VECTOR)
  
  ReDim YDATA_VECTOR(1 To NSIZE, 1 To 1)
  DRIFT_VAL = PARAM_VECTOR(1, 1)
  SIGMA_VAL = PARAM_VECTOR(2, 1)
  
'------------------------------------------------------------------------------------
  PUB_MEAN_VAL = DRIFT_VAL
  PUB_SIGMA_VAL = SIGMA_VAL
'------------------------------------------------------------------------------------
  
  'Populate grid before passing into Trinomial tree class

  For i = 0 To NSIZE - 1
          PUB_CAP_VAL = HW_FAIR_RATE_FUNC(PUB_START_TENOR_VAL, XDATA_VECTOR(i + 1, 1), _
                        PUB_NOM_VAL, PUB_FIXED_RATE_VAL, PUB_START_ARR, PUB_END_ARR, _
                        PUB_ACCR_ARR, PUB_FIX_ARR, PUB_FORW_ARR, _
                        PUB_RF_DISC_ARR, PUB_RF_RATE_ARR, PUB_RF_TENOR_ARR)
          PUB_TENOR_ARR = ARRAY_MERGE_FUNC(PUB_TENOR_ARR, PUB_START_ARR)
  Next i
  
  TEMP_GROUP = HW_TRINOM_TREE_FUNC(PUB_INIT_STEP_VAL, PUB_SIGMA_VAL, _
                    PUB_MEAN_VAL, PUB_TENOR_ARR, PUB_EPS_VAL)
        
  GRAD_ARR = TEMP_GROUP(1)
  PROB_INDEX_ARR = TEMP_GROUP(2)
  PROB_DOWN_ARR = TEMP_GROUP(3)
  PROB_UP_ARR = TEMP_GROUP(4)
  PROB_MED_ARR = TEMP_GROUP(5)
  
  If PUB_VERSION_VAL = 0 Then
    Call HW_FIT_RATE_FUNC(PUB_INIT_STEP_VAL, PUB_RF_TENOR_ARR, _
        PUB_RF_DISC_ARR, PUB_RF_RATE_ARR, FIT_ARR, _
        PRICE_ARR, GRAD_ARR, PUB_TENOR_ARR, PROB_INDEX_ARR, _
        PROB_DOWN_ARR, PROB_UP_ARR, PROB_MED_ARR, PUB_EPS_VAL, 0)
  Else
    Call BK_FIT_RATE_FUNC(PUB_INIT_STEP_VAL, PUB_RF_TENOR_ARR, _
        PUB_RF_DISC_ARR, PUB_RF_RATE_ARR, FIT_ARR, _
        PRICE_ARR, GRAD_ARR, PUB_TENOR_ARR, PROB_INDEX_ARR, _
        PROB_DOWN_ARR, PROB_UP_ARR, PROB_MED_ARR, PUB_EPS_VAL, 0)
  End If
  
For i = 0 To NSIZE - 1
        
    PUB_CAP_VAL = HW_FAIR_RATE_FUNC(PUB_START_TENOR_VAL, XDATA_VECTOR(i + 1, 1), _
                    PUB_NOM_VAL, PUB_FIXED_RATE_VAL, PUB_START_ARR, PUB_END_ARR, _
                    PUB_ACCR_ARR, PUB_FIX_ARR, PUB_FORW_ARR, _
                    PUB_RF_DISC_ARR, PUB_RF_RATE_ARR, PUB_RF_TENOR_ARR)
    
    
    PUB_EXPIRAT_VAL = PUB_END_ARR(UBound(PUB_END_ARR, 1))
    k = HW_NODE_TREE_FUNC(HW_FIND_INDEX_FUNC(k, PUB_TENOR_ARR, 0), _
                    PROB_INDEX_ARR, PUB_EPS_VAL)
    
      ReDim PUB_RATE_ARR(0 To k - 1)
      For j = 0 To k - 1
        PUB_RATE_ARR(j) = 0
      Next j
      Call HW_PRE_ADJ_FUNC(PUB_INIT_STEP_VAL, PUB_NOM_VAL, PUB_CAP_VAL, PUB_EXPIRAT_VAL, _
                        PUB_TENOR_ARR, PUB_START_ARR, PUB_END_ARR, _
                        PUB_ACCR_ARR, PUB_RATE_ARR, FIT_ARR, GRAD_ARR, _
                        PROB_INDEX_ARR, PROB_DOWN_ARR, PROB_UP_ARR, _
                        PROB_MED_ARR, PUB_EPS_VAL, PUB_VERSION_VAL) '  PostAdjustValues

  TO_VAL = PUB_START_ARR(0)
  FROM_VAL = PUB_EXPIRAT_VAL
  If (FROM_VAL < TO_VAL) Then
    HW_BK_SIGMA_OBJ_FUNC = "Lattice: cannot roll the asset back to" & _
            TO_VAL & " (it is already at t = " & FROM_VAL
    Exit Function
  Else
    PUB_RATE_ARR = HW_ROLL_BACK_OBJ_FUNC(FROM_VAL, TO_VAL, PUB_INIT_STEP_VAL, GRAD_ARR, _
                    PUB_TENOR_ARR, FIT_ARR, PROB_INDEX_ARR, PROB_DOWN_ARR, _
                    PROB_UP_ARR, PROB_MED_ARR, PUB_RATE_ARR, PUB_EPS_VAL, 0, _
                    PUB_VERSION_VAL, 1)
  End If
  
    XTEMP_ARR = PUB_RATE_ARR
    YTEMP_ARR = PRICE_ARR(HW_FIND_INDEX_FUNC(PUB_EXPIRAT_VAL, PUB_TENOR_ARR, 0))
    NPV_VAL = 0
    For j = 0 To UBound(XTEMP_ARR, 1)
        NPV_VAL = NPV_VAL + XTEMP_ARR(j) * YTEMP_ARR(j)
    Next j
    PUB_TARGET_VAL = NPV_VAL
    YDATA_VECTOR(i + 1, 1) = BRENT_ZERO_FUNC(PUB_LOWER_BOUND_VAL, PUB_UPPER_BOUND_VAL, "HW_SIGMA_ERROR_OBJ_FUNC", CDbl(PUB_GUESS_VAL), , , PUB_nLOOPS_VAL, PUB_TOL_VAL)
Next i

HW_BK_SIGMA_OBJ_FUNC = YDATA_VECTOR

Exit Function
ERROR_LABEL:
HW_BK_SIGMA_OBJ_FUNC = Err.number
End Function

'// PERFECT

'************************************************************************************
'************************************************************************************



'************************************************************************************
'************************************************************************************
'FUNCTION      : HW_SIGMA_ERROR_OBJ_FUNC
'DESCRIPTION   : IMPLIED SIGMA ERROR FUNCTION
'LIBRARY       : FIXED_INCOME
'GROUP         : HW_2F
'ID            : 007
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 10-06-2008
'************************************************************************************
'************************************************************************************

Private Function HW_SIGMA_ERROR_OBJ_FUNC(ByVal X_DATA_VAL As Double)
 
  Dim i As Long
  Dim NROWS As Long
  
  Dim FIXING As Double
  Dim FORWARD As Double
  
  Dim END_TIME As Double
  Dim ACCR_TIME As Double
  
  Dim TEMP_SUM As Double
  Dim TEMP_MULT As Double
  
  On Error GoTo ERROR_LABEL
  
  NROWS = UBound(PUB_START_ARR, 1)
  TEMP_SUM = 0
  For i = 0 To NROWS
    
    FIXING = PUB_FIX_ARR(i)
    END_TIME = PUB_END_ARR(i)
    ACCR_TIME = PUB_ACCR_ARR(i)
    
    If (END_TIME > 0) Then     'discard expired caplets
      
      TEMP_MULT = HW_ZERO_RATE_FUNC(END_TIME, PUB_RF_TENOR_ARR, _
                PUB_RF_DISC_ARR, PUB_RF_RATE_ARR)
      
      FORWARD = PUB_FORW_ARR(i)
      
      TEMP_SUM = TEMP_SUM + TEMP_MULT * ACCR_TIME * PUB_NOM_VAL * _
                    HW_CAPLET_FUNC(FIXING, FORWARD, _
                    PUB_CAP_VAL, X_DATA_VAL)
    End If
  
  Next i
  
  HW_SIGMA_ERROR_OBJ_FUNC = PUB_TARGET_VAL - TEMP_SUM

Exit Function
ERROR_LABEL:
    HW_SIGMA_ERROR_OBJ_FUNC = Err.number
End Function


'// PERFECT

'************************************************************************************
'************************************************************************************



'************************************************************************************
'************************************************************************************
'FUNCTION      : HW_FIT_RATE_FUNC
'DESCRIPTION   : Calc fitting values for Hull-White
'LIBRARY       : FIXED_INCOME
'GROUP         : HW_2F
'ID            : 008
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 10-06-2008
'************************************************************************************
'************************************************************************************

Private Function HW_FIT_RATE_FUNC(ByVal INIT_STEP_VAL As Double, _
ByVal TENOR_ARR As Variant, _
ByVal DISC_ARR As Variant, _
ByVal RATE_ARR As Variant, _
ByRef FIT_ARR As Variant, _
ByRef PRICE_ARR As Variant, _
ByRef GRAD_ARR As Variant, _
ByRef TIME_ARR As Variant, _
ByRef PROB_INDEX_ARR As Variant, _
ByRef PROB_DOWN_ARR As Variant, _
ByRef PROB_UP_ARR As Variant, _
ByRef PROB_MED_ARR As Variant, _
Optional ByVal epsilon As Double = 2 ^ 52, _
Optional ByVal tolerance As Double = 0.01)

Dim i As Long
Dim j As Long
Dim l As Long

Dim nSTEPS As Long

Dim TEMP_SUM As Double
Dim TEMP_SIZE As Double
Dim TEMP_DELTA As Double
Dim TEMP_GRAD As Double

Dim X_TEMP_VAL As Double
Dim Y_TEMP_VAL As Double

Dim TEMP_ARR As Variant

On Error GoTo ERROR_LABEL

  nSTEPS = UBound(TIME_ARR, 1)
  ReDim PRICE_ARR(0 To nSTEPS - 1)
  
  ReDim TEMP_ARR(0 To 0)
  TEMP_ARR(0) = 1
  PRICE_ARR(0) = TEMP_ARR
  
  ReDim FIT_ARR(0 To nSTEPS - 1)
  l = 0
  For i = 0 To nSTEPS - 1

    Call HW_STATE_PRICE_FUNC(INIT_STEP_VAL, i, l, FIT_ARR, PRICE_ARR, _
        GRAD_ARR, TIME_ARR, PROB_INDEX_ARR, PROB_DOWN_ARR, PROB_UP_ARR, _
        PROB_MED_ARR, epsilon, tolerance, 0)
    
    TEMP_ARR = PRICE_ARR(i)
    TEMP_SIZE = HW_NODE_TREE_FUNC(i, PROB_INDEX_ARR, epsilon)
    
    TEMP_DELTA = TIME_ARR(i + 1) - TIME_ARR(i)
    TEMP_GRAD = GRAD_ARR(i)
    X_TEMP_VAL = HW_NODE_RATE_FUNC(i, 0, INIT_STEP_VAL, GRAD_ARR, _
                                PROB_INDEX_ARR, epsilon)
    
    TEMP_SUM = 0
    For j = 0 To TEMP_SIZE - 1
      TEMP_SUM = TEMP_SUM + TEMP_ARR(j) * Exp(-X_TEMP_VAL * TEMP_DELTA)
      X_TEMP_VAL = X_TEMP_VAL + TEMP_GRAD
    Next j
    
    Y_TEMP_VAL = HW_ZERO_RATE_FUNC(TIME_ARR(i + 1), TENOR_ARR, _
                 DISC_ARR, RATE_ARR) 'Discount Bond
    FIT_ARR(i) = (Log(TEMP_SUM / Y_TEMP_VAL) / Log(Exp(1))) / TEMP_DELTA
  Next i

Exit Function
ERROR_LABEL:
    HW_FIT_RATE_FUNC = Err.number
End Function

'// PERFECT

'************************************************************************************
'************************************************************************************



'************************************************************************************
'************************************************************************************
'FUNCTION      : BK_FIT_RATE_FUNC
'DESCRIPTION   : Calc fitting values for Black-Karansinki
'LIBRARY       : FIXED_INCOME
'GROUP         : HW_2F
'ID            : 009
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 10-06-2008
'************************************************************************************
'************************************************************************************

Private Function BK_FIT_RATE_FUNC(ByVal INIT_STEP_VAL As Double, _
ByVal TENOR_ARR As Variant, _
ByVal DISC_ARR As Variant, _
ByVal RATE_ARR As Variant, _
ByRef FIT_ARR As Variant, _
ByRef PRICE_ARR As Variant, _
ByRef GRAD_ARR As Variant, _
ByRef TIME_ARR As Variant, _
ByRef PROB_INDEX_ARR As Variant, _
ByRef PROB_DOWN_ARR As Variant, _
ByRef PROB_UP_ARR As Variant, _
ByRef PROB_MED_ARR As Variant, _
Optional ByVal epsilon As Double = 2 ^ 52, _
Optional ByVal tolerance As Double = 0.01, _
Optional ByVal BRENT_MIN_VAL As Double = -500, _
Optional ByVal BRENT_MAX_VAL As Double = 500, _
Optional ByVal BRENT_nLOOPS As Long = 100, _
Optional ByVal BRENT_TOLER As Double = 0.00000001)

Dim i As Long
Dim l As Long

Dim nSTEPS As Long

Dim TEMP_SIZE As Double
Dim TEMP_DELTA As Double
Dim TEMP_GRAD As Double

Dim X_TEMP_VAL As Double
Dim Y_TEMP_VAL As Double
Dim Z_TEMP_VAL As Double

Dim TEMP_ARR As Variant

On Error GoTo ERROR_LABEL

nSTEPS = UBound(TIME_ARR, 1)
ReDim PRICE_ARR(0 To nSTEPS - 1)

ReDim TEMP_ARR(0 To 0)
TEMP_ARR(0) = 1
PRICE_ARR(0) = TEMP_ARR

ReDim FIT_ARR(0 To nSTEPS - 1)
l = 0
Z_TEMP_VAL = 1
For i = 0 To nSTEPS - 1
    Call HW_STATE_PRICE_FUNC(INIT_STEP_VAL, i, l, FIT_ARR, PRICE_ARR, GRAD_ARR, TIME_ARR, PROB_INDEX_ARR, PROB_DOWN_ARR, PROB_UP_ARR, PROB_MED_ARR, epsilon, tolerance, 1)
    TEMP_ARR = PRICE_ARR(i)
    TEMP_SIZE = HW_NODE_TREE_FUNC(i, PROB_INDEX_ARR, epsilon)
    TEMP_DELTA = TIME_ARR(i + 1) - TIME_ARR(i)
    TEMP_GRAD = GRAD_ARR(i)
    X_TEMP_VAL = HW_NODE_RATE_FUNC(i, 0, INIT_STEP_VAL, GRAD_ARR, PROB_INDEX_ARR, epsilon)
    PUB_BK_MIN = X_TEMP_VAL
    '    TEMP_SUM = 0
    '   For j = 0 To TEMP_SIZE - 1
    '    TEMP_SUM = TEMP_SUM + TEMP_ARR(j) * Exp(-X_TEMP_VAL * TEMP_DELTA)
    '   X_TEMP_VAL = X_TEMP_VAL + TEMP_GRAD
    'Next j
    Y_TEMP_VAL = HW_ZERO_RATE_FUNC(TIME_ARR(i + 1), TENOR_ARR, DISC_ARR, RATE_ARR) 'Discount Bond
    PUB_BK_DELTA = TEMP_DELTA
    PUB_BK_GRAD = TEMP_GRAD
    PUB_BK_SIZE = TEMP_SIZE
    PUB_BK_PRICE = Y_TEMP_VAL
    PUB_BK_ARR = TEMP_ARR
    Z_TEMP_VAL = BRENT_ZERO_FUNC(BRENT_MIN_VAL, BRENT_MAX_VAL, "BK_FIT_OBJ_FUNC", CDbl(Z_TEMP_VAL), , , BRENT_nLOOPS, BRENT_TOLER)
    FIT_ARR(i) = Z_TEMP_VAL
Next i

Exit Function
ERROR_LABEL:
BK_FIT_RATE_FUNC = Err.number
End Function

'// PERFECT


'************************************************************************************
'************************************************************************************



'************************************************************************************
'************************************************************************************
'FUNCTION      : BK_FIT_OBJ_FUNC
'DESCRIPTION   : Objective function for fitting values in the Black-Karansinki
'LIBRARY       : FIXED_INCOME
'GROUP         : HW_2F
'ID            : 010
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 10-06-2008
'************************************************************************************
'************************************************************************************

Private Function BK_FIT_OBJ_FUNC(ByVal THETA_VAL As Double)
  
  Dim i As Long
  Dim TEMP_DISC As Double
  Dim TEMP_SUM As Double
  Dim TEMP_RESID As Double
  
  On Error GoTo ERROR_LABEL
  
  TEMP_RESID = PUB_BK_PRICE
  TEMP_SUM = PUB_BK_MIN
  
  For i = 0 To PUB_BK_SIZE - 1
    TEMP_DISC = Exp(-Exp(THETA_VAL + TEMP_SUM) * PUB_BK_DELTA)
    TEMP_RESID = TEMP_RESID - PUB_BK_ARR(i) * TEMP_DISC
    TEMP_SUM = TEMP_SUM + PUB_BK_GRAD
  Next i
  
  BK_FIT_OBJ_FUNC = TEMP_RESID
  
Exit Function
ERROR_LABEL:
  BK_FIT_OBJ_FUNC = Err.number
End Function

'// PERFECT
 
'************************************************************************************
'************************************************************************************



'************************************************************************************
'************************************************************************************
'FUNCTION      : HW_CAPLET_FUNC
'DESCRIPTION   : Caplet Function
'LIBRARY       : FIXED_INCOME
'GROUP         : HW_2F
'ID            : 011
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 10-06-2008
'************************************************************************************
'************************************************************************************


Private Function HW_CAPLET_FUNC(ByVal FIXING As Double, _
ByVal FORWARD As Double, _
ByVal STRIKE As Double, _
ByVal SIGMA As Double)
  
On Error GoTo ERROR_LABEL
  
  If FIXING < 0 Then ' the rate was fixed
    HW_CAPLET_FUNC = MAXIMUM_FUNC(FORWARD - STRIKE, 0)
  Else ' forecast numerical inaccuracies can yield a negative answer
    HW_CAPLET_FUNC = MAXIMUM_FUNC(HW_BLACK_FUNC(FORWARD, _
                      STRIKE, SIGMA * FIXING ^ 0.5, 1), 0)
  End If

Exit Function
ERROR_LABEL:
    HW_CAPLET_FUNC = Err.number
End Function

'// PERFECT

'************************************************************************************
'************************************************************************************



'************************************************************************************
'************************************************************************************
'FUNCTION      : HW_BLACK_FUNC
'DESCRIPTION   : Black-Price function
'LIBRARY       : FIXED_INCOME
'GROUP         : HW_2F
'ID            : 012
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 10-06-2008
'************************************************************************************
'************************************************************************************

Private Function HW_BLACK_FUNC(ByVal FORWARD As Double, _
ByVal STRIKE As Double, _
ByVal variance As Double, _
ByVal WEIGHT As Double)
  
  Dim D1_VAL As Double
  Dim D2_VAL As Double
  Const CND_TYPE As Integer = 0
  On Error GoTo ERROR_LABEL

  If Abs(variance) < 2 ^ -52 Then
    HW_BLACK_FUNC = MAXIMUM_FUNC(FORWARD * WEIGHT - STRIKE * WEIGHT, 0)
    Exit Function
  End If
  
  D1_VAL = (Log(FORWARD / STRIKE) / Log(Exp(1))) / variance + 0.5 * variance
  D2_VAL = D1_VAL - variance
  
  HW_BLACK_FUNC = FORWARD * WEIGHT * CND_FUNC(WEIGHT * D1_VAL, CND_TYPE) - _
                    STRIKE * WEIGHT * CND_FUNC(WEIGHT * D2_VAL, CND_TYPE)

Exit Function
ERROR_LABEL:
    HW_BLACK_FUNC = Err.number
End Function

'// PERFECT

'************************************************************************************
'************************************************************************************



'************************************************************************************
'************************************************************************************
'FUNCTION      : HW_TRINOM_TREE_FUNC
'DESCRIPTION   : HW Trinomial Tree Function
'LIBRARY       : FIXED_INCOME
'GROUP         : HW_2F
'ID            : 013
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 10-06-2008
'************************************************************************************
'************************************************************************************

Private Function HW_TRINOM_TREE_FUNC(ByVal INIT_STEP_VAL As Double, _
ByVal SIGMA_VAL As Double, _
ByVal SPEED_VAL As Double, _
ByRef TIME_ARR As Variant, _
Optional ByVal epsilon As Double = 2 ^ 52)

Dim i As Long
Dim j As Long
Dim k As Long
Dim l As Long

Dim nSTEPS As Long

Dim MIN_VAL As Double
Dim MAX_VAL As Double

Dim PU_ARR As Variant
Dim PD_ARR As Variant
Dim PM_ARR As Variant

Dim TEMP_V1 As Double
Dim TEMP_V2 As Double
Dim DELTA_TENOR As Double

Dim ATEMP_VAL As Double
Dim BTEMP_VAL As Double
Dim CTEMP_VAL As Double
Dim DTEMP_VAL As Double

Dim ATEMP_MULT As Double
Dim BTEMP_MULT As Double
Dim CTEMP_MULT As Double
Dim DTEMP_MULT As Double

Dim TEMP_GROUP As Variant
Dim TEMP_ARR As Variant

Dim GRAD_ARR As Variant

Dim PROB_INDEX_ARR As Variant
Dim PROB_DOWN_ARR As Variant
Dim PROB_UP_ARR As Variant
Dim PROB_MED_ARR As Variant

On Error GoTo ERROR_LABEL

'----------------------------SET TRINOMIAL TREE----------------------------
  
  nSTEPS = UBound(TIME_ARR, 1)
  
  ReDim GRAD_ARR(0 To nSTEPS)
'  ReDim PROB_INDEX_ARR(0 To nSTEPS - 1)
  ReDim PROB_INDEX_ARR(0 To nSTEPS)
  ReDim PROB_DOWN_ARR(0 To nSTEPS - 1)
  ReDim PROB_UP_ARR(0 To nSTEPS - 1)
  ReDim PROB_MED_ARR(0 To nSTEPS - 1)

  MIN_VAL = 0
  MAX_VAL = 0
  GRAD_ARR(0) = 0
  
  For i = 0 To nSTEPS - 1
    
    DELTA_TENOR = TIME_ARR(i + 1) - TIME_ARR(i)
    
    'calculate variance and dx for each timestep
    TEMP_V2 = 0.5 * SIGMA_VAL * SIGMA_VAL / SPEED_VAL _
            * (1 - Exp(-2 * SPEED_VAL * DELTA_TENOR))
    
    TEMP_V1 = TEMP_V2 ^ 0.5
    GRAD_ARR(i + 1) = TEMP_V1 * 3 ^ 0.5

    ReDim PU_ARR(0 To MAX_VAL - MIN_VAL)
    ReDim PM_ARR(0 To MAX_VAL - MIN_VAL)
    ReDim PD_ARR(0 To MAX_VAL - MIN_VAL)
    ReDim TEMP_ARR(0 To MAX_VAL - MIN_VAL)
    
    l = 0
    For j = MIN_VAL To MAX_VAL
      ATEMP_VAL = INIT_STEP_VAL + j * GRAD_ARR(i)
      'calculate conditional mean and k at each node
      ATEMP_MULT = ATEMP_VAL * Exp(-SPEED_VAL * DELTA_TENOR)
      
      BTEMP_VAL = (ATEMP_MULT - INIT_STEP_VAL) / GRAD_ARR(i + 1)
      CTEMP_VAL = Sgn(BTEMP_VAL)
      DTEMP_VAL = FLOOR_FUNC(BTEMP_VAL + CTEMP_VAL * 0.5, CTEMP_VAL)
      TEMP_ARR(l) = DTEMP_VAL

      BTEMP_MULT = ATEMP_MULT - (INIT_STEP_VAL + DTEMP_VAL * GRAD_ARR(i + 1))
      CTEMP_MULT = BTEMP_MULT ^ 2
      DTEMP_MULT = BTEMP_MULT * 3 ^ 0.5
      
      PD_ARR(l) = (1 + CTEMP_MULT / TEMP_V2 - DTEMP_MULT / TEMP_V1) / 6
      PM_ARR(l) = (2 - CTEMP_MULT / TEMP_V2) / 3
      PU_ARR(l) = (1 + CTEMP_MULT / TEMP_V2 + DTEMP_MULT / TEMP_V1) / 6
      
      l = l + 1
    Next j
    PROB_DOWN_ARR(i) = PD_ARR
    PROB_MED_ARR(i) = PM_ARR
    PROB_UP_ARR(i) = PU_ARR
    PROB_INDEX_ARR(i) = TEMP_ARR
    
    MIN_VAL = epsilon
    For k = LBound(TEMP_ARR, 1) To UBound(TEMP_ARR, 1)
       If TEMP_ARR(k) < MIN_VAL Then: MIN_VAL = TEMP_ARR(k)
    Next k
    MIN_VAL = MIN_VAL - 1
    
    MAX_VAL = -1 * epsilon
    For k = LBound(TEMP_ARR, 1) To UBound(TEMP_ARR, 1)
       If TEMP_ARR(k) > MAX_VAL Then: MAX_VAL = TEMP_ARR(k)
    Next k

    MAX_VAL = MAX_VAL + 1
  Next i

  ReDim TEMP_GROUP(1 To 5)
  
  TEMP_GROUP(1) = GRAD_ARR
  TEMP_GROUP(2) = PROB_INDEX_ARR
  TEMP_GROUP(3) = PROB_DOWN_ARR
  TEMP_GROUP(4) = PROB_UP_ARR
  TEMP_GROUP(5) = PROB_MED_ARR
  
  HW_TRINOM_TREE_FUNC = TEMP_GROUP
  
Exit Function
ERROR_LABEL:
    HW_TRINOM_TREE_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : HW_BRANCH_PROB_FUNC
'DESCRIPTION   : Calculate branching probabilities from time at each node
'LIBRARY       : FIXED_INCOME
'GROUP         : HW_2F
'ID            : 014
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 10-06-2008
'************************************************************************************
'************************************************************************************

Private Function HW_BRANCH_PROB_FUNC(ByVal i As Long, _
ByVal j As Long, _
ByVal k As Long, _
ByVal l As Long, _
ByVal RHO_VAL As Double, _
ByRef FIRST_GROUP As Variant, _
ByRef SECOND_GROUP As Variant, _
ByRef POS_ADJ_RHO_ARR As Variant, _
ByRef NEG_ADJ_RHO_ARR As Variant, _
Optional ByVal RHO_FACTOR As Double = 36)

'l --> Branch
  Dim ii As Long
  Dim jj As Long
  Dim iii As Double
  Dim jjj As Double
  
  On Error GoTo ERROR_LABEL

  ii = FLOOR_FUNC(l / 3, 1)
  jj = l Mod 3
  iii = HW_PROB_INDEX_FUNC(i, j, ii, FIRST_GROUP(3), _
        FIRST_GROUP(4), FIRST_GROUP(5))
  
  jjj = HW_PROB_INDEX_FUNC(i, k, jj, SECOND_GROUP(3), _
        SECOND_GROUP(4), SECOND_GROUP(5))
  
  If RHO_VAL = 0 Then
    HW_BRANCH_PROB_FUNC = iii * jjj
  ElseIf RHO_VAL > 0 Then
    HW_BRANCH_PROB_FUNC = iii * jjj _
      + (RHO_VAL / RHO_FACTOR) * POS_ADJ_RHO_ARR(l)
  Else
    HW_BRANCH_PROB_FUNC = iii * jjj _
      - (RHO_VAL / RHO_FACTOR) * NEG_ADJ_RHO_ARR(l)
  End If

Exit Function
ERROR_LABEL:
    HW_BRANCH_PROB_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : HW_PROB_INDEX_FUNC
'DESCRIPTION   : HW probability control function
'LIBRARY       : FIXED_INCOME
'GROUP         : HW_2F
'ID            : 015
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 10-06-2008
'************************************************************************************
'************************************************************************************

Private Function HW_PROB_INDEX_FUNC(ByVal i As Long, _
ByVal j As Long, _
ByVal k As Long, _
ByRef PROB_DOWN_ARR As Variant, _
ByRef PROB_UP_ARR As Variant, _
ByRef PROB_MED_ARR As Variant)
  
  Dim TEMP_ARR As Variant
  
  On Error GoTo ERROR_LABEL
  
  If k = 0 Then
    TEMP_ARR = PROB_DOWN_ARR(i)
  ElseIf k = 1 Then
    TEMP_ARR = PROB_MED_ARR(i)
  ElseIf k = 2 Then
    TEMP_ARR = PROB_UP_ARR(i)
  End If
  HW_PROB_INDEX_FUNC = TEMP_ARR(j)

Exit Function
ERROR_LABEL:
    HW_PROB_INDEX_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : HW_NODE_TREE_FUNC
'DESCRIPTION   : Returns the no of tree nodes for a timestep (size)
'LIBRARY       : FIXED_INCOME
'GROUP         : HW_2F
'ID            : 016
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 10-06-2008
'************************************************************************************
'************************************************************************************

Private Function HW_NODE_TREE_FUNC(ByVal i As Long, _
ByRef PROB_INDEX_ARR As Variant, _
Optional ByVal epsilon As Double = 2 ^ 52)
  
  Dim j As Long
  Dim MIN_VAL As Double
  Dim MAX_VAL As Double
  Dim TEMP_ARR As Variant
  
  On Error GoTo ERROR_LABEL
  
  If i = 0 Then
    HW_NODE_TREE_FUNC = 1
    Exit Function
  End If
  TEMP_ARR = PROB_INDEX_ARR(i - 1)
  
  MIN_VAL = epsilon
  For j = LBound(TEMP_ARR, 1) To UBound(TEMP_ARR, 1)
    If TEMP_ARR(j) < MIN_VAL Then: MIN_VAL = TEMP_ARR(j)
  Next j

  MIN_VAL = MIN_VAL - 1
    
  MAX_VAL = -1 * epsilon
  For j = LBound(TEMP_ARR, 1) To UBound(TEMP_ARR, 1)
    If TEMP_ARR(j) > MAX_VAL Then: MAX_VAL = TEMP_ARR(j)
  Next j
  
  MAX_VAL = MAX_VAL + 1
  HW_NODE_TREE_FUNC = MAX_VAL - MIN_VAL + 1

Exit Function
ERROR_LABEL:
    HW_NODE_TREE_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : HW_NODE_RATE_FUNC
'DESCRIPTION   : Returns value of interest rate at the node underlying
'LIBRARY       : FIXED_INCOME
'GROUP         : HW_2F
'ID            : 017
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 10-06-2008
'************************************************************************************
'************************************************************************************

Private Function HW_NODE_RATE_FUNC(ByVal i As Long, _
ByVal INDEX_VAL As Double, _
ByVal INIT_STEP_VAL As Double, _
ByRef GRAD_ARR As Variant, _
ByRef PROB_INDEX_ARR As Variant, _
Optional ByVal epsilon As Double = 2 ^ 52)

  Dim j As Long
  Dim MIN_VAL As Double
  Dim TEMP_ARR As Variant
  
  On Error GoTo ERROR_LABEL
  
  If i = 0 Then
    HW_NODE_RATE_FUNC = INIT_STEP_VAL
    Exit Function
  End If
  TEMP_ARR = PROB_INDEX_ARR(i - 1)
  
  MIN_VAL = epsilon
  For j = LBound(TEMP_ARR, 1) To UBound(TEMP_ARR, 1)
    If TEMP_ARR(j) < MIN_VAL Then: MIN_VAL = TEMP_ARR(j)
  Next j
  
  MIN_VAL = MIN_VAL - 1
  HW_NODE_RATE_FUNC = INIT_STEP_VAL + (MIN_VAL + INDEX_VAL) * GRAD_ARR(i)

Exit Function
ERROR_LABEL:
    HW_NODE_RATE_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : HW_ON_TIME_FUNC
'DESCRIPTION   : Time Grid function
'LIBRARY       : FIXED_INCOME
'GROUP         : HW_2F
'ID            : 018
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 10-06-2008
'************************************************************************************
'************************************************************************************

Private Function HW_ON_TIME_FUNC(ByVal XTEMP_VAL As Double, _
ByVal TENOR_VAL As Double, _
ByRef TENOR_ARR As Variant, _
Optional ByVal tolerance As Double = 0.01)
  
  Dim TEMP_GRID As Double
  
  On Error GoTo ERROR_LABEL
  
  TEMP_GRID = TENOR_ARR(HW_FIND_INDEX_FUNC(XTEMP_VAL, TENOR_ARR, tolerance))
  XTEMP_VAL = TENOR_VAL - TEMP_GRID
  If tolerance = 0 Then: tolerance = 0.000000000000001
  If Abs(XTEMP_VAL) < tolerance Then
        HW_ON_TIME_FUNC = True
  Else: HW_ON_TIME_FUNC = False
  End If

Exit Function
ERROR_LABEL:
    HW_ON_TIME_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : HW_FIND_INDEX_FUNC
'DESCRIPTION   : Time Index Function
'LIBRARY       : FIXED_INCOME
'GROUP         : HW_2F
'ID            : 019
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 10-06-2008
'************************************************************************************
'************************************************************************************

Private Function HW_FIND_INDEX_FUNC(ByVal XTEMP_VAL As Double, _
ByRef TENOR_ARR As Variant, _
Optional ByVal tolerance As Double = 0.01)
  
  Dim i As Long
  Dim TEMP_ERR As Double
  
  On Error GoTo ERROR_LABEL
  '----------------------------------------------------------------------------------
  Select Case tolerance
'----------------------------------------------------------------------------------
'----------------------------------------------------------------------------------
    Case 0 'Yield Calibration Function
'----------------------------------------------------------------------------------
      For i = 0 To UBound(TENOR_ARR, 1)
        If XTEMP_VAL < TENOR_ARR(i) Then
            HW_FIND_INDEX_FUNC = i - 1
            Exit Function
        End If
      Next i
      HW_FIND_INDEX_FUNC = UBound(TENOR_ARR, 1)
'----------------------------------------------------------------------------------
    Case Else 'Bermudean Function
'----------------------------------------------------------------------------------
      For i = 0 To UBound(TENOR_ARR, 1)
        TEMP_ERR = Abs(XTEMP_VAL - TENOR_ARR(i))
        If TEMP_ERR < tolerance Then
          HW_FIND_INDEX_FUNC = i
          Exit Function
        End If
      Next i
      HW_FIND_INDEX_FUNC = -1
'----------------------------------------------------------------------------------
  End Select
'----------------------------------------------------------------------------------
  
Exit Function
ERROR_LABEL:
    HW_FIND_INDEX_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : HW_DESCEND_FUNC
'DESCRIPTION   : Descendent function
'LIBRARY       : FIXED_INCOME
'GROUP         : HW_2F
'ID            : 020
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 10-06-2008
'************************************************************************************
'************************************************************************************

Private Function HW_DESCEND_FUNC(ByVal i As Long, _
ByVal j As Long, _
ByVal INDEX_VAL As Double, _
ByRef PROB_INDEX_ARR As Variant, _
Optional ByVal epsilon As Double = 2 ^ 52)

  Dim k As Long
  Dim MIN_VAL As Double
  Dim TEMP_ARR As Variant
  
  On Error GoTo ERROR_LABEL
  
  TEMP_ARR = PROB_INDEX_ARR(i)
  
  MIN_VAL = epsilon
  For k = LBound(TEMP_ARR, 1) To UBound(TEMP_ARR, 1)
    If TEMP_ARR(k) < MIN_VAL Then: MIN_VAL = TEMP_ARR(k)
  Next k

  MIN_VAL = MIN_VAL - 1
  HW_DESCEND_FUNC = TEMP_ARR(j) - MIN_VAL - 1 + INDEX_VAL

Exit Function
ERROR_LABEL:
    HW_DESCEND_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : HW_PRE_ADJ_FUNC
'DESCRIPTION   :
'LIBRARY       : FIXED_INCOME
'GROUP         : HW_2F
'ID            : 021
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 10-06-2008
'************************************************************************************
'************************************************************************************

Private Function HW_PRE_ADJ_FUNC(ByVal INIT_STEP_VAL As Double, _
ByVal NOMINAL As Double, _
ByVal CAP_VAL As Double, _
ByRef TENOR_VAL As Double, _
ByRef TENOR_ARR As Variant, _
ByRef START_TENOR_ARR As Variant, _
ByRef END_TENOR_ARR As Variant, _
ByRef ACCRUED_ARR As Variant, _
ByRef VALUES_ARR As Variant, _
ByRef FIT_ARR As Variant, _
ByRef GRAD_ARR As Variant, _
ByRef PROB_INDEX_ARR As Variant, _
ByRef PROB_DOWN_ARR As Variant, _
ByRef PROB_UP_ARR As Variant, _
ByRef PROB_MED_ARR As Variant, _
Optional ByVal epsilon As Double = 2 ^ 52, _
Optional ByVal VERSION As Integer = 0)

Dim i As Long
Dim j As Long

Dim ii As Long
Dim jj As Long

Dim SROW As Long
Dim NROWS As Long

Dim TEMP_END As Double
Dim TEMP_START As Double

Dim TO_VAL As Double
Dim FROM_VAL As Double

Dim TEMP_STRIKE As Double
Dim TEMP_ACCRUED As Double

Dim ATEMP_ARR As Variant
Dim BTEMP_ARR As Variant
Dim CTEMP_ARR As Variant

On Error GoTo ERROR_LABEL

For i = 0 To UBound(START_TENOR_ARR, 1)
  TEMP_START = START_TENOR_ARR(i)
  If (HW_ON_TIME_FUNC(TEMP_START, TENOR_VAL, TENOR_ARR, 0)) Then
      TEMP_END = END_TENOR_ARR(i)
      TO_VAL = ACCRUED_ARR(i)

      jj = HW_NODE_TREE_FUNC(HW_FIND_INDEX_FUNC(TEMP_END, TENOR_ARR, 0), _
            PROB_INDEX_ARR, epsilon)
      ReDim CTEMP_ARR(0 To jj - 1)
  
      For j = 0 To jj - 1
        CTEMP_ARR(j) = 1
      Next j
      
  FROM_VAL = TEMP_END
  If (FROM_VAL < TENOR_VAL) Then
    Debug.Print "Lattice: cannot roll the asset back to" & _
            TENOR_VAL & " (it is already at t = " & FROM_VAL
    GoTo 1985
  End If
  
  If (FROM_VAL > TENOR_VAL) Then
    NROWS = HW_FIND_INDEX_FUNC(FROM_VAL, TENOR_ARR, 0)
    SROW = HW_FIND_INDEX_FUNC(TENOR_VAL, TENOR_ARR, 0)
    For ii = NROWS - 1 To SROW Step -1
      ReDim ATEMP_ARR(0 To HW_NODE_TREE_FUNC(ii, PROB_INDEX_ARR, epsilon) - 1)
      BTEMP_ARR = CTEMP_ARR
      Call HW_STEP_BACK_FUNC(ii, INIT_STEP_VAL, GRAD_ARR, TENOR_ARR, FIT_ARR, _
                BTEMP_ARR, ATEMP_ARR, PROB_INDEX_ARR, _
                PROB_DOWN_ARR, PROB_UP_ARR, _
                PROB_MED_ARR, epsilon, 0, VERSION)
      CTEMP_ARR = ATEMP_ARR
    'skip the very last post-adjustment
      'If VERSION <> 0 Then
        If (ii <> SROW) Then
            'Asset PreAdjustValues Routine here
            'Asset PostAdjustValues Routine here
        Else
            'Asset PreAdjustValues Routine here
        End If
      'End If
    Next ii
  End If
1985:
  '-----------------------Post Adjust Values
      TEMP_ACCRUED = 1 + CAP_VAL * TO_VAL
      TEMP_STRIKE = 1 / TEMP_ACCRUED
      For j = 0 To UBound(VALUES_ARR, 1)
        VALUES_ARR(j) = VALUES_ARR(j) + NOMINAL * TEMP_ACCRUED * _
                        MAXIMUM_FUNC(TEMP_STRIKE - CTEMP_ARR(j), 0)
      Next j
   
  End If
Next i

Exit Function
ERROR_LABEL:
    HW_PRE_ADJ_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : HW_FAIR_RATE_FUNC
'DESCRIPTION   :
'LIBRARY       : FIXED_INCOME
'GROUP         : HW_2F
'ID            : 022
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 10-06-2008
'************************************************************************************
'************************************************************************************

Private Function HW_FAIR_RATE_FUNC(ByVal START_TENOR_VAL As Double, _
ByVal END_TENOR_VAL As Double, _
ByVal NOMINAL As Double, _
ByVal FIXED_RATE As Double, _
ByRef START_TENOR_ARR As Variant, _
ByRef END_TENOR_ARR As Variant, _
ByRef ACCRU_ARR As Variant, _
ByRef FIX_ARR As Variant, _
ByRef FORW_ARR As Variant, _
ByRef RF_DISC_ARR As Variant, _
ByRef RF_RATE_ARR As Variant, _
ByRef RF_TENOR_ARR As Variant)
    
  Dim i As Long
  Dim j As Long
  
  Dim NSIZE As Long
  
  Dim END_DISC As Double
  Dim START_DISC As Double
   
  Dim TEMP_TENOR As Double
 
  Dim NPV_VAL As Double
  Dim LEG_BPS_VAL As Double
  Dim FLOAT_COUP_VAL As Double
  Dim FIXED_COUP_VAL As Double
  
  On Error GoTo ERROR_LABEL

  TEMP_TENOR = START_TENOR_VAL
  NSIZE = (END_TENOR_VAL - START_TENOR_VAL) / START_TENOR_VAL
  
  ReDim START_TENOR_ARR(0 To NSIZE - 1)
  ReDim ACCRU_ARR(0 To NSIZE - 1)
  ReDim END_TENOR_ARR(0 To NSIZE - 1)
  ReDim FIX_ARR(0 To NSIZE - 1)
  ReDim FORW_ARR(0 To NSIZE - 1)
  
  NSIZE = 0
  
Do While TEMP_TENOR < END_TENOR_VAL
    
      j = 0
      For i = 0 To UBound(RF_TENOR_ARR, 1)
        If TEMP_TENOR < RF_TENOR_ARR(i) Then
          j = i - 1
          GoTo 1983
        End If
      Next i
1983:
    FORW_ARR(NSIZE) = RF_RATE_ARR(j) 'Inst. Forward
    START_TENOR_ARR(NSIZE) = TEMP_TENOR
    FIX_ARR(NSIZE) = TEMP_TENOR
    END_TENOR_ARR(NSIZE) = TEMP_TENOR + START_TENOR_VAL
    ACCRU_ARR(NSIZE) = START_TENOR_VAL
    NSIZE = NSIZE + 1
    TEMP_TENOR = TEMP_TENOR + START_TENOR_VAL
Loop
  
  TEMP_TENOR = START_TENOR_VAL
  
  NPV_VAL = 0
  LEG_BPS_VAL = 0
  
Do While TEMP_TENOR <= END_TENOR_VAL
    
    START_DISC = HW_ZERO_RATE_FUNC(TEMP_TENOR, RF_TENOR_ARR, _
                    RF_DISC_ARR, RF_RATE_ARR)
    
    END_DISC = HW_ZERO_RATE_FUNC(TEMP_TENOR + START_TENOR_VAL, RF_TENOR_ARR, _
                    RF_DISC_ARR, RF_RATE_ARR)
    
    FLOAT_COUP_VAL = (START_DISC / END_DISC - 1) * NOMINAL
    FIXED_COUP_VAL = START_TENOR_VAL * FIXED_RATE * NOMINAL
    
    NPV_VAL = NPV_VAL - FLOAT_COUP_VAL * HW_ZERO_RATE_FUNC(TEMP_TENOR, _
                RF_TENOR_ARR, RF_DISC_ARR, RF_RATE_ARR)
    
    NPV_VAL = NPV_VAL + FIXED_COUP_VAL * HW_ZERO_RATE_FUNC(TEMP_TENOR, _
                RF_TENOR_ARR, RF_DISC_ARR, RF_RATE_ARR)
    
    LEG_BPS_VAL = LEG_BPS_VAL + START_TENOR_VAL * NOMINAL * _
                HW_ZERO_RATE_FUNC(TEMP_TENOR, RF_TENOR_ARR, _
                RF_DISC_ARR, RF_RATE_ARR)
    
    TEMP_TENOR = TEMP_TENOR + START_TENOR_VAL
  
Loop

  HW_FAIR_RATE_FUNC = FIXED_RATE - NPV_VAL / LEG_BPS_VAL 'GetFairRate

Exit Function
ERROR_LABEL:
    HW_FAIR_RATE_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : HW_SURVIVAL_TREE_FUNC
'DESCRIPTION   :
'LIBRARY       : FIXED_INCOME
'GROUP         : HW_2F
'ID            : 023
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 10-06-2008
'************************************************************************************
'************************************************************************************

Private Function HW_SURVIVAL_TREE_FUNC(ByVal INIT_STEP_VAL As Double, _
ByVal NSIZE As Long, _
ByVal SROW As Long, _
ByVal RHO_VAL As Double, _
ByRef TIME_ARR As Variant, _
ByRef RF_PRICE_ARR As Variant, _
ByRef RF_DISC_ARR As Variant, _
ByRef RF_RATE_ARR As Variant, _
ByRef RF_TENOR_ARR As Variant, _
ByRef RISKY_PRICE_ARR As Variant, _
ByRef RISKY_RATE_ARR As Variant, _
ByRef RISKY_TENOR_ARR As Variant, _
ByRef FIRST_GROUP As Variant, _
ByRef FIRST_FIT_ARR As Variant, _
ByRef FIRST_PRICE_ARR As Variant, _
ByRef SECOND_GROUP As Variant, _
ByRef SECOND_FIT_ARR As Variant, _
ByRef SECOND_PRICE_ARR As Variant, _
ByRef POS_ADJ_RHO_ARR As Variant, _
ByRef NEG_ADJ_RHO_ARR As Variant, _
Optional ByVal RHO_FACTOR As Double = 36, _
Optional ByVal epsilon As Double = 2 ^ 52)

  Dim i As Long
  Dim j As Long
  Dim k As Long
  
  Dim ii As Long
  Dim jj As Long
  
  Dim INDEX_VAL As Double
  
  Dim X1_VAL As Double
  Dim X2_VAL As Double
  
  Dim DX1_VAL As Double
  Dim DX2_VAL As Double
  
  Dim FIRST_VAL As Double
  Dim SECOND_VAL As Double
  Dim THIRD_VAL As Double
  
  Dim TEMP_SUM As Double
  Dim TEMP_DELTA As Double
  Dim TEMP_MULT As Double
  
  Dim FIRST_FACTOR As Double
  Dim SECOND_FACTOR As Double
  
  Dim RATE_ARR As Variant
  Dim DISC_ARR As Variant
  Dim RATES_ARR As Variant
  Dim TENOR_ARR As Variant
  
  Dim TEMP_ARR As Variant
  Dim STATE_PRICES_ARR As Variant
  
  Dim RISKY_DISC_BOND As Double 'Variant
  Dim RF_DISC_ARR_BOND As Double 'Variant
  
  On Error GoTo ERROR_LABEL

  TEMP_ARR = SECOND_FIT_ARR
  ReDim TEMP_ARR(0 To UBound(TIME_ARR, 1) - 1)
  SECOND_FIT_ARR = TEMP_ARR 'Default Fitting Values
  TENOR_ARR = RISKY_TENOR_ARR
  RATE_ARR = RISKY_RATE_ARR
  
  ReDim DISC_ARR(0 To UBound(RISKY_TENOR_ARR, 1))
  ReDim RATES_ARR(0 To UBound(RISKY_TENOR_ARR, 1))
  DISC_ARR(0) = 1
  RATES_ARR(0) = RISKY_RATE_ARR(0)
  
  For k = 1 To UBound(RISKY_TENOR_ARR, 1)
    RATES_ARR(k) = (RISKY_RATE_ARR(k) * (RISKY_TENOR_ARR(k) - _
                            RISKY_TENOR_ARR(k - 1)) + RATES_ARR(k - 1) _
                            * RISKY_TENOR_ARR(k - 1)) / RISKY_TENOR_ARR(k)
    DISC_ARR(k) = Exp(-RATES_ARR(k) * TENOR_ARR(k))
  Next k
  
  
  For k = 0 To UBound(TIME_ARR, 1) - 1
    
    INDEX_VAL = TIME_ARR(k + 1)
    RISKY_DISC_BOND = HW_ZERO_RATE_FUNC(INDEX_VAL, TENOR_ARR, _
                        DISC_ARR, RATE_ARR)
    RF_DISC_ARR_BOND = HW_ZERO_RATE_FUNC(INDEX_VAL, RF_TENOR_ARR, RF_DISC_ARR, _
                            RF_RATE_ARR)
    'survival discount adj --> RISKY_DISC_BOND / RF_DISC_ARR_BOND
    
    If k > SROW Then
        SROW = HW_LIMIT_INDEX_FUNC(INIT_STEP_VAL, NSIZE, SROW, k, RHO_VAL, _
                        TIME_ARR, RISKY_PRICE_ARR, RF_PRICE_ARR, _
                        FIRST_GROUP, FIRST_FIT_ARR, FIRST_PRICE_ARR, _
                        SECOND_GROUP, SECOND_FIT_ARR, SECOND_PRICE_ARR, _
                        POS_ADJ_RHO_ARR, NEG_ADJ_RHO_ARR, RHO_FACTOR, epsilon)
    End If
    
    STATE_PRICES_ARR = RISKY_PRICE_ARR(k)
    
    ii = HW_NODE_TREE_FUNC(k, FIRST_GROUP(2), epsilon)
    jj = HW_NODE_TREE_FUNC(k, SECOND_GROUP(2), epsilon)
    
    TEMP_DELTA = TIME_ARR(k + 1) - TIME_ARR(k)
    DX1_VAL = FIRST_GROUP(1)(k)
    DX2_VAL = SECOND_GROUP(1)(k)
    
    X1_VAL = HW_NODE_RATE_FUNC(k, 0, INIT_STEP_VAL, FIRST_GROUP(1), _
            FIRST_GROUP(2), epsilon)
    X2_VAL = HW_NODE_RATE_FUNC(k, 0, INIT_STEP_VAL, SECOND_GROUP(1), _
            SECOND_GROUP(2), epsilon)
    
    TEMP_SUM = 0
    For i = 0 To ii - 1
      X2_VAL = HW_NODE_RATE_FUNC(k, 0, INIT_STEP_VAL, SECOND_GROUP(1), _
                SECOND_GROUP(2), epsilon)
      For j = 0 To jj - 1
        TEMP_MULT = HW_DISC_FUNC(k, i, INIT_STEP_VAL, FIRST_GROUP(1), TIME_ARR, _
                            FIRST_FIT_ARR, FIRST_GROUP(2), epsilon, 0, 0)
        TEMP_SUM = TEMP_SUM + STATE_PRICES_ARR(i, j) _
                        * TEMP_MULT * Exp(-X2_VAL * TEMP_DELTA)
        X2_VAL = X2_VAL + DX2_VAL
      Next j
      X1_VAL = X1_VAL + DX1_VAL
    Next i
    TEMP_SUM = (Log(TEMP_SUM / RISKY_DISC_BOND) / Log(Exp(1))) / TEMP_DELTA
    SECOND_FIT_ARR(k) = TEMP_SUM
    
'-----------------------------------------------------------------------------------------------
    TEMP_SUM = 0
    For i = 0 To ii - 1
      X2_VAL = HW_NODE_RATE_FUNC(k, 0, INIT_STEP_VAL, SECOND_GROUP(1), _
                        SECOND_GROUP(2), epsilon)
      For j = 0 To jj - 1
        TEMP_MULT = HW_DISC_FUNC(k, i, INIT_STEP_VAL, FIRST_GROUP(1), _
                        TIME_ARR, FIRST_FIT_ARR, FIRST_GROUP(2), epsilon, 0, 0)
        TEMP_SUM = TEMP_SUM + STATE_PRICES_ARR(i, j) _
                        * TEMP_MULT * HW_DISC_FUNC(k, j, INIT_STEP_VAL, SECOND_GROUP(1), _
                        TIME_ARR, SECOND_FIT_ARR, SECOND_GROUP(2), epsilon, 0, 0)
        X2_VAL = X2_VAL + DX2_VAL
      Next j
      X1_VAL = X1_VAL + DX1_VAL
    Next i
    
    FIRST_FACTOR = FLOOR_FUNC(ii / 2, 1)
    SECOND_FACTOR = FLOOR_FUNC(jj / 2, 1)
    
    FIRST_VAL = HW_DISC_FUNC(k, FIRST_FACTOR, INIT_STEP_VAL, _
                    FIRST_GROUP(1), TIME_ARR, _
                    FIRST_FIT_ARR, FIRST_GROUP(2), epsilon, 0, 0)
    SECOND_VAL = HW_DISC_FUNC(k, SECOND_FACTOR, INIT_STEP_VAL, _
                    SECOND_GROUP(1), TIME_ARR, _
                    SECOND_FIT_ARR, SECOND_GROUP(2), epsilon, 0, 0)
    THIRD_VAL = FIRST_VAL * SECOND_VAL
'-----------------------------------------------------------------------------------------------
Next k

HW_SURVIVAL_TREE_FUNC = SROW

Exit Function
ERROR_LABEL:
    HW_SURVIVAL_TREE_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : HW_STEP_BACK_FUNC
'DESCRIPTION   : Populate new values vector for the discrete asset
'LIBRARY       : FIXED_INCOME
'GROUP         : HW_2F
'ID            : 024
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 10-06-2008
'************************************************************************************
'************************************************************************************

Private Function HW_STEP_BACK_FUNC(ByVal k As Long, _
ByVal INIT_STEP_VAL As Double, _
ByRef GRAD_ARR As Variant, _
ByRef TENOR_ARR As Variant, _
ByRef FIT_ARR As Variant, _
ByRef VALUES_ARR As Variant, _
ByRef DATA_ARR As Variant, _
ByRef PROB_INDEX_ARR As Variant, _
ByRef PROB_DOWN_ARR As Variant, _
ByRef PROB_UP_ARR As Variant, _
ByRef PROB_MED_ARR As Variant, _
Optional ByVal epsilon As Double = 2 ^ 52, _
Optional ByVal tolerance As Double = 0.01, _
Optional ByVal VERSION As Integer = 0)
  
  Dim j As Long
  Dim i As Long
  Dim NROWS As Long
  Dim TEMP_VAL As Double
  
  On Error GoTo ERROR_LABEL
  
  NROWS = HW_NODE_TREE_FUNC(k, PROB_INDEX_ARR, epsilon)
  For j = 0 To NROWS - 1
    TEMP_VAL = 0
    For i = 0 To 2
        TEMP_VAL = TEMP_VAL + HW_PROB_INDEX_FUNC(k, j, i, PROB_DOWN_ARR, _
                PROB_UP_ARR, PROB_MED_ARR) * _
                VALUES_ARR(HW_DESCEND_FUNC(k, j, i, _
                PROB_INDEX_ARR, epsilon))
    Next i
    TEMP_VAL = TEMP_VAL * HW_DISC_FUNC(k, j, INIT_STEP_VAL, GRAD_ARR, TENOR_ARR, _
                            FIT_ARR, PROB_INDEX_ARR, _
                            epsilon, tolerance, VERSION)
    DATA_ARR(j) = TEMP_VAL
  Next j
  
Exit Function
ERROR_LABEL:
    HW_STEP_BACK_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : HW_ROLL_BACK_OBJ_FUNC
'DESCRIPTION   :
'LIBRARY       : FIXED_INCOME
'GROUP         : HW_2F
'ID            : 025
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 10-06-2008
'************************************************************************************
'************************************************************************************

Private Function HW_ROLL_BACK_OBJ_FUNC(ByVal START_TENOR_VAL As Double, _
ByVal END_TENOR_VAL As Double, _
ByVal INIT_STEP_VAL As Double, _
ByRef GRAD_ARR As Variant, _
ByRef TENOR_ARR As Variant, _
ByRef FIT_ARR As Variant, _
ByRef PROB_INDEX_ARR As Variant, _
ByRef PROB_DOWN_ARR As Variant, _
ByRef PROB_UP_ARR As Variant, _
ByRef PROB_MED_ARR As Variant, _
Optional ByRef RESULT_ARR As Variant, _
Optional ByVal epsilon As Double = 2 ^ 52, _
Optional ByVal tolerance As Double = 0.01, _
Optional ByVal VERSION As Integer = 0, _
Optional ByVal METHOD As Integer = 0)
 
 Dim i As Long
 Dim j As Long
 
 Dim SROW As Long
 Dim NROWS As Long
 
 Dim ATEMP_ARR As Variant
 Dim BTEMP_ARR As Variant
 
 On Error GoTo ERROR_LABEL
 
  If (START_TENOR_VAL > END_TENOR_VAL) Then
    NROWS = HW_FIND_INDEX_FUNC(START_TENOR_VAL, TENOR_ARR, tolerance)
    
    If IsArray(RESULT_ARR) = False Then
        SROW = HW_NODE_TREE_FUNC(NROWS, PROB_INDEX_ARR, epsilon)
        ReDim RESULT_ARR(0 To SROW - 1) 'Discount Bond Values
        For j = 0 To SROW - 1
          RESULT_ARR(j) = 1
        Next j
    End If
  
    SROW = HW_FIND_INDEX_FUNC(END_TENOR_VAL, TENOR_ARR, tolerance)
    
    For i = NROWS - 1 To SROW Step -1
      ReDim BTEMP_ARR(0 To HW_NODE_TREE_FUNC(i, PROB_INDEX_ARR, epsilon) - 1)
      ATEMP_ARR = RESULT_ARR
      
      Call HW_STEP_BACK_FUNC(i, INIT_STEP_VAL, GRAD_ARR, TENOR_ARR, FIT_ARR, _
                ATEMP_ARR, BTEMP_ARR, PROB_INDEX_ARR, PROB_DOWN_ARR, _
                PROB_UP_ARR, PROB_MED_ARR, epsilon, tolerance, VERSION)
      
      RESULT_ARR = BTEMP_ARR
      
      If METHOD <> 0 Then
          PUB_EXPIRAT_VAL = TENOR_ARR(i)
          If (i <> SROW) Then 'skip the very last post-adjustment
            'AdjustValues Routines:
               Call HW_PRE_ADJ_FUNC(INIT_STEP_VAL, PUB_NOM_VAL, _
                        PUB_CAP_VAL, PUB_EXPIRAT_VAL, _
                        TENOR_ARR, PUB_START_ARR, PUB_END_ARR, _
                        PUB_ACCR_ARR, RESULT_ARR, FIT_ARR, GRAD_ARR, _
                        PROB_INDEX_ARR, PROB_DOWN_ARR, PROB_UP_ARR, _
                        PROB_MED_ARR, epsilon, VERSION) '  PostAdjustValues

            'Asset PostAdjustValues Routine here
          Else
                Call HW_PRE_ADJ_FUNC(INIT_STEP_VAL, PUB_NOM_VAL, _
                        PUB_CAP_VAL, PUB_EXPIRAT_VAL, _
                        TENOR_ARR, PUB_START_ARR, PUB_END_ARR, _
                        PUB_ACCR_ARR, RESULT_ARR, FIT_ARR, GRAD_ARR, _
                        PROB_INDEX_ARR, PROB_DOWN_ARR, PROB_UP_ARR, _
                        PROB_MED_ARR, epsilon, VERSION) '  PostAdjustValues
          End If
      End If
    Next i
  End If
  
  HW_ROLL_BACK_OBJ_FUNC = RESULT_ARR

Exit Function
ERROR_LABEL:
    HW_ROLL_BACK_OBJ_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : HW_STATE_PRICE_FUNC
'DESCRIPTION   :
'LIBRARY       : FIXED_INCOME
'GROUP         : HW_2F
'ID            : 026
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 10-06-2008
'************************************************************************************
'************************************************************************************

Private Function HW_STATE_PRICE_FUNC(ByVal INIT_STEP_VAL As Double, _
ByVal NROWS As Long, _
ByRef SROW As Long, _
ByRef FIT_ARR As Variant, _
ByRef PRICE_ARR As Variant, _
ByRef GRAD_ARR As Variant, _
ByRef TENOR_ARR As Variant, _
ByRef PROB_INDEX_ARR As Variant, _
ByRef PROB_DOWN_ARR As Variant, _
ByRef PROB_UP_ARR As Variant, _
ByRef PROB_MED_ARR As Variant, _
Optional ByVal epsilon As Double = 2 ^ 52, _
Optional ByVal tolerance As Double = 0.01, _
Optional ByVal VERSION As Integer = 0)

Dim j As Long
Dim i As Long

Dim INDEX_VAL As Double

Dim TEMP_DISC As Double
Dim TEMP_DESC As Double

Dim ATEMP_ARR As Variant
Dim BTEMP_ARR As Variant
Dim CTEMP_ARR As Variant

On Error GoTo ERROR_LABEL

    If NROWS > SROW Then 'Compute State Prices
          For i = SROW To NROWS - 1
            ReDim CTEMP_ARR(0 To HW_NODE_TREE_FUNC(i + 1, PROB_INDEX_ARR, epsilon) - 1)
            
            For j = 0 To HW_NODE_TREE_FUNC(i, PROB_INDEX_ARR, epsilon) - 1
              TEMP_DISC = HW_DISC_FUNC(i, j, INIT_STEP_VAL, GRAD_ARR, TENOR_ARR, _
                            FIT_ARR, PROB_INDEX_ARR, epsilon, _
                            tolerance, VERSION)
              
              ATEMP_ARR = PRICE_ARR(i)
              BTEMP_ARR = ATEMP_ARR(j)
              For INDEX_VAL = 0 To 2
                 TEMP_DESC = HW_DESCEND_FUNC(i, j, INDEX_VAL, PROB_INDEX_ARR, epsilon)
                 CTEMP_ARR(TEMP_DESC) = CTEMP_ARR(TEMP_DESC) + BTEMP_ARR * _
                             TEMP_DISC * HW_PROB_INDEX_FUNC(i, j, INDEX_VAL, _
                                         PROB_DOWN_ARR, PROB_UP_ARR, _
                                         PROB_MED_ARR)
              Next INDEX_VAL
            Next j
            PRICE_ARR(i + 1) = CTEMP_ARR
          Next i
          SROW = NROWS
    End If

Exit Function
ERROR_LABEL:
    HW_STATE_PRICE_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : HW_LIMIT_INDEX_FUNC
'DESCRIPTION   :
'LIBRARY       : FIXED_INCOME
'GROUP         : HW_2F
'ID            : 027
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 10-06-2008
'************************************************************************************
'************************************************************************************

Private Function HW_LIMIT_INDEX_FUNC(ByVal INIT_STEP_VAL As Double, _
ByVal NSIZE As Long, _
ByVal SROW As Long, _
ByVal NROWS As Long, _
ByVal RHO_VAL As Double, _
ByRef TENOR_ARR As Variant, _
ByRef RISKY_PRICE_ARR As Variant, _
ByRef RF_PRICE_ARR As Variant, _
ByRef FIRST_GROUP As Variant, _
ByRef FIRST_FIT_ARR As Variant, _
ByRef FIRST_PRICE_ARR As Variant, _
ByRef SECOND_GROUP As Variant, _
ByRef SECOND_FIT_ARR As Variant, _
ByRef SECOND_PRICE_ARR As Variant, _
ByRef POS_ADJ_RHO_ARR As Variant, _
ByRef NEG_ADJ_RHO_ARR As Variant, _
Optional ByVal RHO_FACTOR As Double = 36, _
Optional ByVal epsilon As Double = 2 ^ 52)

Dim i As Long
Dim j As Long
Dim k As Long
Dim l As Long

Dim nSTEPS As Long

Dim RISKY_DISC As Double
Dim RISK_FREE_DISC As Double
        
Dim RF_MATRIX As Variant
Dim RISKY_MATRIX As Variant

Dim FIRST_TREE_DESC As Double
Dim SECOND_TREE_DESC As Double
        
Dim FIRST_TREE_BRANCH As Double
Dim SECOND_TREE_BRANCH As Double
        
Dim TIME_STEP_RISKY As Double 'Current Time Step State Price Risky
Dim TIME_STEP_RISK_FREE As Double 'Current Time Step State Price Risk Free

Dim NEXT_RISKY_ARR As Variant 'Next Step State Prices Risky
Dim NEXT_RISK_FREE_ARR As Variant ' Next Step State Prices Risk Free
  
On Error GoTo ERROR_LABEL
  
  nSTEPS = UBound(TENOR_ARR, 1)
  
  For i = SROW To NROWS - 1
    ReDim NEXT_RISKY_ARR(0 To _
            HW_NODE_TREE_FUNC(i + 1, FIRST_GROUP(2), epsilon) - 1, _
            0 To HW_NODE_TREE_FUNC(i + 1, SECOND_GROUP(2), epsilon) - 1)
    ReDim NEXT_RISK_FREE_ARR(0 To _
            HW_NODE_TREE_FUNC(i + 1, FIRST_GROUP(2), epsilon) - 1, _
            0 To HW_NODE_TREE_FUNC(i + 1, SECOND_GROUP(2), epsilon) - 1)
    
    For j = 0 To HW_NODE_TREE_FUNC(i, FIRST_GROUP(2), epsilon) - 1
      For k = 0 To HW_NODE_TREE_FUNC(i, SECOND_GROUP(2), epsilon) - 1
        RISK_FREE_DISC = HW_DISC_FUNC(i, j, INIT_STEP_VAL, FIRST_GROUP(1), _
                            TENOR_ARR, FIRST_FIT_ARR, FIRST_GROUP(2), _
                            epsilon, 0, 0)
        RISKY_DISC = HW_DISC_FUNC(i, k, INIT_STEP_VAL, SECOND_GROUP(1), TENOR_ARR, _
                            SECOND_FIT_ARR, SECOND_GROUP(2), epsilon, 0, 0)
        
        RISKY_MATRIX = RISKY_PRICE_ARR(i)
        TIME_STEP_RISKY = RISKY_MATRIX(j, k)
        RF_MATRIX = RF_PRICE_ARR(i)
        TIME_STEP_RISK_FREE = RF_MATRIX(j, k)
        
        For l = 0 To NSIZE
          
        
          FIRST_TREE_BRANCH = FLOOR_FUNC(l / 3, 1)
          SECOND_TREE_BRANCH = l Mod 3
          
          FIRST_TREE_DESC = HW_DESCEND_FUNC(i, j, FIRST_TREE_BRANCH, _
                            FIRST_GROUP(2), epsilon)
          SECOND_TREE_DESC = HW_DESCEND_FUNC(i, k, SECOND_TREE_BRANCH, _
                            SECOND_GROUP(2), epsilon)
          
          NEXT_RISKY_ARR(FIRST_TREE_DESC, SECOND_TREE_DESC) = _
            NEXT_RISKY_ARR(FIRST_TREE_DESC, SECOND_TREE_DESC) _
            + TIME_STEP_RISKY * RISK_FREE_DISC _
            * RISKY_DISC * HW_BRANCH_PROB_FUNC(i, j, k, l, RHO_VAL, FIRST_GROUP, _
            SECOND_GROUP, POS_ADJ_RHO_ARR, NEG_ADJ_RHO_ARR, RHO_FACTOR)
            
          NEXT_RISK_FREE_ARR(FIRST_TREE_DESC, SECOND_TREE_DESC) = _
            NEXT_RISK_FREE_ARR(FIRST_TREE_DESC, SECOND_TREE_DESC) _
            + TIME_STEP_RISK_FREE * RISK_FREE_DISC _
            * HW_BRANCH_PROB_FUNC(i, j, k, l, RHO_VAL, FIRST_GROUP, _
            SECOND_GROUP, POS_ADJ_RHO_ARR, NEG_ADJ_RHO_ARR, RHO_FACTOR)
            
        Next l
      Next k
    Next j
    
    RISKY_PRICE_ARR(i + 1) = NEXT_RISKY_ARR
    RF_PRICE_ARR(i + 1) = NEXT_RISK_FREE_ARR
  Next i
  HW_LIMIT_INDEX_FUNC = NROWS

Exit Function
ERROR_LABEL:
    HW_LIMIT_INDEX_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : HW_DISC_FUNC
'DESCRIPTION   :
'LIBRARY       : FIXED_INCOME
'GROUP         : HW_2F
'ID            : 028
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 10-06-2008
'************************************************************************************
'************************************************************************************

Private Function HW_DISC_FUNC(ByVal i As Long, _
ByVal INDEX_VAL As Double, _
ByVal INIT_STEP_VAL As Double, _
ByRef GRAD_ARR As Variant, _
ByRef TIME_ARR As Variant, _
ByRef FIT_ARR As Variant, _
ByRef PROB_INDEX_ARR As Variant, _
Optional ByVal epsilon As Double = 2 ^ 52, _
Optional ByVal tolerance As Double = 0.01, _
Optional ByVal VERSION As Integer = 0)
  
Dim TEMP_RATE As Double
Dim TEMP_DELTA As Double

On Error GoTo ERROR_LABEL

TEMP_RATE = (HW_NODE_RATE_FUNC(i, INDEX_VAL, INIT_STEP_VAL, GRAD_ARR, _
            PROB_INDEX_ARR, epsilon)) + _
            (FIT_ARR(HW_FIND_INDEX_FUNC(TIME_ARR(i), _
            TIME_ARR, tolerance)))
            
If VERSION <> 0 Then: TEMP_RATE = Exp(TEMP_RATE)
            
TEMP_DELTA = (TIME_ARR(i + 1) - TIME_ARR(i))
  
HW_DISC_FUNC = Exp(-1 * TEMP_RATE * TEMP_DELTA)
                'applicable discount rate for
                'a diffusion process at a node

Exit Function
ERROR_LABEL:
    HW_DISC_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : HW_ZERO_RATE_FUNC
'DESCRIPTION   :
'LIBRARY       : FIXED_INCOME
'GROUP         : HW_2F
'ID            : 029
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 10-06-2008
'************************************************************************************
'************************************************************************************

Private Function HW_ZERO_RATE_FUNC(ByVal XTEMP_VAL As Double, _
ByRef TENOR_ARR As Variant, _
ByRef DISC_ARR As Variant, _
ByRef RATE_ARR As Variant)
  
  Dim i As Long
  Dim j As Long
  Dim YTEMP_VAL As Double
  
  On Error GoTo ERROR_LABEL
  
  If XTEMP_VAL = 0 Then
    HW_ZERO_RATE_FUNC = 0
    Exit Function
  End If

  j = 0
  For i = 0 To UBound(TENOR_ARR, 1)
    If XTEMP_VAL < TENOR_ARR(i) Then
      j = i - 1
      GoTo 1983
    End If
  Next i
1983:
  
  YTEMP_VAL = TENOR_ARR(j)
  If XTEMP_VAL = YTEMP_VAL Then
    HW_ZERO_RATE_FUNC = DISC_ARR(j)
  Else
    HW_ZERO_RATE_FUNC = DISC_ARR(j) * Exp(-RATE_ARR(j) * _
                (XTEMP_VAL - YTEMP_VAL))
  End If

Exit Function
ERROR_LABEL:
    HW_ZERO_RATE_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : HW_BK_RESET_VAR_FUNC
'DESCRIPTION   :
'LIBRARY       : FIXED_INCOME
'GROUP         : HW_2F
'ID            : 030
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 10-06-2008
'************************************************************************************
'************************************************************************************

Private Function HW_BK_RESET_VAR_FUNC()

On Error GoTo ERROR_LABEL

HW_BK_RESET_VAR_FUNC = False

   PUB_CAP_VAL = 0
   PUB_NOM_VAL = 0
   PUB_EXPIRAT_VAL = 0
   PUB_INIT_STEP_VAL = 0

   PUB_FIX_ARR = 0
   PUB_START_ARR = 0
   PUB_END_ARR = 0
   PUB_ACCR_ARR = 0
   PUB_FLOOR_ARR = 0
   PUB_FORW_ARR = 0

   PUB_TENOR_ARR = 0
   PUB_RATE_ARR = 0

   PUB_RF_TENOR_ARR = 0
   PUB_RF_RATE_ARR = 0
   PUB_RF_DISC_ARR = 0

   PUB_MEAN_VAL = 0
   PUB_SIGMA_VAL = 0
   PUB_TARGET_VAL = 0

   PUB_START_TENOR_VAL = 0
   PUB_FIXED_RATE_VAL = 0

   PUB_LOWER_BOUND_VAL = 0
   PUB_UPPER_BOUND_VAL = 0

   PUB_GUESS_VAL = 0
   PUB_nLOOPS_VAL = 0
   PUB_TOL_VAL = 0
   PUB_EPS_VAL = 0

   PUB_BK_MIN = 0
   PUB_BK_DELTA = 0
   PUB_BK_PRICE = 0
   PUB_BK_GRAD = 0
   PUB_BK_SIZE = 0
   PUB_BK_ARR = 0

   PUB_VERSION_VAL = 0


HW_BK_RESET_VAR_FUNC = True

Exit Function
ERROR_LABEL:
HW_BK_RESET_VAR_FUNC = False
End Function

'-----------------------------------------------------------------------------------
'The Vasicek and CIR models for pricing bond options do not automatically
'fit today’s term structure. The difference between an equilibrium model
'and a no-arbitrage model is as follows. In an equilibrium model, today’s
'term structure of interest rates is an output. In a no-arbitrage model,
'today’s term structure of interest rates is an input. The following model shows an
'alternative approach for overcoming this limitation. This involves building
'what is known as the Hull White Model.

'In a paper published in 1990, Hull and White explored extensions of the
'Vasicek model that provide an exact fit to the initial term structure .
'One version of the extended Vasicek model that they consider is:

'o   Many analytic results for bond prices and option prices
'o   Two volatility parameters, a and s
'o   Interest rates normally distributed
'o   Standard deviation of a forward rate is a declining function of its maturity


'-----------------------------------------------------------------------------
'By choosing the parameters judiciously, they can be made to provide an
'approximate fit to many of the term structures that are encountered in
'practice. But the fit is not usually an exact one and, in some cases,
'there are significant errors. Most traders find this unsatisfactory. Not
'unreasonably, they argue that they have very little confidence in the
'price of a bond option when the model does not price the underlying bond
'correctly. A 1% error in the price of the underlying bond may lead to a
'25% error in an option price.
 '------------------------------------------------------------------------------------------

