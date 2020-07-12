Attribute VB_Name = "FINAN_PORT_WEIGHTS_MARKOW_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.
                    
Private PUB_TARGET_RETURN As Double
Private PUB_EXPECTED_VECTOR As Variant
Private PUB_COVARIANCE_MATRIX As Variant
Private Const PUB_EPSILON As Double = 2 ^ 52

'************************************************************************************
'************************************************************************************
'FUNCTION      : PORT_MARKOWITZ_WEIGHTS_OPTIMIZER_FUNC (Short sale allowed!!!)

'DESCRIPTION   : This routine calculates optimum weights for components of a stock portfolio.
'The weights are calculated accoring to Markowitz theory for minimizing risk
'for a given target portfolio return

'Markowitz model provides a way to define a portfolio which is efficient - one
'which has highest reward for a given risk, or lowest risk for a given reward

'Create an objective function which takes guess values of portfolio weights and returns
'portfolio variance. Input weights vector has length NSIZE-2. This is becuase the
'other 2 weights are solved by using theses constraints:

'   Weight of all members in portfolio should be 1
'   Expected return of portfolio should match target portfolio return

'Use simplex method to find set of weights which minimize portfolio variance.

'LIBRARY       : PORTFOLIO
'GROUP         : WEIGHTS_MARKOWITZ
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 21/01/2009
'************************************************************************************
'************************************************************************************

Function PORT_MARKOWITZ_WEIGHTS_OPTIMIZER_FUNC(ByRef EXPECTED_RNG As Variant, _
ByRef COVARIANCE_RNG As Variant, _
ByVal TARGET_RETURN As Double, _
Optional ByVal nLOOPS As Long = 1000, _
Optional ByVal tolerance As Double = 0.001, _
Optional ByVal OUTPUT As Integer = 0)

'COVARIANCE_RNG: covariance matrix of stock returns

'EXPECTED_RNG: dont assumed that expected returns for each stock is equal
'to historical returns. This might not be necessarily true!

Dim i As Long
Dim h As Long

Dim NSIZE As Long

Dim Y_VAL As Double
Dim PARAM_VECTOR As Variant
Dim WEIGHTS_VECTOR As Variant

Dim GUESS_VECTOR As Variant 'GuessWeightsVec

On Error GoTo ERROR_LABEL

PUB_TARGET_RETURN = TARGET_RETURN
PUB_COVARIANCE_MATRIX = COVARIANCE_RNG
If UBound(PUB_COVARIANCE_MATRIX, 1) <> UBound(PUB_COVARIANCE_MATRIX, 2) Then: GoTo ERROR_LABEL
NSIZE = UBound(PUB_COVARIANCE_MATRIX, 1)

PUB_EXPECTED_VECTOR = EXPECTED_RNG
If UBound(PUB_EXPECTED_VECTOR, 1) = 1 Then
    PUB_EXPECTED_VECTOR = MATRIX_TRANSPOSE_FUNC(PUB_EXPECTED_VECTOR)
End If
If UBound(PUB_EXPECTED_VECTOR, 1) <> NSIZE Then: GoTo ERROR_LABEL

h = 2
ReDim GUESS_VECTOR(1 To NSIZE - h, 1 To 1)
For i = 1 To NSIZE - h: GUESS_VECTOR(i, 1) = 1 / NSIZE: Next i
PARAM_VECTOR = NELDER_MEAD_OPTIMIZATION3_FUNC("MARKOWITZ_VARIANCE_OBJ_FUNC", GUESS_VECTOR, nLOOPS, tolerance)

'------------------------------------------------------------------------------------
Select Case OUTPUT
'------------------------------------------------------------------------------------
Case Is <= 1
'------------------------------------------------------------------------------------
    ReDim WEIGHTS_VECTOR(1 To NSIZE - h, 1 To 1)
    For i = 1 To NSIZE - h: WEIGHTS_VECTOR(i, 1) = PARAM_VECTOR(i, 1): Next i
    WEIGHTS_VECTOR = MARKOWITZ_ALL_WEIGHTS_FUNC(WEIGHTS_VECTOR, PUB_EXPECTED_VECTOR, PUB_TARGET_RETURN)
    If OUTPUT = 0 Then 'Optimal Weights
        PORT_MARKOWITZ_WEIGHTS_OPTIMIZER_FUNC = WEIGHTS_VECTOR
    Else
        Y_VAL = 0
        For i = 1 To NSIZE: Y_VAL = Y_VAL + WEIGHTS_VECTOR(i, 1) * PUB_EXPECTED_VECTOR(i, 1): Next i
        PORT_MARKOWITZ_WEIGHTS_OPTIMIZER_FUNC = Y_VAL
    End If
'------------------------------------------------------------------------------------
Case Else 'Portfolio Variance
'------------------------------------------------------------------------------------
    PORT_MARKOWITZ_WEIGHTS_OPTIMIZER_FUNC = MARKOWITZ_VARIANCE_OBJ_FUNC(PARAM_VECTOR) ^ 0.5
'------------------------------------------------------------------------------------
End Select
'------------------------------------------------------------------------------------

Exit Function
ERROR_LABEL:
PORT_MARKOWITZ_WEIGHTS_OPTIMIZER_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : MARKOWITZ_VARIANCE_OBJ_FUNC
'DESCRIPTION   :
'LIBRARY       : PORTFOLIO
'GROUP         : WEIGHTS_MARKOWITZ
'ID            : 002
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 21/01/2009
'************************************************************************************
'************************************************************************************

Private Function MARKOWITZ_VARIANCE_OBJ_FUNC(ByRef WEIGHTS_VECTOR As Variant)

Dim i As Long
Dim k As Long
Dim h As Long
Dim NSIZE As Long
Dim X_VAL As Double
Dim Y_VAL As Double

Dim DATA_VECTOR As Variant
'Dim WEIGHTS_VECTOR As Variant

On Error GoTo ERROR_LABEL

'WEIGHTS_VECTOR = WEIGHTS_RNG
'If UBound(WEIGHTS_VECTOR, 1) = 1 Then
'    WEIGHTS_VECTOR = MATRIX_TRANSPOSE_FUNC(WEIGHTS_VECTOR)
'End If
h = 2
NSIZE = UBound(WEIGHTS_VECTOR, 1) + h
ReDim DATA_VECTOR(1 To NSIZE - h, 1 To 1)
X_VAL = 1
For i = 1 To NSIZE - h
    DATA_VECTOR(i, 1) = WEIGHTS_VECTOR(i, 1)
    'If WEIGHTS_VECTOR(i, 1) > 1 Or WEIGHTS_VECTOR(i, 1) < 0 Then: X_VAL = PUB_EPSILON
Next i
DATA_VECTOR = MARKOWITZ_ALL_WEIGHTS_FUNC(DATA_VECTOR, PUB_EXPECTED_VECTOR, PUB_TARGET_RETURN)
Y_VAL = 0
For i = 1 To NSIZE
    For k = 1 To NSIZE
      Y_VAL = Y_VAL + DATA_VECTOR(i, 1) * DATA_VECTOR(k, 1) * PUB_COVARIANCE_MATRIX(i, k)
    Next k
Next i
MARKOWITZ_VARIANCE_OBJ_FUNC = Abs(Y_VAL * X_VAL) ^ 2
   
Exit Function
ERROR_LABEL:
MARKOWITZ_VARIANCE_OBJ_FUNC = PUB_EPSILON
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : MARKOWITZ_ALL_WEIGHTS_FUNC
'DESCRIPTION   :
'LIBRARY       : PORTFOLIO
'GROUP         : WEIGHTS_MARKOWITZ
'ID            : 003
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 21/01/2009
'************************************************************************************
'************************************************************************************

Private Function MARKOWITZ_ALL_WEIGHTS_FUNC(ByRef WEIGHTS_RNG As Variant, _
ByRef EXPECTED_RNG As Variant, _
ByVal TARGET_RETURN As Double)

Dim i As Long
Dim h As Long
Dim NSIZE As Long

Dim TEMP1_SUM As Double
Dim TEMP2_SUM As Double
Dim TEMP3_SUM As Double

Dim TEMP1_VAL As Double
Dim TEMP2_VAL As Double
Dim TEMP3_VAL As Double

Dim DATA_VECTOR As Variant
Dim WEIGHTS_VECTOR As Variant
Dim EXPECTED_VECTOR As Variant
  
On Error GoTo ERROR_LABEL

h = 2
WEIGHTS_VECTOR = WEIGHTS_RNG
If UBound(WEIGHTS_VECTOR, 1) = 1 Then
    WEIGHTS_VECTOR = MATRIX_TRANSPOSE_FUNC(WEIGHTS_VECTOR)
End If

EXPECTED_VECTOR = EXPECTED_RNG
If UBound(EXPECTED_VECTOR, 1) = 1 Then
    EXPECTED_VECTOR = MATRIX_TRANSPOSE_FUNC(EXPECTED_VECTOR)
End If

NSIZE = UBound(EXPECTED_VECTOR, 1)
ReDim DATA_VECTOR(1 To NSIZE, 1 To 1)
TEMP1_VAL = EXPECTED_VECTOR(NSIZE, 1)
TEMP2_VAL = EXPECTED_VECTOR(NSIZE - 1, 1)

TEMP3_VAL = 1
TEMP1_SUM = TARGET_RETURN
  
For i = 1 To NSIZE - h
    DATA_VECTOR(i, 1) = WEIGHTS_VECTOR(i, 1)
    TEMP1_SUM = TEMP1_SUM - EXPECTED_VECTOR(i, 1) * WEIGHTS_VECTOR(i, 1)
Next i
TEMP2_SUM = TEMP3_VAL
For i = 1 To NSIZE - h
    TEMP2_SUM = TEMP2_SUM - WEIGHTS_VECTOR(i, 1)
Next i
DATA_VECTOR(NSIZE - 1, 1) = (TEMP1_SUM - TEMP1_VAL * TEMP2_SUM) / (TEMP2_VAL - TEMP1_VAL)
  
TEMP3_SUM = TEMP3_VAL
For i = 1 To NSIZE - 1: TEMP3_SUM = TEMP3_SUM - DATA_VECTOR(i, 1): Next i
DATA_VECTOR(NSIZE, 1) = TEMP3_SUM
MARKOWITZ_ALL_WEIGHTS_FUNC = DATA_VECTOR

Exit Function
ERROR_LABEL:
MARKOWITZ_ALL_WEIGHTS_FUNC = Err.number
End Function
