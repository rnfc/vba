Attribute VB_Name = "FINAN_DERIV_VARIATIONAL_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.


'************************************************************************************
'************************************************************************************
'FUNCTION      : VARIATIONAL_EUROPEAN_CALL_OPTION_FUNC
'DESCRIPTION   : Calculates Option Price of European call option using
'Variational formulation

'This function performs Finite element method to calculate a call price
'for European option.

'-----------------------------------------------------------------------------
'References:
'-----------------------------------------------------------------------------

'1. Lecture Notes for Finite Element Method
'   http://www.cs.utah.edu/classes/cs6220/cs6220_chapter4_1.pdf

'2.Notes on equations for FEA matrices
'   http://www.math.chalmers.se/cm/education/courses/0506/ppde/lectures/
'    lecture-2/ppde0506-lecture2.pdf
    
'3.American option pricing using Adaptive method - A simplification
'   of BS equation to european option is used for this application
'   http://www.math.umd.edu/research/spotlight/options.pdf

'LIBRARY       : DERIVATIVES
'GROUP         : Variational formulation
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/29/2009
'************************************************************************************
'************************************************************************************

Function VARIATIONAL_EUROPEAN_CALL_OPTION_FUNC(ByVal SPOT As Double, _
ByVal STRIKE As Double, _
ByVal RATE As Double, _
ByVal SIGMA As Double, _
ByVal TENOR As Double, _
Optional ByVal TIME_STEPS As Long = 50, _
Optional ByVal ASSET_STEPS As Long = 20, _
Optional ByVal CND_TYPE As Integer = 0)

Dim i As Long
Dim j As Long
Dim k As Long

Dim nSTEPS As Long

Dim D1_VAL As Double
Dim D2_VAL As Double

Dim TEMP_VAL As Double
Dim TEMP_TARGET As Double

Dim TEMP_LOW As Double
Dim TEMP_HIGH As Double

Dim TEMP_A_STEP As Double 'Asset Step Size
Dim TEMP_T_STEP As Double 'Time Step Size

Dim TEMP_THETA As Double
Dim TEMP_STRIKE As Double 'strike price

Dim ATEMP_ARR As Variant
Dim BTEMP_ARR As Variant
Dim CTEMP_ARR As Variant
Dim DTEMP_ARR As Variant
Dim ETEMP_ARR As Variant
Dim FTEMP_ARR As Variant
  
Dim MASS_MATRIX As Variant
Dim STIFF_MATRIX As Variant
Dim ATEMP_MATRIX As Variant
Dim BTEMP_MATRIX As Variant

Dim STEPS_VECTOR As Variant
Dim RESULT_VECTOR As Variant
Dim TEMP_VECTOR As Variant 'AssestStepsVec

On Error GoTo ERROR_LABEL

TEMP_STRIKE = Log(STRIKE) / Log(Exp(1))
ASSET_STEPS = Int(ASSET_STEPS / 2) * 2

TEMP_TARGET = Log(Exp(RATE * TENOR) * SPOT) / Log(Exp(1))
TEMP_LOW = Log(2) / Log(Exp(1))
TEMP_A_STEP = 2 * (TEMP_LOW + TEMP_TARGET) / ASSET_STEPS
TEMP_HIGH = TEMP_TARGET + 0.5 * ASSET_STEPS * TEMP_A_STEP

TEMP_A_STEP = (TEMP_HIGH + TEMP_LOW) / ASSET_STEPS
TEMP_T_STEP = TENOR / TIME_STEPS
nSTEPS = TIME_STEPS

ReDim STEPS_VECTOR(1 To (ASSET_STEPS - 1), 1 To 1) 'Create Inremental Matrix

TEMP_VAL = (-TEMP_LOW + TEMP_A_STEP)
For i = 1 To (ASSET_STEPS - 1)
  STEPS_VECTOR(i, 1) = TEMP_VAL
  TEMP_VAL = TEMP_VAL + TEMP_A_STEP
Next i

TEMP_THETA = 1 'Rolling Back
ReDim TEMP_VECTOR(1 To (ASSET_STEPS - 1), 1 To (ASSET_STEPS - 1))
For i = 1 To UBound(TEMP_VECTOR, 1) - 1
  TEMP_VECTOR(i, i + 1) = (1 / (6 * ((ASSET_STEPS - 1) + 1)))
  TEMP_VECTOR(i, i) = (2 / (3 * ((ASSET_STEPS - 1) + 1)))
  TEMP_VECTOR(i + 1, i) = (1 / (6 * ((ASSET_STEPS - 1) + 1)))
Next i
TEMP_VECTOR(UBound(TEMP_VECTOR, 1), _
            UBound(TEMP_VECTOR, 1)) = (2 / (3 * ((ASSET_STEPS - 1) + 1)))

MASS_MATRIX = MATRIX_ELEMENTS_MULT_SCALAR_FUNC(TEMP_VECTOR, TEMP_HIGH + TEMP_LOW)
  
ReDim TEMP_VECTOR(1 To (ASSET_STEPS - 1), 1 To (ASSET_STEPS - 1))

For i = 1 To UBound(TEMP_VECTOR, 1) - 1
  TEMP_VECTOR(i, i + 1) = -((ASSET_STEPS - 1) + 1)
  TEMP_VECTOR(i, i) = 2 * ((ASSET_STEPS - 1) + 1)
  TEMP_VECTOR(i + 1, i) = -((ASSET_STEPS - 1) + 1)
Next i
TEMP_VECTOR(UBound(TEMP_VECTOR, 1), _
            UBound(TEMP_VECTOR, 1)) = 2 * ((ASSET_STEPS - 1) + 1)

STIFF_MATRIX = MATRIX_ELEMENTS_MULT_SCALAR_FUNC(TEMP_VECTOR, _
                  0.5 * SIGMA * SIGMA * (1 / (TEMP_HIGH + TEMP_LOW)))

ATEMP_MATRIX = MATRIX_ELEMENTS_ADD_SCALAR_FUNC(STEPS_VECTOR, -TEMP_STRIKE)
BTEMP_MATRIX = ATEMP_MATRIX
ReDim ATEMP_MATRIX(1 To UBound(BTEMP_MATRIX, 1), 1 To UBound(BTEMP_MATRIX, 2))
For i = 1 To UBound(BTEMP_MATRIX, 1)
  For j = 1 To UBound(BTEMP_MATRIX, 2)
    ATEMP_MATRIX(i, j) = Abs(BTEMP_MATRIX(i, j))
  Next j
Next i

ATEMP_MATRIX = MATRIX_ELEMENTS_MULT_SCALAR_FUNC(ATEMP_MATRIX, (-1 / TEMP_A_STEP))
ATEMP_MATRIX = MATRIX_ELEMENTS_ADD_SCALAR_FUNC(ATEMP_MATRIX, 1)
BTEMP_MATRIX = ATEMP_MATRIX
ReDim ATEMP_MATRIX(1 To UBound(BTEMP_MATRIX, 1), 1 To UBound(BTEMP_MATRIX, 2))
For i = 1 To UBound(BTEMP_MATRIX, 1)
  For j = 1 To UBound(BTEMP_MATRIX, 2)
    ATEMP_MATRIX(i, j) = MAXIMUM_FUNC(BTEMP_MATRIX(i, j), 0)
  Next j
Next i

ATEMP_MATRIX = MATRIX_ELEMENTS_MULT_SCALAR_FUNC(ATEMP_MATRIX, 0.5 * SIGMA * _
              SIGMA * Exp(TEMP_STRIKE))

ATEMP_ARR = MATRIX_ELEMENTS_MULT_SCALAR_FUNC(STIFF_MATRIX, TEMP_T_STEP * TEMP_THETA)
ATEMP_ARR = MATRIX_ELEMENTS_ADD_FUNC(MASS_MATRIX, ATEMP_ARR, 1, 1)

BTEMP_ARR = MATRIX_ELEMENTS_MULT_SCALAR_FUNC(STIFF_MATRIX, TEMP_T_STEP * (TEMP_THETA - 1))
BTEMP_ARR = MATRIX_ELEMENTS_ADD_FUNC(MASS_MATRIX, BTEMP_ARR, 1, 1)

ReDim TEMP_VECTOR(1 To (ASSET_STEPS - 1), 1 To 1)

For i = 1 To nSTEPS
  CTEMP_ARR = MMULT_FUNC(BTEMP_ARR, TEMP_VECTOR, 70)
  DTEMP_ARR = MATRIX_ELEMENTS_MULT_SCALAR_FUNC(ATEMP_MATRIX, TEMP_T_STEP)
  ETEMP_ARR = MATRIX_ELEMENTS_ADD_FUNC(CTEMP_ARR, DTEMP_ARR, 1, 1)
  FTEMP_ARR = MATRIX_SVD_INVERSE_FUNC(ATEMP_ARR)
  TEMP_VECTOR = MMULT_FUNC(FTEMP_ARR, ETEMP_ARR, 70)
Next i

k = (UBound(TEMP_VECTOR, 1) + 1) / 2

ReDim RESULT_VECTOR(1 To 2, 1 To 2)
RESULT_VECTOR(1, 1) = "FEA Call price"
RESULT_VECTOR(1, 2) = (TEMP_VECTOR(k, 1) + MAXIMUM_FUNC(Exp(STEPS_VECTOR(k, 1)) - Exp(TEMP_STRIKE), 0)) * Exp(-RATE * TENOR)

D1_VAL = (Log(SPOT / STRIKE) + (RATE + SIGMA ^ 2 / 2) * TENOR) / (SIGMA * Sqr(TENOR))
D2_VAL = D1_VAL - SIGMA * Sqr(TENOR)

RESULT_VECTOR(2, 1) = "Analy Call Price"
RESULT_VECTOR(2, 2) = SPOT * CND_FUNC(D1_VAL, CND_TYPE) - STRIKE * Exp(-RATE * TENOR) * CND_FUNC(D2_VAL, CND_TYPE) 'Analy Call Price

VARIATIONAL_EUROPEAN_CALL_OPTION_FUNC = RESULT_VECTOR

Exit Function
ERROR_LABEL:
VARIATIONAL_EUROPEAN_CALL_OPTION_FUNC = Err.number
End Function
