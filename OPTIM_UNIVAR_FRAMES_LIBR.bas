Attribute VB_Name = "OPTIM_UNIVAR_FRAMES_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.


'************************************************************************************
'************************************************************************************
'FUNCTION      : CALL_TEST_UNIVAR_FRAME_FUNC
'DESCRIPTION   : Generic framework for testing Univariate Optimization functions
'LIBRARY       : OPTIMIZATION
'GROUP         : UNIVAR_FRAMES
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/08/2008
'************************************************************************************
'************************************************************************************

Function CALL_TEST_UNIVAR_FRAME_FUNC(ByVal LOWER_VAL As Double, _
ByVal UPPER_VAL As Double, _
ByVal FUNC_NAME_STR As String, _
Optional ByVal MIN_FLAG As Boolean = True, _
Optional ByVal nLOOPS As Long = 500, _
Optional ByVal epsilon As Double = 10 ^ -15)

Dim COUNTER As Long
Dim XTEMP_VAL As Double
Dim CONST_BOX As Variant
Dim TEMP_MATRIX As Variant

On Error GoTo ERROR_LABEL

ReDim CONST_BOX(1 To 2, 1 To 1)
CONST_BOX(1, 1) = LOWER_VAL
CONST_BOX(2, 1) = UPPER_VAL

ReDim TEMP_MATRIX(0 To 4, 1 To 5)

TEMP_MATRIX(0, 1) = "ALGORITHM"
TEMP_MATRIX(0, 2) = "X_VAL"
TEMP_MATRIX(0, 3) = "Y_VAL"
TEMP_MATRIX(0, 4) = "GRADIENT FD APPROX"
TEMP_MATRIX(0, 5) = "COUNTER"

'-------------------------------------------------------------------------------

COUNTER = 0
TEMP_MATRIX(1, 1) = "Divide-Conquer 1D"
XTEMP_VAL = UNIVAR_MIN_DIVIDE_CONQUER_FUNC(FUNC_NAME_STR, CONST_BOX, MIN_FLAG, COUNTER, _
            nLOOPS, epsilon)

TEMP_MATRIX(1, 2) = XTEMP_VAL
TEMP_MATRIX(1, 3) = Excel.Application.Run(FUNC_NAME_STR, TEMP_MATRIX(1, 2))
TEMP_MATRIX(1, 4) = UNIVAR_FD_GRAD_APPROX_FUNC(FUNC_NAME_STR, _
                        TEMP_MATRIX(1, 2))
TEMP_MATRIX(1, 5) = COUNTER


'-------------------------------------------------------------------------------

'-------------------------------------------------------------------------------

COUNTER = 0
TEMP_MATRIX(2, 1) = "Parabolic Method"
XTEMP_VAL = UNIVAR_MIN_PARABOLIC_FUNC(FUNC_NAME_STR, CONST_BOX, MIN_FLAG, COUNTER, _
            nLOOPS, epsilon)

TEMP_MATRIX(2, 2) = XTEMP_VAL
TEMP_MATRIX(2, 3) = Excel.Application.Run(FUNC_NAME_STR, TEMP_MATRIX(2, 2))
TEMP_MATRIX(2, 4) = UNIVAR_FD_GRAD_APPROX_FUNC(FUNC_NAME_STR, _
                        TEMP_MATRIX(2, 2))
TEMP_MATRIX(2, 5) = COUNTER


'-------------------------------------------------------------------------------
'-------------------------------------------------------------------------------

COUNTER = 0
TEMP_MATRIX(3, 1) = "Gold Method"
XTEMP_VAL = UNIVAR_MIN_GOLD_FUNC(FUNC_NAME_STR, CONST_BOX, MIN_FLAG, COUNTER, _
            nLOOPS, epsilon)

TEMP_MATRIX(3, 2) = XTEMP_VAL
TEMP_MATRIX(3, 3) = Excel.Application.Run(FUNC_NAME_STR, TEMP_MATRIX(3, 2))
TEMP_MATRIX(3, 4) = UNIVAR_FD_GRAD_APPROX_FUNC(FUNC_NAME_STR, _
                        TEMP_MATRIX(3, 2))
TEMP_MATRIX(3, 5) = COUNTER

'-------------------------------------------------------------------------------
'-------------------------------------------------------------------------------

COUNTER = 0
TEMP_MATRIX(4, 1) = "Brent Method"
XTEMP_VAL = UNIVAR_MIN_BRENT_FUNC(FUNC_NAME_STR, CONST_BOX, MIN_FLAG, COUNTER, _
            nLOOPS, epsilon)

TEMP_MATRIX(4, 2) = XTEMP_VAL
TEMP_MATRIX(4, 3) = Excel.Application.Run(FUNC_NAME_STR, TEMP_MATRIX(4, 2))
TEMP_MATRIX(4, 4) = UNIVAR_FD_GRAD_APPROX_FUNC(FUNC_NAME_STR, _
                        TEMP_MATRIX(4, 2))
TEMP_MATRIX(4, 5) = COUNTER

'-------------------------------------------------------------------------------

CALL_TEST_UNIVAR_FRAME_FUNC = TEMP_MATRIX

Exit Function
ERROR_LABEL:
CALL_TEST_UNIVAR_FRAME_FUNC = Err.number
End Function

