Attribute VB_Name = "INTEGRATION_SIMPSON_LIBR"

'-------------------------------------------------------------------------------------
'-------------------------------------------------------------------------------------

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.
'-------------------------------------------------------------------------------------
'-------------------------------------------------------------------------------------

'************************************************************************************
'************************************************************************************
'FUNCTION      : SIMPSON_INTEGRATION_FUNC

'DESCRIPTION   : Integration using the Simpson method with accuracy estimation.
'The integral of the function over [LOWER_BOUND,UPPER_BOUND] is calculated with the
'accuracy of order Epsilon.

'LIBRARY       : INTEGRATION
'GROUP         : SIMPSON
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 21/01/2009
'************************************************************************************
'************************************************************************************

Function SIMPSON_INTEGRATION_FUNC(ByVal FUNC_NAME_STR As String, _
ByRef LOWER_BOUND As Double, _
ByRef UPPER_BOUND As Double, _
Optional ByRef epsilon As Double = 2 ^ -52)
    
Dim i As Long
Dim j As Long

Dim X_VAL As Double

Dim S1_VAL As Double
Dim S2_VAL As Double
Dim S3_VAL As Double
Dim S4_VAL As Double

Dim DELTA_VAL As Double
Dim RESULT_VAL As Double

On Error GoTo ERROR_LABEL

S3_VAL = 1#
DELTA_VAL = UPPER_BOUND - LOWER_BOUND
S1_VAL = Excel.Application.Run(FUNC_NAME_STR, LOWER_BOUND) + _
         Excel.Application.Run(FUNC_NAME_STR, UPPER_BOUND)

i = 0
Do
    S4_VAL = S3_VAL
    DELTA_VAL = DELTA_VAL / 2#
    S2_VAL = 0#
    X_VAL = LOWER_BOUND + DELTA_VAL
    j = 0
    Do
        S2_VAL = S2_VAL + 2# * Excel.Application.Run(FUNC_NAME_STR, X_VAL)
        X_VAL = X_VAL + 2# * DELTA_VAL
    j = j + 1
    Loop Until Not X_VAL < UPPER_BOUND
    S1_VAL = S1_VAL + S2_VAL
    S3_VAL = (S1_VAL + S2_VAL) * DELTA_VAL / 3#
    X_VAL = Abs(S4_VAL - S3_VAL) / 15#
i = i + 1
Loop Until Not X_VAL > epsilon
RESULT_VAL = S3_VAL

SIMPSON_INTEGRATION_FUNC = RESULT_VAL
    
Exit Function
ERROR_LABEL:
SIMPSON_INTEGRATION_FUNC = Err.number
End Function
