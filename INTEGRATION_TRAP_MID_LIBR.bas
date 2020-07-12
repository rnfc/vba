Attribute VB_Name = "INTEGRATION_TRAP_MID_LIBR"

'-------------------------------------------------------------------------------------
'-------------------------------------------------------------------------------------

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.
'-------------------------------------------------------------------------------------
'-------------------------------------------------------------------------------------

'************************************************************************************
'************************************************************************************
'FUNCTION      : MID_POINT_INTEGRATION_FUNC
'DESCRIPTION   : Numerical Integration (Mid Point Method)
'LIBRARY       : INTEGRATION
'GROUP         : TRAP_MID
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 21/01/2009
'************************************************************************************
'************************************************************************************

Function MID_POINT_INTEGRATION_FUNC(ByVal FUNC_NAME_STR As String, _
ByVal START_VALUE As Double, _
ByVal END_VALUE As Double, _
ByVal COUNTER As Long, _
Optional ByVal OUTPUT As Integer = 0)

Dim i As Long
Dim TEMP_DELTA As Double
Dim TEMP_MATRIX As Variant

On Error GoTo ERROR_LABEL:

TEMP_DELTA = (END_VALUE - START_VALUE) / COUNTER

ReDim TEMP_MATRIX(0 To COUNTER, 1 To 5)

For i = 1 To COUNTER
     TEMP_MATRIX(i, 1) = START_VALUE + (TEMP_DELTA * (i - 1)) 'FIRST_VALUE
     TEMP_MATRIX(i, 2) = START_VALUE + (TEMP_DELTA * i) 'SECOND_VALUE
     TEMP_MATRIX(i, 3) = 0.5 * (TEMP_MATRIX(i, 1) + TEMP_MATRIX(i, 2)) 'MID_POINT
     TEMP_MATRIX(i, 4) = TEMP_MATRIX(i - 1, 4) + _
        Excel.Application.Run(FUNC_NAME_STR, TEMP_MATRIX(i, 3)) 'INTEGRAL
     TEMP_MATRIX(i, 5) = TEMP_MATRIX(i, 4) * TEMP_DELTA
Next i

TEMP_MATRIX(0, 1) = "FIRST_VALUE"
TEMP_MATRIX(0, 2) = "SECOND_VALUE"
TEMP_MATRIX(0, 3) = "MID_POINT"
TEMP_MATRIX(0, 4) = "INTEGRAL"
TEMP_MATRIX(0, 5) = "MID INTEGRAL"

Select Case OUTPUT
    Case 0
        MID_POINT_INTEGRATION_FUNC = TEMP_DELTA * TEMP_MATRIX(COUNTER, 4)
    Case Else
        MID_POINT_INTEGRATION_FUNC = TEMP_MATRIX
End Select
    
Exit Function
ERROR_LABEL:
MID_POINT_INTEGRATION_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : TRAP_INTEGRATION_FUNC
'DESCRIPTION   : Numerical Integration (Trap Method)
'LIBRARY       : INTEGRATION
'GROUP         : TRAP_MID
'ID            : 002
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 21/01/2009
'************************************************************************************
'************************************************************************************

Function TRAP_INTEGRATION_FUNC(ByVal FUNC_NAME_STR As String, _
ByVal START_VALUE As Double, _
ByVal END_VALUE As Double, _
ByVal COUNTER As Long, _
Optional ByVal OUTPUT As Integer = 0)

Dim i As Long
Dim Integral As Double
Dim TEMP_DELTA As Double
Dim TEMP_MATRIX As Variant

On Error GoTo ERROR_LABEL
    
TEMP_DELTA = (END_VALUE - START_VALUE) / COUNTER
Integral = Excel.Application.Run(FUNC_NAME_STR, _
    START_VALUE) 'USE THE SAME FORMAT AS THE

ReDim TEMP_MATRIX(0 To COUNTER, 1 To 2)

TEMP_MATRIX(0, 1) = Integral
TEMP_MATRIX(0, 2) = Integral * (TEMP_DELTA / 2)
    
For i = 1 To (COUNTER - 1)
    Integral = Integral + 2 * _
        Excel.Application.Run(FUNC_NAME_STR, START_VALUE + (TEMP_DELTA * i))
    TEMP_MATRIX(i, 1) = Integral
    TEMP_MATRIX(i, 2) = Integral * (TEMP_DELTA / 2)
Next i

TEMP_MATRIX(COUNTER, 1) = Integral + Excel.Application.Run(FUNC_NAME_STR, END_VALUE)
TEMP_MATRIX(COUNTER, 2) = TEMP_MATRIX(COUNTER, 1) * (TEMP_DELTA / 2)

Select Case OUTPUT
    Case 0
        TRAP_INTEGRATION_FUNC = (TEMP_DELTA / 2) * TEMP_MATRIX(COUNTER, 1)
    Case Else
        TRAP_INTEGRATION_FUNC = TEMP_MATRIX
End Select
    
Exit Function
ERROR_LABEL:
TRAP_INTEGRATION_FUNC = Err.number
End Function
