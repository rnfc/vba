Attribute VB_Name = "WEB_NUMBER_TOLERANCE_LIBR"


Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.

'************************************************************************************
'************************************************************************************
'FUNCTION      : MACHINE_TOLERANCE_FUNC
'DESCRIPTION   : Machine Precision (tolerance;epsilon)
'LIBRARY       : WEB_NUMBER
'GROUP         : TOLERANCE
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA

'************************************************************************************
'************************************************************************************

Function MACHINE_TOLERANCE_FUNC()
Dim DATA_VAL As Double
On Error GoTo ERROR_LABEL
DATA_VAL = 1#
Do
    DATA_VAL = 0.5 * DATA_VAL
    If 1# + DATA_VAL = 1# Then Exit Do
Loop
MACHINE_TOLERANCE_FUNC = 2 * DATA_VAL
Exit Function
ERROR_LABEL:
MACHINE_TOLERANCE_FUNC = Err.number
End Function
