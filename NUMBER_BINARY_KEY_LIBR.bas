Attribute VB_Name = "NUMBER_BINARY_KEY_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.

'************************************************************************************
'************************************************************************************
'FUNCTION      : BINARY_KEY_FUNC
'DESCRIPTION   : Generate a Binary Key
'LIBRARY       : NUMBER_BINARY
'GROUP         : KEY
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 14/02/2008
'************************************************************************************
'************************************************************************************

Function BINARY_KEY_FUNC(ByVal DATA_STR As String, _
Optional ByVal NSIZE As Long = 100, _
Optional ByVal MULTIPLIER As Double = 2, _
Optional ByVal LAMBDA As Double = 1983)

Dim i As Long
Dim TEMP_STR As String
Dim TEMP_FACTOR As Double

On Error GoTo ERROR_LABEL

Randomize
TEMP_STR = DATA_STR
TEMP_FACTOR = Rnd * LAMBDA + 1

For i = 1 To NSIZE
    TEMP_STR = TEMP_STR & Chr(MULTIPLIER * i Mod TEMP_FACTOR)
Next i

BINARY_KEY_FUNC = TEMP_STR

Exit Function
ERROR_LABEL:
BINARY_KEY_FUNC = Err.number
End Function

