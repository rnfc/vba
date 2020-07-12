Attribute VB_Name = "WEB_NUMBER_SIGN_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.

'************************************************************************************
'************************************************************************************
'FUNCTION      : LOOK_SIGN_FUNC
'DESCRIPTION   : Break the variable string into the name and its sign (if any)
'LIBRARY       : STRING
'GROUP         : SIGN
'ID            : 001
'UPDATE        : 02/20/2008
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA

'************************************************************************************
'************************************************************************************

Function LOOK_SIGN_FUNC(ByVal DATA_VAL As Double)

Dim TEMP_CHR As String

On Error GoTo ERROR_LABEL

TEMP_CHR = Left(DATA_VAL, 1)

If TEMP_CHR = "-" Or TEMP_CHR = "+" Then
    DATA_VAL = Right(DATA_VAL, Len(DATA_VAL) - 1)
Else
    If IS_NUMERIC_FUNC(DATA_VAL, DECIMAL_SEPARATOR_FUNC()) = True Then
        If DATA_VAL >= 0 Then
              TEMP_CHR = "+"
        Else
              TEMP_CHR = "-"
        End If
    Else
        GoTo ERROR_LABEL
    End If
End If

LOOK_SIGN_FUNC = TEMP_CHR
  
  Exit Function
ERROR_LABEL:
LOOK_SIGN_FUNC = Err.number
End Function
