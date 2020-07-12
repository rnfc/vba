Attribute VB_Name = "WEB_NUMBER_HEX_LIBR"


Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.

'************************************************************************************
'************************************************************************************
'FUNCTION      : CONVERT_NUMBER_HEX_FUNC
'DESCRIPTION   : Returns a string representing the hexadecimal value of a number.
'LIBRARY       : NUMBER
'GROUP         : HEX
'ID            : 001
'UPDATE        : 02/20/2008
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'************************************************************************************
'************************************************************************************

Function CONVERT_NUMBER_HEX_FUNC(ByVal DATA_VAL As Double, _
Optional ByVal VERSION As Integer = 0)

Dim i As Long
Dim TEMP_STR As String

On Error GoTo ERROR_LABEL
  
If VERSION <> 0 Then
    CONVERT_NUMBER_HEX_FUNC = Hex(DATA_VAL)
    Exit Function
End If
  
  TEMP_STR = Hex(Int(DATA_VAL)) & "."
  DATA_VAL = DATA_VAL - Int(DATA_VAL)
  For i = 1 To 16
    DATA_VAL = DATA_VAL * 16
    TEMP_STR = TEMP_STR & Hex(DATA_VAL)
    DATA_VAL = DATA_VAL - Int(DATA_VAL)
  Next i
  CONVERT_NUMBER_HEX_FUNC = TEMP_STR
Exit Function
ERROR_LABEL:
CONVERT_NUMBER_HEX_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : CONVERT_HEX_NUMBER_FUNC
'DESCRIPTION   : Returns a number representing the hexadecimal value of a string.
'LIBRARY       : NUMBER
'GROUP         : HEX
'ID            : 002
'UPDATE        : 02/20/2008
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'************************************************************************************
'************************************************************************************

Function CONVERT_HEX_NUMBER_FUNC(ByVal DATA_STR As String)
  
  Dim TEMP_VAL As Double
    
  On Error GoTo ERROR_LABEL
  
  Const CORREC_FACTOR = 4294967296#
  
  TEMP_VAL = CDec("&H" & DATA_STR)
  If TEMP_VAL > 0 Then
     CONVERT_HEX_NUMBER_FUNC = TEMP_VAL
  Else
     CONVERT_HEX_NUMBER_FUNC = CDec(CORREC_FACTOR + TEMP_VAL)
  End If
  
Exit Function
ERROR_LABEL:
CONVERT_HEX_NUMBER_FUNC = Err.number
End Function
