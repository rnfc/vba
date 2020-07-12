Attribute VB_Name = "NUMBER_REAL_FLOOR_CEILING_LIBR"


Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.

'************************************************************************************
'************************************************************************************
'FUNCTION      : CEILING_FUNC
'DESCRIPTION   : Returns number rounded up, away from zero, to the nearest
'multiple of significance. For example, if you want to avoid using pennies
'in your prices and your product is priced at $4.42, use the
'formula =CEILING_FUNC(4.42,0.05) to round prices up to the nearest
'nickel.
'LIBRARY       : NUMBER_REAL
'GROUP         : FLOOR_CEILING
'ID            : 001
'LAST UPDATE   : 12/02/2008
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA

'************************************************************************************
'************************************************************************************

Function CEILING_FUNC(ByVal DATA_VAL As Double, _
ByVal FACTOR_VAL As Double)

Dim TEMP_VAL As Double

On Error GoTo ERROR_LABEL
  
  FACTOR_VAL = Abs(FACTOR_VAL)
  TEMP_VAL = Int(DATA_VAL / FACTOR_VAL) * FACTOR_VAL
  
  If TEMP_VAL = 0 Then
    CEILING_FUNC = FACTOR_VAL
    Exit Function
  End If
  
  If TEMP_VAL = DATA_VAL Then
    CEILING_FUNC = DATA_VAL
  Else
    CEILING_FUNC = TEMP_VAL + FACTOR_VAL * Sgn(TEMP_VAL)
  End If

Exit Function
ERROR_LABEL:
CEILING_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : FLOOR_FUNC
'DESCRIPTION   : Rounds number down, toward zero, to the nearest
'multiple of significance.
'LIBRARY       : NUMBER_REAL
'GROUP         : FLOOR_CEILING
'ID            : 002
'LAST UPDATE   : 12/02/2008
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'************************************************************************************
'************************************************************************************

Function FLOOR_FUNC(ByVal DATA_VAL As Double, _
ByVal FACTOR_VAL As Double)

On Error GoTo ERROR_LABEL

  If DATA_VAL < 0 And FACTOR_VAL >= 0 Then
    FLOOR_FUNC = 0
    Exit Function
  End If
  
  If DATA_VAL > 0 Then: FACTOR_VAL = Abs(FACTOR_VAL)
  'FACTOR_VAL = Abs(FACTOR_VAL)
  FLOOR_FUNC = Int(DATA_VAL / FACTOR_VAL) * FACTOR_VAL

Exit Function
ERROR_LABEL:
FLOOR_FUNC = 0
End Function
