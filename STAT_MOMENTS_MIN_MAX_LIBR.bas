Attribute VB_Name = "STAT_MOMENTS_MIN_MAX_LIBR"


Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.


'************************************************************************************
'************************************************************************************
'FUNCTION      : MAXIMUM_FUNC
'DESCRIPTION   : COMPARE TWO VALUES, AND RETURNS THE MAX VALUE
'LIBRARY       : STATISTICS
'GROUP         : MAX-MIN
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 14/02/2008
'************************************************************************************
'************************************************************************************


Function MAXIMUM_FUNC(ByVal FIRST_VAL As Variant, _
ByVal SECOND_VAL As Variant)
  
On Error GoTo ERROR_LABEL

MAXIMUM_FUNC = FIRST_VAL
If SECOND_VAL > FIRST_VAL Then MAXIMUM_FUNC = SECOND_VAL
      
Exit Function
ERROR_LABEL:
MAXIMUM_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : MINIMUM_FUNC
'DESCRIPTION   : COMPARE TWO VALUES, AND RETURNS THE MIN VALUE
'LIBRARY       : STATISTICS
'GROUP         : MAX-MIN
'ID            : 002
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 14/02/2008
'************************************************************************************
'************************************************************************************
            
Function MINIMUM_FUNC(ByVal FIRST_VAL As Variant, _
ByVal SECOND_VAL As Variant)
  
On Error GoTo ERROR_LABEL
  
MINIMUM_FUNC = FIRST_VAL
If SECOND_VAL < FIRST_VAL Then MINIMUM_FUNC = SECOND_VAL
  
Exit Function
ERROR_LABEL:
MINIMUM_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : COLLAR_FUNC
'DESCRIPTION   : Collar(a; b; c) = max(a; min(b; c)).
'LIBRARY       : STATISTICS
'GROUP         : MAX-MIN
'ID            : 003
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 14/02/2008
'************************************************************************************
'************************************************************************************

Function COLLAR_FUNC(ByVal MIN_VAL As Variant, _
ByVal VAR_VAL As Variant, _
ByVal MAX_VAL As Variant) 'As Double
  
On Error GoTo ERROR_LABEL

If VAR_VAL < MIN_VAL Then
    COLLAR_FUNC = MIN_VAL
ElseIf VAR_VAL > MAX_VAL Then
    COLLAR_FUNC = MAX_VAL
Else
    COLLAR_FUNC = VAR_VAL
End If

Exit Function
ERROR_LABEL:
COLLAR_FUNC = Err.number
End Function
