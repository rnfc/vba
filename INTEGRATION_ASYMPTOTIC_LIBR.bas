Attribute VB_Name = "INTEGRATION_ASYMPTOTIC_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.

'************************************************************************************
'************************************************************************************
'FUNCTION      : ASYMPTOTIC_EXPANSION_FUNC

'DESCRIPTION   : Evaluates exponential integral function for special case
'n=1 using formula of asymptotic expansion

'LIBRARY       : INTEGRATION
'GROUP         : ASYMPTOTIC
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 21/01/2009
'************************************************************************************
'************************************************************************************

Function ASYMPTOTIC_EXPANSION_FUNC(ByVal X_VAL As Double, _
Optional ByVal nLOOPS As Long = 100, _
Optional ByVal epsilon As Double = 2 ^ -52)
  
Dim i As Long

Dim TEMP_SUM As Double
Dim TEMP_FACT As Double

Dim TEMP_PROD As Double
Dim TEMP_PREV As Double

Dim TEMP_RESID As Double
  
On Error GoTo ERROR_LABEL

If X_VAL <= 0 Then: GoTo ERROR_LABEL
'Expansion function is not defined for x<0"

TEMP_FACT = (-Log(X_VAL) - 0.5772156649)

TEMP_SUM = 0
TEMP_PROD = 1
TEMP_RESID = -1
TEMP_PREV = 0
For i = 1 To nLOOPS
  TEMP_SUM = TEMP_SUM + ((-X_VAL) ^ i) / (i * TEMP_PROD)
  TEMP_PROD = TEMP_PROD * (i + 1)
  TEMP_RESID = TEMP_FACT - TEMP_SUM
  If (Abs(TEMP_PREV - TEMP_RESID) < epsilon) Then
    ASYMPTOTIC_EXPANSION_FUNC = TEMP_RESID
    Exit Function
  End If
  TEMP_PREV = TEMP_RESID
Next i

Exit Function
ERROR_LABEL:
ASYMPTOTIC_EXPANSION_FUNC = Err.number
End Function
