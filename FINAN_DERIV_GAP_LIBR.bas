Attribute VB_Name = "FINAN_DERIV_GAP_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.


'**********************************************************************************
'**********************************************************************************
'FUNCTION      : GAP_OPTION_FUNC
'DESCRIPTION   : Gap options
'LIBRARY       : DERIVATIVES
'GROUP         : GAP
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/31/2009
'**********************************************************************************
'**********************************************************************************

Function GAP_OPTION_FUNC(ByVal SPOT As Double, _
ByVal STRIKE_A As Double, _
ByVal STRIKE_B As Double, _
ByVal TENOR As Double, _
ByVal RATE As Double, _
ByVal CARRY_COST As Double, _
ByVal SIGMA As Double, _
Optional ByVal OPTION_FLAG As Integer = 1, _
Optional ByVal CND_TYPE As Integer = 0)

Dim D1_VAL As Double
Dim D2_VAL As Double

On Error GoTo ERROR_LABEL

If OPTION_FLAG <> 1 Then: OPTION_FLAG = -1

D1_VAL = (Log(SPOT / STRIKE_A) + (CARRY_COST + SIGMA ^ 2 / 2) * TENOR) / _
(SIGMA * Sqr(TENOR))

D2_VAL = D1_VAL - SIGMA * Sqr(TENOR)

Select Case OPTION_FLAG
    Case 1 ', "c", "call"
        GAP_OPTION_FUNC = SPOT * Exp((CARRY_COST - RATE) * TENOR) * _
        CND_FUNC(D1_VAL, CND_TYPE) - STRIKE_B * Exp(-RATE * TENOR) * _
        CND_FUNC(D2_VAL, CND_TYPE)
    Case Else '-1 ', "p", "put"
        GAP_OPTION_FUNC = STRIKE_B * Exp(-RATE * TENOR) * _
        CND_FUNC(-D2_VAL, CND_TYPE) - SPOT * Exp((CARRY_COST - RATE) * _
        TENOR) * CND_FUNC(-D1_VAL, CND_TYPE)
End Select
    
Exit Function
ERROR_LABEL:
GAP_OPTION_FUNC = Err.number
End Function
