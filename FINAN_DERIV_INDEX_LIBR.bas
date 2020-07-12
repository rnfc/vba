Attribute VB_Name = "FINAN_DERIV_INDEX_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
'default first index of arrays.


'**********************************************************************************
'**********************************************************************************
'FUNCTION      : INDEX_OPTION_FUNC
'DESCRIPTION   : Merton (1973) Options on stock indices
'LIBRARY       : DERIVATIVES
'GROUP         : INDEX
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/29/2009
'**********************************************************************************
'**********************************************************************************

Function INDEX_OPTION_FUNC(ByVal SPOT As Double, _
ByVal STRIKE As Double, _
ByVal EXPIRATION As Double, _
ByVal RATE As Double, _
ByVal DIVD As Double, _
ByVal SIGMA As Double, _
Optional ByVal OPTION_FLAG As Integer = 1, _
Optional ByVal CND_TYPE As Integer = 0)
    
Dim D1_VAL As Double
Dim D2_VAL As Double

On Error GoTo ERROR_LABEL

If OPTION_FLAG <> 1 Then: OPTION_FLAG = -1

D1_VAL = (Log(SPOT / STRIKE) + (RATE - DIVD + SIGMA ^ 2 / 2) * EXPIRATION) / _
(SIGMA * Sqr(EXPIRATION))
D2_VAL = D1_VAL - SIGMA * Sqr(EXPIRATION)

Select Case OPTION_FLAG
    Case 1 ', "CALL", "C"
        INDEX_OPTION_FUNC = SPOT * Exp(-DIVD * EXPIRATION) * _
        CND_FUNC(D1_VAL, CND_TYPE) - STRIKE * _
        Exp(-RATE * EXPIRATION) * CND_FUNC(D2_VAL, CND_TYPE)
    Case Else '-1 ', "PUT", "P"
        INDEX_OPTION_FUNC = STRIKE * Exp(-RATE * EXPIRATION) * _
        CND_FUNC(-D2_VAL, CND_TYPE) - SPOT * _
        Exp(-DIVD * EXPIRATION) * CND_FUNC(-D1_VAL, CND_TYPE)
End Select
    
Exit Function
ERROR_LABEL:
INDEX_OPTION_FUNC = Err.number
End Function
