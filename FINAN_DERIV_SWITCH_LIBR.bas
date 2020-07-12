Attribute VB_Name = "FINAN_DERIV_SWITCH_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.

'**********************************************************************************
'**********************************************************************************
'FUNCTION      : TIME_SWITCH_OPTION_FUNC
'DESCRIPTION   : Time switch options (discrete)
'LIBRARY       : DERIVATIVES
'GROUP         : SWITCH
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/29/2009
'**********************************************************************************
'**********************************************************************************

Function TIME_SWITCH_OPTION_FUNC(ByVal SPOT As Double, _
ByVal STRIKE As Double, _
ByVal ACCUMULATED As Double, _
ByVal EXPIRATION As Double, _
ByVal QUANTITY As Double, _
ByVal DELTA_TIME As Double, _
ByVal RATE As Double, _
ByVal CARRY_COST As Double, _
ByVal SIGMA As Double, _
Optional ByVal OPTION_FLAG As Integer = 1, _
Optional ByVal CND_TYPE As Integer = 0)

'ACCUMULATED: Accumulated amount
'EXPIRATION: Time to maturity
'QUANTITY: Number of time units fulfilled
'DELTA_TIME: Time INTERVAL

Dim i As Long
Dim j As Long
Dim k As Long

Dim TEMP_VAL As Double
Dim TEMP_SUM As Double

    On Error GoTo ERROR_LABEL
    
    Select Case OPTION_FLAG
        Case 1 ', "CALL", "C"
            k = 1
        Case Else '-1 ', "PUT", "P"
            k = -1
    End Select
    
    j = EXPIRATION / DELTA_TIME
    TEMP_SUM = 0
    
    For i = 1 To j
        TEMP_VAL = (Log(SPOT / STRIKE) + (CARRY_COST - SIGMA ^ 2 / 2) * i * _
        DELTA_TIME) / (SIGMA * Sqr(i * DELTA_TIME))
        
        TEMP_SUM = TEMP_SUM + CND_FUNC(k * TEMP_VAL, CND_TYPE) * DELTA_TIME
    Next i
    
    TIME_SWITCH_OPTION_FUNC = ACCUMULATED * Exp(-RATE * EXPIRATION) * TEMP_SUM + _
                        DELTA_TIME * ACCUMULATED * Exp(-RATE * _
                        EXPIRATION) * QUANTITY
    
Exit Function
ERROR_LABEL:
TIME_SWITCH_OPTION_FUNC = Err.number
End Function
