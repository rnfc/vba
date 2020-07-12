Attribute VB_Name = "FINAN_FI_ANNUITY_LIBR"

'-----------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------
'The following functions are extremely useful to derive
'the present value of an investment. The present value is the total
'amount that a series of future payments (annuities) is worth now.
'-----------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------

'-----------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------
Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.
'-----------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------


'**********************************************************************************
'**********************************************************************************
'FUNCTION      : SUM_ANNUITY_FUNC
'DESCRIPTION   : Summ Annuity Function
'LIBRARY       : FIXED INCOME
'GROUP         : ANNUITY
'ID            : 001
'LAST UPDATE   : 10-06-2008
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'**********************************************************************************
'**********************************************************************************

Function SUM_ANNUITY_FUNC(ByVal PAYMENT_TODAY As Double, _
ByVal PAYMENT_TENOR As Double, _
ByVal PAYMENT_FIRST As Double, _
ByVal TENOR As Double)

On Error GoTo ERROR_LABEL

SUM_ANNUITY_FUNC = 0.5 * TENOR * (PAYMENT_FIRST + PAYMENT_TENOR) + PAYMENT_TODAY

Exit Function
ERROR_LABEL:
SUM_ANNUITY_FUNC = Err.number
End Function


'**********************************************************************************
'**********************************************************************************
'FUNCTION      : ANNUITY
'DESCRIPTION   : ANNUITY CALCULATION
'LIBRARY       : FIXED INCOME
'GROUP         : ANNUITY
'ID            : 002
'LAST UPDATE   : 10-06-2008
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'**********************************************************************************
'**********************************************************************************

Function ANNUITY_FUNC(ByVal PAYMENT_TODAY As Double, _
ByVal PV_ANNUITY As Double, _
ByVal RATE As Double, _
ByVal TENOR As Double)

On Error GoTo ERROR_LABEL

ANNUITY_FUNC = (PV_ANNUITY - PAYMENT_TODAY) / ((1 - (1 / ((1 + RATE) ^ TENOR))) / RATE)

Exit Function
ERROR_LABEL:
ANNUITY_FUNC = Err.number
End Function


'**********************************************************************************
'**********************************************************************************
'FUNCTION      : PV_ANNUITY_FUNC
'DESCRIPTION   : PV ANNUITY CALCULATION
'LIBRARY       : FIXED INCOME
'GROUP         : ANNUITY
'ID            : 003
'LAST UPDATE   : 10-06-2008
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'**********************************************************************************
'**********************************************************************************

Function PV_ANNUITY_FUNC(ByVal PAYMENT_TODAY As Double, _
ByVal ANNUITY As Double, _
ByVal RATE As Double, _
ByVal TENOR As Double)

On Error GoTo ERROR_LABEL

PV_ANNUITY_FUNC = (ANNUITY * ((1 - (1 / ((1 + RATE) ^ TENOR))) / RATE)) + PAYMENT_TODAY

Exit Function
ERROR_LABEL:
PV_ANNUITY_FUNC = Err.number
End Function

'**********************************************************************************
'**********************************************************************************
'FUNCTION      : PV_GROWTH_ANNUITY1_FUNC
'DESCRIPTION   : PV GROWING ANNUITY CALCULATION (A)
'LIBRARY       : FIXED INCOME
'GROUP         : ANNUITY
'ID            : 004
'LAST UPDATE   : 10-06-2008
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'**********************************************************************************
'**********************************************************************************

Function PV_GROWTH_ANNUITY1_FUNC(ByVal PAYMENT_TODAY As Double, _
ByVal ANNUITY As Double, _
ByVal GROWTH_RATE As Double, _
ByVal RATE As Double, _
ByVal TENOR As Double)

On Error GoTo ERROR_LABEL

PV_GROWTH_ANNUITY1_FUNC = (1 - ((1 + GROWTH_RATE) / (RATE + 1)) ^ TENOR) * ANNUITY * ((1 + GROWTH_RATE) / (RATE - GROWTH_RATE)) + PAYMENT_TODAY

Exit Function
ERROR_LABEL:
PV_GROWTH_ANNUITY1_FUNC = Err.number
End Function


'**********************************************************************************
'**********************************************************************************
'FUNCTION      : PV_GROWTH_ANNUITY2_FUNC
'DESCRIPTION   : PV GROWING ANNUITY CALCULATION (B)
'LIBRARY       : FIXED INCOME
'GROUP         : ANNUITY
'ID            : 005
'LAST UPDATE   : 10-06-2008
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'**********************************************************************************
'**********************************************************************************

Function PV_GROWTH_ANNUITY2_FUNC(ByVal PAYMENT_TODAY As Double, _
ByVal DELTA_ANNUITY As Double, _
ByVal RATE As Double, _
ByVal TENOR As Double)

Dim i As Long
Dim TEMP_SUM As Double

On Error GoTo ERROR_LABEL

TEMP_SUM = 0
For i = 1 To TENOR: TEMP_SUM = TEMP_SUM + (i / (1 + RATE) ^ i): Next i

PV_GROWTH_ANNUITY2_FUNC = DELTA_ANNUITY * TEMP_SUM + PAYMENT_TODAY

Exit Function
ERROR_LABEL:
PV_GROWTH_ANNUITY2_FUNC = Err.number
End Function


'**********************************************************************************
'**********************************************************************************
'FUNCTION      : PV_PERPETUITY_FUNC
'DESCRIPTION   : PV PERPETUITY CALCULATION
'LIBRARY       : FIXED INCOME
'GROUP         : ANNUITY
'ID            : 006
'LAST UPDATE   : 10-06-2008
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'**********************************************************************************
'**********************************************************************************

Function PV_PERPETUITY_FUNC(ByVal PAYMENT_TODAY As Double, _
ByVal ANNUITY As Double, _
ByVal RATE As Double)

On Error GoTo ERROR_LABEL

PV_PERPETUITY_FUNC = (ANNUITY / RATE) + PAYMENT_TODAY

Exit Function
ERROR_LABEL:
PV_PERPETUITY_FUNC = Err.number
End Function


'**********************************************************************************
'**********************************************************************************
'FUNCTION      : PV_GROWTH_PERPETUITY1_FUNC
'DESCRIPTION   : PV GROWING PERPETUITY CALCULATION (A)
'LIBRARY       : FIXED INCOME
'GROUP         : ANNUITY
'ID            : 007
'LAST UPDATE   : 10-06-2008
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'**********************************************************************************
'**********************************************************************************

Function PV_GROWTH_PERPETUITY1_FUNC(ByVal PAYMENT_TODAY As Double, _
ByVal ANNUITY As Double, _
ByVal GROWTH_RATE As Double, _
ByVal RATE As Double)

On Error GoTo ERROR_LABEL

PV_GROWTH_PERPETUITY1_FUNC = PAYMENT_TODAY + ANNUITY * ((1 + GROWTH_RATE) / (RATE - GROWTH_RATE))

Exit Function
ERROR_LABEL:
PV_GROWTH_PERPETUITY1_FUNC = Err.number
End Function


'**********************************************************************************
'**********************************************************************************
'FUNCTION      : PV_GROWTH_PERPETUITY2_FUNC
'DESCRIPTION   : PV GROWING PERPETUITY CALCULATION (B)
'LIBRARY       : FIXED INCOME
'GROUP         : ANNUITY
'ID            : 008
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'**********************************************************************************
'**********************************************************************************

Function PV_GROWTH_PERPETUITY2_FUNC(ByVal PAYMENT_TODAY As Double, _
ByVal ANNUITY As Double, _
ByVal DELTA_ANNUITY As Double, _
ByVal RATE As Double)

On Error GoTo ERROR_LABEL

PV_GROWTH_PERPETUITY2_FUNC = PV_PERPETUITY_FUNC(PAYMENT_TODAY, ANNUITY, RATE) + ((DELTA_ANNUITY / RATE) / RATE)

Exit Function
ERROR_LABEL:
PV_GROWTH_PERPETUITY2_FUNC = Err.number
End Function
