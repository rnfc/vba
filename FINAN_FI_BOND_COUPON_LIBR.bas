Attribute VB_Name = "FINAN_FI_BOND_COUPON_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.

'************************************************************************************
'************************************************************************************
'FUNCTION      : COUPNUM_FUNC
'DESCRIPTION   : Calculates the number of coupons remaining on a bond
'LIBRARY       : BOND
'GROUP         : COUPON
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 21/01/2009
'************************************************************************************
'************************************************************************************

Function COUPNUM_FUNC(ByVal SETTLEMENT As Date, _
ByVal MATURITY As Date, _
Optional ByVal FREQUENCY As Integer = 2)

Dim i As Long
Dim j As Long
Dim DATE_VAL As Date

On Error GoTo ERROR_LABEL

i = 1000

If (FREQUENCY < 1) Then
    COUPNUM_FUNC = 0
    Exit Function
End If

If (MATURITY < SETTLEMENT) Then
    COUPNUM_FUNC = 0
    Exit Function
End If

If (SETTLEMENT = MATURITY) Then
    COUPNUM_FUNC = 1
    Exit Function
End If

j = 0
DATE_VAL = SETTLEMENT

Do While DATE_VAL < MATURITY
    j = j + 1
    DATE_VAL = EDATE_FUNC(DATE_VAL, (12 / FREQUENCY))
    If j > i Then: GoTo ERROR_LABEL
Loop
        
COUPNUM_FUNC = j

'YEARS_VAL = YEARFRAC_FUNC(SETTLEMENT, MATURITY, COUNT_BASIS)
'FRACTION_VAL = (YEARS_VAL * FREQUENCY - Int(YEARS_VAL * FREQUENCY)) / FREQUENCY
'YEARS_VAL = YEARS_VAL - FRACTION_VAL
' Correction if SETTLEMENT is ex-COUPON date
'If FRACTION_VAL = 0 Then
'    FRACTION_VAL = 1 / FREQUENCY
'    YEARS_VAL = YEARS_VAL - 1 / FREQUENCY
'End If
'NUMBER_COUPONS_FUNC = 1 + YEARS_VAL * FREQUENCY

Exit Function
ERROR_LABEL:
COUPNUM_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : COUPDAYBS_FUNC
'DESCRIPTION   : RETURNS THE NUMBER OF DAYS FROM PREVIOUS COUPON PAYMENT
'LIBRARY       : BOND
'GROUP         : COUPON
'ID            : 002
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 21/01/2009
'************************************************************************************
'************************************************************************************

Function COUPDAYBS_FUNC(ByVal SETTLEMENT As Date, _
ByVal MATURITY As Date, _
Optional ByVal FREQUENCY As Integer = 2, _
Optional ByVal COUNT_BASIS As Integer = 0)
  
Dim i As Long
Dim DATE_VAL As Date

On Error GoTo ERROR_LABEL
  
If (FREQUENCY < 1) Or (MATURITY <= SETTLEMENT) Then
    COUPDAYBS_FUNC = 0
    Exit Function
End If

DATE_VAL = COUPPCD_FUNC(SETTLEMENT, MATURITY, FREQUENCY)
i = COUNT_DAYS_FUNC(DATE_VAL, SETTLEMENT, COUNT_BASIS)
  
COUPDAYBS_FUNC = i

Exit Function
ERROR_LABEL:
COUPDAYBS_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : COUPDAYSNC_FUNC
'DESCRIPTION   : Calculates the time to the next coupon payment in days
'LIBRARY       : BOND
'GROUP         : COUPON
'ID            : 003
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 21/01/2009
'************************************************************************************
'************************************************************************************


Function COUPDAYSNC_FUNC(ByVal SETTLEMENT As Date, _
ByVal MATURITY As Date, _
Optional ByVal FREQUENCY As Integer = 2, _
Optional ByVal COUNT_BASIS As Integer = 0)

Dim i As Long
Dim DATE_VAL As Date

On Error GoTo ERROR_LABEL

If (FREQUENCY < 1) Or (MATURITY <= SETTLEMENT) Then
    COUPDAYSNC_FUNC = 0
    Exit Function
End If

DATE_VAL = COUPNCD_FUNC(SETTLEMENT, MATURITY, FREQUENCY)
i = COUNT_DAYS_FUNC(SETTLEMENT, DATE_VAL, COUNT_BASIS)
    
COUPDAYSNC_FUNC = i

' Calculates the time to the next coupon payment in years
'YEARS_VAL = YEARFRAC_FUNC(SETTLEMENT, MATURITY, COUNT_BASIS)
'FRACTION_VAL = (YEARS_VAL * FREQUENCY - Int(YEARS_VAL * FREQUENCY)) / FREQUENCY
' Correction if SETTLEMENT is ex-COUPON date
'If FRACTION_VAL = 0 Then: FRACTION_VAL = 1 / FREQUENCY
'TIME_NEXT_COUPON_FUNC = FRACTION_VAL

Exit Function
ERROR_LABEL:
COUPDAYSNC_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : COUPNCD_FUNC
'DESCRIPTION   : RETURNS THE NEXT COUPON PAYMENT DATE
'LIBRARY       : BOND
'GROUP         : COUPON
'ID            : 004
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 21/01/2009
'************************************************************************************
'************************************************************************************


Function COUPNCD_FUNC(ByVal SETTLEMENT As Date, _
ByVal MATURITY As Date, _
Optional ByVal FREQUENCY As Integer = 2)

Dim j As Long

On Error GoTo ERROR_LABEL

If (FREQUENCY < 1) Or (MATURITY <= SETTLEMENT) Then
    COUPNCD_FUNC = 0
    Exit Function
End If
   
If MATURITY = SETTLEMENT Then
  COUPNCD_FUNC = MATURITY
Else
  j = COUPNUM_FUNC(SETTLEMENT, MATURITY, FREQUENCY) - 1
  COUPNCD_FUNC = EDATE_FUNC(MATURITY, -j * (12 / FREQUENCY))
End If

Exit Function
ERROR_LABEL:
COUPNCD_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : COUPPCD_FUNC
'DESCRIPTION   : Calculates days from previous coupon date until settlement date
'LIBRARY       : BOND
'GROUP         : COUPON
'ID            : 005
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 21/01/2009
'************************************************************************************
'************************************************************************************

Function COUPPCD_FUNC(ByVal SETTLEMENT As Date, _
ByVal MATURITY As Date, _
Optional ByVal FREQUENCY As Integer = 2)

Dim DATE_VAL As Date

On Error GoTo ERROR_LABEL

If (FREQUENCY < 1) Or (MATURITY <= SETTLEMENT) Then
    COUPPCD_FUNC = 0
    Exit Function
End If
  
DATE_VAL = COUPNCD_FUNC(SETTLEMENT, MATURITY, FREQUENCY)
COUPPCD_FUNC = EDATE_FUNC(DATE_VAL, -12 / FREQUENCY)

Exit Function
ERROR_LABEL:
COUPPCD_FUNC = Err.number
End Function

