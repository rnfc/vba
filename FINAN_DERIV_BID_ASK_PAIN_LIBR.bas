Attribute VB_Name = "FINAN_DERIV_BID_ASK_PAIN_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.


'************************************************************************************
'************************************************************************************
'FUNCTION      : WEIGHTED_AVERAGE_STRIKE_PAIN_FUNC

'DESCRIPTION   : Weighted Average Strike of Pain
'This is the weighted average of the "Strike of Pain" calculation for
'each strike price

'SPOT_PRICE: Current price of the Stock/ETF/Index
'CALL/PUT_OPEN_RNG --> Open Interest

'LIBRARY       : DERIVATIVES
'GROUP         : BID_ASK_PAIN
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/29/2009
'************************************************************************************
'************************************************************************************


Function WEIGHTED_AVERAGE_STRIKE_PAIN_FUNC(ByVal SPOT_PRICE As Double, _
ByRef STRIKE_RNG As Variant, _
ByRef CALL_OPEN_RNG As Variant, _
ByRef PUT_OPEN_RNG As Variant, _
Optional ByVal OUTPUT As Integer = 1)

Dim i As Long
Dim NROWS As Long

Dim TEMP_VAL As Double
Dim TEMP_MAX As Double

Dim ATEMP_SUM As Double
Dim BTEMP_SUM As Double

Dim STRIKE_VECTOR As Variant
Dim CALL_OPEN_VECTOR As Variant
Dim PUT_OPEN_VECTOR As Variant
Dim TEMP_MATRIX As Variant

On Error GoTo ERROR_LABEL

STRIKE_VECTOR = STRIKE_RNG
If UBound(STRIKE_VECTOR, 1) = 1 Then: _
    STRIKE_VECTOR = MATRIX_TRANSPOSE_FUNC(STRIKE_VECTOR)

CALL_OPEN_VECTOR = CALL_OPEN_RNG
If UBound(CALL_OPEN_VECTOR, 1) = 1 Then: _
    CALL_OPEN_VECTOR = MATRIX_TRANSPOSE_FUNC(CALL_OPEN_VECTOR)

PUT_OPEN_VECTOR = PUT_OPEN_RNG
If UBound(PUT_OPEN_VECTOR, 1) = 1 Then: _
    PUT_OPEN_VECTOR = MATRIX_TRANSPOSE_FUNC(PUT_OPEN_VECTOR)

If UBound(STRIKE_VECTOR, 1) <> UBound(CALL_OPEN_VECTOR, 1) Then: GoTo ERROR_LABEL
If UBound(STRIKE_VECTOR, 1) <> UBound(PUT_OPEN_VECTOR, 1) Then: GoTo ERROR_LABEL
NROWS = UBound(STRIKE_VECTOR, 1)

ReDim TEMP_MATRIX(0 To NROWS, 1 To 3)

TEMP_MATRIX(0, 1) = "PAIN"
TEMP_MATRIX(0, 2) = "STRIKE"

TEMP_MAX = -2 ^ 52
ATEMP_SUM = 0
BTEMP_SUM = 0
For i = 1 To NROWS
    
    If IsError( _
        MAXIMUM_FUNC(0, (SPOT_PRICE - STRIKE_VECTOR(i, 1)) * CALL_OPEN_VECTOR(i, 1)) + _
        MAXIMUM_FUNC(0, (PUT_OPEN_VECTOR(i, 1) - SPOT_PRICE) * PUT_OPEN_VECTOR(i, 1))) Then
            TEMP_VAL = 0
    Else
            TEMP_VAL = _
        MAXIMUM_FUNC(0, (SPOT_PRICE - STRIKE_VECTOR(i, 1)) * CALL_OPEN_VECTOR(i, 1)) + _
        MAXIMUM_FUNC(0, (PUT_OPEN_VECTOR(i, 1) - SPOT_PRICE) * PUT_OPEN_VECTOR(i, 1))
    End If
    
    TEMP_MATRIX(i, 1) = TEMP_VAL
    If TEMP_MATRIX(i, 1) > TEMP_MAX Then: TEMP_MAX = TEMP_MATRIX(i, 1)
    TEMP_MATRIX(i, 2) = STRIKE_VECTOR(i, 1)
    
    ATEMP_SUM = ATEMP_SUM + (TEMP_MATRIX(i, 1) * TEMP_MATRIX(i, 2))
    BTEMP_SUM = BTEMP_SUM + TEMP_MATRIX(i, 2)
Next i

TEMP_MATRIX(0, 3) = "GRAPH: " & Format(ATEMP_SUM / BTEMP_SUM, "#,##0.00")

'---------------------------------------------------------------------------
Select Case OUTPUT
'---------------------------------------------------------------------------
Case 0
'---------------------------------------------------------------------------
    WEIGHTED_AVERAGE_STRIKE_PAIN_FUNC = ATEMP_SUM / BTEMP_SUM
'---------------------------------------------------------------------------
Case Else
'---------------------------------------------------------------------------
    For i = 1 To NROWS
        TEMP_MATRIX(i, 3) = String(Round(20 * _
        TEMP_MATRIX(i, 1) / TEMP_MAX, 0), "|")
    Next i
    If OUTPUT = 1 Then
        WEIGHTED_AVERAGE_STRIKE_PAIN_FUNC = TEMP_MATRIX
    Else
        WEIGHTED_AVERAGE_STRIKE_PAIN_FUNC = Array(TEMP_MATRIX, ATEMP_SUM / BTEMP_SUM)
    End If
'---------------------------------------------------------------------------
End Select
'---------------------------------------------------------------------------

Exit Function
ERROR_LABEL:
WEIGHTED_AVERAGE_STRIKE_PAIN_FUNC = Err.number
End Function


'The Strike of Pain is the strike price with the lowest In-the-Money
'value for both calls and puts on a given stock for a given expiration
'date. This is the strike where the long option traders will lose the
'most money and the short option traders will gain the most value.
'The Strike of Pain is calculated by taking the open interest of the
'various In-the-Money strike prices multiplied by the amount that the
'call or put is ITM. The ITM values are added together, and the strike
'with the lowest monetary value is the Strike of Pain. The Strike of
'Pain calculation is based on the theory of MAXIMUM_FUNC-Pain(tm).

'This theory takes the value that an option is ITM if
'the stock was trading at a different strike. That value (for the calls or
'puts) is then multiplied by its open interest. This value for all ITM calls
'and puts is then added together. The strike that reflects the lowest ITM
'value is the Strike of Pain. Let’s look at a brief example:
'XYZ is trading at $24.50.

'We would then look at the Full Call and Put Chain for the nearest month
'that would show the Open Interest:
'Call Strike  Call OI  Put Strike  Put OI
'17.5  10  17.5  15
'20  25  20  23
'22.5  50  22.5  35
'25  100  25  50
'27.5  40  27.5  20
'30  5  30  2

'The Strike of Pain Calculator will quickly analyze the total ITM value at
'each strike price. For example, if the stock was trading at 17.5 all of the
'calls would be OTM. However, all of the put strikes (excluding the 17.5)
'would be ITM. The 20 strike put would be $2.50 ITM, the 22.5 put would be
'$5.00 ITM, the 25 put would be $7.50 ITM and so on. To determine the ITM Value
'of the 20-strike put if the stock was trading at $17.50 we would multiply the
'20 put Open Interest, 23, times $2.50 (the amount the put is ITM). This step
'would be repeated for the rest of the ITM puts and then all the values would
'be added together.

'If the stock was trading at $25, the 22.5 call would be $2.50 ITM, the 20 call
'would be $5.00 ITM, and the 17.5 call would be $7.50 ITM. Conversely, the 27.5
'put would be $2.50 ITM and the 30 put would be $5.00 ITM. Again, the ITM values
'would by multiplied times the corresponding strikes Open Interest and the call
'and put ITM values would be added together.

'The strike price with the lowest total ITM value is the Strike of Pain. This is
'the strike price that the stock may have a tendency to gravitate to at expiration.

'Why is this important?

'It is believed that stocks will have a tendency to gravitate to the
'Strike of Pain on expiration day. There is evidence that this tendency
'exists, but there is still debate on whether this is caused by market
'forces or by mere chance. It is important to note that such events such
'as market upheaval, price momentum, breaking news or extreme earnings
'reports will negate the Strike of Pain value. It is also important to
'note that the Strike of Pain price is more reliable within the last
'week until expiration. The Strike of Pain value for a 6-month out target
'month may be inaccurate as the open interest can drastically change between
'now and the 6-month expiration date.
