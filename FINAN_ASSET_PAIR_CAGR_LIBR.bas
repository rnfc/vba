Attribute VB_Name = "FINAN_ASSET_PAIR_CAGR_LIBR"


Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.

Function ASSETS_PAIR_DRAWDOWN_CAGR_PORT_FUNC(ByVal TICKER1_STR As Variant, _
ByVal TICKER2_STR As Variant, _
Optional ByVal START_DATE As Date, _
Optional ByVal END_DATE As Date, _
Optional ByVal INITIAL_INVESTMENT As Double = 1000, _
Optional ByVal WEIGHT1_VAL As Double = 0.75, _
Optional ByVal LONG1_FLAG As Boolean = True, _
Optional ByVal LONG2_FLAG As Boolean = True, _
Optional ByVal COUNT_BASIS As Double = 365, _
Optional ByVal OUTPUT As Integer = 0)

Dim i As Long
Dim k As Long

Dim ii As Long
Dim jj As Long

Dim NROWS As Long

Dim MAX1_VAL As Double
Dim MAX2_VAL As Double 'Max Drawdown
Dim CAGR_VAL As Double

Dim TEMP_ARR As Variant
Dim TEMP_MATRIX As Variant

Dim DATA1_MATRIX As Variant
Dim DATA2_MATRIX As Variant

On Error GoTo ERROR_LABEL

If IsArray(TICKER1_STR) = True Then
    DATA1_MATRIX = TICKER1_STR
    TICKER1_STR = "STOCK1"
Else
    DATA1_MATRIX = YAHOO_HISTORICAL_DATA_SERIE_FUNC(TICKER1_STR, START_DATE, END_DATE, "d", "DA", False, True, True)
End If

If IsArray(TICKER2_STR) = True Then
    DATA2_MATRIX = TICKER2_STR
    TICKER2_STR = "STOCK2"
Else
    DATA2_MATRIX = YAHOO_HISTORICAL_DATA_SERIE_FUNC(TICKER2_STR, START_DATE, END_DATE, "d", "DA", False, True, True)
End If

If UBound(DATA1_MATRIX, 1) <> UBound(DATA2_MATRIX, 1) Then: GoTo ERROR_LABEL

If OUTPUT > 1 Then
    ReDim TEMP_MATRIX(0 To 101, 1 To 3)
    TEMP_MATRIX(0, 1) = TICKER1_STR & " - WEIGHT"
    TEMP_MATRIX(0, 2) = "PORT MAX DRAWDOWN"
    TEMP_MATRIX(0, 3) = "CAGR"
    
    For i = 0 To 100
        TEMP_ARR = ASSETS_PAIR_DRAWDOWN_CAGR_PORT_FUNC(DATA1_MATRIX, DATA2_MATRIX, , , _
                   INITIAL_INVESTMENT, i / 100, LONG1_FLAG, LONG2_FLAG, COUNT_BASIS, 1)
        TEMP_MATRIX(i + 1, 1) = TEMP_ARR(LBound(TEMP_ARR))
        TEMP_MATRIX(i + 1, 2) = TEMP_ARR(LBound(TEMP_ARR) + 1)
        TEMP_MATRIX(i + 1, 3) = TEMP_ARR(LBound(TEMP_ARR) + 2)
    Next i
    Erase TEMP_ARR
    ASSETS_PAIR_DRAWDOWN_CAGR_PORT_FUNC = TEMP_MATRIX
    Exit Function
End If
NROWS = UBound(DATA1_MATRIX, 1)

If LONG1_FLAG = True Then ii = 1 Else ii = -1
If LONG2_FLAG = True Then jj = 1 Else jj = -1

'------------------------------------------------------------------------------
ReDim TEMP_MATRIX(0 To NROWS, 1 To 9)
'------------------------------------------------------------------------------
TEMP_MATRIX(0, 1) = "DATES"
TEMP_MATRIX(0, 2) = TICKER1_STR & " PRICES"
TEMP_MATRIX(0, 3) = TICKER2_STR & " PRICES"

TEMP_MATRIX(0, 4) = TICKER1_STR & " RETURNS"
TEMP_MATRIX(0, 5) = TICKER2_STR & " RETURNS"

TEMP_MATRIX(0, 6) = "PORTFOLIO RETURNS"
TEMP_MATRIX(0, 7) = Format(INITIAL_INVESTMENT, "#,##0.0") & " PORTFOLIO"

i = 1
TEMP_MATRIX(i, 1) = DATA1_MATRIX(i, 1)
TEMP_MATRIX(i, 2) = DATA1_MATRIX(i, 2)
TEMP_MATRIX(i, 3) = DATA2_MATRIX(i, 2)

TEMP_MATRIX(i, 4) = ""
TEMP_MATRIX(i, 5) = ""

TEMP_MATRIX(i, 6) = ""
TEMP_MATRIX(i, 7) = INITIAL_INVESTMENT
MAX1_VAL = TEMP_MATRIX(i, 7)

TEMP_MATRIX(i, 8) = MAX1_VAL
TEMP_MATRIX(i, 9) = TEMP_MATRIX(i, 8) - TEMP_MATRIX(i, 7)
MAX2_VAL = TEMP_MATRIX(i, 9)

For i = 2 To NROWS
    TEMP_MATRIX(i, 1) = DATA1_MATRIX(i, 1)
    TEMP_MATRIX(i, 2) = DATA1_MATRIX(i, 2)
    TEMP_MATRIX(i, 3) = DATA2_MATRIX(i, 2)
    
    TEMP_MATRIX(i, 4) = (TEMP_MATRIX(i, 2) / TEMP_MATRIX(i - 1, 2) - 1) * ii
    TEMP_MATRIX(i, 5) = (TEMP_MATRIX(i, 3) / TEMP_MATRIX(i - 1, 3) - 1) * jj
    
    TEMP_MATRIX(i, 6) = TEMP_MATRIX(i, 4) * WEIGHT1_VAL + TEMP_MATRIX(i, 5) * (1 - WEIGHT1_VAL)
    
    TEMP_MATRIX(i, 7) = TEMP_MATRIX(i - 1, 7) * (1 + TEMP_MATRIX(i, 6))
    If TEMP_MATRIX(i, 7) > MAX1_VAL Then: MAX1_VAL = TEMP_MATRIX(i, 7)
    TEMP_MATRIX(i, 8) = MAX1_VAL
    TEMP_MATRIX(i, 9) = TEMP_MATRIX(i, 8) - TEMP_MATRIX(i, 7)
    If TEMP_MATRIX(i, 9) > MAX2_VAL Then: MAX2_VAL = TEMP_MATRIX(i, 9)
Next i

k = TEMP_MATRIX(NROWS, 1) - TEMP_MATRIX(1, 1) 'No Days
CAGR_VAL = (TEMP_MATRIX(NROWS, 7) / TEMP_MATRIX(1, 7)) ^ (COUNT_BASIS / k) - 1

TEMP_MATRIX(0, 8) = "PORTFOLIO MAX - CAGR = " & Format(CAGR_VAL, "0.00%")
TEMP_MATRIX(0, 9) = "PORTFOLIO DRAWDOWN - MAX DD = " & Format(MAX2_VAL, "#,##0.0")

If OUTPUT = 0 Then
    ASSETS_PAIR_DRAWDOWN_CAGR_PORT_FUNC = TEMP_MATRIX
Else
    ASSETS_PAIR_DRAWDOWN_CAGR_PORT_FUNC = Array(WEIGHT1_VAL, MAX2_VAL, CAGR_VAL) 'OUTPUT = 1
End If

Exit Function
ERROR_LABEL:
ASSETS_PAIR_DRAWDOWN_CAGR_PORT_FUNC = Err.number
End Function
