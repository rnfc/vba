Attribute VB_Name = "FINAN_ASSET_TA_EMA_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.

Function ASSET_TA_EMA_DEVIATION_FUNC(ByVal TICKER_STR As Variant, _
Optional ByVal START_DATE As Date, _
Optional ByVal END_DATE As Date, _
Optional ByVal EMA_PERIOD As Long = 20)

Dim i As Long
Dim j As Long

Dim NROWS As Long
Dim NCOLUMNS As Long

Dim ALPHA_VAL As Double

Dim DATA_MATRIX As Variant
Dim TEMP_MATRIX As Variant

On Error GoTo ERROR_LABEL

If IsArray(TICKER_STR) = False Then
    DATA_MATRIX = YAHOO_HISTORICAL_DATA_SERIE_FUNC(TICKER_STR, START_DATE, END_DATE, _
                  "d", "DOHLCVA", False, True, True)
Else
    DATA_MATRIX = TICKER_STR
End If

NROWS = UBound(DATA_MATRIX, 1)
NCOLUMNS = UBound(DATA_MATRIX, 2)

ALPHA_VAL = 1 - 2 / (EMA_PERIOD + 1)
'-----------------------------------------------------------------------------
ReDim TEMP_MATRIX(0 To NROWS, 1 To 9)
'-----------------------------------------------------------------------------
TEMP_MATRIX(0, 1) = "DATE"
TEMP_MATRIX(0, 2) = "OPEN"
TEMP_MATRIX(0, 3) = "HIGH"
TEMP_MATRIX(0, 4) = "LOW"
TEMP_MATRIX(0, 5) = "CLOSE"
TEMP_MATRIX(0, 6) = "VOLUME"
TEMP_MATRIX(0, 7) = "ADJ CLOSE"
TEMP_MATRIX(0, 8) = "EMA " & EMA_PERIOD
TEMP_MATRIX(0, 9) = "(P-EMA)/P"

i = 1
For j = 1 To NCOLUMNS: TEMP_MATRIX(i, j) = DATA_MATRIX(i, j): Next j
TEMP_MATRIX(i, 6) = TEMP_MATRIX(i, 6) / 1000
TEMP_MATRIX(i, 8) = ALPHA_VAL * TEMP_MATRIX(i, 7) + (1 - ALPHA_VAL) * TEMP_MATRIX(i, 7)
TEMP_MATRIX(i, 9) = 1 - TEMP_MATRIX(i, 8) / TEMP_MATRIX(i, 7)

For i = 2 To NROWS
    For j = 1 To NCOLUMNS: TEMP_MATRIX(i, j) = DATA_MATRIX(i, j): Next j
    TEMP_MATRIX(i, 6) = TEMP_MATRIX(i, 6) / 1000
    TEMP_MATRIX(i, 8) = ALPHA_VAL * TEMP_MATRIX(i - 1, 8) + (1 - ALPHA_VAL) * TEMP_MATRIX(i, 7)
    TEMP_MATRIX(i, 9) = 1 - TEMP_MATRIX(i, 8) / TEMP_MATRIX(i, 7)
Next i

ASSET_TA_EMA_DEVIATION_FUNC = TEMP_MATRIX

Exit Function
ERROR_LABEL:
ASSET_TA_EMA_DEVIATION_FUNC = Err.number
End Function
