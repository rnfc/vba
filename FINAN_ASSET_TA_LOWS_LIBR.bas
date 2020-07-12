Attribute VB_Name = "FINAN_ASSET_TA_LOWS_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.

Function ASSET_TA_NEW_LOWS_FUNC(ByRef TICKER_STR As Variant, _
Optional ByVal START_DATE As Date, _
Optional ByVal END_DATE As Date, _
Optional ByVal NUMBER_DAYS As Long = 10)

'Number Days: lows in previous x days

Dim i As Long
Dim j As Long
Dim l As Long

Dim NROWS As Long

Dim DATA_MATRIX As Variant
Dim TEMP_MATRIX As Variant

On Error GoTo ERROR_LABEL

If IsArray(TICKER_STR) = True Then
    DATA_MATRIX = TICKER_STR
Else
    DATA_MATRIX = YAHOO_HISTORICAL_DATA_SERIE_FUNC(TICKER_STR, START_DATE, END_DATE, "d", "DOHLCVA", False, True, True)
End If
NROWS = UBound(DATA_MATRIX, 1)
If NUMBER_DAYS < 2 Or NUMBER_DAYS > NROWS - 1 Then: GoTo ERROR_LABEL

ReDim TEMP_MATRIX(0 To NROWS, 1 To 9)

TEMP_MATRIX(0, 1) = "DATE"
TEMP_MATRIX(0, 2) = "OPEN"
TEMP_MATRIX(0, 3) = "HIGH"
TEMP_MATRIX(0, 4) = "LOW"
TEMP_MATRIX(0, 5) = "CLOSE"
TEMP_MATRIX(0, 6) = "VOLUME"
TEMP_MATRIX(0, 7) = "ADJ.CLOSE"
TEMP_MATRIX(0, 8) = NUMBER_DAYS & "-DAYS LOWS"

i = 1
For j = 1 To 7: TEMP_MATRIX(i, j) = DATA_MATRIX(i, j): Next j
TEMP_MATRIX(i, 6) = TEMP_MATRIX(i, 6) / 1000
TEMP_MATRIX(i, 8) = 0
TEMP_MATRIX(i, 9) = 0

l = TEMP_MATRIX(i, 8)
For i = 2 To NROWS - 1
    For j = 1 To 7: TEMP_MATRIX(i, j) = DATA_MATRIX(i, j): Next j
    TEMP_MATRIX(i, 6) = TEMP_MATRIX(i, 6) / 1000
    TEMP_MATRIX(i, 8) = IIf(DATA_MATRIX(i, 5) < DATA_MATRIX(i - 1, 5) And DATA_MATRIX(i, 5) < DATA_MATRIX(i + 1, 5), 1, 0)
    If i > NUMBER_DAYS Then
        l = l - TEMP_MATRIX(i - NUMBER_DAYS, 8)
        l = l + TEMP_MATRIX(i, 8)
    Else
        l = l + TEMP_MATRIX(i, 8)
    End If
    TEMP_MATRIX(i, 9) = l
Next i

For j = 1 To 7: TEMP_MATRIX(i, j) = DATA_MATRIX(i, j): Next j
TEMP_MATRIX(i, 6) = TEMP_MATRIX(i, 6) / 1000
TEMP_MATRIX(i, 8) = 0

l = l - TEMP_MATRIX(i - NUMBER_DAYS, 8)
l = l + TEMP_MATRIX(i, 8)
TEMP_MATRIX(i, 9) = l
TEMP_MATRIX(0, 9) = "COUNT-LOWS = " & l

ASSET_TA_NEW_LOWS_FUNC = TEMP_MATRIX

Exit Function
ERROR_LABEL:
ASSET_TA_NEW_LOWS_FUNC = Err.number
End Function
