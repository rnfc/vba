Attribute VB_Name = "FINAN_FUNDAM_SERIE_LIBR"


Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.
                    
Function PER_SHARE_TIME_SERIE_FUNC(ByRef PRICES_RNG As Variant, _
ByRef FUNDAM_RNG As Variant)

Dim i As Long
Dim j As Long
Dim k As Long
Dim NROWS As Long
Dim NSIZE As Long
Dim PRICES_VECTOR As Variant
Dim FUNDAM_VECTOR As Variant
Dim TEMP_MATRIX As Variant

On Error GoTo ERROR_LABEL

PRICES_VECTOR = PRICES_RNG 'C1=DATES / C2=DATA --> Ascending Order
NROWS = UBound(PRICES_VECTOR, 1)
FUNDAM_VECTOR = FUNDAM_RNG 'C1=DATES / C2=DATA --> Descending Order
NSIZE = UBound(FUNDAM_VECTOR, 1)

ReDim TEMP_MATRIX(0 To NROWS, 1 To 4)
TEMP_MATRIX(0, 1) = "DATE"
TEMP_MATRIX(0, 2) = "PRICE"
TEMP_MATRIX(0, 3) = "METRIC"
TEMP_MATRIX(0, 4) = "YIELD" 'Per Share

For i = 1 To NROWS 'Sanitized Montly Prices
    TEMP_MATRIX(i, 1) = DateSerial(Year(PRICES_VECTOR(i, 1)), Month(PRICES_VECTOR(i, 1)), 1)
    TEMP_MATRIX(i, 2) = PRICES_VECTOR(i, 2)
Next i
k = 1
For j = NSIZE - 1 To 1 Step -1
    Do While FUNDAM_VECTOR(j, 1) >= TEMP_MATRIX(k, 1)
        TEMP_MATRIX(k, 3) = FUNDAM_VECTOR(j + 1, 2)
        k = k + 1
        If k > NROWS Then: GoTo 1983
    Loop
Next j
For i = k To NROWS: TEMP_MATRIX(i, 3) = FUNDAM_VECTOR(1, 2): Next i
1983:
For i = 1 To NROWS: TEMP_MATRIX(i, 4) = TEMP_MATRIX(i, 3) / TEMP_MATRIX(i, 2): Next i
PER_SHARE_TIME_SERIE_FUNC = TEMP_MATRIX

Exit Function
ERROR_LABEL:
PER_SHARE_TIME_SERIE_FUNC = Err.number
End Function



