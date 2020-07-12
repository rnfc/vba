Attribute VB_Name = "FINAN_ASSET_TA_CANDLE_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.


'************************************************************************************
'************************************************************************************
'FUNCTION      : ASSET_CANDLESTICK_FUNC
'DESCRIPTION   :
'LIBRARY       : FINAN_ASSET
'GROUP         :
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA

'************************************************************************************
'************************************************************************************

'// PERFECT

Function ASSET_CANDLESTICK_FUNC(ByVal TICKER_STR As Variant, _
Optional ByVal START_DATE As Date, _
Optional ByVal END_DATE As Date, _
Optional ByVal ERROR_VAL As Double = 0.2, _
Optional ByVal SHADOW_VAL As Double = 0.15)

Dim i As Long
Dim j As Long

Dim NROWS As Long
Dim NCOLUMNS As Long

Dim TEMP_MATRIX As Variant
Dim DATA_MATRIX As Variant

On Error GoTo ERROR_LABEL

If IsArray(TICKER_STR) = False Then
    DATA_MATRIX = YAHOO_HISTORICAL_DATA_SERIE_FUNC(TICKER_STR, START_DATE, END_DATE, _
                  "d", "DOHLCV", False, True, True)
Else
    DATA_MATRIX = TICKER_STR
End If
NROWS = UBound(DATA_MATRIX, 1)
NCOLUMNS = UBound(DATA_MATRIX, 2)

ReDim TEMP_MATRIX(0 To NROWS, 1 To 9)
TEMP_MATRIX(0, 1) = "DATE"
TEMP_MATRIX(0, 2) = "OPEN"
TEMP_MATRIX(0, 3) = "HIGH"
TEMP_MATRIX(0, 4) = "LOW"
TEMP_MATRIX(0, 5) = "CLOSE"
TEMP_MATRIX(0, 6) = "VOLUME"

For i = 1 To NROWS
    For j = 1 To NCOLUMNS: TEMP_MATRIX(i, j) = DATA_MATRIX(i, j): Next j
    
    If (TEMP_MATRIX(i, 3) - TEMP_MATRIX(i, 5)) <= ERROR_VAL And _
       (Abs(TEMP_MATRIX(i, 2) - TEMP_MATRIX(i, 5)) < _
       (TEMP_MATRIX(i, 3) - TEMP_MATRIX(i, 4)) / SHADOW_VAL) Then
        TEMP_MATRIX(i, 7) = 1
        TEMP_MATRIX(0, 7) = TEMP_MATRIX(0, 7) + TEMP_MATRIX(i, 7)
    Else
        TEMP_MATRIX(i, 7) = 0
    End If
    
    If (TEMP_MATRIX(i, 3) - TEMP_MATRIX(i, 2)) <= ERROR_VAL And _
       (Abs(TEMP_MATRIX(i, 2) - TEMP_MATRIX(i, 5)) < _
       (TEMP_MATRIX(i, 3) - TEMP_MATRIX(i, 4)) / SHADOW_VAL) Then
        TEMP_MATRIX(i, 8) = 1
        TEMP_MATRIX(0, 8) = TEMP_MATRIX(0, 8) + TEMP_MATRIX(i, 8)
    Else
        TEMP_MATRIX(i, 8) = 0
    End If
    
    If (TEMP_MATRIX(i, 2) - TEMP_MATRIX(i, 5)) <= ERROR_VAL Then
        TEMP_MATRIX(i, 9) = 1
        TEMP_MATRIX(0, 9) = TEMP_MATRIX(0, 9) + TEMP_MATRIX(i, 9)
    Else
        TEMP_MATRIX(i, 9) = 0
    End If
Next i

TEMP_MATRIX(0, 7) = "WHITES = " & TEMP_MATRIX(0, 7)
TEMP_MATRIX(0, 8) = "BLACKS = " & TEMP_MATRIX(0, 8)
TEMP_MATRIX(0, 9) = "DOJIS = " & TEMP_MATRIX(0, 9)

'-------------------------------------------------------------------------------------------

ASSET_CANDLESTICK_FUNC = TEMP_MATRIX

Exit Function
ERROR_LABEL:
ASSET_CANDLESTICK_FUNC = Err.number
End Function

'CandleSticks with Volume Weighted
'There 's also a slider called Amplify that'll rescale the
'candlesticks, increasing the size of those that are associated
'with larger volumes. The amount of rescaling is controlled by
'that slider.

'// PERFECT

Function ASSET_CANDLESTICK_VWEIGHTED_FUNC(ByVal TICKER_STR As Variant, _
Optional ByVal START_DATE As Date, _
Optional ByVal END_DATE As Date, _
Optional ByVal AMPLIFY As Double = 1.4)

Dim i As Long
Dim j As Long
Dim NROWS As Long
Dim NCOLUMNS As Long

Dim TEMP_SUM As Double
Dim TEMP_MATRIX As Variant
Dim DATA_MATRIX As Variant

On Error GoTo ERROR_LABEL

If IsArray(TICKER_STR) = False Then
    DATA_MATRIX = YAHOO_HISTORICAL_DATA_SERIE_FUNC(TICKER_STR, START_DATE, END_DATE, _
                  "d", "DOHLCV", False, True, True)
Else
    DATA_MATRIX = TICKER_STR
End If

NROWS = UBound(DATA_MATRIX, 1)
NCOLUMNS = UBound(DATA_MATRIX, 2)

ReDim TEMP_MATRIX(0 To NROWS, 1 To 10)
TEMP_MATRIX(0, 1) = "DATE"
TEMP_MATRIX(0, 2) = "OPEN"
TEMP_MATRIX(0, 3) = "HIGH"
TEMP_MATRIX(0, 4) = "LOW"
TEMP_MATRIX(0, 5) = "CLOSE"
TEMP_MATRIX(0, 6) = "VOLUME"
TEMP_MATRIX(0, 7) = "VW"
TEMP_MATRIX(0, 8) = "(H-(H-O))*VW"
TEMP_MATRIX(0, 9) = "(H-(H-L))*VW"
TEMP_MATRIX(0, 10) = "(H-(H-C))*VW"

For i = 1 To NROWS
    For j = 1 To NCOLUMNS: TEMP_MATRIX(i, j) = DATA_MATRIX(i, j): Next j
    TEMP_SUM = TEMP_SUM + TEMP_MATRIX(i, 6)
Next i

For i = 1 To NROWS
    TEMP_MATRIX(i, 7) = NROWS * TEMP_MATRIX(i, 6) / TEMP_SUM * AMPLIFY
    TEMP_MATRIX(i, 8) = TEMP_MATRIX(i, 3) - (TEMP_MATRIX(i, 3) - TEMP_MATRIX(i, 2)) * TEMP_MATRIX(i, 7)
    TEMP_MATRIX(i, 9) = TEMP_MATRIX(i, 3) - (TEMP_MATRIX(i, 3) - TEMP_MATRIX(i, 4)) * TEMP_MATRIX(i, 7)
    TEMP_MATRIX(i, 10) = TEMP_MATRIX(i, 3) - (TEMP_MATRIX(i, 3) - TEMP_MATRIX(i, 5)) * TEMP_MATRIX(i, 7)
Next i

ASSET_CANDLESTICK_VWEIGHTED_FUNC = TEMP_MATRIX

Exit Function
ERROR_LABEL:
ASSET_CANDLESTICK_VWEIGHTED_FUNC = Err.number
End Function


'// PERFECT

Function ASSET_CANDLESTICK_COMBOS_FUNC(ByVal TICKER_STR As Variant, _
Optional ByVal START_DATE As Date, _
Optional ByVal END_DATE As Date, _
Optional ByRef PARAM_RNG As Variant, _
Optional ByVal F_VAL As Double = 1.5, _
Optional ByVal G_VAL As Double = 0.5, _
Optional ByVal OUTPUT As Integer = 2)

'----------------------------------------------------------
'PARAM_RNG: (exclude headings)
'Red?=IF(OP>CL,1,0)    Body/Shad =ABS(OP-CL)/30
'----------------------------------------------------------
'       1                 0.633333333333333
'       1                 0.433333333333333
'       1                 0.600000000000000
'       0                 0.233333333333333
'----------------------------------------------------------

Dim h As Long
Dim i As Long
Dim j As Long
Dim k As Long
Dim l As Long
Dim m As Long

Dim NSIZE As Long
Dim NROWS As Long
Dim NCOLUMNS As Long

Dim TEMP_SUM As Double
Dim INDEX_ARR() As Variant
Dim PARAM_VECTOR As Variant
Dim TEMP_MATRIX As Variant
Dim DATA_MATRIX As Variant

On Error GoTo ERROR_LABEL

If IsArray(TICKER_STR) = False Then
    DATA_MATRIX = YAHOO_HISTORICAL_DATA_SERIE_FUNC(TICKER_STR, START_DATE, END_DATE, _
                  "d", "DOHLCVA", False, True, True)
Else
    DATA_MATRIX = TICKER_STR
End If
NROWS = UBound(DATA_MATRIX, 1)
NCOLUMNS = UBound(DATA_MATRIX, 2)
PARAM_VECTOR = PARAM_RNG
If UBound(PARAM_VECTOR, 1) = 1 Then
    PARAM_VECTOR = MATRIX_TRANSPOSE_FUNC(PARAM_VECTOR)
End If
NSIZE = UBound(PARAM_VECTOR, 1)

m = NCOLUMNS + 2 + NSIZE * 2 + 1
ReDim TEMP_MATRIX(0 To NROWS, 1 To m)
TEMP_MATRIX(0, 1) = "DATE"
TEMP_MATRIX(0, 2) = "OPEN"
TEMP_MATRIX(0, 3) = "HIGH"
TEMP_MATRIX(0, 4) = "LOW"
TEMP_MATRIX(0, 5) = "CLOSE"
TEMP_MATRIX(0, 6) = "VOLUME"
TEMP_MATRIX(0, 7) = "ADJ. CLOSE"
TEMP_MATRIX(0, 8) = "BODY/SHAD"
TEMP_MATRIX(0, 9) = "RED?"

For j = 1 To NSIZE
    m = NCOLUMNS + 2 + j
    TEMP_MATRIX(0, m) = "DAY " & CStr(j)
    m = m + NSIZE
    TEMP_MATRIX(0, m) = "DAY " & CStr(j)
Next j
m = NCOLUMNS + 2 + NSIZE * 2 + 1
TEMP_MATRIX(0, m) = "S(DAYS)"

h = 0
l = NSIZE - 1
For i = 1 To NROWS
    For j = 1 To NCOLUMNS: TEMP_MATRIX(i, j) = DATA_MATRIX(i, j): Next j
    TEMP_MATRIX(i, NCOLUMNS + 1) = Abs(DATA_MATRIX(i, 2) - DATA_MATRIX(i, 5)) / (DATA_MATRIX(i, 3) - DATA_MATRIX(i, 4))
    TEMP_MATRIX(i, NCOLUMNS + 2) = IIf(DATA_MATRIX(i, 2) > DATA_MATRIX(i, 5), 1, 0)
    '---------------------------------------------------------------------------------------
    If i > l Then
    '---------------------------------------------------------------------------------------
        For j = 1 To NSIZE
            k = NSIZE - j
            m = NCOLUMNS + 2 + j
            
            If (TEMP_MATRIX(i - k, NCOLUMNS + 1) <= F_VAL * PARAM_VECTOR(j, 2) And _
                     TEMP_MATRIX(i - k, NCOLUMNS + 1) >= G_VAL * PARAM_VECTOR(j, 2)) Then
                TEMP_MATRIX(i, m) = 1
            Else
                TEMP_MATRIX(i, m) = 0
            End If
            
            m = NCOLUMNS + 2 + NSIZE + j
            If TEMP_MATRIX(i - k, NCOLUMNS + 2) = PARAM_VECTOR(j, 1) Then
                TEMP_MATRIX(i, m) = 1
            Else
                TEMP_MATRIX(i, m) = 0
            End If
                
            m = NCOLUMNS + 2 + NSIZE * 2 + 1
            TEMP_MATRIX(i, m) = TEMP_MATRIX(i, m) + (TEMP_MATRIX(i, NCOLUMNS + 2 + j) + TEMP_MATRIX(i, NCOLUMNS + 2 + NSIZE + j))
        Next j
        m = NCOLUMNS + 2 + NSIZE * 2 + 1
        If TEMP_MATRIX(i, m) = (NSIZE * 2) Then
            TEMP_MATRIX(i, m) = 1
            TEMP_SUM = TEMP_SUM + TEMP_MATRIX(i, m)
            h = h + 1
            ReDim Preserve INDEX_ARR(1 To h)
            INDEX_ARR(h) = i
        Else
            TEMP_MATRIX(i, m) = 0
        End If
    '---------------------------------------------------------------------------------------
    End If
    '---------------------------------------------------------------------------------------
Next i

'-------------------------------------------------------------------
If OUTPUT = 0 Or h = 0 Then
'-------------------------------------------------------------------
    ASSET_CANDLESTICK_COMBOS_FUNC = TEMP_MATRIX
'-------------------------------------------------------------------
ElseIf OUTPUT > UBound(INDEX_ARR, 1) Then 'Pattern Matching Index
'-------------------------------------------------------------------
    ASSET_CANDLESTICK_COMBOS_FUNC = INDEX_ARR
'-------------------------------------------------------------------
Else '
'-------------------------------------------------------------------
    ReDim DATA_MATRIX(0 To NSIZE, 1 To 10)
    DATA_MATRIX(0, 1) = "DATE"
    DATA_MATRIX(0, 2) = "OPEN"
    DATA_MATRIX(0, 3) = "HIGH"
    DATA_MATRIX(0, 4) = "LOW"
    DATA_MATRIX(0, 5) = "CLOSE"
    DATA_MATRIX(0, 6) = "VOLUME"
    DATA_MATRIX(0, 7) = "ADJ. CLOSE"
    DATA_MATRIX(0, 8) = "BODY/SHADOW RATIOS"
    DATA_MATRIX(0, 9) = "G-FACTOR"
    DATA_MATRIX(0, 10) = "F-FACTOR"
    
    j = OUTPUT
    j = INDEX_ARR(j)
    For i = 1 To NSIZE
        h = j - NSIZE + i
        For m = 1 To NCOLUMNS: DATA_MATRIX(i, m) = TEMP_MATRIX(h, m): Next m
        DATA_MATRIX(i, 8) = Abs(DATA_MATRIX(i, 2) - DATA_MATRIX(i, 5)) / (DATA_MATRIX(i, 3) - DATA_MATRIX(i, 4)) 'Search Pattern

        DATA_MATRIX(i, 9) = G_VAL * PARAM_VECTOR(i, 2)
        DATA_MATRIX(i, 10) = F_VAL * PARAM_VECTOR(i, 2)
    Next i
    ASSET_CANDLESTICK_COMBOS_FUNC = DATA_MATRIX
'-------------------------------------------------------------------
End If
'-------------------------------------------------------------------
    
Exit Function
ERROR_LABEL:
ASSET_CANDLESTICK_COMBOS_FUNC = Err.number
End Function
