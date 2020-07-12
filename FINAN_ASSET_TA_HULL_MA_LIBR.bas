Attribute VB_Name = "FINAN_ASSET_TA_HULL_MA_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.


Function ASSET_HULL_MA_SYSTEM_FUNC(ByVal TICKER_STR As Variant, _
Optional ByVal START_DATE As Date, _
Optional ByVal END_DATE As Date, _
Optional ByVal DOWN_FACTOR As Double = -0.01, _
Optional ByVal UP_FACTOR As Double = 0.015, _
Optional ByVal MA_PERIOD As Long = 36, _
Optional ByVal MAg_FACTOR As Double = 0.6)

'Reference: http://www.alanhull.com.au/hma/hma.html
'IF Version = 0 Then: MAg, HMA

Dim g As Long
Dim h As Long
Dim i As Long
Dim j As Long
Dim k As Long
Dim l As Long

Dim NROWS As Long

Dim A_VAL As Double
Dim B_VAL As Double
Dim C_VAL As Double
Dim D_VAL As Double
Dim E_VAL As Double

Dim TEMP_VAL As Double
Dim TEMP_SUM As Double

Dim TEMP_MATRIX As Variant
Dim DATA_MATRIX As Variant

On Error GoTo ERROR_LABEL

If IsArray(TICKER_STR) = False Then
    DATA_MATRIX = YAHOO_HISTORICAL_DATA_SERIE_FUNC(TICKER_STR, START_DATE, END_DATE, "d", "DA", False, True, True)
Else
    DATA_MATRIX = TICKER_STR
End If
NROWS = UBound(DATA_MATRIX, 1)

g = Int(MA_PERIOD ^ 0.5)
h = Int(MA_PERIOD / 2)

ReDim TEMP_MATRIX(0 To NROWS, 1 To 15)

TEMP_MATRIX(0, 1) = "DATE"
TEMP_MATRIX(0, 2) = "PRICE"
TEMP_MATRIX(0, 3) = "SMA_" & MA_PERIOD

TEMP_MATRIX(0, 4) = "WMA_" & MA_PERIOD
TEMP_MATRIX(0, 5) = "WMA_" & h

TEMP_MATRIX(0, 6) = "EMA_" & MA_PERIOD
TEMP_MATRIX(0, 7) = "MMA_" & MA_PERIOD
TEMP_MATRIX(0, 8) = "HMA_" & MA_PERIOD

TEMP_MATRIX(0, 9) = "WILDER_MA"
TEMP_MATRIX(0, 10) = "MAg(" & MA_PERIOD & "," & MAg_FACTOR & ")"
TEMP_MATRIX(0, 11) = "EMA_" & h

TEMP_MATRIX(0, 12) = "MAg_BUY"
TEMP_MATRIX(0, 13) = "MAg_SELL"

TEMP_MATRIX(0, 14) = "HMA_BUY"
TEMP_MATRIX(0, 15) = "HMA_SELL"

A_VAL = 1 - 2 / (MA_PERIOD + 1)
B_VAL = MA_PERIOD * (MA_PERIOD + 1) / 2
C_VAL = h * (h + 1) / 2
D_VAL = 1 - 2 / (h + 1)
E_VAL = g * (g + 1) / 2

TEMP_MATRIX(1, 1) = DATA_MATRIX(1, 1)
TEMP_MATRIX(1, 2) = DATA_MATRIX(1, 2)

For j = 3 To 11: TEMP_MATRIX(1, j) = TEMP_MATRIX(1, 2): Next j

TEMP_MATRIX(1, 12) = 0
TEMP_MATRIX(1, 13) = 0
TEMP_MATRIX(1, 14) = 0
TEMP_MATRIX(1, 15) = 0

'----------------------------------------------------------------------------
i = 1
TEMP_SUM = TEMP_MATRIX(i, 2)
'----------------------------------------------------------------------------
For i = 2 To NROWS
'----------------------------------------------------------------------------
    TEMP_MATRIX(i, 1) = DATA_MATRIX(i, 1)
    TEMP_MATRIX(i, 2) = DATA_MATRIX(i, 2)
'----------------------------------------------------------------------------
    TEMP_SUM = TEMP_SUM + TEMP_MATRIX(i, 2)
    If i <= MA_PERIOD Then
        TEMP_MATRIX(i, 3) = TEMP_SUM / i
    Else
        TEMP_MATRIX(i, 3) = TEMP_SUM / (MA_PERIOD + 1)
        k = i - MA_PERIOD
        TEMP_SUM = TEMP_SUM - TEMP_MATRIX(k, 2)
    End If
'----------------------------------------------------------------------------
    k = IIf(i <= MA_PERIOD, i - 1, MA_PERIOD - 1)
    l = 0
    For j = (i - k) To i
        l = l + 1
        TEMP_MATRIX(i, 4) = TEMP_MATRIX(i, 4) + (l * TEMP_MATRIX(j, 2))
    Next j
    TEMP_MATRIX(i, 4) = TEMP_MATRIX(i, 4) / B_VAL
'----------------------------------------------------------------------------
    
    k = IIf(i <= h, i - 1, h - 1)
    l = 0
    For j = (i - k) To i
        l = l + 1
        TEMP_MATRIX(i, 5) = TEMP_MATRIX(i, 5) + (l * TEMP_MATRIX(j, 2))
    Next j
    TEMP_MATRIX(i, 5) = TEMP_MATRIX(i, 5) / C_VAL
    
    TEMP_MATRIX(i, 6) = A_VAL * TEMP_MATRIX(i - 1, 6) + (1 - A_VAL) * TEMP_MATRIX(i, 2)
    TEMP_MATRIX(i, 7) = 2 * TEMP_MATRIX(i, 5) - TEMP_MATRIX(i, 4)
    
'----------------------------------------------------------------------------
    k = IIf(i <= g, i - 1, g - 1)
    l = 0
    For j = (i - k) To i
        l = l + 1
        TEMP_MATRIX(i, 8) = TEMP_MATRIX(i, 8) + (l * TEMP_MATRIX(j, 7))
    Next j
    TEMP_MATRIX(i, 8) = TEMP_MATRIX(i, 8) / E_VAL
'----------------------------------------------------------------------------
    
    TEMP_MATRIX(i, 9) = (TEMP_MATRIX(i, 2) + (MA_PERIOD - 1) * TEMP_MATRIX(i - 1, 9)) / MA_PERIOD
    TEMP_MATRIX(i, 11) = D_VAL * TEMP_MATRIX(i - 1, 11) + (1 - D_VAL) * TEMP_MATRIX(i, 2)
    TEMP_MATRIX(i, 10) = (1 + MAg_FACTOR) * TEMP_MATRIX(i, 11) - MAg_FACTOR * TEMP_MATRIX(i, 6)
    
    If i > 3 Then
        TEMP_VAL = MAXIMUM_FUNC(TEMP_MATRIX(i - 2, 10), TEMP_MATRIX(i - 1, 10))
        If TEMP_MATRIX(i, 10) < (1 + DOWN_FACTOR) * TEMP_VAL Then
            TEMP_MATRIX(i, 12) = TEMP_MATRIX(i, 2)
        Else
            TEMP_MATRIX(i, 12) = 0
        End If
        
        TEMP_VAL = MINIMUM_FUNC(TEMP_MATRIX(i - 2, 10), TEMP_MATRIX(i - 1, 10))
        If TEMP_MATRIX(i, 10) > (1 + UP_FACTOR) * TEMP_VAL Then
            TEMP_MATRIX(i, 13) = TEMP_MATRIX(i, 2)
        Else
            TEMP_MATRIX(i, 13) = 0
        End If
        
        TEMP_VAL = MAXIMUM_FUNC(TEMP_MATRIX(i - 2, 8), TEMP_MATRIX(i - 1, 8))
        If TEMP_MATRIX(i, 10) < (1 + DOWN_FACTOR) * TEMP_VAL Then
            TEMP_MATRIX(i, 14) = TEMP_MATRIX(i, 2)
        Else
            TEMP_MATRIX(i, 14) = 0
        End If
        
        TEMP_VAL = MINIMUM_FUNC(TEMP_MATRIX(i - 2, 8), TEMP_MATRIX(i - 1, 8))
        If TEMP_MATRIX(i, 10) > (1 + UP_FACTOR) * TEMP_VAL Then
            TEMP_MATRIX(i, 15) = TEMP_MATRIX(i, 2)
        Else
            TEMP_MATRIX(i, 15) = 0
        End If
    Else
        TEMP_MATRIX(i, 12) = 0
        TEMP_MATRIX(i, 13) = 0
        TEMP_MATRIX(i, 14) = 0
        TEMP_MATRIX(i, 15) = 0
    End If
'----------------------------------------------------------------------------
Next i
'----------------------------------------------------------------------------

ASSET_HULL_MA_SYSTEM_FUNC = TEMP_MATRIX

Exit Function
ERROR_LABEL:
ASSET_HULL_MA_SYSTEM_FUNC = Err.number
End Function


