Attribute VB_Name = "FINAN_ASSET_TA_MACD_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.

Function ASSET_TA_MACD_BOLLINGER_FUNC(ByVal TICKER_STR As Variant, _
Optional ByVal START_DATE As Date, _
Optional ByVal END_DATE As Date, _
Optional ByVal MA_PERIOD As Long = 15, _
Optional ByVal SD_FACTOR As Double = 2, _
Optional ByVal EMA1_PERIOD As Long = 20, _
Optional ByVal EMA2_PERIOD As Long = 50, _
Optional ByVal VOLUME_WEIGHTED_FLAG As Boolean = True)

Dim i As Long
Dim j As Long

Dim NROWS As Long
Dim NCOLUMNS As Long

Dim ALPHA1_VAL As Double
Dim ALPHA2_VAL As Double
Dim VOLUME_VAL As Double

Dim TEMP1_SUM As Double
Dim TEMP2_SUM As Double

Dim MEAN_VAL As Double
Dim VOLATILITY_VAL As Double

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

ALPHA1_VAL = 1 - 2 / (EMA1_PERIOD + 1)
ALPHA2_VAL = 1 - 2 / (EMA2_PERIOD + 1)
'-----------------------------------------------------------------------------
ReDim TEMP_MATRIX(0 To NROWS, 1 To 17)
'-----------------------------------------------------------------------------
TEMP_MATRIX(0, 1) = "DATE"
TEMP_MATRIX(0, 2) = "OPEN"
TEMP_MATRIX(0, 3) = "HIGH"
TEMP_MATRIX(0, 4) = "LOW"
TEMP_MATRIX(0, 5) = "CLOSE"
TEMP_MATRIX(0, 6) = "VOLUME"

TEMP_MATRIX(0, 7) = "NUM_" & EMA1_PERIOD
TEMP_MATRIX(0, 8) = "DEN_" & EMA1_PERIOD
TEMP_MATRIX(0, 9) = "EMA_" & EMA1_PERIOD

TEMP_MATRIX(0, 10) = "NUM_" & EMA2_PERIOD
TEMP_MATRIX(0, 11) = "DEN_" & EMA2_PERIOD
TEMP_MATRIX(0, 12) = "EMA_" & EMA2_PERIOD

TEMP_MATRIX(0, 13) = "MACD"
TEMP_MATRIX(0, 14) = "MEAN"
TEMP_MATRIX(0, 15) = "SD"
TEMP_MATRIX(0, 16) = "MEAN - " & Format(SD_FACTOR, "0.0") & " x SD"
TEMP_MATRIX(0, 17) = "MEAN + " & Format(SD_FACTOR, "0.0") & " x SD"

i = 1
For j = 1 To NCOLUMNS: TEMP_MATRIX(i, j) = DATA_MATRIX(i, j): Next j
For j = 7 To 17: TEMP_MATRIX(i, j) = "": Next j
TEMP1_SUM = TEMP_MATRIX(i, 5)
TEMP2_SUM = 0

i = 2
For j = 1 To NCOLUMNS: TEMP_MATRIX(i, j) = DATA_MATRIX(i, j): Next j
If VOLUME_WEIGHTED_FLAG = True Then
    TEMP_MATRIX(i, 8) = (1 - ALPHA1_VAL) * TEMP_MATRIX(i, 6) / 1000
    TEMP_MATRIX(i, 11) = (1 - ALPHA2_VAL) * TEMP_MATRIX(i, 6) / 1000
    VOLUME_VAL = TEMP_MATRIX(i, 6) / 1000
Else
    TEMP_MATRIX(i, 8) = 1
    TEMP_MATRIX(i, 11) = 1
    VOLUME_VAL = 1
End If

TEMP_MATRIX(i, 7) = (1 - ALPHA1_VAL) * TEMP_MATRIX(i, 5) * VOLUME_VAL

TEMP_MATRIX(i, 9) = TEMP_MATRIX(i, 7) / TEMP_MATRIX(i, 8)
TEMP_MATRIX(i, 10) = (1 - ALPHA2_VAL) * TEMP_MATRIX(i, 5) * VOLUME_VAL
TEMP_MATRIX(i, 12) = TEMP_MATRIX(i, 10) / TEMP_MATRIX(i, 11)
TEMP_MATRIX(i, 13) = TEMP_MATRIX(i, 9) - TEMP_MATRIX(i, 12)

TEMP_MATRIX(i, 14) = ""
TEMP_MATRIX(i, 15) = ""
TEMP_MATRIX(i, 16) = ""
TEMP_MATRIX(i, 17) = ""
TEMP1_SUM = TEMP1_SUM + TEMP_MATRIX(i, 5)
TEMP2_SUM = 0

For i = 3 To NROWS
    For j = 1 To NCOLUMNS: TEMP_MATRIX(i, j) = DATA_MATRIX(i, j): Next j

    If VOLUME_WEIGHTED_FLAG = True Then
        TEMP_MATRIX(i, 8) = TEMP_MATRIX(i - 1, 8) * ALPHA1_VAL + (1 - ALPHA1_VAL) * TEMP_MATRIX(i, 6) / 1000
        TEMP_MATRIX(i, 11) = TEMP_MATRIX(i - 1, 11) * ALPHA2_VAL + (1 - ALPHA2_VAL) * TEMP_MATRIX(i, 6) / 1000
        VOLUME_VAL = TEMP_MATRIX(i - 1, 6) / 1000
    Else
        TEMP_MATRIX(i, 8) = 1
        TEMP_MATRIX(i, 11) = 1
        VOLUME_VAL = 1
    End If
    
    TEMP_MATRIX(i, 7) = TEMP_MATRIX(i - 1, 7) * ALPHA1_VAL + (1 - ALPHA1_VAL) * _
                        TEMP_MATRIX(i, 5) * VOLUME_VAL
    TEMP_MATRIX(i, 9) = TEMP_MATRIX(i, 7) / TEMP_MATRIX(i, 8)
    
    TEMP_MATRIX(i, 10) = TEMP_MATRIX(i - 1, 10) * ALPHA2_VAL + (1 - ALPHA2_VAL) * _
                         TEMP_MATRIX(i, 5) * VOLUME_VAL
    TEMP_MATRIX(i, 12) = TEMP_MATRIX(i, 10) / TEMP_MATRIX(i, 11)
    TEMP_MATRIX(i, 13) = TEMP_MATRIX(i, 9) - TEMP_MATRIX(i, 12)
    
    TEMP1_SUM = TEMP1_SUM + TEMP_MATRIX(i, 5)
    If i >= MA_PERIOD Then
        MEAN_VAL = TEMP1_SUM / MA_PERIOD
        TEMP_MATRIX(i, 14) = MEAN_VAL
        TEMP1_SUM = TEMP1_SUM - TEMP_MATRIX(i - MA_PERIOD + 1, 5)
                
        TEMP2_SUM = 0
        For j = i To (i - MA_PERIOD + 1) Step -1
            TEMP2_SUM = TEMP2_SUM + (TEMP_MATRIX(j, 5) - MEAN_VAL) ^ 2
        Next j
        VOLATILITY_VAL = (TEMP2_SUM / MA_PERIOD) ^ 0.5
        
        TEMP_MATRIX(i, 15) = VOLATILITY_VAL
        TEMP_MATRIX(i, 16) = TEMP_MATRIX(i, 14) - SD_FACTOR * TEMP_MATRIX(i, 15)
        TEMP_MATRIX(i, 17) = TEMP_MATRIX(i, 14) + SD_FACTOR * TEMP_MATRIX(i, 15)
    Else
        TEMP_MATRIX(i, 14) = ""
        TEMP_MATRIX(i, 15) = ""
        TEMP_MATRIX(i, 16) = ""
        TEMP_MATRIX(i, 17) = ""
    End If
Next i

ASSET_TA_MACD_BOLLINGER_FUNC = TEMP_MATRIX

Exit Function
ERROR_LABEL:
ASSET_TA_MACD_BOLLINGER_FUNC = Err.number
End Function
