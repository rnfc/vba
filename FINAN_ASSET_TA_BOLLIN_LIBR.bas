Attribute VB_Name = "FINAN_ASSET_TA_BOLLIN_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.

'************************************************************************************
'************************************************************************************
'FUNCTION      : ASSET_TA_NBOLLINGER_FUNC

'DESCRIPTION   : Normalized Bollinger Bands: How's the current stock price doing
'compared to the 20-day average? And, what's the upper band value compared to
'the 20-day average.

'LIBRARY       : FINAN_ASSET
'GROUP         : TA
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA

'************************************************************************************
'************************************************************************************

Function ASSET_TA_NBOLLINGER_FUNC(ByVal TICKER_STR As Variant, _
Optional ByVal START_DATE As Date, _
Optional ByVal END_DATE As Date, _
Optional ByVal MA_PERIOD As Long = 20, _
Optional ByVal SD_FACTOR As Double = 2)

Dim i As Long
Dim j As Long

Dim NROWS As Long

Dim TEMP_SUM As Double
Dim MEAN_VAL As Double
Dim SIGMA_VAL As Double
Dim ZTEMP_VAL As Double

Dim TEMP_MATRIX As Variant
Dim DATA_MATRIX As Variant

On Error GoTo ERROR_LABEL

If IsArray(TICKER_STR) = False Then
    DATA_MATRIX = YAHOO_HISTORICAL_DATA_SERIE_FUNC(TICKER_STR, _
                  START_DATE, END_DATE, "DAILY", "DOHLCVA", False, _
                  True, True)
Else
    DATA_MATRIX = TICKER_STR
End If
NROWS = UBound(DATA_MATRIX, 1)

ReDim TEMP_MATRIX(0 To NROWS, 1 To 14)

'------------------------------------------------------------------------------
TEMP_MATRIX(0, 1) = "DATE"
TEMP_MATRIX(0, 2) = "OPEN"
TEMP_MATRIX(0, 3) = "HIGH"
TEMP_MATRIX(0, 4) = "LOW"
TEMP_MATRIX(0, 5) = "CLOSE"
TEMP_MATRIX(0, 6) = "VOLUME"
TEMP_MATRIX(0, 7) = "ADJ CLOSE"
TEMP_MATRIX(0, 8) = "RETURN" '8
TEMP_MATRIX(0, 9) = "PAVG PRICE" '9
TEMP_MATRIX(0, 10) = "LOW RETURN" '10
TEMP_MATRIX(0, 11) = "HIGH RETURN" '11
TEMP_MATRIX(0, 12) = "P/PAVG RETURN" '12
TEMP_MATRIX(0, 13) = "BOLLI-LOW" '13
TEMP_MATRIX(0, 14) = "BOLLI-HIGH" '14

'-------------------------------------------------------------------------------
'-------------------------------------------------------------------------------
i = 1
For j = 1 To 7: TEMP_MATRIX(i, j) = DATA_MATRIX(i, j): Next j
TEMP_MATRIX(i, 6) = TEMP_MATRIX(i, 6) / 1000
For j = 8 To 14: TEMP_MATRIX(i, j) = 0: Next j

TEMP_SUM = TEMP_MATRIX(i, 7)
For i = 2 To NROWS
    For j = 1 To 7: TEMP_MATRIX(i, j) = DATA_MATRIX(i, j): Next j
    TEMP_MATRIX(i, 6) = TEMP_MATRIX(i, 6) / 1000
    TEMP_MATRIX(i, 8) = TEMP_MATRIX(i, 7) / TEMP_MATRIX(i - 1, 7) - 1
    
    TEMP_SUM = TEMP_SUM + TEMP_MATRIX(i, 7)
    If i < MA_PERIOD Then
        MEAN_VAL = TEMP_SUM / i
        SIGMA_VAL = 0
        For j = i To 1 Step -1
            SIGMA_VAL = SIGMA_VAL + (TEMP_MATRIX(j, 7) - MEAN_VAL) ^ 2
        Next j
        SIGMA_VAL = (SIGMA_VAL / i) ^ 0.5
    Else
        MEAN_VAL = TEMP_SUM / (MA_PERIOD)
        TEMP_SUM = TEMP_SUM - TEMP_MATRIX(i - MA_PERIOD + 1, 7)
        
        SIGMA_VAL = 0
        For j = i To (i - MA_PERIOD + 1) Step -1
            SIGMA_VAL = SIGMA_VAL + (TEMP_MATRIX(j, 7) - MEAN_VAL) ^ 2
        Next j
        SIGMA_VAL = (SIGMA_VAL / MA_PERIOD) ^ 0.5
    End If
    
    TEMP_MATRIX(i, 9) = MEAN_VAL
    ZTEMP_VAL = SD_FACTOR * SIGMA_VAL
    TEMP_MATRIX(i, 10) = ZTEMP_VAL / MEAN_VAL * -1
    TEMP_MATRIX(i, 11) = ZTEMP_VAL / MEAN_VAL
    TEMP_MATRIX(i, 12) = TEMP_MATRIX(i, 7) / TEMP_MATRIX(i, 9) - 1
    TEMP_MATRIX(i, 13) = MEAN_VAL - ZTEMP_VAL
    TEMP_MATRIX(i, 14) = MEAN_VAL + ZTEMP_VAL
Next i

ASSET_TA_NBOLLINGER_FUNC = TEMP_MATRIX

Exit Function
ERROR_LABEL:
ASSET_TA_NBOLLINGER_FUNC = Err.number
End Function
