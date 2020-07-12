Attribute VB_Name = "FINAN_ASSET_TA_DMA_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.

'************************************************************************************
'************************************************************************************
'FUNCTION      : ASSET_TA_DMA_FUNC
'DESCRIPTION   : Displaced Moving Averages or DMA
'LIBRARY       : FINAN_ASSET
'GROUP         :
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA

'************************************************************************************
'************************************************************************************

Function ASSET_TA_DMA_FUNC(ByVal TICKER_STR As Variant, _
Optional ByVal START_DATE As Date, _
Optional ByVal END_DATE As Date, _
Optional ByVal MA_PERIOD As Long = 7, _
Optional ByVal NO_PERIODS As Long = 5)

'NO_PERIODS --> Days Advanced

Dim i As Long
Dim j As Long

Dim SROW As Long
Dim NROWS As Long

Dim TEMP_SUM As Double
Dim TEMP_MATRIX As Variant
Dim DATA_MATRIX As Variant

On Error GoTo ERROR_LABEL

If IsArray(TICKER_STR) = False Then
    DATA_MATRIX = YAHOO_HISTORICAL_DATA_SERIE_FUNC(TICKER_STR, START_DATE, END_DATE, _
                  "d", "DA", False, True, True)
Else
    DATA_MATRIX = TICKER_STR
End If

NROWS = UBound(DATA_MATRIX, 1)
SROW = (NO_PERIODS + 1) * MA_PERIOD

ReDim TEMP_MATRIX(0 To NROWS, 1 To 5)

TEMP_MATRIX(0, 1) = "DATE"
TEMP_MATRIX(0, 2) = "PRICE"
TEMP_MATRIX(0, 3) = "RETURN"
TEMP_MATRIX(0, 4) = MA_PERIOD & "_MA"
TEMP_MATRIX(0, 5) = MA_PERIOD & " X_VAL " & NO_PERIODS & " DMA"

TEMP_SUM = 0

i = 1
TEMP_MATRIX(i, 1) = DATA_MATRIX(i, 1)
TEMP_MATRIX(i, 2) = DATA_MATRIX(i, 2)
TEMP_SUM = TEMP_MATRIX(i, 2)
TEMP_MATRIX(i, 3) = 0
TEMP_MATRIX(i, 4) = TEMP_MATRIX(i, 2)
TEMP_MATRIX(i, 5) = TEMP_MATRIX(i, 4)

For i = 2 To NROWS
    TEMP_MATRIX(i, 1) = DATA_MATRIX(i, 1)
    TEMP_MATRIX(i, 2) = DATA_MATRIX(i, 2)
    TEMP_MATRIX(i, 3) = TEMP_MATRIX(i, 2) / TEMP_MATRIX(i - 1, 2) - 1
    If i > MA_PERIOD Then
        j = i - MA_PERIOD
        TEMP_SUM = TEMP_SUM + TEMP_MATRIX(i, 2)
        TEMP_SUM = TEMP_SUM - TEMP_MATRIX(j, 2)
        TEMP_MATRIX(i, 4) = TEMP_SUM / MA_PERIOD
    Else
        TEMP_SUM = TEMP_SUM + TEMP_MATRIX(i, 2)
        TEMP_MATRIX(i, 4) = TEMP_SUM / i
    End If
    
    If i > NO_PERIODS Then
        TEMP_MATRIX(i, 5) = TEMP_MATRIX(i - NO_PERIODS, 4)
    Else
        TEMP_MATRIX(i, 5) = TEMP_MATRIX(i, 4)
    End If
Next i

ASSET_TA_DMA_FUNC = TEMP_MATRIX

Exit Function
ERROR_LABEL:
ASSET_TA_DMA_FUNC = Err.number
End Function

