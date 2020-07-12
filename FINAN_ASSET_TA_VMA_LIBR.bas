Attribute VB_Name = "FINAN_ASSET_TA_VMA_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.


'************************************************************************************
'************************************************************************************
'FUNCTION      : ASSET_TA_VMA_FUNC
'DESCRIPTION   : Volume Moving Average
'LIBRARY       : FINAN_ASSET
'GROUP         :
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'************************************************************************************
'************************************************************************************

Function ASSET_TA_VMA_FUNC(ByVal TICKER_STR As Variant, _
Optional ByVal START_DATE As Date, _
Optional ByVal END_DATE As Date, _
Optional ByVal VMA_PERIOD As Long = 5)

Dim i As Long
Dim j As Long
Dim NROWS As Long
Dim NCOLUMNS As Long

Dim TEMP1_SUM As Double
Dim TEMP2_SUM As Double

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

ReDim TEMP_MATRIX(0 To NROWS, 1 To NCOLUMNS + 1)
TEMP_MATRIX(0, 1) = "DATE"
TEMP_MATRIX(0, 2) = "OPEN"
TEMP_MATRIX(0, 3) = "HIGH"
TEMP_MATRIX(0, 4) = "LOW"
TEMP_MATRIX(0, 5) = "CLOSE"
TEMP_MATRIX(0, 6) = "VOLUME"
TEMP_MATRIX(0, 7) = "CLOSE VW"

TEMP1_SUM = 0: TEMP2_SUM = 0
For i = 1 To NROWS
    For j = 1 To NCOLUMNS: TEMP_MATRIX(i, j) = DATA_MATRIX(i, j): Next j
    If i > VMA_PERIOD Then
        j = i - VMA_PERIOD
        TEMP1_SUM = TEMP1_SUM - (TEMP_MATRIX(j, 5) * TEMP_MATRIX(j, 6))
        TEMP2_SUM = TEMP2_SUM - TEMP_MATRIX(j, 6)
    End If
    TEMP1_SUM = TEMP1_SUM + (TEMP_MATRIX(i, 5) * TEMP_MATRIX(i, 6))
    TEMP2_SUM = TEMP2_SUM + TEMP_MATRIX(i, 6)
    TEMP_MATRIX(i, 7) = TEMP1_SUM / TEMP2_SUM
Next i

ASSET_TA_VMA_FUNC = TEMP_MATRIX

Exit Function
ERROR_LABEL:
ASSET_TA_VMA_FUNC = Err.number
End Function
