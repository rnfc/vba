Attribute VB_Name = "FINAN_ASSET_TA_RSI_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.


'Relative Strength Index (RSI)

Function ASSET_TA_RSI_TABLE_FUNC(ByVal TICKER_STR As Variant, _
Optional ByVal START_DATE As Date, _
Optional ByVal END_DATE As Date, _
Optional ByVal MA_PERIOD As Long = 5)

Dim i As Long
Dim j As Long
Dim NROWS As Long

Dim TEMP1_SUM As Double
Dim TEMP2_SUM As Double

Dim TEMP_MATRIX As Variant
Dim DATA_MATRIX As Variant

On Error GoTo ERROR_LABEL

If IsArray(TICKER_STR) = False Then
    DATA_MATRIX = YAHOO_HISTORICAL_DATA_SERIE_FUNC(TICKER_STR, START_DATE, END_DATE, "d", "DA", False, True, True)
Else
    DATA_MATRIX = TICKER_STR
End If
NROWS = UBound(DATA_MATRIX, 1)

'--------------------------------------------------------------------------------
ReDim TEMP_MATRIX(0 To NROWS, 1 To 7)
'--------------------------------------------------------------------------------

TEMP_MATRIX(0, 1) = "DATE"
TEMP_MATRIX(0, 2) = "CLOSE"
TEMP_MATRIX(0, 3) = "UP BARS"
TEMP_MATRIX(0, 4) = "DOWN BARS"
TEMP_MATRIX(0, 5) = "AVG UP BARS"
TEMP_MATRIX(0, 6) = "AVG DOWN BARS"
TEMP_MATRIX(0, 7) = "RSI"

i = 1
TEMP_MATRIX(i, 1) = DATA_MATRIX(i, 1)
TEMP_MATRIX(i, 2) = DATA_MATRIX(i, 2)

TEMP_MATRIX(i, 3) = 0: TEMP_MATRIX(i, 4) = 0
TEMP_MATRIX(i, 5) = 0: TEMP_MATRIX(i, 6) = 0
TEMP_MATRIX(i, 7) = 0

TEMP1_SUM = 0: TEMP2_SUM = 0
For i = 2 To NROWS
    TEMP_MATRIX(i, 1) = DATA_MATRIX(i, 1)
    TEMP_MATRIX(i, 2) = DATA_MATRIX(i, 2)
    
    If TEMP_MATRIX(i, 2) > TEMP_MATRIX(i - 1, 2) Then
        TEMP_MATRIX(i, 3) = TEMP_MATRIX(i, 2) - TEMP_MATRIX(i - 1, 2)
    Else
        TEMP_MATRIX(i, 3) = 0
    End If
    
    If TEMP_MATRIX(i, 2) < TEMP_MATRIX(i - 1, 2) Then
        TEMP_MATRIX(i, 4) = TEMP_MATRIX(i, 2) - TEMP_MATRIX(i - 1, 2)
    Else
        TEMP_MATRIX(i, 4) = 0
    End If
    
    TEMP1_SUM = TEMP1_SUM + TEMP_MATRIX(i, 3)
    TEMP2_SUM = TEMP2_SUM + TEMP_MATRIX(i, 4)
    
    If i > MA_PERIOD Then
        j = i - MA_PERIOD
        TEMP1_SUM = TEMP1_SUM - TEMP_MATRIX(j, 3)
        TEMP2_SUM = TEMP2_SUM - TEMP_MATRIX(j, 4)
        
        TEMP_MATRIX(i, 5) = TEMP_MATRIX(i, 5) + TEMP1_SUM
        TEMP_MATRIX(i, 6) = TEMP_MATRIX(i, 6) + TEMP2_SUM
        
        If TEMP_MATRIX(i, 6) = 0 Then
            TEMP_MATRIX(i, 7) = 100
        Else
            TEMP_MATRIX(i, 7) = 100 - (100 / (1 + Abs(TEMP_MATRIX(i, 5) / TEMP_MATRIX(i, 6))))
        End If
    End If
Next i

ASSET_TA_RSI_TABLE_FUNC = TEMP_MATRIX

Exit Function
ERROR_LABEL:
ASSET_TA_RSI_TABLE_FUNC = Err.number
End Function
