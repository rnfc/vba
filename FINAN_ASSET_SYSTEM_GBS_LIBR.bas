Attribute VB_Name = "FINAN_ASSET_SYSTEM_GBS_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.


'************************************************************************************
'************************************************************************************
'FUNCTION      : ASSET_GBS_SYSTEM_FUNC
'DESCRIPTION   :
'LIBRARY       : FINAN_ASSET
'GROUP         : SIGNAL
'ID            : 006
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA

'************************************************************************************
'************************************************************************************


Function ASSET_GBS_SYSTEM_FUNC(ByVal TICKER_STR As Variant, _
Optional ByVal START_DATE As Date, _
Optional ByVal END_DATE As Date, _
Optional ByVal INITIAL_CASH As Double = 100000, _
Optional ByVal INITIAL_PERCENTAGE_INVESTED As Double = 0.5, _
Optional ByVal ZOOM_PERCENTAGE As Double = 0.8, _
Optional ByVal ZOOM_PERIOD As Long = 1)

'If you pick ZOOM_PERCENTAGE = 80%, for example, it means you want to Buy
'and Sell when 80% of the candlesticks are red or green, over the
'past ten days;

Dim i As Long
Dim j As Long

Dim SROW As Long
Dim NROWS As Long

Dim DATA_MATRIX As Variant

Dim TEMP1_SUM As Double
Dim TEMP2_SUM As Double
Dim RETURN_VAL As Double

On Error GoTo ERROR_LABEL

If IsArray(TICKER_STR) = False Then
    DATA_MATRIX = YAHOO_HISTORICAL_DATA_SERIE_FUNC(TICKER_STR, START_DATE, END_DATE, _
                  "d", "DOHLCVA", False, True, True)
Else
    DATA_MATRIX = TICKER_STR
End If

SROW = ZOOM_PERIOD + 1
NROWS = UBound(DATA_MATRIX, 1)

ReDim TEMP_MATRIX(0 To NROWS, 1 To 16)

TEMP_MATRIX(0, 1) = "DATE"
TEMP_MATRIX(0, 2) = "OPEN"
TEMP_MATRIX(0, 3) = "HIGH"
TEMP_MATRIX(0, 4) = "LOW"
TEMP_MATRIX(0, 5) = "CLOSE"
TEMP_MATRIX(0, 6) = "VOLUME"
TEMP_MATRIX(0, 7) = "ADJ CLOSE"
TEMP_MATRIX(0, 8) = "C<O"
TEMP_MATRIX(0, 9) = "C>O"
TEMP_MATRIX(0, 10) = "GBS2(BUY SIGNAL)"
TEMP_MATRIX(0, 11) = "GBS2(SELL SIGNAL)"
TEMP_MATRIX(0, 12) = "INVESTMENT"
TEMP_MATRIX(0, 13) = "CASH"
TEMP_MATRIX(0, 14) = "TOTAL"
TEMP_MATRIX(0, 15) = "BUY SIGNAL"
TEMP_MATRIX(0, 16) = "SELL SIGNAL"

For i = 1 To SROW
    For j = 1 To 7: TEMP_MATRIX(i, j) = DATA_MATRIX(i, j): Next j
    GoSub CLOSE_OPEN_LINE
    For j = 10 To 16: TEMP_MATRIX(i, j) = "": Next j
Next i

TEMP_MATRIX(SROW, 12) = (INITIAL_CASH * INITIAL_PERCENTAGE_INVESTED)
TEMP_MATRIX(SROW, 13) = INITIAL_CASH - TEMP_MATRIX(SROW, 12)
TEMP_MATRIX(SROW, 14) = TEMP_MATRIX(SROW, 12) + TEMP_MATRIX(SROW, 13)

i = SROW + 1
For j = 1 To 7: TEMP_MATRIX(i, j) = DATA_MATRIX(i, j): Next j
GoSub CLOSE_OPEN_LINE
GoSub PERCENT_LINE
RETURN_VAL = TEMP_MATRIX(i, 7) / TEMP_MATRIX(i - 1, 7) - 1
If (TEMP_MATRIX(i, 10) > 0) And (TEMP_MATRIX(i - 1, 13) > 0) Then
    TEMP_MATRIX(i, 12) = TEMP_MATRIX(i - 1, 13)
Else
    If (TEMP_MATRIX(i, 11) > 0) Then
        TEMP_MATRIX(i, 12) = 0
    Else
        TEMP_MATRIX(i, 12) = TEMP_MATRIX(i - 1, 12) * (1 + RETURN_VAL)
    End If
End If

If ((TEMP_MATRIX(i - 1, 12) = 0) And (TEMP_MATRIX(i, 12) > 0)) Then
    TEMP_MATRIX(i, 13) = 0
Else
    TEMP_MATRIX(i, 13) = TEMP_MATRIX(i - 1, 13)
End If
TEMP_MATRIX(i, 14) = TEMP_MATRIX(i, 12) + TEMP_MATRIX(i, 13)
TEMP_MATRIX(i, 15) = 0
GoSub SIGNAL_LINE

For i = SROW + 2 To NROWS
    For j = 1 To 7: TEMP_MATRIX(i, j) = DATA_MATRIX(i, j): Next j
    GoSub CLOSE_OPEN_LINE
    GoSub PERCENT_LINE
    RETURN_VAL = TEMP_MATRIX(i, 7) / TEMP_MATRIX(i - 1, 7) - 1
        
    If ((TEMP_MATRIX(i, 10) > 0) And (TEMP_MATRIX(i - 1, 13) > 0)) Then
        TEMP_MATRIX(i, 12) = TEMP_MATRIX(i - 1, 12) * (1 + RETURN_VAL) + TEMP_MATRIX(i - 1, 13)
    Else
        If ((TEMP_MATRIX(i, 11) > 0) And (TEMP_MATRIX(i - 1, 12) > 0)) Then
            TEMP_MATRIX(i, 12) = 0
        Else
            TEMP_MATRIX(i, 12) = TEMP_MATRIX(i - 1, 12) * (1 + RETURN_VAL)
        End If
    End If
    
    If (TEMP_MATRIX(i, 11) > 0) Then
        TEMP_MATRIX(i, 13) = TEMP_MATRIX(i - 1, 12) * (1 + RETURN_VAL) + TEMP_MATRIX(i - 1, 13)
    Else
        If (TEMP_MATRIX(i, 10) > 0) Then
            TEMP_MATRIX(i, 13) = 0
        Else
            TEMP_MATRIX(i, 13) = TEMP_MATRIX(i - 1, 13)
        End If
    End If
    
    TEMP_MATRIX(i, 14) = TEMP_MATRIX(i, 12) + TEMP_MATRIX(i, 13)

    If (TEMP_MATRIX(i - 1, 12) = 0 And TEMP_MATRIX(i, 13) = 0 And TEMP_MATRIX(i - 2, 13) > 0) Then
        TEMP_MATRIX(i, 15) = TEMP_MATRIX(i, 14)
    Else
        TEMP_MATRIX(i, 15) = 0
    End If
    GoSub SIGNAL_LINE
Next i

ASSET_GBS_SYSTEM_FUNC = TEMP_MATRIX

Exit Function
'---------------------------------------------------------------------------------
CLOSE_OPEN_LINE:
'---------------------------------------------------------------------------------
    If TEMP_MATRIX(i, 5) < TEMP_MATRIX(i, 2) Then
        TEMP_MATRIX(i, 8) = 1
    Else
        TEMP_MATRIX(i, 8) = 0
    End If
    
    If TEMP_MATRIX(i, 5) > TEMP_MATRIX(i, 2) Then
        TEMP_MATRIX(i, 9) = 1
    Else
        TEMP_MATRIX(i, 9) = 0
    End If
'---------------------------------------------------------------------------------
Return
'---------------------------------------------------------------------------------
PERCENT_LINE:
'---------------------------------------------------------------------------------
    TEMP1_SUM = 0: TEMP2_SUM = 0
    For j = 0 To ZOOM_PERIOD - 1
        TEMP1_SUM = TEMP1_SUM + TEMP_MATRIX(i - j, 8)
        TEMP2_SUM = TEMP2_SUM + TEMP_MATRIX(i - j, 9)
    Next j
    
    If (TEMP1_SUM / ZOOM_PERIOD >= ZOOM_PERCENTAGE) Then
        TEMP_MATRIX(i, 10) = TEMP_MATRIX(i, 7)
    Else
        TEMP_MATRIX(i, 10) = 0
    End If
    
    If (TEMP2_SUM / ZOOM_PERIOD >= ZOOM_PERCENTAGE) Then
        TEMP_MATRIX(i, 11) = TEMP_MATRIX(i, 7)
    Else
        TEMP_MATRIX(i, 11) = 0
    End If
'---------------------------------------------------------------------------------
Return
'---------------------------------------------------------------------------------
SIGNAL_LINE:
'---------------------------------------------------------------------------------
    If (TEMP_MATRIX(i - 1, 12) > 0 And TEMP_MATRIX(i, 12) = 0) Then
        TEMP_MATRIX(i, 16) = TEMP_MATRIX(i, 14)
    Else
        TEMP_MATRIX(i, 16) = 0
    End If
'---------------------------------------------------------------------------------
Return
'---------------------------------------------------------------------------------
ERROR_LABEL:
ASSET_GBS_SYSTEM_FUNC = Err.number
End Function

