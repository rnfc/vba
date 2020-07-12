Attribute VB_Name = "FINAN_ASSET_SYSTEM_LOW_MA_LIBR"
'----------------------------------------------------------------------------------------------------
'----------------------------------------------------------------------------------------------------
Option Explicit
Option Base 1
'----------------------------------------------------------------------------------------------------
'----------------------------------------------------------------------------------------------------

Function ASSET_LOW_MA_SYSTEM_FUNC(ByVal TICKER_STR As Variant, _
Optional ByVal START_DATE As Date, _
Optional ByVal END_DATE As Date, _
Optional ByVal MA_PERIOD As Long = 10, _
Optional ByVal CASH_RATE As Double = 0.04, _
Optional ByVal COUNT_BASIS As Double = 252, _
Optional ByVal OUTPUT As Integer = 0)

Dim i As Long
Dim j As Long
Dim k As Long 'Max drawdown duration
Dim NROWS As Long

Dim TEMP1_SUM As Double
Dim TEMP2_SUM As Double

Dim LOW_VAL As Double
Dim LMEAN_VAL As Double
Dim LVOLAT_VAL As Double
Dim LSHARPE_VAL As Double
Dim MDRAWDOWN_VAL As Double

Dim TEMP_MATRIX As Variant
Dim DATA_MATRIX As Variant

On Error GoTo ERROR_LABEL

'----------------------------------------------------------------------------------------------------
If IsArray(TICKER_STR) = False Then
    DATA_MATRIX = YAHOO_HISTORICAL_DATA_SERIE_FUNC(TICKER_STR, START_DATE, END_DATE, "d", "DOHLCV", False, True, True)
Else
    DATA_MATRIX = TICKER_STR
End If
'----------------------------------------------------------------------------------------------------
NROWS = UBound(DATA_MATRIX, 1)
'----------------------------------------------------------------------------------------------------
ReDim TEMP_MATRIX(0 To NROWS, 1 To 18)
'----------------------------------------------------------------------------------------------------
TEMP_MATRIX(0, 1) = "DATE"
TEMP_MATRIX(0, 2) = "OPEN"
TEMP_MATRIX(0, 3) = "HIGH"
TEMP_MATRIX(0, 4) = "LOW"
TEMP_MATRIX(0, 5) = "CLOSE"
TEMP_MATRIX(0, 6) = "VOLUME"
'----------------------------------------------------------------------------------------------------
TEMP_MATRIX(0, 7) = MA_PERIOD & "-DAY LOW"
TEMP_MATRIX(0, 8) = "ENTRY"
TEMP_MATRIX(0, 9) = MA_PERIOD & "-DAY AVERAGE"
TEMP_MATRIX(0, 10) = "EXIT"
TEMP_MATRIX(0, 11) = "POSITION"
'----------------------------------------------------------------------------------------------------
TEMP_MATRIX(0, 12) = "ENTRY PRICE"
TEMP_MATRIX(0, 13) = "DAILY RETURN"
TEMP_MATRIX(0, 14) = "EXCESS DAILY RETURN"
'----------------------------------------------------------------------------------------------------
TEMP_MATRIX(0, 15) = "CUMUL. RETURN"
TEMP_MATRIX(0, 16) = "HIGH WATERMARK"
TEMP_MATRIX(0, 17) = "DRAWDOWN"
TEMP_MATRIX(0, 18) = "MAX DRAWDOWN DURATION"
'----------------------------------------------------------------------------------------------------
LOW_VAL = 2 ^ 52
TEMP1_SUM = 0
'----------------------------------------------------------------------------------------------------
For i = 1 To MA_PERIOD
'----------------------------------------------------------------------------------------------------
    For j = 1 To 6: TEMP_MATRIX(i, j) = DATA_MATRIX(i, j): Next j
    TEMP_MATRIX(i, 6) = TEMP_MATRIX(i, 6) / 1000
    If TEMP_MATRIX(i, 4) < LOW_VAL Then: LOW_VAL = TEMP_MATRIX(i, 4)
    For j = 7 To 18: TEMP_MATRIX(i, j) = "": Next j
    TEMP1_SUM = TEMP1_SUM + TEMP_MATRIX(i, 5)
'----------------------------------------------------------------------------------------------------
Next i
'----------------------------------------------------------------------------------------------------
i = MA_PERIOD
TEMP_MATRIX(i, 7) = LOW_VAL
TEMP_MATRIX(i, 11) = 0
'----------------------------------------------------------------------------------------------------
For i = MA_PERIOD + 1 To NROWS
'----------------------------------------------------------------------------------------------------
    For j = 1 To 6: TEMP_MATRIX(i, j) = DATA_MATRIX(i, j): Next j
    TEMP_MATRIX(i, 6) = TEMP_MATRIX(i, 6) / 1000
    
    LOW_VAL = 2 ^ 52
    For j = i To i - MA_PERIOD + 1 Step -1
        If TEMP_MATRIX(j, 4) < LOW_VAL Then: LOW_VAL = TEMP_MATRIX(j, 4)
    Next j
    TEMP_MATRIX(i, 7) = LOW_VAL
    TEMP_MATRIX(i, 8) = IIf(TEMP_MATRIX(i, 4) <= TEMP_MATRIX(i - 1, 7), 1, 0)
    
    j = i - MA_PERIOD
    TEMP1_SUM = TEMP1_SUM - TEMP_MATRIX(j, 5)
    TEMP1_SUM = TEMP1_SUM + TEMP_MATRIX(i, 5)
    TEMP_MATRIX(i, 9) = TEMP1_SUM / MA_PERIOD
    
    TEMP_MATRIX(i, 10) = IIf(TEMP_MATRIX(i, 5) >= TEMP_MATRIX(i, 9), 1, 0)
    If TEMP_MATRIX(i, 8) = 1 Or (TEMP_MATRIX(i - 1, 11) = 1 And TEMP_MATRIX(i, 10) <> 1) Then
        TEMP_MATRIX(i, 11) = 1
    Else
        TEMP_MATRIX(i, 11) = 0
    End If
    
'----------------------------------------------------------------------------------------------------
    If TEMP_MATRIX(i - 1, 11) = 0 Then
'----------------------------------------------------------------------------------------------------
        If TEMP_MATRIX(i, 2) < TEMP_MATRIX(i - 1, 7) Then
            TEMP_MATRIX(i, 12) = TEMP_MATRIX(i, 2)
        Else
            TEMP_MATRIX(i, 12) = TEMP_MATRIX(i - 1, 7)
        End If
        TEMP_MATRIX(i, 13) = TEMP_MATRIX(i - 1, 11) * (TEMP_MATRIX(i, 5) - TEMP_MATRIX(i, 12)) / TEMP_MATRIX(i, 12) + TEMP_MATRIX(i, 11) * (TEMP_MATRIX(i, 5) - TEMP_MATRIX(i, 12)) / TEMP_MATRIX(i, 12)
'----------------------------------------------------------------------------------------------------
    Else
'----------------------------------------------------------------------------------------------------
        TEMP_MATRIX(i, 12) = TEMP_MATRIX(i - 1, 5)
        TEMP_MATRIX(i, 13) = TEMP_MATRIX(i - 1, 11) * (TEMP_MATRIX(i, 5) - TEMP_MATRIX(i, 12)) / TEMP_MATRIX(i, 12) + 0
'----------------------------------------------------------------------------------------------------
    End If
'----------------------------------------------------------------------------------------------------
    TEMP_MATRIX(i, 14) = TEMP_MATRIX(i, 13) - TEMP_MATRIX(i - 1, 11) * CASH_RATE / COUNT_BASIS
'----------------------------------------------------------------------------------------------------
    If i <> MA_PERIOD + 1 Then
'----------------------------------------------------------------------------------------------------
        TEMP_MATRIX(i, 15) = (1 + TEMP_MATRIX(i - 1, 15)) * (1 + TEMP_MATRIX(i, 13)) - 1 'Perfect
        If TEMP_MATRIX(i - 1, 16) > TEMP_MATRIX(i, 15) Then 'Perfect
            TEMP_MATRIX(i, 16) = TEMP_MATRIX(i - 1, 16)
        Else
            TEMP_MATRIX(i, 16) = TEMP_MATRIX(i, 15)
        End If
        TEMP_MATRIX(i, 17) = (1 + TEMP_MATRIX(i, 16)) / (1 + TEMP_MATRIX(i, 15)) - 1
        TEMP_MATRIX(i, 18) = IIf(TEMP_MATRIX(i, 17) = 0, 0, TEMP_MATRIX(i - 1, 18) + 1)
        If TEMP_MATRIX(i, 17) > MDRAWDOWN_VAL Then: MDRAWDOWN_VAL = TEMP_MATRIX(i, 17)
        If TEMP_MATRIX(i, 18) > k Then: k = TEMP_MATRIX(i, 18)
        TEMP2_SUM = TEMP2_SUM + TEMP_MATRIX(i, 14)
'----------------------------------------------------------------------------------------------------
    Else
'----------------------------------------------------------------------------------------------------
        TEMP_MATRIX(i, 15) = TEMP_MATRIX(i, 13) 'Perfect
        TEMP_MATRIX(i, 16) = TEMP_MATRIX(i, 15) 'Perfect
        TEMP_MATRIX(i, 17) = (1 + TEMP_MATRIX(i, 16)) / (1 + TEMP_MATRIX(i, 15)) - 1
        TEMP_MATRIX(i, 18) = IIf(TEMP_MATRIX(i, 17) = 0, 0, 0 + 1)
        MDRAWDOWN_VAL = TEMP_MATRIX(i, 17)
        k = TEMP_MATRIX(i, 18)
        TEMP2_SUM = TEMP_MATRIX(i, 14)
'----------------------------------------------------------------------------------------------------
    End If
'----------------------------------------------------------------------------------------------------
Next i
'----------------------------------------------------------------------------------------------------

'---------------------------------------------------------------------------------------------------
Select Case OUTPUT
'---------------------------------------------------------------------------------------------------
Case 0
'---------------------------------------------------------------------------------------------------
    ASSET_LOW_MA_SYSTEM_FUNC = TEMP_MATRIX
'---------------------------------------------------------------------------------------------------
Case Else
'---------------------------------------------------------------------------------------------------
    j = NROWS - MA_PERIOD
    LMEAN_VAL = TEMP2_SUM / j
    TEMP2_SUM = 0
    For i = MA_PERIOD + 1 To NROWS
        TEMP2_SUM = TEMP2_SUM + (TEMP_MATRIX(i, 14) - LMEAN_VAL) ^ 2
    Next i
    LVOLAT_VAL = (TEMP2_SUM / (j - 1)) ^ 0.5 'sample
    LSHARPE_VAL = LMEAN_VAL / LVOLAT_VAL * COUNT_BASIS ^ 0.5
    ASSET_LOW_MA_SYSTEM_FUNC = Array(LSHARPE_VAL, MDRAWDOWN_VAL, k)
'---------------------------------------------------------------------------------------------------
End Select
'---------------------------------------------------------------------------------------------------

Exit Function
ERROR_LABEL:
ASSET_LOW_MA_SYSTEM_FUNC = Err.number
End Function


