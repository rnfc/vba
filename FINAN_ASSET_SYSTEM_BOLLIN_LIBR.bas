Attribute VB_Name = "FINAN_ASSET_SYSTEM_BOLLIN_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.


Function ASSET_BOLLINGER_BANDS_SIGNAL_FUNC(ByVal TICKER_STR As Variant, _
Optional ByVal START_DATE As Date, _
Optional ByVal END_DATE As Date, _
Optional ByVal BOLLIN_UP_SD As Double = 2, _
Optional ByVal BOLLIN_DN_SD As Double = 2.5, _
Optional ByVal MA_PERIODS As Double = 26, _
Optional ByVal INITIAL_SHARES As Double = 1000, _
Optional ByVal SHARES_TRADED As Double = 50, _
Optional ByVal INITIAL_CASH As Double = 50000, _
Optional ByVal CASH_RATE As Double = 0.02, _
Optional ByVal COUNT_BASIS As Double = 365, _
Optional ByVal OUTPUT As Integer = 0)

Dim i As Long
Dim j As Long
Dim k As Long
Dim l As Long

Dim ii As Long
Dim jj As Long

Dim NROWS As Long

Dim TEMP_STR As String
Dim TEMP_VAL As Double
Dim TEMP_SUM As Double
Dim TEMP_DEV As Double
Dim TEMP_MEAN As Double
Dim TEMP_FACTOR As Double

Dim MEAN_VAL As Double
Dim SIGMA_VAL As Double

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

TEMP_FACTOR = (1 + CASH_RATE) ^ (1 / COUNT_BASIS) - 1

'----------------------------------------------------------------------------
ReDim TEMP_MATRIX(0 To NROWS, 1 To 18)
'----------------------------------------------------------------------------
TEMP_MATRIX(0, 1) = "DATE"
TEMP_MATRIX(0, 2) = "OPEN"
TEMP_MATRIX(0, 3) = "HIGH"
TEMP_MATRIX(0, 4) = "LOW"
TEMP_MATRIX(0, 5) = "CLOSE"
TEMP_MATRIX(0, 6) = "VOLUME"
TEMP_MATRIX(0, 7) = "ADJ CLOSE"
TEMP_MATRIX(0, 8) = "RETURNS"
TEMP_MATRIX(0, 9) = "PORTFOLIO (SCALED)"
TEMP_MATRIX(0, 10) = "BOLLI-UP"
TEMP_MATRIX(0, 11) = "BOLLI-DN"
TEMP_MATRIX(0, 12) = "BUY"
TEMP_MATRIX(0, 13) = "SELL"
TEMP_MATRIX(0, 14) = "SHARES HELD"
TEMP_MATRIX(0, 15) = "EQUITY"
TEMP_MATRIX(0, 16) = "CASH"
TEMP_MATRIX(0, 17) = "SYSTEM"
TEMP_MATRIX(0, 18) = "TRADES"

'----------------------------------------------------------------------------
i = 1
For j = 1 To 7
    TEMP_MATRIX(i, j) = DATA_MATRIX(i, j)
Next j
TEMP_MATRIX(i, 6) = TEMP_MATRIX(i, 6) / 1000
TEMP_MATRIX(i, 8) = ""
TEMP_MATRIX(i, 9) = TEMP_MATRIX(i, 7)
TEMP_MATRIX(i, 10) = ""
TEMP_MATRIX(i, 11) = ""
TEMP_MATRIX(i, 12) = ""
TEMP_MATRIX(i, 13) = ""
TEMP_MATRIX(i, 14) = INITIAL_SHARES
TEMP_MATRIX(i, 15) = TEMP_MATRIX(i, 7) * INITIAL_SHARES
TEMP_MATRIX(i, 16) = INITIAL_CASH
TEMP_MATRIX(i, 17) = TEMP_MATRIX(i, 15) + TEMP_MATRIX(i, 16)
TEMP_MATRIX(i, 18) = ""
'--------------------------------------------------------------------------
MEAN_VAL = 0
TEMP_SUM = TEMP_MATRIX(i, 7)
l = 0
'--------------------------------------------------------------------------
For i = 2 To NROWS
'--------------------------------------------------------------------------
    For j = 1 To 7
        TEMP_MATRIX(i, j) = DATA_MATRIX(i, j)
    Next j
    TEMP_MATRIX(i, 6) = TEMP_MATRIX(i, 6) / 1000
    TEMP_MATRIX(i, 8) = TEMP_MATRIX(i, 7) / TEMP_MATRIX(i - 1, 7) - 1
'--------------------------------------------------------------------------
    If i <= MA_PERIODS Then
'--------------------------------------------------------------------------
        TEMP_SUM = TEMP_SUM + TEMP_MATRIX(i, 7)
        TEMP_MEAN = TEMP_SUM / i
        TEMP_DEV = 0
        For k = i To 1 Step -1
            TEMP_DEV = TEMP_DEV + (TEMP_MATRIX(k, 7) - TEMP_MEAN) ^ 2
        Next k
        TEMP_MATRIX(i, 10) = TEMP_MEAN + BOLLIN_UP_SD * (TEMP_DEV / i) ^ 0.5
        TEMP_MATRIX(i, 11) = TEMP_MEAN - BOLLIN_DN_SD * (TEMP_DEV / i) ^ 0.5
'--------------------------------------------------------------------------
    Else
'--------------------------------------------------------------------------
        
        TEMP_SUM = TEMP_SUM + TEMP_MATRIX(i, 7)
        TEMP_MEAN = TEMP_SUM / (MA_PERIODS + 1)
        TEMP_DEV = 0
        For k = i To (i - MA_PERIODS) Step -1
            TEMP_DEV = TEMP_DEV + (TEMP_MATRIX(k, 7) - TEMP_MEAN) ^ 2
        Next k
        
        TEMP_MATRIX(i, 10) = TEMP_MEAN + BOLLIN_UP_SD * (TEMP_DEV / (MA_PERIODS + 1)) ^ 0.5
        TEMP_MATRIX(i, 11) = TEMP_MEAN - BOLLIN_DN_SD * (TEMP_DEV / (MA_PERIODS + 1)) ^ 0.5
        l = l + 1
        TEMP_SUM = TEMP_SUM - TEMP_MATRIX(l, 7)
'--------------------------------------------------------------------------
    End If
'--------------------------------------------------------------------------

'--------------------------------------------------------------------------
    If i <> 2 Then
'--------------------------------------------------------------------------
        If TEMP_MATRIX(i, 7) < TEMP_MATRIX(i, 11) Then
            TEMP_MATRIX(i, 12) = TEMP_MATRIX(i, 7)
            ii = 1
        Else
            TEMP_MATRIX(i, 12) = 0
            ii = 0
        End If
        
        If TEMP_MATRIX(i, 7) > TEMP_MATRIX(i, 10) Then
            TEMP_MATRIX(i, 13) = TEMP_MATRIX(i, 7)
            jj = 1
        Else
            TEMP_MATRIX(i, 13) = 0
            jj = 0
        End If
'--------------------------------------------------------------------------
    Else
'--------------------------------------------------------------------------
        If (TEMP_MATRIX(i, 7) < TEMP_MATRIX(i, 11) And _
           SHARES_TRADED * TEMP_MATRIX(i, 7) <= TEMP_MATRIX(i - 1, 16) * (1 + CASH_RATE)) Then
            TEMP_MATRIX(i, 12) = TEMP_MATRIX(i, 7)
            ii = 1
        Else
            TEMP_MATRIX(i, 12) = 0
            ii = 0
        End If
        
        If (TEMP_MATRIX(i, 7) > TEMP_MATRIX(i, 10) And TEMP_MATRIX(i - 1, 14) > SHARES_TRADED) Then
            TEMP_MATRIX(i, 13) = TEMP_MATRIX(i, 7)
            jj = 1
        Else
            TEMP_MATRIX(i, 13) = 0
            jj = 0
        End If
'--------------------------------------------------------------------------
    End If
'--------------------------------------------------------------------------
    TEMP_VAL = TEMP_MATRIX(i - 1, 14) + SHARES_TRADED * (ii - jj)
    
    
    TEMP_MATRIX(i, 14) = IIf(TEMP_VAL > 0, TEMP_VAL, 0)
    
    TEMP_STR = ""
    If TEMP_MATRIX(i, 14) > TEMP_MATRIX(i - 1, 14) Then
        ii = 1
        TEMP_STR = "Buy " & Format(TEMP_MATRIX(i, 14) - TEMP_MATRIX(i - 1, 14), "0")
    Else
        ii = 0
    End If
    
    If TEMP_MATRIX(i - 1, 14) > TEMP_MATRIX(i, 14) Then
        jj = 1
        TEMP_STR = "Sell " & Format(TEMP_MATRIX(i - 1, 14) - TEMP_MATRIX(i, 14), "0")
    Else
        jj = 0
    End If
    
    TEMP_MATRIX(i, 15) = TEMP_MATRIX(i, 7) * TEMP_MATRIX(i, 14)
    
    TEMP_VAL = TEMP_MATRIX(i - 1, 16) * (1 + TEMP_FACTOR) - _
              (TEMP_MATRIX(i, 14) - TEMP_MATRIX(i - 1, 14)) * _
               TEMP_MATRIX(i, 12) * ii + (TEMP_MATRIX(i - 1, 14) - _
               TEMP_MATRIX(i, 14)) * TEMP_MATRIX(i, 13) * jj
    
    TEMP_MATRIX(i, 16) = TEMP_VAL 'IIf(TEMP_VAL > 0, TEMP_VAL, 0)
    
    TEMP_MATRIX(i, 17) = TEMP_MATRIX(i, 15) + TEMP_MATRIX(i, 16)
    MEAN_VAL = MEAN_VAL + TEMP_MATRIX(i, 17) / TEMP_MATRIX(i - 1, 17) - 1
    TEMP_MATRIX(i, 18) = TEMP_STR
    
    TEMP_MATRIX(i, 9) = TEMP_MATRIX(i - 1, 9) * TEMP_MATRIX(i, 17) / TEMP_MATRIX(i - 1, 17)
'--------------------------------------------------------------------------
Next i
'--------------------------------------------------------------------------
If OUTPUT = 0 Then
    ASSET_BOLLINGER_BANDS_SIGNAL_FUNC = TEMP_MATRIX
    Exit Function
End If

MEAN_VAL = MEAN_VAL / (NROWS - 1)
SIGMA_VAL = 0
For i = 2 To NROWS
    SIGMA_VAL = SIGMA_VAL + ((TEMP_MATRIX(i, 17) / TEMP_MATRIX(i - 1, 17) - 1) - MEAN_VAL) ^ 2
Next i
SIGMA_VAL = (SIGMA_VAL / (NROWS - 1)) ^ 0.5

If OUTPUT = 1 Then
    ASSET_BOLLINGER_BANDS_SIGNAL_FUNC = MEAN_VAL / SIGMA_VAL
Else
    ASSET_BOLLINGER_BANDS_SIGNAL_FUNC = Array(MEAN_VAL / SIGMA_VAL, MEAN_VAL, SIGMA_VAL)
End If

Exit Function
ERROR_LABEL:
ASSET_BOLLINGER_BANDS_SIGNAL_FUNC = Err.number
End Function
