Attribute VB_Name = "FINAN_ASSET_TA_REPORT_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.

'************************************************************************************
'************************************************************************************
'FUNCTION      : ASSET_TA_REPORT_FUNC
'DESCRIPTION   : Technical Analysis Report
'LIBRARY       : FINAN_ASSET
'GROUP         : TA
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'************************************************************************************
'************************************************************************************

Function ASSET_TA_REPORT_FUNC(ByRef TICKERS_RNG As Variant, _
ByVal START_DATE As Date, _
ByVal END_DATE As Date, _
Optional ByVal EARLY_PERIOD As Long = 5, _
Optional ByVal SD_MULT As Double = 2, _
Optional ByVal MA1_PERIOD As Long = 20, _
Optional ByVal MA2_PERIOD As Long = 50, _
Optional ByVal MA3_PERIOD As Long = 100, _
Optional ByVal EMA1_PERIOD As Long = 12, _
Optional ByVal EMA2_PERIOD As Long = 26, _
Optional ByVal COUNT_BASIS As Double = 252, _
Optional ByRef HOLIDAYS_RNG As Variant)

Dim h As Long
Dim i As Long
Dim j As Long
Dim l As Long

Dim ii As Long
Dim jj As Long

Dim NO_DAYS As Long
Dim NROWS As Long
Dim NCOLUMNS As Long

Dim MIN_VAL As Double
Dim MAX_VAL As Double
Dim RSI_VAL As Double
Dim MAX_VOL_VAL As Double

Dim TEMP1_SUM As Double
Dim TEMP2_SUM As Double
Dim TEMP3_SUM As Double

Dim MEAN_VAL As Double
Dim SIGMA_VAL As Double

Dim LINE_STR As String
Dim TICKER_STR As String
Dim DATA_MATRIX As Variant

Dim TEMP_MATRIX As Variant
Dim TEMP_VECTOR() As String
Dim TICKERS_VECTOR As Variant

On Error GoTo ERROR_LABEL

If IsArray(TICKERS_RNG) = True Then
    TICKERS_VECTOR = TICKERS_RNG
    If UBound(TICKERS_VECTOR, 1) = 1 Then
        TICKERS_VECTOR = MATRIX_TRANSPOSE_FUNC(TICKERS_VECTOR)
    End If
Else
    ReDim TICKERS_VECTOR(1 To 1, 1 To 1)
    TICKERS_VECTOR(1, 1) = TICKERS_RNG
End If

NCOLUMNS = UBound(TICKERS_VECTOR, 1)
ReDim TEMP_VECTOR(1 To NCOLUMNS, 1 To 2)

'----------------------------------------------------------------------------------------------------------------------------------
For l = 1 To NCOLUMNS
'----------------------------------------------------------------------------------------------------------------------------------
    TICKER_STR = TICKERS_VECTOR(l, 1)
    DATA_MATRIX = YAHOO_HISTORICAL_DATA_SERIE_FUNC(TICKER_STR, START_DATE, END_DATE, "d", "DOHLCVA", False, True, True)
    If IsArray(DATA_MATRIX) = False Then: GoTo 1983
    NROWS = UBound(DATA_MATRIX, 1)
    NO_DAYS = NETWORKDAYS_FUNC(DATA_MATRIX(1, 1), DATA_MATRIX(NROWS, 1), HOLIDAYS_RNG)
    ReDim TEMP_MATRIX(0 To NROWS, 1 To 10)
    TEMP_MATRIX(0, 1) = "RETURNS"
    TEMP_MATRIX(0, 2) = "RSI"
    TEMP_MATRIX(0, 3) = "UP BOLLINGER"
    TEMP_MATRIX(0, 4) = "DN BOLLINGER"
    TEMP_MATRIX(0, 5) = EMA1_PERIOD & "-DAY EMA"
    TEMP_MATRIX(0, 6) = EMA2_PERIOD & "-DAY EMA"
    TEMP_MATRIX(0, 7) = "M.A.C.D."
    TEMP_MATRIX(0, 8) = MA1_PERIOD & " DAY E.M.A."
    TEMP_MATRIX(0, 9) = MA2_PERIOD & " DAY E.M.A."
    TEMP_MATRIX(0, 10) = MA3_PERIOD & " DAY E.M.A."
'----------------------------------------------------------------------------------------------------------------------------------
    MEAN_VAL = 0: SIGMA_VAL = 0
    MIN_VAL = 2 ^ 52
    MAX_VAL = MIN_VAL * -1
    MAX_VOL_VAL = MAX_VAL
    ii = 0: jj = 0
    TEMP1_SUM = 0: TEMP2_SUM = 0: TEMP3_SUM = 0
    For i = 1 To NROWS
        If i < MA1_PERIOD Then
            TEMP1_SUM = TEMP1_SUM + DATA_MATRIX(i, 7)
            h = 0
        Else
            TEMP1_SUM = TEMP1_SUM + DATA_MATRIX(i, 7)
            If i <> MA1_PERIOD Then
                h = i - MA1_PERIOD
                TEMP1_SUM = TEMP1_SUM - DATA_MATRIX(h, 7)
            End If
            MEAN_VAL = TEMP1_SUM / MA1_PERIOD
                            
            SIGMA_VAL = 0
            For j = i To h + 1 Step -1
                SIGMA_VAL = SIGMA_VAL + (DATA_MATRIX(j, 7) - MEAN_VAL) ^ 2
            Next j
            SIGMA_VAL = (SIGMA_VAL / MA1_PERIOD) ^ 0.5
                    
            TEMP_MATRIX(i, 3) = MEAN_VAL + SD_MULT * SIGMA_VAL
            TEMP_MATRIX(i, 4) = MEAN_VAL - SD_MULT * SIGMA_VAL
        End If
        
        If i <= MA1_PERIOD + SD_MULT Then
            If i <= MA1_PERIOD Then
                
                TEMP2_SUM = TEMP2_SUM + DATA_MATRIX(i, 7)
                TEMP_MATRIX(i, 5) = TEMP2_SUM / i
                
                TEMP3_SUM = TEMP3_SUM + DATA_MATRIX(i, 7)
                TEMP_MATRIX(i, 6) = TEMP_MATRIX(i, 5)
            Else
                h = i - MA1_PERIOD
                TEMP2_SUM = TEMP2_SUM - DATA_MATRIX(h, 7)

                TEMP2_SUM = TEMP2_SUM + DATA_MATRIX(i, 7)
                TEMP_MATRIX(i, 5) = TEMP2_SUM / MA1_PERIOD
                
                TEMP3_SUM = TEMP3_SUM + DATA_MATRIX(i, 7)
                TEMP_MATRIX(i, 6) = TEMP3_SUM / i
                
            End If
            TEMP_MATRIX(i, 7) = TEMP_MATRIX(i, 5) - TEMP_MATRIX(i, 6)
        Else
            TEMP_MATRIX(i, 5) = (1 - (SD_MULT / (EMA1_PERIOD + 1))) * TEMP_MATRIX(i - 1, 5) + (SD_MULT / (EMA1_PERIOD + 1)) * DATA_MATRIX(i, 7)
            TEMP_MATRIX(i, 6) = (1 - (SD_MULT / (EMA2_PERIOD + 1))) * TEMP_MATRIX(i - 1, 6) + (SD_MULT / (EMA2_PERIOD + 1)) * DATA_MATRIX(i, 7)
            TEMP_MATRIX(i, 7) = TEMP_MATRIX(i, 5) - TEMP_MATRIX(i, 6)
        End If
        GoSub EMA_LINE
        If DATA_MATRIX(i, 3) > MAX_VAL Then
            MAX_VAL = DATA_MATRIX(i, 3)
            ii = i
        End If
        If DATA_MATRIX(i, 4) < MIN_VAL Then
            MIN_VAL = DATA_MATRIX(i, 4)
            jj = i
        End If
        If DATA_MATRIX(i, 6) > MAX_VOL_VAL Then: MAX_VOL_VAL = DATA_MATRIX(i, 6)
    Next i
    
    GoSub REPORT_LINE
    TEMP_VECTOR(l, 1) = TICKER_STR
    TEMP_VECTOR(l, 2) = LINE_STR
    
1983:
Next l

ASSET_TA_REPORT_FUNC = TEMP_VECTOR

Exit Function
'-----------------------------------------------------------------------------------------
EMA_LINE:
'-----------------------------------------------------------------------------------------
    If i <> 1 Then
        TEMP_MATRIX(i, 1) = DATA_MATRIX(i, 7) / DATA_MATRIX(i - 1, 7) - 1
        TEMP_MATRIX(i, 2) = IIf(TEMP_MATRIX(i, 1) > 0, 1, 0)
        RSI_VAL = RSI_VAL + TEMP_MATRIX(i, 2)
        MEAN_VAL = MEAN_VAL + TEMP_MATRIX(i, 1)
        TEMP_MATRIX(i, 8) = (1 - (SD_MULT / (MA1_PERIOD + 1))) * TEMP_MATRIX(i - 1, 8) + (SD_MULT / (MA1_PERIOD + 1)) * DATA_MATRIX(i, 7)
        TEMP_MATRIX(i, 9) = (1 - (SD_MULT / (MA2_PERIOD + 1))) * TEMP_MATRIX(i - 1, 9) + (SD_MULT / (MA2_PERIOD + 1)) * DATA_MATRIX(i, 7)
        TEMP_MATRIX(i, 10) = (1 - (SD_MULT / (MA3_PERIOD + 1))) * TEMP_MATRIX(i - 1, 10) + (SD_MULT / (MA3_PERIOD + 1)) * DATA_MATRIX(i, 7)
    Else
        RSI_VAL = 0
        TEMP_MATRIX(1, 1) = 0
        TEMP_MATRIX(1, 2) = 0
        TEMP_MATRIX(1, 8) = TEMP_MATRIX(1, 5)
        TEMP_MATRIX(1, 9) = TEMP_MATRIX(1, 5)
        TEMP_MATRIX(1, 10) = TEMP_MATRIX(1, 5)
    End If
'-----------------------------------------------------------------------------------------
Return
'-----------------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------------
REPORT_LINE:
'-----------------------------------------------------------------------------------------
    LINE_STR = "For " & TICKER_STR & ", the price was " & _
                         Format(DATA_MATRIX(NROWS - EARLY_PERIOD, 7), "0.00") & _
                         ", " & Format(EARLY_PERIOD, "0") & _
                         " days ago, on " & _
                         Format(DATA_MATRIX(NROWS - EARLY_PERIOD, 1), "mmm d, yyyy") & _
                         ". It has since " & _
                         Format(IIf(DATA_MATRIX(NROWS - EARLY_PERIOD, 7) > _
                         DATA_MATRIX(NROWS, 7), "dropped", "risen"), "") _
                         & " to " & _
                         Format(DATA_MATRIX(NROWS, 7), "0.00") & " on " & _
                         Format(DATA_MATRIX(NROWS, 1), "mmm d, yyyy") & "."
            
    LINE_STR = LINE_STR & " The High and Low for " & Format(DATA_MATRIX(NROWS, 1), _
                         "mmm d, yyyy") & " are " & _
                         Format(DATA_MATRIX(NROWS, 3), "0.00") & _
                         " and " & _
                         Format(DATA_MATRIX(NROWS, 4), "0.00") & "."
            
    LINE_STR = LINE_STR & " The Upper Bollinger is " & _
                         Format(TEMP_MATRIX(NROWS, 3), "0.00") & _
                         ". The Lower Bollinger is " & _
                         Format(TEMP_MATRIX(NROWS, 4), "0.00") & "."
            
    LINE_STR = LINE_STR & " The Volume of trades was " & _
                         Format((DATA_MATRIX(NROWS - _
                         EARLY_PERIOD, 6) / 100000) / 10, "0.0") & _
                         " million, " & _
                         Format(EARLY_PERIOD, "0") & _
                         " days ago. It is now " & _
                         Format((DATA_MATRIX(NROWS, 6) / 100000) / 10, "0.0") & _
                         " million. Maximum shares traded were " _
                         & Format((MAX_VOL_VAL / 100000) / _
                         10, "0.0") & " million."
    
    LINE_STR = LINE_STR & " The minimum price over the last " & Format(NO_DAYS, "0") & _
                          " trading days was " & _
                         Format(MIN_VAL, "0.00") & ", on " & _
                         Format(DATA_MATRIX(jj, 1), "mmm d, yyyy") & _
                         ". The maximum price was " & _
                         Format(MAX_VAL, "0.00") & ", on " _
                         & Format(DATA_MATRIX(ii, 1), "mmm d, yyyy") & "."
    
    LINE_STR = LINE_STR & " The " & Format(TEMP_MATRIX(0, 8), "") & _
                         " was " & Format(TEMP_MATRIX(NROWS - EARLY_PERIOD, 8), "0.00") & _
                         ". It has since " & _
                         Format(IIf(TEMP_MATRIX(NROWS - EARLY_PERIOD, 8) > _
                         TEMP_MATRIX(NROWS, 8), _
                         "dropped", "risen"), "") & " to " & _
                         Format(TEMP_MATRIX(NROWS, 8), "0.00") & "."
    
    LINE_STR = LINE_STR & " The " & Format(TEMP_MATRIX(0, 9), "") & _
                         " was " & Format(TEMP_MATRIX(NROWS - EARLY_PERIOD, 9), "0.00") & _
                         ". It has since " & _
                         Format(IIf(TEMP_MATRIX(NROWS - EARLY_PERIOD, 9) > _
                         TEMP_MATRIX(NROWS, 9), _
                         "dropped", "risen"), "") & " to " & _
                         Format(TEMP_MATRIX(NROWS, 9), "0.00") & "."
    
    LINE_STR = LINE_STR & " The " & Format(TEMP_MATRIX(0, 10), "") & _
                         " was " & Format(TEMP_MATRIX(NROWS - EARLY_PERIOD, 10), "0.00") & _
                         ". It has since " & _
                         Format(IIf(TEMP_MATRIX(NROWS - EARLY_PERIOD, 10) > _
                         TEMP_MATRIX(NROWS, 10), _
                         "dropped", "risen"), "") & " to " & _
                         Format(TEMP_MATRIX(NROWS, 10), "0.00") & "."
    
    LINE_STR = LINE_STR & " The M.A.C.D. was " & _
                         Format(TEMP_MATRIX(NROWS - EARLY_PERIOD, 7), "0.00") & ", " & _
                         Format(EARLY_PERIOD, "0") & _
                         " days ago. It has since " & _
                         Format(IIf(TEMP_MATRIX(NROWS - EARLY_PERIOD, 7) > _
                         TEMP_MATRIX(NROWS, 7), _
                         "dropped", "risen"), "") & " to " & _
                         Format(TEMP_MATRIX(NROWS, 7), "0.00") & "."
    LINE_STR = LINE_STR & " The R.S.I. over the past " & MA3_PERIOD & _
                         " days is " & _
                         Format(RSI_VAL / (NROWS - _
                         MA3_PERIOD), "0%") & "."
    MEAN_VAL = 0
    For i = 2 To NROWS
        MEAN_VAL = MEAN_VAL + TEMP_MATRIX(i, 1)
    Next i
    MEAN_VAL = MEAN_VAL / (NROWS - 1)
    SIGMA_VAL = 0
    For i = 2 To NROWS
        SIGMA_VAL = SIGMA_VAL + (TEMP_MATRIX(i, 1) - MEAN_VAL) ^ 2
    Next i
    SIGMA_VAL = (SIGMA_VAL / (NROWS - 1)) ^ 0.5
    LINE_STR = LINE_STR & " From " & Format(DATA_MATRIX(1, 1), "mmm d, yyyy") & " to " & Format(DATA_MATRIX(NROWS, 1), "mmm d, yyyy")
    LINE_STR = LINE_STR & " the Annualized Volatility for this asset was " & Format(SIGMA_VAL * COUNT_BASIS ^ 0.5, "0.00%")
    LINE_STR = LINE_STR & " and the Annualized Return was " & Format((DATA_MATRIX(NROWS, 7) / DATA_MATRIX(1, 7)) ^ (COUNT_BASIS / NO_DAYS) - 1, "0.00%") & "."
'-----------------------------------------------------------------------------------------
Return
'-----------------------------------------------------------------------------------------
ERROR_LABEL:
ASSET_TA_REPORT_FUNC = Err.number
End Function
