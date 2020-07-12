Attribute VB_Name = "FINAN_ASSET_TA_TABLES_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.

'************************************************************************************
'************************************************************************************
'FUNCTION      : ASSET_TA_TABLE1_FUNC
'DESCRIPTION   :
'LIBRARY       : FINAN_ASSET
'GROUP         : TA_TABLES
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 03/16/2009
'************************************************************************************
'************************************************************************************

Function ASSET_TA_TABLE1_FUNC(ByRef TICKERS_RNG As Variant, _
Optional ByVal INDEX_SYMBOL As String = "^GSPC", _
Optional ByVal START_DATE As Date = 0, _
Optional ByVal END_DATE As Date = 0, _
Optional ByVal CASH_RATE As Double = 0.04, _
Optional ByVal RETURNS_PERIOD As Long = 252, _
Optional ByVal MA_PERIOD As Long = 200, _
Optional ByRef EMA1_PERIOD As Long = 20, _
Optional ByRef EMA2_PERIOD As Long = 5, _
Optional ByVal AROON_PERIOD As Long = 10, _
Optional ByVal WILLIAMS_PERIOD As Long = 14, _
Optional ByVal NBINS As Long = 37, _
Optional ByVal CONFIDENCE_VAL As Double = 0.015)

'CASH_RATE: Annual Risk-free Rate
'RETURNS_PERIOD: T Days
'NBINS: reference to intervals into which you want to group the values
'in the ito probability array

Dim h As Long
Dim i As Long
Dim j As Long
Dim k As Long
Dim l As Long

Dim m As Long
Dim n As Long

Dim NROWS As Long
Dim NSIZE As Long

Dim PI_VAL As Double

Dim TEMP1_VAL As Double
Dim TEMP2_VAL As Double
Dim TEMP3_VAL As Double

Dim RET1_VAL As Double
Dim RET2_VAL As Double
Dim RET3_VAL As Double

Dim SUM_VAL As Double
Dim MEAN_VAL As Double
Dim MAX_VAL As Double
Dim MIN_VAL As Double
Dim MULT_VAL As Double

Dim DEV1_VAL As Double
Dim DEV2_VAL As Double

Dim TEMP_MATRIX As Variant
Dim TICKERS_VECTOR As Variant

Dim DATA_VECTOR As Variant
Dim MARKET_VECTOR As Variant

Dim DATA_MATRIX As Variant
Dim MARKET_MATRIX As Variant

On Error Resume Next

If IsArray(TICKERS_RNG) = True Then
    TICKERS_VECTOR = TICKERS_RNG
    If UBound(TICKERS_VECTOR, 1) = 1 Then: _
        TICKERS_VECTOR = MATRIX_TRANSPOSE_FUNC(TICKERS_VECTOR)
Else
    ReDim TICKERS_VECTOR(1 To 1, 1 To 1)
    TICKERS_VECTOR(1, 1) = TICKERS_RNG
End If

NSIZE = UBound(TICKERS_VECTOR, 1)

If END_DATE = 0 Then
    END_DATE = Now
    END_DATE = DateSerial(Year(END_DATE), Month(END_DATE), Day(END_DATE))
End If

If START_DATE = 0 Then
    START_DATE = DateSerial(Year(END_DATE) - 1, Month(END_DATE), Day(END_DATE) - RETURNS_PERIOD)
End If

PI_VAL = 3.14159265358979
TEMP1_VAL = 1 - 2 / (EMA1_PERIOD + 1)
TEMP2_VAL = 1 - 2 / (EMA2_PERIOD + 1)

MARKET_MATRIX = YAHOO_HISTORICAL_DATA_SERIE_FUNC(Trim(INDEX_SYMBOL), _
                START_DATE, END_DATE, "d", "DOHLCVA", False, True, True)
If IsArray(MARKET_MATRIX) = False Then: GoTo ERROR_LABEL

ReDim TEMP_MATRIX(0 To NSIZE, 1 To 17)
TEMP_MATRIX(0, 1) = "TICKER"
TEMP_MATRIX(0, 2) = "MA_" & MA_PERIOD & "_DAYS" 'Days in Moving Average
TEMP_MATRIX(0, 3) = "PROB_OF_INCREASE"
TEMP_MATRIX(0, 4) = "EMA_" & EMA1_PERIOD & "_DAYS" 'Days in Exp. Moving Average #1
TEMP_MATRIX(0, 5) = "EMA_" & EMA2_PERIOD & "_DAYS" 'Days in Moving Average #2
TEMP_MATRIX(0, 6) = "CAGR" 'Compound Annual Growth Rate (Growth of $1.00)
TEMP_MATRIX(0, 7) = "WILLIAMS_%_R_" & WILLIAMS_PERIOD & "_DAYS"
TEMP_MATRIX(0, 8) = "DRAWDOWN_PERCENT" 'Maximum drop from previous high.
TEMP_MATRIX(0, 9) = "SHARPE_RATIO"
TEMP_MATRIX(0, 10) = INDEX_SYMBOL & "_BETA"
TEMP_MATRIX(0, 11) = "AROON_OSCILLATOR_" & AROON_PERIOD & "_DAYS"
TEMP_MATRIX(0, 12) = "RSL_" & EMA2_PERIOD & "_DAYS"
TEMP_MATRIX(0, 13) = RETURNS_PERIOD & "_DAYS_RETURN"
TEMP_MATRIX(0, 14) = "STDEVP"
TEMP_MATRIX(0, 15) = "MEAN"
TEMP_MATRIX(0, 16) = "PRICE_P0"
TEMP_MATRIX(0, 17) = "T_DAYS"

'-------------------------------------------------------------------------------------
For j = 1 To NSIZE
'-------------------------------------------------------------------------------------
    
    TEMP_MATRIX(j, 1) = Trim(TICKERS_VECTOR(j, 1))
    If TEMP_MATRIX(j, 1) = "" Then: GoTo 1983
    DATA_MATRIX = YAHOO_HISTORICAL_DATA_SERIE_FUNC(TEMP_MATRIX(j, 1), _
                  START_DATE, END_DATE, "d", "DOHLCVA", False, True, True)
    If IsArray(DATA_MATRIX) = False Then: GoTo 1983
    
    NROWS = UBound(DATA_MATRIX, 1) 'Excluding Headings
    
    l = NROWS - RETURNS_PERIOD
    
    If l <= 0 Then: l = 3
    TEMP_MATRIX(j, 6) = 1
    MAX_VAL = TEMP_MATRIX(j, 6)
    TEMP_MATRIX(j, 16) = DATA_MATRIX(l, 7)
    TEMP_MATRIX(j, 17) = RETURNS_PERIOD
    MEAN_VAL = 0
    
    ReDim DATA_VECTOR(1 To l - 1, 1 To 1)
    ReDim MARKET_VECTOR(1 To l - 1, 1 To 1)
'-----------------------------------------------------------------------------------------------
    For i = 1 To l
'-----------------------------------------------------------------------------------------------
        If i >= (l - MA_PERIOD) Then
            TEMP_MATRIX(j, 2) = TEMP_MATRIX(j, 2) + DATA_MATRIX(i, 7)
        End If
        
        If i <> 1 Then 'Delayed Returns
            
            RET2_VAL = RET1_VAL
            TEMP_MATRIX(j, 15) = TEMP_MATRIX(j, 15) + RET2_VAL
            
            RET1_VAL = DATA_MATRIX(i, 7) / DATA_MATRIX(i - 1, 7) - 1
            TEMP_MATRIX(j, 4) = TEMP1_VAL * TEMP_MATRIX(j, 4) + (1 - TEMP1_VAL) * DATA_MATRIX(i, 7)
            TEMP_MATRIX(j, 5) = TEMP2_VAL * TEMP_MATRIX(j, 5) + (1 - TEMP2_VAL) * DATA_MATRIX(i, 7)
        
            DATA_VECTOR(i - 1, 1) = RET1_VAL
            MARKET_VECTOR(i - 1, 1) = MARKET_MATRIX(i, 7) / MARKET_MATRIX(i - 1, 7) - 1
        
        Else
            RET1_VAL = DATA_MATRIX(i, 5) / DATA_MATRIX(i, 2) - 1
            TEMP_MATRIX(j, 4) = DATA_MATRIX(i, 7)
            TEMP_MATRIX(j, 5) = DATA_MATRIX(i, 7)
        End If
        
        MEAN_VAL = MEAN_VAL + RET1_VAL
        TEMP_MATRIX(j, 6) = TEMP_MATRIX(j, 6) * (1 + RET1_VAL)
        
        MAX_VAL = 1
        For k = 1 To i 'Drawdown
            If k = 1 Then
                MULT_VAL = 1
            Else
                If k <> 2 Then
                    RET3_VAL = DATA_MATRIX(k - 1, 7) / DATA_MATRIX(k - 2, 7) - 1
                Else
                    RET3_VAL = DATA_MATRIX(k - 1, 5) / DATA_MATRIX(k - 1, 2) - 1
                End If
                MULT_VAL = MULT_VAL * (1 + RET3_VAL)
            End If
            If MULT_VAL > MAX_VAL Then: MAX_VAL = MULT_VAL
            TEMP_MATRIX(j, 8) = 1 - MULT_VAL / MAX_VAL
        Next k
    Next i
    TEMP_MATRIX(j, 2) = TEMP_MATRIX(j, 2) / (MA_PERIOD + 1)
    TEMP_MATRIX(j, 6) = TEMP_MATRIX(j, 6) - 1
    
    MEAN_VAL = MEAN_VAL / l
    TEMP_MATRIX(j, 15) = TEMP_MATRIX(j, 15) / (l - 1)
    
    DEV1_VAL = 0: DEV2_VAL = 0
'-----------------------------------------------------------------------------------------------
    For i = 1 To l
'-----------------------------------------------------------------------------------------------
        If i <> 1 Then 'Delayed Returns
            RET2_VAL = RET1_VAL
            DEV2_VAL = DEV2_VAL + (RET2_VAL - TEMP_MATRIX(j, 15)) ^ 2
            
            RET1_VAL = DATA_MATRIX(i, 7) / DATA_MATRIX(i - 1, 7) - 1
        Else
            RET1_VAL = DATA_MATRIX(i, 5) / DATA_MATRIX(i, 2) - 1
        End If
        DEV1_VAL = DEV1_VAL + (RET1_VAL - MEAN_VAL) ^ 2
        
        If i >= AROON_PERIOD Then
            MAX_VAL = -2 ^ 52: MIN_VAL = 2 ^ 52
            h = 0: m = 0: n = 0
            For k = (i - AROON_PERIOD + 1) To i
                h = h + 1
                If DATA_MATRIX(k, 5) > MAX_VAL Then
                    MAX_VAL = DATA_MATRIX(k, 5)
                    m = h
                End If
                If DATA_MATRIX(k, 5) < MIN_VAL Then
                    MIN_VAL = DATA_MATRIX(k, 5)
                    n = h
                End If
            Next k
            TEMP_MATRIX(j, 11) = AROON_PERIOD * m - AROON_PERIOD * n
        End If
'-----------------------------------------------------------------------------------------------
    Next i
'-----------------------------------------------------------------------------------------------
    
    TEMP_MATRIX(j, 7) = (DEV1_VAL / l) ^ 0.5 * Sqr(l)
    TEMP_MATRIX(j, 12) = DATA_MATRIX(l, 7) / TEMP_MATRIX(j, 5)
    
    TEMP_MATRIX(j, 14) = (DEV2_VAL / (l - 1)) ^ 0.5
    
    k = l + RETURNS_PERIOD
    If k > NROWS Then: k = NROWS
    
    TEMP_MATRIX(j, 13) = DATA_MATRIX(k, 7) / DATA_MATRIX(l, 7) - 1
    TEMP_MATRIX(j, 9) = (TEMP_MATRIX(j, 15) * l - CASH_RATE) / TEMP_MATRIX(j, 14) / Sqr(l)
    
    TEMP_MATRIX(j, 10) = REGRESSION_SIMPLE_COEF_FUNC(MARKET_VECTOR, DATA_VECTOR, "")(1, 1)

    k = (NBINS - 1) / 2
    SUM_VAL = 0
    TEMP_MATRIX(j, 3) = 0
    MULT_VAL = TEMP_MATRIX(j, 16) * (1 - CONFIDENCE_VAL) ^ k
'-----------------------------------------------------------------------------------------------
    For i = 1 To NBINS 'Ito Probability
'-----------------------------------------------------------------------------------------------
    'Ito's Stochastic Calculus and Probability Theory by
    'S. Watanabe, Hiroshi Kunita, M. Fukushima, N. Ikeda
    'Publisher: Springer-Verlag New York, LLC
    'Pub. Date: January 1996
    'ISBN-13: 9784431701866; 440pp
        TEMP3_VAL = 1 / (TEMP_MATRIX(j, 14) * _
                    Sqr(2 * PI_VAL * TEMP_MATRIX(j, 17))) / MULT_VAL * _
                    Exp(-(1 / (2 * TEMP_MATRIX(j, 17) * _
                    TEMP_MATRIX(j, 14) ^ 2)) * _
                    (Log(MULT_VAL / TEMP_MATRIX(j, 16)) - _
                    (TEMP_MATRIX(j, 15) - 0.5 * TEMP_MATRIX(j, 14) ^ 2) * _
                    TEMP_MATRIX(j, 17)) ^ 2)
        If i <= k Then
            MULT_VAL = TEMP_MATRIX(j, 16) * (1 - CONFIDENCE_VAL) ^ (k - i)
        Else
            MULT_VAL = MULT_VAL * (1 + CONFIDENCE_VAL)
        End If
        If i <= k + 1 Then
            TEMP_MATRIX(j, 3) = TEMP_MATRIX(j, 3) + TEMP3_VAL
        End If
        SUM_VAL = SUM_VAL + TEMP3_VAL
'-----------------------------------------------------------------------------------------------
    Next i
'-----------------------------------------------------------------------------------------------
    TEMP_MATRIX(j, 3) = TEMP_MATRIX(j, 3) / SUM_VAL
1983:
'-------------------------------------------------------------------------------------
Next j
'-------------------------------------------------------------------------------------

ASSET_TA_TABLE1_FUNC = TEMP_MATRIX

Exit Function
ERROR_LABEL:
ASSET_TA_TABLE1_FUNC = Err.number
End Function

Function ASSET_TA_TABLE2_FUNC(ByRef TICKERS_RNG As Variant, _
Optional ByVal START_DATE As Date = 0, _
Optional ByVal END_DATE As Date = 0, _
Optional ByVal CASH_RATE As Double = 0.02, _
Optional ByVal SMA_PERIOD As Long = 20, _
Optional ByVal MACD1_PERIOD As Long = 12, _
Optional ByVal MACD2_PERIOD As Long = 26, _
Optional ByVal EMA_PERIOD As Long = 9, _
Optional ByVal BOLLI_PERIOD As Long = 20, _
Optional ByVal BOLLI_DEVIATIONS As Double = 2, _
Optional ByVal RSI_PERIOD As Long = 14, _
Optional ByVal K_PERIOD As Long = 14, _
Optional ByVal D_PERIOD As Long = 3, _
Optional ByVal ZWEIG_RULE As Double = 0.04, _
Optional ByVal ZWEIG_PERIOD As Long = 14, _
Optional ByVal COUNT_BASIS As Long = 250)

'MACD: Buy if greater than 0.10 // Sell if less than -0.10
'CAGR: Buy if greater than 0.08 // Sell if less than -0.05
'VOLATILITY: Buy if greater than 0.50 // Sell if less than 0.10
'P/SMA: Buy if greater than 1.00 // Sell if less than 0.90
'AVERAGE RETURN: Buy if greater than 0.10 // Sell if less than -0.10
'SHARPE: Buy if greater than 0.00 // Sell if less than 0.00
'100-RSI: Buy if greater than 70.00 // Sell if less than 40.00
'%K: Buy if greater than 80.00 // Sell if less than 20.00
'%D: Buy if greater than 80.00 // Sell if less than 20.00
'SORTINO: Buy if greater than 0.00 // Sell if less than 0.00
'ZWEIG 4%: Buy if greater than 0.04 // Sell if less than -0.04

Dim i As Long
Dim j As Long
Dim k As Long
Dim l As Long

Dim ii As Long 'RSI-UP
Dim jj As Long 'RSI-DN

Dim NROWS As Long
Dim NSIZE As Long

Dim SMA_VAL As Double
Dim BOLLI_MEAN_VAL As Double
Dim BOLLI_SIGMA_VAL As Double

Dim MACD1_VAL As Double
Dim MACD2_VAL As Double
Dim EMA_VAL As Double

Dim MACD1_MULT As Double
Dim MACD2_MULT As Double

Dim EMA_MULT As Double
Dim EMA_FACTOR As Double

Dim TEMP_VAL As Double

Dim RSI_UP_VAL As Double
Dim RSI_DN_VAL As Double

Dim MIN_VAL As Double
Dim MAX_VAL As Double

Dim CL_VAL As Double
Dim HL_VAL As Double

Dim K_SUM As Double
Dim K_ARR() As Double

Dim SORTINO_SUM As Double
Dim ZWEIG_VAL As Double

Dim MEAN_VAL As Double
Dim SIGMA_VAL As Double
Dim YEARFRAC_VAL As Double

Dim TEMP_MATRIX As Variant
Dim TICKERS_VECTOR As Variant

Dim DATA_MATRIX As Variant

Dim tolerance As Double

On Error Resume Next

tolerance = 10 ^ -15
If IsArray(TICKERS_RNG) = True Then
    TICKERS_VECTOR = TICKERS_RNG
    If UBound(TICKERS_VECTOR, 1) = 1 Then: _
        TICKERS_VECTOR = MATRIX_TRANSPOSE_FUNC(TICKERS_VECTOR)
Else
    ReDim TICKERS_VECTOR(1 To 1, 1 To 1)
    TICKERS_VECTOR(1, 1) = TICKERS_RNG
End If

NSIZE = UBound(TICKERS_VECTOR, 1)

If END_DATE = 0 Then
    END_DATE = Now
    END_DATE = DateSerial(Year(END_DATE), Month(END_DATE), Day(END_DATE))
End If

If START_DATE = 0 Then
    START_DATE = DateSerial(Year(END_DATE) - 1, Month(END_DATE), Day(END_DATE))
End If

YEARFRAC_VAL = COUNT_DAYS_FUNC(START_DATE, END_DATE, 1) / COUNT_BASIS '365

MACD1_MULT = 2 / (1 + MACD1_PERIOD)
MACD2_MULT = 2 / (1 + MACD2_PERIOD)
EMA_MULT = 2 / (1 + EMA_PERIOD)

ReDim TEMP_MATRIX(0 To NSIZE, 1 To 17)
TEMP_MATRIX(0, 1) = "TICKER"
TEMP_MATRIX(0, 2) = "ADJ_CLOSE"
TEMP_MATRIX(0, 3) = "SMA: " & SMA_PERIOD
TEMP_MATRIX(0, 4) = "MACD: " & MACD1_PERIOD & "." & MACD2_PERIOD
TEMP_MATRIX(0, 5) = "EMA: " & EMA_PERIOD
TEMP_MATRIX(0, 6) = "BOLLI UP: " & BOLLI_PERIOD & "-" & BOLLI_DEVIATIONS & " SD"
TEMP_MATRIX(0, 7) = "BOLLI DN: " & BOLLI_PERIOD & "-" & BOLLI_DEVIATIONS & " SD"
TEMP_MATRIX(0, 8) = "CAGR"
TEMP_MATRIX(0, 9) = "VOLATILITY (ANNUAL)"
TEMP_MATRIX(0, 10) = "P/SMA"
TEMP_MATRIX(0, 11) = "AVERAGE RETURN (ANNUAL)"
TEMP_MATRIX(0, 12) = "SHARPE RATIO"
TEMP_MATRIX(0, 13) = "100-RSI: " & RSI_PERIOD
'I use 100 - RSI so that BUY signals are greater than something
'rather than less than something
TEMP_MATRIX(0, 14) = "%K: " & K_PERIOD
TEMP_MATRIX(0, 15) = "%D: " & D_PERIOD
TEMP_MATRIX(0, 16) = "SORTINO RATIO"
TEMP_MATRIX(0, 17) = "ZWEIG " & Format(ZWEIG_RULE, "0.0%")

'-------------------------------------------------------------------------------------
For j = 1 To NSIZE
'-------------------------------------------------------------------------------------
    
    TEMP_MATRIX(j, 1) = Trim(TICKERS_VECTOR(j, 1))
    If TEMP_MATRIX(j, 1) = "" Then: GoTo 1983
    DATA_MATRIX = YAHOO_HISTORICAL_DATA_SERIE_FUNC(TEMP_MATRIX(j, 1), _
                  START_DATE, END_DATE, "d", "DOHLCVA", False, True, True)
    If IsArray(DATA_MATRIX) = False Then: GoTo 1983
    
    NROWS = UBound(DATA_MATRIX, 1) 'Excluding Headings
    
    TEMP_MATRIX(j, 2) = DATA_MATRIX(NROWS, 7)
    
    ii = 0: jj = 0
    
    MEAN_VAL = 0: SIGMA_VAL = 0
    SMA_VAL = 0: EMA_VAL = 0
    
    BOLLI_MEAN_VAL = 0: BOLLI_SIGMA_VAL = 0
    RSI_UP_VAL = 0: RSI_DN_VAL = 0
    
    K_SUM = 0
    ReDim K_ARR(1 To NROWS)
    SORTINO_SUM = 0
    
    i = 1
    MACD1_VAL = MACD1_MULT * DATA_MATRIX(i, 7)
    MACD2_VAL = MACD2_MULT * DATA_MATRIX(i, 7)
    EMA_FACTOR = EMA_MULT * DATA_MATRIX(i, 7)
    
    MEAN_VAL = MEAN_VAL + (DATA_MATRIX(i, 5) / DATA_MATRIX(i, 2) - 1)
    For i = 2 To NROWS
        If i >= (NROWS - SMA_PERIOD) Then: SMA_VAL = SMA_VAL + DATA_MATRIX(i, 7)
        If i >= (NROWS - EMA_PERIOD) Then: EMA_VAL = EMA_VAL + DATA_MATRIX(i, 7)
        
        MACD1_VAL = MACD1_MULT * DATA_MATRIX(i, 7) + (1 - MACD1_MULT) * MACD1_VAL
        MACD2_VAL = MACD2_MULT * DATA_MATRIX(i, 7) + (1 - MACD2_MULT) * MACD2_VAL
        EMA_FACTOR = EMA_MULT * DATA_MATRIX(i, 7) + (1 - EMA_MULT) * EMA_FACTOR
        
        If i > (NROWS - BOLLI_PERIOD) Then: BOLLI_MEAN_VAL = BOLLI_MEAN_VAL + DATA_MATRIX(i, 7)
        MEAN_VAL = MEAN_VAL + (DATA_MATRIX(i, 7) / DATA_MATRIX(i - 1, 7) - 1)
    Next i
    
    MEAN_VAL = MEAN_VAL / NROWS
    BOLLI_MEAN_VAL = BOLLI_MEAN_VAL / BOLLI_PERIOD
    
    i = 1
    SIGMA_VAL = SIGMA_VAL + ((DATA_MATRIX(i, 5) / DATA_MATRIX(i, 2) - 1) - MEAN_VAL) ^ 2
    TEMP_VAL = (DATA_MATRIX(i, 7) / DATA_MATRIX(i, 2) - 1)
    SORTINO_SUM = SORTINO_SUM + IIf(TEMP_VAL < MEAN_VAL, (TEMP_VAL - MEAN_VAL) ^ 2, 0)
    
    For i = 2 To NROWS
        If i > (NROWS - BOLLI_PERIOD) Then
            BOLLI_SIGMA_VAL = BOLLI_SIGMA_VAL + (DATA_MATRIX(i, 7) - BOLLI_MEAN_VAL) ^ 2
        End If
        SIGMA_VAL = SIGMA_VAL + ((DATA_MATRIX(i, 7) / DATA_MATRIX(i - 1, 7) - 1) - MEAN_VAL) ^ 2
    
        If i > (NROWS - RSI_PERIOD) Then
            TEMP_VAL = DATA_MATRIX(i, 7) - DATA_MATRIX(i - 1, 7)
            If TEMP_VAL > tolerance Then
                RSI_UP_VAL = RSI_UP_VAL + TEMP_VAL
                ii = ii + 1
            End If
            
            TEMP_VAL = DATA_MATRIX(i - 1, 7) - DATA_MATRIX(i, 7)
            If TEMP_VAL > tolerance Then
                RSI_DN_VAL = RSI_DN_VAL + TEMP_VAL
                jj = jj + 1
            End If
        End If
        
        If i >= K_PERIOD Then
            l = i - K_PERIOD + 1
        Else
            l = 1
        End If
        
        MIN_VAL = 2 ^ 52
        MAX_VAL = -2 ^ 52
        For k = i To l Step -1
            If DATA_MATRIX(k, 7) > MAX_VAL Then: MAX_VAL = DATA_MATRIX(k, 7)
            If DATA_MATRIX(k, 7) < MIN_VAL Then: MIN_VAL = DATA_MATRIX(k, 7)
        Next k
        CL_VAL = DATA_MATRIX(i, 7) - MIN_VAL
        HL_VAL = MAX_VAL - MIN_VAL
        K_ARR(i) = 100 * CL_VAL / HL_VAL
        
        If i > D_PERIOD Then
            l = i - D_PERIOD
            K_SUM = K_SUM - K_ARR(l)
        End If
        K_SUM = K_SUM + K_ARR(i)
        
        TEMP_VAL = (DATA_MATRIX(i, 7) / DATA_MATRIX(i - 1, 7) - 1)
        SORTINO_SUM = SORTINO_SUM + IIf(TEMP_VAL < MEAN_VAL, (TEMP_VAL - MEAN_VAL) ^ 2, 0)
        
        If i > ZWEIG_PERIOD Then
            l = i - ZWEIG_PERIOD
            If DATA_MATRIX(l, 7) > MAX_VAL Then: MAX_VAL = DATA_MATRIX(l, 7)
            If DATA_MATRIX(l, 7) < MIN_VAL Then: MIN_VAL = DATA_MATRIX(l, 7)
        End If
        '0.0% means the Zweig Rules aren't satisfied
        '… else this gives ratios:
        'Price/x-day Min
        '(when it's greater than Rule%)
        'or
        'Price/x-day Max
        '(when it's less than Rule%)
        If DATA_MATRIX(i, 7) > (1 + ZWEIG_RULE) * MIN_VAL Then
            ZWEIG_VAL = DATA_MATRIX(i, 7) / MIN_VAL - 1
        Else
            If DATA_MATRIX(i, 7) < (1 - ZWEIG_RULE) * MAX_VAL Then
                ZWEIG_VAL = DATA_MATRIX(i, 7) / MAX_VAL - 1
            Else
                ZWEIG_VAL = 0
            End If
        End If
    Next i
    Erase K_ARR()
        
    SIGMA_VAL = (SIGMA_VAL / NROWS) ^ 0.5
    BOLLI_SIGMA_VAL = (BOLLI_SIGMA_VAL / BOLLI_PERIOD) ^ 0.5
    
    RSI_UP_VAL = RSI_UP_VAL / ii
    RSI_DN_VAL = RSI_DN_VAL / jj
    
    TEMP_MATRIX(j, 3) = SMA_VAL / (SMA_PERIOD + 1)
    TEMP_MATRIX(j, 4) = MACD1_VAL - MACD2_VAL
    TEMP_MATRIX(j, 5) = EMA_VAL / (EMA_PERIOD + 1)
    TEMP_MATRIX(j, 6) = BOLLI_MEAN_VAL + BOLLI_DEVIATIONS * BOLLI_SIGMA_VAL
    TEMP_MATRIX(j, 7) = BOLLI_MEAN_VAL - BOLLI_DEVIATIONS * BOLLI_SIGMA_VAL
    TEMP_MATRIX(j, 8) = (DATA_MATRIX(NROWS, 7) / DATA_MATRIX(1, 7)) ^ YEARFRAC_VAL - 1
    TEMP_MATRIX(j, 9) = SIGMA_VAL * COUNT_BASIS ^ 0.5
    TEMP_MATRIX(j, 10) = TEMP_MATRIX(j, 2) / TEMP_MATRIX(j, 6)
    TEMP_MATRIX(j, 11) = MEAN_VAL * COUNT_BASIS '365
    TEMP_MATRIX(j, 12) = (TEMP_MATRIX(j, 11) - CASH_RATE) / TEMP_MATRIX(j, 9)
    TEMP_MATRIX(j, 13) = 100 * RSI_DN_VAL / (RSI_UP_VAL + RSI_DN_VAL)
    TEMP_MATRIX(j, 14) = 100 * CL_VAL / HL_VAL
    TEMP_MATRIX(j, 15) = K_SUM / D_PERIOD
    TEMP_MATRIX(j, 16) = (TEMP_MATRIX(j, 11) - CASH_RATE) / _
                        ((SORTINO_SUM / NROWS) ^ 0.5 * (COUNT_BASIS) ^ 0.5) '(365) ^ 0.5
    TEMP_MATRIX(j, 17) = ZWEIG_VAL
1983:
'-------------------------------------------------------------------------------------
Next j
'-------------------------------------------------------------------------------------

ASSET_TA_TABLE2_FUNC = TEMP_MATRIX

Exit Function
ERROR_LABEL:
ASSET_TA_TABLE2_FUNC = Err.number
End Function

Function ASSET_TA_TABLE3_FUNC(ByRef TICKERS_RNG As Variant, _
Optional ByVal INDEX_SYMBOL As String = "^GSPTSE", _
Optional ByVal START_DATE As Date = 0, _
Optional ByVal END_DATE As Date = 0, _
Optional ByVal CASH_RATE As Double = 0.04, _
Optional ByVal MA1_PERIOD As Long = 20, _
Optional ByVal MA2_PERIOD As Long = 50, _
Optional ByVal MA3_PERIOD As Long = 100, _
Optional ByVal COUNT1_BASIS As Double = 365, _
Optional ByVal COUNT2_BASIS As Double = 252)

Dim i As Long
Dim j As Long

Dim NROWS As Long
Dim NSIZE As Long

Dim MEAN_VAL As Double
Dim VOLAT_VAL As Double

Dim DATA1_VECTOR As Variant
Dim DATA2_VECTOR As Variant
Dim DATA3_VECTOR As Variant

Dim TEMP_MATRIX As Variant
Dim DATA_MATRIX As Variant
Dim MARKET_MATRIX As Variant
Dim TICKERS_VECTOR As Variant

On Error Resume Next

If IsArray(TICKERS_RNG) = True Then
    TICKERS_VECTOR = TICKERS_RNG
    If UBound(TICKERS_VECTOR, 1) = 1 Then
        TICKERS_VECTOR = MATRIX_TRANSPOSE_FUNC(TICKERS_VECTOR)
    End If
Else
    ReDim TICKERS_VECTOR(1 To 1, 1 To 1)
    TICKERS_VECTOR(1, 1) = TICKERS_RNG
End If

NSIZE = UBound(TICKERS_VECTOR, 1)

If END_DATE = 0 Then
    END_DATE = Now
    END_DATE = DateSerial(Year(END_DATE), Month(END_DATE), Day(END_DATE))
End If

If START_DATE = 0 Then
    START_DATE = DateSerial(Year(END_DATE) - 1, Month(END_DATE), Day(END_DATE))
End If

MARKET_MATRIX = YAHOO_HISTORICAL_DATA_SERIE_FUNC(Trim(INDEX_SYMBOL), _
                START_DATE, END_DATE, "d", "DOHLCVA", False, True, True)
If IsArray(MARKET_MATRIX) = False Then: GoTo ERROR_LABEL

ReDim TEMP_MATRIX(0 To NSIZE, 1 To 12)
TEMP_MATRIX(0, 1) = "SYMBOL"
TEMP_MATRIX(0, 2) = "MA_" & MA1_PERIOD & "_DAYS" 'Days in Moving Average
TEMP_MATRIX(0, 3) = "MA_" & MA2_PERIOD & "_DAYS" 'Days in Moving Average
TEMP_MATRIX(0, 4) = "MA_" & MA3_PERIOD & "_DAYS" 'Days in Moving Average
TEMP_MATRIX(0, 5) = "CLOSING_PRICE"
TEMP_MATRIX(0, 6) = "CAGR"
TEMP_MATRIX(0, 7) = "ANNUAL_VOLATILITY"
TEMP_MATRIX(0, 8) = "VOLUME/1000"
TEMP_MATRIX(0, 9) = "SERIAL_CORRELATION"
TEMP_MATRIX(0, 10) = "AVG_DAILY_RETURN"
TEMP_MATRIX(0, 11) = "MARKET_BETA"
TEMP_MATRIX(0, 12) = "SHARPE_RATIO"

'-------------------------------------------------------------------------------------
For j = 1 To NSIZE
'-------------------------------------------------------------------------------------
    TEMP_MATRIX(j, 1) = Trim(TICKERS_VECTOR(j, 1))
    If TEMP_MATRIX(j, 1) = "" Then: GoTo 1983
    DATA_MATRIX = YAHOO_HISTORICAL_DATA_SERIE_FUNC(TEMP_MATRIX(j, 1), _
                  START_DATE, END_DATE, "d", "DOHLCVA", False, True, True)
    If IsArray(DATA_MATRIX) = False Then: GoTo 1983
    
    NROWS = UBound(DATA_MATRIX, 1) 'Excluding Headings
'    If NROWS <> UBound(MARKET_MATRIX, 1) Then: GoTo 1983
    
    ReDim DATA1_VECTOR(1 To NROWS - 1, 1 To 1)
    ReDim DATA2_VECTOR(1 To NROWS - 1, 1 To 1)
    ReDim DATA3_VECTOR(1 To NROWS - 1, 1 To 1)
    TEMP_MATRIX(j, 2) = 0
    TEMP_MATRIX(j, 3) = 0
    TEMP_MATRIX(j, 4) = 0
    For i = 1 To NROWS
        If i >= NROWS - MA1_PERIOD Then: TEMP_MATRIX(j, 2) = TEMP_MATRIX(j, 2) + DATA_MATRIX(i, 7)
        If i >= NROWS - MA2_PERIOD Then: TEMP_MATRIX(j, 3) = TEMP_MATRIX(j, 3) + DATA_MATRIX(i, 7)
        If i >= NROWS - MA3_PERIOD Then: TEMP_MATRIX(j, 4) = TEMP_MATRIX(j, 4) + DATA_MATRIX(i, 7)
        If i > 1 Then
            MEAN_VAL = MEAN_VAL + DATA_MATRIX(i, 7) / DATA_MATRIX(i - 1, 7) - 1
            DATA1_VECTOR(i - 1, 1) = DATA_MATRIX(i, 7) / DATA_MATRIX(i - 1, 7) - 1
            If i > 2 Then
                DATA2_VECTOR(i - 1, 1) = DATA1_VECTOR(i - 2, 1)
            ElseIf i = 2 Then
                DATA2_VECTOR(i - 1, 1) = DATA_MATRIX(i - 1, 5) / DATA_MATRIX(i - 1, 2) - 1
            End If
            DATA3_VECTOR(i - 1, 1) = MARKET_MATRIX(i, 7) / MARKET_MATRIX(i - 1, 7) - 1
        Else
            MEAN_VAL = DATA_MATRIX(i, 5) / DATA_MATRIX(i, 2) - 1
        End If
    Next i
    MEAN_VAL = MEAN_VAL / NROWS
    For i = 1 To NROWS
        If i > 1 Then
            VOLAT_VAL = VOLAT_VAL + ((DATA_MATRIX(i, 7) / DATA_MATRIX(i - 1, 7) - 1) - MEAN_VAL) ^ 2
        Else
            VOLAT_VAL = VOLAT_VAL + ((DATA_MATRIX(i, 5) / DATA_MATRIX(i, 2) - 1) - MEAN_VAL) ^ 2
        End If
    Next i
    VOLAT_VAL = (VOLAT_VAL / NROWS) ^ 0.5
    VOLAT_VAL = VOLAT_VAL * COUNT2_BASIS ^ 0.5
    TEMP_MATRIX(j, 2) = TEMP_MATRIX(j, 2) / (MA1_PERIOD + 1)
    TEMP_MATRIX(j, 3) = TEMP_MATRIX(j, 3) / (MA2_PERIOD + 1)
    TEMP_MATRIX(j, 4) = TEMP_MATRIX(j, 4) / (MA3_PERIOD + 1)
    TEMP_MATRIX(j, 5) = DATA_MATRIX(NROWS, 7)
    TEMP_MATRIX(j, 6) = (DATA_MATRIX(NROWS, 7) / DATA_MATRIX(1, 7)) ^ ((DATA_MATRIX(NROWS, 1) - DATA_MATRIX(1, 1)) / COUNT1_BASIS) - 1
    TEMP_MATRIX(j, 7) = VOLAT_VAL
    TEMP_MATRIX(j, 8) = DATA_MATRIX(NROWS, 6) / 1000
    TEMP_MATRIX(j, 9) = CORRELATION_FUNC(DATA1_VECTOR, DATA2_VECTOR, 0, 0)
    TEMP_MATRIX(j, 10) = MEAN_VAL
    TEMP_MATRIX(j, 11) = REGRESSION_SIMPLE_COEF_FUNC(DATA3_VECTOR, DATA1_VECTOR)(1, 1) 'INDEX_VECTOR, DATA_VECTOR)
    TEMP_MATRIX(j, 12) = (MEAN_VAL * COUNT2_BASIS - CASH_RATE) / VOLAT_VAL
1983:
'-------------------------------------------------------------------------------------
Next j
'-------------------------------------------------------------------------------------

ASSET_TA_TABLE3_FUNC = TEMP_MATRIX

Exit Function
ERROR_LABEL:
ASSET_TA_TABLE3_FUNC = Err.number
End Function
