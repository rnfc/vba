Attribute VB_Name = "FINAN_DERIV_DASPP_LIBR"


Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.

Private Const PUB_EPSILON As Double = 2 ^ 52
Private PUB_PERIODS As Double 'trading days per year
Private PUB_CONSTANTS_VECTOR As Variant
Private PUB_DPE_REWEIGHTING_PERIODS As Double

'************************************************************************************
'************************************************************************************
'FUNCTION      : DASPP_SIMULATION_FUNC

'DESCRIPTION   : 'Dynamic Allocation Strategies with Principal Protection (for Prop
'Trading and Hedge Funds). This routine allows to conduct backtest for the CPPI, OBPI
'or DPE (using short/put option since it is a downtrending market). The underlying
'risky asset (can be equity or fixed income instruments) are simulated via general
'innovation processes with specification of stochastic volatility.

'LIBRARY       : PORTFOLIO
'GROUP         : SIMULATION
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 04/07/2011
'************************************************************************************
'************************************************************************************

'-----------------------------------------------------------------------------------------
'References:
'-----------------------------------------------------------------------------------------
'http://zhenwei.wordpress.com/page/2/
'http://www.weizhenstanford.com/
'http://www.barings.com/uk/InstitutionalInvestors/index.htm
'http://www.nobletrading.com/blogs/2008/06/what-is-dynamic-asset-allocation.html
'-----------------------------------------------------------------------------------------

Function DASPP_SIMULATION_FUNC(Optional ByVal INITIAL_WEALTH As Double = 100, _
Optional ByVal INITIAL_RISKY_ASSET As Double = 1, _
Optional ByVal PROTECTION_LEVEL As Double = 90, _
Optional ByVal TIME_TO_MATURITY As Double = 1, _
Optional ByVal CASH_RATE As Double = 0.035, _
Optional ByVal TRANSACTION_COST As Double = 0, _
Optional ByVal RANDOM_TYPE As Integer = 1, Optional ByVal MEAN_RETURN As Double = 0.2, _
Optional ByVal SKEWNESS_VAL As Double = -1, Optional ByVal KURTOSIS_VAL As Double = 6, _
Optional ByVal VOLATILITY_METHOD As Integer = 2, Optional ByVal INITIAL_VOLATILITY As Double = 0.3, _
Optional ByVal LONG_TERM_VOLATILITY As Double = 0.3, Optional ByVal VOLATILITY_MEAN_REVERSION As Double = 0.5, _
Optional ByVal VOLATILITY_OF_VOLATILITY As Double = 0.4, Optional ByVal VOLATILITY_RHO As Double = 0, _
Optional ByVal MIN_RISKY_EXPOSURE As Double = 0, Optional ByVal MAX_RISKY_EXPOSURE As Double = 1, _
Optional ByVal REWEIGHTING_TRIGGER As Double = 0, Optional ByVal DPAA_MULTIPLIER As Double = 6, _
Optional ByVal TARGET_RETURN As Double = 0.3, Optional ByVal CPPI_MULTIPLIER As Double = 6, _
Optional ByVal DPE_REWEIGHTING_PERIODS As Double = 30, Optional ByVal DPE_THETA As Double = 1, _
Optional ByVal DPE_MULTIPLIER As Double = 6, Optional ByVal DPE_REWEIGHTING_TRIGGER As Double = 0, _
Optional ByVal PERIODS_PER_YEAR As Double = 250, _
Optional ByVal CND_TYPE As Integer = 0, _
Optional ByVal nLOOPS As Long = 100, _
Optional ByVal OUTPUT As Integer = 3)

Dim h As Long
Dim i As Long
Dim j As Long
Dim k As Long
Dim l As Long
Dim ii As Long
Dim jj As Long

'Dim START_TIME As Date
'Dim END_TIME As Date

Dim HEADINGS_ARR As Variant
Dim DATA_MATRIX As Variant
Dim TEMP_MATRIX As Variant

On Error GoTo ERROR_LABEL

'-------------------------------------------------------------------------------------------------------
'-------------------------------------------------------------------------------------------------------
'Return/Largest DD/Longest DD/Largest Loss/TurnOver/Average/Vol/Skewness/
'90% VaR/99% VaR/Sharp/Sortino/Upside-P/Win/Loss/Risky/CPPI/OBPI/DPE/DPAA
'-------------------------------------------------------------------------------------------------------
'-------------------------------------------------------------------------------------------------------

HEADINGS_ARR = Array( _
    "RISKY: RETURN", "CPPI: RETURN", "OBPI: RETURN", "DPE: RETURN", "DPAA: RETURN", "RISKY: LARGEST DD", _
    "CPPI: LARGEST DD", "OBPI: LARGEST DD", "DPE: LARGEST DD", "DPAA: LARGEST DD", "RISKY: LONGEST DD", _
    "CPPI: LONGEST DD", "OBPI: LONGEST DD", "DPE: LONGEST DD", "DPAA: LONGEST DD", "RISKY: LARGEST LOSS", _
    "CPPI: LARGEST LOSS", "OBPI: LARGEST LOSS", "DPE: LARGEST LOSS", "DPAA: LARGEST LOSS", "RISKY: TURNOVER", _
    "CPPI: TURNOVER", "OBPI: TURNOVER", "DPE: TURNOVER", "DPAA: TURNOVER", "RISKY: AVERAGE", "CPPI: AVERAGE", _
    "OBPI: AVERAGE", "DPE: AVERAGE", "DPAA: AVERAGE", "RISKY: VOLATILITY", "CPPI: VOLATILITY", "OBPI: VOLATILITY", _
    "DPE: VOLATILITY", "DPAA: VOLATILITY", "RISKY: SKEWNESS", "CPPI: SKEWNESS", "OBPI: SKEWNESS", "DPE: SKEWNESS", _
    "DPAA: SKEWNESS", "RISKY: 90% VAR", "CPPI: 90% VAR", "OBPI: 90% VAR", "DPE: 90% VAR", "DPAA: 90% VAR", _
    "RISKY: 99% VAR", "CPPI: 99% VAR", "OBPI: 99% VAR", "DPE: 99% VAR", "DPAA: 99% VAR", "RISKY: SHARP", _
    "CPPI: SHARP", "OBPI: SHARP", "DPE: SHARP", "DPAA: SHARP", "RISKY: SORTINO", "CPPI: SORTINO", "OBPI: SORTINO", _
    "DPE: SORTINO", "DPAA: SORTINO", "RISKY: UPSIDE-P", "CPPI: UPSIDE-P", "OBPI: UPSIDE-P", "DPE: UPSIDE-P", _
    "DPAA: UPSIDE-P", "RISKY: WIN", "CPPI: WIN", "OBPI: WIN", "DPE: WIN", "DPAA: WIN", _
    "RISKY: LOSS", "CPPI: LOSS", "OBPI: LOSS", "DPE: LOSS", "DPAA: LOSS", _
    "RISKY: RHO - RISKY", "CPPI: RHO - RISKY", "OBPI: RHO - RISKY", "DPE: RHO - RISKY", "DPAA: RHO - RISKY", _
    "RISKY: RHO - CPPI", "CPPI: RHO - CPPI", "OBPI: RHO - CPPI", "DPE: RHO - CPPI", "DPAA: RHO - CPPI", _
    "RISKY: RHO -OBPI", "CPPI: RHO -OBPI", "OBPI: RHO -OBPI", "DPE: RHO -OBPI", "DPAA: RHO -OBPI", _
    "RISKY: RHO -DPE", "CPPI: RHO -DPE", "OBPI: RHO -DPE", "DPE: RHO -DPE", "DPAA: RHO -DPE", _
    "RISKY: RHO -DPAA", "CPPI: RHO -DPAA", "OBPI: RHO -DPAA", "DPE: RHO -DPAA", "DPAA: RHO -DPAA")

'-----------------------------------------------------------------------------------------------------------
ReDim TEMP_MATRIX(0 To nLOOPS, 1 To 100)
'-----------------------------------------------------------------------------------------------------------
If OUTPUT = 0 Then 'Excel Summary Framework
'-----------------------------------------------------------------------------------------------------------
    jj = 1
    For j = 1 To 5 'Headings
        i = j
        For k = 0 To 19
            TEMP_MATRIX(0, jj) = HEADINGS_ARR(k + i)
            jj = jj + 1
            i = i + 4
        Next k
    Next j
    For ii = 1 To nLOOPS
        GoSub RUN_TIME_LINE
        l = 0
        For j = 1 To 4
            If j = 1 Then 'basicAnchor
                h = 0
            ElseIf j = 2 Then 'returnAnchor
                h = 12
            ElseIf j = 3 Then 'ratioAnchor
                h = 18
            Else 'corrAnchor
                h = 30
            End If
            For k = 1 To 5
                jj = l
                For i = 1 To 5
                    TEMP_MATRIX(ii, jj + 1) = DATA_MATRIX(h + i + 1, k + 1)
                    jj = jj + 20
                Next i
                l = l + 1
            Next k
        Next j
    Next ii
    DASPP_SIMULATION_FUNC = TEMP_MATRIX
'-----------------------------------------------------------------------------------------------------------
Else
'-----------------------------------------------------------------------------------------------------------
    'START_TIME = Now
    For j = 1 To 100: TEMP_MATRIX(0, j) = HEADINGS_ARR(j): Next j
    For ii = 1 To nLOOPS
        GoSub RUN_TIME_LINE
        jj = 1
        For j = 1 To 4
            If j = 1 Then 'basicAnchor
                h = 0
            ElseIf j = 2 Then 'returnAnchor
                h = 12
            ElseIf j = 3 Then 'ratioAnchor
                h = 18
            Else 'corrAnchor
                h = 30
            End If
            For k = 1 To 5 'Risky/CPPI/OBPI/DPE/ DPAA
                For i = 1 To 5
                    TEMP_MATRIX(ii, jj) = DATA_MATRIX(h + i + 1, k + 1)
                    jj = jj + 1
                Next i
            Next k
        Next j
        'Application.StatusBar = "Simulation Progress: " & Round(ii / nLOOPS * 100, 1) & "%"
    Next ii
    Erase DATA_MATRIX
'END_TIME = Now
    If OUTPUT = 1 Then
        DASPP_SIMULATION_FUNC = TEMP_MATRIX
    Else
        ReDim DATA_MATRIX(1 To 2)
        DATA_MATRIX(1) = DATA_BASIC_MOMENTS_FUNC(TEMP_MATRIX, 0, 0, 0.05, 1)
        For j = 1 To 100: DATA_MATRIX(1)(j + 1, 1) = HEADINGS_ARR(j): Next j
        If OUTPUT = 2 Then
            DASPP_SIMULATION_FUNC = DATA_MATRIX(1)
        Else
            DATA_MATRIX(2) = DATA_ADVANCED_MOMENTS_FUNC(TEMP_MATRIX, 0, 0, 0.05, 0)
            DATA_MATRIX(2) = MATRIX_ADD_COLUMNS_FUNC(DATA_MATRIX(2), 1, 1)
            DATA_MATRIX(2)(1, 1) = "OBV(s)"
            For j = 1 To 100: DATA_MATRIX(2)(j + 1, 1) = HEADINGS_ARR(j): Next j
            If OUTPUT = 3 Then
                DASPP_SIMULATION_FUNC = DATA_MATRIX(2)
            Else
                DASPP_SIMULATION_FUNC = DATA_MATRIX
            End If
        End If
    End If
'-----------------------------------------------------------------------------------------------------------
End If
'-----------------------------------------------------------------------------------------------------------
Exit Function
'----------------------------------------------------------------------------------------------------
RUN_TIME_LINE:
'----------------------------------------------------------------------------------------------------
    DATA_MATRIX = _
        DASPP_SAMPLING_FUNC( _
            INITIAL_WEALTH, _
            INITIAL_RISKY_ASSET, _
            PROTECTION_LEVEL, _
            TIME_TO_MATURITY, _
            CASH_RATE, _
            TRANSACTION_COST, _
            RANDOM_TYPE, _
            MEAN_RETURN, _
            SKEWNESS_VAL, _
            KURTOSIS_VAL, _
            VOLATILITY_METHOD, _
            INITIAL_VOLATILITY, _
            LONG_TERM_VOLATILITY, VOLATILITY_MEAN_REVERSION, _
            VOLATILITY_OF_VOLATILITY, VOLATILITY_RHO, _
            MIN_RISKY_EXPOSURE, MAX_RISKY_EXPOSURE, _
            REWEIGHTING_TRIGGER, DPAA_MULTIPLIER, _
            TARGET_RETURN, CPPI_MULTIPLIER, _
            DPE_REWEIGHTING_PERIODS, DPE_THETA, _
            DPE_MULTIPLIER, DPE_REWEIGHTING_TRIGGER, _
            PERIODS_PER_YEAR, CND_TYPE, False, 2)
'----------------------------------------------------------------------------------------------------
Return
'----------------------------------------------------------------------------------------------------
ERROR_LABEL:
DASPP_SIMULATION_FUNC = Err.number
End Function


Function DASPP_SAMPLING_FUNC( _
Optional ByVal INITIAL_WEALTH As Double = 100, _
Optional ByVal INITIAL_RISKY_ASSET As Double = 1, _
Optional ByVal PROTECTION_LEVEL As Double = 90, _
Optional ByVal TIME_TO_MATURITY As Double = 1, _
Optional ByVal CASH_RATE As Double = 0.035, _
Optional ByVal TRANSACTION_COST As Double = 0, _
Optional ByVal RANDOM_TYPE As Integer = 1, Optional ByVal MEAN_RETURN As Double = 0.2, _
Optional ByVal SKEWNESS_VAL As Double = -1, Optional ByVal KURTOSIS_VAL As Double = 6, _
Optional ByVal VOLATILITY_METHOD As Integer = 2, Optional ByVal INITIAL_VOLATILITY As Double = 0.3, _
Optional ByVal LONG_TERM_VOLATILITY As Double = 0.3, Optional ByVal VOLATILITY_MEAN_REVERSION As Double = 0.5, _
Optional ByVal VOLATILITY_OF_VOLATILITY As Double = 0.4, Optional ByVal VOLATILITY_RHO As Double = 0, _
Optional ByVal MIN_RISKY_EXPOSURE As Double = 0, Optional ByVal MAX_RISKY_EXPOSURE As Double = 1, _
Optional ByVal REWEIGHTING_TRIGGER As Double = 0, Optional ByVal DPAA_MULTIPLIER As Double = 6, _
Optional ByVal TARGET_RETURN As Double = 0.3, Optional ByVal CPPI_MULTIPLIER As Double = 6, _
Optional ByVal DPE_REWEIGHTING_PERIODS As Double = 30, Optional ByVal DPE_THETA As Double = 1, _
Optional ByVal DPE_MULTIPLIER As Double = 6, Optional ByVal DPE_REWEIGHTING_TRIGGER As Double = 0, _
Optional ByVal PERIODS_PER_YEAR As Double = 250, _
Optional ByVal CND_TYPE As Integer = 0, _
Optional ByVal HEADERS_FLAG As Boolean = True, _
Optional ByVal OUTPUT As Integer = 2)

'-----------------------------------------------------------------------------------------------
'Dynamic trading strategies Sampling & Simulation Routines: CPPI (Constant
'Proportional Portfolio Insurance), OBPI (Option Based Portfolio Insurance),
'DPE (Dynamic Protected Envelope) and DPAA (Dynamic Protected Asset Allocation).
'-----------------------------------------------------------------------------------------------
'Deal Parameters and Simulation Model for Risky Portfolio
'Specified by: r(t) = u(t) dt + sigma(t) dW(t)
'where r(t) is the Instantaneous Return,
'dW(t) is the Innovation,
'sigma(t) is the Volatility.
'-----------------------------------------------------------------------------------------------

'-----------------------------------------------------------------------------------------------
'DEAL PARAMETERS:
'-----------------------------------------------------------------------------------------------
'INITIAL_WEALTH: in XYZ Currency
'INITIAL_RISKY_ASSET: per Contract in XYZ Currency
'PROTECTION_LEVEL: at maturity in XYZ Currency
'TIME_TO_MATURITY: in Years
'CASH_RATE: Risk-Free Return per annum. Assume it is also the Mininum Acceptable Return.
'TRANSACTION_COST: per turnover in XYZ Currency
'-----------------------------------------------------------------------------------------------
'SIMULATION PARAMETERS:
'-----------------------------------------------------------------------------------------------
'RANDOM_TYPE: Innovation Model - 0 for Standard Normal or 1(else) for Normal Inv-Gaussian
'MEAN_RETURN: per annum
'INITIAL_VOLATILITY: per annum
'-----------------------------------------------------------------------------------------------
'MODEL PARAMETERS:
'-----------------------------------------------------------------------------------------------
'VOLATILITY_MODEL: 0 Constant / 1 Heston / 2 Garch
'LONG_TERM_VOLATILITY
'VOLATILITY_MEAN_REVERSION --> B1
'VOLATILITY_OF_VOLATILITY --> B2
'VOLATILITY_RHO --> Only applies for the Heston Model --> Adjust the normal random
'numbers (bivar distribution)
'-----------------------------------------------------------------------------------------------
'TRADE PARAMETERS:
'-----------------------------------------------------------------------------------------------
'MIN_RISKY_EXPOSURE(%)
'MAX_RISKY_EXPOSURE(%)
'REWEIGHTING_TRIGGER(%): Adjust the Risky Asset Weight if the Weight Changes more than the
'Trigger Level.
'DPAA_MULTIPLIER
'TARGET_RETURN(%)
'CPPI_MULTIPLIER
'-----------------------------------------------------------------------------------------------
'DYNAMIC PROTECTED ENVELOPE PARAMETERS:
'-----------------------------------------------------------------------------------------------
'DPE_REWEIGHTING_PERIODS - Reweighting PERIODS
'DPE_THETA - DPE (Dynamic Protection Envelope) use N periods of re-weighting
'Strategies to save Transaction Cost.
'DPE_MULTIPLIER - Multiplier
'DPE_REWEIGHTING_TRIGGER(%) - Reweighting Trigger
'-----------------------------------------------------------------------------------------------
'COUNT BASIS PARAMETERS:
'-----------------------------------------------------------------------------------------------
'PERIODS_PER_YEAR --> Assuming there are 250 Trading Periods in a Year.
'-----------------------------------------------------------------------------------------------

Dim h As Long
Dim i As Long
Dim j As Long

Dim PERIODS As Long 'TRADING DAYS FOR THE TRIAL

Dim TEMP_VAL As Double
Dim TEMP1_SUM As Double
Dim TEMP2_SUM As Double

Dim DPAA_THETA As Double
Dim DPAA_LAMBDA As Double
Dim DPAA_MU As Double

Dim DELTA_TIME As Double
Dim OBPI_SCALE As Double
Dim DPE_FIRST As Double

Dim RANDOM_VECTOR As Variant
Dim SAMPLING_MATRIX As Variant
Dim CHARTING_MATRIX As Variant 'Wealth Processes of Dynamic Trading Strategies
Dim STATISTICS_MATRIX As Variant 'Summary Statistics for Dynamic Trading Strategies

Dim tolerance As Double

On Error GoTo ERROR_LABEL

If HEADERS_FLAG = True Then h = 0 Else h = 1

tolerance = 10 ^ -15
PERIODS = FLOOR_FUNC(PERIODS_PER_YEAR * TIME_TO_MATURITY, 1)
'Total Trading days

DPAA_THETA = DASPP_DPAA_THETA_SAMPLING_FUNC(TIME_TO_MATURITY * 2, CASH_RATE, MEAN_RETURN, VOLATILITY_METHOD, INITIAL_VOLATILITY, LONG_TERM_VOLATILITY, VOLATILITY_MEAN_REVERSION, VOLATILITY_OF_VOLATILITY, VOLATILITY_RHO, RANDOM_TYPE, SKEWNESS_VAL, KURTOSIS_VAL, PERIODS_PER_YEAR)
SAMPLING_MATRIX = DASPP_DPAA_DMV_OPTIMIZER_FUNC(INITIAL_WEALTH, PROTECTION_LEVEL, TIME_TO_MATURITY, CASH_RATE, MEAN_RETURN, DPAA_THETA, 0, 0, CND_TYPE)
DPAA_LAMBDA = SAMPLING_MATRIX(1, 1) '0.609540226833722
DPAA_MU = SAMPLING_MATRIX(2, 1) '0.418822992306894
Erase SAMPLING_MATRIX

DELTA_TIME = 1 / PERIODS_PER_YEAR
RANDOM_VECTOR = DASPP_RANDOM_VECTOR_FUNC(RANDOM_TYPE, PERIODS + 1, SKEWNESS_VAL, KURTOSIS_VAL, 3)

'-------------------------------------------------------------------------------------------------------
'-------------------------------------------------------------------------------------------------------
GoSub HEADERS_LINE
If OUTPUT > 0 Then: GoSub SUMMARY_LINE
'-----------------------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------------------
j = 1
SAMPLING_MATRIX(j, 1) = 0
SAMPLING_MATRIX(j, 2) = RANDOM_VECTOR(1, 1)
SAMPLING_MATRIX(j, 3) = INITIAL_VOLATILITY
SAMPLING_MATRIX(j, 4) = 0
SAMPLING_MATRIX(j, 5) = INITIAL_WEALTH
SAMPLING_MATRIX(j, 6) = PROTECTION_LEVEL / (1 + CASH_RATE * DELTA_TIME) ^ PERIODS_PER_YEAR
'-----------------------------------------------------------------------------------------------

SAMPLING_MATRIX(j, 11) = INITIAL_WEALTH
SAMPLING_MATRIX(j, 9) = 1 '100%

TEMP_VAL = SAMPLING_MATRIX(j, 11) - SAMPLING_MATRIX(j, 6)
If TEMP_VAL > tolerance Then
    SAMPLING_MATRIX(j, 7) = TEMP_VAL
Else
    SAMPLING_MATRIX(j, 7) = 0
End If

TEMP_VAL = SAMPLING_MATRIX(j, 7) * CPPI_MULTIPLIER / INITIAL_WEALTH
If MAX_RISKY_EXPOSURE < TEMP_VAL Then: TEMP_VAL = MAX_RISKY_EXPOSURE
If MIN_RISKY_EXPOSURE > TEMP_VAL Then: TEMP_VAL = MIN_RISKY_EXPOSURE
SAMPLING_MATRIX(j, 8) = TEMP_VAL
SAMPLING_MATRIX(j, 10) = SAMPLING_MATRIX(j, 8)

'------------------------CPPI (Constant Proportional Portfolio Insurance)-------------------------

OBPI_SCALE = DASPP_OBPI_SCALE_FUNC(INITIAL_WEALTH, PROTECTION_LEVEL, TIME_TO_MATURITY, CASH_RATE, INITIAL_VOLATILITY)
TEMP_VAL = DASPP_OBPI_EXPO_FUNC(SAMPLING_MATRIX(j, 5), OBPI_SCALE, SAMPLING_MATRIX(j, 5), PROTECTION_LEVEL, (PERIODS - SAMPLING_MATRIX(j, 1)) / PERIODS_PER_YEAR, CASH_RATE, INITIAL_VOLATILITY, CND_TYPE)
If MAX_RISKY_EXPOSURE < TEMP_VAL Then: TEMP_VAL = MAX_RISKY_EXPOSURE
If MIN_RISKY_EXPOSURE > TEMP_VAL Then: TEMP_VAL = MIN_RISKY_EXPOSURE

SAMPLING_MATRIX(j, 12) = TEMP_VAL
SAMPLING_MATRIX(j, 13) = 1 '100%
SAMPLING_MATRIX(j, 14) = SAMPLING_MATRIX(j, 12)
SAMPLING_MATRIX(j, 15) = INITIAL_WEALTH

'------------------------------DPE (Dynamic Protected Envelope)-----------------------------------
SAMPLING_MATRIX(j, 16) = 1
If DPE_THETA = 1 Then
    DPE_FIRST = PERIODS / DPE_REWEIGHTING_PERIODS
Else
    DPE_FIRST = Round(PERIODS * (DPE_THETA - 1) / (DPE_THETA ^ DPE_REWEIGHTING_PERIODS - 1), 0)
End If

SAMPLING_MATRIX(j, 19) = 1
SAMPLING_MATRIX(j, 21) = INITIAL_WEALTH

TEMP_VAL = SAMPLING_MATRIX(j, 21) - SAMPLING_MATRIX(j, 6)
If TEMP_VAL > tolerance Then
    SAMPLING_MATRIX(j, 17) = TEMP_VAL
Else
    SAMPLING_MATRIX(j, 17) = 0
End If

TEMP_VAL = SAMPLING_MATRIX(j, 17) * CPPI_MULTIPLIER / INITIAL_WEALTH
If MAX_RISKY_EXPOSURE < TEMP_VAL Then: TEMP_VAL = MAX_RISKY_EXPOSURE
If MIN_RISKY_EXPOSURE > TEMP_VAL Then: TEMP_VAL = MIN_RISKY_EXPOSURE
SAMPLING_MATRIX(j, 18) = TEMP_VAL
SAMPLING_MATRIX(j, 20) = SAMPLING_MATRIX(j, 18)

'---------------------------------DPAA (Dynamic Protected Asset Allocation)-----------------------
SAMPLING_MATRIX(j, 22) = DPAA_THETA
SAMPLING_MATRIX(j, 23) = SAMPLING_MATRIX(j, 22) ^ 2
TEMP1_SUM = SAMPLING_MATRIX(j, 23)

SAMPLING_MATRIX(j, 24) = SAMPLING_MATRIX(j, 22) * SAMPLING_MATRIX(j, 2)
TEMP2_SUM = SAMPLING_MATRIX(j, 24)

SAMPLING_MATRIX(j, 25) = DPAA_MU * Exp(-(2 * CASH_RATE - DPAA_THETA ^ 2) * TIME_TO_MATURITY)

SAMPLING_MATRIX(j, 26) = (Log(SAMPLING_MATRIX(j, 25) / DPAA_LAMBDA) + (CASH_RATE + 1 / 2 * DPAA_THETA ^ 2) * (PERIODS - SAMPLING_MATRIX(j, 1)) * DELTA_TIME) / Sqr((PERIODS - SAMPLING_MATRIX(j, 1)) * DELTA_TIME) / DPAA_THETA

SAMPLING_MATRIX(j, 27) = SAMPLING_MATRIX(j, 26) - Sqr((PERIODS - SAMPLING_MATRIX(j, 1)) * DELTA_TIME) * DPAA_THETA

SAMPLING_MATRIX(j, 28) = DPAA_LAMBDA * CND_FUNC(-SAMPLING_MATRIX(j, 27), CND_TYPE) * Exp(-CASH_RATE * (PERIODS - SAMPLING_MATRIX(j, 1)) * DELTA_TIME) - CND_FUNC(-SAMPLING_MATRIX(j, 26), CND_TYPE) * SAMPLING_MATRIX(j, 25)

TEMP_VAL = -(SAMPLING_MATRIX(j, 28) - DPAA_LAMBDA * CND_FUNC(-SAMPLING_MATRIX(j, 27), CND_TYPE) * Exp(-(PERIODS - SAMPLING_MATRIX(j, 1)) * DELTA_TIME * CASH_RATE)) * SAMPLING_MATRIX(j, 22) / SAMPLING_MATRIX(j, 3) * (INITIAL_WEALTH - SAMPLING_MATRIX(1, 6)) * DPAA_MULTIPLIER / SAMPLING_MATRIX(j, 5)

If MAX_RISKY_EXPOSURE < TEMP_VAL Then: TEMP_VAL = MAX_RISKY_EXPOSURE
If MIN_RISKY_EXPOSURE > TEMP_VAL Then: TEMP_VAL = MIN_RISKY_EXPOSURE
    
SAMPLING_MATRIX(j, 29) = TEMP_VAL
SAMPLING_MATRIX(j, 30) = 1 '100%
SAMPLING_MATRIX(j, 31) = SAMPLING_MATRIX(j, 29)
SAMPLING_MATRIX(j, 32) = INITIAL_WEALTH
If OUTPUT > 0 Then: GoSub SUMMARY_LINE

'-------------------------------------------------------------------------------------------------
For i = 1 To PERIODS
'-----------------------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------------------
    j = i + 1
'-----------------------------------------------------------------------------------------------
    
    SAMPLING_MATRIX(j, 1) = i
    SAMPLING_MATRIX(j, 2) = RANDOM_VECTOR(j, 1)
    SAMPLING_MATRIX(j, 3) = DASPP_NEXT_VOLATILITY_FUNC(SAMPLING_MATRIX(i, 2), SAMPLING_MATRIX(i, 3), DELTA_TIME, LONG_TERM_VOLATILITY, VOLATILITY_MEAN_REVERSION, VOLATILITY_OF_VOLATILITY, VOLATILITY_RHO, 0, VOLATILITY_METHOD)
    SAMPLING_MATRIX(j, 4) = MEAN_RETURN * DELTA_TIME + SAMPLING_MATRIX(j, 3) * SAMPLING_MATRIX(j, 2) * Sqr(DELTA_TIME)
    SAMPLING_MATRIX(j, 5) = SAMPLING_MATRIX(i, 5) * (1 + SAMPLING_MATRIX(j, 4))
    SAMPLING_MATRIX(j, 6) = SAMPLING_MATRIX(i, 6) * (1 + CASH_RATE * DELTA_TIME)
    
'---------------------------CPPI (Constant Proportional Portfolio Insurance)------------------------
    SAMPLING_MATRIX(j, 11) = SAMPLING_MATRIX(i, 11) * (1 + SAMPLING_MATRIX(i, 10) * SAMPLING_MATRIX(j, 4) + (1 - SAMPLING_MATRIX(i, 10)) * CASH_RATE * DELTA_TIME) - INITIAL_WEALTH / INITIAL_RISKY_ASSET * Abs(SAMPLING_MATRIX(i, 9)) * TRANSACTION_COST
    TEMP_VAL = SAMPLING_MATRIX(j, 11) - SAMPLING_MATRIX(j, 6)
    If TEMP_VAL > tolerance Then
        SAMPLING_MATRIX(j, 7) = TEMP_VAL
    Else
        SAMPLING_MATRIX(j, 7) = 0
    End If
    TEMP_VAL = SAMPLING_MATRIX(j, 7) * CPPI_MULTIPLIER / INITIAL_WEALTH
    If MAX_RISKY_EXPOSURE < TEMP_VAL Then: TEMP_VAL = MAX_RISKY_EXPOSURE
    If MIN_RISKY_EXPOSURE > TEMP_VAL Then: TEMP_VAL = MIN_RISKY_EXPOSURE
    SAMPLING_MATRIX(j, 8) = TEMP_VAL

    TEMP_VAL = SAMPLING_MATRIX(j, 8) - SAMPLING_MATRIX(i, 10)
    If Abs(TEMP_VAL) > REWEIGHTING_TRIGGER Then
        SAMPLING_MATRIX(j, 9) = TEMP_VAL
    Else
        SAMPLING_MATRIX(j, 9) = 0
    End If
    
    SAMPLING_MATRIX(j, 10) = SAMPLING_MATRIX(i, 10) + SAMPLING_MATRIX(j, 9)
    
'------------------------OBPI (Option Based Portfolio Insurance)----------------------------------
    
    TEMP_VAL = DASPP_OBPI_EXPO_FUNC(SAMPLING_MATRIX(j, 5), TIME_TO_MATURITY, SAMPLING_MATRIX(j, 5), PROTECTION_LEVEL / OBPI_SCALE, (PERIODS - SAMPLING_MATRIX(j, 1)) / PERIODS_PER_YEAR, CASH_RATE, INITIAL_VOLATILITY)
    If MAX_RISKY_EXPOSURE < TEMP_VAL Then: TEMP_VAL = MAX_RISKY_EXPOSURE
    If MIN_RISKY_EXPOSURE > TEMP_VAL Then: TEMP_VAL = MIN_RISKY_EXPOSURE
    
    SAMPLING_MATRIX(j, 12) = TEMP_VAL
    
    TEMP_VAL = SAMPLING_MATRIX(j, 12) - SAMPLING_MATRIX(i, 14)
    If Abs(TEMP_VAL) > REWEIGHTING_TRIGGER Then
        SAMPLING_MATRIX(j, 13) = TEMP_VAL
    Else
        SAMPLING_MATRIX(j, 13) = 0
    End If
    SAMPLING_MATRIX(j, 14) = SAMPLING_MATRIX(j, 13) + SAMPLING_MATRIX(i, 14)


    SAMPLING_MATRIX(j, 15) = SAMPLING_MATRIX(i, 15) * (1 + SAMPLING_MATRIX(i, 14) * SAMPLING_MATRIX(j, 4) + (1 - SAMPLING_MATRIX(i, 14)) * CASH_RATE * DELTA_TIME) - INITIAL_WEALTH / INITIAL_RISKY_ASSET * Abs(SAMPLING_MATRIX(i, 13)) * TRANSACTION_COST
'--------------------------DPE (Dynamic Protected Envelope)----------------------------------
    SAMPLING_MATRIX(j, 16) = DASPP_DPE_INDEX_FUNC(SAMPLING_MATRIX(j, 1), DPE_THETA, DPE_FIRST, DPE_REWEIGHTING_PERIODS)
    SAMPLING_MATRIX(j, 21) = SAMPLING_MATRIX(i, 21) * (1 + SAMPLING_MATRIX(i, 20) * SAMPLING_MATRIX(j, 4) + (1 - SAMPLING_MATRIX(i, 20)) * CASH_RATE * DELTA_TIME) - INITIAL_WEALTH / INITIAL_RISKY_ASSET * Abs(SAMPLING_MATRIX(i, 19)) * TRANSACTION_COST
    TEMP_VAL = SAMPLING_MATRIX(j, 21) - SAMPLING_MATRIX(j, 6)
    If TEMP_VAL > tolerance Then
        SAMPLING_MATRIX(j, 17) = TEMP_VAL
    Else
        SAMPLING_MATRIX(j, 17) = 0
    End If
    If SAMPLING_MATRIX(j, 16) = 1 Then
        TEMP_VAL = SAMPLING_MATRIX(j, 17) * CPPI_MULTIPLIER / INITIAL_WEALTH
        If MAX_RISKY_EXPOSURE < TEMP_VAL Then: TEMP_VAL = MAX_RISKY_EXPOSURE
        If MIN_RISKY_EXPOSURE > TEMP_VAL Then: TEMP_VAL = MIN_RISKY_EXPOSURE
        SAMPLING_MATRIX(j, 18) = TEMP_VAL
    Else
        SAMPLING_MATRIX(j, 18) = SAMPLING_MATRIX(i, 20)
    End If
    TEMP_VAL = SAMPLING_MATRIX(j, 18) - SAMPLING_MATRIX(i, 20)
    If Abs(TEMP_VAL) > DPE_REWEIGHTING_TRIGGER Then
        SAMPLING_MATRIX(j, 19) = TEMP_VAL
    Else
        SAMPLING_MATRIX(j, 19) = 0
    End If
    SAMPLING_MATRIX(j, 20) = SAMPLING_MATRIX(j, 19) + SAMPLING_MATRIX(i, 20)
    
'---------------------------------DPAA (Dynamic Protected Asset Allocation)-----------------------
    SAMPLING_MATRIX(j, 22) = DPAA_THETA
    SAMPLING_MATRIX(j, 23) = SAMPLING_MATRIX(j, 22) ^ 2
    SAMPLING_MATRIX(j, 24) = SAMPLING_MATRIX(j, 22) * SAMPLING_MATRIX(j, 2)
    SAMPLING_MATRIX(j, 25) = DPAA_MU * Exp(-(2 * CASH_RATE - DPAA_THETA ^ 2) * TIME_TO_MATURITY) * Exp(CASH_RATE * SAMPLING_MATRIX(j, 1) * DELTA_TIME - 3 / 2 * TEMP1_SUM * DELTA_TIME - TEMP2_SUM * Sqr(DELTA_TIME))
    TEMP1_SUM = TEMP1_SUM + SAMPLING_MATRIX(j, 23)
    TEMP2_SUM = TEMP2_SUM + SAMPLING_MATRIX(j, 24)
    If i < PERIODS Then
        SAMPLING_MATRIX(j, 26) = (Log(SAMPLING_MATRIX(j, 25) / DPAA_LAMBDA) + (CASH_RATE + 1 / 2 * DPAA_THETA ^ 2) * (PERIODS - SAMPLING_MATRIX(j, 1)) * DELTA_TIME) / Sqr((PERIODS - SAMPLING_MATRIX(j, 1)) * DELTA_TIME) / DPAA_THETA
        SAMPLING_MATRIX(j, 27) = SAMPLING_MATRIX(j, 26) - Sqr((PERIODS - SAMPLING_MATRIX(j, 1)) * DELTA_TIME) * DPAA_THETA
        SAMPLING_MATRIX(j, 28) = DPAA_LAMBDA * CND_FUNC(-SAMPLING_MATRIX(j, 27), CND_TYPE) * Exp(-CASH_RATE * (PERIODS - SAMPLING_MATRIX(j, 1)) * DELTA_TIME) - CND_FUNC(-SAMPLING_MATRIX(j, 26), CND_TYPE) * SAMPLING_MATRIX(j, 25)
        TEMP_VAL = -(SAMPLING_MATRIX(j, 28) - DPAA_LAMBDA * CND_FUNC(-SAMPLING_MATRIX(j, 27), CND_TYPE) * Exp(-(PERIODS - SAMPLING_MATRIX(j, 1)) * DELTA_TIME * CASH_RATE)) * SAMPLING_MATRIX(j, 22) / SAMPLING_MATRIX(j, 3) * (INITIAL_WEALTH - SAMPLING_MATRIX(1, 6)) * DPAA_MULTIPLIER / SAMPLING_MATRIX(j, 5)

        If MAX_RISKY_EXPOSURE < TEMP_VAL Then: TEMP_VAL = MAX_RISKY_EXPOSURE
        If MIN_RISKY_EXPOSURE > TEMP_VAL Then: TEMP_VAL = MIN_RISKY_EXPOSURE
        SAMPLING_MATRIX(j, 29) = TEMP_VAL
    
        TEMP_VAL = SAMPLING_MATRIX(j, 29) - SAMPLING_MATRIX(i, 31)
        If Abs(TEMP_VAL) > tolerance Then
            SAMPLING_MATRIX(j, 30) = TEMP_VAL
        Else
            SAMPLING_MATRIX(j, 30) = 0
        End If
        SAMPLING_MATRIX(j, 31) = SAMPLING_MATRIX(i, 31) + SAMPLING_MATRIX(j, 30)
    Else
        SAMPLING_MATRIX(j, 26) = 0
        SAMPLING_MATRIX(j, 27) = 0
        SAMPLING_MATRIX(j, 28) = 0
        SAMPLING_MATRIX(j, 29) = 0
        SAMPLING_MATRIX(j, 30) = 0
        SAMPLING_MATRIX(j, 31) = 0
    End If
    SAMPLING_MATRIX(j, 32) = SAMPLING_MATRIX(i, 32) * (1 + SAMPLING_MATRIX(i, 31) * SAMPLING_MATRIX(j, 4) + (1 - SAMPLING_MATRIX(i, 31)) * CASH_RATE * DELTA_TIME) - INITIAL_WEALTH / INITIAL_RISKY_ASSET * Abs(SAMPLING_MATRIX(i, 30)) * TRANSACTION_COST
    If OUTPUT > 0 Then: GoSub SUMMARY_LINE
'--------------------------------------------------------------------------------------------------
Next i
'--------------------------------------------------------------------------------------------------
Select Case OUTPUT
'---------------------------------------------------------------------------------------------------
Case 0
    DASPP_SAMPLING_FUNC = SAMPLING_MATRIX
Case 1
    DASPP_SAMPLING_FUNC = STATISTICS_MATRIX
Case 2
    DASPP_SAMPLING_FUNC = DASPP_SAMPLING_SUMMARY_FUNC(STATISTICS_MATRIX)
Case 3
    DASPP_SAMPLING_FUNC = CHARTING_MATRIX
Case Else
    DASPP_SAMPLING_FUNC = Array(SAMPLING_MATRIX, STATISTICS_MATRIX, DASPP_SAMPLING_SUMMARY_FUNC(STATISTICS_MATRIX), CHARTING_MATRIX)
End Select
'----------------------------------------------------------------------------------------------------
Exit Function
'----------------------------------------------------------------------------------------------------
HEADERS_LINE:
'----------------------------------------------------------------------------------------------------
ReDim SAMPLING_MATRIX(h To PERIODS + 1, 1 To 32)
If HEADERS_FLAG = True Then
    SAMPLING_MATRIX(0, 1) = "PERIODS" 'TRADING DAY
    SAMPLING_MATRIX(0, 2) = "SHOCKS"
    SAMPLING_MATRIX(0, 3) = "VOLATILITY"
    SAMPLING_MATRIX(0, 4) = "RISKY RETURN"
    SAMPLING_MATRIX(0, 5) = "RISKY WEALTH"
    SAMPLING_MATRIX(0, 6) = "PROTECTION LEVEL"
    
    'CPPI (Constant Proportional Portfolio Insurance)
    SAMPLING_MATRIX(0, 7) = "CPPI: CUSHION"
    SAMPLING_MATRIX(0, 8) = "CPPI: R-EXPO"
    SAMPLING_MATRIX(0, 9) = "CPPI: TURNOVER"
    SAMPLING_MATRIX(0, 10) = "CPPI: STRATEGY"
    SAMPLING_MATRIX(0, 11) = "CPPI: WEALTH"
    'OBPI (Option Based Portfolio Insurance)
    SAMPLING_MATRIX(0, 12) = "OBPI: R-EXPO"
    SAMPLING_MATRIX(0, 13) = "OBPI: TURNOVER"
    SAMPLING_MATRIX(0, 14) = "OBPI: STRATEGY"
    SAMPLING_MATRIX(0, 15) = "OBPI: WEALTH"
    'DPE (Dynamic Protected Envelope)
    SAMPLING_MATRIX(0, 16) = "DPE: INDICATOR"
    SAMPLING_MATRIX(0, 17) = "DPE: CUSHION"
    SAMPLING_MATRIX(0, 18) = "DPE: R-EXPO"
    SAMPLING_MATRIX(0, 19) = "DPE: TURNOVER"
    SAMPLING_MATRIX(0, 20) = "DPE: STRATEGY"
    SAMPLING_MATRIX(0, 21) = "DPE: WEALTH"
    
    'DPAA (Dynamic Protected Asset Allocation)
    
    SAMPLING_MATRIX(0, 22) = "DPAA: M-PRICE RISK"
    SAMPLING_MATRIX(0, 23) = "DPAA: M-PRICE RISK SQUARED"
    SAMPLING_MATRIX(0, 24) = "DPAA: M-PRICE RISK TIMES dW"
    SAMPLING_MATRIX(0, 25) = "DPAA: yT"
    SAMPLING_MATRIX(0, 26) = "DPAA: d+(t,Yt)"
    SAMPLING_MATRIX(0, 27) = "DPAA: d-(t,Yt)"
    SAMPLING_MATRIX(0, 28) = "DPAA: OPTIMAL WEALTH"
    
    SAMPLING_MATRIX(0, 29) = "DPAA: R-EXPO"
    SAMPLING_MATRIX(0, 30) = "DPAA: TURNOVER"
    SAMPLING_MATRIX(0, 31) = "DPAA: STRATEGY"
    SAMPLING_MATRIX(0, 32) = "DPAA: WEALTH"
End If
'----------------------------------------------------------------------------------------------------
Return
'----------------------------------------------------------------------------------------------------
SUMMARY_LINE:
'----------------------------------------------------------------------------------------------------
If j = 0 Then
'----------------------------------------------------------------------------------------------------
    ReDim CHARTING_MATRIX(h To PERIODS + 1, 1 To 6)
    ReDim STATISTICS_MATRIX(h To PERIODS + 1, 1 To 26)
    
    If HEADERS_FLAG = True Then
        
        CHARTING_MATRIX(0, 1) = "PERIODS" 'TRADING DAY
        CHARTING_MATRIX(0, 2) = "RISKY WEALTH"
        CHARTING_MATRIX(0, 3) = "CPPI WEALTH"
        CHARTING_MATRIX(0, 4) = "OBPI WEALTH"
        CHARTING_MATRIX(0, 5) = "DPE WEALTH"
        CHARTING_MATRIX(0, 6) = "DPAA WEALTH"
        
        STATISTICS_MATRIX(0, 1) = "PERIODS" 'TRADING DAY
        
        STATISTICS_MATRIX(0, 2) = "RISKY ASSET"
        STATISTICS_MATRIX(0, 3) = "RISKY RETURN"
        STATISTICS_MATRIX(0, 4) = "RISKY DRAW DOWNS"
        STATISTICS_MATRIX(0, 5) = "RISKY DD COUNTS"
        STATISTICS_MATRIX(0, 6) = "RISKY TURNOVER"
        
        STATISTICS_MATRIX(0, 7) = "CPPI STRATEGY"
        STATISTICS_MATRIX(0, 8) = "CPPI RETURN"
        STATISTICS_MATRIX(0, 9) = "CPPI DRAW DOWNS"
        STATISTICS_MATRIX(0, 10) = "CPPI DD COUNTS"
        STATISTICS_MATRIX(0, 11) = "CPPI TURNOVER"
        
        STATISTICS_MATRIX(0, 12) = "OBPI STRATEGY"
        STATISTICS_MATRIX(0, 13) = "OBPI RETURN"
        STATISTICS_MATRIX(0, 14) = "OBPI DRAW DOWNS"
        STATISTICS_MATRIX(0, 15) = "OBPI DD COUNTS"
        STATISTICS_MATRIX(0, 16) = "OBPI TURNOVER"
        
        STATISTICS_MATRIX(0, 17) = "DPE STRATEGY"
        STATISTICS_MATRIX(0, 18) = "DPE RETURN"
        STATISTICS_MATRIX(0, 19) = "DPE DRAW DOWNS"
        STATISTICS_MATRIX(0, 20) = "DPE DD COUNTS"
        STATISTICS_MATRIX(0, 21) = "DPE TURNOVER"
        
        STATISTICS_MATRIX(0, 22) = "DPAA STRATEGY"
        STATISTICS_MATRIX(0, 23) = "DPAA RETURN"
        STATISTICS_MATRIX(0, 24) = "DPAA DRAW DOWNS"
        STATISTICS_MATRIX(0, 25) = "DPAA DD COUNTS"
        STATISTICS_MATRIX(0, 26) = "DPAA TURNOVER"
    End If
'----------------------------------------------------------------------------------------------------
ElseIf j = 1 Then
'----------------------------------------------------------------------------------------------------
    CHARTING_MATRIX(j, 1) = SAMPLING_MATRIX(j, 1)
    CHARTING_MATRIX(j, 2) = SAMPLING_MATRIX(j, 5)
    CHARTING_MATRIX(j, 3) = SAMPLING_MATRIX(j, 11)
    CHARTING_MATRIX(j, 4) = SAMPLING_MATRIX(j, 15)
    CHARTING_MATRIX(j, 5) = SAMPLING_MATRIX(j, 21)
    CHARTING_MATRIX(j, 6) = SAMPLING_MATRIX(j, 32)

    STATISTICS_MATRIX(j, 1) = SAMPLING_MATRIX(j, 1)
    STATISTICS_MATRIX(j, 2) = SAMPLING_MATRIX(j, 5)
    STATISTICS_MATRIX(j, 3) = 0
    STATISTICS_MATRIX(j, 4) = 0
    STATISTICS_MATRIX(j, 5) = 0
    STATISTICS_MATRIX(j, 6) = 1 '100%
    STATISTICS_MATRIX(j, 7) = SAMPLING_MATRIX(j, 11)
    STATISTICS_MATRIX(j, 8) = 0
    STATISTICS_MATRIX(j, 9) = 0
    STATISTICS_MATRIX(j, 10) = 0
    STATISTICS_MATRIX(j, 11) = Abs(SAMPLING_MATRIX(j, 9))
    STATISTICS_MATRIX(j, 12) = SAMPLING_MATRIX(j, 15)
    STATISTICS_MATRIX(j, 13) = 0
    STATISTICS_MATRIX(j, 14) = 0
    STATISTICS_MATRIX(j, 15) = 0
    STATISTICS_MATRIX(j, 16) = Abs(SAMPLING_MATRIX(j, 13))
    STATISTICS_MATRIX(j, 17) = SAMPLING_MATRIX(j, 21)
    STATISTICS_MATRIX(j, 18) = 0
    STATISTICS_MATRIX(j, 19) = 0
    STATISTICS_MATRIX(j, 20) = 0
    STATISTICS_MATRIX(j, 21) = Abs(SAMPLING_MATRIX(j, 19))
    STATISTICS_MATRIX(j, 22) = SAMPLING_MATRIX(j, 32)
    STATISTICS_MATRIX(j, 23) = 0
    STATISTICS_MATRIX(j, 24) = 0
    STATISTICS_MATRIX(j, 25) = 0
    STATISTICS_MATRIX(j, 26) = Abs(SAMPLING_MATRIX(j, 30))
'----------------------------------------------------------------------------------------------------
Else
'----------------------------------------------------------------------------------------------------
    CHARTING_MATRIX(j, 1) = SAMPLING_MATRIX(j, 1)
    CHARTING_MATRIX(j, 2) = SAMPLING_MATRIX(j, 5)
    CHARTING_MATRIX(j, 3) = SAMPLING_MATRIX(j, 11)
    CHARTING_MATRIX(j, 4) = SAMPLING_MATRIX(j, 15)
    CHARTING_MATRIX(j, 5) = SAMPLING_MATRIX(j, 21)
    CHARTING_MATRIX(j, 6) = SAMPLING_MATRIX(j, 32)
        
    STATISTICS_MATRIX(j, 1) = SAMPLING_MATRIX(j, 1)
    STATISTICS_MATRIX(j, 2) = SAMPLING_MATRIX(j, 5)
    STATISTICS_MATRIX(j, 3) = STATISTICS_MATRIX(j, 2) / STATISTICS_MATRIX(i, 2) - 1
    STATISTICS_MATRIX(j, 4) = DASPP_DRAWDOWN_FUNC(STATISTICS_MATRIX(i, 4), STATISTICS_MATRIX(j, 3))
    STATISTICS_MATRIX(j, 5) = IIf(STATISTICS_MATRIX(j, 4) < 0, STATISTICS_MATRIX(i, 5) + 1, 0)
    
    STATISTICS_MATRIX(j, 6) = 0
    STATISTICS_MATRIX(j, 7) = SAMPLING_MATRIX(j, 11)
    STATISTICS_MATRIX(j, 8) = STATISTICS_MATRIX(j, 7) / STATISTICS_MATRIX(i, 7) - 1
    STATISTICS_MATRIX(j, 9) = DASPP_DRAWDOWN_FUNC(STATISTICS_MATRIX(i, 9), STATISTICS_MATRIX(j, 8))
    STATISTICS_MATRIX(j, 10) = IIf(STATISTICS_MATRIX(j, 9) < 0, STATISTICS_MATRIX(i, 10) + 1, 0)
    
    STATISTICS_MATRIX(j, 11) = Abs(SAMPLING_MATRIX(j, 9))
    STATISTICS_MATRIX(j, 12) = SAMPLING_MATRIX(j, 15)
    STATISTICS_MATRIX(j, 13) = STATISTICS_MATRIX(j, 12) / STATISTICS_MATRIX(i, 12) - 1
    STATISTICS_MATRIX(j, 14) = DASPP_DRAWDOWN_FUNC(STATISTICS_MATRIX(i, 14), STATISTICS_MATRIX(j, 13))
    STATISTICS_MATRIX(j, 15) = IIf(STATISTICS_MATRIX(j, 14) < 0, STATISTICS_MATRIX(i, 15) + 1, 0)
    
    STATISTICS_MATRIX(j, 16) = Abs(SAMPLING_MATRIX(j, 13))
    STATISTICS_MATRIX(j, 17) = SAMPLING_MATRIX(j, 21)
    STATISTICS_MATRIX(j, 18) = STATISTICS_MATRIX(j, 17) / STATISTICS_MATRIX(i, 17) - 1
    STATISTICS_MATRIX(j, 19) = DASPP_DRAWDOWN_FUNC(STATISTICS_MATRIX(i, 19), STATISTICS_MATRIX(j, 18))
    STATISTICS_MATRIX(j, 20) = IIf(STATISTICS_MATRIX(j, 19) < 0, STATISTICS_MATRIX(i, 20) + 1, 0)
    
    STATISTICS_MATRIX(j, 21) = Abs(SAMPLING_MATRIX(j, 19))
    STATISTICS_MATRIX(j, 22) = SAMPLING_MATRIX(j, 32)
    STATISTICS_MATRIX(j, 23) = STATISTICS_MATRIX(j, 22) / STATISTICS_MATRIX(i, 22) - 1
    STATISTICS_MATRIX(j, 24) = DASPP_DRAWDOWN_FUNC(STATISTICS_MATRIX(i, 24), STATISTICS_MATRIX(j, 23))
    STATISTICS_MATRIX(j, 25) = IIf(STATISTICS_MATRIX(j, 24) < 0, STATISTICS_MATRIX(i, 25) + 1, 0)
    
    STATISTICS_MATRIX(j, 26) = Abs(SAMPLING_MATRIX(j, 30))
'----------------------------------------------------------------------------------------------------
End If
'----------------------------------------------------------------------------------------------------
Return
'----------------------------------------------------------------------------------------------------
ERROR_LABEL:
DASPP_SAMPLING_FUNC = Err.number
End Function

'----------------------------------------------------------------------------------------------------
'----------------------------------------------------------------------------------------------------
Function DASPP_DPAA_THETA_SAMPLING_FUNC(Optional ByVal TIME_TO_MATURITY As Double = 2, _
Optional ByVal CASH_RATE As Double = 0.035, _
Optional ByVal MEAN_RETURN As Double = 0.2, _
Optional ByVal VOLATILITY_METHOD As Integer = 2, _
Optional ByVal INITIAL_VOLATILITY As Double = 0.3, _
Optional ByVal LONG_TERM_VOLATILITY As Double = 0.3, _
Optional ByVal VOLATILITY_MEAN_REVERSION As Double = 0.5, _
Optional ByVal VOLATILITY_OF_VOLATILITY As Double = 0.4, _
Optional ByVal VOLATILITY_RHO As Double = 0, _
Optional ByVal RANDOM_TYPE As Integer = 1, _
Optional ByVal SKEWNESS_VAL As Double = -1, _
Optional ByVal KURTOSIS_VAL As Double = 6, _
Optional ByVal PERIODS_PER_YEAR As Double = 250)

Dim i As Long
Dim j As Long
Dim PERIODS As Long
Dim DELTA_TIME As Double
'Dim TIME_INDEX As Long
Dim PREV_SHOCKS_VAL As Double
Dim SHOCKS_VAL As Double
Dim PREV_VOLATILITY_VAL As Double
Dim VOLATILITY_VAL As Double
Dim THETA_VAL As Double
Dim TEMP_SUM As Double
Dim RANDOM_VECTOR As Variant

On Error GoTo ERROR_LABEL

DELTA_TIME = 1 / PERIODS_PER_YEAR
PERIODS = FLOOR_FUNC(PERIODS_PER_YEAR * TIME_TO_MATURITY, 1)
'Total Trading days

RANDOM_VECTOR = DASPP_RANDOM_VECTOR_FUNC(RANDOM_TYPE, PERIODS + 1, SKEWNESS_VAL, KURTOSIS_VAL, 3)
j = 1
SHOCKS_VAL = RANDOM_VECTOR(j, 1)
VOLATILITY_VAL = INITIAL_VOLATILITY
THETA_VAL = (MEAN_RETURN - CASH_RATE) / VOLATILITY_VAL
TEMP_SUM = THETA_VAL
For i = 1 To PERIODS
    PREV_SHOCKS_VAL = SHOCKS_VAL
    PREV_VOLATILITY_VAL = VOLATILITY_VAL
    j = i + 1
    SHOCKS_VAL = RANDOM_VECTOR(j, 1)
    VOLATILITY_VAL = DASPP_NEXT_VOLATILITY_FUNC(PREV_SHOCKS_VAL, PREV_VOLATILITY_VAL, DELTA_TIME, LONG_TERM_VOLATILITY, VOLATILITY_MEAN_REVERSION, VOLATILITY_OF_VOLATILITY, VOLATILITY_RHO, 0, VOLATILITY_METHOD)
    THETA_VAL = (MEAN_RETURN - CASH_RATE) / VOLATILITY_VAL
    TEMP_SUM = TEMP_SUM + THETA_VAL
Next i

DASPP_DPAA_THETA_SAMPLING_FUNC = TEMP_SUM / (PERIODS + 1)

Exit Function
ERROR_LABEL:
DASPP_DPAA_THETA_SAMPLING_FUNC = Err.number
End Function

Function DASPP_NEXT_VOLATILITY_FUNC(ByVal PREV_SHOCKS_VAL As Double, _
ByVal PREV_VOLATILITY_VAL As Double, _
ByVal DELTA_TIME As Double, _
Optional ByVal LONG_TERM_VOLATILITY As Double = 0.3, _
Optional ByVal VOLATILITY_MEAN_REVERSION As Double = 0.5, _
Optional ByVal VOLATILITY_OF_VOLATILITY As Double = 0.4, _
Optional ByVal VOLATILITY_RHO As Double = 0, _
Optional ByVal NORMAL_RANDOM_VAL As Double = 0, _
Optional ByVal VOLATILITY_METHOD As Integer = 2)
'---------------------------------------------------------------------------------------------------
Dim TEMP_VAL As Double
Dim BETA0_VAL As Double 'Beta 0
Dim STOCH_VAL As Double
'---------------------------------------------------------------------------------------------------
'LONG_TERM_VOLATILITY --> 0.30
'VOLATILITY_MEAN_REVERSION --> beta1 --> 0.50
'VOLATILITY_OF_VOLATILITY --> beta2 --> 0.40
'VOLATILITY_RHO --> Only applies for the Heston Model --> Adjust the normal random
'numbers (bivar distribution)
'---------------------------------------------------------------------------------------------------
On Error GoTo ERROR_LABEL
'---------------------------------------------------------------------------------------------------
Select Case VOLATILITY_METHOD
'---------------------------------------------------------------------------------------------------
Case 0 'Constant
'---------------------------------------------------------------------------------------------------
    DASPP_NEXT_VOLATILITY_FUNC = PREV_VOLATILITY_VAL
'---------------------------------------------------------------------------------------------------
Case 1 'Heston
'---------------------------------------------------------------------------------------------------
'LONG_TERM_VOLATILITY - vbar
'VOLATILITY_MEAN_REVERSION - kappa
'VOLATILITY_OF_VOLATILITY - lambda
'VOLATILITY-RHO
'---------------------------------------------------------------------------------------------------
    If NORMAL_RANDOM_VAL = 0 Then: NORMAL_RANDOM_VAL = RANDOM_NORMAL_FUNC(0, 1, 0)
    STOCH_VAL = PREV_SHOCKS_VAL * VOLATILITY_RHO + NORMAL_RANDOM_VAL * Sqr(1 - VOLATILITY_RHO ^ 2)
    TEMP_VAL = PREV_VOLATILITY_VAL ^ 2 + VOLATILITY_MEAN_REVERSION * (LONG_TERM_VOLATILITY ^ 2 - PREV_VOLATILITY_VAL ^ 2) * DELTA_TIME + VOLATILITY_OF_VOLATILITY * PREV_VOLATILITY_VAL * STOCH_VAL * Sqr(DELTA_TIME)
    If TEMP_VAL < 0 Then: TEMP_VAL = 0
    DASPP_NEXT_VOLATILITY_FUNC = Sqr(TEMP_VAL)
'---------------------------------------------------------------------------------------------------
Case Else 'Garch
'r(t) = mu  + e(t) where
'e(t) = sigma(t) * epsilon(t),
'sigma (T + 1) ^ 2 = beta0 + beta1 * sigma(T) ^ 2 + beta2 * e(t)^2
'and beta0 = (1 - beta1 - beta2) * vbar^2
'where 0 < beta1 + beta2 < 1
'---------------------------------------------------------------------------------------------------
'LONG_TERM_VOLATILITY - vbar
'VOLATILITY_MEAN_REVERSION - BETA1 - between 0 And 1
'VOLATILITY_OF_VOLATILITY - BETA2 - between 0 And 1
    BETA0_VAL = (1 - VOLATILITY_MEAN_REVERSION - VOLATILITY_OF_VOLATILITY) * LONG_TERM_VOLATILITY ^ 2
    TEMP_VAL = BETA0_VAL + VOLATILITY_MEAN_REVERSION * PREV_VOLATILITY_VAL ^ 2 + VOLATILITY_OF_VOLATILITY * PREV_VOLATILITY_VAL ^ 2 * PREV_SHOCKS_VAL ^ 2
    If TEMP_VAL < 0 Then: TEMP_VAL = 0
    DASPP_NEXT_VOLATILITY_FUNC = Sqr(TEMP_VAL)
'---------------------------------------------------------------------------------------------------
End Select
'---------------------------------------------------------------------------------------------------

Exit Function
ERROR_LABEL:
DASPP_NEXT_VOLATILITY_FUNC = PUB_EPSILON
End Function

Function DASPP_RANDOM_VECTOR_FUNC(Optional ByVal RANDOM_TYPE As Integer = 1, _
Optional ByVal NSIZE As Long = 251, _
Optional ByVal SKEWNESS_VAL As Double = -1, _
Optional ByVal KURTOSIS_VAL As Double = 6, _
Optional ByVal FACTOR_VAL As Double = 3)

Dim i As Long
Dim PARAM_VECTOR As Variant
Dim RANDOM_VECTOR As Variant

On Error GoTo ERROR_LABEL

ReDim RANDOM_VECTOR(1 To NSIZE, 1 To 1)

'---------------------------------------------------------------------------------------------------------
Select Case RANDOM_TYPE
'---------------------------------------------------------------------------------------------------------
Case 0
'---------------------------------------------------------------------------------------------------------
    For i = 1 To NSIZE
        RANDOM_VECTOR(i, 1) = RANDOM_NORMAL_FUNC(0, 1, 0)
    Next i
'---------------------------------------------------------------------------------------------------------
Case Else 'Simulate a NIG distributed random variable as a mean-variance mixture (Rydberg-MC method)
'---------------------------------------------------------------------------------------------------------
    PARAM_VECTOR = NIG_MLE_PARAMETERS_FUNC(0, 1, SKEWNESS_VAL, KURTOSIS_VAL + FACTOR_VAL)
    For i = 1 To NSIZE
        RANDOM_VECTOR(i, 1) = NIG_RANDOM_FUNC(PARAM_VECTOR(1, 1), PARAM_VECTOR(2, 1), PARAM_VECTOR(3, 1), PARAM_VECTOR(4, 1))
    Next i
'---------------------------------------------------------------------------------------------------------
End Select
'---------------------------------------------------------------------------------------------------------

DASPP_RANDOM_VECTOR_FUNC = RANDOM_VECTOR

Exit Function
ERROR_LABEL:
DASPP_RANDOM_VECTOR_FUNC = Err.number
End Function

'Add Cumulative Normal Distribution Function

Private Function DASPP_OBPI_SCALE_FUNC(ByVal S_VAL As Double, _
ByVal K_VAL As Double, _
ByVal T_VAL As Double, _
ByVal R_VAL As Double, _
ByVal V_VAL As Double, _
Optional ByVal CND_TYPE As Integer = 0)

Dim D1_VAL As Double
Dim D2_VAL As Double
Dim TEMP_VAL As Double

On Error GoTo ERROR_LABEL

D1_VAL = (Log(S_VAL / K_VAL) + (R_VAL + V_VAL ^ 2 / 2) * T_VAL) / (V_VAL * Sqr(T_VAL))
D2_VAL = D1_VAL - V_VAL * Sqr(T_VAL)
TEMP_VAL = S_VAL - K_VAL * Exp(-R_VAL * T_VAL)

DASPP_OBPI_SCALE_FUNC = TEMP_VAL / (CND_FUNC(D1_VAL, CND_TYPE) * S_VAL - CND_FUNC(D2_VAL, CND_TYPE) * K_VAL * Exp(-R_VAL * T_VAL))

Exit Function
ERROR_LABEL:
DASPP_OBPI_SCALE_FUNC = Err.number
End Function


Private Function DASPP_OBPI_EXPO_FUNC(ByVal W0_VAL As Double, _
ByVal F_VAL As Double, _
ByVal S_VAL_VAL As Double, _
ByVal K_VAL As Double, _
ByVal T_VAL As Double, _
ByVal R_VAL As Double, _
ByVal V_VAL As Double, _
Optional ByVal CND_TYPE As Integer = 0)

Dim D1_VAL As Double
Dim D2_VAL As Double

On Error GoTo ERROR_LABEL

If T_VAL = 0 Then: T_VAL = 0.0000000001

D1_VAL = (Log(S_VAL_VAL / K_VAL) + (R_VAL + V_VAL ^ 2 / 2) * T_VAL) / (V_VAL * Sqr(T_VAL))
D2_VAL = D1_VAL - V_VAL * Sqr(T_VAL)

DASPP_OBPI_EXPO_FUNC = _
    F_VAL * CND_FUNC(D1_VAL, CND_TYPE) * S_VAL_VAL / (F_VAL * _
    CND_FUNC(D1_VAL, CND_TYPE) * S_VAL_VAL - F_VAL * _
    CND_FUNC(D2_VAL, CND_TYPE) * K_VAL * Exp(-R_VAL * _
    T_VAL) + K_VAL * Exp(-R_VAL * T_VAL))

Exit Function
ERROR_LABEL:
DASPP_OBPI_EXPO_FUNC = Err.number
End Function

Private Function DASPP_DPE_INDEX_FUNC(ByVal TIME_INDEX_VAL As Double, _
ByVal THETA_VAL As Double, _
ByVal DPE_FIRST_VAL As Double, _
ByVal PERIODS As Integer) As Integer

Dim i As Long

On Error GoTo ERROR_LABEL

DASPP_DPE_INDEX_FUNC = 0
For i = 1 To PERIODS
    If TIME_INDEX_VAL = Round(THETA_VAL ^ (i - 1) * DPE_FIRST_VAL, 0) Then: DASPP_DPE_INDEX_FUNC = 1
Next i

Exit Function
ERROR_LABEL:
DASPP_DPE_INDEX_FUNC = Err.number
End Function

Function DASPP_DPE_THETA_VALID_FUNC(ByVal DPE_THETA As Double, _
ByVal TIME_TO_MATURITY As Double, _
ByVal DPE_REWEIGHTING_PERIODS As Double, _
Optional ByVal ALOWER_BOUND As Double = 0.1, _
Optional ByVal BLOWER_BOUND As Double = 0.1, _
Optional ByVal AUPPER_BOUND As Double = 2, _
Optional ByVal BUPPER_BOUND As Double = 20, _
Optional ByVal AGUESS_VAL As Double = 0.5, _
Optional ByVal BGUESS_VAL As Double = 2, _
Optional ByVal nLOOPS As Long = 1000, _
Optional ByVal tolerance As Double = 0.00000001, _
Optional ByVal PERIODS_PER_YEAR As Double = 250)

Dim X_VAL As Double
Dim CONVERGE_FLAG As Integer
Dim COUNTER As Long
Dim PARAM_VECTOR As Variant

On Error GoTo ERROR_LABEL

ReDim PARAM_VECTOR(1 To 2, 1 To 1)

PUB_PERIODS = FLOOR_FUNC(PERIODS_PER_YEAR * TIME_TO_MATURITY, 1)
PUB_DPE_REWEIGHTING_PERIODS = DPE_REWEIGHTING_PERIODS

X_VAL = BRENT_ZERO_FUNC(ALOWER_BOUND, AUPPER_BOUND, "DASPP_DPE_THETA_OBJ1_FUNC", AGUESS_VAL, CONVERGE_FLAG, COUNTER, nLOOPS, tolerance)
If X_VAL = PUB_EPSILON Or CONVERGE_FLAG <> 0 Then: GoTo ERROR_LABEL
PARAM_VECTOR(1, 1) = X_VAL

X_VAL = BRENT_ZERO_FUNC(BLOWER_BOUND, BUPPER_BOUND, "DASPP_DPE_THETA_OBJ2_FUNC", BGUESS_VAL, CONVERGE_FLAG, COUNTER, nLOOPS, tolerance)
If X_VAL = PUB_EPSILON Or CONVERGE_FLAG <> 0 Then: GoTo ERROR_LABEL
PARAM_VECTOR(2, 1) = X_VAL

If DPE_THETA <= PARAM_VECTOR(1, 1) Or DPE_THETA >= PARAM_VECTOR(2, 1) Then
'Please input another number!"
    DASPP_DPE_THETA_VALID_FUNC = False
Else
    DASPP_DPE_THETA_VALID_FUNC = True
End If

Exit Function
ERROR_LABEL:
DASPP_DPE_THETA_VALID_FUNC = False
End Function

Private Function DASPP_DPE_THETA_OBJ1_FUNC(ByVal THETA_VAL As Double)

On Error GoTo ERROR_LABEL

If THETA_VAL = 1 Then
    DASPP_DPE_THETA_OBJ1_FUNC = PUB_PERIODS - PUB_DPE_REWEIGHTING_PERIODS
Else
    DASPP_DPE_THETA_OBJ1_FUNC = (THETA_VAL ^ PUB_DPE_REWEIGHTING_PERIODS * (PUB_PERIODS * THETA_VAL - PUB_PERIODS - 1) + 1) / (THETA_VAL - 1)
End If

Exit Function
ERROR_LABEL:
DASPP_DPE_THETA_OBJ1_FUNC = PUB_EPSILON 'Err.number
End Function


Private Function DASPP_DPE_THETA_OBJ2_FUNC(ByVal THETA_VAL As Double)

On Error GoTo ERROR_LABEL

If THETA_VAL = 1 Then
    DASPP_DPE_THETA_OBJ2_FUNC = PUB_DPE_REWEIGHTING_PERIODS - PUB_PERIODS
Else
    DASPP_DPE_THETA_OBJ2_FUNC = (THETA_VAL ^ PUB_DPE_REWEIGHTING_PERIODS - 1 - PUB_PERIODS * (THETA_VAL - 1)) / (THETA_VAL - 1)
End If

Exit Function
ERROR_LABEL:
DASPP_DPE_THETA_OBJ2_FUNC = PUB_EPSILON 'Err.number
End Function


'DPPA Dynamic Market Value (DPAA Parameters)

Function DASPP_DPAA_DMV_OPTIMIZER_FUNC(Optional ByVal INITIAL_WEALTH As Double = 100, _
Optional ByVal PROTECTION_LEVEL As Double = 90, _
Optional ByVal TIME_TO_MATURITY As Double = 1, _
Optional ByVal CASH_RATE As Double = 0.035, _
Optional ByVal MEAN_RETURN As Double = 0.2, _
Optional ByVal DPAA_THETA As Double = 0.930114173871195, _
Optional ByVal X0_VAL As Double = 0, _
Optional ByVal Y0_VAL As Double = 0, _
Optional ByVal CND_TYPE As Integer = 0)

Dim NO_MONTHS As Double
Dim CONVERGE_VAL As Integer
Dim COUNTER As Long
Dim nLOOPS As Long
Dim tolerance As Double
Dim epsilon As Double
Dim YTEMP_VECTOR As Variant

On Error GoTo ERROR_LABEL

NO_MONTHS = TIME_TO_MATURITY * 12 '
'DPAAFV = 1

ReDim PUB_CONSTANTS_VECTOR(1 To 6, 1 To 1)

PUB_CONSTANTS_VECTOR(1, 1) = CASH_RATE
PUB_CONSTANTS_VECTOR(2, 1) = DPAA_THETA
PUB_CONSTANTS_VECTOR(3, 1) = NO_MONTHS
PUB_CONSTANTS_VECTOR(4, 1) = (1 + MEAN_RETURN) ^ TIME_TO_MATURITY - 1 - (PROTECTION_LEVEL / INITIAL_WEALTH)
PUB_CONSTANTS_VECTOR(5, 1) = 1 - (PROTECTION_LEVEL / INITIAL_WEALTH) * Exp(-CASH_RATE * TIME_TO_MATURITY)
PUB_CONSTANTS_VECTOR(6, 1) = CND_TYPE

nLOOPS = 800: tolerance = 0.0000000001: epsilon = 10 ^ -5
YTEMP_VECTOR = NEWTON_BIVAR_ZERO_FUNC("DASPP_DPAA_DMV_OBJ_FUNC", X0_VAL, Y0_VAL, "DASPP_DPAA_DMV_GRAD_FUNC", CONVERGE_VAL, COUNTER, nLOOPS, tolerance, epsilon)

If IsArray(YTEMP_VECTOR) = False Then: GoTo ERROR_LABEL
If CONVERGE_VAL <> 1 Then: GoTo ERROR_LABEL

X0_VAL = Exp(CDbl(YTEMP_VECTOR(1, 1)))
Y0_VAL = Exp(CDbl(YTEMP_VECTOR(2, 1)))

ReDim YTEMP_VECTOR(1 To 2, 1 To 1)
YTEMP_VECTOR(1, 1) = X0_VAL 'lambda
YTEMP_VECTOR(2, 1) = Y0_VAL 'mu
DASPP_DPAA_DMV_OPTIMIZER_FUNC = YTEMP_VECTOR

Exit Function
ERROR_LABEL:
DASPP_DPAA_DMV_OPTIMIZER_FUNC = Err.number
End Function

Public Function DASPP_DPAA_DMV_OBJ_FUNC(ByRef PARAM_VECTOR As Variant) 'X_VAL As Double, _
Y_VAL As Double)

Dim X_VAL As Double
Dim Y_VAL As Double

'Dim PARAM_VECTOR As Variant
Dim YTEMP_VECTOR(1 To 2, 1 To 1) As Variant

On Error GoTo ERROR_LABEL

'PARAM_VECTOR = PARAM_RNG
'If UBound(PARAM_VECTOR, 1) = 1 Then: PARAM_VECTOR = MATRIX_TRANSPOSE_FUNC(PARAM_VECTOR)

X_VAL = PARAM_VECTOR(1, 1)
Y_VAL = PARAM_VECTOR(2, 1)

YTEMP_VECTOR(1, 1) = Exp(X_VAL) * CND_FUNC((X_VAL - Y_VAL + _
    (PUB_CONSTANTS_VECTOR(1, 1) - 1 / 2 * _
    PUB_CONSTANTS_VECTOR(2, 1) ^ 2) * PUB_CONSTANTS_VECTOR(3, 1) / 12) _
    / Sqr(PUB_CONSTANTS_VECTOR(3, 1) / 12) / PUB_CONSTANTS_VECTOR(2, 1), PUB_CONSTANTS_VECTOR(6, 1)) - _
    Exp(Y_VAL) * Exp(-(PUB_CONSTANTS_VECTOR(1, 1) - PUB_CONSTANTS_VECTOR(2, 1) ^ 2) _
    * PUB_CONSTANTS_VECTOR(3, 1) / 12) * CND_FUNC((X_VAL - Y_VAL + _
    (PUB_CONSTANTS_VECTOR(1, 1) - 3 / 2 * PUB_CONSTANTS_VECTOR(2, 1) ^ 2) * _
    PUB_CONSTANTS_VECTOR(3, 1) / 12) / Sqr(PUB_CONSTANTS_VECTOR(3, 1) / 12) / _
    PUB_CONSTANTS_VECTOR(2, 1), PUB_CONSTANTS_VECTOR(6, 1)) - Exp(PUB_CONSTANTS_VECTOR(1, 1) * _
    PUB_CONSTANTS_VECTOR(3, 1) / 12) * PUB_CONSTANTS_VECTOR(5, 1)

YTEMP_VECTOR(2, 1) = Exp(X_VAL) * CND_FUNC((X_VAL - Y_VAL + _
    (PUB_CONSTANTS_VECTOR(1, 1) + 1 / 2 * PUB_CONSTANTS_VECTOR(2, 1) ^ 2) * _
    PUB_CONSTANTS_VECTOR(3, 1) / 12) / Sqr(PUB_CONSTANTS_VECTOR(3, 1) / 12) / _
    PUB_CONSTANTS_VECTOR(2, 1), PUB_CONSTANTS_VECTOR(6, 1)) - Exp(Y_VAL) * Exp(-PUB_CONSTANTS_VECTOR(1, 1) * _
    PUB_CONSTANTS_VECTOR(3, 1) / 12) * CND_FUNC((X_VAL - Y_VAL + _
    (PUB_CONSTANTS_VECTOR(1, 1) - 1 / 2 * PUB_CONSTANTS_VECTOR(2, 1) ^ 2) * _
    PUB_CONSTANTS_VECTOR(3, 1) / 12) / Sqr(PUB_CONSTANTS_VECTOR(3, 1) / 12) / _
    PUB_CONSTANTS_VECTOR(2, 1), PUB_CONSTANTS_VECTOR(6, 1)) - 1 - PUB_CONSTANTS_VECTOR(4, 1)

DASPP_DPAA_DMV_OBJ_FUNC = YTEMP_VECTOR

Exit Function
ERROR_LABEL:
DASPP_DPAA_DMV_OBJ_FUNC = Err.number
End Function

Public Function DASPP_DPAA_DMV_GRAD_FUNC(ByRef PARAM_VECTOR As Variant) 'X_VAL As Double, _
Y_VAL As Double)

Dim X_VAL As Double
Dim Y_VAL As Double

'Dim PARAM_VECTOR As Variant
Dim GRADIENT_MATRIX(1 To 2, 1 To 2) As Variant

On Error GoTo ERROR_LABEL

'PARAM_VECTOR = PARAM_RNG
'If UBound(PARAM_VECTOR, 1) = 1 Then: PARAM_VECTOR = MATRIX_TRANSPOSE_FUNC(PARAM_VECTOR)

X_VAL = PARAM_VECTOR(1, 1)
Y_VAL = PARAM_VECTOR(2, 1)

GRADIENT_MATRIX(1, 1) = _
    Exp(X_VAL) * CND_FUNC((X_VAL - Y_VAL + _
    (PUB_CONSTANTS_VECTOR(1, 1) - 1 / 2 * PUB_CONSTANTS_VECTOR(2, 1) ^ 2) * _
    PUB_CONSTANTS_VECTOR(3, 1) / 12) / Sqr(PUB_CONSTANTS_VECTOR(3, 1) / 12) / _
    PUB_CONSTANTS_VECTOR(2, 1), PUB_CONSTANTS_VECTOR(6, 1)) + Exp(X_VAL) * _
    NORMAL_MASS_DIST_FUNC((X_VAL - Y_VAL + _
    (PUB_CONSTANTS_VECTOR(1, 1) - 1 / 2 * PUB_CONSTANTS_VECTOR(2, 1) ^ 2) * _
    PUB_CONSTANTS_VECTOR(3, 1) / 12) / Sqr(PUB_CONSTANTS_VECTOR(3, 1) / 12) / _
    PUB_CONSTANTS_VECTOR(2, 1)) / Sqr(PUB_CONSTANTS_VECTOR(3, 1) / 12) / _
    PUB_CONSTANTS_VECTOR(2, 1) - Exp(Y_VAL) * Exp(-(PUB_CONSTANTS_VECTOR(1, 1) - _
    PUB_CONSTANTS_VECTOR(2, 1) ^ 2) * PUB_CONSTANTS_VECTOR(3, 1) / 12) * _
    NORMAL_MASS_DIST_FUNC((X_VAL - Y_VAL + (PUB_CONSTANTS_VECTOR(1, 1) - 3 / 2 * _
    PUB_CONSTANTS_VECTOR(2, 1) ^ 2) * PUB_CONSTANTS_VECTOR(3, 1) / 12) / _
    Sqr(PUB_CONSTANTS_VECTOR(3, 1) / 12) / PUB_CONSTANTS_VECTOR(2, 1)) / _
    Sqr(PUB_CONSTANTS_VECTOR(3, 1) / 12) / PUB_CONSTANTS_VECTOR(2, 1)

GRADIENT_MATRIX(1, 2) = _
    -Exp(X_VAL) * NORMAL_MASS_DIST_FUNC((X_VAL - Y_VAL + (PUB_CONSTANTS_VECTOR(1, 1) - _
    1 / 2 * PUB_CONSTANTS_VECTOR(2, 1) ^ 2) * PUB_CONSTANTS_VECTOR(3, 1) / 12) _
    / Sqr(PUB_CONSTANTS_VECTOR(3, 1) / 12) / PUB_CONSTANTS_VECTOR(2, 1)) / _
    Sqr(PUB_CONSTANTS_VECTOR(3, 1) / 12) / PUB_CONSTANTS_VECTOR(2, 1) - _
    Exp(Y_VAL) * Exp(-(PUB_CONSTANTS_VECTOR(1, 1) - PUB_CONSTANTS_VECTOR(2, 1) ^ 2) _
    * PUB_CONSTANTS_VECTOR(3, 1) / 12) * CND_FUNC((X_VAL - Y_VAL + _
    (PUB_CONSTANTS_VECTOR(1, 1) - 3 / 2 * PUB_CONSTANTS_VECTOR(2, 1) ^ 2) * _
    PUB_CONSTANTS_VECTOR(3, 1) / 12) / Sqr(PUB_CONSTANTS_VECTOR(3, 1) / 12) / _
    PUB_CONSTANTS_VECTOR(2, 1), PUB_CONSTANTS_VECTOR(6, 1)) + Exp(Y_VAL) * Exp(-(PUB_CONSTANTS_VECTOR(1, 1) - _
    PUB_CONSTANTS_VECTOR(2, 1) ^ 2) * PUB_CONSTANTS_VECTOR(3, 1) / 12) * _
    NORMAL_MASS_DIST_FUNC((X_VAL - Y_VAL + (PUB_CONSTANTS_VECTOR(1, 1) - 3 / 2 * _
    PUB_CONSTANTS_VECTOR(2, 1) ^ 2) * PUB_CONSTANTS_VECTOR(3, 1) / 12) / _
    Sqr(PUB_CONSTANTS_VECTOR(3, 1) / 12) / PUB_CONSTANTS_VECTOR(2, 1)) / _
    Sqr(PUB_CONSTANTS_VECTOR(3, 1) / 12) / PUB_CONSTANTS_VECTOR(2, 1)

GRADIENT_MATRIX(2, 1) = _
    Exp(X_VAL) * CND_FUNC((X_VAL - Y_VAL + (PUB_CONSTANTS_VECTOR(1, 1) + _
    1 / 2 * PUB_CONSTANTS_VECTOR(2, 1) ^ 2) * PUB_CONSTANTS_VECTOR(3, 1) / 12) / _
    Sqr(PUB_CONSTANTS_VECTOR(3, 1) / 12) / PUB_CONSTANTS_VECTOR(2, 1), PUB_CONSTANTS_VECTOR(6, 1)) + _
    Exp(X_VAL) * NORMAL_MASS_DIST_FUNC((X_VAL - Y_VAL + (PUB_CONSTANTS_VECTOR(1, 1) + 1 / 2 * _
    PUB_CONSTANTS_VECTOR(2, 1) ^ 2) * PUB_CONSTANTS_VECTOR(3, 1) / 12) / _
    Sqr(PUB_CONSTANTS_VECTOR(3, 1) / 12) / PUB_CONSTANTS_VECTOR(2, 1)) / _
    Sqr(PUB_CONSTANTS_VECTOR(3, 1) / 12) / PUB_CONSTANTS_VECTOR(2, 1) - _
    Exp(Y_VAL) * Exp(-PUB_CONSTANTS_VECTOR(1, 1) * PUB_CONSTANTS_VECTOR(3, 1) / _
    12) * NORMAL_MASS_DIST_FUNC((X_VAL - Y_VAL + (PUB_CONSTANTS_VECTOR(1, 1) - 1 / 2 * _
    PUB_CONSTANTS_VECTOR(2, 1) ^ 2) * PUB_CONSTANTS_VECTOR(3, 1) / 12) / _
    Sqr(PUB_CONSTANTS_VECTOR(3, 1) / 12) / PUB_CONSTANTS_VECTOR(2, 1)) / _
    Sqr(PUB_CONSTANTS_VECTOR(3, 1) / 12) / PUB_CONSTANTS_VECTOR(2, 1)

GRADIENT_MATRIX(2, 2) = _
    -Exp(X_VAL) * NORMAL_MASS_DIST_FUNC((X_VAL - Y_VAL + (PUB_CONSTANTS_VECTOR(1, 1) + _
    1 / 2 * PUB_CONSTANTS_VECTOR(2, 1) ^ 2) * PUB_CONSTANTS_VECTOR(3, 1) / 12) _
    / Sqr(PUB_CONSTANTS_VECTOR(3, 1) / 12) / PUB_CONSTANTS_VECTOR(2, 1)) / _
    Sqr(PUB_CONSTANTS_VECTOR(3, 1) / 12) / PUB_CONSTANTS_VECTOR(2, 1) - _
    Exp(Y_VAL) * Exp(-PUB_CONSTANTS_VECTOR(1, 1) * PUB_CONSTANTS_VECTOR(3, 1) / _
    12) * CND_FUNC((X_VAL - Y_VAL + (PUB_CONSTANTS_VECTOR(1, 1) - 1 / 2 * _
    PUB_CONSTANTS_VECTOR(2, 1) ^ 2) * PUB_CONSTANTS_VECTOR(3, 1) / 12) / _
    Sqr(PUB_CONSTANTS_VECTOR(3, 1) / 12) / PUB_CONSTANTS_VECTOR(2, 1), PUB_CONSTANTS_VECTOR(6, 1)) + _
    Exp(Y_VAL) * Exp(-PUB_CONSTANTS_VECTOR(1, 1) * PUB_CONSTANTS_VECTOR(3, 1) / _
    12) * NORMAL_MASS_DIST_FUNC((X_VAL - Y_VAL + (PUB_CONSTANTS_VECTOR(1, 1) - 1 / 2 * _
    PUB_CONSTANTS_VECTOR(2, 1) ^ 2) * PUB_CONSTANTS_VECTOR(3, 1) / 12) / _
    Sqr(PUB_CONSTANTS_VECTOR(3, 1) / 12) / PUB_CONSTANTS_VECTOR(2, 1)) / _
    Sqr(PUB_CONSTANTS_VECTOR(3, 1) / 12) / PUB_CONSTANTS_VECTOR(2, 1)

DASPP_DPAA_DMV_GRAD_FUNC = GRADIENT_MATRIX

Exit Function
ERROR_LABEL:
DASPP_DPAA_DMV_GRAD_FUNC = Err.number
End Function

Private Function DASPP_DRAWDOWN_FUNC(ByVal PREVIOUS_VAL As Double, _
ByVal CURRENT_VAL As Double)

Dim TEMP_VAL As Double

On Error GoTo ERROR_LABEL

If ((CURRENT_VAL < 0 And PREVIOUS_VAL = 0)) Then
    DASPP_DRAWDOWN_FUNC = CURRENT_VAL
Else
    TEMP_VAL = (1 + CURRENT_VAL) * (1 + PREVIOUS_VAL) - 1
    If TEMP_VAL < 0 Then
        DASPP_DRAWDOWN_FUNC = TEMP_VAL
    Else
        DASPP_DRAWDOWN_FUNC = 0
    End If
End If

Exit Function
ERROR_LABEL:
DASPP_DRAWDOWN_FUNC = Err.number
End Function


Private Function DASPP_SAMPLING_SUMMARY_FUNC(ByRef DATA_MATRIX As Variant, _
Optional ByVal INITIAL_WEALTH As Double = 100, _
Optional ByVal CASH_RATE As Double = 0.035, _
Optional ByVal TIME_TO_MATURITY As Double = 1, _
Optional ByVal PERIODS_PER_YEAR As Double = 250)

Dim h As Long
Dim i As Long
Dim l As Long
Dim k As Long
Dim j As Long

Dim NROWS As Long
Dim PERIODS As Long
Dim DELTA_TIME As Double

Dim TEMP_VAL As Double
'Dim DATA_MATRIX As Variant
Dim SUMMARY_MATRIX As Variant 'Strategy Summary Statistics

Dim TEMP_MATRIX As Variant
Dim PRICES_MATRIX As Variant
Dim RETURNS_MATRIX As Variant

Dim tolerance As Double

On Error GoTo ERROR_LABEL
'-------------------------------------------------------------------------------------------
tolerance = 2 ^ 52
PERIODS = FLOOR_FUNC(PERIODS_PER_YEAR * TIME_TO_MATURITY, 1) 'trading days per year
DELTA_TIME = 1 / PERIODS_PER_YEAR
'-------------------------------------------------------------------------------------------
'DATA_MATRIX = DATA_RNG
NROWS = UBound(DATA_MATRIX, 1)

'-------------------------------------------------------------------------------------------
ReDim SUMMARY_MATRIX(1 To 36, 1 To 6)
ReDim PRICES_MATRIX(1 To NROWS, 1 To 5) 'Risky/CPPI/OBPI/DPE/DPAA
ReDim RETURNS_MATRIX(1 To NROWS - 1, 1 To 5) 'Risky/CPPI/OBPI/DPE/DPAA
'-------------------------------------------------------------------------------------------

'-------------------------------------------------------------------------------------------
SUMMARY_MATRIX(1, 1) = "Basics"
SUMMARY_MATRIX(1, 2) = "Return" 'Annualized Final Return of the Strategy
SUMMARY_MATRIX(1, 3) = "Largest DD"
SUMMARY_MATRIX(1, 4) = "Longest DD"
SUMMARY_MATRIX(1, 5) = "Largest Loss"
SUMMARY_MATRIX(1, 6) = "TurnOver"
'-------------------------------------------------------------------------------------------
SUMMARY_MATRIX(7, 1) = "Per Period Returns"
SUMMARY_MATRIX(13, 1) = "Annualized Returns"

l = 7
For i = 1 To 2
    SUMMARY_MATRIX(l, 2) = "Average"
    SUMMARY_MATRIX(l, 3) = "Volatility"
    SUMMARY_MATRIX(l, 4) = "Skewness"
    SUMMARY_MATRIX(l, 5) = "90% VaR"
    SUMMARY_MATRIX(l, 6) = "99% VaR"
    l = l + 6
Next i

'-------------------------------------------------------------------------------------------
'Ratios  Sharp   Sortino Upside-P    Win Loss
SUMMARY_MATRIX(19, 1) = "Ratios"
SUMMARY_MATRIX(19, 2) = "Sharp"
SUMMARY_MATRIX(19, 3) = "Sortino"
SUMMARY_MATRIX(19, 4) = "Upside-P"
SUMMARY_MATRIX(19, 5) = "Win"
SUMMARY_MATRIX(19, 6) = "Loss"
'-------------------------------------------------------------------------------------------
'Asset/Return Corr  Risky   CPPI    OBPI    DPE DPAA
SUMMARY_MATRIX(25, 1) = "Asse Corr"
SUMMARY_MATRIX(31, 1) = "Return Corr"

l = 1
For i = 1 To 6
    SUMMARY_MATRIX(l + 1, 1) = "Risky"
    SUMMARY_MATRIX(l + 2, 1) = "CPPI" 'Constant Proportional Portfolio Insurance
    SUMMARY_MATRIX(l + 3, 1) = "OBPI" 'Option-Based Portfolio Insurance
    SUMMARY_MATRIX(l + 4, 1) = "DPE" 'Dynamic Protected Envelope Strategy
    SUMMARY_MATRIX(l + 5, 1) = "DPAA" 'Dynamic Protected Asset Allocation
    l = l + 6
Next i

l = 25
For i = 1 To 2
    SUMMARY_MATRIX(l, 2) = "Risky"
    SUMMARY_MATRIX(l, 3) = "CPPI" 'Constant Proportional Portfolio Insurance
    SUMMARY_MATRIX(l, 4) = "OBPI" 'Option-Based Portfolio Insurance
    SUMMARY_MATRIX(l, 5) = "DPE" 'Dynamic Protected Envelope Strategy
    SUMMARY_MATRIX(l, 6) = "DPAA" 'Dynamic Protected Asset Allocation
    l = l + 6
Next i


'------------------------------------------------------------------------------------------------------
'------------------------------------X Pass Min/Max/Average
'------------------------------------------------------------------------------------------------------
l = 2

For h = 0 To 4
    SUMMARY_MATRIX(l + h, 3) = tolerance
    SUMMARY_MATRIX(l + h, 4) = -1 * tolerance
    SUMMARY_MATRIX(l + h, 5) = tolerance
    SUMMARY_MATRIX(l + h, 6) = 0
    For j = 2 To 6
        SUMMARY_MATRIX(l + h + 6 * 1, j) = 0
        SUMMARY_MATRIX(l + h + 6 * 2, j) = 0
        SUMMARY_MATRIX(l + h + 6 * 3, j) = 0
        SUMMARY_MATRIX(l + h + 6 * 4, j) = 0
        SUMMARY_MATRIX(l + h + 6 * 5, j) = 0
    Next j
Next h

For i = 1 To NROWS
    j = 1: k = 0
    For h = 0 To 4
        If DATA_MATRIX(i, k + 4) < SUMMARY_MATRIX(l + h, 3) Then: SUMMARY_MATRIX(l + h, 3) = DATA_MATRIX(i, k + 4)
        If DATA_MATRIX(i, k + 5) > SUMMARY_MATRIX(l + h, 4) Then: SUMMARY_MATRIX(l + h, 4) = DATA_MATRIX(i, k + 5)
        If DATA_MATRIX(i, k + 2) < SUMMARY_MATRIX(l + h, 5) Then: SUMMARY_MATRIX(l + h, 5) = DATA_MATRIX(i, k + 2)
        
        SUMMARY_MATRIX(l + h, 6) = SUMMARY_MATRIX(l + h, 6) + IIf(DATA_MATRIX(i, k + 6) = "", 0, DATA_MATRIX(i, k + 6))
        SUMMARY_MATRIX(l + h + 6, 2) = SUMMARY_MATRIX(l + h + 6, 2) + IIf(DATA_MATRIX(i, k + 3) = "", 0, DATA_MATRIX(i, k + 3))
        PRICES_MATRIX(i, j) = DATA_MATRIX(i, k + 2)
        If i > 1 Then: RETURNS_MATRIX(i - 1, j) = DATA_MATRIX(i, k + 3)
        
        k = k + 5
        j = j + 1
    Next h
Next i

j = 2: k = 2
For h = 0 To 4
    SUMMARY_MATRIX(j, 2) = (DATA_MATRIX(NROWS, k) / DATA_MATRIX(1, k)) ^ (1 / TIME_TO_MATURITY) - 1
    SUMMARY_MATRIX(j, 5) = SUMMARY_MATRIX(j, 5) - INITIAL_WEALTH
    SUMMARY_MATRIX(j + 6, 2) = SUMMARY_MATRIX(j + 6, 2) / (NROWS - 1)
    SUMMARY_MATRIX(j + 6 * 2, 2) = SUMMARY_MATRIX(j + 6, 2) * PERIODS
    j = j + 1
    k = k + 5
Next h

'------------------------------------------------------------------------------------------------------
'-----------------------------------X Pass Standard Deviation / Sharpe
'------------------------------------------------------------------------------------------------------
For i = 2 To NROWS
    k = 0
    For h = 0 To 4
        SUMMARY_MATRIX(l + h + 6, 3) = SUMMARY_MATRIX(l + h + 6, 3) + (DATA_MATRIX(i, k + 3) - SUMMARY_MATRIX(l + h + 6, 2)) ^ 2
        k = k + 5
    Next h
Next i

j = 2
For h = 0 To 4
    SUMMARY_MATRIX(j + 6, 3) = (SUMMARY_MATRIX(j + 6, 3) / (NROWS - 2)) ^ 0.5
    SUMMARY_MATRIX(j + 6 * 2, 3) = SUMMARY_MATRIX(j + 6, 3) * Sqr(PERIODS)
    'Sharpe
    SUMMARY_MATRIX(j + 6 * 3, 2) = (SUMMARY_MATRIX(j + 6 * 2, 2) - CASH_RATE) / SUMMARY_MATRIX(j + 6 * 2, 3)
    j = j + 1
Next h

'------------------------------------------------------------------------------------------------------
'----------------------------------X Pass Skewness/VaR/Sortino/Upside P/Win/Loss
'------------------------------------------------------------------------------------------------------
For i = 2 To NROWS
    k = 0
    For h = 0 To 4
        SUMMARY_MATRIX(l + h + 6, 4) = SUMMARY_MATRIX(l + h + 6, 4) + ((DATA_MATRIX(i, k + 3) - SUMMARY_MATRIX(l + h + 6, 2)) / SUMMARY_MATRIX(l + h + 6, 3)) ^ 3
        TEMP_VAL = DATA_MATRIX(i, k + 3)
        If TEMP_VAL < 0 Then
            SUMMARY_MATRIX(l + h + 6 * 3, 6) = SUMMARY_MATRIX(l + h + 6 * 3, 6) + 1
        Else
            SUMMARY_MATRIX(l + h + 6 * 3, 5) = SUMMARY_MATRIX(l + h + 6 * 3, 5) + 1
        End If
        
        TEMP_VAL = TEMP_VAL - CASH_RATE * DELTA_TIME
        If TEMP_VAL < 0 Then
            SUMMARY_MATRIX(l + h + 6 * 3, 3) = SUMMARY_MATRIX(l + h + 6 * 3, 3) + TEMP_VAL ^ 2
        Else
            SUMMARY_MATRIX(l + h + 6 * 3, 4) = SUMMARY_MATRIX(l + h + 6 * 3, 4) + TEMP_VAL
        End If
        k = k + 5
    Next h
Next i

j = 2
For h = 0 To 4
    SUMMARY_MATRIX(j + 6, 4) = SUMMARY_MATRIX(j + 6, 4) / (NROWS - 1)
    SUMMARY_MATRIX(j + 6 * 2, 4) = SUMMARY_MATRIX(j + 6, 4)
    
    TEMP_MATRIX = MATRIX_GET_COLUMN_FUNC(RETURNS_MATRIX, h + 1, 1)
    TEMP_MATRIX = MATRIX_QUICK_SORT_FUNC(TEMP_MATRIX, 1, 1)
    SUMMARY_MATRIX(j + 6, 5) = HISTOGRAM_PERCENTILE_FUNC(TEMP_MATRIX, 0.1, 0)
    SUMMARY_MATRIX(j + 6 * 2, 5) = SUMMARY_MATRIX(j + 6, 5) * Sqr(PERIODS)
    
    SUMMARY_MATRIX(j + 6, 6) = HISTOGRAM_PERCENTILE_FUNC(TEMP_MATRIX, 0.01, 0)
    SUMMARY_MATRIX(j + 6 * 2, 6) = SUMMARY_MATRIX(j + 6, 6) * Sqr(PERIODS)
    
    SUMMARY_MATRIX(j + 6 * 3, 4) = SUMMARY_MATRIX(j + 6 * 3, 4) / Sqr(SUMMARY_MATRIX(j + 6 * 3, 3))
    SUMMARY_MATRIX(j + 6 * 3, 3) = (SUMMARY_MATRIX(j + 6 * 2, 2) - CASH_RATE) / Sqr(SUMMARY_MATRIX(j + 6 * 3, 3))

    SUMMARY_MATRIX(j + 6 * 3, 5) = SUMMARY_MATRIX(j + 6 * 3, 5) / (NROWS - 1)
    SUMMARY_MATRIX(j + 6 * 3, 6) = SUMMARY_MATRIX(j + 6 * 3, 6) / (NROWS - 1)
    j = j + 1
Next h
Erase TEMP_MATRIX

'------------------------------------------------------------------------------------------------------
'---------------------------------------X Pass Correlation
'------------------------------------------------------------------------------------------------------

PRICES_MATRIX = MATRIX_CORRELATION_FUNC(PRICES_MATRIX)
RETURNS_MATRIX = MATRIX_CORRELATION_FUNC(RETURNS_MATRIX)

For j = 1 To 5
    For i = 1 To 5
        SUMMARY_MATRIX(i + 25, j + 1) = PRICES_MATRIX(i, j)
        SUMMARY_MATRIX(i + 31, j + 1) = RETURNS_MATRIX(i, j)
    Next i
Next j
Erase PRICES_MATRIX
Erase RETURNS_MATRIX

DASPP_SAMPLING_SUMMARY_FUNC = SUMMARY_MATRIX

Exit Function
ERROR_LABEL:
DASPP_SAMPLING_SUMMARY_FUNC = Err.number
End Function

Function DASPP_SANITIZED_MATRIX_FUNC(ByRef DATA_MATRIX As Variant)

Dim i As Long
Dim j As Long
Dim k As Long

Dim NROWS As Long

'Dim DATA_MATRIX As Variant
Dim PRICES_MATRIX As Variant

On Error GoTo ERROR_LABEL

'DATA_MATRIX = DATA_RNG
NROWS = UBound(DATA_MATRIX, 1)

ReDim PRICES_MATRIX(1 To NROWS, 1 To 5)
For i = 1 To NROWS
    j = 1: k = 0
    For h = 0 To 4
        PRICES_MATRIX(i, j) = DATA_MATRIX(i, k + 2)
        k = k + 5
        j = j + 1
    Next h
Next i

DASPP_SANITIZED_MATRIX_FUNC = PRICES_MATRIX

Exit Function
ERROR_LABEL:
DASPP_SANITIZED_MATRIX_FUNC = Err.number
End Function

