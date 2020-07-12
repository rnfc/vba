Attribute VB_Name = "FINAN_ASSET_WAVES_CRASH_LIBR"


Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.

'------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------
Private PUB_DATA_MATRIX As Variant
Private PUB_CRASH_LEVEL As Double
Private PUB_PERIODS_FORWARD As Long
Private PUB_HOLIDAYS_RNG As Variant
Private Const PUB_EPSILON As Double = 1E+100
'------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------

'************************************************************************************
'************************************************************************************



'************************************************************************************
'************************************************************************************
'FUNCTION      : ASSET_WAVES_CRASH_FUNC

'DESCRIPTION   : Well, there have been attempts in the past few years to predict
'market crashes. These seem to have centred on identifying certain
'periodicities in a market index, like the S&P 500. It all started in
'the early 1990s, I think, with a French fellow, Didier Sornette, who
'tried to predict the failure of materials by listening to the noises
'they make as the breaking point approached.

'P(n+1) = [ 2-(2p/T)2 ] P(n) - P(n-1) + (2p/T)2 P0
'Here:
'T is some kind of period (like T = 5 days)
'P0 is some parameter (maybe a 10 day Moving Average)
'P (n) Is today 's stock price and P(n-1) is the price yesterday
'P(n+1) is the next stock price in the sequence ... namely tomorrow's price.
 
'We 'll look at the 5-day moving average of S&P500 closing values from
'X to Y.
'(We'll use 5-day average to get something a wee bit smooother than
'daily values.)
'We 'll pick a To and a reduction factor f so that, with T = Tn = To f n,
'which looks like  the magic equation:
'P(n+1) = [ 2-{2p/Tn}2 ] P(n) - P(n-1) + {2p/Tn }2 P0
'For P0 we'll choose the 5-day moving average of actual S&P values.
'We 'll start with P(1) and P(2) equal to the actual S&P 500 values.

'Then we'll use the magic equation, above, and run through a bunch of values
'for To and f and pick out the values that minimize the RMS error between the
'P(n) and the actual S&P values. (The Root-Mean-Square error is expressed as
'a percentage of the starting S&P value else an RMS error of 1.2345 is
'meaningless, eh?)

'After identifying the "best" choice (for To and f) we extrapolate, using the
'predicted values in the magic equation (for the P(n) and P0).

'We 're now out on a limb. Our only numbers are the predicted values. The actual
'S&P 500 values are history. We're looking into the future. Moving averages can
'only be calculated using our predictions. The future evolution of the S&P ...

'LIBRARY       : FINAN_ASSET
'GROUP         : WAVES
'ID            : 001



'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'------------------------------------------------------------------------------------
'REFERENCES    :
'------------------------------------------------------------------------------------
'Why Stock Markets Crash: Critical Events in Complex Financial Systems
'Didier Sornette <http://press.princeton.edu/titles/7341.html>

'http://www.gummy-stuff.org/Wave_Theory.htm
'http://www.gummy-stuff.org/crash.htm
'http://www.ess.ucla.edu/faculty/sornette/
'http://www.google.ca/search?q=Lomb+Periodogram&ie=UTF-8&hl=en&btnG=Google+Search&meta=
'------------------------------------------------------------------------------------
'--> Stock market crashes, Precursors and Replicas
'--> http://arxiv.org/PS_cache/cond-mat/pdf/9510/9510036v1.pdf
'--> http://www.eurekalert.org/pub_releases/2002-12/uoc--smc121402.php
'------------------------------------------------------------------------------------
'************************************************************************************
'************************************************************************************

Function ASSET_WAVES_CRASH_FUNC(ByVal TICKER_STR As Variant, _
Optional ByVal START_DATE As Date, _
Optional ByVal END_DATE As Date, _
Optional ByVal CRASH_LEVEL As Double = 0.2, _
Optional ByVal MA_PERIOD As Long = 5, _
Optional ByVal TN_FACTOR As Double = 3.3, _
Optional ByVal INTENSITY As Variant = 1, _
Optional ByVal PERIODS_FORWARD As Long = 213, _
Optional ByRef HOLIDAYS_RNG As Variant, _
Optional ByVal OUTPUT As Integer = 1)

'CRASH_LEVEL --> Your definition of a Crash

Dim g As Long
Dim h As Long
Dim i As Long
Dim j As Long
Dim k As Long
Dim l As Long

Dim NROWS As Long

Dim PI_VAL As Double

Dim ATEMP_SUM As Double
Dim BTEMP_SUM As Double
Dim CTEMP_SUM As Double
Dim DTEMP_SUM As Double

Dim TEMP_MATRIX As Variant

Dim CRASH_DATE As Date
Dim DATA_MATRIX As Variant

On Error GoTo ERROR_LABEL

If MA_PERIOD < 3 Then: GoTo ERROR_LABEL
If IsArray(TICKER_STR) = False Then
    DATA_MATRIX = YAHOO_HISTORICAL_DATA_SERIE_FUNC(TICKER_STR, START_DATE, END_DATE, _
                  "da", "DOHLCVA", False, True, True)
Else
    DATA_MATRIX = TICKER_STR
End If
PI_VAL = 3.14159265358979

'--------------------------------------------------------------------------------------
h = 2 'Moving Average Threshold
If MA_PERIOD <= h Then: GoTo ERROR_LABEL
h = MA_PERIOD * h
'--------------------------------------------------------------------------------------

NROWS = UBound(DATA_MATRIX, 1)
ReDim TEMP_MATRIX(0 To NROWS + PERIODS_FORWARD, 1 To 10)

TEMP_MATRIX(0, 1) = "DATE"
TEMP_MATRIX(0, 2) = "ACTUAL"

TEMP_MATRIX(0, 3) = "PREDICT"
TEMP_MATRIX(0, 4) = "INTS-FACTOR"
TEMP_MATRIX(0, 5) = "TN-FACTOR"

TEMP_MATRIX(0, 6) = "P0-LEVEL"
TEMP_MATRIX(0, 7) = "MA-ACTUAL"
TEMP_MATRIX(0, 8) = "MA-PREDICT"
TEMP_MATRIX(0, 9) = "RMS-ERROR: "
TEMP_MATRIX(0, 10) = "CRASH-DATE: "

'--------------------------------------------------------------------------

k = 0
ATEMP_SUM = 0: BTEMP_SUM = 0: CTEMP_SUM = 0: DTEMP_SUM = 0
For i = 1 To NROWS
    
    TEMP_MATRIX(i, 1) = DATA_MATRIX(i, 1)
    TEMP_MATRIX(i, 2) = DATA_MATRIX(i, 2)
        
    If i < MA_PERIOD Then
        TEMP_MATRIX(i, 6) = TEMP_MATRIX(i, 2)
        GoSub INTENSITY_LINE
        ATEMP_SUM = ATEMP_SUM + TEMP_MATRIX(i, 2)
        k = k + 1
        
        TEMP_MATRIX(i, 3) = TEMP_MATRIX(i, 2)
        
        BTEMP_SUM = BTEMP_SUM + TEMP_MATRIX(i, 3)
        TEMP_MATRIX(i, 7) = BTEMP_SUM / k
    
        CTEMP_SUM = CTEMP_SUM + TEMP_MATRIX(i, 2)
        TEMP_MATRIX(i, 8) = CTEMP_SUM / k
    
    
    ElseIf i = MA_PERIOD Then
        TEMP_MATRIX(i, 6) = ATEMP_SUM / k
        GoSub INTENSITY_LINE
        
        BTEMP_SUM = BTEMP_SUM + TEMP_MATRIX(i, 3)
        TEMP_MATRIX(i, 7) = BTEMP_SUM / (k + 1)
    
    
        CTEMP_SUM = CTEMP_SUM + TEMP_MATRIX(i, 2)
        TEMP_MATRIX(i, 8) = CTEMP_SUM / (k + 1)
    
    Else
        If i < h Then
            ATEMP_SUM = ATEMP_SUM + TEMP_MATRIX(i - 1, 2)
            k = k + 1
            TEMP_MATRIX(i, 6) = ATEMP_SUM / k
            GoSub INTENSITY_LINE
        
            BTEMP_SUM = BTEMP_SUM + TEMP_MATRIX(i, 3)
            TEMP_MATRIX(i, 7) = BTEMP_SUM / (k + 1)
        
            CTEMP_SUM = CTEMP_SUM + TEMP_MATRIX(i, 2)
            TEMP_MATRIX(i, 8) = CTEMP_SUM / (k + 1)
        
        Else
            If i = h Then
                ATEMP_SUM = 0
                l = i - MA_PERIOD
                For j = i - 1 To l Step -1
                    ATEMP_SUM = ATEMP_SUM + TEMP_MATRIX(j, 2)
                Next j
                TEMP_MATRIX(i, 6) = ATEMP_SUM / MA_PERIOD
                GoSub INTENSITY_LINE
           
                BTEMP_SUM = 0
                CTEMP_SUM = 0
                For j = i To l Step -1
                    BTEMP_SUM = BTEMP_SUM + TEMP_MATRIX(j, 3)
                    CTEMP_SUM = CTEMP_SUM + TEMP_MATRIX(j, 2)
                Next j
                TEMP_MATRIX(i, 7) = BTEMP_SUM / (MA_PERIOD + 1)
'                TEMP_MATRIX(i, 8) = CTEMP_SUM / (MA_PERIOD + 1)
                TEMP_MATRIX(i, 8) = (CTEMP_SUM + TEMP_MATRIX(l - 1, 2)) / (MA_PERIOD + 2)
                'Adjustment!
                
                g = h + MA_PERIOD
            Else
                If i < g Then
                    l = i - MA_PERIOD - 1
                    ATEMP_SUM = ATEMP_SUM - TEMP_MATRIX(l, 2)
                    ATEMP_SUM = ATEMP_SUM + TEMP_MATRIX(i - 1, 2)
                    TEMP_MATRIX(i, 6) = ATEMP_SUM / (MA_PERIOD)
                    GoSub INTENSITY_LINE
                
                    BTEMP_SUM = BTEMP_SUM - TEMP_MATRIX(l, 3)
                    BTEMP_SUM = BTEMP_SUM + TEMP_MATRIX(i, 3)
                    TEMP_MATRIX(i, 7) = BTEMP_SUM / (MA_PERIOD + 1)
                
                    CTEMP_SUM = CTEMP_SUM - TEMP_MATRIX(l, 2)
                    CTEMP_SUM = CTEMP_SUM + TEMP_MATRIX(i, 2)
                    TEMP_MATRIX(i, 8) = CTEMP_SUM / (MA_PERIOD + 1)
                Else
                    If i <> g Then
                        l = i - MA_PERIOD
                        ATEMP_SUM = ATEMP_SUM - TEMP_MATRIX(l, 2)
                        ATEMP_SUM = ATEMP_SUM + TEMP_MATRIX(i - 1, 2)
                        
                        TEMP_MATRIX(i, 6) = ATEMP_SUM / (MA_PERIOD - 1)
                        GoSub INTENSITY_LINE
                    
                        BTEMP_SUM = BTEMP_SUM - TEMP_MATRIX(l, 3)
                        BTEMP_SUM = BTEMP_SUM + TEMP_MATRIX(i, 3)
                        TEMP_MATRIX(i, 7) = BTEMP_SUM / MA_PERIOD
                    
                        CTEMP_SUM = CTEMP_SUM - TEMP_MATRIX(l, 2)
                        CTEMP_SUM = CTEMP_SUM + TEMP_MATRIX(i, 2)
                        TEMP_MATRIX(i, 8) = CTEMP_SUM / MA_PERIOD
                    
                    Else
                        l = i - MA_PERIOD - 1
                        
                        ATEMP_SUM = ATEMP_SUM + TEMP_MATRIX(i - 1, 2)
                        TEMP_MATRIX(i, 6) = ATEMP_SUM / (MA_PERIOD + 1)
                        GoSub INTENSITY_LINE
                        ATEMP_SUM = ATEMP_SUM - TEMP_MATRIX(l, 2)
                        ATEMP_SUM = ATEMP_SUM - TEMP_MATRIX(l + 1, 2)
                                        
                        BTEMP_SUM = BTEMP_SUM - TEMP_MATRIX(l, 3)
                        BTEMP_SUM = BTEMP_SUM + TEMP_MATRIX(i, 3)
                        TEMP_MATRIX(i, 7) = BTEMP_SUM / (MA_PERIOD + 1)
                        BTEMP_SUM = BTEMP_SUM - TEMP_MATRIX(l + 1, 3)
                    
                        CTEMP_SUM = CTEMP_SUM - TEMP_MATRIX(l, 2)
                        CTEMP_SUM = CTEMP_SUM + TEMP_MATRIX(i, 2)
                        TEMP_MATRIX(i, 8) = CTEMP_SUM / (MA_PERIOD + 1)
                        CTEMP_SUM = CTEMP_SUM - TEMP_MATRIX(l + 1, 2)
                    End If
                End If
            End If
        End If
    End If
    
    TEMP_MATRIX(i, 9) = TEMP_MATRIX(i, 7) - TEMP_MATRIX(i, 8)
    DTEMP_SUM = DTEMP_SUM + TEMP_MATRIX(i, 9) ^ 2
    TEMP_MATRIX(i, 10) = ""
Next i

DTEMP_SUM = Sqr(DTEMP_SUM / (NROWS - 2)) / TEMP_MATRIX(1, 2)
If OUTPUT = 2 Then
    ASSET_WAVES_CRASH_FUNC = DTEMP_SUM
    Exit Function
End If
'--------------------------------------------------------------------------------------------
TEMP_MATRIX(0, 9) = TEMP_MATRIX(0, 9) & Format(DTEMP_SUM, "0.00%")
'--------------------------------------------------------------------------------------------
i = NROWS
ATEMP_SUM = 0
l = i - MA_PERIOD + 1
For j = i - 1 To l Step -1
    ATEMP_SUM = ATEMP_SUM + TEMP_MATRIX(j, 3)
Next j

CRASH_DATE = 0
For i = (NROWS + 1) To (NROWS + PERIODS_FORWARD)
    TEMP_MATRIX(i, 1) = WORKDAY2_FUNC(TEMP_MATRIX(i - 1, 1), 1, HOLIDAYS_RNG)
    TEMP_MATRIX(i, 2) = ""

    l = i - MA_PERIOD
    ATEMP_SUM = ATEMP_SUM - TEMP_MATRIX(l, 3)
    ATEMP_SUM = ATEMP_SUM + TEMP_MATRIX(i - 1, 3)
                        
    TEMP_MATRIX(i, 6) = ATEMP_SUM / (MA_PERIOD - 1)
    GoSub INTENSITY_LINE

    BTEMP_SUM = BTEMP_SUM - TEMP_MATRIX(l, 3)
    BTEMP_SUM = BTEMP_SUM + TEMP_MATRIX(i, 3)
    TEMP_MATRIX(i, 7) = BTEMP_SUM / MA_PERIOD
    
    TEMP_MATRIX(i, 8) = ""
    TEMP_MATRIX(i, 9) = ""

    If (TEMP_MATRIX(i, 3) / TEMP_MATRIX(i - 1, 3)) < (1 - CRASH_LEVEL) Then
        TEMP_MATRIX(i, 10) = TEMP_MATRIX(i, 1)
        If CRASH_DATE = 0 Then
            CRASH_DATE = TEMP_MATRIX(i, 1)
            If OUTPUT >= 1 Then: Exit For
            TEMP_MATRIX(0, 10) = TEMP_MATRIX(0, 10) & Format(CRASH_DATE, "mmm dd, yyyy")
        End If
    Else
        TEMP_MATRIX(i, 10) = ""
    End If
Next i

'--------------------------------------------------------------------------------
Select Case OUTPUT
'--------------------------------------------------------------------------------
    Case 0 'LOMB PERIODOGRAM
'--------------------------------------------------------------------------------
        ASSET_WAVES_CRASH_FUNC = TEMP_MATRIX
'--------------------------------------------------------------------------------
    Case 1 'Crash Date
'--------------------------------------------------------------------------------
        ASSET_WAVES_CRASH_FUNC = CRASH_DATE
'--------------------------------------------------------------------------------
    Case Else '>2 'DTEMP_SUM = RMS ERROR
'--------------------------------------------------------------------------------
        ASSET_WAVES_CRASH_FUNC = Array(CRASH_DATE, DTEMP_SUM)
End Select
'--------------------------------------------------------------------------------
Exit Function
'-------------------------------------------------------------------------------
INTENSITY_LINE:
'-------------------------------------------------------------------------------
    If i <> 1 Then
        TEMP_MATRIX(i, 5) = TEMP_MATRIX(i - 1, 5) * INTENSITY
        TEMP_MATRIX(i, 4) = 2 - (2 * PI_VAL / TEMP_MATRIX(i, 5)) ^ 2
    Else
        TEMP_MATRIX(i, 5) = TN_FACTOR
        TEMP_MATRIX(i, 4) = 2 - (2 * PI_VAL / TEMP_MATRIX(i, 5)) ^ 2
    End If
    
    If i < MA_PERIOD Then
        TEMP_MATRIX(i, 3) = TEMP_MATRIX(i, 2)
    Else
        TEMP_MATRIX(i, 3) = TEMP_MATRIX(i, 4) * TEMP_MATRIX(i - 1, 3) - _
                            TEMP_MATRIX(i - 2, 3) + (2 - TEMP_MATRIX(i, 4)) * _
                            TEMP_MATRIX(i, 6)
    End If
'-------------------------------------------------------------------------------
Return
'-------------------------------------------------------------------------------
ERROR_LABEL:
ASSET_WAVES_CRASH_FUNC = PUB_EPSILON
End Function


'************************************************************************************
'************************************************************************************



'************************************************************************************
'************************************************************************************
'FUNCTION      : ASSET_WAVES_CRASH_OPTIMIZER_FUNC
'DESCRIPTION   : That 's when the crash is predicted to occur and, after finding
'some "good" log-periodic fit, they try to extract the time of the
'crash, namely Tc. The ritual is often associated with the Lomb Periodogram.
'Method of spectral analysis for unevenly sampled series, such as the beat-to-beat
'series of an Asset Price.

'LIBRARY       : FINAN_ASSET
'GROUP         : WAVES
'ID            : 002

'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'REFERENCES    : Why Stock Markets Crash: Critical Events in Complex Financial Systems
'Didier Sornette <http://press.princeton.edu/titles/7341.html>
'************************************************************************************
'************************************************************************************

Function ASSET_WAVES_CRASH_OPTIMIZER_FUNC(ByVal TICKER_STR As Variant, _
Optional ByVal START_DATE As Date, _
Optional ByVal END_DATE As Date, _
Optional ByRef PARAM_RNG As Variant, _
Optional ByVal START_CRASH_LEVEL As Double = 0.1, _
Optional ByVal END_CRASH_LEVEL As Double = 0.2, _
Optional ByVal DELTA_CRASH_LEVEL As Double = 0.01, _
Optional ByVal START_MA_PERIOD As Long = 3, _
Optional ByVal END_MA_PERIOD As Double = 4, _
Optional ByVal START_TN As Double = 3, _
Optional ByVal END_TN As Double = 6, _
Optional ByVal DELTA_TN As Double = 0.25, _
Optional ByVal nLOOPS As Long = 10, _
Optional ByVal epsilon As Double = 1.99999999999978E-04, _
Optional ByVal PERIODS_FORWARD As Long = 213, _
Optional ByRef HOLIDAYS_RNG As Variant)
'Optional ByVal epsilon As Double = 1 - 0.9998
Dim hh As Double
Dim ii As Double
Dim jj As Long
Dim kk As Long
Dim ll As Long

Dim TEMP_MULT As Double
Dim TEMP_INTENSITY As Double

Dim CRASH_DATE As Date
Dim ERROR_VAL As Double

Dim PARAM_VECTOR As Variant
Dim TEMP_GROUP As Variant
Dim TEMP_MATRIX As Variant

On Error GoTo ERROR_LABEL

If IsArray(TICKER_STR) = False Then
    PUB_DATA_MATRIX = YAHOO_HISTORICAL_DATA_SERIE_FUNC(TICKER_STR, START_DATE, END_DATE, _
                      "d", "DOHLCVA", False, True, True)
Else
    PUB_DATA_MATRIX = TICKER_STR
End If
    
If IsArray(PARAM_RNG) = True Then
    PARAM_VECTOR = PARAM_RNG
    If UBound(PARAM_VECTOR, 1) = 1 Then: PARAM_VECTOR = PARAM_RNG
    If UBound(PARAM_VECTOR, 1) <> 3 Then: GoTo ERROR_LABEL
    'PARAM_VECTOR(1,1) --> INITIAL_MA_PERIOD As Long = 5
    'PARAM_VECTOR(2,1) --> INITIAL_TN_FACTOR As Double = 3.75
    'PARAM_VECTOR(3,1) --> INITIAL_INTENSITY As Double = 0.999939995799524
    PUB_CRASH_LEVEL = START_CRASH_LEVEL
    PUB_PERIODS_FORWARD = PERIODS_FORWARD
    PUB_HOLIDAYS_RNG = HOLIDAYS_RNG
    
    PARAM_VECTOR = NELDER_MEAD_OPTIMIZATION3_FUNC("ASSET_WAVES_CRASH_OBJ_FUNC", _
                   PARAM_VECTOR, 1000, 10 ^ -15)
    If IsArray(PARAM_VECTOR) = False Then: GoTo ERROR_LABEL
    PARAM_VECTOR(1, 1) = CLng(PARAM_VECTOR(1, 1))
    ASSET_WAVES_CRASH_OPTIMIZER_FUNC = PARAM_VECTOR
    Exit Function
End If

kk = 0
ReDim TEMP_GROUP(1 To 1)

For hh = START_CRASH_LEVEL To END_CRASH_LEVEL Step DELTA_CRASH_LEVEL
    For ll = START_MA_PERIOD To END_MA_PERIOD
        TEMP_MULT = (1 - epsilon) ^ (1 / nLOOPS)
        For ii = START_TN To END_TN Step DELTA_TN
            TEMP_INTENSITY = 1
            For jj = 1 To nLOOPS
                TEMP_MATRIX = ASSET_WAVES_CRASH_FUNC(PUB_DATA_MATRIX, _
                , , hh, ll, ii, TEMP_INTENSITY, PERIODS_FORWARD, HOLIDAYS_RNG, 3)
                If IsArray(TEMP_MATRIX) = True Then
                    CRASH_DATE = TEMP_MATRIX(LBound(TEMP_MATRIX))
                    ERROR_VAL = TEMP_MATRIX(UBound(TEMP_MATRIX))
                    If ERROR_VAL <= 1 And CRASH_DATE <> 0 Then
                        kk = kk + 1
                        ReDim Preserve TEMP_GROUP(1 To kk)
                        TEMP_GROUP(kk) = Array(ll, hh, CRASH_DATE, ERROR_VAL, ii, TEMP_INTENSITY)
                    End If
                End If
                TEMP_INTENSITY = TEMP_INTENSITY * TEMP_MULT
            Next jj
        Next ii
    Next ll
Next hh

kk = UBound(TEMP_GROUP)
ReDim TEMP_MATRIX(0 To kk, 1 To 6)
TEMP_MATRIX(0, 1) = "MA-PERIOD"
TEMP_MATRIX(0, 2) = "CRASH-LEVEL"
TEMP_MATRIX(0, 3) = "CRASH-DATE"
TEMP_MATRIX(0, 4) = "ERROR-VAL"
TEMP_MATRIX(0, 5) = "TN-FACTOR"
TEMP_MATRIX(0, 6) = "INTENSITY"
For jj = 1 To kk
    TEMP_MATRIX(jj, 1) = TEMP_GROUP(jj)(UBound(TEMP_GROUP(jj)) - 5)
    TEMP_MATRIX(jj, 2) = TEMP_GROUP(jj)(UBound(TEMP_GROUP(jj)) - 4)
    TEMP_MATRIX(jj, 3) = TEMP_GROUP(jj)(UBound(TEMP_GROUP(jj)) - 3)
    TEMP_MATRIX(jj, 4) = TEMP_GROUP(jj)(UBound(TEMP_GROUP(jj)) - 2)
    TEMP_MATRIX(jj, 5) = TEMP_GROUP(jj)(UBound(TEMP_GROUP(jj)) - 1)
    TEMP_MATRIX(jj, 6) = TEMP_GROUP(jj)(UBound(TEMP_GROUP(jj)))
Next jj

ASSET_WAVES_CRASH_OPTIMIZER_FUNC = TEMP_MATRIX

Exit Function
ERROR_LABEL:
ASSET_WAVES_CRASH_OPTIMIZER_FUNC = Err.number
End Function

Function ASSET_WAVES_CRASH_OBJ_FUNC(ByRef PARAM_VECTOR As Variant)

On Error GoTo ERROR_LABEL

ASSET_WAVES_CRASH_OBJ_FUNC = ASSET_WAVES_CRASH_FUNC(PUB_DATA_MATRIX, _
    , , PUB_CRASH_LEVEL, PARAM_VECTOR(1, 1), PARAM_VECTOR(2, 1), PARAM_VECTOR(3, 1), _
    PUB_PERIODS_FORWARD, PUB_HOLIDAYS_RNG, 2)

Exit Function
ERROR_LABEL:
ASSET_WAVES_CRASH_OBJ_FUNC = PUB_EPSILON
End Function
