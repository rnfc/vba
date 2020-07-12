Attribute VB_Name = "FINAN_ASSET_SIMUL_LOG_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.
                    
'http://www.gummy-stuff.org/bolli-bands.htm

Function ASSET_LOG_PRICES_DISTRIBUTION_FUNC(ByVal TICKER_STR As Variant, _
Optional ByVal START_DATE As Date, _
Optional ByVal END_DATE As Date, _
Optional ByVal MA_PERIOD As Long = 180, _
Optional ByVal NO_PERIODS As Long = 20, _
Optional ByVal LOWER_DELTA_PRICE As Double = 2, _
Optional ByVal UPPER_DELTA_PRICE As Double = 5, _
Optional ByVal NBINS As Long = 30, _
Optional ByVal NSIZE As Long = 300, _
Optional ByVal OUTPUT As Integer = 0)

'NO_PERIODS:  days into the future

Dim i As Long
Dim j As Long
Dim k As Long
Dim NROWS As Long

Dim L_VAL As Double
Dim U_VAL As Double

Dim P1_VAL As Double
Dim P2_VAL As Double

Dim M1_VAL As Double
Dim S1_VAL As Double

Dim M2_VAL As Double
Dim S2_VAL As Double

Dim M3_VAL As Double
Dim S3_VAL As Double

Dim P1_STR As String
Dim P2_STR As String

Dim MIN_PRICE As Double
Dim MAX_PRICE As Double
Dim DELTA_PRICE As Double
Dim TEMP_PRICE As Double
Dim CURRENT_PRICE As Double

Dim TEMP_MATRIX As Variant
Dim RETURNS_VECTOR As Variant
Dim DATA_MATRIX As Variant

Dim tolerance As Double

On Error GoTo ERROR_LABEL

tolerance = 0.0001

If IsArray(TICKER_STR) = False Then
    DATA_MATRIX = YAHOO_HISTORICAL_DATA_SERIE_FUNC(TICKER_STR, _
                  START_DATE, END_DATE, "DAILY", "DOHLCVA", False, _
                  True, True)
Else
    DATA_MATRIX = TICKER_STR
End If
NROWS = UBound(DATA_MATRIX, 1)

CURRENT_PRICE = DATA_MATRIX(NROWS, 7)
L_VAL = CURRENT_PRICE + LOWER_DELTA_PRICE
U_VAL = CURRENT_PRICE + UPPER_DELTA_PRICE

ReDim RETURNS_VECTOR(1 To NROWS, 1 To 1)
i = 1: k = 0
RETURNS_VECTOR(i, 1) = DATA_MATRIX(i, 5) / DATA_MATRIX(i, 2) - 1
For i = 2 To NROWS
    RETURNS_VECTOR(i, 1) = DATA_MATRIX(i, 7) / DATA_MATRIX(i - 1, 7) - 1
    If i >= (NROWS - MA_PERIOD) And i < NROWS Then
        M1_VAL = M1_VAL + RETURNS_VECTOR(i, 1)
    End If
    j = NROWS - i + 1
    If i <= (NSIZE + 1) Then
        If ((DATA_MATRIX(j, 7) > DATA_MATRIX(j - NO_PERIODS, 7) + LOWER_DELTA_PRICE And _
             DATA_MATRIX(j, 7) < DATA_MATRIX(j - NO_PERIODS, 7) + UPPER_DELTA_PRICE)) Then: k = k + 1
    End If
Next i
P1_VAL = k / NSIZE
P1_STR = "Probability the asset will lie between " & _
            Format(L_VAL, "0.00") & " and " & Format(U_VAL, "0.00") & _
            " (" & Format(NO_PERIODS, "0") & " periods into the future ) is " & _
            Format(P1_VAL, "0.00%") 'discrete

M1_VAL = (M1_VAL / MA_PERIOD)

S1_VAL = 0
For i = NROWS - MA_PERIOD To NROWS - 1
    S1_VAL = S1_VAL + (RETURNS_VECTOR(i, 1) - M1_VAL) ^ 2
Next i
Erase RETURNS_VECTOR

S1_VAL = (S1_VAL / MA_PERIOD) ^ 0.5
M1_VAL = M1_VAL + 1

S2_VAL = Sqr((M1_VAL ^ 2 + S1_VAL ^ 2) ^ NO_PERIODS - M1_VAL ^ (2 * NO_PERIODS))
M2_VAL = M1_VAL ^ NO_PERIODS

S3_VAL = Sqr(Log(1 + S2_VAL ^ 2 / M2_VAL ^ 2))
M3_VAL = Log(M2_VAL) - S3_VAL ^ 2 / 2

P2_VAL = NORMSDIST_FUNC(Log(U_VAL / CURRENT_PRICE), M3_VAL, S3_VAL, 0) - _
         NORMSDIST_FUNC(Log(L_VAL / CURRENT_PRICE), M3_VAL, S3_VAL, 0)

P2_STR = "Probability the asset will lie between " & _
            Format(L_VAL, "0.00") & " and " & Format(U_VAL, "0.00") & _
            " (" & Format(NO_PERIODS, "0") & " periods into the future ) is " & _
            Format(P2_VAL, "0.00%")

If OUTPUT > 0 Then
    ASSET_LOG_PRICES_DISTRIBUTION_FUNC = Array(UCase(P1_STR), UCase(P2_STR))
'    ASSET_LOG_PRICES_DISTRIBUTION_FUNC = Array(P1_VAL, UCase(P1_STR), _
                                               P2_VAL, UCase(P2_STR), _
                                               M3_VAL, S3_VAL) 'log normal
    Exit Function
End If

MIN_PRICE = NORMSINV_FUNC(tolerance, M3_VAL, S3_VAL, 0) * CURRENT_PRICE
MAX_PRICE = NORMSINV_FUNC(1 - tolerance, M3_VAL, S3_VAL, 0) * CURRENT_PRICE
DELTA_PRICE = (MAX_PRICE - MIN_PRICE) / NBINS

ReDim TEMP_MATRIX(0 To NBINS, 1 To 3)
TEMP_MATRIX(0, 1) = "DELTA"
TEMP_MATRIX(0, 2) = "PRICE"
TEMP_MATRIX(0, 3) = "PROBABILITY THAT PRICE IS LESS THAN " & Format(U_VAL, "0.00") & _
                     ", IN " & Format(NO_PERIODS, "0") & " PERIODS INTO THE FUTURE"

TEMP_PRICE = MIN_PRICE
For i = 1 To NBINS
    TEMP_MATRIX(i, 1) = TEMP_PRICE
    TEMP_MATRIX(i, 2) = TEMP_PRICE + CURRENT_PRICE
    TEMP_MATRIX(i, 3) = NORMSDIST_FUNC(Log(1 + TEMP_MATRIX(i, 1) / CURRENT_PRICE), _
                        M3_VAL, S3_VAL, 0)
    
    TEMP_PRICE = TEMP_PRICE + DELTA_PRICE
Next i

ASSET_LOG_PRICES_DISTRIBUTION_FUNC = TEMP_MATRIX

Exit Function
ERROR_LABEL:
ASSET_LOG_PRICES_DISTRIBUTION_FUNC = Err.number
End Function


'Returns the inverse of the lognormal cumulative distribution function of Price(x),
'where ln(x) is normally distributed with parameters mean and standard_dev.
'If p = LOGNORMDIST(x,...) then LOGINV(p,...) = x.
'Use the lognormal distribution to analyze logarithmically transformed
'historical data.

'MENA_VAL --> is the mean of ln(x).
'SIGMA_VAL --> is the standard deviation of ln(x).

Function ASSET_LOG_MOMENTS_SAMPLING_FUNC( _
Optional ByVal CURRENT_PRICE As Double = 10, _
Optional ByVal MEAN_VAL As Double = 0.002, _
Optional ByVal SIGMA_VAL As Double = 0.03, _
Optional ByVal NO_PERIODS As Long = 1000, _
Optional ByVal OUTPUT As Integer = 3)

Dim i As Long
Dim TEMP_MATRIX As Variant 'Generate New Simulated Prices (log normal)

On Error GoTo ERROR_LABEL

ReDim TEMP_MATRIX(0 To NO_PERIODS + 1, 1 To 4)
TEMP_MATRIX(0, 1) = "PERIOD"
TEMP_MATRIX(0, 2) = "GAIN"
TEMP_MATRIX(0, 3) = "PRICE"
TEMP_MATRIX(0, 4) = "RETURN"

i = 1
TEMP_MATRIX(i, 1) = i - 1
TEMP_MATRIX(i, 2) = ""
TEMP_MATRIX(i, 3) = CURRENT_PRICE
TEMP_MATRIX(i, 4) = ""

Randomize
For i = 2 To NO_PERIODS + 1
    TEMP_MATRIX(i, 1) = i - 1

    'RND --> Probability --> is a probability associated with the lognormal distribution.
    
    TEMP_MATRIX(i, 2) = Exp(MEAN_VAL + SIGMA_VAL * NORMSINV_FUNC(Rnd, 0, 1, 0)) 'log normal
    TEMP_MATRIX(i, 3) = TEMP_MATRIX(i - 1, 3) * TEMP_MATRIX(i, 2)
    TEMP_MATRIX(i, 4) = Log(TEMP_MATRIX(i, 3) / TEMP_MATRIX(i - 1, 3))
Next i

'--------------------------------------------------------------------------------
Select Case OUTPUT
'--------------------------------------------------------------------------------
Case 0 'generate Future Price Distribution (with the same simulated data)
'--------------------------------------------------------------------------------
    MEAN_VAL = 0: SIGMA_VAL = 0
    ASSET_LOG_MOMENTS_SAMPLING_FUNC = TEMP_MATRIX
'--------------------------------------------------------------------------------
Case 1 'r & s using Logarithimic
'--------------------------------------------------------------------------------
    MEAN_VAL = 0: SIGMA_VAL = 0
    For i = 2 To NO_PERIODS + 1
        MEAN_VAL = MEAN_VAL + TEMP_MATRIX(i, 4)
    Next i
    MEAN_VAL = MEAN_VAL / NO_PERIODS
    For i = 2 To NO_PERIODS + 1
        SIGMA_VAL = SIGMA_VAL + (TEMP_MATRIX(i, 4) - MEAN_VAL) ^ 2
    Next i
    SIGMA_VAL = (SIGMA_VAL / NO_PERIODS) ^ 0.5
    ASSET_LOG_MOMENTS_SAMPLING_FUNC = Array(MEAN_VAL, SIGMA_VAL)
'--------------------------------------------------------------------------------
Case Else
'--------------------------------------------------------------------------------
    MEAN_VAL = 0: SIGMA_VAL = 0
    For i = 2 To NO_PERIODS + 1
        MEAN_VAL = MEAN_VAL + TEMP_MATRIX(i, 2)
    Next i
    MEAN_VAL = MEAN_VAL / NO_PERIODS
    For i = 2 To NO_PERIODS + 1
        SIGMA_VAL = SIGMA_VAL + (TEMP_MATRIX(i, 2) - MEAN_VAL) ^ 2
    Next i
    SIGMA_VAL = (SIGMA_VAL / NO_PERIODS) ^ 0.5
    
    If OUTPUT = 2 Then 'r & s using simple Mean
        MEAN_VAL = MEAN_VAL - 1
        ASSET_LOG_MOMENTS_SAMPLING_FUNC = Array(MEAN_VAL, SIGMA_VAL)
    Else 'r & s using Annualized
        MEAN_VAL = (TEMP_MATRIX(NO_PERIODS + 1, 3) / TEMP_MATRIX(1, 3)) ^ (1 / NO_PERIODS) - 1
        ASSET_LOG_MOMENTS_SAMPLING_FUNC = Array(MEAN_VAL, SIGMA_VAL)
    End If
'--------------------------------------------------------------------------------
End Select
'--------------------------------------------------------------------------------

Exit Function
ERROR_LABEL:
ASSET_LOG_MOMENTS_SAMPLING_FUNC = Err.number
End Function
