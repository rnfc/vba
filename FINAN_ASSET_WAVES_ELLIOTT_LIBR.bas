Attribute VB_Name = "FINAN_ASSET_WAVES_ELLIOTT_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.

'R. N. Elliott (the U.S. State Department appointed him Chief Accountant for Nicaragua),
'was stricken with a debilitating case of pernicious anemia, became bedridden and was
'forced into retirement at age 58 ... He then studied yearly, monthly, weekly, daily,
'hourly and half-hourly charts of the various market indices covering 75 years of
'historical data.

'According to Elliott, event sequences occur in waves, from the very long (lasting hundreds
'of years, called Grand Super Cycle) to the very short (called sub-minuette, lasting maybe
'a few minutes) and, if you look carefully at the larger waves, you find smaller, or shorter,
'versions embedded within and ...

'Smaller waves inside the bigger waves?
'Yes, apparently, and they're similar in structure. Sort of like fractals where a microscopic
'examination of fractals reveals they are made up of similar, smaller designs.

'Anyway, I started looking for Elliott Waves within the S&P 500 ... those which have a particular
'structure: FIVE up cycles followed by TWO down cycles:

'Well, this is what I did:
'1. I look at the S&P closing value in the month of January, 1930 and compare it to all
'monthly closing values from two years before (Jan, 1928) to two years after (Jan, 1932).

'2. I then move to Feb, 1930 and compare the closing value to the other values in a 4-year
'period centered on Feb/30.

'3. I then move to Mar, 1930 and compare the closing value to the other values in a 4-year
'period centered on Mar/30.

'4. I then move to Apr, 1930 and ...

'I figure that, if there are Elliott Waves of the 5up/3dn variety, I should look at those
'months where the S&P close is the largest value in the 4-year time period.

Function ASSET_ELLIOTT_WAVES_FUNC(ByVal TICKER_STR As Variant, _
Optional ByVal START_DATE As Date, _
Optional ByVal END_DATE As Date, _
Optional ByVal MA_PERIOD As Long = 25, _
Optional ByVal WAVE_NO As Long = 2, _
Optional ByVal NSIZE As Long = 50, _
Optional ByVal OUTPUT As Integer = 1)

Dim h As Long
Dim i As Long
Dim j As Long
Dim k As Long
Dim l As Long 'Start Wave
Dim NROWS As Long
Dim NCOLUMNS As Long

Dim MAX1_VAL As Double
Dim MAX2_VAL As Double
Dim INDEX_ARR() As Long 'Waves Chooser
Dim GOLDEN_RATIO As Double
Dim DATA_MATRIX As Variant

On Error GoTo ERROR_LABEL

If IsArray(TICKER_STR) = False Then
    DATA_MATRIX = YAHOO_HISTORICAL_DATA_SERIE_FUNC(TICKER_STR, START_DATE, END_DATE, _
                  "d", "DHLC", False, True, True)
Else
    DATA_MATRIX = TICKER_STR
End If
NROWS = UBound(DATA_MATRIX, 1)
NCOLUMNS = UBound(DATA_MATRIX, 2)

k = 1
ReDim INDEX_ARR(1 To k)
For i = MA_PERIOD To NROWS
    MAX1_VAL = -2 ^ 52: MAX2_VAL = MAX1_VAL
    For j = 1 To MA_PERIOD - 1
        If j < MA_PERIOD - 1 Then
            If DATA_MATRIX(i - j, 4) > MAX1_VAL Then MAX1_VAL = DATA_MATRIX(i - j, 4)
        End If
        If i + j <= NROWS Then
            If DATA_MATRIX(i + j, 4) > MAX2_VAL Then MAX2_VAL = DATA_MATRIX(i + j, 4)
        End If
    Next j
    If DATA_MATRIX(i, 4) > MAX1_VAL And DATA_MATRIX(i, 4) > MAX2_VAL Then
        If i + NSIZE < NROWS Then
            ReDim Preserve INDEX_ARR(1 To k)
            INDEX_ARR(k) = i
            k = k + 1
        End If
    End If
Next i
k = k - 1

'--------------------------------------------------------------------------------
'Ralph Nelson Elliott (the U.S. State Department appointed him Chief Accountant
'for Nicaragua), was stricken with a debilitating case of pernicious anemia,
'became bedridden and was forced into retirement at age 58 ...
'He then studied yearly, monthly, weekly, daily, hourly and half-hourly charts
'of the various market indices covering 75 years of historical data.

'According to Elliott, event sequences occur in waves, from the very long
'(lasting hundreds of years, called Grand Super Cycle) to the very short
'(called sub-minuette, lasting maybe a few minutes) and, if you look carefully
'at the larger waves, you find smaller, or shorter, versions embedded within and ...
'Smaller waves inside the bigger waves, and they're similar in structure. Sort
'of like fractals where a microscopic examination of fractals reveals they
'are made up of similar, smaller designs.
'-------------------------------------------------------------------------------
Select Case OUTPUT
'-------------------------------------------------------------------------------------
Case 0 'Perfect - 5up/3dn Waves
'-------------------------------------------------------------------------------------
    If (WAVE_NO + 1) > k Then: GoTo ERROR_LABEL 'not enough waves
    '(WAVE_NO + 1) --> Max Location

    GOLDEN_RATIO = (1 + Sqr(5)) / 2
    h = INDEX_ARR(WAVE_NO + 1)
    l = h - MA_PERIOD 'Start Wave
    'l = l + NSIZE 'End Wave
    'If l < 1 Then reduce MA_PERIOD
    
    ReDim TEMP_MATRIX(0 To NSIZE, 1 To NCOLUMNS + 2)
    TEMP_MATRIX(0, 1) = "DATE"
    TEMP_MATRIX(0, 2) = "HIGH"
    TEMP_MATRIX(0, 3) = "LOW"
    TEMP_MATRIX(0, 4) = "CLOSE"
    
    TEMP_MATRIX(0, 5) = "RATIOS: " & _
        Format(DATA_MATRIX(h, NCOLUMNS), "0.00") & " @ " & _
        Format(DATA_MATRIX(h, 1), "mmm dd, yyyy")
    TEMP_MATRIX(0, 6) = "FIBONACCI: " & Format(GOLDEN_RATIO, "0.0000")
    
    For i = 1 To NSIZE
        For j = 1 To NCOLUMNS: TEMP_MATRIX(i, j) = DATA_MATRIX(l, j): Next j
        l = l + 1
        TEMP_MATRIX(i, 5) = DATA_MATRIX(h, NCOLUMNS) / TEMP_MATRIX(i, NCOLUMNS)
        TEMP_MATRIX(i, 6) = GOLDEN_RATIO
    Next i
    ASSET_ELLIOTT_WAVES_FUNC = TEMP_MATRIX
'-------------------------------------------------------------------------------------
Case Else 'Perfect
'-------------------------------------------------------------------------------------
    ReDim TEMP_MATRIX(0 To k, 1 To 6)
    TEMP_MATRIX(0, 1) = "MAX LOCATION: DATE"
    TEMP_MATRIX(0, 2) = "MAX LOCATION: CLOSING PRICE"
    
    TEMP_MATRIX(0, 3) = "START WAVE: DATE"
    TEMP_MATRIX(0, 4) = "START WAVE: CLOSING PRICE"
    
    TEMP_MATRIX(0, 5) = "END WAVE: DATE"
    TEMP_MATRIX(0, 6) = "END WAVE: CLOSING PRICE"
    
    For i = 1 To k
        l = INDEX_ARR(i)
        TEMP_MATRIX(i, 1) = DATA_MATRIX(l, 1)
        TEMP_MATRIX(i, 2) = DATA_MATRIX(l, 4)
        
        l = l - MA_PERIOD
        TEMP_MATRIX(i, 3) = DATA_MATRIX(l, 1)
        TEMP_MATRIX(i, 4) = DATA_MATRIX(l, 4)
        
        l = l + NSIZE - 1
        TEMP_MATRIX(i, 5) = DATA_MATRIX(l, 1)
        TEMP_MATRIX(i, 6) = DATA_MATRIX(l, 4)
    Next i
    ASSET_ELLIOTT_WAVES_FUNC = TEMP_MATRIX
'-------------------------------------------------------------------------------------
End Select
'-------------------------------------------------------------------------------------

Exit Function
ERROR_LABEL:
ASSET_ELLIOTT_WAVES_FUNC = Err.number
End Function
