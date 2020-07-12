Attribute VB_Name = "FINAN_ASSET_MOMENTS_SPLIT_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.


'When a company had its stocks split N to 1 (N is usually 2, but
'can be a fraction like 0.5 as well. When N is smaller than 1, it is
'called a reverse split) with an ex-date of T, all the prices before T
'need to be multiplied by 1/N. Similarly, when a company issued a
'dividend $d per share with an ex-date of T, all the prices before T
'need to be multiplied by the number (Close(T – 1) – d)/Close(T –
'1), where Close(T – 1) is the closing price of the trading day before
'T. Notice that I adjust the historical prices by a multiplier instead
'of subtracting $d so that the historical daily returns will remain the
'same pre- and postadjustment. This is the way Yahoo! Finance adjusts
'its historical data, and is the most common way. (If you adjust
'by subtracting $d instead, the historical daily changes in prices will
'be the same pre- and postadjustment, but not the daily returns.) If
'the historical data are not adjusted, you will find a drop in price
'at the ex-date’s market open from previous day’s close (apart from
'normal market fluctuation), which may trigger an erroneous trading
'signal.

'http://home.dacor.net/norton/finance-math/adjustedClosingPrices.pdf

Function ASSET_ADJUSTING_SPLITS_DIVIDENDS_FUNC(ByRef DATA_RNG As Variant, _
Optional ByRef SPLITS_RNG As Variant, _
Optional ByRef DIVIDENDS_RNG As Variant)

Dim i As Long
Dim k As Long
Dim l As Long

Dim NROWS As Long

Dim DATE_STR As String 'KeyStr
Dim MULTIPLIER_VAL As Double

Dim SPLITS_VECTOR As Variant 'Dates/N
Dim DIVIDENDS_VECTOR As Variant 'Ex-Date/Dividend/PrevClose
Dim DATA_MATRIX As Variant 'Date/Open/High/Low/Close --> Descending Order

Dim COLLECTION_OBJ As New Collection

On Error Resume Next

DATA_MATRIX = DATA_RNG
NROWS = UBound(DATA_MATRIX, 1)

For i = 1 To NROWS
    DATE_STR = CStr(DATA_MATRIX(i, 1))
    COLLECTION_OBJ.Add COLLECTION_OBJ.COUNT + 1, DATE_STR
Next i

If IsArray(SPLITS_RNG) = True Then
    SPLITS_VECTOR = SPLITS_RNG
    For k = LBound(SPLITS_VECTOR, 1) To UBound(SPLITS_VECTOR, 1)
        DATE_STR = CStr(SPLITS_VECTOR(k, 1))
        If CDate(DATE_STR) <= DATA_MATRIX(1, 1) Then: GoTo 1983
        GoSub INDEX_LINE
        If l <= 0 Then: GoTo 1983
        For i = l To 1 Step -1
            DATA_MATRIX(i, 5) = DATA_MATRIX(i, 5) * 1 / SPLITS_VECTOR(k, 2)
        Next i
1983
    Next k
End If

If IsArray(DIVIDENDS_RNG) = True Then
    DIVIDENDS_VECTOR = DIVIDENDS_RNG
    For k = LBound(DIVIDENDS_VECTOR, 1) To UBound(DIVIDENDS_VECTOR, 1)
        DATE_STR = CStr(DIVIDENDS_VECTOR(k, 1))
        If CDate(DATE_STR) <= DATA_MATRIX(1, 1) Then: GoTo 1984
        GoSub INDEX_LINE
        If l <= 0 Then: GoTo 1984
        For i = l To 1 Step -1
            MULTIPLIER_VAL = (DIVIDENDS_VECTOR(k, 3) - DIVIDENDS_VECTOR(k, 2)) / DIVIDENDS_VECTOR(k, 3)
            DATA_MATRIX(i, 5) = DATA_MATRIX(i, 5) * MULTIPLIER_VAL
        Next i
1984
    Next k
End If
ASSET_ADJUSTING_SPLITS_DIVIDENDS_FUNC = DATA_MATRIX

Exit Function
'-------------------------------------------------------------------------------------
INDEX_LINE:
'-------------------------------------------------------------------------------------
    l = CLng(COLLECTION_OBJ(DATE_STR))
    If Err.number <> 0 Then 'Dividend Date > Last Date in the Time Serie
        Err.number = 0
        l = NROWS
    Else
        l = COLLECTION_OBJ(DATE_STR)
        l = l - 1
    End If
'-------------------------------------------------------------------------------------
Return
'-------------------------------------------------------------------------------------
ERROR_LABEL:
'-------------------------------------------------------------------------------------
ASSET_ADJUSTING_SPLITS_DIVIDENDS_FUNC = Err.number
End Function

'Here we look at IGE, an ETF that has had both splits and dividends in its
'history. It had a 2:1 split on June 9, 2005 (the ex-date). Let’s look at the
'unadjusted prices around that date (you can download the historical prices
'of IGE from Yahoo! Finance into an Excel spreadsheet):
                            
'Date    Open    High    Low Close
'6/7/2005    145 146.07  144.11  144.11
'6/8/2005    144.13  146.44  143.75  144.48
'6/9/2005    72.45   73.74   72.23   73.74
'6/10/2005   73.98   74.08   73.31   74
                            
'We need to adjust the prices prior to 6/9/2005 due to this split. This
'is easy: N = 2 here, and all we need to do is to multiply those prices by 1/2
'The following table shows the adjusted prices:
                            
'Date    Open    High    Low Close
'6/7/2005    72.5    73.035  72.055  72.055
'6/8/2005    72.065  73.22   71.875  72.24
'6/9/2005    72.45   73.74   72.23   73.74
'6/10/2005   73.98   74.08   73.31   74

'Now, the astute reader will notice that the adjusted close prices here
'do not match the adjusted close prices displayed in the Yahoo! Finance
'table. The reason for this is that there have been dividends distributed after
'6/9/2005, so the Yahoo! prices have been adjusted for all those as
'well. Since each adjustment is a multiplier, the aggregate adjustment is
'just the product of all the individual multipliers. Here are the dividends
'from 6/9/2005 to November 2007, together with the unadjusted closing
'prices of the previous trading days and the resulting individual multipliers:
                            
'Ex-Date Dividend    Prev Close  Multiplier
'6/21/2005   0.217   77.9    0.997214
'9/26/2005   0.184   89  0.997933
'12/23/2005  0.236   89.87   0.997374
'3/27/2006   0.253   94.79   0.997331
'6/23/2006   0.32    92.2    0.996529
'9/27/2006   0.258   91.53   0.997181
'12/21/2006  0.322   102.61  0.996862
'6/29/2007   0.3 119.44  0.997488
'9/26/2007   0.177   128.08  0.998618
                            
'(Check out the multipliers yourself on Excel using the formula I gave
'above to see if they agree with my values here.) So the aggregate multiplier
'for the dividends is simply 0.998618 × 0.997488 × · · · × 0.997214 =
'0.976773. This multiplier should be applied to all the unadjusted prices
'on or after 6/9/2005. The aggregate multiplier for the dividends and the
'split is 0.976773 × 0.5 = 0.488386, which should be applied to all the
'unadjusted prices before 6/9/2005. So let’s look at the resulting adjusted
'prices after applying these multipliers:
                            
'Date    Open    High    Low Close
'6/7/2005    70.81601    71.33858    70.38135    70.38135
'6/8/2005    70.39111    71.51929    70.20553    70.56205
'6/9/2005    70.76717    72.02721    70.55228    72.02721
'6/10/2005   72.26163    72.35931    71.6072     72.28117
                            
'Now compare with the table from Yahoo! around November 1, 2007:
                            
'Date    Open    High    Low Close   Volume  Adj Close
'6/7/2005    145 146.07  144.11  144.11  58000   70.38
'6/8/2005    144.13  146.44  143.75  144.48  109600  70.56
'6/9/2005    72.45   73.74   72.23   73.74   853200  72.03
'6/10/2005   73.98   74.08   73.31   74  179300  72.28
                            
'You can see that the adjusted closing prices from our calculations and
'from Yahoo! are the same (after rounding to two decimal places). But, of
'course, when you are reading this, IGE will likely have distributed more
'dividends and may have even split further, so your Yahoo! table won’t look
'like the one above. It is a good exercise to check that you can make further
'adjustments based on those dividends and splits that result in the same
'adjusted prices as your current Yahoo! table.

Function ASSET_STOCK_SPLIT_FUNC(ByRef TICKER_STR As Variant, _
Optional ByVal START_DATE As Date, _
Optional ByVal END_DATE As Date, _
Optional ByVal SPLIT_FACTOR As Double = 0.2)

'SPLIT_FACTOR - from 0.01 to 1

Dim i As Long
Dim j As Long

Dim NROWS As Long

Dim TEMP1_VAL As Double
Dim TEMP2_VAL As Double
Dim TEMP_STR As String
Dim TEMP_MULT As Double

Dim TEMP_VECTOR As Variant
Dim DATA_MATRIX As Variant

On Error GoTo ERROR_LABEL

If IsArray(TICKER_STR) = True Then
    DATA_MATRIX = TICKER_STR
Else
    DATA_MATRIX = _
    YAHOO_HISTORICAL_DATA_SERIE_FUNC(TICKER_STR, START_DATE, END_DATE, _
    "d", "DOHLCVA", False, False, True)
End If
NROWS = UBound(DATA_MATRIX, 1)

j = 0
i = 1
TEMP_MULT = DATA_MATRIX(i, 7) / DATA_MATRIX(i, 5)

i = 2
TEMP1_VAL = 0
TEMP2_VAL = TEMP1_VAL

For i = 3 To NROWS
    TEMP1_VAL = IIf((DATA_MATRIX(i, 7) / DATA_MATRIX(i, 5) < (1 + SPLIT_FACTOR) * TEMP_MULT And DATA_MATRIX(i, 7) / DATA_MATRIX(i, 5) > (1 - SPLIT_FACTOR) * TEMP_MULT), 0, DATA_MATRIX(i, 7))
    If (TEMP1_VAL > 0 And TEMP2_VAL = 0) Then
        j = i
        Exit For
    End If
    TEMP2_VAL = TEMP1_VAL
Next i

If j <> 0 Then
    TEMP1_VAL = 1 / (DATA_MATRIX(j, 5) / DATA_MATRIX(j - 1, 5)) 'shares for one
    TEMP_STR = Format(TEMP1_VAL, "0.00") & "x Stock Split : " & Format(DATA_MATRIX(j, 1), "mmm d/yy")
    
    ReDim TEMP_VECTOR(1 To 5, 1 To 1)
    TEMP_VECTOR(1, 1) = TEMP_STR
    TEMP_VECTOR(2, 1) = DATA_MATRIX(j, 1)
    TEMP_VECTOR(3, 1) = DATA_MATRIX(j, 7)
    TEMP_VECTOR(4, 1) = TEMP1_VAL
    TEMP_VECTOR(5, 1) = j
    
    ASSET_STOCK_SPLIT_FUNC = TEMP_VECTOR
Else
    ASSET_STOCK_SPLIT_FUNC = "No Stock Split"
End If

Exit Function
ERROR_LABEL:
ASSET_STOCK_SPLIT_FUNC = Err.number
End Function


