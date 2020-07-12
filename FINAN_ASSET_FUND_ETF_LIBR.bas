Attribute VB_Name = "FINAN_ASSET_FUND_ETF_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.


'You should get all the most recent prices … in particular, Friday's closing prices.
'These prices will be the BASE PRICES from which the expected price of the ETF is
'calculated thoughout the subsequent week.

'On market days:
'The latest Prices (from Yahoo) will be downloaded and their changes from their BASE PRICES
'will be calculated. These changes (weighted accordingly) will be applied to the BASE PRICE
'for the ETF (normally Friday's closing price). You 'll get what should be the latest ETF
'price … maybe.

'P.S. Of course, you can do these on any day after the markets have closed … and
'get closing prices. Then compare subsequent new Prices with them thar closing prices.

'--------------------------------------------------------------------------------------------
'REFERENCES:
'--------------------------------------------------------------------------------------------
'http://67.220.225.70/~gumm5981/CEFs.htm
'http://67.220.225.70/~gumm5981/ETF-NAV.htm
'http://www.claymoreinvestments.ca/etf/fund/cew
'--------------------------------------------------------------------------------------------

Function ETF_CHEAP_RICH_FUNC(ByVal ETF_STR As String, _
ByRef TICKERS_RNG As Variant, _
ByRef WEIGHTS_RNG As Variant, _
Optional ByVal BASE_DATE As Date = 0, _
Optional ByVal OUTPUT As Integer = 0)

'BASE_DATE --> These prices will be the BASE PRICES from which the
'expected price of the ETF is calculated thoughout the subsequent week.
'This is the old BASE Price for the ETF.


Dim i As Long
Dim NROWS As Long

Dim TEMP_SUM As Double
'This is the total of the allocations of old ETF BASE Price.

Dim WEIGHT_SUM As Double
'The is a check to see if the sum of percentages is close to 100%
Dim CURRENT_DATE As Date
Dim TEMP_ARR As Variant
Dim TEMP_MATRIX As Variant
Dim TICKERS_VECTOR As Variant
Dim WEIGHTS_VECTOR As Variant 'The weights assigned to each component.

On Error GoTo ERROR_LABEL

If BASE_DATE = 0 Then 'This is the date when you established the "Base Prices".
    CURRENT_DATE = Now()
    CURRENT_DATE = _
        DateSerial(Year(CURRENT_DATE), Month(CURRENT_DATE), Day(CURRENT_DATE))
    BASE_DATE = DateSerial(Year(CURRENT_DATE), Month(CURRENT_DATE), Day(CURRENT_DATE) - 1)
End If

TICKERS_VECTOR = TICKERS_RNG
If UBound(TICKERS_VECTOR, 1) = 1 Then
    TICKERS_VECTOR = MATRIX_TRANSPOSE_FUNC(TICKERS_VECTOR)
End If

WEIGHTS_VECTOR = WEIGHTS_RNG
If UBound(WEIGHTS_VECTOR, 1) = 1 Then
    WEIGHTS_VECTOR = MATRIX_TRANSPOSE_FUNC(WEIGHTS_VECTOR)
End If

If UBound(TICKERS_VECTOR, 1) <> UBound(WEIGHTS_VECTOR, 1) Then: GoTo ERROR_LABEL
NROWS = UBound(TICKERS_VECTOR, 1)

WEIGHT_SUM = 0
For i = 1 To NROWS
    WEIGHT_SUM = WEIGHT_SUM + WEIGHTS_VECTOR(i, 1)
Next i

ReDim TEMP_MATRIX(0 To NROWS + 1, 1 To 6)

TEMP_MATRIX(0, 1) = "SYMBOL"
TEMP_MATRIX(0, 2) = "WEIGHT"
'The weights assigned to each component.

TEMP_MATRIX(0, 3) = "CURRENT VALUE"
TEMP_MATRIX(0, 4) = "PREVIOUS VALUE"

TEMP_MATRIX(0, 5) = "CHANGE" 'These are the changes since
'the old "Base Prices" were downloaded.
TEMP_MATRIX(0, 6) = "DOLLARS VALUE"
'These are how the old ETF "Base Price" is allocated,
'according the asset weights. This is the total of the
'allocations of old ETF BASE Price.  It should agree with
'ETF Previous Value.

ReDim TEMP_ARR(1 To 1, 1 To 1)
TEMP_ARR(1, 1) = "Last Trade"

TEMP_MATRIX(1, 1) = ETF_STR
TEMP_MATRIX(1, 3) = YAHOO_QUOTES_FUNC(TEMP_MATRIX(1, 1), TEMP_ARR, 0, False, "")(1, 1)
TEMP_MATRIX(1, 4) = YAHOO_HISTORICAL_DATA_LAST_PRICE_FUNC(TEMP_MATRIX(1, 1), BASE_DATE)
TEMP_MATRIX(1, 5) = TEMP_MATRIX(1, 3) / TEMP_MATRIX(1, 4)

TEMP_SUM = 0
TEMP_MATRIX(1, 2) = 0
TEMP_MATRIX(1, 6) = 0

TEMP_ARR = YAHOO_QUOTES_FUNC(TICKERS_VECTOR, TEMP_ARR, 0, False, "")
'These are the latest prices downloaded from Yahoo.

For i = 1 To NROWS
    TEMP_MATRIX(i + 1, 1) = TICKERS_VECTOR(i, 1)
    TEMP_MATRIX(i + 1, 2) = WEIGHTS_VECTOR(i, 1)
    
    TEMP_MATRIX(i + 1, 3) = TEMP_ARR(i, 1)

    TEMP_MATRIX(i + 1, 4) = YAHOO_HISTORICAL_DATA_LAST_PRICE_FUNC(TEMP_MATRIX(i + 1, 1), BASE_DATE)
        
    If TEMP_MATRIX(i + 1, 3) = 0 Or TEMP_MATRIX(i + 1, 4) = 0 Then
        TEMP_MATRIX(i + 1, 5) = 1
    Else
        TEMP_MATRIX(i + 1, 5) = TEMP_MATRIX(i + 1, 3) / TEMP_MATRIX(i + 1, 4)
    End If
    TEMP_MATRIX(i + 1, 6) = TEMP_MATRIX(i + 1, 2) * TEMP_MATRIX(1, 4) / WEIGHT_SUM
    
    TEMP_MATRIX(1, 2) = TEMP_MATRIX(1, 2) + TEMP_MATRIX(i + 1, 2)
    TEMP_MATRIX(1, 6) = TEMP_MATRIX(1, 6) + TEMP_MATRIX(i + 1, 6)
    TEMP_SUM = TEMP_SUM + TEMP_MATRIX(i + 1, 6) * TEMP_MATRIX(i + 1, 5)
Next i

'---------------------------------------------------------------------------------
Select Case OUTPUT
'---------------------------------------------------------------------------------
Case 0
'---------------------------------------------------------------------------------
'This is what the ETF should be worth, based upon the "weighted" changes
'in the component prices.
    ETF_CHEAP_RICH_FUNC = TEMP_MATRIX(1, 4) * TEMP_SUM / TEMP_MATRIX(1, 6)
'---------------------------------------------------------------------------------
Case Else
'---------------------------------------------------------------------------------
    ETF_CHEAP_RICH_FUNC = TEMP_MATRIX
'---------------------------------------------------------------------------------
End Select
'---------------------------------------------------------------------------------

Exit Function
ERROR_LABEL:
ETF_CHEAP_RICH_FUNC = Err.number
End Function

'Mutual Funds invest in a basket of stocks and, at the end of each day, the Net
'Asset Value (NAV) is calculated, depending upon what percentage of the fund is
'invested in what stock.

'Alas, you have to wait until the market has closed to discover the Fund NAV.
'Exchange Traded Funds (ETF) are much like Mutual Funds except that they trade
'throughout the day, just like stocks.

'In fact, their price depends upon the mechanics of bid-and-ask rather than the
'prices of the underlying stocks. Usually, the NAV of the ETF will be available
'on some website after the market has closed and, usually, it's pretty close to
'the last trade, like Closing Market Price --> 39.86 / Closing NAV --> 39.85

'>So what's your point?
'Wouldn 't it be fun to know the NAV throughout the day. Then you could tell whether
'the ETF is trading low or high and buy or sell accordingly.

'>You mean the NAV isn't available thoughout the day?
'Sometimes it is ... sometimes it ain't. See, for example, etfconnect.com

'>But what if the ETF has dozens of components? How would you ...?
'How would you calculate the NAV? As it happens ...

'Here 's what we do:

'    * After the market has closed we download all the component stock prices as well
'      as the ETF closing price. This establishes our BASE Prices.

'    * If the ETF BASE Price is, say, $50, we divide this up according to the weights
'      attached to each component.

'      For example:

'      If the most heavily weighted stock has a weight of 9.25% then we allocate 0.0925*($50)
'      = $4.63 to that first stock.

'      If the next stock has a weight of 8.97% then we allocate 0.0897*($50) = $4.49 to that stock.

'      If the next stock has a weight of 7.32% then we allocate 0.0732*($50) = $3.36 to that stock.
'      If the next stock ...

'>Okay! I get it!
'We now have the ETF BASE Price subdivided according to the weights.

'    * On subsequent days we download the current current stock prices and see how they've
'      changed from their BASE Prices and apply these changes to the ETF Price.
'      For example:

'      If the first stock has changed by 0.7% we change the $4.63 allocation to ($4.63)*(1.007)
'      = $4.66.

'      If the next stock has changed by -0.9% we change the $4.49 allocation to ($4.49)*(0.991)
'      = $4.45.
'      If the next stock has changed by ...

'Adding all the modified allocations, we might have, for example, $51.25 which is our calculated
'NAV at the time we downloaded the current prices.
'Now we compare to the current ETF price (which we downloaded along with all the component prices).
'Then we decide whether the ETF is trading high or low and ...


'>Aren't you assumimg that the downloaded BASE Price is the NAV? Why wouldn't you ...?
'Uh ... yes, you're right.
'You may want to change it to the old NAV which (somehow) you get from a website.


'ETF_CHEAP_RICH_FUNC --> You'll get what should be the latest ETF price … maybe.

