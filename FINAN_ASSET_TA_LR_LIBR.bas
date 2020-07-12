Attribute VB_Name = "FINAN_ASSET_TA_LR_LIBR"


Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.


'************************************************************************************
'************************************************************************************
'FUNCTION      : ASSET_TA_LIQUIDITY_RATIO_FUNC
'DESCRIPTION   : Liquidity Ratio

'There are liquidity ratios and liquidity ratios (see this google list) such as this which considers
'the ratio of a company's assets (cash and securities) and its debt or liabilities. However, that's
'not the one we want to consider. Instead, we'll ...

'A liquid investment is one where you can get in and out easily and quickly ... like your bank
'account. But some stocks sell so few shares each day that if you wanted to buy or sell $10K you
'may have a problem if there aren't many shares trading each day.

'Of course, that wouldn't be the case for GE stock (for example). Maybe $600K worth of stock changes
'hands each day, so buying or selling $10K would be a piece of cake. That's a liquid stock.

'I understand that the Amivest corporation introduced a liquidity ratio that attempts to determine the
'dollar volume of shares which would result in a 1% change in stock price.

'It goes like so (as I understand it):
'Calculate the volume of shares that traded each day (over a month). Call if V(j) ... for day j.
'Pick some representative price for each day. Call if p(j) which may be the closing price for day j.
'Calculate the Dollar Volume for each day. That's V(j)p(j) for day j.
'Calculate the Total Dollar Volume for the month. That's DV = V(1)p(1) + V(2)p(2) + ...
'Next, calculate the percentage changes in daily stock prices (whether it's up or down).
'If the daily percentage changes are r1, r2, ... r22   for 22 market days in the month
'then put:   R = { |r1| + |r2| + ... + |r22| }   the total of the magnitudes of the daily changes
'with r = 1.23 for a 1.23% change
'Since DV generated a TOTAL monthly change of R, we calculate the monthly dollar volume for each 1%
'change in stock price, namely the Ratio:   Liquidity Ratio = DV / R

'We may interpret this ratio as the monthly "dollar volume" necessary to effect a 1% change in stock
'price over one month.
'The "dollar volume" means the total number of dollars that were traded.
'Amivest Liquidity Ratio = (monthly dollar volume) / (sum of absolute value of daily percentage
'changes in stock price)
'or
'Amivest Liquidity Ratio = ( V(1)p(1) + V(2)p(2) + ... + V(n)p(n) ) / ( |r1| + |r2| + ... + |rn| )
'for n market days in the month

'where Vk, pk and rk are the daily volume of shares traded, the daily prices and the percentage changes
'in daily stock prices and, for a 1.23% daily return, we use rk = 1.23,   not 0.0123 !!

'We interpret this ratio as the monthly dollar volume necessary to effect a 1% monthly change in
'stock price. If the ratio is large, it means there needs to be an awful lot of money traded to
'effect a 1% change in stock price.

'It 's August 31, 2004 and I check out smartmoney.com which defines Liquidity Ratio like so:
'This ratio is a measure of how much dollar volume is required to move a stock's price up or down by
'one percentage point. A high ratio indicates a stock that requires relatively heavy trading to move
'its price. A low liquidity ratio indicates a stock that moves on relatively light volume. The ratio
'is calculated by adding the daily percentage changes of a stock's closing price for each trading
'day of the month. Then the total dollar volume for the month is divided by this total-percentage-change
'figure.

'Then I check out GE and find that the Liquidity Ratio is about 774,453.
'We ask: "Does that number make sense?"
'We note the following:
'The average volume of shares traded each day is about 20,000,000.
'The price per share is something like $32.
'That makes the daily dollar volume of trades roughly 32*20,000,000 = $640,000,000.
'Over one month of 22 trading days, the dollar volume of shares traded would be something
'like 22*640,000,000 = $14,080,000,000.
'Is that the numerator of our Liquidity Ratio?
'In order for the Liquidity Ratio to equal 774,453, the denominator must be something like
'14,080,000,000 / 774,453 = 18,181.
'We recognize this as the total of daily percentage changes over one month.

'You mean a 18,181% change ... in one month?
'Well, the actual sum of daily percentage changes was more like 18%
'... so I assume that we take the daily dollar volume divided by 1000

'So, we (hopefully) reproduce the SmartMoney Liquidity Ratio like so:
'1000 * Liquidity Ratio = (monthly dollar volume) / (sum of absolute value of daily percentage
'changes in stock price)
'or
'1000 * Liquidity Ratio = ( V(1)p(1) + V(2)p(2) + ... + V(n)p(n) ) / ( |r1| + |r2| + ... + |rn| )
'for n market days in the month

'where Vk, pk and rk are the daily volume of shares traded, the daily prices and the percentage
'changes in daily stock prices. Now we interpret this Liquidity Ratio as the number of monthly
'kilo-bucks which will effect that 1% change in monthly stock price.

'http://www.investorwords.com/2840/liquidity_ratio.html
'http://www.gummy-stuff.org/liquidity-ratio.htm

'LIBRARY       : FINAN_ASSET
'GROUP         : LIQUIDITY
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'************************************************************************************
'************************************************************************************

Function ASSET_TA_LIQUIDITY_RATIO_FUNC(ByVal TICKER_STR As Variant, _
Optional ByVal START_DATE As Date, _
Optional ByVal END_DATE As Date, _
Optional ByVal LR_PERIOD As Long = 10)

Dim i As Long
Dim j As Long
Dim k As Long

Dim NROWS As Long
Dim MIN_VAL As Double

Dim ATEMP_SUM As Double
Dim BTEMP_SUM As Double
Dim CTEMP_SUM As Double
Dim DTEMP_SUM As Double

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

ReDim TEMP_MATRIX(0 To NROWS, 1 To 12)
TEMP_MATRIX(0, 1) = "DATE"
TEMP_MATRIX(0, 2) = "OPEN"
TEMP_MATRIX(0, 3) = "HIGH"
TEMP_MATRIX(0, 4) = "LOW"
TEMP_MATRIX(0, 5) = "CLOSE"
TEMP_MATRIX(0, 6) = "VOLUME" '6
TEMP_MATRIX(0, 7) = "ADJ.PRICE" '7
TEMP_MATRIX(0, 8) = "RETURNS" '8
TEMP_MATRIX(0, 9) = "$VOLUME" '9
TEMP_MATRIX(0, 10) = "ABS(RETURNS)" '10
TEMP_MATRIX(0, 11) = "DAYS BACK" '11

ATEMP_SUM = 0: BTEMP_SUM = 0
CTEMP_SUM = 0: DTEMP_SUM = 0
MIN_VAL = LR_PERIOD

i = 1
For j = 1 To 7: TEMP_MATRIX(i, j) = DATA_MATRIX(i, j): Next j
TEMP_MATRIX(i, 6) = TEMP_MATRIX(i, 6) / 1000
TEMP_MATRIX(i, 8) = TEMP_MATRIX(i, 5) / TEMP_MATRIX(i, 2) - 1

TEMP_MATRIX(i, 9) = TEMP_MATRIX(i, 6) * TEMP_MATRIX(i, 5)
ATEMP_SUM = ATEMP_SUM + TEMP_MATRIX(i, 9)

TEMP_MATRIX(i, 10) = 100 * Abs(TEMP_MATRIX(i, 8))
BTEMP_SUM = BTEMP_SUM + TEMP_MATRIX(i, 10)

TEMP_MATRIX(i, 11) = IIf(i < MIN_VAL, i, LR_PERIOD) '
TEMP_MATRIX(i, 12) = TEMP_MATRIX(i, 9) / TEMP_MATRIX(i, 10)

For i = 2 To NROWS
    For j = 1 To 7: TEMP_MATRIX(i, j) = DATA_MATRIX(i, j): Next j
    TEMP_MATRIX(i, 6) = TEMP_MATRIX(i, 6) / 1000
    TEMP_MATRIX(i, 8) = TEMP_MATRIX(i, 7) / TEMP_MATRIX(i - 1, 7) - 1
    
    TEMP_MATRIX(i, 9) = TEMP_MATRIX(i, 6) * TEMP_MATRIX(i, 5)
    ATEMP_SUM = ATEMP_SUM + TEMP_MATRIX(i, 9)
    
    TEMP_MATRIX(i, 10) = 100 * Abs(TEMP_MATRIX(i, 8))
    BTEMP_SUM = BTEMP_SUM + TEMP_MATRIX(i, 10)
    
    TEMP_MATRIX(i, 11) = IIf(i < MIN_VAL, i, LR_PERIOD) '
    
    CTEMP_SUM = 0
    DTEMP_SUM = 0
    
    If i <= LR_PERIOD Then
        k = i - TEMP_MATRIX(i, 11) + 1
    Else
        k = i - TEMP_MATRIX(i, 11)
    End If
    
    For j = i To k Step -1
        CTEMP_SUM = CTEMP_SUM + TEMP_MATRIX(j, 9)
        DTEMP_SUM = DTEMP_SUM + TEMP_MATRIX(j, 10)
    Next j

    If DTEMP_SUM <> 0 Then: TEMP_MATRIX(i, 12) = CTEMP_SUM / DTEMP_SUM
Next i

If BTEMP_SUM <> 0 Then
    TEMP_MATRIX(0, 12) = "LIQUIDITY RATIO : " & Format(LR_PERIOD, "0") & _
                          "- PERIOD AVERAGE = " & Format(ATEMP_SUM / BTEMP_SUM, "#,000.0")
Else
    TEMP_MATRIX(0, 12) = "LIQUIDITY RATIO : " & Format(LR_PERIOD, "0") & _
                          "- PERIOD AVERAGE"
End If

ASSET_TA_LIQUIDITY_RATIO_FUNC = TEMP_MATRIX

Exit Function
ERROR_LABEL:
ASSET_TA_LIQUIDITY_RATIO_FUNC = Err.number
End Function

