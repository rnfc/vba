Attribute VB_Name = "FINAN_PORT_TRADE_PERFORM_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.


'************************************************************************************
'************************************************************************************
'FUNCTION      : PORT_TRACK_NET_PERFORMANCE_FUNC
'DESCRIPTION   : Long Short Portfolio - Analysing Active Returns
'LIBRARY       : PORTFOLIO
'GROUP         : TRADE_EXPOSURE
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 21/01/2009
'************************************************************************************
'************************************************************************************

Function PORT_TRACK_NET_PERFORMANCE_FUNC(ByVal NET_EXPOSURE As Double, _
ByVal CASH_RATE As Double, _
ByRef DATE_RNG As Variant, _
ByRef FUND_PERFORMANCE_RNG As Variant, _
ByRef AVG_NET_EXPOSURE_RNG As Variant, _
ByRef MARKET_PERFORMANCE_RNG As Variant, _
Optional ByVal FREQUENCY As Double = 12, _
Optional ByVal OUTPUT As Integer = 0)

Dim i As Long
Dim NCOLUMNS As Long

Dim DATE_VECTOR As Variant
Dim FUND_VECTOR As Variant
Dim EXPOSURE_VECTOR As Variant
Dim MARKET_VECTOR As Variant

Dim TEMP_VECTOR As Variant
Dim TEMP_MATRIX As Variant
Dim SUMMARY_MATRIX As Variant

On Error GoTo ERROR_LABEL

DATE_VECTOR = DATE_RNG
If UBound(DATE_VECTOR, 2) = 1 Then: DATE_VECTOR = MATRIX_TRANSPOSE_FUNC(DATE_VECTOR)
    
FUND_VECTOR = FUND_PERFORMANCE_RNG
If UBound(FUND_VECTOR, 2) = 1 Then: FUND_VECTOR = MATRIX_TRANSPOSE_FUNC(FUND_VECTOR)
If UBound(DATE_VECTOR, 2) <> UBound(FUND_VECTOR, 2) Then: GoTo ERROR_LABEL

EXPOSURE_VECTOR = AVG_NET_EXPOSURE_RNG
If UBound(EXPOSURE_VECTOR, 2) = 1 Then: EXPOSURE_VECTOR = MATRIX_TRANSPOSE_FUNC(EXPOSURE_VECTOR)
If UBound(DATE_VECTOR, 2) <> UBound(EXPOSURE_VECTOR, 2) Then: GoTo ERROR_LABEL

MARKET_VECTOR = MARKET_PERFORMANCE_RNG
If UBound(MARKET_VECTOR, 2) = 1 Then: MARKET_VECTOR = MATRIX_TRANSPOSE_FUNC(MARKET_VECTOR)
If UBound(DATE_VECTOR, 2) <> UBound(MARKET_VECTOR, 2) Then: GoTo ERROR_LABEL


NCOLUMNS = UBound(DATE_VECTOR, 2)
ReDim TEMP_MATRIX(1 To 12, 1 To NCOLUMNS + 1)
ReDim TEMP_VECTOR(1 To NCOLUMNS, 1 To 4)

'First Pass Calculations: Manager Performance

TEMP_MATRIX(1, 1) = "PERIOD"
TEMP_MATRIX(2, 1) = "FUND PERFORMANCE"
TEMP_MATRIX(3, 1) = "CUMMULATIVE FUND PERFORMANCE"
TEMP_MATRIX(4, 1) = "AVERAGE NET EXPOSURE"
TEMP_MATRIX(5, 1) = "MARKET PERFORMANCE"
TEMP_MATRIX(6, 1) = "CUMMULATIVE MARKET PERFORMANCE"
TEMP_MATRIX(7, 1) = "ALPHA (STOCK SELECTION PERFORMANCE)"
TEMP_MATRIX(8, 1) = "CUMMULATIVE ALPHA PERFORMANCE"
TEMP_MATRIX(9, 1) = "MARKET TIMING PERFORMANCE"
TEMP_MATRIX(10, 1) = "CUMMULATIVE MARKET TIMING PERFORMANCE"
TEMP_MATRIX(11, 1) = "TOTAL ACTIVE VALUE"
TEMP_MATRIX(12, 1) = "CUMMULATIVE ACTIVE VALUE"

TEMP_MATRIX(1, 1 + 1) = DATE_VECTOR(1, 1)
TEMP_MATRIX(2, 1 + 1) = FUND_VECTOR(1, 1)
TEMP_MATRIX(3, 1 + 1) = 100 * (1 + TEMP_MATRIX(2, 1 + 1))
TEMP_MATRIX(4, 1 + 1) = EXPOSURE_VECTOR(1, 1)
TEMP_MATRIX(5, 1 + 1) = MARKET_VECTOR(1, 1)
TEMP_MATRIX(6, 1 + 1) = 100 * (1 + TEMP_MATRIX(5, 1 + 1))
TEMP_MATRIX(7, 1 + 1) = TEMP_MATRIX(2, 1 + 1) - (TEMP_MATRIX(4, 1 + 1) * TEMP_MATRIX(5, 1 + 1)) - ((1 - TEMP_MATRIX(4, 1 + 1)) * (CASH_RATE / FREQUENCY))
TEMP_MATRIX(8, 1 + 1) = 100 * (1 + TEMP_MATRIX(7, 1 + 1))
TEMP_MATRIX(9, 1 + 1) = (TEMP_MATRIX(4, 1 + 1) - NET_EXPOSURE) * (TEMP_MATRIX(5, 1 + 1) - CASH_RATE / FREQUENCY)
TEMP_MATRIX(10, 1 + 1) = 100 * (1 + TEMP_MATRIX(9, 1 + 1))
TEMP_MATRIX(11, 1 + 1) = TEMP_MATRIX(2, 1 + 1) - (NET_EXPOSURE * TEMP_MATRIX(5, 1 + 1)) - ((1 - NET_EXPOSURE) * (CASH_RATE / FREQUENCY))
TEMP_MATRIX(12, 1 + 1) = 100 * (1 + TEMP_MATRIX(11, 1 + 1))

TEMP_VECTOR(1, 1) = TEMP_MATRIX(2, 1 + 1)
TEMP_VECTOR(1, 2) = TEMP_MATRIX(7, 1 + 1)
TEMP_VECTOR(1, 3) = TEMP_MATRIX(9, 1 + 1)
TEMP_VECTOR(1, 4) = TEMP_MATRIX(11, 1 + 1)

For i = 2 To NCOLUMNS
    TEMP_MATRIX(1, 1 + i) = DATE_VECTOR(1, i)
    TEMP_MATRIX(2, 1 + i) = FUND_VECTOR(1, i)
    TEMP_MATRIX(3, 1 + i) = TEMP_MATRIX(3, i) * (1 + TEMP_MATRIX(2, 1 + i))
    TEMP_MATRIX(4, 1 + i) = EXPOSURE_VECTOR(1, i)
    TEMP_MATRIX(5, 1 + i) = MARKET_VECTOR(1, i)
    TEMP_MATRIX(6, 1 + i) = TEMP_MATRIX(6, i) * (1 + TEMP_MATRIX(5, 1 + i))
    TEMP_MATRIX(7, 1 + i) = TEMP_MATRIX(2, 1 + i) - (TEMP_MATRIX(4, 1 + i) * TEMP_MATRIX(5, 1 + i)) - ((1 - TEMP_MATRIX(4, 1 + i)) * (CASH_RATE / FREQUENCY))
    TEMP_MATRIX(8, 1 + i) = TEMP_MATRIX(8, i) * (1 + TEMP_MATRIX(7, 1 + i))
    TEMP_MATRIX(9, 1 + i) = (TEMP_MATRIX(4, 1 + i) - NET_EXPOSURE) * (TEMP_MATRIX(5, 1 + i) - CASH_RATE / FREQUENCY)
    TEMP_MATRIX(10, 1 + i) = TEMP_MATRIX(10, i) * (1 + TEMP_MATRIX(9, 1 + i))
    TEMP_MATRIX(11, 1 + i) = TEMP_MATRIX(2, 1 + i) - (NET_EXPOSURE * TEMP_MATRIX(5, 1 + i)) - ((1 - NET_EXPOSURE) * (CASH_RATE / FREQUENCY))
    TEMP_MATRIX(12, 1 + i) = TEMP_MATRIX(12, i) * (1 + TEMP_MATRIX(11, 1 + i))

    TEMP_VECTOR(i, 1) = TEMP_MATRIX(2, 1 + i)
    TEMP_VECTOR(i, 2) = TEMP_MATRIX(7, 1 + i)
    TEMP_VECTOR(i, 3) = TEMP_MATRIX(9, 1 + i)
    TEMP_VECTOR(i, 4) = TEMP_MATRIX(11, 1 + i)
Next i

If OUTPUT = 0 Then
    PORT_TRACK_NET_PERFORMANCE_FUNC = TEMP_MATRIX
    Exit Function
End If

TEMP_VECTOR = MATRIX_STDEV_FUNC(TEMP_VECTOR)

ReDim SUMMARY_MATRIX(1 To 4, 1 To 5)

SUMMARY_MATRIX(1, 1) = "KEY STATISTICS"
SUMMARY_MATRIX(1, 2) = "FUND"
SUMMARY_MATRIX(1, 3) = "ALPHA (STOCK SELECTION)"
SUMMARY_MATRIX(1, 4) = "MARKET TIMING"
SUMMARY_MATRIX(1, 5) = "TOTAL ACTIVE RETURNS"

SUMMARY_MATRIX(2, 1) = "ANNUALISED PERFORMANCE"
SUMMARY_MATRIX(2, 2) = (TEMP_MATRIX(3, 1 + NCOLUMNS) - 100) / 100
SUMMARY_MATRIX(2, 3) = (TEMP_MATRIX(8, 1 + NCOLUMNS) - 100) / 100
SUMMARY_MATRIX(2, 4) = (TEMP_MATRIX(10, 1 + NCOLUMNS) - 100) / 100
SUMMARY_MATRIX(2, 5) = (TEMP_MATRIX(12, 1 + NCOLUMNS) - 100) / 100

SUMMARY_MATRIX(3, 1) = "VOLATILITY"
SUMMARY_MATRIX(3, 2) = TEMP_VECTOR(1, 1) * Sqr(FREQUENCY)
SUMMARY_MATRIX(3, 3) = TEMP_VECTOR(1, 2) * Sqr(FREQUENCY)
SUMMARY_MATRIX(3, 4) = TEMP_VECTOR(1, 3) * Sqr(FREQUENCY)
SUMMARY_MATRIX(3, 5) = TEMP_VECTOR(1, 4) * Sqr(FREQUENCY)

SUMMARY_MATRIX(4, 1) = "RETURN/RISK (INFORMATION)"
SUMMARY_MATRIX(4, 2) = SUMMARY_MATRIX(2, 2) / SUMMARY_MATRIX(3, 2)
SUMMARY_MATRIX(4, 3) = SUMMARY_MATRIX(2, 3) / SUMMARY_MATRIX(3, 3)
SUMMARY_MATRIX(4, 4) = SUMMARY_MATRIX(2, 4) / SUMMARY_MATRIX(3, 4)
SUMMARY_MATRIX(4, 5) = SUMMARY_MATRIX(2, 5) / SUMMARY_MATRIX(3, 5)

'Information Ratio is not the most appropriate title when
'applied to the fund's returns because not all the return
'is active. A better approach would be to remove the 50%
'net exposure bias component.

Select Case OUTPUT
Case 1
    PORT_TRACK_NET_PERFORMANCE_FUNC = SUMMARY_MATRIX
Case Else
    PORT_TRACK_NET_PERFORMANCE_FUNC = Array(SUMMARY_MATRIX, TEMP_MATRIX)
End Select

Exit Function
ERROR_LABEL:
PORT_TRACK_NET_PERFORMANCE_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : PORT_TRACK_RELATIVE_PERFORMANCE_FUNC

'DESCRIPTION   : For an investor, keeping track of personal performance is important.
'Many people believe they know their return rate by calculations in their
'head, but surprisingly, we always overestimate ourselves. Our actual return
'usually comes out a few % lower than we expected. These few percentages can
'add up to the cost of ten of thousands.
'Reference:
'http://www.oldschoolvalue.com/investment-tools/personal-portfolio-vs-sp-500-spreadsheet/

'LIBRARY       : PORTFOLIO
'GROUP         : TRADE_PERFORMANCE
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 21/01/2009
'************************************************************************************
'************************************************************************************

Function PORT_TRACK_RELATIVE_PERFORMANCE_FUNC(ByVal INDEX_STR As String, _
ByRef BUY_TICKERS_RNG As Variant, _
ByRef BUY_QUANTITY_RNG As Variant, _
ByRef BUY_PRICE_RNG As Variant, _
ByRef BUY_COMMISSION_RNG As Variant, _
ByRef BUY_DATE_RNG As Variant, _
ByRef SALE_TICKERS_RNG As Variant, _
ByRef SALE_QUANTITY_RNG As Variant, _
ByRef SALE_BUY_PRICE_RNG As Variant, _
ByRef SALE_BUY_COMMISSION_RNG As Variant, _
ByRef SALE_BUY_DATE_RNG As Variant, _
ByRef SALE_PRICE_RNG As Variant, _
ByRef SALE_DATE_RNG As Variant, _
Optional ByVal CURRENT_DATE As Date = 0, _
Optional ByVal OUTPUT As Integer = 1)

'--------------------------------------------------------------------------------
'INDEX_STR --> Could be an index fund that matches the S&P500
'--------------------------------------------------------------------------------
'BUY --> Stocks still held in portfolio
'SALE --> Stocks in portfolio that have been sold
'--------------------------------------------------------------------------------
'ASSUMPTIONS:
'--------------------------------------------------------------------------------
'Ignores impact of taxes on dividend payouts
'--------------------------------------------------------------------------------
'Ignores commission (assume transaction costs netted out between
'buying issues and buying SPY)
'--------------------------------------------------------------------------------
'http://icarra.com/ is a great portfolio tracker. Lots of statistics and graphs.
'Free and pay options. Sometimes they have problems with their data feeds
'and the quotes are always the prior day close, but it seems to be getting
'better all the time.
'--------------------------------------------------------------------------------

Dim i As Long
Dim j As Long

Dim NSIZE As Long

Dim TEMP_MATRIX As Variant
Dim TEMP_VECTOR As Variant
Dim DATA_MATRIX As Variant

Dim BUY_TICKERS_VECTOR As Variant
Dim BUY_QUANTITY_VECTOR As Variant
Dim BUY_PRICE_VECTOR As Variant
Dim BUY_COMMISSION_VECTOR As Variant
Dim BUY_DATE_VECTOR As Variant
Dim SALE_TICKERS_VECTOR As Variant
Dim SALE_QUANTITY_VECTOR As Variant
Dim SALE_BUY_PRICE_VECTOR As Variant
Dim SALE_BUY_COMMISSION_VECTOR As Variant
Dim SALE_BUY_DATE_VECTOR As Variant
Dim SALE_PRICE_VECTOR As Variant
Dim SALE_DATE_VECTOR As Variant

Dim INDEX_PRICE As Double 'Value of Index on evaluation date

On Error GoTo ERROR_LABEL

If CURRENT_DATE = 0 Then
    CURRENT_DATE = Now()
    CURRENT_DATE = DateSerial(Year(CURRENT_DATE), Month(CURRENT_DATE), Day(CURRENT_DATE))
End If

BUY_TICKERS_VECTOR = BUY_TICKERS_RNG
If UBound(BUY_TICKERS_VECTOR, 1) = 1 Then
    BUY_TICKERS_VECTOR = MATRIX_TRANSPOSE_FUNC(BUY_TICKERS_VECTOR)
End If

BUY_QUANTITY_VECTOR = BUY_QUANTITY_RNG
If UBound(BUY_QUANTITY_VECTOR, 1) = 1 Then
    BUY_QUANTITY_VECTOR = MATRIX_TRANSPOSE_FUNC(BUY_QUANTITY_VECTOR)
End If

BUY_PRICE_VECTOR = BUY_PRICE_RNG
If UBound(BUY_PRICE_VECTOR, 1) = 1 Then
    BUY_PRICE_VECTOR = MATRIX_TRANSPOSE_FUNC(BUY_PRICE_VECTOR)
End If

BUY_COMMISSION_VECTOR = BUY_COMMISSION_RNG
If UBound(BUY_COMMISSION_VECTOR, 1) = 1 Then
    BUY_COMMISSION_VECTOR = MATRIX_TRANSPOSE_FUNC(BUY_COMMISSION_VECTOR)
End If

BUY_DATE_VECTOR = BUY_DATE_RNG
If UBound(BUY_DATE_VECTOR, 1) = 1 Then
    BUY_DATE_VECTOR = MATRIX_TRANSPOSE_FUNC(BUY_DATE_VECTOR)
End If
 
SALE_TICKERS_VECTOR = SALE_TICKERS_RNG
If UBound(SALE_TICKERS_VECTOR, 1) = 1 Then
    SALE_TICKERS_VECTOR = MATRIX_TRANSPOSE_FUNC(SALE_TICKERS_VECTOR)
End If

SALE_QUANTITY_VECTOR = SALE_QUANTITY_RNG
If UBound(SALE_QUANTITY_VECTOR, 1) = 1 Then
    SALE_QUANTITY_VECTOR = MATRIX_TRANSPOSE_FUNC(SALE_QUANTITY_VECTOR)
End If

SALE_BUY_PRICE_VECTOR = SALE_BUY_PRICE_RNG
If UBound(SALE_BUY_PRICE_VECTOR, 1) = 1 Then
    SALE_BUY_PRICE_VECTOR = MATRIX_TRANSPOSE_FUNC(SALE_BUY_PRICE_VECTOR)
End If

SALE_BUY_COMMISSION_VECTOR = SALE_BUY_COMMISSION_RNG
If UBound(SALE_BUY_COMMISSION_VECTOR, 1) = 1 Then
    SALE_BUY_COMMISSION_VECTOR = MATRIX_TRANSPOSE_FUNC(SALE_BUY_COMMISSION_VECTOR)
End If

SALE_BUY_DATE_VECTOR = SALE_BUY_DATE_RNG
If UBound(SALE_BUY_DATE_VECTOR, 1) = 1 Then
    SALE_BUY_DATE_VECTOR = MATRIX_TRANSPOSE_FUNC(SALE_BUY_DATE_VECTOR)
End If

SALE_PRICE_VECTOR = SALE_PRICE_RNG
If UBound(SALE_PRICE_VECTOR, 1) = 1 Then
    SALE_PRICE_VECTOR = MATRIX_TRANSPOSE_FUNC(SALE_PRICE_VECTOR)
End If

SALE_DATE_VECTOR = SALE_DATE_RNG
If UBound(SALE_DATE_VECTOR, 1) = 1 Then
    SALE_DATE_VECTOR = MATRIX_TRANSPOSE_FUNC(SALE_DATE_VECTOR)
End If

NSIZE = UBound(BUY_TICKERS_VECTOR, 1)
If UBound(SALE_TICKERS_VECTOR, 1) > NSIZE Then: NSIZE = UBound(SALE_TICKERS_VECTOR, 1)

'--------------------------------------------------------------------------------

ReDim TEMP_MATRIX(0 To NSIZE + 1, 1 To 28)

'--------------------------------------------------------------------------------
TEMP_MATRIX(0, 1) = "Stock symbol"
TEMP_MATRIX(0, 2) = "# of shares"
TEMP_MATRIX(0, 3) = "Purchase price"
TEMP_MATRIX(0, 4) = "Commission (buy)"
TEMP_MATRIX(0, 5) = "Date of purchase"
TEMP_MATRIX(0, 6) = "Cost basis"
TEMP_MATRIX(0, 7) = "Current Share Price"
TEMP_MATRIX(0, 8) = "Market value"
TEMP_MATRIX(0, 9) = "Gain/Loss"
TEMP_MATRIX(0, 10) = "Percentage Gain/Loss"
TEMP_MATRIX(0, 11) = INDEX_STR & " closing price on date of purchase"
TEMP_MATRIX(0, 12) = "Equivalent # of " & INDEX_STR & " shares"
TEMP_MATRIX(0, 13) = "Market value"
'--------------------------------------------------------------------------------
TEMP_MATRIX(0, 14) = "Stock symbol"
TEMP_MATRIX(0, 15) = "# of shares"
TEMP_MATRIX(0, 16) = "Purchase price"
TEMP_MATRIX(0, 17) = "Commission (buy+sell)"
TEMP_MATRIX(0, 18) = "Date of purchase"
TEMP_MATRIX(0, 19) = "Cost basis"
TEMP_MATRIX(0, 20) = "Sale price"
TEMP_MATRIX(0, 21) = "Date of sale"
TEMP_MATRIX(0, 22) = "Market value"
TEMP_MATRIX(0, 23) = "Gain/Loss"
TEMP_MATRIX(0, 24) = "Percentage Gain/Loss"
TEMP_MATRIX(0, 25) = INDEX_STR & " closing price on date of purchase"
TEMP_MATRIX(0, 26) = INDEX_STR & " closing price on date of sale"
TEMP_MATRIX(0, 27) = "Equivalent # of shares"
TEMP_MATRIX(0, 28) = "Market value"
'--------------------------------------------------------------------------------
TEMP_MATRIX(NSIZE + 1, 1) = "TOTAL"
'--------------------------------------------------------------------------------

ReDim DATA_MATRIX(1 To 1, 1 To 1)
DATA_MATRIX(1, 1) = "Last Trade"

INDEX_PRICE = YAHOO_QUOTES_FUNC(INDEX_STR, DATA_MATRIX, 0, False, "")(1, 1)
DATA_MATRIX = YAHOO_QUOTES_FUNC(BUY_TICKERS_VECTOR, DATA_MATRIX, 0, False, "")

TEMP_MATRIX(NSIZE + 1, 6) = 0
TEMP_MATRIX(NSIZE + 1, 8) = 0
TEMP_MATRIX(NSIZE + 1, 13) = 0
For i = 1 To UBound(BUY_TICKERS_VECTOR, 1)
    TEMP_MATRIX(i, 1) = BUY_TICKERS_VECTOR(i, 1)
    TEMP_MATRIX(i, 2) = BUY_QUANTITY_VECTOR(i, 1)
    TEMP_MATRIX(i, 3) = BUY_PRICE_VECTOR(i, 1)
    TEMP_MATRIX(i, 4) = BUY_COMMISSION_VECTOR(i, 1)
    TEMP_MATRIX(i, 5) = BUY_DATE_VECTOR(i, 1)
    TEMP_MATRIX(i, 6) = TEMP_MATRIX(i, 2) * TEMP_MATRIX(i, 3) + TEMP_MATRIX(i, 4)
    TEMP_MATRIX(NSIZE + 1, 6) = TEMP_MATRIX(NSIZE + 1, 6) + TEMP_MATRIX(i, 6)
    
    TEMP_MATRIX(i, 7) = DATA_MATRIX(i, 1)
    
    TEMP_MATRIX(i, 8) = TEMP_MATRIX(i, 7) * TEMP_MATRIX(i, 2)
    TEMP_MATRIX(NSIZE + 1, 8) = TEMP_MATRIX(NSIZE + 1, 8) + TEMP_MATRIX(i, 8)
    
    TEMP_MATRIX(i, 9) = TEMP_MATRIX(i, 8) - TEMP_MATRIX(i, 6)
    TEMP_MATRIX(i, 10) = IIf((TEMP_MATRIX(i, 6) > 0), TEMP_MATRIX(i, 9) / TEMP_MATRIX(i, 6), 0)
    
    TEMP_MATRIX(i, 11) = YAHOO_HISTORICAL_DATA_LAST_PRICE_FUNC(INDEX_STR, TEMP_MATRIX(i, 5))
    TEMP_MATRIX(i, 12) = IIf((TEMP_MATRIX(i, 6) > 0) And (TEMP_MATRIX(i, 11) > 0), TEMP_MATRIX(i, 6) / TEMP_MATRIX(i, 11), 0)
    TEMP_MATRIX(i, 13) = TEMP_MATRIX(i, 12) * INDEX_PRICE
    TEMP_MATRIX(NSIZE + 1, 13) = TEMP_MATRIX(NSIZE + 1, 13) + TEMP_MATRIX(i, 13)
    
Next i

TEMP_MATRIX(NSIZE + 1, 19) = 0
TEMP_MATRIX(NSIZE + 1, 22) = 0
TEMP_MATRIX(NSIZE + 1, 28) = 0

For i = 1 To UBound(SALE_TICKERS_VECTOR, 1)
    TEMP_MATRIX(i, 14) = SALE_TICKERS_VECTOR(i, 1)
    TEMP_MATRIX(i, 15) = SALE_QUANTITY_VECTOR(i, 1)
    TEMP_MATRIX(i, 16) = SALE_BUY_PRICE_VECTOR(i, 1)
    TEMP_MATRIX(i, 17) = SALE_BUY_COMMISSION_VECTOR(i, 1)
    TEMP_MATRIX(i, 18) = SALE_BUY_DATE_VECTOR(i, 1)
    TEMP_MATRIX(i, 19) = TEMP_MATRIX(i, 15) * TEMP_MATRIX(i, 16) + TEMP_MATRIX(i, 17)
    TEMP_MATRIX(NSIZE + 1, 19) = TEMP_MATRIX(NSIZE + 1, 19) + TEMP_MATRIX(i, 19)
    
    TEMP_MATRIX(i, 20) = SALE_PRICE_VECTOR(i, 1)
    TEMP_MATRIX(i, 21) = SALE_DATE_VECTOR(i, 1)
    TEMP_MATRIX(i, 22) = TEMP_MATRIX(i, 20) * TEMP_MATRIX(i, 15)
    TEMP_MATRIX(NSIZE + 1, 22) = TEMP_MATRIX(NSIZE + 1, 22) + TEMP_MATRIX(i, 22)
    
    TEMP_MATRIX(i, 23) = TEMP_MATRIX(i, 22) - TEMP_MATRIX(i, 19)
    TEMP_MATRIX(i, 24) = IIf((TEMP_MATRIX(i, 19) > 0), TEMP_MATRIX(i, 23) / TEMP_MATRIX(i, 19), 0)
    TEMP_MATRIX(i, 25) = YAHOO_HISTORICAL_DATA_LAST_PRICE_FUNC(INDEX_STR, TEMP_MATRIX(i, 18))
    TEMP_MATRIX(i, 26) = YAHOO_HISTORICAL_DATA_LAST_PRICE_FUNC(INDEX_STR, TEMP_MATRIX(i, 21))
    TEMP_MATRIX(i, 27) = IIf((TEMP_MATRIX(i, 19) > 0) And (TEMP_MATRIX(i, 25) > 0), TEMP_MATRIX(i, 19) / TEMP_MATRIX(i, 25), 0)
    TEMP_MATRIX(i, 28) = TEMP_MATRIX(i, 26) * TEMP_MATRIX(i, 27)
    TEMP_MATRIX(NSIZE + 1, 28) = TEMP_MATRIX(NSIZE + 1, 28) + TEMP_MATRIX(i, 28)
Next i

For j = 1 To 28
    For i = 0 To NSIZE + 1
        If TEMP_MATRIX(i, j) = 0 Then: TEMP_MATRIX(i, j) = ""
    Next i
Next j

'-------------------------------------------------------------------------
Select Case OUTPUT
'-------------------------------------------------------------------------
Case 0
'-------------------------------------------------------------------------
    ReDim TEMP_VECTOR(1 To 6, 1 To 3)
    TEMP_VECTOR(1, 1) = "Buy Return"
    TEMP_VECTOR(1, 2) = TEMP_MATRIX(NSIZE + 1, 8) - TEMP_MATRIX(NSIZE + 1, 6)
    TEMP_VECTOR(1, 3) = TEMP_VECTOR(1, 2) / TEMP_MATRIX(NSIZE + 1, 6)
    
    TEMP_VECTOR(2, 1) = "Buy Return vs Self"
    TEMP_VECTOR(2, 2) = TEMP_MATRIX(NSIZE + 1, 13) - TEMP_MATRIX(NSIZE + 1, 6)
    TEMP_VECTOR(2, 3) = TEMP_VECTOR(2, 2) / TEMP_MATRIX(NSIZE + 1, 6)
    
    TEMP_VECTOR(3, 1) = "Sale Return"
    TEMP_VECTOR(3, 2) = TEMP_MATRIX(NSIZE + 1, 22) - TEMP_MATRIX(NSIZE + 1, 19)
    TEMP_VECTOR(3, 3) = TEMP_VECTOR(3, 2) / TEMP_MATRIX(NSIZE + 1, 19)
    
    TEMP_VECTOR(4, 1) = "Sale Return vs Self"
    TEMP_VECTOR(4, 2) = TEMP_MATRIX(NSIZE + 1, 28) - TEMP_MATRIX(NSIZE + 1, 19)
    TEMP_VECTOR(4, 3) = TEMP_VECTOR(4, 2) / TEMP_MATRIX(NSIZE + 1, 19)
    
    TEMP_VECTOR(5, 1) = "Total Portfolio Return as of " & Format(CURRENT_DATE, "dd-mm-yy")
    TEMP_VECTOR(5, 2) = TEMP_VECTOR(1, 2) + TEMP_VECTOR(3, 2)
    TEMP_VECTOR(5, 3) = (TEMP_VECTOR(1, 2) + TEMP_VECTOR(3, 2)) / (TEMP_MATRIX(NSIZE + 1, 6) + TEMP_MATRIX(NSIZE + 1, 19))
    
    TEMP_VECTOR(6, 1) = "Return of " & INDEX_STR & " as of " & Format(CURRENT_DATE, "dd-mm-yy")
    TEMP_VECTOR(6, 2) = TEMP_VECTOR(2, 2) + TEMP_VECTOR(4, 2)
    TEMP_VECTOR(6, 3) = (TEMP_VECTOR(4, 2) + TEMP_VECTOR(2, 2)) / (TEMP_MATRIX(NSIZE + 1, 19) + TEMP_MATRIX(NSIZE + 1, 6))
        
    PORT_TRACK_RELATIVE_PERFORMANCE_FUNC = TEMP_VECTOR
'-------------------------------------------------------------------------
Case Else
'-------------------------------------------------------------------------
    PORT_TRACK_RELATIVE_PERFORMANCE_FUNC = TEMP_MATRIX
'-------------------------------------------------------------------------
End Select
'-------------------------------------------------------------------------

Exit Function
ERROR_LABEL:
PORT_TRACK_RELATIVE_PERFORMANCE_FUNC = Err.number
End Function

