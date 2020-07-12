Attribute VB_Name = "FINAN_ASSET_FUTURES_FAIR_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.

Function FUTURES_INDEX_FAIR_PRICE_FUNC(ByVal RISK_FREE_RATE_VAL As Double, _
ByVal INDEX_QUOTE_VAL As Double, _
ByRef COMPONENTS_RNG As Variant, _
Optional ByVal YEARFRAC_VAL As Double = 1, _
Optional ByVal VERSION As Integer = 0)

'INDEX_QUOTE_VAL: Current Value of the Index

'COMPONENTS:
'C1 = Last Trade
'C2 = Dividend Yield

'RISK_FREE_RATE_VAL -> Annual Risk-free Rate -->
'10-YEAR TREASURY NOTE(Chicago Options: ^TNX)

'YEARFRAC_VAL --> Years to Expiry

Dim i As Long
Dim j As Long 'Counter for components
Dim NROWS As Long

Dim TEMP_SUM As Double 'SUM of Prices (components/stocks)
Dim DIVISOR_VAL As Double
'A $1 change in any stock price will change the Index by 1 / DIVISOR_VAL
Dim AVG_YIELD_VAL As Double
Dim DIVIDENDS_VAL As Double 'Total Dividends for X days
Dim INDEX_FAIR_VAL As Double
Dim SPREAD_VAL As Double

Dim COMPONENTS_MATRIX As Variant

On Error GoTo ERROR_LABEL

COMPONENTS_MATRIX = COMPONENTS_RNG
If UBound(COMPONENTS_MATRIX, 2) <> 2 Then: GoTo ERROR_LABEL
NROWS = UBound(COMPONENTS_MATRIX, 1)

j = 0
TEMP_SUM = 0
AVG_YIELD_VAL = 0
For i = 1 To NROWS
    TEMP_SUM = TEMP_SUM + COMPONENTS_MATRIX(i, 1)
'-------------------------------------------------------------------------------
    If VERSION = 0 Then 'Exclude non-dividend components
'-------------------------------------------------------------------------------
        If COMPONENTS_MATRIX(i, 2) <> 0 Then
            AVG_YIELD_VAL = AVG_YIELD_VAL + COMPONENTS_MATRIX(i, 2)
            j = j + 1
        End If
'-------------------------------------------------------------------------------
    Else
'-------------------------------------------------------------------------------
        AVG_YIELD_VAL = AVG_YIELD_VAL + COMPONENTS_MATRIX(i, 2)
        j = j + 1
'-------------------------------------------------------------------------------
    End If
'-------------------------------------------------------------------------------
Next i

DIVISOR_VAL = TEMP_SUM / INDEX_QUOTE_VAL
AVG_YIELD_VAL = AVG_YIELD_VAL / j
DIVIDENDS_VAL = TEMP_SUM * AVG_YIELD_VAL * YEARFRAC_VAL
INDEX_FAIR_VAL = INDEX_QUOTE_VAL * (1 + RISK_FREE_RATE_VAL * YEARFRAC_VAL) - DIVIDENDS_VAL
SPREAD_VAL = INDEX_FAIR_VAL - INDEX_QUOTE_VAL

FUTURES_INDEX_FAIR_PRICE_FUNC = Array(INDEX_FAIR_VAL, SPREAD_VAL, _
DIVIDENDS_VAL, AVG_YIELD_VAL, DIVISOR_VAL, TEMP_SUM, j, YEARFRAC_VAL)

Exit Function
ERROR_LABEL:
FUTURES_INDEX_FAIR_PRICE_FUNC = Err.Number
End Function


'-------------------------------------------------------------------------------------------------
'-------------------------------------------------------------------------------------------------
'REFERENCES:
'http://www.financialwebring.org/gummy-stuff/DOW-futures.htm
'http://www.cbot.com/cbot/pub/page/0,3181,1165,00.html
'http://www.cbot.com/cbot/pub/cont_detail/1,3206,1719+8708,00.html
'http://www.indexarb.com/dividendYieldSorteddj.html
'-------------------------------------------------------------------------------------------------
'-------------------------------------------------------------------------------------------------