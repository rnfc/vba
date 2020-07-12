Attribute VB_Name = "FINAN_ASSET_FUND_CEF_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.


'************************************************************************************
'************************************************************************************
'FUNCTION      : ASSET_CEF_DISCOUNT_FUNC
'DESCRIPTION   :
'In a standard, garden variety, open-ended mutual fund, the fund manager issues new
'units whenever somebuddy wants to buy shares. The manager may have to buy more stocks
'with these extra monies - at an inconvenient time ... like when stocks are overpriced.
'Further, if the market tanks and Sam wants to redeem his mutual funds (which he does via
'the fund company, not on the open market), the fund manager may have to sell stocks to
'raise the monies necessary to accommodate Sam. The value of a mutual fund unit, the Net
'Asset Value or NAV, is determined at the end of each day by the value of stocks owned by
'the fund company, the company's cash reserve, its debt and the management expense ratio.

'There are also Closed End Funds, or CEFs, where the number of shares is fixed ... and therefore
'limited. CEFs, however, can be bought and sold like stocks and may have a price above or below
'the NAV. Because their price depends upon the market forces of supply and demand (remember that
'there are a limited number of shares), they tend to be more volatile than the NAV ... and they
'can be traded throughout the day.

'These CEFs also have a NAV: the value of the stocks held divided by the (fixed) number of shares
'outstanding. Let's generate a spreadsheet that downloads the CEF and the NAV and displays the
'percentage that the CEF price is below the NAV price. That's the "discount", but the CEF may
'also trade at a premium to the NAV and ...

'You type in a Yahoo CEF symbol (like jps) and the associated NAV (like XjpsX) and click the button.
'The daily closing CEF prices will be downloaded from Yahoo and saved, for the past year.
'Then the end-of-day NAV prices will be downloaded.
'Both are plotted as well as the discount (or premium) and ...

'REFERENCES:
'http://67.220.225.70/~gumm5981/CEFs.htm
'http://www.cefconnect.com/

'LIBRARY       : FINAN_ASSET
'GROUP         : MF_CEF
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'UPDATE        : 04/01/2009
'************************************************************************************
'************************************************************************************

Function ASSET_CEF_DISCOUNT_FUNC(ByVal CEF_TICKER As Variant, _
ByVal NAV_TICKER As Variant, _
Optional ByVal START_DATE As Date, _
Optional ByVal END_DATE As Date)

Dim i As Long
Dim NROWS As Long

Dim CEF_VECTOR As Variant
Dim NAV_VECTOR As Variant
Dim TEMP_MATRIX As Variant

On Error GoTo ERROR_LABEL
'---------------------------------------------------------------------------------------------
If IsArray(CEF_TICKER) = False Then
    CEF_VECTOR = YAHOO_HISTORICAL_DATA_SERIE_FUNC(CEF_TICKER, START_DATE, END_DATE, "d", "DA", False, True, True)
Else
    CEF_VECTOR = CEF_TICKER
End If
'---------------------------------------------------------------------------------------------
If IsArray(NAV_TICKER) = False Then
    NAV_VECTOR = YAHOO_HISTORICAL_DATA_SERIE_FUNC(NAV_TICKER, START_DATE, END_DATE, "d", "DA", False, True, True)
Else
    NAV_VECTOR = NAV_TICKER
End If
'---------------------------------------------------------------------------------------------
If UBound(CEF_VECTOR, 1) <> UBound(NAV_VECTOR, 1) Then: GoTo ERROR_LABEL
'---------------------------------------------------------------------------------------------
NROWS = UBound(CEF_VECTOR, 1)
'---------------------------------------------------------------------------------------------

ReDim TEMP_MATRIX(0 To NROWS, 1 To 5)

TEMP_MATRIX(0, 1) = "DATE"
TEMP_MATRIX(0, 2) = UCase(CEF_TICKER)
TEMP_MATRIX(0, 3) = UCase(NAV_TICKER)
TEMP_MATRIX(0, 4) = "DISCOUNT"
TEMP_MATRIX(0, 5) = "CEF RETURN"

i = 1
TEMP_MATRIX(i, 1) = CEF_VECTOR(i, 1)
TEMP_MATRIX(i, 2) = CEF_VECTOR(i, 2)
TEMP_MATRIX(i, 3) = NAV_VECTOR(i, 2)
TEMP_MATRIX(i, 4) = 1 - TEMP_MATRIX(i, 2) / TEMP_MATRIX(i, 3)
TEMP_MATRIX(i, 5) = ""

For i = 2 To NROWS
    TEMP_MATRIX(i, 1) = CEF_VECTOR(i, 1)
    TEMP_MATRIX(i, 2) = CEF_VECTOR(i, 2)
    TEMP_MATRIX(i, 3) = NAV_VECTOR(i, 2)
    TEMP_MATRIX(i, 4) = 1 - TEMP_MATRIX(i, 2) / TEMP_MATRIX(i, 3)
    TEMP_MATRIX(i, 5) = TEMP_MATRIX(i, 2) / TEMP_MATRIX(i - 1, 2) - 1
Next i

ASSET_CEF_DISCOUNT_FUNC = TEMP_MATRIX

Exit Function
ERROR_LABEL:
ASSET_CEF_DISCOUNT_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : ASSETS_CEFS_PERFORMANCE_FUNC
'DESCRIPTION   :
'LIBRARY       : FINAN_ASSET
'GROUP         : MF_CEF
'ID            : 002
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'UPDATE        : 04/01/2009
'************************************************************************************
'************************************************************************************

Function ASSETS_CEFS_PERFORMANCE_FUNC(ByRef CEF_TICKERS_RNG As Variant, _
ByRef NAV_TICKERS_RNG As Variant, _
Optional ByVal START_DATE As Date, _
Optional ByVal END_DATE As Date, _
Optional ByVal MA_PERIOD As Long = 200, _
Optional ByVal COUNT_BASIS As Double = 252)

Dim i As Long
Dim j As Long

Dim NROWS As Long
Dim NCOLUMNS As Long

Dim TEMP1_SUM As Double
Dim TEMP2_SUM As Double

Dim DISCOUNT_VAL As Double

Dim CEF_TICKER As String
Dim NAV_TICKER As String

Dim CEF_VECTOR As Variant
Dim NAV_VECTOR As Variant

Dim TEMP_MATRIX As Variant
Dim CEF_TICKERS_VECTOR As Variant
Dim NAV_TICKERS_VECTOR As Variant

On Error GoTo ERROR_LABEL

If IsArray(CEF_TICKERS_RNG) = True Then
    CEF_TICKERS_VECTOR = CEF_TICKERS_RNG
    If UBound(CEF_TICKERS_VECTOR, 1) = 1 Then
        CEF_TICKERS_VECTOR = MATRIX_TRANSPOSE_FUNC(CEF_TICKERS_VECTOR)
    End If
Else
    ReDim CEF_TICKERS_VECTOR(1 To 1, 1 To 1)
    CEF_TICKERS_VECTOR(1, 1) = CEF_TICKERS_RNG
End If

If IsArray(NAV_TICKERS_RNG) = True Then
    NAV_TICKERS_VECTOR = NAV_TICKERS_RNG
    If UBound(NAV_TICKERS_VECTOR, 1) = 1 Then
        NAV_TICKERS_VECTOR = MATRIX_TRANSPOSE_FUNC(NAV_TICKERS_VECTOR)
    End If
Else
    ReDim NAV_TICKERS_VECTOR(1 To 1, 1 To 1)
    NAV_TICKERS_VECTOR(1, 1) = NAV_TICKERS_RNG
End If

If UBound(CEF_TICKERS_VECTOR, 1) <> UBound(NAV_TICKERS_VECTOR, 1) Then: GoTo ERROR_LABEL
NCOLUMNS = UBound(CEF_TICKERS_VECTOR, 1)

ReDim TEMP_MATRIX(0 To NCOLUMNS, 1 To 11)

TEMP_MATRIX(0, 1) = "CEF TICKER"
TEMP_MATRIX(0, 2) = "NAV TICKER"
TEMP_MATRIX(0, 3) = "START DATE"
TEMP_MATRIX(0, 4) = "END DATE"
TEMP_MATRIX(0, 5) = "CURRENT CEF"
TEMP_MATRIX(0, 6) = "CURRENT NAV"
TEMP_MATRIX(0, 7) = "CURRENT DISCOUNT"
TEMP_MATRIX(0, 8) = MA_PERIOD & "-DAY AVERAGE DISCOUNT"
TEMP_MATRIX(0, 9) = "ANNUAL CEF RETURN"
TEMP_MATRIX(0, 10) = "ANNUAL CEF VOLATILITY"
TEMP_MATRIX(0, 11) = "ANNUAL SHARPE"

For j = 1 To NCOLUMNS

    CEF_TICKER = CEF_TICKERS_VECTOR(j, 1)
    NAV_TICKER = NAV_TICKERS_VECTOR(j, 1)
    
    TEMP_MATRIX(j, 1) = UCase(CEF_TICKER)
    TEMP_MATRIX(j, 2) = UCase(NAV_TICKER)

    CEF_VECTOR = YAHOO_HISTORICAL_DATA_SERIE_FUNC(CEF_TICKER, START_DATE, END_DATE, "d", "DA", False, True, True)
    If IsArray(CEF_VECTOR) = False Then: GoTo 1983
    NAV_VECTOR = YAHOO_HISTORICAL_DATA_SERIE_FUNC(NAV_TICKER, START_DATE, END_DATE, "d", "DA", False, True, True)
    If IsArray(NAV_VECTOR) = False Then: GoTo 1983
    If UBound(CEF_VECTOR, 1) <> UBound(NAV_VECTOR, 1) Then: GoTo 1983
    NROWS = UBound(CEF_VECTOR, 1)

    TEMP_MATRIX(j, 3) = CEF_VECTOR(1, 1)
    TEMP_MATRIX(j, 4) = CEF_VECTOR(NROWS, 1)

    TEMP_MATRIX(j, 5) = CEF_VECTOR(NROWS, 2)
    TEMP_MATRIX(j, 6) = NAV_VECTOR(NROWS, 2)
    TEMP_MATRIX(j, 7) = 1 - TEMP_MATRIX(j, 5) / TEMP_MATRIX(j, 6)
    
    TEMP1_SUM = 0: TEMP2_SUM = 0
    For i = 2 To NROWS
        DISCOUNT_VAL = CEF_VECTOR(i, 2) / CEF_VECTOR(i - 1, 2) - 1
        TEMP2_SUM = TEMP2_SUM + DISCOUNT_VAL
        DISCOUNT_VAL = 1 - CEF_VECTOR(i, 2) / NAV_VECTOR(i, 2)
        If i > NROWS - MA_PERIOD + 1 Then
            TEMP1_SUM = TEMP1_SUM + DISCOUNT_VAL
        ElseIf i = NROWS - MA_PERIOD + 1 Then
            TEMP1_SUM = DISCOUNT_VAL
        End If
    Next i
    TEMP_MATRIX(j, 8) = TEMP1_SUM / MA_PERIOD 'MA Average
    TEMP_MATRIX(j, 9) = TEMP2_SUM / (NROWS - 1) 'Average
    TEMP_MATRIX(j, 10) = 0
    For i = 2 To NROWS
        DISCOUNT_VAL = CEF_VECTOR(i, 2) / CEF_VECTOR(i - 1, 2) - 1
        TEMP_MATRIX(j, 10) = TEMP_MATRIX(j, 10) + (DISCOUNT_VAL - TEMP_MATRIX(j, 9)) ^ 2
    Next i
    TEMP_MATRIX(j, 10) = (TEMP_MATRIX(j, 10) / (NROWS - 1)) ^ 0.5 'Sigma
    
    TEMP_MATRIX(j, 9) = TEMP_MATRIX(j, 9) * COUNT_BASIS
    TEMP_MATRIX(j, 10) = TEMP_MATRIX(j, 10) * COUNT_BASIS ^ 0.5
    If TEMP_MATRIX(j, 10) <> 0 Then
        TEMP_MATRIX(j, 11) = TEMP_MATRIX(j, 9) / TEMP_MATRIX(j, 10)
    Else
        TEMP_MATRIX(j, 11) = "N/A"
    End If
    
1983:
Next j

ASSETS_CEFS_PERFORMANCE_FUNC = TEMP_MATRIX

Exit Function
ERROR_LABEL:
ASSETS_CEFS_PERFORMANCE_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : ASSET_MF_MER_FUNC
'DESCRIPTION   : MEASURE MUTUAL FUND PERFORMANCE
'http://67.220.225.70/~gumm5981/MERs_and_stuff.htm
'http://67.220.225.70/~gumm5981/toro.htm

'LIBRARY       : FINAN_ASSET
'GROUP         : MF_CEF
'ID            : 003
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'UPDATE        : 04/01/2009
'************************************************************************************
'************************************************************************************

Function ASSET_MF_MER_FUNC(ByRef TICKERS_RNG As Variant, _
ByRef SIZE_FUND_RNG As Variant, _
ByRef FRONT_END_RNG As Variant, _
ByRef BACK_END_RNG As Variant, _
ByRef ANNUAL_FEES_RNG As Variant, _
ByRef SGA_EXPENSES_RNG As Variant)

'Front-end loads (initial payments) %
'Back-end loads (redemption fees) %
'Annual fees %
'Marketing and promotion expenses %

Dim i As Long
Dim NROWS As Long

Dim TEMP_MATRIX As Variant
Dim FUND_EXPENSES As Double

Dim TICKERS_VECTOR As Variant
Dim SIZE_FUND_VECTOR As Variant
Dim FRONT_END_VECTOR As Variant
Dim BACK_END_VECTOR As Variant
Dim ANNUAL_FEES_VECTOR As Variant
Dim SGA_EXPENSES_VECTOR As Variant

On Error GoTo ERROR_LABEL

TICKERS_VECTOR = TICKERS_RNG
If UBound(TICKERS_VECTOR, 1) = 1 Then
    TICKERS_VECTOR = MATRIX_TRANSPOSE_FUNC(TICKERS_VECTOR)
End If
NROWS = UBound(TICKERS_VECTOR, 1)

SIZE_FUND_VECTOR = SIZE_FUND_RNG
If UBound(SIZE_FUND_VECTOR, 1) = 1 Then
    SIZE_FUND_VECTOR = MATRIX_TRANSPOSE_FUNC(SIZE_FUND_VECTOR)
End If
If UBound(SIZE_FUND_VECTOR, 1) <> NROWS Then: GoTo ERROR_LABEL

FRONT_END_VECTOR = FRONT_END_RNG
If UBound(FRONT_END_VECTOR, 1) = 1 Then
    FRONT_END_VECTOR = MATRIX_TRANSPOSE_FUNC(FRONT_END_VECTOR)
End If
If UBound(FRONT_END_VECTOR, 1) <> NROWS Then: GoTo ERROR_LABEL

BACK_END_VECTOR = BACK_END_RNG
If UBound(BACK_END_VECTOR, 1) = 1 Then
    BACK_END_VECTOR = MATRIX_TRANSPOSE_FUNC(BACK_END_VECTOR)
End If
If UBound(BACK_END_VECTOR, 1) <> NROWS Then: GoTo ERROR_LABEL

ANNUAL_FEES_VECTOR = ANNUAL_FEES_RNG
If UBound(ANNUAL_FEES_VECTOR, 1) = 1 Then
    ANNUAL_FEES_VECTOR = MATRIX_TRANSPOSE_FUNC(ANNUAL_FEES_VECTOR)
End If
If UBound(ANNUAL_FEES_VECTOR, 1) <> NROWS Then: GoTo ERROR_LABEL

SGA_EXPENSES_VECTOR = SGA_EXPENSES_RNG
If UBound(SGA_EXPENSES_VECTOR, 1) = 1 Then
    SGA_EXPENSES_VECTOR = MATRIX_TRANSPOSE_FUNC(SGA_EXPENSES_VECTOR)
End If
If UBound(SGA_EXPENSES_VECTOR, 1) <> NROWS Then: GoTo ERROR_LABEL

ReDim TEMP_MATRIX(0 To NROWS, 1 To 5)

TEMP_MATRIX(0, 1) = "TICKER"
TEMP_MATRIX(0, 2) = "FUND ASSET"
TEMP_MATRIX(0, 3) = "FUND EXPENSES %"
TEMP_MATRIX(0, 4) = "FUND LIABILTIES"
TEMP_MATRIX(0, 5) = "NET ASSET VALUE" 'Changes in NAV (Capital Gain) + Dividends  = Return

For i = 1 To NROWS
    TEMP_MATRIX(i, 1) = TICKERS_VECTOR(i, 1)
    FUND_EXPENSES = (FRONT_END_VECTOR(i, 1) + BACK_END_VECTOR(i, 1) + ANNUAL_FEES_VECTOR(i, 1) + SGA_EXPENSES_VECTOR(i, 1))
    TEMP_MATRIX(i, 2) = SIZE_FUND_VECTOR(i, 1) 'Assets
    TEMP_MATRIX(i, 3) = FUND_EXPENSES
    'NEW_FUND_SIZE = SIZE_FUND_VECTOR(i, 1) * (1 + CAPITAL_GAIN)
    TEMP_MATRIX(i, 4) = TEMP_MATRIX(i, 2) * TEMP_MATRIX(i, 3) 'Liabilities
    TEMP_MATRIX(i, 5) = TEMP_MATRIX(i, 2) - TEMP_MATRIX(i, 4)
    'RETURN = (NEW_NAV / CURRENT_NAV) - 1 + DIVIDENDS
Next i

ASSET_MF_MER_FUNC = TEMP_MATRIX

Exit Function
ERROR_LABEL:
ASSET_MF_MER_FUNC = Err.number
End Function
