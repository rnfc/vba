Attribute VB_Name = "FINAN_ASSET_FUND_BETA_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.


Function MUTUAL_FUNDS_BETA_PERFORMANCE_FUNC(ByRef TICKERS_RNG As Variant)

Dim i As Long
Dim j As Long
Dim k As Long
Dim NROWS As Long

Const ERROR_STR As String = "--"
Const NA_STR As String = "#N/A"
Dim TICKER_STR As String
Dim SRC_URL_STR As String

Dim TEMP_VAL As Variant
Dim TICKERS_VECTOR As Variant
Dim TEMP_MATRIX As Variant

On Error GoTo ERROR_LABEL

If IsArray(TICKERS_RNG) = True Then
    TICKERS_VECTOR = TICKERS_RNG
    If UBound(TICKERS_VECTOR, 1) = 1 Then
        TICKERS_VECTOR = MATRIX_TRANSPOSE_FUNC(TICKERS_VECTOR)
    End If
Else
    ReDim TICKERS_VECTOR(1 To 1, 1 To 1)
    TICKERS_VECTOR(1, 1) = TICKERS_RNG
End If
NROWS = UBound(TICKERS_VECTOR, 1)

ReDim TEMP_MATRIX(0 To NROWS, 1 To 27)

TEMP_MATRIX(0, 1) = "FUND"
TEMP_MATRIX(0, 2) = "NAME"
TEMP_MATRIX(0, 3) = "MS CATEGORY"
TEMP_MATRIX(0, 4) = "AUM ($M)"
TEMP_MATRIX(0, 5) = "MS STARS"
TEMP_MATRIX(0, 6) = "BETA 3YR"
TEMP_MATRIX(0, 7) = "ALPHA 3YR"

TEMP_MATRIX(0, 17) = Format(Now, "YYYY") - 1
For j = 16 To 8 Step -1: TEMP_MATRIX(0, j) = TEMP_MATRIX(0, j + 1) - 1: Next j

TEMP_MATRIX(0, 18) = "YTD"
TEMP_MATRIX(0, 19) = TEMP_MATRIX(0, 16) & "-" & TEMP_MATRIX(0, 17) + 1
TEMP_MATRIX(0, 20) = "3YR RETURN"
TEMP_MATRIX(0, 21) = "5YR RETURN"
TEMP_MATRIX(0, 22) = "10YR RETURN"
TEMP_MATRIX(0, 23) = "5YR STDEV"
TEMP_MATRIX(0, 24) = "BETA 5YR"
TEMP_MATRIX(0, 25) = "ALPHA 5YR"
TEMP_MATRIX(0, 26) = "BETA 10YR"
TEMP_MATRIX(0, 27) = "ALPHA 10YR"

SRC_URL_STR = "http://moneycentral.msn.com/investor/partsub/funds/returns.asp?Symbol="
For i = 1 To NROWS
    TICKER_STR = TICKERS_VECTOR(i, 1)
    If TICKER_STR = "" Then: GoTo 1984
    
    j = 1
    TEMP_MATRIX(i, j) = TICKER_STR
    
    TEMP_VAL = RETRIEVE_WEB_DATA_ELEMENT_FUNC(TICKER_STR, 13862, ERROR_STR) 'Name
    j = j + 1
    If TEMP_VAL <> ERROR_STR Then
        TEMP_MATRIX(i, j) = TEMP_VAL
    Else
        TEMP_MATRIX(i, j) = NA_STR
    End If
    
    TEMP_VAL = RETRIEVE_WEB_DATA_ELEMENT_FUNC(TICKER_STR, 5059, ERROR_STR) 'Category
    j = j + 1
    If TEMP_VAL <> ERROR_STR Then
        TEMP_MATRIX(i, j) = TEMP_VAL
    Else
        TEMP_MATRIX(i, j) = NA_STR
    End If
    
    TEMP_VAL = RETRIEVE_WEB_DATA_ELEMENT_FUNC(TICKER_STR, 5061, ERROR_STR) 'Aum
    j = j + 1
    If TEMP_VAL <> ERROR_STR And IsNumeric(TEMP_VAL) Then
        TEMP_VAL = TEMP_VAL / 1000
        TEMP_MATRIX(i, j) = TEMP_VAL
    Else
        TEMP_MATRIX(i, j) = NA_STR
    End If
    
    TEMP_VAL = RETRIEVE_WEB_DATA_ELEMENT_FUNC(TICKER_STR, 5064, ERROR_STR) 'Ms Stars
    j = j + 1
    If TEMP_VAL <> ERROR_STR And IsNumeric(TEMP_VAL) Then
        TEMP_MATRIX(i, j) = TEMP_VAL
    Else
        TEMP_MATRIX(i, j) = NA_STR
    End If
    
    TEMP_VAL = RETRIEVE_WEB_DATA_ELEMENT_FUNC(TICKER_STR, 5099, ERROR_STR) 'Beta 3yr
    j = j + 1
    If TEMP_VAL <> ERROR_STR And IsNumeric(TEMP_VAL) Then
        TEMP_MATRIX(i, j) = TEMP_VAL
    Else
        TEMP_MATRIX(i, j) = NA_STR
    End If
    
    TEMP_VAL = RETRIEVE_WEB_DATA_ELEMENT_FUNC(TICKER_STR, 5098, ERROR_STR) 'Alpha 3yr
    j = j + 1
    If TEMP_VAL <> ERROR_STR And IsNumeric(TEMP_VAL) Then
        TEMP_MATRIX(i, j) = TEMP_VAL
    Else
        TEMP_MATRIX(i, j) = NA_STR
    End If
    
    For k = -11 To -2
        TEMP_VAL = RETRIEVE_WEB_DATA_CELL_FUNC(SRC_URL_STR & TICKER_STR, k, _
                   "Calendar-Year Total Returns", "Total Return %", , , , , , ERROR_STR)
        j = j + 1
        If TEMP_VAL <> ERROR_STR And IsNumeric(TEMP_VAL) Then
            TEMP_VAL = TEMP_VAL / 100
            TEMP_MATRIX(i, j) = TEMP_VAL
        Else
            TEMP_MATRIX(i, j) = NA_STR
        End If
    Next k
    
    TEMP_VAL = RETRIEVE_WEB_DATA_CELL_FUNC(SRC_URL_STR & TICKER_STR, 1, "Year-to-date", , , , , , , ERROR_STR)
    j = j + 1
    If TEMP_VAL <> ERROR_STR And IsNumeric(TEMP_VAL) Then
        TEMP_VAL = TEMP_VAL / 100
        TEMP_MATRIX(i, j) = TEMP_VAL
    Else
        TEMP_MATRIX(i, j) = NA_STR
    End If
    
    If IsNumeric(TEMP_MATRIX(i, j - 2)) = False Then: GoTo 1983
    If IsNumeric(TEMP_MATRIX(i, j - 1)) = False Then: GoTo 1983
    If IsNumeric(TEMP_MATRIX(i, j - 0)) = False Then: GoTo 1983
    
    TEMP_VAL = (1 + TEMP_MATRIX(i, j - 2)) * (1 + TEMP_MATRIX(i, j - 1)) * (1 + TEMP_MATRIX(i, j - 0)) - 1 '3years Chained-link return
    j = j + 1
    If TEMP_VAL <> ERROR_STR And IsNumeric(TEMP_VAL) Then
        TEMP_MATRIX(i, j) = TEMP_VAL
    Else
1983:
        TEMP_MATRIX(i, j) = NA_STR
    End If
    
    TEMP_VAL = RETRIEVE_WEB_DATA_CELL_FUNC(SRC_URL_STR & TICKER_STR, 1, "3-Year Annualized", , , , , , , ERROR_STR) '3yr Return
    j = j + 1
    If TEMP_VAL <> ERROR_STR And IsNumeric(TEMP_VAL) Then
        TEMP_VAL = TEMP_VAL / 100
        TEMP_MATRIX(i, j) = TEMP_VAL
    Else
        TEMP_MATRIX(i, j) = NA_STR
    End If
    
    TEMP_VAL = RETRIEVE_WEB_DATA_CELL_FUNC(SRC_URL_STR & TICKER_STR, 1, "5-Year Annualized", , , , , , , ERROR_STR) '5yr Return
    j = j + 1
    If TEMP_VAL <> ERROR_STR And IsNumeric(TEMP_VAL) Then
        TEMP_VAL = TEMP_VAL / 100
        TEMP_MATRIX(i, j) = TEMP_VAL
    Else
        TEMP_MATRIX(i, j) = NA_STR
    End If
    
    TEMP_VAL = RETRIEVE_WEB_DATA_CELL_FUNC(SRC_URL_STR & TICKER_STR, 1, "10-Year Annualized", , , , , , , ERROR_STR) '10yr Return
    j = j + 1
    If TEMP_VAL <> ERROR_STR And IsNumeric(TEMP_VAL) Then
        TEMP_VAL = TEMP_VAL / 100
        TEMP_MATRIX(i, j) = TEMP_VAL
    Else
        TEMP_MATRIX(i, j) = NA_STR
    End If
    
    TEMP_VAL = RETRIEVE_WEB_DATA_ELEMENT_FUNC(TICKER_STR, 5109, ERROR_STR) '5yr StDev
    j = j + 1
    If TEMP_VAL <> ERROR_STR And IsNumeric(TEMP_VAL) Then
        TEMP_VAL = TEMP_VAL / 100
        TEMP_MATRIX(i, j) = TEMP_VAL
    Else
        TEMP_MATRIX(i, j) = NA_STR
    End If
    
    TEMP_VAL = RETRIEVE_WEB_DATA_ELEMENT_FUNC(TICKER_STR, 5106, ERROR_STR) 'Beta 5yr
    j = j + 1
    If TEMP_VAL <> ERROR_STR And IsNumeric(TEMP_VAL) Then
        TEMP_MATRIX(i, j) = TEMP_VAL
    Else
        TEMP_MATRIX(i, j) = NA_STR
    End If
    
    TEMP_VAL = RETRIEVE_WEB_DATA_ELEMENT_FUNC(TICKER_STR, 5105, ERROR_STR) 'Alpha 5yr
    j = j + 1
    If TEMP_VAL <> ERROR_STR And IsNumeric(TEMP_VAL) Then
        TEMP_MATRIX(i, j) = TEMP_VAL
    Else
        TEMP_MATRIX(i, j) = NA_STR
    End If
    
    TEMP_VAL = RETRIEVE_WEB_DATA_ELEMENT_FUNC(TICKER_STR, 5113, ERROR_STR) 'Beta 10yr
    j = j + 1
    If TEMP_VAL <> ERROR_STR And IsNumeric(TEMP_VAL) Then
        TEMP_MATRIX(i, j) = TEMP_VAL
    Else
        TEMP_MATRIX(i, j) = NA_STR
    End If
    
    TEMP_VAL = RETRIEVE_WEB_DATA_ELEMENT_FUNC(TICKER_STR, 5112, ERROR_STR) 'Alpha 10yr
    j = j + 1
    If TEMP_VAL <> ERROR_STR And IsNumeric(TEMP_VAL) Then
        TEMP_MATRIX(i, j) = TEMP_VAL
    Else
        TEMP_MATRIX(i, j) = NA_STR
    End If
1984:
Next i

MUTUAL_FUNDS_BETA_PERFORMANCE_FUNC = TEMP_MATRIX

Exit Function
ERROR_LABEL:
MUTUAL_FUNDS_BETA_PERFORMANCE_FUNC = Err.number
End Function
