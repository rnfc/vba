Attribute VB_Name = "FINAN_ASSET_DIVIDENDS_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.

'************************************************************************************
'************************************************************************************
'FUNCTION      : ASSET_DIVIDENDS_PORT_GROWTH_FUNC

'DESCRIPTION   : In 1944 Anne Scheiber, a lifelong federal employee whose income never
'surpassed $3,150 a year, invested $5,000 in blue-chip stocks.
'When she died in 1995 her stocks were worth $22 million and she was
'receiving an annual income of over $1 million in dividends from them.
'The book by RoxAnn Klugman (a tax-law attorney and retirement and estate
'planner), Dividend Growth Investment Strategy, tells how she did it.
'The title of the book noted above says:  "How to Keep Your Retirement Income
'Doubling Every Five Years".

'LIBRARY       : ASSET
'GROUP         : DIVIDENDS
'ID            : 001
'LAST UPDATE   : 27/02/2008
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'************************************************************************************
'************************************************************************************

Function ASSET_DIVIDENDS_PORT_GROWTH_FUNC( _
ByVal TICKER_STR As String, _
Optional ByVal START_YEAR As Integer = 0, _
Optional ByVal END_YEAR As Integer = 0, _
Optional ByVal INITIAL_SHARES As Double = 1000, _
Optional ByVal PAYMENTS_PER_YEAR As Integer = 4)

Dim i As Long
Dim NROWS As Long

Dim KEY_STR As String
Dim ITEM_STR As String

Dim DATA_OBJ As Collection
Dim DATA_VECTOR As Variant
Dim DIVIDENDS_VECTOR As Variant
Dim TEMP_MATRIX As Variant

On Error GoTo ERROR_LABEL

DIVIDENDS_VECTOR = YAHOO_HISTORICAL_DATA_MATRIX_FUNC(TICKER_STR, START_YEAR, 1, 1, END_YEAR, 12, 31, "v", "DA", False, False, False, 0, 0)
If IsArray(DIVIDENDS_VECTOR) = False Then: GoTo ERROR_LABEL
DIVIDENDS_VECTOR = MATRIX_QUICK_SORT_FUNC(DIVIDENDS_VECTOR, 1, 1)
NROWS = UBound(DIVIDENDS_VECTOR, 1)

DATA_VECTOR = YAHOO_HISTORICAL_DATA_MATRIX_FUNC(TICKER_STR, START_YEAR, 1, 1, END_YEAR, 12, 31, "d", "DA", False, False, True, 0, 0)
If IsArray(DATA_VECTOR) = False Then: GoTo ERROR_LABEL
Set DATA_OBJ = New Collection
On Error Resume Next
For i = LBound(DATA_VECTOR, 1) To UBound(DATA_VECTOR, 1)
    ITEM_STR = CStr(DATA_VECTOR(i, 2))
    KEY_STR = CStr(DATA_VECTOR(i, 1))
    Call DATA_OBJ.Add(ITEM_STR, KEY_STR)
    If Err.number <> 0 Then: Err.Clear
Next i

ReDim TEMP_MATRIX(0 To NROWS, 1 To 8)

TEMP_MATRIX(0, 1) = ("DIVIDEND DATE")
TEMP_MATRIX(0, 2) = ("QUATERLY DIVIDENDS")
TEMP_MATRIX(0, 3) = ("ANNUAL DIVIDENDS")
TEMP_MATRIX(0, 4) = ("ANNUAL YIELD")
TEMP_MATRIX(0, 5) = ("CLOSING PRICE")
TEMP_MATRIX(0, 6) = ("NEW SHARES")
TEMP_MATRIX(0, 7) = ("TOTAL SHARES")
TEMP_MATRIX(0, 8) = ("PORTFOLIO WITH REINVESTED DIVIDENDS")

For i = 1 To NROWS
        
    TEMP_MATRIX(i, 1) = DIVIDENDS_VECTOR(i, 1)
    TEMP_MATRIX(i, 2) = DIVIDENDS_VECTOR(i, 2)
    TEMP_MATRIX(i, 3) = DIVIDENDS_VECTOR(i, 2) * PAYMENTS_PER_YEAR
    KEY_STR = DIVIDENDS_VECTOR(i, 1)
    ITEM_STR = DATA_OBJ.Item(KEY_STR)
        
    If Err.number = 0 Then
        TEMP_MATRIX(i, 5) = Val(ITEM_STR)
        If TEMP_MATRIX(i, 5) <> 0 Then
            TEMP_MATRIX(i, 4) = TEMP_MATRIX(i, 3) / TEMP_MATRIX(i, 5)
        Else
            TEMP_MATRIX(i, 4) = 0
        End If
        If i > 1 Then
            TEMP_MATRIX(i, 6) = TEMP_MATRIX(i - 1, 7) * TEMP_MATRIX(i, 2) / TEMP_MATRIX(i, 5)
            TEMP_MATRIX(i, 7) = TEMP_MATRIX(i, 6) + TEMP_MATRIX(i - 1, 7)
            TEMP_MATRIX(i, 8) = TEMP_MATRIX(i, 7) * TEMP_MATRIX(i, 5)
        Else
            TEMP_MATRIX(i, 6) = INITIAL_SHARES * TEMP_MATRIX(i, 2) / TEMP_MATRIX(i, 5)
            TEMP_MATRIX(i, 7) = TEMP_MATRIX(i, 6) + INITIAL_SHARES
            TEMP_MATRIX(i, 8) = TEMP_MATRIX(i, 7) * TEMP_MATRIX(i, 5)
        End If
    Else
        Err.Clear: ITEM_STR = 0
        TEMP_MATRIX(i, 4) = 0: TEMP_MATRIX(i, 5) = 0
        If i > 1 Then
            TEMP_MATRIX(i, 6) = 0
            TEMP_MATRIX(i, 7) = TEMP_MATRIX(i, 6) + TEMP_MATRIX(i - 1, 7)
            TEMP_MATRIX(i, 8) = TEMP_MATRIX(i, 7) * TEMP_MATRIX(i, 5)
        Else
            TEMP_MATRIX(i, 6) = 0
            TEMP_MATRIX(i, 7) = TEMP_MATRIX(i, 6) + INITIAL_SHARES
            TEMP_MATRIX(i, 8) = TEMP_MATRIX(i, 7) * TEMP_MATRIX(i, 5)
        End If
    End If
Next i

ASSET_DIVIDENDS_PORT_GROWTH_FUNC = TEMP_MATRIX

Exit Function
ERROR_LABEL:
ASSET_DIVIDENDS_PORT_GROWTH_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : ASSET_DIVIDENDS_BENCHMARK_FUNC

'DESCRIPTION   : The benchmark is based on the Vanguard Total Stock Market
'VIPERs (VTI) which tracks the Wilshire 5000.

'LIBRARY       : ASSET
'GROUP         : DIVIDENDS
'ID            : 002
'LAST UPDATE   : 27/02/2008
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'************************************************************************************
'************************************************************************************

Function ASSET_DIVIDENDS_BENCHMARK_FUNC( _
ByVal TICKER_STR As String, _
Optional ByVal START_YEAR As Integer = 0, _
Optional ByVal END_YEAR As Integer = 0, _
Optional ByVal TDAYS_PER_YEAR As Double = 365)

Dim h As Long
Dim i As Long
Dim j As Long
Dim k As Long
Dim l() As Long

Dim KEY_STR As String
Dim ITEM_STR As String

Dim NROWS As Long
Dim NSIZE As Long

Dim DATA_VECTOR As Variant
Dim DIVIDENDS_OBJ As Collection
Dim DIVIDENDS_VECTOR As Variant

Dim TEMP_MATRIX As Variant

On Error GoTo ERROR_LABEL

DATA_VECTOR = YAHOO_HISTORICAL_DATA_MATRIX_FUNC(TICKER_STR, START_YEAR, 1, 1, END_YEAR, 12, 31, "d", "DA", False, False, True, 0, 0)
If IsArray(DATA_VECTOR) = False Then: GoTo ERROR_LABEL 'http://www.fintools.com/docs/_Dividend%20Adjusted%20Stock%20Prices.pdf
NROWS = UBound(DATA_VECTOR, 1)

DIVIDENDS_VECTOR = YAHOO_HISTORICAL_DATA_MATRIX_FUNC(TICKER_STR, START_YEAR, 1, 1, END_YEAR, 12, 31, "v", "DA", False, False, False, 0, 0)
If IsArray(DIVIDENDS_VECTOR) = False Then: GoTo ERROR_LABEL
DIVIDENDS_VECTOR = MATRIX_QUICK_SORT_FUNC(DIVIDENDS_VECTOR, 1, 1)
Set DIVIDENDS_OBJ = New Collection
On Error Resume Next
For i = LBound(DATA_VECTOR, 1) To UBound(DATA_VECTOR, 1)
    ITEM_STR = CStr(DIVIDENDS_VECTOR(i, 2))
    KEY_STR = CStr(DIVIDENDS_VECTOR(i, 1))
    Call DIVIDENDS_OBJ.Add(ITEM_STR, KEY_STR)
    If Err.number <> 0 Then: Err.Clear
Next i

NSIZE = 1: i = 1
ReDim Preserve l(1 To 2, 1 To NSIZE)
l(1, NSIZE) = i
k = Month(DATA_VECTOR(i, 1))

For i = 2 To NROWS
    If Month(DATA_VECTOR(i, 1)) <> k Then
        l(2, NSIZE) = i
        NSIZE = NSIZE + 1
        ReDim Preserve l(1 To 2, 1 To NSIZE)
        l(1, NSIZE) = i
        k = Month(DATA_VECTOR(i, 1))
    End If
Next i
l(2, NSIZE) = NROWS

ReDim TEMP_MATRIX(0 To NSIZE, 1 To 13)
TEMP_MATRIX(0, 1) = ("START DATE")
TEMP_MATRIX(0, 2) = ("END DATE")
TEMP_MATRIX(0, 3) = ("NO DAYS")
TEMP_MATRIX(0, 4) = ("STARTING VALUE")
TEMP_MATRIX(0, 5) = ("ENDING VALUE")
TEMP_MATRIX(0, 6) = ("DIVIDEND")
TEMP_MATRIX(0, 7) = ("MONTHLY RETURN")
TEMP_MATRIX(0, 8) = ("% MONTHLY RETURN")
TEMP_MATRIX(0, 9) = ("% ANNUAL RETURN")
TEMP_MATRIX(0, 10) = ("TOTAL DAYS")
TEMP_MATRIX(0, 11) = ("TOTAL RETURN")
TEMP_MATRIX(0, 12) = ("% TOTAL RETURN")
TEMP_MATRIX(0, 13) = ("% TOTAL ANNUAL RETURN")

For k = 1 To NSIZE
    i = l(1, k): j = l(2, k)
    TEMP_MATRIX(k, 1) = DATA_VECTOR(i, 1)
    TEMP_MATRIX(k, 2) = DATA_VECTOR(j, 1)
    TEMP_MATRIX(k, 3) = TEMP_MATRIX(k, 2) - TEMP_MATRIX(k, 1)
    TEMP_MATRIX(k, 4) = DATA_VECTOR(i, 2)
    TEMP_MATRIX(k, 5) = DATA_VECTOR(j, 2)
    TEMP_MATRIX(k, 6) = 0
    For h = i To j
        KEY_STR = DATA_VECTOR(h, 1)
        ITEM_STR = DIVIDENDS_OBJ.Item(KEY_STR)
        If Err.number = 0 Then
            TEMP_MATRIX(k, 6) = TEMP_MATRIX(k, 6) + Val(ITEM_STR)
        Else
            ITEM_STR = 0
            Err.Clear
        End If
    Next h
    If TEMP_MATRIX(k, 6) <> 0 Then
        TEMP_MATRIX(k, 7) = TEMP_MATRIX(k, 5) - TEMP_MATRIX(k, 4) '+ TEMP_MATRIX(k, 6) '--> The Data is already Adjusted for Dividends and Splits
    Else
        TEMP_MATRIX(k, 6) = ""
        TEMP_MATRIX(k, 7) = TEMP_MATRIX(k, 5) - TEMP_MATRIX(k, 4)
    End If
    TEMP_MATRIX(k, 8) = TEMP_MATRIX(k, 7) / TEMP_MATRIX(k, 4)
    TEMP_MATRIX(k, 9) = TEMP_MATRIX(k, 8) * TDAYS_PER_YEAR / TEMP_MATRIX(k, 3)
    If k > 1 Then
        TEMP_MATRIX(k, 10) = TEMP_MATRIX(k - 1, 10) + TEMP_MATRIX(k, 3)
        TEMP_MATRIX(k, 11) = TEMP_MATRIX(k - 1, 11) + TEMP_MATRIX(k, 7)
    Else
        TEMP_MATRIX(k, 10) = TEMP_MATRIX(k, 3)
        TEMP_MATRIX(k, 11) = TEMP_MATRIX(k, 7)
    End If
    TEMP_MATRIX(k, 12) = TEMP_MATRIX(k, 11) / TEMP_MATRIX(k, 4)
    TEMP_MATRIX(k, 13) = TEMP_MATRIX(k, 12) * TDAYS_PER_YEAR / TEMP_MATRIX(k, 10)
Next k

ASSET_DIVIDENDS_BENCHMARK_FUNC = TEMP_MATRIX

Exit Function
ERROR_LABEL:
ASSET_DIVIDENDS_BENCHMARK_FUNC = Err.number
End Function
