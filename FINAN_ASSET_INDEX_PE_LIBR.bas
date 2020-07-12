Attribute VB_Name = "FINAN_ASSET_INDEX_PE_LIBR"


Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.


'************************************************************************************
'************************************************************************************
'FUNCTION      : INDEX_WEIGHTED_PE_FUNC
'DESCRIPTION   : WEIGHTED PE ANALYSIS
'LIBRARY       : FINAN_ASSET
'GROUP         : INDEX
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'************************************************************************************
'************************************************************************************

Function INDEX_WEIGHTED_PE_FUNC(Optional ByVal INDEX_STR As String = "^DJI", _
Optional ByVal REFRESH_CALLER As Variant, _
Optional ByVal SERVER_STR As String = "UNITED STATES", _
Optional ByVal OUTPUT As Integer = 1)

Dim i As Long
Dim NROWS As Long

Dim TEMP1_SUM As Double 'Number of Stocks with Market Cap=0
Dim TEMP2_SUM As Double 'Number of Stocks with Earnings > 0
Dim TEMP3_SUM As Double 'Weighted E/P

Dim TEMP4_SUM As Double 'Total Mkt Cap: Excluding companies with negative earnings
Dim TEMP5_SUM As Double 'Total Mkt Cap

Dim TEMP6_SUM As Double 'Total Earnings*: Excluding companies with negative earnings
Dim TEMP7_SUM As Double 'Total Earnings

Dim DECIMAL_STR As String

Dim TEMP_MATRIX As Variant
Dim TEMP_VECTOR As Variant
Dim DATA_MATRIX As Variant

On Error GoTo ERROR_LABEL

ReDim TEMP_VECTOR(1 To 1, 1 To 4)
TEMP_VECTOR(1, 1) = "Last Trade"
TEMP_VECTOR(1, 2) = "Earnings/Share"
TEMP_VECTOR(1, 3) = "Market Capitalization"
TEMP_VECTOR(1, 4) = "Symbol"

DATA_MATRIX = YAHOO_INDEX_QUOTES_FUNC(INDEX_STR, TEMP_VECTOR, "", False, REFRESH_CALLER, SERVER_STR)
NROWS = UBound(DATA_MATRIX, 1)

ReDim TEMP_MATRIX(0 To NROWS, 1 To 11)

TEMP1_SUM = 0 'Market Cap Sum

TEMP_MATRIX(0, 1) = "TICKER"
TEMP_MATRIX(0, 2) = "LAST TRADE (PRICE ONLY)"
TEMP_MATRIX(0, 3) = "EARNINGS/SHARE"
TEMP_MATRIX(0, 4) = "MARKET CAPITALIZATION"
TEMP_MATRIX(0, 5) = "EARNINGS PER SHARE (EXCLUDING NEGATIVE EARNINGS)"
TEMP_MATRIX(0, 6) = "MARKET CAP/PRICE"
TEMP_MATRIX(0, 7) = "P/E"
TEMP_MATRIX(0, 8) = "E/P"
TEMP_MATRIX(0, 9) = "MARKET CAP (EXCLUDING NEGATIVE EARNINGS)"
TEMP_MATRIX(0, 10) = "PERCENT OF TOTAL"
TEMP_MATRIX(0, 11) = "CUMULATIVE"

'------------------------------First Pass: Basic Calculation-------------------

DECIMAL_STR = DECIMAL_SEPARATOR_FUNC()

For i = 1 To NROWS

    TEMP_MATRIX(i, 1) = DATA_MATRIX(i, 1) & " (" & DATA_MATRIX(i, 5) & ")"
    'Name
    
    If IS_NUMERIC_FUNC(DATA_MATRIX(i, 2), DECIMAL_STR) = True Then
        TEMP_MATRIX(i, 2) = (DATA_MATRIX(i, 2)) 'Last Trade Price
    Else
        TEMP_MATRIX(i, 2) = 0
    End If
    
    If IS_NUMERIC_FUNC(DATA_MATRIX(i, 3), DECIMAL_STR) = True Then
        TEMP_MATRIX(i, 3) = (DATA_MATRIX(i, 3)) 'Earnings per share
    Else
        TEMP_MATRIX(i, 3) = 0
    End If
    
    If IS_NUMERIC_FUNC(DATA_MATRIX(i, 4), DECIMAL_STR) = True Then
        TEMP_MATRIX(i, 4) = DATA_MATRIX(i, 4) 'Market Cap.
    Else
        TEMP_MATRIX(i, 4) = 0
    End If
    
    TEMP5_SUM = TEMP5_SUM + TEMP_MATRIX(i, 4) 'Total Market Cap
    
    If TEMP_MATRIX(i, 4) = 0 Then
        TEMP1_SUM = TEMP1_SUM + 1 'Number of Stocks with Market Cap = 0
    End If
    
    TEMP_MATRIX(i, 5) = TEMP_MATRIX(i, 3) * (TEMP_MATRIX(i, 3) > 0) * -1
    'Earnings per share
    'Excluding Negative Earnings
        
    If TEMP_MATRIX(i, 5) > 0 Then
        TEMP2_SUM = TEMP2_SUM + 1 'Number of Stocks with Earnings > 0
    End If
    If TEMP_MATRIX(i, 2) <> 0 Then
        TEMP_MATRIX(i, 6) = TEMP_MATRIX(i, 4) / TEMP_MATRIX(i, 2)
        'Market Cap / Price
    Else
        TEMP_MATRIX(i, 6) = 0
    End If
    
    TEMP6_SUM = _
    TEMP6_SUM + (TEMP_MATRIX(i, 5) * TEMP_MATRIX(i, 6))
    'Total Earnings*: Excluding companies with negative earnings
    
    TEMP7_SUM = _
    TEMP7_SUM + (TEMP_MATRIX(i, 3) * TEMP_MATRIX(i, 6)) 'Total Earnings

    If TEMP_MATRIX(i, 3) <> 0 Then
        TEMP_MATRIX(i, 7) = TEMP_MATRIX(i, 2) / TEMP_MATRIX(i, 3)
        'P/E
    Else
        TEMP_MATRIX(i, 7) = 0
    End If
    
    If TEMP_MATRIX(i, 7) <> 0 Then
        TEMP_MATRIX(i, 8) = 1 / TEMP_MATRIX(i, 7)
        'E/P
    Else
        TEMP_MATRIX(i, 8) = 0
    End If
    
    TEMP3_SUM = TEMP3_SUM + (TEMP_MATRIX(i, 4) * TEMP_MATRIX(i, 8))

    If TEMP_MATRIX(i, 3) <> 0 Then
        TEMP_MATRIX(i, 9) = TEMP_MATRIX(i, 4) * (TEMP_MATRIX(i, 3) > 0) * -1
            'Market Cap (Excluding Negative Earnings)
    Else
        TEMP_MATRIX(i, 9) = 0
    End If
    
    TEMP4_SUM = TEMP4_SUM + TEMP_MATRIX(i, 9)
    'Market Cap (Excluding Negative Earnings)
Next i

'----------------------------------Second Pass: Cumulative Values calculations

If TEMP5_SUM <> 0 Then
    TEMP_MATRIX(1, 10) = TEMP_MATRIX(1, 4) / TEMP5_SUM 'Market Cap Percent of Total
Else
    TEMP_MATRIX(1, 10) = 0
End If
TEMP_MATRIX(1, 11) = TEMP_MATRIX(1, 10)

For i = 2 To NROWS
    If TEMP5_SUM <> 0 Then
        TEMP_MATRIX(i, 10) = TEMP_MATRIX(i, 4) / TEMP5_SUM
    Else
        TEMP_MATRIX(i, 10) = 0
    End If
    TEMP_MATRIX(i, 11) = TEMP_MATRIX(i - 1, 11) + TEMP_MATRIX(i, 10) 'Cumulative
Next i

ReDim TEMP_VECTOR(0 To 12, 1 To 2)

TEMP_VECTOR(0, 1) = "* EXCLUDING COMPANIES WITH NEGATIVE EARNINGS"
TEMP_VECTOR(0, 2) = ""
        
TEMP_VECTOR(1, 1) = "NUMBER OF STOCKS WITH MARKET CAP=0"
TEMP_VECTOR(1, 2) = TEMP1_SUM

TEMP_VECTOR(2, 1) = "NUMBER OF STOCKS WITH EARNINGS > 0"
TEMP_VECTOR(2, 2) = TEMP2_SUM

TEMP_VECTOR(3, 1) = "WEIGHTED E/P "
TEMP_VECTOR(3, 2) = TEMP3_SUM

TEMP_VECTOR(4, 1) = "P/E"
If TEMP5_SUM <> 0 Then
    TEMP_VECTOR(4, 2) = 1 / (TEMP3_SUM / TEMP5_SUM)
Else
    TEMP_VECTOR(4, 2) = 0
End If
TEMP_VECTOR(5, 1) = "TOTAL MKT CAP*"
TEMP_VECTOR(5, 2) = TEMP4_SUM

TEMP_VECTOR(6, 1) = "TOTAL MKT CAP"
TEMP_VECTOR(6, 2) = TEMP5_SUM
    
TEMP_VECTOR(7, 1) = "TOTAL EARNINGS*"
TEMP_VECTOR(7, 2) = TEMP6_SUM

TEMP_VECTOR(8, 1) = "TOTAL EARNINGS"
TEMP_VECTOR(8, 2) = TEMP7_SUM

TEMP_VECTOR(9, 1) = "PE_RATIOS = " & Format(TEMP_VECTOR(6, 2), "0.0") & "/" & Format(TEMP_VECTOR(8, 2), "0.0")
TEMP_VECTOR(10, 1) = "PE_RATIOS = " & Format(TEMP_VECTOR(6, 2), "0.0") & "/" & Format(TEMP_VECTOR(7, 2), "0.0")
TEMP_VECTOR(11, 1) = "PE_RATIOS = " & Format(TEMP_VECTOR(5, 2), "0.0") & "/" & Format(TEMP_VECTOR(7, 2), "0.0")
TEMP_VECTOR(12, 1) = "PE_RATIOS = " & Format(TEMP_VECTOR(5, 2), "0.0") & "/" & Format(TEMP_VECTOR(8, 2), "0.0")

If TEMP7_SUM <> 0 Then
    TEMP_VECTOR(9, 2) = TEMP5_SUM / TEMP7_SUM
    TEMP_VECTOR(12, 2) = TEMP4_SUM / TEMP7_SUM
Else
    TEMP_VECTOR(9, 2) = 0
    TEMP_VECTOR(12, 2) = 0
End If
    
If TEMP6_SUM <> 0 Then
    TEMP_VECTOR(10, 2) = TEMP5_SUM / TEMP6_SUM
    TEMP_VECTOR(11, 2) = TEMP4_SUM / TEMP6_SUM
Else
    TEMP_VECTOR(10, 2) = 0
    TEMP_VECTOR(11, 2) = 0
End If

Select Case OUTPUT
    Case 0
        INDEX_WEIGHTED_PE_FUNC = TEMP_MATRIX
    Case 1
        INDEX_WEIGHTED_PE_FUNC = TEMP_VECTOR
    Case Else
        INDEX_WEIGHTED_PE_FUNC = Array(TEMP_MATRIX, TEMP_VECTOR)
End Select

Exit Function
ERROR_LABEL:
INDEX_WEIGHTED_PE_FUNC = Err.number
End Function

