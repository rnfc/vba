Attribute VB_Name = "FINAN_ASSET_MOMENTS_CORREL_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.

Function ASSET_PRICE_VOLATILITY_CORRELATION_FUNC(ByVal TICKER_STR As Variant, _
Optional ByVal START_DATE As Date, _
Optional ByVal END_DATE As Date, _
Optional ByRef REFERENCE_DATE As Date = 0, _
Optional ByVal MA_PERIODS As Long = 20, _
Optional ByVal NBINS As Long = 20, _
Optional ByVal VERSION As Integer = 0, _
Optional ByVal METHOD As Integer = 0, _
Optional ByVal OUTPUT As Integer = 0)

'-----------------------------------------------------------------------------
'VERSION = 0 --> SD & Prices
'VERSION else --> SD & Returns
'-----------------------------------------------------------------------------
'METHOD = 0 --> Pearson
'METHOD Else --> Spearman
'-----------------------------------------------------------------------------

Dim h As Long
Dim i As Long
Dim j As Long
Dim k As Long
Dim l As Long
Dim m As Long
Dim n As Long
Dim NROWS As Long

Dim ii As Long 'MIN_VAL
Dim jj As Long 'MAX_VAL
Dim kk As Long

Dim MIN_VAL As Double
Dim MAX_VAL As Double
Dim TEMP_VAL As Double

Dim TEMP1_SUM As Double
Dim TEMP2_SUM As Double

Dim TEMP_MATRIX As Variant
Dim DATA1_VECTOR As Variant
Dim DATA2_VECTOR As Variant
Dim DATA_MATRIX As Variant

On Error GoTo ERROR_LABEL

If IsArray(TICKER_STR) = False Then
    TEMP_MATRIX = YAHOO_HISTORICAL_DATA_SERIE_FUNC(TICKER_STR, START_DATE, END_DATE, "DAILY", "DA", True, False, True)
Else
    TEMP_MATRIX = TICKER_STR
End If
NROWS = UBound(TEMP_MATRIX, 1)

ReDim DATA_MATRIX(1 To NROWS, 1 To 4)
'Date / Adj. Close /Returns /Volatility
For i = 1 To NROWS
    DATA_MATRIX(i, 1) = TEMP_MATRIX(i, 1)
    DATA_MATRIX(i, 2) = TEMP_MATRIX(i, 2)
Next i
Erase TEMP_MATRIX

h = 0
k = 3
i = k - 2
DATA_MATRIX(i, 3) = 0
DATA_MATRIX(i, 4) = 0

TEMP1_SUM = 0
For i = k - 1 To NROWS
    DATA_MATRIX(i, 3) = DATA_MATRIX(i, 2) / DATA_MATRIX(i - 1, 2) - 1
    TEMP1_SUM = TEMP1_SUM + DATA_MATRIX(i, 3)
    If i <= (MA_PERIODS + k) Then
        TEMP2_SUM = 0
        For j = i To k - 1 Step -1
            TEMP2_SUM = TEMP2_SUM + (DATA_MATRIX(j, 3) - (TEMP1_SUM / (i - 1))) ^ 2
        Next j
        TEMP2_SUM = (TEMP2_SUM / (i - 1)) ^ 0.5
    Else
        l = i - (MA_PERIODS + k - 1)
        TEMP1_SUM = TEMP1_SUM - DATA_MATRIX(l, 3)
        TEMP2_SUM = 0
        For j = i To (l + 1) Step -1
            TEMP2_SUM = TEMP2_SUM + (DATA_MATRIX(j, 3) - (TEMP1_SUM / (i - l))) ^ 2
        Next j
        TEMP2_SUM = (TEMP2_SUM / (i - l)) ^ 0.5
    End If
    DATA_MATRIX(i, 4) = TEMP2_SUM
    If DATA_MATRIX(i, 1) = REFERENCE_DATE Then: h = i
Next i


If OUTPUT > 1 Then
    ASSET_PRICE_VOLATILITY_CORRELATION_FUNC = DATA_MATRIX
    Exit Function
End If

'----------------------------------------------------------------------------------------------
If h = 0 Then 'Find Min/Max Correlation
'----------------------------------------------------------------------------------------------
    MIN_VAL = 2 ^ 52
    MAX_VAL = -2 ^ 52
    
    jj = 0: ii = 0
    For i = k - 1 To NROWS - NBINS
        h = i
        GoSub CORRELATION_LINE
        If TEMP_VAL < MIN_VAL Then
            MIN_VAL = TEMP_VAL
            jj = i
        End If
        If TEMP_VAL > MAX_VAL Then
            MAX_VAL = TEMP_VAL
            ii = i
        End If
    Next i

    If OUTPUT = 0 Then
        If jj = 0 Or ii = 0 Then: GoTo ERROR_LABEL
        ASSET_PRICE_VOLATILITY_CORRELATION_FUNC = Array(MIN_VAL, DATA_MATRIX(jj, 1), MAX_VAL, DATA_MATRIX(ii, 1))
    ElseIf OUTPUT = 1 Then
        m = jj: n = ii
        ReDim TEMP_MATRIX(0 To NBINS, 1 To 8)
        TEMP_MATRIX(0, 1) = "MIN CORREL DATE: CORREL = " & Format(MIN_VAL, "0.00%")
        TEMP_MATRIX(0, 2) = "MIN CORREL PRICE"
        TEMP_MATRIX(0, 3) = "MIN CORREL RETURN"
        TEMP_MATRIX(0, 4) = "MIN CORREL VOLATILITY"
        TEMP_MATRIX(0, 5) = "MAX CORREL DATE: CORREL = " & Format(MAX_VAL, "0.00%")
        TEMP_MATRIX(0, 6) = "MAX CORREL PRICE"
        TEMP_MATRIX(0, 7) = "MAX CORREL RETURN"
        TEMP_MATRIX(0, 8) = "MAX CORREL VOLATILITY"
        
        If m = 0 Then: GoTo 1983
        For i = 1 To NBINS
            TEMP_MATRIX(i, 1) = DATA_MATRIX(m, 1)
            TEMP_MATRIX(i, 2) = DATA_MATRIX(m, 2)
            TEMP_MATRIX(i, 3) = DATA_MATRIX(m, 3)
            TEMP_MATRIX(i, 4) = DATA_MATRIX(m, 4)
            m = m + 1
            If m > NROWS Then: Exit For
            
        Next i
1983:
        If n = 0 Then: GoTo 1984
        For i = 1 To NBINS
            TEMP_MATRIX(i, 5) = DATA_MATRIX(n, 1)
            TEMP_MATRIX(i, 6) = DATA_MATRIX(n, 2)
            TEMP_MATRIX(i, 7) = DATA_MATRIX(n, 3)
            TEMP_MATRIX(i, 8) = DATA_MATRIX(n, 4)
            n = n + 1
            If n > NROWS Then: Exit For
        Next i
1984:
        ASSET_PRICE_VOLATILITY_CORRELATION_FUNC = TEMP_MATRIX
    End If
'----------------------------------------------------------------------------------------------
Else
'----------------------------------------------------------------------------------------------
    GoSub CORRELATION_LINE

    If OUTPUT = 0 Then
        ASSET_PRICE_VOLATILITY_CORRELATION_FUNC = TEMP_VAL
    ElseIf OUTPUT = 1 Then
        m = h
        ReDim TEMP_MATRIX(0 To NBINS, 1 To 4)
        TEMP_MATRIX(0, 1) = "CORREL DATE: CORREL = " & Format(TEMP_VAL, "0.00%")
        TEMP_MATRIX(0, 2) = "CORREL PRICE"
        TEMP_MATRIX(0, 3) = "CORREL RETURN"
        TEMP_MATRIX(0, 4) = "CORREL VOLATILITY"
        For i = 1 To NBINS
            TEMP_MATRIX(i, 1) = DATA_MATRIX(m, 1)
            TEMP_MATRIX(i, 2) = DATA_MATRIX(m, 2)
            TEMP_MATRIX(i, 3) = DATA_MATRIX(m, 3)
            TEMP_MATRIX(i, 4) = DATA_MATRIX(m, 4)
            m = m + 1
            If m > NROWS Then: Exit For
        Next i
        ASSET_PRICE_VOLATILITY_CORRELATION_FUNC = TEMP_MATRIX
    End If
'----------------------------------------------------------------------------------------------
End If
'----------------------------------------------------------------------------------------------

'--------------------------------------------------------------------------------
Exit Function
'--------------------------------------------------------------------------------
CORRELATION_LINE:
'--------------------------------------------------------------------------------
    m = h
    ReDim DATA1_VECTOR(1 To NBINS, 1 To 1)
    ReDim DATA2_VECTOR(1 To NBINS, 1 To 1)
    If VERSION = 0 Then j = 2 Else j = 3
    For kk = 1 To NBINS
        DATA1_VECTOR(kk, 1) = DATA_MATRIX(m, 4)
        DATA2_VECTOR(kk, 1) = DATA_MATRIX(m, j)
        m = m + 1
        If m > NROWS Then: Exit For
    Next kk
    If METHOD = 0 Then
        TEMP_VAL = CORRELATION_FUNC(DATA1_VECTOR, DATA2_VECTOR, 0, 0)
    Else
        TEMP_VAL = CORRELATION_SPEARMAN_FUNC(DATA1_VECTOR, DATA2_VECTOR, 0, 0)
    End If
    Erase DATA1_VECTOR
    Erase DATA2_VECTOR
'--------------------------------------------------------------------------------
Return
'--------------------------------------------------------------------------------
ERROR_LABEL:
ASSET_PRICE_VOLATILITY_CORRELATION_FUNC = Err.number
End Function
