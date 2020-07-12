Attribute VB_Name = "FINAN_PORT_MOMENTS_ULCER_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.


'************************************************************************************
'************************************************************************************
'FUNCTION      : PORT_ULCER_INDEXES_FUNC
'------------------------------------------------------------------------------------
'DESCRIPTION   : The Ulcer Index is a stock market risk measure or technical
'analysis indicator devised by Peter Martin in 1987,[1] and published by him
'and Byron McCann in their 1989 book The Investors Guide to Fidelity
'Funds. It's designed as a measure of volatility, but only volatility
'in the downward direction, ie. the amount of drawdown or retracement
'occurring over a period.

'Other volatility measures like standard deviation treat up and down
'movement equally, but a trader doesn't mind upward movement, it's
'the downside that causes stress and stomach ulcers that the index's
'name suggests. (The name pre-dates the discovery, described in the
'ulcer article, that most gastric ulcers are actually caused by a bacteria.)

'The term Ulcer Index has also been used (later) by Steve Shellans,
'editor and publisher of MoniResearch Newsletter for a different
'calculation, also based on the ulcer causing potential of drawdowns.
'[2] Shellans index is not described in this article.

'REFERENCE:
'http://www.tangotools.com/ui/ui.htm
'http://www.gummy-stuff.org/Ulcer-Index.htm
'------------------------------------------------------------------------------------
'LIBRARY       : PORTFOLIO
'GROUP         : MOMENTS_ULCER
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/08/2008
'************************************************************************************
'************************************************************************************

Function PORT_ULCER_INDEXES_FUNC(ByRef DATA_RNG As Variant, _
ByRef WEIGHTS_RNG As Variant, _
Optional ByVal RISK_FREE As Double = 0.04, _
Optional ByVal COUNT_BASIS As Double = 52, _
Optional ByVal OUTPUT As Integer = 0)

'DATA_RNG:
'FIRST ROW --> HEADINGS
'FIRST COLUMN --> DATES
'SECOND COLUMN --> BENCHMARK

Dim i As Long
Dim j As Long
Dim k As Long

Dim NSIZE As Long
Dim NROWS As Long

Dim TEMP1_SUM As Double
Dim TEMP2_SUM As Double

Dim MAX1_VAL As Double
Dim MAX2_VAL As Double

Dim DATA_MATRIX As Variant
Dim WEIGHTS_VECTOR As Variant

Dim TEMP1_MATRIX As Variant
Dim TEMP2_MATRIX As Variant
Dim TEMP3_MATRIX As Variant

Dim TEMP1_VECTOR As Variant
Dim TEMP2_VECTOR As Variant
Dim TEMP3_VECTOR As Variant
Dim TEMP4_VECTOR As Variant

Dim TEMP_VECTOR As Variant
Dim TEMP_GROUP As Variant

On Error GoTo ERROR_LABEL

DATA_MATRIX = DATA_RNG

WEIGHTS_VECTOR = WEIGHTS_RNG
If UBound(WEIGHTS_VECTOR, 2) = 1 Then: WEIGHTS_VECTOR = MATRIX_TRANSPOSE_FUNC(WEIGHTS_VECTOR)

NROWS = UBound(DATA_MATRIX, 1) - 1
NSIZE = UBound(WEIGHTS_VECTOR, 2)
If NSIZE <> UBound(DATA_MATRIX, 2) - 2 Then: GoTo ERROR_LABEL
'Minus 2 because Dates & Benchmark Columns

TEMP1_VECTOR = MATRIX_REMOVE_COLUMNS_FUNC(MATRIX_REMOVE_ROWS_FUNC(DATA_MATRIX, 1, 1), 1, 1)
TEMP2_VECTOR = MATRIX_PERCENT_FUNC(TEMP1_VECTOR, 0) 'Exlude Headings and Dates

TEMP3_VECTOR = MATRIX_MEAN_FUNC(TEMP2_VECTOR)
TEMP4_VECTOR = MATRIX_STDEVP_FUNC(TEMP2_VECTOR)
'---------------------------------------------------------------------------------------
ReDim TEMP_VECTOR(1 To NROWS - 1, 1 To 1)
For i = 1 To NROWS - 1
    TEMP1_SUM = 0
    For j = 1 To NSIZE
        TEMP1_SUM = TEMP1_SUM + WEIGHTS_VECTOR(1, j) * TEMP2_VECTOR(i, j + 1)
    Next j
    TEMP_VECTOR(i, 1) = TEMP1_SUM
    TEMP2_SUM = TEMP2_SUM + TEMP_VECTOR(i, 1)
Next i
TEMP1_SUM = 0
For i = 1 To NROWS - 1
    TEMP1_SUM = TEMP1_SUM + (TEMP_VECTOR(i, 1) - (TEMP2_SUM / (NROWS - 1))) ^ 2
Next i

ReDim TEMP_VECTOR(0 To 7, 1 To NSIZE + 3)
TEMP_VECTOR(0, 1) = "SYMBOL"
TEMP_VECTOR(1, 1) = "MEAN"
TEMP_VECTOR(2, 1) = "VOLATILITY"
TEMP_VECTOR(3, 1) = "MAX DRAWDOWN"
TEMP_VECTOR(4, 1) = "AVG DRAWDOWN"
TEMP_VECTOR(5, 1) = "SHARPE RATIO"
TEMP_VECTOR(6, 1) = "ULCER INDEX"
TEMP_VECTOR(7, 1) = "MARTIN RATIO"
'---------------------------------------------------------------------------------------------

TEMP_VECTOR(0, NSIZE + 3) = "PORTFOLIO"
TEMP_VECTOR(1, NSIZE + 3) = TEMP2_SUM / (NROWS - 1) * COUNT_BASIS
TEMP_VECTOR(2, NSIZE + 3) = ((TEMP1_SUM / (NROWS - 1)) * COUNT_BASIS) ^ 0.5
TEMP_VECTOR(5, NSIZE + 3) = (TEMP_VECTOR(1, NSIZE + 3) - RISK_FREE) / TEMP_VECTOR(2, NSIZE + 3)

For j = 2 To NSIZE + 2 'Include Benchmark
    TEMP_VECTOR(0, j) = DATA_MATRIX(1, j)
    TEMP_VECTOR(1, j) = TEMP3_VECTOR(1, j - 1) * COUNT_BASIS
    TEMP_VECTOR(2, j) = TEMP4_VECTOR(1, j - 1) * Sqr(COUNT_BASIS)
    
    TEMP_VECTOR(5, j) = IIf(TEMP_VECTOR(2, j) <> 0, (TEMP_VECTOR(1, j) - RISK_FREE) / TEMP_VECTOR(2, j), 0)
Next j
'---------------------------------------------------------------------------------------------

ReDim TEMP1_MATRIX(0 To NROWS, 1 To ((NSIZE + 2) * 3))
ReDim TEMP2_MATRIX(0 To NROWS, 1 To (NSIZE + 2))
ReDim TEMP3_MATRIX(0 To NROWS, 1 To (NSIZE + 2))

TEMP1_MATRIX(0, (NSIZE + 2) * 3 - 2) = "PORTFOLIO GROWTH"
TEMP1_MATRIX(0, (NSIZE + 2) * 3 - 1) = "PREV MAX"
TEMP1_MATRIX(0, (NSIZE + 2) * 3) = "AVG DRAWDOWN"

TEMP2_MATRIX(0, (NSIZE + 2)) = "PORTFOLIO"
TEMP3_MATRIX(0, (NSIZE + 2)) = "PORTFOLIO"

k = 2

For j = 1 To ((NSIZE + 1) * 3) Step 3
    TEMP1_MATRIX(0, j) = TEMP_VECTOR(0, k) & " GROWTH"
    TEMP1_MATRIX(0, j + 1) = "PREV MAX"
    TEMP1_MATRIX(0, j + 2) = "AVG DRAWDOWN"
    
    TEMP2_MATRIX(0, k - 1) = TEMP_VECTOR(0, k)
    TEMP3_MATRIX(0, k - 1) = TEMP_VECTOR(0, k)
    
    TEMP1_SUM = 0
    MAX1_VAL = 0
    
    TEMP2_SUM = 0
    MAX2_VAL = 0
    
    For i = 1 To NROWS
        If i = 1 Then
            TEMP1_MATRIX(1, j) = 1
            TEMP1_MATRIX(1, j + 1) = 1
            TEMP1_MATRIX(1, j + 2) = 0
        
            TEMP2_MATRIX(1, k - 1) = 0
            TEMP3_MATRIX(1, k - 1) = 0
        
        Else
        
            TEMP1_MATRIX(i, j) = TEMP1_MATRIX(i - 1, j) * (1 + TEMP2_VECTOR(i - 1, k - 1))
            TEMP1_MATRIX(i, j + 1) = MAXIMUM_FUNC(MAXIMUM_FUNC(TEMP1_MATRIX(i, j), TEMP1_MATRIX(i - 1, j)), TEMP1_MATRIX(i - 1, j + 1))
            TEMP1_MATRIX(i, j + 2) = IIf(TEMP1_MATRIX(i, j + 1) <> 0, 1 - TEMP1_MATRIX(i, j) / TEMP1_MATRIX(i, j + 1), 0)
            If TEMP1_MATRIX(i, j + 1) <> 0 Then
                TEMP2_MATRIX(i, k - 1) = TEMP2_MATRIX(i - 1, k - 1) + (TEMP1_MATRIX(i, j) / TEMP1_MATRIX(i, j + 1) - 1) ^ 2
                TEMP3_MATRIX(i, k - 1) = Sqr(TEMP2_MATRIX(i, k - 1) / i)
            Else
                TEMP2_MATRIX(i, k - 1) = 0
                TEMP3_MATRIX(i, k - 1) = 0
            End If
        End If
        TEMP1_SUM = TEMP1_SUM + TEMP1_MATRIX(i, j + 2)
        MAX1_VAL = MAXIMUM_FUNC(MAX1_VAL, TEMP1_MATRIX(i, j + 2))
        If j > 3 Then 'Exclude Benchmark in Calculation
            TEMP1_MATRIX(i, (NSIZE + 2) * 3 - 2) = TEMP1_MATRIX(i, (NSIZE + 2) * 3 - 2) + TEMP1_MATRIX(i, j) * WEIGHTS_VECTOR(1, (k - 2))
            If i = 1 Then
                TEMP1_MATRIX(i, (NSIZE + 2) * 3 - 1) = TEMP1_MATRIX(i, (NSIZE + 2) * 3 - 2)
                TEMP2_MATRIX(i, NSIZE + 2) = 0
                TEMP3_MATRIX(i, NSIZE + 2) = 0
            Else
                TEMP1_MATRIX(i, (NSIZE + 2) * 3 - 1) = MAXIMUM_FUNC(MAXIMUM_FUNC(TEMP1_MATRIX(i, (NSIZE + 2) * 3 - 2), TEMP1_MATRIX(i - 1, (NSIZE + 2) * 3 - 2)), TEMP1_MATRIX(i - 1, (NSIZE + 2) * 3 - 1))
                If TEMP1_MATRIX(i, (NSIZE + 2) * 3 - 1) <> 0 Then
                    TEMP2_MATRIX(i, NSIZE + 2) = TEMP2_MATRIX(i - 1, NSIZE + 2) + (TEMP1_MATRIX(i, (NSIZE + 2) * 3 - 2) / TEMP1_MATRIX(i, (NSIZE + 2) * 3 - 1) - 1) ^ 2
                    TEMP3_MATRIX(i, NSIZE + 2) = Sqr(TEMP2_MATRIX(i, NSIZE + 2) / i)
                Else
                    TEMP2_MATRIX(i, NSIZE + 2) = 0
                    TEMP3_MATRIX(i, NSIZE + 2) = 0
                End If
            End If
            TEMP1_MATRIX(i, (NSIZE + 2) * 3) = IIf(TEMP1_MATRIX(i, (NSIZE + 2) * 3 - 1) <> 0, 1 - TEMP1_MATRIX(i, (NSIZE + 2) * 3 - 2) / TEMP1_MATRIX(i, (NSIZE + 2) * 3 - 1), 0)
            TEMP2_SUM = TEMP2_SUM + TEMP1_MATRIX(i, (NSIZE + 2) * 3)
            MAX2_VAL = MAXIMUM_FUNC(MAX2_VAL, TEMP1_MATRIX(i, (NSIZE + 2) * 3))
        End If
    Next i
    TEMP_VECTOR(3, k) = MAX1_VAL
    TEMP_VECTOR(4, k) = TEMP1_SUM / NROWS
    If TEMP2_MATRIX(NROWS, k - 1) <> 0 Then
        TEMP_VECTOR(6, k) = Sqr(TEMP2_MATRIX(NROWS, k - 1) / NROWS)
        TEMP_VECTOR(7, k) = (TEMP_VECTOR(1, k) - RISK_FREE) / TEMP_VECTOR(6, k)
    Else
        TEMP_VECTOR(6, k) = 0
        TEMP_VECTOR(7, k) = 0
    End If
    
    k = k + 1
Next j
TEMP_VECTOR(3, NSIZE + 3) = MAX2_VAL
TEMP_VECTOR(4, NSIZE + 3) = TEMP2_SUM / NROWS

If TEMP2_MATRIX(NROWS, NSIZE + 2) <> 0 Then
    TEMP_VECTOR(6, NSIZE + 3) = Sqr(TEMP2_MATRIX(NROWS, NSIZE + 2) / NROWS)
    TEMP_VECTOR(7, NSIZE + 3) = (TEMP_VECTOR(1, NSIZE + 3) - RISK_FREE) / TEMP_VECTOR(6, NSIZE + 3)
Else
    TEMP_VECTOR(6, NSIZE + 3) = 0
    TEMP_VECTOR(7, NSIZE + 3) = 0
End If
'-------------------------------------------------------------------------------------
'-------------------------------------------------------------------------------------
Select Case OUTPUT
'-------------------------------------------------------------------------------------
Case 0 'Annualized Table Summary
'-------------------------------------------------------------------------------------
    PORT_ULCER_INDEXES_FUNC = TEMP_VECTOR
'-------------------------------------------------------------------------------------
Case 1 'Growth/Drawdown Calculation
'-------------------------------------------------------------------------------------
    PORT_ULCER_INDEXES_FUNC = TEMP1_MATRIX
'-------------------------------------------------------------------------------------
Case 2 'Ulcer Indexes Calculation
'-------------------------------------------------------------------------------------
    PORT_ULCER_INDEXES_FUNC = TEMP2_MATRIX
'-------------------------------------------------------------------------------------
Case 3 'Martin Ratio Calculation
'-------------------------------------------------------------------------------------
    PORT_ULCER_INDEXES_FUNC = TEMP3_MATRIX
'-------------------------------------------------------------------------------------
Case Else
'-------------------------------------------------------------------------------------
    ReDim TEMP_GROUP(1 To 3)

    TEMP_GROUP(1) = TEMP1_MATRIX
    TEMP_GROUP(2) = TEMP2_MATRIX
    TEMP_GROUP(3) = TEMP3_MATRIX
    
    PORT_ULCER_INDEXES_FUNC = TEMP_GROUP
'-------------------------------------------------------------------------------------
'-------------------------------------------------------------------------------------
End Select
'-------------------------------------------------------------------------------------
'-------------------------------------------------------------------------------------

Exit Function
ERROR_LABEL:
PORT_ULCER_INDEXES_FUNC = Err.number
End Function
