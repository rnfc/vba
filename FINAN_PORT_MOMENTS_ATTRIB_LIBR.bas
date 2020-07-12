Attribute VB_Name = "FINAN_PORT_MOMENTS_ATTRIB_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.

'************************************************************************************
'************************************************************************************
'FUNCTION      : PORT_LINK_ATTRIBUTES_FUNC
'DESCRIPTION   : Cumulative contribution and attribution effects over time
'(Frongello method).

'http://www.frongello.com/research.html
'http://www.frongello.com/support/CFADigestFeb03.pdf
'http://www.frongello.com/support/JPMWinter20022003.pdf
'http://papers.ssrn.com/sol3/papers.cfm?abstract_id=861844

'LIBRARY       : PORT_MOMENTS
'GROUP         : ATTRIBUTION
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 21/01/2009
'************************************************************************************
'************************************************************************************

Function PORT_LINK_ATTRIBUTES_FUNC(ByRef DATA_RNG As Variant, _
Optional ByRef BENCHMARK_RNG As Variant)

'DATA_RNG:
'   contribution of attribute a (for example, equities)
'   contribution of attribute b (for example, bonds)
'BENCHMARK_RNG:
'   benchmark return

Dim i As Long
Dim j As Long
Dim NROWS As Long
Dim NCOLUMNS As Long

Dim P_ARR As Variant
Dim CP_ARR As Variant
Dim CB_ARR As Variant

Dim TEMP_MATRIX As Variant
Dim DATA_MATRIX As Variant
Dim BENCHMARK_VECTOR As Variant

On Error GoTo ERROR_LABEL

DATA_MATRIX = DATA_RNG
NROWS = UBound(DATA_MATRIX, 1)
NCOLUMNS = UBound(DATA_MATRIX, 2)
If IsArray(BENCHMARK_RNG) Then 'Linking Relative Attributes (Attribution Effects)
    BENCHMARK_VECTOR = BENCHMARK_RNG
    If UBound(BENCHMARK_VECTOR, 1) = 1 Then
        BENCHMARK_VECTOR = MATRIX_TRANSPOSE_FUNC(BENCHMARK_VECTOR)
    End If
Else
    ReDim BENCHMARK_VECTOR(1 To NROWS, 1 To 1) 'Linking Absolute Attributes (Contributions)
End If
If NROWS <> UBound(BENCHMARK_VECTOR, 1) Then: GoTo ERROR_LABEL

ReDim P_ARR(1 To NROWS, 1 To 1)
ReDim CP_ARR(1 To NROWS, 1 To 1)
ReDim CB_ARR(1 To NROWS, 1 To 1)
For i = 1 To NROWS
    P_ARR(i, 1) = BENCHMARK_VECTOR(i, 1)
    For j = 1 To NCOLUMNS
        P_ARR(i, 1) = P_ARR(i, 1) + DATA_MATRIX(i, j)
    Next j
    If i = 1 Then
        CP_ARR(i, 1) = P_ARR(i, 1)
        CB_ARR(i, 1) = BENCHMARK_VECTOR(i, 1)
    Else
        CP_ARR(i, 1) = (1 + CP_ARR(i - 1, 1)) * (1 + P_ARR(i, 1)) - 1
        CB_ARR(i, 1) = (1 + CB_ARR(i - 1, 1)) * (1 + BENCHMARK_VECTOR(i, 1)) - 1
    End If
Next i
ReDim TEMP_MATRIX(1 To NROWS, 1 To NCOLUMNS)
For i = 1 To NROWS 'Periods (1,2,3,4)
    For j = 1 To NCOLUMNS
        If i = 1 Then
            TEMP_MATRIX(1, j) = DATA_MATRIX(i, j)
        Else
            TEMP_MATRIX(i, j) = TEMP_MATRIX(i - 1, j) + 0.5 * DATA_MATRIX(i, j) * (1 + CP_ARR(i - 1, 1) + 1 + CB_ARR(i - 1, 1)) + 0.5 * (P_ARR(i, 1) + BENCHMARK_VECTOR(i, 1)) * TEMP_MATRIX(i - 1, j)
        End If
    Next j
Next i
PORT_LINK_ATTRIBUTES_FUNC = TEMP_MATRIX

Exit Function
ERROR_LABEL:
PORT_LINK_ATTRIBUTES_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : RNG_PORT_SINGLE_PERIOD_PERFORMANCE_ATTRIBUTION_FUNC
'DESCRIPTION   : One-Period Performance Attribution
'LIBRARY       : PORTFOLIO
'GROUP         : TRADE_PERFORMANCE
'ID            : 002
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 21/01/2009
'************************************************************************************
'************************************************************************************
'RNG_PORT_ACTIVE_MGMT_PERFORMANCE_FUNC

Function RNG_PORT_SINGLE_PERIOD_PERFORMANCE_ATTRIBUTION_FUNC( _
ByRef DST_RNG As Excel.Range, _
ByVal NASSETS As Long, _
Optional ByVal ADD_RNG_NAME As Boolean = False)

Dim i As Long

Dim PER_POS_RNG As Excel.Range
Dim PER_RNG As Excel.Range

On Error GoTo ERROR_LABEL

RNG_PORT_SINGLE_PERIOD_PERFORMANCE_ATTRIBUTION_FUNC = False

If NASSETS < 2 Then: GoTo ERROR_LABEL

'---------------------------------------------------------------------------
'---------------------------PERFORMANCE STATISTICS--------------------------
'---------------------------------------------------------------------------

Set PER_POS_RNG = DST_RNG
PER_POS_RNG.Offset(-1, 1).value = "PORTFOLIO"
PER_POS_RNG.Offset(-1, 1).Font.Bold = True

With PER_POS_RNG
   Set PER_RNG = Range(.Offset(NASSETS, 1), .Offset(1, 11))
   If ADD_RNG_NAME = True Then: PER_RNG.name = "PER_STAT"
    
    .Offset(0, 1).value = "Weight"
    .Offset(0, 2).value = "Return"
    .Offset(0, 3).value = "Contribution"
'------------------------------------------
    .Offset(-1, 4).value = "BENCHMARK"
    .Offset(-1, 4).Font.Bold = True

    .Offset(0, 4).value = "Weight"
    .Offset(0, 5).value = "Return"
    .Offset(0, 6).value = "Contribution"

    .Offset(0, 7).value = "Overweight"
    .Offset(0, 8).value = "Performance"
'------------------------------------------
    
    .Offset(-1, 9).value = "ATTRIBUTION"
    .Offset(-1, 9).Font.Bold = True
    
    .Offset(0, 9).value = "Selection"
    .Offset(0, 10).value = "Allocation"
    .Offset(0, 11).value = "Error"

    For i = 1 To NASSETS
      With .Offset(i, 0)
         .value = "Asset " & CStr(i)
         .Font.ColorIndex = 3
      End With
      With .Offset(i, 1)
         .value = 0
         .Font.ColorIndex = 5
      End With
      With .Offset(i, 2)
         .value = 0
         .Font.ColorIndex = 5
      End With
    
      .Offset(i, 3).formula = "=" & .Offset(i, 1).Address & "*" & _
      .Offset(i, 2).Address
    
      With .Offset(i, 4)
         .value = 0
         .Font.ColorIndex = 5
      End With
      With .Offset(i, 5)
         .value = 0
         .Font.ColorIndex = 5
      End With
      
      .Offset(i, 6).formula = "=" & .Offset(i, 4).Address & "*" & _
      .Offset(i, 5).Address
    
      .Offset(i, 7).formula = "=" & .Offset(i, 1).Address & "-" & _
      .Offset(i, 4).Address
    
      .Offset(i, 8).formula = "=" & .Offset(i, 2).Address & "-" & _
      .Offset(i, 5).Address
      
      .Offset(i, 9).formula = "=" & .Offset(i, 8).Address & "*" & _
      .Offset(i, 4).Address
    
      .Offset(i, 10).formula = "=" & .Offset(i, 7).Address & "*" & _
      .Offset(i, 5).Address
    
      .Offset(i, 11).formula = "=" & .Offset(i, 8).Address & "*" & _
      .Offset(i, 7).Address
    
    Next i
    
      .Offset(NASSETS + 1, 2).formula = "=SUM(" & _
        PER_RNG.Columns(2).Address & ")"

      .Offset(NASSETS + 1, 5).formula = "=SUM(" & _
        PER_RNG.Columns(5).Address & ")"

      .Offset(NASSETS + 1, 8).formula = "=" & _
      .Offset(NASSETS + 1, 2).Address & "-" & .Offset(NASSETS + 1, 5).Address

      .Offset(NASSETS + 1, 9).formula = "=SUM(" & _
        PER_RNG.Columns(9).Address & ")"

      .Offset(NASSETS + 1, 10).formula = "=SUM(" & _
        PER_RNG.Columns(10).Address & ")"
        
        .Offset(NASSETS + 3, 10).value = "Active Mgt Effect"
        .Offset(NASSETS + 3, 10).Font.Bold = True
        .Offset(NASSETS + 3, 11).value = "=" & _
        .Offset(NASSETS + 1, 9).Address & "+" & .Offset(NASSETS + 1, 10).Address
        
        .Offset(NASSETS + 4, 10).value = "Error"
        .Offset(NASSETS + 4, 10).Font.Bold = True
        .Offset(NASSETS + 4, 11).value = "=SUM(" & _
        PER_RNG.Columns(11).Address & ")"
        
        .Offset(NASSETS + 5, 10).value = "Performance"
        .Offset(NASSETS + 5, 10).Font.Bold = True
        .Offset(NASSETS + 5, 11).value = "=" & _
        .Offset(NASSETS + 3, 11).Address & "+" & .Offset(NASSETS + 4, 11).Address
End With

RNG_PORT_SINGLE_PERIOD_PERFORMANCE_ATTRIBUTION_FUNC = True

Exit Function
ERROR_LABEL:
RNG_PORT_SINGLE_PERIOD_PERFORMANCE_ATTRIBUTION_FUNC = False
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : RNG_PORT_MULTI_PERIOD_PERFORMANCE_ATTRIBUTION_FUNC
'DESCRIPTION   : Multi-Period Performance Attribution: A routine illustrating various
'approaches to cumulate attribution effects over time.
'LIBRARY       : PORTFOLIO
'GROUP         : TRADE_PERFORMANCE
'ID            : 003
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 21/01/2009
'************************************************************************************
'************************************************************************************

Function RNG_PORT_MULTI_PERIOD_PERFORMANCE_ATTRIBUTION_FUNC( _
ByRef DST_RNG As Excel.Range, _
ByVal NASSETS As Long, _
ByVal PERIODS As Long, _
Optional ByVal FACTOR As Double = 100)

Dim i As Long
Dim j As Long
Dim m As Long

Dim ALPHAS_RNG As Excel.Range
Dim WEIGHTS_RNG As Excel.Range
Dim EXCESS_RNG As Excel.Range

Dim BENCH_TOTAL_RET As Excel.Range
Dim PORT_TOTAL_RET As Excel.Range

Dim BENCH_TOTAL_REBAL As Excel.Range
Dim PORT_TOTAL_WEIGHTS As Excel.Range

Dim CONT_FIRST_Q As Excel.Range 'BM
Dim CONT_SECOND_Q As Excel.Range 'Active Alloc
Dim CONT_THIRD_Q As Excel.Range 'Active Selec
Dim CONT_FORTH_Q As Excel.Range 'PF
Dim CONT_FIFTH_Q As Excel.Range '(wp-wb)*Rb

Dim ADD_ATT_ALLOC_RNG As Excel.Range
Dim ADD_ATT_SELEC_RNG As Excel.Range
Dim ADD_ATT_INTER_RNG As Excel.Range
Dim ADD_ATT_TOTAL_RNG As Excel.Range

Dim TRANS_CONT_ALLOC_RNG As Excel.Range
Dim TRANS_CONT_SELEC_RNG As Excel.Range
Dim TRANS_CONT_INTER_RNG As Excel.Range
Dim TRANS_CONT_TOTAL_RNG As Excel.Range

Dim TRANS_MULT_ALLOC_RNG As Excel.Range
Dim TRANS_MULT_SELEC_RNG As Excel.Range
Dim TRANS_MULT_INTER_RNG As Excel.Range
Dim TRANS_MULT_TOTAL_RNG As Excel.Range

Dim GEO_ATT_ALLOC_RNG As Excel.Range
Dim GEO_ATT_SELEC_RNG As Excel.Range
Dim GEO_ATT_INTER_RNG As Excel.Range
Dim GEO_ATT_TOTAL_RNG As Excel.Range

On Error GoTo ERROR_LABEL

RNG_PORT_MULTI_PERIOD_PERFORMANCE_ATTRIBUTION_FUNC = False

m = 7

'-------------------------1 PASS: SEGMENT ALPHAS---------------------------

Set ALPHAS_RNG = Range(DST_RNG.Offset(2, 1), _
DST_RNG.Offset(NASSETS + 1, PERIODS))

ALPHAS_RNG.Cells(-2, 0).value = "Segment Returns"
ALPHAS_RNG.Cells(-2, 0).Font.Bold = True

ALPHAS_RNG.Cells(0, 0).value = "PERIODS"
For i = 1 To PERIODS
    ALPHAS_RNG.Cells(0, i).value = i
    ALPHAS_RNG.Cells(0, i).Font.ColorIndex = 3
Next i

For i = 1 To NASSETS
    ALPHAS_RNG.Cells(i, 0).value = "ASSETS: " & i
    ALPHAS_RNG.Cells(i, 0).Font.ColorIndex = 3
Next i

ALPHAS_RNG.value = 0
ALPHAS_RNG.Font.ColorIndex = 5

'-------------------------2 PASS: SEGMENT WEIGHTS---------------------------

Set DST_RNG = DST_RNG.Offset(NASSETS + m)

Set WEIGHTS_RNG = Range(DST_RNG.Offset(2, 1), _
DST_RNG.Offset(NASSETS + 1, PERIODS))

WEIGHTS_RNG.Cells(-2, 0).value = "Segment Weights"
WEIGHTS_RNG.Cells(-2, 0).Font.Bold = True

WEIGHTS_RNG.Cells(0, 0).value = "PERIODS"
For i = 1 To PERIODS
    WEIGHTS_RNG.Cells(0, i).formula = "=" & ALPHAS_RNG.Cells(0, i).Address
Next i

For i = 1 To NASSETS
    WEIGHTS_RNG.Cells(i, 0).formula = "=" & ALPHAS_RNG.Cells(i, 0).Address
Next i

WEIGHTS_RNG.value = 0
WEIGHTS_RNG.Font.ColorIndex = 5

'------------------------3 PASS: BENCHMARK RETURNS-----------------------------

Set DST_RNG = DST_RNG.Offset(NASSETS + m)

Set BENCH_TOTAL_RET = Range(DST_RNG.Offset(2, 1), _
DST_RNG.Offset(NASSETS + 1, PERIODS))

BENCH_TOTAL_RET.Cells(-2, 0).value = "Benchmark Returns"
BENCH_TOTAL_RET.Cells(-2, 0).Font.Bold = True

BENCH_TOTAL_RET.Cells(0, 0).value = "PERIODS"
For i = 1 To PERIODS
    BENCH_TOTAL_RET.Cells(0, i).formula = "=" & ALPHAS_RNG.Cells(0, i).Address
Next i

For i = 1 To NASSETS
    BENCH_TOTAL_RET.Cells(i, 0).formula = "=" & ALPHAS_RNG.Cells(i, 0).Address
Next i

BENCH_TOTAL_RET.value = 0
BENCH_TOTAL_RET.Font.ColorIndex = 5
BENCH_TOTAL_RET.Cells(NASSETS + 1, 0).value = "Total"

'---------------------------4 PASS: BENCHMARK REBAL-----------------------------

Set DST_RNG = DST_RNG.Offset(NASSETS + m)

Set BENCH_TOTAL_REBAL = Range(DST_RNG.Offset(2, 1), _
DST_RNG.Offset(NASSETS + 1, PERIODS))

BENCH_TOTAL_REBAL.Cells(-2, 0).value = "Benchmark Rebalance"
BENCH_TOTAL_REBAL.Cells(-2, 0).Font.Bold = True

BENCH_TOTAL_REBAL.Cells(0, 0).value = "PERIODS"
For i = 1 To PERIODS
    BENCH_TOTAL_REBAL.Cells(0, i).formula = "=" & ALPHAS_RNG.Cells(0, i).Address
Next i

For i = 1 To NASSETS
    BENCH_TOTAL_REBAL.Cells(i, 0).formula = "=" & ALPHAS_RNG.Cells(i, 0).Address
Next i

For j = 1 To PERIODS
    For i = 1 To NASSETS
        If j = 1 Then
            BENCH_TOTAL_REBAL.Cells(i, j).value = 0
            BENCH_TOTAL_REBAL.Cells(i, j).Font.ColorIndex = 5
        Else
            BENCH_TOTAL_REBAL.Cells(i, j).formula = "=" & _
            BENCH_TOTAL_REBAL.Cells(i, j - 1).Address & "*(1 + " & _
            BENCH_TOTAL_RET.Cells(i, j - 1).Address & ") / (1 + " & _
            BENCH_TOTAL_RET.Cells(NASSETS + 1, j - 1).Address & ")"
        End If
    Next i
            
    BENCH_TOTAL_RET.Cells(NASSETS + 1, j).formula = "=SUMPRODUCT(" & BENCH_TOTAL_RET.Columns(j).Address & "," & BENCH_TOTAL_REBAL.Columns(j).Address & ")"
    BENCH_TOTAL_REBAL.Cells(NASSETS + 1, j).formula = "=SUM(" & BENCH_TOTAL_REBAL.Columns(j).Address & ")"
Next j

BENCH_TOTAL_RET.Cells(NASSETS + 1, PERIODS + 1).FormulaArray = "=" & "PRODUCT(1+" & BENCH_TOTAL_RET.Rows(NASSETS + 1).Address & ")-1"
BENCH_TOTAL_REBAL.Cells(NASSETS + 1, 0).value = "Total"

'---------------------------5 PASS: PORTFOLIO Returns-----------------------------

Set DST_RNG = DST_RNG.Offset(NASSETS + m)

Set PORT_TOTAL_RET = Range(DST_RNG.Offset(2, 1), _
DST_RNG.Offset(NASSETS + 1, PERIODS))

PORT_TOTAL_RET.Cells(-2, 0).value = "Portfolio Returns"
PORT_TOTAL_RET.Cells(-2, 0).Font.Bold = True

PORT_TOTAL_RET.Cells(0, 0).value = "PERIODS"
For i = 1 To PERIODS
    PORT_TOTAL_RET.Cells(0, i).formula = "=" & ALPHAS_RNG.Cells(0, i).Address
Next i

For i = 1 To NASSETS
    PORT_TOTAL_RET.Cells(i, 0).formula = "=" & ALPHAS_RNG.Cells(i, 0).Address
Next i

For j = 1 To PERIODS
    For i = 1 To NASSETS
        PORT_TOTAL_RET.Cells(i, j).formula = "=" & BENCH_TOTAL_RET.Cells(i, j).Address & "+" & ALPHAS_RNG.Cells(i, j).Address
    Next i
Next j

PORT_TOTAL_RET.Cells(NASSETS + 1, PERIODS + 1).FormulaArray = "=" & "PRODUCT(1+" & PORT_TOTAL_RET.Rows(NASSETS + 1).Address & ")-1"
PORT_TOTAL_RET.Cells(NASSETS + 1, 0).value = "Total"

'---------------------------6 PASS: PORTFOLIO Weights-----------------------------

Set DST_RNG = DST_RNG.Offset(NASSETS + m)

Set PORT_TOTAL_WEIGHTS = Range(DST_RNG.Offset(2, 1), _
DST_RNG.Offset(NASSETS + 1, PERIODS))

PORT_TOTAL_WEIGHTS.Cells(-2, 0).value = "Portfolio Weights"
PORT_TOTAL_WEIGHTS.Cells(-2, 0).Font.Bold = True

PORT_TOTAL_WEIGHTS.Cells(0, 0).value = "PERIODS"
For i = 1 To PERIODS
    PORT_TOTAL_WEIGHTS.Cells(0, i).formula = "=" & ALPHAS_RNG.Cells(0, i).Address
Next i

For i = 1 To NASSETS
    PORT_TOTAL_WEIGHTS.Cells(i, 0).formula = "=" & ALPHAS_RNG.Cells(i, 0).Address
Next i

For j = 1 To PERIODS
    For i = 1 To NASSETS
            PORT_TOTAL_WEIGHTS.Cells(i, j).formula = "=" & _
            BENCH_TOTAL_REBAL.Cells(i, j).Address & "+" & _
            WEIGHTS_RNG.Cells(i, j).Address
    Next i
            PORT_TOTAL_WEIGHTS.Cells(NASSETS + 1, j).formula = "=SUM(" & _
            PORT_TOTAL_WEIGHTS.Columns(j).Address & ")"

            PORT_TOTAL_RET.Cells(NASSETS + 1, j).formula = "=SUMPRODUCT(" & _
            PORT_TOTAL_RET.Columns(j).Address & "," & _
            PORT_TOTAL_WEIGHTS.Columns(j).Address & ")"

Next j

PORT_TOTAL_WEIGHTS.Cells(NASSETS + 1, 0).value = "Total"

'---------------------------7 PASS: FIRST QUARTILE-----------------------------

Set DST_RNG = DST_RNG.Offset(NASSETS + m)

Set CONT_FIRST_Q = Range(DST_RNG.Offset(2, 1), _
DST_RNG.Offset(NASSETS + 1, PERIODS))

CONT_FIRST_Q.Cells(-2, 0).value = "Contributions Q1: BM"
CONT_FIRST_Q.Cells(-2, 0).Font.Bold = True

CONT_FIRST_Q.Cells(0, 0).value = "PERIODS"
For i = 1 To PERIODS
    CONT_FIRST_Q.Cells(0, i).formula = "=" & ALPHAS_RNG.Cells(0, i).Address
Next i

For i = 1 To NASSETS
    CONT_FIRST_Q.Cells(i, 0).formula = "=" & ALPHAS_RNG.Cells(i, 0).Address
Next i

For j = 1 To PERIODS
    For i = 1 To NASSETS
            CONT_FIRST_Q.Cells(i, j).formula = "=" & _
            BENCH_TOTAL_REBAL.Cells(i, j).Address & "*" & _
            BENCH_TOTAL_RET.Cells(i, j).Address
    Next i
            CONT_FIRST_Q.Cells(NASSETS + 1, j).formula = "=SUM(" & _
            CONT_FIRST_Q.Columns(j).Address & ")"

Next j

CONT_FIRST_Q.Cells(NASSETS + 1, PERIODS + 1).FormulaArray = "=" & _
"PRODUCT(1+" & CONT_FIRST_Q.Rows(NASSETS + 1).Address & ")-1"

CONT_FIRST_Q.Cells(NASSETS + 1, 0).value = "Total"

'---------------------------8 PASS: SECOND QUARTILE-----------------------------

Set DST_RNG = DST_RNG.Offset(NASSETS + m)

Set CONT_SECOND_Q = Range(DST_RNG.Offset(2, 1), _
DST_RNG.Offset(NASSETS + 1, PERIODS))

CONT_SECOND_Q.Cells(-2, 0).value = "Contributions Q2: Active Alloc"
CONT_SECOND_Q.Cells(-2, 0).Font.Bold = True

CONT_SECOND_Q.Cells(0, 0).value = "PERIODS"
For i = 1 To PERIODS
    CONT_SECOND_Q.Cells(0, i).formula = "=" & ALPHAS_RNG.Cells(0, i).Address
Next i

For i = 1 To NASSETS
    CONT_SECOND_Q.Cells(i, 0).formula = "=" & ALPHAS_RNG.Cells(i, 0).Address
Next i

For j = 1 To PERIODS
    For i = 1 To NASSETS
            CONT_SECOND_Q.Cells(i, j).formula = "=" & _
            PORT_TOTAL_WEIGHTS.Cells(i, j).Address & "*" & _
            BENCH_TOTAL_RET.Cells(i, j).Address
    Next i
            CONT_SECOND_Q.Cells(NASSETS + 1, j).formula = "=SUM(" & _
            CONT_SECOND_Q.Columns(j).Address & ")"

Next j

CONT_SECOND_Q.Cells(NASSETS + 1, PERIODS + 1).FormulaArray = "=" & _
"PRODUCT(1+" & CONT_SECOND_Q.Rows(NASSETS + 1).Address & ")-1"

CONT_SECOND_Q.Cells(NASSETS + 1, 0).value = "Total"


'---------------------------9 PASS: THIRD QUARTILE-----------------------------

Set DST_RNG = DST_RNG.Offset(NASSETS + m)

Set CONT_THIRD_Q = Range(DST_RNG.Offset(2, 1), _
DST_RNG.Offset(NASSETS + 1, PERIODS))

CONT_THIRD_Q.Cells(-2, 0).value = "Contributions Q3: Active Selec"
CONT_THIRD_Q.Cells(-2, 0).Font.Bold = True

CONT_THIRD_Q.Cells(0, 0).value = "PERIODS"
For i = 1 To PERIODS
    CONT_THIRD_Q.Cells(0, i).formula = "=" & ALPHAS_RNG.Cells(0, i).Address
Next i

For i = 1 To NASSETS
    CONT_THIRD_Q.Cells(i, 0).formula = "=" & ALPHAS_RNG.Cells(i, 0).Address
Next i

For j = 1 To PERIODS
    For i = 1 To NASSETS
            CONT_THIRD_Q.Cells(i, j).formula = "=" & _
            PORT_TOTAL_RET.Cells(i, j).Address & "*" & _
            BENCH_TOTAL_REBAL.Cells(i, j).Address
    Next i
            CONT_THIRD_Q.Cells(NASSETS + 1, j).formula = "=SUM(" & _
            CONT_THIRD_Q.Columns(j).Address & ")"

Next j

CONT_THIRD_Q.Cells(NASSETS + 1, PERIODS + 1).FormulaArray = "=" & _
"PRODUCT(1+" & CONT_THIRD_Q.Rows(NASSETS + 1).Address & ")-1"

CONT_THIRD_Q.Cells(NASSETS + 1, 0).value = "Total"

'---------------------------10 PASS: FORTH QUARTILE-----------------------------

Set DST_RNG = DST_RNG.Offset(NASSETS + m)

Set CONT_FORTH_Q = Range(DST_RNG.Offset(2, 1), _
DST_RNG.Offset(NASSETS + 1, PERIODS))

CONT_FORTH_Q.Cells(-2, 0).value = "Contributions Q4: PF"
CONT_FORTH_Q.Cells(-2, 0).Font.Bold = True

CONT_FORTH_Q.Cells(0, 0).value = "PERIODS"
For i = 1 To PERIODS
    CONT_FORTH_Q.Cells(0, i).formula = "=" & ALPHAS_RNG.Cells(0, i).Address
Next i

For i = 1 To NASSETS
    CONT_FORTH_Q.Cells(i, 0).formula = "=" & ALPHAS_RNG.Cells(i, 0).Address
Next i

For j = 1 To PERIODS
    For i = 1 To NASSETS
            CONT_FORTH_Q.Cells(i, j).formula = "=" & _
            PORT_TOTAL_WEIGHTS.Cells(i, j).Address & "*" & _
            PORT_TOTAL_RET.Cells(i, j).Address
    Next i
            CONT_FORTH_Q.Cells(NASSETS + 1, j).formula = "=SUM(" & _
            CONT_FORTH_Q.Columns(j).Address & ")"

Next j

CONT_FORTH_Q.Cells(NASSETS + 1, PERIODS + 1).FormulaArray = "=" & _
"PRODUCT(1+" & CONT_FORTH_Q.Rows(NASSETS + 1).Address & ")-1"

CONT_FORTH_Q.Cells(NASSETS + 1, 0).value = "Total"

'---------------------------11 PASS: FIFTH QUARTILE-----------------------------

Set DST_RNG = DST_RNG.Offset(NASSETS + m)

Set CONT_FIFTH_Q = Range(DST_RNG.Offset(2, 1), _
DST_RNG.Offset(NASSETS + 1, PERIODS))

CONT_FIFTH_Q.Cells(-2, 0).value = "Contributions Q5: (wp-wb)*Rb"
CONT_FIFTH_Q.Cells(-2, 0).Font.Bold = True

CONT_FIFTH_Q.Cells(0, 0).value = "PERIODS"
For i = 1 To PERIODS
    CONT_FIFTH_Q.Cells(0, i).formula = "=" & ALPHAS_RNG.Cells(0, i).Address
Next i

For i = 1 To NASSETS
    CONT_FIFTH_Q.Cells(i, 0).formula = "=" & ALPHAS_RNG.Cells(i, 0).Address
Next i

For j = 1 To PERIODS
    For i = 1 To NASSETS
            CONT_FIFTH_Q.Cells(i, j).formula = "=(" & _
            PORT_TOTAL_WEIGHTS.Cells(i, j).Address & "-" & _
            BENCH_TOTAL_REBAL.Cells(i, j).Address & ")*" & _
            BENCH_TOTAL_RET.Cells(NASSETS + 1, j).Address
    Next i
            CONT_FIFTH_Q.Cells(NASSETS + 1, j).formula = "=SUM(" & _
            CONT_FIFTH_Q.Columns(j).Address & ")"

Next j

CONT_FIFTH_Q.Cells(NASSETS + 1, PERIODS + 1).FormulaArray = "=" & _
"PRODUCT(1+" & CONT_FIFTH_Q.Rows(NASSETS + 1).Address & ")-1"

CONT_FIFTH_Q.Cells(NASSETS + 1, 0).value = "Total"

'---------------------------12 PASS: EXCESS RETURNS-----------------------------

Set DST_RNG = DST_RNG.Offset(NASSETS + m)

Set EXCESS_RNG = Range(DST_RNG.Offset(2, 1), _
DST_RNG.Offset(5 + 1, PERIODS))

EXCESS_RNG.Cells(-2, 0).value = "Excess Returns"
EXCESS_RNG.Cells(-2, 0).Font.Bold = True

EXCESS_RNG.Cells(0, 0).value = "PERIODS"
For i = 1 To PERIODS
    EXCESS_RNG.Cells(0, i).formula = "=" & ALPHAS_RNG.Cells(0, i).Address
Next i

EXCESS_RNG.Cells(1, 0).value = "Additive"
EXCESS_RNG.Cells(2, 0).value = "Geometric"
EXCESS_RNG.Cells(3, 0).value = "Log"
EXCESS_RNG.Cells(4, 0).value = "k"
EXCESS_RNG.Cells(5, 0).value = "k*=1/{1+r(b)}"

For j = 1 To PERIODS + 1
    EXCESS_RNG.Cells(1, j).formula = "=" & _
        PORT_TOTAL_RET.Cells(NASSETS + 1, j).Address & "-" & _
        BENCH_TOTAL_RET.Cells(NASSETS + 1, j).Address
    EXCESS_RNG.Cells(2, j).formula = "=(1+" & _
        PORT_TOTAL_RET.Cells(NASSETS + 1, j).Address & ")/(1+" & _
        BENCH_TOTAL_RET.Cells(NASSETS + 1, j).Address & ")-1"
    EXCESS_RNG.Cells(3, j).formula = "=LN(1+" & _
        PORT_TOTAL_RET.Cells(NASSETS + 1, j).Address & ")-LN(1+" & _
        BENCH_TOTAL_RET.Cells(NASSETS + 1, j).Address & ")"
    EXCESS_RNG.Cells(4, j).formula = "=" & EXCESS_RNG.Cells(3, j).Address & _
        "/" & EXCESS_RNG.Cells(1, j).Address
    EXCESS_RNG.Cells(5, j).formula = "=" & EXCESS_RNG.Cells(2, j).Address & _
        "/" & EXCESS_RNG.Cells(1, j).Address
Next j


'---------------------------13 PASS: Additive Attribution Effects: Alloc------------------

Set DST_RNG = DST_RNG.Offset(5 + m)

Set ADD_ATT_ALLOC_RNG = Range(DST_RNG.Offset(2, 1), _
DST_RNG.Offset(NASSETS + 1, PERIODS))

ADD_ATT_ALLOC_RNG.Cells(-2, 0).value = "Additive Attribution Effects: Alloc"
ADD_ATT_ALLOC_RNG.Cells(-2, 0).Font.Bold = True

ADD_ATT_ALLOC_RNG.Cells(0, 0).value = "PERIODS"
For i = 1 To PERIODS
    ADD_ATT_ALLOC_RNG.Cells(0, i).formula = "=" & ALPHAS_RNG.Cells(0, i).Address
Next i

For i = 1 To NASSETS
    ADD_ATT_ALLOC_RNG.Cells(i, 0).formula = "=" & ALPHAS_RNG.Cells(i, 0).Address
Next i

For j = 1 To PERIODS
        For i = 1 To NASSETS
            ADD_ATT_ALLOC_RNG.Cells(i, j).formula = "=" & _
            CONT_SECOND_Q.Cells(i, j).Address & "-" & _
            CONT_FIRST_Q.Cells(i, j).Address & "-" & _
            CONT_FIFTH_Q.Cells(i, j).Address
        Next i
            ADD_ATT_ALLOC_RNG.Cells(NASSETS + 1, j).formula = "=SUM(" & _
            ADD_ATT_ALLOC_RNG.Columns(j).Address & ")"
Next j

            ADD_ATT_ALLOC_RNG.Cells(NASSETS + 1, PERIODS + 1).formula = "=SUM(" & _
            ADD_ATT_ALLOC_RNG.Columns(PERIODS + 1).Address & ")"

ADD_ATT_ALLOC_RNG.Cells(NASSETS + 1, PERIODS + 2).formula = "=" & _
            CONT_SECOND_Q.Cells(NASSETS + 1, PERIODS + 1).Address & "-" & _
            CONT_FIRST_Q.Cells(NASSETS + 1, PERIODS + 1).Address & "-" & _
            CONT_FIFTH_Q.Cells(NASSETS + 1, PERIODS + 1).Address

ADD_ATT_ALLOC_RNG.Cells(NASSETS + 1, 0).value = "Total"

'---------------------------14 PASS: Additive Attribution Effects: Selec------------------

Set DST_RNG = DST_RNG.Offset(NASSETS + m)

Set ADD_ATT_SELEC_RNG = Range(DST_RNG.Offset(2, 1), _
DST_RNG.Offset(NASSETS + 1, PERIODS))

ADD_ATT_SELEC_RNG.Cells(-2, 0).value = "Additive Attribution Effects: Selec"
ADD_ATT_SELEC_RNG.Cells(-2, 0).Font.Bold = True

ADD_ATT_SELEC_RNG.Cells(0, 0).value = "PERIODS"
For i = 1 To PERIODS
    ADD_ATT_SELEC_RNG.Cells(0, i).formula = "=" & ALPHAS_RNG.Cells(0, i).Address
Next i

For i = 1 To NASSETS
    ADD_ATT_SELEC_RNG.Cells(i, 0).formula = "=" & ALPHAS_RNG.Cells(i, 0).Address
Next i

For j = 1 To PERIODS
        For i = 1 To NASSETS
            ADD_ATT_SELEC_RNG.Cells(i, j).formula = "=" & _
            CONT_THIRD_Q.Cells(i, j).Address & "-" & _
            CONT_FIRST_Q.Cells(i, j).Address
        Next i
            ADD_ATT_SELEC_RNG.Cells(NASSETS + 1, j).formula = "=SUM(" & _
            ADD_ATT_SELEC_RNG.Columns(j).Address & ")"
Next j

            ADD_ATT_SELEC_RNG.Cells(NASSETS + 1, PERIODS + 1).formula = "=SUM(" & _
            ADD_ATT_SELEC_RNG.Columns(PERIODS + 1).Address & ")"

ADD_ATT_SELEC_RNG.Cells(NASSETS + 1, PERIODS + 2).formula = "=" & _
            CONT_THIRD_Q.Cells(NASSETS + 1, PERIODS + 1).Address & "-" & _
            CONT_FIRST_Q.Cells(NASSETS + 1, PERIODS + 1).Address

ADD_ATT_SELEC_RNG.Cells(NASSETS + 1, 0).value = "Total"


'---------------------------15 PASS: Additive Attribution Effects: Inter------------------

Set DST_RNG = DST_RNG.Offset(NASSETS + m)

Set ADD_ATT_INTER_RNG = Range(DST_RNG.Offset(2, 1), _
DST_RNG.Offset(NASSETS + 1, PERIODS))

ADD_ATT_INTER_RNG.Cells(-2, 0).value = "Additive Attribution Effects: Inter"
ADD_ATT_INTER_RNG.Cells(-2, 0).Font.Bold = True

ADD_ATT_INTER_RNG.Cells(0, 0).value = "PERIODS"
For i = 1 To PERIODS
    ADD_ATT_INTER_RNG.Cells(0, i).formula = "=" & ALPHAS_RNG.Cells(0, i).Address
Next i

For i = 1 To NASSETS
    ADD_ATT_INTER_RNG.Cells(i, 0).formula = "=" & ALPHAS_RNG.Cells(i, 0).Address
Next i

For j = 1 To PERIODS
        For i = 1 To NASSETS
            ADD_ATT_INTER_RNG.Cells(i, j).formula = "=" & _
            CONT_FORTH_Q.Cells(i, j).Address & "-" & _
            CONT_THIRD_Q.Cells(i, j).Address & "-" & _
            CONT_SECOND_Q.Cells(i, j).Address & "+" & _
            CONT_FIRST_Q.Cells(i, j).Address
        Next i
            ADD_ATT_INTER_RNG.Cells(NASSETS + 1, j).formula = "=SUM(" & _
            ADD_ATT_INTER_RNG.Columns(j).Address & ")"
Next j

            ADD_ATT_INTER_RNG.Cells(NASSETS + 1, PERIODS + 1).formula = "=SUM(" & _
            ADD_ATT_INTER_RNG.Columns(PERIODS + 1).Address & ")"

ADD_ATT_INTER_RNG.Cells(NASSETS + 1, PERIODS + 2).formula = "=" & _
            CONT_FORTH_Q.Cells(NASSETS + 1, PERIODS + 1).Address & "-" & _
            CONT_THIRD_Q.Cells(NASSETS + 1, PERIODS + 1).Address & "-" & _
            CONT_SECOND_Q.Cells(NASSETS + 1, PERIODS + 1).Address & "+" & _
            CONT_FIRST_Q.Cells(NASSETS + 1, PERIODS + 1).Address

ADD_ATT_INTER_RNG.Cells(NASSETS + 1, 0).value = "Total"


'---------------------------16 PASS: Additive Attribution Effects: Total------------------

Set DST_RNG = DST_RNG.Offset(NASSETS + m)

Set ADD_ATT_TOTAL_RNG = Range(DST_RNG.Offset(2, 1), _
DST_RNG.Offset(NASSETS + 1, PERIODS))

ADD_ATT_TOTAL_RNG.Cells(-2, 0).value = "Additive Attribution Effects: Total"
ADD_ATT_TOTAL_RNG.Cells(-2, 0).Font.Bold = True

ADD_ATT_TOTAL_RNG.Cells(0, 0).value = "PERIODS"
For i = 1 To PERIODS
    ADD_ATT_TOTAL_RNG.Cells(0, i).formula = "=" & ALPHAS_RNG.Cells(0, i).Address
Next i

For i = 1 To NASSETS
    ADD_ATT_TOTAL_RNG.Cells(i, 0).formula = "=" & ALPHAS_RNG.Cells(i, 0).Address
Next i

For j = 1 To PERIODS
        For i = 1 To NASSETS
            ADD_ATT_TOTAL_RNG.Cells(i, j).formula = "=" & _
            CONT_FORTH_Q.Cells(i, j).Address & "-" & _
            CONT_FIRST_Q.Cells(i, j).Address
        Next i
            ADD_ATT_TOTAL_RNG.Cells(NASSETS + 1, j).formula = "=SUM(" & _
            ADD_ATT_TOTAL_RNG.Columns(j).Address & ")"
Next j


        For i = 1 To NASSETS
            ADD_ATT_TOTAL_RNG.Cells(i, PERIODS + 1).formula = "=" & _
                ADD_ATT_ALLOC_RNG.Cells(i, PERIODS + 1).Address & "+" & _
                ADD_ATT_SELEC_RNG.Cells(i, PERIODS + 1).Address & "+" & _
                ADD_ATT_INTER_RNG.Cells(i, PERIODS + 1).Address
        Next i

            ADD_ATT_TOTAL_RNG.Cells(NASSETS + 1, PERIODS + 1).formula = "=SUM(" & _
            ADD_ATT_TOTAL_RNG.Columns(PERIODS + 1).Address & ")"


ADD_ATT_TOTAL_RNG.Cells(NASSETS + 1, PERIODS + 2).formula = "=" & _
            CONT_FORTH_Q.Cells(NASSETS + 1, PERIODS + 1).Address & "-" & _
            CONT_FIRST_Q.Cells(NASSETS + 1, PERIODS + 1).Address

ADD_ATT_TOTAL_RNG.Cells(NASSETS + 1, 0).value = "Total"



'---------------------------17 PASS: Transformed Add Effects I
'(Continuously Compounding): Alloc------------------

Set DST_RNG = DST_RNG.Offset(NASSETS + m)

Set TRANS_CONT_ALLOC_RNG = Range(DST_RNG.Offset(2, 1), _
DST_RNG.Offset(NASSETS + 1, PERIODS))

TRANS_CONT_ALLOC_RNG.Cells(-2, 0).value = "Transformed Add Effects I " & _
"(Continuously Compounding): Alloc"

TRANS_CONT_ALLOC_RNG.Cells(-2, 0).Font.Bold = True

TRANS_CONT_ALLOC_RNG.Cells(0, 0).value = "PERIODS"
For i = 1 To PERIODS
    TRANS_CONT_ALLOC_RNG.Cells(0, i).formula = "=" & ALPHAS_RNG.Cells(0, i).Address
Next i

For i = 1 To NASSETS
    TRANS_CONT_ALLOC_RNG.Cells(i, 0).formula = "=" & ALPHAS_RNG.Cells(i, 0).Address
Next i

For j = 1 To PERIODS
        For i = 1 To NASSETS
            TRANS_CONT_ALLOC_RNG.Cells(i, j).formula = "=" & _
            EXCESS_RNG.Cells(4, j).Address & "*" & _
            ADD_ATT_ALLOC_RNG.Cells(i, j).Address
        Next i
            TRANS_CONT_ALLOC_RNG.Cells(NASSETS + 1, j).formula = "=SUM(" & _
            TRANS_CONT_ALLOC_RNG.Columns(j).Address & ")"
Next j

        For i = 1 To NASSETS
            TRANS_CONT_ALLOC_RNG.Cells(i, PERIODS + 1).formula = "=SUM(" & _
            TRANS_CONT_ALLOC_RNG.Rows(i).Address & ")"
        Next i

            TRANS_CONT_ALLOC_RNG.Cells(NASSETS + 1, PERIODS + 1).formula = "=SUM(" & _
            TRANS_CONT_ALLOC_RNG.Columns(PERIODS + 1).Address & ")"

TRANS_CONT_ALLOC_RNG.Cells(NASSETS + 1, 0).value = "Total"

        For i = 1 To NASSETS
            ADD_ATT_ALLOC_RNG.Cells(i, PERIODS + 1).formula = "=sum(" & _
                TRANS_CONT_ALLOC_RNG.Rows(i).Address & ")/" & _
                EXCESS_RNG.Cells(4, PERIODS + 1).Address
        Next i

'---------------------------18 PASS: Transformed Add Effects I
'(Continuously Compounding): Selec------------------

Set DST_RNG = DST_RNG.Offset(NASSETS + m)

Set TRANS_CONT_SELEC_RNG = Range(DST_RNG.Offset(2, 1), _
DST_RNG.Offset(NASSETS + 1, PERIODS))

TRANS_CONT_SELEC_RNG.Cells(-2, 0).value = "Transformed Add Effects I " & _
"(Continuously Compounding): Selec"
TRANS_CONT_SELEC_RNG.Cells(-2, 0).Font.Bold = True

TRANS_CONT_SELEC_RNG.Cells(0, 0).value = "PERIODS"
For i = 1 To PERIODS
    TRANS_CONT_SELEC_RNG.Cells(0, i).formula = "=" & ALPHAS_RNG.Cells(0, i).Address
Next i

For i = 1 To NASSETS
    TRANS_CONT_SELEC_RNG.Cells(i, 0).formula = "=" & ALPHAS_RNG.Cells(i, 0).Address
Next i

For j = 1 To PERIODS
        For i = 1 To NASSETS
            TRANS_CONT_SELEC_RNG.Cells(i, j).formula = "=" & _
            EXCESS_RNG.Cells(4, j).Address & "*" & _
            ADD_ATT_SELEC_RNG.Cells(i, j).Address
        Next i
            TRANS_CONT_SELEC_RNG.Cells(NASSETS + 1, j).formula = "=SUM(" & _
            TRANS_CONT_SELEC_RNG.Columns(j).Address & ")"
Next j

        For i = 1 To NASSETS
            TRANS_CONT_SELEC_RNG.Cells(i, PERIODS + 1).formula = "=SUM(" & _
            TRANS_CONT_SELEC_RNG.Rows(i).Address & ")"
        Next i

            TRANS_CONT_SELEC_RNG.Cells(NASSETS + 1, PERIODS + 1).formula = "=SUM(" & _
            TRANS_CONT_SELEC_RNG.Columns(PERIODS + 1).Address & ")"

TRANS_CONT_SELEC_RNG.Cells(NASSETS + 1, 0).value = "Total"

        For i = 1 To NASSETS
            ADD_ATT_SELEC_RNG.Cells(i, PERIODS + 1).formula = "=sum(" & _
                TRANS_CONT_SELEC_RNG.Rows(i).Address & ")/" & _
                EXCESS_RNG.Cells(4, PERIODS + 1).Address
        Next i


'---------------------------19 PASS: Transformed Add Effects I " _
(Continuously Compounding): Inter------------------

Set DST_RNG = DST_RNG.Offset(NASSETS + m)

Set TRANS_CONT_INTER_RNG = Range(DST_RNG.Offset(2, 1), _
DST_RNG.Offset(NASSETS + 1, PERIODS))

TRANS_CONT_INTER_RNG.Cells(-2, 0).value = "Transformed Add Effects I " & _
"(Continuously Compounding): Inter"
TRANS_CONT_INTER_RNG.Cells(-2, 0).Font.Bold = True

TRANS_CONT_INTER_RNG.Cells(0, 0).value = "PERIODS"
For i = 1 To PERIODS
    TRANS_CONT_INTER_RNG.Cells(0, i).formula = "=" & ALPHAS_RNG.Cells(0, i).Address
Next i

For i = 1 To NASSETS
    TRANS_CONT_INTER_RNG.Cells(i, 0).formula = "=" & ALPHAS_RNG.Cells(i, 0).Address
Next i

For j = 1 To PERIODS
        For i = 1 To NASSETS
            TRANS_CONT_INTER_RNG.Cells(i, j).formula = "=" & _
            EXCESS_RNG.Cells(4, j).Address & "*" & _
            ADD_ATT_INTER_RNG.Cells(i, j).Address
        Next i
            TRANS_CONT_INTER_RNG.Cells(NASSETS + 1, j).formula = "=SUM(" & _
            TRANS_CONT_INTER_RNG.Columns(j).Address & ")"
Next j

        For i = 1 To NASSETS
            TRANS_CONT_INTER_RNG.Cells(i, PERIODS + 1).formula = "=SUM(" & _
            TRANS_CONT_INTER_RNG.Rows(i).Address & ")"
        Next i

            TRANS_CONT_INTER_RNG.Cells(NASSETS + 1, PERIODS + 1).formula = "=SUM(" & _
            TRANS_CONT_INTER_RNG.Columns(PERIODS + 1).Address & ")"

TRANS_CONT_INTER_RNG.Cells(NASSETS + 1, 0).value = "Total"

        For i = 1 To NASSETS
            ADD_ATT_INTER_RNG.Cells(i, PERIODS + 1).formula = "=sum(" & _
                TRANS_CONT_INTER_RNG.Rows(i).Address & ")/" & _
                EXCESS_RNG.Cells(4, PERIODS + 1).Address
        Next i

'---------------------------20 PASS: Transformed Add Effects I
'(Continuously Compounding): Total------------------

Set DST_RNG = DST_RNG.Offset(NASSETS + m)

Set TRANS_CONT_TOTAL_RNG = Range(DST_RNG.Offset(2, 1), _
DST_RNG.Offset(NASSETS + 1, PERIODS))

TRANS_CONT_TOTAL_RNG.Cells(-2, 0).value = "Transformed Add Effects I " & _
"(Continuously Compounding): Total"
TRANS_CONT_TOTAL_RNG.Cells(-2, 0).Font.Bold = True

TRANS_CONT_TOTAL_RNG.Cells(0, 0).value = "PERIODS"
For i = 1 To PERIODS
    TRANS_CONT_TOTAL_RNG.Cells(0, i).formula = "=" & ALPHAS_RNG.Cells(0, i).Address
Next i

For i = 1 To NASSETS
    TRANS_CONT_TOTAL_RNG.Cells(i, 0).formula = "=" & ALPHAS_RNG.Cells(i, 0).Address
Next i

For j = 1 To PERIODS
        For i = 1 To NASSETS
            TRANS_CONT_TOTAL_RNG.Cells(i, j).formula = "=" & _
                TRANS_CONT_ALLOC_RNG.Cells(i, j).Address & "+" & _
                TRANS_CONT_SELEC_RNG.Cells(i, j).Address & "+" & _
                TRANS_CONT_INTER_RNG.Cells(i, j).Address
        Next i
            TRANS_CONT_TOTAL_RNG.Cells(NASSETS + 1, j).formula = "=SUM(" & _
            TRANS_CONT_TOTAL_RNG.Columns(j).Address & ")"
Next j

        For i = 1 To NASSETS
            TRANS_CONT_TOTAL_RNG.Cells(i, PERIODS + 1).formula = "=SUM(" & _
            TRANS_CONT_TOTAL_RNG.Rows(i).Address & ")"
        Next i

            TRANS_CONT_TOTAL_RNG.Cells(NASSETS + 1, PERIODS + 1).formula = "=SUM(" & _
            TRANS_CONT_TOTAL_RNG.Columns(PERIODS + 1).Address & ")"

TRANS_CONT_TOTAL_RNG.Cells(NASSETS + 1, 0).value = "Total"


'---------------------------21 PASS: Transformed Add Effects I
'(Continuously Compounding): Alloc------------------

Set DST_RNG = DST_RNG.Offset(NASSETS + m)

Set TRANS_MULT_ALLOC_RNG = Range(DST_RNG.Offset(2, 1), _
DST_RNG.Offset(NASSETS + 1, PERIODS))

TRANS_MULT_ALLOC_RNG.Cells(-2, 0).value = "Transformed Add Effects II " & _
": Multiplicative (Discretely Compounding) - Alloc"

TRANS_MULT_ALLOC_RNG.Cells(-2, 0).Font.Bold = True

TRANS_MULT_ALLOC_RNG.Cells(0, 0).value = "PERIODS"
For i = 1 To PERIODS
    TRANS_MULT_ALLOC_RNG.Cells(0, i).formula = "=" & ALPHAS_RNG.Cells(0, i).Address
Next i

For i = 1 To NASSETS
    TRANS_MULT_ALLOC_RNG.Cells(i, 0).formula = "=" & ALPHAS_RNG.Cells(i, 0).Address
Next i

For j = 1 To PERIODS
        For i = 1 To NASSETS
            TRANS_MULT_ALLOC_RNG.Cells(i, j).formula = "=EXP(" & _
            EXCESS_RNG.Cells(4, j).Address & "*" & _
            ADD_ATT_ALLOC_RNG.Cells(i, j).Address & ")-1"
        Next i
        '=PRODUCT(1+C134:C136)-1
            TRANS_MULT_ALLOC_RNG.Cells(NASSETS + 1, j).FormulaArray = "=PRODUCT(1+" & _
            TRANS_MULT_ALLOC_RNG.Columns(j).Address & ")-1"
Next j

        For i = 1 To NASSETS
            TRANS_MULT_ALLOC_RNG.Cells(i, PERIODS + 1).FormulaArray = "=PRODUCT(1+" & _
            TRANS_MULT_ALLOC_RNG.Rows(i).Address & ")-1"
        Next i

            TRANS_MULT_ALLOC_RNG.Cells(NASSETS + 1, PERIODS + 1).FormulaArray = _
            "=PRODUCT(1+" & TRANS_MULT_ALLOC_RNG.Columns(PERIODS + 1).Address & ")-1"

TRANS_MULT_ALLOC_RNG.Cells(NASSETS + 1, 0).value = "Total"

'---------------------------22 PASS: Transformed Add Effects I
'(Continuously Compounding): Selec------------------

Set DST_RNG = DST_RNG.Offset(NASSETS + m)

Set TRANS_MULT_SELEC_RNG = Range(DST_RNG.Offset(2, 1), _
DST_RNG.Offset(NASSETS + 1, PERIODS))

TRANS_MULT_SELEC_RNG.Cells(-2, 0).value = "Transformed Add Effects II " & _
": Multiplicative (Discretely Compounding) - Selec"
TRANS_MULT_SELEC_RNG.Cells(-2, 0).Font.Bold = True

TRANS_MULT_SELEC_RNG.Cells(0, 0).value = "PERIODS"
For i = 1 To PERIODS
    TRANS_MULT_SELEC_RNG.Cells(0, i).formula = "=" & ALPHAS_RNG.Cells(0, i).Address
Next i

For i = 1 To NASSETS
    TRANS_MULT_SELEC_RNG.Cells(i, 0).formula = "=" & ALPHAS_RNG.Cells(i, 0).Address
Next i

For j = 1 To PERIODS
        For i = 1 To NASSETS
            TRANS_MULT_SELEC_RNG.Cells(i, j).formula = "=EXP(" & _
            EXCESS_RNG.Cells(4, j).Address & "*" & _
            ADD_ATT_SELEC_RNG.Cells(i, j).Address & ")-1"
        Next i
            TRANS_MULT_SELEC_RNG.Cells(NASSETS + 1, j).FormulaArray = "=PRODUCT(1+" & _
            TRANS_MULT_SELEC_RNG.Columns(j).Address & ")-1"
Next j

        For i = 1 To NASSETS
            TRANS_MULT_SELEC_RNG.Cells(i, PERIODS + 1).FormulaArray = "=PRODUCT(1+" & _
            TRANS_MULT_SELEC_RNG.Rows(i).Address & ")-1"
        Next i

            TRANS_MULT_SELEC_RNG.Cells(NASSETS + 1, PERIODS + 1).FormulaArray = _
            "=PRODUCT(1+" & _
            TRANS_MULT_SELEC_RNG.Columns(PERIODS + 1).Address & ")-1"

TRANS_MULT_SELEC_RNG.Cells(NASSETS + 1, 0).value = "Total"

'---------------------------23 PASS: Transformed Add Effects I " _
(Continuously Compounding): Inter------------------

Set DST_RNG = DST_RNG.Offset(NASSETS + m)

Set TRANS_MULT_INTER_RNG = Range(DST_RNG.Offset(2, 1), _
DST_RNG.Offset(NASSETS + 1, PERIODS))

TRANS_MULT_INTER_RNG.Cells(-2, 0).value = "Transformed Add Effects II " & _
": Multiplicative (Discretely Compounding) - Inter"
TRANS_MULT_INTER_RNG.Cells(-2, 0).Font.Bold = True

TRANS_MULT_INTER_RNG.Cells(0, 0).value = "PERIODS"
For i = 1 To PERIODS
    TRANS_MULT_INTER_RNG.Cells(0, i).formula = "=" & ALPHAS_RNG.Cells(0, i).Address
Next i

For i = 1 To NASSETS
    TRANS_MULT_INTER_RNG.Cells(i, 0).formula = "=" & ALPHAS_RNG.Cells(i, 0).Address
Next i

For j = 1 To PERIODS
        For i = 1 To NASSETS
            TRANS_MULT_INTER_RNG.Cells(i, j).formula = "=EXP(" & _
            EXCESS_RNG.Cells(4, j).Address & "*" & _
            ADD_ATT_INTER_RNG.Cells(i, j).Address & ")-1"
        Next i
            TRANS_MULT_INTER_RNG.Cells(NASSETS + 1, j).FormulaArray = "=PRODUCT(1+" & _
            TRANS_MULT_INTER_RNG.Columns(j).Address & ")-1"
Next j

        For i = 1 To NASSETS
            TRANS_MULT_INTER_RNG.Cells(i, PERIODS + 1).FormulaArray = "=PRODUCT(1+" & _
            TRANS_MULT_INTER_RNG.Rows(i).Address & ")-1"
        Next i

            TRANS_MULT_INTER_RNG.Cells(NASSETS + 1, PERIODS + 1).FormulaArray = _
            "=PRODUCT(1+" & _
            TRANS_MULT_INTER_RNG.Columns(PERIODS + 1).Address & ")-1"

TRANS_MULT_INTER_RNG.Cells(NASSETS + 1, 0).value = "Total"

'---------------------------24 PASS: Transformed Add Effects I
'(Continuously Compounding): Total------------------

Set DST_RNG = DST_RNG.Offset(NASSETS + m)

Set TRANS_MULT_TOTAL_RNG = Range(DST_RNG.Offset(2, 1), _
DST_RNG.Offset(NASSETS + 1, PERIODS))

TRANS_MULT_TOTAL_RNG.Cells(-2, 0).value = "Transformed Add Effects II " & _
": Multiplicative (Discretely Compounding) - Total"
TRANS_MULT_TOTAL_RNG.Cells(-2, 0).Font.Bold = True

TRANS_MULT_TOTAL_RNG.Cells(0, 0).value = "PERIODS"
For i = 1 To PERIODS
    TRANS_MULT_TOTAL_RNG.Cells(0, i).formula = "=" & ALPHAS_RNG.Cells(0, i).Address
Next i

For i = 1 To NASSETS
    TRANS_MULT_TOTAL_RNG.Cells(i, 0).formula = "=" & ALPHAS_RNG.Cells(i, 0).Address
Next i

For j = 1 To PERIODS
        For i = 1 To NASSETS
            TRANS_MULT_TOTAL_RNG.Cells(i, j).formula = "=(1+" & _
                TRANS_MULT_ALLOC_RNG.Cells(i, j).Address & ")*(1+" & _
                TRANS_MULT_SELEC_RNG.Cells(i, j).Address & ")*(1+" & _
                TRANS_MULT_INTER_RNG.Cells(i, j).Address & ")-1"
        Next i
            TRANS_MULT_TOTAL_RNG.Cells(NASSETS + 1, j).FormulaArray = "=PRODUCT(1+" & _
            TRANS_MULT_TOTAL_RNG.Columns(j).Address & ")-1"
Next j

        For i = 1 To NASSETS
            TRANS_MULT_TOTAL_RNG.Cells(i, PERIODS + 1).FormulaArray = "=PRODUCT(1+" & _
            TRANS_MULT_TOTAL_RNG.Rows(i).Address & ")-1"
        Next i

            TRANS_MULT_TOTAL_RNG.Cells(NASSETS + 1, PERIODS + 1).formula = "=(1+" & _
                TRANS_MULT_ALLOC_RNG.Cells(NASSETS + 1, PERIODS + 1).Address & ")*(1+" & _
                TRANS_MULT_SELEC_RNG.Cells(NASSETS + 1, PERIODS + 1).Address & ")*(1+" & _
                TRANS_MULT_INTER_RNG.Cells(NASSETS + 1, PERIODS + 1).Address & ")-1"

TRANS_MULT_TOTAL_RNG.Cells(NASSETS + 1, 0).value = "Total"


'---------------------------25 PASS: Geometric Attribution Effects Alloc---------------

Set DST_RNG = DST_RNG.Offset(NASSETS + m)

Set GEO_ATT_ALLOC_RNG = Range(DST_RNG.Offset(2, 1), _
DST_RNG.Offset(NASSETS + 1, PERIODS))

GEO_ATT_ALLOC_RNG.Cells(-2, 0).value = "Geometric Attribution Effects: Alloc"

GEO_ATT_ALLOC_RNG.Cells(-2, 0).Font.Bold = True

GEO_ATT_ALLOC_RNG.Cells(0, 0).value = "PERIODS"
For i = 1 To PERIODS
    GEO_ATT_ALLOC_RNG.Cells(0, i).formula = "=" & ALPHAS_RNG.Cells(0, i).Address
Next i

For i = 1 To NASSETS
    GEO_ATT_ALLOC_RNG.Cells(i, 0).formula = "=" & ALPHAS_RNG.Cells(i, 0).Address
Next i

For j = 1 To PERIODS
        For i = 1 To NASSETS
            GEO_ATT_ALLOC_RNG.Cells(i, j).formula = "=(" & _
                PORT_TOTAL_WEIGHTS.Cells(i, j).Address & "-" & _
                BENCH_TOTAL_REBAL.Cells(i, j).Address & ")*((1+" & _
                BENCH_TOTAL_RET.Cells(i, j).Address & ")/(1+" & _
                BENCH_TOTAL_RET.Cells(NASSETS + 1, j).Address & ")-1)"
        Next i
            GEO_ATT_ALLOC_RNG.Cells(NASSETS + 1, j).formula = "=SUM(" & _
            GEO_ATT_ALLOC_RNG.Columns(j).Address & ")"
Next j

GEO_ATT_ALLOC_RNG.Cells(NASSETS + 1, PERIODS + 1).FormulaArray = "=" & _
    "PRODUCT(1+" & Range(GEO_ATT_ALLOC_RNG.Cells(NASSETS + 1, 1), _
                         GEO_ATT_ALLOC_RNG.Cells(NASSETS + 1, PERIODS)).Address & ")-1"

GEO_ATT_ALLOC_RNG.Cells(NASSETS + 1, 0).value = "Total"


'---------------------------26 PASS: Geometric Attribution Effects Selec---------------

Set DST_RNG = DST_RNG.Offset(NASSETS + m)

Set GEO_ATT_SELEC_RNG = Range(DST_RNG.Offset(2, 1), _
DST_RNG.Offset(NASSETS + 1, PERIODS))

GEO_ATT_SELEC_RNG.Cells(-2, 0).value = "Geometric Attribution Effects: Selec"
GEO_ATT_SELEC_RNG.Cells(-2, 0).Font.Bold = True

GEO_ATT_SELEC_RNG.Cells(0, 0).value = "PERIODS"
For i = 1 To PERIODS
    GEO_ATT_SELEC_RNG.Cells(0, i).formula = "=" & ALPHAS_RNG.Cells(0, i).Address
Next i

For i = 1 To NASSETS
    GEO_ATT_SELEC_RNG.Cells(i, 0).formula = "=" & ALPHAS_RNG.Cells(i, 0).Address
Next i

For j = 1 To PERIODS
        For i = 1 To NASSETS
            GEO_ATT_SELEC_RNG.Cells(i, j).formula = "=" & _
                BENCH_TOTAL_REBAL.Cells(i, j).Address & "*((" & _
                PORT_TOTAL_RET.Cells(i, j).Address & "-" & _
                BENCH_TOTAL_RET.Cells(i, j).Address & ")/(1+" & _
                BENCH_TOTAL_RET.Cells(NASSETS + 1, j).Address & "))"

        Next i
            GEO_ATT_SELEC_RNG.Cells(NASSETS + 1, j).formula = "=SUM(" & _
            GEO_ATT_SELEC_RNG.Columns(j).Address & ")"
Next j
GEO_ATT_SELEC_RNG.Cells(NASSETS + 1, PERIODS + 1).FormulaArray = "=" & _
    "PRODUCT(1+" & Range(GEO_ATT_SELEC_RNG.Cells(NASSETS + 1, 1), _
                         GEO_ATT_SELEC_RNG.Cells(NASSETS + 1, PERIODS)).Address & ")-1"


GEO_ATT_SELEC_RNG.Cells(NASSETS + 1, 0).value = "Total"

'---------------------------27 PASS: Geometric Attribution Effects: Inter---------------


Set DST_RNG = DST_RNG.Offset(NASSETS + m)

Set GEO_ATT_INTER_RNG = Range(DST_RNG.Offset(2, 1), _
DST_RNG.Offset(1 + 1, PERIODS))

GEO_ATT_INTER_RNG.Cells(-2, 0).value = "Geometric Attribution Effects: Inter"
GEO_ATT_INTER_RNG.Cells(-2, 0).Font.Bold = True

GEO_ATT_INTER_RNG.Cells(0, 0).value = "PERIODS"
For i = 1 To PERIODS
    GEO_ATT_INTER_RNG.Cells(0, i).formula = "=" & ALPHAS_RNG.Cells(0, i).Address
Next i

For j = 1 To PERIODS
    GEO_ATT_INTER_RNG.Cells(1 + 1, j).formula = "=(1+" & _
        PORT_TOTAL_RET.Cells(NASSETS + 1, j).Address & ")*(1+" & _
        BENCH_TOTAL_RET.Cells(NASSETS + 1, j).Address & ")/((1+" & _
        CONT_SECOND_Q.Cells(NASSETS + 1, j).Address & ")*(1+" & _
        CONT_THIRD_Q.Cells(NASSETS + 1, j).Address & "))-1"
Next j
GEO_ATT_INTER_RNG.Cells(1 + 1, PERIODS + 1).FormulaArray = "=" & _
    "PRODUCT(1+" & Range(GEO_ATT_INTER_RNG.Cells(1 + 1, 1), _
                         GEO_ATT_INTER_RNG.Cells(1 + 1, PERIODS)).Address & ")-1"


GEO_ATT_INTER_RNG.Cells(1 + 1, 0).value = "Total"


'---------------------------28 PASS: Geometric Attribution Effects: Total---------------

Set DST_RNG = DST_RNG.Offset(1 + m)

Set GEO_ATT_TOTAL_RNG = Range(DST_RNG.Offset(2, 1), _
DST_RNG.Offset(1 + 1, PERIODS))

GEO_ATT_TOTAL_RNG.Cells(-2, 0).value = "Geometric Attribution Effects: Total"
GEO_ATT_TOTAL_RNG.Cells(-2, 0).Font.Bold = True

GEO_ATT_TOTAL_RNG.Cells(0, 0).value = "PERIODS"
For i = 1 To PERIODS
    GEO_ATT_TOTAL_RNG.Cells(0, i).formula = "=" & ALPHAS_RNG.Cells(0, i).Address
Next i

For j = 1 To PERIODS
    GEO_ATT_TOTAL_RNG.Cells(1 + 1, j).formula = "=(1+" & _
        GEO_ATT_ALLOC_RNG.Cells(NASSETS + 1, j).Address & ")*(1+" & _
        GEO_ATT_SELEC_RNG.Cells(NASSETS + 1, j).Address & ")*(1+" & _
        GEO_ATT_INTER_RNG.Cells(1 + 1, j).Address & ")-1"
Next j

GEO_ATT_TOTAL_RNG.Cells(1 + 1, PERIODS + 1).FormulaArray = "=" & _
    "PRODUCT(1+" & Range(GEO_ATT_TOTAL_RNG.Cells(1 + 1, 1), _
                         GEO_ATT_TOTAL_RNG.Cells(1 + 1, PERIODS)).Address & ")-1"

GEO_ATT_TOTAL_RNG.Cells(1 + 1, 0).value = "Total"

RNG_PORT_MULTI_PERIOD_PERFORMANCE_ATTRIBUTION_FUNC = True

Exit Function
ERROR_LABEL:
RNG_PORT_MULTI_PERIOD_PERFORMANCE_ATTRIBUTION_FUNC = False
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : RNG_PORT_PERFORMANCE_ATTRIBUTION_MODELS_FUNC

'DESCRIPTION   : Comparing Performance Attribution Models:
'difference between different performance attribution models.

'Performance Attribution Model
'Performance attribution has the advantage that active management
'decisions can be discussed based on the results over one time
'period, be it monthly or daily. To derive statistically meaningful
'alphas to measure selection skills, for example, requires at least
'36 monthly observations.

'The following routine attempts to explain portfolio performance
'in terms of the active investment management decision "selection"
'and "allocation".

'Characteristics of  a Good Performance Attribution Model
'It is consistent with the investment process and the manager's decision making process.
'It uses a benchmark that reflects the manager's strategic (long-term) asset allocation.
'It measures the effect of the manager's tactical (short-term) allocation shifts.
'It adjusts attribution of returns for systematic risks.

'LIBRARY       : PORTFOLIO
'GROUP         : TRADE_PERFORMANCE
'ID            : 004
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 21/01/2009
'************************************************************************************
'************************************************************************************

Function RNG_PORT_PERFORMANCE_ATTRIBUTION_MODELS_FUNC(ByRef DST_RNG As Excel.Range, _
ByRef LABELS_RNG As Excel.Range, _
ByVal PERIODS As Long, _
Optional ByVal ADD_RNG_NAMES = False)

'INPUTS
'Portfolio segment weights and segment returns
'Benchmark segment weights and segment returns

'Key Notes
'Q1: Benchmark
'Q2: Active Asset Allocation Portfolio
'Q3: Active Stock Selection Portfolio
'Q4: Portfolio
'Calculation of the four notional portfolios...

'Portfolio Benchmark

'Portfolio Segment Weights: Q4=sum[w(p,i)*r(p,i)] / Q2=sum[w(p,i)*r(b,i)]
'Benchmark Sector Weights: Q3=sum[w(b,i)*r(p,i)] / Q1=sum[w(b,i)*r(b,i)]

'w(b,i)... benchmark segment weights
'r(b,i)... benchmark segment returns
'w(p,i)... portfolio segment weights
'w(b,i)... benchmark segment weights

'Strategy Effect: r(b) = Q1 = sum[w(b,i)*r(b,i)]
'Selection Effect: SE = Q3 - Q1 = sum[w(b,i)*{r(p,i)] - r(b,i)}]
'Allocation Effect: AE = Q2 - Q1 = sum[{w(p,i)-w(b,i)}*r(b,i)]
'Interaction Effect: IE = Q4 - Q3 - Q2 + Q1 = sum[{w(p,i)-w(b,i)}*{r(p,i)-r(b,i)}]
'Total Value Added: TVA = r(p) - r(b) = Q4 - Q1 =  SE + AE + IE


'Original Selection Effect: SE = Q3 - Q1
'Interaction effect: IE = Q4 - Q3 - Q2 + Q1
'Modified Selection Effect: SE* = Q3 - Q1 + Q4 - Q3 - Q2 + Q1 = Q4 - Q2
'= sum[w(p,i)*r(p,i)] - sum[w(p,i)*r(b,i)]
'= w(p,i) * { r(p,i) - r(b,i) }

'Within-Benchmark Performance = Allocation + Selection + Interaction
'Out-Of-Benchmark Performance = Sum of Out-Of-Benchmark Segment Contributions

 
Dim i As Long
Dim j As Long
Dim m As Long

Dim NROWS As Long
Dim NPORTS As Long
Dim NASSETS As Long

Dim TEMP_MATRIX As Variant
Dim TEMP_VECTOR As Variant

Dim LABELS_VECTOR As Variant
Dim ASSETS_VECTOR As Variant
Dim POSITION_VECTOR As Variant

Dim BENCH_RET_POS_RNG As Excel.Range
Dim BENCH_RET_RNG As Excel.Range
Dim BENCH_RET_SUM_POS_RNG As Excel.Range
Dim BENCH_RET_SUM_RNG As Excel.Range

Dim BENCH_WEIGHT_POS_RNG As Excel.Range
Dim BENCH_WEIGHTS_RNG As Excel.Range
Dim BENCH_WEIGHT_SUM_POS_RNG As Excel.Range
Dim BENCH_WEIGHT_SUM_RNG As Excel.Range

Dim PORT_RET_POS_RNG As Excel.Range
Dim PORT_RET_RNG As Excel.Range
Dim PORT_RET_SUM_POS_RNG As Excel.Range
Dim PORT_RET_SUM_RNG As Excel.Range

Dim PORT_WEIGHT_POS_RNG As Excel.Range
Dim PORT_WEIGHTS_RNG As Excel.Range
Dim PORT_WEIGHT_SUM_POS_RNG As Excel.Range
Dim PORT_WEIGHT_SUM_RNG As Excel.Range

Dim PERF_POS_RNG As Excel.Range
Dim PERF_RNG As Excel.Range

Dim BENCH_CONT_POS_RNG As Excel.Range
Dim BENCH_CONT_RNG As Excel.Range

Dim PORT_CONT_POS_RNG As Excel.Range
Dim PORT_CONT_RNG As Excel.Range

Dim LEVEL_POS_RNG As Excel.Range
Dim LEVEL_RNG As Excel.Range

Dim SEG_ALLOC_POS_RNG As Excel.Range
Dim SEG_ALLOC_RNG As Excel.Range

Dim SEG_SELEC_POS_RNG As Excel.Range
Dim SEG_SELEC_RNG As Excel.Range

Dim SEG_INTER_POS_RNG As Excel.Range
Dim SEG_INTER_RNG As Excel.Range

Dim SEG_TOTAL_POS_RNG As Excel.Range
Dim SEG_TOTAL_RNG As Excel.Range

Dim SELEC_POS_RNG As Excel.Range
Dim SELEC_RNG As Excel.Range

Dim SELEC_TOTAL_POS_RNG As Excel.Range
Dim SELEC_TOTAL_RNG As Excel.Range

On Error GoTo ERROR_LABEL

RNG_PORT_PERFORMANCE_ATTRIBUTION_MODELS_FUNC = False

m = 4
TEMP_MATRIX = LABELS_RNG
TEMP_MATRIX = MATRIX_TRANSPOSE_FUNC(MATRIX_TRIM_FUNC(MATRIX_TRANSPOSE_FUNC(TEMP_MATRIX), 1, ""))
LABELS_VECTOR = MATRIX_TRANSPOSE_FUNC(MATRIX_GET_ROW_FUNC(TEMP_MATRIX, 1, 1)) 'Portfolio Label Vector
NPORTS = UBound(LABELS_VECTOR, 1)

'---------------1 PASS: SETTING_UP LABELS AND PORT POSITIONS-------------

ReDim POSITION_VECTOR(1 To NPORTS, 1 To 3)
For i = 1 To NPORTS
    TEMP_VECTOR = VECTOR_TRIM_FUNC(MATRIX_GET_COLUMN_FUNC(TEMP_MATRIX, i, 1), "")
    NROWS = UBound(TEMP_VECTOR, 1) - 1 'Remember the Port Headings
    POSITION_VECTOR(i, 3) = NROWS
Next i

ASSETS_VECTOR = MATRIX_ARRAY_CONVERT_FUNC(MATRIX_REMOVE_ROWS_FUNC(TEMP_MATRIX, 1, 1))

NASSETS = UBound(ASSETS_VECTOR, 1)
ReDim TEMP_VECTOR(1 To NASSETS, 1 To 1)
For i = 1 To NASSETS
    TEMP_VECTOR(i, 1) = ASSETS_VECTOR(i)
Next i
TEMP_VECTOR = VECTOR_TRIM_FUNC(TEMP_VECTOR, 0)
NASSETS = UBound(TEMP_VECTOR, 1)

For i = NPORTS To 1 Step -1
    If i = NPORTS Then
        POSITION_VECTOR(i, 1) = NASSETS - POSITION_VECTOR(i, 3) + 1
    Else
        POSITION_VECTOR(i, 1) = POSITION_VECTOR(i + 1, 1) - POSITION_VECTOR(i, 3)
    End If
    POSITION_VECTOR(i, 2) = POSITION_VECTOR(i, 1) + POSITION_VECTOR(i, 3) - 1
Next i

For i = NPORTS To 1 Step -1
    If POSITION_VECTOR(i, 3) = 0 Then
        j = i
        Do While POSITION_VECTOR(j, 3) = 0
            j = j - 1
        Loop
        POSITION_VECTOR(i, 1) = POSITION_VECTOR(j, 1)
        POSITION_VECTOR(i, 2) = POSITION_VECTOR(j, 2)
        LABELS_VECTOR(i, 1) = LABELS_VECTOR(j, 1)
    End If
Next i

'--------------2 PASS: SETTING_UP Benchmark Returns---------------------

Set BENCH_RET_POS_RNG = DST_RNG
With BENCH_RET_POS_RNG
    Set BENCH_RET_RNG = Range(.Offset(2, 1), .Offset(1 + NASSETS, PERIODS))
    If ADD_RNG_NAMES = True Then: BENCH_RET_RNG.name = "BENCH_RET"
    .Offset(0, 0).value = "Benchmark Returns"
    .Offset(0, 0).Font.Bold = True
    
    .Offset(1, 0).value = "Periods"
    .Offset(1, 0).Font.Bold = True
    
    For j = 1 To PERIODS
        .Offset(1, j).value = j
        .Offset(1, j).Font.ColorIndex = 3
    Next j
    
    For i = 1 To NASSETS
        .Offset(1 + i, 0).value = TEMP_VECTOR(i, 1)
        .Offset(1 + i, 0).Font.ColorIndex = 3
    Next i
    
    BENCH_RET_RNG.value = 0
    BENCH_RET_RNG.Font.ColorIndex = 5
    
End With

'--------------3 PASS: SETTING_UP Benchmark Weights---------------------

Set BENCH_WEIGHT_POS_RNG = BENCH_RET_POS_RNG.Offset(1 + NASSETS + m, 0)
With BENCH_WEIGHT_POS_RNG
    Set BENCH_WEIGHTS_RNG = Range(.Offset(2, 1), .Offset(1 + NASSETS, PERIODS))
    If ADD_RNG_NAMES = True Then: BENCH_WEIGHTS_RNG.name = "BENCH_WEIGHT"
    .Offset(0, 0).value = "Benchmark Weights"
    .Offset(0, 0).Font.Bold = True
    
    .Offset(1, 0).value = "Periods"
    .Offset(1, 0).Font.Bold = True
    
    For j = 1 To PERIODS
        .Offset(1, j).formula = "=" & BENCH_RET_POS_RNG.Offset(1, j).Address
    Next j
    
    For i = 1 To NASSETS
        .Offset(1 + i, 0).formula = "=" & BENCH_RET_POS_RNG.Offset(1 + i, 0).Address
    Next i
    
    BENCH_WEIGHTS_RNG.value = 0
    BENCH_WEIGHTS_RNG.Font.ColorIndex = 5
    
    For j = 1 To PERIODS
        .Offset(2 + NASSETS, j).formula = _
            "=SUM(" & BENCH_WEIGHTS_RNG.Columns(j).Address & ")"
    Next j
End With

'--------------4 PASS: SETTING_UP Portfolio Returns---------------------

Set PORT_RET_POS_RNG = BENCH_WEIGHT_POS_RNG.Offset(1 + NASSETS + m, 0)
With PORT_RET_POS_RNG
    Set PORT_RET_RNG = Range(.Offset(2, 1), .Offset(1 + NASSETS, PERIODS))
    If ADD_RNG_NAMES = True Then: PORT_RET_RNG.name = "PORT_RET"
    .Offset(0, 0).value = "Portfolio Returns"
    .Offset(0, 0).Font.Bold = True
    
    .Offset(1, 0).value = "Periods"
    .Offset(1, 0).Font.Bold = True
    
    For j = 1 To PERIODS
        .Offset(1, j).formula = "=" & BENCH_RET_POS_RNG.Offset(1, j).Address
    Next j
    
    For i = 1 To NASSETS
        .Offset(1 + i, 0).formula = "=" & BENCH_RET_POS_RNG.Offset(1 + i, 0).Address
    Next i
    
    PORT_RET_RNG.value = 0
    PORT_RET_RNG.Font.ColorIndex = 5
    
End With

'--------------5 PASS: SETTING_UP Portfolio Weights---------------------

Set PORT_WEIGHT_POS_RNG = PORT_RET_POS_RNG.Offset(1 + NASSETS + m, 0)
With PORT_WEIGHT_POS_RNG
    Set PORT_WEIGHTS_RNG = Range(.Offset(2, 1), .Offset(1 + NASSETS, PERIODS))
    If ADD_RNG_NAMES = True Then: PORT_WEIGHTS_RNG.name = "PORT_WEIGHT"
    .Offset(0, 0).value = "Portfolio Weights"
    .Offset(0, 0).Font.Bold = True
    
    .Offset(1, 0).value = "Periods"
    .Offset(1, 0).Font.Bold = True
    
    For j = 1 To PERIODS
        .Offset(1, j).formula = "=" & BENCH_RET_POS_RNG.Offset(1, j).Address
    Next j
    
    For i = 1 To NASSETS
        .Offset(1 + i, 0).formula = "=" & BENCH_RET_POS_RNG.Offset(1 + i, 0).Address
    Next i
        
    For j = 1 To PERIODS
        .Offset(2 + NASSETS, j).formula = _
            "=SUM(" & PORT_WEIGHTS_RNG.Columns(j).Address & ")"
    Next j
End With

    For j = 1 To PERIODS
        For i = 1 To NASSETS
            PORT_WEIGHTS_RNG.Cells(i, j).formula = "=" & _
            PORT_RET_RNG.Cells(i, j).Address & "+" & _
            BENCH_WEIGHTS_RNG.Cells(i, j).Address
        Next i
    Next j
    PORT_WEIGHTS_RNG.Font.ColorIndex = 3

'--------------6 PASS: SETTING_UP Benchmark Weights Summary ---------------------

Set BENCH_WEIGHT_SUM_POS_RNG = PORT_WEIGHT_POS_RNG.Offset(1 + NASSETS + m, 0)
With BENCH_WEIGHT_SUM_POS_RNG
    Set BENCH_WEIGHT_SUM_RNG = Range(.Offset(2, 1), .Offset(1 + NPORTS, PERIODS))
    If ADD_RNG_NAMES = True Then: BENCH_WEIGHT_SUM_RNG.name = "BENCH_SUM_WEIGHT"

    .Offset(0, 0).value = "Benchmark Weights Summary"
    .Offset(0, 0).Font.Bold = True
    
    .Offset(1, 0).value = "Periods"
    .Offset(1, 0).Font.Bold = True
    
    For j = 1 To PERIODS
        .Offset(1, j).formula = "=" & BENCH_RET_POS_RNG.Offset(1, j).Address
    Next j
    
    For i = 1 To NPORTS
        .Offset(1 + i, 0).value = LABELS_VECTOR(i, 1)
        .Offset(1 + i, 0).Font.ColorIndex = 3
    Next i
        
    For j = 1 To PERIODS
        .Offset(2 + NPORTS, j).formula = _
            "=SUM(" & BENCH_WEIGHT_SUM_RNG.Columns(j).Address & ")"
    Next j
        
End With

    For j = 1 To PERIODS
        For i = 1 To NPORTS
            BENCH_WEIGHT_SUM_RNG.Cells(i, j).formula = "=SUM(" & _
            Range(BENCH_WEIGHTS_RNG.Cells(POSITION_VECTOR(i, 1), j), _
            BENCH_WEIGHTS_RNG.Cells(POSITION_VECTOR(i, 2), j)).Address & ")"
        Next i
    Next j

'--------------7 PASS: SETTING_UP Benchmark Weights Summary ---------------------

Set BENCH_RET_SUM_POS_RNG = BENCH_WEIGHT_SUM_POS_RNG.Offset(1 + NPORTS + m, 0)
With BENCH_RET_SUM_POS_RNG
    Set BENCH_RET_SUM_RNG = Range(.Offset(2, 1), .Offset(1 + NPORTS, PERIODS))
    If ADD_RNG_NAMES = True Then: BENCH_RET_SUM_RNG.name = "BENCH_SUM_RET"

    .Offset(0, 0).value = "Benchmark Returns Summary"
    .Offset(0, 0).Font.Bold = True
    
    .Offset(1, 0).value = "Periods"
    .Offset(1, 0).Font.Bold = True
    
    For j = 1 To PERIODS
        .Offset(1, j).formula = "=" & BENCH_RET_POS_RNG.Offset(1, j).Address
    Next j
    
    For i = 1 To NPORTS
        .Offset(1 + i, 0).formula = "=" & _
            BENCH_WEIGHT_SUM_POS_RNG.Offset(1 + i, 0).Address
    Next i
        
    For j = 1 To PERIODS
        .Offset(2 + NPORTS, j).formula = _
            "=SUMPRODUCT(" & BENCH_RET_SUM_RNG.Columns(j).Address & "," & _
            BENCH_WEIGHT_SUM_RNG.Columns(j).Address & ")"
    Next j
        .Offset(2 + NPORTS, PERIODS + 1).FormulaArray = "=" & _
        "PRODUCT(1+" & Range(.Offset(2 + NPORTS, 1), _
        .Offset(2 + NPORTS, PERIODS)).Address & ")-1"
End With

    For j = 1 To PERIODS
        For i = 1 To NPORTS
            BENCH_RET_SUM_RNG.Cells(i, j).formula = "=SUMPRODUCT(" & _
            Range(BENCH_WEIGHTS_RNG.Cells(POSITION_VECTOR(i, 1), j), _
            BENCH_WEIGHTS_RNG.Cells(POSITION_VECTOR(i, 2), j)).Address & "/" & _
            BENCH_WEIGHT_SUM_RNG.Cells(i, j).Address & "," & _
            Range(BENCH_RET_RNG.Cells(POSITION_VECTOR(i, 1), j), _
            BENCH_RET_RNG.Cells(POSITION_VECTOR(i, 2), j)).Address & ")"
        Next i
    Next j

'--------------8 PASS: SETTING_UP Portfolio Weights Summary ---------------------

Set PORT_WEIGHT_SUM_POS_RNG = BENCH_RET_SUM_POS_RNG.Offset(1 + NPORTS + m, 0)
With PORT_WEIGHT_SUM_POS_RNG
    Set PORT_WEIGHT_SUM_RNG = Range(.Offset(2, 1), .Offset(1 + NPORTS, PERIODS))
    If ADD_RNG_NAMES = True Then: PORT_WEIGHT_SUM_RNG.name = "PORT_SUM_WEIGHT"

    .Offset(0, 0).value = "Portfolio Weights Summary"
    .Offset(0, 0).Font.Bold = True
    
    .Offset(1, 0).value = "Periods"
    .Offset(1, 0).Font.Bold = True
    
    For j = 1 To PERIODS
        .Offset(1, j).formula = "=" & PORT_RET_POS_RNG.Offset(1, j).Address
    Next j
    
    For i = 1 To NPORTS
        .Offset(1 + i, 0).formula = "=" & _
            BENCH_WEIGHT_SUM_POS_RNG.Offset(1 + i, 0).Address
    Next i
        
    For j = 1 To PERIODS
        .Offset(2 + NPORTS, j).formula = _
            "=SUM(" & PORT_WEIGHT_SUM_RNG.Columns(j).Address & ")"
    Next j
        
End With

    For j = 1 To PERIODS
        For i = 1 To NPORTS
            PORT_WEIGHT_SUM_RNG.Cells(i, j).formula = "=SUM(" & _
            Range(PORT_WEIGHTS_RNG.Cells(POSITION_VECTOR(i, 1), j), _
            PORT_WEIGHTS_RNG.Cells(POSITION_VECTOR(i, 2), j)).Address & ")"
        Next i
    Next j

'--------------9 PASS: SETTING_UP Portfolio Returns Summary ---------------------

Set PORT_RET_SUM_POS_RNG = PORT_WEIGHT_SUM_POS_RNG.Offset(1 + NPORTS + m, 0)
With PORT_RET_SUM_POS_RNG
    Set PORT_RET_SUM_RNG = Range(.Offset(2, 1), .Offset(1 + NPORTS, PERIODS))
    If ADD_RNG_NAMES = True Then: PORT_RET_SUM_RNG.name = "PORT_SUM_RET"

    .Offset(0, 0).value = "Portfolio Returns Summary"
    .Offset(0, 0).Font.Bold = True
    
    .Offset(1, 0).value = "Periods"
    .Offset(1, 0).Font.Bold = True
    
    For j = 1 To PERIODS
        .Offset(1, j).formula = "=" & PORT_RET_POS_RNG.Offset(1, j).Address
    Next j
    
    For i = 1 To NPORTS
        .Offset(1 + i, 0).formula = "=" & _
            PORT_WEIGHT_SUM_POS_RNG.Offset(1 + i, 0).Address
    Next i
        
    For j = 1 To PERIODS
        .Offset(2 + NPORTS, j).formula = _
            "=SUMPRODUCT(" & PORT_RET_SUM_RNG.Columns(j).Address & "," & _
            PORT_WEIGHT_SUM_RNG.Columns(j).Address & ")"
    Next j
        .Offset(2 + NPORTS, PERIODS + 1).FormulaArray = "=" & _
        "PRODUCT(1+" & Range(.Offset(2 + NPORTS, 1), _
        .Offset(2 + NPORTS, PERIODS)).Address & ")-1"
End With

For j = 1 To PERIODS
    For i = 1 To NPORTS
        PORT_RET_SUM_RNG.Cells(i, j).formula = "=SUMPRODUCT(" & Range(PORT_WEIGHTS_RNG.Cells(POSITION_VECTOR(i, 1), j), PORT_WEIGHTS_RNG.Cells(POSITION_VECTOR(i, 2), j)).Address & "/" & PORT_WEIGHT_SUM_RNG.Cells(i, j).Address & "," & Range(BENCH_RET_RNG.Cells(POSITION_VECTOR(i, 1), j), BENCH_RET_RNG.Cells(POSITION_VECTOR(i, 2), j)).Address & ")"
    Next i
Next j

'-----------10 PASS: SETTING_UP Performance Weights & Returns---------------

Set PERF_POS_RNG = PORT_RET_SUM_POS_RNG.Offset(1 + NPORTS + m, 0)
With PERF_POS_RNG
    Set PERF_RNG = Range(.Offset(2, 1), .Offset(1 + 3, PERIODS))
    If ADD_RNG_NAMES = True Then: PERF_RNG.name = "PERFORMANCE"

    .Offset(0, 0).value = "Weights & Returns Performance"
    .Offset(0, 0).Font.Bold = True
    .Offset(1, 0).value = "Periods"
    .Offset(1, 0).Font.Bold = True
    For j = 1 To PERIODS
        .Offset(1, j).formula = "=" & PORT_RET_POS_RNG.Offset(1, j).Address
    Next j
    .Offset(2, 0).value = "ADD"
    .Offset(3, 0).value = "LOG"
    .Offset(4, 0).value = "GEO"
    For j = 1 To PERIODS
        .Offset(2 + 3, j).formula = "=" & PERF_POS_RNG.Offset(3, j).Address & "/" & PERF_POS_RNG.Offset(2, j).Address
    Next j
End With

    For j = 1 To PERIODS + 1
        PERF_RNG.Cells(1, j).formula = "=" & _
        PORT_RET_SUM_POS_RNG.Offset(2 + NPORTS, j).Address & "-" & _
        BENCH_RET_SUM_POS_RNG.Offset(2 + NPORTS, j).Address

        PERF_RNG.Cells(2, j).formula = "=LN(1+" & _
        PORT_RET_SUM_POS_RNG.Offset(2 + NPORTS, j).Address & ")-LN(1+" & _
        BENCH_RET_SUM_POS_RNG.Offset(2 + NPORTS, j).Address & ")"

        PERF_RNG.Cells(3, j).formula = "=(1+" & _
        PORT_RET_SUM_POS_RNG.Offset(2 + NPORTS, j).Address & ")/(1+" & _
        BENCH_RET_SUM_POS_RNG.Offset(2 + NPORTS, j).Address & ")-1"
    Next j

'--------------11 PASS: SETTING_UP Benchmark Contribution ---------------------

Set BENCH_CONT_POS_RNG = PERF_POS_RNG.Offset(1 + 3 + m, 0)
With BENCH_CONT_POS_RNG
    Set BENCH_CONT_RNG = Range(.Offset(2, 1), .Offset(1 + NPORTS, PERIODS))
    If ADD_RNG_NAMES = True Then: BENCH_CONT_RNG.name = "BENCH_CONT"

    .Offset(0, 0).value = "Benchmark Contribution"
    .Offset(0, 0).Font.Bold = True
    
    .Offset(1, 0).value = "Periods"
    .Offset(1, 0).Font.Bold = True
    
    For j = 1 To PERIODS
        .Offset(1, j).formula = "=" & PORT_RET_POS_RNG.Offset(1, j).Address
    Next j
    
    For i = 1 To NPORTS
        .Offset(1 + i, 0).formula = "=" & _
            PORT_WEIGHT_SUM_POS_RNG.Offset(1 + i, 0).Address
    Next i
        
    For j = 1 To PERIODS
        .Offset(2 + NPORTS, j).formula = _
            "=SUM(" & BENCH_CONT_RNG.Columns(j).Address & ")"
    Next j

End With

    For j = 1 To PERIODS
        For i = 1 To NPORTS
            BENCH_CONT_RNG.Cells(i, j).formula = "=" & _
            BENCH_WEIGHT_SUM_RNG.Cells(i, j).Address & "*" & _
            BENCH_RET_SUM_RNG.Cells(i, j).Address
        Next i
    Next j

'--------------12 PASS: SETTING_UP Portfolio Contribution ---------------------

Set PORT_CONT_POS_RNG = BENCH_CONT_POS_RNG.Offset(1 + NPORTS + m, 0)
With PORT_CONT_POS_RNG
    Set PORT_CONT_RNG = Range(.Offset(2, 1), .Offset(1 + NPORTS, PERIODS))
    If ADD_RNG_NAMES = True Then: PORT_CONT_RNG.name = "PORT_CONT"

    .Offset(0, 0).value = "Portfolio Contribution"
    .Offset(0, 0).Font.Bold = True
    
    .Offset(1, 0).value = "Periods"
    .Offset(1, 0).Font.Bold = True
    
    For j = 1 To PERIODS
        .Offset(1, j).formula = "=" & PORT_RET_POS_RNG.Offset(1, j).Address
    Next j
    
    For i = 1 To NPORTS
        .Offset(1 + i, 0).formula = "=" & _
            PORT_WEIGHT_SUM_POS_RNG.Offset(1 + i, 0).Address
    Next i
        
    For j = 1 To PERIODS
        .Offset(2 + NPORTS, j).formula = _
            "=SUM(" & PORT_CONT_RNG.Columns(j).Address & ")"
    Next j

End With

    For j = 1 To PERIODS
        For i = 1 To NPORTS
            PORT_CONT_RNG.Cells(i, j).formula = "=" & _
            PORT_WEIGHT_SUM_RNG.Cells(i, j).Address & "*" & _
            PORT_RET_SUM_RNG.Cells(i, j).Address
        Next i
    Next j

'--------------------13 PASS: SETTING_UP Fund Level-------------------


Set LEVEL_POS_RNG = PORT_CONT_POS_RNG.Offset(1 + NPORTS + m, 0)
With LEVEL_POS_RNG
    Set LEVEL_RNG = Range(.Offset(2, 1), .Offset(8, PERIODS))
    If ADD_RNG_NAMES = True Then: LEVEL_RNG.name = "FUND_LEVEL"

    .Offset(0, 0).value = "Fund Level"
    .Offset(0, 0).Font.Bold = True
    
    .Offset(1, 0).value = "Periods"
    .Offset(1, 0).Font.Bold = True
    
    For j = 1 To PERIODS
        .Offset(1, j).formula = "=" & PORT_RET_POS_RNG.Offset(1, j).Address
    Next j
    
        .Offset(2, 0).value = "Q1"
        .Offset(3, 0).value = "Q2"
        .Offset(4, 0).value = "Q3"
        .Offset(5, 0).value = "Q4"
        .Offset(6, 0).value = "ALLOC"
        .Offset(7, 0).value = "SELECT"
        .Offset(8, 0).value = "INTER"
        .Offset(9, 0).value = "VALUE ADDED"
        
    For j = 1 To PERIODS
        .Offset(9, j).formula = _
            "=SUM(" & Range(.Offset(6, j), .Offset(8, j)).Address & ")"
    Next j

End With

    For j = 1 To PERIODS
        LEVEL_RNG.Cells(1, j).formula = "=SUMPRODUCT(" & _
        BENCH_WEIGHT_SUM_RNG.Columns(j).Address & "," & _
        BENCH_RET_SUM_RNG.Columns(j).Address & ")"

        LEVEL_RNG.Cells(2, j).formula = "=SUMPRODUCT(" & _
        PORT_WEIGHT_SUM_RNG.Columns(j).Address & "," & _
        BENCH_RET_SUM_RNG.Columns(j).Address & ")"

        LEVEL_RNG.Cells(3, j).formula = "=SUMPRODUCT(" & _
        BENCH_WEIGHT_SUM_RNG.Columns(j).Address & "," & _
        PORT_RET_SUM_RNG.Columns(j).Address & ")"

        LEVEL_RNG.Cells(4, j).formula = "=SUMPRODUCT(" & _
        PORT_WEIGHT_SUM_RNG.Columns(j).Address & "," & _
        PORT_RET_SUM_RNG.Columns(j).Address & ")"
    
        LEVEL_RNG.Cells(5, j).formula = "=" & _
            LEVEL_RNG.Cells(2, j).Address & "-" & _
            LEVEL_RNG.Cells(1, j).Address
    
        LEVEL_RNG.Cells(6, j).formula = "=" & _
            LEVEL_RNG.Cells(3, j).Address & "-" & _
            LEVEL_RNG.Cells(1, j).Address
    
        LEVEL_RNG.Cells(7, j).formula = "=" & _
            LEVEL_RNG.Cells(4, j).Address & "-" & _
            LEVEL_RNG.Cells(3, j).Address & "-" & _
            LEVEL_RNG.Cells(2, j).Address & "+" & _
            LEVEL_RNG.Cells(1, j).Address
    Next j

        LEVEL_POS_RNG.Offset(2, PERIODS + 1).FormulaArray = "=" & _
                "PRODUCT(1+" & LEVEL_RNG.Rows(1).Address & ")-1"

        LEVEL_POS_RNG.Offset(3, PERIODS + 1).FormulaArray = "=" & _
                "PRODUCT(1+" & LEVEL_RNG.Rows(2).Address & ")-1"
                
        LEVEL_POS_RNG.Offset(4, PERIODS + 1).FormulaArray = "=" & _
                "PRODUCT(1+" & LEVEL_RNG.Rows(3).Address & ")-1"
                
        LEVEL_POS_RNG.Offset(5, PERIODS + 1).FormulaArray = "=" & _
                "PRODUCT(1+" & LEVEL_RNG.Rows(4).Address & ")-1"
                
        LEVEL_POS_RNG.Offset(6, PERIODS + 1).formula = "=" & _
            LEVEL_POS_RNG.Offset(3, PERIODS + 1).Address & "-" & _
            LEVEL_POS_RNG.Offset(2, PERIODS + 1).Address
            
        LEVEL_POS_RNG.Offset(7, PERIODS + 1).formula = "=" & _
            LEVEL_POS_RNG.Offset(4, PERIODS + 1).Address & "-" & _
            LEVEL_POS_RNG.Offset(2, PERIODS + 1).Address

        LEVEL_POS_RNG.Offset(8, PERIODS + 1).formula = "=" & _
            LEVEL_POS_RNG.Offset(5, PERIODS + 1).Address & "-" & _
            LEVEL_POS_RNG.Offset(4, PERIODS + 1).Address & "-" & _
            LEVEL_POS_RNG.Offset(3, PERIODS + 1).Address & "+" & _
            LEVEL_POS_RNG.Offset(2, PERIODS + 1).Address
            
        LEVEL_POS_RNG.Offset(9, PERIODS + 1).FormulaArray = "=" & _
            LEVEL_POS_RNG.Offset(6, PERIODS + 1).Address & "+" & _
            LEVEL_POS_RNG.Offset(7, PERIODS + 1).Address & "+" & _
            LEVEL_POS_RNG.Offset(8, PERIODS + 1).Address



'--------------14 PASS: SETTING_UP Segment Allocation ---------------------

Set SEG_ALLOC_POS_RNG = LEVEL_POS_RNG.Offset(8 + m, 0)
With SEG_ALLOC_POS_RNG
    Set SEG_ALLOC_RNG = Range(.Offset(2, 1), .Offset(1 + NPORTS, PERIODS))
    If ADD_RNG_NAMES = True Then: SEG_ALLOC_RNG.name = "SEG_ALLOC"

    .Offset(0, 0).value = "Segment allocation"
    .Offset(0, 0).Font.Bold = True
    
    .Offset(1, 0).value = "Periods"
    .Offset(1, 0).Font.Bold = True
    
    For j = 1 To PERIODS
        .Offset(1, j).formula = "=" & PORT_RET_POS_RNG.Offset(1, j).Address
    Next j
    
    For i = 1 To NPORTS
        .Offset(1 + i, 0).formula = "=" & _
            PORT_WEIGHT_SUM_POS_RNG.Offset(1 + i, 0).Address
    Next i
        
    For j = 1 To PERIODS
        .Offset(2 + NPORTS, j).formula = _
            "=SUM(" & SEG_ALLOC_RNG.Columns(j).Address & ")"
    Next j

        .Offset(2 + NPORTS, PERIODS + 1).FormulaArray = "=" & _
        "PRODUCT(1+" & Range(.Offset(2 + NPORTS, 1), _
        .Offset(2 + NPORTS, PERIODS)).Address & ")-1"
        
End With

    For j = 1 To PERIODS
        For i = 1 To NPORTS
            SEG_ALLOC_RNG.Cells(i, j).formula = "=(" & _
            PORT_WEIGHT_SUM_RNG.Cells(i, j).Address & "-" & _
            BENCH_WEIGHT_SUM_RNG.Cells(i, j).Address & ") * (" & _
            BENCH_RET_SUM_RNG.Cells(i, j).Address & "-" & _
            BENCH_RET_SUM_POS_RNG.Offset(2 + NPORTS, j).Address & ")"
        Next i
    Next j

'--------------15 PASS: SETTING_UP Segment Selection ---------------------

Set SEG_SELEC_POS_RNG = SEG_ALLOC_POS_RNG.Offset(1 + NPORTS + m, 0)
With SEG_SELEC_POS_RNG
    Set SEG_SELEC_RNG = Range(.Offset(2, 1), .Offset(1 + NPORTS, PERIODS))
    If ADD_RNG_NAMES = True Then: SEG_SELEC_RNG.name = "SEG_SELEC"

    .Offset(0, 0).value = "Segment Selection"
    .Offset(0, 0).Font.Bold = True
    
    .Offset(1, 0).value = "Periods"
    .Offset(1, 0).Font.Bold = True
    
    For j = 1 To PERIODS
        .Offset(1, j).formula = "=" & PORT_RET_POS_RNG.Offset(1, j).Address
    Next j
    
    For i = 1 To NPORTS
        .Offset(1 + i, 0).formula = "=" & _
            PORT_WEIGHT_SUM_POS_RNG.Offset(1 + i, 0).Address
    Next i
        
    For j = 1 To PERIODS
        .Offset(2 + NPORTS, j).formula = _
            "=SUM(" & SEG_SELEC_RNG.Columns(j).Address & ")"
    Next j

        .Offset(2 + NPORTS, PERIODS + 1).FormulaArray = "=" & _
        "PRODUCT(1+" & Range(.Offset(2 + NPORTS, 1), _
        .Offset(2 + NPORTS, PERIODS)).Address & ")-1"
        
End With
    
For j = 1 To PERIODS
    For i = 1 To NPORTS
        SEG_SELEC_RNG.Cells(i, j).formula = "=" & BENCH_WEIGHT_SUM_RNG.Cells(i, j).Address & "*(" & PORT_RET_SUM_RNG.Cells(i, j).Address & "-" & BENCH_RET_SUM_RNG.Cells(i, j).Address & ")"
    Next i
Next j

'--------------15 PASS: SETTING_UP Segment Inter ---------------------

Set SEG_INTER_POS_RNG = SEG_SELEC_POS_RNG.Offset(1 + NPORTS + m, 0)
With SEG_INTER_POS_RNG
    Set SEG_INTER_RNG = Range(.Offset(2, 1), .Offset(1 + NPORTS, PERIODS))
    If ADD_RNG_NAMES = True Then: SEG_INTER_RNG.name = "SEG_INTER"
    .Offset(0, 0).value = "Segment Inter"
    .Offset(0, 0).Font.Bold = True
    .Offset(1, 0).value = "Periods"
    .Offset(1, 0).Font.Bold = True
    For j = 1 To PERIODS
        .Offset(1, j).formula = "=" & PORT_RET_POS_RNG.Offset(1, j).Address
    Next j
    For i = 1 To NPORTS
        .Offset(1 + i, 0).formula = "=" & PORT_WEIGHT_SUM_POS_RNG.Offset(1 + i, 0).Address
    Next i
    For j = 1 To PERIODS
        .Offset(2 + NPORTS, j).formula = "=SUM(" & SEG_INTER_RNG.Columns(j).Address & ")"
    Next j
    .Offset(2 + NPORTS, PERIODS + 1).FormulaArray = "=" & "PRODUCT(1+" & Range(.Offset(2 + NPORTS, 1), .Offset(2 + NPORTS, PERIODS)).Address & ")-1"
End With
    
For j = 1 To PERIODS
    For i = 1 To NPORTS
        SEG_INTER_RNG.Cells(i, j).formula = "=(" & _
        PORT_WEIGHT_SUM_RNG.Cells(i, j).Address & "-" & _
        BENCH_WEIGHT_SUM_RNG.Cells(i, j).Address & ")*(" & _
        PORT_RET_SUM_RNG.Cells(i, j).Address & "-" & _
        BENCH_RET_SUM_RNG.Cells(i, j).Address & ")"
    Next i
Next j

'------------16 PASS: SETTING_UP Segment Total Added Value -----------------

Set SEG_TOTAL_POS_RNG = SEG_INTER_POS_RNG.Offset(1 + NPORTS + m, 0)
With SEG_TOTAL_POS_RNG
    Set SEG_TOTAL_RNG = Range(.Offset(2, 1), .Offset(1 + NPORTS, PERIODS))
    If ADD_RNG_NAMES = True Then: SEG_TOTAL_RNG.name = "SEG_TOTAL"

    .Offset(0, 0).value = "Segment Total Added Value"
    .Offset(0, 0).Font.Bold = True
    .Offset(1, 0).value = "Periods"
    .Offset(1, 0).Font.Bold = True
    For j = 1 To PERIODS
        .Offset(1, j).formula = "=" & PORT_RET_POS_RNG.Offset(1, j).Address
    Next j
    For i = 1 To NPORTS
        .Offset(1 + i, 0).formula = "=" & PORT_WEIGHT_SUM_POS_RNG.Offset(1 + i, 0).Address
    Next i
    For j = 1 To PERIODS
        .Offset(2 + NPORTS, j).formula = "=SUM(" & SEG_TOTAL_RNG.Columns(j).Address & ")"
    Next j
    .Offset(2 + NPORTS, PERIODS + 1).FormulaArray = "=" & "PRODUCT(1+" & Range(.Offset(2 + NPORTS, 1), .Offset(2 + NPORTS, PERIODS)).Address & ")-1"
End With
    
For j = 1 To PERIODS
    For i = 1 To NPORTS
        SEG_TOTAL_RNG.Cells(i, j).formula = "=" & SEG_ALLOC_RNG.Cells(i, j).Address & "+" & SEG_SELEC_RNG.Cells(i, j).Address & "+" & SEG_INTER_RNG.Cells(i, j).Address
    Next i
Next j

'--------------17 PASS: SETTING_UP Asset Selection---------------------

Set SELEC_POS_RNG = SEG_TOTAL_POS_RNG.Offset(1 + NPORTS + m, 0)
With SELEC_POS_RNG
    Set SELEC_RNG = Range(.Offset(2, 1), .Offset(1 + NASSETS, PERIODS))
    If ADD_RNG_NAMES = True Then: SELEC_RNG.name = "ASSET_SELEC"
    .Offset(0, 0).value = "Selection"
    .Offset(0, 0).Font.Bold = True
    .Offset(1, 0).value = "Periods"
    .Offset(1, 0).Font.Bold = True
    For j = 1 To PERIODS
        .Offset(1, j).formula = "=" & BENCH_RET_POS_RNG.Offset(1, j).Address
    Next j
    For i = 1 To NASSETS
        .Offset(1 + i, 0).formula = "=" & BENCH_RET_POS_RNG.Offset(1 + i, 0).Address
    Next i
End With
    
For j = 1 To PERIODS
    For i = 1 To NASSETS
        SELEC_RNG.Cells(i, j).formula = "=(" & PORT_WEIGHTS_RNG.Cells(i, j).Address & "-" & BENCH_WEIGHTS_RNG.Cells(i, j).Address & ")*(" & BENCH_RET_RNG.Cells(i, j).Address & "-" & BENCH_RET_SUM_POS_RNG.Offset(2 + NPORTS, j).Address & ")"
    Next i
Next j

'--------------18 PASS: SETTING_UP Benchmark Weights Summary ---------------------

Set SELEC_TOTAL_POS_RNG = SELEC_POS_RNG.Offset(1 + NASSETS + m, 0)
With SELEC_TOTAL_POS_RNG
    Set SELEC_TOTAL_RNG = Range(.Offset(2, 1), .Offset(1 + NPORTS, PERIODS))
    If ADD_RNG_NAMES = True Then: SELEC_TOTAL_RNG.name = "SELEC_TOTAL"

    .Offset(0, 0).value = "Selection Total Added Value"
    .Offset(0, 0).Font.Bold = True
    
    .Offset(1, 0).value = "Periods"
    .Offset(1, 0).Font.Bold = True
    
    For j = 1 To PERIODS
        .Offset(1, j).formula = "=" & BENCH_RET_POS_RNG.Offset(1, j).Address
    Next j
    
    For i = 1 To NPORTS
        .Offset(1 + i, 0).formula = "=" & PORT_WEIGHT_SUM_POS_RNG.Offset(1 + i, 0).Address
    Next i
        
    For j = 1 To PERIODS + 1
        .Offset(2 + NPORTS, j).formula = "=SUM(" & SELEC_TOTAL_RNG.Columns(j).Address & ")"
    Next j
        
End With

For j = 1 To PERIODS
    For i = 1 To NPORTS
        SELEC_TOTAL_RNG.Cells(i, j).formula = "=SUM(" & Range(SELEC_RNG.Cells(POSITION_VECTOR(i, 1), j), SELEC_RNG.Cells(POSITION_VECTOR(i, 2), j)).Address & ")"
    Next i
Next j
For i = 1 To NPORTS
    SELEC_TOTAL_RNG.Cells(i, PERIODS + 1).FormulaArray = _
    "=PRODUCT(1+" & SELEC_TOTAL_RNG.Rows(i).Address & ")-1"
Next i

RNG_PORT_PERFORMANCE_ATTRIBUTION_MODELS_FUNC = True

Exit Function
ERROR_LABEL:
RNG_PORT_PERFORMANCE_ATTRIBUTION_MODELS_FUNC = False
End Function

'------------------------------------------------------------------------------------

'Key Notes from Professor Steven Forster (Ivey School of Business)

'Asset allocation policy involves the establishment of normal
'asset class weights and is an integral part of investment policy.
'Active asset allocation is the process of managing asset class
'weights relative to the normal weights over time; its aim is to
'enhance the managed portfolio's risk/return tradeoff.

'Quadrant i

'Indicates the total return provided by the investment policy adopted
'by the plan sponsor. The policy "portfolio" thus represents a constant,
'normal allocation to passive asset classes. Investment policy, then,
'identifies the plan's normal portfolio composition. Calculating the
'policy return involves applying the normal weights of each investable
'asset class to the respective passive returns.

'Quadrant II

'Rerports the return attributable to a portfolio reflecting both policy
'and active asset allocation. Whether active allocation involves
'anticipating price moves (market timing) or reacting to market disequilibria
'(fundamental analysis), it results in the under or overweighting of asset
'classes relative to the normal weights indentified by policy.

'The aim of active allocation is to enhance the return and/or reduce the
'risk of the portfolio relative to its policy benchmark. The policy and
'active asset allocation return is computed by applying the actual asset
'class weights to their respective passive benchmark returns.

'Quadrant III

'Presents the returns to a portfolio attributable to policy and security
'selection. Security selection involves active investment decisions concerning
'the securities within each asset class. This framework specifies that the
'return from policy and security selection is obtained by applying the normal
'asset class weights to the actual active returns achieved in each asset class.

'Quadrant IV

'Represents the actual return realized by the plan over the period of
'performance evaluation. This is the result of the plan's actual asset
'class weights interacting with the actual asset class returns.

'The active contribution to total performance is composed of active asset
'allocation, security selection, and the effects of a cross-product term
'that measures the interaction of the security selection and active asset
'allocation decisions.

'Results:

'a) Policy return --> Is the passive portfolio benchmark return, calculated
'as the sum of the policy weighted passive asset class returns, using the
'10 year average asset class weights and a suitable passive index for each
'asset class.

'b) Policy and active asset allocation return --> Is calculated using the
'actual active weights and the appropriate passive index returns. The
'policy and security selection return is calculated using the policy
'weights and the actual active returns.

'Active management not only had no measurable impact on returns, but
'(in the absence of a proxy for the variability of the respetive pension
'liabilities), it appears to have increased riskby a small margin. Given
'the higher risk level of the policy and security selection portfolio, it
'is evident that security selection contributed to actual plan risk. Active
'asset allocation appears to have had a negligible impact on risk relative
'to the benchmark policy.

'BECAUSE ACTIVE ASSET ALLOCATION IS THE PROCESS OF MANAGING ASSET
'CLASS WEIGHTS RELATIVE TO THE NORMAL WEIGHTS, ACTIVE MANAGEMENT IS
'CONDITIONAL ON THE INVESTMENT POLICY. THUS ACTIVE RETURNS ARE
'CONDITIONALLY DISTRIBUTED ON THE POLICY RETURN DISTRIBUTION.

'Besides shifting asset class weights - i.e., external risk positioning -
'a manager or sponsor can change exposure to an asset class within a
'portfolio component - internal risk positoining. Internal methods include
'altering the component's BETA or duration by using long or short feature
'positions, carrying cash or hedging the currency component. Looking at any
'single riskpositioning activity, external or internal, will not give a
'complete or accurate measure of the active portfolio managemnet effect.

'------------------------------------------------------------------------------------
