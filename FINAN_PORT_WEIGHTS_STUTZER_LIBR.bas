Attribute VB_Name = "FINAN_PORT_WEIGHTS_STUTZER_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.

'************************************************************************************
'************************************************************************************
'FUNCTION      : PORT_STUTZER_OBJ_FUNC
'DESCRIPTION   : Performance and Risk Aversion of Funds with Benchmarks
'LIBRARY       : PORTFOLIO
'GROUP         : WEIGHTS_STUTZER
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 21/01/2009
'************************************************************************************
'************************************************************************************

Function PORT_STUTZER_OBJ_FUNC(ByVal GAMMA_VAL As Double, _
ByRef BENCH_RNG As Variant, _
ByRef PORT_RNG As Variant)

'Consider the probability of underperformance, namely Pr[S=0], and
'identify those portfolios for which this decays to zero as n approach infinity.

'The faster that Pr[S=0] --> 0 , the better the portfolio ... since then the
'probability of beating the benchmark, Pr{S>0], is greater.

'The underperformance rate of decay is then the Stutzer Index.
'We vary gamma (that's g) in order to get the largest value of D.
'(Expressed as a percentage, that's our Stutzer Index for this portolio.)
'bunch of allocations and find the "best" (in the sense of the largest
'D-value) is like so.

'FIRST CHART: (EVOLUTION OF D)
'Y-AXIS: DATES or No. Index
'X-AXIS: TEMP_MATRIX(i,6)

'SENSITIVITY: CHART
'X-AXIS = GAMMA
'Y-AXIS = D = TEMP_MATRIX(NROWS,6) 'While doing the sensitivity table
'of the allocation


'ARTICLE: http://www.math.uwaterloo.ca/~sas-adm/fall2004/FosterStutzerFeb04.pdf

Dim i As Long
Dim NROWS As Long

Dim TEMP_SUM As Double

Dim PORT_VECTOR As Variant
Dim BENCH_VECTOR As Variant

Dim TEMP_MATRIX As Variant

On Error GoTo ERROR_LABEL

PORT_VECTOR = PORT_RNG
If UBound(PORT_VECTOR, 1) = 1 Then
    PORT_VECTOR = MATRIX_TRANSPOSE_FUNC(PORT_VECTOR)
End If
BENCH_VECTOR = BENCH_RNG
If UBound(BENCH_VECTOR, 1) = 1 Then: _
    BENCH_VECTOR = MATRIX_TRANSPOSE_FUNC(BENCH_VECTOR)
End If
If UBound(PORT_VECTOR, 1) <> UBound(BENCH_VECTOR, 1) Then: GoTo ERROR_LABEL
NROWS = UBound(PORT_VECTOR, 1)

ReDim TEMP_MATRIX(0 To NROWS, 1 To 6)

TEMP_MATRIX(0, 1) = "INDEX"
TEMP_MATRIX(0, 2) = "PORT GAINS"
TEMP_MATRIX(0, 3) = "BENCH GAINS"
TEMP_MATRIX(0, 4) = "-R^-GAMMA"
TEMP_MATRIX(0, 5) = "AVG[-R^-GAMMA]"
TEMP_MATRIX(0, 6) = "-(1/N)LOG[AVG(-R^-GAMMA)]"

TEMP_MATRIX(1, 1) = 1
TEMP_MATRIX(1, 2) = (PORT_VECTOR(1, 1) + 1)
TEMP_MATRIX(1, 3) = BENCH_VECTOR(1, 1) + 1

TEMP_MATRIX(1, 4) = (-1) * (TEMP_MATRIX(1, 2) / TEMP_MATRIX(1, 3)) ^ (-GAMMA_VAL)

TEMP_SUM = TEMP_MATRIX(1, 4)
TEMP_MATRIX(1, 5) = TEMP_SUM / 1
TEMP_MATRIX(1, 6) = -Log(-TEMP_MATRIX(1, 5)) / TEMP_MATRIX(1, 1)

For i = 2 To NROWS
    TEMP_MATRIX(i, 1) = i
    TEMP_MATRIX(i, 2) = (PORT_VECTOR(i, 1) + 1) * TEMP_MATRIX(i - 1, 2) 'PORTFOLIO GAINS
    TEMP_MATRIX(i, 3) = (BENCH_VECTOR(i, 1) + 1) * (TEMP_MATRIX(i - 1, 3)) 'BENCHMARK GAINS
    TEMP_MATRIX(i, 4) = (-1) * (TEMP_MATRIX(i, 2) / TEMP_MATRIX(i, 3)) ^ (-GAMMA_VAL)
    TEMP_SUM = TEMP_SUM + TEMP_MATRIX(i, 4)
    TEMP_MATRIX(i, 5) = TEMP_SUM / i
    TEMP_MATRIX(i, 6) = -Log(-TEMP_MATRIX(i, 5)) / TEMP_MATRIX(i, 1)
Next i

PORT_STUTZER_OBJ_FUNC = TEMP_MATRIX

'PORT_STUTZER_OBJ_FUNC = TEMP_MATRIX(NROWS, 6)
'---> USE SOLVER TO FIND THE BEST
'ALLOCATION OF (D) By Changing Cells: WEIGHTS_VECTOR --> REMEMBER YOU
'MUST PUT A CONSTRAINT OF NO_SHORT SALES IS ALLOWED

Exit Function
ERROR_LABEL:
PORT_STUTZER_OBJ_FUNC = Err.number
End Function
