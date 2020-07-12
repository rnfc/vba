Attribute VB_Name = "FINAN_PORT_SIMUL_FRONTIER_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.

'************************************************************************************
'************************************************************************************
'FUNCTION      : PORT_FRONTIER_RESAMPLING_FUNC

'- Efficient Asset Management, Richard O. Michaud, 1998
'- Portfolio Resampling: Review and Critique, Bernd Scherrer,
'  Financial Analyst Journal, Noc/Dec 2002

'  http://papers.ssrn.com/sol3/papers.cfm?abstract_id=377503

'LIBRARY       : PORTFOLIO
'GROUP         : FRONTIER_RESAMPLING
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'UPDATE        : 04/01/2008
'************************************************************************************
'************************************************************************************

Function PORT_FRONTIER_RESAMPLING_FUNC(ByRef RISK_TOLERANCE_RNG As Variant, _
ByRef EXPECTED_RETURNS_RNG As Variant, _
ByRef COVAR_RNG As Variant, _
Optional ByVal BUDGET_VAL As Double = 1, _
Optional ByRef LOWER_RNG As Variant = 0, _
Optional ByRef UPPER_RNG As Variant = 1, _
Optional ByVal NO_OBS As Long = 12, _
Optional ByVal nLOOPS As Long = 50, _
Optional ByVal VERSION As Integer = 0, _
Optional ByVal TE_MAX_VAL As Double = 0.1)

Dim h As Long
Dim i As Long
Dim j As Long
Dim k As Long
Dim l As Long

Dim NO_ASSETS As Long
Dim NO_TRIALS As Long

Dim TEMP_GROUP As Variant

Dim XDATA_VECTOR As Variant
Dim YDATA_VECTOR As Variant

Dim XTEMP_VECTOR As Variant
Dim YTEMP_VECTOR As Variant

Dim LOWER_VECTOR As Variant
Dim UPPER_VECTOR As Variant
Dim SIGMA_VECTOR As Variant

Dim COVAR_MATRIX As Variant
Dim CORREL_MATRIX As Variant

Dim TOLERANCE_VECTOR As Variant
Dim EXPECTED_VECTOR As Variant

On Error GoTo ERROR_LABEL

'------------------------------------------------------------------------------
'VERSION --> Uncertainty Type
'------------------------------------------------------------------------------
'0 = Both risk and return
'1 = risk only
'2 = return only
'------------------------------------------------------------------------------

TOLERANCE_VECTOR = RISK_TOLERANCE_RNG
If UBound(TOLERANCE_VECTOR, 1) = 1 Then: _
    TOLERANCE_VECTOR = MATRIX_TRANSPOSE_FUNC(TOLERANCE_VECTOR)
NO_TRIALS = UBound(TOLERANCE_VECTOR, 1)

EXPECTED_VECTOR = EXPECTED_RETURNS_RNG
If UBound(EXPECTED_VECTOR, 1) = 1 Then: _
    EXPECTED_VECTOR = MATRIX_TRANSPOSE_FUNC(EXPECTED_VECTOR)

NO_ASSETS = UBound(EXPECTED_VECTOR, 1)

COVAR_MATRIX = COVAR_RNG
If UBound(COVAR_MATRIX, 1) <> UBound(COVAR_MATRIX, 2) Then: GoTo ERROR_LABEL
If UBound(COVAR_MATRIX, 1) <> NO_ASSETS Then: GoTo ERROR_LABEL

If IsArray(LOWER_RNG) = True Then
    LOWER_VECTOR = LOWER_RNG
    If UBound(LOWER_VECTOR, 1) = 1 Then: _
        LOWER_VECTOR = MATRIX_TRANSPOSE_FUNC(LOWER_VECTOR)
    If NO_ASSETS <> UBound(LOWER_VECTOR, 1) Then: GoTo ERROR_LABEL
Else
    ReDim LOWER_VECTOR(1 To NO_ASSETS, 1 To 1)
    For i = 1 To NO_ASSETS
        LOWER_VECTOR(i, 1) = LOWER_RNG
    Next i
End If

If IsArray(UPPER_RNG) = True Then
    UPPER_VECTOR = UPPER_RNG
    If UBound(UPPER_VECTOR, 1) = 1 Then: _
        UPPER_VECTOR = MATRIX_TRANSPOSE_FUNC(UPPER_VECTOR)
    If NO_ASSETS <> UBound(UPPER_VECTOR, 1) Then: GoTo ERROR_LABEL
Else
    ReDim UPPER_VECTOR(1 To NO_ASSETS, 1 To 1)
    For i = 1 To NO_ASSETS
        UPPER_VECTOR(i, 1) = UPPER_RNG
    Next i
End If

CORREL_MATRIX = PORT_CORREL_COVAR_FUNC(COVAR_MATRIX)
SIGMA_VECTOR = PORT_SIGMA_COVAR_FUNC(COVAR_MATRIX)


ReDim YDATA_VECTOR(1 To NO_ASSETS, 1 To 1)
ReDim YTEMP_VECTOR(1 To NO_ASSETS, 1 To 1)

ReDim TEMP_MATRIX(0 To NO_TRIALS * nLOOPS, 1 To 8)

TEMP_MATRIX(0, 1) = "RESAMPLED FRONTIER RETURN"
TEMP_MATRIX(0, 2) = "RESAMPLED FRONTIER VOLATILITY"
TEMP_MATRIX(0, 3) = "MARKOWITZ PORTFOLIO RETURN"
TEMP_MATRIX(0, 4) = "MARKOWITZ PORTFOLIO VOLATILITY"
TEMP_MATRIX(0, 5) = "RESAMPLED PORTFOLIO RETURN"
TEMP_MATRIX(0, 6) = "RESAMPLED PORTFOLIO VOLATILITY"
TEMP_MATRIX(0, 7) = "TE"
TEMP_MATRIX(0, 8) = "INCL/EXCL"

' loop through risk tolerances
For k = 1 To NO_TRIALS

    ' helper constant to format result matrix
    h = (k - 1) * nLOOPS
    
    ' Calculate Markowitz Efficient Frontier
    XTEMP_VECTOR = PORT_SHARPE_WEIGHTS_OPTIMIZER_FUNC(BUDGET_VAL, _
                   TOLERANCE_VECTOR(k, 1), _
                   EXPECTED_VECTOR, COVAR_MATRIX, _
                   LOWER_VECTOR, UPPER_VECTOR)
                       
    ' counter for resampled portfolios added to resampled frontier
    l = 0
    
    ' loop through simulations
    For i = 1 To nLOOPS
        
        ' resample optimization inputs
        TEMP_GROUP = PORT_FRONTIER_NORMAL_SIMULATION_FUNC(NO_OBS, EXPECTED_VECTOR, _
                     SIGMA_VECTOR, CORREL_MATRIX, 3)
        
        Select Case VERSION ' optimize
            Case 0 ' Uncertainty about risk & return
                XDATA_VECTOR = _
                     PORT_SHARPE_WEIGHTS_OPTIMIZER_FUNC(BUDGET_VAL, _
                     TOLERANCE_VECTOR(k, 1), _
                     TEMP_GROUP(2), TEMP_GROUP(3), LOWER_VECTOR, _
                     UPPER_VECTOR)
            Case 1 ' Uncertainty about risk only
                XDATA_VECTOR = _
                     PORT_SHARPE_WEIGHTS_OPTIMIZER_FUNC(BUDGET_VAL, _
                     TOLERANCE_VECTOR(k, 1), _
                     EXPECTED_VECTOR, TEMP_GROUP(3), LOWER_VECTOR, _
                     UPPER_VECTOR)
            Case Else ' Uncertainty about returns only
                XDATA_VECTOR = _
                     PORT_SHARPE_WEIGHTS_OPTIMIZER_FUNC(BUDGET_VAL, _
                     TOLERANCE_VECTOR(k, 1), _
                     TEMP_GROUP(2), COVAR_MATRIX, LOWER_VECTOR, _
                     UPPER_VECTOR)
        End Select
        
        ' calculate average simulated portfolio allocation & active weights
        For j = 1 To NO_ASSETS
            YTEMP_VECTOR(j, 1) = XDATA_VECTOR(j, 1) - XTEMP_VECTOR(j, 1)
        Next j

        ' calculate tracking error
        TEMP_MATRIX(h + i, 7) = _
            PORT_WEIGHTED_SIGMA_COVAR_FUNC(COVAR_MATRIX, YTEMP_VECTOR)

        ' exclude extreme portfolio
        TEMP_MATRIX(h + i, 8) = TEMP_MATRIX(h + i, 7) <= TE_MAX_VAL
        If TEMP_MATRIX(h + i, 8) Then
            ' calculate risk & return of resampled portfolio
            TEMP_MATRIX(h + i, 1) = PORT_WEIGHTED_RETURN2_FUNC(EXPECTED_VECTOR, XDATA_VECTOR)
            TEMP_MATRIX(h + i, 2) = PORT_WEIGHTED_SIGMA_COVAR_FUNC(COVAR_MATRIX, XDATA_VECTOR)
            ' add resampled portfolio to resampled frontier
            For j = 1 To NO_ASSETS
                YDATA_VECTOR(j, 1) = YDATA_VECTOR(j, 1) + XDATA_VECTOR(j, 1)
            Next j
            ' count resampled portfolio added to resampled frontier
            l = l + 1
        Else
            TEMP_MATRIX(h + i, 1) = CVErr(xlErrNA)
            TEMP_MATRIX(h + i, 2) = CVErr(xlErrNA)
        End If
        
        If i <> 1 Then
            TEMP_MATRIX(h + i, 3) = CVErr(xlErrNA)
            TEMP_MATRIX(h + i, 4) = CVErr(xlErrNA)
            TEMP_MATRIX(h + i, 5) = CVErr(xlErrNA)
            TEMP_MATRIX(h + i, 6) = CVErr(xlErrNA)
        End If
    Next i

    ' calculate risk & return for Markowitz frontier
    TEMP_MATRIX(h + 1, 3) = PORT_WEIGHTED_RETURN2_FUNC(EXPECTED_VECTOR, XTEMP_VECTOR)
    TEMP_MATRIX(h + 1, 4) = PORT_WEIGHTED_SIGMA_COVAR_FUNC(COVAR_MATRIX, XTEMP_VECTOR)
    ' Calculate risk & return of average allocation (=resampled frontier)
    For j = 1 To NO_ASSETS
        YDATA_VECTOR(j, 1) = YDATA_VECTOR(j, 1) / l
    Next j
    TEMP_MATRIX(h + 1, 5) = PORT_WEIGHTED_RETURN2_FUNC(EXPECTED_VECTOR, YDATA_VECTOR)
    TEMP_MATRIX(h + 1, 6) = PORT_WEIGHTED_SIGMA_COVAR_FUNC(COVAR_MATRIX, YDATA_VECTOR)
    ' Markowitz portfolio does not have TE
    TEMP_MATRIX(h + 1, 7) = CVErr(xlErrNA)
    TEMP_MATRIX(h + 1, 8) = CVErr(xlErrNA)
Next k

PORT_FRONTIER_RESAMPLING_FUNC = TEMP_MATRIX

Exit Function
ERROR_LABEL:
PORT_FRONTIER_RESAMPLING_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : PORT_FRONTIER_NORMAL_SIMULATION_FUNC
'DESCRIPTION   :
'LIBRARY       : PORTFOLIO
'GROUP         : FRONTIER_RESAMPLING
'ID            : 002
'UPDATE        : 04/01/2008
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'************************************************************************************
'************************************************************************************

Private Function PORT_FRONTIER_NORMAL_SIMULATION_FUNC(ByVal nLOOPS As Long, _
ByRef EXPECTED_RETURNS_RNG As Variant, _
ByRef VOLATILITIES_RNG As Variant, _
ByRef CORREL_RNG As Variant, _
Optional ByVal OUTPUT As Integer = 0)

Dim i As Long
Dim j As Long
Dim k As Long

Dim NROWS As Long
Dim NCOLUMNS As Long
Dim NO_ASSETS As Long

Dim TEMP_SUM As Double

Dim SIGMA_VECTOR As Variant
Dim CORREL_MATRIX As Variant
Dim EXPECTED_VECTOR As Variant

Dim TEMP_MATRIX As Variant
Dim CHOLESKY_MATRIX As Variant

Dim COVAR_MATRIX As Variant
Dim RETURNS_VECTOR As Variant

On Error GoTo ERROR_LABEL

EXPECTED_VECTOR = EXPECTED_RETURNS_RNG
If UBound(EXPECTED_VECTOR, 1) = 1 Then: _
    EXPECTED_VECTOR = MATRIX_TRANSPOSE_FUNC(EXPECTED_VECTOR)

SIGMA_VECTOR = VOLATILITIES_RNG
If UBound(SIGMA_VECTOR, 1) = 1 Then: _
    SIGMA_VECTOR = MATRIX_TRANSPOSE_FUNC(SIGMA_VECTOR)

NO_ASSETS = UBound(EXPECTED_VECTOR, 1)

CORREL_MATRIX = CORREL_RNG
If UBound(CORREL_MATRIX, 1) <> UBound(CORREL_MATRIX, 2) Then: GoTo ERROR_LABEL
If UBound(CORREL_MATRIX, 1) <> NO_ASSETS Then: GoTo ERROR_LABEL

NROWS = UBound(CORREL_MATRIX, 1)
NCOLUMNS = UBound(CORREL_MATRIX, 2)

ReDim TEMP_MATRIX(1 To NO_ASSETS, 1 To nLOOPS)
For i = 1 To NO_ASSETS
    For j = 1 To nLOOPS
        TEMP_MATRIX(i, j) = RANDOM_NORMAL_FUNC(0, 1, 0)
    Next j
Next i

ReDim CHOLESKY_MATRIX(1 To NROWS, 1 To NROWS)

For j = 1 To NROWS
    TEMP_SUM = 0
    For k = 1 To j - 1
        TEMP_SUM = TEMP_SUM + CHOLESKY_MATRIX(j, k) ^ 2
    Next k
    CHOLESKY_MATRIX(j, j) = CORREL_MATRIX(j, j) - TEMP_SUM
    ' Matrix is not semi-positive definite, no solution exists
    If CHOLESKY_MATRIX(j, j) < 0 Then: GoTo ERROR_LABEL
    CHOLESKY_MATRIX(j, j) = Sqr(MAXIMUM_FUNC(0, CHOLESKY_MATRIX(j, j)))
    
    For i = j + 1 To NROWS
        TEMP_SUM = 0
        For k = 1 To j - 1
            TEMP_SUM = TEMP_SUM + CHOLESKY_MATRIX(i, k) * CHOLESKY_MATRIX(j, k)
        Next k
        If CHOLESKY_MATRIX(j, j) = 0 Then
           CHOLESKY_MATRIX(i, j) = 0
        Else
           CHOLESKY_MATRIX(i, j) = (CORREL_MATRIX(i, j) - TEMP_SUM) / CHOLESKY_MATRIX(j, j)
        End If
    Next i
Next j


TEMP_MATRIX = MMULT_FUNC(CHOLESKY_MATRIX, TEMP_MATRIX, 70)

For i = 1 To NO_ASSETS
    For j = 1 To nLOOPS
        TEMP_MATRIX(i, j) = TEMP_MATRIX(i, j) * _
                            SIGMA_VECTOR(i, 1) + _
                            EXPECTED_VECTOR(i, 1)
    Next j
Next i

Select Case OUTPUT
Case 0
    PORT_FRONTIER_NORMAL_SIMULATION_FUNC = TEMP_MATRIX
Case 1
    TEMP_MATRIX = MATRIX_TRANSPOSE_FUNC(TEMP_MATRIX)
    NROWS = UBound(TEMP_MATRIX, 1)
    NCOLUMNS = UBound(TEMP_MATRIX, 2)
    GoSub 1983
    PORT_FRONTIER_NORMAL_SIMULATION_FUNC = RETURNS_VECTOR
Case 2
    TEMP_MATRIX = MATRIX_TRANSPOSE_FUNC(TEMP_MATRIX)
    NROWS = UBound(TEMP_MATRIX, 1)
    NCOLUMNS = UBound(TEMP_MATRIX, 2)
    GoSub 1983
    GoSub 1984
    PORT_FRONTIER_NORMAL_SIMULATION_FUNC = COVAR_MATRIX
Case Else
    ReDim TEMP_GROUP(1 To 3)
    TEMP_GROUP(1) = TEMP_MATRIX
    
    TEMP_MATRIX = MATRIX_TRANSPOSE_FUNC(TEMP_MATRIX)
    NROWS = UBound(TEMP_MATRIX, 1)
    NCOLUMNS = UBound(TEMP_MATRIX, 2)
    GoSub 1983
    GoSub 1984
    TEMP_GROUP(2) = RETURNS_VECTOR
    TEMP_GROUP(3) = COVAR_MATRIX
    PORT_FRONTIER_NORMAL_SIMULATION_FUNC = TEMP_GROUP
End Select

Exit Function
'-----------------------------------------------------------------------------
1983: 'compute means of data in matrix
'-----------------------------------------------------------------------------

ReDim RETURNS_VECTOR(1 To NCOLUMNS, 1 To 1)
For j = 1 To NCOLUMNS
    TEMP_SUM = 0
    For i = 1 To NROWS
        TEMP_SUM = TEMP_SUM + TEMP_MATRIX(i, j)
    Next i
    RETURNS_VECTOR(j, 1) = TEMP_SUM / NROWS
Next j
'-----------------------------------------------------------------------------
Return
'-----------------------------------------------------------------------------
1984: ' compute covariance matrix for data in matrix
'-----------------------------------------------------------------------------
ReDim COVAR_MATRIX(1 To NCOLUMNS, 1 To NCOLUMNS)
For j = 1 To NCOLUMNS
    For k = 1 To j
        TEMP_SUM = 0
        For i = 1 To NROWS
            TEMP_SUM = TEMP_SUM + (TEMP_MATRIX(i, j) - RETURNS_VECTOR(j, 1)) * (TEMP_MATRIX(i, k) - RETURNS_VECTOR(k, 1))
        Next i
        COVAR_MATRIX(j, k) = TEMP_SUM / NROWS
    Next k
Next j
'-----------------------------------------------------------------------------
For j = 1 To NCOLUMNS
    For k = j + 1 To NCOLUMNS
        COVAR_MATRIX(j, k) = COVAR_MATRIX(k, j)
    Next k
Next j
'-----------------------------------------------------------------------------
Return
'-----------------------------------------------------------------------------
ERROR_LABEL:
PORT_FRONTIER_NORMAL_SIMULATION_FUNC = Err.number
End Function
