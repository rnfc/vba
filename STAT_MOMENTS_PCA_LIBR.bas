Attribute VB_Name = "STAT_MOMENTS_PCA_LIBR"


Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.

'************************************************************************************
'************************************************************************************
'FUNCTION      : MATRIX_PCA_FACTORS_FUNC

'DESCRIPTION   : The Factor-Based Scenario Method provides the user with
'several advantages relative to alternative calculation strategies. First, it
'includes options in its analysis without modifications. By contrast, options
'can provoke extremely misleading results in VAR calculations that linearize
'profit as a function of market variables. Second, the Scenario Method produces
'estimates of risk quickly, unlike methods that depend on extensive simulation.
'A VAR estimate has little value unless it arrives in time to allow adjustments to
'the portfolio. Third, the method provides a useful and easily understood
'summary of the qualitative nature of the risk facing a portfolio. One can
'immediately see that the portfolio is sensitive, say, to rising interest
'rates or to a steepening yield curve. Fourth, the method
'identifies whether an additional trade will increase or decrease risk,
'which provides insight into hedge strategies. Finally, the method allows
'the straightforward aggregation of risks across portfolios maintained and
'valued on different computer systems. The response to a given scenario of
'the combined portfolio is simply the sum of the responses of the individual
'portfolios.

'This routine calculates:
 'a covariance matrix of a set of data
'the eigen values associated with that covariance matrix
'the percent of variance explained by each eigen vector
'the eigen vector associated with each eigenvalue.
'the three largest eigen vectors
    
'LIBRARY       : STATISTICS
'GROUP         : PCA
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/21/2009
'************************************************************************************
'************************************************************************************

Function MATRIX_PCA_FACTORS_FUNC(ByRef DATA_RNG As Variant, _
Optional ByVal DATA_TYPE As Integer = 0, _
Optional ByVal LOG_SCALE As Integer = 0, _
Optional ByVal OUTPUT As Integer = 0)

'The Factor-Based approach to calculating VAR begins with a principal
'components analysis of the yield curve. This decomposes yield curve
'movements into a small number of underlying factors including a “Shift”
'factor that allows rates to rise or fall and a “Twist” factor that allows
'the curve to steepen or flatten. Combining these factors produces specific
'yield curve scenarios used to estimate hypothetical portfolio profit or
'loss. The greatest loss among these scenarios provides an intuitive and
'rapid VAR estimate that tends to provide a conservative estimate of
'the nominal percentile of the loss distribution.

'The Factor-Based Scenario Method is not foolproof, however, and the user
'must judge the appropriateness of the technique for the portfolio at hand.
'Nonetheless the method works well for a class of portfolios that is quite
'important in practice— portfolios that display a concave response to
'changing market prices, such as portfolios dominated by short positions
'in standard options. The negative gamma of such positions
'creates particular concern for risk control purposes.

Dim i As Long
Dim j As Long

Dim NSIZE As Long

Dim TEMP_SUM As Double
Dim DATA_MATRIX As Variant

Dim EIGEN_VECTOR As Variant
Dim EIGEN_MATRIX As Variant

Dim RANK_VECTOR As Variant
Dim RANK_MATRIX As Variant

Dim TEMP_MATRIX As Variant
Dim TEMP_VECTOR As Variant

On Error GoTo ERROR_LABEL

DATA_MATRIX = MATRIX_COVARIANCE_FRAME1_FUNC(DATA_RNG, DATA_TYPE, LOG_SCALE)
TEMP_MATRIX = MATRIX_PCA_FUNC(DATA_MATRIX, False, 2)
EIGEN_MATRIX = TEMP_MATRIX(LBound(TEMP_MATRIX))
EIGEN_VECTOR = TEMP_MATRIX(UBound(TEMP_MATRIX))
Erase TEMP_MATRIX

'----------------------------------------------------------------------------------------
'----------------------------------------------------------------------------------------

'Principal Components Analysis (PCA) can be used to fit a linear regression
'that minimizes the perpendicular distances from the data to the fitted model.
'This is the linear case of what is known as Orthogonal Regression or Total
'Least Squares, and is appropriate when there is no natural distinction between
'predictor and response variables, or when all variables are measured with error.
'This is in contrast to the usual regression assumption that predictor variables
'are measured exactly, and only the response variable has an error component.

'For example, given two data vectors x and y, you can fit a line that minimizes
'the perpendicular distances from each of the points (x(i), y(i)) to the line.
'More generally, with p observed variables, you can fit an r-dimensional
'hyperplane in p-dimensional space (r < p). The choice of r is equivalent to
'choosing the number of components to retain in PCA. It may be based on
'prediction error, or it may simply be a pragmatic choice to reduce data to a
'manageable number of dimensions.
'----------------------------------------------------------------------------------------


NSIZE = UBound(EIGEN_MATRIX, 1)

'----------------------------SETTING HEADINGS-----------------------------
ReDim DATA_MATRIX(1 To NSIZE + 1, 1 To NSIZE)
For j = 1 To NSIZE
    DATA_MATRIX(1, j) = "FACTORS: " & CStr(j)
Next j
'-------------------------------------------------------------------------

DATA_MATRIX(1, 1) = "SHIFT"
DATA_MATRIX(1, 2) = "TWIST"
DATA_MATRIX(1, 3) = "BOW"
DATA_MATRIX(1, 4) = "BOW2"

For j = 1 To NSIZE
    For i = 1 To NSIZE
        DATA_MATRIX(i + 1, j) = EIGEN_VECTOR(j, 1) ^ 0.5 * EIGEN_MATRIX(i, j)
    Next i
Next j

'-----------------------------------------------------------------------------
TEMP_SUM = MATRIX_ELEMENTS_CUMULATIVE_SUM_FUNC(EIGEN_VECTOR)
TEMP_VECTOR = MATRIX_QUICK_SORT_FUNC(EIGEN_VECTOR, 1, 1)
    
ReDim RANK_VECTOR(1 To NSIZE, 1 To 1)

For j = 1 To NSIZE
    For i = 1 To NSIZE
        If TEMP_VECTOR(j, 1) = EIGEN_VECTOR(i, 1) Then
            RANK_VECTOR(j, 1) = i
            Exit For
        End If
    Next i
Next j

'------------------------------------------------------------------------
ReDim TEMP_MATRIX(1 To (NSIZE + 1), 1 To 2)

TEMP_MATRIX(1, 1) = "SORTED EIGEN VALUES"
TEMP_MATRIX(1, 2) = "% VARIANCE EXPLAINED"

For i = 2 To (NSIZE + 1)
    TEMP_MATRIX(i, 1) = TEMP_VECTOR(i - 1, 1)
    TEMP_MATRIX(i, 2) = TEMP_MATRIX(i, 1) / TEMP_SUM
Next i

'--------------------------SETTING HEADINGS-------------------------------

ReDim RANK_MATRIX(1 To NSIZE + 1, 1 To NSIZE)
For j = 1 To NSIZE
    RANK_MATRIX(1, j) = "FACTOR: " & CStr(j)
Next j

'-------------------------------------------------------------------------

For j = 1 To NSIZE
    For i = 1 To NSIZE
        RANK_MATRIX(i + 1, j) = EIGEN_MATRIX(i, RANK_VECTOR(j, 1))
    Next i
Next j

'------------------------------------------------------------------------

Select Case OUTPUT
    Case 0 'Factor Table
        MATRIX_PCA_FACTORS_FUNC = DATA_MATRIX
    Case 1 'Largest Eigen Vectors
        MATRIX_PCA_FACTORS_FUNC = RANK_MATRIX
    Case 2 'Sorted Eigen Values
        MATRIX_PCA_FACTORS_FUNC = TEMP_MATRIX
    Case 3 'Eigen Values
        MATRIX_PCA_FACTORS_FUNC = EIGEN_VECTOR
    Case 4 'Eigen Vectors
        MATRIX_PCA_FACTORS_FUNC = EIGEN_MATRIX
    Case Else
        MATRIX_PCA_FACTORS_FUNC = Array(DATA_MATRIX, RANK_MATRIX, TEMP_MATRIX, EIGEN_VECTOR, EIGEN_MATRIX)
End Select

Exit Function
ERROR_LABEL:
MATRIX_PCA_FACTORS_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : MATRIX_PCA_FUNC

'DESCRIPTION   : This function performs the Principal Component analysis: tridiagonal
'reductional form. Principal Components Analysis (PCA) can be used to fit a linear
'regression that minimizes the perpendicular distances from the data to the fitted
'model. This is the linear case of what is known as Orthogonal Regression or Total
'Least Squares, and is appropriate when there is no natural distinction between
'predictor and response variables, or when all variables are measured with error.
'This is in contrast to the usual regression assumption that predictor variables
'are measured exactly, and only the response variable has an error component.

'For example, given two data vectors x and y, you can fit a line that minimizes the
'perpendicular distances from each of the points (x(i), y(i)) to the line. More
'generally, with p observed variables, you can fit an r-dimensional hyperplane
'in p-dimensional space (r < p). The choice of r is equivalent to choosing the
'number of components to retain in PCA. It may be based on prediction error, or
'it may simply be a pragmatic choice to reduce data to a manageable number of
'dimensions.

'LIBRARY       : STATISTICS
'GROUP         : PCA
'ID            : 002
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/21/2009
'************************************************************************************
'************************************************************************************

Function MATRIX_PCA_FUNC(ByRef DATA_RNG As Variant, _
Optional ByVal INDEXING_FLAG As Boolean = True, _
Optional ByVal OUTPUT As Integer = 0)

Dim h As Long
Dim i As Long
Dim j As Long
Dim k As Long
Dim l As Long
Dim n As Long
Dim NSIZE As Long

Dim D_VAL As Double
Dim B_VAL As Double
Dim S_VAL As Double
Dim R_VAL As Double
Dim C_VAL As Double
Dim P_VAL As Double

Dim H_VAL As Double
Dim I_VAL As Double
Dim F_VAL As Double
Dim G_VAL As Double

Dim TEMP_SUM As Double

Dim DATA_MATRIX As Variant
Dim TEMP_MATRIX As Variant
Dim TEMP_VECTOR As Variant

Dim INDEX_VECTOR As Variant

On Error GoTo ERROR_LABEL

'---------------------------------------------------------------------------------
'Any continuous function can be approximated arbitrarily well by a piecewise linear
'function. As the number of regularly spaced intervals increases, the piecewise linear
'function grows closer to the continuous function it approximates. As a practical
'matter, one needs only a few intervals to approximate the value of non-exotic
'derivatives. (Piecewise approximation may not be appropriate if the portfolio has
'considerable exposure to deals such as digital options.)
'----------------------------------------------------------------------------------

DATA_MATRIX = DATA_RNG
NSIZE = UBound(DATA_MATRIX, 1)

ReDim TEMP_MATRIX(1 To NSIZE, 1 To 1)
ReDim TEMP_VECTOR(1 To NSIZE, 1 To 1)

For i = NSIZE To 2 Step -1
l = i - 1
H_VAL = 0
TEMP_SUM = 0
If l > 1 Then
  For k = 1 To l
    TEMP_SUM = TEMP_SUM + Abs(DATA_MATRIX(i, k))
  Next k
  If TEMP_SUM = 0 Then
    TEMP_VECTOR(i, 1) = DATA_MATRIX(i, l)
  Else
    For k = 1 To l
      DATA_MATRIX(i, k) = DATA_MATRIX(i, k) / TEMP_SUM
      H_VAL = H_VAL + DATA_MATRIX(i, k) ^ 2
    Next k
    F_VAL = DATA_MATRIX(i, l)
    G_VAL = -NEGATIVE_THRESHOLD_FUNC(Sqr(H_VAL), F_VAL)
    TEMP_VECTOR(i, 1) = TEMP_SUM * G_VAL
    H_VAL = H_VAL - F_VAL * G_VAL
    DATA_MATRIX(i, l) = F_VAL - G_VAL
    F_VAL = 0
    For j = 1 To l
      DATA_MATRIX(j, i) = DATA_MATRIX(i, j) / H_VAL
      G_VAL = 0
      For k = 1 To j
        G_VAL = G_VAL + DATA_MATRIX(j, k) * DATA_MATRIX(i, k)
      Next k
      For k = j + 1 To l
        G_VAL = G_VAL + DATA_MATRIX(k, j) * DATA_MATRIX(i, k)
      Next k
      TEMP_VECTOR(j, 1) = G_VAL / H_VAL
      F_VAL = F_VAL + TEMP_VECTOR(j, 1) * DATA_MATRIX(i, j)
    Next j
    I_VAL = F_VAL / (H_VAL + H_VAL)
    For j = 1 To l
      F_VAL = DATA_MATRIX(i, j)
      G_VAL = TEMP_VECTOR(j, 1) - I_VAL * F_VAL
      TEMP_VECTOR(j, 1) = G_VAL
      For k = 1 To j
        DATA_MATRIX(j, k) = DATA_MATRIX(j, k) - F_VAL * TEMP_VECTOR(k, 1) - G_VAL * _
        DATA_MATRIX(i, k)
      Next k
    Next j
  End If
Else
  TEMP_VECTOR(i, 1) = DATA_MATRIX(i, l)
End If
TEMP_MATRIX(i, 1) = H_VAL
Next i
TEMP_MATRIX(1, 1) = 0
TEMP_VECTOR(1, 1) = 0
For i = 1 To NSIZE
l = i - 1
If TEMP_MATRIX(i, 1) <> 0 Then
  For j = 1 To l
    G_VAL = 0
    For k = 1 To l
      G_VAL = G_VAL + DATA_MATRIX(i, k) * DATA_MATRIX(k, j)
    Next k
    For k = 1 To l
      DATA_MATRIX(k, j) = DATA_MATRIX(k, j) - G_VAL * DATA_MATRIX(k, i)
    Next k
  Next j
End If
TEMP_MATRIX(i, 1) = DATA_MATRIX(i, i)
DATA_MATRIX(i, i) = 1
For j = 1 To l
  DATA_MATRIX(i, j) = 0
  DATA_MATRIX(j, i) = 0
Next j
Next i

'-------------------Find evalues and evector of tridiagonal matrix------------------

For i = 2 To NSIZE
  TEMP_VECTOR(i - 1, 1) = TEMP_VECTOR(i, 1)
Next i

TEMP_VECTOR(NSIZE, 1) = 0

For l = 1 To NSIZE
  n = 0
1983:
  For h = l To NSIZE - 1
    D_VAL = Abs(TEMP_MATRIX(h, 1)) + Abs(TEMP_MATRIX(h + 1, 1))
    If Abs(TEMP_VECTOR(h, 1)) + D_VAL = D_VAL Then GoTo 1984
  Next h
  h = NSIZE
  
1984:
    If h <> l Then
      If n = 30 Then: GoTo ERROR_LABEL
      n = n + 1
      
      G_VAL = (TEMP_MATRIX(l + 1, 1) - TEMP_MATRIX(l, 1)) / (2 * TEMP_VECTOR(l, 1))
      R_VAL = PYTHAG_FUNC(G_VAL, 1)
      G_VAL = TEMP_MATRIX(h, 1) - TEMP_MATRIX(l, 1) + TEMP_VECTOR(l, 1) / _
          (G_VAL + NEGATIVE_THRESHOLD_FUNC(R_VAL, G_VAL))
      S_VAL = 1
      C_VAL = 1
      P_VAL = 0
      
      For i = h - 1 To l Step -1
        F_VAL = S_VAL * TEMP_VECTOR(i, 1)
        B_VAL = C_VAL * TEMP_VECTOR(i, 1)
        R_VAL = PYTHAG_FUNC(F_VAL, G_VAL)
        TEMP_VECTOR(i + 1, 1) = R_VAL
        If R_VAL = 0 Then
          TEMP_MATRIX(i + 1, 1) = TEMP_MATRIX(i + 1, 1) - P_VAL
          TEMP_VECTOR(h, 1) = 0
          GoTo 1983
        End If
        S_VAL = F_VAL / R_VAL
        C_VAL = G_VAL / R_VAL
        G_VAL = TEMP_MATRIX(i + 1, 1) - P_VAL
        R_VAL = (TEMP_MATRIX(i, 1) - G_VAL) * S_VAL + 2 * C_VAL * B_VAL
        P_VAL = S_VAL * R_VAL
        TEMP_MATRIX(i + 1, 1) = G_VAL + P_VAL
        G_VAL = C_VAL * R_VAL - B_VAL
        For k = 1 To NSIZE
          F_VAL = DATA_MATRIX(k, i + 1)
          DATA_MATRIX(k, i + 1) = S_VAL * DATA_MATRIX(k, i) + C_VAL * F_VAL
          DATA_MATRIX(k, i) = C_VAL * DATA_MATRIX(k, i) - S_VAL * F_VAL
        Next k
      Next i
       TEMP_MATRIX(l, 1) = TEMP_MATRIX(l, 1) - P_VAL
      TEMP_VECTOR(l, 1) = G_VAL
      TEMP_VECTOR(h, 1) = 0
      GoTo 1983
    End If
Next l

'---------------------------------------------------------------
Select Case INDEXING_FLAG
'---------------------------------------------------------------
Case Is = True 'indexing is true
'---------------------------------------------------------------
    INDEX_VECTOR = MATRIX_PCA_INDEXING_FUNC(TEMP_MATRIX)
    ReDim TEMP_VECTOR(1 To NSIZE, 1 To 1)
    ReDim TEMP_MATRIX(1 To NSIZE, 1 To NSIZE)
    For i = 1 To NSIZE
        For j = 1 To NSIZE
            TEMP_MATRIX(i, j) = DATA_MATRIX(i, INDEX_VECTOR(NSIZE + 1 - j, 1))
            TEMP_VECTOR(i, 1) = TEMP_MATRIX(INDEX_VECTOR(NSIZE + 1 - i, 1), 1)
        Next j
    Next i
            
    Select Case OUTPUT
    Case 0 'EigenVectors
        MATRIX_PCA_FUNC = TEMP_MATRIX
    Case 1 'EigenValues
        MATRIX_PCA_FUNC = TEMP_VECTOR
    Case Else
        MATRIX_PCA_FUNC = Array(TEMP_MATRIX, TEMP_VECTOR)
    End Select
'---------------------------------------------------------------
Case Else
'---------------------------------------------------------------
    Select Case OUTPUT
    Case 0 'EigenVectors
        MATRIX_PCA_FUNC = DATA_MATRIX
    Case 1 'EigenValues
        MATRIX_PCA_FUNC = TEMP_MATRIX
    Case Else
        MATRIX_PCA_FUNC = Array(DATA_MATRIX, TEMP_MATRIX)
    End Select
'---------------------------------------------------------------
End Select
'---------------------------------------------------------------

Exit Function
ERROR_LABEL:
MATRIX_PCA_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : MATRIX_PCA_INDEXING_FUNC
'DESCRIPTION   : This function sorts the EigenVectors
'LIBRARY       : STATISTICS
'GROUP         : PCA
'ID            : 003
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/21/2009
'************************************************************************************
'************************************************************************************

Private Function MATRIX_PCA_INDEXING_FUNC(ByRef DATA_VECTOR As Variant)

Dim h As Long
Dim i As Long
Dim j As Long
Dim k As Long
Dim l As Long
Dim m As Long
Dim n As Long
Dim o As Long

Dim NROWS As Long
Dim TEMP_VAL As Double
Dim INDEX_VECTOR As Variant
Dim STOCK_VECTOR As Variant

On Error GoTo ERROR_LABEL

NROWS = UBound(DATA_VECTOR, 1)

ReDim STOCK_VECTOR(1 To NROWS, 1 To 1)
ReDim INDEX_VECTOR(1 To NROWS, 1 To 1)

For j = 1 To NROWS
  INDEX_VECTOR(j, 1) = j
Next j

o = 0
l = 1
n = NROWS

1983: If n - l < 7 Then
      For j = l + 1 To n
        h = INDEX_VECTOR(j, 1)
        TEMP_VAL = DATA_VECTOR(h, 1)
        For i = j - 1 To 1 Step -1
          If DATA_VECTOR(INDEX_VECTOR(i, 1), 1) <= TEMP_VAL Then GoTo 1984
          INDEX_VECTOR(i + 1, 1) = INDEX_VECTOR(i, 1)
        Next i
        i = 0
1984:    INDEX_VECTOR(i + 1, 1) = h
    Next j
      If o = 0 Then: GoTo 1988
      n = STOCK_VECTOR(o, 1)
      l = STOCK_VECTOR(o - 1, 1)
      o = o - 2
    Else
      k = (l + n) / 2
      m = INDEX_VECTOR(k, 1)
      INDEX_VECTOR(k, 1) = INDEX_VECTOR(l + 1, 1)
      INDEX_VECTOR(l + 1, 1) = m
      If DATA_VECTOR(INDEX_VECTOR(l + 1, 1), 1) > DATA_VECTOR(INDEX_VECTOR(n, 1), 1) Then
        m = INDEX_VECTOR(l + 1, 1)
        INDEX_VECTOR(l + 1, 1) = INDEX_VECTOR(n, 1)
        INDEX_VECTOR(n, 1) = m
      End If
      
      If DATA_VECTOR(INDEX_VECTOR(l, 1), 1) > DATA_VECTOR(INDEX_VECTOR(n, 1), 1) Then
        m = INDEX_VECTOR(l, 1)
        INDEX_VECTOR(l, 1) = INDEX_VECTOR(n, 1)
        INDEX_VECTOR(n, 1) = m
      End If
      If DATA_VECTOR(INDEX_VECTOR(l + 1, 1), 1) > DATA_VECTOR(INDEX_VECTOR(l, 1), 1) Then
        m = INDEX_VECTOR(l + 1, 1)
        INDEX_VECTOR(l + 1, 1) = INDEX_VECTOR(l, 1)
        INDEX_VECTOR(l, 1) = m
      End If
      i = l + 1
      j = n
      h = INDEX_VECTOR(l, 1)
      TEMP_VAL = DATA_VECTOR(h, 1)
1985:
      i = i + 1
      If DATA_VECTOR(INDEX_VECTOR(i, 1), 1) < TEMP_VAL Then GoTo 1985
1986:
      j = j - 1
      If DATA_VECTOR(INDEX_VECTOR(j, 1), 1) > TEMP_VAL Then GoTo 1986
      If j < i Then GoTo 1987
          m = INDEX_VECTOR(i, 1)
          INDEX_VECTOR(i, 1) = INDEX_VECTOR(j, 1)
          INDEX_VECTOR(j, 1) = m
      GoTo 1985

1987:  INDEX_VECTOR(l, 1) = INDEX_VECTOR(j, 1)
      INDEX_VECTOR(j, 1) = h
      o = o + 2
      
      If n - i + 1 >= j - l Then
        STOCK_VECTOR(o, 1) = n
        STOCK_VECTOR(o - 1, 1) = i
        n = j - 1
      Else
        STOCK_VECTOR(o, 1) = j - 1
        STOCK_VECTOR(o - 1, 1) = l
        l = i
      End If
    End If
    GoTo 1983
      
1988: MATRIX_PCA_INDEXING_FUNC = INDEX_VECTOR

Exit Function
ERROR_LABEL:
MATRIX_PCA_INDEXING_FUNC = Err.number
End Function
