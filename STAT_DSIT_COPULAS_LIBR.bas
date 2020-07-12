Attribute VB_Name = "STAT_DSIT_COPULAS_LIBR"


Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.

Private PUB_TARGET_VAL As Double
Private PUB_GUMBEL_VAL As Double
Private PUB_GUMBEL_ALPHA_VAL As Double
Private PUB_GUMBEL_TARGET_VAL As Double

'************************************************************************************
'************************************************************************************
'FUNCTION      : COPULA_RANDOM_FUNC

'DESCRIPTION   : Algorithm to simulate bivariate copulas (gaussian, student
't, clayton, rotated clayton, frank, gumbel, rotated gumbel) and calculating
'Kendall's Tau.

'LIBRARY       : STATISTICS
'GROUP         : COPULAS
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/08/2008
'************************************************************************************
'************************************************************************************

Function COPULA_RANDOM_FUNC(ByVal X_VAL As Double, _
Optional ByVal nDEGREES As Long = 1, _
Optional ByVal VERSION As Integer = 2, _
Optional ByVal RANDOM_TYPE As Integer = 0, _
Optional ByVal epsilon As Double = 0.0000000001)

Dim A_VAL As Double
Dim B_VAL As Double
Dim C_VAL As Double

Dim D_VAL As Double
Dim E_VAL As Double

Dim F_VAL As Double
Dim G_VAL As Double

Dim TEMP_VECTOR As Variant

On Error GoTo ERROR_LABEL

ReDim TEMP_VECTOR(1 To 1, 1 To 2)

'---------------------------------------------------------------------
Select Case VERSION
'---------------------------------------------------------------------
Case 0 'GaussianCopula --> PERFECT
'---------------------------------------------------------------------

    A_VAL = NORMSINV_FUNC(PSEUDO_RANDOM_FUNC(RANDOM_TYPE), 0, 1, 0)
    B_VAL = NORMSINV_FUNC(PSEUDO_RANDOM_FUNC(RANDOM_TYPE), 0, 1, 0)
    
    D_VAL = A_VAL
    E_VAL = X_VAL * A_VAL + Sqr(1 - X_VAL ^ 2) * B_VAL
    
    TEMP_VECTOR(1, 1) = NORMSDIST_FUNC(D_VAL, 0, 1, 0)
    TEMP_VECTOR(1, 2) = NORMSDIST_FUNC(E_VAL, 0, 1, 0)

'---------------------------------------------------------------------
Case 1 'ClaytonCopula --> PERFECT
'---------------------------------------------------------------------

    A_VAL = PSEUDO_RANDOM_FUNC(RANDOM_TYPE)
    B_VAL = PSEUDO_RANDOM_FUNC(RANDOM_TYPE)
    
    TEMP_VECTOR(1, 1) = A_VAL
    TEMP_VECTOR(1, 2) = ((A_VAL ^ (-X_VAL)) * (B_VAL ^ (-X_VAL / _
                        (X_VAL + 1)) - 1) + 1) ^ (-1 / X_VAL)


'---------------------------------------------------------------------
Case 2 'StudentTCopula --> PERFECT
'---------------------------------------------------------------------


    A_VAL = NORMSINV_FUNC(PSEUDO_RANDOM_FUNC(RANDOM_TYPE), 0, 1, 0)
    B_VAL = NORMSINV_FUNC(PSEUDO_RANDOM_FUNC(RANDOM_TYPE), 0, 1, 0)
    C_VAL = INVERSE_CHI_SQUARED_DIST_FUNC(PSEUDO_RANDOM_FUNC(RANDOM_TYPE), nDEGREES, False)
    
    D_VAL = A_VAL
    E_VAL = X_VAL * A_VAL + Sqr(1 - X_VAL ^ 2) * B_VAL
    
    F_VAL = D_VAL * Sqr(nDEGREES / C_VAL)
    G_VAL = E_VAL * Sqr(nDEGREES / C_VAL)
    
    If F_VAL > 0 Then
        TEMP_VECTOR(1, 1) = 1 - TDIST_FUNC(F_VAL, nDEGREES, True)
    Else
        TEMP_VECTOR(1, 1) = TDIST_FUNC(-F_VAL, nDEGREES, True)
    End If
        
    If G_VAL > 0 Then
        TEMP_VECTOR(1, 2) = 1 - TDIST_FUNC(G_VAL, nDEGREES, True)
    Else
        TEMP_VECTOR(1, 2) = TDIST_FUNC(-G_VAL, nDEGREES, True)
    End If


'---------------------------------------------------------------------
Case 3 'RotatedClaytonCopula --> PERFECT
'---------------------------------------------------------------------

    A_VAL = PSEUDO_RANDOM_FUNC(RANDOM_TYPE)
    B_VAL = PSEUDO_RANDOM_FUNC(RANDOM_TYPE)
    
    TEMP_VECTOR(1, 1) = 1 - A_VAL
    TEMP_VECTOR(1, 2) = 1 - ((A_VAL ^ (-X_VAL)) * (B_VAL ^ _
                    (-X_VAL / (X_VAL + 1)) - 1) + _
                    1) ^ (-1 / X_VAL)


'---------------------------------------------------------------------
Case 4 'FrankCopula --> PERFECT
'---------------------------------------------------------------------

    A_VAL = PSEUDO_RANDOM_FUNC(RANDOM_TYPE)
    B_VAL = PSEUDO_RANDOM_FUNC(RANDOM_TYPE)
    
    TEMP_VECTOR(1, 1) = A_VAL
    TEMP_VECTOR(1, 2) = (-1 / X_VAL) * Log(1 + (B_VAL * (1 - _
                        Exp(-X_VAL)) / (B_VAL * (Exp(-X_VAL * _
                        A_VAL) - 1) - Exp(-X_VAL * A_VAL))))


'---------------------------------------------------------------------
Case 5 'Generates random variates from the 2-dimensional Gumbel copula -- PERFECT
'---------------------------------------------------------------------

    A_VAL = PSEUDO_RANDOM_FUNC(RANDOM_TYPE)
    B_VAL = PSEUDO_RANDOM_FUNC(RANDOM_TYPE)
    
    GoSub CONVERGENCE
    C_VAL = D_VAL
    
    TEMP_VECTOR(1, 1) = Exp(-(A_VAL * (-Log(C_VAL)) ^ X_VAL) ^ (1 / X_VAL))
    TEMP_VECTOR(1, 2) = Exp(-((1 - A_VAL) * (-Log(C_VAL)) ^ X_VAL) ^ (1 / X_VAL))



'---------------------------------------------------------------------
Case Else 'RotatedGumbelCopula --> PERFECT
'---------------------------------------------------------------------

    A_VAL = PSEUDO_RANDOM_FUNC(RANDOM_TYPE)
    B_VAL = PSEUDO_RANDOM_FUNC(RANDOM_TYPE)
    
    GoSub CONVERGENCE
    C_VAL = D_VAL
    
    TEMP_VECTOR(1, 1) = 1 - Exp(-(A_VAL * (-Log(C_VAL)) ^ X_VAL) ^ (1 / X_VAL))
    TEMP_VECTOR(1, 2) = 1 - Exp(-((1 - A_VAL) * (-Log(C_VAL)) ^ X_VAL) ^ (1 / X_VAL))

'---------------------------------------------------------------------
End Select
'---------------------------------------------------------------------

COPULA_RANDOM_FUNC = TEMP_VECTOR

Exit Function
'----------------------------------------------------------------------
CONVERGENCE: 'Using Newton
'----------------------------------------------------------------------
D_VAL = epsilon
E_VAL = 0
Do While True
    F_VAL = -(Log(D_VAL) / X_VAL) - (1 / X_VAL) + 1
    G_VAL = (D_VAL - (D_VAL * Log(D_VAL) / X_VAL) - B_VAL) - E_VAL
    If Abs(G_VAL) < epsilon Then Exit Do
    D_VAL = D_VAL + (-G_VAL / F_VAL)
Loop
Return
'----------------------------------------------------------------------
ERROR_LABEL:
'----------------------------------------------------------------------
COPULA_RANDOM_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : COPULA_GUMBEL_FUNC
'DESCRIPTION   : COPULA_GUMBEL_FUNC
'LIBRARY       : STATISTICS
'GROUP         : COPULAS
'ID            : 002
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/08/2008
'************************************************************************************
'************************************************************************************

Function COPULA_GUMBEL_FUNC(ByVal ALPHA_VAL As Double, _
ByVal NROWS As Long, _
Optional ByVal RANDOM_TYPE As Integer = 0)
  
Dim i As Long
Dim ATEMP_VECTOR As Variant
Dim BTEMP_VECTOR As Variant 'p

Dim TEMP_MATRIX As Variant

On Error GoTo ERROR_LABEL

ReDim TEMP_MATRIX(1 To NROWS, 1 To 2)

ATEMP_VECTOR = MATRIX_RANDOM_UNIFORM_FUNC(NROWS, 1, RANDOM_TYPE, 0)
BTEMP_VECTOR = MATRIX_RANDOM_UNIFORM_FUNC(NROWS, 1, RANDOM_TYPE, 0)

PUB_GUMBEL_ALPHA_VAL = ALPHA_VAL

For i = 1 To UBound(ATEMP_VECTOR, 1)
  PUB_GUMBEL_VAL = ATEMP_VECTOR(i, 1)
  PUB_GUMBEL_TARGET_VAL = BTEMP_VECTOR(i, 1)
  
  TEMP_MATRIX(i, 1) = ATEMP_VECTOR(i, 1)
  TEMP_MATRIX(i, 2) = BRENT_ZERO_FUNC(0.00001, 0.9999, "COPULA_GUMBEL_COND_FUNC", 0.5, , , 100, 0.001)

Next i
  
COPULA_GUMBEL_FUNC = TEMP_MATRIX

Exit Function
ERROR_LABEL:
COPULA_GUMBEL_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : COPULA_GUMBEL_COND_FUNC
'DESCRIPTION   : Gumbel Condition CDF
'LIBRARY       : STATISTICS
'GROUP         : COPULAS
'ID            : 003
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/08/2008
'************************************************************************************
'************************************************************************************

Private Function COPULA_GUMBEL_COND_FUNC(ByVal B_VAL As Double)
  
  Dim A_VAL As Double
  Dim C_VAL As Double
  
  Dim MAX_LOG As Double
  Dim FIRST_LOG As Double
  Dim SECOND_LOG As Double
  
  Dim TEMP_LOG As Double
  Dim TEMP_ALPHA As Double
  
  On Error GoTo ERROR_LABEL
  
  A_VAL = PUB_GUMBEL_VAL
  TEMP_ALPHA = PUB_GUMBEL_ALPHA_VAL
  
  FIRST_LOG = -Log(A_VAL)
  SECOND_LOG = -Log(B_VAL)
  
  MAX_LOG = MAXIMUM_FUNC(FIRST_LOG, SECOND_LOG)
  
  TEMP_LOG = -MAX_LOG * ((FIRST_LOG / MAX_LOG) ^ TEMP_ALPHA + _
  (SECOND_LOG / MAX_LOG) ^ TEMP_ALPHA) ^ (1 / TEMP_ALPHA)
  C_VAL = (-FIRST_LOG / TEMP_LOG) ^ (TEMP_ALPHA - 1) * Exp(TEMP_LOG) / A_VAL
  
  COPULA_GUMBEL_COND_FUNC = C_VAL - PUB_GUMBEL_TARGET_VAL

Exit Function
ERROR_LABEL:
COPULA_GUMBEL_COND_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : COPULA_GAUSS_FUNC
'DESCRIPTION   : COPULA_GAUSS_FUNC
'LIBRARY       : STATISTICS
'GROUP         : COPULAS
'ID            : 004
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/08/2008
'************************************************************************************
'************************************************************************************

Function COPULA_GAUSS_FUNC(ByRef CORREL_RNG As Variant, _
Optional ByRef NROWS As Long = 0, _
Optional ByRef MEAN_RNG As Variant, _
Optional ByVal RANDOM_TYPE As Integer = 0)

  Dim i As Long
  Dim j As Long
  
  Dim NSIZE As Long
  Dim NCOLUMNS As Long

  Dim MEAN_VECTOR As Variant
  Dim CORREL_MATRIX As Variant
  Dim TEMP_MATRIX As Variant
  
  On Error GoTo ERROR_LABEL
  
  CORREL_MATRIX = CORREL_RNG
  If UBound(CORREL_MATRIX, 1) <> UBound(CORREL_MATRIX, 2) _
    Then: GoTo ERROR_LABEL 'CORREL_MATRIX must be square
  
  If NROWS = 0 Then: NROWS = UBound(CORREL_MATRIX, 2) _
  'The length of MEAN_VECTOR must equal the number of rows in CORREL_MATRIX

  If IsArray(MEAN_RNG) = False Then
      ReDim MEAN_VECTOR(1 To 1, 1 To UBound(CORREL_MATRIX, 2))
  Else
      MEAN_VECTOR = MEAN_RNG
      If UBound(MEAN_VECTOR, 2) <> UBound(CORREL_MATRIX, 2) _
        Then: GoTo ERROR_LABEL
  End If
  
  NCOLUMNS = UBound(MEAN_VECTOR, 2)
  NSIZE = MAXIMUM_FUNC(UBound(MEAN_VECTOR, 1), NCOLUMNS)
  
  If (UBound(MEAN_VECTOR, 1) * NCOLUMNS <> NSIZE) _
    Then: GoTo ERROR_LABEL 'MEAN_VECTOR must be a vector
  CORREL_MATRIX = MATRIX_CHOLESKY_FUNC(CORREL_MATRIX)
  
  If UBound(MEAN_VECTOR, 1) = NSIZE Then
    MEAN_VECTOR = MATRIX_TRANSPOSE_FUNC(MEAN_VECTOR)
    NCOLUMNS = UBound(MEAN_VECTOR, 2)
  End If
  
  TEMP_MATRIX = MATRIX_GENERATOR_FUNC(NROWS, NCOLUMNS, 1)
  
  For i = 1 To NROWS
    For j = 1 To NCOLUMNS
      TEMP_MATRIX(i, j) = MEAN_VECTOR(1, j)
    Next j
  Next i

  TEMP_MATRIX = (MATRIX_ELEMENTS_ADD_FUNC(MMULT_FUNC(MATRIX_RANDOM_UNIFORM_FUNC(NROWS, _
  NSIZE, RANDOM_TYPE, 0), MATRIX_TRANSPOSE_FUNC(CORREL_MATRIX), 70), TEMP_MATRIX))
  
  For i = 1 To NROWS
    For j = 1 To NCOLUMNS
      TEMP_MATRIX(i, j) = NORMSDIST_FUNC(TEMP_MATRIX(i, j), 0, 1, 0)
    Next j
  Next i

  COPULA_GAUSS_FUNC = TEMP_MATRIX
  
Exit Function
ERROR_LABEL:
COPULA_GAUSS_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : COPULA_T_FUNC
'DESCRIPTION   : COPULA_T_FUNC
'LIBRARY       : STATISTICS
'GROUP         : COPULAS
'ID            : 005
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/08/2008
'************************************************************************************
'************************************************************************************

Function COPULA_T_FUNC(ByRef CORREL_RNG As Variant, _
ByVal DEG_FREEDOM As Double, _
Optional ByRef NROWS As Long = 0, _
Optional ByRef MEAN_RNG As Variant, _
Optional ByVal RANDOM_TYPE As Integer = 0)

  Dim i As Long
  Dim j As Long
  
  Dim NSIZE As Long
  Dim NCOLUMNS As Long

  Dim TEMP_VAL As Double
  
  Dim ATEMP_MATRIX As Variant
  Dim BTEMP_MATRIX As Variant
  Dim CTEMP_MATRIX As Variant
  
  Dim MEAN_VECTOR As Variant
  Dim CORREL_MATRIX As Variant
  
  On Error GoTo ERROR_LABEL
  
  CORREL_MATRIX = CORREL_RNG
  If UBound(CORREL_MATRIX, 1) <> UBound(CORREL_MATRIX, 2) _
    Then: GoTo ERROR_LABEL 'CORREL_MATRIX must be square
  
  If NROWS = 0 Then: NROWS = UBound(CORREL_MATRIX, 2) _
  'The length of MEAN_VECTOR must equal the number of rows in CORREL_MATRIX

  If IsArray(MEAN_RNG) = False Then
      ReDim MEAN_VECTOR(1 To 1, 1 To UBound(CORREL_MATRIX, 2))
  Else
      MEAN_VECTOR = MEAN_RNG
      If UBound(MEAN_VECTOR, 2) <> UBound(CORREL_MATRIX, 2) _
        Then: GoTo ERROR_LABEL
  End If
  
  NCOLUMNS = UBound(MEAN_VECTOR, 2)
  NSIZE = MAXIMUM_FUNC(UBound(MEAN_VECTOR, 1), NCOLUMNS)
  
  If (UBound(MEAN_VECTOR, 1) * NCOLUMNS <> NSIZE) _
    Then: GoTo ERROR_LABEL 'MEAN_VECTOR must be a vector
  CORREL_MATRIX = MATRIX_CHOLESKY_FUNC(CORREL_MATRIX)
  
  If UBound(MEAN_VECTOR, 1) = NSIZE Then
    MEAN_VECTOR = MATRIX_TRANSPOSE_FUNC(MEAN_VECTOR)
    NCOLUMNS = UBound(MEAN_VECTOR, 2)
  End If
  
  ATEMP_MATRIX = MATRIX_GENERATOR_FUNC(NROWS, NCOLUMNS, 1)
  
  For i = 1 To NROWS
    For j = 1 To NCOLUMNS
      ATEMP_MATRIX(i, j) = MEAN_VECTOR(1, j)
    Next j
  Next i

  ATEMP_MATRIX = MMULT_FUNC(MATRIX_RANDOM_UNIFORM_FUNC(NROWS, _
  NSIZE, RANDOM_TYPE, 0), MATRIX_TRANSPOSE_FUNC(CORREL_MATRIX), 70)

  ReDim BTEMP_MATRIX(1 To NROWS, 1 To 1)
  For i = 1 To NROWS
      BTEMP_MATRIX(i, 1) = INVERSE_GAMMA_DIST_FUNC(PSEUDO_RANDOM_FUNC(RANDOM_TYPE), _
      DEG_FREEDOM / 2, 2, True)
  Next i

  ReDim CTEMP_MATRIX(1 To NROWS, 1 To NSIZE)
  
  For i = 1 To NROWS
    TEMP_VAL = (BTEMP_MATRIX(i, 1) / DEG_FREEDOM) ^ 0.5
    For j = 1 To NSIZE
      CTEMP_MATRIX(i, j) = ATEMP_MATRIX(i, j) / TEMP_VAL
    Next j
  Next i
'-------------------------------T-Transformation-----------------------------
  For i = 1 To NROWS
    For j = 1 To NSIZE
        TEMP_VAL = CTEMP_MATRIX(i, j)
            If TEMP_VAL > 0 Then
                  CTEMP_MATRIX(i, j) = 1 - BETA_DIST_FUNC(DEG_FREEDOM / _
                  (DEG_FREEDOM + TEMP_VAL ^ 2), DEG_FREEDOM / _
                  2, 1 / 2, True, True) / 2
            Else
                  CTEMP_MATRIX(i, j) = BETA_DIST_FUNC(DEG_FREEDOM / _
                  (DEG_FREEDOM + TEMP_VAL ^ 2), DEG_FREEDOM / _
                  2, 1 / 2, True, True) / 2
            End If
    Next j
  Next i
'---------------------------------------------------------------------------
  'For the inverse of the t-dist;  where n = DF
    'TEMP_VAL = BETA_INV(2 * MINIMUM_FUNC(x, 1 - x), n / 2, 1 / 2)
    'TEMP_VAL = 1 / TEMP_VAL
    'TINV = (Sgn(x - 0.5) * (n * (TEMP_VAL - 1)) ^ 0.5)
'---------------------------------------------------------------------------
  COPULA_T_FUNC = CTEMP_MATRIX
  
Exit Function
ERROR_LABEL:
COPULA_T_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : COPULA_CLAYTON_FUNC
'DESCRIPTION   :
'LIBRARY       : STATISTICS
'GROUP         : COPULAS
'ID            : 006
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/08/2008
'************************************************************************************
'************************************************************************************

Function COPULA_CLAYTON_FUNC(ByVal ALPHA_VAL As Double, _
ByVal NROWS As Long, _
Optional ByVal RANDOM_TYPE As Integer = 0)
  
  Dim i As Long
  
  Dim ATEMP_VECTOR As Variant
  Dim BTEMP_VECTOR As Variant
  
  Dim TEMP_MATRIX As Variant

  On Error GoTo ERROR_LABEL
    
  ReDim TEMP_MATRIX(1 To NROWS, 1 To 2)
  
  ATEMP_VECTOR = MATRIX_RANDOM_UNIFORM_FUNC(NROWS, 1, RANDOM_TYPE, 0)
  BTEMP_VECTOR = MATRIX_RANDOM_UNIFORM_FUNC(NROWS, 1, RANDOM_TYPE, 0)
  
  For i = 1 To NROWS
    TEMP_MATRIX(i, 1) = ATEMP_VECTOR(i, 1)
    TEMP_MATRIX(i, 2) = ATEMP_VECTOR(i, 1) * (BTEMP_VECTOR(i, 1) ^ _
        (-ALPHA_VAL / (1 + ALPHA_VAL)) - 1 + _
        ATEMP_VECTOR(i, 1) ^ ALPHA_VAL) ^ (-1 / ALPHA_VAL)
  Next i
  
  COPULA_CLAYTON_FUNC = TEMP_MATRIX

Exit Function
ERROR_LABEL:
COPULA_CLAYTON_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : COPULA_KENDALL_FUNC
'DESCRIPTION   : COPULA_KENDALL_FUNC
'LIBRARY       : STATISTICS
'GROUP         : COPULAS
'ID            : 007
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/08/2008
'************************************************************************************
'************************************************************************************

Function COPULA_KENDALL_FUNC(ByVal RHO_VAL As Double, _
Optional ByVal VERSION As Integer = 0)
  
Dim PI_VAL As Double
Dim TEMP_VAL As Double
On Error GoTo ERROR_LABEL

PI_VAL = 3.14159265358979

Select Case VERSION
Case 0 'Gaussin or T-Student
    COPULA_KENDALL_FUNC = 2 * (2 * Atn(RHO_VAL / (1 + Sqr(1 - RHO_VAL * RHO_VAL)))) / PI_VAL _
    '--> Same as  2 * Excel.Application.Asin(RHO_VAL) / Excel.Application.Pi()
    '--> To get the Rho: Sin(Tau * Excel.Application.pi() / 2)
Case 1 'Clayton
    COPULA_KENDALL_FUNC = RHO_VAL / (2 + RHO_VAL)
Case Else 'Frank
    TEMP_VAL = GAULEG7_INTEGRATION_FUNC("COPULA_FRANK_INTEGRAND_FUNC", 0, RHO_VAL, 0.01, 100) / Abs(RHO_VAL)
    TEMP_VAL = Abs(TEMP_VAL)
    COPULA_KENDALL_FUNC = 1 + 4 * (TEMP_VAL - 1) / RHO_VAL
End Select

Exit Function
ERROR_LABEL:
COPULA_KENDALL_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : COPULA_FRANK_INTEGRAND_FUNC
'DESCRIPTION   :
'LIBRARY       : STATISTICS
'GROUP         : COPULAS
'ID            : 008
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/08/2008
'************************************************************************************
'************************************************************************************

Private Function COPULA_FRANK_INTEGRAND_FUNC(ByVal X_VAL As Double)
On Error GoTo ERROR_LABEL
COPULA_FRANK_INTEGRAND_FUNC = X_VAL / (Exp(X_VAL) - 1)
Exit Function
ERROR_LABEL:
COPULA_FRANK_INTEGRAND_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : COPULA_FRANK_FUNC
'DESCRIPTION   :
'LIBRARY       : STATISTICS
'GROUP         : COPULAS
'ID            : 009
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/08/2008
'************************************************************************************
'************************************************************************************

Function COPULA_FRANK_FUNC(ByVal ALPHA_VAL As Double, _
ByVal NROWS As Long, _
Optional ByVal RANDOM_TYPE As Integer = 0)

  Dim i As Long
  
  Dim ATEMP_VECTOR As Variant
  Dim BTEMP_VECTOR As Variant
  
  Dim TEMP_MATRIX As Variant

  On Error GoTo ERROR_LABEL
    
  ReDim TEMP_MATRIX(1 To NROWS, 1 To 2)
  
  ATEMP_VECTOR = MATRIX_RANDOM_UNIFORM_FUNC(NROWS, 1, RANDOM_TYPE, 0)
  BTEMP_VECTOR = MATRIX_RANDOM_UNIFORM_FUNC(NROWS, 1, RANDOM_TYPE, 0)
  
  For i = 1 To NROWS
    TEMP_MATRIX(i, 1) = ATEMP_VECTOR(i, 1)
    TEMP_MATRIX(i, 2) = -1 * Log((Exp(-ALPHA_VAL * ATEMP_VECTOR(i, 1)) * _
        (1 - BTEMP_VECTOR(i, 1)) / BTEMP_VECTOR(i, 1) + Exp(-ALPHA_VAL)) / _
        (1 + Exp(-ALPHA_VAL * ATEMP_VECTOR(i, 1)) * (1 - BTEMP_VECTOR(i, 1)) / _
        BTEMP_VECTOR(i, 1))) / ALPHA_VAL
  Next i
  
  COPULA_FRANK_FUNC = TEMP_MATRIX

Exit Function
ERROR_LABEL:
COPULA_FRANK_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : COPULA_ALPHA_FUNC
'DESCRIPTION   :
'LIBRARY       : STATISTICS
'GROUP         : COPULAS
'ID            : 010
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/08/2008
'************************************************************************************
'************************************************************************************

Function COPULA_ALPHA_FUNC(ByVal TAU_VAL As Double, _
Optional ByVal VERSION As Integer = 0)

On Error GoTo ERROR_LABEL

Select Case VERSION
Case 0 'Clayton
    COPULA_ALPHA_FUNC = 2 * TAU_VAL / (1 - TAU_VAL)
Case 1 'Frank
    PUB_TARGET_VAL = TAU_VAL
    COPULA_ALPHA_FUNC = BRENT_ZERO_FUNC(-100, 100, "COPULA_FRANK_ROOT_FUNC", 0.5, , , 100, 0.01)
Case Else 'Gumbel
    COPULA_ALPHA_FUNC = 1 / (1 - TAU_VAL)
End Select

Exit Function
ERROR_LABEL:
COPULA_ALPHA_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : COPULA_FRANK_ROOT_FUNC
'DESCRIPTION   :
'LIBRARY       : STATISTICS
'GROUP         : COPULAS
'ID            : 011
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/08/2008
'************************************************************************************
'************************************************************************************

Private Function COPULA_FRANK_ROOT_FUNC(ByVal ALPHA As Double)
  
Dim TAU_VAL As Double
Dim TEMP_VAL As Double

On Error GoTo ERROR_LABEL

If Abs(ALPHA) < 2.22E-16 Then
    TAU_VAL = 0
Else
    TEMP_VAL = GAULEG7_INTEGRATION_FUNC("COPULA_FRANK_INTEGRAND_FUNC", 0, ALPHA, 0.01, 100) / Abs(ALPHA)
    TEMP_VAL = Abs(TEMP_VAL)
    TAU_VAL = 1 + 4 * (TEMP_VAL - 1) / ALPHA
End If

COPULA_FRANK_ROOT_FUNC = TAU_VAL - PUB_TARGET_VAL

Exit Function
ERROR_LABEL:
COPULA_FRANK_ROOT_FUNC = Err.number
End Function
