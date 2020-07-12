Attribute VB_Name = "FINAN_FI_HJM_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.

'************************************************************************************
'************************************************************************************
'FUNCTION      : HJM_BUSHY_TREE_FUNC
'DESCRIPTION   : Non-recombining tree for HJM: This program evolves forward rate
'trees on Bushy tree for single factor HJM model as described in
'Robert J. Arrow's book on Interest rate modelling
'LIBRARY       : FIXED_INCOME
'GROUP         : HJM
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 10-06-2008
'************************************************************************************
'************************************************************************************

Function HJM_BUSHY_TREE_FUNC(ByRef FORW_RNG As Variant, _
ByRef SIGMA_RNG As Variant, _
Optional ByVal STEP_SIZE As Double = 0.5)

'STEP_SIZE : time indexing increment size in year

Dim i As Long '
Dim j As Long '
Dim k As Long
  
Dim ii As Long
Dim jj As Long
  
Dim nSTEPS As Long 'no of time steps for tree
Dim TEMP_STR As String
  
Dim TEMP_FU As Double
Dim TEMP_FD As Double
Dim TEMP_FT As Double
  
Dim TEMP_SUM As Double
Dim TEMP_VAL As Double
Dim TEMP_FACT As Double
  
Dim TEMP_DENOM As Double
Dim TEMP_NUMER As Double
  
Dim FORW_ARR As Variant 'initial forward rate vector
Dim TREE_ARR As Variant
Dim TEMP_ARR As Variant
Dim TEMP_GROUP As Variant ''array to store trees of forard rate evolution

Dim FORW_VECTOR As Variant
Dim VOLAT_VECTOR As Variant
Dim TEMP_VECTOR As Variant
Dim TEMP_MATRIX As Variant
  
Dim VOLAT_ARR As Variant

On Error GoTo ERROR_LABEL

FORW_VECTOR = FORW_RNG
If UBound(FORW_VECTOR, 1) = 1 Then
    FORW_VECTOR = MATRIX_TRANSPOSE_FUNC(FORW_VECTOR)
End If

VOLAT_VECTOR = SIGMA_RNG
If UBound(VOLAT_VECTOR, 1) = 1 Then
    VOLAT_VECTOR = MATRIX_TRANSPOSE_FUNC(VOLAT_VECTOR)
End If

ReDim FORW_ARR(0 To UBound(FORW_VECTOR) - 1)
ReDim VOLAT_ARR(0 To UBound(VOLAT_VECTOR) - 1)

j = 1
For i = 0 To UBound(FORW_ARR)
    FORW_ARR(i) = FORW_VECTOR(j, 1)
    j = j + 1
Next i
j = 1
For i = 0 To UBound(VOLAT_ARR)
    VOLAT_ARR(i) = VOLAT_VECTOR(j, 1)
    j = j + 1
Next i
  
nSTEPS = UBound(FORW_ARR, 1) - LBound(FORW_ARR, 1)

ReDim TEMP_GROUP(0 To nSTEPS)
  
For i = 0 To nSTEPS 'TREE_ARR will store evolution of forward rate
    ReDim TREE_ARR(0 To i) As Variant
    For j = 0 To i
      TEMP_FACT = 2 ^ j
      ReDim TEMP_ARR(0 To TEMP_FACT - 1)
      'initialize the first node of each tree to
      'initial forward rate at t=0
      If j = 0 Then: TEMP_ARR(0) = FORW_ARR(i)
      TREE_ARR(j) = TEMP_ARR
    Next j
    TEMP_GROUP(i) = TREE_ARR
Next i

For ii = 0 To nSTEPS - 1
  For i = ii To nSTEPS - 1
    TREE_ARR = TEMP_GROUP(i + 1)
    TEMP_VECTOR = TREE_ARR(ii)
    TEMP_ARR = TREE_ARR(ii + 1)
    jj = 0
    For j = 0 To UBound(TEMP_VECTOR, 1)
      TEMP_FT = TEMP_VECTOR(j)
      TEMP_SUM = 0
      For k = ii + 1 To i + 1
        TEMP_SUM = TEMP_SUM + (Log(FORW_ARR(ii)) * VOLAT_ARR((Abs((ii - k) _
                                * STEP_SIZE)) - 1)) * STEP_SIZE ^ 0.5
      Next k
      TEMP_NUMER = (Exp(TEMP_SUM * STEP_SIZE) + _
                    Exp(-TEMP_SUM * STEP_SIZE)) / 2 'hyperbolic
      'cosine of a number
      TEMP_SUM = 0
      For k = ii + 1 To i
        TEMP_SUM = TEMP_SUM + (Log(FORW_ARR(ii)) * _
                    VOLAT_ARR((Abs((ii - k) * _
                    STEP_SIZE)) - 1)) * STEP_SIZE ^ 0.5
      Next k
      TEMP_DENOM = (Exp(TEMP_SUM * STEP_SIZE) + _
                    Exp(-TEMP_SUM * STEP_SIZE)) / 2 'hyperbolic
      'cosine of a number
      TEMP_VAL = (Log(FORW_ARR(ii)) * VOLAT_ARR((Abs((ii - _
                (i + 1)) * STEP_SIZE)) - 1)) * _
                (STEP_SIZE ^ 0.5) * STEP_SIZE
      TEMP_FU = TEMP_FT * (TEMP_NUMER / TEMP_DENOM) * Exp(TEMP_VAL)
      TEMP_FD = TEMP_FT * (TEMP_NUMER / TEMP_DENOM) * Exp(-TEMP_VAL)
      TEMP_ARR(jj) = TEMP_FD
      jj = jj + 1
      TEMP_ARR(jj) = TEMP_FU
      jj = jj + 1
    Next j
    TREE_ARR(ii + 1) = TEMP_ARR
    TEMP_GROUP(i + 1) = TREE_ARR
  Next i
Next ii

jj = 1
For ii = 0 To nSTEPS
    TEMP_VECTOR = TEMP_GROUP(ii)(UBound(TEMP_GROUP(ii), 1))
    jj = jj + UBound(TEMP_VECTOR, 1) + 2
Next ii

ReDim TEMP_MATRIX(1 To jj, 1 To nSTEPS + 1) 'Reset

jj = 1
For ii = 0 To nSTEPS
    TEMP_STR = "Tree for period t = " & (ii * STEP_SIZE) & _
                " to t = " & ((ii + 1) * STEP_SIZE)
    TEMP_VECTOR = TEMP_GROUP(ii)(UBound(TEMP_GROUP(ii), 1))
    TEMP_MATRIX(jj, 1) = TEMP_STR
    For i = 0 To UBound(TEMP_GROUP(ii), 1)
        TEMP_VECTOR = TEMP_GROUP(ii)(i)
        k = 0
        For j = 0 To UBound(TEMP_VECTOR, 1)
          TEMP_MATRIX(jj + k + 1, i + 1) = TEMP_VECTOR(j)
          k = k + 1
        Next j
    Next i
    jj = jj + UBound(TEMP_VECTOR, 1) + 2
Next ii

'HouseKeeping
For j = LBound(TEMP_MATRIX, 2) To UBound(TEMP_MATRIX, 2)
    For i = LBound(TEMP_MATRIX, 1) To UBound(TEMP_MATRIX, 1)
        If (TEMP_MATRIX(i, j) = 0 Or TEMP_MATRIX(i, j) = "") _
            Then: TEMP_MATRIX(i, j) = ""
    Next i
Next j
HJM_BUSHY_TREE_FUNC = TEMP_MATRIX

Exit Function
ERROR_LABEL:
HJM_BUSHY_TREE_FUNC = Err.number
End Function
