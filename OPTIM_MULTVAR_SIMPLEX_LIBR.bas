Attribute VB_Name = "OPTIM_MULTVAR_SIMPLEX_LIBR"


'-------------------------------------------------------------------------------
'-------------------------------------------------------------------------------
Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.
'-------------------------------------------------------------------------------
Private Const PUB_EPSILON As Double = 2 ^ 52
'-------------------------------------------------------------------------------


'************************************************************************************
'************************************************************************************
'FUNCTION      : SIMPLEX_MINIMUM_OPTIMIZATION_FUNC

'DESCRIPTION   : Simplex minimum function in a multivariable optimization scenario
'In mathematical optimization theory, the simplex algorithm, created by the
'American mathematician George Dantzig in 1947, is a popular algorithm for
'numerical solution of the linear programming problem. The journal Computing
'in Science and Engineering listed it as one of the top 10 algorithms of
'the century.

'An unrelated, but similarly named method is the Nelder-Mead method or downhill
'simplex method due to Nelder & Mead (1965) and is a numerical method for
'optimising many-dimensional unconstrained problems, belonging to the more
'general class of search algorithms.

'In both cases, the method uses the concept of a simplex, which is a polytope of
'N + 1 vertices in N dimensions: a line segment in one dimension, a triangle in
'two dimensions, a tetrahedron in three-dimensional space and so forth.

'LIBRARY       : OPTIMIZATION
'GROUP         : MULTVAR_SIMPLEX
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 02/13/2009
'************************************************************************************
'************************************************************************************

Function SIMPLEX_MINIMUM_OPTIMIZATION_FUNC(ByVal FUNC_NAME_STR As String, _
ByVal CONST_STR_NAME As String, _
ByRef PARAM_RNG As Variant, _
Optional ByRef RELAX_VAL As Double = 0.01, _
Optional ByVal nLOOPS As Single = 200, _
Optional ByVal tolerance As Double = 0.01)

Dim i As Single
Dim j As Single
Dim k As Single
Dim l As Single

Dim NROWS As Single
Dim COUNTER As Single

Dim MIN_VAL As Double
Dim MAX_VAL As Double

Dim XTEMP_VAL As Double
Dim YTEMP_VAL As Double

Dim DELTA_VAL As Double
Dim SAVED_VAL As Double
Dim FACTOR_VAL As Double

Dim TEMP_ARR As Variant
Dim DATA_VECTOR As Variant

Dim SUM_MATRIX As Variant
Dim PARAM_VECTOR As Variant
Dim VERTICES_GROUP As Variant
Dim VALUES_VECTOR As Variant

Dim epsilon As Double

On Error GoTo ERROR_LABEL

epsilon = 2 ^ -52
PARAM_VECTOR = PARAM_RNG
If UBound(PARAM_VECTOR, 1) = 1 Then: PARAM_VECTOR = MATRIX_TRANSPOSE_FUNC(PARAM_VECTOR)
NROWS = UBound(PARAM_VECTOR, 1)
            
ReDim TEMP_ARR(1 To NROWS + 1)
VERTICES_GROUP = TEMP_ARR
For i = 1 To NROWS + 1
    VERTICES_GROUP(i) = PARAM_VECTOR
Next i

For i = 1 To NROWS
    ReDim DATA_VECTOR(1 To NROWS, 1 To 1)
    DATA_VECTOR(i, 1) = 1
    XTEMP_VAL = SIMPLEX_MINIMUM_ROTATION_FUNC(CONST_STR_NAME, VERTICES_GROUP(i + 1), DATA_VECTOR, RELAX_VAL, nLOOPS)
Next i

ReDim VALUES_VECTOR(1 To NROWS + 1, 1 To 1)
For i = 1 To NROWS + 1
    VALUES_VECTOR(i, 1) = CALL_SIMPLEX_MINIMUM_OBJ_FUNC(FUNC_NAME_STR, VERTICES_GROUP(i))
Next i
COUNTER = 2
 
Do While (1 < 2)
    ReDim SUM_MATRIX(1 To NROWS, 1 To 1)
    For i = 1 To NROWS + 1
        SUM_MATRIX = MATRIX_ELEMENTS_ADD_FUNC(SUM_MATRIX, VERTICES_GROUP(i))
    Next i
    'Determine best, worst and 2nd worst vertices
    j = 1
    If (VALUES_VECTOR(1, 1) < VALUES_VECTOR(2, 1)) Then
        k = 2
        l = 1
    Else
        k = 1
        l = 2
    End If
    For i = 2 To NROWS + 1
        If (VALUES_VECTOR(i, 1) > VALUES_VECTOR(k, 1)) Then
            l = k
            k = i
        Else
            If ((VALUES_VECTOR(i, 1) > VALUES_VECTOR(l, 1)) _
            And i <> k) Then
                l = i
            End If
        End If
        If (VALUES_VECTOR(i, 1) < VALUES_VECTOR(j, 1)) Then
            j = i
        End If
    Next i
    
    MIN_VAL = VALUES_VECTOR(j, 1)
    MAX_VAL = VALUES_VECTOR(k, 1)
    DELTA_VAL = 2 * Abs(MAX_VAL - MIN_VAL) / (Abs(MAX_VAL) + Abs(MIN_VAL) + epsilon)
    If (DELTA_VAL < tolerance) Then
        SIMPLEX_MINIMUM_OPTIMIZATION_FUNC = VERTICES_GROUP(j)
        Exit Function
    End If
    
    FACTOR_VAL = -1
    YTEMP_VAL = _
        SIMPLEX_MINIMUM_EXTRAPOLATION_FUNC(FUNC_NAME_STR, CONST_STR_NAME, _
            k, FACTOR_VAL, VERTICES_GROUP, SUM_MATRIX, VALUES_VECTOR, nLOOPS)
    If ((YTEMP_VAL <= VALUES_VECTOR(j, 1)) And (FACTOR_VAL = -1)) Then
        FACTOR_VAL = 2
        XTEMP_VAL = SIMPLEX_MINIMUM_EXTRAPOLATION_FUNC(FUNC_NAME_STR, CONST_STR_NAME, k, FACTOR_VAL, VERTICES_GROUP, SUM_MATRIX, VALUES_VECTOR, nLOOPS)
    Else
        If (YTEMP_VAL >= VALUES_VECTOR(l, 1)) Then
            SAVED_VAL = VALUES_VECTOR(k, 1)
            FACTOR_VAL = 0.5
            YTEMP_VAL = _
                SIMPLEX_MINIMUM_EXTRAPOLATION_FUNC(FUNC_NAME_STR, CONST_STR_NAME, _
                    k, FACTOR_VAL, VERTICES_GROUP, SUM_MATRIX, VALUES_VECTOR, nLOOPS)
            If (YTEMP_VAL >= SAVED_VAL) Then
                For i = 1 To NROWS + 1
                    If (i <> j) Then
                        VERTICES_GROUP(i) = MATRIX_ELEMENTS_MULT_SCALAR_FUNC(MATRIX_ELEMENTS_ADD_FUNC(VERTICES_GROUP(i), VERTICES_GROUP(j)), 0.5)
                        VALUES_VECTOR(i, 1) = CALL_SIMPLEX_MINIMUM_OBJ_FUNC(FUNC_NAME_STR, VERTICES_GROUP(i))
                    End If
                Next i
            End If
        End If
    End If
   COUNTER = COUNTER + 1
   If (COUNTER Mod 100000) = 0 Then: GoTo ERROR_LABEL
Loop

Exit Function
ERROR_LABEL:
SIMPLEX_MINIMUM_OPTIMIZATION_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : CALL_SIMPLEX_MINIMUM_CONST_FUNC
'DESCRIPTION   : Constraint Function
'LIBRARY       : OPTIMIZATION
'GROUP         : MULTVAR_SIMPLEX
'ID            : 002
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 02/13/2009
'************************************************************************************
'************************************************************************************

Private Function CALL_SIMPLEX_MINIMUM_CONST_FUNC(ByVal CONST_STR_NAME As String, _
ByRef PARAM_RNG As Variant) As Boolean
Dim PARAM_VECTOR As Variant
On Error GoTo ERROR_LABEL
PARAM_VECTOR = PARAM_RNG
If UBound(PARAM_VECTOR, 1) = 1 Then: PARAM_VECTOR = MATRIX_TRANSPOSE_FUNC(PARAM_VECTOR)
CALL_SIMPLEX_MINIMUM_CONST_FUNC = Excel.Application.Run(CONST_STR_NAME, PARAM_VECTOR)
Exit Function
ERROR_LABEL:
CALL_SIMPLEX_MINIMUM_CONST_FUNC = False
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : CALL_SIMPLEX_MINIMUM_OBJ_FUNC
'DESCRIPTION   : Objective Function
'LIBRARY       : OPTIMIZATION
'GROUP         : MULTVAR_SIMPLEX
'ID            : 003
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 02/13/2009
'************************************************************************************
'************************************************************************************

Private Function CALL_SIMPLEX_MINIMUM_OBJ_FUNC(ByVal FUNC_NAME_STR As String, _
ByRef PARAM_RNG As Variant) As Double
Dim PARAM_VECTOR As Variant
On Error GoTo ERROR_LABEL
PARAM_VECTOR = PARAM_RNG
If UBound(PARAM_VECTOR, 1) = 1 Then: PARAM_VECTOR = MATRIX_TRANSPOSE_FUNC(PARAM_VECTOR)
CALL_SIMPLEX_MINIMUM_OBJ_FUNC = Excel.Application.Run(FUNC_NAME_STR, PARAM_VECTOR)
Exit Function
ERROR_LABEL:
CALL_SIMPLEX_MINIMUM_OBJ_FUNC = PUB_EPSILON
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : SIMPLEX_MINIMUM_EXTRAPOLATION_FUNC
'DESCRIPTION   : Extrapolate Vertices
'LIBRARY       : OPTIMIZATION
'GROUP         : MULTVAR_SIMPLEX
'ID            : 004
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 02/13/2009
'************************************************************************************
'************************************************************************************


Private Function SIMPLEX_MINIMUM_EXTRAPOLATION_FUNC(ByVal FUNC_NAME_STR As String, _
ByVal CONST_STR_NAME As String, _
ByVal jj As Single, _
ByRef FACTOR_VAL As Double, _
ByRef VERTICES_GROUP As Variant, _
ByRef SUM_MATRIX As Variant, _
ByRef VALUES_VECTOR As Variant, _
Optional ByVal nLOOPS As Single = 200) 'As Double
             
Dim NSIZE As Single
Dim COUNTER As Single
Dim YTEMP_VAL As Double
Dim ATEMP_VAL As Double
Dim BTEMP_VAL As Double

Dim XTEMP_VECTOR As Variant

On Error GoTo ERROR_LABEL

COUNTER = 0
Do While (1 < 2)
    NSIZE = UBound(VALUES_VECTOR, 1) - 1
    ATEMP_VAL = (1 - FACTOR_VAL) / NSIZE
    BTEMP_VAL = ATEMP_VAL - FACTOR_VAL
    XTEMP_VECTOR = MATRIX_ELEMENTS_SUBTRACT_FUNC(MATRIX_ELEMENTS_MULT_SCALAR_FUNC(SUM_MATRIX, ATEMP_VAL), _
                   MATRIX_ELEMENTS_MULT_SCALAR_FUNC(VERTICES_GROUP(jj), BTEMP_VAL))
                
    FACTOR_VAL = FACTOR_VAL * 0.5
    If (CALL_SIMPLEX_MINIMUM_CONST_FUNC(CONST_STR_NAME, XTEMP_VECTOR) = True) Then: Exit Do
    COUNTER = COUNTER + 1
    If (COUNTER Mod nLOOPS) = 0 Then: GoTo ERROR_LABEL
Loop
    
FACTOR_VAL = FACTOR_VAL * 2
YTEMP_VAL = CALL_SIMPLEX_MINIMUM_OBJ_FUNC(FUNC_NAME_STR, XTEMP_VECTOR)
If (YTEMP_VAL < VALUES_VECTOR(jj, 1)) Then
    VALUES_VECTOR(jj, 1) = YTEMP_VAL
    SUM_MATRIX = MATRIX_ELEMENTS_ADD_FUNC(SUM_MATRIX, MATRIX_ELEMENTS_SUBTRACT_FUNC(XTEMP_VECTOR, VERTICES_GROUP(jj)))
    VERTICES_GROUP(jj) = XTEMP_VECTOR
End If
    
SIMPLEX_MINIMUM_EXTRAPOLATION_FUNC = YTEMP_VAL

Exit Function
ERROR_LABEL:
SIMPLEX_MINIMUM_EXTRAPOLATION_FUNC = PUB_EPSILON
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : SIMPLEX_MINIMUM_ROTATION_FUNC
'DESCRIPTION   : Rotation Function
'LIBRARY       : OPTIMIZATION
'GROUP         : MULTVAR_SIMPLEX
'ID            : 005
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 02/13/2009
'************************************************************************************
'************************************************************************************
            
Private Function SIMPLEX_MINIMUM_ROTATION_FUNC(ByVal CONST_STR_NAME As String, _
ByRef PARAM_VECTOR As Variant, _
ByVal DATA_VECTOR As Variant, _
ByRef RELAX_VAL As Double, _
Optional ByVal nLOOPS As Single = 200) 'As Double

Dim COUNTER As Single
Dim DIFF_VAL As Double
Dim CONST_FLAG As Boolean
Dim XTEMP_VECTOR As Variant

On Error GoTo ERROR_LABEL

DIFF_VAL = RELAX_VAL
XTEMP_VECTOR = MATRIX_ELEMENTS_ADD_FUNC(PARAM_VECTOR, _
               MATRIX_ELEMENTS_MULT_SCALAR_FUNC(DATA_VECTOR, DIFF_VAL))
CONST_FLAG = CALL_SIMPLEX_MINIMUM_CONST_FUNC(CONST_STR_NAME, XTEMP_VECTOR)
COUNTER = 0
Do While (CONST_FLAG = False)
    If (COUNTER > nLOOPS) Then: GoTo ERROR_LABEL 'Can't update parameter vector
    DIFF_VAL = DIFF_VAL * 0.5
    COUNTER = COUNTER + 1
    XTEMP_VECTOR = MATRIX_ELEMENTS_ADD_FUNC(PARAM_VECTOR, _
                   MATRIX_ELEMENTS_MULT_SCALAR_FUNC(DATA_VECTOR, DIFF_VAL))
    CONST_FLAG = CALL_SIMPLEX_MINIMUM_CONST_FUNC(CONST_STR_NAME, XTEMP_VECTOR)
Loop
PARAM_VECTOR = MATRIX_ELEMENTS_ADD_FUNC(PARAM_VECTOR, _
               MATRIX_ELEMENTS_MULT_SCALAR_FUNC(DATA_VECTOR, DIFF_VAL))
SIMPLEX_MINIMUM_ROTATION_FUNC = DIFF_VAL

Exit Function
ERROR_LABEL:
SIMPLEX_MINIMUM_ROTATION_FUNC = PUB_EPSILON
End Function
