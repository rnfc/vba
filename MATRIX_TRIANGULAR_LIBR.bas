Attribute VB_Name = "MATRIX_TRIANGULAR_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.


'************************************************************************************
'************************************************************************************
'FUNCTION      : MATRIX_TRIANGULAR_VALIDATE_FUNC
'DESCRIPTION   : Check if a matrix is triangular
'LIBRARY       : MATRIX
'GROUP         : TRIANGULAR
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/27/2009
'************************************************************************************
'************************************************************************************

Function MATRIX_TRIANGULAR_VALIDATE_FUNC(ByRef DATA_RNG As Variant, _
Optional ByVal epsilon As Double = 5 * 10 ^ -16)

'1 = triangular lower
'2 = triangular upper,
'3 = diagonal,
'0 = else

Dim i As Long
Dim j As Long
Dim k As Long

Dim NROWS As Long
Dim NCOLUMNS As Long

Dim LOWER_FLAG As Variant
Dim UPPER_FLAG As Variant
Dim DATA_MATRIX As Variant
     
On Error GoTo ERROR_LABEL
     
DATA_MATRIX = DATA_RNG
NROWS = UBound(DATA_MATRIX, 1)
NCOLUMNS = UBound(DATA_MATRIX, 2)
    
If NROWS <> NCOLUMNS Then GoTo 1984
LOWER_FLAG = True
For i = 1 To NROWS
    For j = i + 1 To NROWS
        If Abs(DATA_MATRIX(i, j)) > epsilon Then LOWER_FLAG = False: GoTo 1983
    Next j
Next i

1983: UPPER_FLAG = True
     For i = 1 To NROWS
        For j = i + 1 To NROWS
            If Abs(DATA_MATRIX(j, i)) > epsilon Then UPPER_FLAG = False: GoTo 1984
        Next j
     Next i

1984: If UPPER_FLAG And LOWER_FLAG Then
        k = 3
     ElseIf UPPER_FLAG And Not LOWER_FLAG Then
        k = 2
     ElseIf Not UPPER_FLAG And LOWER_FLAG Then
        k = 1
     Else
        k = 0
     End If
     
     MATRIX_TRIANGULAR_VALIDATE_FUNC = k
    
Exit Function
ERROR_LABEL:
MATRIX_TRIANGULAR_VALIDATE_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : MATRIX_TRIANGULAR_CHECK_FUNC
'DESCRIPTION   : Check if matrix is triangular upper or lower or any
'LIBRARY       : MATRIX
'GROUP         : TRIANGULAR
'ID            : 002
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/27/2009
'************************************************************************************
'************************************************************************************

Function MATRIX_TRIANGULAR_CHECK_FUNC(ByRef DATA_RNG As Variant, _
Optional ByVal epsilon As Double = 10 ^ -15)

Dim i As Long
Dim j As Long

Dim NROWS As Long
Dim NCOLUMNS As Long

Dim JUMPER_STR As String
Dim DATA_MATRIX As Variant

On Error GoTo ERROR_LABEL

DATA_MATRIX = DATA_RNG

NROWS = UBound(DATA_MATRIX, 1)
NCOLUMNS = UBound(DATA_MATRIX, 2)

JUMPER_STR = ""
If JUMPER_STR = "" Then GoSub Upper     'try upper-triangular
If JUMPER_STR = "" Then GoSub Lower  'try lower-triangular

MATRIX_TRIANGULAR_CHECK_FUNC = JUMPER_STR

Exit Function
'-------------------
Upper:
For i = 1 To NROWS
    For j = 1 To i - 1
        If Abs(DATA_MATRIX(i, j)) > epsilon Then Return
    Next j
Next i
JUMPER_STR = "U"
Return
'-------------------
Lower:
For i = 1 To NROWS
    For j = i + 1 To NCOLUMNS
        If Abs(DATA_MATRIX(i, j)) > epsilon Then Return
    Next j
Next i
JUMPER_STR = "L"
Return

ERROR_LABEL:
MATRIX_TRIANGULAR_CHECK_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : MATRIX_TRIANGULAR_LOWER_SUM_FUNC
'DESCRIPTION   : Returns the Sum of Squares of the Lower Triangle
'LIBRARY       : MATRIX
'GROUP         : TRIANGULAR
'ID            : 003
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/27/2009
'************************************************************************************
'************************************************************************************


Function MATRIX_TRIANGULAR_LOWER_SUM_FUNC(ByRef DATA_RNG As Variant)
       
Dim i As Long
Dim j As Long

Dim NSIZE As Long

Dim NROWS As Long
Dim NCOLUMNS As Long

Dim TEMP_SUM As Double
Dim DATA_MATRIX As Variant

On Error GoTo ERROR_LABEL

DATA_MATRIX = DATA_RNG

NROWS = UBound(DATA_MATRIX, 1)
NCOLUMNS = UBound(DATA_MATRIX, 2)

NSIZE = Int(Sqr(NROWS * NCOLUMNS))
TEMP_SUM = 0
For i = 1 To NSIZE
    For j = i To NSIZE
        TEMP_SUM = TEMP_SUM + (DATA_MATRIX(j, i) ^ 2)
    Next j
Next i
MATRIX_TRIANGULAR_LOWER_SUM_FUNC = TEMP_SUM

Exit Function
ERROR_LABEL:
MATRIX_TRIANGULAR_LOWER_SUM_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : MATRIX_TRIANGULAR_UPPER_SUM_FUNC
'DESCRIPTION   : Returns the Sum of Squares of the Upper Triangle
'LIBRARY       : MATRIX
'GROUP         : TRIANGULAR
'ID            : 004
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/27/2009
'************************************************************************************
'************************************************************************************


Function MATRIX_TRIANGULAR_UPPER_SUM_FUNC(ByRef DATA_RNG As Variant)
    
Dim i As Long
Dim j As Long

Dim NSIZE As Long
Dim NROWS As Long
Dim NCOLUMNS As Long

Dim TEMP_SUM As Double
Dim DATA_MATRIX As Variant

On Error GoTo ERROR_LABEL

DATA_MATRIX = DATA_RNG

NROWS = UBound(DATA_MATRIX, 1)
NCOLUMNS = UBound(DATA_MATRIX, 2)

NSIZE = Int(Sqr(NROWS * NCOLUMNS))

TEMP_SUM = 0
For i = 1 To NSIZE
    For j = i + 1 To NSIZE
        TEMP_SUM = TEMP_SUM + (DATA_MATRIX(i, j) ^ 2)
    Next j
Next i

MATRIX_TRIANGULAR_UPPER_SUM_FUNC = TEMP_SUM

Exit Function
ERROR_LABEL:
MATRIX_TRIANGULAR_UPPER_SUM_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : MATRIX_TRIANGULAR_GJ_LINEAR_SYSTEM_FUNC

'DESCRIPTION   : This function solves a triangular linear system by the forward
'and backward substitution algorithms.


'MATRIX_RNG: is the triangular - upper or lower - system square
'matrix (n x n)

'VECTOR_RNG: is a constant (n x 1) vector or a constant (n x m ) matrix

'OUTPUT: is the unknown (n x 1) vector or the (n x m)
'unknown matrix

'As known, the above linear system has only one solution if - and
'only if -, det(A) <> 0. Otherwise the solutions can be infinite
'or even non-existing. In that case the system is called "singular".

'The parameter DATA_VECTOR can be also a (n x m) matrix. In that case the
'function returns a matrix solution X of the multiple linear system
'Parameter typ = "U" or "L" switches the function from solving for
'the upper-triangular (back substitutions) or lower-triangular system
'(forward substitutions); if omitted, the function automatically detects
'the type of the system.

'Optional parameter epsilon (default is 1E-15) sets the minimum round-off
'error; any value of absolute value less than Tiny will be set to 0.

'LIBRARY       : MATRIX
'GROUP         : TRIANGULAR
'ID            : 005
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/27/2009
'************************************************************************************
'************************************************************************************


Function MATRIX_TRIANGULAR_GJ_LINEAR_SYSTEM_FUNC(ByRef MATRIX_RNG As Variant, _
ByRef VECTOR_RNG As Variant, _
Optional ByVal epsilon As Double = 10 ^ -15, _
Optional ByVal OUTPUT As Integer = 0)

Dim i As Long
Dim j As Long
Dim k As Long

Dim NROWS As Long
Dim NCOLUMNS As Long

Dim CHECK_STR As String
Dim TEMP_VAL As Double
Dim DETERM_VAL As Double

Dim DATA_MATRIX As Variant
Dim DATA_VECTOR As Variant

On Error GoTo ERROR_LABEL

DATA_MATRIX = MATRIX_RNG
DATA_VECTOR = VECTOR_RNG

If UBound(DATA_MATRIX, 1) <> UBound(DATA_MATRIX, 2) Or _
    UBound(DATA_MATRIX, 1) <> UBound(DATA_VECTOR, 1) Then GoTo ERROR_LABEL

NROWS = UBound(DATA_MATRIX, 1)
NCOLUMNS = UBound(DATA_VECTOR, 2)

CHECK_STR = MATRIX_TRIANGULAR_CHECK_FUNC(DATA_MATRIX, epsilon)

Select Case CHECK_STR
Case "L"
    GoSub 1983 'forward substitution
Case "U"
    GoSub 1984 'back substitution
Case Else
    GoTo ERROR_LABEL
End Select

Select Case OUTPUT
    Case 0
        MATRIX_TRIANGULAR_GJ_LINEAR_SYSTEM_FUNC = DATA_VECTOR
    Case 1
        MATRIX_TRIANGULAR_GJ_LINEAR_SYSTEM_FUNC = DETERM_VAL
    Case 2
        MATRIX_TRIANGULAR_GJ_LINEAR_SYSTEM_FUNC = DATA_VECTOR
    Case Else
        MATRIX_TRIANGULAR_GJ_LINEAR_SYSTEM_FUNC = Array(DATA_VECTOR, DETERM_VAL, DATA_VECTOR)
End Select

'-------------------------------------------------------------------------------------------------
Exit Function
'-------------------------------------------------------------------------------------------------
1983: 'forward substitution
'-------------------------------------------------------------------------------------------------
    DETERM_VAL = 1
    For k = 1 To NCOLUMNS
        For i = 1 To NROWS
            If Abs(DATA_MATRIX(i, i)) <= epsilon Then
                DETERM_VAL = 0
                GoTo ERROR_LABEL
            Else
                DETERM_VAL = DETERM_VAL * DATA_MATRIX(i, i)
            End If
            TEMP_VAL = DATA_VECTOR(i, k)
            For j = 1 To i - 1
                TEMP_VAL = TEMP_VAL - DATA_MATRIX(i, j) * DATA_VECTOR(j, k)
            Next j
            DATA_VECTOR(i, k) = TEMP_VAL / DATA_MATRIX(i, i)
        Next i
    Next k
'-------------------------------------------------------------------------------------------------
Return
'-------------------------------------------------------------------------------------------------
1984: 'back substitution
    DETERM_VAL = 1
    For k = 1 To NCOLUMNS
        For i = NROWS To 1 Step -1
            If Abs(DATA_MATRIX(i, i)) <= epsilon Then
                DETERM_VAL = 0
                GoTo ERROR_LABEL
            Else
                DETERM_VAL = DETERM_VAL * DATA_MATRIX(i, i)
            End If
        
            TEMP_VAL = DATA_VECTOR(i, k)
            For j = i + 1 To NROWS
                TEMP_VAL = TEMP_VAL - DATA_MATRIX(i, j) * DATA_VECTOR(j, k)
            Next j
            DATA_VECTOR(i, k) = TEMP_VAL / DATA_MATRIX(i, i)
        Next i
    Next k
'-------------------------------------------------------------------------------------------------
Return
'-------------------------------------------------------------------------------------------------
ERROR_LABEL:
MATRIX_TRIANGULAR_GJ_LINEAR_SYSTEM_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : MATRIX_TRIANGULAR_ERROR_FUNC
'DESCRIPTION   : Return the error for for a triangular matrix
'LIBRARY       : MATRIX
'GROUP         : TRIANGULAR
'ID            : 006
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/27/2009
'************************************************************************************
'************************************************************************************

Function MATRIX_TRIANGULAR_ERROR_FUNC(ByRef DATA_RNG As Variant)

Dim i As Long
Dim j As Long

Dim NROWS As Long
Dim NCOLUMNS As Long

Dim TEMP_MIN As Double
Dim TEMP_SUM As Double

Dim LOWER_BOUND As Double
Dim UPPER_BOUND As Double

Dim DATA_MATRIX As Variant

On Error GoTo ERROR_LABEL

DATA_MATRIX = DATA_RNG
NROWS = UBound(DATA_MATRIX, 1)
NCOLUMNS = (NROWS ^ 2 - NROWS) / 2

'---------------------lower triangular error
TEMP_SUM = 0
For i = 1 To NROWS
    For j = i + 1 To NROWS
        TEMP_SUM = TEMP_SUM + Abs(DATA_MATRIX(i, j))
    Next j
Next i
LOWER_BOUND = TEMP_SUM / NCOLUMNS

TEMP_SUM = 0
For j = 1 To NROWS
    For i = j + 1 To NROWS
      TEMP_SUM = TEMP_SUM + Abs(DATA_MATRIX(i, j))
    Next i
Next j

UPPER_BOUND = TEMP_SUM / NCOLUMNS
TEMP_MIN = MINIMUM_FUNC(UPPER_BOUND, LOWER_BOUND)

MATRIX_TRIANGULAR_ERROR_FUNC = TEMP_MIN
  
Exit Function
ERROR_LABEL:
MATRIX_TRIANGULAR_ERROR_FUNC = Err.number
End Function
