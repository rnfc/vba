Attribute VB_Name = "WEB_STRING_MATH_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.

'************************************************************************************
'************************************************************************************
'FUNCTION      : MATH_EXPRESSION_VALIDATE_FUNC
'DESCRIPTION   : Evaluate Formulas String in a vector
'LIBRARY       : STRING
'GROUP         : FORMULAS
'ID            : 001
'LAST UPDATE   : 14/02/2008
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'************************************************************************************
'************************************************************************************

Function MATH_EXPRESSION_VALIDATE_FUNC(ByVal FORMULA0_STR As String, _
ByRef VALUES_RNG As Variant, _
ByRef VARIABLES_RNG As Variant)

Dim i As Long
Dim NROWS As Long

Dim TEMP_VAL As Double

Dim FORMULA1_STR As String
Dim FORMULA2_STR As String

Dim VALUES_VECTOR As Variant
Dim VARIABLES_VECTOR As Variant

On Error GoTo ERROR_LABEL

FORMULA2_STR = FORMULA0_STR
VALUES_VECTOR = VALUES_RNG
VARIABLES_VECTOR = VARIABLES_RNG

If (IsArray(VALUES_VECTOR) = True) And (IsArray(VARIABLES_VECTOR) = True) Then
    If UBound(VALUES_VECTOR, 1) = 1 Then: VALUES_VECTOR = MATRIX_TRANSPOSE_FUNC(VALUES_VECTOR)
    If UBound(VARIABLES_VECTOR, 1) = 1 Then: VARIABLES_VECTOR = MATRIX_TRANSPOSE_FUNC(VARIABLES_VECTOR)
    If UBound(VALUES_VECTOR, 1) <> UBound(VARIABLES_VECTOR, 1) Then: GoTo ERROR_LABEL
    NROWS = UBound(VALUES_VECTOR, 1)
    For i = 1 To NROWS
       FORMULA1_STR = "(" & Trim(CStr(VALUES_VECTOR(i, 1))) & ")"
       FORMULA2_STR = Replace(FORMULA2_STR, VARIABLES_VECTOR(i, 1), FORMULA1_STR, 1, -1, 0)
    Next i
Else
    FORMULA1_STR = "(" & Trim(CStr(VALUES_VECTOR)) & ")"
    FORMULA2_STR = Replace(FORMULA2_STR, VARIABLES_VECTOR, FORMULA1_STR, 1, -1, 0)
End If

TEMP_VAL = Excel.Application.Evaluate(FORMULA2_STR)
MATH_EXPRESSION_VALIDATE_FUNC = TEMP_VAL

Exit Function
ERROR_LABEL:
MATH_EXPRESSION_VALIDATE_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : MATH_EXPRESSION_SYNTAX_FUNC
'DESCRIPTION   : Validate Syntax in a vector of formulas
'LIBRARY       : STRING
'GROUP         : FORMULAS
'ID            : 002
'LAST UPDATE   : 14/02/2008
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'************************************************************************************
'************************************************************************************

Function MATH_EXPRESSION_SYNTAX_FUNC(ByRef FORMULAS_RNG As Variant)

Dim i As Long
Dim j As Long
Dim k As Long

Dim NROWS As Long

Dim TEMP_STR As String
Dim ERROR_STR As String

Dim TEMP_MATRIX As Variant
Dim TEMP_VECTOR As Variant
Dim FORMULAS_VECTOR As Variant

On Error GoTo ERROR_LABEL

If IsArray(FORMULAS_RNG) = True Then
    FORMULAS_VECTOR = FORMULAS_RNG
    If UBound(FORMULAS_VECTOR, 1) = 1 Then: _
        FORMULAS_VECTOR = MATRIX_TRANSPOSE_FUNC(FORMULAS_VECTOR)
    NROWS = UBound(FORMULAS_VECTOR, 1)
Else
    ReDim FORMULAS_VECTOR(1 To 1, 1 To 1)
    FORMULAS_VECTOR(1, 1) = FORMULAS_RNG
    NROWS = 1
End If

ReDim TEMP_MATRIX(0 To NROWS, 1 To 3)

TEMP_MATRIX(0, 1) = "FORMULA"
TEMP_MATRIX(0, 2) = "SYNTAX_CHECK"
TEMP_MATRIX(0, 3) = "VARIABLES"

For i = 1 To NROWS
    TEMP_VECTOR = MATH_EXPRESSION_CHECK_FUNC(FORMULAS_VECTOR(i, 1), 0)
    If IsArray(TEMP_VECTOR) = False Then
        TEMP_MATRIX(i, 2) = TEMP_VECTOR
        TEMP_MATRIX(i, 3) = ""
        GoTo 1983
    End If
    k = MATH_EXPRESSION_CHECK_FUNC(FORMULAS_VECTOR(i, 1), 1)
    ERROR_STR = MATH_EXPRESSION_CHECK_FUNC(FORMULAS_VECTOR(i, 1), 2)

    TEMP_STR = ""
    For j = 1 To k
        If j < k Then
            TEMP_STR = TEMP_STR + TEMP_VECTOR(j, 1) + " ; "
        Else
            TEMP_STR = TEMP_STR + TEMP_VECTOR(j, 1)
        End If
    Next j
    TEMP_MATRIX(i, 2) = ERROR_STR
    TEMP_MATRIX(i, 3) = TEMP_STR
1983:
    TEMP_MATRIX(i, 1) = FORMULAS_VECTOR(i, 1)
Next i

MATH_EXPRESSION_SYNTAX_FUNC = TEMP_MATRIX

Exit Function
ERROR_LABEL:
MATH_EXPRESSION_SYNTAX_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : MATH_EXPRESSION_CHECK_FUNC
'DESCRIPTION   : Sintax Check Math Expression Routine
'LIBRARY       : STRING
'GROUP         : FORMULAS
'ID            : 003
'LAST UPDATE   : 12/02/2008
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'************************************************************************************
'************************************************************************************

Private Function MATH_EXPRESSION_CHECK_FUNC(ByVal FORMULA0_STR As Variant, _
Optional ByVal OUTPUT As Integer = 0)

Dim h As Long
Dim i As Long
Dim j As Long
Dim k As Long
Dim l As Long
Dim m As Long 'Maximum number of variables

Dim TEMP_ERR As String '
Dim TEMP_ARG As String
Dim TEMP_STR As String
Dim TEMP_CHR As String

Dim ATEMP_FLAG As Boolean
Dim BTEMP_FLAG As Boolean

Dim VALID_FLAG As Boolean

Dim PROCEDURE_STR As String
Dim FUNC_STR_NAMES As String

On Error GoTo ERROR_LABEL

m = 100
FUNC_STR_NAMES = "Abs Atn Cos Exp Fix Int Ln Log Rnd Sgn " & _
"Sin Sqr Tan Acos Asin Cosh Sinh Tanh Acosh Asinh Atanh Fact"

TEMP_ERR = ""
h = 0
j = 0
PROCEDURE_STR = ""
VALID_FLAG = False
l = Len(FORMULA0_STR)

ReDim TEMP_MATRIX(1 To m, 1 To 1)

For i = 1 To l
    TEMP_CHR = Mid(FORMULA0_STR, i, 1)
    Select Case TEMP_CHR
        Case " "
            'skip
        Case "(", "[", "{"
            j = j + 1
            If PROCEDURE_STR <> "" Then
            TEMP_STR = Mid(FORMULA0_STR, i - Len(PROCEDURE_STR), Len(PROCEDURE_STR))
                If TEMP_STR = PROCEDURE_STR Then
                    ATEMP_FLAG = LOOK_STRING_FLAG_FUNC(PROCEDURE_STR, FUNC_STR_NAMES)
                    If ATEMP_FLAG = False Then
                        MATH_EXPRESSION_CHECK_FUNC = "Function <" & _
                            PROCEDURE_STR & "> unknown:" & CStr(i)
                        Exit Function
                    End If
                    PROCEDURE_STR = ""
                Else
                    MATH_EXPRESSION_CHECK_FUNC = "syntax error"
                    Exit Function
                End If
            End If
        Case ")", "]", "}"
            j = j - 1
        Case "+", "-"
             If PROCEDURE_STR = "" Then PROCEDURE_STR = "0"
             VALID_FLAG = True
         Case "*", "/"
             VALID_FLAG = True
        Case "^"
             VALID_FLAG = True
        Case "!"
             VALID_FLAG = True
             TEMP_ARG = PROCEDURE_STR
        Case Else
             PROCEDURE_STR = PROCEDURE_STR + TEMP_CHR
    End Select
    
    GoSub 1983
Next i
VALID_FLAG = True  'catch last argument
GoSub 1983
'check parenthesis
If j <> 0 Then
    MATH_EXPRESSION_CHECK_FUNC = "parenthesis error"
    Exit Function
End If
If TEMP_ERR = "" Then: TEMP_ERR = "OK"

Select Case OUTPUT
    Case 0
        MATH_EXPRESSION_CHECK_FUNC = TEMP_MATRIX
        'VECTOR_TRIM_FUNC(TEMP_MATRIX, "")
    Case 1
        MATH_EXPRESSION_CHECK_FUNC = h
    Case Else
        MATH_EXPRESSION_CHECK_FUNC = TEMP_ERR
End Select

Exit Function
'-------------------------------------------------------------------------------
1983:
'-------------------------------------------------------------------------------
    If VALID_FLAG = True Then
        If PROCEDURE_STR = "" Then
            MATH_EXPRESSION_CHECK_FUNC = "missing argument"
            Exit Function
        End If
        If LCase(PROCEDURE_STR) = "pi" Then
        ElseIf IsNumeric(PROCEDURE_STR) = False Then
            ATEMP_FLAG = LOOK_STRING_FLAG_FUNC(PROCEDURE_STR, FUNC_STR_NAMES)
            BTEMP_FLAG = Not IS_LETTER_FUNC(Left(PROCEDURE_STR, 1), "_")
            If ATEMP_FLAG = True Or BTEMP_FLAG = True Then
                MATH_EXPRESSION_CHECK_FUNC = "variable name not allowed"
                Exit Function
            End If
            
            For k = 1 To h
                If TEMP_MATRIX(k, 1) = PROCEDURE_STR Then: _
                    GoTo 1984 'variable already inserted
            Next k
    
            If k <= m Then
                TEMP_MATRIX(k, 1) = PROCEDURE_STR 'new variable
                h = k   'incremet TEMP_MATRIX counter
            Else
                h = k   'excedeed limit
            End If
            
1984:
            If h > m Then
                MATH_EXPRESSION_CHECK_FUNC = "too many variables"
                Exit Function
            End If
        End If
        PROCEDURE_STR = ""
        VALID_FLAG = False
        'restore the previous argument only for "!" operator
        If TEMP_CHR = "!" Then PROCEDURE_STR = TEMP_ARG
    End If
'-------------------------------------------------------------------------------
Return
'-------------------------------------------------------------------------------
Exit Function
ERROR_LABEL:
MATH_EXPRESSION_CHECK_FUNC = Err.number
End Function
