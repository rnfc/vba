Attribute VB_Name = "POLYNOMIAL_DETERMINANT_LIBR"


Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.

'************************************************************************************
'************************************************************************************
'FUNCTION      : POLYNOMIAL_PARAMETRIC_DETERM_MATRIX_FUNC

'DESCRIPTION   : Computes the parametric determinant D(k) of a (n x n)
'matrix containing a parameter k. This function returns the polynomial
'string D(k) or its vector coefficients depending on the range
'selected. If you have selected one cell the function returns a
'string; if you have selected a vertical range, the function returns
'a vector. The function accepts one parameter.
'Any matrix element can be a linear function of k: Aij = Qij + Mij * k

'LIBRARY       : POLYNOMIAL
'GROUP         : DETERMINANT
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/22/2009
'************************************************************************************
'************************************************************************************

Function POLYNOMIAL_PARAMETRIC_DETERM_MATRIX_FUNC(ByRef DATA_RNG As Variant, _
Optional ByVal INT_FLAG As Boolean = True, _
Optional ByVal OUTPUT As Integer = 1, _
Optional ByVal epsilon As Double = 10 ^ -12)

Dim h As Long
Dim i As Long
Dim j As Long
Dim k As Long

Dim hh As Long
Dim ii As Long
Dim jj As Long
Dim kk As Long

Dim NDEG As Long
Dim NSIZE As Long

Dim VAR_STR As String
Dim TEMP_STR As String
Dim ERROR_STR As String

Dim ATEMP_VECTOR As Variant
Dim BTEMP_VECTOR As Variant
Dim CTEMP_VECTOR As Variant
Dim DTEMP_VECTOR As Variant
Dim ETEMP_VECTOR As Variant
Dim FTEMP_VECTOR As Variant

Dim TEMP_MATRIX As Variant
Dim DATA_MATRIX As Variant

On Error GoTo ERROR_LABEL

ERROR_STR = Err.number

DATA_MATRIX = DATA_RNG
If UBound(DATA_MATRIX, 1) <> UBound(DATA_MATRIX, 2) Then: GoTo ERROR_LABEL

'1) convert symbolic elements into polynomials
'2) compute the max poly NDEG
'3) check if the matrix is integer

'computes the determinant a matrix containing one parameter k
'returning the polynomial coefficients. It uses the indeterminate
'coefficients method

NSIZE = UBound(DATA_MATRIX, 1)

'1) convert symbolic elements into polynomials
'2) compute the max poly NDEG
'3) check if the matrix is integer

ReDim ATEMP_VECTOR(1 To NSIZE, 1 To 1)
ReDim BTEMP_VECTOR(1 To NSIZE, 1 To 1)

For i = 1 To NSIZE
    For j = 1 To NSIZE
        If Not IsNumeric(DATA_MATRIX(i, j)) Then
            hh = hh + 1
            ATEMP_VECTOR(i, 1) = 1
            BTEMP_VECTOR(j, 1) = 1
        Else
            If INT_FLAG Then INT_FLAG = IS_INTEGER_FUNC(DATA_MATRIX(i, j))
        End If
    Next j
Next i
For i = 1 To NSIZE
    ii = ii + ATEMP_VECTOR(i, 1)
    jj = jj + BTEMP_VECTOR(i, 1)
Next i

If ii > jj Then NDEG = jj Else NDEG = ii ' check NDEG limit
If NDEG > 9 Then
    GoTo ERROR_LABEL 'Too many parameters
End If

ReDim DTEMP_VECTOR(1 To hh, 1 To 2)

k = 0
For i = 1 To NSIZE
    For j = 1 To NSIZE
        If Not IsNumeric(DATA_MATRIX(i, j)) Then
            
            VAR_STR = PARSE_POLYNOMIAL_STRING_FUNC(DATA_MATRIX(i, j), 1, "§")
            CTEMP_VECTOR = PARSE_POLYNOMIAL_STRING_FUNC(DATA_MATRIX(i, j), 0, "§")
            'check linear elements, i.e a+b*k
            If UBound(CTEMP_VECTOR, 1) > 1 Then
                ERROR_STR = "Element" & i & "," & j & " is not linear"
                GoTo ERROR_LABEL
            End If
            'check number of parameter. Only one parameter here
            If VAR_STR <> "" Then
                If TEMP_STR = "" Then
                    TEMP_STR = VAR_STR
                Else
                    If TEMP_STR <> VAR_STR Then
                        ERROR_STR = "too many parameters"
                        GoTo ERROR_LABEL
                    End If
                End If
            End If
            'all ok, substitute the the element ij with the polynomial index
            k = k + 1
            DATA_MATRIX(i, j) = "i" & k  'index
            DTEMP_VECTOR(k, 1) = CTEMP_VECTOR(0, 1)
            DTEMP_VECTOR(k, 2) = CTEMP_VECTOR(1, 1)
        End If
    Next j
Next i

ReDim TEMP_MATRIX(1 To NSIZE, 1 To NSIZE)

'build the determinant vector
ReDim FTEMP_VECTOR(1 To NDEG + 1, 1 To 1)
ReDim ETEMP_VECTOR(1 To NDEG + 1, 1 To 1)

h = 1 'step
For k = 1 To NDEG + 1
    ETEMP_VECTOR(k, 1) = (k - 1) * h
    'convert the symbolic elements into numeric
    For i = 1 To NSIZE
        For j = 1 To NSIZE
            If Not IsNumeric(DATA_MATRIX(i, j)) Then
                kk = CLng(Right(DATA_MATRIX(i, j), _
                            Len(DATA_MATRIX(i, j)) - 1))
                TEMP_MATRIX(i, j) = DTEMP_VECTOR(kk, 1) + _
                    DTEMP_VECTOR(kk, 2) * ETEMP_VECTOR(k, 1)
            Else
                TEMP_MATRIX(i, j) = DATA_MATRIX(i, j)
            End If
        Next j
    Next i
    FTEMP_VECTOR(k, 1) = MATRIX_SQUARE_GJ_DETERM_FUNC(TEMP_MATRIX, 0, epsilon)
    'FTEMP_VECTOR(k, 1) = WorksheetFunction.MDeterm(TEMP_MATRIX)
Next k
'build the Vandermonde matrix
ReDim TEMP_MATRIX(1 To NDEG + 1, 1 To NDEG + 1)

For i = 1 To NDEG + 1
    For j = 1 To NDEG + 1
        TEMP_MATRIX(i, j) = ETEMP_VECTOR(i, 1) ^ (j - 1)
    Next j
Next i

'solve the linear system
ReDim ETEMP_VECTOR(1 To NDEG + 1, 1 To 1)
For i = 1 To NDEG + 1
    ETEMP_VECTOR(i, 1) = FTEMP_VECTOR(i, 1)
Next i

ETEMP_VECTOR = MATRIX_GS_REDUCTION_PIVOT_FUNC(TEMP_MATRIX, ETEMP_VECTOR, epsilon, 0)

ReDim GTEMP_VECTOR(0 To NDEG, 1 To 1)
ReDim DTEMP_VECTOR(0 To NDEG, 1 To 1)

For i = 1 To UBound(ETEMP_VECTOR, 1)
    If INT_FLAG Then
        DTEMP_VECTOR(i - 1, 1) = Round(ETEMP_VECTOR(i, 1), 0)
        GTEMP_VECTOR(i - 1, 1) = DTEMP_VECTOR(i - 1, 1)
    Else
        DTEMP_VECTOR(i - 1, 1) = ROUND_FUNC(ETEMP_VECTOR(i, 1), 6, 0)
        GTEMP_VECTOR(i - 1, 1) = ROUND_FUNC(ETEMP_VECTOR(i, 1), 10, 0)
    End If
Next i

Select Case OUTPUT
    Case 0
        POLYNOMIAL_PARAMETRIC_DETERM_MATRIX_FUNC = _
            WRITE_POLYNOMIAL_STRING_FUNC(DTEMP_VECTOR, VAR_STR)
    Case Else
        POLYNOMIAL_PARAMETRIC_DETERM_MATRIX_FUNC = _
            GTEMP_VECTOR
End Select

Exit Function
ERROR_LABEL:
POLYNOMIAL_PARAMETRIC_DETERM_MATRIX_FUNC = ERROR_STR
End Function

