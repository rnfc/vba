Attribute VB_Name = "NUMBER_REAL_SEQUENCE_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.

'************************************************************************************
'************************************************************************************
'FUNCTION      : FIBONACCI_SEQUENCE_FUNC
'DESCRIPTION   : Fibonacci Sequence
'LIBRARY       : NUMBER_REAL
'GROUP         : SEQUENCE
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/08/2008
'************************************************************************************
'************************************************************************************

Function FIBONACCI_SEQUENCE_FUNC(ByVal NSIZE As Long, _
Optional ByVal INIT_VAL As Double = 0)

Dim i As Long
Dim TEMP_MATRIX As Variant

On Error GoTo ERROR_LABEL

ReDim TEMP_MATRIX(0 To NSIZE, 1 To 4)

TEMP_MATRIX(0, 1) = "A"
TEMP_MATRIX(0, 2) = "B"
TEMP_MATRIX(0, 3) = "F=A+B"
TEMP_MATRIX(0, 4) = "F/F(t+1)"

TEMP_MATRIX(1, 1) = INIT_VAL
TEMP_MATRIX(1, 2) = 1
TEMP_MATRIX(1, 3) = TEMP_MATRIX(1, 1) + TEMP_MATRIX(1, 2)

For i = 2 To NSIZE
    TEMP_MATRIX(i, 1) = TEMP_MATRIX(i - 1, 2)
    TEMP_MATRIX(i, 2) = TEMP_MATRIX(i - 1, 3)
    TEMP_MATRIX(i, 3) = TEMP_MATRIX(i, 1) + TEMP_MATRIX(i, 2)
    TEMP_MATRIX(i - 1, 4) = TEMP_MATRIX(i - 1, 3) / TEMP_MATRIX(i, 3)
Next i

TEMP_MATRIX(NSIZE, 4) = "-"

FIBONACCI_SEQUENCE_FUNC = TEMP_MATRIX

Exit Function
ERROR_LABEL:
FIBONACCI_SEQUENCE_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : CORPUT_SEQUENCE_NUMBER_FUNC
'DESCRIPTION   : Returns the equivalent first van der Corput sequence number
'LIBRARY       : REAL
'GROUP         : SEQUENCE
'ID            : 002
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/08/2008
'************************************************************************************
'************************************************************************************

Function CORPUT_SEQUENCE_NUMBER_FUNC(ByVal N1_VAL As Long, _
ByVal N2_VAL As Long)

'If N2_VAL = 2 then it returns the equivalent first van der
'Corput sequence number
    
Dim i As Double
Dim j As Double
Dim k As Double

Dim ii As Double
Dim jj As Double

On Error GoTo ERROR_LABEL

j = N1_VAL
ii = 0
jj = 1 / N2_VAL
Do While j > 0
    k = Int(j / N2_VAL)
    i = j - k * N2_VAL
    ii = ii + jj * i
    jj = jj / N2_VAL
    j = k
Loop

CORPUT_SEQUENCE_NUMBER_FUNC = ii

Exit Function
ERROR_LABEL:
CORPUT_SEQUENCE_NUMBER_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : HALTON_SEQUENCE_FUNC
'DESCRIPTION   : Returns the equivalent first Halton sequence number
'LIBRARY       : NUMBERS
'GROUP         : FACTORIAL
'ID            : 003
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/08/2008
'************************************************************************************
'************************************************************************************

Function HALTON_SEQUENCE_FUNC(ByVal NSIZE As Long)
    
Dim i As Long
Dim j As Long
Dim k As Long

Dim ii As Double
Dim jj As Double

On Error GoTo ERROR_LABEL

j = NSIZE
ii = 0
jj = 1 / 2
Do While j > 0
    k = Int(j / 2)
    i = j - k * 2
    ii = ii + jj * i
    jj = jj / 2
    j = k
Loop
HALTON_SEQUENCE_FUNC = ii

Exit Function
ERROR_LABEL:
HALTON_SEQUENCE_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : CONVEX_HULL_FUNC

'DESCRIPTION   : This convex hull algorithm is very useful for combinatorics
'asset allocation

'LIBRARY       : NUMBERS
'GROUP         : FACTORIAL
'ID            : 004
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/08/2008
'************************************************************************************
'************************************************************************************


Function CONVEX_HULL_FUNC(ByRef DATA1_RNG As Variant, _
ByRef DATA2_RNG As Variant)

Dim i As Long '
Dim j As Long '

Dim hh As Long '
Dim ii As Long '
Dim jj As Long '
Dim kk As Long '

Dim NSIZE As Long

Dim X1_VAL As Double
Dim X2_VAL As Double
Dim X3_VAL As Double

Dim Y1_VAL As Double
Dim Y2_VAL As Double
Dim Y3_VAL As Double

Dim TEMP1_VAL As Double
Dim TEMP2_VAL As Double

Dim TEMP1_VECTOR As Variant
Dim TEMP2_VECTOR As Variant

Dim DATA1_VECTOR As Variant 'Returns
Dim DATA2_VECTOR As Variant 'Sigma

On Error GoTo ERROR_LABEL

DATA1_VECTOR = DATA1_RNG
If UBound(DATA1_VECTOR, 1) = 1 Then
    DATA1_VECTOR = MATRIX_TRANSPOSE_FUNC(DATA1_VECTOR)
End If

DATA2_VECTOR = DATA2_RNG
If UBound(DATA2_VECTOR, 1) = 1 Then
    DATA2_VECTOR = MATRIX_TRANSPOSE_FUNC(DATA2_VECTOR)
End If

If UBound(DATA1_VECTOR, 1) <> UBound(DATA2_VECTOR, 1) Then: GoTo ERROR_LABEL

NSIZE = UBound(DATA1_VECTOR, 1)
ii = MATRIX_FIND_ELEMENT_FUNC(DATA1_VECTOR, MATRIX_ELEMENTS_MAX_FUNC(DATA1_VECTOR, 0), 1, 1, 1)
jj = MATRIX_FIND_ELEMENT_FUNC(DATA2_VECTOR, MATRIX_ELEMENTS_MIN_FUNC(DATA2_VECTOR, 0), 1, 1, 1)

ReDim TEMP1_VECTOR(1 To NSIZE, 1 To 1)
j = 1
TEMP1_VECTOR(j, 1) = jj
hh = jj
Do While 1 = 1
    TEMP1_VAL = 0
    kk = ii
    
    X1_VAL = DATA2_VECTOR(hh, 1)
    X2_VAL = DATA2_VECTOR(ii, 1)
    
    Y1_VAL = DATA1_VECTOR(hh, 1)
    Y2_VAL = DATA1_VECTOR(ii, 1)
    
    For i = 1 To NSIZE
        X3_VAL = DATA2_VECTOR(i, 1)
        Y3_VAL = DATA1_VECTOR(i, 1)
        TEMP2_VAL = ((X2_VAL - X1_VAL) * (Y3_VAL - Y1_VAL) - (X3_VAL - X1_VAL) * (Y2_VAL - Y1_VAL)) / ((X3_VAL ^ 2 + Y3_VAL ^ 2) * (X1_VAL ^ 2 + Y2_VAL ^ 2) ^ 0.5)
        If TEMP2_VAL > TEMP1_VAL Then
            TEMP1_VAL = TEMP2_VAL
            kk = i
        End If
    Next i
    
    If TEMP2_VAL = 0 Then: Exit Do
    j = j + 1
    TEMP1_VECTOR(j, 1) = kk
    hh = kk
Loop

ReDim TEMP2_VECTOR(1 To j, 1 To 1)
For i = 1 To j
    TEMP2_VECTOR(i, 1) = TEMP1_VECTOR(i, 1)
Next i

CONVEX_HULL_FUNC = TEMP2_VECTOR

Exit Function
ERROR_LABEL:
CONVEX_HULL_FUNC = Err.number
End Function
