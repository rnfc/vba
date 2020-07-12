Attribute VB_Name = "NUMBER_REAL_FACTORIAL_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.

'************************************************************************************
'************************************************************************************
'FUNCTION      : FACTORIAL_FUNC
'DESCRIPTION   : Returns the factorial of a number. The factorial of a
'number is equal to 1*2*3*...* number.

'LIBRARY       : NUMBERS
'GROUP         : FACTORIAL
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/08/2008
'************************************************************************************
'************************************************************************************

Function FACTORIAL_FUNC(ByVal X_VAL As Double) 'Limit 196 Temp Numbers

Dim i As Long
Dim NROWS As Long

On Error GoTo ERROR_LABEL

If X_VAL < 0 Then: GoTo ERROR_LABEL

FACTORIAL_FUNC = 1
NROWS = Int(X_VAL)

For i = 1 To NROWS
    FACTORIAL_FUNC = FACTORIAL_FUNC * i
Next i
      
Exit Function
ERROR_LABEL:
FACTORIAL_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : COMBINATIONS_FUNC

'DESCRIPTION   : Combinations j objects, k classes
'Returns the number of combinations for a given number of items.
'Use COMBIN to determine the total possible number of groups for
'a given number of items.

'LIBRARY       : NUMBERS
'GROUP         : FACTORIAL
'ID            : 002
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/08/2008
'************************************************************************************
'************************************************************************************

Function COMBINATIONS_FUNC(ByVal NUMBER_VAL As Double, _
ByVal FACTOR_VAL As Double)
    
'NUMBER_VAL: is the number of items.
'FACTOR_VAL: Number chosen is the number of items in each combination.
    
Dim i As Long
Dim j As Long
Dim k As Long
Dim TEMP_VAL As Variant

On Error GoTo ERROR_LABEL
    
If (NUMBER_VAL < 0) Or (FACTOR_VAL < 0) Then
    TEMP_VAL = 0 'raises in error
End If
    
j = Int(NUMBER_VAL)
k = Int(FACTOR_VAL)
        
If (j < 1) Or (k < 1) Or (k > j) Then
    COMBINATIONS_FUNC = 0
    Exit Function
End If
    
If (k = j) Or (k = 0) Then
    COMBINATIONS_FUNC = 1
    Exit Function
End If
    
TEMP_VAL = j
If k > Int(j / 2) Then k = j - k
For i = 2 To k
    TEMP_VAL = TEMP_VAL * (j + 1 - i) / i
Next i
COMBINATIONS_FUNC = TEMP_VAL

Exit Function
ERROR_LABEL:
COMBINATIONS_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : PERMUTATIONS_FUNC

'DESCRIPTION   : Permutations j objects, k classes
'Returns the number of permutations for a given number of objects
'that can be selected from number objects. A permutation is any set
'or subset of objects or events where internal order is significant.
'Permutations are different from combinations, for which the internal
'order is not significant. Use this function for lottery-style
'probability calculations.

'LIBRARY       : NUMBERS
'GROUP         : FACTORIAL
'ID            : 003
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/08/2008
'************************************************************************************
'************************************************************************************


Function PERMUTATIONS_FUNC(ByVal NUMBER_VAL As Double, _
ByVal FACTOR_VAL As Double)
    
'NUMBER_VAL: is an integer that describes the number of objects.
'FACTOR_VAL: is an integer that describes the number of objects
'in each permutation.
    
Dim i As Long
Dim j As Long
Dim k As Long
Dim TEMP_VAL As Double

On Error GoTo ERROR_LABEL

If NUMBER_VAL < 0 Or FACTOR_VAL < 0 Then: TEMP_VAL = 0
'raises in error
j = Int(NUMBER_VAL)
k = Int(FACTOR_VAL)
If j < 1 Or k < 1 Or k > j Then
    PERMUTATIONS_FUNC = 0
    Exit Function
End If

TEMP_VAL = j
For i = 2 To k
    TEMP_VAL = TEMP_VAL * (j + 1 - i)
Next i

PERMUTATIONS_FUNC = TEMP_VAL

Exit Function
ERROR_LABEL:
PERMUTATIONS_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : COMBINATIONS_ELEMENTS_FUNC

'DESCRIPTION   : Combinations containing k elements drawn from a total of n elements
' without replacement and irrelevant order
'LIBRARY       : NUMBERS
'GROUP         : FACTORIAL
'ID            : 004
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/08/2008
'************************************************************************************
'************************************************************************************


Function COMBINATIONS_ELEMENTS_FUNC(ByVal n As Long, _
ByVal k As Long)

Dim A As Long
Dim B As Long
Dim D As Long

Dim h As Long
Dim i As Long
Dim j As Long

Dim E As Long
Dim f As Long

Dim m As Long

Dim o As Long
Dim p As Long
Dim q As Long

Dim TEMP_MATRIX As Variant

On Error GoTo ERROR_LABEL

o = n + 1
p = 1
For i = o - k To n
    p = p * i
Next i
q = 1
For i = 1 To k
    q = q * i
Next i
m = p / q
ReDim TEMP_MATRIX(1 To m, 1 To k)
f = o - k
For i = 1 To f
    TEMP_MATRIX(i, k) = k + i - 1
Next i
For A = k - 1 To 1 Step -1
    D = f
    h = f
    For i = 1 To f
        TEMP_MATRIX(i, A) = A
    Next i
    For B = A + 1 To A + n - k
        D = D * (o + A - B - k) / (o - B)
        E = f + 1
        f = E + D - 1
        For i = E To f
            TEMP_MATRIX(i, A) = B
        Next i
        For i = 0 To f - E
            For j = 0 To k - (A + 1)
                TEMP_MATRIX(E + i, A + 1 + j) = TEMP_MATRIX(h - D + 1 + i, A + 1 + j)
            Next j
        Next i
    Next B
Next A

COMBINATIONS_ELEMENTS_FUNC = TEMP_MATRIX

Exit Function
ERROR_LABEL:
COMBINATIONS_ELEMENTS_FUNC = Err.number
End Function
