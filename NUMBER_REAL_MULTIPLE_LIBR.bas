Attribute VB_Name = "NUMBER_REAL_MULTIPLE_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.


'************************************************************************************
'************************************************************************************
'FUNCTION      : MOD_FUNC
'DESCRIPTION   : Returns the remainder after number is divided by divisor.
'The result has the same sign as divisor.
'LIBRARY       : NUMBER_REAL
'GROUP         : MULTIPLE
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/08/2008
'************************************************************************************
'************************************************************************************

Function MOD_FUNC(ByVal FIRST_VAL As Double, _
ByVal SECOND_VAL As Double)
'FIRST_VAL - SECOND_VAL * Int(FIRST_VAL / SECOND_VAL)

'FIRST_VAL: is the number for which you want to find the remainder.
'SECOND_VAL: is the number by which you want to divide number.

'Dim ATEMP_VAL As Variant
'Dim BTEMP As Variant
'Dim CTEMP As Variant
    
On Error GoTo ERROR_LABEL

'    ATEMP_VAL = Int(Abs(FIRST_VAL))
'    BTEMP = Int(Abs(SECOND_VAL))
'    CTEMP = Round(ATEMP_VAL - BTEMP * Int(ATEMP_VAL / BTEMP), 0)
'    If FIRST_VAL < 0 Then CTEMP = BTEMP - CTEMP
'    MOD_FUNC = CTEMP

    If FIRST_VAL >= 0 Then
        MOD_FUNC = FIRST_VAL - SECOND_VAL * _
            Int(FIRST_VAL / SECOND_VAL)
    Else
        MOD_FUNC = FIRST_VAL + SECOND_VAL * _
                  (Int(-FIRST_VAL / SECOND_VAL) + 1)
    End If
  
Exit Function
ERROR_LABEL:
MOD_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : MODULUS_FUNC
'DESCRIPTION   : Returns the remainder after number is divided by divisor.
'The result has the same sign as divisor.
'LIBRARY       : NUMBER_REAL
'GROUP         : MULTIPLE
'ID            : 002
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/08/2008
'************************************************************************************
'************************************************************************************

Function MODULUS_FUNC(ByVal FIRST_VAL As Double, _
ByVal SECOND_VAL As Double)
On Error GoTo ERROR_LABEL
MODULUS_FUNC = Round(FIRST_VAL - SECOND_VAL * _
               Int(FIRST_VAL / SECOND_VAL), 0)
Exit Function
ERROR_LABEL:
MODULUS_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : LCM_FUNC
'DESCRIPTION   : Returns the least common multiple of integers. The least
'common multiple is the smallest positive integer that is
'a multiple of all integer arguments number1, number2, and
'so on. Use LCM to add fractions with different denominators.

'LIBRARY       : NUMBER_REAL
'GROUP         : MULTIPLE
'ID            : 003
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/08/2008
'************************************************************************************
'************************************************************************************

Function VECTOR_LCM_FUNC(ByRef DATA_RNG As Variant)
'LCM of NROWS-numbers

Dim i As Long
Dim NROWS As Long
Dim X_VAL As Double
Dim TEMP_VECTOR As Variant

On Error GoTo ERROR_LABEL

TEMP_VECTOR = DATA_RNG
If UBound(TEMP_VECTOR, 1) = 1 Then
    TEMP_VECTOR = MATRIX_TRANSPOSE_FUNC(TEMP_VECTOR)
End If

NROWS = UBound(TEMP_VECTOR, 1)
X_VAL = TEMP_VECTOR(1, 1)
For i = 2 To NROWS
    X_VAL = PAIR_LCM_FUNC(X_VAL, TEMP_VECTOR(i, 1))
Next i
VECTOR_LCM_FUNC = X_VAL
  
Exit Function
ERROR_LABEL:
VECTOR_LCM_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : PAIR_LCM_FUNC
'DESCRIPTION   : Fix the LCM between two integer numbers
'LIBRARY       : NUMBER_REAL
'GROUP         : MULTIPLE
'ID            : 004
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/08/2008
'************************************************************************************
'************************************************************************************

Function PAIR_LCM_FUNC(ByVal FIRST_VAL As Double, _
ByVal SECOND_VAL As Double)
  
Dim ATEMP_VAL As Double
Dim BTEMP_VAL As Double
  
On Error GoTo ERROR_LABEL
  
BTEMP_VAL = Int(Abs(FIRST_VAL))
ATEMP_VAL = Int(Abs(SECOND_VAL))
  
PAIR_LCM_FUNC = Round(ATEMP_VAL * BTEMP_VAL / _
                PAIR_GCD_FUNC(ATEMP_VAL, BTEMP_VAL), 0)
  
Exit Function
ERROR_LABEL:
PAIR_LCM_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : PAIR_MCM_FUNC
'DESCRIPTION   : Find the mcm between two integer numbers
'LIBRARY       : NUMBER_REAL
'GROUP         : MULTIPLE
'ID            : 005
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/08/2008
'************************************************************************************
'************************************************************************************

Function PAIR_MCM_FUNC(ByVal FIRST_VAL As Double, _
ByVal SECOND_VAL As Double)

Dim TEMP_VAL As Double
On Error GoTo ERROR_LABEL

TEMP_VAL = PAIR_GCD_FUNC(FIRST_VAL, SECOND_VAL)
PAIR_MCM_FUNC = FIRST_VAL * SECOND_VAL / TEMP_VAL

Exit Function
ERROR_LABEL:
PAIR_MCM_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : PAIR_GCD_FUNC
'DESCRIP ByVal REFER_STR As String = "NAME*")
'integers. The greatest common divisor is the largest
'integer that divides both number1 and number2 without
'a remainder.
'LIBRARY       : NUMBER_REAL
'GROUP         : MULTIPLE
'ID            : 006
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/08/2008
'************************************************************************************
'************************************************************************************

Function PAIR_GCD_FUNC(ByVal FIRST_VAL As Double, _
ByVal SECOND_VAL As Double)

' Fix the GCD between two integer numbers

Dim ATEMP_VAL As Double
Dim BTEMP_VAL As Double
Dim CTEMP_VAL As Double
  
On Error GoTo ERROR_LABEL
  
  BTEMP_VAL = Int(Abs(FIRST_VAL))
  ATEMP_VAL = Int(Abs(SECOND_VAL))
  Do Until ATEMP_VAL = 0
    CTEMP_VAL = MOD_FUNC(BTEMP_VAL, ATEMP_VAL)
    BTEMP_VAL = ATEMP_VAL
    ATEMP_VAL = CTEMP_VAL
  Loop
  
  PAIR_GCD_FUNC = BTEMP_VAL

Exit Function
ERROR_LABEL:
PAIR_GCD_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : GCD_FUNC
'DESCRIPTION   : GCD of NROWS-numbers
'LIBRARY       : NUMBER_REAL
'GROUP         : MULTIPLE
'ID            : 007
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/08/2008
'************************************************************************************
'************************************************************************************

Function VECTOR_GCD_FUNC(ByRef DATA_RNG As Variant)

Dim i As Long
Dim NROWS As Long
Dim TEMP_VAL As Double
Dim TEMP_VECTOR As Variant

On Error GoTo ERROR_LABEL

TEMP_VECTOR = DATA_RNG
If UBound(TEMP_VECTOR, 1) = 1 Then
    TEMP_VECTOR = MATRIX_TRANSPOSE_FUNC(TEMP_VECTOR)
End If
NROWS = UBound(TEMP_VECTOR, 1)

TEMP_VAL = TEMP_VECTOR(1, 1)
For i = 2 To NROWS
    TEMP_VAL = PAIR_GCD_FUNC(TEMP_VAL, TEMP_VECTOR(i, 1))
Next i

VECTOR_GCD_FUNC = TEMP_VAL
  
Exit Function
ERROR_LABEL:
VECTOR_GCD_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : REDUCED_POWER_FUNC
'DESCRIPTION   : Reduced Function with Threshold
'LIBRARY       : NUMBER_REAL
'GROUP         : MULTIPLE
'ID            : 008
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/08/2008
'************************************************************************************
'************************************************************************************

Function REDUCED_POWER_FUNC(ByVal X_VAL As Double, _
Optional ByVal nLOOPS As Long = 200)
Dim i As Long
On Error GoTo ERROR_LABEL
i = 0
Do While X_VAL Mod 2 = 0
    X_VAL = X_VAL / 2
    i = i + 1
    If i > nLOOPS Then: Exit Do
Loop
If X_VAL = 1 Then
    REDUCED_POWER_FUNC = i
Else
    REDUCED_POWER_FUNC = X_VAL
End If
Exit Function
ERROR_LABEL:
REDUCED_POWER_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : BOUND_FUNC
'DESCRIPTION   : BOUNDING CONSTRAINT FUNCTION
'LIBRARY       : NUMBER_REAL
'GROUP         : MULTIPLE
'ID            : 009
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/08/2008
'************************************************************************************
'************************************************************************************

Function BOUND_FUNC(ByVal THRESHOLD As Double, _
ByVal MIN_VALUE As Double, _
ByVal MAX_VALUE As Double)
  
On Error GoTo ERROR_LABEL
  
If THRESHOLD < MIN_VALUE Then
    BOUND_FUNC = MIN_VALUE
ElseIf THRESHOLD > MAX_VALUE Then
    BOUND_FUNC = MAX_VALUE
Else: BOUND_FUNC = THRESHOLD
End If

Exit Function
ERROR_LABEL:
BOUND_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : BARRIER_FUNC
'DESCRIPTION   : Multiply a number until a threshold is reached
'LIBRARY       : NUMBER_REAL
'GROUP         : MULTIPLE
'ID            : 010
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/08/2008
'************************************************************************************
'************************************************************************************

Function BARRIER_FUNC(ByVal START_VAL As Double, _
ByVal MULTIPLIER As Double, _
ByVal THRESHOLD As Double, _
Optional ByVal nLOOPS As Long = 10000)

Dim j As Long
Dim TEMP_MULT As Double

On Error GoTo ERROR_LABEL

j = 0
TEMP_MULT = START_VAL
Do Until TEMP_MULT >= THRESHOLD
    TEMP_MULT = TEMP_MULT * MULTIPLIER
    j = j + 1
    If j > nLOOPS Then: Exit Do
Loop

BARRIER_FUNC = TEMP_MULT

Exit Function
ERROR_LABEL:
BARRIER_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : MULTIPLE_FUNC
'DESCRIPTION   : Find the first Multiple of FIRST_VAL = SECOND_VAL
'LIBRARY       : NUMBER_REAL
'GROUP         : MULTIPLE
'ID            : 011
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/08/2008
'************************************************************************************
'************************************************************************************

Function MULTIPLE_FUNC(ByVal FIRST_VAL As Double, _
ByVal SECOND_VAL As Double, _
Optional ByVal nLOOPS As Long = 10000)

Dim j As Long
Dim TEMP_SUM As Double

On Error GoTo ERROR_LABEL

If FIRST_VAL < 0 Then: GoTo ERROR_LABEL

j = 0
TEMP_SUM = FIRST_VAL
Do While TEMP_SUM < SECOND_VAL
    TEMP_SUM = TEMP_SUM + FIRST_VAL
    j = j + 1
    If j > nLOOPS Then: Exit Do
Loop

MULTIPLE_FUNC = TEMP_SUM

Exit Function
ERROR_LABEL:
MULTIPLE_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : CLIP_FUNC
'DESCRIPTION   : Clipping function
'LIBRARY       : NUMBER_REAL
'GROUP         : MULTIPLE
'ID            : 012
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/08/2008
'************************************************************************************
'************************************************************************************

Function CLIP_FUNC(ByVal X_VAL As Double, _
ByVal FLOOR_VAL As Double, _
ByVal CEEILING_VAL As Double)

On Error GoTo ERROR_LABEL

If FLOOR_VAL > CEEILING_VAL Then: GoTo ERROR_LABEL 'raise an error

If X_VAL < FLOOR_VAL Then
    CLIP_FUNC = FLOOR_VAL
ElseIf X_VAL > CEEILING_VAL Then
    CLIP_FUNC = CEEILING_VAL
Else
    CLIP_FUNC = X_VAL
End If

Exit Function
ERROR_LABEL:
CLIP_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : MCD2_FUNC
'DESCRIPTION   : Find the MCD between two integer numbers by the Euclid method
'LIBRARY       : NUMBER_REAL
'GROUP         : MULTIPLE
'ID            : 013
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/08/2008
'************************************************************************************
'************************************************************************************

Function MCD2_FUNC(ByVal FIRST_VAL As Double, _
ByVal SECOND_VAL As Double)

Dim Y_VAL As Double
Dim X_VAL As Double
Dim Z_VAL As Double

On Error GoTo ERROR_LABEL

Y_VAL = FIRST_VAL
X_VAL = SECOND_VAL

Do Until X_VAL = 0
    Z_VAL = Y_VAL - X_VAL * Int(Y_VAL / X_VAL)
    Y_VAL = X_VAL
    X_VAL = Z_VAL
Loop

MCD2_FUNC = Y_VAL

Exit Function
ERROR_LABEL:
MCD2_FUNC = Err.number
End Function
