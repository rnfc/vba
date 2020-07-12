Attribute VB_Name = "NUMBER_REAL_PRIME_LIBR"


Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.
    

'************************************************************************************
'************************************************************************************
'FUNCTION      : PRIME_FUNC

'DESCRIPTION   : States whether a NUMBER_VAL is prime or returns the lowest factor
'of a NUMBER_VAL. Works by trial division. Increase the nloops to find bigger
'primes - the function will be slower.

'LIBRARY       : NUMBER
'GROUP         : REAL_PRIME
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/27/2009
'************************************************************************************
'************************************************************************************
    
Function PRIME_FUNC(ByVal NUMBER_VAL As Variant, _
Optional ByVal nLOOPS As Long = 500000)

'nLOOPS --> a limit on the size of divisor it will try

Dim i As Long   'the primes to be used as divisors
Dim SQROOT_VAL  As Double   'the square root of the NUMBER_VAL

On Error GoTo ERROR_LABEL
If NUMBER_VAL <> Int(NUMBER_VAL) Or NUMBER_VAL < 1 Then GoTo ERROR_LABEL

If NUMBER_VAL = 1 Then
    PRIME_FUNC = 1
    Exit Function
ElseIf NUMBER_VAL = 2 Then
    PRIME_FUNC = "P"
    Exit Function
End If

If NUMBER_VAL / 2 = Int(NUMBER_VAL / 2) Then 'Using "Mod" would be an obvious choice
'here but it limits the NUMBER_VAL to 2,147,483,000.
    PRIME_FUNC = 2
    Exit Function
End If

i = 3
SQROOT_VAL = (NUMBER_VAL ^ 0.5)

Do While i <= SQROOT_VAL
    If NUMBER_VAL / i = Int(NUMBER_VAL / i) Then
        PRIME_FUNC = i
        Exit Function
    End If
    i = i + 2
    If i > nLOOPS Then
        PRIME_FUNC = "?"
        Exit Function
    End If
Loop

PRIME_FUNC = "P"

Exit Function
ERROR_LABEL:
PRIME_FUNC = "?"
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : NEXT_PRIME_FUNC
'DESCRIPTION   : Returns the next prime bigger than any NUMBER_VAL.
'LIBRARY       : NUMBER
'GROUP         : REAL_PRIME
'ID            : 002
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/27/2009
'************************************************************************************
'************************************************************************************

Function NEXT_PRIME_FUNC(ByVal NUMBER_VAL As Variant, _
Optional ByVal nLOOPS As Long = 100000)

'nLOOPS --> a limit on the size of divisor it will try

Dim i As Long
Dim SQROOT_VAL As Double
Dim POSS_PRIME_VAL As Double

On Error GoTo ERROR_LABEL

If NUMBER_VAL <> Int(NUMBER_VAL) Or NUMBER_VAL < 1 Then GoTo ERROR_LABEL

If NUMBER_VAL = 1 Then
    NEXT_PRIME_FUNC = 2
    Exit Function
End If

If NUMBER_VAL / 2 = Int(NUMBER_VAL / 2) Then
    POSS_PRIME_VAL = NUMBER_VAL + 1
Else
    POSS_PRIME_VAL = NUMBER_VAL + 2
End If

SQROOT_VAL = POSS_PRIME_VAL ^ 0.5
i = 3

Do Until i > SQROOT_VAL
    i = 3
    Do Until i > SQROOT_VAL
        If POSS_PRIME_VAL / i = Int(POSS_PRIME_VAL / i) Then
            POSS_PRIME_VAL = POSS_PRIME_VAL + 2
            SQROOT_VAL = POSS_PRIME_VAL ^ 0.5
            Exit Do
        End If
        i = i + 2
        If i > nLOOPS Then
            NEXT_PRIME_FUNC = "?"
            Exit Function
        End If
    Loop
Loop
NEXT_PRIME_FUNC = POSS_PRIME_VAL
Exit Function
ERROR_LABEL:
NEXT_PRIME_FUNC = "?"
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : ERATOSTHENE_FUNC
'DESCRIPTION   : Sieve of Eratosthenes is a simple, ancient algorithm for finding all
'prime numbers up to a specified integer. It works efficiently for the smaller
'primes (below 10 million). It was created by Eratosthenes, an ancient Greek mathematician.
'When the Sieve of Eratosthenes is used in computer programming, wheel factorization
'is often applied before the sieve to increase the speed.

'LIBRARY       : NUMBER
'GROUP         : REAL_PRIME
'ID            : 003
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/27/2009
'************************************************************************************
'************************************************************************************

Function ERATOSTHENE_FUNC(ByVal MIN_PRIME_VAL As Long, _
ByVal MAX_PRIME_VAL As Long)

'h = 3,5,7,9,11,13,
'h = i*2+1
'i = 1,2,3,4,5,
'i = (h-1)/2

Dim h As Long
Dim i As Long
Dim j As Long
Dim k As Long
Dim l As Long

Dim SQRT_VAL As Double
Dim TEMP_VECTOR As Variant
'Dim PRIME_FLAG_ARR() As Boolean

On Error GoTo ERROR_LABEL

l = (MAX_PRIME_VAL - 1) / 2
'ReDim PRIME_FLAG_ARR(1 To l) As Boolean
ReDim TEMP_VECTOR(1 To l, 1 To 2)

If MIN_PRIME_VAL < 2 Then k = k + 1: TEMP_VECTOR(k, 1) = 1
If MIN_PRIME_VAL < 3 Then k = k + 1: TEMP_VECTOR(k, 1) = 2
If MIN_PRIME_VAL < 4 Then k = k + 1: TEMP_VECTOR(k, 1) = 3

For i = 1 To l
    'PRIME_FLAG_ARR(i) = True
    TEMP_VECTOR(i, 2) = True
Next i
SQRT_VAL = Sqr(MAX_PRIME_VAL)
h = 3
Do While h < MAX_PRIME_VAL
    If h <= SQRT_VAL Then
    j = h * h
    Do While j <= MAX_PRIME_VAL
        TEMP_VECTOR((j - 1) / 2, 2) = False
        'PRIME_FLAG_ARR((j - 1) / 2) = False
        j = j + 2 * h
    Loop
    End If
    h = h + 2
    If h >= MAX_PRIME_VAL Then
        'MsgBox (CStr(k) + " primes found")
        Exit Do
    End If
'    Do While Not PRIME_FLAG_ARR((h - 1) / 2)
    Do While Not TEMP_VECTOR((h - 1) / 2, 2)
        h = h + 2
        If h >= MAX_PRIME_VAL Then
            'MsgBox (CStr(k) + " primes found")
            Exit Do
        End If
    Loop
    If h > MIN_PRIME_VAL Then
        k = k + 1
        TEMP_VECTOR(k, 1) = h
    End If
Loop

'ReDim Preserve TEMP_VECTOR(1 To k, 1 TO 2)

ERATOSTHENE_FUNC = MATRIX_TRIM_FUNC(TEMP_VECTOR, 1, 0)

Exit Function
ERROR_LABEL:
ERATOSTHENE_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : Totient Euler's function
'DESCRIPTION   :
'LIBRARY       : NUMBER
'GROUP         : REAL_PRIME
'ID            : 004
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/27/2009
'************************************************************************************
'************************************************************************************


Function TOTIENT_EULER_FUNC(ByVal NUMBER_VAL As Variant)

Dim i As Long
Dim TEMP_MULT As Double

Dim FACT_ARR() As Long
Dim EXPO_ARR() As Long
Dim ERROR_STR As String

If NUMBER_VAL <= 1 Then
    TOTIENT_EULER_FUNC = 1
    Exit Function
End If

On Error GoTo ERROR_LABEL
If FACTORIZE_INTEGER_FUNC(NUMBER_VAL, FACT_ARR, EXPO_ARR, _
                          ERROR_STR, 100000) = False Then
    GoTo ERROR_LABEL
End If

TEMP_MULT = 1
For i = 1 To UBound(FACT_ARR)
    TEMP_MULT = TEMP_MULT * (FACT_ARR(i) - 1) * FACT_ARR(i) ^ (EXPO_ARR(i) - 1)
Next i

TOTIENT_EULER_FUNC = TEMP_MULT

Exit Function
ERROR_LABEL:
TOTIENT_EULER_FUNC = "?"
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : MATRICES_FACTORS_FUNC
'DESCRIPTION   : Factorizes an integer number n returns a matrices of factors:
'[factor, exponent]
'LIBRARY       : NUMBER
'GROUP         : REAL_PRIME
'ID            : 005
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/27/2009
'************************************************************************************
'************************************************************************************

Function MATRICES_FACTORS_FUNC(ByVal NUMBER_VAL As Variant)

Dim i As Long
Dim NSIZE As Single

Dim FACT_ARR() As Long
Dim EXPO_ARR() As Long

Dim ERROR_STR As String
Dim TEMP_VECTOR As Variant

On Error GoTo ERROR_LABEL

If NUMBER_VAL > 0 Then
    Call FACTORIZE_INTEGER_FUNC(NUMBER_VAL, FACT_ARR, EXPO_ARR, ERROR_STR, 100000)
    NSIZE = UBound(FACT_ARR)
End If

'load an return array
ReDim TEMP_VECTOR(0 To NSIZE, 1 To 2)
TEMP_VECTOR(0, 1) = ("FACTOR")
TEMP_VECTOR(0, 2) = ("EXP")
For i = 1 To NSIZE
    TEMP_VECTOR(i, 1) = FACT_ARR(i)
    TEMP_VECTOR(i, 2) = EXPO_ARR(i)
Next i
MATRICES_FACTORS_FUNC = TEMP_VECTOR

Exit Function
ERROR_LABEL:
MATRICES_FACTORS_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : FACTORIZE_INTEGER_FUNC
'DESCRIPTION   : Factorizes an integer n
'LIBRARY       : NUMBER
'GROUP         : REAL_PRIME
'ID            : 006
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/27/2009
'************************************************************************************
'************************************************************************************

Private Function FACTORIZE_INTEGER_FUNC(ByVal NUMBER_VAL As Variant, _
ByRef FACT_ARR() As Long, _
ByRef EXPO_ARR() As Long, _
ByRef ERROR_STR As String, _
Optional ByVal nLOOPS As Long = 100000)

Dim i As Long

Dim M_VAL As Double
Dim F_VAL As Double

Dim A_VAL As Variant
Dim B_VAL As Variant
Dim C_VAL As Variant

On Error GoTo ERROR_LABEL

FACTORIZE_INTEGER_FUNC = False

If NUMBER_VAL <= 1 Then: GoTo ERROR_LABEL

ReDim FACT_ARR(1 To 100)
ReDim EXPO_ARR(1 To 100)
ERROR_STR = ""

i = 1

If NUMBER_VAL > 10 ^ 15 Then
    FACT_ARR(i) = "?"
    EXPO_ARR(i) = "?"
    ERROR_STR = "?"
    ReDim Preserve FACT_ARR(1 To 1)
    ReDim Preserve EXPO_ARR(1 To 1)
    Exit Function
End If

'try first with the brute force attack
M_VAL = NUMBER_VAL
C_VAL = NEXT_FACTOR_FUNC(M_VAL, 2, nLOOPS)

If IsNumeric(C_VAL) Then
    F_VAL = C_VAL
    Do While IsNumeric(C_VAL)
        If C_VAL <> F_VAL Then i = i + 1
        FACT_ARR(i) = C_VAL
        EXPO_ARR(i) = EXPO_ARR(i) + 1
        M_VAL = M_VAL / C_VAL
        F_VAL = C_VAL
        C_VAL = NEXT_FACTOR_FUNC(M_VAL, C_VAL, nLOOPS)
    Loop
    If C_VAL = "P" And M_VAL <> 1 Then
        If M_VAL <> F_VAL Then i = i + 1
        FACT_ARR(i) = M_VAL
        EXPO_ARR(i) = EXPO_ARR(i) + 1
        'GoTo Exit_
    End If
ElseIf C_VAL = "P" Then
    FACT_ARR(i) = M_VAL
    EXPO_ARR(i) = 1
End If

If C_VAL = "?" Then
    'try with fermat-lehman attak
    Call FACTOR_FERMAT_LEHMAN_FUNC(M_VAL, A_VAL, B_VAL)
    If B_VAL = "?" Then
        'no factor found
        If FACT_ARR(i) <> "" Then i = i + 1
        FACT_ARR(i) = "?"
        EXPO_ARR(i) = "?"
        ERROR_STR = "?"
    Else
        If B_VAL > 1 Then
            If FACT_ARR(i) <> B_VAL And FACT_ARR(i) <> "" Then i = i + 1
            FACT_ARR(i) = B_VAL
            EXPO_ARR(i) = EXPO_ARR(i) + 1
        End If
        If A_VAL > 1 Then
            If FACT_ARR(i) <> A_VAL And FACT_ARR(i) <> "" Then i = i + 1
            FACT_ARR(i) = A_VAL
            EXPO_ARR(i) = EXPO_ARR(i) + 1
        End If
    End If
End If

ReDim Preserve FACT_ARR(1 To i)
ReDim Preserve EXPO_ARR(1 To i)

FACTORIZE_INTEGER_FUNC = True

Exit Function
ERROR_LABEL:
FACTORIZE_INTEGER_FUNC = False
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : NEXT_FACTOR_FUNC
'DESCRIPTION   :
'LIBRARY       : NUMBER
'GROUP         : REAL_PRIME
'ID            : 007
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/27/2009
'************************************************************************************
'************************************************************************************

Private Function NEXT_FACTOR_FUNC(ByVal NUMBER_VAL As Variant, _
Optional ByVal LOWER_VAL As Long = 2, _
Optional ByVal UPPER_VAL As Long = 500000)

Dim i As Long   'the primes to be used as divisors
Dim SQROOT_VAL As Double   'the square root of the NUMBER_VAL

On Error GoTo ERROR_LABEL

If NUMBER_VAL <> Int(NUMBER_VAL) Or NUMBER_VAL < 1 Then GoTo ERROR_LABEL

If LOWER_VAL < 3 Then
    If NUMBER_VAL = 1 Then
        NEXT_FACTOR_FUNC = 1
        Exit Function
    ElseIf NUMBER_VAL = 2 Then
        NEXT_FACTOR_FUNC = "P"
        Exit Function
    End If
    If NUMBER_VAL / 2 = Int(NUMBER_VAL / 2) Then 'Using "Mod" would be an
    'obvious choice here but it limits the NUMBER_VAL to 2,147,483,000.
        NEXT_FACTOR_FUNC = 2
        Exit Function
    End If
End If

If LOWER_VAL / 2 = Int(LOWER_VAL / 2) Then i = LOWER_VAL + 1 Else i = LOWER_VAL

SQROOT_VAL = (NUMBER_VAL ^ 0.5)
Do While i <= SQROOT_VAL
    If NUMBER_VAL / i = Int(NUMBER_VAL / i) Then
        NEXT_FACTOR_FUNC = i
        Exit Function
    End If
    i = i + 2
    If i > UPPER_VAL Then
        NEXT_FACTOR_FUNC = "?"
        Exit Function
    End If
Loop

NEXT_FACTOR_FUNC = "P"
Exit Function
ERROR_LABEL:
NEXT_FACTOR_FUNC = "?"
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : FACTOR_FERMAT_LEHMAN_FUNC
'DESCRIPTION   : fermat-lehman algorithm
'LIBRARY       : NUMBER
'GROUP         : REAL_PRIME
'ID            : 008
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/27/2009
'************************************************************************************
'************************************************************************************

Private Function FACTOR_FERMAT_LEHMAN_FUNC(ByVal NUMBER_VAL As Variant, _
ByVal A_VAL As Variant, _
ByVal B_VAL As Variant, _
Optional ByVal nLOOPS As Long = 10001)

'trova un fattore fra n^1/3 < a < n^1/2

Dim k As Long
Dim M_VAL As Long

Dim R_VAL As Double

Dim Y_VAL As Double
Dim X_VAL As Double

Dim T_VAL As Variant
Dim T1_VAL As Variant
Dim T2_VAL As Variant
Dim TEMP_VAL As Variant

On Error GoTo ERROR_LABEL

'chek perfect square
R_VAL = Round(Sqr(NUMBER_VAL), 0)
If R_VAL * R_VAL = NUMBER_VAL Then
    A_VAL = R_VAL
    B_VAL = R_VAL
    Exit Function
End If
'
R_VAL = Int(NUMBER_VAL ^ (1 / 3))
k = 1
Do While k <= R_VAL And k < nLOOPS
    T1_VAL = 4 * k * CDec(NUMBER_VAL)
    X_VAL = Int(Sqr(T1_VAL)) + 1
    M_VAL = 0
    Do While M_VAL < Int(Sqr(R_VAL / k))
        'T_VAL = (X_VAL + M_VAL) * (X_VAL + M_VAL) - 4 * k * NUMBER_VAL  'overflow for NUMBER_VAL > 1E11
        T2_VAL = CDec(X_VAL + M_VAL) * (X_VAL + M_VAL)
        T_VAL = T2_VAL - T1_VAL
        If T_VAL > 0 Then
            Y_VAL = Sqr(T_VAL)
            If Y_VAL = Int(Y_VAL) Then
                X_VAL = X_VAL + M_VAL
                A_VAL = MCD2_FUNC(X_VAL + Y_VAL, NUMBER_VAL)
                B_VAL = MCD2_FUNC(X_VAL - Y_VAL, NUMBER_VAL)
                If B_VAL > A_VAL Then
                    TEMP_VAL = A_VAL
                    A_VAL = B_VAL
                    B_VAL = TEMP_VAL
                End If
                Exit Function
            End If
        End If
        M_VAL = M_VAL + 1
    Loop
    k = k + 1
Loop

If k = R_VAL + 1 Then 'prime
    A_VAL = NUMBER_VAL
    B_VAL = 1
ElseIf k = nLOOPS Then 'do not know
    A_VAL = "?"
    B_VAL = "?"
End If

FACTOR_FERMAT_LEHMAN_FUNC = 0

Exit Function
ERROR_LABEL:
FACTOR_FERMAT_LEHMAN_FUNC = Err.number
End Function
