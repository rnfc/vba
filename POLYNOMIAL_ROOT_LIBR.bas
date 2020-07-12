Attribute VB_Name = "POLYNOMIAL_ROOT_LIBR"


Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.

'************************************************************************************
'************************************************************************************
'FUNCTION      : POLYNOMIAL_MATRIX_REGRESSION_FUNC
'DESCRIPTION   : Polynomial Regression.
'LIBRARY       : POLYNOMIAL
'GROUP         : ROOT
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/22/2009
'************************************************************************************
'************************************************************************************

Function POLYNOMIAL_MATRIX_REGRESSION_FUNC(ByVal NDEG As Long, _
ByVal XDATA_RNG As Variant, _
ByVal YDATA_RNG As Variant, _
Optional ByVal INT_FLAG As Boolean = True)

Dim i As Long
Dim j As Long
Dim NSIZE As Long

Dim YDATA_VECTOR As Variant
Dim XDATA_VECTOR As Variant
Dim TEMP_MATRIX As Variant

On Error GoTo ERROR_LABEL

If NDEG = 0 Then: GoTo ERROR_LABEL
If NDEG > 3 Then: GoTo ERROR_LABEL

YDATA_VECTOR = YDATA_RNG
If UBound(YDATA_VECTOR, 1) = 1 Then
    YDATA_VECTOR = MATRIX_TRANSPOSE_FUNC(YDATA_VECTOR)
End If

XDATA_VECTOR = XDATA_RNG
If UBound(XDATA_VECTOR, 1) = 1 Then
    XDATA_VECTOR = MATRIX_TRANSPOSE_FUNC(XDATA_VECTOR)
End If

NSIZE = UBound(XDATA_VECTOR)

ReDim TEMP_MATRIX(1 To NSIZE, 1 To NDEG)
For i = 1 To NSIZE
    TEMP_MATRIX(i, 1) = XDATA_VECTOR(i, 1)
    For j = 2 To NDEG
        TEMP_MATRIX(i, j) = TEMP_MATRIX(i, 1) * TEMP_MATRIX(i, j - 1)
    Next j
    YDATA_VECTOR(i, 1) = YDATA_VECTOR(i, 1)
Next i

POLYNOMIAL_MATRIX_REGRESSION_FUNC = MATRIX_SVD_REGRESSION_FUNC(TEMP_MATRIX, YDATA_VECTOR, INT_FLAG)
'Matrix Regression --> Returns the coefficients vector
'[a0, a1,..am] of linear regression function
'y = a0 + a1*x + a2*x^2 +a3*x^3 +... am*x^m
'It use the SVD decomposition method

Exit Function
ERROR_LABEL:
POLYNOMIAL_MATRIX_REGRESSION_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : POLYNOMIAL_CHARACTERISTIC_COEFFICIENTS_FUNC

'DESCRIPTION   : 'This function returns the coefficients of the characteristic
'polynomial of a square matrix. If the matrix has dimension (NROWS x NROWS),
'then the polynomial has NROWS degree and the coefficients are n+1
'As know, the roots of the characteristic polynomial are the
'eigenvalues of the matrix and vice versa. This function uses the
'Newton-Girard formulas to find all the coefficients.

'Solving a polynomial of 3rd. degree (by any method)
'can be another way to find eigenvalues. Note: Computing
'eigenvalues trough the characteristic polynomial is in general
'less efficient than other decomposition methods (QR, Jacobi),
'but becomes a good choice for low-dimension
'matrices (typically < 6°) and for complex eigenvalues

'LIBRARY       : POLYNOMIAL
'GROUP         : ROOT
'ID            : 002
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/22/2009
'************************************************************************************
'************************************************************************************

Function POLYNOMIAL_CHARACTERISTIC_COEFFICIENTS_FUNC(ByRef DATA_RNG As Variant)

Dim i As Long
Dim k As Long

Dim NROWS As Long
Dim NCOLUMNS As Long

Dim TEMP_SUM As Double

Dim TEMP_VECTOR As Variant
Dim COEF_VECTOR As Variant

Dim ADATA_MATRIX As Variant
Dim BDATA_MATRIX As Variant

Dim TEMP_MATRIX As Variant

On Error GoTo ERROR_LABEL

ADATA_MATRIX = DATA_RNG
NROWS = UBound(ADATA_MATRIX, 1)
NCOLUMNS = UBound(ADATA_MATRIX, 2)
If NROWS <> NCOLUMNS Then GoTo ERROR_LABEL

NROWS = UBound(ADATA_MATRIX)

ReDim TEMP_VECTOR(1 To NROWS, 1 To 1)
ReDim COEF_VECTOR(0 To NROWS, 1 To 1)

COEF_VECTOR(NROWS, 1) = (-1) ^ NROWS
BDATA_MATRIX = ADATA_MATRIX

For k = 1 To NROWS
    For i = 1 To NROWS
        TEMP_VECTOR(k, 1) = TEMP_VECTOR(k, 1) + BDATA_MATRIX(i, i)
    Next i
    TEMP_SUM = 0
    For i = 0 To k - 1
        TEMP_SUM = TEMP_SUM + COEF_VECTOR(NROWS - i, 1) * TEMP_VECTOR(k - i, 1)
    Next i
    
    COEF_VECTOR(NROWS - k, 1) = -TEMP_SUM / k
    
    If k < NROWS Then
        TEMP_MATRIX = MMULT_FUNC(BDATA_MATRIX, ADATA_MATRIX, 70)
        BDATA_MATRIX = TEMP_MATRIX
    End If
Next k

POLYNOMIAL_CHARACTERISTIC_COEFFICIENTS_FUNC = COEF_VECTOR
Exit Function
ERROR_LABEL:
POLYNOMIAL_CHARACTERISTIC_COEFFICIENTS_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : POLYNOMIAL_SR_ROOTS_FUNC

'DESCRIPTION   : This function returns all roots of a given polynomial.
'Poly can be an array of (n+1) coefficients [a0, a1, a2...] or a string like
'"a0+a1x+a2x^2+..." This function uses the Siljak+Ruffini algorithm
'for finding the roots of an nth NDEG polynomial. Eigenvalues. If
'the given polynomial is the characteristic polynomial of a matrix
'this function return all eigenvalues of that matrix.

'Computing eigenvalues through the characteristic polynomial is in
'general less efficient than other decomposition methods (QR, Jacoby),
'but becomes a good choose for low-dimension matrices (typically < 6°)
'and/or for complex eigenvalues.

'LIBRARY       : POLYNOMIAL
'GROUP         : ROOT
'ID            : 003
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/22/2009
'************************************************************************************
'************************************************************************************


Function POLYNOMIAL_SR_ROOTS_FUNC(ByVal INPUT_RNG As Variant, _
Optional ByVal nLOOPS As Long = 200, _
Optional ByVal epsilon As Double = 2 * 10 ^ -15)

Dim h As Long
Dim i As Long
Dim j As Long
Dim k As Long
Dim l As Long

Dim NDEG As Long
Dim NROWS As Long
Dim COUNTER As Long
Dim NTRIALS As Long

Dim ANSIZE As Long
Dim BNSIZE As Long
Dim CNSIZE As Long
Dim DNSIZE As Long

Dim F_VAL As Double
Dim G_VAL As Double
Dim H_VAL As Double
Dim I_VAL As Double
Dim J_VAL As Double

Dim TEMP_MIN As Double
Dim TEMP_MAX As Double
Dim TEMP_FACTOR As Double

Dim ATEMP_VAL As Double
Dim BTEMP_VAL As Double
Dim CTEMP_VAL As Double
Dim DTEMP_VAL As Double

Dim UTEMP_VAL As Double
Dim TTEMP_VAL As Double

Dim XCPLX_VAL As Double
Dim YCPLX_VAL As Double

Dim XTEMP_DELTA As Double
Dim YTEMP_DELTA As Double

Dim DATA_VECTOR As Variant

Dim XTEMP_VECTOR As Variant
Dim YTEMP_VECTOR As Variant

Dim COEF_VECTOR As Variant

Dim REEL_VECTOR As Variant
Dim IMAG_VECTOR As Variant

Dim REEL_COEF_VECTOR As Variant
Dim IMAG_COEF_VECTOR As Variant

Dim REEL_ROOT_VECTOR As Variant
Dim IMAG_ROOT_VECTOR As Variant

Dim FIRST_ROOT_VECTOR As Variant
Dim SECOND_ROOT_VECTOR As Variant

Dim ATEMP_VECTOR As Variant
Dim BTEMP_VECTOR As Variant
Dim CTEMP_VECTOR As Variant
Dim DTEMP_VECTOR As Variant

Dim LAMBDA As Double

On Error GoTo ERROR_LABEL

If VarType(INPUT_RNG) = vbString Then
    DATA_VECTOR = PARSE_POLYNOMIAL_STRING_FUNC(INPUT_RNG, 0, "§")
    NROWS = (UBound(DATA_VECTOR, 1) - LBound(DATA_VECTOR, 1) + 1)
    ReDim COEF_VECTOR(1 To NROWS, 1 To 1)
    For i = 1 To NROWS
        COEF_VECTOR(i, 1) = DATA_VECTOR(i - 1, 1)
    Next i
ElseIf IsArray(INPUT_RNG) = True Then
    DATA_VECTOR = INPUT_RNG
    If LBound(DATA_VECTOR, 1) = 0 Then: DATA_VECTOR = MATRIX_CHANGE_BASE_ONE_FUNC(DATA_VECTOR)
    If UBound(DATA_VECTOR, 1) = 1 Then: DATA_VECTOR = MATRIX_TRANSPOSE_FUNC(DATA_VECTOR)
    COEF_VECTOR = DATA_VECTOR
    NROWS = UBound(COEF_VECTOR, 1)
Else
    GoTo ERROR_LABEL
End If

NDEG = NROWS - 1 'search for small integer roots

DNSIZE = UBound(COEF_VECTOR, 1) - 1
If DNSIZE < 2 Then GoTo 1982
ReDim ATEMP_VECTOR(1 To DNSIZE + 1, 1 To 1)
ReDim FIRST_ROOT_VECTOR(1 To DNSIZE, 1 To 1)

'load vector coefficients in descendent order
For i = 1 To DNSIZE + 1
    ATEMP_VECTOR(i, 1) = COEF_VECTOR(DNSIZE - i + 2, 1)
Next i

'find the lowest possible integer roots
    NTRIALS = 1000
    'ATEMP_VAL = 1
    TEMP_MIN = 0
    ANSIZE = 0
    BNSIZE = DNSIZE
    Do

        j = j + 1
        TEMP_MAX = 10 * Round(Abs(ATEMP_VECTOR(BNSIZE + 1, 1)) ^ _
                (1 / BNSIZE) + Abs(ATEMP_VECTOR(2, 1)) / BNSIZE, 0)
        If TEMP_MAX - TEMP_MIN > NTRIALS Then GoTo 1982
        
        
        ATEMP_VAL = TEMP_MIN + 1
        DTEMP_VAL = Abs(ATEMP_VECTOR(BNSIZE + 1, 1))
        k = 0
        Do 'Find Divisor
            J_VAL = DTEMP_VAL / ATEMP_VAL
            If J_VAL < 1 Or ATEMP_VAL > TEMP_MAX Or k > 1000 Then
                ATEMP_VAL = 0
                Exit Do
            End If
            COUNTER = J_VAL - Int(J_VAL)
            If COUNTER = 0 Then Exit Do
            ATEMP_VAL = ATEMP_VAL + 1
            k = k + 1
        Loop
        
        If ATEMP_VAL = 0 Then GoTo 1981
        'Ruffini's integer rootfinder starts
        Do
            'reduce the polynomial by the syntetic division
            ReDim BTEMP_VECTOR(1 To BNSIZE + 1, 1 To 1)
            BTEMP_VECTOR(1, 1) = ATEMP_VECTOR(1, 1)
            For i = 2 To BNSIZE + 1
                BTEMP_VECTOR(i, 1) = ATEMP_VECTOR(i, 1) + ATEMP_VAL * BTEMP_VECTOR(i - 1, 1)
            Next i
            'check if ATEMP_VAL is an integer FIRST_ROOT_VECTOR
            If BTEMP_VECTOR(BNSIZE + 1, 1) = 0 Then
                'OK. Take the FIRST_ROOT_VECTOR
               ANSIZE = ANSIZE + 1
                FIRST_ROOT_VECTOR(ANSIZE, 1) = ATEMP_VAL
                'load the reduced polynomial
                For i = 1 To BNSIZE + 1
                    ATEMP_VECTOR(i, 1) = BTEMP_VECTOR(i, 1)
                Next i
                BNSIZE = BNSIZE - 1
        Else
            If ATEMP_VAL > 0 Then
                ATEMP_VAL = -ATEMP_VAL 'try with the opposite value
            Else
                TEMP_MIN = TEMP_MIN + 1 'try with the next divisor
                Exit Do
            End If
        End If
    Loop

Loop Until BNSIZE = 0

1981:
ReDim COEF_VECTOR(1 To BNSIZE + 1, 1 To 1)
If BNSIZE > 0 Then
    For i = 0 To BNSIZE
        COEF_VECTOR(i + 1, 1) = ATEMP_VECTOR(BNSIZE + 1 - i, 1)
    Next i
End If

1982:



'search for non-integer and complex roots
If BNSIZE > 0 Then
    ReDim REEL_COEF_VECTOR(0 To BNSIZE, 1 To 1)
    ReDim IMAG_COEF_VECTOR(0 To BNSIZE, 1 To 1)
    ReDim REEL_ROOT_VECTOR(1 To BNSIZE, 1 To 1)
    ReDim IMAG_ROOT_VECTOR(1 To BNSIZE, 1 To 1)
    For i = 0 To BNSIZE
        REEL_COEF_VECTOR(i, 1) = COEF_VECTOR(i + 1, 1)
        IMAG_COEF_VECTOR(i, 1) = 0
    Next i
    

    CNSIZE = UBound(REEL_COEF_VECTOR, 1)
    ReDim XTEMP_VECTOR(0 To CNSIZE, 1 To 1)
    ReDim YTEMP_VECTOR(0 To CNSIZE, 1 To 1)

    J_VAL = CNSIZE
    If CNSIZE = 1 Then GoTo 1983 'Special check for linear polynomial
    Do While CNSIZE > 1
        XCPLX_VAL = 0.1
        YCPLX_VAL = 1
        XTEMP_VECTOR(0, 1) = 1
        YTEMP_VECTOR(0, 1) = 0
        XTEMP_VECTOR(1, 1) = 0.1
        YTEMP_VECTOR(1, 1) = 1
        h = 0
        LAMBDA = epsilon
        GoSub 1984 'Branch to computation of Siljak coefficients
        Do
            G_VAL = F_VAL
            l = 0
            I_VAL = 0
            H_VAL = 0
            h = h + 1
            For k = 1 To CNSIZE
                H_VAL = H_VAL + k * (REEL_COEF_VECTOR(k, 1) * _
                        XTEMP_VECTOR(k - 1, 1) - IMAG_COEF_VECTOR(k, 1) * _
                                YTEMP_VECTOR(k - 1, 1))
                I_VAL = I_VAL + k * (REEL_COEF_VECTOR(k, 1) * _
                        YTEMP_VECTOR(k - 1, 1) + IMAG_COEF_VECTOR(k, 1) * _
                            XTEMP_VECTOR(k - 1, 1))
            Next k
            TEMP_FACTOR = H_VAL * H_VAL + I_VAL * I_VAL
            XTEMP_DELTA = -(UTEMP_VAL * H_VAL + BTEMP_VAL * I_VAL) / TEMP_FACTOR
           YTEMP_DELTA = (UTEMP_VAL * I_VAL - BTEMP_VAL * H_VAL) / TEMP_FACTOR
            'succesive quartering
            Do
                l = l + 1
                If l > 20 Then GoTo 1985
                'Maximum of quartering has been set at 20
                'new root approxiamtions XCPLX_VAL, YCPLX_VAL
                'loaded into Silijak coefficients
                XTEMP_VECTOR(1, 1) = XCPLX_VAL + XTEMP_DELTA
                YTEMP_VECTOR(1, 1) = YCPLX_VAL + YTEMP_DELTA
                'recompute Siljak coefficients
                GoSub 1984
                'if the new error estimate greater than old,
                'quarter the size of Delta XCPLX_VAL and YCPLX_VAL
                If F_VAL <= 10 * G_VAL Then Exit Do
                XTEMP_DELTA = XTEMP_DELTA / 4
                YTEMP_DELTA = YTEMP_DELTA / 4
            Loop
            'increase the tolerance after tot iteration
            If h Mod 20 = 0 Then: If LAMBDA < 10 ^ -3 Then LAMBDA = LAMBDA * 10
            'check if increments are small enough to satisfy the stopping condition
            If (Abs(XTEMP_DELTA) < LAMBDA) And (Abs(YTEMP_DELTA) < LAMBDA) _
                Then Exit Do
            'check if maximum number of iteration has been exceeded
            If h > nLOOPS Then GoTo 1986
            'iterate again
            XCPLX_VAL = XTEMP_VECTOR(1, 1)
            YCPLX_VAL = YTEMP_VECTOR(1, 1)
        Loop

    'Root found. Store computed root in array element
        REEL_ROOT_VECTOR(CNSIZE, 1) = XTEMP_VECTOR(1, 1)
        IMAG_ROOT_VECTOR(CNSIZE, 1) = YTEMP_VECTOR(1, 1)
        'Initialize variables for Synthetic Division algorithm
        REEL_VECTOR = REEL_COEF_VECTOR(CNSIZE, 1)
        IMAG_VECTOR = IMAG_COEF_VECTOR(CNSIZE, 1)
        REEL_COEF_VECTOR(CNSIZE, 1) = 0
        IMAG_COEF_VECTOR(CNSIZE, 1) = 0
        XCPLX_VAL = XTEMP_VECTOR(1, 1)
        YCPLX_VAL = YTEMP_VECTOR(1, 1)
        'Synthetic Division to calculate new polynomial coefficients
        For k = CNSIZE - 1 To 0 Step -1
            CTEMP_VAL = REEL_COEF_VECTOR(k, 1)
            DTEMP_VAL = IMAG_COEF_VECTOR(k, 1)
            UTEMP_VAL = REEL_COEF_VECTOR(k + 1, 1)
            BTEMP_VAL = IMAG_COEF_VECTOR(k + 1, 1)
            REEL_COEF_VECTOR(k, 1) = REEL_VECTOR + XCPLX_VAL * _
                UTEMP_VAL - YCPLX_VAL * BTEMP_VAL
            IMAG_COEF_VECTOR(k, 1) = IMAG_VECTOR + XCPLX_VAL * _
                BTEMP_VAL + YCPLX_VAL * UTEMP_VAL
            REEL_VECTOR = CTEMP_VAL
            IMAG_VECTOR = DTEMP_VAL
        Next k
        CNSIZE = CNSIZE - 1
    Loop

1983:     'Since NDEG of resultant polynomial is one
'compute final root algebraically
    REEL_VECTOR = REEL_COEF_VECTOR(0, 1)
    UTEMP_VAL = REEL_COEF_VECTOR(1, 1)
    IMAG_VECTOR = IMAG_COEF_VECTOR(0, 1)
    BTEMP_VAL = IMAG_COEF_VECTOR(1, 1)
    TTEMP_VAL = UTEMP_VAL * UTEMP_VAL + BTEMP_VAL * BTEMP_VAL
    REEL_ROOT_VECTOR(1, 1) = -(REEL_VECTOR * UTEMP_VAL + _
        IMAG_VECTOR * BTEMP_VAL) / TTEMP_VAL
    IMAG_ROOT_VECTOR(1, 1) = (REEL_VECTOR * BTEMP_VAL - _
        UTEMP_VAL * IMAG_VECTOR) / TTEMP_VAL
    CNSIZE = J_VAL

1987:
    
    ReDim SECOND_ROOT_VECTOR(1 To BNSIZE, 1 To 2)
    For i = 1 To BNSIZE
        SECOND_ROOT_VECTOR(i, 1) = REEL_ROOT_VECTOR(i, 1)
        SECOND_ROOT_VECTOR(i, 2) = IMAG_ROOT_VECTOR(i, 1)
    Next i
End If



ReDim CTEMP_VECTOR(1 To NDEG, 1 To 2)

For i = 1 To ANSIZE
    CTEMP_VECTOR(i, 1) = FIRST_ROOT_VECTOR(i, 1)
    CTEMP_VECTOR(i, 2) = 0
Next i

For i = 1 To BNSIZE
    CTEMP_VECTOR(i + ANSIZE, 1) = SECOND_ROOT_VECTOR(i, 1)
    CTEMP_VECTOR(i + ANSIZE, 2) = SECOND_ROOT_VECTOR(i, 2)
    If Abs(CTEMP_VECTOR(i + ANSIZE, 1)) < epsilon Then CTEMP_VECTOR(i + ANSIZE, 1) = 0
    If Abs(CTEMP_VECTOR(i + ANSIZE, 2)) < epsilon Then CTEMP_VECTOR(i + ANSIZE, 2) = 0
Next i

NROWS = UBound(CTEMP_VECTOR, 1)

ReDim DTEMP_VECTOR(1 To NROWS, 1 To 3)
For i = 1 To NROWS
    DTEMP_VECTOR(i, 2) = CTEMP_VECTOR(i, 1)
    DTEMP_VECTOR(i, 3) = CTEMP_VECTOR(i, 2)
    DTEMP_VECTOR(i, 1) = Sqr(DTEMP_VECTOR(i, 2) ^ 2 + DTEMP_VECTOR(i, 3) ^ 2)
Next i

DTEMP_VECTOR = MATRIX_QUICK_SORT_FUNC(DTEMP_VECTOR, 1, 1)
'DTEMP_VECTOR = MATRIX_SWAP_SORT_FUNC(DTEMP_VECTOR, 1)

ReDim CTEMP_VECTOR(1 To NROWS, 1 To 2)
For i = 1 To NROWS
    CTEMP_VECTOR(i, 1) = DTEMP_VECTOR(i, 2)
    CTEMP_VECTOR(i, 2) = DTEMP_VECTOR(i, 3)
Next i

POLYNOMIAL_SR_ROOTS_FUNC = CTEMP_VECTOR

Exit Function


1984:
    'compute Siljak coefficients: complex binomial power (x+iy)^k
    TEMP_FACTOR = XTEMP_VECTOR(1, 1) * XTEMP_VECTOR(1, 1) + _
        YTEMP_VECTOR(1, 1) * YTEMP_VECTOR(1, 1)
    TTEMP_VAL = 2 * XTEMP_VECTOR(1, 1)
    For k = 0 To CNSIZE - 2
        XTEMP_VECTOR(k + 2, 1) = TTEMP_VAL * XTEMP_VECTOR(k + 1, 1) - _
                TEMP_FACTOR * XTEMP_VECTOR(k, 1)
        YTEMP_VECTOR(k + 2, 1) = TTEMP_VAL * YTEMP_VECTOR(k + 1, 1) - _
            TEMP_FACTOR * YTEMP_VECTOR(k, 1)
    Next k
    'compute complex polinomial value H_VAL(x+iy)
    UTEMP_VAL = 0
    BTEMP_VAL = 0
    For k = 0 To CNSIZE
        UTEMP_VAL = UTEMP_VAL + REEL_COEF_VECTOR(k, 1) * _
                XTEMP_VECTOR(k, 1) - IMAG_COEF_VECTOR(k, 1) * YTEMP_VECTOR(k, 1)
        BTEMP_VAL = BTEMP_VAL + REEL_COEF_VECTOR(k, 1) * _
                YTEMP_VECTOR(k, 1) + IMAG_COEF_VECTOR(k, 1) * XTEMP_VECTOR(k, 1)
    Next k
    F_VAL = UTEMP_VAL * UTEMP_VAL + BTEMP_VAL * BTEMP_VAL
Return

1985:
    CNSIZE = -2  'print error message: Maximum of quartering has been set at 20
    GoTo 1987

1986:     'Print error message: Exceded iteration limit
    CNSIZE = -1
    GoTo 1987

ERROR_LABEL:
POLYNOMIAL_SR_ROOTS_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : POLYNOMIAL_QR_ROOTS_FUNC

'DESCRIPTION   : 'Checking four real and four complex conjugate roots.

'This function returns all roots of a given polynomial. The argument "poly" can
'be an array of (n+1) coefficients [a0, a1, a2...] or a string like
'"a0+a1*x+a2*x^2+...". This function use the QR algorithm for finding the
'eigenvalues of the companion matrix of the given polynomial. This process
'is very fast, robust and stable but may not be converging under certain
'conditions. Usually it is suitable for solving polynomials up to 10th degree

'LIBRARY       : POLYNOMIAL
'GROUP         : ROOT
'ID            : 004
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/22/2009
'************************************************************************************
'************************************************************************************

Function POLYNOMIAL_QR_ROOTS_FUNC(ByVal INPUT_RNG As Variant)

Dim h As Long
Dim i As Long
Dim j As Long
Dim k As Long
Dim l As Long
Dim m As Long
Dim n As Long

Dim o As Long
Dim p As Long
Dim q As Long
Dim r As Long
Dim s As Long

Dim NROWS As Long
Dim nLOOPS As Long
Dim COUNTER As Long

Dim O_VAL As Double
Dim P_VAL As Double

Dim Q_VAL As Double
Dim R_VAL As Double
Dim S_VAL As Double

Dim T_VAL As Double
Dim U_VAL As Double
Dim V_VAL As Double

Dim W_VAL As Double

Dim FIRST_VAL As Double
Dim SECOND_VAL As Double

Dim COEF_VECTOR As Variant
Dim DATA_VECTOR As Variant
Dim TEMP_MATRIX As Variant

Dim REEL_VECTOR As Variant
Dim IMAG_VECTOR As Variant

Dim ATEMP_VECTOR As Variant
Dim BTEMP_VECTOR As Variant

Dim NOT_FLAG As Boolean

On Error GoTo ERROR_LABEL

If VarType(INPUT_RNG) = vbString Then
    DATA_VECTOR = PARSE_POLYNOMIAL_STRING_FUNC(INPUT_RNG, 0, "§")
    NROWS = (UBound(DATA_VECTOR, 1) - LBound(DATA_VECTOR, 1) + 1)
    ReDim COEF_VECTOR(1 To NROWS, 1 To 1)
    For i = 1 To NROWS
        COEF_VECTOR(i, 1) = DATA_VECTOR(i - 1, 1)
    Next i
ElseIf IsArray(INPUT_RNG) = True Then
    DATA_VECTOR = INPUT_RNG
    If LBound(DATA_VECTOR, 1) = 0 Then: DATA_VECTOR = MATRIX_CHANGE_BASE_ONE_FUNC(DATA_VECTOR)
    If UBound(DATA_VECTOR, 1) = 1 Then: DATA_VECTOR = MATRIX_TRANSPOSE_FUNC(DATA_VECTOR)
    COEF_VECTOR = DATA_VECTOR
    NROWS = UBound(COEF_VECTOR, 1)
Else
    GoTo ERROR_LABEL
End If


'build the companion matrix
TEMP_MATRIX = POLYNOMIAL_MONIC_COMPANION_FUNC(COEF_VECTOR)

NROWS = UBound(TEMP_MATRIX)

ReDim REEL_VECTOR(1 To NROWS)
ReDim IMAG_VECTOR(1 To NROWS)

'----------------find eigenvalues of companion matrix
'     THIS SUBROUTINE FINDS THE EIGENVALUES OF A REAL
'     UPPER HESSENBERG MATRIX BY THE QR METHOD.


NOT_FLAG = False
      s = 0
      k = 1

1983:
'
      m = NROWS
      S_VAL = 0#
      nLOOPS = 30 * NROWS
'     .......... SEARCH FOR NEXT EIGENVALUES ..........
1984:
      If (m < 1) Then GoTo 2000
      COUNTER = 0
      o = m - 1
      q = o - 1
'     .......... LOOK FOR SINGLE SMALL SUB-DIAGONAL ELEMENT
'                FOR L=EN STEP -1 UNTIL 1 DO -- ..........
1985:
     For h = 1 To m
            n = m + 1 - h
            If (n = 1) Then GoTo 1986
            R_VAL = Abs(TEMP_MATRIX(n - 1, n - 1)) + Abs(TEMP_MATRIX(n, n))
            If (R_VAL = 0) Then R_VAL = 1   'R_VAL = norm  '
            FIRST_VAL = R_VAL
            SECOND_VAL = FIRST_VAL + Abs(TEMP_MATRIX(n, n - 1))
            If (SECOND_VAL = FIRST_VAL) And _
                Abs(TEMP_MATRIX(n, n - 1)) < 1 Then GoTo 1986
     Next h
'     .......... FORM SHIFT ..........
1986:
      U_VAL = TEMP_MATRIX(m, m)
      If (n = m) Then GoTo 1995
      V_VAL = TEMP_MATRIX(o, o)
      T_VAL = TEMP_MATRIX(m, o) * TEMP_MATRIX(o, m)
      If (n = o) Then GoTo 1996
      If (nLOOPS = 0) Then GoTo 1999
      If ((COUNTER <> 10) And (COUNTER <> 20)) Then GoTo 1987
'     .......... FORM EXCEPTIONAL SHIFT ..........
      S_VAL = S_VAL + U_VAL
'
      For i = 1 To m
        TEMP_MATRIX(i, i) = TEMP_MATRIX(i, i) - U_VAL
      Next i
'
      R_VAL = Abs(TEMP_MATRIX(m, o)) + Abs(TEMP_MATRIX(o, q))
      U_VAL = 0.75 * R_VAL
      V_VAL = U_VAL
      T_VAL = -0.4375 * R_VAL * R_VAL
1987:
      COUNTER = COUNTER + 1
      nLOOPS = nLOOPS - 1
'     .......... LOOK FOR TWO CONSECUTIVE SMALL
'                SUB-DIAGONAL ELEMENTS.
'                FOR M=EN-2 STEP -1 UNTIL n DO -- ..........
      For r = n To q
         l = q + n - r
         W_VAL = TEMP_MATRIX(l, l)
         Q_VAL = U_VAL - W_VAL
         R_VAL = V_VAL - W_VAL
         O_VAL = (Q_VAL * R_VAL - T_VAL) / TEMP_MATRIX(l + 1, l) + TEMP_MATRIX(l, l + 1)
         P_VAL = TEMP_MATRIX(l + 1, l + 1) - W_VAL - Q_VAL - R_VAL
         Q_VAL = TEMP_MATRIX(l + 2, l + 1)
         R_VAL = Abs(O_VAL) + Abs(P_VAL) + Abs(Q_VAL)
         O_VAL = O_VAL / R_VAL
         P_VAL = P_VAL / R_VAL
         Q_VAL = Q_VAL / R_VAL
         If (l = n) Then GoTo 1988
         FIRST_VAL = Abs(O_VAL) * (Abs(TEMP_MATRIX(l - 1, l - 1)) + _
            Abs(W_VAL) + Abs(TEMP_MATRIX(l + 1, l + 1)))
         SECOND_VAL = FIRST_VAL + Abs(TEMP_MATRIX(l, l - 1)) * _
            (Abs(P_VAL) + Abs(Q_VAL))
         If (SECOND_VAL = FIRST_VAL) Then GoTo 1988
      Next r
'
1988:
      p = l + 2
'
      For i = p To m
         TEMP_MATRIX(i, i - 2) = 0#
         If (i <> p) Then TEMP_MATRIX(i, i - 3) = 0#
      Next i
'     .......... DOUBLE QR STEP INVOLVING ROWS n TO m AND
'                COLUMNS l TO m ..........
      For k = l To o
         NOT_FLAG = k <> o
         If (k = l) Then GoTo 1989
         O_VAL = TEMP_MATRIX(k, k - 1)
         P_VAL = TEMP_MATRIX(k + 1, k - 1)
         Q_VAL = 0#
         If (NOT_FLAG) Then Q_VAL = TEMP_MATRIX(k + 2, k - 1)
         U_VAL = Abs(O_VAL) + Abs(P_VAL) + Abs(Q_VAL)
         If (U_VAL = 0#) Then GoTo 1994
         O_VAL = O_VAL / U_VAL
         P_VAL = P_VAL / U_VAL
         Q_VAL = Q_VAL / U_VAL
1989:
         R_VAL = IIf(O_VAL >= 0, Abs(Sqr(O_VAL * O_VAL + P_VAL * P_VAL + Q_VAL * Q_VAL)), _
         -Abs(Sqr(O_VAL * O_VAL + P_VAL * P_VAL + Q_VAL * Q_VAL)))
         
         If (k = l) Then GoTo 1990
         TEMP_MATRIX(k, k - 1) = -R_VAL * U_VAL
         GoTo 1991
1990:
         If (n <> l) Then TEMP_MATRIX(k, k - 1) = -TEMP_MATRIX(k, k - 1)
1991:
         O_VAL = O_VAL + R_VAL
         U_VAL = O_VAL / R_VAL
         V_VAL = P_VAL / R_VAL
         W_VAL = Q_VAL / R_VAL
         P_VAL = P_VAL / O_VAL
         Q_VAL = Q_VAL / O_VAL
         If (NOT_FLAG) Then GoTo 1992
'     .......... ROW MODIFICATION ..........
         For j = k To NROWS
            O_VAL = TEMP_MATRIX(k, j) + P_VAL * TEMP_MATRIX(k + 1, j)
            TEMP_MATRIX(k, j) = TEMP_MATRIX(k, j) - O_VAL * U_VAL
            TEMP_MATRIX(k + 1, j) = TEMP_MATRIX(k + 1, j) - O_VAL * V_VAL
         Next j
'
         j = MINIMUM_FUNC(m, k + 3)
'     .......... COLUMN MODIFICATION ..........
         For i = 1 To j
            O_VAL = U_VAL * TEMP_MATRIX(i, k) + V_VAL * TEMP_MATRIX(i, k + 1)
            TEMP_MATRIX(i, k) = TEMP_MATRIX(i, k) - O_VAL
            TEMP_MATRIX(i, k + 1) = TEMP_MATRIX(i, k + 1) - O_VAL * P_VAL
         Next i
         GoTo 1993
1992:
'     .......... ROW MODIFICATION ..........
         For j = k To NROWS
            O_VAL = TEMP_MATRIX(k, j) + P_VAL * TEMP_MATRIX(k + 1, j) + Q_VAL * _
                TEMP_MATRIX(k + 2, j)
            TEMP_MATRIX(k, j) = TEMP_MATRIX(k, j) - O_VAL * U_VAL
            TEMP_MATRIX(k + 1, j) = TEMP_MATRIX(k + 1, j) - O_VAL * V_VAL
            TEMP_MATRIX(k + 2, j) = TEMP_MATRIX(k + 2, j) - O_VAL * W_VAL
         Next j
'
         j = MINIMUM_FUNC(m, k + 3)
'     .......... COLUMN MODIFICATION ..........
         For i = 1 To j
            O_VAL = U_VAL * TEMP_MATRIX(i, k) + V_VAL * TEMP_MATRIX(i, k + 1) + W_VAL * _
                TEMP_MATRIX(i, k + 2)
            TEMP_MATRIX(i, k) = TEMP_MATRIX(i, k) - O_VAL
            TEMP_MATRIX(i, k + 1) = TEMP_MATRIX(i, k + 1) - O_VAL * P_VAL
            TEMP_MATRIX(i, k + 2) = TEMP_MATRIX(i, k + 2) - O_VAL * Q_VAL
         Next i
1993:
'
      Next k
1994:
'
      GoTo 1985
'     .......... ONE ROOT FOUND ..........
1995:
      REEL_VECTOR(m) = U_VAL + S_VAL
      IMAG_VECTOR(m) = 0#
      m = o
      GoTo 1984
'     .......... TWO ROOTS FOUND ..........
1996:
      O_VAL = (V_VAL - U_VAL) / 2#
      P_VAL = O_VAL * O_VAL + T_VAL
      W_VAL = Sqr(Abs(P_VAL))
      U_VAL = U_VAL + S_VAL
      If (P_VAL < 0#) Then GoTo 1997
'     .......... REAL PAIR ..........
      W_VAL = O_VAL + IIf(O_VAL >= 0, Abs(W_VAL), -Abs(W_VAL))
      
      
      REEL_VECTOR(o) = U_VAL + W_VAL
      REEL_VECTOR(m) = REEL_VECTOR(o)
      If (W_VAL <> 0#) Then REEL_VECTOR(m) = U_VAL - T_VAL / W_VAL
      IMAG_VECTOR(o) = 0#
      IMAG_VECTOR(m) = 0#
      GoTo 1998
'     .......... COMPLEX PAIR ..........
1997:
      REEL_VECTOR(o) = U_VAL + O_VAL
      REEL_VECTOR(m) = U_VAL + O_VAL
      IMAG_VECTOR(o) = W_VAL
      IMAG_VECTOR(m) = -W_VAL
1998:
      m = q
      GoTo 1984
'     .......... SET ERROR -- ALL EIGENVALUES HAVE NOT
'                CONVERGED AFTER 30*NROWS ITERATIONS ..........
1999:
      s = m
2000:

'-------------------------------------------------------------------------------

ReDim ATEMP_VECTOR(1 To NROWS, 1 To 2)
For i = 1 To NROWS
    If i > s Then
        ATEMP_VECTOR(i, 1) = REEL_VECTOR(i)
        ATEMP_VECTOR(i, 2) = IMAG_VECTOR(i)
    Else
        ATEMP_VECTOR(i, 1) = "-"
        ATEMP_VECTOR(i, 2) = "-"
    End If
Next

NROWS = UBound(ATEMP_VECTOR, 1)

ReDim BTEMP_VECTOR(1 To NROWS, 1 To 3)
For i = 1 To NROWS
    BTEMP_VECTOR(i, 2) = ATEMP_VECTOR(i, 1)
    BTEMP_VECTOR(i, 3) = ATEMP_VECTOR(i, 2)
    BTEMP_VECTOR(i, 1) = Sqr(BTEMP_VECTOR(i, 2) ^ 2 + BTEMP_VECTOR(i, 3) ^ 2)
Next i

BTEMP_VECTOR = MATRIX_QUICK_SORT_FUNC(BTEMP_VECTOR, 1, 1)

ReDim ATEMP_VECTOR(1 To NROWS, 1 To 2)
For i = 1 To NROWS
    ATEMP_VECTOR(i, 1) = BTEMP_VECTOR(i, 2)
    ATEMP_VECTOR(i, 2) = BTEMP_VECTOR(i, 3)
Next i

POLYNOMIAL_QR_ROOTS_FUNC = ATEMP_VECTOR

Exit Function
ERROR_LABEL:
POLYNOMIAL_QR_ROOTS_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : POLYNOMIAL_MONIC_COMPANION_FUNC
'DESCRIPTION   : Returns the companion matrix of a monic polynomial
'LIBRARY       : POLYNOMIAL
'GROUP         : ROOT
'ID            : 005
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/22/2009
'************************************************************************************
'************************************************************************************

Function POLYNOMIAL_MONIC_COMPANION_FUNC(ByRef COEF_RNG As Variant)

Dim i As Long
Dim NROWS As Long
Dim NSIZE As Long

Dim DATA_VECTOR As Variant
Dim TEMP_MATRIX As Variant

On Error GoTo ERROR_LABEL

DATA_VECTOR = COEF_RNG
If UBound(DATA_VECTOR, 1) = 1 Then
    DATA_VECTOR = MATRIX_TRANSPOSE_FUNC(DATA_VECTOR)
End If

NROWS = UBound(DATA_VECTOR, 1)
NSIZE = NROWS - 1

ReDim TEMP_MATRIX(1 To NSIZE, 1 To NSIZE) 'build subdiagonal (lower)
For i = 1 To NSIZE - 1
    TEMP_MATRIX(i + 1, i) = 1
Next i
For i = 1 To NSIZE 'insert coefficients into last column
    TEMP_MATRIX(i, NSIZE) = -DATA_VECTOR(i, 1) / DATA_VECTOR(NROWS, 1)
Next i
POLYNOMIAL_MONIC_COMPANION_FUNC = TEMP_MATRIX

Exit Function
ERROR_LABEL:
POLYNOMIAL_MONIC_COMPANION_FUNC = Err.number
End Function

