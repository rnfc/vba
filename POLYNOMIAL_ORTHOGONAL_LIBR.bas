Attribute VB_Name = "POLYNOMIAL_ORTHOGONAL_LIBR"


Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.


'************************************************************************************
'************************************************************************************
'FUNCTION      : POLYNOMIAL_ORTHOGONAL_WEIGHTS_FUNC
'DESCRIPTION   : Orthogonal Polynomial Weights
'LIBRARY       : POLYNOMIAL
'GROUP         : ORTHOGONAL
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/22/2009
'************************************************************************************
'************************************************************************************

Function POLYNOMIAL_ORTHOGONAL_WEIGHTS_FUNC(ByVal NDEG As Long, _
Optional ByVal VERSION As Integer = 0, _
Optional ByVal A_VAL As Double, _
Optional ByVal B_VAL As Double)

Dim i As Long
Dim PI_VAL As Double
Dim TEMP_VAL As Variant

On Error GoTo ERROR_LABEL

PI_VAL = 3.14159265358979

Select Case VERSION ' PERFECT
'----------------------------------------------------------------------------------------
Case 0 'Legendre polynomial weight
    POLYNOMIAL_ORTHOGONAL_WEIGHTS_FUNC = 2 / (2 * NDEG + 1)
'----------------------------------------------------------------------------------------
Case 1 'Jacobi polynomial weight
    TEMP_VAL = (A_VAL + B_VAL + 1) * Log(2) - Log(2 * NDEG + A_VAL + B_VAL + 1)
    TEMP_VAL = TEMP_VAL + GAMMA_LN_FUNC(NDEG + A_VAL + 1) + GAMMA_LN_FUNC(NDEG + B_VAL + 1) - GAMMA_LN_FUNC(NDEG + 1) - GAMMA_LN_FUNC(NDEG + A_VAL + B_VAL + 1)
    POLYNOMIAL_ORTHOGONAL_WEIGHTS_FUNC = Exp(TEMP_VAL)
'----------------------------------------------------------------------------------------
Case 2 'Gegenbauer polynomial weight
    TEMP_VAL = Log(PI_VAL) + (1 - 2 * A_VAL) * Log(2) - Log(NDEG + 1)
    TEMP_VAL = TEMP_VAL + GAMMA_LN_FUNC(NDEG + 2 * A_VAL) - GAMMA_LN_FUNC(NDEG + 1) - 2 * GAMMA_LN_FUNC(A_VAL)
    POLYNOMIAL_ORTHOGONAL_WEIGHTS_FUNC = Exp(TEMP_VAL)
'----------------------------------------------------------------------------------------
Case 3 'Laguerre polynomial weight
    TEMP_VAL = 1
    For i = NDEG + 1 To NDEG + A_VAL: TEMP_VAL = TEMP_VAL * i: Next i
    POLYNOMIAL_ORTHOGONAL_WEIGHTS_FUNC = TEMP_VAL
'----------------------------------------------------------------------------------------
Case 4 'Hermite polynomial weight
    TEMP_VAL = 1
    For i = 1 To NDEG: TEMP_VAL = TEMP_VAL * 2 / i: Next i
    POLYNOMIAL_ORTHOGONAL_WEIGHTS_FUNC = Sqr(PI_VAL) * TEMP_VAL
'----------------------------------------------------------------------------------------
Case 5 'Chebyshev polynomial of the first kind weight
    If NDEG = 0 Then
        POLYNOMIAL_ORTHOGONAL_WEIGHTS_FUNC = 3.14159265358979
    Else
        POLYNOMIAL_ORTHOGONAL_WEIGHTS_FUNC = 1.5707963267949
    End If
'----------------------------------------------------------------------------------------
Case Else 'Chebyshev polynomial of the second kind weight
    POLYNOMIAL_ORTHOGONAL_WEIGHTS_FUNC = 1.5707963267949
'----------------------------------------------------------------------------------------
End Select

Exit Function
ERROR_LABEL:
POLYNOMIAL_ORTHOGONAL_WEIGHTS_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : POLYNOMIAL_TYPE_COEFFICIENTS_FUNC
'DESCRIPTION   : Compute the coefficients of the xType polynomial in
'standard 32-bit precision
'LIBRARY       : POLYNOMIAL
'GROUP         : ORTHOGONAL
'ID            : 002
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/22/2009
'************************************************************************************
'************************************************************************************

Function POLYNOMIAL_TYPE_COEFFICIENTS_FUNC(ByVal NDEG As Long, _
Optional ByVal POLYN_TYPE As Integer = 0, _
Optional ByRef LOWER_VAL As Double = 0, _
Optional ByRef UPPER_VAL As Double = 0)

Dim i As Long
Dim j As Long
Dim NSIZE As Long

Dim C_VAL As Double
Dim D_VAL As Double
Dim E_VAL As Double
Dim F_VAL As Double
Dim G_VAL As Double

Dim TEMP_MATRIX() As Variant
Dim TEMP_ARR() As Variant

Dim tolerance As Double

On Error GoTo ERROR_LABEL

tolerance = 0.000000000000001

Select Case POLYN_TYPE
'---------------------------------------------------------------------------------------
Case 0 'compute the coefficients of the Legendre polynomial (PERFECT)
'---------------------------------------------------------------------------------------
    
    NSIZE = NDEG
    
    ReDim TEMP_MATRIX(0 To NSIZE, 0 To 2), TEMP_ARR(0 To 2)
    
    TEMP_MATRIX(0, 0) = 1
    TEMP_MATRIX(1, 1) = 1
    TEMP_ARR(0) = 1
    TEMP_ARR(1) = 1
    
    j = 2
    Do Until j > NSIZE 'iterate
    
        LOWER_VAL = TEMP_ARR(0) * (2 * j - 1)
        UPPER_VAL = (j - 1) * TEMP_ARR(1)
        TEMP_MATRIX(0, 2) = -UPPER_VAL * TEMP_MATRIX(0, 0)
        
        For i = 1 To j
            TEMP_MATRIX(i, 2) = LOWER_VAL * TEMP_MATRIX(i - 1, 1) - UPPER_VAL * TEMP_MATRIX(i, 0)
        Next i
    'compute the GCD
        C_VAL = j * TEMP_ARR(1) * TEMP_ARR(0)
        TEMP_ARR(2) = C_VAL
        For i = 1 To j
            C_VAL = PAIR_GCD_FUNC(C_VAL, TEMP_MATRIX(i, 2))
        Next i
    'reduce terms
        TEMP_ARR(2) = TEMP_ARR(2) / C_VAL
        For i = 0 To j
            TEMP_MATRIX(i, 2) = TEMP_MATRIX(i, 2) / C_VAL
        Next i
    'shift
        For i = 0 To j
            TEMP_MATRIX(i, 0) = TEMP_MATRIX(i, 1)
            TEMP_MATRIX(i, 1) = TEMP_MATRIX(i, 2)
            TEMP_MATRIX(i, 2) = 0
        Next i
        TEMP_ARR(0) = TEMP_ARR(1)
        TEMP_ARR(1) = TEMP_ARR(2)
        j = j + 1
    Loop

    ReDim COEF_VECTOR(0 To NSIZE, 1 To 1)
    For i = 0 To NSIZE
        COEF_VECTOR(i, 1) = TEMP_MATRIX(i, 1) / TEMP_ARR(1) 'Scale Coefficients
    Next i
'---------------------------------------------------------------------------------------
Case 1 'compute the coefficients of the jacobi polynomial (PERFECT)
'---------------------------------------------------------------------------------------
    NSIZE = NDEG
    
    ReDim TEMP_MATRIX(0 To NSIZE, 0 To 2)
    TEMP_MATRIX(0, 0) = 1
    TEMP_MATRIX(1, 0) = (LOWER_VAL - UPPER_VAL) / 2
    TEMP_MATRIX(1, 1) = (LOWER_VAL + UPPER_VAL + 2) / 2
    
    j = 2
    Do Until j > NSIZE 'iterate
        D_VAL = 2 * j * (j + LOWER_VAL + UPPER_VAL) * (2 * j - 2 + LOWER_VAL + UPPER_VAL)
        E_VAL = (2 * j - 1 + LOWER_VAL + UPPER_VAL) * (LOWER_VAL ^ 2 - UPPER_VAL ^ 2) / D_VAL
        F_VAL = POLYNOMIAL_POCHHAMMER_FUNC(2 * j - 2 + LOWER_VAL + UPPER_VAL, 3) / D_VAL
        G_VAL = -2 * (j - 1 + LOWER_VAL) * (j - 1 + UPPER_VAL) * (2 * j + LOWER_VAL + UPPER_VAL) / D_VAL
        For i = 0 To j
            TEMP_MATRIX(i, 2) = E_VAL * TEMP_MATRIX(i, 1) + G_VAL * TEMP_MATRIX(i, 0)
        Next i
        For i = 1 To j
            TEMP_MATRIX(i, 2) = TEMP_MATRIX(i, 2) + F_VAL * TEMP_MATRIX(i - 1, 1)
        Next i
        'shift
        For i = 0 To j
            TEMP_MATRIX(i, 0) = TEMP_MATRIX(i, 1)
            TEMP_MATRIX(i, 1) = TEMP_MATRIX(i, 2)
            TEMP_MATRIX(i, 2) = 0
        Next i
        j = j + 1
    Loop

    ReDim COEF_VECTOR(0 To NSIZE, 1 To 1)
    For i = 0 To NSIZE
        COEF_VECTOR(i, 1) = TEMP_MATRIX(i, 1)
    Next i

'---------------------------------------------------------------------------------------
Case 2 'compute the coefficients of the GEGENBAUER polynomial (PERFECT)
'---------------------------------------------------------------------------------------
    'In this case LOWER_VAL = Lambda
    If LOWER_VAL = 0 Then: LOWER_VAL = tolerance
    
    NSIZE = NDEG
    ReDim TEMP_MATRIX(0 To NSIZE, 0 To 2)

    TEMP_MATRIX(0, 0) = 1
    TEMP_MATRIX(1, 1) = 2 * LOWER_VAL
    j = 2
    Do Until j > NSIZE
    'iterate
        UPPER_VAL = 2 * (j + LOWER_VAL - 1) / j
        C_VAL = -(j + 2 * LOWER_VAL - 2) / j
        TEMP_MATRIX(0, 2) = C_VAL * TEMP_MATRIX(0, 0)
        For i = 1 To j
            TEMP_MATRIX(i, 2) = UPPER_VAL * TEMP_MATRIX(i - 1, 1) + C_VAL * TEMP_MATRIX(i, 0)
        Next i
    'shift
        For i = 0 To j
            TEMP_MATRIX(i, 0) = TEMP_MATRIX(i, 1)
            TEMP_MATRIX(i, 1) = TEMP_MATRIX(i, 2)
            TEMP_MATRIX(i, 2) = 0
        Next i
        j = j + 1
    Loop
    ReDim COEF_VECTOR(0 To NSIZE, 1 To 1)
    For i = 0 To NSIZE
        COEF_VECTOR(i, 1) = TEMP_MATRIX(i, 1)
    Next i
'---------------------------------------------------------------------------------------
Case 3 'compute the coefficients of the Laguerre polynomial (PERFECT)
'---------------------------------------------------------------------------------------
    'LOWER_VAL --> Epsilon
    NSIZE = NDEG
    ReDim TEMP_MATRIX(0 To NSIZE, 0 To 2)
    ReDim TEMP_ARR(0 To 2)
    TEMP_MATRIX(0, 0) = 1
    TEMP_MATRIX(0, 1) = 1 + LOWER_VAL
    TEMP_MATRIX(1, 1) = -1
    TEMP_ARR(0) = 1
    TEMP_ARR(1) = 1
    j = 2
    Do Until j > NSIZE 'iterate
        UPPER_VAL = TEMP_ARR(0) * (2 * j - 1 + LOWER_VAL)
        C_VAL = (j - 1 + LOWER_VAL) * TEMP_ARR(1)
        For i = 0 To j
            TEMP_MATRIX(i, 2) = UPPER_VAL * TEMP_MATRIX(i, 1) - C_VAL * TEMP_MATRIX(i, 0)
        Next i
        For i = 1 To j
            TEMP_MATRIX(i, 2) = TEMP_MATRIX(i, 2) - TEMP_ARR(0) * TEMP_MATRIX(i - 1, 1)
        Next i
        'compute the GCD
        D_VAL = j * TEMP_ARR(1) * TEMP_ARR(0)
        TEMP_ARR(2) = D_VAL
        For i = 1 To j
            D_VAL = PAIR_GCD_FUNC(D_VAL, TEMP_MATRIX(i, 2))
        Next i
        'reduce terms
        TEMP_ARR(2) = TEMP_ARR(2) / D_VAL
        For i = 0 To j
            TEMP_MATRIX(i, 2) = TEMP_MATRIX(i, 2) / D_VAL
        Next i
        'shift
        For i = 0 To j
            TEMP_MATRIX(i, 0) = TEMP_MATRIX(i, 1)
            TEMP_MATRIX(i, 1) = TEMP_MATRIX(i, 2)
            TEMP_MATRIX(i, 2) = 0
        Next i
        TEMP_ARR(0) = TEMP_ARR(1)
        TEMP_ARR(1) = TEMP_ARR(2)
        j = j + 1
    Loop
    
    ReDim COEF_VECTOR(0 To NSIZE, 1 To 1)
    For i = 0 To NSIZE
        COEF_VECTOR(i, 1) = TEMP_MATRIX(i, 1) / TEMP_ARR(1)
    Next i

'---------------------------------------------------------------------------------------
Case 4 'compute the coefficients of the Hermite polynomial
'---------------------------------------------------------------------------------------
    NSIZE = NDEG
    ReDim TEMP_MATRIX(0 To NSIZE, 0 To 2)
    TEMP_MATRIX(0, 0) = 1
    TEMP_MATRIX(1, 1) = 2
    j = 2
    Do Until j > NSIZE 'iterate
        LOWER_VAL = 2
        UPPER_VAL = -2
        TEMP_MATRIX(0, 2) = (j - 1) * UPPER_VAL * TEMP_MATRIX(0, 0)
        For i = 1 To j
            TEMP_MATRIX(i, 2) = LOWER_VAL * TEMP_MATRIX(i - 1, 1) + UPPER_VAL * (j - 1) * TEMP_MATRIX(i, 0)
        Next i
        'shift
        For i = 0 To j
            TEMP_MATRIX(i, 0) = TEMP_MATRIX(i, 1)
            TEMP_MATRIX(i, 1) = TEMP_MATRIX(i, 2)
            TEMP_MATRIX(i, 2) = 0
        Next i
        j = j + 1
    Loop
    
    ReDim COEF_VECTOR(0 To NSIZE, 1 To 1)
    For i = 0 To NSIZE
        COEF_VECTOR(i, 1) = TEMP_MATRIX(i, 1)
    Next i
'---------------------------------------------------------------------------------------
Case 5 'compute the coefficients of the Chebychev polynomial 1st kind (PERFECT)
'---------------------------------------------------------------------------------------
    NSIZE = NDEG
    ReDim TEMP_MATRIX(0 To NSIZE, 0 To 2)
    TEMP_MATRIX(0, 0) = 1
    TEMP_MATRIX(1, 1) = 1
    
    j = 2
    Do Until j > NSIZE
        'iterate
        LOWER_VAL = 2
        UPPER_VAL = -1
        TEMP_MATRIX(0, 2) = UPPER_VAL * TEMP_MATRIX(0, 0)
        For i = 1 To j
            TEMP_MATRIX(i, 2) = LOWER_VAL * TEMP_MATRIX(i - 1, 1) + UPPER_VAL * TEMP_MATRIX(i, 0)
        Next i
        'shift
        For i = 0 To j
            TEMP_MATRIX(i, 0) = TEMP_MATRIX(i, 1)
            TEMP_MATRIX(i, 1) = TEMP_MATRIX(i, 2)
            TEMP_MATRIX(i, 2) = 0
        Next i
        j = j + 1
    Loop
    
    ReDim COEF_VECTOR(0 To NSIZE, 1 To 1)
    For i = 0 To NSIZE
        COEF_VECTOR(i, 1) = TEMP_MATRIX(i, 1)
    Next i

'---------------------------------------------------------------------------------------
Case Else 'compute the coefficients of the Chebychev polynomial 2nd kind (PERFECT)
'---------------------------------------------------------------------------------------
    NSIZE = NDEG
    ReDim TEMP_MATRIX(0 To NSIZE, 0 To 2)
    TEMP_MATRIX(0, 0) = 1
    TEMP_MATRIX(1, 1) = 2
    
    j = 2
    Do Until j > NSIZE
        'iterate
        LOWER_VAL = 2
        UPPER_VAL = -1
        TEMP_MATRIX(0, 2) = UPPER_VAL * TEMP_MATRIX(0, 0)
        For i = 1 To j
            TEMP_MATRIX(i, 2) = LOWER_VAL * TEMP_MATRIX(i - 1, 1) + UPPER_VAL * TEMP_MATRIX(i, 0)
        Next i
        'shift
        For i = 0 To j
            TEMP_MATRIX(i, 0) = TEMP_MATRIX(i, 1)
            TEMP_MATRIX(i, 1) = TEMP_MATRIX(i, 2)
            TEMP_MATRIX(i, 2) = 0
        Next i
        j = j + 1
    Loop
    
    ReDim COEF_VECTOR(0 To NSIZE, 1 To 1)
    For i = 0 To NSIZE
        COEF_VECTOR(i, 1) = TEMP_MATRIX(i, 1)
    Next i

End Select

POLYNOMIAL_TYPE_COEFFICIENTS_FUNC = COEF_VECTOR

Exit Function
ERROR_LABEL:
POLYNOMIAL_TYPE_COEFFICIENTS_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : POLYNOMIAL_ORTHOGONAL_INTERPOLATION_FUNC
'DESCRIPTION   : Orthogonal Polynomial Interpolation Function
'LIBRARY       : POLYNOMIAL
'GROUP         : ORTHOGONAL
'ID            : 003
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/22/2009
'************************************************************************************
'************************************************************************************

Function POLYNOMIAL_ORTHOGONAL_INTERPOLATION_FUNC(ByRef COEF_RNG As Variant, _
Optional ByVal NBINS As Long = 101, _
Optional ByVal MIN_VAL As Double = 0.5, _
Optional ByVal MAX_VAL As Double = 0.5)

Dim i As Long
Dim k As Long

Dim SROW As Long
Dim NROWS As Long

Dim A_VAL As Double
Dim B_VAL As Double
Dim Y_VAL As Double

Dim TEMP_VECTOR As Variant
Dim COEF_VECTOR As Variant

On Error GoTo ERROR_LABEL
        
COEF_VECTOR = COEF_RNG
If UBound(COEF_VECTOR, 1) = 1 Then
    COEF_VECTOR = MATRIX_TRANSPOSE_FUNC(COEF_VECTOR)
End If
SROW = LBound(COEF_VECTOR, 1)
NROWS = UBound(COEF_VECTOR, 1)
            
B_VAL = 2 / (NBINS - 1)
        
ReDim TEMP_VECTOR(1 To NBINS, 1 To 2)
        
For i = 1 To NBINS
    A_VAL = (i - 1) * B_VAL - 1
    TEMP_VECTOR(i, 1) = MAX_VAL * A_VAL + MIN_VAL 'x
    For k = NROWS To SROW Step -1
        Y_VAL = Y_VAL * A_VAL + COEF_VECTOR(k, 1)
    Next k
    TEMP_VECTOR(i, 2) = Y_VAL 'y
Next i
            
POLYNOMIAL_ORTHOGONAL_INTERPOLATION_FUNC = TEMP_VECTOR

Exit Function
ERROR_LABEL:
POLYNOMIAL_ORTHOGONAL_INTERPOLATION_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : POLYNOMIAL_ORTHOGONAL_ADK_ROOTS_FUNC
'DESCRIPTION   : Robust Rootfinder for orthogonal polynomials using the
'ADK algorithm (Aberth_Durand_Kerner)
'LIBRARY       : POLYNOMIAL
'GROUP         : ORTHOGONAL
'ID            : 004
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/22/2009
'************************************************************************************
'************************************************************************************

Function POLYNOMIAL_ORTHOGONAL_ADK_ROOTS_FUNC(ByVal VERSION As Integer, _
ByVal NDEG As Long, _
ByVal MIN_VAL As Double, _
ByVal MAX_VAL As Double, _
Optional ByVal REF_VALUE As Double = 0, _
Optional ByVal LOWER_VAL As Double = 0, _
Optional ByVal UPPER_VAL As Double = 0, _
Optional ByVal OUTPUT As Integer = 0)

'---------------------------------------------------------------------------------
'REFERENCE: Abramowitz M et al.; "Handbook of Mathematical Functions...",Dover
'Press et al.; "Numerical recipies in fotran77", Cambridge U Press
'---------------------------------------------------------------------------------

Dim i As Long
Dim j As Long
Dim k As Long
Dim l As Long

Dim TEMP_SUM As Double
Dim TEMP_VAL As Double
Dim SAVED_VAL As Double
Dim DELTA_VAL As Double

Dim F1_VAL As Double
Dim F2_VAL As Double

Dim FIRST_VECTOR As Variant
Dim SECOND_VECTOR As Variant

Dim ROOTS_VECTOR As Variant

Dim epsilon As Double
Dim tolerance As Double
Dim LAMBDA As Double

Dim nLOOPS As Long

On Error GoTo ERROR_LABEL

tolerance = 10 ^ -15
epsilon = 0.1
nLOOPS = NDEG * 50

'---------------------------------------------------------------------------------
'If VERSION = 3 Then
'    MAX_VAL = 1.1 * (NDEG * REF_VALUE / _
'    50 + 6 / 5 * REF_VALUE + 18 / 5 * NDEG - 5)
'ElseIf VERSION = 4 Then
'    MIN_VAL = -0.75 * (NDEG - 1)
'    MAX_VAL = -MIN_VAL
'End If

'---------------------------------------------------------------------------------

LAMBDA = (MAX_VAL - MIN_VAL) / (NDEG - 1)

TEMP_VAL = MAX_VAL
SAVED_VAL = TEMP_VAL
l = 0

ReDim ZEROS_VECTOR(1 To NDEG, 1 To 1)

For i = 1 To NDEG
    
    Do
        l = l + 1
        
        Select Case VERSION
'----------------------------------------------------------------------------------------
            Case 0 ' Rutina para calcular el polinomio ortonormal de Legendre de
' orden NDEG y su derivada en TEMP_VAL. Los polinomios de Legendre son un caso
' especial de los de Jacobi con LOWER_VAL = UPPER_VAL = 0. F1_VAL valor
' del polinomio en TEMP_VAL; F2_VAL valor de la derivada del polinomio
' en TEMP_VAL (PERFECT)
'----------------------------------------------------------------------------------------

    ReDim FIRST_VECTOR(0 To 2)
    ReDim SECOND_VECTOR(0 To 2)

    If NDEG = 0 Then
        F1_VAL = 1
        F2_VAL = 0
    ElseIf NDEG = 1 Then
        F1_VAL = TEMP_VAL
        F2_VAL = 1
    Else
        FIRST_VECTOR(0) = 1
        FIRST_VECTOR(1) = TEMP_VAL
        If Abs(TEMP_VAL - 1) < epsilon Or Abs(TEMP_VAL + 1) < epsilon Then
            SECOND_VECTOR(0) = 0
            SECOND_VECTOR(1) = 1
            For k = 1 To NDEG - 1
                FIRST_VECTOR(2) = ((2 * k + 1) * TEMP_VAL * _
                    FIRST_VECTOR(1) - k * FIRST_VECTOR(0)) / (k + 1)
                SECOND_VECTOR(2) = ((2 * k + 1) * (FIRST_VECTOR(1) + _
                    TEMP_VAL * SECOND_VECTOR(1)) - k * _
                        SECOND_VECTOR(0)) / (k + 1)
                FIRST_VECTOR(0) = FIRST_VECTOR(1)
                FIRST_VECTOR(1) = FIRST_VECTOR(2)
                SECOND_VECTOR(0) = SECOND_VECTOR(1)
                SECOND_VECTOR(1) = SECOND_VECTOR(2)
            Next k
            F1_VAL = FIRST_VECTOR(2)
            F2_VAL = SECOND_VECTOR(2)
        Else
            For k = 1 To NDEG - 1
                FIRST_VECTOR(2) = ((2 * k + 1) * TEMP_VAL * _
                    FIRST_VECTOR(1) - k * FIRST_VECTOR(0)) / (k + 1)
                FIRST_VECTOR(0) = FIRST_VECTOR(1)
                FIRST_VECTOR(1) = FIRST_VECTOR(2)
            Next k
            F1_VAL = FIRST_VECTOR(2)
            F2_VAL = NDEG * (TEMP_VAL * FIRST_VECTOR(2) - _
                FIRST_VECTOR(0)) / (TEMP_VAL ^ 2 - 1)
        End If
    End If


GoTo 1985
'----------------------------------------------------------------------------------------
            Case 1 ' Rutina para calcular el polinomio ortonormal de Jacobi de
' orden NDEG y su derivada en TEMP_VAL. F1_VAL valor del polinomio en
' TEMP_VAL; F2_VAL valor de la derivada del polinomio en TEMP_VAL (PERFECT)
'----------------------------------------------------------------------------------------
    If NDEG = 0 Then
        F1_VAL = 1
        F2_VAL = 0
    ElseIf NDEG = 1 Then
        F1_VAL = 0.5 * (2 * (LOWER_VAL + 1) + _
            (LOWER_VAL + UPPER_VAL + 2) * (TEMP_VAL - 1))
        F2_VAL = 0.5 * (LOWER_VAL + UPPER_VAL + 2)
    Else
        F1_VAL = POLYNOMIAL_ORTHOGONAL_JACOBI_ROOTS_FUNC(LOWER_VAL, UPPER_VAL, _
            NDEG, TEMP_VAL)
        F2_VAL = 0.5 * (NDEG + LOWER_VAL + _
            UPPER_VAL + 1) * POLYNOMIAL_ORTHOGONAL_JACOBI_ROOTS_FUNC(LOWER_VAL + 1, _
            UPPER_VAL + 1, NDEG - 1, TEMP_VAL)
    End If
GoTo 1985
'----------------------------------------------------------------------------------------
            Case 2 'Rutina para calcular el polinomio de ortonormal Gegenbauer de
' orden NDEG y su derivada en TEMP_VAL. Los polinomios de Gegenbauer
' son un caso especial de los de Jacobi con LOWER_VAL = UPPER_VAL = l-1/2
' Cuando l=1/2 aparecen los polinomios de Legendre F1_VAL valor del polinomio
' en TEMP_VAL; F2_VAL valor de la derivada del polinomio en TEMP_VAL
'----------------------------------------------------------------------------------------

    If (REF_VALUE = 0) Then: GoTo 1983
    If (REF_VALUE = 1) Then: GoTo 1984

    ReDim FIRST_VECTOR(0 To 2)
    ReDim SECOND_VECTOR(0 To 2)
    
    If NDEG = 0 Then
        F1_VAL = 1
        F2_VAL = 0
    ElseIf NDEG = 1 Then
        F1_VAL = 2 * REF_VALUE * TEMP_VAL
        F2_VAL = 2 * REF_VALUE
    Else
        FIRST_VECTOR(0) = 1
        FIRST_VECTOR(1) = 2 * REF_VALUE * TEMP_VAL
        If Abs(TEMP_VAL - 1) < epsilon Or Abs(TEMP_VAL + 1) < epsilon Then
            SECOND_VECTOR(0) = 0
            SECOND_VECTOR(1) = 2 * REF_VALUE
             For k = 1 To NDEG - 1
                FIRST_VECTOR(2) = (2 * (k + REF_VALUE) * TEMP_VAL * _
                        FIRST_VECTOR(1) - (k + 2 * REF_VALUE - 1) * _
                                FIRST_VECTOR(0)) / (k + 1)
                SECOND_VECTOR(2) = (2 * (k + REF_VALUE) * _
                        FIRST_VECTOR(1) + 2 * (k + REF_VALUE) * _
                                TEMP_VAL * SECOND_VECTOR(1) - _
                                        (k + 2 * REF_VALUE - 1) _
                * SECOND_VECTOR(0)) / (k + 1)
                FIRST_VECTOR(0) = FIRST_VECTOR(1)
                FIRST_VECTOR(1) = FIRST_VECTOR(2)
                SECOND_VECTOR(0) = SECOND_VECTOR(1)
                SECOND_VECTOR(1) = SECOND_VECTOR(2)
            Next k
            F1_VAL = FIRST_VECTOR(2)
            F2_VAL = SECOND_VECTOR(2)
        Else
            For k = 1 To NDEG - 1
                FIRST_VECTOR(2) = (2 * (k + REF_VALUE) * TEMP_VAL * _
                        FIRST_VECTOR(1) - (k + 2 * REF_VALUE - 1) * _
                                FIRST_VECTOR(0)) / (k + 1)
                FIRST_VECTOR(0) = FIRST_VECTOR(1)
                FIRST_VECTOR(1) = FIRST_VECTOR(2)
            Next k
            F1_VAL = FIRST_VECTOR(2)
            F2_VAL = (NDEG * TEMP_VAL * FIRST_VECTOR(2) - _
                    NDEG * FIRST_VECTOR(0)) / (TEMP_VAL ^ 2 - 1)
        End If
    End If

GoTo 1985

'----------------------------------------------------------------------------------------
            Case 3 ' Rutina para calcular el polinomio ortonormal de Laguerre
' asociado REF_VALUE de orden NDEG y su derivada en TEMP_VAL
' F1_VAL valor del polinomio en TEMP_VAL; F2_VAL valor de la
' derivada del polinomio en TEMP_VAL
'----------------------------------------------------------------------------------------
    ReDim FIRST_VECTOR(0 To 2)
    ReDim SECOND_VECTOR(0 To 2)

    If NDEG = 0 Then
        F1_VAL = 1
        F2_VAL = 0
    ElseIf NDEG = 1 Then
        F1_VAL = REF_VALUE + 1 - TEMP_VAL
        F2_VAL = -1
    Else
        FIRST_VECTOR(0) = 1
        FIRST_VECTOR(1) = REF_VALUE + 1 - TEMP_VAL
        If Abs(TEMP_VAL) < epsilon Then
            SECOND_VECTOR(0) = 0
            SECOND_VECTOR(1) = -1
            For k = 1 To NDEG - 1
                FIRST_VECTOR(2) = ((-TEMP_VAL + 2 * k + 1 + REF_VALUE) * _
                        FIRST_VECTOR(1) - (k + REF_VALUE) * FIRST_VECTOR(0)) / (k + 1)
                SECOND_VECTOR(2) = (-FIRST_VECTOR(1) + (-TEMP_VAL + _
                        2 * k + 1 + REF_VALUE) * SECOND_VECTOR(1) - (k + REF_VALUE) * _
                                SECOND_VECTOR(0)) / (k + 1)
                FIRST_VECTOR(0) = FIRST_VECTOR(1)
                FIRST_VECTOR(1) = FIRST_VECTOR(2)
                SECOND_VECTOR(0) = SECOND_VECTOR(1)
                SECOND_VECTOR(1) = SECOND_VECTOR(2)
            Next k
            F1_VAL = FIRST_VECTOR(2)
            F2_VAL = SECOND_VECTOR(2)
        Else
            For k = 1 To NDEG - 1
                FIRST_VECTOR(2) = ((-TEMP_VAL + 2 * k + 1 + REF_VALUE) * _
                        FIRST_VECTOR(1) - (k + REF_VALUE) * FIRST_VECTOR(0)) / (k + 1)
                FIRST_VECTOR(0) = FIRST_VECTOR(1)
                FIRST_VECTOR(1) = FIRST_VECTOR(2)
            Next k
            F1_VAL = FIRST_VECTOR(2)
            F2_VAL = (NDEG * FIRST_VECTOR(2) - (NDEG + REF_VALUE) * _
                    FIRST_VECTOR(0)) / TEMP_VAL
        End If
    End If

GoTo 1985
'----------------------------------------------------------------------------------------
            Case 4 'Hermite Polynomials
'----------------------------------------------------------------------------------------
        
    ReDim FIRST_VECTOR(0 To 2)
    ReDim SECOND_VECTOR(0 To 2)

    If NDEG = 0 Then
        F1_VAL = 1
        F2_VAL = 0
    ElseIf NDEG = 1 Then
        F1_VAL = 2 * TEMP_VAL
        F2_VAL = 2
    Else
        FIRST_VECTOR(0) = 1
        FIRST_VECTOR(1) = 2 * TEMP_VAL
        For k = 1 To NDEG - 1
            FIRST_VECTOR(2) = 2 * TEMP_VAL * FIRST_VECTOR(1) - _
                2 * k * FIRST_VECTOR(0)
            FIRST_VECTOR(0) = FIRST_VECTOR(1)
            FIRST_VECTOR(1) = FIRST_VECTOR(2)
        Next k
        F1_VAL = FIRST_VECTOR(2)
        F2_VAL = 2 * NDEG * FIRST_VECTOR(0)
    End If
        
GoTo 1985
'----------------------------------------------------------------------------------------
            Case 5 'Chebychev Polynomials of 1st kind
'----------------------------------------------------------------------------------------
1983:
    ' Rutina para calcular los ceros del polinomio ortogonal
    ' de Chebychev de primera especie de orden NDEG

    'Const pi  As Double = 3.14159265358979
    'For i = 1 To NDEG
    '    z(i) = Cos(pi * (i - 0.5) / NDEG)
    'Next i

    ReDim FIRST_VECTOR(0 To 2)
    ReDim SECOND_VECTOR(0 To 2)
    
    If NDEG = 0 Then
        F1_VAL = 1
        F2_VAL = 0
    ElseIf NDEG = 1 Then
        F1_VAL = TEMP_VAL
        F2_VAL = 1
    Else
        FIRST_VECTOR(0) = 1
        FIRST_VECTOR(1) = TEMP_VAL
        SECOND_VECTOR(0) = 0
        SECOND_VECTOR(1) = 1
        For k = 1 To NDEG - 1
            FIRST_VECTOR(2) = 2 * TEMP_VAL * _
                FIRST_VECTOR(1) - FIRST_VECTOR(0)
            SECOND_VECTOR(2) = 2 * FIRST_VECTOR(1) + 2 * _
                TEMP_VAL * SECOND_VECTOR(1) - SECOND_VECTOR(0)
            FIRST_VECTOR(0) = FIRST_VECTOR(1)
            FIRST_VECTOR(1) = FIRST_VECTOR(2)
            SECOND_VECTOR(0) = SECOND_VECTOR(1)
            SECOND_VECTOR(1) = SECOND_VECTOR(2)
        Next k
        F1_VAL = FIRST_VECTOR(2)
        F2_VAL = SECOND_VECTOR(2)
    End If

'------------------------------------------------------------------------------------------
        If (VERSION = 2) And (LOWER_VAL = 0) Then
            F1_VAL = F1_VAL * 2 / NDEG
            F2_VAL = F2_VAL * 2 / NDEG
        End If
'------------------------------------------------------------------------------------------

GoTo 1985

'----------------------------------------------------------------------------------------
            Case Else 'Chebychev Polynomials of 2nd kind
'----------------------------------------------------------------------------------------
1984:
    
    ReDim FIRST_VECTOR(0 To 2)
    ReDim SECOND_VECTOR(0 To 2)
    
    ' Rutina para calcular los ceros del polinomio ortogonal
    ' de Chebychev de primera especie de orden NDEG
    '-----------------ZEROS----------------
    'Const pi  As Double = 3.14159265358979
    'For i = 1 To NDEG
    '    z(i) = Cos(pi * i / (NDEG + 1))
    'Next i
    '--------------------------------------
    
    If NDEG = 0 Then
        F1_VAL = 1
        F2_VAL = 0
    ElseIf NDEG = 1 Then
        F1_VAL = 2 * TEMP_VAL
        F2_VAL = 1
    Else
        FIRST_VECTOR(0) = 1
        FIRST_VECTOR(1) = 2 * TEMP_VAL
        SECOND_VECTOR(0) = 0
        SECOND_VECTOR(1) = 2
        For k = 1 To NDEG - 1
            FIRST_VECTOR(2) = 2 * TEMP_VAL * _
                FIRST_VECTOR(1) - FIRST_VECTOR(0)
            SECOND_VECTOR(2) = 2 * FIRST_VECTOR(1) + 2 * _
                TEMP_VAL * SECOND_VECTOR(1) - SECOND_VECTOR(0)
            FIRST_VECTOR(0) = FIRST_VECTOR(1)
            FIRST_VECTOR(1) = FIRST_VECTOR(2)
            SECOND_VECTOR(0) = SECOND_VECTOR(1)
            SECOND_VECTOR(1) = SECOND_VECTOR(2)
        Next k
        F1_VAL = FIRST_VECTOR(2)
        F2_VAL = SECOND_VECTOR(2)
    End If
  
'-------------------------------------------------------------------------------------
        End Select
'-------------------------------------------------------------------------------------
        
1985:
        TEMP_SUM = 0
        For j = 1 To i - 1
            TEMP_SUM = TEMP_SUM + 1 / (TEMP_VAL - ZEROS_VECTOR(j, 1))
        Next j
        SAVED_VAL = TEMP_VAL
        If F2_VAL <> 0 Then
            DELTA_VAL = F1_VAL / F2_VAL
        Else
            DELTA_VAL = tolerance * 10 ^ 6
        End If
        TEMP_VAL = TEMP_VAL - DELTA_VAL / (1 - DELTA_VAL * TEMP_SUM)
    
    Loop Until Abs(TEMP_VAL - SAVED_VAL) < tolerance Or l > nLOOPS
    
    ZEROS_VECTOR(i, 1) = TEMP_VAL
    TEMP_VAL = TEMP_VAL + epsilon * LAMBDA

Next i

Select Case OUTPUT
Case 0
    POLYNOMIAL_ORTHOGONAL_ADK_ROOTS_FUNC = ZEROS_VECTOR
Case Else
    ReDim ROOTS_VECTOR(1 To 2, 1 To 1)
    ROOTS_VECTOR(1, 1) = F1_VAL
    ROOTS_VECTOR(2, 1) = F2_VAL
    
    POLYNOMIAL_ORTHOGONAL_ADK_ROOTS_FUNC = ROOTS_VECTOR
End Select

Exit Function
ERROR_LABEL:
POLYNOMIAL_ORTHOGONAL_ADK_ROOTS_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : POLYNOMIAL_EVALUATE_FUNC
'DESCRIPTION   : Evaluate Polynomials (Legendre , Jacobi, Gegenbauer,
'Laguerre, Hermite, Chebychev 1st & 2nd kind)
'LIBRARY       : POLYNOMIAL
'GROUP         : ORTHOGONAL
'ID            : 005
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/22/2009
'************************************************************************************
'************************************************************************************

Function POLYNOMIAL_EVALUATE_FUNC(ByVal VERSION As Integer, _
ByVal NDEG As Long, _
ByVal X_VAL As Double, _
Optional ByVal REF_VALUE As Double = 0, _
Optional ByVal LOWER_VAL As Double = 0, _
Optional ByVal UPPER_VAL As Double = 0)

'---------------------------------------------------------------------------------
'REFERENCE: Abramowitz M et al.; "Handbook of Mathematical Functions...",Dover
'Press et al.; "Numerical recipies in fotran77", Cambridge U Press
'---------------------------------------------------------------------------------

'Dim i As Long
Dim k As Long

Dim F1_VAL As Double
Dim F2_VAL As Double

Dim F1_ARR As Variant
Dim F2_ARR As Variant

Dim TEMP_VECTOR As Variant

Dim epsilon As Double
Dim tolerance As Double

On Error GoTo ERROR_LABEL

tolerance = 10 ^ -15
epsilon = 0.1


Select Case VERSION
'----------------------------------------------------------------------------------------
            Case 0 ' Rutina para calcular el polinomio ortonormal de Legendre de
' orden NDEG y su derivada en X_VAL. Los polinomios de Legendre son un caso
' especial de los de Jacobi con LOWER_VAL = UPPER_VAL = 0. F1_VAL valor
' del polinomio en X_VAL; F2_VAL valor de la derivada del polinomio
' en X_VAL (PERFECT)
'----------------------------------------------------------------------------------------

    ReDim F1_ARR(0 To 2)
    ReDim F2_ARR(0 To 2)

    If NDEG = 0 Then
        F1_VAL = 1
        F2_VAL = 0
    ElseIf NDEG = 1 Then
        F1_VAL = X_VAL
        F2_VAL = 1
    Else
        F1_ARR(0) = 1
        F1_ARR(1) = X_VAL
        If Abs(X_VAL - 1) < epsilon Or Abs(X_VAL + 1) < epsilon Then
            F2_ARR(0) = 0
            F2_ARR(1) = 1
            For k = 1 To NDEG - 1
                F1_ARR(2) = ((2 * k + 1) * X_VAL * _
                    F1_ARR(1) - k * F1_ARR(0)) / (k + 1)
                F2_ARR(2) = ((2 * k + 1) * (F1_ARR(1) + _
                    X_VAL * F2_ARR(1)) - k * _
                        F2_ARR(0)) / (k + 1)
                F1_ARR(0) = F1_ARR(1)
                F1_ARR(1) = F1_ARR(2)
                F2_ARR(0) = F2_ARR(1)
                F2_ARR(1) = F2_ARR(2)
            Next k
            F1_VAL = F1_ARR(2)
            F2_VAL = F2_ARR(2)
        Else
            For k = 1 To NDEG - 1
                F1_ARR(2) = ((2 * k + 1) * X_VAL * _
                    F1_ARR(1) - k * F1_ARR(0)) / (k + 1)
                F1_ARR(0) = F1_ARR(1)
                F1_ARR(1) = F1_ARR(2)
            Next k
            F1_VAL = F1_ARR(2)
            F2_VAL = NDEG * (X_VAL * F1_ARR(2) - _
                F1_ARR(0)) / (X_VAL ^ 2 - 1)
        End If
    End If

GoTo 1985
'----------------------------------------------------------------------------------------
            Case 1 ' Rutina para calcular el polinomio ortonormal de Jacobi de
' orden NDEG y su derivada en X_VAL. F1_VAL valor del polinomio en
' X_VAL; F2_VAL valor de la derivada del polinomio en X_VAL (PERFECT)
'----------------------------------------------------------------------------------------
    If NDEG = 0 Then
        F1_VAL = 1
        F2_VAL = 0
    ElseIf NDEG = 1 Then
        F1_VAL = 0.5 * (2 * (LOWER_VAL + 1) + _
            (LOWER_VAL + UPPER_VAL + 2) * (X_VAL - 1))
        F2_VAL = 0.5 * (LOWER_VAL + UPPER_VAL + 2)
    Else
        F1_VAL = POLYNOMIAL_ORTHOGONAL_JACOBI_ROOTS_FUNC(LOWER_VAL, UPPER_VAL, _
            NDEG, X_VAL)
        F2_VAL = 0.5 * (NDEG + LOWER_VAL + _
            UPPER_VAL + 1) * POLYNOMIAL_ORTHOGONAL_JACOBI_ROOTS_FUNC(LOWER_VAL + 1, _
            UPPER_VAL + 1, NDEG - 1, X_VAL)
    End If
GoTo 1985
'----------------------------------------------------------------------------------------
            Case 2 'Rutina para calcular el polinomio de ortonormal Gegenbauer de
' orden NDEG y su derivada en X_VAL. Los polinomios de Gegenbauer
' son un caso especial de los de Jacobi con LOWER_VAL = UPPER_VAL = l-1/2
' Cuando l=1/2 aparecen los polinomios de Legendre F1_VAL valor del polinomio
' en X_VAL; F2_VAL valor de la derivada del polinomio en X_VAL
'----------------------------------------------------------------------------------------

    If (REF_VALUE = 0) Then: GoTo 1983
    If (REF_VALUE = 1) Then: GoTo 1984

    ReDim F1_ARR(0 To 2)
    ReDim F2_ARR(0 To 2)
    
    If NDEG = 0 Then
        F1_VAL = 1
        F2_VAL = 0
    ElseIf NDEG = 1 Then
        F1_VAL = 2 * REF_VALUE * X_VAL
        F2_VAL = 2 * REF_VALUE
    Else
        F1_ARR(0) = 1
        F1_ARR(1) = 2 * REF_VALUE * X_VAL
        If Abs(X_VAL - 1) < epsilon Or Abs(X_VAL + 1) < epsilon Then
            F2_ARR(0) = 0
            F2_ARR(1) = 2 * REF_VALUE
             For k = 1 To NDEG - 1
                F1_ARR(2) = (2 * (k + REF_VALUE) * X_VAL * _
                        F1_ARR(1) - (k + 2 * REF_VALUE - 1) * _
                                F1_ARR(0)) / (k + 1)
                F2_ARR(2) = (2 * (k + REF_VALUE) * _
                        F1_ARR(1) + 2 * (k + REF_VALUE) * _
                                X_VAL * F2_ARR(1) - _
                                        (k + 2 * REF_VALUE - 1) _
                * F2_ARR(0)) / (k + 1)
                F1_ARR(0) = F1_ARR(1)
                F1_ARR(1) = F1_ARR(2)
                F2_ARR(0) = F2_ARR(1)
                F2_ARR(1) = F2_ARR(2)
            Next k
            F1_VAL = F1_ARR(2)
            F2_VAL = F2_ARR(2)
        Else
            For k = 1 To NDEG - 1
                F1_ARR(2) = (2 * (k + REF_VALUE) * X_VAL * _
                        F1_ARR(1) - (k + 2 * REF_VALUE - 1) * _
                                F1_ARR(0)) / (k + 1)
                F1_ARR(0) = F1_ARR(1)
                F1_ARR(1) = F1_ARR(2)
            Next k
            F1_VAL = F1_ARR(2)
            F2_VAL = (NDEG * X_VAL * F1_ARR(2) - _
                    NDEG * F1_ARR(0)) / (X_VAL ^ 2 - 1)
        End If
    End If

GoTo 1985
'----------------------------------------------------------------------------------------
            Case 3 ' Rutina para calcular el polinomio ortonormal de Laguerre
' asociado REF_VALUE de orden NDEG y su derivada en X_VAL
' F1_VAL valor del polinomio en X_VAL; F2_VAL valor de la
' derivada del polinomio en X_VAL
'----------------------------------------------------------------------------------------
    ReDim F1_ARR(0 To 2)
    ReDim F2_ARR(0 To 2)

    If NDEG = 0 Then
        F1_VAL = 1
        F2_VAL = 0
    ElseIf NDEG = 1 Then
        F1_VAL = REF_VALUE + 1 - X_VAL
        F2_VAL = -1
    Else
        F1_ARR(0) = 1
        F1_ARR(1) = REF_VALUE + 1 - X_VAL
        If Abs(X_VAL) < epsilon Then
            F2_ARR(0) = 0
            F2_ARR(1) = -1
            For k = 1 To NDEG - 1
                F1_ARR(2) = ((-X_VAL + 2 * k + 1 + REF_VALUE) * _
                        F1_ARR(1) - (k + REF_VALUE) * F1_ARR(0)) / (k + 1)
                F2_ARR(2) = (-F1_ARR(1) + (-X_VAL + _
                        2 * k + 1 + REF_VALUE) * F2_ARR(1) - (k + REF_VALUE) * _
                                F2_ARR(0)) / (k + 1)
                F1_ARR(0) = F1_ARR(1)
                F1_ARR(1) = F1_ARR(2)
                F2_ARR(0) = F2_ARR(1)
                F2_ARR(1) = F2_ARR(2)
            Next k
            F1_VAL = F1_ARR(2)
            F2_VAL = F2_ARR(2)
        Else
            For k = 1 To NDEG - 1
                F1_ARR(2) = ((-X_VAL + 2 * k + 1 + REF_VALUE) * _
                        F1_ARR(1) - (k + REF_VALUE) * F1_ARR(0)) / (k + 1)
                F1_ARR(0) = F1_ARR(1)
                F1_ARR(1) = F1_ARR(2)
            Next k
            F1_VAL = F1_ARR(2)
            F2_VAL = (NDEG * F1_ARR(2) - (NDEG + REF_VALUE) * _
                    F1_ARR(0)) / X_VAL
        End If
    End If

GoTo 1985

'----------------------------------------------------------------------------------------
            Case 4 'Hermite Polynomials
'----------------------------------------------------------------------------------------
        
    ReDim F1_ARR(0 To 2)
    ReDim F2_ARR(0 To 2)

    If NDEG = 0 Then
        F1_VAL = 1
        F2_VAL = 0
    ElseIf NDEG = 1 Then
        F1_VAL = 2 * X_VAL
        F2_VAL = 2
    Else
        F1_ARR(0) = 1
        F1_ARR(1) = 2 * X_VAL
        For k = 1 To NDEG - 1
            F1_ARR(2) = 2 * X_VAL * F1_ARR(1) - _
                2 * k * F1_ARR(0)
            F1_ARR(0) = F1_ARR(1)
            F1_ARR(1) = F1_ARR(2)
        Next k
        F1_VAL = F1_ARR(2)
        F2_VAL = 2 * NDEG * F1_ARR(0)
    End If
        
GoTo 1985

'----------------------------------------------------------------------------------------
            Case 5 'Chebychev Polynomials of 1st kind
'----------------------------------------------------------------------------------------
1983:
    ' Rutina para calcular los ceros del polinomio ortogonal
    ' de Chebychev de primera especie de orden NDEG

    'Const pi  As Double = 3.14159265358979
    'For i = 1 To NDEG
    '    z(i) = Cos(pi * (i - 0.5) / NDEG)
    'Next i

    ReDim F1_ARR(0 To 2)
    ReDim F2_ARR(0 To 2)
    
    If NDEG = 0 Then
        F1_VAL = 1
        F2_VAL = 0
    ElseIf NDEG = 1 Then
        F1_VAL = X_VAL
        F2_VAL = 1
    Else
        F1_ARR(0) = 1
        F1_ARR(1) = X_VAL
        F2_ARR(0) = 0
        F2_ARR(1) = 1
        For k = 1 To NDEG - 1
            F1_ARR(2) = 2 * X_VAL * _
                F1_ARR(1) - F1_ARR(0)
            F2_ARR(2) = 2 * F1_ARR(1) + 2 * _
                X_VAL * F2_ARR(1) - F2_ARR(0)
            F1_ARR(0) = F1_ARR(1)
            F1_ARR(1) = F1_ARR(2)
            F2_ARR(0) = F2_ARR(1)
            F2_ARR(1) = F2_ARR(2)
        Next k
        F1_VAL = F1_ARR(2)
        F2_VAL = F2_ARR(2)
    End If

'------------------------------------------------------------------------------------------
        If (VERSION = 2) And (LOWER_VAL = 0) Then
            F1_VAL = F1_VAL * 2 / NDEG
            F2_VAL = F2_VAL * 2 / NDEG
        End If
'------------------------------------------------------------------------------------------

GoTo 1985

'----------------------------------------------------------------------------------------
            Case Else 'Chebychev Polynomials of 2nd kind
'----------------------------------------------------------------------------------------
1984:
    
    ReDim F1_ARR(0 To 2)
    ReDim F2_ARR(0 To 2)
    
    ' Rutina para calcular los ceros del polinomio ortogonal
    ' de Chebychev de primera especie de orden NDEG
    '-----------------ZEROS----------------
    'Const pi  As Double = 3.14159265358979
    'For i = 1 To NDEG
    '    z(i) = Cos(pi * i / (NDEG + 1))
    'Next i
    '--------------------------------------
    
    If NDEG = 0 Then
        F1_VAL = 1
        F2_VAL = 0
    ElseIf NDEG = 1 Then
        F1_VAL = 2 * X_VAL
        F2_VAL = 1
    Else
        F1_ARR(0) = 1
        F1_ARR(1) = 2 * X_VAL
        F2_ARR(0) = 0
        F2_ARR(1) = 2
        For k = 1 To NDEG - 1
            F1_ARR(2) = 2 * X_VAL * _
                F1_ARR(1) - F1_ARR(0)
            F2_ARR(2) = 2 * F1_ARR(1) + 2 * _
                X_VAL * F2_ARR(1) - F2_ARR(0)
            F1_ARR(0) = F1_ARR(1)
            F1_ARR(1) = F1_ARR(2)
            F2_ARR(0) = F2_ARR(1)
            F2_ARR(1) = F2_ARR(2)
        Next k
        F1_VAL = F1_ARR(2)
        F2_VAL = F2_ARR(2)
    End If
  
'-------------------------------------------------------------------------------------
        End Select
'-------------------------------------------------------------------------------------
1985:

    ReDim TEMP_VECTOR(1 To 2, 1 To 1)
    
    TEMP_VECTOR(1, 1) = F1_VAL
    TEMP_VECTOR(2, 1) = F2_VAL
            
    POLYNOMIAL_EVALUATE_FUNC = TEMP_VECTOR
    
Exit Function
ERROR_LABEL:
POLYNOMIAL_EVALUATE_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : POLYNOMIAL_ORTHOGONAL_JACOBI_ROOTS_FUNC
'DESCRIPTION   : Robust Rootfinder for Jacobi orthogonal polynomials
'LIBRARY       : POLYNOMIAL
'GROUP         : ORTHOGONAL
'ID            : 006
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/22/2009
'************************************************************************************
'************************************************************************************

Private Function POLYNOMIAL_ORTHOGONAL_JACOBI_ROOTS_FUNC( _
ByVal LOWER_VAL As Double, _
ByVal UPPER_VAL As Double, _
ByVal NDEG As Long, _
ByVal X_VAL As Double)
    
Dim j As Long

Dim C_VAL As Double
Dim D_VAL As Double
Dim E_VAL As Double
Dim F_VAL As Double

Dim TEMP_ARR As Variant

On Error GoTo ERROR_LABEL

ReDim TEMP_ARR(0 To 2)

Select Case NDEG
Case 0
    POLYNOMIAL_ORTHOGONAL_JACOBI_ROOTS_FUNC = 1
Case 1
    POLYNOMIAL_ORTHOGONAL_JACOBI_ROOTS_FUNC = 0.5 * (LOWER_VAL - UPPER_VAL + (2 + LOWER_VAL + UPPER_VAL) * X_VAL)
Case Else
    TEMP_ARR(0) = 1
    TEMP_ARR(1) = 0.5 * (LOWER_VAL - UPPER_VAL + _
        (2 + LOWER_VAL + UPPER_VAL) * X_VAL)
    For j = 1 To NDEG - 1
        C_VAL = 2 * (j + 1) * (j + LOWER_VAL + UPPER_VAL + 1) * _
            (2 * j + LOWER_VAL + UPPER_VAL)
        D_VAL = (2 * j + LOWER_VAL + UPPER_VAL + 1) * _
            (LOWER_VAL ^ 2 - UPPER_VAL ^ 2)
        E_VAL = (2 * j + LOWER_VAL + UPPER_VAL) * _
            (2 * j + LOWER_VAL + UPPER_VAL + 1) * _
            (2 * j + LOWER_VAL + UPPER_VAL + 2)
        F_VAL = 2 * (j + LOWER_VAL) * (j + UPPER_VAL) * _
            (2 * j + LOWER_VAL + UPPER_VAL + 2)

        TEMP_ARR(2) = ((D_VAL + E_VAL * X_VAL) * _
            TEMP_ARR(1) - F_VAL * TEMP_ARR(0)) / C_VAL
        TEMP_ARR(0) = TEMP_ARR(1)
        TEMP_ARR(1) = TEMP_ARR(2)
    Next j
    POLYNOMIAL_ORTHOGONAL_JACOBI_ROOTS_FUNC = TEMP_ARR(2)
End Select

Exit Function
ERROR_LABEL:
POLYNOMIAL_ORTHOGONAL_JACOBI_ROOTS_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : POLYNOMIAL_POCHHAMMER_FUNC
'DESCRIPTION   : Rising factorial or "upper factorial"; An algorithm for summing
'orthogonal polynomial series and their derivatives
'LIBRARY       : POLYNOMIAL
'GROUP         : ORTHOGONAL
'ID            : 007
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/22/2009
'************************************************************************************
'************************************************************************************

Private Function POLYNOMIAL_POCHHAMMER_FUNC(ByVal X_VAL As Double, _
ByVal NDEG As Long)

Dim i As Long
Dim MULT_VAL As Double

On Error GoTo ERROR_LABEL

MULT_VAL = X_VAL
For i = 1 To NDEG - 1: MULT_VAL = MULT_VAL * (X_VAL + i): Next i
POLYNOMIAL_POCHHAMMER_FUNC = MULT_VAL

Exit Function
ERROR_LABEL:
POLYNOMIAL_POCHHAMMER_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : POLYNOMIAL_LEGENDRE_FUNC
'DESCRIPTION   : Legendre's polynomials Function
'LIBRARY       : POLYNOMIAL
'GROUP         : ORTHOGONAL
'ID            : 008
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/22/2009
'************************************************************************************
'************************************************************************************

Function POLYNOMIAL_LEGENDRE_FUNC(ByVal X_VAL As Double, _
ByVal NROWS As Long, _
Optional ByVal SROW As Long = 1)

Dim i As Long
Dim A_VAL As Double
Dim B_VAL As Double
Dim C_VAL As Double

On Error GoTo ERROR_LABEL

A_VAL = 0
B_VAL = 1
C_VAL = B_VAL
For i = SROW To NROWS
    C_VAL = (2 * i - 1) / i * X_VAL * B_VAL - (i - 1) / i * A_VAL
    A_VAL = B_VAL
    B_VAL = C_VAL
Next i

POLYNOMIAL_LEGENDRE_FUNC = C_VAL

Exit Function
ERROR_LABEL:
POLYNOMIAL_LEGENDRE_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : POLYNOMIAL_HERMITE_FUNC
'DESCRIPTION   : Hermite's polynomials Function
'LIBRARY       : POLYNOMIAL
'GROUP         : ORTHOGONAL
'ID            : 009
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/22/2009
'************************************************************************************
'************************************************************************************

Function POLYNOMIAL_HERMITE_FUNC(ByVal X_VAL As Double, _
ByVal NROWS As Long, _
Optional ByVal SROW As Long = 1)


Dim i As Long
Dim A_VAL As Double
Dim B_VAL As Double
Dim C_VAL As Double

On Error GoTo ERROR_LABEL

A_VAL = 0
B_VAL = 1
C_VAL = B_VAL
For i = SROW To NROWS
    C_VAL = 2 * X_VAL * B_VAL - 2 * (i - 1) * A_VAL
    A_VAL = B_VAL
    B_VAL = C_VAL
Next i
POLYNOMIAL_HERMITE_FUNC = C_VAL

Exit Function
ERROR_LABEL:
POLYNOMIAL_HERMITE_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : POLYNOMIAL_LAGUERRE_FUNC
'DESCRIPTION   : Laguerre's polynomials Function
'LIBRARY       : POLYNOMIAL
'GROUP         : ORTHOGONAL
'ID            : 010
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/22/2009
'************************************************************************************
'************************************************************************************

Function POLYNOMIAL_LAGUERRE_FUNC(ByVal X_VAL As Double, _
ByVal NROWS As Long, _
Optional ByVal SROW As Long = 1)


Dim i As Long
Dim A_VAL As Double
Dim B_VAL As Double
Dim C_VAL As Double

On Error GoTo ERROR_LABEL

A_VAL = 0
B_VAL = 1
C_VAL = B_VAL
For i = SROW To NROWS
    C_VAL = (2 * i - 1 - X_VAL) * B_VAL - (i - 1) ^ 2 * A_VAL
    A_VAL = B_VAL
    B_VAL = C_VAL
Next i
POLYNOMIAL_LAGUERRE_FUNC = C_VAL

Exit Function
ERROR_LABEL:
POLYNOMIAL_LAGUERRE_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : POLYNOMIAL_CHEBYCEV_FUNC
'DESCRIPTION   : Chebycev's polynomials Function
'LIBRARY       : POLYNOMIAL
'GROUP         : ORTHOGONAL
'ID            : 011
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/22/2009
'************************************************************************************
'************************************************************************************

Function POLYNOMIAL_CHEBYCEV_FUNC(ByVal X_VAL As Double, _
ByVal NROWS As Long, _
Optional ByVal SROW As Long = 1)

Dim i As Long
Dim A_VAL As Double
Dim B_VAL As Double
Dim C_VAL As Double

On Error GoTo ERROR_LABEL

If NROWS = 0 Then
    POLYNOMIAL_CHEBYCEV_FUNC = 1
    Exit Function
End If

If NROWS = 1 Then
    POLYNOMIAL_CHEBYCEV_FUNC = X_VAL
    Exit Function
End If

A_VAL = 1
B_VAL = X_VAL

For i = SROW To NROWS - 1
    C_VAL = 2 * X_VAL * B_VAL - A_VAL
    A_VAL = B_VAL
    B_VAL = C_VAL
Next i
POLYNOMIAL_CHEBYCEV_FUNC = C_VAL

Exit Function
ERROR_LABEL:
POLYNOMIAL_CHEBYCEV_FUNC = Err.number
End Function
