Attribute VB_Name = "FINAN_FI_BOND_INTERPOLATE_LIBR"
Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.

'************************************************************************************
'************************************************************************************
'FUNCTION      : YIELD_POLYNOMIAL_INTERPOLATION_FUNC
'DESCRIPTION   : Intrapolate Rates / Discounts through Least-squares regression
'with polynomials
'LIBRARY       : BOND
'GROUP         : INTERPOLATE
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 21/01/2009
'************************************************************************************
'************************************************************************************

Function YIELD_POLYNOMIAL_INTERPOLATION_FUNC(ByVal NDEG As Integer, _
ByVal SETTLEMENT As Date, _
ByRef XDATA_RNG As Variant, _
ByRef YDATA_RNG As Variant, _
Optional ByVal MATURITY As Variant, _
Optional ByVal OUTPUT As Integer = 0, _
Optional ByVal COUNT_BASIS As Integer = 0)

'MATURITY --> could be an array

Dim i As Long
Dim NROWS As Long

Dim XDATA_VECTOR As Variant
Dim YDATA_VECTOR As Variant

Dim XTEMP_VECTOR As Variant
Dim YTEMP_VECTOR As Variant

Dim COEF_VECTOR As Variant
Dim TEMP_MATRIX As Variant

On Error GoTo ERROR_LABEL

XDATA_VECTOR = XDATA_RNG
If UBound(XDATA_VECTOR, 1) = 1 Then
  XDATA_VECTOR = MATRIX_TRANSPOSE_FUNC(XDATA_VECTOR)
End If
YDATA_VECTOR = YDATA_RNG
If UBound(YDATA_VECTOR, 1) = 1 Then
  YDATA_VECTOR = MATRIX_TRANSPOSE_FUNC(YDATA_VECTOR)
End If
If UBound(YDATA_VECTOR, 1) <> UBound(YDATA_VECTOR, 1) Then: GoTo ERROR_LABEL
NROWS = UBound(XDATA_VECTOR, 1)
ReDim XTEMP_VECTOR(1 To NROWS, 1 To 1)
For i = 1 To NROWS
    XTEMP_VECTOR(i, 1) = YEARFRAC_FUNC(SETTLEMENT, XDATA_VECTOR(i, 1), COUNT_BASIS)
Next i

COEF_VECTOR = POLYNOMIAL_REGRESSION_FUNC(XTEMP_VECTOR, YDATA_VECTOR, CInt(NDEG), 0)
'calculate polynomial coefficients C

'---------------------------------------------------------------------------------
Select Case OUTPUT
'---------------------------------------------------------------------------------
Case 0
'---------------------------------------------------------------------------------
    YIELD_POLYNOMIAL_INTERPOLATION_FUNC = COEF_VECTOR
'---------------------------------------------------------------------------------
Case 1
'---------------------------------------------------------------------------------
    If IsArray(MATURITY) = False Then
        XTEMP_VECTOR = YEARFRAC_FUNC(SETTLEMENT, MATURITY, COUNT_BASIS)
    Else
        XTEMP_VECTOR = MATURITY
        If UBound(XTEMP_VECTOR, 1) = 1 Then
            XTEMP_VECTOR = MATRIX_TRANSPOSE_FUNC(XTEMP_VECTOR)
        End If
        For i = LBound(XTEMP_VECTOR, 1) To UBound(XTEMP_VECTOR, 1)
            XTEMP_VECTOR(i, 1) = YEARFRAC_FUNC(SETTLEMENT, XTEMP_VECTOR(i, 1), COUNT_BASIS)
        Next i
    End If
    YIELD_POLYNOMIAL_INTERPOLATION_FUNC = DISCOUNT_POLYNOMIAL_INTERPOLATION_FUNC(COEF_VECTOR, XTEMP_VECTOR)
'---------------------------------------------------------------------------------
Case 2
'---------------------------------------------------------------------------------
    YTEMP_VECTOR = DISCOUNT_POLYNOMIAL_INTERPOLATION_FUNC(COEF_VECTOR, XTEMP_VECTOR)
    NROWS = UBound(YTEMP_VECTOR, 1)
    ReDim TEMP_MATRIX(0 To NROWS, 1 To 2)
    TEMP_MATRIX(0, 1) = "X"
    TEMP_MATRIX(0, 2) = "Y"
    For i = 1 To NROWS
        TEMP_MATRIX(i, 1) = XTEMP_VECTOR(i, 1)
        TEMP_MATRIX(i, 2) = YTEMP_VECTOR(i, 1)
    Next i
    YIELD_POLYNOMIAL_INTERPOLATION_FUNC = TEMP_MATRIX
'---------------------------------------------------------------------------------
End Select
'---------------------------------------------------------------------------------
  
Exit Function
ERROR_LABEL:
YIELD_POLYNOMIAL_INTERPOLATION_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : DISCOUNT_POLYNOMIAL_INTERPOLATION_FUNC
'DESCRIPTION   : INTRAPOLATE DISCOUNT FACTORS FROM A POLYNOMIAL
'COEFFICIENT VECTOR
'LIBRARY       : BOND
'GROUP         : INTERPOLATE
'ID            : 002
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 21/01/2009
'************************************************************************************
'************************************************************************************

Function DISCOUNT_POLYNOMIAL_INTERPOLATION_FUNC(ByRef COEF_RNG As Variant, _
ByRef XDATA_RNG As Variant)
    
Dim i As Long
Dim j As Long
    
Dim NROWS As Long
Dim NDEG As Long
Dim XDATA_VECTOR As Variant
Dim TEMP_MATRIX As Variant
Dim COEF_VECTOR As Variant

On Error GoTo ERROR_LABEL

'------------------------------------------------------------------------------
If IsArray(XDATA_RNG) = True Then
'------------------------------------------------------------------------------
    COEF_VECTOR = COEF_RNG
    If UBound(COEF_VECTOR, 1) = 1 Then
        COEF_VECTOR = MATRIX_TRANSPOSE_FUNC(COEF_VECTOR)
    End If
    XDATA_VECTOR = XDATA_RNG
    If UBound(XDATA_VECTOR, 1) = 1 Then
        XDATA_VECTOR = MATRIX_TRANSPOSE_FUNC(XDATA_VECTOR)
    End If
    NROWS = UBound(XDATA_VECTOR, 1)
    NDEG = UBound(COEF_VECTOR, 1) '---> THIS IS EQUIVALENT to the No. of Degrees in the Polynomial minus 1
    ReDim TEMP_MATRIX(1 To NROWS, 1 To 1)
    For i = 1 To NROWS
        TEMP_MATRIX(i, 1) = 0
        For j = 1 To NDEG
            TEMP_MATRIX(i, 1) = TEMP_MATRIX(i, 1) + COEF_VECTOR(j, 1) * XDATA_VECTOR(i, 1) ^ (j - 1)
        Next j
    Next i
'------------------------------------------------------------------------------
Else
'------------------------------------------------------------------------------
    COEF_VECTOR = COEF_RNG
    If UBound(COEF_VECTOR, 1) = 1 Then
        COEF_VECTOR = MATRIX_TRANSPOSE_FUNC(COEF_VECTOR)
    End If
    NDEG = UBound(COEF_VECTOR, 1) '---> THIS IS EQUIVALENT to the No. of Degrees in the Polynomial minus 1
    ReDim TEMP_MATRIX(1 To 1, 1 To 1)
    For i = 1 To 1
        TEMP_MATRIX(i, 1) = 0
        For j = 1 To NDEG
            TEMP_MATRIX(i, 1) = TEMP_MATRIX(i, 1) + COEF_VECTOR(j, 1) * XDATA_RNG ^ (j - 1)
        Next j
    Next i
'------------------------------------------------------------------------------
End If
'------------------------------------------------------------------------------

DISCOUNT_POLYNOMIAL_INTERPOLATION_FUNC = TEMP_MATRIX
    
Exit Function
ERROR_LABEL:
DISCOUNT_POLYNOMIAL_INTERPOLATION_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : YIELD_INTERPOLATION_FUNC
'DESCRIPTION   :
'LIBRARY       : BOND
'GROUP         : INTERPOLATE
'ID            : 003
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 10-06-2008
'************************************************************************************
'************************************************************************************

Function YIELD_INTERPOLATION_FUNC(ByVal X0_VAL As Variant, _
ByRef XDATA_RNG As Variant, _
ByRef YDATA_RNG As Variant, _
Optional ByVal METHOD As Variant = 0, _
Optional ByVal OUTPUT As Integer = 0)

'METHOD:
'Case 0, "lin", "lor", ""
'Case 1, "log", "llr"
'Case 2, "lod"
'Case 3, "wei", "raw", "lld"
  
Dim i As Long
Dim j As Long

Dim NROWS As Long

Dim Y1_VAL As Variant
Dim Y2_VAL As Variant
Dim Y0_VAL As Variant

Dim X1_VAL As Variant
Dim X2_VAL As Variant

Dim XDATA_VECTOR As Variant
Dim YDATA_VECTOR As Variant
Dim TEMP_VECTOR As Variant

On Error GoTo ERROR_LABEL

XDATA_VECTOR = XDATA_RNG
If UBound(XDATA_VECTOR, 1) = 1 Then
    XDATA_VECTOR = MATRIX_TRANSPOSE_FUNC(XDATA_VECTOR)
End If

YDATA_VECTOR = YDATA_RNG
If UBound(YDATA_VECTOR, 1) = 1 Then
    YDATA_VECTOR = MATRIX_TRANSPOSE_FUNC(YDATA_VECTOR)
End If

If UBound(XDATA_VECTOR, 1) <> UBound(YDATA_VECTOR, 1) Then: GoTo ERROR_LABEL
NROWS = UBound(XDATA_VECTOR, 1)

If XDATA_VECTOR(NROWS, 1) <= X0_VAL Then
    YIELD_INTERPOLATION_FUNC = YDATA_VECTOR(NROWS, 1)
    Exit Function
End If

For i = 1 To NROWS
    If X0_VAL < XDATA_VECTOR(i, 1) Then
        j = i - 1
        Exit For
    ElseIf X0_VAL = XDATA_VECTOR(i, 1) Then
        j = i
        Exit For
    End If
Next i

Y1_VAL = YDATA_VECTOR(j, 1)
Y2_VAL = YDATA_VECTOR(j + 1, 1)

X1_VAL = XDATA_VECTOR(j, 1)
X2_VAL = XDATA_VECTOR(j + 1, 1)
Y0_VAL = INTERPOLATION_FUNC(X1_VAL, X2_VAL, Y1_VAL, Y2_VAL, X0_VAL, METHOD, 1)

'---------------------------------------------------------------------------
Select Case OUTPUT
'---------------------------------------------------------------------------
Case 0
    YIELD_INTERPOLATION_FUNC = Y0_VAL
'---------------------------------------------------------------------------
Case Else
'---------------------------------------------------------------------------
    ReDim TEMP_VECTOR(1 To 4, 1 To 2)
    
    TEMP_VECTOR(1, 1) = "DATES"
    TEMP_VECTOR(2, 1) = X1_VAL 'X1_VAL
    TEMP_VECTOR(3, 1) = X2_VAL 'X2_VAL
    TEMP_VECTOR(4, 1) = X0_VAL
    
    TEMP_VECTOR(1, 2) = "RATES"
    TEMP_VECTOR(2, 2) = Y1_VAL 'Y1_VAL
    TEMP_VECTOR(3, 2) = Y2_VAL 'Y2_VAL
    TEMP_VECTOR(4, 2) = Y0_VAL
    
    YIELD_INTERPOLATION_FUNC = TEMP_VECTOR
'---------------------------------------------------------------------------
End Select
'---------------------------------------------------------------------------
Exit Function
ERROR_LABEL:
YIELD_INTERPOLATION_FUNC = Err.number
End Function
