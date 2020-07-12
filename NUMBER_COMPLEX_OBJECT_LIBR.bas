Attribute VB_Name = "NUMBER_COMPLEX_OBJECT_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.

Public Type Cplx
    reel As Double
    imag As Double
End Type


'************************************************************************************
'************************************************************************************
'FUNCTION      : COMPLEX_COEFFICIENT_OBJ_FUNC
'DESCRIPTION   : Returns the real coefficient of a complex number
'in x + yi or x + yj text format
'LIBRARY       : NUMBER_COMPLEX
'GROUP         : OBJECT
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/08/2008
'************************************************************************************
'************************************************************************************

Function COMPLEX_COEFFICIENT_OBJ_FUNC(ByVal DATA_STR As String, _
Optional ByVal CPLX_CHR_STR As String = "i") As Cplx

Dim TEMP_VECTOR As Variant
On Error GoTo ERROR_LABEL
    
    TEMP_VECTOR = COMPLEX_CONVERT_STRING_FUNC(DATA_STR, CPLX_CHR_STR)
    
    COMPLEX_COEFFICIENT_OBJ_FUNC.reel = TEMP_VECTOR(1, 1)
    COMPLEX_COEFFICIENT_OBJ_FUNC.imag = TEMP_VECTOR(2, 1)

Exit Function
ERROR_LABEL:
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : COMPLEX_CONCATENATE_OBJ_FUNC
'DESCRIPTION   : Converts real and imaginary coefficients into a
'complex number of the form x + yi or x + yj
'LIBRARY       : NUMBER_COMPLEX
'GROUP         : OBJECT
'ID            : 002
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/08/2008
'************************************************************************************
'************************************************************************************


Function COMPLEX_CONCATENATE_OBJ_FUNC(ByRef COMPLEX_OBJ As Cplx, _
Optional ByVal CPLX_CHR_STR As String = "i")
Dim epsilon As Double
On Error GoTo ERROR_LABEL
    epsilon = 5 * 10 ^ -15
    COMPLEX_CONCATENATE_OBJ_FUNC = COMPLEX_CONVERT_COEFFICIENTS_FUNC(COMPLEX_OBJ.reel, _
                          COMPLEX_OBJ.imag, CPLX_CHR_STR, epsilon)
Exit Function
ERROR_LABEL:
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : COMPLEX_MODULUS_OBJ_FUNC
'DESCRIPTION   : Returns the absolute value (modulus) of a complex
'number in x + yi or x + yj text format
'LIBRARY       : NUMBER_COMPLEX
'GROUP         : OBJECT
'ID            : 003
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/08/2008
'************************************************************************************
'************************************************************************************

Function COMPLEX_MODULUS_OBJ_FUNC(ByRef COMPLEX_OBJ As Cplx) 'Same as Complex ABS
    
    Dim XTEMP_VAL As Double
    Dim YTEMP_VAL As Double
    Dim ZTEMP_VAL As Double
    
    On Error GoTo ERROR_LABEL
    
    XTEMP_VAL = Abs(COMPLEX_OBJ.reel)
    YTEMP_VAL = Abs(COMPLEX_OBJ.imag)
    If XTEMP_VAL = 0 Then
        COMPLEX_MODULUS_OBJ_FUNC = YTEMP_VAL
    ElseIf YTEMP_VAL = 0 Then
        COMPLEX_MODULUS_OBJ_FUNC = XTEMP_VAL
    ElseIf XTEMP_VAL > YTEMP_VAL Then
        ZTEMP_VAL = YTEMP_VAL / XTEMP_VAL
        COMPLEX_MODULUS_OBJ_FUNC = XTEMP_VAL * Sqr(1 + ZTEMP_VAL ^ 2)
    Else
        ZTEMP_VAL = XTEMP_VAL / YTEMP_VAL
        COMPLEX_MODULUS_OBJ_FUNC = YTEMP_VAL * Sqr(1 + ZTEMP_VAL ^ 2)
    End If

Exit Function
ERROR_LABEL:
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : COMPLEX_ZERO_OBJ_FUNC
'DESCRIPTION   : Complex Zero Tolerance
'LIBRARY       : NUMBER_COMPLEX
'GROUP         : OBJECT
'ID            : 004
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/08/2008
'************************************************************************************
'************************************************************************************

Function COMPLEX_ZERO_OBJ_FUNC(ByRef COMPLEX_OBJ As Cplx, _
Optional ByVal epsilon As Double = 0.000000000000001)
On Error GoTo ERROR_LABEL
COMPLEX_ZERO_OBJ_FUNC = Abs(COMPLEX_OBJ.reel) <= epsilon And _
                        Abs(COMPLEX_OBJ.imag) <= epsilon
Exit Function
ERROR_LABEL:
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : COMPLEX_NEGATIVE_OBJ_FUNC
'DESCRIPTION   : Negative of a complex number
'LIBRARY       : NUMBER_COMPLEX
'GROUP         : OBJECT
'ID            : 005
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/08/2008
'************************************************************************************
'************************************************************************************

Function COMPLEX_NEGATIVE_OBJ_FUNC(ByRef COMPLEX_OBJ As Cplx) As Cplx
On Error GoTo ERROR_LABEL
    COMPLEX_NEGATIVE_OBJ_FUNC.reel = -COMPLEX_OBJ.reel
    COMPLEX_NEGATIVE_OBJ_FUNC.imag = -COMPLEX_OBJ.imag
Exit Function
ERROR_LABEL:
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : COMPLEX_INVERSE_OBJ_FUNC
'DESCRIPTION   : Inverse of a number complex number
'LIBRARY       : NUMBER_COMPLEX
'GROUP         : OBJECT
'ID            : 006
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/08/2008
'************************************************************************************
'************************************************************************************

Function COMPLEX_INVERSE_OBJ_FUNC(ByRef COMPLEX_OBJ As Cplx) As Cplx
Dim TEMP_OBJ As Cplx
On Error GoTo ERROR_LABEL
TEMP_OBJ.reel = 1
TEMP_OBJ.imag = 0
COMPLEX_INVERSE_OBJ_FUNC = COMPLEX_QUOTIENT_OBJ_FUNC(TEMP_OBJ, COMPLEX_OBJ)
Exit Function
ERROR_LABEL:
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : COMPLEX_LN_OBJ_FUNC
'DESCRIPTION   : Returns the Ln of a complex number
'LIBRARY       : NUMBER_COMPLEX
'GROUP         : OBJECT
'ID            : 007
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/08/2008
'************************************************************************************
'************************************************************************************

Function COMPLEX_LN_OBJ_FUNC(ByRef COMPLEX_OBJ As Cplx) As Cplx
' Complex Log function
' ln(rh) + atan(ph) *  i
' rh = radix(ATEMP^1 + y^2)
' ph = atan(y / ATEMP)        -pi <  ph  < +pi
    Dim TEMP_OBJ As Cplx
    On Error GoTo ERROR_LABEL
    TEMP_OBJ = COMPLEX_POLAR_OBJ_FUNC(COMPLEX_OBJ)
    COMPLEX_LN_OBJ_FUNC.reel = Log(TEMP_OBJ.reel)
    COMPLEX_LN_OBJ_FUNC.imag = TEMP_OBJ.imag
Exit Function
ERROR_LABEL:
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : COMPLEX_EXPONENTIAL_OBJ_FUNC
'DESCRIPTION   : Returns the exponential of a complex number
'LIBRARY       : NUMBER_COMPLEX
'GROUP         : OBJECT
'ID            : 008
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/08/2008
'************************************************************************************
'************************************************************************************

Function COMPLEX_EXPONENTIAL_OBJ_FUNC(ByRef COMPLEX_OBJ As Cplx) As Cplx
' Complex Exp function
' exp(ATEMP) * (Cos(BTEMP) + Sin(BTEMP) * i))
    Dim TEMP_VAL As Double
    On Error GoTo ERROR_LABEL
    TEMP_VAL = Exp(COMPLEX_OBJ.reel)
    COMPLEX_EXPONENTIAL_OBJ_FUNC.reel = TEMP_VAL * Cos(COMPLEX_OBJ.imag)
    COMPLEX_EXPONENTIAL_OBJ_FUNC.imag = TEMP_VAL * Sin(COMPLEX_OBJ.imag)
Exit Function
ERROR_LABEL:
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : COMPLEX_CONJUGATE_OBJ_FUNC
'DESCRIPTION   : Returns the complex conjugate of a complex number
'LIBRARY       : NUMBER_COMPLEX
'GROUP         : OBJECT
'ID            : 009
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/08/2008
'************************************************************************************
'************************************************************************************

Function COMPLEX_CONJUGATE_OBJ_FUNC(ByRef COMPLEX_OBJ As Cplx) As Cplx
    
    On Error GoTo ERROR_LABEL
    
    COMPLEX_CONJUGATE_OBJ_FUNC.reel = COMPLEX_OBJ.reel
    COMPLEX_CONJUGATE_OBJ_FUNC.imag = -COMPLEX_OBJ.imag

Exit Function
ERROR_LABEL:
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : COMPLEX_SUM_OBJ_FUNC
'DESCRIPTION   : Returns the sum of two complex numbers
'LIBRARY       : NUMBER_COMPLEX
'GROUP         : OBJECT
'ID            : 010
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/08/2008
'************************************************************************************
'************************************************************************************

Function COMPLEX_SUM_OBJ_FUNC(ByRef ACOMPLEX_OBJ As Cplx, _
ByRef BCOMPLEX_OBJ As Cplx) As Cplx
    
On Error GoTo ERROR_LABEL

    COMPLEX_SUM_OBJ_FUNC.reel = ACOMPLEX_OBJ.reel + BCOMPLEX_OBJ.reel
    COMPLEX_SUM_OBJ_FUNC.imag = ACOMPLEX_OBJ.imag + BCOMPLEX_OBJ.imag

Exit Function
ERROR_LABEL:
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : COMPLEX_SUBTRACTION_OBJ_FUNC
'DESCRIPTION   : Returns the subtraction of two complex numbers
'LIBRARY       : NUMBER_COMPLEX
'GROUP         : OBJECT
'ID            : 011
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/08/2008
'************************************************************************************
'************************************************************************************


Function COMPLEX_SUBTRACTION_OBJ_FUNC(ByRef ACOMPLEX_OBJ As Cplx, _
ByRef BCOMPLEX_OBJ As Cplx) As Cplx
    
On Error GoTo ERROR_LABEL

    COMPLEX_SUBTRACTION_OBJ_FUNC.reel = ACOMPLEX_OBJ.reel - BCOMPLEX_OBJ.reel
    COMPLEX_SUBTRACTION_OBJ_FUNC.imag = ACOMPLEX_OBJ.imag - BCOMPLEX_OBJ.imag

Exit Function
ERROR_LABEL:
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : COMPLEX_PRODUCT_OBJ_FUNC
'DESCRIPTION   : Returns the product of two complex numbers
'LIBRARY       : NUMBER_COMPLEX
'GROUP         : OBJECT
'ID            : 012
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/08/2008
'************************************************************************************
'************************************************************************************


Function COMPLEX_PRODUCT_OBJ_FUNC(ByRef ACOMPLEX_OBJ As Cplx, _
ByRef BCOMPLEX_OBJ As Cplx) As Cplx
    
On Error GoTo ERROR_LABEL

COMPLEX_PRODUCT_OBJ_FUNC.reel = ACOMPLEX_OBJ.reel * BCOMPLEX_OBJ.reel - _
                                ACOMPLEX_OBJ.imag * BCOMPLEX_OBJ.imag

COMPLEX_PRODUCT_OBJ_FUNC.imag = ACOMPLEX_OBJ.imag * BCOMPLEX_OBJ.reel + _
                                ACOMPLEX_OBJ.reel * BCOMPLEX_OBJ.imag

Exit Function
ERROR_LABEL:
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : COMPLEX_SCALAR_PRODUCT_OBJ_FUNC
'DESCRIPTION   : Returns the product of a complex numbera and a scalar
'LIBRARY       : NUMBER_COMPLEX
'GROUP         : OBJECT
'ID            : 013
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/08/2008
'************************************************************************************
'************************************************************************************


Function COMPLEX_SCALAR_PRODUCT_OBJ_FUNC(ByRef COMPLEX_OBJ As Cplx, _
ByVal FACTOR_VAL As Double) As Cplx
    
On Error GoTo ERROR_LABEL

    COMPLEX_SCALAR_PRODUCT_OBJ_FUNC.reel = COMPLEX_OBJ.reel * FACTOR_VAL
    COMPLEX_SCALAR_PRODUCT_OBJ_FUNC.imag = COMPLEX_OBJ.imag * FACTOR_VAL

Exit Function
ERROR_LABEL:
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : COMPLEX_QUOTIENT_OBJ_FUNC
'DESCRIPTION   : Returns the quotient of two complex numbers
'LIBRARY       : NUMBER_COMPLEX
'GROUP         : OBJECT
'ID            : 014
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/08/2008
'************************************************************************************
'************************************************************************************

Function COMPLEX_QUOTIENT_OBJ_FUNC(ByRef ACOMPLEX_OBJ As Cplx, _
ByRef BCOMPLEX_OBJ As Cplx) As Cplx
    
    Dim TEMP_MULT As Double
    Dim TEMP_SCALAR As Double
    
    On Error GoTo ERROR_LABEL
    
    If Abs(BCOMPLEX_OBJ.reel) >= Abs(BCOMPLEX_OBJ.imag) Then
        TEMP_SCALAR = BCOMPLEX_OBJ.imag / BCOMPLEX_OBJ.reel
        TEMP_MULT = BCOMPLEX_OBJ.reel + TEMP_SCALAR * BCOMPLEX_OBJ.imag
        COMPLEX_QUOTIENT_OBJ_FUNC.reel = (ACOMPLEX_OBJ.reel + _
            TEMP_SCALAR * ACOMPLEX_OBJ.imag) / TEMP_MULT
        COMPLEX_QUOTIENT_OBJ_FUNC.imag = (ACOMPLEX_OBJ.imag - _
            TEMP_SCALAR * ACOMPLEX_OBJ.reel) / TEMP_MULT
    Else
        TEMP_SCALAR = BCOMPLEX_OBJ.reel / BCOMPLEX_OBJ.imag
        TEMP_MULT = BCOMPLEX_OBJ.imag + TEMP_SCALAR * BCOMPLEX_OBJ.reel
        COMPLEX_QUOTIENT_OBJ_FUNC.reel = (ACOMPLEX_OBJ.reel * _
            TEMP_SCALAR + ACOMPLEX_OBJ.imag) / TEMP_MULT
        COMPLEX_QUOTIENT_OBJ_FUNC.imag = (ACOMPLEX_OBJ.imag * _
            TEMP_SCALAR - ACOMPLEX_OBJ.reel) / TEMP_MULT
    End If

Exit Function
ERROR_LABEL:
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : COMPLEX_POWER_OBJ_FUNC
'DESCRIPTION   : Returns the power of a complex number
'LIBRARY       : NUMBER_COMPLEX
'GROUP         : OBJECT
'ID            : 015
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/08/2008
'************************************************************************************
'************************************************************************************


Function COMPLEX_POWER_OBJ_FUNC(ByRef COMPLEX_OBJ As Cplx, _
ByVal EXPONENT_VAL As Double) As Cplx
    
    Dim TEMP_OBJ As Cplx
    
    On Error GoTo ERROR_LABEL
    
    If COMPLEX_OBJ.imag = 0 Then
        COMPLEX_POWER_OBJ_FUNC.reel = COMPLEX_OBJ.reel ^ EXPONENT_VAL
        COMPLEX_POWER_OBJ_FUNC.imag = 0
    Else
        TEMP_OBJ = COMPLEX_POLAR_OBJ_FUNC(COMPLEX_OBJ)
        TEMP_OBJ.reel = TEMP_OBJ.reel ^ EXPONENT_VAL
        TEMP_OBJ.imag = TEMP_OBJ.imag * EXPONENT_VAL
        COMPLEX_POWER_OBJ_FUNC = COMPLEX_RECT_OBJ_FUNC(TEMP_OBJ)
    End If

Exit Function
ERROR_LABEL:
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : COMPLEX_ROOT_OBJ_FUNC
'DESCRIPTION   : Returns root of a complex number
'LIBRARY       : NUMBER_COMPLEX
'GROUP         : OBJECT
'ID            : 016
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/08/2008
'************************************************************************************
'************************************************************************************

Function COMPLEX_ROOT_OBJ_FUNC(ByRef ACOMPLEX_OBJ As Cplx, _
ByVal FACTOR_VAL As Double) As Cplx
    
    Dim PI_VAL As Double
    
    Dim TEMP_OBJ As Cplx
    
    Dim MTEMP_VAL As Double
    Dim TTEMP_VAL As Double
    Dim FTEMP_VAL As Double
    
    On Error GoTo ERROR_LABEL
    
    PI_VAL = 3.14159265358979
    
    If ACOMPLEX_OBJ.imag = 0 Then
        If ACOMPLEX_OBJ.reel >= 0 Then
            COMPLEX_ROOT_OBJ_FUNC.reel = Sqr(ACOMPLEX_OBJ.reel)
            COMPLEX_ROOT_OBJ_FUNC.imag = 0
        Else
            COMPLEX_ROOT_OBJ_FUNC.reel = 0
            COMPLEX_ROOT_OBJ_FUNC.imag = Sqr(Abs(ACOMPLEX_OBJ.reel))
        End If
    Else
        TEMP_OBJ = COMPLEX_POLAR_OBJ_FUNC(ACOMPLEX_OBJ)
        MTEMP_VAL = TEMP_OBJ.reel ^ (1 / FACTOR_VAL)
        TTEMP_VAL = TEMP_OBJ.imag
        FTEMP_VAL = (TTEMP_VAL + 2 * PI_VAL) / FACTOR_VAL
        COMPLEX_ROOT_OBJ_FUNC.reel = MTEMP_VAL * Cos(FTEMP_VAL)
        COMPLEX_ROOT_OBJ_FUNC.imag = MTEMP_VAL * Sin(FTEMP_VAL)
    End If

Exit Function
ERROR_LABEL:
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : COMPLEX_POLAR_OBJ_FUNC
'DESCRIPTION   : Returns the Complex Polar number
'LIBRARY       : NUMBER_COMPLEX
'GROUP         : OBJECT
'ID            : 017
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/08/2008
'************************************************************************************
'************************************************************************************

Function COMPLEX_POLAR_OBJ_FUNC(ByRef COMPLEX_OBJ As Cplx) As Cplx

Dim PI_VAL As Double

Dim XTEMP_VAL As Double
Dim YTEMP_VAL As Double

On Error GoTo ERROR_LABEL

    PI_VAL = 3.14159265358979

    XTEMP_VAL = COMPLEX_OBJ.reel
    YTEMP_VAL = COMPLEX_OBJ.imag
    COMPLEX_POLAR_OBJ_FUNC.reel = COMPLEX_MODULUS_OBJ_FUNC(COMPLEX_OBJ)

    If XTEMP_VAL = 0 Then
        COMPLEX_POLAR_OBJ_FUNC.imag = Sgn(YTEMP_VAL) * PI_VAL / 2
    ElseIf XTEMP_VAL > 0 Then
        COMPLEX_POLAR_OBJ_FUNC.imag = Atn(YTEMP_VAL / XTEMP_VAL)
    Else
        If YTEMP_VAL <> 0 Then
             COMPLEX_POLAR_OBJ_FUNC.imag = Sgn(YTEMP_VAL) * PI_VAL + Atn(YTEMP_VAL / XTEMP_VAL)
        Else
             COMPLEX_POLAR_OBJ_FUNC.imag = PI_VAL
        End If
    End If

Exit Function
ERROR_LABEL:
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : COMPLEX_RECT_OBJ_FUNC
'DESCRIPTION   : Returns the complex rect value (Sin(r) - Cos(i))
'LIBRARY       : NUMBER_COMPLEX
'GROUP         : OBJECT
'ID            : 018
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/08/2008
'************************************************************************************
'************************************************************************************

Function COMPLEX_RECT_OBJ_FUNC(ByRef COMPLEX_OBJ As Cplx) As Cplx

Dim MTEMP_VAL As Double
Dim TTEMP_VAL As Double

On Error GoTo ERROR_LABEL

MTEMP_VAL = COMPLEX_OBJ.reel
TTEMP_VAL = COMPLEX_OBJ.imag

If MTEMP_VAL = 0 Then
    COMPLEX_RECT_OBJ_FUNC.reel = 0
    COMPLEX_RECT_OBJ_FUNC.imag = 0
    Exit Function
ElseIf TTEMP_VAL = 0 Then
    COMPLEX_RECT_OBJ_FUNC.reel = MTEMP_VAL
    COMPLEX_RECT_OBJ_FUNC.imag = 0
Else
    COMPLEX_RECT_OBJ_FUNC.reel = MTEMP_VAL * Cos(TTEMP_VAL)
    COMPLEX_RECT_OBJ_FUNC.imag = MTEMP_VAL * Sin(TTEMP_VAL)
End If

Exit Function
ERROR_LABEL:
End Function

