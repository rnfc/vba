Attribute VB_Name = "NUMBER_COMPLEX_ARITHMETIC_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.


'************************************************************************************
'************************************************************************************
'FUNCTION      : COMPLEX_CONJUGATE_FUNC
'DESCRIPTION   : Returns the complex conjugate of a complex number in
'x + yi or x + yj text format.
'LIBRARY       : NUMBER_COMPLEX
'GROUP         : ARITHMETIC
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/08/2008
'************************************************************************************
'************************************************************************************


Function COMPLEX_CONJUGATE_FUNC(ByVal DATA_STR As String, _
Optional ByVal CPLX_CHR_STR As String = "i")

On Error GoTo ERROR_LABEL

COMPLEX_CONJUGATE_FUNC = _
COMPLEX_CONCATENATE_OBJ_FUNC(COMPLEX_CONJUGATE( _
COMPLEX_COEFFICIENT_OBJ_FUNC(DATA_STR, CPLX_CHR_STR)))

Exit Function
ERROR_LABEL:
COMPLEX_CONJUGATE_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : COMPLEX_MODULUS_FUNC
'DESCRIPTION   : Returns the absolute value (modulus) of a complex number in
'x + yi or x + yj text format.
'LIBRARY       : NUMBER_COMPLEX
'GROUP         : ARITHMETIC
'ID            : 002
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/08/2008
'************************************************************************************
'************************************************************************************


Function COMPLEX_MODULUS_FUNC(ByVal DATA_STR As String, _
Optional ByVal CPLX_CHR_STR As String = "i")

On Error GoTo ERROR_LABEL

COMPLEX_MODULUS_FUNC = _
COMPLEX_MODULUS_OBJ_FUNC( _
COMPLEX_COEFFICIENT_OBJ_FUNC(DATA_STR, CPLX_CHR_STR))

Exit Function
ERROR_LABEL:
COMPLEX_MODULUS_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : COMPLEX_EXPONENTIAL_FUNC
'DESCRIPTION   : Returns the exponential of a complex number in
'x + yi or x + yj text format.
'LIBRARY       : NUMBER_COMPLEX
'GROUP         : ARITHMETIC
'ID            : 003
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/08/2008
'************************************************************************************
'************************************************************************************


Function COMPLEX_EXPONENTIAL_FUNC(ByVal DATA_STR As String, _
Optional ByVal CPLX_CHR_STR As String = "i")

On Error GoTo ERROR_LABEL

COMPLEX_EXPONENTIAL_FUNC = _
COMPLEX_CONCATENATE_OBJ_FUNC(COMPLEX_EXPONENTIAL_OBJ_FUNC( _
COMPLEX_COEFFICIENT_OBJ_FUNC(DATA_STR, CPLX_CHR_STR)))

Exit Function
ERROR_LABEL:
COMPLEX_EXPONENTIAL_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : COMPLEX_LN_FUNC
'DESCRIPTION   : Returns the Ln of a complex number in
'x + yi or x + yj text format.
'LIBRARY       : NUMBER_COMPLEX
'GROUP         : ARITHMETIC
'ID            : 004
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/08/2008
'************************************************************************************
'************************************************************************************


Function COMPLEX_LN_FUNC(ByVal DATA_STR As String, _
Optional ByVal CPLX_CHR_STR As String = "i")

On Error GoTo ERROR_LABEL

COMPLEX_LN_FUNC = _
COMPLEX_CONCATENATE_OBJ_FUNC(COMPLEX_LN_OBJ_FUNC( _
COMPLEX_COEFFICIENT_OBJ_FUNC(DATA_STR, CPLX_CHR_STR)))

Exit Function
ERROR_LABEL:
COMPLEX_LN_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : COMPLEX_NEGATIVE_FUNC
'DESCRIPTION   : Negative of a complex number in
'x + yi or x + yj text format.
'LIBRARY       : NUMBER_COMPLEX
'GROUP         : ARITHMETIC
'ID            : 005
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/08/2008
'************************************************************************************
'************************************************************************************


Function COMPLEX_NEGATIVE_FUNC(ByVal DATA_STR As String, _
Optional ByVal CPLX_CHR_STR As String = "i")

On Error GoTo ERROR_LABEL

COMPLEX_NEGATIVE_FUNC = _
COMPLEX_CONCATENATE_OBJ_FUNC(COMPLEX_NEGATIVE_OBJ_FUNC( _
COMPLEX_COEFFICIENT_OBJ_FUNC(DATA_STR, CPLX_CHR_STR)))

Exit Function
ERROR_LABEL:
COMPLEX_NEGATIVE_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : COMPLEX_INVERSE_FUNC
'DESCRIPTION   : Inverse of a number complex number of a complex number in
'x + yi or x + yj text format.
'LIBRARY       : NUMBER_COMPLEX
'GROUP         : ARITHMETIC
'ID            : 006
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/08/2008
'************************************************************************************
'************************************************************************************


Function COMPLEX_INVERSE_FUNC(ByVal DATA_STR As String, _
Optional ByVal CPLX_CHR_STR As String = "i")

On Error GoTo ERROR_LABEL

COMPLEX_INVERSE_FUNC = _
COMPLEX_CONCATENATE_OBJ_FUNC(COMPLEX_INVERSE_OBJ_FUNC( _
COMPLEX_COEFFICIENT_OBJ_FUNC(DATA_STR, CPLX_CHR_STR)))

Exit Function
ERROR_LABEL:
COMPLEX_INVERSE_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : COMPLEX_SUM_FUNC
'DESCRIPTION   : Returns the sum of two complex numbers in
'x + yi or x + yj text format.
'LIBRARY       : NUMBER_COMPLEX
'GROUP         : ARITHMETIC
'ID            : 007
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/08/2008
'************************************************************************************
'************************************************************************************

'// PERFECT
Function COMPLEX_SUM_FUNC(ByVal ADATA_STR As String, _
ByVal BDATA_STR As String, _
Optional ByVal CPLX_CHR_STR As String = "i")

On Error GoTo ERROR_LABEL

COMPLEX_SUM_FUNC = COMPLEX_CONCATENATE_OBJ_FUNC(COMPLEX_SUM_OBJ_FUNC( _
COMPLEX_COEFFICIENT_OBJ_FUNC(ADATA_STR, CPLX_CHR_STR), _
COMPLEX_COEFFICIENT_OBJ_FUNC(BDATA_STR, CPLX_CHR_STR)), _
CPLX_CHR_STR)

Exit Function
ERROR_LABEL:
COMPLEX_SUM_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : COMPLEX_SUBTRACTION_FUNC
'DESCRIPTION   : Returns the subtraction of two complex numbers in
'x + yi or x + yj text format.
'LIBRARY       : NUMBER_COMPLEX
'GROUP         : ARITHMETIC
'ID            : 008
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/08/2008
'************************************************************************************
'************************************************************************************


Function COMPLEX_SUBTRACTION_FUNC(ByVal ADATA_STR As String, _
ByVal BDATA_STR As String, _
Optional ByVal CPLX_CHR_STR As String = "i")

On Error GoTo ERROR_LABEL

COMPLEX_SUBTRACTION_FUNC = COMPLEX_CONCATENATE_OBJ_FUNC(COMPLEX_SUBTRACTION_OBJ_FUNC( _
COMPLEX_COEFFICIENT_OBJ_FUNC(ADATA_STR, CPLX_CHR_STR), _
COMPLEX_COEFFICIENT_OBJ_FUNC(BDATA_STR, CPLX_CHR_STR)), _
CPLX_CHR_STR)

Exit Function
ERROR_LABEL:
COMPLEX_SUBTRACTION_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : COMPLEX_PRODUCT_FUNC
'DESCRIPTION   : Returns the product of two complex numbers in
'x + yi or x + yj text format.
'LIBRARY       : NUMBER_COMPLEX
'GROUP         : ARITHMETIC
'ID            : 009
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/08/2008
'************************************************************************************
'************************************************************************************


Function COMPLEX_PRODUCT_FUNC(ByVal ADATA_STR As String, _
ByVal BDATA_STR As String, _
Optional ByVal CPLX_CHR_STR As String = "i")

On Error GoTo ERROR_LABEL

COMPLEX_PRODUCT_FUNC = _
COMPLEX_CONCATENATE_OBJ_FUNC(COMPLEX_PRODUCT_OBJ_FUNC( _
COMPLEX_COEFFICIENT_OBJ_FUNC(ADATA_STR, CPLX_CHR_STR), _
COMPLEX_COEFFICIENT_OBJ_FUNC(BDATA_STR, CPLX_CHR_STR)), _
CPLX_CHR_STR)

Exit Function
ERROR_LABEL:
COMPLEX_PRODUCT_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : COMPLEX_SCALAR_PRODUCT_FUNC
'DESCRIPTION   : Returns the product of a complex number a scalar in
'x + yi or x + yj text format.
'LIBRARY       : NUMBER_COMPLEX
'GROUP         : ARITHMETIC
'ID            : 010
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/08/2008
'************************************************************************************
'************************************************************************************


Function COMPLEX_SCALAR_PRODUCT_FUNC(ByVal DATA_STR As String, _
ByVal SCALAR_VAL As Double, _
Optional ByVal CPLX_CHR_STR As String = "i")

On Error GoTo ERROR_LABEL

COMPLEX_SCALAR_PRODUCT_FUNC = _
COMPLEX_CONCATENATE_OBJ_FUNC(COMPLEX_SCALAR_PRODUCT_OBJ_FUNC( _
COMPLEX_COEFFICIENT_OBJ_FUNC(DATA_STR, CPLX_CHR_STR), _
SCALAR_VAL), CPLX_CHR_STR)

Exit Function
ERROR_LABEL:
COMPLEX_SCALAR_PRODUCT_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : COMPLEX_POWER_FUNC
'DESCRIPTION   : Returns the power of a complex number in
'x + yi or x + yj text format.
'LIBRARY       : NUMBER_COMPLEX
'GROUP         : ARITHMETIC
'ID            : 011
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/08/2008
'************************************************************************************
'************************************************************************************


Function COMPLEX_POWER_FUNC(ByVal DATA_STR As String, _
ByVal EXPONENT_VAL As Double, _
Optional ByVal CPLX_CHR_STR As String = "i")

On Error GoTo ERROR_LABEL

COMPLEX_POWER_FUNC = _
COMPLEX_CONCATENATE_OBJ_FUNC(COMPLEX_POWER_OBJ_FUNC( _
COMPLEX_COEFFICIENT_OBJ_FUNC(DATA_STR, _
CPLX_CHR_STR), EXPONENT_VAL), CPLX_CHR_STR)

Exit Function
ERROR_LABEL:
COMPLEX_POWER_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : COMPLEX_QUOTIENT_FUNC
'DESCRIPTION   : Returns the quotient of a complex number in
'x + yi or x + yj text format.
'LIBRARY       : NUMBER_COMPLEX
'GROUP         : ARITHMETIC
'ID            : 012
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/08/2008
'************************************************************************************
'************************************************************************************


Function COMPLEX_QUOTIENT_FUNC(ByVal ADATA_STR As String, _
ByVal BDATA_STR As String, _
Optional ByVal CPLX_CHR_STR As String = "i")

On Error GoTo ERROR_LABEL

COMPLEX_QUOTIENT_FUNC = _
COMPLEX_CONCATENATE_OBJ_FUNC(COMPLEX_QUOTIENT_OBJ_FUNC( _
COMPLEX_COEFFICIENT_OBJ_FUNC(ADATA_STR, CPLX_CHR_STR), _
COMPLEX_COEFFICIENT_OBJ_FUNC(BDATA_STR, CPLX_CHR_STR)), _
CPLX_CHR_STR)

Exit Function
ERROR_LABEL:
COMPLEX_QUOTIENT_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : COMPLEX_ROOT_FUNC
'DESCRIPTION   : Returns the root of a complex number in
'x + yi or x + yj text format.
'LIBRARY       : NUMBER_COMPLEX
'GROUP         : ARITHMETIC
'ID            : 013
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/08/2008
'************************************************************************************
'************************************************************************************


Function COMPLEX_ROOT_FUNC(ByVal DATA_STR As String, _
ByVal ROOT_VAL As Double, _
Optional ByVal CPLX_CHR_STR As String = "i")

Dim TEMP_STR As String

On Error GoTo ERROR_LABEL

TEMP_STR = _
COMPLEX_CONCATENATE_OBJ_FUNC(COMPLEX_ROOT_OBJ_FUNC(COMPLEX_COEFFICIENT_OBJ_FUNC(DATA_STR, _
CPLX_CHR_STR), ROOT_VAL), CPLX_CHR_STR) 'SECOND ROOT
    
COMPLEX_ROOT_FUNC = COMPLEX_NEGATIVE_FUNC(TEMP_STR, CPLX_CHR_STR)
    'FIRST ROOT
Exit Function
ERROR_LABEL:
COMPLEX_ROOT_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : COMPLEX_QUOTIENT_ARRAY_FUNC
'DESCRIPTION   :
'LIBRARY       : NUMBER_COMPLEX
'GROUP         : ARITHMETIC
'ID            : 014
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/08/2008
'************************************************************************************
'************************************************************************************

Function COMPLEX_QUOTIENT_ARRAY_FUNC(ByVal XR_VAL As Double, _
ByVal XI_VAL As Double, _
ByVal YR_VAL As Double, _
ByVal YI_VAL As Double)
  
  Dim R_VAL As Double
  Dim D_VAL As Double
  
  Dim CPLX_DIV_REAL_VAL As Double
  Dim CPLX_DIV_IMAG_VAL As Double
  
  On Error GoTo ERROR_LABEL
  
  If (Abs(YR_VAL) > Abs(YI_VAL)) Then
    R_VAL = YI_VAL / YR_VAL
    D_VAL = YR_VAL + R_VAL * YI_VAL
    CPLX_DIV_REAL_VAL = (XR_VAL + R_VAL * XI_VAL) / D_VAL
    CPLX_DIV_IMAG_VAL = (XI_VAL - R_VAL * XR_VAL) / D_VAL
  Else
    R_VAL = YR_VAL / YI_VAL
    D_VAL = YI_VAL + R_VAL * YR_VAL
    CPLX_DIV_REAL_VAL = (R_VAL * XR_VAL + XI_VAL) / D_VAL
    CPLX_DIV_IMAG_VAL = (R_VAL * XI_VAL - XR_VAL) / D_VAL
  End If

  COMPLEX_QUOTIENT_ARRAY_FUNC = Array(CPLX_DIV_REAL_VAL, CPLX_DIV_IMAG_VAL)

Exit Function
ERROR_LABEL:
COMPLEX_QUOTIENT_ARRAY_FUNC = Err.number
End Function
