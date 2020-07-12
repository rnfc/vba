Attribute VB_Name = "POLYNOMIAL_PARSE_LIBR"


Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.


'************************************************************************************
'************************************************************************************
'FUNCTION      : WRITE_POLYNOMIAL_STRING_FUNC
'DESCRIPTION   : Write Polynomial String
'LIBRARY       : POLYNOMIAL
'GROUP         : PARSE
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/22/2009
'************************************************************************************
'************************************************************************************

Function WRITE_POLYNOMIAL_STRING_FUNC(ByRef COEF_RNG As Variant, _
Optional ByVal REF_VAR As String = "x")
'100+67x-x^2

Dim i As Long
Dim j As Long
Dim k As Long

Dim NDEG As Long

Dim VAR_STR As String
Dim TEMP_STR As String

Dim TEMP_VAL As Variant
Dim COEF_VAL As Variant

Dim POLYN_STR As String
Dim DEC_SEP_STR As String

Dim COEF_VECTOR As Variant

On Error GoTo ERROR_LABEL

COEF_VECTOR = COEF_RNG

If (UBound(COEF_VECTOR, 1) - LBound(COEF_VECTOR, 1) + 1) = 1 Then
    COEF_VECTOR = MATRIX_TRANSPOSE_FUNC(COEF_VECTOR)
End If

If LBound(COEF_VECTOR, 1) <> 0 Then
    COEF_VECTOR = MATRIX_CHANGE_BASE_ZERO_FUNC(COEF_VECTOR)
End If

NDEG = UBound(COEF_VECTOR, 1) 'begin  building string algorithm

DEC_SEP_STR = DECIMAL_SEPARATOR_FUNC()
POLYN_STR = ""
For i = 0 To NDEG
    j = i
    TEMP_VAL = COEF_VECTOR(i, LBound(COEF_VECTOR, 2))
    If TEMP_VAL <> 0 Then
        'set REF_VAR with power: x^2, x, nothing
        If j = 0 Then
            VAR_STR = ""
        ElseIf j = 1 Then
            VAR_STR = REF_VAR
        Else
            VAR_STR = REF_VAR & "^" & Trim(CStr(j))
        End If
        'set sign
        If TEMP_VAL > 0 Then
            If i > 0 Then
                TEMP_STR = "+"
            Else
                TEMP_STR = "" 'do not write the first sign +
            End If
        Else
            TEMP_STR = "-"   'write always sign -
        End If
        'set coefficient
        If Abs(CDbl(TEMP_VAL)) = 1 And j > 0 Then
            COEF_VAL = ""
            'do no write +1 and -1 before "x, x^2..."
        Else
            COEF_VAL = Trim(CStr(Abs(TEMP_VAL)))
            're-insert the initial zero
            If Left(COEF_VAL, 1) = "." Then COEF_VAL = "0" & COEF_VAL
            'substitute decimal separator
            k = InStr(1, COEF_VAL, ".")
            If k > 0 Then Mid(COEF_VAL, k, 1) = DEC_SEP_STR
        End If
            'build terms i-th : -3x^4, +x
        POLYN_STR = POLYN_STR & TEMP_STR & COEF_VAL & VAR_STR
    End If
Next

If POLYN_STR = "" Then POLYN_STR = "0"
If Len(POLYN_STR) > 255 Then POLYN_STR = "polynomial lenght >255"

TEMP_STR = POLYN_STR 'Trim + Sign
If Left(TEMP_STR, 1) = "+" Then
    POLYN_STR = Right(TEMP_STR, Len(TEMP_STR) - 1)
End If

WRITE_POLYNOMIAL_STRING_FUNC = POLYN_STR

Exit Function
ERROR_LABEL:
WRITE_POLYNOMIAL_STRING_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : PARSE_POLYNOMIAL_STRING_FUNC

'DESCRIPTION   : Converts a polynomial string into a vector coefficient
'"1 -5x +3x^2" -> [1,-5,2]
'"1/3 +4x^2" -> [0.33333, 0, 4]
'this routine reduce similar terms, i.e. 3x+9x -> 12x

'LIBRARY       : POLYNOMIAL
'GROUP         : PARSE
'ID            : 002
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/22/2009
'************************************************************************************
'************************************************************************************

Function PARSE_POLYNOMIAL_STRING_FUNC(ByVal DATA_STR As String, _
Optional ByVal OUTPUT As Integer = 0, _
Optional ByVal REF_CHR As String = "§")

Dim i As Long
Dim j As Long
Dim k As Long

Dim ii As Long
Dim jj As Long
Dim kk As Long

Dim NDEG As Long

Dim VAR_STR As String
Dim CHR_STR As String
Dim TEMP_STR As String
Dim POLY_STR As String

Dim TEMP_MATRIX As Variant
Dim TEMP_VECTOR As Variant

On Error GoTo ERROR_LABEL

ii = 100
ReDim TEMP_MATRIX(1 To 100, 1 To 3)

VAR_STR = ""
POLY_STR = DATA_STR & REF_CHR

i = 1
j = 1
k = 0

jj = Len(POLY_STR)

Do
    CHR_STR = Mid(POLY_STR, i, 1)
    If i > 1 Then
        If CHR_STR = "+" Or CHR_STR = "-" Or CHR_STR = REF_CHR Then
            TEMP_STR = Mid(POLY_STR, j, i - j)
            k = k + 1
            If k > ii Then: GoTo ERROR_LABEL 'too many terms"
            TEMP_VECTOR = PARSE_MONOMIAL_STRING_FUNC(TEMP_STR, REF_CHR)
                TEMP_MATRIX(k, 1) = TEMP_VECTOR(1, 1)
                TEMP_MATRIX(k, 2) = TEMP_VECTOR(2, 1)
                TEMP_MATRIX(k, 3) = TEMP_VECTOR(3, 1)
            j = i
        End If
    End If
    i = i + 1
Loop Until i > jj

' rebuild the polynomial coefficients
kk = k
NDEG = 0
For i = 1 To kk
    If NDEG < TEMP_MATRIX(i, 2) Then
          NDEG = TEMP_MATRIX(i, 2)
    Else: NDEG = i - 1
    End If
    If (TEMP_MATRIX(i, 2) > 0) And (TEMP_MATRIX(i, 3) <> "") Then
        If VAR_STR = "" Then
            VAR_STR = TEMP_MATRIX(i, 3)
        Else
            If VAR_STR <> TEMP_MATRIX(i, 3) Then GoTo ERROR_LABEL _
            'too many variables
        End If
    End If
Next i

Select Case OUTPUT
    Case 0
        ReDim TEMP_VECTOR(0 To NDEG, 1 To 1)
        For k = 1 To kk
            i = TEMP_MATRIX(k, 2)
            TEMP_VECTOR(i, 1) = Val(TEMP_VECTOR(i, 1) + TEMP_MATRIX(k, 1))
        Next k

        PARSE_POLYNOMIAL_STRING_FUNC = TEMP_VECTOR
    Case Else
        PARSE_POLYNOMIAL_STRING_FUNC = VAR_STR
End Select

Exit Function
ERROR_LABEL:
PARSE_POLYNOMIAL_STRING_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : PARSE_MONOMIAL_STRING_FUNC
'DESCRIPTION   : Parse a monomial
'LIBRARY       : POLYNOMIAL
'GROUP         : PARSE
'ID            : 003
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/22/2009
'************************************************************************************
'************************************************************************************

Function PARSE_MONOMIAL_STRING_FUNC(ByVal REF_MON As String, _
Optional ByVal REF_CHR As String = "§")

Dim i As Long
Dim j As Long

Dim VAR_STR As String
Dim CHR_STR As String
Dim MON_STR As String
Dim TEMP_STR As String

Dim ESP_VAL As Variant
Dim COEF_VAL As Variant

Dim ESP_FLAG As Boolean
Dim TEMP_VECTOR As Variant

On Error GoTo ERROR_LABEL

MON_STR = REF_MON & REF_CHR

ESP_VAL = 0
COEF_VAL = 0
VAR_STR = ""

ReDim TEMP_VECTOR(1 To 3, 1 To 1)

ESP_FLAG = False

j = Len(MON_STR)

For i = 1 To j
    CHR_STR = Mid(MON_STR, i, 1)
    If IS_LETTER_FUNC(CHR_STR, "_") Then
            COEF_VAL = PARSE_POLYNOMIAL_COEFFICIENTS_FUNC(TEMP_STR)
            TEMP_STR = "" 'reset
            ESP_VAL = 1
            VAR_STR = CHR_STR
    ElseIf CHR_STR = "^" Then
        ESP_FLAG = True
        TEMP_STR = ""
    ElseIf CHR_STR = REF_CHR Then
        If TEMP_STR <> "" Then
            If ESP_FLAG Then
                ESP_VAL = PARSE_POLYNOMIAL_COEFFICIENTS_FUNC(TEMP_STR)
            Else
                COEF_VAL = PARSE_POLYNOMIAL_COEFFICIENTS_FUNC(TEMP_STR)
            End If
        End If
    ElseIf CHR_STR = " " Or CHR_STR = "*" Or _
           CHR_STR = "(" Or CHR_STR = ")" Then 'nothing to do
    Else
        TEMP_STR = TEMP_STR & CHR_STR
    End If
Next i

'2X^3
TEMP_VECTOR(1, 1) = COEF_VAL '2
TEMP_VECTOR(2, 1) = ESP_VAL '3
TEMP_VECTOR(3, 1) = VAR_STR 'x

PARSE_MONOMIAL_STRING_FUNC = TEMP_VECTOR

Exit Function
ERROR_LABEL:
PARSE_MONOMIAL_STRING_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : PARSE_POLYNOMIAL_COEFFICIENTS_FUNC
'DESCRIPTION   : Convert a string coefficient into double;
'converts also fractional coeff. "3/5" -> 0.6
'LIBRARY       : POLYNOMIAL
'GROUP         : PARSE
'ID            : 004
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/22/2009
'************************************************************************************
'************************************************************************************

Function PARSE_POLYNOMIAL_COEFFICIENTS_FUNC(ByVal DATA_STR As String)

Dim k As Long
Dim TEMP1_VAL As Variant
Dim TEMP2_VAL As Variant

On Error GoTo ERROR_LABEL

If DATA_STR = "" Or DATA_STR = "+" Then
    PARSE_POLYNOMIAL_COEFFICIENTS_FUNC = 1
ElseIf DATA_STR = "-" Then
    PARSE_POLYNOMIAL_COEFFICIENTS_FUNC = -1
Else
    k = InStr(1, DATA_STR, "/")
    If k > 0 Then
        TEMP1_VAL = Left(DATA_STR, k - 1)
        TEMP2_VAL = Right(DATA_STR, Len(DATA_STR) - k)
        If TEMP2_VAL <> 0 Then TEMP1_VAL = TEMP1_VAL / TEMP2_VAL
    Else
        TEMP1_VAL = DATA_STR
    End If
    PARSE_POLYNOMIAL_COEFFICIENTS_FUNC = TEMP1_VAL
End If

Exit Function
ERROR_LABEL:
PARSE_POLYNOMIAL_COEFFICIENTS_FUNC = Err.number
End Function
