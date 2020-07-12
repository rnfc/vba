Attribute VB_Name = "NUMBER_COMPLEX_PARSE_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.


'************************************************************************************
'************************************************************************************
'FUNCTION      : COMPLEX_CHARACTER_STRING_FUNC
'DESCRIPTION   : Get the Complex Character String
'LIBRARY       : NUMBER_COMPLEX
'GROUP         : PARSE
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/08/2008
'************************************************************************************
'************************************************************************************


Function COMPLEX_CHARACTER_STRING_FUNC(ByVal DATA_STR As String, _
Optional ByVal CPLX_CHR_STR As String = "i")

On Error GoTo ERROR_LABEL

If InStr(1, DATA_STR, "i", vbTextCompare) > 0 Then
    COMPLEX_CHARACTER_STRING_FUNC = "i"
ElseIf InStr(1, DATA_STR, "j", vbTextCompare) > 0 Then
    COMPLEX_CHARACTER_STRING_FUNC = "j"
ElseIf InStr(1, DATA_STR, CPLX_CHR_STR, vbTextCompare) > 0 Then
    COMPLEX_CHARACTER_STRING_FUNC = CPLX_CHR_STR
Else
    GoTo ERROR_LABEL
End If

Exit Function
ERROR_LABEL:
COMPLEX_CHARACTER_STRING_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : COMPLEX_EXTRACT_NUMBER_FUNC
'DESCRIPTION   : Extract a complex number from cells, array or string
'LIBRARY       : NUMBER_COMPLEX
'GROUP         : PARSE
'ID            : 002
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/08/2008
'************************************************************************************
'************************************************************************************

Function COMPLEX_EXTRACT_NUMBER_FUNC(ByVal DATA_RNG As Variant, _
Optional ByVal CPLX_CHR_STR As String = "i")

Dim TEMP_VECTOR As Variant
Dim DATA_VECTOR As Variant

On Error GoTo ERROR_LABEL

'-----------------------------------------------------------------------------
If IsNumeric(DATA_RNG) = True Then
'-----------------------------------------------------------------------------
    ReDim TEMP_VECTOR(1 To 2, 1 To 1)
    TEMP_VECTOR(1, 1) = DATA_RNG
    TEMP_VECTOR(2, 1) = 0
'-----------------------------------------------------------------------------
ElseIf IsArray(DATA_RNG) = True Then
'-----------------------------------------------------------------------------
    If IS_2D_ARRAY_FUNC(DATA_RNG) = True Then
        DATA_VECTOR = DATA_RNG
        If UBound(DATA_VECTOR, 1) = 1 Then
            DATA_VECTOR = MATRIX_TRANSPOSE_FUNC(DATA_VECTOR)
        End If
        ReDim TEMP_VECTOR(1 To 2, 1 To 1)
        TEMP_VECTOR(1, 1) = DATA_VECTOR(1, 1)
        TEMP_VECTOR(2, 1) = DATA_VECTOR(2, 1)
    Else
        ReDim TEMP_VECTOR(1 To 2, 1 To 1)
        TEMP_VECTOR(1, 1) = DATA_VECTOR(1)
        TEMP_VECTOR(2, 1) = DATA_VECTOR(2)
    End If
'-----------------------------------------------------------------------------
Else 'If It is String
'-----------------------------------------------------------------------------
    TEMP_VECTOR = COMPLEX_CONVERT_STRING_FUNC(DATA_RNG, CPLX_CHR_STR)
'-----------------------------------------------------------------------------
End If
'-----------------------------------------------------------------------------

COMPLEX_EXTRACT_NUMBER_FUNC = TEMP_VECTOR

Exit Function
ERROR_LABEL:
COMPLEX_EXTRACT_NUMBER_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : COMPLEX_CONVERT_COEFFICIENTS_FUNC
'DESCRIPTION   : Transform a complex number (x,y) into a complex string x+iy

'LIBRARY       : NUMBER_COMPLEX
'GROUP         : PARSE
'ID            : 003
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/08/2008
'************************************************************************************
'************************************************************************************
Function COMPLEX_CONVERT_COEFFICIENTS_FUNC(ByVal REAL_NUMBER As Double, _
ByVal IMAG_NUMBER As Double, _
Optional ByVal CPLX_CHR_STR As String = "i", _
Optional ByVal epsilon As Double = 5 * 10 ^ -15)

'REAL_NUMBER: is the real coefficient of the complex number.
'IMAG_NUMBER: is the imaginary coefficient of the complex number.
'Suffix: is the suffix for the imaginary component of the
'complex number.

Dim k As Long
Dim XTEMP_STR As String
Dim YTEMP_STR As String
Dim VTEMP_STR As String

On Error GoTo ERROR_LABEL

k = 12 'significant digits
'epsilon = 10 ^ -12

If Abs(REAL_NUMBER) < epsilon Then REAL_NUMBER = 0
If Abs(IMAG_NUMBER) < epsilon Then IMAG_NUMBER = 0

If Abs(IMAG_NUMBER - 1) < epsilon Then IMAG_NUMBER = 1
If Abs(IMAG_NUMBER + 1) < epsilon Then IMAG_NUMBER = -1

'round-off limit
REAL_NUMBER = COMPLEX_NUMBER_ROUND_FUNC(REAL_NUMBER, k) '
IMAG_NUMBER = COMPLEX_NUMBER_ROUND_FUNC(IMAG_NUMBER, k) '

If REAL_NUMBER <> 0 Then: XTEMP_STR = COMPLEX_NUMBER_STRING_FUNC(REAL_NUMBER)

If IMAG_NUMBER = 1 Then
    YTEMP_STR = CPLX_CHR_STR
ElseIf IMAG_NUMBER = -1 Then
    YTEMP_STR = "-" & CPLX_CHR_STR
ElseIf IMAG_NUMBER <> 0 Then
    YTEMP_STR = COMPLEX_NUMBER_STRING_FUNC(IMAG_NUMBER) & CPLX_CHR_STR
End If
If IMAG_NUMBER > 0 And XTEMP_STR <> "" Then YTEMP_STR = "+" & YTEMP_STR
VTEMP_STR = XTEMP_STR & YTEMP_STR
If VTEMP_STR = "" Then VTEMP_STR = "0"

COMPLEX_CONVERT_COEFFICIENTS_FUNC = VTEMP_STR

Exit Function
ERROR_LABEL:
COMPLEX_CONVERT_COEFFICIENTS_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : COMPLEX_CONVERT_STRING_FUNC
'DESCRIPTION   : Transform a complex string  x+iy into a complex number (x,y)
'LIBRARY       : NUMBER_COMPLEX
'GROUP         : PARSE
'ID            : 004
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/08/2008
'************************************************************************************
'************************************************************************************

Function COMPLEX_CONVERT_STRING_FUNC(ByVal DATA_STR As String, _
Optional ByVal CPLX_CHR_STR As String = "i")

Dim i As Long
Dim j As Long

Dim CHR_STR As String
Dim REAL_STR As String
Dim IMAG_STR As String
Dim DECIMAL_STR As String
Dim TEMP_VECTOR As Variant

On Error GoTo ERROR_LABEL

DECIMAL_STR = Mid(CStr(1 / 2), 2, 1)
ReDim TEMP_VECTOR(1 To 2, 1 To 1)

DATA_STR = Trim(DATA_STR)

If DATA_STR = "" Or DATA_STR = "0" Then GoTo 1983
If DATA_STR = "i" Or DATA_STR = "j" Then TEMP_VECTOR(2, 1) = 1: GoTo 1983
If DATA_STR = "-i" Or DATA_STR = "-j" Then TEMP_VECTOR(2, 1) = -1: GoTo 1983
If DATA_STR = CPLX_CHR_STR Then TEMP_VECTOR(2, 1) = 1: GoTo 1983
If DATA_STR = "-" & CPLX_CHR_STR Then TEMP_VECTOR(2, 1) = -1: GoTo 1983

If InStr(1, DATA_STR, CPLX_CHR_STR) = 0 And InStr(1, DATA_STR, "i") = 0 _
And InStr(1, DATA_STR, "j") = 0 Then
    TEMP_VECTOR(1, 1) = COMPLEX_STRING_NUMBER_FUNC(DATA_STR, DECIMAL_STR)
    'pure number
    GoTo 1983
End If

i = 2 'parse begins
j = Len(DATA_STR)
Do
    CHR_STR = Mid(DATA_STR, i, 1)
    If CHR_STR = "+" Or CHR_STR = "-" Then
        If UCase(Mid(DATA_STR, i - 1, 1)) <> "E" Then Exit Do
    End If
    i = i + 1
Loop While i <= j

If i <= j Then
    REAL_STR = Trim(Left(DATA_STR, i - 1))
    IMAG_STR = Trim(Mid(DATA_STR, i, j - i))
Else
    REAL_STR = "0"
    IMAG_STR = Trim(Left(DATA_STR, j - 1))
End If

If IMAG_STR = "+" Then IMAG_STR = "1"
If IMAG_STR = "-" Then IMAG_STR = "-1"

TEMP_VECTOR(1, 1) = COMPLEX_STRING_NUMBER_FUNC(REAL_STR, DECIMAL_STR)
TEMP_VECTOR(2, 1) = COMPLEX_STRING_NUMBER_FUNC(IMAG_STR, DECIMAL_STR)

1983:

COMPLEX_CONVERT_STRING_FUNC = TEMP_VECTOR

Exit Function
ERROR_LABEL:
COMPLEX_CONVERT_STRING_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : COMPLEX_MATRIX_FORMAT_FUNC

'DESCRIPTION   : This function has 3 different complex matrix formats:

'----------------------------------------------------------------
' SPLIT FORMAT:
'----------------------------------------------------------------
' 1   2   0  |  0  -1   3
'-1   3  -1  |  0   2  -1
' 0  -1   4  |  0  -2   0

'----------------------------------------------------------------
' INTERLACED FORMAT:
'----------------------------------------------------------------
' 1    0  ||  2   -1  ||  0    3  ||
'-1    0  ||  3    2  || -1   -1  ||
' 0    0  || -1   -2  ||  4    0  ||
'----------------------------------------------------------------
'STRING FORMAT
'----------------------------------------------------------------
' 1    2-i     3i
'-1   3+2i   -1-i
' 0  -1-2i    4

'As we can see, in the Splie format the complex matrix [ Z ] is split
'into two separate matrices: the first contains the real values, and
'the second one the imaginary values. This is the default format.

'In the Interlaced format, the complex values are written as two adjacent
'cells, so that each a individual matrix element occupies two adjacent
'cells. The number of columns is the same as in the first format, but
'the values are interlaced, so that each real column is followed by an
'imaginary column and so on. This format is useful when the elements
'are returned by complex functions. The String format is the well known
'“complex rectangular format”. Each element is written as a text string
'a+ib; therefore the square matrix is still square. This is the most
'compact and intuitive format for integer values. For non-integer values
'the matrix may become illegible. We must also point out that these
'elements, being strings, cannot be formatted with the standard tools
'of Excel, but must be converted back to numbers with the Excel commands

'IMREAL and IMAGINARY.
'LIBRARY       : NUMBER_COMPLEX
'GROUP         : PARSE
'ID            : 005
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/08/2008
'************************************************************************************
'************************************************************************************


Function COMPLEX_MATRIX_FORMAT_FUNC(ByRef DATA_RNG As Variant, _
Optional ByVal CONV_OPT As Integer = 0, _
Optional ByVal CPLX_CHR_STR As String = "i", _
Optional ByVal epsilon As Double = 5 * 10 ^ -15)

'CONV_OPT: 0, 12 --> from 1 to 2
'CONV_OPT: 1, 13 --> from 1 to 3
'CONV_OPT: 2, 21 --> from 2 to 1
'CONV_OPT: 3, 31 --> from 3 to 1

'1  is DATA_MATRIX complex matrix
'      (NROWS x 2m) DATA_MATRIX = [Ar],[Ai] (2 matrices)
'2  is DATA_MATRIX complex matrix
'      (NROWS x 2m) DATA_MATRIX = [Ar,Ai]   (2 cells)
'3  is DATA_MATRIX complex matrix
'      (NROWS x NCOLUMNS)  DATA_MATRIX = [Ar+jAi]   (string)

Dim h As Long
Dim i As Long
Dim j As Long
Dim k As Long

Dim NROWS As Long
Dim NCOLUMNS As Long

Dim ADATA_MATRIX As Variant
Dim BDATA_MATRIX As Variant

Dim TEMP_VECTOR As Variant

On Error GoTo ERROR_LABEL

ADATA_MATRIX = DATA_RNG
BDATA_MATRIX = ADATA_MATRIX

NROWS = UBound(BDATA_MATRIX, 1)
NCOLUMNS = UBound(BDATA_MATRIX, 2)

'-------------------------------------------------------------------------------------
Select Case CONV_OPT
'-------------------------------------------------------------------------------------
    Case 0, 12 'Perfect
'-------------------------------------------------------------------------------------
        If NCOLUMNS Mod 2 <> 0 Then GoTo ERROR_LABEL
        ReDim ADATA_MATRIX(1 To NROWS, 1 To NCOLUMNS)
        k = NCOLUMNS / 2
        For j = 1 To k
            h = 2 * j - 1
            For i = 1 To NROWS
                ADATA_MATRIX(i, h) = BDATA_MATRIX(i, j)
                ADATA_MATRIX(i, h + 1) = BDATA_MATRIX(i, j + k)
            Next i
        Next j
'-------------------------------------------------------------------------------------
    Case 1, 13 'Perfect
'-------------------------------------------------------------------------------------
        If NCOLUMNS Mod 2 <> 0 Then GoTo ERROR_LABEL
        k = NCOLUMNS / 2
        ReDim ADATA_MATRIX(1 To NROWS, 1 To k)
        For j = 1 To k
            For i = 1 To NROWS
                ADATA_MATRIX(i, j) = COMPLEX_CONVERT_COEFFICIENTS_FUNC(BDATA_MATRIX(i, j), _
                BDATA_MATRIX(i, j + k), CPLX_CHR_STR, epsilon)
            Next i
        Next j
'-------------------------------------------------------------------------------------
    Case 2, 21 'Perfect
'-------------------------------------------------------------------------------------
        If NCOLUMNS Mod 2 <> 0 Then GoTo ERROR_LABEL
        ReDim ADATA_MATRIX(1 To NROWS, 1 To NCOLUMNS)
        k = NCOLUMNS / 2
        For j = 1 To k
            h = 2 * j - 1
            For i = 1 To NROWS
                ADATA_MATRIX(i, j) = BDATA_MATRIX(i, h)
                ADATA_MATRIX(i, j + k) = BDATA_MATRIX(i, h + 1)
            Next i
        Next j
'-------------------------------------------------------------------------------------
    Case 3, 31 'Perfect
'-------------------------------------------------------------------------------------
        ReDim ADATA_MATRIX(1 To NROWS, 1 To 2 * NCOLUMNS)
        For j = 1 To NCOLUMNS
            For i = 1 To NROWS
                TEMP_VECTOR = COMPLEX_CONVERT_STRING_FUNC(BDATA_MATRIX(i, j), CPLX_CHR_STR)
                ADATA_MATRIX(i, j) = TEMP_VECTOR(1, 1)
                ADATA_MATRIX(i, j + NCOLUMNS) = TEMP_VECTOR(2, 1)
            Next i
        Next j
'-------------------------------------------------------------------------------------
    Case Else
'-------------------------------------------------------------------------------------
        GoTo ERROR_LABEL
'-------------------------------------------------------------------------------------
End Select
'-------------------------------------------------------------------------------------

COMPLEX_MATRIX_FORMAT_FUNC = ADATA_MATRIX

Exit Function
ERROR_LABEL:
COMPLEX_MATRIX_FORMAT_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : COMPLEX_STRING_NUMBER_FUNC
'DESCRIPTION   : Convert to Double Number (32-bit precision)
'GROUP         : NUMBER_COMPLEX
'GROUP         : PARSE
'ID            : 006
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/08/2008
'************************************************************************************
'************************************************************************************

Private Function COMPLEX_STRING_NUMBER_FUNC(ByVal DATA_VAL As Variant, _
Optional ByVal DECIM_SEPAR_CHR As String = ".")

Dim k As Long
Dim TEMP_STR As Variant
On Error GoTo ERROR_LABEL
TEMP_STR = DATA_VAL
If VarType(TEMP_STR) = vbString Then
    If DECIM_SEPAR_CHR = "" Then DECIM_SEPAR_CHR = Mid(CStr(1 / 2), 2, 1)
    If DECIM_SEPAR_CHR <> "." Then
      k = InStr(1, TEMP_STR, DECIM_SEPAR_CHR)
      If k > 0 Then Mid(TEMP_STR, k, 1) = "."
    End If
      COMPLEX_STRING_NUMBER_FUNC = Val(TEMP_STR)
Else: COMPLEX_STRING_NUMBER_FUNC = TEMP_STR
End If

Exit Function
ERROR_LABEL:
COMPLEX_STRING_NUMBER_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : COMPLEX_NUMBER_STRING_FUNC
'DESCRIPTION   : Convert a number into string following the international
'setting (0.05 ms)
'GROUP         : NUMBER_COMPLEX
'GROUP         : PARSE
'ID            : 007
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/02/2008
'************************************************************************************
'************************************************************************************

Private Function COMPLEX_NUMBER_STRING_FUNC(ByVal DATA_VAL As Variant)

Dim k As Long
Dim TEMP_STR As String
Dim DECIM_SEPAR_CHR As String

On Error GoTo ERROR_LABEL

DECIM_SEPAR_CHR = Mid(CStr(1 / 2), 2, 1)
'Excel.Application.International(xlDecimalSeparator)
TEMP_STR = CStr(DATA_VAL)
k = InStr(1, TEMP_STR, ".")
If k > 0 Then
    If "." <> DECIM_SEPAR_CHR Then Mid(TEMP_STR, k, 1) = DECIM_SEPAR_CHR
Else
    k = InStr(1, TEMP_STR, ",")
    If k > 0 Then
        If "," <> DECIM_SEPAR_CHR Then Mid(TEMP_STR, k, 1) = DECIM_SEPAR_CHR
    End If
End If
COMPLEX_NUMBER_STRING_FUNC = TEMP_STR

Exit Function
ERROR_LABEL:
COMPLEX_NUMBER_STRING_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : COMPLEX_NUMBER_STRING_FUNC
'DESCRIPTION   : Round real value
'GROUP         : NUMBER_COMPLEX
'GROUP         : PARSE
'ID            : 008
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/02/2008
'************************************************************************************
'************************************************************************************

Function COMPLEX_NUMBER_ROUND_FUNC(ByVal DATA_VAL As Double, _
Optional ByVal DECIMALS As Integer = 2)

Dim k As Long
Dim Y_VAL As Double
Dim X_VAL As Double

Dim epsilon As Double

On Error GoTo ERROR_LABEL

epsilon = 2 * 10 ^ -16

X_VAL = DATA_VAL
If Abs(X_VAL) <= epsilon Then 'If x = 0 Then
    COMPLEX_NUMBER_ROUND_FUNC = 0
    Exit Function
Else
    k = Int(Log(Abs(X_VAL)) / Log(10#)) + 1
    Y_VAL = X_VAL / 10 ^ k
    Y_VAL = Round(Y_VAL, DECIMALS)
End If

COMPLEX_NUMBER_ROUND_FUNC = Y_VAL * 10 ^ k

Exit Function
ERROR_LABEL:
COMPLEX_NUMBER_ROUND_FUNC = Err.number
End Function

