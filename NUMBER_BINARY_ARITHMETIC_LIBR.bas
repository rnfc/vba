Attribute VB_Name = "NUMBER_BINARY_ARITHMETIC_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.

'************************************************************************************
'************************************************************************************
'FUNCTION      : BINARY_VARIABLES_COUNTER_FUNC
'DESCRIPTION   : Binary Variables COUNTER
'LIBRARY       : NUMBER_BINARY
'GROUP         : ARITHMETIC
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/08/2008
'************************************************************************************
'************************************************************************************

Function BINARY_VARIABLES_COUNTER_FUNC(ByVal NSIZE As Long)

Dim i As Long
Dim j As Long
Dim k As Long

On Error GoTo ERROR_LABEL

k = 1
j = NSIZE + 1
For i = 1 To j
    If 2 ^ k >= j Then Exit For
    k = k + 1
Next i
BINARY_VARIABLES_COUNTER_FUNC = k

Exit Function
ERROR_LABEL:
BINARY_VARIABLES_COUNTER_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : BINARY_NOT_FUNC
'DESCRIPTION   : Not X --> 0 or 1
'LIBRARY       : NUMBER_BINARY
'GROUP         : ARITHMETIC
'ID            : 002
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/08/2008
'************************************************************************************
'************************************************************************************

'// PERFECT

Function BINARY_NOT_FUNC(ByVal X_BINARY_STR As String)

Dim i As Long
Dim j As Long

Dim ATEMP_STR As String
Dim BTEMP_STR As String
Dim CTEMP_STR As String

On Error GoTo ERROR_LABEL

CTEMP_STR = ""
j = Len(X_BINARY_STR)

For i = j To 1 Step -1
    ATEMP_STR = Mid(X_BINARY_STR, i, 1)
    If ATEMP_STR = "1" Then BTEMP_STR = "0" Else BTEMP_STR = "1"
    CTEMP_STR = BTEMP_STR & CTEMP_STR
Next i

BINARY_NOT_FUNC = CTEMP_STR

Exit Function
ERROR_LABEL:
BINARY_NOT_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : BINARY_AND_FUNC
'DESCRIPTION   : Returns X and Y --> 0 or 1
'LIBRARY       : NUMBER_BINARY
'GROUP         : ARITHMETIC
'ID            : 003
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/08/2008
'************************************************************************************
'************************************************************************************

'// PERFECT

Function BINARY_AND_FUNC(ByVal X_BINARY_STR As String, _
ByVal Y_BINARY_STR As String)

Dim i As Long
Dim j As Long

Dim ATEMP_STR As String
Dim BTEMP_STR As String

On Error GoTo ERROR_LABEL

BTEMP_STR = ""
i = Len(X_BINARY_STR)
j = Len(Y_BINARY_STR)
Do
    ATEMP_STR = "0"
    If Mid(X_BINARY_STR, i, 1) = "1" And _
    Mid(Y_BINARY_STR, j, 1) = "1" Then ATEMP_STR = "1"
    BTEMP_STR = ATEMP_STR & BTEMP_STR
    i = i - 1
    j = j - 1
Loop While i > 0 And j > 0

If i > 0 Then BTEMP_STR = String(i, "0") & BTEMP_STR
If j > 0 Then BTEMP_STR = String(j, "0") & BTEMP_STR

BINARY_AND_FUNC = BTEMP_STR

Exit Function
ERROR_LABEL:
BINARY_AND_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : BINARY_NOT_AND_FUNC
'DESCRIPTION   : Returns not (x and y) --> 0 or 1
'LIBRARY       : NUMBER_BINARY
'GROUP         : ARITHMETIC
'ID            : 004
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/08/2008
'************************************************************************************
'************************************************************************************

'// PERFECT

Function BINARY_NOT_AND_FUNC(ByVal X_BINARY_STR As String, _
ByVal Y_BINARY_STR As String)

Dim i As Long
Dim j As Long

Dim ATEMP_STR As String
Dim BTEMP_STR As String

On Error GoTo ERROR_LABEL

BTEMP_STR = ""
i = Len(X_BINARY_STR)
j = Len(Y_BINARY_STR)
Do
    ATEMP_STR = "1"
    If Mid(X_BINARY_STR, i, 1) = "1" And _
    Mid(Y_BINARY_STR, j, 1) = "1" Then ATEMP_STR = "0"
    BTEMP_STR = ATEMP_STR & BTEMP_STR
    i = i - 1
    j = j - 1
Loop While i > 0 And j > 0

If i > 0 Then BTEMP_STR = String(i, "1") & BTEMP_STR
If j > 0 Then BTEMP_STR = String(j, "1") & BTEMP_STR

BINARY_NOT_AND_FUNC = BTEMP_STR

Exit Function
ERROR_LABEL:
BINARY_NOT_AND_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : BINARY_OR_FUNC
'DESCRIPTION   : Returns X or Y --> 0 or 1
'LIBRARY       : NUMBER_BINARY
'GROUP         : ARITHMETIC
'ID            : 005
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/08/2008
'************************************************************************************
'************************************************************************************

'// PERFECT

Function BINARY_OR_FUNC(ByVal X_BINARY_STR As String, _
ByVal Y_BINARY_STR As String)

Dim i As Long
Dim j As Long

Dim ATEMP_STR As String
Dim BTEMP_STR As String

On Error GoTo ERROR_LABEL

BTEMP_STR = ""
i = Len(X_BINARY_STR)
j = Len(Y_BINARY_STR)
Do
    ATEMP_STR = "0"
    If Mid(X_BINARY_STR, i, 1) = "1" Or _
    Mid(Y_BINARY_STR, j, 1) = "1" Then ATEMP_STR = "1"
    BTEMP_STR = ATEMP_STR & BTEMP_STR
    i = i - 1
    j = j - 1
Loop While i > 0 And j > 0

If i > 0 Then BTEMP_STR = Left(X_BINARY_STR, i) & BTEMP_STR

If j > 0 Then BTEMP_STR = Left(Y_BINARY_STR, j) & BTEMP_STR
BINARY_OR_FUNC = BTEMP_STR

Exit Function
ERROR_LABEL:
BINARY_OR_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : BINARY_NOT_OR_FUNC
'DESCRIPTION   : Returns not (x or y) --> 0 or 1
'LIBRARY       : NUMBER_BINARY
'GROUP         : ARITHMETIC
'ID            : 006
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/08/2008
'************************************************************************************
'************************************************************************************

'// PERFECT

Function BINARY_NOT_OR_FUNC(ByVal X_BINARY_STR As String, _
ByVal Y_BINARY_STR As String)

Dim i As Long
Dim j As Long

Dim ATEMP_STR As String
Dim BTEMP_STR As String

On Error GoTo ERROR_LABEL

BTEMP_STR = ""
i = Len(X_BINARY_STR)
j = Len(Y_BINARY_STR)
Do
    ATEMP_STR = "1"
    If Mid(X_BINARY_STR, i, 1) = "1" Or _
    Mid(Y_BINARY_STR, j, 1) = "1" Then ATEMP_STR = "0"
    BTEMP_STR = ATEMP_STR & BTEMP_STR
    i = i - 1
    j = j - 1
Loop While i > 0 And j > 0

If i > 0 Then BTEMP_STR = BINARY_NOT_FUNC(Left(X_BINARY_STR, i)) & BTEMP_STR
If j > 0 Then BTEMP_STR = BINARY_NOT_FUNC(Left(Y_BINARY_STR, j)) & BTEMP_STR

BINARY_NOT_OR_FUNC = BTEMP_STR

Exit Function
ERROR_LABEL:
BINARY_NOT_OR_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : BINARY_XOR_FUNC
'DESCRIPTION   : Returns X xor Y --> 0 or 1
'LIBRARY       : NUMBER_BINARY
'GROUP         : ARITHMETIC
'ID            : 007
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/08/2008
'************************************************************************************
'************************************************************************************

'// PERFECT

Function BINARY_XOR_FUNC(ByVal X_BINARY_STR As String, _
ByVal Y_BINARY_STR As String)

Dim i As Long
Dim j As Long
Dim ATEMP_STR As String
Dim BTEMP_STR As String

On Error GoTo ERROR_LABEL

BTEMP_STR = ""
i = Len(X_BINARY_STR)
j = Len(Y_BINARY_STR)
Do
    ATEMP_STR = "1"
    If Mid(X_BINARY_STR, i, 1) = Mid(Y_BINARY_STR, j, 1) Then ATEMP_STR = "0"
    BTEMP_STR = ATEMP_STR & BTEMP_STR
    i = i - 1
    j = j - 1
Loop While i > 0 And j > 0

If i > 0 Then BTEMP_STR = Left(X_BINARY_STR, i) & BTEMP_STR
If j > 0 Then BTEMP_STR = Left(Y_BINARY_STR, j) & BTEMP_STR

BINARY_XOR_FUNC = BTEMP_STR
Exit Function
ERROR_LABEL:
BINARY_XOR_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : BINARY_NOT_XOR_FUNC
'DESCRIPTION   : Returns not (x xor y) --> 0 or 1
'LIBRARY       : NUMBER_BINARY
'GROUP         : ARITHMETIC
'ID            : 008
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/08/2008
'************************************************************************************
'************************************************************************************

'// PERFECT

Function BINARY_NOT_XOR_FUNC(ByVal X_BINARY_STR As String, _
ByVal Y_BINARY_STR As String)

Dim i As Long
Dim j As Long

Dim ATEMP_STR As String
Dim BTEMP_STR As String

On Error GoTo ERROR_LABEL

BTEMP_STR = ""
i = Len(X_BINARY_STR)
j = Len(Y_BINARY_STR)
Do
    ATEMP_STR = "0"
    If Mid(X_BINARY_STR, i, 1) = Mid(Y_BINARY_STR, j, 1) Then ATEMP_STR = "1"
    BTEMP_STR = ATEMP_STR & BTEMP_STR
    i = i - 1
    j = j - 1
Loop While i > 0 And j > 0

If i > 0 Then BTEMP_STR = BINARY_NOT_FUNC(Left(X_BINARY_STR, i)) & BTEMP_STR
If j > 0 Then BTEMP_STR = BINARY_NOT_FUNC(Left(Y_BINARY_STR, j)) & BTEMP_STR
BINARY_NOT_XOR_FUNC = BTEMP_STR

Exit Function
ERROR_LABEL:
BINARY_NOT_XOR_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : BINARY_ADD_FUNC
'DESCRIPTION   : Returns x and y
'LIBRARY       : NUMBER_BINARY
'GROUP         : ARITHMETIC
'ID            : 009
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/08/2008
'************************************************************************************
'************************************************************************************

'// PERFECT

Function BINARY_ADD_FUNC(ByVal X_BINARY_STR As String, _
ByVal Y_BINARY_STR As String)

Dim h As Long
Dim i As Long
Dim j As Long
Dim k As Long
Dim l As Long

Dim ii As Long
Dim jj As Long

Dim TEMP_STR As String

On Error GoTo ERROR_LABEL
TEMP_STR = ""
i = Len(X_BINARY_STR)
j = Len(Y_BINARY_STR)
h = 0
Do
    If i > 0 Then ii = CInt(Mid(X_BINARY_STR, i, 1)) Else ii = 0
    If j > 0 Then jj = CInt(Mid(Y_BINARY_STR, j, 1)) Else jj = 0
    k = ii + jj + h
    l = Int(k / 2)
    TEMP_STR = CStr(k - 2 * l) & TEMP_STR
    i = i - 1
    j = j - 1
    h = l
Loop While i > 0 Or j > 0

If l = 1 Then TEMP_STR = "1" & TEMP_STR
BINARY_ADD_FUNC = TEMP_STR

Exit Function
ERROR_LABEL:
BINARY_ADD_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : BINARY_DIFFERENCE_FUNC
'DESCRIPTION   : Difference between two binaries
'LIBRARY       : NUMBER_BINARY
'GROUP         : ARITHMETIC
'ID            : 010
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/08/2008
'************************************************************************************
'************************************************************************************

'// PERFECT

Function BINARY_DIFFERENCE_FUNC(ByVal X_BINARY_STR As String, _
ByVal Y_BINARY_STR As String)

Dim h As Long
Dim i As Long
Dim j As Long
Dim k As Long
Dim l As Long

Dim ii As Long
Dim jj As Long

Dim TEMP_STR As String

On Error GoTo ERROR_LABEL
TEMP_STR = ""
i = Len(X_BINARY_STR)
j = Len(Y_BINARY_STR)
h = 0
Do
    If i > 0 Then ii = CInt(Mid(X_BINARY_STR, i, 1)) Else ii = 0
    If j > 0 Then jj = CInt(Mid(Y_BINARY_STR, j, 1)) Else jj = 0
    k = ii - jj - h
    l = -Int(k / 2)
    TEMP_STR = CStr(k - 2 * Int(k / 2)) & TEMP_STR
    i = i - 1
    j = j - 1
    h = l
Loop While i > 0 Or j > 0
If Left(TEMP_STR, 1) = "0" Then TEMP_STR = Right(TEMP_STR, Len(TEMP_STR) - 1)
BINARY_DIFFERENCE_FUNC = TEMP_STR

Exit Function
ERROR_LABEL:
BINARY_DIFFERENCE_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : BINARY_INCREMENT_FUNC
'DESCRIPTION   : Returns x + 1
'LIBRARY       : NUMBER_BINARY
'GROUP         : ARITHMETIC
'ID            : 011
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/08/2008
'************************************************************************************
'************************************************************************************

'// PERFECT

Function BINARY_INCREMENT_FUNC(ByVal X_BINARY_STR As String)

Dim h As Long
Dim i As Long
Dim j As Long
Dim TEMP_STR As String

On Error GoTo ERROR_LABEL

TEMP_STR = X_BINARY_STR
h = Len(TEMP_STR)
j = 0

For i = h To 1 Step -1
    If Mid(TEMP_STR, i, 1) = "1" Then
        Mid(TEMP_STR, i, 1) = "0"
        j = 1
    Else
        Mid(TEMP_STR, i, 1) = "1"
        j = 0
        Exit For
    End If
Next i

If j = 1 Then TEMP_STR = "1" & TEMP_STR
BINARY_INCREMENT_FUNC = TEMP_STR

Exit Function
ERROR_LABEL:
BINARY_INCREMENT_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : BINARY_COMPLEMENT_FUNC
'DESCRIPTION   : Returns the complement of x
'LIBRARY       : NUMBER_BINARY
'GROUP         : ARITHMETIC
'ID            : 012
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/08/2008
'************************************************************************************
'************************************************************************************

'// PERFECT
Function BINARY_COMPLEMENT_FUNC(ByVal X_BINARY_STR As String)

Dim i As Long
Dim j As Long

Dim TEMP_STR As String

On Error GoTo ERROR_LABEL

If X_BINARY_STR = "0" Then
    BINARY_COMPLEMENT_FUNC = "10"
    Exit Function
End If

TEMP_STR = "0" & X_BINARY_STR
j = Len(TEMP_STR)
For i = j To 1 Step -1
    If Mid(TEMP_STR, i, 1) = "1" Then
        Mid(TEMP_STR, i, 1) = "0"
    Else
        Mid(TEMP_STR, i, 1) = "1"
    End If
Next i

BINARY_COMPLEMENT_FUNC = BINARY_INCREMENT_FUNC(TEMP_STR)

Exit Function
ERROR_LABEL:
BINARY_COMPLEMENT_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : BINARY_MATRIX_MULT_FUNC
'DESCRIPTION   : Returns the binary multiplication of two binary matrix
'LIBRARY       : NUMBER_BINARY
'GROUP         : ARITHMETIC
'ID            : 013
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/08/2008
'************************************************************************************
'************************************************************************************


Function BINARY_MATRIX_MULT_FUNC(ByRef ADATA_RNG As Variant, _
ByRef BDATA_RNG As Variant)

Dim i As Long
Dim j As Long
Dim k As Long

Dim NSIZE As Long
Dim NROWS As Long
Dim NCOLUMNS As Long

Dim TEMP_SUM As Double

Dim ADATA_MATRIX As Variant
Dim BDATA_MATRIX As Variant

Dim TEMP_MATRIX As Variant

On Error GoTo ERROR_LABEL

ADATA_MATRIX = ADATA_RNG
BDATA_MATRIX = BDATA_RNG

NROWS = UBound(ADATA_MATRIX, 1)
NCOLUMNS = UBound(ADATA_MATRIX, 2)
NSIZE = UBound(BDATA_MATRIX, 1)

ReDim TEMP_MATRIX(1 To NSIZE, 1 To NROWS)

For i = 1 To NSIZE
    For j = 1 To NROWS
        TEMP_SUM = 0
        For k = 1 To NCOLUMNS
            If ADATA_MATRIX(j, k) <> 0 Then
                If ADATA_MATRIX(j, k) = 1 Then
                    TEMP_SUM = TEMP_SUM + BDATA_MATRIX(i, k)
                Else
                    If BDATA_MATRIX(i, k) = 0 Then TEMP_SUM = TEMP_SUM + 1
                End If
            End If
        Next k
        TEMP_MATRIX(i, j) = TEMP_SUM Mod 2
    Next j
Next i

BINARY_MATRIX_MULT_FUNC = TEMP_MATRIX

Exit Function
ERROR_LABEL:
BINARY_MATRIX_MULT_FUNC = Err.number
End Function
