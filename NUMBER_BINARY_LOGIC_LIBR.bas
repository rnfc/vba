Attribute VB_Name = "NUMBER_BINARY_LOGIC_LIBR"


Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.

'************************************************************************************
'************************************************************************************
'FUNCTION      : BINARY_LOGIC_FUNC

'DESCRIPTION   : This function returns the logic function of the
'following two forms AND/OR. Where the symbol “dot” is the AND
'operator and “cross” is the OR operator. Each variable can be
'complemented or not. Each (….)  parenthesis is called “implicant”
'in the Logic Network Theory. Each implicant can contain all or part
'of the independet variables.

'LIBRARY       : NUMBER_BINARY
'GROUP         : LOGIC
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 14/02/2008

'************************************************************************************
'************************************************************************************

Function BINARY_LOGIC_FUNC(ByRef IMPLICANT_RNG As Variant, _
ByRef BINARY_RNG As Variant, _
Optional ByVal VERSION As Integer = 0)

Dim i As Long
Dim j As Long
Dim k As Long
Dim h As Long
Dim l As Long

Dim ii As Long
Dim jj As Long
Dim kk As Long

Dim NROWS As Long
Dim NCOLUMNS As Long

Dim TEMP_VECTOR As Variant
Dim BINARY_VECTOR As Variant
Dim IMPLICANT_MATRIX As Variant

On Error GoTo ERROR_LABEL

IMPLICANT_MATRIX = IMPLICANT_RNG
NROWS = UBound(IMPLICANT_MATRIX, 1)
NCOLUMNS = UBound(IMPLICANT_MATRIX, 2)

BINARY_VECTOR = BINARY_RNG
If UBound(BINARY_VECTOR, 1) = 1 Then
    BINARY_VECTOR = MATRIX_TRANSPOSE_FUNC(BINARY_VECTOR)
End If
l = UBound(BINARY_VECTOR, 1)

If VERSION = 0 Then
    ii = 0
    jj = 1    'logic AND
Else
    ii = 1
    jj = 0    'logic OR
End If

ReDim TEMP_VECTOR(1 To l, 1 To 1)

For k = 1 To l
    kk = ii
    For i = 1 To NROWS
        h = jj
        For j = 1 To NCOLUMNS
            If IMPLICANT_MATRIX(i, j) <> 0 Then
                If IMPLICANT_MATRIX(i, j) = 1 And _
                    BINARY_VECTOR(k, j) = ii Or _
                    IMPLICANT_MATRIX(i, j) = -1 And _
                        BINARY_VECTOR(k, j) = jj Then
                    h = ii
                    Exit For
                End If
            End If
        Next j
        If h = jj Then kk = jj: Exit For
    Next i
    TEMP_VECTOR(k, 1) = kk
Next k

BINARY_LOGIC_FUNC = TEMP_VECTOR

Exit Function
ERROR_LABEL:
BINARY_LOGIC_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : BINARY_LOGIC_SYNTHESIS_FUNC

'DESCRIPTION   : This function is usefull to synthesize a logic function from
'its input/output table. We can say that it is the inverse process of the mapping.
'The function, while using the Quine-McKluskey algorithm, returns the logic
'function of the AND logic form.

'Where the symbol “dot” is the AND operator and “cross” is the OR operator.
'Each variable can be complemented or not. Each (….)  parenthesis is called
'“implicant” in the Logic Network Theory. Each implicant can contain all or
'part of the independet variables. The convention for writing an implicant
'is a row-array.

'LIBRARY       : NUMBER_BINARY
'GROUP         : LOGIC
'ID            : 002
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 14/02/2008
'************************************************************************************
'************************************************************************************


Function BINARY_LOGIC_SYNTHESIS_FUNC(ByRef DATA_RNG As Variant, _
Optional ByVal LOGIC_FLAG As Boolean = False, _
Optional ByVal CHR_STR As String = "x", _
Optional ByVal SYMBOL_STR As String = "*", _
Optional ByVal OUTPUT As Integer = 1)

Dim h As Long
Dim i As Long
Dim j As Long
Dim k As Long
Dim l As Long

Dim NSIZE As Long
Dim NROWS As Long
Dim NCOLUMNS As Long

Dim TEMP_STR As String

Dim ATEMP_VECTOR As Variant
Dim BTEMP_VECTOR As Variant
Dim CTEMP_VECTOR As Variant

Dim TEMP_MATRIX As Variant
Dim DATA_MATRIX As Variant

On Error GoTo ERROR_LABEL

DATA_MATRIX = DATA_RNG
NROWS = UBound(DATA_MATRIX, 1)
NCOLUMNS = UBound(DATA_MATRIX, 2)

If LOGIC_FLAG Then ' converts logic True/False --> 1/0
    For i = 1 To NROWS
        For j = 1 To NCOLUMNS
            If DATA_MATRIX(i, j) = True Then DATA_MATRIX(i, j) = 1
            If DATA_MATRIX(i, j) = False Then DATA_MATRIX(i, j) = 0
        Next j
    Next i
End If

ReDim CTEMP_VECTOR(1 To NROWS, 1 To 2)
'---------------------------------------------------------------------------
If NCOLUMNS = 2 Then 'decimal format input
'---------------------------------------------------------------------------
    If VarType(DATA_MATRIX(1, 1)) = vbString Then
        For i = 1 To NROWS
            CTEMP_VECTOR(i, 1) = CONVERT_BINARY_DECIMAL_FUNC(DATA_MATRIX(i, 1))
            CTEMP_VECTOR(i, 2) = DATA_MATRIX(i, 2)
        Next i
    Else
        For i = 1 To NROWS
            CTEMP_VECTOR(i, 1) = DATA_MATRIX(i, 1)
            CTEMP_VECTOR(i, 2) = DATA_MATRIX(i, 2)
        Next i
    End If
'---------------------------------------------------------------------------
Else
'---------------------------------------------------------------------------
    'vector-binary format input
    For i = 1 To NROWS
        TEMP_STR = ""
        For j = 1 To NCOLUMNS - 1
            TEMP_STR = TEMP_STR & CStr(DATA_MATRIX(i, j))
        Next j
        CTEMP_VECTOR(i, 1) = CONVERT_BINARY_DECIMAL_FUNC(TEMP_STR)
        CTEMP_VECTOR(i, 2) = DATA_MATRIX(i, NCOLUMNS)
    Next i
'---------------------------------------------------------------------------
End If
'---------------------------------------------------------------------------
    
CTEMP_VECTOR = BINARY_SORT_FUNC(CTEMP_VECTOR, 1, 1)
k = 1
NCOLUMNS = CTEMP_VECTOR(NROWS, 1) + 1
For i = 1 To NCOLUMNS
    If 2 ^ k >= NCOLUMNS Then Exit For
    k = k + 1
Next i

NSIZE = 2 ^ k
ReDim TEMP_MATRIX(1 To NSIZE, 1 To 1)
For i = 1 To NSIZE
    TEMP_MATRIX(i, 1) = CHR_STR
Next i
For i = 1 To NROWS
    TEMP_MATRIX(i, 1) = CTEMP_VECTOR(i, 2)
Next i
    
ATEMP_VECTOR = BINARY_QUINE_MCCLUSKEY_FUNC(TEMP_MATRIX, _
               CHR_STR, SYMBOL_STR, 0)

BTEMP_VECTOR = BINARY_QUINE_MCCLUSKEY_FUNC(TEMP_MATRIX, _
               CHR_STR, SYMBOL_STR, 1)

'-------------------------------------------------------------------------------
Select Case OUTPUT
'-------------------------------------------------------------------------------
'The Implicants table contains all basic implicants found by the Quine-McKluskey
'algorithm. As known, the implicants of the canonical form AND are a subset of
'the basic implicants table.

'In this table we see another compact form for writing an implicant: “1” means ,
'“0” means  and  “*” means no present.

'The covering matrix is shown at the right. Each row indicates an input binary
'configuration where  y = 1. In our case (5, 6, 7,8, 9, 12, 13, 14). Each column
'indicates a basic implicants: column 1 indicates the implicant 1 of the
'Implicant table, column 2 the implicant 2, and so on.
'The meaning of the cover matrix is show wich implicant “cover” a “1” of the
'function.

'Sometime the logic tables are not defined for every input configuration.
'In that case the algo assumes that the missing configurations can have
'indifferently 1 or 0 and automatically add the missing configuration
'with “X” value or “*” (indifference value).
'-------------------------------------------------------------------------------
Case 0 'Minimum AND form
'-------------------------------------------------------------------------------
        
        NSIZE = UBound(BTEMP_VECTOR, 1)
        l = 0
        For i = 1 To NSIZE
            If BTEMP_VECTOR(i, 2) <> 0 Then: l = l + 1
        Next i
        ReDim CTEMP_VECTOR(1 To l, 1 To Len(TEMP_STR))
        k = Len(BTEMP_VECTOR(1, 1))
        l = 1
        For i = 1 To NSIZE
            If BTEMP_VECTOR(i, 2) <> 0 Then
                TEMP_STR = BTEMP_VECTOR(i, 1)
                For j = 1 To Len(TEMP_STR)
                    If Mid(TEMP_STR, j, 1) = "1" Then
                        h = 1
                    ElseIf Mid(TEMP_STR, j, 1) = "0" Then
                        h = -1
                    Else
                        h = 0
                    End If
                    CTEMP_VECTOR(l, j) = h
                Next j
                l = l + 1
            End If
        Next i

'-------------------------------------------------------------------------------
Case 1 'Implicants
'-------------------------------------------------------------------------------
        NSIZE = UBound(BTEMP_VECTOR, 1)
        ReDim CTEMP_VECTOR(1 To NSIZE, 1 To 2)
        For i = 1 To NSIZE
            CTEMP_VECTOR(i, 1) = i
            CTEMP_VECTOR(i, 2) = BTEMP_VECTOR(i, 1)
        Next i
'-------------------------------------------------------------------------------
Case Else 'Covering Matrix
'-------------------------------------------------------------------------------
        
        NROWS = UBound(ATEMP_VECTOR, 1)
        NCOLUMNS = UBound(ATEMP_VECTOR, 2)
        
        ReDim CTEMP_VECTOR(1 To NROWS + 1, 1 To NCOLUMNS + 1)
        For j = 1 To NCOLUMNS
            CTEMP_VECTOR(1, j + 1) = j
        Next j
        
        For j = 1 To NCOLUMNS
            For i = 1 To NROWS
                CTEMP_VECTOR(1 + i, j + 1) = ATEMP_VECTOR(i, j)
            Next i
        Next j
        
        l = 0
        For i = 1 To UBound(TEMP_MATRIX, 1)
            If TEMP_MATRIX(i, 1) = 1 Then
                l = l + 1
                CTEMP_VECTOR(1 + l, 1) = i - 1
            End If
        Next i
'-------------------------------------------------------------------------------
End Select
'-------------------------------------------------------------------------------

For j = LBound(CTEMP_VECTOR, 2) To UBound(CTEMP_VECTOR, 2)
    For i = LBound(CTEMP_VECTOR, 1) To UBound(CTEMP_VECTOR, 1)
        If IsEmpty(CTEMP_VECTOR(i, j)) = True Then: CTEMP_VECTOR(i, j) = ""
    Next i
Next j
    
BINARY_LOGIC_SYNTHESIS_FUNC = CTEMP_VECTOR

Exit Function
ERROR_LABEL:
BINARY_LOGIC_SYNTHESIS_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : BINARY_QUINE_MCCLUSKEY_FUNC
'DESCRIPTION   : 'The Quine–McCluskey algorithm is a method used for minimization
'of boolean functions. It is functionally identical to Karnaugh mapping, but the
'tabular form makes it more efficient for use in computer algorithms, and it also
'gives a deterministic way to check that the minimal form of a Boolean function has
'been reached. It is sometimes referred to as the tabulation method.

'LIBRARY       : NUMBER_BINARY
'GROUP         : LOGIC
'ID            : 003
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 14/02/2008
'************************************************************************************
'************************************************************************************


Function BINARY_QUINE_MCCLUSKEY_FUNC(ByRef DATA_RNG As Variant, _
Optional ByVal CHR_STR As String = "x", _
Optional ByVal SYMBOL_STR As String = "*", _
Optional ByVal OUTPUT As Integer = 0)

Dim f As Long
Dim g As Long
Dim h As Long
Dim i As Long
Dim j As Long
Dim k As Long
Dim l As Long

Dim ff As Long
Dim hh As Long
Dim ii As Long
Dim jj As Long
Dim kk As Long

Dim fff As Long
Dim hhh As Long
Dim iii As Long
Dim jjj As Long
Dim kkk As Long

Dim A_NROWS As Long
Dim A_NCOLUMNS As Long

Dim B_NROWS As Long
Dim B_NCOLUMNS As Long

Dim ATEMP_STR As String
Dim BTEMP_STR As String
Dim CTEMP_STR As String
Dim BINARY_STR As String

Dim ATEMP_ARR As Variant
Dim BTEMP_ARR As Variant
Dim CTEMP_ARR As Variant

Dim DOM_FLAG As Boolean
Dim FIND_FLAG As Boolean
Dim STOP_FLAG As Boolean
Dim ESSEN_FLAG As Boolean
Dim COMB_FLAG_ARR() As Boolean

Dim DATA_MATRIX As Variant
Dim ATEMP_MATRIX As Variant
Dim BTEMP_MATRIX As Variant

On Error GoTo ERROR_LABEL

DATA_MATRIX = DATA_RNG

A_NROWS = UBound(DATA_MATRIX, 1)
A_NCOLUMNS = UBound(DATA_MATRIX, 2)
f = 0

For i = 1 To A_NROWS
    If DATA_MATRIX(i, 1) <> 0 Then f = f + 1
Next i
hhh = f
g = 1
For i = 1 To A_NROWS
    If 2 ^ g >= A_NROWS Then Exit For
    g = g + 1
Next i

ReDim ATEMP_MATRIX(1 To f, 1 To 2)
ReDim COMB_FLAG_ARR(1 To f)
jj = 0

For ii = 1 To A_NROWS
    If DATA_MATRIX(ii, 1) <> 0 Then
        jj = jj + 1
        ATEMP_MATRIX(jj, 1) = CONVERT_DECIMAL_BINARY_FUNC(ii - 1, 0)
        ATEMP_MATRIX(jj, 1) = _
            String(g - Len(ATEMP_MATRIX(jj, 1)), "0") & ATEMP_MATRIX(jj, 1)
        k = 0
        j = Len(ATEMP_MATRIX(jj, 1))
        For i = 1 To j
            If Mid(ATEMP_MATRIX(jj, 1), i, 1) = "1" Then k = k + 1
        Next i
        ATEMP_MATRIX(jj, 2) = k
    End If
Next ii

ATEMP_MATRIX = BINARY_SORT_FUNC(ATEMP_MATRIX, 2, 1)
ReDim BTEMP_MATRIX(1 To A_NROWS, 1 To 2)

Do
    
    ReDim BTEMP_ARR(1 To 4 * A_NROWS, 1 To 2)
    ff = 0
    For ii = 1 To f
        For jj = ii + 1 To f
            If ATEMP_MATRIX(jj, 2) > ATEMP_MATRIX(ii, 2) + 1 Then Exit For
            If ATEMP_MATRIX(jj, 2) = ATEMP_MATRIX(ii, 2) + 1 Then
                l = 0
                ATEMP_STR = ""
                For i = 1 To Len(ATEMP_MATRIX(ii, 1))
                    If Mid(ATEMP_MATRIX(ii, 1), i, 1) <> _
                        Mid(ATEMP_MATRIX(jj, 1), i, 1) Then
                        ATEMP_STR = ATEMP_STR & SYMBOL_STR
                        l = l + 1
                    Else
                        ATEMP_STR = ATEMP_STR & Mid(ATEMP_MATRIX(ii, 1), i, 1)
                    End If
                Next i
                
                If l = 1 Then
                    COMB_FLAG_ARR(ii) = True
                    COMB_FLAG_ARR(jj) = True
                    FIND_FLAG = False
                    For kk = 1 To ff
                        If BTEMP_ARR(kk, 2) > ATEMP_MATRIX(ii, 2) Then Exit For
                        If BTEMP_ARR(kk, 1) = ATEMP_STR Then FIND_FLAG = True: Exit For
                    Next kk
                    If Not FIND_FLAG Then
                        ff = ff + 1
                        BTEMP_ARR(ff, 1) = ATEMP_STR
                        BTEMP_ARR(ff, 2) = ATEMP_MATRIX(ii, 2)
                    End If
                End If
            End If
        Next jj
    Next ii
    
    For i = 1 To f
        If Not COMB_FLAG_ARR(i) Then
            fff = fff + 1
            BTEMP_MATRIX(fff, 1) = ATEMP_MATRIX(i, 1)
            BTEMP_MATRIX(fff, 2) = ATEMP_MATRIX(i, 2)
        End If
    Next i
                    
    If ff > 0 Then
        f = ff
        ReDim ATEMP_MATRIX(1 To f, 1 To 2)
        ReDim COMB_FLAG_ARR(1 To f)
        For i = 1 To f
            ATEMP_MATRIX(i, 1) = BTEMP_ARR(i, 1)
            ATEMP_MATRIX(i, 2) = BTEMP_ARR(i, 2)
        Next i
    End If
    
Loop Until ff = 0
    
Erase ATEMP_MATRIX, BTEMP_ARR
ReDim ATEMP_MATRIX(1 To fff, 1 To A_NROWS)

For i = 1 To fff
    h = 0
    iii = InStr(1, BTEMP_MATRIX(i, 1), SYMBOL_STR)
    Do While iii > 0
        h = h + 1
        iii = InStr(iii + 1, BTEMP_MATRIX(i, 1), SYMBOL_STR)
    Loop

    If h = 0 Then
        A_NCOLUMNS = 1
        ReDim ATEMP_ARR(1 To A_NCOLUMNS)
        ATEMP_ARR(1) = CONVERT_BINARY_DECIMAL_FUNC(BTEMP_MATRIX(i, 1))
    Else
        A_NCOLUMNS = 2 ^ h
        ReDim ATEMP_ARR(1 To A_NCOLUMNS)
        For k = 1 To A_NCOLUMNS
            BINARY_STR = CONVERT_DECIMAL_BINARY_FUNC(k - 1, 0)
            BINARY_STR = String(h - Len(BINARY_STR), "0") & BINARY_STR
            iii = 0
            BTEMP_STR = BTEMP_MATRIX(i, 1)
            For j = 1 To h
                CTEMP_STR = Mid(BINARY_STR, j, 1)
                iii = InStr(iii + 1, BTEMP_MATRIX(i, 1), SYMBOL_STR)
                Mid(BTEMP_STR, iii, 1) = CTEMP_STR
            Next j
            ATEMP_ARR(k) = CONVERT_BINARY_DECIMAL_FUNC(BTEMP_STR)
        Next k
    End If
        
    For j = 1 To UBound(ATEMP_ARR)
        ii = ATEMP_ARR(j) + 1
        If DATA_MATRIX(ii, 1) = 1 Then
            ATEMP_MATRIX(i, ii) = CHR_STR
        End If
    Next j
Next i
    
B_NROWS = UBound(ATEMP_MATRIX, 1)
B_NCOLUMNS = UBound(ATEMP_MATRIX, 2)

ReDim CTEMP_ARR(1 To B_NCOLUMNS)
For i = 1 To UBound(CTEMP_ARR)
    CTEMP_ARR(i) = 1
Next i

For i = 1 To UBound(BTEMP_MATRIX)
    BTEMP_MATRIX(i, 2) = 1
Next i

GoSub 1983

Do
    GoSub 1983
    GoSub 1984
    ESSEN_FLAG = False
    For j = 1 To B_NCOLUMNS
        If CTEMP_ARR(j) = 1 Then
            For i = 1 To B_NROWS
                If BTEMP_MATRIX(i, 2) > 0 Then
                    If ATEMP_MATRIX(i, j) = CHR_STR Then Exit For
                End If
            Next i
            BTEMP_MATRIX(i, 2) = -1
            For hh = 1 To B_NCOLUMNS
                If ATEMP_MATRIX(i, hh) = CHR_STR Then CTEMP_ARR(hh) = 0
            Next hh
            ESSEN_FLAG = True
        End If
    Next j
    
    GoSub 1984
    DOM_FLAG = False
    For ii = 1 To B_NROWS
        If BTEMP_MATRIX(ii, 2) > 0 Then
            For jj = ii + 1 To B_NROWS
                If BTEMP_MATRIX(jj, 2) > 0 Then
                    jjj = 0
                    kkk = 0
                    For j = 1 To B_NCOLUMNS
                        If CTEMP_ARR(j) > 0 Then
                            If ATEMP_MATRIX(ii, j) = CHR_STR And _
                                ATEMP_MATRIX(jj, j) = "" Then jjj = jjj + 1
                            If ATEMP_MATRIX(ii, j) = "" And _
                                ATEMP_MATRIX(jj, j) = CHR_STR Then kkk = kkk + 1
                        End If
                    Next j
                    If jjj >= 0 And kkk = 0 Then
                        BTEMP_MATRIX(jj, 2) = 0
                        DOM_FLAG = True
                    ElseIf jjj = 0 And kkk > 0 Then
                        BTEMP_MATRIX(ii, 2) = 0
                        DOM_FLAG = True
                    End If
                End If
            Next jj
        End If
    Next ii
    
    If Not DOM_FLAG And Not ESSEN_FLAG Then
        For ii = 1 To B_NROWS
            If BTEMP_MATRIX(ii, 2) > 0 Then BTEMP_MATRIX(ii, 2) = 0: Exit For
        Next
    End If

Loop Until STOP_FLAG

ReDim BTEMP_ARR(1 To fff, 1 To 2)
For i = 1 To fff
    BTEMP_ARR(i, 1) = BTEMP_MATRIX(i, 1)
    BTEMP_ARR(i, 2) = BTEMP_MATRIX(i, 2)
Next i
ReDim BTEMP_MATRIX(1 To fff, 1 To 2)
For i = 1 To fff
    BTEMP_MATRIX(i, 1) = BTEMP_ARR(i, 1)
    BTEMP_MATRIX(i, 2) = BTEMP_ARR(i, 2)
Next i
    
ReDim BTEMP_ARR(1 To UBound(ATEMP_MATRIX, 1), 1 To UBound(ATEMP_MATRIX, 2))
For i = 1 To UBound(ATEMP_MATRIX, 1)
    For j = 1 To UBound(ATEMP_MATRIX, 2)
        BTEMP_ARR(i, j) = ATEMP_MATRIX(i, j)
    Next j
Next i
    
hhh = 0
For i = 1 To A_NROWS
    If DATA_MATRIX(i, 1) = 1 Then hhh = hhh + 1
Next i
    
ReDim ATEMP_MATRIX(1 To hhh, 1 To fff)
hh = 0
For j = 1 To UBound(BTEMP_ARR, 2)
    If DATA_MATRIX(j, 1) = 1 Then
        hh = hh + 1
        For i = 1 To UBound(BTEMP_ARR, 1)
            ATEMP_MATRIX(hh, i) = BTEMP_ARR(i, j)
        Next i
    End If
Next j

Select Case OUTPUT
    Case 0
        BINARY_QUINE_MCCLUSKEY_FUNC = ATEMP_MATRIX
    Case 1
        BINARY_QUINE_MCCLUSKEY_FUNC = BTEMP_MATRIX
    Case Else
        BINARY_QUINE_MCCLUSKEY_FUNC = DATA_MATRIX
End Select

'----------------------------------------------------------------------------------
Exit Function
'----------------------------------------------------------------------------------
1983: 'Column Sum
'----------------------------------------------------------------------------------
    STOP_FLAG = True
    For j = 1 To UBound(ATEMP_MATRIX, 2)
        If CTEMP_ARR(j) > 0 Then
            CTEMP_ARR(j) = 0
            For i = 1 To UBound(ATEMP_MATRIX, 1)
                If ATEMP_MATRIX(i, j) = CHR_STR And _
                    BTEMP_MATRIX(i, 2) > 0 Then CTEMP_ARR(j) = CTEMP_ARR(j) + 1
            Next i
            If CTEMP_ARR(j) > 0 Then STOP_FLAG = False
        End If
    Next j
'----------------------------------------------------------------------------------
Return
'----------------------------------------------------------------------------------
'----------------------------------------------------------------------------------
1984: 'Row sum
'----------------------------------------------------------------------------------
    For i = 1 To UBound(ATEMP_MATRIX, 1)
        If BTEMP_MATRIX(i, 2) > 0 Then
            BTEMP_MATRIX(i, 2) = 0
            For j = 1 To UBound(ATEMP_MATRIX, 2)
                If ATEMP_MATRIX(i, j) = CHR_STR And _
                    CTEMP_ARR(j) > 0 Then BTEMP_MATRIX(i, 2) = BTEMP_MATRIX(i, 2) + 1
            Next j
        End If
    Next i
'----------------------------------------------------------------------------------
Return
'----------------------------------------------------------------------------------
'----------------------------------------------------------------------------------
ERROR_LABEL:
BINARY_QUINE_MCCLUSKEY_FUNC = Err.number
End Function
