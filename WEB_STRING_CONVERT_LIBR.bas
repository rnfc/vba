Attribute VB_Name = "WEB_STRING_CONVERT_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.
                    
Private Type DATA_MAT_OBJ
    iRow As Long
    iCol As Long
    valo As String
End Type


'************************************************************************************
'************************************************************************************
'FUNCTION      : CONVERT_MATRIX_STRING2_FUNC
'DESCRIPTION   : Convert Matrix to Text
'LIBRARY       : STRING
'GROUP         : CONVERT
'ID            : 001
'LAST UPDATE   : 12/02/2008
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'************************************************************************************
'************************************************************************************

Function CONVERT_MATRIX_STRING2_FUNC(ByRef DATA_RNG As Variant, _
Optional ByVal DELIM_CHR As String = "|", _
Optional ByVal VERSION As Integer = 0)

Dim i As Long
Dim j As Long

Dim SROW As Long
Dim NROWS As Long

Dim SCOLUMN As Long
Dim NCOLUMNS As Long

Dim TEMP_CHR As String
Dim TEMP_STR As String
Dim TEMP_CRLF As String

Dim TEMP_ARR() As String
Dim DATA_MATRIX As Variant

On Error GoTo ERROR_LABEL
    
DATA_MATRIX = DATA_RNG

SROW = LBound(DATA_MATRIX, 1)
NROWS = UBound(DATA_MATRIX, 1)

SCOLUMN = LBound(DATA_MATRIX, 2)
NCOLUMNS = UBound(DATA_MATRIX, 2)

TEMP_STR = ""

'------------------------------------------------------------------------------------
Select Case VERSION
'------------------------------------------------------------------------------------
Case 0
'------------------------------------------------------------------------------------
        TEMP_CRLF = Chr(13) + Chr(10)
        For i = 1 To NROWS - 1
            For j = 1 To NCOLUMNS - 1
                TEMP_CHR = DATA_MATRIX(SROW + i - 1, SCOLUMN + j - 1)
                TEMP_STR = TEMP_STR & TEMP_CHR
                TEMP_STR = TEMP_STR & DELIM_CHR
            Next j
            TEMP_CHR = DATA_MATRIX(SROW + i - 1, SCOLUMN + NCOLUMNS - 1)
            TEMP_STR = TEMP_STR & TEMP_CHR
            TEMP_STR = TEMP_STR & DELIM_CHR
            TEMP_STR = TEMP_STR & TEMP_CRLF
        Next i
        For j = 1 To NCOLUMNS - 1
            TEMP_CHR = DATA_MATRIX(SROW + NROWS - 1, SCOLUMN + j - 1)
            TEMP_STR = TEMP_STR & TEMP_CHR
            TEMP_STR = TEMP_STR & DELIM_CHR
        Next j
        TEMP_CHR = DATA_MATRIX(SROW + NROWS - 1, SCOLUMN + NCOLUMNS - 1)
        TEMP_STR = TEMP_STR & TEMP_CHR
        TEMP_STR = TEMP_STR & TEMP_CRLF
        CONVERT_MATRIX_STRING2_FUNC = TEMP_STR
        Exit Function
'------------------------------------------------------------------------------------
Case 1 'Output in one-dimen array
'------------------------------------------------------------------------------------
        ReDim TEMP_ARR(1 To NROWS)
        For i = 1 To NROWS - 1
            TEMP_STR = ""
            For j = 1 To NCOLUMNS - 1
                TEMP_CHR = DATA_MATRIX(SROW + i - 1, SCOLUMN + j - 1)
                TEMP_STR = TEMP_STR & TEMP_CHR
                TEMP_STR = TEMP_STR & DELIM_CHR
            Next j
            TEMP_CHR = DATA_MATRIX(SROW + i - 1, SCOLUMN + NCOLUMNS - 1)
            TEMP_STR = TEMP_STR & TEMP_CHR
            TEMP_STR = TEMP_STR & DELIM_CHR
            TEMP_ARR(i) = TEMP_STR
        Next i
        TEMP_STR = ""
        For j = 1 To NCOLUMNS - 1
            TEMP_CHR = DATA_MATRIX(SROW + NROWS - 1, SCOLUMN + j - 1)
            TEMP_STR = TEMP_STR & TEMP_CHR
            TEMP_STR = TEMP_STR & DELIM_CHR
        Next j
        TEMP_CHR = DATA_MATRIX(SROW + NROWS - 1, SCOLUMN + NCOLUMNS - 1)
        TEMP_STR = TEMP_STR & TEMP_CHR
        TEMP_ARR(NROWS) = TEMP_STR
        CONVERT_MATRIX_STRING2_FUNC = TEMP_ARR
        Exit Function
'------------------------------------------------------------------------------------
Case Else 'Everything in two-dimen array
'------------------------------------------------------------------------------------
        ReDim TEMP_ARR(1 To NROWS, 1 To 1)
        For i = 1 To NROWS - 1
            TEMP_STR = ""
            For j = 1 To NCOLUMNS - 1
                TEMP_CHR = DATA_MATRIX(SROW + i - 1, SCOLUMN + j - 1)
                TEMP_STR = TEMP_STR & TEMP_CHR
                TEMP_STR = TEMP_STR & DELIM_CHR
            Next j
            TEMP_CHR = DATA_MATRIX(SROW + i - 1, SCOLUMN + NCOLUMNS - 1)
            TEMP_STR = TEMP_STR & TEMP_CHR
            TEMP_STR = TEMP_STR & DELIM_CHR
            TEMP_ARR(i, 1) = TEMP_STR
        Next i
        TEMP_STR = ""
        For j = 1 To NCOLUMNS - 1
            TEMP_CHR = DATA_MATRIX(SROW + NROWS - 1, SCOLUMN + j - 1)
            TEMP_STR = TEMP_STR & TEMP_CHR
            TEMP_STR = TEMP_STR & DELIM_CHR
        Next j
        TEMP_CHR = DATA_MATRIX(SROW + NROWS - 1, SCOLUMN + NCOLUMNS - 1)
        TEMP_STR = TEMP_STR & TEMP_CHR
        TEMP_ARR(NROWS, 1) = TEMP_STR
        CONVERT_MATRIX_STRING2_FUNC = TEMP_ARR
        Exit Function
'----------------------------------------------------------------------------------
End Select
'----------------------------------------------------------------------------------

Exit Function
ERROR_LABEL:
CONVERT_MATRIX_STRING2_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : CONVERT_STRING_MATRIX2_FUNC
'DESCRIPTION   : Convert Text to Matrix
'LIBRARY       : STRING
'GROUP         : CONVERT
'ID            : 002
'LAST UPDATE   : 12/02/2008
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'************************************************************************************
'************************************************************************************

Function CONVERT_STRING_MATRIX2_FUNC(ByRef DATA_STR As String, _
Optional ByVal DELIM_CHR As String = "|")

Dim i As Long
Dim j As Long
Dim k As Long

Dim ii As Long
Dim jj As Long
Dim kk As Long
Dim hh As Long

Dim NROWS As Long
Dim NCOLUMNS As Long

Dim ROW_STR As String
Dim COLUMN_STR As String

Dim TEMP_STR As String
Dim TEMP_CRLF As String

Dim TEMP_VECTOR() As String
Dim TEMP_MATRIX() As String

On Error GoTo ERROR_LABEL
    
i = 0
TEMP_STR = DATA_STR
TEMP_CRLF = Chr(13) + Chr(10)

ii = 1
jj = InStr(ii, TEMP_STR, TEMP_CRLF)
If jj = 0 Then jj = Len(TEMP_STR) + 1 'Length of the string + 1

NCOLUMNS = COUNT_STRING_FUNC(Trim(Mid(TEMP_STR, ii, jj - ii)), _
DELIM_CHR, DELIM_CHR, 1, 1)
    
NROWS = COUNT_CHARACTERS_FUNC(Trim(TEMP_STR), TEMP_CRLF)
 
ReDim TEMP_VECTOR(1 To NROWS * (NCOLUMNS + 1), 1 To 3)

NROWS = 0
NCOLUMNS = 0

Do While jj > 0
    ROW_STR = Trim(Mid(TEMP_STR, ii, jj - ii)) 'bounding string
    kk = 1
    hh = InStr(kk, ROW_STR, DELIM_CHR)
    If hh = 0 Then: GoTo 1984

    If hh = 0 Then hh = Len(ROW_STR)
    COLUMN_STR = Trim(Mid(ROW_STR, kk, hh - kk)) 'Get each Character
        i = i + 1  'row COUNTER
        If NROWS < i Then NROWS = i
        j = 0
        Do
            j = j + 1
            If NCOLUMNS < j Then NCOLUMNS = j
            k = k + 1
            TEMP_VECTOR(k, 1) = CStr(i)
            TEMP_VECTOR(k, 2) = CStr(j)
            TEMP_VECTOR(k, 3) = Trim(COLUMN_STR)
            kk = hh + 1
            If kk > Len(ROW_STR) Then Exit Do ' Perfect
            hh = InStr(kk, ROW_STR, DELIM_CHR)
            If hh = 0 Then hh = Len(ROW_STR) ' Perfect
            COLUMN_STR = Trim(Mid(ROW_STR, kk, hh - kk))
       Loop
1984:

    ii = jj + 2
    jj = InStr(ii, TEMP_STR, TEMP_CRLF)
Loop

NROWS = CLng(TEMP_VECTOR(k, 1))
NCOLUMNS = CLng(TEMP_VECTOR(k, 2))

If NROWS > 0 And NCOLUMNS > 0 Then
    ReDim TEMP_MATRIX(1 To NROWS, 1 To NCOLUMNS)
    For ii = 1 To k
        i = CLng(TEMP_VECTOR(ii, 1))
        j = CLng(TEMP_VECTOR(ii, 2))
        TEMP_MATRIX(i, j) = TEMP_VECTOR(ii, 3)
    Next ii
    CONVERT_STRING_MATRIX2_FUNC = TEMP_MATRIX

Else
    GoTo ERROR_LABEL
End If

Exit Function
ERROR_LABEL:
CONVERT_STRING_MATRIX2_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : CONVERT_STRING_MATRIX_FUNC
'DESCRIPTION   : Convert string to matrix
'LIBRARY       : STRING
'GROUP         : CONVERT
'ID            : 003
'UPDATE        : 01/21/2008
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'************************************************************************************
'************************************************************************************

Function CONVERT_STRING_MATRIX_FUNC(ByRef DATA_STR As String, _
Optional ByVal VERSION As Integer = 1)

Dim i As Long
Dim j As Long
Dim k As Long
Dim l As Long

Dim ii As Long
Dim jj As Long
Dim kk As Long
Dim ll As Long

Dim NROWS As Long
Dim NCOLUMNS As Long

Dim COUNTER As Long

Dim COL_STR As String
Dim ROW_STR As String

Dim ATEMP_STR As String
Dim BTEMP_STR As String
Dim TEMP_VAL As Variant
Dim TEMP_STR As String

Dim CRLF_STR As String
Dim TEMP_MATRIX() As String
Dim TEMP_GROUP() As DATA_MAT_OBJ

On Error GoTo ERROR_LABEL
ReDim TEMP_GROUP(1 To 1)

CRLF_STR = Chr(13) + Chr(10)

Select Case VERSION
'-------------------------------------------------------------------------------------
'-------------------------------------------------------------------------------------
Case 0 'If InStr(1, DATA_STR, "Array") > 0 Then 'list format
'-------------------------------------------------------------------------------------
'DATA_STR = "Array(1..3, 1..3," & Chr(13) + Chr(10) & _
"(1, 1) = 1," & Chr(13) + Chr(10) & _
"(1, 2) = 2," & Chr(13) + Chr(10) & _
"(1, 3) = 3," & Chr(13) + Chr(10) & _
"(2, 1) = 4," & Chr(13) + Chr(10) & _
"(2, 2) = 5," & Chr(13) + Chr(10) & _
"(2, 3) = 6," & Chr(13) + Chr(10) & _
"(3, 1) = 7," & Chr(13) + Chr(10) & _
"(3, 2) = 8," & Chr(13) + Chr(10) & _
"(3, 3) = 234 " & Chr(13) + Chr(10) & _
")"
'-------------------------------------------------------------------------------------
    i = 1
    j = InStr(i, DATA_STR, CRLF_STR) 'join string
    
    TEMP_STR = ""
    Do
        TEMP_STR = TEMP_STR + Mid(DATA_STR, i, j - i)
        If Right(TEMP_STR, 1) = "\" Then
            TEMP_STR = Left(TEMP_STR, Len(TEMP_STR) - 1)
        End If
        'Debug.Print TEMP_STR
        i = j + 2
        j = InStr(i, DATA_STR, CRLF_STR)
    Loop Until j = 0
    DATA_STR = TEMP_STR
    
    i = InStr(1, DATA_STR, "..")
    j = InStr(1, DATA_STR, ",")
    NROWS = CLng(Mid(DATA_STR, i + 2, j - i - 2))
    i = InStr(i + 1, DATA_STR, "..")
    j = InStr(j + 1, DATA_STR, ",")
    NCOLUMNS = CLng(Mid(DATA_STR, i + 2, j - i - 2))
    
    j = InStr(1, DATA_STR, "(")
    Do
        
        i = InStr(j + 1, DATA_STR, "(")
        j = InStr(i, DATA_STR, ")")
        j = InStr(j, DATA_STR, ",")
        If j > 0 Then
            ROW_STR = Trim(Mid(DATA_STR, i, j - i))
        Else
            ROW_STR = Trim(Mid(DATA_STR, i, Len(DATA_STR) - i))
        End If
        
        TEMP_VAL = 0: ii = 0: jj = 0
        k = InStr(1, ROW_STR, "(")
        l = InStr(1, ROW_STR, ",")
        ii = CLng(Mid(ROW_STR, k + 1, l - k - 1))
        k = l
        l = InStr(l, ROW_STR, ")")
        jj = CLng(Mid(ROW_STR, k + 1, l - k - 1))
        k = InStr(1, ROW_STR, "=")
        l = InStr(k + 1, ROW_STR, ",")
        
        If l = 0 Then: l = Len(ROW_STR) + 1
        TEMP_VAL = (Mid(ROW_STR, k + 1, l - k - 1))
        
        If ii > 0 And jj > 0 Then
            COUNTER = COUNTER + 1
            ReDim Preserve TEMP_GROUP(1 To COUNTER)
            TEMP_GROUP(COUNTER).iRow = ii
            TEMP_GROUP(COUNTER).iCol = jj
            TEMP_GROUP(COUNTER).valo = TEMP_VAL
        End If
    Loop Until j = 0

'-------------------------------------------------------------------------------------
'-------------------------------------------------------------------------------------
Case 1 'ElseIf InStr(1, DATA_STR, "|") Or InStr(1, DATA_STR, "[") Or _
            InStr(1, DATA_STR, "{") > 0 Then
'-------------------------------------------------------------------------------------
'DATA_STR = "[" & Chr(13) + Chr(10) & _
"[1,2,3] " & Chr(13) + Chr(10) & _
"[4,5,6] " & Chr(13) + Chr(10) & _
"[7,8,9] " & Chr(13) + Chr(10) & _
"]"
'-------------------------------------------------------------------------------------
    'matrix/vector format
'-------------------------------------------------------------------------------------
        i = 1
        j = InStr(i, DATA_STR, CRLF_STR)
        If j = 0 Then j = Len(DATA_STR) + 1
        Do While j > 0
            ROW_STR = Trim(Mid(DATA_STR, i, j - i)) 'bounding string
            kk = 0
            If Left(ROW_STR, 1) = "|" Then kk = InStr(2, ROW_STR, "|")
            If Left(ROW_STR, 1) = "[" Then kk = InStr(2, ROW_STR, "]")
            If Left(ROW_STR, 1) = "{" Then kk = InStr(2, ROW_STR, "}")
            If kk > 0 Then
                ROW_STR = Mid(ROW_STR, 1, kk)
                k = 2
                l = InStr(k, ROW_STR, ",")
                If l = 0 Then l = Len(ROW_STR)
                COL_STR = Trim(Mid(ROW_STR, k, l - k))
                If COL_STR <> "" Then
                    ii = ii + 1  'row counter
                    If NROWS < ii Then NROWS = ii
                    jj = 0
                    Do
                        If COL_STR <> "" Then
                            jj = jj + 1
                            If NCOLUMNS < jj Then NCOLUMNS = jj
                            'catch this entry
                            COUNTER = COUNTER + 1
                            ReDim Preserve TEMP_GROUP(1 To COUNTER)
                            TEMP_GROUP(COUNTER).iRow = ii
                            TEMP_GROUP(COUNTER).iCol = jj
                            TEMP_GROUP(COUNTER).valo = (COL_STR)
                        End If
                        'next column
                        k = l + 1
                        If k >= Len(ROW_STR) Then: Exit Do
                        l = InStr(k, ROW_STR, ",")
                        If l = 0 Then l = Len(ROW_STR)
                        COL_STR = Trim(Mid(ROW_STR, k, l - k))
                   Loop
                End If
            End If
            i = j + 2
            j = InStr(i, DATA_STR, CRLF_STR)
        Loop
    
'-------------------------------------------------------------------------------------
'-------------------------------------------------------------------------------------
Case 2 'Same as 'ElseIf InStr(1, DATA_STR, "|") Or InStr(1, DATA_STR, "[") Or _
InStr(1, DATA_STR, "{") > 0 Then; but without CRLF
'-------------------------------------------------------------------------------------
'DATA_STR = [[1,2,3],[4,5,6],[7,8,9]]
'-------------------------------------------------------------------------------------

     DATA_STR = Mid(DATA_STR, InStr(1, DATA_STR, "[", 1) + 1, _
                                 InStrRev(DATA_STR, "]", -1, 1) - _
                                 InStr(1, DATA_STR, "[", 1) + 1) _
                                 'PARSE INITIAL BRACKETS
      i = 1
      j = Len(DATA_STR) + 1 'Length of the string + 1
    
      Do
            ROW_STR = Trim(Mid(DATA_STR, i, j - i)) 'bounding string
            kk = 0
            If Left(ROW_STR, 1) = "|" Then kk = InStr(2, ROW_STR, "|")
            If Left(ROW_STR, 1) = "[" Then kk = InStr(2, ROW_STR, "]")
            If Left(ROW_STR, 1) = "{" Then kk = InStr(2, ROW_STR, "}")
            
            If kk > 0 Then 'Analyze each row
                ROW_STR = Mid(ROW_STR, 1, kk)
                ll = Len(ROW_STR)
                k = 2
                l = InStr(k, ROW_STR, ",")
                If l = 0 Then l = Len(ROW_STR)
                COL_STR = Trim(Mid(ROW_STR, k, l - k)) 'Get each Character
                If COL_STR <> "" Then
                    ii = ii + 1  'row COUNTER
                    If NROWS < ii Then NROWS = ii
                    jj = 0
                    Do
                        If COL_STR <> "" Then
                            jj = jj + 1
                            If NCOLUMNS < jj Then NCOLUMNS = jj
                            'catch this entry
                            COUNTER = COUNTER + 1
                            ReDim Preserve TEMP_GROUP(1 To COUNTER)
                            TEMP_GROUP(COUNTER).iRow = ii
                            TEMP_GROUP(COUNTER).iCol = jj
                            TEMP_GROUP(COUNTER).valo = (COL_STR)
                        End If
                        'next column
                        k = l + 1
                        If k >= Len(ROW_STR) Then Exit Do
                        l = InStr(k, ROW_STR, ",")
                        If l = 0 Then l = Len(ROW_STR)
                        COL_STR = Trim(Mid(ROW_STR, k, l - k))
                   Loop
                End If
            End If
            i = i + ll + 1
    Loop While (j - i) > 0

'-------------------------------------------------------------------------------------
'-------------------------------------------------------------------------------------
Case Else
'-------------------------------------------------------------------------------------
'DATA_STR = "1 2 3 4 5 6" & Chr(13) + Chr(10) & _
"4 5 6 7 8 9" & Chr(13) + Chr(10) & _
"10 11 12 13 14 15" & Chr(13) + Chr(10)
'-------------------------------------------------------------------------------------
    i = 1
    j = InStr(1, DATA_STR, CRLF_STR)
    Do While j > 0
        ROW_STR = Trim(Mid(DATA_STR, i, j - i)) + " " 'check columns wrap
        If InStr(1, ROW_STR, "Column", vbTextCompare) > 0 Then
            ii = 0:  kk = NCOLUMNS
        Else
            ii = ii + 1  'row counter
            If NROWS < ii Then: NROWS = ii
            k = 1
            jj = kk  'column counter
            For l = 1 To Len(ROW_STR) - 1
                ATEMP_STR = Trim(Mid(ROW_STR, l, 1))
                BTEMP_STR = Trim(Mid(ROW_STR, l + 1, 1))
                If ATEMP_STR <> "" And BTEMP_STR = "" Then  'edge found
                    COL_STR = Trim(Mid(ROW_STR, k, l - k + 1))
                    jj = jj + 1
                    If NCOLUMNS < jj Then NCOLUMNS = jj 'catch this entry
                    COUNTER = COUNTER + 1
                    ReDim Preserve TEMP_GROUP(1 To COUNTER)
                    TEMP_GROUP(COUNTER).iRow = ii
                    TEMP_GROUP(COUNTER).iCol = jj
                    TEMP_GROUP(COUNTER).valo = Val(COL_STR)
                    k = l + 1
                End If
            Next l
        End If
        i = j + 2
        j = InStr(i, DATA_STR, CRLF_STR)
    Loop

'-------------------------------------------------------------------------------------
'-------------------------------------------------------------------------------------
End Select
'-------------------------------------------------------------------------------------
'-------------------------------------------------------------------------------------

'transfer data from TEMP_GROUP to array
If NROWS > 0 And NCOLUMNS > 0 Then
    ReDim TEMP_MATRIX(1 To NROWS, 1 To NCOLUMNS)
    For i = 1 To COUNTER
        ii = TEMP_GROUP(i).iRow
        jj = TEMP_GROUP(i).iCol
        TEMP_MATRIX(ii, jj) = TEMP_GROUP(i).valo
    Next i
    CONVERT_STRING_MATRIX_FUNC = TEMP_MATRIX
Else
    GoTo ERROR_LABEL '"Incorrect dimension"
End If

'NICO_TICO = "1 2 3 4 5 6" & Chr(13) + Chr(10) & _
"4 5 6 7 8 9" & Chr(13) + Chr(10) & _
"10 11 12 13 14 15" & Chr(13) + Chr(10)

'NICO_TICO = "[" & Chr(13) + Chr(10) & _
"[1,2,3] " & Chr(13) + Chr(10) & _
"[4,5,6] " & Chr(13) + Chr(10) & _
"[7,8,9] " & Chr(13) + Chr(10) & _
"]"

'NICO_TICO = "Array(1..3, 1..3," & Chr(13) + Chr(10) & _
"(1, 1) = 1," & Chr(13) + Chr(10) & _
"(1, 2) = 2," & Chr(13) + Chr(10) & _
"(1, 3) = 3," & Chr(13) + Chr(10) & _
"(2, 1) = 4," & Chr(13) + Chr(10) & _
"(2, 2) = 5," & Chr(13) + Chr(10) & _
"(2, 3) = 6," & Chr(13) + Chr(10) & _
"(3, 1) = 7," & Chr(13) + Chr(10) & _
"(3, 2) = 8," & Chr(13) + Chr(10) & _
"(3, 3) = 234 " & Chr(13) + Chr(10) & _
")"

Exit Function
ERROR_LABEL:
CONVERT_STRING_MATRIX_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : CONVERT_STRING_MATRIX_SPARSE_FUNC
'DESCRIPTION   : Convert String to Matrix (Sparse Method)
'LIBRARY       : STRING
'GROUP         : CONVERT
'ID            : 004
'LAST UPDATE   : 12/02/2008
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'************************************************************************************
'************************************************************************************

Function CONVERT_STRING_MATRIX_SPARSE_FUNC(ByRef DATA_STR As String, _
Optional ByVal DELIM_CHR As String = "|")

Dim i As Long
Dim j As Long
Dim k As Long
Dim l As Long

Dim ii As Long
Dim jj As Long

Dim NROWS As Long
Dim NCOLUMNS As Long

Dim COUNTER As Long

Dim COL_STR As String
Dim ROW_STR As String

Dim FIRST_STR As String
Dim SECOND_STR As String
Dim CRLF_STR As String

Dim TEMP_MATRIX As Variant
Dim TEMP_GROUP() As DATA_MAT_OBJ

On Error GoTo ERROR_LABEL

CRLF_STR = Chr(13) + Chr(10)
    
If Right(DATA_STR, 2) <> CRLF_STR Then: DATA_STR = DATA_STR & CRLF_STR

i = 1
j = InStr(1, DATA_STR, CRLF_STR)
Do While j > 0
    ROW_STR = Trim(Mid(DATA_STR, i, j - i)) + " " 'check columns wrap
        ii = ii + 1  'row counter
        If NROWS < ii Then: NROWS = ii
        k = 1
        jj = 0  'column counter
        For l = 1 To Len(ROW_STR) - 1
            FIRST_STR = Trim(Mid(ROW_STR, l, 1))
            SECOND_STR = Trim(Mid(ROW_STR, l + 1, 1))
            If FIRST_STR <> "" And SECOND_STR = DELIM_CHR Then  'edge found
'                    COL_STR = Trim(Mid(ROW_STR, k, l - k + 1))
                COL_STR = Trim(Mid(ROW_STR, IIf(k = 1, k, k + 1), _
                          IIf(k = 1, l - k + 1, l - k)))
                
                jj = jj + 1
                If NCOLUMNS < jj Then NCOLUMNS = jj 'catch this entry
                COUNTER = COUNTER + 1
                ReDim Preserve TEMP_GROUP(1 To COUNTER)
                TEMP_GROUP(COUNTER).iRow = ii
                TEMP_GROUP(COUNTER).iCol = jj
                TEMP_GROUP(COUNTER).valo = COL_STR
                k = l + 1
            End If
        Next l
    i = j + 2
    j = InStr(i, DATA_STR, CRLF_STR)
Loop

'transfer data from TEMP_GROUP to array
If NROWS > 0 And NCOLUMNS > 0 Then
    ReDim TEMP_MATRIX(1 To NROWS, 1 To NCOLUMNS)
    For i = 1 To COUNTER
        ii = TEMP_GROUP(i).iRow
        jj = TEMP_GROUP(i).iCol
        TEMP_MATRIX(ii, jj) = TEMP_GROUP(i).valo
    Next i
    CONVERT_STRING_MATRIX_SPARSE_FUNC = CONVERT_STRING_NUMBER_FUNC(TEMP_MATRIX)
Else
    GoTo ERROR_LABEL '"Incorrect dimension"
End If

Exit Function
ERROR_LABEL:
CONVERT_STRING_MATRIX_SPARSE_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : CONVERT_MATRIX_STRING_FUNC
'DESCRIPTION   : Convert matrix to string
'LIBRARY       : STRING
'GROUP         : CONVERT
'ID            : 005
'UPDATE        : 01/21/2008
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'************************************************************************************
'************************************************************************************

Function CONVERT_MATRIX_STRING_FUNC(ByRef DATA_RNG As Variant, _
Optional ByVal VERSION As Integer = 0)

Dim i As Long
Dim j As Long

Dim SROW As Long
Dim NROWS As Long

Dim SCOLUMN As Long
Dim NCOLUMNS As Long

Dim TEMP_STR As String
Dim DATA_STR As String
Dim DATA_MATRIX As Variant

On Error GoTo ERROR_LABEL

DATA_MATRIX = DATA_RNG
    
SROW = LBound(DATA_MATRIX, 1)
NROWS = UBound(DATA_MATRIX, 1)

SCOLUMN = LBound(DATA_MATRIX, 2)
NCOLUMNS = UBound(DATA_MATRIX, 2)
    
'--------------------------------------------------------------------------------
Select Case VERSION
'--------------------------------------------------------------------------------
Case 0 'OUTPUT FORMAT: [[1,2,3],[4,5,6],[7,8,9]]
'--------------------------------------------------------------------------------
    DATA_STR = ""
    For i = 1 To NROWS
        DATA_STR = DATA_STR & "["
        For j = 1 To NCOLUMNS
            TEMP_STR = DATA_MATRIX(SROW + i - 1, SCOLUMN + j - 1)
            DATA_STR = DATA_STR & TEMP_STR
            If j < NCOLUMNS Then DATA_STR = DATA_STR & ","
        Next j
        DATA_STR = DATA_STR & "]"
        If i < NROWS Then DATA_STR = DATA_STR & ","
    Next i
    DATA_STR = LCase("[" & DATA_STR & "]")
'     DATA_STR = "[" & DATA_STR & "]"
'--------------------------------------------------------------------------------
Case Else 'OUTPUT FORMAT: [ 1, 2, 3; 4, 5, 6; 7, 8, 9]
'--------------------------------------------------------------------------------
    DATA_STR = "["
    For i = 1 To NROWS
        For j = 1 To NCOLUMNS
            TEMP_STR = DATA_MATRIX(SROW + i - 1, SCOLUMN + j - 1)
            DATA_STR = DATA_STR & TEMP_STR
            If j < NCOLUMNS Then DATA_STR = DATA_STR & ","
        Next j
        If i < NROWS Then DATA_STR = DATA_STR & ";"
    Next i
    DATA_STR = DATA_STR & "]"
'--------------------------------------------------------------------------------
End Select
'--------------------------------------------------------------------------------

CONVERT_MATRIX_STRING_FUNC = DATA_STR

Exit Function
ERROR_LABEL:
CONVERT_MATRIX_STRING_FUNC = Err.number
End Function
