Attribute VB_Name = "WEB_STRING_SPLIT_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.

'************************************************************************************
'************************************************************************************
'FUNCTION      : SPLIT_ARRAY_FUNC
'DESCRIPTION   : Returns a one-based, two-dimensional array containing a
'specified number of substrings inside an array.
'LIBRARY       : STRING
'GROUP         : SPLIT
'ID            : 001
'LAST UPDATE   : 12/02/2008
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'************************************************************************************
'************************************************************************************

Function SPLIT_ARRAY_FUNC(ByRef DATA_RNG As Variant, _
Optional ByVal DELIM_CHR As String = "|", _
Optional ByVal SKIP_STR As Variant = "", _
Optional ByVal AROW As Long = 1)

Dim i As Long
Dim j As Long

Dim ii As Long
Dim jj As Long
Dim kk As Long

Dim NROWS As Long
Dim NCOLUMNS As Long

Dim TEMP_STR As String
Dim DATA_STR As String
Dim TEMP_MATRIX() As String

On Error GoTo ERROR_LABEL

DATA_STR = DATA_RNG(AROW)
If DATA_STR = SKIP_STR Then: GoTo ERROR_LABEL

If Right(DATA_STR, Len(DELIM_CHR)) <> DELIM_CHR Then
    DATA_STR = DATA_STR & DELIM_CHR
End If

j = Len(DATA_STR)
NCOLUMNS = 0
For i = 1 To j
    If Mid(DATA_STR, i, Len(DELIM_CHR)) = DELIM_CHR Then: NCOLUMNS = NCOLUMNS + 1
Next i

NROWS = UBound(DATA_RNG) - LBound(DATA_RNG) + 1

ReDim TEMP_MATRIX(1 To NROWS, 1 To NCOLUMNS)

ii = 1
For kk = LBound(DATA_RNG) To UBound(DATA_RNG)
    DATA_STR = DATA_RNG(kk)
    If DATA_STR = SKIP_STR Then: GoTo 1983
    If Right(DATA_STR, Len(DELIM_CHR)) <> DELIM_CHR Then
        DATA_STR = DATA_STR & DELIM_CHR
    End If
    i = 1
    jj = 1
    Do
        j = InStr(i, DATA_STR, DELIM_CHR)
        If j = 0 Then: GoTo 1983
        TEMP_STR = Mid(DATA_STR, i, j - i)
        i = j + Len(DELIM_CHR)
        TEMP_MATRIX(ii, jj) = TEMP_STR
        jj = jj + 1
    Loop Until jj > UBound(TEMP_MATRIX, 2)
1983:
    ii = ii + 1
Next kk

SPLIT_ARRAY_FUNC = TEMP_MATRIX

Exit Function
ERROR_LABEL:
SPLIT_ARRAY_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : SPLIT_STRING_FUNC
'DESCRIPTION   : Returns a zero-based, one-dimensional array containing a
'specified number of substrings.
'LIBRARY       : STRING
'GROUP         : SPLIT
'ID            : 002
'LAST UPDATE   : 12/02/2008
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'************************************************************************************
'************************************************************************************

Function SPLIT_STRING_FUNC(ByVal DATA_STR As String, _
Optional ByVal DELIM_CHR As String = " ", _
Optional ByVal LIMIT_VAL As Long = -1, _
Optional ByVal COMP_VAL As Long = 1)

Dim i As Long
Dim j As Long
Dim k As Long
Dim TEMP_STR As String
Dim TEMP_ARR() As Variant
    
'TEMP_STR: String expression containing substrings
'and delimiters. If expression is a zero-length string(""),
'Split returns an empty array, that is, an array with no
'elements and no data.

'DELIM_CHR: String character used to identify substring
'limits. If omitted, the space character (" ") is assumed
'to be the delimiter. If delimiter is a zero-length string,
'a single-element array containing the entire expression
'string is returned.

'LIMIT_VAL: Number of substrings to be returned; –1 indicates
'that all substrings are returned.

'COMP_VAL: Numeric value indicating the kind of comparison to
'use when evaluating substrings

'vbUseCompareOption (–1 ) --> Performs a comparison using the
'setting of the Option Compare statement.
'vbBinaryCompare ( 0 ) --> Performs a binary comparison.
'vbTextCompare ( 1 ) -->  Performs a textual comparison.
'vbDatabaseCompare ( 2 )--> Microsoft Access only. Performs
'a comparison based on information in your database.

On Error GoTo ERROR_LABEL

TEMP_STR = DATA_STR
i = 0
If DELIM_CHR <> "" Then
    k = Len(DELIM_CHR)
    j = InStr(1, TEMP_STR, DELIM_CHR, COMP_VAL)
    Do While j
        ReDim Preserve TEMP_ARR(0 To i)
        TEMP_ARR(i) = Left(TEMP_STR, j - 1)
        TEMP_STR = Mid(TEMP_STR, j + k)
        i = i + 1
        If LIMIT_VAL <> -1 And i >= LIMIT_VAL Then Exit Do
        j = InStr(1, TEMP_STR, DELIM_CHR, COMP_VAL)
    Loop
End If

ReDim Preserve TEMP_ARR(0 To i)
TEMP_ARR(i) = TEMP_STR

SPLIT_STRING_FUNC = TEMP_ARR

Exit Function
ERROR_LABEL:
SPLIT_STRING_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : SPLIT_TEXT_FUNC
'DESCRIPTION   : This function splits the sentence into
'words and returns a string array of the words. Each
'element of the array contains one word.
'LIBRARY       : STRING
'GROUP         : SPLIT
'ID            : 003
'LAST UPDATE   : 12/02/2008
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA

'************************************************************************************
'************************************************************************************

Function SPLIT_TEXT_FUNC(ByVal DATA_STR As String, _
Optional ByVal DELIM_CHR As String = " ", _
Optional ByVal EXTRA_CHR As String = ":")
    
Dim i As Long
Dim j As Long
Dim TEMP_STR As String
     
On Error GoTo ERROR_LABEL

TEMP_STR = DATA_STR
' Replace tab characters with space characters.
TEMP_STR = Trim(Replace(TEMP_STR, vbTab, " ", 1, -1, 0))

' Filter all specified characters from the string.
i = Len(EXTRA_CHR)
For j = 1 To i
    TEMP_STR = Trim(Replace(TEMP_STR, Mid(EXTRA_CHR, j, 1), " ", 1, -1, 0))
Next j

' Loop until all consecutive space characters are
' replaced by a single space character.
Do While InStr(TEMP_STR, "  ")
    TEMP_STR = Replace(TEMP_STR, "  ", " ", 1, -1, 0)
Loop

' Split the sentence into an array of words and return
' the array. If a DELIM_CHR is specified, use it.

If Len(DELIM_CHR) = 0 Then
    SPLIT_TEXT_FUNC = SPLIT_STRING_FUNC(TEMP_STR, "", -1, 1)
Else
    SPLIT_TEXT_FUNC = SPLIT_STRING_FUNC(TEMP_STR, DELIM_CHR, -1, 1)
End If
    
Exit Function
ERROR_LABEL:
SPLIT_TEXT_FUNC = Err.number
End Function
