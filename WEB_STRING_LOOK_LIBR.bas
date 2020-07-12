Attribute VB_Name = "WEB_STRING_LOOK_LIBR"


Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.

'************************************************************************************
'************************************************************************************
'FUNCTION      : LOOK_STRING_FUNC
'DESCRIPTION   : Look for a String
'LIBRARY       : STRING
'GROUP         : LOOK
'ID            : 001
'UPDATE        : 02/20/2008
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'************************************************************************************
'************************************************************************************

Function LOOK_STRING_FUNC(ByRef DATA_STR As String, _
ByRef LOOK_STR As String, _
Optional ByVal DELIM_CHR As String = ";")

Dim i As Long
Dim j As Long
Dim k As Long

On Error GoTo ERROR_LABEL

i = InStr(1, DATA_STR, LOOK_STR, 1)
If i = 0 Then: GoTo ERROR_LABEL

j = InStrRev(DATA_STR, DELIM_CHR, i)
k = InStr(i, DATA_STR, DELIM_CHR)

DATA_STR = Mid(DATA_STR, j + 1, k - j - 1)

LOOK_STRING_FUNC = DATA_STR

Exit Function
ERROR_LABEL:
LOOK_STRING_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : LOOK_STRING_MAT_FUNC
'DESCRIPTION   : Look string inside a matrix
'LIBRARY       : STRING
'GROUP         : LOOK
'ID            : 002
'UPDATE        : 02/20/2008
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'************************************************************************************
'************************************************************************************

Function LOOK_STRING_MAT_FUNC(ByRef DATA_RNG As Variant, _
ByVal LOOK_STR As String, _
Optional ByVal SROW As Long = 1, _
Optional ByVal OUTPUT As Integer = 0)

Dim i As Long
Dim j As Long
Dim k As Long

Dim ii As Long
Dim jj As Long
Dim kk As Long

Dim NROWS As Long
Dim NCOLUMNS As Long

Dim TEMP_STR As String
Dim DELIM_CHR As String

Dim TEMP_ARR As Variant
Dim DATA_MATRIX As Variant

On Error GoTo ERROR_LABEL

DATA_MATRIX = DATA_RNG
NCOLUMNS = UBound(DATA_MATRIX, 2)

If LOOK_STR = "" Then kk = 1: Exit Function 'extract sep char
DELIM_CHR = DECIMAL_SEPARATOR_FUNC() '"," or

If InStr(1, LOOK_STR, "+") > 0 Then DELIM_CHR = "+" 'AND

ii = 1
i = 0
ReDim TEMP_ARR(0 To i)

Do
    jj = InStr(ii, LOOK_STR, DELIM_CHR)
    If jj = 0 Then jj = Len(LOOK_STR) + 1
    i = i + 1
    ReDim Preserve TEMP_ARR(i)
    TEMP_ARR(i) = LCase(Trim(Mid(LOOK_STR, ii, jj - ii)))
    ii = jj + 1
Loop Until ii > Len(LOOK_STR)

NROWS = i

kk = 0

For i = 1 To NROWS
    k = 0
    For j = 1 To NCOLUMNS
        TEMP_STR = LCase(DATA_MATRIX(SROW, j))
        If TEMP_STR Like "*" + TEMP_ARR(i) + "*" Then
            k = 1:  Exit For
        End If
    Next j
    
    If k = 0 Then
        kk = 0
        If DELIM_CHR = "+" Then Exit For
    Else
        kk = 1
        If DELIM_CHR = "," Then Exit For
    End If
Next i
    
    Select Case OUTPUT
        Case 0
            LOOK_STRING_MAT_FUNC = TEMP_ARR
        Case Else
            LOOK_STRING_MAT_FUNC = TEMP_STR
    End Select
    
Exit Function
ERROR_LABEL:
LOOK_STRING_MAT_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : LOOK_STRING_FLAG_FUNC
'DESCRIPTION   : Check for an expression to occur in a string
'LIBRARY       : STRING
'GROUP         : LOOK
'ID            : 003
'UPDATE        : 02/20/2008
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'************************************************************************************
'************************************************************************************

Function LOOK_STRING_FLAG_FUNC(ByVal DATA_STR As String, _
ByVal LOOK_STR As String)

Dim i As Long
Dim ATEMP_STR As String
Dim BTEMP_STR As String

On Error GoTo ERROR_LABEL

LOOK_STRING_FLAG_FUNC = False

ATEMP_STR = " " & LOOK_STR & " "
BTEMP_STR = " " & DATA_STR & " "
i = InStr(1, ATEMP_STR, BTEMP_STR, 1)
If i > 0 Then: LOOK_STRING_FLAG_FUNC = True 'function mono variable

Exit Function
ERROR_LABEL:
LOOK_STRING_FLAG_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : LOOK_STRING_SUFFIX_FUNC
'DESCRIPTION   : Returns the extension of the specified string
'LIBRARY       : STRING
'GROUP         : LOOK
'ID            : 004
'UPDATE        : 02/20/2008
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'************************************************************************************
'************************************************************************************

Function LOOK_STRING_SUFFIX_FUNC(ByRef DATA_STR As String, _
Optional ByVal DELIM_CHR As String = ".")

Dim i As Long

On Error GoTo ERROR_LABEL

i = InStrRev(DATA_STR, DELIM_CHR)

If i = 0 Then
    LOOK_STRING_SUFFIX_FUNC = vbNullString
Else
    LOOK_STRING_SUFFIX_FUNC = Mid(DATA_STR, i + 1)
End If

Exit Function
ERROR_LABEL:
LOOK_STRING_SUFFIX_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : LOOK_STRING_SUFFIX_ARR_FUNC
'DESCRIPTION   : Extract string extension from a string array
'LIBRARY       : STRING
'GROUP         : LOOK
'ID            : 005
'UPDATE        : 02/20/2008
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'************************************************************************************
'************************************************************************************

Function LOOK_STRING_SUFFIX_ARR_FUNC(ByRef DATA_RNG As Variant, _
ByRef LOOK_STR As String)

Dim i As Long
Dim j As Long

Dim DATA_STR As String

On Error GoTo ERROR_LABEL

j = Len(LOOK_STR)
i = LBound(DATA_RNG)
DATA_STR = DATA_RNG(i)

Do Until DATA_STR = ""
    If DATA_STR <> vbNullString Then
        If StrComp(Left(DATA_STR, j), LOOK_STR) = 0 Then 'Equal to...
            DATA_STR = Mid(DATA_STR, j + 1)
            Exit Do
        End If
    End If
    i = i + 1
    DATA_STR = DATA_RNG(i)
    If i >= UBound(DATA_RNG) Then: Exit Do
Loop
LOOK_STRING_SUFFIX_ARR_FUNC = Array(DATA_STR, i)

Exit Function
ERROR_LABEL:
LOOK_STRING_SUFFIX_ARR_FUNC = Err.number
End Function
