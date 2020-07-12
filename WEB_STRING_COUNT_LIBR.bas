Attribute VB_Name = "WEB_STRING_COUNT_LIBR"


Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.

'************************************************************************************
'************************************************************************************
'FUNCTION      : COUNT_STRING_FUNC
'DESCRIPTION   : LOCATE AND COUNT CHR IN A STRING
'LIBRARY       : STRING
'GROUP         : COUNT
'ID            : 001
'UPDATE        : 02/20/2008
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'************************************************************************************
'************************************************************************************

Function COUNT_STRING_FUNC(ByRef DATA_STR As String, _
ByVal LOOK_CHR As String, _
ByVal DELIM_CHR As String, _
Optional ByVal START_VAL As Long = 1, _
Optional ByVal VERSION As Integer = 0)

Dim i As Long
Dim j As Long
Dim k As Long

Dim ii As Long
Dim jj As Long
Dim kk As Long

On Error GoTo ERROR_LABEL

ii = 0
kk = InStr(START_VAL, DATA_STR, LOOK_CHR)
k = 0

i = 0
Do
    i = i + 1
    jj = InStr(ii + 1, DATA_STR, DELIM_CHR)
    If jj <> 0 Then: j = j + 1
    If jj = (kk - 1) Then: k = j
    ii = jj
    If i > 100000 Then: Exit Do
Loop While jj > 0

Select Case VERSION
Case 0
    COUNT_STRING_FUNC = k
    'How Many Delim you need to Reach a Position in a
    'String e.g., A|B|C|D --> 2 to reach C
Case Else
    COUNT_STRING_FUNC = j 'Number of Delim Chr in a String
End Select

Exit Function
ERROR_LABEL:
COUNT_STRING_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : COUNT_CHARACTERS_FUNC
'DESCRIPTION   : Count number of Delim sets in a string
'LIBRARY       : STRING
'GROUP         : COUNT
'ID            : 002
'UPDATE        : 02/20/2008
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'************************************************************************************
'************************************************************************************

Function COUNT_CHARACTERS_FUNC(ByVal DATA_STR As String, _
Optional ByVal DELIM_CHR As String = "|")

Dim i As Long
Dim j As Long

On Error GoTo ERROR_LABEL

j = 0
i = InStr(1, DATA_STR, DELIM_CHR)
Do While i > 0
  i = i + 1
  j = j + 1
  i = InStr(i, DATA_STR, DELIM_CHR)
Loop
COUNT_CHARACTERS_FUNC = j

Exit Function
ERROR_LABEL:
COUNT_CHARACTERS_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : COUNT_COMMAS_FUNC
'DESCRIPTION   : Count number of commas between parenthesis. Ignore
'commas in subparenthesis
'LIBRARY       : STRING
'GROUP         : COUNT
'ID            : 003
'UPDATE        : 02/20/2008
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'************************************************************************************
'************************************************************************************

Function COUNT_COMMAS_FUNC(ByVal DATA_STR As String)

Dim i As Long
Dim j As Long 'numCommas
Dim k As Long 'numOpenPar

On Error GoTo ERROR_LABEL

Do
    i = i + 1
    If Mid(DATA_STR, i, 1) = "(" Then
        k = k + 1
    ElseIf Mid(DATA_STR, i, 1) = ")" Then
        k = k - 1
    ElseIf Mid(DATA_STR, i, 1) = "," And k = 1 Then
        j = j + 1
    End If
Loop While k > 0 And i < Len(DATA_STR)
If i = Len(DATA_STR) And k > 0 Then j = 0

COUNT_COMMAS_FUNC = j 'j + 1

Exit Function
ERROR_LABEL:
COUNT_COMMAS_FUNC = Err.number
End Function


Function COUNT_TICKERS_FUNC(ByRef DATA_RNG As Variant)

Dim i As Long
Dim j As Long

Dim NROWS As Long
Dim DATA_VECTOR As Variant

On Error GoTo ERROR_LABEL

DATA_VECTOR = DATA_RNG
If UBound(DATA_VECTOR, 1) = 1 Then: _
DATA_VECTOR = MATRIX_TRANSPOSE_FUNC(DATA_VECTOR)

NROWS = UBound(DATA_VECTOR, 1)

j = NROWS
For i = 1 To NROWS
    If IsError(DATA_VECTOR(i, 1)) = True Then
        j = j - 1
    Else
        If Trim(DATA_VECTOR(i, 1)) = "" Then: j = j - 1
    End If
Next i

COUNT_TICKERS_FUNC = j

Exit Function
ERROR_LABEL:
COUNT_TICKERS_FUNC = Err.number
End Function


