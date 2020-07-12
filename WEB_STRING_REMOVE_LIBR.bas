Attribute VB_Name = "WEB_STRING_REMOVE_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.

'************************************************************************************
'************************************************************************************
'FUNCTION      : REMOVE_CHARACTER_FUNC
'DESCRIPTION   : Remove Character
'LIBRARY       : STRING
'GROUP         : REMOVE
'ID            : 001
'UPDATE        : 02/20/2008
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'************************************************************************************
'************************************************************************************

Function REMOVE_CHARACTER_FUNC(ByVal DATA_STR As String, _
ByVal DELIM_CHR As String)
    
Dim i As Long
Dim j As Long
Dim TEMP_STR As String           'working string

On Error GoTo ERROR_LABEL

j = Len(DATA_STR)
TEMP_STR = ""
For i = 1 To j
    If Mid(DATA_STR, i, 1) <> DELIM_CHR Then
        TEMP_STR = TEMP_STR & Mid(DATA_STR, i, 1)
    End If
Next i
REMOVE_CHARACTER_FUNC = TEMP_STR

Exit Function
ERROR_LABEL:
REMOVE_CHARACTER_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : REMOVE_BRACKETS_FUNC
'DESCRIPTION   : Remove brackets
'LIBRARY       : STRING
'GROUP         : BRACKETS
'ID            : 002
'UPDATE        : 02/20/2008
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA

'************************************************************************************
'************************************************************************************

Function REMOVE_BRACKETS_FUNC(ByRef DATA_STR As String)
    
Dim i As Long
Dim j As Long

On Error GoTo ERROR_LABEL
    
If IsError(DATA_STR) Or DATA_STR = vbNullString Or _
    Not DATA_STR Like "*[[]*[]]*" Then
        REMOVE_BRACKETS_FUNC = DATA_STR
        Exit Function
End If

i = InStr(1, DATA_STR, "[", 1) + 1
j = InStrRev(DATA_STR, "]", -1, 1)
j = j - i

DATA_STR = Mid(DATA_STR, i, j)

If DATA_STR = vbNullString Then: GoTo ERROR_LABEL
REMOVE_BRACKETS_FUNC = DATA_STR

Exit Function
ERROR_LABEL:
REMOVE_BRACKETS_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : REMOVE_QUOTES_FUNC
'DESCRIPTION   : Removes the quotes from Text and returns the result.
'LIBRARY       : STRING
'GROUP         : QUOTES
'ID            : 003
'UPDATE        : 02/20/2008
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'************************************************************************************
'************************************************************************************

Function REMOVE_QUOTES_FUNC(ByVal DATA_STR As String)
On Error GoTo ERROR_LABEL
REMOVE_QUOTES_FUNC = Replace(DATA_STR, Chr(34), vbNullString, 1, -1, 0)
Exit Function
ERROR_LABEL:
REMOVE_QUOTES_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : REMOVE_CHARACTERS_FUNC
'DESCRIPTION   :
'LIBRARY       : STRING
'GROUP         : LINE
'ID            : 004
'UPDATE        : 02/20/2008
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'************************************************************************************
'************************************************************************************

Function REMOVE_CHARACTERS_FUNC(ByRef DATA_STR As String, _
Optional ByVal ILLEGAL_CHR As String = "/-:;!@#$%^&*()+=,<>")

Dim i As Long
Dim j As Long

On Error GoTo ERROR_LABEL

i = Len(ILLEGAL_CHR)
For j = 1 To i
    DATA_STR = Replace(DATA_STR, Mid(ILLEGAL_CHR, j, 1), "", 1, -1, vbBinaryCompare)
Next j

' If we made out of the loop, the name is valid.
REMOVE_CHARACTERS_FUNC = DATA_STR

Exit Function
ERROR_LABEL:
REMOVE_CHARACTERS_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : REMOVE_EXTRA_SPACES_FUNC
'DESCRIPTION   : SINGLE SPACE IN A STRING
'LIBRARY       : STRING
'GROUP         : LINE
'ID            : 005
'UPDATE        : 02/20/2008
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'************************************************************************************
'************************************************************************************

Function REMOVE_EXTRA_SPACES_FUNC(ByVal TEMP_STR As String) 'DATA_STR

Dim i As Long
'Dim TEMP_STR As String

On Error GoTo ERROR_LABEL

'TEMP_STR = DATA_STR
i = InStr(1, TEMP_STR, Space(2), 0)
Do Until i = 0
    TEMP_STR = Replace(TEMP_STR, Space(2), Space(1), 1, -1, 0)
    i = InStr(1, TEMP_STR, Space(2), 0)
Loop

REMOVE_EXTRA_SPACES_FUNC = TEMP_STR

Exit Function
ERROR_LABEL:
REMOVE_EXTRA_SPACES_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : REMOVE_NEW_LINE_FUNC
'DESCRIPTION   : Find vbNewLine for Chr(10).
'LIBRARY       : STRING
'GROUP         : LINE
'ID            : 006
'UPDATE        : 02/20/2008
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'************************************************************************************
'************************************************************************************

Function REMOVE_NEW_LINE_FUNC(ByVal DATA_STR As String)
On Error GoTo ERROR_LABEL
    REMOVE_NEW_LINE_FUNC = Replace(DATA_STR, vbNewLine, Chr(10), 1, -1, 0)
Exit Function
ERROR_LABEL:
REMOVE_NEW_LINE_FUNC = Err.number
End Function
