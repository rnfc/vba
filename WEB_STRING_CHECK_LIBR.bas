Attribute VB_Name = "WEB_STRING_CHECK_LIBR"
'--------------------------------------------------------------------------------
'--------------------------------------------------------------------------------
Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.
'--------------------------------------------------------------------------------
'--------------------------------------------------------------------------------


'************************************************************************************
'************************************************************************************
'FUNCTION      : CHECK_STRING_FUNC
'DESCRIPTION   : This returns True if DATA_STR is a syntactically valid
'string, otherwise False
'LIBRARY       : STRING
'GROUP         : VALID
'ID            : 001
'UPDATE        : 02/20/2008
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'************************************************************************************
'************************************************************************************

Function CHECK_STRING_FUNC(ByVal DATA_STR As String, _
Optional ByVal VALID_STR As String = "[A-Z]*.[A-Z]*")
On Error GoTo ERROR_LABEL
    CHECK_STRING_FUNC = (DATA_STR Like VALID_STR)
Exit Function
ERROR_LABEL:
CHECK_STRING_FUNC = False
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : CHECK_STRING_CHARACTERS_FUNC

'DESCRIPTION   : This function returns True or False indicating whether the
'string is valid. The following characters, and the space characters,
'are invalid in names:  / - : ; ! @ # $ % ^ & *( ) + = , < >

'LIBRARY       : STRING
'GROUP         : VALID
'ID            : 002
'UPDATE        : 02/20/2008
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'************************************************************************************
'************************************************************************************

Function CHECK_STRING_CHARACTERS_FUNC(ByVal NAME_STR As String, _
Optional ByVal ILLEGAL_CHR As String = " /-:;!@#$%^&*()+=,<>")

Dim i As Long
Dim j As Long
Dim k As Long
Dim l As Long

On Error GoTo ERROR_LABEL

If Trim(NAME_STR) = vbNullString Then
    CHECK_STRING_CHARACTERS_FUNC = False
    Exit Function
End If

' Test each character in NAME_STR against each character in
' ILLEGAL_CHR. If a match is found, get out and return False.

k = Len(NAME_STR)
l = Len(ILLEGAL_CHR)

For i = 1 To k
    For j = 1 To l
        If StrComp(Mid(NAME_STR, i, 1), _
            Mid(ILLEGAL_CHR, j, 1), _
            vbBinaryCompare) = 0 Then
                CHECK_STRING_CHARACTERS_FUNC = False
                Exit Function
        End If
    Next j
Next i

' If we made out of the loop, the name is valid.
CHECK_STRING_CHARACTERS_FUNC = True

Exit Function
ERROR_LABEL:
CHECK_STRING_CHARACTERS_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : CHECK_STRING_LENGTH_FUNC
'DESCRIPTION   : This function returns True or False indicating whether
'the string length is longer than j
'LIBRARY       : STRING
'GROUP         : VALID
'ID            : 003
'UPDATE        : 02/20/2008
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'************************************************************************************
'************************************************************************************

Function CHECK_STRING_LENGTH_FUNC(ByVal DATA_STR As String, _
Optional ByVal j As Long = 10)
On Error GoTo ERROR_LABEL
CHECK_STRING_LENGTH_FUNC = (Len(DATA_STR) <= j) And (Len(Trim(DATA_STR)) > 0)
Exit Function
ERROR_LABEL:
CHECK_STRING_LENGTH_FUNC = False
End Function
