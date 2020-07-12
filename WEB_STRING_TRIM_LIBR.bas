Attribute VB_Name = "WEB_STRING_TRIM_LIBR"


Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.

'-------------------------------------------------------------------------------
Private Declare PtrSafe Function lstrlenW Lib "kernel32" (ByVal lpString As Long) As Long
'-------------------------------------------------------------------------------

'************************************************************************************
'************************************************************************************
'FUNCTION      : TRIM_CHARACTERS_FUNC

'DESCRIPTION   : If REVER_FLAG is False, the function returns the portion of
' TEXT_STR that is to the left of the first occurrence of LOOK_CHR.
' If REVER_FLAG is True, the function returns the portion of
' Text that is to the left of the last occurrence of LOOK_CHR.
' If Char is not found, the entire string TEXT_STR is returned.
' If COMP_TYPE is vbBinaryCompare, text is compared in
' a CASE-SENSITIVE manner ("A"<>"a"). If COMP_TYPE is any
' other value, text is compared in CASE-INSENSITIVE mode ("A" = "a").

'LIBRARY       : STRING
'GROUP         : TRIM
'ID            : 001
'UPDATE        : 02/20/2008
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'************************************************************************************
'************************************************************************************

Function TRIM_CHARACTERS_FUNC(ByVal DATA_STR As String, _
ByVal LOOK_CHR As String, _
Optional ByVal REVER_FLAG As Boolean = False, _
Optional ByVal COMP_TYPE As Long = 0)

'COMP_TYPE:
'vbUseCompareOption; –1 ; Performs a comparison using the
'setting of the Option Compare statement.

'vbBinaryCompare; 0 ; Performs a binary comparison.

'vbTextCompare; 1 ; Performs a textual comparison.

'vbDatabaseCompare ; 2 ; Microsoft Access only. Performs
'a comparison based on information in your database.


Dim i As Long

On Error GoTo ERROR_LABEL

If REVER_FLAG = False Then
    i = InStr(1, DATA_STR, LOOK_CHR, COMP_TYPE)
Else
    i = InStrRev(DATA_STR, LOOK_CHR, -1, COMP_TYPE)
End If

If i Then
    TRIM_CHARACTERS_FUNC = Left(DATA_STR, i - 1)
Else
    TRIM_CHARACTERS_FUNC = DATA_STR
End If

Exit Function
ERROR_LABEL:
TRIM_CHARACTERS_FUNC = Err.number
End Function



'************************************************************************************
'************************************************************************************



'************************************************************************************
'************************************************************************************
'FUNCTION      : STRIP_NULL_CHARACTERS_FUNC
'DESCRIPTION   : Return a string without the chr$(0) terminator.
'LIBRARY       : STRING
'GROUP         : NULL
'ID            : 002
'UPDATE        : 02/20/2008
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'************************************************************************************
'************************************************************************************
