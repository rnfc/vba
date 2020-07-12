Attribute VB_Name = "EXCEL_OUTLOOK_VALIDATE_LIBR"


Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.

'************************************************************************************
'************************************************************************************
'FUNCTION      : OUTLOOK_VALIDATE_EMAIL_ADDRESS_FUNC
'DESCRIPTION   : Verify string e-mail address
'LIBRARY       : WEB
'GROUP         : EMAIL
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'************************************************************************************
'************************************************************************************

Function OUTLOOK_VALIDATE_EMAIL_ADDRESS_FUNC(ByVal EMAIL_ADDRESS As String)

Dim EVAL_STR As String
Dim TEMP_STR As String

On Error GoTo ERROR_LABEL

TEMP_STR = EMAIL_ADDRESS

'----------------------------------------------------------------------------
If InStr(1, TEMP_STR, "@") Then
'----------------------------------------------------------------------------
    EVAL_STR = Left(EMAIL_ADDRESS, (InStr(1, EMAIL_ADDRESS, "@")) - 1)
    If Len(EVAL_STR) < 1 Then
        OUTLOOK_VALIDATE_EMAIL_ADDRESS_FUNC = False
        Exit Function
    End If
    EVAL_STR = Mid(TEMP_STR, (InStr(1, TEMP_STR, "@")) + 1, Len(TEMP_STR))
    If (Len(EVAL_STR) - 1) < 1 Then
        OUTLOOK_VALIDATE_EMAIL_ADDRESS_FUNC = False
        Exit Function
    End If
    If InStr(1, EVAL_STR, ".") Then
        EVAL_STR = Mid(EVAL_STR, InStr(1, EVAL_STR, ".") + 1, Len(TEMP_STR))
        EVAL_STR = Right(EVAL_STR, 3)
        If (Len(EVAL_STR)) < 2 Then
            OUTLOOK_VALIDATE_EMAIL_ADDRESS_FUNC = False
            Exit Function
        End If
    Else
        OUTLOOK_VALIDATE_EMAIL_ADDRESS_FUNC = False
        Exit Function
    End If
'----------------------------------------------------------------------------
Else
'----------------------------------------------------------------------------
    OUTLOOK_VALIDATE_EMAIL_ADDRESS_FUNC = False
    Exit Function
'----------------------------------------------------------------------------
End If
'----------------------------------------------------------------------------
OUTLOOK_VALIDATE_EMAIL_ADDRESS_FUNC = True
Exit Function
ERROR_LABEL:
OUTLOOK_VALIDATE_EMAIL_ADDRESS_FUNC = False
End Function
