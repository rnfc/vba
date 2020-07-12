Attribute VB_Name = "DATE_PARSE_LIBR"


Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.



'************************************************************************************
'************************************************************************************
'FUNCTION      : PARSE_CURRENT_TIME_FUNC
'DESCRIPTION   : Parse current time
'LIBRARY       : DATE
'GROUP         : TIME
'ID            : 002
'LAST UPDATE   : 12/02/2008
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA

'************************************************************************************
'************************************************************************************

Function PARSE_CURRENT_TIME_FUNC(Optional ByVal DELIM_CHR As String = "_")

On Error GoTo ERROR_LABEL

    PARSE_CURRENT_TIME_FUNC = Replace(Replace(Replace(Now, ":", DELIM_CHR, _
                    1, -1, vbBinaryCompare), "/", DELIM_CHR, _
                    1, -1, vbBinaryCompare), " ", DELIM_CHR, _
                    1, -1, vbBinaryCompare)
Exit Function
ERROR_LABEL:
PARSE_CURRENT_TIME_FUNC = Err.number
End Function


Function PARSE_DATE_TIME_FUNC(ByVal DATE_VAL As Variant, _
Optional ByVal VERSION As Integer = 1)

Dim DATE_STR As Variant

On Error GoTo ERROR_LABEL

DATE_STR = CLng(DATE_VAL)

'-------------------------------------------------------------------------------
Select Case VERSION
'-------------------------------------------------------------------------------
Case 0
'-------------------------------------------------------------------------------
'Digits  Example     Remarks
'4   9298    Converted to 2-Sep-1998
'5   11298   Converted to 12-Jan-1998, NOT 2-Nov-1998
'6   112298  Converted to 22-Nov-1998
'7   1231998     Converted to 23-Jan-1998, NOT 3-Dec-1998
'8   11221998    Converted to 22-Nov-1998
    
    Select Case Len(DATE_STR)
        Case 4 ' e.g., 9298 = 2-Sep-1998
            DATE_STR = Left(DATE_STR, 1) & "/" & Mid(DATE_STR, 2, 1) & "/" & Right(DATE_STR, 2)
        Case 5 ' e.g., 11298 = 12-Jan-1998 NOT 2-Nov-1998
            DATE_STR = Left(DATE_STR, 1) & "/" & Mid(DATE_STR, 2, 2) & "/" & Right(DATE_STR, 2)
        Case 6 ' e.g., 090298 = 2-Sep-1998
            DATE_STR = Left(DATE_STR, 2) & "/" & Mid(DATE_STR, 3, 2) & "/" & Right(DATE_STR, 2)
        Case 7 ' e.g., 1231998 = 23-Jan-1998 NOT 3-Dec-1998
            DATE_STR = Left(DATE_STR, 1) & "/" & Mid(DATE_STR, 2, 2) & "/" & Right(DATE_STR, 4)
        Case 8 ' e.g., 09021998 = 2-Sep-1998
            DATE_STR = Left(DATE_STR, 2) & "/" & Mid(DATE_STR, 3, 2) & "/" & Right(DATE_STR, 4)
        Case Else
            GoTo ERROR_LABEL
    End Select
    PARSE_DATE_TIME_FUNC = DateValue(DATE_STR)
'-------------------------------------------------------------------------------
Case Else
'-------------------------------------------------------------------------------
'The next procedure will test the location and value of the changed cell.
'If it is in the range A1:A10, the value is converted to a proper time.
'The value in the cell must be between 1 and 6 numbers in length. Otherwise
'an error will occur.  The rules for conversion are described below:
'Digits  Example     Remarks
'1   1   Converted to 12:01:00 AM
'2   23  Converted to 12:23:00 AM
'3   123     Converted to 1:23:00 AM
'4   1234    Converted to 12:34:00
'5   12345   Converted to 1:23:45, NOT 12:03:45
'6   123456  Converted to 12:34:56
    Select Case Len(DATE_STR)
        Case 1 ' e.g., 1 = 00:01 AM
            DATE_STR = "00:0" & DATE_STR
        Case 2 ' e.g., 12 = 00:12 AM
            DATE_STR = "00:" & DATE_STR
        Case 3 ' e.g., 735 = 7:35 AM
            DATE_STR = Left(DATE_STR, 1) & ":" & Right(DATE_STR, 2)
        Case 4 ' e.g., 1234 = 12:34
            DATE_STR = Left(DATE_STR, 2) & ":" & Right(DATE_STR, 2)
        Case 5 ' e.g., 12345 = 1:23:45 NOT 12:03:45
            DATE_STR = Left(DATE_STR, 1) & ":" & Mid(DATE_STR, 2, 2) & ":" & Right(DATE_STR, 2)
        Case 6 ' e.g., 123456 = 12:34:56
            DATE_STR = Left(DATE_STR, 2) & ":" & Mid(DATE_STR, 3, 2) & ":" & Right(DATE_STR, 2)
        Case Else
            GoTo ERROR_LABEL
    End Select
    PARSE_DATE_TIME_FUNC = TimeValue(DATE_STR)
'-------------------------------------------------------------------------------
End Select
'-------------------------------------------------------------------------------

Exit Function
ERROR_LABEL:
PARSE_DATE_TIME_FUNC = Err.number
End Function

