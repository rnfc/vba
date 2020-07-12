Attribute VB_Name = "WEB_NUMBER_CONVERT_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.


'************************************************************************************
'************************************************************************************
'FUNCTION      : CONVERT_STRING_NUMBER_FUNC
'DESCRIPTION   : Parse Web Number Function
'LIBRARY       : WEB_NUMBER
'GROUP         : CONVERT
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 2012.04.07
'************************************************************************************
'************************************************************************************

Function CONVERT_STRING_NUMBER_FUNC(ByRef DATA_RNG As Variant)

Dim i As Long
Dim j As Long

Dim SROW As Long
Dim NROWS As Long

Dim SCOLUMN As Long
Dim NCOLUMNS As Long

Dim TEMP1_VAL As Variant
Dim TEMP2_VAL As Variant
Dim FACTOR_VAL As Double
Dim DATA_MATRIX As Variant

'Dim DCHR_STR As String

On Error GoTo ERROR_LABEL

'DCHR_STR = DECIMAL_SEPARATOR_FUNC()
'-------------------------------------------------------------------------------
If IsArray(DATA_RNG) = True Then
'-------------------------------------------------------------------------------
    SROW = LBound(DATA_RNG, 1): NROWS = UBound(DATA_RNG, 1)
    SCOLUMN = LBound(DATA_RNG, 2): NCOLUMNS = UBound(DATA_RNG, 2)
    ReDim DATA_MATRIX(SROW To NROWS, SCOLUMN To NCOLUMNS)
    For j = SCOLUMN To NCOLUMNS
        For i = SROW To NROWS
            TEMP1_VAL = Trim(DATA_RNG(i, j)): GoSub PARSE_LINE
            DATA_MATRIX(i, j) = TEMP1_VAL
        Next i
    Next j
    CONVERT_STRING_NUMBER_FUNC = DATA_MATRIX
'-------------------------------------------------------------------------------
Else
'-------------------------------------------------------------------------------
    TEMP1_VAL = Trim(DATA_RNG): GoSub PARSE_LINE
    CONVERT_STRING_NUMBER_FUNC = TEMP1_VAL
'-------------------------------------------------------------------------------
End If
'-------------------------------------------------------------------------------

'-------------------------------------------------------------------------------
Exit Function
'-------------------------------------------------------------------------------
PARSE_LINE:
'-------------------------------------------------------------------------------
    If TEMP1_VAL = "-" Then TEMP1_VAL = "0"
    If TEMP1_VAL = "--" Then TEMP1_VAL = "0"
    If TEMP1_VAL = "---" Then TEMP1_VAL = "0"
    If TEMP1_VAL = Chr(150) Then TEMP1_VAL = "0"
    If Left(TEMP1_VAL, 1) = "$" Then TEMP1_VAL = Mid(TEMP1_VAL, 2)
    If Left(TEMP1_VAL, 1) = "(" And Right(TEMP1_VAL, 1) = ")" Then
        TEMP1_VAL = "-" & Mid(TEMP1_VAL, 2, Len(TEMP1_VAL) - 2)
    End If
    TEMP2_VAL = TEMP1_VAL
    '---------------------------------------------------------------------------------------------
    If (Right(UCase(TEMP2_VAL), 2) = "PM") Or (Right(UCase(TEMP2_VAL), 2) = "AM") Then
    '---------------------------------------------------------------------------------------------
        TEMP1_VAL = TEMP2_VAL
        'If IsDate(TEMP2_VAL) = True Then: TEMP1_VAL = CDate(TEMP2_VAL)
    '---------------------------------------------------------------------------------------------
    Else
    '---------------------------------------------------------------------------------------------
        Select Case True
        Case Right(TEMP2_VAL, 1) = "B"
            TEMP2_VAL = Left(TEMP2_VAL, Len(TEMP2_VAL) - 1)
            FACTOR_VAL = 1000000
        Case Right(TEMP2_VAL, 1) = "M"
            TEMP2_VAL = Left(TEMP2_VAL, Len(TEMP2_VAL) - 1)
            FACTOR_VAL = 1000
        Case Right(TEMP2_VAL, 1) = "K"
            TEMP2_VAL = Left(TEMP2_VAL, Len(TEMP2_VAL) - 1)
            FACTOR_VAL = 1000
        Case Right(TEMP2_VAL, 1) = "%"
            TEMP2_VAL = Left(TEMP2_VAL, Len(TEMP2_VAL) - 1)
            FACTOR_VAL = 0.01
        Case Right(TEMP2_VAL, 4) = " Mil"
            TEMP2_VAL = Left(TEMP2_VAL, Len(TEMP2_VAL) - 4)
            FACTOR_VAL = 1000000
        Case Right(TEMP2_VAL, 5) = " Mill"
            TEMP2_VAL = Left(TEMP2_VAL, Len(TEMP2_VAL) - 5)
            FACTOR_VAL = 1000000
        Case Right(TEMP2_VAL, 4) = " Bil"
            TEMP2_VAL = Left(TEMP2_VAL, Len(TEMP2_VAL) - 4)
            FACTOR_VAL = 1000000000
        Case Right(TEMP2_VAL, 5) = " Bill"
            TEMP2_VAL = Left(TEMP2_VAL, Len(TEMP2_VAL) - 5)
            FACTOR_VAL = 1000000000
        Case Else
            FACTOR_VAL = 1
        End Select
        If IsNumeric(TEMP2_VAL) = True Then 'IS_NUMERIC_FUNC(TEMP2_VAL,DCHR_STR)
            TEMP1_VAL = CDec(TEMP2_VAL) * FACTOR_VAL
        ElseIf IsDate(TEMP2_VAL) = True Then
            TEMP1_VAL = DateSerial(Year(TEMP2_VAL), Month(TEMP2_VAL), Day(TEMP2_VAL))
        Else
            TEMP1_VAL = TEMP2_VAL
        End If
    '---------------------------------------------------------------------------------------------
    End If
    '---------------------------------------------------------------------------------------------
'-------------------------------------------------------------------------------
Return
'-------------------------------------------------------------------------------
ERROR_LABEL:
CONVERT_STRING_NUMBER_FUNC = Err.number
End Function

Function CONVERT_NUMBER_UNITS_FUNC(ByVal DATA_VAL As Integer, _
ByRef NUMB_TEXT_ARR As Variant, _
Optional ByRef USE_AND_FLAG As Boolean = False)

Dim kk As Integer
 
On Error GoTo ERROR_LABEL

If DATA_VAL > 99 Then
   kk = DATA_VAL \ 100
   CONVERT_NUMBER_UNITS_FUNC = NUMB_TEXT_ARR(kk) & " Hundred "
   DATA_VAL = DATA_VAL - (kk * 100)
End If
 
If USE_AND_FLAG = True Then
   CONVERT_NUMBER_UNITS_FUNC = CONVERT_NUMBER_UNITS_FUNC & "and "
End If

If DATA_VAL > 20 Then
   kk = DATA_VAL \ 10
   CONVERT_NUMBER_UNITS_FUNC = CONVERT_NUMBER_UNITS_FUNC & NUMB_TEXT_ARR(kk + 18) & " "
   DATA_VAL = DATA_VAL - (kk * 10)
End If
 
If DATA_VAL > 0 Then
   CONVERT_NUMBER_UNITS_FUNC = CONVERT_NUMBER_UNITS_FUNC & NUMB_TEXT_ARR(DATA_VAL) & " "
End If

Exit Function
ERROR_LABEL:
CONVERT_NUMBER_UNITS_FUNC = Err.number
End Function

Function CONVERT_STRING_LETTERS_NUMBER_FUNC(ByVal DATA_VAL As Variant)

Dim TEMP_STR As String
On Error GoTo ERROR_LABEL
TEMP_STR = DATA_VAL
If Asc(TEMP_STR) > 57 Then 'TEMP_STR is a letter
    CONVERT_STRING_LETTERS_NUMBER_FUNC = Asc(TEMP_STR) - 55
Else
    CONVERT_STRING_LETTERS_NUMBER_FUNC = Val(TEMP_STR)
End If

Exit Function
ERROR_LABEL:
CONVERT_STRING_LETTERS_NUMBER_FUNC = Err.number
End Function

Private Function CONVERT_STRING_DECIMALS_FUNC(ByVal DATA_VAL As Variant)
  
Dim TEMP_STR As String
On Error GoTo ERROR_LABEL

TEMP_STR = DATA_VAL

If Not DECIMAL_SEPARATOR_FUNC() = "." Then _
TEMP_STR = Replace(TEMP_STR, ".", ",", 1, -1, 0)

CONVERT_STRING_DECIMALS_FUNC = CDec(TEMP_STR)

Exit Function
ERROR_LABEL:
CONVERT_STRING_DECIMALS_FUNC = Err.number
End Function

Function CONVERT_NUMBER_TIME_FUNC(ByVal DATA_VAL As String)
'1 1 Converted to 12:01:00 AM
'2 23 Converted to 12:23:00 AM
'3 123 Converted to 1:23:00 AM
'4 1234 Converted to 12:34:00
'5 12345 Converted to 1:23:45, NOT 12:03:45
'6 123456 Converted to 12:34:56
    
Dim TIME_STR As String

On Error GoTo ERROR_LABEL
    
Select Case Len(DATA_VAL)
Case 1 ' e.g., 1 = 00:01 AM
    TIME_STR = "00:0" & DATA_VAL
Case 2 ' e.g., 12 = 00:12 AM
    TIME_STR = "00:" & DATA_VAL
Case 3 ' e.g., 735 = 7:35 AM
    TIME_STR = Left(DATA_VAL, 1) & ":" & Right(DATA_VAL, 2)
Case 4 ' e.g., 1234 = 12:34
    TIME_STR = Left(DATA_VAL, 2) & ":" & Right(DATA_VAL, 2)
Case 5 ' e.g., 12345 = 1:23:45 NOT 12:03:45
    TIME_STR = Left(DATA_VAL, 1) & ":" & Mid(DATA_VAL, 2, 2) & ":" & Right(DATA_VAL, 2)
Case 6 ' e.g., 123456 = 12:34:56
    TIME_STR = Left(DATA_VAL, 2) & ":" & Mid(DATA_VAL, 3, 2) & ":" & Right(DATA_VAL, 2)
Case Else
    GoTo ERROR_LABEL
End Select

CONVERT_NUMBER_TIME_FUNC = TIME_STR 'TimeValue(TIME_STR)
    
Exit Function
ERROR_LABEL:
CONVERT_NUMBER_TIME_FUNC = Err.number
End Function


Function CONVERT_NUMBER_DATE_FUNC(ByVal DATA_VAL As String)
'Digits Example Remarks
'4 9298 Converted to 2-Sep-1998
'5 11298 Converted to 12-Jan-1998, NOT 2-Nov-1998
'6 112298 Converted to 22-Nov-1998
'7 1231998 Converted to 23-Jan-1998, NOT 3-Dec-1998
'8 11221998 Converted to 22-Nov-1998

Dim DATE_STR As String

On Error GoTo ERROR_LABEL

Select Case Len(DATA_VAL)
Case 4 ' e.g., 9298 = 2-Sep-1998
    DATE_STR = Left(DATA_VAL, 1) & "/" & Mid(DATA_VAL, 2, 1) & "/" & Right(DATA_VAL, 2)
Case 5 ' e.g., 11298 = 12-Jan-1998 NOT 2-Nov-1998
    DATE_STR = Left(DATA_VAL, 1) & "/" & Mid(DATA_VAL, 2, 2) & "/" & Right(DATA_VAL, 2)
Case 6 ' e.g., 090298 = 2-Sep-1998
    DATE_STR = Left(DATA_VAL, 2) & "/" & Mid(DATA_VAL, 3, 2) & "/" & Right(DATA_VAL, 2)
Case 7 ' e.g., 1231998 = 23-Jan-1998 NOT 3-Dec-1998
    DATE_STR = Left(DATA_VAL, 1) & "/" & Mid(DATA_VAL, 2, 2) & "/" & Right(DATA_VAL, 4)
Case 8 ' e.g., 09021998 = 2-Sep-1998
    DATE_STR = Left(DATA_VAL, 2) & "/" & Mid(DATA_VAL, 3, 2) & "/" & Right(DATA_VAL, 4)
Case Else
    GoTo ERROR_LABEL
End Select

CONVERT_NUMBER_DATE_FUNC = DateValue(DATE_STR)
    
Exit Function
ERROR_LABEL:
CONVERT_NUMBER_DATE_FUNC = Err.number
End Function
