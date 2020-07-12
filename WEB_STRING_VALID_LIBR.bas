Attribute VB_Name = "WEB_STRING_VALID_LIBR"


Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.

'************************************************************************************
'************************************************************************************
'FUNCTION      : IS_LETTER_FUNC
'DESCRIPTION   : Returns an Integer representing the character code
'corresponding to the first letter in a string.
'LIBRARY       : STRING
'GROUP         : VALIDATION
'ID            : 001
'LAST UPDATE   : 12/02/2008
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA

'************************************************************************************
'************************************************************************************

Function IS_LETTER_FUNC(ByRef DATA_STR As Variant, _
Optional ByVal REFER_CHR As String = "_")
  
Dim i As Long
Dim TEMP_STR As String

On Error GoTo ERROR_LABEL

TEMP_STR = DATA_STR
If TEMP_STR = "" Then
  IS_LETTER_FUNC = False
  Exit Function
End If

i = Asc(TEMP_STR)

If (65 <= i And i <= 90) Or (97 <= i _
And i <= 122) Or (TEMP_STR = REFER_CHR) Then
  IS_LETTER_FUNC = True
Else
  IS_LETTER_FUNC = False
End If

Exit Function
ERROR_LABEL:
IS_LETTER_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : IS_DIGIT_FUNC
'DESCRIPTION   : Check if it is a digit
'LIBRARY       : STRING
'GROUP         : VALIDATION
'ID            : 002
'LAST UPDATE   : 12/02/2008
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA

'************************************************************************************
'************************************************************************************

Function IS_DIGIT_FUNC(ByVal DATA_STR As String)
Dim i As Long
Dim TEMP_STR As String
On Error GoTo ERROR_LABEL
TEMP_STR = DATA_STR
i = Asc(TEMP_STR)
IS_DIGIT_FUNC = (48 <= i And i <= 57)
Exit Function
ERROR_LABEL:
IS_DIGIT_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : IS_VARIABLE_FUNC
'DESCRIPTION   : Check if it is a variable name
'LIBRARY       : STRING
'GROUP         : VALIDATION
'ID            : 003
'LAST UPDATE   : 12/02/2008
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'************************************************************************************
'************************************************************************************

Function IS_VARIABLE_FUNC(ByRef DATA_STR As Variant)
  
Dim i As Long
Dim ATEMP_STR As Variant
Dim BTEMP_STR As String

On Error GoTo ERROR_LABEL

ATEMP_STR = DATA_STR
If IS_LETTER_FUNC(Left(ATEMP_STR, 1), "_") Then
    For i = 2 To Len(ATEMP_STR)
        BTEMP_STR = Mid(ATEMP_STR, i, 1)
        If Not IS_LETTER_FUNC(BTEMP_STR, "_") Then _
        If Not IS_DIGIT_FUNC(BTEMP_STR) Then _
        IS_VARIABLE_FUNC = False: Exit Function
    Next i
    IS_VARIABLE_FUNC = True
Else
    IS_VARIABLE_FUNC = False
End If

Exit Function
ERROR_LABEL:
IS_VARIABLE_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : IS_INTEGER_FUNC
'DESCRIPTION   : Check if value is integer
'LIBRARY       : STRING
'GROUP         : VALIDATION
'ID            : 004
'LAST UPDATE   : 12/02/2008
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'************************************************************************************
'************************************************************************************

Function IS_INTEGER_FUNC(ByVal DATA_VAL As Variant)

Dim epsilon As Variant

On Error GoTo ERROR_LABEL

epsilon = 5 * 10 ^ -14

If IsNumeric(DATA_VAL) = False Then
    IS_INTEGER_FUNC = False
    Exit Function
End If

If Abs(Round(DATA_VAL, 0) - DATA_VAL) > epsilon Then
      IS_INTEGER_FUNC = False
Else: IS_INTEGER_FUNC = True
End If

'check if DATA_VAL value is integer
'raises an error if variable DATA_VAL is not integer
'Dim DATA_VAL As Double
'Dim BTEMP_VAL As Double
'Dim ATEMP_VAL As Double
'epsilon = 5 * 10 ^ -14
'ATEMP_VAL = Round(DATA_VAL, 0)
'BTEMP_VAL = Abs(ATEMP_VAL - DATA_VAL)
'If BTEMP_VAL > epsilon Then CHECK_INT_FUNC = "" 'raises an error
'CHECK_INT_FUNC = ATEMP_VAL

Exit Function
ERROR_LABEL:
IS_INTEGER_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : IS_NUMERIC_FUNC
'DESCRIPTION   : Check if the reference is numeric
'LIBRARY       : STRING
'GROUP         : VALIDATION
'ID            : 005
'LAST UPDATE   : 12/02/2008
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'************************************************************************************
'************************************************************************************

Function IS_NUMERIC_FUNC(ByVal DATA_VAL As Variant, _
Optional ByRef DECIM_SEPAR_CHR As String = ".")

Dim i As Long
Dim TEMP_VAL As Variant

On Error GoTo ERROR_LABEL

TEMP_VAL = DATA_VAL
If DECIM_SEPAR_CHR = "" Then: DECIM_SEPAR_CHR = DECIMAL_SEPARATOR_FUNC()

IS_NUMERIC_FUNC = False
If IS_LETTER_FUNC(TEMP_VAL, "_") = True Then: Exit Function

If DECIM_SEPAR_CHR = "." Then
'the decimal separator is the period (.)
    
    If InStr(1, TEMP_VAL, ",") > 0 Then: Exit Function
    'Comma is not allowed as decimal separator.
    
    If InStr(1, TEMP_VAL, "d", 1) > 0 Then: Exit Function
    
    i = InStr(1, TEMP_VAL, DECIM_SEPAR_CHR)
    If i > 0 Then If InStr(i + 1, TEMP_VAL, DECIM_SEPAR_CHR) > 0 _
    Then Exit Function
    'Too many periods.
    
    If i = 1 And Len(TEMP_VAL) > 1 Then TEMP_VAL = "0" & TEMP_VAL _
    'Let ".25" = "0.25"
    
    IS_NUMERIC_FUNC = IsNumeric(TEMP_VAL)
Else
'the decimal separator is the comma (,)
    If InStr(1, TEMP_VAL, ".") > 0 Then Exit Function _
    'point is not allowed as decimal separator.
    
    If InStr(1, TEMP_VAL, "d", 1) > 0 Then Exit Function
    i = InStr(1, TEMP_VAL, DECIM_SEPAR_CHR)
    If i > 0 Then If InStr(i + 1, TEMP_VAL, DECIM_SEPAR_CHR) > 0 _
    Then Exit Function
    'Too many commas.
    
    If i = 1 And Len(TEMP_VAL) > 1 Then TEMP_VAL = "0" & TEMP_VAL _
    'Let ",25" = "0,25"
    IS_NUMERIC_FUNC = IsNumeric(TEMP_VAL)
End If

Exit Function
ERROR_LABEL:
IS_NUMERIC_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : STRING_VAR_CHECK_FUNC
'DESCRIPTION   : Validate data types. Each option will perform a
'primitive datatype validation and return a boolean value
 
'The variant data type is used for all parameters as
'the form sets the value property for a text/combo box
'to NULL if the user does not supply data.  Variant is
'the only data type that stores the NULL value.

'LIBRARY       : STRING
'GROUP         : VALIDATION
'ID            : 006
'LAST UPDATE   : 12/02/2008
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'************************************************************************************
'************************************************************************************
 
Function STRING_VAR_CHECK_FUNC(ByVal DATA_VAR As Variant, _
Optional ByVal VERSION As Integer = 0)

'Validate a required integer field
'Criteria: cannot be null, must be a whole number
 
Dim TEMP_VAR As Variant

On Error GoTo ERROR_LABEL

STRING_VAR_CHECK_FUNC = False

TEMP_VAR = DATA_VAR
Select Case VERSION
   '------------------------------------------------------------------------
Case 0 'Validate a required integer field
     If IsNumeric(TEMP_VAR) Then
          If CDbl(TEMP_VAR) = CLng(TEMP_VAR) Then: STRING_VAR_CHECK_FUNC = True
     End If
   '------------------------------------------------------------------------
Case 1
'    Validate a required numeric field
'    Criteria: cannot be null, must be a number
 
     If IsNumeric(TEMP_VAR) Then: STRING_VAR_CHECK_FUNC = True
   '------------------------------------------------------------------------
Case 2
'    Validate a required date field
'    Criteria: cannot be null, must be a date
 
     If Not IsNumeric(TEMP_VAR) Then
          If IsDate(TEMP_VAR) Then: STRING_VAR_CHECK_FUNC = True
     End If
   '------------------------------------------------------------------------
Case 3
'    Validate a required text field
'    Criteria: cannot be null, cannot be a date, cannot be a number
     
     If Not IsNull(TEMP_VAR) Then
          If Not IsDate(TEMP_VAR) Then
               If Not IsNumeric(TEMP_VAR) Then: STRING_VAR_CHECK_FUNC = True
          End If
     End If
   '------------------------------------------------------------------------
Case 4 '--> PERFECT
'    Validate a required text field
'    Criteria: cannot be null, cannot be a date
 
     If Not IsNull(TEMP_VAR) Then
          If Not IsDate(TEMP_VAR) Then: STRING_VAR_CHECK_FUNC = True
     End If
   '------------------------------------------------------------------------
Case 5
'    Validate a not required integer field
'    Criteria: must be a whole number or null
     
     If IsNull(TEMP_VAR) Then
          STRING_VAR_CHECK_FUNC = True
     Else
          If IsNumeric(TEMP_VAR) Then
               If CDbl(TEMP_VAR) = CLng(TEMP_VAR) Then: STRING_VAR_CHECK_FUNC = True
          End If
     End If
   '------------------------------------------------------------------------
Case 6
'    Validate a not required numeric field
'    Criteria: must be a number or null
 
     If IsNull(TEMP_VAR) Then
          STRING_VAR_CHECK_FUNC = True
     Else
          If IsNumeric(TEMP_VAR) Then: STRING_VAR_CHECK_FUNC = True
     End If

  '------------------------------------------------------------------------
Case 7
'    Validate a not required date field
'    Criteria: must be a date or null

     If IsNull(TEMP_VAR) Then
          STRING_VAR_CHECK_FUNC = True
     Else
          If IsDate(TEMP_VAR) Then: STRING_VAR_CHECK_FUNC = True
     End If
 
 '------------------------------------------------------------------------
Case 8
 'Validate a not required text field
 '    Criteria: can be null or any other value except a date or a number

     If IsNull(TEMP_VAR) Then
          STRING_VAR_CHECK_FUNC = True
     Else
          If Not IsDate(TEMP_VAR) Then
               If Not IsNumeric(TEMP_VAR) Then: STRING_VAR_CHECK_FUNC = True
          End If
     End If
'------------------------------------------------------------------------
Case Else
'    Assume invalid data
     If IsNull(TEMP_VAR) Then
          STRING_VAR_CHECK_FUNC = True
     Else
          If Not IsDate(TEMP_VAR) Then: STRING_VAR_CHECK_FUNC = True
     End If
End Select

Exit Function
ERROR_LABEL:
STRING_VAR_CHECK_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : STRING_VAR_VALID_FUNC

'DESCRIPTION   : This test the VarType of Var to see if it is valid to be used
' as a registry key value. Note that all numeric data types (Singles,
' Doubles, etc) are considered value, even though their values will
' be changed when converted to longs.

'LIBRARY       : STRING
'GROUP         : VALID
'ID            : 007
'LAST UPDATE   : 12/02/2008
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'************************************************************************************
'************************************************************************************

Function STRING_VAR_VALID_FUNC(ByVal DATA_VAR As Variant)

Dim TEMP_VAR As Variant

On Error GoTo ERROR_LABEL

TEMP_VAR = DATA_VAR
If VarType(TEMP_VAR) >= vbArray Then
    STRING_VAR_VALID_FUNC = False
    Exit Function
End If
If IsArray(TEMP_VAR) = True Then
    STRING_VAR_VALID_FUNC = False
    Exit Function
End If
If IsObject(TEMP_VAR) = True Then
    STRING_VAR_VALID_FUNC = False
    Exit Function
End If

Select Case VarType(TEMP_VAR)
    Case vbBoolean, vbByte, vbCurrency, vbDate, vbDouble, vbInteger, _
        vbLong, vbSingle, vbString
        STRING_VAR_VALID_FUNC = True
    Case Else
        STRING_VAR_VALID_FUNC = False
End Select

Exit Function
ERROR_LABEL:
STRING_VAR_VALID_FUNC = False
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : VAR_IS_ARRAY_FUNC
'DESCRIPTION   : Returns a Boolean value indicating whether the source
'is an array.
'LIBRARY       : STRING
'GROUP         : VALIDATION
'ID            : 008
'LAST UPDATE   : 12/02/2008
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'************************************************************************************
'************************************************************************************

Function VAR_IS_ARRAY_FUNC(ByRef DATA_RNG As Variant)

Dim DATA_ARR As Variant

On Error GoTo ERROR_LABEL

VAR_IS_ARRAY_FUNC = False
DATA_ARR = DATA_RNG

If VarType(DATA_ARR) >= vbArray Then
    VAR_IS_ARRAY_FUNC = True
ElseIf (VarType(DATA_ARR) = vbUserDefinedType) Or _
       (VarType(DATA_ARR) = vbObject) Then
    VAR_IS_ARRAY_FUNC = False
End If

Exit Function
ERROR_LABEL:
VAR_IS_ARRAY_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : VAR_IS_RANGE_FUNC
'DESCRIPTION   : Returns a Boolean value indicating whether the source is a range.
'LIBRARY       : STRING
'GROUP         : VALIDATION
'ID            : 009
'LAST UPDATE   : 12/02/2008
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'************************************************************************************
'************************************************************************************

Function VAR_IS_RANGE_FUNC(ByVal SRC_RNG As Variant) As Boolean
On Error GoTo ERROR_LABEL
VAR_IS_RANGE_FUNC = False
    If IsObject(SRC_RNG) = True Then: VAR_IS_RANGE_FUNC = True
Exit Function
ERROR_LABEL:
VAR_IS_RANGE_FUNC = False
End Function

