Attribute VB_Name = "WEB_NUMBER_FRACTION_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.

'************************************************************************************
'************************************************************************************
'FUNCTION      : NUMBER_FRACTION_STRING1_FUNC
'DESCRIPTION   : Convert A, B in string fraction
'LIBRARY       : NUMBERS
'GROUP         : FRACTION
'ID            : 001
'LAST UPDATE   : 12/02/2008
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'************************************************************************************
'************************************************************************************

Function NUMBER_FRACTION_STRING1_FUNC(ByRef NUMERATOR_VAL As Variant, _
ByRef DENOMINATOR_VAL As Variant)
'As Currency
Dim TEMP_STR As String
Dim TEMP_CHR As String

On Error GoTo ERROR_LABEL

TEMP_CHR = ""
If (NUMERATOR_VAL * DENOMINATOR_VAL) < 0 Then: TEMP_CHR = "-"

If Not IS_INTEGER_FUNC(NUMERATOR_VAL) Or Not _
       IS_INTEGER_FUNC(DENOMINATOR_VAL) Then
    TEMP_STR = CStr((NUMERATOR_VAL / DENOMINATOR_VAL))
Else

    If Abs(NUMERATOR_VAL) = 1 And Abs(DENOMINATOR_VAL) = 1 Then
        TEMP_STR = TEMP_CHR & "1"
    ElseIf Abs(NUMERATOR_VAL) = 1 And Abs(DENOMINATOR_VAL) > 1 Then
        TEMP_STR = TEMP_CHR & "1/" & Trim(CStr((Abs(DENOMINATOR_VAL))))
    ElseIf Abs(NUMERATOR_VAL) > 1 And Abs(DENOMINATOR_VAL) = 1 Then
        TEMP_STR = TEMP_CHR & Trim(CStr((Abs(NUMERATOR_VAL))))
    Else
        TEMP_STR = TEMP_CHR & Trim(CStr((Abs(NUMERATOR_VAL)))) & "/" & _
            Trim(CStr((Abs(DENOMINATOR_VAL))))
    End If
End If

NUMBER_FRACTION_STRING1_FUNC = TEMP_STR

Exit Function
ERROR_LABEL:
NUMBER_FRACTION_STRING1_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : NUMBER_FRACTION_STRING2_FUNC
'DESCRIPTION   : Convertit un nombre décimal en une fraction normalisée et
'détermine automatiquement un denominator entre 2 and 8.
'LIBRARY       : NUMBER
'GROUP         : FRACTION
'ID            : 002
'LAST UPDATE   : 12/02/2008
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'************************************************************************************
'************************************************************************************

Function NUMBER_FRACTION_STRING2_FUNC(ByRef DATA_VAL As Double)

Dim TEMP_STR As String
Dim TEMP_FIXED As Double

On Error GoTo ERROR_LABEL

If (VarType(DATA_VAL) < 2) Or (VarType(DATA_VAL) > 6) Then
  NUMBER_FRACTION_STRING2_FUNC = DATA_VAL
Else
  DATA_VAL = Abs(DATA_VAL)
  TEMP_FIXED = Int(DATA_VAL)
  If TEMP_FIXED > 0 Then
    TEMP_STR = CStr(TEMP_FIXED)
  End If
  Select Case DATA_VAL - TEMP_FIXED
    Case Is < 0.1
      If TEMP_FIXED > 0 Then
        TEMP_STR = TEMP_STR
      Else
        TEMP_STR = CStr(DATA_VAL)
      End If
    Case 0.1 To 0.145
      TEMP_STR = TEMP_STR + " 1/8"
    Case 0.145 To 0.182
      TEMP_STR = TEMP_STR + " 1/6"
    Case 0.182 To 0.225
      TEMP_STR = TEMP_STR + " 1/5"
    Case 0.225 To 0.29
      TEMP_STR = TEMP_STR + " 1/4"
    Case 0.29 To 0.35
      TEMP_STR = TEMP_STR + " 1/3"
    Case 0.35 To 0.3875
      TEMP_STR = TEMP_STR + " 3/8"
    Case 0.3875 To 0.45
      TEMP_STR = TEMP_STR + " 2/5"
    Case 0.45 To 0.55
      TEMP_STR = TEMP_STR + " 1/2"
    Case 0.55 To 0.6175
      TEMP_STR = TEMP_STR + " 3/5"
    Case 0.6175 To 0.64
      TEMP_STR = TEMP_STR + " 5/8"
    Case 0.64 To 0.7
      TEMP_STR = TEMP_STR + " 2/3"
    Case 0.7 To 0.775
      TEMP_STR = TEMP_STR + " 3/4"
    Case 0.775 To 0.8375
      TEMP_STR = TEMP_STR + " 4/5"
    Case 0.8735 To 0.91
      TEMP_STR = TEMP_STR + " 7/8"
    Case Is > 0.91
      TEMP_STR = CStr(Int(DATA_VAL) + 1)
  End Select
  NUMBER_FRACTION_STRING2_FUNC = TEMP_STR
End If

Exit Function
ERROR_LABEL:
NUMBER_FRACTION_STRING2_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : CONVERT_FRACTION_STRING_NUMBER_FUNC
'DESCRIPTION   : Analyse une fraction standard de format "a/b" ou "a b/c"
' et retourne un nombre. Par exemple "2/5" ou "3 1/2" sont
' des entrées valides.
'LIBRARY       : NUMBERS
'GROUP         : FRACTION
'ID            : 003
'LAST UPDATE   : 12/02/2008
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'************************************************************************************
'************************************************************************************

Function CONVERT_FRACTION_STRING_NUMBER_FUNC(ByRef DATA_STR As String)

Dim i As Long
Dim TEMP_STR As String

Dim TEMP_VAL As Double
Dim NUMENATOR_VAL As Double
Dim DENOMINATOR_VAL As Double

On Error GoTo ERROR_LABEL
  
  If VarType(DATA_STR) < 2 Or VarType(DATA_STR) = 7 Then
    CONVERT_FRACTION_STRING_NUMBER_FUNC = Null
  ElseIf VarType(DATA_STR) <> 8 Then
    CONVERT_FRACTION_STRING_NUMBER_FUNC = DATA_STR
  Else
    TEMP_STR = Trim$(DATA_STR)
    i = InStr(TEMP_STR, " ")
    If i = 0 Then
      If InStr(TEMP_STR, "/") = 0 Then
        TEMP_VAL = Val(TEMP_STR)
      Else
        TEMP_VAL = 0
      End If
    Else
      TEMP_VAL = Val(Left$(TEMP_STR, i - 1))
      TEMP_STR = Mid$(TEMP_STR, i + 1)
    End If
    i = InStr(TEMP_STR, "/")
    If i <> 0 Then
      NUMENATOR_VAL = Val(Left$(TEMP_STR, i - 1))
      DENOMINATOR_VAL = Val(Mid$(TEMP_STR, i + 1))
      If DENOMINATOR_VAL <> 0 Then
        TEMP_VAL = TEMP_VAL + NUMENATOR_VAL / DENOMINATOR_VAL
      End If
    End If
    CONVERT_FRACTION_STRING_NUMBER_FUNC = TEMP_VAL
  End If

Exit Function
ERROR_LABEL:
CONVERT_FRACTION_STRING_NUMBER_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : CONVERT_NUMBER_FRACTION1_FUNC
'DESCRIPTION   : Real number to fraction expantion
'LIBRARY       : NUMBER
'GROUP         : FRACTION
'ID            : 004
'LAST UPDATE   : 12/02/2008
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'************************************************************************************
'************************************************************************************

Function CONVERT_NUMBER_FRACTION1_FUNC(ByRef DATA_VAL As Double)

Dim i As Long

Dim ATEMP As Double
Dim BTEMP As Double
Dim CTEMP As Double

Dim dTemp As Double
Dim ETEMP As Double
Dim FTEMP As Double

Dim FIRST_VAL As Double
Dim SECOND_VAL As Double
Dim THIRD_VAL As Double

Dim X1_VAL As Double
Dim X2_VAL As Double

Dim Y1_VAL As Double
Dim Y2_VAL As Double

Dim DELTA_VAL As Double
Dim NUMENATOR_VAL As Double
Dim DENOMINATOR_VAL As Double

Dim epsilon As Double

Dim TEMP_VECTOR() As Variant

On Error GoTo ERROR_LABEL

ReDim TEMP_VECTOR(1 To 3, 1 To 1)

If DATA_VAL < 0 Then
    DATA_VAL = Abs(DATA_VAL)
    THIRD_VAL = -1
ElseIf DATA_VAL > 0 Then
    THIRD_VAL = 1
Else
    NUMENATOR_VAL = 0
    DENOMINATOR_VAL = 1
    
    TEMP_VECTOR(1, 1) = 0
    TEMP_VECTOR(2, 1) = NUMENATOR_VAL
    TEMP_VECTOR(3, 1) = DENOMINATOR_VAL
    
    CONVERT_NUMBER_FRACTION1_FUNC = TEMP_VECTOR
    Exit Function
End If

epsilon = 10 ^ -14

X1_VAL = 1
Y1_VAL = GET_DECIMALS_FUNC(DATA_VAL) 'decimal part
FIRST_VAL = Int(DATA_VAL)  'integer part

ATEMP = 0
BTEMP = 1
CTEMP = FIRST_VAL

dTemp = 1
ETEMP = 0
FTEMP = 1

DELTA_VAL = Abs((DATA_VAL - CTEMP / FTEMP) / DATA_VAL)
epsilon = epsilon
i = 0
'begin iterative method of continued fraction expantion
Do Until DELTA_VAL < epsilon
    i = i + 1
    If Y1_VAL = 0 Then Exit Do
    SECOND_VAL = Int(X1_VAL / Y1_VAL)
    X2_VAL = Y1_VAL
    Y2_VAL = X1_VAL - SECOND_VAL * Y1_VAL
'shift variables
    ATEMP = BTEMP
    BTEMP = CTEMP
    dTemp = ETEMP
    ETEMP = FTEMP
    FIRST_VAL = SECOND_VAL
    X1_VAL = X2_VAL
    Y1_VAL = Y2_VAL
    CTEMP = FIRST_VAL * BTEMP + ATEMP
    FTEMP = FIRST_VAL * ETEMP + dTemp
    DELTA_VAL = Abs((DATA_VAL - CTEMP / FTEMP) / DATA_VAL)
Loop

NUMENATOR_VAL = THIRD_VAL * CTEMP
DENOMINATOR_VAL = FTEMP

TEMP_VECTOR(1, 1) = NUMENATOR_VAL & "/" & DENOMINATOR_VAL
TEMP_VECTOR(2, 1) = NUMENATOR_VAL
TEMP_VECTOR(3, 1) = DENOMINATOR_VAL
    
CONVERT_NUMBER_FRACTION1_FUNC = TEMP_VECTOR

Exit Function
ERROR_LABEL:
CONVERT_NUMBER_FRACTION1_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : CONVERT_NUMBER_FRACTION2_FUNC
'DESCRIPTION   : Convertit un nombre décimal en fraction mais ne la
'normalise pas. exemple 3 2/4 -> 3 2/4
'LIBRARY       : NUMBER
'GROUP         : FRACTION
'ID            : 005
'LAST UPDATE   : 12/02/2008
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'************************************************************************************
'************************************************************************************


Function CONVERT_NUMBER_FRACTION2_FUNC(ByRef DATA_VAL As Double, _
ByRef DENOMINATOR_VAL As Double)

Dim TEMP_STR As String
Dim TEMP_FIXED As Double
Dim NUMENATOR_VAL As Double

On Error GoTo ERROR_LABEL

If (VarType(DATA_VAL) < 2) Or (VarType(DATA_VAL) > 6) Then
  CONVERT_NUMBER_FRACTION2_FUNC = DATA_VAL
  Exit Function
End If
DATA_VAL = Abs(DATA_VAL)
TEMP_FIXED = Int(DATA_VAL)
NUMENATOR_VAL = Int((DATA_VAL - TEMP_FIXED) * DENOMINATOR_VAL + 0.5)
'Arrondi arithmétique
If NUMENATOR_VAL = DENOMINATOR_VAL Then
  TEMP_FIXED = TEMP_FIXED + 1
  NUMENATOR_VAL = 0
End If
If TEMP_FIXED > 0 Then
  TEMP_STR = CStr(TEMP_FIXED)
End If
If NUMENATOR_VAL > 0 Then
  TEMP_STR = TEMP_STR & " " & NUMENATOR_VAL & "/" & DENOMINATOR_VAL
End If
CONVERT_NUMBER_FRACTION2_FUNC = TEMP_STR

Exit Function
ERROR_LABEL:
CONVERT_NUMBER_FRACTION2_FUNC = Err.number
End Function


