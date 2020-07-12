Attribute VB_Name = "WEB_NUMBER_ROUND_LIBR"



Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.

'************************************************************************************
'************************************************************************************
'FUNCTION      : ROUND_FUNC
'DESCRIPTION   : Rounds a number x taking only d significant digits
'LIBRARY       : NUMBERS
'GROUP         : ROUNDING
'ID            : 001
'LAST UPDATE   : 12/02/2008
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'************************************************************************************
'************************************************************************************

Function ROUND_FUNC(ByVal DATA_VAL As Double, _
Optional ByVal DECIMALS As Integer = 2, _
Optional ByVal VERSION As Integer = 0)

Dim k As Long
Dim TEMP_VAL As Double
Dim Y_VAL As Double
Dim X_VAL As Double

Dim epsilon As Double

On Error GoTo ERROR_LABEL

epsilon = 2 * 10 ^ -16

X_VAL = DATA_VAL
If Abs(X_VAL) <= epsilon Then 'If x = 0 Then
    ROUND_FUNC = 0
Else
    Select Case VERSION
    Case 0
        k = Int(Log(Abs(X_VAL)) / Log(10#)) + 1
        Y_VAL = X_VAL / 10 ^ k
        Y_VAL = Round(Y_VAL, DECIMALS)
        ROUND_FUNC = Y_VAL * 10 ^ k
    Case 1
        TEMP_VAL = DECIMALS - Int(Log(Abs(X_VAL)) / Log(10)) - 1
        If TEMP_VAL < 0 Then: TEMP_VAL = 0
        ROUND_FUNC = Round(X_VAL, TEMP_VAL)
    Case 2
        ROUND_FUNC = Int(X_VAL * 100 + 0.5) / 100
    Case 3
        ROUND_FUNC = CCur(Int(X_VAL * 100 + 0.5) / 100)
    Case Else
        ROUND_FUNC = Int(X_VAL * (10 ^ DECIMALS) + 0.5) / (10 ^ DECIMALS)
    End Select
End If

Exit Function
ERROR_LABEL:
ROUND_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : ASYM_DOWN_FUNC
'DESCRIPTION   : Asymmetrically rounds numbers down - similar to Int().
'Negative numbers get more negative.
'LIBRARY       : NUMBERS
'GROUP         : ROUNDING
'ID            : 002
'LAST UPDATE   : 12/02/2008
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA

'************************************************************************************
'************************************************************************************

Function ASYM_DOWN_FUNC(ByVal DATA_VAL As Double, _
Optional ByVal FACTOR_VAL As Double = 1)

On Error GoTo ERROR_LABEL

ASYM_DOWN_FUNC = Int(DATA_VAL * FACTOR_VAL) / FACTOR_VAL

Exit Function
ERROR_LABEL:
ASYM_DOWN_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : SYM_DOWN_FUNC
'DESCRIPTION   : Symmetrically rounds numbers down - similar to Fix().
'Truncates all numbers toward 0. Same as ASYM_DOWN_FUNC for positive numbers.
'LIBRARY       : NUMBERS
'GROUP         : ROUNDING
'ID            : 003
'LAST UPDATE   : 12/02/2008
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA

'************************************************************************************
'************************************************************************************

Function SYM_DOWN_FUNC(ByVal DATA_VAL As Double, _
Optional ByVal FACTOR_VAL As Double = 1)

On Error GoTo ERROR_LABEL

SYM_DOWN_FUNC = Fix(DATA_VAL * FACTOR_VAL) / FACTOR_VAL
   
'  Alternately:
'  SYM_DOWN_FUNC = ASYM_DOWN_FUNC(Abs(DATA_VAL), FACTOR_VAL) * Sgn(DATA_VAL)

Exit Function
ERROR_LABEL:
SYM_DOWN_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : ASYMUP_FUNC
'DESCRIPTION   : Asymmetrically rounds numbers fractions up.
'Same as SYM_DOWN_FUNC for negative numbers. Similar to Ceiling.
'LIBRARY       : NUMBERS
'GROUP         : ROUNDING
'ID            : 004
'LAST UPDATE   : 12/02/2008
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA

'************************************************************************************
'************************************************************************************

Function ASYMUP_FUNC(ByVal DATA_VAL As Double, _
Optional ByVal FACTOR_VAL As Double = 1)

Dim TEMP_VAL As Double
On Error GoTo ERROR_LABEL

TEMP_VAL = Int(DATA_VAL * FACTOR_VAL)
ASYMUP_FUNC = (TEMP_VAL + IIf(DATA_VAL = TEMP_VAL, 0, 1)) / FACTOR_VAL

Exit Function
ERROR_LABEL:
ASYMUP_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : SYMUP_FUNC (SAME AS ROUNDUP)
'DESCRIPTION   : Symmetrically rounds fractions up - that is, away from 0.
'Same as ASYMUP_FUNC for positive numbers.
'Same as ASYM_DOWN_FUNC for negative numbers.

'LIBRARY       : NUMBERS
'GROUP         : ROUNDING
'ID            : 005
'LAST UPDATE   : 12/02/2008
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'************************************************************************************
'************************************************************************************

Function SYMUP_FUNC(ByVal DATA_VAL As Double, _
Optional ByVal FACTOR_VAL As Double = 1)

Dim TEMP_VAL As Double

On Error GoTo ERROR_LABEL

TEMP_VAL = Fix(DATA_VAL * FACTOR_VAL)
SYMUP_FUNC = (TEMP_VAL + IIf(DATA_VAL = TEMP_VAL, 0, Sgn(DATA_VAL))) / FACTOR_VAL

Exit Function
ERROR_LABEL:
SYMUP_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : ASYM_ARITH_FUNC
'DESCRIPTION   : Asymmetric arithmetic rounding - rounds .5 up always.
'Similar to Java worksheet Round function.
'LIBRARY       : NUMBERS
'GROUP         : ROUNDING
'ID            : 006
'LAST UPDATE   : 12/02/2008
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'************************************************************************************
'************************************************************************************

Function ASYM_ARITH_FUNC(ByVal DATA_VAL As Double, _
Optional ByVal FACTOR_VAL As Double = 1)

On Error GoTo ERROR_LABEL

ASYM_ARITH_FUNC = Int(DATA_VAL * FACTOR_VAL + 0.5) / FACTOR_VAL

Exit Function
ERROR_LABEL:
ASYM_ARITH_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : SYM_ARITH_FUNC
'DESCRIPTION   : Symmetric arithmetic rounding - rounds .5 away from 0.
'Same as ASYM_ARITH_FUNC for positive numbers.
'Similar to Excel Worksheet Round function.
'LIBRARY       : NUMBERS
'GROUP         : ROUNDING
'ID            : 007
'LAST UPDATE   : 12/02/2008
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'************************************************************************************
'************************************************************************************

Function SYM_ARITH_FUNC(ByVal DATA_VAL As Double, _
Optional ByVal FACTOR_VAL As Double = 1)

On Error GoTo ERROR_LABEL

SYM_ARITH_FUNC = Fix(DATA_VAL * FACTOR_VAL + 0.5 * Sgn(DATA_VAL)) / FACTOR_VAL
'  Alternately:
'  SYM_ARITH_FUNC = Abs(ASYM_ARITH_FUNC(DATA_VAL, FACTOR_VAL)) * Sgn(DATA_VAL)
Exit Function
ERROR_LABEL:
SYM_ARITH_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : BROUND_FUNC
'DESCRIPTION   : Rounds .5 up or down to achieve an even number.
'Symmetrical by definition.
'LIBRARY       : NUMBERS
'GROUP         : ROUNDING
'ID            : 008
'LAST UPDATE   : 12/02/2008
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'************************************************************************************
'************************************************************************************

Function BROUND_FUNC(ByVal DATA_VAL As Double, _
Optional ByVal FACTOR_VAL As Double = 1)
   
'  For smaller numbers:
'  BROUND_FUNC = CLng(DATA_VAL * FACTOR_VAL) / FACTOR_VAL

Dim TEMP_VAL As Double
Dim FIX_VAL As Double

On Error GoTo ERROR_LABEL
     
TEMP_VAL = DATA_VAL * FACTOR_VAL
FIX_VAL = Fix(TEMP_VAL + 0.5 * Sgn(DATA_VAL))
' Handle rounding of .5 in a special manner
If TEMP_VAL - Int(TEMP_VAL) = 0.5 Then
    If FIX_VAL / 2 <> Int(FIX_VAL / 2) Then ' Is TEMP_VAL odd
    ' Reduce Magnitude by 1 to make even
        FIX_VAL = FIX_VAL - Sgn(DATA_VAL)
    End If
End If
BROUND_FUNC = FIX_VAL / FACTOR_VAL
Exit Function
ERROR_LABEL:
BROUND_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : RND_ROUND_FUNC
'DESCRIPTION   : Random rounding. Rounds .5 up or down in a random fashion.
'LIBRARY       : NUMBERS
'GROUP         : ROUNDING
'ID            : 009
'LAST UPDATE   : 12/02/2008
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA

'************************************************************************************
'************************************************************************************

Function RND_ROUND_FUNC(ByVal DATA_VAL As Double, _
Optional ByVal FACTOR_VAL As Double = 1)
   ' Should Execute Randomize statement somewhere prior to calling.
Dim TEMP_VAL As Double
Dim FIX_VAL As Double

On Error GoTo ERROR_LABEL

TEMP_VAL = DATA_VAL * FACTOR_VAL
FIX_VAL = Fix(TEMP_VAL + 0.5 * Sgn(DATA_VAL))
' Handle rounding of .5 in a special manner.
If TEMP_VAL - Int(TEMP_VAL) = 0.5 Then
  ' Reduce Magnitude by 1 in half the cases.
  FIX_VAL = FIX_VAL - Int(Rnd * 2) * Sgn(DATA_VAL)
End If
RND_ROUND_FUNC = FIX_VAL / FACTOR_VAL

Exit Function
ERROR_LABEL:
RND_ROUND_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : ALT_ROUND_FUNC
'DESCRIPTION   : Alternates between rounding .5 up or down.
'LIBRARY       : NUMBERS
'GROUP         : ROUNDING
'ID            : 010
'LAST UPDATE   : 12/02/2008
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'************************************************************************************
'************************************************************************************

Function ALT_ROUND_FUNC(ByVal DATA_VAL As Double, _
Optional ByVal FACTOR_VAL As Double = 1)

Static RED_FLAG As Boolean
Dim TEMP_VAL As Double
Dim FIX_VAL As Double

On Error GoTo ERROR_LABEL
     
TEMP_VAL = DATA_VAL * FACTOR_VAL
FIX_VAL = Fix(TEMP_VAL + 0.5 * Sgn(DATA_VAL))
' Handle rounding of .5 in a special manner.
If TEMP_VAL - Int(TEMP_VAL) = 0.5 Then
    ' Alternate between rounding .5 down (negative) and up (positive).
    If (RED_FLAG And Sgn(DATA_VAL) = 1) Or (Not RED_FLAG And Sgn(DATA_VAL) = -1) Then
    ' Or, Replace the previous If statement with the following to
    ' alternate between rounding .5 to reduce magnitude and increase
    ' magnitude.
    ' If RED_FLAG Then
        FIX_VAL = FIX_VAL - Sgn(DATA_VAL)
    End If
    RED_FLAG = Not RED_FLAG
End If
ALT_ROUND_FUNC = FIX_VAL / FACTOR_VAL

Exit Function
ERROR_LABEL:
ALT_ROUND_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : ADOWN_DIG_FUNC

'DESCRIPTION   : With the exception of Excel's MRound() worksheet function,
'the built- in rounding functions take arguments in the manner of
'ADOWN_DIG_FUNC, where the second argument specifies the number of
'digits instead of a factor.

'The rounding implementations presented here use a factor, like
'MRound(), which is more flexible because you do not have to round
'to a power of 10. You can write wrapper functions in the manner of
'ADOWN_DIG_FUNC.

'All of the rounding implementations presented here use the double data
'type, which can represent approximately 15 decimal digits.

'Since not all fractional values can be expressed exactly, you might get
'unexpected results because the display value does not match the stored value.

'For example, the number 2.25 might be stored internally as 2.2499999...,
'which would round down with arithmetic rounding, instead of up as you
'might expect. Also, the more calculations a number is put through, the
'greater possibility that the stored binary value will deviate from the
'ideal decimal value.

'If this is the case, you may want to choose a different data type, such
'as Currency, which is exact to 4 decimal places.

'LIBRARY       : NUMBERS
'GROUP         : ROUNDING
'ID            : 011
'LAST UPDATE   : 12/02/2008
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA

'************************************************************************************
'************************************************************************************

Function ADOWN_DIG_FUNC(ByVal DATA_VAL As Double, _
Optional ByVal DECIMALS As Integer = 0)

On Error GoTo ERROR_LABEL

ADOWN_DIG_FUNC = ASYM_DOWN_FUNC(DATA_VAL, 10 ^ DECIMALS)

Exit Function
ERROR_LABEL:
ADOWN_DIG_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : ROUND_2CB_FUNC

'DESCRIPTION   : You might also consider making the data types Variant
'and use CDec() to convert everything to the Decimal data type, which
'can be exact to 28 decimal digits.

'When you use the Currency data type, which is exact to 4 decimal digits,
'you typically want to round to 2 decimal digits for cents.

'The ROUND_2CB_FUNC function below is a hard-coded variation that performs
'banker's rounding to 2 decimal digits, but does not multiply the
'original number. This avoids a possible overflow condition if the
'monetary amount is approaching the limits of the Currency data type.

'LIBRARY       : NUMBERS
'GROUP         : ROUNDING
'ID            : 012
'LAST UPDATE   : 12/02/2008
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA

'************************************************************************************
'************************************************************************************

Function ROUND_2CB_FUNC(ByVal DATA_VAL As Currency)
On Error GoTo ERROR_LABEL
ROUND_2CB_FUNC = CCur(DATA_VAL / 100) * 100
Exit Function
ERROR_LABEL:
ROUND_2CB_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : ASYM_ARITH_DEC_FUNC
'DESCRIPTION   : The following is an example of asymmetric arithmetic
'rounding using the Decimal data type
'LIBRARY       : NUMBERS
'GROUP         : ROUNDING
'ID            : 013
'LAST UPDATE   : 12/02/2008
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA

'************************************************************************************
'************************************************************************************

Function ASYM_ARITH_DEC_FUNC(ByVal DATA_VAL As Double, _
Optional ByVal FACTOR_VAL As Double = 1)

On Error GoTo ERROR_LABEL
     
     If Not IsNumeric(DATA_VAL) Then
       ASYM_ARITH_DEC_FUNC = DATA_VAL
     Else
       If Not IsNumeric(FACTOR_VAL) Then FACTOR_VAL = 1
       ASYM_ARITH_DEC_FUNC = Int(CDec(DATA_VAL * FACTOR_VAL) + 0.5)
     End If

Exit Function
ERROR_LABEL:
ASYM_ARITH_DEC_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : ROUND_TABLE_FUNC
'DESCRIPTION   : Run Test of Rounding Functions
'LIBRARY       : NUMBERS
'GROUP         : ROUNDING
'ID            : 014
'LAST UPDATE   : 12/02/2008
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA

'************************************************************************************
'************************************************************************************

Function ROUND_TABLE_FUNC(ByVal DATA_VAL As Double, _
Optional ByVal DECIMALS As Integer = 2)

Dim FACTOR_VAL As Double
Dim TEMP_VECTOR(1 To 17, 1 To 2)
On Error GoTo ERROR_LABEL
     
FACTOR_VAL = 10 ^ DECIMALS
TEMP_VECTOR(1, 1) = "ROUND_FUNC_0"
TEMP_VECTOR(1, 2) = ROUND_FUNC(DATA_VAL, DECIMALS, 0)

TEMP_VECTOR(2, 1) = "ROUND_FUNC_1"
TEMP_VECTOR(2, 2) = ROUND_FUNC(DATA_VAL, DECIMALS, 1)

TEMP_VECTOR(3, 1) = "ROUND_FUNC_2"
TEMP_VECTOR(3, 2) = ROUND_FUNC(DATA_VAL, DECIMALS, 2)

TEMP_VECTOR(4, 1) = "ROUND_FUNC_3"
TEMP_VECTOR(4, 2) = ROUND_FUNC(DATA_VAL, DECIMALS, 3)

TEMP_VECTOR(5, 1) = "ROUND_FUNC_4"
TEMP_VECTOR(5, 2) = ROUND_FUNC(DATA_VAL, DECIMALS, 4)

TEMP_VECTOR(6, 1) = "ASYM_DOWN_FUNC"
TEMP_VECTOR(6, 2) = ASYM_DOWN_FUNC(DATA_VAL, FACTOR_VAL)

TEMP_VECTOR(7, 1) = "SYM_DOWN_FUNC"
TEMP_VECTOR(7, 2) = SYM_DOWN_FUNC(DATA_VAL, FACTOR_VAL)

TEMP_VECTOR(8, 1) = "ASYMUP_FUNC"
TEMP_VECTOR(8, 2) = ASYMUP_FUNC(DATA_VAL, FACTOR_VAL)

TEMP_VECTOR(9, 1) = "SYMUP_FUNC"
TEMP_VECTOR(9, 2) = SYMUP_FUNC(DATA_VAL, FACTOR_VAL)

TEMP_VECTOR(10, 1) = "ASYM_ARITH_FUNC"
TEMP_VECTOR(10, 2) = ASYM_ARITH_FUNC(DATA_VAL, FACTOR_VAL)

TEMP_VECTOR(11, 1) = "SYM_ARITH_FUNC"
TEMP_VECTOR(11, 2) = SYM_ARITH_FUNC(DATA_VAL, FACTOR_VAL)

TEMP_VECTOR(12, 1) = "BROUND_FUNC"
TEMP_VECTOR(12, 2) = BROUND_FUNC(DATA_VAL, FACTOR_VAL)

TEMP_VECTOR(13, 1) = "RND_ROUND_FUNC"
TEMP_VECTOR(13, 2) = RND_ROUND_FUNC(DATA_VAL, FACTOR_VAL)

TEMP_VECTOR(14, 1) = "ALT_ROUND_FUNC"
TEMP_VECTOR(14, 2) = ALT_ROUND_FUNC(DATA_VAL, FACTOR_VAL)

TEMP_VECTOR(15, 1) = "ADOWN_DIG_FUNC"
TEMP_VECTOR(15, 2) = ADOWN_DIG_FUNC(DATA_VAL, DECIMALS)

TEMP_VECTOR(16, 1) = "ROUND_2CB_FUNC"
TEMP_VECTOR(16, 2) = ROUND_2CB_FUNC(CCur(DATA_VAL))

TEMP_VECTOR(17, 1) = "ASYM_ARITH_DEC_FUNC"
TEMP_VECTOR(17, 2) = ASYM_ARITH_DEC_FUNC(DATA_VAL, FACTOR_VAL) / FACTOR_VAL

ROUND_TABLE_FUNC = TEMP_VECTOR
Exit Function
ERROR_LABEL:
ROUND_TABLE_FUNC = Err.number
End Function
