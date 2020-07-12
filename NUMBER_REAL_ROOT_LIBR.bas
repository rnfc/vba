Attribute VB_Name = "NUMBER_REAL_ROOT_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.


'************************************************************************************
'************************************************************************************
'FUNCTION      : PYTHAG_FUNC
'DESCRIPTION   : Pythagora function
'LIBRARY       : NUMBER_REAL
'GROUP         : ROOT
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/08/2008
'************************************************************************************
'************************************************************************************

Function PYTHAG_FUNC(ByVal A_VAL As Double, _
ByVal B_VAL As Double)
  
On Error GoTo ERROR_LABEL
  
If Abs(A_VAL) > Abs(B_VAL) Then
    PYTHAG_FUNC = Abs(A_VAL) * Sqr(1 + (Abs(B_VAL) / Abs(A_VAL)) ^ 2)
ElseIf Abs(B_VAL) = 0 Then
    PYTHAG_FUNC = 0
Else
    PYTHAG_FUNC = Abs(B_VAL) * Sqr(1 + (Abs(A_VAL) / Abs(B_VAL)) ^ 2)
End If

Exit Function
ERROR_LABEL:
PYTHAG_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : ROOT_FUNC
'DESCRIPTION   : Algebric extension of root for a<0
'LIBRARY       : NUMBER_REAL
'GROUP         : ROOT
'ID            : 002
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/08/2008
'************************************************************************************
'************************************************************************************


Function ROOT_FUNC(ByVal X_VAL As Double, _
Optional ByVal NROWS As Long = 2)
    
On Error GoTo ERROR_LABEL
    
If NROWS = 0 Then
    ROOT_FUNC = "" 'raise an error
ElseIf MOD_FUNC(NROWS, 2) = 0 Then 'NROWS is even => _
root in a<0 doesn´t exist
    If X_VAL < 0 Then
        ROOT_FUNC = "" 'raise an error
    Else
        ROOT_FUNC = X_VAL ^ (1 / NROWS)
    End If
Else  'NROWS is odd => root in a<0 exists
    ROOT_FUNC = Sgn(X_VAL) * Abs(X_VAL) ^ (1 / NROWS)
End If
  
Exit Function
ERROR_LABEL:
ROOT_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : SQRT_FUNC
'DESCRIPTION   : Squared-Root Function
'LIBRARY       : REAL_NUMBER
'GROUP         : ROOT
'ID            : 003
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/08/2008
'************************************************************************************
'************************************************************************************

Function SQRT_FUNC(ByVal X_VAL As Double)

On Error GoTo ERROR_LABEL

SQRT_FUNC = X_VAL ^ 0.5

Exit Function
ERROR_LABEL:
SQRT_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : HYPOT_FUNC
'DESCRIPTION   : returns the square root of the sum of the squares of the arguments
'LIBRARY       : REAL_NUMBER
'GROUP         : ROOT
'ID            : 004
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/08/2008
'************************************************************************************
'************************************************************************************
                    
Function HYPOT_FUNC(ByVal A_VAL As Double, _
ByVal B_VAL As Double)
    
Dim C_VAL As Double
  
On Error GoTo ERROR_LABEL
  
If A_VAL = 0 Then
  HYPOT_FUNC = Abs(B_VAL)
Else
  C_VAL = B_VAL / A_VAL
  HYPOT_FUNC = Abs(A_VAL) * (1 + C_VAL * C_VAL) ^ 0.5
End If

Exit Function
ERROR_LABEL:
HYPOT_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : CONTINUED_FRACTION_SQUARE_ROOT_FUNC
'DESCRIPTION   : return the continued fraction of a square root of an integer number
'LIBRARY       : NUMBER
'GROUP         : REAL_ROOT
'ID            : 005
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/27/2009
'************************************************************************************
'************************************************************************************

Function CONTINUED_FRACTION_SQUARE_ROOT_FUNC(ByVal NUMBER_VAL As Variant, _
Optional ByVal nLOOPS As Long = 10000)

Dim i As Long
Dim j As Long

Dim NSIZE As Long
Dim TEMP_ARR As Variant
Dim TEMP_VECTOR As Variant

Dim A_VAL As Double
Dim S_VAL As Double
Dim T_VAL As Double

Dim R_VAL As Double
Dim R1_VAL As Double

Dim U_VAL As Double
Dim U1_VAL As Double

Dim V_VAL As Double
Dim V1_VAL As Double

Dim P_VAL As Double
Dim P1_VAL As Double

On Error GoTo ERROR_LABEL

'continued fraction expansion of a n square-root
'returns the vector q=[q0,q1,q2...qn] representing n^(1/2)
'the lenght-1 of the vector q represents the period of the
'continued fraction

ReDim TEMP_ARR(0 To nLOOPS)
R_VAL = Int(Sqr(NUMBER_VAL))
TEMP_ARR(0) = R_VAL
R1_VAL = 2 * R_VAL
U1_VAL = R1_VAL
V_VAL = NUMBER_VAL - R_VAL ^ 2

If V_VAL = 0 Then  'n is perfect square
    ReDim Preserve TEMP_ARR(0 To 0)
    GoTo 1983
End If

V1_VAL = 1
A_VAL = Int(R1_VAL / V_VAL)
TEMP_ARR(1) = A_VAL
U_VAL = R1_VAL - MODULUS_FUNC(R1_VAL, V_VAL)
P1_VAL = R_VAL
P_VAL = MODULUS_FUNC(A_VAL * R_VAL + 1, NUMBER_VAL)
S_VAL = 1
i = 1
Do
    If TEMP_ARR(i) = 2 * TEMP_ARR(0) Then Exit Do
    i = i + 1
    T_VAL = V_VAL
    V_VAL = A_VAL * (U1_VAL - U_VAL) + V1_VAL
    V1_VAL = T_VAL
    A_VAL = Int(U_VAL / V_VAL)
    TEMP_ARR(i) = A_VAL
    U1_VAL = U_VAL
    U_VAL = R1_VAL - MODULUS_FUNC(U_VAL, V_VAL)
    S_VAL = 1 - S_VAL
    T_VAL = P_VAL
    P_VAL = MODULUS_FUNC(A_VAL * P_VAL + P1_VAL, NUMBER_VAL)
    P1_VAL = T_VAL
Loop Until i = nLOOPS

ReDim Preserve TEMP_ARR(0 To i)

1983:
NSIZE = UBound(TEMP_ARR) - LBound(TEMP_ARR) + 1
ReDim TEMP_VECTOR(1 To NSIZE, 1 To 1)

j = 1
For i = LBound(TEMP_ARR) To UBound(TEMP_ARR)
    TEMP_VECTOR(j, 1) = TEMP_ARR(i)
    j = j + 1
Next i
CONTINUED_FRACTION_SQUARE_ROOT_FUNC = TEMP_VECTOR

Exit Function
ERROR_LABEL:
CONTINUED_FRACTION_SQUARE_ROOT_FUNC = Err.number
End Function
