Attribute VB_Name = "FINAN_FI_BOND_FORWARD_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.

Function FORWARD_RATE_FUNC(ByVal END_DATE As Date, _
ByVal END_RATE As Double, _
ByVal MID_DATE As Date, _
ByVal MID_RATE As Double, _
ByVal START_DATE As Date)
  
On Error GoTo ERROR_LABEL
  
If END_DATE <> MID_DATE Then
  FORWARD_RATE_FUNC = (END_RATE * (END_DATE - START_DATE) - MID_RATE * (MID_DATE - START_DATE)) / (END_DATE - MID_DATE)
Else
  FORWARD_RATE_FUNC = END_RATE
End If

Exit Function
ERROR_LABEL:
FORWARD_RATE_FUNC = Err.number
End Function

Function FORWARD_FORWARD_VOLATILITY_FUNC(ByVal END_DATE As Date, _
ByVal END_VOLATILITY As Double, _
ByVal MID_DATE As Date, _
ByVal MID_VOLATILITY As Double, _
ByVal START_DATE As Date)

On Error GoTo ERROR_LABEL

If END_DATE > MID_DATE Then
  FORWARD_FORWARD_VOLATILITY_FUNC = Sqr(MAXIMUM_FUNC((END_VOLATILITY ^ 2 * (END_DATE - START_DATE) - MID_VOLATILITY ^ 2 * (MID_DATE - START_DATE)) / (END_DATE - MID_DATE), 0))
Else
  FORWARD_FORWARD_VOLATILITY_FUNC = END_VOLATILITY
End If

Exit Function
ERROR_LABEL:
FORWARD_FORWARD_VOLATILITY_FUNC = Err.number
End Function

Function FUTURES_INTERPOLATION_FUNC(ByVal X1_VAL As Variant, _
ByVal X2_VAL As Variant, _
ByVal Y1_VAL As Double, _
ByVal Y2_VAL As Double, _
ByVal X0_VAL As Variant, _
ByVal Y0_VAL As Double) As Double

Dim CARRY(0 To 2) As Double

On Error GoTo ERROR_LABEL

CARRY(2) = 1 / X2_VAL * Log(Y2_VAL / Y0_VAL)
If X1_VAL > 0 Then
    CARRY(1) = 1 / X1_VAL * Log(Y1_VAL / Y0_VAL)
    CARRY(0) = INTERPOLATION_FUNC(X1_VAL, X2_VAL, CARRY(1), CARRY(2), X0_VAL, "lin", 1)
Else
    CARRY(0) = CARRY(2)
End If
FUTURES_INTERPOLATION_FUNC = Y0_VAL * Exp(CARRY(0) * X0_VAL)

Exit Function
ERROR_LABEL:
FUTURES_INTERPOLATION_FUNC = Err.number
End Function

Function FORWARD_INTERPOLATION_FUNC(ByVal X1_VAL As Variant, _
ByVal X2_VAL As Variant, _
ByVal Y1_VAL As Double, _
ByVal Y2_VAL As Double, _
ByVal X0_VAL As Variant, _
Optional ByVal METHOD_VAL As Variant = "lin", _
Optional ByVal POWER_VAL As Double = 1) As Double

On Error GoTo ERROR_LABEL

If X0_VAL < X1_VAL Then
    FORWARD_INTERPOLATION_FUNC = FORWARD_INTERPOLATION_FUNC(X1_VAL, X2_VAL, Y1_VAL, Y2_VAL, X1_VAL, METHOD_VAL, POWER_VAL)
ElseIf X0_VAL > X2_VAL Then
    FORWARD_INTERPOLATION_FUNC = FORWARD_INTERPOLATION_FUNC(X1_VAL, X2_VAL, Y1_VAL, Y2_VAL, X2_VAL, METHOD_VAL, POWER_VAL)
Else
    Select Case LCase(Left(METHOD_VAL, 3))
    Case 0, "lin", "lor", ""
        FORWARD_INTERPOLATION_FUNC = (2 * X0_VAL - X1_VAL) / (X2_VAL - X1_VAL) * Y2_VAL + (X2_VAL - 2 * X0_VAL) / (X2_VAL - X1_VAL) * Y1_VAL
    Case 1, "log", "llr"
        FORWARD_INTERPOLATION_FUNC = INTERPOLATION_FUNC(X1_VAL, X2_VAL, Y1_VAL, Y2_VAL, X0_VAL, "lin", 1) * (X0_VAL / (X2_VAL - X1_VAL) * Log(Y2_VAL / Y1_VAL) + 1)
    Case 2, "lod"
        FORWARD_INTERPOLATION_FUNC = (Exp(-X1_VAL * Y1_VAL) - Exp(-X2_VAL * Y2_VAL)) / ((X0_VAL - X1_VAL) * Exp(-X2_VAL * Y2_VAL) + (X2_VAL - X0_VAL) * Exp(-X1_VAL * Y1_VAL))
    Case 3, "wei", "raw", "lld"
        FORWARD_INTERPOLATION_FUNC = ((Y2_VAL ^ POWER_VAL * X2_VAL - Y1_VAL ^ POWER_VAL * X1_VAL) / (X2_VAL - X1_VAL)) ^ (1 / POWER_VAL)
    Case Else
        FORWARD_INTERPOLATION_FUNC = FORWARD_INTERPOLATION_FUNC(X1_VAL, X2_VAL, Y1_VAL, Y2_VAL, X0_VAL)
    End Select
End If

Exit Function
ERROR_LABEL:
FORWARD_INTERPOLATION_FUNC = Err.number
End Function
