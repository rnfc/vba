Attribute VB_Name = "STAT_INTERPOLATE_LIBR"
Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 0       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.

Public PUB_ILAST_INDEX_VAL As Integer
Public PUB_LLAST_INDEX_VAL As Integer

Function INTERPOLATION_FUNC(ByVal X1_VAL As Variant, _
ByVal X2_VAL As Variant, _
ByVal Y1_VAL As Double, _
ByVal Y2_VAL As Double, _
ByVal X0_VAL As Variant, _
Optional METHOD_VAL As Variant = "lin", _
Optional POWER_VAL As Double = 1) As Double

On Error GoTo ERROR_LABEL

If X0_VAL <= X1_VAL Then
    INTERPOLATION_FUNC = Y1_VAL
ElseIf X0_VAL >= X2_VAL Then
    INTERPOLATION_FUNC = Y2_VAL
Else
    Select Case LCase(Left(METHOD_VAL, 3))
    Case 0, "lin", "lor", ""
        INTERPOLATION_FUNC = (X0_VAL - X1_VAL) / (X2_VAL - X1_VAL) * Y2_VAL + (X2_VAL - X0_VAL) / (X2_VAL - X1_VAL) * Y1_VAL
    Case 1, "log", "llr"
        INTERPOLATION_FUNC = Y2_VAL ^ ((X0_VAL - X1_VAL) / (X2_VAL - X1_VAL)) * Y1_VAL ^ ((X2_VAL - X0_VAL) / (X2_VAL - X1_VAL))
    Case 2, "lod"
        INTERPOLATION_FUNC = -1 / X0_VAL * Log((X0_VAL - X1_VAL) / (X2_VAL - X1_VAL) * Exp(-X2_VAL * Y2_VAL) + (X2_VAL - X0_VAL) / (X2_VAL - X1_VAL) * Exp(-X1_VAL * Y1_VAL))
    Case 3, "wei", "raw", "lld"
        INTERPOLATION_FUNC = ((X0_VAL - X1_VAL) / (X2_VAL - X1_VAL) * X2_VAL / X0_VAL * Y2_VAL ^ POWER_VAL + (X2_VAL - X0_VAL) / (X2_VAL - X1_VAL) * X1_VAL / X0_VAL * Y1_VAL ^ POWER_VAL) ^ (1 / POWER_VAL)
    Case Else
        INTERPOLATION_FUNC = Y1_VAL
    End Select
End If

Exit Function
ERROR_LABEL:
INTERPOLATION_FUNC = Err.number
End Function


Private Function ARRAY_INTERPOLATION_FUNC(ByRef XDATA_ARR As Variant, _
ByRef YDATA_ARR As Variant, _
ByVal X0_VAL As Variant, _
Optional ByVal METHOD_VAL As Variant = "LIN", _
Optional ByVal POWER_VAL As Double = 1) As Double

'XDATA_ARR MUST BE SORTED IN ASCENDING ORDER
'Debug.Print ARRAY_INTERPOLATION_FUNC(Array(1, 2, 3, 4, 5), Array(2.4, 2.5, 2.6, 2.7, 2.8), 4.5)

Dim k As Integer

On Error GoTo ERROR_LABEL

k = UBound(XDATA_ARR)
PUB_LLAST_INDEX_VAL = LAST_INDEX_ARRAY_INTERPOLATION_FUNC(XDATA_ARR, YDATA_ARR, X0_VAL)
Select Case PUB_LLAST_INDEX_VAL
Case 0
    ARRAY_INTERPOLATION_FUNC = ARRAY_INTERPOLATION_FUNC(XDATA_ARR, YDATA_ARR, XDATA_ARR(1), METHOD_VAL, POWER_VAL)
Case k
    ARRAY_INTERPOLATION_FUNC = ARRAY_INTERPOLATION_FUNC(XDATA_ARR, YDATA_ARR, XDATA_ARR(k), METHOD_VAL, POWER_VAL)
Case Else
    ARRAY_INTERPOLATION_FUNC = INTERPOLATION_FUNC(XDATA_ARR(PUB_LLAST_INDEX_VAL), XDATA_ARR(PUB_LLAST_INDEX_VAL + 1), CDbl(YDATA_ARR(PUB_LLAST_INDEX_VAL)), CDbl(YDATA_ARR(PUB_LLAST_INDEX_VAL + 1)), X0_VAL, METHOD_VAL, POWER_VAL)
End Select

Exit Function
ERROR_LABEL:
ARRAY_INTERPOLATION_FUNC = Err.number
End Function


Private Function LAST_INDEX_ARRAY_INTERPOLATION_FUNC(ByRef XDATA_ARR As Variant, _
ByRef YDATA_ARR As Variant, _
ByVal X0_VAL As Variant) As Integer

Dim k As Integer

On Error GoTo ERROR_LABEL

k = UBound(XDATA_ARR)

Do
    If X0_VAL >= XDATA_ARR(PUB_ILAST_INDEX_VAL) Then
        If X0_VAL > XDATA_ARR(k) Then
            LAST_INDEX_ARRAY_INTERPOLATION_FUNC = k
            Exit Function
        ElseIf X0_VAL <= XDATA_ARR(PUB_ILAST_INDEX_VAL + 1) Then
            LAST_INDEX_ARRAY_INTERPOLATION_FUNC = PUB_ILAST_INDEX_VAL
            Exit Function
        Else
            PUB_ILAST_INDEX_VAL = PUB_ILAST_INDEX_VAL + 1
        End If
    Else
        If X0_VAL < XDATA_ARR(1) Then
            LAST_INDEX_ARRAY_INTERPOLATION_FUNC = 0
            Exit Function
        ElseIf X0_VAL >= XDATA_ARR(PUB_ILAST_INDEX_VAL - 1) Then
            LAST_INDEX_ARRAY_INTERPOLATION_FUNC = PUB_ILAST_INDEX_VAL - 1
            Exit Function
        Else
            PUB_ILAST_INDEX_VAL = PUB_ILAST_INDEX_VAL - 2
        End If
    End If
Loop

Exit Function
ERROR_LABEL:
LAST_INDEX_ARRAY_INTERPOLATION_FUNC = Err.number
End Function
