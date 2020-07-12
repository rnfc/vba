Attribute VB_Name = "NUMBER_REAL_EUCLIDEAN_LIBR"


Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.

'************************************************************************************
'************************************************************************************
'FUNCTION      : ABS_FUNC
'DESCRIPTION   : Returns a value of the same type that is passed to it
'specifying the absolute value of a number.

'LIBRARY       : NUMBER_REAL
'GROUP         : EUCLIDEAN
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/08/2008
'************************************************************************************
'************************************************************************************


Function ABS_FUNC(ByVal X_VAL As Double)

On Error GoTo ERROR_LABEL

'The absolute value of a number is its unsigned magnitude. For example,
'Abs(-1) and Abs(1) both return 1.

If X_VAL < 0 Then
    ABS_FUNC = X_VAL * -1
Else
    ABS_FUNC = X_VAL
End If

Exit Function
ERROR_LABEL:
ABS_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : SIGN_FUNC
'DESCRIPTION   : Returns a Variant (Integer) indicating the sign of a number.
'The sign of the number argument determines the return
'value of the Sgn function.
'LIBRARY       : NUMBER_REAL
'GROUP         : EUCLIDEAN
'ID            : 002
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/08/2008
'************************************************************************************
'************************************************************************************


Function SIGN_FUNC(ByVal X_VAL As Double)
On Error GoTo ERROR_LABEL

If X_VAL < 0 Then
    SIGN_FUNC = -1
ElseIf X_VAL = 0 Then
    SIGN_FUNC = 0
Else 'Greater than zero 1
    SIGN_FUNC = 1
End If

Exit Function
ERROR_LABEL:
SIGN_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : NEGATIVE_THRESHOLD_FUNC
'DESCRIPTION   : Returns the sign of the number if a certain
'threshold is reached (THRESHOLD)
'LIBRARY       : NUMBER_REAL
'GROUP         : EUCLIDEAN
'ID            : 003
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/08/2008
'************************************************************************************
'************************************************************************************

Function NEGATIVE_THRESHOLD_FUNC(ByVal X_VAL As Double, _
Optional ByVal THRESHOLD As Double = 0)

Dim A_VAL As Double
Dim B_VAL As Double

On Error GoTo ERROR_LABEL

If (X_VAL >= 0) Then
    A_VAL = X_VAL
Else
    A_VAL = -X_VAL
End If

If (THRESHOLD >= 0) Then
    B_VAL = A_VAL
Else
    B_VAL = -A_VAL
End If

NEGATIVE_THRESHOLD_FUNC = B_VAL

'SAME AS:
'If THRESHOLD >= 0 Then
'   NEGATIVE_THRESHOLD_FUNC = Abs(TEMP)
'Else
'   NEGATIVE_THRESHOLD_FUNC = -Abs(TEMP)
'End If

Exit Function
ERROR_LABEL:
NEGATIVE_THRESHOLD_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : EUCLIDEAN_FUNC
'DESCRIPTION   : This algorithm handles negative numbers by making
'them positive
'REFERENCE: Chabert, Jean-Luc (ed.), A History of Algorithms: From the
'Pebble to the Microchip - Springer, Berlin, 1999

'LIBRARY       : NUMBER_REAL
'GROUP         : EUCLIDEAN
'ID            : 004
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/08/2008
'************************************************************************************
'************************************************************************************

Function EUCLIDEAN_FUNC(ByVal A_VAL As Double, _
ByVal B_VAL As Double)

Dim U_VAL As Double
Dim V_VAL As Double

On Error GoTo ERROR_LABEL

If A_VAL < 0 Then: A_VAL = -A_VAL
If B_VAL < 0 Then: B_VAL = -B_VAL

U_VAL = A_VAL
V_VAL = B_VAL

If (A_VAL = 0) Or (B_VAL = 0) Then
    EUCLIDEAN_FUNC = 0
    Exit Function
End If

Do While U_VAL <> V_VAL
    If U_VAL > V_VAL Then
        U_VAL = U_VAL - V_VAL
    Else
        V_VAL = V_VAL - U_VAL
    End If
Loop

EUCLIDEAN_FUNC = U_VAL

Exit Function
ERROR_LABEL:
EUCLIDEAN_FUNC = Err.number
End Function
