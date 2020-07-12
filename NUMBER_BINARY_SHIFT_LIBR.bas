Attribute VB_Name = "NUMBER_BINARY_SHIFT_LIBR"


Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.
                    
'************************************************************************************
'************************************************************************************
'FUNCTION      : BINARY_SHIFTER_FUNC
'DESCRIPTION   : Binary shift (+n shift right, -NSIZE shift left)
'LIBRARY       : NUMBER_BINARY
'GROUP         : SHIFT
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 14/02/2008
'************************************************************************************
'************************************************************************************

'// PERFECT

Function BINARY_SHIFTER_FUNC(ByVal BINARY_STR As String, _
Optional ByVal NSIZE As Long = 1, _
Optional ByVal RING_FLAG As Boolean = False)

On Error GoTo ERROR_LABEL

If NSIZE > 0 Then
    BINARY_SHIFTER_FUNC = BINARY_SHIFT_RIGHT_FUNC(BINARY_STR, NSIZE, RING_FLAG)
ElseIf NSIZE < 0 Then
    BINARY_SHIFTER_FUNC = BINARY_SHIFT_LEFT_FUNC(BINARY_STR, Abs(NSIZE), RING_FLAG)
Else
    BINARY_SHIFTER_FUNC = BINARY_STR
End If

Exit Function
ERROR_LABEL:
BINARY_SHIFTER_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : BINARY_SHIFT_RIGHT_FUNC
'DESCRIPTION   : Binary shift-right
'LIBRARY       : NUMBER_BINARY
'GROUP         : SHIFT
'ID            : 002
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 14/02/2008
'************************************************************************************
'************************************************************************************

'// PERFECT

Function BINARY_SHIFT_RIGHT_FUNC(ByVal BINARY_STR As String, _
Optional ByVal NSIZE As Long = 1, _
Optional ByVal RING_FLAG As Boolean = False)

Dim i As Long
Dim j As Long
Dim ATEMP_STR As String
Dim BTEMP_STR As String

On Error GoTo ERROR_LABEL

i = Len(BINARY_STR)
BTEMP_STR = ""
j = NSIZE Mod i
If RING_FLAG Then
    If j > 0 Then BTEMP_STR = Right(BINARY_STR, j)
    ATEMP_STR = Left(BINARY_STR, i - j)
Else
    If NSIZE < i Then
        BTEMP_STR = String(j, "0")
        ATEMP_STR = Left(BINARY_STR, i - j)
    Else
        ATEMP_STR = String(i, "0")
    End If
End If

BINARY_SHIFT_RIGHT_FUNC = BTEMP_STR & ATEMP_STR
Exit Function
ERROR_LABEL:
BINARY_SHIFT_RIGHT_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : BINARY_SHIFT_LEFT_FUNC
'DESCRIPTION   : Binary shift-left
'LIBRARY       : NUMBER_BINARY
'GROUP         : SHIFT
'ID            : 003
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 14/02/2008
'************************************************************************************
'************************************************************************************

'// PERFECT
Function BINARY_SHIFT_LEFT_FUNC(ByVal BINARY_STR As Variant, _
Optional ByVal NSIZE As Variant = 1, _
Optional ByVal RING_FLAG As Boolean = False)

Dim i As Long
Dim j As Long
Dim ATEMP_STR As String
Dim BTEMP_STR As String

On Error GoTo ERROR_LABEL

i = Len(BINARY_STR)
BTEMP_STR = ""
j = NSIZE Mod i
If RING_FLAG Then
    If j > 0 Then BTEMP_STR = Left(BINARY_STR, j)
    ATEMP_STR = Right(BINARY_STR, i - j)
Else
    If NSIZE < i Then
        BTEMP_STR = String(j, "0")
        ATEMP_STR = Right(BINARY_STR, i - j)
    Else
        ATEMP_STR = String(i, "0")
    End If
End If
BINARY_SHIFT_LEFT_FUNC = ATEMP_STR & BTEMP_STR

Exit Function
ERROR_LABEL:
BINARY_SHIFT_LEFT_FUNC = Err.number
End Function
