Attribute VB_Name = "STAT_REGRESSION_MEDIAN_LIBR"


Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.

'************************************************************************************
'************************************************************************************
'FUNCTION      : MEDIAN_REGRESSION_FUNC

'DESCRIPTION   : This function performs the robust linear regression
'with 3 different methods:
'     - SM: simple median
'     - RM: Repeated median
'     - LMS: least median squared
'These methods are suitable for data containing wrong points. When data
'has noise (experimental data is always noisy), the basic problem is
'that classic LMS (least minimum squared) is highly affected by noisy
'points. The main goal of robust methods is to minimize as much as
'possible the influence of the noise

'returns coefficients [a1,a0] of linear regression YDATA_RNG = a1*x+a0

'LIBRARY       : STATISTICS
'GROUP         : REGRESSION_MEDIAN
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/22/2009
'************************************************************************************
'************************************************************************************

Function MEDIAN_REGRESSION_FUNC(ByRef XDATA_RNG As Variant, _
ByRef YDATA_RNG As Variant, _
Optional ByVal VERSION As Integer = 2)

Dim i As Long
Dim j As Long
Dim k As Long

Dim SROW As Long
Dim NROWS As Long

Dim PEND_VAL As Double
Dim ORDEN_VAL As Double

Dim XDATA_VECTOR As Variant
Dim YDATA_VECTOR As Variant

Dim TEMP1_VECTOR As Variant
Dim TEMP2_VECTOR As Variant

Dim RECT_VECTOR As Variant
Dim RESID_VECTOR As Variant
Dim MEDIAN_VECTOR As Variant
Dim PEND_VECTOR As Variant

On Error GoTo ERROR_LABEL

XDATA_VECTOR = XDATA_RNG
If UBound(XDATA_VECTOR, 1) = 1 Then
    XDATA_VECTOR = MATRIX_TRANSPOSE_FUNC(XDATA_VECTOR)
End If

YDATA_VECTOR = YDATA_RNG
If UBound(YDATA_VECTOR, 1) = 1 Then
    YDATA_VECTOR = MATRIX_TRANSPOSE_FUNC(YDATA_VECTOR)
End If
If UBound(XDATA_VECTOR, 1) <> UBound(YDATA_VECTOR, 1) Then: GoTo ERROR_LABEL

SROW = LBound(XDATA_VECTOR, 1)
NROWS = UBound(XDATA_VECTOR, 1)

'---------------------------------------------------------------------
Select Case VERSION
'---------------------------------------------------------------------
Case 0 ', "L", "LMS" 'simple median
'---------------------------------------------------------------------
    ReDim RECT_VECTOR(SROW To (NROWS - SROW) * (NROWS + 1 - SROW) / 2, 1 To 3)
    ReDim RESID_VECTOR(SROW To NROWS, 1 To 1)
    ReDim MEDIAN_VECTOR(SROW To (NROWS - SROW) * (NROWS + 1 - SROW) / 2, 1 To 1)
    k = SROW
    For i = SROW To NROWS - 1
        For j = i + 1 To NROWS
            RECT_VECTOR(k, 1) = (YDATA_VECTOR(i, 1) - YDATA_VECTOR(j, 1)) / (XDATA_VECTOR(i, 1) - XDATA_VECTOR(j, 1))
            RECT_VECTOR(k, 2) = YDATA_VECTOR(i, 1) - (RECT_VECTOR(k, 1) * XDATA_VECTOR(i, 1))
            k = k + 1
        Next j
    Next i
    For i = SROW To ((NROWS - SROW) * (NROWS + 1 - SROW) / 2)
        For j = SROW To NROWS
            RESID_VECTOR(j, 1) = (YDATA_VECTOR(j, 1) - RECT_VECTOR(i, 1) * XDATA_VECTOR(j, 1) - RECT_VECTOR(i, 2)) ^ 2
        Next j
        Call MEDIAN_REGRESSION_SORT_FUNC(RESID_VECTOR, SROW, NROWS)
        If ((NROWS + 1 - SROW) Mod 2) = 0 Then
            MEDIAN_VECTOR(i, 1) = (RESID_VECTOR((NROWS + 1 - SROW) / 2, 1) + RESID_VECTOR((NROWS + 3 - SROW) / 2, 1)) / 2
        Else
            MEDIAN_VECTOR(i, 1) = RESID_VECTOR((NROWS + 2 - SROW) / 2, 1)
        End If
        RECT_VECTOR(i, 3) = MEDIAN_VECTOR(i, 1)
    Next i
    NROWS = ((NROWS - SROW) * (NROWS + 1 - SROW) / 2)
    Call MEDIAN_REGRESSION_SORT_FUNC(MEDIAN_VECTOR, SROW, NROWS)
    i = 1
    Do
        If MEDIAN_VECTOR(1, 1) = RECT_VECTOR(i, 3) Then
            PEND_VAL = RECT_VECTOR(i, 1)
            ORDEN_VAL = RECT_VECTOR(i, 2)
            i = NROWS
        End If
        i = i + 1
    Loop While i <= NROWS
'----------------------------------------------------------------------------
Case 1 ', "R", "RM" 'Repeated median
'---------------------------------------------------------------------
    ReDim PEND_VECTOR(SROW To NROWS - 1, SROW To NROWS - 1)
    ReDim TEMP1_VECTOR(SROW To NROWS - 1, 1 To 1)
    ReDim TEMP2_VECTOR(SROW To NROWS, 1 To 1)
    k = SROW
    For i = SROW To NROWS - 1
        For j = i + 1 To NROWS
            PEND_VECTOR(i, j - 1) = (YDATA_VECTOR(i, 1) - YDATA_VECTOR(j, 1)) / (XDATA_VECTOR(i, 1) - XDATA_VECTOR(j, 1))
        Next j
    Next i
    NROWS = NROWS - 1
    For i = SROW To NROWS + 1
        k = SROW
        For j = SROW To NROWS 'here
            If j >= i Then
                TEMP1_VECTOR(k, 1) = PEND_VECTOR(i, j)
                k = k + 1
            ElseIf j < i Then
                TEMP1_VECTOR(k, 1) = PEND_VECTOR(j, i - 1)
                k = k + 1
            End If
        Next j
        Call MEDIAN_REGRESSION_SORT_FUNC(TEMP1_VECTOR, SROW, NROWS)
        If ((NROWS + 1 - SROW) Mod 2) = 0 Then
            TEMP2_VECTOR(i, 1) = (TEMP1_VECTOR((NROWS + 1 - SROW) / 2, 1) + TEMP1_VECTOR((NROWS + 3 - SROW) / 2, 1)) / 2
        Else
            TEMP2_VECTOR(i, 1) = TEMP1_VECTOR((NROWS + 2 - SROW) / 2, 1)
        End If
    Next i
    Call MEDIAN_REGRESSION_SORT_FUNC(TEMP2_VECTOR, SROW, NROWS + 1)
    If ((NROWS + 2 - SROW) Mod 2) = 0 Then
        PEND_VAL = (TEMP2_VECTOR((NROWS + 2 - SROW) / 2, 1) + TEMP2_VECTOR((NROWS + 4 - SROW) / 2, 1)) / 2
    Else
        PEND_VAL = TEMP2_VECTOR((NROWS + 3 - SROW) / 2, 1)
    End If
    NROWS = UBound(XDATA_VECTOR, 1)
    ReDim TEMP1_VECTOR(SROW To NROWS, 1 To 1)
    For i = SROW To NROWS
        TEMP1_VECTOR(i, 1) = YDATA_VECTOR(i, 1) - PEND_VAL * XDATA_VECTOR(i, 1)
    Next i
    Call MEDIAN_REGRESSION_SORT_FUNC(TEMP1_VECTOR, SROW, NROWS)
    If ((NROWS + 1 - SROW) Mod 2) = 0 Then
        ORDEN_VAL = (TEMP1_VECTOR((NROWS + 1 - SROW) / 2, 1) + TEMP1_VECTOR((NROWS + 3 - SROW) / 2, 1)) / 2
    Else
        ORDEN_VAL = TEMP1_VECTOR((NROWS + 2 - SROW) / 2, 1)
    End If
'----------------------------------------------------------------------------
Case Else 'least median squared
'----------------------------------------------------------------------------
    ReDim PEND_VECTOR(SROW To (NROWS - SROW) * (NROWS + 1 - SROW) / 2, 1 To 1)
    k = SROW
    For i = SROW To NROWS - 1
        For j = i + 1 To NROWS
            PEND_VECTOR(k, 1) = (YDATA_VECTOR(i, 1) - YDATA_VECTOR(j, 1)) / (XDATA_VECTOR(i, 1) - XDATA_VECTOR(j, 1))
            k = k + 1
        Next j
    Next i
    Call MEDIAN_REGRESSION_SORT_FUNC(PEND_VECTOR, SROW, UBound(PEND_VECTOR, 1))
    NROWS = (NROWS - SROW) * (NROWS + 1 - SROW) / 2
    If ((NROWS + 1 - SROW) Mod 2) = 0 Then
        PEND_VAL = (PEND_VECTOR((NROWS + 1 - SROW) / 2, 1) + PEND_VECTOR((NROWS + 3 - SROW) / 2, 1)) / 2
    Else
        PEND_VAL = PEND_VECTOR((NROWS + 2 - SROW) / 2, 1)
    End If
    NROWS = UBound(XDATA_VECTOR, 1)
    ReDim PEND_VECTOR(SROW To NROWS, 1 To 1)
    For i = SROW To NROWS
        PEND_VECTOR(i, 1) = YDATA_VECTOR(i, 1) - PEND_VAL * XDATA_VECTOR(i, 1)
    Next i
    Call MEDIAN_REGRESSION_SORT_FUNC(PEND_VECTOR, SROW, NROWS)
    If ((NROWS + 1 - SROW) Mod 2) = 0 Then
        ORDEN_VAL = (PEND_VECTOR((NROWS + 1 - SROW) / 2, 1) + PEND_VECTOR((NROWS + 3 - SROW) / 2, 1)) / 2
    Else
        ORDEN_VAL = PEND_VECTOR((NROWS + 2 - SROW) / 2, 1)
    End If
'----------------------------------------------------------------------------
End Select
'----------------------------------------------------------------------------

MEDIAN_REGRESSION_FUNC = Array(PEND_VAL, ORDEN_VAL)

Exit Function
ERROR_LABEL:
MEDIAN_REGRESSION_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : MEDIAN_REGRESSION_SORT_FUNC
'DESCRIPTION   : ROBUST_DATA_SORT FOR ROBUST REGRESSION
'LIBRARY       : STATISTICS
'GROUP         : REGRESSION_MEDIAN
'ID            : 002
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/22/2009
'************************************************************************************
'************************************************************************************

Private Function MEDIAN_REGRESSION_SORT_FUNC(ByRef DATA_VECTOR As Variant, _
ByVal SROW As Long, _
ByVal NROWS As Long)

Dim i As Long
Dim j As Long

Dim A_VAL As Double
Dim B_VAL As Double

On Error GoTo ERROR_LABEL

MEDIAN_REGRESSION_SORT_FUNC = False

i = SROW
j = NROWS
A_VAL = DATA_VECTOR((i + j) / 2, 1)

Do
    Do While DATA_VECTOR(i, 1) < A_VAL
        i = i + 1
    Loop
    Do While DATA_VECTOR(j, 1) > A_VAL
        j = j - 1
    Loop
    If i <= j Then
        B_VAL = DATA_VECTOR(i, 1)
        DATA_VECTOR(i, 1) = DATA_VECTOR(j, 1)
        DATA_VECTOR(j, 1) = B_VAL
        i = i + 1
        j = j - 1
    End If
Loop Until i > j

If SROW < j Then: Call MEDIAN_REGRESSION_SORT_FUNC(DATA_VECTOR, SROW, j)
If i < NROWS Then: Call MEDIAN_REGRESSION_SORT_FUNC(DATA_VECTOR, i, NROWS)

MEDIAN_REGRESSION_SORT_FUNC = True

Exit Function
ERROR_LABEL:
MEDIAN_REGRESSION_SORT_FUNC = False
End Function
