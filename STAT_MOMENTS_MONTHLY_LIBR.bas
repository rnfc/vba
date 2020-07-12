Attribute VB_Name = "STAT_MOMENTS_MONTHLY_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.


'************************************************************************************
'************************************************************************************
'FUNCTION      : MONTHLY_CYCLE_CHART_FUNC
'DESCRIPTION   : Monthly Cycle Chart - Another Way to Look at Trend Data
'LIBRARY       : STATISTICS
'GROUP         :
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 14/02/2008
'************************************************************************************
'************************************************************************************

Function MONTHLY_CYCLE_CHART_FUNC(ByRef DATE_RNG As Variant, _
ByRef DATA_RNG As Variant, _
Optional ByVal YAXIS_MIN_ZOOM As Double = 0.82, _
Optional ByVal YAXIS_MAX_ZOOM As Double = 1.12, _
Optional ByVal YAXIS_POS As Double = 0, _
Optional ByVal OUTPUT As Integer = 1)

'Optional ByVal YAXIS_MAJOR_UNITS As Double = 0, _
'Optional ByVal XAXIS_MIN_ZOOM As Double = 0.95, _
'Optional ByVal XAXIS_MAX_ZOOM As Double = 1.05, _

Dim h As Long
Dim i As Long
Dim j As Long
Dim k As Long
Dim l As Long
Dim n As Long

Dim NROWS As Long
Dim NCOLUMNS As Long

Dim YEAR_INT As Long
Dim MONTH_INT As Long
Dim DATE_VAL As Date
Dim MONTH_STR As String

Dim MONTHLY_MATRIX As Variant
Dim SUMMARY_MATRIX As Variant
Dim CHART_MATRIX As Variant

Dim DATA_VECTOR As Variant
Dim DATE_VECTOR As Variant

Dim INDEX_OBJ As New Collection

On Error GoTo ERROR_LABEL

DATA_VECTOR = DATA_RNG
If UBound(DATA_VECTOR, 1) = 1 Then
    DATA_VECTOR = MATRIX_TRANSPOSE_FUNC(DATA_VECTOR)
End If

DATE_VECTOR = DATE_RNG
If UBound(DATE_VECTOR, 1) = 1 Then
    DATE_VECTOR = MATRIX_TRANSPOSE_FUNC(DATE_VECTOR)
End If
If UBound(DATA_VECTOR, 1) <> UBound(DATE_VECTOR, 1) Then: GoTo ERROR_LABEL
NROWS = UBound(DATA_VECTOR, 1)
If NROWS < 12 Then: GoTo ERROR_LABEL

'----------------------------------------------------------------------------
i = 1: l = 1
YEAR_INT = Year(DATE_VECTOR(i, 1))
Call INDEX_OBJ.Add(i, CStr(l))
For i = 1 To NROWS
    n = Year(DATE_VECTOR(i, 1))
    If n <> YEAR_INT Then
        l = l + 1
        Call INDEX_OBJ.Add(i, CStr(l))
        YEAR_INT = n
    End If
Next i

ReDim SUMMARY_MATRIX(1 To l + 5, 1 To 12 + 2)

DATE_VAL = Now
SUMMARY_MATRIX(1, 1) = "YEAR"
For h = 1 To 12
    MONTH_STR = Format(DateSerial(Year(DATE_VAL), h, 1), "mmm")
    SUMMARY_MATRIX(1, 1 + h) = MONTH_STR
    SUMMARY_MATRIX(l + 4, 1 + h) = -2 ^ 52
    SUMMARY_MATRIX(l + 5, 1 + h) = 2 ^ 52
Next h
SUMMARY_MATRIX(1, h + 1) = "AVERAGE"

SUMMARY_MATRIX(l + 2, 1) = "AVERAGE"
SUMMARY_MATRIX(l + 3, 1) = "COUNTA"

SUMMARY_MATRIX(l + 4, 1) = "MAX"
SUMMARY_MATRIX(l + 5, 1) = "MIN"

For h = 1 To l
    i = CLng(INDEX_OBJ.Item(h))
    SUMMARY_MATRIX(1 + h, 1) = Year(DATE_VECTOR(i, 1))
    If h <> l Then
        j = CLng(INDEX_OBJ.Item(h + 1)) - 1
    Else
        j = NROWS
    End If
    n = 0
    For k = i To j 'For each month per year
        MONTH_INT = Month(DATE_VECTOR(k, 1))
        If DATA_VECTOR(k, 1) <> "" And DATA_VECTOR(k, 1) <> 0 Then
            n = n + 1
            SUMMARY_MATRIX(1 + h, 1 + MONTH_INT) = DATA_VECTOR(k, 1)
            SUMMARY_MATRIX(1 + h, 2 + 12) = SUMMARY_MATRIX(1 + h, 2 + 12) + DATA_VECTOR(k, 1)
            
            SUMMARY_MATRIX(l + 2, 1 + MONTH_INT) = SUMMARY_MATRIX(l + 2, 1 + MONTH_INT) + DATA_VECTOR(k, 1)
            SUMMARY_MATRIX(l + 3, 1 + MONTH_INT) = SUMMARY_MATRIX(l + 3, 1 + MONTH_INT) + 1
            
            If DATA_VECTOR(k, 1) > SUMMARY_MATRIX(l + 4, 1 + MONTH_INT) Then: SUMMARY_MATRIX(l + 4, 1 + MONTH_INT) = DATA_VECTOR(k, 1)
            If DATA_VECTOR(k, 1) < SUMMARY_MATRIX(l + 5, 1 + MONTH_INT) Then: SUMMARY_MATRIX(l + 5, 1 + MONTH_INT) = DATA_VECTOR(k, 1)
        End If
    Next k
    If n <> 0 Then: SUMMARY_MATRIX(1 + h, 2 + 12) = SUMMARY_MATRIX(1 + h, 2 + 12) / n
Next h

SUMMARY_MATRIX(l + 3, 2 + 12) = "MAX/MIN"
SUMMARY_MATRIX(l + 4, 2 + 12) = -2 ^ 52: SUMMARY_MATRIX(l + 5, 2 + 12) = 2 ^ 52
For h = 1 To 12
    If SUMMARY_MATRIX(l + 4, 1 + h) > SUMMARY_MATRIX(l + 4, 2 + 12) Then: SUMMARY_MATRIX(l + 4, 2 + 12) = SUMMARY_MATRIX(l + 4, 1 + h)
    If SUMMARY_MATRIX(l + 5, 1 + h) < SUMMARY_MATRIX(l + 5, 2 + 12) Then: SUMMARY_MATRIX(l + 5, 2 + 12) = SUMMARY_MATRIX(l + 5, 1 + h)
    If SUMMARY_MATRIX(l + 3, 1 + h) <> 0 Then: SUMMARY_MATRIX(l + 2, 1 + h) = SUMMARY_MATRIX(l + 2, 1 + h) / SUMMARY_MATRIX(l + 3, 1 + h)
    SUMMARY_MATRIX(l + 2, 2 + 12) = SUMMARY_MATRIX(l + 2, 2 + 12) + SUMMARY_MATRIX(l + 2, 1 + h) / 12
Next h

Select Case OUTPUT
Case 0
    MONTHLY_CYCLE_CHART_FUNC = SUMMARY_MATRIX
Case 1
    GoSub MONTHLY_LINE
    MONTHLY_CYCLE_CHART_FUNC = MONTHLY_MATRIX
Case Else
    GoSub CHART_LINE
    If OUTPUT = 2 Then
        MONTHLY_CYCLE_CHART_FUNC = CHART_MATRIX
    Else
        GoSub MONTHLY_LINE
        MONTHLY_CYCLE_CHART_FUNC = Array(SUMMARY_MATRIX, MONTHLY_MATRIX, CHART_MATRIX)
    End If
End Select

Exit Function
'------------------------------------------------------------------------------------------------
MONTHLY_LINE:
'------------------------------------------------------------------------------------------------
ReDim MONTHLY_MATRIX(1 To l + 1, 1 To 24 + 1)
h = 1: k = 0 '1
For j = 1 To 24 Step 2
    MONTH_STR = SUMMARY_MATRIX(1, 1 + h)
    h = h + 1
    MONTHLY_MATRIX(1, 1 + j) = MONTH_STR & ": XPOS"
    MONTHLY_MATRIX(1, 1 + j + 1) = MONTH_STR & ": YPOS"
    For i = 1 To l
        MONTHLY_MATRIX(1 + i, 1 + j) = k
        k = k + 1
    Next i
Next j
For i = 1 To l
    MONTHLY_MATRIX(1 + i, 1) = SUMMARY_MATRIX(1 + i, 1)
    k = 1
    For j = 1 To 24 Step 2
        MONTHLY_MATRIX(1 + i, 1 + j + 1) = SUMMARY_MATRIX(1 + i, 1 + k)
        k = k + 1
    Next j
Next i
'------------------------------------------------------------------------------------------------
Return
'------------------------------------------------------------------------------------------------
CHART_LINE:
'------------------------------------------------------------------------------------------------
NROWS = 13: NCOLUMNS = 10
ReDim CHART_MATRIX(0 To NROWS, 1 To NCOLUMNS)
For i = 1 To NROWS: For j = 1 To NCOLUMNS: CHART_MATRIX(i, j) = "": Next j: Next i
CHART_MATRIX(0, 1) = "MONTH"
CHART_MATRIX(0, 2) = "X AXIS LABELS: X-POS"
CHART_MATRIX(0, 3) = "X AXIS LABELS: Y-POS"
CHART_MATRIX(0, 4) = "Y AXIS DIVIDER: X-POS"
CHART_MATRIX(0, 5) = "Y AXIS DIVIDER: Y-POS"
CHART_MATRIX(0, 6) = "Y AXIS DIVIDER: Y-HGT"
CHART_MATRIX(0, 7) = "MONTHLY AVG: MONTH IDX"
CHART_MATRIX(0, 8) = "MONTHLY AVG: X PLOT POS"
CHART_MATRIX(0, 9) = "MONTHLY AVG: AVG MONTH"
CHART_MATRIX(0, 10) = "MONTHLY AVG: DELTA-X"

i = 0
j = Int(l / 2)
For h = 1 To 12
    CHART_MATRIX(h, 1) = SUMMARY_MATRIX(1, 1 + h)
    CHART_MATRIX(h, 2) = j
    j = j + l
    CHART_MATRIX(h, 3) = SUMMARY_MATRIX(l + 5, 2 + 12) * YAXIS_MIN_ZOOM
    CHART_MATRIX(h, 4) = i
    i = i + l
    CHART_MATRIX(h, 5) = YAXIS_POS
    CHART_MATRIX(h, 6) = SUMMARY_MATRIX(l + 4, 2 + 12) * YAXIS_MAX_ZOOM
    CHART_MATRIX(h, 7) = h
    CHART_MATRIX(h, 8) = (h - 1) * l
    CHART_MATRIX(h, 9) = SUMMARY_MATRIX(l + 2, 1 + h)
    CHART_MATRIX(h, 10) = l
Next h
CHART_MATRIX(1, 6) = ""
CHART_MATRIX(NROWS, 1) = "YEAR"
CHART_MATRIX(NROWS, 4) = i
CHART_MATRIX(NROWS, 5) = YAXIS_POS
CHART_MATRIX(NROWS, 6) = SUMMARY_MATRIX(l + 4, 2 + 12) * YAXIS_MAX_ZOOM
'------------------------------------------------------------------------------------------------
Return
'------------------------------------------------------------------------------------------------
ERROR_LABEL:
MONTHLY_CYCLE_CHART_FUNC = Err.number
End Function

