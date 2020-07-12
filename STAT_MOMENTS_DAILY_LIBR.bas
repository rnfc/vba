Attribute VB_Name = "STAT_MOMENTS_DAILY_LIBR"
Option Explicit
Option Base 1

'************************************************************************************
'************************************************************************************
'FUNCTION      : DAILY_ANNUAL_STATS_FUNC

'DESCRIPTION   :
'General purpose tool to allow users to convert raw station daily data to annual
'average values and to conduct analysis of annual trends.

'Procedure to process raw temp data and produce annual stats
'Procedure works on sheet specified in next line. If you have
'more than 1 source data sheet, make sure to start with first sheet,
'then proceed through other sheets in cronological order

'LIBRARY       : STATISTICS
'GROUP         : TREND_DAILY
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/22/2009
'************************************************************************************
'************************************************************************************

Function DAILY_ANNUAL_STATS_FUNC(ByRef DATE_RNG As Variant, _
ByRef DATA_RNG As Variant, _
Optional ByVal START_BASE_VAL As Long = 1951, _
Optional ByVal END_BASE_VAL As Long = 1980)
'ANOMALY Base Period

Dim h As Long
Dim i As Long
Dim j As Long
Dim k As Long
Dim l As Long
Dim m As Long
Dim n As Long

Dim NROWS As Long
Dim TEMP_VAL As Double

Dim YEAR_INT As Long
Dim FORMAT_STR As String
Dim DMEAN_VAL As Double
Dim TMEAN_VAL As Double

Dim INDEX_OBJ As New Collection

Dim TEMP_VECTOR As Variant
Dim DATE_VECTOR As Variant
Dim DATA_VECTOR As Variant
Dim TEMP_MATRIX As Variant ' Annual Temp array

On Error GoTo ERROR_LABEL

DATE_VECTOR = DATE_RNG
If UBound(DATE_VECTOR, 1) = 1 Then
    DATE_VECTOR = MATRIX_TRANSPOSE_FUNC(DATE_VECTOR)
End If
DATA_VECTOR = DATA_RNG
If UBound(DATA_VECTOR, 1) = 1 Then
    DATA_VECTOR = MATRIX_TRANSPOSE_FUNC(DATA_VECTOR)
End If
If UBound(DATA_VECTOR, 1) <> UBound(DATE_VECTOR, 1) Then: GoTo ERROR_LABEL
NROWS = UBound(DATE_VECTOR, 1)

i = 1: l = 1
Call INDEX_OBJ.Add(i, CStr(l))

YEAR_INT = Year(DATE_VECTOR(i, 1))
l = l + 1
For i = 2 To NROWS
    n = Year(DATE_VECTOR(i, 1))
    If n <> YEAR_INT Then
        YEAR_INT = n
        Call INDEX_OBJ.Add(i, CStr(l))
        l = l + 1
    End If
Next i
i = i - 1
Call INDEX_OBJ.Add(i, CStr(l))

NROWS = l - 1
If NROWS <= 0 Then: GoTo ERROR_LABEL
'Calculate running total, number observation, avg for year
ReDim TEMP_MATRIX(0 To NROWS, 1 To 12)
TEMP_MATRIX(0, 1) = "YEAR"
TEMP_MATRIX(0, 2) = "STARTING VALUE"
TEMP_MATRIX(0, 3) = "ENDING VALUE"
TEMP_MATRIX(0, 4) = "OBS"

TEMP_MATRIX(0, 5) = "25th PERCENTILE"
TEMP_MATRIX(0, 6) = "MINIMUM"
TEMP_MATRIX(0, 7) = "MEAN"
TEMP_MATRIX(0, 8) = "50th PERCENTILE"
TEMP_MATRIX(0, 9) = "MAXIMUM"
TEMP_MATRIX(0, 10) = "75th PERCENTILE"

TEMP_MATRIX(0, 11) = "RANGE"
TEMP_MATRIX(0, 12) = "SUM"

l = 0
For h = 1 To NROWS
    i = CLng(INDEX_OBJ(h))
    j = CLng(INDEX_OBJ(h + 1)) - 1
    m = j - i + 1
    TEMP_MATRIX(h, 1) = Year(DATE_VECTOR(i, 1))
    
    TEMP_MATRIX(h, 2) = DATA_VECTOR(i, 1)
    TEMP_MATRIX(h, 3) = DATA_VECTOR(j, 1)
    TEMP_MATRIX(h, 4) = 0 'm
    ReDim TEMP_VECTOR(1 To m, 1 To 1)
    TEMP_MATRIX(h, 6) = 2 ^ 52: TEMP_MATRIX(h, 9) = -2 ^ 52
    TEMP_MATRIX(h, 12) = 0
    m = 1
    For k = i To j
        TEMP_MATRIX(h, 4) = TEMP_MATRIX(h, 4) + 1
        TEMP_VAL = DATA_VECTOR(k, 1)
        TEMP_VECTOR(m, 1) = TEMP_VAL
        m = m + 1
        If TEMP_VAL < TEMP_MATRIX(h, 6) Then: TEMP_MATRIX(h, 6) = TEMP_VAL
        If TEMP_VAL > TEMP_MATRIX(h, 9) Then: TEMP_MATRIX(h, 9) = TEMP_VAL
        TEMP_MATRIX(h, 12) = TEMP_MATRIX(h, 12) + TEMP_VAL
    Next k
    TEMP_MATRIX(h, 7) = TEMP_MATRIX(h, 12) / TEMP_MATRIX(h, 4)
    DMEAN_VAL = DMEAN_VAL + TEMP_MATRIX(h, 7)
    If TEMP_MATRIX(h, 1) >= START_BASE_VAL And TEMP_MATRIX(h, 1) <= END_BASE_VAL Then
        l = l + 1
        TMEAN_VAL = TMEAN_VAL + TEMP_MATRIX(h, 7)
    End If
    TEMP_MATRIX(h, 11) = TEMP_MATRIX(h, 9) - TEMP_MATRIX(h, 6)
    TEMP_VECTOR = MATRIX_QUICK_SORT_FUNC(TEMP_VECTOR, 1, 1)
    TEMP_MATRIX(h, 5) = HISTOGRAM_PERCENTILE_FUNC(TEMP_VECTOR, 0.25, 0)
    TEMP_MATRIX(h, 8) = HISTOGRAM_PERCENTILE_FUNC(TEMP_VECTOR, 0.5, 0)
    TEMP_MATRIX(h, 10) = HISTOGRAM_PERCENTILE_FUNC(TEMP_VECTOR, 0.75, 0)
Next h

If START_BASE_VAL = 0 And END_BASE_VAL = 0 Then
    DAILY_ANNUAL_STATS_FUNC = TEMP_MATRIX
    Exit Function
End If

DMEAN_VAL = DMEAN_VAL / NROWS
TMEAN_VAL = TMEAN_VAL / l
FORMAT_STR = "#,##0.0000"
'Get Baseline period
ReDim Preserve TEMP_MATRIX(0 To NROWS, 1 To 15)
TEMP_MATRIX(0, 13) = "NEG ANOMALY: " & Format(TMEAN_VAL, FORMAT_STR)
'Procedure to calculate anomalies for baseline period for annual data
TEMP_MATRIX(0, 14) = "POS ANOMALY"
'Procedure to calculate anomalies for baseline period for annual data
TEMP_MATRIX(0, 15) = "CUSUM: " & Format(DMEAN_VAL, FORMAT_STR)
'Procedure to calculate CuSum

h = 1
TEMP_MATRIX(h, 13) = "": TEMP_MATRIX(h, 14) = ""
TEMP_MATRIX(h, 15) = TEMP_MATRIX(h, 7) - DMEAN_VAL
TEMP_VAL = TEMP_MATRIX(h, 7) - TMEAN_VAL
If TEMP_VAL < 0 Then TEMP_MATRIX(h, 13) = TEMP_VAL ' neg ANOMALY
If TEMP_VAL > 0 Then TEMP_MATRIX(h, 14) = TEMP_VAL ' pos ANOMALY
For h = 2 To NROWS
    TEMP_MATRIX(h, 13) = "": TEMP_MATRIX(h, 14) = ""
    TEMP_VAL = TEMP_MATRIX(h, 7) - TMEAN_VAL
    If TEMP_VAL < 0 Then TEMP_MATRIX(h, 13) = TEMP_VAL ' neg ANOMALY
    If TEMP_VAL > 0 Then TEMP_MATRIX(h, 14) = TEMP_VAL ' pos ANOMALY
    TEMP_MATRIX(h, 15) = TEMP_MATRIX(h - 1, 15) + (TEMP_MATRIX(h, 7) - DMEAN_VAL)
Next h

DAILY_ANNUAL_STATS_FUNC = TEMP_MATRIX

Exit Function
ERROR_LABEL:
DAILY_ANNUAL_STATS_FUNC = Err.number
End Function
