Attribute VB_Name = "FINAN_FI_BOND_TENOR_LIBR"
Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.


'************************************************************************************
'************************************************************************************
'FUNCTION      : BOND_DATES_BOND_TENOR_FUNC
'DESCRIPTION   : FROM DATES TO TENOR
'LIBRARY       : FI_BOND
'GROUP         : TENOR
'ID            : 001
'LAST UPDATE   : 11 / 02 / 2004
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'************************************************************************************
'************************************************************************************

Function BOND_DATES_BOND_TENOR_FUNC(ByVal SETTLEMENT As Date, _
ByVal MATURITY As Date, _
Optional ByVal FREQUENCY As Integer = 2, _
Optional ByVal COUNT_BASIS As Integer = 0)

'TENOR_RNG --> Maturity time vector in years
    
Dim i As Long
Dim j As Long
    
Dim NSIZE As Long
Dim PDAYS_VAL As Long
Dim NDAYS_VAL As Long
    
Dim TEMP_MULT As Double
Dim TEMP_FACTOR As Double
Dim TEMP_VECTOR As Variant

On Error GoTo ERROR_LABEL

If (FREQUENCY < 1) Then
    BOND_DATES_BOND_TENOR_FUNC = 0
    Exit Function
End If

If (MATURITY < SETTLEMENT) Then
    BOND_DATES_BOND_TENOR_FUNC = 0
    Exit Function
End If

If (SETTLEMENT = MATURITY) Then
    ReDim TEMP_VECTOR(1 To 1, 1 To 1)
    TEMP_VECTOR(1, 1) = MATURITY
    BOND_DATES_BOND_TENOR_FUNC = TEMP_VECTOR
    Exit Function
End If
 
PDAYS_VAL = COUPDAYBS_FUNC(SETTLEMENT, MATURITY, FREQUENCY, COUNT_BASIS)
NDAYS_VAL = COUPDAYSNC_FUNC(SETTLEMENT, MATURITY, FREQUENCY, COUNT_BASIS)
NSIZE = PDAYS_VAL + NDAYS_VAL

Select Case COUNT_BASIS
Case 0, 4 'US (NASD) 30/360 ; European 30/360
    TEMP_FACTOR = PDAYS_VAL / NSIZE
Case 1 'Actual / Actual
    TEMP_FACTOR = PDAYS_VAL / NSIZE
Case 2 'Actual / 360
    TEMP_MULT = NSIZE / (360 / FREQUENCY)
    TEMP_FACTOR = PDAYS_VAL / NSIZE * TEMP_MULT
Case 3 'Actual / 365
    TEMP_MULT = NSIZE / (365 / FREQUENCY)
    TEMP_FACTOR = PDAYS_VAL / NSIZE * TEMP_MULT
End Select

j = COUPNUM_FUNC(SETTLEMENT, MATURITY, FREQUENCY)

ReDim TEMP_VECTOR(1 To j, 1 To 1)

TEMP_VECTOR(1, 1) = (1 / FREQUENCY) - (TEMP_FACTOR / FREQUENCY)
For i = 2 To j
    TEMP_VECTOR(i, 1) = TEMP_VECTOR(i - 1, 1) + (1 / FREQUENCY)
Next i
   
BOND_DATES_BOND_TENOR_FUNC = TEMP_VECTOR

Exit Function
ERROR_LABEL:
BOND_DATES_BOND_TENOR_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : BOND_TENOR_DATES_FUNC
'DESCRIPTION   : FROM TENORS TO DATES
'LIBRARY       : FI_BOND
'GROUP         : TENOR
'ID            : 002
'LAST UPDATE   : 11 / 02 / 2004
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'************************************************************************************
'************************************************************************************

Function BOND_TENOR_DATES_FUNC(ByRef TENOR_RNG As Variant, _
Optional ByVal SETTLEMENT As Date, _
Optional ByRef HOLIDAYS_RNG As Variant)

'TENOR_RNG --> Maturity time vector in years

Dim i As Long
Dim j As Double
Dim NROWS As Long
Dim TEMP_VECTOR As Variant
Dim TENOR_VECTOR As Variant

On Error GoTo ERROR_LABEL

If SETTLEMENT = 0 Then
    SETTLEMENT = Now: SETTLEMENT = DateSerial(Year(SETTLEMENT), Month(SETTLEMENT), Day(SETTLEMENT))
End If
TENOR_VECTOR = TENOR_RNG
If UBound(TENOR_VECTOR, 1) = 1 Then
    TENOR_VECTOR = MATRIX_TRANSPOSE_FUNC(TENOR_VECTOR)
End If
NROWS = UBound(TENOR_VECTOR, 1)
ReDim TEMP_VECTOR(1 To NROWS, 1 To 1)
For i = 1 To NROWS
    j = TENOR_VECTOR(i, 1) * 12
    TEMP_VECTOR(i, 1) = WORKMONTH_FUNC(SETTLEMENT, j, HOLIDAYS_RNG)
Next i
BOND_TENOR_DATES_FUNC = TEMP_VECTOR

Exit Function
ERROR_LABEL:
BOND_TENOR_DATES_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : BOND_TENOR_TABLE_FUNC
'DESCRIPTION   : GENERIC TENOR FRAME MAKER IN DAY
'LIBRARY       : FI_BOND
'GROUP         : TENOR
'ID            : 003
'LAST UPDATE   : 11 / 02 / 2004
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'************************************************************************************
'************************************************************************************

Function BOND_TENOR_TABLE_FUNC(Optional ByVal OUTPUT As Integer = 0, _
Optional ByVal SETTLEMENT As Date, _
Optional ByRef HOLIDAYS_RNG As Variant)

Dim i As Long
Const j As Long = 44
Dim k As Long
Dim HEADINGS_ARR As Variant 'Year/Month/Day
Dim TENOR_VECTOR As Variant

On Error GoTo ERROR_LABEL

ReDim TENOR_VECTOR(1 To j, 1 To 3)

If SETTLEMENT = 0 Then: SETTLEMENT = DateSerial(Year(Now), Month(Now), Day(Now))

'FIRST PASS: YEARLY INTEGERS
HEADINGS_ARR = Array(0, 0, 1, 0, 1, 0, 0, 2, 0, 0, 3, 0, 0, 6, 0, 0, 9, 0, 1, 0, 0, _
1, 3, 0, 1, 6, 0, 1, 9, 0, 2, 0, 0, 2, 3, 0, 2, 6, 0, 2, 9, 0, 3, 0, 0, 3, 3, 0, 3, _
6, 0, 3, 9, 0, 4, 0, 0, 4, 3, 0, 4, 6, 0, 4, 9, 0, 5, 0, 0, 5, 3, 0, 5, 6, 0, 5, 9, _
0, 6, 0, 0, 6, 3, 0, 6, 6, 0, 6, 9, 0, 7, 0, 0, 7, 3, 0, 7, 6, 0, 7, 9, 0, 8, 0, 0, _
8, 6, 0, 9, 0, 0, 9, 6, 0, 10, 0, 0, 12, 0, 0, 15, 0, 0, 20, 0, 0, 25, 0, 0, 30, 0, 0)

k = LBound(HEADINGS_ARR)
For i = 1 To j
    TENOR_VECTOR(i, 1) = HEADINGS_ARR(k + 0)
    TENOR_VECTOR(i, 2) = HEADINGS_ARR(k + 1)
    TENOR_VECTOR(i, 3) = HEADINGS_ARR(k + 2)
    k = k + 3
Next i

Select Case OUTPUT
Case 0
    BOND_TENOR_TABLE_FUNC = MATRIX_GET_COLUMN_FUNC(BOND_TENOR_VECTOR_FUNC(TENOR_VECTOR, SETTLEMENT, "Y", "M", "W", "TN", HOLIDAYS_RNG), 2, 1)
Case 1
    BOND_TENOR_TABLE_FUNC = BOND_TENOR_VECTOR_FUNC(TENOR_VECTOR, SETTLEMENT, "Y", "M", "W", "TN", HOLIDAYS_RNG)
Case Else
    BOND_TENOR_TABLE_FUNC = TENOR_VECTOR
End Select

Exit Function
ERROR_LABEL:
BOND_TENOR_TABLE_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : BOND_TENOR_VECTOR_FUNC
'DESCRIPTION   : Returns an array of numbers, and each number represents a date
'that is the indicated number for the tenor
'LIBRARY       : FI_BOND
'GROUP         : TENOR
'ID            : 004
'LAST UPDATE   : 11 / 02 / 2004
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'************************************************************************************
'************************************************************************************

Function BOND_TENOR_VECTOR_FUNC(ByRef DATA_RNG As Variant, _
Optional ByVal SETTLEMENT As Date, _
Optional ByVal YEAR_CHR As String = "Y", _
Optional ByVal MONTH_CHR As String = "M", _
Optional ByVal WEEK_CHR As String = "W", _
Optional ByVal REF_CHR As String = "TN", _
Optional ByVal HOLIDAYS_RNG As Variant)

'COLUMN 1 MUST BE YEARLY INTEGERS IN THE FORM #Y#M#W
'COLUMN 2 MUST BE MONTHLY INTEGERS IN THE FORM #Y#M#W
'COLUMN 3 MUST BE WEEKLY INTEGERS IN THE FORM #Y#M#W

Dim i As Long
Dim NROWS As Long

Dim j As Double
Dim TEMP_CHR As String
Dim DATA_MATRIX As Variant
Dim TEMP_VECTOR As Variant

On Error GoTo ERROR_LABEL

DATA_MATRIX = DATA_RNG
If SETTLEMENT = 0 Then: SETTLEMENT = DateSerial(Year(Now), Month(Now), Day(Now))

NROWS = UBound(DATA_MATRIX, 1)

ReDim TEMP_VECTOR(0 To NROWS, 1 To 2)

TEMP_VECTOR(0, 1) = REF_CHR
TEMP_VECTOR(0, 2) = WORKDAY2_FUNC(SETTLEMENT, 1, HOLIDAYS_RNG)

For i = 1 To NROWS
'THE FOLLOWING CALCULATIONS USE THE Number OF MONTHS AS REFERENCE
    TEMP_CHR = ""
    j = 0
    '---------------------------------YEARLY------------------------------
    If IsEmpty(DATA_MATRIX(i, 1)) = False And DATA_MATRIX(i, 1) > 0 Then
        TEMP_CHR = DATA_MATRIX(i, 1) & YEAR_CHR
        j = DATA_MATRIX(i, 1) * 12
    End If
    '--------------------------------MONTHLY------------------------------
    If IsEmpty(DATA_MATRIX(i, 2)) = False And DATA_MATRIX(i, 2) > 0 Then
        TEMP_CHR = TEMP_CHR & DATA_MATRIX(i, 2) & MONTH_CHR
        j = j + DATA_MATRIX(i, 2)
    End If
    '---------------------------------WEEKLY------------------------------
    If IsEmpty(DATA_MATRIX(i, 3)) = False And DATA_MATRIX(i, 3) > 0 Then
        TEMP_CHR = TEMP_CHR & DATA_MATRIX(i, 3) & WEEK_CHR
        TEMP_VECTOR(i, 1) = TEMP_CHR 'TENOR
        TEMP_VECTOR(i, 2) = EDATE_FUNC(WORKDAY2_FUNC(SETTLEMENT, 5 * DATA_MATRIX(i, 3), HOLIDAYS_RNG), j)
    Else
        TEMP_VECTOR(i, 1) = TEMP_CHR 'TENOR
        TEMP_VECTOR(i, 2) = WORKMONTH_FUNC(SETTLEMENT, j, HOLIDAYS_RNG)
    End If
Next i

BOND_TENOR_VECTOR_FUNC = TEMP_VECTOR

Exit Function
ERROR_LABEL:
BOND_TENOR_VECTOR_FUNC = Err.number
End Function
