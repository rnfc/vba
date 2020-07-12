Attribute VB_Name = "STAT_MOMENTS_WEEKLY_LIBR"

'----------------------------------------------------------------------------------------------
'----------------------------------------------------------------------------------------------
Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.
'----------------------------------------------------------------------------------------------
'----------------------------------------------------------------------------------------------


'************************************************************************************
'************************************************************************************
'FUNCTION      : WEEKLY_TREND_INTENSITY_FUNC
'DESCRIPTION   : http://chandoo.org/wp/2010/01/15/flu-trends-chart-review/
'http://www.google.org/flutrends/intl/en_us/us/#US

'LIBRARY       : STATISTICS
'GROUP         :
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 03/08/2010
'************************************************************************************
'************************************************************************************

Function WEEKLY_TREND_INTENSITY_FUNC(ByRef DATA_RNG As Variant, _
Optional ByVal YEAR_INDEX As Long = 2, _
Optional ByVal ASSET_INDEX As Long = 3, _
Optional ByVal START_WEEK As Integer = 13, _
Optional ByVal OUTPUT As Integer = 3)
'DataRng Must Include Headings

Dim h As Long
Dim i As Long
Dim j As Long
Dim k As Long
Dim l As Long
Dim m As Long

Dim SROW As Long
Dim NROWS As Long
Dim NCOLUMNS As Long
Dim NO_PERIODS As Long

Dim MIN_VAL As Double 'Min Intensity
Dim MAX_VAL As Double 'Max Intensity
Dim DATE_VAL As Date
Dim INDEX_ARR As Variant
Dim TEMP_VECTOR As Variant 'Intensities
Dim TEMP_MATRIX As Variant
Dim DATA_MATRIX As Variant

On Error GoTo ERROR_LABEL

DATA_MATRIX = DATA_RNG
SROW = LBound(DATA_MATRIX, 1)
NROWS = UBound(DATA_MATRIX, 1)
NCOLUMNS = UBound(DATA_MATRIX, 2)

'---------------------------------------------------------------------------
GoSub INDEX_LINE
'---------------------------------------------------------------------------
Select Case OUTPUT
'---------------------------------------------------------------------------
Case 0 'Select Dates for comparison
'---------------------------------------------------------------------------
    WEEKLY_TREND_INTENSITY_FUNC = INDEX_ARR
'---------------------------------------------------------------------------
Case Else 'Actual Values
'---------------------------------------------------------------------------
    ReDim TEMP_MATRIX(0 To NO_PERIODS, 1 To k + 2)
    TEMP_MATRIX(0, 1) = UCase(DATA_MATRIX(SROW, ASSET_INDEX + 1))
    TEMP_MATRIX(0, k + 2) = "COMPARED YEAR"
    
    MAX_VAL = -2 ^ 52
    For j = 1 To k
        TEMP_MATRIX(0, j + 1) = INDEX_ARR(2, j + 1)
        For i = 1 To NO_PERIODS: TEMP_MATRIX(i, j + 1) = CVErr(xlErrNA): Next i
        i = IIf(j = 1, START_WEEK, 1)
        l = INDEX_ARR(3, j + 1)
        m = INDEX_ARR(4, j + 1)
        For h = l To m
            TEMP_MATRIX(i, j + 1) = DATA_MATRIX(h, ASSET_INDEX + 1)
            If TEMP_MATRIX(i, j + 1) <> 0 Then
                If TEMP_MATRIX(i, j + 1) > MAX_VAL Then: MAX_VAL = TEMP_MATRIX(i, j + 1)
            Else
                TEMP_MATRIX(i, j + 1) = CVErr(xlErrNA)
            End If
            i = i + 1
        Next h
    Next j
    For i = 1 To NO_PERIODS
        TEMP_MATRIX(i, 1) = i
        If YEAR_INDEX <> 0 Then
            TEMP_MATRIX(i, k + 2) = TEMP_MATRIX(i, 1 + YEAR_INDEX)
        Else
            TEMP_MATRIX(i, k + 2) = CVErr(xlErrNA)
        End If
    Next i
    '---------------------------------------------------------------------------
    If OUTPUT = 1 Then
    '---------------------------------------------------------------------------
        WEEKLY_TREND_INTENSITY_FUNC = TEMP_MATRIX
    '---------------------------------------------------------------------------
    Else 'Maximum Intensity All Time
    '---------------------------------------------------------------------------
        DATA_MATRIX = TEMP_MATRIX
        MIN_VAL = 2 ^ 52
        For i = 1 To NO_PERIODS
            For j = 2 To k + 2
                If IsNumeric(DATA_MATRIX(i, j)) = True Then
                    DATA_MATRIX(i, j) = DATA_MATRIX(i, j) / MAX_VAL * 100
                    If DATA_MATRIX(i, j) < MIN_VAL Then: MIN_VAL = DATA_MATRIX(i, j)
                End If
            Next j
        Next i
        '---------------------------------------------------------------------------
        If OUTPUT = 2 Then
        '---------------------------------------------------------------------------
            Erase TEMP_MATRIX
            WEEKLY_TREND_INTENSITY_FUNC = DATA_MATRIX 'Intensity All Time
        '---------------------------------------------------------------------------
        Else 'If OUTPUT = 3 Then
        '---------------------------------------------------------------------------
            ReDim INDEX_ARR(1 To 3, 0 To NO_PERIODS)
            For j = 1 To NO_PERIODS: For i = 1 To 3: INDEX_ARR(i, j) = "": Next i: Next j
            INDEX_ARR(1, 0) = DATA_MATRIX(0, k + 1)
            INDEX_ARR(2, 0) = DATA_MATRIX(0, YEAR_INDEX + 1)
            INDEX_ARR(3, 0) = ""
            INDEX_ARR(3, 1) = START_WEEK
            For j = 1 To NO_PERIODS
                If IsNumeric(DATA_MATRIX(j, k + 1)) Then
                    INDEX_ARR(1, j) = DATA_MATRIX(j, k + 1)
                Else
                    INDEX_ARR(1, j) = 0
                End If
                If YEAR_INDEX <> 0 Then
                    If IsNumeric(DATA_MATRIX(j, YEAR_INDEX + 1)) Then
                        INDEX_ARR(2, j) = DATA_MATRIX(j, YEAR_INDEX + 1)
                    Else
                        INDEX_ARR(2, j) = 0
                    End If
                Else
                    INDEX_ARR(2, j) = 0
                End If
                If INDEX_ARR(3, j - 1) >= NO_PERIODS Then
                    INDEX_ARR(3, j) = 1
                Else
                    INDEX_ARR(3, j) = INDEX_ARR(3, j - 1) + 1
                End If
            Next j
            
            If OUTPUT = 3 Then
                Erase DATA_MATRIX: Erase TEMP_MATRIX
                WEEKLY_TREND_INTENSITY_FUNC = INDEX_ARR
            Else
                WEEKLY_TREND_INTENSITY_FUNC = Array(TEMP_MATRIX, INDEX_ARR, DATA_MATRIX)
            End If
        End If
    End If
'---------------------------------------------------------------------------
End Select
'---------------------------------------------------------------------------

'---------------------------------------------------------------------------
Exit Function
'-----------------------------------------------------------------------------
INDEX_LINE:
'-----------------------------------------------------------------------------
    i = 1 + SROW: k = 1
    DATE_VAL = DATA_MATRIX(i, 1)
    NO_PERIODS = 52
    
    If START_WEEK = 0 Then 'This returns the week number of the date in
    'DT based on Week 1 starting on January 1 of the year of DT, regardless
    'of what day of week that might be.
        START_WEEK = Int(((DATE_VAL - DateSerial(Year(DATE_VAL), 1, 0)) + 6) / 7)
    ElseIf START_WEEK < 1 Then
        START_WEEK = 1
    ElseIf START_WEEK > NO_PERIODS Then
        START_WEEK = NO_PERIODS
    End If
    
    ReDim INDEX_ARR(1 To 4, 1 To k + 1)
    INDEX_ARR(1, 1) = "START DATE"
    INDEX_ARR(2, 1) = "START YEAR"
    INDEX_ARR(3, 1) = "START ROW" 'Row Starts for Each Week
    INDEX_ARR(4, 1) = "END ROW"
    
    h = Year(DATE_VAL)
    INDEX_ARR(1, k + 1) = DATE_VAL
    INDEX_ARR(2, k + 1) = h & " - " & h + 1
    
    INDEX_ARR(3, k + 1) = i
    INDEX_ARR(4, k + 1) = NO_PERIODS - START_WEEK + i
    l = 1
    For i = 2 + SROW To NROWS
        DATE_VAL = DATA_MATRIX(i, 1)
        If Year(DATE_VAL) <> Year(INDEX_ARR(1, k + 1)) Then
            k = k + 1
            ReDim Preserve INDEX_ARR(1 To 4, 1 To k + 1)
            h = Year(DATE_VAL)
            INDEX_ARR(1, k + 1) = DATE_VAL
            INDEX_ARR(2, k + 1) = h & " - " & h + 1
            If l > NO_PERIODS Then
                INDEX_ARR(3, k + 1) = INDEX_ARR(4, k) + 1 + (l - NO_PERIODS)
            Else
                INDEX_ARR(3, k + 1) = INDEX_ARR(4, k) + 1
            End If
            INDEX_ARR(4, k + 1) = INDEX_ARR(3, k + 1) + NO_PERIODS - 1
            If INDEX_ARR(4, k + 1) > NROWS Then: INDEX_ARR(4, k + 1) = NROWS
            l = 1
        Else
            l = l + 1
        End If
    Next i
    ReDim Preserve INDEX_ARR(1 To 4, 1 To k)
    k = k - 1
'-----------------------------------------------------------------------------
Return
'-----------------------------------------------------------------------------
ERROR_LABEL:
WEEKLY_TREND_INTENSITY_FUNC = Err.number
End Function
