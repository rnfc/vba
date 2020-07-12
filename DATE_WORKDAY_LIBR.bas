Attribute VB_Name = "DATE_WORKDAY_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.

'''''''''''''''''''''''''''''''''''''''''''''''''''''
' EDaysOfWeek by CPearson.com
' Days of the week to exclude. This is a bit-field
' enum, so that its values can be added or OR'd
' together to specify more than one day. E.g,.
' to exclude Tuesday and Saturday, use
' (Tuesday+Saturday), or (Tuesday OR Saturday)
'''''''''''''''''''''''''''''''''''''''''''''''''''''
Enum EDaysOfWeek 'Exclude Day Of Week?
    Sunday = 1      ' 2 ^ (vbSunday - 1)
    Monday = 2      ' 2 ^ (vbMonday - 1)
    Tuesday = 4     ' 2 ^ (vbTuesday - 1)
    Wednesday = 8   ' 2 ^ (vbWednesday - 1)
    Thursday = 16   ' 2 ^ (vbThursday - 1)
    Friday = 32     ' 2 ^ (vbFriday - 1)
    Saturday = 64   ' 2 ^ (vbSaturday - 1)
End Enum

'**********************************************************************************
'**********************************************************************************
'FUNCTION      : WORKDAY1_FUNC

'DESCRIPTION   : This is a replacement for the ATP WORKDAY function. It
'expands on WORKDAY by allowing you to specify any number
'of days of the week to exclude.

'START_DATE: The date on which the period starts.
'NO_DAYS: The number of workdays to include
'                   in the period.
'EXCLUDA_DAY: The sum of the values in EDaysOfWeek
'to exclude. E..g, to exclude Tuesday
'and Saturday, pass Tuesday+Saturday in
'this parameter.

'HOLIDAYS_RNG: an array or range of dates to exclude
'from the period.

'RESULT: A date that is NO_DAYS past
'START_DATE, excluding HOLIDAYS_RNG and
'excluded days of the week.

' Because it is possible that combinations of HOLIDAYS_RNG and
' excluded days of the week could make an end date impossible
' to determine (e.g., exclude all days of the week), the latest
' date that will be calculated is START_DATE + (10 * NO_DAYS).
' This limit is controlled by the k variable.
' If NO_DAYS is less than zero, the result is 0. If
' the k value is exceeded, the result is 0.

'LIBRARY       : DATE
'GROUP         : WORKDAY
'ID            : 001
'**********************************************************************************
'**********************************************************************************

Function WORKDAY1_FUNC(ByVal START_DATE As Date, _
ByVal NO_DAYS As Long, _
ByVal EXCLUDA_DAY As EDaysOfWeek, _
Optional ByVal HOLIDAYS_RNG As Variant) As Variant

'As you can see, each value assigned to a day of week is a power of 2. Each day of week
'turns on one bit of the Enum value. This allows you to specify more than one day of week
'by simply adding the corresponding values together. For example, to exclude Tuesdays and
'Fridays, you would use Tuesday + Friday. Since Tuesday is equal to 4, it has a binary
'representation of 00000100. In binary, Friday, which equals 32, is 00100000. When these
'are added together, the result is 00100100. This shows that the bits for Tuesday and Friday
'are turned on, and all the other day's bits are off. Note that the values used for the
'weekdays are not the same as the constants used by Excel and by VBA (the relationship between
'the Enum's values and the VBA values is shown in the comments within the Enum). If you
'specify that all days of week are to be excluded (EXCLUDA_DAY >= 127), the function will
'return a #VALUE error.

'For example, to find the date that is 15 days past 5-January-2009, excluding Tuesdays and
'Fridays, you can use =WORKDAY1_FUNC(DATE(2009,1,5),15,4+32). The result is 26-January-2009. To
'exclude HOLIDAYS_RNG, put the dates to exclude in some range of cells, say K1:K10, and pass
'that range as the final parameter to WORKDAY1_FUNC:
'=WORKDAY1_FUNC(DATE(2009,1,5),15,4+32,K1:K10).

'START_DATE is the date from which the counting of days begins. NO_DAYS is the number of work
'days after START_DATE whose is to be returned. EXCLUDA_DAY is a value that indicates which days of
'the week to exclude. This is explained below. HOLIDAYS_RNG is an array or range contains the dates
'of HOLIDAYS_RNG to exclude from the calculation.

Dim h As Long
Dim i As Long ' generic counter
Dim j As Long ' days actually worked
Dim k As Long ' prevent infinite looping

Dim NROWS As Long

Dim TEMP_DATE As Date ' incrementing date
Dim TEMP_DAY As EDaysOfWeek ' day of week of TEMP_DATE

Dim HOLIDAY_FLAG As Boolean ' is TEMP_DATE a holiday?
Dim HOLIDAYS_VECTOR As Variant

On Error GoTo ERROR_LABEL

If NO_DAYS < 0 Then: GoTo ERROR_LABEL
' day required must be greater than or equal to zero.
If NO_DAYS = 0 Then
    WORKDAY1_FUNC = START_DATE
    Exit Function
End If

If EXCLUDA_DAY >= (Sunday + Monday + Tuesday + Wednesday + _
            Thursday + Friday + Saturday) Then: GoTo ERROR_LABEL
    ' all days of week excluded. get out with error.

NROWS = 0
If IsArray(HOLIDAYS_RNG) = True Then
    HOLIDAYS_VECTOR = HOLIDAYS_RNG
    If UBound(HOLIDAYS_VECTOR, 1) = 1 Then
        HOLIDAYS_VECTOR = MATRIX_TRANSPOSE_FUNC(HOLIDAYS_VECTOR)
    End If
    NROWS = UBound(HOLIDAYS_VECTOR, 1)
End If

' this prevents an infinite loop which is possible
' under certain circumstances.
k = NO_DAYS * 10
i = 0
j = 0
' loop until the number of actual days worked (j)
' is equal to the specified NO_DAYS.
Do Until j = NO_DAYS
    i = i + 1
    TEMP_DATE = START_DATE + i
    TEMP_DAY = 2 ^ (Weekday(TEMP_DATE) - 1)
    If (TEMP_DAY And EXCLUDA_DAY) = 0 Then
        ' not excluded day of week. continue.
        HOLIDAY_FLAG = False
        If NROWS <> 0 Then: GoSub HOLIDAYS_LINE
        If HOLIDAY_FLAG = False Then: j = j + 1 'TEMP_DATE is not a holiday. Include the date.
    End If
    If i > k Then: GoTo ERROR_LABEL
        ' out of control loop. get out with #VALUE
Loop
' return the result
WORKDAY1_FUNC = START_DATE + i

Exit Function
'---------------------------------------------------------------------------
HOLIDAYS_LINE:
'---------------------------------------------------------------------------
If IsArray(HOLIDAYS_RNG) = True Then
    For h = 1 To NROWS
        If HOLIDAYS_VECTOR(h, 1) = TEMP_DATE Then
            HOLIDAY_FLAG = True
            ' TEMP_DATE is a holiday. get out and
            ' don't count it.
            Exit For
        End If
    Next h
End If
'---------------------------------------------------------------------------
Return
'---------------------------------------------------------------------------
ERROR_LABEL:
WORKDAY1_FUNC = Err.number
End Function

'**********************************************************************************
'**********************************************************************************
'FUNCTION      : NETWORKDAYS_FUNC
'DESCRIPTION   : Returns the number of whole working days between start_date
'and end_date. Working days exclude weekends and any dates identified in
'holidays

'LIBRARY       : DATE
'GROUP         : WORKDAY
'ID            : 002
'LAST UPDATE   : 01/09/2009
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA

'**********************************************************************************
'**********************************************************************************

Function NETWORKDAYS_FUNC(ByVal FIRST_DATE As Date, _
ByVal SECOND_DATE As Date, _
Optional ByVal HOLIDAYS_RNG As Variant)

Dim k As Long
Dim l As Long
Dim START_DATE As Date
Dim END_DATE As Date
Dim SETTLEMENT As Date

On Error GoTo ERROR_LABEL

If FIRST_DATE <= SECOND_DATE Then
    START_DATE = FIRST_DATE
    END_DATE = SECOND_DATE
    l = 1
Else
    START_DATE = SECOND_DATE
    END_DATE = FIRST_DATE
    l = -1
End If

SETTLEMENT = DateAdd("d", 1, START_DATE)
Do While SETTLEMENT <= END_DATE
    Select Case Weekday(SETTLEMENT)
        Case 2 To 6
            k = k + 1
        Case Else
    End Select
    SETTLEMENT = DateAdd("d", 1, SETTLEMENT)
Loop

'TEMP_STR = "NETWORKDAYS(" & """" & Format(START_DATE) & _
            """,""" & Format(END_DATE) & """)"
' NETWORKDAYS_FUNC= Evaluate(TEMP_STR)

If IsArray(HOLIDAYS_RNG) = True Then
    NETWORKDAYS_FUNC = (k - HOLIDAYS_COUNT_FUNC(START_DATE, END_DATE, HOLIDAYS_RNG)) * l
Else
    NETWORKDAYS_FUNC = k * l
End If

Exit Function
ERROR_LABEL:
NETWORKDAYS_FUNC = Err.number
End Function

'**********************************************************************************
'**********************************************************************************
'FUNCTION      : WORKDAY2_FUNC

'DESCRIPTION   : Returns a number that represents a date that is the
'indicated number of working days before or after a date (the
'starting date). Working days exclude weekends and any dates
'identified as holidays.

'LIBRARY       : DATE
'GROUP         : WORKDAY
'ID            : 003
'LAST UPDATE   : 01/09/2009
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA

'**********************************************************************************
'**********************************************************************************

Function WORKDAY2_FUNC(Optional ByVal SETTLEMENT As Date = 0, _
Optional ByVal nDAYS As Long = 1, _
Optional ByRef HOLIDAYS_RNG As Variant)

'SETTLEMENT: is a date that represents the start date.

'NDAYS: is the number of nonweekend and nonholiday days before
'or after SETTLEMENT. A positive value for days yields a future date;
'a negative value yields a past date.

'HOLIDAYS_RNG: is an optional list of one or more dates to
'exclude from the working calendar, such as state and federal
'holidays and floating holidays. The list can be either a range
'of cells that contain the dates or an array constant of the serial
'numbers that represent the dates.


Dim i As Long
Dim j As Long
Dim k As Long
Dim MATCH_FLAG As Boolean
Dim DATE_VAL As Date
Dim HOLIDAYS_OBJ As Collection
Dim HOLIDAYS_VECTOR As Variant
On Error GoTo ERROR_LABEL

If SETTLEMENT = 0 Then: SETTLEMENT = Now

If nDAYS = 0 Then
    WORKDAY2_FUNC = SETTLEMENT
    Exit Function
End If

'------------------------------------------------------------------------------------------------
If IsArray(HOLIDAYS_RNG) = True Then
'------------------------------------------------------------------------------------------------
    GoSub LOAD_LINE
    If nDAYS > 0 Then
        k = 0
        DATE_VAL = SETTLEMENT
        Do While k <> nDAYS
            DATE_VAL = DateAdd("d", 1, DATE_VAL)
            GoSub MATCH_LINE
            If MATCH_FLAG = True Then: GoTo 1983
            Select Case Weekday(DATE_VAL)
                Case 2 To 6
                    k = k + 1
                Case Else
            End Select
1983:
        Loop
    Else
        k = Abs(nDAYS)
        DATE_VAL = SETTLEMENT
        Do While k <> 0
            DATE_VAL = DateAdd("d", -1, DATE_VAL)
            GoSub MATCH_LINE
            If MATCH_FLAG = True Then: GoTo 1984
            Select Case Weekday(DATE_VAL)
                Case 2 To 6
                    k = k - 1
                Case Else
            End Select
1984:
        Loop
    End If
'------------------------------------------------------------------------------------------------
Else
'------------------------------------------------------------------------------------------------
    If nDAYS > 0 Then
        k = 0
        DATE_VAL = SETTLEMENT
        Do While k <> nDAYS
            DATE_VAL = DateAdd("d", 1, DATE_VAL)
            Select Case Weekday(DATE_VAL)
                Case 2 To 6
                    k = k + 1
                Case Else
            End Select
        Loop
    Else
        k = Abs(nDAYS)
        DATE_VAL = SETTLEMENT
        Do While k <> 0
            DATE_VAL = DateAdd("d", -1, DATE_VAL)
            Select Case Weekday(DATE_VAL)
                Case 2 To 6
                    k = k - 1
                Case Else
            End Select
        Loop
    End If
'------------------------------------------------------------------------------------------------
End If
'------------------------------------------------------------------------------------------------

WORKDAY2_FUNC = DATE_VAL

Exit Function
'------------------------------------------------------------------------------------------------
LOAD_LINE:
'------------------------------------------------------------------------------------------------
HOLIDAYS_VECTOR = HOLIDAYS_RNG
If UBound(HOLIDAYS_VECTOR, 1) = 1 Then
    HOLIDAYS_VECTOR = MATRIX_TRANSPOSE_FUNC(HOLIDAYS_VECTOR)
End If
j = UBound(HOLIDAYS_VECTOR, 1)
Set HOLIDAYS_OBJ = New Collection
On Error Resume Next
For i = 1 To j
    If IsDate(HOLIDAYS_VECTOR(i, 1)) Then
        DATE_VAL = HOLIDAYS_VECTOR(i, 1)
        Call HOLIDAYS_OBJ.Add(CStr(i), CStr(DATE_VAL))
    End If
Next i
If Err.number <> 0 Then: Err.Clear
'------------------------------------------------------------------------------------------------
Return
'------------------------------------------------------------------------------------------------
MATCH_LINE:
'------------------------------------------------------------------------------------------------
    j = 0: j = HOLIDAYS_OBJ(CStr(DATE_VAL))
    If j > 0 Then
        MATCH_FLAG = True
    Else
        MATCH_FLAG = False
    End If
'------------------------------------------------------------------------------------------------
Return
'------------------------------------------------------------------------------------------------
ERROR_LABEL:
WORKDAY2_FUNC = DATE_VAL
End Function


'**********************************************************************************
'**********************************************************************************
'FUNCTION      : WORKMONTH_FUNC
'DESCRIPTION   : Returns a number that represents a date that is the
'indicated number of working days before or after a date
'LIBRARY       : DATE
'GROUP         : WORKING
'ID            : 004
'LAST UPDATE   : 01/09/2009
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA

'**********************************************************************************
'**********************************************************************************

Function WORKMONTH_FUNC(ByVal SETTLEMENT As Date, _
Optional ByVal MONTHS_VAL As Integer = 2, _
Optional ByRef HOLIDAYS_RNG As Variant)

On Error GoTo ERROR_LABEL

WORKMONTH_FUNC = WORKDAY2_FUNC(WORKDAY2_FUNC(EDATE_FUNC(SETTLEMENT, MONTHS_VAL), 1, HOLIDAYS_RNG), -1, HOLIDAYS_RNG)

Exit Function
ERROR_LABEL:
WORKMONTH_FUNC = Err.number
End Function


'**********************************************************************************
'**********************************************************************************
'FUNCTION      : WORKMONTH_FUNC
'DESCRIPTION   : Returns a number that represents the working days per year,
'starting on jan 1.
'LIBRARY       : DATE
'GROUP         : WORKING
'ID            : 005
'LAST UPDATE   : 01/09/2009
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA

'**********************************************************************************
'**********************************************************************************

Function WORKDAYS_FUNC(Optional ByVal SETTLEMENT As Date = 0, _
Optional ByVal HOLIDAYS_RNG As Variant)

On Error GoTo ERROR_LABEL

If SETTLEMENT = 0 Then: SETTLEMENT = Now
WORKDAYS_FUNC = NETWORKDAYS_FUNC(DateSerial(Year(SETTLEMENT), 1, 1), DateSerial(Year(SETTLEMENT) + 1, 1, 1), HOLIDAYS_RNG)
Exit Function
ERROR_LABEL:
WORKDAYS_FUNC = Err.number
End Function
