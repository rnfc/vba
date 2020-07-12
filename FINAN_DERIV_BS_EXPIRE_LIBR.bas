Attribute VB_Name = "FINAN_DERIV_BS_EXPIRE_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.


'Calculate the number of days to option expiration, given the current date
'and the option's expiration month and year.

'MONTH_VAL: option expiration month (1 - 12)
'YEAR_VAL: option expiration year (e.g. 2002)

Function OPTION_EXPIRE_DAYS_COUNT_FUNC(ByVal CURRENT_DATE As Date, _
ByVal MONTH_VAL As Double, _
ByVal YEAR_VAL As Double) As Integer
    
Dim DATE_VAL As Date

On Error GoTo ERROR_LABEL

Select Case True
Case IsDate(CURRENT_DATE) = False
    GoTo ERROR_LABEL
Case IsNumeric(MONTH_VAL) = False
    GoTo ERROR_LABEL
Case IsNumeric(YEAR_VAL) = False
    GoTo ERROR_LABEL
Case DateSerial(YEAR_VAL, MONTH_VAL, Day(CURRENT_DATE)) < CURRENT_DATE
    GoTo ERROR_LABEL
Case MONTH_VAL > 12
    GoTo ERROR_LABEL
Case YEAR_VAL > (Year(CURRENT_DATE) + 5)
    GoTo ERROR_LABEL
End Select

'find first day of expiration month
DATE_VAL = Int(CURRENT_DATE)
Do While YEAR_VAL > Year(DATE_VAL): DATE_VAL = DATE_VAL + 1: Loop
Do While MONTH_VAL > Month(DATE_VAL): DATE_VAL = DATE_VAL + 1: Loop
DATE_VAL = DATE_VAL - Day(DATE_VAL) + 1
'find first Friday of expiration month
Do While Weekday(DATE_VAL) <> 6: DATE_VAL = DATE_VAL + 1: Loop
'days to expiration is first Friday + 14 days.
OPTION_EXPIRE_DAYS_COUNT_FUNC = DATE_VAL + 14 - Int(CURRENT_DATE) + 1
If OPTION_EXPIRE_DAYS_COUNT_FUNC < 0 Then OPTION_EXPIRE_DAYS_COUNT_FUNC = 0
    
Exit Function
ERROR_LABEL:
OPTION_EXPIRE_DAYS_COUNT_FUNC = Err.number
End Function


'Call and put option symbols include a strike price and expiration month. The second to last letter
'in the symbol denotes the month of expiration and the last term the price. So, for example, GEKD would
'be the symbol for the General Electric November 20.00 calls. GERT would be the June 17.50 puts.
'Here is a Select…Case structure using the char data type to determine the month of expiration

Function OPTION_EXPIRE_MONTH_STR_FUNC(ByVal CHR_MONTH As String)

Dim MONTH_STR As String

On Error GoTo ERROR_LABEL

Select Case LCase(CHR_MONTH)
Case "a", "m"
    MONTH_STR = "January"
Case "b", "n"
    MONTH_STR = "February"
Case "c", "o"
    MONTH_STR = "March"
Case "d", "p"
    MONTH_STR = "April"
Case "e", "q"
    MONTH_STR = "May"
Case "f", "r"
    MONTH_STR = "June"
Case "g", "s"
    MONTH_STR = "July"
Case "h", "t"
    MONTH_STR = "August"
Case "i", "u"
    MONTH_STR = "September"
Case "j", "v"
    MONTH_STR = "October"
Case "k", "w"
    MONTH_STR = "November"
Case "l", "x"
    MONTH_STR = "December"
End Select

OPTION_EXPIRE_MONTH_STR_FUNC = MONTH_STR

Exit Function
ERROR_LABEL:
OPTION_EXPIRE_MONTH_STR_FUNC = Err.number
End Function


Function OPTION_EXPIRE_DATE_FUNC(ByVal TICKER_STR As String, _
Optional ByVal SETTLEMENT As Date = 0)

Dim i As Long
Dim CHR_STR As String
Dim HEADINGS1_ARR As Variant
Dim HEADINGS2_ARR As Variant

On Error GoTo ERROR_LABEL

If SETTLEMENT = 0 Then: SETTLEMENT = Now
CHR_STR = Trim(Left(Right(TICKER_STR, 2), 1))

HEADINGS1_ARR = Array("A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X")
HEADINGS2_ARR = Array(1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12)
For i = LBound(HEADINGS1_ARR) To UBound(HEADINGS1_ARR)
    If HEADINGS1_ARR(i) = CHR_STR Then '3rd Friday
        OPTION_EXPIRE_DATE_FUNC = NTH_DAY_WEEK_FUNC(Year(SETTLEMENT), CInt(HEADINGS2_ARR(i)), 3, 6)
        Exit Function
    End If
Next i
GoTo ERROR_LABEL

Exit Function
ERROR_LABEL:
OPTION_EXPIRE_DATE_FUNC = Err.number
End Function


Function OPTION_EXPIRE_MONTH_FUNC(ByVal OPTION_SYMBOL As String, _
Optional ByVal VERSION As Integer = 0)

Dim i As Long
Dim MONTH_STR As String 'Option Month Code
Dim HEADINGS_ARR As Variant

On Error GoTo ERROR_LABEL

MONTH_STR = UCase(Trim(Left(Right(OPTION_SYMBOL, 2), 1)))

Select Case VERSION
Case 0 'Option Expiration Month (text)
    HEADINGS_ARR = Array("A", "-01", "B", "-02", "C", "-03", "D", "-04", "E", "-05", "F", "-06", "G", "-07", "H", "-08", "I", "-09", "J", "-10", "K", "-11", "L", "-12", "M", "-01", "N", "-02", "O", "-03", "P", "-04", "Q", "-05", "R", "-06", "S", "-07", "T", "-08", "U", "-09", "V", "-10", "W", "-11", "X", "-12")
Case Else 'Option Expiration Month
    HEADINGS_ARR = Array("A", "Jan", "B", "Feb", "C", "Mar", "D", "Apr", "E", "May", "F", "Jun", "G", "Jul", "H", "Aug", "I", "Sep", "J", "Oct", "K", "Nov", "L", "Dec", "M", "Jan", "N", "Feb", "O", "Mar", "P", "Apr", "Q", "May", "R", "Jun")
End Select

For i = LBound(HEADINGS_ARR) To UBound(HEADINGS_ARR)
    If HEADINGS_ARR(i) = MONTH_STR Then
        MONTH_STR = HEADINGS_ARR(i + 1)
        Exit For
    End If
Next i

OPTION_EXPIRE_MONTH_FUNC = MONTH_STR

Exit Function
ERROR_LABEL:
OPTION_EXPIRE_MONTH_FUNC = Err.number
End Function
