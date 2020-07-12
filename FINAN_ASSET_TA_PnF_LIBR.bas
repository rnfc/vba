Attribute VB_Name = "FINAN_ASSET_TA_PnF_LIBR"


Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.
                    

'************************************************************************************
'************************************************************************************
'FUNCTION      : ASSET_TA_PnF_FUNC
'DESCRIPTION   : PnF Grid Function
'LIBRARY       : FINAN_ASSET
'GROUP         : HIGH-LOW
'ID            : 006
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA

'************************************************************************************
'************************************************************************************

Function ASSET_TA_PnF_FUNC(ByVal TICKER_STR As Variant, _
Optional ByVal START_DATE As Date, _
Optional ByVal END_DATE As Date, _
Optional ByVal REVERSE_VAL As Integer = 3, _
Optional ByVal BOX_VAL As Integer = 1, _
Optional ByVal INC_FACTOR As Double = 1, _
Optional ByVal FIRST_STR As String = "O", _
Optional ByVal SECOND_STR As String = "X", _
Optional ByVal CONFIDENCE_VAL As Double = 0.05, _
Optional ByVal OUTPUT As Integer = 0)

Dim i As Integer
Dim j As Integer

Dim ii As Integer

Dim H1 As Integer
Dim L1 As Integer

Dim H2 As Integer
Dim L2 As Integer

Dim NROWS As Integer

Dim ROW_VAL As Integer
Dim COL_VAL As Integer

Dim MODE_STR As String

Dim LOW_MATRIX As Variant
Dim HIGH_MATRIX As Variant
Dim DATA_MATRIX As Variant

Dim HIGH_VECTOR As Variant
Dim LOW_VECTOR As Variant
Dim PRICES_VECTOR As Variant

On Error GoTo ERROR_LABEL

If IsArray(TICKER_STR) = True Then
    DATA_MATRIX = TICKER_STR
Else
    DATA_MATRIX = _
    YAHOO_HISTORICAL_DATA_SERIE_FUNC(TICKER_STR, START_DATE, END_DATE, _
    "d", "DOHLCVA", False, True, True)
End If
NROWS = UBound(DATA_MATRIX, 1)

ReDim HIGH_VECTOR(1 To NROWS, 1 To 1)
ReDim LOW_VECTOR(1 To NROWS, 1 To 1)
ReDim PRICES_VECTOR(1 To NROWS, 1 To 1)

For i = 1 To NROWS
    HIGH_VECTOR(i, 1) = Round(DATA_MATRIX(i, 3) * INC_FACTOR, 0)
    LOW_VECTOR(i, 1) = Round(DATA_MATRIX(i, 4) * INC_FACTOR, 0)
    PRICES_VECTOR(i, 1) = DATA_MATRIX(i, 7)
Next i

j = 0
H1 = Round(MATRIX_ELEMENTS_MAX_FUNC(PRICES_VECTOR, 0) * (1 + CONFIDENCE_VAL), 0)
L1 = Round(MATRIX_ELEMENTS_MIN_FUNC(PRICES_VECTOR, 0) * (1 - CONFIDENCE_VAL), 0)

ReDim DATA_MATRIX(1 To 1 + H1 - L1, 1 To 2)
ReDim HIGH_MATRIX(1 To 1, 1 To 2)
ReDim LOW_MATRIX(1 To 1, 1 To 2)

For i = 1 To 1 + H1 - L1
    DATA_MATRIX(i, 1) = H1 - j
    j = j + 1
Next i


H2 = HIGH_VECTOR(1, 1)
L2 = LOW_VECTOR(1, 1)

LOW_MATRIX(1, 1) = L2  ' current Low
HIGH_MATRIX(1, 1) = H2  ' current High

For i = 1 + H1 - H2 To 1 + H1 - L2
    DATA_MATRIX(i, 2) = FIRST_STR
Next i

ROW_VAL = i - 1   ' last row
COL_VAL = 1 ' current column

MODE_STR = FIRST_STR  ' current mode is O
For ii = 2 To NROWS
    Call PnF_DO_FUNC(ii, DATA_MATRIX, LOW_MATRIX, _
                HIGH_MATRIX, LOW_VECTOR, HIGH_VECTOR, _
                ROW_VAL, COL_VAL, MODE_STR, REVERSE_VAL, _
                BOX_VAL, INC_FACTOR, FIRST_STR, SECOND_STR)
Next ii

'------------------------------SOME HOUSE KEEPING----------------------------
DATA_MATRIX = MATRIX_TRIM_FUNC(DATA_MATRIX, 1, "")
HIGH_MATRIX = MATRIX_TRANSPOSE_FUNC(VECTOR_TRIM_FUNC(MATRIX_TRANSPOSE_FUNC(HIGH_MATRIX), ""))
LOW_MATRIX = MATRIX_TRANSPOSE_FUNC(VECTOR_TRIM_FUNC(MATRIX_TRANSPOSE_FUNC(LOW_MATRIX), ""))
'----------------------------------------------------------------------------

For j = LBound(DATA_MATRIX, 2) To UBound(DATA_MATRIX, 2)
    For i = LBound(DATA_MATRIX, 1) To UBound(DATA_MATRIX, 1)
        If IsEmpty(DATA_MATRIX(i, j)) = True Then: DATA_MATRIX(i, j) = ""
    Next i
Next j

Select Case OUTPUT
    Case 0
        ASSET_TA_PnF_FUNC = DATA_MATRIX 'PERFECT
    Case 1
        ASSET_TA_PnF_FUNC = HIGH_MATRIX 'PERFECT
    Case 2
        ASSET_TA_PnF_FUNC = LOW_MATRIX 'PERFECT
    Case Else
        ASSET_TA_PnF_FUNC = Array(DATA_MATRIX, HIGH_MATRIX, LOW_MATRIX)
End Select

Exit Function
ERROR_LABEL:
ASSET_TA_PnF_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************



'************************************************************************************
'************************************************************************************
'FUNCTION      : PnF_DO_FUNC
'DESCRIPTION   :
'LIBRARY       : FINAN_ASSET
'GROUP         : HIGH-LOW
'ID            : 007



'AUTHOR        : RAFAEL NICOLAS FERMIN COTA

'************************************************************************************
'************************************************************************************

Private Function PnF_DO_FUNC(ByVal ii As Integer, _
ByRef DATA_MATRIX As Variant, _
ByRef LOW_MATRIX As Variant, _
ByRef HIGH_MATRIX As Variant, _
ByRef LOW_VECTOR As Variant, _
ByRef HIGH_VECTOR As Variant, _
ByRef ROW_VAL As Integer, _
ByRef COL_VAL As Integer, _
ByRef MODE_STR As String, _
Optional ByVal REVERSE_VAL As Integer = 3, _
Optional ByVal BOX_VAL As Integer = 1, _
Optional ByVal INC_FACTOR As Double = 1, _
Optional ByVal FIRST_STR As String = "O", _
Optional ByVal SECOND_STR As String = "X")

Dim NSIZE As Integer

Dim k As Integer
Dim kk As Integer

Dim TEMP_BOX As Integer
Dim TEMP_REVERSE As Integer
Dim TEMP_ROW As Integer

Dim TEMP_COL As Integer
Dim TEMP_LOW As Integer

Dim TEMP_HIGH As Integer
Dim TEMP_MODE As String

On Error GoTo ERROR_LABEL

PnF_DO_FUNC = True

NSIZE = UBound(DATA_MATRIX, 1)

TEMP_MODE = MODE_STR
TEMP_BOX = BOX_VAL
TEMP_REVERSE = REVERSE_VAL

TEMP_ROW = ROW_VAL   ' current TEMP_ROW
TEMP_COL = COL_VAL   ' current column

ReDim Preserve DATA_MATRIX(1 To NSIZE, 1 To TEMP_COL + 2)
ReDim Preserve HIGH_MATRIX(1 To 1, 1 To TEMP_COL + 1)
ReDim Preserve LOW_MATRIX(1 To 1, 1 To TEMP_COL + 1)

TEMP_LOW = LOW_MATRIX(1, TEMP_COL) ' previous TEMP_LOW
TEMP_HIGH = HIGH_MATRIX(1, TEMP_COL) ' previous TEMP_HIGH

If TEMP_MODE = FIRST_STR Then
    kk = TEMP_LOW - LOW_VECTOR(ii, 1)     ' Check kk in Lows
    If kk >= TEMP_BOX Then                   ' Does it drop by $TEMP_BOX?
        For k = 1 To kk                 ' If so, add some Os
            DATA_MATRIX(TEMP_ROW + k, TEMP_COL + 1) = FIRST_STR
        Next k
        ROW_VAL = TEMP_ROW + kk          ' display current TEMP_ROW
        LOW_MATRIX(1, TEMP_COL) = LOW_VECTOR(ii, 1) ' display new TEMP_LOW
        Exit Function
    End If
    
    ' Check if TEMP_HIGH increases by $TEMP_REVERSE
    kk = HIGH_VECTOR(ii, 1) - TEMP_LOW
    If kk >= TEMP_REVERSE Then ' kk TEMP_MODE
        TEMP_MODE = SECOND_STR
        MODE_STR = TEMP_MODE
        Call PnF_REVERSE_FUNC(ii, DATA_MATRIX, LOW_MATRIX, _
                HIGH_MATRIX, LOW_VECTOR, HIGH_VECTOR, _
                ROW_VAL, COL_VAL, MODE_STR, REVERSE_VAL, _
                BOX_VAL, INC_FACTOR, FIRST_STR, SECOND_STR)    ' Switch to X-TEMP_MODE
        Exit Function
    End If
End If

If TEMP_MODE = SECOND_STR Then
    kk = HIGH_VECTOR(ii, 1) - TEMP_HIGH    ' Check kk in Highs
    If kk >= TEMP_BOX Then                   ' Does it increase by $TEMP_BOX?
        For k = 1 To kk                 ' If so, add some Xs
            DATA_MATRIX(TEMP_ROW - k, TEMP_COL + 1) = SECOND_STR
        Next k
       ROW_VAL = TEMP_ROW - kk           ' display current TEMP_ROW
       HIGH_MATRIX(1, TEMP_COL) = HIGH_VECTOR(ii, 1) ' display new TEMP_HIGH
       Exit Function
    End If
    ' Check if TEMP_LOW decreases by $TEMP_REVERSE
    kk = TEMP_HIGH - LOW_VECTOR(ii, 1)
    If kk >= TEMP_REVERSE Then
        TEMP_MODE = FIRST_STR
        MODE_STR = TEMP_MODE
        Call PnF_REVERSE_FUNC(ii, DATA_MATRIX, LOW_MATRIX, _
                HIGH_MATRIX, LOW_VECTOR, HIGH_VECTOR, _
                ROW_VAL, COL_VAL, MODE_STR, REVERSE_VAL, _
                BOX_VAL, INC_FACTOR, FIRST_STR, SECOND_STR)  ' Switch to O-TEMP_MODE
        Exit Function
    End If
End If

PnF_DO_FUNC = False

Exit Function
ERROR_LABEL:
PnF_DO_FUNC = False
End Function

'************************************************************************************
'************************************************************************************



'************************************************************************************
'************************************************************************************
'FUNCTION      : PnF_REVERSE_FUNC
'DESCRIPTION   :
'LIBRARY       : FINAN_ASSET
'GROUP         : HIGH-LOW
'ID            : 008



'AUTHOR        : RAFAEL NICOLAS FERMIN COTA

'************************************************************************************
'************************************************************************************

Private Function PnF_REVERSE_FUNC(ByVal ii As Integer, _
ByRef DATA_MATRIX As Variant, _
ByRef LOW_MATRIX As Variant, _
ByRef HIGH_MATRIX As Variant, _
ByRef LOW_VECTOR As Variant, _
ByRef HIGH_VECTOR As Variant, _
ByRef ROW_VAL As Integer, _
ByRef COL_VAL As Integer, _
ByRef MODE_STR As String, _
Optional ByVal REVERSE_VAL As Integer = 3, _
Optional ByVal BOX_VAL As Integer = 1, _
Optional ByVal INC_FACTOR As Double = 1, _
Optional ByVal FIRST_STR As String = "O", _
Optional ByVal SECOND_STR As String = "X")

Dim k As Integer
Dim kk As Integer

Dim TEMP_ROW As Integer
Dim TEMP_COL As Integer
Dim TEMP_LOW As Integer
Dim TEMP_HIGH As Integer

Dim TEMP_MODE As String

On Error GoTo ERROR_LABEL

PnF_REVERSE_FUNC = False

TEMP_MODE = MODE_STR
TEMP_ROW = ROW_VAL   ' current TEMP_ROW
TEMP_COL = COL_VAL   ' current column
TEMP_HIGH = HIGH_MATRIX(1, TEMP_COL) ' previous TEMP_HIGH
TEMP_LOW = LOW_MATRIX(1, TEMP_COL) ' previous TEMP_LOW

COL_VAL = COL_VAL + 1   ' kk column
TEMP_COL = COL_VAL               ' display new column

If TEMP_MODE = FIRST_STR Then
    kk = TEMP_HIGH - LOW_VECTOR(ii, 1) ' Get height of Os
    For k = 1 To kk                 ' If so, add some Os
        DATA_MATRIX(TEMP_ROW + k, TEMP_COL + 1) = FIRST_STR
    Next k
    ROW_VAL = TEMP_ROW + kk       ' current TEMP_ROW
    HIGH_MATRIX(1, TEMP_COL) = HIGH_MATRIX(1, TEMP_COL - 1) - 1 ' display previous TEMP_HIGH
    LOW_MATRIX(1, TEMP_COL) = LOW_VECTOR(ii, 1)     ' display new TEMP_LOW
End If

If TEMP_MODE = SECOND_STR Then
    kk = HIGH_VECTOR(ii, 1) - TEMP_LOW        ' Get height of Xs
    For k = 1 To kk                          ' Add some Xs
        DATA_MATRIX(TEMP_ROW - k, TEMP_COL + 1) = SECOND_STR
    Next k
    ROW_VAL = TEMP_ROW - kk          ' current TEMP_ROW
    HIGH_MATRIX(1, TEMP_COL) = HIGH_VECTOR(ii, 1)       ' display new TEMP_HIGH
    LOW_MATRIX(1, TEMP_COL) = LOW_MATRIX(1, TEMP_COL - 1) + 1 ' display previous TEMP_LOW
End If

PnF_REVERSE_FUNC = True

Exit Function
ERROR_LABEL:
PnF_REVERSE_FUNC = False
End Function


'Point & Figure Charts
'Here 's how it works:
'Each day we look at the High and Low for the day.
'For the first day, we enter a number of Os from the High to the Low.
'Example: High=$27 to Low=$24
'The next day we see how much the Low has gone lower.
'If the Low did decrease, but by less than $1, we do nothing.
'If the decrease is greater than $1, we add another set of Os, down to that new Low.
'Example: Low decreased to $22 so we add 2 more Os
'We continue doing nothing or adding Os ... so long as the Low is decreasing.
'If the Lows do not decrease, only then do we look at the High for the day
'... to see if the High has increased by at least $3.
'If not, we again do nothing.

'You spend a lot of time doing nothing!
'The objective is to follow the Lows as they go down and ignore any local ups or down that aren't significant.
'We ignore decreases in Lows (or increases in Highs) that are less than $1.
'Okay, so we keep adding Os ... but what about when the Highs start going up significantly?
'Aah, then we start a new column of Xs.
'For our example (where the bottom O was at $22), the High would have to be at least $22 + $3 =
'$25 for us to start that new column .... else we'd do nothing.
'If the Lows stop decreasing and the High increases by at least $3 we add a new column of Xs.
'Example: The Lows stop decreasing and the High goes up to $26
'We now continue to watch the Highs, so long as they are increasing, adding more Xs to the top of
'the column each time the High increases by at least $1.

'I thought $3 was for the Highs and $1 for the Lows.
'No, we continue with either the Lows or Highs as long as the changes, day-to-day, are at least $1.
'That $1 is called the box size.

'We switch from O to Xs (starting a new column) when the Lows stop decreasing and the High increases by
'at least $3. That $3 is called the reversal size.

'We switch from Xs to Os when the Highs stop increasing and the Lows decrease by ... uh ...
'By at least the reversal size.

'Notice some interesting things:
'1. There are never less than 3 Os or Xs in each column, if the reversal
'size is 3.
'2. Minor changes are ignored. A reversal must be "significant" in order to incorporate
'it into our P&F chart.
'3. The horizontal axis has nothing to do with the time. It does not count days ... but the
'number of times we get a reversal.
'4. Uptrends, downtrends, support and resistance levels are easily recognized ... so thay say.
'5. The (sometimes) misleading effect of time is ...

'Why would anybody want to use a Factor other than "1"?
'Maybe You 'd use Factor = 0.1 if you're lookin' at the S&P500
'that 'd give a P&F chart with smaller numbers, for, say the S&P500.
'Huh? Why smaller numbers? If they're small enough you can use your fingers to keep
'track and ... Just remember that, with Factor = 0.1, a Box and Reverse of 1 and 3
'really means changes of 10 and 30 ... and that'd be a good idea for the S&P, eh?

'For a 10,000 DOW, you might want to use Factor = 0.01 so Box = 1 means changes of 100
'(or 1%), like so It ain 't days ... it's reversals.

'The main benefit is that support and resistance levels jump right out at you (at least they should) because PnF filters out so much noise.
'By adjusting the box size and reversal amount, you can visually determine trading opportunities.
'For instance you might decide that with a certain box size and reversal amount that 80% of breakouts result in a profit of 4 boxes.
'You can set your initial stop on breakout to be the reversal amount (say 3 boxes).
'Then trail the stop as the trade moves in your favor, using the same reversal amount as the stop distance.
'The possibilities are endless.

'References:
'http://208.234.169.12/vb/forumdisplay.php?s=11df9000f1b4cebaefddfdc7cec629a4&forumid=33&daysprune=
'http://www.stockcharts.com/support/pnfCharts.html
'http://www.incrediblecharts.com/technical/point_and_figure.htm
'http://www.investorsintelligence.com/x/using_point_and_figure_charts.html
'http://dorseywright.com/cgi-bin/foxweb.exe/fwuniv
'http://www.gummy-stuff.org/PandF-charts.htm
