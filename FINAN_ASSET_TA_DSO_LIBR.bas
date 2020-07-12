Attribute VB_Name = "FINAN_ASSET_TA_DSO_LIBR"


Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.


Function ASSET_DSO_FUNC(ByVal TICKER_STR As Variant, _
Optional ByVal START_DATE As Date, _
Optional ByVal END_DATE As Date, _
Optional ByVal OUTPUT As Integer = 0)

'Daily Stock Oscillation

Dim i As Long
Dim j As Long

Dim NROWS As Long
Dim NCOLUMNS As Long

Dim TEMP1_SUM As Double
Dim TEMP2_SUM As Double

Dim TEMP_MATRIX As Variant
Dim DATA_MATRIX As Variant

On Error GoTo ERROR_LABEL

If IsArray(TICKER_STR) = False Then
    DATA_MATRIX = YAHOO_HISTORICAL_DATA_SERIE_FUNC(TICKER_STR, START_DATE, END_DATE, "d", "DOHLCV", False, True, True)
Else
    DATA_MATRIX = TICKER_STR
End If
NROWS = UBound(DATA_MATRIX, 1)
NCOLUMNS = UBound(DATA_MATRIX, 2)

ReDim TEMP_MATRIX(0 To NROWS, 1 To 8)
TEMP_MATRIX(0, 1) = "DATE"
TEMP_MATRIX(0, 2) = "OPEN"
TEMP_MATRIX(0, 3) = "HIGH"
TEMP_MATRIX(0, 4) = "LOW"
TEMP_MATRIX(0, 5) = "CLOSE"
TEMP_MATRIX(0, 6) = "VOLUME"

TEMP_MATRIX(0, 7) = "(H-L)/O: "
'Daily Stock Activity (DSA) = (High - Low) / Open
'If it is 2.0%, it means the spread between High and Low was 2.0% of
'the Opening price ... and that'd be significant.
TEMP_MATRIX(0, 8) = "(H/O)*(C/L)-1: " '(High/Open)*(Close/Low) - 1

TEMP1_SUM = 0: TEMP2_SUM = 0
For i = 1 To NROWS
    For j = 1 To NCOLUMNS: TEMP_MATRIX(i, j) = DATA_MATRIX(i, j): Next j
    TEMP_MATRIX(i, 7) = (TEMP_MATRIX(i, 3) - TEMP_MATRIX(i, 4)) / TEMP_MATRIX(i, 2)
    TEMP1_SUM = TEMP1_SUM + TEMP_MATRIX(i, 7)
    TEMP_MATRIX(i, 8) = (TEMP_MATRIX(i, 3) / TEMP_MATRIX(i, 2)) * (TEMP_MATRIX(i, 5) / TEMP_MATRIX(i, 4)) - 1
    TEMP2_SUM = TEMP2_SUM + TEMP_MATRIX(i, 8)
Next i
TEMP1_SUM = TEMP1_SUM / NROWS
TEMP2_SUM = TEMP2_SUM / NROWS

TEMP_MATRIX(0, 7) = TEMP_MATRIX(0, 7) & Format(TEMP1_SUM, "0.00%")
TEMP_MATRIX(0, 8) = TEMP_MATRIX(0, 8) & Format(TEMP1_SUM, "0.00%")

Select Case OUTPUT
Case 0
    ASSET_DSO_FUNC = TEMP_MATRIX
Case Else 'Average DSO
    ASSET_DSO_FUNC = Array(TEMP1_SUM, TEMP2_SUM)
End Select

Exit Function
ERROR_LABEL:
ASSET_DSO_FUNC = Err.number
End Function


