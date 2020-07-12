Attribute VB_Name = "FINAN_PORT_WEIGHTS_ROTAT_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.
                    
'************************************************************************************
'************************************************************************************
'FUNCTION      : PORT_ASSET_ROTATION_FUNC

'DESCRIPTION   :
'LIBRARY       : PORTFOLIO
'GROUP         : TRADE_SWITCH
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 21/01/2009

'REFERENCES    :
'http://www.gummy-stuff.org/rotation.htm
'http://www.gummy-stuff.org/sector-rotation.htm
'http://www.gummy-stuff.org/sector-math.htm
'http://www.gummy-stuff.org/sector-explain.htm
'http://www.gummy-stuff.org/sectors.htm
'http://www.gummy-stuff.org/sector-returns.htm
'************************************************************************************
'************************************************************************************

Function PORT_ASSET_ROTATION_FUNC(ByRef DATA_RNG As Variant, _
ByVal INITIAL_CASH_VAL As Double, _
Optional ByRef PREFER_ASSET_INDEX_VAL As Long = 2, _
Optional ByRef MA_PERIODS As Double = 8, _
Optional ByRef SWITCH_FACTOR As Double = 18, _
Optional ByRef VERSION As Integer = 1, _
Optional ByRef OUTPUT As Integer = 0)

'VERION = + 1: Choose the sector which has the largest gMA, meaning the sector whose price
'is farthest above the Moving Average. That's assuming there's a upward trend with that sector
'and we want to be in on that trend. Indeed, since the gMA is the rate of increase of that
'Weighted Moving Average, we want the sector with the largest Rate of Increase.

'VERION = - 1: Choose the sector whose price is farthest below its Moving Average. That's
'consistent with Buy Low. We anticipate that we'll being buying a bargain.

'Note: gMA = CurrentPrice - MovingAverage and you switch from your Favourite only

Dim i As Long
Dim j As Long
Dim k As Long

Dim NROWS As Long
Dim NCOLUMNS As Long

Dim TEMP_SUM As Double
Dim MAX_VAL As Double

Dim DATA_MATRIX As Variant
Dim DATES_VECTOR As Variant
Dim TICKERS_VECTOR As Variant

Dim TEMP1_MATRIX As Variant
Dim TEMP2_MATRIX As Variant
Dim TEMP3_MATRIX As Variant
Dim TEMP4_MATRIX As Variant

Dim COLLECTION_OBJ As Collection

On Error Resume Next

If VERSION <> 1 Then: VERSION = -1

DATA_MATRIX = DATA_RNG
TICKERS_VECTOR = MATRIX_GET_ROW_FUNC(DATA_MATRIX, 1, 1)
TICKERS_VECTOR = MATRIX_REMOVE_COLUMNS_FUNC(TICKERS_VECTOR, 1, 1)

DATES_VECTOR = MATRIX_GET_COLUMN_FUNC(DATA_MATRIX, 1, 1)
DATES_VECTOR = MATRIX_REMOVE_ROWS_FUNC(DATES_VECTOR, 1, 1)

DATA_MATRIX = MATRIX_REMOVE_ROWS_FUNC(DATA_MATRIX, 1, 1)
DATA_MATRIX = MATRIX_REMOVE_COLUMNS_FUNC(DATA_MATRIX, 1, 1)

NROWS = UBound(DATA_MATRIX, 1)
NCOLUMNS = UBound(DATA_MATRIX, 2)

If MA_PERIODS > NROWS Then: GoTo ERROR_LABEL
If MA_PERIODS < 2 Then: GoTo ERROR_LABEL

If PREFER_ASSET_INDEX_VAL > NCOLUMNS Then: GoTo ERROR_LABEL
If PREFER_ASSET_INDEX_VAL < 1 Then: GoTo ERROR_LABEL

Set COLLECTION_OBJ = New Collection
For j = 1 To NCOLUMNS - 1
    Call COLLECTION_OBJ.Add(CStr(j), CStr(TICKERS_VECTOR(1, j)))
Next j

ReDim TEMP1_MATRIX(1 To NROWS, 1 To NCOLUMNS)
ReDim TEMP2_MATRIX(1 To NROWS, 1 To NCOLUMNS)

For i = 1 To NROWS
    For j = 1 To NCOLUMNS
        If DATA_MATRIX(i, j) <> 0 Then
            TEMP1_MATRIX(i, j) = DATA_MATRIX(i, j) / DATA_MATRIX(1, j)
        Else
            TEMP1_MATRIX(i, j) = DATA_MATRIX(i - 1, j) / DATA_MATRIX(1, j)
        End If
        TEMP2_MATRIX(i, j) = TEMP1_MATRIX(i, j) * INITIAL_CASH_VAL
    Next j
Next i

ReDim TEMP3_MATRIX(1 To NROWS, 1 To NCOLUMNS)
For j = 1 To NCOLUMNS
    TEMP_SUM = 0
    For i = 1 To NROWS
        TEMP_SUM = TEMP_SUM + TEMP1_MATRIX(i, j)
        If i > MA_PERIODS + 1 Then
            TEMP_SUM = TEMP_SUM - TEMP1_MATRIX(i - MA_PERIODS - 1, j)
            TEMP3_MATRIX(i, j) = TEMP1_MATRIX(i, j) - (TEMP_SUM / (MA_PERIODS + 1))
        Else
            TEMP3_MATRIX(i, j) = TEMP1_MATRIX(i, j) - (TEMP_SUM / i)
        End If
    Next i
Next j

i = 0
ReDim TEMP4_MATRIX(i To NROWS, 1 To 10)
TEMP4_MATRIX(i, 1) = ("DATE")
TEMP4_MATRIX(i, 2) = TICKERS_VECTOR(1, PREFER_ASSET_INDEX_VAL) 'FAVOURITE GMA
TEMP4_MATRIX(i, 3) = "MAX - GMA"
TEMP4_MATRIX(i, 4) = ("CURRENT ASSET")
TEMP4_MATRIX(i, 5) = ("NEW PRICE")
TEMP4_MATRIX(i, 6) = ("OLD PRICE")
TEMP4_MATRIX(i, 7) = ("DOLLARS AVAILABLE")
TEMP4_MATRIX(i, 8) = ("UNITS")
TEMP4_MATRIX(i, 9) = ("PORTFOLIO")
TEMP4_MATRIX(i, 10) = ("GAINS")

i = 1
TEMP4_MATRIX(i, 1) = DATES_VECTOR(i, 1)
TEMP4_MATRIX(i, 2) = TEMP3_MATRIX(i, PREFER_ASSET_INDEX_VAL)
GoSub MAX_LINE: TEMP4_MATRIX(i, 3) = MAX_VAL
TEMP4_MATRIX(i, 4) = TEMP4_MATRIX(i - 1, 2)
k = 0: k = CLng(COLLECTION_OBJ(CStr(TEMP4_MATRIX(i, 4))))
If k = 0 Then: k = NCOLUMNS
TEMP4_MATRIX(i, 5) = TEMP2_MATRIX(i, k)
TEMP4_MATRIX(i, 6) = TEMP2_MATRIX(i, k)
TEMP4_MATRIX(i, 7) = 0
TEMP4_MATRIX(i, 8) = INITIAL_CASH_VAL / TEMP4_MATRIX(i, 5)
TEMP4_MATRIX(i, 9) = INITIAL_CASH_VAL
TEMP4_MATRIX(i, 10) = ""

'------------------------------------------------------------------------------------------------------------------
For i = 2 To NROWS
'------------------------------------------------------------------------------------------------------------------
    TEMP4_MATRIX(i, 1) = DATES_VECTOR(i, 1)
    TEMP4_MATRIX(i, 2) = TEMP3_MATRIX(i, PREFER_ASSET_INDEX_VAL)
    GoSub MAX_LINE: TEMP4_MATRIX(i, 3) = MAX_VAL
    If VERSION * TEMP4_MATRIX(i, 3) > VERSION * SWITCH_FACTOR * TEMP4_MATRIX(i, 2) Then
        For j = 1 To NCOLUMNS
            If TEMP3_MATRIX(i, j) = TEMP4_MATRIX(i, 3) Then: Exit For
        Next j
        TEMP4_MATRIX(i, 4) = TICKERS_VECTOR(1, j)
    Else
        TEMP4_MATRIX(i, 4) = TEMP4_MATRIX(i - 1, 4)
    End If
    
    k = 0: k = CLng(COLLECTION_OBJ(CStr(TEMP4_MATRIX(i, 4))))
    If k = 0 Then: k = NCOLUMNS
    TEMP4_MATRIX(i, 5) = TEMP2_MATRIX(i, k)
    
    k = 0: k = CLng(COLLECTION_OBJ(CStr(TEMP4_MATRIX(i - 1, 4))))
    If k = 0 Then: k = NCOLUMNS
    TEMP4_MATRIX(i, 6) = TEMP2_MATRIX(i, k)
    TEMP4_MATRIX(i, 7) = TEMP4_MATRIX(i, 6) * TEMP4_MATRIX(i - 1, 8)
    
    If (TEMP4_MATRIX(i, 4) <> TEMP4_MATRIX(i - 1, 4)) And (TEMP4_MATRIX(i, 5) <> 0) Then
        TEMP4_MATRIX(i, 8) = TEMP4_MATRIX(i, 7) / TEMP4_MATRIX(i, 5)
    Else
        TEMP4_MATRIX(i, 8) = TEMP4_MATRIX(i - 1, 8)
    End If
    
    If (TEMP4_MATRIX(i, 4) <> TEMP4_MATRIX(i - 1, 4)) And (TEMP4_MATRIX(i, 5) <> 0) Then
        TEMP4_MATRIX(i, 8) = TEMP4_MATRIX(i, 7) / TEMP4_MATRIX(i, 5)
    Else
        TEMP4_MATRIX(i, 8) = TEMP4_MATRIX(i - 1, 8)
    End If
    
    TEMP4_MATRIX(i, 9) = TEMP4_MATRIX(i, 8) * TEMP4_MATRIX(i, 5)
    If (TEMP4_MATRIX(i - 1, 9) <> 0) Then
        TEMP4_MATRIX(i, 10) = TEMP4_MATRIX(i, 9) / TEMP4_MATRIX(i - 1, 9) - 1
    Else
        TEMP4_MATRIX(i, 10) = TEMP4_MATRIX(i, 9)
    End If
1983:
'------------------------------------------------------------------------------------------------------------------
Next i
'------------------------------------------------------------------------------------------------------------------

'----------------------------------------------------------------------------------------------------------------
Select Case OUTPUT
'----------------------------------------------------------------------------------------------------------------
Case 0
'----------------------------------------------------------------------------------------------------------------
    PORT_ASSET_ROTATION_FUNC = TEMP4_MATRIX 'Sector Switching Summary
'----------------------------------------------------------------------------------------------------------------
Case 1
'----------------------------------------------------------------------------------------------------------------
    PORT_ASSET_ROTATION_FUNC = TEMP3_MATRIX 'gMA
'----------------------------------------------------------------------------------------------------------------
Case 2
'----------------------------------------------------------------------------------------------------------------
    PORT_ASSET_ROTATION_FUNC = TEMP2_MATRIX ' Portfolio $Growth
'----------------------------------------------------------------------------------------------------------------
Case 3
'----------------------------------------------------------------------------------------------------------------
    PORT_ASSET_ROTATION_FUNC = TEMP1_MATRIX ' Asset Growth
'----------------------------------------------------------------------------------------------------------------
Case Else
'----------------------------------------------------------------------------------------------------------------
    ReDim DATA_MATRIX(0 To NROWS, 1 To NCOLUMNS)
    For j = 1 To NCOLUMNS
        DATA_MATRIX(0, j) = TICKERS_VECTOR(1, j) 'Charting Summary
        For i = 1 To NROWS
            DATA_MATRIX(i, j) = IIf(DATA_MATRIX(0, j) = TEMP4_MATRIX(i, 4), 0.5, 0)
        Next i
    Next j
    If OUTPUT = 4 Then
        PORT_ASSET_ROTATION_FUNC = DATA_MATRIX
    Else
        PORT_ASSET_ROTATION_FUNC = Array(TEMP4_MATRIX, TEMP3_MATRIX, TEMP2_MATRIX, TEMP1_MATRIX, DATA_MATRIX)
    End If
'----------------------------------------------------------------------------------------------------------------
End Select
'----------------------------------------------------------------------------------------------------------------
Exit Function
'----------------------------------------------------------------------------------------------------------------
MAX_LINE:
'----------------------------------------------------------------------------------------------------------------
    MAX_VAL = -2 ^ 52
    For j = 1 To NCOLUMNS
        If TEMP3_MATRIX(i, j) > MAX_VAL Then: MAX_VAL = TEMP3_MATRIX(i, j)
    Next j
'----------------------------------------------------------------------------------------------------------------
Return
'----------------------------------------------------------------------------------------------------------------
ERROR_LABEL:
PORT_ASSET_ROTATION_FUNC = Err.number
End Function
