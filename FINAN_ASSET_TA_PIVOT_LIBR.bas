Attribute VB_Name = "FINAN_ASSET_TA_PIVOT_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.


'************************************************************************************
'************************************************************************************
'FUNCTION      : ASSETS_PIVOT_POINTS_FUNC
'DESCRIPTION   :
'http://www.gummy-stuff.org/pivot-points.htm
'http://www.investopedia.com/articles/technical/04/041404.asp

'LIBRARY       : FINAN_ASSET_TA
'GROUP         : PIVOT
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'************************************************************************************
'************************************************************************************

Function ASSETS_PIVOT_POINTS_FUNC(ByRef TICKERS_RNG As Variant, _
Optional ByVal REFRESH_CALLER As Variant, _
Optional ByVal SERVER_STR As String = "UNITED STATES")

Dim i As Long
Dim j As Long
Dim NROWS As Long

Dim TEMP_MATRIX As Variant

Dim DATA_MATRIX As Variant
Dim TICKERS_VECTOR As Variant

On Error GoTo ERROR_LABEL

If IsArray(TICKERS_RNG) = True Then
    TICKERS_VECTOR = TICKERS_RNG
    If UBound(TICKERS_VECTOR, 1) = 1 Then
        TICKERS_VECTOR = MATRIX_TRANSPOSE_FUNC(TICKERS_VECTOR)
    End If
Else
    ReDim TICKERS_VECTOR(1 To 1, 1 To 1)
    TICKERS_VECTOR(1, 1) = TICKERS_RNG
End If

'nd1vhg l1
NROWS = UBound(TICKERS_VECTOR)

ReDim TEMP_MATRIX(1 To 1, 1 To 9)
TEMP_MATRIX(1, 1) = "Name"
TEMP_MATRIX(1, 2) = "Symbol"
TEMP_MATRIX(1, 3) = "time of last trade"
TEMP_MATRIX(1, 4) = "Volume"
TEMP_MATRIX(1, 5) = "High"
TEMP_MATRIX(1, 6) = "Low"
TEMP_MATRIX(1, 7) = "Last Trade"

DATA_MATRIX = YAHOO_QUOTES_FUNC(TICKERS_VECTOR, TEMP_MATRIX, _
              REFRESH_CALLER, False, SERVER_STR)

If IsArray(DATA_MATRIX) = False Then: GoTo ERROR_LABEL

ReDim TEMP_MATRIX(0 To NROWS, 1 To 14)
For j = 1 To 14: For i = 0 To NROWS: _
TEMP_MATRIX(i, j) = "": Next i: Next j

TEMP_MATRIX(0, 1) = "Name"
TEMP_MATRIX(0, 2) = "Symbol"
TEMP_MATRIX(0, 3) = "Time"
TEMP_MATRIX(0, 4) = "Volume"

TEMP_MATRIX(0, 5) = "Day High"
TEMP_MATRIX(0, 6) = "Day Low"
TEMP_MATRIX(0, 7) = "Close"
TEMP_MATRIX(0, 8) = "Pivot"

TEMP_MATRIX(0, 9) = "R1"
TEMP_MATRIX(0, 10) = "R2"
TEMP_MATRIX(0, 11) = "R3"
TEMP_MATRIX(0, 12) = "S1"
TEMP_MATRIX(0, 13) = "S2"
TEMP_MATRIX(0, 14) = "S3"

'-------------------------------------------------------------------------------
For i = 1 To NROWS
'-------------------------------------------------------------------------------
    For j = 1 To 7: TEMP_MATRIX(i, j) = DATA_MATRIX(i, j): Next j
    TEMP_MATRIX(i, 8) = (TEMP_MATRIX(i, 5) + TEMP_MATRIX(i, 6) + TEMP_MATRIX(i, 7)) / 3
    
    TEMP_MATRIX(i, 9) = 2 * TEMP_MATRIX(i, 8) - TEMP_MATRIX(i, 6)
    TEMP_MATRIX(i, 12) = 2 * TEMP_MATRIX(i, 8) - TEMP_MATRIX(i, 5)
    
    TEMP_MATRIX(i, 10) = TEMP_MATRIX(i, 8) + (TEMP_MATRIX(i, 9) - TEMP_MATRIX(i, 12))
    TEMP_MATRIX(i, 13) = TEMP_MATRIX(i, 8) - (TEMP_MATRIX(i, 9) - TEMP_MATRIX(i, 12))
    
    TEMP_MATRIX(i, 11) = TEMP_MATRIX(i, 5) + 2 * (TEMP_MATRIX(i, 8) - TEMP_MATRIX(i, 6))
    TEMP_MATRIX(i, 14) = TEMP_MATRIX(i, 6) - 2 * (TEMP_MATRIX(i, 5) - TEMP_MATRIX(i, 8))
'-------------------------------------------------------------------------------
Next i
'-------------------------------------------------------------------------------

ASSETS_PIVOT_POINTS_FUNC = TEMP_MATRIX

Exit Function
ERROR_LABEL:
ASSETS_PIVOT_POINTS_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : ASSET_PIVOT_POINTS_FUNC
'DESCRIPTION   :
'http://www.gummy-stuff.org/pivot-points.htm
'http://www.investopedia.com/articles/technical/04/041404.asp

'LIBRARY       : FINAN_ASSET_TA
'GROUP         : PIVOT
'ID            : 002
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'************************************************************************************
'************************************************************************************

Function ASSET_PIVOT_POINTS_FUNC(ByRef TICKER_STR As Variant, _
Optional ByVal START_DATE As Date, _
Optional ByVal END_DATE As Date, _
Optional ByVal PIVOT_INDEX As Integer = 1)

Dim i As Long
Dim j As Long
Dim NROWS As Long
Dim PIVOT_VAL As Double
Dim TEMP_MATRIX As Variant
Dim DATA_MATRIX As Variant

On Error GoTo ERROR_LABEL

If IsArray(TICKER_STR) = True Then
    DATA_MATRIX = TICKER_STR
Else
    DATA_MATRIX = _
    YAHOO_HISTORICAL_DATA_SERIE_FUNC(TICKER_STR, START_DATE, END_DATE, _
    "d", "DOHLCV", False, False, True)
End If
NROWS = UBound(DATA_MATRIX, 1)

ReDim TEMP_MATRIX(0 To NROWS, 1 To 13)
TEMP_MATRIX(0, 1) = "DATE"
TEMP_MATRIX(0, 2) = "OPEN"
TEMP_MATRIX(0, 3) = "HIGH"
TEMP_MATRIX(0, 4) = "LOW"
TEMP_MATRIX(0, 5) = "CLOSE"
TEMP_MATRIX(0, 6) = "VOLUME"
TEMP_MATRIX(0, 7) = "PIVOT"
TEMP_MATRIX(0, 8) = "R1"
TEMP_MATRIX(0, 9) = "R2"
TEMP_MATRIX(0, 10) = "R3"
TEMP_MATRIX(0, 11) = "S1"
TEMP_MATRIX(0, 12) = "S2"
TEMP_MATRIX(0, 13) = "S3"

'-----------------------------------------------------------------------------------------------
Select Case PIVOT_INDEX
'-----------------------------------------------------------------------------------------------
Case 0
'-----------------------------------------------------------------------------------------------
    i = 1
    For j = 1 To 6: TEMP_MATRIX(i, j) = DATA_MATRIX(i, j): Next j
    TEMP_MATRIX(i, 6) = TEMP_MATRIX(i, 6) / 1000
    
    TEMP_MATRIX(i, 7) = (TEMP_MATRIX(i, 3) + TEMP_MATRIX(i, 4) + TEMP_MATRIX(i, 5)) / 3
    TEMP_MATRIX(i, 8) = 2 * TEMP_MATRIX(i, 7) - TEMP_MATRIX(i, 4)
    TEMP_MATRIX(i, 11) = 2 * TEMP_MATRIX(i, 7) - TEMP_MATRIX(i, 3)
    TEMP_MATRIX(i, 9) = TEMP_MATRIX(i, 7) + (TEMP_MATRIX(i, 3) - TEMP_MATRIX(i, 4))
    TEMP_MATRIX(i, 12) = TEMP_MATRIX(i, 7) - (TEMP_MATRIX(i, 3) - TEMP_MATRIX(i, 4))
    TEMP_MATRIX(i, 10) = TEMP_MATRIX(i, 3) + 2 * (TEMP_MATRIX(i, 7) - TEMP_MATRIX(i, 4))
    TEMP_MATRIX(i, 13) = TEMP_MATRIX(i, 4) - 2 * (TEMP_MATRIX(i, 3) - TEMP_MATRIX(i, 7))
    
    For i = 2 To NROWS
        For j = 1 To 6: TEMP_MATRIX(i, j) = DATA_MATRIX(i, j): Next j
        TEMP_MATRIX(i, 6) = TEMP_MATRIX(i, 6) / 1000
        
        TEMP_MATRIX(i, 7) = (TEMP_MATRIX(i, 3) + TEMP_MATRIX(i, 4) + TEMP_MATRIX(i, 5)) / 3
        
        If TEMP_MATRIX(i, 5) > TEMP_MATRIX(i - 1, 8) Then
            TEMP_MATRIX(i, 8) = 2 * TEMP_MATRIX(i, 7) - TEMP_MATRIX(i, 4)
        Else
            If TEMP_MATRIX(i, 5) < TEMP_MATRIX(i - 1, 11) Then
                TEMP_MATRIX(i, 8) = 2 * TEMP_MATRIX(i, 7) - TEMP_MATRIX(i, 4)
            Else
                TEMP_MATRIX(i, 8) = TEMP_MATRIX(i - 1, 8)
            End If
        End If
        
        TEMP_MATRIX(i, 9) = TEMP_MATRIX(i - 1, 9)
        
        If (TEMP_MATRIX(i, 5) > TEMP_MATRIX(i - 1, 10) Or TEMP_MATRIX(i, 5) < TEMP_MATRIX(i - 1, 13)) Then
            TEMP_MATRIX(i, 10) = TEMP_MATRIX(i, 3) + 2 * (TEMP_MATRIX(i, 7) - TEMP_MATRIX(i, 4))
        Else
            TEMP_MATRIX(i, 10) = TEMP_MATRIX(i - 1, 10)
        End If
        
        TEMP_MATRIX(i, 11) = TEMP_MATRIX(i - 1, 11)
        TEMP_MATRIX(i, 12) = TEMP_MATRIX(i - 1, 12)
        
        If (TEMP_MATRIX(i, 5) > TEMP_MATRIX(i - 1, 10) Or TEMP_MATRIX(i, 5) < TEMP_MATRIX(i - 1, 13)) Then
            TEMP_MATRIX(i, 13) = TEMP_MATRIX(i, 4) - 2 * (TEMP_MATRIX(i, 3) - TEMP_MATRIX(i, 7))
        Else
            TEMP_MATRIX(i, 13) = TEMP_MATRIX(i - 1, 13)
        End If
    Next i
'-----------------------------------------------------------------------------------------------
Case Else
'-----------------------------------------------------------------------------------------------
    i = PIVOT_INDEX
    If i > NROWS Then: i = NROWS
    PIVOT_VAL = (DATA_MATRIX(i, 3) + DATA_MATRIX(i, 4) + DATA_MATRIX(i, 5)) / 3
    TEMP_MATRIX(0, 7) = "PIVOT: " & Format(DATA_MATRIX(i, 1), "mmm dd, yyyy")

    i = 1
    For j = 1 To 6: TEMP_MATRIX(i, j) = DATA_MATRIX(i, j): Next j
    TEMP_MATRIX(i, 6) = TEMP_MATRIX(i, 6) / 1000
    
    TEMP_MATRIX(i, 7) = PIVOT_VAL
    TEMP_MATRIX(i, 8) = 2 * PIVOT_VAL - TEMP_MATRIX(i, 4)
    TEMP_MATRIX(i, 9) = PIVOT_VAL + (TEMP_MATRIX(i, 3) - TEMP_MATRIX(i, 4))
    TEMP_MATRIX(i, 10) = TEMP_MATRIX(i, 3) + 2 * (PIVOT_VAL - TEMP_MATRIX(i, 4))
    TEMP_MATRIX(i, 11) = 2 * PIVOT_VAL - TEMP_MATRIX(i, 3)
    TEMP_MATRIX(i, 12) = PIVOT_VAL - (TEMP_MATRIX(i, 3) - TEMP_MATRIX(i, 4))
    TEMP_MATRIX(i, 13) = TEMP_MATRIX(i, 4) - 2 * (TEMP_MATRIX(i, 3) - PIVOT_VAL)
    
    For i = 2 To NROWS
        For j = 1 To 6: TEMP_MATRIX(i, j) = DATA_MATRIX(i, j): Next j
        TEMP_MATRIX(i, 6) = TEMP_MATRIX(i, 6) / 1000
        TEMP_MATRIX(i, 7) = PIVOT_VAL
        For j = 8 To 13
            TEMP_MATRIX(i, j) = TEMP_MATRIX(i - 1, j)
        Next j
    Next i
'-----------------------------------------------------------------------------------------------
End Select
'-----------------------------------------------------------------------------------------------

ASSET_PIVOT_POINTS_FUNC = TEMP_MATRIX

Exit Function
ERROR_LABEL:
ASSET_PIVOT_POINTS_FUNC = Err.number
End Function
