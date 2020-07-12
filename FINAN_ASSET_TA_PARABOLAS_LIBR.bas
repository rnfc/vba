Attribute VB_Name = "FINAN_ASSET_TA_PARABOLAS_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.

'http://www.marketmasters.com.au/86.0.html
'http://www.gummy-stuff.org/parabolas.htm
'http://stockcharts.com/school/doku.php?id=chart_school:technical_indicators:parabolic_sar
'http://finance.yahoo.com/q/ta?s=GE&t=2y&l=on&z=m&q=l&p=p&a=&c=

Function ASSET_TA_PARABOLAS_FUNC(ByVal TICKER_STR As Variant, _
Optional ByVal START_DATE As Date, _
Optional ByVal END_DATE As Date, _
Optional ByVal NO_PERIODS As Double = 700, _
Optional ByVal SHIFT_FACTOR As Double = -4#, _
Optional ByVal VERSION As Integer = 1, _
Optional ByVal OUTPUT As Integer = 0)

'Base Parabola on first NO PERIODS
'IF VERSION = 1 Then --> CUBIC
'IF OUTPUT > 0 Then --> COEFFICIENTS

Dim i As Long
Dim j As Long

Dim NROWS As Long
Dim NCOLUMNS As Long

Dim TEMP_STR As String
Dim FACTOR_VAL As Double
Dim TEMP_VECTOR As Variant 'Weighted Prices/Coefficients
Dim TEMP_MATRIX As Variant
Dim DATA_MATRIX As Variant

On Error GoTo ERROR_LABEL

If IsArray(TICKER_STR) = False Then
    DATA_MATRIX = YAHOO_HISTORICAL_DATA_SERIE_FUNC(TICKER_STR, START_DATE, END_DATE, _
                  "d", "DOHLCVA", False, True, True)
Else
    DATA_MATRIX = TICKER_STR
End If

NROWS = UBound(DATA_MATRIX, 1)
NCOLUMNS = UBound(DATA_MATRIX, 2)

FACTOR_VAL = NO_PERIODS * (NO_PERIODS + 1) / 2 'f1
'----------------------------------------------------------------
Select Case VERSION
'----------------------------------------------------------------
Case 0 'f1, f2, f3, f4
'----------------------------------------------------------------
    ReDim TEMP_VECTOR(1 To 3, 1 To 1)
    ReDim TEMP_MATRIX(1 To 3, 1 To 3)
    
    TEMP_MATRIX(1, 1) = NO_PERIODS * (NO_PERIODS + 1) * (2 * NO_PERIODS + 1) * _
                      (3 * NO_PERIODS * NO_PERIODS + 3 * NO_PERIODS - 1) / 30
    TEMP_MATRIX(2, 1) = FACTOR_VAL ^ 2
    TEMP_MATRIX(3, 1) = NO_PERIODS * (NO_PERIODS + 1) * (2 * NO_PERIODS + 1) / 6
    
    TEMP_MATRIX(1, 2) = TEMP_MATRIX(2, 1)
    TEMP_MATRIX(2, 2) = TEMP_MATRIX(3, 1)
    TEMP_MATRIX(3, 2) = FACTOR_VAL
    
    TEMP_MATRIX(1, 3) = TEMP_MATRIX(3, 1)
    TEMP_MATRIX(2, 3) = FACTOR_VAL
    TEMP_MATRIX(3, 3) = NO_PERIODS

    TEMP_VECTOR(1, 1) = 0: TEMP_VECTOR(2, 1) = 0: TEMP_VECTOR(3, 1) = 0
    For i = 1 To NO_PERIODS - 6
        TEMP_VECTOR(1, 1) = TEMP_VECTOR(1, 1) + DATA_MATRIX(i, 7) * (i - 1) ^ 2
        TEMP_VECTOR(2, 1) = TEMP_VECTOR(2, 1) + DATA_MATRIX(i, 7) * (i - 1)
        TEMP_VECTOR(3, 1) = TEMP_VECTOR(3, 1) + DATA_MATRIX(i, 7)
    Next i
    TEMP_VECTOR = MMULT_FUNC(MATRIX_INVERSE_FUNC(TEMP_MATRIX, 0), TEMP_VECTOR, 70)
    If OUTPUT <> 0 Then
        ASSET_TA_PARABOLAS_FUNC = TEMP_VECTOR
        Exit Function
    End If
    ReDim TEMP_MATRIX(0 To NROWS, 1 To NCOLUMNS + 2)
    For i = 1 To NROWS
        For j = 1 To 7: TEMP_MATRIX(i, j) = DATA_MATRIX(i, j): Next j
        TEMP_MATRIX(i, 6) = TEMP_MATRIX(i, 6) / 1000
        
        TEMP_MATRIX(i, NCOLUMNS + 1) = TEMP_VECTOR(1, 1) * (i - 1) ^ 2 + _
                                       TEMP_VECTOR(2, 1) * (i - 1) ^ 1 + _
                                       TEMP_VECTOR(3, 1)
        TEMP_MATRIX(i, NCOLUMNS + 2) = TEMP_MATRIX(i, NCOLUMNS + 1) + SHIFT_FACTOR
    Next i
    TEMP_STR = "gSAR = (" & Format(TEMP_VECTOR(1, 1), "0.0000") & ")n^2 + (" & _
                            Format(TEMP_VECTOR(2, 1), "0.000") & ")n + (" & _
                            Format(TEMP_MATRIX(1, NCOLUMNS + 1), "0.00") & ")"

'----------------------------------------------------------------
Case Else 'f1, f2, f3, f4, f5, f6 --> CUBIC
'----------------------------------------------------------------
    ReDim TEMP_VECTOR(1 To 4, 1 To 1)
    ReDim TEMP_MATRIX(1 To 4, 1 To 4)
    TEMP_MATRIX(1, 1) = (6 * NO_PERIODS ^ 7 + 21 * NO_PERIODS ^ 6 + _
                        21 * NO_PERIODS ^ 5 - 7 * NO_PERIODS ^ 3 + NO_PERIODS) / 42
    TEMP_MATRIX(2, 1) = NO_PERIODS ^ 2 * (NO_PERIODS + 1) ^ 2 * (2 * NO_PERIODS ^ 2 + _
                        2 * NO_PERIODS - 1) / 12
    TEMP_MATRIX(3, 1) = NO_PERIODS * (NO_PERIODS + 1) * (2 * NO_PERIODS + 1) * _
                        (3 * NO_PERIODS * NO_PERIODS + 3 * NO_PERIODS - 1) / 30
    TEMP_MATRIX(4, 1) = FACTOR_VAL ^ 2
    
    TEMP_MATRIX(1, 2) = TEMP_MATRIX(2, 1)
    TEMP_MATRIX(2, 2) = TEMP_MATRIX(3, 1)
    TEMP_MATRIX(3, 2) = TEMP_MATRIX(4, 1)
    TEMP_MATRIX(4, 2) = NO_PERIODS * (NO_PERIODS + 1) * (2 * NO_PERIODS + 1) / 6
    
    TEMP_MATRIX(1, 3) = TEMP_MATRIX(3, 1)
    TEMP_MATRIX(2, 3) = TEMP_MATRIX(4, 1)
    TEMP_MATRIX(3, 3) = TEMP_MATRIX(4, 2)
    TEMP_MATRIX(4, 3) = FACTOR_VAL
    
    TEMP_MATRIX(1, 4) = TEMP_MATRIX(4, 1)
    TEMP_MATRIX(2, 4) = TEMP_MATRIX(4, 2)
    TEMP_MATRIX(3, 4) = TEMP_MATRIX(4, 3)
    TEMP_MATRIX(4, 4) = NO_PERIODS
    
    TEMP_VECTOR(1, 1) = 0: TEMP_VECTOR(2, 1) = 0
    TEMP_VECTOR(3, 1) = 0: TEMP_VECTOR(4, 1) = 0
    For i = 1 To NO_PERIODS - 6
        TEMP_VECTOR(1, 1) = TEMP_VECTOR(1, 1) + DATA_MATRIX(i, 7) * (i - 1) ^ 3
        TEMP_VECTOR(2, 1) = TEMP_VECTOR(2, 1) + DATA_MATRIX(i, 7) * (i - 1) ^ 2
        TEMP_VECTOR(3, 1) = TEMP_VECTOR(3, 1) + DATA_MATRIX(i, 7) * (i - 1)
        TEMP_VECTOR(4, 1) = TEMP_VECTOR(4, 1) + DATA_MATRIX(i, 7)
    Next i
    TEMP_VECTOR = MMULT_FUNC(MATRIX_INVERSE_FUNC(TEMP_MATRIX, 0), TEMP_VECTOR, 70)
    If OUTPUT <> 0 Then
        ASSET_TA_PARABOLAS_FUNC = TEMP_VECTOR
        Exit Function
    End If
    
    ReDim TEMP_MATRIX(0 To NROWS, 1 To NCOLUMNS + 2)
    For i = 1 To NROWS
        For j = 1 To 7: TEMP_MATRIX(i, j) = DATA_MATRIX(i, j): Next j
        TEMP_MATRIX(i, 6) = TEMP_MATRIX(i, 6) / 1000

        TEMP_MATRIX(i, NCOLUMNS + 1) = TEMP_VECTOR(1, 1) * (i - 1) ^ 3 + _
                                       TEMP_VECTOR(2, 1) * (i - 1) ^ 2 + _
                                       TEMP_VECTOR(3, 1) * (i - 1) ^ 1 + _
                                       TEMP_VECTOR(4, 1)
        TEMP_MATRIX(i, NCOLUMNS + 2) = TEMP_MATRIX(i, NCOLUMNS + 1) + SHIFT_FACTOR
    Next i
    TEMP_STR = "g3SAR = (" & Format(TEMP_VECTOR(1, 1), "0.0000000") & ")n^3 + (" & _
                             Format(TEMP_VECTOR(2, 1), "0.0000") & ")n^2 + (" & _
                             Format(TEMP_MATRIX(1, NCOLUMNS + 1), "0.00") & ") "

'----------------------------------------------------------------
End Select
'----------------------------------------------------------------

TEMP_MATRIX(0, 1) = "DATE"
TEMP_MATRIX(0, 2) = "OPEN"
TEMP_MATRIX(0, 3) = "HIGH"
TEMP_MATRIX(0, 4) = "LOW"
TEMP_MATRIX(0, 5) = "CLOSE"
TEMP_MATRIX(0, 6) = "VOLUME"
TEMP_MATRIX(0, 7) = "ADJ CLOSE"
TEMP_MATRIX(0, 8) = "PARABOLA"
TEMP_MATRIX(0, 9) = TEMP_STR

ASSET_TA_PARABOLAS_FUNC = TEMP_MATRIX

Exit Function
ERROR_LABEL:
ASSET_TA_PARABOLAS_FUNC = Err.number
End Function
