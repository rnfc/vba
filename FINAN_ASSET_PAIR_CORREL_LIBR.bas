Attribute VB_Name = "FINAN_ASSET_PAIR_CORREL_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.

Function ASSET_PERFECT_CORRELATION_FUNC( _
ByVal TICKER_STR As Variant, _
Optional ByVal START_DATE As Date, _
Optional ByVal END_DATE As Date, _
Optional ByVal INITIAL_INVESTMENT As Double = 1000, _
Optional ByVal INITIAL_ALLOCATION As Double = 0.6, _
Optional ByVal P_VAL As Double = 1, _
Optional ByVal A_VAL As Double = -0.1, _
Optional ByVal B_VAL As Double = 0.2, _
Optional ByVal N_VAL As Long = 1, _
Optional ByVal K_VAL As Double = 0, _
Optional ByVal VERSION As Integer = 0)

'INITIAL ALLOCATION IN TICKER_STR
'P_VAL --> CLOSING PRICE OF THE A2 STOCK, THE RETURN OF THIS ASSET
'IS DERIVED FROM THE TICKER_STR

Dim i As Long
Dim NROWS As Long
Dim MEAN_VAL As Double
Dim DIFF_VAL As Double

Dim TEMP1_SUM As Double
Dim TEMP2_SUM As Double

Dim RHO_VAL As Double
Dim ALLOC1_VAL As Double
Dim ALLOC2_VAL As Double

Dim DATA1_VECTOR As Variant
Dim DATA2_VECTOR As Variant

Dim TEMP_MATRIX As Variant
Dim DATA_MATRIX As Variant

On Error GoTo ERROR_LABEL

If IsArray(TICKER_STR) = False Then
    DATA_MATRIX = _
    YAHOO_HISTORICAL_DATA_SERIE_FUNC(TICKER_STR, _
    START_DATE, END_DATE, "DAILY", "DA", True, False, True)
Else
    DATA_MATRIX = TICKER_STR
    TICKER_STR = "A1"
End If
NROWS = UBound(DATA_MATRIX, 1)

ReDim DATA1_VECTOR(1 To NROWS - 1, 1 To 1)
ReDim DATA2_VECTOR(1 To NROWS - 1, 1 To 1)
ReDim TEMP_MATRIX(0 To NROWS, 1 To 8)
'PERIOD  Stock X Stock Y Portfolio   Stock X Stock Y (x-M)^2 (x-M)^3
TEMP_MATRIX(0, 1) = "DATE"

i = 1
TEMP_MATRIX(i, 1) = DATA_MATRIX(i, 1)
TEMP_MATRIX(i, 2) = DATA_MATRIX(i, 2)
TEMP_MATRIX(i, 3) = ""
TEMP_MATRIX(i, 4) = P_VAL
TEMP_MATRIX(i, 5) = ""
TEMP_MATRIX(i, 6) = INITIAL_INVESTMENT
TEMP_MATRIX(i, 7) = ""
TEMP_MATRIX(i, 8) = ""

'------------------------------------------------------------------------------------------------------
For i = 2 To NROWS
'------------------------------------------------------------------------------------------------------
    TEMP_MATRIX(i, 1) = DATA_MATRIX(i, 1)
    TEMP_MATRIX(i, 2) = DATA_MATRIX(i, 2)
    TEMP_MATRIX(i, 3) = DATA_MATRIX(i, 2) / DATA_MATRIX(i - 1, 2) - 1
    MEAN_VAL = MEAN_VAL + TEMP_MATRIX(i, 3)
    
    
    If VERSION = 0 Then
    'To get Correlation = 0
    'set
    '1) y = a*(x-K)^2 + b*(x-K)
    '2) b to -a*S_2 / S_1
    '3) n to 2
    '4) K to M
        DIFF_VAL = (TEMP_MATRIX(i, 3) - K_VAL)
        TEMP_MATRIX(i, 5) = A_VAL * DIFF_VAL ^ N_VAL + B_VAL * IIf(K_VAL <> 0, DIFF_VAL, 1)
    Else
    'To get Correlation = +1 or -1
    'set
    '1) y = a*x^n + b
    '2) a and b to any number
    '3) n to 1
    '4) K to 0
        TEMP_MATRIX(i, 5) = A_VAL * TEMP_MATRIX(i, 3) ^ N_VAL + B_VAL
    End If
    
    TEMP_MATRIX(i, 4) = TEMP_MATRIX(i - 1, 4) * (1 + TEMP_MATRIX(i, 5))
    DATA1_VECTOR(i - 1, 1) = TEMP_MATRIX(i, 3)
    DATA2_VECTOR(i - 1, 1) = TEMP_MATRIX(i, 5)
    
    ALLOC1_VAL = INITIAL_ALLOCATION * TEMP_MATRIX(i, 3)
    ALLOC2_VAL = (1 - INITIAL_ALLOCATION) * TEMP_MATRIX(i, 5)
    
    TEMP_MATRIX(i, 6) = TEMP_MATRIX(i - 1, 6) * (1 + ALLOC1_VAL + ALLOC2_VAL)
'------------------------------------------------------------------------------------------------------
Next i
'------------------------------------------------------------------------------------------------------

RHO_VAL = CORRELATION_FUNC(DATA1_VECTOR, DATA2_VECTOR, 0, 0)
MEAN_VAL = MEAN_VAL / (NROWS - 1)

TEMP1_SUM = 0: TEMP2_SUM = 0
For i = 2 To NROWS
    DIFF_VAL = (TEMP_MATRIX(i, 3) - MEAN_VAL)
    TEMP_MATRIX(i, 7) = DIFF_VAL ^ 2
    TEMP1_SUM = TEMP1_SUM + TEMP_MATRIX(i, 7)
    
    TEMP_MATRIX(i, 8) = DIFF_VAL ^ 3
    TEMP2_SUM = TEMP2_SUM + TEMP_MATRIX(i, 8)
Next i

TEMP_MATRIX(0, 2) = "ASSET " & TICKER_STR
TEMP_MATRIX(0, 3) = "RETURN " & TICKER_STR & "- RHO: " & Format(RHO_VAL, "0.00%")
TEMP_MATRIX(0, 4) = "ASSET2"
TEMP_MATRIX(0, 5) = "RETURN2" & "- RHO: " & Format(RHO_VAL, "0.00%")
TEMP_MATRIX(0, 6) = "PORTFOLIO " & " - WEIGHT: " & Format(INITIAL_ALLOCATION, "0.00%") & " / " & Format(1 - INITIAL_ALLOCATION, "0.00%")
TEMP_MATRIX(0, 7) = "( " & TICKER_STR & " - MEAN_" & TICKER_STR & " )^ 2: SUM = " & Format(TEMP1_SUM, "0.00%")
TEMP_MATRIX(0, 8) = "( " & TICKER_STR & " - MEAN_" & TICKER_STR & " )^ 3: SUM = " & Format(TEMP2_SUM, "0.00%")

ASSET_PERFECT_CORRELATION_FUNC = TEMP_MATRIX

Exit Function
ERROR_LABEL:
ASSET_PERFECT_CORRELATION_FUNC = Err.number
End Function
