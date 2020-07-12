Attribute VB_Name = "FINAN_FI_BDT_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.

'**********************************************************************************
'**********************************************************************************
'FUNCTION      : BDT_SWAPTION_FUNC
'DESCRIPTION   : SWAPTION GRID MODEL
'LIBRARY       : FIXED_INCOME
'GROUP         : SWAP_BDT
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 10-06-2008
'**********************************************************************************
'**********************************************************************************

Function BDT_SWAPTION_FUNC(ByVal PERIODS As Long, _
ByVal FIXED_RATE As Double, _
ByVal EXPIRATION As Double, _
ByVal STRIKE As Double, _
ByRef SPOT_RNG As Variant, _
ByRef FAIR_RNG As Variant, _
ByRef SIGMA_RNG As Variant, _
Optional ByVal PROB_UP As Double = 0.5, _
Optional ByVal RATE_FACTOR As Double = 100, _
Optional ByVal OUTPUT As Integer = 0)

Dim i As Long
Dim j As Long
Dim k As Long

Dim FAIR_ARR As Variant
Dim TEMP_MATRIX As Variant

On Error GoTo ERROR_LABEL

'--------------------REMEMBER YOU MUST CALIBRATE THE YIELD CURVE FIRST----------------
FAIR_ARR = Excel.Application.Run("BDT_SWAPTION_GRID_FUNC", 2, SPOT_RNG, FAIR_RNG, SIGMA_RNG, _
PROB_UP, RATE_FACTOR)
'-------------------------------------------------------------------------------------

If PERIODS > UBound(FAIR_ARR, 2) - 1 Then: GoTo ERROR_LABEL

ReDim TEMP_MATRIX(1 To PERIODS + 1, 1 To PERIODS + 1)

'-------------------------FIRST PASS: SETTING UP THE GRID------------------
TEMP_MATRIX(1, 1) = ("SWAPTION LATTICE")
For i = 1 To PERIODS 'Note that the values at a node are the discounted values
'of the nodes 1 period ahead. We therefore start from t=(PERIODS -1) even though
'final payoff occurs at t=PERIODS
    TEMP_MATRIX(1, i + 1) = i - 1
    TEMP_MATRIX(i + 1, 1) = PERIODS - i
Next i

k = (UBound(FAIR_ARR, 1) - UBound(TEMP_MATRIX, 1)) ' For locating cells

'-------------------------SECOND PASS: FILLING THE GRID-------------------

For i = 2 To PERIODS + 1
    For j = PERIODS + 1 To 2 Step -1
    
        If TEMP_MATRIX(1, j) < EXPIRATION Then
            If TEMP_MATRIX(i, 1) <= TEMP_MATRIX(1, j) Then
                TEMP_MATRIX(i, j) = (PROB_UP * TEMP_MATRIX(i - 1, j + 1) + _
                (1 - PROB_UP) * TEMP_MATRIX(i, j + 1)) / (1 + FAIR_ARR(i + k, j) _
                / RATE_FACTOR)
            Else
                TEMP_MATRIX(i, j) = ""
            End If

        ElseIf TEMP_MATRIX(1, j) = EXPIRATION Then
            If TEMP_MATRIX(i, 1) <= TEMP_MATRIX(1, j) Then
               TEMP_MATRIX(i, j) = MAXIMUM_FUNC((FAIR_ARR(i + k, j) / RATE_FACTOR _
               - FIXED_RATE) / (1 + FAIR_ARR(i + k, j) / RATE_FACTOR) _
                 + (PROB_UP * TEMP_MATRIX(i - 1, j + 1) + (1 - PROB_UP) * _
                 TEMP_MATRIX(i, j + 1)) / (1 + FAIR_ARR(i + k, j) / RATE_FACTOR) _
                 - STRIKE, 0)
            Else
                TEMP_MATRIX(i, j) = ""
            End If
        Else
            If TEMP_MATRIX(1, j) = (PERIODS - 1) Then
                If TEMP_MATRIX(i, 1) <= TEMP_MATRIX(1, j) Then
                    TEMP_MATRIX(i, j) = (FAIR_ARR(i + k, j) / RATE_FACTOR _
                    - FIXED_RATE) / (1 + FAIR_ARR(i + k, j) / RATE_FACTOR)
                Else
                    TEMP_MATRIX(i, j) = ""
                End If
             Else
                If TEMP_MATRIX(i, 1) <= TEMP_MATRIX(1, j) Then
                    TEMP_MATRIX(i, j) = (FAIR_ARR(i + k, j) / RATE_FACTOR - FIXED_RATE) _
                    / (1 + FAIR_ARR(i + k, j) / RATE_FACTOR) + _
                    (PROB_UP * TEMP_MATRIX(i - 1, j + 1) + (1 - PROB_UP) * _
                    TEMP_MATRIX(i, j + 1)) / (1 + FAIR_ARR(i + k, j) / RATE_FACTOR)
                Else
                    TEMP_MATRIX(i, j) = ""
                End If
            End If
        End If
    Next j
Next i

'---------------------------------------------------------------------------------------
For j = 1 To UBound(TEMP_MATRIX, 2)
    For i = 1 To UBound(TEMP_MATRIX, 1)
        If IsEmpty(TEMP_MATRIX(i, j)) = True Then: TEMP_MATRIX(i, j) = ""
    Next i
Next j
'---------------------------------------------------------------------------------------

Select Case OUTPUT
    Case 0
        BDT_SWAPTION_FUNC = TEMP_MATRIX
    Case Else
        BDT_SWAPTION_FUNC = Excel.Application.Run("BDT_SWAPTION_GRID_FUNC", 0, _
        SPOT_RNG, FAIR_RNG, SIGMA_RNG, _
        PROB_UP, RATE_FACTOR)
End Select

Exit Function
ERROR_LABEL:
BDT_SWAPTION_FUNC = Err.number
End Function


'**********************************************************************************
'**********************************************************************************
'FUNCTION      : BDT_SWAPTION_GRID_FUNC
'DESCRIPTION   : BLACK DERMAN TOY SWAPTION GRID MODEL
'LIBRARY       : FIXED_INCOME
'GROUP         : SWAP_BDT
'ID            : 002
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 10-06-2008
'**********************************************************************************
'**********************************************************************************

Private Function BDT_SWAPTION_GRID_FUNC(ByVal OUTPUT As Integer, _
ByRef SPOT_RNG As Variant, _
ByRef FAIR_RNG As Variant, _
ByRef SIGMA_RNG As Variant, _
Optional ByVal DELTA_STEP As Double = 1, _
Optional ByVal PROB_UP As Double = 0.5, _
Optional ByVal RATE_FACTOR As Double = 100)

Dim i As Long
Dim j As Long

Dim SPOT_VECTOR As Variant
Dim FAIR_VECTOR As Variant
Dim SIGMA_VECTOR As Variant

Dim SHORT_RATE_ARR As Variant
Dim ASSET_PRICE_ARR As Variant
Dim ZERO_PRICE_ARR As Variant
Dim TEMP_MATRIX As Variant

Dim TEMP_SUM As Double

Dim PERIODS As Variant

On Error GoTo ERROR_LABEL

SPOT_VECTOR = SPOT_RNG
If UBound(SPOT_VECTOR, 2) = 1 Then
    SPOT_VECTOR = MATRIX_TRANSPOSE_FUNC(SPOT_VECTOR)
End If

FAIR_VECTOR = FAIR_RNG
If UBound(FAIR_VECTOR, 2) = 1 Then
    FAIR_VECTOR = MATRIX_TRANSPOSE_FUNC(FAIR_VECTOR)
End If

SIGMA_VECTOR = SIGMA_RNG
If UBound(SIGMA_VECTOR, 2) = 1 Then
    SIGMA_VECTOR = MATRIX_TRANSPOSE_FUNC(SIGMA_VECTOR)
End If

PERIODS = UBound(SPOT_VECTOR, 2) - 1

ReDim SHORT_RATE_ARR(1 To PERIODS + 2, 1 To PERIODS + 2)
ReDim ASSET_PRICE_ARR(1 To PERIODS + 3, 1 To PERIODS + 3)

'-------------------------FIRST PASS: SETTING UP THE GRIDS---------
SHORT_RATE_ARR(1, 1) = ("SHORT RATE LATTICE")
For i = 1 To PERIODS + 1
    SHORT_RATE_ARR(1, i + 1) = i - 1
    SHORT_RATE_ARR(i + 1, 1) = PERIODS - i + 1
Next i

For i = 1 To PERIODS + 2
    ASSET_PRICE_ARR(1, i + 1) = i - 1
    ASSET_PRICE_ARR(i + 1, 1) = PERIODS - i + 2
Next i

'-------------------------SECOND PASS: FILLING THE GRIDS-----------

For i = 2 To PERIODS + 2
    For j = 2 To PERIODS + 2
        If SHORT_RATE_ARR(i, 1) <= SHORT_RATE_ARR(1, j) Then
            If IsArray(SIGMA_VECTOR) = True Then
                SHORT_RATE_ARR(i, j) = FAIR_VECTOR(1, j - 1) * _
                (Exp(SIGMA_VECTOR(1, j - 1) * SHORT_RATE_ARR(i, 1) * _
                Sqr(DELTA_STEP)))
            Else
                SHORT_RATE_ARR(i, j) = FAIR_VECTOR(1, j - 1) * _
                (Exp(SIGMA_VECTOR * SHORT_RATE_ARR(i, 1)))
            End If
        End If
    Next j
Next i

ASSET_PRICE_ARR(PERIODS + 3, 2) = 1

For i = (PERIODS + 3) To 2 Step -1 'Backward Looping
    For j = 3 To PERIODS + 3
          If ASSET_PRICE_ARR(i, 1) = 0 Then
              ASSET_PRICE_ARR(i, j) = (1 - PROB_UP) * ASSET_PRICE_ARR(i, j - 1) _
               / (1 + SHORT_RATE_ARR(i - 1, j - 1) / RATE_FACTOR)
          ElseIf ASSET_PRICE_ARR(i, 1) = ASSET_PRICE_ARR(1, j) Then
                ASSET_PRICE_ARR(i, j) = PROB_UP * ASSET_PRICE_ARR(i + 1, j - 1) _
                / (1 + SHORT_RATE_ARR(i, j - 1) / RATE_FACTOR)
          ElseIf (0 < ASSET_PRICE_ARR(i, 1)) And _
          (ASSET_PRICE_ARR(i, 1) < ASSET_PRICE_ARR(1, j)) Then
                ASSET_PRICE_ARR(i, j) = PROB_UP * ASSET_PRICE_ARR(i + 1, j - 1) _
                / (1 + SHORT_RATE_ARR(i, j - 1) / RATE_FACTOR) + (1 - PROB_UP) _
                * ASSET_PRICE_ARR(i, j - 1) / (1 + SHORT_RATE_ARR(i - 1, j - 1) / _
                RATE_FACTOR)
          Else
                ASSET_PRICE_ARR(i, j) = 0
          End If
    Next j
Next i

'-------------------We need first to take out the PERIODS HEADINGS----------------------------
ZERO_PRICE_ARR = MATRIX_SUM_FUNC(MATRIX_REMOVE_COLUMNS_FUNC( _
                 MATRIX_REMOVE_ROWS_FUNC(ASSET_PRICE_ARR, 1, 1), 1, 1))
'---------------------------------------------------------------------------------------------

'-------------------------------------------------------------------------------------------
Select Case OUTPUT
'-------------------------------------------------------------------------------------------
Case 0
'-------------------------------------------------------------------------------------------
    ReDim TEMP_MATRIX(0 To 5, 1 To UBound(ZERO_PRICE_ARR, 2) + 1)
    
    TEMP_MATRIX(0, 1) = "PERIOD (STEP)"
    
    TEMP_MATRIX(1, 1) = "SPOT_RATES"
    TEMP_MATRIX(2, 1) = "ZERO_PRICES"
    TEMP_MATRIX(3, 1) = "ESTIMATED_SPOT_RATES"
    TEMP_MATRIX(4, 1) = "SQUARED_DIFFERENCES"
    TEMP_MATRIX(5, 1) = "OBJECTIVE_FUNCTION"
    
    TEMP_MATRIX(1, 2) = SPOT_VECTOR(1, 1)
    TEMP_MATRIX(2, 2) = ASSET_PRICE_ARR(PERIODS + 3, 2)
    
    TEMP_MATRIX(3, 2) = 0
    TEMP_MATRIX(4, 2) = 0
    
    TEMP_SUM = 0
    For i = 2 To UBound(ZERO_PRICE_ARR, 2)
        TEMP_MATRIX(0, i + 1) = i - 1
        
        If i <> UBound(ZERO_PRICE_ARR, 2) Then
            TEMP_MATRIX(1, i + 1) = SPOT_VECTOR(1, i)
        Else
            TEMP_MATRIX(1, i + 1) = 0
        End If
        
        TEMP_MATRIX(2, i + 1) = ZERO_PRICE_ARR(1, i)
        TEMP_MATRIX(3, i + 1) = ((1 / ZERO_PRICE_ARR(1, i)) ^ (1 / (i - 1)) - 1) * 100
        TEMP_MATRIX(4, i + 1) = (TEMP_MATRIX(1, i) - TEMP_MATRIX(3, i + 1)) ^ 2
        TEMP_SUM = TEMP_SUM + TEMP_MATRIX(4, i + 1)
        TEMP_MATRIX(5, i + 1) = 0
    Next i
    
    TEMP_MATRIX(5, 2) = TEMP_SUM '---> OBJECTIVE FUNCTION
    'Now use solver to match the term structure of zero prices
    'by setting the objective function to 0: CHANGING CELLS: ACTUAL SPOT RANGE
    
    BDT_SWAPTION_GRID_FUNC = TEMP_MATRIX
'-------------------------------------------------------------------------------------------
Case 1
'-------------------------------------------------------------------------------------------
    For j = 1 To UBound(ZERO_PRICE_ARR, 2)
        For i = 1 To UBound(ZERO_PRICE_ARR, 1)
            If IsEmpty(ZERO_PRICE_ARR(i, j)) = True Then: ZERO_PRICE_ARR(i, j) = ""
        Next i
    Next j
    BDT_SWAPTION_GRID_FUNC = ZERO_PRICE_ARR
'-------------------------------------------------------------------------------------------
Case 2
'-------------------------------------------------------------------------------------------
    For j = 1 To UBound(SHORT_RATE_ARR, 2)
        For i = 1 To UBound(SHORT_RATE_ARR, 1)
            If IsEmpty(SHORT_RATE_ARR(i, j)) = True Then: SHORT_RATE_ARR(i, j) = ""
        Next i
    Next j
    BDT_SWAPTION_GRID_FUNC = SHORT_RATE_ARR
'-------------------------------------------------------------------------------------------
Case 3
'-------------------------------------------------------------------------------------------
    For j = 1 To UBound(ASSET_PRICE_ARR, 2)
        For i = 1 To UBound(ASSET_PRICE_ARR, 1)
            If IsEmpty(ASSET_PRICE_ARR(i, j)) = True Then: ASSET_PRICE_ARR(i, j) = ""
        Next i
    Next j
    BDT_SWAPTION_GRID_FUNC = ASSET_PRICE_ARR
'-------------------------------------------------------------------------------------------
End Select
'-------------------------------------------------------------------------------------------

Exit Function
ERROR_LABEL:
BDT_SWAPTION_GRID_FUNC = Err.number
End Function

