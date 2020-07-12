Attribute VB_Name = "FINAN_COMMOD_GOLD_MINE_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.

Function GOLD_MINE_PROFIT_FUNC( _
ByRef ORE_TONNES_RNG As Variant, _
ByRef GOLD_GRADE_RNG As Variant, _
ByRef GOLD_PRICE_RNG As Variant, _
ByRef EXCHANGE_RATE_RNG As Variant, _
ByRef MINE_UNIT_COST_RNG As Variant, _
ByRef PROCESS_UNIT_COST_RNG As Variant, _
Optional ByVal OUNCE_GRAMS_VAL As Double = 31.1035)

'----------------------------------------------------------------------------------------------------------
'----------------------------------------------------------------------------------------------------------
'Ore Tonnes (t): 1,000,000
'Gold Grade (g/t Au): 1.68
'Gold price (U$/Oz Au): 1,200
'Exchange Rate (A$/U$): 0.92
'Mine Unit Cost ($/t): 40
'Process Unit Cost ($/t): 25
'Total Cost ($): 65,000,000
'Revenue ($): 70,452,018
'Profit ($): 5,452,018
'one ounce/31.1035 grams
'----------------------------------------------------------------------------------------------------------
'For Simulation Purpose:
'Gold Price
'EXCHANGE RATE
'----------------------------------------------------------------------------------------------------------
'Scenario No.    Ore Tonnes  Gold Grade  Gold Price  Ex Rate Mine Cost   Process Cost
'----------------------------------------------------------------------------------------------------------
'1   Base Case      1,000,000   1.68   1200     0.92   40  25
'2   High Tonnes    1,200,000   1.68   1200     0.92   40  25
'3   Low Tonnes       800,000   1.68   1200     0.92   40  25
'4   High Grade     1,000,000   1.80   1200     0.92   40  25
'5   Low Grade      1,000,000   1.68   1200     0.92   40  25
'6   High Ex Rate   1,000,000   1.50   1200     1.00   40  25
'7   Low Ex Rate    1,000,000   1.68   1200     0.85   40  25
'8   High Cost      1,000,000   1.68   1200     0.92   50  35
'9   Low Cost       1,000,000   1.68   1200     0.92   30  20
'10  Hi Gold Price  1,000,000   1.68   1400     0.92   40  25
'11  Low Gold Price 1,000,000   1.68   1000     0.92   40  25
'12  Worst Case       800,000   1.58   1000     1.00   50  35
'----------------------------------------------------------------------------------------------------------
'----------------------------------------------------------------------------------------------------------

Dim i As Long
Dim NROWS As Long

Dim ORE_TONNES_VECTOR As Variant
Dim GOLD_GRADE_VECTOR As Variant
Dim GOLD_PRICE_VECTOR As Variant
Dim EXCHANGE_RATE_VECTOR As Variant
Dim MINE_UNIT_COST_VECTOR As Variant
Dim PROCESS_UNIT_COST_VECTOR As Variant

Dim TEMP_MATRIX As Variant
On Error GoTo ERROR_LABEL

If IsArray(ORE_TONNES_RNG) = True Then
    ORE_TONNES_VECTOR = ORE_TONNES_RNG
    If UBound(ORE_TONNES_VECTOR, 1) = 1 Then
        ORE_TONNES_VECTOR = MATRIX_TRANSPOSE_FUNC(ORE_TONNES_VECTOR)
    End If
Else
    ReDim ORE_TONNES_VECTOR(1 To 1, 1 To 1)
    ORE_TONNES_VECTOR(1, 1) = ORE_TONNES_RNG
End If
NROWS = UBound(ORE_TONNES_VECTOR, 1)

If IsArray(GOLD_GRADE_RNG) = True Then
    GOLD_GRADE_VECTOR = GOLD_GRADE_RNG
    If UBound(GOLD_GRADE_VECTOR, 1) = 1 Then
        GOLD_GRADE_VECTOR = MATRIX_TRANSPOSE_FUNC(GOLD_GRADE_VECTOR)
    End If
Else
    ReDim GOLD_GRADE_VECTOR(1 To 1, 1 To 1)
    GOLD_GRADE_VECTOR(1, 1) = GOLD_GRADE_RNG
End If
If NROWS <> UBound(GOLD_GRADE_VECTOR, 1) Then: GoTo ERROR_LABEL

If IsArray(GOLD_PRICE_RNG) = True Then
    GOLD_PRICE_VECTOR = GOLD_PRICE_RNG
    If UBound(GOLD_PRICE_VECTOR, 1) = 1 Then
        GOLD_PRICE_VECTOR = MATRIX_TRANSPOSE_FUNC(GOLD_PRICE_VECTOR)
    End If
Else
    ReDim GOLD_PRICE_VECTOR(1 To 1, 1 To 1)
    GOLD_PRICE_VECTOR(1, 1) = GOLD_PRICE_RNG
End If
If NROWS <> UBound(GOLD_PRICE_VECTOR, 1) Then: GoTo ERROR_LABEL

If IsArray(EXCHANGE_RATE_RNG) = True Then
    EXCHANGE_RATE_VECTOR = EXCHANGE_RATE_RNG
    If UBound(EXCHANGE_RATE_VECTOR, 1) = 1 Then
        EXCHANGE_RATE_VECTOR = MATRIX_TRANSPOSE_FUNC(EXCHANGE_RATE_VECTOR)
    End If
Else
    ReDim EXCHANGE_RATE_VECTOR(1 To NROWS, 1 To 1)
    For i = 1 To NROWS
        EXCHANGE_RATE_VECTOR(i, 1) = EXCHANGE_RATE_RNG
    Next i
End If
If NROWS <> UBound(EXCHANGE_RATE_VECTOR, 1) Then: GoTo ERROR_LABEL

If IsArray(MINE_UNIT_COST_RNG) = True Then
    MINE_UNIT_COST_VECTOR = MINE_UNIT_COST_RNG
    If UBound(MINE_UNIT_COST_VECTOR, 1) = 1 Then
        MINE_UNIT_COST_VECTOR = MATRIX_TRANSPOSE_FUNC(MINE_UNIT_COST_VECTOR)
    End If
Else
    ReDim MINE_UNIT_COST_VECTOR(1 To 1, 1 To 1)
    MINE_UNIT_COST_VECTOR(1, 1) = MINE_UNIT_COST_RNG
End If
If NROWS <> UBound(MINE_UNIT_COST_VECTOR, 1) Then: GoTo ERROR_LABEL

If IsArray(PROCESS_UNIT_COST_RNG) = True Then
    PROCESS_UNIT_COST_VECTOR = PROCESS_UNIT_COST_RNG
    If UBound(PROCESS_UNIT_COST_VECTOR, 1) = 1 Then
        PROCESS_UNIT_COST_VECTOR = MATRIX_TRANSPOSE_FUNC(PROCESS_UNIT_COST_VECTOR)
    End If
Else
    ReDim PROCESS_UNIT_COST_VECTOR(1 To 1, 1 To 1)
    PROCESS_UNIT_COST_VECTOR(1, 1) = PROCESS_UNIT_COST_RNG
End If
If NROWS <> UBound(PROCESS_UNIT_COST_VECTOR, 1) Then: GoTo ERROR_LABEL

ReDim TEMP_MATRIX(0 To NROWS, 1 To 10)
TEMP_MATRIX(0, 1) = "ORE TONNES"
TEMP_MATRIX(0, 2) = "GOLD GRADE"
TEMP_MATRIX(0, 3) = "GOLD PRICE"
TEMP_MATRIX(0, 4) = "EXCHANGE RATE"
TEMP_MATRIX(0, 5) = "MINE UNIT COST"
TEMP_MATRIX(0, 6) = "PROCESS UNIT COST"
TEMP_MATRIX(0, 7) = "TOTAL COST"
TEMP_MATRIX(0, 8) = "COST/Oz"
TEMP_MATRIX(0, 9) = "REVENUE"
TEMP_MATRIX(0, 10) = "PROFIT"

For i = 1 To NROWS
    TEMP_MATRIX(i, 1) = ORE_TONNES_VECTOR(i, 1)
    TEMP_MATRIX(i, 2) = GOLD_GRADE_VECTOR(i, 1)
    TEMP_MATRIX(i, 3) = GOLD_PRICE_VECTOR(i, 1)
    TEMP_MATRIX(i, 4) = EXCHANGE_RATE_VECTOR(i, 1)
    TEMP_MATRIX(i, 5) = MINE_UNIT_COST_VECTOR(i, 1)
    TEMP_MATRIX(i, 6) = PROCESS_UNIT_COST_VECTOR(i, 1)
    TEMP_MATRIX(i, 7) = TEMP_MATRIX(i, 1) * (TEMP_MATRIX(i, 5) + TEMP_MATRIX(i, 6))
    TEMP_MATRIX(i, 8) = TEMP_MATRIX(i, 7) / (TEMP_MATRIX(i, 1) * TEMP_MATRIX(i, 2) / OUNCE_GRAMS_VAL)
    TEMP_MATRIX(i, 9) = TEMP_MATRIX(i, 1) * (TEMP_MATRIX(i, 2) / OUNCE_GRAMS_VAL) * (TEMP_MATRIX(i, 3) / TEMP_MATRIX(i, 4))
    TEMP_MATRIX(i, 10) = TEMP_MATRIX(i, 9) - TEMP_MATRIX(i, 7)
Next i

GOLD_MINE_PROFIT_FUNC = TEMP_MATRIX

Exit Function
ERROR_LABEL:
GOLD_MINE_PROFIT_FUNC = Err.number
End Function
