Attribute VB_Name = "FINAN_ENERGY_OIL_LIBR"


Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.

'************************************************************************************
'************************************************************************************
'FUNCTION      : AVG_BARREL_VALUE_FUNC
'DESCRIPTION   : AVG. VALUE OF A BARREL OF OIL EXTRACTED
'LIBRARY       : ENERGY
'GROUP         : OIL
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/29/2009
'************************************************************************************
'************************************************************************************

Function AVG_BARREL_VALUE_FUNC(ByVal NET_VALUE_BARREL As Double, _
ByVal EXTRACTION_RATE As Double, _
ByVal DISCOUNT_FACTOR As Double, _
Optional ByVal FACTOR_VAL As Double = 0.5)

Dim BARREL_VAL As Double
Dim HALF_LIFE As Double ''The value of a barrel of oil in the
'ground depends on the time taken to get it above ground.

On Error GoTo ERROR_LABEL

HALF_LIFE = Log(FACTOR_VAL) / Log(1 - EXTRACTION_RATE)
'So, a barrel of oil will, on average,
'be sold in HALF_LIFE YEARS
BARREL_VAL = NET_VALUE_BARREL * 1 / (1 + DISCOUNT_FACTOR) ^ HALF_LIFE

AVG_BARREL_VALUE_FUNC = Array(HALF_LIFE, BARREL_VAL)

Exit Function
ERROR_LABEL:
AVG_BARREL_VALUE_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : STORABLE_COMMODITIES_FUTURES_FUNC
'DESCRIPTION   : FUTURES ON STORABLE COMMODITIES
'LIBRARY       : ENERGY
'GROUP         : OIL
'ID            : 002
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/29/2009
'************************************************************************************
'************************************************************************************

Function STORABLE_COMMODITIES_FUTURES_FUNC(ByVal SPOT_PRICE As Double, _
ByVal BORROWING_RATE As Double, _
ByVal LENDING_RATE As Double, _
ByVal STORAGE_COST As Double, _
ByVal COST_SHORT_SALES As Double, _
ByVal TENOR As Double, _
ByVal PERCENTAGE_COLLECTED_SHORT_SELLER As Double)

'Current spot price
'Borrowing rate (annualized)
'Lending rate (annualized)

'STORAGE_COST: Storage cost (In $/unit per year)

'COST_SHORT_SALES: Transactions cost for short sales (per unit)

'TENOR: Time to expiration (In years)

'PERCENTAGE_COLLECTED_SHORT_SELLER: _
% of storage cost collected by short seller

Dim LOWER_VAL As Double
Dim UPPER_VAL As Double

On Error GoTo ERROR_LABEL

UPPER_VAL = SPOT_PRICE * (1 + BORROWING_RATE) ^ TENOR + STORAGE_COST * TENOR
LOWER_VAL = (SPOT_PRICE - COST_SHORT_SALES) * (1 + LENDING_RATE) ^ TENOR + _
            STORAGE_COST * TENOR * PERCENTAGE_COLLECTED_SHORT_SELLER

STORABLE_COMMODITIES_FUTURES_FUNC = Array(LOWER_VAL, UPPER_VAL)

Exit Function
ERROR_LABEL:
STORABLE_COMMODITIES_FUTURES_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : OIL_RESERVE_LEVEL_FUNC
'DESCRIPTION   : Here’s what reserves will be in the following years
'LIBRARY       : ENERGY
'GROUP         : OIL
'ID            : 003
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/29/2009
'************************************************************************************
'************************************************************************************


Function OIL_RESERVE_LEVEL_FUNC(ByVal BARRELS_TODAY As Double, _
ByRef EXTRACTION_RNG As Variant)

Dim i As Long
Dim NROWS As Long

Dim DATA_VECTOR As Variant
Dim TEMP_MATRIX As Variant

On Error GoTo ERROR_LABEL

DATA_VECTOR = EXTRACTION_RNG 'Extraction Rates
If UBound(DATA_VECTOR, 1) = 1 Then
    DATA_VECTOR = MATRIX_TRANSPOSE_FUNC(DATA_VECTOR)
End If

NROWS = UBound(DATA_VECTOR, 1)

ReDim TEMP_MATRIX(0 To NROWS, 1 To 4)

TEMP_MATRIX(0, 1) = 0 'Periods
TEMP_MATRIX(0, 2) = BARRELS_TODAY ' BARRELS_LEFT
TEMP_MATRIX(0, 3) = "" 'Extraction Rate
TEMP_MATRIX(0, 4) = "" 'Extraction Value

For i = 1 To NROWS
    TEMP_MATRIX(i, 1) = i
    TEMP_MATRIX(i, 2) = TEMP_MATRIX(i - 1, 2) * (1 - DATA_VECTOR(i, 1))
    TEMP_MATRIX(i, 3) = DATA_VECTOR(i, 1)
    TEMP_MATRIX(i, 4) = TEMP_MATRIX(i - 1, 2) - TEMP_MATRIX(i, 2)
Next i

OIL_RESERVE_LEVEL_FUNC = TEMP_MATRIX

Exit Function
ERROR_LABEL:
OIL_RESERVE_LEVEL_FUNC = Err.number
End Function
