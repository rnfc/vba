Attribute VB_Name = "FINAN_FI_FED_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.


'************************************************************************************
'************************************************************************************
'FUNCTION      : FED_RATES_PROBABILITIES_FUNC
'DESCRIPTION   : Fred Rates Table
'LIBRARY       : FIXED INCOME
'GROUP         : FED
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 10-06-2008
'************************************************************************************
'************************************************************************************

Function FED_RATES_PROBABILITIES_FUNC(ByRef DATA_RNG As Variant, _
Optional ByVal FACTOR As Long = 100)

Dim j As Long
Dim NCOLUMNS As Long

Dim TEMP_VECTOR As Variant
Dim TEMP_MATRIX As Variant

On Error GoTo ERROR_LABEL

TEMP_MATRIX = DATA_RNG

'ROW 1: Current Fed Rate (%)
'ROW 2: New Target Fed Rate (%)
'ROW 3: Fed Futures
'ROW 4: Number of days in Month
'ROW 5: Current day of the Month

NCOLUMNS = UBound(TEMP_MATRIX, 2)
ReDim TEMP_VECTOR(1 To 1, 1 To NCOLUMNS)

For j = 1 To NCOLUMNS
    TEMP_VECTOR(1, j) = FED_RATES_PREDICTION_FUNC(TEMP_MATRIX(1, j), _
                        TEMP_MATRIX(2, j), TEMP_MATRIX(3, j), _
                        TEMP_MATRIX(4, j), TEMP_MATRIX(5, j), FACTOR)
Next j

FED_RATES_PROBABILITIES_FUNC = TEMP_VECTOR

Exit Function
ERROR_LABEL:
FED_RATES_PROBABILITIES_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : FED_RATES_PROBABILITIES_FUNC

'DESCRIPTION   : Predicting Fred Rates
'1) http://www.clevelandfed.org/Research/policy/fedfunds/index.cfm
'3) http://www.bus.ucf.edu/ssmith/FedFundsProbability.doc
'2) http://www.cbot.com/cbot/pub/cont_detail/0,3206,991+23425,00.html

'LIBRARY       : FIXED INCOME
'GROUP         : FED
'ID            : 002
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 10-06-2008
'************************************************************************************
'************************************************************************************


Function FED_RATES_PREDICTION_FUNC(ByVal CURR_FED_RATE As Double, _
ByVal NEW_TARG_FED_RATE As Double, _
ByVal FED_FUTURE As Double, _
ByVal NO_DAYS_MONTH As Long, _
ByVal CURR_DAY_MONTH As Long, _
Optional ByVal FACTOR As Long = 100)

'For the Fed Future check: http://www.cbot.com/cbot/pub/page/0,3181,1525,00.html

On Error GoTo ERROR_LABEL

FED_RATES_PREDICTION_FUNC = (((FACTOR - FED_FUTURE) / FACTOR - _
                CURR_FED_RATE) / (NEW_TARG_FED_RATE - CURR_FED_RATE)) * _
                (NO_DAYS_MONTH / (NO_DAYS_MONTH - CURR_DAY_MONTH))

Exit Function
ERROR_LABEL:
FED_RATES_PREDICTION_FUNC = Err.number
End Function

