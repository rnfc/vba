Attribute VB_Name = "FINAN_FI_REAL_LIBR"


Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.


'************************************************************************************
'************************************************************************************
'FUNCTION      : REAL_NPV_FUNC
'DESCRIPTION   : Real Net Present Value
'LIBRARY       : FINAN_FI
'GROUP         : NPV
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 03/05/2009
'************************************************************************************
'************************************************************************************

Function REAL_NPV_FUNC(ByVal interest_rate As Double, _
ByVal INFLATION_RATE As Double, _
ByRef CASH_FLOW_RNG As Variant, _
Optional ByVal OUTPUT As Integer = 0)

Dim i As Long
Dim NROWS As Long

Dim REAL_RATE As Double
Dim DATA_VECTOR As Variant
Dim TEMP_MATRIX As Variant

On Error GoTo ERROR_LABEL

DATA_VECTOR = CASH_FLOW_RNG
If UBound(DATA_VECTOR, 1) = 1 Then: DATA_VECTOR = MATRIX_TRANSPOSE_FUNC(DATA_VECTOR)
NROWS = UBound(DATA_VECTOR, 1)
ReDim TEMP_MATRIX(0 To NROWS, 1 To 6)

REAL_RATE = REAL_INTEREST_RATE_FUNC(interest_rate, INFLATION_RATE)

TEMP_MATRIX(0, 1) = ("CASH FLOW")
TEMP_MATRIX(0, 2) = ("CONSTANT DOLLAR FACTOR")
TEMP_MATRIX(0, 3) = ("REAL CASH FLOW")
TEMP_MATRIX(0, 4) = ("REAL PV FACTOR")
TEMP_MATRIX(0, 5) = ("REAL PV")
TEMP_MATRIX(0, 6) = ("CUMULATIVE REAL PV")

TEMP_MATRIX(1, 1) = DATA_VECTOR(1, 1)
TEMP_MATRIX(1, 2) = 1
TEMP_MATRIX(1, 3) = TEMP_MATRIX(1, 1) * TEMP_MATRIX(1, 2)
TEMP_MATRIX(1, 4) = 1
TEMP_MATRIX(1, 5) = TEMP_MATRIX(1, 3) * TEMP_MATRIX(1, 4)
TEMP_MATRIX(1, 6) = TEMP_MATRIX(1, 5)

For i = 2 To NROWS
    TEMP_MATRIX(i, 1) = DATA_VECTOR(i, 1)
    TEMP_MATRIX(i, 2) = TEMP_MATRIX(i - 1, 2) / (1 + INFLATION_RATE) _
    'Previous Constant Dollar Factor / ( 1 + Inflation Rate )
    TEMP_MATRIX(i, 3) = TEMP_MATRIX(i, 1) * TEMP_MATRIX(i, 2)
    TEMP_MATRIX(i, 4) = TEMP_MATRIX(i - 1, 4) / (1 + REAL_RATE)
    TEMP_MATRIX(i, 5) = TEMP_MATRIX(i, 3) * TEMP_MATRIX(i, 4)
    TEMP_MATRIX(i, 6) = TEMP_MATRIX(i - 1, 6) + TEMP_MATRIX(i, 5)
Next i

Select Case OUTPUT
Case 0
    REAL_NPV_FUNC = TEMP_MATRIX
Case Else
    REAL_NPV_FUNC = TEMP_MATRIX(NROWS, 6) 'REAL NPV
End Select

Exit Function
ERROR_LABEL:
REAL_NPV_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : CONSTANT_DOLLAR_FUNC
'DESCRIPTION   : Constant Dollar Calculation
'LIBRARY       : FINAN_FI
'GROUP         : NPV
'ID            : 002
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 03/05/2009
'************************************************************************************
'************************************************************************************

Function CONSTANT_DOLLAR_FUNC(ByVal CURRENT_PRICE As Double, _
ByVal INFLATION_RATE As Double, _
ByVal TENOR As Double)

Dim i As Long
Dim TEMP_MATRIX As Variant

On Error GoTo ERROR_LABEL

ReDim TEMP_MATRIX(0 To TENOR + 1, 1 To 4)

TEMP_MATRIX(0, 1) = ("REAL PRICE")
TEMP_MATRIX(0, 2) = ("INFLATION FACTOR")
TEMP_MATRIX(0, 3) = ("CONSTANT DOLLAR FACTOR")
TEMP_MATRIX(0, 4) = ("PRICE IN CONSTANT DOLLARS")

TEMP_MATRIX(1, 1) = CURRENT_PRICE
TEMP_MATRIX(1, 2) = 1
TEMP_MATRIX(1, 3) = 1
TEMP_MATRIX(1, 4) = TEMP_MATRIX(1, 1) * TEMP_MATRIX(1, 3)

For i = 2 To TENOR + 1
    TEMP_MATRIX(i, 1) = TEMP_MATRIX(i - 1, 1) * (1 + INFLATION_RATE)
    TEMP_MATRIX(i, 2) = TEMP_MATRIX(i - 1, 2) * (1 + INFLATION_RATE)
    TEMP_MATRIX(i, 3) = TEMP_MATRIX(i - 1, 3) / (1 + INFLATION_RATE)
    'Previous Constant Dollar Factor / ( 1 + Inflation Rate )
    TEMP_MATRIX(i, 4) = TEMP_MATRIX(i, 1) * TEMP_MATRIX(i, 3)
Next i

CONSTANT_DOLLAR_FUNC = TEMP_MATRIX

Exit Function
ERROR_LABEL:
CONSTANT_DOLLAR_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : REAL_INTEREST_RATE_FUNC
'DESCRIPTION   : Real Interest Rate Calculation
'LIBRARY       : FINAN_FI
'GROUP         : NPV
'ID            : 003
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 03/05/2009
'************************************************************************************
'************************************************************************************

Function REAL_INTEREST_RATE_FUNC(ByVal interest_rate As Double, _
ByVal INFLATION_RATE As Double)
On Error GoTo ERROR_LABEL
REAL_INTEREST_RATE_FUNC = (1 + interest_rate) / (1 + INFLATION_RATE) - 1
Exit Function
ERROR_LABEL:
REAL_INTEREST_RATE_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : TODAY_DOLLAR_PRICE_FUNC
'DESCRIPTION   : Today Dollar Price
'LIBRARY       : FINAN_FI
'GROUP         : NPV
'ID            : 004
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 03/05/2009
'************************************************************************************
'************************************************************************************

Function TODAY_DOLLAR_PRICE_FUNC(ByVal CASH_FLOW As Double, _
ByVal interest_rate As Double, _
ByVal INFLATION_RATE As Double)
On Error GoTo ERROR_LABEL
TODAY_DOLLAR_PRICE_FUNC = CASH_FLOW * (1 + REAL_INTEREST_RATE_FUNC(interest_rate, INFLATION_RATE))
Exit Function
ERROR_LABEL:
TODAY_DOLLAR_PRICE_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : TOMORROW_DOLLAR_PRICE_FUNC
'DESCRIPTION   : Tomorrow Dollar Price
'LIBRARY       : FINAN_FI
'GROUP         : NPV
'ID            : 005
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 03/05/2009
'************************************************************************************
'************************************************************************************

Function TOMORROW_DOLLAR_PRICE_FUNC(ByVal CASH_FLOW As Double, _
ByVal interest_rate As Double, _
ByVal INFLATION_RATE As Double)
On Error GoTo ERROR_LABEL
TOMORROW_DOLLAR_PRICE_FUNC = CASH_FLOW * (1 + REAL_INTEREST_RATE_FUNC(interest_rate, INFLATION_RATE)) * (1 + INFLATION_RATE)
Exit Function
ERROR_LABEL:
TOMORROW_DOLLAR_PRICE_FUNC = Err.number
End Function
