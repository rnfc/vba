Attribute VB_Name = "FINAN_FI_CMS_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.


'************************************************************************************
'************************************************************************************
'FUNCTION      : SWAP_CMS_FUNC
'DESCRIPTION   : Calculates price of constant maturity swap / with Convexity
'Adjustment. Results are compared to value mentioned in page 600 of Hull's book Options
'and Derivatives. http://www.rotman.utoronto.ca/~hull/ofod/errata/
'LIBRARY       : SWAP
'GROUP         : SWAP_CMS
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 10-06-2008
'************************************************************************************
'************************************************************************************

Function SWAP_CMS_FUNC(ByVal PRINCIPAL As Double, _
ByVal CMS_TENOR As Double, _
ByVal FIXED_RATE As Double, _
ByVal SWAP_RATE_TENOR As Double, _
ByVal FORWARD_SIGMA As Double, _
ByVal SWAP_SIGMA As Double, _
ByVal RHO_VAL As Double, _
Optional ByVal FREQUENCY As Integer = 2, _
Optional ByVal PAR_VAL As Double = 100, _
Optional ByVal MAX_TENOR As Double = 100, _
Optional ByVal epsilon As Double = 0.0001, _
Optional ByVal OUTPUT As Integer = 0)

'RHO_VAL: correlation of forward rate and swap rate
'epsilon: used to calculate first and second order derivatives

Dim i As Long
Dim j As Long

Dim NROWS As Long

Dim MIN_TENOR As Double
Dim DELTA_TENOR As Double
Dim TEMP_SUM As Double
Dim TEMP_MATRIX As Variant

On Error GoTo ERROR_LABEL

'-----------------------------------------------------------------------
'Example
'-----------------------------------------------------------------------
'Principal: $100,000,000
'CMS Maturity yrs: 6
'Fixed Rate: 5%
'Swap rate of yrs: 5
'vol forward: 20%
'vol swap: 15%
'correlation: 70%
'-----------------------------------------------------------------------
'Swap Price: 159,838.63
'-----------------------------------------------------------------------

MIN_TENOR = 0
DELTA_TENOR = 1 / FREQUENCY

NROWS = (MAX_TENOR - MIN_TENOR) / DELTA_TENOR + 1
ReDim TEMP_MATRIX(0 To NROWS, 1 To 13)

TEMP_MATRIX(0, 1) = "INDEX"
TEMP_MATRIX(0, 2) = "TENOR"
TEMP_MATRIX(0, 3) = "DELTA"
TEMP_MATRIX(0, 4) = "FORWARD RATE"
TEMP_MATRIX(0, 5) = "DISCOUNT FACTOR"
TEMP_MATRIX(0, 6) = "CUMULATIVE"
TEMP_MATRIX(0, 7) = "SWAP YIELD"
TEMP_MATRIX(0, 8) = "G'"
TEMP_MATRIX(0, 9) = "G''"
TEMP_MATRIX(0, 10) = "CONVEX ADJ"
TEMP_MATRIX(0, 11) = "TIMING ADJ"
TEMP_MATRIX(0, 12) = "TOTAL ADJ"
TEMP_MATRIX(0, 13) = "NPV" 'Net Payment Present Value

TEMP_MATRIX(1, 1) = 0
TEMP_MATRIX(1, 2) = 0
TEMP_MATRIX(1, 3) = DELTA_TENOR
TEMP_MATRIX(1, 4) = FIXED_RATE
TEMP_MATRIX(1, 5) = 1
TEMP_MATRIX(1, 6) = 0
TEMP_MATRIX(1, 7) = ""
TEMP_MATRIX(1, 8) = ""
TEMP_MATRIX(1, 9) = ""
TEMP_MATRIX(1, 10) = ""
TEMP_MATRIX(1, 11) = ""
TEMP_MATRIX(1, 12) = ""
TEMP_MATRIX(1, 13) = "" 'Net Payment Present Value


For i = 2 To NROWS
    TEMP_MATRIX(i, 1) = i - 1
    TEMP_MATRIX(i, 2) = TEMP_MATRIX(i - 1, 2) + DELTA_TENOR
    TEMP_MATRIX(i, 3) = DELTA_TENOR
    TEMP_MATRIX(i, 4) = FIXED_RATE
    TEMP_MATRIX(i, 5) = 1 / (1 + TEMP_MATRIX(i, 4) / FREQUENCY) ^ TEMP_MATRIX(i, 1)
    TEMP_MATRIX(i, 6) = TEMP_MATRIX(i - 1, 6) + TEMP_MATRIX(i, 5) * TEMP_MATRIX(i, 3)
Next i

j = Int(SWAP_RATE_TENOR / DELTA_TENOR) + 1
i = 2
Do While (j <= NROWS) And ((i + 1) <= NROWS)
    TEMP_MATRIX(i, 7) = (TEMP_MATRIX(i, 5) - TEMP_MATRIX(j, 5)) / (TEMP_MATRIX(j, 6) - TEMP_MATRIX(i, 6)) 'forward swap yield for X yr swap
    TEMP_MATRIX(i, 8) = (SWAP_CMS_GFN_FUNC(TEMP_MATRIX(i, 7) + epsilon, FIXED_RATE, SWAP_RATE_TENOR, FREQUENCY, PAR_VAL) - SWAP_CMS_GFN_FUNC(TEMP_MATRIX(i, 7) - epsilon, FIXED_RATE, SWAP_RATE_TENOR, FREQUENCY, PAR_VAL)) / (2 * epsilon)
    TEMP_MATRIX(i, 9) = (SWAP_CMS_GFN_FUNC(TEMP_MATRIX(i, 7) + epsilon, FIXED_RATE, SWAP_RATE_TENOR, FREQUENCY, PAR_VAL) - 2 * SWAP_CMS_GFN_FUNC(TEMP_MATRIX(i, 7), FIXED_RATE, SWAP_RATE_TENOR, FREQUENCY, PAR_VAL) + SWAP_CMS_GFN_FUNC(TEMP_MATRIX(i, 7) - epsilon, FIXED_RATE, SWAP_RATE_TENOR, FREQUENCY, PAR_VAL)) / (epsilon * epsilon)
    TEMP_MATRIX(i, 10) = 0.5 * TEMP_MATRIX(i, 7) * TEMP_MATRIX(i, 7) * SWAP_SIGMA * SWAP_SIGMA * TEMP_MATRIX(i, 2) * TEMP_MATRIX(i, 9) / TEMP_MATRIX(i, 8)
    TEMP_MATRIX(i, 11) = TEMP_MATRIX(i, 7) * TEMP_MATRIX(i, 3) * TEMP_MATRIX(i, 4) * RHO_VAL * SWAP_SIGMA * FORWARD_SIGMA * TEMP_MATRIX(i, 2) / (1 + TEMP_MATRIX(i, 4) * TEMP_MATRIX(i, 3))
    TEMP_MATRIX(i, 12) = TEMP_MATRIX(i, 10) + TEMP_MATRIX(i, 11)
    TEMP_MATRIX(i, 13) = -PRINCIPAL * TEMP_MATRIX(i - 1, 3) * TEMP_MATRIX(i, 12) * TEMP_MATRIX(i + 1, 5)
    j = j + 1
    i = i + 1
Loop

'--------------------------------------------------------------
Select Case OUTPUT
'--------------------------------------------------------------
Case 0 'Swap price
'--------------------------------------------------------------
    TEMP_SUM = 0
    j = Int(SWAP_RATE_TENOR / DELTA_TENOR) + 1
    For i = 2 To (j + 1)
        TEMP_SUM = TEMP_SUM + TEMP_MATRIX(i, 13)
    Next i
    
    SWAP_CMS_FUNC = TEMP_SUM
'--------------------------------------------------------------
Case Else
'--------------------------------------------------------------
    SWAP_CMS_FUNC = TEMP_MATRIX
'--------------------------------------------------------------
End Select
'--------------------------------------------------------------

Exit Function
ERROR_LABEL:
SWAP_CMS_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : SWAP_CMS_GFN_FUNC
'DESCRIPTION   : GFN PV FACTOR
'LIBRARY       : SWAP
'GROUP         : SWAP_CMS
'ID            : 002
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 10-06-2008
'************************************************************************************
'************************************************************************************

Private Function SWAP_CMS_GFN_FUNC(ByVal Y_VAL As Double, _
ByVal FIXED_RATE As Double, _
ByVal SWAP_RATE_TENOR As Long, _
Optional ByVal FREQUENCY As Integer = 2, _
Optional ByVal PAR_VAL As Double = 100)

Dim i As Long
Dim NSIZE As Long
Dim TEMP_SUM As Double

On Error GoTo ERROR_LABEL

NSIZE = SWAP_RATE_TENOR * FREQUENCY
TEMP_SUM = 0
For i = 1 To NSIZE
    TEMP_SUM = TEMP_SUM + (FIXED_RATE / FREQUENCY) * PAR_VAL / (1 + Y_VAL / 2) ^ i
Next i
TEMP_SUM = TEMP_SUM + PAR_VAL / (1 + Y_VAL / 2) ^ NSIZE
SWAP_CMS_GFN_FUNC = TEMP_SUM

Exit Function
ERROR_LABEL:
SWAP_CMS_GFN_FUNC = Err.number
End Function
