Attribute VB_Name = "FINAN_FI_IRR_LIBR"

'////////////////////////////////////////////////////////////////////////////////////////////////////
'////////////////////////////////////////////////////////////////////////////////////////////////////
Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.
'////////////////////////////////////////////////////////////////////////////////////////////////////
'////////////////////////////////////////////////////////////////////////////////////////////////////
Private PUB_NPV_VAL As Double
Private PUB_NPV_VECTOR As Variant
Private PUB_DATE_VECTOR As Variant
Private PUB_DATA_VECTOR As Variant
Private Const PUB_EPSILON As Double = 2 ^ 52
'////////////////////////////////////////////////////////////////////////////////////////////////////
'////////////////////////////////////////////////////////////////////////////////////////////////////


'**********************************************************************************
'**********************************************************************************
'FUNCTION      : IRR_FUNC
'DESCRIPTION   :
'LIBRARY       : FI
'GROUP         : IRR
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 17/06/2010
'**********************************************************************************
'**********************************************************************************

Function IRR_FUNC(ByRef DATA_RNG As Variant, _
Optional ByRef DATE_RNG As Variant, _
Optional ByVal GUESS_VAL As Double = 0.1, _
Optional ByVal nLOOPS As Long = 1000, _
Optional ByVal tolerance As Double = 10 ^ -10) '0.00000000001

Dim COUNTER As Long
Dim CONVERG_VAL As Integer
Dim FORMULA_STR As String

On Error GoTo ERROR_LABEL

PUB_DATA_VECTOR = DATA_RNG
If UBound(PUB_DATA_VECTOR, 1) = 1 Then
    PUB_DATA_VECTOR = MATRIX_TRANSPOSE_FUNC(PUB_DATA_VECTOR)
End If

If IsArray(DATE_RNG) = True Then
    PUB_DATE_VECTOR = DATE_RNG
    If UBound(PUB_DATE_VECTOR, 1) = 1 Then
        PUB_DATE_VECTOR = MATRIX_TRANSPOSE_FUNC(PUB_DATE_VECTOR)
    End If
    If UBound(PUB_DATE_VECTOR, 1) <> UBound(PUB_DATA_VECTOR, 1) Then: GoTo ERROR_LABEL
    FORMULA_STR = "CALL_XIRR_OBJ_FUNC"
Else
    FORMULA_STR = "CALL_IRR_OBJ_FUNC"
End If
If GUESS_VAL <= 0 Then: GUESS_VAL = 10 ^ -1
'If GUESS_VAL >= 1 Then: GUESS_VAL = 1 - tolerance
'IRR_FUNC = NEWTON_ZERO_FUNC(GUESS_VAL, "CALL_XIRR_OBJ_FUNC", "", CONVERG_VAL, COUNTER, nLOOPS, tolerance)
'IRR_FUNC = CALL_TEST_ZERO_FRAME_FUNC(LOWER_GUESS_VAL, UPPER_GUESS_VAL, "CALL_XIRR_OBJ_FUNC")
IRR_FUNC = MULLER_ZERO_FUNC(-GUESS_VAL, GUESS_VAL, FORMULA_STR, CONVERG_VAL, COUNTER, nLOOPS, tolerance)

Exit Function
ERROR_LABEL:
IRR_FUNC = Err.number
End Function


'**********************************************************************************
'**********************************************************************************
'FUNCTION      : NPV_FUNC
'DESCRIPTION   : Returns the NPV for a schedule of cash flows that is not necessarily periodic.
'LIBRARY       : FI
'GROUP         : IRR
'ID            : 002
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 17/06/2010
'ANATOMY OF XIRR
'http://office.microsoft.com/en-us/excel/HP052093411033.aspx
'**********************************************************************************
'**********************************************************************************

Function NPV_FUNC(ByVal RATE_VAL As Double, _
ByRef DATA_RNG As Variant, _
Optional ByRef DATE_RNG As Variant, _
Optional ByVal OUTPUT As Integer = 0)

On Error GoTo ERROR_LABEL

PUB_DATA_VECTOR = DATA_RNG
If UBound(PUB_DATA_VECTOR, 1) = 1 Then
    PUB_DATA_VECTOR = MATRIX_TRANSPOSE_FUNC(PUB_DATA_VECTOR)
End If

If IsArray(DATE_RNG) = True Then
    PUB_DATE_VECTOR = DATE_RNG
    If UBound(PUB_DATE_VECTOR, 1) = 1 Then
        PUB_DATE_VECTOR = MATRIX_TRANSPOSE_FUNC(PUB_DATE_VECTOR)
    End If
    If UBound(PUB_DATE_VECTOR, 1) <> UBound(PUB_DATA_VECTOR, 1) Then: GoTo ERROR_LABEL
    Call CALL_XIRR_OBJ_FUNC(RATE_VAL)
Else
    Call CALL_IRR_OBJ_FUNC(RATE_VAL)
End If

Select Case OUTPUT
Case 0
    NPV_FUNC = PUB_NPV_VAL
Case Else
    NPV_FUNC = PUB_NPV_VECTOR
End Select

Exit Function
ERROR_LABEL:
NPV_FUNC = Err.number
End Function

'**********************************************************************************
'**********************************************************************************
'FUNCTION      : CALL_IRR_OBJ_FUNC
'DESCRIPTION   :
'LIBRARY       : FI
'GROUP         : IRR
'ID            : 003
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 17/06/2010
'**********************************************************************************
'**********************************************************************************

Private Function CALL_IRR_OBJ_FUNC(ByVal X_VAL As Double)

Dim i As Long
Dim SROW As Long
Dim NROWS As Long

On Error GoTo ERROR_LABEL

SROW = LBound(PUB_DATA_VECTOR, 1): NROWS = UBound(PUB_DATA_VECTOR, 1)
ReDim PUB_NPV_VECTOR(SROW To NROWS, 1 To 1): PUB_NPV_VAL = 0
For i = SROW To NROWS
    If PUB_DATA_VECTOR(i, 1) <> "" Then
        PUB_NPV_VECTOR(i, 1) = PUB_DATA_VECTOR(i, 1) / (1 + X_VAL) ^ (i - 1)
        PUB_NPV_VAL = PUB_NPV_VAL + PUB_NPV_VECTOR(i, 1)
    End If
Next i
CALL_IRR_OBJ_FUNC = PUB_NPV_VAL 'Abs(PUB_NPV_VAL) ^ 2

Exit Function
ERROR_LABEL:
CALL_IRR_OBJ_FUNC = PUB_EPSILON
End Function

'**********************************************************************************
'**********************************************************************************
'FUNCTION      : CALL_XIRR_OBJ_FUNC
'DESCRIPTION   : Objective function for the internal rate of return for a
'schedule of cash flows that is not necessarily periodic.
'LIBRARY       : FI
'GROUP         : IRR
'ID            : 004
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 17/06/2010
'**********************************************************************************
'**********************************************************************************

Private Function CALL_XIRR_OBJ_FUNC(ByVal X_VAL As Double)

Dim i As Long
Dim SROW As Long
Dim NROWS As Long

On Error GoTo ERROR_LABEL

SROW = LBound(PUB_DATE_VECTOR, 1): NROWS = UBound(PUB_DATE_VECTOR, 1)
ReDim PUB_NPV_VECTOR(SROW To NROWS, 1 To 2): PUB_NPV_VAL = 0
For i = SROW To NROWS
    PUB_NPV_VECTOR(i, 1) = (PUB_DATE_VECTOR(i, 1) - PUB_DATE_VECTOR(SROW, 1)) / 365
    PUB_NPV_VECTOR(i, 2) = PUB_DATA_VECTOR(i, 1) / (1 + X_VAL) ^ PUB_NPV_VECTOR(i, 1)
    PUB_NPV_VAL = PUB_NPV_VAL + PUB_NPV_VECTOR(i, 2)
Next i
CALL_XIRR_OBJ_FUNC = PUB_NPV_VAL 'Abs(PUB_NPV_VAL) ^ 2

Exit Function
ERROR_LABEL:
CALL_XIRR_OBJ_FUNC = PUB_EPSILON
End Function

'**********************************************************************************
'**********************************************************************************
'FUNCTION      : MIRR_FUNC

'DESCRIPTION   : Returns the modified internal rate of return for a series of
'periodic cash flows. MIRR considers both the cost of the investment and the
'interest received on reinvestment of cash.

'MIRR(values,finance_rate,reinvest_rate)
'Values: is an array or a reference to cells that contain numbers. These numbers represent
'a series of payments (negative values) and income (positive values) occurring at regular
'periods.

'Values must contain at least one positive value and one negative value to calculate the
'modified internal rate of return. Otherwise, MIRR returns the #DIV/0! error value.
'If an array or reference argument contains text, logical values, or empty cells, those
'values are ignored; however, cells with the value zero are included.

'Finance_rate: is the interest rate you pay on the money used in the cash flows.
'Reinvest_rate: is the interest rate you receive on the cash flows as you reinvest them.

'MIRR uses the order of values to interpret the order of cash flows. Be sure to enter your
'payment and income values in the sequence you want and with the correct signs (positive values
'for cash received, negative values for cash paid).

 'Data / Description
'-$120,000 Initial cost
'39,000 Return first year
'30,000 Return second year
'21,000 Return third year
'37,000 Return fourth year
'46,000 Return fifth year
'10.00% Annual interest rate for the 120,000 loan
'12.00% Annual interest rate for the reinvested profits
'=MIRR(A2:A7, A8, A9) Investment's modified rate of return after five years (13%)
'=MIRR(A2:A5, A8, A9) Modified rate of return after three years (-5%)
'=MIRR(A2:A7, A8, 14%) Five-year modified rate of return based on a reinvest_rate of 14 percent (13%)
 
'PV*(1+MIRR)^t = FV

'LIBRARY       : FI
'GROUP         : IRR
'ID            : 005
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 17/06/2010
'REFERENCES    : http://www.gummy-stuff.org/MIRR.htm
'**********************************************************************************
'**********************************************************************************

Function MIRR_FUNC(ByRef DATA_RNG As Variant, _
Optional ByVal FINANCE_RATE_VAL As Double = 0.05, _
Optional ByVal REINVEST_RATE_VAL As Double = 0.08, _
Optional ByVal VERSION As Integer = 1)

'-----------------------------------------------------------------------
'Finance and Reinvest Rates --> PER PERIOD!!!!
'-----------------------------------------------------------------------
'VERSION = 0:
'-----------------------------------------------------------------------
    'You borrow $x for t years @r1%
    'You make additional loans and/or invest profits @r2%
    'The final value of your enterprise is $y.
    'PV = present value of your loans
    'FV = Future Value of your invested profits.
    'There is a magic equation to solve for MIRR, namely:
    'PV*(1+MIRR)^t = FV --> that's MIRR
'-----------------------------------------------------------------------
'ELSE
'-----------------------------------------------------------------------
    'You borrow $X to buy some company.
    'Each year you withdraw some profit (entered as a negative) and/or make another loan (entered as a positive).
    'Your loans have a Finance rate of r1.
    'Your profits are invested at an Investment rate of r2.
    'Your profits amount to PV after t years.
    'However, at the end of the time period you must pay off the loan(s). L_VAL
'-----------------------------------------------------------------------
'END IF
'-----------------------------------------------------------------------

Dim i As Long
Dim j As Long
Dim NROWS As Long 'No Periods

Dim B_VAL As Double 'BankRoll
Dim L_VAL As Double 'Sum Loans
Dim P_VAL As Double 'Sum P

Dim PV_VAL As Double 'Sum PV
Dim FV_VAL As Double 'Sum FV

Dim MIRR_VAL As Double
Dim GIRR_VAL As Double
Dim GRR_VAL As Double

Dim DATA_VECTOR As Variant

On Error GoTo ERROR_LABEL

DATA_VECTOR = DATA_RNG
If UBound(DATA_VECTOR, 1) = 1 Then
    DATA_VECTOR = MATRIX_TRANSPOSE_FUNC(DATA_VECTOR)
End If
NROWS = UBound(DATA_VECTOR, 1)
j = NROWS - 1 't = 0

'----------------------------------------------------------------------------------------------------------
Select Case VERSION
'----------------------------------------------------------------------------------------------------------
Case 0
'----------------------------------------------------------------------------------------------------------
    PV_VAL = 0: FV_VAL = 0
    For i = 1 To NROWS
        If DATA_VECTOR(i, 1) > 0 Then
            FV_VAL = FV_VAL + DATA_VECTOR(i, 1) * (1 + REINVEST_RATE_VAL) ^ (j - (i - 1))
        End If
        If DATA_VECTOR(i, 1) < 0 Then
            PV_VAL = PV_VAL + -DATA_VECTOR(i, 1) / (1 + FINANCE_RATE_VAL) ^ (i - 1)
        End If
    Next i
    MIRR_VAL = (FV_VAL / PV_VAL) ^ (1 / j) - 1
    MIRR_FUNC = MIRR_VAL
'----------------------------------------------------------------------------------------------------------
Case Else
'----------------------------------------------------------------------------------------------------------
    PV_VAL = 0: FV_VAL = 0: L_VAL = 0
    For i = 1 To NROWS
        If DATA_VECTOR(i, 1) > 0 Then
            PV_VAL = PV_VAL + DATA_VECTOR(i, 1) / (1 + FINANCE_RATE_VAL) ^ (i - 1)
            P_VAL = P_VAL + DATA_VECTOR(i, 1) / (1 + REINVEST_RATE_VAL) ^ (i - 1) 'Reinvestments
            L_VAL = L_VAL + DATA_VECTOR(i, 1) * (1 + FINANCE_RATE_VAL) ^ (j - (i - 1)) 'Loans
        End If
        If DATA_VECTOR(i, 1) < 0 Then
            FV_VAL = FV_VAL + -DATA_VECTOR(i, 1) * (1 + REINVEST_RATE_VAL) ^ (j - (i - 1))
        End If
    Next i
    MIRR_VAL = (FV_VAL / PV_VAL) ^ (1 / j) - 1
    GIRR_VAL = (FV_VAL / P_VAL) ^ (1 / j) - 1
    B_VAL = L_VAL / (1 + REINVEST_RATE_VAL) ^ j 'Bankroll
    GRR_VAL = (FV_VAL / B_VAL) ^ (1 / j) - 1
    MIRR_FUNC = Array(MIRR_VAL, GIRR_VAL, GRR_VAL)
'----------------------------------------------------------------------------------------------------------
End Select
'----------------------------------------------------------------------------------------------------------

Exit Function
ERROR_LABEL:
MIRR_FUNC = Err.number
End Function