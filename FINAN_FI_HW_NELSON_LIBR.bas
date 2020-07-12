Attribute VB_Name = "FINAN_FI_HW_NELSON_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.



'**********************************************************************************
'**********************************************************************************
'FUNCTION      : HW_NELSON_OPTION_TREE_FUNC
'DESCRIPTION   : Hull-White Nelson Option Model
'LIBRARY       : FIXED_INCOME
'GROUP         : NS_OPTION
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 10-06-2008
'**********************************************************************************
'**********************************************************************************

Function HW_NELSON_OPTION_TREE_FUNC(ByVal OPT_VALUE As Double, _
ByVal POLICY As Double, _
ByVal GUARANTEE As Double, _
ByVal STEPS As Double, _
ByVal TENOR As Double, _
ByVal KAPPA As Double, _
ByVal SIGMA As Double, _
ByVal BETA0_VAL As Double, _
ByVal BETA1_VAL As Double, _
ByVal BETA2_VAL As Double, _
ByVal BETA3_VAL As Double, _
ByVal TAU1_VAL As Double, _
ByVal TAU2_VAL As Double, _
Optional ByVal VERSION As Integer = 0, _
Optional ByVal YIELD_FUNC_NAME As String = "NS_HW_ZERO_FUNC")

'The parameter SIGMA determines the short rate's instantaneous standard deviation.
'The reversion rate parameter, KAPPA, determines the rate at which standard
'deviations decline with maturity. The higher KAPPA, the greater the decline.
'When it is = 0, the model reduces the Ho and Lee, and zero-coupon bond price
'volatilities are a linear function of maturity with the instantaneous standard
'deviations of both spot and forward rates being constant.

'-------------------------------------------------------------------------------------
'In choosing the correct value for the Strike of the Bond Option, the precise terms
'of the option are, therefore, important.

'1) IF the strike price is defined as the cash amount that is exchanged for the bond
'when the option is exercised, X should be put equal to this strike price.

'2) IF, as is more common, the strike price is the quoted price applicable when
'the option is exercised, X should be set equal to teh strke pric eplus accrued
'interest at the expiration date of the option

'(Remember that traders refer to the quoted price of a bond as the "clean price" and
'the cash price as the "dirty price.")
'-------------------------------------------------------------------------------------

Dim i As Double
Dim j As Double

Dim SROW As Double
Dim NROWS As Double

Dim SCOLUMN As Double
Dim NCOLUMNS As Double

Dim PERIODS As Double

Dim DV_VAL As Double
Dim CV_VAL As Double

Dim EXERCISE_ARR As Variant
Dim SHORT_RATE_ARR As Variant
Dim TEMP_MATRIX As Variant

On Error GoTo ERROR_LABEL

SHORT_RATE_ARR = HW_NELSON_RATE_TREE_FUNC(GUARANTEE, STEPS, TENOR, KAPPA, _
                SIGMA, BETA0_VAL, BETA1_VAL, BETA2_VAL, BETA3_VAL, _
                TAU1_VAL, TAU2_VAL, YIELD_FUNC_NAME)

EXERCISE_ARR = HW_NELSON_EXERCISE_TREE_FUNC(OPT_VALUE, POLICY, GUARANTEE, _
                STEPS, TENOR, KAPPA, SIGMA, BETA0_VAL, BETA1_VAL, _
                BETA2_VAL, BETA3_VAL, _
                TAU1_VAL, TAU2_VAL, YIELD_FUNC_NAME)

PERIODS = TENOR / STEPS

NROWS = PERIODS * 6 + 7
NCOLUMNS = PERIODS * 2 + 2

ReDim TEMP_MATRIX(0 To NROWS, 1 To NCOLUMNS)

For i = 0 To PERIODS
    SROW = i * 6 + 2
    TEMP_MATRIX(SROW, 1) = STEPS * i
Next i

SROW = (PERIODS - 1) * 6 + 2
SCOLUMN = (PERIODS - 1) * 2 + 1

SROW = (PERIODS - 1) * 6 + 2
SCOLUMN = (PERIODS - 1) * 2 + 1

For j = 1 To SCOLUMN
    DV_VAL = _
        EXERCISE_ARR(SROW + 10, j + 1) * SHORT_RATE_ARR(SROW + 2, j + 1) * _
    Exp(-STEPS * SHORT_RATE_ARR(SROW + 6, j + 1)) + _
        EXERCISE_ARR(SROW + 10, j + 2) * SHORT_RATE_ARR(SROW + 3, j + 1) * _
    Exp(-STEPS * SHORT_RATE_ARR(SROW + 6, j + 2)) + _
        EXERCISE_ARR(SROW + 10, j + 3) * SHORT_RATE_ARR(SROW + 4, j + 1) * _
    Exp(-STEPS * SHORT_RATE_ARR(SROW + 6, j + 3))

    CV_VAL = EXERCISE_ARR(SROW + 4, j + 1)

    TEMP_MATRIX(SROW, j + 1) = MAXIMUM_FUNC(DV_VAL, CV_VAL)
    TEMP_MATRIX(SROW + 1, j + 1) = DV_VAL
    TEMP_MATRIX(SROW + 2, j + 1) = CV_VAL
Next j

For i = 2 To PERIODS
    SROW = (PERIODS - i) * 6 + 2
    SCOLUMN = (PERIODS - i) * 2 + 1
    For j = 1 To SCOLUMN
        DV_VAL = _
        TEMP_MATRIX(SROW + 6, j + 1) * SHORT_RATE_ARR(SROW + 2, j + 1) * _
        Exp(-STEPS * SHORT_RATE_ARR(SROW + 6, j + 1)) + _
        TEMP_MATRIX(SROW + 6, j + 2) * SHORT_RATE_ARR(SROW + 3, j + 1) * _
        Exp(-STEPS * SHORT_RATE_ARR(SROW + 6, j + 2)) + _
        TEMP_MATRIX(SROW + 6, j + 3) * SHORT_RATE_ARR(SROW + 4, j + 1) * _
        Exp(-STEPS * SHORT_RATE_ARR(SROW + 6, j + 3))

        CV_VAL = EXERCISE_ARR(SROW + 4, j + 1)
        TEMP_MATRIX(SROW, j + 1) = MAXIMUM_FUNC(DV_VAL, CV_VAL)
        TEMP_MATRIX(SROW + 1, j + 1) = DV_VAL
        TEMP_MATRIX(SROW + 2, j + 1) = CV_VAL
    Next j
Next i

'----------------------------------HOUSE_KEEPING-----------------------------------
For i = 0 To NROWS
    For j = 1 To NCOLUMNS
        If IsEmpty(TEMP_MATRIX(i, j)) = True Then: TEMP_MATRIX(i, j) = ""
    Next j
Next i
'----------------------------------------------------------------------------------

Select Case VERSION
    Case 0
        If EXERCISE_ARR(UBound(EXERCISE_ARR, 1), 1) > _
            EXERCISE_ARR(UBound(EXERCISE_ARR, 1), 2) Then
            HW_NELSON_OPTION_TREE_FUNC = TEMP_MATRIX(2, 2) & ": Guarantee unrealistic!"
        Else
            HW_NELSON_OPTION_TREE_FUNC = TEMP_MATRIX(2, 2) 'Option Value
        End If
    Case 1
        HW_NELSON_OPTION_TREE_FUNC = TEMP_MATRIX(3, 2) 'Option value if hold to the next step
    Case 2
        HW_NELSON_OPTION_TREE_FUNC = TEMP_MATRIX(4, 2)
        'Option value if surrenderred now = intrinsic value
    Case Else
        HW_NELSON_OPTION_TREE_FUNC = TEMP_MATRIX
End Select

Exit Function
ERROR_LABEL:
HW_NELSON_OPTION_TREE_FUNC = Err.number
End Function

'**********************************************************************************
'**********************************************************************************
'FUNCTION      : HW_NELSON_EXERCISE_TREE_FUNC
'DESCRIPTION   : A GENERAL HW EXERCISE TREE_BUILDING PROCEDURE.
'LIBRARY       : FIXED_INCOME
'GROUP         : NS_OPTION
'ID            : 002
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 10-06-2008
'**********************************************************************************
'**********************************************************************************


Function HW_NELSON_EXERCISE_TREE_FUNC(ByVal OPT_VALUE As Double, _
ByVal POLICY As Double, _
ByVal GUARANTEE As Double, _
ByVal STEPS As Double, _
ByVal TENOR As Double, _
ByVal KAPPA As Double, _
ByVal SIGMA As Double, _
ByVal BETA0_VAL As Double, _
ByVal BETA1_VAL As Double, _
ByVal BETA2_VAL As Double, _
ByVal BETA3_VAL As Double, _
ByVal TAU1_VAL As Double, _
ByVal TAU2_VAL As Double, _
Optional ByVal YIELD_FUNC_NAME As String = "NS_HW_ZERO_FUNC")

Dim i As Double
Dim j As Double

Dim SROW As Double
Dim NROWS As Double

Dim SCOLUMN As Double
Dim NCOLUMNS As Double

Dim PERIODS As Double

Dim UPO_VAL As Double ' units of policies
Dim UBO_VAL As Double ' units of bonds

Dim POLICY_VAL As Double ' policy initial value
Dim BOND_VAL As Double ' bond initial value

Dim P_VAL As Double
Dim B_VAL As Double

Dim SHORT_RATE_ARR As Variant
Dim TEMP_MATRIX As Variant

On Error GoTo ERROR_LABEL

SHORT_RATE_ARR = HW_NELSON_RATE_TREE_FUNC(GUARANTEE, STEPS, TENOR, KAPPA, _
    SIGMA, BETA0_VAL, BETA1_VAL, BETA2_VAL, BETA3_VAL, TAU1_VAL, _
    TAU2_VAL, YIELD_FUNC_NAME)

POLICY_VAL = SHORT_RATE_ARR(7, 1) 'GV or Policy Value
BOND_VAL = SHORT_RATE_ARR(7, 2) 'AS

B_VAL = OPT_VALUE
P_VAL = POLICY * B_VAL

PERIODS = TENOR / STEPS

NROWS = PERIODS * 6 + 7
NCOLUMNS = PERIODS * 2 + 2

ReDim TEMP_MATRIX(0 To NROWS, 1 To NCOLUMNS)
UPO_VAL = P_VAL / POLICY_VAL
UBO_VAL = B_VAL / BOND_VAL


'--------------------------Intrisic value of option = max (0, GV - AS)-----------------

For i = 0 To PERIODS
    SROW = i * 6 + 2
    SCOLUMN = i * 2 + 1
    TEMP_MATRIX(SROW, 1) = SHORT_RATE_ARR(SROW, 1)
    TEMP_MATRIX(SROW + 5, 1) = SHORT_RATE_ARR(SROW + 5, 1) * UPO_VAL
    For j = 1 To SCOLUMN
        TEMP_MATRIX(SROW + 5, j + 1) = SHORT_RATE_ARR(SROW + 5, j + 1) * UBO_VAL
    Next j
Next i

' -----------------------------------Option payoff-------------------------------------
For i = 0 To PERIODS
    SROW = i * 6 + 2
    SCOLUMN = i * 2 + 1
    For j = 1 To SCOLUMN
        TEMP_MATRIX(SROW + 4, j + 1) = MAXIMUM_FUNC(0, _
        TEMP_MATRIX(SROW + 5, 1) - TEMP_MATRIX(SROW + 5, j + 1))
    Next j
Next i

'--------------------------------SOME HOUSE KEEPING------------------------------------
For i = 0 To NROWS
    For j = 1 To NCOLUMNS
        If IsEmpty(TEMP_MATRIX(i, j)) = True Then: TEMP_MATRIX(i, j) = ""
    Next j
Next i
'--------------------------------------------------------------------------------------

HW_NELSON_EXERCISE_TREE_FUNC = TEMP_MATRIX

Exit Function
ERROR_LABEL:
    HW_NELSON_EXERCISE_TREE_FUNC = Err.number
End Function


'**********************************************************************************
'**********************************************************************************
'FUNCTION      : HW_NELSON_RATE_TREE_FUNC

'DESCRIPTION   : The following grid is widely used for pricing instruments
'when the simpler one-factor models are inappropriate.
'They are easy to implement and, if used carefully, can ensure that
'most non-standard interest rate derivatives are priced consistently
'with actively traded instruments such as interest rate caps, European
'swap options, and European bond options. Two limitations of the models are:
'  1.  They involve one factor (that is, one source of uncertainty).
'  2.  They do not give the user complete freedom in choosing the volatility
'  structure.

'LIBRARY       : FIXED_INCOME
'GROUP         : NS_OPTION
'ID            : 003
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 10-06-2008
'**********************************************************************************
'**********************************************************************************

Function HW_NELSON_RATE_TREE_FUNC(ByVal GUARANTEE As Double, _
ByVal STEPS As Double, _
ByVal TENOR As Double, _
ByVal KAPPA As Double, _
ByVal SIGMA As Double, _
ByVal BETA0_VAL As Double, _
ByVal BETA1_VAL As Double, _
ByVal BETA2_VAL As Double, _
ByVal BETA3_VAL As Double, _
ByVal TAU1_VAL As Double, _
ByVal TAU2_VAL As Double, _
Optional ByVal YIELD_FUNC_NAME As String = "NS_HW_ZERO_FUNC")

Dim i As Double
Dim j As Double

Dim SROW As Double
Dim SCOLUMN As Double

Dim NROWS As Double
Dim NCOLUMNS As Double

Dim START_RATE As Double ' start RATE

Dim PERIODS As Double

Dim SHORT_RATE As Double ' short RATE
Dim DELTA_RATE As Double ' increment in short RATE

Dim UPPER_VAL As Double ' barrier value
Dim LOWER_VAL As Double

Dim TEMP_MATRIX As Variant

On Error GoTo ERROR_LABEL

UPPER_VAL = 1.84 / (KAPPA * STEPS)
LOWER_VAL = -UPPER_VAL

DELTA_RATE = SIGMA * (3 * STEPS)

PERIODS = TENOR / STEPS

NROWS = PERIODS * 6 + 7
NCOLUMNS = PERIODS * 2 + 2

START_RATE = BETA0_VAL + BETA1_VAL + BETA2_VAL

ReDim TEMP_MATRIX(0 To NROWS, 1 To NCOLUMNS)

For i = 0 To PERIODS
    SROW = i * 6 + 2
    TEMP_MATRIX(SROW, 1) = i * STEPS
    TEMP_MATRIX(SROW + 5, 1) = Exp(-GUARANTEE * (TENOR - i * STEPS))

    If i = 0 Then
        TEMP_MATRIX(SROW, 2) = START_RATE
        TEMP_MATRIX(SROW + 1, 2) = 0
        TEMP_MATRIX(SROW + 5, 2) = Excel.Application.Run(YIELD_FUNC_NAME, i * STEPS, TENOR, KAPPA, _
        SIGMA, BETA0_VAL, BETA1_VAL, BETA2_VAL, BETA3_VAL, TAU1_VAL, _
        TAU2_VAL, TEMP_MATRIX(SROW, 2))
        '-----> YOU CAN CHANGE THIS
    Else
        TEMP_MATRIX(SROW, 2) = TEMP_MATRIX(SROW - 6, 2) + DELTA_RATE
        TEMP_MATRIX(SROW + 1, 2) = TEMP_MATRIX(SROW - 5, 2) + 1
        TEMP_MATRIX(SROW + 5, 2) = Excel.Application.Run(YIELD_FUNC_NAME, i * STEPS, TENOR, KAPPA, _
            SIGMA, BETA0_VAL, BETA1_VAL, BETA2_VAL, BETA3_VAL, TAU1_VAL, _
            TAU2_VAL, TEMP_MATRIX(SROW, 2))
        '-----> YOU CAN CHANGE THIS
    End If

    SCOLUMN = i * 2 + 1
    For j = 1 To SCOLUMN
        If j > 1 Then
            TEMP_MATRIX(SROW, j + 1) = TEMP_MATRIX(SROW, j) - DELTA_RATE
            TEMP_MATRIX(SROW + 1, j + 1) = TEMP_MATRIX(SROW + 1, j) - 1
            TEMP_MATRIX(SROW + 5, j + 1) = Excel.Application.Run(YIELD_FUNC_NAME, i * STEPS, TENOR, _
                KAPPA, SIGMA, BETA0_VAL, BETA1_VAL, BETA2_VAL, BETA3_VAL, _
                TAU1_VAL, TAU2_VAL, TEMP_MATRIX(SROW, j + 1))
        End If
    Next j
Next i

PERIODS = TENOR / STEPS

For i = 0 To PERIODS - 1 ' Determine the trinomial tree risk-neutral probabilities

    SROW = i * 6
    SCOLUMN = i * 2 + 1
    For j = 1 To SCOLUMN

    SHORT_RATE = TEMP_MATRIX(SROW + 2, j + 1)
    
    If SHORT_RATE > UPPER_VAL Then
        TEMP_MATRIX(SROW + 4, j + 1) = 7 / 6 + _
        0.5 * (KAPPA ^ 2 * STEPS ^ 2 * TEMP_MATRIX(SROW + 3, j + 1) ^ 2 - 3 _
        * KAPPA * STEPS * TEMP_MATRIX(SROW + 3, j + 1))
       TEMP_MATRIX(SROW + 5, j + 1) = -1 / 3 - 1 * (KAPPA ^ 2 * STEPS ^ 2 * _
        TEMP_MATRIX(SROW + 3, j + 1) ^ 2 - 2 * KAPPA * _
        STEPS * TEMP_MATRIX(SROW + 3, j + 1))
        TEMP_MATRIX(SROW + 6, j + 1) = 1 / 6 + 0.5 * _
        (KAPPA ^ 2 * STEPS ^ 2 * _
        TEMP_MATRIX(SROW + 3, j + 1) ^ 2 - 1 * KAPPA * _
        STEPS * TEMP_MATRIX(SROW + 3, j + 1))
    ElseIf SHORT_RATE < LOWER_VAL Then
        TEMP_MATRIX(SROW + 4, j + 1) = 1 / 6 + 0.5 * _
        (KAPPA ^ 2 * STEPS ^ 2 * _
        TEMP_MATRIX(SROW + 3, j + 1) ^ 2 + 1 * KAPPA * _
        STEPS * TEMP_MATRIX(SROW + 3, j + 1))
        TEMP_MATRIX(SROW + 5, j + 1) = -1 / 3 - 1 * _
        (KAPPA ^ 2 * STEPS ^ 2 * _
        TEMP_MATRIX(SROW + 3, j + 1) ^ 2 + 2 * KAPPA * _
        STEPS * TEMP_MATRIX(SROW + 3, j + 1))
        TEMP_MATRIX(SROW + 6, j + 1) = 7 / 6 + 0.5 * _
        (KAPPA ^ 2 * STEPS ^ 2 * _
        TEMP_MATRIX(SROW + 3, j + 1) ^ 2 + 3 * KAPPA * _
        STEPS * TEMP_MATRIX(SROW + 3, j + 1))
    Else
        TEMP_MATRIX(SROW + 4, j + 1) = 1 / 6 + 0.5 * _
        (KAPPA ^ 2 * STEPS ^ 2 * _
        TEMP_MATRIX(SROW + 3, j + 1) ^ 2 - 1 * KAPPA * _
        STEPS * TEMP_MATRIX(SROW + 3, j + 1))
        TEMP_MATRIX(SROW + 5, j + 1) = 2 / 3 - 1 * _
        (KAPPA ^ 2 * STEPS ^ 2 * _
        TEMP_MATRIX(SROW + 3, j + 1) ^ 2 - 0 * KAPPA * _
        STEPS * TEMP_MATRIX(SROW + 3, j + 1))
        TEMP_MATRIX(SROW + 6, j + 1) = 1 / 6 + 0.5 * _
        (KAPPA ^ 2 * STEPS ^ 2 * _
        TEMP_MATRIX(SROW + 3, j + 1) ^ 2 + 1 * KAPPA * _
        STEPS * TEMP_MATRIX(SROW + 3, j + 1))
    End If
    Next j
Next i

'--------------------------------SOME HOUSE_KEEEPING---------------------------------
For i = 0 To NROWS
    For j = 1 To NCOLUMNS
        If IsEmpty(TEMP_MATRIX(i, j)) = True Then: TEMP_MATRIX(i, j) = ""
    Next j
Next i
'------------------------------------------------------------------------------------

HW_NELSON_RATE_TREE_FUNC = TEMP_MATRIX

Exit Function
ERROR_LABEL:
HW_NELSON_RATE_TREE_FUNC = Err.number
End Function


'**********************************************************************************
'**********************************************************************************
'FUNCTION      : HW_NELSON_OPTION_GUARANTEE_FUNC
'DESCRIPTION   : OPTION TEST GUARANTEE
'LIBRARY       : FIXED_INCOME
'GROUP         : NS_OPTION
'ID            : 004
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 10-06-2008
'**********************************************************************************
'**********************************************************************************

Function HW_NELSON_OPTION_GUARANTEE_FUNC(ByVal GUARANTEE As Double, _
ByVal OPT_VALUE As Double, _
ByVal POLICY As Double, _
ByVal TENOR As Double)

On Error GoTo ERROR_LABEL

HW_NELSON_OPTION_GUARANTEE_FUNC = Exp(TENOR * GUARANTEE) * OPT_VALUE * POLICY

'OPT_VALUE * POLICY --> INITIAL GUARANTEE

Exit Function
ERROR_LABEL:
HW_NELSON_OPTION_GUARANTEE_FUNC = Err.number
End Function
