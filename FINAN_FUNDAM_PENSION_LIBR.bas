Attribute VB_Name = "FINAN_FUNDAM_PENSION_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.



'************************************************************************************
'************************************************************************************
'FUNCTION      : PENSION_SIMULATION_FUNC

'DESCRIPTION   : This function shows a simplistic Monte Carlo Simulation
'of savings for retirement.

'OUTPUT: Retirement accumulation and annuity values are determined.
'LIMITATIONS OF THE FUNCTION: To take into account the fact that contributions
'occur throughout the year, the model assumes that contributions to the risky
'asset account made during a given year earn half the return (positive or
'negative) of the account. But, this model DOES NOT allow you to change the
'rate at which you save for retirement at different points in your career nor
'your allocation between risky and risk-free assets.

'LIBRARY       : FUNDAMENTAL
'GROUP         : PENSION
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 03/05/2009
'************************************************************************************
'************************************************************************************

Function PENSION_SIMULATION_FUNC(ByVal nLOOPS As Long, _
ByVal CURRENT_AGE As Double, _
ByVal RETIREMENT_AGE As Double, _
ByVal LIFE_EXPECTANCY As Double, _
ByVal CURRENT_SALARY As Double, _
ByVal SALARY_GROWTH_RATE As Double, _
ByVal SAVINGS_RATE As Double, _
ByVal PORT_ALLOCATION As Double, _
ByVal SAFE_ASSET_RETURN As Double, _
ByVal PORT_RETURN As Double, _
ByVal PORT_SIGMA As Double, _
ByVal CURRENT_WEALTH As Double, _
Optional ByVal FACTOR As Double = 100, _
Optional ByVal RANDOM_FLAG As Boolean = True)

'KEY NOTES:

'(1) nLOOPS: Number of repetitions in the Monte Carlo analysis

'(2) CURRENT AGE: Your current age.

'(3) RETIREMENT AGE: The age at which you expect to retire.

'(4) LIFE EXPECTANCY: Your life expectancy.  The model is inflexible--if you _
' choose 105, it assumes you will live exactly that long. Your life expectancy
' minus the retirement age determines the number of years of retirement and
' therefore how long your savings must last.

'(5) CURRENT SALARY: Your current salary.

'(6) SALARY GROWTH RATE (drift): The rate at which you assume your salary
'    will grow in real terms.

'(7) SAVINGS RATE: The rate at which you are saving for retirement.

'(9) PORT_ALLOCATION: The asset allocation parameter is the fraction of
'    your retirement portfolio that you put into risky assets (e.g.,
'    stock accounts),  The rest goes into the risk-free account.

'(10) SAFE_ASSET_RETURN: You can also select the return to the inflation-indexed
'     account. Since the real return is constant, the SD is zero.

'11) PORT_RETURN: 'You can select the average return to risky assets and the
'    standard deviation of returns. I will recommend dealing with real, as
'    opposed to nominal, returns. I am from the Dominican Republic and I have
'    seen inflation rates going over the skies (over 60% percent in just one day).

'12) PORT_SIGMA: Your portfolio standard deviation. Remember that standard deviation
'    is just a convenient measure of how far returns deviate from their average. Big
'    standard deviations DOES NOT MEAN a big probability of a loss ... or even big
'    uncertainty.

'13) CURRENT_WEALTH: The current accumulation in your retirement account.

Dim i As Long
Dim j As Long

Dim ANNUITY_FACTOR As Double
Dim EXPECTED_RETURN As Double
Dim TEMP_CONTRIBUTION As Double
Dim TEMP_WEALTH As Double

Dim TENOR As Double
Dim TEMP_MATRIX As Variant
Dim NORMAL_RANDOM_MATRIX As Variant

On Error GoTo ERROR_LABEL

TENOR = RETIREMENT_AGE + 1 - CURRENT_AGE ' number of years in simulation
'THE FORMULA FOR RETIREMENT TENOR --> LIFE_EXPECTANCY + 1 - CURRENT_AGE

ReDim TEMP_MATRIX(1 To nLOOPS, 1 To 2)
ReDim NORMAL_RANDOM_MATRIX(1 To TENOR, 1 To 1) ' Results: final wealth and annuity

ANNUITY_FACTOR = CURRENT_WEALTH / ((1 - (1 / (1 + SAFE_ASSET_RETURN) ^ _
(LIFE_EXPECTANCY - RETIREMENT_AGE + 1))) / SAFE_ASSET_RETURN)

ANNUITY_FACTOR = ANNUITY_FACTOR / FACTOR

If RANDOM_FLAG = True Then: Randomize

For j = 1 To nLOOPS
     
     TEMP_CONTRIBUTION = CURRENT_SALARY * SAVINGS_RATE
     TEMP_WEALTH = CURRENT_WEALTH
     
     NORMAL_RANDOM_MATRIX = MATRIX_RANDOM_NORMAL_FUNC(TENOR, 1, 0, PORT_RETURN, PORT_SIGMA, 0)
     
     For i = 1 To TENOR
         EXPECTED_RETURN = PORT_ALLOCATION * NORMAL_RANDOM_MATRIX(i, 1) + (1 - PORT_ALLOCATION) * SAFE_ASSET_RETURN
         TEMP_CONTRIBUTION = TEMP_CONTRIBUTION * (1 + SALARY_GROWTH_RATE)
         TEMP_WEALTH = (TEMP_WEALTH + 0.5 * TEMP_CONTRIBUTION) * Exp(EXPECTED_RETURN) + 0.5 * TEMP_CONTRIBUTION  ' lognormal returns
     Next i
     
     TEMP_MATRIX(j, 1) = TEMP_WEALTH
     TEMP_MATRIX(j, 2) = ANNUITY_FACTOR * TEMP_WEALTH

Next j

PENSION_SIMULATION_FUNC = TEMP_MATRIX

Exit Function
ERROR_LABEL:
PENSION_SIMULATION_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : PENSION_WITHDRAWAL_RATE_FUNC

'DESCRIPTION   : WITHDRAWAL RATE FUNCTION
'You 'll note that, if you expect to withdraw from your portfolio at
'some "safe" rate when you retire, then you may have to invest a LARGE
'fraction of your current salary !!!

'LIBRARY       : FUNDAMENTAL
'GROUP         : PENSION
'ID            : 002
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 03/05/2009
'************************************************************************************
'************************************************************************************

Function PENSION_WITHDRAWAL_RATE_FUNC(ByVal CURRENT_AGE As Double, _
ByVal AGE_RETIREMENT As Double, _
ByVal AGE_DEATH As Double, _
ByVal NET_INCOME As Double, _
ByVal TAX_RATE As Double, _
ByVal PORT_ROI As Double, _
ByVal INFLATION_RATE As Double, _
Optional ByVal OUTPUT As Integer = 1)

'CURRENT_AGE: Your Current Age
'AGE_RETIREMENT: Age at Retirement
'AGE_DEATH: Age at Death
'NET_INCOME: After-tax Income you need NOW
'TAX_RATE: Tax RATE
'PORT_ROI: Return on Investment
'INFLATION_RATE: Assumed Inflation

Dim ATEMP_VECTOR As Variant
Dim BTEMP_VECTOR As Variant

On Error GoTo ERROR_LABEL

ReDim ATEMP_VECTOR(1 To 5, 1 To 3)

ATEMP_VECTOR(1, 1) = "Real Return"
ATEMP_VECTOR(1, 2) = (PORT_ROI - INFLATION_RATE) / (1 + INFLATION_RATE)
ATEMP_VECTOR(1, 3) = "reduced from " & Format(PORT_ROI, "0.0%") & " because of inflation @ " & Format(INFLATION_RATE, "0.0%")

ATEMP_VECTOR(2, 1) = "Before-tax Income  you need NOW"
ATEMP_VECTOR(2, 2) = NET_INCOME / (1 - TAX_RATE)
ATEMP_VECTOR(2, 3) = "at your current age of " & Format(CURRENT_AGE, "0") & " (increased from " & Format(NET_INCOME, "$0,000") & " because of tax rate @ " & Format(TAX_RATE, "0%") & ")"

ATEMP_VECTOR(3, 1) = "Before-tax Income you'll need"
ATEMP_VECTOR(3, 2) = ATEMP_VECTOR(2, 2) * (1 + INFLATION_RATE) ^ (AGE_RETIREMENT - CURRENT_AGE)
ATEMP_VECTOR(3, 3) = "at your retirement age of " & Format(AGE_RETIREMENT, "0") & " (increased from " & Format(ATEMP_VECTOR(2, 2), "$0,000") & " because of inflation @ " & Format(INFLATION_RATE, "0%") & ")"

ATEMP_VECTOR(4, 1) = "Required Portfolio you'll need"
ATEMP_VECTOR(4, 2) = ATEMP_VECTOR(3, 2) * (1 - (1 + ATEMP_VECTOR(1, 2)) ^ (AGE_RETIREMENT - AGE_DEATH)) / ATEMP_VECTOR(1, 2)
ATEMP_VECTOR(4, 3) = "at your retirement age of " & Format(AGE_RETIREMENT, "0") & " so portfolio will last to age " & Format(AGE_DEATH, "0") & " with annual withdrawals of " & Format(ATEMP_VECTOR(3, 2), "$0,000") & " (assuming your portfolio grows at a Real Return of " & Format(ATEMP_VECTOR(1, 2), "0.00%") & ")"

ATEMP_VECTOR(5, 1) = "WITHDRAWAL RATE"
ATEMP_VECTOR(5, 2) = ATEMP_VECTOR(3, 2) / ATEMP_VECTOR(4, 2)
ATEMP_VECTOR(5, 3) = "at your retirement age of " & Format(AGE_RETIREMENT, "0") & " as a percentage of " & Format(ATEMP_VECTOR(4, 2), "$0,000")

ReDim BTEMP_VECTOR(1 To 5, 1 To 1)

BTEMP_VECTOR(1, 1) = "  You are " & Format(CURRENT_AGE, "0") & _
" years old and can live comfortably on " & Format(NET_INCOME, "$0,000") & _
" after-taxes, meaning " & Format(ATEMP_VECTOR(2, 2), "$0,000") & _
" before-taxes @ " _
& Format(TAX_RATE, "0%") & "."

BTEMP_VECTOR(2, 1) = "  At an inflation rate of " & _
Format(INFLATION_RATE, "0.0%") & _
", that means a before-tax income of " & Format(ATEMP_VECTOR(3, 2), _
"$0,000") & _
" by the time you reach age " & Format(AGE_RETIREMENT, "0") & "."

BTEMP_VECTOR(3, 1) = "  If your portfolio grows at a Real Return Rate of " & _
Format(ATEMP_VECTOR(1, 2), "0.00%") & _
" and you want your portfolio to last to age " & _
Format(AGE_DEATH, "0") & ","

BTEMP_VECTOR(4, 1) = "  then, by age " & Format(AGE_RETIREMENT, "0") & _
", you must have a portfolio of " & Format(ATEMP_VECTOR(4, 2), "$0,000") & _
" so that you can withdraw " & Format(ATEMP_VECTOR(3, 2), "$0,000") & " per year."

BTEMP_VECTOR(5, 1) = "  That corresponds to a withdrawal rate (initially) of " & _
Format(ATEMP_VECTOR(5, 2), "0.00%") & "."

Select Case OUTPUT
    Case 0
        PENSION_WITHDRAWAL_RATE_FUNC = ATEMP_VECTOR
    Case Else
        PENSION_WITHDRAWAL_RATE_FUNC = BTEMP_VECTOR
End Select

Exit Function
ERROR_LABEL:
PENSION_WITHDRAWAL_RATE_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : LIFE_INCOME_FUND_FUNCTION
'DESCRIPTION   : LIFE INCOME FUND CALCULATOR
'LIBRARY       : FUNDAMENTAL
'GROUP         : PENSION
'ID            : 003
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 03/05/2009
'************************************************************************************
'************************************************************************************

Function LIF_MAX_WITHDRAWAL_FUNC(ByVal LIF_BALANCE As Double, _
ByVal GOVT_LONG_BOND_RATE As Double, _
ByVal OTHER_REFEREN_RATE As Double, _
Optional ByVal CURRENT_AGE As Double = 60, _
Optional ByVal TEMP_AGE_LAST As Double = 15, _
Optional ByVal AGE_CONVERTION As Double = 80, _
Optional ByVal AGE_BOUND As Double = 90)

'--> Rules: Maximum percentage is 1/PV (PV=Present Value)
'--> of $1.00 per year for 15 (TEMP_AGE_LAST) years after initiating the LIF
'--> (using the MAX Rate), PLUS  1/PV of $1.00
'--> per year thereafter until age 90 (AGE_BOUND) (using the other reference rate
'--> normally 6.0%). This percentage Is applied to the January 1 LIF balance.
'--> AGE = as of January 1. By age 80 (AGE_CONVERTION), LIF must be converted
'--> to a Life Annuity.

'LIFE_INCOME_FUND_CHART:
    'X-AXIS: TEMP_MATRIX(i,7) --> AGE
    
    'Y-AXIS-RIGHT: TEMP_MATRIX(i,8) --> MONTHLY_INCOME
    'Y-AXIS_LEFT: TEMP_MATRIX(i,6) -->LIF_MAX_WITHDRAWAL
    
'CHART_TITLE: =Maximum Monthly Withdrawals if January 1 LIF balance = LIF_BALANCE

Dim i As Long

Dim NROWS As Long
Dim TEMP_AGE As Double

Dim MAX_RATE As Double

Dim TEMP_MATRIX As Variant

Dim FIRST_FACTOR As Double
Dim SECOND_FACTOR As Double

On Error GoTo ERROR_LABEL

NROWS = Int(AGE_CONVERTION - CURRENT_AGE) + 1
MAX_RATE = MAXIMUM_FUNC(GOVT_LONG_BOND_RATE, OTHER_REFEREN_RATE)

FIRST_FACTOR = 1 / (1 + MAX_RATE)
SECOND_FACTOR = 1 / (1 + OTHER_REFEREN_RATE)

ReDim TEMP_MATRIX(0 To NROWS, 1 To 8)

TEMP_MATRIX(0, 1) = "PV"
TEMP_MATRIX(0, 2) = "n"
TEMP_MATRIX(0, 3) = "m"
TEMP_MATRIX(0, 4) = "PV"
TEMP_MATRIX(0, 5) = "1/PV"
TEMP_MATRIX(0, 6) = "LIF_MAX_WITH"
TEMP_MATRIX(0, 7) = "AGE"
TEMP_MATRIX(0, 8) = "MONTHLY_INCOME"

'Suppose your age is A and you have $D in your pocket and you want to
'buy a Life Annuity with this money, the monthly payments to start
'immediately and end when you drop dead ... or maybe the payments are
'transferred to your spouse when you die ... and maybe you want umpteen
'years of payments guaranteed, even if you drop dead immediately.

'The insurance company has to determine how long you'll live (they use
'mortality tables) whether you're male or female (the tables are different),
'how long your spouse will live (if the payments pass to the survivor) and
'what return they can get by investing the $D (the payments will depend upon
'a long term bond rate) and how much they want to keep and how much they want
'to give you, monthly ...

'Alas ... so many factors ... aah, but we can provide a rough (sometimes quite
'rough) estimate like so:

'We invest our $D with an annual return of R% and take annual withdrawals from
'this investment portfolio so as to last until we're half-way to 90 years old ...
'and, if we're now at age A, that means we withdraw for Y = (90 - A)/2 years.

'there 's a neat formula for this and it's:

'$P = $D R/(1 - (1+R)-Y)
                 
For i = 1 To NROWS

    TEMP_AGE = CURRENT_AGE + i - 1

    If TEMP_AGE <= CURRENT_AGE + (TEMP_AGE_LAST - 1) Then
        TEMP_MATRIX(i, 1) = (1 - FIRST_FACTOR ^ (CURRENT_AGE + _
            TEMP_AGE_LAST - TEMP_AGE)) / _
        (1 - FIRST_FACTOR) + FIRST_FACTOR ^ (CURRENT_AGE + _
            TEMP_AGE_LAST - TEMP_AGE) * _
        (1 - SECOND_FACTOR ^ ((AGE_BOUND - TEMP_AGE_LAST) - CURRENT_AGE)) / _
        (1 - SECOND_FACTOR)
    Else: TEMP_MATRIX(i, 1) = (1 - SECOND_FACTOR ^ (AGE_BOUND - TEMP_AGE)) / _
        (1 - SECOND_FACTOR)
    End If

    TEMP_MATRIX(i, 2) = MINIMUM_FUNC(TEMP_AGE_LAST, AGE_BOUND - TEMP_AGE)
    
    TEMP_MATRIX(i, 3) = MAXIMUM_FUNC(AGE_BOUND - TEMP_AGE - TEMP_MATRIX(i, 2), 0)
    
    TEMP_MATRIX(i, 4) = (1 - FIRST_FACTOR ^ TEMP_MATRIX(i, 2)) / (1 - FIRST_FACTOR) + _
    FIRST_FACTOR ^ TEMP_MATRIX(i, 2) * (1 - SECOND_FACTOR ^ TEMP_MATRIX(i, 3)) / _
    (1 - SECOND_FACTOR)
    
    TEMP_MATRIX(i, 5) = 1 / TEMP_MATRIX(i, 4)
    
    If TEMP_AGE <= AGE_CONVERTION Then
       TEMP_MATRIX(i, 6) = 1 / TEMP_MATRIX(i, 1)
    Else: TEMP_MATRIX(i, 6) = ""
    End If

    If i <> 1 Then
        TEMP_MATRIX(i, 7) = TEMP_MATRIX(i - 1, 7) + 1
    Else: TEMP_MATRIX(i, 7) = CURRENT_AGE
    End If
    
    TEMP_MATRIX(i, 8) = TEMP_MATRIX(i, 6) * LIF_BALANCE / 12

Next i

LIF_MAX_WITHDRAWAL_FUNC = TEMP_MATRIX

Exit Function
ERROR_LABEL:
LIF_MAX_WITHDRAWAL_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : PENSION_TABLE_FUNC
'DESCRIPTION   : TOTAL PENSION TABLE
'LIBRARY       : FUNDAMENTAL
'GROUP         : PENSION
'ID            : 004
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 03/05/2009
'************************************************************************************
'************************************************************************************

Function PENSION_TABLE_FUNC(ByVal RETIREMENT_AGE As Double, _
ByVal YEARS_SERVICE As Double, _
ByVal PENSION_SIZE As Double, _
ByVal PENSION_GROWTH As Double, _
ByVal PENSION_SALARY As Double, _
ByVal PENSION_SALARY_GROWTH As Double, _
ByVal DROP_DEAD As Double)

'---------------------------------------------------------------------
'------------INDEX(LOOK_RNG,MATCH(TARGET_VALUE,REF_RNG,0))------------
'------------REF_RNG and TARGET_VALUE must be correlated--------------
'---------------------------------------------------------------------

'-----------------------------------------------------------
'CHART:
'a) PENSION
        'X-AXIS: TEMP_MATRIX(1,2)
        'Y-AXIS: TEMP_MATRIX(4,2) or TEMP_MATRIX(6,2)
    'SENSITIVITY BASED ON YEARS OF SERVICE

'b) Life Expectancy
        'X-AXIS: AGE
        'Y-AXIS: DROP_DEAD
    'SENSITIVITY BASED ON THE AGE AT DIFFERENT DROP DEADS

'----------------------------------------------------------

'PENSION_SIZE: Pension is x% of your salary for each year of service

'PENSION_GROWTH: is x% per year after retirement

'PENSION_SALARY: at RETIREMENT_AGE is $x.xx increasing at
'PENSION_SALARY_GROWTH until retirement.

'DROP_DEAD: Life Expectancy (stick in any convenient numbers)

Dim TEMP_VECTOR As Variant

On Error GoTo ERROR_LABEL

ReDim TEMP_VECTOR(1 To 6, 1 To 2)

TEMP_VECTOR(1, 1) = "RETIREMENT_AGE"
TEMP_VECTOR(1, 2) = RETIREMENT_AGE

TEMP_VECTOR(2, 1) = "YEARS_SERVICE"
TEMP_VECTOR(2, 2) = YEARS_SERVICE

TEMP_VECTOR(3, 1) = "ANNUAL_SALARY"
TEMP_VECTOR(3, 2) = PENSION_SALARY

TEMP_VECTOR(4, 1) = "ANNUAL_PENSION"
TEMP_VECTOR(4, 2) = PENSION_SIZE * TEMP_VECTOR(3, 2) * _
TEMP_VECTOR(2, 2)

TEMP_VECTOR(5, 1) = "LIFE_EXPECTANCY"
TEMP_VECTOR(5, 2) = DROP_DEAD - RETIREMENT_AGE

TEMP_VECTOR(6, 1) = "TOTAL_PENSION"
TEMP_VECTOR(6, 2) = TEMP_VECTOR(4, 2) * ((1 + PENSION_GROWTH) ^ _
TEMP_VECTOR(5, 2) - 1) / PENSION_GROWTH

PENSION_TABLE_FUNC = TEMP_VECTOR

Exit Function
ERROR_LABEL:
PENSION_TABLE_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : PENSION_FUND_MODEL_FUNC
'DESCRIPTION   : PENSION_FUND_MODEL
'LIBRARY       : FUNDAMENTAL
'GROUP         : PENSION
'ID            : 005
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 03/05/2009
'************************************************************************************
'************************************************************************************

Function PENSION_FUND_MODEL_FUNC(ByVal RATE_RETIRED As Double, _
ByVal INFLATION_RATE As Double, _
ByVal FROM_AGE As Double, _
ByVal TO_AGE As Double, _
Optional ByVal DELTA_AGE As Double = 1, _
Optional ByVal MIN_TENOR As Double = 0, _
Optional ByVal MAX_TENOR As Double = 50, _
Optional ByVal DELTA_TENOR As Double = 1, _
Optional ByVal MALE_AGE_BOUND As Double = 81, _
Optional ByVal MALE_TENOR_BOUND As Double = 10.5, _
Optional ByVal FEMALE_AGE_BOUND As Double = 86, _
Optional ByVal FEMALE_TENOR_BOUND As Double = 10)

'At retirement, most individuals face a choice between voluntary annuitization
'and discretionary management of assets with systematic withdrawals for consumption
'purposes. Annuitization — buying a life annuity from an insurance company — assures
'a lifelong consumption stream that can not be outlived, but at the expense of a
'complete loss of liquidity. On the other hand, discretionary management and
'consumption from assets — self annuitization — preserves flexibility but with the
'distinct risk that a constant standard of living will not be maintainable.


Dim i As Long
Dim NROWS As Long

Dim TEMP_AGE As Double
Dim TEMP_MATRIX As Variant
Dim REAL_RETURN As Double

On Error GoTo ERROR_LABEL

REAL_RETURN = (RATE_RETIRED - INFLATION_RATE) / (1 + INFLATION_RATE)

NROWS = (TO_AGE - FROM_AGE) / DELTA_AGE + 1

ReDim TEMP_MATRIX(0 To NROWS, 1 To 5)


'---------------------------------FUNDS_REQUIRED-----------------------------
TEMP_MATRIX(0, 1) = "AGE" '--> x-axis
TEMP_MATRIX(0, 2) = "FUNDS REQUIRED: MALE" '--> y-axis
TEMP_MATRIX(0, 3) = "ANNUITY RATE: MALE" '--> y-axis
TEMP_MATRIX(0, 4) = "FUNDS REQUIRED: FEMALE" '--> y-axis
TEMP_MATRIX(0, 5) = "ANNUITY RATE: FEMALE" '--> y-axis
'----------------------------------------------------------------------------

TEMP_AGE = FROM_AGE

For i = 1 To NROWS

    TEMP_MATRIX(i, 1) = TEMP_AGE
    
    TEMP_MATRIX(i, 2) = PENSION_REQUIRED_FUND_FUNC(REAL_RETURN, TEMP_AGE, MIN_TENOR, _
    MAX_TENOR, DELTA_TENOR, MALE_AGE_BOUND, MALE_TENOR_BOUND, 1)
    
    TEMP_MATRIX(i, 3) = 100 / TEMP_MATRIX(i, 2)
    
    TEMP_MATRIX(i, 4) = PENSION_REQUIRED_FUND_FUNC(REAL_RETURN, TEMP_AGE, MIN_TENOR, _
    MAX_TENOR, DELTA_TENOR, FEMALE_AGE_BOUND, FEMALE_TENOR_BOUND, 1)
    
    TEMP_MATRIX(i, 5) = 100 / TEMP_MATRIX(i, 4)
    
    TEMP_AGE = TEMP_AGE + DELTA_AGE
Next i

PENSION_FUND_MODEL_FUNC = TEMP_MATRIX

Exit Function
ERROR_LABEL:
PENSION_FUND_MODEL_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : PENSION_REQUIRED_FUND_FUNC
'DESCRIPTION   : FUND REQUIRED CALCULATOR
'LIBRARY       : FUNDAMENTAL
'GROUP         : PENSION
'ID            : 006
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 03/05/2009
'************************************************************************************
'************************************************************************************


Function PENSION_REQUIRED_FUND_FUNC(ByVal REAL_RETURN As Double, _
ByVal CURRENT_AGE As Double, _
Optional ByVal MIN_TENOR As Double = 0, _
Optional ByVal MAX_TENOR As Double = 50, _
Optional ByVal DELTA_TENOR As Double = 1, _
Optional ByVal AGE_BOUND As Double = 81, _
Optional ByVal TENOR_BOUND As Double = 10.5, _
Optional ByVal OUTPUT As Integer = 0)

'GENDER: 0 - MALE ; 1 - FEMALE

Dim i As Long
Dim NROWS As Long

Dim TEMP_TENOR As Double
Dim TEMP_SUM As Double
Dim TEMP_MATRIX As Variant

On Error GoTo ERROR_LABEL

TEMP_SUM = 0
NROWS = (MAX_TENOR - MIN_TENOR) / DELTA_TENOR

ReDim TEMP_MATRIX(0 To NROWS, 1 To 3)

TEMP_MATRIX(0, 1) = "TENOR" '--> x-axis
TEMP_MATRIX(0, 2) = "PROB_LIVE" '--> Probability of surviving
'for t years (y-axis)
TEMP_MATRIX(0, 3) = "PENSION_NEEDED"

'Funds Required per $1.00 pension

TEMP_SUM = 0
TEMP_TENOR = MIN_TENOR

For i = 1 To NROWS
    TEMP_MATRIX(i, 1) = TEMP_TENOR
       
    TEMP_MATRIX(i, 2) = PENSION_PROBABILITY_LIVE_FUNC(CURRENT_AGE, TEMP_MATRIX(i, 1), _
    AGE_BOUND, TENOR_BOUND)
    
    TEMP_MATRIX(i, 3) = TEMP_MATRIX(i, 2) / (1 + REAL_RETURN) ^ TEMP_MATRIX(i, 1)
    
    TEMP_SUM = TEMP_SUM + TEMP_MATRIX(i, 3)
    
    TEMP_TENOR = TEMP_TENOR + DELTA_TENOR
Next i

'I want to provide a pension of $1.00 per year to Kristen, every year
'until Kristen drops dead.

'Today, I have $A in my bank account, invested at r%.
'I ask: "How large should A be so that I will have enough to pay Kristen,
'until she dies?"

'If Kristen is still living after, say, 7 years then I'd be paying her that
'$1.00 seven years from now.

'Of course, if the $1.00 were indexed to, say i% inflation, I'd have to
'pay Kristen $1.00 (1+i)7.

'I then need (1+i)7/(1+r)7 today   in order to pay Kristen $1.00 plus seven
'years of inflation ... seven years from now.

'To estimate the required value of A we do this:

'Assume that I have N such Sams and I pay each of them until they drop dead.
'If, after 1 year, N1 are still alive then I'd have to pay $N1(1+i) which means
'I should have N1(1+i)/(1+r) today   in my bank account

'If, after 2 years, N2 are still alive then I'd have to pay $N2(1+i)2 which means
'I should have N2(1+i)2/(1+r)2 today

'If, after 3 years, N3 are still alive then I'd have to pay $N3(1+i)3 which means
'I should have N3(1+i)3/(1+r)3 today etc. etc.

'So the amount I'd need, for N people like Kristen,
'is N1(1+i)/(1+r) + N2(1+i)2/(1+r)2 + N3(1+i)3/(1+r)3 + ...

'So, for just one person (namely Kristen) I divide by N and get:
'A = (N1/N)(1+i)/(1+r) + (N2/N)(1+i)2/(1+r)2 + (N3/N)(1+i)3/(1+r)3 + ...

'We now recognize things like p(k) = Nk/N as the fraction of people
'(like Kristen) who are expected to survive for k years. That'd
'give: A = S p(k) {(1+i)/ (1+r)}k, where p(k) is the probability
'of surviving to year k.

'-----------------------------------Note-----------------------------------
'We can Replace (1+r) / (1+i) with 1+R where R = (r-i)/(1+i) is the
'inflation-adjusted (or "real") return.
'--------------------------------------------------------------------------

Select Case OUTPUT
Case 0
    PENSION_REQUIRED_FUND_FUNC = TEMP_MATRIX
Case Else
    PENSION_REQUIRED_FUND_FUNC = TEMP_SUM
End Select

Exit Function
ERROR_LABEL:
PENSION_REQUIRED_FUND_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : PENSION_PROBABILITY_LIVE_FUNC

'DESCRIPTION   : Probability that an n-year old will die before t-years have
'elapsed [following Milevsky & Robinson]
'http://www.yorku.ca/milevsky/Papers/NAAJ2000A.pdf

'LIBRARY       : FUNDAMENTAL
'GROUP         : PENSION
'ID            : 007
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 03/05/2009
'************************************************************************************
'************************************************************************************


Function PENSION_PROBABILITY_LIVE_FUNC(ByVal CURRENT_AGE As Double, _
ByVal TENOR As Double, _
Optional ByVal AGE_BOUND As Double = 81, _
Optional ByVal TENOR_BOUND As Double = 10.5)

'Probability you will live that many years [tenor]
'MALE_AGE_BOUND = 81
'MALE_TENOR_BOUND = 10.5

'FEMALE_AGE_BOUND = 86
'FEMALE_TENOR_BOUND = 10

'Y-AXIS: PROB
'X-AXIS: TENOR
'Sensitivity based on TENOR per Age

'GENDER: 0 - MALE ; 1 - FEMALE

On Error GoTo ERROR_LABEL

Dim DENSITY As Double

If TENOR = 0 Then: TENOR = 0.000000000001

DENSITY = Exp((CURRENT_AGE - AGE_BOUND) / TENOR_BOUND)

'Useful Explanation: http://www.gummy-stuff.org/annuities-2.htm

PENSION_PROBABILITY_LIVE_FUNC = Exp(DENSITY * (1 - Exp(TENOR / TENOR_BOUND)))

Exit Function
ERROR_LABEL:
PENSION_PROBABILITY_LIVE_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : PENSION_ANNUAL_SAVINGS_FUNC
'DESCRIPTION   : REQUIRED_ANNUAL_SAVINGS_FUNCTION
'LIBRARY       : FUNDAMENTAL
'GROUP         : PENSION
'ID            : 008
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 03/05/2009
'************************************************************************************
'************************************************************************************

Function PENSION_ANNUAL_SAVINGS_FUNC(ByVal PORT_GROWTH As Double, _
ByVal INFLATION_RATE As Double, _
ByVal TENOR_RETIREMENT As Double, _
ByVal SALARY_RETIREMENT As Double, _
ByVal CURRENT_SALARY As Double, _
ByVal WITHDRAWAL As Double)

'PORT_GROWTH: Expected Annual Portfolio Growth
'INFLATION_RATE: Assumed Inflation Rate
'TENOR_RETIREMENT: Years until Retirement
'SALARY_RETIREMENT: Required salary from portfolio _
(at retirement, but in today's dollars)
'CURRENT_SALARY: Your current salary
'WITHDRAWAL: Assumed "safe" Withdrawal Rate

Dim FIRST_FACTOR As Double
Dim SECOND_FACTOR As Double
Dim TEMP_VECTOR As Variant

On Error GoTo ERROR_LABEL

ReDim TEMP_VECTOR(1 To 4, 1 To 2)

FIRST_FACTOR = (1 + INFLATION_RATE) / (1 + PORT_GROWTH)
SECOND_FACTOR = (FIRST_FACTOR - 1) / (FIRST_FACTOR ^ TENOR_RETIREMENT _
- 1) / (1 + PORT_GROWTH) ^ TENOR_RETIREMENT

TEMP_VECTOR(1, 1) = "Salary in " & Format(TENOR_RETIREMENT, "0") & _
" years (from your portfolio):"
TEMP_VECTOR(1, 2) = SALARY_RETIREMENT * (1 + INFLATION_RATE) ^ TENOR_RETIREMENT

TEMP_VECTOR(2, 1) = "Required portfolio (after " & Format(TENOR_RETIREMENT, "0") _
& " years):"

TEMP_VECTOR(2, 2) = TEMP_VECTOR(1, 2) / WITHDRAWAL

TEMP_VECTOR(3, 1) = "Required annual investment (initially, but " & _
"increasing with inflation)"
TEMP_VECTOR(3, 2) = TEMP_VECTOR(2, 2) * SECOND_FACTOR

TEMP_VECTOR(4, 1) = "Required annual savings (initially, as percentage " & _
"of current salary)"
TEMP_VECTOR(4, 2) = TEMP_VECTOR(3, 2) / CURRENT_SALARY

PENSION_ANNUAL_SAVINGS_FUNC = TEMP_VECTOR

'You 'll note that, if you expect to withdraw from your portfolio at
'some "safe" rate when you retire, then you may have to invest a LARGE
'fraction of your current salary !!!


Exit Function
ERROR_LABEL:
PENSION_ANNUAL_SAVINGS_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : PENSION_FUND_GROWTH_FUNC
'DESCRIPTION   : PENSION_FUND_GROWTH_FUNC_FUNCTION
'LIBRARY       : FUNDAMENTAL
'GROUP         : PENSION
'ID            : 009
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 03/05/2009
'************************************************************************************
'************************************************************************************

Function PENSION_FUND_GROWTH_FUNC(ByVal AVG_SALARY As Double, _
ByVal SALARY_GROWTH As Double, _
ByVal TENOR_SERVICE As Double, _
ByVal TENOR_RETIRE As Double, _
ByVal AFTER_RETIRE_TENOR As Double, _
ByVal CONTRIBUTION_PERCENTAGE As Double, _
ByVal CONTRIBUTION_TENOR As Double, _
ByVal SALARY_REPLACEMENT_PERCENTAGE As Double, _
ByVal INFLATION_RATE As Double)

'Avg.Salary: 50,000
'Salary Growth: 2%
'Inflation: 3%
'Years of Service [tenor service]: 20 yrs
'Expect to retire [tenor retire] :10 yrs
'Live for years after retire [after retire tenor]: 20 yrs
'Contribution Amount [contribution percentage] : 16%
'Contribution Years [contribution tenor]: 30
'Salary Replacement: 50%

Dim REAL_RETURN As Double
Dim NOMINAL_RETURN As Double
Dim TEMP_VALUE As Double

On Error GoTo ERROR_LABEL

REAL_RETURN = (SALARY_REPLACEMENT_PERCENTAGE / CONTRIBUTION_PERCENTAGE) _
^ (1 / CONTRIBUTION_TENOR) - 1

NOMINAL_RETURN = REAL_RETURN + INFLATION_RATE

TEMP_VALUE = (AVG_SALARY * SALARY_GROWTH) * (TENOR_SERVICE + TENOR_RETIRE)
TEMP_VALUE = PV_ANNUITY_FUNC(0, TEMP_VALUE, NOMINAL_RETURN, AFTER_RETIRE_TENOR)

PENSION_FUND_GROWTH_FUNC = PV_ANNUITY_FUNC(0, TEMP_VALUE, _
NOMINAL_RETURN, TENOR_RETIRE)

Exit Function
ERROR_LABEL:
PENSION_FUND_GROWTH_FUNC = Err.number
End Function
