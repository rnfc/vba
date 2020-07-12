Attribute VB_Name = "FINAN_DERIV_AMER_LSM_GU_LIBR"


Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.

Private PUB_STRIKE_VAL As Double
Private PUB_OPTION_FLAG As Integer
Private PUB_PARAMETERS As Variant

'************************************************************************************
'************************************************************************************
'FUNCTION      : LSM_GU_AMERICAN_OPTION_REGRESS_NOW_LATER_FUNC
'DESCRIPTION   : Regress now or later : Longstaff Schwartz vs. Glasserman Yu comparision
'LIBRARY       : DERIVATIVES
'GROUP         : AMER_LSM_GU
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/29/2009
'http://www.quantcode.com/modules/mydownloads/viewcat.php?cid=9&min=70&orderby=dateA&show=5
'************************************************************************************
'************************************************************************************

Function LSM_GU_AMERICAN_OPTION_REGRESS_NOW_LATER_FUNC( _
ByVal SPOT As Double, _
ByVal STRIKE As Double, _
ByVal EXPIRATION As Double, _
ByVal RISK_FREE_RATE As Double, _
ByVal VOLATILITY As Double, _
Optional ByVal NTRIALS As Long = 50, _
Optional ByVal nSTEPS As Long = 100, _
Optional ByVal nLOOPS As Long = 500, _
Optional ByRef OPTION_FLAG As Integer = -1)
  
'There are 2 popular regression based methods for evaluating american
'option price using MC simulation:
'1.Longsatff Schwartz
'2.Glasserman Yu

'The difference between the 2 methods lies on the calculation of regression
'variables.
'(1) calculates regression parameters using stock price at current step and
'next step discounted option values

'(2) calculates regression parameters using stock price at next step and
'next step option values

'This algo implements the 2 mehtods within a single function so that code
'can be compared for implementation differences between the 2 models.

'Glasserman 's method was suggested as an improvement over Longstaff's method
'in sense that it is closer to the correct value. In this spredsheet we try to
'verify the result by calculating option prices multiple times and find average.
'eg., in current run displayed on sheet, Longstaff gives a mean of 4.575098651
'while Glasserman's method gives a mean of 4.531829723.

'While the correct option value suggested as per Quantlib run is :
'Option type = Put
'Maturity = May 17th, 1999
'Underlying Price = 36
'STRIKE = 40
'Risk-free interest rate = 6.000000 %
'Dividend yield = 0.000000 %
'Volatility = 20.000000 %

'Method European Bermudan American
'Black-Scholes 3.844308 N/A N/A
'Barone-Adesi/Whaley N/A N/A 4.459628
'Bjerksund/Stensland N/A N/A 4.453064
'Integral 3.844309 N/A N/A
'Finite differences 3.844342 4.360807 4.486118
'Binomial Jarrow-Rudd 3.844132 4.361174 4.486552
'Binomial Cox-Ross-Rubinstein 3.843504 4.360861 4.486415

'Hence the (2) method is closer to value 4.49 and seems to be a better
'alternative. While I have done this trial mutiple times and GU does seem
'to be closer to correct option values, I am not able to understand why
'the variance is more. Any thoughts, or might be a flaw in my implementation?

Dim h As Long
Dim i As Long
Dim j As Long
Dim k As Long
Dim l As Long

Dim DISC_VAL As Double
Dim DELTA_VAL As Double

Dim TEMP_SUM As Double

Dim XTEMP_VAL As Double
Dim YTEMP_VAL As Double

Dim SIMUL_MATRIX As Variant
Dim CONVERG_MATRIX As Variant

Dim CE_MATRIX As Variant 'cash flow from exercise
Dim CCLS_MATRIX As Variant 'cash flow from continuation in Longstaff
Dim CCGU_MATRIX As Variant 'cash flow from continuation in Glasserman

Dim EFLS_MATRIX As Variant 'exercise flags in Longstaff
Dim EFGU_MATRIX As Variant 'exercise flags in Glasserman

Dim XDATA_GU_MATRIX As Variant
Dim XDATA_LS_MATRIX As Variant

Dim YDATA_GU_VECTOR As Variant
Dim YDATA_LS_VECTOR As Variant

Dim OLS_GU_MATRIX As Variant
Dim OLS_LS_MATRIX As Variant

On Error GoTo ERROR_LABEL

ReDim CONVERG_MATRIX(1 To NTRIALS, 1 To 2)
If OPTION_FLAG <> 1 Then: OPTION_FLAG = -1 'Put
DELTA_VAL = EXPIRATION / nSTEPS

For l = 1 To NTRIALS

    SIMUL_MATRIX = AMERICAN_OPTION_PATH_SIMULATION_FUNC(nLOOPS, nSTEPS, SPOT, RISK_FREE_RATE, VOLATILITY, EXPIRATION)
    ReDim CCLS_MATRIX(1 To nLOOPS, 1 To nSTEPS)
    'cash flow from continuation in Longstaff
    ReDim CCGU_MATRIX(1 To nLOOPS, 1 To nSTEPS)
    'cash flow from continuation in Glasserman
    ReDim CE_MATRIX(1 To nLOOPS, 1 To nSTEPS)
    'cash flow from exercise
    ReDim EFLS_MATRIX(1 To nLOOPS, 1 To nSTEPS)
    'exercise flags in Longstaff
    ReDim EFGU_MATRIX(1 To nLOOPS, 1 To nSTEPS)
    'exercise flags in Glasserman

    'Initialize the period at option EXPIRATION
    For i = 1 To nLOOPS
        XTEMP_VAL = SIMUL_MATRIX(i, nSTEPS)
        CE_MATRIX(i, nSTEPS) = MAXIMUM_FUNC(OPTION_FLAG * (XTEMP_VAL - STRIKE), 0)
        CCLS_MATRIX(i, nSTEPS) = CE_MATRIX(i, nSTEPS)
        CCGU_MATRIX(i, nSTEPS) = CE_MATRIX(i, nSTEPS)
        If CE_MATRIX(i, nSTEPS) > 0 Then
            EFLS_MATRIX(i, nSTEPS) = 1
            EFGU_MATRIX(i, nSTEPS) = 1
        Else
            EFLS_MATRIX(i, nSTEPS) = 0
            EFGU_MATRIX(i, nSTEPS) = 0
        End If
    Next i

    DISC_VAL = Exp(-RISK_FREE_RATE * DELTA_VAL)

    For k = nSTEPS - 1 To 2 Step -1
        'Need to Regress discounted continuation value at next time step
        ' to S variables at current time step
        j = 0
        For i = 1 To nLOOPS
            CE_MATRIX(i, k) = MAXIMUM_FUNC(OPTION_FLAG * (SIMUL_MATRIX(i, k) - STRIKE), 0)
            If CE_MATRIX(i, k) > 0 Then
            j = j + 1
            End If
        Next i
        'only the positive payoff points are input for regression
        ReDim XDATA_LS_MATRIX(1 To j, 1 To 3)
        'will become independent variables matrix
        ReDim YDATA_LS_VECTOR(1 To j, 1 To 1)
        'will become observations matrix
        ReDim XDATA_GU_MATRIX(1 To j, 1 To 3)
        'will become independent variables matrix
        ReDim YDATA_GU_VECTOR(1 To j, 1 To 1)
        'will become observations matrix
        
        j = 1
        For i = 1 To nLOOPS
            If CE_MATRIX(i, k) > 0 Then
                XTEMP_VAL = SIMUL_MATRIX(i, k)
                'this step prices have to be considered for regress now
                XDATA_LS_MATRIX(j, 1) = 1
                XDATA_LS_MATRIX(j, 2) = XTEMP_VAL
                XDATA_LS_MATRIX(j, 3) = XTEMP_VAL * XTEMP_VAL
                YDATA_LS_VECTOR(j, 1) = CCLS_MATRIX(i, k + 1) * DISC_VAL
                'this step CC_MATRIX values are used
                XTEMP_VAL = SIMUL_MATRIX(i, k + 1)
                'next step prices have to be considered for regress later
                XDATA_GU_MATRIX(j, 1) = 1
                XDATA_GU_MATRIX(j, 2) = XTEMP_VAL
                XDATA_GU_MATRIX(j, 3) = XTEMP_VAL * XTEMP_VAL
                YDATA_GU_VECTOR(j, 1) = CCGU_MATRIX(i, k + 1)
                'next step CC_MATRIX values are used
                j = j + 1
            End If
        Next i
        
        OLS_LS_MATRIX = REGRESSION_MULT_COEF_FUNC(XDATA_LS_MATRIX, YDATA_LS_VECTOR, False, 0)
        OLS_GU_MATRIX = REGRESSION_MULT_COEF_FUNC(XDATA_GU_MATRIX, YDATA_GU_VECTOR, False, 0)
        
        For i = 1 To nLOOPS
        
            'compare regressed continuation value with immediate
            'exercise value to determine decision for exercise
            
            XTEMP_VAL = SIMUL_MATRIX(i, k)
            YTEMP_VAL = OLS_LS_MATRIX(1, 1) + OLS_LS_MATRIX(2, 1) * XTEMP_VAL + OLS_LS_MATRIX(3, 1) * XTEMP_VAL * XTEMP_VAL
            If CE_MATRIX(i, k) > 0 And CE_MATRIX(i, k) > YTEMP_VAL Then
            'option should be exercised
                EFLS_MATRIX(i, k) = 1
                For h = k + 1 To nSTEPS
                    EFLS_MATRIX(i, h) = 0
                Next h
                CCLS_MATRIX(i, k) = CE_MATRIX(i, k)
            Else
                'option should not be exercised
                'hence take continuation value by discounting
                'from next period continuation value
                CCLS_MATRIX(i, k) = CCLS_MATRIX(i, k + 1) * DISC_VAL
                EFLS_MATRIX(i, k) = 0
            End If
            YTEMP_VAL = OLS_GU_MATRIX(1, 1) + OLS_GU_MATRIX(2, 1) * XTEMP_VAL + OLS_GU_MATRIX(3, 1) * XTEMP_VAL * XTEMP_VAL
            If CE_MATRIX(i, k) > 0 And CE_MATRIX(i, k) > YTEMP_VAL Then
                'option should be exercised
                EFGU_MATRIX(i, k) = 1
                For h = k + 1 To nSTEPS
                    EFGU_MATRIX(i, h) = 0
                Next h
                CCGU_MATRIX(i, k) = CE_MATRIX(i, k)
            Else
                'option should not be exercised
                'hence take continuation value by discounting
                'from next period continuation value
                CCGU_MATRIX(i, k) = CCGU_MATRIX(i, k + 1) * DISC_VAL
                EFGU_MATRIX(i, k) = 0
            End If
        Next i
    Next k
    TEMP_SUM = 0
    For i = 1 To nLOOPS
        For k = 2 To nSTEPS
            If EFLS_MATRIX(i, k) = 1 Then
                XTEMP_VAL = SIMUL_MATRIX(i, k)
                TEMP_SUM = TEMP_SUM + Exp(-RISK_FREE_RATE * (k - 1) * DELTA_VAL) * MAXIMUM_FUNC(OPTION_FLAG * (XTEMP_VAL - STRIKE), 0)
                Exit For
            End If
        Next k
    Next i
    CONVERG_MATRIX(l, 1) = TEMP_SUM / nLOOPS 'MCAmericanPriceLS
    TEMP_SUM = 0
    For i = 1 To nLOOPS
        For k = 2 To nSTEPS
            If EFGU_MATRIX(i, k) = 1 Then
                XTEMP_VAL = SIMUL_MATRIX(i, k)
                TEMP_SUM = TEMP_SUM + Exp(-RISK_FREE_RATE * (k - 1) * DELTA_VAL) * MAXIMUM_FUNC(OPTION_FLAG * (XTEMP_VAL - STRIKE), 0)
                Exit For
            End If
        Next k
    Next i
    CONVERG_MATRIX(l, 2) = TEMP_SUM / nLOOPS 'MCAmericanPriceGU
Next l

'C1: LSM / C2: GU
LSM_GU_AMERICAN_OPTION_REGRESS_NOW_LATER_FUNC = CONVERG_MATRIX

Exit Function
ERROR_LABEL:
LSM_GU_AMERICAN_OPTION_REGRESS_NOW_LATER_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************



'************************************************************************************
'************************************************************************************
'FUNCTION      : LSM_AMERICAN_OPTION_TEST_FUNC
'DESCRIPTION   : This algo implements valuation of a call/put option of american
'exercise type. (Increasing the no of timesteps tend to approach american price)
'LIBRARY       : DERIVATIVES
'GROUP         : AMER_LSM_GU
'ID            : 002
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/29/2009
'************************************************************************************
'************************************************************************************

Function LSM_AMERICAN_OPTION_TEST_FUNC(ByVal SPOT As Double, _
ByVal STRIKE As Double, _
ByVal EXPIRATION As Double, _
ByVal RISK_FREE_RATE As Double, _
ByVal VOLATILITY As Double, _
Optional ByVal nSTEPS As Long = 100, _
Optional ByVal nLOOPS As Long = 10000, _
Optional ByRef OPTION_FLAG As Integer = -1, _
Optional ByVal OUTPUT As Integer = 0)
  
'American Option pricing by Longstaff and schwartz Least Squares
'Longstaff and Schwartz (1992) had devised a method to value American
'options using Monte carlo simulation.

'Assume:
'N : no of simulations
'M : no of time steps
'Algorithm:
'1.Initialize N x M stock prices grid with stock price evolution using
'GBM dynamics.
'2.Create these N x M matrices
'CC : stores continuation values
'CE : store immediate exercise values
'EF : Stores exercise flag
'3.Initialize last column of CE,CE,EF using option payoff
'4.Start Iterating for each time step backwards
'4a.Create a matrix of regression independent variables using basis
'functions 1, X, X^2 and set X=stock price at the node. Note : Logstaff
'recommends to pick only those points which have a positive immediate
'exercise value. rest of the nodes with 0 or negative payoff are ignored
'for regresion input
'4b.Set input dependent/ observations vector to immediate exercise values
'4c.Perform Ordinary Least sqaures to get parameters for coeffcients of
'basis functions. It is basically a regression of stock variables at the
'current time step to cotinuation values at the next timestep.
'4d.Now compute back continuation values using the parameter values obtained
'from 4c.
'4e.Compare continuation values obtained from regression vis-a-vis the
'immediate exercise values
'4f.If immediate exercise value is more than regressed continuation value,
'choose to exercise and Populate EF matrix flag at the node.
'4g.Populate continuation value : If exercise flag at node is set then set
'it to immediate exercise value else set it to discounted value of
'continuation value at next timestep
'4h.Repeat 4a until simualtions are done
'5.Observe the EF matrix and think that actual exercise decision was made when
'EF flag=1. Accordingly calculate expected payoff by averaging. This gives
'the MC option price
'NOTE: I have tried comparing results to following run from Quantlib:
'Option type = Put
'Maturity = May 17th, 1999
'Underlying Price = 36
'STRIKE = 40
'Risk-free interest RISK_FREE_RATE = 6.000000 %
'Dividend yield = 0.000000 %
'Volatility = 20.000000 %'

'Method European Bermudan American
'Black-Scholes 3.844308 N/A N/A
'Barone-Adesi/Whaley N/A N/A 4.459628
'Bjerksund/Stensland N/A N/A 4.453064
'Integral 3.844309 N/A N/A
'Finite differences 3.844342 4.360807 4.486118
'Binomial Jarrow-Rudd 3.844132 4.361174 4.486552
'Binomial Cox-Ross-Rubinstein 3.843504 4.360861 4.486415

'Using same input parameters and timesteps=100 and simualtions=10000,
'I am getting price around 4.50

'To get an idea of the underlying process, there is an additional option
'of dumping the matrices of stock prices,continuaiton values,exercise flags
'and immediate exercise values at each node. To enable this option,
'tiemsteps and simulations should be less than 100

Dim h As Long
Dim i As Long
Dim j As Long
Dim k As Long

Dim DISC_VAL As Double
Dim DELTA_VAL As Double

Dim TEMP_SUM As Double

Dim XTEMP_VAL As Double
Dim YTEMP_VAL As Double

Dim CE_MATRIX As Variant 'cash flow from exercise
Dim CC_MATRIX As Variant 'cash flow from continuation
Dim EF_MATRIX As Variant 'exercise flags

Dim OLS_MATRIX As Variant
Dim XDATA_MATRIX As Variant
Dim YDATA_VECTOR As Variant

Dim TEMP_GROUP As Variant
Dim SIMUL_MATRIX As Variant

On Error GoTo ERROR_LABEL

If OPTION_FLAG <> 1 Then: OPTION_FLAG = -1 'Put

DELTA_VAL = EXPIRATION / nSTEPS
SIMUL_MATRIX = AMERICAN_OPTION_PATH_SIMULATION_FUNC(nLOOPS, nSTEPS, SPOT, RISK_FREE_RATE, VOLATILITY, EXPIRATION)

ReDim CC_MATRIX(1 To nLOOPS, 1 To nSTEPS) 'cash flow from continuation
ReDim CE_MATRIX(1 To nLOOPS, 1 To nSTEPS) 'cash flow from exercise
ReDim EF_MATRIX(1 To nLOOPS, 1 To nSTEPS) 'exercise flags

'Initialize the period at option EXPIRATION
For i = 1 To nLOOPS
    XTEMP_VAL = SIMUL_MATRIX(i, nSTEPS)
    CE_MATRIX(i, nSTEPS) = MAXIMUM_FUNC(OPTION_FLAG * (XTEMP_VAL - STRIKE), 0)
    CC_MATRIX(i, nSTEPS) = CE_MATRIX(i, nSTEPS)
    If CE_MATRIX(i, nSTEPS) > 0 Then
        EF_MATRIX(i, nSTEPS) = 1
    Else
        EF_MATRIX(i, nSTEPS) = 0
    End If
Next i

DISC_VAL = Exp(-RISK_FREE_RATE * DELTA_VAL)
For k = nSTEPS - 1 To 2 Step -1
    'Need to Regress discounted continuation value at next time step
    ' to S variables at current time step
    j = 0
    For i = 1 To nLOOPS
        CE_MATRIX(i, k) = MAXIMUM_FUNC(OPTION_FLAG * (SIMUL_MATRIX(i, k) - STRIKE), 0)
        If CE_MATRIX(i, k) > 0 Then
            j = j + 1
        End If
    Next i
    'only the positive payoff points are input for regression
    ReDim XDATA_MATRIX(1 To j, 1 To 3)
    'will become independent variables matrix
    ReDim YDATA_VECTOR(1 To j, 1 To 1)
    'will become observations matrix
    j = 1
    For i = 1 To nLOOPS
        If CE_MATRIX(i, k) > 0 Then
            XTEMP_VAL = SIMUL_MATRIX(i, k)
            XDATA_MATRIX(j, 1) = 1
            XDATA_MATRIX(j, 2) = XTEMP_VAL
            XDATA_MATRIX(j, 3) = XTEMP_VAL * XTEMP_VAL
            YDATA_VECTOR(j, 1) = CC_MATRIX(i, k + 1) * DISC_VAL
            j = j + 1
        End If
    Next i
    
    OLS_MATRIX = REGRESSION_MULT_COEF_FUNC(XDATA_MATRIX, YDATA_VECTOR, False, 0)
    For i = 1 To nLOOPS
        'compare regressed continuation value with immediate
        'exercise value to determine decision for exercise
        XTEMP_VAL = SIMUL_MATRIX(i, k)
        YTEMP_VAL = OLS_MATRIX(1, 1) + OLS_MATRIX(2, 1) * XTEMP_VAL + OLS_MATRIX(3, 1) * XTEMP_VAL * XTEMP_VAL
        If CE_MATRIX(i, k) > 0 And CE_MATRIX(i, k) > YTEMP_VAL Then
            'option should be exercised
            EF_MATRIX(i, k) = 1
            For h = k + 1 To nSTEPS
                EF_MATRIX(i, h) = 0
            Next h
            CC_MATRIX(i, k) = CE_MATRIX(i, k)
        Else
            'option should not be exercised
            'hence take continuation value by discounting
            'from next period continuation value
            CC_MATRIX(i, k) = CC_MATRIX(i, k + 1) * DISC_VAL
            EF_MATRIX(i, k) = 0
        End If
    Next i
Next k

TEMP_SUM = 0

For i = 1 To nLOOPS
    For k = 2 To nSTEPS
        If EF_MATRIX(i, k) = 1 Then
            XTEMP_VAL = SIMUL_MATRIX(i, k)
            TEMP_SUM = TEMP_SUM + Exp(-RISK_FREE_RATE * (k - 1) * DELTA_VAL) * MAXIMUM_FUNC(OPTION_FLAG * (XTEMP_VAL - STRIKE), 0)
            Exit For
        End If
    Next k
Next i

Select Case OUTPUT
Case 0 'MCAmericanPrice
    LSM_AMERICAN_OPTION_TEST_FUNC = TEMP_SUM / nLOOPS
Case 1
    LSM_AMERICAN_OPTION_TEST_FUNC = SIMUL_MATRIX
Case 2
    LSM_AMERICAN_OPTION_TEST_FUNC = EF_MATRIX
Case 3
    LSM_AMERICAN_OPTION_TEST_FUNC = CE_MATRIX
Case 4
    LSM_AMERICAN_OPTION_TEST_FUNC = CC_MATRIX
Case Else
    ReDim TEMP_GROUP(1 To 5)
    TEMP_GROUP(1) = TEMP_SUM / nLOOPS
    TEMP_GROUP(2) = SIMUL_MATRIX
    TEMP_GROUP(3) = EF_MATRIX
    TEMP_GROUP(4) = CE_MATRIX
    TEMP_GROUP(5) = CC_MATRIX
    LSM_AMERICAN_OPTION_TEST_FUNC = TEMP_GROUP
End Select

Exit Function
ERROR_LABEL:
LSM_AMERICAN_OPTION_TEST_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************



'************************************************************************************
'************************************************************************************
'FUNCTION      : LSM_EXERCISE_BOUNDARY_FUNC
'DESCRIPTION   : Optimal exercise boundary in Longstaff & Schwartz model is obtained by
'solving the basis function equation with the immediate exercise value.
'Exercise decision is made when regressed continuation value is less than
'immediate exercise value.

'LIBRARY       : DERIVATIVES
'GROUP         : AMER_LSM_GU
'ID            : 003
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/29/2009
'************************************************************************************
'************************************************************************************

Function LSM_EXERCISE_BOUNDARY_FUNC(ByVal SPOT As Double, _
ByVal STRIKE As Double, _
ByVal EXPIRATION As Double, _
ByVal RISK_FREE_RATE As Double, _
ByVal VOLATILITY As Double, _
Optional ByVal nSTEPS As Long = 20, _
Optional ByVal nLOOPS As Long = 20000, _
Optional ByRef OPTION_FLAG As Integer = -1, _
Optional ByVal OUTPUT As Integer = 1)

'Regressed Value: p (1) + p(2) * x + p(3) * x ^ 2
'Immediate exercise value : exerciseflag*(x-Strike)
'where x is stock price
'Hence the boundary can be obtained by solving the equation
'p (1) + p(2) * x + p(3) * x ^ 2 = exerciseflag * (x - STRIKE)
'Due to quadratic nature of this equation 2 roots are obtained. Hence
'continuation value is less then immediate exercise value whenever stock
'price lies between root1 and root2. This is thus the exercise region
'suggested by LSM model.
'This function plots optimal exercise frontier for an American Put after t=0 to
'maturity. Note that if call option is selected it gives unreliable results.
'I am thinking this is because it is never optimal to exercise a call option
'before expiry unless it encounters a dividend period.

Dim h As Long
Dim i As Long
Dim j As Long
Dim k As Long

Dim DISC_VAL As Double
Dim DELTA_VAL As Double
Dim TEMP_SUM As Double

Dim ATEMP_VAL As Double
Dim BTEMP_VAL As Double
Dim CTEMP_VAL As Double
Dim DTEMP_VAL As Double

Dim XTEMP_VAL As Double
Dim YTEMP_VAL As Double

Dim CE_MATRIX As Variant 'cash flow from exercise
Dim CC_MATRIX As Variant 'cash flow from continuation
Dim EF_MATRIX As Variant 'exercise flags

Dim OLS_MATRIX As Variant
Dim XDATA_MATRIX As Variant
Dim YDATA_VECTOR As Variant
Dim TEMP_GROUP As Variant

Dim SIMUL_MATRIX As Variant
Dim PARAM_MATRIX As Variant
Dim ROOTS_MATRIX As Variant

On Error GoTo ERROR_LABEL

If OPTION_FLAG <> 1 Then: OPTION_FLAG = -1 'Put

DELTA_VAL = EXPIRATION / nSTEPS
SIMUL_MATRIX = AMERICAN_OPTION_PATH_SIMULATION_FUNC(nLOOPS, nSTEPS, SPOT, RISK_FREE_RATE, VOLATILITY, EXPIRATION)
ReDim CC_MATRIX(1 To nLOOPS, 1 To nSTEPS) 'cash flow from continuation
ReDim CE_MATRIX(1 To nLOOPS, 1 To nSTEPS) 'cash flow from exercise
ReDim EF_MATRIX(1 To nLOOPS, 1 To nSTEPS) 'exercise flags

ReDim PARAM_MATRIX(1 To 3, 1 To nSTEPS) 'for debugging estimation parameters
ReDim ROOTS_MATRIX(1 To nSTEPS, 1 To 3)

'Initialize the period at option EXPIRATION
For i = 1 To nLOOPS
    XTEMP_VAL = SIMUL_MATRIX(i, nSTEPS)
    CE_MATRIX(i, nSTEPS) = MAXIMUM_FUNC(OPTION_FLAG * (XTEMP_VAL - STRIKE), 0)
    CC_MATRIX(i, nSTEPS) = CE_MATRIX(i, nSTEPS)
    If CE_MATRIX(i, nSTEPS) > 0 Then
      EF_MATRIX(i, nSTEPS) = 1
    Else
      EF_MATRIX(i, nSTEPS) = 0
    End If
Next i

ROOTS_MATRIX(nSTEPS, 1) = EXPIRATION
If OPTION_FLAG = -1 Then
    ROOTS_MATRIX(nSTEPS, 2) = STRIKE
    ROOTS_MATRIX(nSTEPS, 3) = 0
Else
    ROOTS_MATRIX(nSTEPS, 2) = STRIKE * 2
    ROOTS_MATRIX(nSTEPS, 3) = STRIKE
End If
DISC_VAL = Exp(-RISK_FREE_RATE * DELTA_VAL)

For k = nSTEPS - 1 To 2 Step -1
  'Need to Regress discounted continuation value at next time step
  ' to S variables at current time step
    j = 0
    For i = 1 To nLOOPS
        CE_MATRIX(i, k) = MAXIMUM_FUNC(OPTION_FLAG * (SIMUL_MATRIX(i, k) - STRIKE), 0)
        If CE_MATRIX(i, k) > 0 Then
            j = j + 1
        End If
    Next i
    'only the positive payoff points are input for regression
    ReDim XDATA_MATRIX(1 To j, 1 To 3)
    'will become independent variables matrix
    ReDim YDATA_VECTOR(1 To j, 1 To 1)
    'will become observations matrix
    j = 1
    For i = 1 To nLOOPS
        If CE_MATRIX(i, k) > 0 Then
            XTEMP_VAL = SIMUL_MATRIX(i, k)
            XDATA_MATRIX(j, 1) = 1
            XDATA_MATRIX(j, 2) = XTEMP_VAL
            XDATA_MATRIX(j, 3) = XTEMP_VAL * XTEMP_VAL
            YDATA_VECTOR(j, 1) = CC_MATRIX(i, k + 1) * DISC_VAL
            j = j + 1
        End If
    Next i
    
    OLS_MATRIX = REGRESSION_MULT_COEF_FUNC(XDATA_MATRIX, YDATA_VECTOR, False, 0)
    For i = 1 To 3
        PARAM_MATRIX(i, k) = OLS_MATRIX(i, 1)
    Next i
    For i = 1 To nLOOPS
        'compare regressed continuation value with immediate
        'exercise value to determine decision for exercise
        XTEMP_VAL = SIMUL_MATRIX(i, k)
        YTEMP_VAL = OLS_MATRIX(1, 1) + OLS_MATRIX(2, 1) * XTEMP_VAL + _
        OLS_MATRIX(3, 1) * XTEMP_VAL * XTEMP_VAL
        If CE_MATRIX(i, k) > 0 And CE_MATRIX(i, k) > YTEMP_VAL Then
            'option should be exercised
            EF_MATRIX(i, k) = 1
            For h = k + 1 To nSTEPS
                EF_MATRIX(i, h) = 0
            Next h
            CC_MATRIX(i, k) = CE_MATRIX(i, k)
        Else
            'option should not be exercised
            'hence take continuation value by discounting
            'from next period continuation value
            CC_MATRIX(i, k) = CC_MATRIX(i, k + 1) * DISC_VAL
            EF_MATRIX(i, k) = 0
        End If
    Next i
    'need to find roots of the equation ax^2+bx+c=0
    'the 2 roots are -b+-(BTEMP_VAL^2-4ac)^0.5/2a
    ATEMP_VAL = OLS_MATRIX(3, 1)
    BTEMP_VAL = OLS_MATRIX(2, 1) - OPTION_FLAG
    CTEMP_VAL = OLS_MATRIX(1, 1) + OPTION_FLAG * STRIKE
    DTEMP_VAL = (BTEMP_VAL ^ 2 - 4 * ATEMP_VAL * CTEMP_VAL)
    If DTEMP_VAL > 0 Then
        DTEMP_VAL = DTEMP_VAL ^ 0.5
        ROOTS_MATRIX(k, 1) = k * DELTA_VAL
        ROOTS_MATRIX(k, 2) = (-BTEMP_VAL + DTEMP_VAL) / (2 * ATEMP_VAL)
        ROOTS_MATRIX(k, 3) = (-BTEMP_VAL - DTEMP_VAL) / (2 * ATEMP_VAL)
    Else
        ROOTS_MATRIX(k, 1) = k * DELTA_VAL
        ROOTS_MATRIX(k, 2) = 0
        ROOTS_MATRIX(k, 3) = 0
    End If

Next k

TEMP_SUM = 0
For i = 1 To nLOOPS
    For k = 2 To nSTEPS
        If EF_MATRIX(i, k) = 1 Then
            XTEMP_VAL = SIMUL_MATRIX(i, k)
            TEMP_SUM = TEMP_SUM + Exp(-RISK_FREE_RATE * (k - 1) * DELTA_VAL) * MAXIMUM_FUNC(OPTION_FLAG * (XTEMP_VAL - STRIKE), 0)
            Exit For
        End If
    Next k
Next i


'lets evaluate optimal exercise frontier
'For k = nSTEPS - 1 To 2 Step -1
'Next k

Select Case OUTPUT
Case 0 'MCAmericanPrice
    LSM_EXERCISE_BOUNDARY_FUNC = TEMP_SUM / nLOOPS
Case 1 'Exercise Boundary Region; Time/Root1/Root2
    LSM_EXERCISE_BOUNDARY_FUNC = ROOTS_MATRIX
Case 2
    LSM_EXERCISE_BOUNDARY_FUNC = SIMUL_MATRIX
Case 3
    LSM_EXERCISE_BOUNDARY_FUNC = EF_MATRIX
Case 4
    LSM_EXERCISE_BOUNDARY_FUNC = CE_MATRIX
Case 5
    LSM_EXERCISE_BOUNDARY_FUNC = CC_MATRIX
Case 6
    LSM_EXERCISE_BOUNDARY_FUNC = PARAM_MATRIX
Case Else
    ReDim TEMP_GROUP(1 To 7)
    TEMP_GROUP(1) = TEMP_SUM / nLOOPS
    TEMP_GROUP(2) = ROOTS_MATRIX
    TEMP_GROUP(3) = SIMUL_MATRIX
    TEMP_GROUP(4) = EF_MATRIX
    TEMP_GROUP(5) = CE_MATRIX
    TEMP_GROUP(6) = CC_MATRIX
    TEMP_GROUP(7) = PARAM_MATRIX
    
    LSM_EXERCISE_BOUNDARY_FUNC = TEMP_GROUP
End Select

Exit Function
ERROR_LABEL:
LSM_EXERCISE_BOUNDARY_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : GU_AMERICAN_OPTION_REGRESS_LATER_FUNC

'DESCRIPTION   : Regress later is a monte carlo method suggested by Glasserman
'and Yu. It is different from Longstaff and Schwartz method that the regession
'takes place one time step ahead. To do this the authors suggest use
'basis functions in regression which are martingales.

'LIBRARY       : DERIVATIVES
'GROUP         : AMER_LSM_GU
'ID            : 004
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/29/2009
'************************************************************************************
'************************************************************************************

Function GU_AMERICAN_OPTION_REGRESS_LATER_FUNC( _
ByVal SPOT As Double, _
ByVal STRIKE As Double, _
ByVal EXPIRATION As Double, _
ByVal RISK_FREE_RATE As Double, _
ByVal VOLATILITY As Double, _
Optional ByVal nSTEPS As Long = 100, _
Optional ByVal nLOOPS As Long = 10000, _
Optional ByRef OPTION_FLAG As Integer = -1, _
Optional ByVal OUTPUT As Integer = 0)


'Assume:
'N : no of simulations
'M : no of time steps
'Algorithm:
'1.Initialize N x M stock prices grid with stock price evolution using
'GBM dynamics.
'2.Create these N x M matrices
'CC : stores continuation values
'CE : store immediate exercise values
'EF : Stores exercise flag
'3.Initialize last column of CE,CE,EF using option payoff
'4.Start Iterating for each time step backwards
'4a.Create a matrix of regression independent variables using basis functions
'1, X, X^2 and set X=stock price at the node stepof the next time
'Note : pick only those points which have a positive immediate exercise value.
'rest of the nodes with 0 or negative payoff are ignored for regresion input
'4b.Set input dependent/ observations vector to immediate exercise values of
'the next time step
'4c.Perform Ordinary Least sqaures to get parameters for coeffcients of basis
'functions. It is basically a regression of stock variables at the next time
'step to option values at the next timestep. Since both the independent and
'dependent variables are taken at same timestep, the values for comparision
'using these regression parameters do not need to get discounted.
'4d.Now compute continuation values at this node using the parameter values
'obtained from 4c.Note that Longstaff method comapres discounted contunation
'value from next step to current immediate exercise value.
'While Glasserman compares regressed values whose parameters were basically
'computed using stock variables and option values at the next step.
'4e.Compare continuation values obtained from regression in 4d vis-a-vis the
'immediate exercise values
'4f.If immediate exercise value is more than regressed continuation value,
'choose to exercise and Populate EF matrix flag at the node.
'4g.Populate continuation value : If exercise flag at node is set then set it
'to immediate exercise value else set it to discounted value of continuation
'value at next timestep
'4h.Repeat 4a until simualtions are done
'5.Observe the EF matrix and think that actual exercise decision was made when
'EF flag=1. Accordingly calculate expected payoff by averaging. This gives
'the MC option price
'NOTE: I have tried comparing results to following run from Quantlib:
'Option type = Put
'Maturity = May 17th, 1999
'Underlying Price = 36
'STRIKE = 40
'Risk-free interest rate = 6.000000 %
'Dividend yield = 0.000000 %
'Volatility = 20.000000 %

'Method European Bermudan American
'Black-Scholes 3.844308 N/A N/A
'Barone-Adesi/Whaley N/A N/A 4.459628
'Bjerksund/Stensland N/A N/A 4.453064
'Integral 3.844309 N/A N/A
'Finite differences 3.844342 4.360807 4.486118
'Binomial Jarrow-Rudd 3.844132 4.361174 4.486552
'Binomial Cox-Ross-Rubinstein 3.843504 4.360861 4.486415

'Using same input parameters in excel file and timesteps=100 and
'simualtions=10000, I am getting price around 4.47

'To get an idea of the underlying process, there is an additional option
'of dumping the matrices of stock prices,continuaiton values,exercise flags
'and immediate exercise values at each node. To enable this option, tiemsteps
'and simulations should be less than 100

Dim h As Long
Dim i As Long
Dim j As Long
Dim k As Long

Dim DISC_VAL As Double
Dim DELTA_VAL As Double

Dim TEMP_SUM As Double

Dim XTEMP_VAL As Double
Dim YTEMP_VAL As Double

Dim CE_MATRIX As Variant 'cash flow from exercise
Dim CC_MATRIX As Variant 'cash flow from continuation
Dim EF_MATRIX As Variant 'exercise flags

Dim OLS_MATRIX As Variant
Dim XDATA_MATRIX As Variant
Dim YDATA_VECTOR As Variant

Dim TEMP_GROUP As Variant
Dim SIMUL_MATRIX As Variant

On Error GoTo ERROR_LABEL

If OPTION_FLAG <> 1 Then: OPTION_FLAG = -1 'Put

DELTA_VAL = EXPIRATION / nSTEPS
SIMUL_MATRIX = AMERICAN_OPTION_PATH_SIMULATION_FUNC(nLOOPS, nSTEPS, SPOT, RISK_FREE_RATE, VOLATILITY, EXPIRATION)

ReDim CC_MATRIX(1 To nLOOPS, 1 To nSTEPS)
'cash flow from continuation
ReDim CE_MATRIX(1 To nLOOPS, 1 To nSTEPS)
'cash flow from exercise
ReDim EF_MATRIX(1 To nLOOPS, 1 To nSTEPS)
'exercise flags

'Initialize the period at option EXPIRATION
For i = 1 To nLOOPS
    XTEMP_VAL = SIMUL_MATRIX(i, nSTEPS)
    CE_MATRIX(i, nSTEPS) = MAXIMUM_FUNC(OPTION_FLAG * (XTEMP_VAL - STRIKE), 0)
    CC_MATRIX(i, nSTEPS) = CE_MATRIX(i, nSTEPS)
    If CE_MATRIX(i, nSTEPS) > 0 Then
        EF_MATRIX(i, nSTEPS) = 1
    Else
        EF_MATRIX(i, nSTEPS) = 0
    End If
Next i

DISC_VAL = Exp(-RISK_FREE_RATE * DELTA_VAL)

For k = nSTEPS - 1 To 2 Step -1
'Need to Regress discounted continuation value at next time step
'to S variables at current time step
    j = 0
    For i = 1 To nLOOPS
        CE_MATRIX(i, k) = MAXIMUM_FUNC(OPTION_FLAG * (SIMUL_MATRIX(i, k) - STRIKE), 0)
        If CE_MATRIX(i, k) > 0 Then
            j = j + 1
        End If
    Next i
    'only the positive payoff points are input for regression
    ReDim XDATA_MATRIX(1 To j, 1 To 3) 'will become independent variables matrix
    ReDim YDATA_VECTOR(1 To j, 1 To 1) 'will become observations matrix
    j = 1
    For i = 1 To nLOOPS
        If CE_MATRIX(i, k) > 0 Then
            XTEMP_VAL = SIMUL_MATRIX(i, k + 1)
            'next step prices have to be considered for regress later
            XDATA_MATRIX(j, 1) = 1
            XDATA_MATRIX(j, 2) = XTEMP_VAL
            XDATA_MATRIX(j, 3) = XTEMP_VAL * XTEMP_VAL
            YDATA_VECTOR(j, 1) = CC_MATRIX(i, k + 1)
            'next step CC_MATRIX values are used
            j = j + 1
        End If
    Next i
    
    OLS_MATRIX = REGRESSION_MULT_COEF_FUNC(XDATA_MATRIX, YDATA_VECTOR, False, 0)
    For i = 1 To nLOOPS
        'compare regressed continuation value with immediate
        'exercise value to determine decision for exercise
        XTEMP_VAL = SIMUL_MATRIX(i, k)
        YTEMP_VAL = OLS_MATRIX(1, 1) + OLS_MATRIX(2, 1) * XTEMP_VAL + _
        OLS_MATRIX(3, 1) * XTEMP_VAL * XTEMP_VAL
        If CE_MATRIX(i, k) > 0 And CE_MATRIX(i, k) > YTEMP_VAL Then
            'option should be exercised
            EF_MATRIX(i, k) = 1
            For h = k + 1 To nSTEPS
                EF_MATRIX(i, h) = 0
            Next h
            CC_MATRIX(i, k) = CE_MATRIX(i, k)
        Else
            'option should not be exercised
            'hence take continuation value by discounting
            'from next period continuation value
            CC_MATRIX(i, k) = CC_MATRIX(i, k + 1) * DISC_VAL
            EF_MATRIX(i, k) = 0
        End If
    Next i
Next k

TEMP_SUM = 0
For i = 1 To nLOOPS
    For k = 2 To nSTEPS
        If EF_MATRIX(i, k) = 1 Then
            XTEMP_VAL = SIMUL_MATRIX(i, k)
            TEMP_SUM = TEMP_SUM + Exp(-RISK_FREE_RATE * (k - 1) * DELTA_VAL) * MAXIMUM_FUNC(OPTION_FLAG * (XTEMP_VAL - STRIKE), 0)
            Exit For
        End If
    Next k
Next i

Select Case OUTPUT
Case 0 'MCAmericanPrice
    GU_AMERICAN_OPTION_REGRESS_LATER_FUNC = TEMP_SUM / nLOOPS
Case 1
    GU_AMERICAN_OPTION_REGRESS_LATER_FUNC = SIMUL_MATRIX
Case 2
    GU_AMERICAN_OPTION_REGRESS_LATER_FUNC = EF_MATRIX
Case 3
    GU_AMERICAN_OPTION_REGRESS_LATER_FUNC = CE_MATRIX
Case 4
    GU_AMERICAN_OPTION_REGRESS_LATER_FUNC = CC_MATRIX
Case Else
    ReDim TEMP_GROUP(1 To 5)
    TEMP_GROUP(1) = TEMP_SUM / nLOOPS
    TEMP_GROUP(2) = SIMUL_MATRIX
    TEMP_GROUP(3) = EF_MATRIX
    TEMP_GROUP(4) = CE_MATRIX
    TEMP_GROUP(5) = CC_MATRIX
    GU_AMERICAN_OPTION_REGRESS_LATER_FUNC = TEMP_GROUP
End Select

Exit Function
ERROR_LABEL:
GU_AMERICAN_OPTION_REGRESS_LATER_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : LSM_AMERICAN_OPTION_POLICY_FUNC
'DESCRIPTION   : Exercise Policy in Longstaff Schwartz Model
'LIBRARY       : DERIVATIVES
'GROUP         : AMER_LSM_GU
'ID            : 005
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/29/2009
'************************************************************************************
'************************************************************************************

Function LSM_AMERICAN_OPTION_POLICY_FUNC(ByVal SPOT As Double, _
ByVal STRIKE As Double, _
ByVal EXPIRATION As Double, _
ByVal RISK_FREE_RATE As Double, _
ByVal VOLATILITY As Double, _
ByVal TIME_STEP As Long, _
ByVal MIN_STOCK_PRICE As Double, _
ByVal MAX_STOCK_PRICE As Double, _
ByVal STOCK_PRICE_POINTS As Double, _
Optional ByRef PARAM_RNG As Variant, _
Optional ByVal nSTEPS As Long = 10, _
Optional ByVal nLOOPS As Long = 10000, _
Optional ByRef OPTION_FLAG As Integer = -1, _
Optional ByVal OUTPUT As Integer = 0)

Dim h As Long
Dim i As Long
Dim j As Long
Dim k As Long

Dim DISC_VAL As Double
Dim DELTA_VAL As Double
Dim STEP_SIZE As Double

Dim TEMP_SUM As Double

Dim XTEMP_VAL As Double
Dim YTEMP_VAL As Double

Dim CE_MATRIX As Variant 'cash flow from exercise
Dim CC_MATRIX As Variant 'cash flow from continuation
Dim EF_MATRIX As Variant 'exercise flags

Dim OLS_MATRIX As Variant
Dim XDATA_MATRIX As Variant
Dim YDATA_VECTOR As Variant

Dim TEMP_GROUP As Variant
Dim TEMP_MATRIX As Variant

Dim SIMUL_MATRIX As Variant
Dim PARAM_MATRIX As Variant

Dim BRENT_GUESS_VAL As Double
Dim BRENT_LOWER_BOUND As Double
Dim BRENT_UPPER_BOUND As Double
Dim BRENT_COUNTER As Long
Dim BRENT_CONVERG_VAL As Integer
Dim BRENT_nLOOPS As Long
Dim BRENT_epsilon As Double

On Error GoTo ERROR_LABEL

'-------------------------------------------------------------------------------
'TESTING PARAMETERS
'-------------------------------------------------------------------------------

BRENT_GUESS_VAL = 36
BRENT_LOWER_BOUND = 35
BRENT_UPPER_BOUND = 40
BRENT_nLOOPS = 600
BRENT_epsilon = 0.001

If IsArray(PARAM_RNG) = False Then
    ReDim PARAM_RNG(1 To 3, 1 To 1)
    PARAM_RNG(1, 1) = 46.60673456
    PARAM_RNG(2, 1) = -1.469053663
    PARAM_RNG(3, 1) = 0.007899828
End If

'-------------------------------------------------------------------------------
'-------------------------------------------------------------------------------

If TIME_STEP >= nSTEPS Or TIME_STEP = 0 Then
    GoTo ERROR_LABEL ' time step being analyzed is out of range"
End If
If OPTION_FLAG <> 1 Then: OPTION_FLAG = -1 'Put

Randomize
DELTA_VAL = EXPIRATION / nSTEPS
SIMUL_MATRIX = AMERICAN_OPTION_PATH_SIMULATION_FUNC(nLOOPS, nSTEPS, SPOT, RISK_FREE_RATE, VOLATILITY, EXPIRATION)

ReDim CC_MATRIX(1 To nLOOPS, 1 To nSTEPS) 'cash flow from continuation
ReDim CE_MATRIX(1 To nLOOPS, 1 To nSTEPS) 'cash flow from exercise
ReDim EF_MATRIX(1 To nLOOPS, 1 To nSTEPS) 'exercise flags
ReDim PARAM_MATRIX(1 To 3, 1 To nSTEPS) 'for debugging estimation parameters

'Initialize the period at option EXPIRATION
For i = 1 To nLOOPS
    XTEMP_VAL = SIMUL_MATRIX(i, nSTEPS)
    CE_MATRIX(i, nSTEPS) = MAXIMUM_FUNC(OPTION_FLAG * (XTEMP_VAL - STRIKE), 0)
    CC_MATRIX(i, nSTEPS) = CE_MATRIX(i, nSTEPS)
    If CE_MATRIX(i, nSTEPS) > 0 Then
        EF_MATRIX(i, nSTEPS) = 1
    Else
        EF_MATRIX(i, nSTEPS) = 0
    End If
Next i

DISC_VAL = Exp(-RISK_FREE_RATE * DELTA_VAL)

For k = nSTEPS - 1 To 2 Step -1
'Need to Regress discounted continuation value at next time step
'to S variables at current time step
    j = 0
    For i = 1 To nLOOPS
        CE_MATRIX(i, k) = MAXIMUM_FUNC(OPTION_FLAG * (SIMUL_MATRIX(i, k) - STRIKE), 0)
        If CE_MATRIX(i, k) > 0 Then
            j = j + 1
        End If
    Next i
    'only the positive payoff points are input for regression
    ReDim XDATA_MATRIX(1 To j, 1 To 3) 'will become independent variables matrix
    ReDim YDATA_VECTOR(1 To j, 1 To 1) 'will become observations matrix
    j = 1
    For i = 1 To nLOOPS
        If CE_MATRIX(i, k) > 0 Then
            XTEMP_VAL = SIMUL_MATRIX(i, k)
            XDATA_MATRIX(j, 1) = 1
            XDATA_MATRIX(j, 2) = XTEMP_VAL
            XDATA_MATRIX(j, 3) = XTEMP_VAL * XTEMP_VAL
            YDATA_VECTOR(j, 1) = CC_MATRIX(i, k + 1) * DISC_VAL
            j = j + 1
        End If
    Next i
    
    OLS_MATRIX = REGRESSION_MULT_COEF_FUNC(XDATA_MATRIX, YDATA_VECTOR, False, 0)
    For i = 1 To 3
        PARAM_MATRIX(i, k) = OLS_MATRIX(i, 1)
    Next i
    For i = 1 To nLOOPS
        'compare regressed continuation value with immediate
        'exercise value to determine decision for exercise
        XTEMP_VAL = SIMUL_MATRIX(i, k)
        YTEMP_VAL = OLS_MATRIX(1, 1) + OLS_MATRIX(2, 1) * _
        XTEMP_VAL + OLS_MATRIX(3, 1) * XTEMP_VAL * XTEMP_VAL
        If CE_MATRIX(i, k) > 0 And CE_MATRIX(i, k) > YTEMP_VAL Then
            'option should be exercised
            EF_MATRIX(i, k) = 1
            For h = k + 1 To nSTEPS
                EF_MATRIX(i, h) = 0
            Next h
            CC_MATRIX(i, k) = CE_MATRIX(i, k)
        Else
            'option should not be exercised
            'hence take continuation value by discounting
            'from next period continuation value
            CC_MATRIX(i, k) = CC_MATRIX(i, k + 1) * DISC_VAL
            EF_MATRIX(i, k) = 0
        End If
    Next i
    
    If k = TIME_STEP Then
        ReDim TEMP_MATRIX(1 To STOCK_PRICE_POINTS, 1 To 3)
        STEP_SIZE = (MAX_STOCK_PRICE - MIN_STOCK_PRICE) / STOCK_PRICE_POINTS
        XTEMP_VAL = MIN_STOCK_PRICE
        For i = 1 To STOCK_PRICE_POINTS
            TEMP_MATRIX(i, 1) = XTEMP_VAL
            TEMP_MATRIX(i, 2) = OLS_MATRIX(1, 1) + OLS_MATRIX(2, 1) * XTEMP_VAL + OLS_MATRIX(3, 1) * XTEMP_VAL * XTEMP_VAL
            TEMP_MATRIX(i, 3) = MAXIMUM_FUNC(OPTION_FLAG * (XTEMP_VAL - STRIKE), 0)
            XTEMP_VAL = XTEMP_VAL + STEP_SIZE
        Next i
        LSM_AMERICAN_OPTION_POLICY_FUNC = TEMP_MATRIX
        'ExerciseRegion/ExerciseAnalysis
        'Stock Price
        'Regressed Continuation Value
        'Immediate Exercise Value
        Exit Function
    End If
    
    If IsArray(PARAM_RNG) = True Then
        PUB_STRIKE_VAL = STRIKE
        PUB_OPTION_FLAG = OPTION_FLAG
        PUB_PARAMETERS = PARAM_RNG 'OLS_MATRIX
        If UBound(PUB_PARAMETERS, 1) = 1 Then
            PUB_PARAMETERS = MATRIX_TRANSPOSE_FUNC(PUB_PARAMETERS)
        End If
        LSM_AMERICAN_OPTION_POLICY_FUNC = BRENT_ZERO_FUNC(BRENT_LOWER_BOUND, BRENT_UPPER_BOUND, "LSM_AMERICAN_OPTION_POLICY_OBJ_FUNC", BRENT_GUESS_VAL, BRENT_CONVERG_VAL, BRENT_COUNTER, BRENT_nLOOPS, BRENT_epsilon)
        Exit Function
    End If
Next k

TEMP_SUM = 0

For i = 1 To nLOOPS
    For k = 2 To nSTEPS
        If EF_MATRIX(i, k) = 1 Then
            XTEMP_VAL = SIMUL_MATRIX(i, k)
            TEMP_SUM = TEMP_SUM + Exp(-RISK_FREE_RATE * (k - 1) * DELTA_VAL) * MAXIMUM_FUNC(OPTION_FLAG * (XTEMP_VAL - STRIKE), 0)
            Exit For
        End If
    Next k
Next i

'lets evaluate optimal exercise frontier
'For k = nSTEPS - 1 To 2 Step -1
'Next k

Select Case OUTPUT
Case 0 'MCAmericanPrice
    LSM_AMERICAN_OPTION_POLICY_FUNC = TEMP_SUM / nLOOPS
Case 1
    LSM_AMERICAN_OPTION_POLICY_FUNC = SIMUL_MATRIX
Case 2
    LSM_AMERICAN_OPTION_POLICY_FUNC = EF_MATRIX
Case 3
    LSM_AMERICAN_OPTION_POLICY_FUNC = CE_MATRIX
Case 4
    LSM_AMERICAN_OPTION_POLICY_FUNC = CC_MATRIX
Case 5
    LSM_AMERICAN_OPTION_POLICY_FUNC = PARAM_MATRIX
Case Else
    ReDim TEMP_GROUP(1 To 6)
    TEMP_GROUP(1) = TEMP_SUM / nLOOPS
    TEMP_GROUP(2) = SIMUL_MATRIX
    TEMP_GROUP(3) = EF_MATRIX
    TEMP_GROUP(4) = CE_MATRIX
    TEMP_GROUP(5) = CC_MATRIX
    TEMP_GROUP(6) = PARAM_MATRIX
    LSM_AMERICAN_OPTION_POLICY_FUNC = TEMP_GROUP
End Select

Exit Function
ERROR_LABEL:
LSM_AMERICAN_OPTION_POLICY_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************



'************************************************************************************
'************************************************************************************
'FUNCTION      : LSM_AMERICAN_OPTION_POLICY_OBJ_FUNC
'DESCRIPTION   :
'LIBRARY       : DERIVATIVES
'GROUP         : AMER_LSM_GU
'ID            : 006
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/29/2009
'************************************************************************************
'************************************************************************************

Private Function LSM_AMERICAN_OPTION_POLICY_OBJ_FUNC( _
ByVal XTEMP_VAL As Double) As Double

Dim ATEMP_VAL As Double
Dim BTEMP_VAL As Double

On Error GoTo ERROR_LABEL

ATEMP_VAL = PUB_PARAMETERS(1, 1) + _
            PUB_PARAMETERS(2, 1) * XTEMP_VAL + _
            PUB_PARAMETERS(3, 1) * XTEMP_VAL * XTEMP_VAL

BTEMP_VAL = MAXIMUM_FUNC(PUB_OPTION_FLAG * (XTEMP_VAL - PUB_STRIKE_VAL), 0)

LSM_AMERICAN_OPTION_POLICY_OBJ_FUNC = ATEMP_VAL - BTEMP_VAL

Exit Function
ERROR_LABEL:
LSM_AMERICAN_OPTION_POLICY_OBJ_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************



'************************************************************************************
'************************************************************************************
'FUNCTION      : AMERICAN_OPTION_PATH_SIMULATION_FUNC
'DESCRIPTION   :
'LIBRARY       : DERIVATIVES
'GROUP         : AMER_LSM_GU
'ID            : 007
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/29/2009
'************************************************************************************
'************************************************************************************

Private Function AMERICAN_OPTION_PATH_SIMULATION_FUNC(ByVal nLOOPS As Long, _
ByVal nSTEPS As Long, _
ByVal SPOT As Double, _
ByVal RISK_FREE_RATE As Double, _
ByVal VOLATILITY As Double, _
ByVal EXPIRATION As Double)
  
Dim i As Long
Dim j As Long
Dim k As Long

Dim MULT_VAL As Double
Dim DELTA_VAL As Double
Dim DRIFT_VAL As Double
Dim FACTOR_VAL As Double

Dim RND_ARR As Variant
Dim TEMP_MATRIX As Variant

On Error GoTo ERROR_LABEL

DELTA_VAL = EXPIRATION / nSTEPS
DRIFT_VAL = (RISK_FREE_RATE - 0.5 * VOLATILITY ^ 2) * DELTA_VAL
FACTOR_VAL = VOLATILITY * DELTA_VAL ^ 0.5
RND_ARR = VECTOR_RANDOM_BOX_MULLER_FUNC(nLOOPS * nSTEPS)

k = 1
ReDim TEMP_MATRIX(1 To nLOOPS, 1 To nSTEPS)

For i = 1 To nLOOPS
  TEMP_MATRIX(i, 1) = SPOT
  MULT_VAL = SPOT
  For j = 2 To nSTEPS
    MULT_VAL = MULT_VAL * Exp(DRIFT_VAL + FACTOR_VAL * RND_ARR(k))
    k = k + 1
    TEMP_MATRIX(i, j) = MULT_VAL
  Next j
Next i

AMERICAN_OPTION_PATH_SIMULATION_FUNC = TEMP_MATRIX

Exit Function
ERROR_LABEL:
AMERICAN_OPTION_PATH_SIMULATION_FUNC = Err.number
End Function
