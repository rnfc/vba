Attribute VB_Name = "FINAN_DERIV_EURO_VG_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.

'************************************************************************************
'************************************************************************************
'FUNCTION      : VG_EUROPEAN_OPTION_FUNC
'DESCRIPTION   : Calculates Price of European Option in VG model using
'Finite Difference

'Reference: Madan, Carr, Chang, The Variance Gamma
'Process and Option Pricing (1998), Appendix, p.99 ff

'The scenario is the following:

'integrate cdfN(f(a,b,x))*pdfGamma(c,x) over the positive reals  with f(x) =
'= a/sqrt(x)+b*sqrt(x), cdfN is the cumulative normal distribution, pdfGamma the
'density of the gamma distribution cdfGamma (called P in Abramowitz & Stegun).

'The limits of f depend on the signs of a and b, so the first argument will
'reach 0 or 1. Or 1/2. Now estimate where to Replace it by that constant (and
'justify why ... but i guess for mj's solution one would have to as well).

'So for being 'close to 0' one switches to a cdfGamma and for the infinite
'tail use an adaptive integration (instead of cutting of) which stops if the
'integrand is decreasing and becomes small.

'For c = time/nu getting large (i.e. beyond 170, where most implementations
'will quitt - but then impl volatility is almost constant) one can switch to
'asymptotics (which i did without extreme care).

'For a=0 or b=0 one can even write down 'easy' explicite solutions in terms of
'hypergeometric functions.
'--------------------------------------------------------------------------------
'--------------------------------------------------------------------------------

'LIBRARY       : DERIVATIVES
'GROUP         : EURO_VG
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/29/2009
'************************************************************************************
'************************************************************************************

Function VG_EUROPEAN_OPTION_FUNC(ByVal SPOT As Double, _
ByVal STRIKE As Double, _
ByVal TENOR As Double, _
ByVal RATE As Double, _
ByVal DIVD As Double, _
ByVal SIGMA As Double, _
ByVal VG_NU As Double, _
ByVal VG_THETA As Double, _
Optional ByVal STOCK_STEPS As Long = 200, _
Optional ByVal TIME_STEPS As Long = 200, _
Optional ByVal OPTION_FLAG As Integer = 1)

'VG_NU : nu parameter of VG model
'VG_THETA : theta parameter for VG model
'STOCK_STEPS : no of stock steps
'TIME_STEPS : no of time steps

  Dim i As Long '
  Dim j As Long
  Dim k As Long
  Dim l As Long
  
  Dim SROW As Long
  Dim NROWS As Long
  Dim NSIZE As Long
  
  Dim Y1_VAL As Double
  Dim Y2_VAL As Double
  
  Dim X1_VAL As Double
  Dim X2_VAL As Double
  
  Dim TEMP_SUM As Double
  Dim TEMP_MULT As Double
  
  Dim MIN_VAL As Double
  Dim MAX_VAL As Double
  
  Dim OPTION_ARR As Variant
  Dim LSH_MATRIX As Variant
  
  Dim BP_ARR As Variant
  Dim DP_ARR As Variant
  
  Dim RHS_ARR As Variant
  Dim RESID_ARR As Variant
  Dim TEMP_ARR As Variant

  Dim ATEMP_ARR As Variant
  Dim BTEMP_ARR As Variant
  Dim CTEMP_ARR As Variant
  Dim DTEMP_ARR As Variant
'--------------------------------------------------------------------------
  Dim DELTA_VAL As Double 'stock step size
  Dim DELTA_TIME As Double 'timestep size
  Dim SPOT_ARR As Variant 'spot price
'--------------------------------------------------------------------------
'cached vectors for PIDE
  Dim EXP_INT_LP_ARR As Variant
  Dim EXP_INT_LN_ARR As Variant
  Dim EXP_LP_ARR As Variant
  Dim EXP_LN_ARR As Variant
  Dim EXP_PLUS_ARR As Variant
  Dim EXP_MIN_ARR As Variant
'--------------------------------------------------------------------------
'coefficinets for tridiagonal matrix
  Dim ATEMP_VAL As Double
  Dim BTEMP_VAL As Double
  Dim CTEMP_VAL As Double
'--------------------------------------------------------------------------
'cached values
  Dim omega As Double
  Dim FIRST_LAMBDA As Double
  Dim SECOND_LAMBDA As Double
'--------------------------------------------------------------------------
  
  On Error GoTo ERROR_LABEL
  
  omega = Log(1 - VG_THETA * VG_NU - SIGMA * SIGMA * VG_NU / 2) / VG_NU
  FIRST_LAMBDA = ((VG_THETA ^ 2) / (SIGMA ^ 4) + 2 / _
            (SIGMA ^ 2 * VG_NU)) ^ 0.5 + VG_THETA / (SIGMA ^ 2)
  SECOND_LAMBDA = ((VG_THETA ^ 2) / (SIGMA ^ 4) + 2 / _
            (SIGMA ^ 2 * VG_NU)) ^ 0.5 - VG_THETA / (SIGMA ^ 2)
  
  MIN_VAL = Log(4)
  MAX_VAL = Log(SPOT * 2)
  DELTA_VAL = (MAX_VAL - MIN_VAL) / STOCK_STEPS
  DELTA_TIME = TENOR / TIME_STEPS
  
  ReDim SPOT_ARR(0 To STOCK_STEPS)
  For i = 0 To STOCK_STEPS
    SPOT_ARR(i) = Exp(MIN_VAL + DELTA_VAL * i)
  Next i

'-------------------------------------------------------------------------------------
'-------------------------------------------------------------------------------------
'------------------stores vectors for performance improvement-------------------------
'-------------------------------------------------------------------------------------
'-------------------------------------------------------------------------------------
  
  ReDim EXP_INT_LP_ARR(0 To STOCK_STEPS)
  ReDim EXP_INT_LN_ARR(0 To STOCK_STEPS)
  ReDim EXP_LP_ARR(0 To STOCK_STEPS)
  ReDim EXP_LN_ARR(0 To STOCK_STEPS)
  ReDim EXP_PLUS_ARR(0 To STOCK_STEPS)
  ReDim EXP_MIN_ARR(0 To STOCK_STEPS)
  
  For i = 0 To STOCK_STEPS
    EXP_INT_LP_ARR(i) = ASYMPTOTIC_EXPANSION_FUNC((i + 1) * _
                        DELTA_VAL * SECOND_LAMBDA, 100, 2 ^ (-52))
    EXP_INT_LN_ARR(i) = ASYMPTOTIC_EXPANSION_FUNC((i + 1) * _
                        DELTA_VAL * FIRST_LAMBDA, 100, 2 ^ (-52))
    EXP_LP_ARR(i) = Exp(-(i + 1) * DELTA_VAL * SECOND_LAMBDA)
    EXP_LN_ARR(i) = Exp(-(i + 1) * DELTA_VAL * FIRST_LAMBDA)
    EXP_PLUS_ARR(i) = ASYMPTOTIC_EXPANSION_FUNC((i + 1) * _
                        DELTA_VAL * (FIRST_LAMBDA + 1), 100, 2 ^ (-52))
    EXP_MIN_ARR(i) = ASYMPTOTIC_EXPANSION_FUNC((i + 1) * _
                        DELTA_VAL * (SECOND_LAMBDA - 1), 100, 2 ^ (-52))
  Next i
  

'rollback time axis till the current time
'return the last option price vector
  
'------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------
'---------------Form the left hand side matrix for set of equations------------------
'------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------
  
  ReDim LSH_MATRIX(1 To STOCK_STEPS - 1, 1 To STOCK_STEPS - 1)
  
  ATEMP_VAL = (RATE - DIVD + omega) * DELTA_TIME / (2 * DELTA_VAL)
  CTEMP_VAL = -ATEMP_VAL
  BTEMP_VAL = 1 + RATE * DELTA_TIME
  
  For i = 1 To STOCK_STEPS - 2
    LSH_MATRIX(i + 1, i) = ATEMP_VAL
  Next i
  For i = 1 To STOCK_STEPS - 1
    LSH_MATRIX(i, i) = BTEMP_VAL
  Next i
  For i = 1 To STOCK_STEPS - 2
    LSH_MATRIX(i, i + 1) = CTEMP_VAL
  Next i
  
  ReDim TEMP_ARR(0 To STOCK_STEPS) 'returns payoff at final node
  
  For i = STOCK_STEPS To 0 Step -1
    Select Case OPTION_FLAG
        Case 1 ', "c", "call"
          TEMP_ARR(i) = IIf((SPOT_ARR(i) - STRIKE) >= 0, (SPOT_ARR(i) - STRIKE), 0)
        Case Else 'PUT
          TEMP_ARR(i) = IIf((STRIKE - SPOT_ARR(i)) >= 0, (STRIKE - SPOT_ARR(i)), 0)
    End Select
  Next i
  OPTION_ARR = TEMP_ARR
    
For j = TIME_STEPS - 1 To 0 Step -1
    Select Case OPTION_FLAG
        Case 1 ', "c", "call"
      OPTION_ARR(0) = 0
      OPTION_ARR(STOCK_STEPS) = SPOT_ARR(STOCK_STEPS) * Exp(-DIVD * _
                            (TIME_STEPS - j) * DELTA_TIME) - _
                            STRIKE * Exp(-RATE * (TIME_STEPS - j) * DELTA_TIME)
        Case Else 'Put
      OPTION_ARR(0) = STRIKE * Exp(-RATE * (TIME_STEPS - j) * DELTA_TIME) - _
                            SPOT_ARR(0) * Exp(-DIVD * (TIME_STEPS - j) _
                            * DELTA_TIME)
      OPTION_ARR(STOCK_STEPS) = 0
    End Select
    
'returns right hand side vector for set of linear equations
'OPTION_ARR is the option price vec for previous time step
'j is index of time period
  
  ReDim TEMP_ARR(1 To STOCK_STEPS - 1)

  For i = 1 To STOCK_STEPS - 1 'Move up the stock axis
    TEMP_SUM = 0
    
    For k = 1 To STOCK_STEPS - i - 1
      TEMP_SUM = TEMP_SUM + (OPTION_ARR(i + k + 1) - OPTION_ARR(i + k)) * _
                (EXP_LP_ARR(k - 1) - EXP_LP_ARR(k)) / (VG_NU * _
                DELTA_VAL * SECOND_LAMBDA)
      TEMP_SUM = TEMP_SUM + (OPTION_ARR(i + k) - OPTION_ARR(i) - k * _
                (OPTION_ARR(i + k + 1) - OPTION_ARR(i + k))) * _
                (EXP_INT_LP_ARR(k - 1) - EXP_INT_LP_ARR(k)) / VG_NU
    Next k
    
    For k = 1 To i - 1
      TEMP_SUM = TEMP_SUM + (OPTION_ARR(i - k - 1) - OPTION_ARR(i - k)) * _
                (EXP_LN_ARR(k - 1) - EXP_LN_ARR(k)) / (VG_NU * _
                DELTA_VAL * FIRST_LAMBDA)
      TEMP_SUM = TEMP_SUM + (OPTION_ARR(i - k) - OPTION_ARR(i) - k * _
                (OPTION_ARR(i - k - 1) - OPTION_ARR(i - k))) * _
                (EXP_INT_LN_ARR(k - 1) - EXP_INT_LN_ARR(k)) / VG_NU
    Next k
    
    TEMP_SUM = TEMP_SUM + (OPTION_ARR(i + 1) - OPTION_ARR(i)) * _
                (1 - EXP_LP_ARR(0)) / (VG_NU * DELTA_VAL * SECOND_LAMBDA)
    
    TEMP_SUM = TEMP_SUM + (OPTION_ARR(i - 1) - OPTION_ARR(i)) * _
                (1 - EXP_LN_ARR(0)) / (VG_NU * DELTA_VAL * FIRST_LAMBDA)
    
    'The remaining integral terms will change depending on type of option
    Select Case OPTION_FLAG
        Case 1 ', "c", "call"

      TEMP_SUM = TEMP_SUM + (Exp(-DIVD * (TIME_STEPS - j - 1) * DELTA_TIME) * _
                SPOT_ARR(i) * EXP_MIN_ARR(STOCK_STEPS - i - 1) _
                - (STRIKE * Exp(-RATE * (TIME_STEPS - j - 1) * DELTA_TIME) _
                + OPTION_ARR(i)) * EXP_INT_LP_ARR(STOCK_STEPS - i - 1) _
                - EXP_INT_LN_ARR(i - 1) * OPTION_ARR(i)) / VG_NU
        Case Else 'Put
      TEMP_SUM = TEMP_SUM + (((STRIKE * Exp(-RATE * (TIME_STEPS - j - 1) * _
                DELTA_TIME) - OPTION_ARR(i)) * EXP_INT_LN_ARR(i - 1) _
                - Exp(-DIVD * (TIME_STEPS - j - 1) * _
                DELTA_TIME) * SPOT_ARR(i) * EXP_PLUS_ARR(i - 1) - _
                EXP_INT_LP_ARR(STOCK_STEPS - i - 1) * OPTION_ARR(i)) / VG_NU)
    End Select
  
    TEMP_ARR(i) = OPTION_ARR(i) + DELTA_TIME * TEMP_SUM
  
  Next i
  
  TEMP_ARR(1) = TEMP_ARR(1) - ATEMP_VAL * OPTION_ARR(0)
  TEMP_ARR(STOCK_STEPS - 1) = TEMP_ARR(STOCK_STEPS - 1) - _
                            CTEMP_VAL * OPTION_ARR(STOCK_STEPS)
  RHS_ARR = TEMP_ARR

'---------------------------------------------------------------------------------
'---------------------------Solve Tridiagonal System------------------------------
'---------------------------------------------------------------------------------
'Solves a tridiagonal system of equations
'using procedure as described in: Tridiagonal matrix algorithm
'http://en.wikipedia.org/wiki/Tridiagonal_matrix_algorithm
'---------------------------------------------------------------------------------
'LSH_MATRIX is the left hand side tridiagonal matrix
'RHS_ARR is right hand side vector
'returned vector is RESID_ARR
'---------------------------------------------------------------------------------
  
  SROW = LBound(LSH_MATRIX, 1)
  NROWS = UBound(LSH_MATRIX, 1)
  NSIZE = NROWS - SROW + 1
  
  'Extract the vectors A,B,C of the tridiagonal matrix
  
  ReDim ATEMP_ARR(1 To NSIZE)
  ReDim BTEMP_ARR(1 To NSIZE)
  ReDim CTEMP_ARR(1 To NSIZE)
  ReDim DTEMP_ARR(1 To NSIZE)
  
  ATEMP_ARR(1) = 0
  CTEMP_ARR(NSIZE) = 0
  
  For i = 1 To NSIZE
    If i > 1 Then: ATEMP_ARR(i) = LSH_MATRIX(SROW + i - 1, SROW + i - 2)
    BTEMP_ARR(i) = LSH_MATRIX(SROW + i - 1, SROW + i - 1)
    If i < NSIZE Then: CTEMP_ARR(i) = LSH_MATRIX(SROW + i - 1, SROW + i)
  Next i
  
  DTEMP_ARR = RHS_ARR
  ReDim BP_ARR(1 To NSIZE)
  ReDim DP_ARR(1 To NSIZE)
  
  BP_ARR(1) = BTEMP_ARR(1)
  DP_ARR(1) = DTEMP_ARR(1)
    
  For k = 2 To NSIZE
    TEMP_MULT = ATEMP_ARR(k) / BP_ARR(k - 1)
    BP_ARR(k) = BTEMP_ARR(k) - TEMP_MULT * CTEMP_ARR(k - 1)
    DP_ARR(k) = DTEMP_ARR(k) - TEMP_MULT * DP_ARR(k - 1)
  Next k
  
  ReDim RESID_ARR(1 To NSIZE)
  
  RESID_ARR(NSIZE) = DP_ARR(NSIZE) / BP_ARR(NSIZE)
  
  For k = NSIZE - 1 To 1 Step -1
    RESID_ARR(k) = (DP_ARR(k) - CTEMP_ARR(k) * RESID_ARR(k + 1)) / BP_ARR(k)
  Next k
  
'---------------------------------------------------------------------------------
  l = 1
  For i = LBound(RESID_ARR, 1) To UBound(RESID_ARR, 1)
      OPTION_ARR(l) = RESID_ARR(i)
      l = l + 1
  Next i
Next j
'------------------------OPTION VAL INTERPOLATION--------------------------------
  
  If ((LBound(SPOT_ARR, 1) <> LBound(OPTION_ARR, 1)) Or _
      (UBound(SPOT_ARR, 1) <> UBound(OPTION_ARR, 1))) Then: GoTo ERROR_LABEL
    'Interpolate :spot and opt do not match
  
  If ((SPOT < SPOT_ARR(LBound(OPTION_ARR, 1))) Or _
        (SPOT > SPOT_ARR(UBound(OPTION_ARR, 1)))) Then: GoTo ERROR_LABEL
  'Interpolate : SPOT out of range

  For i = LBound(SPOT_ARR, 1) To UBound(SPOT_ARR, 1)
    If SPOT <= SPOT_ARR(i) Then
      If SPOT = SPOT_ARR(i) Then
        VG_EUROPEAN_OPTION_FUNC = OPTION_ARR(i)
        Exit Function
      Else
        X1_VAL = SPOT_ARR(i - 1)
        Y1_VAL = OPTION_ARR(i - 1)
        X2_VAL = SPOT_ARR(i)
        Y2_VAL = OPTION_ARR(i)
        VG_EUROPEAN_OPTION_FUNC = Y1_VAL + (Y2_VAL - Y1_VAL) * _
                    (SPOT - X1_VAL) / (X2_VAL - X1_VAL)
        Exit Function
      End If
    End If
  Next i
  
'--------------------------Simulation parameters settings------------------------------
'S=100; %underlying price
'K=101; %strike
'T=1; %maturity
'sigma=0.3; %volatility for VG model
'r=0.066; %risk free rate
'VG_nu=.25; %nu for VG model
'VG_theta=-0.3; %theta of VG model
'nsimulations=100000; % no. of MC simulations
'CallPutFlag="P"; %Enter C for Call and P for Put
'--------------------------------------------------------------------------------------
'omega=(1/VG_nu)*( log(1-VG_theta*VG_nu-sigma*sigma*VG_nu/2) );

'dt=T;
'thmean=dt;
'thvar=VG_nu*dt;
'theta=thmean/thvar;
'alpha=thmean*theta;

'G2vec=gamma_rnd(alpha,theta,nsimulations,1);
'X2vec=normal_rnd(VG_theta*G2vec,sigma*sigma*G2vec);
'S2vec=exp( log(S)+r*T+omega*T+X2vec );
'if CallPutFlag=="C",
'    payoffvec=max(S2vec-k,0);
'Else
'    payoffvec=max(k-S2vec,0);
'End If
'mc_callprice = Exp(-r * T) * mean(payoffvec)
'--------------------------------------------------------------------------------------
  
Exit Function
ERROR_LABEL:
VG_EUROPEAN_OPTION_FUNC = Err.number
End Function

