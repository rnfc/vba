Attribute VB_Name = "FINAN_DERIV_AMER_RND_LIBR"
Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.

'************************************************************************************
'************************************************************************************
'FUNCTION      : RANDOM_TREE_AMERICAN_OPTION_FUNC
'DESCRIPTION   : Pricing American options using random tree, as described in
'Monte Carlo Methods in Financial Engineering
'(Stochastic Modelling and Applied Probability)
'http://www2.gsb.columbia.edu/faculty/pglasserman/Other/
'LIBRARY       : DERIVATIVES
'GROUP         : AMER_RND_LIBR
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/29/2009
'************************************************************************************
'************************************************************************************

Function RANDOM_TREE_AMERICAN_OPTION_FUNC(ByVal SPOT As Double, _
ByVal STRIKE As Double, _
ByVal RISK_FREE_RATE As Double, _
ByVal VOLATILITY As Double, _
ByVal EXPIRATION As Double, _
Optional ByVal BRANCHING As Long = 8, _
Optional ByVal nSTEPS As Long = 4, _
Optional ByVal nLOOPS As Long = 50)

Dim h As Long
Dim i As Long
Dim j As Long
Dim k As Long
Dim l As Long
Dim m As Long

Dim NSIZE As Long

Dim HIGH_SUM As Double
Dim LOW_SUM As Double

Dim TEMP_VAL As Double
Dim TEMP_DELTA As Double
Dim TEMP_VOLAT As Double
Dim TEMP_DRIFT As Double
Dim TEMP_SUM As Double 'value of expected payoff while ignoring
  'the current node

Dim HIGH_ARR As Variant
Dim LOW_ARR As Variant

Dim HIGH_TREE As Variant
Dim LOW_TREE As Variant

Dim IN_ARR As Variant
Dim OUT_ARR As Variant

Dim PREV_ARR As Variant
Dim PRICE_TREE_ARR As Variant

Dim INIT_ARR As Variant
Dim ATEMP_ARR As Variant
Dim BTEMP_ARR As Variant
Dim CTEMP_ARR As Variant 'Option price array for successor nodes
  
Dim CURRENT_PAYOFF As Double
Dim EXPECTED_PAYOFF As Double

On Error GoTo ERROR_LABEL

ReDim HIGH_ARR(0 To nLOOPS)
ReDim LOW_ARR(0 To nLOOPS)

HIGH_SUM = 0
LOW_SUM = 0

For i = 0 To nLOOPS

'----------------------------GET RANDOM TREE---------------------------
  ReDim BTEMP_ARR(0 To nSTEPS)
  ReDim INIT_ARR(0 To 0)
  INIT_ARR(0) = SPOT
  BTEMP_ARR(0) = INIT_ARR
  TEMP_DELTA = EXPIRATION / nSTEPS
  TEMP_DRIFT = (RISK_FREE_RATE - VOLATILITY ^ 2 / 2) * TEMP_DELTA
  TEMP_VOLAT = VOLATILITY * Sqr(TEMP_DELTA)

  For m = 1 To nSTEPS
    NSIZE = BRANCHING ^ m
    PREV_ARR = BTEMP_ARR(m - 1)
    ReDim ATEMP_ARR(0 To NSIZE - 1)
    For j = 0 To NSIZE - 1
      h = Int(j / BRANCHING)
      TEMP_VAL = PREV_ARR(h)
      ATEMP_ARR(j) = TEMP_VAL * Exp(TEMP_DRIFT + TEMP_VOLAT * _
                     TEMP_DELTA ^ 0.5 * RANDOM_POLAR_MARSAGLIA_FUNC())
    Next j
    BTEMP_ARR(m) = ATEMP_ARR
  Next m
  PRICE_TREE_ARR = BTEMP_ARR

'----------------------------Get High Estimators---------------------------
  
  IN_ARR = PRICE_TREE_ARR(UBound(PRICE_TREE_ARR, 1))
  
  ReDim ATEMP_ARR(0 To UBound(IN_ARR, 1))
  ReDim OUT_ARR(0 To UBound(PRICE_TREE_ARR, 1))
  For j = 0 To UBound(ATEMP_ARR, 1)
    TEMP_VAL = IN_ARR(j)
    ATEMP_ARR(j) = MAXIMUM_FUNC(TEMP_VAL - STRIKE, 0)
  Next j
  OUT_ARR(UBound(OUT_ARR, 1)) = ATEMP_ARR
  
  For m = UBound(PRICE_TREE_ARR, 1) - 1 To 0 Step -1
    BTEMP_ARR = OUT_ARR(m + 1)
    IN_ARR = PRICE_TREE_ARR(m)
    ReDim ATEMP_ARR(0 To UBound(IN_ARR, 1))
    For j = 0 To UBound(BTEMP_ARR, 1) Step BRANCHING
      h = Int(j / BRANCHING)
      EXPECTED_PAYOFF = 0
      For k = j To j + BRANCHING - 1
        EXPECTED_PAYOFF = EXPECTED_PAYOFF + BTEMP_ARR(k)
      Next k
      EXPECTED_PAYOFF = EXPECTED_PAYOFF / BRANCHING
       TEMP_VAL = IN_ARR(h)
      CURRENT_PAYOFF = MAXIMUM_FUNC(TEMP_VAL - STRIKE, 0)
      'at each node compare current payoff vs. expected payoff
      ATEMP_ARR(h) = MAXIMUM_FUNC(CURRENT_PAYOFF, EXPECTED_PAYOFF)
    Next j
    OUT_ARR(m) = ATEMP_ARR
  Next m
  
  HIGH_TREE = OUT_ARR

'----------------------------Get Low Estimators---------------------------
  
  IN_ARR = PRICE_TREE_ARR(UBound(PRICE_TREE_ARR, 1))
  
  ReDim ATEMP_ARR(0 To UBound(IN_ARR, 1))
  ReDim OUT_ARR(0 To UBound(PRICE_TREE_ARR, 1))
  For j = 0 To UBound(ATEMP_ARR, 1)
    TEMP_VAL = IN_ARR(j)
    ATEMP_ARR(j) = MAXIMUM_FUNC(TEMP_VAL - STRIKE, 0)
  Next j
  OUT_ARR(UBound(OUT_ARR, 1)) = ATEMP_ARR
  
  For m = UBound(PRICE_TREE_ARR, 1) - 1 To 0 Step -1
    BTEMP_ARR = OUT_ARR(m + 1)
    IN_ARR = PRICE_TREE_ARR(m)
    ReDim ATEMP_ARR(0 To UBound(IN_ARR, 1))
    For j = 0 To UBound(BTEMP_ARR, 1) Step BRANCHING
      h = Int(j / BRANCHING)
      
      ReDim CTEMP_ARR(0 To BRANCHING - 1)
      TEMP_VAL = IN_ARR(h)
      
      CURRENT_PAYOFF = MAXIMUM_FUNC(TEMP_VAL - STRIKE, 0)
      For k = 0 To BRANCHING - 1
        TEMP_SUM = 0
        For l = 0 To BRANCHING - 1
          If l <> k Then: TEMP_SUM = TEMP_SUM + BTEMP_ARR(j + l)
        Next l
        TEMP_SUM = TEMP_SUM / (BRANCHING - 1)
        If TEMP_SUM > CURRENT_PAYOFF Then 'it is better to continue as
        'we will get more payoff
          CTEMP_ARR(k) = BTEMP_ARR(j + k)
        Else 'do not continue and exercise option now
          CTEMP_ARR(k) = CURRENT_PAYOFF
        End If
      Next k
      TEMP_SUM = 0
      For k = 0 To BRANCHING - 1
        TEMP_SUM = TEMP_SUM + CTEMP_ARR(k)
      Next k
      ATEMP_ARR(h) = TEMP_SUM / BRANCHING
    Next j
    OUT_ARR(m) = ATEMP_ARR
  Next m
  
  LOW_TREE = OUT_ARR
'-----------------------------------------------------------------------
  HIGH_SUM = HIGH_SUM + HIGH_TREE(0)(0)
  LOW_SUM = LOW_SUM + LOW_TREE(0)(0)
Next i

'High Estimator, Low Estimator
RANDOM_TREE_AMERICAN_OPTION_FUNC = Array(HIGH_SUM / nLOOPS, LOW_SUM / nLOOPS)

Exit Function
ERROR_LABEL:
RANDOM_TREE_AMERICAN_OPTION_FUNC = Err.number
End Function
