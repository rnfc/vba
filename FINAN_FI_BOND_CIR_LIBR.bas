Attribute VB_Name = "FINAN_FI_BOND_CIR_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.

'**********************************************************************************
'**********************************************************************************
'FUNCTION      : CIR_RATE_FUNC
'DESCRIPTION   : Function CIR_DF to calculate CIR's zero discount function,
'  i.e. present value factor or price of a zero bond with maturity.
'LIBRARY       : BOND
'GROUP         : CIR
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 21/01/2009
'**********************************************************************************
'**********************************************************************************


Function CIR_RATE_FUNC(ByVal SHORT_RATE As Double, _
ByVal EQUILIBRIUM As Double, _
ByVal PULL_BACK As Double, _
ByVal SIGMA As Double, _
ByVal TENOR As Double, _
Optional ByVal OUTPUT As Integer = 0)

' SHORT_RATE: Interest rate at time MIN_TENOR.

' EQUILIBRIUM: Long-run short-term interest rate (equals b in Hull notation)
' long-term equilibrium of mean reverting spot rate process.

' PULL_BACK: "Pull-back" factor of interest rate (strength of mean reversion)
' a in Hull notation --> Speed of Adjustment.

' SIGMA: Instantaneous standard deviation of spot rate.

' TENOR: Maturity time (T)

Dim MEAN_RATE As Double

Dim TEMP_VECTOR As Variant

On Error GoTo ERROR_LABEL

ReDim TEMP_VECTOR(1 To 8, 1 To 2)

'----------------------------------------------------------------------------

TEMP_VECTOR(2, 1) = ("Y IN CIR MODEL")
TEMP_VECTOR(3, 1) = ("B IN CIR MODEL")
TEMP_VECTOR(4, 1) = ("A IN CIR MODEL")

TEMP_VECTOR(2, 2) = CIR_GAMMA_FUNC(PULL_BACK, SIGMA)
TEMP_VECTOR(3, 2) = CIR_B_FACT_FUNC(PULL_BACK, SIGMA, TENOR)
TEMP_VECTOR(4, 2) = CIR_A_FACT_FUNC(EQUILIBRIUM, PULL_BACK, SIGMA, TENOR)

'-----------------------------------------------------------------------------

TEMP_VECTOR(5, 1) = ("INFINITELY-LONG RATE")
TEMP_VECTOR(6, 1) = ("CIR ZERO RATE")
TEMP_VECTOR(7, 1) = ("CIR SIGMA OF ZERO RATE")
TEMP_VECTOR(8, 1) = ("CIR DISCOUNT FACTOR")

TEMP_VECTOR(5, 2) = CIR_INFIN_RATE_FUNC(EQUILIBRIUM, PULL_BACK, SIGMA)
TEMP_VECTOR(6, 2) = CIR_ZERO_RATE_FUNC(SHORT_RATE, EQUILIBRIUM, PULL_BACK, SIGMA, TENOR)
TEMP_VECTOR(7, 2) = CIR_INSTAN_SIGM_FUNC(SHORT_RATE, EQUILIBRIUM, PULL_BACK, SIGMA, TENOR)
TEMP_VECTOR(8, 2) = CIR_DISC_FUNC(SHORT_RATE, EQUILIBRIUM, PULL_BACK, SIGMA, TENOR)

'-----------------------------------------------------------------------------
MEAN_RATE = TEMP_VECTOR(6, 2)
TEMP_VECTOR(1, 1) = ("STEADY STATE PROBABILITY DENSITY FUNCTION: " & Format(MEAN_RATE, "0.0%") & "")
TEMP_VECTOR(1, 2) = CIR_ZERO_PROB_FUNC(MEAN_RATE, EQUILIBRIUM, PULL_BACK, SIGMA)

Select Case OUTPUT
    Case 0
        CIR_RATE_FUNC = TEMP_VECTOR
    Case Else
        CIR_RATE_FUNC = MATRIX_GET_COLUMN_FUNC(TEMP_VECTOR, 2, 1)
End Select

Exit Function
ERROR_LABEL:
CIR_RATE_FUNC = Err.number
End Function


'**********************************************************************************
'**********************************************************************************
'FUNCTION      : CIR_YIELD_FUNC
'DESCRIPTION   : Term structure in Cox, C.J. Ingersoll, and J.E. Ross Model
'LIBRARY       : BOND
'GROUP         : CIR
'ID            : 002
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 21/01/2009
'**********************************************************************************
'**********************************************************************************

Function CIR_YIELD_FUNC(ByVal SETTLEMENT As Date, _
ByVal SHORT_RATE As Double, _
ByVal EQUILIBRIUM As Double, _
ByVal PULL_BACK As Double, _
ByVal SIGMA As Double, _
Optional ByRef TENOR_RNG As Variant, _
Optional ByRef HOLIDAYS_RNG As Variant, _
Optional ByVal COUNT_BASIS As Integer = 0, _
Optional ByVal OUTPUT As Integer = 1)

Dim i As Long
Dim j As Long

Dim TEMP_TENOR As Double
Dim TENOR_VECTOR As Variant
Dim TEMP_MATRIX As Variant
Dim TEMP_VECTOR As Variant

' SHORT_RATE: Short Term Interest rate today
' EQUILIBRIUM: Long-run short-term interest rate (equals b in Hull notation)
' --> long-term equilibrium of mean reverting spot rate process
' PULL_BACK: "Pull-back" factor of interest rate (strength of mean reversion)
' a in Hull notation --> SPEED OF ADJUSTMENT
' SIGMA: Instantaneous standard deviation of spot rate
' TENOR_RNG --> Maturity Time Vector (IN YEARS)

TENOR_VECTOR = TENOR_RNG

If IsArray(TENOR_VECTOR) = False Then
   TENOR_VECTOR = BOND_TENOR_TABLE_FUNC(1, SETTLEMENT, HOLIDAYS_RNG)
Else
   If UBound(TENOR_VECTOR, 1) = 1 Then: TENOR_VECTOR = MATRIX_TRANSPOSE_FUNC(TENOR_VECTOR)
End If

j = UBound(TENOR_VECTOR, 1)

ReDim TEMP_MATRIX(0 To j, 1 To 4)

TEMP_MATRIX(0, 1) = ("TN")
TEMP_MATRIX(0, 2) = ("DATE")
TEMP_MATRIX(0, 3) = ("DISCOUNT_FACTOR")
TEMP_MATRIX(0, 4) = ("ZERO_RATE")

For i = 1 To j
     
     TEMP_MATRIX(i, 1) = TENOR_VECTOR(i, 1)
     
     TEMP_MATRIX(i, 2) = TENOR_VECTOR(i, 2)
     
     TEMP_TENOR = YEARFRAC_FUNC(SETTLEMENT, TEMP_MATRIX(i, 2), COUNT_BASIS)
     
     TEMP_MATRIX(i, 3) = CIR_DISC_FUNC(SHORT_RATE, EQUILIBRIUM, PULL_BACK, _
     SIGMA, TEMP_TENOR)
     
     TEMP_MATRIX(i, 4) = CIR_ZERO_RATE_FUNC(SHORT_RATE, EQUILIBRIUM, PULL_BACK, _
     SIGMA, TEMP_TENOR)
Next i

Select Case OUTPUT
Case 0
    CIR_YIELD_FUNC = TEMP_MATRIX
Case Else 'FOR CHARTING PURPOSE
    ReDim TEMP_VECTOR(1 To 8, 1 To 2)
   
   TEMP_VECTOR(1, 1) = ("LONG-TERM EQUILIBRIUM RATE")
   TEMP_VECTOR(2, 1) = Format(SETTLEMENT, "D-MMM-YY")
   TEMP_VECTOR(3, 1) = Format(TEMP_MATRIX(j, 2), "D-MMM-YY")
   TEMP_VECTOR(4, 1) = ("SHORT-RATE AT SETTLEMENT")
   TEMP_VECTOR(5, 1) = Format(SETTLEMENT, "D-MMM-YY")
   TEMP_VECTOR(6, 1) = ("INFINITELY LONG RATE")
   TEMP_VECTOR(7, 1) = Format(SETTLEMENT, "D-MMM-YY")
   TEMP_VECTOR(8, 1) = Format(TEMP_MATRIX(j, 2), "D-MMM-YY")
   
   TEMP_VECTOR(1, 2) = ""
   TEMP_VECTOR(2, 2) = EQUILIBRIUM
   TEMP_VECTOR(3, 2) = EQUILIBRIUM
   TEMP_VECTOR(4, 2) = ""
   TEMP_VECTOR(5, 2) = SHORT_RATE
   TEMP_VECTOR(6, 2) = ""
   TEMP_VECTOR(7, 2) = CIR_INFIN_RATE_FUNC(EQUILIBRIUM, PULL_BACK, SIGMA)
   TEMP_VECTOR(8, 2) = TEMP_VECTOR(7, 2)

    CIR_YIELD_FUNC = TEMP_VECTOR
End Select

Exit Function
ERROR_LABEL:
CIR_YIELD_FUNC = Err.number
End Function

'**********************************************************************************
'**********************************************************************************
'FUNCTION      : CIR_BOND_OPT_FUNC
'DESCRIPTION   : CIR BOND OPTION MODEL
'LIBRARY       : BOND
'GROUP         : CIR
'ID            : 003
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 21/01/2009
'**********************************************************************************
'**********************************************************************************

Function CIR_BOND_OPT_FUNC(ByVal FACE As Double, _
ByVal STRIKE As Double, _
ByVal BOND_TENOR As Double, _
ByVal OPT_TENOR As Double, _
ByVal SHORT_RATE As Double, _
ByVal EQUILIBRIUM As Double, _
ByVal PULL_BACK As Double, _
ByVal SIGMA As Double, _
Optional ByVal OPTION_FLAG As Integer = 1)

'EQUILIBRIUM --> Mean reversion level ( q )
'PULL_BACK --> Speed of mean reversion ( k )
'SIGMA --> Volatility (s)
  
Dim ATEMP_VAL As Double
Dim BTEMP_VAL As Double

Dim FIRST_COMPONENT As Double
Dim SECOND_COMPONENT As Double

On Error GoTo ERROR_LABEL

If OPTION_FLAG <> 1 Then: OPTION_FLAG = -1

STRIKE = STRIKE / FACE
FIRST_COMPONENT = CIR_DISC_FUNC(SHORT_RATE, EQUILIBRIUM, _
                    PULL_BACK, SIGMA, OPT_TENOR)
SECOND_COMPONENT = CIR_DISC_FUNC(SHORT_RATE, EQUILIBRIUM, _
                PULL_BACK, SIGMA, BOND_TENOR)

BTEMP_VAL = Sqr(SIGMA ^ 2 * (1 - Exp(-2 * PULL_BACK * OPT_TENOR)) / (2 * PULL_BACK)) _
* (1 - Exp(-PULL_BACK * (BOND_TENOR - OPT_TENOR))) / PULL_BACK

ATEMP_VAL = 1 / BTEMP_VAL * Log(SECOND_COMPONENT / (FIRST_COMPONENT * STRIKE)) + BTEMP_VAL / 2

Select Case OPTION_FLAG
Case 1 ', "c"
    CIR_BOND_OPT_FUNC = FACE * (SECOND_COMPONENT * CND_FUNC(ATEMP_VAL) - STRIKE * FIRST_COMPONENT * CND_FUNC(ATEMP_VAL - BTEMP_VAL))
Case Else
    CIR_BOND_OPT_FUNC = FACE * (STRIKE * FIRST_COMPONENT * CND_FUNC(-ATEMP_VAL + BTEMP_VAL) - SECOND_COMPONENT * CND_FUNC(-ATEMP_VAL))
End Select

Exit Function
ERROR_LABEL:
CIR_BOND_OPT_FUNC = Err.number
End Function

'**********************************************************************************
'**********************************************************************************
'FUNCTION      : CIR_PROB_TBL_FUNC
'DESCRIPTION   : CIR PROB TABLE
'LIBRARY       : BOND
'GROUP         : CIR
'ID            : 004
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 21/01/2009
'**********************************************************************************
'**********************************************************************************


Function CIR_PROB_TBL_FUNC(ByVal EQUILIBRIUM As Double, _
ByVal PULL_BACK As Double, _
ByVal SIGMA As Double, _
Optional ByVal MIN_FACTOR As Double = -4.5, _
Optional ByVal MAX_FACTOR As Double = 4.5, _
Optional ByVal DELTA_FACTOR As Double = 0.5, _
Optional ByVal MULT_FACTOR As Double = 1.2, _
Optional ByVal OUTPUT As Integer = 1)

Dim i As Long
Dim NROWS As Long
Dim TEMP_MATRIX As Variant
Dim TEMP_VECTOR As Variant

On Error GoTo ERROR_LABEL

NROWS = Abs(MAX_FACTOR - MIN_FACTOR) / DELTA_FACTOR + 1

ReDim TEMP_MATRIX(0 To NROWS, 1 To 3)

TEMP_MATRIX(0, 1) = ("SIGMA_FACTOR")
TEMP_MATRIX(0, 2) = ("ZERO_RATE")
TEMP_MATRIX(0, 3) = ("PROB_DENSITY")

TEMP_MATRIX(1, 1) = MIN_FACTOR
TEMP_MATRIX(1, 2) = MAXIMUM_FUNC(EQUILIBRIUM + (SIGMA * _
        Sqr(EQUILIBRIUM / 2 / PULL_BACK)) _
        * TEMP_MATRIX(1, 1), 0)
TEMP_MATRIX(1, 3) = CIR_ZERO_PROB_FUNC(TEMP_MATRIX(1, 2), _
                EQUILIBRIUM, PULL_BACK, SIGMA)

For i = 2 To NROWS
    TEMP_MATRIX(i, 1) = TEMP_MATRIX(i - 1, 1) + DELTA_FACTOR
    TEMP_MATRIX(i, 2) = MAXIMUM_FUNC(EQUILIBRIUM + _
            (SIGMA * Sqr(EQUILIBRIUM / 2 / PULL_BACK)) _
            * TEMP_MATRIX(i, 1), 0)
    TEMP_MATRIX(i, 3) = CIR_ZERO_PROB_FUNC(TEMP_MATRIX(i, 2), _
                EQUILIBRIUM, PULL_BACK, SIGMA)
Next i

Select Case OUTPUT
Case 0
   CIR_PROB_TBL_FUNC = TEMP_MATRIX

Case Else 'FOR CHARTING PURPOSE

    ReDim TEMP_VECTOR(1 To 9, 1 To 2)
    
    TEMP_VECTOR(1, 1) = ("MEAN_RATE")
    TEMP_VECTOR(1, 2) = ""
    
    TEMP_VECTOR(2, 1) = TEMP_MATRIX(Abs(Int((MIN_FACTOR / DELTA_FACTOR) - 1)), 2)
    TEMP_VECTOR(2, 2) = 0
    
    TEMP_VECTOR(3, 1) = TEMP_VECTOR(2, 1)
    TEMP_VECTOR(3, 2) = TEMP_MATRIX(Abs(Int((MIN_FACTOR / DELTA_FACTOR) - 1)), 3) * MULT_FACTOR
    
    
    TEMP_VECTOR(4, 1) = ("+ SIGMA")
    TEMP_VECTOR(4, 2) = ""
    
    TEMP_VECTOR(5, 1) = EQUILIBRIUM + (SIGMA * Sqr(EQUILIBRIUM / 2 / PULL_BACK))
    TEMP_VECTOR(5, 2) = TEMP_VECTOR(2, 2)
    
    TEMP_VECTOR(6, 1) = TEMP_VECTOR(5, 1)
    TEMP_VECTOR(6, 2) = TEMP_VECTOR(3, 2)
    
    
    TEMP_VECTOR(7, 1) = ("- SIGMA")
    TEMP_VECTOR(7, 2) = ""
    
    TEMP_VECTOR(8, 1) = EQUILIBRIUM - (SIGMA * Sqr(EQUILIBRIUM / 2 / PULL_BACK))
    TEMP_VECTOR(8, 2) = TEMP_VECTOR(2, 2)
    
    TEMP_VECTOR(9, 1) = TEMP_VECTOR(8, 1)
    TEMP_VECTOR(9, 2) = TEMP_VECTOR(3, 2)
    
    CIR_PROB_TBL_FUNC = TEMP_VECTOR
End Select

Exit Function
ERROR_LABEL:
CIR_PROB_TBL_FUNC = Err.number
End Function


'**********************************************************************************
'**********************************************************************************
'FUNCTION      : CIR_GAMMA_FUNC
'DESCRIPTION   : CIR_GAMMA_FUNC FUNCTION
'LIBRARY       : BOND
'GROUP         : CIR
'ID            : 005
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 21/01/2009
'**********************************************************************************
'**********************************************************************************

Function CIR_GAMMA_FUNC(ByVal PULL_BACK As Double, _
ByVal SIGMA As Double)

On Error GoTo ERROR_LABEL

CIR_GAMMA_FUNC = Sqr(PULL_BACK ^ 2 + 2 * SIGMA ^ 2)

Exit Function
ERROR_LABEL:
CIR_GAMMA_FUNC = Err.number
End Function


'**********************************************************************************
'**********************************************************************************
'FUNCTION      : CIR_A_FACT_FUNC
'DESCRIPTION   : CIR MODEL FIRST FACTOR
'LIBRARY       : BOND
'GROUP         : CIR
'ID            : 006
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 21/01/2009
'**********************************************************************************
'**********************************************************************************

Function CIR_A_FACT_FUNC(ByVal EQUILIBRIUM As Double, _
ByVal PULL_BACK As Double, _
ByVal SIGMA As Double, _
ByVal TENOR As Double)

Dim GAMMA As Double

On Error GoTo ERROR_LABEL

GAMMA = CIR_GAMMA_FUNC(PULL_BACK, SIGMA)

CIR_A_FACT_FUNC = ((2 * GAMMA * Exp((PULL_BACK + GAMMA) * TENOR _
* 0.5)) / ((GAMMA + PULL_BACK) * (Exp(GAMMA * TENOR) - 1) + 2 _
* GAMMA)) ^ (2 * PULL_BACK * EQUILIBRIUM / SIGMA ^ 2)

Exit Function
ERROR_LABEL:
CIR_A_FACT_FUNC = Err.number
End Function


'**********************************************************************************
'**********************************************************************************
'FUNCTION      : CIR_B_FACT_FUNC
'DESCRIPTION   : CIR SECOND FACTOR
'LIBRARY       : BOND
'GROUP         : CIR
'ID            : 007
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 21/01/2009
'**********************************************************************************
'**********************************************************************************

Function CIR_B_FACT_FUNC(ByVal PULL_BACK As Double, _
ByVal SIGMA As Double, _
ByVal TENOR As Double)

Dim GAMMA As Double

On Error GoTo ERROR_LABEL

GAMMA = CIR_GAMMA_FUNC(PULL_BACK, SIGMA)

CIR_B_FACT_FUNC = 2 * (Exp(GAMMA * TENOR) - 1) / _
    ((GAMMA + PULL_BACK) * (Exp(GAMMA * TENOR) - 1) + 2 * GAMMA)

Exit Function
ERROR_LABEL:
CIR_B_FACT_FUNC = Err.number
End Function


'**********************************************************************************
'**********************************************************************************
'FUNCTION      : CIR_INFIN_RATE_FUNC
'DESCRIPTION   : CIR LONG RATE FUNCTION
'LIBRARY       : BOND
'GROUP         : CIR
'ID            : 008
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 21/01/2009
'**********************************************************************************
'**********************************************************************************

Function CIR_INFIN_RATE_FUNC(ByVal EQUILIBRIUM As Double, _
ByVal PULL_BACK As Double, _
ByVal SIGMA As Double)

Dim GAMMA As Double

On Error GoTo ERROR_LABEL

GAMMA = CIR_GAMMA_FUNC(PULL_BACK, SIGMA)

CIR_INFIN_RATE_FUNC = (2 * PULL_BACK * EQUILIBRIUM) / (GAMMA + PULL_BACK)

Exit Function
ERROR_LABEL:
CIR_INFIN_RATE_FUNC = Err.number
End Function


'**********************************************************************************
'**********************************************************************************
'FUNCTION      : CIR_ZERO_RATE_FUNC
'DESCRIPTION   : CIR ZERO RATE FUNCTION
'LIBRARY       : BOND
'GROUP         : CIR
'ID            : 009
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 21/01/2009
'**********************************************************************************
'**********************************************************************************

Function CIR_ZERO_RATE_FUNC(ByVal SHORT_RATE As Double, _
ByVal EQUILIBRIUM As Double, _
ByVal PULL_BACK As Double, _
ByVal SIGMA As Double, _
ByVal TENOR As Double)

Dim ATEMP_VAL As Double
Dim BTEMP_VAL As Double

On Error GoTo ERROR_LABEL

ATEMP_VAL = CIR_A_FACT_FUNC(EQUILIBRIUM, PULL_BACK, SIGMA, TENOR)
BTEMP_VAL = CIR_B_FACT_FUNC(PULL_BACK, SIGMA, TENOR)

CIR_ZERO_RATE_FUNC = Log(1 / (ATEMP_VAL * Exp(-BTEMP_VAL * SHORT_RATE))) / TENOR

Exit Function
ERROR_LABEL:
CIR_ZERO_RATE_FUNC = Err.number
End Function


'**********************************************************************************
'**********************************************************************************
'FUNCTION      : CIR_INSTAN_SIGM_FUNC
'DESCRIPTION   : CIR INSTANTANEOUS SIGMA FUNCTION
'LIBRARY       : BOND
'GROUP         : CIR
'ID            : 010
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 21/01/2009
'**********************************************************************************
'**********************************************************************************

Function CIR_INSTAN_SIGM_FUNC(ByVal SHORT_RATE As Double, _
ByVal EQUILIBRIUM As Double, _
ByVal PULL_BACK As Double, _
ByVal SIGMA As Double, _
ByVal TENOR As Double)

Dim BTEMP_VAL As Double

On Error GoTo ERROR_LABEL

BTEMP_VAL = CIR_B_FACT_FUNC(PULL_BACK, SIGMA, TENOR)

If TENOR > 0 Then
    CIR_INSTAN_SIGM_FUNC = SIGMA * (SHORT_RATE ^ 0.5) * (BTEMP_VAL / TENOR)
Else
    CIR_INSTAN_SIGM_FUNC = 1
End If

Exit Function
ERROR_LABEL:
CIR_INSTAN_SIGM_FUNC = Err.number
End Function

'**********************************************************************************
'**********************************************************************************
'FUNCTION      : CIR_DISC_FUNC
'DESCRIPTION   : CIR_DISC_FUNC FACTOR FUNCTION
'LIBRARY       : BOND
'GROUP         : CIR
'ID            : 011
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 21/01/2009
'**********************************************************************************
'**********************************************************************************

Function CIR_DISC_FUNC(ByVal SHORT_RATE As Double, _
ByVal EQUILIBRIUM As Double, _
ByVal PULL_BACK As Double, _
ByVal SIGMA As Double, _
ByVal TENOR As Double)

Dim ATEMP_VAL As Double
Dim BTEMP_VAL As Double

On Error GoTo ERROR_LABEL

ATEMP_VAL = CIR_A_FACT_FUNC(EQUILIBRIUM, PULL_BACK, SIGMA, TENOR)
BTEMP_VAL = CIR_B_FACT_FUNC(PULL_BACK, SIGMA, TENOR)

CIR_DISC_FUNC = ATEMP_VAL * Exp(-BTEMP_VAL * SHORT_RATE)

Exit Function
ERROR_LABEL:
CIR_DISC_FUNC = Err.number
End Function

'**********************************************************************************
'**********************************************************************************
'FUNCTION      : CIR_ZERO_PROB_FUNC
'DESCRIPTION   : Long-term distribution of Spot Rate (Steady State
'Probability Density Function)
'LIBRARY       : BOND
'GROUP         : CIR
'ID            : 012
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 21/01/2009
'**********************************************************************************
'**********************************************************************************

Function CIR_ZERO_PROB_FUNC(ByVal MEAN_RATE As Double, _
ByVal EQUILIBRIUM As Double, _
ByVal PULL_BACK As Double, _
ByVal SIGMA As Double)

Dim ALPHA As Double

On Error GoTo ERROR_LABEL

ALPHA = 2 * PULL_BACK * EQUILIBRIUM / SIGMA ^ 2
CIR_ZERO_PROB_FUNC = GAMMA_DIST_FUNC(MEAN_RATE, ALPHA, SIGMA ^ 2 / (2 * PULL_BACK), False)

Exit Function
ERROR_LABEL:
CIR_ZERO_PROB_FUNC = Err.number
End Function


'**********************************************************************************
'**********************************************************************************
'FUNCTION      : CIR_DISCR_FUNC
'DESCRIPTION   : Parameters for the Cox, Ingersoll, Ross Model (discrete version)
'LIBRARY       : BOND
'GROUP         : CIR
'ID            : 013
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 21/01/2009
'**********************************************************************************
'**********************************************************************************

Function CIR_DISCR_FUNC(ByVal DELTA_TENOR As Double, _
ByVal SPOT_RATE As Double, _
ByVal EQUILIBRIUM As Double, _
ByVal PULL_BACK As Double, _
ByVal SIGMA As Double, _
ByVal NORMAL_RANDOM_VAL As Double)

'Because the volatility is proportional to the square root of r,
'r cannot become negative. As the rates increase, their volatility
'increases. At the same time, the model has the same mean-reverting
'or "pull-back" properties as the Vasicek model.

On Error GoTo ERROR_LABEL

CIR_DISCR_FUNC = PULL_BACK * (EQUILIBRIUM - SPOT_RATE) * DELTA_TENOR + _
SIGMA * (SPOT_RATE ^ 0.5) * NORMAL_RANDOM_VAL * (DELTA_TENOR ^ 0.5) + SPOT_RATE

Exit Function
ERROR_LABEL:
CIR_DISCR_FUNC = Err.number
End Function


'**********************************************************************************
'**********************************************************************************
'FUNCTION      : CIR_SIGMA_MATCH_FUNC
'DESCRIPTION   : Divided by Sqr of Short Rate to make volatility of both models
'comparable to the vasicek
'LIBRARY       : BOND
'GROUP         : CIR
'ID            : 014
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 21/01/2009
'**********************************************************************************
'**********************************************************************************


Function CIR_SIGMA_MATCH_FUNC(ByVal SHORT_RATE As Double, _
ByVal SIGMA_VASICEK As Double)

On Error GoTo ERROR_LABEL

CIR_SIGMA_MATCH_FUNC = SIGMA_VASICEK / SHORT_RATE ^ 0.5

Exit Function
ERROR_LABEL:
CIR_SIGMA_MATCH_FUNC = Err.number
End Function

'----------------------------------------------------------------------------------
'Hull, John C., Options, Futures & Other Derivatives. Fourth edition (2000).
'Prentice-Hall. p. 570.
'-------------------------------------Model----------------------------------------
'  Cox, C.J. Ingersoll, J.E. Ross, S.A. (1985) . "A Theory of the Term Structure
'  of Interest Rates". Econometrica, 53 (1985), p. 385-407

'  Steady-state probability density function formula for Vasicek model from Wilmott,
'  Paul. Paul Wilmott on Quantitative Finance, Volume 2, p. 563, John Wiley 2000.
 
'  Formula infinitely long rate, volatility zero rates from Jackson, M. Staunton, M.
'  "Advanced Modelling in Finance using Excel and VBA", Wiley Finance (2001). p. 238.
'----------------------------------------------------------------------------------
