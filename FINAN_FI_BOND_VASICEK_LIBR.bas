Attribute VB_Name = "FINAN_FI_BOND_VASICEK_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.


'**********************************************************************************
'**********************************************************************************
'FUNCTION      : VASICEK_RATE_FUNC
'DESCRIPTION   : Calculate VASICEK's zero discount function,
'i.e. present value factor or price of a zero bond with maturity
'LIBRARY       : BOND
'GROUP         : VASICEK
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 21/01/2009
'**********************************************************************************
'**********************************************************************************


Function VASICEK_RATE_FUNC(ByVal SHORT_RATE As Double, _
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

ReDim TEMP_VECTOR(1 To 7, 1 To 2)

'-----------------------------------------------------------------------------

TEMP_VECTOR(2, 1) = ("B IN VASICEK MODEL")
TEMP_VECTOR(3, 1) = ("A IN VASICEK MODEL")
TEMP_VECTOR(4, 1) = ("INFINITELY-LONG RATE")
TEMP_VECTOR(5, 1) = ("VASICEK ZERO RATE")
TEMP_VECTOR(6, 1) = ("VASICEK SIGMA OF ZERO RATE")
TEMP_VECTOR(7, 1) = ("VASICEK DISCOUNT FACTOR")

TEMP_VECTOR(2, 2) = VASICEK_B_FACT_FUNC(PULL_BACK, TENOR)
TEMP_VECTOR(3, 2) = VASICEK_A_FACT_FUNC(EQUILIBRIUM, PULL_BACK, SIGMA, TENOR)

'-----------------------------------------------------------------------------

TEMP_VECTOR(4, 2) = VASICEK_INFIN_RATE_FUNC(EQUILIBRIUM, PULL_BACK, SIGMA)
TEMP_VECTOR(5, 2) = VASICEK_ZERO_RATE_FUNC(SHORT_RATE, EQUILIBRIUM, PULL_BACK, SIGMA, TENOR)
TEMP_VECTOR(6, 2) = VASICEK_INSTAN_SIGM_FUNC(PULL_BACK, SIGMA, TENOR)
TEMP_VECTOR(7, 2) = VASICEK_DISC_FUNC(SHORT_RATE, EQUILIBRIUM, PULL_BACK, SIGMA, TENOR)

'-----------------------------------------------------------------------------
MEAN_RATE = TEMP_VECTOR(6, 2)

TEMP_VECTOR(1, 1) = ("STEADY STATE PROBABILITY DENSITY FUNCTION: " & Format(MEAN_RATE, "0.0%") & "")
TEMP_VECTOR(1, 2) = VASICEK_ZERO_PROB_FUNC(MEAN_RATE, EQUILIBRIUM, PULL_BACK, SIGMA)

Select Case OUTPUT
    Case 0
        VASICEK_RATE_FUNC = TEMP_VECTOR
    Case Else
        VASICEK_RATE_FUNC = MATRIX_GET_COLUMN_FUNC(TEMP_VECTOR, 2, 1)
End Select

Exit Function
ERROR_LABEL:
VASICEK_RATE_FUNC = Err.number
End Function


'**********************************************************************************
'**********************************************************************************
'FUNCTION      : VASICEK_YIELD_FUNC
'DESCRIPTION   : Term structure in Vasicek Model
'LIBRARY       : BOND
'GROUP         : VASICEK
'ID            : 002
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 21/01/2009
'**********************************************************************************
'**********************************************************************************

'PULL_BACK ---> kappa
'EQUILIBRIUM --> thetha

Function VASICEK_YIELD_FUNC(ByVal SETTLEMENT As Date, _
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
      TEMP_MATRIX(i, 3) = VASICEK_DISC_FUNC(SHORT_RATE, EQUILIBRIUM, PULL_BACK, SIGMA, TEMP_TENOR)
      TEMP_MATRIX(i, 4) = VASICEK_ZERO_RATE_FUNC(SHORT_RATE, EQUILIBRIUM, PULL_BACK, SIGMA, TEMP_TENOR)
 Next i

Select Case OUTPUT
Case 0
    VASICEK_YIELD_FUNC = TEMP_MATRIX
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
    TEMP_VECTOR(7, 2) = VASICEK_INFIN_RATE_FUNC(EQUILIBRIUM, PULL_BACK, SIGMA)
    TEMP_VECTOR(8, 2) = TEMP_VECTOR(7, 2)
 
    VASICEK_YIELD_FUNC = TEMP_VECTOR
End Select

Exit Function
ERROR_LABEL:
VASICEK_YIELD_FUNC = Err.number
End Function

'**********************************************************************************
'**********************************************************************************



'**********************************************************************************
'**********************************************************************************
'FUNCTION      : VASICEK_BOND_OPT_FUNC
'DESCRIPTION   : VASICEK BOND OPTION MODEL
'LIBRARY       : BOND
'GROUP         : VASICEK
'ID            : 003
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 21/01/2009
'**********************************************************************************
'**********************************************************************************

Function VASICEK_BOND_OPT_FUNC(ByVal FACE As Double, _
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

FIRST_COMPONENT = VASICEK_DISC_FUNC(SHORT_RATE, EQUILIBRIUM, _
PULL_BACK, SIGMA, OPT_TENOR)

SECOND_COMPONENT = VASICEK_DISC_FUNC(SHORT_RATE, EQUILIBRIUM, _
PULL_BACK, SIGMA, BOND_TENOR)

BTEMP_VAL = Sqr(SIGMA ^ 2 * (1 - Exp(-2 * PULL_BACK * OPT_TENOR)) / (2 * PULL_BACK)) _
* (1 - Exp(-PULL_BACK * (BOND_TENOR - OPT_TENOR))) / PULL_BACK

ATEMP_VAL = 1 / BTEMP_VAL * Log(SECOND_COMPONENT / (FIRST_COMPONENT * STRIKE)) + BTEMP_VAL / 2

Select Case OPTION_FLAG
Case 1 ', "c"
    VASICEK_BOND_OPT_FUNC = FACE * (SECOND_COMPONENT * CND_FUNC(ATEMP_VAL) _
    - STRIKE * FIRST_COMPONENT * CND_FUNC(ATEMP_VAL - BTEMP_VAL))
Case Else
    VASICEK_BOND_OPT_FUNC = FACE * (STRIKE * FIRST_COMPONENT * _
    CND_FUNC(-ATEMP_VAL + BTEMP_VAL) - SECOND_COMPONENT * CND_FUNC(-ATEMP_VAL))
End Select

Exit Function
ERROR_LABEL:
VASICEK_BOND_OPT_FUNC = Err.number
End Function

'**********************************************************************************
'**********************************************************************************



'**********************************************************************************
'**********************************************************************************
'FUNCTION      : VASICEK_PROB_TBL_FUNC
'DESCRIPTION   : VASICEK PROB TABLE
'LIBRARY       : BOND
'GROUP         : VASICEK
'ID            : 004
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 21/01/2009
'**********************************************************************************
'**********************************************************************************

Function VASICEK_PROB_TBL_FUNC(ByVal EQUILIBRIUM As Double, _
ByVal PULL_BACK As Double, _
ByVal SIGMA As Double, _
Optional ByVal MIN_FACTOR As Double = -4.5, _
Optional ByVal MAX_FACTOR As Double = 4.5, _
Optional ByVal DELTA_FACTOR As Double = 0.5, _
Optional ByVal MULT_FACTOR As Double = 1.2, _
Optional ByVal OUTPUT As Integer = 0)

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
TEMP_MATRIX(1, 2) = EQUILIBRIUM + (SIGMA / Sqr(2 * PULL_BACK)) _
* TEMP_MATRIX(1, 1)
TEMP_MATRIX(1, 3) = VASICEK_ZERO_PROB_FUNC(TEMP_MATRIX(1, 2), EQUILIBRIUM, PULL_BACK, SIGMA)

For i = 2 To NROWS
    TEMP_MATRIX(i, 1) = TEMP_MATRIX(i - 1, 1) + DELTA_FACTOR
    TEMP_MATRIX(i, 2) = EQUILIBRIUM + (SIGMA / Sqr(2 * PULL_BACK)) _
    * TEMP_MATRIX(i, 1)
    TEMP_MATRIX(i, 3) = VASICEK_ZERO_PROB_FUNC(TEMP_MATRIX(i, 2), EQUILIBRIUM, PULL_BACK, SIGMA)
Next i

Select Case OUTPUT
Case 0
    VASICEK_PROB_TBL_FUNC = TEMP_MATRIX
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
    TEMP_VECTOR(5, 1) = EQUILIBRIUM + (SIGMA / Sqr(2 * PULL_BACK))
    TEMP_VECTOR(5, 2) = TEMP_VECTOR(2, 2)
    TEMP_VECTOR(6, 1) = TEMP_VECTOR(5, 1)
    TEMP_VECTOR(6, 2) = TEMP_VECTOR(3, 2)
    TEMP_VECTOR(7, 1) = ("- SIGMA")
    TEMP_VECTOR(7, 2) = ""
    TEMP_VECTOR(8, 1) = EQUILIBRIUM - (SIGMA / Sqr(2 * PULL_BACK))
    TEMP_VECTOR(8, 2) = TEMP_VECTOR(2, 2)
    TEMP_VECTOR(9, 1) = TEMP_VECTOR(8, 1)
    TEMP_VECTOR(9, 2) = TEMP_VECTOR(3, 2)

    VASICEK_PROB_TBL_FUNC = TEMP_VECTOR
End Select

Exit Function
ERROR_LABEL:
VASICEK_PROB_TBL_FUNC = Err.number
End Function

'**********************************************************************************
'**********************************************************************************



'**********************************************************************************
'**********************************************************************************
'FUNCTION      : VASICEK_A_FACT_FUNC
'DESCRIPTION   : VASICEK FIRST FACTOR
'LIBRARY       : BOND
'GROUP         : VASICEK
'ID            : 005
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 21/01/2009
'**********************************************************************************
'**********************************************************************************

Function VASICEK_A_FACT_FUNC(ByVal EQUILIBRIUM As Double, _
ByVal PULL_BACK As Double, _
ByVal SIGMA As Double, _
ByVal TENOR As Double)

Dim BTEMP_VAL As Double
Dim ATEMP_VAL As Double

Dim FIRST_TERM As Double
Dim SECOND_TERM As Double
Dim THIRD_TERM As Double

On Error GoTo ERROR_LABEL

BTEMP_VAL = VASICEK_B_FACT_FUNC(PULL_BACK, TENOR)
FIRST_TERM = BTEMP_VAL - TENOR

SECOND_TERM = (PULL_BACK ^ 2) * EQUILIBRIUM - (SIGMA ^ 2) / 2
THIRD_TERM = (SIGMA ^ 2 * BTEMP_VAL ^ 2)
ATEMP_VAL = Exp((FIRST_TERM * SECOND_TERM) / (PULL_BACK ^ 2) - _
THIRD_TERM / (4 * PULL_BACK))

VASICEK_A_FACT_FUNC = Exp((FIRST_TERM * SECOND_TERM) / (PULL_BACK ^ 2) _
- THIRD_TERM / (4 * PULL_BACK))

Exit Function
ERROR_LABEL:
VASICEK_A_FACT_FUNC = Err.number
End Function

'**********************************************************************************
'**********************************************************************************



'**********************************************************************************
'**********************************************************************************
'FUNCTION      : VASICEK_B_FACT_FUNC
'DESCRIPTION   : VASICEK SECOND FACTOR
'LIBRARY       : BOND
'GROUP         : VASICEK
'ID            : 006
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 21/01/2009
'**********************************************************************************
'**********************************************************************************

Function VASICEK_B_FACT_FUNC(ByVal PULL_BACK As Double, _
ByVal TENOR As Double)

On Error GoTo ERROR_LABEL

VASICEK_B_FACT_FUNC = (1 - Exp(-PULL_BACK * TENOR)) / PULL_BACK

Exit Function
ERROR_LABEL:
VASICEK_B_FACT_FUNC = Err.number
End Function

'**********************************************************************************
'**********************************************************************************



'**********************************************************************************
'**********************************************************************************
'FUNCTION      : VASICEK_INFIN_RATE_FUNC
'DESCRIPTION   : VASICEK LONG RATE
'LIBRARY       : BOND
'GROUP         : VASICEK
'ID            : 007
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 21/01/2009
'**********************************************************************************
'**********************************************************************************

Function VASICEK_INFIN_RATE_FUNC(ByVal EQUILIBRIUM As Double, _
ByVal PULL_BACK As Double, _
ByVal SIGMA As Double)

On Error GoTo ERROR_LABEL

VASICEK_INFIN_RATE_FUNC = EQUILIBRIUM - (SIGMA ^ 2 / (2 * PULL_BACK ^ 2))

Exit Function
ERROR_LABEL:
VASICEK_INFIN_RATE_FUNC = Err.number
End Function

'**********************************************************************************
'**********************************************************************************



'**********************************************************************************
'**********************************************************************************
'FUNCTION      : VASICEK_ZERO_RATE_FUNC
'DESCRIPTION   : VASICEK ZERO RATE FUNCTION
'LIBRARY       : BOND
'GROUP         : VASICEK
'ID            : 008
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 21/01/2009
'**********************************************************************************
'**********************************************************************************

Function VASICEK_ZERO_RATE_FUNC(ByVal SHORT_RATE As Double, _
ByVal EQUILIBRIUM As Double, _
ByVal PULL_BACK As Double, _
ByVal SIGMA As Double, _
ByVal TENOR As Double)

Dim ATEMP_VAL As Double
Dim BTEMP_VAL As Double

On Error GoTo ERROR_LABEL

ATEMP_VAL = VASICEK_A_FACT_FUNC(EQUILIBRIUM, PULL_BACK, SIGMA, TENOR)
BTEMP_VAL = VASICEK_B_FACT_FUNC(PULL_BACK, TENOR)

VASICEK_ZERO_RATE_FUNC = Log(1 / (ATEMP_VAL * Exp(-BTEMP_VAL * SHORT_RATE))) / TENOR

Exit Function
ERROR_LABEL:
VASICEK_ZERO_RATE_FUNC = Err.number
End Function

'**********************************************************************************
'**********************************************************************************



'**********************************************************************************
'**********************************************************************************
'FUNCTION      : VASICEK_INSTAN_SIGM_FUNC
'DESCRIPTION   : VASICEK INSTANTANEOUS SIGMA FUNCTION
'LIBRARY       : BOND
'GROUP         : VASICEK
'ID            : 009
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 21/01/2009
'**********************************************************************************
'**********************************************************************************


Function VASICEK_INSTAN_SIGM_FUNC(ByVal PULL_BACK As Double, _
ByVal SIGMA As Double, _
ByVal TENOR As Double)

Dim BTEMP_VAL As Double

On Error GoTo ERROR_LABEL

BTEMP_VAL = VASICEK_B_FACT_FUNC(PULL_BACK, TENOR)

If TENOR > 0 Then
    VASICEK_INSTAN_SIGM_FUNC = SIGMA * (BTEMP_VAL / TENOR)
Else
    VASICEK_INSTAN_SIGM_FUNC = 1
End If

Exit Function
ERROR_LABEL:
VASICEK_INSTAN_SIGM_FUNC = Err.number
End Function

'**********************************************************************************
'**********************************************************************************



'**********************************************************************************
'**********************************************************************************
'FUNCTION      : VASICEK_DISC_FUNC
'DESCRIPTION   : VASICEK DISCOUNT FUNCTION
'LIBRARY       : BOND
'GROUP         : VASICEK
'ID            : 010
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 21/01/2009
'**********************************************************************************
'**********************************************************************************


Function VASICEK_DISC_FUNC(ByVal SHORT_RATE As Double, _
ByVal EQUILIBRIUM As Double, _
ByVal PULL_BACK As Double, _
ByVal SIGMA As Double, _
ByVal TENOR As Double)

Dim ATEMP_VAL As Double
Dim BTEMP_VAL As Double

On Error GoTo ERROR_LABEL

ATEMP_VAL = VASICEK_A_FACT_FUNC(EQUILIBRIUM, PULL_BACK, SIGMA, TENOR)
BTEMP_VAL = VASICEK_B_FACT_FUNC(PULL_BACK, TENOR)

VASICEK_DISC_FUNC = ATEMP_VAL * Exp(-BTEMP_VAL * SHORT_RATE)

Exit Function
ERROR_LABEL:
VASICEK_DISC_FUNC = Err.number
End Function

'**********************************************************************************
'**********************************************************************************



'**********************************************************************************
'**********************************************************************************
'FUNCTION      : VASICEK_ZERO_PROB_FUNC
'DESCRIPTION   : Long-term distribution of Spot Rate (Steady State
'Probability Density Function)
'LIBRARY       : BOND
'GROUP         : VASICEK
'ID            : 011
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 21/01/2009
'**********************************************************************************
'**********************************************************************************

Function VASICEK_ZERO_PROB_FUNC(ByVal MEAN_RATE As Double, _
ByVal EQUILIBRIUM As Double, _
ByVal PULL_BACK As Double, _
ByVal SIGMA As Double)

On Error GoTo ERROR_LABEL

    VASICEK_ZERO_PROB_FUNC = NORMDIST_FUNC(MEAN_RATE, EQUILIBRIUM, _
    (SIGMA / Sqr(2 * PULL_BACK)), 1)

Exit Function
ERROR_LABEL:
VASICEK_ZERO_PROB_FUNC = Err.number
End Function


'**********************************************************************************
'**********************************************************************************



'**********************************************************************************
'**********************************************************************************
'FUNCTION      : VASICEK_DISCR_FUNC
'DESCRIPTION   : DISCRETE ZERO VALUE
'LIBRARY       : BOND
'GROUP         : VASICEK
'ID            : 012
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 21/01/2009
'**********************************************************************************
'**********************************************************************************

Function VASICEK_DISCR_FUNC(ByVal DELTA_TENOR As Double, _
ByVal SPOT_RATE As Double, _
ByVal EQUILIBRIUM As Double, _
ByVal PULL_BACK As Double, _
ByVal SIGMA As Double, _
ByVal NORMAL_RANDOM_VAL As Double)

On Error GoTo ERROR_LABEL

'Parameters for the Vasicek (discrete version):

' : "strength" at which r is pulled back to g
' : long-term equilibrium of short-term rates
' : volatility superimposed(annualized)
' : is a random drawing from a standardized normal distribution, F(0,1)

VASICEK_DISCR_FUNC = PULL_BACK * (EQUILIBRIUM - SPOT_RATE) * (DELTA_TENOR) + _
SIGMA * NORMAL_RANDOM_VAL * (DELTA_TENOR) ^ 0.5 + SPOT_RATE

Exit Function
ERROR_LABEL:
VASICEK_DISCR_FUNC = Err.number
End Function

'---------------------------------------------------------------------------------------
'The Black 's model for pricing interest rate options such as bond options,
'interest rate caps/floors, and swap options make the assumption that the
'probability distribution of an interest rate, a bond price, or some other
'variable at a future point in time is lognormal. They are widely used for
'valuing instruments such as caps, European bond options, and European swap
'options. These models do not provide a description of the stochastic behaviour
'of interest rates and bond prices. Black's model is concerned only with
'describing the probability distribution of a single variable at a single point
'in time Consequently, they cannot be used for valuing interest rate derivatives
'such as American-style swap options, callable bonds, and structured notes.This
'note discusses an alternative approach to overcome these limitations.
'---------------------------------------------------------------------------------------
'----------------------Term structure in Vasicek Model----------------------------------
'---------------------------------------------------------------------------------------

'Hull, John C., Options, Futures & Other Derivatives. Fourth edition (2000).
'Prentice-Hall. p. 567.

'Model:

'Vasicek, O. 1977 "An Equilibrium Characterization of the term structure." Journal
'of Financial Economics 5: 177-188.

'Steady-state probability density function formula for Vasicek model from Wilmott,
'Paul. Paul Wilmott on Quantitative Finance, Volume 2, p. 563, John Wiley 2000.

'Formula infinitely long rate from Holden, Craig W. Spreadsheet Modeling in
'Investments. Prentice Hall. 2002 edition. p.49

'---------------------------------------------------------------------------------------
