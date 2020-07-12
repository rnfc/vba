Attribute VB_Name = "FINAN_FI_NS_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.

'**********************************************************************************
'**********************************************************************************
'FUNCTION      : NS_SVENSSON_YIELD_FRAME_FUNC
'DESCRIPTION   : Function to calculate NELSON zero discount function,
'i.e. present value factor or price of a zero bond with maturity
'LIBRARY       : FIXED_INCOME
'GROUP         : NS_YIELD
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 10-06-2008
'**********************************************************************************
'**********************************************************************************

Function NS_SVENSSON_YIELD_FRAME_FUNC(ByVal TENOR As Double, _
ByVal BETA0_VAL As Double, _
ByVal BETA1_VAL As Double, _
ByVal BETA2_VAL As Double, _
ByVal BETA3_VAL As Double, _
ByVal TAU1_VAL As Double, _
ByVal TAU2_VAL As Double, _
Optional ByVal DELTA_TENOR_VAL As Double = 1, _
Optional ByVal OUTPUT As Integer = 0)

Dim TEMP_MATRIX As Variant

On Error GoTo ERROR_LABEL

ReDim TEMP_MATRIX(1 To 9, 1 To 2)

TEMP_MATRIX(1, 1) = "FIRST_COMPONENT"
TEMP_MATRIX(1, 2) = NS_SVENSSON_COMPONENT_FUNC(TENOR, BETA0_VAL, BETA1_VAL, BETA2_VAL, BETA3_VAL, TAU1_VAL, TAU2_VAL, 0)
'-----------------------------------------------------------------------------
TEMP_MATRIX(2, 1) = "SECOND_COMPONENT"
TEMP_MATRIX(2, 2) = NS_SVENSSON_COMPONENT_FUNC(TENOR, BETA0_VAL, BETA1_VAL, BETA2_VAL, BETA3_VAL, TAU1_VAL, TAU2_VAL, 1)

TEMP_MATRIX(3, 1) = "THIRD_COMPONENT"
TEMP_MATRIX(3, 2) = NS_SVENSSON_COMPONENT_FUNC(TENOR, BETA0_VAL, BETA1_VAL, BETA2_VAL, BETA3_VAL, TAU1_VAL, TAU2_VAL, 2)

TEMP_MATRIX(4, 1) = "FORTH_COMPONENT"
TEMP_MATRIX(4, 2) = NS_SVENSSON_COMPONENT_FUNC(TENOR, BETA0_VAL, BETA1_VAL, BETA2_VAL, BETA3_VAL, TAU1_VAL, TAU2_VAL, 3)
'---------------------------------------------------------------------------------------
TEMP_MATRIX(5, 1) = "YIELD"
TEMP_MATRIX(5, 2) = NS_SVENSSON_YIELD_FUNC(TENOR, BETA0_VAL, BETA1_VAL, BETA2_VAL, BETA3_VAL, TAU1_VAL, TAU2_VAL)
TEMP_MATRIX(6, 1) = "DISCOUNT_FACTOR"
TEMP_MATRIX(6, 2) = Exp(-1 * TENOR * TEMP_MATRIX(5, 2))
'----------------------------------------------------------------------------------------
TEMP_MATRIX(7, 1) = "INITIAL YIELD"
TEMP_MATRIX(7, 2) = NS_SVENSSON_SMOOTHING_FUNC(DELTA_TENOR_VAL, BETA0_VAL, BETA1_VAL, BETA2_VAL, BETA3_VAL, TAU1_VAL, TAU2_VAL)
TEMP_MATRIX(8, 1) = "FORWARD YIELD"
TEMP_MATRIX(8, 2) = NS_SVENSSON_FORWARD_FUNC(DELTA_TENOR_VAL, BETA0_VAL, BETA1_VAL, BETA2_VAL, BETA3_VAL, TAU1_VAL, TAU2_VAL)
TEMP_MATRIX(9, 1) = "FORWARD DERIVATIVE"
TEMP_MATRIX(9, 2) = NS_SVENSSON_GRADIENT_FUNC(DELTA_TENOR_VAL, BETA0_VAL, BETA1_VAL, BETA2_VAL, BETA3_VAL, TAU1_VAL, TAU2_VAL)


Select Case OUTPUT
Case 0
    NS_SVENSSON_YIELD_FRAME_FUNC = TEMP_MATRIX
Case Else
    NS_SVENSSON_YIELD_FRAME_FUNC = MATRIX_GET_COLUMN_FUNC(TEMP_MATRIX, 2, 1)
End Select

Exit Function
ERROR_LABEL:
NS_SVENSSON_YIELD_FRAME_FUNC = Err.number
End Function


'**********************************************************************************
'**********************************************************************************
'FUNCTION      : NS_SVENSSON_YIELD_FUNC
'DESCRIPTION   : Function the extension according to Nelson Siegel &
'Svensson model
'LIBRARY       : FIXED_INCOME
'GROUP         : NS_YIELD
'ID            : 002
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 10-06-2008
'**********************************************************************************
'**********************************************************************************

Function NS_SVENSSON_YIELD_FUNC(ByVal TENOR As Double, _
ByVal BETA0_VAL As Double, _
ByVal BETA1_VAL As Double, _
ByVal BETA2_VAL As Double, _
ByVal BETA3_VAL As Double, _
ByVal TAU1_VAL As Double, _
ByVal TAU2_VAL As Double)

On Error GoTo ERROR_LABEL

If TAU2_VAL = 0 Then: TAU2_VAL = TAU1_VAL
'for basic Nelson Siegel TAU2_VAL = TAU1_VAL
If TENOR = 0 Then: TENOR = 0.000000001

If BETA3_VAL = 0 Then 'Extended Nelson & Siegel Spot Rate Model
    NS_SVENSSON_YIELD_FUNC = BETA0_VAL + BETA1_VAL * ((1 - Exp(-TENOR / TAU1_VAL)) / TENOR * TAU1_VAL) + BETA2_VAL * (((1 - Exp(-TENOR / TAU2_VAL)) / TENOR * TAU2_VAL) - Exp(-TENOR / TAU2_VAL))
Else 'Extended Nelson & Siegel Spot Rate Model with Svensson 1994 Extension
    NS_SVENSSON_YIELD_FUNC = BETA0_VAL + BETA1_VAL * ((1 - Exp(-TENOR / TAU1_VAL)) / TENOR * TAU1_VAL) + BETA2_VAL * (((1 - Exp(-TENOR / TAU1_VAL)) / TENOR * TAU1_VAL) - Exp(-TENOR / TAU1_VAL)) + BETA3_VAL * (((1 - Exp(-TENOR / TAU2_VAL)) / TENOR * TAU2_VAL) - Exp(-TENOR / TAU2_VAL))
End If

Exit Function
ERROR_LABEL:
NS_SVENSSON_YIELD_FUNC = Err.number
End Function


'**********************************************************************************
'**********************************************************************************
'FUNCTION      : NS_SVENSSON_COMPONENT_FUNC
'DESCRIPTION   : Nelson & Svensson Component Calculator
'LIBRARY       : FIXED_INCOME
'GROUP         : NS_YIELD
'ID            : 003
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 10-06-2008
'**********************************************************************************
'**********************************************************************************

Function NS_SVENSSON_COMPONENT_FUNC(ByVal TENOR As Double, _
ByVal BETA0_VAL As Double, _
ByVal BETA1_VAL As Double, _
ByVal BETA2_VAL As Double, _
ByVal BETA3_VAL As Double, _
ByVal TAU1_VAL As Double, _
ByVal TAU2_VAL As Double, _
Optional ByVal OUTPUT As Integer = 0)

On Error GoTo ERROR_LABEL

If TAU2_VAL = 0 Then: TAU2_VAL = TAU1_VAL  'for basic Nelson Siegel TAU2_VAL = TAU1_VAL
If TENOR = 0 Then: TENOR = 0.000000001

Select Case OUTPUT
Case 0 'FIRST COMPONENT
    NS_SVENSSON_COMPONENT_FUNC = BETA0_VAL
Case 1  'SECOND COMPONENT
    NS_SVENSSON_COMPONENT_FUNC = BETA1_VAL * ((1 - Exp(-TENOR / TAU1_VAL)) / TENOR * TAU1_VAL)
Case 2 'THIRD COMPONENT
    If BETA3_VAL = 0 Then
        NS_SVENSSON_COMPONENT_FUNC = BETA2_VAL * (((1 - Exp(-TENOR / TAU2_VAL)) / TENOR * TAU2_VAL) - Exp(-TENOR / TAU2_VAL))
    Else
        NS_SVENSSON_COMPONENT_FUNC = BETA2_VAL * (((1 - Exp(-TENOR / TAU1_VAL)) / TENOR * TAU1_VAL) - Exp(-TENOR / TAU1_VAL))
    End If
Case Else 'FORTH COMPONENT with SVENSSON 1994 Extension
    NS_SVENSSON_COMPONENT_FUNC = BETA3_VAL * (((1 - Exp(-TENOR / TAU2_VAL)) / TENOR * TAU2_VAL) - Exp(-TENOR / TAU2_VAL))
End Select

Exit Function
ERROR_LABEL:
NS_SVENSSON_COMPONENT_FUNC = Err.number
End Function


'**********************************************************************************
'**********************************************************************************
'FUNCTION      : NS_SVENSSON_SMOOTHING_FUNC
'DESCRIPTION   : Nelson-Siegel smoothing of the initial yield curve
'LIBRARY       : FIXED_INCOME
'GROUP         : NS_YIELD
'ID            : 004
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 10-06-2008
'**********************************************************************************
'**********************************************************************************

Function NS_SVENSSON_SMOOTHING_FUNC(ByVal DELTA_TENOR_VAL As Double, _
ByVal BETA0_VAL As Double, _
ByVal BETA1_VAL As Double, _
ByVal BETA2_VAL As Double, _
ByVal BETA3_VAL As Double, _
ByVal TAU1_VAL As Double, _
ByVal TAU2_VAL As Double)

Dim TEMP1_VAL As Double
Dim TEMP2_VAL As Double

On Error GoTo ERROR_LABEL

If TAU2_VAL = 0 Then: TAU2_VAL = TAU1_VAL
'for basic Nelson Siegel TAU2_VAL = TAU1_VAL
If TAU1_VAL = 0 Then: TAU1_VAL = 0.0001

'Long-run levels of interest rates (BETA0_VAL)
'Short-run component (BETA1_VAL)
'Medium-term component (BETA2_VAL) -->
'determines magnitude and the direction of the hump
'Long-Term component (BETA3_VAL) --> optional parameter proposed by Svensson (1994)
'Decay parameter 1 --> determines decay of short-term component, must be > 0
'Decay parameter 2 --> determines decay of medium-term component, must be > 0
'Spot rate at time t

Select Case DELTA_TENOR_VAL
Case 0
  NS_SVENSSON_SMOOTHING_FUNC = BETA0_VAL + BETA1_VAL + BETA2_VAL
Case Else
  TEMP1_VAL = Exp(-DELTA_TENOR_VAL / TAU1_VAL)
  TEMP2_VAL = Exp(-DELTA_TENOR_VAL / TAU2_VAL)
  NS_SVENSSON_SMOOTHING_FUNC = BETA0_VAL + BETA1_VAL * TEMP1_VAL + BETA2_VAL * (1 - TEMP1_VAL) / (DELTA_TENOR_VAL / TAU1_VAL) + BETA3_VAL * ((1 - TEMP2_VAL) / (DELTA_TENOR_VAL / TAU2_VAL) - TEMP2_VAL)
End Select

Exit Function
ERROR_LABEL:
NS_SVENSSON_SMOOTHING_FUNC = Err.number
End Function


'**********************************************************************************
'**********************************************************************************
'FUNCTION      : NS_SVENSSON_FORWARD_FUNC
'DESCRIPTION   : Calculate the process followed by the yield form the process
'from the parameters of the Nelson Siegal Model
'LIBRARY       : FIXED_INCOME
'GROUP         : NS_YIELD
'ID            : 005
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 10-06-2008
'**********************************************************************************
'**********************************************************************************

Function NS_SVENSSON_FORWARD_FUNC(ByVal DELTA_TENOR_VAL As Double, _
ByVal BETA0_VAL As Double, _
ByVal BETA1_VAL As Double, _
ByVal BETA2_VAL As Double, _
ByVal BETA3_VAL As Double, _
ByVal TAU1_VAL As Double, _
ByVal TAU2_VAL As Double)

Dim TEMP1_VAL As Double
Dim TEMP2_VAL As Double
Dim epsilon As Double

On Error GoTo ERROR_LABEL

epsilon = 0.0001
If TAU2_VAL = 0 Then: TAU2_VAL = TAU1_VAL
'for basic Nelson Siegel TAU2_VAL = TAU1_VAL

If (DELTA_TENOR_VAL > 0) Then
    TEMP1_VAL = DELTA_TENOR_VAL * (1 + epsilon) * NS_SVENSSON_SMOOTHING_FUNC(DELTA_TENOR_VAL * (1 + epsilon), BETA0_VAL, BETA1_VAL, BETA2_VAL, BETA3_VAL, TAU1_VAL, TAU2_VAL)
    TEMP2_VAL = DELTA_TENOR_VAL * NS_SVENSSON_SMOOTHING_FUNC(DELTA_TENOR_VAL, BETA0_VAL, BETA1_VAL, BETA2_VAL, BETA3_VAL, TAU1_VAL, TAU2_VAL)
    NS_SVENSSON_FORWARD_FUNC = (TEMP1_VAL - TEMP2_VAL) / (epsilon * DELTA_TENOR_VAL)
Else
    NS_SVENSSON_FORWARD_FUNC = BETA0_VAL + BETA1_VAL + BETA2_VAL
End If

Exit Function
ERROR_LABEL:
NS_SVENSSON_FORWARD_FUNC = Err.number
End Function


'**********************************************************************************
'**********************************************************************************
'FUNCTION      : NS_SVENSSON_GRADIENT_FUNC
'DESCRIPTION   : Grad Function of the Nelson Svensson Model
'LIBRARY       : FIXED_INCOME
'GROUP         : NS_YIELD
'ID            : 006
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 10-06-2008
'**********************************************************************************
'**********************************************************************************


Function NS_SVENSSON_GRADIENT_FUNC(ByVal DELTA_TENOR_VAL As Double, _
ByVal BETA0_VAL As Double, _
ByVal BETA1_VAL As Double, _
ByVal BETA2_VAL As Double, _
ByVal BETA3_VAL As Double, _
ByVal TAU1_VAL As Double, _
ByVal TAU2_VAL As Double)

Dim TEMP1_VAL As Double
Dim TEMP2_VAL As Double
Dim epsilon As Double

'Long-run levels of interest rates (BETA0_VAL)
'Short-run component (BETA1_VAL)
'Medium-term component (BETA2_VAL) -->
'determines magnitude and the direction of the hump
'Long-Term component (BETA3_VAL) --> optional parameter proposed by Svensson (1994)
'Decay parameter 1 --> determines decay of short-term component, must be > 0
'Decay parameter 2 --> determines decay of medium-term component, must be > 0
'Spot rate at time t

On Error GoTo ERROR_LABEL

epsilon = 0.0001
If DELTA_TENOR_VAL > 0 Then
  TEMP1_VAL = NS_SVENSSON_FORWARD_FUNC(DELTA_TENOR_VAL * (1 + epsilon), BETA0_VAL, BETA1_VAL, BETA2_VAL, BETA3_VAL, TAU1_VAL, TAU2_VAL)
  TEMP2_VAL = NS_SVENSSON_FORWARD_FUNC(DELTA_TENOR_VAL, BETA0_VAL, BETA1_VAL, BETA2_VAL, BETA3_VAL, TAU1_VAL, TAU2_VAL)
  NS_SVENSSON_GRADIENT_FUNC = (TEMP1_VAL - TEMP2_VAL) / (epsilon * DELTA_TENOR_VAL)
Else
  TEMP1_VAL = NS_SVENSSON_FORWARD_FUNC(epsilon, BETA0_VAL, BETA1_VAL, BETA2_VAL, BETA3_VAL, TAU1_VAL, TAU2_VAL)
  TEMP2_VAL = NS_SVENSSON_FORWARD_FUNC(0, BETA0_VAL, BETA1_VAL, BETA2_VAL, BETA3_VAL, TAU1_VAL, TAU2_VAL)
  NS_SVENSSON_GRADIENT_FUNC = (TEMP1_VAL - TEMP2_VAL) / (epsilon)
End If

Exit Function
ERROR_LABEL:
NS_SVENSSON_GRADIENT_FUNC = Err.number
End Function


'**********************************************************************************
'**********************************************************************************
'FUNCTION      : NS_DISCOUNT_FACTOR_FUNC
'DESCRIPTION   : Returns the discount factor per period
'LIBRARY       : FIXED_INCOME
'GROUP         : NS_YIELD
'ID            : 007
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 10-06-2008
'**********************************************************************************
'**********************************************************************************

Function NS_DISCOUNT_FACTOR_FUNC(ByVal TENOR_VAL As Double, _
ByVal YIELD_VAL As Double, _
Optional ByVal FREQUENCY As Double = 1, _
Optional ByVal VERSION As Integer = 0)

On Error GoTo ERROR_LABEL

Select Case VERSION
Case 0
   NS_DISCOUNT_FACTOR_FUNC = Exp(-TENOR_VAL * YIELD_VAL * FREQUENCY)
Case Else
   NS_DISCOUNT_FACTOR_FUNC = 1 / (1 + (YIELD_VAL / FREQUENCY)) ^ (TENOR_VAL * FREQUENCY)
End Select

Exit Function
ERROR_LABEL:
NS_DISCOUNT_FACTOR_FUNC = Err.number
End Function


'**********************************************************************************
'**********************************************************************************
'FUNCTION      : NS_HW_ZERO_FUNC
'DESCRIPTION   : Calculates the Zero Rate using the HW & Nelson Siegel Frame
'LIBRARY       : FIXED_INCOME
'GROUP         : NS_YIELD
'ID            : 008
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 10-06-2008
'**********************************************************************************
'**********************************************************************************

Function NS_HW_ZERO_FUNC(ByVal DELTA_TENOR_VAL As Double, _
ByVal TENOR As Double, _
ByVal KAPPA As Double, _
ByVal SIGMA As Double, _
ByVal BETA0_VAL As Double, _
ByVal BETA1_VAL As Double, _
ByVal BETA2_VAL As Double, _
ByVal BETA3_VAL As Double, _
ByVal TAU1_VAL As Double, _
ByVal TAU2_VAL As Double, _
ByVal SHORT_RATE As Double)

'--> Calibration Of the Hull-White Two Factor Yield model all the way through
'Nelson Siegal Parameters. Remember the volatility structure in the future is
'liable to be quite different form that observed in the market todya. So, this
'technique provides a richer pattern of term structure movements and a richer
'pattern of volatilities than one-factor models of the short rate. Nevertheles,
'it does not give the user complete freedom in choosing the volatility structure.
'So, once I have more free time I will work on the HJM and LMM models.

On Error GoTo ERROR_LABEL

If TAU2_VAL = 0 Then: TAU2_VAL = TAU1_VAL
'for basic Nelson Siegel TAU2_VAL = TAU1_VAL

'Long-run levels of interest rates (BETA0_VAL)

'Short-run component (BETA1_VAL)

'Short-Rate = BETA 0 + BETA1_VAL

'Medium-term component (BETA2_VAL) --> determines magnitude and the
'direction of the hump

'Long-Term component (BETA3_VAL) --> optional parameter proposed by Svensson (1994)

'Decay parameter 1 --> determines decay of short-term component, must be > 0

'Decay parameter 2 --> determines decay of medium-term component, must be > 0
'Spot rate at time t

NS_HW_ZERO_FUNC = NS_HW_A_FUNC(DELTA_TENOR_VAL, TENOR, KAPPA, SIGMA, BETA0_VAL, BETA1_VAL, BETA2_VAL, BETA3_VAL, TAU1_VAL, TAU2_VAL) * Exp(-NS_HW_B_FUNC(TENOR - DELTA_TENOR_VAL, SIGMA) * SHORT_RATE)

Exit Function
ERROR_LABEL:
NS_HW_ZERO_FUNC = Err.number
End Function


'**********************************************************************************
'**********************************************************************************
'FUNCTION      : NS_HW_A_FUNC
'DESCRIPTION   : First Factor of the Hull-White Model
'LIBRARY       : FIXED_INCOME
'GROUP         : NS_YIELD
'ID            : 009
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 10-06-2008
'**********************************************************************************
'**********************************************************************************

Function NS_HW_A_FUNC(ByVal DELTA_TENOR_VAL As Double, _
ByVal TENOR As Double, _
ByVal KAPPA As Double, _
ByVal SIGMA As Double, _
ByVal BETA0_VAL As Double, _
ByVal BETA1_VAL As Double, _
ByVal BETA2_VAL As Double, _
ByVal BETA3_VAL As Double, _
ByVal TAU1_VAL As Double, _
ByVal TAU2_VAL As Double)

Dim W1_VAL As Double
Dim W2_VAL As Double
Dim W3_VAL As Double
Dim W4_VAL As Double
Dim YIELD_VAL As Double

On Error GoTo ERROR_LABEL

'Long-run levels of interest rates (BETA0_VAL)

'Short-run component (BETA1_VAL)

'Medium-term component (BETA2_VAL) -->
'determines magnitude and the direction of the hump

'Long-Term component (BETA3_VAL) --> optional parameter proposed by Svensson (1994)

'Decay parameter 1 --> determines decay of short-term component, must be > 0

'Decay parameter 2 --> determines decay of medium-term component, must be > 0

'Spot rate at time t

YIELD_VAL = NS_SVENSSON_FORWARD_FUNC(DELTA_TENOR_VAL, BETA0_VAL, BETA1_VAL, BETA2_VAL, BETA3_VAL, TAU1_VAL, TAU2_VAL)
W1_VAL = Exp(-TENOR * NS_SVENSSON_SMOOTHING_FUNC(TENOR, BETA0_VAL, BETA1_VAL, BETA2_VAL, BETA3_VAL, TAU1_VAL, TAU2_VAL))
W2_VAL = Exp(-DELTA_TENOR_VAL * NS_SVENSSON_SMOOTHING_FUNC(DELTA_TENOR_VAL, BETA0_VAL, BETA1_VAL, BETA2_VAL, BETA3_VAL, TAU1_VAL, TAU2_VAL))
W3_VAL = Log(W1_VAL / W2_VAL) + NS_HW_B_FUNC(TENOR - DELTA_TENOR_VAL, SIGMA) * YIELD_VAL - (1 / (4 * (SIGMA ^ 3))) * (KAPPA ^ 2) * ((Exp(-SIGMA * TENOR) - Exp(-SIGMA * DELTA_TENOR_VAL)) ^ 2) * (Exp(2 * SIGMA * DELTA_TENOR_VAL) - 1)
W4_VAL = Exp(W3_VAL)
NS_HW_A_FUNC = W4_VAL

Exit Function
ERROR_LABEL:
NS_HW_A_FUNC = Err.number
End Function


'**********************************************************************************
'**********************************************************************************
'FUNCTION      : NS_HW_B_FUNC
'DESCRIPTION   : Second Factor of the Hull-White Model
'LIBRARY       : FIXED_INCOME
'GROUP         : NS_YIELD
'ID            : 010
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'**********************************************************************************
'**********************************************************************************

Function NS_HW_B_FUNC(ByVal DELTA_TENOR_VAL As Double, _
ByVal SIGMA As Double)

On Error GoTo ERROR_LABEL

NS_HW_B_FUNC = (1 / SIGMA) * (1 - Exp(-SIGMA * DELTA_TENOR_VAL))

Exit Function
ERROR_LABEL:
NS_HW_B_FUNC = Err.number
End Function
