Attribute VB_Name = "FINAN_PORT_SIMUL_WITHDR_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.

'************************************************************************************
'************************************************************************************
'FUNCTION      : PORTFOLIO_WITHDRAWAL_SIMULATION_FUNC

'DESCRIPTION   :

'This algorithm uses Monte Carlo simulation to model the retirement draw-down
'phase of an investment portfolio under varying withdrawal rate, inflation,
'investment performance, and withdrawal policy assumptions.

'Simulation uses a stochastic (Monte Carlo) rather than drawing from an actual
'historical return series. Each period (year), the simulation draws a random
'value for return and for inflation that is normally disributed around the
'respective mean and standardard deviation for the variable.

'Returns of individual asset classes, as well as covariance between asset classes
'are not modeled.

'References:
'    March 2006 FPA Journal Article by Jonathan Guyton and
'    William Klinger
'        http://www.fpanet.org/journal/articles/2006_Issues/jfp0306-art6.cfm

'LIBRARY       : PORTFOLIO
'GROUP         : WITHDRAWAL
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 21/01/2009
'************************************************************************************
'************************************************************************************


Function PORTFOLIO_WITHDRAWAL_SIMULATION_FUNC(ByRef PERIODS As Long, _
ByRef RETURN_AVG_VAL As Double, _
ByRef RETURN_STDEV_VAL As Double, _
ByRef INFLATION_AVG_VAL As Double, _
ByRef INFLATION_STDEV_VAL As Double, _
ByRef START_BALANCE_VAL As Double, _
ByRef PORTFOLIO_TAX_RATE_VAL As Double, _
ByRef START_WITHDRAWAL_VAL As Double, _
ByRef AVERAGE_ENDING_WITHDRAWAL_VAL As Double, _
ByRef AVERAGE_ENDING_BALANCE_VAL As Double, _
ByRef SUCCESS_PROBABILITY_VAL As Double, _
ByRef AVERAGE_PERCENT_MAINTAINED_VAL As Double, _
ByRef ENDING_BALANCE_ARR() As Double, _
ByRef ENDING_WITHDRAWAL_ARR() As Double, _
Optional ByVal WITHDRAWAL_POLICY_OPT As Integer = 1, _
Optional ByVal nLOOPS As Long = 10000, _
Optional ByVal REFRESH_PERIOD As Long = 0, _
Optional ByRef PROGRESS_OBJECT As Object)

'WITHDRAWAL_POLICY_OPT
' CONSERVATIVE = 1
' FLEXIBLE = 2
' STABLE = 3

'------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------
'Withdrawal Policy Name  Description
'Stable PP --> Start withdrawing the specified percent of the total portfolio
'balance.

'Adjust the withdrawal for inflation each year to maintain purchasing power.

'Conservative --> Start withdrawing the specified percent of the total portfolio
'balance.Adjust the withdrawal for inflation, as long as 1) the portfolio balance
'didn't fall in the past year and 2) Portfolio Balance hasn't dropped below the
'starting balance (inflation adjusted). Missed cost of living adjustments don't get
'made up if things turn around, even if the portfolio balance skyrockets.

'Flexible --> Same as Conservative in the lean years, but this policy allows the
'withdrawal to increasebeyond the starting value and by more than the inflation
'rate if the portfolio is doing really well (The extra increase in any year can
'be as high as the inflation rate (double cola) if the portfolio has doubled in
'value since the start.
'------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------
                
Dim i As Long 'Iteration
Dim j As Long 'Year
Dim k As Long 'SuccessCount

Dim INFLATION_VAL As Double
Dim PORTFOLIO_RETURN_VAL As Double

Dim WITHDRAWAL_VAL As Double
Dim PREVIOUS_WITHDRAWAL_VAL As Double
Dim ENDING_WITHDRAWAL_TOTAL_VAL As Double
Dim ENDING_BALANCE_TOTAL_VAL As Double

Dim PREVIOUS_BALANCE_VAL As Double
Dim PORTFOLIO_BALANCE_VAL As Double

'REFRESH_PERIOD: How often should we call out to let the world
'know what we're up to?

On Error GoTo ERROR_LABEL
  
PORTFOLIO_WITHDRAWAL_SIMULATION_FUNC = False

' Initialize the random number generator used by the
' RANDOM_NORMAL_FAST_FUNC() function.
Randomize

ReDim ENDING_BALANCE_ARR(0 To nLOOPS)
ReDim ENDING_WITHDRAWAL_ARR(0 To nLOOPS)

ENDING_WITHDRAWAL_TOTAL_VAL = 0
ENDING_BALANCE_TOTAL_VAL = 0
k = 0

For i = 1 To nLOOPS
  PORTFOLIO_BALANCE_VAL = START_BALANCE_VAL
  PREVIOUS_BALANCE_VAL = START_BALANCE_VAL
  PREVIOUS_WITHDRAWAL_VAL = START_WITHDRAWAL_VAL
  WITHDRAWAL_VAL = START_WITHDRAWAL_VAL
  
  For j = 1 To PERIODS
  
   ' Get normally distributed random values for inflation and return
   INFLATION_VAL = (RANDOM_NORMAL_FAST_FUNC() * INFLATION_STDEV_VAL) + _
   INFLATION_AVG_VAL
   
   PORTFOLIO_RETURN_VAL = (RANDOM_NORMAL_FAST_FUNC() * RETURN_STDEV_VAL) + _
   RETURN_AVG_VAL

   PORTFOLIO_BALANCE_VAL = TAXABLE_BALANCE_FUNC(PREVIOUS_BALANCE_VAL, _
   PREVIOUS_WITHDRAWAL_VAL, PORTFOLIO_TAX_RATE_VAL, PORTFOLIO_RETURN_VAL, _
   INFLATION_VAL)
   
   If (PORTFOLIO_BALANCE_VAL <= 0) Then
       PORTFOLIO_BALANCE_VAL = 0 ' Bummer, we ran out of money on this run!!!
   Else
      If (WITHDRAWAL_POLICY_OPT = 1) Then
        WITHDRAWAL_VAL = CONSERVATIVE_WITHDRAWAL_FUNC(PORTFOLIO_BALANCE_VAL, _
        PREVIOUS_BALANCE_VAL, START_BALANCE_VAL, PREVIOUS_WITHDRAWAL_VAL, _
        START_WITHDRAWAL_VAL, INFLATION_VAL)
      Else
        If (WITHDRAWAL_POLICY_OPT = 2) Then
          WITHDRAWAL_VAL = FLEXIBLE_WITHDRAWAL_FUNC(PORTFOLIO_BALANCE_VAL, _
          PREVIOUS_BALANCE_VAL, START_BALANCE_VAL, PREVIOUS_WITHDRAWAL_VAL, _
          START_WITHDRAWAL_VAL, INFLATION_VAL)
        Else ' For Stable PP, there's no change in withdrawal,
             ' which means it gets to exactly keep up with inflation
        End If
      End If
   End If
   
   PREVIOUS_WITHDRAWAL_VAL = WITHDRAWAL_VAL
   PREVIOUS_BALANCE_VAL = PORTFOLIO_BALANCE_VAL
  Next j
  
  ENDING_BALANCE_ARR(i) = PORTFOLIO_BALANCE_VAL
  ENDING_WITHDRAWAL_ARR(i) = WITHDRAWAL_VAL

  If (WITHDRAWAL_VAL > 0) Then
    ENDING_WITHDRAWAL_TOTAL_VAL = ENDING_WITHDRAWAL_TOTAL_VAL + WITHDRAWAL_VAL
  End If
  
  If (PORTFOLIO_BALANCE_VAL > 0) Then
    k = k + 1
    ENDING_BALANCE_TOTAL_VAL = ENDING_BALANCE_TOTAL_VAL + PORTFOLIO_BALANCE_VAL
  End If
 
  ' Only call out to update progress in the UI periodically
  ' This check is essential for decent simulation progress
  If Not PROGRESS_OBJECT Is Nothing Then
      If ((i Mod REFRESH_PERIOD) = 0) Then
        PROGRESS_OBJECT.value = i
      End If
  End If
Next i

If Not PROGRESS_OBJECT Is Nothing Then
    ' Update Progress one last time (important because of mod() above)
    PROGRESS_OBJECT.value = nLOOPS
End If
  
' Update results values (passed by reference)
AVERAGE_ENDING_WITHDRAWAL_VAL = ENDING_WITHDRAWAL_TOTAL_VAL / nLOOPS
AVERAGE_ENDING_BALANCE_VAL = ENDING_BALANCE_TOTAL_VAL / nLOOPS
SUCCESS_PROBABILITY_VAL = k / nLOOPS
AVERAGE_PERCENT_MAINTAINED_VAL = AVERAGE_ENDING_WITHDRAWAL_VAL / _
START_WITHDRAWAL_VAL

PORTFOLIO_WITHDRAWAL_SIMULATION_FUNC = True
  
Exit Function
ERROR_LABEL:
PORTFOLIO_WITHDRAWAL_SIMULATION_FUNC = False
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : CONSERVATIVE_WITHDRAWAL_FUNC

'DESCRIPTION   :
' Compute the withdrawal amount for the current year using the CONSERVATIVE method
' The general idea of this method is to skip COLAs (cost of living adjustments) on
' the withdrawal amount during periods of poor portfolio performance to save cash.
'
' The code implements this by subtracting out the current period's inflation rate
' from the withdrawal amount.  The subtraction is needed because all dollar amounts in the
' simulation are maintained through the years at present value (in start year dollars).
' In order to not give the retiree a cost of living adjustment, you need to subtract it
' from the withdrawal amount.  In real life, the retiree's withdrawal would be frozen at
' the same level as it was in the year prior to when the conservation action kicked in.
'
' Finally, notice that with this method, there's no way to make up for missed COLAs
' and there are no bonus increases in the withdrawal amount if things are going great.

'LIBRARY       : PORTFOLIO
'GROUP         : WITHDRAWAL
'ID            : 002
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 21/01/2009
'************************************************************************************
'************************************************************************************

Private Function CONSERVATIVE_WITHDRAWAL_FUNC( _
ByRef CURRENT_BALANCE_VAL As Double, _
ByRef PREVIOUS_BALANCE_VAL As Double, _
ByRef START_BALANCE_VAL As Double, _
ByRef PREVIOUS_WITHDRAWAL_VAL As Double, _
ByRef START_WITHDRAWAL_VAL As Double, _
ByRef INFLATION_RATE_VAL As Double)
                            
On Error GoTo ERROR_LABEL

' If portfolio is not doing well so the retiree doesn't get a cost of living increase
If ((CURRENT_BALANCE_VAL < PREVIOUS_BALANCE_VAL) And _
    (CURRENT_BALANCE_VAL < START_BALANCE_VAL)) Then
   CONSERVATIVE_WITHDRAWAL_FUNC = PREVIOUS_WITHDRAWAL_VAL * (1 - _
   INFLATION_RATE_VAL)
Else
   ' The portfolio is doing ok so the retiree gets a standard COLA based on the
   ' inflation rate. See above for explaination of why withdrawal isn't increased.
   CONSERVATIVE_WITHDRAWAL_FUNC = PREVIOUS_WITHDRAWAL_VAL
End If
  
Exit Function
ERROR_LABEL:
CONSERVATIVE_WITHDRAWAL_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : FLEXIBLE_WITHDRAWAL_FUNC

'DESCRIPTION   :
' Compute the withdrawal amount for the current period using the FLEXIBLE method.
' Same as Conservative in the lean years, but this policy allows the withdrawal to
' increase by more than the inflation rate if the portfolio is doing really
' well.  The extra increase in any year can be as high as the inflation rate (double COLA)
' if the portfolio has doubled in value since the start.

'LIBRARY       : PORTFOLIO
'GROUP         : WITHDRAWAL
'ID            : 003
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 21/01/2009
'************************************************************************************
'************************************************************************************

Private Function FLEXIBLE_WITHDRAWAL_FUNC(ByRef CURRENT_BALANCE_VAL As Double, _
ByRef PREVIOUS_BALANCE_VAL As Double, _
ByRef START_BALANCE_VAL As Double, _
ByRef PREVIOUS_WITHDRAWAL_VAL As Double, _
ByRef START_WITHDRAWAL_VAL As Double, _
ByRef INFLATION_RATE_VAL As Double)

On Error GoTo ERROR_LABEL

If ((CURRENT_BALANCE_VAL < PREVIOUS_BALANCE_VAL) And _
    (CURRENT_BALANCE_VAL < START_BALANCE_VAL)) Then
   ' Portfolio Value is shrinking, no COLA this year (reduce
   ' withdrawal by inflation rate)
   FLEXIBLE_WITHDRAWAL_FUNC = PREVIOUS_WITHDRAWAL_VAL * (1 - _
   INFLATION_RATE_VAL)
Else
   If (CURRENT_BALANCE_VAL > START_BALANCE_VAL) Then
       ' If portfolio balance is ok, then restore some extra purchasing power
       ' by increasing withdrawal by recent inflation rate, this increase
       If (CURRENT_BALANCE_VAL > (2 * START_BALANCE_VAL)) Then
          ' Portfolio has doubled in size, increase withdrawal by inflation
          ' rate twice
          FLEXIBLE_WITHDRAWAL_FUNC = PREVIOUS_WITHDRAWAL_VAL * _
                                    (1 + (INFLATION_RATE_VAL))
       Else
          ' Portfolio is above the start value, but not doubled, give a
          ' smaller increase.
          FLEXIBLE_WITHDRAWAL_FUNC = PREVIOUS_WITHDRAWAL_VAL * _
                                    (1 + (INFLATION_RATE_VAL / 4))
       End If
   Else
       ' Not a bad year, but portfolio isn't growing much so just give a
       ' normal COLA
       FLEXIBLE_WITHDRAWAL_FUNC = PREVIOUS_WITHDRAWAL_VAL
   End If
End If
  
Exit Function
ERROR_LABEL:
FLEXIBLE_WITHDRAWAL_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : TAXABLE_BALANCE_FUNC

'DESCRIPTION   :
'LIBRARY       : PORTFOLIO
'GROUP         : WITHDRAWAL
'ID            : 004
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 21/01/2009
'************************************************************************************
'************************************************************************************

Private Function TAXABLE_BALANCE_FUNC(ByRef PREVIOUS_BALANCE_VAL As Double, _
ByRef CURRENT_WITHDRAWAL_VAL As Double, _
ByRef TAX_RATE_VAL As Double, _
ByRef PORTFOLIO_RETURN_VAL As Double, _
ByRef INFLATION_RATE_VAL As Double)
  
Dim BALANCE_VAL As Double
Dim AFTER_TAX_GAIN_VAL As Double

On Error GoTo ERROR_LABEL

' Add in the portfolio return on the prior balance minus the effect of taxes
' First subtract out withdrawal
BALANCE_VAL = PREVIOUS_BALANCE_VAL - CURRENT_WITHDRAWAL_VAL

If (PORTFOLIO_RETURN_VAL >= 0) Then
   AFTER_TAX_GAIN_VAL = BALANCE_VAL * (PORTFOLIO_RETURN_VAL - _
                       (PORTFOLIO_RETURN_VAL * TAX_RATE_VAL))
Else
    ' If return was negative, don't deduct for taxes that year
    AFTER_TAX_GAIN_VAL = BALANCE_VAL * PORTFOLIO_RETURN_VAL
End If

TAXABLE_BALANCE_FUNC = BALANCE_VAL + AFTER_TAX_GAIN_VAL - _
                      (BALANCE_VAL * INFLATION_RATE_VAL)

If (TAXABLE_BALANCE_FUNC < 0) Then
   TAXABLE_BALANCE_FUNC = 0
End If
  
Exit Function
ERROR_LABEL:
TAXABLE_BALANCE_FUNC = Err.number
End Function
