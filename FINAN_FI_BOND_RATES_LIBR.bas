Attribute VB_Name = "FINAN_FI_BOND_RATES_LIBR"


Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.

Private PUB_GROWTH_VAL As Double


Function INTEREST_RATE_DISCOUNT_FUNC(ByVal END_DATE As Date, _
ByVal START_DATE As Date, _
ByVal CARRY_RATE As Double, _
ByVal X0_VAL As Double, _
Optional ByVal METHOD_VAL As Variant = "nacc", _
Optional ByVal COUNT_BASIS As Integer = 0) As Double
'X0_VAL -> VALUE

Dim TAU_VAL As Double
On Error GoTo ERROR_LABEL
TAU_VAL = YEARFRAC_FUNC(START_DATE, END_DATE, COUNT_BASIS)
Select Case LCase(METHOD_VAL)
Case "nacc", ""
'    INTEREST_RATE_DISCOUNT_FUNC = X0_VAL * Exp(-(END_DATE - START_DATE) / 365 * CARRY_RATE)
    INTEREST_RATE_DISCOUNT_FUNC = X0_VAL * Exp(-TAU_VAL * CARRY_RATE)
Case Else
'    INTEREST_RATE_DISCOUNT_FUNC = X0_VAL * Exp(-(END_DATE - START_DATE) / 365 * INTEREST_RATE_CONVERTER_FUNC(CARRY_RATE, METHOD_VAL, 0, START_DATE, END_DATE))
    INTEREST_RATE_DISCOUNT_FUNC = X0_VAL * Exp(-TAU_VAL * INTEREST_RATE_CONVERTER_FUNC(CARRY_RATE, METHOD_VAL, 0, START_DATE, END_DATE))
End Select

Exit Function
ERROR_LABEL:
INTEREST_RATE_DISCOUNT_FUNC = Err.number
End Function

Function INTEREST_RATE_COMPOUND_FUNC(ByVal END_DATE As Date, _
ByVal START_DATE As Date, _
ByVal CARRY_RATE As Double, _
ByVal X0_VAL As Double, _
Optional ByVal METHOD_VAL As Variant = "nacc", _
Optional ByVal COUNT_BASIS As Integer = 0) As Double
'X0_VAL -> VALUE

Dim TAU_VAL As Double
On Error GoTo ERROR_LABEL
TAU_VAL = YEARFRAC_FUNC(START_DATE, END_DATE, COUNT_BASIS)
Select Case LCase(METHOD_VAL)
Case "nacc", ""
'    INTEREST_RATE_COMPOUND_FUNC = X0_VAL * Exp((END_DATE - START_DATE) / 365 * CARRY_RATE)
    INTEREST_RATE_COMPOUND_FUNC = X0_VAL * Exp(TAU_VAL * CARRY_RATE)
Case Else
'    INTEREST_RATE_COMPOUND_FUNC = X0_VAL * Exp((END_DATE - START_DATE) / 365 * INTEREST_RATE_CONVERTER_FUNC(CARRY_RATE, METHOD_VAL, 0, START_DATE, END_DATE))
    INTEREST_RATE_COMPOUND_FUNC = X0_VAL * Exp(TAU_VAL * INTEREST_RATE_CONVERTER_FUNC(CARRY_RATE, METHOD_VAL, 0, START_DATE, END_DATE))
End Select

Exit Function
ERROR_LABEL:
INTEREST_RATE_COMPOUND_FUNC = Err.number
End Function


'IC0_PERIOD_VAL --> INPUT_COMPOUNDING_PERIOD
'OC0_PERIOD_VAL --> OUTPUT_COMPOUNDING_PERIOD

Function INTEREST_RATE_CONVERTER_FUNC(ByVal RATE_VAL As Double, _
ByVal IC0_PERIOD_VAL As Variant, _
ByVal OC0_PERIOD_VAL As Variant, _
Optional ByVal START_DATE As Date, _
Optional ByVal FORWARD_DATE As Date, _
Optional ByVal COUNT_BASIS As Integer = 0) As Double

'INPUT_RATE --> nominal interest rate
'COUNT BASIS:
'   0 or omitted US (NASD) 30/360
'   1 Actual/actual
'   2 Actual/360
'   3 Actual/365
'   4 European 30 / 360

On Error GoTo ERROR_LABEL

If IsNumeric(IC0_PERIOD_VAL) And IsNumeric(OC0_PERIOD_VAL) Then 'Convert Rate In/Out
    INTEREST_RATE_CONVERTER_FUNC = INTEREST_RATE_QUICK_FUNC(RATE_VAL, CDbl(IC0_PERIOD_VAL), CDbl(OC0_PERIOD_VAL))
Else 'Returns the effective/nominal interest rate
    Dim TAU_VAL As Double
    Dim IC1_PERIOD_VAL, OC1_PERIOD_VAL As Double
    'TAU_VAL = (FORWARD_DATE - START_DATE) / 365
    TAU_VAL = YEARFRAC_FUNC(START_DATE, FORWARD_DATE, COUNT_BASIS)
    If TAU_VAL = 0 Then TAU_VAL = 1
    Select Case LCase(CStr(IC0_PERIOD_VAL)) 'GROWTH FACTOR CALCULATION
    Case "simple", "jibar"
        If TAU_VAL >= 0 Then PUB_GROWTH_VAL = 1 + RATE_VAL * TAU_VAL Else PUB_GROWTH_VAL = 1 / (1 - RATE_VAL * TAU_VAL)
    Case 0, "nacc", "continuous", ""
        PUB_GROWTH_VAL = Exp(RATE_VAL * TAU_VAL)
    Case Else
        If IsNumeric(IC0_PERIOD_VAL) = True Then IC1_PERIOD_VAL = CDbl(IC0_PERIOD_VAL) Else IC1_PERIOD_VAL = INTEREST_RATE_PERIOD_FUNC(CStr(IC0_PERIOD_VAL))
        PUB_GROWTH_VAL = (1 + RATE_VAL * IC1_PERIOD_VAL) ^ (TAU_VAL / IC1_PERIOD_VAL)
    End Select
    Select Case LCase(CStr(OC0_PERIOD_VAL))     'OUTPUT CALCULATION
    Case "simple", "jibar"
        INTEREST_RATE_CONVERTER_FUNC = (PUB_GROWTH_VAL - 1) / TAU_VAL
    Case 0, "nacc", ""
        INTEREST_RATE_CONVERTER_FUNC = Log(PUB_GROWTH_VAL) / TAU_VAL
    Case Else
        If IsNumeric(OC0_PERIOD_VAL) = True Then OC1_PERIOD_VAL = CDbl(OC0_PERIOD_VAL) Else OC1_PERIOD_VAL = INTEREST_RATE_PERIOD_FUNC(CStr(OC0_PERIOD_VAL))
        INTEREST_RATE_CONVERTER_FUNC = 1 / OC1_PERIOD_VAL * (PUB_GROWTH_VAL ^ (OC1_PERIOD_VAL / TAU_VAL) - 1)
    End Select
End If

Exit Function
ERROR_LABEL:
INTEREST_RATE_CONVERTER_FUNC = Err.number
End Function

Private Function INTEREST_RATE_QUICK_FUNC(ByVal RATE_VAL As Double, _
ByVal IC0_PERIOD_VAL As Double, _
ByVal OC0_PERIOD_VAL As Double) As Double

On Error GoTo ERROR_LABEL

If IC0_PERIOD_VAL = 0 Then PUB_GROWTH_VAL = Exp(RATE_VAL) Else PUB_GROWTH_VAL = (1 + RATE_VAL * IC0_PERIOD_VAL) ^ (1 / IC0_PERIOD_VAL)
If OC0_PERIOD_VAL = 0 Then INTEREST_RATE_QUICK_FUNC = Log(PUB_GROWTH_VAL) Else INTEREST_RATE_QUICK_FUNC = (PUB_GROWTH_VAL ^ OC0_PERIOD_VAL - 1) / OC0_PERIOD_VAL

Exit Function
ERROR_LABEL:
INTEREST_RATE_QUICK_FUNC = Err.number
End Function

Private Function INTEREST_RATE_PERIOD_FUNC(ByVal METHOD_VAL As Variant) As Double

On Error GoTo ERROR_LABEL

Select Case LCase(Trim(METHOD_VAL))
Case "naca"
    INTEREST_RATE_PERIOD_FUNC = 1
Case "nacs"
    INTEREST_RATE_PERIOD_FUNC = 0.5
Case "nacd"
    INTEREST_RATE_PERIOD_FUNC = 1 / 365
Case "nacq"
    INTEREST_RATE_PERIOD_FUNC = 0.25
Case "nacm"
    INTEREST_RATE_PERIOD_FUNC = 1 / 12
Case "nacc"
    INTEREST_RATE_PERIOD_FUNC = 0
End Select

Exit Function
ERROR_LABEL:
INTEREST_RATE_PERIOD_FUNC = Err.number
End Function


'A statement by a bank that the interest rate on one-year deposits is 10%
'per annum sounds straightforward and unambiguous. In fact, its precise
'meaning depends on the way the interest rate is measured.
'If the interest rate is measured with annual compounding, the bank's
'statement that the interest rate is 10% means that $100 grows to:
'$100 x 1.1  = $110
'at the end of one year. When the interest rate is measured with semiannual
'compounding, it means that we earn 5% every six months, with the interest
'being reinvested. In this case $100 grows to $100 x 1.05 x 1.05 = $110.25
'at the end of one year. When the intrest rate is measured with quaterly
'compounding, the bank's statement menas that we earn 2.5% every three months,
'with the interest being reinvested. The $100 then grows to
'$100 x 1.0254 = $110.38
'At the end of one year. The model included with this note called, the
'"Magic of compounding", shows the effect of increasing the compounding
'frequency further. The compounding frequency defines the units in which
'an interest rate is measured. A rate expressed with one compounding
'frequency can be converted into an equivalent rate with a different
'compounding frequency. For example:

'Compounding Frequency   Effect Rate   Frequency
'Annually                   10.000%        1
'Semiannually               10.250%        2
'Quaterly                   10.381%        4
'Monthly                    10.471%       12
'Weekly                     10.506%       52
'Daily                      10.516%      365
'Continuous                  9.531%        e

'From this table we see that 10.25% with annual compounding is equivalent to
'10% with semiannual compounding. We can think of the difference between one
'compounding frequency and another to be analogous to the difference between
'kilometers and miles. They are two different units of measurement.

'To generalize our results, suppose that an amount A is invested for n years
'at an interest rate of R per annum. If the rate is compounded once per annum,
'the terminal value of the investment is: A ( 1 + R ) n

'If the rate is compounded m times per annum, the terminal value of the investment
'is: A ( 1 + R / m ) n x m
'When m =1 the rate is sometimes referred to as the equivalent annual interest
'rate.

'---> Continuous Compounding
'The limit as the compounding frequency, m, tends to infinity is known as
'continuous compounding. With continuous compounding, it can be shown that
'an amount A invested for n years at rate R grows to A x e^(r) x n
'Where e = 2.71828. The function ex is the exponential function and is
'built into most calculators, so the computation of the expression present
'no problems.
