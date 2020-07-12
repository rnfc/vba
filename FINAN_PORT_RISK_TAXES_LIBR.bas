Attribute VB_Name = "FINAN_PORT_RISK_TAXES_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.

'You begin with a $10,000 portfolio, with constant annual return and additional
'annual investments of $100. Your initial after-tax dividend is $70, with constant
'annual percentage increases.

'Each year a percentage f = 25% of your portfolio is taxed at the capital gains rate
'CAPITAL_GAINS_TAX_RATE_PER_PERIOD = 35%. After all Capital Gains taxes, what is your
'portfolio worth … after TEMP_MATRIX(i, 1) years?

'Reference: http://www.gummy-stuff.org/bogle-problem1.htm
'Final after-tax Portfolio: { (1-T) gn + T } A0 + (1 - T) D (gn - xn) / (g - x) + T D (xn - 1) / (x - 1)

Function FINAL_AFTER_TAX_PORTFOLIO_FUNC( _
Optional ByVal RETURN_PER_PERIOD As Double = 0.06, _
Optional ByVal DIVIDEND_INCREASE_PER_PERIOD As Double = 0.05, _
Optional ByVal INITIAL_PORTFOLIO As Double = 10000, _
Optional ByVal ADDITIONAL_INVESTMENT As Double = 100, _
Optional ByVal FIRST_DIVIDEND As Double = 70, _
Optional ByVal PORTFOLIO_FRACTION As Double = 0.25, _
Optional ByVal CAPITAL_GAINS_TAX_RATE_PER_PERIOD As Double = 0.35, _
Optional ByVal MIN_PERIOD As Double = 0, _
Optional ByVal DELTA_PERIOD As Double = 1, _
Optional ByVal NBINS As Long = 21)

'FIRST_DIVIDEND --> first after-tax dividend (re-invested)
'FIRST_DIVIDEND --> d(1-t); d = annual dividend; t = dividend tax rate
'PORTFOLIO_FRACTION--> fraction of period portfolio that is capital-gain taxed

Dim i As Long

Dim G1_VAL As Double 'period growth factor for the assets in the portfolio
'(so $1 grows to $g in 1 year)
Dim G2_VAL As Double 'period growth factor for the dividends
'(so a $1 dividend in one period would be $X the following period
Dim G3_VAL As Double 'tax-reduced annual gain factor

Dim TEMP_PERIOD As Double
Dim TEMP_MATRIX As Variant

On Error GoTo ERROR_LABEL

ReDim TEMP_MATRIX(0 To NBINS, 1 To 5)
TEMP_MATRIX(0, 1) = "N"
TEMP_MATRIX(0, 2) = "P1"
TEMP_MATRIX(0, 3) = "P2"
TEMP_MATRIX(0, 4) = "P3"
TEMP_MATRIX(0, 5) = "P1+P2+P3"

G1_VAL = RETURN_PER_PERIOD + 1
G2_VAL = DIVIDEND_INCREASE_PER_PERIOD + 1
G3_VAL = G1_VAL - CAPITAL_GAINS_TAX_RATE_PER_PERIOD * _
         PORTFOLIO_FRACTION * (G1_VAL - 1)

TEMP_PERIOD = MIN_PERIOD
i = 1
TEMP_MATRIX(i, 1) = TEMP_PERIOD
TEMP_MATRIX(i, 2) = INITIAL_PORTFOLIO
TEMP_MATRIX(i, 3) = 0
TEMP_MATRIX(i, 4) = 0
TEMP_MATRIX(i, 5) = TEMP_MATRIX(i, 2) + TEMP_MATRIX(i, 3) + TEMP_MATRIX(i, 4)
TEMP_PERIOD = TEMP_PERIOD + 1

For i = 2 To NBINS
    TEMP_MATRIX(i, 1) = TEMP_PERIOD
    TEMP_MATRIX(i, 2) = ((1 - CAPITAL_GAINS_TAX_RATE_PER_PERIOD) * G3_VAL ^ TEMP_MATRIX(i, 1) + _
                        CAPITAL_GAINS_TAX_RATE_PER_PERIOD) * INITIAL_PORTFOLIO
    
    TEMP_MATRIX(i, 3) = (1 - CAPITAL_GAINS_TAX_RATE_PER_PERIOD) * FIRST_DIVIDEND * _
                        (G3_VAL ^ TEMP_MATRIX(i, 1) - G2_VAL ^ TEMP_MATRIX(i, 1)) / _
                        (G3_VAL - G2_VAL) + CAPITAL_GAINS_TAX_RATE_PER_PERIOD * _
                        FIRST_DIVIDEND * (G2_VAL ^ TEMP_MATRIX(i, 1) - 1) / (G2_VAL - 1)
    
    TEMP_MATRIX(i, 4) = ((1 - CAPITAL_GAINS_TAX_RATE_PER_PERIOD) * (G3_VAL ^ _
                        TEMP_MATRIX(i, 1) - 1) / (G3_VAL - 1) + TEMP_MATRIX(i, 1) * _
                        CAPITAL_GAINS_TAX_RATE_PER_PERIOD) * ADDITIONAL_INVESTMENT
    
    TEMP_MATRIX(i, 5) = TEMP_MATRIX(i, 2) + TEMP_MATRIX(i, 3) + TEMP_MATRIX(i, 4)
    TEMP_PERIOD = TEMP_PERIOD + 1
Next i

FINAL_AFTER_TAX_PORTFOLIO_FUNC = TEMP_MATRIX

Exit Function
ERROR_LABEL:
FINAL_AFTER_TAX_PORTFOLIO_FUNC = Err.number
End Function
