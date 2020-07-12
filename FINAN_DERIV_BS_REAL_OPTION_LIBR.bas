Attribute VB_Name = "FINAN_DERIV_BS_REAL_OPTION_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                'default first index of arrays.

'************************************************************************************
'************************************************************************************
'FUNCTION      : CALL_EQUITY_VALUE_FUNC
'DESCRIPTION   : This program calculates the value of equity as a call option
'**  on the value of the underlying firm.
'LIBRARY       : DERIVATIVES
'GROUP         : REAL OPTIONS
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/29/2009
'************************************************************************************
'************************************************************************************

Function CALL_EQUITY_VALUE_FUNC(ByVal VALUE_FIRM As Double, _
ByVal CUMUL_FACE_VALUE As Double, _
ByVal AVG_DURATION As Double, _
ByVal VOLATILITY_CASH_FLOW As Double, _
ByVal DIVD As Double, _
ByVal RISK_FREE As Double, _
Optional ByVal OUTPUT As Integer = 0, _
Optional ByVal CND_TYPE As Integer = 0)

'VALUE_FIRM = Enter the value of the firm

'VOLATILITY_CASH_FLOW = There are three ways of estimating standard
'deviation. One is to use the firm's own stock and bond prices
'to estimate it. The other is to use the variance of the industry
'to which your firm belongs. Or you can just
'enter the annualized standard deviation in ln(value) of asset

'DIVD = Enter the expected dividend yield.

'CUMUL_FACE_VALUE = Cumulated face value of outstanding debt.
'Add coupons to the face value of debt (nominal terms)

'AVG_DURATION = Average duration of outstanding debt
'[Weighted by the face value of the debt]

'RISK_FREE = Enter the riskless rate that corresponds to the
'option lifetime

Dim D1_VAL As Double
Dim D2_VAL As Double

Dim SPOT As Double
Dim STRIKE As Double
Dim EXPIRATION As Double
Dim RATE As Double
Dim VOLATILITY As Double

Dim CALL_VALUE As Double

On Error GoTo ERROR_LABEL

SPOT = VALUE_FIRM
STRIKE = CUMUL_FACE_VALUE
EXPIRATION = AVG_DURATION
RATE = RISK_FREE
VOLATILITY = VOLATILITY_CASH_FLOW

D1_VAL = (Log(SPOT / STRIKE) + ((RATE - DIVD) + VOLATILITY ^ 2 / 2) * EXPIRATION) / (VOLATILITY * Sqr(EXPIRATION))
D2_VAL = D1_VAL - VOLATILITY * Sqr(EXPIRATION)
CALL_VALUE = (Exp((0 - DIVD) * EXPIRATION)) * SPOT * CND_FUNC(D1_VAL, CND_TYPE) - STRIKE * (Exp((0 - RATE) * EXPIRATION)) * CND_FUNC(D2_VAL, CND_TYPE)

Select Case OUTPUT
Case 0
    CALL_EQUITY_VALUE_FUNC = CALL_VALUE 'Value of equity as a call
Case 1
    CALL_EQUITY_VALUE_FUNC = SPOT - CALL_VALUE  'Value of outstanding debt
Case Else
    CALL_EQUITY_VALUE_FUNC = (STRIKE / (SPOT - CALL_VALUE)) ^ (1 / EXPIRATION) - 1 'Appropriate interest rate for debt (annualized)
End Select

Exit Function
ERROR_LABEL:
CALL_EQUITY_VALUE_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : STOCK_BOND_OPTION_VARIANCE_FUNC
'DESCRIPTION   :
'LIBRARY       : DERIVATIVES
'GROUP         : REAL OPTIONS
'ID            : 002
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/29/2009
'************************************************************************************
'************************************************************************************

Function STOCK_BOND_OPTION_VARIANCE_FUNC(ByVal VOLATILITY_STOCK As Double, _
ByVal VOLATILITY_BOND As Double, _
ByVal RHO_VAL As Double, _
ByVal DE_RATIO As Double)

On Error GoTo ERROR_LABEL

'VOLATILITY_STOCK: Standard deviation in the firm's stock price (ln)
'VOLATILITY_BOND: Standard deviation in the firm's bond price (ln)
'RHO: Correlation between the stock and bond prices
'DE_RATIO: Average D/(D+E) ratio during the variance estimation period

STOCK_BOND_OPTION_VARIANCE_FUNC = (((1 - DE_RATIO) ^ 2) * (VOLATILITY_STOCK ^ 2) + DE_RATIO ^ 2 * VOLATILITY_BOND ^ 2 + 2 * DE_RATIO * (1 - DE_RATIO) * VOLATILITY_STOCK * VOLATILITY_BOND * RHO_VAL) ^ (0.5)

Exit Function
ERROR_LABEL:
STOCK_BOND_OPTION_VARIANCE_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : ABANDONMENT_OPTION_FUNC
'DESCRIPTION   : Calculates the value of an abandonment option
'Firms often use the excess capacity that they have on an existing plant,
'storage facility or computer resource for a new project. When they do so,
'they make one of two assumptions:

'1. They assume that excess capacity is free, since it is not being
'used currently and cannot be sold off or rented, in most cases.
'2. They allocate a portion of the book value of the plant or resource
'to the project. Thus, if the plant has a book value of $ 100 million and
'the new project uses 40% of it, $ 40 million will be allocated to the project.

'We will argue that neither of these approaches considers the opportunity
'cost of using excess capacity, since the opportunity cost comes usually
'comes from costs that the firm will face in the future as a consequence
'of using up excess capacity today. By using up excess capacity on a new
'project, the firm will run out of capacity sooner than it would if it
'did not take the project. When it does run out of capacity, it has to
'take one of two paths:

'•  Mac183; New capacity will have to be bought or built when
'capacity runs  out, in which case the opportunity cost will be
'the higher cost in present value terms of doing this earlier
'rather than later.

'•  Mac183; Production will have to be cut back on one of the
'product lines, leading to a loss in cash flows that would have
'been generated by the lost sales.

'Again, this choice is not random, since the logical action to take
'is the one that leads to the lower cost, in present value terms, for
'the firm. Thus, if it cheaper to lose sales rather than build new
'capacity, the opportunity cost for the project being considered
'should be based on the lost sales.

'LIBRARY       : DERIVATIVES
'GROUP         : REAL OPTIONS
'ID            : 003
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/29/2009
'************************************************************************************
'************************************************************************************

Function ABANDONMENT_OPTION_FUNC(ByVal PV_CONT_PROJECT As Double, _
ByVal VOLATILITY_CASH_FLOW As Double, _
ByVal REMAINING_LIFE_PROJECT As Double, _
ByVal VALUE_RECEIVED As Double, _
ByVal TENOR As Double, _
ByVal RISK_FREE As Double, _
Optional ByVal CND_TYPE As Integer = 0)


'PV_CONT_PROJECT = Enter the present value of the cashflows from
'continuing with project.

'VOLATILITY_CASH_FLOW = Enter the annualized standard deviation in
'ln(present value of CF)

'REMAINING_LIFE_PROJECT = This is the remaining life of the project.
'You might not have the power to abandon over the entire life.

'VALUE_RECEIVED = This is the expected net proceeds, if
'the project is abandoned.

'TENOR = Enter the number of years for which the abandonment option holds.

'RISK_FREE = Enter the riskless rate that corresponds to the option lifetime

Dim D1_VAL As Double
Dim D2_VAL As Double

Dim SPOT As Double
Dim STRIKE As Double
Dim EXPIRATION As Double
Dim RATE As Double
Dim DIVD As Double
Dim VOLATILITY As Double

On Error GoTo ERROR_LABEL

SPOT = PV_CONT_PROJECT
STRIKE = VALUE_RECEIVED
EXPIRATION = TENOR
RATE = RISK_FREE
DIVD = (1 / REMAINING_LIFE_PROJECT)
VOLATILITY = VOLATILITY_CASH_FLOW

D1_VAL = (Log(SPOT / STRIKE) + ((RATE - DIVD) + VOLATILITY ^ 2 / 2) * EXPIRATION) / (VOLATILITY * Sqr(EXPIRATION))
D2_VAL = D1_VAL - VOLATILITY * Sqr(EXPIRATION)

ABANDONMENT_OPTION_FUNC = (Exp((0 - DIVD) * EXPIRATION)) * SPOT * CND_FUNC(D1_VAL, CND_TYPE) - STRIKE * (Exp((0 - RATE) * EXPIRATION)) * CND_FUNC(D2_VAL, CND_TYPE) - (Exp((0 - DIVD) * EXPIRATION)) * SPOT + STRIKE * (Exp((0 - RATE) * EXPIRATION))

Exit Function
ERROR_LABEL:
ABANDONMENT_OPTION_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : DELAY_OPTION_FUNC

'DESCRIPTION   :
'This program calculates the value of the option to delay
'making an investment.
'1. The firm has exclusive rights to the project for a fixed
'period. If it does not have exclusive rights in a competitive
'sector, the project will be taken be a competing firm as soon
'as it becomes a value-creating project. In other words, the
'option will be exercised by someone else as soon as S>K.

'2. There have to be factors that will cause the present value
'of the cash flows from taking the project (eg. technological
'or market shifts) to vary across time. If there is no variance
'in the present value of the cash flows, there can be no value
'to the option.

'LIBRARY       : DERIVATIVES
'GROUP         : REAL OPTIONS
'ID            : 004
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/29/2009
'************************************************************************************
'************************************************************************************

Function DELAY_OPTION_FUNC(ByVal PV_INV_PROJECT As Double, _
ByVal VOLATILITY_CASH_FLOW As Double, _
ByVal INITIAL_INVST As Double, _
ByVal TENOR As Double, _
ByVal RISK_FREE As Double, _
Optional ByVal CND_TYPE As Integer = 0)

'PV_INV_PROJECT = Enter the present value of cash flows from investing in
'the project today, not including the initial investment.

'VOLATILITY_CASH_FLOW = Enter the standard deviation in the expected present value.
'This can be estimated either from a simulation or by looking at industry
'averages.

'INITIAL_INVST = Enter the investment needed to take this project today.

'TENOR = Enter the number of years for which you have  rights to this project.
'(if you do not have exclusive rights, enter the number of years for which
'you will have a significant competitive advantage).

'RISK_FREE = Enter the riskless rate that corresponds to the option lifetime

Dim D1_VAL As Double
Dim D2_VAL As Double

Dim SPOT As Double
Dim STRIKE As Double
Dim EXPIRATION As Double
Dim RATE As Double
Dim DIVD As Double
Dim VOLATILITY As Double

On Error GoTo ERROR_LABEL

SPOT = PV_INV_PROJECT
STRIKE = INITIAL_INVST
EXPIRATION = TENOR
RATE = RISK_FREE
DIVD = (1 / TENOR)
VOLATILITY = VOLATILITY_CASH_FLOW

D1_VAL = (Log(SPOT / STRIKE) + ((RATE - DIVD) + VOLATILITY ^ 2 / 2) * EXPIRATION) / (VOLATILITY * Sqr(EXPIRATION))
D2_VAL = D1_VAL - VOLATILITY * Sqr(EXPIRATION)

DELAY_OPTION_FUNC = (Exp((0 - DIVD) * EXPIRATION)) * SPOT * CND_FUNC(D1_VAL, CND_TYPE) - STRIKE * (Exp((0 - RATE) * EXPIRATION)) * CND_FUNC(D2_VAL, CND_TYPE)

Exit Function
ERROR_LABEL:
DELAY_OPTION_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : EXPANSION_OPTION_FUNC
'DESCRIPTION   : This program calculates the value of an expansion option
'The option to expand a project adds value to the current
'project if and only if the following conditions are met:

'1. The current project has to be taken in order for the
'expansion to be viable later on. In other words, if the
'firm can take the expanded version of the project later
'without taking the current project, it is not appropriate
'to credit the current project with the value of this
'option. In real world projects, the current project may
'provide either the information that is necessary to make
'the expansion decision, or the brand name visibility and
'technical skill that is required for the expansion to work.

'2. There have to be factors that will cause the present
'value of the cash flows from expansion to vary across
'time. If there is no variance in the present value of
'the cash flows, there can be no value to the option.

'LIBRARY       : DERIVATIVES
'GROUP         : REAL OPTIONS
'ID            : 005
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/29/2009
'************************************************************************************
'************************************************************************************


Function EXPANSION_OPTION_FUNC(ByVal PV_EXPANSION As Double, _
ByVal VOLATILITY_CASH_FLOW As Double, _
ByVal INITIAL_INVST As Double, _
ByVal TENOR As Double, _
ByVal COST_WAITING As Double, _
ByVal RISK_FREE As Double, _
Optional ByVal CND_TYPE As Integer = 0)


'PV_EXPANSION = This is your estimate of the present value of the cash
'flows that will accrue from expansion, as estimated today.

'VOLATILITY_CASH_FLOW = This can either be the standard deviation from a capital
'budgeting simulation, or the industry average standard deviation in firm value.

'INITIAL_INVST = This is your assessment of the expected cost of expansion.

'TENOR = This is the number of years for which the firm will have rights to
'the expansion option. If the firm does not have exclusive rights, this is
'the number of years for which the firm will have a significant competitive
'advantage.

'COST_WAITING = This measures the cashflow (as a percent of the present value
'of the expansion potential) that will be foregone by waiting a year once
'expansion becomes viable (present value > initial investment).

'RISK_FREE = Enter the riskless rate that corresponds to the option lifetime

Dim D1_VAL As Double
Dim D2_VAL As Double

Dim SPOT As Double
Dim STRIKE As Double
Dim EXPIRATION As Double
Dim RATE As Double
Dim DIVD As Double
Dim VOLATILITY As Double

On Error GoTo ERROR_LABEL

SPOT = PV_EXPANSION
STRIKE = INITIAL_INVST
EXPIRATION = TENOR
RATE = RISK_FREE
DIVD = COST_WAITING '--> COST OF DELAY
VOLATILITY = VOLATILITY_CASH_FLOW

D1_VAL = (Log(SPOT / STRIKE) + ((RATE - DIVD) + VOLATILITY ^ 2 / 2) * EXPIRATION) / (VOLATILITY * Sqr(EXPIRATION))
D2_VAL = D1_VAL - VOLATILITY * Sqr(EXPIRATION)

EXPANSION_OPTION_FUNC = (Exp((0 - DIVD) * EXPIRATION)) * SPOT * CND_FUNC(D1_VAL, CND_TYPE) - STRIKE * (Exp((0 - RATE) * EXPIRATION)) * CND_FUNC(D2_VAL, CND_TYPE)

Exit Function
ERROR_LABEL:
EXPANSION_OPTION_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : FLEXIBILITY_OPTION_FUNC

'DESCRIPTION   : This program calculates the value of financial flexibility
'on an annualized basis. It can be used to determine if firms should
'maintain excess debt capacity.

'When making financial decisions, managers consider the effects such
'decisions will have on their capacity to take new projects or meet
'unanticipated contingencies in future periods. Practically, this
'translates into firms maintaining excess debt capacity or larger cash
'balances than are warranted by current needs, to meet unexpected future
'requirements. While maintaining this financing flexibility has value to
'firms, it also has a cost; the large cash balances earn low returns and
'excess debt capacity implies that the firm is giving up some value and
'has a higher cost of capital.

'The value of flexibility can be analyzed using the option pricing
'framework; a firm maintains large cash balances and excess debt capacity
'in order to have the option to take projects that might arise in the
'future. The value of this option will depend upon two key variables:

'1. Quality of the Firm's Projects: It is the excess return that the
'firm earns on its projects that provides the value to flexibility.
'Other things remaining equal, firms operating in businesses where
'projects earn substantially higher returns than their hurdle rates
'should value flexibility more than those that operate in stable
'businesses where excess returns are small.

'2. Uncertainty about Future Projects: If flexibility is viewed as an
'option, its value will increase when there is greater uncertainty about
'future projects; thus, firms with predictable capital expenditures should
'value flexibility less than those with high variability in capital
'expenditures.

'This option framework would imply that firms such as Berkshire, which earn
'large excess returns on their projects and face more uncertainty about future
'investment needs, can justify holding large cash balances and excess debt
'capacity, whereas a firm such as Chrysler, with much smaller excess returns
'and more predictable investment needs, should hold a much smaller cash balance
'and less excess debt. In fact, the value of flexibility can be calculated as a
'percentage of firm value, with the following inputs for the option pricing model.

'S = Annual Net Capital Expenditures as percent of Firm Value (1 + Excess Return)
'k = Annual Net Capital Expenditures as percent of Firm Value
'Tenor = 1 year
'Var = Variance in ln(Net Capital Expenditures)
'y = Annual Cost of Holding Cash or Maintaining Excess
'Debt Capacity as % of Firm Value
'To illustrate, assume that a firm which earns 18% on its projects has a
'cost of capital of 13%, and that net capital expenditures are 10% of firm
'value; the variance in ln(net capital expenditures) is 0.04. Also assume
'that the firm could have a cost of capital of 12% if it used its excess
'debt capacity. The value of flexibility as a percentage of firm value can
'be estimated as follows:

'S = 10% (1.05) = 10.50% [Excess Return = 18% - 13% = 5%]
'k = 10
't = 1 year
'Var = 0.04
'y = 13 - 12 = 1

'Based on these inputs and a riskless rate of 5%, the value of flexibility
'is 1.31% of firm value.

'LIBRARY       : DERIVATIVES
'GROUP         : REAL OPTIONS
'ID            : 006
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/29/2009
'************************************************************************************
'************************************************************************************


Function FLEXIBILITY_OPTION_FUNC(ByVal REINV_NEED As Double, _
ByVal VOLATILITY_REINV_NEED As Double, _
ByVal NEED_WO_FLEX As Double, _
ByVal NEED_WI_FLEX As Double, _
ByVal COST_CAPITAL As Double, _
ByVal RETURN_CAPITAL As Double, _
ByVal EXPIRATION As Double, _
ByVal RISK_FREE As Double, _
Optional ByVal CND_TYPE As Integer = 0)


'REINV_NEED = Use the historical average of (Net Cap Ex including acquisitions
'+ change in Working Capital)/Market Value of Firm

'VOLATILITY_REINV_NEED = Enter the standard deviation in the
'ln(reinvestment/value ratio)

'NEED_WI_FLEX = If the firm does not want to use external financing,
'this will be the ratio of internal funds available, after debt payments,
'to the value of the firm. FCFE/Value of the Firm

'NEED_WO_FLEX = This is the maximum reinvestment needs, as a percent of
'firm value, that can be financed with flexibility.

'COST_CAPITAL = Enter the firm's current cost of capital.

'RETURN_CAPITAL = This should be the expected marginal return on capital
'on future projects. You can use the firm's current return on capital or
'the industry average, as an estimate.

'RISK_FREE = Enter the riskless rate that corresponds to the option lifetime
'Value of Call of financial flexibility

Dim D1_MIN_VAL As Double
Dim D2_MIN_VAL As Double

Dim D1_MAX_VAL As Double
Dim D2_MAX_VAL As Double

Dim SPOT As Double
Dim STRIKE As Double
Dim RATE As Double
Dim DIVD As Double
Dim VOLATILITY As Double

Dim EXCESS_RETURN As Double
Dim MAX_FLEXIBILITY As Double

On Error GoTo ERROR_LABEL

SPOT = REINV_NEED
STRIKE = NEED_WO_FLEX
RATE = RISK_FREE
DIVD = 0 '--> COST OF DELAY
VOLATILITY = VOLATILITY_REINV_NEED

EXCESS_RETURN = RETURN_CAPITAL - COST_CAPITAL
MAX_FLEXIBILITY = NEED_WI_FLEX

D1_MIN_VAL = (Log(SPOT / STRIKE) + (RISK_FREE - DIVD + (VOLATILITY ^ 2 / 2)) * EXPIRATION) / ((VOLATILITY) * (EXPIRATION ^ 0.5))
D2_MIN_VAL = D1_MIN_VAL - VOLATILITY * Sqr(EXPIRATION)

D1_MAX_VAL = (Log(SPOT / MAX_FLEXIBILITY) + (RISK_FREE - DIVD + (VOLATILITY ^ 2 / 2)) * EXPIRATION) / ((VOLATILITY) * (EXPIRATION ^ 0.5))
D2_MAX_VAL = D1_MAX_VAL - VOLATILITY * Sqr(EXPIRATION)

FLEXIBILITY_OPTION_FUNC = (((Exp((0 - DIVD) * EXPIRATION)) * SPOT * CND_FUNC(D1_MIN_VAL, CND_TYPE) - STRIKE * (Exp((0 - RISK_FREE) * EXPIRATION)) * CND_FUNC(D2_MIN_VAL, CND_TYPE)) - ((Exp((0 - DIVD) * EXPIRATION)) * SPOT * CND_FUNC(D1_MAX_VAL, CND_TYPE) - MAX_FLEXIBILITY * (Exp((0 - RISK_FREE) * EXPIRATION)) * CND_FUNC(D2_MAX_VAL, CND_TYPE))) * EXCESS_RETURN / COST_CAPITAL _
'This is the annual value of financial flexibility, as a percent of firm value. It should be compared to the cost of maintaining this flexibility

Exit Function
ERROR_LABEL:
FLEXIBILITY_OPTION_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : NATURAL_RESOURCES_OPTION_FUNC
'DESCRIPTION   : This program calculates the value of the option to delay
'making an investment.
'LIBRARY       : DERIVATIVES
'GROUP         : REAL OPTIONS
'ID            : 007
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/29/2009
'************************************************************************************
'************************************************************************************

Function NATURAL_RESOURCES_OPTION_FUNC(ByVal ESTIMATED_RESERVE As Double, _
ByVal PRICE_UNIT As Double, _
ByVal COST_UNIT As Double, _
ByVal VOLATILITY_PRICE As Double, _
ByVal CASH_FLOW As Double, _
ByVal DEVELOP_RESOURCE As Double, _
ByVal RELIQUISHED_TENOR As Double, _
ByVal RISK_FREE As Double, _
Optional ByVal CND_TYPE As Integer = 0)

'ESTIMATED_RESERVE = This is the estimated quantity (in barrels, ounces, tonnes)
'of the resource in the reserve.

'PRICE_UNIT = This is the current price per unit of the resource.

'COST_UNIT = This is the cost associated with extracting each unit of the
'resource.

'VOLATILITY_PRICE = Enter the standard deviation in the price of the natural
'resource (ln).

'CASH_FLOW = This is the expected after-tax cash flow each year after the
'resource is developed.

'DEVELOP_RESOURCE = This is the up-front development cost to make the
'undeveloped reserve into a developed one

'RELIQUISHED_TENOR = Enter when the rights to resource will be relinquished

'RISK_FREE = Enter the riskless rate that corresponds to the option lifetime

Dim D1_VAL As Double
Dim D2_VAL As Double

Dim SPOT As Double
Dim STRIKE As Double
Dim EXPIRATION As Double
Dim RATE As Double
Dim DIVD As Double
Dim VOLATILITY As Double

On Error GoTo ERROR_LABEL

SPOT = ESTIMATED_RESERVE * (PRICE_UNIT - COST_UNIT)
STRIKE = DEVELOP_RESOURCE
EXPIRATION = RELIQUISHED_TENOR
RATE = RISK_FREE
DIVD = CASH_FLOW / SPOT
VOLATILITY = VOLATILITY_PRICE

D1_VAL = (Log(SPOT / STRIKE) + ((RATE - DIVD) + VOLATILITY ^ 2 / 2) * EXPIRATION) / (VOLATILITY * Sqr(EXPIRATION))
D2_VAL = D1_VAL - VOLATILITY * Sqr(EXPIRATION)

NATURAL_RESOURCES_OPTION_FUNC = (Exp((0 - DIVD) * EXPIRATION)) * SPOT * CND_FUNC(D1_VAL, CND_TYPE) - STRIKE * (Exp((0 - RATE) * EXPIRATION)) * CND_FUNC(D2_VAL, CND_TYPE) 'Value of the natural resource option

Exit Function
ERROR_LABEL:
NATURAL_RESOURCES_OPTION_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : PATENT_OPTION_FUNC
'DESCRIPTION   : This program calculates the value of the option to delay
'making an investment.
'LIBRARY       : DERIVATIVES
'GROUP         : REAL OPTIONS
'ID            : 008
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/29/2009
'************************************************************************************
'************************************************************************************

Function PATENT_OPTION_FUNC(ByVal PV_CASH_FLOW As Double, _
ByVal VOLATILITY_CASH_FLOW As Double, _
ByVal DEVELOPING_COST As Double, _
ByVal LIFE_PATENT As Double, _
ByVal COST_DELAYING As Double, _
ByVal RISK_FREE As Double, _
Optional ByVal CND_TYPE As Integer = 0)

'PV_CASH_FLOW = Estimate the present value of the expected cash flows from
'developing the patent now, not counting the initial development cost.

'VOLATILITY_CASH_FLOW = This can be best obtained by looking at the standard
'deviations in firm value of other firms in this business, or by running
'capital budgeting simulations.

'DEVELOPING_COST = This is the expected cost associated with converting
'the patent into a commercial product.

'LIFE_PATENT = Enter the remaining number of years in the patent / project
'rights

'COST_DELAYING = This is the cash flow that will be lost because of not
'developing stated as a percent of the present value of the net cash flows.
'As a default, we will assume that you if you do not invest in the project
'once it becomes viable, you will lose one year of protection and that
'your cash flows will decline proportionately (1/remaining life of the patent)

'RISK_FREE = Enter the riskless rate that corresponds to the option lifetime

Dim D1_VAL As Double
Dim D2_VAL As Double

Dim SPOT As Double
Dim STRIKE As Double
Dim EXPIRATION As Double
Dim RATE As Double
Dim DIVD As Double
Dim VOLATILITY As Double

On Error GoTo ERROR_LABEL

SPOT = PV_CASH_FLOW
STRIKE = DEVELOPING_COST

EXPIRATION = LIFE_PATENT

If COST_DELAYING = 0 Then
    DIVD = 1 / LIFE_PATENT
Else
    DIVD = COST_DELAYING
End If

VOLATILITY = VOLATILITY_CASH_FLOW
RATE = RISK_FREE

D1_VAL = (Log(SPOT / STRIKE) + ((RATE - DIVD) + VOLATILITY ^ 2 / 2) * EXPIRATION) / (VOLATILITY * Sqr(EXPIRATION))
D2_VAL = D1_VAL - VOLATILITY * Sqr(EXPIRATION)

PATENT_OPTION_FUNC = (Exp((0 - DIVD) * EXPIRATION)) * SPOT * CND_FUNC(D1_VAL, CND_TYPE) - STRIKE * (Exp((0 - RATE) * EXPIRATION)) * CND_FUNC(D2_VAL, CND_TYPE)

Exit Function
ERROR_LABEL:
PATENT_OPTION_FUNC = Err.number
End Function
