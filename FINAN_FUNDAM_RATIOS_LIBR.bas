Attribute VB_Name = "FINAN_FUNDAM_RATIOS_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.

'************************************************************************************
'************************************************************************************
'FUNCTION      : EXTERNAL_FINANCING_FUNC

'DESCRIPTION   : Function for calculating External Financing Required
'((Assets that change with Sales / Sales) x Projected Sales Increase) -
'((Liabilities that change with Sales / Sales) x Projected Sales Increase) -
'(Profit Margin x Projected Sales x Retention Ratio)

'LIBRARY       : FUNDAMENTAL
'GROUP         : RATIOS
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 03/02/2008
'************************************************************************************
'************************************************************************************

Function EXTERNAL_FINANCING_FUNC(ByVal CURRENT_REVENUES As Double, _
ByVal PROJECTED_REVENUES As Double, _
ByVal CURRENT_NET_INCOME As Double, _
ByVal CURRENT_DIVD_PREFERRED As Double, _
ByVal CURRENT_DIVD_COMMON As Double, _
ByVal CURRENT_ASSETS As Double, _
ByVal CURRENT_LIABILITIES As Double)

Dim ASSETS_SALES As Double
Dim DELTA_REVENUES As Double
Dim LIABILITIES_SALES As Double
Dim PROFIT_MARGIN As Double
Dim RETENTION_RATIO As Double

'DIVD PREFERRED: All dividends declared for the
'reporting period on preferred stock.

'DIVD COMMON: All dividends declared and
'appropriated for the reporting period on common stock.

On Error GoTo ERROR_LABEL

ASSETS_SALES = CURRENT_ASSETS / CURRENT_REVENUES
'Select only current assets, but Fixed Assets may also change given movements with Sales.
DELTA_REVENUES = (PROJECTED_REVENUES - CURRENT_REVENUES)
LIABILITIES_SALES = CURRENT_LIABILITIES / CURRENT_REVENUES
'Select only current liabilities, but longterm liabilities may move given changes with sales.
PROFIT_MARGIN = CURRENT_NET_INCOME / CURRENT_REVENUES
RETENTION_RATIO = 1 - (CURRENT_DIVD_COMMON / (CURRENT_NET_INCOME - CURRENT_DIVD_PREFERRED))

EXTERNAL_FINANCING_FUNC = (ASSETS_SALES * DELTA_REVENUES) - _
                          (LIABILITIES_SALES * DELTA_REVENUES) - _
                          (PROFIT_MARGIN * PROJECTED_REVENUES * RETENTION_RATIO)

Exit Function
ERROR_LABEL:
EXTERNAL_FINANCING_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : ROE_RATIO_FUNC
'DESCRIPTION   : RETURN ON EQUITY
'LIBRARY       : FUNDAMENTAL
'GROUP         : RATIOS
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 03/02/2008
'************************************************************************************
'************************************************************************************

Function ROE_RATIO_FUNC(ByVal ROA_VAL As Double, _
ByVal DEBT_RATIO_VAL As Double, _
ByVal INTEREST_VAL As Double, _
ByVal TAX_RATE_VAL As Double)

'DEBT_RATIO_VAL = D / (D + E)
Dim DEBT_EQUITY_VAL As Double '= D / E
On Error GoTo ERROR_LABEL
DEBT_EQUITY_VAL = DEBT_RATIO_VAL / (1 - DEBT_RATIO_VAL)
ROE_RATIO_FUNC = ROA_VAL + DEBT_EQUITY_VAL * (ROA_VAL - INTEREST_VAL * (1 - TAX_RATE_VAL))
Exit Function
ERROR_LABEL:
ROE_RATIO_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : IMPLIED_ROE_RATIO_FUNC
'DESCRIPTION   : IMPLIED RETURN ON EQUITY
'LIBRARY       : FUNDAMENTAL
'GROUP         : RATIOS
'ID            : 002
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 03/02/2008
'************************************************************************************
'************************************************************************************

Function IMPLIED_ROE_RATIO_FUNC(ByVal UNLEVERED_GROWTH_VAL As Double, _
ByVal PAYOUT_VAL As Double, _
ByVal DEBT_RATIO_VAL As Double, _
ByVal INTEREST_VAL As Double, _
ByVal TAX_RATE_VAL As Double)
'DEBT_RATIO_VAL = D / (D + E)
On Error GoTo ERROR_LABEL
IMPLIED_ROE_RATIO_FUNC = LEVERED_GROWTH2_FUNC(UNLEVERED_GROWTH_VAL, PAYOUT_VAL, DEBT_RATIO_VAL, INTEREST_VAL, TAX_RATE_VAL) / (1 - PAYOUT_VAL)
Exit Function
ERROR_LABEL:
IMPLIED_ROE_RATIO_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : IMPLIED_ROA_RATIO_FUNC
'DESCRIPTION   : IMPLIED RETURN ON ASSETS
'LIBRARY       : FUNDAMENTAL
'GROUP         : RATIOS
'ID            : 003
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 03/02/2008
'************************************************************************************
'************************************************************************************

Function IMPLIED_ROA_RATIO_FUNC(ByVal UNLEVERED_GROWTH_VAL As Double, _
ByVal PAYOUT_VAL As Double)
On Error GoTo ERROR_LABEL
IMPLIED_ROA_RATIO_FUNC = UNLEVERED_GROWTH_VAL / (1 - PAYOUT_VAL)
Exit Function
ERROR_LABEL:
IMPLIED_ROA_RATIO_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : RANK_PEG_RATIO_FUNC
'DESCRIPTION   : PEG_MULTIPLE_RANKER (dimensionless stock rating)
'LIBRARY       : FUNDAMENTAL
'GROUP         : RATIOS
'ID            : 004
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 03/02/2008
'************************************************************************************
'************************************************************************************

Function RANK_PEG_RATIO_FUNC(ByVal PEG_RATIO_VAL As Double, _
ByVal SIGMA_VAL As Double, _
Optional ByVal FACTOR_VAL As Double = 10000)
'SIGMA_VAL --> Must Be Annualized
On Error GoTo ERROR_LABEL
RANK_PEG_RATIO_FUNC = PEG_RATIO_VAL * SIGMA_VAL ^ 2 * FACTOR_VAL
Exit Function
ERROR_LABEL:
RANK_PEG_RATIO_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : EPS_GROWTH_FUNC
'DESCRIPTION   : GROWTH_EARNINGS_RANKER
'LIBRARY       : FUNDAMENTAL
'GROUP         : RATIOS
'ID            : 005
'LAST UPDATE   : 03/02/2008
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'************************************************************************************
'************************************************************************************

Function EPS_GROWTH_FUNC(ByVal ENTERPRISE_VAL As Double, _
ByVal CURRENT_EBITDA As Double, _
ByVal EARNINGS_GROWTH As Double, _
ByVal SIGMA_VAL As Double, _
Optional ByVal FACTOR_VAL As Double = 100)
'ENTERPRISE_VAL = Mkt Cap + Debt - Cash/Securities
'EBITDA = Operating Income --> Earnings before Interest, Taxes, Depreciation,
'and amortization
On Error GoTo ERROR_LABEL
EPS_GROWTH_FUNC = FACTOR_VAL * (ENTERPRISE_VAL / CURRENT_EBITDA) * SIGMA_VAL ^ 2 / EARNINGS_GROWTH
Exit Function
ERROR_LABEL:
EPS_GROWTH_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : UNLEVERED_BREAK_EVEN_GROWTH_FUNC
'DESCRIPTION   : BREAK_EVEN_GROWTH: UNLEVERED_GROWTH_VAL = GROWTH_LEVERED
'LIBRARY       : FUNDAMENTAL
'GROUP         : RATIOS
'ID            : 006
'LAST UPDATE   : 03/02/2008
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'************************************************************************************
'************************************************************************************

Function UNLEVERED_BREAK_EVEN_GROWTH_FUNC(ByVal PAYOUT_VAL As Double, _
ByVal INTEREST_VAL As Double, _
ByVal TAX_RATE_VAL As Double)
On Error GoTo ERROR_LABEL
UNLEVERED_BREAK_EVEN_GROWTH_FUNC = (1 - PAYOUT_VAL) * INTEREST_VAL * (1 - TAX_RATE_VAL)
Exit Function
ERROR_LABEL:
UNLEVERED_BREAK_EVEN_GROWTH_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : UNLEVERED_GROWTH_FUNC
'DESCRIPTION   : UNLEVERED_GROWTH_VAL_FIRM
'LIBRARY       : FUNDAMENTAL
'GROUP         : RATIOS
'ID            : 007
'LAST UPDATE   : 03/02/2008
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'************************************************************************************
'************************************************************************************

Function UNLEVERED_GROWTH_FUNC(ByVal ROA_VAL As Double, _
ByVal PAYOUT_VAL As Double)
On Error GoTo ERROR_LABEL
UNLEVERED_GROWTH_FUNC = (1 - PAYOUT_VAL) * ROA_VAL
Exit Function
ERROR_LABEL:
UNLEVERED_GROWTH_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : LEVERED_GROWTH1_FUNC
'DESCRIPTION   : Visualizing Relationship Between Levered and Unlevered
'Growth Rate
'LIBRARY       : FUNDAMENTAL
'GROUP         : RATIOS
'ID            : 008
'LAST UPDATE   : 03/02/2008
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'************************************************************************************
'************************************************************************************

Function LEVERED_GROWTH1_FUNC(ByVal ROA_VAL As Double, _
ByVal PAYOUT_VAL As Double, _
ByVal DEBT_RATIO_VAL As Double, _
ByVal INTEREST_VAL As Double, _
ByVal TAX_RATE_VAL As Double, _
Optional ByVal VERSION As Integer = 0)
'DEBT_RATIO_VAL = D / (D + E)
Dim DEBT_EQUITY_VAL As Double '= D / E
On Error GoTo ERROR_LABEL
DEBT_EQUITY_VAL = DEBT_RATIO_VAL / (1 - DEBT_RATIO_VAL)
Select Case VERSION
Case 0
    LEVERED_GROWTH1_FUNC = (1 - PAYOUT_VAL) * ROE_RATIO_FUNC(ROA_VAL, DEBT_RATIO_VAL, INTEREST_VAL, TAX_RATE_VAL)
Case Else
    LEVERED_GROWTH1_FUNC = (1 + DEBT_EQUITY_VAL) * UNLEVERED_GROWTH_FUNC(ROA_VAL, PAYOUT_VAL) - (DEBT_EQUITY_VAL * (1 - PAYOUT_VAL) * INTEREST_VAL * (1 - TAX_RATE_VAL))
End Select
Exit Function
ERROR_LABEL:
LEVERED_GROWTH1_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : LEVERED_GROWTH2_FUNC
'DESCRIPTION   : Unlevered firm growth as input parameter could for
'example be estimated as a function of expected economic growth.
'LIBRARY       : FUNDAMENTAL
'GROUP         : RATIOS
'ID            : 009
'LAST UPDATE   : 03/02/2008
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'************************************************************************************
'************************************************************************************

Function LEVERED_GROWTH2_FUNC(ByVal UNLEVERED_GROWTH_VAL As Double, _
ByVal PAYOUT_VAL As Double, _
ByVal DEBT_RATIO_VAL As Double, _
ByVal INTEREST_VAL As Double, _
ByVal TAX_RATE_VAL As Double)
'DEBT_RATIO_VAL = D / (D + E)
Dim DEBT_EQUITY_VAL As Double '= D / E
On Error GoTo ERROR_LABEL
DEBT_EQUITY_VAL = DEBT_RATIO_VAL / (1 - DEBT_RATIO_VAL)
LEVERED_GROWTH2_FUNC = (1 + DEBT_EQUITY_VAL) * UNLEVERED_GROWTH_VAL - (DEBT_EQUITY_VAL * (1 - PAYOUT_VAL) * INTEREST_VAL * (1 - TAX_RATE_VAL))
Exit Function
ERROR_LABEL:
LEVERED_GROWTH2_FUNC = Err.number
End Function
