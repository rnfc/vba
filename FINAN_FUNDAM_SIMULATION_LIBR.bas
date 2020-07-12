Attribute VB_Name = "FINAN_FUNDAM_SIMULATION_LIBR"


Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.

'************************************************************************************
'************************************************************************************
'FUNCTION      : INTERNET_COMPANIES_VALUATION_FUNC

'DESCRIPTION   : The valuation of Internet companies is a subject of much
'discussion in the financial press and among financial economists.
'The previous function is based fundamentally on assumptions about the
'expected growth rate of revenues and on expectations about the cost structure
'of the company. Because these expectations are likely to change continuously
'as new information becomes available, the model generates company values and
'stock prices that are highly volatile. The mode gives a systematic way to think
'about the drivers of value of Internet companies, however, and directs attention
'to the parameters that are most important in the valuation.

'To implement the model, we had to make many assumptions about possible future
'financing, about future cash distributions to shareholders and bondholders,
'about the horizon of the estimation, and so on. Alternative assumptions are
'possible and easily incorporated in the analysis. Potential users of a model
'such as the one presented here would need a deep knowledfe of the company and
'its industry in order to make reasonable assumptions.

'We conclude that, depending on the parameters chosen and given high enough
'growth rates of revenues, the value of an Internet stock may be rational.
'Even when the chance that a company may go bankrupt is real, if the initial
'growth rates are sufficiently high and if there is enough volatility in this
'growth over time, valuations can be what would otherwise appear to be
'unbelievably high. In addition, we found the valuation has great sensitivity to
'initial conditions and exact specification of the parameters. This finding is
'consistent with observations that the returns of Internet stock have been
'strikingly volatile.

'LIBRARY       : FUNDAMENTAL
'GROUP         : SIMULATION
'ID            : 001



'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'REFERENCES    : Schwartz, Eduardo S. Moon, Mark. Rational Pricing of
'Internet Companies. Financial Analysts Journal May/June 2000, p. 62-74.
'http://dse.univr.it/safe/Workshops/PhD/2003/Paper2-Schwartz.pdf
'http://dse.univr.it/safe/Workshops/PhD/2003/RealOptions.htm
'http://pages.stern.nyu.edu/~adamodar/pdfiles/valrisk/ch8.pdf
'http://www.ucm.es/centros/cont/descargas/documento15894.pdf
'************************************************************************************
'************************************************************************************

Function INTERNET_COMPANIES_VALUATION_FUNC(ByRef REVENUES_RNG As Variant, _
ByVal TENOR As Double, _
ByVal EBITDA_MULT As Double, _
ByVal CASH_BALANCE As Double, _
ByVal LOSS_FORWARD As Double, _
ByVal COGS_PERCENT As Double, _
ByVal VC_COST_PERCENT As Double, _
ByVal FIXED_COST As Double, _
ByVal RISK_FREE As Double, _
ByVal CORP_TAX_RATE As Double, _
Optional ByVal THRESHOLD As Double = 0.0000000000001, _
Optional ByVal COUNT_BASIS As Integer = 1)

'RISK_FREE --> Discount Rate to discount Cash Flows

'************************************************************************************

'Valuation of High Growth Companies
'This is a pure implementation of a model described by Schwartz and Moon 2000 in their
'paper "Rationale Pricing of Internet Companies".

'The tech bubble has burst in the meantime but the approach is very illustrative (though
'not too practical) for any kind of high growth company.

'Schwartz/Moon propose a simulation approach with both revenue and growth of revenue
'being stochastic processes whose volatility and mean approach some long-term equilibrium.
'After all, such high growth companies often exhibit growth rates that cannot possibly be
'maintained in the long-run.

'In many instances the company will not generate enough funds to finance its growth strategy
'(initial cash resources are used up) and it will go bankrupt as indicated by a bankruptcy line.

'Reference:
'Schwartz, Eduardo S. Moon, Mark. Rational Pricing of Internet Companies.. Financial Analysts
'Journal May/June 2000, p. 62-74

'************************************************************************************


Dim i As Long
Dim j As Long

Dim SROW As Long
Dim nLOOPS As Long 'nLOOPS
Dim NROWS As Double 'PERIODS
Dim D_VAL As Double

Dim CASH_VAL As Double 'Cash Balance
Dim LOSS_VAL As Double 'Loss Carry Forward
Dim INCOME_VAL As Double
Dim TAX_VAL As Double
Dim EBT_VAL As Double 'Pre-tax Profit
Dim INTEREST_VAL As Double 'Interest Income
Dim EBITDA_VAL As Double
Dim SGA_VAL As Double 'Selling, General & Admin. Expenses
Dim GROSS_VAL As Double 'Gross Profit
Dim COGS_VAL As Double 'Cost of Goods Sold
Dim REVENUE_VAL As Double 'Simulated revenue

Dim TEMP_MATRIX As Variant
Dim DATA_MATRIX As Variant

On Error GoTo ERROR_LABEL

DATA_MATRIX = REVENUES_RNG
SROW = LBound(DATA_MATRIX, 1)
NROWS = UBound(DATA_MATRIX, 1) ' First entry in PERIODS includes the
'initial revenue

nLOOPS = UBound(DATA_MATRIX, 2) '--> Per Period
D_VAL = TENOR / (NROWS - 1)

ReDim TEMP_MATRIX(1 To nLOOPS, 1 To 2)

For j = 1 To nLOOPS
    CASH_VAL = CASH_BALANCE 'Initial Cash Balance
    LOSS_VAL = LOSS_FORWARD 'Initial Loss Carry Forward
    
    For i = SROW + 1 To NROWS 'start at i = SROW + 1 to exclude t=0.
    
        REVENUE_VAL = DATA_MATRIX(i, j)
        COGS_VAL = COGS_PERCENT * REVENUE_VAL
        GROSS_VAL = REVENUE_VAL - COGS_VAL
        SGA_VAL = REVENUE_VAL * VC_COST_PERCENT + FIXED_COST
        EBITDA_VAL = GROSS_VAL - SGA_VAL
        
        INTEREST_VAL = (Exp(RISK_FREE / COUNT_BASIS * D_VAL) - 1) * CASH_VAL
        EBT_VAL = EBITDA_VAL + INTEREST_VAL
        TAX_VAL = -MINIMUM_FUNC(0, (LOSS_VAL - EBT_VAL) * CORP_TAX_RATE)
        INCOME_VAL = EBT_VAL - TAX_VAL
        LOSS_VAL = MAXIMUM_FUNC(LOSS_VAL - INCOME_VAL, 0)
        CASH_VAL = INCOME_VAL + CASH_VAL
        
        If CASH_VAL < THRESHOLD Then
            TEMP_MATRIX(j, 2) = i 'THRESHOLD_FLAG
        End If
        
        If CASH_VAL < 0 Then: Exit For
    Next i
        
    If CASH_VAL > 0 Then '--> YOU CAN CHANGE THIS
        TEMP_MATRIX(j, 1) = (CASH_VAL + (EBITDA_VAL * EBITDA_MULT)) * Exp(-RISK_FREE * (NROWS - 1) / COUNT_BASIS)
        'Present Value of the firm (NROWS - 1 to exclude t = 0)
    Else
        TEMP_MATRIX(j, 1) = 0
    End If
Next j

INTERNET_COMPANIES_VALUATION_FUNC = TEMP_MATRIX

'----------------------------------------------------------------------------------
'NOTES   : Issues arising in the valuation of high-tech companies
'ARTICLE : Issues arising in the valuation of high-tech companies
'AUTHOR  : Tzur Fenigstein

'The balance sheet of a typical high-tech company, especially in its early
'stages of life, does not reflect the real economic value of the assets that
'these companies own because the majority of them are intangibles such as
'knowledge, patents, brand names etc. There are not usually recorded in the
'company's financial records unless they have been purchased from another
'company.
'--------------------Characteristics of High-Tech companies-----------------------

'a.  High losses or small profits in the first few years of operation
'b.  On avg. enjoy high growth rates
'c.  In its early stages can expect that it will grow on avg. much more
'    rapidly than a traditional business or the economy as a whole.
'd.  Extreme uncertainty regarding its future
'i.  the probability of the business failing is high compare to more traditional
'    business
'e.  Minimal tangible assets but a number of intangible assets.
        'i.  Little property, plant and equipment and the majority of
'            its assets are not in the company's balance sheet.
        'ii. Assets are the knowledge and experience in the minds of the
        '    people that comprise the software company.

' Considerations when evaluating high-tech companies

    'a.  The value of any business, whether it is traditional with tons
    '    of tangible assets or high-tech, stems from its ability to generate
    '    cash
    'b.  The basic fundamental principles of investment remain the same.
    '    It doesn't matter if we are investing in a food company or a
    '    software company. ' This means that in both cases we should try to
    '    analyze the basic on the critical success factors based on which a
    '    business can survive and succeed.
    'c.  Due to the difficulty in analyzing and making forecasts for high-tech
    '    businesses we should try to use more than one valuation approach.

' Value drivers that influence a high-tech company's value:

'a.  Business opportunities
    'i.  The intensity of opportunities is stronger as is the pace
    '    at which each opportunity appears
'b.  Unique Value Proposition
    'i.  That the business offers in the market and whether there is
    '    a sustainable competitive advantage.
'c.  Customer Base of the Company
    'i.  The Bigger the customer base the bigger the success prospects
    'of the business
'd.  Competitive advantage
    'i.  Capability of the business to cope with competition
'e.  Management team
'    i.  How much management can deal with foreseen and unforeseen market
'    opportunities and risks.
'f.  General Market Condition
'    i.  This is the most important driver, because this is what is going to
'    drive funding.

'Valuation approaches

'1)  Cost Approach: Establish value based on the cost of producing or
'    replacing an asset. The principle behind this approach is that the
'    fair value of an asset should not exceed the cost of obtaining a
'    substitute asset of comparable features and functionality (replacement
'    cost is the max. amount to pay).

'        a.  This approach is usually appropriate when we are evaluating
'        tangible assets ' Real state.

'        b.  When we are dealing with high-tech companies where one of the
'        characteristics is that they do not have large tangible assets we
'        may find the cost approach mostly inappropriate. As real economic
'        benefits and value of such intangibles are different and often far
'        above the cost invested.

'2)  Market Approach: Is used to estimate value through the analyst of recent
'    sales of comparable companies or assets.

'         a.  This analysis is based on earnings; book value; cash flows; and
'         revenues.

'         b.  For unique high-tech companies this analyze is based on
'         non-financial measures such as number of subscribers in cable
'         companies or number of hits in cases of Internet companies
'         which are typically applied to the appropriate financial indicators
'         or operating indicators of the subject entity to determine a range
'         of values.

   'However, there are some major disadvantages in applying the market
   'approach. Firstly there is always a difficulty in identifying
   'comparable companies or comparable transactions and when we are
   'tying to evaluate specific assets it is almost impossible to identify
   'active markets or relevant prices for comparable assets. Also, as opposed
   'to the income approach, the market approach does not reflect all the unique
   'characteristics of the subject business. Applying the market approach to
   'intangible assets is also very limited as usually there is no truly comparable
   'business.

'3)  Income Approach: This approach measures value by reference to the
'    enterprise's expected future debt free cash flows from business
'    operations.

'a.  This typically involves a projection of income and expense and other
'    sources and uses a cash, the assignment of a terminal or residual value
'    at the end of the projection period that is reasonably consistent with
'    the key assumptions and long-term growth of the business and the
'    determination of an appropriate discount rate that reflects the risk of
'    achieving the projections. When we apply the income approach we try to look,
'    analyse and decide on the business's long-term direction and ignore occasional
'    downturns or upturns which are influenced by general market conditions.

        'i.  Factors that form the basis for expected future financial performance
        '    include historical and projected growth rates, business plans or operating
        '    budgets for the enterprise in question, prevailing relevant business
        '    conditions and industry trends including growth expectations in light of
        '    general market growth, competitive market environment and market position.

        'ii. Typically a 10 year projection period of FCF plus an estimated terminal
        '    value, which represents the value of the business enterprise beyond the
        '    projected period. This is discounted to present value through the
        '    application of a discount rate that reflects the weighted average cost of
        '    capital for the subject enterprise.

        'iii.The PV of aggregate annual FCF plus the TV represents the combined debt
        '    and equity capital or enterprise value of the company.

'b.  The income approach is also the most suitable for evaluating individual
'    intangible assets such as knowledge or other intellectual property.

'c.  The income approach does however have certain disadvantage. It is the
'    hardest approach to apply, as it requires a full financial model that
'    forecasts future cash-flows of the subject business. This is harder to
'    estimate in times of turbulence, especially for high-tech companies.
'    However, using probability based option analysis to calculate the expected
'    cash flows can refine this approach.
'--------------------------------------------------------------------------------------


Exit Function
ERROR_LABEL:
INTERNET_COMPANIES_VALUATION_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : INTERNET_COMPANIES_SIMULATION_FUNC

'DESCRIPTION   : Developing Continuous - Time Model for REVENUE. Consider a
'company with instantaneous rate of revenues (or sales) at time t given by Rt.
'Assume that the dynamics of these revenues are given by the following stochastic
'differential equation: dRt / Rt = u dt + Ot d z1. Where, µt, the drift, is
'the expected rate of growth in revenues and is assumed to follow a
'mean-reverting process with a long-term average drift. O is volatility in
'the rate of revenue growth; and z1 is a random variable that reflects the
'draw from a normal distribution. That is, the initial very high growth
'rates of the company are assumed to converge stochastically to the more
'reasonable and sustainable rate of growth for the industry to
'which the company belongs.

'Probably no recent investment topic elicits stronger feelings that Google stock
'valuation. The skyrocketing valuation of this company have made employees
'(at Googles) millionaires and billionaires, while the actual company's promises
'were not generating significant revenue.

'Personally I believe that Googles stock valuation have been bid upward
'irrationally by traders. Such traders see the current frezy as a spectacular
'example of the market bubble of 2001. These traditionalists fear significant
'negative consequences to the real economy if the bubble bursts. Most of them
'believe that some of the pledge will rapidly grow to dominate and even make
'irrelevant their traditional bricks-and-mortar competitors, such as Yahoo and
'Microsoft.

'In the following function I use real-options theory and modern capital-budgeting
'techniques to the problem of valuing Googles. We formulate the model using a
'technique proposed by Schwartz, Eduardo S. Moon, and Mark in continuous time.
'I use both revenue and growth of revenue being stochastic processes whose
'volatility and mean approach some long-term equilibrium. After compiling all
'the news releases of Googles, and analyzing the accounting data available for
'Googles, I realized that the range of $200-$500 per share cannot possibly be
'maintained in the long-run.

'LIBRARY       : FUNDAMENTAL
'GROUP         : SIMULATION
'ID            : 002

'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'REFERENCES    : Schwartz, Eduardo S. Moon, Mark. Rational Pricing of
'Internet Companies. Financial Analysts Journal May/June 2000, p. 62-74.
'************************************************************************************
'************************************************************************************

Function INTERNET_COMPANIES_SIMULATION_FUNC(ByVal nLOOPS As Long, _
ByVal TENOR As Double, _
ByVal INIT_REVENUE_VAL As Double, _
ByVal GROW_REVENUE_VAL As Double, _
ByVal SIGMA_REVENUE_VAL As Double, _
ByVal SIGMA_GROWTH_REVENUE_VAL As Double, _
ByVal PEARSON As Double, _
ByVal LT_GROWTH_REVENUE_VAL As Double, _
ByVal LT_SIGMA_GROWTH_REVENUE_VAL As Double, _
ByVal ADJ_GROWTH_REVENUE_VAL As Double, _
ByVal ADJ_SIGMA_REVENUE_VAL As Double, _
ByVal ADJ_SIGMA_GROWTH_REVENUE_VAL As Double, _
ByVal MARK_PRICE_REVENUE_VAL As Double, _
ByVal MARK_PRICE_GROWTH_REVENUE_VAL As Double, _
Optional ByVal COUNT_BASIS As Integer = 1, _
Optional ByVal RANDOM_FLAG As Boolean = True, _
Optional ByVal OUTPUT As Integer = 0)

'INIT_REVENUE_VAL --> Initial REV: Observable from current Income Statement.
'GROW_REVENUE_VAL --> Initial expected rate of REV growth

'SIGMA_REVENUE_VAL --> Initial volatility of REV. From past income statements and
'projections of future growth

'SIGMA_GROWTH_REVENUE_VAL --> Initial volatility of expected rates of growth in REV.
'Standard deviation of percentage change in revenues over the recent past

'PEARSON --> Correlation between percentage change in. Inferred from market
'volatility of stock price

'REV and change in expected rate of growth. Estimated from past company or
'cross-sectional data

'LT_GROWTH_REVENUE_VAL --> Long-term rate of growth in REV. Rate of growth in revenues for
'a stable company in the same industry as the company being valued

'LT_SIGMA_GROWTH_REVENUE_VAL --> Long-term volatility of the rate of growth in REV. Volatility
'of percentage changes in revenues for a stable company in the same industry as
'the company being valued

'ADJ_GROWTH_REVENUE_VAL --> Speed of adjustment for rate of growth process. Estimated from
'assumptions about the half-life of the process to Long-term rate of growth
'in revenues

'ADJ_SIGMA_REVENUE_VAL --> Speed of adjustment for the volatility of REV process.
'Estimated from assumptions about the half-life of the process to Long-term
'volatility of the rate of growth in revenues

'ADJ_SIGMA_GROWTH_REVENUE_VAL --> Speed of adjustment for the volatility of the rate of
'REV growth process. Estimated from assumptions about the half-life of the
'process to zero

'MARK_PRICE_REVENUE_VAL --> Market price of risk for REV factor. Obtained from the product
'of the correlation between percentage changes in revenues and return on aggregate
'wealth multiplied by the standard deviation of aggregate wealth.

'MARK_PRICE_GROWTH_REVENUE_VAL --> Market price of risk for the expected rate of growth
'in REV factor: Obtained from the product of the correlation between changes
'in growth rates in revenues and return on aggregate wealth multiplied by the
'standard deviation of aggregate wealth.

Dim i As Long
Dim j As Long

Dim PERIODS As Double

Dim D_VAL As Double 'Time increment for discrete version of model
Dim NRV_VAL As Double
Dim CNRV_VAL As Double

Dim TEMP_MATRIX As Variant
Dim DATA_MATRIX As Variant
Dim RANDOM_VECTOR As Variant

On Error GoTo ERROR_LABEL

PERIODS = TENOR * COUNT_BASIS
D_VAL = TENOR / PERIODS

ReDim TEMP_MATRIX(0 To PERIODS, 1 To 7)
ReDim DATA_MATRIX(0 To PERIODS, 1 To nLOOPS)

TEMP_MATRIX(0, 1) = 0
TEMP_MATRIX(0, 2) = SIGMA_GROWTH_REVENUE_VAL _
'Initial volatility of expected rate of REV growth
TEMP_MATRIX(0, 3) = SIGMA_REVENUE_VAL 'Initial volatility of REV
TEMP_MATRIX(0, 4) = GROW_REVENUE_VAL 'Initial expected rate of REV growth
TEMP_MATRIX(0, 5) = GROW_REVENUE_VAL
TEMP_MATRIX(0, 6) = INIT_REVENUE_VAL 'Initial REV
TEMP_MATRIX(0, 7) = INIT_REVENUE_VAL 'Initial REV
    
If RANDOM_FLAG = True Then: Randomize
    
For j = 1 To nLOOPS
    
    DATA_MATRIX(0, j) = TEMP_MATRIX(0, 7)
    RANDOM_VECTOR = MATRIX_RANDOM_NORMAL_FUNC(PERIODS, 2, 0, 0, 1, 0)

    For i = 1 To PERIODS
        NRV_VAL = RANDOM_VECTOR(i, 1)
        CNRV_VAL = NRV_VAL * PEARSON + Sqr(1 - PEARSON ^ 2) * RANDOM_VECTOR(i, 2)
        TEMP_MATRIX(i, 1) = TEMP_MATRIX(i - 1, 1) + D_VAL
        TEMP_MATRIX(i, 2) = SIGMA_GROWTH_REVENUE_VAL * Exp(-ADJ_SIGMA_GROWTH_REVENUE_VAL * i) 'VOLATILITY GROWTH REV
        TEMP_MATRIX(i, 3) = SIGMA_REVENUE_VAL * Exp(-ADJ_SIGMA_REVENUE_VAL * i) + LT_SIGMA_GROWTH_REVENUE_VAL * (1 - Exp(-ADJ_SIGMA_REVENUE_VAL * i)) 'VOLATILITY REV
        TEMP_MATRIX(i, 4) = Exp(-ADJ_GROWTH_REVENUE_VAL * D_VAL) * TEMP_MATRIX(i - 1, 4) + (1 - Exp(-ADJ_GROWTH_REVENUE_VAL * D_VAL)) * (LT_GROWTH_REVENUE_VAL - (MARK_PRICE_GROWTH_REVENUE_VAL * TEMP_MATRIX(i - 1, 2) / ADJ_GROWTH_REVENUE_VAL)) 'DRIFT MEAN REV
        TEMP_MATRIX(i, 5) = Exp(-ADJ_GROWTH_REVENUE_VAL * D_VAL) * TEMP_MATRIX(i - 1, 5) + (1 - Exp(-ADJ_GROWTH_REVENUE_VAL * D_VAL)) * (LT_GROWTH_REVENUE_VAL - (MARK_PRICE_GROWTH_REVENUE_VAL * TEMP_MATRIX(i - 1, 2) / ADJ_SIGMA_REVENUE_VAL)) + Sqr((1 - Exp(-2 * ADJ_GROWTH_REVENUE_VAL * D_VAL)) / (2 * ADJ_GROWTH_REVENUE_VAL)) * TEMP_MATRIX(i - 1, 2) * CNRV_VAL 'MEAN REV
'--------------------------------------------------------------------------------------
        TEMP_MATRIX(i, 6) = TEMP_MATRIX(i - 1, 6) * Exp((TEMP_MATRIX(i - 1, 4) - MARK_PRICE_REVENUE_VAL * TEMP_MATRIX(i - 1, 3) - (TEMP_MATRIX(i - 1, 3) ^ 2 / 2)) * D_VAL) 'DRIFT REV
        TEMP_MATRIX(i, 7) = DATA_MATRIX(i - 1, j) * Exp((TEMP_MATRIX(i - 1, 5) - MARK_PRICE_REVENUE_VAL * TEMP_MATRIX(i - 1, 3) - (TEMP_MATRIX(i - 1, 3) ^ 2 / 2)) * D_VAL + TEMP_MATRIX(i - 1, 3) * Sqr(D_VAL) * NRV_VAL) 'SIMULATED REV
'----------------------------------------------------------------------------------
        DATA_MATRIX(i, j) = TEMP_MATRIX(i, 7)
'----------------------------------------------------------------------------------
    Next i
Next j

Select Case OUTPUT
Case 0
    INTERNET_COMPANIES_SIMULATION_FUNC = DATA_MATRIX
Case Else
    INTERNET_COMPANIES_SIMULATION_FUNC = TEMP_MATRIX
End Select

Exit Function
ERROR_LABEL:
INTERNET_COMPANIES_SIMULATION_FUNC = Err.number
End Function

'- Athanassakos (2007), "Valuying Internet Ventures", Journal of Business Valuation and Economic Loss Analysis, Vol. 2, Issue 1, Article 2.
'- Berger, Phillip G., Ofek, Eli and Swary, Itzhak (1996), "Investor Valuation of the Abandonment Option", Journal of Financial Economics 42, 1996, 257-287.
'- Blakely and Estrada (1999), "The Pricing of Internet Stocks", IESE.
'- Braunschweig (2003), "The New Internet Gamble", VCJ, December 2003.
'- Bartov, Mohantam, and Seethamraju (2002), "Valuation of Internet Stocks - An IPO Perspective", Journal of Accounting Research, Vol.40 N°2 May 2002.
'- Cao and Summer (2006), "Debating Google: Is this high-flier ridiculously overvalued?", Morningstar StockInvestor.
'- Cunningham, Christopher R. (2006), "House Price Uncertainty, Timing of Development and Vacant Land Prices: Evidence of Real Options in Seattle", Journal of Urban Economics 59, Issue 1, 2006, 1-31.
'- Demers, E. and B. Lev, (2001), "A Rude Awakening: Internet Shakeout in 2000", forthcoming in Review of Accounting Studies.
'- Estrada (2001), "The Pricing of Internet Stocks (II)", IESE.
'- Estrada (2002), "Adjusting P/E Ratios by Growth and Risk: The PERG Ratio", IESE.
'- Corr (2006), "Boom, Bust, Boom: Internet Company Valuations – From Netscape to Google", The Financier, Vol. 13/14, 2006-2007.
'- Geyskens, Gielens, et Dekimpe (2003), "Comment le marché évalue-t-il l'ajout d'un canal de distribution sur Internet ?", Recherche et Applications en Marketing, Vol. 18, n°2/2003.
'- Hand, J., (2000a), "Profits, Losses, and the Pricing of Internet Stocks", Working Paper, Kenan-Flagler Business School, UNC Chapel Hill.
'- Hand, J., (2000b), "The Role of Economic Fundamentals, Web Traffic, and Supply and Demand in the Pricing of U.S. Internet Stocks", Working Paper, Kenan-Flagler Business School, UNC Chapel Hill.
'- Higson and Briginshaw (2000), "Valuing Internet Businesses", Business Strategy Review, 2000, Volume 11 Issue 1, pp 10-20.
'- Jorion, Talmor (2001), "Value Relevance of Financial and Nonfinancial Information in Emerging Industries: The Changing Role of Web Traffic Data", Working Paper, IIBRTAU, CRITO.
'- Keating (2000), "Discussion of The Eyeballs Have It: Searching for the Value of Internet Stocks", Journal of Accounting Research, Vol. 38 Supplement 2000.
'- Kozberg (2001), "The Value Drivers of Internet Stocks: A Business Models Approach", Working Paper, Baruch College.
'- Moel, Alberto and Tufano, Peter (2002), "When Are Real Options Exercised? An Empirical Study of Mine Closings." Review of Financial Studies 15, Issue 1, 2002, 35-64.
'- Noe and Parker (2005), "Winner Take All: Competition, Strategy and the Structure of Returns in the Internet Economy", Journal of Economics & Management Strategy, Volume 14, Number 1, Spring 2005, 141–164.
'- Paddock, James L., Siegel, Daniel R. and Smith, James L. (1998), "Option Valuation of Claims on Real Assets: The Case of Offshore Petroleum Leases", Quarterly Journal of Economics 103, Issue 3, 1988, 479-508.
'- Perotti, Enrico, and Rossetto, Silvia (2000), "Internet Portals as Portfolio of Entry Options", November 2000, University of Amsterdam and CEPR.
'- Pinches, George E., Narayanan, V. K. and Kelm, Katherine M. (1996), "How the Market Values the Different Stages of Corporate R&D - Initiation, Progress and Commercialization", Journal of Applied Corporate Finance 9, Issue 1, Spring 1996, 60-69.
'- Quigg, Laura (1993), "Empirical Testing of Real Option-Pricing Models", Journal of Finance 48, Issue 2, 1993, 621-640.
'- Rajgopal, S., S. Kotha and M. Venkatachalam, (2000), "The Relevance of Web Traffic for Internet Stock Prices", Working Paper, University of Washington and Stanford University Berkeley.
'- Sanders and Boivie (2003), "Sorting Things Out: Valuation of new firms in uncertain markets", Strategic Management Journal, 25: 167-184.
'- Schwartz, Eduardo S., and Moon, Mark (2000), "Rational Valuation of Internet Companies", Financial Analysts Journal 56, Issue 3, 2000, 62-75.
'- Schwartz, Eduardo S., and Moon, Mark (2001), "Rational Pricing of Internet Companies Revisited", The Financial Review 36, Issue 4, 2001, 7-25.
'- Tan (2000), "Real Options Valuation of eBusiness".
'- Trueman, Wong, and Zhang (2000), "The Eyeballs Have It: Searching for the Value in Internet Stocks", Journal of Accouting Research, Vol. 38 Supplement 2000.
