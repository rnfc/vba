Attribute VB_Name = "FINAN_FUNDAM_COMPARAB_LIBR"


Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.

'************************************************************************************
'************************************************************************************
'FUNCTION      : NON_FINANCIAL_INSTITUTIONS_COMPARABLE_FUNC
'DESCRIPTION   :
'LIBRARY       : FUNDAMENTAL
'GROUP         : COMPARABLES
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 08/04/2009
'************************************************************************************
'************************************************************************************

Function NON_FINANCIAL_INSTITUTIONS_COMPARABLE_FUNC(ByRef DATA_RNG As Variant, _
Optional ByVal COUNT_BASIS As Long = 365)

'--------------------------------------------------------------------------------
'Abbreviations and Definitions
'--------------------------------------------------------------------------------
'AR:  Accounts receivable
'CA: Current Assets
'CAPEX :  "Capital expenditures" or "Acquisitions of Plant, Property,
'and Equipment"
'Cash Flow :  Has many definitions and has to be separately defined or
'elaborated upon in a given context.
'CEO Total Compensation:  Sum of salary, bonuses and other compensation
'in millions of dollars. Do not include options.
'Chg: Change
'COG: Cost of Goods
'C&E: Cash & Equivalents
'CL: Current Liabilities
'CY:   Current Year
'DIO: Days Inventory Outstanding
'DPO: Days Payables Outstanding
'DSO: Days Sales Outstanding
'Free Cash Flow :  Cash provided by Operations - dividends paid - Capital
'expenditures
'LTD: Long Term Debt
'MS: Marketable Securities (if any)
'NI: Net Income
'Net Profit :  Same as Net Income or Net Earnings
'NCO: Net Cash from Operations
'PP&E :  Plant, Property, and Equipment.   The total that exists now.
'As distinguished from Acquisitions of Plant Property and Equipment(CAPEX)
'PTI :  Pre tax Income
'PY:   Prior Year
'ROIC: Return on Invested Capital
'STI: Short Term Investments
'STD: Short Term Debt


'From Statement of Income
'10K and Proxy (SEC Schedule 14A) data
'http://edgarscan.pwcglobal.com/servlets/edgarscan


'Accounts Payable = Payables Accounts Receivable = Trade Receivables =
'Receivables
'Additional Paid-in Capital = Capital in Excess of Stated Value =
'Capital Surplus = Paid-in Capital Balance Sheet = Statement of
'Financial Condition = Consolidated Balance Sheets Capital Expenditures
'includes: Acquisition of Property and Equipment Capital Expenditures
'Capitalized Software Costs Cash includes: Cash and Equivalents Marketable
'Securities Short-term Marketable Securities Investment Securities Other
'Securities Short-term Investments Trading Assets Cost of Goods Sold
'includes: Costs of Sales Cost of Revenue Cost of Products Sold Cost of
'Services Sold Costs, Materials, and Production Current Assets includes:
'Cash Accounts Receivable Trade Accounts Receivable Other receivables Loan
'receivable Inventories (includes raw materials, work-in-process,
'semi-finishedgoods, and finished goods) Deferred tax Prepaid income tax
'Prepaid assets Other prepaid expenses and receivables Other current
'assets Current Liabilities includes: Accounts payable Income taxes Current
'portion of long-term debt Accrued liabilities (expenses) Deferred/unearned
'revenue Other current liabilities Earnings = Net Income = Net Profit Income
'Statement = Earnings Statement = Statement of Operations = Profit & Loss
'Statement = Consolidated Statement of Income Inventories = Merchandise
'Inventories Earnings Before Income Taxes = Income (Loss) Before Income
'Taxes = Earnings Before Provision for Income Taxes Earnings per Share =
'Net Income per Share = Net Income per Common Share Long-Term Debt
'includes: Notes/Loans payable Bank line of credit Capital lease obligation
'Preferred stock Convertible notes Net Income = Net Profit = Net Earnings
'Operating Cash Flow includes: Net cash provided by (used in) Operating
'Activities It is positive if "provided by"and negative if "üsed in"
'Operating Activities. Revenues = Sales = Net Sales Shareholder Equity =
'Shareholders' Investment = Stockholders' Equity Short-Term Debt includes:
'Debt Payable Within One Year Current Portion of Long-Term Debt Notes Payable
'Short-term borrowings (Some annual reports don't like to use the word "debt".)

'There are a lot of opportunities to make mistakes in this section. Such as:
'All input must be in Millions of dollars. Most annual reports are in
'thousands. Sign conventions in the Cash Flow Inputs can be either minus
'or plus depending as you will see. Input of capital expenditures or
'dividends are usually input as positive whereas a cash flow statement
'would list them as negative. The input of numbers of shares must be in
'"diluted average weighted outstanding shares",etc.

'DATA_RNG:
'ROW 1: COMPANY
'ROW 2: FISCAL YEAR
'ROW 3: SALES - CURRENT YEAR CY (M); Revenue, Sales, or Net Sales
'ROW 4: SALES - PRIOR YEAR  PY (M)
'ROW 5: COST OF GOODS SOLD - CURRENT YEAR; Or Cost of Services, Cost of
'Products, Cost of Revenue, etc.
'ROW 6: COST OF GOODS SOLD - PRIOR YEAR
'ROW 7: INTEREST EXPENSE
'ROW 8: PRE TAX INCOME (PTI); Or Earnings Before Provision for Income
'Taxes, Or Earnings Before Income Taxes, etc. Or Income before taxes,
'profit before taxes
'ROW 9: INCOME TAXES; Or Provision for Income Taxes, etc.
'ROW 10: NET PROFIT / NET INCOME / NET EARNINGS
'ROW 11: DILUTED AVG WTD NO SHARES CY (M); If not available on statement
'of income, see "selected financial data table" in annual report for
'"diluted average weighted number of shares"
'ROW 12: DILUTED AVG WTD NO SHARES PY (M)
'ROW 13: C&E+SHTM INV+MARKETABLE SEC(CY); The sum of Cash & Equivalents,
'Short Term Investments, and Marketable Securities, etc. (if any)
'ROW 14: C&E+SHTM INV+MARKETABLE SEC(PY)
'ROW 15: ACCOUNTS RECEIVABLE - CY; Or Receivables or trade receivables
'ROW 16: ACCOUNTS RECEIVABLE - PY
'ROW 17: INVENTORIES - CURRENT YEAR; Merchandise Inventories, etc.
'ROW 18: INVENTORIES - PRIOR YEAR
'ROW 19: CURRENT ASSETS
'ROW 20: PROPERTY, PLANT  & EQUIPMENT; Property and Equipment,etc.
'ROW 21: TOTAL ASSETS (CY)
'ROW 22: TOTAL ASSETS (PY)
'ROW 23: SHORT TERM DEBT; Current Portion of Long Term Debt,Short Term
'borrowings,Debt Payable within one year,etc.
'ROW 24: ACCOUNTS PAYABLE
'ROW 25: CURRENT LIABILITIES
'ROW 26: LONG TERM DEBT; Long Term Liabilities, less current portion
'ROW 27: RETAINED EARNINGS; Retained Earnings: When a company starts out
'it only has paid in capital as part of shareholder equity. By this is
'meant the dollars paid to the company by shareholders to obtain shares
'(not via a public market like the NYSE)but like in an IPO. Thus retained
'earnings would be zero. But as the company grows and makes a profit the
'retained earnings will generally grow, unless reduced by paying dividends
'or other expenditures, and especially in mature companies retained earnings
'will become a larger percentage of shareholder equity.
'ROW 28: SHAREHOLDER EQUITY - CURRENT YEAR; Total Stockholder's Equity;
'Total Shareholder's Equity
'ROW 29: SHAREHOLDER EQUITY - PRIOR YEAR
'ROW 30: NET CASH FROM OPERATING ACTIVITIES(CY); The sign to use on the
'input sheet should be positive if the Operating Activity does not have
'parenthesis and the text says cash "provided by" Operating Activities.
'In the rare case that the cash is "used in" Operating Activities and/or
'is in parenthesis, then it is negative.
'ROW 31: NET CASH FROM INVESTING ACTIVITIES; The sign to be input for
'Investing Activities should be negative if the Operating Activity is
'in parenthesis or the text says cash "used in" Investing Activities.
'It is positive if it is not in parenthesis and the text says "provided
'by".
'ROW 32: CAPITAL EXPENDITURES (CAPEX); Or called purchases of Property,
'Plant or Equiptment, etc.Is almost always positive unless more equiptment
'was sold off than purchased. Then it would be negative.
'ROW 33: NET CASH FROM FINANCING ACTIVITIES; Financing Activities are
'positive if not in parenthesis and cash is "provided by" the activity.
'Negative if in parenthesis or "used in" the activity.
'ROW 34: DIVIDENDS PAID; Input dividends as positive regardless of the
'sign used in the annual report.
'ROW 35: TOTAL DEBT (OPTIONAL); Do not use this option. Input a zero.
'Otherwise, the Valueline Total Debt could be input if no other data
'is available for long and short term debt.
'ROW 36: CEO TOTAL COMPENSATION; Go to Proxy Statement (Schedule 14A at
'Edgarscan online) and add up CEO's Salary,Bonus, and All Other compensation.
'Input in Millions of dollars. Do not include long term compensation or
'stock options
'ROW 37: STOCK OPTION GRANTED SHARES(M); Annual report notes will list
'Stock Options outstanding at the close of the fiscal year of the report.
'Input in Millions of shares.
'ROW 38: NET INCOME IF STK OPTIONS EXPENSED($M); "Net Income if Stock
'Options Expensed" (input in millions of dollars). Look in the footnotes
'of the Annual Report or 10K for a section on "stock based compensation"
'that starts out with the "Net Income as reported" (same value as on the
'Income Statement) and then subtracts the "Pro forma employee compensation
'cost of stock based compensation plans." The Net Income if stock options
'are expensed will be labeled "pro forma". The new law is that all must
'be expensed. If the difference is between Net Income Reported and the pro
'forma amount is too great, they have been using your company as a personal
'piggy bank and usually try to buy back stock so the shares are not too
'badly or noticably diluted.

Dim k As Long

Dim NROWS As Long
Dim NCOLUMNS As Long

Dim TEMP_MATRIX As Variant
Dim DATA_MATRIX As Variant

On Error GoTo ERROR_LABEL

DATA_MATRIX = DATA_RNG
NROWS = UBound(DATA_MATRIX, 1)
NCOLUMNS = UBound(DATA_MATRIX, 2)

ReDim TEMP_MATRIX(1 To 54, 1 To NCOLUMNS + 1)

TEMP_MATRIX(1, 1) = "COMPANY"
TEMP_MATRIX(2, 1) = "REPORT DATE"
TEMP_MATRIX(3, 1) = "PRETAX PROFIT MARGIN  | GC: >15%"
TEMP_MATRIX(4, 1) = "NET PROFIT MARGIN (PROFITABILITY) | GC: >10%"
TEMP_MATRIX(5, 1) = "ASSET TURNOVER(EFFICIENCY) | GC: chk chgs"
TEMP_MATRIX(6, 1) = "FINANCIAL LEVERAGE(GEARING) | GC: chk chgs"
TEMP_MATRIX(7, 1) = "RETURN ON EQUITY  | GC: >15%"
TEMP_MATRIX(8, 1) = "RETAINED TO COMMON EQTY | GC: >= grth rt"
TEMP_MATRIX(9, 1) = "GROWTH OF DEBT | GC: < roe"
TEMP_MATRIX(10, 1) = "RETURN ON INVESTED CAPITAL  | GC: >15%"
TEMP_MATRIX(11, 1) = "RETURN ON TOTAL ASSETS | GC: cmp ind"
TEMP_MATRIX(12, 1) = "CASH FROM OPERATIONS/NET INCOME % | GC: > 0%"
TEMP_MATRIX(13, 1) = "EPS | GC: > prior yr"
TEMP_MATRIX(14, 1) = "CASH FROM OPERATIONS/PER SHARE | GC: >eps "
TEMP_MATRIX(15, 1) = "GROWTH IN CASH FROM OPERATIONS/SHARE % | GC: >eps gr"
TEMP_MATRIX(16, 1) = "QUALITY OF EARNINGS | GC: < 3%"
TEMP_MATRIX(17, 1) = "IMPACT ON NI OF EXPENSING STK OPTIONS | GC: >-5%"
TEMP_MATRIX(18, 1) = "RETAINED EARNINGS/SHAREHOLDER EQUITY | GC: sgr "
TEMP_MATRIX(19, 1) = "FREE CASH FLOW ($M) | GC: > prior yr"
TEMP_MATRIX(20, 1) = "FREE CASH FLOW MARGIN % | GC: >10%"
TEMP_MATRIX(21, 1) = "OPERATING CASH FLOW MARGIN | GC: > 15%"
TEMP_MATRIX(22, 1) = "OPERATING CASH FLOW COVERAGE  | GC: > 0.9"
TEMP_MATRIX(23, 1) = "PRETAX PROFIT"
TEMP_MATRIX(24, 1) = "CAPEX / SALES"
TEMP_MATRIX(25, 1) = "TOTAL DEBT TO EQUITY RATIO | GC: <. 33"
TEMP_MATRIX(26, 1) = "TOTAL DEBT/ NCO | GC: <3.3"
TEMP_MATRIX(27, 1) = "LTD / 2X LAST YEARS EARNINGS | GC: <1"
TEMP_MATRIX(28, 1) = "SHARE BUYBACK/(DILUTION)  | GC: > 0.0%"
TEMP_MATRIX(29, 1) = "STK OPTION SHARES/TOTAL SHARES % | GC: <5%"
TEMP_MATRIX(30, 1) = "NET C&E+STI+MS-TOTAL DEBT | GC: >0"
TEMP_MATRIX(31, 1) = "NET C&E+STI+MS-TOTAL DEBT PER SH | GC: >0"
TEMP_MATRIX(32, 1) = "C&E+STI+MS VS DEBT RATIO | GC: > 1.5"
TEMP_MATRIX(33, 1) = "C&E+STI+MS RATIO CY/PY % | GC: > 0%"
TEMP_MATRIX(34, 1) = "INTEREST COVERAGE RATIO  | GC: > 5"
TEMP_MATRIX(35, 1) = "CURRENT RATIO | GC: > 2.0"
TEMP_MATRIX(36, 1) = "QUICK ASSETS RATIO | GC: > 1.0"
TEMP_MATRIX(37, 1) = "FOOLISH FLOW RATIO | GC: <1.25"
TEMP_MATRIX(38, 1) = "% CHANGE IN ACCOUNTS RECEIVABLE(AR) | GC: < 0%"
TEMP_MATRIX(39, 1) = "CHANGE IN SALES % | GC: >12%"
TEMP_MATRIX(40, 1) = "% CHANGE IN AR VS SALES | GC: < 0%"
TEMP_MATRIX(41, 1) = "DAYS SALES OUTSTANDING CY | GC: < py"
TEMP_MATRIX(42, 1) = "DAYS SALES OUTSTANDING PY"
TEMP_MATRIX(43, 1) = "INVENTORY TURNOVER RATE (CY) | GC: > py"
TEMP_MATRIX(44, 1) = "INVENTORY TURNOVER RATE (PY)"
TEMP_MATRIX(45, 1) = "% CHANGE IN INVENTORY  "
TEMP_MATRIX(46, 1) = "% CHANGE IN INVENTORY VS SALES  | GC: < 0%"
TEMP_MATRIX(47, 1) = "PLANT TURNOVER RATIO | GC: incr yty"
TEMP_MATRIX(48, 1) = "CEO PAY AS % OF NET INCOME"
TEMP_MATRIX(49, 1) = "CASH CONVERSION CYCLE(CCC) | GC: ind cmp"
TEMP_MATRIX(50, 1) = "TOTAL DEBT   | GC:  "
TEMP_MATRIX(51, 1) = "AVERAGE EQUITY"
TEMP_MATRIX(52, 1) = "AVERAGE ASSETS"
TEMP_MATRIX(53, 1) = "NET PROFIT / NET INCOME / NET EARNINGS "
TEMP_MATRIX(54, 1) = "EPS GROWTH RATE"

'--------------------------------------------------------------------------------
For k = 1 To NCOLUMNS
'--------------------------------------------------------------------------------
'A color code of red, yellow, or green colors is used to signify high risk,
'caution, or good performance, respectively, for each calculated parameter.
'The value of each calculated number is also given so the user can make his
'or her own assessment of quality or risk. The color code also provides a
'quick overview of stocks in the entire portfolio that may be in trouble as
'measured by a multitude of parameters indicating "red", or conversely in
'"green" for those companies that have excellent quality.
    If DATA_MATRIX(1, k) = 0 Then
        TEMP_MATRIX(1, k + 1) = ""
    Else
        TEMP_MATRIX(1, k + 1) = DATA_MATRIX(1, k)
    End If

'--------------------------------------------------------------------------------
    If DATA_MATRIX(2, k) = 0 Then
        TEMP_MATRIX(2, k + 1) = ""
    Else
        TEMP_MATRIX(2, k + 1) = DATA_MATRIX(2, k)
    End If
    If DATA_MATRIX(35, k) = 0 Then
        TEMP_MATRIX(50, k + 1) = DATA_MATRIX(26, k) + DATA_MATRIX(23, k)
    Else
        TEMP_MATRIX(50, k + 1) = DATA_MATRIX(35, k)
    End If
'--------------------------------------------------------------------------------
    If (DATA_MATRIX(29, k) > 0 Or DATA_MATRIX(29, k) < 0) Then
        TEMP_MATRIX(51, k + 1) = (DATA_MATRIX(28, k) + DATA_MATRIX(29, k)) / 2
    Else
        TEMP_MATRIX(51, k + 1) = 0
    End If
'--------------------------------------------------------------------------------
    If (DATA_MATRIX(21, k) > 0 And DATA_MATRIX(22, k) > 0) Then
        TEMP_MATRIX(52, k + 1) = (DATA_MATRIX(21, k) + DATA_MATRIX(22, k)) / 2
    Else
        TEMP_MATRIX(52, k + 1) = 0
    End If
'--------------------------------------------------------------------------------
    If (DATA_MATRIX(10, k) > 0 Or DATA_MATRIX(10, k) < 0) Then
        TEMP_MATRIX(53, k + 1) = DATA_MATRIX(10, k)
    Else
        TEMP_MATRIX(53, k + 1) = 0
    End If
'--------------------------------------------------------------------------------
'% Pretax profit margin=(pretax profit / sales ) x 100. Shows how profitable the
'company is, taking into account all income and all costs before paying income
'taxes. Higher is better. Compare companies in the same industry, since margins
'differ significantly between industries.
    If DATA_MATRIX(3, k) > 0 Then
        TEMP_MATRIX(3, k + 1) = DATA_MATRIX(8, k) / DATA_MATRIX(3, k)
    Else
        TEMP_MATRIX(3, k + 1) = ""
    End If

'--------------------------------------------------------------------------------
'% Net profit margin = (net profit /sales ) x 100. % Profitability compared to
'sales after all taxes have been paid. A component of Return On Equity (ROE).
    If DATA_MATRIX(3, k) > 0 Then
        TEMP_MATRIX(4, k + 1) = DATA_MATRIX(10, k) / DATA_MATRIX(3, k)
    Else
        TEMP_MATRIX(4, k + 1) = ""
    End If

'--------------------------------------------------------------------------------
'Sales / Avg Annual Assets
    If TEMP_MATRIX(52, k + 1) = 0 Then
        TEMP_MATRIX(5, k + 1) = ""
    Else
        TEMP_MATRIX(5, k + 1) = DATA_MATRIX(3, k) / TEMP_MATRIX(52, k + 1)
    End If

'--------------------------------------------------------------------------------
'Avg Annual Assets / Avg Shareholder Equity
    If TEMP_MATRIX(51, k + 1) = 0 Then
        TEMP_MATRIX(6, k + 1) = ""
    Else
        TEMP_MATRIX(6, k + 1) = TEMP_MATRIX(52, k + 1) / TEMP_MATRIX(51, k + 1)
    End If

'--------------------------------------------------------------------------------
'
'% Return on Equity (ROE) = (net profit / average equity) x 100 The rate of
'profit the company earns on the stockholder's equity entrusted to management to
'use. In this case, the equity is calculated by averaging the equity at the
'beginning and end of the fiscal year. ROE is also useful to look at to determine
'how much internally generated return on capital is generated to finance future
'growth. Percent Return On Equity is useful in measuring the "efficiency" of management
'compared to competitors. Since the ROE is a key factor in the growth of the company's
'earnings, breaking it into its 3 components allows us to analyze the sources of
'earnings growth and their trends. ROE = (net profit / sales) X (sales / assets) X
'(assets / equity). We do not calculate ROE this way, but use this fact to examine
'the year to year trend of each component on the ROE trend. Component No. 1 (net
'profit / sales ) is the Net Profit Margin described earlier. Component No. 2
'( sales / assets ) is called the Asset Turnover (Efficiency
') Is how efficient and intensively is management utilizing the assets of the
'company. Check when significant changes occur from year to year.. Component
'No. 3 (assets / equity ) is called Financial Leverage (Gearing) or Balance Sheet
'Leverage. From our discussion of Assets and Equity above, we can write:
'(assets / equity) = (assets / (assets - liabilities) Note that Increased debt
'increases both the assets and the liabilities, so the denominator is not significantly
'changed, but the numerator goes up. Therefore increased Debt increases the Financial
'Leverage directly assuming other changes in the balance sheet are relatively less
'significant. However, be aware that the higher the leverage with increased debt,
'the higher the Return on Equity might be, but also the higher the risk. Some leverage
'can increase the returns to the shareholders, but increased debt leverage increases the
'risk of failure and/or bankruptcy. In comparing the year to year changes in Financial
'Leverage and ROE, one should also note the increased percentage of debt. An increase in
'assets due to more debt will raise Financial Leverage, but will tend to lower Asset
'Turnover unless an offsetting increase in sales is made. Note: The three components of
'Return On Equity are usually printed out on the Output worksheet on the three rows
'immediately above the Return On Equity. But Components No.2 and 3 of these 3 rows have
'been hidden in this spreadsheet in order to keep it simple. If you want, you can unhide
'these values.
    
    If TEMP_MATRIX(51, k + 1) = 0 Then
        TEMP_MATRIX(7, k + 1) = ""
    Else
        TEMP_MATRIX(7, k + 1) = DATA_MATRIX(10, k) / TEMP_MATRIX(51, k + 1)
    End If

'--------------------------------------------------------------------------------
'% ROE x (earnings - dividends)/ earnings
    If (DATA_MATRIX(10, k) > 0 Or DATA_MATRIX(10, k) < 0) Then
        TEMP_MATRIX(8, k + 1) = TEMP_MATRIX(7, k + 1) * (DATA_MATRIX(10, k) - DATA_MATRIX(34, k)) / DATA_MATRIX(10, k)
    Else
        TEMP_MATRIX(8, k + 1) = ""
    End If

'--------------------------------------------------------------------------------
'% Growth of annual debt
    If (TEMP_MATRIX(1, k + 1) = TEMP_MATRIX(1, k) And TEMP_MATRIX(50, k + 1) > 0) Then
        TEMP_MATRIX(9, k + 1) = (TEMP_MATRIX(50, k + 1) - TEMP_MATRIX(50, k)) / TEMP_MATRIX(50, k)
    Else
        TEMP_MATRIX(9, k + 1) = "NA"
    End If

'--------------------------------------------------------------------------------
'ROIC = (Net Profit) / (Avg. Equity + LTD) % Return on invested capital =
'(net profit) x 100. / (avg. equity + long term debt) Return on invested capital
'is a measure of the efficiency by which management is utilizing both the equity
'and the long term debt under its care. It is a far better measure of management
'than return on equity when a company has large long term debt. If long term debt
'were zero, the return in invested capital and return on equity would be the same.
    
    If DATA_MATRIX(10, k) = 0 Then
        TEMP_MATRIX(10, k + 1) = ""
    Else
        TEMP_MATRIX(10, k + 1) = DATA_MATRIX(10, k) / (TEMP_MATRIX(51, k + 1) + DATA_MATRIX(26, k))
    End If

'--------------------------------------------------------------------------------
'Return on total assets = (Net Profit / total assets) %Return on total assets=
'(net profit / total assets) x 100. Having a high return on total assets as well
'as a high return on equity is important, since a poor company can show a high
'return on equity in a given year simply by showing a profit with a tiny amount
'of equity. Both return on assets and return on equity should be examined.
    
    If DATA_MATRIX(21, k) > 0 Then
        TEMP_MATRIX(11, k + 1) = DATA_MATRIX(10, k) / DATA_MATRIX(21, k)
    Else
        TEMP_MATRIX(11, k + 1) = ""
    End If

'--------------------------------------------------------------------------------
'(Net Cash from Operating Activities(NCO)/Net Income)-1 % Cash From Operations
'To Net Income =((Cash From Operations / Net Income) -1) X 100.) This is a measure
'of the percentage by which the cash from operations exceeds the net income. On the
'Statement Of Cash Flows, it is desirable that the Net Cash provided by Operating
'Activities be close to or exceed the Net Income to demonstrate a higher quality
'of earnings and that neither item be a user of cash (negative) rather than a provider
'of cash (positive). This spreadsheet flags and quantifies such situations so they
'do not go unoticed.
    
    If TEMP_MATRIX(53, k + 1) = 0 Then
        TEMP_MATRIX(12, k + 1) = ""
    Else
        TEMP_MATRIX(12, k + 1) = (DATA_MATRIX(30, k) / TEMP_MATRIX(53, k + 1)) - 1
    End If

'--------------------------------------------------------------------------------
'Net Income / Weighted Average Shares
    
    If TEMP_MATRIX(53, k + 1) = 0 Then
        TEMP_MATRIX(13, k + 1) = "NA"
    Else
        TEMP_MATRIX(13, k + 1) = TEMP_MATRIX(53, k + 1) / DATA_MATRIX(11, k)
    End If

'--------------------------------------------------------------------------------
'Net Cash from Operations / Avg Weighted Shares
    
    
    If DATA_MATRIX(30, k) = 0 Then
        TEMP_MATRIX(14, k + 1) = ""
    Else
        TEMP_MATRIX(14, k + 1) = DATA_MATRIX(30, k) / DATA_MATRIX(11, k)
    End If

'--------------------------------------------------------------------------------
'(CFO/share CY - CFO/share PY) /CFO/Share PY
    
    If TEMP_MATRIX(1, k + 1) = TEMP_MATRIX(1, k) Then
        TEMP_MATRIX(15, k + 1) = (TEMP_MATRIX(14, k + 1) - TEMP_MATRIX(14, k)) / TEMP_MATRIX(14, k)
    Else
        TEMP_MATRIX(15, k + 1) = "NA"
    End If

'--------------------------------------------------------------------------------
'(NI-NCO)/(Total Assets CY+Total Assets PY)/2
    
    
    If DATA_MATRIX(22, k) = 0 Then
        TEMP_MATRIX(16, k + 1) = "NA"
    Else
        TEMP_MATRIX(16, k + 1) = (DATA_MATRIX(10, k) - DATA_MATRIX(30, k)) / ((DATA_MATRIX(21, k) + DATA_MATRIX(22, k)) / 2)
    End If

'--------------------------------------------------------------------------------
'(NI if Stk Opt Expensed-NI)/NI % Impact on Net Income of Expensing Stock Options.
'Shows the percentage by which EPS were overstated. Thus the reported EPS would be
'lower by this percentage if stock options had been considered and thus is another
'measure of the quality of earnings. When projecting EPS for the next 5 years, it
'is good to know what the latest year EPS would have been if options had been
'expensed. Some companies are already reporting EPS reduced by the expensing of
'stock options. It will now be required.
    
    If DATA_MATRIX(38, k) <= 0 Then
        TEMP_MATRIX(17, k + 1) = "NA"
    Else
        TEMP_MATRIX(17, k + 1) = (DATA_MATRIX(38, k) - DATA_MATRIX(10, k)) / DATA_MATRIX(10, k)
    End If

'--------------------------------------------------------------------------------
'Retained Earnings/Shareholder Equity % Retained to Common Equity (term used in
'Valueline) Also means % Reinvestment Rate = % Internal Growth Rate = Implied
'Growth Rate = ROE x (earnings - dividends) / earnings This is very important
'because this is the rate of return of money left over from the return on equity
'after paying dividends (if any). If no dividends are paid, the ROE and the %
'Retained to Common Equity are the same. The (earnings - dividends) / (earnings)
'term is called the "retention rate". This is also equal to 1.0 minus (dividends /
'earnings) or 1 minus the "payout ratio". Note that this "implied growth rate" is
'theoretical, but is useful at estimating whether the company is generating enough
'funds to pay for expansion to maintain the estimated growth rate. It is also obvious
'that a company that wants to grow rapidly would prefer not to pay dividends. A
'company can grow by either available reinvestment funds or by borrowing money. A useful
'criteria would be to compare the Implied Growth Rate to the SSG projected growth rate.
    
    If (DATA_MATRIX(27, k) > 0 And DATA_MATRIX(28, k) > 0) Then
        TEMP_MATRIX(18, k + 1) = DATA_MATRIX(27, k) / DATA_MATRIX(28, k)
    Else
        TEMP_MATRIX(18, k + 1) = "NA"
    End If

'--------------------------------------------------------------------------------
'Net Cash from Operations - CAPEX

    If DATA_MATRIX(3, k) = "" Then
        TEMP_MATRIX(19, k + 1) = ""
    Else
        TEMP_MATRIX(19, k + 1) = (DATA_MATRIX(30, k) - DATA_MATRIX(32, k))
    End If

'--------------------------------------------------------------------------------
'(NCO-CAPEX-Dvds)/Sales % Margin of Free Cash Flow To Sales = (Cash Flow
'Provided By Operations - Capital Expenditures - Dividends Paid) X 100. /
'Sales The Net Cash from Operating Activities minus capital expenditures and
'dividends paid should be positive. It is expressed in this case as a percentage
'of Sales. Capital expenditures are also referred to as CAPEX.
        
    If DATA_MATRIX(3, k) = "" Then
        TEMP_MATRIX(20, k + 1) = ""
    Else
        TEMP_MATRIX(20, k + 1) = (DATA_MATRIX(30, k) - DATA_MATRIX(32, k) - DATA_MATRIX(34, k)) / DATA_MATRIX(3, k)
    End If

'--------------------------------------------------------------------------------
'%OCF Margin=Net Cash from Operations/Revenue %Operating Cash Flow Margin= (net
'cash from Operations / sales ) Measures the effectiveness of generating cash for
'every dollar of sales. The net cash from Operations is given on the Cash Flow
'Statement. The Sales (or Revenue) is located on the Statement of Income.
    
    If DATA_MATRIX(3, k) > 0 Then
        TEMP_MATRIX(21, k + 1) = DATA_MATRIX(30, k) / DATA_MATRIX(3, k)
    Else
        TEMP_MATRIX(21, k + 1) = ""
    End If

'--------------------------------------------------------------------------------
'
'NCO / (Absolute Value (NCI+NCF) Operating Cash Flow Coverage Ratio = Net Cash
'"provided by"/ ("used in") Operating Activities (Net Cash "provided by"/ ("used in")
'Investing Activities + Net Cash "provided by"/ ("used in") Financing Activities) The
'main operating business is a source of cash and is called Operating Activities (in
'the numerator above) and is divided by the combined total (in the denominator) of
'Investing Activities (which is usually mostly capitol expenditures) plus Financing
'Activities (which is usually mostly debt financing costs) Each of these three items
'listed in the Statement of Cash Flow can have a plus sign (if they are a provider of
'cash) or a minus sign (if they are a user of cash). The Operating Activities (numerator)
'is usually (and hopefully) a provider of cash and has a plus sign, whereas the two items
'in the denominator, when algebraically totaled up are usually net users of cash. If the
'numerator is negative, the program will assign a negative sign for the entire Operating Cash
'Flow Coverage Ratio. The larger this ratio is (say one or above) , the more successful
'the main Operating business is and the less dependence on spending for capitol expenditures
'and debt repayment. Examining the three parts of the Cash Flow Statement is of major
'importance in determining the quality of the total cash flow of the company and goes far
'beyond the Cash Flow Coverage Ratio discussed here. It is important that this be studied
'separately as well. (i.e. What do each of the three components tabulated on the cash
'flow statement consist of?)
   
    If (DATA_MATRIX(31, k) = 0 And DATA_MATRIX(33, k) = 0) Then
        TEMP_MATRIX(22, k + 1) = ""
    Else
        TEMP_MATRIX(22, k + 1) = DATA_MATRIX(30, k) / Abs((DATA_MATRIX(31, k) + DATA_MATRIX(33, k)))
    End If

'--------------------------------------------------------------------------------
   TEMP_MATRIX(23, k + 1) = DATA_MATRIX(8, k)
'--------------------------------------------------------------------------------
    If (DATA_MATRIX(32, k) <> 0 And DATA_MATRIX(3, k) <> 0 And DATA_MATRIX(32, k) <> "" And DATA_MATRIX(3, k) <> "") Then
        TEMP_MATRIX(24, k + 1) = DATA_MATRIX(32, k) / DATA_MATRIX(3, k)
    Else
        TEMP_MATRIX(24, k + 1) = "NA"
    End If

'--------------------------------------------------------------------------------
'Total debt to equity ratio = total debt / equity Total Debt to Equity Ratio =
'Total Debt / Average Equity Lower debt may permit management to have greater
'flexibility during difficult economic times and to pay less interest costs in
'servicing the debt. Normally long term debt is used. However, total debt is
'used here and therefore a much more pessimistic number is produced. Total debt
'is the sum of short term and long term debt. GE was recently criticized by noted
'Bond Analyst Bill Gross (March 2002) for having an excessive short term debt in
'the form of commercial paper. Debt is not necessarily bad if properly managed.
    
    If DATA_MATRIX(35, k) > 0 Then
        TEMP_MATRIX(25, k + 1) = DATA_MATRIX(35, k) / DATA_MATRIX(28, k)
    Else
        If TEMP_MATRIX(50, k + 1) = 0 Then
            TEMP_MATRIX(25, k + 1) = "No Debt"
        Else
            If DATA_MATRIX(28, k) < 0 Then
                TEMP_MATRIX(25, k + 1) = "Neg Eqty"
            Else
                TEMP_MATRIX(25, k + 1) = TEMP_MATRIX(50, k + 1) / DATA_MATRIX(28, k)
            End If
        End If
    End If
'--------------------------------------------------------------------------------
'Total Debt /Net Cash from Operating Activities(NCO)
    
    If TEMP_MATRIX(50, k + 1) < 0.01 Then
        TEMP_MATRIX(26, k + 1) = "No Debt"
    Else
        TEMP_MATRIX(26, k + 1) = TEMP_MATRIX(50, k + 1) / (DATA_MATRIX(30, k))
    End If
'--------------------------------------------------------------------------------
'Long Term Debt / 2 times last years Net Income
    
    If DATA_MATRIX(26, k) = 0 Then
        TEMP_MATRIX(27, k + 1) = "No LTD"
    Else
        If DATA_MATRIX(10, k) > 0 Then
            TEMP_MATRIX(27, k + 1) = DATA_MATRIX(26, k) / (DATA_MATRIX(10, k) * 2)
        Else
            TEMP_MATRIX(27, k + 1) = "Loss"
        End If
    End If
'--------------------------------------------------------------------------------
'(# shares last year - # shares this year) / # shares last year Share Buyback (Vs Share
'Dilution) = (Shares Last Year - Shares This Year) / Shares Last year If a company actually
'buys back its shares (rather than just authorizes it and talks about it!), the earnings
'per share will increase and existing stockholders will own a greater percentage of the
'company. If the calculated equation comes out negative, the number of shares this year has
'been increased and we have "share dilution" which is the opposite effect. From studying
'several companies, it will be noted that share dilution is far more common as a result of
'an overage of stock options granted or new stock offerings all exceeding the number of
'actual buybacks.
    
    If DATA_MATRIX(12, k) > 0 Then
        TEMP_MATRIX(28, k + 1) = (DATA_MATRIX(12, k) - DATA_MATRIX(11, k)) / DATA_MATRIX(12, k)
    Else
        TEMP_MATRIX(28, k + 1) = ""
    End If
'--------------------------------------------------------------------------------
'Stk Opt Sh Granted/Total Shares Stock Option Shares / Total Number of Shares Shows the
'percentage of shares that management has awarded of the total shares outstanding. Our
'color criteria is 5%. Some of these shares may be "in the money" and thus counted in
'the diluted shares and some may not. Thus one needs to look more deeply to determine
'the impact. Since we buy a stock expecting the price to rise and if it does then these
'awarded shares may be excercised and thus further dilute shareholder earnings and equity,
'it seems to be an area the investor should understand before investing in such a company.
'Perhaps a company can still show suitable growth and potential total return even with
'such a large options overhang but it is an area that should be considered.
    
    If DATA_MATRIX(37, k) = "" Then
        TEMP_MATRIX(29, k + 1) = "NA"
    Else
        TEMP_MATRIX(29, k + 1) = DATA_MATRIX(37, k) / DATA_MATRIX(11, k)
    End If
'--------------------------------------------------------------------------------
'Cash & Equiv.+Sh Tm Inv.+Marketable Securities-Total Debt Net Cash= (Cash &
'Equivalents - Long Term Debt): C&E as stated here really includes short term
'investments and short term marketable securities in addition to Cash and Cash
'Equivalents. This will be a positive number for a company with lots of cash and
'little or no debt. A large positive number is an ideal signal of financial strength,
'but a lot of good companies will also have negative net cash.
    
    If DATA_MATRIX(13, k) > 0 Then
        TEMP_MATRIX(30, k + 1) = DATA_MATRIX(13, k) - TEMP_MATRIX(50, k + 1)
    Else
        TEMP_MATRIX(30, k + 1) = ""
    End If
'--------------------------------------------------------------------------------
'Cash & Equiv.+Sh Tm Inv.+Marketable Securities-Total Debt/Shares
    
    If DATA_MATRIX(13, k) = 0 Then
        TEMP_MATRIX(31, k + 1) = ""
    Else
        TEMP_MATRIX(31, k + 1) = (DATA_MATRIX(13, k) - TEMP_MATRIX(50, k + 1)) / DATA_MATRIX(11, k)
    End If
'--------------------------------------------------------------------------------
'(Cash & Equiv+ Sh Tm Inv+MS) / Total Debt % Ratio of Cash From Operations to Total
'Debt = (Cash From Operations / Total Debt) X 100. This is intended to show how well
'the Cash From Operations covers the total outstanding debt. A measure of less than 25%
'indicates limited financial ability.
        
    If TEMP_MATRIX(50, k + 1) = 0 Then
        TEMP_MATRIX(32, k + 1) = "No Debt"
    Else
        TEMP_MATRIX(32, k + 1) = (DATA_MATRIX(13, k) / TEMP_MATRIX(50, k + 1))
    End If
'--------------------------------------------------------------------------------
'Cash & Equiv +Sh Tm Inv+Marketable Securities CY/Same for PY Cash & Equivalents Ratio
'CY/PY = Cash & Equivalents (Current Year) / Cash & Equivalents (Prior Year) C&E as stated
'here really includes short term investments and short term marketable securities in addition
'to Cash and Cash equivalents. This is a measure if the company is improving on the Cash &
'Equivalents available this year over last year.
    
    If DATA_MATRIX(14, k) = 0 Then
        TEMP_MATRIX(33, k + 1) = ""
    Else
        TEMP_MATRIX(33, k + 1) = (DATA_MATRIX(13, k) / DATA_MATRIX(14, k)) - 1
    End If
'--------------------------------------------------------------------------------
'(Pretax Profit + Interest Expense) / Interest Expense Interest Coverage Ratio =
'(Pretax Profit + Interest Expense) / Interest Expense This is expressed as a ratio
'of the number of times the excess of pretax profit plus the interest expense exceeds
'the interest expense. Higher numbers show an increased capability to handle interest
'costs. It is at a maximum if there is no debt and therefore no interest is paid. In
'this case, the words "No Debt" are printed out.
    
    If DATA_MATRIX(7, k) > 0 Then
        TEMP_MATRIX(34, k + 1) = (TEMP_MATRIX(23, k + 1) + DATA_MATRIX(7, k)) / DATA_MATRIX(7, k)
    Else
        If DATA_MATRIX(7, k) = 0 Then
            TEMP_MATRIX(34, k + 1) = ">99"
        Else
            TEMP_MATRIX(34, k + 1) = ""
        End If
    End If
'--------------------------------------------------------------------------------
'Current ratio= Current Assets / Current Liabilities Current Ratio = Current Assets /
'Current Liabilities This is a measure of short term liquidity where current assets
'(available in one year) could be used to pay off current liabilities (debt due within
'one year). A ratio of 1 to 2 is typical depending on the nature of the company's business.
'Higher is better, but numbers over 3 or 4 indicate excess cash may not be put to work
'efficiently. Less that 1 is called negative working capital and is rare or used as an
'interest free way of raising cash. (Omnicom, OMC, an advertising company, for example
'is negative (ratio less than 1) every year.) For some businesses, negative working
'capitol may not be a bad thing. Sonic tends to be negative.
    
    If (DATA_MATRIX(25, k) > 0 Or DATA_MATRIX(25, k) < 0) Then
        TEMP_MATRIX(35, k + 1) = DATA_MATRIX(19, k) / DATA_MATRIX(25, k)
    Else
        TEMP_MATRIX(35, k + 1) = "NA"
    End If
'--------------------------------------------------------------------------------
'Quick Assets Ratio= (C&E+STI+MS+AR) / Current Liabilities Quick Assets Ratio = (Cash &
'Equivalents+Short TermInvestments+Short Term Marketable Securites+Accounts Receivable) /
'Current Liabilities: Cash and other assets which can or will be converted into cash fairly
'soon, such as accounts receivable and marketable securities; or equivalently, current
'assets minus inventory.
    
    If (DATA_MATRIX(25, k) > 0 Or DATA_MATRIX(25, k) < 0) Then
        TEMP_MATRIX(36, k + 1) = (DATA_MATRIX(13, k) + DATA_MATRIX(15, k)) / DATA_MATRIX(25, k)
    Else
        TEMP_MATRIX(36, k + 1) = "NA"
    End If
'--------------------------------------------------------------------------------
'Foolish Flow Ratio = (CA-C&E)/CL-STD) Foolish Flow Ratio= (Current Assets -
'Cash& Equivalents) / (Current Liabilities - Short Term Debt) is taken from
'The Motley Fool. It is a ratio of "bad assets to good liabilities." What do we
'mean by a "bad asset" and "good liability"? A bad asset is high inventories and
'high accounts receivable. A good liability is a high accounts payable. If you
'look at a typical Balance Sheet, Current Assets consist of "bad assets" (where
'people owe you money such as accounts receivable, or that cost you money such as
'inventories that you have to wait till you can sell) and "good assets" like Cash
'and Equivalents that are immediately available. Current Liabilites consist of "good
'liabilities" such as accounts payable (that you owe other people but have not paid
'yet) and "bad liabilities" such as short term debt that must be paid right away.
'The type of companies that have high inventories to sell will not be as able to keep
'this ratio down as easily as the types of businesses that have no inventory. So this
'ratio has to be compared to peer companies to be useful. An arbitrary criteria of 1.25
'was used only to flag the need to investigate the nature of a company that might have
'a number higher than its peers in the same industry.
    
    If (DATA_MATRIX(25, k) = DATA_MATRIX(23, k) Or DATA_MATRIX(25, k) = 0) Then
        TEMP_MATRIX(37, k + 1) = "NA"
    Else
        TEMP_MATRIX(37, k + 1) = (DATA_MATRIX(19, k) - DATA_MATRIX(13, k)) / (DATA_MATRIX(25, k) - DATA_MATRIX(23, k))
    End If
'--------------------------------------------------------------------------------
'AR CY / AR PY
    If DATA_MATRIX(16, k) > 0 Then
        TEMP_MATRIX(38, k + 1) = (DATA_MATRIX(15, k) / DATA_MATRIX(16, k)) - 1
    Else
        TEMP_MATRIX(38, k + 1) = 0
    End If
'--------------------------------------------------------------------------------
'Sales CY / Sales PY Change in Sales This measures the percent change in sales for the
'current year (CY) compared to the prior year (PY).
    
    If DATA_MATRIX(4, k) > 0 Then
        TEMP_MATRIX(39, k + 1) = (DATA_MATRIX(3, k) / DATA_MATRIX(4, k)) - 1
    Else
        TEMP_MATRIX(39, k + 1) = ""
    End If
'--------------------------------------------------------------------------------
'Chg AR% - Chg Sales% Change in AR/Sales This measures the percent change in the
'ratios of Accounts Receivable (AR) to Sales for current year (CY) to prior year (PY)
    
    If TEMP_MATRIX(38, k + 1) = 0 Then
        TEMP_MATRIX(40, k + 1) = 0
    Else
        TEMP_MATRIX(40, k + 1) = TEMP_MATRIX(38, k + 1) - TEMP_MATRIX(39, k + 1)
    End If
'--------------------------------------------------------------------------------
'Days Sales Outstanding ( CY) The Days Sales Outstanding (DSO) indicates how many days it
'is taking to convert the uncollected sales to cash. A short turnover period and a stable
'or declining trend are positive indicators of receivable quality. (Net Sales/365 days) =
'Sales per Day (Average Accounts Receivable / Sales per Day) = DSO
    
    If DATA_MATRIX(3, k) > 0 Then
        TEMP_MATRIX(41, k + 1) = DATA_MATRIX(15, k) / (DATA_MATRIX(3, k) / COUNT_BASIS)
    Else
        TEMP_MATRIX(41, k + 1) = ""
    End If
'--------------------------------------------------------------------------------
'AR PY / (Sales PY / 365)
    
    If DATA_MATRIX(3, k) > 0 Then
        TEMP_MATRIX(42, k + 1) = DATA_MATRIX(16, k) / (DATA_MATRIX(4, k) / COUNT_BASIS)
    Else
        TEMP_MATRIX(42, k + 1) = ""
    End If
'--------------------------------------------------------------------------------
'Inventory Turnover Rate (CY) The inventory turnover rate for the current year (CY) is the
'number of times annually that the dollar value of the current inventory can be sold in a
'given year. The more inventory "turns" that can be achieved per year , the greater the
'liquidity of the inventories. Turnover rate (CY) = Cost of Goods Sold (CY) / Inventory (CY)
    
    If DATA_MATRIX(17, k) > 0 Then
        TEMP_MATRIX(43, k + 1) = DATA_MATRIX(5, k) / DATA_MATRIX(17, k)
    Else
        TEMP_MATRIX(43, k + 1) = "No Inv"
    End If
'--------------------------------------------------------------------------------
'COG PY / Inv PY
    
    If DATA_MATRIX(17, k) > 0 Then
        TEMP_MATRIX(44, k + 1) = DATA_MATRIX(6, k) / DATA_MATRIX(18, k)
    Else
        TEMP_MATRIX(44, k + 1) = "No Inv"
    End If
'--------------------------------------------------------------------------------
'INV CY / INV PY
    
    If DATA_MATRIX(18, k) > 0 Then
        TEMP_MATRIX(45, k + 1) = (DATA_MATRIX(17, k) / DATA_MATRIX(18, k)) - 1
    Else
        TEMP_MATRIX(45, k + 1) = "No Inv"
    End If
'--------------------------------------------------------------------------------
'Chg Inv% - Chg Sales%
    
    If DATA_MATRIX(18, k) = 0 Then
        TEMP_MATRIX(46, k + 1) = 0
    Else
        TEMP_MATRIX(46, k + 1) = TEMP_MATRIX(45, k + 1) - TEMP_MATRIX(39, k + 1)
    End If
'--------------------------------------------------------------------------------
'Plant Turnover Ratio = Sales / Plant, Property, and Equipment (PP&E) This measures the
'Sales that are returned relative to the value of the plant, property, and equipment that
'exists. Higher values are better in that they indicate a better return on the facilities.
'Additional capital expenditures (CAPEX) to add to the total PP&E should be met with improved
'sales in subsequent years in order to maintain or improve the Plant Turnover Ratio.
    
    If DATA_MATRIX(20, k) > 0 Then
        TEMP_MATRIX(47, k + 1) = DATA_MATRIX(3, k) / DATA_MATRIX(20, k)
    Else
        TEMP_MATRIX(47, k + 1) = "NA"
    End If
'--------------------------------------------------------------------------------
'CEO Total Compensation /Net Income
    
    If DATA_MATRIX(36, k) = 0 Then
        TEMP_MATRIX(48, k + 1) = "NA"
    Else
        If DATA_MATRIX(10, k) < 0 Then
            TEMP_MATRIX(48, k + 1) = ">100%"
        Else
            TEMP_MATRIX(48, k + 1) = DATA_MATRIX(36, k) / DATA_MATRIX(10, k)
        End If
    End If
'--------------------------------------------------------------------------------
'CCC=DIO+DSO-DPO (From Fool) Cash Conversion Cycle (CCC) = Days Inv Outstanding +
'Days Sales Outstanding - Days Payables Outstanding CCC is from the Fool and is
'briefly the time required to turn a dollar spent on goods sold back into cash.
    
    If (DATA_MATRIX(5, k) > 0 And DATA_MATRIX(17, k)) Then
        TEMP_MATRIX(49, k + 1) = (DATA_MATRIX(17, k) / (DATA_MATRIX(5, k) / COUNT_BASIS)) + (DATA_MATRIX(15, k) / (DATA_MATRIX(3, k) / COUNT_BASIS)) - (DATA_MATRIX(24, k) / (DATA_MATRIX(5, k) / COUNT_BASIS))
    Else
        TEMP_MATRIX(49, k + 1) = "NA"
    End If
'--------------------------------------------------------------------------------
    If TEMP_MATRIX(1, k + 1) = TEMP_MATRIX(1, k) Then
        TEMP_MATRIX(54, k + 1) = (TEMP_MATRIX(13, k + 1) - TEMP_MATRIX(13, k)) / TEMP_MATRIX(13, k)
    Else
        TEMP_MATRIX(54, k + 1) = "NA"
    End If
'--------------------------------------------------------------------------------
Next k
'--------------------------------------------------------------------------------

NON_FINANCIAL_INSTITUTIONS_COMPARABLE_FUNC = TEMP_MATRIX

Exit Function
ERROR_LABEL:
NON_FINANCIAL_INSTITUTIONS_COMPARABLE_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : FINANCIAL_INSTITUTIONS_COMPARABLE_FUNC

'DESCRIPTION   : This comparable valuation algorithm uses valuation ratios
'of financial institution(s) and apply ratios and proportions to the companies
'being analyzed. The financial condition, results of operations and asset quality
'of a financial institution are significantly dependent on the macroeconomic, social
'and political conditions prevailing in the countries in which the Firm operates.
'Accordingly, decreases in the growth rate, periods of negative growth, increases in
'inflation, changes in law, regulation, policy, or future judicial rulings and
'interpretations of policies involving exchange controls and other matters such as
'(but not limited to) currency depreciation, inflation, interest rates, taxation,
'banking laws and regulations and other political or economic developments in or
'affecting the countries may affect the overall business environment and may in turn
'impact our financial condition and results of operations.

'LIBRARY       : FUNDAMENTAL
'GROUP         : COMPARABLES
'ID            : 002
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 08/04/2009
'************************************************************************************
'************************************************************************************

Function FINANCIAL_INSTITUTIONS_COMPARABLE_FUNC(ByRef DATA_RNG As Variant)

'In comparing the data between banks, the comparisons should also be made:
'1. Over a couple years time for a given bank 2. Only between banks in a given
'peer group. (i.e. Compare credit card companies to their peers; small regional
'banks to their peers, etc.) Recall that Assets - Liabilities = Equity Keep in
'mind that as the bank grows, all parts of the equation must grow as well.
'Larger banks will tend to have a large amount of investment securities as
'well, although loans will likely be the majority asset in most all banks.

'INCOME STATEMENT There is nothing complicated about understanding banks;
'you just have to be familiar with the language they use and how their
'business is conducted. IT IS QUITE DIFFERENT! Their "Operating Revenue"
'consists of Net Interest Income from loans PLUS Non Interest Income from
'fees, etc, MINUS a Loan Loss Provision (or LLP) for anticipated bad loans
'PLUS a Tax Equivalent Adjustment (or TEA) used to adjust tax exempt income
'up to a pre tax level based on the overall corporate calculated tax rate.
'(In many cases, the Net Interest Income is already adjusted for any tax
'exempt income (It may say "taxable equivalent", possibly in a footnote).
'In this case, for the Tax Equivalent Adjustment, you may input a zero. If
'it does not say, look in places like the Management's Discussion and Analysis
'or tables that may give the Tax Equivalent Adjustment, or calculate the Net
'Interest Income with and without the added TEA. The SEC does not allow the
'Net Interest Income that appears in the Income Statement to include the
'TEA So to restate: Operating Revenue= Net Interest Income before Loan
'Losses LLP + Non Interest Income -LLP + TEA. That was the Income. Now for
'the expenses: You have the Total Interest Expense and Non Interest Expense
'(more properly called "Other Expense") Total Interest Expense consists of
'interest that the banks have to pay out for deposits, other borrowed money,
'and long term debt. Non Interest Expense (or "Other Expense") consists of
'Salaries, benefits, Occupancy, furniture and equipment, Office, Audit and
'regulatory fees, marketing, and other. Notice these are almost fixed fees
'that are the cost of doing business . Non Interest Expense is NOT tied to
'Non Interest Income. Mentally connecting these terms causes a lot of confusion.
'Many banks therefore PROPERLY call Non Interest Expenses "Other Expenses" and
'call Non Interest Income "Other Income." Let’s talk about Loan Loss Provisions
'(LLP) on the Income Statement: On the Income Statement and Balance Sheet, we
'like to compare the current year (CY) with the prior year (PY) for the Loan
'Loss Provision which the banks set aside for anticipated bad loans. A bank
'may decide to set aside more money for LLP this year relative to last year
'because IN MANAGEMENT'S JUDGEMENT they have doubts about the collectability
'of the loan portfolio due to factors such as the nature of the portfolio,
'trends in loss experience, economic conditions, etc. Since LLP affects the
'calculation of Operating Revenue (as we have seen above) , yet is subjective,
'LLP has to be accounted for somewhere along the line from year to year. An
'overly big LLP takes away from reported income at year end, and an overly
'small one adds to reported income, but with a price to pay later as we will
'show. It could only be that management is being conservative, not necessarily
'actually be faced with big losses. Where it is that management has to "pay
'the piper" for a bad guess one way or the other? That is where the LOAN LOSS
'RESERVE (LLR) comes in. It is analogous to a "kitty" that management either
'contributes to at end of the year if they made the LOAN LOSS PROVISION too
'high; or takes funds out of it if they made the LLP too low. Obviously a
'bank is in a more conservative position if they have a BIG loan loss reserve
'(LLR) as a percentage of loans. The loan loss provision is an item on the
'Income Statement. The loan loss reserve is on the balance sheet. If a bank
'used more money to cover bad loans than was set aside in the LOAN LOSS
'PROVISION, that deficit would be taken out of the LOAN LOSS RESERVE account.
'If they used less, the surplus is added to the LLR. The bank doesn't want to
'be over reserved, because putting too much money aside to cover bad loans can
'take away from other businesses. On the other hand, if a bank is under-reserved,
'it can get caught by surprise by bad loans, and be forced to take a charge
'against earnings. Regulatory agencies can require that a bank set aside
'additional amounts in the LLP and LLR. Net Income is defined the same as for
'any other business. It is what is left over after subtracting from the
'Operating Revenue all the various costs of doing business and all the taxes.

'Be sure you input the Net Interest Income before Loan Losses, because the
'computer will subtract out the loan losses, which are called input the "Net
'Interest Income Before Loan Losses because the computer will subtract out
'the loan losses, which are called the Loan Loss Provision (LLP). The provision
'for loan losses will be input for the Current Year CY and for the Prior Year
'PY, right next to it. The Total Non Interest Income and Total Non Interest
'Expense are sometimes called "Other Income" and "Other Expense", respectively.
'After inputting "Net Income", you are done with the Income Statement.

'
'CONSOLIDATED BALANCE SHEET The Total Investment Loans are THE major part
'of a bank's Assets. It is the source of most of their income. Loans are
'Assets. Deposits from the bank customers on the other hand are the major
'part of the Liabilities. We like to see the Loans increase from prior year
'(PY) to current year (CY) by a substantial percentage as well as the see
'the deposits grow which provide the funds for the bank to make loans. Total
'Assets include the loans, which as we said are a major part of the Assets.
'The cash and cash equivalents of the bank plus the bank premises and other
'assets make up the remainder of the Assets. We also like to see Total Assets
'grow from PY to CY. In calculating the Return on Assets, we divide Net Income
'by the Average Assets from the beginning of the year to the end of the year.
'Total Stockholder's Equity is the Total Assets minus the Total Liability.
'We input this for both the CY and PY so that we may calculate the return on
'Average Equity from the beginning to the end of the year.

'The Total Investment Loans for current year CY and prior year PY, right next
'to it, are often called by other names such as Loans, net of unearned Income
'or Loans and Leases, or Total Loans. Next, you will need to input Total Assets
'for the Current Year CY and Prior Year PY. right next to it. Near the bottom
'of the page, input Shareholder's Equity for CY and PY. Also called
'Stockholder's equity

'SELECTED FINANCIAL DATA: (as available) No need to look for Tax Equivalent
'Adjustment if the Net Interest Income is defined as "taxable equivalent".
'In that case, it has already been added in to Net Interest Income, so we
'input zero for TEA. If it is not said whether the Net Interest Income is
'a "taxable equivalent", then we have to find the TEA from other sections
'of the annual report. The next three items are taken from either the
'Valueline or from the Annual Report Selected Financial Data section or from
'the 10K. They are not always easy to find, or may be using different
'terminology. I prefer to get Non Performing Assets as a % of Loans, Net
'Charge Offs as a percent of average loans, and Loan Loss Reserves as a %
'of loans from the Valueline first for most convenience, then from the
'Annual Report. This input to the spreadsheet is merely reprinted on the
'output sheet. The first two should be less than 1% and as small as possible.
'Loan loss reserves can and should be higher at typically 1 to 2% or more to
'be conservative. Non Performing Assets as a percent of Total Loans means Non
'Performing Loans. (Remember Loans Assets are non performing loans (Loans are
'one type of Assets.) Unlike LLP which is POTENTIAL bad loans, Non Performing
'Assets are ACTUAL experience with bad loans. Non Performing Assets are loans
'where the interest is past due for over 90 days or more; loans that are not
'being paid on schedule; or loans that are being paid at a reduced rate. The
'lower this ratio, the better. It shows that the bank is keeping its bad loans
'under control. Keeping this ratio under 1% or lower is normal. Net Charge Off
'as a percent of average loans is calculated as follows; Net Charge Off =
'(Loans written off - Collected Bad Loans) / Average Total Assets. This is
'apparently the case where some loans may eventually be collected and some
'must be written off as hopeless. The net charge off ratio should also be
'less that 1%. For this number to be positive, as it apparently is, obviously
'the loans written off are larger than the Collected Bad Loans. Actually
'collecting on part of the loans keeps this number down. Loan Loss Reserve
'as a % of loans. This is the "kitty" where the bank goes to get funds when
'they had more bad loans than were estimated in the Loan Loss Provision. This
'number is usually well over 1% and ranges from 1 to 2% normally. The higher
'numbers are more conservative. The loan to deposit ratio is either found or
'can be calculated from the annual report and allows you to see where the
'bank gets its funds needed to finance its loaning activities. If this ratio
'is greater than 1, the bank deposits may not be sufficient to support its
'loaning activities. If this ratio is less than 1, the bank may not be fully
'using its deposits as a profitable source of loans. However anywhere from 60
'to 110% is not abnormal, with smaller banks having a smaller loan to deposit
'ratio. Lower is considered to be better. A trend toward larger ratios with
'time indicates higher leveraging without depositor support. Real Estate
'Owned, called REO or OREO, is foreclosed property. The input is in millions
'of dollars. On annual report it is sometimes called Other Real Estate Owned,
'hence OREO.

'The only remaining green shaded line that must be filled in here is the Tax
'Equivalent Adjustment or TEA. This is an adjustment which takes into account
'that some of the bank's assets are in tax exempt securities, and it adjusts
'that interest income on a fully taxable equivalent basis. Many annual reports
'have already added in this adjustment into their statements and report their
'figures as "taxable equivalent". If this is the case, we input a zero on this
'line. If it does not say, you may have to look into the 10K or the footnotes,
'etc to find out. When a bank has many tax exempt securities, this will not
'be negligible. So if it does not say, you will have to search for the TEA.
'The remaining data in this section are optional but highly desirable input
'that are available in Valueline and/or the annual report. The data as input
'are merely output again. In some cases, the data such as loan loss experience
'are color coded (e.g. green, yellow or red) depending if the numbers are
'better or worse than desired.

'DATA_RNG:
'ROW 1: BANK
'ROW 2: REPORT DATE
'ROW 3: TOTAL INTEREST INCOME
'ROW 4: TOTAL INTEREST EXPENSE
'ROW 5: NET INTEREST INCOME BEFORE LOAN LOSSES (LLP)
'ROW 6: LOAN LOSS PROVISION (LLP)  CY
'ROW 7: LOAN LOSS PROVISION (LLP)  PY
'ROW 8: TOTAL NON INTEREST INCOME  OR ("OTHER INCOME")
'ROW 9: TOTAL NON INTEREST EXPENSE   ("OTHER EXPENSE")
'ROW 10: NET INCOME
'ROW 11: TOTAL INVESTMENT LOANS CY (OR LOANS & LEASES)
'ROW 12: TOTAL INVESTMENT LOANS PY (OR LOANS & LEASES)
'ROW 13: TOTAL ASSETS CY
'ROW 14: TOTAL ASSETS PY
'ROW 15: TOTAL STOCKHOLDER'S EQUITY CY
'ROW 16: TOTAL STOCKHOLDER'S EQUITY PY
'ROW 17: TEA  TAX EQUIVALENT ADJUSTMENT ($M)
'ROW 18: NON PERFORMING ASSETS AS A % OF TOTAL ASSETS
'ROW 19: NET CHARGE-OFFS AS % OF AVERAGE LOANS
'ROW 20: LOAN LOSS "RESERVE" AS A % OF LOANS
'ROW 21: LOAN TO DEPOSIT RATIO
'ROW 22: REAL ESTATE OWNED REO (FORCLSD PROPERTY)$M

Dim k As Long

Dim NROWS As Long
Dim NCOLUMNS As Long

Dim TEMP_MATRIX As Variant
Dim DATA_MATRIX As Variant

On Error GoTo ERROR_LABEL

DATA_MATRIX = DATA_RNG
NROWS = UBound(DATA_MATRIX, 1)
NCOLUMNS = UBound(DATA_MATRIX, 2)

'---------------------------------------------------------------------------
ReDim TEMP_MATRIX(1 To 22, 1 To NCOLUMNS + 1)
'---------------------------------------------------------------------------

TEMP_MATRIX(1, 1) = "COMPANY"
TEMP_MATRIX(2, 1) = "REPORT DATE"
TEMP_MATRIX(3, 1) = "OPERATING REV. (PLUS TEA LESS LOAN LOSS PROVISION)"
TEMP_MATRIX(4, 1) = "EFFICIENCY RATIO | GC: lower, < 58%"
TEMP_MATRIX(5, 1) = "ROA RETURN ON AVERAGE ASSETS | GC: >1.1%"
TEMP_MATRIX(6, 1) = "ROE RETURN ON AVERAGE EQUITY | GC: >15%"
TEMP_MATRIX(7, 1) = "NET INTEREST MARGIN NIM | GC: >4.5-5%, >5%"
TEMP_MATRIX(8, 1) = "AVG SHLDR. EQUITY / AVG TOTAL ASSETS | GC: >5.5%,>=7.5%"
TEMP_MATRIX(9, 1) = "LOANS-- GROWTH FROM PRIOR YEAR | GC: positive"
TEMP_MATRIX(10, 1) = "LOAN LOSS PROVISION GRWTH FROM PY | GC: decr yr to yr"
TEMP_MATRIX(11, 1) = "ARE LOANS GROWING FASTER THAN LOAN LOSS PROVISION? | GC: yes"
TEMP_MATRIX(12, 1) = "PROVIS FOR LOAN LOSS AS % OF TOT LNS CY | GC: <1%"
TEMP_MATRIX(13, 1) = "PROVIS FOR LOANN LOSS AS % OF TOT LNS PY | GC: <1%"
TEMP_MATRIX(14, 1) = "HAS LOAN LOSS PROVISION IMPROVED IN CURRENT YEAR | GC: yes"
TEMP_MATRIX(15, 1) = "INT. INC. AFTER LOAN LOSS PROVISION AS % OF OPER. REV. | GC: higher is best"
TEMP_MATRIX(16, 1) = "NON PERF ASSETS AS A % OF TOTAL ASSETS <1% | GC: <1%"
TEMP_MATRIX(17, 1) = "NET CHARGEOFFS AS A % OF AVG LOANS | GC: <1%"
TEMP_MATRIX(18, 1) = "LOAN LOSS RESERVE AS A % OF LOANS"
TEMP_MATRIX(19, 1) = "LOAN TO DEPOSIT RATIO | GC: lower is best"
TEMP_MATRIX(20, 1) = "REAL ESTATE OWNED AS % OF AV ASSETS | GC:  not > 1%"
TEMP_MATRIX(21, 1) = "AVG TOTAL ASSETS"
TEMP_MATRIX(22, 1) = "AVG SHAREHOLDER EQUITY"

'---------------------------------------------------------------------------
For k = 1 To NCOLUMNS
'---------------------------------------------------------------------------
'A color code of red, yellow, or green colors is used to signify high risk, caution, or good
'performance, respectively, for each calculated parameter. The value of each calculated number
'is also given so the user can make his or her own assessment of quality or risk. The color code
'also provides a quick overview of stocks in the entire portfolio that may be in trouble as measured
'by a multitude of parameters indicating "red", or conversely in "green" for those companies that
'have excellent quality.
    If DATA_MATRIX(1, k) <> "" Then
        TEMP_MATRIX(1, k + 1) = DATA_MATRIX(1, k)
    Else
        TEMP_MATRIX(1, k + 1) = ""
    End If
'---------------------------------------------------------------------------
    If DATA_MATRIX(2, k) <> "" Then
        TEMP_MATRIX(2, k + 1) = DATA_MATRIX(2, k)
    Else
        TEMP_MATRIX(2, k + 1) = ""
    End If
'---------------------------------------------------------------------------

'OPERATING REVENUE Operating Revenue consists of Net Interest Income from loans PLUS Non Interest
'Income from fees, etc, MINUS a Loan Loss Provision (or LLP) for anticipated bad loans PLUS a Tax
'Equivalent Adjustment (or TEA) used to adjust tax exempt income up to a pre tax level based on the
'overall corporate calculated tax rate. (In many cases, the Net Interest Income has already had the
'Tax Equivalent Adjustment added to it. (It may say "taxable equivalent", possibly in a footnote).
'In this case, for the Tax Equivalent Adjustment, you may input a zero. If it does not say, look in
'places like the Management's Discussion and Analysis or tables that may give the Tax Equivalent
'Adjustment, or calculate the Net Interest Income with and without the added TEA. So to restate:
'Operating Revenue = Net Interest Income before Loan Losses LLP + Non Interest Income - LLP + TEA.
    TEMP_MATRIX(3, k + 1) = DATA_MATRIX(5, k) + DATA_MATRIX(8, k) - DATA_MATRIX(6, k) + DATA_MATRIX(17, k)
'---------------------------------------------------------------------------
'(Total Assets CY + Total Assets PY) / 2.0
    TEMP_MATRIX(21, k + 1) = (DATA_MATRIX(13, k) + DATA_MATRIX(14, k)) / 2
'---------------------------------------------------------------------------
'(Equity CY + Equity PY) / 2.0
    TEMP_MATRIX(22, k + 1) = (DATA_MATRIX(15, k) + DATA_MATRIX(16, k)) / 2
'---------------------------------------------------------------------------
'EFFICIENCY RATIO Efficiency Ratio has many definitions used by various banks,
'so it must be defined. The definition we are using is: Efficiency Ratio =
'Non Interest Expense / Operating Expense Where: Non Interest Expense (also
'called "Other Expenses" or “Operating Expenses”) is Salaries and benefits,
'Furniture and equipment, Office, Audit and regulatory fees, marketing, etc.
'Note that these expenses have nothing to do with what kind of income is being
'made (i.e. Interest or Non Interest Income) Operating Revenue is as defined in
'our discussion above. The term "efficiency ratio" is a misnomer. Usually
'efficiency is (what you get out) / (what you put in) , whereas in this case it
'appears to be "upside down" or = (what you put in) / (what you get out) =
'(Operating Expense) / (Operating Revenue Therefore LOWER IS BETTER !! ( We did
'not make the term, but we just have to live with it!) Amy Crane read in the Wall
'Street Journal that a global measure of Efficiency Ratio is 58%. Therefore our
'color conditional formatting of the output is in Green for banks having ratios
'at or below 58%, and in Red for banks having ratios above 58%.
    
    If TEMP_MATRIX(3, k + 1) = 0 Then
        TEMP_MATRIX(4, k + 1) = ""
    Else
        TEMP_MATRIX(4, k + 1) = DATA_MATRIX(9, k) / TEMP_MATRIX(3, k + 1)
    End If
'---------------------------------------------------------------------------
'ROA, RETURN ON AVERAGE ASSETS: Return on Average Assets = Net Income / Average
'Total Assets for the current year (CY), where Average Total Assets (CY) = (Total
'Assets CY + Total Assets PY) / 2.0 Loans make up the majority of Assets for a
'bank, but assets also include bank premises and equipment, etc. Assets will be
'growing during the year and from year to year, therefore we use Average Assets.
'An industry average for Return on Average Assets is 1.1%. Therefore output ROA
'of 1.1% or higher is color coded Green, and lower than 1.1% is color coded in
'Red. Don't let the small number fool you Small differences of .01 - .02% can mean
'many thousands of dollars because the Assets are so high.
    If TEMP_MATRIX(21, k + 1) = 0 Then
        TEMP_MATRIX(5, k + 1) = ""
    Else
        TEMP_MATRIX(5, k + 1) = DATA_MATRIX(10, k) / TEMP_MATRIX(21, k + 1)
    End If
'---------------------------------------------------------------------------
'ROE, RETURN ON AVERAGE EQUITY: Return on Average Equity = Net Income / Average
'Stockholder Equity (CY) where Average Stockholder Equity (CY) = (Stockholder
'Equity CY + Stockholder Equity PY) / 2.0 Return on Equity at or above 15% is
'considered good and is color coded green. Under 15% is coded Red.

    If TEMP_MATRIX(22, k + 1) = 0 Then
        TEMP_MATRIX(6, k + 1) = ""
    Else
        TEMP_MATRIX(6, k + 1) = DATA_MATRIX(10, k) / TEMP_MATRIX(22, k + 1)
    End If
'---------------------------------------------------------------------------
'NET INTEREST MARGIN: This definition can vary between many banks. The
'definition we are using is: Net Interest Margin NIM = (Net Interest Income
'before LLP - LLP) / Investment Loans CY. The Net Interest Income before the
'Loan Loss Provision (LLP) is from the Statement of Income, immediately followed
'by the Loan Loss Provision (LLP) which must be subtracted. Investment Loans for
'the current year (CY) are located on the Balance Sheet (under Assets). Investment
'Loans are also called by names such as "Loans, net of unearned Income" etc. The
'Net Interest Income is the difference between interests earned from loans minus
'interest paid to depositors. A Loan Loss Provision is then subtracted to determine
'the "keepable" part. It is then divided by Investment Loans for the current year.
'In other words, NIM is how much money you are making from interest income as a
'percentage of the Loans made. This parameter combines the measurement of two
'things: The difference between interest earned on loans and interest paid to de
'positors as well as how profitable this difference is relative to the Total Investment
'Loans that exist. This is a very closely watched measure of how well a bank is doing.
'If interest rates go up or down, this measure would not be affected as long as the
'difference between interest made on loans and interest paid to deposits move up or
'down together. A bank examiner on the I Club List in 1998 said that a bank with a Net
'Interest Margin of 4.5 to 5% was in good shape. We have color coded between 4.5 to
'5% as being yellow. Greater than or equal to 5% was coded Green. Less than or equal to
'4.5% was coded Red. ( It is not important if you agree on color coding limits, just
'compare between banks.)
    If DATA_MATRIX(11, k) = 0 Then
        TEMP_MATRIX(7, k + 1) = ""
    Else
        TEMP_MATRIX(7, k + 1) = (DATA_MATRIX(5, k) - DATA_MATRIX(6, k)) / DATA_MATRIX(11, k)
    End If
'---------------------------------------------------------------------------
'Average STOCKHOLDER 'S EQUITY / AVG. TOTAL ASSETS (OR CAPITAL RATIO) Peter
'Lynch, in his book, "Beating The Street" Chapter 12 & 13 on banks, considers
'this ratio to be the most fundamental measure of financial strength. Before he
'invests in any bank, he likes to see this ratio to be at least 7.5% He discusses
'it in great detail as applied to banks he considers buying. The Capital Ratio =
'Avg. Stockholder's Equity CY / Avg. Total Assets CY (expressed as a percentage);
'where Average Stockholder Equity (CY)= (Stockholder Equity CY - Stockholder
'Equity PY) / 2.0 ;where Average Total Assets(CY)= (Total Assets CY - Total Assets PY)
'/ 2.0 Federal institutions require banks to maintain a minimum level of equity-capital.
'If a bank increases its level of risk, it should increase its equity as well in order
'to absorb unanticipated losses and protect depositors. The major capital adequacy ratio
'is the above defined Capital Ratio. For observations of typical Capital Ratios, between
'8 to 10 percent is considered good (co or coded green 5.5 to 7.5% is considered good
'(color coded yellow); and greater than or equal to 10% is very good (color coded green).
'Less than or equal to 8%, we 7.5% is considered very good (color coded green). Less than
'5.5% is color coded Red. However Federal Banking agencies permit lower, but we suggest,
'in this case, reviewing the Regulatory capital section of the Annual Report / 10K to
'investigate further what they have to say. If this ratio increases over time, it means
'that the bank is taking less risk, and conversely if it decreases, more risk. Federal
'banking agencies set minimum capital requirements in several categories such as Total
'Risk Adjusted Capital, Tier 1 Capital, and Tier 1 Leverage Ratios. The actual definition
'of each category is not important here, but it should be noted that Federal regulators
'could raise the minimum requirements. Depending on a banks's capital adequacy, the bank
'can be placed into a regulatory category ranging from well capitalized to critically
'"undercapitalize". Classification in the undercapitalized categories can have a material
'effect on the bank's operations. For instance, the minimum regulatory capital requirement
'for a bank to be considered "well capitalized" would be if it maintained a minimum "Total",
'"Tier 1", and "Leverage" of 10%, 6%, and 5% respectively. It is important to glance over
'the Regulatory Capital section of the Annual Report / 10K to look at trends over time. For
'instance in the Commerce Bankcorp 2001 Annual Report, the bank reported that it would be
'floating a 200 million dollar Convertible Trust Preferred Offering in 2002 in order to
'increase its Capitol Ratio by about one percentage point, thereby putting it about 2 1/4
'points above the "well capitalized" minimum for its "Leverage" Ratio

    If DATA_MATRIX(13, k) = 0 Then
        TEMP_MATRIX(8, k + 1) = ""
    Else
        TEMP_MATRIX(8, k + 1) = TEMP_MATRIX(22, k + 1) / TEMP_MATRIX(21, k + 1)
    End If
'---------------------------------------------------------------------------
'LOANS --- GROWTH FROM PRIOR YEAR (Loans CY - Loans PY) / Loans PY Expressed as a
'percentage growth in dollars, the growth of Loans should be at least positive and
'preferably in the neighborhood of 15% or greater in order to fuel the growth in
'Operating Revenue. It is presently color coded Red only if it is less than or equal
'to 0. Green if greater than 0 (i.e.positive.)

    If DATA_MATRIX(12, k) = 0 Then
        TEMP_MATRIX(9, k + 1) = ""
    Else
        TEMP_MATRIX(9, k + 1) = (DATA_MATRIX(11, k) - DATA_MATRIX(12, k)) / DATA_MATRIX(12, k)
    End If
'---------------------------------------------------------------------------
'LOAN LOSS PROVISION GROWTH FROM PRIOR YEAR Growth in Loan Loss Provision ( LLP)
'is calculated for the current year relative to the prior year and is expressed as
'a percentage. (LLP CY - LLP PY) / LLP PY A bank is required to set aside a portion
'of its revenues for bad loans. This allowance for loan losses is maintained at a
'level that , IN MANAGEMENTS JUDGEMENT, is adequate to absorb credit losses inherent
'in the loan portfolio. The amount of the allowance is based on management's evaluation
'of the collectability of the loan portfolio, credit concentrations, trends in historical
'loss experience, specific impaired loans, and current economic conditions. An increase
'in the loan loss parameter may have been necessitated by an increase in the prior year
'Non Performing Assets as a percentage of Total Loans (to be discussed later) It is
'therefore not color coded.

    If DATA_MATRIX(7, k) = 0 Then
        TEMP_MATRIX(10, k + 1) = ""
    Else
        TEMP_MATRIX(10, k + 1) = (DATA_MATRIX(6, k) - DATA_MATRIX(7, k)) / DATA_MATRIX(7, k)
    End If
'---------------------------------------------------------------------------
'ARE LOANS GROWING FASTER THAN THE LOAN LOSS PROVISION? Loan Growth greater than the
'Loan Loss Provision Growth results in the program printing out YES If Loan Growth >
'LLP Growth is YES This is a healthy sign, but not a necessity in the short run and
'is not color coded. The previous year's loss experience may have rendered it prudent
'for management to increase the Loan Loss Provision this year.

    If (TEMP_MATRIX(9, k + 1) <> "" And TEMP_MATRIX(10, k + 1) <> "") Then
        If TEMP_MATRIX(9, k + 1) >= TEMP_MATRIX(10, k + 1) Then
            TEMP_MATRIX(11, k + 1) = "YES"
        Else
            TEMP_MATRIX(11, k + 1) = "NO"
        End If
    Else
        TEMP_MATRIX(11, k + 1) = ""
    End If
'---------------------------------------------------------------------------
'PROVISION FOR LOAN LOSS AS A % OF TOTAL LOANS IN THE CURRENT YEAR (CY):
'LLP CY / Total Investment Loans CY It is desirable to keep the necessity to
'have a Loan Loss Provision for the Current Year CY below 1% of the Total
'Investment Loans. Less than 1% is color coded Green. Equal to or greater
'than 1% is color coded Red.

    If DATA_MATRIX(11, k) = 0 Then
        TEMP_MATRIX(12, k + 1) = ""
    Else
        TEMP_MATRIX(12, k + 1) = (DATA_MATRIX(6, k) / DATA_MATRIX(11, k))
    End If
'---------------------------------------------------------------------------
'PROVISION FOR LOAN LOSS AS A % OF TOTAL LOANS IN THE PRIOR YEAR (PY):
'LLP PY / Total Investment Loans PY . Same as above, except for the Prior Year (PY)

    If DATA_MATRIX(12, k) = 0 Then
        TEMP_MATRIX(13, k + 1) = ""
    Else
        TEMP_MATRIX(13, k + 1) = (DATA_MATRIX(7, k) / DATA_MATRIX(12, k))
    End If
'---------------------------------------------------------------------------
'HAS LOAN LOSS IMPROVED IN CURRENT YEAR? If the percentage of LLP CY / Total
'Loans is less than LLP PY / Total Loans (both above) , then the answer YES
'is printed out. It is a favorable sign to reduce the Loan Losses as a percent
'of Total Loans relative to the prior year. This is not color coded.
    
    If (TEMP_MATRIX(12, k + 1) <> "" And TEMP_MATRIX(13, k + 1) <> "") Then
        If TEMP_MATRIX(12, k + 1) < TEMP_MATRIX(13, k + 1) Then
            TEMP_MATRIX(14, k + 1) = "YES"
        Else
            TEMP_MATRIX(14, k + 1) = "NO"
        End If
    Else
        TEMP_MATRIX(14, k + 1) = ""
    End If
'---------------------------------------------------------------------------
'NET INTEREST INCOME (AFTER LOAN LOSS PROVISION) AS A PERCENTAGE OF OPERATING
'INCOME: (Net Interest Income - LLP) / Operating Revenue The Net Interest Income
'(after LLP) as a percentage of the Operating Revenue is a measure of the
'profitability of the loan operation. The tendency is for interest income to be
'more profitable than non interest income, which is also part of the denominator.
'(i.e. Is it is better to be less diversified?) Higher is better. It is not color
'coded.
    
    If TEMP_MATRIX(3, k + 1) <> 0 Then
        TEMP_MATRIX(15, k + 1) = (DATA_MATRIX(5, k) - DATA_MATRIX(6, k)) / TEMP_MATRIX(3, k + 1)
    Else
        TEMP_MATRIX(15, k + 1) = ""
    End If
'---------------------------------------------------------------------------
'NON PERFORMING ASSETS AS A % OF TOTAL LOANS TOTAL ASSETS Non Performing
'Assets as a percent of Total Loans means the same as Non Performing Loans.
'Assets are the same as non performing loans (Remember Loans are one type of
'Assets.) Unlike LLP which provides for POTENTIAL bad loans, Non Performing
'Assets are ACTUAL experience with bad loans. Non Performing Assets are loans
'where the interest is past due for over 90 days or more; loans that are not
'being paid on schedule; or loans that are being paid at a reduced rate. The
'lower this ratio, the better. It shows that the bank is keeping its bad loans
'under control. Keeping this ratio under 1% or lower is normal. It is color
'coded green if less than 1%; red if equal to or greater than 1%. Whenever
'Non Performing Assets as a % of Total Loans comes out unacceptably high,
'the tendency of Total Assets comes out unacceptably high, the tendency of
'management is to increase the Loan Loss Provision as a % of Total Loans in
'the following year to exceed the Non Performing Assets in order to more than
'cover a possible repeat of such an experience. In a recent event (at the time
'of this writing) where Synovus had experienced a high Non Performing Assets ratio,
'the management chose a Loan Loss Provision for the following year that was three
'times the Non Performing Assets ratio.
    
    If DATA_MATRIX(18, k) = 0 Then
        TEMP_MATRIX(16, k + 1) = ""
    Else
        TEMP_MATRIX(16, k + 1) = DATA_MATRIX(18, k)
    End If
'---------------------------------------------------------------------------
' NET CHARGE OFF AS A % OF AVERAGE LOANS: Net Charge Off as a percent of average
'loans is calculated as follows; Net Charge Off = (Loans written off - Collected
'Bad Loans) / Average Total Assets. This is the case where some non performing
'loans may eventually be collected and some must be written off as hopeless. The
'net charge off ratio should also be less that 1%. For this number to be positive,
'as it is, obviously the loans written off are larger than the Collected Bad Loans.
'Actually collecting on part of the non performing loans keeps the charge off down.
'It is color coded green if less than 1%; red if equal to or greater than 1%.
   If DATA_MATRIX(19, k) = 0 Then
        TEMP_MATRIX(17, k + 1) = ""
    Else
        TEMP_MATRIX(17, k + 1) = DATA_MATRIX(19, k)
    End If
'---------------------------------------------------------------------------
 'LOAN LOSS RESERVE AS A % OF LOANS: Loan Loss Reserve as a % of loans. This is
 'the "kitty" where the bank goes to get funds when they had more bad loans than
 'were estimated in the Loan Loss Provision. This number is usually well over 1%
 'and ranges from 1 to 2% normally. The higher numbers are more conservative. It
 'is not color coded.
    If DATA_MATRIX(20, k) = 0 Then
        TEMP_MATRIX(18, k + 1) = ""
    Else
        TEMP_MATRIX(18, k + 1) = DATA_MATRIX(20, k)
    End If
'---------------------------------------------------------------------------
'LOAN TO DEPOSIT RATIO: The loan to deposit ratio is either found or can
'be calculated from the annual report and allows you to see where the bank
'gets its funds needed to finance its loaning activities. If this ratio is
'greater than 1, the bank deposits may not be sufficient to support its
'loaning activities. If this ratio is less than 1, the bank may not be
'fully using its deposits as a profitable source of loans. However ranges
'from 60 to 110% are not abnormal, with smaller banks having a smaller loan
'to deposit ratio. Lower is considered to be better. A trend toward larger
'ratios with time indicates higher leveraging without depositor support.
'It is not color coded.
    
    If DATA_MATRIX(21, k) = 0 Then
        TEMP_MATRIX(19, k + 1) = ""
    Else
        TEMP_MATRIX(19, k + 1) = DATA_MATRIX(21, k)
    End If
'---------------------------------------------------------------------------
'OTHER REAL ESTATE OWNED (ORE) (FORCLOSED PROPERTY) Y: Real Estate Owned as
'a % of Average Assets=Real Estate Owned / Average Total Assets This is
'property on which the bank has already forclosed. The REO or OREO as it
'is called for "other real estate owned has been written off as a loss on
'the books as a part of non performing assets. It is worry some if it is
'on the rise, and if there is a lot of it, you can assume the bank is having
'trouble getting rid of it. Less than 1% is color coded green. 1% or greater
'is color coded red.
    
    If DATA_MATRIX(22, k) = 0 Then
        TEMP_MATRIX(20, k + 1) = ""
    Else
        TEMP_MATRIX(20, k + 1) = DATA_MATRIX(22, k) / TEMP_MATRIX(21, k + 1)
    End If
'---------------------------------------------------------------------------
Next k
'---------------------------------------------------------------------------

FINANCIAL_INSTITUTIONS_COMPARABLE_FUNC = TEMP_MATRIX

Exit Function
ERROR_LABEL:
FINANCIAL_INSTITUTIONS_COMPARABLE_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : INTERNATIONAL_INSTITUTIONS_COMPARABLE_FUNC
'DESCRIPTION   :
'LIBRARY       : FUNDAMENTAL
'GROUP         : CREDIT
'ID            : 003
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 02/08/2010
'************************************************************************************
'************************************************************************************

Function INTERNATIONAL_INSTITUTIONS_COMPARABLE_FUNC(ByRef TICKERS_RNG As Variant, _
ByRef EXCHANGE_RNG As Variant, _
ByRef INDUSTRY_RNG As Variant, _
ByRef INTERCEPT_RNG As Variant, _
ByRef BETA_RNG As Variant, _
ByRef STANDARD_ERROR_RNG As Variant, _
ByRef RSQUARED_RNG As Variant, _
ByRef MARKET_RNG As Variant, _
ByRef OPERATING_RNG As Variant, _
ByRef EBITDA_RNG As Variant, _
ByRef DEPRECIATION_RNG As Variant, _
ByRef PREVIOUS_LONG_TERM_DEBT_RNG As Variant, ByRef LONG_TERM_DEBT_RNG As Variant, _
ByRef PREVIOUS_SHORT_TERM_DEBT_RNG As Variant, ByRef SHORT_TERM_DEBT_RNG As Variant, _
ByRef CASH_BALANCE_RNG As Variant, ByRef TAXES_RNG As Variant, _
ByRef PREVIOUS_EQUITY_RNG As Variant, ByRef EQUITY_RNG As Variant, _
ByRef NET_INCOME_RNG As Variant, ByRef DIVIDENDS_RNG As Variant, _
ByRef INTEREST_RNG As Variant, ByRef COUNTRY_RNG As Variant, _
ByRef INDUSTRY_AVG_BETA_RNG As Variant, _
ByRef INTEREST_COVERAGE1_RNG As Variant, ByRef INTEREST_COVERAGE2_RNG As Variant, _
Optional ByVal INTEREST_COVERAGE_THRESHOLD As Double = 10000, _
Optional ByVal POWER_COEFFICIENT As Double = 3, _
Optional ByVal CASH_RATE As Double = 0.04, _
Optional ByVal COUNT_BASIS As Double = 52, _
Optional ByVal OUTPUT As Integer = 0)
'Bottom-up Beta
'-----------------------------------------------------------------------------
'DATA_RNG (EXCLUDE HEADINGS):
'-----------------------------------------------------------------------------
'INDUSTRY_RNG: Industry Name
'EXCHANGE_RNG: Exchange Code
'TICKERS_RNG: Name
'MARKET_RNG: Current Market Capitalization
'OPERATING_RNG: Operating Income
'EBITDA_RNG: EBITDA(Earn Bef Int Dep & Amo)
'DEPRECIATION_RNG: Depreciation & Amortization
'TAXES_RNG: Effective Tax Rate
'EQUITY_RNG: Total Shareholders's Equity
'NET_INCOME_RNG: Net Income
'DIVIDENDS_RNG: Current Dividends ($)
'SHORT_TERM_DEBT_RNG: ST Borrowings
'LONG_TERM_DEBT_RNG: LT Borrowings
'CASH_BALANCE_RNG: Current Cash Balance
'INTEREST_RNG: Interest Expense
'-----------------------------------------------------------------------------
'COUNTRY_RNG (EXCLUDE HEADINGS):
'-----------------------------------------------------------------------------
'COLUMN 1: EXCHANGE CODE (per country)
'COLUMN 2: GOVERNMENT's BOND RATE (TEN YEAR BOND RATE)
'COLUMN 3: MARGINAL TAX RATE per Exchange
'COLUMN 4: RISK PREMIUM
'COLUMN 5: COUNTRY DEFAULT SPREAD
'-----------------------------------------------------------------------------
'INTEREST_COVERAGE1_RNG (EXCLUDE HEADINGS): interest coverage ratio
'-----------------------------------------------------------------------------
'COLUMN 1: >
'COLUMN 2: <= to
'COLUMN 3: Rating is
'COLUMN 4: Spread is
'-----------------------------------------------------------------------------
'INTEREST_COVERAGE2_RNG (EXCLUDE HEADINGS): interest coverage ratio for smaller
'and riskier firms
'-----------------------------------------------------------------------------
'COLUMN 1: >
'COLUMN 2: <= to
'COLUMN 3: Rating is
'COLUMN 4: Spread is
'-----------------------------------------------------------------------------
'-----------------------------------------------------------------------------

Dim i As Long
Dim j As Long
Dim NROWS As Long

Dim MARGINAL_TAX_RATE As Double

Dim TICKERS_VECTOR As Variant
Dim EXCHANGE_VECTOR As Variant
Dim INDUSTRY_VECTOR As Variant

Dim INTERCEPT_VECTOR As Variant
Dim BETA_VECTOR As Variant 'Raw Beta
Dim STANDARD_ERROR_VECTOR As Variant
Dim RSQUARED_VECTOR As Variant

Dim MARKET_VECTOR As Variant
Dim OPERATING_VECTOR As Variant
Dim EBITDA_VECTOR As Variant
Dim DEPRECIATION_VECTOR As Variant 'Depreciation & Amortization

Dim TAXES_VECTOR As Variant
Dim EQUITY_VECTOR As Variant

Dim PREVIOUS_EQUITY_VECTOR As Variant
Dim PREVIOUS_SHORT_TERM_DEBT_VECTOR As Variant
Dim PREVIOUS_LONG_TERM_DEBT_VECTOR As Variant

Dim SHORT_TERM_DEBT_VECTOR As Variant
Dim LONG_TERM_DEBT_VECTOR As Variant
Dim INTEREST_VECTOR As Variant
Dim CASH_VECTOR As Variant
Dim NET_INCOME_VECTOR As Variant
Dim DIVIDENDS_VECTOR As Variant

Dim COUNTRY_MATRIX As Variant
Dim INTEREST_COVERAGE1_MATRIX As Variant
Dim INTEREST_COVERAGE2_MATRIX As Variant
Dim INDUSTRY_AVG_BETA_MATRIX As Variant

Dim TEMP1_MATRIX As Variant
Dim TEMP2_MATRIX As Variant
Dim TEMP3_MATRIX As Variant
Dim TEMP4_MATRIX As Variant

On Error Resume Next
'----------------------------------------------------------------------------
TICKERS_VECTOR = TICKERS_RNG
If UBound(TICKERS_VECTOR, 1) = 1 Then
    TICKERS_VECTOR = MATRIX_TRANSPOSE_FUNC(TICKERS_VECTOR)
End If
NROWS = UBound(TICKERS_VECTOR, 1)

EXCHANGE_VECTOR = EXCHANGE_RNG
If UBound(EXCHANGE_VECTOR, 1) = 1 Then
    EXCHANGE_VECTOR = MATRIX_TRANSPOSE_FUNC(EXCHANGE_VECTOR)
End If
If NROWS <> UBound(EXCHANGE_VECTOR, 1) Then: GoTo ERROR_LABEL

INDUSTRY_VECTOR = INDUSTRY_RNG
If UBound(INDUSTRY_VECTOR, 1) = 1 Then
    INDUSTRY_VECTOR = MATRIX_TRANSPOSE_FUNC(INDUSTRY_VECTOR)
End If
If NROWS <> UBound(INDUSTRY_VECTOR, 1) Then: GoTo ERROR_LABEL

'----------------------------------------------------------------------------
INTERCEPT_VECTOR = INTERCEPT_RNG
If UBound(INTERCEPT_VECTOR, 1) = 1 Then
    INTERCEPT_VECTOR = MATRIX_TRANSPOSE_FUNC(INTERCEPT_VECTOR)
End If
If NROWS <> UBound(INTERCEPT_VECTOR, 1) Then: GoTo ERROR_LABEL

BETA_VECTOR = BETA_RNG
If UBound(BETA_VECTOR, 1) = 1 Then
    BETA_VECTOR = MATRIX_TRANSPOSE_FUNC(BETA_VECTOR)
End If
If NROWS <> UBound(BETA_VECTOR, 1) Then: GoTo ERROR_LABEL

STANDARD_ERROR_VECTOR = STANDARD_ERROR_RNG
If UBound(STANDARD_ERROR_VECTOR, 1) = 1 Then
    STANDARD_ERROR_VECTOR = MATRIX_TRANSPOSE_FUNC(STANDARD_ERROR_VECTOR)
End If
If NROWS <> UBound(STANDARD_ERROR_VECTOR, 1) Then: GoTo ERROR_LABEL

RSQUARED_VECTOR = RSQUARED_RNG
If UBound(RSQUARED_VECTOR, 1) = 1 Then
    RSQUARED_VECTOR = MATRIX_TRANSPOSE_FUNC(RSQUARED_VECTOR)
End If
If NROWS <> UBound(RSQUARED_VECTOR, 1) Then: GoTo ERROR_LABEL
'----------------------------------------------------------------------------
MARKET_VECTOR = MARKET_RNG
If UBound(MARKET_VECTOR, 1) = 1 Then
    MARKET_VECTOR = MATRIX_TRANSPOSE_FUNC(MARKET_VECTOR)
End If
If NROWS <> UBound(MARKET_VECTOR, 1) Then: GoTo ERROR_LABEL

PREVIOUS_EQUITY_VECTOR = PREVIOUS_EQUITY_RNG
If UBound(PREVIOUS_EQUITY_VECTOR, 1) = 1 Then
    PREVIOUS_EQUITY_VECTOR = MATRIX_TRANSPOSE_FUNC(PREVIOUS_EQUITY_VECTOR)
End If
If NROWS <> UBound(PREVIOUS_EQUITY_VECTOR, 1) Then: GoTo ERROR_LABEL

EQUITY_VECTOR = EQUITY_RNG
If UBound(EQUITY_VECTOR, 1) = 1 Then
    EQUITY_VECTOR = MATRIX_TRANSPOSE_FUNC(EQUITY_VECTOR)
End If
If NROWS <> UBound(EQUITY_VECTOR, 1) Then: GoTo ERROR_LABEL
'----------------------------------------------------------------------------
OPERATING_VECTOR = OPERATING_RNG
If UBound(OPERATING_VECTOR, 1) = 1 Then
    OPERATING_VECTOR = MATRIX_TRANSPOSE_FUNC(OPERATING_VECTOR)
End If
If NROWS <> UBound(OPERATING_VECTOR, 1) Then: GoTo ERROR_LABEL

EBITDA_VECTOR = EBITDA_RNG
If UBound(EBITDA_VECTOR, 1) = 1 Then
    EBITDA_VECTOR = MATRIX_TRANSPOSE_FUNC(EBITDA_VECTOR)
End If
If NROWS <> UBound(EBITDA_VECTOR, 1) Then: GoTo ERROR_LABEL

DEPRECIATION_VECTOR = DEPRECIATION_RNG
If UBound(DEPRECIATION_VECTOR, 1) = 1 Then
    DEPRECIATION_VECTOR = MATRIX_TRANSPOSE_FUNC(DEPRECIATION_VECTOR)
End If
If NROWS <> UBound(DEPRECIATION_VECTOR, 1) Then: GoTo ERROR_LABEL

'----------------------------------------------------------------------------
TAXES_VECTOR = TAXES_RNG
If UBound(TAXES_VECTOR, 1) = 1 Then
    TAXES_VECTOR = MATRIX_TRANSPOSE_FUNC(TAXES_VECTOR)
End If
If NROWS <> UBound(TAXES_VECTOR, 1) Then: GoTo ERROR_LABEL

'----------------------------------------------------------------------------
PREVIOUS_SHORT_TERM_DEBT_VECTOR = PREVIOUS_SHORT_TERM_DEBT_RNG
If UBound(PREVIOUS_SHORT_TERM_DEBT_VECTOR, 1) = 1 Then
    PREVIOUS_SHORT_TERM_DEBT_VECTOR = MATRIX_TRANSPOSE_FUNC(PREVIOUS_SHORT_TERM_DEBT_VECTOR)
End If
If NROWS <> UBound(PREVIOUS_SHORT_TERM_DEBT_VECTOR, 1) Then: GoTo ERROR_LABEL

SHORT_TERM_DEBT_VECTOR = SHORT_TERM_DEBT_RNG
If UBound(SHORT_TERM_DEBT_VECTOR, 1) = 1 Then
    SHORT_TERM_DEBT_VECTOR = MATRIX_TRANSPOSE_FUNC(SHORT_TERM_DEBT_VECTOR)
End If
If NROWS <> UBound(SHORT_TERM_DEBT_VECTOR, 1) Then: GoTo ERROR_LABEL

PREVIOUS_LONG_TERM_DEBT_VECTOR = PREVIOUS_LONG_TERM_DEBT_RNG
If UBound(PREVIOUS_LONG_TERM_DEBT_VECTOR, 1) = 1 Then
    PREVIOUS_LONG_TERM_DEBT_VECTOR = MATRIX_TRANSPOSE_FUNC(PREVIOUS_LONG_TERM_DEBT_VECTOR)
End If
If NROWS <> UBound(PREVIOUS_LONG_TERM_DEBT_VECTOR, 1) Then: GoTo ERROR_LABEL

LONG_TERM_DEBT_VECTOR = LONG_TERM_DEBT_RNG
If UBound(LONG_TERM_DEBT_VECTOR, 1) = 1 Then
    LONG_TERM_DEBT_VECTOR = MATRIX_TRANSPOSE_FUNC(LONG_TERM_DEBT_VECTOR)
End If
If NROWS <> UBound(LONG_TERM_DEBT_VECTOR, 1) Then: GoTo ERROR_LABEL

'----------------------------------------------------------------------------
CASH_VECTOR = CASH_BALANCE_RNG
If UBound(CASH_VECTOR, 1) = 1 Then
    CASH_VECTOR = MATRIX_TRANSPOSE_FUNC(CASH_VECTOR)
End If
If NROWS <> UBound(CASH_VECTOR, 1) Then: GoTo ERROR_LABEL

NET_INCOME_VECTOR = NET_INCOME_RNG
If UBound(NET_INCOME_VECTOR, 1) = 1 Then
    NET_INCOME_VECTOR = MATRIX_TRANSPOSE_FUNC(NET_INCOME_VECTOR)
End If
If NROWS <> UBound(NET_INCOME_VECTOR, 1) Then: GoTo ERROR_LABEL

DIVIDENDS_VECTOR = DIVIDENDS_RNG
If UBound(DIVIDENDS_VECTOR, 1) = 1 Then
    DIVIDENDS_VECTOR = MATRIX_TRANSPOSE_FUNC(DIVIDENDS_VECTOR)
End If
If NROWS <> UBound(DIVIDENDS_VECTOR, 1) Then: GoTo ERROR_LABEL

'----------------------------------------------------------------------------
INTEREST_VECTOR = INTEREST_RNG
If UBound(INTEREST_VECTOR, 1) = 1 Then
    INTEREST_VECTOR = MATRIX_TRANSPOSE_FUNC(INTEREST_VECTOR)
End If
If NROWS <> UBound(INTEREST_VECTOR, 1) Then: GoTo ERROR_LABEL
'----------------------------------------------------------------------------
COUNTRY_MATRIX = COUNTRY_RNG
'----------------------------------------------------------------------------
INTEREST_COVERAGE1_MATRIX = INTEREST_COVERAGE1_RNG
INTEREST_COVERAGE2_MATRIX = INTEREST_COVERAGE2_RNG
INDUSTRY_AVG_BETA_MATRIX = INDUSTRY_AVG_BETA_RNG
'----------------------------------------------------------------------------
'---------------------------------------------------------------------------------
'FIRST PASS: CORPORATE CREDIT ANALYSIS
'---------------------------------------------------------------------------------

ReDim TEMP1_MATRIX(0 To NROWS, 1 To 24)

TEMP1_MATRIX(0, 1) = "CORPORATE CREDIT ANALYSIS"
TEMP1_MATRIX(0, 2) = "EXCHANGE"
TEMP1_MATRIX(0, 3) = "INDUSTRY"
TEMP1_MATRIX(0, 4) = "BOND RATE" 'TEN YEAR
TEMP1_MATRIX(0, 5) = "EBIT"
TEMP1_MATRIX(0, 6) = "INTEREST EXPENSE"
TEMP1_MATRIX(0, 7) = "INTEREST COVERAGE RATIO"
TEMP1_MATRIX(0, 8) = "SYNTHETIC RATING"
TEMP1_MATRIX(0, 9) = "DEFAULT SPREAD"
TEMP1_MATRIX(0, 10) = "PRE-TAX COST OF DEBT"
TEMP1_MATRIX(0, 11) = "TAX RATE"
TEMP1_MATRIX(0, 12) = "AFTER-TAX COST OF DEBT"
TEMP1_MATRIX(0, 13) = "MARKET EQUITY"
TEMP1_MATRIX(0, 14) = "BOOK EQUITY"
TEMP1_MATRIX(0, 15) = "LT BORROWINGS"
TEMP1_MATRIX(0, 16) = "ST BORROWINGS"
TEMP1_MATRIX(0, 17) = "BOOK DEBT"
TEMP1_MATRIX(0, 18) = "MARKET VALUE OF DEBT"
TEMP1_MATRIX(0, 19) = "D/(D+E): BOOK"
TEMP1_MATRIX(0, 20) = "D/(D+E): MARKET"
TEMP1_MATRIX(0, 21) = "EBITDA/ MARKET VALUE"
TEMP1_MATRIX(0, 22) = "CASH BALANCE"
TEMP1_MATRIX(0, 23) = "FIRM VALUE"
TEMP1_MATRIX(0, 24) = "CASH / VALUE"

For i = 1 To NROWS
    TEMP1_MATRIX(i, 1) = TICKERS_VECTOR(i, 1)
    TEMP1_MATRIX(i, 2) = EXCHANGE_VECTOR(i, 1)
    TEMP1_MATRIX(i, 3) = INDUSTRY_VECTOR(i, 1)
'-----------------------------------------------------------------------------
    For j = LBound(COUNTRY_MATRIX, 1) To UBound(COUNTRY_MATRIX, 1)
        If COUNTRY_MATRIX(j, 1) = EXCHANGE_VECTOR(i, 1) Then
            TEMP1_MATRIX(i, 4) = COUNTRY_MATRIX(j, 2)
            TEMP1_MATRIX(i, 9) = COUNTRY_MATRIX(j, 5)
            MARGINAL_TAX_RATE = COUNTRY_MATRIX(j, 3)
            'TEMP1_MATRIX(i, 11) = COUNTRY_MATRIX(j, 3)
    'Debug.Print EXCHANGE_VECTOR(i, 1), COUNTRY_MATRIX(j, 1), TAXES_VECTOR(i, 1), MARGINAL_TAX_RATE
            Exit For
        Else
            TEMP1_MATRIX(i, 4) = CASH_RATE
            TEMP1_MATRIX(i, 9) = 0
            MARGINAL_TAX_RATE = 0
            'TEMP1_MATRIX(i, 11) = TAXES_VECTOR(i, 1) 'Effective Tax Rate
        End If
    Next j
    If TEMP1_MATRIX(i, 11) = 0 Then: TEMP1_MATRIX(i, 11) = 0.5
'-----------------------------------------------------------------------------
    TEMP1_MATRIX(i, 5) = OPERATING_VECTOR(i, 1)
    TEMP1_MATRIX(i, 6) = INTEREST_VECTOR(i, 1)
    
    If TEMP1_MATRIX(i, 6) <> 0 Then
        TEMP1_MATRIX(i, 7) = TEMP1_MATRIX(i, 5) / TEMP1_MATRIX(i, 6)
    Else
        TEMP1_MATRIX(i, 7) = 100000
    End If
    
    If MARKET_VECTOR(i, 1) > INTEREST_COVERAGE_THRESHOLD Then
        For j = LBound(INTEREST_COVERAGE1_MATRIX, 1) To UBound(INTEREST_COVERAGE1_MATRIX, 1)
            If TEMP1_MATRIX(i, 7) > INTEREST_COVERAGE1_MATRIX(j, 1) And (TEMP1_MATRIX(i, 7) < INTEREST_COVERAGE1_MATRIX(j, 2) Or TEMP1_MATRIX(i, 7) = INTEREST_COVERAGE1_MATRIX(j, 2)) Then
                TEMP1_MATRIX(i, 8) = INTEREST_COVERAGE1_MATRIX(j, 3)
                TEMP1_MATRIX(i, 9) = TEMP1_MATRIX(i, 9) + INTEREST_COVERAGE1_MATRIX(j, 4)
                Exit For
            Else
                TEMP1_MATRIX(i, 8) = 0
            End If
        Next j
    Else
        For j = LBound(INTEREST_COVERAGE2_MATRIX, 1) To UBound(INTEREST_COVERAGE2_MATRIX, 1)
            If TEMP1_MATRIX(i, 7) > INTEREST_COVERAGE2_MATRIX(j, 1) And (TEMP1_MATRIX(i, 7) < INTEREST_COVERAGE2_MATRIX(j, 2) Or TEMP1_MATRIX(i, 7) = INTEREST_COVERAGE2_MATRIX(j, 2)) Then
                TEMP1_MATRIX(i, 8) = INTEREST_COVERAGE2_MATRIX(j, 3)
                TEMP1_MATRIX(i, 9) = TEMP1_MATRIX(i, 9) + INTEREST_COVERAGE2_MATRIX(j, 4)
                Exit For
            Else
                TEMP1_MATRIX(i, 8) = 0
            End If
        Next j
    End If
    
    TEMP1_MATRIX(i, 10) = TEMP1_MATRIX(i, 4) + TEMP1_MATRIX(i, 9)

'-----------------------------------------------------------------------------
    If TAXES_VECTOR(i, 1) < MARGINAL_TAX_RATE Then
        TEMP1_MATRIX(i, 11) = MARGINAL_TAX_RATE
    Else
        If TAXES_VECTOR(i, 1) > 0.5 Then
            TEMP1_MATRIX(i, 11) = 0.5
        Else
            TEMP1_MATRIX(i, 11) = TAXES_VECTOR(i, 1)
        End If
    End If
'-----------------------------------------------------------------------------
    
    TEMP1_MATRIX(i, 13) = MARKET_VECTOR(i, 1)
    TEMP1_MATRIX(i, 14) = EQUITY_VECTOR(i, 1)
    TEMP1_MATRIX(i, 15) = LONG_TERM_DEBT_VECTOR(i, 1)
    TEMP1_MATRIX(i, 16) = SHORT_TERM_DEBT_VECTOR(i, 1)
    TEMP1_MATRIX(i, 17) = TEMP1_MATRIX(i, 15) + TEMP1_MATRIX(i, 16)
    
    TEMP1_MATRIX(i, 12) = TEMP1_MATRIX(i, 10) * (1 - TEMP1_MATRIX(i, 11))
    TEMP1_MATRIX(i, 18) = TEMP1_MATRIX(i, 6) * (1 - (1 + TEMP1_MATRIX(i, 10)) ^ (-1 * POWER_COEFFICIENT)) / TEMP1_MATRIX(i, 10) + TEMP1_MATRIX(i, 17) / (1 + TEMP1_MATRIX(i, 10)) ^ POWER_COEFFICIENT
    TEMP1_MATRIX(i, 20) = TEMP1_MATRIX(i, 18) / (TEMP1_MATRIX(i, 18) + TEMP1_MATRIX(i, 13))
    TEMP1_MATRIX(i, 21) = EBITDA_VECTOR(i, 1) / (TEMP1_MATRIX(i, 18) + TEMP1_MATRIX(i, 13))
    TEMP1_MATRIX(i, 19) = TEMP1_MATRIX(i, 17) / (TEMP1_MATRIX(i, 17) + TEMP1_MATRIX(i, 14))

    TEMP1_MATRIX(i, 22) = CASH_VECTOR(i, 1)
    TEMP1_MATRIX(i, 23) = TEMP1_MATRIX(i, 13) + TEMP1_MATRIX(i, 15) + TEMP1_MATRIX(i, 16)
    TEMP1_MATRIX(i, 24) = TEMP1_MATRIX(i, 22) / TEMP1_MATRIX(i, 23)
Next i

If OUTPUT = 1 Then 'Credit Analysis
    INTERNATIONAL_INSTITUTIONS_COMPARABLE_FUNC = TEMP1_MATRIX
    Exit Function
End If

'---------------------------------------------------------------------------------
'SECOND PASS: RISK & RETURN ANALYSIS
'---------------------------------------------------------------------------------

ReDim TEMP2_MATRIX(0 To NROWS, 1 To 25)

TEMP2_MATRIX(0, 1) = "RISK & RETURN ANALYSIS"
TEMP2_MATRIX(0, 2) = "EXCHANGE"
TEMP2_MATRIX(0, 3) = "INDUSTRY"
TEMP2_MATRIX(0, 4) = "UNLEVERED BETA"
TEMP2_MATRIX(0, 5) = "MARKET DEBT"
TEMP2_MATRIX(0, 6) = "MARKET EQUITY"
TEMP2_MATRIX(0, 7) = "DEBT TO EQUITY"
TEMP2_MATRIX(0, 8) = "TAX RATE"
TEMP2_MATRIX(0, 9) = "LEVERED BETA"

TEMP2_MATRIX(0, 10) = "INTERCEPT"
TEMP2_MATRIX(0, 11) = "RF(1-BETA)"
TEMP2_MATRIX(0, 12) = "JENSEN'S ALPHA"
TEMP2_MATRIX(0, 13) = "ANNUALIZED"
TEMP2_MATRIX(0, 14) = "RAW BETA"
TEMP2_MATRIX(0, 15) = "STANDARD ERROR"
TEMP2_MATRIX(0, 16) = "UPPER BOUND"
TEMP2_MATRIX(0, 17) = "LOWER BOUND"
TEMP2_MATRIX(0, 18) = "R-SQUARED"
TEMP2_MATRIX(0, 19) = "BOND RATE" 'TEN YEAR
TEMP2_MATRIX(0, 20) = "RISK PREMIUM"
TEMP2_MATRIX(0, 21) = "EXPECTED RETURN"

TEMP2_MATRIX(0, 22) = "NET INCOME (EARNINGS)"
TEMP2_MATRIX(0, 23) = "DIVIDENDS"
TEMP2_MATRIX(0, 24) = "PAYOUT RATIO"
TEMP2_MATRIX(0, 25) = "DIVIDEND YIELD"

For i = 1 To NROWS
    TEMP2_MATRIX(i, 1) = TICKERS_VECTOR(i, 1)
    TEMP2_MATRIX(i, 2) = EXCHANGE_VECTOR(i, 1)
    TEMP2_MATRIX(i, 3) = INDUSTRY_VECTOR(i, 1)
    For j = LBound(INDUSTRY_AVG_BETA_MATRIX, 1) To UBound(INDUSTRY_AVG_BETA_MATRIX, 1)
        If TEMP2_MATRIX(i, 3) = INDUSTRY_AVG_BETA_MATRIX(j, 1) Then
            TEMP2_MATRIX(i, 4) = INDUSTRY_AVG_BETA_MATRIX(j, 8)
            Exit For
        Else
            TEMP2_MATRIX(i, 4) = 0
        End If
    Next j
    TEMP2_MATRIX(i, 5) = TEMP1_MATRIX(i, 18)
    TEMP2_MATRIX(i, 6) = TEMP1_MATRIX(i, 13)
    
    TEMP2_MATRIX(i, 7) = TEMP2_MATRIX(i, 5) / TEMP2_MATRIX(i, 6)
    TEMP2_MATRIX(i, 8) = TEMP1_MATRIX(i, 11)
    TEMP2_MATRIX(i, 9) = TEMP2_MATRIX(i, 4) * (1 + (1 - TEMP2_MATRIX(i, 8)) * TEMP2_MATRIX(i, 7))
    
    TEMP2_MATRIX(i, 10) = INTERCEPT_VECTOR(i, 1)
    TEMP2_MATRIX(i, 11) = (CASH_RATE / COUNT_BASIS) * (1 - BETA_VECTOR(i, 1))
    TEMP2_MATRIX(i, 12) = TEMP2_MATRIX(i, 10) - TEMP2_MATRIX(i, 11)
    TEMP2_MATRIX(i, 13) = (1 + TEMP2_MATRIX(i, 12)) ^ COUNT_BASIS - 1
    TEMP2_MATRIX(i, 14) = BETA_VECTOR(i, 1)
    TEMP2_MATRIX(i, 15) = STANDARD_ERROR_VECTOR(i, 1)
    TEMP2_MATRIX(i, 16) = TEMP2_MATRIX(i, 14) + TEMP2_MATRIX(i, 15)
    TEMP2_MATRIX(i, 17) = TEMP2_MATRIX(i, 14) - TEMP2_MATRIX(i, 15)
    TEMP2_MATRIX(i, 18) = RSQUARED_VECTOR(i, 1)
    
    For j = LBound(COUNTRY_MATRIX, 1) To UBound(COUNTRY_MATRIX, 1)
        If COUNTRY_MATRIX(j, 1) = EXCHANGE_VECTOR(i, 1) Then
            TEMP2_MATRIX(i, 19) = COUNTRY_MATRIX(j, 2)
            TEMP2_MATRIX(i, 20) = COUNTRY_MATRIX(j, 4)
            Exit For
        Else
            TEMP2_MATRIX(i, 19) = CASH_RATE
            TEMP2_MATRIX(i, 20) = 0
        End If
    Next j
    TEMP2_MATRIX(i, 21) = TEMP2_MATRIX(i, 19) + TEMP2_MATRIX(i, 14) * TEMP2_MATRIX(i, 20)
    TEMP2_MATRIX(i, 22) = NET_INCOME_VECTOR(i, 1)
    TEMP2_MATRIX(i, 23) = DIVIDENDS_VECTOR(i, 1)
    If TEMP2_MATRIX(i, 22) > 0 Then
        TEMP2_MATRIX(i, 24) = TEMP2_MATRIX(i, 23) / TEMP2_MATRIX(i, 22)
    Else
        TEMP2_MATRIX(i, 24) = 0
    End If
    TEMP2_MATRIX(i, 25) = TEMP2_MATRIX(i, 23) / TEMP1_MATRIX(i, 13)
Next i

If OUTPUT = 2 Then 'Risk & Return Analysis
    INTERNATIONAL_INSTITUTIONS_COMPARABLE_FUNC = TEMP2_MATRIX
    Exit Function
End If

'---------------------------------------------------------------------------------
'THIRD PASS: COST OF CAPITAL ANALYSIS
'---------------------------------------------------------------------------------

ReDim TEMP3_MATRIX(0 To NROWS, 1 To 23)

TEMP3_MATRIX(0, 1) = "COST OF CAPITAL ANALYSIS"
TEMP3_MATRIX(0, 2) = "EXCHANGE"
TEMP3_MATRIX(0, 3) = "INDUSTRY"
TEMP3_MATRIX(0, 4) = "RISK-FREE RATE"
TEMP3_MATRIX(0, 5) = "RISK PREMIUM"
TEMP3_MATRIX(0, 6) = "LEVERED BETA"
TEMP3_MATRIX(0, 7) = "COST OF EQUITY"
TEMP3_MATRIX(0, 8) = "AFTER-TAX COST OF DEBT"
TEMP3_MATRIX(0, 9) = "MARKET DEBT"
TEMP3_MATRIX(0, 10) = "MARKET EQUITY"
TEMP3_MATRIX(0, 11) = "D/(D+E)"
TEMP3_MATRIX(0, 12) = "COST OF CAPITAL"
TEMP3_MATRIX(0, 13) = "OPERATING INCOME"
TEMP3_MATRIX(0, 14) = "TAX RATE"
TEMP3_MATRIX(0, 15) = "EBIT (1-T)"
TEMP3_MATRIX(0, 16) = "PREVIOUS SHORT TERM DEBT" 'LAST YEAR
TEMP3_MATRIX(0, 17) = "PREVIOUS LT TERM DEBT" 'LAST YEAR
TEMP3_MATRIX(0, 18) = "PREVIOUS BOOK EQUITY" 'LAST YEAR
TEMP3_MATRIX(0, 19) = "PREVIOUS TOTAL CAPITAL" 'LAST YEAR
TEMP3_MATRIX(0, 20) = "ROC"
TEMP3_MATRIX(0, 21) = "WACC"
TEMP3_MATRIX(0, 22) = "ROC-WACC"
TEMP3_MATRIX(0, 23) = "EVA"

For i = 1 To NROWS
    TEMP3_MATRIX(i, 1) = TICKERS_VECTOR(i, 1)
    TEMP3_MATRIX(i, 2) = EXCHANGE_VECTOR(i, 1)
    TEMP3_MATRIX(i, 3) = INDUSTRY_VECTOR(i, 1)
    TEMP3_MATRIX(i, 4) = TEMP2_MATRIX(i, 19)
    TEMP3_MATRIX(i, 5) = TEMP2_MATRIX(i, 20)
    TEMP3_MATRIX(i, 6) = TEMP2_MATRIX(i, 9)
    TEMP3_MATRIX(i, 7) = TEMP3_MATRIX(i, 4) + TEMP2_MATRIX(i, 9) * TEMP2_MATRIX(i, 20)
    TEMP3_MATRIX(i, 8) = TEMP1_MATRIX(i, 12)
    TEMP3_MATRIX(i, 9) = TEMP1_MATRIX(i, 18)
    TEMP3_MATRIX(i, 10) = TEMP1_MATRIX(i, 13)
    TEMP3_MATRIX(i, 11) = TEMP1_MATRIX(i, 18) / (TEMP1_MATRIX(i, 18) + TEMP1_MATRIX(i, 13))
    TEMP3_MATRIX(i, 12) = TEMP3_MATRIX(i, 8) * TEMP3_MATRIX(i, 11) + TEMP3_MATRIX(i, 7) * (1 - TEMP3_MATRIX(i, 11))
    TEMP3_MATRIX(i, 13) = OPERATING_VECTOR(i, 1)
    TEMP3_MATRIX(i, 14) = TEMP1_MATRIX(i, 11)
    
    TEMP3_MATRIX(i, 15) = TEMP3_MATRIX(i, 13) * (1 - TEMP3_MATRIX(i, 14))
    TEMP3_MATRIX(i, 16) = PREVIOUS_SHORT_TERM_DEBT_VECTOR(i, 1)
    TEMP3_MATRIX(i, 17) = PREVIOUS_LONG_TERM_DEBT_VECTOR(i, 1)
    TEMP3_MATRIX(i, 18) = PREVIOUS_EQUITY_VECTOR(i, 1)
    
    TEMP3_MATRIX(i, 19) = TEMP3_MATRIX(i, 16) + TEMP3_MATRIX(i, 17) + TEMP3_MATRIX(i, 18)
    TEMP3_MATRIX(i, 20) = TEMP3_MATRIX(i, 15) / TEMP3_MATRIX(i, 19)
    TEMP3_MATRIX(i, 21) = TEMP3_MATRIX(i, 12)
    TEMP3_MATRIX(i, 22) = TEMP3_MATRIX(i, 20) - TEMP3_MATRIX(i, 21)
    TEMP3_MATRIX(i, 23) = TEMP3_MATRIX(i, 22) * TEMP3_MATRIX(i, 19)
Next i

If OUTPUT = 3 Then 'Capital Cost Analysis
    INTERNATIONAL_INSTITUTIONS_COMPARABLE_FUNC = TEMP3_MATRIX
    Exit Function
End If

'---------------------------------------------------------------------------------
'FORTH PASS: SUMMARY
'---------------------------------------------------------------------------------

ReDim TEMP4_MATRIX(0 To NROWS, 1 To 18)
TEMP4_MATRIX(0, 1) = "SUMMARY"
TEMP4_MATRIX(0, 2) = "EXCHANGE"
TEMP4_MATRIX(0, 3) = "RISKFREE RATE"
TEMP4_MATRIX(0, 4) = "EQUITY RISK PREMIUM"
TEMP4_MATRIX(0, 5) = "RAW BETA" 'REGRE BETA
TEMP4_MATRIX(0, 6) = "JENSEN'S ALPHA" 'Weekly
TEMP4_MATRIX(0, 7) = "ANNUALIZED" 'Yearly
TEMP4_MATRIX(0, 8) = "R-SQUARED"
TEMP4_MATRIX(0, 9) = "UNLEVERED BETA"
TEMP4_MATRIX(0, 10) = "MARKET EQUITY"
TEMP4_MATRIX(0, 11) = "MARKET DEBT"
TEMP4_MATRIX(0, 12) = "DEBT TO EQUITY"
TEMP4_MATRIX(0, 13) = "LEVERED BETA"
TEMP4_MATRIX(0, 14) = "COST OF EQUITY"
TEMP4_MATRIX(0, 15) = "SYNTHETIC RATING"
TEMP4_MATRIX(0, 16) = "PRE-TAX COST OF DEBT"
TEMP4_MATRIX(0, 17) = "COST OF CAPITAL"
TEMP4_MATRIX(0, 18) = "ROC"

For i = 1 To NROWS
    TEMP4_MATRIX(i, 1) = TEMP3_MATRIX(i, 1)
    TEMP4_MATRIX(i, 2) = TEMP2_MATRIX(i, 2)
    TEMP4_MATRIX(i, 3) = TEMP3_MATRIX(i, 4)
    TEMP4_MATRIX(i, 4) = TEMP3_MATRIX(i, 5)
    TEMP4_MATRIX(i, 5) = TEMP2_MATRIX(i, 14)
    TEMP4_MATRIX(i, 6) = TEMP2_MATRIX(i, 12)
    TEMP4_MATRIX(i, 7) = TEMP2_MATRIX(i, 13)
    TEMP4_MATRIX(i, 8) = TEMP2_MATRIX(i, 18)
    TEMP4_MATRIX(i, 9) = TEMP2_MATRIX(i, 4)
    TEMP4_MATRIX(i, 10) = TEMP3_MATRIX(i, 10)
    TEMP4_MATRIX(i, 11) = TEMP3_MATRIX(i, 9)
    TEMP4_MATRIX(i, 12) = TEMP2_MATRIX(i, 7)
    TEMP4_MATRIX(i, 13) = TEMP3_MATRIX(i, 6)
    TEMP4_MATRIX(i, 14) = TEMP3_MATRIX(i, 7)
    TEMP4_MATRIX(i, 15) = TEMP1_MATRIX(i, 8)
    TEMP4_MATRIX(i, 16) = TEMP1_MATRIX(i, 10)
    TEMP4_MATRIX(i, 17) = TEMP3_MATRIX(i, 12)
    TEMP4_MATRIX(i, 18) = TEMP3_MATRIX(i, 20)
Next i

If OUTPUT = 4 Then 'Summary
    INTERNATIONAL_INSTITUTIONS_COMPARABLE_FUNC = TEMP4_MATRIX
Else 'If OUTPUT = 0
    INTERNATIONAL_INSTITUTIONS_COMPARABLE_FUNC = Array(TEMP1_MATRIX, TEMP2_MATRIX, TEMP3_MATRIX, TEMP4_MATRIX)
End If

Exit Function
ERROR_LABEL:
INTERNATIONAL_INSTITUTIONS_COMPARABLE_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : RNG_COMPARABLE_TRANSACTIONS_FUNC
'DESCRIPTION   :
'LIBRARY       : FUNDAMENTAL
'GROUP         : COMPARABLES
'ID            : 004
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 08/04/2009
'************************************************************************************
'************************************************************************************

Function RNG_COMPARABLE_TRANSACTIONS_FUNC(ByVal TICKER_STR As String, _
Optional ByRef SCREENER_RNG As Variant, _
Optional ByVal SRC_WBOOK As Excel.Workbook)
    
Dim h As Long
Dim i As Long
Dim j As Long
Dim k As Long
Dim l As Long

Dim ii As Long
Dim jj As Long

Dim POS_VAL As Long
Dim NSIZE As Long
Dim NROWS As Long
Dim NCOLUMNS As Long

Dim INDEX_ARR As Variant
Dim SCREENER_VECTOR As Variant

Dim TEMP_STR As String
Dim DATE_STR As String
Dim TEMP_VAL As Variant
Dim ADDRESS_STR As String
Dim DST_RNG As Excel.Range
Dim TEMP_RNG As Excel.Range
Dim FORMAT_RNG As Excel.Range
Dim DST_WSHEET As Excel.Worksheet
Dim DATA_MATRIX As Variant

On Error GoTo ERROR_LABEL

RNG_COMPARABLE_TRANSACTIONS_FUNC = False

If IsArray(SCREENER_RNG) = True Then
    SCREENER_VECTOR = SCREENER_RNG
    If UBound(SCREENER_VECTOR, 1) = 1 Then
        SCREENER_VECTOR = MATRIX_TRANSPOSE_FUNC(SCREENER_RNG)
    End If
    If UBound(SCREENER_VECTOR, 1) <> 8 Then: GoTo ERROR_LABEL
    DATA_MATRIX = FINVIZ_COMPARABLE_ANALYSIS_FUNC(TICKER_STR, 1)
    If IsArray(DATA_MATRIX) = False Then: GoTo ERROR_LABEL
    DATA_MATRIX = RNG_COMPARABLE_SCREENER_FUNC(DATA_MATRIX, SCREENER_VECTOR)
    If IsArray(DATA_MATRIX) = False Then: GoTo ERROR_LABEL
Else 'No Screener
1982:
    DATA_MATRIX = FINVIZ_COMPARABLE_ANALYSIS_FUNC(TICKER_STR, 1)
    If IsArray(DATA_MATRIX) = False Then: GoTo ERROR_LABEL
End If

NROWS = UBound(DATA_MATRIX, 1)
NCOLUMNS = UBound(DATA_MATRIX, 2)

NSIZE = NROWS - 2 'Headings & Reference ticker

If SRC_WBOOK Is Nothing Then: Set SRC_WBOOK = ActiveWorkbook
Set DST_WSHEET = WSHEET_ADD_FUNC(PARSE_CURRENT_TIME_FUNC("_"), SRC_WBOOK)
ActiveWindow.DisplayGridlines = False

DATE_STR = Format(Now, "mmm dd, yyyy")

Set FORMAT_RNG = DST_WSHEET.Cells
GoSub RNG_FORMAT_CELLS
Set DST_RNG = FORMAT_RNG.Cells(1, 1)
'--------------------------------------------------------------------------------------------
h = 3
POS_VAL = 9 + h * (NSIZE + 2)
Set DST_RNG = DST_RNG.Offset(POS_VAL - 1, 1)

With DST_RNG
    .value = "Financial Information of Comparable Companies"
    Set FORMAT_RNG = DST_RNG
    GoSub FORMAT_MAINHEAD
            
    With .Offset(1, 0)
        .value = "as of " & DATE_STR
        .HorizontalAlignment = xlGeneral
        Set FORMAT_RNG = DST_RNG.Offset(1, 0)
        GoSub FORMAT_SUBHEAD
    End With
    
    Set TEMP_RNG = Range(.Offset(3, 0), .Offset(3, 18))
    With TEMP_RNG
        .value = Array("Company", "Revenue (M)", "EBITDA (M)", "Net Earning (M, ttm)", _
                "Net Earning (M, forward)", "Market Cap (M)", "Book Value (M, mrq)", _
                "Cash (M)", "Debt (M)", "Most Recent Quarter (mrq)", _
                "Fiscal Year Ends (fye)", "P/E (ttm)", "P/E (forward)", "P/S (ttm)", _
                "P/B (mrq)", "EV / EBITDA (ttm)", "EV / REV (ttm)", "Diluted EPS", _
                "Diluted Shares Outstanding (M)")
        .Font.Bold = True
        
        Set FORMAT_RNG = TEMP_RNG
        GoSub RNG_FORMAT_BORDERS_TB
        
        INDEX_ARR = Array(11, 12, 13, 2, 15, 16, 10, 9, 3, 4, 5, 6, 8, 7, 14)
        For k = 1 To 2
            If k = 1 Then l = 3 Else l = 4
            ii = 1
            For j = 0 To 18
                For i = 1 To NSIZE
                    If k = 1 Then: i = 0
                    If j = 0 Then
                        With .Cells(i + l, j + 1)
                            .value = DATA_MATRIX(i + 2, 1)
                            TEMP_STR = .Text
                            .Font.Bold = True
                        End With
                        .Worksheet.Hyperlinks.Add Anchor:=.Cells(i + l, j + 1), _
                        Address:="http://finance.yahoo.com/q/ks?s=" & TEMP_STR, _
                        TextToDisplay:=TEMP_STR
                    ElseIf j = 4 Then
                        .Cells(i + l, j + 1).formula = _
                            "=" & .Cells(i + l, 5 + 1).Address(False, False) & "/" & .Cells(i + l, 12 + 1).Address(False, False)
                    ElseIf j = 6 Then
                        .Cells(i + l, j + 1).formula = _
                            "=" & .Cells(i + l, 5 + 1).Address(False, False) & "/" & .Cells(i + l, 14 + 1).Address(False, False)
                    ElseIf j = 18 Then
                        .Cells(i + l, j + 1).formula = _
                            "=" & .Cells(i + l, 3 + 1).Address(False, False) & "/" & .Cells(i + l, 17 + 1).Address(False, False)
                    Else
                        With .Cells(i + l, j + 1)
                            jj = INDEX_ARR(ii)
                            TEMP_VAL = DATA_MATRIX(i + 2, jj)
                            .value = IIf(TEMP_VAL = "", 0, TEMP_VAL)
                            .Font.Color = -4165632
                        End With
                    End If
                    If k = 1 Then: i = NSIZE + 1
                Next i
                If j <> 0 And j <> 4 And j <> 6 And j <> 18 Then: ii = ii + 1
            Next j
        Next k
        
        Set FORMAT_RNG = TEMP_RNG.Offset(2, 0)
        GoSub RNG_FORMAT_BORDERS_BL
        
        Set FORMAT_RNG = TEMP_RNG.Offset(NSIZE + 3, 0)
        GoSub RNG_FORMAT_BORDERS_BL
    End With
    
    With .Offset(8 + NSIZE)
        .value = "Abbreviation Guide: "
        .Font.Bold = True
    End With
    .Offset(9 + NSIZE).value = "M = Millions"
    .Offset(10 + NSIZE).value = "mrq = Most Recent Quarter"
    .Offset(11 + NSIZE).value = "ttm = Trailing Twelve Months"
    .Offset(12 + NSIZE).value = "lfy = Last Fiscal Year"
    .Offset(13 + NSIZE).value = "fye = Fiscal Year Ending"
    Range(.Offset(8 + NSIZE), .Offset(13 + NSIZE)).HorizontalAlignment = xlGeneral

End With

'---------------------------------------------------------------------------
Set DST_RNG = DST_RNG.Offset(-POS_VAL + 2, 0)
'---------------------------------------------------------------------------
With DST_RNG
'---------------------------------------------------------------------------
    .value = "Comparable Valuation"
    Set FORMAT_RNG = DST_RNG
    GoSub FORMAT_MAINHEAD
    
    .Offset(1, 0).value = "as of " & DATE_STR
    Set FORMAT_RNG = DST_RNG.Offset(1, 0)
    GoSub FORMAT_SUBHEAD
    
    .Offset(0, 4).value = "Input:"
    Set FORMAT_RNG = DST_RNG.Offset(0, 4)
    GoSub FORMAT_MAINHEAD
    
    With .Offset(0, 5)
        .value = "Minority Interest (M)"
        .Font.Bold = True
    End With
    
    With .Offset(1, 5)
        .value = "Preferred Shares (M)"
        .Font.Bold = True
    End With
    
    With .Offset(0, 6)
        .value = 0
        .Font.Color = -4165632
    End With
    Set FORMAT_RNG = DST_RNG.Offset(0, 6)
    GoSub RNG_FORMAT_BORDERS_BOX
    
    With .Offset(1, 6)
        .value = 0
        .Font.Color = -4165632
    End With
    Set FORMAT_RNG = DST_RNG.Offset(1, 6)
    GoSub RNG_FORMAT_BORDERS_BOX

    Set TEMP_RNG = Range(.Offset(3, 0), .Offset(3, 6))
    With TEMP_RNG
        .value = Array("Company", "P/E (ttm)", "P/E (forward)", "P/S (ttm)", "P/B (mrq)", _
                "EV / EBITDA (ttm)", "EV / REV (ttm)")
        .Font.Bold = True
        
        Set FORMAT_RNG = TEMP_RNG
        GoSub RNG_FORMAT_BORDERS_TB
    
        For k = 1 To 2
            j = 1
            If k = 1 Then
'----------------------------------------------------------------
                For i = 1 To NSIZE
                    For ii = 1 To 2 'ii = 1 if it's first row
                        If ii = 1 Then jj = 3 Else jj = i * h - h + 3
                            .Cells(jj, j).formula = "=" & .Cells(12 + h * (NSIZE + 2), 1).Address(False, False)
                            .Cells(jj, j).Font.Bold = True

                        If ii = 2 Then
                            .Cells(jj, j).formula = "=" & .Cells(11 + h * (NSIZE + 2) + i, 1).Address(False, False)
                            .Cells(jj, j).Font.Bold = True
                        End If
                    Next ii
                Next i
'----------------------------------------------------------------
            End If
            
            For j = 2 To 7
                For i = 1 To NSIZE
                    For ii = 1 To 2
                        If k = 1 Then
                            If ii = 1 Then jj = 3 Else jj = i * h - h + 3
                                .Cells(jj, j).formula = _
                                "=" & .Cells(12 + h * (NSIZE + 2), 10 + j).Address(False, False)
                                Set FORMAT_RNG = .Cells(jj, j)
                                    GoSub RNG_FORMAT_NUM_MULTIPLE
                            If ii = 2 Then
                                .Cells(jj, j).formula = _
                                "=" & .Cells(11 + h * (NSIZE + 2) + i, 10 + j).Address(False, False)
                            End If
                        Else
                            If ii = 1 Then jj = 4 Else jj = i * h - h + 4
                                If j = 2 Then
                                    .Cells(jj, j).formula = _
                                    "=" & .Cells(jj - 1, j).Address(False, False) _
                                    & "*" & .Cells(10 + h * (NSIZE + 2), 4).Address(False, False) _
                                    & "/" & .Cells(10 + h * (NSIZE + 2), 19).Address(False, False)
                                    Set FORMAT_RNG = .Cells(jj, j)
                                        GoSub RNG_FORMAT_NUM_ACCOUNTING
                                    
                                ElseIf j = 3 Then
                                    .Cells(jj, j).formula = _
                                    "=" & .Cells(jj - 1, j).Address(False, False) _
                                    & "*" & .Cells(10 + h * (NSIZE + 2), 5).Address(False, False) _
                                    & "/" & .Cells(10 + h * (NSIZE + 2), 19).Address(False, False)
                                    Set FORMAT_RNG = .Cells(jj, j)
                                        GoSub RNG_FORMAT_NUM_ACCOUNTING
                                        
                                ElseIf j = 4 Then
                                    .Cells(jj, j).formula = _
                                    "=" & .Cells(jj - 1, j).Address(False, False) _
                                    & "*" & .Cells(10 + h * (NSIZE + 2), 2).Address(False, False) _
                                    & "/" & .Cells(10 + h * (NSIZE + 2), 19).Address(False, False)
                                    Set FORMAT_RNG = .Cells(jj, j)
                                        GoSub RNG_FORMAT_NUM_ACCOUNTING
                                
                                ElseIf j = 5 Then
                                    .Cells(jj, j).formula = _
                                    "=" & .Cells(jj - 1, j).Address(False, False) _
                                    & "*" & .Cells(10 + h * (NSIZE + 2), 7).Address(False, False) _
                                    & "/" & .Cells(10 + h * (NSIZE + 2), 19).Address(False, False)
                                    Set FORMAT_RNG = .Cells(jj, j)
                                        GoSub RNG_FORMAT_NUM_ACCOUNTING
                                
                                ElseIf j = 6 Then
                                    .Cells(jj, j).formula = _
                                    "=if(" & .Cells(jj - 1, j).Address(False, False) & "=0,0," _
                                    & "(" & .Cells(jj - 1, j).Address(False, False) _
                                    & "*" & .Cells(10 + h * (NSIZE + 2), 3).Address(False, False) _
                                    & "+" & .Cells(10 + h * (NSIZE + 2), 8).Address(False, False) _
                                    & "-" & .Cells(10 + h * (NSIZE + 2), 9).Address(False, False) _
                                    & "-" & .Cells(-2, 7).Address(False, False) _
                                    & "-" & .Cells(-1, 7).Address(False, False) _
                                    & ")/" & .Cells(10 + h * (NSIZE + 2), 19).Address(False, False) & ")"
                                    Set FORMAT_RNG = .Cells(jj, j)
                                        GoSub RNG_FORMAT_NUM_ACCOUNTING
                                
                                Else
                                    .Cells(jj, j).formula = _
                                    "=if(" & .Cells(jj - 1, j).Address(False, False) & "=0,0," _
                                    & "(" & .Cells(jj - 1, j).Address(False, False) _
                                    & "*" & .Cells(10 + h * (NSIZE + 2), 2).Address(False, False) _
                                    & "+" & .Cells(10 + h * (NSIZE + 2), 8).Address(False, False) _
                                    & "-" & .Cells(10 + h * (NSIZE + 2), 9).Address(False, False) _
                                    & "-" & .Cells(-2, 7).Address(False, False) _
                                    & "-" & .Cells(-1, 7).Address(False, False) _
                                    & ")/" & .Cells(10 + h * (NSIZE + 2), 19).Address(False, False) & ")"
                                    Set FORMAT_RNG = .Cells(jj, j)
                                        GoSub RNG_FORMAT_NUM_ACCOUNTING
                                End If
                                .Cells(jj, j).Font.Bold = True
                        End If
                    Next ii
                Next i
            Next j
        Next k
'---------------------------------------------------------------------------------------------------------------
'Nico revised
        For i = 0 To 1 'Perfect
            j = 1
            If i = 0 Then
                .Cells(h * (NSIZE + i) + 3, j) = "Average"
                    Set FORMAT_RNG = .Cells(h * (NSIZE + i) + 3, j)
                        GoSub RNG_FORMAT_BORDERS_TT
            Else
                .Cells(h * (NSIZE + i) + 3, j) = "Median"
                    Set FORMAT_RNG = .Cells(h * (NSIZE + i) + 3, j)
                        GoSub RNG_FORMAT_BORDERS_TT
            End If
            .Cells(h * (NSIZE + i) + 3, j).Font.Bold = True
            
            If i = 0 Then
                '------------------------------------------------------------------------------------------------
                For j = 2 To 7
                '------------------------------------------------------------------------------------------------
                    l = 1
                    ADDRESS_STR = .Cells(3 + h * (NSIZE) - h * l, j).Address(False, False)
                    For l = 2 To NSIZE
                        ADDRESS_STR = ADDRESS_STR & "," & .Cells(3 + h * (NSIZE) - h * l, j).Address(False, False)
                    Next l
                    .Cells(h * (NSIZE + i) + 3, j).formula = "=AVERAGE(" & ADDRESS_STR & ")"
                    Set FORMAT_RNG = .Cells(h * (NSIZE + i) + 3, j)
                        GoSub RNG_FORMAT_BORDERS_TT
                    Set FORMAT_RNG = .Cells(h * (NSIZE + i) + 3, j)
                        GoSub RNG_FORMAT_NUM_MULTIPLE
                
                    l = 1
                    ADDRESS_STR = .Cells(4 + h * (NSIZE + i) - h * l, j).Address(False, False)
                    For l = 2 To NSIZE
                        ADDRESS_STR = ADDRESS_STR & "," & .Cells(4 + h * (NSIZE + i) - h * l, j).Address(False, False)
                    Next l
                    .Cells(h * (NSIZE + i) + 4, j).formula = "=AVERAGE(" & ADDRESS_STR & ")"
                    .Cells(h * (NSIZE + i) + 4, j).Font.Bold = True
                    Set FORMAT_RNG = .Cells(h * (NSIZE + i) + 4, j)
                        GoSub RNG_FORMAT_NUM_ACCOUNTING
                '------------------------------------------------------------------------------------------------
                Next j
                '------------------------------------------------------------------------------------------------
            Else
                '------------------------------------------------------------------------------------------------
                For j = 2 To 7
                '------------------------------------------------------------------------------------------------
                    l = 1
                    ADDRESS_STR = .Cells(3 + h * (NSIZE + i - 1) - h * l, j).Address(False, False)
                    For l = 2 To NSIZE
                        ADDRESS_STR = ADDRESS_STR & "," & .Cells(3 + h * (NSIZE + i - 1) - h * l, j).Address(False, False)
                    Next l
                    .Cells(h * (NSIZE + i) + 3, j).formula = "=MEDIAN(" & ADDRESS_STR & ")"
                    Set FORMAT_RNG = .Cells(h * (NSIZE + i) + 3, j)
                        GoSub RNG_FORMAT_BORDERS_TT
                    Set FORMAT_RNG = .Cells(h * (NSIZE + i) + 3, j)
                        GoSub RNG_FORMAT_NUM_MULTIPLE
                '------------------------------------------------------------------------------------------------
                    l = 1
                    ADDRESS_STR = .Cells(4 + h * (NSIZE + i - 1) - h * l, j).Address(False, False)
                    For l = 2 To NSIZE
                        ADDRESS_STR = ADDRESS_STR & "," & .Cells(4 + h * (NSIZE + i - 1) - h * l, j).Address(False, False)
                    Next l
                    .Cells(h * (NSIZE + i) + 4, j).formula = "=MEDIAN(" & ADDRESS_STR & ")"
                    .Cells(h * (NSIZE + i) + 4, j).Font.Bold = True
                    Set FORMAT_RNG = .Cells(h * (NSIZE + i) + 4, j)
                        GoSub RNG_FORMAT_NUM_ACCOUNTING
                '------------------------------------------------------------------------------------------------
                Next j
                '------------------------------------------------------------------------------------------------
            End If
        Next i
    End With
'---------------------------------------------------------------------------
End With
'---------------------------------------------------------------------------

'---------------------------------------------------------------------------
RNG_COMPARABLE_TRANSACTIONS_FUNC = True
'---------------------------------------------------------------------------
Exit Function
'---------------------------------------------------------------------------
RNG_FORMAT_CELLS:
'---------------------------------------------------------------------------
    With FORMAT_RNG
        .Style = "Comma"
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
        .RowHeight = 15
        .ColumnWidth = 15
        With .Columns("A")
            .ColumnWidth = 3
        End With
        With .Font
            .name = "Arial"
            .Size = 8
        End With
    End With
Return
'---------------------------------------------------------------------------
FORMAT_MAINHEAD:
'---------------------------------------------------------------------------
    With FORMAT_RNG
        .HorizontalAlignment = xlGeneral
        With .Font
            .Size = 12
            .Bold = True
        End With
    End With
Return
'---------------------------------------------------------------------------
FORMAT_SUBHEAD:
'---------------------------------------------------------------------------
    With FORMAT_RNG
        .HorizontalAlignment = xlGeneral
        With .Font
            .Size = 9
            .Italic = True
        End With
    End With
Return
'---------------------------------------------------------------------------
RNG_FORMAT_BORDERS_TB:
'---------------------------------------------------------------------------
    With FORMAT_RNG
        With .Borders(xlEdgeTop)
            .LineStyle = xlContinuous
            .ColorIndex = xlAutomatic
            .TintAndShade = 0
            .WEIGHT = xlMedium
        End With
        With .Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .ColorIndex = xlAutomatic
            .TintAndShade = 0
            .WEIGHT = xlMedium
        End With
    
    End With
Return
'---------------------------------------------------------------------------
RNG_FORMAT_BORDERS_BL:
'---------------------------------------------------------------------------
    With FORMAT_RNG.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .WEIGHT = xlThin
    End With
Return
'---------------------------------------------------------------------------
RNG_FORMAT_BORDERS_TT:
'---------------------------------------------------------------------------
    With FORMAT_RNG.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .WEIGHT = xlMedium
    End With
Return
'---------------------------------------------------------------------------
RNG_FORMAT_BORDERS_BOX:
'---------------------------------------------------------------------------
    With FORMAT_RNG
        With .Borders(xlEdgeLeft)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .WEIGHT = xlThin
        End With
        With .Borders(xlEdgeTop)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .WEIGHT = xlThin
        End With
        With .Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .WEIGHT = xlThin
        End With
        With .Borders(xlEdgeRight)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .WEIGHT = xlThin
        End With
    End With
Return
'---------------------------------------------------------------------------
RNG_FORMAT_NUM_MULTIPLE:
'---------------------------------------------------------------------------
    With FORMAT_RNG
        .NumberFormat = "#,##0.00 x;[Red](#,##0.00 x)"
    End With
Return
'---------------------------------------------------------------------------
RNG_FORMAT_NUM_ACCOUNTING:
'---------------------------------------------------------------------------
    With FORMAT_RNG
        .NumberFormat = "$#,##0.00;[Red]($#,##0.00)"
    End With
Return
'---------------------------------------------------------------------------
ERROR_LABEL:
RNG_COMPARABLE_TRANSACTIONS_FUNC = False
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : RNG_COMPARABLE_SCREENER_FUNC
'DESCRIPTION   :
'LIBRARY       : FUNDAMENTAL
'GROUP         : COMPARABLES
'ID            : 005
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 08/04/2009
'************************************************************************************
'************************************************************************************
        
Private Function RNG_COMPARABLE_SCREENER_FUNC(ByRef DATA_RNG As Variant, _
ByRef SCREENER_RNG As Variant)

Dim i As Long
Dim j As Long
Dim k As Long
Dim l As Long

Dim ii As Long
Dim jj As Long

Dim NROWS As Long
Dim NCOLUMNS As Long

Dim INDEX_ARR As Variant
Dim SYMBOLS_ARR As Variant
Dim SCREENER_VECTOR As Variant
Dim TEMP_MATRIX As Variant
Dim DATA_MATRIX As Variant
Dim MATCH_FLAG As Boolean

On Error GoTo ERROR_LABEL

DATA_MATRIX = DATA_RNG
NROWS = UBound(DATA_MATRIX, 1)
SCREENER_VECTOR = SCREENER_RNG
If UBound(SCREENER_VECTOR, 1) = 1 Then
    SCREENER_VECTOR = MATRIX_TRANSPOSE_FUNC(SCREENER_VECTOR)
End If

INDEX_ARR = Array(2, 17, 21, 22, 13, 14, 24, 28)
NCOLUMNS = UBound(INDEX_ARR)

l = 0
ReDim SYMBOLS_ARR(1 To l + 1)
For i = 3 To NROWS 'Skip Headings & Reference
    MATCH_FLAG = True
    For j = 1 To NCOLUMNS 'No Criteria
        If SCREENER_VECTOR(j, 1) = "" Then: GoTo 1984 'Unless it is blank
        k = INDEX_ARR(j)
        If DATA_MATRIX(2, k) = "" Or DATA_MATRIX(2, k) = 0 Then: GoTo 1984 'Unless it is blank
        
        If DATA_MATRIX(i, k) = "" Or DATA_MATRIX(i, k) = 0 Then
            MATCH_FLAG = False
            Exit For
        ElseIf Abs(DATA_MATRIX(i, k) / DATA_MATRIX(2, k) - 1) > SCREENER_VECTOR(j, 1) Then
            MATCH_FLAG = False
            Exit For
        End If
1984:
    Next j
    If MATCH_FLAG = True Then
        l = l + 1
        ReDim Preserve SYMBOLS_ARR(1 To l)
        SYMBOLS_ARR(l) = i
    End If
Next i
If l = 0 Then: GoTo ERROR_LABEL
INDEX_ARR = Array(1, 2, 4, 5, 7, 8, 9, 10, 11, 12, 17, 21, 22, 23, 25, 27)
NROWS = l
NCOLUMNS = UBound(INDEX_ARR)

ReDim TEMP_MATRIX(1 To NROWS + 2, 1 To NCOLUMNS)
For jj = 1 To NCOLUMNS
    j = INDEX_ARR(jj)
    TEMP_MATRIX(1, jj) = DATA_MATRIX(1, j)
    TEMP_MATRIX(2, jj) = DATA_MATRIX(2, j)
    For ii = 1 To NROWS
        i = SYMBOLS_ARR(ii)
        TEMP_MATRIX(ii + 2, jj) = DATA_MATRIX(i, j)
    Next ii
Next jj

RNG_COMPARABLE_SCREENER_FUNC = TEMP_MATRIX

Exit Function
ERROR_LABEL:
RNG_COMPARABLE_SCREENER_FUNC = Err.number
End Function
