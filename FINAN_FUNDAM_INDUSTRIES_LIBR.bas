Attribute VB_Name = "FINAN_FUNDAM_INDUSTRIES_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.

Private Const P0 = 0
' --> VARIABLE = INDUSTRY ( INDUSTRIES / COMPANIES POSITION )
Private Const p1 = 1
' --> VARIABLE = NO_FIRMS ( NO. OF FIRMS POSITION )
Private Const p2 = 2
' --> VARIABLE = DIVD ( DIVIDENDS POSITION )
Private Const P3 = 3
' --> VARIABLE = VL_BETA ( VL BETA POSITION )
Private Const P4 = 4
' --> VARIABLE = CURRENT_PE ( CURRENT P / E POSITION )
Private Const P5 = 5
' --> VARIABLE = FORDWARD_PE ( FORWARD PE POSITION )
Private Const P6 = 6
' --> VARIABLE = TRAILING_PE ( TRAILING PE POSITION )
Private Const P7 = 7
' --> VARIABLE = DIVD_YIELD ( DIVID YLD POSITION )
Private Const P8 = 8
' --> VARIABLE = TAX_RATE ( TAX RATE POSITION )
Private Const P9 = 9
' --> VARIABLE = INSIDERS ( INSIDERS % POSITION )
Private Const P10 = 10
' --> VARIABLE = INSTITUT_INVEST ( INSTITUTIONAL % POSITION )
Private Const P11 = 11
' --> VARIABLE = MARKET_CAP ( MARKET CAP POSITION )
Private Const P12 = 12
' --> VARIABLE = TOTAL_DEBT ( TOTAL DEBT POSITION )
Private Const P13 = 13
' --> VARIABLE = FIRM_VALUE ( FIRM VALUE POSITION )
Private Const P14 = 14
' --> VARIABLE = EV ( ENTERPRISE VALUE POSITION )
Private Const P15 = 15
' --> VARIABLE = INV_CAPITAL ( INVESTED CAPITAL (NET OF CASH) POSITION )
Private Const P16 = 16
' --> VARIABLE = NON_CASH_WC ( NON-CASH WC POSITION )
Private Const P17 = 17
' --> VARIABLE = CHG_NON_CASH_WC ( Change in non-CASH WC Position )
Private Const P18 = 18
' --> VARIABLE = REINVEST ( Reinvestment Position )
Private Const P19 = 19
' --> VARIABLE = SALES ( SALES Position )
Private Const P20 = 20
' --> VARIABLE = SGA_EXPENSES ( SG&A Expenses Position )
Private Const P21 = 21
' --> VARIABLE = EBIT ( EBIT Position )
Private Const P22 = 22
' --> VARIABLE = EBITDA ( EBITDA Position )
Private Const P23 = 23
' --> VARIABLE = EBIT_ATAX ( EBIT (1 - t) Position )
Private Const P24 = 24
' --> VARIABLE = DEPRECIATION ( Depreciation Position )
Private Const p25 = 25
' --> VARIABLE = CAPEX ( Capital Expenditures Position )
Private Const P26 = 26
' --> VARIABLE = NET_INCOME ( Net Income Position )
Private Const P27 = 27
' --> VARIABLE = TRAILING_NET_INCOME ( Trailing Net Income Position )
Private Const P28 = 28
' --> VARIABLE = CASH ( CASH Position )
Private Const P29 = 29
' --> VARIABLE = AR ( Acc Rec Position )
Private Const P30 = 30
' --> VARIABLE = INVENTORY ( Inventory Position )
Private Const P31 = 31
' --> VARIABLE = NET_PLANT ( Net Plant Position )
Private Const P32 = 32
' --> VARIABLE = TOTAL_ASSETS ( Total Assets Position )
Private Const P33 = 33
' --> VARIABLE = AP ( Acc Payable Position )
Private Const P34 = 34
' --> VARIABLE = BV_EQUITY ( BV Equity Position )
Private Const P35 = 35
' --> VARIABLE = SALES_GROWTH_5YR ( SALES Growth: 5 yr Position )
Private Const P36 = 36
' --> VARIABLE = EPS_GROWTH_5YR ( EPS Growth: 5 yr Position )
Private Const P37 = 37
' --> VARIABLE = RETURN_5YR ( Return: 5 yr Position )
Private Const P38 = 38
' --> VARIABLE = BETA_5YR ( Beta: 5 yr Position )
Private Const P39 = 39
' --> VARIABLE = CORRELATION ( Estd correlation Position )
Private Const P40 = 40
' --> VARIABLE = ST_DEV_3YR ( Std dev: 3 yr Position )
Private Const P41 = 41
' --> VARIABLE = CURRENT_EPS_GROWTH ( Proj EPS Gr Position )
Private Const P42 = 42
' --> VARIABLE = DIVD_GROWTH_5YR ( Dividend Growth: 5 yr Position )

Private PUB_NSIZE As Long


'************************************************************************************
'************************************************************************************
'FUNCTION      : PRINT_INDUSTRIES_REPORTS_FUNC
'DESCRIPTION   : PRINT INDUSTRY REPORT
'LIBRARY       : FUNDAMENTAL
'GROUP         : INDUSTRIES
'ID            : 001
'LAST UPDATE   : 14 / 07 / 2009
'AUTHORS       : RAFAEL NICOLAS FERMIN COTA & CHRISTOPHER GILPIN
'REFERENCES    : http://pages.stern.nyu.edu/~adamodar/
'************************************************************************************
'************************************************************************************

Public Function PRINT_INDUSTRIES_REPORTS_FUNC(ByRef SRC_RNG As Excel.Range, _
ByRef DST_RNG As Excel.Range, _
ByRef LABELS_RNG As Excel.Range, _
ByRef CONTROL_SHAPE As Object, _
ByVal NSIZE As Long)

Dim SRC_SHEET As Excel.Worksheet
Dim DST_SHEET As Excel.Worksheet

Dim REPORT_FLAG As Boolean

On Error GoTo ERROR_LABEL

PRINT_INDUSTRIES_REPORTS_FUNC = False

Set SRC_SHEET = SRC_RNG.Worksheet
Set DST_SHEET = DST_RNG.Worksheet

PUB_NSIZE = NSIZE
DST_RNG.CurrentRegion.ClearContents

Select Case CONTROL_SHAPE.ControlFormat.value
Case 1
    LABELS_RNG.value = "Accounts Payable Model"
    REPORT_FLAG = INDUSTRIES_ACCOUNTS_PAYABLE_FUNC(SRC_RNG, DST_RNG)
Case 2
    LABELS_RNG.value = "Accounts Receivable Model"
    REPORT_FLAG = INDUSTRIES_ACCOUNTS_RECEIVABLE_FUNC(SRC_RNG, DST_RNG)
Case 3
    LABELS_RNG.value = "Beta Model"
    REPORT_FLAG = INDUSTRIES_BETA_FUNC(SRC_RNG, DST_RNG)
Case 4
    LABELS_RNG.value = "CAPEX Model"
    REPORT_FLAG = INDUSTRIES_CAPEX_FUNC(SRC_RNG, DST_RNG)
Case 5
    LABELS_RNG.value = "CASH Flow Model"
    REPORT_FLAG = INDUSTRIES_CASH_FLOW_FUNC(SRC_RNG, DST_RNG)
Case 6
    LABELS_RNG.value = "CASH Model"
    REPORT_FLAG = INDUSTRIES_CASH_FUNC(SRC_RNG, DST_RNG)
Case 7
    LABELS_RNG.value = "DEBT Model"
    REPORT_FLAG = INDUSTRIES_DEBT_FUNC(SRC_RNG, DST_RNG)
Case 8
    LABELS_RNG.value = "DIVIDENDS Model"
    REPORT_FLAG = INDUSTRIES_DIVIDENDS_FUNC(SRC_RNG, DST_RNG)
Case 9
    LABELS_RNG.value = "EBITDA Model"
    REPORT_FLAG = INDUSTRIES_EBITDA_FUNC(SRC_RNG, DST_RNG)
Case 10
    LABELS_RNG.value = "Economic Value Added Model"
    REPORT_FLAG = INDUSTRIES_EVA_FUNC(SRC_RNG, DST_RNG)
Case 11
    LABELS_RNG.value = "EFFECTIVE TAX RATE MODEL"
    REPORT_FLAG = INDUSTRIES_TAX_FUNC(SRC_RNG, DST_RNG)
Case 12
    LABELS_RNG.value = "FREE CASH FLOW MODEL"
    REPORT_FLAG = INDUSTRIES_FCF_FUNC(SRC_RNG, DST_RNG)
Case 13
    LABELS_RNG.value = "Free Cash Flow to Equity Model"
    REPORT_FLAG = INDUSTRIES_FCFE_FUNC(SRC_RNG, DST_RNG)
Case 14
    LABELS_RNG.value = "Growth Model"
    REPORT_FLAG = INDUSTRIES_GROWTH_FUNC(SRC_RNG, DST_RNG)
Case 15
    LABELS_RNG.value = " Historical Growth Model"
    REPORT_FLAG = INDUSTRIES_HISTORICAL_GROWTH_FUNC(SRC_RNG, DST_RNG)
Case 16
    LABELS_RNG.value = " Holdings Model"
    REPORT_FLAG = INDUSTRIES_HOLDINGS_FUNC(SRC_RNG, DST_RNG)
Case 17
    LABELS_RNG.value = "Inventory Model"
    REPORT_FLAG = INDUSTRIES_INVENTORIES_FUNC(SRC_RNG, DST_RNG)
Case 18
    LABELS_RNG.value = "Margins Model"
    REPORT_FLAG = INDUSTRIES_MARGINS_FUNC(SRC_RNG, DST_RNG)
Case 19
    LABELS_RNG.value = "PE Model"
    REPORT_FLAG = INDUSTRIES_PE_FUNC(SRC_RNG, DST_RNG)
Case 20
    LABELS_RNG.value = "Price to Book Value Model"
    REPORT_FLAG = INDUSTRIES_PB_FUNC(SRC_RNG, DST_RNG)
Case 21
    LABELS_RNG.value = "Regression Stats Model"
    REPORT_FLAG = INDUSTRIES_PS_FUNC(SRC_RNG, DST_RNG)
Case 22
    LABELS_RNG.value = "Price to Sale Model"
    REPORT_FLAG = INDUSTRIES_REGRESSION_FUNC(SRC_RNG, DST_RNG)
Case 23
    LABELS_RNG.value = "Retention Model"
    REPORT_FLAG = INDUSTRIES_RETENTION_FUNC(SRC_RNG, DST_RNG)
Case 24
    LABELS_RNG.value = "Return on Capital Model"
    REPORT_FLAG = INDUSTRIES_ROC_FUNC(SRC_RNG, DST_RNG)
Case 25
    LABELS_RNG.value = "Return on Equity Model"
    REPORT_FLAG = INDUSTRIES_ROE_FUNC(SRC_RNG, DST_RNG)
Case 26
    LABELS_RNG.value = "Standard Deviation Model"
    REPORT_FLAG = INDUSTRIES_VOLATILITY_FUNC(SRC_RNG, DST_RNG)
Case 27
    LABELS_RNG.value = "Valuation Model"
    REPORT_FLAG = INDUSTRIES_VALUATION_FUNC(SRC_RNG, DST_RNG)
Case 28
    LABELS_RNG.value = "WACC Model"
    REPORT_FLAG = INDUSTRIES_WACC_FUNC(SRC_RNG, DST_RNG)
Case 29
    LABELS_RNG.value = "Working Capital Model"
    REPORT_FLAG = INDUSTRIES_WORKING_CAPITAL_FUNC(SRC_RNG, DST_RNG)
Case Else

End Select

If REPORT_FLAG = False Then: GoTo ERROR_LABEL
PRINT_INDUSTRIES_REPORTS_FUNC = True

Exit Function
ERROR_LABEL:
PRINT_INDUSTRIES_REPORTS_FUNC = False
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : INDUSTRIES_REPORT_HEADINGS_FUNC
'DESCRIPTION   : For many of the ratios, estimated on a sector/industry basis,
'we use the cumulated values. As an example, the PE ratio for
'the sector is not a simple average of the PE ratios of individual firms
'in the sector. Instead, it is obtained by dividing the cumulated net income
'for the sector (obtained by adding up the net income of each firm in the
'sector) by the cumulated market value of equity of firm in the sector (
'obtained by adding up the market values of all of the firms in the sector).

'LIBRARY       : FUNDAMENTAL
'GROUP         : INDUSTRIES
'ID            : 002
'LAST UPDATE   : 14 / 07 / 2009
'AUTHORS       : RAFAEL NICOLAS FERMIN COTA & CHRISTOPHER GILPIN
'REFERENCES    : http://pages.stern.nyu.edu/~adamodar/
'************************************************************************************
'************************************************************************************

Public Function INDUSTRIES_REPORT_HEADINGS_FUNC()

INDUSTRIES_REPORT_HEADINGS_FUNC = _
    Array("INDUSTRY/COMPANY NAME", "NO FIRMS/EXCHANGE", " DIVIDENDS", _
    " VALUE LINE BETA", " CURRENT PE", _
    " FORWARD PE", " TRAILING PE", _
    " DIVIDEND YIELD", " EFF TAX RATE", _
    " INSIDER HOLDINGS", " INSTITUTIONAL HOLDINGS", _
    " MARKET CAP", " TOTAL DEBT", " FIRM VALUE", _
    " ENTERPRISE VALUE", " INVESTED CAPITAL", " NON-CASH WC", _
    " CHG IN NON-CASH WC", " REINVESTMENT", _
    " REVENUES: LAST YR", " SG&A EXPENSES", " EBIT", _
    " EBITDA", " EBIT(1-T)", " DEPRECIATION", _
    " CAPITAL EXPENDITURES", " NET INCOME", _
    " TRAILING NET INCOME", " CASH", " ACCOUNTS RECEIVABLE", _
    " INVENTORIES", " NET PLANT", " TOTAL ASSETS", _
    " ACCOUNTS PAYABLE", " SHAREHOLDERS EQUITY", _
    " SALES GROWTH 1-YEAR", " EPS GROWTH 5-YEAR", _
    " TOTAL RETURN 5-YEAR", " BETA 5-YEAR", _
    " CORRELATION", " STD DEV 3-YEAR", _
    " PROJ EPS GROWTH RATE", " DIVIDEND GROWTH 5-YEAR")

'   "SUM OF DIVIDENDS", AVERAGE OF VALUE LINE BETA", "AVERAGE OF CURRENT PE", _
    "AVERAGE OF FORWARD PE", "AVERAGE OF TRAILING PE", _
    "AVERAGE OF DIVIDEND YIELD", "AVERAGE OF EFF TAX RATE", _
    "AVERAGE OF INSIDER HOLDINGS", "AVERAGE OF INSTITUTIONAL HOLDINGS", _
    "SUM OF MARKET CAP", "SUM OF TOTAL DEBT", "SUM OF FIRM VALUE", _
    "SUM OF ENTERPRISE VALUE", "SUM OF INVESTED CAPITAL", "SUM OF NON-CASH WC", _
    "SUM OF CHG IN NON-CASH WC", "SUM OF REINVESTMENT", _
    "SUM OF REVENUES: LAST YR", "SUM OF SG&A EXPENSES", "SUM OF EBIT", _
    "SUM OF EBITDA", "SUM OF EBIT(1-T)", "SUM OF DEPRECIATION", _
    "SUM OF CAPITAL EXPENDITURES", "SUM OF NET INCOME", _
    "SUM OF TRAILING NET INCOME", "SUM OF CASH", "SUM OF ACCOUNTS RECEIVABLE", _
    "SUM OF INVENTORIES", "SUM OF NET PLANT", "SUM OF TOTAL ASSETS", _
    "SUM OF ACCOUNTS PAYABLE", "SUM OF SHAREHOLDERS EQUITY", _
    "AVERAGE OF SALES GROWTH 1-YEAR", "AVERAGE OF EPS GROWTH 5-YEAR", _
    "AVERAGE OF TOTAL RETURN 5-YEAR", "AVERAGE OF BETA 5-YEAR", _
    "AVERAGE OF CORRELATION", "AVERAGE OF STD DEV 3-YEAR", _
    "AVERAGE OF PROJ EPS GROWTH RATE", "AVERAGE OF DIVIDEND GROWTH 5-YEAR")
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : INDUSTRIES_ACCOUNTS_PAYABLE_FUNC
'DESCRIPTION   : ACCOUNTS PAYABLE TABLE
'LIBRARY       : FUNDAMENTAL
'GROUP         : INDUSTRIES
'ID            : 003
'LAST UPDATE   : 14 / 07 / 2009
'AUTHORS       : RAFAEL NICOLAS FERMIN COTA & CHRISTOPHER GILPIN
'REFERENCES    : http://pages.stern.nyu.edu/~adamodar/
'************************************************************************************
'************************************************************************************

Private Function INDUSTRIES_ACCOUNTS_PAYABLE_FUNC(ByRef SRC_RNG As Excel.Range, _
ByRef DST_RNG As Excel.Range)

Dim i As Long

Dim SRC_WSHEET_NAME As String

Dim INDUSTRY As String
Dim NO_FIRMS As String
Dim SALES As String
Dim AP As String

On Error GoTo ERROR_LABEL

INDUSTRIES_ACCOUNTS_PAYABLE_FUNC = False

SRC_WSHEET_NAME = SRC_RNG.Worksheet.name & "!"

Range(DST_RNG.Cells(1, 1), DST_RNG.Cells(1, 1 + 3)).value = _
    Array("Industries/Companies", _
    "No. of Firms/Exchange", _
    "Accounts Payable/SALES", _
    "Number of Days of SALES in Accounts Payable")

With DST_RNG
    For i = 1 To PUB_NSIZE
        INDUSTRY = SRC_RNG.Offset(i, P0).Address
        NO_FIRMS = SRC_RNG.Offset(i, p1).Address
        SALES = SRC_RNG.Offset(i, P19).Address
        AP = SRC_RNG.Offset(i, P33).Address
        
        .Offset(i, 0).formula = "=" & SRC_WSHEET_NAME & INDUSTRY
        .Offset(i, 1).formula = "=" & SRC_WSHEET_NAME & NO_FIRMS
        .Offset(i, 2).formula = "=IF(" & SRC_WSHEET_NAME & SALES & ">0," & _
            SRC_WSHEET_NAME & AP & "/" & SRC_WSHEET_NAME & SALES & ",""NA"")"
        .Offset(i, 3).formula = "=IF(" & SRC_WSHEET_NAME & SALES & ">0," & _
            SRC_WSHEET_NAME & AP & "/(" & SRC_WSHEET_NAME & SALES & "/365),""NA"")"
    Next i
End With

INDUSTRIES_ACCOUNTS_PAYABLE_FUNC = True

Exit Function
ERROR_LABEL:
INDUSTRIES_ACCOUNTS_PAYABLE_FUNC = False
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : INDUSTRIES_ACCOUNTS_RECEIVABLE_FUNC
'DESCRIPTION   : ACCOUNTS RECEIVABLE TABLE
'LIBRARY       : FUNDAMENTAL
'GROUP         : INDUSTRIES
'ID            : 004
'LAST UPDATE   : 14 / 07 / 2009
'AUTHORS       : RAFAEL NICOLAS FERMIN COTA & CHRISTOPHER GILPIN
'REFERENCES    : http://pages.stern.nyu.edu/~adamodar/
'************************************************************************************
'************************************************************************************

Private Function INDUSTRIES_ACCOUNTS_RECEIVABLE_FUNC(ByRef SRC_RNG As Excel.Range, _
ByRef DST_RNG As Excel.Range)

Dim i As Long
Dim SRC_WSHEET_NAME As String

Dim INDUSTRY As String
Dim NO_FIRMS As String
Dim SALES As String
Dim AR As String
Dim EV As String

On Error GoTo ERROR_LABEL

INDUSTRIES_ACCOUNTS_RECEIVABLE_FUNC = False
SRC_WSHEET_NAME = SRC_RNG.Worksheet.name & "!"

Range(DST_RNG.Cells(1, 1), DST_RNG.Cells(1, 1 + 4)).value = _
    Array("Industries/Companies", _
            "No. of Firms/Exchange", _
            "Accounts Receivable/SALES", _
            "Number of Days of SALES in Accounts Receivable", _
            "Accounts Receivable/ Enterprise Value")
With DST_RNG
    For i = 1 To PUB_NSIZE
        INDUSTRY = SRC_RNG.Offset(i, P0).Address
        NO_FIRMS = SRC_RNG.Offset(i, p1).Address
        SALES = SRC_RNG.Offset(i, P19).Address
        AR = SRC_RNG.Offset(i, P29).Address
        EV = SRC_RNG.Offset(i, P14).Address
        
        .Offset(i, 0).formula = "=" & SRC_WSHEET_NAME & INDUSTRY
        
        .Offset(i, 1).formula = "=" & SRC_WSHEET_NAME & NO_FIRMS
        
        .Offset(i, 2).formula = "=IF(" & SRC_WSHEET_NAME & SALES & ">0," & _
        SRC_WSHEET_NAME & AR & "/" & SRC_WSHEET_NAME & SALES & ",""NA"")"
        
        .Offset(i, 3).formula = "=IF(" & SRC_WSHEET_NAME & SALES & ">0," & _
        SRC_WSHEET_NAME & AR & "/(" & SRC_WSHEET_NAME & SALES & "/365),""NA"")"
        
        .Offset(i, 4).formula = "=" & SRC_WSHEET_NAME & AR & "/" & SRC_WSHEET_NAME & EV
    Next i
End With

INDUSTRIES_ACCOUNTS_RECEIVABLE_FUNC = True

Exit Function
ERROR_LABEL:
INDUSTRIES_ACCOUNTS_RECEIVABLE_FUNC = False
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : INDUSTRIES_BETA_FUNC
'DESCRIPTION   : BETA TABLE
'LIBRARY       : FUNDAMENTAL
'GROUP         : INDUSTRIES
'ID            : 005
'LAST UPDATE   : 14 / 07 / 2009
'AUTHORS       : RAFAEL NICOLAS FERMIN COTA & CHRISTOPHER GILPIN
'REFERENCES    : http://pages.stern.nyu.edu/~adamodar/
'************************************************************************************
'************************************************************************************

Private Function INDUSTRIES_BETA_FUNC(ByRef SRC_RNG As Excel.Range, _
ByRef DST_RNG As Excel.Range)

Dim i As Long
Dim SRC_WSHEET_NAME As String

Dim INDUSTRY As String
Dim NO_FIRMS As String
Dim VL_BETA As String
Dim BETA_5YR As String
Dim TOTAL_DEBT As String
Dim Market_Cap As String
Dim TAX_RATE As String
Dim FirmValue As String
Dim cash As String
Dim CORRELATION As String

On Error GoTo ERROR_LABEL

INDUSTRIES_BETA_FUNC = False
SRC_WSHEET_NAME = SRC_RNG.Worksheet.name & "!"

Range(DST_RNG.Cells(1, 1), DST_RNG.Cells(1, 1 + 9)).value = _
    Array("Industries/Companies", _
            "No. of Firms/Exchange", _
            "Average Beta", _
            "Market D/E Ratio", _
            "Tax Rate", _
            "Unlevered Beta", _
            "CASH/Firm Value", _
            "Unlevered Beta corrected for CASH", _
            "Correlation with market", _
            "Total Beta (Unlevered)")

With DST_RNG
    For i = 1 To PUB_NSIZE
        INDUSTRY = SRC_RNG.Offset(i, P0).Address
        NO_FIRMS = SRC_RNG.Offset(i, p1).Address
        VL_BETA = SRC_RNG.Offset(i, P3).Address
        BETA_5YR = SRC_RNG.Offset(i, P38).Address
        TOTAL_DEBT = SRC_RNG.Offset(i, P12).Address
        Market_Cap = SRC_RNG.Offset(i, P11).Address
        TAX_RATE = SRC_RNG.Offset(i, P8).Address
        FirmValue = SRC_RNG.Offset(i, P13).Address
        cash = SRC_RNG.Offset(i, P28).Address
        CORRELATION = SRC_RNG.Offset(i, P39).Address
        
        .Offset(i, 0).formula = "=" & SRC_WSHEET_NAME & INDUSTRY 'INDUSTRY
        .Offset(i, 1).formula = "=" & SRC_WSHEET_NAME & NO_FIRMS 'NO_FIRMS
        
        .Offset(i, 2).formula = "=MAX(" & SRC_WSHEET_NAME & VL_BETA & "," & _
        SRC_WSHEET_NAME & BETA_5YR & ")" 'AvgBeta
        
        .Offset(i, 3).formula = "=" & SRC_WSHEET_NAME & TOTAL_DEBT & "/" & _
        SRC_WSHEET_NAME & Market_Cap 'MarketDE
        
        .Offset(i, 4).formula = "=IF(" & SRC_WSHEET_NAME & TAX_RATE & _
        ">0.5,0.5,IF(" & SRC_WSHEET_NAME & TAX_RATE & "<0,0," & SRC_WSHEET_NAME & _
        TAX_RATE & "))" 'TAX_RATE
        
        .Offset(i, 5).formula = "=" & .Offset(i, 2).Address & "/(1+(1-" & _
        .Offset(i, 4).Address & ")* " & .Offset(i, 3).Address & ")" 'UnLevBeta
        
        .Offset(i, 6).formula = "=" & SRC_WSHEET_NAME & cash & "/" & _
        SRC_WSHEET_NAME & FirmValue 'CASHFirmVal
        
        .Offset(i, 7).formula = "=" & .Offset(i, 5).Address & "/(1-" & _
        .Offset(i, 6).Address & ")" 'UnlevBetaCorrected
        
        .Offset(i, 8).formula = "=" & SRC_WSHEET_NAME & CORRELATION
        
        .Offset(i, 9).formula = "=" & .Offset(i, 7).Address & "/" & _
        .Offset(i, 8).Address 'TotalBeta
    
    Next i
End With

INDUSTRIES_BETA_FUNC = True

Exit Function
ERROR_LABEL:
INDUSTRIES_BETA_FUNC = False
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : INDUSTRIES_CAPEX_FUNC
'DESCRIPTION   : CAPEX TABLE
'LIBRARY       : FUNDAMENTAL
'GROUP         : INDUSTRIES
'ID            : 006
'LAST UPDATE   : 14 / 07 / 2009
'AUTHORS       : RAFAEL NICOLAS FERMIN COTA & CHRISTOPHER GILPIN
'REFERENCES    : http://pages.stern.nyu.edu/~adamodar/
'************************************************************************************
'************************************************************************************

Private Function INDUSTRIES_CAPEX_FUNC(ByRef SRC_RNG As Excel.Range, _
ByRef DST_RNG As Excel.Range)

Dim i As Long
Dim SRC_WSHEET_NAME As String

Dim INDUSTRY As String
Dim NO_FIRMS As String
Dim CAPEX As String
Dim DEPRECIATION As String
Dim SALES As String
Dim EBIT_ATAX As String
Dim INV_CAPITAL As String

On Error GoTo ERROR_LABEL

INDUSTRIES_CAPEX_FUNC = False

SRC_WSHEET_NAME = SRC_RNG.Worksheet.name & "!"


Range(DST_RNG.Cells(1, 1), _
DST_RNG.Cells(1, 1 + 7)).value = _
Array("Industries/Companies", "No. of Firms/Exchange", "Capital Expenditures", _
"Depreciation", "Cap Ex / Deprecn", "Net Cap Ex/SALES", "Net Cap Ex/ EBIT (1-t)", _
"SALES/Capital")

With DST_RNG
    For i = 1 To PUB_NSIZE
        INDUSTRY = SRC_RNG.Offset(i, P0).Address
        NO_FIRMS = SRC_RNG.Offset(i, p1).Address
        CAPEX = SRC_RNG.Offset(i, p25).Address
        DEPRECIATION = SRC_RNG.Offset(i, P24).Address
        SALES = SRC_RNG.Offset(i, P19).Address
        EBIT_ATAX = SRC_RNG.Offset(i, P23).Address 'EBIT (1-T)
        INV_CAPITAL = SRC_RNG.Offset(i, P15).Address 'Invested Capital (net of CASH)
        
        .Offset(i, 0).formula = "=" & SRC_WSHEET_NAME & INDUSTRY 'INDUSTRY
        .Offset(i, 1).formula = "=" & SRC_WSHEET_NAME & NO_FIRMS 'NO_FIRMS
        
        .Offset(i, 2).formula = "=" & SRC_WSHEET_NAME & CAPEX
        .Offset(i, 3).formula = "=" & SRC_WSHEET_NAME & DEPRECIATION
        
        .Offset(i, 4).formula = "=IF(" & .Offset(i, 3).Address & ">0," & _
        .Offset(i, 2).Address & "/" & .Offset(i, 3).Address & ", ""NA"")" 'Cap Ex/Deprecn
        
        .Offset(i, 5).formula = "=IF(" & SRC_WSHEET_NAME & SALES & ">0,(" & _
        .Offset(i, 2).Address & "-" & .Offset(i, 3).Address & ")" & "/" & _
        SRC_WSHEET_NAME & SALES & ", ""NA"")" ' NET Cap Ex/Deprecn
        
        .Offset(i, 6).formula = "=IF(" & SRC_WSHEET_NAME & EBIT_ATAX & ">0,(" & _
        .Offset(i, 2).Address & "-" & .Offset(i, 3).Address & ")" & "/" & _
        SRC_WSHEET_NAME & EBIT_ATAX & ", ""NA"")" ' NET Cap Ex/EBIT_ATAX
        
        .Offset(i, 7).formula = "=IF(" & SRC_WSHEET_NAME & SALES & ">0," & _
        SRC_WSHEET_NAME & SALES & "/" & SRC_WSHEET_NAME & INV_CAPITAL & ", ""NA"")"
        'SALES/Capital
    Next i
End With

INDUSTRIES_CAPEX_FUNC = True

Exit Function
ERROR_LABEL:
INDUSTRIES_CAPEX_FUNC = False
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : INDUSTRIES_CASH_FLOW_FUNC
'DESCRIPTION   : CASH FLOW TABLE
'LIBRARY       : FUNDAMENTAL
'GROUP         : INDUSTRIES
'ID            : 007
'LAST UPDATE   : 14 / 07 / 2009
'AUTHORS       : RAFAEL NICOLAS FERMIN COTA & CHRISTOPHER GILPIN
'REFERENCES    : http://pages.stern.nyu.edu/~adamodar/
'************************************************************************************
'************************************************************************************

Private Function INDUSTRIES_CASH_FLOW_FUNC(ByRef SRC_RNG As Excel.Range, _
ByRef DST_RNG As Excel.Range)

Dim i As Long
Dim SRC_WSHEET_NAME As String

Dim INDUSTRY As String
Dim NO_FIRMS As String
Dim TAX_RATE As String
Dim EBIT_ATAX As String
Dim NON_CASH_WC As String
Dim SALES As String

On Error GoTo ERROR_LABEL

INDUSTRIES_CASH_FLOW_FUNC = False
SRC_WSHEET_NAME = SRC_RNG.Worksheet.name & "!"

Range(DST_RNG.Cells(1, 1), _
DST_RNG.Cells(1, 1 + 4)).value = _
Array("Industries/Companies", "No. of Firms/Exchange", "After-tax Operating Margin", _
"Tax Rates", "Non-CASH Working Capital as % of Revenues")

With DST_RNG
    For i = 1 To PUB_NSIZE
    
        INDUSTRY = SRC_RNG.Offset(i, P0).Address
        NO_FIRMS = SRC_RNG.Offset(i, p1).Address
        
        SALES = SRC_RNG.Offset(i, P19).Address
        EBIT_ATAX = SRC_RNG.Offset(i, P23).Address 'EBIT (1-T)
        
        TAX_RATE = SRC_RNG.Offset(i, P8).Address
        NON_CASH_WC = SRC_RNG.Offset(i, P16).Address
        
        .Offset(i, 0).formula = "=" & SRC_WSHEET_NAME & INDUSTRY 'INDUSTRY
        .Offset(i, 1).formula = "=" & SRC_WSHEET_NAME & NO_FIRMS 'NO_FIRMS
        
        .Offset(i, 2).formula = "=IF(" & SRC_WSHEET_NAME & SALES & ">0," & _
            SRC_WSHEET_NAME & EBIT_ATAX & "/" & SRC_WSHEET_NAME & SALES & _
            ", ""NA"")" 'After-tax Operating Margin
        
        .Offset(i, 3).formula = "=IF(" & SRC_WSHEET_NAME & TAX_RATE & ">0.5,0.5,IF(" & _
            SRC_WSHEET_NAME & TAX_RATE & "<0,0," & SRC_WSHEET_NAME & _
            TAX_RATE & "))" 'Tax Rate
        
        .Offset(i, 4).formula = "=IF(" & SRC_WSHEET_NAME & SALES & ">0," & _
            SRC_WSHEET_NAME & NON_CASH_WC & "/" & SRC_WSHEET_NAME & SALES & _
            ", ""NA"")" 'Non-CASH Working Capital as % of Revenues
    
    Next i
End With

INDUSTRIES_CASH_FLOW_FUNC = True

Exit Function
ERROR_LABEL:
INDUSTRIES_CASH_FLOW_FUNC = False
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : INDUSTRIES_CASH_FUNC
'DESCRIPTION   : CASH TABLE
'LIBRARY       : FUNDAMENTAL
'GROUP         : INDUSTRIES
'ID            : 008
'LAST UPDATE   : 14 / 07 / 2009
'AUTHORS       : RAFAEL NICOLAS FERMIN COTA & CHRISTOPHER GILPIN
'REFERENCES    : http://pages.stern.nyu.edu/~adamodar/
'************************************************************************************
'************************************************************************************

Private Function INDUSTRIES_CASH_FUNC(ByRef SRC_RNG As Excel.Range, _
ByRef DST_RNG As Excel.Range)

Dim i As Long
Dim SRC_WSHEET_NAME As String

Dim INDUSTRY As String
Dim NO_FIRMS As String
Dim cash As String
Dim FIRM_VALUE As String
Dim TOTAL_ASSETS As String
Dim SALES As String

On Error GoTo ERROR_LABEL

INDUSTRIES_CASH_FUNC = False

SRC_WSHEET_NAME = SRC_RNG.Worksheet.name & "!"

Range(DST_RNG.Cells(1, 1), _
DST_RNG.Cells(1, 1 + 4)).value = _
Array("Industries/Companies", "No. of Firms/Exchange", _
"CASH/Firm Value", "CASH/Revenues", "CASH/Total Assets")

With DST_RNG
    For i = 1 To PUB_NSIZE
    
        INDUSTRY = SRC_RNG.Offset(i, P0).Address
        NO_FIRMS = SRC_RNG.Offset(i, p1).Address
        
        cash = SRC_RNG.Offset(i, P28).Address
        FIRM_VALUE = SRC_RNG.Offset(i, P13).Address
        TOTAL_ASSETS = SRC_RNG.Offset(i, P32).Address
        SALES = SRC_RNG.Offset(i, P19).Address
        
        .Offset(i, 0).formula = "=" & SRC_WSHEET_NAME & INDUSTRY 'INDUSTRY
        .Offset(i, 1).formula = "=" & SRC_WSHEET_NAME & NO_FIRMS 'NO_FIRMS
        
        .Offset(i, 2).formula = "=" & SRC_WSHEET_NAME & cash & " / " & _
            SRC_WSHEET_NAME & FIRM_VALUE 'CASH / Firm Value"
        .Offset(i, 3).formula = "=IF(" & SRC_WSHEET_NAME & SALES & ">0," & _
            SRC_WSHEET_NAME & cash & "/" & SRC_WSHEET_NAME & SALES & _
            ", ""NA"")" 'CASH as % of Revenues
        .Offset(i, 4).formula = "=" & SRC_WSHEET_NAME & cash & " / " & _
            SRC_WSHEET_NAME & TOTAL_ASSETS 'CASH / Total Assets
    
    Next i
End With

INDUSTRIES_CASH_FUNC = True

Exit Function
ERROR_LABEL:
INDUSTRIES_CASH_FUNC = False
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : INDUSTRIES_DEBT_FUNC
'DESCRIPTION   : DEBT TABLE
'LIBRARY       : FUNDAMENTAL
'GROUP         : INDUSTRIES
'ID            : 009
'LAST UPDATE   : 14 / 07 / 2009
'AUTHORS       : RAFAEL NICOLAS FERMIN COTA & CHRISTOPHER GILPIN
'REFERENCES    : http://pages.stern.nyu.edu/~adamodar/
'************************************************************************************
'************************************************************************************

Private Function INDUSTRIES_DEBT_FUNC(ByRef SRC_RNG As Excel.Range, _
ByRef DST_RNG As Excel.Range)

Dim i As Long
Dim SRC_WSHEET_NAME As String

Dim INDUSTRY As String
Dim NO_FIRMS As String
Dim FIRM_VALUE As String
Dim TOTAL_DEBT As String
Dim INV_CAPITAL As String
Dim cash As String
Dim TAX_RATE As String
Dim INSIDERS As String
Dim ST_DEV_3YR As String
Dim EBITDA As String
Dim NET_PLANT As String
Dim CAPEX As String

On Error GoTo ERROR_LABEL

INDUSTRIES_DEBT_FUNC = False
SRC_WSHEET_NAME = SRC_RNG.Worksheet.name & "!"

Range(DST_RNG.Cells(1, 1), _
DST_RNG.Cells(1, 1 + 9)).value = _
Array("Industries/Companies", "No. of Firms/Exchange", "MV Debt Ratio", "BV Debt Ratio", _
"Effective Tax Rate", "INSIDERS Holdings", "Std Deviation in Prices", "EBITDA/Value", _
"Fixed Assets/BV of Capital", "Capital Spending/BV of Capital")

With DST_RNG
    For i = 1 To PUB_NSIZE
    
        INDUSTRY = SRC_RNG.Offset(i, P0).Address
        NO_FIRMS = SRC_RNG.Offset(i, p1).Address
        
        FIRM_VALUE = SRC_RNG.Offset(i, P13).Address
        TOTAL_DEBT = SRC_RNG.Offset(i, P12).Address
        
        INV_CAPITAL = SRC_RNG.Offset(i, P15).Address 'Invested Capital (net of CASH)
        cash = SRC_RNG.Offset(i, P28).Address
        
        TAX_RATE = SRC_RNG.Offset(i, P8).Address
        
        INSIDERS = SRC_RNG.Offset(i, P9).Address
        ST_DEV_3YR = SRC_RNG.Offset(i, P40).Address
        
        EBITDA = SRC_RNG.Offset(i, P22).Address
        
        NET_PLANT = SRC_RNG.Offset(i, P31).Address
        CAPEX = SRC_RNG.Offset(i, p25).Address
        
        .Offset(i, 0).formula = "=" & SRC_WSHEET_NAME & INDUSTRY 'INDUSTRY
        .Offset(i, 1).formula = "=" & SRC_WSHEET_NAME & NO_FIRMS 'NO_FIRMS
        
        .Offset(i, 2).formula = "=" & SRC_WSHEET_NAME & TOTAL_DEBT & " / " & _
            SRC_WSHEET_NAME & FIRM_VALUE 'MV Debt Ratio
        .Offset(i, 3).formula = "=" & SRC_WSHEET_NAME & TOTAL_DEBT & " / (" & _
            SRC_WSHEET_NAME & INV_CAPITAL & " + " & SRC_WSHEET_NAME & cash & ")" 'BV Debt Ratio
        
        .Offset(i, 4).formula = "=" & SRC_WSHEET_NAME & TAX_RATE 'Effective Tax Rate
        .Offset(i, 5).formula = "=" & SRC_WSHEET_NAME & INSIDERS 'INSIDER HOLDINGs
        .Offset(i, 6).formula = "=" & SRC_WSHEET_NAME & ST_DEV_3YR 'ST_DEVIATIONS
        
        .Offset(i, 7).formula = "=" & SRC_WSHEET_NAME & EBITDA & " / " & _
        SRC_WSHEET_NAME & FIRM_VALUE 'EBITDA per Value
        
        .Offset(i, 8).formula = "=" & SRC_WSHEET_NAME & NET_PLANT & " / (" & _
        SRC_WSHEET_NAME & INV_CAPITAL & " + " & SRC_WSHEET_NAME & cash & _
        ")" ''Fixed Assets/BV of Capital
        .Offset(i, 9).formula = "=" & SRC_WSHEET_NAME & CAPEX & " / (" & _
        SRC_WSHEET_NAME & INV_CAPITAL & " + " & SRC_WSHEET_NAME & cash & _
        ")" ''Capital Spending/BV of Capital

    Next i
End With

INDUSTRIES_DEBT_FUNC = True

Exit Function
ERROR_LABEL:
INDUSTRIES_DEBT_FUNC = False
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : INDUSTRIES_DIVIDENDS_FUNC
'DESCRIPTION   : DIVD TABLE
'LIBRARY       : FUNDAMENTAL
'GROUP         : INDUSTRIES
'ID            : 010

'LAST UPDATE   : 14 / 07 / 2009

'AUTHORS       : RAFAEL NICOLAS FERMIN COTA & CHRISTOPHER GILPIN
'REFERENCES    : http://pages.stern.nyu.edu/~adamodar/
'************************************************************************************
'************************************************************************************

Private Function INDUSTRIES_DIVIDENDS_FUNC(ByRef SRC_RNG As Excel.Range, _
ByRef DST_RNG As Excel.Range)

Dim i As Long
Dim SRC_WSHEET_NAME As String

Dim INDUSTRY As String
Dim NO_FIRMS As String
Dim FIRM_VALUE As String
Dim TOTAL_DEBT As String
Dim NET_INCOME As String
Dim DIVD_YIELD As String
Dim INSIDERS As String
Dim ST_DEV_3YR As String
Dim DIVD As String
Dim INSTITUT_INVEST As String
Dim Market_Cap As String
Dim BV_EQUITY As String

On Error GoTo ERROR_LABEL

INDUSTRIES_DIVIDENDS_FUNC = False
SRC_WSHEET_NAME = SRC_RNG.Worksheet.name & "!"

Range(DST_RNG.Cells(1, 1), _
DST_RNG.Cells(1, 1 + 8)).value = _
Array("Industries/Companies", "No. of Firms/Exchange", "Dividend Yield", _
"Dividend Payout", "Market Cap", "ROE", "Insider Holdings", _
"Institutional Holdings", "Std Dev in Stock Prices")

With DST_RNG
    For i = 1 To PUB_NSIZE
        INDUSTRY = SRC_RNG.Offset(i, P0).Address
        NO_FIRMS = SRC_RNG.Offset(i, p1).Address
        
        FIRM_VALUE = SRC_RNG.Offset(i, P13).Address
        TOTAL_DEBT = SRC_RNG.Offset(i, P12).Address
        
        DIVD_YIELD = SRC_RNG.Offset(i, P7).Address
        NET_INCOME = SRC_RNG.Offset(i, P26).Address
        
        DIVD = SRC_RNG.Offset(i, p2).Address
        Market_Cap = SRC_RNG.Offset(i, P11).Address
        
        BV_EQUITY = SRC_RNG.Offset(i, P34).Address
        INSIDERS = SRC_RNG.Offset(i, P9).Address
        INSTITUT_INVEST = SRC_RNG.Offset(i, P10).Address
        ST_DEV_3YR = SRC_RNG.Offset(i, P40).Address
        
        .Offset(i, 0).formula = "=" & SRC_WSHEET_NAME & INDUSTRY 'INDUSTRY
        .Offset(i, 1).formula = "=" & SRC_WSHEET_NAME & NO_FIRMS 'NO_FIRMS
        
        .Offset(i, 2).formula = "=" & SRC_WSHEET_NAME & DIVD_YIELD 'DIVD YIELD
        .Offset(i, 3).formula = "=IF(" & SRC_WSHEET_NAME & NET_INCOME & ">0," & _
        SRC_WSHEET_NAME & DIVD & "/" & SRC_WSHEET_NAME & NET_INCOME & ", ""NA"")" 'DIVD. Payout
        .Offset(i, 4).formula = "=" & SRC_WSHEET_NAME & Market_Cap 'MARKET CAPITALIZATION
        
        .Offset(i, 5).formula = "=" & SRC_WSHEET_NAME & NET_INCOME & " / " & _
        SRC_WSHEET_NAME & BV_EQUITY 'ROE
        
        .Offset(i, 6).formula = "=" & SRC_WSHEET_NAME & INSIDERS
        .Offset(i, 7).formula = "=" & SRC_WSHEET_NAME & INSTITUT_INVEST
        .Offset(i, 8).formula = "=" & SRC_WSHEET_NAME & ST_DEV_3YR
    Next i
End With

INDUSTRIES_DIVIDENDS_FUNC = True

Exit Function
ERROR_LABEL:
INDUSTRIES_DIVIDENDS_FUNC = False
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : INDUSTRIES_EBITDA_FUNC
'DESCRIPTION   : EBITDA TABLE
'LIBRARY       : FUNDAMENTAL
'GROUP         : INDUSTRIES
'ID            : 011

'LAST UPDATE   : 14 / 07 / 2009

'AUTHORS       : RAFAEL NICOLAS FERMIN COTA & CHRISTOPHER GILPIN
'REFERENCES    : http://pages.stern.nyu.edu/~adamodar/
'************************************************************************************
'************************************************************************************

Private Function INDUSTRIES_EBITDA_FUNC(ByRef SRC_RNG As Excel.Range, _
ByRef DST_RNG As Excel.Range)

Dim i As Long
Dim SRC_WSHEET_NAME As String

Dim INDUSTRY As String
Dim NO_FIRMS As String
Dim EBIT As String
Dim EBITDA As String
Dim EBIT_ATAX As String
Dim EV As String

On Error GoTo ERROR_LABEL

INDUSTRIES_EBITDA_FUNC = False
SRC_WSHEET_NAME = SRC_RNG.Worksheet.name & "!"

Range(DST_RNG.Cells(1, 1), _
DST_RNG.Cells(1, 1 + 4)).value = _
Array("Industries/Companies", "No. of Firms/Exchange", "Value/EBITDA", _
"Value/EBIT", "Value/EBIT(1-t)")

With DST_RNG
    For i = 1 To PUB_NSIZE
        INDUSTRY = SRC_RNG.Offset(i, P0).Address
        NO_FIRMS = SRC_RNG.Offset(i, p1).Address
        EBIT = SRC_RNG.Offset(i, P21).Address
        EBITDA = SRC_RNG.Offset(i, P22).Address
        EBIT_ATAX = SRC_RNG.Offset(i, P23).Address
        EV = SRC_RNG.Offset(i, P14).Address
        
        .Offset(i, 0).formula = "=" & SRC_WSHEET_NAME & INDUSTRY
        .Offset(i, 1).formula = "=" & SRC_WSHEET_NAME & NO_FIRMS
        .Offset(i, 2).formula = "=IF(" & SRC_WSHEET_NAME & EBITDA & ">0," & _
        SRC_WSHEET_NAME & EV & "/" & SRC_WSHEET_NAME & EBITDA & ",""NA"")"
        
        .Offset(i, 3).formula = "=IF(" & SRC_WSHEET_NAME & EBIT & ">0," & _
        SRC_WSHEET_NAME & EV & "/" & SRC_WSHEET_NAME & EBIT & ",""NA"")"
        .Offset(i, 4).formula = "=IF(" & SRC_WSHEET_NAME & EBIT_ATAX & ">0," & _
        SRC_WSHEET_NAME & EV & "/" & SRC_WSHEET_NAME & EBIT_ATAX & ",""NA"")"
    Next i
End With

INDUSTRIES_EBITDA_FUNC = True

Exit Function
ERROR_LABEL:
INDUSTRIES_EBITDA_FUNC = False
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : INDUSTRIES_EVA_FUNC
'DESCRIPTION   : EVA TABLE
'LIBRARY       : FUNDAMENTAL
'GROUP         : INDUSTRIES
'ID            : 012

'LAST UPDATE   : 14 / 07 / 2009

'AUTHORS       : RAFAEL NICOLAS FERMIN COTA & CHRISTOPHER GILPIN
'REFERENCES    : http://pages.stern.nyu.edu/~adamodar/
'************************************************************************************
'************************************************************************************

Function INDUSTRIES_EVA_FUNC(ByRef SRC_RNG As Excel.Range, _
ByRef DST_RNG As Excel.Range)

Dim i As Long
Dim SRC_WSHEET_NAME As String

Dim INDUSTRY As String
Dim NO_FIRMS As String
Dim SALES As String
Dim BETA_5YR As String
Dim BV_EQUITY As String
Dim NET_INCOME As String
Dim EBIT_ATAX As String
Dim VL_BETA As String
Dim INV_CAPITAL As String
Dim ST_DEV_3YR As String
Dim TAX_RATE As String
Dim TOTAL_DEBT As String
Dim FIRM_VALUE As String
Dim Market_Cap As String

On Error GoTo ERROR_LABEL

INDUSTRIES_EVA_FUNC = False
SRC_WSHEET_NAME = SRC_RNG.Worksheet.name & "!"

Range(DST_RNG.Cells(1, 1), _
DST_RNG.Cells(1, 1 + 18)).value = _
Array("Industries/Companies", "No. of Firms/Exchange", "Beta", "ROE", _
"Cost of Equity", "(ROE - COE)", "BV of Equity", "Equity EVA", "ROC", _
"Cost of Capital", "(ROC - WACC)", "BV of Capital", "EVA", "E/(D+E)", _
"Std Dev in Stock", "Cost of Debt", "Tax Rate", "After-tax Cost of Debt", "D/(D+E)")

With DST_RNG
    For i = 1 To PUB_NSIZE
        INDUSTRY = SRC_RNG.Offset(i, P0).Address
        NO_FIRMS = SRC_RNG.Offset(i, p1).Address
        VL_BETA = SRC_RNG.Offset(i, P3).Address
        BETA_5YR = SRC_RNG.Offset(i, P38).Address
        BV_EQUITY = SRC_RNG.Offset(i, P34).Address
        NET_INCOME = SRC_RNG.Offset(i, P26).Address
        SALES = SRC_RNG.Offset(i, P19).Address
        EBIT_ATAX = SRC_RNG.Offset(i, P23).Address
        INV_CAPITAL = SRC_RNG.Offset(i, P15).Address
        ST_DEV_3YR = SRC_RNG.Offset(i, P40).Address
        TAX_RATE = SRC_RNG.Offset(i, P8).Address
        TOTAL_DEBT = SRC_RNG.Offset(i, P12).Address
        FIRM_VALUE = SRC_RNG.Offset(i, P13).Address
        Market_Cap = SRC_RNG.Offset(i, P11).Address
        
        
        .Offset(i, 0).formula = "=" & SRC_WSHEET_NAME & INDUSTRY 'INDUSTRIES / COMPANIES
        .Offset(i, 1).formula = "=" & SRC_WSHEET_NAME & NO_FIRMS 'No. of Firms
        .Offset(i, 2).formula = "=MAX(" & SRC_WSHEET_NAME & VL_BETA & " , " & _
        SRC_WSHEET_NAME & BETA_5YR & ")" 'Beta
        .Offset(i, 3).formula = "=" & SRC_WSHEET_NAME & NET_INCOME & "/" & _
        SRC_WSHEET_NAME & BV_EQUITY 'ROE
        .Offset(i, 4).formula = "= LT_Bonds + " & .Offset(i, 4).Offset(0, -2).Address & _
        " * RISK_PREMIUM " 'Cost of Equity
        
        .Offset(i, 5).formula = "=" & .Offset(i, 5).Offset(0, -2).Address & "-" & _
        .Offset(i, 5).Offset(0, -1).Address  '(ROE - COE)
        .Offset(i, 6).formula = "=" & SRC_WSHEET_NAME & BV_EQUITY 'BV of Equity
        .Offset(i, 7).formula = "=" & .Offset(i, 7).Offset(0, -1).Address & " * " & _
        .Offset(i, 7).Offset(0, -2).Address 'equity EVA
        
        .Offset(i, 8).formula = "=IF(IF(" & SRC_WSHEET_NAME & SALES & ">0," & _
        SRC_WSHEET_NAME & EBIT_ATAX & "/" & SRC_WSHEET_NAME & SALES _
        & ",""NA"")=""NA"",""NA"",IF(" & SRC_WSHEET_NAME & SALES & ">0," & _
        EBIT_ATAX & "/" & SRC_WSHEET_NAME _
        & SALES & ",""NA"")* IF(" & SRC_WSHEET_NAME & SALES & ">0," & _
        SRC_WSHEET_NAME & SALES & "/" & _
        SRC_WSHEET_NAME & INV_CAPITAL & ",""NA""))"     '
        
        .Offset(i, 9).formula = "=" & .Offset(i, 9).Offset(0, -5).Address & _
        "*(1-" & .Offset(i, 9).Offset(0, 9).Address & ")+" & _
        .Offset(i, 9).Offset(0, 8).Address & "*" & .Offset(i, 9).Offset(0, 9).Address _
        ' Cost of Capital
        
        .Offset(i, 10).formula = "=IF(" & .Offset(i, 10).Offset(0, -2).Address & _
        "= ""NA"", ""NA""," & .Offset(i, 10).Offset(0, -2).Address & " - " & _
        .Offset(i, 10).Offset(0, -1).Address & ")"  '(ROC - WACC)"
        
        .Offset(i, 11).formula = "=" & SRC_WSHEET_NAME & INV_CAPITAL 'BV of Capital
        .Offset(i, 12).formula = "=IF(" & .Offset(i, 12).Offset(0, -2).Address & _
        "= ""NA"", ""NA""," & .Offset(i, 12).Offset(0, -2).Address & "*" & _
        .Offset(i, 12).Offset(0, -1).Address & ")" 'EVA"
        
        
        .Offset(i, 13).formula = "=1 -" & .Offset(i, 13).Offset(0, 5).Address 'E/(D+E)
        .Offset(i, 14).formula = "=" & SRC_WSHEET_NAME & ST_DEV_3YR 'Std Dev in Stock
        .Offset(i, 15).formula = "= LT_Bonds" & "+ VLOOKUP(" & _
        .Offset(i, 15).Offset(0, -1).Address & ", Debt_Table, 3)" 'Cost of Debt
        
        .Offset(i, 16).formula = "=" & SRC_WSHEET_NAME & TAX_RATE 'Tax Rate
        .Offset(i, 17).formula = "=" & .Offset(i, 17).Offset(0, -2).Address & _
        "*(1-" & .Offset(i, 17).Offset(0, -1).Address & ")" 'After-tax Cost of Debt"
        .Offset(i, 18).formula = "=" & SRC_WSHEET_NAME & TOTAL_DEBT & "/" & _
        SRC_WSHEET_NAME & FIRM_VALUE  ' D/(D+E)
    Next i
End With

INDUSTRIES_EVA_FUNC = True

Exit Function
ERROR_LABEL:
INDUSTRIES_EVA_FUNC = False
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : INDUSTRIES_TAX_FUNC
'DESCRIPTION   : TAX TABLE
'LIBRARY       : FUNDAMENTAL
'GROUP         : INDUSTRIES
'ID            : 013

'LAST UPDATE   : 14 / 07 / 2009

'AUTHORS       : RAFAEL NICOLAS FERMIN COTA & CHRISTOPHER GILPIN
'REFERENCES    : http://pages.stern.nyu.edu/~adamodar/
'************************************************************************************
'************************************************************************************

Private Function INDUSTRIES_TAX_FUNC(ByRef SRC_RNG As Excel.Range, _
ByRef DST_RNG As Excel.Range)

Dim i As Long
Dim SRC_WSHEET_NAME As String

Dim INDUSTRY As String
Dim NO_FIRMS As String
Dim TAX_RATE As String

On Error GoTo ERROR_LABEL

INDUSTRIES_TAX_FUNC = False
SRC_WSHEET_NAME = SRC_RNG.Worksheet.name & "!"

Range(DST_RNG.Cells(1, 1), _
DST_RNG.Cells(1, 1 + 2)).value = _
Array("Industries/Companies", "No. of Firms/Exchange", "Effect. Tax Rate")

With DST_RNG
    For i = 1 To PUB_NSIZE
        INDUSTRY = SRC_RNG.Offset(i, P0).Address
        NO_FIRMS = SRC_RNG.Offset(i, p1).Address
        TAX_RATE = SRC_RNG.Offset(i, P8).Address
        
        .Offset(i, 0).formula = "=" & SRC_WSHEET_NAME & INDUSTRY
        .Offset(i, 1).formula = "=" & SRC_WSHEET_NAME & NO_FIRMS
        .Offset(i, 2).formula = "=" & SRC_WSHEET_NAME & TAX_RATE
    Next i
End With

INDUSTRIES_TAX_FUNC = True

Exit Function
ERROR_LABEL:
INDUSTRIES_TAX_FUNC = False
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : INDUSTRIES_FCF_FUNC
'DESCRIPTION   : FCF TABLE
'LIBRARY       : FUNDAMENTAL
'GROUP         : INDUSTRIES
'ID            : 014

'LAST UPDATE   : 14 / 07 / 2009

'AUTHORS       : RAFAEL NICOLAS FERMIN COTA & CHRISTOPHER GILPIN
'REFERENCES    : http://pages.stern.nyu.edu/~adamodar/
'************************************************************************************
'************************************************************************************

Private Function INDUSTRIES_FCF_FUNC(ByRef SRC_RNG As Excel.Range, _
ByRef DST_RNG As Excel.Range)

Dim i As Long
Dim SRC_WSHEET_NAME As String

Dim INDUSTRY As String
Dim NO_FIRMS As String
Dim EBIT_ATAX As String
Dim REINVEST As String
Dim CHG_NON_CASH_WC As String

On Error GoTo ERROR_LABEL

INDUSTRIES_FCF_FUNC = False
SRC_WSHEET_NAME = SRC_RNG.Worksheet.name & "!"

Range(DST_RNG.Cells(1, 1), _
DST_RNG.Cells(1, 1 + 3)).value = _
Array("Industries/Companies", "No. of Firms/Exchange", "EBIT(1-t)", "FCFF")

With DST_RNG
    For i = 1 To PUB_NSIZE
    
        INDUSTRY = SRC_RNG.Offset(i, P0).Address
        NO_FIRMS = SRC_RNG.Offset(i, p1).Address
        EBIT_ATAX = SRC_RNG.Offset(i, P23).Address
        REINVEST = SRC_RNG.Offset(i, P18).Address
        CHG_NON_CASH_WC = SRC_RNG.Offset(i, P17).Address
        
        
        .Offset(i, 0).formula = "=" & SRC_WSHEET_NAME & INDUSTRY
        .Offset(i, 1).formula = "=" & SRC_WSHEET_NAME & NO_FIRMS
        .Offset(i, 2).formula = "=" & SRC_WSHEET_NAME & EBIT_ATAX
        .Offset(i, 3).formula = "=" & .Offset(i, 3).Offset(0, -1).Address & "-" _
        & SRC_WSHEET_NAME & REINVEST & "-" & SRC_WSHEET_NAME & CHG_NON_CASH_WC
    
    Next i
End With

INDUSTRIES_FCF_FUNC = True

Exit Function
ERROR_LABEL:
INDUSTRIES_FCF_FUNC = False
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : INDUSTRIES_FCFE_FUNC
'DESCRIPTION   : FCFE TABLE
'LIBRARY       : FUNDAMENTAL
'GROUP         : INDUSTRIES
'ID            : 015

'LAST UPDATE   : 14 / 07 / 2009
'AUTHORS       : RAFAEL NICOLAS FERMIN COTA & CHRISTOPHER GILPIN
'REFERENCES    : http://pages.stern.nyu.edu/~adamodar/
'************************************************************************************
'************************************************************************************

Private Function INDUSTRIES_FCFE_FUNC(ByRef SRC_RNG As Excel.Range, _
ByRef DST_RNG As Excel.Range)

Dim i As Long
Dim SRC_WSHEET_NAME As String

Dim INDUSTRY As String
Dim NO_FIRMS As String
Dim DIVD As String
Dim NET_INCOME As String
Dim CAPEX As String
Dim DEPRECIATION As String
Dim CHG_NON_CASH_WC As String
Dim INV_CAPITAL As String
Dim TOTAL_DEBT As String
Dim cash As String

On Error GoTo ERROR_LABEL

INDUSTRIES_FCFE_FUNC = False
SRC_WSHEET_NAME = SRC_RNG.Worksheet.name & "!"

Range(DST_RNG.Cells(1, 1), _
DST_RNG.Cells(1, 1 + 6)).value = _
Array("Industries/Companies", "No. of Firms/Exchange", "Dividends", "Net Income", _
"FCFE", "Payout", "Dividends/FCFE")

With DST_RNG

    For i = 1 To PUB_NSIZE
        
        INDUSTRY = SRC_RNG.Offset(i, P0).Address
        NO_FIRMS = SRC_RNG.Offset(i, p1).Address
        DIVD = SRC_RNG.Offset(i, p2).Address
        NET_INCOME = SRC_RNG.Offset(i, P26).Address
        CHG_NON_CASH_WC = SRC_RNG.Offset(i, P17).Address
        CAPEX = SRC_RNG.Offset(i, p25).Address
        DEPRECIATION = SRC_RNG.Offset(i, P24).Address
        INV_CAPITAL = SRC_RNG.Offset(i, P15).Address
        TOTAL_DEBT = SRC_RNG.Offset(i, P12).Address
        cash = SRC_RNG.Offset(i, P28).Address
        
        
        .Offset(i, 0).formula = "=" & SRC_WSHEET_NAME & INDUSTRY
        .Offset(i, 1).formula = "=" & SRC_WSHEET_NAME & NO_FIRMS
        .Offset(i, 2).formula = "=" & SRC_WSHEET_NAME & DIVD
        .Offset(i, 3).formula = "=" & SRC_WSHEET_NAME & NET_INCOME
        
        .Offset(i, 4).formula = "=" & SRC_WSHEET_NAME & NET_INCOME & "-(" & _
        SRC_WSHEET_NAME & CAPEX & "-" & SRC_WSHEET_NAME & DEPRECIATION & "-" & _
        SRC_WSHEET_NAME & CHG_NON_CASH_WC & ")*( 1-(" & SRC_WSHEET_NAME & TOTAL_DEBT & _
        "/(" & SRC_WSHEET_NAME & INV_CAPITAL & "+" & SRC_WSHEET_NAME & cash & ")))"
        
        .Offset(i, 5).formula = "=" & .Offset(i, 5).Offset(0, -3).Address & _
        "/" & .Offset(i, 5).Offset(0, -2).Address
        
        .Offset(i, 6).formula = "=IF(" & .Offset(i, 6).Offset(0, -2).Address & ">0," & _
        .Offset(i, 6).Offset(0, -4).Address & "/" & .Offset(i, 6).Offset(0, -2).Address & _
        ", ""NA"")"
            
    Next i
End With

INDUSTRIES_FCFE_FUNC = True

Exit Function
ERROR_LABEL:
INDUSTRIES_FCFE_FUNC = False
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : INDUSTRIES_GROWTH_FUNC
'DESCRIPTION   : GROWTH TABLE
'LIBRARY       : FUNDAMENTAL
'GROUP         : INDUSTRIES
'ID            : 016

'LAST UPDATE   : 14 / 07 / 2009

'AUTHORS       : RAFAEL NICOLAS FERMIN COTA & CHRISTOPHER GILPIN
'REFERENCES    : http://pages.stern.nyu.edu/~adamodar/
'************************************************************************************
'************************************************************************************

Private Function INDUSTRIES_GROWTH_FUNC(ByRef SRC_RNG As Excel.Range, _
ByRef DST_RNG As Excel.Range)

Dim i As Long
Dim SRC_WSHEET_NAME As String

Dim INDUSTRY As String
Dim NO_FIRMS As String
Dim NET_INCOME As String
Dim BV_EQUITY As String
Dim DIVD As String

On Error GoTo ERROR_LABEL

INDUSTRIES_GROWTH_FUNC = False
SRC_WSHEET_NAME = SRC_RNG.Worksheet.name & "!"

Range(DST_RNG.Cells(1, 1), _
DST_RNG.Cells(1, 1 + 4)).value = _
Array("Industries/Companies", "No. of Firms/Exchange", "ROE", "Retention Ratio", _
"Fundamental Growth")

With DST_RNG

    For i = 1 To PUB_NSIZE
    
        INDUSTRY = SRC_RNG.Offset(i, P0).Address
        NO_FIRMS = SRC_RNG.Offset(i, p1).Address
        BV_EQUITY = SRC_RNG.Offset(i, P34).Address
        NET_INCOME = SRC_RNG.Offset(i, P26).Address
        DIVD = SRC_RNG.Offset(i, p2).Address
        
        .Offset(i, 0).formula = "=" & SRC_WSHEET_NAME & INDUSTRY
        .Offset(i, 1).formula = "=" & SRC_WSHEET_NAME & NO_FIRMS
        .Offset(i, 2).formula = "=" & SRC_WSHEET_NAME & NET_INCOME & "/" & _
        SRC_WSHEET_NAME & BV_EQUITY
        
        .Offset(i, 3).formula = "=IF(" & SRC_WSHEET_NAME & NET_INCOME & _
        "> 0,1-" & SRC_WSHEET_NAME & DIVD & "/" & SRC_WSHEET_NAME & _
        NET_INCOME & ",""NA"")"
        
        .Offset(i, 4).formula = "=IF(" & .Offset(i, 4).Offset(0, -1).Address & _
        "= ""NA"", ""NA""," & .Offset(i, 4).Offset(0, -2).Address & "*" & _
        .Offset(i, 4).Offset(0, -1).Address & ")"
    
    Next i
End With

INDUSTRIES_GROWTH_FUNC = True

Exit Function
ERROR_LABEL:
INDUSTRIES_GROWTH_FUNC = False
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : INDUSTRIES_HISTORICAL_GROWTH_FUNC
'DESCRIPTION   : HIST.GROWTH TABLE
'LIBRARY       : FUNDAMENTAL
'GROUP         : INDUSTRIES
'ID            : 017

'LAST UPDATE   : 14 / 07 / 2009

'AUTHORS       : RAFAEL NICOLAS FERMIN COTA & CHRISTOPHER GILPIN
'REFERENCES    : http://pages.stern.nyu.edu/~adamodar/
'************************************************************************************
'************************************************************************************

Private Function INDUSTRIES_HISTORICAL_GROWTH_FUNC(ByRef SRC_RNG As Excel.Range, _
ByRef DST_RNG As Excel.Range)

Dim i As Long

Dim SRC_WSHEET_NAME As String

Dim INDUSTRY As String
Dim NO_FIRMS As String
Dim EPS_GROWTH_5YR As String
Dim SALES_GROWTH_5YR As String
Dim DIVD_GROWTH_5YR As String

On Error GoTo ERROR_LABEL

INDUSTRIES_HISTORICAL_GROWTH_FUNC = False
SRC_WSHEET_NAME = SRC_RNG.Worksheet.name & "!"

Range(DST_RNG.Cells(1, 1), _
DST_RNG.Cells(1, 1 + 4)).value = _
Array("Industries/Companies", "No. of Firms/Exchange", "Growth in EPS", _
"Growth in Sales", "Growth in Dividends")

With DST_RNG

    For i = 1 To PUB_NSIZE
    
        INDUSTRY = SRC_RNG.Offset(i, P0).Address
        NO_FIRMS = SRC_RNG.Offset(i, p1).Address
        EPS_GROWTH_5YR = SRC_RNG.Offset(i, P36).Address
        SALES_GROWTH_5YR = SRC_RNG.Offset(i, P35).Address
        DIVD_GROWTH_5YR = SRC_RNG.Offset(i, P42).Address
        
        .Offset(i, 0).formula = "=" & SRC_WSHEET_NAME & INDUSTRY
        .Offset(i, 1).formula = "=" & SRC_WSHEET_NAME & NO_FIRMS
        .Offset(i, 2).formula = "=" & SRC_WSHEET_NAME & EPS_GROWTH_5YR
        .Offset(i, 3).formula = "=" & SRC_WSHEET_NAME & SALES_GROWTH_5YR
        .Offset(i, 4).formula = "=IF(" & SRC_WSHEET_NAME & DIVD_GROWTH_5YR & _
        "> 0," & SRC_WSHEET_NAME & DIVD_GROWTH_5YR & ",""NA"")"
    
    Next i
End With

INDUSTRIES_HISTORICAL_GROWTH_FUNC = True

Exit Function
ERROR_LABEL:
INDUSTRIES_HISTORICAL_GROWTH_FUNC = False
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : INDUSTRIES_HOLDINGS_FUNC
'DESCRIPTION   : HOLDINGS TABLE
'LIBRARY       : FUNDAMENTAL
'GROUP         : INDUSTRIES
'ID            : 018

'LAST UPDATE   : 14 / 07 / 2009

'AUTHORS       : RAFAEL NICOLAS FERMIN COTA & CHRISTOPHER GILPIN
'REFERENCES    : http://pages.stern.nyu.edu/~adamodar/
'************************************************************************************
'************************************************************************************

Private Function INDUSTRIES_HOLDINGS_FUNC(ByRef SRC_RNG As Excel.Range, _
ByRef DST_RNG As Excel.Range)

Dim i As Long
Dim SRC_WSHEET_NAME As String

Dim INDUSTRY As String
Dim NO_FIRMS As String
Dim INSIDERS As String
Dim INSTITUT_INVEST As String

On Error GoTo ERROR_LABEL

INDUSTRIES_HOLDINGS_FUNC = False
SRC_WSHEET_NAME = SRC_RNG.Worksheet.name & "!"

Range(DST_RNG.Cells(1, 1), _
DST_RNG.Cells(1, 1 + 3)).value = _
Array("Industries/Companies", "No. of Firms/Exchange", "Insider Holdings", _
"Institutional Holdings")

With DST_RNG

    For i = 1 To PUB_NSIZE
    
        INDUSTRY = SRC_RNG.Offset(i, P0).Address
        NO_FIRMS = SRC_RNG.Offset(i, p1).Address
        INSIDERS = SRC_RNG.Offset(i, P9).Address
        INSTITUT_INVEST = SRC_RNG.Offset(i, P10).Address
        
        .Offset(i, 0).formula = "=" & SRC_WSHEET_NAME & INDUSTRY
        .Offset(i, 1).formula = "=" & SRC_WSHEET_NAME & NO_FIRMS
        .Offset(i, 2).formula = "=" & SRC_WSHEET_NAME & INSIDERS
        .Offset(i, 3).formula = "=" & SRC_WSHEET_NAME & INSTITUT_INVEST
    
    Next i
End With

INDUSTRIES_HOLDINGS_FUNC = True

Exit Function
ERROR_LABEL:
INDUSTRIES_HOLDINGS_FUNC = False
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : INDUSTRIES_INVENTORIES_FUNC
'DESCRIPTION   : INVENTORY TABLE
'LIBRARY       : FUNDAMENTAL
'GROUP         : INDUSTRIES
'ID            : 019

'LAST UPDATE   : 14 / 07 / 2009

'AUTHORS       : RAFAEL NICOLAS FERMIN COTA & CHRISTOPHER GILPIN
'REFERENCES    : http://pages.stern.nyu.edu/~adamodar/
'************************************************************************************
'************************************************************************************

Private Function INDUSTRIES_INVENTORIES_FUNC(ByRef SRC_RNG As Excel.Range, _
ByRef DST_RNG As Excel.Range)

Dim i As Long
Dim SRC_WSHEET_NAME As String

Dim INDUSTRY As String
Dim NO_FIRMS As String
Dim SALES As String
Dim INVENTORY As String
Dim EV As String

On Error GoTo ERROR_LABEL

INDUSTRIES_INVENTORIES_FUNC = True
SRC_WSHEET_NAME = SRC_RNG.Worksheet.name & "!"


Range(DST_RNG.Cells(1, 1), _
DST_RNG.Cells(1, 1 + 4)).value = _
Array("Industries/Companies", "No. of Firms/Exchange", "Inventory/Sales", _
"Inventory/Enterprise Value", "Number of Days Sales in Inventory")


With DST_RNG
    For i = 1 To PUB_NSIZE
    
        INDUSTRY = SRC_RNG.Offset(i, P0).Address
        NO_FIRMS = SRC_RNG.Offset(i, p1).Address
        SALES = SRC_RNG.Offset(i, P19).Address
        INVENTORY = SRC_RNG.Offset(i, P30).Address
        EV = SRC_RNG.Offset(i, P14).Address
        
        .Offset(i, 0).formula = "=" & SRC_WSHEET_NAME & INDUSTRY
        .Offset(i, 1).formula = "=" & SRC_WSHEET_NAME & NO_FIRMS
        .Offset(i, 2).formula = "=IF(" & SRC_WSHEET_NAME & SALES & "> 0," _
            & SRC_WSHEET_NAME & INVENTORY & "/" & SRC_WSHEET_NAME & SALES & ", ""NA"")"
        .Offset(i, 3).formula = "=" & SRC_WSHEET_NAME & INVENTORY & "/" _
            & SRC_WSHEET_NAME & EV
        .Offset(i, 4).formula = "=IF(" & SRC_WSHEET_NAME & SALES & "> 0," & _
            SRC_WSHEET_NAME & INVENTORY & "/(" & SRC_WSHEET_NAME & SALES & "/365),""NA"")"
    
    Next i
End With

INDUSTRIES_INVENTORIES_FUNC = True

Exit Function
ERROR_LABEL:
INDUSTRIES_INVENTORIES_FUNC = False
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : INDUSTRIES_MARGINS_FUNC
'DESCRIPTION   : MARGINS TABLE
'LIBRARY       : FUNDAMENTAL
'GROUP         : INDUSTRIES
'ID            : 020

'LAST UPDATE   : 14 / 07 / 2009

'AUTHORS       : RAFAEL NICOLAS FERMIN COTA & CHRISTOPHER GILPIN
'REFERENCES    : http://pages.stern.nyu.edu/~adamodar/
'************************************************************************************
'************************************************************************************

Private Function INDUSTRIES_MARGINS_FUNC(ByRef SRC_RNG As Excel.Range, _
ByRef DST_RNG As Excel.Range)

Dim i As Long
Dim SRC_WSHEET_NAME As String

Dim INDUSTRY As String
Dim NO_FIRMS As String
Dim SALES As String
Dim SGA_EXPENSES As String
Dim EBITDA As String
Dim EBIT As String
Dim EBIT_ATAX As String
Dim TRAILING_NET_INCOME As String

On Error GoTo ERROR_LABEL

INDUSTRIES_MARGINS_FUNC = False

SRC_WSHEET_NAME = SRC_RNG.Worksheet.name & "!"

Range(DST_RNG.Cells(1, 1), _
DST_RNG.Cells(1, 1 + 6)).value = _
Array("Industries/Companies", "No. of Firms/Exchange", "EBITDASG&A/Sales", _
"EBITDA/Sales", "EBIT/Sales", "After-tax Operating Margin (EBIT(1-t)", "Net Margin")

With DST_RNG

    For i = 1 To PUB_NSIZE
        
        INDUSTRY = SRC_RNG.Offset(i, P0).Address
        NO_FIRMS = SRC_RNG.Offset(i, p1).Address
        SALES = SRC_RNG.Offset(i, P19).Address
        SGA_EXPENSES = SRC_RNG.Offset(i, P20).Address
        EBIT = SRC_RNG.Offset(i, P21).Address
        EBITDA = SRC_RNG.Offset(i, P22).Address
        EBIT_ATAX = SRC_RNG.Offset(i, P23).Address
        TRAILING_NET_INCOME = SRC_RNG.Offset(i, P27).Address
        
        
        .Offset(i, 0).formula = "=" & SRC_WSHEET_NAME & INDUSTRY
        .Offset(i, 1).formula = "=" & SRC_WSHEET_NAME & NO_FIRMS
        .Offset(i, 2).formula = "=IF(" & SRC_WSHEET_NAME & SALES & "> 0,(" & _
            SRC_WSHEET_NAME & EBITDA & "+" & SRC_WSHEET_NAME & SGA_EXPENSES & ")/" & _
            SRC_WSHEET_NAME & SALES & ", ""NA"")"
        .Offset(i, 3).formula = "=IF(" & SRC_WSHEET_NAME & SALES & "> 0," & _
            SRC_WSHEET_NAME & EBITDA & "/" & SRC_WSHEET_NAME & SALES & ", ""NA"")"
        .Offset(i, 4).formula = "=IF(" & SRC_WSHEET_NAME & SALES & "> 0," & _
            SRC_WSHEET_NAME & EBIT & "/" & SRC_WSHEET_NAME & SALES & ", ""NA"")"
        .Offset(i, 5).formula = "=IF(" & SRC_WSHEET_NAME & SALES & "> 0," & _
            SRC_WSHEET_NAME & EBIT_ATAX & "/" & SRC_WSHEET_NAME & SALES & ", ""NA"")"
        .Offset(i, 6).formula = "=IF(" & SRC_WSHEET_NAME & SALES & "> 0," & _
            SRC_WSHEET_NAME & TRAILING_NET_INCOME & "/" & SRC_WSHEET_NAME & SALES & ", ""NA"")"
        
    Next i
End With

INDUSTRIES_MARGINS_FUNC = True

Exit Function
ERROR_LABEL:
INDUSTRIES_MARGINS_FUNC = False
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : INDUSTRIES_PE_FUNC
'DESCRIPTION   : PE TABLE
'LIBRARY       : FUNDAMENTAL
'GROUP         : INDUSTRIES
'ID            : 021

'LAST UPDATE   : 14 / 07 / 2009

'AUTHORS       : RAFAEL NICOLAS FERMIN COTA & CHRISTOPHER GILPIN
'REFERENCES    : http://pages.stern.nyu.edu/~adamodar/
'************************************************************************************
'************************************************************************************

Private Function INDUSTRIES_PE_FUNC(ByRef SRC_RNG As Excel.Range, _
ByRef DST_RNG As Excel.Range)

Dim i As Long
Dim SRC_WSHEET_NAME As String

Dim INDUSTRY As String
Dim NO_FIRMS As String
Dim NET_INCOME As String
Dim Market_Cap As String
Dim CURRENT_PE As String
Dim TRAILING_PE As String
Dim FORDWARD_PE As String
Dim CURRENT_EPS_GROWTH As String
Dim DIVD As String
Dim VL_BETA As String
Dim BETA_5YR As String

On Error GoTo ERROR_LABEL

INDUSTRIES_PE_FUNC = False
SRC_WSHEET_NAME = SRC_RNG.Worksheet.name & "!"


Range(DST_RNG.Cells(1, 1), _
DST_RNG.Cells(1, 1 + 8)).value = _
Array("Industries/Companies", "No. of Firms/Exchange", _
"Aggregate Market Cap/ Aggregate Net Income", _
"Price/Current EPS", "Price/Trailing EPS", _
"Price/Forward PE", "Expected Growth", _
"Payout", "Beta")


With DST_RNG

    For i = 1 To PUB_NSIZE

        INDUSTRY = SRC_RNG.Offset(i, P0).Address
        NO_FIRMS = SRC_RNG.Offset(i, p1).Address
        DIVD = SRC_RNG.Offset(i, p2).Address
        VL_BETA = SRC_RNG.Offset(i, P3).Address
        CURRENT_PE = SRC_RNG.Offset(i, P4).Address
        FORDWARD_PE = SRC_RNG.Offset(i, P5).Address
        TRAILING_PE = SRC_RNG.Offset(i, P6).Address
        Market_Cap = SRC_RNG.Offset(i, P11).Address
        NET_INCOME = SRC_RNG.Offset(i, P26).Address
        CURRENT_EPS_GROWTH = SRC_RNG.Offset(i, P41).Address
        BETA_5YR = SRC_RNG.Offset(i, P38).Address
        
        .Offset(i, 0).formula = "=" & SRC_WSHEET_NAME & INDUSTRY
        .Offset(i, 1).formula = "=" & SRC_WSHEET_NAME & NO_FIRMS
        .Offset(i, 2).formula = "=IF(" & SRC_WSHEET_NAME & NET_INCOME & "> 0," & _
            SRC_WSHEET_NAME & Market_Cap & "/" & SRC_WSHEET_NAME & NET_INCOME & ", ""NA"")"
        .Offset(i, 3).formula = "=" & SRC_WSHEET_NAME & CURRENT_PE
        .Offset(i, 4).formula = "=" & SRC_WSHEET_NAME & TRAILING_PE
        .Offset(i, 5).formula = "=" & SRC_WSHEET_NAME & FORDWARD_PE
        .Offset(i, 6).formula = "=" & SRC_WSHEET_NAME & CURRENT_EPS_GROWTH
        .Offset(i, 7).formula = "=IF(" & SRC_WSHEET_NAME & NET_INCOME & "> 0," & _
            SRC_WSHEET_NAME & DIVD & "/" & SRC_WSHEET_NAME & NET_INCOME & ", ""NA"")"
        .Offset(i, 8).formula = "=MAX(" & SRC_WSHEET_NAME & VL_BETA & ", " & _
            SRC_WSHEET_NAME & BETA_5YR & ")"
        
    Next i
End With

INDUSTRIES_PE_FUNC = True

Exit Function
ERROR_LABEL:
INDUSTRIES_PE_FUNC = False
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : INDUSTRIES_PB_FUNC
'DESCRIPTION   : Price to Book Value Table
'LIBRARY       : FUNDAMENTAL
'GROUP         : INDUSTRIES
'ID            : 022

'LAST UPDATE   : 14 / 07 / 2009

'AUTHORS       : RAFAEL NICOLAS FERMIN COTA & CHRISTOPHER GILPIN
'REFERENCES    : http://pages.stern.nyu.edu/~adamodar/
'************************************************************************************
'************************************************************************************

Private Function INDUSTRIES_PB_FUNC(ByRef SRC_RNG As Excel.Range, _
ByRef DST_RNG As Excel.Range)

Dim i As Long
Dim SRC_WSHEET_NAME As String

Dim INDUSTRY As String
Dim NO_FIRMS As String
Dim BV_EQUITY As String
Dim Market_Cap As String
Dim NET_INCOME As String
Dim CURRENT_EPS_GROWTH As String
Dim DIVD As String
Dim VL_BETA As String
Dim FIRM_VALUE As String
Dim BETA_5YR As String
Dim INV_CAPITAL As String
Dim SALES As String
Dim EBIT_ATAX As String

On Error GoTo ERROR_LABEL

INDUSTRIES_PB_FUNC = False
SRC_WSHEET_NAME = SRC_RNG.Worksheet.name & "!"

Range(DST_RNG.Cells(1, 1), _
DST_RNG.Cells(1, 1 + 8)).value = _
Array("Industries/Companies", "No. of Firms/Exchange", "Price/BV", "ROE", _
"Expected Growth in EPS", "Payout", "Beta", "Value/BV", "ROC")

With DST_RNG

    For i = 1 To PUB_NSIZE
    
        INDUSTRY = SRC_RNG.Offset(i, P0).Address
        NO_FIRMS = SRC_RNG.Offset(i, p1).Address
        BV_EQUITY = SRC_RNG.Offset(i, P34).Address
        Market_Cap = SRC_RNG.Offset(i, P11).Address
        NET_INCOME = SRC_RNG.Offset(i, P26).Address
        CURRENT_EPS_GROWTH = SRC_RNG.Offset(i, P41).Address
        DIVD = SRC_RNG.Offset(i, p2).Address
        VL_BETA = SRC_RNG.Offset(i, P3).Address
        BETA_5YR = SRC_RNG.Offset(i, P38).Address
        FIRM_VALUE = SRC_RNG.Offset(i, P13).Address
        INV_CAPITAL = SRC_RNG.Offset(i, P15).Address
        SALES = SRC_RNG.Offset(i, P19).Address
        EBIT_ATAX = SRC_RNG.Offset(i, P23).Address
        
        
        .Offset(i, 0).formula = "=" & SRC_WSHEET_NAME & INDUSTRY
        .Offset(i, 1).formula = "=" & SRC_WSHEET_NAME & NO_FIRMS
        .Offset(i, 2).formula = "=" & SRC_WSHEET_NAME & Market_Cap & "/" & _
            SRC_WSHEET_NAME & BV_EQUITY
        .Offset(i, 3).formula = "=" & SRC_WSHEET_NAME & NET_INCOME & "/" & _
            SRC_WSHEET_NAME & BV_EQUITY
        .Offset(i, 4).formula = "=" & SRC_WSHEET_NAME & CURRENT_EPS_GROWTH
        .Offset(i, 5).formula = "=IF(" & SRC_WSHEET_NAME & NET_INCOME & "> 0," & _
            SRC_WSHEET_NAME & DIVD & "/" & SRC_WSHEET_NAME & NET_INCOME & ", ""NA"")"
        .Offset(i, 6).formula = "=MAX(" & SRC_WSHEET_NAME & VL_BETA & ", " & _
            SRC_WSHEET_NAME & BETA_5YR & ")"
        .Offset(i, 7).formula = "=" & SRC_WSHEET_NAME & FIRM_VALUE & "/" & _
            SRC_WSHEET_NAME & INV_CAPITAL
        .Offset(i, 8).formula = "=If(if(" & SRC_WSHEET_NAME & SALES & ">0," & _
            SRC_WSHEET_NAME & EBIT_ATAX & "/" & SRC_WSHEET_NAME & SALES & _
            ", ""NA"")=""NA"",""NA"",IF(" & SRC_WSHEET_NAME & SALES & "> 0, " & _
            SRC_WSHEET_NAME & EBIT_ATAX & "/" & SRC_WSHEET_NAME & SALES & _
            ", ""NA"")*IF(" & SRC_WSHEET_NAME & SALES & ">0," & SRC_WSHEET_NAME & SALES & _
            "/" & SRC_WSHEET_NAME & INV_CAPITAL & ",""NA""))"
    
    Next i
End With

INDUSTRIES_PB_FUNC = True

Exit Function
ERROR_LABEL:
INDUSTRIES_PB_FUNC = False
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : INDUSTRIES_REGRESSION_FUNC
'DESCRIPTION   : FUNDAMENTAL REGRESSION TABLE
'LIBRARY       : FUNDAMENTAL
'GROUP         : INDUSTRIES
'ID            : 023

'LAST UPDATE   : 14 / 07 / 2009

'AUTHORS       : RAFAEL NICOLAS FERMIN COTA & CHRISTOPHER GILPIN
'REFERENCES    : http://pages.stern.nyu.edu/~adamodar/
'************************************************************************************
'************************************************************************************

Private Function INDUSTRIES_REGRESSION_FUNC(ByRef SRC_RNG As Excel.Range, _
ByRef DST_RNG As Excel.Range)

Dim i As Long
Dim SRC_WSHEET_NAME As String

Dim INDUSTRY As String
Dim NO_FIRMS As String
Dim RETURN_5YR As String
Dim BETA_5YR As String
Dim CORRELATION As String

On Error GoTo ERROR_LABEL

INDUSTRIES_REGRESSION_FUNC = False
SRC_WSHEET_NAME = SRC_RNG.Worksheet.name & "!"

Range(DST_RNG.Cells(1, 1), _
DST_RNG.Cells(1, 1 + 4)).value = _
Array("Industries/Companies", "No. of Firms/Exchange", "Jensen's Alpha", "Beta", "R-Squared")


With DST_RNG

    For i = 1 To PUB_NSIZE
    
        INDUSTRY = SRC_RNG.Offset(i, P0).Address
        NO_FIRMS = SRC_RNG.Offset(i, p1).Address
        RETURN_5YR = SRC_RNG.Offset(i, P37).Address
        BETA_5YR = SRC_RNG.Offset(i, P38).Address
        CORRELATION = SRC_RNG.Offset(i, P39).Address
        
        .Offset(i, 0).formula = "=" & SRC_WSHEET_NAME & INDUSTRY
        .Offset(i, 1).formula = "=" & SRC_WSHEET_NAME & NO_FIRMS
        .Offset(i, 2).formula = "=" & SRC_WSHEET_NAME & RETURN_5YR & "-" & _
        .Offset(i, 2).Offset(0, 1).Address & "*" & SRC_WSHEET_NAME & RETURN_5YR
        .Offset(i, 3).formula = "=" & SRC_WSHEET_NAME & BETA_5YR
        .Offset(i, 4).formula = "=" & SRC_WSHEET_NAME & CORRELATION & "^2"
            
    Next i
End With

INDUSTRIES_REGRESSION_FUNC = True

Exit Function
ERROR_LABEL:
INDUSTRIES_REGRESSION_FUNC = False
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : INDUSTRIES_PS_FUNC
'DESCRIPTION   : Price to Sale Table
'LIBRARY       : FUNDAMENTAL
'GROUP         : INDUSTRIES
'ID            : 024

'LAST UPDATE   : 14 / 07 / 2009

'AUTHORS       : RAFAEL NICOLAS FERMIN COTA & CHRISTOPHER GILPIN
'REFERENCES    : http://pages.stern.nyu.edu/~adamodar/
'************************************************************************************
'************************************************************************************

Private Function INDUSTRIES_PS_FUNC(ByRef SRC_RNG As Excel.Range, _
ByRef DST_RNG As Excel.Range)

Dim i As Long
Dim SRC_WSHEET_NAME As String

Dim INDUSTRY As String
Dim NO_FIRMS As String
Dim SALES As String
Dim Market_Cap As String
Dim TRAILING_NET_INCOME As String
Dim CURRENT_EPS_GROWTH As String
Dim EBIT_ATAX As String
Dim NET_INCOME As String
Dim DIVD As String
Dim VL_BETA As String
Dim BETA_5YR As String
Dim EV As String

On Error GoTo ERROR_LABEL

INDUSTRIES_PS_FUNC = False
SRC_WSHEET_NAME = SRC_RNG.Worksheet.name & "!"

Range(DST_RNG.Cells(1, 1), _
DST_RNG.Cells(1, 1 + 8)).value = _
Array("Industries/Companies", "No. of Firms/Exchange", "Price/Sales", "Net Margin", _
"Expected Growth", "Payout", "Beta", "Value/Sales", "After-tax Operating Margin")


With DST_RNG

    For i = 1 To PUB_NSIZE
    
        INDUSTRY = SRC_RNG.Offset(i, P0).Address
        NO_FIRMS = SRC_RNG.Offset(i, p1).Address
        Market_Cap = SRC_RNG.Offset(i, P11).Address
        SALES = SRC_RNG.Offset(i, P19).Address
        TRAILING_NET_INCOME = SRC_RNG.Offset(i, P27).Address
        CURRENT_EPS_GROWTH = SRC_RNG.Offset(i, P41).Address
        NET_INCOME = SRC_RNG.Offset(i, P26).Address
        DIVD = SRC_RNG.Offset(i, p2).Address
        VL_BETA = SRC_RNG.Offset(i, P3).Address
        BETA_5YR = SRC_RNG.Offset(i, P38).Address
        EV = SRC_RNG.Offset(i, P14).Address
        EBIT_ATAX = SRC_RNG.Offset(i, P23).Address
        
        
        .Offset(i, 0).formula = "=" & SRC_WSHEET_NAME & INDUSTRY
        .Offset(i, 1).formula = "=" & SRC_WSHEET_NAME & NO_FIRMS
        .Offset(i, 2).formula = "=IF(" & SRC_WSHEET_NAME & SALES & "> 0," & _
        SRC_WSHEET_NAME & Market_Cap & "/" & SRC_WSHEET_NAME & SALES & ", ""NA"")"
        
        .Offset(i, 3).formula = "=IF(" & SRC_WSHEET_NAME & SALES & "> 0," & _
        SRC_WSHEET_NAME & TRAILING_NET_INCOME & "/" & SRC_WSHEET_NAME & SALES & ", ""NA"")"
        
        .Offset(i, 4).formula = "=" & SRC_WSHEET_NAME & CURRENT_EPS_GROWTH
        .Offset(i, 5).formula = "=IF(" & SRC_WSHEET_NAME & NET_INCOME & "> 0," & _
        SRC_WSHEET_NAME & DIVD & "/" & SRC_WSHEET_NAME & NET_INCOME & ", ""NA"")"
        
        .Offset(i, 6).formula = "=MAX(" & SRC_WSHEET_NAME & VL_BETA & ", " & _
        SRC_WSHEET_NAME & BETA_5YR & ")"
        
        .Offset(i, 7).formula = "=IF(" & SRC_WSHEET_NAME & SALES & "> 0," & _
        SRC_WSHEET_NAME & EV & "/" & SRC_WSHEET_NAME & SALES & ", ""NA"")"
        
        .Offset(i, 8).formula = "=if(" & SRC_WSHEET_NAME & SALES & ">0," & _
        SRC_WSHEET_NAME & EBIT_ATAX & "/" & SRC_WSHEET_NAME & SALES & ", ""NA"")"
                
    Next i
End With

INDUSTRIES_PS_FUNC = True

Exit Function
ERROR_LABEL:
INDUSTRIES_PS_FUNC = False
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : INDUSTRIES_RETENTION_FUNC
'DESCRIPTION   : RETENTION TABLE
'LIBRARY       : FUNDAMENTAL
'GROUP         : INDUSTRIES
'ID            : 025

'LAST UPDATE   : 14 / 07 / 2009

'AUTHORS       : RAFAEL NICOLAS FERMIN COTA & CHRISTOPHER GILPIN
'REFERENCES    : http://pages.stern.nyu.edu/~adamodar/
'************************************************************************************
'************************************************************************************


Private Function INDUSTRIES_RETENTION_FUNC(ByRef SRC_RNG As Excel.Range, _
ByRef DST_RNG As Excel.Range)

Dim i As Long
Dim SRC_WSHEET_NAME As String

Dim INDUSTRY As String
Dim NO_FIRMS As String
Dim SALES As String
Dim RETURN_5YR As String
Dim BETA_5YR As String
Dim CORRELATION As String
Dim EBIT_ATAX As String
Dim INV_CAPITAL As String
Dim REINVEST As String

On Error GoTo ERROR_LABEL

INDUSTRIES_RETENTION_FUNC = False
SRC_WSHEET_NAME = SRC_RNG.Worksheet.name & "!"

Range(DST_RNG.Cells(1, 1), _
DST_RNG.Cells(1, 1 + 4)).value = _
Array("Industries/Companies", "No. of Firms/Exchange", "ROC", "Reinvestment Rate", _
"Expected Growth in EBIT")

With DST_RNG

    For i = 1 To PUB_NSIZE
    
        INDUSTRY = SRC_RNG.Offset(i, P0).Address
        NO_FIRMS = SRC_RNG.Offset(i, p1).Address
        SALES = SRC_RNG.Offset(i, P19).Address
        EBIT_ATAX = SRC_RNG.Offset(i, P23).Address
        INV_CAPITAL = SRC_RNG.Offset(i, P15).Address
        REINVEST = SRC_RNG.Offset(i, P18).Address
        
        
        RETURN_5YR = SRC_RNG.Offset(i, P37).Address
        BETA_5YR = SRC_RNG.Offset(i, P38).Address
        CORRELATION = SRC_RNG.Offset(i, P39).Address
        
        .Offset(i, 0).formula = "=" & SRC_WSHEET_NAME & INDUSTRY
        .Offset(i, 1).formula = "=" & SRC_WSHEET_NAME & NO_FIRMS
        .Offset(i, 2).formula = "=If(if(" & SRC_WSHEET_NAME & SALES & ">0," & _
        SRC_WSHEET_NAME & EBIT_ATAX & "/" & SRC_WSHEET_NAME & SALES & _
        ", ""NA"")=""NA"",""NA"",IF(" & SRC_WSHEET_NAME & SALES & _
        "> 0, " & SRC_WSHEET_NAME & EBIT_ATAX & "/" & SRC_WSHEET_NAME & SALES & _
        ", ""NA"")*IF(" & SRC_WSHEET_NAME & SALES & ">0," & SRC_WSHEET_NAME & SALES & _
        "/" & SRC_WSHEET_NAME & INV_CAPITAL & ",""NA""))"
        
        .Offset(i, 3).formula = "=If(" & SRC_WSHEET_NAME & EBIT_ATAX & ">0," & _
        SRC_WSHEET_NAME & REINVEST & "/" & SRC_WSHEET_NAME & EBIT_ATAX & ",""NA"")"
        .Offset(i, 4).formula = "=If(OR(" & .Offset(i, 4).Offset(0, -1).Address & _
        "=""NA""," & .Offset(i, 4).Offset(0, -2).Address & "=""NA""),""NA""," & _
        .Offset(i, 4).Offset(0, -2).Address & "*" & .Offset(i, 4).Offset(0, -1).Address & ")"
        
    
    Next i
End With

INDUSTRIES_RETENTION_FUNC = True

Exit Function
ERROR_LABEL:
INDUSTRIES_RETENTION_FUNC = False
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : INDUSTRIES_ROC_FUNC
'DESCRIPTION   : ROC TABLE
'LIBRARY       : FUNDAMENTAL
'GROUP         : INDUSTRIES
'ID            : 026

'LAST UPDATE   : 14 / 07 / 2009

'AUTHORS       : RAFAEL NICOLAS FERMIN COTA & CHRISTOPHER GILPIN
'REFERENCES    : http://pages.stern.nyu.edu/~adamodar/
'************************************************************************************
'************************************************************************************

Private Function INDUSTRIES_ROC_FUNC(ByRef SRC_RNG As Excel.Range, _
ByRef DST_RNG As Excel.Range)

Dim i As Long
Dim SRC_WSHEET_NAME As String

Dim INDUSTRY As String
Dim NO_FIRMS As String
Dim SALES As String
Dim EBIT_ATAX As String
Dim RETURN_5YR As String
Dim BETA_5YR As String
Dim CORRELATION As String
Dim INV_CAPITAL As String

On Error GoTo ERROR_LABEL

INDUSTRIES_ROC_FUNC = False
SRC_WSHEET_NAME = SRC_RNG.Worksheet.name & "!"

Range(DST_RNG.Cells(1, 1), _
DST_RNG.Cells(1, 1 + 4)).value = _
Array("Industries/Companies", "No. of Firms/Exchange", "After-tax Operating Margin", _
"Sales/Capital", "Return on Capital")

With DST_RNG

    For i = 1 To PUB_NSIZE
    
        INDUSTRY = SRC_RNG.Offset(i, P0).Address
        NO_FIRMS = SRC_RNG.Offset(i, p1).Address
        SALES = SRC_RNG.Offset(i, P19).Address
        EBIT_ATAX = SRC_RNG.Offset(i, P23).Address
        INV_CAPITAL = SRC_RNG.Offset(i, P15).Address
        
        
        RETURN_5YR = SRC_RNG.Offset(i, P37).Address
        BETA_5YR = SRC_RNG.Offset(i, P38).Address
        CORRELATION = SRC_RNG.Offset(i, P39).Address
        
        .Offset(i, 0).formula = "=" & SRC_WSHEET_NAME & INDUSTRY
        
        .Offset(i, 1).formula = "=" & SRC_WSHEET_NAME & NO_FIRMS
        
        .Offset(i, 2).formula = "=If(" & SRC_WSHEET_NAME & SALES & ">0," _
        & SRC_WSHEET_NAME & EBIT_ATAX & "/" & SRC_WSHEET_NAME & SALES & ", ""NA"")"
        .Offset(i, 3).formula = "=If(" & SRC_WSHEET_NAME & SALES & ">0," & _
        SRC_WSHEET_NAME & SALES & "/" & SRC_WSHEET_NAME & INV_CAPITAL & ", ""NA"")"
        
        .Offset(i, 4).formula = "=If(" & .Offset(i, 4).Offset(0, -2).Address & _
        "=""NA"",""NA""," & .Offset(i, 4).Offset(0, -2).Address & "*" & _
        .Offset(i, 4).Offset(0, -1).Address & ")"
    
    Next i
End With

INDUSTRIES_ROC_FUNC = True

Exit Function
ERROR_LABEL:
INDUSTRIES_ROC_FUNC = False
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : INDUSTRIES_ROE_FUNC
'DESCRIPTION   : ROE TABLE
'LIBRARY       : FUNDAMENTAL
'GROUP         : INDUSTRIES
'ID            : 027

'LAST UPDATE   : 14 / 07 / 2009

'AUTHORS       : RAFAEL NICOLAS FERMIN COTA & CHRISTOPHER GILPIN
'REFERENCES    : http://pages.stern.nyu.edu/~adamodar/
'************************************************************************************
'************************************************************************************

Private Function INDUSTRIES_ROE_FUNC(ByRef SRC_RNG As Excel.Range, _
ByRef DST_RNG As Excel.Range)

Dim i As Long
Dim SRC_WSHEET_NAME As String

Dim INDUSTRY As String
Dim NO_FIRMS As String
Dim SALES As String
Dim EBIT_ATAX As String
Dim INV_CAPITAL As String
Dim BV_EQUITY As String
Dim TOTAL_DEBT As String
Dim cash As String
Dim NET_INCOME As String

On Error GoTo ERROR_LABEL

INDUSTRIES_ROE_FUNC = False
SRC_WSHEET_NAME = SRC_RNG.Worksheet.name & "!"

Range(DST_RNG.Cells(1, 1), _
DST_RNG.Cells(1, 1 + 5)).value = _
Array("Industries/Companies", "No. of Firms/Exchange", "ROC", "Book D/E", "Non-cash ROE", "ROE")

With DST_RNG

    For i = 1 To PUB_NSIZE
        
        INDUSTRY = SRC_RNG.Offset(i, P0).Address
        NO_FIRMS = SRC_RNG.Offset(i, p1).Address
        SALES = SRC_RNG.Offset(i, P19).Address
        EBIT_ATAX = SRC_RNG.Offset(i, P23).Address
        INV_CAPITAL = SRC_RNG.Offset(i, P15).Address
        TOTAL_DEBT = SRC_RNG.Offset(i, P12).Address
        BV_EQUITY = SRC_RNG.Offset(i, P34).Address
        cash = SRC_RNG.Offset(i, P28).Address
        NET_INCOME = SRC_RNG.Offset(i, P26).Address
        
        .Offset(i, 0).formula = "=" & SRC_WSHEET_NAME & INDUSTRY
        
        .Offset(i, 1).formula = "=" & SRC_WSHEET_NAME & NO_FIRMS
        
        .Offset(i, 2).formula = "=If(If(" & SRC_WSHEET_NAME & SALES & _
        ">0," & SRC_WSHEET_NAME & EBIT_ATAX & "/" & SRC_WSHEET_NAME & SALES & _
        ", ""NA"")=""NA"",""NA"",IF(" & SRC_WSHEET_NAME & SALES & "> 0, " & _
        SRC_WSHEET_NAME & EBIT_ATAX & "/" & SRC_WSHEET_NAME & SALES & _
        ", ""NA"")*IF(" & SRC_WSHEET_NAME & SALES & ">0," & SRC_WSHEET_NAME & _
        SALES & "/" & SRC_WSHEET_NAME & INV_CAPITAL & ",""NA""))"
        
        .Offset(i, 3).formula = "=" & SRC_WSHEET_NAME & TOTAL_DEBT & "/" & _
        SRC_WSHEET_NAME & BV_EQUITY
        
        .Offset(i, 4).formula = "=" & SRC_WSHEET_NAME & NET_INCOME & "/(" & _
        SRC_WSHEET_NAME & BV_EQUITY & "-" & SRC_WSHEET_NAME & cash & ")"
        
        .Offset(i, 5).formula = "=" & SRC_WSHEET_NAME & NET_INCOME & "/" & _
        SRC_WSHEET_NAME & BV_EQUITY
        
    Next i
End With

INDUSTRIES_ROE_FUNC = True

Exit Function
ERROR_LABEL:
INDUSTRIES_ROE_FUNC = False
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : INDUSTRIES_VOLATILITY_FUNC
'DESCRIPTION   : STD. DEV. TABLE
'LIBRARY       : FUNDAMENTAL
'GROUP         : INDUSTRIES
'ID            : 028

'LAST UPDATE   : 14 / 07 / 2009

'AUTHORS       : RAFAEL NICOLAS FERMIN COTA & CHRISTOPHER GILPIN
'REFERENCES    : http://pages.stern.nyu.edu/~adamodar/
'************************************************************************************
'************************************************************************************

Private Function INDUSTRIES_VOLATILITY_FUNC(ByRef SRC_RNG As Excel.Range, _
ByRef DST_RNG As Excel.Range)

Dim i As Long
Dim SRC_WSHEET_NAME As String

Dim INDUSTRY As String
Dim NO_FIRMS As String
Dim ST_DEV_3YR As String
Dim FIRM_VALUE As String
Dim Market_Cap As String

On Error GoTo ERROR_LABEL

INDUSTRIES_VOLATILITY_FUNC = False
SRC_WSHEET_NAME = SRC_RNG.Worksheet.name & "!"

Range(DST_RNG.Cells(1, 1), _
DST_RNG.Cells(1, 1 + 5)).value = _
Array("Industries/Companies", "No. of Firms/Exchange", "Std Deviation in Equity", _
"Std Deviation in Firm Value", "E/(D+E)", "D/(D+E)")

With DST_RNG

    For i = 1 To PUB_NSIZE
        
        INDUSTRY = SRC_RNG.Offset(i, P0).Address
        NO_FIRMS = SRC_RNG.Offset(i, p1).Address
        ST_DEV_3YR = SRC_RNG.Offset(i, P40).Address
        FIRM_VALUE = SRC_RNG.Offset(i, P13).Address
        Market_Cap = SRC_RNG.Offset(i, P11).Address
        
        
        .Offset(i, 0).formula = "=" & SRC_WSHEET_NAME & INDUSTRY
        .Offset(i, 1).formula = "=" & SRC_WSHEET_NAME & NO_FIRMS
        .Offset(i, 2).formula = "=" & SRC_WSHEET_NAME & ST_DEV_3YR
        
        .Offset(i, 3).formula = "=(" & .Offset(i, 3).Offset(0, 1).Address & "^2*" & _
            .Offset(i, 3).Offset(0, -1).Address & "^2+" & .Offset(i, 3).Offset(0, 2).Address & _
            "^2*" & .Offset(i, 3).Offset(0, -1).Address & "^2/((1/Sigma))+ 2*" & _
            .Offset(i, 3).Offset(0, 1).Address & "*" & .Offset(i, 3).Offset(0, 2).Address & _
            "* Pearson *" & .Offset(i, 3).Offset(0, -1).Address & _
            "*(" & .Offset(i, 3).Offset(0, -1).Address & "/(1/Sigma)^2))^0.5"
        
        .Offset(i, 4).formula = "=" & SRC_WSHEET_NAME & Market_Cap & "/" & _
        SRC_WSHEET_NAME & FIRM_VALUE
        
        .Offset(i, 5).formula = "=1-" & .Offset(i, 5).Offset(0, -1).Address
    
    Next i
End With

INDUSTRIES_VOLATILITY_FUNC = True

Exit Function
ERROR_LABEL:
INDUSTRIES_VOLATILITY_FUNC = False
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : INDUSTRIES_VALUATION_FUNC
'DESCRIPTION   : VALUATION TABLE
'LIBRARY       : FUNDAMENTAL
'GROUP         : INDUSTRIES
'ID            : 029
'LAST UPDATE   : 14 / 07 / 2009
'AUTHORS       : RAFAEL NICOLAS FERMIN COTA & CHRISTOPHER GILPIN
'REFERENCES    : http://pages.stern.nyu.edu/~adamodar/
'************************************************************************************
'************************************************************************************

Private Function INDUSTRIES_VALUATION_FUNC(ByRef SRC_RNG As Excel.Range, _
ByRef DST_RNG As Excel.Range)

Dim i As Long
Dim SRC_WSHEET_NAME As String

Dim INDUSTRY As String
Dim NO_FIRMS As String
Dim VL_BETA As String
Dim TAX_RATE As String
Dim Market_Cap As String
Dim TOTAL_DEBT As String
Dim FIRM_VALUE As String
Dim EV As String
Dim INV_CAPITAL As String
Dim NON_CASH_WC As String
Dim REINVEST As String
Dim SALES As String
Dim EBIT As String
Dim EBIT_ATAX As String
Dim DEPRECIATION As String
Dim CAPEX As String
Dim NET_INCOME As String
Dim TRAILING_NET_INCOME As String
Dim DIVD As String
Dim cash As String
Dim BV_EQUITY As String
Dim BETA_5YR As String


On Error GoTo ERROR_LABEL

INDUSTRIES_VALUATION_FUNC = False
SRC_WSHEET_NAME = SRC_RNG.Worksheet.name & "!"

Range(DST_RNG.Cells(1, 1), _
DST_RNG.Cells(1, 1 + 17)).value = _
Array("Industries/Companies", "No. of Firms/Exchange", "Levered Beta", _
"Unlevered Beta", "Market D/E", "Market Debt/Capital", "ROE", _
"ROC", "Effective Tax Rate", "Pre-tax Operating Margin", _
"After-tax Operating Margin", "Net Margin", "Cap Ex/ Depreciation", _
"Non-cash WC/ Revenues", "Payout Ratio", "Reinvestment Rate", _
"Sales/Capital", "EV/Sales")

With DST_RNG

    For i = 1 To PUB_NSIZE
    
        INDUSTRY = SRC_RNG.Offset(i, P0).Address
        NO_FIRMS = SRC_RNG.Offset(i, p1).Address
        DIVD = SRC_RNG.Offset(i, p2).Address
        VL_BETA = SRC_RNG.Offset(i, P3).Address
        TAX_RATE = SRC_RNG.Offset(i, P8).Address
        Market_Cap = SRC_RNG.Offset(i, P11).Address
        TOTAL_DEBT = SRC_RNG.Offset(i, P12).Address
        FIRM_VALUE = SRC_RNG.Offset(i, P13).Address
        EV = SRC_RNG.Offset(i, P14).Address
        INV_CAPITAL = SRC_RNG.Offset(i, P15).Address
        NON_CASH_WC = SRC_RNG.Offset(i, P16).Address
        REINVEST = SRC_RNG.Offset(i, P18).Address
        SALES = SRC_RNG.Offset(i, P19).Address
        EBIT = SRC_RNG.Offset(i, P21).Address
        EBIT_ATAX = SRC_RNG.Offset(i, P23).Address
        DEPRECIATION = SRC_RNG.Offset(i, P24).Address
        CAPEX = SRC_RNG.Offset(i, p25).Address
        NET_INCOME = SRC_RNG.Offset(i, P26).Address
        TRAILING_NET_INCOME = SRC_RNG.Offset(i, P27).Address
        cash = SRC_RNG.Offset(i, P28).Address
        BV_EQUITY = SRC_RNG.Offset(i, P34).Address
        BETA_5YR = SRC_RNG.Offset(i, P38).Address
        
        
        .Offset(i, 0).formula = "=" & SRC_WSHEET_NAME & INDUSTRY
        .Offset(i, 1).formula = "=" & SRC_WSHEET_NAME & NO_FIRMS
        .Offset(i, 2).formula = "=MAX(" & SRC_WSHEET_NAME & VL_BETA & "," & _
        SRC_WSHEET_NAME & BETA_5YR & ")"
        
        .Offset(i, 3).formula = "=(MAX(" & SRC_WSHEET_NAME & VL_BETA & "," & _
        SRC_WSHEET_NAME & BETA_5YR & ")/(1+(1-(If(" & SRC_WSHEET_NAME & TAX_RATE & _
        "> 0.5, 0.5, If(" & SRC_WSHEET_NAME & TAX_RATE & _
        "< 0, 0," & SRC_WSHEET_NAME & TAX_RATE & " ))))*( " & _
        SRC_WSHEET_NAME & TOTAL_DEBT & "/" & SRC_WSHEET_NAME & Market_Cap & _
        ")))/(1-(" & SRC_WSHEET_NAME & cash & "/" & _
        SRC_WSHEET_NAME & FIRM_VALUE & "))"
        
        .Offset(i, 4).formula = "=" & SRC_WSHEET_NAME & TOTAL_DEBT & "/" & _
        SRC_WSHEET_NAME & Market_Cap
        
        .Offset(i, 5).formula = "=" & SRC_WSHEET_NAME & TOTAL_DEBT & "/" & _
        SRC_WSHEET_NAME & FIRM_VALUE
        
        .Offset(i, 6).formula = "=" & SRC_WSHEET_NAME & NET_INCOME & "/(" & _
        SRC_WSHEET_NAME & BV_EQUITY & "-" & SRC_WSHEET_NAME & cash & ")"
        
        .Offset(i, 7).formula = "=If(if(" & SRC_WSHEET_NAME & SALES & ">0," & _
        SRC_WSHEET_NAME & EBIT_ATAX & "/" & SRC_WSHEET_NAME & SALES & _
        ", ""NA"")=""NA"",""NA"",IF(" & SRC_WSHEET_NAME & SALES & "> 0, " & _
        SRC_WSHEET_NAME & EBIT_ATAX & "/" & SRC_WSHEET_NAME & SALES & _
        ", ""NA"")*IF(" & SRC_WSHEET_NAME & SALES & ">0," & SRC_WSHEET_NAME & _
        SALES & "/" & SRC_WSHEET_NAME & INV_CAPITAL & ",""NA""))"
        
        .Offset(i, 8).formula = "=If(" & SRC_WSHEET_NAME & TAX_RATE & _
        "> 0.5, 0.5, If(" & SRC_WSHEET_NAME & TAX_RATE & "< 0, 0," & _
        SRC_WSHEET_NAME & TAX_RATE & "))"
        
        .Offset(i, 9).formula = "=If(" & SRC_WSHEET_NAME & SALES & "> 0,(" & _
        SRC_WSHEET_NAME & EBIT & "/" & SRC_WSHEET_NAME & SALES & "),""NA"")"
        
        .Offset(i, 10).formula = "=If(" & SRC_WSHEET_NAME & SALES & ">0," & _
        SRC_WSHEET_NAME & EBIT_ATAX & "/" & SRC_WSHEET_NAME & SALES & ", ""NA"")"
        
        .Offset(i, 11).formula = "=If(" & SRC_WSHEET_NAME & SALES & ">0," & _
        SRC_WSHEET_NAME & TRAILING_NET_INCOME & "/" & SRC_WSHEET_NAME & SALES & ", ""NA"")"
        
        .Offset(i, 12).formula = "=If(" & SRC_WSHEET_NAME & DEPRECIATION & _
        ">0," & SRC_WSHEET_NAME & CAPEX & "/" & SRC_WSHEET_NAME & DEPRECIATION & ", ""NA"")"
        
        .Offset(i, 13).formula = "=If(" & SRC_WSHEET_NAME & SALES & "> 0,(" & _
        SRC_WSHEET_NAME & NON_CASH_WC & "/" & SRC_WSHEET_NAME & SALES & "),""NA"")"
        
        .Offset(i, 14).formula = "=If(" & SRC_WSHEET_NAME & NET_INCOME & "> 0,(" & _
        SRC_WSHEET_NAME & DIVD & "/" & SRC_WSHEET_NAME & NET_INCOME & "),""NA"")"
        
        .Offset(i, 15).formula = "=If(" & SRC_WSHEET_NAME & EBIT_ATAX & "> 0, If(" & _
        SRC_WSHEET_NAME & EBIT_ATAX & "> 0," & SRC_WSHEET_NAME & REINVEST & "/" & _
        SRC_WSHEET_NAME & EBIT_ATAX & ",""NA""),""NA"")"
        
        .Offset(i, 16).formula = "=If(" & SRC_WSHEET_NAME & SALES & ">0," & _
        SRC_WSHEET_NAME & SALES & "/" & SRC_WSHEET_NAME & INV_CAPITAL & ", ""NA"")"
        
        .Offset(i, 17).formula = "=If(" & SRC_WSHEET_NAME & SALES & ">0," & _
        SRC_WSHEET_NAME & EV & "/" & SRC_WSHEET_NAME & SALES & ", ""NA"")"
    
    Next i
End With

INDUSTRIES_VALUATION_FUNC = True

Exit Function
ERROR_LABEL:
INDUSTRIES_VALUATION_FUNC = False
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : INDUSTRIES_WACC_FUNC
'DESCRIPTION   : WACC TABLE
'LIBRARY       : FUNDAMENTAL
'GROUP         : INDUSTRIES
'ID            : 030

'LAST UPDATE   : 14 / 07 / 2009

'AUTHORS       : RAFAEL NICOLAS FERMIN COTA & CHRISTOPHER GILPIN
'REFERENCES    : http://pages.stern.nyu.edu/~adamodar/
'************************************************************************************
'************************************************************************************

Private Function INDUSTRIES_WACC_FUNC(ByRef SRC_RNG As Excel.Range, _
ByRef DST_RNG As Excel.Range)

Dim i As Long
Dim SRC_WSHEET_NAME As String

Dim INDUSTRY As String
Dim NO_FIRMS As String
Dim BETA_5YR As String
Dim VL_BETA As String
Dim Market_Cap As String
Dim FIRM_VALUE As String
Dim ST_DEV_3YR As String
Dim TAX_RATE As String

On Error GoTo ERROR_LABEL

INDUSTRIES_WACC_FUNC = False

SRC_WSHEET_NAME = SRC_RNG.Worksheet.name & "!"

Range(DST_RNG.Cells(1, 1), _
DST_RNG.Cells(1, 1 + 10)).value = _
Array("Industries/Companies", "No. of Firms/Exchange", "Beta", "Cost of Equity", _
"E/(D+E)", "Std Dev in Stock", "Cost of Debt", "Tax Rate", "After-tax Cost of Debt", _
"D/(D+E)", "Cost of Capital")


With DST_RNG

    For i = 1 To PUB_NSIZE
    
        INDUSTRY = SRC_RNG.Offset(i, P0).Address
        NO_FIRMS = SRC_RNG.Offset(i, p1).Address
        VL_BETA = SRC_RNG.Offset(i, P3).Address
        BETA_5YR = SRC_RNG.Offset(i, P38).Address
        Market_Cap = SRC_RNG.Offset(i, P11).Address
        FIRM_VALUE = SRC_RNG.Offset(i, P13).Address
        ST_DEV_3YR = SRC_RNG.Offset(i, P40).Address
        TAX_RATE = SRC_RNG.Offset(i, P8).Address
        
        
        
        .Offset(i, 0).formula = "=" & SRC_WSHEET_NAME & INDUSTRY
        .Offset(i, 1).formula = "=" & SRC_WSHEET_NAME & NO_FIRMS
        .Offset(i, 2).formula = "=MAX(" & SRC_WSHEET_NAME & VL_BETA & "," & _
        SRC_WSHEET_NAME & BETA_5YR & ")"
        
        .Offset(i, 3).formula = "= LT_Bonds +" & .Offset(i, 3).Offset(0, -1).Address & _
        " * RISK_PREMIUM"
        
        .Offset(i, 4).formula = "=" & SRC_WSHEET_NAME & Market_Cap & "/" & _
        SRC_WSHEET_NAME & FIRM_VALUE
        
        .Offset(i, 5).formula = "=" & SRC_WSHEET_NAME & ST_DEV_3YR
        
        .Offset(i, 6).formula = "= LT_Bonds + VLookup(" & _
        .Offset(i, 6).Offset(0, -1).Address & ", Debt_Table, 3)"
        
        .Offset(i, 7).formula = "=" & SRC_WSHEET_NAME & TAX_RATE
        
        .Offset(i, 8).formula = "=IF(TAX_SWITCHER=""YES"", (1 - MARGINAL_TAX_RATE) * " _
        & .Offset(i, 8).Offset(0, -2).Address & "," & .Offset(i, 8).Offset(0, -2).Address & _
        "*(1-" & .Offset(i, 8).Offset(0, -1).Address & "))"
        
        .Offset(i, 9).formula = "=1 -" & .Offset(i, 9).Offset(0, -5).Address
        
        .Offset(i, 10).formula = "=" & .Offset(i, 10).Offset(0, -7).Address & "*" & _
        .Offset(i, 10).Offset(0, -6).Address & "+" & .Offset(i, 10).Offset(0, -2).Address & _
        "*" & .Offset(i, 10).Offset(0, -1).Address
            
    Next i
End With

INDUSTRIES_WACC_FUNC = True

Exit Function
ERROR_LABEL:
INDUSTRIES_WACC_FUNC = False
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : INDUSTRIES_WORKING_CAPITAL_FUNC
'DESCRIPTION   : WORKKING CAPITAL TABLE
'LIBRARY       : FUNDAMENTAL
'GROUP         : INDUSTRIES
'ID            : 031
'LAST UPDATE   : 14 / 07 / 2009
'AUTHORS       : RAFAEL NICOLAS FERMIN COTA & CHRISTOPHER GILPIN
'REFERENCES    : http://pages.stern.nyu.edu/~adamodar/
'************************************************************************************
'************************************************************************************

Private Function INDUSTRIES_WORKING_CAPITAL_FUNC(ByRef SRC_RNG As Excel.Range, _
ByRef DST_RNG As Excel.Range)

Dim i As Long
Dim SRC_WSHEET_NAME As String

Dim INDUSTRY As String
Dim NO_FIRMS As String
Dim SALES As String
Dim AP As String
Dim AR As String
Dim INVENTORY As String
Dim cash As String
Dim NON_CASH_WC As String

On Error GoTo ERROR_LABEL

INDUSTRIES_WORKING_CAPITAL_FUNC = False

SRC_WSHEET_NAME = SRC_RNG.Worksheet.name & "!"

Range(DST_RNG.Cells(1, 1), _
DST_RNG.Cells(1, 1 + 7)).value = _
Array("Industries/Companies", "No. of Firms/Exchange", "Accounts Receivable/Sales", _
"Inventory/Sales", "Accounts Payable/Sales", "Cash/Sales", _
"Non-cash Working Capital/Sales", "Working capital/ Sales")

With DST_RNG
    For i = 1 To PUB_NSIZE
        INDUSTRY = SRC_RNG.Offset(i, P0).Address
        NO_FIRMS = SRC_RNG.Offset(i, p1).Address
        SALES = SRC_RNG.Offset(i, P19).Address
        AR = SRC_RNG.Offset(i, P29).Address
        AP = SRC_RNG.Offset(i, P33).Address
        INVENTORY = SRC_RNG.Offset(i, P30).Address
        cash = SRC_RNG.Offset(i, P28).Address
        NON_CASH_WC = SRC_RNG.Offset(i, P16).Address
        
        .Offset(i, 0).formula = "=" & SRC_WSHEET_NAME & INDUSTRY
        
        .Offset(i, 1).formula = "=" & SRC_WSHEET_NAME & NO_FIRMS
        
        .Offset(i, 2).formula = "=If(" & SRC_WSHEET_NAME & SALES & "> 0,(" & _
        SRC_WSHEET_NAME & AR & "/" & SRC_WSHEET_NAME & SALES & "),""NA"")"
        
        .Offset(i, 3).formula = "=If(" & SRC_WSHEET_NAME & SALES & "> 0,(" & _
        SRC_WSHEET_NAME & INVENTORY & "/" & SRC_WSHEET_NAME & SALES & "),""NA"")"
        
        .Offset(i, 4).formula = "=If(" & SRC_WSHEET_NAME & SALES & "> 0,(" & _
        SRC_WSHEET_NAME & AP & "/" & SRC_WSHEET_NAME & SALES & "),""NA"")"
        
        .Offset(i, 5).formula = "=If(" & SRC_WSHEET_NAME & SALES & "> 0,(" & _
        SRC_WSHEET_NAME & cash & "/" & SRC_WSHEET_NAME & SALES & "),""NA"")"
        
        .Offset(i, 6).formula = "=If(" & SRC_WSHEET_NAME & SALES & "> 0,(" & _
        SRC_WSHEET_NAME & NON_CASH_WC & "/" & SRC_WSHEET_NAME & SALES & "),""NA"")"
        
        .Offset(i, 7).formula = "=If(" & .Offset(i, 7).Offset(0, -2).Address & _
        "=""NA"",""NA"",(" & .Offset(i, 7).Offset(0, -1).Address & "+" & _
        .Offset(i, 7).Offset(0, -2).Address & "))"
    
    Next i
End With

INDUSTRIES_WORKING_CAPITAL_FUNC = True

Exit Function
ERROR_LABEL:
INDUSTRIES_WORKING_CAPITAL_FUNC = False
End Function
