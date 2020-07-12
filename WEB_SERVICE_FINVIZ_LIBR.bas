Attribute VB_Name = "WEB_SERVICE_FINVIZ_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.

'v=111&f=ind_stocksonly&ft=1&ta=1&p=d&r=1
'v=111&f=ind_exchangetradedfund&ft=1&ta=1&p=d&r=1

Public Const PUB_FINVIZ_URL_STR As String = "http://finviz.com/export.ashx?v=152"
Public Const PUB_FINVIZ_DATA_SUFFIX_STR As String = "&ft=1&ta=1&p=d&r=1&c=" & _
"1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16,17,18,19,20,21,22,23,24,25,26,27," & _
"28,29,30,31,32,33,34,35,36,37,38,39,40,41,42,43,44,45,46,47,48,49,50,51,52," & _
"53,54,55,56,57,58,59,60,61,62,63,64,65,66,67,68"

Private PUB_FINVIZ_INDUSTRY_OBJ As Collection

Public Sub SHOW_FINVIZ_FORM_FUNC()
On Error GoTo ERROR_LABEL
frmFinviz.show
Exit Sub
ERROR_LABEL:
End Sub

Function FINVIZ_COMPARABLE_ANALYSIS_FUNC(ByVal TICKER_STR As String, _
Optional ByVal OUTPUT As Integer = 1)

Dim i As Long
Dim DATA_STR As String
Dim SUFFIX_STR As String
Dim SRC_URL_STR As String
Dim TEMP_MATRIX As Variant

On Error GoTo ERROR_LABEL

SRC_URL_STR = PUB_FINVIZ_URL_STR & "&ft=1&t=" & TICKER_STR & "&ta=1&p=d&r=1&c=4"

DATA_STR = SAVE_WEB_DATA_PAGE_FUNC(SRC_URL_STR, 0, False, 0, False)
If DATA_STR = PUB_WEB_DATA_SYSTEM_ERROR_STR Then: GoTo ERROR_LABEL
DATA_STR = Replace(DATA_STR, Chr(34), "")
For i = 8 To 14: DATA_STR = Replace(DATA_STR, Chr(i), ""): Next i 'Remove HTML Syntax
If DATA_STR = "Industry" Then: GoTo ERROR_LABEL 'Wrong ticker symbol
DATA_STR = Trim(Replace(DATA_STR, "Industry", ""))
SUFFIX_STR = FINVIZ_INDUSTRY_TAG_FUNC(DATA_STR) 'Getting Industry Tag
If SUFFIX_STR = "0" Then: GoTo ERROR_LABEL

'------------------------------------------------------------------------------------------------
Select Case OUTPUT
'------------------------------------------------------------------------------------------------
Case 0
'------------------------------------------------------------------------------------------------
    SUFFIX_STR = "v=152&f=" & SUFFIX_STR & PUB_FINVIZ_DATA_SUFFIX_STR
    FINVIZ_COMPARABLE_ANALYSIS_FUNC = FINVIZ_SCREENER_FUNC(SUFFIX_STR)
'------------------------------------------------------------------------------------------------
Case Else
'------------------------------------------------------------------------------------------------
    SUFFIX_STR = "v=152&f=" & SUFFIX_STR & "&ft=1&ta=1&p=d&r=1&c=1"
    TEMP_MATRIX = FINVIZ_SCREENER_FUNC(SUFFIX_STR)
    If IsArray(TEMP_MATRIX) = False Then: GoTo ERROR_LABEL
    
    TEMP_MATRIX(1, 1) = UCase(TICKER_STR)
    TEMP_MATRIX = ARRAY_REMOVE_DUPLICATES_FUNC(TEMP_MATRIX, 0)
'    Debug.Print UBound(TEMP_MATRIX, 1), TEMP_MATRIX(1, 1), TEMP_MATRIX(UBound(TEMP_MATRIX, 1), 1)
    
    If OUTPUT = 1 Then
'        Debug.Print UBound(TEMP_MATRIX, 1)
        FINVIZ_COMPARABLE_ANALYSIS_FUNC = YAHOO_KEY_STATISTICS_FUNC(TEMP_MATRIX)
    Else
        FINVIZ_COMPARABLE_ANALYSIS_FUNC = TEMP_MATRIX
    End If
'------------------------------------------------------------------------------------------------
End Select
'------------------------------------------------------------------------------------------------

Exit Function
ERROR_LABEL:
FINVIZ_COMPARABLE_ANALYSIS_FUNC = Err.number
End Function

Function FINVIZ_BROKER_RECOMMENDATIONS_FUNC(ByVal TICKER_STR As String)

Dim h As Long
Dim i As Long
Dim j As Long
Dim k As Long
Dim l As Long
Dim TEMP_STR As String
Dim LOOK_STR As String
Dim DELIM_STR As String
Dim DATA_STR As String
Dim DATA_ARR As Variant
Dim SRC_URL_STR As String

On Error GoTo ERROR_LABEL

'Download current quotes
'------------------------------------------------------------------------------------------------
SRC_URL_STR = "http://finviz.com/quote.ashx?t=" & TICKER_STR
DATA_STR = SAVE_WEB_DATA_PAGE_FUNC(SRC_URL_STR, 0, False, 0, False)
'------------------------------------------------------------------------------------------------
i = InStr(1, DATA_STR, "fullview-ratings-outer")
If i = 0 Then: GoTo ERROR_LABEL
j = InStr(i, DATA_STR, "<tr><td>")
'------------------------------------------------------------------------------------------------
DATA_STR = Mid(DATA_STR, i, j - i): DELIM_STR = "|"
DATA_STR = Replace(DATA_STR, "<b>", "")
DATA_STR = Replace(DATA_STR, "</b>", "")
DATA_STR = Replace(DATA_STR, "&rarr; ", "-> ")
'------------------------------------------------------------------------------------------------
i = 1: h = 0: ReDim DATA_ARR(1 To 1)
Do
    LOOK_STR = "align=": l = Len(LOOK_STR): TEMP_STR = ""
    For k = 1 To 5
        i = InStr(i, DATA_STR, LOOK_STR)
        If i = 0 Then: Exit For
        i = i + l
        i = InStr(i, DATA_STR, ">") + 1
        j = InStr(i, DATA_STR, "<")
        TEMP_STR = TEMP_STR & Mid(DATA_STR, i, j - i) & DELIM_STR
        i = j + 1
    Next k
    i = InStr(i, DATA_STR, "</td>")
    h = h + 1: ReDim Preserve DATA_ARR(1 To h): DATA_ARR(h) = TEMP_STR
Loop Until InStr(i, DATA_STR, LOOK_STR) = 0

ReDim TEMP_MATRIX(0 To h, 1 To 5)
TEMP_MATRIX(0, 1) = "DATE"
TEMP_MATRIX(0, 2) = "ACTION"
TEMP_MATRIX(0, 3) = "RESEARCH FIRM"
TEMP_MATRIX(0, 4) = "FROM/TO"
TEMP_MATRIX(0, 5) = "TARGET"

For l = 1 To h
    TEMP_STR = DATA_ARR(l): i = 1
    For k = 1 To 5
        j = InStr(i, TEMP_STR, DELIM_STR)
        TEMP_MATRIX(l, k) = Mid(TEMP_STR, i, j - i)
        i = j + 1
    Next k
Next l
FINVIZ_BROKER_RECOMMENDATIONS_FUNC = CONVERT_STRING_NUMBER_FUNC(TEMP_MATRIX)

Exit Function
ERROR_LABEL:
FINVIZ_BROKER_RECOMMENDATIONS_FUNC = Err.number
End Function

Function FINVIZ_SCREENER_FUNC( _
Optional ByVal SUFFIX_STR As String = _
"v=111&f=fa_curratio_o2,fa_debteq_u0.3,fa_fpe_low,fa_pb_u2," & _
"fa_pe_low,fa_peg_low,fa_quickratio_o2&ft=2&ta=1&p=d&o=pb&r=1")

'Debt/Equity <30% --> dont require tons of debt to operate
'Current Ratio >2 --> Have adequate short term liquidity
'Quick Ratio > 2 --> Have adequate short term liquidity without their inventories
'P/B < 2 --> Trading near its liquidation value
'Trailing P/E < 15 --> Producing significant earnings relative to its marcap cap (both past and future)
'FORWARD PE < 15
'PEG < 1
'PE-PEG -are subject to downward revisions as the
'macro picture deteriorates more and more.

Dim DATA_STR As String
Dim SRC_URL_STR As String

On Error GoTo ERROR_LABEL

'Download current quotes
SRC_URL_STR = Replace(PUB_FINVIZ_URL_STR, "v=152", "") & SUFFIX_STR
DATA_STR = SAVE_WEB_DATA_PAGE_FUNC(SRC_URL_STR, 0, False, 0, False) '& Chr(13)

'Parse returned data
'DATA_STR = Replace(DATA_STR, Chr(10), Chr(13))
'DATA_STR = Replace(DATA_STR, Chr(13) & Chr(13), Chr(13))
'DATA_STR = Replace(DATA_STR, Chr(13) & Chr(13), Chr(13))

FINVIZ_SCREENER_FUNC = FINVIZ_PARSER_FUNC(DATA_STR)

Exit Function
ERROR_LABEL:
FINVIZ_SCREENER_FUNC = Err.number
End Function

'Screen on over 7000 US stocks. Compare your screening results to market cap weighted aggregates at
'the industry and sector level

'http://www.gummy-stuff.org/Mkt_Indexes.htm
'http://www.gummy-stuff.org/Indexes2.htm
'http://www.gummy-stuff.org/Wil-GDP.htm
'http://www.gummy-stuff.org/PE-ratios.htm
'http://www.gummy-stuff.org/gRANK.htm

Function FINVIZ_MARKET_CAP_WEIGHTED_REPORT_FUNC(Optional ByRef DATA_RNG As Variant, _
Optional ByVal NCOLUMNS As Long = 62) '58)

Dim h As Long
Dim i As Long
Dim j As Long
Dim k As Long

Const l As Long = 5
'v=151&c=1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16,17,18,19,20,21,22,23,24,25,26,27,28,29,30,
'31,32,33,34,35,36,37,38,39,40,41,42,43,44,45,46,47,48,49,50,51,52,53,54,55,56,57,
'58,59,60,61,62,63,64,65,66,67,68

Dim SROW As Long
Dim NROWS As Long

Dim NSIZE As Long

Dim NUMER_VAL As Double
Dim DENOM_VAL As Double

Dim SECTOR_STR As String
Dim INDUSTRY_STR As String

Dim SECTORS_ARR As Variant
Dim INDUSTRIES_ARR As Variant

Dim TEMP_VECTOR As Variant
Dim TEMP_MATRIX As Variant
Dim DATA_MATRIX As Variant

On Error GoTo ERROR_LABEL

If IsArray(DATA_RNG) = True Then
    DATA_MATRIX = DATA_RNG
Else
    DATA_MATRIX = FINVIZ_DATABASE_FUNC
End If
SROW = LBound(DATA_MATRIX, 1)
NROWS = UBound(DATA_MATRIX, 1)
If NCOLUMNS + l > UBound(DATA_MATRIX, 2) Then: NCOLUMNS = UBound(DATA_MATRIX, 2) - l
'Ticker  Company Sector  Industry    Country
GoSub LOAD_HEADINGS

h = SROW + 1
ReDim TEMP_VECTOR(1 To 3, 1 To NCOLUMNS) 'All
For i = SROW + 1 To NROWS
    For j = 1 To NCOLUMNS
        GoSub FILTERS_LINE
    Next j
Next i
GoSub LOAD_MATRIX
h = h + 1
For k = LBound(SECTORS_ARR) To UBound(SECTORS_ARR) 'Sectors
    TEMP_MATRIX(h, 1) = SECTORS_ARR(k)
    SECTOR_STR = SECTORS_ARR(k)
    For i = SROW + 1 To NROWS
        If DATA_MATRIX(i, l - 2) = SECTOR_STR Then
            For j = 1 To NCOLUMNS
                GoSub FILTERS_LINE
            Next j
        End If
    Next i
    GoSub LOAD_MATRIX
    h = h + 1
Next k
For k = LBound(INDUSTRIES_ARR) To UBound(INDUSTRIES_ARR) 'Industries
    TEMP_MATRIX(h, 1) = INDUSTRIES_ARR(k)
    INDUSTRY_STR = INDUSTRIES_ARR(k)
    For i = SROW + 1 To NROWS
        If DATA_MATRIX(i, l - 1) = INDUSTRY_STR Then
            For j = 1 To NCOLUMNS
                GoSub FILTERS_LINE
            Next j
        End If
    Next i
    GoSub LOAD_MATRIX
    h = h + 1
Next k

Erase TEMP_VECTOR: Erase DATA_MATRIX
FINVIZ_MARKET_CAP_WEIGHTED_REPORT_FUNC = TEMP_MATRIX

Exit Function
'----------------------------------------------------------------------------------------
LOAD_HEADINGS:
'----------------------------------------------------------------------------------------
    NSIZE = 1 'All
    SECTORS_ARR = Array("Basic Materials", "Conglomerates", "Consumer Goods", _
                        "Financial", "Healthcare", "Industrial Goods", _
                        "Services", "Technology", "Utilities")
    NSIZE = NSIZE + UBound(SECTORS_ARR) - LBound(SECTORS_ARR) + 1
    INDUSTRIES_ARR = Array("Accident & Health Insurance", "Advertising Agencies", "Aerospace/Defense - Major Diversified", "Aerospace/Defense Products & Services", "Agricultural Chemicals", "Air Delivery & Freight Services", _
    "Air Services, Other", "Aluminum", "Apparel Stores", "Appliances", "Application Software", "Asset Management", "Auto Dealerships", "Auto Manufacturers - Major", "Auto Parts", "Auto Parts Stores", "Auto Parts Wholesale", _
    "Basic Materials Wholesale", "Beverages - Brewers", "Beverages - Soft Drinks", "Beverages - Wineries & Distillers", "Biotechnology", "Broadcasting - Radio", "Broadcasting - TV", "Building Materials Wholesale", "Business Equipment", _
    "Business Services", "Business Software & Services", "Catalog & Mail Order Houses", "CATV Systems", "Cement", "Chemicals - Major Diversified", "Cigarettes", "Cleaning Products", "Closed-End Fund - Debt", "Closed-End Fund - Equity", _
    "Closed-End Fund - Foreign", "Communication Equipment", "Computer Based Systems", "Computer Peripherals", "Computers Wholesale", "Confectioners", "Conglomerates", "Consumer Services", "Copper", "Credit Services", "Dairy Products", _
    "Data Storage Devices", "Department Stores", "Diagnostic Substances", "Discount, Variety Stores", "Diversified Communication Services", "Diversified Computer Systems", "Diversified Electronics", "Diversified Investments", "Diversified Machinery", "Diversified Utilities", _
    "Drug Delivery", "Drug Manufacturers - Major", "Drug Manufacturers - Other", "Drug Related Products", "Drug Stores", "Drugs - Generic", "Drugs Wholesale", "Education & Training Services", "Electric Utilities", "Electronic Equipment", "Electronics Stores", "Electronics Wholesale", _
    "Entertainment - Diversified", "Exchange Traded Fund", "Farm & Construction Machinery", "Farm Products", "Food - Major Diversified", "Food Wholesale", "Foreign Money Center Banks", "Foreign Regional Banks", "Foreign Utilities", "Gaming Activities", "Gas Utilities", "General Building Materials", "General Contractors", _
    "General Entertainment", "Gold", "Grocery Stores", "Health Care Plans", "Healthcare Information Services", "Heavy Construction", "Home Furnishing Stores", "Home Furnishings & Fixtures", "Home Health Care", "Home Improvement Stores", "Hospitals", "Housewares & Accessories", "Independent Oil & Gas", _
    "Industrial Electrical Equipment", "Industrial Equipment & Components", "Industrial Equipment Wholesale", "Industrial Metals & Minerals", "Information & Delivery Services", "Information Technology Services", "Insurance Brokers", "Internet Information Providers", "Internet Service Providers", "Internet Software & Services", "Investment Brokerage - National", _
    "Investment Brokerage - Regional", "Jewelry Stores", "Life Insurance", "Lodging", "Long Distance Carriers", "Long-Term Care Facilities", "Lumber, Wood Production", "Machine Tools & Accessories", "Major Airlines", "Major Integrated Oil & Gas", "Management Services", "Manufactured Housing", "Marketing Services", "Meat Products", "Medical Appliances & Equipment", "Medical Equipment Wholesale", _
    "Medical Instruments & Supplies", "Medical Laboratories & Research", "Medical Practitioners", "Metal Fabrication", "Money Center Banks", "Mortgage Investment", "Movie Production, Theaters", "Multimedia & Graphics Software", "Music & Video Stores", "Networking & Communication Devices", "Nonmetallic Mineral Mining", "Office Supplies", "Oil & Gas Drilling & Exploration", _
    "Oil & Gas Equipment & Services", "Oil & Gas Pipelines", "Oil & Gas Refining & Marketing", "Packaging & Containers", "Paper & Paper Products", "Personal Computers", "Personal Products", "Personal Services", "Photographic Equipment & Supplies", "Pollution & Treatment Controls", "Printed Circuit Boards", "Processed & Packaged Goods", "Processing Systems & Products", "Property & Casualty Insurance", "Property Management", _
    "Publishing - Books", "Publishing - Newspapers", "Publishing - Periodicals", "Railroads", "Real Estate Development", "Recreational Goods, Other", "Recreational Vehicles", "Regional - Mid-Atlantic Banks", "Regional - Midwest Banks", "Regional - Northeast Banks", "Regional - Pacific Banks", "Regional - Southeast Banks", "Regional - Southwest Banks", "Regional Airlines", "REIT - Diversified", "REIT - Healthcare Facilities", "REIT - Hotel/Motel", _
    "REIT - Industrial", "REIT - Office", "REIT - Residential", "REIT - Retail", "Rental & Leasing Services", "Research Services", "Residential Construction", "Resorts & Casinos", "Restaurants", "Rubber & Plastics", "Savings & Loans", "Scientific & Technical Instruments", "Security & Protection Services", "Security Software & Services", "Semiconductor - Broad Line", "Semiconductor - Integrated Circuits", _
    "Semiconductor - Specialized", "Semiconductor Equipment & Materials", "Semiconductor- Memory Chips", "Shipping", "Silver", "Small Tools & Accessories", "Specialized Health Services", "Specialty Chemicals", "Specialty Eateries", "Specialty Retail, Other", "Sporting Activities", "Sporting Goods", "Sporting Goods Stores", "Staffing & Outsourcing Services", "Steel & Iron", "Surety & Title Insurance", "Synthetics", "Technical & System Software", "Technical Services", _
    "Telecom Services - Domestic", "Telecom Services - Foreign", "Textile - Apparel Clothing", "Textile - Apparel Footwear & Accessories", "Textile Industrial", "Tobacco Products, Other", "Toy & Hobby Stores", "Toys & Games", "Trucking", "Trucks & Other Vehicles", "Waste Management", "Water Utilities", "Wholesale, Other", "Wireless Communications")
    NSIZE = NSIZE + UBound(INDUSTRIES_ARR) - LBound(INDUSTRIES_ARR) + 1
    ReDim TEMP_MATRIX(SROW To SROW + NSIZE, 1 To NCOLUMNS + 1)
    TEMP_MATRIX(SROW, 1) = "Sectors/Industries"
    TEMP_MATRIX(SROW + 1, 1) = "All"
    For j = 1 To NCOLUMNS
        TEMP_MATRIX(SROW, j + 1) = DATA_MATRIX(SROW, l + j)
    Next j
Return
'----------------------------------------------------------------------------------------
FILTERS_LINE:
'----------------------------------------------------------------------------------------
    If IsNumeric(DATA_MATRIX(i, l + j)) Then
        NUMER_VAL = DATA_MATRIX(i, l + j)
        DENOM_VAL = 1
    Else
        NUMER_VAL = 0
        DENOM_VAL = 0
    End If
    If IsNumeric(DATA_MATRIX(i, l + 1)) Then
        NUMER_VAL = NUMER_VAL * DATA_MATRIX(i, l + 1)
        DENOM_VAL = DENOM_VAL * DATA_MATRIX(i, l + 1)
    Else
        NUMER_VAL = NUMER_VAL * 0
        DENOM_VAL = DENOM_VAL * 0
    End If
    TEMP_VECTOR(1, j) = TEMP_VECTOR(1, j) + NUMER_VAL
    TEMP_VECTOR(2, j) = TEMP_VECTOR(2, j) + DENOM_VAL
    If TEMP_VECTOR(2, j) <> 0 Then: TEMP_VECTOR(3, j) = TEMP_VECTOR(3, j) + 1
'----------------------------------------------------------------------------------------
Return
'----------------------------------------------------------------------------------------
LOAD_MATRIX:
'----------------------------------------------------------------------------------------
    TEMP_MATRIX(h, 1) = TEMP_MATRIX(h, 1) & " (" & TEMP_VECTOR(3, 1) & ")"
    For j = 1 To NCOLUMNS
        If IsNumeric(TEMP_VECTOR(1, j)) And IsNumeric(TEMP_VECTOR(2, j)) Then
            NUMER_VAL = TEMP_VECTOR(1, j)
            DENOM_VAL = TEMP_VECTOR(2, j)
            If DENOM_VAL <> 0 Then
                TEMP_MATRIX(h, j + 1) = TEMP_MATRIX(h, j + 1) + NUMER_VAL / DENOM_VAL
            End If
        End If
        TEMP_VECTOR(1, j) = 0
        TEMP_VECTOR(2, j) = 0
        TEMP_VECTOR(3, j) = 0
    Next j
'----------------------------------------------------------------------------------------
Return
'----------------------------------------------------------------------------------------
ERROR_LABEL:
FINVIZ_MARKET_CAP_WEIGHTED_REPORT_FUNC = Err.number
End Function

Function FINVIZ_DATABASE_FUNC()

Dim SRC_URL_STR As String
Dim DATA_STR As String

On Error GoTo ERROR_LABEL

SRC_URL_STR = PUB_FINVIZ_URL_STR & PUB_FINVIZ_DATA_SUFFIX_STR
DATA_STR = SAVE_WEB_DATA_PAGE_FUNC(SRC_URL_STR, 0, False, 0, False)

FINVIZ_DATABASE_FUNC = FINVIZ_PARSER_FUNC(DATA_STR)

Exit Function
ERROR_LABEL:
FINVIZ_DATABASE_FUNC = Err.number
End Function

Private Function FINVIZ_PARSER_FUNC(ByVal DATA_STR As String)

Dim i As Long
Dim j As Long

Dim ii As Long
Dim jj As Long

Dim SROW As Long
Dim NROWS As Long
Dim NCOLUMNS As Long
Dim TEMP_STR As String

Dim TEMP1_VAL As Variant
Dim TEMP2_VAL As Variant

Dim FACTOR_VAL As Variant

Dim DATA_ARR As Variant
Dim TEMP_MATRIX As Variant

On Error GoTo ERROR_LABEL

'DATA_STR = Replace(DATA_STR, Chr(34), "")
'DATA_STR = Split(DATA_STR, Chr(10))

DATA_STR = Replace(DATA_STR, vbCrLf, vbLf)
DATA_ARR = Split(DATA_STR, vbLf)
SROW = LBound(DATA_ARR, 1)
NROWS = UBound(DATA_ARR, 1)
NCOLUMNS = 0
TEMP_STR = DATA_ARR(SROW)
i = 1
Do
    NCOLUMNS = NCOLUMNS + 1
    j = InStr(i, TEMP_STR, ",")
    i = j + 1
Loop Until j = 0

ReDim TEMP_MATRIX(1 To NROWS, 1 To NCOLUMNS)

For ii = SROW To NROWS - 1
    i = 1
    For jj = 1 To NCOLUMNS
        If i > Len(DATA_ARR(ii)) Then Exit For
        TEMP_STR = IIf(Mid(DATA_ARR(ii), i, 1) = Chr(34), Chr(34), "") & ","
        j = InStr(i, DATA_ARR(ii) & ",", TEMP_STR)
        
        TEMP1_VAL = Mid(DATA_ARR(ii), i + Len(TEMP_STR) - 1, j - i - Len(TEMP_STR) + 1)
        TEMP2_VAL = Trim(TEMP1_VAL)
        
        If Right(TEMP2_VAL, 1) = "%" Then
           FACTOR_VAL = 100
           TEMP2_VAL = Left(TEMP2_VAL, Len(TEMP2_VAL) - 1)
        Else
           FACTOR_VAL = 1
        End If
        
        On Error Resume Next
            TEMP1_VAL = CDec(TEMP2_VAL) / FACTOR_VAL
        On Error GoTo ERROR_LABEL
        
        TEMP_MATRIX(ii + 1, jj) = TEMP1_VAL
        i = j + Len(TEMP_STR)
    Next jj
Next ii

FINVIZ_PARSER_FUNC = TEMP_MATRIX

Exit Function
ERROR_LABEL:
FINVIZ_PARSER_FUNC = Err.number
End Function

Private Function FINVIZ_INDUSTRY_TAG_FUNC(ByVal INDUSTRY_STR As String)

Dim i As Long
Dim j As Long
Dim k As Long
Dim l As Long
Dim ATEMP_STR As String
Dim BTEMP_STR As String
Dim HEADINGS_STR As String

On Error GoTo ERROR_LABEL

If PUB_FINVIZ_INDUSTRY_OBJ Is Nothing = True Then
    HEADINGS_STR = "accident & health insurance|ind_accidenthealthinsurance|advertising agencies|ind_advertisingagencies|aerospace/defense - major diversified|ind_aerospacedefensemajordiversified|aerospace/defense products & services|ind_aerospacedefenseproductsservices|agricultural chemicals|ind_agriculturalchemicals|air delivery & freight services|ind_airdeliveryfreightservices|air services, other|ind_airservicesother|aluminum|ind_aluminum|apparel stores|ind_apparelstores|appliances|ind_appliances|application software|ind_applicationsoftware|asset management|ind_assetmanagement|auto dealerships|ind_autodealerships|auto manufacturers - major|ind_automanufacturersmajor|auto parts|ind_autoparts|auto parts stores|ind_autopartsstores|auto parts wholesale|ind_autopartswholesale|basic materials wholesale|ind_basicmaterialswholesale|beverages - brewers|ind_beveragesbrewers|beverages - soft drinks|ind_beveragessoftdrinks|beverages - wineries & distillers|ind_beverageswineriesdistillers|"
    HEADINGS_STR = HEADINGS_STR & "biotechnology|ind_biotechnology|broadcasting - radio|ind_broadcastingradio|broadcasting - tv|ind_broadcastingtv|building materials wholesale|ind_buildingmaterialswholesale|business equipment|ind_businessequipment|business services|ind_businessservices|business software & services|ind_businesssoftwareservices|catalog & mail order houses|ind_catalogmailorderhouses|catv systems|ind_catvsystems|cement|ind_cement|chemicals - major diversified|ind_chemicalsmajordiversified|cigarettes|ind_cigarettes|cleaning products|ind_cleaningproducts|closed-end fund - debt|ind_closedendfunddebt|closed-end fund - equity|ind_closedendfundequity|closed-end fund - foreign|ind_closedendfundforeign|communication equipment|ind_communicationequipment|computer based systems|ind_computerbasedsystems|computer peripherals|ind_computerperipherals|computers wholesale|ind_computerswholesale|confectioners|ind_confectioners|"
    HEADINGS_STR = HEADINGS_STR & "conglomerates|ind_conglomerates|consumer services|ind_consumerservices|copper|ind_copper|credit services|ind_creditservices|dairy products|ind_dairyproducts|data storage devices|ind_datastoragedevices|department stores|ind_departmentstores|diagnostic substances|ind_diagnosticsubstances|discount, variety stores|ind_discountvarietystores|diversified communication services|ind_diversifiedcommunicationservices|diversified computer systems|ind_diversifiedcomputersystems|diversified electronics|ind_diversifiedelectronics|diversified investments|ind_diversifiedinvestments|diversified machinery|ind_diversifiedmachinery|diversified utilities|ind_diversifiedutilities|drug delivery|ind_drugdelivery|drug manufacturers - major|ind_drugmanufacturersmajor|drug manufacturers - other|ind_drugmanufacturersother|drug related products|ind_drugrelatedproducts|drug stores|ind_drugstores|drugs - generic|ind_drugsgeneric|"
    HEADINGS_STR = HEADINGS_STR & "drugs wholesale|ind_drugswholesale|education & training services|ind_educationtrainingservices|electric utilities|ind_electricutilities|electronic equipment|ind_electronicequipment|electronics stores|ind_electronicsstores|electronics wholesale|ind_electronicswholesale|entertainment - diversified|ind_entertainmentdiversified|exchange traded fund|ind_exchangetradedfund|farm & construction machinery|ind_farmconstructionmachinery|farm products|ind_farmproducts|food - major diversified|ind_foodmajordiversified|food wholesale|ind_foodwholesale|foreign money center banks|ind_foreignmoneycenterbanks|foreign regional banks|ind_foreignregionalbanks|foreign utilities|ind_foreignutilities|gaming activities|ind_gamingactivities|gas utilities|ind_gasutilities|general building materials|ind_generalbuildingmaterials|general contractors|ind_generalcontractors|general entertainment|ind_generalentertainment|gold|ind_gold|"
    HEADINGS_STR = HEADINGS_STR & "grocery stores|ind_grocerystores|health care plans|ind_healthcareplans|healthcare information services|ind_healthcareinformationservices|heavy construction|ind_heavyconstruction|home furnishing stores|ind_homefurnishingstores|home furnishings & fixtures|ind_homefurnishingsfixtures|home health care|ind_homehealthcare|home improvement stores|ind_homeimprovementstores|hospitals|ind_hospitals|housewares & accessories|ind_housewaresaccessories|independent oil & gas|ind_independentoilgas|industrial electrical equipment|ind_industrialelectricalequipment|industrial equipment & components|ind_industrialequipmentcomponents|industrial equipment wholesale|ind_industrialequipmentwholesale|industrial metals & minerals|ind_industrialmetalsminerals|information & delivery services|ind_informationdeliveryservices|information technology services|ind_informationtechnologyservices|insurance brokers|ind_insurancebrokers|internet information providers|ind_internetinformationproviders|" & _
    "internet service providers|ind_internetserviceproviders|internet software & services|ind_internetsoftwareservices|"
    HEADINGS_STR = HEADINGS_STR & "investment brokerage - national|ind_investmentbrokeragenational|investment brokerage - regional|ind_investmentbrokerageregional|jewelry stores|ind_jewelrystores|life insurance|ind_lifeinsurance|lodging|ind_lodging|long distance carriers|ind_longdistancecarriers|long-term care facilities|ind_longtermcarefacilities|lumber, wood production|ind_lumberwoodproduction|machine tools & accessories|ind_machinetoolsaccessories|major airlines|ind_majorairlines|major integrated oil & gas|ind_majorintegratedoilgas|management services|ind_managementservices|manufactured housing|ind_manufacturedhousing|marketing services|ind_marketingservices|meat products|ind_meatproducts|medical appliances & equipment|ind_medicalappliancesequipment|medical equipment wholesale|ind_medicalequipmentwholesale|medical instruments & supplies|ind_medicalinstrumentssupplies|medical laboratories & research|ind_medicallaboratoriesresearch|medical practitioners|ind_medicalpractitioners|metal fabrication|" & _
    "ind_metalfabrication|"
    HEADINGS_STR = HEADINGS_STR & "money center banks|ind_moneycenterbanks|mortgage investment|ind_mortgageinvestment|movie production, theaters|ind_movieproductiontheaters|multimedia & graphics software|ind_multimediagraphicssoftware|music & video stores|ind_musicvideostores|networking & communication devices|ind_networkingcommunicationdevices|nonmetallic mineral mining|ind_nonmetallicmineralmining|office supplies|ind_officesupplies|oil & gas drilling & exploration|ind_oilgasdrillingexploration|oil & gas equipment & services|ind_oilgasequipmentservices|oil & gas pipelines|ind_oilgaspipelines|oil & gas refining & marketing|ind_oilgasrefiningmarketing|packaging & containers|ind_packagingcontainers|paper & paper products|ind_paperpaperproducts|personal computers|ind_personalcomputers|personal products|ind_personalproducts|personal services|ind_personalservices|photographic equipment & supplies|ind_photographicequipmentsupplies|pollution & treatment controls|ind_pollutiontreatmentcontrols|" & _
    "printed circuit boards|ind_printedcircuitboards|processed & packaged goods|ind_processedpackagedgoods|"
    HEADINGS_STR = HEADINGS_STR & "processing systems & products|ind_processingsystemsproducts|property & casualty insurance|ind_propertycasualtyinsurance|property management|ind_propertymanagement|publishing - books|ind_publishingbooks|publishing - newspapers|ind_publishingnewspapers|publishing - periodicals|ind_publishingperiodicals|railroads|ind_railroads|real estate development|ind_realestatedevelopment|recreational goods, other|ind_recreationalgoodsother|recreational vehicles|ind_recreationalvehicles|regional - mid-atlantic banks|ind_regionalmidatlanticbanks|regional - midwest banks|ind_regionalmidwestbanks|regional - northeast banks|ind_regionalnortheastbanks|regional - pacific banks|ind_regionalpacificbanks|regional - southeast banks|ind_regionalsoutheastbanks|regional - southwest banks|ind_regionalsouthwestbanks|regional airlines|ind_regionalairlines|reit - diversified|ind_reitdiversified|reit - healthcare facilities|ind_reithealthcarefacilities|reit - hotel/motel|ind_reithotelmotel|" & _
    "reit - industrial|ind_reitindustrial|"
    HEADINGS_STR = HEADINGS_STR & "reit - office|ind_reitoffice|reit - residential|ind_reitresidential|reit - retail|ind_reitretail|rental & leasing services|ind_rentalleasingservices|research services|ind_researchservices|residential construction|ind_residentialconstruction|resorts & casinos|ind_resortscasinos|restaurants|ind_restaurants|rubber & plastics|ind_rubberplastics|savings & loans|ind_savingsloans|scientific & technical instruments|ind_scientifictechnicalinstruments|security & protection services|ind_securityprotectionservices|security software & services|ind_securitysoftwareservices|semiconductor - broad line|ind_semiconductorbroadline|semiconductor - integrated circuits|ind_semiconductorintegratedcircuits|semiconductor - specialized|ind_semiconductorspecialized|semiconductor equipment & materials|ind_semiconductorequipmentmaterials|semiconductor- memory chips|ind_semiconductormemorychips|shipping|ind_shipping|silver|ind_silver|small tools & accessories|ind_smalltoolsaccessories|"
    HEADINGS_STR = HEADINGS_STR & "specialized health services|ind_specializedhealthservices|specialty chemicals|ind_specialtychemicals|specialty eateries|ind_specialtyeateries|specialty retail, other|ind_specialtyretailother|sporting activities|ind_sportingactivities|sporting goods|ind_sportinggoods|sporting goods stores|ind_sportinggoodsstores|staffing & outsourcing services|ind_staffingoutsourcingservices|steel & iron|ind_steeliron|stocks only (ex-funds)|ind_stocksonly|surety & title insurance|ind_suretytitleinsurance|synthetics|ind_synthetics|technical & system software|ind_technicalsystemsoftware|technical services|ind_technicalservices|telecom services - domestic|ind_telecomservicesdomestic|telecom services - foreign|ind_telecomservicesforeign|textile - apparel clothing|ind_textileapparelclothing|textile - apparel footwear & accessories|ind_textileapparelfootwearaccessories|textile industrial|ind_textileindustrial|tobacco products, other|ind_tobaccoproductsother|toy & hobby stores|ind_toyhobbystores|"
    HEADINGS_STR = HEADINGS_STR & "toys & games|ind_toysgames|trucking|ind_trucking|trucks & other vehicles|ind_trucksothervehicles|waste management|ind_wastemanagement|water utilities|ind_waterutilities|wholesale, other|ind_wholesaleother|wireless communications|ind_wirelesscommunications|"
    j = Len(HEADINGS_STR)
    l = 0
    For i = 1 To j
        If Mid(HEADINGS_STR, i, 1) = "|" Then: l = l + 1
    Next i
    
    Set PUB_FINVIZ_INDUSTRY_OBJ = New Collection
    i = 1
    For k = 1 To l / 2
        j = InStr(i, HEADINGS_STR, "|")
        ATEMP_STR = Mid(HEADINGS_STR, i, j - i)
        i = j + 1
        j = InStr(i, HEADINGS_STR, "|")
        BTEMP_STR = Mid(HEADINGS_STR, i, j - i)
        Call PUB_FINVIZ_INDUSTRY_OBJ.Add(ATEMP_STR, BTEMP_STR)
        Call PUB_FINVIZ_INDUSTRY_OBJ.Add(BTEMP_STR, ATEMP_STR)
        i = j + 1
    Next k
End If

FINVIZ_INDUSTRY_TAG_FUNC = PUB_FINVIZ_INDUSTRY_OBJ(INDUSTRY_STR)

Exit Function
ERROR_LABEL:
FINVIZ_INDUSTRY_TAG_FUNC = Err.number
End Function
