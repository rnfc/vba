Attribute VB_Name = "FINAN_ENERGY_CALC_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1
'as the default first index of arrays.


'************************************************************************************
'************************************************************************************
'FUNCTION      : ENERGY_SYSTEMS_CALC_FUNC

'DESCRIPTION   : Converts a number from one measurement system to another.
'For example, this function can translate a table of distances in miles to
'a table of distances in kilometers.

'LIBRARY       : ENERGY
'GROUP         : CALC
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/29/2009
'************************************************************************************
'************************************************************************************

Function ENERGY_SYSTEMS_CALC_FUNC(ByVal VERSION As Integer, _
ByVal DATA_VAL As Variant, _
ByVal BASE_VAL As Integer, _
Optional ByVal OUTPUT As Integer = 0)

Dim i As Long
Dim NROWS As Long

Dim CALC_STR As String
Dim CALC_FUNC_STR As String
Dim CALC_FIELD_STR As String

Dim TEMP_MATRIX As Variant

On Error GoTo ERROR_LABEL

Select Case VERSION
    Case 0
        NROWS = 9 'AREA_CONSTANT_FUNC HAS 9 ENTRIES
        CALC_STR = "AREA"
    Case 1
        NROWS = 14
        CALC_STR = "LENGTH"
    Case 2
        NROWS = 21
        CALC_STR = "MASS"
    Case 3
        NROWS = 10
        CALC_STR = "WEIGHT"
    Case 4
        NROWS = 8
        CALC_STR = "TIME"
    Case 5
        NROWS = 4
        CALC_STR = "TEMPERATURE"
    Case 6
        NROWS = 20
        CALC_STR = "VOLUME"
    Case Else
        NROWS = 15
        CALC_STR = "ENERGY"
End Select

CALC_FUNC_STR = CALC_STR & "_CONSTANT_FUNC"
CALC_FIELD_STR = CALC_STR & "_CONSTANT_FIELD_FUNC"

'----------------------------------------------------------------------
Select Case OUTPUT
'----------------------------------------------------------------------
Case 0
'----------------------------------------------------------------------
    ReDim TEMP_MATRIX(1 To NROWS, 1 To 2)
    For i = 1 To NROWS
        If CALC_STR = "TEMPERATURE" Then
            TEMP_MATRIX(i, 2) = _
            Excel.Application.Run(CALC_FUNC_STR, DATA_VAL, BASE_VAL, 0) * _
            Excel.Application.Run(CALC_FUNC_STR, DATA_VAL, i, 1) + _
            Excel.Application.Run(CALC_FUNC_STR, DATA_VAL, i, 2)
        Else
            TEMP_MATRIX(i, 2) = _
            Excel.Application.Run(CALC_FUNC_STR, DATA_VAL, BASE_VAL, 0) * _
            Excel.Application.Run(CALC_FUNC_STR, DATA_VAL, i, 1)
        End If
        TEMP_MATRIX(i, 1) = Excel.Application.Run(CALC_FIELD_STR, i)
    Next i
    
    ENERGY_SYSTEMS_CALC_FUNC = TEMP_MATRIX
'----------------------------------------------------------------------
Case Else
'----------------------------------------------------------------------
    ENERGY_SYSTEMS_CALC_FUNC = CStr("" & CALC_STR & " CALCULATOR: BASE <" & Excel.Application.Run(CALC_FIELD_STR, BASE_VAL) & ">")
End Select

Exit Function
ERROR_LABEL:
ENERGY_SYSTEMS_CALC_FUNC = Err.number
End Function

'-----------------------------------------------------------------------------------
'-------------------------LINKS TO CONVERSION FACTORS-------------------------------
'-----------------------------------------------------------------------------------

'Metric Conversion Factors
'http://www.eia.doe.gov/emeu/aer/pdf/pages/sec13_12.pdf

'Metric Prefixes and Miscellaneous Conversion Factors
'http://www.eia.doe.gov/emeu/aer/pdf/pages/sec13_13.pdf
'http://www.eia.doe.gov/neic/experts/heatcalc.xls
 
'Others
 'http://www.eppo.go.th/ref/UNIT-OIL.html
 'http://astro.berkeley.edu/~wright/fuel_energy.html
 'http://www.cngcorp.com/customer_sales_service/fuel_cost_charts.html
 'http://www.physics.uci.edu/~silverma/units.html
'-----------------------------------------------------------------------------------

'-----------------------------------------------------------------------------------------------------------------------------------------
'http://www.bp.com/productlanding.do?categoryId=6929&contentId=7044622
'www.eia.doe.gov/emeu/steo/pub/xls/chart-gallery.xls
'www.eia.doe.gov/emeu/steo/pub/xls/STEO_m.xls
'www.eia.doe.gov/emeu/steo/pub/fsheets/real_prices.xls
'www.eia.doe.gov/emeu/cabs/AOMC/images/chron_2008.xls
'www.eia.doe.gov/neic/experts/heatcalc.xls
'www.eia.doe.gov/oiaf/servicerpt/stimulus/excel/stimulus.xls
'www.eia.doe.gov/pub/oil_gas/natural_gas/feature_articles/2009/ngyir2008/ngyir2008.xls
'www.eia.doe.gov/emeu/international/Crude2.xls
'www.cmegroup.com/trading/energy/nymex-daily-reports.html
'www.google.ca/search?hl=en&client=firefox-a&rls=org.mozilla:en-US:official&hs=HSm&q=filetype:xls+site:www.cmegroup.com&start=0&sa=N
'-----------------------------------------------------------------------------------------------------------------------------------------
'http://www.gams.com/specialists/
'http://www.amsterdamoptimization.com/
'http://www.power.uwaterloo.ca/~fmilano/archive/marbella.pdf
'http://www.electricitycommission.govt.nz/opdev/modelling/gem
'http://www.erc.uct.ac.za/Research/Snapp/snapp.htm
'http://www.google.ca/search?hl=en&q=filetype:xlsm+site:http://www.chpcenternw.org/NwChpDocs/&aq=f&aqi=&aql=&oq=&gs_rfai=
'http://www.google.ca/search?hl=en&q=filetype:ppt+site:http://www.chpcenternw.org/NwChpDocs/&aq=f&aqi=&aql=&oq=&gs_rfai=
'-----------------------------------------------------------------------------------------------------------------------------------------

'-----------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------

'The 1999 Overture: Unit-of-measure mix-up tied to loss of $125Million
'Mars Orbiter NASA’s Mars Climate Orbiter was lost because engineers
'did not make a simple conversion from English units to metric, an
'embarrassing lapse that sent the $125 million craft off course.
'. . .. . . The navigators ( JPL ) assumed metric units
'of force per second, or newtons.  In fact, the numbers were in pounds
'of force per second as supplied ‘by Lockheed Martin ( the contractor ).
'Source: Kathy Sawyer, Boston Globe, October 1, 1999, page 1.

'-----------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------
