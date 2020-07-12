Attribute VB_Name = "NUMBER_REAL_CONSTANT_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.

'************************************************************************************
'************************************************************************************
'FUNCTION      : SYMBOLIC_CONSTANT_FUNC
'DESCRIPTION   : Translate a symbolic Constant to its double value
'LIBRARY       : NUMBER_REAL
'GROUP         : CONSTANT
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/08/2008
'************************************************************************************
'************************************************************************************

Function SYMBOLIC_CONSTANT_FUNC(ByRef CONSTANT_STR As String)
    
Dim PI_VAL As Double

On Error GoTo ERROR_LABEL

PI_VAL = 3.14159265358979

Select Case UCase(CONSTANT_STR)
Case UCase("pi")
    SYMBOLIC_CONSTANT_FUNC = PI_VAL              'pi-greek
Case UCase("pi2")
    SYMBOLIC_CONSTANT_FUNC = PI_VAL / 2                'pi-greek/2
Case UCase("pi3")
    SYMBOLIC_CONSTANT_FUNC = PI_VAL / 3                'pi-greek/3
Case UCase("pi4")
    SYMBOLIC_CONSTANT_FUNC = PI_VAL / 4                'pi-greek/4
Case UCase("e")
    SYMBOLIC_CONSTANT_FUNC = 2.71828182845905       'Euler-Napier constant
Case UCase("eu")
    SYMBOLIC_CONSTANT_FUNC = 0.577215664901533      'Euler-Mascheroni constant
Case UCase("phi")
    SYMBOLIC_CONSTANT_FUNC = 1.61803398874989       'golden ratio
Case UCase("g")
    SYMBOLIC_CONSTANT_FUNC = 9.80665                'Acceleration due to gravity
Case UCase("G")
    SYMBOLIC_CONSTANT_FUNC = 6.672 * 10 ^ -11       'Gravitational constant
Case UCase("R")
    SYMBOLIC_CONSTANT_FUNC = 8.31451                'Gas constant
Case UCase("eps")
    SYMBOLIC_CONSTANT_FUNC = 8.854187817 * 10 ^ -12 'Permittivity of vacuum
Case UCase("mu")
    SYMBOLIC_CONSTANT_FUNC = 12.566370614 * 10 ^ -7 'Permeability of vacuum
Case UCase("c")
    SYMBOLIC_CONSTANT_FUNC = 2.99792458 * 10 ^ 8    'Speed of light
Case UCase("q")
    SYMBOLIC_CONSTANT_FUNC = 1.60217733 * 10 ^ -19  'Elementary charge
Case UCase("me")
    SYMBOLIC_CONSTANT_FUNC = 9.1093897 * 10 ^ -31   'Electron rest mass
Case UCase("mp")
    SYMBOLIC_CONSTANT_FUNC = 1.6726231 * 10 ^ -27   'Proton rest mass
Case UCase("mn")
    SYMBOLIC_CONSTANT_FUNC = 1.6749286 * 10 ^ -27   'Neutron rest mass
Case UCase("K")
    SYMBOLIC_CONSTANT_FUNC = 1.380658 * 10 ^ -23    'Boltzmann constant
Case UCase("h")
    SYMBOLIC_CONSTANT_FUNC = 6.6260755 * 10 ^ -34   'Planck constant
Case UCase("A")
    SYMBOLIC_CONSTANT_FUNC = 6.0221367 * 10 ^ 23    'Avogadro number
Case Else ' support intrinsic date/time values
    Select Case UCase(CONSTANT_STR)
    Case "DATE"  'or date
        SYMBOLIC_CONSTANT_FUNC = CDbl(Date)
    Case "TIME"  'or time
        SYMBOLIC_CONSTANT_FUNC = CDbl(Time)
    Case "NOW"   'or now
        SYMBOLIC_CONSTANT_FUNC = CDbl(Now)
    Case Else
        GoTo ERROR_LABEL
    End Select
End Select

Exit Function
ERROR_LABEL:
SYMBOLIC_CONSTANT_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : TIME_CONSTANT_FUNC
'DESCRIPTION   : Convert to years (365.2564 days), days, hours, minutes ,seconds
'millisecond, microseconds, nanoseconds
'LIBRARY       : NUMBER_REAL
'GROUP         : CONSTANT
'ID            : 002
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/08/2008
'************************************************************************************
'************************************************************************************


Function TIME_CONSTANT_FUNC(ByVal DATA_VAL As Variant, _
ByVal BASE_VAL As Integer, _
Optional ByVal OUTPUT As Integer = 0)

Dim TEMP_VAL As Variant

On Error GoTo ERROR_LABEL

If BASE_VAL = 1 Then '"years (365.2564 days)"
    TEMP_VAL = 1 / 31558152.96
ElseIf BASE_VAL = 2 Then '"days"
    TEMP_VAL = 1 / 86400
ElseIf BASE_VAL = 3 Then '"hours"
    TEMP_VAL = 2.77777777777778E-04
ElseIf BASE_VAL = 4 Then '"minutes"
    TEMP_VAL = 1.66666666666667E-02
ElseIf BASE_VAL = 5 Then '"seconds"
    TEMP_VAL = 1
ElseIf BASE_VAL = 6 Then '"milliseconds"
    TEMP_VAL = 1000
ElseIf BASE_VAL = 7 Then '"microseconds"
    TEMP_VAL = 1000000
ElseIf BASE_VAL = 8 Then '"nanoseconds"
    TEMP_VAL = 1000000000
Else
    GoTo ERROR_LABEL
End If

Select Case OUTPUT
Case 0
    TIME_CONSTANT_FUNC = (DATA_VAL / TEMP_VAL)
Case Else
    TIME_CONSTANT_FUNC = TEMP_VAL
End Select

Exit Function
ERROR_LABEL:
TIME_CONSTANT_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : TIME_CONSTANT_FIELD_FUNC
'DESCRIPTION   : Time Calculator Tags
'LIBRARY       : NUMBER_REAL
'GROUP         : CONSTANT
'ID            : 003
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/08/2008
'************************************************************************************
'************************************************************************************

Function TIME_CONSTANT_FIELD_FUNC(ByVal FIELD_STR As String)

On Error GoTo ERROR_LABEL

If FIELD_STR = 1 Then
    TIME_CONSTANT_FIELD_FUNC = "years (365.2564 days)"
ElseIf FIELD_STR = 2 Then
    TIME_CONSTANT_FIELD_FUNC = "days"
ElseIf FIELD_STR = 3 Then
    TIME_CONSTANT_FIELD_FUNC = "hours"
ElseIf FIELD_STR = 4 Then
    TIME_CONSTANT_FIELD_FUNC = "minutes"
ElseIf FIELD_STR = 5 Then
    TIME_CONSTANT_FIELD_FUNC = "seconds"
ElseIf FIELD_STR = 6 Then
    TIME_CONSTANT_FIELD_FUNC = "milliseconds"
ElseIf FIELD_STR = 7 Then
    TIME_CONSTANT_FIELD_FUNC = "microseconds"
ElseIf FIELD_STR = 8 Then
    TIME_CONSTANT_FIELD_FUNC = "nanoseconds"
Else
    GoTo ERROR_LABEL
End If

Exit Function
ERROR_LABEL:
TIME_CONSTANT_FIELD_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : AREA_CONSTANT_FUNC
'DESCRIPTION   : Convert to sqmi, sqkm, hect, acre, sqm, sqyd, sqft,
'sqin, sqcm
'LIBRARY       : NUMBER_REAL
'GROUP         : CONSTANT
'ID            : 004
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/08/2008
'************************************************************************************
'************************************************************************************

Function AREA_CONSTANT_FUNC(ByVal DATA_VAL As Variant, _
ByVal BASE_VAL As Integer, _
Optional ByVal OUTPUT As Integer = 0)

Dim TEMP_VAL As Variant

On Error GoTo ERROR_LABEL

If BASE_VAL = 1 Then '"sqmi"
    TEMP_VAL = 3.8610215854185E-07
ElseIf BASE_VAL = 2 Then '"sqkm"
    TEMP_VAL = 0.000001
ElseIf BASE_VAL = 3 Then '"hect"
    TEMP_VAL = 0.0001
ElseIf BASE_VAL = 4 Then '"acre"
    TEMP_VAL = 2.47105381467165E-04
ElseIf BASE_VAL = 5 Then '"sqm"
    TEMP_VAL = 1 'SQUARED METER IS THE METRIC BASE
ElseIf BASE_VAL = 6 Then '"sqyd"
    TEMP_VAL = 1.19599004630108
ElseIf BASE_VAL = 7 Then '"sqft"
    TEMP_VAL = 10.7639104167097
ElseIf BASE_VAL = 8 Then '"sqin"
    TEMP_VAL = 1550.0031000062
ElseIf BASE_VAL = 9 Then '"sqcm"
    TEMP_VAL = 10000
Else
    GoTo ERROR_LABEL
End If

Select Case OUTPUT
    Case 0
        AREA_CONSTANT_FUNC = (DATA_VAL / TEMP_VAL)
    Case Else
        AREA_CONSTANT_FUNC = TEMP_VAL
End Select

Exit Function
ERROR_LABEL:
AREA_CONSTANT_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : AREA_CONSTANT_FIELD_FUNC
'DESCRIPTION   : Area Calculator Tags
'LIBRARY       : NUMBER_REAL
'GROUP         : CONSTANT
'ID            : 005
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/08/2008
'************************************************************************************
'************************************************************************************

Function AREA_CONSTANT_FIELD_FUNC(ByVal FIELD_STR As String)

On Error GoTo ERROR_LABEL

If FIELD_STR = 1 Then
    AREA_CONSTANT_FIELD_FUNC = "sqmi"
ElseIf FIELD_STR = 2 Then
    AREA_CONSTANT_FIELD_FUNC = "sqkm"
ElseIf FIELD_STR = 3 Then
    AREA_CONSTANT_FIELD_FUNC = "hect"
ElseIf FIELD_STR = 4 Then
    AREA_CONSTANT_FIELD_FUNC = "acre"
ElseIf FIELD_STR = 5 Then
    AREA_CONSTANT_FIELD_FUNC = "sqm"
ElseIf FIELD_STR = 6 Then
    AREA_CONSTANT_FIELD_FUNC = "sqyd"
ElseIf FIELD_STR = 7 Then
    AREA_CONSTANT_FIELD_FUNC = "sqft"
ElseIf FIELD_STR = 8 Then
    AREA_CONSTANT_FIELD_FUNC = "sqin"
ElseIf FIELD_STR = 9 Then
    AREA_CONSTANT_FIELD_FUNC = "sqcm"
Else
    GoTo ERROR_LABEL
End If

Exit Function
ERROR_LABEL:
AREA_CONSTANT_FIELD_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : LENGTH_CONSTANT_FUNC
'DESCRIPTION   : Conver to nautm, stmi, mi, km, chain, rod, fath, m, yd,
'survft ft, in, cm, mm
'LIBRARY       : NUMBER_REAL
'GROUP         : CONSTANT
'ID            : 006
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/08/2008
'************************************************************************************
'************************************************************************************

Function LENGTH_CONSTANT_FUNC(ByVal DATA_VAL As Variant, _
ByVal BASE_VAL As Integer, _
Optional ByVal OUTPUT As Integer = 0)

Dim TEMP_VAL As Variant

On Error GoTo ERROR_LABEL

If BASE_VAL = 1 Then '"nautm"
    TEMP_VAL = 5.39956803455723E-04
ElseIf BASE_VAL = 2 Then '"stmi"
    TEMP_VAL = 6.21369949494948E-04
ElseIf BASE_VAL = 3 Then '"mi"
    TEMP_VAL = 6.21371192237334E-04
ElseIf BASE_VAL = 4 Then '"km"
    TEMP_VAL = 0.001
ElseIf BASE_VAL = 5 Then '"chain"
    TEMP_VAL = 4.97095959595959E-02
ElseIf BASE_VAL = 6 Then '"rod"
    TEMP_VAL = 0.198838383838384
ElseIf BASE_VAL = 7 Then '"fath"
    TEMP_VAL = 0.546805555555557
ElseIf BASE_VAL = 8 Then '"m"
    TEMP_VAL = 1
ElseIf BASE_VAL = 9 Then '"yd"
    TEMP_VAL = 1.09361329833771
ElseIf BASE_VAL = 10 Then '"survft"
    TEMP_VAL = 3.28083333333334
ElseIf BASE_VAL = 11 Then '"ft"
    TEMP_VAL = 3.28083989501312
ElseIf BASE_VAL = 12 Then '"in"
    TEMP_VAL = 39.3700787401575
ElseIf BASE_VAL = 13 Then '"cm"
    TEMP_VAL = 100
ElseIf BASE_VAL = 14 Then '"mm"
    TEMP_VAL = 1000
Else
    GoTo ERROR_LABEL
End If

Select Case OUTPUT
    Case 0
        LENGTH_CONSTANT_FUNC = (DATA_VAL / TEMP_VAL)
    Case Else
        LENGTH_CONSTANT_FUNC = TEMP_VAL
End Select

Exit Function
ERROR_LABEL:
LENGTH_CONSTANT_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : LENGTH_CONSTANT_FIELD_FUNC
'DESCRIPTION   : Length Calculator BASE_VAL Tags
'LIBRARY       : NUMBER_REAL
'GROUP         : CONSTANT
'ID            : 007
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/08/2008
'************************************************************************************
'************************************************************************************

Function LENGTH_CONSTANT_FIELD_FUNC(ByVal FIELD_STR As String)

On Error GoTo ERROR_LABEL

If FIELD_STR = 1 Then
    LENGTH_CONSTANT_FIELD_FUNC = "nautm"
ElseIf FIELD_STR = 2 Then
    LENGTH_CONSTANT_FIELD_FUNC = "stmi"
ElseIf FIELD_STR = 3 Then
    LENGTH_CONSTANT_FIELD_FUNC = "mi"
ElseIf FIELD_STR = 4 Then
    LENGTH_CONSTANT_FIELD_FUNC = "km"
ElseIf FIELD_STR = 5 Then
    LENGTH_CONSTANT_FIELD_FUNC = "chain"
ElseIf FIELD_STR = 6 Then
    LENGTH_CONSTANT_FIELD_FUNC = "rod"
ElseIf FIELD_STR = 7 Then
    LENGTH_CONSTANT_FIELD_FUNC = "fath"
ElseIf FIELD_STR = 8 Then
    LENGTH_CONSTANT_FIELD_FUNC = "m"
ElseIf FIELD_STR = 9 Then
    LENGTH_CONSTANT_FIELD_FUNC = "yd"
ElseIf FIELD_STR = 10 Then
    LENGTH_CONSTANT_FIELD_FUNC = "survft"
ElseIf FIELD_STR = 11 Then
    LENGTH_CONSTANT_FIELD_FUNC = "ft"
ElseIf FIELD_STR = 12 Then
    LENGTH_CONSTANT_FIELD_FUNC = "in"
ElseIf FIELD_STR = 13 Then
    LENGTH_CONSTANT_FIELD_FUNC = "cm"
ElseIf FIELD_STR = 14 Then
    LENGTH_CONSTANT_FIELD_FUNC = "mm"
Else
    GoTo ERROR_LABEL
End If

Exit Function
ERROR_LABEL:
LENGTH_CONSTANT_FIELD_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : WEIGHT_CONSTANT_FUNC
'DESCRIPTION   : Convert to mton, ton, cwt, kg, lb, troz, oz, g, mg, µg
'LIBRARY       : NUMBER_REAL
'GROUP         : CONSTANT
'ID            : 008
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/08/2008
'************************************************************************************
'************************************************************************************


Function WEIGHT_CONSTANT_FUNC(ByVal DATA_VAL As Variant, _
ByVal BASE_VAL As Integer, _
Optional ByVal OUTPUT As Integer = 0)

Dim TEMP_VAL As Variant

On Error GoTo ERROR_LABEL

If BASE_VAL = 1 Then '"mton": t, tonne(s) --> 1000 kilograms
    TEMP_VAL = 1
ElseIf BASE_VAL = 2 Then '"ton": tn, ton(s) --> 2000 pounds, short ton.
    TEMP_VAL = 1.10231131092439
ElseIf BASE_VAL = 3 Then '"cwt"
    TEMP_VAL = 22.0462262184878
ElseIf BASE_VAL = 4 Then '"kg"
    TEMP_VAL = 1000
ElseIf BASE_VAL = 5 Then '"lb": lb, pound(s)  --> 16 ounces
    TEMP_VAL = 2204.62262184878
ElseIf BASE_VAL = 6 Then '"troz"
    TEMP_VAL = 32150.748429235
ElseIf BASE_VAL = 7 Then '"oz": oz, ounce(s) --> avoirdupois ounce (avdp oz), 1/16 pound.
    TEMP_VAL = 35273.9619495804
ElseIf BASE_VAL = 8 Then '"g": g, gram(s) --> Metric weight ~half a dime
    TEMP_VAL = 1000000
ElseIf BASE_VAL = 9 Then '"mg"
    TEMP_VAL = 1000000000
ElseIf BASE_VAL = 10 Then '"µg"
    TEMP_VAL = 1000000000000#
Else
    GoTo ERROR_LABEL
End If

Select Case OUTPUT
    Case 0
        WEIGHT_CONSTANT_FUNC = (DATA_VAL / TEMP_VAL)
    Case Else
        WEIGHT_CONSTANT_FUNC = TEMP_VAL
End Select

Exit Function
ERROR_LABEL:
WEIGHT_CONSTANT_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : WEIGHT_CONSTANT_FIELD_FUNC
'DESCRIPTION   : Weight Calculator BASE_VAL Tags
'LIBRARY       : NUMBER_REAL
'GROUP         : CONSTANT
'ID            : 009
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/08/2008
'************************************************************************************
'************************************************************************************

Function WEIGHT_CONSTANT_FIELD_FUNC(ByVal FIELD_STR As String)

On Error GoTo ERROR_LABEL


If FIELD_STR = 1 Then
    WEIGHT_CONSTANT_FIELD_FUNC = "mton"
ElseIf FIELD_STR = 2 Then
    WEIGHT_CONSTANT_FIELD_FUNC = "ton"
ElseIf FIELD_STR = 3 Then
    WEIGHT_CONSTANT_FIELD_FUNC = "cwt"
ElseIf FIELD_STR = 4 Then
    WEIGHT_CONSTANT_FIELD_FUNC = "kg"
ElseIf FIELD_STR = 5 Then
    WEIGHT_CONSTANT_FIELD_FUNC = "lb"
ElseIf FIELD_STR = 6 Then
    WEIGHT_CONSTANT_FIELD_FUNC = "troz"
ElseIf FIELD_STR = 7 Then
    WEIGHT_CONSTANT_FIELD_FUNC = "oz"
ElseIf FIELD_STR = 8 Then
    WEIGHT_CONSTANT_FIELD_FUNC = "g"
ElseIf FIELD_STR = 9 Then
    WEIGHT_CONSTANT_FIELD_FUNC = "mg"
ElseIf FIELD_STR = 10 Then
    WEIGHT_CONSTANT_FIELD_FUNC = "µg"
Else
    GoTo ERROR_LABEL
End If

Exit Function
ERROR_LABEL:
WEIGHT_CONSTANT_FIELD_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : MASS_CONSTANT_FUNC
'DESCRIPTION   : Convert to yotta, zetta, exa, peta, tera, giga, mega, kilo,
'hecto, deca, units, deci, centi, mili, micro, nano, pico, femto, atto,
'zepto, yocto
'LIBRARY       : NUMBER_REAL
'GROUP         : CONSTANT
'ID            : 010
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/08/2008
'************************************************************************************
'************************************************************************************

Function MASS_CONSTANT_FUNC(ByVal DATA_VAL As Variant, _
ByVal BASE_VAL As Integer, _
Optional ByVal OUTPUT As Integer = 0)

Dim TEMP_VAL As Variant

On Error GoTo ERROR_LABEL

If BASE_VAL = 1 Then '"yotta"
    TEMP_VAL = 10 ^ -24
ElseIf BASE_VAL = 2 Then '"zetta"
    TEMP_VAL = 10 ^ -21
ElseIf BASE_VAL = 3 Then '"exa"
    TEMP_VAL = 10 ^ -18
ElseIf BASE_VAL = 4 Then '"peta"
    TEMP_VAL = 10 ^ -15
ElseIf BASE_VAL = 5 Then '"tera"
    TEMP_VAL = 10 ^ -12
ElseIf BASE_VAL = 6 Then '"giga"
    TEMP_VAL = 10 ^ -9
ElseIf BASE_VAL = 7 Then '"mega"
    TEMP_VAL = 10 ^ -6
ElseIf BASE_VAL = 8 Then '"kilo"
    TEMP_VAL = 10 ^ -3
ElseIf BASE_VAL = 9 Then '"hecto"
    TEMP_VAL = 10 ^ -2
ElseIf BASE_VAL = 10 Then '"deca"
    TEMP_VAL = 10 ^ -1
ElseIf BASE_VAL = 11 Then '"unit"
    TEMP_VAL = 1
ElseIf BASE_VAL = 12 Then '"deci"
    TEMP_VAL = 10 ^ 1
ElseIf BASE_VAL = 13 Then '"centi"
    TEMP_VAL = 10 ^ 2
ElseIf BASE_VAL = 14 Then '"mili"
    TEMP_VAL = 10 ^ 3
ElseIf BASE_VAL = 15 Then '"micro"
    TEMP_VAL = 10 ^ 6
ElseIf BASE_VAL = 16 Then '"nano"
    TEMP_VAL = 10 ^ 9
ElseIf BASE_VAL = 17 Then '"pico"
    TEMP_VAL = 10 ^ 12
ElseIf BASE_VAL = 18 Then '"femto"
    TEMP_VAL = 10 ^ 15
ElseIf BASE_VAL = 19 Then '"atto"
    TEMP_VAL = 10 ^ 18
ElseIf BASE_VAL = 20 Then '"zepto"
    TEMP_VAL = 10 ^ 21
ElseIf BASE_VAL = 21 Then '"yocto"
    TEMP_VAL = 10 ^ 24
Else
    GoTo ERROR_LABEL
End If

Select Case OUTPUT
    Case 0
        MASS_CONSTANT_FUNC = (DATA_VAL / TEMP_VAL)
    Case Else
        MASS_CONSTANT_FUNC = TEMP_VAL
End Select

Exit Function
ERROR_LABEL:
MASS_CONSTANT_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : MASS_CONSTANT_FIELD_FUNC
'DESCRIPTION   : Mass Calculator BASE_VAL Tags
'LIBRARY       : NUMBER_REAL
'GROUP         : CONSTANT
'ID            : 011
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/08/2008
'************************************************************************************
'************************************************************************************

Function MASS_CONSTANT_FIELD_FUNC(ByVal FIELD_STR As String)

On Error GoTo ERROR_LABEL

If FIELD_STR = 1 Then
    MASS_CONSTANT_FIELD_FUNC = "yotta"
ElseIf FIELD_STR = 2 Then
    MASS_CONSTANT_FIELD_FUNC = "zetta"
ElseIf FIELD_STR = 3 Then
    MASS_CONSTANT_FIELD_FUNC = "exa"
ElseIf FIELD_STR = 4 Then
    MASS_CONSTANT_FIELD_FUNC = "peta"
ElseIf FIELD_STR = 5 Then
    MASS_CONSTANT_FIELD_FUNC = "tera"
ElseIf FIELD_STR = 6 Then
    MASS_CONSTANT_FIELD_FUNC = "giga"
ElseIf FIELD_STR = 7 Then
    MASS_CONSTANT_FIELD_FUNC = "mega"
ElseIf FIELD_STR = 8 Then
    MASS_CONSTANT_FIELD_FUNC = "kilo"
ElseIf FIELD_STR = 9 Then
    MASS_CONSTANT_FIELD_FUNC = "hecto"
ElseIf FIELD_STR = 10 Then
    MASS_CONSTANT_FIELD_FUNC = "deca"
ElseIf FIELD_STR = 11 Then
    MASS_CONSTANT_FIELD_FUNC = "unit"
ElseIf FIELD_STR = 12 Then
    MASS_CONSTANT_FIELD_FUNC = "deci"
ElseIf FIELD_STR = 13 Then
    MASS_CONSTANT_FIELD_FUNC = "centi"
ElseIf FIELD_STR = 14 Then
    MASS_CONSTANT_FIELD_FUNC = "mili"
ElseIf FIELD_STR = 15 Then
    MASS_CONSTANT_FIELD_FUNC = "micro"
ElseIf FIELD_STR = 16 Then
    MASS_CONSTANT_FIELD_FUNC = "nano"
ElseIf FIELD_STR = 17 Then
    MASS_CONSTANT_FIELD_FUNC = "pico"
ElseIf FIELD_STR = 18 Then
    MASS_CONSTANT_FIELD_FUNC = "femto"
ElseIf FIELD_STR = 19 Then
    MASS_CONSTANT_FIELD_FUNC = "atto"
ElseIf FIELD_STR = 20 Then
    MASS_CONSTANT_FIELD_FUNC = "zepto"
ElseIf FIELD_STR = 21 Then
    MASS_CONSTANT_FIELD_FUNC = "yocto"
Else
    GoTo ERROR_LABEL
End If

Exit Function
ERROR_LABEL:
MASS_CONSTANT_FIELD_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : TEMPERATURE_CONSTANT_FUNC
'DESCRIPTION   : Convert to °c, °f, °k, °r
'LIBRARY       : NUMBER_REAL
'GROUP         : CONSTANT
'ID            : 012
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/08/2008
'************************************************************************************
'************************************************************************************


Function TEMPERATURE_CONSTANT_FUNC(ByVal DATA_VAL As Variant, _
ByVal BASE_VAL As Integer, _
Optional ByVal OUTPUT As Integer = 0)

Dim ATEMP_BASE As Variant
Dim BTEMP_BASE As Variant

On Error GoTo ERROR_LABEL

If BASE_VAL = 1 Then '"°c"
    ATEMP_BASE = 1
    BTEMP_BASE = 0
ElseIf BASE_VAL = 2 Then '"°f"
    ATEMP_BASE = 1.8
    BTEMP_BASE = 32
ElseIf BASE_VAL = 3 Then '"°k"
    ATEMP_BASE = 1
    BTEMP_BASE = 273.166666666667
ElseIf BASE_VAL = 4 Then '"°r"
    ATEMP_BASE = 1.8
    BTEMP_BASE = 491.7
Else
    GoTo ERROR_LABEL
End If

Select Case OUTPUT
    Case 0
        TEMPERATURE_CONSTANT_FUNC = ((DATA_VAL - BTEMP_BASE) / ATEMP_BASE)
    Case 1
        TEMPERATURE_CONSTANT_FUNC = ATEMP_BASE
    Case Else
        TEMPERATURE_CONSTANT_FUNC = BTEMP_BASE
End Select

Exit Function
ERROR_LABEL:
TEMPERATURE_CONSTANT_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : TEMPERATURE_CONSTANT_FIELD_FUNC
'DESCRIPTION   : Temperature BASE_VAL Tags
'LIBRARY       : NUMBER_REAL
'GROUP         : CONSTANT
'ID            : 013
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/08/2008
'************************************************************************************
'************************************************************************************

Function TEMPERATURE_CONSTANT_FIELD_FUNC(ByVal FIELD_STR As String)

On Error GoTo ERROR_LABEL

If FIELD_STR = 1 Then
    TEMPERATURE_CONSTANT_FIELD_FUNC = "°c"
ElseIf FIELD_STR = 2 Then
    TEMPERATURE_CONSTANT_FIELD_FUNC = "°f"
ElseIf FIELD_STR = 3 Then
    TEMPERATURE_CONSTANT_FIELD_FUNC = "°k"
ElseIf FIELD_STR = 4 Then
    TEMPERATURE_CONSTANT_FIELD_FUNC = "°r"
Else
    GoTo ERROR_LABEL
End If

Exit Function
ERROR_LABEL:
TEMPERATURE_CONSTANT_FIELD_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : VOLUME_CONSTANT_FUNC
'DESCRIPTION   : Convert to cum, cuyd, bushl, cuft, peck, igal, gal, ltr
'quart, pint, cup, floz, cuin, tbsp, tsp, ml
'LIBRARY       : NUMBER_REAL
'GROUP         : CONSTANT
'ID            : 014
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/08/2008
'************************************************************************************
'************************************************************************************


Function VOLUME_CONSTANT_FUNC(ByVal DATA_VAL As Variant, _
ByVal BASE_VAL As Integer, _
Optional ByVal OUTPUT As Integer = 0)

Dim TEMP_VAL As Variant

On Error GoTo ERROR_LABEL

If BASE_VAL = 1 Then '"cord": cord(s) of wood --> 128 cubic feet
    TEMP_VAL = 0.275895800556912
ElseIf BASE_VAL = 2 Then '"cum": cubic meter(s) --> 31% more than a cubic yard.
    TEMP_VAL = 1
ElseIf BASE_VAL = 3 Then '"cuyd"
    TEMP_VAL = 1.30795061931439
ElseIf BASE_VAL = 4 Then '"bbl":  barrel(s) 42 american gallons. This is the
'international oil barrel (bbl or bo). Commerical barrels in the US are
'officially 31.5 gallons, though beer barrels are usually 31 gallons,
'and there are many other variations.
    TEMP_VAL = 6.28981077043211
ElseIf BASE_VAL = 5 Then '"bushl"
    TEMP_VAL = 28.3775933927882
ElseIf BASE_VAL = 6 Then '"cuft"
    TEMP_VAL = 35.3146667214886
ElseIf BASE_VAL = 7 Then '"peck"
    TEMP_VAL = 113.510373571153
ElseIf BASE_VAL = 8 Then '"igal": Brittish gallon, larger than a US gallon.
    TEMP_VAL = 219.96915152619
ElseIf BASE_VAL = 9 Then '"gal": American gallon
    TEMP_VAL = 264.172052358148
ElseIf BASE_VAL = 10 Then '"ltr"
    TEMP_VAL = 1000
ElseIf BASE_VAL = 11 Then '"quart"
    TEMP_VAL = 1056.68820943259
ElseIf BASE_VAL = 12 Then '"pint"
    TEMP_VAL = 2113.37641886519
ElseIf BASE_VAL = 13 Then '"cup"
    TEMP_VAL = 4226.75283773038
ElseIf BASE_VAL = 14 Then '"jig": jigger(s) --> 1.5 fluid ounces <fl oz US>
    TEMP_VAL = 22542.6807054818
ElseIf BASE_VAL = 15 Then '"floz-us": there are 32 US fluid ounces in a US quart.
    TEMP_VAL = 33814.0210582226
ElseIf BASE_VAL = 16 Then '"floz-uk": there are 20 imperial ounces in an imperial pint.
    TEMP_VAL = 35195.0791085072
ElseIf BASE_VAL = 17 Then '"cuin"
    TEMP_VAL = 61023.7440947323
ElseIf BASE_VAL = 18 Then '"tbsp"
    TEMP_VAL = 67628.045403686
ElseIf BASE_VAL = 19 Then '"tsp"
        TEMP_VAL = 202884.136211058
ElseIf BASE_VAL = 20 Then '"ml"
    TEMP_VAL = 1000000
Else
    GoTo ERROR_LABEL
End If

Select Case OUTPUT
    Case 0
        VOLUME_CONSTANT_FUNC = (DATA_VAL / TEMP_VAL)
    Case Else
        VOLUME_CONSTANT_FUNC = TEMP_VAL
End Select

Exit Function
ERROR_LABEL:
VOLUME_CONSTANT_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : VOLUME_CONSTANT_FIELD_FUNC
'DESCRIPTION   : Volume Calculator BASE_VAL Tags
'LIBRARY       : NUMBER_REAL
'GROUP         : CONSTANT
'ID            : 015
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/08/2008
'************************************************************************************
'************************************************************************************

Function VOLUME_CONSTANT_FIELD_FUNC(ByVal FIELD_STR As String)

On Error GoTo ERROR_LABEL

If FIELD_STR = 1 Then
    VOLUME_CONSTANT_FIELD_FUNC = "cord"
ElseIf FIELD_STR = 2 Then
    VOLUME_CONSTANT_FIELD_FUNC = "cum"
ElseIf FIELD_STR = 3 Then
    VOLUME_CONSTANT_FIELD_FUNC = "cuyd"
ElseIf FIELD_STR = 4 Then
    VOLUME_CONSTANT_FIELD_FUNC = "bbl"
ElseIf FIELD_STR = 5 Then
    VOLUME_CONSTANT_FIELD_FUNC = "bushl"
ElseIf FIELD_STR = 6 Then
    VOLUME_CONSTANT_FIELD_FUNC = "cuft"
ElseIf FIELD_STR = 7 Then
    VOLUME_CONSTANT_FIELD_FUNC = "peck"
ElseIf FIELD_STR = 8 Then
    VOLUME_CONSTANT_FIELD_FUNC = "igal"
ElseIf FIELD_STR = 9 Then
    VOLUME_CONSTANT_FIELD_FUNC = "gal"
ElseIf FIELD_STR = 10 Then
    VOLUME_CONSTANT_FIELD_FUNC = "ltr"
ElseIf FIELD_STR = 11 Then
    VOLUME_CONSTANT_FIELD_FUNC = "quart"
ElseIf FIELD_STR = 12 Then
    VOLUME_CONSTANT_FIELD_FUNC = "pint"
ElseIf FIELD_STR = 13 Then
    VOLUME_CONSTANT_FIELD_FUNC = "cup"
ElseIf FIELD_STR = 14 Then
    VOLUME_CONSTANT_FIELD_FUNC = "jig"
ElseIf FIELD_STR = 15 Then
    VOLUME_CONSTANT_FIELD_FUNC = "floz-us"
ElseIf FIELD_STR = 16 Then
    VOLUME_CONSTANT_FIELD_FUNC = "floz-uk"
ElseIf FIELD_STR = 17 Then
    VOLUME_CONSTANT_FIELD_FUNC = "cuin"
ElseIf FIELD_STR = 18 Then
    VOLUME_CONSTANT_FIELD_FUNC = "tbsp"
ElseIf FIELD_STR = 19 Then
    VOLUME_CONSTANT_FIELD_FUNC = "tsp"
ElseIf FIELD_STR = 20 Then
    VOLUME_CONSTANT_FIELD_FUNC = "ml"
Else
    GoTo ERROR_LABEL
End If

Exit Function
ERROR_LABEL:
    VOLUME_CONSTANT_FIELD_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : ENERGY_CONSTANT_FUNC
'DESCRIPTION   : Conver to J, ft-lb, Cal,Btu, Wh, kCal, MJ, hp-hr, therm,
'BOE, TCE, TOE, MTon, nuke, Quad
'LIBRARY       : NUMBER_REAL
'GROUP         : CONSTANT
'ID            : 016
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/08/2008
'************************************************************************************
'************************************************************************************

Function ENERGY_CONSTANT_FUNC(ByVal DATA_VAL As Variant, _
ByVal BASE_VAL As Integer, _
Optional ByVal OUTPUT As Integer = 0)

Dim TEMP_VAL As Variant

On Error GoTo ERROR_LABEL

If BASE_VAL = 1 Then '"J": J, joule(s)-->One watt second.
'The energy used by a 1 watt light bulb (e.g. ~ for a flashlight) in 1 second.
'A joule is a watt-second, or 1/1055 Btu
    TEMP_VAL = 1 / 1

ElseIf BASE_VAL = 2 Then '"ft -lb" : ft-lb, foot-pound(s)-->The energy required
'to raise 1 pound by 1 foot.
    TEMP_VAL = 1 / 1.3558179

ElseIf BASE_VAL = 3 Then '"Cal" : cal, calorie(s)-->The energy needed to raise
'the temperagure of 1 gram of water 1 degree Centigrade.
    TEMP_VAL = 1 / 4.1868

ElseIf BASE_VAL = 4 Then '"Btu" : Btu, British thermal unit(s)-->A Btu (or BTU)
'is a British thermal unit, the energy needed to heat 1 pound of water 1
'degree Fahrenheit.
    TEMP_VAL = 1 / 1055.05585262
    'According to http://www.unc.edu/~rowlett/units/
    'British thermal unit = 1055.055895 J
ElseIf BASE_VAL = 5 Then '"Wh" : Wh, watt-hour(s)-->One watt for one hour. The
'energy needed to light one 60 watt light bulb for one minute.
    TEMP_VAL = 1 / 3600

ElseIf BASE_VAL = 6 Then '"kCal" : kcal, food calorie(s)-->The energy need to
'heat 1 kilogram (2.2 pounds) of water 1 deg. C
    TEMP_VAL = 1 / 4186.8

ElseIf BASE_VAL = 7 Then '"MJ" : MJ, mega joule(s)-->One million joules.
    TEMP_VAL = 1 / 1000000

ElseIf BASE_VAL = 8 Then '"hp -hr" : hp hr, horsepower hours(s)-->The energy
'provided by a horse pulling a load for one hour.
    TEMP_VAL = 1 / 2684519 '--> 3600 * 745.7 Watt (745.7 Watt = 1 HorsePower)

ElseIf BASE_VAL = 9 Then '"therm" : thm, therm(s)-->100,000 Btu (British
'thermal units).
    TEMP_VAL = 1 / 105505585.262

ElseIf BASE_VAL = 10 Then '"BOE" : BOE, barrel(s) of oil equivalent-->A BOE
'(also boe, bboe) is a "barrels of oil equivalent."
    TEMP_VAL = 1 / 5711869031.3779
    'Also remember that 1 BOE Barrels of oil equivalent = 5,658.53 ft^3 NG
    'http://www.spe.org/spe/jsp/basic/0,,1104_1732,00.html
ElseIf BASE_VAL = 11 Then '"TCE" : TCE, tonne(s) of coal equivalent-->Tonnes
'(metric tons) of coal equivalent.
    TEMP_VAL = 1 / 29307600000#

ElseIf BASE_VAL = 12 Then '"TOE" : TOE, tonne(s) of oil equivalent-->Tonnes
'(metric tons) of oil equivalent.
    TEMP_VAL = 1 / 41868000000# 'http://www.iea.org/dbtw-wpd/Textbase/stats/unit.asp
    'but according to http://unstats.un.org/unsd/energy/balance/conversion.htm
    '1 Tonne of oil equivalent = 42.6216*(10^9) = 42621600000
ElseIf BASE_VAL = 13 Then '"Mton": Mton, megaton(s)-->Megatons (Mton or Mt) are
'used for measuring yieds of A- and H-bombs, e.g. a 1 megaton bomb.
    TEMP_VAL = 1 / 4.184E+15

ElseIf BASE_VAL = 14 Then '"nuke": N-yr, Nuclear-plant-year equivalent-->The output
'of one roughly typical (1+ GW) nuclear plant for one year. Defined as 1GW for
'8760 hours.
    TEMP_VAL = 1 / 3.1536E+16

ElseIf BASE_VAL = 15 Then '"Quad": Quad, quadrillion Btu-->One quadrillion
'(10<sup>15</sup>) Btu.
    TEMP_VAL = 1 / 1.05505585262E+18
Else
    GoTo ERROR_LABEL
End If

Select Case OUTPUT
    Case 0
ENERGY_CONSTANT_FUNC = (DATA_VAL / TEMP_VAL)
    Case Else
ENERGY_CONSTANT_FUNC = TEMP_VAL
End Select

Exit Function
ERROR_LABEL:
ENERGY_CONSTANT_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : ENERGY_CONSTANT_FIELD_FUNC
'DESCRIPTION   : Energy Calculator BASE_VAL Tags
'LIBRARY       : NUMBER_REAL
'GROUP         : CONSTANT
'ID            : 017
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/08/2008
'************************************************************************************
'************************************************************************************

Function ENERGY_CONSTANT_FIELD_FUNC(ByVal FIELD_STR As String)

On Error GoTo ERROR_LABEL

'-----------------------------------------------------------------------------------
'--------------------------------KEY NOTE-------------------------------------------
'MM --> MM, million --> 1e6 --> One million (6 zeros). This is not used with
'electricity or in the metric system. MM was meant to indicate one thousand
'thousand, M being the Roman numeral 1000. However, MM actually means 2000,
'not one million, in Roman numeration. Still used in some traditional units
'such as MMBtu and MMb (million barrels of oil).
'-----------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------

If FIELD_STR = 1 Then
    ENERGY_CONSTANT_FIELD_FUNC = "J"
ElseIf FIELD_STR = 2 Then
    ENERGY_CONSTANT_FIELD_FUNC = "ft -lb"
ElseIf FIELD_STR = 3 Then
    ENERGY_CONSTANT_FIELD_FUNC = "Cal"
ElseIf FIELD_STR = 4 Then
    ENERGY_CONSTANT_FIELD_FUNC = "Btu"
ElseIf FIELD_STR = 5 Then
    ENERGY_CONSTANT_FIELD_FUNC = "Wh"
ElseIf FIELD_STR = 6 Then
    ENERGY_CONSTANT_FIELD_FUNC = "kCal"
ElseIf FIELD_STR = 7 Then
    ENERGY_CONSTANT_FIELD_FUNC = "MJ"
ElseIf FIELD_STR = 8 Then
    ENERGY_CONSTANT_FIELD_FUNC = "hp -hr"
ElseIf FIELD_STR = 9 Then
    ENERGY_CONSTANT_FIELD_FUNC = "therm"
ElseIf FIELD_STR = 10 Then
    ENERGY_CONSTANT_FIELD_FUNC = "BOE"
ElseIf FIELD_STR = 11 Then
    ENERGY_CONSTANT_FIELD_FUNC = "TCE"
ElseIf FIELD_STR = 12 Then
    ENERGY_CONSTANT_FIELD_FUNC = "TOE"
ElseIf FIELD_STR = 13 Then
    ENERGY_CONSTANT_FIELD_FUNC = "Mton"
ElseIf FIELD_STR = 14 Then
    ENERGY_CONSTANT_FIELD_FUNC = "nuke"
ElseIf FIELD_STR = 15 Then
    ENERGY_CONSTANT_FIELD_FUNC = "Quad"
Else
    GoTo ERROR_LABEL
End If

Exit Function
ERROR_LABEL:
ENERGY_CONSTANT_FIELD_FUNC = Err.number
End Function


'-----------------------------------------------------------------------------------
'----------------------LINKS TO CONVERSION FACTORS----------------------------
'-----------------------------------------------------------------------------------

'Metric Conversion Factors
'http://www.eia.doe.gov/emeu/aer/pdf/pages/sec13_12.pdf

'Metric Prefixes and Miscellaneous Conversion Factors
'http://www.eia.doe.gov/emeu/aer/pdf/pages/sec13_13.pdf
 
'Others
 'http://www.eppo.go.th/ref/UNIT-OIL.html
 'http://astro.berkeley.edu/~wright/fuel_energy.html
 'http://www.cngcorp.com/customer_sales_service/fuel_cost_charts.html
 'http://www.physics.uci.edu/~silverma/units.html

'-----------------------------------------------------------------------------------
'--------------------------------KEY NOTE-------------------------------------------
'MM --> MM, million --> 1e6 --> One million (6 zeros). This is not used with
'electricity or in the metric system. MM was meant to indicate one thousand
'thousand, M being the Roman numeral 1000. However, MM actually means 2000,
'not one million, in Roman numeration. Still used in some traditional units
'such as MMBtu and MMb (million barrels of oil).
'-----------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------