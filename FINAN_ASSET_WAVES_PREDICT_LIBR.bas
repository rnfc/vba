Attribute VB_Name = "FINAN_ASSET_WAVES_PREDICT_LIBR"


'------------------------------------------------------------------------------------
Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.
'------------------------------------------------------------------------------------
Private Const PUB_EPSILON As Double = 2 ^ 52
Private PUB_DATA_MATRIX As Variant
'------------------------------------------------------------------------------------


'************************************************************************************
'************************************************************************************
'FUNCTION      : ASSET_WAVE_THEORY_PREDICT_FUNC
'DESCRIPTION   : Asset Waves Function
'LIBRARY       : FINAN_ASSET
'GROUP         : WAVES
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'REFERENCE: http://www.gummy-stuff.org/Wave_Theory.htm
'************************************************************************************
'************************************************************************************

Function ASSET_WAVE_THEORY_PREDICT_FUNC(ByVal TICKER_STR As Variant, _
Optional ByVal START_DATE As Date, _
Optional ByVal END_DATE As Date, _
Optional ByVal EXTRAPOLATE_PERIODS As Long = 2, _
Optional ByVal PRICE_AVG_START_PERIOD As Long = 3, _
Optional ByVal P0_AVG_START_PERIOD As Long = 3, _
Optional ByVal TN_PERIODS As Double = 3, _
Optional ByVal OUTPUT As Integer = 3)

'EXTRAPOLATE_PERIODS: Extrapolate last x periods
'P0_AVG_START_PERIOD: T START

Dim i As Long
Dim j As Long
Dim k As Long

Dim NROWS As Long

Dim PI_VAL As Double

Dim TEMP1_SUM As Double
Dim TEMP2_SUM As Double
Dim TEMP3_SUM As Double

Dim NEXT_PRICE_VAL As Double

Dim ERROR_VAL As Double
Dim UPS_DOWNS_VAL As Double

Dim FACTOR_VAL As Variant
Dim TEMP_MATRIX As Variant
Dim DATA_MATRIX As Variant

On Error GoTo ERROR_LABEL

If IsArray(TICKER_STR) = False Then
    DATA_MATRIX = YAHOO_HISTORICAL_DATA_SERIE_FUNC(TICKER_STR, START_DATE, END_DATE, "d", "DA", False, False, True)
Else
    DATA_MATRIX = TICKER_STR
End If
NROWS = UBound(DATA_MATRIX, 1)

If OUTPUT > 2 Then
    PUB_DATA_MATRIX = DATA_MATRIX
    ReDim PARAM_VECTOR(1 To 4, 1 To 1)
    PARAM_VECTOR(1, 1) = EXTRAPOLATE_PERIODS
    PARAM_VECTOR(2, 1) = PRICE_AVG_START_PERIOD
    PARAM_VECTOR(3, 1) = P0_AVG_START_PERIOD
    PARAM_VECTOR(4, 1) = TN_PERIODS
    
    PARAM_VECTOR = NELDER_MEAD_OPTIMIZATION3_FUNC( _
            "ASSET_WAVE_THEORY_OBJ_FUNC", _
            PARAM_VECTOR, 1000, 10 ^ -5)
            
    Erase PUB_DATA_MATRIX
    ASSET_WAVE_THEORY_PREDICT_FUNC = PARAM_VECTOR
    Exit Function
End If

PI_VAL = 3.14159265358979
FACTOR_VAL = 2 - (2 * PI_VAL / TN_PERIODS) ^ 2
'2-(2*p/T)^2; best: (2*p/T)^2

ReDim TEMP_MATRIX(0 To NROWS, 1 To 9)

TEMP_MATRIX(0, 1) = "DATE"
TEMP_MATRIX(0, 2) = "ACTUAL_PRICES"
TEMP_MATRIX(0, 3) = "P(N): " & PRICE_AVG_START_PERIOD & " - PERIODS PRICE AVG"
TEMP_MATRIX(0, 4) = "PO: " & P0_AVG_START_PERIOD & " AVG OF P(N)"
TEMP_MATRIX(0, 6) = "EXTRAPOLATION"
TEMP_MATRIX(0, 7) = "AVG PO"
TEMP_MATRIX(0, 9) = "GOOD PREDICTION: P(N)-P(N+1)"

k = 0
TEMP1_SUM = 0: TEMP2_SUM = 0
TEMP3_SUM = 0: UPS_DOWNS_VAL = 0
ERROR_VAL = 0
For i = 1 To NROWS
    TEMP_MATRIX(i, 1) = DATA_MATRIX(i, 1)
    TEMP_MATRIX(i, 2) = DATA_MATRIX(i, 2)

    If i <= PRICE_AVG_START_PERIOD Then
        TEMP1_SUM = TEMP1_SUM + TEMP_MATRIX(i, 2)
        TEMP_MATRIX(i, 3) = TEMP1_SUM / i
    Else
        j = i - PRICE_AVG_START_PERIOD
        TEMP1_SUM = TEMP1_SUM - TEMP_MATRIX(j, 2)
        TEMP1_SUM = TEMP1_SUM + TEMP_MATRIX(i, 2)
        TEMP_MATRIX(i, 3) = TEMP1_SUM / PRICE_AVG_START_PERIOD
    End If
    
    If i < P0_AVG_START_PERIOD Then
        TEMP2_SUM = TEMP2_SUM + TEMP_MATRIX(i, 3)
        TEMP_MATRIX(i, 4) = TEMP2_SUM / i
        
        TEMP_MATRIX(i, 5) = TEMP_MATRIX(i, 3)
        TEMP3_SUM = TEMP3_SUM + TEMP_MATRIX(i, 5)
        
        TEMP_MATRIX(i, 6) = ""
        TEMP_MATRIX(i, 7) = TEMP_MATRIX(i, 4)
        TEMP_MATRIX(i, 8) = ""
        TEMP_MATRIX(i, 9) = ""
    
    Else
        
        j = i - P0_AVG_START_PERIOD
        If i <> P0_AVG_START_PERIOD Then
            TEMP2_SUM = TEMP2_SUM - TEMP_MATRIX(j, 3)
            TEMP3_SUM = TEMP3_SUM - TEMP_MATRIX(j, 5)
        End If
        
        TEMP2_SUM = TEMP2_SUM + TEMP_MATRIX(i, 3)
        TEMP_MATRIX(i, 4) = TEMP2_SUM / P0_AVG_START_PERIOD
    
        If i < (NROWS + 1 - EXTRAPOLATE_PERIODS) Then
            TEMP_MATRIX(i, 5) = (2 - FACTOR_VAL) * _
                                TEMP_MATRIX(i - 1, 3) - _
                                TEMP_MATRIX(i - 2, 3) + _
                                FACTOR_VAL * TEMP_MATRIX(i - 1, 4)
        
            TEMP3_SUM = TEMP3_SUM + TEMP_MATRIX(i, 5)

            TEMP_MATRIX(i, 6) = ""
            TEMP_MATRIX(i, 7) = TEMP_MATRIX(i, 4)
        
        Else
            TEMP_MATRIX(i, 5) = (2 - FACTOR_VAL) * _
                                TEMP_MATRIX(i - 1, 5) - _
                                TEMP_MATRIX(i - 2, 5) + _
                                FACTOR_VAL * TEMP_MATRIX(i - 1, 7)
            TEMP_MATRIX(i, 6) = TEMP_MATRIX(i, 5)
            TEMP3_SUM = TEMP3_SUM + TEMP_MATRIX(i, 5)
                
            TEMP_MATRIX(i, 7) = TEMP3_SUM / P0_AVG_START_PERIOD
        End If
    
        If ((TEMP_MATRIX(i, 3) - TEMP_MATRIX(i - 1, 3)) * _
            (TEMP_MATRIX(i, 5) - TEMP_MATRIX(i - 1, 5)) > 0) Then
            
            TEMP_MATRIX(i, 8) = 1
            TEMP_MATRIX(i, 9) = TEMP_MATRIX(i, 5)
        Else
            TEMP_MATRIX(i, 8) = 0
            TEMP_MATRIX(i, 9) = ""
        End If
        UPS_DOWNS_VAL = UPS_DOWNS_VAL + TEMP_MATRIX(i, 8)
        k = k + 1
    End If
    If i < NROWS Then: ERROR_VAL = ERROR_VAL + Abs(DATA_MATRIX(i + 1, 2) - TEMP_MATRIX(i, 5)) ^ 2

'    If i >= UPS_DOWNS_START_PERIODS Then
'        UPS_DOWNS_VAL = UPS_DOWNS_VAL + TEMP_MATRIX(i, 8)
'        k = k + 1
'    End If
    
'    If i < NROWS Then: ERROR_VAL = ERROR_VAL + (TEMP_MATRIX(i, 5) - DATA_MATRIX(i + 1, 2))

Next i
    
UPS_DOWNS_VAL = UPS_DOWNS_VAL / k
ERROR_VAL = (ERROR_VAL / (NROWS - 1)) ^ 0.5
NEXT_PRICE_VAL = TEMP_MATRIX(NROWS, 5)

TEMP_MATRIX(0, 5) = "PREDICTED P(N+1) - RMS ERROR: " & Format(ERROR_VAL, "0.00")
TEMP_MATRIX(0, 8) = "UPS/DOWNS: " & Format(UPS_DOWNS_VAL, "0.0%")

Select Case OUTPUT
Case 0
    ASSET_WAVE_THEORY_PREDICT_FUNC = TEMP_MATRIX
Case 1
    ASSET_WAVE_THEORY_PREDICT_FUNC = Array(NEXT_PRICE_VAL, UPS_DOWNS_VAL, ERROR_VAL)
Case 2
    ASSET_WAVE_THEORY_PREDICT_FUNC = ERROR_VAL 'minimize
    'UPS_DOWNS_VAL '--> maximize
End Select

Exit Function
ERROR_LABEL:
ASSET_WAVE_THEORY_PREDICT_FUNC = PUB_EPSILON
End Function


Function ASSET_WAVE_THEORY_OBJ_FUNC(ByVal PARAM_VECTOR As Variant)

On Error GoTo ERROR_LABEL

ASSET_WAVE_THEORY_OBJ_FUNC = ASSET_WAVE_THEORY_PREDICT_FUNC(PUB_DATA_MATRIX, , , _
                                    PARAM_VECTOR(1, 1), _
                                    PARAM_VECTOR(2, 1), _
                                    PARAM_VECTOR(3, 1), _
                                    PARAM_VECTOR(4, 1), _
                                    2)
Exit Function
ERROR_LABEL:
ASSET_WAVE_THEORY_OBJ_FUNC = PUB_EPSILON
End Function
