Attribute VB_Name = "FINAN_DERIV_BS_SURFACE_LIBR"

'-------------------------------------------------------------------------------------
Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.
'-------------------------------------------------------------------------------------

Private PUB_TARGET_VAL As Double

Private PUB_SPOT_VAL As Double
Private PUB_RATE_VAL As Double
Private PUB_SIGMA_VAL As Double 'DefaultVolatility

Private PUB_STRIKE_VAL As Double
Private PUB_EXPIRATION_VAL As Double
Private PUB_OPTION_FLAG As Integer

Private PUB_CND_TYPE As Integer

Private PUB_TENOR_ARR As Variant
Private PUB_STRIKE_ARR As Variant

Private PUB_IMPL_BID_ARR As Variant
Private PUB_IMPL_ASK_ARR As Variant
Private PUB_IMPL_SIGMA_ARR As Variant

Private PUB_3D_MATRIX As Variant
Private PUB_INDEX_ARR As Variant
Private PUB_GROUP_ARR As Variant

Private Const PUB_EPSILON As Double = 2 ^ 52


'************************************************************************************
'************************************************************************************
'FUNCTION      : IMPLIED_VOLATILITY_SURFACE_CBOE_FUNC
'DESCRIPTION   : CBOE Implied Volatility Surface
'LIBRARY       : DERIVATIVES
'GROUP         : IMPLIED
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 21/10/2013
'http://www.cboe.com/DelayedQuote/QuoteTableDownload.aspx
'************************************************************************************
'************************************************************************************

Function IMPLIED_VOLATILITY_SURFACE_CBOE_FUNC(ByVal FILE_PATH_STR As String, _
Optional ByVal RATE As Double = 0.04, _
Optional ByVal FILTERING_EPS As Double = 0.01, _
Optional ByVal GRID_EPS As Double = 0.001, _
Optional ByVal MATURITY_STEPS As Long = 50, _
Optional ByVal STRIKE_STEPS As Long = 80, _
Optional ByVal DEFAULT_SIGMA As Double = 0.01, _
Optional ByVal GUESS_SIGMA As Double = 0.2, _
Optional ByVal LOWER_VAL As Double = 0.0000001, _
Optional ByVal UPPER_VAL As Double = 10, _
Optional ByVal OUTPUT As Integer = 0, _
Optional ByVal TDAYS_PER_YEAR As Double = 360, _
Optional ByVal CND_TYPE As Integer = 0)

'FILTERING_EPS: Reduce this to make ImpliedVolGraph smoother
'GRID_EPS: Reduce this to make SmoothVolGraph smoother
'MATURITY_STEPS: Extrapolation points for SmoothVolGraph
'STRIKE_STEPS: Extrapolation points for SmoothVolGraph

Dim i As Long
Dim NROWS As Long
Dim SPOT As Double
Dim DATE_STR As String
Dim SETTLEMENT As Date
Dim MATURITY_VEC As Variant
Dim STRIKE_VEC As Variant
Dim CALL_BID_VEC As Variant
Dim CALL_ASK_VEC As Variant
Dim PUT_BID_VEC As Variant
Dim PUT_ASK_VEC As Variant
Dim DATA_MATRIX As Variant

'On Error GoTo ERROR_LABEL

DATA_MATRIX = CONVERT_STRING_NUMBER_FUNC(CONVERT_TEXT_FILE_MATRIX_FUNC(FILE_PATH_STR, , , ","))

SPOT = DATA_MATRIX(1, 2)
DATE_STR = DATA_MATRIX(2, 1) 'Jun 07 2006 @ 19:24 ET (Data 20 Minutes Delayed)
DATE_STR = Left(DATE_STR, 6) & "," & Mid(DATE_STR, 8, 4)
SETTLEMENT = CDate(DATE_STR)

'---------------------------------------------------------------------------------
'DATA --> http://www.cboe.com/DelayedQuote/QuoteTableDownload.aspx
DATA_MATRIX = IMPLIED_VOLATILITY_SURFACE_BID_ASK_EXTRACT_FUNC(DATA_MATRIX, 0, 4, 5, 11, 12, , , 3)
DATA_MATRIX = IMPLIED_VOLATILITY_SURFACE_BID_ASK_MEAN_FUNC(DATA_MATRIX)
'---------------------------------------------------------------------------------
NROWS = UBound(DATA_MATRIX, 1)
ReDim MATURITY_VEC(1 To NROWS, 1 To 1)
ReDim STRIKE_VEC(1 To NROWS, 1 To 1)
ReDim CALL_BID_VEC(1 To NROWS, 1 To 1)
ReDim CALL_ASK_VEC(1 To NROWS, 1 To 1)
ReDim PUT_BID_VEC(1 To NROWS, 1 To 1)
ReDim PUT_ASK_VEC(1 To NROWS, 1 To 1)

For i = 1 To NROWS
    MATURITY_VEC(i, 1) = DATA_MATRIX(i, 1)
    STRIKE_VEC(i, 1) = DATA_MATRIX(i, 2)
    CALL_BID_VEC(i, 1) = DATA_MATRIX(i, 3)
    CALL_ASK_VEC(i, 1) = DATA_MATRIX(i, 4)
    PUT_BID_VEC(i, 1) = DATA_MATRIX(i, 5)
    PUT_ASK_VEC(i, 1) = DATA_MATRIX(i, 6)
Next i
'---------------------------------------------------------------------------------

IMPLIED_VOLATILITY_SURFACE_CBOE_FUNC = IMPLIED_VOLATILITY_SURFACE_FUNC(SPOT, RATE, SETTLEMENT, MATURITY_VEC, STRIKE_VEC, CALL_BID_VEC, CALL_ASK_VEC, PUT_BID_VEC, PUT_ASK_VEC, FILTERING_EPS, GRID_EPS, MATURITY_STEPS, STRIKE_STEPS, DEFAULT_SIGMA, GUESS_SIGMA, LOWER_VAL, UPPER_VAL, OUTPUT, TDAYS_PER_YEAR, CND_TYPE)

Exit Function
ERROR_LABEL:
IMPLIED_VOLATILITY_SURFACE_CBOE_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : IMPLIED_VOLATILITY_SURFACE_FUNC
'DESCRIPTION   : Implied Surface Function
'LIBRARY       : DERIVATIVES
'GROUP         : IMPLIED
'ID            : 002
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 21/10/2013
'************************************************************************************
'************************************************************************************

'1: Form quotes summary information of call and put price bid and
'   ask in Avg Summary

'2: Implied vol bid ask limits are placed in ImpliedVols Table

'3: Implied vols are filtered using objective function of minimizing
'   the vol gradients on the strike/maturity axes. Constraint is the
'   bid-ask limits

'4: Implied vols are then extrapolated on the non-quoted coordinates,
'   using objective of minimizing gradient and vol > 0 for constraint

'5: A Bicubic spline is used to smoothen the surface

Function IMPLIED_VOLATILITY_SURFACE_FUNC(ByVal SPOT As Double, _
ByVal RATE As Double, _
ByVal SETTLEMENT As Variant, _
ByRef MATURITY_RNG As Variant, _
ByRef STRIKE_RNG As Variant, _
ByRef CALL_BID_RNG As Variant, _
ByRef CALL_ASK_RNG As Variant, _
ByRef PUT_BID_RNG As Variant, _
ByRef PUT_ASK_RNG As Variant, _
Optional ByVal FILTERING_EPS As Double = 0.01, _
Optional ByVal GRID_EPS As Double = 0.001, _
Optional ByVal EXPIRATION_STEPS As Long = 50, _
Optional ByVal STRIKE_STEPS As Long = 80, _
Optional ByVal DEFAULT_SIGMA As Double = 0.01, _
Optional ByVal GUESS_SIGMA As Double = 0.2, _
Optional ByVal LOWER_VAL As Double = 0.0000001, _
Optional ByVal UPPER_VAL As Double = 10, _
Optional ByVal OUTPUT As Integer = 0, _
Optional ByVal TDAYS_PER_YEAR As Double = 252, _
Optional ByVal CND_TYPE As Integer = 0)

'----------------------------------------------------------------------------
'Call IMPLIED_VOLATILITY_SURFACE_RESET_FUNC
'----------------------------------------------------------------------------
'----------------------------------------------------------------------------
PUB_SPOT_VAL = SPOT
PUB_RATE_VAL = RATE
PUB_SIGMA_VAL = DEFAULT_SIGMA 'DefaultVolatility

PUB_CND_TYPE = CND_TYPE
'-----------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------

'1) FILTERING = This tolerance reduce this to make ImpliedVolGraph
'   smoother

'2) GRID = This tolerance reduce this to make SmoothVolGraph smoother

'3) EXPIRATION_STEPS = Surface Maturity Steps; <=Extrapolation points for
'   SmoothVolGraph

'4) STRIKE_STEPS = Surface Strike Steps; <=Extrapolation points for
'   SmoothVolGraph

'-----------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------
Dim i As Long
Dim j As Long
Dim k As Long
'-----------------------------------------------
Dim ii As Long
Dim jj As Long
'-----------------------------------------------
Dim NSIZE As Long
'-----------------------------------------------
Dim PUT_BID_VECTOR As Variant
Dim PUT_ASK_VECTOR As Variant
'-----------------------------------------------
Dim CALL_BID_VECTOR As Variant
Dim CALL_ASK_VECTOR As Variant
'-----------------------------------------------
Dim STRIKE_VECTOR As Variant
Dim MATURITY_VECTOR As Variant
'-----------------------------------------------
Dim TENOR_VAL As Double
Dim STRIKE_VAL As Double
Dim MATURITY_VAL As Double
'-----------------------------------------------
Dim CALL_BID_VAL As Double
Dim CALL_IMPL_BID_VAL As Double

Dim CALL_ASK_VAL As Double
Dim CALL_IMPL_ASK_VAL As Double
'-----------------------------------------------
Dim PUT_BID_VAL As Double
Dim PUT_IMPL_BID_VAL As Double
'-----------------------------------------------
Dim PUT_ASK_VAL As Double
Dim PUT_IMPL_ASK_VAL As Double
'-----------------------------------------------
Dim BID_IMPLIED_VAL As Double
Dim ASK_IMPLIED_VAL As Double
'-----------------------------------------------
Dim DPRICE_VAL As Double
Dim IPRICE_VAL As Double
'-----------------------------------------------
Dim TEMP_VAL As Double
'-----------------------------------------------
Dim TEMP1_ARR As Variant
Dim TEMP2_ARR As Variant
Dim TEMP_GROUP As Variant
'-----------------------------------------------
Dim XKEY_GROUP_ARR As Variant
Dim XDATA_GROUP_ARR As Variant
'-----------------------------------------------
Dim YKEY_GROUP_ARR As Variant
Dim YDATA_GROUP_ARR As Variant
'-----------------------------------------------
'-----------------------------------------------
Dim PARAM_VECTOR As Variant
'-----------------------------------------------

'On Error GoTo ERROR_LABEL

'--------------------------------------------------------------------------------
'1) MARKET DIRECTION --> BEARISH
'--------------------------------------------------------------------------------
'    a) IMPLIED VOLATILITY: LOW

'Buy Naked Puts
'Bear Vertical Spreads:
'Buy ATM Call/Sell ITM Call
'Buy ATM Put/Sell OTM Put
'Sell OTM (ITM) Call (Put) Butterflies
'Buy ITM (OTM) Call (Put) Time Spreads
    

'    b) IMPLIED VOLATILITY: NEUTRAL
'Sell the Underlying

'    c) IMPLIED VOLATILITY: HIGH
'Sell Naked Calls
'Bear Vertical Spreads:
'Buy OTM Call/Sell ATM Call
'Buy OTM (ITM) Call (Put) Time Spreads
'Buy ITM (OTM) Call (Put) Butterflies
'Sell OTM (ITM) Call (Put) Time Spreads

'--------------------------------------------------------------------------------
'2) MARKET DIRECTION --> NEUTRAL
'--------------------------------------------------------------------------------

'    a) IMPLIED VOLATILITY: LOW
'Backspreads
'Buy Straddles / Strangles
'Sell ATM Call Or Put Butterflies
'Buy ATM Call Or Put Time Spreads

'    b) IMPLIED VOLATILITY: NEUTRAL
'Do Nothing

'    c) IMPLIED VOLATILITY: HIGH
'Ratio Vertical Spreads
'Sell Straddles / Strangles
'Buy ATM Call Or Put Butterflies
'Sell ATM Call Or Put Time Spreads

'--------------------------------------------------------------------------------
'3) MARKET DIRECTION --> BULLISH
'--------------------------------------------------------------------------------
'    a) IMPLIED VOLATILITY: LOW
'Buy Naked Calls
'Bull Vertical Spreads
'Buy ATM Call/Sell OTM Call
'--------------------------------------------------------------------------------


MATURITY_VECTOR = MATURITY_RNG
If UBound(MATURITY_VECTOR, 1) = 1 Then
    MATURITY_VECTOR = MATRIX_TRANSPOSE_FUNC(MATURITY_VECTOR)
End If

STRIKE_VECTOR = STRIKE_RNG
If UBound(STRIKE_VECTOR, 1) = 1 Then
    STRIKE_VECTOR = MATRIX_TRANSPOSE_FUNC(STRIKE_VECTOR)
End If

CALL_BID_VECTOR = CALL_BID_RNG
If UBound(CALL_BID_VECTOR, 1) = 1 Then
    CALL_BID_VECTOR = MATRIX_TRANSPOSE_FUNC(CALL_BID_VECTOR)
End If

CALL_ASK_VECTOR = CALL_ASK_RNG
If UBound(CALL_ASK_VECTOR, 1) = 1 Then
    CALL_ASK_VECTOR = MATRIX_TRANSPOSE_FUNC(CALL_ASK_VECTOR)
End If

PUT_BID_VECTOR = PUT_BID_RNG
If UBound(PUT_BID_VECTOR, 1) = 1 Then
    PUT_BID_VECTOR = MATRIX_TRANSPOSE_FUNC(PUT_BID_VECTOR)
End If

PUT_ASK_VECTOR = PUT_ASK_RNG
If UBound(PUT_ASK_VECTOR, 1) = 1 Then
    PUT_ASK_VECTOR = MATRIX_TRANSPOSE_FUNC(PUT_ASK_VECTOR)
End If
NSIZE = UBound(MATURITY_VECTOR, 1)

k = 0

ReDim PUB_GROUP_ARR(1 To 6)
ReDim PUB_TENOR_ARR(1 To 1)
ReDim PUB_STRIKE_ARR(1 To 1)

ReDim PUB_IMPL_BID_ARR(1 To 1)
ReDim PUB_IMPL_ASK_ARR(1 To 1)

ReDim FLAG_ARR(1 To 1) As Variant
ReDim OPT_PRICE_BID_ARR(1 To 1) As Variant
ReDim OPT_PRICE_ASK_ARR(1 To 1) As Variant

'-----------------------------------------------------------------------------
'-----------------------------------------------------------------------------
'--------------------------------IMPLIED-VOLATILITY---------------------------
'-----------------------------------------------------------------------------
'-----------------------------------------------------------------------------
For i = 1 To NSIZE
'----------------------------------------------------------------------------------
    STRIKE_VAL = STRIKE_VECTOR(i, 1)
    MATURITY_VAL = MATURITY_VECTOR(i, 1)
    'TENOR_VAL = NETWORKDAYS_FUNC(CDate(SETTLEMENT), CDate(MATURITY_VAL), HOLIDAYS_RNG) / TDAYS_PER_YEAR
    TENOR_VAL = COUNT_DAYS_FUNC(SETTLEMENT, MATURITY_VAL, 1) / TDAYS_PER_YEAR
'----------------------------------------------------------------------------------
    DPRICE_VAL = CALL_BID_VECTOR(i, 1)
    PUB_TARGET_VAL = DPRICE_VAL
    PUB_STRIKE_VAL = STRIKE_VAL
    PUB_EXPIRATION_VAL = TENOR_VAL
    PUB_OPTION_FLAG = 1
    IPRICE_VAL = IMPLIED_VOLATILITY_SURFACE_SOLVER_FUNC(LOWER_VAL, UPPER_VAL, GUESS_SIGMA)
    If IPRICE_VAL = PUB_EPSILON Or IPRICE_VAL <= 0 Then
      CALL_BID_VAL = IMPLIED_VOLATILITY_SURFACE_BSM_FUNC(PUB_SIGMA_VAL, 1)
      CALL_IMPL_BID_VAL = PUB_SIGMA_VAL
    Else
      CALL_BID_VAL = DPRICE_VAL
      CALL_IMPL_BID_VAL = IPRICE_VAL
    End If
'---------------------------------------------------------------------------------
    DPRICE_VAL = CALL_ASK_VECTOR(i, 1)
    PUB_TARGET_VAL = DPRICE_VAL
    PUB_STRIKE_VAL = STRIKE_VAL
    PUB_EXPIRATION_VAL = TENOR_VAL
    PUB_OPTION_FLAG = 1
    IPRICE_VAL = IMPLIED_VOLATILITY_SURFACE_SOLVER_FUNC(LOWER_VAL, UPPER_VAL, GUESS_SIGMA)
    If IPRICE_VAL = PUB_EPSILON Or IPRICE_VAL <= 0 Then
        CALL_ASK_VAL = IMPLIED_VOLATILITY_SURFACE_BSM_FUNC(PUB_SIGMA_VAL, 1)
        CALL_IMPL_ASK_VAL = PUB_SIGMA_VAL
    Else
        CALL_ASK_VAL = DPRICE_VAL
        CALL_IMPL_ASK_VAL = IPRICE_VAL
    End If
'---------------------------------------------------------------------------------
'----------------------------------------------------------------------------------
    DPRICE_VAL = PUT_BID_VECTOR(i, 1)
    PUB_TARGET_VAL = DPRICE_VAL
    PUB_STRIKE_VAL = STRIKE_VAL
    PUB_EXPIRATION_VAL = TENOR_VAL
    PUB_OPTION_FLAG = -1
    IPRICE_VAL = IMPLIED_VOLATILITY_SURFACE_SOLVER_FUNC(LOWER_VAL, UPPER_VAL, GUESS_SIGMA)
    If IPRICE_VAL = PUB_EPSILON Or IPRICE_VAL <= 0 Then
        PUT_BID_VAL = IMPLIED_VOLATILITY_SURFACE_BSM_FUNC(PUB_SIGMA_VAL, 1)
        PUT_IMPL_BID_VAL = PUB_SIGMA_VAL
    Else
        PUT_BID_VAL = DPRICE_VAL
        PUT_IMPL_BID_VAL = IPRICE_VAL
    End If
'---------------------------------------------------------------------------------
    DPRICE_VAL = PUT_ASK_VECTOR(i, 1)
    PUB_TARGET_VAL = DPRICE_VAL
    PUB_STRIKE_VAL = STRIKE_VAL
    PUB_EXPIRATION_VAL = TENOR_VAL
    PUB_OPTION_FLAG = -1
    IPRICE_VAL = IMPLIED_VOLATILITY_SURFACE_SOLVER_FUNC(LOWER_VAL, UPPER_VAL, GUESS_SIGMA)
    If IPRICE_VAL = PUB_EPSILON Or IPRICE_VAL <= 0 Then
      PUT_ASK_VAL = IMPLIED_VOLATILITY_SURFACE_BSM_FUNC(PUB_SIGMA_VAL, 1)
      PUT_IMPL_ASK_VAL = PUB_SIGMA_VAL
    Else
      PUT_ASK_VAL = DPRICE_VAL
      PUT_IMPL_ASK_VAL = IPRICE_VAL
    End If
'---------------------------------------------------------------------------------
    BID_IMPLIED_VAL = MINIMUM_FUNC(CALL_IMPL_BID_VAL, CALL_IMPL_ASK_VAL)
    BID_IMPLIED_VAL = MINIMUM_FUNC(BID_IMPLIED_VAL, PUT_IMPL_BID_VAL)
    BID_IMPLIED_VAL = MINIMUM_FUNC(BID_IMPLIED_VAL, PUT_IMPL_ASK_VAL)
    ASK_IMPLIED_VAL = MAXIMUM_FUNC(CALL_IMPL_BID_VAL, CALL_IMPL_ASK_VAL)
    ASK_IMPLIED_VAL = MAXIMUM_FUNC(ASK_IMPLIED_VAL, PUT_IMPL_BID_VAL)
    ASK_IMPLIED_VAL = MAXIMUM_FUNC(ASK_IMPLIED_VAL, PUT_IMPL_ASK_VAL)
'---------------------------------------------------------------------------------
    If BID_IMPLIED_VAL < PUB_SIGMA_VAL Then: BID_IMPLIED_VAL = PUB_SIGMA_VAL
    If ASK_IMPLIED_VAL < PUB_SIGMA_VAL Then: ASK_IMPLIED_VAL = PUB_SIGMA_VAL
    If (CALL_ASK_VAL > CALL_BID_VAL) And (PUT_ASK_VAL > PUT_BID_VAL) Then
      ReDim Preserve PUB_TENOR_ARR(1 To k + 1)
      ReDim Preserve PUB_STRIKE_ARR(1 To k + 1)
      ReDim Preserve FLAG_ARR(1 To k + 1)
      ReDim Preserve OPT_PRICE_BID_ARR(1 To k + 1)
      ReDim Preserve OPT_PRICE_ASK_ARR(1 To k + 1)
      ReDim Preserve PUB_IMPL_BID_ARR(1 To k + 1)
      ReDim Preserve PUB_IMPL_ASK_ARR(1 To k + 1)
      PUB_TENOR_ARR(k + 1) = TENOR_VAL
      PUB_STRIKE_ARR(k + 1) = STRIKE_VAL
      PUB_IMPL_BID_ARR(k + 1) = BID_IMPLIED_VAL
      PUB_IMPL_ASK_ARR(k + 1) = ASK_IMPLIED_VAL
      k = k + 1
    End If
Next i

ReDim TEMP2_ARR(1 To UBound(PUB_TENOR_ARR), 1 To 4)
For i = 1 To UBound(PUB_TENOR_ARR)
    TEMP2_ARR(i, 1) = PUB_TENOR_ARR(i)
    TEMP2_ARR(i, 2) = PUB_STRIKE_ARR(i)
    TEMP2_ARR(i, 3) = PUB_IMPL_BID_ARR(i)
    TEMP2_ARR(i, 4) = PUB_IMPL_ASK_ARR(i)
Next i
PUB_GROUP_ARR(1) = TEMP2_ARR 'Implied Vols
If OUTPUT = 3 Then
    IMPLIED_VOLATILITY_SURFACE_FUNC = PUB_GROUP_ARR(1) 'IMPLIED - VOLATILITY TABLE
    Exit Function
End If
'------------------------------------------------------------------------------
'------------------------------------------------------------------------------
'------------------------------FILTERED VOLS-----------------------------------
'------------------------------------------------------------------------------
'------------------------------------------------------------------------------
ReDim PARAM_VECTOR(1 To UBound(PUB_IMPL_BID_ARR, 1), 1 To 1)
For i = 1 To UBound(PUB_IMPL_BID_ARR, 1)
    PARAM_VECTOR(i, 1) = (PUB_IMPL_BID_ARR(i) + PUB_IMPL_ASK_ARR(i)) / 2
Next i
' TEMP1_ARR = PARAM_VECTOR
TEMP1_ARR = SIMPLEX_MINIMUM_OPTIMIZATION_FUNC("IMPLIED_VOLATILITY_SURFACE_FILTERED_OBJ_FUNC", _
            "IMPLIED_VOLATILITY_SURFACE_FILTERED_CONST_FUNC", PARAM_VECTOR, 0.01, 200, FILTERING_EPS)
'Debug.Print IMPLIED_VOLATILITY_SURFACE_FILTERED_CONST_FUNC(PARAM_VECTOR)
ReDim PUB_IMPL_SIGMA_ARR(1 To UBound(PUB_IMPL_BID_ARR, 1))
For i = 1 To UBound(PUB_IMPL_BID_ARR, 1)
    PUB_IMPL_SIGMA_ARR(i) = TEMP1_ARR(i, 1)
Next i
ReDim TEMP2_ARR(1 To UBound(PUB_TENOR_ARR, 1), 1 To 3)
For i = 1 To UBound(PUB_TENOR_ARR, 1)
    TEMP2_ARR(i, 1) = PUB_TENOR_ARR(i)
    TEMP2_ARR(i, 2) = PUB_STRIKE_ARR(i)
    TEMP2_ARR(i, 3) = PUB_IMPL_SIGMA_ARR(i)
Next i
PUB_GROUP_ARR(2) = TEMP2_ARR 'Filtered Vols
If OUTPUT = 2 Then
    IMPLIED_VOLATILITY_SURFACE_FUNC = PUB_GROUP_ARR(2) 'FILTERED - VOLATILITY TABLE
    Exit Function
End If
'---------------------------------------------------------------------------
'---------------------------------------------------------------------------
'---------------------------VOLATILITY SURFACE------------------------------
'---------------------------------------------------------------------------
'---------------------------------------------------------------------------

ReDim TEMP2_ARR(1 To UBound(PUB_TENOR_ARR, 1), 1 To 2)
For i = 1 To UBound(PUB_TENOR_ARR, 1)
    TEMP2_ARR(i, 1) = PUB_STRIKE_ARR(i)
    TEMP2_ARR(i, 2) = PUB_IMPL_SIGMA_ARR(i)
Next i

TEMP_GROUP = IMPLIED_VOLATILITY_SURFACE_MATRIX_GROUP_FUNC(PUB_TENOR_ARR, TEMP2_ARR, 2)
XDATA_GROUP_ARR = TEMP_GROUP(1)
XKEY_GROUP_ARR = TEMP_GROUP(2)
ReDim TEMP2_ARR(1 To UBound(PUB_STRIKE_ARR, 1), 1 To 2)
For i = 1 To UBound(PUB_STRIKE_ARR, 1)
    TEMP2_ARR(i, 1) = PUB_TENOR_ARR(i)
    TEMP2_ARR(i, 2) = PUB_IMPL_SIGMA_ARR(i)
Next i
TEMP_GROUP = IMPLIED_VOLATILITY_SURFACE_MATRIX_GROUP_FUNC(PUB_STRIKE_ARR, TEMP2_ARR, 2)
YDATA_GROUP_ARR = TEMP_GROUP(1)
YKEY_GROUP_ARR = TEMP_GROUP(2)
'Axes of the GRID_EPS is formed by the XKEY_GROUP_ARR and YKEY_GROUP_ARR
'parameters to be passed for optimization are all the points
'which do not belong to original points
ReDim PUB_3D_MATRIX(1 To UBound(XKEY_GROUP_ARR, 1), _
1 To UBound(YKEY_GROUP_ARR, 1), 1 To 4)
For i = 1 To UBound(XKEY_GROUP_ARR, 1)
    For j = 1 To UBound(YKEY_GROUP_ARR, 1)
        PUB_3D_MATRIX(i, j, 1) = XKEY_GROUP_ARR(i)
        PUB_3D_MATRIX(i, j, 2) = YKEY_GROUP_ARR(j)
        PUB_3D_MATRIX(i, j, 3) = 0 '0 indicates Z-point is fixed
    Next j
Next i

For i = 1 To UBound(XDATA_GROUP_ARR, 1)
    k = i
    TEMP1_ARR = XDATA_GROUP_ARR(i) 'matrix of y,z
    For j = 1 To UBound(TEMP1_ARR, 1)
        TEMP_VAL = TEMP1_ARR(j, 1)
        jj = -1
        For ii = 1 To UBound(YKEY_GROUP_ARR, 1)
            If TEMP_VAL = YKEY_GROUP_ARR(ii) Then
                jj = ii
                GoTo 1983
            End If
        Next ii
1983:
        If jj <= 0 Then: GoTo ERROR_LABEL  'error in GridHelper.Extrapolate
        PUB_3D_MATRIX(k, jj, 3) = 1 '1 indicates Z-point has data
        PUB_3D_MATRIX(k, jj, 4) = TEMP1_ARR(j, 2)
    Next j
Next i

'now array of parameters is to be made, which will be used for
'optimization .They are the points where PUB_3D_MATRIX(k, jj, 3)=0

k = 0
For i = 1 To UBound(PUB_3D_MATRIX, 1)
    For j = 1 To UBound(PUB_3D_MATRIX, 2)
        If PUB_3D_MATRIX(i, j, 3) = 0 Then: k = k + 1
    Next j
Next i

ReDim PARAM_VECTOR(1 To k, 1 To 1)
ReDim PUB_INDEX_ARR(1 To k, 1 To 2) 'contains index points for x,y

k = 1
For i = 1 To UBound(PUB_3D_MATRIX, 1)
    For j = 1 To UBound(PUB_3D_MATRIX, 2)
        If PUB_3D_MATRIX(i, j, 3) = 0 Then
            PARAM_VECTOR(k, 1) = GUESS_SIGMA
            PUB_INDEX_ARR(k, 1) = i
            PUB_INDEX_ARR(k, 2) = j
            k = k + 1
        End If
    Next j
Next i

' TEMP1_ARR = PARAM_VECTOR
TEMP1_ARR = SIMPLEX_MINIMUM_OPTIMIZATION_FUNC("IMPLIED_VOLATILITY_SURFACE_OBJ_FUNC", _
            "IMPLIED_VOLATILITY_SURFACE_CONST_FUNC", PARAM_VECTOR, 0.01, 200, GRID_EPS)
'Debug.Print IMPLIED_VOLATILITY_SURFACE_CONST_FUNC(PARAM_VECTOR)

For i = 1 To UBound(TEMP1_ARR, 1)
    ii = PUB_INDEX_ARR(i, 1)
    jj = PUB_INDEX_ARR(i, 2)
    PUB_3D_MATRIX(ii, jj, 4) = TEMP1_ARR(i, 1)
Next i

ReDim TEMP2_ARR(1 To UBound(PUB_3D_MATRIX, 1) + 1, 1 To UBound(PUB_3D_MATRIX, 2) + 1)
For i = 1 To UBound(TEMP2_ARR, 1) - 1
    For j = 1 To UBound(TEMP2_ARR, 2) - 1
        TEMP2_ARR(i + 1, j + 1) = PUB_3D_MATRIX(i, j, 4)
    Next j
Next i
For i = 1 To UBound(XKEY_GROUP_ARR)
    TEMP2_ARR(i + 1, 1) = XKEY_GROUP_ARR(i)
Next i
For i = 1 To UBound(YKEY_GROUP_ARR)
    TEMP2_ARR(1, i + 1) = YKEY_GROUP_ARR(i)
Next i
PUB_GROUP_ARR(3) = TEMP2_ARR 'Original Implied Volatility Surface
If OUTPUT = 1 Then
    IMPLIED_VOLATILITY_SURFACE_FUNC = PUB_GROUP_ARR(3) 'SURFACE - VOLATILITY TABLE
    Exit Function
End If

'---------------------------------------------------------------------------
'---------------------------------------------------------------------------
'----------------------------SMOOTH_CUBIC_SPLINE----------------------------
TEMP2_ARR = IMPLIED_VOLATILITY_SURFACE_GET_FUNC(TEMP2_ARR, 2, 2, EXPIRATION_STEPS, STRIKE_STEPS)
'---------------------------------------------------------------------------
'---------------------------------------------------------------------------
'---------------------------------------------------------------------------
PUB_GROUP_ARR(4) = TEMP2_ARR 'Smoothed Implied Volatility Surface
'---------------------------------------------------------------------------

Select Case OUTPUT
Case 0
    IMPLIED_VOLATILITY_SURFACE_FUNC = PUB_GROUP_ARR(4) 'SMOOTH - VOLATILITY TABLE
Case Else
    IMPLIED_VOLATILITY_SURFACE_FUNC = PUB_GROUP_ARR
End Select

Exit Function
ERROR_LABEL:
IMPLIED_VOLATILITY_SURFACE_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : IMPLIED_VOLATILITY_SURFACE_FILTERED_OBJ_FUNC
'DESCRIPTION   : Filtered Volatility Objective Function
'LIBRARY       : DERIVATIVES
'GROUP         : IMPLIED
'ID            : 003
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 21/10/2013
'************************************************************************************
'************************************************************************************

Function IMPLIED_VOLATILITY_SURFACE_FILTERED_OBJ_FUNC(ByRef PARAM_RNG As Variant)
  
Dim i As Long
Dim j As Long

Dim TEMP_SUM As Double
Dim STRIKE_VAL As Double
Dim TEMP_SIGMA As Double
Dim MATURITY_VAL As Double

Dim CURVE_VAL As Double
Dim SLOPE1_VAL As Double
Dim SLOPE2_VAL As Double

Dim TEMP_GROUP As Variant

Dim PARAM_VECTOR As Variant
Dim TEMP_MATRIX As Variant 'cols strike,impliedvols
Dim XDATA_GROUP_ARR As Variant

'On Error GoTo ERROR_LABEL

PARAM_VECTOR = PARAM_RNG
TEMP_SUM = 0

ReDim TEMP_MATRIX(1 To UBound(PUB_TENOR_ARR, 1), 1 To 2)
For i = 1 To UBound(PUB_TENOR_ARR, 1)
  TEMP_MATRIX(i, 1) = PUB_STRIKE_ARR(i)
  TEMP_MATRIX(i, 2) = PARAM_VECTOR(i, 1)
Next i

XDATA_GROUP_ARR = IMPLIED_VOLATILITY_SURFACE_MATRIX_GROUP_FUNC(PUB_TENOR_ARR, TEMP_MATRIX, 0)

For i = 1 To UBound(XDATA_GROUP_ARR, 1)   'loop over the maturity
  TEMP_GROUP = XDATA_GROUP_ARR(i) 'matrix of strike,impliedVols
  'loop over the strikes
  For j = 2 To UBound(TEMP_GROUP, 1) - 1
    TEMP_SIGMA = TEMP_GROUP(j, 2) - TEMP_GROUP(j - 1, 2)
    STRIKE_VAL = TEMP_GROUP(j, 1) - TEMP_GROUP(j - 1, 1)
    SLOPE1_VAL = TEMP_SIGMA / STRIKE_VAL
    
    TEMP_SIGMA = TEMP_GROUP(j + 1, 2) - TEMP_GROUP(j, 2)
    STRIKE_VAL = TEMP_GROUP(j + 1, 1) - TEMP_GROUP(j, 1)
    SLOPE2_VAL = TEMP_SIGMA / STRIKE_VAL
    CURVE_VAL = (SLOPE2_VAL - SLOPE1_VAL) / (TEMP_GROUP(j + 1, 1) - TEMP_GROUP(j - 1, 1))
    CURVE_VAL = Abs(CURVE_VAL)
    CURVE_VAL = CURVE_VAL * 1000
    TEMP_SUM = TEMP_SUM + CURVE_VAL
  Next j
Next i
  

ReDim TEMP_MATRIX(1 To UBound(PUB_STRIKE_ARR, 1), 1 To 2)
For i = 1 To UBound(PUB_STRIKE_ARR, 1)
  TEMP_MATRIX(i, 1) = PUB_TENOR_ARR(i) 'maturity time is 2st col
  TEMP_MATRIX(i, 2) = PARAM_VECTOR(i, 1) 'implied vol is 2nd col
Next i

XDATA_GROUP_ARR = IMPLIED_VOLATILITY_SURFACE_MATRIX_GROUP_FUNC(PUB_STRIKE_ARR, TEMP_MATRIX, 0)

For i = 1 To UBound(XDATA_GROUP_ARR, 1) 'loop over the strike
  TEMP_GROUP = XDATA_GROUP_ARR(i) 'matrix of maturity,impliedVols
  'loop over the maturity
  For j = 2 To UBound(TEMP_GROUP, 1) - 1
    
    TEMP_SIGMA = TEMP_GROUP(j, 2) - TEMP_GROUP(j - 1, 2)
    MATURITY_VAL = TEMP_GROUP(j, 1) - TEMP_GROUP(j - 1, 1)
    SLOPE1_VAL = TEMP_SIGMA / MATURITY_VAL
    
    TEMP_SIGMA = TEMP_GROUP(j + 1, 2) - TEMP_GROUP(j, 2)
    MATURITY_VAL = TEMP_GROUP(j + 1, 1) - TEMP_GROUP(j, 1)
    SLOPE2_VAL = TEMP_SIGMA / MATURITY_VAL
    CURVE_VAL = (SLOPE2_VAL - SLOPE1_VAL) / (TEMP_GROUP(j + 1, 1) - TEMP_GROUP(j - 1, 1))
    CURVE_VAL = Abs(CURVE_VAL)
    CURVE_VAL = CURVE_VAL * 1000
    TEMP_SUM = TEMP_SUM + CURVE_VAL
  Next j
Next i

IMPLIED_VOLATILITY_SURFACE_FILTERED_OBJ_FUNC = TEMP_SUM

Exit Function
ERROR_LABEL:
IMPLIED_VOLATILITY_SURFACE_FILTERED_OBJ_FUNC = PUB_EPSILON
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : IMPLIED_VOLATILITY_SURFACE_FILTERED_CONST_FUNC
'DESCRIPTION   : Filtered Volatility Constraint Function
'LIBRARY       : DERIVATIVES
'GROUP         : IMPLIED
'ID            : 004
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 21/10/2013
'************************************************************************************
'************************************************************************************

Function IMPLIED_VOLATILITY_SURFACE_FILTERED_CONST_FUNC(ByRef PARAM_VECTOR As Variant)
  
Dim i As Long
Dim X_VAL As Double

'On Error GoTo ERROR_LABEL

IMPLIED_VOLATILITY_SURFACE_FILTERED_CONST_FUNC = True

For i = 1 To UBound(PARAM_VECTOR, 1)
    X_VAL = PARAM_VECTOR(i, 1)
    If (X_VAL < PUB_IMPL_BID_ARR(i)) Or (X_VAL > PUB_IMPL_ASK_ARR(i)) Then
        IMPLIED_VOLATILITY_SURFACE_FILTERED_CONST_FUNC = False
        Exit Function
    End If
Next i

Exit Function
ERROR_LABEL:
IMPLIED_VOLATILITY_SURFACE_FILTERED_CONST_FUNC = False
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : IMPLIED_VOLATILITY_SURFACE_CONST_FUNC
'DESCRIPTION   : Volatility Surface Constraint Function
'LIBRARY       : DERIVATIVES
'GROUP         : IMPLIED
'ID            : 005
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 21/10/2013
'************************************************************************************
'************************************************************************************

Function IMPLIED_VOLATILITY_SURFACE_CONST_FUNC(ByRef PARAM_VECTOR As Variant)
  
Dim i As Long

Dim X_VAL As Double

'On Error GoTo ERROR_LABEL

IMPLIED_VOLATILITY_SURFACE_CONST_FUNC = True

For i = 1 To UBound(PARAM_VECTOR, 1)
    X_VAL = PARAM_VECTOR(i, 1)
    If X_VAL <= 0 Then
        IMPLIED_VOLATILITY_SURFACE_CONST_FUNC = False
        Exit Function
    End If
Next i

Exit Function
ERROR_LABEL:
IMPLIED_VOLATILITY_SURFACE_CONST_FUNC = False
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : IMPLIED_VOLATILITY_SURFACE_OBJ_FUNC
'DESCRIPTION   : Volatility Surface Objective Function
'LIBRARY       : DERIVATIVES
'GROUP         : IMPLIED
'ID            : 006
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 21/10/2013
'************************************************************************************
'************************************************************************************

Function IMPLIED_VOLATILITY_SURFACE_OBJ_FUNC(ByRef PARAM_RNG As Variant)
  
Dim i As Long
Dim j As Long

Dim ii As Long
Dim jj As Long

Dim DX_VAL As Double
Dim DY_VAL As Double
Dim DZ_VAL As Double

Dim SLOPE1_VAL As Double
Dim SLOPE2_VAL As Double

Dim CURVE_VAL As Double
Dim RESULT_VAL As Double

Dim TEMP_MATRIX As Variant
'Dim PARAM_RNG As Variant

'On Error GoTo ERROR_LABEL

'PARAM_RNG = PARAM_RNG
TEMP_MATRIX = PUB_3D_MATRIX
For i = 1 To UBound(PARAM_RNG, 1)
    ii = PUB_INDEX_ARR(i, 1)
    jj = PUB_INDEX_ARR(i, 2)
    TEMP_MATRIX(ii, jj, 4) = PARAM_RNG(i, 1)
Next i
  
For i = 1 To UBound(TEMP_MATRIX, 1)
  For j = 2 To UBound(TEMP_MATRIX, 2) - 1
      DZ_VAL = TEMP_MATRIX(i, j, 4) - TEMP_MATRIX(i, j - 1, 4)
      DY_VAL = TEMP_MATRIX(i, j, 2) - TEMP_MATRIX(i, j - 1, 2)
      SLOPE1_VAL = DZ_VAL / DY_VAL
      DZ_VAL = TEMP_MATRIX(i, j + 1, 4) - TEMP_MATRIX(i, j, 4)
      DY_VAL = TEMP_MATRIX(i, j + 1, 2) - TEMP_MATRIX(i, j, 2)
      SLOPE2_VAL = DZ_VAL / DY_VAL
      CURVE_VAL = (SLOPE2_VAL - SLOPE1_VAL) / (TEMP_MATRIX(i, j + 1, 2) - TEMP_MATRIX(i, j - 1, 2))
      CURVE_VAL = Abs(CURVE_VAL)
      CURVE_VAL = CURVE_VAL * 1000
      'RESULT_VAL = RESULT_VAL + CURVE_VAL
  Next j
Next i

For i = 2 To UBound(TEMP_MATRIX, 1) - 1
  For j = 1 To UBound(TEMP_MATRIX, 2)
        DZ_VAL = TEMP_MATRIX(i, j, 4) - TEMP_MATRIX(i - 1, j, 4)
        DX_VAL = TEMP_MATRIX(i, j, 1) - TEMP_MATRIX(i - 1, j, 1)
        SLOPE1_VAL = DZ_VAL / DX_VAL
        DZ_VAL = TEMP_MATRIX(i + 1, j, 4) - TEMP_MATRIX(i, j, 4)
        DX_VAL = TEMP_MATRIX(i + 1, j, 1) - TEMP_MATRIX(i, j, 1)
        SLOPE2_VAL = DZ_VAL / DX_VAL
        CURVE_VAL = (SLOPE2_VAL - SLOPE1_VAL) / (TEMP_MATRIX(i + 1, j, 1) - TEMP_MATRIX(i - 1, j, 1))
        CURVE_VAL = Abs(CURVE_VAL)
        CURVE_VAL = CURVE_VAL * 1000
        RESULT_VAL = RESULT_VAL + CURVE_VAL
  Next j
Next i

IMPLIED_VOLATILITY_SURFACE_OBJ_FUNC = RESULT_VAL

Exit Function
ERROR_LABEL:
IMPLIED_VOLATILITY_SURFACE_OBJ_FUNC = PUB_EPSILON
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : IMPLIED_VOLATILITY_SURFACE_SOLVER_FUNC
'DESCRIPTION   : Implied Volatility Zero Function
'LIBRARY       : DERIVATIVES
'GROUP         : IMPLIED
'ID            : 007
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 21/10/2013
'************************************************************************************
'************************************************************************************

Private Function IMPLIED_VOLATILITY_SURFACE_SOLVER_FUNC( _
Optional ByVal LOWER_VAL As Double = 0.0000001, _
Optional ByVal UPPER_VAL As Double = 10, _
Optional ByVal GUESS_SIGMA As Double = 0.2)

Dim CONVERG_VAL As Integer
Dim COUNTER As Long
Dim Y_VAL As Double

'On Error GoTo ERROR_LABEL
CONVERG_VAL = 0
COUNTER = 0

'Y_VAL = BRENT_ZERO_FUNC(LOWER_VAL, UPPER_VAL, "IMPLIED_VOLATILITY_SURFACE_BSM_FUNC", GUESS_SIGMA, CONVERG_VAL, COUNTER, 100, 0.0001)
Y_VAL = MULLER_ZERO_FUNC(LOWER_VAL, UPPER_VAL, "IMPLIED_VOLATILITY_SURFACE_BSM_FUNC", CONVERG_VAL, COUNTER, 1000, 10 ^ -10)
'Debug.Print CONVERG_VAL
If CONVERG_VAL <> 0 Then: GoTo ERROR_LABEL 'Y_VAL = 0
IMPLIED_VOLATILITY_SURFACE_SOLVER_FUNC = Y_VAL

Exit Function
ERROR_LABEL:
IMPLIED_VOLATILITY_SURFACE_SOLVER_FUNC = PUB_EPSILON
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : IMPLIED_VOLATILITY_SURFACE_BSM_FUNC
'DESCRIPTION   : Implied Volatility Solver Function
'LIBRARY       : DERIVATIVES
'GROUP         : IMPLIED
'ID            : 008
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 21/10/2013
'************************************************************************************
'************************************************************************************

Private Function IMPLIED_VOLATILITY_SURFACE_BSM_FUNC(ByVal X_VAL As Double, _
Optional ByVal OUTPUT As Integer = 0)

'OUTPUT for Brent Optimization must be = 0
Dim Y_VAL As Double

'On Error GoTo ERROR_LABEL

Select Case PUB_OPTION_FLAG
Case 1 ', "c", "call"
    Y_VAL = BLACK_SCHOLES_OPTION_FUNC(PUB_SPOT_VAL, PUB_STRIKE_VAL, PUB_EXPIRATION_VAL, PUB_RATE_VAL, X_VAL, 1, PUB_CND_TYPE)
Case Else '-1, "p", "put"
    Y_VAL = BLACK_SCHOLES_OPTION_FUNC(PUB_SPOT_VAL, PUB_STRIKE_VAL, PUB_EXPIRATION_VAL, PUB_RATE_VAL, X_VAL, -1, PUB_CND_TYPE)
End Select

Select Case OUTPUT
Case 0
    IMPLIED_VOLATILITY_SURFACE_BSM_FUNC = Y_VAL - PUB_TARGET_VAL
Case Else
    IMPLIED_VOLATILITY_SURFACE_BSM_FUNC = Y_VAL
End Select

Exit Function
ERROR_LABEL:
IMPLIED_VOLATILITY_SURFACE_BSM_FUNC = PUB_EPSILON
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : IMPLIED_VOLATILITY_SURFACE_GET_FUNC
'DESCRIPTION   : Get Surface Function
'LIBRARY       : DERIVATIVES
'GROUP         : IMPLIED
'ID            : 009
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 21/10/2013
'************************************************************************************
'************************************************************************************

Private Function IMPLIED_VOLATILITY_SURFACE_GET_FUNC(ByRef DATA_RNG As Variant, _
Optional ByVal SROW As Long = 2, _
Optional ByVal SCOLUMN As Long = 2, _
Optional ByVal EXPIRATION_STEPS As Long = 50, _
Optional ByVal STRIKE_STEPS As Long = 80)
  
Dim i As Long
Dim j As Long
Dim k As Long

Dim XDATA_ARR As Variant
Dim YDATA_ARR As Variant

Dim XI_VAL As Double
Dim YI_VAL As Double

Dim XT_VAL As Double
Dim YT_VAL As Double
Dim ZT_VAL As Double

Dim TEMP_ARR As Variant
Dim TEMP_GROUP As Variant

Dim XTEMP_VECTOR As Variant
Dim YTEMP_VECTOR As Variant

Dim SPLINE_ARR As Variant
Dim ZTEMP_ARR As Variant
Dim ZTEMP_VECTOR As Variant

'On Error GoTo ERROR_LABEL

TEMP_GROUP = IMPLIED_VOLATILITY_SURFACE_GAXES_FUNC(DATA_RNG, SROW, SCOLUMN)

XTEMP_VECTOR = TEMP_GROUP(1) 'X_AXIS_ARR
YTEMP_VECTOR = TEMP_GROUP(2) 'Y_AXIS_ARR
ZTEMP_ARR = TEMP_GROUP(3) 'Z_AXIS_ARR
SPLINE_ARR = TEMP_GROUP(4)
  
XI_VAL = (XTEMP_VECTOR(UBound(XTEMP_VECTOR, 1)) - XTEMP_VECTOR(1)) / (STRIKE_STEPS - 1)
YI_VAL = (YTEMP_VECTOR(UBound(YTEMP_VECTOR, 1)) - YTEMP_VECTOR(1)) / (EXPIRATION_STEPS - 1)
XT_VAL = XTEMP_VECTOR(1)
ReDim TEMP_ARR(1 To EXPIRATION_STEPS + 1, 1 To STRIKE_STEPS + 1)

ReDim XDATA_ARR(1 To STRIKE_STEPS)
ReDim YDATA_ARR(1 To EXPIRATION_STEPS)

For i = 1 To STRIKE_STEPS ' + 1
    XDATA_ARR(i) = XT_VAL
    YT_VAL = YTEMP_VECTOR(1)
    For j = 1 To EXPIRATION_STEPS 'yCases + 1
        ReDim TEMP_VECTOR(1 To UBound(ZTEMP_ARR, 1))
        For k = 1 To UBound(TEMP_VECTOR)
            TEMP_VECTOR(k) = IMPLIED_VOLATILITY_SURFACE_YAXES_FUNC(SPLINE_ARR(k, 1), SPLINE_ARR(k, 2), SPLINE_ARR(k, 3), XT_VAL)
        Next k
        ZTEMP_VECTOR = IMPLIED_VOLATILITY_SURFACE_SPLINE_FUNC(YTEMP_VECTOR, TEMP_VECTOR)
        ZT_VAL = IMPLIED_VOLATILITY_SURFACE_YAXES_FUNC(YTEMP_VECTOR, TEMP_VECTOR, ZTEMP_VECTOR, YT_VAL)
        TEMP_ARR(j + 1, i + 1) = ZT_VAL
        YDATA_ARR(j) = YT_VAL
        YT_VAL = YT_VAL + YI_VAL
    Next j
    XT_VAL = XT_VAL + XI_VAL
Next i

TEMP_ARR(1, 1) = ""
'----------------------------------------------------------------------
For i = 1 To UBound(YDATA_ARR)
    TEMP_ARR(i + 1, 1) = YDATA_ARR(i)
Next i
For i = 1 To UBound(XDATA_ARR)
    TEMP_ARR(1, i + 1) = XDATA_ARR(i)
Next i
'----------------------------------------------------------------------
IMPLIED_VOLATILITY_SURFACE_GET_FUNC = TEMP_ARR
'----------------------------------------------------------------------
  
Exit Function
ERROR_LABEL:
IMPLIED_VOLATILITY_SURFACE_GET_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : IMPLIED_VOLATILITY_SURFACE_CURVE_FUNC
'DESCRIPTION   : Get Curve Function
'LIBRARY       : DERIVATIVES
'GROUP         : IMPLIED
'ID            : 010
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 21/10/2013
'************************************************************************************
'************************************************************************************

Private Function IMPLIED_VOLATILITY_SURFACE_CURVE_FUNC(ByRef XDATA_RNG As Variant, _
ByRef YDATA_RNG As Variant, _
ByRef ZDATA_RNG As Variant, _
ByVal NSIZE As Long)
  
Dim i As Long

Dim XT_VAL As Double
Dim YT_VAL As Double
Dim XI_VAL As Double

Dim TEMP_MATRIX As Variant
Dim XDATA_VECTOR As Variant
Dim YDATA_VECTOR As Variant
Dim ZDATA_VECTOR As Variant
  
'On Error GoTo ERROR_LABEL

XDATA_VECTOR = XDATA_RNG
YDATA_VECTOR = YDATA_RNG
ZDATA_VECTOR = ZDATA_RNG

XI_VAL = (XDATA_VECTOR(UBound(XDATA_VECTOR, 1)) - XDATA_VECTOR(1)) / NSIZE
XT_VAL = XDATA_VECTOR(1)

ReDim TEMP_MATRIX(1 To NSIZE + 1, 1 To 2)

For i = 1 To NSIZE + 1
    YT_VAL = IMPLIED_VOLATILITY_SURFACE_YAXES_FUNC(XDATA_VECTOR, YDATA_VECTOR, ZDATA_VECTOR, XT_VAL)
    TEMP_MATRIX(i, 1) = XT_VAL
    TEMP_MATRIX(i, 2) = YT_VAL
    XT_VAL = XT_VAL + XI_VAL
Next i

IMPLIED_VOLATILITY_SURFACE_CURVE_FUNC = TEMP_MATRIX
   
Exit Function
ERROR_LABEL:
IMPLIED_VOLATILITY_SURFACE_CURVE_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : IMPLIED_VOLATILITY_SURFACE_GAXES_FUNC
'DESCRIPTION   : Set Axis Function
'LIBRARY       : DERIVATIVES
'GROUP         : IMPLIED
'ID            : 011
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 21/10/2013
'************************************************************************************
'************************************************************************************

Private Function IMPLIED_VOLATILITY_SURFACE_GAXES_FUNC(ByRef DATA_RNG As Variant, _
Optional ByVal SROW As Long = 2, _
Optional ByVal SCOLUMN As Long = 2)

'XTEMP_VECTOR-coordinates are before the first row
'YTEMP_VECTOR-coordinates are before the first column

Dim i As Long
Dim j As Long
Dim k As Long

Dim ii As Long
Dim jj As Long

Dim NROWS As Long
Dim NCOLUMNS As Long
  
Dim XDATA_ARR As Variant
Dim YDATA_ARR As Variant
Dim ZDATA_ARR As Variant

Dim YTEMP_ARR As Variant
Dim SPLINE_ARR As Variant

Dim TEMP_GROUP As Variant

Dim DATA_MATRIX As Variant

'On Error GoTo ERROR_LABEL

DATA_MATRIX = DATA_RNG

NROWS = UBound(DATA_MATRIX, 1)
NCOLUMNS = UBound(DATA_MATRIX, 2)

ReDim XDATA_ARR(1 To (NCOLUMNS - SCOLUMN + 1))
ReDim YDATA_ARR(1 To (NROWS - SROW + 1))

k = 1
For i = SCOLUMN To NCOLUMNS
  XDATA_ARR(k) = DATA_MATRIX(SROW - 1, i)
  k = k + 1
Next i
k = 1
For i = SROW To NROWS
  YDATA_ARR(k) = DATA_MATRIX(i, SCOLUMN - 1)
  k = k + 1
Next i

ReDim ZDATA_ARR(1 To (NROWS - SROW + 1), 1 To (NCOLUMNS - SCOLUMN + 1))
ii = 1
For i = SROW To NROWS
  jj = 1
  For j = SCOLUMN To NCOLUMNS
    ZDATA_ARR(ii, jj) = DATA_MATRIX(i, j)
    jj = jj + 1
  Next j
  ii = ii + 1
Next i

ReDim SPLINE_ARR(1 To UBound(ZDATA_ARR, 1), 1 To 3)
For i = 1 To UBound(SPLINE_ARR, 1)
  ReDim YTEMP_ARR(1 To UBound(ZDATA_ARR, 2))
    For j = 1 To UBound(ZDATA_ARR, 2)
      YTEMP_ARR(j) = ZDATA_ARR(i, j)
    Next j
    SPLINE_ARR(i, 1) = XDATA_ARR
    SPLINE_ARR(i, 2) = YTEMP_ARR
    SPLINE_ARR(i, 3) = IMPLIED_VOLATILITY_SURFACE_SPLINE_FUNC(XDATA_ARR, YTEMP_ARR)
Next i
  ReDim TEMP_GROUP(1 To 4)
  TEMP_GROUP(1) = XDATA_ARR
  TEMP_GROUP(2) = YDATA_ARR
  TEMP_GROUP(3) = ZDATA_ARR
  TEMP_GROUP(4) = SPLINE_ARR
  IMPLIED_VOLATILITY_SURFACE_GAXES_FUNC = TEMP_GROUP

Exit Function
ERROR_LABEL:
IMPLIED_VOLATILITY_SURFACE_GAXES_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : IMPLIED_VOLATILITY_SURFACE_YAXES_FUNC
'DESCRIPTION   : Y Axis Function
'LIBRARY       : DERIVATIVES
'GROUP         : IMPLIED
'ID            : 012
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 21/10/2013
'************************************************************************************
'************************************************************************************

Private Function IMPLIED_VOLATILITY_SURFACE_YAXES_FUNC(ByRef XDATA_RNG As Variant, _
ByRef YDATA_RNG As Variant, _
ByRef ZDATA_RNG As Variant, _
ByVal X_VAL As Double)
    
Dim i As Single
Dim hh As Single
Dim ii As Single
Dim jj As Single

Dim xx As Single
Dim yy As Single
Dim zz As Single

Dim XTEMP_VECTOR As Variant
Dim YTEMP_VECTOR As Variant
Dim ZTEMP_VECTOR As Variant

'On Error GoTo ERROR_LABEL

XTEMP_VECTOR = XDATA_RNG
YTEMP_VECTOR = YDATA_RNG
ZTEMP_VECTOR = ZDATA_RNG

If X_VAL < XTEMP_VECTOR(1) Then: X_VAL = XTEMP_VECTOR(1)
If X_VAL > XTEMP_VECTOR(UBound(XTEMP_VECTOR, 1)) Then: X_VAL = XTEMP_VECTOR(UBound(XTEMP_VECTOR, 1))
For i = 1 To UBound(XTEMP_VECTOR, 1)
    If XTEMP_VECTOR(i) > X_VAL Then
        xx = i - 1
        Exit For
    End If
Next i

If X_VAL = XTEMP_VECTOR(UBound(XTEMP_VECTOR, 1)) Then
    IMPLIED_VOLATILITY_SURFACE_YAXES_FUNC = YTEMP_VECTOR(UBound(YTEMP_VECTOR, 1))
    Exit Function
End If

hh = (XTEMP_VECTOR(xx + 1) - XTEMP_VECTOR(xx))
ii = (XTEMP_VECTOR(xx + 1) - X_VAL) / hh
jj = 1 - ii
yy = xx: zz = xx
IMPLIED_VOLATILITY_SURFACE_YAXES_FUNC = ii * YTEMP_VECTOR(yy) + jj * YTEMP_VECTOR(yy + 1) + ((ii * ii * ii - ii) * ZTEMP_VECTOR(zz) + (jj * jj * jj - jj) * ZTEMP_VECTOR(zz + 1)) * (hh * hh) / 6
   
Exit Function
ERROR_LABEL:
IMPLIED_VOLATILITY_SURFACE_YAXES_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : IMPLIED_VOLATILITY_SURFACE_SPLINE_FUNC
'DESCRIPTION   : Smooth Spline Function
'LIBRARY       : DERIVATIVES
'GROUP         : IMPLIED
'ID            : 013
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 21/10/2013
'************************************************************************************
'************************************************************************************

Private Function IMPLIED_VOLATILITY_SURFACE_SPLINE_FUNC(ByRef XDATA_ARR As Variant, _
ByRef YDATA_ARR As Variant)

Dim i As Long

Dim TEMP1_VAL As Double
Dim TEMP2_VAL As Double

'Dim XDATA_ARR As Variant
'Dim YDATA_ARR As Variant

Dim TEMP1_ARR As Variant
Dim TEMP2_ARR As Variant

'On Error GoTo ERROR_LABEL

'XDATA_ARR = XDATA_RNG
'YDATA_ARR = YDATA_RNG
ReDim TEMP1_ARR(1 To UBound(XDATA_ARR, 1) - 1)
ReDim TEMP2_ARR(1 To UBound(XDATA_ARR, 1))
  
TEMP1_ARR(1) = 0
TEMP2_ARR(1) = 0

For i = 2 To UBound(XDATA_ARR, 1) - 1
    TEMP1_VAL = (XDATA_ARR(i) - XDATA_ARR(i - 1)) / (XDATA_ARR(i + 1) - XDATA_ARR(i - 1))
    TEMP2_VAL = TEMP1_VAL * TEMP2_ARR(i - 1) + 2
    TEMP2_ARR(i) = (TEMP1_VAL - 1) / TEMP2_VAL
    TEMP1_ARR(i) = (YDATA_ARR(i + 1) - YDATA_ARR(i)) / (XDATA_ARR(i + 1) - XDATA_ARR(i)) - (YDATA_ARR(i) - YDATA_ARR(i - 1)) / (XDATA_ARR(i) - XDATA_ARR(i - 1))
    TEMP1_ARR(i) = (6 * TEMP1_ARR(i) / (XDATA_ARR(i + 1) - XDATA_ARR(i - 1)) - TEMP1_VAL * TEMP1_ARR(i - 1)) / TEMP2_VAL
Next i
  
TEMP2_ARR(UBound(XDATA_ARR, 1)) = 0
For i = UBound(XDATA_ARR, 1) - 1 To 1 Step -1
    TEMP2_ARR(i) = TEMP2_ARR(i) * TEMP2_ARR(i + 1) + TEMP1_ARR(i)
Next i

IMPLIED_VOLATILITY_SURFACE_SPLINE_FUNC = TEMP2_ARR

Exit Function
ERROR_LABEL:
IMPLIED_VOLATILITY_SURFACE_SPLINE_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : IMPLIED_VOLATILITY_SURFACE_BID_ASK_MEAN_FUNC
'DESCRIPTION   : Avg. Bid Ask Table
'LIBRARY       : DERIVATIVES
'GROUP         : SURFACE
'ID            : 014
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 21/10/2013
'************************************************************************************
'************************************************************************************

Function IMPLIED_VOLATILITY_SURFACE_BID_ASK_MEAN_FUNC(ByRef DATA_MATRIX As Variant)

Dim ii As Long
Dim jj As Long
  
Dim TEMP_ARR As Variant

Dim GROUP1_ARR As Variant
Dim GROUP2_ARR As Variant

Dim TEMP_MATRIX As Variant
'Dim DATA_MATRIX As Variant
  
Dim CALL_BID_VAL As Double
Dim CALL_ASK_VAL As Double
Dim PUT_BID_VAL As Double
Dim PUT_ASK_VAL As Double
    
Dim CALL_BID_TVAL As Double
Dim CALL_ASK_TVAL As Double
Dim PUT_BID_TVAL As Double
Dim PUT_ASK_TVAL As Double
  
Dim CALL_BID_INT As Long
Dim CALL_ASK_INT As Long
Dim PUT_BID_INT As Long
Dim PUT_ASK_INT As Long
  
'On Error GoTo ERROR_LABEL

'DATA_MATRIX = DATA_RNG
DATA_MATRIX = MATRIX_DOUBLE_SORT_FUNC(DATA_MATRIX)
GROUP2_ARR = IMPLIED_VOLATILITY_SURFACE_BID_ASK_QUOTES_FUNC(DATA_MATRIX)
ReDim TEMP_ARR(1 To UBound(GROUP2_ARR), 1 To 6)
For ii = 1 To UBound(GROUP2_ARR)
    GROUP1_ARR = GROUP2_ARR(ii)
    TEMP_ARR(ii, 1) = GROUP1_ARR(1)
    TEMP_ARR(ii, 2) = GROUP1_ARR(2)
    TEMP_MATRIX = GROUP1_ARR(3)
    
    CALL_BID_VAL = 0: CALL_ASK_VAL = 0
    PUT_BID_VAL = 0: PUT_ASK_VAL = 0
    CALL_BID_TVAL = 0: CALL_ASK_TVAL = 0
    PUT_BID_TVAL = 0: PUT_ASK_TVAL = 0
    CALL_BID_INT = 0: CALL_ASK_INT = 0
    PUT_BID_INT = 0: PUT_ASK_INT = 0
    
    For jj = 1 To UBound(TEMP_MATRIX, 1)
        If TEMP_MATRIX(jj, 1) > 0 Then
            CALL_BID_TVAL = CALL_BID_TVAL + TEMP_MATRIX(jj, 1)
            CALL_BID_INT = CALL_BID_INT + 1
        End If
        If TEMP_MATRIX(jj, 2) > 0 Then
            CALL_ASK_TVAL = CALL_ASK_TVAL + TEMP_MATRIX(jj, 2)
            CALL_ASK_INT = CALL_ASK_INT + 1
        End If
        If TEMP_MATRIX(jj, 3) > 0 Then
            PUT_BID_TVAL = PUT_BID_TVAL + TEMP_MATRIX(jj, 3)
            PUT_BID_INT = PUT_BID_INT + 1
        End If
        If TEMP_MATRIX(jj, 4) > 0 Then
            PUT_ASK_TVAL = PUT_ASK_TVAL + TEMP_MATRIX(jj, 4)
            PUT_ASK_INT = PUT_ASK_INT + 1
        End If
    Next jj
    
    If CALL_BID_INT > 0 Then: CALL_BID_VAL = CALL_BID_TVAL / CALL_BID_INT
    If CALL_ASK_INT > 0 Then: CALL_ASK_VAL = CALL_ASK_TVAL / CALL_ASK_INT
    If PUT_BID_INT > 0 Then: PUT_BID_VAL = PUT_BID_TVAL / PUT_BID_INT
    If PUT_ASK_INT > 0 Then: PUT_ASK_VAL = PUT_ASK_TVAL / PUT_ASK_INT
    TEMP_ARR(ii, 3) = CALL_BID_VAL
    TEMP_ARR(ii, 4) = CALL_ASK_VAL
    
    TEMP_ARR(ii, 5) = PUT_BID_VAL
    TEMP_ARR(ii, 6) = PUT_ASK_VAL
Next ii

IMPLIED_VOLATILITY_SURFACE_BID_ASK_MEAN_FUNC = TEMP_ARR

Exit Function
ERROR_LABEL:
IMPLIED_VOLATILITY_SURFACE_BID_ASK_MEAN_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : IMPLIED_VOLATILITY_SURFACE_BID_ASK_QUOTES_FUNC
'DESCRIPTION   : Aggregate Quotes
'LIBRARY       : DERIVATIVES
'GROUP         : SURFACE
'ID            : 015
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 21/10/2013
'************************************************************************************
'************************************************************************************

Function IMPLIED_VOLATILITY_SURFACE_BID_ASK_QUOTES_FUNC(ByRef DATA_MATRIX As Variant)
  
Dim i As Long
Dim j As Long
  
Dim ii As Long
Dim jj As Long
  
Dim SROW As Long
Dim NROWS As Long
  
Dim STRIKE As Double
Dim MATURITY As Date
  
Dim TEMP_ARR As Variant
Dim TEMP_VECTOR As Variant
Dim TEMP_MATRIX As Variant
'Dim DATA_MATRIX As Variant
  
Dim GROUP1_ARR As Variant
Dim GROUP2_ARR As Variant
  
'On Error GoTo ERROR_LABEL
  
'DATA_MATRIX = DATA_RNG
  
SROW = LBound(DATA_MATRIX, 1)
NROWS = UBound(DATA_MATRIX, 1)
  
MATURITY = DATA_MATRIX(1, 1)
STRIKE = DATA_MATRIX(1, 2)
  
ReDim GROUP1_ARR(1 To 3)
GROUP1_ARR(1) = MATURITY
GROUP1_ARR(2) = STRIKE
  
ReDim TEMP_MATRIX(1 To 1, 1 To 4)
TEMP_MATRIX(1, 1) = DATA_MATRIX(1, 3)
TEMP_MATRIX(1, 2) = DATA_MATRIX(1, 4)
TEMP_MATRIX(1, 3) = DATA_MATRIX(1, 5)
TEMP_MATRIX(1, 4) = DATA_MATRIX(1, 6)
GROUP1_ARR(3) = TEMP_MATRIX

ReDim GROUP2_ARR(1 To 1) 'Aggregate
GROUP2_ARR(1) = GROUP1_ARR
j = 1

For i = (SROW + 1) To NROWS 'update the option quotes mat
    MATURITY = DATA_MATRIX(i, 1)
    STRIKE = DATA_MATRIX(i, 2)
    
    If ((MATURITY = DATA_MATRIX(i - 1, 1)) And (STRIKE = DATA_MATRIX(i - 1, 2))) Then
      GROUP1_ARR = GROUP2_ARR(j)
      TEMP_MATRIX = GROUP1_ARR(3)
      
      ReDim TEMP_VECTOR(1 To 4, 1 To 1)
      TEMP_VECTOR(1, 1) = DATA_MATRIX(i, 3)
      TEMP_VECTOR(2, 1) = DATA_MATRIX(i, 4)
      TEMP_VECTOR(3, 1) = DATA_MATRIX(i, 5)
      TEMP_VECTOR(4, 1) = DATA_MATRIX(i, 6)
    
      ReDim TEMP_ARR(1 To UBound(TEMP_MATRIX, 1) + 1, 1 To UBound(TEMP_MATRIX, 2))
      'Insert Matrix Row
        
      For ii = 1 To UBound(TEMP_MATRIX, 1)
        For jj = 1 To UBound(TEMP_MATRIX, 2)
            TEMP_ARR(ii, jj) = TEMP_MATRIX(ii, jj)
        Next jj
      Next ii
        
      For jj = 1 To UBound(TEMP_MATRIX, 2)
        TEMP_ARR(ii, jj) = TEMP_VECTOR(jj, 1)
      Next jj
        
      TEMP_MATRIX = TEMP_ARR
      GROUP1_ARR(3) = TEMP_MATRIX
      GROUP2_ARR(j) = GROUP1_ARR
    Else
      ReDim GROUP1_ARR(1 To 3)
      GROUP1_ARR(1) = MATURITY
      GROUP1_ARR(2) = STRIKE
      
      ReDim TEMP_MATRIX(1 To 1, 1 To 4)
      TEMP_MATRIX(1, 1) = DATA_MATRIX(i, 3)
      TEMP_MATRIX(1, 2) = DATA_MATRIX(i, 4)
      TEMP_MATRIX(1, 3) = DATA_MATRIX(i, 5)
      TEMP_MATRIX(1, 4) = DATA_MATRIX(i, 6)
          
      GROUP1_ARR(3) = TEMP_MATRIX
      j = j + 1
      ReDim Preserve GROUP2_ARR(1 To j)
      GROUP2_ARR(j) = GROUP1_ARR
    End If
Next i
  
IMPLIED_VOLATILITY_SURFACE_BID_ASK_QUOTES_FUNC = GROUP2_ARR

Exit Function
ERROR_LABEL:
IMPLIED_VOLATILITY_SURFACE_BID_ASK_QUOTES_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : IMPLIED_VOLATILITY_SURFACE_BID_ASK_EXTRACT_FUNC
'DESCRIPTION   : Extract Bid Ask Values
'LIBRARY       : DERIVATIVES
'GROUP         : SURFACE
'ID            : 016
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 21/10/2013
'************************************************************************************
'************************************************************************************

Function IMPLIED_VOLATILITY_SURFACE_BID_ASK_EXTRACT_FUNC(ByRef DATA_RNG As Variant, _
Optional ByVal VERSION As Integer = 0, _
Optional ByVal COL_CALL_BID As Long = 4, _
Optional ByVal COL_CALL_ASK As Long = 5, _
Optional ByVal COL_PUT_BID As Long = 11, _
Optional ByVal COL_PUT_ASK As Long = 12, _
Optional ByVal COL_MATURITY As Long = 2, _
Optional ByVal COL_STRIKE As Long = 8, _
Optional ByVal SROW As Long = 3)

Dim i As Long
Dim j As Long
Dim NROWS As Long

Dim STRIKE As Double
Dim MATURITY As Date

Dim DATA_MATRIX As Variant
Dim TEMP_MATRIX As Variant
Dim TEMP_VECTOR As Variant

'On Error GoTo ERROR_LABEL

DATA_MATRIX = DATA_RNG
NROWS = UBound(DATA_MATRIX, 1)

'----------------------------------------------------------------------------------------------
Select Case VERSION
'----------------------------------------------------------------------------------------------
Case 0 'CBOE
'----------------------------------------------------------------------------------------------
    ReDim TEMP_MATRIX(1 To NROWS - SROW, 1 To 6)
    
    j = 1
    For i = (SROW + 1) To NROWS
      TEMP_VECTOR = IMPLIED_VOLATILITY_SURFACE_PARSE_FUNC(DATA_MATRIX(i, 1))
      
      MATURITY = TEMP_VECTOR(1, 1)
      STRIKE = TEMP_VECTOR(2, 1)
      
      TEMP_MATRIX(j, 1) = MATURITY
      TEMP_MATRIX(j, 2) = STRIKE
      TEMP_MATRIX(j, 3) = CDec(DATA_MATRIX(i, COL_CALL_BID))
      TEMP_MATRIX(j, 4) = CDec(DATA_MATRIX(i, COL_CALL_ASK))
      TEMP_MATRIX(j, 5) = CDec(DATA_MATRIX(i, COL_PUT_BID))
      TEMP_MATRIX(j, 6) = CDec(DATA_MATRIX(i, COL_PUT_ASK))
      
      j = j + 1
    Next i
'----------------------------------------------------------------------------------------------
Case Else 'Yahoo
'----------------------------------------------------------------------------------------------
    ReDim TEMP_MATRIX(1 To NROWS, 1 To 6)
    
    j = 1
    For i = 1 To NROWS
      
      TEMP_MATRIX(j, 1) = CDate(DATA_MATRIX(i, COL_MATURITY))
      TEMP_MATRIX(j, 2) = CDec(DATA_MATRIX(i, COL_STRIKE))
      TEMP_MATRIX(j, 3) = CDec(DATA_MATRIX(i, COL_CALL_BID))
      TEMP_MATRIX(j, 4) = CDec(DATA_MATRIX(i, COL_CALL_ASK))
      TEMP_MATRIX(j, 5) = CDec(DATA_MATRIX(i, COL_PUT_BID))
      TEMP_MATRIX(j, 6) = CDec(DATA_MATRIX(i, COL_PUT_ASK))
      
      j = j + 1
    Next i
'----------------------------------------------------------------------------------------------
End Select
'----------------------------------------------------------------------------------------------

IMPLIED_VOLATILITY_SURFACE_BID_ASK_EXTRACT_FUNC = TEMP_MATRIX

Exit Function
ERROR_LABEL:
IMPLIED_VOLATILITY_SURFACE_BID_ASK_EXTRACT_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : IMPLIED_VOLATILITY_SURFACE_MATRIX_GROUP_FUNC
'DESCRIPTION   : Aggregate Vector & Matrix --> Creating a Group of Data
'LIBRARY       : MATRIX
'GROUP         : GROUP
'ID            : 017
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 10/21/2013
'************************************************************************************
'************************************************************************************

Function IMPLIED_VOLATILITY_SURFACE_MATRIX_GROUP_FUNC(ByRef KEY_VECTOR As Variant, _
ByRef DATA_MATRIX As Variant, _
Optional ByVal OUTPUT As Integer = 0)
  
Dim h As Long
Dim i As Long
Dim j As Long
Dim k As Long '--> Temp Index No.
Dim l As Long

Dim NSIZE As Long

Dim KEY_VAL As Variant 'key for indexing
  
Dim TEMP1_ARR As Variant
Dim TEMP2_ARR As Variant 'the row of data to be indexed

Dim TEMP_GROUP As Variant
Dim TEMP_MATRIX As Variant

Dim KEY_GROUP_ARR As Variant
Dim DATA_GROUP_ARR As Variant

On Error GoTo ERROR_LABEL
  
ReDim DATA_GROUP_ARR(0 To 0)
ReDim KEY_GROUP_ARR(0 To 0)

'------------------------------------------------------------------------------
For i = 1 To UBound(KEY_VECTOR, 1)
'------------------------------------------------------------------------------
    KEY_VAL = KEY_VECTOR(i) '---> ONE DIMENSION ARRAY
    
    ReDim TEMP1_ARR(1 To UBound(DATA_MATRIX, 2)) 'ReadMatrixRowIntoArray
    For j = 1 To UBound(DATA_MATRIX, 2)
      TEMP1_ARR(j) = DATA_MATRIX(i, j)
    Next j
    
    TEMP2_ARR = TEMP1_ARR
    k = -1
    For j = 1 To UBound(KEY_GROUP_ARR, 1) 'Find Group
      If KEY_VAL = KEY_GROUP_ARR(j) Then
        k = j
        GoTo 1983
      End If
    Next j
1983:
'------------------------------------------------------------------------------
    If k > 0 Then 'InsertRowInGroup
'------------------------------------------------------------------------------
      TEMP_GROUP = DATA_GROUP_ARR(k)
      If UBound(TEMP_GROUP, 1) > 0 Then 'InsertRowInMatrix
            ReDim TEMP1_ARR(1 To UBound(TEMP_GROUP, 1) + 1, 1 To UBound(TEMP_GROUP, 2))
      Else
            ReDim TEMP1_ARR(1 To 1, 1 To UBound(TEMP_GROUP, 2))
      End If
      
      For h = 1 To UBound(TEMP_GROUP, 1)
          For l = 1 To UBound(TEMP_GROUP, 2)
            TEMP1_ARR(h, l) = TEMP_GROUP(h, l)
          Next l
      Next h
      
      For l = 1 To UBound(TEMP2_ARR, 1)
            TEMP1_ARR(UBound(TEMP_GROUP, 1) + 1, l) = TEMP2_ARR(l)
      Next l
      
      TEMP_GROUP = TEMP1_ARR
      DATA_GROUP_ARR(k) = TEMP_GROUP
'------------------------------------------------------------------------------
    Else 'Add Group
'------------------------------------------------------------------------------
        NSIZE = UBound(DATA_GROUP_ARR, 1)
        
        ReDim Preserve DATA_GROUP_ARR(1 To NSIZE + 1)
        ReDim Preserve KEY_GROUP_ARR(1 To NSIZE + 1)
        ReDim TEMP_MATRIX(0 To 0, 1 To UBound(TEMP2_ARR, 1))
        
        KEY_GROUP_ARR(NSIZE + 1) = KEY_VAL
            
        If UBound(TEMP_MATRIX, 1) > 0 Then 'InsertRowInMatrix
              ReDim TEMP1_ARR(1 To UBound(TEMP_MATRIX, 1) + 1, 1 To UBound(TEMP_MATRIX, 2))
        Else
              ReDim TEMP1_ARR(1 To 1, 1 To UBound(TEMP_MATRIX, 2))
        End If
        
        For h = 1 To UBound(TEMP_MATRIX, 1)
          For l = 1 To UBound(TEMP_MATRIX, 2)
            TEMP1_ARR(h, l) = TEMP_MATRIX(h, l)
          Next l
        Next h
        
        For l = 1 To UBound(TEMP2_ARR, 1)
            TEMP1_ARR(UBound(TEMP_MATRIX, 1) + 1, l) = TEMP2_ARR(l)
        Next l
        
        DATA_GROUP_ARR(NSIZE + 1) = TEMP1_ARR
'------------------------------------------------------------------------------
    End If
'------------------------------------------------------------------------------
Next i
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
Select Case OUTPUT
'------------------------------------------------------------------------------
Case 0
'------------------------------------------------------------------------------
    IMPLIED_VOLATILITY_SURFACE_MATRIX_GROUP_FUNC = DATA_GROUP_ARR
'------------------------------------------------------------------------------
Case 1
'------------------------------------------------------------------------------
    IMPLIED_VOLATILITY_SURFACE_MATRIX_GROUP_FUNC = KEY_GROUP_ARR
'------------------------------------------------------------------------------
Case Else
'------------------------------------------------------------------------------
    ReDim TEMP_GROUP(1 To 2)
    TEMP_GROUP(1) = DATA_GROUP_ARR
    TEMP_GROUP(2) = KEY_GROUP_ARR
    
    IMPLIED_VOLATILITY_SURFACE_MATRIX_GROUP_FUNC = TEMP_GROUP
'------------------------------------------------------------------------------
End Select
'------------------------------------------------------------------------------

Exit Function
ERROR_LABEL:
IMPLIED_VOLATILITY_SURFACE_MATRIX_GROUP_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : IMPLIED_VOLATILITY_SURFACE_PARSE_FUNC
'DESCRIPTION   : Parse Third Friday Function (weekday of thursday is 5)
'LIBRARY       : DERIVATIVES
'GROUP         : SURFACE
'ID            : 018
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 21/10/2013
'************************************************************************************
'************************************************************************************

Function IMPLIED_VOLATILITY_SURFACE_PARSE_FUNC(ByVal DATE_STR As String)
  
Dim i As Long
Dim j As Long
Dim k As Long

Dim NSIZE As Long
Dim MONTH_STR As String
Dim YEAR_STR As String
Dim TEMP_STR As String
Dim TEMP_VECTOR As Variant

'On Error GoTo ERROR_LABEL

ReDim TEMP_VECTOR(1 To 2, 1 To 1)
  
'------------------------------------------------------------------------------
MONTH_STR = Mid(DATE_STR, 4, 3)
YEAR_STR = "20" & Left(DATE_STR, 2)
For i = 1 To 7
    TEMP_STR = MONTH_STR & " 0" & i & "," & YEAR_STR
    NSIZE = Weekday(CDate(TEMP_STR))
    If NSIZE = 6 Then Exit For
Next i
j = 14 + i
TEMP_STR = MONTH_STR & " " & CStr(j) & "," & YEAR_STR
TEMP_VECTOR(1, 1) = CDate(TEMP_STR)
'------------------------------------------------------------------------------
k = InStr(1, DATE_STR, "(", vbTextCompare)
If k = 0 Then: GoTo ERROR_LABEL
TEMP_VECTOR(2, 1) = CDec(Mid(DATE_STR, 7, k - 7)) 'Out STRIKE
'------------------------------------------------------------------------------
IMPLIED_VOLATILITY_SURFACE_PARSE_FUNC = TEMP_VECTOR

Exit Function
ERROR_LABEL:
IMPLIED_VOLATILITY_SURFACE_PARSE_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : IMPLIED_VOLATILITY_SURFACE_SPOT_PRICE_FUNC
'DESCRIPTION   : OPTION SPOT PRICE
'LIBRARY       : DERIVATIVES
'GROUP         : SURFACE
'ID            : 019
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 21/10/2013
'************************************************************************************
'************************************************************************************

Function IMPLIED_VOLATILITY_SURFACE_SPOT_PRICE_FUNC(ByVal FULL_PATH_NAME As String)
'On Error GoTo ERROR_LABEL
IMPLIED_VOLATILITY_SURFACE_SPOT_PRICE_FUNC = _
    CDec(CONVERT_TEXT_FILE_MATRIX_FUNC(FULL_PATH_NAME, 5, 4, ",")(1, 2))
Exit Function
ERROR_LABEL:
IMPLIED_VOLATILITY_SURFACE_SPOT_PRICE_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : IMPLIED_VOLATILITY_SURFACE_SETTLEMENT_FUNC
'DESCRIPTION   : OPT SETTLEMENT DATE
'LIBRARY       : DERIVATIVES
'GROUP         : SURFACE
'ID            : 020
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 21/10/2013
'************************************************************************************
'************************************************************************************

Function IMPLIED_VOLATILITY_SURFACE_SETTLEMENT_FUNC(ByVal FULL_PATH_NAME As String)
Dim DATE_STR As String

'On Error GoTo ERROR_LABEL
DATE_STR = CONVERT_TEXT_FILE_MATRIX_FUNC(FULL_PATH_NAME, 5, 4, ",")(2, 1)
'Jun 07 2006 @ 19:24 ET (Data 20 Minutes Delayed)
DATE_STR = Left(DATE_STR, 6) & "," & Mid(DATE_STR, 8, 4)
IMPLIED_VOLATILITY_SURFACE_SETTLEMENT_FUNC = CDate(DATE_STR)

Exit Function
ERROR_LABEL:
IMPLIED_VOLATILITY_SURFACE_SETTLEMENT_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : IMPLIED_VOLATILITY_SURFACE_RESET_FUNC
'DESCRIPTION   : Reset Implied Surface Variables
'LIBRARY       : DERIVATIVES
'GROUP         : IMPLIED
'ID            : 021
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 21/10/2013
'************************************************************************************
'************************************************************************************

Private Function IMPLIED_VOLATILITY_SURFACE_RESET_FUNC()

'On Error GoTo ERROR_LABEL

IMPLIED_VOLATILITY_SURFACE_RESET_FUNC = False

PUB_TARGET_VAL = 0

PUB_SPOT_VAL = 0
PUB_RATE_VAL = 0
PUB_SIGMA_VAL = 0 'DefaultVolatility

PUB_STRIKE_VAL = 0
PUB_EXPIRATION_VAL = 0
PUB_OPTION_FLAG = 0

PUB_CND_TYPE = 0

PUB_TENOR_ARR = 0
PUB_STRIKE_ARR = 0

PUB_IMPL_BID_ARR = 0
PUB_IMPL_ASK_ARR = 0
PUB_IMPL_SIGMA_ARR = 0

PUB_GROUP_ARR = 0
PUB_3D_MATRIX = 0
PUB_INDEX_ARR = 0

IMPLIED_VOLATILITY_SURFACE_RESET_FUNC = True

Exit Function
ERROR_LABEL:
IMPLIED_VOLATILITY_SURFACE_RESET_FUNC = False
End Function
