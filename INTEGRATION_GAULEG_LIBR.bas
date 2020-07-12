Attribute VB_Name = "INTEGRATION_GAULEG_LIBR"

'-------------------------------------------------------------------------------------
'-------------------------------------------------------------------------------------

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.
'-------------------------------------------------------------------------------------
'-------------------------------------------------------------------------------------


'************************************************************************************
'************************************************************************************
'FUNCTION      : GAULEG7_INTEGRATION_FUNC
'DESCRIPTION   : Find the area under the curvey = f (x), -1 = x = 1.
'weights for 7-point Gauss-Legendre integration
'(only 4 values out of 7 are given as they are symmetric)
'LIBRARY       : INTEGRATION
'GROUP         : GAULEG
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 21/01/2009
'************************************************************************************
'************************************************************************************

Function GAULEG7_INTEGRATION_FUNC(ByVal FUNC_NAME_STR As String, _
ByVal LOWER_VAL As Double, _
ByVal UPPER_VAL As Double, _
ByVal tolerance As Double, _
Optional ByVal nLOOPS As Long = 30)
  
' weights for 7-point Gauss-Legendre integration
' (only 4 values out of 7 are given as they are symmetric)

Dim i As Single
Dim j As Single

Dim k As Long

Dim C_VAL As Double
Dim H_VAL As Double

Dim TEMP1_VAL As Double 'TEMP1_VAL (abscissa) and f(TEMP1_VAL)
Dim TEMP2_VAL As Double 'will be result of TEMP2_VAL integral
Dim CTEMP_VAL As Double 'will be result of CTEMP_VAL integral

Dim TEMP_SUM As Double
Dim FUNC_VAL As Double

Dim TEMP1_ARR() As Double
Dim TEMP2_ARR() As Double
Dim TEMP3_ARR() As Double

On Error GoTo ERROR_LABEL

k = 0
ReDim TEMP1_ARR(0 To 3)

TEMP1_ARR(0) = 0.417959183673469
TEMP1_ARR(1) = 0.381830050505119
TEMP1_ARR(2) = 0.279705391489277
TEMP1_ARR(3) = 0.12948496616887

ReDim TEMP2_ARR(0 To 7) ' weights for 15-point Gauss-Kronrod integration

TEMP2_ARR(0) = 0.209482141084728
TEMP2_ARR(1) = 0.204432940075298
TEMP2_ARR(2) = 0.190350578064785
TEMP2_ARR(3) = 0.169004726639267
TEMP2_ARR(4) = 0.140653259715525
TEMP2_ARR(5) = 0.10479001032225
TEMP2_ARR(6) = 0.063092092629979
TEMP2_ARR(7) = 0.022935322010529

ReDim TEMP3_ARR(0 To 7) ' abscissae (evaluation points) for 15-point
'Gauss-Kronrod integration

TEMP3_ARR(0) = 0#
TEMP3_ARR(1) = 0.207784955007898
TEMP3_ARR(2) = 0.405845151377397
TEMP3_ARR(3) = 0.586087235467691
TEMP3_ARR(4) = 0.741531185599394
TEMP3_ARR(5) = 0.864864423359769
TEMP3_ARR(6) = 0.949107912342758
TEMP3_ARR(7) = 0.991455371120813

H_VAL = (UPPER_VAL - LOWER_VAL) / 2
C_VAL = (LOWER_VAL + UPPER_VAL) / 2
FUNC_VAL = Excel.Application.Run(FUNC_NAME_STR, C_VAL)
TEMP2_VAL = FUNC_VAL * TEMP1_ARR(0)
CTEMP_VAL = FUNC_VAL * TEMP2_ARR(0)

' calculate TEMP2_VAL and half of CTEMP_VAL
i = 2
For j = 1 To 3
  TEMP1_VAL = H_VAL * TEMP3_ARR(i)
  TEMP_SUM = Excel.Application.Run(FUNC_NAME_STR, (C_VAL - TEMP1_VAL)) + Excel.Application.Run(FUNC_NAME_STR, (C_VAL + TEMP1_VAL))
  TEMP2_VAL = TEMP2_VAL + TEMP_SUM * TEMP1_ARR(j)
  CTEMP_VAL = CTEMP_VAL + TEMP_SUM * TEMP2_ARR(i)
  i = i + 2
Next j

' calculate other half of CTEMP_VAL
For i = 1 To 7 Step 2
    TEMP1_VAL = H_VAL * TEMP3_ARR(i)
    TEMP_SUM = Excel.Application.Run(FUNC_NAME_STR, (C_VAL - TEMP1_VAL)) + Excel.Application.Run(FUNC_NAME_STR, (C_VAL + TEMP1_VAL))
    CTEMP_VAL = CTEMP_VAL + TEMP_SUM * TEMP2_ARR(i)
Next i

' multiply by (LOWER_VAL - UPPER_VAL) / 2
TEMP2_VAL = H_VAL * TEMP2_VAL
CTEMP_VAL = H_VAL * CTEMP_VAL

' 15 more function evaluations have been used
k = k + 15

' error is <= CTEMP_VAL - TEMP2_VAL
' if error is larger than tolerance then split the interval
' in two and integrate recursively
If (Abs(CTEMP_VAL - TEMP2_VAL) < tolerance) Then
   GAULEG7_INTEGRATION_FUNC = CTEMP_VAL
Else
    If k + 30 > nLOOPS Then: GoTo ERROR_LABEL 'maximum number of function evaluations exceeded"
   GAULEG7_INTEGRATION_FUNC = GAULEG7_INTEGRATION_FUNC(FUNC_NAME_STR, LOWER_VAL, C_VAL, tolerance / 2, nLOOPS) + GAULEG7_INTEGRATION_FUNC(FUNC_NAME_STR, C_VAL, UPPER_VAL, tolerance / 2, nLOOPS)
End If

Exit Function
ERROR_LABEL:
GAULEG7_INTEGRATION_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : GAULEG8_INTEGRATION_FUNC
'DESCRIPTION   : Gauss-Legendre performed over 8 subintervals
'LIBRARY       : INTEGRATION
'GROUP         : GAULEG
'ID            : 002
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 21/01/2009
'************************************************************************************
'************************************************************************************

Function GAULEG8_INTEGRATION_FUNC(ByVal FORMULA_STR As String, _
ByVal VARIABLE_STR As String, _
ByVal LOWER_VAL As Double, _
ByVal UPPER_VAL As Double, _
Optional ByVal RANK_VAL As Long = 8)

'-------------------------------------------------------------------
'a = 0
'b = 1

'Debug.Print GAULEG8_INTEGRATION_FUNC("sin(x)", "x", a, b, 8)
'Debug.Print " 0.4596976941318603 (exact value)"

'Debug.Print GAULEG8_INTEGRATION_FUNC("exp(@x)", "@x", a, b, 8)
'Debug.Print " 1.7182818284590452 (exact value)"
'-------------------------------------------------------------------

'RANK_VAL--> SubInterval: change as you like to any positive integer

Dim i As Long   ' COUNTER
Dim j As Long   ' COUNTER
Dim k As Long   ' degrees

Dim A_VAL As Double ' bounds to give +- 1 by change of variables
Dim B_VAL As Double
Dim L_VAL As Double ' of length (UPPER_VAL-LOWER_VAL)/pieces
Dim V_VAL As Double ' integration variable
Dim T_VAL As Double  ' summing over subintervals
Dim Y_VAL As Double

Dim TEMP_SUM As Double  ' to sum up
Dim TEMP_STR As String

Dim TEMP1_VAL As Double      ' subinterval TEMP1 ... TEMP2
Dim TEMP2_VAL As Double

Dim Z_ARR() As Double
Dim W_ARR() As Double

On Error GoTo ERROR_LABEL

'--------------------------------------------------------------------------
k = 16
'--------------------------------------------------------------------------
ReDim Z_ARR(1 To k) ' Gauss zeros
ReDim W_ARR(1 To k) ' Gauss weights
'--------------------------------------------------------------------------
'----Pre-computed data for Gauss-Legendre integration, do not change this--
'--------------------------------------------------------------------------

Z_ARR(1) = -0.98940093499165
Z_ARR(2) = -0.944575023073233
Z_ARR(3) = -0.865631202387832
Z_ARR(4) = -0.755404408355003
Z_ARR(5) = -0.617876244402644
Z_ARR(6) = -0.458016777657227
Z_ARR(7) = -0.281603550779259
Z_ARR(8) = -9.50125098376374E-02
Z_ARR(9) = 9.50125098376374E-02
Z_ARR(10) = 0.281603550779259
Z_ARR(11) = 0.458016777657227
Z_ARR(12) = 0.617876244402644
Z_ARR(13) = 0.755404408355003
Z_ARR(14) = 0.865631202387832
Z_ARR(15) = 0.944575023073233
Z_ARR(16) = 0.98940093499165
'-------------------------------------------------------------------------
W_ARR(1) = 2.71524594117541E-02
W_ARR(2) = 6.22535239386479E-02
W_ARR(3) = 9.51585116824928E-02
W_ARR(4) = 0.124628971255534
W_ARR(5) = 0.149595988816577
W_ARR(6) = 0.169156519395003
W_ARR(7) = 0.182603415044924
W_ARR(8) = 0.189450610455069
W_ARR(9) = 0.189450610455069
W_ARR(10) = 0.182603415044924
W_ARR(11) = 0.169156519395003
W_ARR(12) = 0.149595988816577
W_ARR(13) = 0.124628971255534
W_ARR(14) = 9.51585116824928E-02
W_ARR(15) = 6.22535239386479E-02
W_ARR(16) = 2.71524594117541E-02
'-------------------------------------------------------------------------

L_VAL = (UPPER_VAL - LOWER_VAL) / RANK_VAL
TEMP1_VAL = LOWER_VAL - L_VAL
TEMP2_VAL = TEMP1_VAL + L_VAL

T_VAL = 0
For j = 1 To RANK_VAL
    TEMP1_VAL = TEMP1_VAL + L_VAL
    TEMP2_VAL = TEMP1_VAL + L_VAL
    
    A_VAL = (TEMP2_VAL - TEMP1_VAL) / 2
    B_VAL = (TEMP2_VAL + TEMP1_VAL) / 2
    TEMP_SUM = 0
    For i = 1 To k
        V_VAL = A_VAL * Z_ARR(i) + B_VAL  ' change of variables for integration bounds
        GoSub CALC_LINE
        TEMP_SUM = TEMP_SUM + W_ARR(i) * Y_VAL
    Next i
    T_VAL = T_VAL + TEMP_SUM * A_VAL _
    'change of variables for integration bounds
Next j

GAULEG8_INTEGRATION_FUNC = T_VAL

Exit Function
CALC_LINE:
    TEMP_STR = Replace(CStr(V_VAL), ",", ".") ' xlDecimalSeparator
    TEMP_STR = Replace(FORMULA_STR, VARIABLE_STR, TEMP_STR)
    Y_VAL = Excel.Application.Evaluate(TEMP_STR)
Return
ERROR_LABEL:
GAULEG8_INTEGRATION_FUNC = Err.number
End Function
