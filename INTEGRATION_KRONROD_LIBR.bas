Attribute VB_Name = "INTEGRATION_KRONROD_LIBR"


'-------------------------------------------------------------------------------------
'-------------------------------------------------------------------------------------

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.
'-------------------------------------------------------------------------------------
'-------------------------------------------------------------------------------------

'************************************************************************************
'************************************************************************************
'FUNCTION      : GAUSS_KRONROD_INTEGRATION_FUNC
'DESCRIPTION   : 1-Dimensional adaptive Gauss Kronrod integration
'LIBRARY       : INTEGRATION
'GROUP         : KRONROD
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 21/01/2009
'************************************************************************************
'************************************************************************************

Function GAUSS_KRONROD_INTEGRATION_FUNC(ByVal FUNC_NAME_STR As Variant, _
ByVal LOWER_VAL As Double, _
ByVal UPPER_VAL As Double, _
Optional ByVal nLOOPS As Integer = 400, _
Optional ByVal tolerance As Double = 10 ^ -14)

'nLOOPS --> maximal number of subdivisions

Dim i As Integer
Dim j As Integer
Dim k As Integer
Dim h As Integer

Dim TEMP1_ARR(0 To 5) As Double
Dim TEMP2_ARR(0 To 11) As Double
Dim TEMP3_ARR(0 To 11) As Double
Dim TEMP4_ARR() As Double
Dim TEMP5_ARR() As Double
Dim TEMP6_ARR() As Double
Dim TEMP7_ARR() As Double

Dim TEMP_FIN As Double
Dim TEMP_ERR As Double
Dim TEMP_OLD As Double

Dim TEMP_LOWER As Double
Dim TEMP_UPPER As Double
Dim TEMP_HALF As Double

Dim TEMP_FACTOR As Double
Dim TEMP_CENTRE As Double
Dim TEMP_ABSCIS As Double

Dim TEMP1_SUM As Double
Dim TEMP2_SUM As Double
Dim TEMP3_SUM As Double

Dim KRON_INT As Double
Dim KRON_ERR As Double

Dim COUNTER As Long     ' function evaluations

On Error GoTo ERROR_LABEL

ReDim TEMP6_ARR(0 To nLOOPS)
ReDim TEMP7_ARR(0 To nLOOPS)
ReDim TEMP4_ARR(0 To nLOOPS)
ReDim TEMP5_ARR(0 To nLOOPS)

'--------------------------------------------------------------------------------
'--------------------------------------------------------------------------------
'What numerical integration technique would be best for a definite
'integral, evaluated over a finite interval.

'* Definite integral over a finite interval, z = z1 to z2.
'* The integrand is continuous over the interval, as are its
'  derivatives.

'* The integrand is well-behaved in magnitude over MOST of the
'  interval.

'* Over a TINY part of the interval, it makes some oscillations
'  many orders of magnitude above the values along the rest of the
'  interval.  For example, over 99% of the interval the integrand
'  varies smoothly with values between +/- 10e5.  In a tiny piece of
'  the interval it makes a sudden oscillation between +/- 10e11.
'--------------------------------------------------------------------------------
'--------------------------------------------------------------------------------


GoSub 1984 ' initialize Gauss points

k = 1
j = 1
TEMP4_ARR(1) = LOWER_VAL
TEMP5_ARR(1) = UPPER_VAL

COUNTER = 0

Do
  j = j + 1
  TEMP5_ARR(j) = TEMP5_ARR(k)
  TEMP4_ARR(j) = (TEMP4_ARR(k) + TEMP5_ARR(k)) / 2#
  TEMP5_ARR(k) = TEMP4_ARR(j)
  TEMP_LOWER = TEMP4_ARR(k)
  TEMP_UPPER = TEMP5_ARR(k)
  
  GoSub 1983
  
  TEMP7_ARR(k) = KRON_INT
  TEMP6_ARR(k) = KRON_ERR
  TEMP_LOWER = TEMP4_ARR(j)
  TEMP_UPPER = TEMP5_ARR(j)
  
  GoSub 1983
  
  TEMP7_ARR(j) = KRON_INT
  TEMP6_ARR(j) = KRON_ERR
  TEMP_ERR = 0
  TEMP_FIN = 0
  
  For i = 1 To j
    If TEMP6_ARR(i) > TEMP6_ARR(k) Then k = i
      TEMP_FIN = TEMP_FIN + TEMP7_ARR(i)
      TEMP_ERR = TEMP_ERR + TEMP6_ARR(i) ^ 2
  Next i
  TEMP_ERR = Sqr(TEMP_ERR)
  
  If Abs(TEMP_OLD - TEMP_FIN) < tolerance / 10 Then
    TEMP_OLD = TEMP_FIN 'Exit Do
  Else: TEMP_OLD = TEMP_FIN
  End If
  
Loop While 4 * TEMP_ERR > tolerance And j < nLOOPS

GAUSS_KRONROD_INTEGRATION_FUNC = TEMP_FIN
Exit Function

'---------------------------------------------------------------------
1983:  'Kronrod Rule on interval [TEMP_LOWER,TEMP_UPPER]
    TEMP_HALF = (TEMP_UPPER - TEMP_LOWER) / 2#
    TEMP_CENTRE = (TEMP_LOWER + TEMP_UPPER) / 2#
    TEMP_FACTOR = Excel.Application.Run(FUNC_NAME_STR, TEMP_CENTRE)
    TEMP3_SUM = TEMP_FACTOR * TEMP1_ARR(0)
    TEMP2_SUM = TEMP_FACTOR * TEMP2_ARR(0)
    For h = 1 To 11
        TEMP_ABSCIS = TEMP_HALF * TEMP3_ARR(h)
        TEMP1_SUM = Excel.Application.Run(FUNC_NAME_STR, TEMP_CENTRE - TEMP_ABSCIS) + Excel.Application.Run(FUNC_NAME_STR, TEMP_CENTRE + TEMP_ABSCIS)
        TEMP2_SUM = TEMP2_SUM + TEMP2_ARR(h) * TEMP1_SUM
        If h Mod 2 = 0 Then TEMP3_SUM = TEMP3_SUM + TEMP1_ARR(h / 2) * TEMP1_SUM
    Next h
    KRON_INT = TEMP2_SUM * TEMP_HALF
    KRON_ERR = Abs(TEMP3_SUM - TEMP2_SUM) * TEMP_HALF
    COUNTER = 23 + COUNTER
Return  ' end sub routine Kronrod
1984:    ' Gauss-Legendre points
    '---------------------------------------------------------------------
    TEMP1_ARR(0) = 0.272925086777901
    TEMP1_ARR(1) = 5.56685671161745E-02
    TEMP1_ARR(2) = 0.125580369464905
    TEMP1_ARR(3) = 0.186290210927735
    TEMP1_ARR(4) = 0.233193764591991
    TEMP1_ARR(5) = 0.262804544510248
    '---------------------------------------------------------------------
    TEMP2_ARR(0) = 0.136577794711118
    TEMP2_ARR(1) = 9.76544104596129E-03
    TEMP2_ARR(2) = 2.71565546821044E-02
    TEMP2_ARR(3) = 4.58293785644267E-02
    TEMP2_ARR(4) = 6.30974247503748E-02
    TEMP2_ARR(5) = 7.86645719322276E-02
    TEMP2_ARR(6) = 9.29530985969007E-02
    TEMP2_ARR(7) = 0.105872074481389
    TEMP2_ARR(8) = 0.116739502461047
    TEMP2_ARR(9) = 0.125158799100319
    TEMP2_ARR(10) = 0.131280684229806
    TEMP2_ARR(11) = 0.135193572799885
    '---------------------------------------------------------------------
    TEMP3_ARR(0) = 0#
    TEMP3_ARR(1) = 0.996369613889543
    TEMP3_ARR(2) = 0.978228658146057
    TEMP3_ARR(3) = 0.941677108578068
    TEMP3_ARR(4) = 0.887062599768095
    TEMP3_ARR(5) = 0.816057456656221
    TEMP3_ARR(6) = 0.730152005574049
    TEMP3_ARR(7) = 0.630599520161965
    TEMP3_ARR(8) = 0.519096129206812
    TEMP3_ARR(9) = 0.397944140952378
    TEMP3_ARR(10) = 0.269543155952345
    TEMP3_ARR(11) = 0.136113000799362
Return  ' end sub routine Kronroddata

Exit Function
ERROR_LABEL:
GAUSS_KRONROD_INTEGRATION_FUNC = Err.number
End Function
