Attribute VB_Name = "STAT_DIST_NORMAL_BIVAR_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.


'************************************************************************************
'************************************************************************************
'FUNCTION      : BIVAR_CUMUL_NORM_FUNC
'DESCRIPTION   : Bivariate Cumulative Normal Distribution Functions
'LIBRARY       : STATISTICS
'GROUP         : DIST_NORMAL_BIVAR
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/22/2009
'************************************************************************************
'************************************************************************************

Function BIVAR_CUMUL_NORM_FUNC(ByVal BIVAR_TYPE As String, _
ByVal UNIVAR_TYPE As String, _
ByVal X1_VAL As Double, _
ByVal X2_VAL As Double, _
ByVal RHO As Double) As Double

On Error GoTo ERROR_LABEL

'For UNIVAR_TYPE a series approach by Marsaglia and the solution given at Genz
'are compared. It turns out that for small values an interpolation should
'be used (cdfN_Hart), for medium size Marsaglia cdfN_Marsaglia is worth its
'cost and for larger ones an asymptotic is the choice.

'In cdfN_Marsaglia functions are 'Taylored' around 0,...,7 and this should
'stop after 20 steps with Excel's exactness. The asymptotic I think is due
'to Legendre (and for example is used in Maple).


  Select Case BIVAR_TYPE
'------------------------------------------------------------------------------------
'Thus the method of Drezner & Wesolowsky given by Genz results in good speed
'and exactness up to 14 or 15 digits - based on a good choice for cdfN1:
'may be in Fortran it is not neccessary to modify the orignal cdfN, but at
'least Excel needs that.
 
  Case "drezner"
    BIVAR_CUMUL_NORM_FUNC = DREZNER_BIVAR_NORM_FUNC(UNIVAR_TYPE, _
                            X1_VAL, X2_VAL, RHO)
    Exit Function
'------------------------------------------------------------------------------------
  Case "fixed drezner"
    BIVAR_CUMUL_NORM_FUNC = FIX_DREZNER_BIVAR_NORM_FUNC(UNIVAR_TYPE, _
                            X1_VAL, X2_VAL, RHO)
    Exit Function
'------------------------------------------------------------------------------------

  Case "drezwes first"
    BIVAR_CUMUL_NORM_FUNC = DREZWES_A_BIVAR_NORM_FUNC(UNIVAR_TYPE, _
                            X1_VAL, X2_VAL, RHO)
    Exit Function
'------------------------------------------------------------------------------------
  Case "drezwes second"
    BIVAR_CUMUL_NORM_FUNC = DREZWES_B_BIVAR_NORM_FUNC(UNIVAR_TYPE, _
                            X1_VAL, X2_VAL, RHO)
    Exit Function
'------------------------------------------------------------------------------------
  Case "fixed drezwes second"
    BIVAR_CUMUL_NORM_FUNC = FIX_DREZWES_B_BIVAR_NORM_FUNC(UNIVAR_TYPE, _
                            X1_VAL, X2_VAL, RHO)
    Exit Function
'------------------------------------------------------------------------------------
  Case "genz"
    BIVAR_CUMUL_NORM_FUNC = GENZ_BIVAR_NORM_FUNC(UNIVAR_TYPE, _
                            X1_VAL, X2_VAL, RHO)
    Exit Function
'------------------------------------------------------------------------------------
' In cdfN_Marsaglia functions are 'Taylored' around 0,...,7 and this should
' stop after 20 steps with Excel's exactness. The asymptotic I think is due
' to Legendre (and for example is used in Maple).

  Case "Marsaglia"
    BIVAR_CUMUL_NORM_FUNC = MARSAG_BIVAR_NORM_FUNC(X1_VAL, X2_VAL, RHO, 100)
    Exit Function
'------------------------------------------------------------------------------------
  End Select
'------------------------------------------------------------------------------------


Exit Function
ERROR_LABEL:
BIVAR_CUMUL_NORM_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : DREZNER_BIVAR_NORM_FUNC
'DESCRIPTION   : Bivariate Drezner Normal Distribution Function
'LIBRARY       : STATISTICS
'GROUP         : DIST_NORMAL_BIVAR
'ID            : 002
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/22/2009
'************************************************************************************
'************************************************************************************

Function DREZNER_BIVAR_NORM_FUNC(ByVal UNIVAR_TYPE As String, _
ByVal X1_VAL As Double, _
ByVal X2_VAL As Double, _
ByVal RHO_VAL As Double) As Double
    
Dim ii As Integer
Dim jj As Integer

Dim XTEMP_ARR As Variant
Dim YTEMP_ARR As Variant

Dim FIRST_RHO As Double
Dim SECOND_RHO As Double

Dim ATEMP_VAL As Double
Dim BTEMP_VAL As Double

Dim TEMP_SUM As Double
Dim TEMP_DELTA As Double

Dim PI_VAL As Double

On Error GoTo ERROR_LABEL

PI_VAL = 3.14159265358979

XTEMP_ARR = Array(0.24840615, 0.39233107, 0.21141819, 0.03324666, 0.00082485334)
YTEMP_ARR = Array(0.10024215, 0.48281397, 1.0609498, 1.7797294, 2.6697604)

ATEMP_VAL = X1_VAL / Sqr(2 * (1 - RHO_VAL ^ 2))
BTEMP_VAL = X2_VAL / Sqr(2 * (1 - RHO_VAL ^ 2))

If X1_VAL <= 0 And X2_VAL <= 0 And RHO_VAL <= 0 Then
    TEMP_SUM = 0
    For ii = LBound(XTEMP_ARR, 1) To UBound(XTEMP_ARR, 1)
        For jj = LBound(XTEMP_ARR, 1) To UBound(XTEMP_ARR, 1)
            TEMP_SUM = TEMP_SUM + XTEMP_ARR(ii) * _
                        XTEMP_ARR(jj) * Exp(ATEMP_VAL * (2 * _
                        YTEMP_ARR(ii) - ATEMP_VAL) _
                        + BTEMP_VAL * (2 * YTEMP_ARR(jj) - BTEMP_VAL) + 2 * _
                        RHO_VAL * (YTEMP_ARR(ii) - ATEMP_VAL) * _
                        (YTEMP_ARR(jj) - BTEMP_VAL))
        Next jj
    Next ii
    DREZNER_BIVAR_NORM_FUNC = Sqr(1 - RHO_VAL ^ 2) / PI_VAL * TEMP_SUM

ElseIf X1_VAL <= 0 And X2_VAL >= 0 And RHO_VAL >= 0 Then
    DREZNER_BIVAR_NORM_FUNC = _
        UNIVAR_CUMUL_NORM_FUNC(UNIVAR_TYPE, X1_VAL) - _
            DREZNER_BIVAR_NORM_FUNC(UNIVAR_TYPE, X1_VAL, -X2_VAL, -RHO_VAL)

ElseIf X1_VAL >= 0 And X2_VAL <= 0 And RHO_VAL >= 0 Then
    DREZNER_BIVAR_NORM_FUNC = UNIVAR_CUMUL_NORM_FUNC(UNIVAR_TYPE, X2_VAL) - _
        DREZNER_BIVAR_NORM_FUNC(UNIVAR_TYPE, -X1_VAL, X2_VAL, -RHO_VAL)

ElseIf X1_VAL >= 0 And X2_VAL >= 0 And RHO_VAL <= 0 Then
    DREZNER_BIVAR_NORM_FUNC = UNIVAR_CUMUL_NORM_FUNC(UNIVAR_TYPE, X1_VAL) + _
        UNIVAR_CUMUL_NORM_FUNC(UNIVAR_TYPE, X2_VAL) - 1 + _
    DREZNER_BIVAR_NORM_FUNC(UNIVAR_TYPE, -X1_VAL, -X2_VAL, RHO_VAL)

ElseIf X1_VAL * X2_VAL * RHO_VAL > 0 Then
    FIRST_RHO = (RHO_VAL * X1_VAL - X2_VAL) * _
            Sgn(X1_VAL) / Sqr(X1_VAL ^ 2 - 2 * _
            RHO_VAL * X1_VAL * X2_VAL + X2_VAL ^ 2)
    SECOND_RHO = (RHO_VAL * X2_VAL - X1_VAL) * Sgn(X2_VAL) / _
            Sqr(X1_VAL ^ 2 - 2 * RHO_VAL * X1_VAL * X2_VAL _
            + X2_VAL ^ 2)
    TEMP_DELTA = (1 - Sgn(X1_VAL) * Sgn(X2_VAL)) / 4
    DREZNER_BIVAR_NORM_FUNC = _
                DREZNER_BIVAR_NORM_FUNC(UNIVAR_TYPE, X1_VAL, 0, FIRST_RHO) + _
                DREZNER_BIVAR_NORM_FUNC(UNIVAR_TYPE, X2_VAL, 0, _
    SECOND_RHO) - TEMP_DELTA
End If

Exit Function
ERROR_LABEL:
DREZNER_BIVAR_NORM_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : FIX_DREZNER_BIVAR_NORM_FUNC
'DESCRIPTION   : Improved bivariate drezner - deals with cases where |rho| = 1
'LIBRARY       : STATISTICS
'GROUP         : DIST_NORMAL_BIVAR
'ID            : 003
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/22/2009
'************************************************************************************
'************************************************************************************

Function FIX_DREZNER_BIVAR_NORM_FUNC(ByVal UNIVAR_TYPE As String, _
ByVal X1_VAL As Double, _
ByVal X2_VAL As Double, _
ByVal RHO_VAL As Double) As Double
  
Dim ii As Integer
Dim jj As Integer

Dim XTEMP_ARR As Variant
Dim YTEMP_ARR As Variant

Dim FIRST_RHO As Double
Dim SECOND_RHO As Double

Dim TEMP_DELTA As Double
Dim TEMP_SUM As Double

Dim ATEMP_VAL As Double
Dim BTEMP_VAL As Double

Dim epsilon As Double
Dim PI_VAL As Double

On Error GoTo ERROR_LABEL

PI_VAL = 3.14159265358979
epsilon = 10 ^ -16

XTEMP_ARR = Array(0.24840615, 0.39233107, 0.21141819, 0.03324666, 0.00082485334)
YTEMP_ARR = Array(0.10024215, 0.48281397, 1.0609498, 1.7797294, 2.6697604)

If Abs(RHO_VAL) > 1 - epsilon Then
  If RHO_VAL > 0 Then
    FIX_DREZNER_BIVAR_NORM_FUNC = UNIVAR_CUMUL_NORM_FUNC(UNIVAR_TYPE, _
                  MAXIMUM_FUNC(X1_VAL, X2_VAL))
  Else
    If X2_VAL <= -X1_VAL Then
      FIX_DREZNER_BIVAR_NORM_FUNC = 0
    Else
      FIX_DREZNER_BIVAR_NORM_FUNC = UNIVAR_CUMUL_NORM_FUNC(UNIVAR_TYPE, _
                   X1_VAL) + UNIVAR_CUMUL_NORM_FUNC(UNIVAR_TYPE, _
                   X2_VAL) - 1
    End If
  End If
  Exit Function
End If

ATEMP_VAL = X1_VAL / Sqr(2 * (1 - RHO_VAL ^ 2))
BTEMP_VAL = X2_VAL / Sqr(2 * (1 - RHO_VAL ^ 2))

If X1_VAL <= 0 And X2_VAL <= 0 And RHO_VAL <= 0 Then
  TEMP_SUM = 0
  For ii = LBound(XTEMP_ARR, 1) To UBound(XTEMP_ARR, 1)
    For jj = LBound(XTEMP_ARR, 1) To UBound(XTEMP_ARR, 1)
      TEMP_SUM = TEMP_SUM + XTEMP_ARR(ii) * XTEMP_ARR(jj) * _
                  Exp(ATEMP_VAL * (2 * YTEMP_ARR(ii) - ATEMP_VAL) _
                  + BTEMP_VAL * (2 * YTEMP_ARR(jj) - BTEMP_VAL) + 2 * _
                  RHO_VAL * (YTEMP_ARR(ii) - ATEMP_VAL) * _
                  (YTEMP_ARR(jj) - BTEMP_VAL))
    Next jj
  Next ii
  FIX_DREZNER_BIVAR_NORM_FUNC = Sqr(1 - RHO_VAL ^ 2) / PI_VAL * TEMP_SUM

ElseIf X1_VAL <= 0 And X2_VAL >= 0 And RHO_VAL >= 0 Then
  FIX_DREZNER_BIVAR_NORM_FUNC = UNIVAR_CUMUL_NORM_FUNC(UNIVAR_TYPE, _
                  X1_VAL) - FIX_DREZNER_BIVAR_NORM_FUNC(UNIVAR_TYPE, _
                  X1_VAL, -X2_VAL, -RHO_VAL)

ElseIf X1_VAL >= 0 And X2_VAL <= 0 And RHO_VAL >= 0 Then
  FIX_DREZNER_BIVAR_NORM_FUNC = UNIVAR_CUMUL_NORM_FUNC(UNIVAR_TYPE, _
                  X2_VAL) - FIX_DREZNER_BIVAR_NORM_FUNC(UNIVAR_TYPE, _
                  -X1_VAL, X2_VAL, -RHO_VAL)

ElseIf X1_VAL >= 0 And X2_VAL >= 0 And RHO_VAL <= 0 Then
  FIX_DREZNER_BIVAR_NORM_FUNC = UNIVAR_CUMUL_NORM_FUNC(UNIVAR_TYPE, _
                  X1_VAL) + UNIVAR_CUMUL_NORM_FUNC(UNIVAR_TYPE, _
                  X2_VAL) - 1 + FIX_DREZNER_BIVAR_NORM_FUNC(UNIVAR_TYPE, _
                  -X1_VAL, -X2_VAL, RHO_VAL)

ElseIf X1_VAL * X2_VAL * RHO_VAL > 0 Then
  
  FIRST_RHO = (RHO_VAL * X1_VAL - X2_VAL) * Sgn(X1_VAL) / _
          Sqr(X1_VAL ^ 2 - 2 * RHO_VAL * X1_VAL * _
          X2_VAL + X2_VAL ^ 2)
  
  SECOND_RHO = (RHO_VAL * X2_VAL - X1_VAL) * Sgn(X2_VAL) / _
          Sqr(X1_VAL ^ 2 - 2 * RHO_VAL * X1_VAL * _
          X2_VAL + X2_VAL ^ 2)
  TEMP_DELTA = (1 - Sgn(X1_VAL) * Sgn(X2_VAL)) / 4
  
  FIX_DREZNER_BIVAR_NORM_FUNC = FIX_DREZNER_BIVAR_NORM_FUNC(UNIVAR_TYPE, _
              X1_VAL, 0, FIRST_RHO) + FIX_DREZNER_BIVAR_NORM_FUNC(UNIVAR_TYPE, _
              X2_VAL, 0, SECOND_RHO) - TEMP_DELTA
End If

Exit Function
ERROR_LABEL:
FIX_DREZNER_BIVAR_NORM_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : DREZWES_A_BIVAR_NORM_FUNC
'DESCRIPTION   : First Drez & Wes bivariate normal distribution function
'LIBRARY       : STATISTICS
'GROUP         : DIST_NORMAL_BIVAR
'ID            : 004
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/22/2009
'************************************************************************************
'************************************************************************************

Function DREZWES_A_BIVAR_NORM_FUNC(ByVal UNIVAR_TYPE As String, _
ByVal X1_VAL As Double, _
ByVal X2_VAL As Double, _
ByVal RHO_VAL As Double) As Double
  
Dim ii As Integer

Dim TEMP_SUM As Double
Dim ATEMP_VAL As Double
Dim BTEMP_VAL As Double

Dim XTEMP_ARR As Variant
Dim YTEMP_ARR As Variant

On Error GoTo ERROR_LABEL

XTEMP_ARR = Array(0.018854042, 0.038088059, 0.0452707394, 0.038088059, 0.018854042)
YTEMP_ARR = Array(0.04691008, 0.23076534, 0.5, 0.76923466, 0.95308992)

TEMP_SUM = 0
For ii = LBound(XTEMP_ARR, 1) To UBound(XTEMP_ARR, 1)
  ATEMP_VAL = YTEMP_ARR(ii) * RHO_VAL
  BTEMP_VAL = 1 - ATEMP_VAL ^ 2
  TEMP_SUM = TEMP_SUM + XTEMP_ARR(ii) * Exp((2 * X1_VAL * _
        X2_VAL * ATEMP_VAL - X1_VAL ^ 2 - _
        X2_VAL ^ 2) / BTEMP_VAL / 2) / Sqr(BTEMP_VAL)
Next ii

DREZWES_A_BIVAR_NORM_FUNC = RHO_VAL * TEMP_SUM + _
          UNIVAR_CUMUL_NORM_FUNC(UNIVAR_TYPE, X1_VAL) * _
          UNIVAR_CUMUL_NORM_FUNC(UNIVAR_TYPE, X2_VAL)

Exit Function
ERROR_LABEL:
DREZWES_A_BIVAR_NORM_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : DREZWES_B_BIVAR_NORM_FUNC
'DESCRIPTION   : Second Drez & Wes bivariate normal distribution function
'LIBRARY       : STATISTICS
'GROUP         : DIST_NORMAL_BIVAR
'ID            : 005
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/22/2009
'************************************************************************************
'************************************************************************************

Function DREZWES_B_BIVAR_NORM_FUNC(ByVal UNIVAR_TYPE As String, _
ByVal X1_VAL As Double, _
ByVal X2_VAL As Double, _
ByVal RHO_VAL As Double) As Double

Dim ii As Integer

Dim ATEMP_VAL As Double
Dim BTEMP_VAL As Double

Dim CTEMP_VAL As Double
Dim DTEMP_VAL As Double
Dim ETEMP_VAL As Double
Dim FTEMP_VAL As Double
Dim GTEMP_VAL As Double
Dim HTEMP_VAL As Double

Dim ITEMP_VAL As Double
Dim JTEMP_VAL As Double
Dim KTEMP_VAL As Double

Dim LTEMP_VAL As Double
Dim MTEMP_VAL As Double
Dim NTEMP_VAL As Double

Dim XTEMP_ARR As Variant
Dim YTEMP_ARR As Variant

On Error GoTo ERROR_LABEL

XTEMP_ARR = Array(0.04691008, 0.23076534, 0.5, 0.76923466, 0.95308992)
YTEMP_ARR = Array(0.018854042, 0.038088059, 0.0452707394, 0.038088059, 0.018854042)

ATEMP_VAL = X1_VAL
BTEMP_VAL = X2_VAL
DTEMP_VAL = (ATEMP_VAL * ATEMP_VAL + BTEMP_VAL * BTEMP_VAL) / 2

If Abs(RHO_VAL) >= 0.7 Then
  MTEMP_VAL = 1 - RHO_VAL * RHO_VAL
  NTEMP_VAL = Sqr(MTEMP_VAL)
  If RHO_VAL < 0 Then BTEMP_VAL = -BTEMP_VAL
  
  ETEMP_VAL = ATEMP_VAL * BTEMP_VAL
  HTEMP_VAL = Exp(-ETEMP_VAL / 2)
  
  If Abs(RHO_VAL) < 1 Then
    
    GTEMP_VAL = Abs(ATEMP_VAL - BTEMP_VAL)
    FTEMP_VAL = GTEMP_VAL * GTEMP_VAL / 2
    GTEMP_VAL = GTEMP_VAL / NTEMP_VAL
    JTEMP_VAL = 0.5 - ETEMP_VAL / 8
    KTEMP_VAL = 3 - 2 * JTEMP_VAL * FTEMP_VAL
    CTEMP_VAL = 0.13298076 * GTEMP_VAL * KTEMP_VAL * (1 - _
        UNIVAR_CUMUL_NORM_FUNC(UNIVAR_TYPE, GTEMP_VAL)) - Exp(-FTEMP_VAL / _
    MTEMP_VAL) * (KTEMP_VAL + JTEMP_VAL * MTEMP_VAL) * 0.053051647
    
    For ii = LBound(XTEMP_ARR, 1) To UBound(XTEMP_ARR, 1)
      LTEMP_VAL = NTEMP_VAL * XTEMP_ARR(ii)
      ITEMP_VAL = LTEMP_VAL * LTEMP_VAL
      MTEMP_VAL = Sqr(1 - ITEMP_VAL)
      CTEMP_VAL = CTEMP_VAL - YTEMP_ARR(ii) * Exp(-FTEMP_VAL / ITEMP_VAL) * _
                  (Exp(-ETEMP_VAL / (1 + MTEMP_VAL)) / MTEMP_VAL / HTEMP_VAL - 1 _
      - JTEMP_VAL * ITEMP_VAL)
    Next ii
  End If
  DREZWES_B_BIVAR_NORM_FUNC = CTEMP_VAL * NTEMP_VAL * HTEMP_VAL + _
          UNIVAR_CUMUL_NORM_FUNC(UNIVAR_TYPE, _
          MAXIMUM_FUNC(ATEMP_VAL, BTEMP_VAL))
  If RHO_VAL < 0 Then
    DREZWES_B_BIVAR_NORM_FUNC = _
      UNIVAR_CUMUL_NORM_FUNC(UNIVAR_TYPE, ATEMP_VAL) - DREZWES_B_BIVAR_NORM_FUNC
  End If
Else
  ETEMP_VAL = ATEMP_VAL * BTEMP_VAL
  If RHO_VAL <> 0 Then
    For ii = LBound(XTEMP_ARR, 1) To UBound(XTEMP_ARR, 1)
      LTEMP_VAL = RHO_VAL * XTEMP_ARR(ii)
      MTEMP_VAL = 1 - LTEMP_VAL * LTEMP_VAL
      CTEMP_VAL = CTEMP_VAL + YTEMP_ARR(ii) * _
                  Exp((LTEMP_VAL * ETEMP_VAL - DTEMP_VAL) / _
                  MTEMP_VAL) / Sqr(MTEMP_VAL)
    Next ii
  End If
  DREZWES_B_BIVAR_NORM_FUNC = UNIVAR_CUMUL_NORM_FUNC(UNIVAR_TYPE, ATEMP_VAL) * _
                 UNIVAR_CUMUL_NORM_FUNC(UNIVAR_TYPE, BTEMP_VAL) + _
                 RHO_VAL * CTEMP_VAL
End If
   
Exit Function
ERROR_LABEL:
DREZWES_B_BIVAR_NORM_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : FIX_DREZWES_B_BIVAR_NORM_FUNC

'DESCRIPTION   : Modified/corrected from the second function in Drez & Wes
'0/0 case resolved by l'H rule

'LIBRARY       : STATISTICS
'GROUP         : DIST_NORMAL_BIVAR
'ID            : 006
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/22/2009
'************************************************************************************
'************************************************************************************

Function FIX_DREZWES_B_BIVAR_NORM_FUNC(ByVal UNIVAR_TYPE As String, _
ByVal X1_VAL As Double, _
ByVal X2_VAL As Double, _
ByVal RHO_VAL As Double) As Double

Dim ii As Integer

Dim XTEMP_ARR As Variant
Dim YTEMP_ARR As Variant

Dim ATEMP_VAL As Double
Dim BTEMP_VAL As Double
Dim CTEMP_VAL As Double

Dim DTEMP_VAL As Double
Dim ETEMP_VAL As Double
Dim FTEMP_VAL As Double

Dim GTEMP_VAL As Double
Dim HTEMP_VAL As Double
Dim ITEMP_VAL As Double

Dim JTEMP_VAL As Double
Dim KTEMP_VAL As Double
Dim LTEMP_VAL As Double

Dim MTEMP_VAL As Double
Dim NTEMP_VAL As Double
Dim OTEMP_VAL As Double

On Error GoTo ERROR_LABEL

XTEMP_ARR = Array(0.04691008, 0.23076534, 0.5, 0.76923466, 0.95308992)
YTEMP_ARR = Array(0.018854042, 0.038088059, 0.0452707394, 0.038088059, 0.018854042)

ATEMP_VAL = X1_VAL
BTEMP_VAL = X2_VAL
DTEMP_VAL = (ATEMP_VAL * ATEMP_VAL + BTEMP_VAL * BTEMP_VAL) / 2

If Abs(RHO_VAL) >= 0.7 Then
  NTEMP_VAL = 1 - RHO_VAL * RHO_VAL
  OTEMP_VAL = Sqr(NTEMP_VAL)
  If RHO_VAL < 0 Then BTEMP_VAL = -BTEMP_VAL
  ETEMP_VAL = ATEMP_VAL * BTEMP_VAL
  HTEMP_VAL = Exp(-ETEMP_VAL / 2)
  
  If Abs(RHO_VAL) < 1 Then
    GTEMP_VAL = Abs(ATEMP_VAL - BTEMP_VAL)
    FTEMP_VAL = GTEMP_VAL * GTEMP_VAL / 2
    GTEMP_VAL = GTEMP_VAL / OTEMP_VAL
    KTEMP_VAL = 0.5 - ETEMP_VAL / 8
    LTEMP_VAL = 3 - 2 * KTEMP_VAL * FTEMP_VAL
    CTEMP_VAL = 0.13298076 * GTEMP_VAL * LTEMP_VAL * _
              (1 - UNIVAR_CUMUL_NORM_FUNC(UNIVAR_TYPE, GTEMP_VAL)) - _
              Exp(-FTEMP_VAL / NTEMP_VAL) * (LTEMP_VAL + KTEMP_VAL * _
              NTEMP_VAL) * 0.053051647
    
    For ii = LBound(XTEMP_ARR, 1) To UBound(XTEMP_ARR, 1)
      MTEMP_VAL = OTEMP_VAL * XTEMP_ARR(ii)
      JTEMP_VAL = MTEMP_VAL * MTEMP_VAL
      NTEMP_VAL = Sqr(1 - JTEMP_VAL)
      If HTEMP_VAL = 0 Then
        ITEMP_VAL = 0
      Else
        ITEMP_VAL = Exp(-ETEMP_VAL / (1 + NTEMP_VAL)) / NTEMP_VAL / HTEMP_VAL
      End If
      CTEMP_VAL = CTEMP_VAL - YTEMP_ARR(ii) * Exp(-FTEMP_VAL / _
                  JTEMP_VAL) * (ITEMP_VAL - 1 - KTEMP_VAL * JTEMP_VAL)
    Next ii
  End If
  FIX_DREZWES_B_BIVAR_NORM_FUNC = CTEMP_VAL * OTEMP_VAL * HTEMP_VAL + _
              UNIVAR_CUMUL_NORM_FUNC(UNIVAR_TYPE, _
              MAXIMUM_FUNC(ATEMP_VAL, BTEMP_VAL))
  If RHO_VAL < 0 Then
    FIX_DREZWES_B_BIVAR_NORM_FUNC = _
      UNIVAR_CUMUL_NORM_FUNC(UNIVAR_TYPE, ATEMP_VAL) - _
      FIX_DREZWES_B_BIVAR_NORM_FUNC
  End If
Else
  ETEMP_VAL = ATEMP_VAL * BTEMP_VAL
  If RHO_VAL <> 0 Then
    For ii = LBound(XTEMP_ARR, 1) To UBound(XTEMP_ARR, 1)
      MTEMP_VAL = RHO_VAL * XTEMP_ARR(ii)
      NTEMP_VAL = 1 - MTEMP_VAL * MTEMP_VAL
      CTEMP_VAL = CTEMP_VAL + YTEMP_ARR(ii) * Exp((MTEMP_VAL * _
                  ETEMP_VAL - DTEMP_VAL) / NTEMP_VAL) / Sqr(NTEMP_VAL)
    Next ii
  End If
  FIX_DREZWES_B_BIVAR_NORM_FUNC = _
          UNIVAR_CUMUL_NORM_FUNC(UNIVAR_TYPE, ATEMP_VAL) * _
          UNIVAR_CUMUL_NORM_FUNC(UNIVAR_TYPE, BTEMP_VAL) + RHO_VAL * CTEMP_VAL
End If
   
Exit Function
ERROR_LABEL:
FIX_DREZWES_B_BIVAR_NORM_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : GENZ_BIVAR_NORM_FUNC
'DESCRIPTION   : Function for computing bivariate normal probabilities.
'--> used for option valuation

'       Alan Genz
'       Department of Mathematics
'       Washington State University
'       Pullman, WA 99164-3113
'       Email : alangenz@wsu.edu
'    This function is based on the method described by
'        Drezner, Z and G.O. Wesolowsky, (1989),
'        On the computation of the bivariate normal integral,
'        Journal of Statist. Comput. Simul. 35, pp. 101-107,
'    with major modifications for double precision, and for |R| close to 1.

'LIBRARY       : STATISTICS
'GROUP         : DIST_NORMAL_BIVAR
'ID            : 007
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/22/2009
'************************************************************************************
'************************************************************************************

Function GENZ_BIVAR_NORM_FUNC(ByVal UNIVAR_TYPE As String, _
ByVal X1_VAL As Double, _
ByVal X2_VAL As Double, _
ByVal RHO_VAL As Double) As Double


Dim ii As Integer
Dim jj As Integer
Dim ll As Integer
Dim hh As Integer

Dim XTEMP_ARR(10, 3) As Double
Dim YTEMP_ARR(10, 3) As Double

Dim ATEMP_VAL As Double
Dim BTEMP_VAL As Double
Dim CTEMP_VAL As Double

Dim DTEMP_VAL As Double
Dim ETEMP_VAL As Double
Dim FTEMP_VAL As Double

Dim GTEMP_VAL As Double
Dim HTEMP_VAL As Double
Dim ITEMP_VAL As Double

Dim JTEMP_VAL As Double
Dim KTEMP_VAL As Double
Dim LTEMP_VAL As Double

Dim MTEMP_VAL As Double
Dim NTEMP_VAL As Double
Dim OTEMP_VAL As Double

Dim PI_VAL As Double
    
On Error GoTo ERROR_LABEL
    
PI_VAL = 3.14159265358979

YTEMP_ARR(1, 1) = 0.17132449237917
XTEMP_ARR(1, 1) = -0.932469514203152
YTEMP_ARR(2, 1) = 0.360761573048138
XTEMP_ARR(2, 1) = -0.661209386466265
YTEMP_ARR(3, 1) = 0.46791393457269
XTEMP_ARR(3, 1) = -0.238619186083197

YTEMP_ARR(1, 2) = 4.71753363865118E-02
XTEMP_ARR(1, 2) = -0.981560634246719
YTEMP_ARR(2, 2) = 0.106939325995318
XTEMP_ARR(2, 2) = -0.904117256370475
YTEMP_ARR(3, 2) = 0.160078328543346
XTEMP_ARR(3, 2) = -0.769902674194305
YTEMP_ARR(4, 2) = 0.203167426723066
XTEMP_ARR(4, 2) = -0.587317954286617
YTEMP_ARR(5, 2) = 0.233492536538355
XTEMP_ARR(5, 2) = -0.36783149899818
YTEMP_ARR(6, 2) = 0.249147045813403
XTEMP_ARR(6, 2) = -0.125233408511469

YTEMP_ARR(1, 3) = 1.76140071391521E-02
XTEMP_ARR(1, 3) = -0.993128599185095
YTEMP_ARR(2, 3) = 4.06014298003869E-02
XTEMP_ARR(2, 3) = -0.963971927277914
YTEMP_ARR(3, 3) = 6.26720483341091E-02
XTEMP_ARR(3, 3) = -0.912234428251326
YTEMP_ARR(4, 3) = 8.32767415767048E-02
XTEMP_ARR(4, 3) = -0.839116971822219
YTEMP_ARR(5, 3) = 0.10193011981724
XTEMP_ARR(5, 3) = -0.746331906460151
YTEMP_ARR(6, 3) = 0.118194531961518
XTEMP_ARR(6, 3) = -0.636053680726515
YTEMP_ARR(7, 3) = 0.131688638449177
XTEMP_ARR(7, 3) = -0.510867001950827
YTEMP_ARR(8, 3) = 0.142096109318382
XTEMP_ARR(8, 3) = -0.37370608871542
YTEMP_ARR(9, 3) = 0.149172986472604
XTEMP_ARR(9, 3) = -0.227785851141645
YTEMP_ARR(10, 3) = 0.152753387130726
XTEMP_ARR(10, 3) = -7.65265211334973E-02
      
If Abs(RHO_VAL) < 0.3 Then
  hh = 1
  ll = 3
ElseIf Abs(RHO_VAL) < 0.75 Then
  hh = 2
  ll = 6
Else
  hh = 3
  ll = 10
End If
      
CTEMP_VAL = -X1_VAL
DTEMP_VAL = -X2_VAL
ETEMP_VAL = CTEMP_VAL * DTEMP_VAL
GTEMP_VAL = 0
      
If Abs(RHO_VAL) < 0.925 Then
  If Abs(RHO_VAL) > 0 Then
    FTEMP_VAL = (CTEMP_VAL * CTEMP_VAL + DTEMP_VAL * DTEMP_VAL) / 2
    
    If Abs(RHO_VAL) = 1 Then
        ITEMP_VAL = Sgn(RHO_VAL) * PI_VAL / 2
    Else: ITEMP_VAL = Atn(RHO_VAL / Sqr(1 - RHO_VAL ^ 2))
    End If
    
    For ii = 1 To ll
      For jj = -1 To 1 Step 2
        JTEMP_VAL = Sin(ITEMP_VAL * (jj * XTEMP_ARR(ii, hh) + 1) / 2)
        GTEMP_VAL = GTEMP_VAL + YTEMP_ARR(ii, hh) * _
                    Exp((JTEMP_VAL * ETEMP_VAL - FTEMP_VAL) / _
                    (1 - JTEMP_VAL * JTEMP_VAL))
      Next jj
    Next ii
    GTEMP_VAL = GTEMP_VAL * ITEMP_VAL / (4 * PI_VAL)
  End If
  GTEMP_VAL = GTEMP_VAL + UNIVAR_CUMUL_NORM_FUNC(UNIVAR_TYPE, -CTEMP_VAL) * _
            UNIVAR_CUMUL_NORM_FUNC(UNIVAR_TYPE, -DTEMP_VAL)
Else
  If RHO_VAL < 0 Then
    DTEMP_VAL = -DTEMP_VAL
    ETEMP_VAL = -ETEMP_VAL
  End If
  If Abs(RHO_VAL) < 1 Then
    HTEMP_VAL = (1 - RHO_VAL) * (1 + RHO_VAL)
    ATEMP_VAL = Sqr(HTEMP_VAL)
    KTEMP_VAL = (CTEMP_VAL - DTEMP_VAL) ^ 2
    LTEMP_VAL = (4 - ETEMP_VAL) / 8
    MTEMP_VAL = (12 - ETEMP_VAL) / 16
    ITEMP_VAL = -(KTEMP_VAL / HTEMP_VAL + ETEMP_VAL) / 2
    If ITEMP_VAL > -100 Then GTEMP_VAL = ATEMP_VAL * _
                Exp(ITEMP_VAL) * (1 - LTEMP_VAL * _
                (KTEMP_VAL - HTEMP_VAL) * _
                (1 - MTEMP_VAL * KTEMP_VAL / 5) / 3 + _
                LTEMP_VAL * MTEMP_VAL * HTEMP_VAL * HTEMP_VAL / 5)
    If -ETEMP_VAL < 100 Then
      BTEMP_VAL = Sqr(KTEMP_VAL)
      GTEMP_VAL = GTEMP_VAL - Exp(-ETEMP_VAL / 2) * Sqr(2 * PI_VAL) * _
            UNIVAR_CUMUL_NORM_FUNC(UNIVAR_TYPE, _
            -BTEMP_VAL / ATEMP_VAL) * BTEMP_VAL * (1 - LTEMP_VAL * _
            KTEMP_VAL * (1 - MTEMP_VAL * KTEMP_VAL / 5) / 3)
    End If
    ATEMP_VAL = ATEMP_VAL / 2
    For ii = 1 To ll
      For jj = -1 To 1 Step 2
        NTEMP_VAL = (ATEMP_VAL * (jj * XTEMP_ARR(ii, hh) + 1)) ^ 2
        OTEMP_VAL = Sqr(1 - NTEMP_VAL)
        ITEMP_VAL = -(KTEMP_VAL / NTEMP_VAL + ETEMP_VAL) / 2
        If ITEMP_VAL > -100 Then
           GTEMP_VAL = GTEMP_VAL + ATEMP_VAL * _
                    YTEMP_ARR(ii, hh) * Exp(ITEMP_VAL) * _
                    (Exp(-ETEMP_VAL * (1 - OTEMP_VAL) / _
                    (2 * (1 + OTEMP_VAL))) / OTEMP_VAL - (1 + LTEMP_VAL * _
                    NTEMP_VAL * (1 + MTEMP_VAL * NTEMP_VAL)))
        End If
      Next jj
    Next ii
    GTEMP_VAL = -GTEMP_VAL / (2 * PI_VAL)
  End If
  If RHO_VAL > 0 Then
    GTEMP_VAL = GTEMP_VAL + UNIVAR_CUMUL_NORM_FUNC(UNIVAR_TYPE, _
                        -MINIMUM_FUNC(CTEMP_VAL, DTEMP_VAL))
  Else
    GTEMP_VAL = -GTEMP_VAL
    If DTEMP_VAL > CTEMP_VAL Then GTEMP_VAL = GTEMP_VAL + _
                    UNIVAR_CUMUL_NORM_FUNC(UNIVAR_TYPE, DTEMP_VAL) - _
                    UNIVAR_CUMUL_NORM_FUNC(UNIVAR_TYPE, CTEMP_VAL)
  End If
End If
GENZ_BIVAR_NORM_FUNC = GTEMP_VAL

Exit Function
ERROR_LABEL:
GENZ_BIVAR_NORM_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : MARSAG_BIVAR_NORM_FUNC

'DESCRIPTION   : Marsaglia Bivariate Normal Distribution

'For lower correlation Marsaglia's method can be used to write
'down recursions for the Taylor series, while for higher correlations it is
'better to use a method of Vasicek (which I modified a bit), both more or
'less are recursions for the incomplete Gamma function involved. For both
'the expansion is stopped by machine precision (and the cdfN1 used). The
'code looks a bit ugly as I generated it using Maple procedures (to have
'comparable results for tests). Note that the series approach is limited
'by ~ 1e-15 for exactness: a decomposition in two (very exact) summands is
'done, but a system exactness 1 + eps <> 1 is involved.


'LIBRARY       : STATISTICS
'GROUP         : DIST_NORMAL_BIVAR
'ID            : 008
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/22/2009
'************************************************************************************
'************************************************************************************

Function MARSAG_BIVAR_NORM_FUNC( _
ByVal X1_VAL As Double, _
ByVal X2_VAL As Double, _
ByVal RHO_VAL As Double, _
ByVal nSTEPS As Integer) As Double
  
Dim ATEMP_VAL As Double
Dim BTEMP_VAL As Double
Dim CTEMP_VAL As Double
Dim DTEMP_VAL As Double
Dim PI_VAL As Double
    
On Error GoTo ERROR_LABEL
  
PI_VAL = 3.14159265358979

If (1 < Abs(RHO_VAL)) Then
ElseIf (Abs(RHO_VAL) = 1) Then
  
  MARSAG_BIVAR_NORM_FUNC = CDbl(N2_ABS_RHO_FUNC(X1_VAL, X2_VAL, Sgn(RHO_VAL)))
  Exit Function
Else
End If

If (Sgn(RHO_VAL) = 0) Then
  MARSAG_BIVAR_NORM_FUNC = CDbl(MARSAG2_CUMUL_NORM_FUNC(X1_VAL) * _
              MARSAG2_CUMUL_NORM_FUNC(X2_VAL))
ElseIf (Sgn(X1_VAL) = 0 And Sgn(X2_VAL) = 0) Then
  MARSAG_BIVAR_NORM_FUNC = CDbl((PI_VAL + 2# * Atn(RHO_VAL * _
              ((1# - RHO_VAL * RHO_VAL) ^ (-1# / 2#)))) / PI_VAL / 4#)
Else
End If

If (Sgn(X1_VAL) = 0) Then
  MARSAG_BIVAR_NORM_FUNC = CDbl(MARSAG_N2_REDUC_SUM_FUNC(X2_VAL, RHO_VAL, nSTEPS))
ElseIf (Sgn(X2_VAL) = 0) Then
  MARSAG_BIVAR_NORM_FUNC = CDbl(MARSAG_N2_REDUC_SUM_FUNC(X1_VAL, RHO_VAL, nSTEPS))
Else
End If

ATEMP_VAL = 1# / Sqr(X1_VAL * X1_VAL - 2# * _
          RHO_VAL * X1_VAL * X2_VAL + X2_VAL * X2_VAL)
BTEMP_VAL = (RHO_VAL * X1_VAL - X2_VAL) * ATEMP_VAL * CDbl(Sgn(X1_VAL))
CTEMP_VAL = (RHO_VAL * X2_VAL - X1_VAL) * ATEMP_VAL * CDbl(Sgn(X2_VAL))
If (Sgn(X2_VAL * X1_VAL) = -1) Then
  DTEMP_VAL = MARSAG_N2_REDUC_SUM_FUNC(X1_VAL, BTEMP_VAL, nSTEPS) + _
      MARSAG_N2_REDUC_SUM_FUNC(X2_VAL, CTEMP_VAL, nSTEPS) - 1# / 2#
ElseIf (Sgn(X2_VAL * X1_VAL) = 1) Then
  DTEMP_VAL = MARSAG_N2_REDUC_SUM_FUNC(X1_VAL, BTEMP_VAL, nSTEPS) + _
      MARSAG_N2_REDUC_SUM_FUNC(X2_VAL, CTEMP_VAL, nSTEPS)
Else
End If

MARSAG_BIVAR_NORM_FUNC = CDbl(DTEMP_VAL)

Exit Function
ERROR_LABEL:
MARSAG_BIVAR_NORM_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : BINORMAL_DENSITY_FUNC
'DESCRIPTION   : Compute standard binormal density function
'LIBRARY       : STATISTICS
'GROUP         : DIST_NORMAL_BIVAR
'ID            : 009
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/22/2009
'************************************************************************************
'************************************************************************************

Function BINORMAL_DENSITY_FUNC(ByVal X1_VAL As Double, _
ByVal X2_VAL As Double, _
ByVal RHO_VAL As Double)

Dim PI_2_VAL As Double
Dim RESULT_VAL As Double

Dim TEMP_FACTOR As Double

Dim FIRST_VAL As Double
Dim SECOND_VAL As Double

On Error GoTo ERROR_LABEL

PI_2_VAL = (2 * 3.14159265358979)
TEMP_FACTOR = 1 - RHO_VAL * RHO_VAL
FIRST_VAL = 1 / (PI_2_VAL * Sqr(TEMP_FACTOR))

SECOND_VAL = -(X1_VAL * X1_VAL - RHO_VAL * 2 * _
        X1_VAL * X2_VAL + X2_VAL * X2_VAL) / (TEMP_FACTOR * 2)

RESULT_VAL = FIRST_VAL * Exp(SECOND_VAL)

BINORMAL_DENSITY_FUNC = RESULT_VAL

Exit Function
ERROR_LABEL:
BINORMAL_DENSITY_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : MARSAG_N2_REDUC_SUM_FUNC
'DESCRIPTION   :
'LIBRARY       : STATISTICS
'GROUP         : DIST_NORMAL_BIVAR
'ID            : 010
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/22/2009
'************************************************************************************
'************************************************************************************

Private Function MARSAG_N2_REDUC_SUM_FUNC( _
ByVal X1_VAL As Double, _
ByVal RHO_VAL As Double, _
ByVal nSTEPS As Integer) As Double
  
Dim PI_VAL As Double
  
On Error GoTo ERROR_LABEL
  
PI_VAL = 3.14159265358979

'  If (1# <= Abs(RHO_VAL)) Then
' End If
' else generate an error (message)

If (Sgn(X1_VAL) = 0) Then
  MARSAG_N2_REDUC_SUM_FUNC = CDbl(1# / 4# + 1# / PI_VAL * Atn(RHO_VAL * _
  ((1# - RHO_VAL * RHO_VAL) ^ (-1# / 2#))) / 2#)
End If

If (Sgn(RHO_VAL) = 0) Then
  MARSAG_N2_REDUC_SUM_FUNC = CDbl(MARSAG2_CUMUL_NORM_FUNC(X1_VAL) / 2#)
End If

If (0# < X1_VAL And 0# < RHO_VAL) Then
  MARSAG_N2_REDUC_SUM_FUNC = MARSAG_N2_SERIES_FUNC(X1_VAL, RHO_VAL, nSTEPS)

ElseIf (0# < X1_VAL And 0# < -RHO_VAL) Then
  MARSAG_N2_REDUC_SUM_FUNC = CDbl(MARSAG2_CUMUL_NORM_FUNC(X1_VAL)) - _
                      MARSAG_N2_REDUC_SUM_FUNC(X1_VAL, -RHO_VAL, nSTEPS)

ElseIf (0# < -X1_VAL And 0# < RHO_VAL) Then
  MARSAG_N2_REDUC_SUM_FUNC = CDbl(1# / 2# - MARSAG2_CUMUL_NORM_FUNC(-X1_VAL)) + _
                      MARSAG_N2_REDUC_SUM_FUNC(-X1_VAL, RHO_VAL, nSTEPS)

ElseIf (0# < -X1_VAL And 0# < -RHO_VAL) Then
  MARSAG_N2_REDUC_SUM_FUNC = 0.5 - _
              MARSAG_N2_REDUC_SUM_FUNC(-X1_VAL, -RHO_VAL, nSTEPS)
Else
End If

Exit Function
ERROR_LABEL:
MARSAG_N2_REDUC_SUM_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : MARSAG_N2_SERIES_FUNC
'DESCRIPTION   :
'LIBRARY       : STATISTICS
'GROUP         : DIST_NORMAL_BIVAR
'ID            : 011
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/22/2009
'************************************************************************************
'************************************************************************************

Private Function MARSAG_N2_SERIES_FUNC( _
ByVal X1_VAL As Double, _
ByVal RHO_VAL As Double, _
ByVal nSTEPS As Integer) As Double
  
On Error GoTo ERROR_LABEL

If (Abs(RHO_VAL) <= CDbl(1# / Sqr(2#))) Then
  MARSAG_N2_SERIES_FUNC = N2_MARSAG_SERIES_FUNC(X1_VAL, -RHO_VAL, nSTEPS)
ElseIf (Abs(RHO_VAL) < 1#) Then
  MARSAG_N2_SERIES_FUNC = MARSAG_N2_VASICEK_SERIES_FUNC(X1_VAL, RHO_VAL, nSTEPS)
Else
End If

Exit Function
ERROR_LABEL:
MARSAG_N2_SERIES_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : MARSAG_N2_VASICEK_SERIES_FUNC

'DESCRIPTION   : For correlation Marsaglia's method can be used to write
' down recursions for the Taylor series while for higher correlations it is
' better to use a method of Vasicek (which i modified a bit), both more or
' less are recursions for the incomplete Gamma function involved. For both
' the expansion is stopped by machine precision.

'REFERENCE:
' Oldrich Alfons Vasicek (1998), Moody's KMV,
' http://www.moodyskmv.com/research/whitepaper/a_Series_Expansion_for_the__
' Bivariate_Normal_Integral.pdf
' as series in 1-rho^2 (which is like Drezner & Wesolowsky, cf Genz)
' i did re-write that to understand it

'LIBRARY       : STATISTICS
'GROUP         : DIST_NORMAL_BIVAR
'ID            : 012
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/22/2009
'************************************************************************************
'************************************************************************************

Private Function MARSAG_N2_VASICEK_SERIES_FUNC( _
ByVal X1_VAL As Double, _
ByVal RHO As Double, _
ByVal nSTEPS As Integer) As Double
  
Dim ii As Integer

Dim TEMP_SUM As Double

Dim ATEMP_VAL As Double
Dim BTEMP_VAL As Double
Dim CTEMP_VAL As Double

Dim PI_VAL As Double
  
On Error GoTo ERROR_LABEL
  
PI_VAL = 3.14159265358979

CTEMP_VAL = Sqr(1# - RHO * RHO) * CDbl(Sgn(X1_VAL)) * Exp(1# / (RHO + 1#) / _
          (RHO - 1#) * X1_VAL * X1_VAL / 2#)
TEMP_SUM = 0#

ATEMP_VAL = CDbl(-2# * X1_VAL * Sqr(2# * PI_VAL) * _
          MARSAG2_CUMUL_NORM_FUNC(-Abs(X1_VAL) * _
          ((1# - RHO * RHO) ^ (-1# / 2#))) + 2# * CTEMP_VAL)
TEMP_SUM = CDbl(TEMP_SUM + ATEMP_VAL)

BTEMP_VAL = CDbl(-2# * CTEMP_VAL * (1# - RHO * RHO))
ATEMP_VAL = CDbl(-X1_VAL * X1_VAL * ATEMP_VAL / 6# - BTEMP_VAL / 6#)
TEMP_SUM = CDbl(TEMP_SUM + ATEMP_VAL)

For ii = 2 To nSTEPS
  BTEMP_VAL = CDbl(-(-1# + RHO * RHO) * CDbl(2 * ii - 1) / _
              CDbl(ii - 1) * BTEMP_VAL / 2#)
  ATEMP_VAL = CDbl((X1_VAL * X1_VAL * ATEMP_VAL * _
              CDbl(-2 * ii + 1) - BTEMP_VAL) / CDbl(ii) / _
  CDbl(1 + 2 * ii) / 2#)
  TEMP_SUM = CDbl(TEMP_SUM + ATEMP_VAL)
  If (TEMP_SUM = TEMP_SUM - ATEMP_VAL) Then
    Exit For
  End If
Next

MARSAG_N2_VASICEK_SERIES_FUNC = 0.5 - CDbl(TEMP_SUM / PI_VAL / 4#)

Exit Function
ERROR_LABEL:
MARSAG_N2_VASICEK_SERIES_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : MARSAG_N2_VASICEK_SERIES_FUNC

'DESCRIPTION   : George Marsaglia, Evaluating the Normal Distribution,
' Journal of Statistical Software (2004),
' http://www.jstatsoft.org/counter.php?id=100&url=v11/i04/v11i04.pdf&ct=1
' for the univariate case, which is used to get the bivariate case here

'LIBRARY       : STATISTICS
'GROUP         : DIST_NORMAL_BIVAR
'ID            : 013
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/22/2009
'************************************************************************************
'************************************************************************************

Private Function N2_MARSAG_SERIES_FUNC( _
ByVal X1_VAL As Double, _
ByVal RHO_VAL As Double, _
ByVal nSTEPS As Integer) As Double
  
Dim ii As Integer

Dim TEMP_SUM As Double

Dim ATEMP_VAL As Double
Dim BTEMP_VAL As Double
Dim CTEMP_VAL As Double

Dim PI_VAL As Double
  
On Error GoTo ERROR_LABEL
  
PI_VAL = 3.14159265358979
CTEMP_VAL = CDbl(RHO_VAL / Sqr(1# - RHO_VAL * RHO_VAL))

ATEMP_VAL = 0#
BTEMP_VAL = 0#
TEMP_SUM = 0#

ATEMP_VAL = CDbl(-1# / (1# + CTEMP_VAL * CTEMP_VAL) * _
              Exp(-(1# + CTEMP_VAL * CTEMP_VAL) * X1_VAL * X1_VAL / 2#) / _
PI_VAL * CTEMP_VAL / 2#)
BTEMP_VAL = CDbl(CTEMP_VAL * CTEMP_VAL * X1_VAL * X1_VAL * ATEMP_VAL)
TEMP_SUM = CDbl(ATEMP_VAL)

For ii = 1 To nSTEPS - 1
  ATEMP_VAL = CDbl((2# * CDbl(ii) * CTEMP_VAL * CTEMP_VAL / _
              (1# + CTEMP_VAL * CTEMP_VAL) * _
              ATEMP_VAL + BTEMP_VAL) / CDbl(1 + 2 * ii))
  BTEMP_VAL = CDbl(CTEMP_VAL * CTEMP_VAL * X1_VAL * X1_VAL * _
              BTEMP_VAL / CDbl(1 + 2 * ii))
  TEMP_SUM = CDbl(TEMP_SUM + ATEMP_VAL)
  If (TEMP_SUM = TEMP_SUM - ATEMP_VAL) Then
    Exit For
  End If
Next

N2_MARSAG_SERIES_FUNC = TEMP_SUM + 0.5 * CDbl(MARSAG2_CUMUL_NORM_FUNC(X1_VAL))

Exit Function
ERROR_LABEL:
N2_MARSAG_SERIES_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : N2_ABS_RHO_FUNC
'DESCRIPTION   :
'LIBRARY       : STATISTICS
'GROUP         : DIST_NORMAL_BIVAR
'ID            : 014
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/22/2009
'************************************************************************************
'************************************************************************************

Private Function N2_ABS_RHO_FUNC( _
ByVal X1_VAL As Double, _
ByVal X2_VAL As Double, _
ByVal RHO_VAL As Double) As Double
  
On Error GoTo ERROR_LABEL

If (RHO_VAL = 1) Then
  If (X1_VAL <= X2_VAL) Then
    N2_ABS_RHO_FUNC = MARSAG2_CUMUL_NORM_FUNC(X1_VAL)
    Exit Function
  ElseIf (X2_VAL <= X1_VAL) Then
    N2_ABS_RHO_FUNC = MARSAG2_CUMUL_NORM_FUNC(X2_VAL)
    Exit Function
  Else
  End If
ElseIf (RHO_VAL = -1) Then
  If (-X2_VAL <= X1_VAL) Then
    N2_ABS_RHO_FUNC = MARSAG2_CUMUL_NORM_FUNC(X1_VAL) + _
                        MARSAG2_CUMUL_NORM_FUNC(X2_VAL) - 1
    Exit Function
  ElseIf (X1_VAL <= -X2_VAL) Then
    N2_ABS_RHO_FUNC = 0
    Exit Function
  Else
  End If
Else
End If

Exit Function
ERROR_LABEL:
N2_ABS_RHO_FUNC = Err.number
End Function
