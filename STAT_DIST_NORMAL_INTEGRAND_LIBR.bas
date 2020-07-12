Attribute VB_Name = "STAT_DIST_NORMAL_INTEGRAND_LIBR"

'---------------------------------------------------------------------------------------
'---------------------------------------------------------------------------------------
'---------------------------------------------------------------------------------------
Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.
'---------------------------------------------------------------------------------------
'---------------------------------------------------------------------------------------

Private CDFN_G_Y As Double
Private CDFN_G_RHO As Double

Private CDFN_G_X1_VAL As Double
Private CDFN_G_X2_VAL As Double
Private CDFN_G_X3_VAL As Double

Private CDFN_G_RHO12_VAL As Double
Private CDFN_G_RHO13_VAL As Double
Private CDFN_G_RHO23_VAL As Double

'---------------------------------------------------------------------------------
'---------------------------------------------------------------------------------

'************************************************************************************
'************************************************************************************
'FUNCTION      : CDFN_2_INTEGRAND_FUNC
'DESCRIPTION   :
'LIBRARY       : STATISTICS
'GROUP         : DIST_NORMAL_INTEGRAND
'ID            : 003
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/22/2009
'************************************************************************************
'************************************************************************************

Private Function CDFN_2_INTEGRAND_FUNC(ByVal X_TEMP_VAL As Double) As Double
Dim Z_TEMP_VAL As Double
On Error GoTo ERROR_LABEL
Z_TEMP_VAL = (CDFN_G_Y - CDFN_G_RHO * (-X_TEMP_VAL)) / Sqr(1 - CDFN_G_RHO ^ 2)

CDFN_2_INTEGRAND_FUNC = Exp(-0.5 * X_TEMP_VAL * X_TEMP_VAL - 0.918938533204673) _
                  * UNIVAR_CUMUL_NORM_FUNC("", Z_TEMP_VAL)

Exit Function
ERROR_LABEL:
CDFN_2_INTEGRAND_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : CDFN_3_INTEGRAND_FUNC
'DESCRIPTION   :
'LIBRARY       : STATISTICS
'GROUP         : DIST_NORMAL_INTEGRAND
'ID            : 004
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/22/2009
'************************************************************************************
'************************************************************************************

Private Function CDFN_3_INTEGRAND_FUNC(ByVal X_TEMP_VAL As Double) As Double
On Error GoTo ERROR_LABEL
CDFN_3_INTEGRAND_FUNC = Exp(-0.5 * -X_TEMP_VAL * -X_TEMP_VAL - 0.918938533204673) * _
                  CDFN_BIVAR_FUNC(-X_TEMP_VAL)
Exit Function
ERROR_LABEL:
CDFN_3_INTEGRAND_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : CDFN_BIVAR_FUNC
'DESCRIPTION   :
'LIBRARY       : STATISTICS
'GROUP         : DIST_NORMAL_INTEGRAND
'ID            : 007
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/22/2009
'************************************************************************************
'************************************************************************************

Private Function CDFN_BIVAR_FUNC(ByVal x As Double) As Double

Dim RHO_VAL As Double
Dim ATEMP_VAL As Double
Dim BTEMP_VAL As Double

On Error GoTo ERROR_LABEL

ATEMP_VAL = (CDFN_G_X2_VAL - CDFN_G_RHO12_VAL * x) / Sqr(1 - CDFN_G_RHO12_VAL ^ 2)
BTEMP_VAL = (CDFN_G_X3_VAL - CDFN_G_RHO13_VAL * x) / Sqr(1 - CDFN_G_RHO13_VAL ^ 2)
RHO_VAL = (CDFN_G_RHO23_VAL - CDFN_G_RHO13_VAL * CDFN_G_RHO12_VAL) / _
           Sqr(1 - CDFN_G_RHO12_VAL ^ 2) / Sqr(1 - CDFN_G_RHO13_VAL ^ 2)

CDFN_BIVAR_FUNC = BIVAR_CUMUL_NORM_FUNC("genz", "", ATEMP_VAL, BTEMP_VAL, RHO_VAL)

Exit Function
ERROR_LABEL:
CDFN_BIVAR_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : CDFN_3_INTEGR1_FUNC
'DESCRIPTION   :
'LIBRARY       : STATISTICS
'GROUP         : DIST_NORMAL_INTEGRAND
'ID            : 008
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/22/2009
'************************************************************************************
'************************************************************************************

Private Function CDFN_3_INTEGR1_FUNC(ByVal X_TEMP_VAL As Double) As Double

Dim ATEMP_VAL As Double
Dim BTEMP_VAL As Double
Dim RHO_VAL As Double

On Error GoTo ERROR_LABEL

ATEMP_VAL = (CDFN_G_X2_VAL - CDFN_G_RHO12_VAL * _
            (-X_TEMP_VAL)) / Sqr(1 - CDFN_G_RHO12_VAL ^ 2)

BTEMP_VAL = (CDFN_G_X3_VAL - CDFN_G_RHO13_VAL * _
            (-X_TEMP_VAL)) / Sqr(1 - CDFN_G_RHO13_VAL ^ 2)
RHO_VAL = (CDFN_G_RHO23_VAL - CDFN_G_RHO13_VAL * CDFN_G_RHO12_VAL) / _
           Sqr(1 - CDFN_G_RHO12_VAL ^ 2) / Sqr(1 - CDFN_G_RHO13_VAL ^ 2)

CDFN_3_INTEGR1_FUNC = Exp(-0.5 * X_TEMP_VAL * X_TEMP_VAL - 0.918938533204673) * _
                        BIVAR_CUMUL_NORM_FUNC("genz", "", ATEMP_VAL, _
                        BTEMP_VAL, RHO_VAL)
                        
Exit Function
ERROR_LABEL:
CDFN_3_INTEGR1_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : CDFN_3_INTEGR2_FUNC
'DESCRIPTION   :
'LIBRARY       : STATISTICS
'GROUP         : DIST_NORMAL_INTEGRAND
'ID            : 009
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/22/2009
'************************************************************************************
'************************************************************************************

Private Function CDFN_3_INTEGR2_FUNC(ByVal Z_TEMP_VAL As Double) As Double

On Error GoTo ERROR_LABEL

CDFN_3_INTEGR2_FUNC = Exp(-Z_TEMP_VAL * Z_TEMP_VAL / 2 + _
                        Z_TEMP_VAL * CDFN_G_X1_VAL) * _
                        CDFN_BIVAR_FUNC(CDFN_G_X1_VAL - Z_TEMP_VAL)
Exit Function
ERROR_LABEL:
CDFN_3_INTEGR2_FUNC = Err.number
End Function
