Attribute VB_Name = "NUMBER_REAL_AREA_VOLUME_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.


'************************************************************************************
'************************************************************************************
'FUNCTION      : AREA_CIRCLE_FUNC
'DESCRIPTION   : Aire d'un cercle à partir du rayon
'LIBRARY       : NUMBER_REAL
'GROUP         : AREA_VOLUME
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/08/2008
'************************************************************************************
'************************************************************************************

Function AREA_CIRCLE_FUNC(ByVal RAD_VAL As Double)
On Error GoTo ERROR_LABEL
AREA_CIRCLE_FUNC = RAD_VAL * RAD_VAL * (Atn(1) * 4)
Exit Function
ERROR_LABEL:
AREA_CIRCLE_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : AREA_RECT_FUNC
'DESCRIPTION   : Aire d'un rectangle
'LIBRARY       : NUMBER_REAL
'GROUP         : AREA_VOLUME
'ID            : 002
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/08/2008
'************************************************************************************
'************************************************************************************

Function AREA_RECT_FUNC(ByVal L_VAL As Double, _
ByVal W_VAL As Double)
On Error GoTo ERROR_LABEL
AREA_RECT_FUNC = L_VAL * W_VAL
Exit Function
ERROR_LABEL:
AREA_RECT_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : AREA_RING_FUNC
'DESCRIPTION   : Aire d'un anneau définit à partir de 2 rayons
'LIBRARY       : NUMBER_REAL
'GROUP         : AREA_VOLUME
'ID            : 003
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/08/2008
'************************************************************************************
'************************************************************************************


Function AREA_RING_FUNC(ByVal INT_VAL As Double, _
ByVal EXT_VAL As Double)

On Error GoTo ERROR_LABEL
  AREA_RING_FUNC = AREA_CIRCLE_FUNC(EXT_VAL) - AREA_CIRCLE_FUNC(INT_VAL)
Exit Function
ERROR_LABEL:
AREA_RING_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : AREA_SPHERE_FUNC
'DESCRIPTION   : Aire d'une sphère à partir du rayon
'LIBRARY       : NUMBER_REAL
'GROUP         : AREA_VOLUME
'ID            : 004
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/08/2008
'************************************************************************************
'************************************************************************************


Function AREA_SPHERE_FUNC(ByVal R_VAL As Double)
On Error GoTo ERROR_LABEL
  AREA_SPHERE_FUNC = 4 * (Atn(1) * 4) * R_VAL * R_VAL
Exit Function
ERROR_LABEL:
AREA_SPHERE_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : AREA_CUBE_FUNC
'DESCRIPTION   : Aire du carré en fonction du s
'LIBRARY       : NUMBER_REAL
'GROUP         : AREA_VOLUME
'ID            : 005
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/08/2008
'************************************************************************************
'************************************************************************************

Function AREA_CUBE_FUNC(ByVal S_VAL As Double)

On Error GoTo ERROR_LABEL
  AREA_CUBE_FUNC = S_VAL * S_VAL
Exit Function
ERROR_LABEL:
AREA_CUBE_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : AREA_TRAP_FUNC
'DESCRIPTION   : Aire du Trapèze à partir de la longueur des côtés parallèles
' et de la hauteur perpendiculaire
'LIBRARY       : NUMBER_REAL
'GROUP         : AREA_VOLUME
'ID            : 006
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/08/2008
'************************************************************************************
'************************************************************************************

Function AREA_TRAP_FUNC(ByVal H_VAL As Double, _
ByVal LENGTH_1_VAL As Double, _
ByVal LENGTH_2_VAL As Double)
'
On Error GoTo ERROR_LABEL
  AREA_TRAP_FUNC = H_VAL * (LENGTH_1_VAL + LENGTH_2_VAL) / 2
Exit Function
ERROR_LABEL:
AREA_TRAP_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : AREA_TRIANG1_FUNC
'DESCRIPTION   : Aire du triangle à partir de la longueur d'un
' côté et de la hauteur perpendiculaire
'LIBRARY       : NUMBER_REAL
'GROUP         : AREA_VOLUME
'ID            : 007
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/08/2008
'************************************************************************************
'************************************************************************************

Function AREA_TRIANG1_FUNC(ByVal L_VAL As Double, _
ByVal H_VAL As Double)
'
On Error GoTo ERROR_LABEL
  AREA_TRIANG1_FUNC = L_VAL * H_VAL / 2
Exit Function
ERROR_LABEL:
AREA_TRIANG1_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : AREA_TRIANG2_FUNC
'DESCRIPTION   : Aire du triangle à partir de la longueur des 3 côtés
'LIBRARY       : NUMBER_REAL
'GROUP         : AREA_VOLUME
'ID            : 008
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/08/2008
'************************************************************************************
'************************************************************************************

Function AREA_TRIANG2_FUNC(ByVal A_TEMP As Double, _
ByVal B_VAL As Double, _
ByVal C_VAL As Double)

Dim COSC_VAL As Double

On Error GoTo ERROR_LABEL
  COSC_VAL = (A_TEMP * A_TEMP + B_VAL * B_VAL - _
                C_VAL * C_VAL) / (2 * A_TEMP * B_VAL)
  
  AREA_TRIANG2_FUNC = A_TEMP * B_VAL * Sqr(1 - COSC_VAL * COSC_VAL) / 2

Exit Function
ERROR_LABEL:
AREA_TRIANG2_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : VOL_CONE_FUNC
'DESCRIPTION   : Volume d'un cône en fonction du rayon de sa base et de sa hauteur
'LIBRARY       : NUMBER_REAL
'GROUP         : AREA_VOLUME
'ID            : 009
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/08/2008
'************************************************************************************
'************************************************************************************

Function VOL_CONE_FUNC(ByVal H_VAL As Double, _
ByVal R_VAL As Double)
'
On Error GoTo ERROR_LABEL
  VOL_CONE_FUNC = H_VAL * R_VAL * R_VAL * (Atn(1) * 4) / 3
Exit Function
ERROR_LABEL:
VOL_CONE_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : VOL_CYLIN_FUNC
'DESCRIPTION   : Volume d'un Cylindre en fonction de sa hauteur et du rayon
'LIBRARY       : NUMBER_REAL
'GROUP         : AREA_VOLUME
'ID            : 010
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/08/2008
'************************************************************************************
'************************************************************************************


Function VOL_CYLIN_FUNC(ByVal H_VAL As Double, _
ByVal R_VAL As Double)

On Error GoTo ERROR_LABEL
  VOL_CYLIN_FUNC = (Atn(1) * 4) * R_VAL * R_VAL * H_VAL
Exit Function
ERROR_LABEL:
VOL_CYLIN_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : VOL_PIPE_FUNC
'DESCRIPTION   : Volume d'un Tuyau en soustrayant 2 cylindres
'LIBRARY       : NUMBER_REAL
'GROUP         : AREA_VOLUME
'ID            : 011
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/08/2008
'************************************************************************************
'************************************************************************************

Function VOL_PIPE_FUNC(ByVal H_VAL As Double, _
ByVal EXT_VAL As Double, _
ByVal INT_VAL As Double)
On Error GoTo ERROR_LABEL
VOL_PIPE_FUNC = VOL_CYLIN_FUNC(H_VAL, EXT_VAL) - VOL_CYLIN_FUNC(H_VAL, INT_VAL)
Exit Function
ERROR_LABEL:
VOL_PIPE_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : VOL_PYRAM_FUNC
'DESCRIPTION   : Volume d'une pyramide ou d'un cône en fonction de l'aire de sa
' base et de sa hauteur
'LIBRARY       : NUMBER_REAL
'GROUP         : AREA_VOLUME
'ID            : 012
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/08/2008
'************************************************************************************
'************************************************************************************

Function VOL_PYRAM_FUNC(ByVal H_VAL As Double, _
ByVal BASE_VAL As Double)
On Error GoTo ERROR_LABEL
VOL_PYRAM_FUNC = H_VAL * BASE_VAL / 3
Exit Function
ERROR_LABEL:
VOL_PYRAM_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : VOL_SPHERE_FUNC
'DESCRIPTION   : Volume d'une sphère en fonction de son rayon
'LIBRARY       : NUMBER_REAL
'GROUP         : AREA_VOLUME
'ID            : 013
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/08/2008
'************************************************************************************
'************************************************************************************


Function VOL_SPHERE_FUNC(ByVal R_VAL As Double)
On Error GoTo ERROR_LABEL
VOL_SPHERE_FUNC = (Atn(1) * 4) * R_VAL * R_VAL * R_VAL * 4 / 3
Exit Function
ERROR_LABEL:
VOL_SPHERE_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : VOL_PYRAM_TRUNC_FUNC
'DESCRIPTION   : Volume d'une triangle
'LIBRARY       : NUMBER_REAL
'GROUP         : AREA_VOLUME
'ID            : 014
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/08/2008
'************************************************************************************
'************************************************************************************

Function VOL_PYRAM_TRUNC_FUNC(ByVal H_VAL As Double, _
ByVal BASE_1_VAL As Double, _
ByVal BASE_2_VAL As Double)
On Error GoTo ERROR_LABEL
VOL_PYRAM_TRUNC_FUNC = H_VAL * (BASE_1_VAL + BASE_2_VAL + Sqr(BASE_1_VAL) * Sqr(BASE_2_VAL)) / 3
Exit Function
ERROR_LABEL:
VOL_PYRAM_TRUNC_FUNC = Err.number
End Function
