Attribute VB_Name = "NUMBER_REAL_DEGREE_LIBR"


Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.


'************************************************************************************
'************************************************************************************
'FUNCTION      : MOD_CRS_FUNC
'DESCRIPTION   : Return value in range 0<value<=2*pi
'LIBRARY       : NUMBER_REAL
'GROUP         : REAL_DEGREE
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/08/2008
'************************************************************************************
'************************************************************************************

Function MOD_CRS_FUNC(ByVal X_VAL As Double)
Dim PI_VAL As Double
On Error GoTo ERROR_LABEL
PI_VAL = 3.14159265358979
MOD_CRS_FUNC = (2 * PI_VAL) - MOD_FUNC((2 * PI_VAL) - X_VAL, (2 * PI_VAL))
Exit Function
ERROR_LABEL:
MOD_CRS_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : MOD_LON_FUNC
'DESCRIPTION   : Return value in range -pi<=value<pi
'LIBRARY       : NUMBER_REAL
'GROUP         : REAL_DEGREE
'ID            : 002
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/08/2008
'************************************************************************************
'************************************************************************************

Function MOD_LON_FUNC(ByVal LON_VAL As Double)
Dim PI_VAL As Double
On Error GoTo ERROR_LABEL
PI_VAL = 3.14159265358979
MOD_LON_FUNC = MOD_FUNC(LON_VAL + PI_VAL, 2 * PI_VAL) - PI_VAL
Exit Function
ERROR_LABEL:
MOD_LON_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : DEGREES_TO_RADIANS_FUNC
'DESCRIPTION   : Converts degrees to radians
'LIBRARY       : NUMBER_REAL
'GROUP         : REAL_DEGREE
'ID            : 003
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/08/2008
'************************************************************************************
'************************************************************************************

Function DEGREES_TO_RADIANS_FUNC(ByVal X_VAL As Double)
    Dim PI_VAL As Double
    On Error GoTo ERROR_LABEL
    PI_VAL = 3.14159265358979
    DEGREES_TO_RADIANS_FUNC = (PI_VAL / 180) * X_VAL
Exit Function
ERROR_LABEL:
DEGREES_TO_RADIANS_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : ADJUST_DEGREES_FUNC
'DESCRIPTION   : Adjust Degrees
'LIBRARY       : NUMBER_REAL
'GROUP         : REAL_DEGREE
'ID            : 004
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/08/2008
'************************************************************************************
'************************************************************************************

Function ADJUST_DEGREES_FUNC(ByVal X_VAL As Double)
        
    Dim TEMP_VAL As Double
    
    On Error GoTo ERROR_LABEL
    
    TEMP_VAL = X_VAL
    Do While TEMP_VAL < 0
        TEMP_VAL = TEMP_VAL + 360
    Loop
    Do While TEMP_VAL > 360
        TEMP_VAL = TEMP_VAL - 360
    Loop
    If TEMP_VAL = 0 Then TEMP_VAL = 360
    
    ADJUST_DEGREES_FUNC = TEMP_VAL

Exit Function
ERROR_LABEL:
ADJUST_DEGREES_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : DISTANCE_TO_RADIANS_FUNC
'DESCRIPTION   : Multiplier to turn distance to radians
'LIBRARY       : NUMBER_REAL
'GROUP         : REAL_DEGREE
'ID            : 005
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/08/2008
'************************************************************************************
'************************************************************************************

Function DISTANCE_TO_RADIANS_FUNC(Optional VERSION As Integer = 1)

Dim PI_VAL As Double

On Error GoTo ERROR_LABEL

PI_VAL = 3.14159265358979

If (VERSION = 1) Then
  DISTANCE_TO_RADIANS_FUNC = PI_VAL / (180 * 60) 'nm
ElseIf (VERSION = 2) Then
  DISTANCE_TO_RADIANS_FUNC = PI_VAL / (180 * 60 * 1.852) 'km
Else 'If (VERSION = 3) Then
  DISTANCE_TO_RADIANS_FUNC = PI_VAL / (180 * 60 * 1.150779) 'statute miles
End If

Exit Function
ERROR_LABEL:
DISTANCE_TO_RADIANS_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : REDUCE_DEGREES_FUNC
'DESCRIPTION   : angle reduction
'LIBRARY       : NUMBER_REAL
'GROUP         : REAL_DEGREE
'ID            : 006
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/08/2008
'************************************************************************************
'************************************************************************************

Function REDUCE_DEGREES_FUNC(ByVal X_VAL As Double)
    
    Dim i As Long
    
    Dim ATEMP_VAL As Double
    Dim BTEMP_VAL As Double
    Dim CTEMP_VAL As Double
    Dim DTEMP_VAL As Double
    Dim ETEMP_VAL As Double
    
    Dim PI_VAL As Double
    Dim HALF_PI_VAL As Double
    
    On Error GoTo ERROR_LABEL
    
    PI_VAL = 3.14159265358979

    HALF_PI_VAL = PI_VAL / 2
    
    i = Int(X_VAL / HALF_PI_VAL)
    BTEMP_VAL = i - 4 * Int(i / 4) + 1
    CTEMP_VAL = CDec(X_VAL)
    
    If i <> 0 And (BTEMP_VAL = 1 Or BTEMP_VAL = 3) Then
        DTEMP_VAL = CTEMP_VAL / i
        ETEMP_VAL = (HALF_PI_VAL - DTEMP_VAL) 'complement to 90°
        ATEMP_VAL = (ETEMP_VAL * -i)
    ElseIf i <> -1 And (BTEMP_VAL = 2 Or BTEMP_VAL = 4) Then
        DTEMP_VAL = CTEMP_VAL / (i + 1)
        ETEMP_VAL = (HALF_PI_VAL - DTEMP_VAL) 'complement to 90°
        ATEMP_VAL = ETEMP_VAL * (i + 1)
    ElseIf i = 0 Then
        ATEMP_VAL = CTEMP_VAL
    ElseIf i = -1 Then
        ATEMP_VAL = -CTEMP_VAL
    End If
    
    'ATEMP_VAL 'angle reduced output
    'BTEMP_VAL 'quadrante  output
    
    REDUCE_DEGREES_FUNC = Array(ATEMP_VAL, BTEMP_VAL)

Exit Function
ERROR_LABEL:
REDUCE_DEGREES_FUNC = Err.number
End Function
