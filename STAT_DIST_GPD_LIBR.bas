Attribute VB_Name = "STAT_DIST_GPD_LIBR"


Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.

'************************************************************************************
'************************************************************************************
'FUNCTION      : GPD_DIST_FUNC
'DESCRIPTION   : Helper function in the analysis of discrete distribution
'LIBRARY       : STATISTICS
'GROUP         : DIST_GPD
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/22/2009
'************************************************************************************
'************************************************************************************

Function GPD_DIST_FUNC(ByRef DATA_RNG As Variant, _
ByVal THRESHOLD As Double)

Dim i As Long
Dim j As Long

Dim NROWS As Long

Dim TEMP_VAL As Double
Dim TEMP_SUM As Double

Dim TEMP_VECTOR As Variant
Dim DATA_VECTOR As Variant

On Error GoTo ERROR_LABEL

DATA_VECTOR = DATA_RNG
If UBound(DATA_VECTOR, 1) = 1 Then: DATA_VECTOR = MATRIX_TRANSPOSE_FUNC(DATA_VECTOR)
DATA_VECTOR = MATRIX_QUICK_SORT_FUNC(DATA_VECTOR, 1, 1)
NROWS = UBound(DATA_VECTOR, 1)

ReDim TEMP_VECTOR(1 To NROWS, 1 To 2)
j = 0
TEMP_SUM = 0
For i = 1 To NROWS
    TEMP_VAL = DATA_VECTOR(i, 1)
    If TEMP_VAL <= THRESHOLD Then
        TEMP_VECTOR(i, 1) = TEMP_VAL
        j = j + 1
        TEMP_SUM = TEMP_SUM + TEMP_VAL
    Else
        TEMP_VECTOR(i, 1) = "N/A"
        TEMP_VECTOR(i, 2) = "N/A"
    End If
Next i

For i = 1 To j
    TEMP_VECTOR(i, 2) = i / (j + 1)
Next i

GPD_DIST_FUNC = TEMP_VECTOR
Exit Function
ERROR_LABEL:
GPD_DIST_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : GPD_PDF_FUNC
'DESCRIPTION   : GPD pdf distribution
'LIBRARY       : STATISTICS
'GROUP         : DIST_GPD
'ID            : 002
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/22/2009
'************************************************************************************
'************************************************************************************

Function GPD_PDF_FUNC(ByVal X_VAL As Double, _
ByVal SHAPE_VAL As Double, _
ByVal SCALE_VAL As Double, _
ByVal LOCATION_VAL As Double)

' validate domain

On Error GoTo ERROR_LABEL

If SHAPE_VAL >= 0 Then
    If Not (LOCATION_VAL <= X_VAL) Then: GoTo ERROR_LABEL
Else
    If Not ((LOCATION_VAL <= X_VAL) And (X_VAL <= LOCATION_VAL - (SCALE_VAL / SHAPE_VAL))) Then: GoTo ERROR_LABEL
End If

' calc values
If SHAPE_VAL = 0 Then
    GPD_PDF_FUNC = (1 / SCALE_VAL) * Exp(-(X_VAL - LOCATION_VAL) / SCALE_VAL)
Else
    GPD_PDF_FUNC = (1 / SCALE_VAL) * (1 + SHAPE_VAL * (X_VAL - LOCATION_VAL) / SCALE_VAL) ^ (-1 - 1 / SHAPE_VAL)
End If

Exit Function
ERROR_LABEL:
GPD_PDF_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : GPD_CDF_FUNC
'DESCRIPTION   : GPD cdf distribution
'LIBRARY       : STATISTICS
'GROUP         : DIST_GPD
'ID            : 003
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/22/2009
'************************************************************************************
'************************************************************************************

Function GPD_CDF_FUNC(ByVal X_VAL As Double, _
ByVal SHAPE_VAL As Double, _
ByVal SCALE_VAL As Double, _
ByVal LOCATION_VAL As Double)

'--------------------------------------------------------------------------
'----------------------------Validate domain-------------------------------
'--------------------------------------------------------------------------

On Error GoTo ERROR_LABEL

If SHAPE_VAL >= 0 Then
    If Not (LOCATION_VAL <= X_VAL) Then: GoTo ERROR_LABEL
Else
    If Not ((LOCATION_VAL <= X_VAL) And _
            (X_VAL <= _
            LOCATION_VAL - SCALE_VAL / _
            SHAPE_VAL)) Then: GoTo ERROR_LABEL
End If

'--------------------------------------------------------------------------
'-----------------------------calc values----------------------------------
'--------------------------------------------------------------------------

If SHAPE_VAL = 0 Then
    GPD_CDF_FUNC = 1 - Exp(-(X_VAL - LOCATION_VAL) / SCALE_VAL)
Else
    GPD_CDF_FUNC = 1 - (1 + SHAPE_VAL * (X_VAL - _
                LOCATION_VAL) / SCALE_VAL) ^ (-1 / SHAPE_VAL)
End If

Exit Function
ERROR_LABEL:
GPD_CDF_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : GPD_INV_CDF_FUNC
'DESCRIPTION   : GPD inverse cdf distribution
'LIBRARY       : STATISTICS
'GROUP         : DIST_GPD
'ID            : 004
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/22/2009
'************************************************************************************
'************************************************************************************

Function GPD_INV_CDF_FUNC(ByVal X_VAL As Double, _
ByVal SHAPE_VAL As Double, _
ByVal SCALE_VAL As Double, _
ByVal LOCATION_VAL As Double)

' calc values
On Error GoTo ERROR_LABEL

If SHAPE_VAL = 0 Then
    GPD_INV_CDF_FUNC = LOCATION_VAL - SCALE_VAL * Log(1 - X_VAL)
Else
    GPD_INV_CDF_FUNC = LOCATION_VAL + (SCALE_VAL / _
                SHAPE_VAL) * ((1 - X_VAL) ^ (-SHAPE_VAL) - 1)
End If
Exit Function
ERROR_LABEL:
GPD_INV_CDF_FUNC = Err.number
End Function

