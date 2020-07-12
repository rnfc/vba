Attribute VB_Name = "FINAN_PORT_MOMENTS_DISPERS_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.


'************************************************************************************
'************************************************************************************
'FUNCTION      : PORT_ASSET_WEIGHTED_STDEV_FUNC
'DESCRIPTION   : Calculate the asset-weighted standard deviation ("dispersion")
'LIBRARY       : PORTFOLIO
'GROUP         : MOMENTS_DISPERSION
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/08/2008
'************************************************************************************
'************************************************************************************

Function PORT_ASSET_WEIGHTED_STDEV_FUNC(ByRef WEIGHTS_RNG As Variant, _
ByRef RETURNS_RNG As Variant, _
Optional ByVal OUTPUT As Integer = 0)

Dim i As Long
Dim j As Long
Dim NROWS As Long

Dim TEMP1_SUM As Double
Dim TEMP2_SUM As Double
Dim TEMP3_SUM As Double
Dim TEMP4_SUM As Double
Dim TEMP5_SUM As Double

Dim TEMP_MATRIX As Variant
Dim WEIGHTS_VECTOR As Variant
Dim RETURNS_VECTOR As Variant

On Error GoTo ERROR_LABEL


WEIGHTS_VECTOR = WEIGHTS_RNG
If UBound(WEIGHTS_VECTOR, 1) = 1 Then
    WEIGHTS_VECTOR = MATRIX_TRANSPOSE_FUNC(WEIGHTS_VECTOR)
End If

RETURNS_VECTOR = RETURNS_RNG
If UBound(RETURNS_VECTOR, 1) = 1 Then
    RETURNS_VECTOR = MATRIX_TRANSPOSE_FUNC(RETURNS_VECTOR)
End If

If UBound(WEIGHTS_VECTOR, 1) <> UBound(RETURNS_VECTOR, 1) Then: GoTo ERROR_LABEL
NROWS = UBound(WEIGHTS_VECTOR, 1)

'------------------------------------------------------------------------------------
If OUTPUT = 0 Then
'------------------------------------------------------------------------------------

    ReDim TEMP_MATRIX(0 To NROWS + 3, 1 To 7)
    
    TEMP_MATRIX(0, 1) = "WEIGHTS"
    TEMP_MATRIX(0, 2) = "RETURNS"
    TEMP_MATRIX(0, 3) = "WEIGHTED RETURNS"
    TEMP_MATRIX(0, 4) = "CONTRIBUTIONS"
    TEMP_MATRIX(0, 5) = "SQR.DEV FROM EQUAL-WEIGHTED MEAN"
    TEMP_MATRIX(0, 6) = "SQR.DEV FROM ASSET-WEIGHTED MEAN"
    TEMP_MATRIX(0, 7) = "WEIGHTED SDAWM"
    
    TEMP3_SUM = 0
    For i = 1 To NROWS
        TEMP3_SUM = TEMP3_SUM + WEIGHTS_VECTOR(i, 1)
    Next i
    
    TEMP4_SUM = 0
    TEMP5_SUM = 0
    For i = 1 To NROWS
        TEMP_MATRIX(i, 1) = WEIGHTS_VECTOR(i, 1)
        
        TEMP_MATRIX(i, 2) = RETURNS_VECTOR(i, 1)
        TEMP4_SUM = TEMP4_SUM + TEMP_MATRIX(i, 2)
        
        TEMP_MATRIX(i, 3) = TEMP_MATRIX(i, 1) / TEMP3_SUM
        
        TEMP_MATRIX(i, 4) = TEMP_MATRIX(i, 2) * TEMP_MATRIX(i, 3)
        TEMP5_SUM = TEMP5_SUM + TEMP_MATRIX(i, 4)
    Next i
    
    For i = 1 To NROWS
        TEMP_MATRIX(i, 5) = (TEMP_MATRIX(i, 2) - (TEMP4_SUM / NROWS)) ^ 2
        TEMP_MATRIX(i, 6) = (TEMP_MATRIX(i, 2) - TEMP5_SUM) ^ 2
        TEMP_MATRIX(i, 7) = TEMP_MATRIX(i, 3) * TEMP_MATRIX(i, 6)
    Next i
    
    For i = 1 To 7
        TEMP_MATRIX(NROWS + 1, i) = ""
        TEMP_MATRIX(NROWS + 2, i) = 0
    Next i


    For i = 1 To NROWS
        For j = 1 To 7: TEMP_MATRIX(NROWS + 2, j) = TEMP_MATRIX(NROWS + 2, j) + TEMP_MATRIX(i, j): Next j
    Next i

    TEMP_MATRIX(NROWS + 3, 1) = TEMP_MATRIX(NROWS + 2, 1) / NROWS
    TEMP_MATRIX(NROWS + 3, 2) = TEMP_MATRIX(NROWS + 2, 2) / NROWS
    TEMP_MATRIX(NROWS + 3, 3) = TEMP_MATRIX(NROWS + 2, 3) / NROWS
    TEMP_MATRIX(NROWS + 3, 4) = TEMP_MATRIX(NROWS + 2, 4) / NROWS
    TEMP_MATRIX(NROWS + 3, 5) = (TEMP_MATRIX(NROWS + 2, 5) / NROWS) ^ 0.5
    TEMP_MATRIX(NROWS + 3, 6) = NROWS
    TEMP_MATRIX(NROWS + 3, 7) = TEMP_MATRIX(NROWS + 2, 7) ^ 0.5
    
    PORT_ASSET_WEIGHTED_STDEV_FUNC = TEMP_MATRIX

'------------------------------------------------------------------------------------
Else
'------------------------------------------------------------------------------------
    TEMP3_SUM = 0
    For i = 1 To NROWS: TEMP3_SUM = TEMP3_SUM + WEIGHTS_VECTOR(i, 1): Next i
    TEMP1_SUM = 0
    For i = 1 To NROWS     ' compute asset-weighted average return
        TEMP1_SUM = TEMP1_SUM + RETURNS_VECTOR(i, 1) * WEIGHTS_VECTOR(i, 1) / TEMP3_SUM
    Next i
    ' compute asset-weighted deviations
    TEMP2_SUM = 0
    For i = 1 To NROWS
        TEMP2_SUM = TEMP2_SUM + (WEIGHTS_VECTOR(i, 1) / TEMP3_SUM) * (RETURNS_VECTOR(i, 1) - TEMP1_SUM) ^ 2
    Next i

    PORT_ASSET_WEIGHTED_STDEV_FUNC = TEMP2_SUM ^ 0.5
'------------------------------------------------------------------------------------
End If
'------------------------------------------------------------------------------------

Exit Function
ERROR_LABEL:
PORT_ASSET_WEIGHTED_STDEV_FUNC = Err.number
End Function


