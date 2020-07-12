Attribute VB_Name = "GEOMETRY_MANDELBROT_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.

Function MANDELBROT_FUNC( _
Optional ByVal XMIN_VAL As Double = -2, _
Optional ByVal XDELTA_VAL As Double = 3 / 100, _
Optional ByVal XBINS_VAL As Long = 100, _
Optional ByVal YMIN_VAL As Double = -1, _
Optional ByVal YDELTA_VAL As Double = 2 / 100, _
Optional ByVal YBINS_VAL As Long = 100, _
Optional ByVal nLOOPS As Long = 1000, _
Optional ByVal POWER_VAL As Double = 2, _
Optional ByVal ESCAPE_VAL As Double = 4)

Dim i As Long
Dim j As Long
Dim k As Long
Dim l As Variant

Dim XT_VAL As Double
Dim YT_VAL As Double

Dim X0_VAL As Double
Dim Y0_VAL As Double

Dim X1_VAL As Double
Dim Y1_VAL As Double

Dim RSQ_VAL As Double

Dim MAX_VAL As Double
Dim TEMP_MATRIX() As Double

'Const epsilon As Double = 10 ^ -15
On Error GoTo ERROR_LABEL

ReDim TEMP_MATRIX(0 To YBINS_VAL + 1, 0 To XBINS_VAL + 1)
For j = 0 To XBINS_VAL
    XT_VAL = XMIN_VAL + j * XDELTA_VAL
    'XT_VAL = XT_VAL - epsilon
    TEMP_MATRIX(0, j + 1) = XT_VAL
    For i = 0 To YBINS_VAL
        YT_VAL = YMIN_VAL + i * YDELTA_VAL
        'YT_VAL = YT_VAL - epsilon
        If j = 0 Then: TEMP_MATRIX(i + 1, 0) = YT_VAL
        GoSub RSQ_LINE
        TEMP_MATRIX(i + 1, j + 1) = l
    Next i
Next j
TEMP_MATRIX(0, 0) = l

MANDELBROT_FUNC = TEMP_MATRIX

Exit Function
'-----------------------------------------------------------------------------------
RSQ_LINE:
'-----------------------------------------------------------------------------------
    k = 1: l = "": MAX_VAL = -2 ^ 52
    X0_VAL = 0: Y0_VAL = 0
    RSQ_VAL = X0_VAL ^ 2 + Y0_VAL ^ 2 'RSQ
    MAX_VAL = RSQ_VAL: l = k
    For k = 2 To nLOOPS
        If RSQ_VAL > ESCAPE_VAL Then: Return
        X1_VAL = X0_VAL: Y1_VAL = Y0_VAL
        X0_VAL = X1_VAL ^ POWER_VAL - Y1_VAL ^ POWER_VAL + XT_VAL
        Y0_VAL = 2 * X1_VAL * Y1_VAL + YT_VAL
        RSQ_VAL = X0_VAL ^ 2 + Y0_VAL ^ 2 'RSQ
        If RSQ_VAL > MAX_VAL Then
            MAX_VAL = RSQ_VAL: l = k
        End If
    Next k
'-----------------------------------------------------------------------------------
Return
'-----------------------------------------------------------------------------------
ERROR_LABEL:
MANDELBROT_FUNC = Err.number
End Function


