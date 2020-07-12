Attribute VB_Name = "STAT_REGRESSION_MIN_MAX_LIBR"


Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.


'************************************************************************************
'************************************************************************************
'FUNCTION      : REGRESSION_MIN_MAX_FUNC

'DESCRIPTION   : This function performs the linear regression with the
'Min-Max criterion (also called Chebycev approximation)
'of a discrete dataset (x, y)

'Returns the Min-Max Linear Regression Polynomial (Cebychev)
'Uses the exchange algorithm

'LIBRARY       : STATISTICS
'GROUP         : REGRESSION_MIN_MAX
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/22/2009
'************************************************************************************
'************************************************************************************

Function REGRESSION_MIN_MAX_FUNC(ByRef XDATA_RNG As Variant, _
ByRef YDATA_RNG As Variant, _
Optional ByVal nLOOPS As Long = 1000)

Dim i As Long
Dim k As Long
Dim l As Long

Dim NROWS As Long
Dim NCOLUMNS As Long

Dim X1_VAL As Double
Dim X2_VAL As Double
Dim X3_VAL As Double

Dim Y1_VAL As Double
Dim Y2_VAL As Double
Dim Y3_VAL As Double

Dim A_VAL As Double
Dim B_VAL As Double
Dim C_VAL As Double

Dim YDATA_VECTOR As Variant
Dim XDATA_VECTOR As Variant
Dim TEMP_VECTOR As Variant

Dim FIRST_SCALAR As Double
Dim SECOND_SCALAR As Double

Dim epsilon As Double
Dim tolerance As Double

On Error GoTo ERROR_LABEL

XDATA_VECTOR = XDATA_RNG
If UBound(XDATA_VECTOR, 1) = 1 Then: _
    XDATA_VECTOR = MATRIX_TRANSPOSE_FUNC(XDATA_VECTOR)

YDATA_VECTOR = YDATA_RNG
If UBound(YDATA_VECTOR, 1) = 1 Then: _
    YDATA_VECTOR = MATRIX_TRANSPOSE_FUNC(YDATA_VECTOR)

NROWS = UBound(YDATA_VECTOR, 1)

If NROWS < 3 Or NROWS <> UBound(XDATA_VECTOR, 1) Then: GoTo ERROR_LABEL

epsilon = 10 ^ -15

ReDim TEMP_VECTOR(1 To NROWS, 1 To 1)
NCOLUMNS = UBound(XDATA_VECTOR, 2)

A_VAL = 1
B_VAL = 2
C_VAL = 3

X1_VAL = XDATA_VECTOR(A_VAL, 1)
X2_VAL = XDATA_VECTOR(B_VAL, 1)
X3_VAL = XDATA_VECTOR(C_VAL, 1)
Y1_VAL = YDATA_VECTOR(A_VAL, 1)
Y2_VAL = YDATA_VECTOR(B_VAL, 1)
Y3_VAL = YDATA_VECTOR(C_VAL, 1)

l = nLOOPS

Do
    NCOLUMNS = (Y1_VAL - Y3_VAL) / (X1_VAL - X3_VAL)
    FIRST_SCALAR = ((Y1_VAL + 2 * Y2_VAL + Y3_VAL) - NCOLUMNS * _
        (X1_VAL + 2 * X2_VAL + X3_VAL)) / 4
    SECOND_SCALAR = ((Y1_VAL - 2 * Y2_VAL + Y3_VAL) - NCOLUMNS * _
        (X1_VAL - 2 * X2_VAL + X3_VAL)) / 4
    
'----------------To compute the line of min-max error (Cebysev) use----------
'NCOLUMNS = (Y1_VAL - Y3_VAL) / (X1_VAL - X3_VAL)
'FIRST_SCALAR = ((Y1_VAL + 2 * Y2_VAL + Y3_VAL) - NCOLUMNS * (X1_VAL + 2 * X2_VAL + X3_VAL)) / 4
'SECOND_SCALAR = ((Y1_VAL - 2 * Y2_VAL + Y3_VAL) - NCOLUMNS * (X1_VAL - 2 * X2_VAL + X3_VAL)) / 4
'----------------------------------------------------------------------------
    
    'search for max error point
    tolerance = 0
    k = 0
    For i = 1 To NROWS
        TEMP_VECTOR(i, 1) = YDATA_VECTOR(i, 1) - _
            NCOLUMNS * XDATA_VECTOR(i, 1) - FIRST_SCALAR
        If Abs(TEMP_VECTOR(i, 1)) > tolerance Then
            tolerance = Abs(TEMP_VECTOR(i, 1))
            k = i
        End If
    Next
    
    'check end algorithm
    If tolerance <= Abs(SECOND_SCALAR) + epsilon Then Exit Do
    'insert the new Pmax value into the tree point set
    '[P1,P2,P3] and eliminate one
    If XDATA_VECTOR(k, 1) > X3_VAL Then 'case  P1<P2<P3<Pmax
        If Sgn(TEMP_VECTOR(k, 1)) <> _
            Sgn(TEMP_VECTOR(C_VAL, 1)) Then 'shift left
            X1_VAL = X2_VAL
            X2_VAL = X3_VAL
            Y1_VAL = Y2_VAL
            Y2_VAL = Y3_VAL
            A_VAL = B_VAL
            B_VAL = C_VAL
        End If
        X3_VAL = XDATA_VECTOR(k, 1)
        Y3_VAL = YDATA_VECTOR(k, 1)
        C_VAL = k

    ElseIf XDATA_VECTOR(k, 1) < X1_VAL Then  'case  Pmax<P1<P2<P3
        If Sgn(TEMP_VECTOR(k, 1)) <> _
            Sgn(TEMP_VECTOR(A_VAL, 1)) Then 'shift right
            X3_VAL = X2_VAL
            X2_VAL = X1_VAL
            Y3_VAL = Y2_VAL
            Y2_VAL = Y1_VAL
            C_VAL = B_VAL
            B_VAL = A_VAL
        End If
        X1_VAL = XDATA_VECTOR(k, 1)
        Y1_VAL = YDATA_VECTOR(k, 1)
        A_VAL = k

    ElseIf XDATA_VECTOR(k, 1) < X2_VAL Then   'case  P1<Pmax<P2<P3
        If Sgn(TEMP_VECTOR(k, 1)) = Sgn(TEMP_VECTOR(B_VAL, 1)) Then
            X2_VAL = XDATA_VECTOR(k, 1)
            Y2_VAL = YDATA_VECTOR(k, 1)
            B_VAL = k
        Else
            X1_VAL = XDATA_VECTOR(k, 1)
            Y1_VAL = YDATA_VECTOR(k, 1)
            A_VAL = k
        End If

    ElseIf XDATA_VECTOR(k, 1) > X2_VAL Then   'case  P1<P2<Pmax<P3
        If Sgn(TEMP_VECTOR(k, 1)) = Sgn(TEMP_VECTOR(B_VAL, 1)) Then
            X2_VAL = XDATA_VECTOR(k, 1)
            Y2_VAL = YDATA_VECTOR(k, 1)
            B_VAL = k
        Else
            X3_VAL = XDATA_VECTOR(k, 1)
            Y3_VAL = YDATA_VECTOR(k, 1)
            C_VAL = k
        End If
    End If
    l = l - 1
Loop Until l = 0

If l > 0 Then
    REGRESSION_MIN_MAX_FUNC = Array(FIRST_SCALAR, NCOLUMNS)
Else
    GoTo ERROR_LABEL
End If

Exit Function
ERROR_LABEL:
REGRESSION_MIN_MAX_FUNC = Err.number
End Function
