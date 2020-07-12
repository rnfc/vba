Attribute VB_Name = "FINAN_PORT_WEIGHTS_FUTURES_LIBR"
'---------------------------------------------------------------------------
'---------------------------------------------------------------------------
Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.
'---------------------------------------------------------------------------
'---------------------------------------------------------------------------

Private Const PUB_EPSILON As Double = 2 ^ 52


'************************************************************************************
'************************************************************************************
'FUNCTION      : PORT_FUTURES_HEDGE_FUNC
'DESCRIPTION   : Optimal allocation for hedged/unhedged currencies returns
'LIBRARY       : FINAN_PORT_
'GROUP         : HEDGE_FUTURES
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/06/2011
'************************************************************************************
'************************************************************************************

Function PORT_FUTURES_HEDGE_FUNC(ByRef DATES_RNG As Variant, _
ByRef ASSET_PRICES_RNG As Variant, _
ByRef ASSET_TICKER_RNG As Variant, _
ByRef ASSET_CURRENCY_RNG As Variant, _
ByRef SPOT_FX_PRICES_RNG As Variant, _
ByRef SPOT_FX_CURRENCY_RNG As Variant, _
ByRef FORWARD_PRICES_RNG As Variant, _
ByRef FORWARD_CURRENCY_RNG As Variant, _
Optional ByVal OUTPUT As Integer = 4)
'Optional ByVal BASE_CURRENCY_STR As String = "USD", _

Dim i As Long
Dim j As Long
Dim k As Long

Dim NROWS As Long
Dim NCOLUMNS As Long

Dim TEMP_STR As String
Dim TEMP_VAL As Double

Dim TEMP1_SUM As Double
Dim TEMP2_SUM As Double

Dim SPOT_FX_COLLECTION_OBJ As Collection
Dim FORWARD_COLLECTION_OBJ As Collection

Dim TEMP0_MATRIX As Variant
Dim TEMP1_MATRIX As Variant
Dim TEMP2_MATRIX As Variant
Dim TEMP3_MATRIX As Variant
Dim TEMP4_MATRIX As Variant

Dim MEAN_VECTOR As Variant
Dim DATA_MATRIX As Variant
Dim COVARIANCE_MATRIX As Variant
Dim VARIANCE_VECTOR As Variant

Dim DATES_VECTOR As Variant
Dim ASSET_PRICES_MATRIX As Variant
Dim ASSET_TICKER_VECTOR As Variant
Dim ASSET_CURRENCY_VECTOR As Variant
Dim SPOT_FX_PRICES_MATRIX As Variant
Dim SPOT_FX_CURRENCY_VECTOR As Variant
Dim FORWARD_PRICES_MATRIX As Variant
Dim FORWARD_CURRENCY_VECTOR As Variant

'-----------------------------------------------------------------------------------------------------
On Error GoTo ERROR_LABEL
'-----------------------------------------------------------------------------------------------------

DATES_VECTOR = DATES_RNG
If UBound(DATES_VECTOR, 1) = 1 Then
    DATES_VECTOR = MATRIX_TRANSPOSE_FUNC(DATES_VECTOR)
End If
NROWS = UBound(DATES_VECTOR, 1)

ASSET_PRICES_MATRIX = ASSET_PRICES_RNG
If UBound(ASSET_PRICES_MATRIX, 1) = 1 Then
    ASSET_PRICES_MATRIX = MATRIX_TRANSPOSE_FUNC(ASSET_PRICES_MATRIX)
End If
If NROWS <> UBound(ASSET_PRICES_MATRIX, 1) Then: GoTo ERROR_LABEL
NCOLUMNS = UBound(ASSET_PRICES_MATRIX, 2)

ASSET_TICKER_VECTOR = ASSET_TICKER_RNG
If UBound(ASSET_TICKER_VECTOR, 2) = 1 Then
    ASSET_TICKER_VECTOR = MATRIX_TRANSPOSE_FUNC(ASSET_TICKER_VECTOR)
End If
If NCOLUMNS <> UBound(ASSET_TICKER_VECTOR, 2) Then: GoTo ERROR_LABEL

ASSET_CURRENCY_VECTOR = ASSET_CURRENCY_RNG
If UBound(ASSET_CURRENCY_VECTOR, 2) = 1 Then
    ASSET_CURRENCY_VECTOR = MATRIX_TRANSPOSE_FUNC(ASSET_CURRENCY_VECTOR)
End If
If NCOLUMNS <> UBound(ASSET_CURRENCY_VECTOR, 2) Then: GoTo ERROR_LABEL

SPOT_FX_PRICES_MATRIX = SPOT_FX_PRICES_RNG
If UBound(SPOT_FX_PRICES_MATRIX, 1) = 1 Then
    SPOT_FX_PRICES_MATRIX = MATRIX_TRANSPOSE_FUNC(SPOT_FX_PRICES_MATRIX)
End If
If NROWS <> UBound(SPOT_FX_PRICES_MATRIX, 1) Then: GoTo ERROR_LABEL
SPOT_FX_CURRENCY_VECTOR = SPOT_FX_CURRENCY_RNG
If UBound(SPOT_FX_CURRENCY_VECTOR, 2) = 1 Then
    SPOT_FX_CURRENCY_VECTOR = MATRIX_TRANSPOSE_FUNC(SPOT_FX_CURRENCY_VECTOR)
End If

FORWARD_PRICES_MATRIX = FORWARD_PRICES_RNG
If UBound(FORWARD_PRICES_MATRIX, 1) = 1 Then
    FORWARD_PRICES_MATRIX = MATRIX_TRANSPOSE_FUNC(FORWARD_PRICES_MATRIX)
End If
If NROWS <> UBound(FORWARD_PRICES_MATRIX, 1) Then: GoTo ERROR_LABEL
FORWARD_CURRENCY_VECTOR = FORWARD_CURRENCY_RNG
If UBound(FORWARD_CURRENCY_VECTOR, 2) = 1 Then
    FORWARD_CURRENCY_VECTOR = MATRIX_TRANSPOSE_FUNC(FORWARD_CURRENCY_VECTOR)
End If
If UBound(SPOT_FX_CURRENCY_VECTOR, 2) <> UBound(FORWARD_CURRENCY_VECTOR, 2) Then: GoTo ERROR_LABEL

GoSub INDEX_LINE
ReDim TEMP0_MATRIX(0 To NROWS - 1, 1 To NCOLUMNS + 1)
TEMP0_MATRIX(0, 1) = "DATES"
For j = 1 To NCOLUMNS
    TEMP0_MATRIX(0, j + 1) = ASSET_TICKER_VECTOR(1, j)
Next j
For i = 2 To NROWS
    TEMP0_MATRIX(i - 1, 1) = DATES_VECTOR(i, 1)
    For j = 1 To NCOLUMNS
        TEMP_VAL = ASSET_PRICES_MATRIX(i, j) / ASSET_PRICES_MATRIX(i - 1, j) - 1
        TEMP0_MATRIX(i - 1, j + 1) = TEMP_VAL
    Next j
Next i
'-----------------------------------------------------------------------------------------------------
If OUTPUT = 0 Then 'Asset Returns
'-----------------------------------------------------------------------------------------------------
    DATA_MATRIX = TEMP0_MATRIX: GoSub COVARIANCE_LINE: Erase DATA_MATRIX
    PORT_FUTURES_HEDGE_FUNC = Array(TEMP0_MATRIX, COVARIANCE_MATRIX, MEAN_VECTOR, VARIANCE_VECTOR)
    Exit Function
'-----------------------------------------------------------------------------------------------------
End If
'-----------------------------------------------------------------------------------------------------
ReDim TEMP1_MATRIX(0 To NROWS, 1 To NCOLUMNS + 1)
ReDim TEMP2_MATRIX(0 To NROWS, 1 To NCOLUMNS + 1)
TEMP1_MATRIX(0, 1) = "DATES"
TEMP2_MATRIX(0, 1) = "DATES"
For j = 1 To NCOLUMNS
    TEMP1_MATRIX(0, j + 1) = ASSET_CURRENCY_VECTOR(1, j) & ": " & ASSET_TICKER_VECTOR(1, j)
    TEMP2_MATRIX(0, j + 1) = ASSET_CURRENCY_VECTOR(1, j) & ": " & ASSET_TICKER_VECTOR(1, j)
Next j
For i = 1 To NROWS
    TEMP1_MATRIX(i, 1) = DATES_VECTOR(i, 1)
    TEMP2_MATRIX(i, 1) = DATES_VECTOR(i, 1)
    For j = 1 To NCOLUMNS
        TEMP_STR = ASSET_CURRENCY_VECTOR(1, j)
        k = 0: k = CLng(SPOT_FX_COLLECTION_OBJ(TEMP_STR))
        TEMP_VAL = ASSET_PRICES_MATRIX(i, j) / SPOT_FX_PRICES_MATRIX(i, k)
        TEMP1_MATRIX(i, j + 1) = TEMP_VAL
        k = 0: k = CLng(FORWARD_COLLECTION_OBJ(TEMP_STR))
        TEMP_VAL = ASSET_PRICES_MATRIX(i, j) / FORWARD_PRICES_MATRIX(i, k)
        TEMP2_MATRIX(i, j + 1) = TEMP_VAL
    Next j
Next i
'-----------------------------------------------------------------------------------------------------
If OUTPUT = 1 Then 'Unhedged Price (@spot)
'-----------------------------------------------------------------------------------------------------
    PORT_FUTURES_HEDGE_FUNC = TEMP1_MATRIX
    Exit Function
'-----------------------------------------------------------------------------------------------------
ElseIf OUTPUT = 2 Then 'Unhedged Price (@Forward)
'-----------------------------------------------------------------------------------------------------
    PORT_FUTURES_HEDGE_FUNC = TEMP2_MATRIX
    Exit Function
'-----------------------------------------------------------------------------------------------------
End If
'-----------------------------------------------------------------------------------------------------
ReDim TEMP3_MATRIX(0 To NROWS - 1, 1 To NCOLUMNS + 1)
ReDim TEMP4_MATRIX(0 To NROWS - 1, 1 To NCOLUMNS + 1)
TEMP3_MATRIX(0, 1) = "DATES"
TEMP4_MATRIX(0, 1) = "DATES"
For j = 1 To NCOLUMNS
    TEMP3_MATRIX(0, j + 1) = ASSET_CURRENCY_VECTOR(1, j) & ": " & ASSET_TICKER_VECTOR(1, j)
    TEMP4_MATRIX(0, j + 1) = ASSET_CURRENCY_VECTOR(1, j) & ": " & ASSET_TICKER_VECTOR(1, j)
Next j
For i = 2 To NROWS
    TEMP3_MATRIX(i - 1, 1) = DATES_VECTOR(i, 1)
    TEMP4_MATRIX(i - 1, 1) = DATES_VECTOR(i, 1)
    For j = 1 To NCOLUMNS
        TEMP_VAL = TEMP1_MATRIX(i, j + 1) / TEMP1_MATRIX(i - 1, j + 1) - 1
        TEMP3_MATRIX(i - 1, j + 1) = TEMP_VAL
        TEMP_STR = ASSET_CURRENCY_VECTOR(1, j)
        k = 0: k = CLng(SPOT_FX_COLLECTION_OBJ(TEMP_STR))
        TEMP_VAL = (TEMP1_MATRIX(i - 1, j + 1) - TEMP2_MATRIX(i - 1, j + 1) + ((ASSET_PRICES_MATRIX(i, j) - ASSET_PRICES_MATRIX(i - 1, j)) / SPOT_FX_PRICES_MATRIX(i, k))) / TEMP1_MATRIX(i - 1, j + 1)
        TEMP4_MATRIX(i - 1, j + 1) = TEMP_VAL
    Next j
Next i
'-----------------------------------------------------------------------------------------------------
If OUTPUT = 3 Then 'Unhedged Returns
'-----------------------------------------------------------------------------------------------------
    DATA_MATRIX = TEMP3_MATRIX: GoSub COVARIANCE_LINE: Erase DATA_MATRIX
    PORT_FUTURES_HEDGE_FUNC = Array(TEMP3_MATRIX, COVARIANCE_MATRIX, MEAN_VECTOR, VARIANCE_VECTOR)
'-----------------------------------------------------------------------------------------------------
ElseIf OUTPUT = 4 Then 'Hedged Returns
'-----------------------------------------------------------------------------------------------------
    DATA_MATRIX = TEMP4_MATRIX: GoSub COVARIANCE_LINE: Erase DATA_MATRIX
    PORT_FUTURES_HEDGE_FUNC = Array(TEMP4_MATRIX, COVARIANCE_MATRIX, MEAN_VECTOR, VARIANCE_VECTOR)
'-----------------------------------------------------------------------------------------------------
Else
'-----------------------------------------------------------------------------------------------------
    PORT_FUTURES_HEDGE_FUNC = Array(TEMP4_MATRIX, TEMP3_MATRIX, TEMP2_MATRIX, TEMP1_MATRIX, TEMP0_MATRIX)
'-----------------------------------------------------------------------------------------------------
End If
'-----------------------------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------------------------
Exit Function
'-----------------------------------------------------------------------------------------------------
INDEX_LINE:
'-----------------------------------------------------------------------------------------------------
    Set SPOT_FX_COLLECTION_OBJ = New Collection
    Set FORWARD_COLLECTION_OBJ = New Collection
    On Error Resume Next
    For j = LBound(SPOT_FX_CURRENCY_VECTOR, 2) To UBound(SPOT_FX_CURRENCY_VECTOR, 2)
        TEMP_STR = SPOT_FX_CURRENCY_VECTOR(1, j)
        Call SPOT_FX_COLLECTION_OBJ.Add(TEMP_STR, CStr(j))
        Call SPOT_FX_COLLECTION_OBJ.Add(CStr(j), TEMP_STR)
    
        TEMP_STR = FORWARD_CURRENCY_VECTOR(1, j)
        Call FORWARD_COLLECTION_OBJ.Add(TEMP_STR, CStr(j))
        Call FORWARD_COLLECTION_OBJ.Add(CStr(j), TEMP_STR)
    Next j
    Err.Clear
    On Error GoTo ERROR_LABEL
'-----------------------------------------------------------------------------------------------------
Return
'-----------------------------------------------------------------------------------------------------
COVARIANCE_LINE: 'Compute Covariance Matrix
'-----------------------------------------------------------------------------------------------
    ReDim MEAN_VECTOR(1 To NCOLUMNS, 1 To 1) 'compute means of data in matrix (calc with n-1)
    For j = 1 To NCOLUMNS
        TEMP1_SUM = 0
        For i = 1 To NROWS - 1
            TEMP1_SUM = TEMP1_SUM + DATA_MATRIX(i, j + 1)
        Next i
        MEAN_VECTOR(j, 1) = TEMP1_SUM / (NROWS - 1)
    Next j
    ReDim COVARIANCE_MATRIX(1 To NCOLUMNS, 1 To NCOLUMNS)
    ReDim VARIANCE_VECTOR(1 To NCOLUMNS, 1 To 1)
    For j = 1 To NCOLUMNS
        For k = 1 To j
            TEMP1_SUM = 0: TEMP2_SUM = 0
            For i = 1 To NROWS - 1
                TEMP1_SUM = TEMP1_SUM + (DATA_MATRIX(i, j + 1) - MEAN_VECTOR(j, 1)) * (DATA_MATRIX(i, k + 1) - MEAN_VECTOR(k, 1))
                If k = 1 Then: TEMP2_SUM = TEMP2_SUM + (DATA_MATRIX(i, j + 1) - MEAN_VECTOR(j, 1)) ^ 2
            Next i
            COVARIANCE_MATRIX(j, k) = TEMP1_SUM / (NROWS - 2)
            If k = 1 Then: VARIANCE_VECTOR(j, 1) = (TEMP2_SUM / (NROWS - 2)) '^ 0.5 'Stdev
        Next k
    Next j
    For j = 1 To NCOLUMNS: For k = j + 1 To NCOLUMNS: COVARIANCE_MATRIX(j, k) = COVARIANCE_MATRIX(k, j): Next k: Next j
'-------------------------------------------------------------------------------------------------
Return
'-------------------------------------------------------------------------------------------------
ERROR_LABEL:
PORT_FUTURES_HEDGE_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : PORT_FUTURES_TARGET_VAR_OBJ_FUNC
'DESCRIPTION   : Objective Function for Hedged/Unhedged Portfolio Optimizer
'LIBRARY       : FINAN_PORT_
'GROUP         : HEDGE_FUTURES
'ID            : 002
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/06/2011
'************************************************************************************
'************************************************************************************

Function PORT_FUTURES_TARGET_VAR_OBJ_FUNC(ByRef MEAN_RNG As Variant, _
ByRef WEIGHTS_RNG As Variant, _
ByRef COVAR_UNHEDGED_RNG As Variant, _
Optional ByRef COVAR_HEDGED_RNG As Variant, _
Optional ByVal TARGET_VAR_VAL As Double = 1.74643034618447E-03, _
Optional ByVal LONG_LIMIT_VAL As Double = 1, _
Optional ByVal SHORT_LIMIT_VAL As Double = 0.2)

'COVAR_UNHEDGED_RNG --> Use UnHedged Returns
'COVAR_HEDGED_RNG --> Hedged Returns --> Use Hedge Returns

Dim i As Long
Dim j As Long
Dim NCOLUMNS As Long

Dim OK_FLAG As Boolean

Dim TEMP0_SUM As Double
Dim TEMP1_SUM As Double
Dim TEMP2_SUM As Double
Dim TEMP3_SUM As Double 'Xn^2(sn^2)
Dim PORT_VAR_VAL As Double
Dim PORT_RETURN_VAL As Double

Dim MEAN_VECTOR As Variant
Dim WEIGHTS_VECTOR As Variant
Dim COVAR_UNHEDGED_MATRIX As Variant
Dim COVAR_HEDGED_MATRIX As Variant

Const tolerance1 As Double = 10 ^ -10 'Long Limit
Const tolerance2 As Double = 10 ^ -10 'Short Limit
Const tolerance3 As Double = 10 ^ -10 'Port Sum
Const tolerance4 As Double = 10 ^ -5 'Port Var

On Error GoTo ERROR_LABEL

MEAN_VECTOR = MEAN_RNG
If UBound(MEAN_VECTOR, 1) = 1 Then
    MEAN_VECTOR = MATRIX_TRANSPOSE_FUNC(MEAN_VECTOR)
End If
NCOLUMNS = UBound(MEAN_VECTOR, 1)

WEIGHTS_VECTOR = WEIGHTS_RNG
If UBound(WEIGHTS_VECTOR, 1) = 1 Then
    WEIGHTS_VECTOR = MATRIX_TRANSPOSE_FUNC(WEIGHTS_VECTOR)
End If
If NCOLUMNS <> UBound(WEIGHTS_VECTOR, 1) Then: GoTo ERROR_LABEL
COVAR_UNHEDGED_MATRIX = COVAR_UNHEDGED_RNG
If NCOLUMNS <> UBound(COVAR_UNHEDGED_MATRIX, 1) Then: GoTo ERROR_LABEL
If NCOLUMNS <> UBound(COVAR_UNHEDGED_MATRIX, 2) Then: GoTo ERROR_LABEL

If IsArray(COVAR_HEDGED_RNG) = False Then
    COVAR_HEDGED_MATRIX = COVAR_UNHEDGED_RNG
Else
    COVAR_HEDGED_MATRIX = COVAR_HEDGED_RNG
End If
If NCOLUMNS <> UBound(COVAR_HEDGED_MATRIX, 1) Then: GoTo ERROR_LABEL
If NCOLUMNS <> UBound(COVAR_HEDGED_MATRIX, 2) Then: GoTo ERROR_LABEL

TEMP0_SUM = 0: TEMP1_SUM = 0
TEMP2_SUM = 0: TEMP3_SUM = 0
For i = 1 To NCOLUMNS 'Matrix for (Xi)(Xj)(Corij)(si)(sj)
    TEMP0_SUM = TEMP0_SUM + WEIGHTS_VECTOR(i, 1)
    TEMP1_SUM = TEMP1_SUM + WEIGHTS_VECTOR(i, 1) * MEAN_VECTOR(i, 1)
    For j = i To NCOLUMNS
        If j <> i Then
            TEMP2_SUM = TEMP2_SUM + (WEIGHTS_VECTOR(i, 1) * WEIGHTS_VECTOR(j, 1)) * COVAR_HEDGED_MATRIX(i, j)
        Else
            TEMP3_SUM = TEMP3_SUM + (WEIGHTS_VECTOR(i, 1) * WEIGHTS_VECTOR(i, 1)) * COVAR_UNHEDGED_MATRIX(i, j) 'Xn^2(sn^2)
        End If
    Next j
Next i
PORT_RETURN_VAL = TEMP1_SUM
TEMP2_SUM = TEMP2_SUM * 2
PORT_VAR_VAL = TEMP2_SUM + TEMP3_SUM

'-------------------------------------------------------------------------------------------------------------------
'Directions
'-------------------------------------------------------------------------------------------------------------------
'Step 1: Set all the weights to zero
'Step 2: Decide on the target MONTHLY variance, and enter that value in the var "target variance".
'Step 3: Enter any desired short and long limits
'Step 4: Optimize.  The expected return will appear. By modifying the solver settings and inputing
'a value in the "target return" box, this Optimizer can also be used to determine a variance.
If (TEMP0_SUM <= LONG_LIMIT_VAL + tolerance1) And (TEMP0_SUM >= (-SHORT_LIMIT_VAL - tolerance2)) And (Abs(TEMP0_SUM - 1) <= tolerance3) And (Abs(PORT_VAR_VAL - TARGET_VAR_VAL) <= tolerance4) Then 'Weight/Var Check
    OK_FLAG = True
Else
    OK_FLAG = False
End If
'We could also modify the solver routine and inputing a value in the "target return", this
'Optimizer can also be used to determine a variance.

PORT_FUTURES_TARGET_VAR_OBJ_FUNC = Array(PORT_RETURN_VAL, PORT_VAR_VAL, TEMP0_SUM, OK_FLAG)

Exit Function
ERROR_LABEL:
PORT_FUTURES_TARGET_VAR_OBJ_FUNC = PUB_EPSILON
End Function
