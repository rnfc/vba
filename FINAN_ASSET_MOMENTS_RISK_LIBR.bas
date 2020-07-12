Attribute VB_Name = "FINAN_ASSET_MOMENTS_RISK_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.

'************************************************************************************
'************************************************************************************
'FUNCTION      : ASSET_RISK_ADJUSTED_FUNC

'DESCRIPTION   : New Risk-Adjusted Performance Measures: Calculations for Calmar
'Ratio, Sterling Ratio, Burke Ratio, Excess Return on VaR, Modified Sharpe
'Ratio, Conditional Sharpe Ratio, Gain-Loss-Ratio, Sortino Ratio,
'Kappa, Omega and Upside-Potential-Ratio.

'LIBRARY       : FINAN_ASSET
'GROUP         : RISK_ADJ
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'************************************************************************************
'************************************************************************************

Function ASSET_RISK_ADJUSTED_FUNC(ByRef DATA_RNG As Variant, _
Optional ByVal CASH_RATE As Double = 0, _
Optional ByVal TRETURN_VAL As Double = 0, _
Optional ByVal NDRAWDOWNS_VAL As Integer = 5, _
Optional ByVal CONFIDENCE_VAL As Variant = 0.05, _
Optional ByVal DATA_TYPE As Integer = 0, _
Optional ByVal LOG_SCALE As Integer = 0, _
Optional ByVal PERIODS As Integer = 12, _
Optional ByVal COUNT_BASIS As Integer = 12)

'DATA_RNG: ...returns should be "excess returns" over a "riskfree rate"

Dim i As Long
Dim j As Long
Dim k As Long

Dim NROWS As Long
Dim NCOLUMNS As Long

Dim HEADINGS_STR As String

Dim MEAN_VAL As Double
Dim SIGMA_VAL As Double

Dim TEMP_VECTOR As Variant
Dim DATA_VECTOR As Variant
Dim DATA_MATRIX As Variant
Dim TEMP_MATRIX As Variant

On Error GoTo ERROR_LABEL

DATA_MATRIX = DATA_RNG
If DATA_TYPE <> 0 Then: DATA_MATRIX = MATRIX_PERCENT_FUNC(DATA_MATRIX, LOG_SCALE)
NROWS = UBound(DATA_MATRIX, 1)
NCOLUMNS = UBound(DATA_MATRIX, 2)
For j = 1 To NCOLUMNS: For i = 1 To NROWS: DATA_MATRIX(i, j) = DATA_MATRIX(i, j) - CASH_RATE: Next i: Next j

HEADINGS_STR = "ID,NORMAL VAR,MODIFIED VAR,CONDITIONAL VAR,MAXIMUM DRAWDOWN,AVERAGE DRAWDOWN,SQUARE ROOT SQUARED DRAWDOWNS,LPM(1),HPM(1),LPM(2),LPM(3),SHARPE,CALMAR,STERLING,BURKE,EXCESS RETURN ON VAR,CONDITIONAL SHARPE RATIO,MODIFIED SHARPE RATIO,GAIN LOSS,SORTINO,KAPPA,OMEGA,UPSIDE POTENTIAL,"
ReDim TEMP_MATRIX(0 To NCOLUMNS, 1 To 23)
i = 1
For k = 1 To 23
    j = InStr(i, HEADINGS_STR, ",")
    TEMP_MATRIX(0, k) = Mid(HEADINGS_STR, i, j - i)
    i = j + 1
Next k

On Error GoTo 1983

For i = 1 To NCOLUMNS
    TEMP_MATRIX(i, 1) = i
    
    DATA_VECTOR = MATRIX_GET_COLUMN_FUNC(DATA_MATRIX, i, 1)
    TEMP_VECTOR = ASSET_VARS_FUNC(DATA_VECTOR, CONFIDENCE_VAL, 0, 0, PERIODS, COUNT_BASIS)
    If IsArray(TEMP_VECTOR) = False Then: GoTo 1983
    
    MEAN_VAL = TEMP_VECTOR(LBound(TEMP_VECTOR) + 0)
    SIGMA_VAL = TEMP_VECTOR(LBound(TEMP_VECTOR) + 1)
    TEMP_MATRIX(i, 2) = TEMP_VECTOR(LBound(TEMP_VECTOR) + 7)
    TEMP_MATRIX(i, 3) = TEMP_VECTOR(LBound(TEMP_VECTOR) + 8)
    TEMP_MATRIX(i, 4) = TEMP_VECTOR(LBound(TEMP_VECTOR) + 9)

    TEMP_VECTOR = ASSET_DD_DISCRETE_FUNC(DATA_VECTOR, 0)
    If IsArray(TEMP_VECTOR) = False Then: GoTo 1983
    TEMP_MATRIX(i, 5) = TEMP_VECTOR(1, 1)
    TEMP_MATRIX(i, 6) = 0: TEMP_MATRIX(i, 7) = 0
    For j = 1 To NDRAWDOWNS_VAL
        TEMP_MATRIX(i, 6) = TEMP_MATRIX(i, 6) + TEMP_VECTOR(j, 1)
        TEMP_MATRIX(i, 7) = TEMP_MATRIX(i, 7) + TEMP_VECTOR(j, 1) ^ 2
    Next j
    
    TEMP_MATRIX(i, 6) = TEMP_MATRIX(i, 6) / NDRAWDOWNS_VAL
    TEMP_MATRIX(i, 7) = TEMP_MATRIX(i, 7) ^ 0.5
    TEMP_MATRIX(i, 8) = ASSET_LPM_FUNC(DATA_VECTOR, TRETURN_VAL, 1, 0, 0)
    TEMP_MATRIX(i, 9) = ASSET_HPM_FUNC(DATA_VECTOR, TRETURN_VAL, 1, 0, 0)
    TEMP_MATRIX(i, 10) = ASSET_LPM_FUNC(DATA_VECTOR, TRETURN_VAL, 2, 0, 0)
    TEMP_MATRIX(i, 11) = ASSET_LPM_FUNC(DATA_VECTOR, TRETURN_VAL, 3, 0, 0)
    
    If SIGMA_VAL <> 0 Then
        TEMP_MATRIX(i, 12) = MEAN_VAL / SIGMA_VAL
    Else
        TEMP_MATRIX(i, 12) = CVErr(xlErrNA)
    End If
    
    If TEMP_MATRIX(i, 5) <> 0 Then
        TEMP_MATRIX(i, 13) = MEAN_VAL / (TEMP_MATRIX(i, 5) * -1)
    Else
        TEMP_MATRIX(i, 13) = CVErr(xlErrNA)
    End If
    
    If TEMP_MATRIX(i, 6) <> 0 Then
        TEMP_MATRIX(i, 14) = MEAN_VAL / (TEMP_MATRIX(i, 6) * -1)
    Else
        TEMP_MATRIX(i, 14) = CVErr(xlErrNA)
    End If
    
    If TEMP_MATRIX(i, 7) <> 0 Then
        TEMP_MATRIX(i, 15) = MEAN_VAL / TEMP_MATRIX(i, 7)
    Else
        TEMP_MATRIX(i, 15) = CVErr(xlErrNA)
    End If
    
    If TEMP_MATRIX(i, 2) <> 0 Then
        TEMP_MATRIX(i, 16) = MEAN_VAL / (TEMP_MATRIX(i, 2) * -1)
    Else
        TEMP_MATRIX(i, 16) = CVErr(xlErrNA)
    End If
    
    If TEMP_MATRIX(i, 4) <> 0 Then
        TEMP_MATRIX(i, 17) = MEAN_VAL / (TEMP_MATRIX(i, 4) * -1)
    Else
        TEMP_MATRIX(i, 17) = CVErr(xlErrNA)
    End If
    
    If TEMP_MATRIX(i, 3) <> 0 Then
        TEMP_MATRIX(i, 18) = MEAN_VAL / (TEMP_MATRIX(i, 3) * -1)
    Else
        TEMP_MATRIX(i, 18) = CVErr(xlErrNA)
    End If
    
    If TEMP_MATRIX(i, 8) <> 0 Then
        TEMP_MATRIX(i, 19) = TEMP_MATRIX(i, 9) / TEMP_MATRIX(i, 8)
        TEMP_MATRIX(i, 22) = 1 + (MEAN_VAL - TRETURN_VAL) / TEMP_MATRIX(i, 8)
    Else
        TEMP_MATRIX(i, 19) = CVErr(xlErrNA)
        TEMP_MATRIX(i, 22) = CVErr(xlErrNA)
    End If
    
    If TEMP_MATRIX(i, 10) <> 0 Then
        TEMP_MATRIX(i, 20) = (MEAN_VAL - TRETURN_VAL) / (TEMP_MATRIX(i, 10) ^ (1 / 2))
        TEMP_MATRIX(i, 23) = TEMP_MATRIX(i, 9) / (TEMP_MATRIX(i, 10) ^ (1 / 2))
    Else
        TEMP_MATRIX(i, 20) = CVErr(xlErrNA)
        TEMP_MATRIX(i, 23) = CVErr(xlErrNA)
    End If
    
    If TEMP_MATRIX(i, 11) <> 0 Then
        TEMP_MATRIX(i, 21) = (MEAN_VAL - TRETURN_VAL) / (TEMP_MATRIX(i, 11) ^ (1 / 3))
    Else
        TEMP_MATRIX(i, 21) = CVErr(xlErrNA)
    End If
    
1983:
Next i

ASSET_RISK_ADJUSTED_FUNC = TEMP_MATRIX

Exit Function
ERROR_LABEL:
ASSET_RISK_ADJUSTED_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : ASSET_GMS_FUNC
'DESCRIPTION   : Gms function
'LIBRARY       : FINAN_ASSET
'GROUP         : RISK_ADJ
'ID            : 002
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'************************************************************************************
'************************************************************************************

Function ASSET_GMS_FUNC(ByRef DATA_RNG As Variant, _
Optional ByVal DATA_TYPE As Integer = 0)

Dim i As Long
Dim NROWS As Long

Dim MULT_VAL As Double
Dim TEMP_SUM As Double
Dim DATA_VECTOR As Variant

On Error GoTo ERROR_LABEL

DATA_VECTOR = DATA_RNG
If DATA_TYPE <> 0 Then: DATA_VECTOR = MATRIX_PERCENT_FUNC(DATA_VECTOR, 0)

NROWS = UBound(DATA_VECTOR)
MULT_VAL = 1
TEMP_SUM = 0
For i = 1 To NROWS
    MULT_VAL = MULT_VAL * (1 + DATA_VECTOR(i, 1))
    TEMP_SUM = TEMP_SUM + (1 / MULT_VAL)
Next i
ASSET_GMS_FUNC = TEMP_SUM

Exit Function
ERROR_LABEL:
ASSET_GMS_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : ASSET_LPM_FUNC
'DESCRIPTION   : Calculates lower partial moment of order for a vector of returns
'LIBRARY       : FINAN_ASSET
'GROUP         : RISK_ADJ
'ID            : 003
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'************************************************************************************
'************************************************************************************

Function ASSET_LPM_FUNC(ByRef DATA_RNG As Variant, _
Optional ByVal TRETURN_VAL As Double = 0, _
Optional ByVal EFACTOR_VAL As Integer = 1, _
Optional ByVal DATA_TYPE As Integer = 0, _
Optional ByVal LOG_SCALE As Integer = 0)

Dim i As Long
Dim NROWS As Long
Dim TEMP_VAL As Double
Dim TEMP_SUM As Double
Dim DATA_VECTOR As Variant

On Error GoTo ERROR_LABEL

DATA_VECTOR = DATA_RNG
If UBound(DATA_VECTOR, 1) = 1 Then: DATA_VECTOR = MATRIX_TRANSPOSE_FUNC(DATA_VECTOR)
If DATA_TYPE <> 0 Then: DATA_VECTOR = MATRIX_PERCENT_FUNC(DATA_VECTOR, LOG_SCALE)

NROWS = UBound(DATA_VECTOR, 1)
TEMP_SUM = 0
For i = 1 To NROWS
    TEMP_VAL = TRETURN_VAL - DATA_VECTOR(i, 1)
    If TEMP_VAL > 0 Then: TEMP_SUM = TEMP_SUM + TEMP_VAL ^ EFACTOR_VAL
Next i

ASSET_LPM_FUNC = TEMP_SUM / NROWS
    
Exit Function
ERROR_LABEL:
ASSET_LPM_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : ASSET_HPM_FUNC
'DESCRIPTION   : Calculates higher partial moment of order for a vector of returns
'LIBRARY       : FINAN_ASSET
'GROUP         : RISK_ADJ
'ID            : 004
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'************************************************************************************
'************************************************************************************

Function ASSET_HPM_FUNC(ByRef DATA_RNG As Variant, _
Optional ByVal TRETURN_VAL As Double = 0, _
Optional ByVal EFACTOR_VAL As Integer = 1, _
Optional ByVal DATA_TYPE As Integer = 0, _
Optional ByVal LOG_SCALE As Integer = 0)

Dim i As Long
Dim NROWS As Long

Dim TEMP_VAL As Double
Dim TEMP_SUM As Double
Dim DATA_VECTOR As Variant

On Error GoTo ERROR_LABEL

DATA_VECTOR = DATA_RNG
If UBound(DATA_VECTOR, 1) = 1 Then: DATA_VECTOR = MATRIX_TRANSPOSE_FUNC(DATA_VECTOR)
If DATA_TYPE <> 0 Then: DATA_VECTOR = MATRIX_PERCENT_FUNC(DATA_VECTOR, LOG_SCALE)

NROWS = UBound(DATA_VECTOR, 1)
TEMP_SUM = 0
For i = 1 To NROWS
    TEMP_VAL = DATA_VECTOR(i, 1) - TRETURN_VAL
    If TEMP_VAL > 0 Then: TEMP_SUM = TEMP_SUM + TEMP_VAL ^ EFACTOR_VAL
Next i

ASSET_HPM_FUNC = TEMP_SUM / NROWS
    
Exit Function
ERROR_LABEL:
ASSET_HPM_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : ASSET_CALMAR_FUNC
'DESCRIPTION   : Calmar Ratio
'LIBRARY       : FINAN_ASSET
'GROUP         : RISK_ADJ
'ID            : 005
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA

'************************************************************************************
'************************************************************************************

Function ASSET_CALMAR_FUNC(ByRef DATA_RNG As Variant, _
Optional ByVal DATA_TYPE As Integer = 0, _
Optional ByVal LOG_SCALE As Integer = 0)

Dim i As Long
Dim MIN_VAL As Double
Dim MEAN_VAL As Double
Dim DATA_VECTOR As Variant

On Error GoTo ERROR_LABEL
       
DATA_VECTOR = DATA_RNG
If UBound(DATA_VECTOR, 1) = 1 Then: DATA_VECTOR = MATRIX_TRANSPOSE_FUNC(DATA_VECTOR)
If DATA_TYPE <> 0 Then: DATA_VECTOR = MATRIX_PERCENT_FUNC(DATA_VECTOR, LOG_SCALE)

MEAN_VAL = MATRIX_MEAN_FUNC(DATA_VECTOR)(1, 1)
DATA_VECTOR = ASSET_DD_DISCRETE_FUNC(DATA_VECTOR, 0)
MIN_VAL = 2 ^ 52
For i = LBound(DATA_VECTOR, 1) To UBound(DATA_VECTOR, 1)
    If DATA_VECTOR(i, 1) < MIN_VAL Then: MIN_VAL = DATA_VECTOR(i, 1)
Next i

If MIN_VAL <> 0 Then
    ASSET_CALMAR_FUNC = MEAN_VAL / (MIN_VAL * -1)
Else
    ASSET_CALMAR_FUNC = CVErr(xlErrNA)
End If

Exit Function
ERROR_LABEL:
ASSET_CALMAR_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : ASSET_STERLING_FUNC
'DESCRIPTION   : Sterling Ratio
'LIBRARY       : FINAN_ASSET
'GROUP         : RISK_ADJ
'ID            : 006
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA

'************************************************************************************
'************************************************************************************

Function ASSET_STERLING_FUNC(ByRef DATA_RNG As Variant, _
Optional ByVal NDRAWDOWNS_VAL As Integer = 5, _
Optional ByVal DATA_TYPE As Integer = 0, _
Optional ByVal LOG_SCALE As Integer = 0)

Dim i As Long
Dim MEAN_VAL As Double
Dim TEMP_SUM As Double
Dim DATA_VECTOR As Variant

On Error GoTo ERROR_LABEL

DATA_VECTOR = DATA_RNG
If UBound(DATA_VECTOR, 1) = 1 Then: DATA_VECTOR = MATRIX_TRANSPOSE_FUNC(DATA_VECTOR)
If DATA_TYPE <> 0 Then: DATA_VECTOR = MATRIX_PERCENT_FUNC(DATA_VECTOR, LOG_SCALE)
MEAN_VAL = MATRIX_MEAN_FUNC(DATA_VECTOR)(1, 1)
' Number of drawdowns to be considered
DATA_VECTOR = ASSET_DD_DISCRETE_FUNC(DATA_VECTOR, 0)
TEMP_SUM = 0
For i = 1 To NDRAWDOWNS_VAL
    TEMP_SUM = TEMP_SUM + DATA_VECTOR(i, 1)
Next i

If TEMP_SUM <> 0 Then
    ASSET_STERLING_FUNC = MEAN_VAL / ((TEMP_SUM / NDRAWDOWNS_VAL) * -1)
Else
    ASSET_STERLING_FUNC = CVErr(xlErrNA)
End If

Exit Function
ERROR_LABEL:
ASSET_STERLING_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : ASSET_BURKE_FUNC
'DESCRIPTION   : Burke Ratio
'LIBRARY       : FINAN_ASSET
'GROUP         : RISK_ADJ
'ID            : 007
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA

'************************************************************************************
'************************************************************************************

Function ASSET_BURKE_FUNC(ByRef DATA_RNG As Variant, _
Optional ByVal NDRAWDOWNS_VAL As Integer = 5, _
Optional ByVal DATA_TYPE As Integer = 0, _
Optional ByVal LOG_SCALE As Integer = 0)

Dim i As Long
Dim MEAN_VAL As Double
Dim TEMP_SUM As Double

Dim DATA_VECTOR As Variant

On Error GoTo ERROR_LABEL

DATA_VECTOR = DATA_RNG
If UBound(DATA_VECTOR, 1) = 1 Then: DATA_VECTOR = MATRIX_TRANSPOSE_FUNC(DATA_VECTOR)
If DATA_TYPE <> 0 Then: DATA_VECTOR = MATRIX_PERCENT_FUNC(DATA_VECTOR, LOG_SCALE)

' Number of drawdowns to be considered
MEAN_VAL = MATRIX_MEAN_FUNC(DATA_VECTOR)(1, 1)
DATA_VECTOR = ASSET_DD_DISCRETE_FUNC(DATA_VECTOR, 0)
TEMP_SUM = 0
For i = 1 To NDRAWDOWNS_VAL
    TEMP_SUM = TEMP_SUM + DATA_VECTOR(i, 1) ^ 2
Next i
If TEMP_SUM <> 0 Then
    ASSET_BURKE_FUNC = MEAN_VAL / TEMP_SUM ^ (1 / 2)
Else
    ASSET_BURKE_FUNC = CVErr(xlErrNA)
End If

Exit Function
ERROR_LABEL:
ASSET_BURKE_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : ASSET_SHARPE_FUNC
'DESCRIPTION   : Sharpe Ratio
'LIBRARY       : FINAN_ASSET
'GROUP         : RISK_ADJ
'ID            : 008
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'************************************************************************************
'************************************************************************************

Function ASSET_SHARPE_FUNC(ByRef DATA_RNG As Variant, _
Optional ByVal DATA_TYPE As Integer = 0, _
Optional ByVal LOG_SCALE As Integer = 0)

Dim MEAN_VAL As Double
Dim STDEV_VAL As Double

Dim DATA_VECTOR As Variant

On Error GoTo ERROR_LABEL

DATA_VECTOR = DATA_RNG
If UBound(DATA_VECTOR, 1) = 1 Then: DATA_VECTOR = MATRIX_TRANSPOSE_FUNC(DATA_VECTOR)
If DATA_TYPE <> 0 Then: DATA_VECTOR = MATRIX_PERCENT_FUNC(DATA_VECTOR, LOG_SCALE)
MEAN_VAL = MATRIX_MEAN_FUNC(DATA_VECTOR)(1, 1)
STDEV_VAL = MATRIX_STDEV_FUNC(DATA_VECTOR)(1, 1)

If STDEV_VAL <> 0 Then
    ASSET_SHARPE_FUNC = MEAN_VAL / STDEV_VAL
Else
    ASSET_SHARPE_FUNC = CVErr(xlErrNA)
End If

Exit Function
ERROR_LABEL:
ASSET_SHARPE_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : ASSET_EXCESS_VAR_FUNC
'DESCRIPTION   : Excess Return On VaR
'LIBRARY       : FINAN_ASSET
'GROUP         : RISK_ADJ
'ID            : 009
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'************************************************************************************
'************************************************************************************

Function ASSET_EXCESS_VAR_FUNC(ByRef DATA_RNG As Variant, _
Optional ByVal CONFIDENCE_VAL As Variant = 0.05, _
Optional ByVal DATA_TYPE As Integer = 0, _
Optional ByVal LOG_SCALE As Integer = 0, _
Optional ByVal PERIODS As Integer = 12, _
Optional ByVal COUNT_BASIS As Integer = 12)

Dim NVAR_VAL As Double
Dim MEAN_VAL As Double

Dim DATA_VECTOR As Variant
Dim TEMP_VECTOR As Variant

On Error GoTo ERROR_LABEL
      
DATA_VECTOR = DATA_RNG
If UBound(DATA_VECTOR, 1) = 1 Then: DATA_VECTOR = MATRIX_TRANSPOSE_FUNC(DATA_VECTOR)
If DATA_TYPE <> 0 Then: DATA_VECTOR = MATRIX_PERCENT_FUNC(DATA_VECTOR, LOG_SCALE)
TEMP_VECTOR = ASSET_VARS_FUNC(DATA_VECTOR, CONFIDENCE_VAL, 0, 0, PERIODS, COUNT_BASIS)
MEAN_VAL = TEMP_VECTOR(LBound(TEMP_VECTOR) + 0)
NVAR_VAL = TEMP_VECTOR(LBound(TEMP_VECTOR) + 7)

If NVAR_VAL <> 0 Then
    ASSET_EXCESS_VAR_FUNC = MEAN_VAL / (NVAR_VAL * -1)
Else
    ASSET_EXCESS_VAR_FUNC = CVErr(xlErrNA)
End If

Exit Function
ERROR_LABEL:
ASSET_EXCESS_VAR_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : ASSET_CONDITIONAL_SHARPE_FUNC
'DESCRIPTION   : Conditional Sharpe Ratio
'LIBRARY       : FINAN_ASSET
'GROUP         : RISK_ADJ
'ID            : 010
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'************************************************************************************
'************************************************************************************

Function ASSET_CONDITIONAL_SHARPE_FUNC(ByRef DATA_RNG As Variant, _
Optional ByVal CONFIDENCE_VAL As Variant = 0.05, _
Optional ByVal DATA_TYPE As Integer = 0, _
Optional ByVal LOG_SCALE As Integer = 0, _
Optional ByVal PERIODS As Integer = 12, _
Optional ByVal COUNT_BASIS As Integer = 12)

Dim CVAR_VAL As Double
Dim MEAN_VAL As Double

Dim DATA_VECTOR As Variant
Dim TEMP_VECTOR As Variant

On Error GoTo ERROR_LABEL
      
DATA_VECTOR = DATA_RNG
If UBound(DATA_VECTOR, 1) = 1 Then: DATA_VECTOR = MATRIX_TRANSPOSE_FUNC(DATA_VECTOR)
If DATA_TYPE <> 0 Then: DATA_VECTOR = MATRIX_PERCENT_FUNC(DATA_VECTOR, LOG_SCALE)
TEMP_VECTOR = ASSET_VARS_FUNC(DATA_VECTOR, CONFIDENCE_VAL, 0, 0, PERIODS, COUNT_BASIS)
MEAN_VAL = TEMP_VECTOR(LBound(TEMP_VECTOR) + 0)
CVAR_VAL = TEMP_VECTOR(LBound(TEMP_VECTOR) + 9)

If CVAR_VAL <> 0 Then
    ASSET_CONDITIONAL_SHARPE_FUNC = MEAN_VAL / (CVAR_VAL * -1)
Else
    ASSET_CONDITIONAL_SHARPE_FUNC = CVErr(xlErrNA)
End If

Exit Function
ERROR_LABEL:
ASSET_CONDITIONAL_SHARPE_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : ASSET_MODIFIED_SHARPE_FUNC
'DESCRIPTION   : Modified Sharpe Ratio
'LIBRARY       : FINAN_ASSET
'GROUP         : RISK_ADJ
'ID            : 011
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'************************************************************************************
'************************************************************************************

Function ASSET_MODIFIED_SHARPE_FUNC(ByRef DATA_RNG As Variant, _
Optional ByVal CONFIDENCE_VAL As Variant = 0.05, _
Optional ByVal DATA_TYPE As Integer = 0, _
Optional ByVal LOG_SCALE As Integer = 0, _
Optional ByVal PERIODS As Integer = 12, _
Optional ByVal COUNT_BASIS As Integer = 12)

Dim MVAR_VAL As Double
Dim MEAN_VAL As Double

Dim DATA_VECTOR As Variant
Dim TEMP_VECTOR As Variant

On Error GoTo ERROR_LABEL
     
DATA_VECTOR = DATA_RNG
If UBound(DATA_VECTOR, 1) = 1 Then: DATA_VECTOR = MATRIX_TRANSPOSE_FUNC(DATA_VECTOR)
If DATA_TYPE <> 0 Then: DATA_VECTOR = MATRIX_PERCENT_FUNC(DATA_VECTOR, LOG_SCALE)
TEMP_VECTOR = ASSET_VARS_FUNC(DATA_VECTOR, CONFIDENCE_VAL, 0, 0, PERIODS, COUNT_BASIS)
MEAN_VAL = TEMP_VECTOR(LBound(TEMP_VECTOR) + 0)
MVAR_VAL = TEMP_VECTOR(LBound(TEMP_VECTOR) + 8)

If MVAR_VAL <> 0 Then
    ASSET_MODIFIED_SHARPE_FUNC = MEAN_VAL / (MVAR_VAL * -1)
Else
    ASSET_MODIFIED_SHARPE_FUNC = CVErr(xlErrNA)
End If

Exit Function
ERROR_LABEL:
ASSET_MODIFIED_SHARPE_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : ASSET_GAIN_LOSS_FUNC
'DESCRIPTION   : Gain-Loss-Ratio (assuming a minimum required excess return of zero)
'LIBRARY       : FINAN_ASSET
'GROUP         : RISK_ADJ
'ID            : 012
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'************************************************************************************
'************************************************************************************

Function ASSET_GAIN_LOSS_FUNC(ByRef DATA_RNG As Variant, _
Optional ByVal TRETURN_VAL As Double = 0#, _
Optional ByVal DATA_TYPE As Integer = 0, _
Optional ByVal LOG_SCALE As Integer = 0)

Dim LPM_VAL As Double
Dim HPM_VAL As Double

Dim DATA_VECTOR As Variant

On Error GoTo ERROR_LABEL

DATA_VECTOR = DATA_RNG
If UBound(DATA_VECTOR, 1) = 1 Then: DATA_VECTOR = MATRIX_TRANSPOSE_FUNC(DATA_VECTOR)
If DATA_TYPE <> 0 Then: DATA_VECTOR = MATRIX_PERCENT_FUNC(DATA_VECTOR, LOG_SCALE)

LPM_VAL = ASSET_LPM_FUNC(DATA_VECTOR, TRETURN_VAL, 1)
HPM_VAL = ASSET_HPM_FUNC(DATA_VECTOR, TRETURN_VAL, 1)

If LPM_VAL <> 0 Then
    ASSET_GAIN_LOSS_FUNC = HPM_VAL / LPM_VAL
Else
    ASSET_GAIN_LOSS_FUNC = CVErr(xlErrNA)
End If

Exit Function
ERROR_LABEL:
ASSET_GAIN_LOSS_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : ASSET_SORTINO_FUNC
'DESCRIPTION   : Sortino Ratio
'LIBRARY       : FINAN_ASSET
'GROUP         : RISK_ADJ
'ID            : 013
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA

'************************************************************************************
'************************************************************************************

Function ASSET_SORTINO_FUNC(ByRef DATA_RNG As Variant, _
Optional ByVal TRETURN_VAL As Double = 0#, _
Optional ByVal DATA_TYPE As Integer = 0, _
Optional ByVal LOG_SCALE As Integer = 0)

Dim LPM_VAL As Double
Dim MEAN_VAL As Double
Dim DATA_VECTOR As Variant

On Error GoTo ERROR_LABEL
      
DATA_VECTOR = DATA_RNG
If UBound(DATA_VECTOR, 1) = 1 Then: DATA_VECTOR = MATRIX_TRANSPOSE_FUNC(DATA_VECTOR)
If DATA_TYPE <> 0 Then: DATA_VECTOR = MATRIX_PERCENT_FUNC(DATA_VECTOR, LOG_SCALE)
MEAN_VAL = MATRIX_MEAN_FUNC(DATA_VECTOR)(1, 1)
LPM_VAL = ASSET_LPM_FUNC(DATA_VECTOR, TRETURN_VAL, 2)

If LPM_VAL <> 0 Then
    ASSET_SORTINO_FUNC = (MEAN_VAL - TRETURN_VAL) / LPM_VAL ^ (1 / 2)
Else
    ASSET_SORTINO_FUNC = CVErr(xlErrNA)
End If

Exit Function
ERROR_LABEL:
ASSET_SORTINO_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : ASSET_KAPPA_FUNC
'DESCRIPTION   : Kappa
'LIBRARY       : FINAN_ASSET
'GROUP         : RISK_ADJ
'ID            : 014



'AUTHOR        : RAFAEL NICOLAS FERMIN COTA

'************************************************************************************
'************************************************************************************

Function ASSET_KAPPA_FUNC(ByRef DATA_RNG As Variant, _
Optional ByVal TRETURN_VAL As Double = 0#, _
Optional ByVal DATA_TYPE As Integer = 0, _
Optional ByVal LOG_SCALE As Integer = 0)

Dim LPM_VAL As Double
Dim MEAN_VAL As Double
Dim DATA_VECTOR As Variant

On Error GoTo ERROR_LABEL
      
DATA_VECTOR = DATA_RNG
If UBound(DATA_VECTOR, 1) = 1 Then: DATA_VECTOR = MATRIX_TRANSPOSE_FUNC(DATA_VECTOR)
If DATA_TYPE <> 0 Then: DATA_VECTOR = MATRIX_PERCENT_FUNC(DATA_VECTOR, LOG_SCALE)
MEAN_VAL = MATRIX_MEAN_FUNC(DATA_VECTOR)(1, 1)
LPM_VAL = ASSET_LPM_FUNC(DATA_VECTOR, TRETURN_VAL, 3)

If LPM_VAL <> 0 Then
    ASSET_KAPPA_FUNC = (MEAN_VAL - TRETURN_VAL) / LPM_VAL ^ (1 / 3)
Else
    ASSET_KAPPA_FUNC = CVErr(xlErrNA)
End If

Exit Function
ERROR_LABEL:
ASSET_KAPPA_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : ASSET_OMEGA_FUNC
'DESCRIPTION   : Omega
'LIBRARY       : FINAN_ASSET
'GROUP         : RISK_ADJ
'ID            : 015
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'************************************************************************************
'************************************************************************************

Function ASSET_OMEGA_FUNC(ByRef DATA_RNG As Variant, _
Optional ByVal TRETURN_VAL As Double = 0#, _
Optional ByVal DATA_TYPE As Integer = 0, _
Optional ByVal LOG_SCALE As Integer = 0)

Dim MEAN_VAL As Double
Dim LPM_VAL As Double
Dim DATA_VECTOR As Variant

On Error GoTo ERROR_LABEL
     
DATA_VECTOR = DATA_RNG
If UBound(DATA_VECTOR, 1) = 1 Then: DATA_VECTOR = MATRIX_TRANSPOSE_FUNC(DATA_VECTOR)
If DATA_TYPE <> 0 Then: DATA_VECTOR = MATRIX_PERCENT_FUNC(DATA_VECTOR, LOG_SCALE)
MEAN_VAL = MATRIX_MEAN_FUNC(DATA_VECTOR)(1, 1)
LPM_VAL = ASSET_LPM_FUNC(DATA_VECTOR, TRETURN_VAL, 1)

If LPM_VAL <> 0 Then
    ASSET_OMEGA_FUNC = 1 + (MEAN_VAL - TRETURN_VAL) / LPM_VAL
Else
    ASSET_OMEGA_FUNC = CVErr(xlErrNA)
End If

Exit Function
ERROR_LABEL:
ASSET_OMEGA_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : ASSET_UP_POTENTIAL_FUNC
'DESCRIPTION   : Upside-Potential-Ratio
'LIBRARY       : FINAN_ASSET
'GROUP         : RISK_ADJ
'ID            : 016
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'************************************************************************************
'************************************************************************************

Function ASSET_UP_POTENTIAL_FUNC(ByRef DATA_RNG As Variant, _
Optional ByVal TRETURN_VAL As Double = 0#, _
Optional ByVal DATA_TYPE As Integer = 0, _
Optional ByVal LOG_SCALE As Integer = 0)

Dim LPM_VAL As Double
Dim HPM_VAL As Double

Dim DATA_VECTOR As Variant
On Error GoTo ERROR_LABEL
DATA_VECTOR = DATA_RNG
If UBound(DATA_VECTOR, 1) = 1 Then: DATA_VECTOR = MATRIX_TRANSPOSE_FUNC(DATA_VECTOR)
If DATA_TYPE <> 0 Then: DATA_VECTOR = MATRIX_PERCENT_FUNC(DATA_VECTOR, LOG_SCALE)

LPM_VAL = ASSET_LPM_FUNC(DATA_VECTOR, TRETURN_VAL, 2)
HPM_VAL = ASSET_HPM_FUNC(DATA_VECTOR, TRETURN_VAL, 1)

If LPM_VAL <> 0 Then
    ASSET_UP_POTENTIAL_FUNC = HPM_VAL / LPM_VAL ^ (1 / 2)
Else
    ASSET_UP_POTENTIAL_FUNC = CVErr(xlErrNA)
End If

Exit Function
ERROR_LABEL:
ASSET_UP_POTENTIAL_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : ASSET_DD_DISCRETE_FUNC
'DESCRIPTION   : Calculates drawdown vector from a vector of discrete returns of format
'LIBRARY       : FINAN_ASSET
'GROUP         : RISK_ADJ
'ID            : 017
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 06/04/2010
'************************************************************************************
'************************************************************************************

Private Function ASSET_DD_DISCRETE_FUNC(ByRef DATA_RNG As Variant, _
Optional ByVal DATA_TYPE As Integer = 0)

Dim i As Long
Dim NROWS As Long

Dim MAX1_VAL As Double
Dim TEMP_VECTOR As Variant

Dim DATA_VECTOR As Variant

On Error GoTo ERROR_LABEL

DATA_VECTOR = DATA_RNG
If DATA_TYPE <> 0 Then: DATA_VECTOR = MATRIX_PERCENT_FUNC(DATA_VECTOR, 0)
NROWS = UBound(DATA_VECTOR, 1)

ReDim TEMP_VECTOR(1 To NROWS, 1 To 2)

i = 1
TEMP_VECTOR(i, 2) = 1 + DATA_VECTOR(i, 1)
MAX1_VAL = TEMP_VECTOR(i, 2)
TEMP_VECTOR(i, 1) = -1 + TEMP_VECTOR(i, 2) / MAX1_VAL
        
For i = 2 To NROWS
    TEMP_VECTOR(i, 2) = TEMP_VECTOR(i - 1, 2) * (1 + DATA_VECTOR(i, 1))
    MAX1_VAL = MAXIMUM_FUNC(TEMP_VECTOR(i, 2), MAX1_VAL)
    TEMP_VECTOR(i, 1) = -1 + TEMP_VECTOR(i, 2) / MAX1_VAL
Next i

TEMP_VECTOR = MATRIX_QUICK_SORT_FUNC(TEMP_VECTOR, 1, 1)
ASSET_DD_DISCRETE_FUNC = TEMP_VECTOR

Exit Function
ERROR_LABEL:
ASSET_DD_DISCRETE_FUNC = Err.number
End Function

