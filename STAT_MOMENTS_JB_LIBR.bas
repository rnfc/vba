Attribute VB_Name = "STAT_MOMENTS_JB_LIBR"


Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.

'************************************************************************************
'************************************************************************************
'FUNCTION      : JARQUE_BERA_HYPOTHESIS_MODEL_FUNC
'DESCRIPTION   : Jarque-Bera hypothesis model for normality
'LIBRARY       : STATISTICS
'GROUP         : JARQUE_BERA
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/22/2009
'************************************************************************************
'************************************************************************************

Function JARQUE_BERA_HYPOTHESIS_MODEL_FUNC(ByVal SKEW As Double, _
ByVal KURTOSIS As Double, _
ByVal NSIZE As Long, _
Optional ByVal CI_FACTOR As Double = 0.05, _
Optional ByVal OUTPUT As Integer = 0)

' p-Value: Likelihood, under the assumption that the data is normal, that the
'data would yield the obtained results. If the p-value is less than the significance
'level, then the result constitutes evidence against the null hypothesis.

Dim PValue As Double
Dim JSTAT_VAL As Double

'SKEW & KURTOSIS: Calculate 3rd and 4th moments (Skewness and Kurtosis)

On Error GoTo ERROR_LABEL
    
JSTAT_VAL = NSIZE * ((SKEW ^ 2) / 6 + (KURTOSIS ^ 2) / 24)
PValue = CHI_SQUARED_DIST_FUNC(JSTAT_VAL, 2, True, False)

Select Case OUTPUT
    Case 0
        JARQUE_BERA_HYPOTHESIS_MODEL_FUNC = (CI_FACTOR < PValue)
    Case 1
        JARQUE_BERA_HYPOTHESIS_MODEL_FUNC = JSTAT_VAL 'J-Stat
    Case Else
        JARQUE_BERA_HYPOTHESIS_MODEL_FUNC = PValue 'J-PValue
End Select

Exit Function
ERROR_LABEL:
JARQUE_BERA_HYPOTHESIS_MODEL_FUNC = Err.number 'CVErr(xlErrValue)
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : JARQUE_BERA_NORMALITY_TEST_FUNC
'DESCRIPTION   : Test for normality with the Jarque-Bera test
'LIBRARY       : STATISTICS
'GROUP         : JARQUE_BERA
'ID            : 002
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/22/2009
'************************************************************************************
'************************************************************************************

Function JARQUE_BERA_NORMALITY_TEST_FUNC(ByRef DATA_RNG As Variant, _
Optional ByVal CONFIDENCE_VAL As Double = 0.05, _
Optional ByVal DATA_TYPE As Integer = 0, _
Optional ByVal LOG_SCALE As Integer = 0)

Dim i As Long
Dim NROWS As Long

Dim P_VAL As Double
Dim JB_VAL As Double

Dim MEAN_VAL As Double
Dim SIGMA_VAL As Double
Dim SKEW_VAL As Double
Dim KURT_VAL As Double

Dim DATA_VECTOR As Variant
Dim NORM_VECTOR As Variant

On Error GoTo ERROR_LABEL

DATA_VECTOR = DATA_RNG
If DATA_TYPE <> 0 Then: DATA_VECTOR = MATRIX_PERCENT_FUNC(DATA_VECTOR, LOG_SCALE)
NROWS = UBound(DATA_VECTOR, 1)

MEAN_VAL = MATRIX_MEAN_FUNC(DATA_VECTOR)(1, 1)
SIGMA_VAL = MATRIX_STDEV_FUNC(DATA_VECTOR)(1, 1)

' Normalize returns
ReDim NORM_VECTOR(1 To NROWS, 1 To 1)
For i = 1 To NROWS
    NORM_VECTOR(i, 1) = (DATA_VECTOR(i, 1) - MEAN_VAL) / SIGMA_VAL
Next i

' Calculate 3rd and 4th moments (skewness and kurtosis)
SKEW_VAL = 0
KURT_VAL = 0
For i = 1 To NROWS
    SKEW_VAL = SKEW_VAL + NORM_VECTOR(i, 1) ^ 3
    KURT_VAL = KURT_VAL + NORM_VECTOR(i, 1) ^ 4
Next i
SKEW_VAL = SKEW_VAL / NROWS
KURT_VAL = KURT_VAL / NROWS - 3

JB_VAL = NROWS * ((SKEW_VAL ^ 2) / 6 + (KURT_VAL ^ 2) / 24)
P_VAL = CHI_SQUARED_DIST_FUNC(JB_VAL, 2, True, False)

JARQUE_BERA_NORMALITY_TEST_FUNC = (CONFIDENCE_VAL < P_VAL)

Exit Function
ERROR_LABEL:
JARQUE_BERA_NORMALITY_TEST_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : JARQUE_BERA_PVALUE_FUNC
'DESCRIPTION   : JB P-Value
'LIBRARY       : STATISTICS
'GROUP         : JARQUE_BERA
'ID            : 003
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/22/2009
'************************************************************************************
'************************************************************************************


Function JARQUE_BERA_PVALUE_FUNC(ByRef DATA_RNG As Variant, _
Optional ByVal CONFIDENCE_VAL As Double = 0.05, _
Optional ByVal DATA_TYPE As Integer = 0, _
Optional ByVal LOG_SCALE As Integer = 0)

Dim i As Long
Dim NROWS As Long

Dim JB_VAL As Double

Dim MEAN_VAL As Double
Dim SIGMA_VAL As Double
Dim SKEW_VAL As Double
Dim KURT_VAL As Double

Dim DATA_VECTOR As Variant
Dim NORM_VECTOR As Variant

On Error GoTo ERROR_LABEL

DATA_VECTOR = DATA_RNG
If DATA_TYPE <> 0 Then: DATA_VECTOR = MATRIX_PERCENT_FUNC(DATA_VECTOR, LOG_SCALE)
NROWS = UBound(DATA_VECTOR, 1)

MEAN_VAL = MATRIX_MEAN_FUNC(DATA_VECTOR)(1, 1)
SIGMA_VAL = MATRIX_STDEV_FUNC(DATA_VECTOR)(1, 1)

' Normalize returns
ReDim NORM_VECTOR(1 To NROWS, 1 To 1)
For i = 1 To NROWS
    NORM_VECTOR(i, 1) = (DATA_VECTOR(i, 1) - MEAN_VAL) / SIGMA_VAL
Next i

' Calculate 3rd and 4th moments (skewness and kurtosis)
SKEW_VAL = 0: KURT_VAL = 0
For i = 1 To NROWS
    SKEW_VAL = SKEW_VAL + NORM_VECTOR(i, 1) ^ 3
    KURT_VAL = KURT_VAL + NORM_VECTOR(i, 1) ^ 4
Next i
SKEW_VAL = SKEW_VAL / NROWS
KURT_VAL = KURT_VAL / NROWS - 3

JB_VAL = NROWS * ((SKEW_VAL ^ 2) / 6 + (KURT_VAL ^ 2) / 24)
JARQUE_BERA_PVALUE_FUNC = CHI_SQUARED_DIST_FUNC(JB_VAL, 2, True, False)

Exit Function
ERROR_LABEL:
JARQUE_BERA_PVALUE_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : JARQUE_BERA_CRITICAL_VALUE_FUNC
'DESCRIPTION   : JB Critical Value
'LIBRARY       : STATISTICS
'GROUP         : JARQUE_BERA
'ID            : 004
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/22/2009
'************************************************************************************
'************************************************************************************

Function JARQUE_BERA_CRITICAL_VALUE_FUNC(Optional ByVal CONFIDENCE_VAL As Double = 0.05)
On Error GoTo ERROR_LABEL
JARQUE_BERA_CRITICAL_VALUE_FUNC = INVERSE_CHI_SQUARED_DIST_FUNC(CONFIDENCE_VAL, 2, False)
Exit Function
ERROR_LABEL:
JARQUE_BERA_CRITICAL_VALUE_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : JARQUE_BERA_STATISTICS_FUNC
'DESCRIPTION   : Jarque-Bera Statistics
'LIBRARY       : STATISTICS
'GROUP         : JARQUE_BERA
'ID            : 005
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/22/2009
'************************************************************************************
'************************************************************************************

Function JARQUE_BERA_STATISTICS_FUNC(ByRef DATA_RNG As Variant, _
Optional ByVal CONFIDENCE_VAL As Double = 0.05, _
Optional ByVal DATA_TYPE As Integer = 0, _
Optional ByVal LOG_SCALE As Integer = 0)

Dim i As Long
Dim NROWS As Long

Dim MEAN_VAL As Double
Dim SIGMA_VAL As Double
Dim SKEW_VAL As Double
Dim KURT_VAL As Double

Dim DATA_VECTOR As Variant
Dim NORM_VECTOR As Variant

On Error GoTo ERROR_LABEL

DATA_VECTOR = DATA_RNG
If DATA_TYPE <> 0 Then: DATA_VECTOR = MATRIX_PERCENT_FUNC(DATA_VECTOR, LOG_SCALE)
NROWS = UBound(DATA_VECTOR, 1)

MEAN_VAL = MATRIX_MEAN_FUNC(DATA_VECTOR)(1, 1)
SIGMA_VAL = MATRIX_STDEV_FUNC(DATA_VECTOR)(1, 1)

' Normalize returns
ReDim NORM_VECTOR(1 To NROWS, 1 To 1)
For i = 1 To NROWS
    NORM_VECTOR(i, 1) = (DATA_VECTOR(i, 1) - MEAN_VAL) / SIGMA_VAL
Next i

' Calculate 3rd and 4th moments (skewness and kurtosis)
SKEW_VAL = 0
KURT_VAL = 0
For i = 1 To NROWS
    SKEW_VAL = SKEW_VAL + NORM_VECTOR(i, 1) ^ 3
    KURT_VAL = KURT_VAL + NORM_VECTOR(i, 1) ^ 4
Next i
SKEW_VAL = SKEW_VAL / NROWS
KURT_VAL = KURT_VAL / NROWS - 3

JARQUE_BERA_STATISTICS_FUNC = NROWS * ((SKEW_VAL ^ 2) / 6 + (KURT_VAL ^ 2) / 24)

Exit Function
ERROR_LABEL:
JARQUE_BERA_STATISTICS_FUNC = Err.number
End Function

