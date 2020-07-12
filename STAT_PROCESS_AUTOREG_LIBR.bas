Attribute VB_Name = "STAT_PROCESS_AUTOREG_LIBR"

'-------------------------------------------------------------------------------------------------
'-------------------------------------------------------------------------------------------------
Option Base 1
Option Explicit
'-------------------------------------------------------------------------------------------------
'-------------------------------------------------------------------------------------------------

'************************************************************************************
'************************************************************************************
'FUNCTION      : VECTOR_AUTOREGRESSION_FUNC
'DESCRIPTION   : vector autoregression on multiple time series
'LIBRARY       : STATISTICS
'GROUP         :
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 05/26/2011
'************************************************************************************
'************************************************************************************

Function VECTOR_AUTOREGRESSION_FUNC(ByRef DATA_RNG As Variant, _
Optional ByVal nLAGS As Long = 2, _
Optional ByVal OUTPUT As Integer = 0)

Dim i As Long
Dim j As Long
Dim k As Long
Dim l As Long

Dim NSIZE As Long
Dim NROWS As Long
Dim NCOLUMNS As Long

Dim TEMP_SUM As Double

Dim OLS_VECTOR As Variant
Dim COEFFICIENTS_MATRIX As Variant 'coefficient of matrix of dependent variables
Dim FITTED_MATRIX As Variant 'values of timeseries calculated after regression
Dim INPUTS_MATRIX As Variant 'values of timeseries after discarding the first nlag terms

Dim TEMP1_MATRIX As Variant
Dim TEMP2_MATRIX As Variant
Dim TEMP3_MATRIX As Variant
Dim DATA_MATRIX As Variant

On Error GoTo ERROR_LABEL

DATA_MATRIX = DATA_RNG
NROWS = UBound(DATA_MATRIX, 1)
NSIZE = NROWS - nLAGS
NCOLUMNS = UBound(DATA_MATRIX, 2)
ReDim TEMP1_MATRIX(1 To NROWS, 1 To NCOLUMNS * nLAGS)
l = 1
For j = 1 To NCOLUMNS
    For k = 1 To nLAGS
        For i = 1 To NROWS - k
            TEMP1_MATRIX(i + k, l) = DATA_MATRIX(i, j)
        Next i
        l = l + 1
    Next k
Next j
ReDim TEMP2_MATRIX(1 To NSIZE, 1 To NCOLUMNS * nLAGS + 1)
k = NCOLUMNS * nLAGS
For i = 1 To NSIZE
    For j = 1 To k
        TEMP2_MATRIX(i, j) = TEMP1_MATRIX(i + nLAGS, j)
    Next j
Next i
For i = 1 To NSIZE
    TEMP2_MATRIX(i, NCOLUMNS * nLAGS + 1) = 1
Next i

ReDim YTEMP_VECTOR(1 To NSIZE, 1 To 1)
ReDim COEFFICIENTS_MATRIX(1 To k + 1, 1 To NCOLUMNS) ' * NLAGS + 1)
ReDim FITTED_MATRIX(1 To NSIZE, 1 To NCOLUMNS)
ReDim INPUTS_MATRIX(1 To NSIZE, 1 To NCOLUMNS)

For j = 1 To NCOLUMNS
    For i = 1 To NSIZE
        YTEMP_VECTOR(i, 1) = DATA_MATRIX(i + nLAGS, j)
    Next i
    TEMP3_MATRIX = MATRIX_TRANSPOSE_FUNC(TEMP2_MATRIX)
    OLS_VECTOR = MMULT_FUNC(MMULT_FUNC(MATRIX_INVERSE_FUNC(MMULT_FUNC(TEMP3_MATRIX, TEMP2_MATRIX, 70), 2), TEMP3_MATRIX, 70), YTEMP_VECTOR, 70)
    For i = 1 To UBound(OLS_VECTOR, 1): COEFFICIENTS_MATRIX(i, j) = OLS_VECTOR(i, 1): Next i
    For i = 1 To NSIZE
        TEMP_SUM = 0
        For k = 1 To UBound(OLS_VECTOR, 1): TEMP_SUM = TEMP_SUM + OLS_VECTOR(k, 1) * TEMP2_MATRIX(i, k): Next k
        FITTED_MATRIX(i, j) = TEMP_SUM
        INPUTS_MATRIX(i, j) = YTEMP_VECTOR(i, 1)
    Next i
Next j
'----------------------------------------------------------------------------------------
Select Case OUTPUT
'----------------------------------------------------------------------------------------
Case 0 'Beta Coefficients --> Perfect
'----------------------------------------------------------------------------------------
    VECTOR_AUTOREGRESSION_FUNC = COEFFICIENTS_MATRIX
'----------------------------------------------------------------------------------------
Case 1 'Fitted Values
'----------------------------------------------------------------------------------------
    VECTOR_AUTOREGRESSION_FUNC = FITTED_MATRIX
'----------------------------------------------------------------------------------------
Case 2 'Inputs Value with lags
'----------------------------------------------------------------------------------------
    VECTOR_AUTOREGRESSION_FUNC = INPUTS_MATRIX
'----------------------------------------------------------------------------------------
Case Else
'----------------------------------------------------------------------------------------
    VECTOR_AUTOREGRESSION_FUNC = Array(COEFFICIENTS_MATRIX, FITTED_MATRIX, INPUTS_MATRIX)
'----------------------------------------------------------------------------------------
End Select
'----------------------------------------------------------------------------------------
Exit Function
ERROR_LABEL:
VECTOR_AUTOREGRESSION_FUNC = Err.number
End Function

