Attribute VB_Name = "STAT_MOMENTS_CROSS_LIBR"


Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.

'************************************************************************************
'************************************************************************************
'FUNCTION      : MATRIX_CROSS_SECTIONAL_FUNC
'DESCRIPTION   : Functions to calculate cross-sectional volatility, covariance,
'correlations, and average volatilities, covariances and correlations.
'LIBRARY       : STATISTICS
'GROUP         : CROSS-SECTIONAL
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/08/2008
'************************************************************************************
'************************************************************************************

Function MATRIX_CROSS_SECTIONAL_FUNC(ByRef DATA_RNG As Variant, _
Optional ByVal MA_PERIOD As Long = 12, _
Optional ByVal MULT_FACTOR As Double = 100)


Dim h As Long
Dim i As Long
Dim j As Long
Dim k As Long
Dim l As Long
Dim m As Long

Dim NROWS As Long
Dim NCOLUMNS As Long

Dim TEMP1_SUM As Double
Dim TEMP2_SUM As Double

Dim TEMP3_SUM As Double
Dim TEMP4_SUM As Double

Dim TEMP5_SUM As Double
Dim TEMP6_SUM As Double

Dim TEMP7_SUM As Double
Dim TEMP8_SUM As Double

Dim MEAN_VAL As Double
Dim SIGMA_VAL As Double

Dim TEMP_MATRIX As Variant
Dim COVAR_MATRIX As Variant
Dim DATA_MATRIX As Variant

On Error GoTo ERROR_LABEL

DATA_MATRIX = DATA_RNG
NROWS = UBound(DATA_MATRIX, 1)
NCOLUMNS = UBound(DATA_MATRIX, 2)

ReDim TEMP_MATRIX(0 To NROWS, 1 To 11)

TEMP_MATRIX(0, 1) = "DISPERSION"
TEMP_MATRIX(0, 2) = "AVERAGE DISPERSION"
TEMP_MATRIX(0, 3) = "CROSS-SECTIONAL VOLATILITY"
TEMP_MATRIX(0, 4) = "CROSS-SECTIONAL COVARIANCE"
TEMP_MATRIX(0, 5) = "CROSS-SECTIONAL CORRELATION"
TEMP_MATRIX(0, 6) = "AVG CROSS-SECTIONAL VOLATILITY"
TEMP_MATRIX(0, 7) = "AVG CROSS-SECTIONAL COVARIANCE"
TEMP_MATRIX(0, 8) = "AVG CROSS-SECTIONAL CORRELATION"
TEMP_MATRIX(0, 9) = "AVG VOLATILITY"
TEMP_MATRIX(0, 10) = "AVG COVARIANCE"
TEMP_MATRIX(0, 11) = "AVG CORRELATION"

k = 0
TEMP1_SUM = 0
For l = 1 To NROWS
    MEAN_VAL = 0
    For j = 1 To NCOLUMNS
        MEAN_VAL = MEAN_VAL + DATA_MATRIX(l, j)
        TEMP_MATRIX(l, 3) = TEMP_MATRIX(l, 3) + DATA_MATRIX(l, j) ^ 2
    Next j
    MEAN_VAL = MEAN_VAL / NCOLUMNS
    TEMP_MATRIX(l, 3) = (TEMP_MATRIX(l, 3) / NCOLUMNS) ^ 0.5 * MULT_FACTOR
    'computes cross-sectional volatility from a vector of returns
    
    SIGMA_VAL = 0
    For j = 1 To NCOLUMNS
        SIGMA_VAL = SIGMA_VAL + (MEAN_VAL - DATA_MATRIX(l, j)) ^ 2
        For i = 1 To NCOLUMNS
            If j <> i Then
                TEMP_MATRIX(l, 4) = TEMP_MATRIX(l, 4) + DATA_MATRIX(l, i) * DATA_MATRIX(l, j)
            End If
        Next i
    Next j
    SIGMA_VAL = (SIGMA_VAL / NCOLUMNS) ^ 0.5
    
    TEMP_MATRIX(l, 4) = (TEMP_MATRIX(l, 4) / (NCOLUMNS ^ 2 - NCOLUMNS)) * MULT_FACTOR ^ 2
    ' computes cross-sectional covariance from a vector of returns
    TEMP_MATRIX(l, 5) = (TEMP_MATRIX(l, 4) / TEMP_MATRIX(l, 3) ^ 2)
    ' computes cross-sectional correlation from a vector of returns
'--------------------------------------------------------------
    TEMP_MATRIX(l, 1) = MULT_FACTOR * SIGMA_VAL 'Dispersion

    TEMP1_SUM = TEMP1_SUM + TEMP_MATRIX(l, 1)
    TEMP2_SUM = TEMP1_SUM 'Average Dispersion

    TEMP3_SUM = TEMP3_SUM + TEMP_MATRIX(l, 3)
    TEMP4_SUM = TEMP3_SUM 'Avg Cross-Sectional Volatility

    TEMP5_SUM = TEMP5_SUM + TEMP_MATRIX(l, 4)
    TEMP6_SUM = TEMP5_SUM 'Avg Cross-Sectional Covariance

    TEMP7_SUM = TEMP7_SUM + TEMP_MATRIX(l, 5)
    TEMP8_SUM = TEMP7_SUM 'Avg Cross-Sectional Correlation

'--------------------------------------------------------------
    If l >= MA_PERIOD Then
'--------------------------------------------------------------
        If l > MA_PERIOD Then
            For j = 1 To k
                TEMP2_SUM = TEMP2_SUM - TEMP_MATRIX(j, 1)
                TEMP4_SUM = TEMP4_SUM - TEMP_MATRIX(j, 3)
                TEMP6_SUM = TEMP6_SUM - TEMP_MATRIX(j, 4)
                TEMP8_SUM = TEMP8_SUM - TEMP_MATRIX(j, 5)
            Next j
            TEMP_MATRIX(l, 2) = TEMP2_SUM / MA_PERIOD 'Average Dispersion
            TEMP_MATRIX(l, 6) = TEMP4_SUM / MA_PERIOD 'Avg Cross-Sectional Volatility
            TEMP_MATRIX(l, 7) = TEMP6_SUM / MA_PERIOD 'Avg Cross-Sectional Covariance
            TEMP_MATRIX(l, 8) = TEMP8_SUM / MA_PERIOD 'Avg Cross-Sectional Correlation
            k = k + 1
        Else 'If l = MA_PERIOD Then
            TEMP_MATRIX(l, 2) = TEMP2_SUM / MA_PERIOD 'Average Dispersion
            TEMP_MATRIX(l, 6) = TEMP4_SUM / MA_PERIOD 'Avg Cross-Sectional Volatility
            TEMP_MATRIX(l, 7) = TEMP6_SUM / MA_PERIOD 'Avg Cross-Sectional Covariance
            TEMP_MATRIX(l, 8) = TEMP8_SUM / MA_PERIOD 'Avg Cross-Sectional Correlation
            k = k + 1
        End If
        
        ReDim TEMP_ARR(1 To NCOLUMNS) 'average for each column
        ReDim COVAR_MATRIX(1 To NCOLUMNS, 1 To NCOLUMNS)
        'compute the average
        For j = 1 To NCOLUMNS
            h = k
            TEMP_ARR(j) = 0
            For i = 1 To MA_PERIOD
                TEMP_ARR(j) = TEMP_ARR(j) + DATA_MATRIX(h, j)
                h = h + 1
            Next i
            TEMP_ARR(j) = TEMP_ARR(j) / MA_PERIOD
        Next j
        'compute the cross covariance matrix
        For i = 1 To NCOLUMNS
            For j = 1 To NCOLUMNS
                If j < i Then
                    COVAR_MATRIX(i, j) = COVAR_MATRIX(j, i)
                Else
                    COVAR_MATRIX(i, j) = 0
                    h = k
                    For m = 1 To MA_PERIOD
                        COVAR_MATRIX(i, j) = COVAR_MATRIX(i, j) + (DATA_MATRIX(h, i) - TEMP_ARR(i)) * (DATA_MATRIX(h, j) - TEMP_ARR(j))
                        h = h + 1
                    Next m
                    COVAR_MATRIX(i, j) = COVAR_MATRIX(i, j) / MA_PERIOD
                End If
            Next j
        Next i
'--------------------------------------------------------------
        'ReDim COVAR_MATRIX(1 To MA_PERIOD, 1 To NCOLUMNS)
        'h = k
        'For i = 1 To MA_PERIOD
            'For j = 1 To NCOLUMNS
                'COVAR_MATRIX(i, j) = DATA_MATRIX(h, j)
            'Next j
            'h = h + 1
        'Next i
       ' COVAR_MATRIX = MATRIX_COVARIANCE_FRAME2_FUNC(COVAR_MATRIX)
'--------------------------------------------------------------
        TEMP_MATRIX(l, 9) = 0
        For i = 1 To NCOLUMNS
            TEMP_MATRIX(l, 9) = TEMP_MATRIX(l, 9) + COVAR_MATRIX(i, i) ^ 0.5
        Next i
        TEMP_MATRIX(l, 9) = TEMP_MATRIX(l, 9) / NCOLUMNS * MULT_FACTOR
        'Avg Volatility
'--------------------------------------------------------------
        TEMP_MATRIX(l, 10) = 0
        For i = 1 To NCOLUMNS - 1
            For j = i + 1 To NCOLUMNS
               TEMP_MATRIX(l, 10) = TEMP_MATRIX(l, 10) + COVAR_MATRIX(i, j)
            Next j
        Next i
        TEMP_MATRIX(l, 10) = TEMP_MATRIX(l, 10) * 2 / (NCOLUMNS * (NCOLUMNS - 1)) * MULT_FACTOR ^ 2
        'Avg Covariance
'--------------------------------------------------------------
        'COVAR_MATRIX = MATRIX_CORRELATION_COVARIANCE_FUNC(COVAR_MATRIX)
        TEMP_MATRIX(l, 11) = 0
        For i = 1 To NCOLUMNS - 1
            For j = i + 1 To NCOLUMNS
                TEMP_MATRIX(l, 11) = TEMP_MATRIX(l, 11) + COVAR_MATRIX(i, j) / ((COVAR_MATRIX(i, i) * COVAR_MATRIX(j, j)) ^ 0.5)  'COVAR_MATRIX(i, j)
            Next j
        Next i
        TEMP_MATRIX(l, 11) = TEMP_MATRIX(l, 11) * 2 / (NCOLUMNS * (NCOLUMNS - 1))
        'Avg Correlation
'--------------------------------------------------------------
    Else
        TEMP_MATRIX(l, 2) = ""
        For j = 6 To 11: TEMP_MATRIX(l, j) = "": Next j
    End If
'--------------------------------------------------------------
Next l
'--------------------------------------------------------------

MATRIX_CROSS_SECTIONAL_FUNC = TEMP_MATRIX

Exit Function
ERROR_LABEL:
MATRIX_CROSS_SECTIONAL_FUNC = Err.number
End Function
