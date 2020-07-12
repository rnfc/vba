Attribute VB_Name = "STAT_PROCESS_DW_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
'default first index of arrays.

'************************************************************************************
'************************************************************************************
'FUNCTION      : DURBIN_WATSON_AC_TEST_FUNC
'DESCRIPTION   :
'LIBRARY       : STATISTICS
'GROUP         : DW
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 11/11/2012
'************************************************************************************
'************************************************************************************

Function DURBIN_WATSON_AC_TEST_FUNC(ByVal XDATA_RNG As Variant, _
ByVal RESIDUALS_RNG As Variant, _
Optional ByVal INTERCEPT_FLAG As Boolean = True)

Dim i As Long
Dim j As Long
Dim k As Long

Dim NROWS As Long
Dim NCOLUMNS As Long

Dim TEMP1_VAL As Double
Dim TEMP2_VAL As Double

Dim DF_VAL As Long
Dim RHO_VAL As Double

Dim DW_PVAL As Variant
Dim DW_SVAL As Double

Dim EP1_VAL As Double
Dim EP2_VAL As Double
Dim EP_ARR As Variant

Dim MEAN_VAL As Double
Dim SIGMA_VAL As Double

Dim RESIDUAL_VAL As Double
Dim ELEMENT_VAL As Double

Dim TEMP1_VECTOR As Variant
Dim TEMP2_VECTOR As Variant
Dim TEMP_MATRIX As Variant

Dim M_MATRIX As Variant
Dim A_MATRIX As Variant

Dim X_MATRIX As Variant
Dim XT_MATRIX As Variant
Dim XTX_MATRIX As Variant
Dim XTXV_MATRIX As Variant
Dim XTXVD_MATRIX As Variant
Dim XXTXV_MATRIX As Variant
Dim XXTXVXT_MATRIX As Variant

Dim MA_MATRIX As Variant
Dim MAM_MATRIX As Variant

Dim DATA_ARR As Variant
Dim XDATA_MATRIX As Variant
Dim RESIDUALS_VECTOR As Variant

Dim ERROR_STR As String
On Error GoTo ERROR_LABEL

XDATA_MATRIX = XDATA_RNG
NROWS = UBound(XDATA_MATRIX, 1) 'No Obs
NCOLUMNS = UBound(XDATA_MATRIX, 2) 'No XVars

RESIDUALS_VECTOR = RESIDUALS_RNG
If UBound(RESIDUALS_VECTOR, 1) = 1 Then
    RESIDUALS_VECTOR = MATRIX_TRANSPOSE_FUNC(RESIDUALS_VECTOR)
End If
If UBound(RESIDUALS_VECTOR, 1) <> NROWS Then: GoTo ERROR_LABEL

'Create and fill X and XTranspose matrices
'Put in the column of 1s if there is an intercept term
If INTERCEPT_FLAG = True Then
    NCOLUMNS = NCOLUMNS + 1
    ReDim X_MATRIX(1 To NROWS, 1 To NCOLUMNS)
    ReDim XT_MATRIX(1 To NCOLUMNS, 1 To NROWS)
    For i = 1 To NROWS
        X_MATRIX(i, 1) = 1
        XT_MATRIX(1, i) = 1
        For j = 2 To NCOLUMNS
            X_MATRIX(i, j) = XDATA_MATRIX(i, j - 1)
            XT_MATRIX(j, i) = XDATA_MATRIX(i, j - 1)
        Next j
    Next i
Else
    ReDim X_MATRIX(1 To NROWS, 1 To NCOLUMNS)
    ReDim XT_MATRIX(1 To NCOLUMNS, 1 To NROWS)
    For i = 1 To NROWS
        For j = 1 To NCOLUMNS
            X_MATRIX(i, j) = XDATA_MATRIX(i, j)
            XT_MATRIX(j, i) = XDATA_MATRIX(i, j)
        Next j
    Next i
End If

' X'X work
' Get X'X
'ReDim XTX_MATRIX(1 To NCOLUMNS, 1 To NCOLUMNS)
XTX_MATRIX = MMULT_FUNC(XT_MATRIX, X_MATRIX, 70)

' (X'X)-1 work
' Get (X'X)-1
XTXV_MATRIX = MATRIX_INVERSE_FUNC(XTX_MATRIX, 0) 'Application.WorksheetFunction.MInverse(XTX_MATRIX)
ReDim XTXVD_MATRIX(1 To NCOLUMNS, 1 To NCOLUMNS)
For i = 1 To NCOLUMNS
    For j = 1 To NCOLUMNS
        XTXVD_MATRIX(i, j) = XTXV_MATRIX(i, j)
    Next j
Next i

'XTransposeXInverseD = VariantToDouble(XTransposeXInverse) 'This works also with () in Variant

' Form the M matrix -- I - X(X'X)-1X'
' Get X(X'X)-1
XXTXV_MATRIX = MMULT_FUNC(X_MATRIX, XTXVD_MATRIX, 70)
'Get X(X'X)-1X'
XXTXVXT_MATRIX = MMULT_FUNC(XXTXV_MATRIX, XT_MATRIX, 70)
' Get I - X(X'X)-1X'
ReDim M_MATRIX(1 To NROWS, 1 To NROWS)
For i = 1 To NROWS
    For j = 1 To NROWS
        If i = j Then
            M_MATRIX(i, j) = 1 - XXTXVXT_MATRIX(i, j)
        Else
            M_MATRIX(i, j) = 0 - XXTXVXT_MATRIX(i, j)
        End If
    Next j
Next i

' DW work
' Find DW d statistic
' Numerator is e'Ae
' Find A
ReDim A_MATRIX(1 To NROWS, 1 To NROWS)
For i = 1 To NROWS
    For j = 1 To NROWS
        'the diagonal
        If i = j Then
            If i = 1 Or i = NROWS Then
                A_MATRIX(i, j) = 1
            Else
                A_MATRIX(i, j) = 2
            End If
        'the -1 off the diagonal
        ElseIf Abs(i - j) = 1 Then
            A_MATRIX(i, j) = -1
            'the zeroes
        Else
            A_MATRIX(i, j) = 0
        End If
    Next j
Next i

' Find e'A
ReDim EP_ARR(1 To NROWS)
For i = 1 To NROWS
    RESIDUAL_VAL = RESIDUALS_VECTOR(i, 1)
    For j = 1 To NROWS
        ELEMENT_VAL = RESIDUAL_VAL * A_MATRIX(i, j)
        EP_ARR(j) = EP_ARR(j) + ELEMENT_VAL
    Next j
Next i

' Find e'Ae
For i = 1 To NROWS
    ELEMENT_VAL = RESIDUALS_VECTOR(i, 1) * EP_ARR(i)
    EP2_VAL = EP2_VAL + ELEMENT_VAL
Next i

' Denominator is e'e

For i = 1 To NROWS
    ELEMENT_VAL = RESIDUALS_VECTOR(i, 1) ^ 2
    EP1_VAL = EP1_VAL + ELEMENT_VAL
Next i

' DW d statistic
DW_SVAL = EP2_VAL / EP1_VAL

' Find MAM
' Form MA
'starttime = Timer
MA_MATRIX = MMULT_FUNC(M_MATRIX, A_MATRIX, 70)
' Form MAM
MAM_MATRIX = MMULT_FUNC(MA_MATRIX, M_MATRIX, 70)

' Get the DF for distribution
DF_VAL = NROWS - NCOLUMNS

'There is an issue to handle when DF is large.  The Farebrother
'algorithm may fail, then you use the normal approx to compute the P value.
'We'll give the user a chance to just use the normal approx to save the
'time required to attempt to get the exact DW P Value using Farebrother
MEAN_VAL = 0: SIGMA_VAL = 0
If DF_VAL > 100 Then
1984:
    'We need to compute the mean and variance for the normal approximation
    TEMP_MATRIX = MMULT_FUNC(MA_MATRIX, MA_MATRIX, 70)
    TEMP1_VAL = 0: TEMP2_VAL = 0
    For i = 1 To NROWS
        TEMP1_VAL = MA_MATRIX(i, i) + TEMP1_VAL
        TEMP2_VAL = TEMP_MATRIX(i, i) + TEMP2_VAL
    Next i
    'so, we have the mean and variance then we can use Normdist
    'to compute the P Value for the DWStat value
    'Use the normal sheet of this workbook!
    '1983 the Eigenvalue and PValue calculations below
    MEAN_VAL = TEMP1_VAL / DF_VAL 'Mean
    SIGMA_VAL = Sqr((2 / (DF_VAL ^ 2)) * (TEMP2_VAL - (TEMP1_VAL * TEMP1_VAL / DF_VAL))) 'Sigma
    DW_PVAL = NORMDIST_FUNC(DW_SVAL, MEAN_VAL, SIGMA_VAL, 0)
    If (DW_SVAL - MEAN_VAL) < 0 Then
        DW_PVAL = 2 * DW_PVAL
    Else
        DW_PVAL = 2 * (1 - DW_PVAL)
    End If
Else 'Exact DW P Value. If it fails, the normal approximation will automatically be used.
    ' Get the eigenvalues of MAM and get the DW PValue
    ' See below for more on this function
    ERROR_STR = ""
    DW_PVAL = DURBIN_WATSON_PVALUE_FUNC(MAM_MATRIX, NROWS, DW_SVAL, DF_VAL, ERROR_STR)
    If ERROR_STR <> "" Or DW_PVAL = 2 Then: GoTo 1984
End If

'Compute the r(Res,LagRes) and report as the estimated rho
ReDim TEMP1_VECTOR(1 To NROWS - 1, 1 To 1) 'residuals2ton
ReDim TEMP2_VECTOR(1 To NROWS - 1, 1 To 1) 'residuals2tonlagged
For k = 1 To NROWS - 1
    TEMP1_VECTOR(k, 1) = RESIDUALS_VECTOR(k, 1)
    TEMP2_VECTOR(k, 1) = RESIDUALS_VECTOR(k + 1, 1)
Next k
RHO_VAL = CORRELATION_FUNC(TEMP1_VECTOR, TEMP2_VECTOR, 0, 0)

ReDim DATA_ARR(1 To 7)
DATA_ARR(1) = DW_SVAL
DATA_ARR(2) = DW_PVAL
DATA_ARR(3) = RHO_VAL 'Estimated rho (using r(Res, LagRes))
DATA_ARR(4) = DF_VAL
DATA_ARR(5) = "The DW statistic is " & DW_SVAL & " and the DW P Value is " & DW_PVAL & ". Estimated rho (computed via the correlation of the residuals and lagged residuals is " & RHO_VAL & "."
DATA_ARR(6) = MEAN_VAL
DATA_ARR(7) = SIGMA_VAL

DURBIN_WATSON_AC_TEST_FUNC = DATA_ARR

Exit Function
ERROR_LABEL:
DURBIN_WATSON_AC_TEST_FUNC = ERROR_STR
End Function



'Run the DW Analysis

'The Durbin-Watson analysis assumes a first-order autoregressive process. Both the DW d statistic and
'its P-value are reported.

Function DURBIN_WATSON1_FUNC(ByRef RESIDUALS_RNG As Variant)

Dim i As Integer
Dim j As Integer
Dim NROWS As Integer
Dim NCOLUMNS As Integer

Dim NSUM_VAL As Double 'SumNumerator
Dim DSUM_VAL As Double 'SumDenominator

Dim LAG_VAL As Double 'LaggedResiduals
Dim SQUARED_VAL As Double 'ResidualDifferenceSquared

Dim DATA_MATRIX As Variant
Dim TEMP_VECTOR As Variant

On Error GoTo ERROR_LABEL

DATA_MATRIX = RESIDUALS_RNG
NROWS = UBound(DATA_MATRIX, 1)
NCOLUMNS = UBound(DATA_MATRIX, 2)
ReDim TEMP_VECTOR(1 To 1, 1 To NCOLUMNS)

For j = 1 To NCOLUMNS
    'ReDim LAG_VAL(1 To NROWS - 1)
    'ReDim SQUARED_VAL(1 To NROWS - 1)
    NSUM_VAL = 0
    For i = 1 To NROWS - 1
        'LAG_VAL(i) = DATA_MATRIX(i + 1, j)
        LAG_VAL = DATA_MATRIX(i + 1, j)
        'SQUARED_VAL(i) = (DATA_MATRIX(i, j) - LAG_VAL(i)) ^ 2
        SQUARED_VAL = (DATA_MATRIX(i, j) - LAG_VAL) ^ 2
        'NSUM_VAL = NSUM_VAL + SQUARED_VAL(i)
        NSUM_VAL = NSUM_VAL + SQUARED_VAL
    Next i
    
    DSUM_VAL = 0
    For i = 1 To NROWS
        DSUM_VAL = DSUM_VAL + DATA_MATRIX(i, j) ^ 2
    Next i
    If DSUM_VAL <> 0 Then
        TEMP_VECTOR(1, j) = NSUM_VAL / DSUM_VAL
    Else
        TEMP_VECTOR(1, j) = CVErr(xlErrNA)
    End If
Next j

DURBIN_WATSON1_FUNC = TEMP_VECTOR

Exit Function
ERROR_LABEL:
DURBIN_WATSON1_FUNC = Err.number
End Function

Function DURBIN_WATSON2_FUNC(ByVal RESIDUALS_RNG As Variant)

Dim i As Integer
Dim j As Integer
Dim k As Integer

Dim NROWS As Integer
Dim NCOLUMNS As Integer

Dim EP1_VAL As Double
Dim EP2_VAL As Double

Dim RESID_VAL As Double
Dim ELEMENT_VAL As Double

Dim DATA_ARR() As Double
Dim TEMP_MATRIX() As Single
Dim DATA_MATRIX As Variant
Dim TEMP_VECTOR As Variant

On Error GoTo ERROR_LABEL

'Get number of observations, NROWS
DATA_MATRIX = RESIDUALS_RNG
NROWS = UBound(DATA_MATRIX, 1)
NCOLUMNS = UBound(DATA_MATRIX, 2)
ReDim TEMP_VECTOR(1 To 1, 1 To NCOLUMNS)

For k = 1 To NCOLUMNS
    ' Numerator is e'Ae
    ReDim TEMP_MATRIX(1 To NROWS, 1 To NROWS)
    For i = 1 To NROWS
        For j = 1 To NROWS
            'the diagonal
            If i = j Then
                If i = 1 Or i = NROWS Then
                    TEMP_MATRIX(i, j) = 1
                Else
                    TEMP_MATRIX(i, j) = 2
                End If
            ElseIf Abs(i - j) = 1 Then 'the -1 off the diagonal
                TEMP_MATRIX(i, j) = -1 'the zeroes
            Else
                TEMP_MATRIX(i, j) = 0
            End If
        Next j
    Next i
    
    ' Find e'A
    ReDim DATA_ARR(1 To NROWS)
    For i = 1 To NROWS
        RESID_VAL = DATA_MATRIX(i, k)
        For j = 1 To NROWS
            ELEMENT_VAL = RESID_VAL * TEMP_MATRIX(i, j)
            DATA_ARR(j) = DATA_ARR(j) + ELEMENT_VAL
        Next j
    Next i
    
    ' Find e'Ae
    EP1_VAL = 0
    For i = 1 To NROWS
        ELEMENT_VAL = DATA_MATRIX(i, k) * DATA_ARR(i)
        EP1_VAL = EP1_VAL + ELEMENT_VAL
    Next i
    
    ' Denominator is e'e
    EP2_VAL = 0
    For i = 1 To NROWS
        ELEMENT_VAL = DATA_MATRIX(i, k) ^ 2
        EP2_VAL = EP2_VAL + ELEMENT_VAL
    Next i
    
    ' DW d statistic
    If EP2_VAL <> 0 Then
        TEMP_VECTOR(1, k) = EP1_VAL / EP2_VAL
    Else
        TEMP_VECTOR(1, k) = CVErr(xlErrNA)
    End If
Next k

DURBIN_WATSON2_FUNC = TEMP_VECTOR

Exit Function
ERROR_LABEL:
DURBIN_WATSON2_FUNC = Err.number
End Function

'// PERFECT

' This algorithm does a Durbin-Watson analysis of a
' first-order autoregressive process
' It proceeds by forming the X and MAM matrices (including
' the e'Ae/e'e (DW stat) quadratic form), then it accesses
' the DURBIN_WATSON_EIGENVAL_FUNC function, and finally the Durbin_Watson_PVALUE_FUNC function.

' It is assumed that the user has a run a regression of y
' on x1, x2, . . . xK (called NumXVar in the code)
' with N observations (called NObs in the code) and has the
' residuals (by using, for example, LINEST to get the
' predicted values and then calculating residual = actual
' Y - Predicted Y).
' It is assumed that the data are in columns on an
' Excel worksheet.
' The user clicks on the Durbin-Watson Analysis button
' and an informational message is displayed, followed
' by a dialog box where the X data and residuals are
' entered.

' Sources:
' See Judge, et al. 2nd ed. p. 384-401 for more on matrix
' manipulation and A, e'Ae, MAM forms
'
' The code itself is a combination of several sources including
' Judge, et. al. 2nd ed.
' J. P. Imhoff, "Computing the distribution of quadratic
' forms in normal variables," Biometrika, vol. 48 No. 3 and 4
' (1961), pp. 419-426
' http://lib.stat.cmu.edu/apstat/ AS (applied statistics)
' library with its on-line collection of algorithms

' A large part of the code itself has been adapted
' from the following three Fortran 77 programs:
' 1)
' algorithm as 60.1 appl.statist. (1973) vol.22 no.2
'   reduces real symmetric matrix to tridiagonal form
' 2)
' algorithm as 60.2 appl.statist. (1973) vol.22, no.2
'   finds latent roots and vectors of tridiagonal matrix
' 3)
' algorithm AS 153 (AS R52), VOL. 33, 363-366, 1984.
'   ORIGINALLY FROM STATLIB.  REVISED 5/3/1996 BY CLINT CUMMINS
'   finds the p-value of quadratic form appropriate for
'   DW stat
' The algorithms are downloadable from a variety of sites
' on the Internet (including http://lib.stat.cmu.edu/apstat/)
' and the accompanying explanatory articles are in the
' journal titled: APPLIED STATISTICS

' This work was financially supported by the Wabash Center
' for Theology and Teaching.

' Get the P-value

Private Function DURBIN_WATSON_PVALUE_FUNC(ByRef DATA_RNG As Variant, _
ByVal NROWS As Long, _
ByVal DW_VAL As Double, _
ByVal DF_VAL As Long, _
ByRef ERROR_STR As String)

Dim i As Long
'Dim j As Long
Dim k As Long
Dim l As Long

Dim DD As Long

Dim J1 As Long
Dim J2 As Long
Dim J3 As Long
Dim J4 As Long

Dim L1 As Long
Dim L2 As Long

Dim N1 As Long
Dim N2 As Long

Dim H1 As Integer
Dim H2 As Integer

Dim U_VAL As Double
Dim V_VAL As Double

Dim X_VAL As Double
Dim Y_VAL As Double

Dim C_VAL As Double
Dim NUM_VAL As Double
Dim PIN_VAL As Double
Dim DBL_VAL As Double
Dim PROD_VAL As Double
Dim SGN_VAL As Double

Dim ZERO_VAL As Double
Dim ONE_VAL As Double
Dim HALF_VAL As Double
Dim TWO_VAL As Double

Dim SUM1_VAL As Double
Dim SUM2_VAL As Double
Dim A_ARR() As Double
Dim EIGEN_ARR As Variant

On Error GoTo ERROR_LABEL

ERROR_STR = ""
EIGEN_ARR = DURBIN_WATSON_EIGENVAL_FUNC(DATA_RNG, NROWS)
If IsArray(EIGEN_ARR) = False Then: GoTo ERROR_LABEL

'l is NOT the number of observations
'l is the number summations that are used to compute the integral
'See Farebrother (1980) p. 224-225
'l=200 is taken to be exact
'l=24 is the highest value in the table, but accuracy improves as
'l gets bigger
l = 24
ZERO_VAL = 0#
ONE_VAL = 1#
HALF_VAL = 0.5
TWO_VAL = 2#

'For DW C=0
C_VAL = 0
ReDim A_ARR(0 To DF_VAL)
'Fill eigenvalue vector with DW stat as A_ARR(0)
A_ARR(0) = DW_VAL
X_VAL = A_ARR(0)
For i = 1 To DF_VAL: A_ARR(i) = EIGEN_ARR(i): Next i
'SET N1 = INDEX OF 1ST A_ARR(I) >= X_VAL.
'ALLOW FOR THE A_ARR'S BEING IN REVERSE ORDER.
If (A_ARR(1) > A_ARR(DF_VAL)) Then
    H2 = DF_VAL
    k = -1
    i = 1
Else
    H2 = 1
    k = 1
    i = DF_VAL
End If

For N1 = H2 To i Step k
    If (A_ARR(N1) > X_VAL) Then GoTo 20
Next N1

'IF ALL A_ARR'S ARE -VE AND C >= 0, THEN PROBABILITY = 1.

If (C_VAL > ZERO_VAL) Then
    DURBIN_WATSON_PVALUE_FUNC = ONE_VAL 'Nov 02 Fix
    Exit Function
End If

'SIMILARLY IF ALL THE A_ARR'S ARE +VE AND C <= 0, THEN PROBABILITY = 0.

20:
If (N1 = H2 And C_VAL <= ZERO_VAL) Then
    DURBIN_WATSON_PVALUE_FUNC = ZERO_VAL 'Nov 02 Fix
    Exit Function
End If

If (k = 1) Then N1 = N1 - 1
H2 = DF_VAL - N1
If (C_VAL = ZERO_VAL) Then
    Y_VAL = H2 - N1
Else
    Y_VAL = C_VAL * (A_ARR(1) - A_ARR(DF_VAL))
End If

If (Y_VAL >= ZERO_VAL) Then
    DD = 2
    H2 = N1
    k = -k
    J1 = 0
    J2 = 2
    J3 = 3
    J4 = 1
Else
    DD = -2
    N1 = N1 + 1
    J1 = DF_VAL - 2
    J2 = DF_VAL - 1
    J3 = DF_VAL + 1
    J4 = DF_VAL
End If
PIN_VAL = TWO_VAL * Atn(ONE_VAL) / l
SUM2_VAL = HALF_VAL * (k + 1)
DBL_VAL = l
SGN_VAL = k / DBL_VAL
N2 = l + l - 1

'FIRST INTEGRALS
'In the original Fortran code, the VB translation was
'For L1 = H2 - H2 / 2 * 2 To 0 Step -1
'This was puzzling since there doesn't seem to be any reason for the loop
'The key is if H2 is even or odd.  The code is looking for the element
'in the list of sorted eigenvalues for which the DWStat value is higher than
'If the element is even, all is well and you go through the loop once.
'If it is odd, you go through twice.
'This is incorporated in the code by DIMming H1 as an Integer.

H1 = H2 / 2
For L1 = 2 * H1 - H2 To 0 Step -1
    For L2 = J2 To N1 Step DD
        SUM1_VAL = A_ARR(J4)
        ' PROD_VAL = A_ARR(J2)
        PROD_VAL = A_ARR(L2)
        U_VAL = HALF_VAL * (SUM1_VAL + PROD_VAL)
        V_VAL = HALF_VAL * (SUM1_VAL - PROD_VAL)
        SUM1_VAL = ZERO_VAL
        For i = 1 To N2 Step 2
            'MsgBox "I is " & I & "N2 is " & N2
            Y_VAL = U_VAL - V_VAL * Cos(i * PIN_VAL)
            NUM_VAL = Y_VAL - X_VAL
            PROD_VAL = Exp(-C_VAL / NUM_VAL)
            For k = 1 To J1
                'MsgBox "Num is " & NUM_VAL & "C is " & C & "Y is " & Y_VAL & "A(K) is " & A_ARR(K)
                PROD_VAL = PROD_VAL * NUM_VAL / (Y_VAL - A_ARR(k))
                'If K = 750 Then
                    'MsgBox "hi"
                'End If
Break:
            Next k
            For k = J3 To DF_VAL
                PROD_VAL = PROD_VAL * NUM_VAL / (Y_VAL - A_ARR(k))
            Next k
            SUM1_VAL = SUM1_VAL + Sqr(Abs(PROD_VAL))
        Next i
        SGN_VAL = -SGN_VAL
        SUM2_VAL = SUM2_VAL + SGN_VAL * SUM1_VAL
        J1 = J1 + DD
        J3 = J3 + DD
        J4 = J4 + DD
    Next L2
    'SECOND INTEGRAL.
    If (DD = 2) Then
        J3 = J3 - 1
    Else
        J1 = J1 + 1
    End If
    J2 = 0
    N1 = 0
Next L1

'When DF is large this algorithm may fail:
'The Imtest Package says this:
'The p value is computed using a Fortran version of the Applied Statistics Algorithm AS 153 by Farebrother (1980, 1984).
'This algorithm is called "pan" or "gradsol". For large sample sizes
'the algorithm might fail tocompute the p value; in that case a warning
'is printed and an approximate p value will be given; this p value
'is computed using a normal approximation with mean and variance of
'the Durbin-Watson test statistic.
'We will adopt this procedure.
'The algorithm "fails" if SUM2_VAL > 1.

'You can see what's going if you put a break in the line that says Break:
'and run the code.
'View the Locals Window and watch what happens to Prod
'When DF is high, the cumulation of many terms makes Prod explode
'This does not happen when DF is low
'Nov 02: I do not know what controls Prod to explode

If SUM2_VAL > 1 Then
'Report warning to user
    ERROR_STR = "The Durbin-Watson P Value calculation cannot be completed.  This occurs when l is high.  A normal approximation for the P Value will be computed and reported." ', vbInformation, "Can't Compute Exact DW P Value"
'Source is original Durbin Watson article (I), p. 420
'Mean is trace MA which is back in the MatrixMult macro so we'll go back up
' there to compute it.
'I'll make the sum=2 and then check it in MatrixMult
    SUM2_VAL = 2
End If

DURBIN_WATSON_PVALUE_FUNC = SUM2_VAL

Exit Function
ERROR_LABEL:
ERROR_STR = Err.number
DURBIN_WATSON_PVALUE_FUNC = ERROR_STR
End Function

'// PERFECT
' Get the eigenvalues of a real symmetric matrix, MAM
Private Function DURBIN_WATSON_EIGENVAL_FUNC(ByRef DATA_RNG As Variant, _
ByVal NROWS As Integer)

' Get the machine-dependent constants down

Dim h As Integer
Dim i As Integer
Dim j As Integer
Dim k As Integer
Dim l As Integer
Dim m As Integer
Dim n As Integer
Dim o As Integer
Dim p As Integer
Dim q As Integer

Dim B_VAL As Double
Dim C_VAL As Double
Dim G_VAL As Double
Dim F_VAL As Double
Dim H_VAL As Double
Dim I_VAL As Double

Dim P_VAL As Double
Dim Q_VAL As Double
Dim R_VAL As Double
Dim S_VAL As Double

Dim MITS_VAL As Double
Dim ZERO_VAL As Double
Dim ONE_VAL As Double
Dim TWO_VAL As Double

Dim E_ARR() As Double
' D is the eigenvalue vector
Dim D_ARR() As Double
Dim DATA_MATRIX As Variant

Dim epsilon As Double
Dim tolerance As Double
Dim ETA_VAL As Double
Dim PRECISION_VAL As Double

On Error GoTo ERROR_LABEL

DATA_MATRIX = DATA_RNG
epsilon = 7.105427358E-15
tolerance = 3.131513063E-294
PRECISION_VAL = 0.00000000000001
ETA_VAL = epsilon * tolerance

i = NROWS
ReDim E_ARR(1 To NROWS)
ReDim D_ARR(1 To NROWS)

' NOTE: This code is essentially the same as that in
' AS algorithms documented above. Please see these
' algorithms for more detailed documentation.

For h = 2 To NROWS
    l = i - 2
    F_VAL = DATA_MATRIX(i, i - 1)
    G_VAL = 0
    If (l < 1) Then GoTo 30
    For k = 1 To l
        G_VAL = G_VAL + DATA_MATRIX(i, k) ^ 2
    Next k
30:
    H_VAL = G_VAL + F_VAL * F_VAL
    
    'if G_VAL is too small for orthogonality to be guaranteed, the transformation Is skipped
    If (G_VAL > tolerance) Then GoTo 40
    E_ARR(i) = F_VAL
    
    D_ARR(i) = 0
    GoTo 100
40:
    l = l + 1
    G_VAL = Sqr(H_VAL)
    If (F_VAL >= 0) Then G_VAL = -G_VAL
    E_ARR(i) = G_VAL
    H_VAL = H_VAL - F_VAL * G_VAL
    DATA_MATRIX(i, i - 1) = F_VAL - G_VAL
    F_VAL = 0
    For j = 1 To l
        DATA_MATRIX(j, i) = DATA_MATRIX(i, j) / H_VAL
        G_VAL = 0
        'form element of a * u
        For k = 1 To j
            G_VAL = G_VAL + DATA_MATRIX(j, k) * DATA_MATRIX(i, k)
        Next k
        If (j >= l) Then GoTo 70
        q = j + 1
        For k = q To l
            G_VAL = G_VAL + DATA_MATRIX(k, j) * DATA_MATRIX(i, k)
        Next k
        'form element of P_VAL
70:
        E_ARR(j) = G_VAL / H_VAL
        F_VAL = F_VAL + G_VAL * DATA_MATRIX(j, i)
    Next j
    'form k
    I_VAL = F_VAL / (H_VAL + H_VAL)
    'form reduced a
    For j = 1 To l
        F_VAL = DATA_MATRIX(i, j)
        G_VAL = E_ARR(j) - I_VAL * F_VAL
        E_ARR(j) = G_VAL
        For k = 1 To j
            DATA_MATRIX(j, k) = DATA_MATRIX(j, k) - F_VAL * E_ARR(k) - G_VAL * DATA_MATRIX(i, k)
            'MsgBox "DATA_MATRIX(J,K) is " & DATA_MATRIX(J, K)
        Next k
    Next j
    '90      CONTINUE
    D_ARR(i) = H_VAL
100:
    i = i - 1
Next h

D_ARR(1) = 0
E_ARR(1) = 0

'accumulation of transformation matrices
For i = 1 To NROWS
    l = i - 1
    If (D_ARR(i) = 0 Or l = 0) Then GoTo 140
    For j = 1 To l
        G_VAL = 0
        For k = 1 To l
            G_VAL = G_VAL + DATA_MATRIX(i, k) * DATA_MATRIX(k, j)
        Next k
        For k = 1 To l
            DATA_MATRIX(k, j) = DATA_MATRIX(k, j) - G_VAL * DATA_MATRIX(k, i)
        Next k
    Next j
140:
    D_ARR(i) = DATA_MATRIX(i, i)
    DATA_MATRIX(i, i) = 1
    If (l = 0) Then GoTo 1983
    For j = 1 To l
        DATA_MATRIX(i, j) = 0
        DATA_MATRIX(j, i) = 0
    Next j
1983:
Next i

MITS_VAL = 30
ZERO_VAL = 0#
ONE_VAL = 1#
TWO_VAL = 2#
PRECISION_VAL = 0.00000000000001

o = NROWS - 1
For i = 2 To NROWS
    E_ARR(i - 1) = E_ARR(i)
Next i
E_ARR(NROWS) = ZERO_VAL
B_VAL = ZERO_VAL
F_VAL = ZERO_VAL

For l = 1 To NROWS
    p = 0
    H_VAL = PRECISION_VAL * (Abs(D_ARR(l)) + Abs(E_ARR(l)))
    If (B_VAL < H_VAL) Then B_VAL = H_VAL
    'look for small sub-diagonal element
    For n = l To NROWS
        m = n
        If (Abs(E_ARR(m)) <= B_VAL) Then GoTo 1984
    Next n
1984:
    If (m = l) Then GoTo 90
1985:
    If (p = MITS_VAL) Then GoTo ERROR_LABEL 'Exit Function
    p = p + 1
    'form shift
    P_VAL = (D_ARR(l + 1) - D_ARR(l)) / (TWO_VAL * E_ARR(l))
    On Error Resume Next
    R_VAL = Sqr(P_VAL * P_VAL + ONE_VAL)
    
    If Err.number = 6 Then P_VAL = 0
    Q_VAL = P_VAL + R_VAL
    If (P_VAL < ZERO_VAL) Then Q_VAL = P_VAL - R_VAL
    H_VAL = D_ARR(l) - E_ARR(l) / Q_VAL
    For i = l To NROWS
        D_ARR(i) = D_ARR(i) - H_VAL
    Next i
    F_VAL = F_VAL + H_VAL
    'ql transformation
    P_VAL = D_ARR(m)
    C_VAL = ONE_VAL
    S_VAL = ZERO_VAL
    n = m - 1
    i = m
    For h = l To n
        j = i
        i = i - 1
        G_VAL = C_VAL * E_ARR(i)
        H_VAL = C_VAL * P_VAL
        If (Abs(P_VAL) >= Abs(E_ARR(i))) Then GoTo 60
        C_VAL = P_VAL / E_ARR(i)
        R_VAL = Sqr(C_VAL * C_VAL + ONE_VAL)
        E_ARR(j) = S_VAL * E_ARR(i) * R_VAL
        S_VAL = ONE_VAL / R_VAL
        C_VAL = C_VAL / R_VAL
        GoTo 1986
60:
        C_VAL = E_ARR(i) / P_VAL
        R_VAL = Sqr(C_VAL * C_VAL + ONE_VAL)
        E_ARR(j) = S_VAL * P_VAL * R_VAL
        S_VAL = C_VAL / R_VAL
        C_VAL = ONE_VAL / R_VAL
1986:
        P_VAL = C_VAL * D_ARR(i) - S_VAL * G_VAL
        D_ARR(j) = H_VAL + S_VAL * (C_VAL * G_VAL + S_VAL * D_ARR(i))
        'form vector
        For k = 1 To NROWS
            H_VAL = DATA_MATRIX(k, j)
            DATA_MATRIX(k, j) = S_VAL * DATA_MATRIX(k, i) + C_VAL * H_VAL
            DATA_MATRIX(k, i) = C_VAL * DATA_MATRIX(k, i) - S_VAL * H_VAL
        Next k
    Next h
    
    E_ARR(l) = S_VAL * P_VAL
    D_ARR(l) = C_VAL * P_VAL
    If (Abs(E_ARR(l)) > B_VAL) Then GoTo 1985
90:
    D_ARR(l) = D_ARR(l) + F_VAL
Next l

'order latent roots and vectors
For i = 1 To o
    k = i
    P_VAL = D_ARR(i)
    h = i + 1
    For j = h To NROWS
        If (D_ARR(j) <= P_VAL) Then GoTo 110
        k = j
        P_VAL = D_ARR(j)
110:
    Next j
    If (k = i) Then GoTo 130
    D_ARR(k) = D_ARR(i)
    D_ARR(i) = P_VAL
    'MsgBox D_ARR(I)
    For j = 1 To NROWS
        P_VAL = DATA_MATRIX(j, i)
        DATA_MATRIX(j, i) = DATA_MATRIX(j, k)
        DATA_MATRIX(j, k) = P_VAL
    Next j
    'MsgBox D_ARR(I)
130:
Next i

' Assigns the eigenvalues to the D_ARR() vector
DURBIN_WATSON_EIGENVAL_FUNC = D_ARR

Exit Function
ERROR_LABEL:
DURBIN_WATSON_EIGENVAL_FUNC = Err.number
End Function
