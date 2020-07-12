Attribute VB_Name = "FINAN_FI_BOND_RISKY_LIBR"
'------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------
'Longstaff & Schwartz 1995 Valuing Risky Debt
'Reference: Longstaff, Francis A., Schwartz, Eduardo S. A simple
'approach to Valuing Risky Fixed and Floating Rate Debt. Journal of
'Finance. Vol 3. July 1995.
'------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.
                    

'Risky Debt Valuation Model Longstaff & Schwartz 95

Public Function LS_RISKY_DEBT_VALUATION_FUNC(ByRef RISK_FREE_RATE_RNG As Variant, _
ByRef TENOR_RNG As Variant, _
ByRef BETA_RNG As Variant, _
ByRef H2_RNG As Variant, _
ByRef ALPHA_RNG As Variant, _
ByRef X_RNG As Variant, _
ByRef W_RNG As Variant, _
ByRef SIGMA2_RNG As Variant, _
ByRef RHO_RNG As Variant, _
Optional ByVal nLOOPS_RNG As Variant = 100, _
Optional ByVal CND_TYPE_RNG As Variant = 0, _
Optional ByVal VERSION As Integer = 0)

Dim i As Long
Dim j As Long
Dim NROWS As Long
Dim NCOLUMNS As Long

Dim ATEMP_VAL As Double
Dim BTEMP_VAL As Double
Dim CTEMP_VAL As Double

Dim RISK_FREE_RATE As Double 'Rate r0 at t=0
Dim TENOR As Double 'Maturity Time(years)
Dim BETA_VAL As Double '"Pullback"
Dim H2_VAL As Double 'Instantaneous Variance of short rate
Dim ALPHA_VAL As Double 'a in L&S = z + constant
Dim X_VAL As Double 'V/K =X
Dim W_VAL As Double 'Writedown = 1 - Recovery Rate
Dim SIGMA2_VAL As Double 'Volatility of asset value process
Dim RHO_VAL As Double 'Instantaneous correl. Asset/interest rate
Dim nLOOPS As Long 'Iterations for Q
Dim CND_TYPE As Integer

Dim RISK_FREE_RATE_VECTOR As Variant
Dim TENOR_VECTOR As Variant
Dim BETA_VECTOR As Variant
Dim H2_VECTOR As Variant
Dim ALPHA_VECTOR As Variant
Dim X_VECTOR As Variant
Dim W_VECTOR As Variant
Dim SIGMA2_VECTOR As Variant
Dim RHO_VECTOR As Variant
Dim nLOOPS_VECTOR As Variant
Dim CND_TYPE_VECTOR As Variant

Dim TEMP_MATRIX As Variant

On Error GoTo ERROR_LABEL

If IsArray(RISK_FREE_RATE_RNG) = True Then
    RISK_FREE_RATE_VECTOR = RISK_FREE_RATE_RNG
    If UBound(RISK_FREE_RATE_VECTOR, 2) = 1 Then
        RISK_FREE_RATE_VECTOR = MATRIX_TRANSPOSE_FUNC(RISK_FREE_RATE_VECTOR)
    End If
Else
    ReDim RISK_FREE_RATE_VECTOR(1 To 1, 1 To 1)
    RISK_FREE_RATE_VECTOR(1, 1) = RISK_FREE_RATE_RNG
End If
NCOLUMNS = UBound(RISK_FREE_RATE_VECTOR, 2)


If IsArray(TENOR_RNG) = True Then
    TENOR_VECTOR = TENOR_RNG
    If VERSION = 0 Then
        If UBound(TENOR_VECTOR, 2) = 1 Then
            TENOR_VECTOR = MATRIX_TRANSPOSE_FUNC(TENOR_VECTOR)
        End If
    Else
        If UBound(TENOR_VECTOR, 1) = 1 Then
            TENOR_VECTOR = MATRIX_TRANSPOSE_FUNC(TENOR_VECTOR)
        End If
    End If
Else
    ReDim TENOR_VECTOR(1 To 1, 1 To 1)
    TENOR_VECTOR(1, 1) = TENOR_RNG
End If

If VERSION = 0 Then
    If UBound(TENOR_VECTOR, 2) <> NCOLUMNS Then: GoTo ERROR_LABEL
End If

If IsArray(BETA_RNG) = True Then
    BETA_VECTOR = BETA_RNG
    If UBound(BETA_VECTOR, 2) = 1 Then
        BETA_VECTOR = MATRIX_TRANSPOSE_FUNC(BETA_VECTOR)
    End If
Else
    ReDim BETA_VECTOR(1 To 1, 1 To 1)
    BETA_VECTOR(1, 1) = BETA_RNG
End If
If UBound(BETA_VECTOR, 2) <> NCOLUMNS Then: GoTo ERROR_LABEL


If IsArray(H2_RNG) = True Then
    H2_VECTOR = H2_RNG
    If UBound(H2_VECTOR, 2) = 1 Then
        H2_VECTOR = MATRIX_TRANSPOSE_FUNC(H2_VECTOR)
    End If
Else
    ReDim H2_VECTOR(1 To 1, 1 To 1)
    H2_VECTOR(1, 1) = H2_RNG
End If
If UBound(H2_VECTOR, 2) <> NCOLUMNS Then: GoTo ERROR_LABEL


If IsArray(ALPHA_RNG) = True Then
    ALPHA_VECTOR = ALPHA_RNG
    If UBound(ALPHA_VECTOR, 2) = 1 Then
        ALPHA_VECTOR = MATRIX_TRANSPOSE_FUNC(ALPHA_VECTOR)
    End If
Else
    ReDim ALPHA_VECTOR(1 To 1, 1 To 1)
    ALPHA_VECTOR(1, 1) = ALPHA_RNG
End If
If UBound(ALPHA_VECTOR, 2) <> NCOLUMNS Then: GoTo ERROR_LABEL


If IsArray(X_RNG) = True Then
    X_VECTOR = X_RNG
    If UBound(X_VECTOR, 2) = 1 Then
        X_VECTOR = MATRIX_TRANSPOSE_FUNC(X_VECTOR)
    End If
Else
    ReDim X_VECTOR(1 To 1, 1 To 1)
    X_VECTOR(1, 1) = X_RNG
End If
If UBound(X_VECTOR, 2) <> NCOLUMNS Then: GoTo ERROR_LABEL


If IsArray(W_RNG) = True Then
    W_VECTOR = W_RNG
    If UBound(W_VECTOR, 2) = 1 Then
        W_VECTOR = MATRIX_TRANSPOSE_FUNC(W_VECTOR)
    End If
Else
    ReDim W_VECTOR(1 To 1, 1 To 1)
    W_VECTOR(1, 1) = W_RNG
End If
If UBound(W_VECTOR, 2) <> NCOLUMNS Then: GoTo ERROR_LABEL


If IsArray(SIGMA2_RNG) = True Then
    SIGMA2_VECTOR = SIGMA2_RNG
    If UBound(SIGMA2_VECTOR, 2) = 1 Then
        SIGMA2_VECTOR = MATRIX_TRANSPOSE_FUNC(SIGMA2_VECTOR)
    End If
Else
    ReDim SIGMA2_VECTOR(1 To 1, 1 To 1)
    SIGMA2_VECTOR(1, 1) = SIGMA2_RNG
End If
If UBound(SIGMA2_VECTOR, 2) <> NCOLUMNS Then: GoTo ERROR_LABEL


If IsArray(RHO_RNG) = True Then
    RHO_VECTOR = RHO_RNG
    If UBound(RHO_VECTOR, 2) = 1 Then
        RHO_VECTOR = MATRIX_TRANSPOSE_FUNC(RHO_VECTOR)
    End If
Else
    ReDim RHO_VECTOR(1 To 1, 1 To 1)
    RHO_VECTOR(1, 1) = RHO_RNG
End If
If UBound(RHO_VECTOR, 2) <> NCOLUMNS Then: GoTo ERROR_LABEL


If IsArray(nLOOPS_RNG) = True Then
    nLOOPS_VECTOR = nLOOPS_RNG
    If UBound(nLOOPS_VECTOR, 2) = 1 Then
        nLOOPS_VECTOR = MATRIX_TRANSPOSE_FUNC(nLOOPS_VECTOR)
    End If
Else
    ReDim nLOOPS_VECTOR(1 To 1, 1 To NCOLUMNS)
    For j = 1 To NCOLUMNS
        nLOOPS_VECTOR(1, j) = nLOOPS_RNG
    Next j
End If
If UBound(nLOOPS_VECTOR, 2) <> NCOLUMNS Then: GoTo ERROR_LABEL


If IsArray(CND_TYPE_RNG) = True Then
    CND_TYPE_VECTOR = CND_TYPE_RNG
    If UBound(CND_TYPE_VECTOR, 2) = 1 Then
        CND_TYPE_VECTOR = MATRIX_TRANSPOSE_FUNC(CND_TYPE_VECTOR)
    End If
Else
    ReDim CND_TYPE_VECTOR(1 To 1, 1 To NCOLUMNS)
    For j = 1 To NCOLUMNS
        CND_TYPE_VECTOR(1, j) = CND_TYPE_RNG
    Next j
End If
If UBound(CND_TYPE_VECTOR, 2) <> NCOLUMNS Then: GoTo ERROR_LABEL

'-----------------------------------------------------------------------------------------------
If VERSION = 0 Then
'-----------------------------------------------------------------------------------------------
    ReDim TEMP_MATRIX(1 To 9, 1 To NCOLUMNS + 1)
    
    TEMP_MATRIX(1, 1) = _
    "RISKY DISCOUNT BOND: VALUE RISK FREE DISCOUNT BOND (VASICEK)"
    
    TEMP_MATRIX(2, 1) = _
    "RISKY DISCOUNT BOND: YIELD RISKFREE BOND"
    
    TEMP_MATRIX(3, 1) = _
    "RISKY DISCOUNT BOND: PROBABILITY OF DEFAULT (RISK NEUTRAL)"
    
    TEMP_MATRIX(4, 1) = _
    "RISKY DISCOUNT BOND: VALUE RISKY DISCOUNT BOND (L&S 95)"
    
    TEMP_MATRIX(5, 1) = _
    "RISKY DISCOUNT BOND: YIELD RISKY DISCOUNT BOND"
    
    TEMP_MATRIX(6, 1) = _
    "FLOATING-RATE COUPON PAYMENT: TIME OF FLOATING RATE PAYMENT (<=T)"
    
    TEMP_MATRIX(7, 1) = _
    "FLOATING-RATE COUPON PAYMENT: EXPECTED VALUE R AT T (RISK NEUTRAL PROCESS)"
    
    TEMP_MATRIX(8, 1) = _
    "FLOATING-RATE COUPON PAYMENT: CORRELATION ADJUSTMENT"
    
    TEMP_MATRIX(9, 1) = _
    "FLOATING-RATE COUPON PAYMENT: VALUE OF FLOATING RATE PAYMENT AT TIME TAU"
    
    For j = 1 To NCOLUMNS
        RISK_FREE_RATE = RISK_FREE_RATE_VECTOR(1, j)
        TENOR = TENOR_VECTOR(1, j)
        BETA_VAL = BETA_VECTOR(1, j)
        H2_VAL = H2_VECTOR(1, j)
        
        ALPHA_VAL = ALPHA_VECTOR(1, j)
        X_VAL = X_VECTOR(1, j)
        W_VAL = W_VECTOR(1, j)
        SIGMA2_VAL = SIGMA2_VECTOR(1, j)
        RHO_VAL = RHO_VECTOR(1, j)
        nLOOPS = nLOOPS_VECTOR(1, j)
        CND_TYPE = CND_TYPE_VECTOR(1, j)
        
        TEMP_MATRIX(1, j + 1) = LS_D_FUNC(RISK_FREE_RATE, TENOR, ALPHA_VAL, BETA_VAL, H2_VAL)
        TEMP_MATRIX(2, j + 1) = Log(1 / TEMP_MATRIX(1, j + 1)) / TENOR
        'Log(TEMP_MATRIX(1, j + 1)) / TENOR
        
        TEMP_MATRIX(3, j + 1) = LS_Q_FUNC(X_VAL, RISK_FREE_RATE, TENOR, ALPHA_VAL, _
                                BETA_VAL, SIGMA2_VAL ^ 0.5, SIGMA2_VAL, H2_VAL ^ 0.5, _
                                H2_VAL, RHO_VAL, nLOOPS, CND_TYPE, 1)
        
        TEMP_MATRIX(4, j + 1) = LS_P_FUNC(TEMP_MATRIX(1, j + 1), W_VAL, TEMP_MATRIX(3, j + 1))
        TEMP_MATRIX(5, j + 1) = Log(1 / TEMP_MATRIX(4, j + 1)) / TENOR
        TEMP_MATRIX(6, j + 1) = TENOR
        
        TEMP_MATRIX(7, j + 1) = LS_R_EXP_FUNC(RISK_FREE_RATE, TEMP_MATRIX(6, j + 1), TENOR, _
                                ALPHA_VAL, BETA_VAL, H2_VAL)
        
        TEMP_MATRIX(8, j + 1) = LS_G_FUNC(X_VAL, RISK_FREE_RATE, TEMP_MATRIX(6, j + 1), TENOR, nLOOPS, _
                                ALPHA_VAL, BETA_VAL, SIGMA2_VAL ^ 0.5, SIGMA2_VAL, H2_VAL ^ 0.5, _
                                H2_VAL, RHO_VAL, CND_TYPE)
        
        TEMP_MATRIX(9, j + 1) = LS_F_FUNC(W_VAL, TEMP_MATRIX(1, j + 1), TEMP_MATRIX(4, j + 1), _
                                TEMP_MATRIX(7, j + 1), TEMP_MATRIX(8, j + 1))
    Next j
'-----------------------------------------------------------------------------------------------
Else
'-----------------------------------------------------------------------------------------------
    NROWS = UBound(TENOR_VECTOR, 1)
    
    ReDim TEMP_MATRIX(0 To NROWS, 1 To 6)
    TEMP_MATRIX(0, 1) = "TN"
    
    TEMP_MATRIX(0, 2) = _
    "RISKY DISCOUNT BOND: VALUE RISK FREE DISCOUNT BOND (VASICEK)"
    
    TEMP_MATRIX(0, 3) = _
    "RISKY DISCOUNT BOND: RATE RISKFREE BOND"
    
    TEMP_MATRIX(0, 4) = _
    "RISKY DISCOUNT BOND: VALUE RISKY DISCOUNT BOND (L&S 95)"
    
    TEMP_MATRIX(0, 5) = _
    "RISKY DISCOUNT BOND: RATE RISKY DISCOUNT BOND"
    
    TEMP_MATRIX(0, 6) = _
    "FLOATING-RATE COUPON PAYMENT: VALUE OF FLOATING RATE PAYMENT AT TIME TAU"
    
    For i = 1 To NROWS
        TENOR = TENOR_VECTOR(i, 1)
        
        RISK_FREE_RATE = RISK_FREE_RATE_VECTOR(1, 1)
        BETA_VAL = BETA_VECTOR(1, 1)
        H2_VAL = H2_VECTOR(1, 1)
        
        ALPHA_VAL = ALPHA_VECTOR(1, 1)
        X_VAL = X_VECTOR(1, 1)
        W_VAL = W_VECTOR(1, 1)
        SIGMA2_VAL = SIGMA2_VECTOR(1, 1)
        RHO_VAL = RHO_VECTOR(1, 1)
        nLOOPS = nLOOPS_VECTOR(1, 1)
        CND_TYPE = CND_TYPE_VECTOR(1, 1)
    
        TEMP_MATRIX(i, 1) = TENOR
        
        TEMP_MATRIX(i, 2) = LS_D_FUNC(RISK_FREE_RATE, TENOR, ALPHA_VAL, BETA_VAL, H2_VAL)
        TEMP_MATRIX(i, 3) = Log(1 / TEMP_MATRIX(i, 2)) / (TEMP_MATRIX(i, 1))
        
        ATEMP_VAL = LS_Q_FUNC(X_VAL, RISK_FREE_RATE, TENOR, ALPHA_VAL, _
                                BETA_VAL, SIGMA2_VAL ^ 0.5, SIGMA2_VAL, H2_VAL ^ 0.5, _
                                H2_VAL, RHO_VAL, nLOOPS, CND_TYPE, 1)
        
        TEMP_MATRIX(i, 4) = LS_P_FUNC(TEMP_MATRIX(i, 2), W_VAL, ATEMP_VAL)
        TEMP_MATRIX(i, 5) = Log(1 / TEMP_MATRIX(i, 4)) / (TEMP_MATRIX(i, 1))
        
        
        BTEMP_VAL = LS_R_EXP_FUNC(RISK_FREE_RATE, TEMP_MATRIX(i, 1), TENOR, _
                                ALPHA_VAL, BETA_VAL, H2_VAL)
        
        CTEMP_VAL = LS_G_FUNC(X_VAL, RISK_FREE_RATE, TEMP_MATRIX(i, 1), TENOR, nLOOPS, _
                                ALPHA_VAL, BETA_VAL, SIGMA2_VAL ^ 0.5, SIGMA2_VAL, H2_VAL ^ 0.5, _
                                H2_VAL, RHO_VAL, CND_TYPE)
        
        TEMP_MATRIX(i, 6) = LS_F_FUNC(W_VAL, TEMP_MATRIX(i, 2), TEMP_MATRIX(i, 4), _
                                BTEMP_VAL, CTEMP_VAL)

    
    Next i
'-----------------------------------------------------------------------------------------------
End If
'-----------------------------------------------------------------------------------------------

LS_RISKY_DEBT_VALUATION_FUNC = TEMP_MATRIX

Exit Function
ERROR_LABEL:
LS_RISKY_DEBT_VALUATION_FUNC = Err.number
End Function

'Risky Debt Valuation Model Longstaff & Schwartz 95

Public Function LS_RISKY_COUPON_DEBT_VALUATION_FUNC(ByRef RISK_FREE_RATE_RNG As Variant, _
ByRef TENOR_RNG As Variant, _
ByRef BETA_RNG As Variant, _
ByRef H2_RNG As Variant, _
ByRef ALPHA_RNG As Variant, _
ByRef X_RNG As Variant, _
ByRef W_RNG As Variant, _
ByRef SIGMA2_RNG As Variant, _
ByRef RHO_RNG As Variant, _
ByRef FIXED_COUPON_RNG As Variant, _
Optional ByVal nLOOPS_RNG As Variant = 100, _
Optional ByVal CND_TYPE_RNG As Variant = 0, _
Optional ByVal VERSION As Integer = 0)

Dim i As Long
Dim j As Long
Dim NROWS As Long
Dim NCOLUMNS As Long

Dim RISK_FREE_RATE As Double 'Rate r0 at t=0
Dim TENOR As Double 'Maturity Time(years)
Dim BETA_VAL As Double '"Pullback"
Dim H2_VAL As Double 'Instantaneous Variance of short rate
Dim ALPHA_VAL As Double 'a in L&S = z + constant
Dim X_VAL As Double 'V/K =X
Dim W_VAL As Double 'Writedown = 1 - Recovery Rate
Dim SIGMA2_VAL As Double 'Volatility of asset value process
Dim RHO_VAL As Double 'Instantaneous correl. Asset/interest rate
Dim FIXED_COUPON_VAL As Double
Dim nLOOPS As Long 'Iterations for Q
Dim CND_TYPE As Integer

Dim RISK_FREE_RATE_VECTOR As Variant
Dim TENOR_VECTOR As Variant
Dim BETA_VECTOR As Variant
Dim H2_VECTOR As Variant
Dim ALPHA_VECTOR As Variant
Dim X_VECTOR As Variant
Dim W_VECTOR As Variant
Dim SIGMA2_VECTOR As Variant
Dim RHO_VECTOR As Variant
Dim FIXED_COUPON_VECTOR As Variant
Dim nLOOPS_VECTOR As Variant
Dim CND_TYPE_VECTOR As Variant

Dim TEMP_MATRIX As Variant

On Error GoTo ERROR_LABEL

If IsArray(RISK_FREE_RATE_RNG) = True Then
    RISK_FREE_RATE_VECTOR = RISK_FREE_RATE_RNG
    If UBound(RISK_FREE_RATE_VECTOR, 2) = 1 Then
        RISK_FREE_RATE_VECTOR = MATRIX_TRANSPOSE_FUNC(RISK_FREE_RATE_VECTOR)
    End If
Else
    ReDim RISK_FREE_RATE_VECTOR(1 To 1, 1 To 1)
    RISK_FREE_RATE_VECTOR(1, 1) = RISK_FREE_RATE_RNG
End If
NCOLUMNS = UBound(RISK_FREE_RATE_VECTOR, 2)


If IsArray(TENOR_RNG) = True Then
    TENOR_VECTOR = TENOR_RNG
    If VERSION = 0 Then
        If UBound(TENOR_VECTOR, 2) = 1 Then
            TENOR_VECTOR = MATRIX_TRANSPOSE_FUNC(TENOR_VECTOR)
        End If
    Else
        If UBound(TENOR_VECTOR, 1) = 1 Then
            TENOR_VECTOR = MATRIX_TRANSPOSE_FUNC(TENOR_VECTOR)
        End If
    End If
Else
    ReDim TENOR_VECTOR(1 To 1, 1 To 1)
    TENOR_VECTOR(1, 1) = TENOR_RNG
End If

If VERSION = 0 Then
    If UBound(TENOR_VECTOR, 2) <> NCOLUMNS Then: GoTo ERROR_LABEL
End If

If IsArray(BETA_RNG) = True Then
    BETA_VECTOR = BETA_RNG
    If UBound(BETA_VECTOR, 2) = 1 Then
        BETA_VECTOR = MATRIX_TRANSPOSE_FUNC(BETA_VECTOR)
    End If
Else
    ReDim BETA_VECTOR(1 To 1, 1 To 1)
    BETA_VECTOR(1, 1) = BETA_RNG
End If
If UBound(BETA_VECTOR, 2) <> NCOLUMNS Then: GoTo ERROR_LABEL


If IsArray(H2_RNG) = True Then
    H2_VECTOR = H2_RNG
    If UBound(H2_VECTOR, 2) = 1 Then
        H2_VECTOR = MATRIX_TRANSPOSE_FUNC(H2_VECTOR)
    End If
Else
    ReDim H2_VECTOR(1 To 1, 1 To 1)
    H2_VECTOR(1, 1) = H2_RNG
End If
If UBound(H2_VECTOR, 2) <> NCOLUMNS Then: GoTo ERROR_LABEL


If IsArray(ALPHA_RNG) = True Then
    ALPHA_VECTOR = ALPHA_RNG
    If UBound(ALPHA_VECTOR, 2) = 1 Then
        ALPHA_VECTOR = MATRIX_TRANSPOSE_FUNC(ALPHA_VECTOR)
    End If
Else
    ReDim ALPHA_VECTOR(1 To 1, 1 To 1)
    ALPHA_VECTOR(1, 1) = ALPHA_RNG
End If
If UBound(ALPHA_VECTOR, 2) <> NCOLUMNS Then: GoTo ERROR_LABEL


If IsArray(X_RNG) = True Then
    X_VECTOR = X_RNG
    If UBound(X_VECTOR, 2) = 1 Then
        X_VECTOR = MATRIX_TRANSPOSE_FUNC(X_VECTOR)
    End If
Else
    ReDim X_VECTOR(1 To 1, 1 To 1)
    X_VECTOR(1, 1) = X_RNG
End If
If UBound(X_VECTOR, 2) <> NCOLUMNS Then: GoTo ERROR_LABEL


If IsArray(W_RNG) = True Then
    W_VECTOR = W_RNG
    If UBound(W_VECTOR, 2) = 1 Then
        W_VECTOR = MATRIX_TRANSPOSE_FUNC(W_VECTOR)
    End If
Else
    ReDim W_VECTOR(1 To 1, 1 To 1)
    W_VECTOR(1, 1) = W_RNG
End If
If UBound(W_VECTOR, 2) <> NCOLUMNS Then: GoTo ERROR_LABEL


If IsArray(SIGMA2_RNG) = True Then
    SIGMA2_VECTOR = SIGMA2_RNG
    If UBound(SIGMA2_VECTOR, 2) = 1 Then
        SIGMA2_VECTOR = MATRIX_TRANSPOSE_FUNC(SIGMA2_VECTOR)
    End If
Else
    ReDim SIGMA2_VECTOR(1 To 1, 1 To 1)
    SIGMA2_VECTOR(1, 1) = SIGMA2_RNG
End If
If UBound(SIGMA2_VECTOR, 2) <> NCOLUMNS Then: GoTo ERROR_LABEL


If IsArray(RHO_RNG) = True Then
    RHO_VECTOR = RHO_RNG
    If UBound(RHO_VECTOR, 2) = 1 Then
        RHO_VECTOR = MATRIX_TRANSPOSE_FUNC(RHO_VECTOR)
    End If
Else
    ReDim RHO_VECTOR(1 To 1, 1 To 1)
    RHO_VECTOR(1, 1) = RHO_RNG
End If
If UBound(RHO_VECTOR, 2) <> NCOLUMNS Then: GoTo ERROR_LABEL


If IsArray(FIXED_COUPON_RNG) = True Then
    FIXED_COUPON_VECTOR = FIXED_COUPON_RNG
    If UBound(FIXED_COUPON_VECTOR, 2) = 1 Then
        FIXED_COUPON_VECTOR = MATRIX_TRANSPOSE_FUNC(FIXED_COUPON_VECTOR)
    End If
Else
    ReDim FIXED_COUPON_VECTOR(1 To 1, 1 To 1)
    FIXED_COUPON_VECTOR(1, 1) = FIXED_COUPON_RNG
End If
If UBound(FIXED_COUPON_VECTOR, 2) <> NCOLUMNS Then: GoTo ERROR_LABEL

If IsArray(nLOOPS_RNG) = True Then
    nLOOPS_VECTOR = nLOOPS_RNG
    If UBound(nLOOPS_VECTOR, 2) = 1 Then
        nLOOPS_VECTOR = MATRIX_TRANSPOSE_FUNC(nLOOPS_VECTOR)
    End If
Else
    ReDim nLOOPS_VECTOR(1 To 1, 1 To NCOLUMNS)
    For j = 1 To NCOLUMNS
        nLOOPS_VECTOR(1, j) = nLOOPS_RNG
    Next j
End If
If UBound(nLOOPS_VECTOR, 2) <> NCOLUMNS Then: GoTo ERROR_LABEL


If IsArray(CND_TYPE_RNG) = True Then
    CND_TYPE_VECTOR = CND_TYPE_RNG
    If UBound(CND_TYPE_VECTOR, 2) = 1 Then
        CND_TYPE_VECTOR = MATRIX_TRANSPOSE_FUNC(CND_TYPE_VECTOR)
    End If
Else
    ReDim CND_TYPE_VECTOR(1 To 1, 1 To NCOLUMNS)
    For j = 1 To NCOLUMNS
        CND_TYPE_VECTOR(1, j) = CND_TYPE_RNG
    Next j
End If
If UBound(CND_TYPE_VECTOR, 2) <> NCOLUMNS Then: GoTo ERROR_LABEL

'-----------------------------------------------------------------------------------------------
If VERSION = 0 Then
'-----------------------------------------------------------------------------------------------
    ReDim TEMP_MATRIX(1 To 8, 1 To NCOLUMNS + 1)
    
    TEMP_MATRIX(1, 1) = "FIXED COUPON"
    TEMP_MATRIX(2, 1) = "RISKY (CLEAN PRICE)"
    TEMP_MATRIX(3, 1) = "RISKY (INCL. ACCRUED INTEREST)"
    TEMP_MATRIX(4, 1) = "YIELD TO MATURITY RISKY BOND"
    TEMP_MATRIX(5, 1) = "RISKLESS COUPON BOND (SIGMA=0)"
    TEMP_MATRIX(6, 1) = "RISKLESS  (INCL. ACCRUED INTEREST)"
    TEMP_MATRIX(7, 1) = "YIELD TO MATURITY RISKLESS BOND"
    TEMP_MATRIX(8, 1) = "CREDIT SPREAD (IN BPS)"
    
    For j = 1 To NCOLUMNS
        RISK_FREE_RATE = RISK_FREE_RATE_VECTOR(1, j)
        TENOR = TENOR_VECTOR(1, j)
        BETA_VAL = BETA_VECTOR(1, j)
        H2_VAL = H2_VECTOR(1, j)
        
        ALPHA_VAL = ALPHA_VECTOR(1, j)
        X_VAL = X_VECTOR(1, j)
        W_VAL = W_VECTOR(1, j)
        SIGMA2_VAL = SIGMA2_VECTOR(1, j)
        RHO_VAL = RHO_VECTOR(1, j)
        FIXED_COUPON_VAL = FIXED_COUPON_VECTOR(1, j)
        nLOOPS = nLOOPS_VECTOR(1, j)
        CND_TYPE = CND_TYPE_VECTOR(1, j)
        
        TEMP_MATRIX(1, j + 1) = FIXED_COUPON_VAL
        
        TEMP_MATRIX(2, j + 1) = _
        LS_COUPON_BOND_FUNC(X_VAL, RISK_FREE_RATE, TENOR, TEMP_MATRIX(1, j + 1), _
        W_VAL, ALPHA_VAL, BETA_VAL, SIGMA2_VAL ^ 0.5, H2_VAL ^ 0.5, RHO_VAL, nLOOPS)
        
        TEMP_MATRIX(3, j + 1) = _
        TEMP_MATRIX(2, j + 1) + LS_ACCRUED_COUPON_FUNC(TENOR, TEMP_MATRIX(1, j + 1), 1)
        
        TEMP_MATRIX(4, j + 1) = _
        LS_YTM_FUNC(TEMP_MATRIX(3, j + 1), TENOR, TEMP_MATRIX(1, j + 1), 1, 1)
        
        TEMP_MATRIX(5, j + 1) = _
        LS_COUPON_BOND_FUNC(X_VAL, RISK_FREE_RATE, TENOR, TEMP_MATRIX(1, j + 1), _
        W_VAL, ALPHA_VAL, BETA_VAL, 0, H2_VAL ^ 0.5, RHO_VAL, nLOOPS)
        
        TEMP_MATRIX(6, j + 1) = _
        TEMP_MATRIX(5, j + 1) + LS_ACCRUED_COUPON_FUNC(TENOR, TEMP_MATRIX(1, j + 1), 1)
        
        TEMP_MATRIX(7, j + 1) = _
        LS_YTM_FUNC(TEMP_MATRIX(6, j + 1), TENOR, TEMP_MATRIX(1, j + 1), 1, 1)
        
        If IsNumeric(TEMP_MATRIX(4, j + 1)) And IsNumeric(TEMP_MATRIX(7, j + 1)) Then
            TEMP_MATRIX(8, j + 1) = (TEMP_MATRIX(4, j + 1) - TEMP_MATRIX(7, j + 1)) * 10000
        Else
            TEMP_MATRIX(8, j + 1) = "N/A"
        End If
    Next j
'-----------------------------------------------------------------------------------------------
Else
'-----------------------------------------------------------------------------------------------
    NROWS = UBound(TENOR_VECTOR, 1)
    
    ReDim TEMP_MATRIX(0 To NROWS, 1 To 8)
    TEMP_MATRIX(0, 1) = "TENOR"
    TEMP_MATRIX(0, 2) = "RISKY (CLEAN PRICE)"
    TEMP_MATRIX(0, 3) = "RISKY (INCL. ACCRUED INTEREST)"
    TEMP_MATRIX(0, 4) = "YIELD TO MATURITY RISKY BOND"
    TEMP_MATRIX(0, 5) = "RISKLESS COUPON BOND (SIGMA=0)"
    TEMP_MATRIX(0, 6) = "RISKLESS  (INCL. ACCRUED INTEREST)"
    TEMP_MATRIX(0, 7) = "YIELD TO MATURITY RISKLESS BOND"
    TEMP_MATRIX(0, 8) = "CREDIT SPREAD (IN BPS)"
    
    For i = 1 To NROWS
        TENOR = TENOR_VECTOR(i, 1)
        
        RISK_FREE_RATE = RISK_FREE_RATE_VECTOR(1, 1)
        BETA_VAL = BETA_VECTOR(1, 1)
        H2_VAL = H2_VECTOR(1, 1)
        
        ALPHA_VAL = ALPHA_VECTOR(1, 1)
        X_VAL = X_VECTOR(1, 1)
        W_VAL = W_VECTOR(1, 1)
        SIGMA2_VAL = SIGMA2_VECTOR(1, 1)
        RHO_VAL = RHO_VECTOR(1, 1)
        FIXED_COUPON_VAL = FIXED_COUPON_VECTOR(1, 1)
        nLOOPS = nLOOPS_VECTOR(1, 1)
        CND_TYPE = CND_TYPE_VECTOR(1, 1)
    
        TEMP_MATRIX(i, 1) = TENOR

        TEMP_MATRIX(i, 2) = _
        LS_COUPON_BOND_FUNC(X_VAL, RISK_FREE_RATE, TEMP_MATRIX(i, 1), FIXED_COUPON_VAL, _
        W_VAL, ALPHA_VAL, BETA_VAL, SIGMA2_VAL ^ 0.5, H2_VAL ^ 0.5, RHO_VAL, nLOOPS, CND_TYPE)
        
        TEMP_MATRIX(i, 3) = _
        TEMP_MATRIX(i, 2) + LS_ACCRUED_COUPON_FUNC(TEMP_MATRIX(i, 1), FIXED_COUPON_VAL, 1)
        
        TEMP_MATRIX(i, 4) = _
        LS_YTM_FUNC(TEMP_MATRIX(i, 3), TEMP_MATRIX(i, 1), FIXED_COUPON_VAL, 1, 1, 0, 2)
        
        TEMP_MATRIX(i, 5) = _
        LS_COUPON_BOND_FUNC(X_VAL, RISK_FREE_RATE, TEMP_MATRIX(i, 1), FIXED_COUPON_VAL, _
        W_VAL, ALPHA_VAL, BETA_VAL, 0, H2_VAL ^ 0.5, RHO_VAL, nLOOPS, CND_TYPE)
        
        TEMP_MATRIX(i, 6) = _
        TEMP_MATRIX(i, 5) + LS_ACCRUED_COUPON_FUNC(TEMP_MATRIX(i, 1), FIXED_COUPON_VAL, 1)
        
        TEMP_MATRIX(i, 7) = _
        LS_YTM_FUNC(TEMP_MATRIX(i, 6), TEMP_MATRIX(i, 1), FIXED_COUPON_VAL, 1, 1, 0, 2)
        
        If IsNumeric(TEMP_MATRIX(i, 4)) And IsNumeric(TEMP_MATRIX(i, 7)) Then
            TEMP_MATRIX(i, 8) = (TEMP_MATRIX(i, 4) - TEMP_MATRIX(i, 7)) * 10000
        Else
            TEMP_MATRIX(i, 8) = "N/A"
        End If
    Next i
'-----------------------------------------------------------------------------------------------
End If
'-----------------------------------------------------------------------------------------------

LS_RISKY_COUPON_DEBT_VALUATION_FUNC = TEMP_MATRIX

Exit Function
ERROR_LABEL:
LS_RISKY_COUPON_DEBT_VALUATION_FUNC = Err.number
End Function

'The value of a riskless discount bond given in a the Vasicek(1977) framework.

Public Function LS_D_FUNC(ByVal RISK_FREE_RATE As Double, _
ByVal TENOR As Double, _
ByVal ALPHA As Double, _
ByVal BETA_VAL As Double, _
ByVal H2_VAL As Double)

Dim A_VAL As Double
Dim B_VAL As Double

On Error GoTo ERROR_LABEL

A_VAL = (H2_VAL / (2 * BETA_VAL ^ 2) - ALPHA / BETA_VAL) * TENOR _
   + (H2_VAL / BETA_VAL ^ 3 - ALPHA / BETA_VAL ^ 2) * (Exp(-BETA_VAL * TENOR) - 1) _
   - (H2_VAL / (4 * BETA_VAL ^ 3)) * (Exp(-2 * BETA_VAL * TENOR) - 1)

B_VAL = (1 - Exp(-BETA_VAL * TENOR)) / BETA_VAL
LS_D_FUNC = Exp(A_VAL - B_VAL * RISK_FREE_RATE)

Exit Function
ERROR_LABEL:
LS_D_FUNC = Err.number
End Function


Private Function LS_P_FUNC(ByVal D_VAL As Double, _
ByVal W_VAL As Double, _
ByVal Q_VAL As Double)

On Error GoTo ERROR_LABEL
   
   LS_P_FUNC = D_VAL * (1 - W_VAL * Q_VAL)
   
Exit Function
ERROR_LABEL:
LS_P_FUNC = Err.number
End Function

'Probability under risk neutral measure that default occurs. This measure may
'differ from actual probability of default because the upward drift of the
'asset process under the risk neutral process is different (i.e. lower) than
'the drift of the actual process.

'The value is returned as an array containing all components of Q (q 1 .. nLOOPS
'in L&S notation). For subsequent use, it must first be summed up using
'a user-defined function

Private Function LS_Q_FUNC(ByVal X_VAL As Double, _
ByVal RISK_FREE_RATE As Double, _
ByVal TENOR As Double, _
ByVal ALPHA As Double, _
ByVal BETA_VAL As Double, _
ByVal SIGMA_VAL As Double, _
ByVal SIGMA2_VAL As Double, _
ByVal H_VAL As Double, _
ByVal H2_VAL As Double, _
ByVal RHO_VAL As Double, _
ByVal nLOOPS As Long, _
Optional ByVal CND_TYPE As Integer = 0, _
Optional ByVal OUTPUT As Integer = 0)

Dim i As Long
Dim j As Long
Dim TEMP_SUM As Double
Dim CUMUL_SUM As Double

On Error GoTo ERROR_LABEL

ReDim TEMP_ARR(1 To nLOOPS)

CUMUL_SUM = 0
For i = 1 To nLOOPS
    TEMP_ARR(i) = CND_FUNC((-Log(X_VAL) - LS_M_FUNC(i * TENOR / nLOOPS, TENOR, _
                  RISK_FREE_RATE, ALPHA, BETA_VAL, SIGMA_VAL, SIGMA2_VAL, _
                  H_VAL, H2_VAL, RHO_VAL)) / Sqr(LS_S_FUNC(i * TENOR / nLOOPS, _
                  ALPHA, BETA_VAL, SIGMA_VAL, SIGMA2_VAL, H_VAL, H2_VAL, RHO_VAL)), CND_TYPE)
    If i > 1 Then 'substract sum term if i >1
        TEMP_SUM = 0
        
        For j = 1 To i - 1
          TEMP_SUM = TEMP_SUM + TEMP_ARR(j) * CND_FUNC((LS_M_FUNC(j * TENOR / nLOOPS, TENOR, _
          RISK_FREE_RATE, ALPHA, BETA_VAL, SIGMA_VAL, SIGMA2_VAL, H_VAL, H2_VAL, RHO_VAL) _
          - LS_M_FUNC(i * TENOR / nLOOPS, TENOR, RISK_FREE_RATE, ALPHA, BETA_VAL, SIGMA_VAL, _
          SIGMA2_VAL, H_VAL, H2_VAL, RHO_VAL)) / Sqr(LS_S_FUNC(i * TENOR / nLOOPS, ALPHA, _
          BETA_VAL, SIGMA_VAL, SIGMA2_VAL, H_VAL, H2_VAL, RHO_VAL) - LS_S_FUNC(j * TENOR / nLOOPS, _
          ALPHA, BETA_VAL, SIGMA_VAL, SIGMA2_VAL, H_VAL, H2_VAL, RHO_VAL)), CND_TYPE)
        Next j
        
        TEMP_ARR(i) = TEMP_ARR(i) - TEMP_SUM
    End If
    CUMUL_SUM = CUMUL_SUM + TEMP_ARR(i)
Next i

Select Case OUTPUT
Case 0
    LS_Q_FUNC = TEMP_ARR
Case Else 'Probability of default (risk neutral)
    LS_Q_FUNC = CUMUL_SUM
End Select

Exit Function
ERROR_LABEL:
LS_Q_FUNC = Err.number
End Function

' M term in Longstaff & Schwartz 1995 Valuing Risky Debt

Private Function LS_M_FUNC(ByVal SMT_VAL As Double, _
ByVal TENOR As Double, _
ByVal RISK_FREE_RATE As Double, _
ByVal ALPHA As Double, _
ByVal BETA_VAL As Double, _
ByVal SIGMA_VAL As Double, _
ByVal SIGMA2_VAL As Double, _
ByVal H_VAL As Double, _
ByVal H2_VAL As Double, _
ByVal RHO_VAL As Double)

Dim X1_VAL As Double
Dim X2_VAL As Double
Dim X3_VAL As Double

On Error GoTo ERROR_LABEL

X1_VAL = Exp(-BETA_VAL * TENOR) * (Exp(BETA_VAL * SMT_VAL) - 1)
X2_VAL = (1 - Exp(-BETA_VAL * SMT_VAL))
X3_VAL = Exp(-BETA_VAL * TENOR) * X2_VAL
LS_M_FUNC = ((ALPHA - RHO_VAL * SIGMA_VAL * H_VAL) / BETA_VAL - H2_VAL / _
    BETA_VAL ^ 2 - SIGMA2_VAL / 2) * SMT_VAL _
    + ((RHO_VAL * SIGMA_VAL * H_VAL) / BETA_VAL ^ 2 + H2_VAL / (2 * BETA_VAL ^ 3)) * X1_VAL _
    + (RISK_FREE_RATE / BETA_VAL - ALPHA / BETA_VAL ^ 2 + H2_VAL / BETA_VAL ^ 3) * X2_VAL _
    - (H2_VAL / (2 * BETA_VAL ^ 3)) * X3_VAL

Exit Function
ERROR_LABEL:
LS_M_FUNC = Err.number
End Function

' S term in Longstaff & Schwartz 1995 Valuing Risky Debt

Private Function LS_S_FUNC(ByVal SMT_VAL As Double, _
ByVal ALPHA As Double, _
ByVal BETA_VAL As Double, _
ByVal SIGMA_VAL As Double, _
ByVal SIGMA2_VAL As Double, _
ByVal H_VAL As Double, _
ByVal H2_VAL As Double, _
ByVal RHO_VAL As Double)

Dim X1_VAL As Double
Dim X2_VAL As Double

On Error GoTo ERROR_LABEL

X1_VAL = (1 - Exp(-BETA_VAL * SMT_VAL))
X2_VAL = (1 - Exp(-2 * BETA_VAL * SMT_VAL))
LS_S_FUNC = ((RHO_VAL * SIGMA_VAL * H_VAL) / BETA_VAL + _
    H2_VAL / BETA_VAL ^ 2 + SIGMA2_VAL) * SMT_VAL _
    - ((RHO_VAL * SIGMA_VAL * H_VAL) / BETA_VAL ^ 2 + _
    (2 * H2_VAL) / BETA_VAL ^ 3) * X1_VAL _
    + (H2_VAL / (2 * BETA_VAL ^ 3)) * X2_VAL

Exit Function
ERROR_LABEL:
LS_S_FUNC = Err.number
End Function

' The value of a risky floating-rate payment
' TAU_VAL: time at which coupon occurs (<=T)
' The following parameters were added to reduce number of recalculations of the
' same values (P, DC, and elements of Q which are required for G term
'   W_VAL:     Write-down in case of default (= 1 - recovery rate)
'   D_VAL: value of riskless discount bond
'   P_VAL: value of risky discount bond
'   R_VAL: Expected value of RISK_FREE_RATE at time TAU_VAL(RISK_FREE_RATE term in L&S 95)
'   G_VAL: Correlation adjustment (G term in L&S 95)
' subparameters see function P above

Private Function LS_F_FUNC(ByVal W_VAL As Double, _
ByVal D_VAL As Double, _
ByVal P_VAL As Double, _
ByVal R_VAL As Double, _
ByVal G_VAL As Double)

On Error GoTo ERROR_LABEL

LS_F_FUNC = P_VAL * R_VAL + W_VAL * D_VAL * G_VAL

Exit Function
ERROR_LABEL:
LS_F_FUNC = Err.number
End Function


' Expected value of risk free rate at time TAU_VAL under risk neutral process

Private Function LS_R_EXP_FUNC(ByVal RISK_FREE_RATE As Double, _
ByVal TAU_VAL As Double, _
ByVal TENOR As Double, _
ByVal ALPHA As Double, _
ByVal BETA_VAL As Double, _
ByVal H2_VAL As Double)

On Error GoTo ERROR_LABEL

LS_R_EXP_FUNC = RISK_FREE_RATE * Exp(-BETA_VAL * TAU_VAL) _
       + (ALPHA / BETA_VAL - H2_VAL / BETA_VAL ^ 2) * (1 - Exp(-BETA_VAL * TAU_VAL)) _
       + (H2_VAL / (2 * BETA_VAL ^ 2)) * Exp(-BETA_VAL * TENOR) * (Exp(BETA_VAL * TAU_VAL) - _
       Exp(-BETA_VAL * TAU_VAL))

Exit Function
ERROR_LABEL:
LS_R_EXP_FUNC = Err.number
End Function

' Correlation adjustment in expression for value of floating-rate payment
' Special parameter:
' QTERM_ARR: array with all components of Q term in L&S 95

Private Function LS_G_FUNC(ByVal X_VAL As Double, _
ByVal RISK_FREE_RATE As Double, _
ByVal TAU_VAL As Double, _
ByVal TENOR As Double, _
ByVal nLOOPS As Long, _
ByVal ALPHA As Double, _
ByVal BETA_VAL As Double, _
ByVal SIGMA_VAL As Double, _
ByVal SIGMA2_VAL As Double, _
ByVal H_VAL As Double, _
ByVal H2_VAL As Double, _
ByVal RHO_VAL As Double, _
Optional ByVal CND_TYPE As Integer = 0)

Dim i As Long
Dim C_VAL As Double
Dim TEMP_SUM As Double
Dim SMT_VAL As Double
Dim MIN_TAU_VAL As Double
Dim QTERM_ARR As Variant

On Error GoTo ERROR_LABEL

QTERM_ARR = LS_Q_FUNC(X_VAL, RISK_FREE_RATE, TENOR, ALPHA, BETA_VAL, _
            SIGMA_VAL, SIGMA2_VAL, H_VAL, H2_VAL, RHO_VAL, nLOOPS, _
            CND_TYPE, 0)

If IsArray(QTERM_ARR) = False Then: GoTo ERROR_LABEL

TEMP_SUM = 0
C_VAL = (RHO_VAL * SIGMA_VAL * H_VAL / BETA_VAL + H2_VAL / BETA_VAL ^ 2) * _
        Exp(-BETA_VAL * TAU_VAL) * (Exp(1))

For i = 1 To nLOOPS
    SMT_VAL = i * TENOR / nLOOPS
    MIN_TAU_VAL = MINIMUM_FUNC(TAU_VAL, SMT_VAL)
    C_VAL = (RHO_VAL * SIGMA_VAL * H_VAL / BETA_VAL + H2_VAL / BETA_VAL ^ 2) _
            * Exp(-BETA_VAL * TAU_VAL) * (Exp(BETA_VAL * MIN_TAU_VAL) - 1) _
            - H2_VAL / (2 * BETA_VAL ^ 2) * Exp(-BETA_VAL * TAU_VAL) * _
            Exp(-BETA_VAL * SMT_VAL) * (Exp(2 * BETA_VAL * MIN_TAU_VAL) - 1)
        
    TEMP_SUM = TEMP_SUM + QTERM_ARR(i) * C_VAL * LS_M_FUNC(SMT_VAL, TENOR, RISK_FREE_RATE, _
               ALPHA, BETA_VAL, SIGMA_VAL, SIGMA2_VAL, H_VAL, H2_VAL, RHO_VAL) / _
               LS_S_FUNC(SMT_VAL, ALPHA, BETA_VAL, SIGMA_VAL, SIGMA2_VAL, H_VAL, H2_VAL, RHO_VAL)
Next i
LS_G_FUNC = TEMP_SUM

Exit Function
ERROR_LABEL:
LS_G_FUNC = Err.number
End Function

' PRICE of a Risky Coupon Bond
' according to Longstaff & Schwartz 1995 Valuing Risky Debt

' Parameters:
' W_VAL:     Write-down in case of default (= 1 - recovery rate)
' X_VAL:     ratio of V/K - firm asset value as % of bankruptcy threshold
' RISK_FREE_RATE:     initial short-term riskless interest rate
' TENOR:     time to maturity
' COUPON:    fixed annual coupon in % of face value
' SIGMA_VAL:   instanteaneous stdev of asset process (constant) - SIGMA2_VAL = SIGMA_VAL^2
' H_VAL:     spot rate volatility (constant) - H2_VAL = H_VAL^2
' ALPHA: zeta (long-term equilibrium of mean reverting process plus a constant
'        to represent market PRICE of risk
' BETA_VAL:  "pull-back" factor - speed of adjustment (constant)
' RHO_VAL:   correlation between asset and interest rate process
' nLOOPS:     number of iterations in the calculation of term Q in the formula. Convergence
'        quite good for values > 100.

Private Function LS_COUPON_BOND_FUNC(ByVal X_VAL As Double, _
ByVal RISK_FREE_RATE As Double, _
ByVal TENOR As Double, _
ByVal COUPON As Double, _
ByVal W_VAL As Double, _
ByVal ALPHA As Double, _
ByVal BETA_VAL As Double, _
ByVal SIGMA_VAL As Double, _
ByVal H_VAL As Double, _
ByVal RHO_VAL As Double, _
ByVal nLOOPS As Long, _
Optional ByVal CND_TYPE As Integer = 0)

Dim i As Long
Dim D_VAL As Double
Dim Q_VAL As Double

Dim TEMP_SUM As Double
Dim PERIODS As Double
Dim TEMP_TENOR As Double

Dim SIGMA2_VAL As Double
Dim H2_VAL As Double

On Error GoTo ERROR_LABEL

SIGMA2_VAL = SIGMA_VAL ^ 2
H2_VAL = H_VAL ^ 2
TEMP_SUM = 0
If Int(TENOR) <> TENOR Then
    PERIODS = Int(TENOR) + 1
Else
    PERIODS = Int(TENOR)
End If

For i = 1 To PERIODS
    TEMP_TENOR = TENOR - PERIODS + i
    D_VAL = LS_D_FUNC(RISK_FREE_RATE, TEMP_TENOR, ALPHA, BETA_VAL, H2_VAL)
    Q_VAL = LS_Q_FUNC(X_VAL, RISK_FREE_RATE, TEMP_TENOR, ALPHA, BETA_VAL, _
            SIGMA_VAL, SIGMA2_VAL, H_VAL, H2_VAL, RHO_VAL, nLOOPS, CND_TYPE, 1)
    TEMP_SUM = TEMP_SUM + LS_P_FUNC(D_VAL, W_VAL, Q_VAL) * COUPON  'sum PV of coupons
Next i
TEMP_SUM = TEMP_SUM + LS_P_FUNC(D_VAL, W_VAL, Q_VAL) 'add principal
LS_COUPON_BOND_FUNC = TEMP_SUM

Exit Function
ERROR_LABEL:
LS_COUPON_BOND_FUNC = Err.number
End Function


'PRICE / accrued coupon / YIELD to maturity (continuously compounded)
'of a coupon bond

'Parameters:
'YIELD: YIELD to maturity (continously compounded)
'TENOR:     time to maturity
'COUPON:    fixed annual coupon in % of face value
'FREQUENCY: coupon frequency per year
'REDEMPTION: redemption amount as % of face value

Private Function LS_PV_FUNC(ByVal YIELD As Double, _
ByVal TENOR As Double, _
ByVal COUPON As Double, _
ByVal FREQUENCY As Double, _
ByVal REDEMPTION As Double)

Dim i As Long
Dim PERIODS As Double

Dim TEMP_SUM As Double
Dim TEMP_TENOR As Double

On Error GoTo ERROR_LABEL

TEMP_SUM = 0
If Int(TENOR) = TENOR Then
    PERIODS = TENOR * FREQUENCY
Else
    PERIODS = Int(TENOR * FREQUENCY) + 1
End If

For i = 1 To PERIODS
    TEMP_TENOR = (TENOR * FREQUENCY - PERIODS + i) / FREQUENCY
    TEMP_SUM = TEMP_SUM + Exp(-TEMP_TENOR * YIELD) * COUPON / FREQUENCY
Next i
TEMP_SUM = TEMP_SUM + Exp(-TEMP_TENOR * YIELD) * REDEMPTION
LS_PV_FUNC = TEMP_SUM

Exit Function
ERROR_LABEL:
LS_PV_FUNC = Err.number
End Function

'LS Yield to Mautirity Coupon Bond - Bisec Method
'22 Steps for Convergence

Private Function LS_YTM_FUNC(ByVal CLEAN_PRICE As Double, _
ByVal TENOR As Double, _
ByVal COUPON As Double, _
ByVal FREQUENCY As Double, _
ByVal REDEMPTION As Double, _
Optional LOWER_BOUND As Double = 0, _
Optional UPPER_BOUND As Double = 2)

Dim k As Long
Dim HIGH_VAL As Double
Dim LOW_VAL As Double
Dim TARGET_VAL As Double
Dim ERROR_STR As String
Dim tolerance As Double
    
On Error GoTo ERROR_LABEL
    
tolerance = 0.000001
HIGH_VAL = UPPER_BOUND
LOW_VAL = LOWER_BOUND
k = 1
TARGET_VAL = CLEAN_PRICE

Do While (HIGH_VAL - LOW_VAL) > tolerance
    If LS_PV_FUNC((HIGH_VAL + LOW_VAL) / 2, TENOR, COUPON, FREQUENCY, REDEMPTION) + _
       LS_ACCRUED_COUPON_FUNC(TENOR, COUPON, FREQUENCY) < TARGET_VAL Then
        HIGH_VAL = (HIGH_VAL + LOW_VAL) / 2
    Else
        LOW_VAL = (HIGH_VAL + LOW_VAL) / 2
    End If
    k = k + 1
    If k > 100 Then
        ERROR_STR = "No Convergence"
        GoTo ERROR_LABEL
    End If
Loop

'Debug.Print k
If HIGH_VAL = UPPER_BOUND Or LOW_VAL = LOWER_BOUND Then
    ERROR_STR = "No Solution"
    GoTo ERROR_LABEL
Else
    LS_YTM_FUNC = (HIGH_VAL + LOW_VAL) / 2
End If

Exit Function
ERROR_LABEL:
If ERROR_STR = Empty Then
    ERROR_STR = "# " & Str(Err.number) & " was generated by " _
                    & Err.source & " " & Err.Description
End If
LS_YTM_FUNC = ERROR_STR
End Function

Private Function LS_ACCRUED_COUPON_FUNC(ByVal TENOR As Double, _
ByVal COUPON As Double, _
ByVal FREQUENCY As Double)

On Error GoTo ERROR_LABEL

If Int(TENOR) = TENOR Then
   LS_ACCRUED_COUPON_FUNC = COUPON / FREQUENCY
Else
   LS_ACCRUED_COUPON_FUNC = (TENOR - Int(FREQUENCY * TENOR) / FREQUENCY) * COUPON
End If

Exit Function
ERROR_LABEL:
LS_ACCRUED_COUPON_FUNC = Err.number
End Function

'------------------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------------------
'In their 1995 paper F. Longstaff and Eduardo Schwartz develop a closed-form solutions for the
'valuation of fixed an floating rate debt subject to both credit and interest rate risk.

'While the algorithm implements both the discount bond and floating rate payment valuation, it
'also generates yield / maturity table for risky fixed coupon bonds. These can be considered as
'a portfolio of risky discount bonds and are valued accordingly.

'As a note of caution, parameter n is used to calculate one of the terms iteratively in this model.
'For precise values, n would have to be set to values of 100 to 200. Yet this will slow the model,
'particularly the recalculation of the data tables. So to gain a qualitative appreciation of
'parameter sensitivities, leave the value in the 10 to 20 range.

'Reference:
'Longstaff, Francis A., Schwartz, Eduardo S. A simple approach to Valuing Risky Fixed and Floating
'Rate Debt. Journal of Finance. Vol 3. July 1995.
'------------------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------------------
