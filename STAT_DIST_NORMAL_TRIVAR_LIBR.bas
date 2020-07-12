Attribute VB_Name = "STAT_DIST_NORMAL_TRIVAR_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.

'************************************************************************************
'************************************************************************************
'FUNCTION      : STECK_TRIVAR_CUMUL_NORM_FUNC

'DESCRIPTION   :
'George P.Steck function for Computing Trivariate Normal Probabilities
'REFERENCE:
'George p.Steck
'Source: Ann. Math. Statist. Volume 29, Number 3 (1958), 780-800.
'Related Items:
'See Correction: George P. Steck, Correction Notes:
'Corrections to "A Table for Computing Trivariate Normal Probabilities".
'Ann. Math. Statist., Volume 30, Number 4 (1959), 1267--1267.

'LIBRARY       : STATISTICS
'GROUP         : DIST_NORMAL_TRIVAR
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/22/2009
'************************************************************************************
'************************************************************************************

Function STECK_TRIVAR_CUMUL_NORM_FUNC(ByVal X1_VAL As Double, _
ByVal X2_VAL As Double, _
ByVal X3_VAL As Double, _
ByVal RHO_12_VAL As Double, _
ByVal RHO_31_VAL As Double, _
ByVal RHO_23_VAL As Double)

'     FOR THE TRIVARIATE NORMAL R.V. (X,Y,W) WITH ZERO MEANS, UNIT
'     VARIANCE AND CORR(X,Y)=RHO12,CORR(Y,W)RHO23,CORR(W,X)=RHO31,TNOR
'     USES THE METHOD OF SECTION 2 OF STECK (1958) TO COMPUTE Z= PR
'     (X<X1,Y<X2,W<X3) ACCURACY OF THE RESULT DEPENDS ON THE
'     SUBPROGRAMS T AND STK.

'     format is  call tnorm(p,1.97,-1.96,2.84,.1,.2,.3,999)
'         (  999 is alternate return label  )
'        SUBROUTINE T AND STK USE CLOSED FORM APPROXIMATIONS
'        ACCURATE TO ABOUT .00025,  TIME IS ABOUT .003

Dim ii As Integer
Dim jj As Integer
Dim kk As Integer

Dim A_TEMP_VAL As Double
Dim B_TEMP_VAL As Double
Dim C_TEMP_VAL As Double
Dim D_TEMP_VAL As Double
Dim E_TEMP_VAL As Double
Dim F_TEMP_VAL As Double
Dim G_TEMP_VAL As Double
Dim H_TEMP_VAL As Double

Dim ATEMP_ARR As Variant
Dim BTEMP_ARR As Variant
Dim KTEMP_ARR As Variant
Dim LTEMP_ARR As Variant
Dim CTEMP_ARR As Variant
Dim DTEMP_ARR As Variant
Dim HTEMP_ARR As Variant
Dim ITEMP_ARR As Variant
Dim JTEMP_ARR As Variant
Dim ETEMP_ARR As Variant
Dim FTEMP_ARR As Variant
Dim GTEMP_ARR As Variant

ReDim ATEMP_ARR(0 To 5)
ReDim BTEMP_ARR(0 To 5)
ReDim CTEMP_ARR(0 To 5)
ReDim DTEMP_ARR(0 To 5)
ReDim ETEMP_ARR(0 To 5)
ReDim FTEMP_ARR(0 To 2)
ReDim GTEMP_ARR(0 To 5)
ReDim HTEMP_ARR(0 To 5)
ReDim ITEMP_ARR(0 To 5)
ReDim JTEMP_ARR(0 To 2)
ReDim KTEMP_ARR(0 To 5)
ReDim LTEMP_ARR(0 To 5)

On Error GoTo ERROR_LABEL

    DTEMP_ARR(0) = X1_VAL
    DTEMP_ARR(1) = X2_VAL
    DTEMP_ARR(2) = X3_VAL
    
    For ii = 1 To 3
        A_TEMP_VAL = DTEMP_ARR(ii - 1)
        If (Abs(A_TEMP_VAL) >= 0.000000000000001) Then: GoTo 1983
        DTEMP_ARR(ii - 1) = 0.000000000000001
1983:
    Next ii

    CTEMP_ARR(0) = RHO_12_VAL
    CTEMP_ARR(1) = RHO_23_VAL
    CTEMP_ARR(2) = RHO_31_VAL
    kk = 0

'   IND=0 FOR ALL DTEMP_ARR'DTEMP_ARR OF SAME SIGN, =1 OTHERWISE.
'   IF DTEMP_ARR'S ARE OF
'   DIFFERENT SIGNS, THEN PERMUTE TO HAVE DTEMP_ARR(3) OF DIFF'T SIGN
'   AND CHANGE DTEMP_ARR(3),R(2), AND R(3). SIGNS.
    
    If (DTEMP_ARR(0) <> 0#) Then
        GoTo 1985
    Else
    GoTo 1984
    End If
1984:
    If ((DTEMP_ARR(1) * DTEMP_ARR(2)) >= 0#) Then
        GoTo 1991
    Else
        GoTo 1990
    End If
1985:
    If ((DTEMP_ARR(0) * DTEMP_ARR(1)) >= 0#) Then
        GoTo 1986
    Else
        GoTo 1987
    End If
1986:
    If ((DTEMP_ARR(0) * DTEMP_ARR(2)) >= 0#) Then
        GoTo 1991
    Else
        GoTo 1990
    End If
1987:
    If ((DTEMP_ARR(0) * DTEMP_ARR(2)) >= 0#) Then
        GoTo 1989
    Else
        GoTo 1988
    End If
1988:
    
    H_TEMP_VAL = DTEMP_ARR
    DTEMP_ARR = DTEMP_ARR(2)
    DTEMP_ARR(2) = H_TEMP_VAL
        
    H_TEMP_VAL = CTEMP_ARR
    CTEMP_ARR = CTEMP_ARR(1)
    CTEMP_ARR(1) = H_TEMP_VAL
    
    GoTo 1990
1989:
    
    H_TEMP_VAL = DTEMP_ARR(1)
    DTEMP_ARR(1) = DTEMP_ARR(2)
    DTEMP_ARR(2) = H_TEMP_VAL
        
    H_TEMP_VAL = CTEMP_ARR
    CTEMP_ARR = CTEMP_ARR(2)
    CTEMP_ARR(2) = H_TEMP_VAL
    
1990:
    kk = 1
    DTEMP_ARR(2) = -DTEMP_ARR(2)
    CTEMP_ARR(1) = -CTEMP_ARR(1)
    CTEMP_ARR(2) = -CTEMP_ARR(2)
1991:
    For ii = 1 To 3
        HTEMP_ARR(ii - 1) = 1# - CTEMP_ARR(ii - 1) * CTEMP_ARR(ii - 1)
    Next ii
    E_TEMP_VAL = HTEMP_ARR(0) + HTEMP_ARR(1) + _
                 HTEMP_ARR(2) - 2 + CTEMP_ARR(0) * 2 * _
                 CTEMP_ARR(1) * CTEMP_ARR(2)
    If (E_TEMP_VAL > 0.000000000000001) Then: GoTo 1992
    
    If (E_TEMP_VAL < -0.0001) Then
        STECK_TRIVAR_CUMUL_NORM_FUNC = 0
        Exit Function
    End If
    F_TEMP_VAL = 0.999999
1992:
    E_TEMP_VAL = 1 / Sqr(E_TEMP_VAL)
    G_TEMP_VAL = 0
    GTEMP_ARR(0) = 1 / Sqr(HTEMP_ARR(0))
    GTEMP_ARR(1) = 1 / Sqr(HTEMP_ARR(1))
    GTEMP_ARR(2) = 1 / Sqr(HTEMP_ARR(2))
    
    For ii = 1 To 3
        jj = ii + 3
        DTEMP_ARR(jj - 1) = DTEMP_ARR(ii - 1)
        FTEMP_ARR(ii - 1) = 1 / DTEMP_ARR(ii - 1)
    
        Call STECK_PHI_FUNC(DTEMP_ARR(ii - 1), LTEMP_ARR(ii - 1), _
                        B_TEMP_VAL, C_TEMP_VAL, D_TEMP_VAL)
    
        LTEMP_ARR(jj - 1) = LTEMP_ARR(ii - 1)
        HTEMP_ARR(jj - 1) = HTEMP_ARR(ii - 1)
        GTEMP_ARR(jj - 1) = GTEMP_ARR(ii - 1)
        CTEMP_ARR(jj - 1) = CTEMP_ARR(ii - 1)
        JTEMP_ARR(ii - 1) = CTEMP_ARR(ii - 1) * CTEMP_ARR(ii + 1) - CTEMP_ARR(ii)
        ITEMP_ARR(ii - 1) = DTEMP_ARR(ii) - DTEMP_ARR(ii - 1) * CTEMP_ARR(ii - 1)
        A_TEMP_VAL = ITEMP_ARR(ii - 1)
        If (Abs(A_TEMP_VAL) > 1E-25) Then: GoTo 1993
        ITEMP_ARR(ii - 1) = 1E-25
1993:
        ITEMP_ARR(jj - 1) = DTEMP_ARR(ii + 1) - DTEMP_ARR(ii - 1) * _
                        CTEMP_ARR(ii + 1)
        A_TEMP_VAL = ITEMP_ARR(jj - 1)
        If (Abs(A_TEMP_VAL) >= 1E-25) Then: GoTo 1994
        ITEMP_ARR(jj - 1) = 1E-25
1994:
        ATEMP_ARR(ii - 1) = ITEMP_ARR(ii - 1) * FTEMP_ARR(ii - 1) * _
                            GTEMP_ARR(ii - 1)
        ATEMP_ARR(jj - 1) = ITEMP_ARR(jj - 1) * FTEMP_ARR(ii - 1) * _
                            GTEMP_ARR(ii + 1)
        BTEMP_ARR(ii - 1) = E_TEMP_VAL * (JTEMP_ARR(ii - 1) + _
                            HTEMP_ARR(ii - 1) * ITEMP_ARR(jj - 1) / _
                            ITEMP_ARR(ii - 1))
        BTEMP_ARR(jj - 1) = E_TEMP_VAL * (JTEMP_ARR(ii - 1) + _
                            HTEMP_ARR(ii + 1) * ITEMP_ARR(ii - 1) / _
                            ITEMP_ARR(jj - 1))
    
        G_TEMP_VAL = G_TEMP_VAL + (1# - STECK_DEL_FUNC(ATEMP_ARR(ii - 1), _
                    ATEMP_ARR(jj - 1))) * LTEMP_ARR(ii - 1)
    Next ii
    
    G_TEMP_VAL = G_TEMP_VAL * 0.5
    
    For ii = 1 To 6
        
        ETEMP_ARR(ii - 1) = STECK_T_FUNC(DTEMP_ARR(ii - 1), _
                            ATEMP_ARR(ii - 1), LTEMP_ARR(ii - 1))
        
        KTEMP_ARR(ii - 1) = STECK_S_FN_FUNC(DTEMP_ARR(ii - 1), _
                            ATEMP_ARR(ii - 1), BTEMP_ARR(ii - 1), _
                            LTEMP_ARR(ii - 1), ETEMP_ARR(ii - 1))
        G_TEMP_VAL = G_TEMP_VAL - ETEMP_ARR(ii - 1) * 0.5 - KTEMP_ARR(ii - 1)
    
    Next ii
    
    If (kk = 0) Then
        STECK_TRIVAR_CUMUL_NORM_FUNC = G_TEMP_VAL
        Exit Function
    End If
    
    G_TEMP_VAL = (LTEMP_ARR(0) + LTEMP_ARR(4) - _
                STECK_DEL_FUNC(DTEMP_ARR(0), DTEMP_ARR(1))) * _
                0.5 - ETEMP_ARR(0) - ETEMP_ARR(4) - G_TEMP_VAL
    
    STECK_TRIVAR_CUMUL_NORM_FUNC = G_TEMP_VAL

Exit Function
ERROR_LABEL:
STECK_TRIVAR_CUMUL_NORM_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : STECK_S_FN_FUNC

'DESCRIPTION   :
'COMPUTES STECK'S (1958 ANN MATH STAT) S-FN WITH MAX ERROR E-9,
'THOUGH COMPUTER ROUNDING ERRORS MAY REDUCE THIS ACCURACY. CALLS
'FUNCTION SUBPROGRAMS STK(X,A,B) IN WHICH DABS(B) < 1.1, PHI(X), AND
'T(X,A,PHIX) NOTE THAT PHIX = PHI(X) AND TXA = T(X,A,PHIX).

'LIBRARY       : STATISTICS
'GROUP         : DIST_NORMAL_TRIVAR
'ID            : 002
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/22/2009
'************************************************************************************
'************************************************************************************

Private Function STECK_S_FN_FUNC(ByVal X_TEMP_VAL As Double, _
ByVal A_TEMP_VAL As Double, _
ByVal B_TEMP_VAL As Double, _
ByVal C_TEMP_VAL As Double, _
ByVal D_TEMP_VAL As Double)

Dim E_TEMP_VAL As Double
Dim F_TEMP_VAL As Double
Dim G_TEMP_VAL As Double
Dim H_TEMP_VAL As Double
Dim I_TEMP_VAL As Double
Dim J_TEMP_VAL As Double
Dim K_TEMP_VAL As Double
Dim L_TEMP_VAL As Double
Dim M_TEMP_VAL As Double
Dim N_TEMP_VAL As Double
Dim O_TEMP_VAL As Double
Dim P_TEMP_VAL As Double
Dim Q_TEMP_VAL As Double
Dim R_TEMP_VAL As Double

Dim RESULT_VAL As Double
    
On Error GoTo ERROR_LABEL
    
    J_TEMP_VAL = Abs(B_TEMP_VAL)
    If (J_TEMP_VAL > 1.1) Then: GoTo 1983
    
    RESULT_VAL = STECK_STK_FUNC(X_TEMP_VAL, A_TEMP_VAL, B_TEMP_VAL)
    STECK_S_FN_FUNC = RESULT_VAL
    Exit Function
1983:
    L_TEMP_VAL = Abs(D_TEMP_VAL)
    I_TEMP_VAL = Abs(A_TEMP_VAL)
    R_TEMP_VAL = 1 / J_TEMP_VAL
    E_TEMP_VAL = C_TEMP_VAL - 0.5
    N_TEMP_VAL = Abs(E_TEMP_VAL)
    M_TEMP_VAL = Abs(X_TEMP_VAL)
    P_TEMP_VAL = M_TEMP_VAL * I_TEMP_VAL * J_TEMP_VAL
    Call STECK_PHI_FUNC(P_TEMP_VAL, O_TEMP_VAL, _
        F_TEMP_VAL, G_TEMP_VAL, H_TEMP_VAL)
    Q_TEMP_VAL = O_TEMP_VAL - 0.5
    If (I_TEMP_VAL <= 1) Then: GoTo 1984
    
    E_TEMP_VAL = 1 / I_TEMP_VAL
    RESULT_VAL = Q_TEMP_VAL * L_TEMP_VAL - N_TEMP_VAL * _
            STECK_T_FUNC(P_TEMP_VAL, R_TEMP_VAL, O_TEMP_VAL) + _
            (N_TEMP_VAL - Q_TEMP_VAL) * 0.25 + _
            STECK_STK_FUNC(P_TEMP_VAL, R_TEMP_VAL, E_TEMP_VAL)
    GoTo 1985
1984:
    K_TEMP_VAL = I_TEMP_VAL * J_TEMP_VAL
    E_TEMP_VAL = 1 / K_TEMP_VAL
    RESULT_VAL = (N_TEMP_VAL + 0.5) * 0.25 + Q_TEMP_VAL * L_TEMP_VAL - _
                STECK_STK_FUNC(P_TEMP_VAL, E_TEMP_VAL, I_TEMP_VAL) - _
                STECK_STK_FUNC(M_TEMP_VAL, K_TEMP_VAL, R_TEMP_VAL)
1985:
    If (X_TEMP_VAL >= 0) Then: GoTo 1986
    
    RESULT_VAL = Atn(J_TEMP_VAL / Sqr(A_TEMP_VAL * A_TEMP_VAL * _
                (B_TEMP_VAL * B_TEMP_VAL + 1) + 1)) * _
                0.1591549 - RESULT_VAL
1986:
    If (B_TEMP_VAL >= 0) Then
        STECK_S_FN_FUNC = RESULT_VAL
        Exit Function
    End If
    RESULT_VAL = -RESULT_VAL
    STECK_S_FN_FUNC = RESULT_VAL

Exit Function
ERROR_LABEL:
STECK_S_FN_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : STECK_T_FUNC
'DESCRIPTION   : OWENS T DOUBLE PRECISION FUNCTION ACCURATE TO .00005
' D_TEMP_VAL IS A DUMMY ARGUMENT

'LIBRARY       : STATISTICS
'GROUP         : DIST_NORMAL_TRIVAR
'ID            : 003
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/22/2009
'************************************************************************************
'************************************************************************************

Private Function STECK_T_FUNC(ByVal Z_TEMP_VAL As Double, _
ByVal A_TEMP_VAL As Double, _
ByVal D_TEMP_VAL As Double)

Dim E_TEMP_VAL As Double
Dim F_TEMP_VAL As Double
Dim G_TEMP_VAL As Double
Dim H_TEMP_VAL As Double
Dim I_TEMP_VAL As Double
Dim J_TEMP_VAL As Double
Dim K_TEMP_VAL As Double
Dim L_TEMP_VAL As Double
Dim M_TEMP_VAL As Double
Dim N_TEMP_VAL As Double
Dim O_TEMP_VAL As Double
Dim P_TEMP_VAL As Double

Dim RESULT_VAL As Double
    
On Error GoTo ERROR_LABEL
    
    L_TEMP_VAL = Z_TEMP_VAL
    J_TEMP_VAL = A_TEMP_VAL
    M_TEMP_VAL = 0
    K_TEMP_VAL = 1
    RESULT_VAL = 0
    If (A_TEMP_VAL >= 0) Then: GoTo 1983
    J_TEMP_VAL = -J_TEMP_VAL
    K_TEMP_VAL = -1
1983:
    If (J_TEMP_VAL <= 1) Then: GoTo 1984
    
    L_TEMP_VAL = Abs(L_TEMP_VAL)
    
    Call STECK_PHI_FUNC(L_TEMP_VAL, N_TEMP_VAL, F_TEMP_VAL, G_TEMP_VAL, H_TEMP_VAL)
    L_TEMP_VAL = J_TEMP_VAL * L_TEMP_VAL
    J_TEMP_VAL = 1 / J_TEMP_VAL
    
    Call STECK_PHI_FUNC(L_TEMP_VAL, O_TEMP_VAL, F_TEMP_VAL, G_TEMP_VAL, H_TEMP_VAL)
    M_TEMP_VAL = (N_TEMP_VAL + O_TEMP_VAL) * 0.5 - N_TEMP_VAL * O_TEMP_VAL
    M_TEMP_VAL = M_TEMP_VAL * K_TEMP_VAL
    K_TEMP_VAL = -K_TEMP_VAL
1984:
    I_TEMP_VAL = Atn(J_TEMP_VAL)
    P_TEMP_VAL = L_TEMP_VAL * L_TEMP_VAL
    If (P_TEMP_VAL > 150) Then: GoTo 1985
    
    RESULT_VAL = K_TEMP_VAL * I_TEMP_VAL * 0.15915494 * _
                Exp(P_TEMP_VAL * -0.5 * J_TEMP_VAL / I_TEMP_VAL)
    'Computing 4th power
    E_TEMP_VAL = L_TEMP_VAL * J_TEMP_VAL
    E_TEMP_VAL = E_TEMP_VAL * E_TEMP_VAL
    RESULT_VAL = RESULT_VAL * (E_TEMP_VAL * E_TEMP_VAL * 0.00868 + 1)
    RESULT_VAL = RESULT_VAL + M_TEMP_VAL
1985:
    STECK_T_FUNC = RESULT_VAL

Exit Function
ERROR_LABEL:
STECK_T_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : STECK_PHI_FUNC

'DESCRIPTION   : ROUTINE RETURNS CUMULATIVE NORMAL OF Z IN ret, DENSITY IN F,
'-F/(1.-PHI) IN G, F/PHI IN H

'LIBRARY       : STATISTICS
'GROUP         : DIST_NORMAL_TRIVAR
'ID            : 004
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/22/2009
'************************************************************************************
'************************************************************************************

Private Function STECK_PHI_FUNC(ByRef W_TEMP_VAL As Variant, _
ByRef Z_TEMP_VAL As Variant, _
ByRef F_TEMP_VAL As Variant, _
ByRef G_TEMP_VAL As Variant, _
ByRef H_TEMP_VAL As Variant)

Dim A1_TEMP_VAL As Double
Dim A2_TEMP_VAL As Double
Dim A3_TEMP_VAL As Double
Dim A4_TEMP_VAL As Double
Dim A5_TEMP_VAL As Double

Dim B1_TEMP_VAL As Double
Dim B2_TEMP_VAL As Double

Dim C1_TEMP_VAL As Double
Dim C2_TEMP_VAL As Double
Dim C3_TEMP_VAL As Double

Dim D1_TEMP_VAL As Double
Dim D2_TEMP_VAL As Double
Dim D3_TEMP_VAL As Double

Dim RESULT_VAL As Double

Dim V_TEMP_VAL As Double
Dim X_TEMP_VAL As Double
Dim Y_TEMP_VAL As Double

On Error GoTo ERROR_LABEL

    A1_TEMP_VAL = 1.330274429
    A2_TEMP_VAL = 1.821255978
    A3_TEMP_VAL = 1.781477937
    A4_TEMP_VAL = 0.356563782
    A5_TEMP_VAL = 0.31938153
    B1_TEMP_VAL = 0.2316419
    B2_TEMP_VAL = 0.39894228
    C1_TEMP_VAL = 0
    C2_TEMP_VAL = 1
    C3_TEMP_VAL = 7.5
    
    V_TEMP_VAL = W_TEMP_VAL
    If (Abs(V_TEMP_VAL) > 7.5) Then
    
        If (C3_TEMP_VAL > 0) Then
          X_TEMP_VAL = C3_TEMP_VAL
        Else
          X_TEMP_VAL = -C3_TEMP_VAL
        End If
        If (V_TEMP_VAL > 0) Then
              Y_TEMP_VAL = X_TEMP_VAL
        Else: Y_TEMP_VAL = -X_TEMP_VAL
        End If
        V_TEMP_VAL = Y_TEMP_VAL
    End If
    'Computing 2nd power
    D1_TEMP_VAL = V_TEMP_VAL
    
    F_TEMP_VAL = B2_TEMP_VAL * Exp(-(D1_TEMP_VAL * D1_TEMP_VAL) / 2)
    
    D3_TEMP_VAL = C2_TEMP_VAL / (C2_TEMP_VAL + B1_TEMP_VAL * Abs(V_TEMP_VAL))
    
    D2_TEMP_VAL = ((((A1_TEMP_VAL * D3_TEMP_VAL - A2_TEMP_VAL) * D3_TEMP_VAL + _
            A3_TEMP_VAL) * D3_TEMP_VAL - A4_TEMP_VAL) * _
            D3_TEMP_VAL + A5_TEMP_VAL) * D3_TEMP_VAL
    
    RESULT_VAL = F_TEMP_VAL * D2_TEMP_VAL
    If (V_TEMP_VAL > C1_TEMP_VAL) Then: RESULT_VAL = C2_TEMP_VAL - RESULT_VAL
    
    G_TEMP_VAL = -F_TEMP_VAL / (C2_TEMP_VAL - RESULT_VAL)
    H_TEMP_VAL = F_TEMP_VAL / RESULT_VAL
    Z_TEMP_VAL = RESULT_VAL

Exit Function
ERROR_LABEL:
STECK_PHI_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : STECK_STK_FUNC
'DESCRIPTION   :
'LIBRARY       : STATISTICS
'GROUP         : DIST_NORMAL_TRIVAR
'ID            : 005
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/22/2009
'************************************************************************************
'************************************************************************************

Private Function STECK_STK_FUNC(ByVal X_TEMP_VAL As Double, _
ByVal A_TEMP_VAL As Double, _
ByVal B_TEMP_VAL As Double)

Dim C_TEMP_VAL As Double
Dim D_TEMP_VAL As Double
Dim E_TEMP_VAL As Double
Dim F_TEMP_VAL As Double

Dim G_TEMP_VAL As Double
Dim H_TEMP_VAL As Double
Dim I_TEMP_VAL As Double

Dim RESULT_VAL As Double

On Error GoTo ERROR_LABEL

'ACCURATE TO .0001
C_TEMP_VAL = A_TEMP_VAL * A_TEMP_VAL + 1

'Computing 2nd power
F_TEMP_VAL = A_TEMP_VAL * 0.5 * B_TEMP_VAL
D_TEMP_VAL = Sqr(C_TEMP_VAL + F_TEMP_VAL * F_TEMP_VAL)

'Computing 2nd power
F_TEMP_VAL = A_TEMP_VAL * B_TEMP_VAL
E_TEMP_VAL = B_TEMP_VAL / Sqr(C_TEMP_VAL + F_TEMP_VAL * F_TEMP_VAL)
F_TEMP_VAL = X_TEMP_VAL * D_TEMP_VAL

Call STECK_PHI_FUNC(F_TEMP_VAL, RESULT_VAL, G_TEMP_VAL, H_TEMP_VAL, I_TEMP_VAL)

RESULT_VAL = RESULT_VAL * 0.1591549 * Atn(E_TEMP_VAL)
STECK_STK_FUNC = RESULT_VAL

Exit Function
ERROR_LABEL:
STECK_STK_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : STECK_DEL_FUNC
'DESCRIPTION   :
'LIBRARY       : STATISTICS
'GROUP         : DIST_NORMAL_TRIVAR
'ID            : 006
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/22/2009
'************************************************************************************
'************************************************************************************

Private Function STECK_DEL_FUNC(ByVal X_TEMP_VAL As Double, _
ByVal Y_TEMP_VAL As Double)

Dim TEMP_MULT

On Error GoTo ERROR_LABEL

    TEMP_MULT = X_TEMP_VAL * Y_TEMP_VAL
    
    If (TEMP_MULT < 0#) Then
        GoTo 1983
    ElseIf (TEMP_MULT = 0) Then
        GoTo 1985
    Else
        GoTo 1984
    End If
1983:
    STECK_DEL_FUNC = 1
    Exit Function
1984:
    STECK_DEL_FUNC = 0
    Exit Function
1985:
    If (X_TEMP_VAL + Y_TEMP_VAL >= 0#) Then
        GoTo 1984
    Else
        GoTo 1983
    End If

Exit Function
ERROR_LABEL:
STECK_DEL_FUNC = Err.number
End Function

