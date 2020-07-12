Attribute VB_Name = "MATRIX_EIGEN_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.


'************************************************************************************
'************************************************************************************
'FUNCTION      : MATRIX_EIGEN_SQUARE_FUNC

'DESCRIPTION   : This function is used for doing eigen value decomposition on
'a square matrix. Additionally includes an option to sort the eigen values and
'corresponding eigen vectors.

'Reference: Handbook for Auto. Comp., Vol.ii-Linear Algebra, and
'the corresponding Fortran subroutine in EISPACK.
'by Bowdler, Martin, Reinsch, and Wilkinson,

'LIBRARY       : MATRIX
'GROUP         : EIGEN
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/28/2009
'************************************************************************************
'************************************************************************************

Function MATRIX_EIGEN_SQUARE_FUNC(ByRef DATA_RNG As Variant, _
Optional ByVal SORT_FLAG As Boolean = True, _
Optional ByVal OUTPUT As Integer = 0)
  
Dim h As Single
Dim i As Single
Dim j As Single
Dim k As Single
Dim l As Single
Dim m As Single
Dim n As Single

Dim kk As Single
Dim ll As Single
Dim mm As Single
Dim nn As Single
 
Dim A_VAL As Double
Dim B_VAL As Double
Dim C_VAL As Double
Dim D_VAL As Double
Dim E_VAL As Double
Dim F_VAL As Double
Dim G_VAL As Double
Dim H_VAL As Double
Dim I_VAL As Double
Dim J_VAL As Double
Dim K_VAL As Double
Dim L_VAL As Double
Dim M_VAL As Double
Dim N_VAL As Double
 
Dim O_VAL As Double
Dim P_VAL As Double
Dim Q_VAL As Double
Dim R_VAL As Double
 
Dim X_VAL As Double
Dim Y_VAL As Double
Dim Z_VAL As Double
  
Dim RA_VAL As Double
Dim SA_VAL As Double
  
Dim VR_VAL As Double
Dim VI_VAL As Double
  
Dim NSIZE As Single
Dim LOW_VAL As Single
Dim HIGH_VAL As Single
Dim nLOOPS As Single
Dim SHIFT_VAL As Double
Dim COUNTER As Single
Dim NORM_VAL As Double
Dim INT_VAL As Single
Dim NOT_LAST_VAL As Single
   
Dim CPLX_DIV_REAL_VAL As Double
Dim CPLX_DIV_IMAG_VAL As Double
   
Dim TEMP_ARR As Variant
Dim TEMP_VECTOR As Variant
Dim TEMP_MATRIX As Variant
Dim DATA_MATRIX As Variant
 
Dim EIGEN_ARR As Variant
Dim HESSEN_MATRIX As Variant
Dim ORTHOGONAL_ARR As Variant
Dim REAL_EIGEN_VALUES_MATRIX As Variant 'GetRealEigenValues
Dim IMAG_EIGEN_VALUES_MATRIX As Variant 'GetImaginaryEigenValues
 
Dim EIGEN_VALUES_MATRIX As Variant
Dim EIGEN_VECTORS_MATRIX As Variant 'GetEigenVectorsMatrix
 
Dim epsilon As Double
  
On Error GoTo ERROR_LABEL
  
DATA_MATRIX = DATA_RNG
NSIZE = UBound(DATA_MATRIX, 2) - LBound(DATA_MATRIX, 2) + 1
ll = NSIZE
DATA_MATRIX = MATRIX_CHANGE_BASE_ZERO_FUNC(DATA_MATRIX)
ReDim EIGEN_VECTORS_MATRIX(0 To ll - 1, 0 To ll - 1)
ReDim REAL_EIGEN_VALUES_MATRIX(0 To ll - 1)
ReDim IMAG_EIGEN_VALUES_MATRIX(0 To ll - 1)
kk = 1
  
j = 0
Do While ((j < ll) And (kk = 1))
    Do While ((i < ll) And (kk = 1))
        If DATA_MATRIX(i, j) = DATA_MATRIX(j, i) Then
            kk = 1
        Else
            kk = 0
        End If
        i = i + 1
    Loop
    j = j + 1
Loop
  
'-----------------------------------------------------------------
If kk = 1 Then     ' Triagonalize.
'-----------------------------------------------------------------
    For i = 0 To ll - 1
        For j = 0 To ll - 1
            EIGEN_VECTORS_MATRIX(i, j) = DATA_MATRIX(i, j)
        Next j
    Next i
    ll = NSIZE
    For j = 0 To ll - 1
        REAL_EIGEN_VALUES_MATRIX(j) = EIGEN_VECTORS_MATRIX(ll - 1, j)
    Next j

    For i = ll - 1 To 1 Step -1 'Householder reduction to tridiagonal form
        J_VAL = 0 'Scale to avoid under/overflow
        I_VAL = 0
        For k = 0 To i - 1
            J_VAL = J_VAL + Abs(REAL_EIGEN_VALUES_MATRIX(k))
        Next k
        '------------------------------------------------------------------------
        If (J_VAL = 0) Then
        '------------------------------------------------------------------------
            IMAG_EIGEN_VALUES_MATRIX(i) = REAL_EIGEN_VALUES_MATRIX(i - 1)
            For j = 0 To i - 1
                REAL_EIGEN_VALUES_MATRIX(j) = EIGEN_VECTORS_MATRIX(i - 1, j)
                EIGEN_VECTORS_MATRIX(i, j) = 0
                EIGEN_VECTORS_MATRIX(j, i) = 0
            Next j
        '------------------------------------------------------------------------
         Else 'Generate Householder vector
        '------------------------------------------------------------------------
            For k = 0 To i - 1
               REAL_EIGEN_VALUES_MATRIX(k) = REAL_EIGEN_VALUES_MATRIX(k) / J_VAL
               I_VAL = I_VAL + REAL_EIGEN_VALUES_MATRIX(k) * _
                                     REAL_EIGEN_VALUES_MATRIX(k)
            Next k
            F_VAL = REAL_EIGEN_VALUES_MATRIX(i - 1)
            G_VAL = SQRT_FUNC(I_VAL)
            If (F_VAL > 0) Then
               G_VAL = -G_VAL
            End If
            IMAG_EIGEN_VALUES_MATRIX(i) = J_VAL * G_VAL
            I_VAL = I_VAL - F_VAL * G_VAL
            REAL_EIGEN_VALUES_MATRIX(i - 1) = F_VAL - G_VAL
            For j = 0 To i - 1
               IMAG_EIGEN_VALUES_MATRIX(j) = 0
            Next j
            For j = 0 To i - 1 'Apply similarity transformation to remaining columns
               F_VAL = REAL_EIGEN_VALUES_MATRIX(j)
               EIGEN_VECTORS_MATRIX(j, i) = F_VAL
               G_VAL = IMAG_EIGEN_VALUES_MATRIX(j) + EIGEN_VECTORS_MATRIX(j, j) * F_VAL
               For k = j + 1 To i - 1
                  G_VAL = G_VAL + EIGEN_VECTORS_MATRIX(k, j) * REAL_EIGEN_VALUES_MATRIX(k)
                  IMAG_EIGEN_VALUES_MATRIX(k) = IMAG_EIGEN_VALUES_MATRIX(k) + _
                                                EIGEN_VECTORS_MATRIX(k, j) * F_VAL
               Next k
               IMAG_EIGEN_VALUES_MATRIX(j) = G_VAL
            Next j
            F_VAL = 0
            For j = 0 To i - 1
               IMAG_EIGEN_VALUES_MATRIX(j) = IMAG_EIGEN_VALUES_MATRIX(j) / I_VAL
               F_VAL = F_VAL + IMAG_EIGEN_VALUES_MATRIX(j) * REAL_EIGEN_VALUES_MATRIX(j)
            Next j
            H_VAL = F_VAL / (I_VAL + I_VAL)
            For j = 0 To i - 1
               IMAG_EIGEN_VALUES_MATRIX(j) = IMAG_EIGEN_VALUES_MATRIX(j) - _
                                        H_VAL * REAL_EIGEN_VALUES_MATRIX(j)
            Next j
            For j = 0 To i - 1
               F_VAL = REAL_EIGEN_VALUES_MATRIX(j)
               G_VAL = IMAG_EIGEN_VALUES_MATRIX(j)
               For k = j To i - 1
                  EIGEN_VECTORS_MATRIX(k, j) = EIGEN_VECTORS_MATRIX(k, j) - _
                                          (F_VAL * IMAG_EIGEN_VALUES_MATRIX(k) + _
                                          G_VAL * REAL_EIGEN_VALUES_MATRIX(k))
               Next k
               REAL_EIGEN_VALUES_MATRIX(j) = EIGEN_VECTORS_MATRIX(i - 1, j)
               EIGEN_VECTORS_MATRIX(i, j) = 0
            Next j
        '------------------------------------------------------------------------
         End If
        '------------------------------------------------------------------------
         REAL_EIGEN_VALUES_MATRIX(i) = I_VAL
      Next i
      
      For i = 0 To ll - 2 'Accumulate transformations
         EIGEN_VECTORS_MATRIX(ll - 1, i) = EIGEN_VECTORS_MATRIX(i, i)
         EIGEN_VECTORS_MATRIX(i, i) = 1
         I_VAL = REAL_EIGEN_VALUES_MATRIX(i + 1)
         If (I_VAL <> 0) Then
            For k = 0 To i
               REAL_EIGEN_VALUES_MATRIX(k) = EIGEN_VECTORS_MATRIX(k, i + 1) / I_VAL
            Next k
            For j = 0 To i
               G_VAL = 0
               For k = 0 To i
                  G_VAL = G_VAL + EIGEN_VECTORS_MATRIX(k, i + 1) * EIGEN_VECTORS_MATRIX(k, j)
               Next k
               For k = 0 To i
                  EIGEN_VECTORS_MATRIX(k, j) = EIGEN_VECTORS_MATRIX(k, j) - _
                                           G_VAL * REAL_EIGEN_VALUES_MATRIX(k)
               Next k
            Next j
         End If
         For k = 0 To i
            EIGEN_VECTORS_MATRIX(k, i + 1) = 0
         Next k
      Next i
      
      For j = 0 To ll - 1
         REAL_EIGEN_VALUES_MATRIX(j) = EIGEN_VECTORS_MATRIX(ll - 1, j)
         EIGEN_VECTORS_MATRIX(ll - 1, j) = 0
      Next j
      EIGEN_VECTORS_MATRIX(ll - 1, ll - 1) = 1
      IMAG_EIGEN_VALUES_MATRIX(0) = 0
'-----------------------------------------------------------------
    ' Diagonalize.
'-----------------------------------------------------------------
      For i = 1 To ll - 1 'Symmetric tridiagonal QL algorithm
         IMAG_EIGEN_VALUES_MATRIX(i - 1) = IMAG_EIGEN_VALUES_MATRIX(i)
      Next i
      IMAG_EIGEN_VALUES_MATRIX(ll - 1) = 0
      F_VAL = 0
      R_VAL = 0
      epsilon = 2 ^ -52
      For l = 0 To ll - 1
         ' Find small subdiagonal element
         R_VAL = MAXIMUM_FUNC(R_VAL, Abs(REAL_EIGEN_VALUES_MATRIX(l)) + Abs(IMAG_EIGEN_VALUES_MATRIX(l)))
         m = l
         ' Original while-loop from Java code
         Do While (m < ll)
            If (Abs(IMAG_EIGEN_VALUES_MATRIX(m)) <= epsilon * R_VAL) Then
               Exit Do
            End If
            m = m + 1
         Loop
         ' If m == l, REAL_EIGEN_VALUES_MATRIX(l) is an eigenvalue,
         ' otherwise, iterate.
         If (m > l) Then
            nLOOPS = 0
            Do While (Abs(IMAG_EIGEN_VALUES_MATRIX(l)) > epsilon * R_VAL)
               nLOOPS = nLOOPS + 1  ' (Could check iteration count here.)
               ' Compute implicit shift
               G_VAL = REAL_EIGEN_VALUES_MATRIX(l)
               M_VAL = (REAL_EIGEN_VALUES_MATRIX(l + 1) - G_VAL) / (2 * IMAG_EIGEN_VALUES_MATRIX(l))
               N_VAL = HYPOT_FUNC(M_VAL, 1)
               If (M_VAL < 0) Then
                  N_VAL = -N_VAL
               End If
               REAL_EIGEN_VALUES_MATRIX(l) = IMAG_EIGEN_VALUES_MATRIX(l) / (M_VAL + N_VAL)
               REAL_EIGEN_VALUES_MATRIX(l + 1) = IMAG_EIGEN_VALUES_MATRIX(l) * (M_VAL + N_VAL)
               E_VAL = REAL_EIGEN_VALUES_MATRIX(l + 1)
               I_VAL = G_VAL - REAL_EIGEN_VALUES_MATRIX(l)
               For i = l + 2 To ll - 1
                  REAL_EIGEN_VALUES_MATRIX(i) = REAL_EIGEN_VALUES_MATRIX(i) - I_VAL
               Next i
               F_VAL = F_VAL + I_VAL
               ' Implicit QL transformation.
               M_VAL = REAL_EIGEN_VALUES_MATRIX(m)
               A_VAL = 1
               B_VAL = A_VAL
               C_VAL = A_VAL
               D_VAL = IMAG_EIGEN_VALUES_MATRIX(l + 1)
               K_VAL = 0
               L_VAL = 0
               For i = m - 1 To l Step -1
                  C_VAL = B_VAL
                  B_VAL = A_VAL
                  L_VAL = K_VAL
                  G_VAL = A_VAL * IMAG_EIGEN_VALUES_MATRIX(i)
                  I_VAL = A_VAL * M_VAL
                  N_VAL = HYPOT_FUNC(M_VAL, CDbl(IMAG_EIGEN_VALUES_MATRIX(i)))
                  IMAG_EIGEN_VALUES_MATRIX(i + 1) = K_VAL * N_VAL
                  K_VAL = IMAG_EIGEN_VALUES_MATRIX(i) / N_VAL
                  A_VAL = M_VAL / N_VAL
                  M_VAL = A_VAL * REAL_EIGEN_VALUES_MATRIX(i) - K_VAL * G_VAL
                  REAL_EIGEN_VALUES_MATRIX(i + 1) = I_VAL + K_VAL * (A_VAL * G_VAL + K_VAL * _
                  REAL_EIGEN_VALUES_MATRIX(i))
                  ' Accumulate transformation.
                  For k = 0 To ll - 1
                     I_VAL = EIGEN_VECTORS_MATRIX(k, i + 1)
                     EIGEN_VECTORS_MATRIX(k, i + 1) = K_VAL * EIGEN_VECTORS_MATRIX(k, i) + A_VAL * I_VAL
                     EIGEN_VECTORS_MATRIX(k, i) = A_VAL * EIGEN_VECTORS_MATRIX(k, i) - K_VAL * I_VAL
                  Next k
               Next i
               M_VAL = -K_VAL * L_VAL * C_VAL * D_VAL * IMAG_EIGEN_VALUES_MATRIX(l) / E_VAL
               IMAG_EIGEN_VALUES_MATRIX(l) = K_VAL * M_VAL
               REAL_EIGEN_VALUES_MATRIX(l) = A_VAL * M_VAL
               ' Check for convergence.
            Loop
         End If 'if (m > l)
         REAL_EIGEN_VALUES_MATRIX(l) = REAL_EIGEN_VALUES_MATRIX(l) + F_VAL
         IMAG_EIGEN_VALUES_MATRIX(l) = 0
      Next l
      ' Sort eigenvalues and corresponding vectors.
      For i = 0 To ll - 2
         k = i
         M_VAL = REAL_EIGEN_VALUES_MATRIX(i)
         For j = i + 1 To ll - 1
            If (REAL_EIGEN_VALUES_MATRIX(j) < M_VAL) Then
               k = j
               M_VAL = REAL_EIGEN_VALUES_MATRIX(j)
            End If
         Next j
         If (k <> i) Then
            REAL_EIGEN_VALUES_MATRIX(k) = REAL_EIGEN_VALUES_MATRIX(i)
            REAL_EIGEN_VALUES_MATRIX(i) = M_VAL
            For j = 0 To ll - 1
               M_VAL = EIGEN_VECTORS_MATRIX(j, i)
               EIGEN_VECTORS_MATRIX(j, i) = EIGEN_VECTORS_MATRIX(j, k)
               EIGEN_VECTORS_MATRIX(j, k) = M_VAL
            Next j
         End If
      Next i

'-----------------------------------------------------------------
  Else
'-----------------------------------------------------------------
    ReDim HESSEN_MATRIX(0 To ll - 1, 0 To ll - 1)
    ReDim ORTHOGONAL_ARR(0 To ll - 1)
    For j = 0 To ll - 1
      For i = 0 To ll - 1
        HESSEN_MATRIX(i, j) = DATA_MATRIX(i, j)
      Next i
    Next j
'-----------------------------------------------------------------
    'Reduce to Hessenberg form.
'-----------------------------------------------------------------
    LOW_VAL = 0
    ll = NSIZE
    HIGH_VAL = ll - 1
    For h = LOW_VAL + 1 To HIGH_VAL - 1
    ' Scale column.
      J_VAL = 0
      For i = h To HIGH_VAL
         J_VAL = J_VAL + Abs(HESSEN_MATRIX(i, h - 1))
      Next i
      If J_VAL <> 0 Then
        'Compute Householder transformation.
        I_VAL = 0
        For i = HIGH_VAL To h Step -1
          ORTHOGONAL_ARR(i) = HESSEN_MATRIX(i, h - 1) / J_VAL
          I_VAL = I_VAL + ORTHOGONAL_ARR(i) * ORTHOGONAL_ARR(i)
        Next i
        G_VAL = I_VAL ^ 0.5
        If ORTHOGONAL_ARR(h) > 0 Then
          G_VAL = -G_VAL
        End If
        I_VAL = I_VAL - ORTHOGONAL_ARR(h) * G_VAL
        ORTHOGONAL_ARR(h) = ORTHOGONAL_ARR(h) - G_VAL
        ' Apply Householder similarity transformation
        ' HESSEN_MATRIX = (I-u*u'/HESSEN_MATRIX)*HESSEN_MATRIX*(I-u*u')/HESSEN_MATRIX)
        For j = h To ll - 1
          F_VAL = 0
          For i = HIGH_VAL To h Step -1
            F_VAL = F_VAL + ORTHOGONAL_ARR(i) * HESSEN_MATRIX(i, j)
          Next i
          F_VAL = F_VAL / I_VAL
          For i = h To HIGH_VAL
            HESSEN_MATRIX(i, j) = HESSEN_MATRIX(i, j) - F_VAL * ORTHOGONAL_ARR(i)
          Next i
        Next j
        For i = 0 To HIGH_VAL
          F_VAL = 0
          For j = HIGH_VAL To h Step -1
            F_VAL = F_VAL + ORTHOGONAL_ARR(j) * HESSEN_MATRIX(i, j)
          Next j
          F_VAL = F_VAL / I_VAL
          For j = h To HIGH_VAL
            HESSEN_MATRIX(i, j) = HESSEN_MATRIX(i, j) - F_VAL * ORTHOGONAL_ARR(j)
          Next j
        Next i
        ORTHOGONAL_ARR(h) = J_VAL * ORTHOGONAL_ARR(h)
        HESSEN_MATRIX(h, h - 1) = J_VAL * G_VAL
      End If
    Next h
    'Accumulate transformations (Algol'K_VAL ortran).
    For i = 0 To ll - 1
      For j = 0 To ll - 1
        If i = j Then
          EIGEN_VECTORS_MATRIX(i, j) = 1
        Else
          EIGEN_VECTORS_MATRIX(i, j) = 0
        End If
      Next j
    Next i
    For h = HIGH_VAL - 1 To LOW_VAL + 1 Step -1
      If HESSEN_MATRIX(h, h - 1) <> 0 Then
        For i = h + 1 To HIGH_VAL
          ORTHOGONAL_ARR(i) = HESSEN_MATRIX(i, h - 1)
        Next i
        For j = h To HIGH_VAL
          G_VAL = 0
          For i = h To HIGH_VAL
            G_VAL = G_VAL + ORTHOGONAL_ARR(i) * EIGEN_VECTORS_MATRIX(i, j)
          Next i
          'Double division avoids possible underflow
          G_VAL = (G_VAL / ORTHOGONAL_ARR(h)) / HESSEN_MATRIX(h, h - 1)
          For i = h To HIGH_VAL
            EIGEN_VECTORS_MATRIX(i, j) = EIGEN_VECTORS_MATRIX(i, j) + G_VAL * ORTHOGONAL_ARR(i)
          Next i
        Next j
      End If
    Next h
'-----------------------------------------------------------------
    ' Reduce Hessenberg to real Schur form.
'-----------------------------------------------------------------
  ll = NSIZE
  n = ll - 1
  LOW_VAL = 0
  HIGH_VAL = ll - 1
  epsilon = 2 ^ -52
  COUNTER = 1
  SHIFT_VAL = 0
  M_VAL = 0
  Q_VAL = 0
  N_VAL = 0
  K_VAL = 0
  Z_VAL = 0
  'Store roots isolated by balanc and compute matrix NORM_VAL
  NORM_VAL = 0
  For i = 0 To ll - 1
    If ((i < LOW_VAL) Or (i > HIGH_VAL)) Then
      REAL_EIGEN_VALUES_MATRIX(i) = HESSEN_MATRIX(i, i)
      IMAG_EIGEN_VALUES_MATRIX(i) = 0
    End If
    INT_VAL = MAXIMUM_FUNC(i - 1, 0)
    For j = INT_VAL To ll - 1
      NORM_VAL = NORM_VAL + Abs(HESSEN_MATRIX(i, j))
    Next j
  Next i
   
  'Outer loop over eigenvalue index
  nLOOPS = 0
  Do While (n >= LOW_VAL)
    'Look for single small sub-diagonal element
    nn = n
    Do While (nn > LOW_VAL)
      K_VAL = Abs(HESSEN_MATRIX(nn - 1, nn - 1)) + Abs(HESSEN_MATRIX(nn, nn))
      If K_VAL = 0 Then
        K_VAL = NORM_VAL
      End If
      If (Abs(HESSEN_MATRIX(nn, nn - 1)) < epsilon * K_VAL) Then
        Exit Do
      End If
      nn = nn - 1
    Loop
    
    ' Check for convergence
    ' One root found
    If nn = n Then
      HESSEN_MATRIX(n, n) = HESSEN_MATRIX(n, n) + SHIFT_VAL
      REAL_EIGEN_VALUES_MATRIX(n) = HESSEN_MATRIX(n, n)
      IMAG_EIGEN_VALUES_MATRIX(n) = 0
      n = n - 1
      nLOOPS = 0
      ' Two roots found
    ElseIf nn = n - 1 Then
      P_VAL = HESSEN_MATRIX(n, n - 1) * HESSEN_MATRIX(n - 1, n)
      M_VAL = (HESSEN_MATRIX(n - 1, n - 1) - HESSEN_MATRIX(n, n)) / 2
      Q_VAL = M_VAL * M_VAL + P_VAL
      Z_VAL = (Abs(Q_VAL)) ^ 0.5
      HESSEN_MATRIX(n, n) = HESSEN_MATRIX(n, n) + SHIFT_VAL
      HESSEN_MATRIX(n - 1, n - 1) = HESSEN_MATRIX(n - 1, n - 1) + SHIFT_VAL
      X_VAL = HESSEN_MATRIX(n, n)
      ' Real pair
      If Q_VAL >= 0 Then
        If M_VAL >= 0 Then
          Z_VAL = M_VAL + Z_VAL
        Else
          Z_VAL = M_VAL - Z_VAL
        End If
        REAL_EIGEN_VALUES_MATRIX(n - 1) = X_VAL + Z_VAL
        REAL_EIGEN_VALUES_MATRIX(n) = REAL_EIGEN_VALUES_MATRIX(n - 1)
        If Z_VAL <> 0 Then
          REAL_EIGEN_VALUES_MATRIX(n) = X_VAL - P_VAL / Z_VAL
        End If
        IMAG_EIGEN_VALUES_MATRIX(n - 1) = 0
        IMAG_EIGEN_VALUES_MATRIX(n) = 0
        X_VAL = HESSEN_MATRIX(n, n - 1)
        K_VAL = Abs(X_VAL) + Abs(Z_VAL)
        M_VAL = X_VAL / K_VAL
        Q_VAL = Z_VAL / K_VAL
        N_VAL = (M_VAL * M_VAL + Q_VAL * Q_VAL) ^ 0.5
        M_VAL = M_VAL / N_VAL
        Q_VAL = Q_VAL / N_VAL
        
        'Row modification
        For j = n - 1 To ll - 1
          Z_VAL = HESSEN_MATRIX(n - 1, j)
          HESSEN_MATRIX(n - 1, j) = Q_VAL * Z_VAL + M_VAL * HESSEN_MATRIX(n, j)
          HESSEN_MATRIX(n, j) = Q_VAL * HESSEN_MATRIX(n, j) - M_VAL * Z_VAL
        Next j
        
        'Column modification
        For i = 0 To n
          Z_VAL = HESSEN_MATRIX(i, n - 1)
          HESSEN_MATRIX(i, n - 1) = Q_VAL * Z_VAL + M_VAL * HESSEN_MATRIX(i, n)
          HESSEN_MATRIX(i, n) = Q_VAL * HESSEN_MATRIX(i, n) - M_VAL * Z_VAL
        Next i
        
        'Accumulate transformations
        For i = LOW_VAL To HIGH_VAL
          Z_VAL = EIGEN_VECTORS_MATRIX(i, n - 1)
          EIGEN_VECTORS_MATRIX(i, n - 1) = Q_VAL * Z_VAL + M_VAL * EIGEN_VECTORS_MATRIX(i, n)
          EIGEN_VECTORS_MATRIX(i, n) = Q_VAL * EIGEN_VECTORS_MATRIX(i, n) - M_VAL * Z_VAL
        Next i
      'Complex pair
      Else
        REAL_EIGEN_VALUES_MATRIX(n - 1) = X_VAL + M_VAL
        REAL_EIGEN_VALUES_MATRIX(n) = X_VAL + M_VAL
        IMAG_EIGEN_VALUES_MATRIX(n - 1) = Z_VAL
        IMAG_EIGEN_VALUES_MATRIX(n) = -Z_VAL
      End If
      n = n - 2
      nLOOPS = 0
      ' No convergence yet
    Else
            ' Form shift
            X_VAL = HESSEN_MATRIX(n, n)
            Y_VAL = 0
            P_VAL = 0
            If nn < n Then
               Y_VAL = HESSEN_MATRIX(n - 1, n - 1)
               P_VAL = HESSEN_MATRIX(n, n - 1) * HESSEN_MATRIX(n - 1, n)
            End If
   
            ' Wilkinson'K_VAL original ad hoc shift
   
            If nLOOPS = 10 Then
               SHIFT_VAL = SHIFT_VAL + X_VAL
               For i = LOW_VAL To n
                  HESSEN_MATRIX(i, i) = HESSEN_MATRIX(i, i) - X_VAL
               Next i
               K_VAL = Abs(HESSEN_MATRIX(n, n - 1)) + Abs(HESSEN_MATRIX(n - 1, n - 2))
               X_VAL = Y_VAL = 0.75 * K_VAL
               P_VAL = -0.4375 * K_VAL * K_VAL
            End If

            ' MATLAB'K_VAL new ad hoc shift

            If nLOOPS = 30 Then
                K_VAL = (Y_VAL - X_VAL) / 2
                K_VAL = K_VAL * K_VAL + P_VAL
                If K_VAL > 0 Then
                    K_VAL = SQRT_FUNC(K_VAL)
                    If Y_VAL < X_VAL Then
                       K_VAL = -K_VAL
                    End If
                    K_VAL = X_VAL - P_VAL / ((Y_VAL - X_VAL) / 2 + K_VAL)
                    For i = LOW_VAL To n
                       HESSEN_MATRIX(i, i) = HESSEN_MATRIX(i, i) - K_VAL
                    Next i
                    SHIFT_VAL = SHIFT_VAL + K_VAL
                    P_VAL = 0.964
                    Y_VAL = P_VAL
                    X_VAL = Y_VAL
                End If
            End If
   
            nLOOPS = nLOOPS + 1   ' (Could check iteration count here.)
                ' Look for two consecutive small sub-diagonal elements
            m = n - 2
            Do While (m >= nn)
               Z_VAL = HESSEN_MATRIX(m, m)
               N_VAL = X_VAL - Z_VAL
               K_VAL = Y_VAL - Z_VAL
               M_VAL = (N_VAL * K_VAL - P_VAL) / HESSEN_MATRIX(m + 1, m) + HESSEN_MATRIX(m, m + 1)
               Q_VAL = HESSEN_MATRIX(m + 1, m + 1) - Z_VAL - N_VAL - K_VAL
               N_VAL = HESSEN_MATRIX(m + 2, m + 1)
               K_VAL = Abs(M_VAL) + Abs(Q_VAL) + Abs(N_VAL)
               M_VAL = M_VAL / K_VAL
               Q_VAL = Q_VAL / K_VAL
               N_VAL = N_VAL / K_VAL
               If (m = nn) Then
                  Exit Do
               End If
               If (Abs(HESSEN_MATRIX(m, m - 1)) * (Abs(Q_VAL) + Abs(N_VAL)) < _
                  epsilon * (Abs(M_VAL) * (Abs(HESSEN_MATRIX(m - 1, m - 1)) + Abs(Z_VAL) + _
                  Abs(HESSEN_MATRIX(m + 1, m + 1))))) Then
                     Exit Do
               End If
               m = m - 1
            Loop
   
            For i = m + 2 To n
               HESSEN_MATRIX(i, i - 2) = 0
               If (i > m + 2) Then
                  HESSEN_MATRIX(i, i - 3) = 0
               End If
            Next i
   
            ' Double QR step involving rows nn:n and columns m:n
            For k = m To n - 1
              If (k <> n - 1) Then
                NOT_LAST_VAL = 1
              Else
                NOT_LAST_VAL = 0
              End If
              If (k <> m) Then
                  M_VAL = HESSEN_MATRIX(k, k - 1)
                  Q_VAL = HESSEN_MATRIX(k + 1, k - 1)
                  If NOT_LAST_VAL > 0 Then
                    N_VAL = HESSEN_MATRIX(k + 2, k - 1)
                  Else
                    N_VAL = 0
                  End If
                  X_VAL = Abs(M_VAL) + Abs(Q_VAL) + Abs(N_VAL)
                  If (X_VAL <> 0) Then
                     M_VAL = M_VAL / X_VAL
                     Q_VAL = Q_VAL / X_VAL
                     N_VAL = N_VAL / X_VAL
                  End If
               End If
               If (X_VAL = 0) Then
                  Exit Do
               End If
               K_VAL = SQRT_FUNC(M_VAL * M_VAL + Q_VAL * Q_VAL + N_VAL * N_VAL)
               If (M_VAL < 0) Then
                  K_VAL = -K_VAL
               End If
               If (K_VAL <> 0) Then
                  If (k <> m) Then
                     HESSEN_MATRIX(k, k - 1) = -K_VAL * X_VAL
                  ElseIf (nn <> m) Then
                     HESSEN_MATRIX(k, k - 1) = -HESSEN_MATRIX(k, k - 1)
                  End If
                  M_VAL = M_VAL + K_VAL
                  X_VAL = M_VAL / K_VAL
                  Y_VAL = Q_VAL / K_VAL
                  Z_VAL = N_VAL / K_VAL
                  Q_VAL = Q_VAL / M_VAL
                  N_VAL = N_VAL / M_VAL
   
                  ' Row modification
   
                  For j = k To ll - 1
                     M_VAL = HESSEN_MATRIX(k, j) + Q_VAL * HESSEN_MATRIX(k + 1, j)
                     If (NOT_LAST_VAL) Then
                        M_VAL = M_VAL + N_VAL * HESSEN_MATRIX(k + 2, j)
                        HESSEN_MATRIX(k + 2, j) = HESSEN_MATRIX(k + 2, j) - M_VAL * Z_VAL
                     End If
                     HESSEN_MATRIX(k, j) = HESSEN_MATRIX(k, j) - M_VAL * X_VAL
                     HESSEN_MATRIX(k + 1, j) = HESSEN_MATRIX(k + 1, j) - M_VAL * Y_VAL
                  Next j
   
                  ' Column modification
   
                  For i = 0 To MINIMUM_FUNC(Int(n), Int(k + 3))
                     M_VAL = X_VAL * HESSEN_MATRIX(i, k) + Y_VAL * HESSEN_MATRIX(i, k + 1)
                     If (NOT_LAST_VAL) Then
                        M_VAL = M_VAL + Z_VAL * HESSEN_MATRIX(i, k + 2)
                        HESSEN_MATRIX(i, k + 2) = HESSEN_MATRIX(i, k + 2) - M_VAL * N_VAL
                     End If
                     HESSEN_MATRIX(i, k) = HESSEN_MATRIX(i, k) - M_VAL
                     HESSEN_MATRIX(i, k + 1) = HESSEN_MATRIX(i, k + 1) - M_VAL * Q_VAL
                  Next i
   
                  ' Accumulate transformations
   
                  For i = LOW_VAL To HIGH_VAL
                     M_VAL = X_VAL * EIGEN_VECTORS_MATRIX(i, k) + Y_VAL * EIGEN_VECTORS_MATRIX(i, k + 1)
                     If (NOT_LAST_VAL) Then
                        M_VAL = M_VAL + Z_VAL * EIGEN_VECTORS_MATRIX(i, k + 2)
                        EIGEN_VECTORS_MATRIX(i, k + 2) = EIGEN_VECTORS_MATRIX(i, k + 2) - M_VAL * N_VAL
                     End If
                     EIGEN_VECTORS_MATRIX(i, k) = EIGEN_VECTORS_MATRIX(i, k) - M_VAL
                     EIGEN_VECTORS_MATRIX(i, k + 1) = EIGEN_VECTORS_MATRIX(i, k + 1) - M_VAL * Q_VAL
                  Next i
               End If ' (K_VAL != 0)
            Next k
    End If ' check convergence
  Loop 'while (n >= LOW_VAL)
      ' Backsubstitute to find vectors of upper triangular form
      If (NORM_VAL = 0) Then
        n = ll + 1
        GoTo 1983
      End If
   
      For n = ll - 1 To 0 Step -1
         M_VAL = REAL_EIGEN_VALUES_MATRIX(n)
         Q_VAL = IMAG_EIGEN_VALUES_MATRIX(n)
   
         ' Real vector
   
         If (Q_VAL = 0) Then
            
            nn = n
            HESSEN_MATRIX(n, n) = 1
            For i = n - 1 To 0 Step -1
               P_VAL = HESSEN_MATRIX(i, i) - M_VAL
               N_VAL = 0
               For j = nn To n
                  N_VAL = N_VAL + HESSEN_MATRIX(i, j) * HESSEN_MATRIX(j, n)
               Next j
               If (IMAG_EIGEN_VALUES_MATRIX(i) < 0) Then
                  Z_VAL = P_VAL
                  K_VAL = N_VAL
               Else
                  nn = i
                  If (IMAG_EIGEN_VALUES_MATRIX(i) = 0) Then
                     If (P_VAL <> 0) Then
                        HESSEN_MATRIX(i, n) = -N_VAL / P_VAL
                     Else
                        HESSEN_MATRIX(i, n) = -N_VAL / (epsilon * NORM_VAL)
                     End If
   
                  ' Solve real equations
   
                  Else
                     X_VAL = HESSEN_MATRIX(i, i + 1)
                     Y_VAL = HESSEN_MATRIX(i + 1, i)
                     Q_VAL = (REAL_EIGEN_VALUES_MATRIX(i) - M_VAL) * (REAL_EIGEN_VALUES_MATRIX(i) - _
                     M_VAL) + IMAG_EIGEN_VALUES_MATRIX(i) * IMAG_EIGEN_VALUES_MATRIX(i)
                     O_VAL = (X_VAL * K_VAL - Z_VAL * N_VAL) / Q_VAL
                     HESSEN_MATRIX(i, n) = O_VAL
                     If (Abs(X_VAL) > Abs(Z_VAL)) Then
                        HESSEN_MATRIX(i + 1, n) = (-N_VAL - P_VAL * O_VAL) / X_VAL
                     Else
                        HESSEN_MATRIX(i + 1, n) = (-K_VAL - Y_VAL * O_VAL) / Z_VAL
                     End If
                  End If
   
                  ' Overflow control
   
                  O_VAL = Abs(HESSEN_MATRIX(i, n))
                  If ((epsilon * O_VAL) * O_VAL > 1) Then
                     For j = i To n
                        HESSEN_MATRIX(j, n) = HESSEN_MATRIX(j, n) / O_VAL
                     Next j
                  End If
               End If ' If (IMAG_EIGEN_VALUES_MATRIX(i) < 0)
            Next i
   
         ' Complex vector
   
         ElseIf (Q_VAL < 0) Then
            nn = n - 1

            ' Last vector component imaginary so matrix is triangular
   
            If (Abs(HESSEN_MATRIX(n, n - 1)) > Abs(HESSEN_MATRIX(n - 1, n))) Then
               HESSEN_MATRIX(n - 1, n - 1) = Q_VAL / HESSEN_MATRIX(n, n - 1)
               HESSEN_MATRIX(n - 1, n) = -(HESSEN_MATRIX(n, n) - M_VAL) / HESSEN_MATRIX(n, n - 1)
            Else
               TEMP_ARR = COMPLEX_QUOTIENT_ARRAY_FUNC(0, -HESSEN_MATRIX(n - 1, n), _
               HESSEN_MATRIX(n - 1, n - 1) - M_VAL, Q_VAL)
               CPLX_DIV_REAL_VAL = LBound(TEMP_ARR)
               CPLX_DIV_IMAG_VAL = UBound(TEMP_ARR)
               
               HESSEN_MATRIX(n - 1, n - 1) = CPLX_DIV_REAL_VAL
               HESSEN_MATRIX(n - 1, n) = CPLX_DIV_IMAG_VAL
            End If
            HESSEN_MATRIX(n, n - 1) = 0
            HESSEN_MATRIX(n, n) = 1
            For i = n - 2 To 0 Step -1
               RA_VAL = 0
               SA_VAL = 0
               For j = nn To n
                  RA_VAL = RA_VAL + HESSEN_MATRIX(i, j) * HESSEN_MATRIX(j, n - 1)
                  SA_VAL = SA_VAL + HESSEN_MATRIX(i, j) * HESSEN_MATRIX(j, n)
               Next j
               P_VAL = HESSEN_MATRIX(i, i) - M_VAL
   
               If (IMAG_EIGEN_VALUES_MATRIX(i) < 0) Then
                  Z_VAL = P_VAL
                  N_VAL = RA_VAL
                  K_VAL = SA_VAL
               Else
                  nn = i
                  If (IMAG_EIGEN_VALUES_MATRIX(i) = 0) Then
                     TEMP_ARR = COMPLEX_QUOTIENT_ARRAY_FUNC(-RA_VAL, -SA_VAL, P_VAL, Q_VAL)
                     CPLX_DIV_REAL_VAL = LBound(TEMP_ARR)
                     CPLX_DIV_IMAG_VAL = UBound(TEMP_ARR)
                     HESSEN_MATRIX(i, n - 1) = CPLX_DIV_REAL_VAL
                     HESSEN_MATRIX(i, n) = CPLX_DIV_IMAG_VAL
                  Else
   
                     ' Solve complex equations
   
                     X_VAL = HESSEN_MATRIX(i, i + 1)
                     Y_VAL = HESSEN_MATRIX(i + 1, i)
                     VR_VAL = (REAL_EIGEN_VALUES_MATRIX(i) - M_VAL) * (REAL_EIGEN_VALUES_MATRIX(i) - _
                     M_VAL) + IMAG_EIGEN_VALUES_MATRIX(i) * IMAG_EIGEN_VALUES_MATRIX(i) - Q_VAL * Q_VAL
                     VI_VAL = (REAL_EIGEN_VALUES_MATRIX(i) - M_VAL) * 2 * Q_VAL
                     If ((VR_VAL = 0) And (VI_VAL = 0)) Then
                        VR_VAL = epsilon * NORM_VAL * (Abs(P_VAL) + Abs(Q_VAL) + _
                        Abs(X_VAL) + Abs(Y_VAL) + Abs(Z_VAL))
                     End If
                     TEMP_ARR = COMPLEX_QUOTIENT_ARRAY_FUNC(X_VAL * N_VAL - Z_VAL * RA_VAL + Q_VAL * SA_VAL, _
                     X_VAL * K_VAL - Z_VAL * SA_VAL - Q_VAL * RA_VAL, VR_VAL, VI_VAL)
                     CPLX_DIV_REAL_VAL = LBound(TEMP_ARR)
                     CPLX_DIV_IMAG_VAL = UBound(TEMP_ARR)
                     HESSEN_MATRIX(i, n - 1) = CPLX_DIV_REAL_VAL
                     HESSEN_MATRIX(i, n) = CPLX_DIV_IMAG_VAL
                     If (Abs(X_VAL) > (Abs(Z_VAL) + Abs(Q_VAL))) Then
                        HESSEN_MATRIX(i + 1, n - 1) = (-RA_VAL - P_VAL * HESSEN_MATRIX(i, n - 1) + _
                        Q_VAL * HESSEN_MATRIX(i, n)) / X_VAL
                        HESSEN_MATRIX(i + 1, n) = (-SA_VAL - P_VAL * HESSEN_MATRIX(i, n) - Q_VAL * _
                        HESSEN_MATRIX(i, n - 1)) / X_VAL
                     Else
                        TEMP_ARR = COMPLEX_QUOTIENT_ARRAY_FUNC(-N_VAL - Y_VAL * HESSEN_MATRIX(i, n - 1), _
                        -K_VAL - Y_VAL * HESSEN_MATRIX(i, n), Z_VAL, Q_VAL)
                        CPLX_DIV_REAL_VAL = LBound(TEMP_ARR)
                        CPLX_DIV_IMAG_VAL = UBound(TEMP_ARR)
                        HESSEN_MATRIX(i + 1, n - 1) = CPLX_DIV_REAL_VAL
                        HESSEN_MATRIX(i + 1, n) = CPLX_DIV_IMAG_VAL
                     End If
                  End If
   
                  ' Overflow control

                  O_VAL = MAXIMUM_FUNC(Abs(HESSEN_MATRIX(i, n - 1)), Abs(HESSEN_MATRIX(i, n)))
                  If ((epsilon * O_VAL) * O_VAL > 1) Then
                     For j = i To n
                        HESSEN_MATRIX(j, n - 1) = HESSEN_MATRIX(j, n - 1) / O_VAL
                        HESSEN_MATRIX(j, n) = HESSEN_MATRIX(j, n) / O_VAL
                     Next j
                  End If
               End If 'If (IMAG_EIGEN_VALUES_MATRIX(i) < 0)
            Next i
         End If 'If (Q_VAL = 0)
      Next n
   
      ' Vectors of isolated roots
      For i = 0 To ll - 1
         If ((i < LOW_VAL) Or (i > HIGH_VAL)) Then
            For j = i To ll - 1
               EIGEN_VECTORS_MATRIX(i, j) = HESSEN_MATRIX(i, j)
            Next j
         End If
      Next i
   
      ' Back transformation to get eigenvectors of original matrix
      For j = ll - 1 To LOW_VAL Step -1
         For i = LOW_VAL To HIGH_VAL
            Z_VAL = 0
            For k = LOW_VAL To MINIMUM_FUNC(Int(j), Int(HIGH_VAL))
               Z_VAL = Z_VAL + EIGEN_VECTORS_MATRIX(i, k) * HESSEN_MATRIX(k, j)
            Next k
            EIGEN_VECTORS_MATRIX(i, j) = Z_VAL
         Next i
      Next j
1983:
'-----------------------------------------------------------------
  End If
'-----------------------------------------------------------------

'----------------------------------------------------------------------------------
  If SORT_FLAG = True Then
'----------------------------------------------------------------------------------
      ll = NSIZE
      ReDim EIGEN_ARR(1 To ll)
      For i = 1 To ll
        EIGEN_ARR(i) = REAL_EIGEN_VALUES_MATRIX(i - 1)
      Next i
      'ReDim EIGEN_VECTORS_MATRIX(1 To ll, 1 To ll)
      EIGEN_VECTORS_MATRIX = MATRIX_CHANGE_BASE_ONE_FUNC(EIGEN_VECTORS_MATRIX)
      ll = UBound(EIGEN_ARR, 1)
      ReDim TEMP_MATRIX(1 To ll, 1 To 2)
      For i = 1 To ll
        TEMP_MATRIX(i, 1) = EIGEN_ARR(i)
        TEMP_MATRIX(i, 2) = i
      Next i
      TEMP_MATRIX = MATRIX_SORT_COLUMNS_FUNC(TEMP_MATRIX, 2)
      ReDim TEMP_VECTOR(1 To ll, 1 To 2)
      mm = 1
      For i = ll To 1 Step -1
        TEMP_VECTOR(mm, 1) = TEMP_MATRIX(i, 1)
        TEMP_VECTOR(mm, 2) = TEMP_MATRIX(i, 2)
        mm = mm + 1
      Next i
      TEMP_MATRIX = TEMP_VECTOR
      ReDim EIGEN_ARR(1 To ll)
      ReDim EIGEN_VALUES_MATRIX(1 To ll, 1 To ll)
      For i = 1 To ll
        EIGEN_ARR(i) = TEMP_MATRIX(i, 1)
        h = TEMP_MATRIX(i, 2)
        For j = 1 To ll
          EIGEN_VALUES_MATRIX(j, i) = EIGEN_VECTORS_MATRIX(j, h)
        Next j
      Next i
      EIGEN_VECTORS_MATRIX = EIGEN_VALUES_MATRIX
      For i = 1 To ll
        REAL_EIGEN_VALUES_MATRIX(i - 1) = EIGEN_ARR(i)
      Next i
      EIGEN_VECTORS_MATRIX = MATRIX_CHANGE_BASE_ZERO_FUNC(EIGEN_VECTORS_MATRIX)
'----------------------------------------------------------------------------------
  End If
'----------------------------------------------------------------------------------

'----------------------------------------------------------------------------------
Select Case OUTPUT
'----------------------------------------------------------------------------------
Case 0, 1 'Get Eigen Values Matrix
'----------------------------------------------------------------------------------
  ReDim TEMP_MATRIX(0 To ll - 1, 0 To ll - 1)
  For i = 0 To ll - 1
    For j = 0 To ll - 1
      TEMP_MATRIX(i, j) = 0
    Next j
    TEMP_MATRIX(i, i) = REAL_EIGEN_VALUES_MATRIX(i)
    If (IMAG_EIGEN_VALUES_MATRIX(i) > 0) Then
      TEMP_MATRIX(i, i + 1) = IMAG_EIGEN_VALUES_MATRIX(i)
    ElseIf (IMAG_EIGEN_VALUES_MATRIX(i) < 0) Then
      TEMP_MATRIX(i, i - 1) = IMAG_EIGEN_VALUES_MATRIX(i)
    End If
  Next i
  
  If OUTPUT = 0 Then
    MATRIX_EIGEN_SQUARE_FUNC = MATRIX_CHANGE_BASE_ONE_FUNC(TEMP_MATRIX)
  Else
    MATRIX_EIGEN_SQUARE_FUNC = Array(MATRIX_CHANGE_BASE_ONE_FUNC(TEMP_MATRIX), _
                                                EIGEN_VECTORS_MATRIX, REAL_EIGEN_VALUES_MATRIX, _
                                                IMAG_EIGEN_VALUES_MATRIX)
  End If
'----------------------------------------------------------------------------------
Case 2
'----------------------------------------------------------------------------------
  MATRIX_EIGEN_SQUARE_FUNC = EIGEN_VECTORS_MATRIX 'Get Eigen Vectors Matrix
'----------------------------------------------------------------------------------
Case 3
'----------------------------------------------------------------------------------
  MATRIX_EIGEN_SQUARE_FUNC = REAL_EIGEN_VALUES_MATRIX 'Get Real Eigen Values
'----------------------------------------------------------------------------------
Case Else
'----------------------------------------------------------------------------------
  MATRIX_EIGEN_SQUARE_FUNC = IMAG_EIGEN_VALUES_MATRIX 'Get Imaginary EigenValues
'----------------------------------------------------------------------------------
End Select
'----------------------------------------------------------------------------------

Exit Function
ERROR_LABEL:
MATRIX_EIGEN_SQUARE_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : MATRIX_JACOBI_EIGEN_FUNC

'DESCRIPTION   : Returns the approx eigenvalues/eigenvectors of a symmetric matrix
'uses the fast Jacobi iterative algorithm

'LIBRARY       : MATRIX
'GROUP         : EIGEN
'ID            : 002
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/28/2009
'************************************************************************************
'************************************************************************************

Function MATRIX_JACOBI_EIGEN_FUNC(ByRef DATA_RNG As Variant, _
Optional ByVal nLOOPS As Long = 200, _
Optional ByVal OUTPUT As Integer = 1, _
Optional ByVal epsilon As Double = 2 * 10 ^ -14)

Dim i As Long
Dim j As Long
Dim l As Long

Dim ii As Long
Dim jj As Long

Dim NROWS As Long

Dim TEMP_COS As Double
Dim TEMP_SIN As Double

Dim ATEMP_VAL As Double
Dim BTEMP_VAL As Double
Dim CTEMP_VAL As Double
Dim DTEMP_VAL As Double

Dim DATA_MATRIX As Variant
Dim TEMP_MATRIX As Variant
Dim UTEMP_MATRIX As Variant
Dim WTEMP_MATRIX As Variant

On Error GoTo ERROR_LABEL

DATA_MATRIX = DATA_RNG
NROWS = UBound(DATA_MATRIX, 1)

l = 1
TEMP_MATRIX = MATRIX_IDENTITY_FUNC(NROWS)

Do Until l > nLOOPS
    l = l + 1
    'search for max value out of the first diagonal
    ATEMP_VAL = 0
    
    ii = 0
    jj = 0
    
    For i = 1 To NROWS
        For j = 1 To NROWS
            If i <> j And Abs(DATA_MATRIX(i, j)) > ATEMP_VAL Then
                ATEMP_VAL = Abs(DATA_MATRIX(i, j))
                ii = i
                jj = j
            End If
        Next j
    Next i

    If ii = 0 Then Exit Do
    
    BTEMP_VAL = DATA_MATRIX(jj, jj) - DATA_MATRIX(ii, ii)
    If BTEMP_VAL = 0 Then
        CTEMP_VAL = 1
    Else
        DTEMP_VAL = BTEMP_VAL / DATA_MATRIX(ii, jj) / 2
        CTEMP_VAL = Sgn(DTEMP_VAL) / (Abs(DTEMP_VAL) + Sqr(DTEMP_VAL ^ 2 + 1))
    End If
    
    TEMP_COS = 1 / (CTEMP_VAL ^ 2 + 1) ^ 0.5 'cosine
    TEMP_SIN = CTEMP_VAL * TEMP_COS         'sine
        
    For i = 1 To NROWS
        ATEMP_VAL = DATA_MATRIX(i, ii)
        BTEMP_VAL = DATA_MATRIX(i, jj)
        DATA_MATRIX(i, ii) = TEMP_COS * ATEMP_VAL - TEMP_SIN * BTEMP_VAL
        DATA_MATRIX(i, jj) = TEMP_SIN * ATEMP_VAL + TEMP_COS * BTEMP_VAL
    Next i

    For j = 1 To NROWS
        ATEMP_VAL = DATA_MATRIX(ii, j)
        BTEMP_VAL = DATA_MATRIX(jj, j)
        DATA_MATRIX(ii, j) = TEMP_COS * ATEMP_VAL - TEMP_SIN * BTEMP_VAL
        DATA_MATRIX(jj, j) = TEMP_SIN * ATEMP_VAL + TEMP_COS * BTEMP_VAL
    Next j
    
    ReDim UTEMP_MATRIX(1 To NROWS, 1 To NROWS)
    ReDim WTEMP_MATRIX(1 To NROWS, 1 To NROWS)
    
    For i = 1 To NROWS
        UTEMP_MATRIX(i, i) = 1
    Next i
    
    UTEMP_MATRIX(ii, ii) = TEMP_COS
    UTEMP_MATRIX(ii, jj) = TEMP_SIN
    UTEMP_MATRIX(jj, ii) = -TEMP_SIN
    UTEMP_MATRIX(jj, jj) = TEMP_COS
    TEMP_MATRIX = MMULT_FUNC(TEMP_MATRIX, UTEMP_MATRIX, 70)
Loop

Select Case OUTPUT
Case 0 'Eigen Values
    MATRIX_JACOBI_EIGEN_FUNC = _
        MATRIX_TRIM_SMALL_VALUES_FUNC(DATA_MATRIX, epsilon)
Case 1 'EigenVectors
    MATRIX_JACOBI_EIGEN_FUNC = TEMP_MATRIX
Case Else
    MATRIX_JACOBI_EIGEN_FUNC = Array(MATRIX_TRIM_SMALL_VALUES_FUNC(DATA_MATRIX, epsilon), TEMP_MATRIX)
End Select

Exit Function
ERROR_LABEL:
MATRIX_JACOBI_EIGEN_FUNC = Err.number
End Function
'-----------------------------------------------------------------------------------------------

'************************************************************************************
'************************************************************************************
'FUNCTION      : MATRIX_JACOBI_EIGEN_SORT_FUNC
'DESCRIPTION   : This function perform the sort of the eigenvalues and returns
'the eigenvectors associated to the absolute highest eigenvalues.
'EigvalM is the diagonal eigenvalues (n x n ) matrix and EigvectM
'is the (n x n ) eigenvector unitary matrix.

'LIBRARY       : MATRIX
'GROUP         : EIGEN
'ID            : 003
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/28/2009
'************************************************************************************
'************************************************************************************


Function MATRIX_JACOBI_EIGEN_SORT_FUNC(ByRef EIGEN_VALUES_RNG As Variant, _
ByRef EIGEN_VECTORS_RNG As Variant)

Dim i As Long
Dim j As Long
Dim k As Long
Dim NROWS As Long

Dim ATEMP_ARR As Variant
Dim BTEMP_ARR As Variant

Dim TEMP_MATRIX As Variant
Dim EIGEN_VALUES_MATRIX As Variant
Dim EIGEN_VECTORS_MATRIX As Variant

On Error GoTo ERROR_LABEL

EIGEN_VALUES_MATRIX = EIGEN_VALUES_RNG ' eigen values matrix
NROWS = UBound(EIGEN_VALUES_MATRIX, 1)

EIGEN_VECTORS_MATRIX = EIGEN_VECTORS_RNG ' eigen vectors matrix

ReDim ATEMP_ARR(1 To NROWS)
ReDim BTEMP_ARR(1 To NROWS)

For i = 1 To NROWS ' extract eigen values
    ATEMP_ARR(i) = EIGEN_VALUES_MATRIX(i, i)
    BTEMP_ARR(i) = i
Next i

For i = 1 To NROWS
    For j = i + 1 To NROWS
        If Abs(ATEMP_ARR(BTEMP_ARR(j))) > _
                Abs(ATEMP_ARR(BTEMP_ARR(i))) Then
            k = BTEMP_ARR(i)
            BTEMP_ARR(i) = BTEMP_ARR(j)
            BTEMP_ARR(j) = k
        End If
    Next j
Next i

' now populate result matrix
ReDim TEMP_MATRIX(1 To NROWS, 1 To NROWS)
For i = 1 To NROWS
    For j = 1 To NROWS
        TEMP_MATRIX(j, i) = EIGEN_VECTORS_MATRIX(j, BTEMP_ARR(i))
    Next j
Next i
MATRIX_JACOBI_EIGEN_SORT_FUNC = TEMP_MATRIX
      
Exit Function
ERROR_LABEL:
MATRIX_JACOBI_EIGEN_SORT_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : MATRIX_EIGENVECTOR_FUNC
'DESCRIPTION   : This function returns the eigenvector associated with the given
'eigenvalue of a matrix A (n x n)

'If Eigenvalues is a single value, the function returns a (n x 1)
'vector. Otherwise if Eigenvalues is a vector of all eigenvalues of
'matrix A, the function returns a matrix (n x n) of eigenvector.
'Note: the eigenvector returned by this function is not normalized.
'The optional parameter epsilon  is useful only if your eigenvalues
'are affected by error. In that case the epsilon should be proportionally
'adapted. Otherwise the result may be a NULL matrix.  If omitted, the
'function tries to detect by itself the best error parameter for the
'approximate eigenvalues.

'In case of difficult eigenvalue affected by poor accuracy, it is better
'to use the alternative inverse iterative method performed by the function

'LIBRARY       : MATRIX
'GROUP         : EIGEN
'ID            : 004
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/28/2009
'************************************************************************************
'************************************************************************************

Function MATRIX_EIGENVECTOR_FUNC(ByRef DATA_RNG As Variant, _
ByRef EIGEN_RNG As Variant, _
Optional ByVal epsilon As Double = 2 * 10 ^ -15)

Dim i As Long
Dim j As Long
Dim k As Long
Dim l As Long

Dim NSIZE As Long
Dim NROWS As Long
Dim NCOLUMNS As Long

Dim TEMP_SGN As Double
Dim TEMP_VAL As Double '

Dim DATA_MATRIX As Variant
Dim UTEMP_MATRIX As Variant
Dim WTEMP_MATRIX As Variant
Dim EIGEN_MATRIX As Variant

On Error GoTo ERROR_LABEL

DATA_MATRIX = DATA_RNG
EIGEN_MATRIX = EIGEN_RNG
If IsArray(EIGEN_MATRIX) Then
    If UBound(EIGEN_MATRIX, 1) = 1 Then
        EIGEN_MATRIX = MATRIX_TRANSPOSE_FUNC(EIGEN_MATRIX)
    End If
    NSIZE = UBound(EIGEN_MATRIX, 1)
Else
    NSIZE = 1
End If

NROWS = UBound(DATA_MATRIX, 1)
ReDim WTEMP_MATRIX(1 To NROWS, 1 To NROWS)
TEMP_VAL = 0
k = 1
Do Until k > NSIZE
    If IsArray(EIGEN_MATRIX) Then
        TEMP_VAL = EIGEN_MATRIX(k, 1)
    Else
        TEMP_VAL = EIGEN_MATRIX
    End If
    If k > 1 Then DATA_MATRIX = DATA_RNG  'reinitialize
    For i = 1 To NROWS
        DATA_MATRIX(i, i) = DATA_MATRIX(i, i) - TEMP_VAL
    Next i
    
    UTEMP_MATRIX = MATRIX_GS_SINGULAR_LINEAR_SYSTEM_FUNC(DATA_MATRIX, , 0, epsilon, 0)
    l = 0
    For i = 1 To NROWS
        '      COUNTER=0 => non singular matrix
        '      COUNTER=1 => eigenvalue simple
        '      COUNTER>1 => eigenvalue multiple
        If UTEMP_MATRIX(i, i) = 1 Then l = l + 1
    Next i
    
    If l = 0 Then 'matrix not singular
        MATRIX_EIGENVECTOR_FUNC = UTEMP_MATRIX
        Exit Function
    End If
    
    For j = 1 To NROWS
        If UTEMP_MATRIX(j, j) = 1 Then
            For i = 1 To NROWS
                WTEMP_MATRIX(i, k) = UTEMP_MATRIX(i, j)
            Next i
            k = k + 1
        End If
    Next j
Loop

'normalize the sign of each eigenvector making positive the first
'non zero element  |aij| > epsilon

NROWS = UBound(WTEMP_MATRIX, 1)
NCOLUMNS = UBound(WTEMP_MATRIX, 2)

For j = 1 To NCOLUMNS
    TEMP_SGN = 0
    For i = 1 To NROWS
        If Abs(WTEMP_MATRIX(i, j)) > 1000 * epsilon Then
            If TEMP_SGN = 0 Then TEMP_SGN = Sgn(WTEMP_MATRIX(i, j))
            If TEMP_SGN < 0 Then
                WTEMP_MATRIX(i, j) = -WTEMP_MATRIX(i, j)
            Else
                Exit For 'exit inner for
            End If
        End If
    Next i
Next j

WTEMP_MATRIX = MATRIX_NORMALIZED_VECTOR_FUNC(WTEMP_MATRIX, 2, epsilon)
MATRIX_EIGENVECTOR_FUNC = WTEMP_MATRIX

Exit Function
ERROR_LABEL:
MATRIX_EIGENVECTOR_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : MATRIX_QR_EIGENVALUES_FUNC

'DESCRIPTION   : Find real and complex eigenvalues with the iterative QR method

'This function performs the diagonal reduction of a matrix with
'QR method, returning the approximate the n x 2 array of eigenvalues,
'real or complex. The example below show that the given matrix has
'two complex conjugate eigenvalues and only one real eigenvalue.

'Being a symmetric there are only n real distinct eigenvalues. So the
'function returns only an array n x 1

'References

'This function uses a reduction of the LAPACK FORTRAN HQR and ELMHES
'subroutines (April 1983)

'---------------------------------------------------------------------------
'HQR IS A TRANSLATION OF THE ALGOL PROCEDURE
'NUM. MATH. 14, 219-231(1970) BY MARTIN, PETERS, AND WILKINSON.
'HANDBOOK FOR AUTO. COMP., VOL.II-LINEAR ALGEBRA, 359-371(1971).

'ELMHES IS A TRANSLATION OF THE ALGOL PROCEDURE,
'NUM. MATH. 12, 349-368(1968) BY MARTIN AND WILKINSON.
'HANDBOOK FOR AUTO. COMP., VOL.II-LINEAR ALGEBRA, 339-358(1971).
'---------------------------------------------------------------------------

'LIBRARY       : MATRIX
'GROUP         : EIGEN
'ID            : 005
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/28/2009
'************************************************************************************
'************************************************************************************

Function MATRIX_QR_EIGENVALUES_FUNC(ByRef DATA_RNG As Variant, _
Optional ByVal SORT_FLAG As Boolean = True, _
Optional ByVal epsilon As Double = 2 * 10 ^ -15)

Dim h As Long '
Dim i As Long
Dim j As Long
Dim k As Long
Dim l As Long '
Dim m As Long '
Dim n As Long '

Dim hh As Long '
Dim ii As Long
Dim jj As Long '
Dim kk As Long '
Dim ll As Long '

Dim NSIZE As Long
Dim nLOOPS As Long
Dim COUNTER As Long

Dim P_VAL As Double
Dim Q_VAL As Double
Dim R_VAL As Double
Dim S_VAL As Double
Dim T_VAL As Double
Dim W_VAL As Double
Dim X_VAL As Double
Dim Y_VAL As Double
Dim Z_VAL As Double

Dim ATEMP_SUM As Double
Dim BTEMP_SUM As Double

Dim XTEMP_VAL As Double
Dim YTEMP_VAL As Double
Dim ZTEMP_VAL As Double

Dim DATA_MATRIX As Variant
Dim TEMP_MATRIX As Variant

Dim NOTLAS_FLAG As Boolean

On Error GoTo ERROR_LABEL

NOTLAS_FLAG = False

DATA_MATRIX = DATA_RNG
NSIZE = UBound(DATA_MATRIX, 1)

ReDim RTEMP_VECTOR(1 To NSIZE)
ReDim ITEMP_VECTOR(1 To NSIZE)

      For k = 2 To NSIZE - 1
         i = k
         XTEMP_VAL = 0
         For j = k To NSIZE
            If (Abs(DATA_MATRIX(j, k - 1)) > Abs(XTEMP_VAL)) Then
               XTEMP_VAL = DATA_MATRIX(j, k - 1)
               i = j
            End If
         Next j
         If (i <> k) Then
            For j = k - 1 To NSIZE
                   ZTEMP_VAL = DATA_MATRIX(i, j)
                   DATA_MATRIX(i, j) = DATA_MATRIX(k, j)
                   DATA_MATRIX(k, j) = ZTEMP_VAL
            Next j
            For j = 1 To NSIZE
                   ZTEMP_VAL = DATA_MATRIX(j, i)
                   DATA_MATRIX(j, i) = DATA_MATRIX(j, k)
                   DATA_MATRIX(j, k) = ZTEMP_VAL
            Next j
         End If
         If (XTEMP_VAL <> 0) Then
            For i = k + 1 To NSIZE
               YTEMP_VAL = DATA_MATRIX(i, k - 1)
               If (YTEMP_VAL <> 0) Then
                  YTEMP_VAL = YTEMP_VAL / XTEMP_VAL
                  DATA_MATRIX(i, k - 1) = YTEMP_VAL
                  For j = k To NSIZE
                     DATA_MATRIX(i, j) = DATA_MATRIX(i, j) - _
                        YTEMP_VAL * DATA_MATRIX(k, j)
                  Next j
                  For j = 1 To NSIZE
                     DATA_MATRIX(j, k) = DATA_MATRIX(j, k) + _
                        YTEMP_VAL * DATA_MATRIX(j, i)
                  Next j
               End If
            Next i
         End If
      Next k


'----------------------------------------------------------------------------
'     THIS SUBROUTINE FINDS THE EIGENVALUES OF A REAL UPPER HESSENBERG
'     MATRIX BY THE QR METHOD.
'
'     THIS SUBROUTINE IS A TRANSLATION OF THE ALGOL PROCEDURE HQR,
'     NUM. MATH. 14, 219-231(1970) BY MARTIN, PETERS, AND WILKINSON.
'     HANDBOOK FOR AUTO. COMP., VOL.II-LINEAR ALGEBRA, 359-371(1971).
'
'        DATA_MATRIX: HAS BEEN DESTROYED.  THEREFORE, IT MUST BE SAVED
'          BEFORE CALLING  HQR  IF SUBSEQUENT CALCULATION AND
'          BACK TRANSFORMATION OF EIGENVECTORS IS TO BE PERFORMED.
'
'        RTEMP_VECTOR AND ITEMP_VECTOR: CONTAIN THE REAL AND IMAGINARY PARTS,
'          RESPECTIVELY, OF THE EIGENVALUES.  THE EIGENVALUES
'          ARE UNORDERED EXCEPT THAT COMPLEX CONJUGATE PAIRS
'          OF VALUES APPEAR CONSECUTIVELY WITH THE EIGENVALUE
'          HAVING THE POSITIVE IMAGINARY PART FIRST.  IF AN
'          ERROR EXIT IS MADE, THE EIGENVALUES SHOULD BE CORRECT
'          FOR INDICES IERR+1,...,NSIZE.
'
'        COUNTER: IS SET TO
'          ZERO       FOR NORMAL RETURN,
'          J          IF THE LIMIT OF 30*NSIZE ITERATIONS IS EXHAUSTED
'                     WHILE THE J-TH EIGENVALUE IS BEING SOUGHT.
'------------------------------------------------------------------------------


      COUNTER = 0
      k = 1
      
      ii = NSIZE
      T_VAL = 0#
      nLOOPS = 30 * NSIZE
'     .......... SEARCH FOR NEXT EIGENVALUES ..........
1983:
      If (ii < 1) Then GoTo 1999
      kk = 0
      n = ii - 1
      hh = n - 1
'     .......... LOOK FOR SINGLE SMALL SUB-DIAGONAL ELEMENT
1984:
     For ll = 1 To ii
            l = ii + 1 - ll
            If (l = 1) Then GoTo 1985
            S_VAL = Abs(DATA_MATRIX(l - 1, l - 1)) + Abs(DATA_MATRIX(l, l))
            If (S_VAL = 0) Then S_VAL = 1
            ATEMP_SUM = S_VAL
            BTEMP_SUM = ATEMP_SUM + Abs(DATA_MATRIX(l, l - 1))
            If (BTEMP_SUM = ATEMP_SUM) And Abs(DATA_MATRIX(l, l - 1)) < 1 _
                Then GoTo 1985
     Next ll
'     .......... FORM SHIFT ..........
1985:
      X_VAL = DATA_MATRIX(ii, ii)
      If (l = ii) Then GoTo 1994
      Y_VAL = DATA_MATRIX(n, n)
      W_VAL = DATA_MATRIX(ii, n) * DATA_MATRIX(n, ii)
      If (l = n) Then GoTo 1995
      If (nLOOPS = 0) Then GoTo 1998
      If ((kk <> 10) And (kk <> 20)) Then GoTo 1986
'     .......... FORM EXCEPTIONAL SHIFT ..........
      T_VAL = T_VAL + X_VAL
'
      For i = 1 To ii
        DATA_MATRIX(i, i) = DATA_MATRIX(i, i) - X_VAL
      Next i
'
      S_VAL = Abs(DATA_MATRIX(ii, n)) + Abs(DATA_MATRIX(n, hh))
      X_VAL = 0.75 * S_VAL
      Y_VAL = X_VAL
      W_VAL = -0.4375 * S_VAL * S_VAL
1986:
      kk = kk + 1
      nLOOPS = nLOOPS - 1
'     .......... LOOK FOR TWO CONSECUTIVE SMALL
'                SUB-DIAGONAL ELEMENTS.
      For jj = l To hh
         m = hh + l - jj
         Z_VAL = DATA_MATRIX(m, m)
         R_VAL = X_VAL - Z_VAL
         S_VAL = Y_VAL - Z_VAL
         P_VAL = (R_VAL * S_VAL - W_VAL) / DATA_MATRIX(m + 1, m) + DATA_MATRIX(m, m + 1)
         Q_VAL = DATA_MATRIX(m + 1, m + 1) - Z_VAL - R_VAL - S_VAL
         R_VAL = DATA_MATRIX(m + 2, m + 1)
         S_VAL = Abs(P_VAL) + Abs(Q_VAL) + Abs(R_VAL)
         P_VAL = P_VAL / S_VAL
         Q_VAL = Q_VAL / S_VAL
         R_VAL = R_VAL / S_VAL
         If (m = l) Then GoTo 1987
         ATEMP_SUM = Abs(P_VAL) * (Abs(DATA_MATRIX(m - 1, m - 1)) + Abs(Z_VAL) + _
            Abs(DATA_MATRIX(m + 1, m + 1)))
         BTEMP_SUM = ATEMP_SUM + Abs(DATA_MATRIX(m, m - 1)) * (Abs(Q_VAL) + Abs(R_VAL))
         If (BTEMP_SUM = ATEMP_SUM) Then GoTo 1987
      Next jj
'
1987:
      h = m + 2
'
      For i = h To ii
         DATA_MATRIX(i, i - 2) = 0#
         If (i <> h) Then DATA_MATRIX(i, i - 3) = 0#
      Next i
'     .......... DOUBLE QR STEP INVOLVING ROWS L TO ii AND
'                COLUMNS M TO ii ..........
      For k = m To n
         NOTLAS_FLAG = k <> n
         If (k = m) Then GoTo 1988
         P_VAL = DATA_MATRIX(k, k - 1)
         Q_VAL = DATA_MATRIX(k + 1, k - 1)
         R_VAL = 0#
         If (NOTLAS_FLAG) Then R_VAL = DATA_MATRIX(k + 2, k - 1)
         X_VAL = Abs(P_VAL) + Abs(Q_VAL) + Abs(R_VAL)
         If (X_VAL = 0#) Then GoTo 1993
         P_VAL = P_VAL / X_VAL
         Q_VAL = Q_VAL / X_VAL
         R_VAL = R_VAL / X_VAL
1988:
         
         S_VAL = IIf(P_VAL >= 0, Abs(Sqr(P_VAL * P_VAL + Q_VAL * Q_VAL + R_VAL * R_VAL)), _
                        -Abs(Sqr(P_VAL * P_VAL + Q_VAL * Q_VAL + R_VAL * R_VAL)))
         
         If (k = m) Then GoTo 1989
         DATA_MATRIX(k, k - 1) = -S_VAL * X_VAL
         GoTo 1990
1989:
         If (l <> m) Then DATA_MATRIX(k, k - 1) = -DATA_MATRIX(k, k - 1)
1990:
         P_VAL = P_VAL + S_VAL
         X_VAL = P_VAL / S_VAL
         Y_VAL = Q_VAL / S_VAL
         Z_VAL = R_VAL / S_VAL
         Q_VAL = Q_VAL / P_VAL
         R_VAL = R_VAL / P_VAL
         If (NOTLAS_FLAG) Then GoTo 1991
'     .......... ROW MODIFICATION ..........
         For j = k To NSIZE
            P_VAL = DATA_MATRIX(k, j) + Q_VAL * DATA_MATRIX(k + 1, j)
            DATA_MATRIX(k, j) = DATA_MATRIX(k, j) - P_VAL * X_VAL
            DATA_MATRIX(k + 1, j) = DATA_MATRIX(k + 1, j) - P_VAL * Y_VAL
         Next j
'
         j = MINIMUM_FUNC(ii, k + 3)
'     .......... COLUMN MODIFICATION ..........
         For i = 1 To j
            P_VAL = X_VAL * DATA_MATRIX(i, k) + Y_VAL * DATA_MATRIX(i, k + 1)
            DATA_MATRIX(i, k) = DATA_MATRIX(i, k) - P_VAL
            DATA_MATRIX(i, k + 1) = DATA_MATRIX(i, k + 1) - P_VAL * Q_VAL
         Next i
         GoTo 1992
1991:
'     .......... ROW MODIFICATION ..........
         For j = k To NSIZE
            P_VAL = DATA_MATRIX(k, j) + Q_VAL * DATA_MATRIX(k + 1, j) + _
                R_VAL * DATA_MATRIX(k + 2, j)
            DATA_MATRIX(k, j) = DATA_MATRIX(k, j) - P_VAL * X_VAL
            DATA_MATRIX(k + 1, j) = DATA_MATRIX(k + 1, j) - P_VAL * Y_VAL
            DATA_MATRIX(k + 2, j) = DATA_MATRIX(k + 2, j) - P_VAL * Z_VAL
         Next j
'
         j = MINIMUM_FUNC(ii, k + 3)
'     .......... COLUMN MODIFICATION ..........
         For i = 1 To j
            P_VAL = X_VAL * DATA_MATRIX(i, k) + Y_VAL * DATA_MATRIX(i, k + 1) + _
                Z_VAL * DATA_MATRIX(i, k + 2)
            DATA_MATRIX(i, k) = DATA_MATRIX(i, k) - P_VAL
            DATA_MATRIX(i, k + 1) = DATA_MATRIX(i, k + 1) - P_VAL * Q_VAL
            DATA_MATRIX(i, k + 2) = DATA_MATRIX(i, k + 2) - P_VAL * R_VAL
         Next i
1992:
'
      Next k
1993:
'
      GoTo 1984
'     .......... ONE ROOT FOUND ..........
1994:
      RTEMP_VECTOR(ii) = X_VAL + T_VAL
      ITEMP_VECTOR(ii) = 0#
      ii = n
      GoTo 1983
'     .......... TWO ROOTS FOUND ..........
1995:
      P_VAL = (Y_VAL - X_VAL) / 2#
      Q_VAL = P_VAL * P_VAL + W_VAL
      Z_VAL = Sqr(Abs(Q_VAL))
      X_VAL = X_VAL + T_VAL
      If (Q_VAL < 0#) Then GoTo 1996
'     .......... REAL PAIR ..........
      Z_VAL = P_VAL + IIf(P_VAL >= 0, Abs(Z_VAL), -Abs(Z_VAL))
      
      
      RTEMP_VECTOR(n) = X_VAL + Z_VAL
      RTEMP_VECTOR(ii) = RTEMP_VECTOR(n)
      If (Z_VAL <> 0#) Then RTEMP_VECTOR(ii) = X_VAL - W_VAL / Z_VAL
      ITEMP_VECTOR(n) = 0#
      ITEMP_VECTOR(ii) = 0#
      GoTo 1997
'     .......... COMPLEX PAIR ..........
1996:
      RTEMP_VECTOR(n) = X_VAL + P_VAL
      RTEMP_VECTOR(ii) = X_VAL + P_VAL
      ITEMP_VECTOR(n) = Z_VAL
      ITEMP_VECTOR(ii) = -Z_VAL
1997:
      ii = hh
      GoTo 1983
'     .......... SET ERROR -- ALL EIGENVALUES HAVE NOT
'                CONVERGED AFTER 30*NSIZE ITERATIONS ..........
1998:
      COUNTER = ii
1999:

ReDim TEMP_MATRIX(1 To NSIZE, 1 To 2)

For i = 1 To NSIZE
    If i > COUNTER Then
        TEMP_MATRIX(i, 1) = RTEMP_VECTOR(i)
        TEMP_MATRIX(i, 2) = ITEMP_VECTOR(i)
    Else
        TEMP_MATRIX(i, 1) = "-"
        TEMP_MATRIX(i, 2) = "-"
    End If
Next i

TEMP_MATRIX = MATRIX_TRIM_SMALL_VALUES_FUNC(TEMP_MATRIX, epsilon)
If SORT_FLAG = True Then: TEMP_MATRIX = MATRIX_QUICK_SORT_FUNC(TEMP_MATRIX, 1, 1)

MATRIX_QR_EIGENVALUES_FUNC = TEMP_MATRIX

Exit Function
ERROR_LABEL:
MATRIX_QR_EIGENVALUES_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : MATRIX_QL_EIGENVALUES_TRIDIAGONAL_FUNC

'DESCRIPTION   : This function returns the real eigenvalues of a tridiagonal
'symmetric matrix. It works also for an asymmetrical tridiagonal
'matrix having all real eigenvalues.

'The optional parameter Itermax sets the max number of iteration
'allowed (default nLOOPS =200). This function uses the efficient
'QL algorithm. If the matrix does not have all-real eigenvalues,
'this function returns "-". This function accepts tridiagonal
'matrices in both square (n x n) and (n x 3 ) rectangular formats.
'Note that the rectangular form (n x 3) is very useful for large matrices.

'LIBRARY       : MATRIX
'GROUP         : EIGEN
'ID            : 006
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/28/2009
'************************************************************************************
'************************************************************************************

Function MATRIX_QL_EIGENVALUES_TRIDIAGONAL_FUNC(ByRef DATA_RNG As Variant, _
Optional ByVal nLOOPS As Long = 200)

Dim i As Long
Dim j As Long
Dim k As Long
Dim l As Long

Dim NROWS As Long

Dim ATEMP_VAL As Double
Dim BTEMP_VAL As Double
Dim CTEMP_VAL As Double
Dim DTEMP_VAL As Double
Dim ETEMP_VAL As Double
Dim FTEMP_VAL As Double

Dim UPPER_ARR As Variant
Dim LOWER_ARR As Variant
Dim DIAGONAL_ARR As Variant

Dim DATA_MATRIX As Variant
Dim WTEMP_MATRIX As Variant

On Error GoTo ERROR_LABEL

DATA_MATRIX = DATA_RNG
DATA_MATRIX = MATRIX_TRIDIAGONAL_LOAD_FUNC(DATA_MATRIX)
DATA_MATRIX = MATRIX_TRIDIAGONAL_SYMMETRIZE_FUNC(DATA_MATRIX)

NROWS = UBound(DATA_MATRIX, 1)

ReDim UPPER_ARR(1 To NROWS)
ReDim LOWER_ARR(1 To NROWS)
ReDim DIAGONAL_ARR(1 To NROWS)

For i = 1 To NROWS
    LOWER_ARR(i) = DATA_MATRIX(i, 1)
    DIAGONAL_ARR(i) = DATA_MATRIX(i, 2)
    UPPER_ARR(i) = DATA_MATRIX(i, 3)
Next i

'-----------------------------------------------------------------------------
'The following subroutines find eigenvalues and eigenvector of
'symmetric tridiagonal matrix with QL algorithm at the end diag
'will contains the eigenvalues
'-----------------------------------------------------------------------------

For j = 1 To NROWS
  l = 0
  Do Until l > nLOOPS
      For k = j To NROWS - 1
          FTEMP_VAL = Abs(DIAGONAL_ARR(k)) + Abs(DIAGONAL_ARR(k + 1))
          If (Abs(UPPER_ARR(k)) + FTEMP_VAL = FTEMP_VAL) Then Exit For
      Next k
      If (k = j) Then Exit Do  '--> exit loop
      l = l + 1
      ATEMP_VAL = (DIAGONAL_ARR(j + 1) - DIAGONAL_ARR(j)) / (2# * UPPER_ARR(j))
      CTEMP_VAL = Sqr(ATEMP_VAL ^ 2 + 1)
      WTEMP_MATRIX = Abs(CTEMP_VAL)
      If ATEMP_VAL < 0 Then WTEMP_MATRIX = -WTEMP_MATRIX
      ATEMP_VAL = DIAGONAL_ARR(k) - DIAGONAL_ARR(j) + _
                UPPER_ARR(j) / (ATEMP_VAL + WTEMP_MATRIX)
      DTEMP_VAL = 1
      ETEMP_VAL = 1
      FTEMP_VAL = 0
      For i = k - 1 To j Step -1
            BTEMP_VAL = DTEMP_VAL * UPPER_ARR(i)
            DATA_MATRIX = ETEMP_VAL * UPPER_ARR(i)
            CTEMP_VAL = Sqr(BTEMP_VAL ^ 2 + ATEMP_VAL ^ 2)
            UPPER_ARR(i + 1) = CTEMP_VAL
            If (CTEMP_VAL = 0) Then
                DIAGONAL_ARR(i + 1) = DIAGONAL_ARR(i + 1) - FTEMP_VAL
                UPPER_ARR(k) = 0#
                Exit For
            End If
            DTEMP_VAL = BTEMP_VAL / CTEMP_VAL
            ETEMP_VAL = ATEMP_VAL / CTEMP_VAL
            ATEMP_VAL = DIAGONAL_ARR(i + 1) - FTEMP_VAL
            CTEMP_VAL = (DIAGONAL_ARR(i) - ATEMP_VAL) * DTEMP_VAL + 2# * ETEMP_VAL * DATA_MATRIX
            FTEMP_VAL = DTEMP_VAL * CTEMP_VAL
            DIAGONAL_ARR(i + 1) = ATEMP_VAL + FTEMP_VAL
            ATEMP_VAL = ETEMP_VAL * CTEMP_VAL - DATA_MATRIX
      Next i
      If CTEMP_VAL <> 0 Then
        DIAGONAL_ARR(j) = DIAGONAL_ARR(j) - FTEMP_VAL
        UPPER_ARR(j) = ATEMP_VAL
      End If
      UPPER_ARR(k) = 0#
  Loop
  'check convergency
  If l > nLOOPS Then: GoTo ERROR_LABEL 'convergency fails
Next j

ReDim DATA_MATRIX(1 To NROWS, 1 To 1)
For i = 1 To NROWS
    DATA_MATRIX(i, 1) = DIAGONAL_ARR(i)
Next i

DATA_MATRIX = MATRIX_QUICK_SORT_FUNC(DATA_MATRIX, 1, 1)

MATRIX_QL_EIGENVALUES_TRIDIAGONAL_FUNC = DATA_MATRIX

Exit Function
ERROR_LABEL:
MATRIX_QL_EIGENVALUES_TRIDIAGONAL_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : MATRIX_TOEPLITZ_EIGENVALUES_TRIDIAGONAL_FUNC
'DESCRIPTION   : Returns all eigenvalues of a tridiagonal toeplitz matrix n x n .
'It has been demonstrated that these matrices:
'·   for n even,  have all eigenvalues real if a*c>0; all eigenvalues
'    complex otherwise.
'·   for n odd - have all n-1 eigenvalues real if a*c>0; all n-1 eigenvalues
'    complex otherwise. The last n eigenvalue is always b.
'Example. Find the eigenvalues of the 40 x 40 tridiagonal toeplitz matrix
'having a = 1, b = 3, c = 2
'Because a*c = 2 > 0 , then all eigenvalues are real.

'LIBRARY       : MATRIX
'GROUP         : EIGEN
'ID            : 007
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/28/2009
'************************************************************************************
'************************************************************************************

Function MATRIX_TOEPLITZ_EIGENVALUES_TRIDIAGONAL_FUNC(ByVal NSIZE As Long, _
ByVal A_VAL As Double, _
ByVal B_VAL As Double, _
ByVal C_VAL As Double)

Dim i As Long
Dim j As Long
Dim k As Long

Dim PI_VAL As Double
Dim ATEMP_VAL As Double
Dim BTEMP_VAL As Double
Dim TEMP_MATRIX As Variant

On Error GoTo ERROR_LABEL

PI_VAL = 3.14159265358979
ReDim TEMP_MATRIX(1 To NSIZE, 1 To 2)
If A_VAL * C_VAL = 0 Then
    For k = 1 To NSIZE
        TEMP_MATRIX(k, 1) = B_VAL
        TEMP_MATRIX(k, 2) = 0
    Next k
Else 'one eigenvalue with multiplicity NSIZE
    ATEMP_VAL = 2 * Sqr(Abs(A_VAL * C_VAL))
    For k = 1 To Int(NSIZE / 2)
        BTEMP_VAL = ATEMP_VAL / Sqr(1 + (Tan(k * PI_VAL / (NSIZE + 1))) ^ 2)
        j = NSIZE + 1 - k
        i = k
        If A_VAL * C_VAL > 0 Then   'all eigenvalues are real
            TEMP_MATRIX(i, 1) = B_VAL - BTEMP_VAL
            TEMP_MATRIX(i, 2) = 0
            TEMP_MATRIX(j, 1) = B_VAL + BTEMP_VAL
            TEMP_MATRIX(j, 2) = 0
        Else                'all eigenvalues are complex
            TEMP_MATRIX(i, 1) = B_VAL
            TEMP_MATRIX(i, 2) = -BTEMP_VAL
            TEMP_MATRIX(j, 1) = B_VAL
            TEMP_MATRIX(j, 2) = BTEMP_VAL
        End If
    Next k
    If NSIZE Mod 2 <> 0 Then 'odd matrix
        j = (NSIZE + 1) / 2
        TEMP_MATRIX(j, 1) = B_VAL
        TEMP_MATRIX(j, 2) = 0
    End If
End If

MATRIX_TOEPLITZ_EIGENVALUES_TRIDIAGONAL_FUNC = TEMP_MATRIX

Exit Function
ERROR_LABEL:
MATRIX_TOEPLITZ_EIGENVALUES_TRIDIAGONAL_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : MATRIX_EIGENVECTOR_TRIDIAGONAL_FUNC

'DESCRIPTION   : 'This function returns the eigenvector associated with the given
'eigenvalue of a tridiagonal matrix a. If Eigenvalues is a single
'value, the function returns an (n x 1) eigenvector. Otherwise if
'Eigenvalues is a vector of all eigenvalues, the function returns
'the (n x n) matrix of eigenvectors.

'Note: the eigenvectors returned by this function are not normalized.
'The optional parameter tolerance  is useful only if your eigenvalues are
'affected by an error. In that case the tolerance should be proportionally
'adapted. Otherwise the result may be a NULL matrix.  If omitted, the
'function tries to detect by itself the best error parameter for the
'approximate eigenvalues.

'This function accepts both tridiagonal square (n x n) matrices and
'(n x 3 ) rectangular matrices.
'The second form is useful for large matrices.

'Example.
'Given the 19 x 19 tridiagonal matrix having eigenvalue EIGEN_VECTOR = 1, find its
'associate eigenvector.

'LIBRARY       : MATRIX
'GROUP         : EIGEN
'ID            : 008
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/28/2009
'************************************************************************************
'************************************************************************************

Function MATRIX_EIGENVECTOR_TRIDIAGONAL_FUNC(ByRef DATA_RNG As Variant, _
ByRef EIGEN_RNG As Variant, _
Optional ByVal tolerance As Double = 0)

Dim i As Long
Dim j As Long
Dim k As Long
Dim l As Long

Dim ii As Long
Dim jj As Long

Dim NSIZE As Long
Dim NROWS As Long
Dim NCOLUMNS As Long

Dim TEMP_DET As Double 'Variant
Dim TEMP_SGN As Double
Dim TEMP_CHECK As Double

Dim ATEMP_VAL As Double
Dim BTEMP_VAL As Double

Dim ATEMP_VECTOR As Variant
Dim BTEMP_VECTOR As Variant

Dim ATEMP_MATRIX As Variant
Dim BTEMP_MATRIX As Variant
Dim CTEMP_MATRIX As Variant
Dim DTEMP_MATRIX As Variant

Dim EIGENVALUES_MATRIX As Variant

Dim epsilon As Double
Dim DETERM_FLAG As Boolean

On Error GoTo ERROR_LABEL

ATEMP_MATRIX = MATRIX_TRIDIAGONAL_LOAD_FUNC(DATA_RNG)
BTEMP_MATRIX = ATEMP_MATRIX
EIGENVALUES_MATRIX = EIGEN_RNG

If IsArray(EIGENVALUES_MATRIX) Then
    If UBound(EIGENVALUES_MATRIX, 1) = 1 Then
        EIGENVALUES_MATRIX = MATRIX_TRANSPOSE_FUNC(EIGENVALUES_MATRIX)
    End If
    ii = UBound(EIGENVALUES_MATRIX)
Else
    ii = 1
End If

NSIZE = UBound(BTEMP_MATRIX, 1)
ReDim DTEMP_MATRIX(1 To NSIZE, 1 To NSIZE)
ReDim CTEMP_MATRIX(1 To NSIZE, 1 To 1)

ATEMP_VAL = 0
k = 1
Do Until k > ii
    If IsArray(EIGENVALUES_MATRIX) Then
        ATEMP_VAL = EIGENVALUES_MATRIX(k, 1)
    Else
        ATEMP_VAL = EIGENVALUES_MATRIX
    End If
    If k > 1 Then BTEMP_MATRIX = ATEMP_MATRIX  'reinitialize
    
    For i = 1 To NSIZE
        BTEMP_MATRIX(i, 2) = BTEMP_MATRIX(i, 2) - ATEMP_VAL
        CTEMP_MATRIX(i, 1) = 0
    Next i

    If tolerance = 0 Then 'try to estimate epsilon
        ATEMP_VECTOR = MATRIX_TRIDIAGONAL_LOAD_FUNC(BTEMP_MATRIX)
        ReDim BTEMP_VECTOR(1 To UBound(ATEMP_VECTOR, 1), 1 To 1)
        TEMP_DET = MATRIX_TRIDIAGONAL_GJ_LINEAR_SYSTEM_FUNC(ATEMP_VECTOR, BTEMP_VECTOR, 10 ^ -13, 1)
        'If IsNumeric(TEMP_DET) = False Then: TEMP_DET = 0
        If TEMP_DET <= 10 ^ -13 Then epsilon = 10 ^ -13 Else _
            epsilon = 100 * Abs(TEMP_DET)
        If epsilon > 10 ^ -6 Then epsilon = 10 ^ -6
    Else
        epsilon = tolerance
    End If
    
    DETERM_FLAG = True
    NROWS = UBound(BTEMP_MATRIX, 1)
    NCOLUMNS = UBound(CTEMP_MATRIX, 2)

    BTEMP_MATRIX(1, 1) = 0
    TEMP_DET = 1

    For i = 1 To NROWS - 1
        TEMP_CHECK = 0
        For j = 1 To 3
            If Abs(BTEMP_MATRIX(i, j)) > epsilon Then TEMP_CHECK = 1: Exit For
        Next j
        If TEMP_CHECK = 0 Then
            TEMP_DET = 0: GoTo 1983  'singular matrix
        End If
    
        BTEMP_MATRIX(i, 1) = BTEMP_MATRIX(i, 2)
        BTEMP_MATRIX(i, 2) = BTEMP_MATRIX(i, 3)
        BTEMP_MATRIX(i, 3) = 0
    
        If Abs(BTEMP_MATRIX(i + 1, 1)) > epsilon Then
            If Abs(BTEMP_MATRIX(i, 1)) < Abs(BTEMP_MATRIX(i + 1, 1)) Then
                If DETERM_FLAG Then TEMP_DET = -TEMP_DET
                BTEMP_MATRIX = MATRIX_SWAP_ROW_FUNC(BTEMP_MATRIX, i + 1, i)
                CTEMP_MATRIX = MATRIX_SWAP_ROW_FUNC(CTEMP_MATRIX, i + 1, i)
            End If
        
            If DETERM_FLAG Then TEMP_DET = TEMP_DET * BTEMP_MATRIX(i, 1)
            BTEMP_VAL = -BTEMP_MATRIX(i + 1, 1) / BTEMP_MATRIX(i, 1)
            BTEMP_MATRIX = MATRIX_LINEAR_ROWS_COMBINATION_FUNC(BTEMP_MATRIX, i + 1, i, BTEMP_VAL)
            CTEMP_MATRIX = MATRIX_LINEAR_ROWS_COMBINATION_FUNC(CTEMP_MATRIX, i + 1, i, BTEMP_VAL)
        End If
    
        BTEMP_MATRIX(i + 1, 1) = 0
    Next i

    'determinant computation

    BTEMP_MATRIX(NROWS, 1) = BTEMP_MATRIX(NROWS, 2)
    BTEMP_MATRIX(NROWS, 2) = BTEMP_MATRIX(NROWS, 3)
    BTEMP_MATRIX(NROWS, 3) = 0

    If Abs(BTEMP_MATRIX(NROWS, 1)) <= epsilon Then
        BTEMP_MATRIX(NROWS, 1) = 0
        TEMP_DET = 0 '"singular"
    End If

    If DETERM_FLAG Then
        TEMP_DET = TEMP_DET * BTEMP_MATRIX(NROWS, 1)
        If Abs(TEMP_DET) <= epsilon Then
            TEMP_DET = 0
            GoTo 1983 'singular matrix
        End If
    End If

    For i = 1 To NROWS '1984 last row
        BTEMP_MATRIX(i, 2) = BTEMP_MATRIX(i, 2) / BTEMP_MATRIX(i, 1)
        BTEMP_MATRIX(i, 3) = BTEMP_MATRIX(i, 3) / BTEMP_MATRIX(i, 1)
    
        For j = 1 To NCOLUMNS
            CTEMP_MATRIX(i, j) = CTEMP_MATRIX(i, j) / BTEMP_MATRIX(i, 1)
        Next j
    Next i

    For i = NROWS - 1 To 1 Step -1 'backsubstitution
        For j = 1 To NCOLUMNS
            CTEMP_MATRIX(i, j) = CTEMP_MATRIX(i, j) - _
                BTEMP_MATRIX(i, 2) * CTEMP_MATRIX(i + 1, j)
            If i < NROWS - 1 Then CTEMP_MATRIX(i, j) = _
                CTEMP_MATRIX(i, j) - BTEMP_MATRIX(i, 3) * CTEMP_MATRIX(i + 2, j)
        Next j
    Next i

1983:
    
    If TEMP_DET > epsilon Then  'matrix not singular
        MATRIX_EIGENVECTOR_TRIDIAGONAL_FUNC = CTEMP_MATRIX
        Exit Function
    End If
    
    'inspection of reduced matrix
    l = 0
    For i = 1 To NSIZE
        '      COUNTER=0 => non singular matrix
        '      COUNTER=1 => eigenvalue simple
        '      COUNTER>1 => eigenvalue multiple
        If Abs(BTEMP_MATRIX(i, 1)) < epsilon Then _
            l = l + 1
    Next i
    
    For j = 1 To l
        jj = NSIZE - j + 1
        DTEMP_MATRIX(jj, k) = 1
        For i = jj - 1 To 1 Step -1
            If i = NSIZE - 1 Then
                DTEMP_MATRIX(i, k) = -BTEMP_MATRIX(i, 2) * _
                    DTEMP_MATRIX(i + 1, k) / BTEMP_MATRIX(i, 1)
            Else
                DTEMP_MATRIX(i, k) = (-BTEMP_MATRIX(i, 2) * _
                    DTEMP_MATRIX(i + 1, k) - BTEMP_MATRIX(i, 3) * _
                        DTEMP_MATRIX(i + 2, k)) / BTEMP_MATRIX(i, 1)
            End If
        Next
        k = k + 1
    Next j
Loop

NROWS = UBound(DTEMP_MATRIX, 1)
NCOLUMNS = UBound(DTEMP_MATRIX, 2)

For j = 1 To NCOLUMNS 'normalize the sign of each eigenvector making
'positive the first non zero element  |aij| > epsilon

    TEMP_SGN = 0
    For i = 1 To NROWS
        If Abs(DTEMP_MATRIX(i, j)) > 1000 * epsilon Then
            If TEMP_SGN = 0 Then TEMP_SGN = Sgn(DTEMP_MATRIX(i, j))
            If TEMP_SGN < 0 Then
                DTEMP_MATRIX(i, j) = -DTEMP_MATRIX(i, j)
            Else
                Exit For 'exit inner for
            End If
        End If
    Next i
Next j

MATRIX_EIGENVECTOR_TRIDIAGONAL_FUNC = MATRIX_NORMALIZED_VECTOR_FUNC(DTEMP_MATRIX, 2, epsilon)

Exit Function
ERROR_LABEL:
MATRIX_EIGENVECTOR_TRIDIAGONAL_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : MATRIX_DOMINANT_EIGEN_FUNC

'DESCRIPTION   : Returns the dominant real eigenvalue or Eigenvectors of a
'matrix. A dominant eigenvalue, if it exists, is the one with the maximum
'absolute value. 'Returns the dominant eigenvector of a matrix, related to
'the dominant real eigenvalue. nLOOPS (optional) sets the maximum number of
'iterations allowed (default 1000). This function uses the power iterative
'method. This algorithm is started with a random generic vector. Often it
'converges, but sometimes not.

'LIBRARY       : MATRIX
'GROUP         : EIGEN
'ID            : 009
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/28/2009
'************************************************************************************
'************************************************************************************

Function MATRIX_DOMINANT_EIGEN_FUNC(ByRef DATA_RNG As Variant, _
Optional ByVal NORM_FLAG As Boolean = False, _
Optional ByVal nLOOPS As Long = 1000, _
Optional ByVal OUTPUT As Integer = 0, _
Optional ByVal epsilon As Double = 10 ^ -14)

Dim i As Long
Dim j As Long
Dim l As Long

Dim NSIZE As Long
Dim NROWS As Long
Dim NCOLUMNS As Long

Dim ATEMP_VAL As Double
Dim BTEMP_VAL As Double
Dim CTEMP_VAL As Double
Dim DTEMP_VAL As Double

Dim TEMP_ERR As Double
Dim TEMP_DELTA As Double
Dim TEMP_SIGN As Double
Dim TEMP_SUM As Double

Dim ATEMP_VECTOR As Variant
Dim BTEMP_VECTOR As Variant
Dim EIGEN_VECTOR As Variant

Dim TEMP_MATRIX As Variant
Dim DATA_MATRIX As Variant

On Error GoTo ERROR_LABEL

DATA_MATRIX = DATA_RNG
If UBound(DATA_MATRIX, 1) <> UBound(DATA_MATRIX, 2) Then: GoTo ERROR_LABEL

NSIZE = UBound(DATA_MATRIX, 1)
ReDim ATEMP_VECTOR(1 To NSIZE, 1 To 1)
ReDim BTEMP_VECTOR(1 To NSIZE, 1 To 1)

ReDim EIGEN_VECTOR(1 To NSIZE, 1 To 1)
If NSIZE = 1 Then
    ATEMP_VAL = DATA_MATRIX(1, 1)
    EIGEN_VECTOR(1, 1) = 1
    GoTo 1983
End If
'initialize starting vector
Randomize
For i = 1 To NSIZE
    ATEMP_VECTOR(i, 1) = Rnd
Next i

DTEMP_VAL = 0
l = 0
Do
    TEMP_MATRIX = MMULT_FUNC(DATA_MATRIX, ATEMP_VECTOR, 70)
    CTEMP_VAL = 0
    BTEMP_VAL = 0
    For i = 1 To NSIZE
        CTEMP_VAL = CTEMP_VAL + TEMP_MATRIX(i, 1) * ATEMP_VECTOR(i, 1)
        BTEMP_VAL = BTEMP_VAL + ATEMP_VECTOR(i, 1) * ATEMP_VECTOR(i, 1)
    Next i
    
    TEMP_DELTA = CTEMP_VAL / BTEMP_VAL 'Rayleigh coefficient
    
    CTEMP_VAL = 0
    TEMP_SUM = 0
    
    For i = 1 To NSIZE
        ATEMP_VECTOR(i, 1) = TEMP_MATRIX(i, 1) / Sqr(BTEMP_VAL)
        'load and rescaling next vector
        
        TEMP_SUM = TEMP_SUM + _
            Abs(Abs(ATEMP_VECTOR(i, 1)) - Abs(BTEMP_VECTOR(i, 1)))
            'vector error evaluation
    Next i
    
    If BTEMP_VAL > 1 Then TEMP_SUM = TEMP_SUM / BTEMP_VAL
    
    TEMP_ERR = Abs(TEMP_DELTA - DTEMP_VAL)
    If Abs(TEMP_DELTA) > 1 Then TEMP_ERR = TEMP_ERR / Abs(TEMP_DELTA)
    TEMP_ERR = TEMP_ERR + TEMP_SUM
    DTEMP_VAL = TEMP_DELTA
    For i = 1 To NSIZE
        BTEMP_VECTOR(i, 1) = ATEMP_VECTOR(i, 1)    'save current vector
    Next i
    l = l + 1
Loop While l < nLOOPS And (TEMP_ERR > 4 * 10 ^ -15)

ATEMP_VAL = TEMP_DELTA
'normalize to max value

TEMP_DELTA = 0
For i = 1 To NSIZE
    If Abs(ATEMP_VECTOR(i, 1)) > Abs(TEMP_DELTA) Then _
        TEMP_DELTA = ATEMP_VECTOR(i, 1)
Next i

For i = 1 To NSIZE
    EIGEN_VECTOR(i, 1) = ATEMP_VECTOR(i, 1) / TEMP_DELTA
Next i


NROWS = UBound(ATEMP_VECTOR, 1)
NCOLUMNS = UBound(ATEMP_VECTOR, 2)

For j = 1 To NCOLUMNS 'normalize the sign of each eigenvector
'making positive the first non zero element  |aij| > tol
    TEMP_SIGN = 0
    For i = 1 To NROWS
        If Abs(ATEMP_VECTOR(i, j)) > 1000 * (4 * 10 ^ -15) Then
            If TEMP_SIGN = 0 Then TEMP_SIGN = Sgn(ATEMP_VECTOR(i, j))
            If TEMP_SIGN < 0 Then
                ATEMP_VECTOR(i, j) = -ATEMP_VECTOR(i, j)
            Else
                Exit For 'exit inner for
            End If
        End If
    Next i
Next j

1983:
If l >= nLOOPS And TEMP_ERR > 10 ^ 6 * epsilon Then: GoTo ERROR_LABEL
'If l >= nLOOPS Then: GoTo ERROR_LABEL
'convergence fails

Select Case OUTPUT
    Case 0 ' Dominant EigenValue
        MATRIX_DOMINANT_EIGEN_FUNC = ATEMP_VAL
    Case Else 'Dominant EigenVector
        If NORM_FLAG Then EIGEN_VECTOR = VECTOR_ABSOLUTE_NORM_FUNC(EIGEN_VECTOR, 0)
        MATRIX_DOMINANT_EIGEN_FUNC = EIGEN_VECTOR
End Select

Exit Function
ERROR_LABEL:
MATRIX_DOMINANT_EIGEN_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : MATRIX_POWER_EIGEN_SQUARE_FUNC

'DESCRIPTION   : This function returns all real eigenvalues or EigenVectors of a
'square matrix. Optional parameters are: nLOOPS sets the maximum number of
'iterations allowed (default 1000). NORM_FLAG if TRUE, the function
'returns a normalized vector |v|=1  (default FALSE)
'Remark: This function uses the power iterative method. This
'algorithm works also for asymmetric matrices with low-moderate
'dimension. This algorithm is started with a random generic vector.
'Often it converges, but sometimes not. So if one of these functions
'returns the error “limit iterations exceeded”, do not worry. Simply, re-try it.

'Uses the power algorithm and the matrix reduction method

'LIBRARY       : MATRIX
'GROUP         : EIGEN
'ID            : 010
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/28/2009
'************************************************************************************
'************************************************************************************


Function MATRIX_POWER_EIGEN_SQUARE_FUNC(ByRef DATA_RNG As Variant, _
Optional ByVal NORM_FLAG As Boolean = False, _
Optional ByVal OUTPUT As Integer = 0, _
Optional ByVal nLOOPS As Long = 1000)

Dim i As Long
Dim j As Long
Dim k As Long
Dim l As Long

Dim ii As Long
Dim jj As Long
Dim kk As Long
Dim ll As Long

Dim NROWS As Long
Dim NCOLUMNS As Long

Dim ANSIZE As Long
Dim BNSIZE As Long

Dim ATEMP_VAL As Double
Dim BTEMP_VAL As Double
Dim CTEMP_VAL As Double
Dim DTEMP_VAL As Double
Dim ETEMP_VAL As Double

Dim TEMP_ERR As Double
Dim TEMP_SIGN As Double
Dim TEMP_EIGEN As Double

Dim ATEMP_SUM As Double
Dim BTEMP_SUM As Double
Dim CTEMP_SUM As Double

Dim ATEMP_MATRIX As Variant
Dim BTEMP_MATRIX As Variant
Dim CTEMP_MATRIX As Variant
Dim DTEMP_MATRIX As Variant
Dim ETEMP_MATRIX As Variant

Dim ATEMP_VECTOR As Variant
Dim BTEMP_VECTOR As Variant
Dim CTEMP_VECTOR As Variant
Dim DTEMP_VECTOR As Variant
Dim ETEMP_VECTOR As Variant

Dim DATA_MATRIX As Variant
Dim TEMP_MATRIX As Variant

On Error GoTo ERROR_LABEL

DATA_MATRIX = DATA_RNG

If UBound(DATA_MATRIX, 1) <> UBound(DATA_MATRIX, 2) Then: GoTo ERROR_LABEL

ANSIZE = UBound(DATA_MATRIX, 1)

ReDim ATEMP_MATRIX(1 To ANSIZE, 1 To ANSIZE)
ReDim BTEMP_MATRIX(1 To ANSIZE, 1 To ANSIZE)
ReDim CTEMP_MATRIX(1 To ANSIZE, 1 To ANSIZE)
ReDim DTEMP_MATRIX(1 To ANSIZE, 1 To ANSIZE)

ReDim ATEMP_VECTOR(1 To ANSIZE, 1 To 1)
ReDim BTEMP_VECTOR(1 To ANSIZE, 1 To 1)

For i = 1 To ANSIZE
    BTEMP_VECTOR(i, 1) = True
Next i
k = 1

'------------------------------------------------------------------------------------------
Do
'------------------------------------------------------------------------------------------
    kk = ANSIZE - k + 1 'initialize and extract the reduced matrix
    ReDim TEMP_MATRIX(1 To kk, 1 To kk)
    ReDim CTEMP_VECTOR(1 To kk, 1 To 1)
    ii = 0
    For i = 1 To ANSIZE
        If BTEMP_VECTOR(i, 1) Then
            ii = ii + 1
            jj = 0
            For j = 1 To ANSIZE
                If BTEMP_VECTOR(j, 1) Then
                    jj = jj + 1
                    TEMP_MATRIX(ii, jj) = DATA_MATRIX(i, j)
                End If
            Next j
        End If
    Next i
    
'-------------search for the dominant eigenvalues and relative eigenvector---------------

    If UBound(TEMP_MATRIX, 1) <> UBound(TEMP_MATRIX, 2) Then: GoTo 1983

    BNSIZE = UBound(TEMP_MATRIX, 1)
    ReDim CTEMP_VECTOR(1 To BNSIZE, 1 To 1)
    ReDim DTEMP_VECTOR(1 To BNSIZE, 1 To 1)
    ReDim ETEMP_VECTOR(1 To BNSIZE, 1 To 1)
    If BNSIZE = 1 Then
        TEMP_EIGEN = TEMP_MATRIX(1, 1)
        CTEMP_VECTOR(1, 1) = 1
        GoTo 1983
    End If
    
    Randomize
    For i = 1 To BNSIZE 'initialize starting vector
        DTEMP_VECTOR(i, 1) = Rnd
    Next i
    
    ATEMP_VAL = 0
    ll = 0
'------------------------------------------------------------------------------------------
    Do
'------------------------------------------------------------------------------------------
        'multiplication matrix
        ETEMP_MATRIX = MMULT_FUNC(TEMP_MATRIX, DTEMP_VECTOR, 70)
        'Rayleigh coefficient computing
        ATEMP_SUM = 0
        BTEMP_SUM = 0
        For i = 1 To BNSIZE
            ATEMP_SUM = ATEMP_SUM + ETEMP_MATRIX(i, 1) * DTEMP_VECTOR(i, 1)
            BTEMP_SUM = BTEMP_SUM + DTEMP_VECTOR(i, 1) * DTEMP_VECTOR(i, 1)
        Next i
        BTEMP_VAL = ATEMP_SUM / BTEMP_SUM 'Rayleigh coefficient
        ATEMP_SUM = 0
        CTEMP_SUM = 0
        For i = 1 To BNSIZE
            DTEMP_VECTOR(i, 1) = ETEMP_MATRIX(i, 1) / (BTEMP_SUM) ^ 0.5
            'load and rescaling next vector
            CTEMP_SUM = CTEMP_SUM + Abs(Abs(DTEMP_VECTOR(i, 1)) - _
            Abs(ETEMP_VECTOR(i, 1))) 'vector error evaluation
        Next i
        If BTEMP_SUM > 1 Then CTEMP_SUM = CTEMP_SUM / BTEMP_SUM
        'eigenvalue error
        TEMP_ERR = Abs(BTEMP_VAL - ATEMP_VAL)
        If Abs(BTEMP_VAL) > 1 Then TEMP_ERR = TEMP_ERR / Abs(BTEMP_VAL)
        TEMP_ERR = TEMP_ERR + CTEMP_SUM
        ATEMP_VAL = BTEMP_VAL
        For i = 1 To BNSIZE
            ETEMP_VECTOR(i, 1) = DTEMP_VECTOR(i, 1)    'save current vector
        Next i
        ll = ll + 1
'------------------------------------------------------------------------------------------
    Loop While ll < nLOOPS And (TEMP_ERR > (4 * 10 ^ -15))   '
'------------------------------------------------------------------------------------------
    TEMP_EIGEN = BTEMP_VAL
    'normalize to max value
    BTEMP_VAL = 0

    For i = 1 To BNSIZE
        If Abs(DTEMP_VECTOR(i, 1)) > Abs(BTEMP_VAL) Then BTEMP_VAL = DTEMP_VECTOR(i, 1)
    Next i
    For i = 1 To BNSIZE
        CTEMP_VECTOR(i, 1) = DTEMP_VECTOR(i, 1) / BTEMP_VAL
    Next i
    
'------------normalize the sign of each eigenvector making positive the first-----------
    NROWS = UBound(DTEMP_VECTOR, 1)
    NCOLUMNS = UBound(DTEMP_VECTOR, 2)
    For j = 1 To NCOLUMNS
        TEMP_SIGN = 0
        For i = 1 To NROWS
            If Abs(DTEMP_VECTOR(i, j)) > 1000 * (4 * 10 ^ -15) Then
                If TEMP_SIGN = 0 Then TEMP_SIGN = Sgn(DTEMP_VECTOR(i, j))
                If TEMP_SIGN < 0 Then
                    DTEMP_VECTOR(i, j) = -DTEMP_VECTOR(i, j)
                Else
                    Exit For 'exit inner for
                End If
            End If
        Next i
    Next j
'---------------------------------------------------------------------------------------
1983:

    If ll >= nLOOPS And TEMP_ERR > 10 ^ 6 * (10 ^ -15) Then GoTo ERROR_LABEL
    'save results
    ATEMP_VECTOR(k, 1) = TEMP_EIGEN
    ii = 0
    For i = 1 To ANSIZE
        If BTEMP_VECTOR(i, 1) Then
            ii = ii + 1
            BTEMP_MATRIX(i, k) = CTEMP_VECTOR(ii, 1)
        End If
    Next i
    'search for first 1 - element
    For i = 1 To ANSIZE
        If BTEMP_MATRIX(i, k) = 1 Then
            l = i
            Exit For
        End If
    Next i
    'save l-row
    For j = 1 To ANSIZE
        DTEMP_MATRIX(k, j) = DATA_MATRIX(l, j)
    Next j
    'compute new reduce matrix
    For i = 1 To ANSIZE
        For j = 1 To ANSIZE
            CTEMP_MATRIX(i, j) = BTEMP_MATRIX(i, k) * DATA_MATRIX(l, j)
        Next j
    Next i
    For i = 1 To ANSIZE
        For j = 1 To ANSIZE
            DATA_MATRIX(i, j) = DATA_MATRIX(i, j) - CTEMP_MATRIX(i, j)
        Next j
    Next i
    k = k + 1
    BTEMP_VECTOR(l, 1) = False
'------------------------------------------------------------------------------------------
Loop Until k > ANSIZE
'------------------------------------------------------------------------------------------

For k = 1 To ANSIZE 'eigenvector transformation

    For i = 1 To ANSIZE
        ATEMP_MATRIX(i, k) = BTEMP_MATRIX(i, k)
    Next i
    
    For j = 1 To k - 1
        CTEMP_VAL = ATEMP_VECTOR(k, 1) - ATEMP_VECTOR(k - j, 1)
        DTEMP_VAL = 0
        For i = 1 To ANSIZE
            DTEMP_VAL = DTEMP_VAL + DTEMP_MATRIX(k - j, i) _
                    * ATEMP_MATRIX(i, k)
        Next i
        For i = 1 To ANSIZE
            ATEMP_MATRIX(i, k) = CTEMP_VAL * ATEMP_MATRIX(i, k) + _
                    DTEMP_VAL * BTEMP_MATRIX(i, k - j)
        Next i
    Next j
    'rescaling each vector to its max value
    CTEMP_VAL = 0
    ETEMP_VAL = 0
    For i = 1 To ANSIZE
        If Abs(ATEMP_MATRIX(i, k)) > CTEMP_VAL Then _
                    CTEMP_VAL = Abs(ATEMP_MATRIX(i, k))
        If CTEMP_VAL > (10 ^ -15) And ETEMP_VAL = 0 _
                Then ETEMP_VAL = Sgn(ATEMP_MATRIX(i, k)) 'first non zero element
    Next i
    
    If CTEMP_VAL > (10 ^ -15) Then
        For i = 1 To ANSIZE
            ATEMP_MATRIX(i, k) = ETEMP_VAL * ATEMP_MATRIX(i, k) / CTEMP_VAL
        Next i
    End If
Next k

'------------------------------------------------------------------------------------------
Select Case OUTPUT
'------------------------------------------------------------------------------------------
Case 0 ' EigenValues
'------------------------------------------------------------------------------------------
    MATRIX_POWER_EIGEN_SQUARE_FUNC = ATEMP_VECTOR
'------------------------------------------------------------------------------------------
Case Else 'EigenVectors
'------------------------------------------------------------------------------------------
    NROWS = UBound(ATEMP_MATRIX, 1)
    NCOLUMNS = UBound(ATEMP_MATRIX, 2)
    For j = 1 To NCOLUMNS 'normalize the sign of each eigenvector
    'making positive the first non zero element  |aij| > tol
        TEMP_SIGN = 0
        For i = 1 To NROWS
            If Abs(ATEMP_MATRIX(i, j)) > 1000 * (2 * 10 ^ -15) Then
                If TEMP_SIGN = 0 Then TEMP_SIGN = Sgn(ATEMP_MATRIX(i, j))
                If TEMP_SIGN < 0 Then
                    ATEMP_MATRIX(i, j) = -ATEMP_MATRIX(i, j)
                Else
                    Exit For 'exit inner for
                End If
            End If
        Next i
    Next j
    If NORM_FLAG Then ATEMP_MATRIX = MATRIX_NORMALIZED_VECTOR_FUNC(ATEMP_MATRIX, 2, 2 * 10 ^ -14)
    MATRIX_POWER_EIGEN_SQUARE_FUNC = ATEMP_MATRIX
'------------------------------------------------------------------------------------------
End Select
'------------------------------------------------------------------------------------------

Exit Function
ERROR_LABEL:
MATRIX_POWER_EIGEN_SQUARE_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : MATRIX_INVERSE_EIGENVECTOR_FUNC

'DESCRIPTION   : This function returns the eigenvector associated with the given
'EIGEN_VALUES of a matrix A using the inverse iteration algorithm
'If "EIGEN_VALUES" is a single value, the function returns a (n x 1)
'vector. Otherwise if Eigenvalues is a vector of all eigenvalues of
'matrix A, the function returns a matrix (n x n) of eigenvector.
'The eigenvector returned is normalized with norm=2 .

'The Inverse iteration method is adapt for eigenvalues affected by
'large errors.

'LIBRARY       : MATRIX
'GROUP         : EIGEN
'ID            : 011
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/28/2009
'************************************************************************************
'************************************************************************************

Function MATRIX_INVERSE_EIGENVECTOR_FUNC(ByRef DATA_RNG As Variant, _
ByRef EIGEN_RNG As Variant, _
Optional ByVal nLOOPS As Long = 20, _
Optional ByVal epsilon As Double = 10 ^ -10)
'return the real eigenvectors associated to the given real eigenvalues
'uses the inverse iteration algorithm

Dim i As Long
Dim j As Long
Dim k As Long
Dim l As Long

Dim NROWS As Long
Dim NCOLUMNS As Long

Dim TEMP_DOT As Double
Dim TEMP_RES As Double
Dim TEMP_SGN As Double
Dim TEMP_ERR As Double
Dim TEMP_PERT As Double

Dim EIGEN_VECTOR As Variant

Dim ATEMP_VECTOR As Variant
Dim BTEMP_VECTOR As Variant

Dim DATA_MATRIX As Variant

'Dim CTEMP_MATRIX As Variant
Dim BTEMP_MATRIX As Variant

On Error GoTo ERROR_LABEL

EIGEN_VECTOR = EIGEN_RNG
If UBound(EIGEN_VECTOR, 1) = 1 Then
    EIGEN_VECTOR = MATRIX_TRANSPOSE_FUNC(EIGEN_VECTOR)
End If
DATA_MATRIX = DATA_RNG
If UBound(DATA_MATRIX, 1) <> UBound(DATA_MATRIX, 2) Then GoTo ERROR_LABEL

NROWS = UBound(DATA_MATRIX, 1)
NCOLUMNS = UBound(EIGEN_VECTOR)

ReDim ATEMP_VECTOR(1 To NROWS, 1 To 1)
ReDim BTEMP_VECTOR(1 To NROWS, 1 To 1)

ReDim ATEMP_MATRIX(1 To NROWS, 1 To NCOLUMNS)
ReDim BTEMP_MATRIX(1 To NROWS, 1 To NROWS)

TEMP_PERT = Exp(0.5 * Log(epsilon)) 'compute the perturbation factor
'
ReDim ATEMP_MATRIX(1 To NROWS, 1 To NCOLUMNS)  'eigenvector matrix

For k = 1 To NCOLUMNS
    'load initial random values
    For i = 1 To NROWS
        ATEMP_VECTOR(i, 1) = 1 + 0.3 * Rnd
    Next i
    l = 0
    Do
        'build the iteration matrix
        For i = 1 To NROWS
            For j = 1 To NROWS
                BTEMP_MATRIX(i, j) = DATA_MATRIX(i, j)
            Next j
            BTEMP_MATRIX(i, i) = BTEMP_MATRIX(i, i) - _
                EIGEN_VECTOR(k, 1) * (1 + TEMP_PERT)
        Next i
        
        ATEMP_VECTOR = VECTOR_ABSOLUTE_NORM_FUNC(ATEMP_VECTOR, 0)
        
        For i = 1 To NROWS
            BTEMP_VECTOR(i, 1) = ATEMP_VECTOR(i, 1)
        Next i
        
'        CTEMP_MATRIX = BTEMP_MATRIX
 '       BTEMP_MATRIX = MATRIX_GS_REDUCTION_PIVOT_FUNC(BTEMP_MATRIX, BTEMP_VECTOR, epsilon, 1)
'        BTEMP_VECTOR = MATRIX_GS_REDUCTION_PIVOT_FUNC(CTEMP_MATRIX, BTEMP_VECTOR, epsilon, 0)
        BTEMP_VECTOR = MATRIX_GS_REDUCTION_PIVOT_FUNC(BTEMP_MATRIX, BTEMP_VECTOR, epsilon, 0)

        TEMP_DOT = 0
        For i = 1 To NROWS
            TEMP_DOT = TEMP_DOT + ATEMP_VECTOR(i, 1) * BTEMP_VECTOR(i, 1)
        Next i
        
        TEMP_RES = 0
        For i = 1 To NROWS
            TEMP_RES = TEMP_RES + (BTEMP_VECTOR(i, 1) - _
                TEMP_DOT * ATEMP_VECTOR(i, 1)) ^ 2
        Next i
        TEMP_RES = Sqr(TEMP_RES)
        
        TEMP_ERR = 0
        If TEMP_RES <> 0 Then TEMP_ERR = TEMP_RES / Abs(TEMP_DOT)
        
        For i = 1 To NROWS
            ATEMP_VECTOR(i, 1) = BTEMP_VECTOR(i, 1)
        Next i
        l = l + 1
    Loop Until l > nLOOPS Or TEMP_ERR <= epsilon
    
    ATEMP_VECTOR = VECTOR_ABSOLUTE_NORM_FUNC(ATEMP_VECTOR, 0)
    
    For i = 1 To NROWS
        ATEMP_MATRIX(i, k) = ATEMP_VECTOR(i, 1)
        If Abs(ATEMP_MATRIX(i, k)) < epsilon Then ATEMP_MATRIX(i, k) = 0
    Next i
Next k

NROWS = UBound(ATEMP_MATRIX, 1)
NCOLUMNS = UBound(ATEMP_MATRIX, 2)

For j = 1 To NCOLUMNS 'normalize the sign of each eigenvector making
'positive the first non zero element  |aij| > epsilon
TEMP_SGN = 0
    For i = 1 To NROWS
        If Abs(ATEMP_MATRIX(i, j)) > 1000 * epsilon Then
            If TEMP_SGN = 0 Then TEMP_SGN = Sgn(ATEMP_MATRIX(i, j))
            If TEMP_SGN < 0 Then
                ATEMP_MATRIX(i, j) = -ATEMP_MATRIX(i, j)
            Else
                Exit For 'exit inner for
            End If
        End If
    Next i
Next j

MATRIX_INVERSE_EIGENVECTOR_FUNC = ATEMP_MATRIX

Exit Function
ERROR_LABEL:
MATRIX_INVERSE_EIGENVECTOR_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : MATRIX_RANDOM_EIGENVALUES_FUNC
'DESCRIPTION   : Returns a matrix with a given set of eigenvalues
'LIBRARY       : MATRIX
'GROUP         : EIGEN
'ID            : 012
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/28/2009
'************************************************************************************
'************************************************************************************

Function MATRIX_RANDOM_EIGENVALUES_FUNC(ByRef EIGEN_RNG As Variant, _
Optional ByVal INT_FLAG As Boolean = True)

'INT_FLAG = True (default) for Integer matrix, False for decimal matrix

Dim i As Long
Dim j As Long
Dim k As Long

Dim NROWS As Long
Dim NCOLUMNS As Long

Dim DEV_VAL As Double
Dim MEAN_VAL As Double

Dim ATEMP_VAL As Double
Dim BTEMP_VAL As Double 'shaffer

Dim UPPER_BOUND As Double
Dim LOWER_BOUND As Double

Dim ATEMP_MATRIX As Variant
Dim BTEMP_MATRIX As Variant
Dim CTEMP_MATRIX As Variant
Dim DTEMP_MATRIX As Variant

Dim DATA_VECTOR As Variant

On Error GoTo ERROR_LABEL

DATA_VECTOR = EIGEN_RNG
If UBound(DATA_VECTOR, 1) = 1 Then
    DATA_VECTOR = MATRIX_TRANSPOSE_FUNC(DATA_VECTOR)
End If

NROWS = UBound(DATA_VECTOR, 1)
NCOLUMNS = NROWS

MEAN_VAL = 0
DEV_VAL = 2
BTEMP_VAL = NROWS

ReDim ATEMP_MATRIX(1 To NROWS, 1 To NCOLUMNS)
ReDim BTEMP_MATRIX(1 To NROWS, 1 To NCOLUMNS)
ReDim CTEMP_MATRIX(1 To NROWS, 1 To NCOLUMNS)

UPPER_BOUND = (MEAN_VAL + DEV_VAL)
LOWER_BOUND = (MEAN_VAL - DEV_VAL)

Randomize
For i = 1 To NROWS
    For j = 1 To NCOLUMNS
        BTEMP_MATRIX(i, j) = (UPPER_BOUND - LOWER_BOUND + 1) * _
            Rnd + LOWER_BOUND
        If INT_FLAG Then BTEMP_MATRIX(i, j) = Int(BTEMP_MATRIX(i, j))
        If i > j Then BTEMP_MATRIX(i, j) = 0
    Next j
Next i

For i = 1 To NROWS 'set determinant 1
    BTEMP_MATRIX(i, i) = 1
Next

For i = 2 To NROWS 'shaffer
    ATEMP_VAL = Int((UPPER_BOUND - LOWER_BOUND + 1) * Rnd + LOWER_BOUND)
    BTEMP_MATRIX = MATRIX_LINEAR_ROWS_COMBINATION_FUNC(BTEMP_MATRIX, i, 1, ATEMP_VAL)
Next i

k = MINIMUM_FUNC(NROWS, NCOLUMNS) 'Matrice diagonale DATA_VECTOR
For i = 1 To k
    ATEMP_MATRIX(i, i) = DATA_VECTOR(i, 1)
Next i
CTEMP_MATRIX = MATRIX_LU_INVERSE_FUNC(BTEMP_MATRIX)
    
ReDim DTEMP_MATRIX(1 To NROWS, 1 To NROWS)
    
For i = 1 To NROWS
    For j = 1 To NROWS
        For k = 1 To NROWS
            DTEMP_MATRIX(i, j) = DTEMP_MATRIX(i, j) + _
            CTEMP_MATRIX(i, k) * ATEMP_MATRIX(k, j)
        Next k
    Next j
Next i
CTEMP_MATRIX = DTEMP_MATRIX
    
ReDim DTEMP_MATRIX(1 To NROWS, 1 To NROWS) 'Reset the Array
For i = 1 To NROWS
    For j = 1 To NROWS
        For k = 1 To NROWS
            DTEMP_MATRIX(i, j) = DTEMP_MATRIX(i, j) + _
            CTEMP_MATRIX(i, k) * BTEMP_MATRIX(k, j)
        Next k
    Next j
Next i
    
MATRIX_RANDOM_EIGENVALUES_FUNC = DTEMP_MATRIX
    
Exit Function
ERROR_LABEL:
MATRIX_RANDOM_EIGENVALUES_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : MATRIX_RANDOM_EIGENVALUES_SQUARE_FUNC
'DESCRIPTION   : Returns a square matrix having the given eigenvaluess
'LIBRARY       : MATRIX
'GROUP         : EIGEN
'ID            : 013
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/28/2009
'************************************************************************************
'************************************************************************************

Function MATRIX_RANDOM_EIGENVALUES_SQUARE_FUNC(ByRef EIGEN_RNG As Variant, _
Optional ByVal MEAN_VAL As Double = 0, _
Optional ByVal DEV_VAL As Double = 1)

Dim i As Long
Dim j As Long
Dim k As Long

Dim NSIZE As Long

Dim UPPER_BOUND As Double
Dim LOWER_BOUND As Double

Dim ATEMP_MATRIX As Variant
Dim BTEMP_MATRIX As Variant
Dim CTEMP_MATRIX As Variant
Dim DTEMP_MATRIX As Variant

Dim TEMP_VECTOR As Variant
Dim DATA_VECTOR As Variant

On Error GoTo ERROR_LABEL

DATA_VECTOR = EIGEN_RNG
If UBound(DATA_VECTOR, 1) = 1 Then
    DATA_VECTOR = MATRIX_TRANSPOSE_FUNC(DATA_VECTOR)
End If
NSIZE = UBound(DATA_VECTOR, 1)

ReDim TEMP_VECTOR(1 To NSIZE, 1 To 1)

UPPER_BOUND = (MEAN_VAL + DEV_VAL)
LOWER_BOUND = (MEAN_VAL - DEV_VAL)

Randomize
For i = 1 To NSIZE
    TEMP_VECTOR(i, 1) = (UPPER_BOUND - _
    LOWER_BOUND + 1) * Rnd + LOWER_BOUND
Next i

ReDim ATEMP_MATRIX(1 To NSIZE, 1 To NSIZE)
For i = 1 To NSIZE
    ATEMP_MATRIX(i, i) = DATA_VECTOR(i, 1)
Next i

BTEMP_MATRIX = MATRIX_HOUSEHOLDER_FUNC(TEMP_VECTOR)  'build Houseolder matrix H

ReDim DTEMP_MATRIX(1 To NSIZE, 1 To NSIZE)

For i = 1 To NSIZE
    For j = 1 To NSIZE
        For k = 1 To NSIZE
            DTEMP_MATRIX(i, j) = DTEMP_MATRIX(i, j) + ATEMP_MATRIX(i, k) * _
                BTEMP_MATRIX(k, j)
        Next k
    Next j
Next i

CTEMP_MATRIX = DTEMP_MATRIX
ReDim DTEMP_MATRIX(1 To NSIZE, 1 To NSIZE)

For i = 1 To NSIZE
    For j = 1 To NSIZE
        For k = 1 To NSIZE
            DTEMP_MATRIX(i, j) = DTEMP_MATRIX(i, j) + BTEMP_MATRIX(i, k) * _
                CTEMP_MATRIX(k, j)
        Next k
    Next j
Next i

MATRIX_RANDOM_EIGENVALUES_SQUARE_FUNC = DTEMP_MATRIX

Exit Function
ERROR_LABEL:
MATRIX_RANDOM_EIGENVALUES_SQUARE_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : MATRIX_CHECK_DEFINITENESS_FUNC
'DESCRIPTION   : Checks definiteness of symmetric matrices using their eigenvalues
'LIBRARY       : MATRIX
'GROUP         : EIGEN
'ID            : 014
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/28/2009
'************************************************************************************
'************************************************************************************

Function MATRIX_CHECK_DEFINITENESS_FUNC(ByRef EIGEN_RNG As Variant)
    
Dim NSIZE As Long
    
Dim TEMP_SUM As Double
Dim TEMP_MIN As Double
Dim TEMP_MAX As Double
Dim TEMP_VAL As Double
    
Dim DATA_VECTOR As Variant
    
On Error GoTo ERROR_LABEL

DATA_VECTOR = EIGEN_RNG
If UBound(DATA_VECTOR, 1) = 1 Then
    DATA_VECTOR = MATRIX_TRANSPOSE_FUNC(DATA_VECTOR)
End If
    
NSIZE = UBound(DATA_VECTOR, 1)
DATA_VECTOR = VECTOR_SIGN_VALUES_FUNC(DATA_VECTOR, 1)
    
TEMP_SUM = MATRIX_ELEMENTS_CUMULATIVE_SUM_FUNC(DATA_VECTOR)
TEMP_MIN = MATRIX_ELEMENTS_MIN_FUNC(DATA_VECTOR, 0)
TEMP_MAX = MATRIX_ELEMENTS_MAX_FUNC(DATA_VECTOR, 0)
    
If TEMP_SUM = NSIZE Then
    TEMP_VAL = 1 '(+ve def)
ElseIf TEMP_SUM = -NSIZE Then
    TEMP_VAL = -1 '(-ve def)
ElseIf TEMP_SUM >= 0 And TEMP_MIN >= 0 Then
    TEMP_VAL = 0.5 '(+ve semi-def)
ElseIf TEMP_SUM <= 0 And TEMP_MAX <= 0 Then
    TEMP_VAL = -0.5 '(-ve semi-def)
Else
    TEMP_VAL = 0
End If

MATRIX_CHECK_DEFINITENESS_FUNC = TEMP_VAL

Exit Function
ERROR_LABEL:
MATRIX_CHECK_DEFINITENESS_FUNC = Err.number
End Function

