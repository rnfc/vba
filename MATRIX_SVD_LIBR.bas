Attribute VB_Name = "MATRIX_SVD_LIBR"

'------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------
Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.
'------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------

'------------------------------------------------------------------------------------
'Singular Value Decomposition of a (n x m) matrix A provides three matrices, U, D, V
'performing the following decomposition1:

'A =U ×D×VT; where: p = Min(n, m)
'U is an orthogonal matrix (n x p)
'D is a square diagonal matrix (p x p)
'V is an orthogonal matrix (m x p)

'From the D matrix of singular values we get the max and min values to compute the
'condition number, used to measure the ill-conditioning of a matrix.

'The SVD decomposition of a square matrix always returns square matrices of the same
'size, but for a rectangular matrix we should pay a bit more attention to the correct
'dimensions.

'Nomenclature: The matrices returned by the SVD are sometimes called
'U ( hanger ), D (stretcher), V (aligner).
'So the decomposition for a matrix A can be written as1
'(any matrix) = ( hanger ) x (stretcher) x (aligner)
'------------------------------------------------------------------------------------


'************************************************************************************
'************************************************************************************
'FUNCTION      : MATRIX_SVD_FACT_FUNC (Self-autonomous function)

'DESCRIPTION   : 'Singular Value Decomposition.
'For an m-by-n matrix A with m >= n, the singular value decomposition is
'an m-by-n orthogonal matrix U, an n-by-n diagonal matrix S, and
'an n-by-n orthogonal matrix V so that A = U*S*V'.
'The singular values, sigma[k] = S[k][k], are ordered so that
'sigma[0] >= sigma[1] >= ... >= sigma[n-1].
'The singular value decompostion always exists, so the constructor will
'never fail.  The matrix condition number and the effective numerical
'rank can be computed from this decomposition.

'LIBRARY       : MATRIX
'GROUP         : SVD
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/19/2009
'************************************************************************************
'************************************************************************************

Function MATRIX_SVD_FACT_FUNC(ByRef DATA_RNG As Variant, _
Optional ByVal OUTPUT As Integer = 0)
  
  Dim h As Single
  Dim i As Single
  Dim j As Single
  Dim k As Single
  Dim l As Single
  Dim m As Single

  Dim NSIZE As Single
  Dim NROWS As Single
  Dim NCOLUMNS As Single
  
  Dim P_VAL As Single
  Dim Q_VAL As Single
  Dim U_VAL As Single
  Dim V_VAL As Single

  Dim ATEMP_VAL As Double
  Dim BTEMP_VAL As Double
  Dim CTEMP_VAL As Double
  Dim DTEMP_VAL As Double
  Dim ETEMP_VAL As Double
  Dim FTEMP_VAL As Double
  Dim GTEMP_VAL As Double
  Dim HTEMP_VAL As Double
  Dim ITEMP_VAL As Double
  Dim JTEMP_VAL As Double
  Dim KTEMP_VAL As Double
  Dim LTEMP_VAL As Double
  Dim MTEMP_VAL As Double
  Dim NTEMP_VAL As Double
  
  Dim SHIFT_VAL As Double
  Dim SCALE_VAL As Double
  
  Dim DATA_MATRIX As Variant
  Dim TEMP_MATRIX As Variant
  
  Dim ETEMP_ARR() As Double
  Dim STEMP_ARR As Variant
  Dim WTEMP_ARR As Variant
  
  Dim STEMP_MATRIX As Variant
  Dim UTEMP_MATRIX As Variant
  Dim VTEMP_MATRIX As Variant

  Dim nLOOPS As Single
  Dim VERSION As Single
  Dim epsilon As Double

  Dim TRANSPOSE_FLAG As Boolean

  On Error GoTo ERROR_LABEL

  DATA_MATRIX = DATA_RNG
  If UBound(DATA_MATRIX, 1) < UBound(DATA_MATRIX, 2) Then
    TRANSPOSE_FLAG = True
    DATA_MATRIX = MATRIX_TRANSPOSE_FUNC(DATA_MATRIX)
  End If
  DATA_MATRIX = MATRIX_CHANGE_BASE_ZERO_FUNC(DATA_MATRIX)
  
  NROWS = UBound(DATA_MATRIX, 1) + 1
  NCOLUMNS = UBound(DATA_MATRIX, 2) + 1
  NSIZE = MINIMUM_FUNC(NROWS, NCOLUMNS)
  
  ReDim STEMP_ARR(0 To MINIMUM_FUNC(NROWS + 1, NCOLUMNS) - 1)
  ReDim UTEMP_MATRIX(0 To NROWS - 1, 0 To NSIZE - 1)
  ReDim VTEMP_MATRIX(0 To NCOLUMNS - 1, 0 To NCOLUMNS - 1)
  
  ReDim ETEMP_ARR(0 To NCOLUMNS - 1)
  ReDim WTEMP_ARR(0 To NROWS - 1)
  
  TEMP_MATRIX = DATA_MATRIX
  U_VAL = 1
  V_VAL = 1
  
  ' Reduce TEMP_MATRIX to bidiagonal form, storing the diagonal elements
  ' in STEMP_ARR and the super-diagonal elements in ETEMP_ARR.
  h = MINIMUM_FUNC(NROWS - 1, NCOLUMNS)
  l = MAXIMUM_FUNC(0, MINIMUM_FUNC(NCOLUMNS - 2, NROWS))
  
'------------------------------------------------------------------------------------
For k = 0 To MAXIMUM_FUNC(h, l) - 1
'----------------------------------------------------------------------------
  If k < h Then
    ' Compute the transformation for the k-th column and
    ' place the k-th diagonal in STEMP_ARR(k).
    ' Compute 2-norm of k-th column without under/overflow.
'----------------------------------------------------------------------------
    STEMP_ARR(k) = 0
    For i = k To NROWS - 1
      ATEMP_VAL = STEMP_ARR(k)
      BTEMP_VAL = TEMP_MATRIX(i, k)
      STEMP_ARR(k) = HYPOT_FUNC(ATEMP_VAL, BTEMP_VAL)
    Next i
    If STEMP_ARR(k) <> 0 Then
        If TEMP_MATRIX(k, k) < 0 Then
            STEMP_ARR(k) = -STEMP_ARR(k)
        End If
        For i = k To NROWS - 1
          TEMP_MATRIX(i, k) = TEMP_MATRIX(i, k) / STEMP_ARR(k)
        Next i
        TEMP_MATRIX(k, k) = TEMP_MATRIX(k, k) + 1
    End If
    STEMP_ARR(k) = -STEMP_ARR(k)
  End If

  For j = k + 1 To NCOLUMNS - 1
    If ((k < h) And (STEMP_ARR(k) <> 0)) Then
        ' Apply the transformation.
        CTEMP_VAL = 0
        For i = k To NROWS - 1
            CTEMP_VAL = CTEMP_VAL + TEMP_MATRIX(i, k) * TEMP_MATRIX(i, j)
        Next i
        CTEMP_VAL = -CTEMP_VAL / TEMP_MATRIX(k, k)
        For i = k To NROWS - 1
            TEMP_MATRIX(i, j) = TEMP_MATRIX(i, j) + CTEMP_VAL * TEMP_MATRIX(i, k)
        Next i
    End If
    ' Place the k-th row of TEMP_MATRIX into ETEMP_ARR for the
    ' subsequent calculation of the row transformation.
    ETEMP_ARR(j) = TEMP_MATRIX(k, j)
  Next j

  If (U_VAL And (k < h)) Then
  ' Place the transformation in UTEMP_MATRIX for subsequent back
  ' multiplication.
    For i = k To NROWS - 1
      UTEMP_MATRIX(i, k) = TEMP_MATRIX(i, k)
    Next i
  End If
'----------------------------------------------------------------------------
  If k < l Then
    ' Compute the k-th row transformation and place the
    ' k-th super-diagonal in ETEMP_ARR(k).
    ' Compute 2-norm without under/overflow.
'----------------------------------------------------------------------------
    ETEMP_ARR(k) = 0
    For i = k + 1 To NCOLUMNS - 1
        ETEMP_ARR(k) = HYPOT_FUNC(ETEMP_ARR(k), ETEMP_ARR(i))
    Next i
    If (ETEMP_ARR(k) <> 0) Then
        If (ETEMP_ARR(k + 1) < 0) Then
            ETEMP_ARR(k) = -ETEMP_ARR(k)
        End If
        For i = k + 1 To NCOLUMNS - 1
            ETEMP_ARR(i) = ETEMP_ARR(i) / ETEMP_ARR(k)
        Next i
        ETEMP_ARR(k + 1) = ETEMP_ARR(k + 1) + 1
    End If
    ETEMP_ARR(k) = -ETEMP_ARR(k)
    If ((k + 1 < NROWS) And (ETEMP_ARR(k) <> 0)) Then
        ' Apply the transformation.
        For i = k + 1 To NROWS - 1
            WTEMP_ARR(i) = 0
        Next i
        For j = k + 1 To NCOLUMNS - 1
            For i = k + 1 To NROWS - 1
                WTEMP_ARR(i) = WTEMP_ARR(i) + ETEMP_ARR(j) * TEMP_MATRIX(i, j)
            Next i
        Next j
        For j = k + 1 To NCOLUMNS - 1
            CTEMP_VAL = -ETEMP_ARR(j) / ETEMP_ARR(k + 1)
            For i = k + 1 To NROWS - 1
                TEMP_MATRIX(i, j) = TEMP_MATRIX(i, j) + CTEMP_VAL * WTEMP_ARR(i)
            Next i
        Next j
    End If
    If (V_VAL) Then
        ' Place the transformation in VTEMP_MATRIX for subsequent
        ' back multiplication.
        For i = k + 1 To NCOLUMNS - 1
            VTEMP_MATRIX(i, k) = ETEMP_ARR(i)
        Next i
    End If
'------------------------------------------------------------------------------------
  End If
'------------------------------------------------------------------------------------
Next k
'------------------------------------------------------------------------------------

' Set up the final bidiagonal matrix or order P_VAL.

    P_VAL = MINIMUM_FUNC(NCOLUMNS, NROWS + 1)
    If (h < NCOLUMNS) Then: STEMP_ARR(h) = TEMP_MATRIX(h, h)
    If (NROWS < P_VAL) Then: STEMP_ARR(P_VAL - 1) = 0
    If (l + 1 < P_VAL) Then: ETEMP_ARR(l) = TEMP_MATRIX(l, P_VAL - 1)
    ETEMP_ARR(P_VAL - 1) = 0

    ' If required, generate UTEMP_MATRIX.
    If (U_VAL > 0) Then
        For j = h To NSIZE - 1
            For i = 0 To NROWS - 1
                UTEMP_MATRIX(i, j) = 0
            Next i
            UTEMP_MATRIX(j, j) = 1
        Next j
        For k = h - 1 To 0 Step -1
            If (STEMP_ARR(k) <> 0) Then
                For j = k + 1 To NSIZE - 1
                    'double CTEMP_VAL = 0;
                    CTEMP_VAL = 0
                    For i = k To NROWS - 1
                        CTEMP_VAL = CTEMP_VAL + UTEMP_MATRIX(i, k) * _
                                    UTEMP_MATRIX(i, j)
                    Next i
                    CTEMP_VAL = -CTEMP_VAL / UTEMP_MATRIX(k, k)
                    For i = k To NROWS - 1
                        UTEMP_MATRIX(i, j) = UTEMP_MATRIX(i, j) + _
                            CTEMP_VAL * UTEMP_MATRIX(i, k)
                    Next i
                Next j
                For i = k To NROWS - 1
                    UTEMP_MATRIX(i, k) = -UTEMP_MATRIX(i, k)
                Next i
                UTEMP_MATRIX(k, k) = 1 + UTEMP_MATRIX(k, k)
                For i = 0 To k - 2
                    UTEMP_MATRIX(i, k) = 0
                Next i
            Else
                For i = 0 To NROWS - 1
                    UTEMP_MATRIX(i, k) = 0
                Next i
                UTEMP_MATRIX(k, k) = 1
            End If
        Next k
    End If
    ' If required, generate VTEMP_MATRIX.
    If (V_VAL > 0) Then
        For k = NCOLUMNS - 1 To 0 Step -1
            If ((k < l) And (ETEMP_ARR(k) <> 0)) Then
                For j = k + 1 To NSIZE - 1
                    CTEMP_VAL = 0
                    For i = k + 1 To NCOLUMNS - 1
                        CTEMP_VAL = CTEMP_VAL + VTEMP_MATRIX(i, k) * _
                                    VTEMP_MATRIX(i, j)
                    Next i
                    CTEMP_VAL = -CTEMP_VAL / VTEMP_MATRIX(k + 1, k)
                    For i = k + 1 To NCOLUMNS - 1
                        VTEMP_MATRIX(i, j) = VTEMP_MATRIX(i, j) + _
                                             CTEMP_VAL * VTEMP_MATRIX(i, k)
                    Next i
                Next j
            End If
            For i = 0 To NCOLUMNS - 1
                VTEMP_MATRIX(i, k) = 0
            Next i
            VTEMP_MATRIX(k, k) = 1
        Next k
    End If

' Main iteration loop for the singular values.
            
    Q_VAL = P_VAL - 1
    nLOOPS = 0
    epsilon = 2 ^ -52
'   while (P_VAL > 0) {
'------------------------------------------------------------------------------------
   ' Here is where TEMP_MATRIX test for too many iterations would go.
   ' This section of the program inspects for
   ' negligible elements in the STEMP_ARR and ETEMP_ARR arrays.  On
   ' completion the variables VERSION and k are set as follows.
    Do While P_VAL > 0
'------------------------------------------------------------------------------------
                
        For k = P_VAL - 2 To -1 Step -1
            If (k = -1) Then: Exit For
            If (Abs(ETEMP_ARR(k)) <= epsilon * (Abs(STEMP_ARR(k)) + _
                Abs(STEMP_ARR(k + 1)))) Then
                ETEMP_ARR(k) = 0
                Exit For
            End If
        Next k

        If (k = P_VAL - 2) Then
            VERSION = 4
        Else
            For m = P_VAL - 1 To k Step -1
                If (m = k) Then: Exit For
                CTEMP_VAL = 0
                If m <> P_VAL Then
                    CTEMP_VAL = Abs(ETEMP_ARR(m))
                Else
                    CTEMP_VAL = 0
                End If
                If m <> k + 1 Then
                    CTEMP_VAL = CTEMP_VAL + Abs(ETEMP_ARR(m - 1))
                Else
                    CTEMP_VAL = CTEMP_VAL + 0
                End If
                If (Abs(STEMP_ARR(m)) <= epsilon * CTEMP_VAL) Then
                    STEMP_ARR(m) = 0
                    Exit For
                End If
            Next m
            If (m = k) Then
                VERSION = 3
            ElseIf (m = P_VAL - 1) Then
                VERSION = 1
            Else
                VERSION = 2
                k = m
            End If
        End If
        k = k + 1

' Perform the task indicated by VERSION.
    
'----------------------------------------------------------------------------------
    If VERSION = 1 Then
    'if STEMP_ARR(P_VAL) and ETEMP_ARR(k-1) are negligible and k<p
'----------------------------------------------------------------------------------
 ' Deflate negligible STEMP_ARR(P_VAL).
        FTEMP_VAL = ETEMP_ARR(P_VAL - 2)
        ETEMP_ARR(P_VAL - 2) = 0
        For j = P_VAL - 2 To k Step -1
            CTEMP_VAL = 0
            ATEMP_VAL = STEMP_ARR(j)
            CTEMP_VAL = HYPOT_FUNC(ATEMP_VAL, FTEMP_VAL)
            DTEMP_VAL = STEMP_ARR(j) / CTEMP_VAL
            ETEMP_VAL = FTEMP_VAL / CTEMP_VAL
            STEMP_ARR(j) = CTEMP_VAL
            If (j <> k) Then
                FTEMP_VAL = -ETEMP_VAL * ETEMP_ARR(j - 1)
                ETEMP_ARR(j - 1) = DTEMP_VAL * ETEMP_ARR(j - 1)
            End If
            If (V_VAL) Then
                For i = 0 To NCOLUMNS - 1
                    CTEMP_VAL = DTEMP_VAL * VTEMP_MATRIX(i, j) + _
                                ETEMP_VAL * VTEMP_MATRIX(i, P_VAL - 1)
                    VTEMP_MATRIX(i, P_VAL - 1) = _
                                -ETEMP_VAL * VTEMP_MATRIX(i, j) + _
                                DTEMP_VAL * VTEMP_MATRIX(i, P_VAL - 1)
                    VTEMP_MATRIX(i, j) = CTEMP_VAL
                Next i
            End If
        Next j
'----------------------------------------------------------------------------------
    ElseIf VERSION = 2 Then
    ' Split at negligible STEMP_ARR(k).
    ' if STEMP_ARR(k) is negligible and k<p
'----------------------------------------------------------------------------------
        FTEMP_VAL = ETEMP_ARR(k - 1)
        ETEMP_ARR(k - 1) = 0
        For j = k To P_VAL - 1
            ATEMP_VAL = STEMP_ARR(j)
            CTEMP_VAL = HYPOT_FUNC(ATEMP_VAL, FTEMP_VAL)
            DTEMP_VAL = STEMP_ARR(j) / CTEMP_VAL
            ETEMP_VAL = FTEMP_VAL / CTEMP_VAL
            STEMP_ARR(j) = CTEMP_VAL
            FTEMP_VAL = -ETEMP_VAL * ETEMP_ARR(j)
            ETEMP_ARR(j) = DTEMP_VAL * ETEMP_ARR(j)
            If (U_VAL) Then
                For i = 0 To NROWS - 1
                    CTEMP_VAL = DTEMP_VAL * UTEMP_MATRIX(i, j) + _
                                ETEMP_VAL * UTEMP_MATRIX(i, k - 1)
                    UTEMP_MATRIX(i, k - 1) = -ETEMP_VAL * UTEMP_MATRIX(i, j) + _
                                DTEMP_VAL * UTEMP_MATRIX(i, k - 1)
                    UTEMP_MATRIX(i, j) = CTEMP_VAL
                Next i
            End If
        Next j
'----------------------------------------------------------------------------------
    ElseIf VERSION = 3 Then
    'Calculate the SHIFT_VAL.
    'if ETEMP_ARR(k-1) is negligible, k<p, and
    'STEMP_ARR(k), ..., STEMP_ARR(P_VAL) are not
    'negligible (qr step).
'----------------------------------------------------------------------------------
        SCALE_VAL = _
        MAXIMUM_FUNC(MAXIMUM_FUNC(MAXIMUM_FUNC(MAXIMUM_FUNC(Abs(STEMP_ARR(P_VAL - 1)), _
        Abs(STEMP_ARR(P_VAL - 2))), Abs(ETEMP_ARR(P_VAL - 2))), Abs(STEMP_ARR(k))), _
        Abs(ETEMP_ARR(k)))
                      
        GTEMP_VAL = STEMP_ARR(P_VAL - 1) / SCALE_VAL
        HTEMP_VAL = STEMP_ARR(P_VAL - 2) / SCALE_VAL
        ITEMP_VAL = ETEMP_ARR(P_VAL - 2) / SCALE_VAL
        JTEMP_VAL = STEMP_ARR(k) / SCALE_VAL
        KTEMP_VAL = ETEMP_ARR(k) / SCALE_VAL
        LTEMP_VAL = ((HTEMP_VAL + GTEMP_VAL) * (HTEMP_VAL - GTEMP_VAL) + _
                    ITEMP_VAL * ITEMP_VAL) / 2
        MTEMP_VAL = (GTEMP_VAL * ITEMP_VAL) * (GTEMP_VAL * ITEMP_VAL)
        SHIFT_VAL = 0

        If ((LTEMP_VAL <> 0) Or _
            (MTEMP_VAL <> 0)) Then
            SHIFT_VAL = (LTEMP_VAL * LTEMP_VAL + MTEMP_VAL) ^ 0.5
            If (LTEMP_VAL < 0) Then: SHIFT_VAL = -SHIFT_VAL
            SHIFT_VAL = MTEMP_VAL / (LTEMP_VAL + SHIFT_VAL)
        End If
        FTEMP_VAL = (JTEMP_VAL + GTEMP_VAL) * (JTEMP_VAL - GTEMP_VAL) + SHIFT_VAL
        NTEMP_VAL = JTEMP_VAL * KTEMP_VAL
        ' Chase zeros.
        For j = k To P_VAL - 2
            CTEMP_VAL = HYPOT_FUNC(FTEMP_VAL, NTEMP_VAL)
            DTEMP_VAL = FTEMP_VAL / CTEMP_VAL
            ETEMP_VAL = NTEMP_VAL / CTEMP_VAL
            If (j <> k) Then: ETEMP_ARR(j - 1) = CTEMP_VAL
            FTEMP_VAL = DTEMP_VAL * STEMP_ARR(j) + ETEMP_VAL * ETEMP_ARR(j)
            ETEMP_ARR(j) = DTEMP_VAL * ETEMP_ARR(j) - ETEMP_VAL * STEMP_ARR(j)
            NTEMP_VAL = ETEMP_VAL * STEMP_ARR(j + 1)
            STEMP_ARR(j + 1) = DTEMP_VAL * STEMP_ARR(j + 1)
            If (V_VAL) Then
                For i = 0 To NCOLUMNS - 1
                    CTEMP_VAL = DTEMP_VAL * VTEMP_MATRIX(i, j) + _
                                ETEMP_VAL * VTEMP_MATRIX(i, j + 1)
                    VTEMP_MATRIX(i, j + 1) = -ETEMP_VAL * VTEMP_MATRIX(i, j) + _
                                DTEMP_VAL * VTEMP_MATRIX(i, j + 1)
                    VTEMP_MATRIX(i, j) = CTEMP_VAL
                Next i
            End If
            CTEMP_VAL = HYPOT_FUNC(FTEMP_VAL, NTEMP_VAL)
            DTEMP_VAL = FTEMP_VAL / CTEMP_VAL
            ETEMP_VAL = NTEMP_VAL / CTEMP_VAL
            STEMP_ARR(j) = CTEMP_VAL
            FTEMP_VAL = DTEMP_VAL * ETEMP_ARR(j) + ETEMP_VAL * STEMP_ARR(j + 1)
            STEMP_ARR(j + 1) = -ETEMP_VAL * ETEMP_ARR(j) + _
                        DTEMP_VAL * STEMP_ARR(j + 1)
            NTEMP_VAL = ETEMP_VAL * ETEMP_ARR(j + 1)
            ETEMP_ARR(j + 1) = DTEMP_VAL * ETEMP_ARR(j + 1)
            If ((U_VAL > 0) And (j < NROWS - 1)) Then
                For i = 0 To NROWS - 1
                    CTEMP_VAL = DTEMP_VAL * UTEMP_MATRIX(i, j) + _
                                ETEMP_VAL * UTEMP_MATRIX(i, j + 1)
                    UTEMP_MATRIX(i, j + 1) = -ETEMP_VAL * UTEMP_MATRIX(i, j) + _
                                DTEMP_VAL * UTEMP_MATRIX(i, j + 1)
                    UTEMP_MATRIX(i, j) = CTEMP_VAL
                Next i
            End If
        Next j
        ETEMP_ARR(P_VAL - 2) = FTEMP_VAL
        nLOOPS = nLOOPS + 1
'----------------------------------------------------------------------------------
    ElseIf VERSION = 4 Then
    'Make the singular values positive.
    'if ETEMP_ARR(P_VAL-1) is negligible (convergence).
'----------------------------------------------------------------------------------
        If (STEMP_ARR(k) <= 0) Then
            If STEMP_ARR(k) < 0 Then
                STEMP_ARR(k) = -STEMP_ARR(k)
            Else
                STEMP_ARR(k) = 0
            End If
            If (V_VAL) Then
                For i = 0 To Q_VAL
                    VTEMP_MATRIX(i, k) = -VTEMP_MATRIX(i, k)
                Next i
            End If
        End If
        ' Order the singular values.
        Do While (k < Q_VAL)
            If (STEMP_ARR(k) >= STEMP_ARR(k + 1)) Then: Exit Do
            CTEMP_VAL = STEMP_ARR(k)
            STEMP_ARR(k) = STEMP_ARR(k + 1)
            STEMP_ARR(k + 1) = CTEMP_VAL
            If (V_VAL And (k < NCOLUMNS - 1)) Then
                For i = 0 To NCOLUMNS - 1
                    CTEMP_VAL = VTEMP_MATRIX(i, k + 1)
                    VTEMP_MATRIX(i, k + 1) = VTEMP_MATRIX(i, k)
                    VTEMP_MATRIX(i, k) = CTEMP_VAL
                Next i
            End If
            If (U_VAL And (k < NROWS - 1)) Then
                For i = 0 To NROWS - 1
                    CTEMP_VAL = UTEMP_MATRIX(i, k + 1)
                    UTEMP_MATRIX(i, k + 1) = UTEMP_MATRIX(i, k)
                    UTEMP_MATRIX(i, k) = CTEMP_VAL
                Next i
            End If
            k = k + 1
        Loop
        nLOOPS = 0
        P_VAL = P_VAL - 1
'--------------------------------------------------------------------------------------
    End If
'--------------------------------------------------------------------------------------
Loop
'--------------------------------------------------------------------------------------
  VTEMP_MATRIX = MATRIX_CHANGE_BASE_ONE_FUNC(VTEMP_MATRIX)
            
  ReDim TEMP_MATRIX(0 To NCOLUMNS - 1, 0 To NCOLUMNS - 1)
  For i = 0 To NCOLUMNS - 1
      For j = 0 To NCOLUMNS - 1
          TEMP_MATRIX(i, j) = 0
      Next j
      TEMP_MATRIX(i, i) = STEMP_ARR(i)
  Next i
  STEMP_MATRIX = MATRIX_CHANGE_BASE_ONE_FUNC(TEMP_MATRIX)
            
            
  NSIZE = MINIMUM_FUNC(NROWS + 1, NCOLUMNS)
  'my code for temp fix begins here
  If NSIZE > UBound(UTEMP_MATRIX, 2) + 1 Then
    NSIZE = UBound(UTEMP_MATRIX, 2) + 1
  End If
  ReDim TEMP_MATRIX(0 To NROWS - 1, 0 To NSIZE - 1)
  For i = 0 To NROWS - 1
    For j = 0 To NSIZE - 1
      TEMP_MATRIX(i, j) = UTEMP_MATRIX(i, j)
    Next j
  Next i
  TEMP_MATRIX = MATRIX_CHANGE_BASE_ONE_FUNC(TEMP_MATRIX)
  UTEMP_MATRIX = TEMP_MATRIX
  
  Select Case OUTPUT
  Case 0
    ReDim TEMP_GROUP(1 To 3) As Variant
    TEMP_GROUP(1) = UTEMP_MATRIX
    TEMP_GROUP(2) = VTEMP_MATRIX
    TEMP_GROUP(3) = STEMP_MATRIX
    MATRIX_SVD_FACT_FUNC = TEMP_GROUP
  Case 1
    MATRIX_SVD_FACT_FUNC = UTEMP_MATRIX
  Case 2
    MATRIX_SVD_FACT_FUNC = VTEMP_MATRIX
  Case Else
    MATRIX_SVD_FACT_FUNC = STEMP_MATRIX
  End Select
            
Exit Function
ERROR_LABEL:
MATRIX_SVD_FACT_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : MATRIX_SVD_INVERSE_FUNC
'DESCRIPTION   : INVERSE OF A MATRIX - USING SVD
'LIBRARY       : MATRIX
'GROUP         : SVD
'ID            : 002
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/19/2009
'************************************************************************************
'************************************************************************************

Function MATRIX_SVD_INVERSE_FUNC(ByRef DATA_RNG As Variant)
  
  Dim i As Long
  Dim j As Long
    
  Dim STEMP_VAL As Double
  Dim VTEMP_VAL As Double
  
  Dim DATA_MATRIX As Variant
  Dim DATA_GROUP As Variant
  
  Dim UTEMP_MATRIX As Variant
  Dim VTEMP_MATRIX As Variant
  Dim STEMP_MATRIX As Variant
  
  On Error GoTo ERROR_LABEL
  
  DATA_GROUP = MATRIX_SVD_FACT_FUNC(DATA_RNG, 0)

  UTEMP_MATRIX = DATA_GROUP(1)
  VTEMP_MATRIX = DATA_GROUP(2)
  STEMP_MATRIX = DATA_GROUP(3)
    
  If ((UBound(STEMP_MATRIX, 1) = 1) And (UBound(STEMP_MATRIX, 2) = 1)) Then
    STEMP_VAL = 1 / STEMP_MATRIX(1, 1)
    VTEMP_VAL = VTEMP_MATRIX(1, 1)
    ReDim DATA_MATRIX(1 To UBound(UTEMP_MATRIX, 1), 1 To UBound(UTEMP_MATRIX, 2))
      For i = 1 To UBound(UTEMP_MATRIX, 1)
        For j = 1 To UBound(UTEMP_MATRIX, 2)
          DATA_MATRIX(i, j) = UTEMP_MATRIX(i, j) * (VTEMP_VAL * STEMP_VAL)
        Next j
      Next i
  Else
    ReDim DATA_MATRIX(1 To UBound(STEMP_MATRIX, 1), 1 To UBound(STEMP_MATRIX, 2))
       For i = 1 To UBound(STEMP_MATRIX, 1)
          DATA_MATRIX(i, i) = 1 / STEMP_MATRIX(i, i)
       Next i
    DATA_MATRIX = MMULT_FUNC(MMULT_FUNC(VTEMP_MATRIX, DATA_MATRIX, 70), _
               MATRIX_TRANSPOSE_FUNC(UTEMP_MATRIX), 70)
  End If
  
  MATRIX_SVD_INVERSE_FUNC = DATA_MATRIX

Exit Function
ERROR_LABEL:
MATRIX_SVD_INVERSE_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : MATRIX_SVD_MOORE_PENROSE_INVERSE_FUNC
'DESCRIPTION   : Computes the Moore-Penrose pseudo-inverse of a (n x m) matrix
'Def: the minimum-norm least squares solution x to a linear system:
'Ax=b --> min||Ax-b|| ; is the vector x = (A^t A)^-1 A^t =A+ b

'The matrix is called the pseudo-inverse of A if the matrix A has
'dimension (n x m), its pseudo-inverse has dimension (m x n)
'One of the most important applications of the SVD decomposition is
'A = U D V^t; A+ = V D^-1 U^t

'Note the pseudo-inverse coincides with the inverse for non-singular
'square matrices

'LIBRARY       : MATRIX
'GROUP         : SVD
'ID            : 003
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/19/2009
'************************************************************************************
'************************************************************************************

Function MATRIX_SVD_MOORE_PENROSE_INVERSE_FUNC(ByRef DATA_RNG As Variant, _
Optional ByVal epsilon As Double = 2 * 10 ^ -15)

Dim i As Long
Dim NROWS As Long
Dim NCOLUMNS As Long

Dim DATA_GROUP As Variant
Dim TEMP_MATRIX As Variant

Dim UDATA_MATRIX As Variant
Dim SDATA_MATRIX As Variant
Dim VDATA_MATRIX As Variant

On Error GoTo ERROR_LABEL

UDATA_MATRIX = DATA_RNG

NROWS = UBound(UDATA_MATRIX, 1)
NCOLUMNS = UBound(UDATA_MATRIX, 2)

DATA_GROUP = MATRIX_SVD_DECOMPOSITION_FUNC(UDATA_MATRIX, NROWS, NCOLUMNS, 1)
UDATA_MATRIX = DATA_GROUP(LBound(DATA_GROUP))
VDATA_MATRIX = DATA_GROUP(LBound(DATA_GROUP) + 1)
SDATA_MATRIX = DATA_GROUP(LBound(DATA_GROUP) + 2)
Erase DATA_GROUP

'invert the diagonal element
For i = 1 To UBound(SDATA_MATRIX, 1)
    If Abs(SDATA_MATRIX(i, 1)) < epsilon Then
          SDATA_MATRIX(i, 1) = 0
    Else
          SDATA_MATRIX(i, 1) = 1 / SDATA_MATRIX(i, 1)
    End If
Next i

TEMP_MATRIX = MMULT_FUNC(VECTOR_DIAGONAL_MATRIX_FUNC((SDATA_MATRIX)), _
              MATRIX_TRANSPOSE_FUNC(UDATA_MATRIX), 70)

TEMP_MATRIX = MMULT_FUNC(VDATA_MATRIX, TEMP_MATRIX, 70)

MATRIX_SVD_MOORE_PENROSE_INVERSE_FUNC = TEMP_MATRIX

Exit Function
ERROR_LABEL:
MATRIX_SVD_MOORE_PENROSE_INVERSE_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : MATRIX_SVD_DECOMPOSITION_FUNC
'DESCRIPTION   : SVD Routine
'LIBRARY       : MATRIX
'GROUP         : SVD
'ID            : 004
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/19/2009
'************************************************************************************
'************************************************************************************

Function MATRIX_SVD_DECOMPOSITION_FUNC(ByRef DATA_RNG As Variant, _
Optional ByRef NROWS As Long = 0, _
Optional ByRef NCOLUMNS As Long = 0, _
Optional ByVal OUTPUT As Integer = 0)

Dim i As Single
Dim j As Single

Dim ii As Single
Dim jj As Single

Dim k As Single
Dim kk As Single

Dim nLOOPS As Single

Dim NORM_VAL As Double
Dim SCALE_VAL As Double

Dim ATEMP_VAL As Double
Dim BTEMP_VAL As Double
Dim CTEMP_VAL As Double
Dim DTEMP_VAL As Double
Dim ETEMP_VAL As Double
Dim FTEMP_VAL As Double
Dim GTEMP_VAL As Double
Dim HTEMP_VAL As Double

Dim TEMP_VECTOR As Variant
Dim UDATA_MATRIX As Variant
Dim VDATA_MATRIX As Variant
Dim SDATA_MATRIX As Variant

Dim EXCHANGE_FLAG As Boolean

On Error GoTo ERROR_LABEL

UDATA_MATRIX = DATA_RNG
If NROWS = 0 Then: NROWS = UBound(UDATA_MATRIX, 1)
If NCOLUMNS = 0 Then: NCOLUMNS = UBound(UDATA_MATRIX, 2)

ReDim TEMP_VECTOR(1 To NCOLUMNS, 1 To 1)
ReDim SDATA_MATRIX(1 To NCOLUMNS, 1 To 1)
ReDim VDATA_MATRIX(1 To NCOLUMNS, 1 To NCOLUMNS)

CTEMP_VAL = 0 'Householder reduction to bidiagonal form.
SCALE_VAL = 0
NORM_VAL = 0
For i = 1 To NCOLUMNS
  ii = i + 1
  TEMP_VECTOR(i, 1) = SCALE_VAL * CTEMP_VAL
  CTEMP_VAL = 0
  ETEMP_VAL = 0
  SCALE_VAL = 0
  If (i <= NROWS) Then
    For k = i To NROWS
      SCALE_VAL = SCALE_VAL + Abs(UDATA_MATRIX(k, i))
    Next k
    If (SCALE_VAL <> 0) Then
      For k = i To NROWS
        UDATA_MATRIX(k, i) = UDATA_MATRIX(k, i) / SCALE_VAL
        ETEMP_VAL = ETEMP_VAL + UDATA_MATRIX(k, i) * UDATA_MATRIX(k, i)
      Next k
      BTEMP_VAL = UDATA_MATRIX(i, i)
      CTEMP_VAL = -1 * IIf(BTEMP_VAL >= 0, Abs(Sqr(ETEMP_VAL)), -Abs(Sqr(ETEMP_VAL)))
      DTEMP_VAL = BTEMP_VAL * CTEMP_VAL - ETEMP_VAL
      UDATA_MATRIX(i, i) = BTEMP_VAL - CTEMP_VAL
      For j = ii To NCOLUMNS
        ETEMP_VAL = 0
        For k = i To NROWS
          ETEMP_VAL = ETEMP_VAL + UDATA_MATRIX(k, i) * UDATA_MATRIX(k, j)
        Next k
        BTEMP_VAL = ETEMP_VAL / DTEMP_VAL     'bug
        For k = i To NROWS
          UDATA_MATRIX(k, j) = UDATA_MATRIX(k, j) + BTEMP_VAL * UDATA_MATRIX(k, i)
        Next k
      Next j
      For k = i To NROWS
        UDATA_MATRIX(k, i) = SCALE_VAL * UDATA_MATRIX(k, i)
      Next k
    End If
  End If
  SDATA_MATRIX(i, 1) = SCALE_VAL * CTEMP_VAL
  CTEMP_VAL = 0
  ETEMP_VAL = 0
  SCALE_VAL = 0
  If ((i <= NROWS) And (i <> NCOLUMNS)) Then
    For k = ii To NCOLUMNS
      SCALE_VAL = SCALE_VAL + Abs(UDATA_MATRIX(i, k))
    Next k
    If (SCALE_VAL <> 0) Then
      For k = ii To NCOLUMNS
        UDATA_MATRIX(i, k) = UDATA_MATRIX(i, k) / SCALE_VAL
        ETEMP_VAL = ETEMP_VAL + UDATA_MATRIX(i, k) * UDATA_MATRIX(i, k)
      Next k
      BTEMP_VAL = UDATA_MATRIX(i, ii)
      CTEMP_VAL = -1 * IIf(BTEMP_VAL >= 0, Abs(Sqr(ETEMP_VAL)), -Abs(Sqr(ETEMP_VAL)))
      DTEMP_VAL = BTEMP_VAL * CTEMP_VAL - ETEMP_VAL
      UDATA_MATRIX(i, ii) = BTEMP_VAL - CTEMP_VAL
      For k = ii To NCOLUMNS
        TEMP_VECTOR(k, 1) = UDATA_MATRIX(i, k) / DTEMP_VAL
      Next k
      For j = ii To NROWS
        ETEMP_VAL = 0
        For k = ii To NCOLUMNS
          ETEMP_VAL = ETEMP_VAL + UDATA_MATRIX(j, k) * UDATA_MATRIX(i, k)
        Next k
        For k = ii To NCOLUMNS
          UDATA_MATRIX(j, k) = UDATA_MATRIX(j, k) + ETEMP_VAL * TEMP_VECTOR(k, 1)
        Next k
      Next j
      For k = ii To NCOLUMNS
        UDATA_MATRIX(i, k) = SCALE_VAL * UDATA_MATRIX(i, k)
      Next k
    End If
  End If
  NORM_VAL = MAXIMUM_FUNC(NORM_VAL, (Abs(SDATA_MATRIX(i, 1)) + Abs(TEMP_VECTOR(i, 1))))
Next i

For i = NCOLUMNS To 1 Step -1 'Accumulation of right-hand transformations.
   If (i < NCOLUMNS) Then
      If (CTEMP_VAL <> 0) Then
         For j = ii To NCOLUMNS 'Double division to avoid possible under ow.
            VDATA_MATRIX(j, i) = (UDATA_MATRIX(i, j) / _
                UDATA_MATRIX(i, ii)) / CTEMP_VAL
         Next j
         For j = ii To NCOLUMNS
            ETEMP_VAL = 0
            For k = ii To NCOLUMNS
               ETEMP_VAL = ETEMP_VAL + UDATA_MATRIX(i, k) * VDATA_MATRIX(k, j)
            Next k
            For k = ii To NCOLUMNS
               VDATA_MATRIX(k, j) = VDATA_MATRIX(k, j) + _
                ETEMP_VAL * VDATA_MATRIX(k, i)
            Next k
         Next j
      End If
      For j = ii To NCOLUMNS
         VDATA_MATRIX(i, j) = 0
         VDATA_MATRIX(j, i) = 0
      Next j
   End If
   VDATA_MATRIX(i, i) = 1
   CTEMP_VAL = TEMP_VECTOR(i, 1)
   ii = i
Next i

For i = MINIMUM_FUNC(NROWS, NCOLUMNS) To 1 Step -1 'Accumulation of
'left-hand transformations.
   ii = i + 1
   CTEMP_VAL = SDATA_MATRIX(i, 1)
   For j = ii To NCOLUMNS
      UDATA_MATRIX(i, j) = 0
   Next j
   If (CTEMP_VAL <> 0) Then
      CTEMP_VAL = 1 / CTEMP_VAL
      For j = ii To NCOLUMNS
         ETEMP_VAL = 0
         For k = ii To NROWS
            ETEMP_VAL = ETEMP_VAL + UDATA_MATRIX(k, i) * UDATA_MATRIX(k, j)
         Next k
         BTEMP_VAL = (ETEMP_VAL / UDATA_MATRIX(i, i)) * CTEMP_VAL
         For k = i To NROWS
            UDATA_MATRIX(k, j) = UDATA_MATRIX(k, j) + BTEMP_VAL * UDATA_MATRIX(k, i)
         Next k
      Next j
      For j = i To NROWS
         UDATA_MATRIX(j, i) = UDATA_MATRIX(j, i) * CTEMP_VAL
      Next j
   Else
      For j = i To NROWS
         UDATA_MATRIX(j, i) = 0
      Next j
   End If
   UDATA_MATRIX(i, i) = UDATA_MATRIX(i, i) + 1
Next i

For k = NCOLUMNS To 1 Step -1 'Diagonalization of the bidiagonal
'form: Loop over singular values, and over allowed iterations.
    For nLOOPS = 1 To 30
       For ii = k To 1 Step -1 'Test for splitting.
          kk = ii - 1
          If ((Abs(TEMP_VECTOR(ii, 1)) + NORM_VAL) = NORM_VAL) _
            Then GoTo 1984
          If ((Abs(SDATA_MATRIX(kk, 1)) + NORM_VAL) = NORM_VAL) _
            Then GoTo 1983
       Next ii
1983:
    ATEMP_VAL = 0 'Cancellation  if ii > 1.
    ETEMP_VAL = 1
    For i = ii To k
       BTEMP_VAL = ETEMP_VAL * TEMP_VECTOR(i, 1)
       TEMP_VECTOR(i, 1) = ATEMP_VAL * TEMP_VECTOR(i, 1)
       If ((Abs(BTEMP_VAL) + NORM_VAL) = NORM_VAL) Then GoTo 1984
       CTEMP_VAL = SDATA_MATRIX(i, 1)
       DTEMP_VAL = PYTHAG_FUNC(BTEMP_VAL, CTEMP_VAL)
       SDATA_MATRIX(i, 1) = DTEMP_VAL
       DTEMP_VAL = 1 / DTEMP_VAL
       ATEMP_VAL = (CTEMP_VAL * DTEMP_VAL)
       ETEMP_VAL = -(BTEMP_VAL * DTEMP_VAL)
       For j = 1 To NROWS
          GTEMP_VAL = UDATA_MATRIX(j, kk)
          HTEMP_VAL = UDATA_MATRIX(j, i)
          UDATA_MATRIX(j, kk) = (GTEMP_VAL * ATEMP_VAL) + (HTEMP_VAL * ETEMP_VAL)
          UDATA_MATRIX(j, i) = -(GTEMP_VAL * ETEMP_VAL) + (HTEMP_VAL * ATEMP_VAL)
       Next j
    Next i
1984:
    HTEMP_VAL = SDATA_MATRIX(k, 1)
    If (ii = k) Then 'Convergence.
       If (HTEMP_VAL < 0) Then 'Singular value is made nonnegative.
           SDATA_MATRIX(k, 1) = -HTEMP_VAL
           For j = 1 To NCOLUMNS
              VDATA_MATRIX(j, k) = -VDATA_MATRIX(j, k)
           Next j
       End If
       GoTo 1985
    End If
    
    If (nLOOPS = 30) Then GoTo ERROR_LABEL 'no convergence in MATRIX_SVD_DECOMPOSITION_FUNC'
    FTEMP_VAL = SDATA_MATRIX(ii, 1) 'Shift from bottom 2-by-2 minor.
    kk = k - 1
    GTEMP_VAL = SDATA_MATRIX(kk, 1)
    CTEMP_VAL = TEMP_VECTOR(kk, 1)
    DTEMP_VAL = TEMP_VECTOR(k, 1)
    BTEMP_VAL = ((GTEMP_VAL - HTEMP_VAL) * (GTEMP_VAL + HTEMP_VAL) + (CTEMP_VAL - _
        DTEMP_VAL) * (CTEMP_VAL + DTEMP_VAL)) / (2 * DTEMP_VAL * GTEMP_VAL)
    CTEMP_VAL = PYTHAG_FUNC(BTEMP_VAL, 1)
    BTEMP_VAL = ((FTEMP_VAL - HTEMP_VAL) * (FTEMP_VAL + HTEMP_VAL) + DTEMP_VAL * _
        ((GTEMP_VAL / (BTEMP_VAL + (IIf(BTEMP_VAL >= 0, Abs(CTEMP_VAL), _
            -Abs(CTEMP_VAL))))) - DTEMP_VAL)) / FTEMP_VAL
    ATEMP_VAL = 1 'Next QR transformation:
    ETEMP_VAL = 1
    For j = ii To kk
       i = j + 1
       CTEMP_VAL = TEMP_VECTOR(i, 1)
       GTEMP_VAL = SDATA_MATRIX(i, 1)
       DTEMP_VAL = ETEMP_VAL * CTEMP_VAL
       CTEMP_VAL = ATEMP_VAL * CTEMP_VAL
       HTEMP_VAL = PYTHAG_FUNC(BTEMP_VAL, DTEMP_VAL)
       TEMP_VECTOR(j, 1) = HTEMP_VAL
       ATEMP_VAL = BTEMP_VAL / HTEMP_VAL
       ETEMP_VAL = DTEMP_VAL / HTEMP_VAL
       BTEMP_VAL = (FTEMP_VAL * ATEMP_VAL) + (CTEMP_VAL * ETEMP_VAL)
       CTEMP_VAL = -(FTEMP_VAL * ETEMP_VAL) + (CTEMP_VAL * ATEMP_VAL)
       DTEMP_VAL = GTEMP_VAL * ETEMP_VAL
       GTEMP_VAL = GTEMP_VAL * ATEMP_VAL
       For jj = 1 To NCOLUMNS
          FTEMP_VAL = VDATA_MATRIX(jj, j)
          HTEMP_VAL = VDATA_MATRIX(jj, i)
          VDATA_MATRIX(jj, j) = (FTEMP_VAL * ATEMP_VAL) + (HTEMP_VAL * ETEMP_VAL)
          VDATA_MATRIX(jj, i) = -(FTEMP_VAL * ETEMP_VAL) + (HTEMP_VAL * ATEMP_VAL)
       Next jj
       HTEMP_VAL = PYTHAG_FUNC(BTEMP_VAL, DTEMP_VAL)
       SDATA_MATRIX(j, 1) = HTEMP_VAL 'Rotation can be arbitrary if HTEMP_VAL = 0.
       If (HTEMP_VAL <> 0) Then
          HTEMP_VAL = 1 / HTEMP_VAL
          ATEMP_VAL = BTEMP_VAL * HTEMP_VAL
          ETEMP_VAL = DTEMP_VAL * HTEMP_VAL
       End If
       BTEMP_VAL = (ATEMP_VAL * CTEMP_VAL) + (ETEMP_VAL * GTEMP_VAL)
       FTEMP_VAL = -(ETEMP_VAL * CTEMP_VAL) + (ATEMP_VAL * GTEMP_VAL)
       For jj = 1 To NROWS
          GTEMP_VAL = UDATA_MATRIX(jj, j)
          HTEMP_VAL = UDATA_MATRIX(jj, i)
          UDATA_MATRIX(jj, j) = (GTEMP_VAL * ATEMP_VAL) + (HTEMP_VAL * ETEMP_VAL)
          UDATA_MATRIX(jj, i) = -(GTEMP_VAL * ETEMP_VAL) + (HTEMP_VAL * ATEMP_VAL)
       Next jj
    Next j
    TEMP_VECTOR(ii, 1) = 0
    TEMP_VECTOR(k, 1) = BTEMP_VAL
    SDATA_MATRIX(k, 1) = FTEMP_VAL
 Next nLOOPS
1985:
 'continue
Next k
    
'-----------------------------------------------------------------------------------------------
Select Case OUTPUT
'-----------------------------------------------------------------------------------------------
Case 0
'-----------------------------------------------------------------------------------------------
    MATRIX_SVD_DECOMPOSITION_FUNC = Array(UDATA_MATRIX, VDATA_MATRIX, SDATA_MATRIX)
'-----------------------------------------------------------------------------------------------
Case Else 'Descending Sort for the SVD decomposition
'-----------------------------------------------------------------------------------------------
    ii = LBound(SDATA_MATRIX, 1)
    jj = UBound(SDATA_MATRIX, 1)
    
    Do
        EXCHANGE_FLAG = False
        For i = ii To jj Step 2
            j = i + 1
            If j > jj Then Exit For
            ATEMP_VAL = Abs(SDATA_MATRIX(i, 1))
            BTEMP_VAL = Abs(SDATA_MATRIX(j, 1))
            If (ATEMP_VAL < BTEMP_VAL) Then
                CTEMP_VAL = SDATA_MATRIX(j, 1)
                SDATA_MATRIX(j, 1) = SDATA_MATRIX(i, 1)
                SDATA_MATRIX(i, 1) = CTEMP_VAL
                UDATA_MATRIX = MATRIX_SWAP_COLUMN_FUNC(UDATA_MATRIX, j, i)
                VDATA_MATRIX = MATRIX_SWAP_COLUMN_FUNC(VDATA_MATRIX, j, i)
                EXCHANGE_FLAG = True
            End If
        Next
        If ii = LBound(SDATA_MATRIX, 1) Then
            ii = LBound(SDATA_MATRIX, 1) + 1
        Else
            ii = LBound(SDATA_MATRIX, 1)
        End If
    Loop Until EXCHANGE_FLAG = False And ii = LBound(SDATA_MATRIX, 1)
    
    If OUTPUT = 1 Then
        MATRIX_SVD_DECOMPOSITION_FUNC = Array(UDATA_MATRIX, VDATA_MATRIX, SDATA_MATRIX)
    ElseIf OUTPUT = 2 Then
        MATRIX_SVD_DECOMPOSITION_FUNC = UDATA_MATRIX
    ElseIf OUTPUT = 3 Then
        MATRIX_SVD_DECOMPOSITION_FUNC = VDATA_MATRIX
    Else
        MATRIX_SVD_DECOMPOSITION_FUNC = SDATA_MATRIX
    End If

'-----------------------------------------------------------------------------------------------
End Select
'-----------------------------------------------------------------------------------------------

Exit Function
ERROR_LABEL:
MATRIX_SVD_DECOMPOSITION_FUNC = Err.number
End Function
