Attribute VB_Name = "NUMBER_COMPLEX_EIGEN_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.


'************************************************************************************
'************************************************************************************
'FUNCTION      : COMPLEX_MATRIX_QR_EIGENVALUES_FUNC

'DESCRIPTION   : Find real and complex eigenvalues of complex matrix with
'the iterative QR method.

'This function uses a reduction of the EISPACK FORTRAN COMQR and
'CORTH subroutines (April 1983). COMQR IS A TRANSLATION OF THE
'ALGOL PROCEDURE math. 12, 369-376(1968) by Martin  and Wilkinson.).

'This function performs the diagonal reduction of a given complex matrix
'with QR method, and returns the approximate eigenvalues real or complex
'as an (n x 2) array. This function supports 3 different formats:
'1 = split, 2 = interlaced, 3 = string

'LIBRARY       : NUMBER_COMPLEX
'GROUP         : EIGEN
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/08/2008
'************************************************************************************
'************************************************************************************

Function COMPLEX_MATRIX_QR_EIGENVALUES_FUNC(ByRef DATA_RNG As Variant, _
Optional ByVal CPLX_FORMAT As Integer = 1, _
Optional ByVal CPLX_CHR_STR As String = "i", _
Optional ByVal epsilon As Double = 2 * 10 ^ -14)

Dim h As Long
Dim i As Long
Dim j As Long
Dim k As Long
Dim E As Long

Dim EE As Long
Dim ff As Long '
Dim gg As Long '
Dim hh As Long
Dim ii As Long '
Dim jj As Long
Dim kk As Long

Dim NSIZE As Long
Dim ERROR_VAL As Long

Dim F_VAL As Double
Dim G_VAL As Double '
Dim H_VAL As Double

Dim SI_VAL As Double
Dim SR_VAL As Double
      
Dim TI_VAL As Double
Dim TR_VAL As Double
      
Dim XI_VAL As Double
Dim XR_VAL As Double
      
Dim YI_VAL As Double
Dim YR_VAL As Double

Dim ZI_VAL As Double
Dim ZR_VAL As Double
      
Dim NORM_VAL As Double
Dim ATEMP_VAL As Double
Dim BTEMP_VAL As Double
      
Dim IMAG_VAL As Double
Dim REAL_VAL As Double
Dim SCALE_VAL As Double

Dim ROOT_COMPLEX_OBJ As Cplx
Dim FIRST_COMPLEX_OBJ As Cplx
Dim SECOND_COMPLEX_OBJ As Cplx
Dim THIRD_COMPLEX_OBJ As Cplx
Dim FORTH_COMPLEX_OBJ As Cplx

Dim AREAL_VECTOR As Variant
Dim AIMAG_VECTOR As Variant

Dim BREAL_VECTOR As Variant
Dim BIMAG_VECTOR As Variant

Dim REAL_MATRIX As Variant
Dim IMAG_MATRIX As Variant

Dim DATA_MATRIX As Variant
Dim TEMP_MATRIX As Variant

On Error GoTo ERROR_LABEL

DATA_MATRIX = DATA_RNG
If CPLX_FORMAT = 2 Then DATA_MATRIX = _
    COMPLEX_MATRIX_FORMAT_FUNC(DATA_MATRIX, 21, CPLX_CHR_STR, epsilon)
If CPLX_FORMAT = 3 Then DATA_MATRIX = _
    COMPLEX_MATRIX_FORMAT_FUNC(DATA_MATRIX, 31, CPLX_CHR_STR, epsilon)

'check dimension. Only square matrix here
If 2 * UBound(DATA_MATRIX, 1) <> UBound(DATA_MATRIX, 2) Then: GoTo ERROR_LABEL

NSIZE = UBound(DATA_MATRIX, 1) 'matrix dimension

ReDim REAL_MATRIX(1 To NSIZE, 1 To NSIZE)
ReDim IMAG_MATRIX(1 To NSIZE, 1 To NSIZE)

For i = 1 To NSIZE
    For j = 1 To NSIZE
        REAL_MATRIX(i, j) = DATA_MATRIX(i, j)
        IMAG_MATRIX(i, j) = DATA_MATRIX(i, j + NSIZE)
    Next j
Next i

'------------------------------------------------------------------------------------
 'hessember transformation: this subroutine is a translation of a
 'complex analogue of the algol procedure orthes, num. math. 12,
 '349-368(1968) by martin and wilkinson. handbook for auto. comp.,
 'vol.ii-linear algebra, 339-358(1971).
 
' Given a complex general matrix, this subroutine reduces a submatrix
' situated in rows and columns 1 through NSIZE to upper hessenberg form
' by unitary similarity transformations.
'------------------------------------------------------------------------------------

      ReDim AREAL_VECTOR(1 To NSIZE)
      ReDim AIMAG_VECTOR(1 To NSIZE)

      kk = NSIZE - 1
      gg = 1 + 1
      hh = k
      
      If (kk < gg) Then GoTo 1982

      For k = gg To kk
         H_VAL = 0#
         AREAL_VECTOR(k) = 0#
         AIMAG_VECTOR(k) = 0#
         SCALE_VAL = 0#
         
         For i = k To NSIZE
            SCALE_VAL = SCALE_VAL + Abs(REAL_MATRIX(i, k - 1)) + _
                                    Abs(IMAG_MATRIX(i, k - 1))
         Next i

         If (SCALE_VAL = 0#) Then Exit For
         ff = k + NSIZE
         
         For ii = k To NSIZE
            i = ff - ii
            AREAL_VECTOR(i) = REAL_MATRIX(i, k - 1) / SCALE_VAL
            AIMAG_VECTOR(i) = IMAG_MATRIX(i, k - 1) / SCALE_VAL
            H_VAL = H_VAL + AREAL_VECTOR(i) * AREAL_VECTOR(i) + _
                    AIMAG_VECTOR(i) * AIMAG_VECTOR(i)
         Next ii

         G_VAL = Sqr(H_VAL)
         F_VAL = Sqr(AREAL_VECTOR(k) ^ 2 + AIMAG_VECTOR(k) ^ 2)
         If (F_VAL = 0#) Then 'GoTo 103
            AREAL_VECTOR(k) = G_VAL
            REAL_MATRIX(k, k - 1) = SCALE_VAL
         Else
            H_VAL = H_VAL + F_VAL * G_VAL
            G_VAL = G_VAL / F_VAL
            AREAL_VECTOR(k) = (1# + G_VAL) * AREAL_VECTOR(k)
            AIMAG_VECTOR(k) = (1# + G_VAL) * AIMAG_VECTOR(k)
         End If
         
         For j = k To NSIZE
            REAL_VAL = 0#
            IMAG_VAL = 0#
            
            For ii = k To NSIZE
               i = ff - ii
               REAL_VAL = REAL_VAL + AREAL_VECTOR(i) * REAL_MATRIX(i, j) + _
                         AIMAG_VECTOR(i) * IMAG_MATRIX(i, j)
               IMAG_VAL = IMAG_VAL + AREAL_VECTOR(i) * IMAG_MATRIX(i, j) - _
                         AIMAG_VECTOR(i) * REAL_MATRIX(i, j)
            Next ii
            REAL_VAL = REAL_VAL / H_VAL
            IMAG_VAL = IMAG_VAL / H_VAL
            For i = k To NSIZE
               REAL_MATRIX(i, j) = REAL_MATRIX(i, j) - REAL_VAL * _
                                   AREAL_VECTOR(i) + IMAG_VAL * AIMAG_VECTOR(i)
               IMAG_MATRIX(i, j) = IMAG_MATRIX(i, j) - REAL_VAL * _
                                  AIMAG_VECTOR(i) - IMAG_VAL * AREAL_VECTOR(i)
            Next i
         Next j
         
         For i = 1 To NSIZE
            REAL_VAL = 0#
            IMAG_VAL = 0#
            
            For jj = k To NSIZE
               j = ff - jj
               REAL_VAL = REAL_VAL + AREAL_VECTOR(j) * REAL_MATRIX(i, j) - _
                         AIMAG_VECTOR(j) * IMAG_MATRIX(i, j)
               IMAG_VAL = IMAG_VAL + AREAL_VECTOR(j) * IMAG_MATRIX(i, j) + _
                         AIMAG_VECTOR(j) * REAL_MATRIX(i, j)
            
            Next jj
            REAL_VAL = REAL_VAL / H_VAL
            IMAG_VAL = IMAG_VAL / H_VAL
            For j = k To NSIZE
               REAL_MATRIX(i, j) = REAL_MATRIX(i, j) - REAL_VAL * _
                                  AREAL_VECTOR(j) - IMAG_VAL * AIMAG_VECTOR(j)
               IMAG_MATRIX(i, j) = IMAG_MATRIX(i, j) + REAL_VAL * _
                                  AIMAG_VECTOR(j) - IMAG_VAL * AREAL_VECTOR(j)
            Next j
         Next i
         
         AREAL_VECTOR(k) = SCALE_VAL * AREAL_VECTOR(k)
         AIMAG_VECTOR(k) = SCALE_VAL * AIMAG_VECTOR(k)
         REAL_MATRIX(k, k - 1) = -G_VAL * REAL_MATRIX(k, k - 1)
         IMAG_MATRIX(k, k - 1) = -G_VAL * IMAG_MATRIX(k, k - 1)
      Next k
1982:

'------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------
' This subroutine is a translation of a unitary analogue of the algol procedure
' comlr, num. math. 12, 369-376(1968) by martin and wilkinson. Handbook for auto.
' comp., vol.ii-linear algebra, 396-403(1971). The unitary analogue substitutes
' the qr algorithm of francis (comp. jour. 4, 332-345(1962)) for the lr algorithm.
' This subroutine finds the eigenvalues of a complex upper hessenberg matrix by
' the qr method.
'-----------------------------------------------------------------------------------

      ReDim BREAL_VECTOR(1 To NSIZE)
      ReDim BIMAG_VECTOR(1 To NSIZE)

      ERROR_VAL = 0
      jj = NSIZE
      If (1 = NSIZE) Then GoTo 1984
      
      k = 1 + 1
      
      For i = k To NSIZE
         ii = MINIMUM_FUNC(i + 1, NSIZE)
         If (IMAG_MATRIX(i, i - 1) = 0#) Then GoTo 1983
         NORM_VAL = PYTHAG_FUNC(REAL_MATRIX(i, i - 1), IMAG_MATRIX(i, i - 1))
         YR_VAL = REAL_MATRIX(i, i - 1) / NORM_VAL
         YI_VAL = IMAG_MATRIX(i, i - 1) / NORM_VAL
         REAL_MATRIX(i, i - 1) = NORM_VAL
         IMAG_MATRIX(i, i - 1) = 0#
'
         For j = i To NSIZE
            SI_VAL = YR_VAL * IMAG_MATRIX(i, j) - _
                YI_VAL * REAL_MATRIX(i, j)
            REAL_MATRIX(i, j) = YR_VAL * REAL_MATRIX(i, j) + _
                YI_VAL * IMAG_MATRIX(i, j)
            IMAG_MATRIX(i, j) = SI_VAL
         Next j
'
         For j = 1 To ii
            SI_VAL = YR_VAL * IMAG_MATRIX(j, i) + _
                YI_VAL * REAL_MATRIX(j, i)
            REAL_MATRIX(j, i) = YR_VAL * REAL_MATRIX(j, i) _
                - YI_VAL * IMAG_MATRIX(j, i)
            IMAG_MATRIX(j, i) = SI_VAL
         Next j
1983:   Next i
'     .......... store roots isolated by cbal ..........
1984:  For i = 1 To NSIZE
         If (i >= 1 And i <= NSIZE) Then GoTo 1985
         BREAL_VECTOR(i) = REAL_MATRIX(i, i)
         BIMAG_VECTOR(i) = IMAG_MATRIX(i, i)
1985:  Next i
'
      kk = NSIZE
      TR_VAL = 0#
      TI_VAL = 0#
      E = 30 * NSIZE
'     .......... search for next eigenvalue ..........
Do
1986:  If (kk < 1) Then GoTo 1999
      EE = 0
      hh = kk - 1
'     .......... look for single small sub-diagonal element
'                for l=en step -1 until 1 d0 -- ..........
    Do
1987:  For ii = 1 To kk
         k = kk + 1 - ii
         If (k = 1) Then GoTo 1989
         ATEMP_VAL = Abs(REAL_MATRIX(k - 1, k - 1)) + _
            Abs(IMAG_MATRIX(k - 1, k - 1)) + Abs(REAL_MATRIX(k, k)) + _
                Abs(IMAG_MATRIX(k, k))
         BTEMP_VAL = ATEMP_VAL + Abs(REAL_MATRIX(k, k - 1))
         If (BTEMP_VAL = ATEMP_VAL) Then GoTo 1989
1988:  Next ii
'     .......... form shift ..........
1989:  If (k = kk) Then GoTo 1997
      If (E = 0) Then GoTo 1998
      If (EE = 10 Or EE = 20) Then GoTo 1991
      
      SR_VAL = REAL_MATRIX(kk, kk)
      SI_VAL = IMAG_MATRIX(kk, kk)
      XR_VAL = REAL_MATRIX(hh, kk) * REAL_MATRIX(kk, hh)
      XI_VAL = IMAG_MATRIX(hh, kk) * REAL_MATRIX(kk, hh)
      
      If (XR_VAL = 0# And XI_VAL = 0#) Then GoTo 1992
      
      YR_VAL = (REAL_MATRIX(hh, hh) - SR_VAL) / 2#
      YI_VAL = (IMAG_MATRIX(hh, hh) - SI_VAL) / 2#
      
      FIRST_COMPLEX_OBJ.reel = YR_VAL ^ 2 - YI_VAL ^ 2 + XR_VAL
      FIRST_COMPLEX_OBJ.imag = 2# * YR_VAL * YI_VAL + XI_VAL
      
      ROOT_COMPLEX_OBJ = COMPLEX_ROOT_OBJ_FUNC(FIRST_COMPLEX_OBJ, 2)
      ZR_VAL = ROOT_COMPLEX_OBJ.reel
      ZI_VAL = ROOT_COMPLEX_OBJ.imag
      
      If (YR_VAL * ZR_VAL + YI_VAL * ZI_VAL >= 0#) Then GoTo 1990
      ZR_VAL = -ZR_VAL
      ZI_VAL = -ZI_VAL
1990:
      
      SECOND_COMPLEX_OBJ.reel = XR_VAL
      SECOND_COMPLEX_OBJ.imag = XI_VAL
      
      THIRD_COMPLEX_OBJ.reel = YR_VAL + ZR_VAL
      THIRD_COMPLEX_OBJ.imag = YI_VAL + ZI_VAL
      FORTH_COMPLEX_OBJ = COMPLEX_QUOTIENT_OBJ_FUNC(SECOND_COMPLEX_OBJ, THIRD_COMPLEX_OBJ)

      XR_VAL = FORTH_COMPLEX_OBJ.reel
      XI_VAL = FORTH_COMPLEX_OBJ.imag

      SR_VAL = SR_VAL - XR_VAL
      SI_VAL = SI_VAL - XI_VAL
      GoTo 1992
'     .......... form exceptional shift ..........
1991:  SR_VAL = Abs(REAL_MATRIX(kk, hh)) + _
        Abs(REAL_MATRIX(hh, kk - 2))
       SI_VAL = 0#
'
1992:  For i = 1 To kk
         REAL_MATRIX(i, i) = REAL_MATRIX(i, i) - SR_VAL
         IMAG_MATRIX(i, i) = IMAG_MATRIX(i, i) - SI_VAL
1993:  Next i
'
      TR_VAL = TR_VAL + SR_VAL
      TI_VAL = TI_VAL + SI_VAL
      EE = EE + 1
      E = E - 1
      
'     .......... reduce to triangle (rows) ..........
      h = k + 1
'
      For i = h To kk
         SR_VAL = REAL_MATRIX(i, i - 1)
         REAL_MATRIX(i, i - 1) = 0#
         NORM_VAL = PYTHAG_FUNC(PYTHAG_FUNC(REAL_MATRIX(i - 1, i - 1), _
                    IMAG_MATRIX(i - 1, i - 1)), SR_VAL)
         
         XR_VAL = REAL_MATRIX(i - 1, i - 1) / NORM_VAL
         BREAL_VECTOR(i - 1) = XR_VAL
         XI_VAL = IMAG_MATRIX(i - 1, i - 1) / NORM_VAL
         BIMAG_VECTOR(i - 1) = XI_VAL
         REAL_MATRIX(i - 1, i - 1) = NORM_VAL
         IMAG_MATRIX(i - 1, i - 1) = 0#
         IMAG_MATRIX(i, i - 1) = SR_VAL / NORM_VAL
'
         For j = i To kk
            YR_VAL = REAL_MATRIX(i - 1, j)
            YI_VAL = IMAG_MATRIX(i - 1, j)
            ZR_VAL = REAL_MATRIX(i, j)
            ZI_VAL = IMAG_MATRIX(i, j)
            REAL_MATRIX(i - 1, j) = XR_VAL * YR_VAL + _
                    XI_VAL * YI_VAL + IMAG_MATRIX(i, i - 1) * ZR_VAL
            IMAG_MATRIX(i - 1, j) = XR_VAL * YI_VAL - _
                    XI_VAL * YR_VAL + IMAG_MATRIX(i, i - 1) * ZI_VAL
            REAL_MATRIX(i, j) = XR_VAL * ZR_VAL - XI_VAL * _
                    ZI_VAL - IMAG_MATRIX(i, i - 1) * YR_VAL
            IMAG_MATRIX(i, j) = XR_VAL * ZI_VAL + XI_VAL * _
                    ZR_VAL - IMAG_MATRIX(i, i - 1) * YI_VAL
         Next j
      Next i
'
      SI_VAL = IMAG_MATRIX(kk, kk)
      If (SI_VAL = 0#) Then GoTo 1994
      NORM_VAL = PYTHAG_FUNC(REAL_MATRIX(kk, kk), SI_VAL)
      SR_VAL = REAL_MATRIX(kk, kk) / NORM_VAL
      SI_VAL = SI_VAL / NORM_VAL
      REAL_MATRIX(kk, kk) = NORM_VAL
      IMAG_MATRIX(kk, kk) = 0#
'     .......... inverse operation (columns) ..........
1994:  For j = h To kk
         XR_VAL = BREAL_VECTOR(j - 1)
         XI_VAL = BIMAG_VECTOR(j - 1)
'
         For i = k To j
            YR_VAL = REAL_MATRIX(i, j - 1)
            YI_VAL = 0#
            ZR_VAL = REAL_MATRIX(i, j)
            ZI_VAL = IMAG_MATRIX(i, j)
            If (i = j) Then GoTo 1995
            YI_VAL = IMAG_MATRIX(i, j - 1)
            IMAG_MATRIX(i, j - 1) = XR_VAL * YI_VAL + _
                    XI_VAL * YR_VAL + IMAG_MATRIX(j, j - 1) * ZI_VAL
1995:        REAL_MATRIX(i, j - 1) = XR_VAL * YR_VAL - XI_VAL * _
                    YI_VAL + IMAG_MATRIX(j, j - 1) * ZR_VAL
            REAL_MATRIX(i, j) = XR_VAL * ZR_VAL + XI_VAL * _
                    ZI_VAL - IMAG_MATRIX(j, j - 1) * YR_VAL
            IMAG_MATRIX(i, j) = XR_VAL * ZI_VAL - XI_VAL * _
                    ZR_VAL - IMAG_MATRIX(j, j - 1) * YI_VAL
         Next i
1996:  Next j
'
      If (SI_VAL <> 0#) Then
        For i = k To kk
           YR_VAL = REAL_MATRIX(i, kk)
           YI_VAL = IMAG_MATRIX(i, kk)
           REAL_MATRIX(i, kk) = SR_VAL * YR_VAL - SI_VAL * YI_VAL
           IMAG_MATRIX(i, kk) = SR_VAL * YI_VAL + SI_VAL * YR_VAL
        Next i
      End If

    Loop 'GoTo 1987: look for single small sub-diagonal element
'     ..........  root found ..........
1997:  BREAL_VECTOR(kk) = REAL_MATRIX(kk, kk) + TR_VAL
      BIMAG_VECTOR(kk) = IMAG_MATRIX(kk, kk) + TI_VAL
      kk = hh
'
Loop 'GoTo 1986 : search for next eigenvalue
'     .......... set error -- all eigenvalues have not
'                converged after 30*NSIZE iterations ..........
1998:
    ERROR_VAL = kk
1999:

'------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------


ReDim TEMP_MATRIX(1 To NSIZE, 1 To 2)
For i = 1 To NSIZE
    If i > ERROR_VAL Then
        TEMP_MATRIX(i, 1) = BREAL_VECTOR(i)
        TEMP_MATRIX(i, 2) = BIMAG_VECTOR(i)
    Else
        TEMP_MATRIX(i, 1) = "-"
        TEMP_MATRIX(i, 2) = "-"
    End If
Next

TEMP_MATRIX = MATRIX_TRIM_SMALL_VALUES_FUNC(TEMP_MATRIX, epsilon)
TEMP_MATRIX = MATRIX_QUICK_SORT_FUNC(TEMP_MATRIX, 1, 1)
COMPLEX_MATRIX_QR_EIGENVALUES_FUNC = TEMP_MATRIX

Exit Function
ERROR_LABEL:
COMPLEX_MATRIX_QR_EIGENVALUES_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : COMPLEX_MATRIX_INVERSE_EIGENVECTORS_FUNC

'DESCRIPTION   : Eigenvectors of complex matrix with inverse iteration
'This function returns the eigenvector associated to ist
'eigenvalue of a complex matrix A (n x n) by the inverse
'iteration algorithm. If "Eigenvalue" is a single value,
'the function returns a complex vector. Otherwise if
'"Eigenvalues" is a complex vector, the function returns
'the complex matrix of the associated eigenvectors.
'The inverse iteration method is adapt for eigenvalues affected
'by large error, because is more stable than the singular system
'resolution of the eigenvector function associated to complex eigenvalue

'LIBRARY       : NUMBER_COMPLEX
'GROUP         : EIGEN
'ID            : 002
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/08/2008
'************************************************************************************
'************************************************************************************

Function COMPLEX_MATRIX_INVERSE_EIGENVECTORS_FUNC(ByRef DATA_RNG As Variant, _
ByRef EIGEN_RNG As Variant, _
Optional ByVal CPLX_FORMAT As Integer = 1, _
Optional ByVal CPLX_CHR_STR As String = "i", _
Optional ByVal epsilon As Double = 5 * 10 ^ -15)

Dim i As Long
Dim j As Long
Dim k As Long

Dim o As Long
Dim p As Long
Dim q As Long

Dim EE As Long
Dim ff As Long
Dim hh As Long
Dim ii As Long
Dim jj As Long
Dim kk As Long

Dim pp As Long
Dim qq As Long
Dim RR As Long
Dim ss As Long

Dim F_VAL As Double '
Dim G_VAL As Double
Dim H_VAL As Double

Dim NROWS As Long
Dim NSIZE As Long
Dim ERROR_VAL As Long

Dim TEMP_ROOT As Double
Dim TEMP_SUM As Double
Dim TEMP_VAL As Double

Dim TEMP_IMAG As Double
Dim TEMP_REEL As Double

Dim TEMP_SCALE As Double

Dim IMAG_SCALE As Double
Dim REAL_SCALE As Double

Dim TEMP_NORM As Double
Dim TEMP_THRESD As Double
Dim TEMP_FACTOR As Double

Dim XTEMP_VAL As Double
Dim YTEMP_VAL As Double

Dim IMAG_LAMBDA As Double
Dim REAL_LAMBDA As Double

Dim ATEMP_VECTOR As Variant
Dim BTEMP_VECTOR As Variant

Dim AREAL_VECTOR As Variant
Dim AIMAG_VECTOR As Variant

Dim BREAL_VECTOR As Variant
Dim BIMAG_VECTOR As Variant

Dim REAL_MATRIX As Variant
Dim IMAG_MATRIX As Variant

Dim EIGEN_VECTORS As Variant
Dim DATA_MATRIX As Variant
Dim TEMP_MATRIX As Variant

Dim ORTHOG_REAL_VECTOR As Variant
Dim ORTHOG_IMAG_VECTOR As Variant

Dim ATEMP_COMPLEX_OBJ As Cplx
Dim BTEMP_COMPLEX_OBJ As Cplx
Dim CTEMP_COMPLEX_OBJ As Cplx

Dim LAMBDA As Double
Dim tolerance As Double

On Error GoTo ERROR_LABEL

LAMBDA = 10 ^ -13

DATA_MATRIX = DATA_RNG
If CPLX_FORMAT = 2 Then DATA_MATRIX = _
        COMPLEX_MATRIX_FORMAT_FUNC(DATA_MATRIX, 21, CPLX_CHR_STR, epsilon)
If CPLX_FORMAT = 3 Then DATA_MATRIX = _
        COMPLEX_MATRIX_FORMAT_FUNC(DATA_MATRIX, 31, CPLX_CHR_STR, epsilon)
'optional complex part for format=1 and square matrices.
If CPLX_FORMAT = 1 And (UBound(DATA_MATRIX, 1) = UBound(DATA_MATRIX, 2)) Then
    ReDim Preserve DATA_MATRIX(1 To UBound(DATA_MATRIX, 1), _
    1 To 2 * UBound(DATA_MATRIX, 1))
End If
'check dimension. Only square matrix here
If 2 * UBound(DATA_MATRIX, 1) <> UBound(DATA_MATRIX, 2) _
Then: GoTo ERROR_LABEL

'load complex matrix
NROWS = UBound(DATA_MATRIX, 1) 'matrix dimension
ReDim REAL_MATRIX(1 To NROWS, 1 To NROWS)
ReDim IMAG_MATRIX(1 To NROWS, 1 To NROWS)
For i = 1 To NROWS
    For j = 1 To NROWS
        REAL_MATRIX(i, j) = DATA_MATRIX(i, j)
        IMAG_MATRIX(i, j) = DATA_MATRIX(i, j + NROWS)
    Next j
Next i

'load one or more eigenvalues
If VarType(EIGEN_RNG) = vbString Then
    ReDim TEMP_MATRIX(1 To 1, 1 To 1)
    TEMP_MATRIX(1, 1) = EIGEN_RNG
ElseIf IsArray(EIGEN_RNG) = True Then
    TEMP_MATRIX = EIGEN_RNG
Else
    GoTo ERROR_LABEL
End If

If CPLX_FORMAT = 2 Then TEMP_MATRIX = _
        COMPLEX_MATRIX_FORMAT_FUNC(TEMP_MATRIX, 21, CPLX_CHR_STR, epsilon)
If CPLX_FORMAT = 3 Then TEMP_MATRIX = _
    COMPLEX_MATRIX_FORMAT_FUNC(TEMP_MATRIX, 31, CPLX_CHR_STR, epsilon)
'check dimension. Only square matrix here

If Not (UBound(TEMP_MATRIX, 1) <= NROWS And _
UBound(TEMP_MATRIX, 2) = 2) Then GoTo ERROR_LABEL

'load complex eigenvalues
ReDim AREAL_VECTOR(1 To NROWS)
ReDim AIMAG_VECTOR(1 To NROWS)
ReDim EIGEN_VECTORS(1 To NROWS)
For i = 1 To UBound(TEMP_MATRIX)
    AREAL_VECTOR(i) = TEMP_MATRIX(i, 1)
    AIMAG_VECTOR(i) = TEMP_MATRIX(i, 2)
    EIGEN_VECTORS(i) = True
Next i

'Erase DATA_MATRIX, TEMP_MATRIX

'-------------------------transform into hessemberg---------------------------------
      ReDim ORTHOG_REAL_VECTOR(1 To NROWS)
      ReDim ORTHOG_IMAG_VECTOR(1 To NROWS)


'--------------------------------------------------------------------------------
'     given a complex general matrix, this subroutine
'     reduces a submatrix situated in rows and columns
'     1 through NROWS to upper hessenberg form by
'     unitary similarity transformations.
'--------------------------------------------------------------------------------
'
      pp = NROWS - 1
      qq = 1 + 1
      EE = NSIZE
      If (pp < qq) Then GoTo 1982
'
      For NSIZE = qq To pp
         H_VAL = 0#
         ORTHOG_REAL_VECTOR(NSIZE) = 0#
         ORTHOG_IMAG_VECTOR(NSIZE) = 0#
         TEMP_SCALE = 0#
'     .......... TEMP_SCALE column (algol tol then not needed) ..........
         For i = NSIZE To NROWS
            TEMP_SCALE = TEMP_SCALE + Abs(REAL_MATRIX(i, NSIZE - 1)) _
                + Abs(IMAG_MATRIX(i, NSIZE - 1))
         Next i
'
         If (TEMP_SCALE = 0#) Then Exit For
         p = NSIZE + NROWS
'     .......... for i=igh step -1 until NSIZE do -- ..........
         For ii = NSIZE To NROWS
            i = p - ii
            ORTHOG_REAL_VECTOR(i) = REAL_MATRIX(i, NSIZE - 1) / TEMP_SCALE
            ORTHOG_IMAG_VECTOR(i) = IMAG_MATRIX(i, NSIZE - 1) / TEMP_SCALE
            H_VAL = H_VAL + ORTHOG_REAL_VECTOR(i) * ORTHOG_REAL_VECTOR(i) + _
                ORTHOG_IMAG_VECTOR(i) * ORTHOG_IMAG_VECTOR(i)
         Next ii
'
         G_VAL = Sqr(H_VAL)
         F_VAL = Sqr(ORTHOG_REAL_VECTOR(NSIZE) ^ 2 + ORTHOG_IMAG_VECTOR(NSIZE) ^ 2)
         If (F_VAL = 0#) Then 'GoTo 103
            ORTHOG_REAL_VECTOR(NSIZE) = G_VAL
            REAL_MATRIX(NSIZE, NSIZE - 1) = TEMP_SCALE
         Else
            H_VAL = H_VAL + F_VAL * G_VAL
            G_VAL = G_VAL / F_VAL
            ORTHOG_REAL_VECTOR(NSIZE) = (1# + G_VAL) * ORTHOG_REAL_VECTOR(NSIZE)
            ORTHOG_IMAG_VECTOR(NSIZE) = (1# + G_VAL) * ORTHOG_IMAG_VECTOR(NSIZE)
         End If
'     .......... form (i-(u*ut)/H_VAL) * a ..........
         For j = NSIZE To NROWS
            TEMP_REEL = 0#
            TEMP_IMAG = 0#
'     .......... for i=igh step -1 until NSIZE do -- ..........
            For ii = NSIZE To NROWS
               i = p - ii
               TEMP_REEL = TEMP_REEL + ORTHOG_REAL_VECTOR(i) * _
                    REAL_MATRIX(i, j) + ORTHOG_IMAG_VECTOR(i) * IMAG_MATRIX(i, j)
               TEMP_IMAG = TEMP_IMAG + ORTHOG_REAL_VECTOR(i) * _
                    IMAG_MATRIX(i, j) - ORTHOG_IMAG_VECTOR(i) * REAL_MATRIX(i, j)
            Next ii
            TEMP_REEL = TEMP_REEL / H_VAL
            TEMP_IMAG = TEMP_IMAG / H_VAL
            For i = NSIZE To NROWS
               REAL_MATRIX(i, j) = REAL_MATRIX(i, j) - TEMP_REEL * _
                    ORTHOG_REAL_VECTOR(i) + TEMP_IMAG * ORTHOG_IMAG_VECTOR(i)
               IMAG_MATRIX(i, j) = IMAG_MATRIX(i, j) - TEMP_REEL * _
                    ORTHOG_IMAG_VECTOR(i) - TEMP_IMAG * ORTHOG_REAL_VECTOR(i)
            Next i
         Next j
'     .......... form (i-(u*ut)/H_VAL)*a*(i-(u*ut)/H_VAL) ..........
         For i = 1 To NROWS
            TEMP_REEL = 0#
            TEMP_IMAG = 0#
'     .......... for j=igh step -1 until NSIZE do -- ..........
            For jj = NSIZE To NROWS
               j = p - jj
               TEMP_REEL = TEMP_REEL + ORTHOG_REAL_VECTOR(j) * _
                    REAL_MATRIX(i, j) - ORTHOG_IMAG_VECTOR(j) * IMAG_MATRIX(i, j)
               TEMP_IMAG = TEMP_IMAG + ORTHOG_REAL_VECTOR(j) * _
                    IMAG_MATRIX(i, j) + ORTHOG_IMAG_VECTOR(j) * REAL_MATRIX(i, j)
'            continue
            Next jj
            TEMP_REEL = TEMP_REEL / H_VAL
            TEMP_IMAG = TEMP_IMAG / H_VAL
            For j = NSIZE To NROWS
               REAL_MATRIX(i, j) = REAL_MATRIX(i, j) - TEMP_REEL * _
                    ORTHOG_REAL_VECTOR(j) - TEMP_IMAG * ORTHOG_IMAG_VECTOR(j)
               IMAG_MATRIX(i, j) = IMAG_MATRIX(i, j) + TEMP_REEL * _
                    ORTHOG_IMAG_VECTOR(j) - TEMP_IMAG * ORTHOG_REAL_VECTOR(j)
            Next j
         Next i
'
         ORTHOG_REAL_VECTOR(NSIZE) = TEMP_SCALE * ORTHOG_REAL_VECTOR(NSIZE)
         ORTHOG_IMAG_VECTOR(NSIZE) = TEMP_SCALE * ORTHOG_IMAG_VECTOR(NSIZE)
         REAL_MATRIX(NSIZE, NSIZE - 1) = -G_VAL * REAL_MATRIX(NSIZE, NSIZE - 1)
         IMAG_MATRIX(NSIZE, NSIZE - 1) = -G_VAL * IMAG_MATRIX(NSIZE, NSIZE - 1)
      Next NSIZE
1982:
'----------------eigenvector of the hessemberg by inverse-iteration-----------------
      ReDim BREAL_VECTOR(1 To NROWS, 1 To NROWS)
      ReDim BIMAG_VECTOR(1 To NROWS, 1 To NROWS)
      
      ReDim DATA_MATRIX(1 To NROWS, 1 To NROWS)
      ReDim TEMP_MATRIX(1 To NROWS, 1 To NROWS)
      
      ReDim ATEMP_VECTOR(1 To NROWS)
      ReDim BTEMP_VECTOR(1 To NROWS)

'-------------------------------------------------------------------------------
'     this subroutine finds those eigenvectors of a complex upper
'     hessenberg matrix corresponding to specified eigenvalues,
'     using inverse iteration.
'-------------------------------------------------------------------------------
      
      ERROR_VAL = 0
      q = 0
      hh = 1
      EE = NROWS
      ff = NROWS
'
      For k = 1 To NROWS
         If (Not EIGEN_VECTORS(k)) Then GoTo 1994  'next k
         If (hh > ff) Then GoTo 1995
         'set error -- underestimate of eigenvector
         If (q >= k) Then GoTo 1983
'     .......... check for possible splitting ..........
         For q = k To NROWS - 1
            If (REAL_MATRIX(q + 1, q) = 0# And _
                IMAG_MATRIX(q + 1, q) = 0#) Then Exit For 'GoTo 140
         Next q
'     .......... compute infinity TEMP_NORM of leading q by q
'                (hessenberg) matrix ..........
         TEMP_NORM = 0#
         p = 1
'
         For i = 1 To q
            XTEMP_VAL = 0#
            For j = p To q
                XTEMP_VAL = XTEMP_VAL + _
                    PYTHAG_FUNC(REAL_MATRIX(i, j), IMAG_MATRIX(i, j))
            Next j
'
            If (XTEMP_VAL > TEMP_NORM) Then TEMP_NORM = XTEMP_VAL
            p = i
         Next i
'     .......... tolerance replaces zero pivot in decomposition
'                and close roots are modified by tolerance ..........
         If (TEMP_NORM = 0#) Then TEMP_NORM = 1#
         tolerance = TEMP_NORM * 2 * 10 ^ -16
'     .......... TEMP_FACTOR is the criterion for growth ..........
         TEMP_ROOT = q
         TEMP_ROOT = Sqr(TEMP_ROOT)
         TEMP_FACTOR = 0.1 / TEMP_ROOT
1983:     REAL_LAMBDA = AREAL_VECTOR(k)
         IMAG_LAMBDA = AIMAG_VECTOR(k)
         If (k = 1) Then GoTo 1987
         o = k - 1
         GoTo 1985
'     .......... perturb eigenvalue if it is close to any previous eigenvalue ..........
1984:     REAL_LAMBDA = REAL_LAMBDA + tolerance
'     .......... for i=k-1 step -1 until 1 do -- ..........
1985:     For ii = 1 To o
            i = k - ii
            If (EIGEN_VECTORS(i) And Abs(AREAL_VECTOR(i) - _
                REAL_LAMBDA) < tolerance And Abs(AIMAG_VECTOR(i) - _
                    IMAG_LAMBDA) < tolerance) Then GoTo 1984
1986:     Next ii
'
         AREAL_VECTOR(k) = REAL_LAMBDA

1987:     p = 1
         For i = 1 To q
'
            For j = p To q
               DATA_MATRIX(i, j) = REAL_MATRIX(i, j)
               TEMP_MATRIX(i, j) = IMAG_MATRIX(i, j)
'            continue
            Next j
'
            DATA_MATRIX(i, i) = DATA_MATRIX(i, i) - REAL_LAMBDA
            TEMP_MATRIX(i, i) = TEMP_MATRIX(i, i) - IMAG_LAMBDA
            p = i
            ATEMP_VECTOR(i) = tolerance
         Next i
'     .......... triangular decomposition with interchanges,
'                replacing zero pivots by tolerance ..........
         If (q = 1) Then GoTo 1990
'
         For i = 2 To q
            p = i - 1
            If (PYTHAG_FUNC(DATA_MATRIX(i, p), _
                TEMP_MATRIX(i, p)) <= PYTHAG_FUNC(DATA_MATRIX(p, p), _
                    TEMP_MATRIX(p, p))) Then GoTo 1988
'
            For j = p To q
               YTEMP_VAL = DATA_MATRIX(i, j)
               DATA_MATRIX(i, j) = DATA_MATRIX(p, j)
               DATA_MATRIX(p, j) = YTEMP_VAL
               YTEMP_VAL = TEMP_MATRIX(i, j)
               TEMP_MATRIX(i, j) = TEMP_MATRIX(p, j)
               TEMP_MATRIX(p, j) = YTEMP_VAL
            Next j
'
1988:        If (DATA_MATRIX(p, p) = 0# And TEMP_MATRIX(p, p) = 0#) _
                Then DATA_MATRIX(p, p) = tolerance
            
            ATEMP_COMPLEX_OBJ.reel = DATA_MATRIX(i, p)
            ATEMP_COMPLEX_OBJ.imag = TEMP_MATRIX(i, p)
            BTEMP_COMPLEX_OBJ.reel = DATA_MATRIX(p, p)
            BTEMP_COMPLEX_OBJ.imag = TEMP_MATRIX(p, p)
            CTEMP_COMPLEX_OBJ = COMPLEX_QUOTIENT_OBJ_FUNC(ATEMP_COMPLEX_OBJ, BTEMP_COMPLEX_OBJ)
            XTEMP_VAL = CTEMP_COMPLEX_OBJ.reel
            YTEMP_VAL = CTEMP_COMPLEX_OBJ.imag
            
            
            If (XTEMP_VAL = 0# And YTEMP_VAL = 0#) Then GoTo 1989
'
            For j = p To q
               DATA_MATRIX(i, j) = DATA_MATRIX(i, j) - _
                        XTEMP_VAL * DATA_MATRIX(p, j) + _
                                YTEMP_VAL * TEMP_MATRIX(p, j)
               TEMP_MATRIX(i, j) = TEMP_MATRIX(i, j) - _
                        XTEMP_VAL * TEMP_MATRIX(p, j) - _
                                YTEMP_VAL * DATA_MATRIX(p, j)
            Next j
1989:     Next i
'
1990:     If (DATA_MATRIX(q, q) = 0# And TEMP_MATRIX(q, q) = 0#) _
                Then DATA_MATRIX(q, q) = tolerance
         kk = 0
'     .......... back substitution  -- ..........
         Do
             For ii = 1 To q
                i = q + 1 - ii
                XTEMP_VAL = ATEMP_VECTOR(i)
                YTEMP_VAL = 0#
                If (i <> q) Then
                    ss = i + 1
                    For j = ss To q
                       XTEMP_VAL = XTEMP_VAL - _
                            DATA_MATRIX(i, j) * ATEMP_VECTOR(j) + _
                                    TEMP_MATRIX(i, j) * BTEMP_VECTOR(j)
                       YTEMP_VAL = YTEMP_VAL - DATA_MATRIX(i, j) * _
                                BTEMP_VECTOR(j) - TEMP_MATRIX(i, j) * _
                                        ATEMP_VECTOR(j)
                    Next j
                End If
                
                ATEMP_COMPLEX_OBJ.reel = XTEMP_VAL
                ATEMP_COMPLEX_OBJ.imag = YTEMP_VAL
                BTEMP_COMPLEX_OBJ.reel = DATA_MATRIX(i, i)
                BTEMP_COMPLEX_OBJ.imag = TEMP_MATRIX(i, i)
                CTEMP_COMPLEX_OBJ = COMPLEX_QUOTIENT_OBJ_FUNC(ATEMP_COMPLEX_OBJ, BTEMP_COMPLEX_OBJ)
                ATEMP_VECTOR(i) = CTEMP_COMPLEX_OBJ.reel
                BTEMP_VECTOR(i) = CTEMP_COMPLEX_OBJ.imag
             
             Next ii
    '     .......... acceptance test for eigenvector and normalization ..........
             kk = kk + 1
             TEMP_NORM = 0#
             TEMP_THRESD = 0#
    '
             For i = 1 To q
                XTEMP_VAL = PYTHAG_FUNC(ATEMP_VECTOR(i), BTEMP_VECTOR(i))
                If Not (TEMP_THRESD >= XTEMP_VAL) Then
                    TEMP_THRESD = XTEMP_VAL
                    j = i
                End If
                TEMP_NORM = TEMP_NORM + XTEMP_VAL
             Next i
    '
             If (TEMP_NORM < TEMP_FACTOR) Then GoTo 1991
    '     .......... accept vector ..........
             XTEMP_VAL = ATEMP_VECTOR(j)
             YTEMP_VAL = BTEMP_VECTOR(j)
    '
             For i = 1 To q

                ATEMP_COMPLEX_OBJ.reel = ATEMP_VECTOR(i)
                ATEMP_COMPLEX_OBJ.imag = BTEMP_VECTOR(i)
                BTEMP_COMPLEX_OBJ.reel = XTEMP_VAL
                BTEMP_COMPLEX_OBJ.imag = YTEMP_VAL
                CTEMP_COMPLEX_OBJ = COMPLEX_QUOTIENT_OBJ_FUNC(ATEMP_COMPLEX_OBJ, BTEMP_COMPLEX_OBJ)
                BREAL_VECTOR(i, hh) = CTEMP_COMPLEX_OBJ.reel
                BIMAG_VECTOR(i, hh) = CTEMP_COMPLEX_OBJ.imag
             
             Next i
    '
             If (q = NROWS) Then GoTo 1993
             j = q + 1
             GoTo 1992
    '     .......... in-line procedure for choosing
    '                a new starting vector ..........
1991:         If (kk >= q) Then Exit Do
             XTEMP_VAL = TEMP_ROOT
             YTEMP_VAL = tolerance / (XTEMP_VAL + 1#)
             ATEMP_VECTOR(1) = tolerance
    '
             For i = 2 To q
                ATEMP_VECTOR(i) = YTEMP_VAL
             Next i
             j = q - kk + 1
             ATEMP_VECTOR(j) = ATEMP_VECTOR(j) - tolerance * XTEMP_VAL
         Loop  'back substitution
'     .......... set error -- unaccepted eigenvector ..........
         j = 1
         ERROR_VAL = -k
'     .......... set remaining vector components to zero ..........
1992:     For i = j To NROWS
            BREAL_VECTOR(i, hh) = 0#
            BIMAG_VECTOR(i, hh) = 0#
         Next i
'
1993:     hh = hh + 1
'      continue
1994:  Next k
'
      NSIZE = hh - 1

'     .......... set error -- underestimate of eigenvector
'                space required ..........
1995: If (ERROR_VAL <> 0) Then ERROR_VAL = ERROR_VAL - NROWS
      If (ERROR_VAL = 0) Then ERROR_VAL = -(2 * NROWS + 1)
      NSIZE = hh - 1
'---------------------back-transformation of the eigenvectors-----------------------
'-----------------------------------------------------------------------------------
'     this subroutine forms the eigenvectors of a complex general
'     matrix by back transforming those of the corresponding
'     upper hessenberg matrix determined by  corth.
'-----------------------------------------------------------------------------------
'
      If (NSIZE = 0) Then GoTo 1997
      pp = NROWS - 1
      qq = 1 + 1
      If (pp < qq) Then GoTo 1997
'     .......... for mp=igh-1 step -1 until low+1 do -- ..........
      For ff = qq To pp
         p = 1 + NROWS - ff
         If (REAL_MATRIX(p, p - 1) = 0# And IMAG_MATRIX(p, p - 1) = 0#) Then GoTo 1996
'     .......... H_VAL below is negative of H_VAL formed in corth ..........
         H_VAL = REAL_MATRIX(p, p - 1) * ORTHOG_REAL_VECTOR(p) + _
                IMAG_MATRIX(p, p - 1) * ORTHOG_IMAG_VECTOR(p)
         RR = p + 1
         For i = RR To NROWS
            ORTHOG_REAL_VECTOR(i) = REAL_MATRIX(i, p - 1)
            ORTHOG_IMAG_VECTOR(i) = IMAG_MATRIX(i, p - 1)
         Next i
         For j = 1 To NSIZE
            REAL_SCALE = 0#
            IMAG_SCALE = 0#
            For i = p To NROWS
               REAL_SCALE = REAL_SCALE + ORTHOG_REAL_VECTOR(i) * _
                    BREAL_VECTOR(i, j) + ORTHOG_IMAG_VECTOR(i) * BIMAG_VECTOR(i, j)
               IMAG_SCALE = IMAG_SCALE + ORTHOG_REAL_VECTOR(i) * _
                    BIMAG_VECTOR(i, j) - ORTHOG_IMAG_VECTOR(i) * BREAL_VECTOR(i, j)
            Next i
            REAL_SCALE = REAL_SCALE / H_VAL
            IMAG_SCALE = IMAG_SCALE / H_VAL
            For i = p To NROWS
               BREAL_VECTOR(i, j) = BREAL_VECTOR(i, j) + REAL_SCALE * _
                    ORTHOG_REAL_VECTOR(i) - IMAG_SCALE * ORTHOG_IMAG_VECTOR(i)
               BIMAG_VECTOR(i, j) = BIMAG_VECTOR(i, j) + REAL_SCALE * _
                    ORTHOG_IMAG_VECTOR(i) + IMAG_SCALE * ORTHOG_REAL_VECTOR(i)
            Next i
         Next j
1996:  Next ff
1997:

ReDim TEMP_MATRIX(1 To NROWS, 1 To 2 * NSIZE)
For i = 1 To NROWS
    For j = 1 To NSIZE
        TEMP_MATRIX(i, j) = BREAL_VECTOR(i, j)
        TEMP_MATRIX(i, j + NSIZE) = BIMAG_VECTOR(i, j)
    Next j
Next i


'--------------------Normalize Complex Matrix--------------------------

NROWS = UBound(TEMP_MATRIX, 1)
NSIZE = UBound(TEMP_MATRIX, 2) / 2


For j = 1 To 2 * NSIZE
    For i = 1 To NROWS
        If Abs(TEMP_MATRIX(i, j)) < (2 * LAMBDA) Then TEMP_MATRIX(i, j) = 0
    Next i
Next j

For j = 1 To NSIZE
    TEMP_SUM = 0 '
    For i = 1 To NROWS
        TEMP_VAL = TEMP_MATRIX(i, j) ^ 2 + _
                TEMP_MATRIX(i, j + NSIZE) ^ 2
                TEMP_SUM = TEMP_SUM + TEMP_VAL
    Next i
    TEMP_SUM = Sqr(TEMP_SUM)
    
    If Abs(TEMP_SUM) > (2 * LAMBDA) Then
            For i = 1 To NROWS
                TEMP_MATRIX(i, j) = TEMP_MATRIX(i, j) / TEMP_SUM
                TEMP_MATRIX(i, j + NSIZE) = TEMP_MATRIX(i, j + _
                    NSIZE) / TEMP_SUM
            Next i
    End If
Next j

'--------------------------------------------------------------------------------

DATA_MATRIX = MATRIX_TRIM_SMALL_VALUES_FUNC(TEMP_MATRIX, LAMBDA)  'clean-up
If CPLX_FORMAT = 2 Then DATA_MATRIX = _
            COMPLEX_MATRIX_FORMAT_FUNC(DATA_MATRIX, 12, CPLX_CHR_STR, epsilon)
If CPLX_FORMAT = 3 Then DATA_MATRIX = _
            COMPLEX_MATRIX_FORMAT_FUNC(DATA_MATRIX, 13, CPLX_CHR_STR, epsilon)

COMPLEX_MATRIX_INVERSE_EIGENVECTORS_FUNC = DATA_MATRIX

Exit Function
ERROR_LABEL:
COMPLEX_MATRIX_INVERSE_EIGENVECTORS_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : COMPLEX_COMPANION_MATRIX_QR_EIGENVALUES_FUNC

'DESCRIPTION   : Find real and complex eigenvalues of complex matrix with the
'iterative QR method. This function performs the diagonal reduction of a given
'complex matrix with QR method, and returns the approximate eigenvalues
'real or complex as an (n x 2) array. This function supports 3
'different formats: 1 = split, 2 = interlaced, 3 = string
'Optional parameter CPLX_FORMAT sets the complex input/output
'format (default = 1)

'REFERENCE: This function uses a reduction of the EISPACK FORTRAN
'COMQR and CORTH subroutines (April 1983) COMQR IS A TRANSLATION
'OF THE ALGOL PROCEDURE math. 12, 369-376(1968) by Martin  and Wilkinson.).

'LIBRARY       : CPLX
'GROUP         : EIGEN
'ID            : 003
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/08/2008
'************************************************************************************
'************************************************************************************

Function COMPLEX_COMPANION_MATRIX_QR_EIGENVALUES_FUNC(ByRef DATA_RNG As Variant, _
Optional ByVal CPLX_FORMAT As Integer = 1, _
Optional ByVal CPLX_CHR_STR As String = "i", _
Optional ByVal epsilon As Double = 2 * 10 ^ -14)

Dim i As Long
Dim j As Long
Dim l As Long
Dim D As Long

Dim p As Long '

Dim EE As Long
Dim ii As Long
Dim jj As Long
Dim kk As Long

Dim ll As Long '
Dim nn As Long

Dim pp As Long '
Dim qq As Long

Dim ss As Long '
Dim tt As Long

Dim NSIZE As Long
Dim nLOOPS As Long
Dim ERROR_VAL As Long

Dim F_VAL As Double
Dim G_VAL As Double
Dim H_VAL As Double

Dim NORM_VAL As Double
Dim ATEMP_VAL As Double
Dim BTEMP_VAL As Double

Dim IMAG_VAL As Double
Dim REAL_VAL As Double
Dim SCALE_VAL As Double

Dim STEMP_IMAG As Double
Dim STEMP_REEL As Double

Dim TTEMP_IMAG As Double
Dim TTEMP_REEL As Double

Dim XTEMP_IMAG As Double
Dim XTEMP_REEL As Double

Dim YTEMP_IMAG As Double
Dim YTEMP_REEL As Double

Dim ZTEMP_IMAG As Double
Dim ZTEMP_REEL As Double

Dim ATEMP_COMPLEX_OBJ As Cplx
Dim BTEMP_COMPLEX_OBJ As Cplx
Dim CTEMP_COMPLEX_OBJ As Cplx

Dim DTEMP_COMPLEX_OBJ As Cplx
Dim ETEMP_COMPLEX_OBJ As Cplx

Dim REAL_MATRIX As Variant
Dim IMAG_MATRIX As Variant

Dim REAL_VECTOR As Variant
Dim IMAG_VECTOR As Variant

Dim DATA_MATRIX As Variant
Dim TEMP_MATRIX As Variant

On Error GoTo ERROR_LABEL

DATA_MATRIX = DATA_RNG
If CPLX_FORMAT = 2 Then DATA_MATRIX = _
    COMPLEX_MATRIX_FORMAT_FUNC(DATA_MATRIX, 21, CPLX_CHR_STR, epsilon)

If CPLX_FORMAT = 3 Then DATA_MATRIX = _
    COMPLEX_MATRIX_FORMAT_FUNC(DATA_MATRIX, 31, CPLX_CHR_STR, epsilon)

'check dimension. Only square matrix here
If 2 * UBound(DATA_MATRIX, 1) <> UBound(DATA_MATRIX, 2) Then GoTo ERROR_LABEL
    
NSIZE = UBound(DATA_MATRIX, 1) 'matrix dimension
ReDim REAL_MATRIX(1 To NSIZE, 1 To NSIZE)
ReDim IMAG_MATRIX(1 To NSIZE, 1 To NSIZE)
For i = 1 To NSIZE
    For j = 1 To NSIZE
        REAL_MATRIX(i, j) = DATA_MATRIX(i, j)
        IMAG_MATRIX(i, j) = DATA_MATRIX(i, j + NSIZE)
    Next j
Next i

'------------------------hessember transformation-----------------------------------
      ReDim ORTHOG_REAL_VECTOR(1 To NSIZE)
      ReDim ORTHOG_IMAG_VECTOR(1 To NSIZE)
'--------------------------------------------------------------------------------
'     given a complex general matrix, this subroutine
'     reduces a submatrix situated in rows and columns
'     1 through NSIZE to upper hessenberg form by
'     unitary similarity transformations.
'--------------------------------------------------------------------------------
'
      pp = NSIZE - 1
      qq = 1 + 1
      EE = nn
      If (pp < qq) Then GoTo 1982
'
      For nn = qq To pp
         H_VAL = 0#
         ORTHOG_REAL_VECTOR(nn) = 0#
         ORTHOG_IMAG_VECTOR(nn) = 0#
         SCALE_VAL = 0#
'     .......... SCALE_VAL column (algol tol then not needed) ..........
         For i = nn To NSIZE
            SCALE_VAL = SCALE_VAL + Abs(REAL_MATRIX(i, nn - 1)) _
                + Abs(IMAG_MATRIX(i, nn - 1))
         Next i
'
         If (SCALE_VAL = 0#) Then Exit For
         p = nn + NSIZE
'     .......... for i=igh step -1 until nn do -- ..........
         For ii = nn To NSIZE
            i = p - ii
            ORTHOG_REAL_VECTOR(i) = REAL_MATRIX(i, nn - 1) / SCALE_VAL
            ORTHOG_IMAG_VECTOR(i) = IMAG_MATRIX(i, nn - 1) / SCALE_VAL
            H_VAL = H_VAL + ORTHOG_REAL_VECTOR(i) * ORTHOG_REAL_VECTOR(i) + _
                ORTHOG_IMAG_VECTOR(i) * ORTHOG_IMAG_VECTOR(i)
         Next ii
'
         G_VAL = Sqr(H_VAL)
         F_VAL = Sqr(ORTHOG_REAL_VECTOR(nn) ^ 2 + ORTHOG_IMAG_VECTOR(nn) ^ 2)
         If (F_VAL = 0#) Then 'GoTo 103
            ORTHOG_REAL_VECTOR(nn) = G_VAL
            REAL_MATRIX(nn, nn - 1) = SCALE_VAL
         Else
            H_VAL = H_VAL + F_VAL * G_VAL
            G_VAL = G_VAL / F_VAL
            ORTHOG_REAL_VECTOR(nn) = (1# + G_VAL) * ORTHOG_REAL_VECTOR(nn)
            ORTHOG_IMAG_VECTOR(nn) = (1# + G_VAL) * ORTHOG_IMAG_VECTOR(nn)
         End If
'     .......... form (i-(u*ut)/H_VAL) * a ..........
         For j = nn To NSIZE
            REAL_VAL = 0#
            IMAG_VAL = 0#
'     .......... for i=igh step -1 until nn do -- ..........
            For ii = nn To NSIZE
               i = p - ii
               REAL_VAL = REAL_VAL + ORTHOG_REAL_VECTOR(i) * _
                    REAL_MATRIX(i, j) + ORTHOG_IMAG_VECTOR(i) * IMAG_MATRIX(i, j)
               IMAG_VAL = IMAG_VAL + ORTHOG_REAL_VECTOR(i) * _
                    IMAG_MATRIX(i, j) - ORTHOG_IMAG_VECTOR(i) * REAL_MATRIX(i, j)
            Next ii
            REAL_VAL = REAL_VAL / H_VAL
            IMAG_VAL = IMAG_VAL / H_VAL
            For i = nn To NSIZE
               REAL_MATRIX(i, j) = REAL_MATRIX(i, j) - REAL_VAL * _
                    ORTHOG_REAL_VECTOR(i) + IMAG_VAL * ORTHOG_IMAG_VECTOR(i)
               IMAG_MATRIX(i, j) = IMAG_MATRIX(i, j) - REAL_VAL * _
                    ORTHOG_IMAG_VECTOR(i) - IMAG_VAL * ORTHOG_REAL_VECTOR(i)
            Next i
         Next j
'     .......... form (i-(u*ut)/H_VAL)*a*(i-(u*ut)/H_VAL) ..........
         For i = 1 To NSIZE
            REAL_VAL = 0#
            IMAG_VAL = 0#
'     .......... for j=igh step -1 until nn do -- ..........
            For jj = nn To NSIZE
               j = p - jj
               REAL_VAL = REAL_VAL + ORTHOG_REAL_VECTOR(j) * _
                    REAL_MATRIX(i, j) - ORTHOG_IMAG_VECTOR(j) * IMAG_MATRIX(i, j)
               IMAG_VAL = IMAG_VAL + ORTHOG_REAL_VECTOR(j) * _
                    IMAG_MATRIX(i, j) + ORTHOG_IMAG_VECTOR(j) * REAL_MATRIX(i, j)
'            continue
            Next jj
            REAL_VAL = REAL_VAL / H_VAL
            IMAG_VAL = IMAG_VAL / H_VAL
            For j = nn To NSIZE
               REAL_MATRIX(i, j) = REAL_MATRIX(i, j) - REAL_VAL * _
                    ORTHOG_REAL_VECTOR(j) - IMAG_VAL * ORTHOG_IMAG_VECTOR(j)
               IMAG_MATRIX(i, j) = IMAG_MATRIX(i, j) + REAL_VAL * _
                    ORTHOG_IMAG_VECTOR(j) - IMAG_VAL * ORTHOG_REAL_VECTOR(j)
            Next j
         Next i
'
         ORTHOG_REAL_VECTOR(nn) = SCALE_VAL * ORTHOG_REAL_VECTOR(nn)
         ORTHOG_IMAG_VECTOR(nn) = SCALE_VAL * ORTHOG_IMAG_VECTOR(nn)
         REAL_MATRIX(nn, nn - 1) = -G_VAL * REAL_MATRIX(nn, nn - 1)
         IMAG_MATRIX(nn, nn - 1) = -G_VAL * IMAG_MATRIX(nn, nn - 1)
      Next nn
1982:

'---------------------eigenvalues of hessemberg matrix------------------------------

      ReDim REAL_VECTOR(1 To NSIZE)
      ReDim IMAG_VECTOR(1 To NSIZE)

'-------------------------------------------------------------------------
'     this subroutine finds the eigenvalues of a complex
'     upper hessenberg matrix by the qr method.
'-------------------------------------------------------------------------
      ERROR_VAL = 0
      EE = NSIZE
      If (1 = NSIZE) Then GoTo 1984
'     .......... create real subdiagonal elements ..........
      l = 1 + 1
'
      For i = l To NSIZE
         ll = MINIMUM_FUNC(i + 1, NSIZE)
         If (IMAG_MATRIX(i, i - 1) = 0#) Then GoTo 1983
         NORM_VAL = PYTHAG_FUNC(REAL_MATRIX(i, i - 1), IMAG_MATRIX(i, i - 1))
         YTEMP_REEL = REAL_MATRIX(i, i - 1) / NORM_VAL
         YTEMP_IMAG = IMAG_MATRIX(i, i - 1) / NORM_VAL
         REAL_MATRIX(i, i - 1) = NORM_VAL
         IMAG_MATRIX(i, i - 1) = 0#
'
         For j = i To NSIZE
            STEMP_IMAG = YTEMP_REEL * IMAG_MATRIX(i, j) - _
                    YTEMP_IMAG * REAL_MATRIX(i, j)
            REAL_MATRIX(i, j) = YTEMP_REEL * REAL_MATRIX(i, j) + _
                    YTEMP_IMAG * IMAG_MATRIX(i, j)
            IMAG_MATRIX(i, j) = STEMP_IMAG
         Next j
'
         For j = 1 To ll
            STEMP_IMAG = YTEMP_REEL * IMAG_MATRIX(j, i) + _
                    YTEMP_IMAG * REAL_MATRIX(j, i)
            REAL_MATRIX(j, i) = YTEMP_REEL * REAL_MATRIX(j, i) - _
                    YTEMP_IMAG * IMAG_MATRIX(j, i)
            IMAG_MATRIX(j, i) = STEMP_IMAG
         Next j
1983:   Next i


'     .......... store roots isolated by cbal ..........
1984:  For i = 1 To NSIZE
         If (i >= 1 And i <= NSIZE) Then GoTo 1985
         REAL_VECTOR(i) = REAL_MATRIX(i, i)
         IMAG_VECTOR(i) = IMAG_MATRIX(i, i)
1985:  Next i
'
      D = NSIZE
      TTEMP_REEL = 0#
      TTEMP_IMAG = 0#
      nLOOPS = 30 * NSIZE
'     .......... search for next eigenvalue ..........
Do
1986:  If (D < 1) Then GoTo 1999
      kk = 0
      tt = D - 1
'     .......... look for single small sub-diagonal element
'                for l=en step -1 until 1 d0 -- ..........
    Do
1987:  For ll = 1 To D
         l = D + 1 - ll
         If (l = 1) Then GoTo 1989
         ATEMP_VAL = Abs(REAL_MATRIX(l - 1, l - 1)) + _
                Abs(IMAG_MATRIX(l - 1, l - 1)) + _
                        Abs(REAL_MATRIX(l, l)) + Abs(IMAG_MATRIX(l, l))
         BTEMP_VAL = ATEMP_VAL + Abs(REAL_MATRIX(l, l - 1))
         If (BTEMP_VAL = ATEMP_VAL) Then GoTo 1989
1988:  Next ll
'     .......... form shift ..........
1989:  If (l = D) Then GoTo 1997
      If (nLOOPS = 0) Then GoTo 1998
      If (kk = 10 Or kk = 20) Then GoTo 1991
      STEMP_REEL = REAL_MATRIX(D, D)
      STEMP_IMAG = IMAG_MATRIX(D, D)
      XTEMP_REEL = REAL_MATRIX(tt, D) * REAL_MATRIX(D, tt)
      XTEMP_IMAG = IMAG_MATRIX(tt, D) * REAL_MATRIX(D, tt)
      If (XTEMP_REEL = 0# And XTEMP_IMAG = 0#) Then GoTo 1992
      YTEMP_REEL = (REAL_MATRIX(tt, tt) - STEMP_REEL) / 2#
      YTEMP_IMAG = (IMAG_MATRIX(tt, tt) - STEMP_IMAG) / 2#
        
        DTEMP_COMPLEX_OBJ.reel = YTEMP_REEL ^ 2 - YTEMP_IMAG ^ 2 + XTEMP_REEL
        DTEMP_COMPLEX_OBJ.imag = 2# * YTEMP_REEL * YTEMP_IMAG + XTEMP_IMAG
        
        ETEMP_COMPLEX_OBJ = COMPLEX_ROOT_OBJ_FUNC(DTEMP_COMPLEX_OBJ, 2)
        ZTEMP_REEL = ETEMP_COMPLEX_OBJ.reel
        ZTEMP_IMAG = ETEMP_COMPLEX_OBJ.imag
      
      
      
      If (YTEMP_REEL * ZTEMP_REEL + YTEMP_IMAG * _
            ZTEMP_IMAG >= 0#) Then GoTo 1990
      ZTEMP_REEL = -ZTEMP_REEL
      ZTEMP_IMAG = -ZTEMP_IMAG
1990:
            ATEMP_COMPLEX_OBJ.reel = XTEMP_REEL
            ATEMP_COMPLEX_OBJ.imag = XTEMP_IMAG
            BTEMP_COMPLEX_OBJ.reel = YTEMP_REEL + ZTEMP_REEL
            BTEMP_COMPLEX_OBJ.imag = YTEMP_IMAG + ZTEMP_IMAG
            
            CTEMP_COMPLEX_OBJ = COMPLEX_QUOTIENT_OBJ_FUNC(ATEMP_COMPLEX_OBJ, BTEMP_COMPLEX_OBJ)
            XTEMP_REEL = CTEMP_COMPLEX_OBJ.reel
            XTEMP_IMAG = CTEMP_COMPLEX_OBJ.imag
      
      
      STEMP_REEL = STEMP_REEL - XTEMP_REEL
      STEMP_IMAG = STEMP_IMAG - XTEMP_IMAG
      GoTo 1992
'     .......... form exceptional shift ..........
1991:  STEMP_REEL = Abs(REAL_MATRIX(D, tt)) + Abs(REAL_MATRIX(tt, D - 2))
      STEMP_IMAG = 0#
'
1992:  For i = 1 To D
         REAL_MATRIX(i, i) = REAL_MATRIX(i, i) - STEMP_REEL
         IMAG_MATRIX(i, i) = IMAG_MATRIX(i, i) - STEMP_IMAG
1993:  Next i
'
      TTEMP_REEL = TTEMP_REEL + STEMP_REEL
      TTEMP_IMAG = TTEMP_IMAG + STEMP_IMAG
      kk = kk + 1
      nLOOPS = nLOOPS - 1
'     .......... reduce to triangle (rows) ..........
      ss = l + 1
'
      For i = ss To D
         STEMP_REEL = REAL_MATRIX(i, i - 1)
         REAL_MATRIX(i, i - 1) = 0#
         NORM_VAL = PYTHAG_FUNC(PYTHAG_FUNC(REAL_MATRIX(i - 1, i - 1), _
                IMAG_MATRIX(i - 1, i - 1)), STEMP_REEL)
         XTEMP_REEL = REAL_MATRIX(i - 1, i - 1) / NORM_VAL
         REAL_VECTOR(i - 1) = XTEMP_REEL
         XTEMP_IMAG = IMAG_MATRIX(i - 1, i - 1) / NORM_VAL
         IMAG_VECTOR(i - 1) = XTEMP_IMAG
         REAL_MATRIX(i - 1, i - 1) = NORM_VAL
         IMAG_MATRIX(i - 1, i - 1) = 0#
         IMAG_MATRIX(i, i - 1) = STEMP_REEL / NORM_VAL
'
         For j = i To D
            YTEMP_REEL = REAL_MATRIX(i - 1, j)
            YTEMP_IMAG = IMAG_MATRIX(i - 1, j)
            ZTEMP_REEL = REAL_MATRIX(i, j)
            ZTEMP_IMAG = IMAG_MATRIX(i, j)
            REAL_MATRIX(i - 1, j) = XTEMP_REEL * _
                    YTEMP_REEL + XTEMP_IMAG * YTEMP_IMAG + _
                            IMAG_MATRIX(i, i - 1) * ZTEMP_REEL
            IMAG_MATRIX(i - 1, j) = XTEMP_REEL * YTEMP_IMAG - _
                    XTEMP_IMAG * YTEMP_REEL + _
                            IMAG_MATRIX(i, i - 1) * ZTEMP_IMAG
            REAL_MATRIX(i, j) = XTEMP_REEL * ZTEMP_REEL - _
                        XTEMP_IMAG * ZTEMP_IMAG - _
                                IMAG_MATRIX(i, i - 1) * YTEMP_REEL
            IMAG_MATRIX(i, j) = XTEMP_REEL * ZTEMP_IMAG + _
                            XTEMP_IMAG * ZTEMP_REEL - _
                                    IMAG_MATRIX(i, i - 1) * YTEMP_IMAG
         Next j
      Next i
'
      STEMP_IMAG = IMAG_MATRIX(D, D)
      If (STEMP_IMAG = 0#) Then GoTo 1994
      NORM_VAL = PYTHAG_FUNC(REAL_MATRIX(D, D), STEMP_IMAG)
      STEMP_REEL = REAL_MATRIX(D, D) / NORM_VAL
      STEMP_IMAG = STEMP_IMAG / NORM_VAL
      REAL_MATRIX(D, D) = NORM_VAL
      IMAG_MATRIX(D, D) = 0#
'     .......... inverse operation (columns) ..........
1994:  For j = ss To D
         XTEMP_REEL = REAL_VECTOR(j - 1)
         XTEMP_IMAG = IMAG_VECTOR(j - 1)
'
         For i = l To j
            YTEMP_REEL = REAL_MATRIX(i, j - 1)
            YTEMP_IMAG = 0#
            ZTEMP_REEL = REAL_MATRIX(i, j)
            ZTEMP_IMAG = IMAG_MATRIX(i, j)
            If (i = j) Then GoTo 1995
            YTEMP_IMAG = IMAG_MATRIX(i, j - 1)
            IMAG_MATRIX(i, j - 1) = XTEMP_REEL * YTEMP_IMAG + _
                        XTEMP_IMAG * YTEMP_REEL + _
                                IMAG_MATRIX(j, j - 1) * ZTEMP_IMAG
1995:        REAL_MATRIX(i, j - 1) = XTEMP_REEL * YTEMP_REEL - _
                        XTEMP_IMAG * YTEMP_IMAG + _
                                IMAG_MATRIX(j, j - 1) * ZTEMP_REEL
            REAL_MATRIX(i, j) = XTEMP_REEL * ZTEMP_REEL + _
                        XTEMP_IMAG * ZTEMP_IMAG - _
                                IMAG_MATRIX(j, j - 1) * YTEMP_REEL
            IMAG_MATRIX(i, j) = XTEMP_REEL * ZTEMP_IMAG - _
                        XTEMP_IMAG * ZTEMP_REEL - _
                                IMAG_MATRIX(j, j - 1) * YTEMP_IMAG
         Next i
1996:  Next j
'
      If (STEMP_IMAG <> 0#) Then
        For i = l To D
           YTEMP_REEL = REAL_MATRIX(i, D)
           YTEMP_IMAG = IMAG_MATRIX(i, D)
           REAL_MATRIX(i, D) = STEMP_REEL * YTEMP_REEL - STEMP_IMAG * YTEMP_IMAG
           IMAG_MATRIX(i, D) = STEMP_REEL * YTEMP_IMAG + STEMP_IMAG * YTEMP_REEL
        Next i
      End If
'
    Loop 'GoTo 1987: look for single small sub-diagonal element
'     .......... a root found ..........
1997:  REAL_VECTOR(D) = REAL_MATRIX(D, D) + TTEMP_REEL
      IMAG_VECTOR(D) = IMAG_MATRIX(D, D) + TTEMP_IMAG
      D = tt
'
Loop 'GoTo 1986 : search for next eigenvalue
'     .......... set error -- all eigenvalues have not
'                converged after 30*NSIZE iterations ..........
1998:
      ERROR_VAL = D
1999:

ReDim TEMP_MATRIX(1 To NSIZE, 1 To 2) 'data output
For i = 1 To NSIZE
    If i > ERROR_VAL Then
        TEMP_MATRIX(i, 1) = REAL_VECTOR(i)
        TEMP_MATRIX(i, 2) = IMAG_VECTOR(i)
    Else
        TEMP_MATRIX(i, 1) = "-"
        TEMP_MATRIX(i, 2) = "-"
    End If
Next
TEMP_MATRIX = MATRIX_TRIM_SMALL_VALUES_FUNC(TEMP_MATRIX, epsilon)
TEMP_MATRIX = MATRIX_QUICK_SORT_FUNC(TEMP_MATRIX, 1, 1)

COMPLEX_COMPANION_MATRIX_QR_EIGENVALUES_FUNC = TEMP_MATRIX

Exit Function
ERROR_LABEL:
COMPLEX_COMPANION_MATRIX_QR_EIGENVALUES_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : COMPLEX_MATRIX_EIGENVECTORS_FUNC

'DESCRIPTION   : This function returns the complex eigenvector associated
'with the given complex eigenvalue of a real or complex matrix A (n x n)

'LIBRARY       : NUMBER_COMPLEX
'GROUP         : EIGEN
'ID            : 004
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/08/2008
'************************************************************************************
'************************************************************************************

Function COMPLEX_MATRIX_EIGENVECTORS_FUNC(ByRef DATA_RNG As Variant, _
ByRef CPLX_EIGEN_RNG As Variant, _
Optional ByVal CPLX_FORMAT As Integer = 1, _
Optional ByVal CPLX_CHR_STR As String = "i", _
Optional ByVal epsilon As Double = 0.0000000001)

'The function returns an array of two columns (n x 2): the first
'column contains the real part, the second column the imaginary part.
'It returns an array of four columns (n x 4) if the eigenvalue is
'double. The first two columns contain the first complex eigenvector;
'the last two columns contain the second complex eigenvector. And so on.

'The optional parameter epsilon (default 1E-10) is useful only if your
'eigenvalue has an error. In that case the MaxErr should be proportionally
'increased (1E-8, 1E-6, etc.). Otherwise the result may be a NULL matrix.

Dim i As Long
Dim j As Long
Dim k As Long

Dim NROWS As Long

Dim REAL_VAL As Double
Dim IMAG_VAL As Double

Dim DATA_MATRIX As Variant

Dim ATEMP_MATRIX As Variant
Dim CTEMP_MATRIX As Variant
Dim BTEMP_MATRIX As Variant

Dim EIGEN_MATRIX As Variant

On Error GoTo ERROR_LABEL

DATA_MATRIX = DATA_RNG
EIGEN_MATRIX = COMPLEX_EXTRACT_NUMBER_FUNC(CPLX_EIGEN_RNG)

'---------------------------------------------------------------------------------
If UBound(DATA_MATRIX, 1) = UBound(DATA_MATRIX, 2) Then 'Real Matrix
'---------------------------------------------------------------------------------

    If CPLX_FORMAT = 3 Or CPLX_FORMAT = 2 Then: GoTo 1982
    
    NROWS = UBound(DATA_MATRIX, 1)
    ReDim CTEMP_MATRIX(1 To NROWS, 1 To 2 * NROWS)
    ReDim ATEMP_MATRIX(1 To 2 * NROWS, 1 To 2 * NROWS)
    REAL_VAL = EIGEN_MATRIX(1, 1)
    IMAG_VAL = EIGEN_MATRIX(2, 1)
    'load matrix for complex system
    'real matrix
    For i = 1 To NROWS
        For j = 1 To NROWS
            ATEMP_MATRIX(i, j) = DATA_MATRIX(i, j)
            ATEMP_MATRIX(i + NROWS, j + NROWS) = ATEMP_MATRIX(i, j)
        Next j
        ATEMP_MATRIX(i, i) = ATEMP_MATRIX(i, i) - REAL_VAL
        ATEMP_MATRIX(i + NROWS, i + NROWS) = ATEMP_MATRIX(i, i)
        ATEMP_MATRIX(i, i + NROWS) = IMAG_VAL
        ATEMP_MATRIX(i + NROWS, i) = -IMAG_VAL
    Next i

    BTEMP_MATRIX = MATRIX_GS_SINGULAR_LINEAR_SYSTEM_FUNC(ATEMP_MATRIX, , 0, epsilon, 0)

    k = 1
    If IMAG_VAL = 0 Then   'only real eigenvector
        For j = 1 To NROWS
            If BTEMP_MATRIX(j, j) = 1 Then
                For i = 1 To NROWS
                    CTEMP_MATRIX(i, k) = BTEMP_MATRIX(i, j)
                Next i
                k = k + 2
            End If
        Next j
    Else          'complex eigenvector
        For j = 1 To NROWS Step 2
            If BTEMP_MATRIX(j + NROWS, j + NROWS) = 1 Then
                For i = 1 To NROWS
                    CTEMP_MATRIX(i, k) = BTEMP_MATRIX(i + NROWS, j + NROWS)
                    CTEMP_MATRIX(i, k + 1) = -BTEMP_MATRIX(i, j + NROWS)
                Next i
                k = k + 2
            End If
        Next j
    End If
'---------------------------------------------------------------------------------
Else 'Complex Matrix
'---------------------------------------------------------------------------------
1982:
    If CPLX_FORMAT = 2 Then DATA_MATRIX = _
            COMPLEX_MATRIX_FORMAT_FUNC(DATA_MATRIX, 21, CPLX_CHR_STR, epsilon)
    If CPLX_FORMAT = 3 Then DATA_MATRIX = _
            COMPLEX_MATRIX_FORMAT_FUNC(DATA_MATRIX, 31, CPLX_CHR_STR, epsilon)

    NROWS = UBound(DATA_MATRIX, 1)
    ReDim CTEMP_MATRIX(1 To NROWS, 1 To 2 * NROWS)
    ReDim ATEMP_MATRIX(1 To 2 * NROWS, 1 To 2 * NROWS)
    REAL_VAL = EIGEN_MATRIX(1, 1)
    IMAG_VAL = EIGEN_MATRIX(2, 1)
    'load matrix for complex system
    'complex matrix
    For i = 1 To NROWS
        For j = 1 To NROWS
            'real part
            ATEMP_MATRIX(i, j) = DATA_MATRIX(i, j)
            ATEMP_MATRIX(i + NROWS, j + NROWS) = DATA_MATRIX(i, j)
            'complex part
            ATEMP_MATRIX(i, j + NROWS) = -DATA_MATRIX(i, j + NROWS)
            ATEMP_MATRIX(i + NROWS, j) = DATA_MATRIX(i, j + NROWS)
        Next j
        ATEMP_MATRIX(i, i) = ATEMP_MATRIX(i, i) - REAL_VAL
        ATEMP_MATRIX(i + NROWS, i + NROWS) = ATEMP_MATRIX(i, i)
        ATEMP_MATRIX(i, i + NROWS) = ATEMP_MATRIX(i, i + NROWS) + IMAG_VAL
        ATEMP_MATRIX(i + NROWS, i) = ATEMP_MATRIX(i + NROWS, i) - IMAG_VAL
    Next i

    BTEMP_MATRIX = MATRIX_GS_SINGULAR_LINEAR_SYSTEM_FUNC(ATEMP_MATRIX, , 0, epsilon, 0)
    k = 1
    For j = 1 To 2 * NROWS 'Step 2
        If BTEMP_MATRIX(j, j) = 1 Then
            If j > NROWS Then
                For i = 1 To NROWS
                    CTEMP_MATRIX(i, k) = BTEMP_MATRIX(i + NROWS, j)
                    CTEMP_MATRIX(i, k + 1) = -BTEMP_MATRIX(i, j)
                Next i
            Else
                For i = 1 To NROWS
                    CTEMP_MATRIX(i, k) = BTEMP_MATRIX(i, j)
                    CTEMP_MATRIX(i, k + 1) = BTEMP_MATRIX(i + NROWS, j)
                Next i
            End If
            k = k + 2
            If k > NROWS Then GoTo 1983
        End If
    Next j

'---------------------------------------------------------------------------------
End If
'---------------------------------------------------------------------------------

1983:

CTEMP_MATRIX = MATRIX_TRIM_SMALL_VALUES_FUNC(CTEMP_MATRIX, epsilon)
If CPLX_FORMAT = 2 Then CTEMP_MATRIX = _
        COMPLEX_MATRIX_FORMAT_FUNC(CTEMP_MATRIX, 12, CPLX_CHR_STR, epsilon)
'If CPLX_FORMAT = 2 Then CTEMP_MATRIX = _
'        COMPLEX_MATRIX_FORMAT_FUNC(CTEMP_MATRIX, 12, CPLX_CHR_STR, epsilon)

If CPLX_FORMAT = 3 Then CTEMP_MATRIX = _
        COMPLEX_MATRIX_FORMAT_FUNC(CTEMP_MATRIX, 13, CPLX_CHR_STR, epsilon)
'If CPLX_FORMAT = 3 Then CTEMP_MATRIX = _
'        COMPLEX_MATRIX_FORMAT_FUNC(CTEMP_MATRIX, 13, CPLX_CHR_STR, epsilon)

COMPLEX_MATRIX_EIGENVECTORS_FUNC = CTEMP_MATRIX
Exit Function
ERROR_LABEL:
COMPLEX_MATRIX_EIGENVECTORS_FUNC = Err.number
End Function
