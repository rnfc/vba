Attribute VB_Name = "OPTIM_ODE_GALERKIN_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.

'-----------------------------------------------------------------------------------------------------
Private PUB_PARAM_ARR As Variant
Private PUB_INDEX_VAL As Single
Private PUB_BASIS_POINTS As Single
Private PUB_START_POINT As Double
Private PUB_END_POINT As Double
Private PUB_FUNC_NAME_STR As String


Private FunctionName_ As String
'-----------------------------------------------------------------------------------------------------

'************************************************************************************
'************************************************************************************
'FUNCTION      : GALERKIN_ODE_SOLVER_FUNC

'DESCRIPTION   : Galerkin to solve ODE. Set of galerkin equations are solved
'using NLE_FSOLVE_FUNC. Set of equations are formed by using Galerkin method.

'After calculating coefficients, y[x] is plotted against x for the range
'under consideration.

'Based on paper:
'http://math.fullerton.edu/mathews/n2003/galerkin/GalerkinMod/Links/GalerkinMod_lnk_5.html

'LIBRARY       : OPTIMIZATION
'GROUP         : ODE_GALERKIN
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/08/2008
'************************************************************************************
'************************************************************************************

Function GALERKIN_ODE_SOLVER_FUNC(ByVal FUNC_NAME_STR As String, _
ByVal START_POINT As Double, _
ByVal END_POINT As Double, _
Optional ByVal BASIS_POINTS As Long = 7, _
Optional ByVal NO_POINTS As Long = 10, _
Optional ByVal GUESS_RNG As Variant = 0, _
Optional ByVal OUTPUT As Integer = 1)
  
'Start Point(a)
'End point (b)
'No. of basis functions
'No. of points for plot
  
Dim i As Long

Dim X_VAL As Double
Dim Y_VAL As Double
Dim DELTA_POINTS As Double

Dim GUESS_ARR As Variant
Dim GUESS_VECTOR As Variant

Dim TEMP_ARR As Variant
Dim TEMP1_MATRIX As Variant
Dim TEMP2_MATRIX As Variant

On Error GoTo ERROR_LABEL

PUB_BASIS_POINTS = BASIS_POINTS
PUB_START_POINT = START_POINT
PUB_END_POINT = END_POINT
PUB_FUNC_NAME_STR = FUNC_NAME_STR

'-----------------------------------------------------------------------------
If IsArray(GUESS_RNG) = True Then
    GUESS_VECTOR = GUESS_RNG
    If UBound(GUESS_VECTOR, 1) = 1 Then
        GUESS_VECTOR = MATRIX_TRANSPOSE_FUNC(GUESS_VECTOR)
    End If
Else
    ReDim GUESS_VECTOR(1 To BASIS_POINTS, 1 To 1)
    For i = 1 To BASIS_POINTS
        GUESS_VECTOR(i, 1) = GUESS_RNG
    Next i
End If
If UBound(GUESS_VECTOR, 1) <> BASIS_POINTS Then: GoTo ERROR_LABEL
'-----------------------------------------------------------------------------

ReDim GUESS_ARR(1 To BASIS_POINTS)
For i = 1 To BASIS_POINTS
    GUESS_ARR(i) = GUESS_VECTOR(i, 1)
Next i

TEMP_ARR = NLE_FSOLVE_FUNC("GALERKIN_INTEGRAND_FUNC", GUESS_ARR)

ReDim TEMP1_MATRIX(1 To BASIS_POINTS, 1 To 2)

For i = 1 To BASIS_POINTS
    TEMP1_MATRIX(i, 1) = i
    TEMP1_MATRIX(i, 2) = TEMP_ARR(i)
Next i

If OUTPUT = 0 Then 'Coefficients for basis functions
' No/value
    GALERKIN_ODE_SOLVER_FUNC = TEMP1_MATRIX
    Exit Function
End If

DELTA_POINTS = (END_POINT - START_POINT) / NO_POINTS
X_VAL = START_POINT

ReDim TEMP2_MATRIX(1 To NO_POINTS + 1, 1 To 2)

For i = 1 To NO_POINTS + 1
    Y_VAL = GALERKIN_STORES_FUNC(X_VAL, TEMP_ARR)
    TEMP2_MATRIX(i, 1) = X_VAL
    TEMP2_MATRIX(i, 2) = Y_VAL
    X_VAL = X_VAL + DELTA_POINTS
Next i

If OUTPUT = 1 Then 'values for plot of y[x]
'x/y[x]
    GALERKIN_ODE_SOLVER_FUNC = TEMP2_MATRIX
Else
    GALERKIN_ODE_SOLVER_FUNC = Array(TEMP1_MATRIX, TEMP2_MATRIX)
End If

Exit Function
ERROR_LABEL:
GALERKIN_ODE_SOLVER_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : GALERKIN_STORES_FUNC
'DESCRIPTION   :
'LIBRARY       : OPTIMIZATION
'GROUP         : ODE_GALERKIN
'ID            : 002
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/08/2008
'************************************************************************************
'************************************************************************************

Private Function GALERKIN_STORES_FUNC(ByVal X_VAL As Double, _
ByRef PARAM_ARR As Variant)
  
Dim i As Long
Dim YX_VAL As Double 'stores y[X_VAL]
Dim Y0_VAL As Double 'initial value

On Error GoTo ERROR_LABEL
  
Y0_VAL = 1: YX_VAL = 1
'substitute y[X_VAL] with basis functions
For i = 1 To PUB_BASIS_POINTS
    YX_VAL = YX_VAL + PARAM_ARR(i) * X_VAL ^ i
Next i
GALERKIN_STORES_FUNC = YX_VAL

Exit Function
ERROR_LABEL:
GALERKIN_STORES_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : GALERKIN_INTEGRAND_FUNC
'DESCRIPTION   :
'LIBRARY       : OPTIMIZATION
'GROUP         : ODE_GALERKIN
'ID            : 003
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/08/2008
'************************************************************************************
'************************************************************************************
Private Function GALERKIN_INTEGRAND_FUNC(ByRef PARAM_ARR As Variant)
  
Dim i As Long
Dim nLOOPS As Long
Dim TEMP_ARR As Variant
Dim tolerance As Double
    
On Error GoTo ERROR_LABEL
  
tolerance = 0.000001 'tolerance for integral
nLOOPS = 1000
  
ReDim TEMP_ARR(1 To PUB_BASIS_POINTS)
  
For i = 1 To PUB_BASIS_POINTS
    PUB_INDEX_VAL = i
    PUB_PARAM_ARR = PARAM_ARR
    TEMP_ARR(i) = GAULEG7_INTEGRATION_FUNC(PUB_FUNC_NAME_STR, PUB_START_POINT, PUB_END_POINT, tolerance, nLOOPS)
Next i

GALERKIN_INTEGRAND_FUNC = TEMP_ARR

Exit Function
ERROR_LABEL:
GALERKIN_INTEGRAND_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : GALERKIN_ODE_TEST_FUNC1

'DESCRIPTION   : This routine demonstrates how to solve ODE by using Galerkin method.
'-1 -x + sin(2*x) + 5*y[x] + y'[x]    with initial condition y[0]=1
'Basis function substitution is :
'y[x] = y0 + c1*x + c2*x^2 + c3*x3 and so on

'LIBRARY       : OPTIMIZATION
'GROUP         : ODE_GALERKIN
'ID            : 004
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/08/2008
'************************************************************************************
'************************************************************************************

Function GALERKIN_ODE_TEST_FUNC1(ByVal X_VAL As Double) 'As Double

Dim i As Single
Dim RX_VAL As Double
Dim YX_VAL As Double 'stores y[X_VAL]
Dim Y0_VAL As Double 'iniatial value
Dim YDX_VAL As Double 'stores first derivative of y[X_VAL]

On Error GoTo ERROR_LABEL

Y0_VAL = 1
YX_VAL = 1
'substitute y[X_VAL] with basis functions
For i = 1 To PUB_BASIS_POINTS
  YX_VAL = YX_VAL + PUB_PARAM_ARR(i) * X_VAL ^ i
Next i
YDX_VAL = 0
For i = 1 To PUB_BASIS_POINTS
  YDX_VAL = YDX_VAL + i * PUB_PARAM_ARR(i) * (X_VAL ^ (i - 1))
Next i
RX_VAL = -1 - X_VAL + Sin(2 * X_VAL) + 5 * YX_VAL + YDX_VAL 'residual

'RX_VAL = -1 - X_VAL + Sin(2 * X_VAL) + PUB_PARAM_ARR(1) + 2 * X_VAL * PUB_PARAM_ARR(2) + _
3 * X_VAL ^ 2 * PUB_PARAM_ARR(3) + 4 * X_VAL ^ 3 * PUB_PARAM_ARR(4) + 5 * X_VAL ^ 4 * PUB_PARAM_ARR(5) + _
6 * X_VAL ^ 5 * PUB_PARAM_ARR(6) + 7 * X_VAL ^ 6 * PUB_PARAM_ARR(7) + 5 * (1 + X_VAL * PUB_PARAM_ARR(1) + _
X_VAL ^ 2 * PUB_PARAM_ARR(2) + X_VAL ^ 3 * PUB_PARAM_ARR(3) + X_VAL ^ 4 * PUB_PARAM_ARR(4) + _
X_VAL ^ 5 * PUB_PARAM_ARR(5) + X_VAL ^ 6 * PUB_PARAM_ARR(6) + X_VAL ^ 7 * PUB_PARAM_ARR(7))

GALERKIN_ODE_TEST_FUNC1 = (X_VAL ^ PUB_INDEX_VAL) * RX_VAL

Exit Function
ERROR_LABEL:
GALERKIN_ODE_TEST_FUNC1 = Err.number
End Function


Function NLE_FSOLVE_FUNC(FunctionName As String, _
xguessvec As Variant) As Variant

FunctionName_ = FunctionName
Dim j As Single
Dim n As Single
On Error GoTo err_sub
n = UBound(xguessvec)
GoTo no_error
err_sub:
n = 1
tmp = xguessvec
ReDim xguessvec(1 To 1)
xguessvec(1) = tmp
no_error:
Dim MAXFEV As Single
Dim ML As Single
Dim mu As Single
Dim MODE As Single
Dim NPRINT As Single
Dim Info As Single
Dim NFEV As Single
Dim LDFJAC As Single
Dim LR As Single
Dim NWRITE As Single
Dim XTOL As Double
Dim EPSFCN As Double
Dim FACTOR As Double

ReDim x(1 To n)
ReDim fvec(1 To n)
ReDim diag(1 To n)
ReDim fjac(1 To n, 1 To n)
LR = (n * (n + 1)) / 2
ReDim r(1 To LR)
ReDim QTF(1 To n)
ReDim WA1(1 To n)
ReDim WA2(1 To n)
ReDim WA3(1 To n)
ReDim WA4(1 To n)


NWRITE = 6
'N = 9
 '
 '     THE FOLLOWING STARTING VALUES PROVIDE A ROUGH SOLUTION.
 '
       For j = 1 To n
          x(j) = xguessvec(j)
       Next j
 '
       LDFJAC = n
       
 '
 '     SET XTOL TO THE SQUARE ROOT OF THE MACHINE PRECISION.
 '     UNLESS HIGH PRECISION SOLUTIONS ARE REQUIRED,
 '     THIS IS THE RECOMMENDED SETTING.
 '
       'XTOL = 0.0001
       XTOL = 0.000000000000001
MAXFEV = 2000
       ML = 1
       'MU = 1
       mu = n - 1
       EPSFCN = 0
       MODE = 2
       For j = 1 To n
          diag(j) = 1
       Next j
       FACTOR = 100
       NPRINT = 0
Dim fcn2 As String
fcn2 = "dummy"
 hybrd fcn2, n, x, fvec, XTOL, MAXFEV, ML, mu, EPSFCN, diag, _
                 MODE, FACTOR, NPRINT, Info, NFEV, fjac, LDFJAC, _
                 r, LR, QTF, WA1, WA2, WA3, WA4
      
      NLE_FSOLVE_FUNC = x
'      iflag = 1
 '     fcn N, x, fvec, iflag
  '    FNORM = myenorm(N, fvec)

End Function


Private Sub fcn(ByRef n As Single, ByRef x As Variant, ByRef fvec As Variant, _
                            ByRef iflag As Variant)
    If (iflag <> 0) Then GoTo L5
           Exit Sub
L5:
    If n = 1 Then
        fvec(1) = Application.Run(FunctionName_, x(1))
    Else
        fvec = Application.Run(FunctionName_, x)
    End If
End Sub
Private Sub fcntest(ByRef n As Single, ByRef x As Variant, ByRef fvec As Variant, _
ByRef iflag As Variant)
'         calculate the functions at x and
'         return this vector in fvec.
Dim k As Single
Dim temp As Double
Dim temp1 As Double

'       DOUBLE PRECISION ONE,TEMP,TEMP1,TEMP2,THREE,TWO,ZERO
'       DATA ZERO,ONE,TWO,THREE /0.D0,1.D0,2.D0,3.D0/
'
       If (iflag <> 0) Then GoTo L5
 '
 '     INSERT PRINT STATEMENTS HERE WHEN NPRINT IS POSITIVE.
 '
       Return
L5:
       For k = 1 To n
          temp = (3 - 2 * x(k)) * x(k)
          temp1 = 0
          If (k <> 1) Then temp1 = x(k - 1)
          temp2 = 0
          If (k <> n) Then temp2 = x(k + 1)
          fvec(k) = temp - temp1 - 2 * temp2 + 1
       Next k
End Sub


Private Sub hybrd(fcn2 As String, ByRef n As Single, ByRef x As Variant, ByRef fvec As Variant, _
ByRef XTOL As Double, ByRef MAXFEV As Single, ByRef ML As Single, ByRef mu As Single, _
ByRef EPSFCN As Double, ByRef diag As Variant, _
ByRef MODE As Single, ByRef FACTOR As Double, ByRef NPRINT As Single, ByRef Info As Single, _
ByRef NFEV As Single, ByRef fjac As Variant, ByRef LDFJAC As Single, ByRef r As Variant, _
ByRef LR As Single, ByRef QTF As Variant, ByRef WA1 As Variant, _
ByRef WA2 As Variant, ByRef WA3 As Variant, ByRef WA4 As Variant)
                
                      
      'integer n,maxfev,ml,mu,mode,nprint,info,nfev,ldfjac,lr
      'double precision xtol,epsfcn,factor
      'double precision x(n),fvec(n),diag(n),fjac(ldfjac,n),r(lr),
     '                 qtf(n),wa1(n),wa2(n),wa3(n),wa4(n)
     
     ' external fcn
     Dim i As Single
     Dim iflag As Single
     Dim iter As Single
     Dim j As Single
     Dim jm1 As Single
     Dim l As Single
     Dim msum As Single
     Dim ncfail As Single
     Dim ncsuc As Single
     Dim nslow1 As Single
     Dim nslow2 As Single
     ReDim iwa(1 To 1)
     Dim jeval As Boolean
     Dim sing As Boolean
     
      'integer i,iflag,iter,j,jm1,l,msum,ncfail,ncsuc,nslow1,nslow2
      'integer iwa(1)
      'logical jeval, sing
      Dim actred As Double
      Dim DELTA As Double
      Dim epsmch As Double
      Dim FNORM As Double
      Dim fnorm1 As Double
      Dim one As Double
      Dim pnorm As Double
      Dim prered As Double
      Dim p1 As Double
      Dim P5 As Double
      Dim p001 As Double
      Dim p0001 As Double
      Dim ratio As Double
      Dim Sum As Double
      Dim temp As Double
      Dim xnorm As Double
      Dim ZERO As Double
      
      'double precision actred,delta,epsmch,fnorm,fnorm1,one,pnorm,
     '*                 prered,p1,p5,p001,p0001,ratio,sum,temp,xnorm,
     '*                 zero
      'double precision dpmpar,enorm
      one = 1
      p1 = 0.1
      P5 = 0.5
      p001 = 0.001
      p0001 = 0.0001
      ZERO = 0
      
      'data one, p1, p5, p001, p0001, zero
     '*     /1.0d0,1.0d-1,5.0d-1,1.0d-3,1.0d-4,0.0d0/
'
'     epsmch is the machine precision.
'
      epsmch = 1.2E-16
'
      Info = 0
      iflag = 0
      NFEV = 0
'
'     check the input parameters for errors.
'
      If ((n <= 0) Or (XTOL < ZERO) Or (MAXFEV <= 0) _
         Or (ML < 0) Or (mu < 0) Or (FACTOR <= ZERO) _
         Or (LDFJAC < n) Or (LR < (n * (n + 1)) / 2)) Then GoTo L300
         
      If (MODE <> 2) Then GoTo L20
      For j = 1 To n
         If (diag(j) <= ZERO) Then GoTo L300
      Next j
L20:
'
'     evaluate the function at the starting point
'     and calculate its norm.
'
      iflag = 1
      fcn n, x, fvec, iflag
      NFEV = 1
      If (iflag < 0) Then GoTo L300
      FNORM = myenorm(n, fvec)
'
'     determine the number of calls to fcn needed to compute
'     the jacobian matrix.
'
      msum = min0(ML + mu + 1, n)
'
'     initialize iteration counter and monitors.
'
      iter = 1
      ncsuc = 0
      ncfail = 0
      nslow1 = 0
      nslow2 = 0
'
'     beginning of the outer loop.
'
L30:
         jeval = True
'
'        calculate the jacobian matrix.
'
         iflag = 2
         fdjac1 fcn2, n, x, fvec, fjac, LDFJAC, iflag, ML, mu, EPSFCN, WA1, WA2
         NFEV = NFEV + msum
         If (iflag < 0) Then GoTo L300
'
'        compute the qr factorization of the jacobian.
'
         qrfac n, n, fjac, LDFJAC, False, iwa, 1, WA1, WA2, WA3
'
'        on the first iteration and if mode is 1, scale according
'        to the norms of the columns of the initial jacobian.
'
         If (iter <> 1) Then GoTo L70
         If (MODE = 2) Then GoTo L50
         For j = 1 To n
            diag(j) = WA2(j)
            If (WA2(j) = ZERO) Then diag(j) = one
         Next j
L50:
'
'        on the first iteration, calculate the norm of the scaled x
'        and initialize the step bound delta.
'
         For j = 1 To n
            WA3(j) = diag(j) * x(j)
         Next j
         xnorm = myenorm(n, WA3)
         DELTA = FACTOR * xnorm
         If (DELTA = ZERO) Then DELTA = FACTOR
L70:
'
'        form (q transpose)*fvec and store in qtf.
'
         For i = 1 To n
            QTF(i) = fvec(i)
         Next i
         For j = 1 To n
            If (fjac(j, j) = ZERO) Then GoTo L110
            Sum = ZERO
            For i = j To n
               Sum = Sum + fjac(i, j) * QTF(i)
            Next i
            temp = -Sum / fjac(j, j)
            For i = j To n
               QTF(i) = QTF(i) + fjac(i, j) * temp
            Next i
L110:
         Next j
'
'        copy the triangular factor of the qr factorization into r.
'
         sing = False
         For j = 1 To n
            l = j
            jm1 = j - 1
            If (jm1 < 1) Then GoTo L140
            For i = 1 To jm1
               r(l) = fjac(i, j)
               l = l + n - i
            Next i
L140:
            r(l) = WA1(j)
            If (WA1(j) = ZERO) Then sing = True
         Next j
'
'        accumulate the orthogonal factor in fjac.
'
          qform n, n, fjac, LDFJAC, WA1
'
'        rescale if necessary.
'
         If (MODE = 2) Then GoTo L170
         For j = 1 To n
            diag(j) = dmax1(diag(j), WA2(j))
         Next j
L170:
'
'        beginning of the inner loop.
'
L180:
'
'           if requested, call fcn to enable printing of iterates.
'
            If (NPRINT <= 0) Then GoTo L190
            iflag = 0
            If (mymod(iter - 1, NPRINT) = 0) Then
              fcn n, x, fvec, iflag
            End If
            If (iflag < 0) Then GoTo L300
L190:
'
'           determine the direction p.
'
            dogleg n, r, LR, diag, QTF, DELTA, WA1, WA2, WA3
'
'           store the direction p and x + p. calculate the norm of p.
'
            For j = 1 To n
               WA1(j) = -WA1(j)
               WA2(j) = x(j) + WA1(j)
               WA3(j) = diag(j) * WA1(j)
            Next j
            pnorm = myenorm(n, WA3)
'
'           on the first iteration, adjust the initial step bound.
'
            If (iter = 1) Then DELTA = dmin1(DELTA, pnorm)
'
'           evaluate the function at x + p and calculate its norm.
'
            iflag = 1
            fcn n, WA2, WA4, iflag
            NFEV = NFEV + 1
            If (iflag < 0) Then GoTo L300
            fnorm1 = myenorm(n, WA4)
'
'           compute the scaled actual reduction.
'
            actred = -one
            If (fnorm1 < FNORM) Then actred = one - (fnorm1 / FNORM) ^ 2
'
'           compute the scaled predicted reduction.
'
            l = 1
            For i = 1 To n
               Sum = ZERO
               For j = i To n
                  Sum = Sum + r(l) * WA1(j)
                  l = l + 1
               Next j
               WA3(i) = QTF(i) + Sum
            Next i
            temp = myenorm(n, WA3)
            prered = ZERO
            If (temp < FNORM) Then prered = one - (temp / FNORM) ^ 2
'
'           compute the ratio of the actual to the predicted
'           reduction.
'
            ratio = ZERO
            If (prered > ZERO) Then ratio = actred / prered
'
'           update the step bound.
'
            If (ratio >= p1) Then GoTo L230
               ncsuc = 0
               ncfail = ncfail + 1
               DELTA = P5 * DELTA
               GoTo L240
L230:
               ncfail = 0
               ncsuc = ncsuc + 1
               If ((ratio >= P5) Or (ncsuc > 1)) Then DELTA = dmax1(DELTA, pnorm / P5)
               If (dabs(ratio - one) <= p1) Then DELTA = pnorm / P5
L240:
'
'           test for successful iteration.
'
            If (ratio < p0001) Then GoTo L260
'
'           successful iteration. update x, fvec, and their norms.
'
            For j = 1 To n
               x(j) = WA2(j)
               WA2(j) = diag(j) * x(j)
               fvec(j) = WA4(j)
            Next j
            xnorm = myenorm(n, WA2)
            FNORM = fnorm1
            iter = iter + 1
L260:
'
'           determine the progress of the iteration.
'
            nslow1 = nslow1 + 1
            If (actred >= p001) Then nslow1 = 0
            If (jeval) Then nslow2 = nslow2 + 1
            If (actred >= p1) Then nslow2 = 0
'
'           test for convergence.
'
            If ((DELTA <= XTOL * xnorm) Or (FNORM = ZERO)) Then Info = 1
            If (Info <> 0) Then GoTo L300
'
'           tests for termination and stringent tolerances.
'
            If (NFEV >= MAXFEV) Then Info = 2
            If (p1 * dmax1(p1 * DELTA, pnorm) <= epsmch * xnorm) Then Info = 3
            If (nslow2 = 5) Then Info = 4
            If (nslow1 = 10) Then Info = 5
            If (Info <> 0) Then GoTo L300
'
'           criterion for recalculating jacobian approximation
'           by forward differences.
'
            If (ncfail = 2) Then GoTo L290
'
'           calculate the rank one modification to the jacobian
'           and update qtf if necessary.
'
            For j = 1 To n
               Sum = ZERO
               For i = 1 To n
                  Sum = Sum + fjac(i, j) * WA4(i)
               Next i
               WA2(j) = (Sum - WA3(j)) / pnorm
               WA1(j) = diag(j) * ((diag(j) * WA1(j)) / pnorm)
               If (ratio >= p0001) Then QTF(j) = Sum
            Next j
'
'           compute the qr factorization of the updated jacobian.
'
            r1updt n, n, r, LR, WA1, WA2, WA3, sing
            r1mpyq n, n, fjac, LDFJAC, WA2, WA3
            r1mpyqforqtf 1, n, QTF, 1, WA2, WA3
'
'           end of the inner loop.
'
            jeval = False
            GoTo L180
L290:
'
'        end of the outer loop.
'
         GoTo L30
L300:
'
'     termination, either normal or user imposed.
'
      If (iflag < 0) Then Info = iflag
      iflag = 0
      If (NPRINT > 0) Then fcn n, x, fvec, iflag
End Sub


 
 Private Sub r1updt(ByRef m As Single, ByRef n As Single, ByRef s As Variant, _
 ByRef ls As Single, ByRef U As Variant, ByRef V As Variant, _
 ByRef W As Variant, ByRef sing As Boolean)
 

      'integer m,n,ls
      'logical sing
      'double precision s(ls),u(m),v(n),w(m)
      Dim i As Single
      Dim j As Single
      Dim jj As Single
      Dim l As Single
      Dim nmj As Single
      Dim nm1 As Single
      'integer i,j,jj,l,nmj,nm1
      Dim mycos As Double
      Dim mycotan As Double
      Dim giant As Double
      Dim one As Double
      Dim P5 As Double
      Dim p25 As Double
      Dim mysin As Double
      Dim mytan As Double
      Dim tau As Double
      Dim temp As Double
      Dim ZERO As Double
      'double precision cos,cotan,giant,one,p5,p25,sin,tan,tau,temp,
     '*                 zero
      'double precision dpmpar
      one = 1
      P5 = 0.5
      p25 = 0.25
      ZERO = 0
      'data one,p5,p25,zero /1.0d0,5.0d-1,2.5d-1,0.0d0/
'
'     giant is the largest magnitude.
'
      giant = 2 ^ 55
'
'     initialize the diagonal element pointer.
'
      jj = (n * (2 * m - n + 1)) / 2 - (m - n)
'
'     move the nontrivial part of the last column of s into w.
'
      l = jj
      For i = n To m
         W(i) = s(l)
         l = l + 1
      Next i
'
'     rotate the vector v into a multiple of the n-th unit vector
'     in such a way that a spike is introduced into w.
'
      nm1 = n - 1
      If (nm1 < 1) Then GoTo L70
      For nmj = 1 To nm1
         j = n - nmj
         jj = jj - (m - j + 1)
         W(j) = ZERO
         If (V(j) = ZERO) Then GoTo L50
'
'        determine a givens rotation which eliminates the
'        j-th element of v.
'
         If (dabs(V(n)) >= dabs(V(j))) Then GoTo L20
            mycotan = V(n) / V(j)
            mysin = P5 / dsqrt(p25 + p25 * mycotan ^ 2)
            mycos = mysin * mycotan
            tau = one
            If (dabs(mycos) * giant > one) Then tau = one / mycos
            GoTo L30
L20:
            mytan = V(j) / V(n)
            mycos = P5 / dsqrt(p25 + p25 * mytan ^ 2)
            mysin = mycos * mytan
            tau = mysin
L30:
'
'        apply the transformation to v and store the information
'        necessary to recover the givens rotation.
'
         V(n) = mysin * V(j) + mycos * V(n)
         V(j) = tau
'
'        apply the transformation to s and extend the spike in w.
'
         l = jj
         For i = j To m
            temp = mycos * s(l) - mysin * W(i)
            W(i) = mysin * s(l) + mycos * W(i)
            s(l) = temp
            l = l + 1
         Next i
L50:
      Next nmj
L70:
'
'     add the spike from the rank 1 update to w.
'
      For i = 1 To m
         W(i) = W(i) + V(n) * U(i)
      Next i
'
'     eliminate the spike.
'
      sing = False
      If (nm1 < 1) Then GoTo L140
      For j = 1 To nm1
         If (W(j) = ZERO) Then GoTo L120
'
'        determine a givens rotation which eliminates the
'        j-th element of the spike.
'
         If (dabs(s(jj)) >= dabs(W(j))) Then GoTo L90
            mycotan = s(jj) / W(j)
            mysin = P5 / dsqrt(p25 + p25 * mycotan ^ 2)
            mycos = mysin * mycotan
            tau = one
            If (dabs(mycos) * giant > one) Then tau = one / mycos
            GoTo L100
L90:
            mytan = W(j) / s(jj)
            mycos = P5 / dsqrt(p25 + p25 * mytan ^ 2)
            mysin = mycos * mytan
            tau = mysin
L100:
'
'        apply the transformation to s and reduce the spike in w.
'
         l = jj
         For i = j To m
            temp = mycos * s(l) + mysin * W(i)
            W(i) = -mysin * s(l) + mycos * W(i)
            s(l) = temp
            l = l + 1
         Next i
'
'        store the information necessary to recover the
'        givens rotation.
'
         W(j) = tau
L120:
'
'        test for zero diagonal elements in the output s.
'
         If (s(jj) = ZERO) Then sing = True
         jj = jj + (m - j + 1)
      Next j
L140:
'
'     move w back into the last column of the output s.
'
      l = jj
      For i = n To m
         s(l) = W(i)
         l = l + 1
      Next i
      If (s(jj) = ZERO) Then sing = True
 End Sub
 
 
 
 
 Private Function min1single(A As Single, B As Single)
If A <= B Then
  min1single = A
Else
  min1single = B
End If
End Function
 Private Sub qrfac(ByRef m As Single, ByRef n As Single, ByRef A As Variant, ByRef lda As Single, _
   ByRef pivot As Boolean, ByRef ipvt As Variant, ByRef lipvt As Single, ByRef rdiag As Variant, _
    ByRef acnorm As Variant, ByRef WA As Variant)
    Dim i As Single
    Dim j As Single
    Dim k As Single
    Dim jp1 As Single
    Dim kmax As Single
    Dim minmn As Single
    Dim ajnorm As Double
    Dim epsmch As Double
    Dim one As Double
    Dim p05 As Double
    Dim Sum As Double
    Dim temp As Double
    Dim ZERO As Double
    one = 1
    p05 = 0.05
    ZERO = 0
    
    ReDim tmpaarr(1 To m) As Variant
    Dim tmpj As Single
    
'
'     epsmch is the machine precision.
'
      epsmch = 1.2E-16
'
'     compute the initial column norms and initialize several arrays.
'
      For j = 1 To n
         For tmpj = 1 To m
           tmpaarr(tmpj) = A(tmpj, j)
         Next tmpj
         'acnorm(j) = enorm(m, a(1, j))
         
         'a3 = enorm(m, a2)
         'testenorm tmpaarr
         acnorm(j) = myenorm(m, tmpaarr)
         rdiag(j) = acnorm(j)
         WA(j) = rdiag(j)
         If (pivot) Then ipvt(j) = j
      Next j
'
'     reduce a to r with householder transformations.
'
      minmn = min1single(m, n)
      For j = 1 To minmn
         If (Not pivot) Then GoTo L40
'
'        bring the column of largest norm into the pivot position.
'
         kmax = j
         For k = j To n
            If (rdiag(k) > rdiag(kmax)) Then kmax = k
         Next k
         If (kmax = j) Then GoTo L40
         For i = 1 To m
            temp = A(i, j)
            A(i, j) = A(i, kmax)
            A(i, kmax) = temp
         Next i
         rdiag(kmax) = rdiag(j)
         WA(kmax) = WA(j)
         k = ipvt(j)
         ipvt(j) = ipvt(kmax)
         ipvt(kmax) = k
L40:
'
'        compute the householder transformation to reduce the
'        j-th column of a to a multiple of the j-th unit vector.
'
         ReDim tmpaarr(1 To m - j + 1) As Variant
         For tmpj = 1 To m - j + 1
           tmpaarr(tmpj) = A(tmpj + j - 1, j)
         Next tmpj
         'ajnorm = myenorm(m - j + 1, a(j, j))
         ajnorm = myenorm(m - j + 1, tmpaarr)
         If (ajnorm = ZERO) Then GoTo L100
         If (A(j, j) < ZERO) Then ajnorm = -ajnorm
         For i = j To m
            A(i, j) = A(i, j) / ajnorm
         Next i
         A(j, j) = A(j, j) + one
'
'        apply the transformation to the remaining columns
'        and update the norms.
'
         jp1 = j + 1
         If (n < jp1) Then GoTo L100
         For k = jp1 To n
            Sum = ZERO
            For i = j To m
               Sum = Sum + A(i, j) * A(i, k)
            Next i
            temp = Sum / A(j, j)
            For i = j To m
               A(i, k) = A(i, k) - temp * A(i, j)
            Next i
            If ((Not pivot) Or (rdiag(k) = ZERO)) Then GoTo L80
            temp = A(j, k) / rdiag(k)
            rdiag(k) = rdiag(k) * dsqrt(dmax1(ZERO, one - temp * temp))
            If (p05 * (rdiag(k) / WA(k)) ^ 2 > epsmch) Then GoTo L80
            rdiag(k) = myenorm(m - j, A(jp1, k))
            WA(k) = rdiag(k)
L80:
         Next k
L100:
         rdiag(j) = -ajnorm
      Next j
      'Return
'
'     last card of subroutine qrfac.
'
 End Sub
Private Function mymod(x As Single, Y As Single) As Single
mymod = 1
End Function

Private Sub dogleg(ByRef n As Single, ByRef r As Variant, ByRef LR As Single, _
ByRef diag As Variant, ByRef qtb As Variant, ByRef DELTA As Double, _
ByRef x As Variant, ByRef WA1 As Variant, ByRef WA2 As Variant)
'subroutine dogleg(n, r, lr, diag, qtb, delta, x, wa1, wa2)
      'integer n,lr
      'double precision delta
      'double precision r(lr),diag(n),qtb(n),x(n),wa1(n),wa2(n)
Dim i As Single
Dim j As Single
Dim jj As Single
Dim jp1 As Single
Dim k As Single
Dim l As Single
Dim ALPHA As Double
Dim bnorm As Double
Dim epsmch As Double
Dim gnorm As Double
Dim one As Double
Dim qnorm As Double
Dim sgnorm As Double
Dim Sum As Double
Dim temp As Double
Dim ZERO As Double
one = 1
ZERO = 0
epsmch = 1E-16
'      integer i,j,jj,jp1,k,l
'      double precision alpha,bnorm,epsmch,gnorm,one,qnorm,sgnorm,sum,
'     *                 temp,zero
'      double precision dpmpar,enorm
'      data one,zero /1.0d0,0.0d0/
      'epsmch = dpmpar(1)
'
'     first, calculate the gauss-newton direction.
'
      jj = (n * (n + 1)) / 2 + 1
      For k = 1 To n
         j = n - k + 1
         jp1 = j + 1
         jj = jj - k
         l = jj + 1
         Sum = ZERO
         If (n < jp1) Then GoTo L20
         For i = jp1 To n
            Sum = Sum + r(l) * x(i)
            l = l + 1
         Next i
L20:
         temp = r(jj)
         If (temp <> ZERO) Then GoTo L40
         l = j
         For i = 1 To j
            temp = dmax1(temp, dabs(r(l)))
            l = l + n - i
         Next i
         temp = epsmch * temp
         If (temp = ZERO) Then temp = epsmch
L40:
         x(j) = (qtb(j) - Sum) / temp
      Next k
'
'     test whether the gauss-newton direction is acceptable.
'
      For j = 1 To n
         WA1(j) = ZERO
         WA2(j) = diag(j) * x(j)
      Next j
      qnorm = myenorm(n, WA2)
      If (qnorm <= DELTA) Then GoTo L140
'
'     the gauss-newton direction is not acceptable.
'     next, calculate the scaled gradient direction.
'
      l = 1
      For j = 1 To n
         temp = qtb(j)
         For i = j To n
            WA1(i) = WA1(i) + r(l) * temp
            l = l + 1
         Next i
         WA1(j) = WA1(j) / diag(j)
      Next j
'
'     calculate the norm of the scaled gradient and test for
'     the special case in which the scaled gradient is zero.
'
      gnorm = myenorm(n, WA1)
      sgnorm = ZERO
      ALPHA = DELTA / qnorm
      If (gnorm = ZERO) Then GoTo L120
'
'     calculate the point along the scaled gradient
'     at which the quadratic is minimized.
'
      For j = 1 To n
         WA1(j) = (WA1(j) / gnorm) / diag(j)
      Next j
      l = 1
      For j = 1 To n
         Sum = ZERO
         For i = j To n
            Sum = Sum + r(l) * WA1(i)
            l = l + 1
         Next i
         WA2(j) = Sum
      Next j
      temp = myenorm(n, WA2)
      sgnorm = (gnorm / temp) / temp
'
'     test whether the scaled gradient direction is acceptable.
'
      ALPHA = ZERO
      If (sgnorm >= DELTA) Then GoTo L120
'
'     the scaled gradient direction is not acceptable.
'     finally, calculate the point along the dogleg
'     at which the quadratic is minimized.
'
      bnorm = myenorm(n, qtb)
      temp = (bnorm / gnorm) * (bnorm / qnorm) * (sgnorm / DELTA)
      temp = temp - (DELTA / qnorm) * (sgnorm / DELTA) ^ 2 _
            + dsqrt((temp - (DELTA / qnorm)) ^ 2 _
                    + (one - (DELTA / qnorm) ^ 2) * (one - (sgnorm / DELTA) ^ 2))
      ALPHA = ((DELTA / qnorm) * (one - (sgnorm / DELTA) ^ 2)) / temp
L120:
'
'     form appropriate convex combination of the gauss-newton
'     direction and the scaled gradient direction.
'
      temp = (one - ALPHA) * dmin1(sgnorm, DELTA)
      For j = 1 To n
         x(j) = temp * WA1(j) + ALPHA * x(j)
      Next j
L140:
      
End Sub
Private Function dmin1(A As Double, B As Double)
If A <= B Then
  dmin1 = A
Else
  dmin1 = B
End If
End Function
Private Function dmax1(A, B)
If A >= B Then
  dmax1 = A
Else
  dmax1 = B
End If
End Function

Private Sub fdjac1(fcn2 As String, ByRef n As Single, ByRef x As Variant, ByRef fvec As Variant, _
ByRef fjac As Variant, ByRef LDFJAC As Variant, ByRef iflag As Single, _
ByRef ML As Single, ByRef mu As Single, ByRef EPSFCN As Double, _
ByRef WA1 As Variant, ByRef WA2 As Variant)



 'subroutine fdjac1(fcn,n,x,fvec,fjac,ldfjac,iflag,ml,mu,epsfcn,
 '    *                  wa1,wa2)
 '     integer n,ldfjac,iflag,ml,mu
 '     double precision epsfcn
 '     double precision x(n),fvec(n),fjac(ldfjac,n),wa1(n),wa2(n)
 Dim i As Single
 Dim j As Single
 Dim k As Single
 Dim msum As Single
 Dim eps As Double
 Dim epsmch As Double
 Dim h As Double
 Dim temp As Double
 Dim ZERO As Double
 ZERO = 0
 
 '     integer i,j,k,msum
 '     double precision eps,epsmch,h,temp,zero
 '     double precision dpmpar
 '     data zero /0.0d0/
      epsmch = 1E-16
'
      eps = dsqrt(dmax1(EPSFCN, epsmch))
      msum = ML + mu + 1
      If (msum < n) Then GoTo L40
'
'        computation of dense approximate jacobian.
'
         For j = 1 To n
            temp = x(j)
            h = eps * dabs(temp)
            If (h = ZERO) Then h = eps
            x(j) = temp + h
            'redim paramArray(1 to
            fcn n, x, WA1, iflag
            If (iflag < 0) Then GoTo L30
            x(j) = temp
            For i = 1 To n
               fjac(i, j) = (WA1(i) - fvec(i)) / h
            Next i
         Next j
L30:
         GoTo L110
L40:
'
'        computation of banded approximate jacobian.
'
         For k = 1 To msum
            For j = k To n Step msum
               WA2(j) = x(j)
               h = eps * dabs(WA2(j))
               If (h = ZERO) Then h = eps
               x(j) = WA2(j) + h
            Next j
            fcn n, x, WA1, iflag
            If (iflag < 0) Then GoTo L100
            For j = k To n Step msum
               x(j) = WA2(j)
               h = eps * dabs(WA2(j))
               If (h = ZERO) Then h = eps
               For i = 1 To n
                  fjac(i, j) = ZERO
                  If ((i >= (j - mu)) And (i <= j + ML)) Then
                    fjac(i, j) = (WA1(i) - fvec(i)) / h
                  End If
               Next i
            Next j
         Next k
L100:
L110:
End Sub



Private Function myenorm(n As Single, x As Variant) As Double
      Dim i As Single
      Dim agiant As Double
      Dim floatn As Double
      Dim one As Double
      Dim rdwarf As Double
      Dim rgiant As Double
      Dim s1 As Double
      Dim s2 As Double
      Dim s3 As Double
      Dim xabs As Double
      Dim x1max As Double
      Dim x3max As Double
      Dim ZERO As Double
      
      one = 1
      ZERO = 0
      rdwarf = 3.834E-20
      rgiant = 1.304E+19
      s1 = ZERO
      s2 = ZERO
      s3 = ZERO
      x1max = ZERO
      x3max = ZERO
      floatn = n
      agiant = rgiant / floatn
      For i = 1 To n
         xabs = dabs(CDbl(x(i)))
         If ((xabs > rdwarf) And (xabs < agiant)) Then GoTo L70
            If (xabs <= rdwarf) Then GoTo L30
'
'              sum for large components.
'
               If (xabs <= x1max) Then GoTo L10
                  s1 = one + s1 * (x1max / xabs) ^ 2
                  x1max = xabs
                  GoTo L20
L10:
                  s1 = s1 + (xabs / x1max) ^ 2
L20:
               GoTo L60
L30:
'
'              sum for small components.
'
               If (xabs <= x3max) Then GoTo L40
                  s3 = one + s3 * (x3max / xabs) ^ 2
                  x3max = xabs
                  GoTo L50
L40:
                  If (xabs <> ZERO) Then s3 = s3 + (xabs / x3max) ^ 2
L50:
L60:
            GoTo L80
L70:
'
'           sum for intermediate components.
'
            s2 = s2 + xabs ^ 2
L80:
      Next i
'
'     calculation of norm.
'
      If (s1 = ZERO) Then GoTo L100
         myenorm = x1max * dsqrt(s1 + (s2 / x1max) / x1max)
         GoTo L130
L100:
         If (s2 = ZERO) Then GoTo L110
            If (s2 >= x3max) Then _
              myenorm = dsqrt(s2 * (one + (x3max / s2) * (x3max * s3)))
            If (s2 < x3max) Then _
              myenorm = dsqrt(x3max * ((s2 / x3max) + (x3max * s3)))
            GoTo L120
L110:
            myenorm = x3max * dsqrt(s3)
L120:
L130:
'
'     last card of function enorm.
'
End Function


Private Function dsqrt(x)
dsqrt = x ^ 0.5
End Function
Private Function dabs(x)
dabs = Abs(x)
End Function
Private Sub r1mpyq(ByRef m As Single, ByRef n As Single, ByRef A As Variant, _
ByRef lda As Single, ByRef V As Variant, ByRef W As Variant)


      'subroutine r1mpyq(m, n, a, lda, v, w)
      'integer m,n,lda
      'double precision a(lda,n),v(n),w(n)
Dim i As Single
Dim j As Single
Dim nmj As Single
Dim nm1 As Single
Dim mycos As Double
Dim one As Double
Dim mysin As Double
Dim temp As Double
one = 1

      'integer i,j,nmj,nm1
'      double precision cos,one,sin,temp
'      data one /1.0d0/
'
'     apply the first set of givens rotations to a.
'
      nm1 = n - 1
      If (nm1 < 1) Then GoTo L50
      For nmj = 1 To nm1
         j = n - nmj
         If (dabs(V(j)) > one) Then mycos = one / V(j)
         If (dabs(V(j)) > one) Then mysin = dsqrt(one - mycos ^ 2)
         If (dabs(V(j)) <= one) Then mysin = V(j)
         If (dabs(V(j)) <= one) Then mycos = dsqrt(one - mysin ^ 2)
         For i = 1 To m
            temp = mycos * A(i, j) - mysin * A(i, n)
            A(i, n) = mysin * A(i, j) + mycos * A(i, n)
            A(i, j) = temp
         Next i
      Next nmj
'
'     apply the second set of givens rotations to a.
'
      For j = 1 To nm1
         If (dabs(W(j)) > one) Then mycos = one / W(j)
         If (dabs(W(j)) > one) Then mysin = dsqrt(one - mycos ^ 2)
         If (dabs(W(j)) <= one) Then mysin = W(j)
         If (dabs(W(j)) <= one) Then mycos = dsqrt(one - mysin ^ 2)
         For i = 1 To m
            temp = mycos * A(i, j) + mysin * A(i, n)
            A(i, n) = -mysin * A(i, j) + mycos * A(i, n)
            A(i, j) = temp
         Next i
      Next j
L50:
End Sub
Private Sub r1mpyqforqtf(ByRef m As Single, ByRef n As Single, ByRef A As Variant, _
ByRef lda As Single, ByRef V As Variant, ByRef W As Variant)


      'subroutine r1mpyq(m, n, a, lda, v, w)
      'integer m,n,lda
      'double precision a(lda,n),v(n),w(n)
Dim i As Single
Dim j As Single
Dim nmj As Single
Dim nm1 As Single
Dim mycos As Double
Dim one As Double
Dim mysin As Double
Dim temp As Double
one = 1

      'integer i,j,nmj,nm1
'      double precision cos,one,sin,temp
'      data one /1.0d0/
'
'     apply the first set of givens rotations to a.
'
      nm1 = n - 1
      If (nm1 < 1) Then GoTo L50
      For nmj = 1 To nm1
         j = n - nmj
         If (dabs(V(j)) > one) Then mycos = one / V(j)
         If (dabs(V(j)) > one) Then mysin = dsqrt(one - mycos ^ 2)
         If (dabs(V(j)) <= one) Then mysin = V(j)
         If (dabs(V(j)) <= one) Then mycos = dsqrt(one - mysin ^ 2)
         For i = 1 To m
            temp = mycos * A(j) - mysin * A(n)
            A(n) = mysin * A(j) + mycos * A(n)
            A(j) = temp
         Next i
      Next nmj
'
'     apply the second set of givens rotations to a.
'
      For j = 1 To nm1
         If (dabs(W(j)) > one) Then mycos = one / W(j)
         If (dabs(W(j)) > one) Then mysin = dsqrt(one - mycos ^ 2)
         If (dabs(W(j)) <= one) Then mysin = W(j)
         If (dabs(W(j)) <= one) Then mycos = dsqrt(one - mysin ^ 2)
         For i = 1 To m
            temp = mycos * A(j) + mysin * A(n)
            A(n) = -mysin * A(j) + mycos * A(n)
            A(j) = temp
         Next i
      Next j
L50:
End Sub




Private Function min0(A, B)
If A <= B Then
  min0 = A
Else
  min0 = B
End If
End Function
Private Sub qform(ByRef m As Single, ByRef n As Single, ByRef q As Variant, _
ByRef ldq As Single, ByRef WA As Variant)

 Dim i As Single
 Dim j As Single
 Dim jm1 As Single
 Dim k As Single
 Dim l As Single
 Dim minmn As Single
 Dim np1 As Single
 Dim one As Double
 Dim Sum As Double
 Dim temp As Double
 Dim ZERO As Double
 one = 1
 Sum = 0
 temp = 1
 ZERO = 0
 
 '     integer i,j,jm1,k,l,minmn,np1
  '    double precision one,sum,temp,zero
   '   data one,zero /1.0d0,0.0d0/
'
'     zero out upper triangle of q in the first min(m,n) columns.
'
      minmn = min0(m, n)
      If minmn < 2 Then GoTo L30
      For j = 2 To minmn
         jm1 = j - 1
         For i = 1 To jm1
            q(i, j) = ZERO
         Next i
      Next j
      
L30:
'
'     initialize remaining columns to those of the identity matrix.
'
      np1 = n + 1
      If (m < np1) Then GoTo L60
      For j = np1 To m
        For i = 1 To m
            q(i, j) = ZERO
        Next i
        q(j, j) = one
      Next j
L60:
'
'     accumulate q from its factored form.
'
      For l = 1 To minmn
         k = minmn - l + 1
         For i = k To m
            WA(i) = q(i, k)
            q(i, k) = ZERO
         Next i
         q(k, k) = one
         If (WA(k) = ZERO) Then GoTo L110
         For j = k To m
            Sum = ZERO
            For i = k To m
               Sum = Sum + q(i, j) * WA(i)
            Next i
            temp = Sum / WA(k)
            For i = k To m
               q(i, j) = q(i, j) - temp * WA(i)
            Next i
         Next j
L110:
      Next l
End Sub


