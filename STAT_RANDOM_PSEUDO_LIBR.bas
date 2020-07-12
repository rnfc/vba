Attribute VB_Name = "STAT_RANDOM_PSEUDO_LIBR"


Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.


'************************************************************************************
'************************************************************************************
'FUNCTION      : RANDOM_FMRG_FUNC
'DESCRIPTION   : Function for multiple recursive generators of modulus p and order
'k where all nonzero coefficients of the recurrence are equal. This method
'is proposed by Deng and Lin (2000) as a special case. The advantage of this
'kind of generator is that a single multiplication is needed to compute the
'recurrence, so the generator would run faster than the general case.

'LIBRARY       : STATISTICS
'GROUP         : RANDOM_PSEUDO
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 14/02/2008
'************************************************************************************
'************************************************************************************

Function RANDOM_FMRG_FUNC(Optional ByVal VOLATILE_FLAG As Boolean = True)

If VOLATILE_FLAG = True Then: Excel.Application.Volatile (True)

Static TEMP_VAL As Currency '>15 digits
Static FIRST_VAL As Currency '>15 digits
Static SECOND_VAL As Currency '>15 digits
Static LAUNCH_FLAG As Long

Dim ii As Long
Dim jj As Integer
Dim kk As Integer
Dim FMRG_VAL As Double
Dim PSEUD_ARR As Variant

On Error GoTo ERROR_LABEL

'Debug.Print LAUNCH_FLAG

If LAUNCH_FLAG = 0 Then: GoSub 1983 'runs on first time used no need to run
'again in any macro

ii = 2147483647
'This calculates the next "random" number in the sequence
TEMP_VAL = (LAUNCH_FLAG * SECOND_VAL - FIRST_VAL) - _
  ii * Int((LAUNCH_FLAG * SECOND_VAL - FIRST_VAL) / ii)

FMRG_VAL = TEMP_VAL / ii
SECOND_VAL = FIRST_VAL
FIRST_VAL = TEMP_VAL

RANDOM_FMRG_FUNC = FMRG_VAL

Exit Function
'-------------------------------------------------------------------------------
1983:
'-------------------------------------------------------------------------------
Randomize
FIRST_VAL = Rnd * 2 ^ 24
SECOND_VAL = Rnd * 2 ^ 24

'Load values from p. 147 of Deng and Lin, "Random Number Generation for
'the New Century," The American Statistician, May 2000, vol. 54, no. 2

PSEUD_ARR = Array(26403, 27149, 29812, 30229, 31332, 33236, _
                33986, 34601, 36098, 36181, 36673, 36848, 37097, _
                37877, 39613, 40851, 40961, 42174, 42457, 43199, _
                43693, 44314, 44530, 45670, 46338)

If LBound(PSEUD_ARR) = 1 Then 'If Base =1 Then: kk = 24
    kk = UBound(PSEUD_ARR) - LBound(PSEUD_ARR)
Else 'If Base = 0 Then: kk = 25
    kk = UBound(PSEUD_ARR) - LBound(PSEUD_ARR) + 1
End If

jj = Rnd * kk + 1
LAUNCH_FLAG = PSEUD_ARR(jj)
'Pick a val, each with equal probability
Return

ERROR_LABEL:
RANDOM_FMRG_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : RANDOM_MERSENNE_TWISTER_FUNC
'DESCRIPTION   : This function is a pseudorandom number generator developed in
'1997 by Makoto Matsumoto and Takuji Nishimura that is based on a matrix linear
'recurrence over a finite binary field. It provides for fast generation of very
'high quality pseudorandom numbers, having been designed specifically to rectify
'many of the flaws found in older algorithms.

'Its name derives from the fact that period length is chosen to be a Mersenne prime.
'There are at least two common variants of the algorithm, differing only in the size
'of the Mersenne primes used. The newer and more commonly used one is the Mersenne
'Twister MT19937, with 32-bit word length. There is also a variant with 64-bit word
'length, MT19937-64, which generates a different sequence.
'See: http://en.wikipedia.org/wiki/Mersenne_Twister

'LIBRARY       : STATISTICS
'GROUP         : RANDOM_PSEUDO
'ID            : 002
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 14/02/2008
'************************************************************************************
'************************************************************************************

Function RANDOM_MERSENNE_TWISTER_FUNC(Optional ByVal OUTPUT As Integer = 0, _
Optional ByVal VOLATILE_FLAG As Boolean = False)

On Error GoTo ERROR_LABEL

Excel.Application.Volatile (VOLATILE_FLAG)

Select Case OUTPUT
Case 0
    RANDOM_MERSENNE_TWISTER_FUNC = genrand_real1()
'double genrand_real1(void)
'/* generates a random number on [0,1]-real-interval */
'return genrand_int32()*(1.0/4294967295.0);     '/* divided by 2^32-1 */
Case 1
    RANDOM_MERSENNE_TWISTER_FUNC = genrand_real2()
'double genrand_real2(void)
'/* generates a random number on [0,1)-real-interval */
'return genrand_int32()*(1.0/4294967296.0);     '/* divided by 2^32 */
Case 2
    RANDOM_MERSENNE_TWISTER_FUNC = genrand_real2b()
'Returns results in the
'range [0,1) == [0, 1-kMT_Gap2]
'Its lowest value is : 0.0
'Its highest value is: 0.9999999999990
Case 3
    RANDOM_MERSENNE_TWISTER_FUNC = genrand_real2c()
'Returns results in the range
'(0,1] == [0+kMT_Gap2, 1.0]
'Its lowest value is : 0.0000000000010  (1E-12)
'Its highest value is: 1.0
Case 4
    RANDOM_MERSENNE_TWISTER_FUNC = genrand_real3()
'double genrand_real3(void)
'/* generates a random number on (0,1)-real-interval */
'return (((double)genrand_int32()) + 0.5)*(1.0/4294967296.0);
'/* divided by 2^32 */

Case 5
    RANDOM_MERSENNE_TWISTER_FUNC = genrand_real3b()
'double genrand_real3(void)
'Returns results in the range (0,1) == [0+kMT_Gap, 1-kMT_Gap]
'Its lowest value is : 0.0000000000005  (5E-13)
'Its highest value is: 0.9999999999995
Case 6
    RANDOM_MERSENNE_TWISTER_FUNC = genrand_int31()
'long genrand_int31(void)
'/* generates a random number on [0,0x7fffffff]-interval */
'return (long)(genrand_int32()>>1);
Case 7
    RANDOM_MERSENNE_TWISTER_FUNC = genrand_int32()
'unsigned long genrand()
'Returns a value in the range [0, 2^32-1] (that is: [0, 4294967295] )
'   - The return type of the function is Double, not Long, but the values
'     returned are integers.
'   - If you want Long values in the range
'    [-2^31, 2^31-1] ([-2147483648, 2147483647]),
'     then call genrand_int32SignedLong() instead of this function.
Case 8
    RANDOM_MERSENNE_TWISTER_FUNC = genrand_int32SignedLong()
'unsigned long genrand_int32(void)
'This is the translation to VBA of the original C code for genrand_int32(),
'but renamed as explained in the section "Differences with the original C
'functions and source file"
'/* generates a random number on [0,0xffffffff]-interval */
'(Yes, BUT RETURNS IT AS A (signed) Long in the range [-2^31, 2^31-1])

Case 9
    RANDOM_MERSENNE_TWISTER_FUNC = genrand_real4b()
'Returns results in the
'range [-1,1] == [-1.0, 1.0]
'Its lowest value is : -1.0
'Its highest value is: 1.0

Case 10
    RANDOM_MERSENNE_TWISTER_FUNC = genrand_real5b()
'Returns results in the
'range (-1,1) == [-kMT_GapInterval, kMT_GapInterval]
'Its lowest value is : -0.9999999999990
'Its highest value is: 0.9999999999990
Case Else
    RANDOM_MERSENNE_TWISTER_FUNC = genrand_res53()
'double genrand_res53(void)
'/* generates a random number on [0,1) with 53-bit resolution*/
'unsigned long a=genrand_int32()>>5, b=genrand_int32()>>6;
'return(a*67108864.0+b)*(1.0/9007199254740992.0);

End Select

Exit Function
ERROR_LABEL:
RANDOM_MERSENNE_TWISTER_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : PSEUDO_RANDOM_FUNC
'DESCRIPTION   : Return a uniform a random number [0 1]
'LIBRARY       : STATISTICS
'GROUP         : RANDOM_PSEUDO
'ID            : 003
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 14/02/2008
'************************************************************************************
'************************************************************************************

Function PSEUDO_RANDOM_FUNC(Optional ByVal RANDOM_TYPE As Integer = 0)

Dim epsilon As Double

On Error GoTo ERROR_LABEL

epsilon = 0.999999999999999

'---------------------------------------------------------------
Select Case RANDOM_TYPE
'---------------------------------------------------------------
Case 0
'---------------------------------------------------------------
    PSEUDO_RANDOM_FUNC = Rnd
'---------------------------------------------------------------
Case Else 'TO AVOID ERRORS
'---------------------------------------------------------------
    PSEUDO_RANDOM_FUNC = Rnd
    If PSEUDO_RANDOM_FUNC = 1 Then
        PSEUDO_RANDOM_FUNC = epsilon
    End If
    If PSEUDO_RANDOM_FUNC = 0 Then
        PSEUDO_RANDOM_FUNC = 1 - epsilon
    End If
'---------------------------------------------------------------
End Select
'---------------------------------------------------------------

Exit Function
ERROR_LABEL:
PSEUDO_RANDOM_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : RANDOM_BETWEEN_FUNC
'DESCRIPTION   : Return random numbers from a Truncated UNIFORM Distribution
'LIBRARY       : STATISTICS
'GROUP         : RANDOM_PSEUDO
'ID            : 004
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 14/02/2008
'************************************************************************************
'************************************************************************************

Function RANDOM_BETWEEN_FUNC(ByVal LOW_MEAN As Double, _
ByVal HIGH_SIGMA As Double, _
Optional ByVal VERSION As Integer = 0)

'--> LOW_MEAN is equal either LOW_VALUE or MEAN_VALUE
'--> HIGH_SIGMA is equal either HIGH_VALUE or SIGMA_VALUE

Dim A_VAL As Double
Dim B_VAL As Double

Dim RANDOM_VAL As Double

On Error GoTo ERROR_LABEL

RANDOM_VAL = PSEUDO_RANDOM_FUNC(0)

Select Case VERSION
Case 0
    RANDOM_BETWEEN_FUNC = RANDOM_VAL * (HIGH_SIGMA - LOW_MEAN + 1) + LOW_MEAN
Case Else
    B_VAL = (2 * LOW_MEAN + Sqr((2 * LOW_MEAN) ^ 2 - 4 * (LOW_MEAN ^ 2 - 3 * (HIGH_SIGMA ^ 2)))) / 2
    A_VAL = 2 * LOW_MEAN - B_VAL
    RANDOM_BETWEEN_FUNC = RANDOM_VAL * (B_VAL - A_VAL) + A_VAL
End Select

Exit Function
ERROR_LABEL:
RANDOM_BETWEEN_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : VECTOR_RANDOM_INDEX_FUNC
'DESCRIPTION   : Return random numbers Vector Index
'LIBRARY       : STATISTICS
'GROUP         : RANDOM_PSEUDO
'ID            : 005
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 14/02/2008
'************************************************************************************
'************************************************************************************

Function VECTOR_RANDOM_INDEX_FUNC(ByVal NSIZE As Long, _
Optional ByVal RANDOM_TYPE As Integer = 0)

Dim i As Long
Dim j As Long
Dim TEMP_VAL As Double
Dim TEMP_VECTOR As Variant

On Error GoTo ERROR_LABEL

ReDim TEMP_VECTOR(1 To NSIZE, 1 To 1)
For i = 1 To NSIZE
    TEMP_VECTOR(i, 1) = i
Next i
For i = 1 To NSIZE
    j = Int(PSEUDO_RANDOM_FUNC(RANDOM_TYPE) * NSIZE) + 1
    TEMP_VAL = TEMP_VECTOR(i, 1)
    TEMP_VECTOR(i, 1) = TEMP_VECTOR(j, 1)
    TEMP_VECTOR(j, 1) = TEMP_VAL
Next i

VECTOR_RANDOM_INDEX_FUNC = TEMP_VECTOR

Exit Function
ERROR_LABEL:
VECTOR_RANDOM_INDEX_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : MATRIX_RANDOM_UNIFORM_FUNC
'DESCRIPTION   : Returns an array with uniformly distributed random numbers [0 1]
'LIBRARY       : STATISTICS
'GROUP         : RANDOM_PSEUDO
'ID            : 006
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 14/02/2008
'************************************************************************************
'************************************************************************************

Function MATRIX_RANDOM_UNIFORM_FUNC(ByVal NROWS As Long, _
ByVal NCOLUMNS As Long, _
Optional ByVal RANDOM_TYPE As Integer = 0, _
Optional ByVal VERSION As Integer = 0)

Dim i As Long
Dim j As Long

Dim TEMP_MATRIX As Variant

On Error GoTo ERROR_LABEL

ReDim TEMP_MATRIX(1 To NROWS, 1 To NCOLUMNS)

'-------------------------------------------------------------
Select Case VERSION
'-------------------------------------------------------------
Case 0
'-------------------------------------------------------------
    For j = 1 To NCOLUMNS
        For i = 1 To NROWS
            TEMP_MATRIX(i, j) = PSEUDO_RANDOM_FUNC(RANDOM_TYPE)
        Next i
    Next j
'-------------------------------------------------------------
Case Else
'-------------------------------------------------------------
    For i = 1 To NROWS
        TEMP_MATRIX(i, 1) = PSEUDO_RANDOM_FUNC(RANDOM_TYPE)
        For j = 2 To NCOLUMNS
            TEMP_MATRIX(i, j) = TEMP_MATRIX(i, 1)
        Next j
    Next i
'-------------------------------------------------------------
End Select
'-------------------------------------------------------------
    
MATRIX_RANDOM_UNIFORM_FUNC = TEMP_MATRIX

Exit Function
ERROR_LABEL:
MATRIX_RANDOM_UNIFORM_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : RANDOM_TRIANGULAR_FUNC
'DESCRIPTION   : Returns a random numbers from a Triangular Distribution
'LIBRARY       : STATISTICS
'GROUP         : RANDOM_PSEUDO
'ID            : 007
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 14/02/2008
'************************************************************************************
'************************************************************************************

Function RANDOM_TRIANGULAR_FUNC(ByVal PEAK_VALUE As Double, _
ByVal MIN_VALUE As Double, _
ByVal MAX_VALUE As Double, _
ByVal RANDOM_VAL As Double)

Dim TEMP_VAL As Double
Dim RATIO_VAL As Double

On Error GoTo ERROR_LABEL

RATIO_VAL = (PEAK_VALUE - MIN_VALUE) / (MAX_VALUE - MIN_VALUE)

If RANDOM_VAL <= RATIO_VAL Then
    TEMP_VAL = (MIN_VALUE + (MAX_VALUE - MIN_VALUE) * (RANDOM_VAL * RATIO_VAL) ^ 0.5)
Else
    TEMP_VAL = MIN_VALUE + (MAX_VALUE - MIN_VALUE) * (1 - ((1 - RATIO_VAL) * (1 - RANDOM_VAL)) ^ 0.5)
End If

RANDOM_TRIANGULAR_FUNC = TEMP_VAL

Exit Function
ERROR_LABEL:
RANDOM_TRIANGULAR_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : MATRIX_RANDOM_TRIANGULAR_FUNC
'DESCRIPTION   : Returns random numbers from a Triangular Distribution
'LIBRARY       : STATISTICS
'GROUP         : RANDOM_PSEUDO
'ID            : 008
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 14/02/2008
'************************************************************************************
'************************************************************************************

Function MATRIX_RANDOM_TRIANGULAR_FUNC(ByVal NROWS As Long, _
ByVal NCOLUMNS As Long, _
ByVal PEAK_VALUE As Double, _
ByVal MIN_VALUE As Double, _
ByVal MAX_VALUE As Double, _
Optional ByVal RANDOM_TYPE As Integer = 0)

Dim i As Long
Dim j As Long

Dim RATIO_VAL As Double
Dim RANDOM_MATRIX As Variant

On Error GoTo ERROR_LABEL

'The Triangular distribution is often used when no or little data is
'available.  It has 3 parameters, the minimum and the maximum that
'defines the range, and the more likely (the peak).  The distribution
'is skewed to the left when the peak is closed to the minimum and to
'the right when the peak is closed to the maximum.  It is a simple distribution
'that as its name implied, has a triangular shape.

ReDim TEMP_VECTOR(1 To NROWS, 1 To NCOLUMNS)

RATIO_VAL = (PEAK_VALUE - MIN_VALUE) / (MAX_VALUE - MIN_VALUE)
    
RANDOM_MATRIX = MATRIX_RANDOM_UNIFORM_FUNC(NROWS, NCOLUMNS, RANDOM_TYPE, 0)
For i = 1 To NROWS
    For j = 1 To NCOLUMNS
        RANDOM_MATRIX(i, j) = RANDOM_TRIANGULAR_FUNC(PEAK_VALUE, MIN_VALUE, MAX_VALUE, RANDOM_MATRIX(i, j))
    Next j
Next i

MATRIX_RANDOM_TRIANGULAR_FUNC = RANDOM_MATRIX

Exit Function
ERROR_LABEL:
MATRIX_RANDOM_TRIANGULAR_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : RANDOM_NORMAL_FUNC
'DESCRIPTION   : Return random numbers from a Normal Distribution
'LIBRARY       : STATISTICS
'GROUP         : RANDOM_PSEUDO
'ID            : 009
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 14/02/2008
'************************************************************************************
'************************************************************************************

Function RANDOM_NORMAL_FUNC(Optional ByVal MEAN_VAL As Double = 0, _
Optional ByVal SIGMA_VAL As Double = 1, _
Optional ByVal VERSION As Integer = 0)

Dim i As Long

Dim V1_VAL As Double
Dim V2_VAL As Double
Dim RAD_VAL As Double
Dim FACT_VAL As Double
Dim RANDOM_VAL As Double

Dim TEMP_SUM As Double

On Error GoTo ERROR_LABEL

'-------------------------------------------------------------------------
Select Case VERSION
'-------------------------------------------------------------------------
Case 0
'-------------------------------------------------------------------------
    RANDOM_VAL = PSEUDO_RANDOM_FUNC(1)
    RANDOM_NORMAL_FUNC = NORMSINV_FUNC(RANDOM_VAL, MEAN_VAL, SIGMA_VAL, 0)
'-------------------------------------------------------------------------
Case 1
'-------------------------------------------------------------------------
'Options, Futures and Other Derivatives: John C. Hull Sixth Edition.
    TEMP_SUM = 0
    For i = 1 To 12
        RANDOM_VAL = PSEUDO_RANDOM_FUNC(0)
        TEMP_SUM = TEMP_SUM + RANDOM_VAL
    Next i
    RANDOM_NORMAL_FUNC = (TEMP_SUM - 6) * SIGMA_VAL + MEAN_VAL
'-------------------------------------------------------------------------
Case Else
'-------------------------------------------------------------------------

'This case is mathematically an exact method, however, it may magnify
'the problem of a poor pseudo-random number generator, but at least the
'functions are implemented to machine accuracy with Excel.

1983: RANDOM_VAL = PSEUDO_RANDOM_FUNC(0)
    V1_VAL = 2 * RANDOM_VAL - 1
    V2_VAL = 2 * RANDOM_VAL - 1
    RAD_VAL = V1_VAL ^ 2 + V2_VAL ^ 2
    If RAD_VAL >= 1 Or RAD_VAL = 0 Then GoTo 1983
    FACT_VAL = Sqr(-2 * Log(RAD_VAL) / RAD_VAL)
    
    RANDOM_NORMAL_FUNC = V2_VAL * FACT_VAL * SIGMA_VAL + MEAN_VAL

'-------------------------------------------------------------------------
End Select
'-------------------------------------------------------------------------

Exit Function
ERROR_LABEL:
RANDOM_NORMAL_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : RANDOM_NORMAL_BETWEEN_FUNC
'DESCRIPTION   : Return random numbers from a Truncated Normal Distribution
'LIBRARY       : STATISTICS
'GROUP         : RANDOM_PSEUDO
'ID            : 010
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 14/02/2008
'************************************************************************************
'************************************************************************************

Function RANDOM_NORMAL_BETWEEN_FUNC(ByVal MEAN_VAL As Double, _
ByVal SIGMA_VAL As Double, _
Optional ByVal LEFT_BOUND As Double, _
Optional ByVal RIGHT_BOUND As Double, _
Optional ByVal RANDOM_TYPE As Integer = 0)

Dim TEMP_VAL As Double

On Error GoTo ERROR_LABEL

1983:    TEMP_VAL = RANDOM_NORMAL_FUNC(MEAN_VAL, SIGMA_VAL, RANDOM_TYPE)

If RIGHT_BOUND = 0 Then
    If TEMP_VAL < LEFT_BOUND Then GoTo 1983
Else
    If (RIGHT_BOUND < TEMP_VAL) Or (TEMP_VAL < LEFT_BOUND) Then GoTo 1983
End If

RANDOM_NORMAL_BETWEEN_FUNC = TEMP_VAL

Exit Function
ERROR_LABEL:
RANDOM_NORMAL_BETWEEN_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : MATRIX_RANDOM_NORMAL_FUNC
'DESCRIPTION   : Returns an array with normally distributed random numbers [0 1]
'LIBRARY       : STATISTICS
'GROUP         : RANDOM_PSEUDO
'ID            : 011
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 14/02/2008
'************************************************************************************
'************************************************************************************

Function MATRIX_RANDOM_NORMAL_FUNC(ByVal NROWS As Long, _
ByVal NCOLUMNS As Long, _
Optional ByVal NORM_TYPE As Integer = 0, _
Optional ByVal MEAN_VAL As Double = 0, _
Optional ByVal SIGMA_VAL As Double = 1, _
Optional ByVal VERSION As Integer = 0)

Dim i As Long
Dim j As Long

Dim TEMP_MATRIX As Variant

On Error GoTo ERROR_LABEL

ReDim TEMP_MATRIX(1 To NROWS, 1 To NCOLUMNS)

'-------------------------------------------------------------------
Select Case VERSION
'-------------------------------------------------------------------
Case 0
'-------------------------------------------------------------------
    For j = 1 To NCOLUMNS
        For i = 1 To NROWS
            TEMP_MATRIX(i, j) = RANDOM_NORMAL_FUNC(MEAN_VAL, SIGMA_VAL, NORM_TYPE)
        Next i
    Next j
'-------------------------------------------------------------------
Case Else 'REPEATING THE SAME RND NUMBER PER ROW
'-------------------------------------------------------------------
    For i = 1 To NROWS
        TEMP_MATRIX(i, 1) = RANDOM_NORMAL_FUNC(MEAN_VAL, SIGMA_VAL, NORM_TYPE)
        For j = 2 To NCOLUMNS
            TEMP_MATRIX(i, j) = TEMP_MATRIX(i, 1)
        Next j
    Next i
'-------------------------------------------------------------------
End Select
'-------------------------------------------------------------------
    
MATRIX_RANDOM_NORMAL_FUNC = TEMP_MATRIX

Exit Function
ERROR_LABEL:
MATRIX_RANDOM_NORMAL_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : RANDOM_NORMAL_FAST_FUNC

'DESCRIPTION   : Short and fast NormInv() equivalent function used to generate normally
' distributed random numbers.

' Note that for exact equivalence to NormInv, multiply return value by sigma
' and add in Mu
'
' http://www.wilmott.com/messageview.cfm?catid=10&threadid=38771

'LIBRARY       : STATISTICS
'GROUP         : RANDOM_PSEUDO
'ID            : 012
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 14/02/2008
'************************************************************************************
'************************************************************************************

Function RANDOM_NORMAL_FAST_FUNC(Optional ByVal VERSION As Long = 0)

Dim T_VAL As Double
Dim Q_VAL As Double

Dim X_VAL As Double
Dim P_VAL As Double

Dim C0_VAL As Double
Dim C1_VAL As Double
Dim C2_VAL As Double

Dim D1_VAL As Double
Dim D2_VAL As Double
Dim D3_VAL As Double

Dim S1_VAL As Double
Dim S2_VAL As Double

Dim X1_VAL As Double

Dim RAND1_VAL As Double
Dim RAND2_VAL As Double
 
On Error GoTo ERROR_LABEL

'------------------------------------------------------------------------------------
Select Case VERSION
'------------------------------------------------------------------------------------
Case 0
'------------------------------------------------------------------------------------
Q_VAL = Rnd

If (Q_VAL = 0.5) Then
     RANDOM_NORMAL_FAST_FUNC = 0
Else
    Q_VAL = 1# - Q_VAL
    If ((Q_VAL > 0) And (Q_VAL < 0.5)) Then
        P_VAL = Q_VAL
    Else
        If (Q_VAL = 1) Then
            P_VAL = 1 - 0.9999999 ' JPR - attempt to fix divide by zero below
        Else
            P_VAL = 1# - Q_VAL
        End If
    End If
    T_VAL = Sqr(Log(1# / (P_VAL * P_VAL)))
    C0_VAL = 2.515517
    C1_VAL = 0.802853
    C2_VAL = 0.010328
    D1_VAL = 1.432788
    D2_VAL = 0.189269
    D3_VAL = 0.001308
    X_VAL = T_VAL - (C0_VAL + C1_VAL * T_VAL + C2_VAL * (T_VAL * T_VAL)) / (1# + D1_VAL * T_VAL + D2_VAL * (T_VAL * T_VAL) + D3_VAL * (T_VAL ^ 3))
    If (Q_VAL > 0.5) Then: X_VAL = -1# * X_VAL
End If

RANDOM_NORMAL_FAST_FUNC = X_VAL
'----------------------------------------------------------------------------------------
Case Else
'----------------------------------------------------------------------------------------
' Another NormInv() like function
' Below is the BoxMuller-Algorithm which delivers a normally distributed random number
' It comes from a posting to the Wilmott.com forum by a member named fuez
' URL http://www.wilmott.com/messageview.cfm?catid=10&threadid=38771&forumid=1
1983:
    RAND1_VAL = 2 * Rnd - 1
    RAND2_VAL = 2 * Rnd - 1
    S1_VAL = RAND1_VAL ^ 2 + RAND2_VAL ^ 2
    If S1_VAL > 1 Then GoTo 1983
    S2_VAL = Sqr(-2 * Log(S1_VAL) / S1_VAL)
    X1_VAL = RAND1_VAL * S2_VAL
    'X2_VAL = RAND2_VAL * S2_VAL ' Not necessary to calculate.
    'return one or the other
    RANDOM_NORMAL_FAST_FUNC = X1_VAL
'---------------------------------------------------------------------------------------
End Select

Exit Function
ERROR_LABEL:
RANDOM_NORMAL_FAST_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : RANDOM_GAUSS_FUNC

'DESCRIPTION   : From Numerical Recipes in C, second edition, page 289
'returns a random number from Gaussian dis with mean and sigma specified

'LIBRARY       : STATISTICS
'GROUP         : RANDOM_PSEUDO
'ID            : 013
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 14/02/2008
'************************************************************************************
'************************************************************************************

Function RANDOM_GAUSS_FUNC(Optional ByVal MEAN_VAL As Double = 0, _
Optional ByVal SIGMA_VAL As Double = 1)

Static NEXT_RND As Double
Static RND_WAIT As Boolean
Static RANDOMIZE_FLAG As Boolean

Dim FAC_VAL As Double
Dim RSQ_VAL As Double
Dim V1_VAL As Double
Dim V2_VAL As Double
Dim STD_VAL As Double

On Error GoTo ERROR_LABEL

If Not (RANDOMIZE_FLAG) Then
   Randomize
   RANDOMIZE_FLAG = True
End If

If Not (RND_WAIT) Then
  Do
    V1_VAL = 2# * Rnd() - 1#
    V2_VAL = 2# * Rnd() - 1#
    RSQ_VAL = V1_VAL * V1_VAL + V2_VAL * V2_VAL
  Loop Until RSQ_VAL <= 1#
    FAC_VAL = Sqr(-2# * Log(RSQ_VAL) / RSQ_VAL) 'natural log
    NEXT_RND = V1_VAL * FAC_VAL
    RND_WAIT = True
    STD_VAL = V2_VAL * FAC_VAL
Else
    RND_WAIT = False
    STD_VAL = NEXT_RND
End If
'STD_VAL has mean zero and SD=1.
RANDOM_GAUSS_FUNC = (STD_VAL * SIGMA_VAL) + MEAN_VAL

Exit Function
ERROR_LABEL:
RANDOM_GAUSS_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : RANDOM_POLAR_MARSAGLIA_FUNC

'DESCRIPTION   : Generate normally distributed random numbers with zero mean and
' unit variance, N(0,1), using the Polar Marsaglia method

'LIBRARY       : STATISTICS
'GROUP         : RANDOM_PSEUDO
'ID            : 014
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 14/02/2008
'************************************************************************************
'************************************************************************************

Function RANDOM_POLAR_MARSAGLIA_FUNC() As Double

Dim Z1_VAL As Double
' uniformly distributed random number on (-1,1)
Dim Z2_VAL As Double
' uniformly distributed random number on (-1,1)
Dim TEMP_RAD As Double
' Z1_VAL^2 + Z2_VAL^2
Dim TEMP_VAL As Double
' temporary variable, to speed things up
Dim X1_VAL As Double
' the first N(0,1) normally distributed number

' Since the Polar Marsaglia method uses two U(-1,1) numbers to generate two N(0,1)
' random numbers, it is much more efficient to return the first N(0,1) number and
' save the second one. Then when the function is called again, it simply returns the
' one it saved last time, rather than computing two new N(0,1) random numbers.

' To implement this scheme, we have to

'  1.  Save the second N(0,1) random number so it is available the
'next time we call the function. Typically, variables inside a function
'in (most) programming languages are local and cease to exist once the
'function returns. In VBA, in order to preserve a variable so it doesn't
'cease to exist we have to declare it as Static

Static X2_VAL As Double
' the second N(0,1) normally distributed number, declared as
'a Static so it doesn't disappear once the function returns
'the first N(0,1)

'  2. Have some way of knowing whether or not we already have an N(0,1)
'  available or if we have to generate some new ones. Obviously this
'  variable has to be Static so it is preserved across function calls.
'  We choose a Boolean variable which is TRUE if an N(0,1) number is
'  already available (in which case we just return
'  the number) and FALSE if we have to generate two new N(0,1) numbers.
'  The code relies on the fact that when a function is first called in VBA all its
'  local variables are set to zero, and zero for a Boolean variable is equivalent to
'  FALSE.

Static AVAIL_FLAG As Boolean
' this flag that tells us whether an N(0,1) variable is already
' available (TRUE) or whether we have to compute two new
' ones (FALSE). When the function is called for the first time
' VBA sets this variable to FALSE. It is declared as Static so
' that its value is preserved across function calls.

On Error GoTo ERROR_LABEL

If (Not AVAIL_FLAG) Then  ' we have to generate two new N(0,1) numbers

  Do
  
    Z1_VAL = 2 * Rnd() - 1
    'generate two new U(-1,1) numbers
    Z2_VAL = 2 * Rnd() - 1
    TEMP_RAD = Z1_VAL * Z1_VAL + Z2_VAL * Z2_VAL
    'compute (the square of) their distance from the origin
  
  Loop While (TEMP_RAD > 1)
  'if they don't lie inside the unit circle in (Z1_VAL,Z2_VAL) space
  'then reject them and try again. If we were being ultra
  'cautious, we would also check that TEMP_RAD > 0.
                                           
  TEMP_VAL = Sqr(-2 * Log(TEMP_RAD) / TEMP_RAD)
  'the two new U(-1,1) numbers lie inside the unit circle
  X1_VAL = TEMP_VAL * Z1_VAL
  ' now generate the two new N(0,1) numbers
  X2_VAL = TEMP_VAL * Z2_VAL
  ' and save the second one in the Static variable X2_VAL
  
  AVAIL_FLAG = True
  'set the TEMP_FLAG to indicate that X2_VAL contains a new N(0,1) number
  RANDOM_POLAR_MARSAGLIA_FUNC = X1_VAL
  'and return X1_VAL
  
Else
  'already have X2_VAL ready, so no need to do anything except
  'set the TEMP_FLAG to indicate that next time we have to generate
  AVAIL_FLAG = False
  'two new N(0,1) random numbers
  RANDOM_POLAR_MARSAGLIA_FUNC = X2_VAL
  'and then return X2_VAL
  
End If

Exit Function
ERROR_LABEL:
RANDOM_POLAR_MARSAGLIA_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : VECTOR_RANDOM_BOX_MULLER_FUNC
'DESCRIPTION   : Box muller transformation array; returns an array of
'NSIZE normally distributed variables
'LIBRARY       : STATISTICS
'GROUP         : RANDOM_PSEUDO
'ID            : 015
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 14/02/2008
'************************************************************************************
'************************************************************************************

Function VECTOR_RANDOM_BOX_MULLER_FUNC(ByVal NSIZE As Long)
  
Dim ii As Long
Dim jj As Long
Dim NROWS As Long

Dim V1_VAL As Double
Dim V2_VAL As Double

Dim TEMP_VAL As Double
Dim TEMP_FACT As Double

Dim TEMP_ARR() As Variant

On Error GoTo ERROR_LABEL

ReDim TEMP_ARR(1 To NSIZE)

NROWS = FLOOR_FUNC(NSIZE / 2, 1)

jj = 0
For ii = 1 To NROWS
    Do
        V1_VAL = 2 * Rnd - 1
        V2_VAL = 2 * Rnd - 1
        TEMP_VAL = V1_VAL * V1_VAL + V2_VAL * V2_VAL
    Loop Until TEMP_VAL <= 1
    TEMP_FACT = Sqr(-2 * Log(TEMP_VAL) / TEMP_VAL)
    jj = jj + 1
    TEMP_ARR(jj) = V1_VAL * TEMP_FACT
    jj = jj + 1
    TEMP_ARR(jj) = V2_VAL * TEMP_FACT
Next ii

If (NSIZE > (NROWS * 2)) Then
    Do
        V1_VAL = 2 * Rnd - 1
        V2_VAL = 2 * Rnd - 1
        TEMP_VAL = V1_VAL * V1_VAL + V2_VAL * V2_VAL
    Loop Until TEMP_VAL <= 1
    TEMP_FACT = Sqr(-2 * Log(TEMP_VAL) / TEMP_VAL)
    jj = jj + 1
    TEMP_ARR(jj) = V2_VAL * TEMP_FACT
End If

VECTOR_RANDOM_BOX_MULLER_FUNC = TEMP_ARR
  
Exit Function
ERROR_LABEL:
VECTOR_RANDOM_BOX_MULLER_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : RANDOM_LOG_NORMAL_FUNC
'DESCRIPTION   : Return random numbers from a Log Normal Distribution
'LIBRARY       : STATISTICS
'GROUP         : RANDOM_PSEUDO
'ID            : 016
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 14/02/2008
'************************************************************************************
'************************************************************************************

Function RANDOM_LOG_NORMAL_FUNC(Optional ByVal MEAN_VAL As Double = 0, _
Optional ByVal SIGMA_VAL As Double = 1, _
Optional ByVal MU_SD_TYPE As Boolean = False, _
Optional ByVal VERSION As Integer = 0)

'The log-normal distribution is often assumed to be the distribution of
'a stock price.  A distribution is log-normally distributed when the
'natural log of the set of the random variables in that distribution
'is a normally distributed.  In plain English, if you take the natural
'log of each of the random numbers from a log-normal distribution, the
'new number set will be normally distribution.  Like the normal distribution,
'log-normal distribtuion is also defined with mean and standard deviation.

Dim RANDOM_VAL As Double
On Error GoTo ERROR_LABEL

RANDOM_VAL = PSEUDO_RANDOM_FUNC(1)
RANDOM_LOG_NORMAL_FUNC = INVERSE_LOGNORMDIST_FUNC(RANDOM_VAL, MEAN_VAL, SIGMA_VAL, MU_SD_TYPE, VERSION)
Exit Function
ERROR_LABEL:
RANDOM_LOG_NORMAL_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : RANDOM_LOG_PEARSON_FUNC
'DESCRIPTION   : Random Number Generator - Log Pearson Type III Distribution
'LIBRARY       : STATISTICS
'GROUP         : RANDOM_PSEUDO
'ID            : 017
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 14/02/2008
'************************************************************************************
'************************************************************************************

'The Log Pearson Type III distribution is commonly used in hydraulic
'studies. It is somehow similar to normal distribution, except instead
'of two parameters, stanand deviation and mean, it also has skew.  When
'the skew is small, Log Pearson Type III distribution approximates normal.

Function RANDOM_LOG_PEARSON_FUNC(ByVal RANDOM_VAL As Double, _
ByVal SKEW_VAL As Double, _
Optional ByVal MEAN_VAL As Double = 0, _
Optional ByVal SIGMA_VAL As Double = 1)

Dim TEMP_VAL As Double

On Error GoTo ERROR_LABEL

TEMP_VAL = (((SKEW_VAL / 6) * (RANDOM_VAL - SKEW_VAL / 6) + 1) ^ 3 - 1) * (2 / SKEW_VAL)
RANDOM_LOG_PEARSON_FUNC = TEMP_VAL * SIGMA_VAL + MEAN_VAL

Exit Function
ERROR_LABEL:
RANDOM_LOG_PEARSON_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : RANDOM_NUMBERS_SIMULATION_FUNC
'DESCRIPTION   :
'LIBRARY       : STATISTICS
'GROUP         : RANDOM_PSEUDO
'ID            : 018
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 16/07/2010
'************************************************************************************
'************************************************************************************

Function RANDOM_NUMBERS_SIMULATION_FUNC(Optional ByVal VERSION As Integer = 4, _
Optional ByVal MU_VAL As Double = 0, _
Optional ByVal SE_VAL As Double = 1, _
Optional ByVal POWER_VAL As Long = 1, _
Optional ByVal FACTOR_VAL As Double = 0.5, _
Optional ByVal RANDOM_FLAG As Boolean = True, _
Optional ByVal NO_PERIODS As Long = 10, _
Optional ByVal nLOOPS As Long = 1000, _
Optional ByVal OUTPUT As Integer = 2)

Dim i As Long
Dim j As Long

Dim RANDOM_VAL As Double

On Error GoTo ERROR_LABEL
If RANDOM_FLAG = True Then: Randomize
ReDim TEMP_MATRIX(1 To nLOOPS, 1 To NO_PERIODS)
'-----------------------------------------------------------------------------------------
Select Case VERSION
'-----------------------------------------------------------------------------------------
Case 0 ' Uniform
'-----------------------------------------------------------------------------------------
    For i = 1 To nLOOPS
        For j = 1 To NO_PERIODS
            RANDOM_VAL = PSEUDO_RANDOM_FUNC(1)
            TEMP_MATRIX(i, j) = SE_VAL * (NO_PERIODS) ^ 0.5 * (RANDOM_VAL - FACTOR_VAL) 'Factor --> Bound
        Next j
    Next i
'-----------------------------------------------------------------------------------------------
Case 1 'nonlinear, deterministic
'-----------------------------------------------------------------------------------------------
    i = 1
    For j = 1 To NO_PERIODS
        TEMP_MATRIX(i, j) = j ^ POWER_VAL
    Next j
    For i = 2 To nLOOPS
        For j = 1 To NO_PERIODS
            TEMP_MATRIX(i, j) = TEMP_MATRIX(i - 1, j)
        Next j
    Next i
'-----------------------------------------------------------------------------------------------
Case 2 'nonlinear, Random
'-----------------------------------------------------------------------------------------------
    For i = 1 To nLOOPS
        For j = 1 To NO_PERIODS
            RANDOM_VAL = PSEUDO_RANDOM_FUNC(1)
            TEMP_MATRIX(i, j) = (j ^ POWER_VAL) + NORMSINV_FUNC(RANDOM_VAL, MU_VAL, SE_VAL * FACTOR_VAL, 0) * -1 'Factor --> Influence
        Next j
    Next i
'-----------------------------------------------------------------------------------------
Case 3 'random homoscedastic (Normal)
'-----------------------------------------------------------------------------------------
    For i = 1 To nLOOPS
        For j = 1 To NO_PERIODS
            RANDOM_VAL = PSEUDO_RANDOM_FUNC(1)
            TEMP_MATRIX(i, j) = NORMSINV_FUNC(RANDOM_VAL, MU_VAL, SE_VAL, 0)
        Next j
    Next i
'-----------------------------------------------------------------------------------------------
Case 4 'random heteroscedastic
'-----------------------------------------------------------------------------------------------
    For i = 1 To nLOOPS
        For j = 1 To NO_PERIODS
            RANDOM_VAL = PSEUDO_RANDOM_FUNC(1)
            TEMP_MATRIX(i, j) = NORMSINV_FUNC(RANDOM_VAL, MU_VAL, SE_VAL, 0) * j ^ 0.5 ' this makes the data heteroscedastic
        Next j
    Next i
'-----------------------------------------------------------------------------------------
Case 5 'Log-Normal
'-----------------------------------------------------------------------------------------
    For i = 1 To nLOOPS
        For j = 1 To NO_PERIODS
            RANDOM_VAL = PSEUDO_RANDOM_FUNC(1)
            TEMP_MATRIX(i, j) = INVERSE_LOGNORMDIST_FUNC(RANDOM_VAL, MU_VAL, SE_VAL, False, 0)
        Next j
    Next i
'-----------------------------------------------------------------------------------------
Case Else 'T-Dist 'Probability associated with the two-tailed
'-----------------------------------------------------------------------------------------
    'Student's t-distribution.
    For i = 1 To nLOOPS
        For j = 1 To NO_PERIODS
            RANDOM_VAL = PSEUDO_RANDOM_FUNC(1)
            TEMP_MATRIX(i, j) = INVERSE_TDIST_FUNC(RANDOM_VAL / 2, FACTOR_VAL) 'Factor - Degrees of Freedom
        Next j
    Next i
'-----------------------------------------------------------------------------------------
End Select
'-----------------------------------------------------------------------------------------

Select Case OUTPUT
Case 0
    RANDOM_NUMBERS_SIMULATION_FUNC = TEMP_MATRIX
Case 1
    RANDOM_NUMBERS_SIMULATION_FUNC = DATA_BASIC_MOMENTS_FUNC(TEMP_MATRIX, 0, 0, 0.05, 1)
    'approximates the averages and SDs of different elements in the data series using Monte Carlo simulation
Case 2
    RANDOM_NUMBERS_SIMULATION_FUNC = MATRIX_CORRELATION_FUNC(TEMP_MATRIX)
Case Else
    RANDOM_NUMBERS_SIMULATION_FUNC = Array(TEMP_MATRIX, DATA_BASIC_MOMENTS_FUNC(TEMP_MATRIX, 0, 0, 0.05, 1), MATRIX_CORRELATION_FUNC(TEMP_MATRIX))
End Select

Exit Function
ERROR_LABEL:
RANDOM_NUMBERS_SIMULATION_FUNC = Err.number
End Function
