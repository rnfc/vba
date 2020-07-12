Attribute VB_Name = "FINAN_DERIV_MOMENTS_LIBR"


Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.

'************************************************************************************
'************************************************************************************
'FUNCTION      : OPTION_PRICES_FOUR_MOMENTS_FUNC

'DESCRIPTION   : This function shows, how one can estimate the descriptive
'statistics for a Risk neutral density directly from option prices using an
'approach similar to the VIX construction (where i use a somewhat different
'discretization), no interpolation of volatility or prices is needed.

'Parity Table & Centered 4 moments (e.g., mean, variance, skewness, kurtosis)

'LIBRARY       : DERIVATIVES
'GROUP         : MOMENTS
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/29/2009
'************************************************************************************
'************************************************************************************

Function OPTION_PRICES_FOUR_MOMENTS_FUNC(ByVal FORWARD As Double, _
ByVal EXPIRATION As Double, _
ByVal RATE As Double, _
ByRef STRIKE_RNG As Variant, _
ByRef CALL_RNG As Variant, _
ByRef PUT_RNG As Variant, _
Optional ByVal OUTPUT As Integer = 1)

Dim i As Long
Dim NROWS As Long

Dim SPOT As Double

Dim PUT_VECTOR As Variant
Dim CALL_VECTOR As Variant
Dim STRIKE_VECTOR As Variant

Dim PUT_DATA_VECTOR As Variant
Dim CALL_DATA_VECTOR As Variant
Dim STRIKE_DATA_VECTOR As Variant

Dim TEMP1_MATRIX As Variant
Dim TEMP2_MATRIX As Variant
Dim TEMP3_MATRIX As Variant

On Error GoTo ERROR_LABEL

STRIKE_DATA_VECTOR = STRIKE_RNG
If UBound(STRIKE_DATA_VECTOR, 1) = 1 Then
    STRIKE_DATA_VECTOR = MATRIX_TRANSPOSE_FUNC(STRIKE_DATA_VECTOR)
End If

CALL_DATA_VECTOR = CALL_RNG
If UBound(CALL_DATA_VECTOR, 1) = 1 Then
    CALL_DATA_VECTOR = MATRIX_TRANSPOSE_FUNC(CALL_DATA_VECTOR)
End If
If UBound(STRIKE_DATA_VECTOR, 1) <> UBound(CALL_DATA_VECTOR, 1) Then: GoTo ERROR_LABEL

PUT_DATA_VECTOR = PUT_RNG
If UBound(PUT_DATA_VECTOR, 1) = 1 Then
    PUT_DATA_VECTOR = MATRIX_TRANSPOSE_FUNC(PUT_DATA_VECTOR)
End If
If UBound(STRIKE_DATA_VECTOR, 1) <> UBound(PUT_DATA_VECTOR, 1) Then: GoTo ERROR_LABEL

NROWS = UBound(STRIKE_DATA_VECTOR, 1)

SPOT = FORWARD / Exp(RATE * EXPIRATION)

    ReDim TEMP1_MATRIX(0 To NROWS, 1 To 7)

    TEMP1_MATRIX(0, 1) = "NORM_STRIKE"
    TEMP1_MATRIX(0, 2) = "NORM_CALL"
    TEMP1_MATRIX(0, 3) = "NORM_PUT"

    'BSPut(1,k,1,0,v) - BSCall(1,k,1,0,v) = k - 1
    TEMP1_MATRIX(0, 4) = "PARIT_CALL"
    TEMP1_MATRIX(0, 5) = "PARIT_PUT"
    'simple arb. checks
    TEMP1_MATRIX(0, 6) = "FIRST_CHECK"
    TEMP1_MATRIX(0, 7) = "SECOND_CHECK"

    For i = NROWS To 1 Step -1
        TEMP1_MATRIX(i, 1) = STRIKE_DATA_VECTOR(i, 1) / FORWARD
        TEMP1_MATRIX(i, 2) = 1 / SPOT * CALL_DATA_VECTOR(i, 1)
        TEMP1_MATRIX(i, 3) = 1 / SPOT * PUT_DATA_VECTOR(i, 1)
        TEMP1_MATRIX(i, 4) = IIf(TEMP1_MATRIX(i, 1) <= 1, TEMP1_MATRIX(i, 2), TEMP1_MATRIX(i, 3) - TEMP1_MATRIX(i, 1) + 1)
        TEMP1_MATRIX(i, 5) = IIf(1 < TEMP1_MATRIX(i, 1), TEMP1_MATRIX(i, 3), Abs(TEMP1_MATRIX(i, 2) + TEMP1_MATRIX(i, 1) - 1))
        TEMP1_MATRIX(i, 6) = TEMP1_MATRIX(i, 4) - TEMP1_MATRIX(i, 5) + TEMP1_MATRIX(i, 1)
        If i <> NROWS Then
            TEMP1_MATRIX(i, 7) = IIf(TEMP1_MATRIX(i + 1, 4) < TEMP1_MATRIX(i, 4), 0, TEMP1_MATRIX(i, 4) - TEMP1_MATRIX(i + 1, 5)) + IIf(TEMP1_MATRIX(i, 5) < TEMP1_MATRIX(i + 1, 5), 0, TEMP1_MATRIX(i + 1, 5) - TEMP1_MATRIX(i, 5))
        Else
            TEMP1_MATRIX(i, 7) = ""
        End If
    Next i

'------------------------------------------------------------------------------------
Select Case OUTPUT
'------------------------------------------------------------------------------------
Case 0 'Parity Table
'------------------------------------------------------------------------------------
        OPTION_PRICES_FOUR_MOMENTS_FUNC = TEMP1_MATRIX
'------------------------------------------------------------------------------------
Case Else 'First 4 Moments
'------------------------------------------------------------------------------------

   ReDim STRIKE_VECTOR(1 To NROWS, 1 To 1)
   ReDim CALL_VECTOR(1 To NROWS, 1 To 1)
   ReDim PUT_VECTOR(1 To NROWS, 1 To 1)

   For i = 1 To NROWS
      STRIKE_VECTOR(i, 1) = TEMP1_MATRIX(i, 1)
      CALL_VECTOR(i, 1) = TEMP1_MATRIX(i, 4)
      PUT_VECTOR(i, 1) = TEMP1_MATRIX(i, 5)
   Next i

   TEMP2_MATRIX = OPTION_PRICES_MOMENTS_STATISTIC_FUNC(STRIKE_VECTOR, CALL_VECTOR, PUT_VECTOR)

   ReDim TEMP3_MATRIX(1 To 4, 1 To 2) ' translate to centered moments
       
   TEMP3_MATRIX(1, 1) = "MEAN"
   TEMP3_MATRIX(2, 1) = "VAR"
   TEMP3_MATRIX(3, 1) = "SKEW"
   TEMP3_MATRIX(4, 1) = "KURT"
   
   TEMP3_MATRIX(1, 2) = TEMP2_MATRIX(1, 1)
   TEMP3_MATRIX(2, 2) = TEMP2_MATRIX(2, 1)
   TEMP3_MATRIX(3, 2) = TEMP2_MATRIX(3, 1) / (Sqr(TEMP2_MATRIX(2, 1)) ^ 3)
   TEMP3_MATRIX(4, 2) = TEMP2_MATRIX(4, 1) / (TEMP2_MATRIX(2, 1) ^ 2)
   
   OPTION_PRICES_FOUR_MOMENTS_FUNC = TEMP3_MATRIX

'------------------------------------------------------------------------------------
End Select
'------------------------------------------------------------------------------------

Exit Function
ERROR_LABEL:
OPTION_PRICES_FOUR_MOMENTS_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : OPTION_PRICES_MOMENTS_STATISTIC_FUNC
'DESCRIPTION   : Four Moments Option Statistic
'LIBRARY       : DERIVATIVES
'GROUP         : MOMENTS
'ID            : 002
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/29/2009
'************************************************************************************
'************************************************************************************


Private Function OPTION_PRICES_MOMENTS_STATISTIC_FUNC(ByRef STRIKE_RNG As Variant, _
ByRef CALL_RNG As Variant, _
ByRef PUT_RNG As Variant)

Dim i As Long
Dim j As Long
Dim k As Long

Dim NSIZE As Long
Dim LOWER_NROWS As Long
Dim UPPER_NROWS As Long

Dim G1_VAL As Double
Dim G2_VAL As Double

Dim TEMP_K0 As Double
Dim TEMP_VAL As Double
Dim TEMP_SUM As Double

Dim CALL_VECTOR As Variant
Dim PUT_VECTOR As Variant
Dim STRIKE_VECTOR As Variant

Dim TEMP1_VECTOR As Variant
Dim TEMP2_VECTOR As Variant

Dim LOWER_PUT_VECTOR As Variant
Dim UPPER_CALL_VECTOR As Variant

Dim LOWER_STRIKE_VECTOR As Variant
Dim UPPER_STRIKE_VECTOR As Variant

Dim LOWER_PUT_WEIGHTS_VECTOR As Variant
Dim UPPER_CALL_WEIGHTS_VECTOR As Variant

On Error GoTo ERROR_LABEL

CALL_VECTOR = CALL_RNG
If UBound(CALL_VECTOR, 1) = 1 Then
    CALL_VECTOR = MATRIX_TRANSPOSE_FUNC(CALL_VECTOR)
End If

PUT_VECTOR = PUT_RNG
If UBound(PUT_VECTOR, 1) = 1 Then
    PUT_VECTOR = MATRIX_TRANSPOSE_FUNC(PUT_VECTOR)
End If

STRIKE_VECTOR = STRIKE_RNG
If UBound(STRIKE_VECTOR, 1) = 1 Then
    STRIKE_VECTOR = MATRIX_TRANSPOSE_FUNC(STRIKE_VECTOR)
End If
NSIZE = UBound(STRIKE_VECTOR, 1)

If (NSIZE = UBound(CALL_VECTOR, 1)) And (NSIZE = UBound(PUT_VECTOR, 1)) Then
Else
  Exit Function
End If

' TEMP_K0 = max(strike | strike <= 1)

For i = 1 To NSIZE
  If (STRIKE_VECTOR(i, 1) <= 1#) Then
    TEMP_K0 = STRIKE_VECTOR(i, 1)
  Else
    Exit For
  End If
Next i

' initialize
ReDim LOWER_STRIKE_VECTOR(1 To NSIZE, 1 To 1)
ReDim UPPER_STRIKE_VECTOR(1 To NSIZE, 1 To 1)

' separate strikes below and above TEMP_K0 by filling arrK_.....
' and set those arrays to correct lengths each

For i = 1 To NSIZE
  If (STRIKE_VECTOR(i, 1) <= TEMP_K0) Then
    LOWER_STRIKE_VECTOR(i, 1) = STRIKE_VECTOR(i, 1)
  Else
    Exit For
  End If
Next i
  
LOWER_NROWS = i - 1
j = 1
For i = 1 To NSIZE
  If (TEMP_K0 <= STRIKE_VECTOR(i, 1)) Then
    UPPER_STRIKE_VECTOR(j, 1) = STRIKE_VECTOR(i, 1)
    j = j + 1
  End If
Next i
  
UPPER_NROWS = j - 1

TEMP1_VECTOR = LOWER_STRIKE_VECTOR
ReDim LOWER_STRIKE_VECTOR(1 To LOWER_NROWS, 1 To 1)
For i = 1 To LOWER_NROWS
    LOWER_STRIKE_VECTOR(i, 1) = TEMP1_VECTOR(i, 1)
Next i

TEMP2_VECTOR = UPPER_STRIKE_VECTOR
ReDim UPPER_STRIKE_VECTOR(1 To UPPER_NROWS, 1 To 1) ' initialize arrays
For i = 1 To UPPER_NROWS
    UPPER_STRIKE_VECTOR(i, 1) = TEMP2_VECTOR(i, 1)
Next i

' initialize arrays
ReDim LOWER_PUT_VECTOR(1 To LOWER_NROWS, 1 To 1)
ReDim UPPER_CALL_VECTOR(1 To UPPER_NROWS, 1 To 1)

' provide put and call prices below and above TEMP_K0 by filling arr PC ...

For i = 1 To UBound(STRIKE_VECTOR, 1)
  If (STRIKE_VECTOR(i, 1) <= TEMP_K0) Then
    LOWER_PUT_VECTOR(i, 1) = PUT_VECTOR(i, 1)
  Else
    Exit For
  End If
Next i
  
j = 1
For i = 1 To UBound(STRIKE_VECTOR, 1)
  If (TEMP_K0 <= STRIKE_VECTOR(i, 1)) Then
    UPPER_CALL_VECTOR(j, 1) = CALL_VECTOR(i, 1)
    j = j + 1
  End If
Next i
    
TEMP1_VECTOR = LOWER_PUT_VECTOR
ReDim LOWER_PUT_VECTOR(1 To LOWER_NROWS, 1 To 1)
For i = 1 To LOWER_NROWS
    LOWER_PUT_VECTOR(i, 1) = TEMP1_VECTOR(i, 1)
Next i

TEMP2_VECTOR = UPPER_CALL_VECTOR
ReDim UPPER_CALL_VECTOR(1 To UPPER_NROWS, 1 To 1) ' initialize arrays
For i = 1 To UPPER_NROWS
    UPPER_CALL_VECTOR(i, 1) = TEMP2_VECTOR(i, 1)
Next i

' compute the first 4 un-centered moments
ReDim TEMP1_VECTOR(1 To 4, 1 To 1) ' un-centered moments
ReDim TEMP2_VECTOR(1 To 4, 1 To 1) ' centered moments

ReDim LOWER_PUT_WEIGHTS_VECTOR(1 To UBound(LOWER_STRIKE_VECTOR, 1), 1 To 1)
ReDim UPPER_CALL_WEIGHTS_VECTOR(1 To UBound(UPPER_STRIKE_VECTOR, 1), 1 To 1)

For k = 1 To 4
  
  ' weight the prices by TEMP_VAL(strike) by filling LOWER_PUT_WEIGHTS_VECTOR and
  'UPPER_CALL_WEIGHTS_VECTOR
'-------------------------------------------------------------------------------
    For i = 1 To UBound(LOWER_STRIKE_VECTOR, 1)

        If k = 1 Then TEMP_VAL = (-1 / (LOWER_STRIKE_VECTOR(i, 1) * _
                    LOWER_STRIKE_VECTOR(i, 1))): GoTo 1983
        If k = 2 Then TEMP_VAL = 2 * (-1# + (Log(LOWER_STRIKE_VECTOR(i, 1)))) * _
                    (-1 / (LOWER_STRIKE_VECTOR(i, 1) * LOWER_STRIKE_VECTOR(i, 1))): GoTo 1983
        If k = 3 Then TEMP_VAL = 3 * (Log(LOWER_STRIKE_VECTOR(i, 1))) * _
                    (-2# + (Log(LOWER_STRIKE_VECTOR(i, 1)))) * (-1 / _
                    (LOWER_STRIKE_VECTOR(i, 1) * LOWER_STRIKE_VECTOR(i, 1))): GoTo 1983
        If k = 4 Then TEMP_VAL = 4 * (Log(LOWER_STRIKE_VECTOR(i, 1))) * _
                    (Log(LOWER_STRIKE_VECTOR(i, 1))) * (-3# + (Log(LOWER_STRIKE_VECTOR(i, 1)))) _
                    * (-1 / (LOWER_STRIKE_VECTOR(i, 1) * LOWER_STRIKE_VECTOR(i, 1))): GoTo 1983

        TEMP_VAL = CDbl(k) * (1# - CDbl(k) + (Log(LOWER_STRIKE_VECTOR(i, 1)))) * _
                (Log(LOWER_STRIKE_VECTOR(i, 1))) ^ (k - 2)
1983:
  
      LOWER_PUT_WEIGHTS_VECTOR(i, 1) = LOWER_PUT_VECTOR(i, 1) * TEMP_VAL
    Next i
'-------------------------------------------------------------------------------

    For i = 1 To UBound(UPPER_STRIKE_VECTOR, 1)
        If k = 1 Then TEMP_VAL = (-1 / (UPPER_STRIKE_VECTOR(i, 1) * UPPER_STRIKE_VECTOR(i, 1))): GoTo 1984
        If k = 2 Then TEMP_VAL = 2 * (-1# + (Log(UPPER_STRIKE_VECTOR(i, 1)))) * _
                    (-1 / (UPPER_STRIKE_VECTOR(i, 1) * UPPER_STRIKE_VECTOR(i, 1))): GoTo 1984
        If k = 3 Then TEMP_VAL = 3 * (Log(UPPER_STRIKE_VECTOR(i, 1))) * _
                    (-2# + (Log(UPPER_STRIKE_VECTOR(i, 1)))) * (-1 / _
                    (UPPER_STRIKE_VECTOR(i, 1) * UPPER_STRIKE_VECTOR(i, 1))): GoTo 1984
        If k = 4 Then TEMP_VAL = 4 * (Log(UPPER_STRIKE_VECTOR(i, 1))) * _
                    (Log(UPPER_STRIKE_VECTOR(i, 1))) * (-3# + (Log(UPPER_STRIKE_VECTOR(i, 1)))) _
                    * (-1 / (UPPER_STRIKE_VECTOR(i, 1) * UPPER_STRIKE_VECTOR(i, 1))): GoTo 1984

        TEMP_VAL = CDbl(k) * (1# - CDbl(k) + (Log(UPPER_STRIKE_VECTOR(i, 1)))) * _
                (Log(UPPER_STRIKE_VECTOR(i, 1))) ^ (k - 2)
1984:

        UPPER_CALL_WEIGHTS_VECTOR(i, 1) = UPPER_CALL_VECTOR(i, 1) * TEMP_VAL
    Next i
'-------------------------------------------------------------------------------


    If k = 1 Then G1_VAL = (Log(TEMP_K0)): GoTo 1985
    If k = 2 Then G1_VAL = (Log(TEMP_K0)) * (Log(TEMP_K0)): GoTo 1985
    If k = 3 Then G1_VAL = (Log(TEMP_K0)) * (Log(TEMP_K0)) * _
                    (Log(TEMP_K0)): GoTo 1985
    If k = 4 Then G1_VAL = (Log(TEMP_K0)) * (Log(TEMP_K0)) * _
                    (Log(TEMP_K0)) * (Log(TEMP_K0)): GoTo 1985

    G1_VAL = (Log(TEMP_K0)) ^ k
1985:

'-------------------------------------------------------------------------------
    If k = 1 Then G2_VAL = 1# / TEMP_K0: GoTo 1986
    If k = 2 Then G2_VAL = 2# * (Log(TEMP_K0)) / _
TEMP_K0: GoTo 1986
    If k = 3 Then G2_VAL = 3# * (Log(TEMP_K0)) * _
                    (Log(TEMP_K0)) / TEMP_K0: GoTo 1986
    If k = 4 Then G2_VAL = 4# * (Log(TEMP_K0)) * _
                    (Log(TEMP_K0)) * (Log(TEMP_K0)) / TEMP_K0: GoTo 1986

    G2_VAL = CDbl(k) / TEMP_K0 * Log(TEMP_K0) ^ (k - 1)
'-------------------------------------------------------------------------------
1986:
'-------------------------------------------------------------------------------

  ' compute discrete integrals and correcting term
  TEMP_SUM = OPTION_PRICES_MOMENTS_INTEGRAND_FUNC(LOWER_STRIKE_VECTOR, LOWER_PUT_WEIGHTS_VECTOR)
  TEMP_SUM = TEMP_SUM + OPTION_PRICES_MOMENTS_INTEGRAND_FUNC(UPPER_STRIKE_VECTOR, _
                UPPER_CALL_WEIGHTS_VECTOR)
  TEMP1_VECTOR(k, 1) = TEMP_SUM + G1_VAL - (TEMP_K0 - 1#) * G2_VAL
Next k

' translate to centered moments
TEMP2_VECTOR(1, 1) = TEMP1_VECTOR(1, 1)
TEMP2_VECTOR(2, 1) = TEMP1_VECTOR(2, 1) - TEMP1_VECTOR(1, 1) ^ 2
TEMP2_VECTOR(3, 1) = TEMP1_VECTOR(3, 1) - 3# * TEMP1_VECTOR(1, 1) * _
                  TEMP1_VECTOR(2, 1) + 2# * TEMP1_VECTOR(1, 1) ^ 3
TEMP2_VECTOR(4, 1) = TEMP1_VECTOR(4, 1) - 4# * TEMP1_VECTOR(1, 1) * _
                  TEMP1_VECTOR(3, 1) + 6# * TEMP1_VECTOR(1, 1) ^ 2# * _
                  TEMP1_VECTOR(2, 1) - 3# * TEMP1_VECTOR(1, 1) ^ 4

OPTION_PRICES_MOMENTS_STATISTIC_FUNC = TEMP2_VECTOR

Exit Function
ERROR_LABEL:
OPTION_PRICES_MOMENTS_STATISTIC_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : OPTION_PRICES_MOMENTS_INTEGRAND_FUNC
'DESCRIPTION   : Moments Integration Function - using simpson Integral Function
'LIBRARY       : DERIVATIVES
'GROUP         : MOMENTS
'ID            : 003
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/29/2009
'************************************************************************************
'************************************************************************************

Private Function OPTION_PRICES_MOMENTS_INTEGRAND_FUNC(ByRef XDATA_RNG As Variant, _
ByRef YDATA_RNG As Variant)

Dim i As Long
Dim j As Long
Dim k As Long

Dim ii As Double
Dim jj As Double
Dim kk As Double
Dim ll As Double
Dim mm As Double

Dim X1_VAL As Double
Dim X2_VAL As Double
Dim X3_VAL As Double
Dim Y1_VAL As Double
Dim Y2_VAL As Double
Dim Y3_VAL As Double

Dim TEMP_SUM As Double
Dim FINAL_VAL As Double

Dim XTEMP_VECTOR As Variant
Dim YTEMP_VECTOR As Variant

On Error GoTo ERROR_LABEL

XTEMP_VECTOR = XDATA_RNG
If UBound(XTEMP_VECTOR, 1) = 1 Then
    XTEMP_VECTOR = MATRIX_TRANSPOSE_FUNC(XTEMP_VECTOR)
End If

YTEMP_VECTOR = YDATA_RNG
If UBound(YTEMP_VECTOR, 1) = 1 Then
    YTEMP_VECTOR = MATRIX_TRANSPOSE_FUNC(YTEMP_VECTOR)
End If

j = UBound(XTEMP_VECTOR, 1)

If (j = 1) Then
  OPTION_PRICES_MOMENTS_INTEGRAND_FUNC = 0#
  Exit Function
End If
If (j = 2) Then
  OPTION_PRICES_MOMENTS_INTEGRAND_FUNC = (YTEMP_VECTOR(1, 1) + YTEMP_VECTOR(2, 1)) / (XTEMP_VECTOR(2, 1) - XTEMP_VECTOR(1, 1)) / 2#
  Exit Function
End If

If (j Mod 2 <> 1) Then
  i = j - 1
Else
  i = j
End If

TEMP_SUM = 0#
For k = 1 To i - 2 Step 2
  ii = XTEMP_VECTOR(k + 1, 1) - XTEMP_VECTOR(k, 1)
  kk = XTEMP_VECTOR(k + 2, 1) - XTEMP_VECTOR(k + 1, 1)
  jj = YTEMP_VECTOR(k, 1)
  ll = YTEMP_VECTOR(k + 1, 1)
  mm = YTEMP_VECTOR(k + 2, 1)
  TEMP_SUM = TEMP_SUM + (kk / 3# - 1# / kk * ii * ii / 6# + ii / 6#) * mm + (-kk * kk / ii / 6# + kk / 6# + ii / 3#) * jj + (kk * kk / ii / 6# + kk / 2# + 1# / kk * ii * ii / 6# + ii / 2#) * ll
Next k

If (j = i + 1) Then
  
  X1_VAL = XTEMP_VECTOR(j - 2, 1)
  X2_VAL = XTEMP_VECTOR(j - 1, 1)
  X3_VAL = XTEMP_VECTOR(j, 1)
  Y1_VAL = YTEMP_VECTOR(j - 2, 1)
  Y2_VAL = YTEMP_VECTOR(j - 1, 1)
  Y3_VAL = YTEMP_VECTOR(j, 1)
  
  FINAL_VAL = (X2_VAL * X2_VAL - 2# * X3_VAL * X2_VAL + X3_VAL * X3_VAL) * Y1_VAL
  FINAL_VAL = (-X3_VAL * X3_VAL + (4# * X3_VAL + 2# * X2_VAL) * X1_VAL - 2# * X3_VAL * X2_VAL - 3# * X1_VAL * X1_VAL) * Y2_VAL + FINAL_VAL
  FINAL_VAL = (-X2_VAL * X2_VAL + (2# * X3_VAL + 4# * X2_VAL) * X1_VAL - 2# * X3_VAL * X2_VAL - 3# * X1_VAL * X1_VAL) * Y3_VAL + FINAL_VAL
  FINAL_VAL = (X2_VAL - X3_VAL) * FINAL_VAL / (-X2_VAL + X1_VAL) / (X1_VAL - X3_VAL) / 6#
  TEMP_SUM = TEMP_SUM + FINAL_VAL
End If

OPTION_PRICES_MOMENTS_INTEGRAND_FUNC = TEMP_SUM

Exit Function
ERROR_LABEL:
OPTION_PRICES_MOMENTS_INTEGRAND_FUNC = Err.number
End Function
