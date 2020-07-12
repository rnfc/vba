Attribute VB_Name = "FINAN_FI_HW_TREE_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.



'**********************************************************************************
'**********************************************************************************
'FUNCTION      : HW_INTEREST_RATE_TRINOMIAL_TREE_FUNC

'DESCRIPTION   : Function that returns trinomial interest rate tree as described in
'Hull, John C., Options, Futures & Other Derivatives. Fourth edition (2000).
'Prentice-Hall. p. 580ff.

'LIBRARY       : HULL-WHITE
'GROUP         : TREE
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 10-06-2008
'**********************************************************************************
'**********************************************************************************

Function HW_INTEREST_RATE_TRINOMIAL_TREE_FUNC(ByVal EXPIRATION As Double, _
ByVal STEPS As Long, _
ByVal VOLATILITY As Double, _
ByVal ALPHA As Double, _
ByRef XDATA_RNG As Variant, _
ByRef YDATA_RNG As Variant, _
Optional ByVal THRESHOLD_VAL As Double = 0.184, _
Optional ByVal epsilon As Double = 0.00001, _
Optional ByVal VERSION As Integer = 1, _
Optional ByVal OUTPUT As Integer = 2)
    
' ALPHA: Pullback factor of mean reversion process
' VOLATILITY: Volatiltiy of process
' VERSION: 0 --> log normal mode, else --> normal Hull & White mode

' THRESHOLD_VAL: factor to calculate the maximum number of steps in tree, i.e
'. level at which interest rates start branching down for high rates, respectively up
'for low rates. Hull recommends 0.184 which is also default value.
    
Dim l As Long
Dim i As Long
Dim j As Long
Dim k As Long

Dim SROW As Long
Dim NROWS As Long

Dim NSIZE As Long

Dim O_VAL As Double
Dim P_VAL As Double
Dim Q_VAL As Double

Dim DT_VAL As Double
Dim DR_VAL As Double
Dim DX_VAL As Double

Dim Z_VAL As Double
Dim WEIGHT_VAL As Double

Dim ATEMP_SUM As Double
Dim BTEMP_SUM As Double
Dim CTEMP_SUM As Double

Dim ALPHA_ARR As Variant

Dim RTEMP_MATRIX As Variant
Dim XTEMP_MATRIX As Variant
Dim PTEMP_MATRIX As Variant
Dim QTEMP_MATRIX As Variant

Dim XDATA_VECTOR As Variant
Dim YDATA_VECTOR As Variant

Dim tolerance As Double

On Error GoTo ERROR_LABEL

tolerance = 0.0001
XDATA_VECTOR = XDATA_RNG
If UBound(XDATA_VECTOR, 1) = 1 Then
    XDATA_VECTOR = MATRIX_TRANSPOSE_FUNC(XDATA_VECTOR)
End If
NSIZE = UBound(XDATA_VECTOR, 1)

YDATA_VECTOR = YDATA_RNG
If UBound(YDATA_VECTOR, 1) = 1 Then
    YDATA_VECTOR = MATRIX_TRANSPOSE_FUNC(YDATA_VECTOR)
End If
If NSIZE <> UBound(YDATA_VECTOR, 1) Then: GoTo ERROR_LABEL

DT_VAL = EXPIRATION / STEPS
DR_VAL = VOLATILITY * Sqr(3 * DT_VAL) 'i and j steps in tree
k = Int(THRESHOLD_VAL / (ALPHA * DT_VAL)) + 1

ReDim PTEMP_MATRIX(-k To k, 1 To 3)
'Calculate probabilities
PTEMP_MATRIX(-k, 1) = 1 / 6 + (ALPHA ^ 2 * k ^ 2 * DT_VAL ^ 2 + ALPHA * -k * DT_VAL) / 2
PTEMP_MATRIX(-k, 2) = -1 / 3 - ALPHA ^ 2 * k ^ 2 * DT_VAL ^ 2 - 2 * ALPHA * -k * DT_VAL
PTEMP_MATRIX(-k, 3) = 7 / 6 + (ALPHA ^ 2 * k ^ 2 * DT_VAL ^ 2 + 3 * ALPHA * -k * DT_VAL) / 2
For j = -k + 1 To 0
    PTEMP_MATRIX(j, 1) = 1 / 6 + (ALPHA ^ 2 * j ^ 2 * DT_VAL ^ 2 - ALPHA * j * DT_VAL) / 2
    PTEMP_MATRIX(j, 2) = 2 / 3 - ALPHA ^ 2 * j ^ 2 * DT_VAL ^ 2
    PTEMP_MATRIX(j, 3) = 1 / 6 + (ALPHA ^ 2 * j ^ 2 * DT_VAL ^ 2 + ALPHA * j * DT_VAL) / 2
Next j
For j = 1 To k
    PTEMP_MATRIX(j, 1) = PTEMP_MATRIX(-j, 3)
    PTEMP_MATRIX(j, 2) = PTEMP_MATRIX(-j, 2)
    PTEMP_MATRIX(j, 3) = PTEMP_MATRIX(-j, 1)
Next j

ReDim ALPHA_ARR(0 To STEPS)
ReDim QTEMP_MATRIX(0 To STEPS, -k To k)
ReDim RTEMP_MATRIX(0 To STEPS, -k To k)

If VERSION = 0 Then
    ReDim XTEMP_MATRIX(0 To STEPS, -k To k) 'Log normal tree
End If

'determine QTEMP_MATRIX values
'---------------------------------------------------------------------------------------
Select Case VERSION
'---------------------------------------------------------------------------------------
Case 0 'log
'---------------------------------------------------------------------------------------
    Z_VAL = DT_VAL
    GoSub 1983 'Interpolation
    RTEMP_MATRIX(0, 0) = WEIGHT_VAL
    ALPHA_ARR(0) = Log(RTEMP_MATRIX(0, 0))
    XTEMP_MATRIX(0, 0) = ALPHA_ARR(0)
    QTEMP_MATRIX(0, 0) = 1
    DX_VAL = DR_VAL
'---------------------------------------------------------------------------------------
Case Else 'normal
'---------------------------------------------------------------------------------------
    Z_VAL = DT_VAL
    GoSub 1983 'Interpolation
    ALPHA_ARR(0) = WEIGHT_VAL
    RTEMP_MATRIX(0, 0) = ALPHA_ARR(0)
    QTEMP_MATRIX(0, 0) = 1
'---------------------------------------------------------------------------------------
End Select
'---------------------------------------------------------------------------------------

For i = 0 To STEPS - 1
    SROW = MAXIMUM_FUNC(-k + 1, -i)
    NROWS = MINIMUM_FUNC(k - 1, i)
    For j = SROW To NROWS
        Select Case VERSION
        Case 0 'Log
          QTEMP_MATRIX(i + 1, j + 1) = PTEMP_MATRIX(j, 1) * _
            Exp(-Exp(ALPHA_ARR(i) + j * DX_VAL) * DT_VAL) * QTEMP_MATRIX(i, j) _
            + QTEMP_MATRIX(i + 1, j + 1)
          
          QTEMP_MATRIX(i + 1, j) = PTEMP_MATRIX(j, 2) * _
            Exp(-Exp(ALPHA_ARR(i) + j * DX_VAL) * DT_VAL) * QTEMP_MATRIX(i, j) _
            + QTEMP_MATRIX(i + 1, j)
          
          QTEMP_MATRIX(i + 1, j - 1) = PTEMP_MATRIX(j, 3) * _
            Exp(-Exp(ALPHA_ARR(i) + j * DX_VAL) * DT_VAL) * QTEMP_MATRIX(i, j) _
            + QTEMP_MATRIX(i + 1, j - 1)
        Case Else 'Normal
          QTEMP_MATRIX(i + 1, j + 1) = PTEMP_MATRIX(j, 1) * _
            Exp(-(ALPHA_ARR(i) + j * DR_VAL) * DT_VAL) * QTEMP_MATRIX(i, j) _
            + QTEMP_MATRIX(i + 1, j + 1)
          
          QTEMP_MATRIX(i + 1, j) = PTEMP_MATRIX(j, 2) * _
            Exp(-(ALPHA_ARR(i) + j * DR_VAL) * DT_VAL) * QTEMP_MATRIX(i, j) _
            + QTEMP_MATRIX(i + 1, j)
          
          QTEMP_MATRIX(i + 1, j - 1) = PTEMP_MATRIX(j, 3) * _
            Exp(-(ALPHA_ARR(i) + j * DR_VAL) * DT_VAL) * QTEMP_MATRIX(i, j) _
            + QTEMP_MATRIX(i + 1, j - 1)
        End Select
    Next j
    
    If i >= k Then
'---------------------------------------------------------------------------------------
        Select Case VERSION
'---------------------------------------------------------------------------------------
        Case 0 'Log
'---------------------------------------------------------------------------------------
      'upward branching tree
            QTEMP_MATRIX(i + 1, -k + 2) = PTEMP_MATRIX(-k, 1) * _
                Exp(-Exp(ALPHA_ARR(i) + -k * DX_VAL) * DT_VAL) * QTEMP_MATRIX(i, -k) _
                  + QTEMP_MATRIX(i + 1, -k + 2)
            
            QTEMP_MATRIX(i + 1, -k + 1) = PTEMP_MATRIX(-k, 2) * _
                Exp(-Exp(ALPHA_ARR(i) + -k * DX_VAL) * DT_VAL) * QTEMP_MATRIX(i, -k) _
                  + QTEMP_MATRIX(i + 1, -k + 1)
            
            QTEMP_MATRIX(i + 1, -k) = PTEMP_MATRIX(-k, 3) * _
                Exp(-Exp(ALPHA_ARR(i) + -k * DX_VAL) * DT_VAL) * QTEMP_MATRIX(i, -k) _
                  + QTEMP_MATRIX(i + 1, -k)
      'downward branching tree
            QTEMP_MATRIX(i + 1, k) = PTEMP_MATRIX(k, 1) * _
                Exp(-Exp(ALPHA_ARR(i) + k * DX_VAL) * DT_VAL) * QTEMP_MATRIX(i, k) _
                  + QTEMP_MATRIX(i + 1, k)
            
            QTEMP_MATRIX(i + 1, k - 1) = PTEMP_MATRIX(k, 2) * _
                Exp(-Exp(ALPHA_ARR(i) + k * DX_VAL) * DT_VAL) * QTEMP_MATRIX(i, k) _
                  + QTEMP_MATRIX(i + 1, k - 1)
            QTEMP_MATRIX(i + 1, k - 2) = PTEMP_MATRIX(k, 3) * _
                Exp(-Exp(ALPHA_ARR(i) + k * DX_VAL) * DT_VAL) * QTEMP_MATRIX(i, k) _
                  + QTEMP_MATRIX(i + 1, k - 2)
'---------------------------------------------------------------------------------------
        Case Else
'---------------------------------------------------------------------------------------
      'upward branching tree
            QTEMP_MATRIX(i + 1, -k + 2) = PTEMP_MATRIX(-k, 1) * _
                Exp(-(ALPHA_ARR(i) + -k * DR_VAL) * DT_VAL) * QTEMP_MATRIX(i, -k) _
                  + QTEMP_MATRIX(i + 1, -k + 2)
            
            QTEMP_MATRIX(i + 1, -k + 1) = PTEMP_MATRIX(-k, 2) * _
                Exp(-(ALPHA_ARR(i) + -k * DR_VAL) * DT_VAL) * QTEMP_MATRIX(i, -k) _
                  + QTEMP_MATRIX(i + 1, -k + 1)
            
            QTEMP_MATRIX(i + 1, -k) = PTEMP_MATRIX(-k, 3) * _
                Exp(-(ALPHA_ARR(i) + -k * DR_VAL) * DT_VAL) * QTEMP_MATRIX(i, -k) _
                  + QTEMP_MATRIX(i + 1, -k)
      'downward branching tree
            QTEMP_MATRIX(i + 1, k) = PTEMP_MATRIX(k, 1) * _
                Exp(-(ALPHA_ARR(i) + k * DR_VAL) * DT_VAL) * QTEMP_MATRIX(i, k) _
                  + QTEMP_MATRIX(i + 1, k)
            
            QTEMP_MATRIX(i + 1, k - 1) = PTEMP_MATRIX(k, 2) * _
                Exp(-(ALPHA_ARR(i) + k * DR_VAL) * DT_VAL) * QTEMP_MATRIX(i, k) _
                  + QTEMP_MATRIX(i + 1, k - 1)
            
            QTEMP_MATRIX(i + 1, k - 2) = PTEMP_MATRIX(k, 3) * _
                Exp(-(ALPHA_ARR(i) + k * DR_VAL) * DT_VAL) * QTEMP_MATRIX(i, k) _
                  + QTEMP_MATRIX(i + 1, k - 2)
'---------------------------------------------------------------------------------------
      End Select
'---------------------------------------------------------------------------------------
    End If

    ' Find ALPHA_ARR
'---------------------------------------------------------------------------------------
    Select Case VERSION
'---------------------------------------------------------------------------------------
    Case 0 'Log
'---------------------------------------------------------------------------------------
        'Find ALPHA_ARR(i+1) with Newton-Raphson
        ALPHA_ARR(i + 1) = ALPHA_ARR(i) 'set initial guess
        Z_VAL = (i + 2) * DT_VAL
        GoSub 1983 'Interpolation
        
        P_VAL = Exp(-(i + 2) * DT_VAL * WEIGHT_VAL)
        Do
            ATEMP_SUM = 0
            BTEMP_SUM = 0
            CTEMP_SUM = 0
            SROW = MAXIMUM_FUNC(-k, -i - 1)
            NROWS = MINIMUM_FUNC(k, i + 1)
            For j = SROW To NROWS
                ATEMP_SUM = ATEMP_SUM + QTEMP_MATRIX(i + 1, j) * _
                            Exp(-Exp(ALPHA_ARR(i + 1) + j * DX_VAL) * DT_VAL)
                
                BTEMP_SUM = BTEMP_SUM + QTEMP_MATRIX(i + 1, j) * _
                            Exp(-Exp(ALPHA_ARR(i + 1) + epsilon + j * DX_VAL) * DT_VAL)
                
                CTEMP_SUM = CTEMP_SUM + QTEMP_MATRIX(i + 1, j) * _
                            Exp(-Exp(ALPHA_ARR(i + 1) - epsilon + j * DX_VAL) * DT_VAL)
            Next j
            O_VAL = ATEMP_SUM - P_VAL
            Q_VAL = (BTEMP_SUM - CTEMP_SUM) / (2 * epsilon)
            ' determine value of slope as central difference
            ALPHA_ARR(i + 1) = ALPHA_ARR(i + 1) - O_VAL / Q_VAL
        Loop Until Abs(O_VAL) < tolerance
        
        'fill x, r arrays (ALPHA_ARR(i) corresponds to XTEMP_MATRIX(i))
        
        SROW = MAXIMUM_FUNC(-k, -i - 1)
        NROWS = MINIMUM_FUNC(k, i + 1)
        For j = SROW To NROWS
            XTEMP_MATRIX(i + 1, j) = DX_VAL * j + ALPHA_ARR(i + 1)
            RTEMP_MATRIX(i + 1, j) = Exp(DX_VAL * j + ALPHA_ARR(i + 1))
        Next j
'---------------------------------------------------------------------------------------
    Case Else
'---------------------------------------------------------------------------------------
        ATEMP_SUM = 0
        SROW = MAXIMUM_FUNC(-k, -i - 1)
        NROWS = MINIMUM_FUNC(k, i + 1)
        For j = SROW To NROWS
            ATEMP_SUM = ATEMP_SUM + QTEMP_MATRIX(i + 1, j) * Exp(-j * DR_VAL * DT_VAL)
        Next j
        
        Z_VAL = (i + 2) * DT_VAL
        GoSub 1983 'Interpolation
        
        ALPHA_ARR(i + 1) = (1 / DT_VAL) * Log(ATEMP_SUM / _
                              Exp(-(i + 2) * DT_VAL * WEIGHT_VAL))

        SROW = MAXIMUM_FUNC(-k, -i - 1)
        NROWS = MINIMUM_FUNC(k, i + 1)
        
        For j = SROW To NROWS
            RTEMP_MATRIX(i + 1, j) = DR_VAL * j + ALPHA_ARR(i + 1)
        Next j
'---------------------------------------------------------------------------------------
    End Select
'---------------------------------------------------------------------------------------
Next i

'------------------------------------------------------------------------------------
Select Case OUTPUT
'------------------------------------------------------------------------------------
Case 0
'------------------------------------------------------------------------------------
    Select Case VERSION
    Case 0 'Log(Rates)
        HW_INTEREST_RATE_TRINOMIAL_TREE_FUNC = MATRIX_TRANSPOSE_FUNC(XTEMP_MATRIX)
    Case Else 'Rates(R)
        HW_INTEREST_RATE_TRINOMIAL_TREE_FUNC = MATRIX_TRANSPOSE_FUNC(RTEMP_MATRIX)
    End Select
'------------------------------------------------------------------------------------
Case 1 'Rates(R)
'------------------------------------------------------------------------------------
    HW_INTEREST_RATE_TRINOMIAL_TREE_FUNC = MATRIX_TRANSPOSE_FUNC(RTEMP_MATRIX)
'------------------------------------------------------------------------------------
Case Else 'Probabilities
'------------------------------------------------------------------------------------
    HW_INTEREST_RATE_TRINOMIAL_TREE_FUNC = PTEMP_MATRIX
'------------------------------------------------------------------------------------
End Select
'------------------------------------------------------------------------------------

Exit Function
'-------------------------------------------------------
1983:
' determine a linear interpolated value
' x/y with x in ascending order required
'Z_VAL: x-value for which a Y-value should be interpolated
'-------------------------------------------------------
    WEIGHT_VAL = 0
    If Z_VAL <= XDATA_VECTOR(1, 1) Then
        WEIGHT_VAL = YDATA_VECTOR(1, 1) 'for smaller value set equal to smalles value in Y vector
        GoTo 1984
    Else
        l = 1
        Do While XDATA_VECTOR(l, 1) < Z_VAL ' Find out where Z_VAL is
            l = l + 1
            If l > NSIZE Then 'for larger values set equal to largest value in Y vector
                WEIGHT_VAL = YDATA_VECTOR(NSIZE, 1)
                GoTo 1984
            End If
        Loop
        WEIGHT_VAL = (XDATA_VECTOR(l, 1) - Z_VAL) / _
                     (XDATA_VECTOR(l, 1) - XDATA_VECTOR(l - 1, 1))
        
        WEIGHT_VAL = WEIGHT_VAL * YDATA_VECTOR(l - 1, 1) + _
                    (1 - WEIGHT_VAL) * YDATA_VECTOR(l, 1)
    End If
1984:
'-------------------------------------------------------
Return
'-------------------------------------------------------
ERROR_LABEL:
HW_INTEREST_RATE_TRINOMIAL_TREE_FUNC = Err.number
End Function


'**********************************************************************************
'**********************************************************************************
'FUNCTION      : HW_PROB_TABLE_FUNC
'DESCRIPTION   : HW PROBABILITIES ON ALL TREE BRANCHES
'LIBRARY       : HULL-WHITE
'GROUP         : TREE
'ID            : 002
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 10-06-2008
'**********************************************************************************
'**********************************************************************************

Private Function HW_PROB_TABLE_FUNC(ByVal TENOR As Double, _
ByVal STEPS As Double, _
ByVal KAPPA_VAL As Double, _
ByVal SIGMA As Double, _
Optional ByVal THRESHOLD_VAL As Double = 0.184)

'The parameter SIGMA determines the short rate's instantaneous
'standard deviation. The reversion rate parameter, KAPPA_VAL, determines
'the rate at which standard deviations decline with maturity. The
'higher KAPPA_VAL, the greater the decline.

'When it is = 0, the model reduces the Ho and Lee, and zero-coupon bond price
'volatilities are a linear function of maturity with the instantaneous standard
'deviations of both spot and forward rates being constant.

Dim i As Double
Dim j As Double

Dim MIN_VAL As Double
Dim MAX_VAL As Double

Dim DELTA_TENOR As Double
Dim DELTA_RATE As Double

Dim TEMP_MATRIX() As Variant

On Error GoTo ERROR_LABEL

DELTA_TENOR = TENOR / STEPS
DELTA_RATE = SIGMA * Sqr(3 * DELTA_TENOR)
MAX_VAL = Int(THRESHOLD_VAL / KAPPA_VAL / DELTA_TENOR + 1)
MIN_VAL = -1 * MAX_VAL

ReDim TEMP_MATRIX(1 To MAX_VAL * 2 + 1, 1 To 7)

j = j + 1
For i = (MAX_VAL) To (MIN_VAL) Step -1
    TEMP_MATRIX(j, 1) = i
    TEMP_MATRIX(j, 2) = i * DELTA_RATE
    TEMP_MATRIX(j, 3) = Exp(-i * DELTA_RATE)
    TEMP_MATRIX(j, 4) = HW_PROB_UP_FUNC(KAPPA_VAL, i, DELTA_TENOR, MAX_VAL)
    TEMP_MATRIX(j, 5) = HW_PROB_MID_FUNC(KAPPA_VAL, i, DELTA_TENOR, MAX_VAL)
    TEMP_MATRIX(j, 6) = HW_PROB_DOWN_FUNC(KAPPA_VAL, i, DELTA_TENOR, MAX_VAL)
    TEMP_MATRIX(j, 7) = TEMP_MATRIX(j, 6) + TEMP_MATRIX(j, 5) + TEMP_MATRIX(j, 4)
    
    j = j + 1
Next i

HW_PROB_TABLE_FUNC = TEMP_MATRIX

Exit Function
ERROR_LABEL:
HW_PROB_TABLE_FUNC = Err.number
End Function

'**********************************************************************************
'**********************************************************************************
'FUNCTION      : HW_NEXT_PROB_FUNC
'DESCRIPTION   : HW NEXT PROBABILY
'LIBRARY       : HULL-WHITE
'GROUP         : TREE
'ID            : 003
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 10-06-2008
'**********************************************************************************
'**********************************************************************************

Private Function HW_NEXT_PROB_FUNC(ByVal SIGMA As Double, _
ByVal KAPPA_VAL As Double, _
ByVal j As Double, _
ByVal STEP_VAL As Double, _
ByVal MAX_VAL As Double)

On Error GoTo ERROR_LABEL

If SIGMA = -1 Then
    HW_NEXT_PROB_FUNC = HW_PROB_UP_FUNC(KAPPA_VAL, j, STEP_VAL, MAX_VAL)
Else
    If SIGMA = 0 Then
        HW_NEXT_PROB_FUNC = HW_PROB_MID_FUNC(KAPPA_VAL, j, STEP_VAL, MAX_VAL)
    Else
        HW_NEXT_PROB_FUNC = HW_PROB_DOWN_FUNC(KAPPA_VAL, j, STEP_VAL, MAX_VAL)
    End If
End If

Exit Function
ERROR_LABEL:
HW_NEXT_PROB_FUNC = Err.number
End Function


'**********************************************************************************
'**********************************************************************************
'FUNCTION      : HW_PROB_UP_FUNC
'DESCRIPTION   : HW UP PROBABILY
'LIBRARY       : HULL-WHITE
'GROUP         : TREE
'ID            : 004
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 10-06-2008
'**********************************************************************************
'**********************************************************************************

Private Function HW_PROB_UP_FUNC(ByVal KAPPA_VAL As Double, _
ByVal j As Double, _
ByVal STEP_VAL As Double, _
ByVal MAX_VAL As Double)

Dim TEMP_VAL As Double
Dim FACTOR_VAL As Double

On Error GoTo ERROR_LABEL

FACTOR_VAL = KAPPA_VAL * j * STEP_VAL
If MAX_VAL = 0 Or (j < MAX_VAL And j > -MAX_VAL) Then
    TEMP_VAL = 1 / 6 + FACTOR_VAL * (FACTOR_VAL - 1) / 2
Else
    If j = MAX_VAL Then
        TEMP_VAL = 7 / 6 + FACTOR_VAL * (FACTOR_VAL - 3) / 2
    Else
        If j = -MAX_VAL Then
            TEMP_VAL = 1 / 6 + FACTOR_VAL * (FACTOR_VAL + 1) / 2
        End If
    End If
End If
HW_PROB_UP_FUNC = TEMP_VAL

Exit Function
ERROR_LABEL:
HW_PROB_UP_FUNC = Err.number
End Function


'**********************************************************************************
'**********************************************************************************
'FUNCTION      : HW_PROB_DOWN_FUNC
'DESCRIPTION   : HW DOWN PROBABILY
'LIBRARY       : HULL-WHITE
'GROUP         : TREE
'ID            : 005
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 10-06-2008
'**********************************************************************************
'**********************************************************************************

Private Function HW_PROB_DOWN_FUNC(ByVal KAPPA_VAL As Double, _
ByVal j As Double, _
ByVal STEP_VAL As Double, _
ByVal MAX_VAL As Double)

Dim TEMP_VAL As Double
Dim FACTOR_VAL As Double

On Error GoTo ERROR_LABEL

FACTOR_VAL = KAPPA_VAL * j * STEP_VAL
If MAX_VAL = 0 Or (j < MAX_VAL And j > -MAX_VAL) Then
    TEMP_VAL = 1 / 6 + FACTOR_VAL * (FACTOR_VAL + 1) / 2
Else
    If j = MAX_VAL Then
        TEMP_VAL = 1 / 6 + FACTOR_VAL * (FACTOR_VAL - 1) / 2
    Else
        If j = -MAX_VAL Then
            TEMP_VAL = 7 / 6 + FACTOR_VAL * (FACTOR_VAL + 3) / 2
        End If
    End If
End If
HW_PROB_DOWN_FUNC = TEMP_VAL

Exit Function
ERROR_LABEL:
HW_PROB_DOWN_FUNC = Err.number
End Function


'**********************************************************************************
'**********************************************************************************
'FUNCTION      : HW_PROB_MID_FUNC
'DESCRIPTION   : HW MID PROBABILY
'LIBRARY       : HULL-WHITE
'GROUP         : TREE
'ID            : 006
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 10-06-2008
'**********************************************************************************
'**********************************************************************************

Private Function HW_PROB_MID_FUNC(ByVal KAPPA_VAL As Double, _
ByVal j As Double, _
ByVal STEP_VAL As Double, _
ByVal MAX_VAL As Double)

Dim TEMP_VAL As Double
Dim FACTOR_VAL As Double

On Error GoTo ERROR_LABEL

FACTOR_VAL = KAPPA_VAL * j * STEP_VAL
If MAX_VAL = 0 Or (j < MAX_VAL And j > -MAX_VAL) Then
    TEMP_VAL = 2 / 3 - FACTOR_VAL ^ 2
Else
    If j = MAX_VAL Then
        TEMP_VAL = -1 / 3 + FACTOR_VAL * (-FACTOR_VAL + 2)
    Else
        If j = -MAX_VAL Then
            TEMP_VAL = -1 / 3 + FACTOR_VAL * (-FACTOR_VAL - 2)
        End If
    End If
End If
HW_PROB_MID_FUNC = TEMP_VAL

Exit Function
ERROR_LABEL:
HW_PROB_MID_FUNC = Err.number
End Function

