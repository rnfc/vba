Attribute VB_Name = "STAT_PROBABILITY_BAYER_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.

'************************************************************************************
'************************************************************************************
'FUNCTION      : BAYES_PROBABILITY_TABLE_FUNC
'DESCRIPTION   : BAYES PROBABILITY TABLE
'LIBRARY       : STATISTICS
'GROUP         : PROBABILITY_BAYER
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/22/2009
'************************************************************************************
'************************************************************************************

Function BAYES_PROBABILITY_TABLE_FUNC(ByRef PRIOR_PROB_RNG As Variant, _
ByRef COND_PROB_RNG As Variant)

'Probabilities --> New Info --> Application of Bayes' Theorem -->
'Posterior Prob....
'http://en.wikipedia.org/wiki/Bayes'_theorem

Dim i As Long
Dim NROWS As Long

Dim ATEMP_SUM As Double
Dim BTEMP_SUM As Double

Dim PRIOR_PROB_VECTOR As Variant 'THE SUM MUST BE EQUAL 100%
Dim COND_PROB_VECTOR As Variant

On Error GoTo ERROR_LABEL

PRIOR_PROB_VECTOR = PRIOR_PROB_RNG
If UBound(PRIOR_PROB_VECTOR, 1) = 1 Then
    PRIOR_PROB_VECTOR = MATRIX_TRANSPOSE_FUNC(PRIOR_PROB_VECTOR)
End If

COND_PROB_VECTOR = COND_PROB_RNG
If UBound(COND_PROB_VECTOR, 1) = 1 Then
    COND_PROB_VECTOR = MATRIX_TRANSPOSE_FUNC(COND_PROB_VECTOR)
End If

If UBound(PRIOR_PROB_VECTOR, 1) <> UBound(COND_PROB_VECTOR, 1) Then: GoTo ERROR_LABEL

NROWS = UBound(PRIOR_PROB_VECTOR, 1)

ReDim TEMP_MATRIX(0 To NROWS, 1 To 8)

TEMP_MATRIX(0, 1) = "P[Ai]" 'Prior Probabilities
TEMP_MATRIX(0, 2) = ""
    
TEMP_MATRIX(0, 3) = "P[X|Ai]" 'Conditional Probabilities
TEMP_MATRIX(0, 4) = "P[Y|Ai]"
    
TEMP_MATRIX(0, 5) = "P[Ai^X]" 'Probability of Outcome <Joint Probabilities>
TEMP_MATRIX(0, 6) = "P[Ai^Y]"
    
TEMP_MATRIX(0, 7) = "P[Ai|X]" 'Posterior Probabilities
TEMP_MATRIX(0, 8) = "P[Ai|Y]"

For i = 1 To NROWS

'-----------------------First Pass: Prior Probabilities (Up to NROWS)
    
    TEMP_MATRIX(i, 1) = "P[A" & CStr(i) & "]"
    TEMP_MATRIX(i, 2) = PRIOR_PROB_VECTOR(i, 1)
    
'-----------------------Second Pass: Conditional Probabilities (Up to NROWS)
    
    TEMP_MATRIX(i, 3) = COND_PROB_VECTOR(i, 1)
    TEMP_MATRIX(i, 4) = 1 - TEMP_MATRIX(i, 3)
    
'-----------------------Third Pass: Probability of Outcome <Joint
'Probabilities> (Up to NROWS)
    
    TEMP_MATRIX(i, 5) = TEMP_MATRIX(i, 2) * TEMP_MATRIX(i, 3)
    TEMP_MATRIX(i, 6) = TEMP_MATRIX(i, 2) * TEMP_MATRIX(i, 4)
    
    ATEMP_SUM = ATEMP_SUM + TEMP_MATRIX(i, 5)
    BTEMP_SUM = BTEMP_SUM + TEMP_MATRIX(i, 6)
    
Next i

'-----------------------Last Pass: Posterior Probabilities
For i = 1 To NROWS
    TEMP_MATRIX(i, 7) = TEMP_MATRIX(i, 5) / ATEMP_SUM
    TEMP_MATRIX(i, 8) = TEMP_MATRIX(i, 6) / BTEMP_SUM
Next i

BAYES_PROBABILITY_TABLE_FUNC = TEMP_MATRIX

Exit Function
ERROR_LABEL:
BAYES_PROBABILITY_TABLE_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : EXPECTED_VALUE_FUNC
'DESCRIPTION   : You flip a coin 25(tenor) times, getting Heads [and a 12.0%
'gain (UP_GAIN )] 70% (PROB_UP_GAIN) of the time - and a Loss of 10.7%
'(1 - PROB_UP_GAIN) when you get Tails. How does you portfolio [initially $100,000
'(X_VAL)] fare?

'LIBRARY       : STATISTICS
'GROUP         : BAYER
'ID            : 002
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/22/2009
'************************************************************************************
'************************************************************************************

Function EXPECTED_VALUE_FUNC(ByVal X_VAL As Double, _
ByVal UP_GAIN As Double, _
ByVal PROB_UP_GAIN As Double, _
ByVal TENOR As Long)

'X_VAL: Current Portfolio
'UP_GAIN: Estimated Annual UP Gain (g)
'PROB_UP_GAIN: Assumed Probability of an Up (PROB_UP_GAIN)

Dim i As Long

Dim U_VAL As Double
Dim D_VAL As Double
Dim TEMP_MATRIX As Variant
Dim RETURN_VAL As Double 'expected return

On Error GoTo ERROR_LABEL

U_VAL = (1 + UP_GAIN)
D_VAL = 1 / U_VAL
RETURN_VAL = PROB_UP_GAIN * U_VAL + (1 - PROB_UP_GAIN) * D_VAL - 1

ReDim TEMP_MATRIX(0 To TENOR, 1 To 4)

TEMP_MATRIX(0, 1) = 0

TEMP_MATRIX(0, 2) = X_VAL

TEMP_MATRIX(0, 3) = FACTORIAL_FUNC(TENOR) / FACTORIAL_FUNC(TENOR - 0) / FACTORIAL_FUNC(0) _
* (PROB_UP_GAIN) ^ 0 * ((1 - PROB_UP_GAIN)) ^ (TENOR - 0) 'Probability of
'getting i heads in in TENOR coin tosses (x = i, y = prob)

TEMP_MATRIX(0, 4) = X_VAL * U_VAL ^ 0 * D_VAL ^ (TENOR - 0)

For i = 1 To TENOR
    TEMP_MATRIX(i, 1) = i
    
    TEMP_MATRIX(i, 2) = TEMP_MATRIX(i - 1, 2) * (1 + RETURN_VAL)
    
    TEMP_MATRIX(i, 3) = FACTORIAL_FUNC(TENOR) / FACTORIAL_FUNC(TENOR - i) / FACTORIAL_FUNC(i) _
    * (PROB_UP_GAIN) ^ i * ((1 - PROB_UP_GAIN)) ^ (TENOR - i)
    'Probability of getting i heads in in TENOR coin tosses (x = i, y = prob)
    
    TEMP_MATRIX(i, 4) = X_VAL * U_VAL ^ i * D_VAL ^ (TENOR - i) 'Probability of
    'getting a particular portfolio after x(tenor) years
    '(x = $Value, y = Prob)
Next i

EXPECTED_VALUE_FUNC = TEMP_MATRIX

Exit Function
ERROR_LABEL:
EXPECTED_VALUE_FUNC = Err.number
End Function
