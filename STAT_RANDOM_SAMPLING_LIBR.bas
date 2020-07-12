Attribute VB_Name = "STAT_RANDOM_SAMPLING_LIBR"


Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.


'************************************************************************************
'************************************************************************************
'FUNCTION      : RANDOM_SAMPLING_CHOOSER_FUNC
'DESCRIPTION   : Choser for different types of Random Sampling
'LIBRARY       : STATISTICS
'GROUP         : RANDOM_SAMPLING
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 14/02/2008
'************************************************************************************
'************************************************************************************

Function RANDOM_SAMPLING_CHOOSER_FUNC(ByRef DATA_RNG As Variant, _
ByVal nDRAWS As Long, _
Optional ByVal VERSION As Integer = 1, _
Optional ByVal DRAWS_ONCE As Integer = 0, _
Optional ByVal RANDOM_FLAG As Boolean = True)

On Error GoTo ERROR_LABEL

Select Case VERSION
Case 0
    RANDOM_SAMPLING_CHOOSER_FUNC = RANDOM_SAMPLING_WITHOUT_REPLACEMENT(DATA_RNG, nDRAWS, RANDOM_FLAG)
    'WITHOUT REPLACEMENT
Case 1
    RANDOM_SAMPLING_CHOOSER_FUNC = RANDOM_SAMPLING_WITH_REPLACEMENT(DATA_RNG, nDRAWS, RANDOM_FLAG)
    'WITH REPLACEMENT
Case Else
    RANDOM_SAMPLING_CHOOSER_FUNC = RANDOM_SAMPLING_CONSECUTIVE_FUNC(DATA_RNG, nDRAWS, DRAWS_ONCE, RANDOM_FLAG)
    'CONSECUTIVE DRAWS
End Select

Exit Function
ERROR_LABEL:
RANDOM_SAMPLING_CHOOSER_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : RANDOM_SAMPLING_WITH_REPLACEMENT
'DESCRIPTION   : DRAWS WITH REPLACEMENT
'LIBRARY       : STATISTICS
'GROUP         : RANDOM_SAMPLING
'ID            : 002
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 14/02/2008
'************************************************************************************
'************************************************************************************

Function RANDOM_SAMPLING_WITH_REPLACEMENT(ByRef DATA_RNG As Variant, _
ByVal nDRAWS As Long, _
Optional ByVal RANDOM_FLAG As Boolean = True)

Dim i As Long
Dim j As Long
Dim k As Long ' Random Row

Dim SROW As Long
Dim NROWS As Long

Dim NSIZE As Long
Dim NO_VAR As Long

Dim RANDOM_VAL As Double

Dim DATA_MATRIX As Variant
Dim TEMP_MATRIX As Variant

On Error GoTo ERROR_LABEL

If nDRAWS = 0 Then: GoTo ERROR_LABEL

DATA_MATRIX = DATA_RNG

SROW = LBound(DATA_MATRIX, 1)
NROWS = UBound(DATA_MATRIX, 1)

NSIZE = (NROWS - SROW + 1)

NO_VAR = UBound(DATA_MATRIX, 2)

If nDRAWS > 1000000 Then
    RANDOM_SAMPLING_WITH_REPLACEMENT = _
    "The number of draws with replacement cannot be greater than " _
        & CStr(Format(1000000, 0, 0#)) & " (DATA POINTS)."
    Exit Function
End If

ReDim TEMP_MATRIX(1 To nDRAWS, 1 To NO_VAR)
If RANDOM_FLAG = True Then: Randomize

For i = 1 To nDRAWS
    RANDOM_VAL = Rnd
    k = Int(NSIZE * RANDOM_VAL + 1)
    For j = 1 To NO_VAR
        TEMP_MATRIX(i, j) = DATA_MATRIX(k, j)
    Next j
Next i

RANDOM_SAMPLING_WITH_REPLACEMENT = TEMP_MATRIX

Exit Function
ERROR_LABEL:
RANDOM_SAMPLING_WITH_REPLACEMENT = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : RANDOM_SAMPLING_WITHOUT_REPLACEMENT
'DESCRIPTION   : DRAWS WITHOUT REPLACEMENT
'LIBRARY       : STATISTICS
'GROUP         : RANDOM_SAMPLING
'ID            : 003
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 14/02/2008
'************************************************************************************
'************************************************************************************

Function RANDOM_SAMPLING_WITHOUT_REPLACEMENT(ByRef DATA_RNG As Variant, _
ByVal nDRAWS As Long, _
Optional ByVal RANDOM_FLAG As Boolean = True)

Dim i As Long
Dim j As Long
Dim k As Long ' Random Row

Dim NO_VAR As Long
Dim NSIZE As Long

Dim SROW As Long
Dim NROWS As Long

Dim BUMP_VAL As Variant
Dim CHOSE_VAL As Variant
Dim RANDOM_VAL As Double

Dim DATA_MATRIX As Variant
Dim TEMP_MATRIX As Variant

On Error GoTo ERROR_LABEL

If nDRAWS = 0 Then: GoTo ERROR_LABEL

DATA_MATRIX = DATA_RNG

NROWS = UBound(DATA_MATRIX, 1)
SROW = LBound(DATA_MATRIX, 1)
NSIZE = NROWS - SROW + 1
NO_VAR = UBound(DATA_MATRIX, 2)

If nDRAWS > NSIZE Then 'nDRAWS = NSIZE
    RANDOM_SAMPLING_WITHOUT_REPLACEMENT = "The number of draws without replacement " & _
    "cannot be greater than " & CStr(Format(NSIZE, 0, 0)) _
        & " (number of DATA POINTS)."
    Exit Function
End If

If RANDOM_FLAG = True Then: Randomize

ReDim TEMP_MATRIX(1 To nDRAWS, 1 To NO_VAR)

For i = 1 To nDRAWS
    RANDOM_VAL = Rnd
    k = Int((NROWS - SROW + 1) * RANDOM_VAL + SROW)
    For j = 1 To NO_VAR
        BUMP_VAL = DATA_MATRIX(i, j)
        CHOSE_VAL = DATA_MATRIX(k, j)
        DATA_MATRIX(i, j) = CHOSE_VAL
        DATA_MATRIX(k, j) = BUMP_VAL
        TEMP_MATRIX(i, j) = DATA_MATRIX(i, j) ' Put the sample into results vector
    Next j
    SROW = SROW + 1
Next i

RANDOM_SAMPLING_WITHOUT_REPLACEMENT = TEMP_MATRIX

Exit Function
ERROR_LABEL:
RANDOM_SAMPLING_WITHOUT_REPLACEMENT = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : RANDOM_SAMPLING_CONSECUTIVE_FUNC
'DESCRIPTION   : CONSECUTIVE DRAWS
'LIBRARY       : STATISTICS
'GROUP         : RANDOM_SAMPLING
'ID            : 004
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 14/02/2008
'************************************************************************************
'************************************************************************************

Function RANDOM_SAMPLING_CONSECUTIVE_FUNC(ByRef DATA_RNG As Variant, _
ByVal nDRAWS As Long, _
Optional ByVal DRAWS_ONCE As Long = 0, _
Optional ByVal RANDOM_FLAG As Boolean = True)

' This function draws a certain number of ENTRIES in a row.
' E.g. if the k is 1 and DRAWS_ONCE = 3, then we draw X, Y and h
' When there is overlap we go back to the beginning.

'Enter the number of consecutive draws: from 2 to nDRAWS. Your choice must
'perfectly divide (without a remainder) the number nDRAWS
'[e.g. with 10 draws, 2, 5, and 10 are allowed]

' DRAWS_ONCE: how many you draw at once
' m: No Passes
' nDRAWS = m * DRAWS_ONCE

Dim h As Long
Dim i As Long
Dim j As Long
Dim k As Long ' Random Row
Dim l As Long
Dim m As Long

Dim SROW As Long
Dim NROWS As Long

Dim NSIZE As Long
Dim NO_VAR As Long

Dim RANDOM_VAL As Double

Dim TEMP_MATRIX As Variant
Dim DATA_MATRIX As Variant

On Error GoTo ERROR_LABEL

If DRAWS_ONCE = 0 Then: GoTo ERROR_LABEL

DATA_MATRIX = DATA_RNG

SROW = LBound(DATA_MATRIX, 1)
NROWS = UBound(DATA_MATRIX, 1)

NSIZE = (NROWS - SROW + 1)

NO_VAR = UBound(DATA_MATRIX, 2)

If (nDRAWS > NSIZE) Then 'nDRAWS = NSIZE
    RANDOM_SAMPLING_CONSECUTIVE_FUNC = "The number of draws cannot be greater than " _
        & CStr(Format(NSIZE, 0, 0)) & " (number of DATA POINTS)."
    Exit Function
End If

If ((DRAWS_ONCE > nDRAWS) Or (DRAWS_ONCE < 2) Or (nDRAWS Mod DRAWS_ONCE <> 0) _
Or (IsNumeric(DRAWS_ONCE) = False)) Then
    RANDOM_SAMPLING_CONSECUTIVE_FUNC = "You did not enter a number greater than 2 " & _
        " that perfectly divides the number of draws. Please try again."
    Exit Function
End If

m = Int(nDRAWS / DRAWS_ONCE)

ReDim TEMP_MATRIX(1 To NSIZE, 1 To NO_VAR)

If RANDOM_FLAG = True Then: Randomize

For i = 1 To m
        
        h = (i - 1) * DRAWS_ONCE
        RANDOM_VAL = Rnd
        k = Int(NSIZE * RANDOM_VAL + 1)

'--------------------------------------Check for overlaps

        If k + DRAWS_ONCE - 1 > NSIZE Then
            For l = k To NSIZE
                h = h + 1
                For j = 1 To NO_VAR
                    TEMP_MATRIX(h, j) = DATA_MATRIX(l, j)
                Next j
            Next l
            
            For l = 1 To DRAWS_ONCE - NSIZE + k - 1
                h = h + 1
                For j = 1 To NO_VAR
                    TEMP_MATRIX(h, j) = DATA_MATRIX(l, j)
                Next j
            Next l
        Else
            For l = 1 To DRAWS_ONCE
                h = h + 1
                For j = 1 To NO_VAR
                    TEMP_MATRIX(h, j) = DATA_MATRIX(k + l - 1, j)
                Next j
            Next l
        End If
Next i


RANDOM_SAMPLING_CONSECUTIVE_FUNC = MATRIX_TRIM_FUNC(TEMP_MATRIX, 1, 0)

Exit Function
ERROR_LABEL:
RANDOM_SAMPLING_CONSECUTIVE_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : RANDOM_SAMPLING_SE_ADJUSTMENT_FUNC
'DESCRIPTION   : STANDARD_ERROR_FUNCTION
'LIBRARY       : STATISTICS
'GROUP         : RANDOM_SAMPLING
'ID            : 005
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 14/02/2008
'************************************************************************************
'************************************************************************************


Function RANDOM_SAMPLING_SE_ADJUSTMENT_FUNC(ByVal SIGMA_VAL As Double, _
ByVal nDRAWS As Long, _
ByVal nLOOPS As Long, _
Optional ByVal OUTPUT As Integer = 0)

'1) Hooke 's Law: A worse measuring device leads to a less precise estimate.
'As the spread of the errors in the box increases, so does the SE of the
'sample slope.

'2) Hooke 's Law: More spread in the weights leads to a more precise estimate.
'As the spread of the indep. variable increases, the SE of the sample
'slope decreases.

'3) Hooke 's Law: More observations leads to a more precise estimate.
'In other words, as the number of draws increases, the SE of the
'sample slope decreases.


On Error GoTo ERROR_LABEL

Select Case OUTPUT
Case 0 'SE WITH REPLACEMENT
    RANDOM_SAMPLING_SE_ADJUSTMENT_FUNC = SIGMA_VAL / Sqr(nDRAWS)
Case Else 'SE WITHOUT REPLACEMENT
    'CORRECT_FACTOR = Sqr((nLOOPS - (nDRAWS)) / nLOOPS)
    RANDOM_SAMPLING_SE_ADJUSTMENT_FUNC = (SIGMA_VAL / Sqr(nDRAWS)) * (Sqr((nLOOPS - (nDRAWS)) / nLOOPS))
End Select

Exit Function
ERROR_LABEL:
RANDOM_SAMPLING_SE_ADJUSTMENT_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : RANDOM_SAMPLING_NORMAL_ADJUSTMENT_FUNC
'DESCRIPTION   : Area from Z_SCORE to positive infinity: Using the Standard
'Formula for the Standard Error
'LIBRARY       : STATISTICS
'GROUP         : RANDOM_SAMPLING
'ID            : 006
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 14/02/2008
'************************************************************************************
'************************************************************************************

Function RANDOM_SAMPLING_NORMAL_ADJUSTMENT_FUNC(ByVal X_VAL As Double, _
ByVal MEAN_VAL As Double, _
ByVal SIGMA_VAL As Double, _
ByVal nLOOPS As Long, _
ByVal nDRAWS As Long, _
Optional ByVal VERSION As Integer = 1)

Dim Z_SCORE As Double

On Error GoTo ERROR_LABEL

Z_SCORE = (X_VAL - MEAN_VAL) / RANDOM_SAMPLING_SE_ADJUSTMENT_FUNC(SIGMA_VAL, nDRAWS, nLOOPS, VERSION)
RANDOM_SAMPLING_NORMAL_ADJUSTMENT_FUNC = 1 - NORMSDIST_FUNC(Z_SCORE, 0, 1, 0)

Exit Function
ERROR_LABEL:
RANDOM_SAMPLING_NORMAL_ADJUSTMENT_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : RANDOM_SAMPLING_SIMULATION_FUNC
'DESCRIPTION   : Run a monte carlo simulation
'LIBRARY       : STATISTICS
'GROUP         : RANDOM_SAMPLING
'ID            : 007
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 14/02/2008
'************************************************************************************
'************************************************************************************

Function RANDOM_SAMPLING_SIMULATION_FUNC(ByRef DATA_RNG As Variant, _
ByVal nLOOPS As Long, _
ByVal nDRAWS As Long, _
ByVal BOUND_VALUE As Variant, _
Optional ByVal BOUND_TYPE As Integer = 0, _
Optional ByVal VERSION As Integer = 0, _
Optional ByVal DRAWS_ONCE As Integer = 0, _
Optional ByVal RANDOM_FLAG As Boolean = True)

Dim h As Long
Dim i As Long
Dim j As Long
Dim l As Long

Dim NO_VAR As Long 'No. Variables in the array

Dim DATA_MATRIX As Variant
Dim TEMP_MATRIX As Variant

Dim ATEMP_VECTOR As Variant
Dim BTEMP_VECTOR As Variant

On Error GoTo ERROR_LABEL

DATA_MATRIX = DATA_RNG
NO_VAR = UBound(DATA_MATRIX, 2)

ReDim TEMP_MATRIX(1 To nLOOPS, 1 To 1)
ReDim BTEMP_VECTOR(1 To 1, 1 To NO_VAR)

For h = 1 To NO_VAR
    If NO_VAR <> 1 Then: ATEMP_VECTOR = MATRIX_GET_COLUMN_FUNC(DATA_MATRIX, h, 1)
    For j = 1 To nLOOPS 'Repetitions
        ATEMP_VECTOR = RANDOM_SAMPLING_CHOOSER_FUNC(ATEMP_VECTOR, nDRAWS, VERSION, DRAWS_ONCE, RANDOM_FLAG)
        l = 0
        For i = 1 To nDRAWS 'Entries
            Select Case BOUND_TYPE
            Case 0
                If ATEMP_VECTOR(i, 1) = BOUND_VALUE Then
                    l = l + 1
                End If
            Case 1
                If ATEMP_VECTOR(i, 1) >= BOUND_VALUE Then
                    l = l + 1
                End If
            Case Else
                If ATEMP_VECTOR(i, 1) <= BOUND_VALUE Then
                    l = l + 1
                End If
            End Select
        Next i
        
        TEMP_MATRIX(j, 1) = l
    Next j
    BTEMP_VECTOR(1, h) = MATRIX_ELEMENTS_CUMULATIVE_SUM_FUNC(TEMP_MATRIX) / (nLOOPS * nDRAWS)
Next h

RANDOM_SAMPLING_SIMULATION_FUNC = BTEMP_VECTOR

Exit Function
ERROR_LABEL:
RANDOM_SAMPLING_SIMULATION_FUNC = Err.number
End Function
