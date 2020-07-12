Attribute VB_Name = "FINAN_FI_BOND_TREASURY_LIBR"


Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.

'************************************************************************************
'************************************************************************************
'FUNCTION      : TREASURY_YIELD_INTERPOLATION_FUNC
'DESCRIPTION   :
'LIBRARY       : BOND
'GROUP         : TREASURY
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 10-06-2008
'************************************************************************************
'************************************************************************************

Function TREASURY_YIELD_INTERPOLATION_FUNC(ByRef TENOR_RNG As Variant, _
ByRef YIELD_RNG As Variant, _
ByRef TERM_RNG As Variant, _
Optional ByVal nSTEPS As Long = 15, _
Optional ByVal TENOR_VAL As Long = 5, _
Optional ByRef FREQUENCY_VAL As Double = 1, _
Optional ByVal FACTOR_VAL As Double = 100, _
Optional ByVal tolerance As Double = 0.01, _
Optional ByVal VERSION As Long = 0)

'FREQUENCY = 1 --> Annually
'FREQUENCY = 4 --> Quaterly
'FREQUENCY = 12 --> Monthly

'VERSION
'0: Interest rates move up tolerance annually for five years and then remain
'   steady for n years.

'1: Move up tolerance annually for TENOR_VAL years, then down tolerance
'   annually for TENOR_VAL years, then steady.

'Else: Rates move down tolerance annually for five years, and then remain steady
'   for n years.

Dim i As Long
Dim j As Long

Dim NROWS As Long
Dim NSIZE As Long

Dim FREQUENCY As Double

Dim TERM_VECTOR As Variant
Dim TENOR_VECTOR As Variant

Dim YIELD_VECTOR As Variant
Dim TEMP_VECTOR As Variant
Dim TEMP_MATRIX As Variant

On Error GoTo ERROR_LABEL

TENOR_VECTOR = TENOR_RNG
If UBound(TENOR_VECTOR, 1) = 1 Then
    TENOR_VECTOR = MATRIX_TRANSPOSE_FUNC(TENOR_VECTOR)
End If

YIELD_VECTOR = YIELD_RNG
If UBound(YIELD_VECTOR, 1) = 1 Then
    YIELD_VECTOR = MATRIX_TRANSPOSE_FUNC(YIELD_VECTOR)
End If

If UBound(TENOR_VECTOR, 1) <> UBound(YIELD_VECTOR, 1) Then: GoTo ERROR_LABEL
NROWS = UBound(TENOR_VECTOR, 1)

TERM_VECTOR = TERM_RNG
If UBound(TERM_VECTOR, 1) = 1 Then
    TERM_VECTOR = MATRIX_TRANSPOSE_FUNC(TERM_VECTOR)
End If

NSIZE = UBound(TERM_VECTOR, 1)

ReDim TEMP_MATRIX(0 To nSTEPS, 0 To NSIZE - 1)
TEMP_VECTOR = TREASURY_INTERPOLATION_OBJ_FUNC(TENOR_VECTOR, YIELD_VECTOR, TERM_VECTOR, FACTOR_VAL)
For i = 0 To NSIZE - 1 'Interpolation Yield (%)
    TEMP_MATRIX(0, i) = TEMP_VECTOR(i + 1, 2) / FACTOR_VAL
    'Set data of step 0 by interpolation data.
Next i

FREQUENCY = 1 / FREQUENCY_VAL

'---------------------------------------------------------------------------------
Select Case VERSION
'---------------------------------------------------------------------------------
Case 0
'---------------------------------------------------------------------------------
    For i = 0 To NSIZE - 1
        For j = 1 To nSTEPS
            TEMP_MATRIX(j, i) = TEMP_MATRIX(j - 1, i) + _
                               FREQUENCY * tolerance * _
                               (-((j - TENOR_VAL / FREQUENCY) <= 0))
                               'Curve moves up tolerance annually for first
                               'TENOR_VAL years then remain steady.
        Next j
    Next i
'---------------------------------------------------------------------------------
Case 1
'---------------------------------------------------------------------------------
    For i = 0 To NSIZE - 1
        For j = 1 To nSTEPS
            TEMP_MATRIX(j, i) = TEMP_MATRIX(j - 1, i) + _
                               FREQUENCY * tolerance * (-((j - _
                               TENOR_VAL / FREQUENCY) <= 0)) - _
                               FREQUENCY * tolerance * (-((j - _
                               TENOR_VAL / FREQUENCY) > 0 And _
                               (j - (TENOR_VAL * 2) / FREQUENCY) <= 0))
                               'Curve moves up tolerance annually for first
                               'TENOR_VAL
                               'years and down tolerance annually for another
                               'TENOR_VAL years then remain steady.
        Next j
    Next i
'---------------------------------------------------------------------------------
Case Else
'---------------------------------------------------------------------------------
    For i = 0 To NSIZE - 1
        For j = 1 To nSTEPS
            TEMP_MATRIX(j, i) = TEMP_MATRIX(j - 1, i) - _
                                FREQUENCY * tolerance * (-((j - _
                                TENOR_VAL / FREQUENCY) <= 0))
            'Curve moves down tolerance annually for first TENOR_VAL years
            'then remain steady.
        Next j
    Next i
'---------------------------------------------------------------------------------
End Select
'---------------------------------------------------------------------------------

For i = 0 To NSIZE - 1
    TEMP_MATRIX(0, i) = TEMP_MATRIX(0, i) * FACTOR_VAL
    For j = 1 To nSTEPS 'Steps from 0 to X
            If TEMP_MATRIX(j, i) <= 0 Then
                TEMP_MATRIX(j, i) = "" 'Set all negtive value by NULL.
            Else
                TEMP_MATRIX(j, i) = TEMP_MATRIX(j, i) * FACTOR_VAL
            End If
    Next j
Next i

TREASURY_YIELD_INTERPOLATION_FUNC = TEMP_MATRIX

Exit Function
ERROR_LABEL:
TREASURY_YIELD_INTERPOLATION_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************



'************************************************************************************
'************************************************************************************
'FUNCTION      : TREASURY_INTERPOLATION_OBJ_FUNC
'DESCRIPTION   :
'LIBRARY       : BOND
'GROUP         : TREASURY
'ID            : 002
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 10-06-2008
'************************************************************************************
'************************************************************************************

Private Function TREASURY_INTERPOLATION_OBJ_FUNC(ByRef TENOR_RNG As Variant, _
ByRef YIELD_RNG As Variant, _
ByRef TERM_RNG As Variant, _
Optional ByVal FACTOR_VAL As Double = 100)

Dim i As Long
Dim j As Long 'Define loop index.

Dim NROWS As Long
Dim NSIZE As Long

Dim TERM_VECTOR As Variant
Dim TENOR_VECTOR As Variant

Dim YIELD_VECTOR As Variant
Dim TEMP_VECTOR As Variant
Dim TEMP_MATRIX As Variant

On Error GoTo ERROR_LABEL

TENOR_VECTOR = TENOR_RNG
If UBound(TENOR_VECTOR, 1) = 1 Then: _
    TENOR_VECTOR = MATRIX_TRANSPOSE_FUNC(TENOR_VECTOR)
    
YIELD_VECTOR = YIELD_RNG
If UBound(YIELD_VECTOR, 1) = 1 Then: _
    YIELD_VECTOR = MATRIX_TRANSPOSE_FUNC(YIELD_VECTOR)
    
If UBound(TENOR_VECTOR, 1) <> UBound(YIELD_VECTOR, 1) Then: GoTo ERROR_LABEL
NROWS = UBound(TENOR_VECTOR, 1)

TERM_VECTOR = TERM_RNG
If UBound(TERM_VECTOR, 1) = 1 Then: _
    TERM_VECTOR = MATRIX_TRANSPOSE_FUNC(TERM_VECTOR)
NSIZE = UBound(TERM_VECTOR, 1)

For i = 1 To NROWS
    YIELD_VECTOR(i, 1) = YIELD_VECTOR(i, 1) / FACTOR_VAL
Next i

ReDim TEMP_VECTOR(1 To NSIZE, 1 To 1)

'Inperpolation loops.
i = 0
j = 0   'Initiate pointers.
Do While i < NROWS
    If TERM_VECTOR(j + 1, 1) = TENOR_VECTOR(i + 1, 1) Then
        TEMP_VECTOR(j + 1, 1) = YIELD_VECTOR(i + 1, 1)
        i = i + 1
        j = j + 1   'Pick out existed annual data. Both pointers
        'move ahead.
    ElseIf TERM_VECTOR(j + 1, 1) < TENOR_VECTOR(i + 1, 1) Then
        TEMP_VECTOR(j + 1, 1) = YIELD_VECTOR(i - 1 + 1, 1) + _
        (YIELD_VECTOR(i + 1, 1) - _
        YIELD_VECTOR(i - 1 + 1, 1)) * (TERM_VECTOR(j + 1, 1) - _
        TENOR_VECTOR(i - 1 + 1, 1)) / (TENOR_VECTOR(i + 1, 1) - _
        TENOR_VECTOR(i - 1 + 1, 1))
        j = j + 1   'Interpolate these gaps when TENOR_VECTOR greater
        'than interpolation term. Then the pointer to interpolation
        'data moves ahead.
    Else
        i = i + 1   'Pointer to input data moves ahead when
        'TENOR_VECTOR less than interpolation term.
    End If
Loop

'Output interpolation data to output range.
ReDim TEMP_MATRIX(1 To NSIZE, 1 To 2)

For i = 1 To NSIZE
    TEMP_MATRIX(i, 1) = TERM_VECTOR(i, 1)
    TEMP_MATRIX(i, 2) = TEMP_VECTOR(i, 1) * FACTOR_VAL
Next i

TREASURY_INTERPOLATION_OBJ_FUNC = TEMP_MATRIX

Exit Function
ERROR_LABEL:
TREASURY_INTERPOLATION_OBJ_FUNC = Err.number
End Function
