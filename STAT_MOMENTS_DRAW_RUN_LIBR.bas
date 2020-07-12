Attribute VB_Name = "STAT_MOMENTS_DRAW_RUN_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.

'************************************************************************************
'************************************************************************************
'FUNCTION      : VECTOR_NMDD_FUNC
'DESCRIPTION   : NROWS-th non-overplapping drawdown
'LIBRARY       : STATISTICS
'GROUP         : DRAWDOWN
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 04/06/2009
'************************************************************************************
'************************************************************************************

Function VECTOR_NMDD_FUNC(ByRef DATA_RNG As Variant, _
Optional ByRef NROWS As Long = 0)

Dim i As Long
Dim j As Long

Dim TEMP_VECTOR As Variant
Dim DATA_VECTOR As Variant

On Error GoTo ERROR_LABEL

DATA_VECTOR = DATA_RNG
If UBound(DATA_VECTOR, 1) = 1 Then
    DATA_VECTOR = MATRIX_TRANSPOSE_FUNC(DATA_VECTOR)
End If
If NROWS = 0 Or NROWS > UBound(DATA_VECTOR, 1) Then
    NROWS = UBound(DATA_VECTOR, 1)
End If

For i = 1 To NROWS
    TEMP_VECTOR = VECTOR_MDD_FUNC(DATA_VECTOR)
    If IsArray(TEMP_VECTOR) Then
        For j = TEMP_VECTOR(2, 1) To TEMP_VECTOR(3, 1)
            DATA_VECTOR(j, 1) = 0
        Next j
    Else
        VECTOR_NMDD_FUNC = CVErr(xlErrNA)
        Exit Function
    End If
Next i

VECTOR_NMDD_FUNC = MATRIX_TRANSPOSE_FUNC(TEMP_VECTOR)

Exit Function
ERROR_LABEL:
VECTOR_NMDD_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : VECTOR_NMRU_FUNC
'DESCRIPTION   : NROWS-th non-overplapping run up
'LIBRARY       : STATISTICS
'GROUP         : DRAWDOWN
'ID            : 009
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 04/06/2009
'************************************************************************************
'************************************************************************************

Function VECTOR_NMRU_FUNC(ByRef DATA_RNG As Variant, _
Optional ByRef NROWS As Long = 0)

Dim i As Long
Dim j As Long

Dim TEMP_VECTOR As Variant
Dim DATA_VECTOR As Variant

On Error GoTo ERROR_LABEL

DATA_VECTOR = DATA_RNG
If UBound(DATA_VECTOR, 1) = 1 Then
    DATA_VECTOR = MATRIX_TRANSPOSE_FUNC(DATA_VECTOR)
End If
If NROWS = 0 Or NROWS > UBound(DATA_VECTOR, 1) Then
    NROWS = UBound(DATA_VECTOR, 1)
End If

For i = 1 To NROWS
    TEMP_VECTOR = VECTOR_MRU_FUNC(DATA_VECTOR)
    If IsArray(TEMP_VECTOR) Then
        For j = TEMP_VECTOR(2, 1) To TEMP_VECTOR(3, 1)
            DATA_VECTOR(j, 1) = 0
        Next j
    Else
        VECTOR_NMRU_FUNC = CVErr(xlErrNA)
        Exit Function
    End If
Next i
VECTOR_NMRU_FUNC = MATRIX_TRANSPOSE_FUNC(TEMP_VECTOR)

Exit Function
ERROR_LABEL:
VECTOR_NMRU_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : VECTOR_MDD_FUNC
'DESCRIPTION   : Max drawdown
'LIBRARY       : STATISTICS
'GROUP         : DRAWDOWN
'ID            : 010
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 04/06/2009
'************************************************************************************
'************************************************************************************

Function VECTOR_MDD_FUNC(ByRef DATA_RNG As Variant)

'DATA_RNG
'time-series vector (format: indexed (start=1) % figures)

Dim h As Long
Dim i As Long 'Start index of maximum drawdown phase
Dim j As Long 'End index of maximum drawdown phase (lowest point)
Dim k As Long 'End index of recovery phase
Dim l As Long

Dim NROWS As Long
Dim DATA_VECTOR As Variant

Dim MAX_VAL As Double
Dim MAX_LAST_VAL As Double
Dim MAX_TEMP_VAL As Double
Dim MAX_DD_VAL As Double
Dim MAX_DRAWDOWN_VAL As Double 'Maximum drawdown (relative, positive sign)

On Error GoTo ERROR_LABEL

DATA_VECTOR = DATA_RNG
If UBound(DATA_VECTOR, 1) = 1 Then
    DATA_VECTOR = MATRIX_TRANSPOSE_FUNC(DATA_VECTOR)
End If
NROWS = UBound(DATA_VECTOR, 1)

MAX_VAL = 1 + MAXIMUM_FUNC(DATA_VECTOR(1, 1), 0)
MAX_DRAWDOWN_VAL = 0
MAX_LAST_VAL = 1
i = 1
k = 0
If DATA_VECTOR(1, 1) < 0 Then
    l = 0
Else
    l = 1
End If
'-----------------------------------------------------------------------
For h = 1 To NROWS
'-----------------------------------------------------------------------
    MAX_TEMP_VAL = MAX_LAST_VAL * (1 + DATA_VECTOR(h, 1))
    If MAX_TEMP_VAL > MAX_VAL Then
        MAX_VAL = MAX_TEMP_VAL
        l = h
    End If
    MAX_DD_VAL = 1 - (MAX_TEMP_VAL / MAX_VAL)
    If MAX_DD_VAL > MAX_DRAWDOWN_VAL Then
        MAX_DRAWDOWN_VAL = MAX_DD_VAL
        i = l
        j = h
    End If
    If i = l Then: k = h + 1
    MAX_LAST_VAL = MAX_TEMP_VAL
'-----------------------------------------------------------------------
Next h
'-----------------------------------------------------------------------

'-----------------------------------------------------------------------
If MAX_DRAWDOWN_VAL <> 0 Then
'-----------------------------------------------------------------------
    ReDim TEMP_VECTOR(1 To 4, 1 To 1)
    TEMP_VECTOR(1, 1) = -MAX_DRAWDOWN_VAL 'Value
    TEMP_VECTOR(2, 1) = i + 1 'Start
    TEMP_VECTOR(3, 1) = j 'End
    If k > NROWS Then
        TEMP_VECTOR(4, 1) = CVErr(xlErrNA) 'Recovery
    Else
        TEMP_VECTOR(4, 1) = k 'Recovery
    End If
    VECTOR_MDD_FUNC = TEMP_VECTOR
'-----------------------------------------------------------------------
Else
'-----------------------------------------------------------------------
    VECTOR_MDD_FUNC = CVErr(xlErrNA)
'-----------------------------------------------------------------------
End If
'-----------------------------------------------------------------------

Exit Function
ERROR_LABEL:
VECTOR_MDD_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : VECTOR_MRU_FUNC
'DESCRIPTION   : Maximum Run Up
'LIBRARY       : STATISTICS
'GROUP         : DRAWDOWN
'ID            : 011
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 04/06/2009
'************************************************************************************
'************************************************************************************

Function VECTOR_MRU_FUNC(ByRef DATA_RNG As Variant)
'DATA_RNG: time-series vector (format: indexed (start=1) % figures)

Dim h As Long
Dim i As Long 'Start index of MINimum Run Up phase
Dim j As Long 'End index of MINimum Run Up phase (lowest point)
Dim k As Long 'End index of recovery phase
Dim l As Long

Dim NROWS As Long
Dim DATA_VECTOR As Variant

Dim MIN_VAL As Double
Dim MIN_LAST_VAL As Double
Dim MIN_TEMP_VAL As Double
Dim MIN_DD_VAL As Double
Dim MAX_RUNUP_VAL As Double 'MINimum Run Up (relative, positive sign)

On Error GoTo ERROR_LABEL

DATA_VECTOR = DATA_RNG
If UBound(DATA_VECTOR, 1) = 1 Then
    DATA_VECTOR = MATRIX_TRANSPOSE_FUNC(DATA_VECTOR)
End If
NROWS = UBound(DATA_VECTOR, 1)

MIN_VAL = 1 + MINIMUM_FUNC(DATA_VECTOR(1, 1), 0)
MAX_RUNUP_VAL = 0
MIN_LAST_VAL = 1
i = 1
k = 0
If DATA_VECTOR(1, 1) > 0 Then
    l = 0
Else
    l = 1
End If
'--------------------------------------------------------------------
For h = 1 To NROWS
'--------------------------------------------------------------------
    MIN_TEMP_VAL = MIN_LAST_VAL * (1 + DATA_VECTOR(h, 1))
    If MIN_TEMP_VAL < MIN_VAL Then
        MIN_VAL = MIN_TEMP_VAL
        l = h
    End If
    MIN_DD_VAL = 1 - (MIN_TEMP_VAL / MIN_VAL)
    If MIN_DD_VAL < MAX_RUNUP_VAL Then
        MAX_RUNUP_VAL = MIN_DD_VAL
        i = l
        j = h
    End If
    If i = l Then: k = h + 1
    MIN_LAST_VAL = MIN_TEMP_VAL
'--------------------------------------------------------------------
Next h
'--------------------------------------------------------------------

If MAX_RUNUP_VAL <> 0 Then 'MINimum Run Up
    ReDim TEMP_VECTOR(1 To 4, 1 To 1)
    TEMP_VECTOR(1, 1) = -MAX_RUNUP_VAL 'Value
    TEMP_VECTOR(2, 1) = i + 1 'Start
    TEMP_VECTOR(3, 1) = j 'End
    If k > NROWS Then
        TEMP_VECTOR(4, 1) = CVErr(xlErrNA) 'Recovery
    Else
        TEMP_VECTOR(4, 1) = k 'Recovery
    End If
    VECTOR_MRU_FUNC = TEMP_VECTOR
Else
    VECTOR_MRU_FUNC = CVErr(xlErrNA)
End If

Exit Function
ERROR_LABEL:
VECTOR_MRU_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : VECTOR_UW_FUNC
'DESCRIPTION   : Underwater
'LIBRARY       : STATISTICS
'GROUP         : DRAWDOWN
'ID            : 012
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 04/06/2009
'************************************************************************************
'************************************************************************************

Function VECTOR_UW_FUNC(ByRef DATA_RNG As Variant)
'DATA_RNG: time-series vector (format: indexed (start=1) % figures)

Dim j As Long
Dim NROWS As Long

Dim MAX_VAL As Double
Dim MAX_LAST_VAL As Double
Dim MAX_TEMP_VAL As Double

Dim TEMP_VECTOR As Variant
Dim DATA_VECTOR As Variant

On Error GoTo ERROR_LABEL

DATA_VECTOR = DATA_RNG
If UBound(DATA_VECTOR, 1) = 1 Then
    DATA_VECTOR = MATRIX_TRANSPOSE_FUNC(DATA_VECTOR)
End If
NROWS = UBound(DATA_VECTOR, 1)

MAX_VAL = 1 + MAXIMUM_FUNC(DATA_VECTOR(1, 1), 0)
MAX_LAST_VAL = 1
ReDim TEMP_VECTOR(1 To NROWS, 1 To 1)
For j = 1 To NROWS
    MAX_TEMP_VAL = MAX_LAST_VAL * (1 + DATA_VECTOR(j, 1))
    If MAX_TEMP_VAL > MAX_VAL Then: MAX_VAL = MAX_TEMP_VAL
    TEMP_VECTOR(j, 1) = (MAX_TEMP_VAL / MAX_VAL) - 1
    MAX_LAST_VAL = MAX_TEMP_VAL
Next j

VECTOR_UW_FUNC = TEMP_VECTOR
'unsorted vector of all drawdown values (with negative sign)

Exit Function
ERROR_LABEL:
VECTOR_UW_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : VECTOR_OW_FUNC
'DESCRIPTION   : Overwater
'LIBRARY       : STATISTICS
'GROUP         : DRAWDOWN
'ID            : 013
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 04/06/2009
'************************************************************************************
'************************************************************************************

Function VECTOR_OW_FUNC(ByRef DATA_RNG As Variant)
'DATA_RNG: time-series vector (format: indexed (start=1) % figures)

Dim j As Long
Dim NROWS As Long

Dim MAX_VAL As Double
Dim MAX_LAST_VAL As Double
Dim MAX_TEMP_VAL As Double

Dim TEMP_VECTOR As Variant
Dim DATA_VECTOR As Variant

On Error GoTo ERROR_LABEL

DATA_VECTOR = DATA_RNG
If UBound(DATA_VECTOR, 1) = 1 Then
    DATA_VECTOR = MATRIX_TRANSPOSE_FUNC(DATA_VECTOR)
End If
NROWS = UBound(DATA_VECTOR, 1)

MAX_VAL = 1 + MINIMUM_FUNC(DATA_VECTOR(1, 1), 0)
MAX_LAST_VAL = 1

ReDim TEMP_VECTOR(1 To NROWS, 1 To 1)

For j = 1 To NROWS
    MAX_TEMP_VAL = MAX_LAST_VAL * (1 + DATA_VECTOR(j, 1))
    If MAX_TEMP_VAL < MAX_VAL Then
        MAX_VAL = MAX_TEMP_VAL
    End If
    TEMP_VECTOR(j, 1) = (MAX_TEMP_VAL / MAX_VAL) - 1
    MAX_LAST_VAL = MAX_TEMP_VAL
Next j
VECTOR_OW_FUNC = TEMP_VECTOR
'unsorted vector of all drawdown values (with negative sign)

Exit Function
ERROR_LABEL:
VECTOR_OW_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : VECTOR_DD_FUNC
'DESCRIPTION   : Drawdowns
'LIBRARY       : STATISTICS
'GROUP         : DRAWDOWN
'ID            : 014
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 04/06/2009
'************************************************************************************
'************************************************************************************

Function VECTOR_DD_FUNC(ByRef DATA_RNG As Variant, _
Optional ByVal k As Long = 0)

' DATA_VECTOR...  time-series vector (format: indexed (start=1) % figures)
' k (optional)... k-th drawdown (all if missing)

Dim j As Long

Dim NROWS As Long
Dim TEMP_VECTOR As Variant
Dim DATA_VECTOR As Variant

Dim MAX_VAL As Double
Dim MAX_LAST_VAL As Double
Dim MAX_TEMP_VAL As Double

On Error GoTo ERROR_LABEL

DATA_VECTOR = DATA_RNG
If UBound(DATA_VECTOR, 1) = 1 Then
    DATA_VECTOR = MATRIX_TRANSPOSE_FUNC(DATA_VECTOR)
End If
NROWS = UBound(DATA_VECTOR, 1)

MAX_VAL = 1 + MAXIMUM_FUNC(DATA_VECTOR(1, 1), 0)
MAX_LAST_VAL = 1
ReDim TEMP_VECTOR(1 To NROWS, 1 To 1)
For j = 1 To NROWS
    MAX_TEMP_VAL = MAX_LAST_VAL * (1 + DATA_VECTOR(j, 1))
    If MAX_TEMP_VAL > MAX_VAL Then: MAX_VAL = MAX_TEMP_VAL
    TEMP_VECTOR(j, 1) = 1 - (MAX_TEMP_VAL / MAX_VAL)
    MAX_LAST_VAL = MAX_TEMP_VAL
Next j
TEMP_VECTOR = MATRIX_QUICK_SORT_FUNC(TEMP_VECTOR, 1, 0)
If k = 0 Then
    ReDim DATA_VECTOR(1 To NROWS, 1 To 1)
    For j = 1 To NROWS
        DATA_VECTOR(j, 1) = -TEMP_VECTOR(j, 1)
    Next j
    VECTOR_DD_FUNC = DATA_VECTOR
    'NROWS-th drawdown or vector with all drawdowns (positive sign)
Else
    VECTOR_DD_FUNC = -TEMP_VECTOR(k, 1)
    If VECTOR_DD_FUNC = 0 Then VECTOR_DD_FUNC = CVErr(xlErrNA)
End If

Exit Function
ERROR_LABEL:
VECTOR_DD_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : VECTOR_RU_FUNC
'DESCRIPTION   : Run ups
'LIBRARY       : STATISTICS
'GROUP         : DRAWDOWN
'ID            : 015
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 04/06/2009
'************************************************************************************
'************************************************************************************

Function VECTOR_RU_FUNC(ByRef DATA_RNG As Variant, _
Optional ByVal k As Long = 0)

'DATA_RNG: time-series vector (format: indexed (start=1) % figures)
' k (optional)... k-th drawdown (all if missing)

Dim j As Long

Dim NROWS As Long
Dim TEMP_VECTOR As Variant
Dim DATA_VECTOR As Variant

Dim MAX_VAL As Double
Dim MAX_LAST_VAL As Double
Dim MAX_TEMP_VAL As Double

On Error GoTo ERROR_LABEL

DATA_VECTOR = DATA_RNG
If UBound(DATA_VECTOR, 1) = 1 Then
    DATA_VECTOR = MATRIX_TRANSPOSE_FUNC(DATA_VECTOR)
End If
NROWS = UBound(DATA_VECTOR, 1)

MAX_VAL = 1 + MINIMUM_FUNC(DATA_VECTOR(1, 1), 0)
MAX_LAST_VAL = 1
ReDim TEMP_VECTOR(1 To NROWS, 1 To 1)
For j = 1 To NROWS
    MAX_TEMP_VAL = MAX_LAST_VAL * (1 + DATA_VECTOR(j, 1))
    If MAX_TEMP_VAL < MAX_VAL Then: MAX_VAL = MAX_TEMP_VAL
    TEMP_VECTOR(j, 1) = 1 - (MAX_TEMP_VAL / MAX_VAL)
    MAX_LAST_VAL = MAX_TEMP_VAL
Next j
TEMP_VECTOR = MATRIX_QUICK_SORT_FUNC(TEMP_VECTOR, 1, 1)
If k = 0 Then
    ReDim DATA_VECTOR(1 To NROWS, 1 To 1)
    For j = 1 To NROWS
        DATA_VECTOR(j, 1) = -TEMP_VECTOR(j, 1)
    Next j
    VECTOR_RU_FUNC = DATA_VECTOR
    'NROWS-th drawdown or vector with all drawdowns (positive sign)
Else
    VECTOR_RU_FUNC = -TEMP_VECTOR(k, 1)
    If VECTOR_RU_FUNC = 0 Then VECTOR_RU_FUNC = CVErr(xlErrNA)
End If

Exit Function
ERROR_LABEL:
VECTOR_RU_FUNC = Err.number
End Function
