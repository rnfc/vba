Attribute VB_Name = "FINAN_PORT_MOMENTS_INDEX_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.


'************************************************************************************
'************************************************************************************
'FUNCTION      : PORT_INDEX_MOVER_FUNC
'DESCRIPTION   : Portfolio Index Mover
'LIBRARY       : PORTFOLIO
'GROUP         : TRADE_CONTRIBUTION
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 21/01/2009
'************************************************************************************
'************************************************************************************

Function PORT_INDEX_MOVER_FUNC(ByVal INDEX_LEVEL As Double, _
ByRef SEGMENT_RNG As Variant, _
ByRef WEIGHTS_RNG As Variant, _
ByRef RETURN_RNG As Variant, _
Optional ByVal FACTOR_VAL As Double = 100, _
Optional ByVal OUTPUT As Integer = 1)

Dim i As Long
Dim NROWS As Long

Dim TEMP1_SUM As Double
Dim TEMP2_SUM As Double
Dim TEMP3_SUM As Double

Dim WEIGHTS_VECTOR As Variant
Dim RETURN_VECTOR As Variant
Dim SEGMENT_VECTOR As Variant

Dim TEMP_VECTOR As Variant
Dim TEMP_MATRIX As Variant

On Error GoTo ERROR_LABEL

SEGMENT_VECTOR = SEGMENT_RNG
If UBound(SEGMENT_VECTOR, 1) = 1 Then
    SEGMENT_VECTOR = MATRIX_TRANSPOSE_FUNC(SEGMENT_VECTOR)
End If
WEIGHTS_VECTOR = WEIGHTS_RNG
If UBound(WEIGHTS_VECTOR, 1) = 1 Then
    WEIGHTS_VECTOR = MATRIX_TRANSPOSE_FUNC(WEIGHTS_VECTOR)
End If
If UBound(SEGMENT_VECTOR, 1) <> UBound(WEIGHTS_VECTOR, 1) Then: GoTo ERROR_LABEL
RETURN_VECTOR = RETURN_RNG
If UBound(RETURN_VECTOR, 1) = 1 Then
    RETURN_VECTOR = MATRIX_TRANSPOSE_FUNC(RETURN_VECTOR)
End If
If UBound(SEGMENT_VECTOR, 1) <> UBound(RETURN_VECTOR, 1) Then: GoTo ERROR_LABEL

NROWS = UBound(SEGMENT_VECTOR, 1)

ReDim TEMP_MATRIX(0 To NROWS + 1, 1 To 5)

TEMP_MATRIX(0, 1) = "SEGMENT"
TEMP_MATRIX(0, 2) = "WEIGHT"
TEMP_MATRIX(0, 3) = "RETURN"
TEMP_MATRIX(0, 4) = "CONTRIBUTION"
TEMP_MATRIX(0, 5) = "POINTS"

TEMP1_SUM = 0: TEMP2_SUM = 0: TEMP3_SUM = 0
For i = 1 To NROWS
    TEMP_MATRIX(i, 1) = SEGMENT_VECTOR(i, 1)
    TEMP_MATRIX(i, 2) = WEIGHTS_VECTOR(i, 1)
    TEMP1_SUM = TEMP1_SUM + TEMP_MATRIX(i, 2)
    
    TEMP_MATRIX(i, 3) = RETURN_VECTOR(i, 1)
    TEMP_MATRIX(i, 4) = TEMP_MATRIX(i, 2) * TEMP_MATRIX(i, 3)
    TEMP2_SUM = TEMP2_SUM + TEMP_MATRIX(i, 4)
    
    TEMP_MATRIX(i, 5) = INDEX_LEVEL * TEMP_MATRIX(i, 4) / FACTOR_VAL
    TEMP3_SUM = TEMP3_SUM + TEMP_MATRIX(i, 5)
Next i

TEMP_MATRIX(NROWS + 1, 1) = "SUMS"
TEMP_MATRIX(NROWS + 1, 2) = TEMP1_SUM
TEMP_MATRIX(NROWS + 1, 3) = ""
TEMP_MATRIX(NROWS + 1, 4) = TEMP2_SUM
TEMP_MATRIX(NROWS + 1, 5) = TEMP3_SUM

Select Case OUTPUT
Case 0
    PORT_INDEX_MOVER_FUNC = TEMP_MATRIX
    
Case Else 'Overall Index
    
    ReDim TEMP_VECTOR(1 To 2, 1 To 3)
    
    TEMP_VECTOR(1, 1) = "I0"
    TEMP_VECTOR(2, 1) = INDEX_LEVEL
    
    TEMP_VECTOR(1, 2) = "I1"
    TEMP_VECTOR(2, 2) = TEMP_VECTOR(2, 1) * (1 + TEMP2_SUM / FACTOR_VAL)
    
    TEMP_VECTOR(1, 3) = "DIFF"
    TEMP_VECTOR(2, 3) = TEMP_VECTOR(2, 1) - TEMP_VECTOR(2, 2)
    
    PORT_INDEX_MOVER_FUNC = TEMP_VECTOR
End Select

Exit Function
ERROR_LABEL:
PORT_INDEX_MOVER_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : RNG_PORT_INDEX_CONTRIBUTION_FUNC
'DESCRIPTION   : Portfolio Index Contribution
'LIBRARY       : PORTFOLIO
'GROUP         : TRADE_CONTRIBUTION
'ID            : 002
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 21/01/2009
'************************************************************************************
'************************************************************************************

Function RNG_PORT_INDEX_CONTRIBUTION_FUNC(ByRef DST_RNG As Excel.Range, _
ByVal NASSETS As Long, _
ByVal PERIODS As Long, _
Optional ByVal FACTOR As Double = 100)

Dim i As Long
Dim j As Long
Dim m As Long

Dim RETURNS_RNG As Excel.Range
Dim BUY_HOLD_RNG As Excel.Range

Dim DISC_RNG As Excel.Range
Dim CONT_RNG As Excel.Range
Dim INDEX_RNG As Excel.Range
Dim INDEX_MOVER_RNG As Excel.Range
Dim INDEX_CUMUL_RNG As Excel.Range

Dim TEMP1_STR As String
Dim TEMP2_STR As String

On Error GoTo ERROR_LABEL

RNG_PORT_INDEX_CONTRIBUTION_FUNC = False
m = 7

'-------------------------FIRST PASS: SEGMENT RETURNS---------------------------

Set RETURNS_RNG = Range(DST_RNG.Offset(2, 1), _
DST_RNG.Offset(NASSETS + 1, PERIODS))

RETURNS_RNG.Cells(-2, 0).value = "SEGMENT RETURNS"
RETURNS_RNG.Cells(-2, 0).Font.Bold = True

RETURNS_RNG.Cells(0, 0).value = "PERIODS"
For i = 1 To PERIODS
    RETURNS_RNG.Cells(0, i).value = i
    RETURNS_RNG.Cells(0, i).Font.ColorIndex = 3
Next i

For i = 1 To NASSETS
    RETURNS_RNG.Cells(i, 0).value = "ASSETS: " & i
    RETURNS_RNG.Cells(i, 0).Font.ColorIndex = 3
    
    RETURNS_RNG.Cells(i, PERIODS + 1).FormulaArray = "=" & _
    FACTOR & "*(PRODUCT(1+" & RETURNS_RNG.Rows(i).Address _
    & "/" & FACTOR & ")-1)"
    
Next i

RETURNS_RNG.value = 0
RETURNS_RNG.Font.ColorIndex = 5

'-------------------------SECOND PASS: SEGMENT WEIGHTS---------------------------

Set DST_RNG = DST_RNG.Offset(NASSETS + m)

Set BUY_HOLD_RNG = Range(DST_RNG.Offset(2, 1), _
DST_RNG.Offset(NASSETS + 1, PERIODS))

BUY_HOLD_RNG.Cells(-4, 0).value = "Segment Weights Buy And Hold"
BUY_HOLD_RNG.Cells(-4, 0).Font.Bold = True

BUY_HOLD_RNG.Cells(-2, 0).value = "PERIODS"
For i = 1 To PERIODS
    BUY_HOLD_RNG.Cells(-2, i).formula = "=" & RETURNS_RNG.Cells(0, i).Address
Next i

For i = 1 To NASSETS
    BUY_HOLD_RNG.Cells(i, 0).value = "=" & RETURNS_RNG.Cells(i, 0).Address
Next i

BUY_HOLD_RNG.Cells(-1, 0).value = "PORT_RETURNS_LOG"
BUY_HOLD_RNG.Cells(0, 0).value = "PORT_RETURNS_LIN"

For j = 1 To PERIODS
    TEMP1_STR = ""
    TEMP2_STR = ""
    For i = 1 To NASSETS
        
        TEMP1_STR = TEMP1_STR & _
        BUY_HOLD_RNG.Cells(i, j).Address & "*" & _
        RETURNS_RNG.Cells(i, j).Address & IIf(i <> NASSETS, "+", "")
    
        TEMP2_STR = TEMP2_STR & _
        BUY_HOLD_RNG.Cells(i, j).Address & "*" & FACTOR & "*LN(1+" & _
        RETURNS_RNG.Cells(i, j).Address & "/" & FACTOR & ")" & IIf(i <> NASSETS, "+", "")
    
    Next i
    
    BUY_HOLD_RNG.Cells(-1, j).formula = "=" & TEMP2_STR
    BUY_HOLD_RNG.Cells(-1, j).Font.Bold = True
    BUY_HOLD_RNG.Cells(-1, j).Font.ColorIndex = 3
    
    BUY_HOLD_RNG.Cells(0, j).formula = "=" & TEMP1_STR
    BUY_HOLD_RNG.Cells(0, j).Font.Bold = True
    BUY_HOLD_RNG.Cells(0, j).Font.ColorIndex = 3

    For i = 1 To NASSETS
        If j = 1 Then
            BUY_HOLD_RNG.Cells(i, j).value = 0
            BUY_HOLD_RNG.Cells(i, j).Font.ColorIndex = 5
        Else
            BUY_HOLD_RNG.Cells(i, j).formula = "=" & _
                BUY_HOLD_RNG.Cells(i, j - 1).Address & _
                "*(1+" & RETURNS_RNG.Cells(i, j - 1).Address & "/" & _
                FACTOR & ")/(1+" & BUY_HOLD_RNG.Cells(0, j - 1).Address & _
                "/" & FACTOR & ")"
        End If
    Next i
    BUY_HOLD_RNG.Cells(NASSETS + 1, j).formula = _
        "=SUM(" & BUY_HOLD_RNG.Columns(j).Address & ")"
Next j

BUY_HOLD_RNG.Cells(NASSETS + 1, 0).value = "TOTAL"

    TEMP1_STR = ""
    For i = 1 To NASSETS
        TEMP1_STR = TEMP1_STR & _
        BUY_HOLD_RNG.Cells(i, 1).Address & "*" & _
        RETURNS_RNG.Cells(i, PERIODS + 1).Address & IIf(i <> NASSETS, "+", "")
    Next i
    BUY_HOLD_RNG.Cells(0, PERIODS + 1).formula = "=" & TEMP1_STR
    BUY_HOLD_RNG.Cells(-1, PERIODS + 1).value = "Chain-Linked Monthly Portfolio Returns"
    
    BUY_HOLD_RNG.Cells(0, PERIODS + 2).FormulaArray = "=" & _
    FACTOR & "*(PRODUCT(1+" & BUY_HOLD_RNG.Rows(0).Address _
    & "/" & FACTOR & ")-1)"
    BUY_HOLD_RNG.Cells(-1, PERIODS + 2).value = "Weighted (bop) Chain-Linked " & _
    "Segment Returns"

'----------------------THIRD PASS: SEGMENT DISCRETE RETURNS--------------------------

Set DST_RNG = DST_RNG.Offset(NASSETS + m)

Set DISC_RNG = Range(DST_RNG.Offset(2, 1), _
DST_RNG.Offset(NASSETS + 1, PERIODS))

DISC_RNG.Cells(-2, 0).value = "Discrete Returns"
DISC_RNG.Cells(-2, 0).Font.Bold = True

DISC_RNG.Cells(0, 0).value = "PERIODS"
For i = 1 To PERIODS
    DISC_RNG.Cells(0, i).formula = "=" & RETURNS_RNG.Cells(0, i).Address
Next i

For i = 1 To NASSETS
    DISC_RNG.Cells(i, 0).value = "=" & RETURNS_RNG.Cells(i, 0).Address
Next i

For j = 1 To PERIODS + 1
    For i = 1 To NASSETS
        If j <> PERIODS + 1 Then
            DISC_RNG.Cells(i, j).formula = "=" & _
                RETURNS_RNG.Cells(i, j).Address & "*" & _
            BUY_HOLD_RNG.Cells(i, j).Address
        Else
            DISC_RNG.Cells(i, j).FormulaArray = "=" & _
                FACTOR & "*(PRODUCT(1+" & DISC_RNG.Rows(i).Address _
                & "/" & FACTOR & ")-1)"
        End If
    Next i
    DISC_RNG.Cells(NASSETS + 1, j).formula = _
        "=SUM(" & DISC_RNG.Columns(j).Address & ")"
Next j

DISC_RNG.Cells(NASSETS + 1, 0).value = "TOTAL"

'----------------------FORTH PASS: SEGMENT CONT RETURNS--------------------------

Set DST_RNG = DST_RNG.Offset(NASSETS + m)

Set CONT_RNG = Range(DST_RNG.Offset(2, 1), _
DST_RNG.Offset(NASSETS + 1, PERIODS))

CONT_RNG.Cells(-2, 0).value = "Continuous Returns"
CONT_RNG.Cells(-2, 0).Font.Bold = True

CONT_RNG.Cells(0, 0).value = "PERIODS"
For i = 1 To PERIODS
    CONT_RNG.Cells(0, i).formula = "=" & RETURNS_RNG.Cells(0, i).Address
Next i

For i = 1 To NASSETS
    CONT_RNG.Cells(i, 0).value = "=" & RETURNS_RNG.Cells(i, 0).Address
Next i

For j = 1 To PERIODS + 1
    For i = 1 To NASSETS
        If j <> PERIODS + 1 Then
            CONT_RNG.Cells(i, j).formula = "=" & _
                BUY_HOLD_RNG.Cells(i, j).Address & "*" & _
            FACTOR & "*LN(1+" & RETURNS_RNG.Cells(i, j).Address & "/" & _
            FACTOR & ")"
        Else
            CONT_RNG.Cells(i, j).formula = "=SUM(" & _
                CONT_RNG.Rows(i).Address & ")"
        End If
    Next i
    CONT_RNG.Cells(NASSETS + 1, j).formula = _
        "=SUM(" & CONT_RNG.Columns(j).Address & ")"
Next j

CONT_RNG.Cells(NASSETS + 1, 0).value = "TOTAL"

'----------------------FIFTH PASS: INDEX PORTFOLIO--------------------------

Set DST_RNG = DST_RNG.Offset(NASSETS + m)

Set INDEX_RNG = Range(DST_RNG.Offset(2, 1), _
DST_RNG.Offset(NASSETS + 1, PERIODS))

INDEX_RNG.Cells(-2, 0).value = "Index Portfolio"
INDEX_RNG.Cells(-2, 0).Font.Bold = True

INDEX_RNG.Cells(0, 0).value = 100
INDEX_RNG.Cells(0, 0).Font.Bold = True
INDEX_RNG.Cells(0, 0).Font.ColorIndex = 3

INDEX_RNG.Cells(1, 0).value = "Indexed Portfolio Value"
INDEX_RNG.Cells(2, 0).value = "Index Change"

For i = 1 To PERIODS
    INDEX_RNG.Cells(0, i).formula = "=" & RETURNS_RNG.Cells(0, i).Address
Next i


For j = 1 To PERIODS
    If j = 1 Then
        INDEX_RNG.Cells(1, j).formula = "=" & _
            INDEX_RNG.Cells(0, 0).Address & "*(1+" & _
            BUY_HOLD_RNG.Cells(0, j).Address & "/" & _
            FACTOR & ")"
    
        INDEX_RNG.Cells(2, j).formula = "=" & _
            INDEX_RNG.Cells(1, j).Address & "-" & _
            INDEX_RNG.Cells(0, 0).Address
    Else
        INDEX_RNG.Cells(1, j).formula = "=" & _
            INDEX_RNG.Cells(1, j - 1).Address & "*(1+" & _
            BUY_HOLD_RNG.Cells(0, j).Address & "/" & _
            FACTOR & ")"
        INDEX_RNG.Cells(2, j).formula = "=" & _
            INDEX_RNG.Cells(1, j).Address & "-" & _
            INDEX_RNG.Cells(1, j - 1).Address
    End If
Next j

INDEX_RNG.Cells(2, PERIODS + 1).formula = "=SUM(" & INDEX_RNG.Rows(2).Address & ")"

'----------------------SIXTH PASS: INDEX_MOVER --------------------------

Set DST_RNG = DST_RNG.Offset(m, 0)

Set INDEX_MOVER_RNG = Range(DST_RNG.Offset(2, 1), _
DST_RNG.Offset(NASSETS + 1, PERIODS))

INDEX_MOVER_RNG.Cells(-2, 0).value = "Index Movers"
INDEX_MOVER_RNG.Cells(-2, 0).Font.Bold = True

INDEX_MOVER_RNG.Cells(0, 0).value = "PERIODS"
For i = 1 To PERIODS
    INDEX_MOVER_RNG.Cells(0, i).formula = "=" & RETURNS_RNG.Cells(0, i).Address
Next i

For i = 1 To NASSETS
    INDEX_MOVER_RNG.Cells(i, 0).value = "=" & RETURNS_RNG.Cells(i, 0).Address
Next i

For j = 1 To PERIODS + 1
    For i = 1 To NASSETS
        If j <> PERIODS + 1 Then
            If j = 1 Then
                INDEX_MOVER_RNG.Cells(i, j).formula = "=" & _
                    INDEX_RNG.Cells(0, 0).Address & "*" & _
                    DISC_RNG.Cells(i, j).Address & "/" & _
                    FACTOR
            Else
                INDEX_MOVER_RNG.Cells(i, j).formula = "=" & _
                    INDEX_RNG.Cells(1, j - 1).Address & "*" & _
                    DISC_RNG.Cells(i, j).Address & "/" & _
                    FACTOR
            End If
        Else
            INDEX_MOVER_RNG.Cells(i, j).formula = "=SUM(" & _
                INDEX_MOVER_RNG.Rows(i).Address & ")"
        End If
    Next i
    INDEX_MOVER_RNG.Cells(NASSETS + 1, j).formula = _
        "=SUM(" & INDEX_MOVER_RNG.Columns(j).Address & ")"
Next j

INDEX_MOVER_RNG.Cells(NASSETS + 1, 0).value = "TOTAL"

'----------------------SEVENTH PASS: INDEX_CUMUL --------------------------

Set DST_RNG = DST_RNG.Offset(NASSETS + m)

Set INDEX_CUMUL_RNG = Range(DST_RNG.Offset(2, 1), _
DST_RNG.Offset(NASSETS + 1, PERIODS))

INDEX_CUMUL_RNG.Cells(-2, 0).value = "Index Movers Cumulated"
INDEX_CUMUL_RNG.Cells(-2, 0).Font.Bold = True

INDEX_CUMUL_RNG.Cells(0, 0).value = "PERIODS"
For i = 1 To PERIODS
    INDEX_CUMUL_RNG.Cells(0, i).formula = "=" & RETURNS_RNG.Cells(0, i).Address
Next i

For i = 1 To NASSETS
    INDEX_CUMUL_RNG.Cells(i, 0).value = "=" & RETURNS_RNG.Cells(i, 0).Address
Next i

'

For j = 1 To PERIODS + 1
    For i = 1 To NASSETS
        If j <> PERIODS + 1 Then
            If j = 1 Then
                INDEX_CUMUL_RNG.Cells(i, j).formula = "=" & _
                    INDEX_MOVER_RNG.Cells(i, j).Address
            Else
                INDEX_CUMUL_RNG.Cells(i, j).formula = "=" & _
                    INDEX_MOVER_RNG.Cells(i, j).Address & "+" & _
                    INDEX_CUMUL_RNG.Cells(i, j - 1).Address
            End If
        Else
            INDEX_CUMUL_RNG.Cells(i, j).formula = "=SUM(" & _
                INDEX_CUMUL_RNG.Rows(i).Address & ")"
        End If
    Next i
    
    INDEX_CUMUL_RNG.Cells(NASSETS + 1, j).formula = _
        "=SUM(" & INDEX_CUMUL_RNG.Columns(j).Address & ")"
    
    If j <> 1 Then
        INDEX_CUMUL_RNG.Cells(NASSETS + 2, j).formula = _
            "=" & FACTOR & "*((1+" & _
                INDEX_CUMUL_RNG.Cells(NASSETS + 2, j - 1).Address & _
                    "/" & FACTOR & ")*(1+" & _
                BUY_HOLD_RNG.Cells(0, j).Address & _
                    "/" & FACTOR & ")-1)"
    Else
        INDEX_CUMUL_RNG.Cells(NASSETS + 2, j).formula = _
            "=" & BUY_HOLD_RNG.Cells(0, j).Address
    End If
Next j

INDEX_CUMUL_RNG.Cells(NASSETS + 1, 0).value = "TOTAL"
INDEX_CUMUL_RNG.Cells(NASSETS + 2, 0).value = "Cummulative Chain-Linked Return"

RNG_PORT_INDEX_CONTRIBUTION_FUNC = True

Exit Function
ERROR_LABEL:
RNG_PORT_INDEX_CONTRIBUTION_FUNC = False
End Function
