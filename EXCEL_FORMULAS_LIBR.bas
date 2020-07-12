Attribute VB_Name = "EXCEL_FORMULAS_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.


'************************************************************************************
'************************************************************************************
'FUNCTION      : RNG_SHOW_FORMULA_FUNC
'DESCRIPTION   : Check if the specify cell is part of an array formula
'LIBRARY       : EXCEL
'GROUP         : FORMULAS
'ID            : 001
'LAST UPDATE   : 14/02/2008
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'************************************************************************************
'************************************************************************************

Function RNG_SHOW_FORMULA_FUNC(ByRef DATA_RNG As Excel.Range)

Dim TEMP_STR As String

On Error GoTo ERROR_LABEL
    
If DATA_RNG.value = "" Then
    TEMP_STR = ""
    GoTo 1983
End If

If DATA_RNG.HasFormula Then
    TEMP_STR = DATA_RNG.FormulaR1C1Local
    If DATA_RNG.HasArray Then: TEMP_STR = "{" & TEMP_STR & "}"
Else
    TEMP_STR = DATA_RNG.value
End If

1983:
RNG_SHOW_FORMULA_FUNC = TEMP_STR

Exit Function
ERROR_LABEL:
RNG_SHOW_FORMULA_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : RNG_INSERT_FORMULA_FUNC
'DESCRIPTION   : Sums to each column of a Excel range (UsedRange),
'but requires passed variables for the starting column and starting row.
'LIBRARY       : EXCEL
'GROUP         : FORMULAS
'ID            : 003
'LAST UPDATE   : 14/02/2008
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'************************************************************************************
'************************************************************************************

Function RNG_INSERT_FORMULA_FUNC(Optional ByVal SROW As Long = 5, _
Optional ByVal SCOLUMN As Long = 2, _
Optional ByVal FORMULA_STR As String = "Sum", _
Optional ByRef SRC_WSHEET As Excel.Worksheet)

Dim i As Long
Dim j As Long

On Error GoTo ERROR_LABEL

If SRC_WSHEET Is Nothing Then: SRC_WSHEET = ActiveSheet

For j = SCOLUMN To SRC_WSHEET.UsedRange.Columns.COUNT
    i = SRC_WSHEET.Cells(SROW, j).End(xlDown).row
    SRC_WSHEET.Cells(i + 1, j).formula = "=" & FORMULA_STR & "(" & SRC_WSHEET.Cells(SROW, j).Address & ":" & SRC_WSHEET.Cells(i, j).Address & ")"
Next j

RNG_INSERT_FORMULA_FUNC = True

Exit Function
ERROR_LABEL:
RNG_INSERT_FORMULA_FUNC = False
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : RNG_EVALUATE_FORMULA_FUNC
'DESCRIPTION   : The following functions show how to evaluate a function
'expression (given as a string) so it works like a numerical function on the
'argument. Here the function expression is given as named range in a sheet.
'LIBRARY       : EXCEL
'GROUP         : FORMULAS
'ID            : 004
'LAST UPDATE   : 14/02/2008
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'************************************************************************************
'************************************************************************************

Function RNG_EVALUATE_FORMULA_FUNC(ByVal X_DATA_VAL As Double, _
ByVal RNG_STR_NAME As String, _
Optional ByVal VAR_CHR_STR As String = "@x", _
Optional ByRef SRC_WBOOK As Excel.Workbook)

Dim TEMP_STR As String

On Error GoTo ERROR_LABEL

If SRC_WBOOK Is Nothing Then: Set SRC_WBOOK = ActiveWorkbook

' substitute the variable @x by value, value casted to string
TEMP_STR = Evaluate(SRC_WBOOK.Names.Item(RNG_STR_NAME).value)

TEMP_STR = Replace(TEMP_STR, VAR_CHR_STR, CStr(X_DATA_VAL), 1, -1, 0) 'Convert
'the temp value as String then Replace all the variables for the value

RNG_EVALUATE_FORMULA_FUNC = Evaluate(Evaluate("(" & TEMP_STR & ")"))

Exit Function
ERROR_LABEL:
RNG_EVALUATE_FORMULA_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : RNG_REPLACE_FORMULAS_PARAMETERS_FUNC
'DESCRIPTION   :
'LIBRARY       : EXCEL
'GROUP         : FORMULAS
'ID            : 005
'LAST UPDATE   : 14/02/2008
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'************************************************************************************
'************************************************************************************

Function RNG_REPLACE_FORMULAS_PARAMETERS_FUNC(ByVal FORMULA_STR As Variant, _
ByRef VAR_NAME_RNG As Variant, _
ByRef DATA_RNG As Excel.Range, _
ByRef PARAM_RNG As Excel.Range, _
Optional ByVal REFER_VAR_STR As Variant = "param")

'FORMULA_STR = "param1*EXP(-(((x-param2)/param3)^2))+y"
'VAR_NAME_RNG(1, 1) = "x"
'VAR_NAME_RNG(1, 2) = "y"

Dim i As Long
Dim j As Long
Dim k As Long

Dim SROW As Long
Dim SCOLUMN As Long

Dim NSIZE As Long
Dim NROWS As Long 'No Observations
Dim NCOLUMNS As Long 'No Parameters

Dim TEMP1_STR As String
Dim TEMP2_STR As String
Dim TEMP3_STR As String

Dim TEMP_VECTOR As Variant
Dim TEMP_MATRIX As Variant

Dim SRC_WSHEET As Excel.Worksheet

On Error GoTo ERROR_LABEL

'susbtitute parameters-cell

Set SRC_WSHEET = PARAM_RNG.Parent
SROW = PARAM_RNG.Cells(1, 1).row
SCOLUMN = PARAM_RNG.Cells(1, 1).Column
NSIZE = PARAM_RNG.Cells.COUNT

NROWS = DATA_RNG.Rows.COUNT
NCOLUMNS = DATA_RNG.Columns.COUNT

If IsArray(VAR_NAME_RNG) = True Then
    TEMP_VECTOR = VAR_NAME_RNG
    If UBound(TEMP_VECTOR, 2) = 1 Then: _
        TEMP_VECTOR = MATRIX_TRANSPOSE_FUNC(TEMP_VECTOR)
    If NCOLUMNS > UBound(TEMP_VECTOR, 2) Then: _
        NCOLUMNS = UBound(TEMP_VECTOR, 2)
Else
    ReDim TEMP_VECTOR(1 To 1, 1 To 1)
    TEMP_VECTOR(1, 1) = VAR_NAME_RNG
    NCOLUMNS = 1
End If

ReDim TEMP_MATRIX(1 To NROWS, 1 To 1)

FORMULA_STR = CONVERT_FACTORIAL_FORMULA_FUNC(FORMULA_STR, "FACT_FUNC")

For i = 1 To NROWS
    TEMP2_STR = FORMULA_STR
    For j = 1 To NCOLUMNS
        TEMP1_STR = DATA_RNG.Cells(i, j).Address(False, False) _
            'substitute variable-cell
        TEMP2_STR = Replace(TEMP2_STR, TEMP_VECTOR(1, j), _
                TEMP1_STR, 1, -1, 0)
    Next j
    For k = 1 To NSIZE
        If PARAM_RNG.Rows.COUNT <> 1 Then
            TEMP1_STR = SRC_WSHEET.Cells(SROW + k - 1, SCOLUMN).Address(True, True)
        Else
            TEMP1_STR = SRC_WSHEET.Cells(SROW, SCOLUMN + k - 1).Address(True, True)
        End If
        TEMP3_STR = REFER_VAR_STR & CStr(k)
        TEMP2_STR = Replace(TEMP2_STR, TEMP3_STR, TEMP1_STR, 1, -1, 0)
    Next k
    TEMP_MATRIX(i, 1) = "=" & TEMP2_STR
Next i

RNG_REPLACE_FORMULAS_PARAMETERS_FUNC = TEMP_MATRIX

Exit Function
ERROR_LABEL:
RNG_REPLACE_FORMULAS_PARAMETERS_FUNC = Err.number
End Function

Function RNG_GET_WSHEET_FORMULAS_VALUES_FUNC( _
Optional ByRef SRC_WSHEET As Excel.Worksheet)

Dim h As Long
Dim i As Long
Dim j As Long
Dim k As Long

Dim TEMP_STR As String
Dim KEY_STR As String
Dim DELIM_STR As String
Dim ADDRESS_STR As String
Dim FORMULA_STR As String
Dim DATA_ARR() As String
Dim DATA_VECTOR() As String
Dim DCELL As Excel.Range
Dim COLLECTION_OBJ As New clsTypeHash

On Error GoTo ERROR_LABEL

If SRC_WSHEET Is Nothing Then: Set SRC_WSHEET = ActiveSheet
DELIM_STR = "|"

h = SRC_WSHEET.UsedRange.Cells.COUNT
COLLECTION_OBJ.SetSize h 'SIZE
COLLECTION_OBJ.IgnoreCase = False

h = 0
For Each DCELL In SRC_WSHEET.UsedRange
    If DCELL.HasArray Then
        ADDRESS_STR = DCELL.CurrentArray.Address
        FORMULA_STR = DCELL.FormulaArray
        If FORMULA_STR <> "" And ADDRESS_STR <> "" Then: GoSub LOAD_ARRAY
    Else
        If DCELL.HasFormula Then
            ADDRESS_STR = DCELL.Address
            FORMULA_STR = DCELL.formula
            If FORMULA_STR <> "" And ADDRESS_STR <> "" Then: GoSub LOAD_ARRAY
        ElseIf IsEmpty(DCELL.value) = False Then
            ADDRESS_STR = DCELL.Address
            FORMULA_STR = DCELL.value
            If FORMULA_STR <> "" And ADDRESS_STR <> "" Then: GoSub LOAD_ARRAY
        End If
    End If
Next DCELL

ReDim DATA_VECTOR(1 To h, 1 To 2)
For k = 1 To h
    TEMP_STR = DATA_ARR(k)
    i = 1: j = InStr(i, TEMP_STR, DELIM_STR)
    DATA_VECTOR(k, 1) = Trim(Mid(TEMP_STR, i, j - i))
    i = j + 1: j = Len(TEMP_STR)
    DATA_VECTOR(k, 2) = Trim(Mid(TEMP_STR, i, j - i + 1))
Next k

Set COLLECTION_OBJ = Nothing
RNG_GET_WSHEET_FORMULAS_VALUES_FUNC = DATA_VECTOR

Exit Function
'---------------------------------------------------------------------
LOAD_ARRAY:
'---------------------------------------------------------------------
    KEY_STR = ADDRESS_STR & DELIM_STR & FORMULA_STR
    If COLLECTION_OBJ.Exists(KEY_STR) = False Then
        h = h + 1
        Call COLLECTION_OBJ.Add(KEY_STR, CStr(h))
        ReDim Preserve DATA_ARR(1 To h)
        DATA_ARR(h) = KEY_STR
    End If
'---------------------------------------------------------------------
Return
'---------------------------------------------------------------------
ERROR_LABEL:
RNG_GET_WSHEET_FORMULAS_VALUES_FUNC = Err.number
End Function


Function RNG_TRIM_FORMULAS_FUNC(ByRef DATA_RNG As Variant, _
Optional ByVal REFER_STR As String = "DATA_SHEET!", _
Optional ByVal OUTPUT As Integer = 0)

'=IF($C$2=FALSE,"-",DATA_SHEET!I11)  -->  DATA_SHEET!G4
'=IF($C$2=FALSE,"-",DATA_SHEET!G11)  -->  DATA_SHEET!C30
'=IF($C$2=FALSE,"-",DATA_SHEET!G4)   -->  DATA_SHEET!D30
'=IF($C$2=FALSE,"-",DATA_SHEET!M91+DATA_SHEET!M87+DATA_SHEET!M88)
'DATA_SHEET!M91,DATA_SHEET!M87, DATA_SHEET!M88

Dim i As Long
Dim j As Long
Dim k As Long

Dim TEMP_STR As String
Dim DATA_STR As String
Dim TEMP_ARR As Variant
Dim DATA_VECTOR As Variant

On Error GoTo ERROR_LABEL

DATA_VECTOR = DATA_RNG

TEMP_STR = ""
For k = 1 To UBound(DATA_VECTOR, 1)
    DATA_STR = DATA_VECTOR(k, 1)
    i = InStr(1, DATA_STR, REFER_STR)
    Do
        j = i + Len(REFER_STR)
        Do While IsNumeric(Mid(DATA_STR, j, 1)) = False: j = j + 1: Loop
        Do While IsNumeric(Mid(DATA_STR, j, 1)) = True: j = j + 1: Loop
        If TEMP_STR = "" Then
            TEMP_STR = Mid(DATA_STR, i, j - i)
        Else
            TEMP_STR = TEMP_STR & "|" & Mid(DATA_STR, i, j - i)
        End If
        i = InStr(j, DATA_STR, "REFER_STR")
    Loop Until i = 0
Next k
TEMP_ARR = Split(TEMP_STR, "|", -1)
ReDim DATA_VECTOR(1 To UBound(TEMP_ARR), 1 To 1)
For i = 1 To UBound(TEMP_ARR): DATA_VECTOR(i, 1) = TEMP_ARR(i): Next i

Select Case OUTPUT
Case 0
    RNG_TRIM_FORMULAS_FUNC = ARRAY_REMOVE_DUPLICATES_FUNC(DATA_VECTOR, 0)
Case Else
    RNG_TRIM_FORMULAS_FUNC = DATA_VECTOR
End Select

Exit Function
ERROR_LABEL:
RNG_TRIM_FORMULAS_FUNC = Err.number
End Function


Function RNG_PARSE_FORMULAS_FUNC(ByRef DATA_RNG As Excel.Range)

Dim i As Long
Dim j As Long

Dim ii As Long
Dim jj As Long

Dim NROWS As Long
Dim NCOLUMNS As Long
Dim CUMUL_STR As String
Dim TEMP_STR As String
Dim TEMP_VECTOR() As String
On Error GoTo ERROR_LABEL

NROWS = DATA_RNG.Rows.COUNT
NCOLUMNS = DATA_RNG.Columns.COUNT

ReDim TEMP_VECTOR(1 To NROWS, 1 To 1)
For i = 1 To NROWS
    CUMUL_STR = ""
    For j = 1 To NCOLUMNS
        If DATA_RNG.Cells(i, j).HasFormula Then
            TEMP_STR = Trim(DATA_RNG.Cells(i, j).formula)
            jj = Len(TEMP_STR) - 1: ii = jj
            Do While Mid(TEMP_STR, ii, 1) <> ",": ii = ii - 1: Loop
            ii = ii + 1: TEMP_STR = Mid(TEMP_STR, ii, jj - ii)
            If j = 1 Then
                CUMUL_STR = TEMP_STR
            ElseIf j = NCOLUMNS Then
                CUMUL_STR = CUMUL_STR & "|" & TEMP_STR & "|"
            Else
                CUMUL_STR = CUMUL_STR & "|" & TEMP_STR
            End If
        Else
            CUMUL_STR = CUMUL_STR & "" & DATA_RNG.Cells(i, j).Text
        End If
    Next j
    TEMP_VECTOR(i, 1) = CUMUL_STR
Next i
RNG_PARSE_FORMULAS_FUNC = TEMP_VECTOR

Exit Function
ERROR_LABEL:
RNG_PARSE_FORMULAS_FUNC = Err.number
End Function
