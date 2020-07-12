Attribute VB_Name = "DATE_SERIES_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.


'************************************************************************************
'************************************************************************************

'************************************************************************************
'************************************************************************************
'FUNCTION      : MATCH_DATES_VECTOR1_FUNC
'DESCRIPTION   :
'LIBRARY       : DATE
'GROUP         : SERIES
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/19/2009
'************************************************************************************
'************************************************************************************

Function MATCH_DATES_VECTOR1_FUNC(ByRef DATA_MATRIX As Variant, _
Optional ByVal RESORT_FLAG As Integer = 0)

Dim h As Long
Dim i As Long
Dim j As Long
Dim k As Long
Dim l As Long

Dim hh As Long
Dim ii As Long
Dim jj As Long
Dim kk As Long
Dim ll As Long

Dim NROWS As Long
Dim NCOLUMNS As Long

Dim DELIM_STR As String

Dim DATE_VAL As String
Dim DATA_VAL As String

Dim TEMP_STR As String
Dim LEFT_STR As String
Dim RIGHT_STR As String
Dim LINE_STR As String

Dim DATE_VECTOR As Variant
Dim TEMP_MATRIX As Variant
'Dim DATA_MATRIX As Variant
Dim DATE_VECTOR_OBJ As clsTypeHash

On Error GoTo ERROR_LABEL

Set DATE_VECTOR_OBJ = New clsTypeHash
DATE_VECTOR_OBJ.SetSize 100000
DATE_VECTOR_OBJ.IgnoreCase = False

DELIM_STR = ","

'DATA_MATRIX = DATA_RNG
NROWS = UBound(DATA_MATRIX, 1)
NCOLUMNS = UBound(DATA_MATRIX, 2)

ll = Int(NCOLUMNS / 2)
If NCOLUMNS Mod ll <> 0 Then: GoTo ERROR_LABEL
LINE_STR = DELIM_STR
For ii = 1 To ll
    LINE_STR = LINE_STR & "0" & DELIM_STR
Next ii

hh = 0

jj = 2
Do While jj <= NCOLUMNS

    l = NROWS
    Do While DATA_MATRIX(l, jj - 1) = 0
        l = l - 1
        If l = 0 Then: GoTo 1983
    Loop
    
    Do Until l = 0
        DATE_VAL = DATA_MATRIX(l, jj - 1)
        If DATE_VECTOR_OBJ.Exists(DATE_VAL) = False Then
            TEMP_STR = DATE_VAL & LINE_STR
            Call DATE_VECTOR_OBJ.Add(DATE_VAL, TEMP_STR)
            hh = hh + 1
        End If
        DATA_VAL = Trim(DATA_MATRIX(l, jj))

        If DATA_VAL = "0" Or DATA_VAL = "" Then: DATA_VAL = 0
        TEMP_STR = DATE_VECTOR_OBJ.Item(DATE_VAL)
        i = 1
        j = InStr(i, TEMP_STR, DELIM_STR)
        If j = 0 Then GoTo 1982
        i = j + 1
        ll = Int(jj / 2)
        For kk = 1 To ll
            LEFT_STR = Mid(TEMP_STR, 1, i - 1)
            j = InStr(i, TEMP_STR, DELIM_STR)
            If j = 0 Then GoTo 1982
            RIGHT_STR = Mid(TEMP_STR, j, Len(TEMP_STR) - j + 1)
            i = j + 1
        Next kk
        DATE_VECTOR_OBJ.Item(DATE_VAL) = LEFT_STR & DATA_VAL & RIGHT_STR
1982:
        l = l - 1
    Loop

1983:
jj = jj + 2
Loop

ll = Int(NCOLUMNS / 2)
ReDim TEMP_MATRIX(1 To hh, 1 To ll + 1)
ReDim DATE_VECTOR(1 To hh, 1 To 1)
For ii = 1 To hh
    h = ii - 1
    DATE_VAL = DATE_VECTOR_OBJ.GetKey(h)
    DATE_VECTOR(ii, 1) = CDate(DATE_VAL)
Next ii
DATE_VECTOR = MATRIX_QUICK_SORT_FUNC(DATE_VECTOR, 1, IIf(RESORT_FLAG = True, 0, 1))


For ii = 1 To hh
    TEMP_MATRIX(ii, 1) = DATE_VECTOR(ii, 1)
    DATE_VAL = DATE_VECTOR(ii, 1)
    DATA_VAL = DATE_VECTOR_OBJ.Item(DATE_VAL)
    k = Len(DATA_VAL)
    i = InStr(1, DATA_VAL, DELIM_STR)
    If i = 0 Then: GoTo 1984
    i = i + 1
    jj = 1
    Do
        j = InStr(i, DATA_VAL, DELIM_STR)
        If j = 0 Then: Exit Do
        TEMP_MATRIX(ii, jj + 1) = CDec(Mid(DATA_VAL, i, j - i))
        i = j + 1
        jj = jj + 1
    Loop Until jj > ll
1984:
Next ii

MATCH_DATES_VECTOR1_FUNC = TEMP_MATRIX

Exit Function
ERROR_LABEL:
MATCH_DATES_VECTOR1_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************

'************************************************************************************
'************************************************************************************
'FUNCTION      : MATCH_DATES_VECTOR2_FUNC
'DESCRIPTION   : Dates must be in descending order
'LIBRARY       : DATE
'GROUP         : SERIES
'ID            : 002
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/19/2009
'************************************************************************************
'************************************************************************************

Function MATCH_DATES_VECTOR2_FUNC(ByRef DATA_MATRIX As Variant, _
Optional ByVal RESORT_FLAG As Integer = 0)

Dim h As Long
Dim i As Long
Dim k As Long
Dim j As Long
Dim l As Long

Dim NROWS As Long
Dim NCOLUMNS As Long

Dim TEMP_MATRIX As Variant
'Dim DATA_MATRIX As Variant

On Error GoTo ERROR_LABEL

'DATA_MATRIX = DATA_RNG

NROWS = UBound(DATA_MATRIX, 1)
NCOLUMNS = UBound(DATA_MATRIX, 2)
h = Int(NCOLUMNS / 2)
If NCOLUMNS Mod h <> 0 Then: GoTo ERROR_LABEL
ReDim TEMP_MATRIX(1 To NROWS, 1 To h + 1)

j = 2: k = 2
Do While k <= NCOLUMNS
    i = 1
    l = NROWS
    Do While DATA_MATRIX(l, k - 1) = 0
        l = l - 1
        If l = 0 Then: GoTo 1983
    Loop

    Do Until l = 0
'--------------------------------------------------------------------------------------
        If TEMP_MATRIX(i, 1) = 0 And DATA_MATRIX(l, k - 1) <> 0 Then
'--------------------------------------------------------------------------------------
            TEMP_MATRIX(i, 1) = DATA_MATRIX(l, k - 1)
            TEMP_MATRIX(i, j) = DATA_MATRIX(l, k)
            i = i + 1
            l = l - 1
            GoTo 1982
'--------------------------------------------------------------------------------------
        ElseIf TEMP_MATRIX(i, 1) = DATA_MATRIX(l, k - 1) Then
'--------------------------------------------------------------------------------------
            TEMP_MATRIX(i, j) = DATA_MATRIX(l, k)
            i = i + 1
            l = l - 1
            GoTo 1982
'--------------------------------------------------------------------------------------
        ElseIf TEMP_MATRIX(i, 1) > DATA_MATRIX(l, k - 1) Then
'--------------------------------------------------------------------------------------
                TEMP_MATRIX = MATRIX_ADD_ROWS_FUNC(TEMP_MATRIX, i, 1)
                i = i + 1
                If i <> 1 Then
                    i = i - 1
                    TEMP_MATRIX(i, 1) = DATA_MATRIX(l, k - 1)
                    TEMP_MATRIX(i, j) = DATA_MATRIX(l, k)
                    i = i + 1
                    l = l - 1
                End If
            GoTo 1982
'--------------------------------------------------------------------------------------
        ElseIf TEMP_MATRIX(i, 1) < DATA_MATRIX(l, k - 1) Then 'go to the next
'--------------------------------------------------------------------------------------
        'destination with the same source
            i = i + 1
        End If
1982:
    Loop
1983:
   j = j + 1
   k = k + 2
Loop
'----------------------------------------HOUSE_KEEPING----------------------------------
TEMP_MATRIX = MATRIX_TRIM_FUNC(TEMP_MATRIX, 1, 0)
'---------------------------------------------------------------------------------------

If RESORT_FLAG = 0 Then
    MATCH_DATES_VECTOR2_FUNC = TEMP_MATRIX
Else
    MATCH_DATES_VECTOR2_FUNC = MATRIX_REVERSE_FUNC(TEMP_MATRIX)
End If

Exit Function
ERROR_LABEL:
MATCH_DATES_VECTOR2_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : SANITISED_DATES_VECTOR_FUNC
'DESCRIPTION   :
'LIBRARY       : DATE
'GROUP         : SERIES
'ID            : 003
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/19/2009
'************************************************************************************
'************************************************************************************

Function SANITISED_DATES_VECTOR_FUNC(ByRef DATA_MATRIX As Variant, _
Optional ByVal PERIOD_STR As String = "m", _
Optional ByVal NSIZE As Long = 0, _
Optional ByVal OUTPUT As Integer = 0)

Dim h As Long
Dim i As Long
Dim j As Long
Dim k As Long

Dim SROW As Long
Dim NROWS As Long

Dim SCOLUMN As Long
Dim NCOLUMNS As Long

Dim TEMP_MATRIX As Variant
        
On Error GoTo ERROR_LABEL

'DATA_MATRIX = DATA_RNG
'---------------------------------------------------------------------------------------
If NSIZE = 0 Then
'---------------------------------------------------------------------------------------
    Select Case OUTPUT
    '---------------------------------------------------------------------------------------
    '---------------------------------------------------------------------------------------
    Case 0 'Ascending without headings
    '---------------------------------------------------------------------------------------
        h = 1
        GoSub REDIM1_LINE
        For j = SCOLUMN To NCOLUMNS
            h = 0: i = SROW
            For k = NROWS To (SROW + 1) Step -1
                GoSub NORM1_LINE
                i = i + 1
            Next k
        Next j
    '---------------------------------------------------------------------------------------
    Case 1 'Ascending with headings
    '---------------------------------------------------------------------------------------
        h = 0
        GoSub REDIM1_LINE
        For j = SCOLUMN To NCOLUMNS
            TEMP_MATRIX(SROW, j) = DATA_MATRIX(SROW, j)
            h = 0: i = SROW + 1
            For k = NROWS To (SROW + 1) Step -1
                GoSub NORM1_LINE
                i = i + 1
            Next k
        Next j
    '---------------------------------------------------------------------------------------
    Case 2 'Descending without headings
    '---------------------------------------------------------------------------------------
        h = 1
        GoSub REDIM1_LINE
        For j = SCOLUMN To NCOLUMNS
            h = 1: i = SROW + 1
            For k = SROW + 1 To NROWS Step 1
                GoSub NORM1_LINE
                i = i + 1
            Next k
        Next j
    '---------------------------------------------------------------------------------------
    Case Else 'Descending with headings
    '---------------------------------------------------------------------------------------
        h = 0
        GoSub REDIM1_LINE
        For j = SCOLUMN To NCOLUMNS
            TEMP_MATRIX(SROW, j) = DATA_MATRIX(SROW, j)
            h = 0: i = SROW + 1
            For k = SROW + 1 To NROWS Step 1
                GoSub NORM1_LINE
                i = i + 1
            Next k
        Next j
    '---------------------------------------------------------------------------------------
    End Select
    '---------------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------
Else ' Multiple
'-----------------------------------------------------------------------------------
    Select Case OUTPUT
    '-----------------------------------------------------------------------------------
    Case 0 'Sanitised Matrix without Headings
    '-----------------------------------------------------------------------------------
        h = 1
        GoSub REDIM2_LINE
        h = 0: k = 1
        For j = SCOLUMN To NCOLUMNS Step 2
            GoSub NORM2_LINE
            k = k + 1
        Next j
    '-----------------------------------------------------------------------------------
    Case 1 'Sanitised Matrix with Headings
    '-----------------------------------------------------------------------------------
        h = 0
        GoSub REDIM2_LINE
        h = 1: k = 1
        For j = SCOLUMN To NCOLUMNS Step 2
            TEMP_MATRIX(SROW, j) = "DATE"
            TEMP_MATRIX(SROW, j + 1) = DATA_MATRIX(k, 1)
            GoSub NORM2_LINE
            k = k + 1
        Next j
    '-----------------------------------------------------------------------------------
    Case 2 'Sanitised Matrix & Match Dates without Headings
    '-----------------------------------------------------------------------------------
        h = 1
        GoSub REDIM2_LINE
        h = 0: k = 1
        For j = SCOLUMN To NCOLUMNS Step 2
            GoSub NORM2_LINE
            k = k + 1
        Next j
        TEMP_MATRIX = MATCH_DATES_VECTOR1_FUNC(TEMP_MATRIX, 0)
    '-----------------------------------------------------------------------------------
    Case Else 'Sanitised Matrix & Match Dates with Headings
    '-----------------------------------------------------------------------------------
        h = 1
        GoSub REDIM2_LINE
        h = 0: k = 1
        For j = SCOLUMN To NCOLUMNS Step 2
            GoSub NORM2_LINE
            k = k + 1
        Next j
        TEMP_MATRIX = MATCH_DATES_VECTOR1_FUNC(TEMP_MATRIX, 0)
        TEMP_MATRIX = MATRIX_ADD_ROWS_FUNC(TEMP_MATRIX, 1, 1)
        
        SROW = LBound(TEMP_MATRIX, 1): NROWS = UBound(TEMP_MATRIX, 1)
        SCOLUMN = LBound(TEMP_MATRIX, 2): NCOLUMNS = UBound(TEMP_MATRIX, 2)
        TEMP_MATRIX(SROW, SCOLUMN) = "DATE"
        For j = SCOLUMN + 1 To NCOLUMNS
            TEMP_MATRIX(SROW, j) = DATA_MATRIX(j - 1, 1)
        Next j
    '-----------------------------------------------------------------------------------
    End Select
    '-----------------------------------------------------------------------------------
End If
'-----------------------------------------------------------------------------------

SANITISED_DATES_VECTOR_FUNC = TEMP_MATRIX

Exit Function
'-----------------------------------------------------------------------------------
REDIM1_LINE:
'-----------------------------------------------------------------------------------
    SROW = LBound(DATA_MATRIX, 1)
    NROWS = UBound(DATA_MATRIX, 1)
    SCOLUMN = LBound(DATA_MATRIX, 2)
    NCOLUMNS = UBound(DATA_MATRIX, 2)
    ReDim TEMP_MATRIX(SROW To NROWS - h, SCOLUMN To NCOLUMNS)
'-----------------------------------------------------------------------------------
Return
'-----------------------------------------------------------------------------------
REDIM2_LINE:
'-----------------------------------------------------------------------------------
    SROW = h
    NROWS = NSIZE
    SCOLUMN = 1
    NCOLUMNS = UBound(DATA_MATRIX, 1) * 2
    ReDim TEMP_MATRIX(SROW To NROWS, SCOLUMN To NCOLUMNS)
'-----------------------------------------------------------------------------------
Return
'-----------------------------------------------------------------------------------
NORM1_LINE:
'-----------------------------------------------------------------------------------
    If j <> SCOLUMN Then
        TEMP_MATRIX(i - h, j) = CDec(DATA_MATRIX(k, j))
    Else
        TEMP_MATRIX(i - h, j) = NORMALIZE_DATES_VECTOR_FUNC(CDate(DATA_MATRIX(k, j)), PERIOD_STR)
    End If
'-----------------------------------------------------------------------------------
Return
'-----------------------------------------------------------------------------------
NORM2_LINE:
'-----------------------------------------------------------------------------------
    NROWS = UBound(DATA_MATRIX(k, 2), 1)
    For i = SROW + h To NROWS
        TEMP_MATRIX(i, j) = NORMALIZE_DATES_VECTOR_FUNC(CDate(DATA_MATRIX(k, 2)(i, 1)), PERIOD_STR)
        TEMP_MATRIX(i, j + 1) = CDec(DATA_MATRIX(k, 3)(i, 1))
    Next i
'-----------------------------------------------------------------------------------
Return
'-----------------------------------------------------------------------------------
ERROR_LABEL:
SANITISED_DATES_VECTOR_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : NORMALIZE_DATES_VECTOR_FUNC
'DESCRIPTION   : Normalize Dates
'LIBRARY       : DATE
'GROUP         : SERIES
'ID            : 004
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/19/2009
'************************************************************************************
'************************************************************************************

Function NORMALIZE_DATES_VECTOR_FUNC(ByVal DATE_VAL As Variant, _
Optional ByVal PERIOD_STR As Variant = "m")

On Error GoTo ERROR_LABEL

If IsDate(DATE_VAL) = False Then: GoTo 1983
Select Case PERIOD_STR
Case 0, "m"
   Do While Day(DATE_VAL) <> 1
       DATE_VAL = DateAdd("d", -1, DATE_VAL)
   Loop
Case 1, "w"
   Do While Weekday(DATE_VAL) <> vbMonday
       DATE_VAL = DateAdd("d", -1, DATE_VAL)
   Loop
Case Else '2,"", "d", "v"
End Select

1983:
NORMALIZE_DATES_VECTOR_FUNC = DATE_VAL
    
Exit Function
ERROR_LABEL:
NORMALIZE_DATES_VECTOR_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : FORMAT_DATES_VECTOR_FUNC
'DESCRIPTION   : Format Time Series Dates
'LIBRARY       : DATE
'GROUP         : SERIES
'ID            : 005
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/19/2009
'************************************************************************************
'************************************************************************************

Function FORMAT_DATES_VECTOR_FUNC(ByVal DATE_VAL As Variant, _
Optional ByVal PERIOD_STR As Variant = "m")

On Error GoTo ERROR_LABEL

If IsDate(DATE_VAL) = False Then: GoTo 1983
Select Case PERIOD_STR
Case 0, "m"
    DATE_VAL = Format(DATE_VAL, "mmm-yy")
Case 1, "w"
    DATE_VAL = Format(DATE_VAL, "d-mmm-yy")
Case Else '2,"", "d", "v"
    DATE_VAL = Format(DATE_VAL, "d-mmm-yy")
End Select

1983:
FORMAT_DATES_VECTOR_FUNC = DATE_VAL

Exit Function
ERROR_LABEL:
FORMAT_DATES_VECTOR_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : RNG_FORMAT_DATES_VECTOR_FUNC
'DESCRIPTION   : Format Time Series Rng
'LIBRARY       : DATE
'GROUP         : SERIES
'ID            : 006
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/19/2009
'************************************************************************************
'************************************************************************************


Function RNG_FORMAT_DATES_VECTOR_FUNC(ByRef DATA_RNG As Excel.Range, _
Optional ByVal LABELS_FLAG As Boolean = True, _
Optional ByVal VERSION As Integer = 0)

Dim j As Long

On Error GoTo ERROR_LABEL

RNG_FORMAT_DATES_VECTOR_FUNC = False

If LABELS_FLAG = True Then
'--------------------------------------------------------------------
    With DATA_RNG.Rows(1)
'--------------------------------------------------------------------
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
        .Font.Bold = True
'--------------------------------------------------------------------
        With .Borders(xlEdgeLeft)
            .LineStyle = xlContinuous
            .WEIGHT = xlThin
            .ColorIndex = xlAutomatic
        End With
        With .Borders(xlEdgeTop)
            .LineStyle = xlContinuous
            .WEIGHT = xlThin
            .ColorIndex = xlAutomatic
        End With
        With .Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .WEIGHT = xlThin
            .ColorIndex = xlAutomatic
        End With
        With .Borders(xlEdgeRight)
            .LineStyle = xlContinuous
            .WEIGHT = xlThin
            .ColorIndex = xlAutomatic
        End With

        With .Interior
            .ColorIndex = 36
            .Pattern = xlSolid
        End With
'--------------------------------------------------------------------
    End With
'--------------------------------------------------------------------
End If

'--------------------------------------------------------------------
    With DATA_RNG
'--------------------------------------------------------------------
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
'--------------------------------------------------------------------
    End With

'--------------------------------------------------------------------
Select Case VERSION
Case 0
'--------------------------------------------------------------------
        DATA_RNG.Style = "Comma"
        DATA_RNG.Columns(1).NumberFormat = "m/d/yyyy"
'--------------------------------------------------------------------
Case Else
'--------------------------------------------------------------------
        DATA_RNG.Style = "Comma"
        For j = 1 To DATA_RNG.Columns.COUNT Step 2
            DATA_RNG.Columns(j).NumberFormat = "m/d/yyyy"
        Next j
'--------------------------------------------------------------------
End Select
'--------------------------------------------------------------------

RNG_FORMAT_DATES_VECTOR_FUNC = True

Exit Function
ERROR_LABEL:
RNG_FORMAT_DATES_VECTOR_FUNC = False
End Function
