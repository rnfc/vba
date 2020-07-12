Attribute VB_Name = "MATRIX_RNG_COPY_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.


'************************************************************************************
'************************************************************************************
'FUNCTION      : RNG_MATRIX_COPY_FUNC
'DESCRIPTION   : The following function can be used for copying several different
'matrix formats in a range: diagonal, triangular, tridiagonal, adjoint, etc.
'LIBRARY       : MATRIX-VECTOR
'GROUP         : RNG_COPY
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/19/2009
'************************************************************************************
'************************************************************************************

Function RNG_MATRIX_COPY_FUNC(ByRef SRC_RNG As Excel.Range, _
ByRef DST_RNG As Excel.Range, _
Optional ByVal FILL_FLAG As Long = 1, _
Optional ByVal VERSION As Variant = 0)

Dim h As Long
Dim hh As Long

Dim i As Long
Dim ii As Long
Dim iii As Long

Dim j As Long
Dim jj As Long

Dim k As Long
Dim kk As Long

Dim R1 As Long
Dim C1 As Long

Dim r2 As Long
Dim C2 As Long

Dim r1s As Long
Dim c1s As Long

Dim r1t As Long
Dim c1t As Long

Dim NSIZE As Long
Dim NCOLUMNS As Long

Dim ATEMP_MATRIX As Variant
Dim BTEMP_MATRIX As Variant

Dim TEMP_RNG As Excel.Range
Dim DST_WSHEET As Excel.Worksheet

On Error GoTo ERROR_LABEL

RNG_MATRIX_COPY_FUNC = False
Set DST_WSHEET = DST_RNG.Worksheet 'sheet target

'coordinate target cell
r1t = DST_RNG.row
c1t = DST_RNG.Column
'coordinate first cell
r1s = SRC_RNG.Areas(1).row
c1s = SRC_RNG.Areas(1).Column

If (FILL_FLAG = 0 And VERSION = "ASIS") _
Or (FILL_FLAG = 0 And VERSION = "TRANSP") _
Or (FILL_FLAG = 0 And VERSION = 0) _
Or (FILL_FLAG = 0 And VERSION = 4) Then
'fill the others (complementary) cells with zero
    NSIZE = SRC_RNG.Areas.COUNT
    With SRC_RNG.Areas(NSIZE)
        r2 = .Rows.COUNT + .row - 1
        C2 = .Columns.COUNT + .Column - 1
    End With
    With DST_WSHEET
      .Range(.Cells(r1t, c1t), .Cells(r1t + r2 - r1s, c1t + C2 - c1s)) = 0
    End With
End If

Select Case VERSION
'---------------------------------------------------------------------------------
Case 0, "ASIS"   'only values
'---------------------------------------------------------------------------------
    
    For Each TEMP_RNG In SRC_RNG.Areas
        R1 = TEMP_RNG.row
        C1 = TEMP_RNG.Column
        ii = r1t + R1 - r1s
        jj = c1t + C1 - c1s
        NSIZE = TEMP_RNG.Rows.COUNT
        NCOLUMNS = TEMP_RNG.Columns.COUNT
        ATEMP_MATRIX = TEMP_RNG
        With DST_WSHEET
            .Range(.Cells(ii, jj), _
            .Cells(ii + NSIZE - 1, jj + NCOLUMNS - 1)) = ATEMP_MATRIX
        End With
    Next TEMP_RNG

'---------------------------------------------------------------------------------
Case 1, "VERT"
'---------------------------------------------------------------------------------
        iii = r1t
        For Each TEMP_RNG In SRC_RNG.Areas
            NSIZE = TEMP_RNG.Cells.COUNT
            ReDim BTEMP_MATRIX(1 To NSIZE, 1 To 1)
            For i = 1 To NSIZE
                BTEMP_MATRIX(i, 1) = TEMP_RNG.Cells(i)
            Next i
            ii = iii
            iii = ii + NSIZE - 1
            With DST_WSHEET
                .Range(.Cells(ii, c1t), .Cells(iii, c1t)) = BTEMP_MATRIX
            End With
            iii = iii + 1
        Next TEMP_RNG

'---------------------------------------------------------------------------------
Case 2, "ORIZ"
'---------------------------------------------------------------------------------
        iii = c1t
        For Each TEMP_RNG In SRC_RNG.Areas
            NSIZE = TEMP_RNG.Cells.COUNT
            ReDim BTEMP_MATRIX(1 To 1, 1 To NSIZE)
            For i = 1 To NSIZE
                BTEMP_MATRIX(1, i) = TEMP_RNG.Cells(i)
            Next
            ii = iii
            iii = ii + NSIZE - 1
            With DST_WSHEET
                .Range(.Cells(r1t, ii), .Cells(r1t, iii)) = BTEMP_MATRIX
            End With
            iii = iii + 1
        Next TEMP_RNG

'---------------------------------------------------------------------------------
Case 3, "DIAG"
'---------------------------------------------------------------------------------
        If FILL_FLAG = 0 Then
            kk = 0
            For Each TEMP_RNG In SRC_RNG.Areas
                kk = kk + TEMP_RNG.Cells.COUNT
            Next
            Range(DST_WSHEET.Cells(r1t, c1t), _
                  DST_WSHEET.Cells(r1t + kk - 1, c1t + kk - 1)) = 0
        End If
        i = r1t: j = c1t
        For Each TEMP_RNG In SRC_RNG.Areas
           For k = 1 To TEMP_RNG.Cells.COUNT
                DST_WSHEET.Cells(i, j) = TEMP_RNG.Cells(k).value
                i = i + 1
                j = j + 1
           Next k
        Next TEMP_RNG

'---------------------------------------------------------------------------------
Case 4, "TRANSP"
'---------------------------------------------------------------------------------
        For Each TEMP_RNG In SRC_RNG.Areas
            NSIZE = TEMP_RNG.Rows.COUNT
            NCOLUMNS = TEMP_RNG.Columns.COUNT
            ATEMP_MATRIX = TEMP_RNG
            ReDim BTEMP_MATRIX(1 To NCOLUMNS, 1 To NSIZE)
            If NSIZE = 1 And NCOLUMNS = 1 Then
                BTEMP_MATRIX(1, 1) = ATEMP_MATRIX
            Else
                For i = 1 To NSIZE
                For j = 1 To NCOLUMNS
                    BTEMP_MATRIX(j, i) = ATEMP_MATRIX(i, j)
                Next j, i
            End If
            R1 = TEMP_RNG.row
            C1 = TEMP_RNG.Column
            ii = r1t + R1 - r1s
            jj = c1t + C1 - c1s
            With DST_WSHEET
                .Range(.Cells(ii, jj), _
                .Cells(ii + NCOLUMNS - 1, jj + NSIZE - 1)) = BTEMP_MATRIX
            End With
        Next TEMP_RNG

'---------------------------------------------------------------------------------
Case 5, "FLIPH"
'---------------------------------------------------------------------------------
    If SRC_RNG.Areas.COUNT = 1 Then
       Set TEMP_RNG = SRC_RNG
        NSIZE = TEMP_RNG.Rows.COUNT
        NCOLUMNS = TEMP_RNG.Columns.COUNT
        ATEMP_MATRIX = TEMP_RNG
        ReDim BTEMP_MATRIX(1 To NSIZE, 1 To NCOLUMNS)
        If NSIZE = 1 And NCOLUMNS = 1 Then
            BTEMP_MATRIX(1, 1) = ATEMP_MATRIX
        Else
            For i = 1 To NSIZE
            For j = 1 To NCOLUMNS
                BTEMP_MATRIX(i, NCOLUMNS + 1 - j) = ATEMP_MATRIX(i, j)
            Next j, i
        End If
        R1 = TEMP_RNG.row
        C1 = TEMP_RNG.Column
        ii = r1t + R1 - r1s
        jj = c1t + C1 - c1s
        With DST_WSHEET
            .Range(.Cells(ii, jj), _
            .Cells(ii + NSIZE - 1, jj + NCOLUMNS - 1)) = BTEMP_MATRIX
        End With
    Else
        GoTo ERROR_LABEL 'Cannot flip multiselection
    End If

'---------------------------------------------------------------------------------
Case 6, "FLIPV"
'---------------------------------------------------------------------------------
    If SRC_RNG.Areas.COUNT = 1 Then
       Set TEMP_RNG = SRC_RNG
        NSIZE = TEMP_RNG.Rows.COUNT
        NCOLUMNS = TEMP_RNG.Columns.COUNT
        ATEMP_MATRIX = TEMP_RNG
        ReDim BTEMP_MATRIX(1 To NSIZE, 1 To NCOLUMNS)
        If NSIZE = 1 And NCOLUMNS = 1 Then
            BTEMP_MATRIX(1, 1) = ATEMP_MATRIX
        Else
            For i = 1 To NSIZE
            For j = 1 To NCOLUMNS
                BTEMP_MATRIX(NSIZE + 1 - i, j) = ATEMP_MATRIX(i, j)
            Next j, i
        End If
        R1 = TEMP_RNG.row
        C1 = TEMP_RNG.Column
        ii = r1t + R1 - r1s
        jj = c1t + C1 - c1s
        With DST_WSHEET
            .Range(.Cells(ii, jj), _
            .Cells(ii + NSIZE - 1, jj + NCOLUMNS - 1)) = BTEMP_MATRIX
        End With
    Else
        GoTo ERROR_LABEL 'Cannot flip multiselection
    End If

'---------------------------------------------------------------------------------
Case Else 'Referring to VERSION as "ADJOINT"
'---------------------------------------------------------------------------------
    If IsEmpty(DST_RNG) Then
        GoTo ERROR_LABEL 'No matrix cell selected
    Else
        With DST_RNG
            ii = .row - .CurrentRegion.row + 1
            jj = .Column - .CurrentRegion.Column + 1
            NSIZE = .CurrentRegion.Rows.COUNT - 1
            NCOLUMNS = .CurrentRegion.Columns.COUNT - 1
        End With
        ATEMP_MATRIX = DST_RNG.CurrentRegion
        If NSIZE = 0 Or NCOLUMNS = 0 Then
            GoTo ERROR_LABEL 'no matrix found
        End If
        ReDim BTEMP_MATRIX(1 To NSIZE, 1 To NCOLUMNS)
        For i = 1 To NSIZE
            If i >= ii Then h = i + 1 Else h = i
            For j = 1 To NCOLUMNS
                If j >= jj Then hh = j + 1 Else hh = j
                BTEMP_MATRIX(i, j) = ATEMP_MATRIX(h, hh)
            Next j
        Next i
        With DST_WSHEET
            .Range(.Cells(r1t, c1t), _
            .Cells(r1t + NSIZE - 1, c1t + NCOLUMNS - 1)) = BTEMP_MATRIX
        End With
        If SRC_RNG.Areas.COUNT = 1 Then: _
            Set SRC_RNG = RNG_MATRIX_SELECT_FUNC(SRC_RNG, 1, 5)
    End If
End Select

RNG_MATRIX_COPY_FUNC = True

Exit Function
ERROR_LABEL:
RNG_MATRIX_COPY_FUNC = False
End Function



'************************************************************************************
'************************************************************************************
'FUNCTION      : RNG_MATRIX_SELECT_FUNC
'DESCRIPTION   : The following function can be used for selecting several different
'matrix formats: diagonal, triangular, tridiagonal, adjoint, etc.
'Simply enter any cell into the matrix and choose the VERSION that
'you want. Automatically the function works on the max area bordered
'by empty cells that usually correspond to the full matrix. If you
'want to restrict the area simply select the sub-matrix that you want.
'LIBRARY       : MATRIX-VECTOR
'GROUP         : RNG_COPY
'ID            : 002
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/19/2009
'************************************************************************************
'************************************************************************************

Private Function RNG_MATRIX_SELECT_FUNC(ByRef SRC_RNG As Excel.Range, _
Optional ByVal SWITCH_FLAG As Boolean = True, _
Optional ByVal VERSION As Long = 2) As Excel.Range

Dim r As Long
Dim c As Long

Dim R0 As Long
Dim C0 As Long

Dim R1 As Long
Dim C1 As Long

Dim r2 As Long
Dim C2 As Long

Dim ra As Long
Dim CA As Long

Dim rb As Long
Dim CB As Long

Dim DST_RNG As Excel.Range
Dim DST_WSHEET As Excel.Worksheet

On Error GoTo ERROR_LABEL

Set DST_WSHEET = SRC_RNG.Worksheet

R1 = SRC_RNG.row
C1 = SRC_RNG.Column
r2 = SRC_RNG.Rows.COUNT + R1 - 1
C2 = SRC_RNG.Columns.COUNT + C1 - 1

Select Case VERSION
'---------------------------------------------------------------------------------
    Case 0 'full current region from left-upper corner
'---------------------------------------------------------------------------------
        Set SRC_RNG = SRC_RNG.Cells(1, 1)
        R0 = SRC_RNG.CurrentRegion.row
        C0 = SRC_RNG.CurrentRegion.Column
        r2 = SRC_RNG.CurrentRegion.Rows.COUNT + R0 - 1
        C2 = SRC_RNG.CurrentRegion.Columns.COUNT + C0 - 1
        Set DST_RNG = Range(DST_WSHEET.Cells(R1, C1), DST_WSHEET.Cells(r2, C2))
'---------------------------------------------------------------------------------
    Case 1 'diagonal
'---------------------------------------------------------------------------------
        If SWITCH_FLAG = True Then 'diagonal main
            r = R1: c = C1
            Set DST_RNG = DST_WSHEET.Cells(r, c)
            Do While r < r2 And c < C2
                r = r + 1: c = c + 1
                Set DST_RNG = Union(DST_RNG, DST_WSHEET.Cells(r, c))
            Loop
        Else  'diagonal secondary
            r = R1: c = C2
            Set DST_RNG = DST_WSHEET.Cells(r, c)
            Do While r < r2 And c > C1
                r = r + 1: c = c - 1
                Set DST_RNG = Union(DST_RNG, DST_WSHEET.Cells(r, c))
            Loop
        End If
'---------------------------------------------------------------------------------
    Case 2  'triangular
'---------------------------------------------------------------------------------
        If SWITCH_FLAG = True Then 'triangular lower
            r = R1: c = C1
            Set DST_RNG = Range(DST_WSHEET.Cells(r, c), DST_WSHEET.Cells(r2, c))
            Do While r < r2 And c < C2
                r = r + 1: c = c + 1
                Set DST_RNG = Union(DST_RNG, _
                    Range(DST_WSHEET.Cells(r, c), DST_WSHEET.Cells(r2, c)))
            Loop
        Else  'triangular upper
            r = R1: c = C1
            Set DST_RNG = Range(DST_WSHEET.Cells(R1, c), DST_WSHEET.Cells(r, c))
            Do While r < r2 And c < C2
                r = r + 1: c = c + 1
                Set DST_RNG = Union(DST_RNG, _
                    Range(DST_WSHEET.Cells(R1, c), DST_WSHEET.Cells(r, c)))
            Loop
        End If
'---------------------------------------------------------------------------------
    Case 3  'triangular without diagonal
'---------------------------------------------------------------------------------
        If SWITCH_FLAG = True Then 'lower
            r = R1 + 1: c = C1
            Set DST_RNG = Range(DST_WSHEET.Cells(r, c), DST_WSHEET.Cells(r2, c))
            Do While r < r2 And c < C2
                r = r + 1: c = c + 1
                Set DST_RNG = Union(DST_RNG, _
                    Range(DST_WSHEET.Cells(r, c), DST_WSHEET.Cells(r2, c)))
            Loop
        Else  'upper
            r = R1: c = C1 + 1
            Set DST_RNG = Range(DST_WSHEET.Cells(R1, c), DST_WSHEET.Cells(r, c))
            Do While r < r2 And c < C2
                r = r + 1: c = c + 1
                Set DST_RNG = Union(DST_RNG, _
                    Range(DST_WSHEET.Cells(R1, c), DST_WSHEET.Cells(r, c)))
            Loop
        End If
'---------------------------------------------------------------------------------
    Case 4   'tridiagonal
'---------------------------------------------------------------------------------
        If SWITCH_FLAG = True Then 'tridiagonal main
            r = R1: c = C1
            Set DST_RNG = Union(DST_WSHEET.Cells(r, c), DST_WSHEET.Cells(r, c + 1))
            Do While r < r2 And c < C2
                r = r + 1: c = c + 1
                Set DST_RNG = Union(DST_RNG, DST_WSHEET.Cells(r, c))
                If c < C2 Then Set DST_RNG = Union(DST_RNG, DST_WSHEET.Cells(r, c + 1))
                If c > 1 Then Set DST_RNG = Union(DST_RNG, DST_WSHEET.Cells(r, c - 1))
            Loop
        Else  'tridiagonal secondary
            r = R1: c = C2
            Set DST_RNG = Union(DST_WSHEET.Cells(r, c - 1), DST_WSHEET.Cells(r, c))
            Do While r < r2 And c > C1
                r = r + 1: c = c - 1
                Set DST_RNG = Union(DST_RNG, DST_WSHEET.Cells(r, c))
                If c > C1 Then Set DST_RNG = Union(DST_RNG, DST_WSHEET.Cells(r, c - 1))
                If c < C2 Then Set DST_RNG = Union(DST_RNG, DST_WSHEET.Cells(r, c + 1))
            Loop
        End If
        
'---------------------------------------------------------------------------------
    Case Else
'---------------------------------------------------------------------------------
        If Not IsEmpty(SRC_RNG) Then
            Set SRC_RNG = SRC_RNG.Cells(1, 1)
            With SRC_RNG
                R0 = .row
                C0 = .Column
                R1 = .CurrentRegion.row
                C1 = .CurrentRegion.Column
                r2 = .CurrentRegion.Rows.COUNT + R1 - 1
                C2 = .CurrentRegion.Columns.COUNT + C1 - 1
            End With
            Set DST_RNG = DST_WSHEET.Cells(R0, C0)
            ra = R1: CA = C1
            rb = R0 - 1: CB = C0 - 1
            If rb >= ra And CB >= CA Then _
                Set DST_RNG = Union(DST_RNG, _
                    Range(DST_WSHEET.Cells(ra, CA), DST_WSHEET.Cells(rb, CB)))
            ra = R0 + 1: CA = C1
            rb = r2: CB = C0 - 1
            If rb >= ra And CB >= CA Then _
                Set DST_RNG = Union(DST_RNG, _
                    Range(DST_WSHEET.Cells(ra, CA), DST_WSHEET.Cells(rb, CB)))
            ra = R1: CA = C0 + 1
            rb = R0 - 1: CB = C2
            If rb >= ra And CB >= CA Then _
                Set DST_RNG = Union(DST_RNG, _
                    Range(DST_WSHEET.Cells(ra, CA), DST_WSHEET.Cells(rb, CB)))
            ra = R0 + 1: CA = C0 + 1
            rb = r2: CB = C2
            If rb >= ra And CB >= CA Then _
                Set DST_RNG = Union(DST_RNG, _
                    Range(DST_WSHEET.Cells(ra, CA), DST_WSHEET.Cells(rb, CB)))
        End If
End Select

Set RNG_MATRIX_SELECT_FUNC = DST_RNG

Exit Function
ERROR_LABEL:
RNG_MATRIX_SELECT_FUNC = Err.number
End Function
