Attribute VB_Name = "MATRIX_MATCH_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.

'************************************************************************************
'************************************************************************************
'FUNCTION      : MATRIX_MATCH_DATA_FUNC
'DESCRIPTION   : Match and Set the position inside a two-dimension array
'LIBRARY       : MATRIX
'GROUP         : MATCH
'ID            : 007
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/19/2009
'************************************************************************************
'************************************************************************************

Function MATRIX_MATCH_DATA_FUNC(ByRef DATA_RNG As Variant, _
Optional ByVal OUTPUT As Integer = 2)

Dim i As Long
Dim j As Long

Dim SCOLUMN As Long
Dim NCOLUMNS As Long

Dim NSIZE As Long
Dim NO_VAR As Long

Dim TEMP_ARR As Variant
Dim TEMP_VECTOR As Variant
Dim DATA_MATRIX As Variant

Dim LABEL_VECTOR As Variant
Dim POSITION_VECTOR As Variant

On Error GoTo ERROR_LABEL

DATA_MATRIX = DATA_RNG
DATA_MATRIX = _
MATRIX_TRANSPOSE_FUNC(MATRIX_TRIM_FUNC(MATRIX_TRANSPOSE_FUNC(DATA_MATRIX), 1, ""))

LABEL_VECTOR = _
MATRIX_TRANSPOSE_FUNC(MATRIX_GET_ROW_FUNC(DATA_MATRIX, 1, 1)) 'Label Vector

SCOLUMN = LBound(LABEL_VECTOR, 1)
NCOLUMNS = UBound(LABEL_VECTOR, 1)

ReDim POSITION_VECTOR(SCOLUMN To NCOLUMNS, 1 To 3)

For i = SCOLUMN To NCOLUMNS
    TEMP_VECTOR = VECTOR_TRIM_FUNC(MATRIX_GET_COLUMN_FUNC(DATA_MATRIX, i, 1), "")
    NSIZE = (UBound(TEMP_VECTOR, 1) - LBound(TEMP_VECTOR, 1) + 1) - 1 _
    'Exclude Headings
    POSITION_VECTOR(i, 3) = NSIZE
Next i

TEMP_ARR = _
MATRIX_ARRAY_CONVERT_FUNC(MATRIX_REMOVE_ROWS_FUNC(DATA_MATRIX, _
LBound(DATA_MATRIX, 1), 1))

ReDim TEMP_VECTOR(LBound(TEMP_ARR, 1) To UBound(TEMP_ARR, 1), 1 To 1)
For i = LBound(TEMP_ARR, 1) To UBound(TEMP_ARR, 1)
    TEMP_VECTOR(i, 1) = TEMP_ARR(i)
Next i

TEMP_VECTOR = VECTOR_TRIM_FUNC(TEMP_VECTOR, 0)
NO_VAR = UBound(TEMP_VECTOR, 1) - LBound(TEMP_VECTOR, 1) + 1

For i = NCOLUMNS To SCOLUMN Step -1
    If i = NCOLUMNS Then
        POSITION_VECTOR(i, 1) = NO_VAR - POSITION_VECTOR(i, 3) + 1
    Else
        POSITION_VECTOR(i, 1) = POSITION_VECTOR(i + 1, 1) - POSITION_VECTOR(i, 3)
    End If
    POSITION_VECTOR(i, 2) = POSITION_VECTOR(i, 1) + POSITION_VECTOR(i, 3) - 1
Next i

For i = NCOLUMNS To SCOLUMN Step -1
    If POSITION_VECTOR(i, 3) = 0 Then
        j = i
        Do While POSITION_VECTOR(j, 3) = 0
            j = j - 1
        Loop
        POSITION_VECTOR(i, 1) = POSITION_VECTOR(j, 1)
        POSITION_VECTOR(i, 2) = POSITION_VECTOR(j, 2)
        LABEL_VECTOR(i, 1) = LABEL_VECTOR(j, 1)
    End If
Next i

Select Case OUTPUT
    Case 0
        MATRIX_MATCH_DATA_FUNC = Array(POSITION_VECTOR, _
                              LABEL_VECTOR, TEMP_VECTOR, _
                              TEMP_ARR, DATA_MATRIX)
    Case 1
        MATRIX_MATCH_DATA_FUNC = POSITION_VECTOR
    Case 2
        MATRIX_MATCH_DATA_FUNC = LABEL_VECTOR
    Case 3
        MATRIX_MATCH_DATA_FUNC = TEMP_VECTOR
    Case 4
        MATRIX_MATCH_DATA_FUNC = TEMP_ARR
    Case Else
        MATRIX_MATCH_DATA_FUNC = DATA_MATRIX
End Select

Exit Function
ERROR_LABEL:
MATRIX_MATCH_DATA_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : ARRAY_MATCH_DATA_FUNC
'DESCRIPTION   : This function use two custom classes that allow you to store and
'look up multiple data items quickly.

'LIBRARY       : MATRIX
'GROUP         : MATCH
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/19/2009
'************************************************************************************
'************************************************************************************

Function ARRAY_MATCH_DATA_FUNC(ByRef ADATA_ARR As Variant, _
ByRef BDATA_ARR As Variant)

Dim i As Long
Dim j As Long
Dim k As Long

Dim DATA_OBJ As New Collection
'Collections are very useful, but you can only store one item, and you
'can't store 'types'.

Dim TEMP_VAL As Variant
Dim TEMP_ARR() As Variant
'Dim T_VAL As Single

On Error GoTo ERROR_LABEL

k = 1
For i = LBound(BDATA_ARR) To UBound(BDATA_ARR)
  DATA_OBJ.Add CStr(k), CStr(BDATA_ARR(i))
  'first item is array sequence , s
  'econd is string we lookup on both must be strings
   k = k + 1
Next i

'search
'set error trapping on
On Error Resume Next

ReDim TEMP_ARR(1 To 1)
j = 0: k = 1
For i = LBound(ADATA_ARR) To UBound(ADATA_ARR)
  'next line looks up the item from ADATA_ARR in the collection, and
  'converts the result to a number. If no error, this gives us the
  'sequence number, eg if the number is 45, it means B2(45,1) is a match
  'If you get an error there was no match
  TEMP_VAL = DATA_OBJ(CStr(ADATA_ARR(k)))
  If Err = 0 Then 'no error, we have a match
    j = j + 1
    ReDim Preserve TEMP_ARR(1 To j)
    TEMP_ARR(j) = TEMP_VAL
  Else 'no match, reset error
    Err = 0
  End If
  k = k + 1
Next i

'report matches and time
'T_VAL = Timer - T_VAL 'freeze timer
'MsgBox "I found " & j & " matches in " & Format(T_VAL, "0.00") & _
" seconds", vbInformation, "Method 1"
ARRAY_MATCH_DATA_FUNC = TEMP_ARR



'------------------------------------------------------------------------------

Exit Function
ERROR_LABEL:
ARRAY_MATCH_DATA_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : MATRIX_MATCH_COUNT_FUNC
'DESCRIPTION   : Count number of times an entry appear in a matrix
'LIBRARY       : MATRIX
'GROUP         : MATCH
'ID            : 002
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/19/2009
'************************************************************************************
'************************************************************************************

Function MATRIX_MATCH_COUNT_FUNC(ByRef DATA_RNG As Variant, _
ByVal REF_VAL As Variant, _
Optional ByVal METHOD As Integer = 0, _
Optional ByVal SCOLUMN As Long = 1, _
Optional ByVal SROW As Long = 1)

Dim kk As Long
Dim DATA_MATRIX As Variant

On Error GoTo ERROR_LABEL

DATA_MATRIX = DATA_RNG

kk = 0

Do Until SCOLUMN > UBound(DATA_MATRIX, 2)
      If METHOD = 0 Then
        If DATA_MATRIX(SROW, SCOLUMN) Like REF_VAL Then _
        kk = kk + 1
      Else
        If Not DATA_MATRIX(SROW, SCOLUMN) Like REF_VAL Then _
        kk = kk + 1
      End If

      SCOLUMN = SCOLUMN + 1
Loop

MATRIX_MATCH_COUNT_FUNC = kk

Exit Function
ERROR_LABEL:
MATRIX_MATCH_COUNT_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : MATRIX_MATCH_WITHOUT_REPLACEMENT_FUNC (DYNAMIC_LOADING_WO)
'DESCRIPTION   : LOAD AN ARRAY BASED ON A REFERENCE VECTOR --> WITHOUT REPLACEMENT
'LIBRARY       : MATRIX
'GROUP         : MATCH
'ID            : 003
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/19/2009
'************************************************************************************
'************************************************************************************

Function MATRIX_MATCH_WITHOUT_REPLACEMENT_FUNC(ByRef DATA_RNG As Variant, _
ByRef REFER_RNG As Variant)

Dim ii As Long
Dim jj As Long

Dim iii As Long
Dim jjj As Long

Dim NROWS As Long
Dim NCOLUMNS As Long

Dim DATA_MATRIX As Variant
Dim REFER_VECTOR As Variant
Dim TEMP_MATRIX As Variant

Dim TEMP_FLAG As Boolean

On Error GoTo ERROR_LABEL

DATA_MATRIX = DATA_RNG
REFER_VECTOR = REFER_RNG
If UBound(REFER_VECTOR, 2) = 1 Then
    REFER_VECTOR = MATRIX_TRANSPOSE_FUNC(REFER_VECTOR)
End If

iii = 1
jjj = 0

NROWS = UBound(DATA_MATRIX, 1)
NCOLUMNS = UBound(DATA_MATRIX, 2)

TEMP_FLAG = False

ReDim TEMP_MATRIX(1 To NROWS, 1 To NCOLUMNS)

Do Until iii > NCOLUMNS
    
    TEMP_FLAG = False
    
        For jj = 1 To UBound(REFER_VECTOR, 2)
            If REFER_VECTOR(1, jj) Like DATA_MATRIX(1, iii) Then
                TEMP_FLAG = True
                Exit For
            End If
        Next jj
        
        If TEMP_FLAG = True Then
           jjj = jjj + 1
           ReDim Preserve TEMP_MATRIX(1 To NROWS, 1 To jjj)
           For ii = 1 To NROWS
                TEMP_MATRIX(ii, jjj) = DATA_MATRIX(ii, iii)
            Next ii
        End If
    
    iii = iii + 1

Loop
    
MATRIX_MATCH_WITHOUT_REPLACEMENT_FUNC = TEMP_MATRIX

Exit Function
ERROR_LABEL:
MATRIX_MATCH_WITHOUT_REPLACEMENT_FUNC = Err.number
End Function



'************************************************************************************
'************************************************************************************
'FUNCTION      : MATRIX_MATCH_WITH_REPLACEMENT_FUNC (DYNAMIC_LOADING_WI)

'DESCRIPTION   : LOAD AN ARRAY BASED ON A REFERENCE VECTOR --> WITH REPLACEMENT
'FIRST ROW OF DATA_RNG MUST HAVE THE TICKET (HEADING); DATA_RNG CANNOT HAVE
'REPEATED TICKET; REFERENCE TICKET MUST BE UNIQUE

'LIBRARY       : MATRIX
'GROUP         : MATCH
'ID            : 004
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/19/2009
'************************************************************************************
'************************************************************************************

Function MATRIX_MATCH_WITH_REPLACEMENT_FUNC(ByRef DATA_RNG As Variant, _
ByRef REFER_RNG As Variant)

Dim ii As Long
Dim jj As Long

Dim iii As Long
Dim jjj As Long

Dim NROWS As Long
Dim NCOLUMNS As Long

Dim TEMP_VAL As Variant

Dim REFER_VECTOR As Variant
Dim DATA_MATRIX As Variant
Dim TEMP_MATRIX As Variant

On Error GoTo ERROR_LABEL

DATA_MATRIX = DATA_RNG
REFER_VECTOR = REFER_RNG
If UBound(REFER_VECTOR, 2) = 1 Then: _
    REFER_VECTOR = MATRIX_TRANSPOSE_FUNC(REFER_VECTOR)

iii = 1
jjj = 0

NROWS = UBound(DATA_MATRIX, 1)
NCOLUMNS = UBound(DATA_MATRIX, 2)

ReDim TEMP_MATRIX(1 To NROWS, 1 To NCOLUMNS)

Do Until iii > UBound(REFER_VECTOR, 2)
    TEMP_VAL = REFER_VECTOR(1, iii)
    jj = MATRIX_FIND_ELEMENT_FUNC(DATA_MATRIX, TEMP_VAL, 1, 1, 0)
    If jj <> 0 Then
        jjj = jjj + 1 'COUNTER FOR THE COLUMN
        ReDim Preserve TEMP_MATRIX(1 To NROWS, 1 To jjj)
        For ii = 1 To NROWS
            TEMP_MATRIX(ii, jjj) = DATA_MATRIX(ii, jj)
        Next ii
    End If
    iii = iii + 1
Loop

MATRIX_MATCH_WITH_REPLACEMENT_FUNC = TEMP_MATRIX

Exit Function
ERROR_LABEL:
MATRIX_MATCH_WITH_REPLACEMENT_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : MATRIX_MATCH_EXCLUDE_FUNC (EXCLUDE_VARIABLES)
'DESCRIPTION   : LOAD RANGE EXCLUDING REFERENCE VARIABLES(SUBTRACTING COLUMN METHOD)
'LIBRARY       : MATRIX
'GROUP         : MATCH
'ID            : 005
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/19/2009
'************************************************************************************
'************************************************************************************

Function MATRIX_MATCH_EXCLUDE_FUNC(ByRef DATA_RNG As Variant, _
ByRef REFER_RNG As Variant, _
Optional ByVal REFER_STR As Variant = "INDEPENDENT")

'REFER_RNG --> Must be "Dependent" or "Independent" String Vector

Dim jj As Long
Dim kk As Long

Dim NSIZE As Long
Dim NROWS As Long
Dim NCOLUMNS As Long

Dim TEMP_VECTOR As Variant
Dim TEMP_MATRIX As Variant

Dim DATA_MATRIX As Variant
Dim REFER_VECTOR As Variant

On Error GoTo ERROR_LABEL

DATA_MATRIX = DATA_RNG
REFER_VECTOR = REFER_RNG
If UBound(REFER_VECTOR, 2) = 1 Then
    REFER_VECTOR = MATRIX_TRANSPOSE_FUNC(REFER_VECTOR)
End If

If UBound(DATA_MATRIX, 2) <> UBound(REFER_VECTOR, 2) Then: GoTo ERROR_LABEL

NROWS = UBound(DATA_MATRIX, 1)
NCOLUMNS = UBound(DATA_MATRIX, 2)
NSIZE = MATRIX_MATCH_COUNT_FUNC(REFER_VECTOR, REFER_STR, 0, 1, 1)

kk = 0
jj = 1

TEMP_MATRIX = DATA_MATRIX

Do
    If REFER_VECTOR(1, jj) Like REFER_STR Then
        TEMP_VECTOR = MATRIX_REMOVE_COLUMNS_FUNC(TEMP_MATRIX, jj - kk, 1)
        TEMP_MATRIX = TEMP_VECTOR
        kk = kk + 1
    End If
    jj = jj + 1
Loop Until kk = NSIZE

MATRIX_MATCH_EXCLUDE_FUNC = TEMP_MATRIX

Exit Function
ERROR_LABEL:
MATRIX_MATCH_EXCLUDE_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : MATRIX_MATCH_RE_ARRANGE_FUNC (RE_ARRANGE_ENTRIES)
'DESCRIPTION   : RE-ARRANGE ENTRIES IN MATRIX USING A REFERENCE VALUE AS A BASE
'LIBRARY       : MATRIX
'GROUP         : MATCH
'ID            : 006
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/19/2009
'************************************************************************************
'************************************************************************************

Function MATRIX_MATCH_RE_ARRANGE_FUNC(ByRef DATA_RNG As Variant, _
Optional ByVal REF_VAL As Variant = 0)

Dim hh As Long
Dim ii As Long
Dim jj As Long
Dim kk As Long
Dim ll As Long

Dim NROWS As Long
Dim NCOLUMNS As Long

Dim TEMP_MATRIX As Variant
Dim TEMP_VECTOR As Variant
Dim DATA_MATRIX As Variant

On Error GoTo ERROR_LABEL

DATA_MATRIX = DATA_RNG

NROWS = UBound(DATA_MATRIX, 1)
NCOLUMNS = UBound(DATA_MATRIX, 2)

ReDim TEMP_MATRIX(1 To NROWS, 1 To NCOLUMNS)

For jj = 1 To NCOLUMNS
      For ii = 1 To jj
            hh = 0
            ReDim TEMP_VECTOR(1 To NROWS, 1 To 2)
            For kk = 1 To NROWS
                If (DATA_MATRIX(kk, ii) = REF_VAL) Then
                    hh = hh + 1
                Else
                    TEMP_VECTOR(kk - hh, 1) = DATA_MATRIX(kk, jj)
                    TEMP_VECTOR(kk - hh, 2) = DATA_MATRIX(kk, ii)
                End If
            Next kk
            
            'ReDim Preserve TEMP_MATRIX(1 To NROWS, 1 To NCOLUMNS)
            'Re-Arrange Entries
            For ll = 1 To NROWS - hh
                TEMP_MATRIX(ll, jj) = TEMP_VECTOR(ll, 2)
            Next ll
      Next ii
Next jj

MATRIX_MATCH_RE_ARRANGE_FUNC = TEMP_MATRIX

Exit Function
ERROR_LABEL:
MATRIX_MATCH_RE_ARRANGE_FUNC = Err.number
End Function


Function VECTOR_MATCH_DATA_FUNC(ByRef DATA_RNG As Variant, _
Optional ByVal REFER_STR As String = "NAME*")

' <<< Change "NAME*" to the cell text that indicates
' the start of a new block. The code uses LIKE so
' wildcards are allowed.

Dim h As Long
Dim i As Long
Dim j As Long
Dim k As Long

Dim NROWS As Long

Dim DATA_VECTOR As Variant
Dim TEMP_MATRIX As Variant

On Error GoTo ERROR_LABEL

DATA_VECTOR = DATA_RNG
If UBound(DATA_VECTOR, 1) = 1 Then
    DATA_VECTOR = MATRIX_TRANSPOSE_FUNC(DATA_VECTOR)
End If
NROWS = UBound(DATA_VECTOR, 1)

'The first item in the list must conform to REFER_STR '" + REFER_STR & "'."
If Not UCase(DATA_VECTOR(1, 1)) Like UCase(REFER_STR) Then: GoTo ERROR_LABEL

h = 0: i = 0: j = 0
For k = 1 To NROWS
    If UCase(DATA_VECTOR(k, 1)) Like UCase(REFER_STR) Then
        j = j + 1
        If i > h Then: h = i
        i = 0
    Else
        i = i + 1
    End If
Next k
If i > h Then: h = i

ReDim TEMP_MATRIX(0 To h, 1 To j)

i = 0: j = 0
For k = 1 To NROWS
    If UCase(DATA_VECTOR(k, 1)) Like UCase(REFER_STR) Then
        i = 0
        j = j + 1
        TEMP_MATRIX(0, j) = DATA_VECTOR(k, 1)
    Else
        i = i + 1
        TEMP_MATRIX(i, j) = DATA_VECTOR(k, 1)
    End If
Next k

VECTOR_MATCH_DATA_FUNC = TEMP_MATRIX

Exit Function
ERROR_LABEL:
VECTOR_MATCH_DATA_FUNC = Err.number
End Function
