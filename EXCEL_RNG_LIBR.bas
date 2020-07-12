Attribute VB_Name = "EXCEL_RNG_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.

Private PUB_VALID_RNG_ARR(0 To 1) As Boolean

'************************************************************************************
'************************************************************************************
'FUNCTION      : RNG_DELETE_BLANK_ROWS_FUNC
'DESCRIPTION   : Delete blank rows
'LIBRARY       : EXCEL
'GROUP         : BLANK
'ID            : 001
'LAST UPDATE   : 15/02/2008
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'************************************************************************************
'************************************************************************************

Function RNG_DELETE_BLANK_ROWS_FUNC(ByRef DATA_RNG As Excel.Range)

Dim i As Long
Dim j As Long

Dim NROWS As Long
Dim NCOLUMNS As Long

Dim TEMP_FLAG As Boolean

On Error GoTo ERROR_LABEL

NROWS = DATA_RNG.Rows.COUNT
NCOLUMNS = DATA_RNG.Columns.COUNT

i = 1 'initialize for first row
Do While i < NROWS
    TEMP_FLAG = False 'initialize for each row
    j = 1 'initialize for first column
    Do While IsEmpty(DATA_RNG(i, j).value) = True _
        And j < NCOLUMNS
        j = j + 1
    Loop
    
    If IsEmpty(DATA_RNG(i, j).value) = True And _
        j = NCOLUMNS Then TEMP_FLAG = True
        'Means all the columns in the row were blank so the row is empty
    If TEMP_FLAG = True Then
            DATA_RNG.Rows(i).Delete 'this row is to be deleted
            NROWS = NROWS - 1 'this shrinks the size of the range
    Else
        i = i + 1
    End If
Loop

RNG_DELETE_BLANK_ROWS_FUNC = True

Exit Function
ERROR_LABEL:
RNG_DELETE_BLANK_ROWS_FUNC = False
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : RNG_CHECK_BLANKS_FUNC
'DESCRIPTION   : The following routines check for empty cells within a dataset
'If the application finds an empty cell, you will have to enter a numerical value
'LIBRARY       : EXCEL
'GROUP         : VALIDATION
'ID            : 003
'LAST UPDATE   : 15/02/2008
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'************************************************************************************
'************************************************************************************

Function RNG_CHECK_BLANKS_FUNC(ByRef DATA_RNG() As Excel.Range, _
Optional ByVal START_ROW As Long = 2, _
Optional ByVal AROW As Long = 1, _
Optional ByVal SROW As Long = 2, _
Optional ByVal ACOLUMN As Long = 1, _
Optional ByVal SCOLUMN As Long = 2)

Dim MSG_STR As String

On Error GoTo ERROR_LABEL

If DATA_RNG(START_ROW, ACOLUMN) = "" Then
    MsgBox "Please insert/download data first.", _
    vbInformation, "WARNING!"
    Exit Function
End If

MSG_STR = _
"The following routine checks for empty cells within the dataset." _
& " If the algorithm finds an empty cell, you will have to enter a" _
& " numerical value."

MsgBox MSG_STR, vbExclamation, "DATA VALIDATION"

Call RNG_VALIDATE_CELLS_FUNC(DATA_RNG(), AROW, SROW, ACOLUMN, SCOLUMN, "")

Exit Function
ERROR_LABEL:
PUB_VALID_RNG_ARR(1) = False
MsgBox "There was an error, please reload the application", _
vbInformation, "DATA VALIDATION"
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : RNG_VALIDATE_CELLS_FUNC
'DESCRIPTION   : The following routine launch an input box to fill empty cells
'LIBRARY       : EXCEL
'GROUP         : VALIDATION
'ID            : 004
'LAST UPDATE   : 15/02/2008
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'************************************************************************************
'************************************************************************************

Function RNG_VALIDATE_CELLS_FUNC(ByRef DATA_RNG() As Excel.Range, _
Optional ByVal AROW As Long = 1, _
Optional ByVal SROW As Long = 2, _
Optional ByVal ACOLUMN As Long = 1, _
Optional ByVal SCOLUMN As Long = 2, _
Optional ByVal REFER_STR As String = "")

Dim i As Long
Dim j As Long

Dim NROWS As Long
Dim NCOLUMNS As Long

Dim TEMP_VAL As Variant
Dim TEMP_RNG As Excel.Range

On Error GoTo ERROR_LABEL

PUB_VALID_RNG_ARR(1) = False

NROWS = UBound(DATA_RNG(), 1)
NCOLUMNS = UBound(DATA_RNG(), 2)


For j = SCOLUMN To NCOLUMNS
    For i = SROW To NROWS
        Set TEMP_RNG = DATA_RNG(i, j)
            
            If (TEMP_RNG.value = REFER_STR) Or _
                (TEMP_RNG.value = 0) Or _
                (Trim(TEMP_RNG.Text) = "") Or _
                (IsNumeric(TEMP_RNG) = False) Then
                    
                    TEMP_RNG.Activate
                    
                    TEMP_VAL = InputBox("Type an adjusted closing price.", _
                    "Symbol <" & DATA_RNG(AROW, j) & " : " & _
                    Format(DATA_RNG(i, ACOLUMN), "dd/mmm/yyyy") & ">")
                        
                    If (IsNumeric(TEMP_VAL) = False) Or _
                    (IsMissing(TEMP_VAL) = True) Or (TEMP_VAL = 0) Then
                        GoTo ERROR_LABEL
                    Else: TEMP_RNG = TEMP_VAL
                    End If
            End If
    Next i
Next j

PUB_VALID_RNG_ARR(1) = True

MsgBox "The dataset has been validated. Feel free to continue working with" _
& " the dataset.", vbExclamation, "DATA VALIDATION"

Exit Function
ERROR_LABEL:

PUB_VALID_RNG_ARR(1) = False

MsgBox "There was an error adjusting the closing prices," & _
" please reload the application", vbInformation, _
"Error adjusting data -  <numerical values only>"

End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : RNG_VALIDATE_CELLS_FLAG_FUNC
'DESCRIPTION   : Data Validation Flag - In case of an empty cell
'LIBRARY       : EXCEL
'GROUP         : VALIDATION
'ID            : 005
'LAST UPDATE   : 15/02/2008
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'************************************************************************************
'************************************************************************************

Function RNG_VALIDATE_CELLS_FLAG_FUNC()

Dim TEMP_OBJ As Variant
Dim STYLE_STR As String
Dim TITLE_STR As String
Dim MSG_STR As String

On Error GoTo ERROR_LABEL  ' Enable error-handling routine

MSG_STR = "Please validate dataset."
STYLE_STR = vbCritical
TITLE_STR = "Warning!"

    If PUB_VALID_RNG_ARR(1) = False Then
        TEMP_OBJ = MsgBox(MSG_STR, STYLE_STR, TITLE_STR)
        RNG_VALIDATE_CELLS_FLAG_FUNC = False
        Exit Function
    End If

RNG_VALIDATE_CELLS_FLAG_FUNC = True

Exit Function
ERROR_LABEL:
RNG_VALIDATE_CELLS_FLAG_FUNC = False
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : RNG_REMOVE_REPEATED_RECORDS_FUNC

'DESCRIPTION   : This function uses Advanced Filter to remove duplicate
'records from the rows spanned by DATA_RNG. A row is considered to
' be a duplicate of another row if the columns spanned by ColumnRangeOfDuplictes
' are equal. Columns outside of those spanned by DATA_RNG
' are not tested.  The function returns the number of rows deleted, including
' 0 if there were no duplicates, or -1 if an error occurred, such as a
' protected sheet or a DATA_RNG range with multiple areas.
' Note that Advanced Filter considers the first row to be the header row
' of the data, so it will never be deleted.

'LIBRARY       : EXCEL
'GROUP         : DELETE
'ID            : 006
'LAST UPDATE   : 15/02/2008
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'************************************************************************************
'************************************************************************************

Function RNG_REMOVE_REPEATED_RECORDS_FUNC(ByRef DATA_RNG As Excel.Range)

Dim i As Long
Dim j As Long

Dim TEMP_RNG As Excel.Range
Dim DELETE_RNG As Excel.Range

On Error GoTo ERROR_LABEL

' Allow only one area.
If DATA_RNG.Areas.COUNT > 1 Then
    RNG_REMOVE_REPEATED_RECORDS_FUNC = -1
    Exit Function
End If

If DATA_RNG.Worksheet.ProtectContents = True Then
    RNG_REMOVE_REPEATED_RECORDS_FUNC = -1
    Exit Function
End If

' Change application settings for speed.
i = DATA_RNG.Rows.COUNT

DATA_RNG.AdvancedFilter Action:=xlFilterInPlace, Unique:=True
'AutoFilter the range.

' Loop through and build a range of hidden rows.
For Each TEMP_RNG In DATA_RNG
    If TEMP_RNG.EntireRow.Hidden = True Then
        If DELETE_RNG Is Nothing Then
            Set DELETE_RNG = TEMP_RNG.EntireRow
        Else
            Set DELETE_RNG = Excel.Application.Union(DELETE_RNG, TEMP_RNG.EntireRow)
        End If
    End If
Next TEMP_RNG
'''''''''''''''''''''''''
' Delete the hidden rows.
'''''''''''''''''''''''''
DELETE_RNG.Delete Shift:=xlUp
'''''''''''''''''''''''''
' Turn off the filter.
'''''''''''''''''''''''''
DATA_RNG.Worksheet.ShowAllData
j = DATA_RNG.Rows.COUNT

RNG_REMOVE_REPEATED_RECORDS_FUNC = True 'i - j

Exit Function:
ERROR_LABEL:
RNG_REMOVE_REPEATED_RECORDS_FUNC = False
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : RNG_FILL_SET_ARR_FUNC
'DESCRIPTION   : Fill and set an array with multiples ranges
'LIBRARY       : EXCEL
'GROUP         : LOADING
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'************************************************************************************
'************************************************************************************

Function RNG_FILL_SET_ARR_FUNC(ByRef SRC_RNG As Excel.Range, _
ByRef DST_RNG() As Excel.Range)

Dim i As Long
Dim j As Long

Dim NROWS As Long
Dim NCOLUMNS As Long

On Error GoTo ERROR_LABEL

RNG_FILL_SET_ARR_FUNC = False

NROWS = SRC_RNG.Rows.COUNT
NCOLUMNS = SRC_RNG.Columns.COUNT
        
ReDim DST_RNG(1 To NROWS, 1 To NCOLUMNS)
For j = 1 To NCOLUMNS
    For i = 1 To NROWS
        Set DST_RNG(i, j) = SRC_RNG.Cells(i, j)
    Next i
Next j

RNG_FILL_SET_ARR_FUNC = True

Exit Function
ERROR_LABEL:
RNG_FILL_SET_ARR_FUNC = False
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : RNG_EXPAND_RNG_FUNC
'DESCRIPTION   : Expand range with current region and Special Cells
'LIBRARY       : EXCEL
'GROUP         : OFFSET
'ID            : 002
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA

'************************************************************************************
'************************************************************************************


Function RNG_EXPAND_RNG_FUNC(ByRef START_POS As Excel.Range, _
Optional ByVal VERSION As Long = 0) As Excel.Range

Dim FIRST_POSITION As Excel.Range
Dim LAST_POSITION As Excel.Range

Dim TEMP_RANGE As Excel.Range

On Error GoTo ERROR_LABEL

Select Case VERSION
    Case 0 'USE THIS OPTION ONLY IF YOU FIRST CLEAN THE ENTIRE SHEET
        Set FIRST_POSITION = START_POS
        Set LAST_POSITION = FIRST_POSITION.SpecialCells(xlCellTypeLastCell)

        If FIRST_POSITION.row < LAST_POSITION.row Then
            Set TEMP_RANGE = Range(FIRST_POSITION, LAST_POSITION)
        End If
        Set RNG_EXPAND_RNG_FUNC = TEMP_RANGE
    Case Else
        Set RNG_EXPAND_RNG_FUNC = START_POS.CurrentRegion
End Select

Exit Function
ERROR_LABEL:
        Set RNG_EXPAND_RNG_FUNC = Nothing
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : RNG_RESIZE_RNG_FUNC
'DESCRIPTION   : Resize range
'LIBRARY       : EXCEL
'GROUP         : OFFSET
'ID            : 004
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA

'************************************************************************************
'************************************************************************************

Function RNG_RESIZE_RNG_FUNC(ByRef SRC_RNG As Excel.Range, _
Optional ByVal SROW As Long = 0, _
Optional ByVal SCOLUMN As Long = 0) As Excel.Range

Dim NROWS As Long
Dim NCOLUMNS As Long

On Error GoTo ERROR_LABEL

NROWS = SRC_RNG.Rows.COUNT
NCOLUMNS = SRC_RNG.Columns.COUNT

Set RNG_RESIZE_RNG_FUNC = SRC_RNG.Resize(NROWS - SROW, NCOLUMNS - SCOLUMN)

Exit Function
ERROR_LABEL:
Set RNG_RESIZE_RNG_FUNC = Nothing
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : RNG_SELECT_BOX_FUNC
'DESCRIPTION   : Set Box Range
'LIBRARY       : EXCEL
'GROUP         : OFFSET
'ID            : 006
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA

'************************************************************************************
'************************************************************************************

Function RNG_SELECT_BOX_FUNC(ByRef SRC_RNG As Excel.Range) As Excel.Range

Dim SROW As Long
Dim SCOLUMN As Long

Dim NROWS As Long
Dim NCOLUMNS As Long

Dim TEMP_RNG As Excel.Range

On Error GoTo ERROR_LABEL

    If SRC_RNG.Cells.COUNT = 1 Then
        Set TEMP_RNG = SRC_RNG
        With TEMP_RNG
            SROW = .row
            SCOLUMN = .Column
            NROWS = .CurrentRegion.Rows.COUNT + .CurrentRegion.row - 1
            NCOLUMNS = .CurrentRegion.Columns.COUNT + .CurrentRegion.Column - 1
        End With
    Else
       Set TEMP_RNG = SRC_RNG
        With TEMP_RNG
            SROW = .row
            SCOLUMN = .Column
            NROWS = SROW + .Rows.COUNT - 1
            NCOLUMNS = SCOLUMN + .Columns.COUNT - 1
        End With
    End If
   
   Set RNG_SELECT_BOX_FUNC = TEMP_RNG

Exit Function
ERROR_LABEL:
Set RNG_SELECT_BOX_FUNC = Nothing
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : RNG_SNAPSHOT_FUNC
'DESCRIPTION   : Snapshot of a range
'LIBRARY       : EXCEL
'GROUP         : PICTURE
'ID            : 001



'AUTHOR        : RAFAEL NICOLAS FERMIN COTA

'************************************************************************************
'************************************************************************************

Function RNG_SNAPSHOT_FUNC(ByRef SRC_RNG As Excel.Range, _
ByRef DST_RNG As Excel.Range)

On Error GoTo ERROR_LABEL
    
    RNG_SNAPSHOT_FUNC = False
    SRC_RNG.CopyPicture Appearance:=xlScreen, Format:=xlPicture
    DST_RNG.Activate
    DST_RNG.Worksheet.Paste
    Excel.Application.CutCopyMode = False
    RNG_SNAPSHOT_FUNC = True

Exit Function
ERROR_LABEL:
RNG_SNAPSHOT_FUNC = False
End Function

'Subroutine to insert/update a picture within a cell based on cell content

Function RNG_INSERT_PICTURE_FUNC(ByVal SRC_RNG As Excel.Range, _
Optional ByRef SRC_WSHEET As Excel.Worksheet)
    
Dim TEMP_VAL As Double
Dim TEMP_CELL As Excel.Range

On Error GoTo ERROR_LABEL

    RNG_INSERT_PICTURE_FUNC = False
    
    If SRC_WSHEET Is Nothing Then: Set SRC_WSHEET = ActiveSheet
    
    For Each TEMP_CELL In SRC_RNG
        On Error Resume Next
            SRC_WSHEET.Shapes("Image:" & TEMP_CELL.Address).Delete
        On Error GoTo 1983

        If Left(TEMP_CELL.value, 7) = "Image: " Then
           With SRC_WSHEET.Pictures.Insert(Mid(TEMP_CELL.value, 8, 999))
                .Left = TEMP_CELL.Left + 1
                .Top = TEMP_CELL.Top + 1
                .name = "Image:" & TEMP_CELL.Address
                TEMP_CELL.RowHeight = .Height + 2
                TEMP_VAL = TEMP_CELL.Width / TEMP_CELL.ColumnWidth
                TEMP_CELL.ColumnWidth = .Width / TEMP_VAL + 2
           End With
        End If
1983:
    Next TEMP_CELL

    RNG_INSERT_PICTURE_FUNC = True

Exit Function
ERROR_LABEL:
RNG_INSERT_PICTURE_FUNC = False
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : RNG_LAST_ROW_POS
'DESCRIPTION   : Last Row number
'LIBRARY       : EXCEL
'GROUP         : POSITION
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA

'************************************************************************************
'************************************************************************************

Function RNG_LAST_ROW_POS(ByVal AROW As Long, _
ByVal ACOLUMN As Long, _
Optional ByRef SRC_WSHEET As Excel.Worksheet)
  
  Dim i As Long
  Dim TEMP_VALUE As String
  
  On Error GoTo ERROR_LABEL
  
  If SRC_WSHEET Is Nothing Then: Set SRC_WSHEET = ActiveSheet
  
  i = AROW
  TEMP_VALUE = SRC_WSHEET.Cells(i, ACOLUMN).value
  While TEMP_VALUE <> ""
    TEMP_VALUE = SRC_WSHEET.Cells(i, ACOLUMN).value
    i = i + 1
  Wend
  
  RNG_LAST_ROW_POS = i - 2

Exit Function
ERROR_LABEL:
    RNG_LAST_ROW_POS = Err.number    'convergence not met
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : RNG_LAST_COL_POS
'DESCRIPTION   : Last Column number
'LIBRARY       : EXCEL
'GROUP         : POSITION
'ID            : 002
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA

'************************************************************************************
'************************************************************************************

Function RNG_LAST_COL_POS(ByVal AROW As Variant, _
ByVal ACOLUMN As Long, _
Optional ByRef SRC_WSHEET As Excel.Worksheet)
  
  Dim i As Long
  Dim TEMP_VALUE As Variant
  
  On Error GoTo ERROR_LABEL

  If SRC_WSHEET Is Nothing Then: Set SRC_WSHEET = ActiveSheet

  i = ACOLUMN
  TEMP_VALUE = SRC_WSHEET.Cells(AROW, i).value
  While TEMP_VALUE <> ""
    TEMP_VALUE = SRC_WSHEET.Cells(AROW, i).value
    i = i + 1
  Wend
  RNG_LAST_COL_POS = i - 2

Exit Function
ERROR_LABEL:
    RNG_LAST_COL_POS = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : RNG_PRECEDENTS_FUNC
'DESCRIPTION   : Returns a Range object that represents the range containing
'all the direct precedents of a cell. This can be a multiple selection (a
'union of Range objects) if there’s more than one precedent
'LIBRARY       : EXCEL
'GROUP         : PRECEDENTS
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA

'************************************************************************************
'************************************************************************************

Function RNG_PRECEDENTS_FUNC(ByRef SRC_RNG As Excel.Range, _
Optional ByVal VERSION As Integer = 0)

'If ActiveCell.HasFormula Then
'    FUNC_STR = Selection.Address
'    ADDRESS_STR = RNG_PRECEDENTS_FUNC(Range(FUNC_STR), 0)
'    MsgBox ADDRESS_STR
'End If

Dim FUNC_STR As String
Dim ADDRESS_STR As String

On Error GoTo ERROR_LABEL
    
If SRC_RNG.HasFormula Then
    FUNC_STR = SRC_RNG.Address
    Select Case VERSION
        Case 0
            If Range(FUNC_STR).Precedents.Address(False, False) <> "" Then
                ADDRESS_STR = Range(FUNC_STR).Precedents.Address
                RNG_PRECEDENTS_FUNC = ADDRESS_STR
            Else
                GoTo ERROR_LABEL
            End If
        Case Else
            Set RNG_PRECEDENTS_FUNC = Range(FUNC_STR).DirectPrecedents
    End Select
Else
    GoTo ERROR_LABEL
End If

Exit Function
ERROR_LABEL:
RNG_PRECEDENTS_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : RNG_REPLACE_ELEMENTS_FUNC
'DESCRIPTION   : Replace values in a range in which a specified element has been
'replaced with another value.
'LIBRARY       : EXCEL
'GROUP         : REPLACE
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA

'************************************************************************************
'************************************************************************************


Function RNG_REPLACE_ELEMENTS_FUNC(ByRef ELEM_RNG As Variant, _
ByRef SRC_RNG As Excel.Range)

Dim i As Long
Dim NROWS As Long

Dim ELEMENT_STR As Variant
Dim ELEM_ARR As Variant

Dim DATA_CELL As Excel.Range 'RANDOM CELL

On Error GoTo ERROR_LABEL

RNG_REPLACE_ELEMENTS_FUNC = False
ELEM_ARR = ELEM_RNG

If UBound(ELEM_ARR, 1) = 1 Then: _
    ELEM_ARR = MATRIX_TRANSPOSE_FUNC(ELEM_ARR)

NROWS = UBound(ELEM_ARR, 1)

For i = 1 To NROWS
    ELEMENT_STR = ELEM_ARR(i, 1)
    For Each DATA_CELL In SRC_RNG
        If Mid(DATA_CELL, 1, Len(ELEMENT_STR)) Like ELEMENT_STR Then
                 DATA_CELL = ELEMENT_STR 'RESET VALUES
             Exit For
        End If
    Next DATA_CELL
Next i

RNG_REPLACE_ELEMENTS_FUNC = True

Exit Function
ERROR_LABEL:
RNG_REPLACE_ELEMENTS_FUNC = False
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : RNG_CLEARCONTENTS_FUNC
'DESCRIPTION   : Reset Results
'LIBRARY       : EXCEL
'GROUP         : RESET
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA

'************************************************************************************
'************************************************************************************

Function RNG_CLEARCONTENTS_FUNC(ByVal SRC_RNG_STR As String, _
Optional ByVal VERSION As Long = 0)

Dim i As Long
Dim SROW As Long
Dim SCOLUMN As Long

Dim NROWS As Long
Dim SRC_WSHEET As Excel.Worksheet

On Error GoTo ERROR_LABEL

RNG_CLEARCONTENTS_FUNC = False
If SRC_RNG_STR = "" Then GoTo ERROR_LABEL
Set SRC_WSHEET = Range(SRC_RNG_STR).Worksheet
SROW = Range(SRC_RNG_STR).row
SCOLUMN = Range(SRC_RNG_STR).Column

If VERSION = 0 Then
    i = 0
ElseIf VERSION = 1 Then
    NROWS = SRC_WSHEET.Cells(SROW, SCOLUMN).CurrentRegion.Rows.COUNT + SROW
    Range(SRC_WSHEET.Cells(SROW, SCOLUMN), _
          SRC_WSHEET.Cells(NROWS, SCOLUMN + 1)).ClearContents
    i = 0
End If

RNG_CLEARCONTENTS_FUNC = True

Exit Function
ERROR_LABEL:
RNG_CLEARCONTENTS_FUNC = False
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : RNG_NAME_ADDRESS_FUNC
'DESCRIPTION   : Returns the range reference in the language of the macro
'LIBRARY       : EXCEL
'GROUP         : SOURCE
'ID            : 002
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA

'************************************************************************************
'************************************************************************************

Function RNG_NAME_ADDRESS_FUNC(ByRef DATA_CELL As Excel.Range, _
Optional ByVal VERSION As Long = 0)

On Error GoTo ERROR_LABEL

Select Case VERSION
    Case 0 'e.g., =$E$5
        RNG_NAME_ADDRESS_FUNC = DATA_CELL.Address(External:=False)
    Case 1 'e.g., =VISUALIZATION!$E$5
        RNG_NAME_ADDRESS_FUNC = DATA_CELL.name
    Case 2 'e.g., = [ORG_BOOK.xls]VISUALIZATION!$E$5
        RNG_NAME_ADDRESS_FUNC = DATA_CELL.Address(External:=True)
    Case Else 'e.g., = NICO --> which is the name of the cell $E$5
        RNG_NAME_ADDRESS_FUNC = DATA_CELL.name.name
End Select
Exit Function
ERROR_LABEL:
    RNG_NAME_ADDRESS_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : RNG_INDIRECT_FUNC
'DESCRIPTION   :
'LIBRARY       : EXCEL
'GROUP         : SOURCE
'ID            : 003
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'************************************************************************************
'************************************************************************************

Function RNG_INDIRECT_FUNC(ByVal SRC_RNG_STR As String, _
Optional ByVal VERSION As Integer = 0, _
Optional ByVal INDEX_NO As Long = 1, _
Optional ByVal VOLAT_FLAG As Boolean = False)

'EXTREMELY USEFUL WITH GOAL_SEEK
    
If VOLAT_FLAG = True Then: Excel.Application.Volatile (True)

Select Case VERSION
Case 0
    Set RNG_INDIRECT_FUNC = Range(SRC_RNG_STR)
Case 1
    Set RNG_INDIRECT_FUNC = Range(SRC_RNG_STR).Columns(INDEX_NO)
Case 2
    Set RNG_INDIRECT_FUNC = Range(SRC_RNG_STR).Rows(INDEX_NO)
Case Else
    RNG_INDIRECT_FUNC = Range(SRC_RNG_STR).value
End Select

Exit Function
ERROR_LABEL:
RNG_INDIRECT_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : RNG_ADDRESS_STRING_FUNC
'DESCRIPTION   : RETURNS THE ADDRESS REFERENCE SPECIFIED BY A RANGE
'LIBRARY       : EXCEL
'GROUP         : SOURCE
'ID            : 004
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA

'************************************************************************************
'************************************************************************************

Function RNG_ADDRESS_STRING_FUNC(ByRef SRC_RNG As Excel.Range)

Dim SROW As Long
Dim NROWS As Long

Dim SCOLUMN As Long
Dim NCOLUMNS As Long

Dim TEMP_STR As String

On Error GoTo ERROR_LABEL

SROW = SRC_RNG.row
SCOLUMN = SRC_RNG.Column

NROWS = SRC_RNG.row + SRC_RNG.Rows.COUNT - 1
NCOLUMNS = SRC_RNG.Column + SRC_RNG.Columns.COUNT - 1

If SROW = NROWS And SCOLUMN < NCOLUMNS Then
    TEMP_STR = Range(SRC_RNG.Worksheet.Cells(SROW + 1, SCOLUMN), _
                SRC_RNG.Worksheet.Cells(NROWS + 2, NCOLUMNS)).Address
Else
    TEMP_STR = Range(SRC_RNG.Worksheet.Cells(SROW, SCOLUMN + 1), _
                SRC_RNG.Worksheet.Cells(NROWS, SCOLUMN + 2)).Address
End If

RNG_ADDRESS_STRING_FUNC = TEMP_STR

Exit Function
ERROR_LABEL:
RNG_ADDRESS_STRING_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : RNG_BOX_ADDRESS_FUNC
'DESCRIPTION   : RETURNS THE ADDRESS REFERENCE SPECIFIED BY ROW & COLUMN
'LIBRARY       : EXCEL
'GROUP         : SOURCE
'ID            : 005
'LAST UPDATE   : 15/02/2008
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'************************************************************************************
'************************************************************************************

Function RNG_BOX_ADDRESS_FUNC(ByVal SROW As Long, _
ByVal SCOLUMN As Long, _
ByVal NROWS As Long, _
ByVal NCOLUMNS As Long, _
Optional ByRef SRC_WSHEET As Excel.Worksheet)

On Error GoTo ERROR_LABEL

If SRC_WSHEET Is Nothing Then: Set SRC_WSHEET = ActiveSheet

If SRC_WSHEET.Cells(SROW, SCOLUMN) <> "" And _
    IsNumeric(SRC_WSHEET.Cells(SROW, SCOLUMN)) Then
    NROWS = SRC_WSHEET.Cells(SROW, SCOLUMN).CurrentRegion.Rows.COUNT + _
             SRC_WSHEET.Cells(SROW, SCOLUMN).CurrentRegion.row - 1 'Get Last Row
    NCOLUMNS = SCOLUMN + 1
    RNG_BOX_ADDRESS_FUNC = Range(SRC_WSHEET.Cells(SROW, SCOLUMN), _
                            SRC_WSHEET.Cells(NROWS, NCOLUMNS)).Address
Else
    GoTo ERROR_LABEL
End If

Exit Function
ERROR_LABEL:
RNG_BOX_ADDRESS_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : RNG_FIND_POSITION_FUNC
'DESCRIPTION   : Finds specific information in a range, and returns an array
'with the information found
'LIBRARY       : EXCEL
'GROUP         : RNG-LOOK
'ID            : 001
'LAST UPDATE   : 15/02/2008
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'************************************************************************************
'************************************************************************************

Function RNG_FIND_POSITION_FUNC(ByVal KEY_WORD As String, _
ByRef SRC_RNG As Excel.Range, _
Optional ByVal VERSION As Long = 1, _
Optional ByVal A_SRC_ROW As Long = 0, _
Optional ByVal A_SRC_COLUMN As Long = 0, _
Optional ByVal B_SRC_ROW As Long = 0, _
Optional ByVal B_SRC_COLUMN As Long = 0)

Dim i As Long
Dim j As Long

Dim NROWS As Long
Dim NCOLUMNS As Long

Dim TEMP_RNG As Excel.Range
Dim TEMP_MATRIX As Variant

On Error GoTo ERROR_LABEL

j = 10000 'Loop Limit
i = 0

Set TEMP_RNG = SRC_RNG.Find(What:=KEY_WORD, AFTER:=SRC_RNG.Cells(1, 1), _
                LookIn:=xlValues, LookAt:=xlPart, _
                SearchOrder:=xlByRows, SearchDirection:=xlNext, _
                MatchCase:=False, SearchFormat:=False)

Do Until Left(TEMP_RNG.value, Len(KEY_WORD)) Like KEY_WORD
    i = i + 1
    Set TEMP_RNG = SRC_RNG.Find(What:=KEY_WORD, AFTER:=TEMP_RNG, _
                LookIn:=xlValues, LookAt:=xlPart, _
                SearchOrder:=xlByRows, SearchDirection:=xlNext, _
                MatchCase:=False, SearchFormat:=False)
    If i = j Then Exit Do
Loop

'----------------------------------------------------------------------------
If Not TEMP_RNG Is Nothing Then
'----------------------------------------------------------------------------
    
    Set TEMP_RNG = TEMP_RNG.Offset(A_SRC_ROW, A_SRC_COLUMN)
    
    Select Case VERSION
    Case 0
        Set TEMP_RNG = _
            TEMP_RNG.Offset(B_SRC_ROW, B_SRC_COLUMN)
    Case 1
        With TEMP_RNG
            Set TEMP_RNG = _
                Range(.Offset(B_SRC_ROW, B_SRC_COLUMN), .End(xlDown))
        End With
    Case 2
        With TEMP_RNG
            Set TEMP_RNG = _
            Range(.Offset(B_SRC_ROW, B_SRC_COLUMN).End(xlToRight), _
            .Offset(0, 0))
        End With
    Case 3
        With TEMP_RNG
            Set TEMP_RNG = _
            Range(.Offset(B_SRC_ROW, B_SRC_COLUMN).End(xlToRight), _
            .End(xlDown))
        End With
    Case 4
        With TEMP_RNG
            Set TEMP_RNG = _
            Range(.Offset(B_SRC_ROW, B_SRC_COLUMN), .End(xlToLeft))
        End With
    Case 5
        With TEMP_RNG
            Set TEMP_RNG = _
            Range(.Offset(B_SRC_ROW, B_SRC_COLUMN).End(xlToLeft), _
            .Offset(0, 0))
        End With
    Case 6
        With TEMP_RNG
            Set TEMP_RNG = _
            Range(.Offset(B_SRC_ROW, B_SRC_COLUMN).End(xlToLeft), _
            .End(xlDown))
        End With
    Case 7
        With TEMP_RNG
            Set TEMP_RNG = _
            Range(.Offset(0, 0), .Offset(B_SRC_ROW, B_SRC_COLUMN))
        End With
    End Select
    
    NROWS = TEMP_RNG.Rows.COUNT
    NCOLUMNS = TEMP_RNG.Columns.COUNT
        
    ReDim TEMP_MATRIX(1 To NROWS, 1 To NCOLUMNS)
    For i = 1 To NROWS
        For j = 1 To NCOLUMNS
            TEMP_MATRIX(i, j) = TEMP_RNG.Cells(i, j)
        Next j
    Next i
    
    RNG_FIND_POSITION_FUNC = TEMP_MATRIX
'----------------------------------------------------------------------------
End If
'----------------------------------------------------------------------------

Exit Function
ERROR_LABEL:
RNG_FIND_POSITION_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : RNG_TRANSPOSE_2D_ARRAY_FUNC

'DESCRIPTION   : This function tranposes the array DATA_RNG. DATA_RNG must be
'a two dimensional array. If DATA_RNG is not an array, the
'result is just DATA_RNG itself. If DATA_RNG is a 1-dimensional
'array, the result is just DATA_RNG itself. If you need to
'transpose a 1-dimensional array from a row to a column
'in order to properly return it to a worksheet, use
'RNG_TRANSPOSE_1D_ARRAY_FUNC. If DATA_RNG has more than three dimensions,
'an error value is returned.

'LIBRARY       : RANGE
'GROUP         : TRANSPOSE
'ID            : 003
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA

'************************************************************************************
'************************************************************************************

Function RNG_TRANSPOSE_2D_ARRAY_FUNC(ByRef DATA_RNG As Variant)

Dim i As Long
Dim j As Long

Dim ii As Long
Dim jj As Long

Dim SROW As Long
Dim SCOLUMN As Long

Dim NROWS As Long
Dim NCOLUMNS As Long

Dim NSIZE As Long

Dim TEMP_MATRIX() As Variant

On Error GoTo ERROR_LABEL

If IsArray(DATA_RNG) = False Then
    RNG_TRANSPOSE_2D_ARRAY_FUNC = DATA_RNG
    Exit Function
End If

NSIZE = ARRAY_DIMENSION_FUNC(DATA_RNG)
Select Case NSIZE
    Case 0
        If IsObject(DATA_RNG) = True Then
            Set RNG_TRANSPOSE_2D_ARRAY_FUNC = DATA_RNG
        Else
            RNG_TRANSPOSE_2D_ARRAY_FUNC = DATA_RNG
        End If
    Case 1
        RNG_TRANSPOSE_2D_ARRAY_FUNC = DATA_RNG
    Case 2
        
        SROW = LBound(DATA_RNG, 1)
        NROWS = UBound(DATA_RNG, 1)
        SCOLUMN = LBound(DATA_RNG, 2)
        NCOLUMNS = UBound(DATA_RNG, 2)
        
        ii = SROW
        jj = SCOLUMN
        
        ReDim TEMP_MATRIX(SCOLUMN To NCOLUMNS, SROW To NROWS)
        
        For i = SROW To NROWS
            For j = SCOLUMN To NCOLUMNS
                TEMP_MATRIX(j, i) = DATA_RNG(i, j)
                jj = jj + 1
            Next j
            ii = ii + 1
        Next i
        RNG_TRANSPOSE_2D_ARRAY_FUNC = TEMP_MATRIX
    End Select

Exit Function
ERROR_LABEL:
RNG_TRANSPOSE_2D_ARRAY_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : RNG_TRANSPOSE_1D_ARRAY_FUNC

'DESCRIPTION   : This function transforms a 1-dim array to a 2-dim array and
' transposes it. This is required when returning arrays back to
' worksheet cells. The ROW_FLAG parameter determines if the array is
' to be returned to the worksheet as a row (TRUE) or as a columns (FALSE).

'LIBRARY       : MATRIX
'GROUP         : TRANSPOSE
'ID            : 004



'AUTHOR        : RAFAEL NICOLAS FERMIN COTA

'************************************************************************************
'************************************************************************************

Function RNG_TRANSPOSE_1D_ARRAY_FUNC(ByRef DATA_RNG As Variant, _
Optional ByVal ROW_FLAG As Boolean = True)

Dim ii As Long
Dim SROW As Long
Dim NROWS As Long
Dim TEMP_MATRIX As Variant

On Error GoTo ERROR_LABEL

If IsArray(DATA_RNG) = False Then: GoTo ERROR_LABEL
If ARRAY_DIMENSION_FUNC(DATA_RNG) <> 1 Then: GoTo ERROR_LABEL

If ROW_FLAG = True Then
    SROW = LBound(DATA_RNG)
    NROWS = UBound(DATA_RNG)
    ReDim TEMP_MATRIX(SROW To SROW, SROW To NROWS)
    For ii = SROW To NROWS
        TEMP_MATRIX(SROW, ii) = DATA_RNG(ii)
    Next ii
Else
    SROW = LBound(DATA_RNG)
    NROWS = UBound(DATA_RNG)
    ReDim TEMP_MATRIX(SROW To NROWS, SROW To SROW)
    For ii = SROW To NROWS
        TEMP_MATRIX(ii, SROW) = DATA_RNG(ii)
    Next ii
End If

RNG_TRANSPOSE_1D_ARRAY_FUNC = TEMP_MATRIX

Exit Function
ERROR_LABEL:
RNG_TRANSPOSE_1D_ARRAY_FUNC = Err.number
End Function

'TrimRange function removes all the blank (="") rows of a
'range and returns a shortened range
'(without changing the sequence of elements)
'A blank in the first column will determine whether a row is removed.
'This function is used to clean the input to the LINEST worksheet function

Function RNG_TRIM_FUNC(ByRef SRC_RNG As Excel.Range)

Dim i As Long
Dim j As Long
Dim k As Long

Dim NROWS As Long
Dim NCOLUMNS As Long

Dim TEMP_MATRIX() As Variant

On Error GoTo ERROR_LABEL

NROWS = SRC_RNG.Rows.COUNT
NCOLUMNS = SRC_RNG.Columns.COUNT
j = 0
For i = 1 To SRC_RNG.Rows.COUNT
  If SRC_RNG(i, 1) <> "" Then  'count empty rows (first column is looked at only)
    j = j + 1
  End If
Next i

If j > 0 Then 'non-empty rows were found
  ReDim TEMP_MATRIX(1 To j, 1 To NCOLUMNS)
  j = 0
  For i = 1 To NROWS
    If SRC_RNG(i, 1) <> "" Then
      j = j + 1
      For k = 1 To NCOLUMNS
       TEMP_MATRIX(j, k) = SRC_RNG(i, k)
      Next k
    End If
  Next i
End If

RNG_TRIM_FUNC = TEMP_MATRIX

Exit Function
ERROR_LABEL:
RNG_TRIM_FUNC = Err.number
End Function


