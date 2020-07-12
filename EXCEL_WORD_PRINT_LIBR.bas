Attribute VB_Name = "EXCEL_WORD_PRINT_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.

'************************************************************************************
'************************************************************************************
'FUNCTION      : WORD_PRINT_FORMULAS_FUNC
'DESCRIPTION   : This routine will print all of the entries in a matrix
'LIBRARY       : EXCEL
'GROUP         : MICROSOFT WORD
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 03/02/2008
'************************************************************************************
'************************************************************************************

Function WORD_PRINT_DATA_FUNC(ByRef DATA_RNG As Variant, _
Optional ByVal REPORT_NAME_STR As String = "TEST_REPORT")

Dim i As Long
Dim j As Long

Dim SROW As Long
Dim NROWS As Long

Dim SCOLUMN As Long
Dim NCOLUMNS As Long

Dim DATA_MATRIX As Variant

Dim WORD_OBJ As Object

On Error Resume Next

WORD_PRINT_DATA_FUNC = False
Err.number = 0

Set WORD_OBJ = GetObject(, "Word.Application.8")
If Err.number = 429 Then
    Set WORD_OBJ = CreateObject("Word.Application.8")
    Err.number = 0
End If

DATA_MATRIX = DATA_RNG

SROW = LBound(DATA_MATRIX, 1)
NROWS = UBound(DATA_MATRIX, 1)

SCOLUMN = LBound(DATA_MATRIX, 2)
NCOLUMNS = UBound(DATA_MATRIX, 2)

WORD_OBJ.Visible = True
WORD_OBJ.Documents.Add

With WORD_OBJ.Selection
    .Font.name = "Courier New"
    .TypeText "Report: " & REPORT_NAME_STR
    .TypeParagraph
    .TypeText Text:="Date: " + Format(Now(), "dd-mmm-yy hh:mm")
    .TypeParagraph 'New Line
    .TypeParagraph 'New Line
End With

With WORD_OBJ.Selection
    For j = SCOLUMN To NCOLUMNS 'j = SCOLUMN + 1 To NCOLUMNS
        For i = SROW To NROWS
            '.Font.Bold = True
            '.TypeText DATA_MATRIX(i, SCOLUMN) & ": "
            .Font.Bold = False
            .TypeText DATA_MATRIX(i, j)
            .TypeParagraph 'New Line
        Next i
        .TypeParagraph 'New Line
    Next j
End With

Set WORD_OBJ = Nothing

If Err.number = 0 Then
    WORD_PRINT_DATA_FUNC = True 'Done printing formulas to Word
Else
    WORD_PRINT_DATA_FUNC = False
End If
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : WORD_IMPORT_TEXT_FUNC
'DESCRIPTION   : Get data from Microsoft Word
'LIBRARY       : EXCEL
'GROUP         : MICROSOFT WORD
'ID            : 002
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 03/02/2008
'************************************************************************************
'************************************************************************************


Function WORD_IMPORT_TEXT_FUNC(ByVal FULL_PATH_NAME As String, _
ByVal START_DOC As Variant, _
ByVal END_DOC As Variant)

Dim WORD_OBJ As Object
Dim WORD_RNG As Object
Dim DATA_STR As String

On Error Resume Next
WORD_IMPORT_TEXT_FUNC = False
Err.number = 0

Set WORD_OBJ = GetObject(, "Word.Application.8")
If Err.number = 429 Then
    Set WORD_OBJ = CreateObject("Word.Application.8")
    Err.number = 0
End If

WORD_OBJ.Visible = False

WORD_OBJ.Documents.Open FileName:=FULL_PATH_NAME
'Set WORD_RNG = WORD_OBJ.ActiveDocument.Range(START_DOC, END_DOC) 'By Line
Set WORD_RNG = WORD_OBJ.ActiveDocument.Range( _
    WORD_OBJ.ActiveDocument.Paragraphs(START_DOC).Range.Start, _
    WORD_OBJ.ActiveDocument.Paragraphs(END_DOC).Range.End) 'By Paragrah

DATA_STR = WORD_RNG.Text
WORD_OBJ.ActiveDocument.Close (False)
WORD_OBJ.Application.Quit

Set WORD_OBJ = Nothing

If Err.number = 0 Then
    WORD_IMPORT_TEXT_FUNC = DATA_STR
Else
    WORD_IMPORT_TEXT_FUNC = Err.number
End If
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : WORD_PRINT_COMMENTS_FUNC
'DESCRIPTION   : This routine will print all of the cell comments in the
' active workbook to a Word file.  Word will be left running.
'LIBRARY       : EXCEL
'GROUP         : MICROSOFT WORD
'ID            : 003
'LAST UPDATE   : 03/02/2008
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'************************************************************************************
'************************************************************************************


Function WORD_PRINT_COMMENTS_FUNC(ByRef SRC_WBOOK As Excel.Workbook, _
Optional ByVal VERSION As Integer = 0)

Dim i As Long
Dim WORD_OBJ As Object
Dim TEMP_FLAG As Boolean
Dim DATA_STR As String
Dim TEMP_CELL As Excel.Range
Dim SRC_WSHEET As Excel.Worksheet

On Error Resume Next
Err.number = 0

If SRC_WBOOK Is Nothing Then: Set SRC_WBOOK = ActiveWorkbook
WORD_PRINT_COMMENTS_FUNC = False

Select Case VERSION 'Do want to print cell values with comments
    Case 0 'vbYes
        TEMP_FLAG = True
    Case Else 'vbNo
        TEMP_FLAG = False
End Select

Set WORD_OBJ = GetObject(, "Word.Application.8")
If Err.number = 429 Then
    Set WORD_OBJ = CreateObject("Word.Application.8")
    Err.number = 0
End If

WORD_OBJ.Visible = True
WORD_OBJ.Documents.Add
With WORD_OBJ.Selection
    .TypeText Text:="Cell Comments In Workbook: " + ActiveWorkbook.name
    .TypeParagraph
    .TypeText Text:="Date: " + Format(Now(), "dd-mmm-yy hh:mm")
    .TypeParagraph
    .TypeParagraph
End With
    
For Each SRC_WSHEET In SRC_WBOOK.Worksheets
    For i = 1 To SRC_WSHEET.comments.COUNT
       Set TEMP_CELL = SRC_WSHEET.comments(i).Parent
       DATA_STR = SRC_WSHEET.comments(i).Text
       With WORD_OBJ.Selection
            .TypeText Text:="Comment In Cell: " + _
             TEMP_CELL.Address(False, False, xlA1) + " on sheet: " + _
                                SRC_WSHEET.name
             If TEMP_FLAG = True Then
                .TypeText Text:="  Cell Value: " + Format(TEMP_CELL.value)
             End If
            .TypeParagraph
            .TypeText Text:=DATA_STR
            .TypeParagraph
            .TypeParagraph
        End With
    Next i
Next SRC_WSHEET

Set WORD_OBJ = Nothing
If Err.number = 0 Then
    WORD_PRINT_COMMENTS_FUNC = True 'Finished Printing Comments To Word
Else
    WORD_PRINT_COMMENTS_FUNC = False
End If
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : WORD_PRINT_FORMULAS_FUNC
'DESCRIPTION   : This routine will print all of the cell comments in the
' active workbook to a Word file.  Word will be left running.
'LIBRARY       : EXCEL
'GROUP         : MICROSOFT WORD
'ID            : 004
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 03/02/2008

'************************************************************************************
'************************************************************************************

Function WORD_PRINT_FORMULAS_FUNC(ByRef DATA_RNG As Excel.Range)

Dim DATA_STR As String
Dim TEMP_CELL As Excel.Range
Dim WORD_OBJ As Object
Dim TEMP_FLAG As Boolean

On Error Resume Next

WORD_PRINT_FORMULAS_FUNC = False
Err.number = 0

Set WORD_OBJ = GetObject(, "Word.Application.8")
If Err.number = 429 Then
    Set WORD_OBJ = CreateObject("Word.Application.8")
    Err.number = 0
End If

WORD_OBJ.Visible = True
WORD_OBJ.Documents.Add

With WORD_OBJ.Selection
    .Font.name = "Courier New"
    .TypeText "Formulas In Worksheet: " + ActiveSheet.name
    .TypeParagraph
    .TypeText Text:="Date: " + Format(Now(), "dd-mmm-yy hh:mm")
    .TypeParagraph
    .TypeText "Cells: " + DATA_RNG.Cells(1, 1).Address(False, False, xlA1) & _
                        " to " & DATA_RNG.Cells(DATA_RNG.Rows.COUNT, _
                        DATA_RNG.Columns.COUNT).Address(False, False, xlA1)
    .TypeParagraph
    .TypeParagraph
End With

For Each TEMP_CELL In DATA_RNG
    TEMP_FLAG = TEMP_CELL.HasArray
    DATA_STR = TEMP_CELL.formula
    If TEMP_FLAG Then
        DATA_STR = "{" + DATA_STR + "}"
    End If
    If DATA_STR <> "" Then
        With WORD_OBJ.Selection
            .Font.Bold = True
            .TypeText TEMP_CELL.Address(False, False, xlA1) & ": "
            .Font.Bold = False
            .TypeText DATA_STR
            .TypeParagraph
            .TypeParagraph
        End With
    End If
Next TEMP_CELL
Set WORD_OBJ = Nothing

If Err.number = 0 Then
    WORD_PRINT_FORMULAS_FUNC = True 'Done printing formulas to Word
Else
    WORD_PRINT_FORMULAS_FUNC = False
End If
End Function
