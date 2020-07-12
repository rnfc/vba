Attribute VB_Name = "EXCEL_WSHEET_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.

Function WSHEET_ADD_FUNC(ByVal SHEET_NAME As String, _
Optional ByRef SRC_WBOOK As Excel.Workbook) As Excel.Worksheet

Dim TEMP_FLAG As Boolean
On Error GoTo ERROR_LABEL

If SRC_WBOOK Is Nothing Then: Set SRC_WBOOK = ActiveWorkbook
SRC_WBOOK.Worksheets.Add , SRC_WBOOK.ActiveSheet
TEMP_FLAG = WSHEET_RENAME_FUNC(SHEET_NAME, SRC_WBOOK)
If TEMP_FLAG = True Then
    Set WSHEET_ADD_FUNC = SRC_WBOOK.Worksheets(SHEET_NAME)
Else
    Set WSHEET_ADD_FUNC = SRC_WBOOK.ActiveSheet
End If

Exit Function
ERROR_LABEL:
Set WSHEET_ADD_FUNC = Nothing
End Function

Sub WSHEETS_REMOVE_CURRENT_FUNC()

Dim kk As Integer
Dim nSHEET As Excel.Worksheet
Dim DST_WBOOK As Excel.Workbook
Dim WSHEET_NAME As String

On Error GoTo ERROR_LABEL

Set DST_WBOOK = ActiveWorkbook

Call EXCEL_TURN_OFF_EVENTS_FUNC
Call EXCEL_DISPLAY_ALERTS_FUNC(False)

WSHEET_NAME = CStr(Date)
WSHEET_NAME = Replace(WSHEET_NAME, "/", "_")

kk = Len(WSHEET_NAME)
For Each nSHEET In DST_WBOOK.Worksheets
    If Left(nSHEET.name, kk) Like WSHEET_NAME Then: nSHEET.Delete
Next nSHEET

Call EXCEL_TURN_ON_EVENTS_FUNC
Call EXCEL_DISPLAY_ALERTS_FUNC(True)

Exit Sub
ERROR_LABEL:
Call EXCEL_TURN_ON_EVENTS_FUNC
Call EXCEL_DISPLAY_ALERTS_FUNC(True)
'ADD MSG HERE; Err.Description
End Sub

Function WSHEETS_REMOVE_FUNC(ByRef DATA_ARR As Variant, _
Optional ByRef SRC_WBOOK As Excel.Workbook)

Dim i As Long
Dim TEMP_SHEET As Excel.Worksheet
Dim TEMP_FLAG As Boolean
        
On Error GoTo ERROR_LABEL
        
WSHEETS_REMOVE_FUNC = False
If SRC_WBOOK Is Nothing Then: Set SRC_WBOOK = ActiveWorkbook

For Each TEMP_SHEET In SRC_WBOOK.Worksheets
    TEMP_FLAG = False
    For i = LBound(DATA_ARR) To UBound(DATA_ARR)
         If DATA_ARR(i) = TEMP_SHEET.name Then
             TEMP_FLAG = True
             Exit For
         End If
    Next i
    If TEMP_FLAG = False Then: TEMP_SHEET.Delete
Next TEMP_SHEET

WSHEETS_REMOVE_FUNC = True

Exit Function
ERROR_LABEL:
WSHEETS_REMOVE_FUNC = False
End Function


Function WSHEET_REMOVE_FUNC(Optional ByVal NAME_STR As Variant = Null, _
Optional ByRef SRC_WBOOK As Excel.Workbook)

Dim TEMP_STR As String
Dim TEMP_SHEET As Excel.Worksheet

On Error GoTo ERROR_LABEL

WSHEET_REMOVE_FUNC = False
If SRC_WBOOK Is Nothing Then: Set SRC_WBOOK = ActiveWorkbook

If VarType(NAME_STR) = vbNull Then
    TEMP_STR = PARSE_CURRENT_TIME_FUNC("_")
    For Each TEMP_SHEET In SRC_WBOOK.Worksheets
        If Left(TEMP_SHEET.name, 5) Like Left(TEMP_STR, 5) _
        Then: TEMP_SHEET.Delete
    Next TEMP_SHEET
Else
    TEMP_STR = NAME_STR
    For Each TEMP_SHEET In SRC_WBOOK.Worksheets
        If TEMP_SHEET.name = TEMP_STR Then: TEMP_SHEET.Delete
    Next TEMP_SHEET
End If

WSHEET_REMOVE_FUNC = True

Exit Function
ERROR_LABEL:
WSHEET_REMOVE_FUNC = False
End Function

'Excel.Application.ScreenUpdating = False
'TEMP_GROUP = WSHEETS_COMPARE_FUNC(Range("a1"), Worksheets("FCFF_CALCS"), Worksheets("FCFF_MODEL"), 2, 500, True)

Function WSHEETS_COMPARE_FUNC(ByRef DST_RNG As Excel.Range, _
ByRef FIRST_WORKSHEET As Excel.Worksheet, _
ByRef SECOND_WORKSHEET As Excel.Worksheet, _
Optional ByVal VERSION As Long = 2, _
Optional ByVal NSIZE As Long = 500, _
Optional ByVal FORMAT_FLAG As Boolean = True)

'FORMAT_FLAG: Include cells that differ only in number format

'IF VERSION 0 = Formulas differ
'IF VERSION 1 = Values differ
'IF VERSION 2 = Either formulas or values differ

Dim i As Long
Dim j As Long

Dim hh As Long
Dim ii As Long
Dim jj As Long
Dim kk As Long

Dim AROW As Long 'Last Row
Dim NROWS As Long

Dim FIRST_VAL As String
Dim SECOND_VAL As String

Dim EQUAL_STR As String
Dim SEPAR_STR As String

Dim TEMP_GROUP As Variant
Dim LABEL_ARR As Variant

Dim FIRST_CELL As Excel.Range
Dim SECOND_CELL As Excel.Range

Dim FIRST_WBOOK As String
Dim SECOND_WBOOK As String

Dim FIRST_FLAG As Boolean
Dim SECOND_FLAG As Boolean

On Error GoTo ERROR_LABEL

FIRST_WBOOK = "[" & FIRST_WORKSHEET.Parent.name & "]" & FIRST_WORKSHEET.name
SECOND_WBOOK = "[" & SECOND_WORKSHEET.Parent.name & "]" & SECOND_WORKSHEET.name

NROWS = DST_RNG.Worksheet.Rows.COUNT
AROW = 1

ReDim LABEL_ARR(1 To 3)
LABEL_ARR(1) = "Formula"
LABEL_ARR(2) = "Value"
LABEL_ARR(3) = "Numberformat"

ReDim TEMP_GROUP(0 To NSIZE, 1 To 4) As Variant
TEMP_GROUP(0, 1) = "Address"
TEMP_GROUP(0, 2) = "Difference"
TEMP_GROUP(0, 3) = FIRST_WBOOK
TEMP_GROUP(0, 4) = SECOND_WBOOK

kk = 0

ii = Excel.Application.max( _
FIRST_WORKSHEET.Range("A1").SpecialCells(xlLastCell).row, _
SECOND_WORKSHEET.Range("A1").SpecialCells(xlLastCell).row)

jj = Excel.Application.max( _
FIRST_WORKSHEET.Range("A1").SpecialCells(xlLastCell).Column, _
SECOND_WORKSHEET.Range("A1").SpecialCells(xlLastCell).Column)

For i = 1 To ii
For j = 1 To jj

Set FIRST_CELL = FIRST_WORKSHEET.Cells(i, j)
Set SECOND_CELL = SECOND_WORKSHEET.Cells(i, j)

hh = 0

Select Case VERSION
'------------------------------------------------------------------------------
Case 0 'Compare formulas
'------------------------------------------------------------------------------
FIRST_VAL = FIRST_CELL.formula
SECOND_VAL = SECOND_CELL.formula

If FIRST_VAL <> SECOND_VAL Then
FIRST_FLAG = FIRST_CELL.HasFormula
SECOND_FLAG = SECOND_CELL.HasFormula

'1 indicates a formula difference, 2 a value difference
hh = (FIRST_FLAG Or SECOND_FLAG) + 2

If FIRST_FLAG = False Then FIRST_VAL = FIRST_CELL.value
If SECOND_FLAG = False Then SECOND_VAL = SECOND_CELL.value
End If
'------------------------------------------------------------------------------
Case 1 'Compare Values
'------------------------------------------------------------------------------
FIRST_VAL = FIRST_CELL.value
SECOND_VAL = SECOND_CELL.value
If TypeName(FIRST_VAL) <> TypeName(SECOND_VAL) Then
hh = 2
ElseIf FIRST_VAL <> SECOND_VAL Then
hh = 2
End If
'------------------------------------------------------------------------------
Case 2 'Compare Both
'------------------------------------------------------------------------------
FIRST_VAL = FIRST_CELL.formula
SECOND_VAL = SECOND_CELL.formula

If FIRST_VAL <> SECOND_VAL Then
FIRST_FLAG = FIRST_CELL.HasFormula
SECOND_FLAG = SECOND_CELL.HasFormula

'1 indicates a formula difference, 2 a value difference
hh = (FIRST_FLAG Or SECOND_FLAG) + 2

If FIRST_FLAG = False Then FIRST_VAL = FIRST_CELL.value
If SECOND_FLAG = False Then SECOND_VAL = SECOND_CELL.value
End If

If hh = 0 Then
FIRST_VAL = FIRST_CELL.value
SECOND_VAL = SECOND_CELL.value
If TypeName(FIRST_VAL) <> TypeName(SECOND_VAL) Then
hh = 2
ElseIf FIRST_VAL <> SECOND_VAL Then
hh = 2
End If
End If
'------------------------------------------------------------------------------
End Select

If hh = 0 And FORMAT_FLAG = True Then
If FIRST_CELL.NumberFormat <> SECOND_CELL.NumberFormat Then
hh = 3
FIRST_VAL = " " & FIRST_CELL.NumberFormat
SECOND_VAL = " " & SECOND_CELL.NumberFormat
End If
End If

If hh Then

EQUAL_STR = "="
SEPAR_STR = " "

If kk = NSIZE Then
If kk = 0 Then GoTo 1982  'nothing to write
'will all entries fit? if not, write as many as possible
If (NROWS - AROW) < kk Then kk = (NROWS - AROW)
DST_RNG.Cells(AROW + 1, 1).Resize(kk, 4).value = TEMP_GROUP
AROW = AROW + kk
kk = 0
1982:
End If

If Not IsError(FIRST_VAL) Then
If Left$(FIRST_VAL, 1) = EQUAL_STR Then
FIRST_VAL = SEPAR_STR & FIRST_VAL
End If
End If

If Not IsError(SECOND_VAL) Then
If Left$(SECOND_VAL, 1) = EQUAL_STR Then
SECOND_VAL = SEPAR_STR & SECOND_VAL
End If
End If

kk = kk + 1
TEMP_GROUP(kk, 1) = FIRST_CELL.Address
TEMP_GROUP(kk, 2) = LABEL_ARR(hh)
TEMP_GROUP(kk, 3) = FIRST_VAL
TEMP_GROUP(kk, 4) = SECOND_VAL

End If
If AROW >= NROWS Then 'Too many differences
WSHEETS_COMPARE_FUNC = False
Exit Function
End If
Next j
Next i


'write anything left in buffer to worksheet
If kk = 0 Then GoTo 1983  'nothing to write
'will all entries fit? if not, write as many as possible
If (NROWS - AROW) < kk Then kk = (NROWS - AROW)
DST_RNG.Cells(AROW + 1, 1).Resize(kk + 1, 4).value = TEMP_GROUP
AROW = AROW + kk
kk = 0
1983:

With DST_RNG.Worksheet.UsedRange.Columns
  .AutoFit
  .HorizontalAlignment = xlLeft
End With

WSHEETS_COMPARE_FUNC = True

Exit Function
ERROR_LABEL: 'No differences found!
WSHEETS_COMPARE_FUNC = False
End Function

Function WSHEETS_COUNT_FUNC(Optional ByRef SRC_WBOOK As Excel.Workbook)

Dim ii As Long
Dim EACH_SHEET As Excel.Worksheet

On Error GoTo ERROR_LABEL

If SRC_WBOOK Is Nothing Then: Set SRC_WBOOK = ActiveWorkbook

ii = 0
For Each EACH_SHEET In SRC_WBOOK.Worksheets
    ii = ii + 1
Next EACH_SHEET

WSHEETS_COUNT_FUNC = ii

Exit Function
ERROR_LABEL:
WSHEETS_COUNT_FUNC = Err.number
End Function

Function WSHEETS_HIDE_FUNC(ByRef DATA_ARR As Variant, _
Optional ByVal HIDE_OPT As Boolean = False, _
Optional ByRef SRC_WBOOK As Excel.Workbook)

Dim i As Long
Dim HIDE_FLAG As Boolean
Dim EACH_SHEET As Excel.Worksheet
        
On Error GoTo ERROR_LABEL
        
WSHEETS_HIDE_FUNC = False
If SRC_WBOOK Is Nothing Then: Set SRC_WBOOK = ActiveWorkbook

For Each EACH_SHEET In SRC_WBOOK.Worksheets
      HIDE_FLAG = False
      For i = LBound(DATA_ARR) To UBound(DATA_ARR)
           If DATA_ARR(i) = EACH_SHEET.name Then
               HIDE_FLAG = True
               Exit For
           End If
      Next i
      If HIDE_FLAG = False Then: EACH_SHEET.Visible = HIDE_OPT
Next EACH_SHEET

WSHEETS_HIDE_FUNC = True

Exit Function
ERROR_LABEL:
WSHEETS_HIDE_FUNC = False
End Function


Function WSHEET_LOOK_FUNC(ByVal SHEET_NAME As String, _
Optional ByVal VERSION As Integer = 0, _
Optional ByRef SRC_WBOOK As Excel.Workbook)

Dim ii As Long
Dim TEMP_FLAG As Boolean
Dim EACH_SHEET As Excel.Worksheet

On Error GoTo ERROR_LABEL

If SRC_WBOOK Is Nothing Then: Set SRC_WBOOK = ActiveWorkbook

TEMP_FLAG = False
ii = 0
For Each EACH_SHEET In SRC_WBOOK.Worksheets
    ii = ii + 1
    If EACH_SHEET.name = SHEET_NAME Then
        TEMP_FLAG = True
        Exit For
    End If
Next EACH_SHEET

If TEMP_FLAG = True Then
    Select Case VERSION
        Case 0
    WSHEET_LOOK_FUNC = True
        Case Else
    WSHEET_LOOK_FUNC = ii
    End Select
Else
    WSHEET_LOOK_FUNC = False
End If

Exit Function
ERROR_LABEL:
    WSHEET_LOOK_FUNC = False
End Function

Function WSHEET_PARSE_NAME_FUNC(ByVal WORK_NAME_STR As String, _
Optional ByVal REF_CHR As String = "]", _
Optional ByVal OUTPUT As Integer = 0)
  
  Dim ii As String
  
  On Error GoTo ERROR_LABEL
  
  ii = InStr(1, WORK_NAME_STR, REF_CHR, 0)
  If ii = 0 Then
    WSHEET_PARSE_NAME_FUNC = WORK_NAME_STR
    Exit Function
  End If
  
  Select Case OUTPUT
    Case 0
          WSHEET_PARSE_NAME_FUNC = Mid$(WORK_NAME_STR, ii + 1) 'Worksheet's Name
    Case Else
          WSHEET_PARSE_NAME_FUNC = Mid$(WORK_NAME_STR, 2, ii - 2) 'Workbook's Name
  End Select
  
Exit Function
ERROR_LABEL:
    WSHEET_PARSE_NAME_FUNC = Err.number
End Function

Function WSHEET_PRINT_FUNC()

Dim TEMP_RESP As Variant

On Error GoTo ERROR_LABEL

WSHEET_PRINT_FUNC = False
    TEMP_RESP = MsgBox("Would you like to Print the ActiveSheet." _
    & " Press OK to proceed.", vbOKCancel + vbExclamation)
    
    If TEMP_RESP = vbOK Then
        ActiveWindow.SelectedSheets.PrintOut Copies:=1
        ActiveChart.PrintOut Copies:=1
    End If
WSHEET_PRINT_FUNC = True

Exit Function
ERROR_LABEL:
WSHEET_PRINT_FUNC = False
End Function


Function WSHEET_PROTECT_FUNC(ByVal SHEET_NAME As String, _
ByVal PASSWORD_STR As String, _
Optional ByRef SRC_WBOOK As Excel.Workbook)

Dim SRC_WSHEET As Excel.Worksheet

On Error GoTo ERROR_LABEL
    
WSHEET_PROTECT_FUNC = False
If SRC_WBOOK Is Nothing Then: Set SRC_WBOOK = ActiveWorkbook

Set SRC_WSHEET = SRC_WBOOK.Worksheets(SHEET_NAME)
SRC_WSHEET.Protect _
    password:=PASSWORD_STR, _
    DrawingObjects:=True, _
    Contents:=True, _
    Scenarios:=True, _
    userinterfaceonly:=True
WSHEET_PROTECT_FUNC = True

Exit Function
ERROR_LABEL:
WSHEET_PROTECT_FUNC = False
End Function

Function WSHEET_UNPROTECT_FUNC(ByVal SHEET_NAME As String, _
ByVal PASSWORD_STR As String, _
Optional ByRef SRC_WBOOK As Excel.Workbook)

Dim SRC_WSHEET As Excel.Worksheet

On Error GoTo ERROR_LABEL
    
WSHEET_UNPROTECT_FUNC = False
If SRC_WBOOK Is Nothing Then: Set SRC_WBOOK = ActiveWorkbook

Set SRC_WSHEET = SRC_WBOOK.Worksheets(SHEET_NAME)
'remove all protection
SRC_WSHEET.Unprotect password:=PASSWORD_STR
WSHEET_UNPROTECT_FUNC = True

Exit Function
ERROR_LABEL:
WSHEET_UNPROTECT_FUNC = False
End Function

Function WSHEET_RENAME_FUNC(ByVal SHEET_NAME As String, _
Optional ByRef SRC_WBOOK As Excel.Workbook)

''on insertion, forces name change
''if valid name, changes, otherwise loops

Dim SRC_WSHEET As Excel.Worksheet

On Error GoTo ERROR_LABEL

WSHEET_RENAME_FUNC = False
If SRC_WBOOK Is Nothing Then: Set SRC_WBOOK = ActiveWorkbook
        
        'verifies that name is not sheet
        If LCase(Left(SHEET_NAME, 5)) = "sheet" _
            Or LCase(Left(SHEET_NAME, 5)) = "chart" Then
            MsgBox _
            "Worksheet names cannot begin with the word 'sheet' or 'chart'.", _
                    vbOKOnly, "Worksheet Name"
            'resets variable to nothing to continue loop
            GoTo ERROR_LABEL
        Else
            'loop through worksheets and verifies name is unique
            For Each SRC_WSHEET In SRC_WBOOK.Worksheets
                'Replace illegal strings values with underscore
                SHEET_NAME = WSHEET_CLEAN_NAME_FUNC(SHEET_NAME, "_")
                
                If LCase(Left(SHEET_NAME, 31)) = LCase(SRC_WSHEET.name) Then
                    'if name exists, notifies user, resets variable
                    MsgBox _
                    "Worksheet name already exists. Please provide a unique name.", _
                            vbOKOnly, "Worksheet Name"
                    'resets string to nothing
                    GoTo ERROR_LABEL
                End If
            Next SRC_WSHEET
        End If

Set SRC_WSHEET = SRC_WBOOK.ActiveSheet
SRC_WSHEET.name = Left(SHEET_NAME, 31)

WSHEET_RENAME_FUNC = True
Exit Function
ERROR_LABEL:
WSHEET_RENAME_FUNC = False
End Function

'************************************************************************************
'************************************************************************************



'************************************************************************************
'************************************************************************************
'FUNCTION      : WSHEET_CLEAN_NAME_FUNC
'DESCRIPTION   : Function accepts string and replacement value; replaces string with
'value for a set of illegal characters, i.e.,[,],\,/,:, ', ?, and *.
'LIBRARY       : EXCEL
'GROUP         : RENAME
'ID            : 002



'AUTHOR        : RAFAEL NICOLAS FERMIN COTA

'************************************************************************************
'************************************************************************************

Private Function WSHEET_CLEAN_NAME_FUNC(ByVal SHEET_NAME As String, _
ByVal REF_CHR As String)

''Replace illegal strings values with underscore
''changes original string

On Error GoTo ERROR_LABEL

    SHEET_NAME = Replace(SHEET_NAME, "*", REF_CHR, 1, -1, 0)
    SHEET_NAME = Replace(SHEET_NAME, "\", REF_CHR, 1, -1, 0)
    SHEET_NAME = Replace(SHEET_NAME, ":", REF_CHR, 1, -1, 0)
    SHEET_NAME = Replace(SHEET_NAME, "'", REF_CHR, 1, -1, 0)
    SHEET_NAME = Replace(SHEET_NAME, "?", REF_CHR, 1, -1, 0)
    SHEET_NAME = Replace(SHEET_NAME, "/", REF_CHR, 1, -1, 0)
    SHEET_NAME = Replace(SHEET_NAME, "[", REF_CHR, 1, -1, 0)
    SHEET_NAME = Replace(SHEET_NAME, "]", REF_CHR, 1, -1, 0)

WSHEET_CLEAN_NAME_FUNC = SHEET_NAME

Exit Function
ERROR_LABEL:
WSHEET_CLEAN_NAME_FUNC = Err.number
End Function



'************************************************************************************
'************************************************************************************
'FUNCTION      : WSHEET_SEARCH_RECORD_FUNC
'DESCRIPTION   : This searches the range specified by SRC_RNG and returns a
'Range object that contains all the cells in which REFER_VAL was found. The
'search parameters to this function have the same meaning and effect as they
'do with the Range.Find method. If the value was not found, the function
'return Nothing.
'LIBRARY       : EXCEL
'GROUP         : WSHEET-LOOK
'ID            : 003
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA

'************************************************************************************
'************************************************************************************

Private Function WSHEET_SEARCH_RECORD_FUNC(ByRef SRC_RNG As Excel.Range, _
ByVal REFER_VAL As Variant, _
Optional ByVal LOOK_IN_VAR As XlFindLookIn = xlValues, _
Optional ByVal LOOK_AT_VAR As XlLookAt = xlWhole, _
Optional ByVal SEARCH_ORDER_VAR As XlSearchOrder = xlByRows, _
Optional ByVal MATCH_CASE_FLAG As Boolean = False) As Excel.Range

Dim FOUND_RNG As Excel.Range
Dim FIRST_RNG As Excel.Range
Dim LAST_RNG As Excel.Range
Dim TEMP_RNG As Excel.Range

On Error GoTo ERROR_LABEL

With SRC_RNG
    Set LAST_RNG = .Cells(.Cells.COUNT)
End With
'On Error Resume Next
On Error GoTo 0
Set FOUND_RNG = SRC_RNG.Find(REFER_VAL, _
        LAST_RNG, LOOK_IN_VAR, _
        LOOK_AT_VAR, SEARCH_ORDER_VAR, MATCH_CASE_FLAG)

If Not FOUND_RNG Is Nothing Then
    Set FIRST_RNG = FOUND_RNG
    Set TEMP_RNG = FOUND_RNG
    Set FOUND_RNG = SRC_RNG.FindNext(FOUND_RNG)
    Do Until False ' Loop forever. We'll "Exit Do" when necessary.
        If (FOUND_RNG Is Nothing) Then
            Exit Do
        End If
        If (FOUND_RNG.Address = FIRST_RNG.Address) Then
            Exit Do
        End If
        Set TEMP_RNG = Excel.Application.Union(TEMP_RNG, FOUND_RNG)
        Set FOUND_RNG = SRC_RNG.FindNext(FOUND_RNG)
    Loop
End If
    
Set WSHEET_SEARCH_RECORD_FUNC = TEMP_RNG

Exit Function
ERROR_LABEL:
Set WSHEET_SEARCH_RECORD_FUNC = Nothing
End Function
