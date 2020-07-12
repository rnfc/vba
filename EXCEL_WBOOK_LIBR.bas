Attribute VB_Name = "EXCEL_WBOOK_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.

                            
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.

'************************************************************************************
'************************************************************************************
'FUNCTION      : WBOOK_GET_DATA_FUNC
'DESCRIPTION   : Find specific information in a workbook by referring to a
'kew word, and then copy the information found in the destination range
'LIBRARY       : WORKBOOK
'GROUP         : BOOLEAN
'ID            : 001

'AUTHOR        : RAFAEL NICOLAS FERMIN COTA

'************************************************************************************
'************************************************************************************

Function WBOOK_GET_DATA_FUNC(ByRef DST_RNG As Excel.Range, _
ByVal KEY_WORD As String, _
ByVal LOAD_OPT As Integer, _
ByVal SRC_ROW As Long, _
ByVal SRC_COLUMN As Long, _
ByVal SRC_URL_STR As String, _
Optional ByVal SHEET_NAME As Variant = 1)

Dim i As Long
Dim j As Long

Dim NROWS As Long
Dim NCOLUMNS As Long

Dim TEMP_RNG As Excel.Range
Dim TEMP_MATRIX As Variant

Dim SRC_WBOOK As Excel.Workbook

On Error GoTo ERROR_LABEL

WBOOK_GET_DATA_FUNC = False

Set SRC_WBOOK = Workbooks.Open(SRC_URL_STR)

Set TEMP_RNG = SRC_WBOOK.Worksheets(SHEET_NAME).UsedRange

TEMP_MATRIX = RNG_FIND_POSITION_FUNC(KEY_WORD, TEMP_RNG, LOAD_OPT, SRC_ROW, SRC_COLUMN, 0, 0)

SRC_WBOOK.Close (False)

NROWS = UBound(TEMP_MATRIX, 1)
NCOLUMNS = UBound(TEMP_MATRIX, 2)

For j = 1 To NCOLUMNS
    For i = 1 To NROWS
        DST_RNG.Cells(i, j) = TEMP_MATRIX(i, j)
    Next i
Next j

WBOOK_GET_DATA_FUNC = True

Exit Function
ERROR_LABEL:
WBOOK_GET_DATA_FUNC = False
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : WBOOK_COUNT_SHEETS_FUNC
'DESCRIPTION   : COUNT WORKSHEETS WITHIN MULT. WORKBOOKS
'LIBRARY       : WORKBOOKS
'GROUP         : COUNT
'ID            : 001



'AUTHOR        : RAFAEL NICOLAS FERMIN COTA

'************************************************************************************
'************************************************************************************


Function WBOOK_COUNT_SHEETS_FUNC(Optional ByVal OUTPUT As Integer = 0)
  
  Dim hh As Long
  Dim ii As Long
  Dim jj As Long
  
  Dim TEMP_STR As String
  Dim TEMP_ARR As Variant
  Dim SRC_WBOOK As Excel.Workbook
  Dim SRC_WSHEET As Excel.Worksheet
    
  On Error GoTo ERROR_LABEL
  
  jj = 10
  ii = Workbooks.COUNT * jj
  ReDim TEMP_ARR(0 To ii)

  hh = -1
  For Each SRC_WBOOK In Workbooks
    If SRC_WBOOK.name <> ThisWorkbook.name Then
      TEMP_STR = "[" & SRC_WBOOK.name & "]"
      For Each SRC_WSHEET In SRC_WBOOK.Worksheets
        If SRC_WSHEET.Visible = True And _
           SRC_WSHEET.ProtectContents = False Then
          hh = hh + 1
          If hh > ii Then
            ii = ii + jj
            ReDim Preserve TEMP_ARR(0 To ii)
          End If
          TEMP_ARR(hh) = TEMP_STR & SRC_WSHEET.name
        End If  'visible, not protected
      Next SRC_WSHEET
    End If  'not ThisWorkbook
  Next SRC_WBOOK

  If hh >= 0 Then
    ReDim Preserve TEMP_ARR(0 To hh)
    TEMP_ARR = ARRAY_SHELL_SORT_FUNC(TEMP_ARR)
  Else
    ReDim TEMP_ARR(0 To 0)
  End If
  If OUTPUT = 0 Then
      WBOOK_COUNT_SHEETS_FUNC = hh + 1
  Else
      WBOOK_COUNT_SHEETS_FUNC = TEMP_ARR
  End If
Exit Function
ERROR_LABEL:
WBOOK_COUNT_SHEETS_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : WBOOK_FIND_FUNC
'DESCRIPTION   : SEARCHES FOR ALL XLS FILES LOCATED IN THE DRIVE
'LIBRARY       : WORKBOOKS
'GROUP         : LOOK
'ID            : 001



'AUTHOR        : RAFAEL NICOLAS FERMIN COTA

'************************************************************************************
'************************************************************************************

'FULL_PATH_NAME = C:\Documents and Settings\N\Desktop
Function WBOOK_FIND_FUNC(ByVal FULL_PATH_NAME As String, _
Optional ByVal EXT_STR As Variant = ".xls")

Dim i As Long
Dim APPL_OBJ As Object
Dim TEMP_ARR As Variant

On Error GoTo ERROR_LABEL

Set APPL_OBJ = Excel.Application.FileSearch

With APPL_OBJ
    .LookIn = FULL_PATH_NAME
    .FileName = "*" & EXT_STR
    If .Execute > 0 Then
        'MsgBox "There were " & .FoundFiles.Count & _
            " file(s) found."
        ReDim TEMP_ARR(.FoundFiles.COUNT, 1)
        For i = 1 To .FoundFiles.COUNT
            TEMP_ARR(i, 1) = .FoundFiles(i)
        Next i
        WBOOK_FIND_FUNC = TEMP_ARR
    Else
        GoTo ERROR_LABEL
    End If
End With

Exit Function
ERROR_LABEL:
WBOOK_FIND_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : WBOOK_PATH_FUNC
'DESCRIPTION   : DISPLAY THE WORKBOOK PATH
'LIBRARY       : WORKBOOKS
'GROUP         : LOOK
'ID            : 002



'AUTHOR        : RAFAEL NICOLAS FERMIN COTA

'************************************************************************************
'************************************************************************************


Function WBOOK_PATH_FUNC(Optional ByRef SRC_WBOOK As Excel.Workbook, _
Optional ByVal OUTPUT As Integer = 0)

On Error GoTo ERROR_LABEL

If SRC_WBOOK Is Nothing Then: Set SRC_WBOOK = ActiveWorkbook

Select Case OUTPUT
    Case 0
        WBOOK_PATH_FUNC = SRC_WBOOK.Path
    Case Else
        WBOOK_PATH_FUNC = SRC_WBOOK.Path & _
                    Excel.Application.PathSeparator & SRC_WBOOK.name
End Select

Exit Function
ERROR_LABEL:
    WBOOK_PATH_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : WBOOK_PARSE_NAME_FUNC
'DESCRIPTION   : Please select a range of cells from which you want.
'You may click on another sheet in another workbook
'if needed
'LIBRARY       : WORKBOOKS
'GROUP         : NAME
'ID            : 001



'AUTHOR        : RAFAEL NICOLAS FERMIN COTA

'************************************************************************************
'************************************************************************************
                    
Function WBOOK_PARSE_NAME_FUNC(ByRef SRC_RNG As Excel.Range)

Dim ii As Long
Dim jj As Long

On Error GoTo ERROR_LABEL

ii = InStr(1, SRC_RNG.Address(External:=True), "[")
jj = InStr(2, SRC_RNG.Address(External:=True), "]")
WBOOK_PARSE_NAME_FUNC = Mid(SRC_RNG.Address(External:=True), ii + 1, jj - ii - 1)

Exit Function
ERROR_LABEL:
WBOOK_PARSE_NAME_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : WBOOK_COUNT_NAMES_FUNC
'DESCRIPTION   : COUNT NAMES IN A WORKBOOK
'LIBRARY       : WORKBOOKS
'GROUP         : NAME
'ID            : 002



'AUTHOR        : RAFAEL NICOLAS FERMIN COTA

'************************************************************************************
'************************************************************************************

Function WBOOK_COUNT_NAMES_FUNC(Optional ByRef SRC_WBOOK As Excel.Workbook)
On Error GoTo ERROR_LABEL
If SRC_WBOOK Is Nothing Then: Set SRC_WBOOK = ActiveWorkbook
WBOOK_COUNT_NAMES_FUNC = SRC_WBOOK.Names.COUNT
Exit Function
ERROR_LABEL:
WBOOK_COUNT_NAMES_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : WBOOK_REMOVE_NAMES_FUNC
'DESCRIPTION   : DELETE NAMES IN THE SPECIFIED RANGE
'LIBRARY       : WORKBOOKS
'GROUP         : NAME
'ID            : 006



'AUTHOR        : RAFAEL NICOLAS FERMIN COTA

'************************************************************************************
'************************************************************************************

Function WBOOK_REMOVE_NAMES_FUNC(Optional ByRef DATA_ARR As Variant, _
Optional ByRef SRC_WBOOK As Excel.Workbook)

Dim i As Long
Dim EACH_NAME As Excel.name
Dim TEMP_FLAG As Boolean
        
On Error GoTo ERROR_LABEL
                
WBOOK_REMOVE_NAMES_FUNC = False
If SRC_WBOOK Is Nothing Then: Set SRC_WBOOK = ActiveWorkbook

For Each EACH_NAME In SRC_WBOOK.Names
      TEMP_FLAG = False
      If IsArray(DATA_ARR) = True Then
        For i = LBound(DATA_ARR) To UBound(DATA_ARR)
           If DATA_ARR(i) = EACH_NAME.name Then
               TEMP_FLAG = True
               Exit For
           End If
        Next i
      End If
      If TEMP_FLAG = False Then: EACH_NAME.Delete
Next EACH_NAME
WBOOK_REMOVE_NAMES_FUNC = True

Exit Function
ERROR_LABEL:
WBOOK_REMOVE_NAMES_FUNC = False
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : WBOOK_OPEN_CLOSE_FUNC
'DESCRIPTION   : OPEN OR CLOSE A WORKBOOK
'LIBRARY       : WORKBOOKS
'GROUP         : OPEN
'ID            : 001



'AUTHOR        : RAFAEL NICOLAS FERMIN COTA

'************************************************************************************
'************************************************************************************

Function WBOOK_OPEN_CLOSE_FUNC(ByVal FULL_PATH_NAME As String, _
Optional ByVal VERSION As Integer = 0, _
Optional ByVal READ_FLAG As Boolean = False, _
Optional ByVal SAVE_FLAG As Boolean = True)

On Error GoTo ERROR_LABEL

WBOOK_OPEN_CLOSE_FUNC = False

Select Case VERSION
Case 0
    Workbooks.Open (FULL_PATH_NAME), , (READ_FLAG)
Case Else
    Workbooks(FULL_PATH_NAME).Close (SAVE_FLAG)
End Select

WBOOK_OPEN_CLOSE_FUNC = True

Exit Function
ERROR_LABEL:
WBOOK_OPEN_CLOSE_FUNC = False
End Function
