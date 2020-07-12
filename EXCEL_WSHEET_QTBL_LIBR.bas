Attribute VB_Name = "EXCEL_WSHEET_QTBL_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.


'************************************************************************************
'************************************************************************************
'FUNCTION      : WSHEET_QTBL_ADD_FUNC
'DESCRIPTION   : Create a Query Table
'LIBRARY       : QUERY TABLE
'GROUP         : ADD DELETE
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'************************************************************************************
'************************************************************************************

Function WSHEET_QTBL_ADD_FUNC(ByVal SHEET_NAME As String, _
Optional ByVal QUERY_ROW As Long = 2, _
Optional ByVal QUERY_COL As Long = 2, _
Optional ByVal URL_STR As String = "URL;", _
Optional ByRef QUERY_WBOOK As Excel.Workbook) As Excel.QueryTable

Dim TEMP_FLAG As Boolean
Dim DST_RNG As Excel.Range
Dim QUERY_TABLE As Excel.QueryTable
Dim QUERY_SHEET As Excel.Worksheet

On Error GoTo ERROR_LABEL

If QUERY_WBOOK Is Nothing Then: Set QUERY_WBOOK = ActiveWorkbook

If WSHEET_LOOK_FUNC(SHEET_NAME, 0, QUERY_WBOOK) = False _
    Then: GoTo ERROR_LABEL
    
    Set QUERY_SHEET = QUERY_WBOOK.Worksheets(SHEET_NAME)
    Set DST_RNG = QUERY_SHEET.Cells(QUERY_ROW, QUERY_COL)
    TEMP_FLAG = WSHEET_QTBL_RNG_CHECK_FUNC(DST_RNG)
    Do While TEMP_FLAG = True
        Set DST_RNG = DST_RNG.Offset(0, 1)
        TEMP_FLAG = WSHEET_QTBL_RNG_CHECK_FUNC(DST_RNG)
    Loop
    Set QUERY_TABLE = QUERY_SHEET.QueryTables.Add(URL_STR, DST_RNG)
    
    Set WSHEET_QTBL_ADD_FUNC = QUERY_TABLE
Exit Function
ERROR_LABEL:
Set WSHEET_QTBL_ADD_FUNC = Nothing
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : WSHEET_QTBL_REMOVE_FUNC
'DESCRIPTION   : Delete a Query Table
'LIBRARY       : QUERY TABLE
'GROUP         : ADD DELETE
'ID            : 002
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA

'************************************************************************************
'************************************************************************************

Function WSHEET_QTBL_REMOVE_FUNC(ByVal QUERY_NAME As String, _
Optional ByRef QUERY_SHEET As Excel.Worksheet)

'Dim i As Long
Dim j As Long
Dim k As Long

Dim POS_ARR As Variant
Dim TEMP_ARR As Variant

On Error GoTo ERROR_LABEL

WSHEET_QTBL_REMOVE_FUNC = False
If QUERY_SHEET Is Nothing Then: Set QUERY_SHEET = ActiveSheet

TEMP_ARR = WSHEET_QTBL_FIND_FUNC(QUERY_NAME, QUERY_SHEET)
'i = TEMP_ARR(LBound(TEMP_ARR))
j = TEMP_ARR(LBound(TEMP_ARR) + 1)
POS_ARR = TEMP_ARR(UBound(TEMP_ARR) - 1)
k = TEMP_ARR(UBound(TEMP_ARR))

If k = 0 And j = 0 Then
    WSHEET_QTBL_REMOVE_FUNC = False
ElseIf k = 1 And j = 0 Then
    QUERY_SHEET.QueryTables(POS_ARR(LBound(POS_ARR))).Delete
    WSHEET_QTBL_REMOVE_FUNC = True
Else 'If k = 0 And j > 1 Then
    QUERY_SHEET.QueryTables(POS_ARR(j)).Delete
    WSHEET_QTBL_REMOVE_FUNC = True
End If

Exit Function
ERROR_LABEL:
WSHEET_QTBL_REMOVE_FUNC = False
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : WSHEET_QTBLS_REMOVE_FUNC
'DESCRIPTION   : Delete all Queries in a Worksheet
'LIBRARY       : QUERY TABLE
'GROUP         : ADD DELETE
'ID            : 003
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA

'************************************************************************************
'************************************************************************************

Function WSHEET_QTBLS_REMOVE_FUNC(Optional ByRef NAMES_RNG As Variant, _
Optional ByRef QUERY_SHEET As Excel.Worksheet)

Dim i As Long

Dim NAME_STR As String
Dim TEMP_FLAG As Boolean

Dim EACH_QUERY As Excel.QueryTable
        
On Error GoTo ERROR_LABEL

WSHEET_QTBLS_REMOVE_FUNC = False
If QUERY_SHEET Is Nothing Then: Set QUERY_SHEET = ActiveSheet


'--------------------------------------------------------------------------------
If IsArray(NAMES_RNG) = False Then
'--------------------------------------------------------------------------------
    For Each EACH_QUERY In QUERY_SHEET.QueryTables
         EACH_QUERY.Delete
    Next EACH_QUERY
'--------------------------------------------------------------------------------
Else
'--------------------------------------------------------------------------------
    For Each EACH_QUERY In QUERY_SHEET.QueryTables

          TEMP_FLAG = False
          NAME_STR = EACH_QUERY.name
          If IsArray(NAMES_RNG) = True Then
              For i = LBound(NAMES_RNG) To UBound(NAMES_RNG)
                   If NAMES_RNG(i) = NAME_STR Then
                       TEMP_FLAG = True
                       Exit For
                   End If
              Next i
          End If
          If TEMP_FLAG = False Then: EACH_QUERY.Delete
    Next EACH_QUERY
'--------------------------------------------------------------------------------
End If
'--------------------------------------------------------------------------------

WSHEET_QTBLS_REMOVE_FUNC = True

Exit Function
ERROR_LABEL:
WSHEET_QTBLS_REMOVE_FUNC = False
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : WSHEET_QTBLS_COUNT_FUNC
'DESCRIPTION   : Count No. Queries in a Worksheet
'LIBRARY       : QUERY TABLE
'GROUP         : COUNT
'ID            : 001



'AUTHOR        : RAFAEL NICOLAS FERMIN COTA

'************************************************************************************
'************************************************************************************

Function WSHEET_QTBLS_COUNT_FUNC(Optional ByRef QUERY_SHEET As Excel.Worksheet)

Dim i As Long
Dim EACH_QUERY As Excel.QueryTable

On Error GoTo ERROR_LABEL

If QUERY_SHEET Is Nothing Then: Set QUERY_SHEET = ActiveSheet

i = 0
For Each EACH_QUERY In QUERY_SHEET.QueryTables
    i = i + 1
Next EACH_QUERY

WSHEET_QTBLS_COUNT_FUNC = i

Exit Function
ERROR_LABEL:
WSHEET_QTBLS_COUNT_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : WSHEET_QTBL_FORMAT_FUNC
'DESCRIPTION   : Format Query Parameters
'LIBRARY       : QUERY TABLE
'GROUP         : FORMAT
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA

'************************************************************************************
'************************************************************************************

Function WSHEET_QTBL_FORMAT_FUNC(ByRef QUERY_TABLE As Excel.QueryTable, _
Optional ByVal SAVE_OPT As Boolean = True, _
Optional ByVal BACK_OPT As Boolean = True, _
Optional ByVal FORM_INDEX As Integer = 0, _
Optional ByVal REFR_INDEX As Integer = 1, _
Optional ByVal SELEC_INDEX As Integer = 1, _
Optional ByVal REFR_PER As Integer = 15, _
Optional ByVal REFER_TABLES As String = "0,1", _
Optional ByVal PRESER_FORM As Boolean = True, _
Optional ByVal WIDTH_FORM As Boolean = True, _
Optional ByVal ENAB_EDIT As Boolean = False, _
Optional ByVal ENAB_REFR As Boolean = True)

On Error GoTo ERROR_LABEL

WSHEET_QTBL_FORMAT_FUNC = False
    
    With QUERY_TABLE

        .EnableEditing = ENAB_EDIT 'True if the user can edit the
        'specified query table. False if the user can only refresh
        'the query table. Read/write Boolean

        .PreserveFormatting = PRESER_FORM 'For PivotTable reports, this property is
        'True if formatting is preserved when the report is refreshed or recalculated
        'by operations such as pivoting, sorting, or changing page field items.
        'For query tables, this property is True if any formatting common to the
        'first five rows of data are applied to new rows of data in the query table.
        'Unused cells aren’t formatted. The property is False if the last AutoFormat
        'applied to the query table is applied to new rows of data. The default value
        'is True (unless the query table was created in Microsoft Excel 97 and the
        'HasAutoFormat property is True, in which case PreserveFormatting is False
        
        .EnableRefresh = ENAB_REFR 'True if the PivotTable cache or query table
        'can be refreshed by the user. The default value is True. Read/write Boolean.

        .BackgroundQuery = BACK_OPT 'True if queries for the PivotTable report or
        'query table are performed asynchronously (in the background). Read/write
        'Boolean.
                
        If REFR_INDEX = 0 Then 'Returns or sets the way rows on the specified
        'worksheet are added or deleted to accommodate the number of rows in
        'a recordset returned by a query

            .RefreshStyle = xlOverwriteCells 'No new cells or rows are added
            'to the worksheet. Data in surrounding cells is overwritten to
            'accommodate any overflow
        ElseIf REFR_INDEX = 1 Then
            .RefreshStyle = xlInsertDeleteCells 'Partial rows are inserted
            'or deleted to match the exact number of rows required for the
            'new recordset.
        Else
            .RefreshStyle = xlInsertEntireRows 'Entire rows are inserted,
            'if necessary, to accommodate any overflow. No cells or rows
            'are deleted from the worksheet
        End If
        
        .SaveData = SAVE_OPT 'True if data for the PivotTable report is
        'saved with the workbook. False if only the report definition is
        'saved. Read/write Boolean.
        
        .AdjustColumnWidth = WIDTH_FORM 'True if the column widths are
        'automatically adjusted for the best fit each time you refresh
        'the specified query table or XML map. False if the column
        'widths aren’t automatically adjusted with each refresh. The
        'default value is True. Read/write Boolean.
        
        .RefreshPeriod = REFR_PER 'Returns or sets the number of minutes
        'between refreshes. Read/write Long.

        If SELEC_INDEX = 0 Then 'Returns or sets a value that determines
        'whether an entire Web page, all tables on the Web page, or only
        'specific tables on the Web page are imported into a query table.
            .WebSelectionType = xlEntirePage
        ElseIf SELEC_INDEX = 1 Then
            .WebSelectionType = xlSpecifiedTables
            .WebTables = REFER_TABLES 'Returns or sets a comma-delimited
            'list of table names or table index numbers when you import
            'a Web page into a query table.
        Else
            .WebSelectionType = xlAllTables
        End If
        
'---------------------------------------------------------------------------------------
            .FieldNames = True 'if field names from the data source appear as
            'column headings for the returned data.
'---------------------------------------------------------------------------------------
            .RowNumbers = False 'True if row numbers are added as the first
            'column of the specified query table
'---------------------------------------------------------------------------------------
        
        If FORM_INDEX = 0 Then 'Returns or sets a value that determines how
        'much formatting from a Web page, if any, is applied when you import
        'the page into a query table. Read/write
            .WebFormatting = xlWebFormattingNone
        ElseIf FORM_INDEX = 1 Then
            .WebFormatting = xlWebFormattingRTF
        Else 'If FORM_INDEX = 2
            .WebFormatting = xlWebFormattingAll
            .WebPreFormattedTextToColumns = True 'Returns or sets whether data
            'contained within HTML <PRE> tags in the Web page is parsed into columns
            'when you import the page into a query table
            .WebConsecutiveDelimitersAsOne = True 'True if consecutive delimiters
            'are treated as a single delimiter when you import data from HTML <PRE>
            'tags in a Web page into a query table, and if the data is to be
            'parsed into columns. False if you want to treat consecutive delimiters as
            'multiple delimiters
            .WebSingleBlockTextImport = True 'True if data from the HTML <PRE> tags
            'in the specified Web page is processed all at once when you import the
            'page into a query table. False if the data is imported in blocks of
            'contiguous rows so that header rows will be recognized as such.
            .WebDisableDateRecognition = False 'True if data that resembles dates is
            'parsed as text when you import a Web page into a query table. False if
            'date recognition is used
        End If
            .WebDisableRedirections = False
    End With

WSHEET_QTBL_FORMAT_FUNC = True

Exit Function
ERROR_LABEL:
WSHEET_QTBL_FORMAT_FUNC = False
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : WSHEET_QTBL_CTRL_FUNC
'DESCRIPTION   : Set Query Table Frame
'LIBRARY       : QUERY TABLE
'GROUP         : FRAME
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA

'************************************************************************************
'************************************************************************************


Function WSHEET_QTBL_CTRL_FUNC( _
Optional ByRef QUERY_SHEET As Excel.Worksheet, _
Optional ByRef QUERY_TABLE As Excel.QueryTable, _
Optional ByVal QUERY_NAME As Variant = Null, _
Optional ByVal QUERY_ROW As Long = 2, _
Optional ByVal QUERY_COL As Long = 2, _
Optional ByRef QUERY_WBOOK As Excel.Workbook, _
Optional ByVal VERSION As Integer = 0)

Dim i As Long
Dim j As Long
Dim k As Long
Dim h As Long

Dim ENAB_REFR As Boolean
Dim BACK_OPT As Boolean

Dim SAVE_OPT As Boolean
Dim FORM_INDEX As Integer
Dim REFR_INDEX As Integer
Dim SELEC_INDEX As Integer
Dim REFR_PER As Integer
Dim REFER_TABLES As String
Dim PRESER_FORM As Boolean
Dim WIDTH_FORM As Boolean
Dim ENAB_EDIT As Boolean

Dim TEMP_ARR As Variant
Dim TEMP_FLAG As Boolean

On Error GoTo ERROR_LABEL

WSHEET_QTBL_CTRL_FUNC = False

Select Case VERSION
'---------------------------------------------------------------------
Case 0 'Useful for historical data sets
'---------------------------------------------------------------------
    ENAB_REFR = True
    BACK_OPT = False
'---------------------------------------------------------------------
    
    SAVE_OPT = False
    FORM_INDEX = 0
    REFR_INDEX = 0
    SELEC_INDEX = 0
    REFR_PER = 15
    REFER_TABLES = ""
    PRESER_FORM = True
    WIDTH_FORM = False
    ENAB_EDIT = True
'---------------------------------------------------------------------
Case 1 'Useful for quotes
'---------------------------------------------------------------------
    
    ENAB_REFR = True
    BACK_OPT = True
'---------------------------------------------------------------------
    
    SAVE_OPT = False
    FORM_INDEX = 0 'Perfect
    REFR_INDEX = 0 'Perfect
    SELEC_INDEX = 0 'Perfect
    REFR_PER = 15 'Perfect
    REFER_TABLES = ""
    PRESER_FORM = True
    WIDTH_FORM = False
    ENAB_EDIT = True
'---------------------------------------------------------------------
Case Else
'---------------------------------------------------------------------

'---------------------------------------------------------------------
End Select
'---------------------------------------------------------------------

'---------------------------------------------------------------------
If QUERY_TABLE Is Nothing Then
'---------------------------------------------------------------------
   If QUERY_WBOOK Is Nothing Then: Set QUERY_WBOOK = ActiveWorkbook
   
   If VarType(QUERY_NAME) <> vbNull Then 'Historical
        
           TEMP_ARR = WSHEET_QTBL_FIND_FUNC(QUERY_NAME, QUERY_SHEET)
           j = TEMP_ARR(UBound(TEMP_ARR)) 'NICO
           i = TEMP_ARR(LBound(TEMP_ARR) + 1) 'NICO_X
       
           If j = 0 And i = 0 Then
              Set QUERY_TABLE = _
                WSHEET_QTBL_ADD_FUNC(QUERY_SHEET.name, QUERY_ROW, _
                QUERY_COL, "URL;", QUERY_WBOOK)
              QUERY_TABLE.name = QUERY_NAME
           ElseIf j = 1 And i = 0 Then
              Set QUERY_TABLE = QUERY_SHEET.QueryTables(LBound(TEMP_ARR(3)))
           Else
              h = 0
              For k = LBound(TEMP_ARR(3)) To UBound(TEMP_ARR(3))
                    If TEMP_ARR(3)(k) = QUERY_NAME Then
                        h = k
                        Exit For
                    End If
              Next k
              If h = 0 Then: h = UBound(TEMP_ARR(3))
              Set QUERY_TABLE = QUERY_SHEET.QueryTables(h)
           End If
    Else 'Quotes
           Set QUERY_TABLE = WSHEET_QTBL_ADD_FUNC(QUERY_SHEET.name, QUERY_ROW, _
                            QUERY_COL, "URL;", QUERY_WBOOK)
    End If
'---------------------------------------------------------------------
End If
'---------------------------------------------------------------------

TEMP_FLAG = WSHEET_QTBL_FORMAT_FUNC(QUERY_TABLE, SAVE_OPT, BACK_OPT, _
                    FORM_INDEX, REFR_INDEX, SELEC_INDEX, REFR_PER, _
                    REFER_TABLES, PRESER_FORM, WIDTH_FORM, _
                    ENAB_EDIT, ENAB_REFR)
If TEMP_FLAG = False Then: GoTo ERROR_LABEL


WSHEET_QTBL_CTRL_FUNC = True

Exit Function
ERROR_LABEL:
WSHEET_QTBL_CTRL_FUNC = False
End Function
      
'************************************************************************************
'************************************************************************************
'FUNCTION      : WSHEET_QTBL_FIND_FUNC
'DESCRIPTION   : LOOK FOR A QUERY MAP
'LIBRARY       : QUERY TABLE
'GROUP         : LOOK
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA

'************************************************************************************
'************************************************************************************

Function WSHEET_QTBL_FIND_FUNC(ByVal QUERY_NAME As String, _
Optional ByRef QUERY_SHEET As Excel.Worksheet)

Dim i As Long
Dim hh As Long
Dim POS_ARR As Variant
Dim j As Long

Dim NAME_STR As String
Dim MATCH_FLAG As Boolean

Dim SUFFIX_STR As String

Dim EACH_QUERY As Excel.QueryTable

On Error GoTo ERROR_LABEL

MATCH_FLAG = False
If QUERY_SHEET Is Nothing Then: Set QUERY_SHEET = ActiveSheet

hh = 0: i = 0: j = 0
ReDim POS_ARR(0 To 0)
'-------------------------------------------------------------------------------------------
For Each EACH_QUERY In QUERY_SHEET.QueryTables
      hh = hh + 1
      MATCH_FLAG = False
'-------------------------------------------------------------------------------------------
      NAME_STR = EACH_QUERY.name
'-------------------------------------------------------------------------------------------
      If NAME_STR = QUERY_NAME Then
       MATCH_FLAG = True
       j = 1
       POS_ARR(0) = hh
      End If
      
'-------------------------------------------------------------------------------------------
      If COUNT_STRING_FUNC(NAME_STR, "_", "_", 1, 1) > 0 Then
         SUFFIX_STR = Mid(NAME_STR, InStr(1, NAME_STR, "_", 0), Len(NAME_STR))
         NAME_STR = Trim(Replace(NAME_STR, SUFFIX_STR, "", 1, -1, 1))
      End If
      If NAME_STR Like QUERY_NAME Then: MATCH_FLAG = True
      If (MATCH_FLAG = True) Then
        i = i + 1
        ReDim Preserve POS_ARR(0 To i)
        POS_ARR(i) = hh
      End If
'-------------------------------------------------------------------------------------------
Next EACH_QUERY
'-------------------------------------------------------------------------------------------
WSHEET_QTBL_FIND_FUNC = Array(hh, i, POS_ARR, j)
'-------------------------------------------------------------------------------------------
'hh -> COUNTER
'i -> NO. OF NAMES with "_"
'POS_ARR -> Array of position
'j -> Dummy for Variable

Exit Function
ERROR_LABEL:
    WSHEET_QTBL_FIND_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : WSHEET_QTBL_NAME_FUNC
'DESCRIPTION   : CHECK QUERY TABLE NAME
'LIBRARY       : QUERY TABLE
'GROUP         : NAME
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA

'************************************************************************************
'************************************************************************************

Function WSHEET_QTBL_NAME_FUNC( _
Optional ByVal QUERY_NAME As String = "ExternalData", _
Optional ByRef QUERY_SHEET As Excel.Worksheet)

Dim j As Long
Dim k As Long

Dim TEMP_ARR As Variant

On Error GoTo ERROR_LABEL

If QUERY_SHEET Is Nothing Then: Set QUERY_SHEET = ActiveSheet

TEMP_ARR = WSHEET_QTBL_FIND_FUNC(QUERY_NAME, QUERY_SHEET)

'i = TEMP_ARR(LBound(TEMP_ARR))
k = TEMP_ARR(UBound(TEMP_ARR)) 'NICO
j = TEMP_ARR(LBound(TEMP_ARR) + 1) 'NICO_

If k = 0 And j = 0 Then
    WSHEET_QTBL_NAME_FUNC = QUERY_NAME & "_" & 1
ElseIf k = 1 And j = 0 Then
    WSHEET_QTBL_NAME_FUNC = QUERY_NAME & "_" & 1
ElseIf k = 0 And j > 0 Then
    WSHEET_QTBL_NAME_FUNC = QUERY_NAME & "_" & (j + 1)
ElseIf k = 1 And j > 0 Then
    WSHEET_QTBL_NAME_FUNC = QUERY_NAME & "_" & (j + 1 - k)
End If

Exit Function
ERROR_LABEL:
WSHEET_QTBL_NAME_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : WSHEET_QTBL_PARAMETERS_FUNC
'DESCRIPTION   : Defines a parameter for the specified query table
'LIBRARY       : QUERY TABLE
'GROUP         : PARAMETERS
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA

'************************************************************************************
'************************************************************************************

Function WSHEET_QTBL_PARAMETERS_FUNC( _
ByRef QUERY_TABLE As Excel.QueryTable, _
ByRef QUERY_VALUE As Variant, _
ByVal QUERY_PARAM As Variant, _
Optional ByVal VERSION As Integer = 0)

Dim SRC_PARAM As Parameter

On Error GoTo ERROR_LABEL

'QUERY_VALUE: The value of the specified parameter, as shown in the
'description of the Type argument.

WSHEET_QTBL_PARAMETERS_FUNC = False
Set SRC_PARAM = QUERY_TABLE.Parameters.Add(QUERY_PARAM, xlParamTypeVarChar) _
'--> YOU CAN CHANGE THIS
    
    Select Case VERSION 'can be one of these XlParameterType constants
'----------------------------------------------------------------------------
        Case 0 'Uses the value specified by the Value argument
            SRC_PARAM.SetParam xlConstant, QUERY_VALUE
'----------------------------------------------------------------------------
        Case 1 'Displays a dialog box that prompts the user for the _
        value. The Value argument specifies the text shown in the _
        dialog box
            SRC_PARAM.SetParam xlPrompt, QUERY_VALUE
'----------------------------------------------------------------------------
        Case Else 'Uses the value of the cell in the upper-left _
        corner of the range. The Value argument specifies a Range _
        object
            SRC_PARAM.SetParam xlRange, QUERY_VALUE
'----------------------------------------------------------------------------
    End Select

WSHEET_QTBL_PARAMETERS_FUNC = True

Exit Function
ERROR_LABEL:
WSHEET_QTBL_PARAMETERS_FUNC = False
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : WSHEET_QTBL_PARAMETERS_TABLE_FUNC
'DESCRIPTION   : List Query Parameters
'LIBRARY       : QUERY TABLE
'GROUP         : PARAMETERS
'ID            : 002
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA

'************************************************************************************
'************************************************************************************

Function WSHEET_QTBL_PARAMETERS_TABLE_FUNC( _
Optional ByRef QUERY_SHEET As Excel.Worksheet)

  Dim i As Long
  
  Dim NROWS As Long
  Dim NCOLUMNS As Long

  Dim nQuery As Excel.QueryTable
  
  Dim TEMP_ARR As Variant
  
  On Error GoTo ERROR_LABEL
  
  If QUERY_SHEET Is Nothing Then: Set QUERY_SHEET = ActiveSheet
  
  If QUERY_SHEET.QueryTables.COUNT = 0 Then
        WSHEET_QTBL_PARAMETERS_TABLE_FUNC = ".... not WebQueries"
      Exit Function
  End If
  
  NROWS = 7
  NCOLUMNS = QUERY_SHEET.QueryTables.COUNT + 1
  
  ReDim TEMP_ARR(0 To NROWS, 1 To NCOLUMNS)

  TEMP_ARR(0, 1) = "Worksheet Name"
  TEMP_ARR(1, 1) = "Query Name"
  TEMP_ARR(2, 1) = "Query Selection"
  TEMP_ARR(3, 1) = "Query Connection"
  TEMP_ARR(4, 1) = "Query Destination"
  TEMP_ARR(5, 1) = "Query Result Range"
  TEMP_ARR(6, 1) = "Query Adjacent Formulas"
  TEMP_ARR(7, 1) = "Query Parameters Count"

    i = 1
    For Each nQuery In QUERY_SHEET.QueryTables
      
      i = i + 1
      TEMP_ARR(0, i) = QUERY_SHEET.name
      
          Select Case nQuery.QueryType
            Case xlWebQuery
              TEMP_ARR(1, i) = "Query(" & i - 1 & ")" & " ; " & nQuery.name
            Case Else
              TEMP_ARR(1, i) = "..not WebQuery"
                  GoTo 1983
          End Select
      
          Select Case nQuery.WebSelectionType
            Case xlEntirePage
              TEMP_ARR(2, i) = "Entire Page"
            Case xlAllTables
              TEMP_ARR(2, i) = "All Tables"
            Case xlSpecifiedTables
              TEMP_ARR(2, i) = "Specified Table(s): " & nQuery.WebTables
          End Select
          
      TEMP_ARR(3, i) = nQuery.Connection
      TEMP_ARR(4, i) = nQuery.Destination.Address
      TEMP_ARR(5, i) = nQuery.ResultRange.Address
      TEMP_ARR(6, i) = nQuery.FillAdjacentFormulas
      TEMP_ARR(7, i) = nQuery.Parameters.COUNT
1983:
    Next nQuery
    
    WSHEET_QTBL_PARAMETERS_TABLE_FUNC = TEMP_ARR
    
Exit Function
ERROR_LABEL:
    WSHEET_QTBL_PARAMETERS_TABLE_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : WSHEET_QTBL_CONNECTION_FUNC
'DESCRIPTION   : Establish Query Connection
'LIBRARY       : QUERY TABLE
'GROUP         : STATUS
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA

'************************************************************************************
'************************************************************************************

Function WSHEET_QTBL_CONNECTION_FUNC(ByRef QUERY_TABLE As Excel.QueryTable, _
ByVal QUERY_URL As String, _
Optional ByVal URL_STR As String = "URL;")

On Error GoTo ERROR_LABEL

WSHEET_QTBL_CONNECTION_FUNC = False
QUERY_TABLE.Connection = URL_STR & QUERY_URL
WSHEET_QTBL_CONNECTION_FUNC = True

Exit Function
ERROR_LABEL:
WSHEET_QTBL_CONNECTION_FUNC = False
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : WSHEET_QTBL_REFRESH_FUNC
'DESCRIPTION   : Refresh Query Connection
'LIBRARY       : QUERY TABLE
'GROUP         : STATUS
'ID            : 002
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA

'************************************************************************************
'************************************************************************************

'// PERFECT

Function WSHEET_QTBL_REFRESH_FUNC(ByRef QUERY_TABLE As Excel.QueryTable, _
Optional ByVal STATUS_OPT As Boolean = False)

On Error GoTo ERROR_LABEL

WSHEET_QTBL_REFRESH_FUNC = False
    If STATUS_OPT = (True) Then: Call WSHEET_QTBL_STATUS_FUNC(QUERY_TABLE)
    QUERY_TABLE.Refresh
    Excel.Application.StatusBar = False
WSHEET_QTBL_REFRESH_FUNC = True

Exit Function
ERROR_LABEL:
WSHEET_QTBL_REFRESH_FUNC = False
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : WSHEET_QTBLS_REFRESH_FUNC
'DESCRIPTION   : Refresh ALL Query Connections
'LIBRARY       : QUERY TABLE
'GROUP         : STATUS
'ID            : 003
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA

'************************************************************************************
'************************************************************************************

Function WSHEET_QTBLS_REFRESH_FUNC(ByRef QUERY_SHEET As Excel.Worksheet, _
Optional ByVal STATUS_OPT As Boolean = False)

Dim EACH_QUERY As Excel.QueryTable

On Error GoTo ERROR_LABEL

WSHEET_QTBLS_REFRESH_FUNC = False
For Each EACH_QUERY In QUERY_SHEET.QueryTables
     If STATUS_OPT = (True) Then: Call WSHEET_QTBL_STATUS_FUNC(EACH_QUERY)
     EACH_QUERY.Refresh
     Excel.Application.StatusBar = False
Next EACH_QUERY
WSHEET_QTBLS_REFRESH_FUNC = True

Exit Function
ERROR_LABEL:
WSHEET_QTBLS_REFRESH_FUNC = False
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : WSHEET_QTBL_STATUS_FUNC
'DESCRIPTION   : Print Query Status Bar
'LIBRARY       : QUERY TABLE
'GROUP         : STATUS
'ID            : 004
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA

'************************************************************************************
'************************************************************************************

Function WSHEET_QTBL_STATUS_FUNC(ByRef QUERY_TABLE As Excel.QueryTable)
On Error GoTo ERROR_LABEL
WSHEET_QTBL_STATUS_FUNC = False
    Excel.Application.StatusBar = "Requesting " & _
        CStr(QUERY_TABLE.Connection) & "..."
WSHEET_QTBL_STATUS_FUNC = True
Exit Function
ERROR_LABEL:
WSHEET_QTBL_STATUS_FUNC = False
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : WSHEET_QTBL_RNG_CHECK_FUNC
'DESCRIPTION   : Check if there is a query in a cell
'LIBRARY       : QUERY TABLE
'GROUP         : VALIDATION
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA

'************************************************************************************
'************************************************************************************

Function WSHEET_QTBL_RNG_CHECK_FUNC(ByRef QUERY_RNG As Excel.Range)

  Dim nQuery As Excel.QueryTable
  Dim TEMP_MATCH As Boolean
  
  On Error GoTo ERROR_LABEL
  
  TEMP_MATCH = False 'Find the QueryTable containing the active cell (if any).
  For Each nQuery In QUERY_RNG.Worksheet.QueryTables
    If nQuery.QueryType <> xlWebQuery Then: GoTo 1983
    If nQuery.Parameters.COUNT <> 0 Then: GoTo 1983
    
    If Not (Intersect(nQuery.ResultRange, QUERY_RNG) Is Nothing) Then
      TEMP_MATCH = True
      Exit For
    End If
1983:
  Next nQuery
  
  WSHEET_QTBL_RNG_CHECK_FUNC = TEMP_MATCH 'IF TEMP_MATCH = False Then: Active Cell _
  isn't within any Query Table

Exit Function
ERROR_LABEL:
WSHEET_QTBL_RNG_CHECK_FUNC = False
End Function
