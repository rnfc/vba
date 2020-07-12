Attribute VB_Name = "EXCEL_PIVOT_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.


'************************************************************************************
'************************************************************************************
'FUNCTION      : PIVOT_ADD_FUNC
'DESCRIPTION   : CREATE A PIVOT TABLE
'LIBRARY       : PIVOT
'GROUP         : ADD
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'************************************************************************************
'************************************************************************************

Function PIVOT_ADD_FUNC(ByRef DST_RANGE As Excel.Range, _
ByRef SRC_RANGE As Excel.Range, _
ByVal PIVOT_STR_NAME As String, _
Optional ByRef DST_WBOOK As Excel.Workbook)

On Error GoTo ERROR_LABEL

PIVOT_ADD_FUNC = False
If DST_WBOOK Is Nothing Then: Set DST_WBOOK = ActiveWorkbook

With DST_WBOOK
    .PivotCaches.Add(SourceType:=xlDatabase, SourceData:= _
    SRC_RANGE).CreatePivotTable TableDestination:= _
    DST_RANGE, Tablename:=PIVOT_STR_NAME, DefaultVersion:= _
    xlPivotTableVersion10
End With

PIVOT_ADD_FUNC = True

Exit Function
ERROR_LABEL:
PIVOT_ADD_FUNC = False
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : PIVOT_DATA_FIELD_FUNC
'DESCRIPTION   : ADD DATA FIELD IN A PIVOT TABLE
'LIBRARY       : PIVOT
'GROUP         : ADD
'ID            : 002
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'************************************************************************************
'************************************************************************************

Function PIVOT_DATA_FIELD_FUNC(ByVal PIVOT_STR_NAME As String, _
ByVal FIELD_STR_NAME As String, _
Optional ByVal VERSION As Integer = 1, _
Optional ByRef SRC_WSHEET As Excel.Worksheet)

On Error GoTo ERROR_LABEL

PIVOT_DATA_FIELD_FUNC = False
If SRC_WSHEET Is Nothing Then: Set SRC_WSHEET = ActiveSheet

Select Case VERSION
Case 1
    With SRC_WSHEET.PivotTables(PIVOT_STR_NAME)
        .AddDataField SRC_WSHEET.PivotTables( _
            PIVOT_STR_NAME).PivotFields(FIELD_STR_NAME), , xlCount
    End With
Case 2
    With SRC_WSHEET.PivotTables(PIVOT_STR_NAME)
        .AddDataField SRC_WSHEET.PivotTables( _
            PIVOT_STR_NAME).PivotFields(FIELD_STR_NAME), , xlSum
    End With
Case 3
    With SRC_WSHEET.PivotTables(PIVOT_STR_NAME)
        .AddDataField SRC_WSHEET.PivotTables( _
            PIVOT_STR_NAME).PivotFields(FIELD_STR_NAME), , xlAverage
    End With
Case Else
    GoTo ERROR_LABEL
End Select

PIVOT_DATA_FIELD_FUNC = True

Exit Function
ERROR_LABEL:
PIVOT_DATA_FIELD_FUNC = False
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : PIVOT_REPORT_FUNC
'DESCRIPTION   : FORMAT A PIVOT REPORT
'LIBRARY       : PIVOT
'GROUP         : FORMAT
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'************************************************************************************
'************************************************************************************

Function PIVOT_REPORT_FUNC(ByVal PIVOT_STR_NAME As String, _
Optional ByRef SRC_WSHEET As Excel.Worksheet)

On Error GoTo ERROR_LABEL

PIVOT_REPORT_FUNC = False

If SRC_WSHEET Is Nothing Then: Set SRC_WSHEET = ActiveSheet
SRC_WSHEET.PivotTables(PIVOT_STR_NAME).Format xlReport1
PIVOT_REPORT_FUNC = True

Exit Function
ERROR_LABEL:
PIVOT_REPORT_FUNC = False
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : PIVOT_PARAM_FUNC
'DESCRIPTION   : FORMAT PIVOT TABLE
'LIBRARY       : PIVOT
'GROUP         : FORMAT
'ID            : 002
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'************************************************************************************
'************************************************************************************

Function PIVOT_PARAM_FUNC(ByVal PIVOT_STR_NAME As String, _
Optional ByRef SRC_WSHEET As Excel.Worksheet)

On Error GoTo ERROR_LABEL

If SRC_WSHEET Is Nothing Then: Set SRC_WSHEET = ActiveSheet

PIVOT_PARAM_FUNC = False
With SRC_WSHEET.PivotTables(PIVOT_STR_NAME)
        .DisplayErrorString = True
        .ErrorString = "NA"
        .NullString = "0"
End With
PIVOT_PARAM_FUNC = True

Exit Function
ERROR_LABEL:
PIVOT_PARAM_FUNC = False
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : PIVOT_FORMAT_FUNC
'DESCRIPTION   : FORMAT A PIVOT FIELD
'LIBRARY       : PIVOT
'GROUP         : FORMAT
'ID            : 003
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'************************************************************************************
'************************************************************************************

Function PIVOT_FORMAT_FUNC(ByVal PIVOT_STR_NAME As String, _
ByVal FIELD_STR_NAME As String, _
Optional ByRef SRC_WSHEET As Excel.Worksheet)

On Error GoTo ERROR_LABEL

If SRC_WSHEET Is Nothing Then: Set SRC_WSHEET = ActiveSheet

PIVOT_FORMAT_FUNC = False
With SRC_WSHEET.PivotTables(PIVOT_STR_NAME)
    With .PivotFields(FIELD_STR_NAME)
        .Orientation = xlRowField
        .Position = 1
    End With
End With

PIVOT_FORMAT_FUNC = True

Exit Function
ERROR_LABEL:
PIVOT_FORMAT_FUNC = False
End Function
