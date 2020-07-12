Attribute VB_Name = "EXCEL_CHART_SERIES_LIBR"
'--------------------------------------------------------------------------------------------------------
'--------------------------------------------------------------------------------------------------------
Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.
'--------------------------------------------------------------------------------------------------------
'--------------------------------------------------------------------------------------------------------

'************************************************************************************
'************************************************************************************
'FUNCTION      : EXCEL_CHART_COUNT_SERIES_FUNC
'DESCRIPTION   : COUNT SERIES IN CHART
'LIBRARY       : EXCEL_CHART
'GROUP         : SERIES
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'************************************************************************************
'************************************************************************************

Function EXCEL_CHART_COUNT_SERIES_FUNC(Optional ByVal CHART_NAME_STR As Variant = "", _
Optional ByRef SRC_WSHEET As Excel.Worksheet)

Dim i As Long
Dim CHART_OBJ As Excel.Chart 'Object
Dim SERIE_OBJ As Excel.Series

On Error GoTo ERROR_LABEL

If SRC_WSHEET Is Nothing Then: Set SRC_WSHEET = ActiveSheet
i = SRC_WSHEET.ChartObjects.COUNT
If i = 0 Then: GoTo ERROR_LABEL
If CHART_NAME_STR = "" Then: CHART_NAME_STR = SRC_WSHEET.ChartObjects(i).name
Set CHART_OBJ = SRC_WSHEET.ChartObjects(CHART_NAME_STR).Chart

i = 0
For Each SERIE_OBJ In CHART_OBJ.SeriesCollection: i = i + 1: Next SERIE_OBJ
EXCEL_CHART_COUNT_SERIES_FUNC = i

Exit Function
ERROR_LABEL:
EXCEL_CHART_COUNT_SERIES_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : EXCEL_CHART_LIST_SERIES_FUNC
'DESCRIPTION   :
'LIBRARY       : EXCEL_CHART
'GROUP         : SERIES
'ID            : 002
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'************************************************************************************
'************************************************************************************

Function EXCEL_CHART_LIST_SERIES_FUNC(Optional ByVal CHART_NAME_STR As Variant = "", _
Optional ByVal SRC_WSHEET As Excel.Worksheet)

Dim i As Long
Dim CHART_OBJ As Excel.Chart 'Object
Dim SERIE_OBJ As Excel.Series
Dim TEMP_MATRIX As Variant

On Error GoTo ERROR_LABEL

If SRC_WSHEET Is Nothing Then: Set SRC_WSHEET = ActiveSheet
i = SRC_WSHEET.ChartObjects.COUNT
If i = 0 Then: GoTo ERROR_LABEL
If CHART_NAME_STR = "" Then: CHART_NAME_STR = SRC_WSHEET.ChartObjects(i).name
Set CHART_OBJ = SRC_WSHEET.ChartObjects(CHART_NAME_STR).Chart

i = 1
ReDim TEMP_MATRIX(1 To 2, 1 To i + 1)
TEMP_MATRIX(1, 1) = "SERIES"
TEMP_MATRIX(2, 1) = "POINTS"

For Each SERIE_OBJ In CHART_OBJ.SeriesCollection
    ReDim Preserve TEMP_MATRIX(1 To 2, 1 To i + 1)
    TEMP_MATRIX(1, i + 1) = SERIE_OBJ.formula
    TEMP_MATRIX(2, i + 1) = SERIE_OBJ.Points.COUNT
    i = i + 1
Next SERIE_OBJ

EXCEL_CHART_LIST_SERIES_FUNC = WorksheetFunction.Transpose(TEMP_MATRIX)

Exit Function
ERROR_LABEL:
EXCEL_CHART_LIST_SERIES_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : EXCEL_CHART_LOOK_SERIES_FUNC
'DESCRIPTION   :
'LIBRARY       : EXCEL_CHART
'GROUP         : SERIES
'ID            : 003
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'************************************************************************************
'************************************************************************************

Function EXCEL_CHART_LOOK_SERIES_FUNC(ByVal SERIE_NAME_STR As Variant, _
Optional ByVal CHART_NAME_STR As Variant = "", _
Optional ByVal OUTPUT As Integer = 0, _
Optional ByRef SRC_WSHEET As Excel.Worksheet)

Dim i As Long
Dim CHART_OBJ As Excel.Chart 'Object
Dim SERIE_OBJ As Excel.Series
Dim MATCH_FLAG As Boolean

On Error GoTo ERROR_LABEL

If SRC_WSHEET Is Nothing Then: Set SRC_WSHEET = ActiveSheet
i = SRC_WSHEET.ChartObjects.COUNT
If i = 0 Then: GoTo ERROR_LABEL
If CHART_NAME_STR = "" Then: CHART_NAME_STR = SRC_WSHEET.ChartObjects(i).name
Set CHART_OBJ = SRC_WSHEET.ChartObjects(CHART_NAME_STR).Chart

MATCH_FLAG = False
i = 0
For Each SERIE_OBJ In CHART_OBJ.SeriesCollection
    i = i + 1
    If SERIE_OBJ.name = SERIE_NAME_STR Then
        MATCH_FLAG = True
        Exit For
    End If
Next SERIE_OBJ

If MATCH_FLAG = True Then
    Select Case OUTPUT
    Case 0
        EXCEL_CHART_LOOK_SERIES_FUNC = True
    Case Else
        EXCEL_CHART_LOOK_SERIES_FUNC = i
    End Select
Else
    EXCEL_CHART_LOOK_SERIES_FUNC = False
End If

Exit Function
ERROR_LABEL:
EXCEL_CHART_LOOK_SERIES_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : EXCEL_CHART_ADD_SERIE_FUNC
'DESCRIPTION   :
'LIBRARY       : EXCEL_CHART
'GROUP         : SERIES
'ID            : 004
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'************************************************************************************
'************************************************************************************

Function EXCEL_CHART_ADD_SERIE_FUNC(ByVal SERIE_NAME_STR As Variant, _
Optional ByVal CHART_NAME_STR As Variant = "", _
Optional ByRef XAXIS_RNG As Excel.Range, _
Optional ByRef YAXIS_RNG As Excel.Range, _
Optional ByVal COLOR_INDEX As Long = 3, _
Optional ByVal BORDER_WEIGHT_VAL As Variant = xlThin, _
Optional ByRef SRC_WSHEET As Excel.Worksheet)

Dim i As Long
Dim j As Variant
Dim CHART_OBJ As Excel.ChartObject

On Error GoTo ERROR_LABEL

EXCEL_CHART_ADD_SERIE_FUNC = False

If SRC_WSHEET Is Nothing Then: Set SRC_WSHEET = ActiveSheet
i = SRC_WSHEET.ChartObjects.COUNT
If i = 0 Then: GoTo ERROR_LABEL
If CHART_NAME_STR = "" Then: CHART_NAME_STR = SRC_WSHEET.ChartObjects(i).name

If EXCEL_CHART_LOOK_FUNC(CHART_NAME_STR, 0, SRC_WSHEET) = True Then
    Set CHART_OBJ = SRC_WSHEET.ChartObjects(CHART_NAME_STR)
    If EXCEL_CHART_LOOK_SERIES_FUNC(SERIE_NAME_STR, CHART_NAME_STR, 0, SRC_WSHEET) = True Then
        CHART_OBJ.Chart.SeriesCollection(SERIE_NAME_STR).Delete
    End If
    CHART_OBJ.Chart.SeriesCollection.NewSeries

    j = EXCEL_CHART_COUNT_SERIES_FUNC(CHART_NAME_STR, SRC_WSHEET)
    'Locate the POSITION of the new Serie
    With CHART_OBJ.Chart
        .SeriesCollection(j).name = SERIE_NAME_STR
        .SeriesCollection(j).Values = XAXIS_RNG
        .SeriesCollection(j).Values = YAXIS_RNG
        .SeriesCollection(j).Border.ColorIndex = COLOR_INDEX
        .SeriesCollection(j).Border.WEIGHT = BORDER_WEIGHT_VAL
    End With
Else
    GoTo ERROR_LABEL
End If

EXCEL_CHART_ADD_SERIE_FUNC = True

Exit Function
ERROR_LABEL:
EXCEL_CHART_ADD_SERIE_FUNC = False
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : EXCEL_CHART_ADD_SERIES_FUNC
'DESCRIPTION   :
'LIBRARY       : EXCEL_CHART
'GROUP         : SERIES
'ID            : 005
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'************************************************************************************
'************************************************************************************

Function EXCEL_CHART_ADD_SERIES_FUNC(ByVal CHART_NAME_STR As Variant, _
ByRef NAMES_VECTOR() As Variant, _
ByRef XAXIS_RNG() As Excel.Range, _
ByRef YAXIS_RNG() As Excel.Range, _
Optional ByVal BORDERS_ARR As Variant, _
Optional ByRef COLORS_ARR As Variant, _
Optional ByRef SRC_WSHEET As Excel.Worksheet)

Dim i As Long
Dim NSERIES As Long
Dim CHART_OBJ As Excel.ChartObject

On Error GoTo ERROR_LABEL

EXCEL_CHART_ADD_SERIES_FUNC = False
If SRC_WSHEET Is Nothing Then: Set SRC_WSHEET = ActiveSheet

If EXCEL_CHART_LOOK_FUNC(CHART_NAME_STR, 0, SRC_WSHEET) = True Then
    Set CHART_OBJ = SRC_WSHEET.ChartObjects(CHART_NAME_STR)
    NSERIES = UBound(NAMES_VECTOR) - LBound(NAMES_VECTOR) + 1
    With CHART_OBJ.Chart
        If NSERIES >= .SeriesCollection.COUNT Then
            Do Until .SeriesCollection.COUNT >= NSERIES
               .SeriesCollection.NewSeries
            Loop
        Else
            Do Until .SeriesCollection.COUNT = NSERIES
               .SeriesCollection(.SeriesCollection.COUNT).Delete
            Loop
        End If
        For i = 1 To NSERIES
            .SeriesCollection(i).name = NAMES_VECTOR(i)
            .SeriesCollection(i).XValues = XAXIS_RNG(i)
            .SeriesCollection(i).Values = YAXIS_RNG(i)
        Next i
        If IsArray(BORDERS_ARR) Then
            For i = 1 To NSERIES
                If BORDERS_ARR(i) <> "" Then: .SeriesCollection(i).Border.WEIGHT = BORDERS_ARR(i) 'xlThin
            Next i
        End If
        If IsArray(COLORS_ARR) Then
            For i = 1 To NSERIES
                If COLORS_ARR(i) <> "" Then: .SeriesCollection(i).Border.ColorIndex = COLORS_ARR(i)
            Next i
        End If
    End With
Else
    GoTo ERROR_LABEL
End If

EXCEL_CHART_ADD_SERIES_FUNC = True

Exit Function
ERROR_LABEL:
EXCEL_CHART_ADD_SERIES_FUNC = False
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : EXCEL_CHART_REMOVE_SERIE_FUNC
'DESCRIPTION   :
'LIBRARY       : EXCEL_CHART
'GROUP         : SERIES
'ID            : 006
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'************************************************************************************
'************************************************************************************

Function EXCEL_CHART_REMOVE_SERIE_FUNC(ByVal CHART_NAME_STR As Variant, _
ByRef DATA_VECTOR As Variant, _
Optional ByRef SRC_WSHEET As Excel.Worksheet)

Dim i As Long
Dim SERIE_OBJ As Excel.Series
Dim MATCH_FLAG As Boolean
        
On Error GoTo ERROR_LABEL

EXCEL_CHART_REMOVE_SERIE_FUNC = False
If SRC_WSHEET Is Nothing Then: Set SRC_WSHEET = ActiveSheet

If EXCEL_CHART_LOOK_FUNC(CHART_NAME_STR, 0, SRC_WSHEET) = True Then
    For Each SERIE_OBJ In SRC_WSHEET.ChartObjects(CHART_NAME_STR).Chart.SeriesCollection
          MATCH_FLAG = False
          For i = LBound(DATA_VECTOR) To UBound(DATA_VECTOR)
               If DATA_VECTOR(i) = SERIE_OBJ.name Then
                   MATCH_FLAG = True
                   Exit For
               End If
          Next i
          If MATCH_FLAG = False Then: SERIE_OBJ.Delete
    Next SERIE_OBJ
Else
    GoTo ERROR_LABEL
End If

EXCEL_CHART_REMOVE_SERIE_FUNC = True

Exit Function
ERROR_LABEL:
EXCEL_CHART_REMOVE_SERIE_FUNC = False
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : EXCEL_CHART_REMOVE_SERIES_FUNC
'DESCRIPTION   :
'LIBRARY       : EXCEL_CHART
'GROUP         : SERIES
'ID            : 007
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'************************************************************************************
'************************************************************************************

Function EXCEL_CHART_REMOVE_SERIES_FUNC(ByRef CHART_OBJ As Excel.ChartObject)

Dim SERIES_OBJ As Excel.Series

On Error GoTo ERROR_LABEL

EXCEL_CHART_REMOVE_SERIES_FUNC = False

For Each SERIES_OBJ In CHART_OBJ.Chart.SeriesCollection: SERIES_OBJ.Delete: Next SERIES_OBJ
Set SERIES_OBJ = Nothing 'REMOVE ALL UNWANTED SERIES FROM CHART, IF ANY

EXCEL_CHART_REMOVE_SERIES_FUNC = True

Exit Function
ERROR_LABEL:
EXCEL_CHART_REMOVE_SERIES_FUNC = False
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : EXCEL_CHART_ADD_SERIES_FRAME1_FUNC
'DESCRIPTION   :
'LIBRARY       : EXCEL_CHART
'GROUP         : SERIES
'ID            : 008
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'************************************************************************************
'************************************************************************************

Function EXCEL_CHART_ADD_SERIES_FRAME1_FUNC(ByRef DATA_RNG As Excel.Range, _
ByVal CHART_NAME_STR As Variant, _
Optional ByRef SRC_WSHEET As Excel.Worksheet)

Dim i As Long
Dim j As Long

Dim NCOLUMNS As Long
Dim XAXIS_RNG() As Excel.Range
Dim YAXIS_RNG() As Excel.Range
Dim NAMES_ARR() As Variant

Dim MATCH_FLAG As Boolean

On Error GoTo ERROR_LABEL

EXCEL_CHART_ADD_SERIES_FRAME1_FUNC = False
NCOLUMNS = DATA_RNG.Columns.COUNT
If NCOLUMNS Mod NCOLUMNS / 2 <> 0 Then: GoTo ERROR_LABEL

If SRC_WSHEET Is Nothing Then: Set SRC_WSHEET = ActiveSheet

i = 0
For j = 1 To NCOLUMNS Step 2
    i = i + 1
    ReDim Preserve NAMES_ARR(i)
    ReDim Preserve YAXIS_RNG(i)
    ReDim Preserve XAXIS_RNG(i)
    NAMES_ARR(i) = DATA_RNG.Columns(j).Cells(1).value
    Set XAXIS_RNG(i) = RNG_RESIZE_RNG_FUNC(DATA_RNG.Columns(j).Offset(1, 0), 1)
    Set YAXIS_RNG(i) = RNG_RESIZE_RNG_FUNC(DATA_RNG.Columns(j + 1).Offset(1, 0), 1)
Next j

MATCH_FLAG = EXCEL_CHART_ADD_SERIES_FUNC(CHART_NAME_STR, NAMES_ARR(), XAXIS_RNG(), YAXIS_RNG(), , , SRC_WSHEET)
If MATCH_FLAG = False Then: GoTo ERROR_LABEL
EXCEL_CHART_ADD_SERIES_FRAME1_FUNC = True

Exit Function
ERROR_LABEL:
EXCEL_CHART_ADD_SERIES_FRAME1_FUNC = False
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : EXCEL_CHART_ADD_SERIES_FRAME2_FUNC
'DESCRIPTION   :
'LIBRARY       : EXCEL_CHART
'GROUP         : SERIES
'ID            : 009
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'************************************************************************************
'************************************************************************************

Function EXCEL_CHART_ADD_SERIES_FRAME2_FUNC(ByVal CHART_OBJ As Excel.Chart, _
ByRef DATA_GROUP As Variant)

Dim i As Long
Dim j As Long
Dim k As Long
Dim SERIES_OBJ As Excel.Series

On Error GoTo ERROR_LABEL

EXCEL_CHART_ADD_SERIES_FRAME2_FUNC = False
If IsArray(DATA_GROUP) = True Then
    k = UBound(DATA_GROUP) - LBound(DATA_GROUP) + 1
    For i = LBound(DATA_GROUP) To UBound(DATA_GROUP)
        j = LBound(DATA_GROUP(i))
        Set SERIES_OBJ = CHART_OBJ.SeriesCollection.NewSeries
        With SERIES_OBJ
            If k <= 2 Then
                .name = DATA_GROUP(i)(j + 0) '"Chart Series 1"
                .Values = DATA_GROUP(i)(j + 1) 'Array(10, 20, 30, 20, 10)
                .XValues = DATA_GROUP(i)(j + 2) 'Array("alpha", "beta", "gamma", "delta", "epsilon")
            Else 'No Labels
                .Values = DATA_GROUP(i)(j + 0)
                .XValues = DATA_GROUP(i)(j + 1)
            End If
        End With
        Set SERIES_OBJ = Nothing
    Next i
End If
EXCEL_CHART_ADD_SERIES_FRAME2_FUNC = True
Exit Function
ERROR_LABEL:
EXCEL_CHART_ADD_SERIES_FRAME2_FUNC = False
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : EXCEL_CHART_ADD_SERIE_FRAME1_FUNC
'DESCRIPTION   : LOAD CHART SHEET WITH DATA
'LIBRARY       : EXCEL_CHART
'GROUP         : LOAD
'ID            : 009
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'************************************************************************************
'************************************************************************************

Function EXCEL_CHART_ADD_SERIE_FRAME1_FUNC(ByRef CHART_OBJ As Excel.Chart, _
ByRef XAXIS_RNG As Excel.Range, _
ByRef YAXIS_RNG As Excel.Range, _
Optional ByVal SERIE_NAME_STR As Variant = 1)

On Error GoTo ERROR_LABEL
 
EXCEL_CHART_ADD_SERIE_FRAME1_FUNC = False
If CHART_OBJ.SeriesCollection.COUNT <> 0 Then
    With CHART_OBJ.SeriesCollection(SERIE_NAME_STR)
          .XValues = XAXIS_RNG
          .Values = YAXIS_RNG
    End With
Else
    CHART_OBJ.SeriesCollection.NewSeries
    With CHART_OBJ.SeriesCollection(1)
          .XValues = XAXIS_RNG
          .Values = YAXIS_RNG
    End With
End If

EXCEL_CHART_ADD_SERIE_FRAME1_FUNC = True

Exit Function
ERROR_LABEL:
EXCEL_CHART_ADD_SERIE_FRAME1_FUNC = False
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : EXCEL_CHART_ADD_SERIE_FRAME2_FUNC
'DESCRIPTION   :
'LIBRARY       : EXCEL_CHART
'GROUP         : LOAD
'ID            : 010
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'************************************************************************************
'************************************************************************************

'Sheets("Chart1").Select
'Call EXCEL_CHART_ADD_SERIE_FRAME2_FUNC(ActiveChart, Range("xr"), Range("yr"), 1, 0)

Function EXCEL_CHART_ADD_SERIE_FRAME2_FUNC(ByRef CHART_OBJ As Excel.Chart, _
ByRef DATA_RNG As Excel.Range, _
Optional ByRef DATES_RNG As Excel.Range, _
Optional ByVal SERIE_NAME_STR As Variant = 1, _
Optional ByVal VERSION As Integer = 1)
    
On Error GoTo ERROR_LABEL
    
EXCEL_CHART_ADD_SERIE_FRAME2_FUNC = False
Select Case VERSION
Case 0
    With CHART_OBJ
        .SetSourceData DATA_RNG
        .SeriesCollection(SERIE_NAME_STR).XValues = DATES_RNG
    End With
Case Else
    With CHART_OBJ
        .SetSourceData DATA_RNG
    End With
End Select

EXCEL_CHART_ADD_SERIE_FRAME2_FUNC = True

Exit Function
ERROR_LABEL:
EXCEL_CHART_ADD_SERIE_FRAME2_FUNC = False
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : EXCEL_CHART_FORMAT_SERIES_FUNC
'DESCRIPTION   :
'LIBRARY       : EXCEL_CHART
'GROUP         : SERIES
'ID            : 014
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'************************************************************************************
'************************************************************************************

Function EXCEL_CHART_FORMAT_SERIES_FUNC(ByRef SERIES_OBJ As Excel.Series) 'MyNewSrs
On Error GoTo ERROR_LABEL
EXCEL_CHART_FORMAT_SERIES_FUNC = False
With SERIES_OBJ
    With .Border
        .ColorIndex = 1
        .WEIGHT = xlHairline
        .LineStyle = xlNone
    End With
    .MarkerBackgroundColorIndex = 2 'xlAutomatic
    .MarkerForegroundColorIndex = xlAutomatic
    .MarkerStyle = xlCircle 'xlNone
    .Smooth = False
    .MarkerSize = 5
    .Shadow = False
End With
EXCEL_CHART_FORMAT_SERIES_FUNC = False
Exit Function
ERROR_LABEL:
EXCEL_CHART_FORMAT_SERIES_FUNC = False
End Function
