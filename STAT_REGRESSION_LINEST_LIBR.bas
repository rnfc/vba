Attribute VB_Name = "STAT_REGRESSION_LINEST_LIBR"
Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.

'************************************************************************************
'************************************************************************************
'FUNCTION      : EXCEL_LINEST1_FUNC

'DESCRIPTION   : Calculates the statistics for a line by using the "least squares"
'method to calculate a straight line that best fits your data, and then returns an
'array that describes the line

'LIBRARY       : STATISTICS
'GROUP         : REGRESSION_LINEST
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/22/2009
'************************************************************************************
'************************************************************************************

Function EXCEL_LINEST1_FUNC(ByRef XDATA_RNG As Variant, _
ByRef YDATA_RNG As Variant, _
Optional ByVal INTERCEPT_FLAG As Boolean = False)

Dim i As Long
Dim j As Long
Dim k As Long  ' count the number of valid observations

Dim XNROWS As Long
Dim YNROWS As Long

Dim XNCOLUMNS As Long
Dim YNCOLUMNS As Long

Dim ERROR_STR As String

Dim YDATA_VECTOR As Variant
Dim XDATA_MATRIX As Variant

Dim TEMP1_MATRIX As Variant
Dim TEMP2_MATRIX As Variant

Dim YTEMP_VECTOR() As Double
Dim XTEMP_MATRIX() As Double 'The actual values go here

On Error GoTo ERROR_LABEL

XDATA_MATRIX = XDATA_RNG
If UBound(XDATA_MATRIX, 1) = 1 Then
    XDATA_MATRIX = MATRIX_TRANSPOSE_FUNC(XDATA_MATRIX)
End If

YDATA_VECTOR = YDATA_RNG
If UBound(YDATA_VECTOR, 1) = 1 Then
    YDATA_VECTOR = MATRIX_TRANSPOSE_FUNC(YDATA_VECTOR)
End If

ERROR_STR = ""
'---------------------------------------------------------------------------------
XNCOLUMNS = UBound(XDATA_MATRIX, 2)
'---------------------------------------------------------------------------------
If XNCOLUMNS > 51 Then
    ERROR_STR = "Unfortunately, this function cannot handle more than 51 independent variables.  You've selected " & XNCOLUMNS & ". Sorry!"
        GoTo ERROR_LABEL
End If
'---------------------------------------------------------------------------------
' go through cells in each area (determine how many there will be)
' labels need to be found
XNROWS = UBound(XDATA_MATRIX, 1) - 1
YNROWS = UBound(YDATA_VECTOR, 1) - 1
YNCOLUMNS = UBound(YDATA_VECTOR, 2)

If XNROWS <> YNROWS Then
    ERROR_STR = "You must select the same number of rows for both the X variable(s) and the Y variable. Please try again."
    GoTo ERROR_LABEL
End If
'---------------------------------------------------------------------------------
' Check that we have just one Y column
If YNCOLUMNS > 1 Then
    ERROR_STR = "You must select only one column for the Y variable. Please try again."
    GoTo ERROR_LABEL
End If

'---------------------------------------------------------------------------------
' Check on labels
ReDim XTEMP_MATRIX(1 To XNROWS, 1 To XNCOLUMNS)
ReDim YTEMP_VECTOR(1 To XNROWS, 1 To 1)
'---------------------------------------------------------------------------------
For j = 1 To XNCOLUMNS
    If IsNumeric(XDATA_MATRIX(1, j)) = True Then
        ERROR_STR = "The X variable label in column " & j & " you've chosen is a number. Please try again."
        'Potential Label Problem
        GoTo ERROR_LABEL
    End If
Next j

'---------------------------------------------------------------------------------
If IsNumeric(YDATA_VECTOR(1, 1)) = True Then
    ERROR_STR = "The Y variable label you've chosen is a number. Please try again." 'Potential Label Problem
    GoTo ERROR_LABEL
End If
'---------------------------------------------------------------------------------
' Start reading the data
' Must read in one SROW at a time across Y variable and X variables
' Data is assumed to be in columnar format!
'---------------------------------------------------------------------------------
k = 0
For i = 1 To XNROWS
    On Error GoTo 1982
    ' Read y data first
    k = k + 1
     'remember first SROW is label so must add one
    ' We are sent to error handling if this isn't a number
    ' Now check for blanks
    YTEMP_VECTOR(k, 1) = YDATA_VECTOR(i + 1, 1)
    If IsEmpty(YDATA_VECTOR(i + 1, 1)) = True Then
        k = k - 1 ' we are going to skip this obs.
        GoTo 1983
    End If
    ' If we've passed, go to the x variables
    For j = 1 To XNCOLUMNS
        On Error GoTo 1982
     '   Check for empty values
        If IsEmpty(XDATA_MATRIX(i + 1, j)) = True Then
            k = k - 1
            GoTo 1983
        Else
            XTEMP_MATRIX(k, j) = XDATA_MATRIX(i + 1, j)
        End If
    Next j
    GoTo 1983
1982:
    k = k - 1
'Resume 1983
1983:
Next i

'---------------------------------------------------------------------------------
' End reading in data
If k < XNCOLUMNS Then
    ERROR_STR = "There aren't enough observations with non-missing values to obtain parameter estimates. Try again."
    GoTo ERROR_LABEL
End If
'---------------------------------------------------------------------------------
'-------------------------------------------------------------------------
TEMP1_MATRIX = WorksheetFunction.LinEst(YTEMP_VECTOR(), XTEMP_MATRIX(), INTERCEPT_FLAG, True)
'-------------------------------------------------------------------------
ReDim TEMP2_MATRIX(0 To UBound(TEMP1_MATRIX, 1), _
0 To IIf(UBound(TEMP1_MATRIX, 2) < 5, 5, UBound(TEMP1_MATRIX, 2)))
'-------------------------------------------------------------------------
TEMP2_MATRIX(0, 0) = "Variables"
TEMP2_MATRIX(1, 0) = "Coefficients"
TEMP2_MATRIX(2, 0) = "Standard Error"
TEMP2_MATRIX(3, 0) = "Coefficient of Determination"
TEMP2_MATRIX(4, 0) = "F-Statistic"
TEMP2_MATRIX(5, 0) = "Regression Sum of Squares"
'-------------------------------------------------------------------------
For j = 1 To UBound(TEMP1_MATRIX, 2)
    For i = 1 To UBound(TEMP1_MATRIX, 1)
        TEMP2_MATRIX(i, j) = TEMP1_MATRIX(i, j)
    Next i
Next j
'-------------------------------------------------------------------------
TEMP2_MATRIX(3, 3) = TEMP2_MATRIX(3, 2)
TEMP2_MATRIX(4, 3) = TEMP2_MATRIX(4, 2)
TEMP2_MATRIX(5, 3) = TEMP2_MATRIX(5, 2)

TEMP2_MATRIX(3, 2) = "Standard Error for the Y Estimate"
TEMP2_MATRIX(4, 2) = "Degrees of Freedom"
TEMP2_MATRIX(5, 2) = "Residual Sum of Squares"
'-------------------------------------------------------------------------
If INTERCEPT_FLAG = True Then
'-------------------------------------------------------------------------
    TEMP2_MATRIX(0, UBound(TEMP1_MATRIX, 2)) = "Y0: " & YDATA_VECTOR(1, 1)
    For i = 1 To XNCOLUMNS
        TEMP2_MATRIX(0, UBound(TEMP1_MATRIX, 2) + 1 - i - 1) = "X" & i & ": " & XDATA_MATRIX(1, i)
    Next i
'-------------------------------------------------------------------------
Else
'-------------------------------------------------------------------------
    For i = 1 To XNCOLUMNS
        TEMP2_MATRIX(0, UBound(TEMP1_MATRIX, 2) - i) = "X" & i & ": " & XDATA_MATRIX(1, i)
    Next i
'-------------------------------------------------------------------------
End If
'-------------------------------------------------------------------------

TEMP2_MATRIX(3, 4) = "No. Var"
TEMP2_MATRIX(3, 5) = XNCOLUMNS

TEMP2_MATRIX(4, 4) = "No. Obs."
TEMP2_MATRIX(4, 5) = k

TEMP2_MATRIX(5, 4) = "No. Missing Obs."
TEMP2_MATRIX(5, 5) = XNROWS - k

EXCEL_LINEST1_FUNC = TEMP2_MATRIX

Exit Function
ERROR_LABEL:
If ERROR_STR = "" Then
    EXCEL_LINEST1_FUNC = Err.number
Else
    EXCEL_LINEST1_FUNC = ERROR_STR
End If
End Function



'************************************************************************************
'************************************************************************************
'FUNCTION      : EXCEL_LINEST2_FUNC

'DESCRIPTION   : Calculates the statistics for a line by using the "least squares"
'method to calculate a straight line that best fits your data, and then returns an
'array that describes the line

'LIBRARY       : STATISTICS
'GROUP         : REGRESSION_LINEST
'ID            : 002
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/22/2009
'************************************************************************************
'************************************************************************************

Function EXCEL_LINEST2_FUNC(ByRef XDATA_RNG As Excel.Range, _
ByRef YDATA_RNG As Excel.Range, _
Optional ByVal INTERCEPT_FLAG As Boolean = True)

Dim h As Long  'searching through cells in an area
Dim i As Long
Dim j As Long
Dim k As Long
Dim l As Long ' index for x variables

Dim hh As Long
Dim ii As Long 'keep track of which X areas have problems
Dim jj As Long
Dim ll As Long 'count the number of valid observations

Dim NSIZE As Long
Dim XNROWS As Long
Dim YNROWS As Long

Dim XNCOLUMNS As Long
Dim YNCOLUMNS As Long

Dim ROWS_ARR() As Long    ' keep track of how many obs in each area
Dim COLUMNS_ARR() As Long ' keep track of how many x variables in each area

Dim ERROR_STR As String
Dim YVAR_LABEL_STR As String
Dim XVAR_LABEL_ARR() As String

Dim TEMP1_MATRIX As Variant ' The matrix will contain output
Dim TEMP2_MATRIX As Variant

Dim XTEMP_MATRIX() As Double 'The actual values go here
Dim YTEMP_VECTOR() As Double

Dim ERROR_MATCH_FLAG As Boolean
'indicator for mismatch between no of obs in X cols

On Error GoTo ERROR_LABEL

ERROR_STR = ""
NSIZE = XDATA_RNG.Areas.COUNT
XNCOLUMNS = 0
ERROR_MATCH_FLAG = False
ReDim COLUMNS_ARR(1 To NSIZE)
ReDim ROWS_ARR(1 To NSIZE)
'-------------------------------------------------------------------------
For i = 1 To NSIZE
    COLUMNS_ARR(i) = XDATA_RNG.Areas(i).Columns.COUNT
    ROWS_ARR(i) = XDATA_RNG.Areas(i).Rows.COUNT
    XNCOLUMNS = XNCOLUMNS + COLUMNS_ARR(i)
    If i > 1 Then
        If ROWS_ARR(i) <> ROWS_ARR(i - 1) Then
            ERROR_MATCH_FLAG = True
            ii = i - 1
            jj = i
        End If
    End If
Next i
'-------------------------------------------------------------------------
If XNCOLUMNS > 51 Then
    ERROR_STR = "Unfortunately, this function cannot handle more than 51 independent variables.  You've selected " & XNCOLUMNS & ". Sorry!"
        GoTo ERROR_LABEL
End If
'-------------------------------------------------------------------------
' Warning if ERROR_MATCH_FLAG is true
If ERROR_MATCH_FLAG = True Then
    ERROR_STR = "The number of rows in X area " & ii & " does not equal the number of observations in X area " & jj & ". Please try again."
    GoTo ERROR_LABEL
End If
'-------------------------------------------------------------------------
' go through cells in each area (determine how many there will be)
' labels need to be found
XNROWS = ROWS_ARR(1) - 1
ReDim XVAR_LABEL_ARR(1 To XNCOLUMNS) As String
YNROWS = YDATA_RNG.Rows.COUNT - 1
YNCOLUMNS = YDATA_RNG.Columns.COUNT
If XNROWS <> YNROWS Then
    ERROR_STR = "You must select the same number of rows for both the X variable(s) and the Y variable. Please try again."
    GoTo ERROR_LABEL
End If
'-------------------------------------------------------------------------
' Check that we have just one Y column
If YNCOLUMNS > 1 Then
    ERROR_STR = "You must select only one column for the Y variable. Please try again."
    GoTo ERROR_LABEL
End If
'-------------------------------------------------------------------------
' Check on labels
ReDim XTEMP_MATRIX(1 To XNROWS, 1 To XNCOLUMNS)
ReDim YTEMP_VECTOR(1 To XNROWS, 1 To 1)
l = 0
For i = 1 To NSIZE
    For j = 1 To COLUMNS_ARR(i)
        l = l + 1
        h = j
        XVAR_LABEL_ARR(l) = XDATA_RNG.Areas(i).Cells(h)
        If IsNumeric(XVAR_LABEL_ARR(l)) = True Then
            hh = MsgBox("The X variable label in column " & l & " you've chosen is a number.  Do you really want the variable label to be " & XVAR_LABEL_ARR(l) & "?", vbYesNo, Title:="Potential Label Problem")
            If hh = vbNo Then GoTo ERROR_LABEL
        End If
    Next j
Next i
'-------------------------------------------------------------------------
YVAR_LABEL_STR = YDATA_RNG(1)
If IsNumeric(YVAR_LABEL_STR) = True Then
    hh = _
    MsgBox("The Y variable label you've chosen is a number. Do you really want the variable label to be " & YVAR_LABEL_STR & "?", vbYesNo, Title:="Potential Label Problem")
    If hh = vbNo Then GoTo ERROR_LABEL
End If
'-------------------------------------------------------------------------
' Start reading the data
' Must read in one SROW at a time
' across Y variable and X variables
' Data is assumed to be in columnar format!
ll = 0
For i = 1 To XNROWS
    On Error GoTo 1982
    ' Read y data first
    ll = ll + 1
     'remember first SROW is label so must add one
    ' We are sent to error handling if this isn't a number
    ' Now check for blanks
    YTEMP_VECTOR(ll, 1) = YDATA_RNG(i + 1, 1)
    If IsEmpty(YDATA_RNG(i + 1, 1)) = True Then
        ll = ll - 1 ' we are going to skip this obs.
        GoTo 1983
    End If
    ' If we've passed, go to the x variables
    l = 0
    For j = 1 To NSIZE
        For k = 1 To COLUMNS_ARR(j)
            h = i * COLUMNS_ARR(j) + k
            l = l + 1
            On Error GoTo 1982
     '   XTEMP_MATRIX(ll, j) = XDATA_RNG(i + 1, j)
     '   Check for empty values
            If IsEmpty(XDATA_RNG.Areas(j).Cells(h)) = True Then
                ll = ll - 1
                GoTo 1983
            Else
                XTEMP_MATRIX(ll, l) = XDATA_RNG.Areas(j).Cells(h)
            End If
        Next k
    Next j
    GoTo 1983
1982:
    ll = ll - 1
    'Resume 1983
1983:
Next i
'---------------------------------------------------------------------
' End reading in data
If ll < XNCOLUMNS Then
    ERROR_STR = "There aren't enough observations with non-missing values to obtain parameter estimates.  Try again."
    GoTo ERROR_LABEL
End If
'-------------------------------------------------------------------------
TEMP1_MATRIX = WorksheetFunction.LinEst(YTEMP_VECTOR(), XTEMP_MATRIX(), INTERCEPT_FLAG, True)
If IsArray(TEMP1_MATRIX) = False Then: GoTo ERROR_LABEL

'-------------------------------------------------------------------------
ReDim TEMP2_MATRIX(0 To UBound(TEMP1_MATRIX, 1), 0 To IIf(UBound(TEMP1_MATRIX, 2) < 5, 5, UBound(TEMP1_MATRIX, 2)))
'-------------------------------------------------------------------------
TEMP2_MATRIX(0, 0) = "Variables"
TEMP2_MATRIX(1, 0) = "Coefficients"
TEMP2_MATRIX(2, 0) = "Standard Error"
TEMP2_MATRIX(3, 0) = "Coefficient of Determination"
TEMP2_MATRIX(4, 0) = "F-Statistic"
TEMP2_MATRIX(5, 0) = "Regression Sum of Squares"
'-------------------------------------------------------------------------
For j = 1 To UBound(TEMP1_MATRIX, 2)
    For i = 1 To UBound(TEMP1_MATRIX, 1)
        TEMP2_MATRIX(i, j) = TEMP1_MATRIX(i, j)
    Next i
Next j
'-------------------------------------------------------------------------
TEMP2_MATRIX(3, 3) = TEMP2_MATRIX(3, 2)
TEMP2_MATRIX(4, 3) = TEMP2_MATRIX(4, 2)
TEMP2_MATRIX(5, 3) = TEMP2_MATRIX(5, 2)

TEMP2_MATRIX(3, 2) = "Standard Error for the Y Estimate"
TEMP2_MATRIX(4, 2) = "Degrees of Freedom"
TEMP2_MATRIX(5, 2) = "Residual Sum of Squares"
'-------------------------------------------------------------------------
If INTERCEPT_FLAG = True Then
'-------------------------------------------------------------------------
    TEMP2_MATRIX(0, UBound(TEMP1_MATRIX, 2)) = "Y0: " & YVAR_LABEL_STR
    For i = 1 To XNCOLUMNS
        TEMP2_MATRIX(0, UBound(TEMP1_MATRIX, 2) + 1 - i - 1) = "X" & i & ": " & XVAR_LABEL_ARR(i)
    Next i
'-------------------------------------------------------------------------
Else
'-------------------------------------------------------------------------
    For i = 1 To XNCOLUMNS
        TEMP2_MATRIX(0, UBound(TEMP1_MATRIX, 2) - i) = "X" & i & ": " & XVAR_LABEL_ARR(i)
    Next i
'-------------------------------------------------------------------------
End If
'-------------------------------------------------------------------------

TEMP2_MATRIX(3, 4) = "No. Var"
TEMP2_MATRIX(3, 5) = XNCOLUMNS

TEMP2_MATRIX(4, 4) = "No. Obs."
TEMP2_MATRIX(4, 5) = ll

TEMP2_MATRIX(5, 4) = "No. Missing Obs."
TEMP2_MATRIX(5, 5) = XNROWS - ll

EXCEL_LINEST2_FUNC = TEMP2_MATRIX

Exit Function
ERROR_LABEL:
If ERROR_STR = "" Then
    EXCEL_LINEST2_FUNC = Err.number
Else
    EXCEL_LINEST2_FUNC = ERROR_STR
End If
End Function

Function EXCEL_LINEST3_FUNC(ByRef XDATA_RNG As Variant, _
ByRef YDATA_RNG As Variant, _
Optional ByVal INTERCEPT_FLAG As Boolean = True, _
Optional ByVal LAGS_RNG As Variant = 4)

Dim h As Long
Dim i As Long
Dim j As Long
Dim k As Long
Dim l As Long

Dim NROWS As Long
Dim NO_OBS As Long
Dim NCOLUMNS As Long

Dim DF_VAL As Long 'df
Dim NO_LAGS As Long
Dim LABEL_STR As String

Dim TEMP_VAL As Variant
Dim DATA_ARR As Variant
Dim DATA_MATRIX As Variant
Dim LAGS_VECTOR As Variant
Dim XTEMP_MATRIX As Variant

Dim YDATA_VECTOR As Variant
Dim XDATA_MATRIX As Variant

On Error GoTo ERROR_LABEL

YDATA_VECTOR = YDATA_RNG
If UBound(YDATA_VECTOR, 1) = 1 Then
    YDATA_VECTOR = MATRIX_TRANSPOSE_FUNC(YDATA_VECTOR)
End If
NROWS = UBound(YDATA_VECTOR, 1)

XDATA_MATRIX = XDATA_RNG
NO_OBS = UBound(XDATA_MATRIX, 1)
NCOLUMNS = UBound(XDATA_MATRIX, 2)

'----------------------------------------------------------------------------------------------------
If IsArray(LAGS_RNG) = True Then
'----------------------------------------------------------------------------------------------------
    LAGS_VECTOR = LAGS_RNG
    If UBound(LAGS_VECTOR, 2) <> NCOLUMNS Then: GoTo ERROR_LABEL
    NO_LAGS = UBound(LAGS_VECTOR, 1)
    For i = 1 To NO_LAGS
        For j = 1 To NCOLUMNS
            k = LAGS_VECTOR(i, j)
            If NO_OBS - NROWS + 1 <= k Then
                LAGS_VECTOR(i, j) = NO_OBS - NROWS
            Else
                LAGS_VECTOR(i, j) = k
            End If
        Next j
    Next i
'----------------------------------------------------------------------------------------------------
Else
'----------------------------------------------------------------------------------------------------
    NO_LAGS = LAGS_RNG + 1
    ReDim LAGS_VECTOR(1 To NO_LAGS, 1 To NCOLUMNS)
    For i = 0 To LAGS_RNG
        For j = 1 To NCOLUMNS
            If NO_OBS - NROWS + 1 <= i Then
                LAGS_VECTOR(i + 1, j) = NO_OBS - NROWS
            Else
                LAGS_VECTOR(i + 1, j) = i
            End If
        Next j
    Next i
'----------------------------------------------------------------------------------------------------
End If
'----------------------------------------------------------------------------------------------------
If INTERCEPT_FLAG = True Then h = 1 Else h = 0
DF_VAL = NROWS - NCOLUMNS - h
ReDim DATA_MATRIX(0 To NO_LAGS, 1 To NCOLUMNS + (NCOLUMNS + h) * 4 + 3)
j = 1
'----------------------------------------------------------------------------------------------------
For k = 1 To 5
'----------------------------------------------------------------------------------------------------
    If k = 1 Then
        LABEL_STR = "Lag"
    ElseIf k = 2 Then
        LABEL_STR = "Coefficient"
    ElseIf k = 3 Then
        LABEL_STR = "Standard Error"
    ElseIf k = 4 Then
        LABEL_STR = "t Stat"
    Else
        LABEL_STR = "P-value"
    End If
    For l = 1 To NCOLUMNS
        DATA_MATRIX(0, j) = LABEL_STR & ": " & "X" & l
        j = j + 1
    Next l
    If k > 1 And h = 1 Then
        DATA_MATRIX(0, j) = LABEL_STR & ": " & "Y"
        j = j + 1
    End If
'----------------------------------------------------------------------------------------------------
Next k
'----------------------------------------------------------------------------------------------------
DATA_MATRIX(0, j) = "Rsq"
j = j + 1
DATA_MATRIX(0, j) = "F"
j = j + 1
DATA_MATRIX(0, j) = "Significance F"
'----------------------------------------------------------------------------------------------------
For i = 1 To NO_LAGS
'----------------------------------------------------------------------------------------------------
    ReDim XTEMP_MATRIX(1 To NROWS, 1 To NCOLUMNS)
    For j = 1 To NCOLUMNS
        DATA_MATRIX(i, j) = LAGS_VECTOR(i, j)
        l = NO_OBS - LAGS_VECTOR(i, j)
        For k = NROWS To 1 Step -1
            XTEMP_MATRIX(k, j) = XDATA_MATRIX(l, j)
            l = l - 1
        Next k
    Next j
    
    DATA_ARR = WorksheetFunction.LinEst(YDATA_VECTOR, XTEMP_MATRIX, INTERCEPT_FLAG, True)
    If IsArray(DATA_ARR) = False Then: GoTo 1983
    DF_VAL = DATA_ARR(4, 2) 'NROWS - NCOLUMNS - h '#DF Perfect
    
    l = NCOLUMNS
    k = (NCOLUMNS + h)
    For j = 1 To k
        DATA_MATRIX(i, l + j) = DATA_ARR(1, j)
        DATA_MATRIX(i, l + j + k * 1) = DATA_ARR(2, j)
        
        TEMP_VAL = DATA_ARR(1, j) / DATA_ARR(2, j)
        DATA_MATRIX(i, l + j + k * 2) = TEMP_VAL
        
        TEMP_VAL = WorksheetFunction.tdist(Abs(TEMP_VAL), DF_VAL, 2)
        DATA_MATRIX(i, l + j + k * 3) = TEMP_VAL
    Next j
    
    j = 1 + l + k * 4
    TEMP_VAL = DATA_ARR(3, 1) 'Rsq
    DATA_MATRIX(i, j) = TEMP_VAL
    j = j + 1
    
    TEMP_VAL = DATA_ARR(4, 1) 'F
    DATA_MATRIX(i, j) = TEMP_VAL
    j = j + 1
   
    TEMP_VAL = WorksheetFunction.FDist(TEMP_VAL, NROWS - DF_VAL - h, DF_VAL) 'Significance F
    DATA_MATRIX(i, j) = TEMP_VAL
    j = j + 1
1983:
'----------------------------------------------------------------------------------------------------
Next i
'----------------------------------------------------------------------------------------------------


EXCEL_LINEST3_FUNC = DATA_MATRIX

Exit Function
ERROR_LABEL:
EXCEL_LINEST3_FUNC = Err.number
End Function
