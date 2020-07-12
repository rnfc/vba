Attribute VB_Name = "EXCEL_CONTROL_LIBR"


Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.

'************************************************************************************
'************************************************************************************
'FUNCTION      : EXCEL_SHOW_ADDIN_FUNC
'DESCRIPTION   : Show Addin
'LIBRARY       : EXCEL
'GROUP         : ADD_IN
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA

'************************************************************************************
'************************************************************************************

Function EXCEL_SHOW_ADDIN_FUNC(Optional ByRef SRC_WBOOK As Excel.Workbook)
On Error GoTo ERROR_LABEL
    EXCEL_SHOW_ADDIN_FUNC = False
        If SRC_WBOOK Is Nothing Then: Set SRC_WBOOK = ActiveWorkbook
        SRC_WBOOK.IsAddin = False
    EXCEL_SHOW_ADDIN_FUNC = True
Exit Function
ERROR_LABEL:
    EXCEL_SHOW_ADDIN_FUNC = False
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : ADDIN_HIDE
'DESCRIPTION   : Hide Addin
'LIBRARY       : EXCEL
'GROUP         : ADD_IN
'ID            : 002
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA

'************************************************************************************
'************************************************************************************

Function EXCEL_HIDE_ADDIN_FUNC(Optional ByRef SRC_WBOOK As Excel.Workbook)
On Error GoTo ERROR_LABEL
EXCEL_HIDE_ADDIN_FUNC = False
    If SRC_WBOOK Is Nothing Then: Set SRC_WBOOK = ActiveWorkbook
    SRC_WBOOK.IsAddin = True
EXCEL_HIDE_ADDIN_FUNC = True
Exit Function
ERROR_LABEL:
EXCEL_HIDE_ADDIN_FUNC = False
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : EXCEL_CHANGE_REFERENCE_STYLE_FUNC
'DESCRIPTION   : Change Reference Style
'LIBRARY       : EXCEL
'GROUP         : CONTROL
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA

'************************************************************************************
'************************************************************************************

Function EXCEL_CHANGE_REFERENCE_STYLE_FUNC(Optional ByVal VERSION As Long = 0)

On Error GoTo ERROR_LABEL

EXCEL_CHANGE_REFERENCE_STYLE_FUNC = False
    
    Select Case VERSION
        Case 0
            Excel.Application.ReferenceStyle = xlR1C1
        Case Else
            Excel.Application.ReferenceStyle = xlA1
    End Select

EXCEL_CHANGE_REFERENCE_STYLE_FUNC = True

Exit Function
ERROR_LABEL:
EXCEL_CHANGE_REFERENCE_STYLE_FUNC = False
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : EXCEL_DISPLAY_ALERTS_FUNC
'DESCRIPTION   : Turn on/off Display Alerts
'LIBRARY       : EXCEL
'GROUP         : CONTROL
'ID            : 002
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA

'************************************************************************************
'************************************************************************************

Function EXCEL_DISPLAY_ALERTS_FUNC(Optional ByVal VERSION As Boolean = True)

On Error GoTo ERROR_LABEL

EXCEL_DISPLAY_ALERTS_FUNC = False

Select Case VERSION
Case Is = True
    Excel.Application.DisplayAlerts = True
Case Is = False
    Excel.Application.DisplayAlerts = False
End Select

EXCEL_DISPLAY_ALERTS_FUNC = True

Exit Function
ERROR_LABEL:
EXCEL_DISPLAY_ALERTS_FUNC = False
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : EXCEL_TURN_OFF_EVENTS_FUNC
'DESCRIPTION   : Turn off events in Excel
'LIBRARY       : EXCEL
'GROUP         : CONTROL
'ID            : 006
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA

'************************************************************************************
'************************************************************************************

Function EXCEL_TURN_OFF_EVENTS_FUNC()
    
On Error GoTo ERROR_LABEL

EXCEL_TURN_OFF_EVENTS_FUNC = False
With Excel.Application
    .EnableEvents = False
    .Calculation = xlCalculationManual
    .ScreenUpdating = False
    .Cursor = xlWait
    .StatusBar = False
End With
EXCEL_TURN_OFF_EVENTS_FUNC = True

Exit Function
ERROR_LABEL:
EXCEL_TURN_OFF_EVENTS_FUNC = False
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : EXCEL_TURN_ON_EVENTS_FUNC
'DESCRIPTION   : Turn on events in Excel
'LIBRARY       : EXCEL
'GROUP         : CONTROL
'ID            : 007
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA

'************************************************************************************
'************************************************************************************

Function EXCEL_TURN_ON_EVENTS_FUNC()

On Error GoTo ERROR_LABEL
    
EXCEL_TURN_ON_EVENTS_FUNC = False
With Excel.Application
    .Calculation = xlCalculationAutomatic
    .ScreenUpdating = True
    .EnableEvents = True
    .Cursor = xlDefault
    .StatusBar = False
End With
EXCEL_TURN_ON_EVENTS_FUNC = True
Exit Function
ERROR_LABEL:
EXCEL_TURN_ON_EVENTS_FUNC = False
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : EXCEL_CALCULATE_FULL_REBUILD_FUNC
'DESCRIPTION   : Calculate Full Rebuild Sheet
'LIBRARY       : EXCEL
'GROUP         : CONTROL
'ID            : 009
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA

'************************************************************************************
'************************************************************************************
Sub EXCEL_CALCULATE_FULL_REBUILD_FUNC()
On Error Resume Next
If Val(Excel.Application.VERSION) < 10 Then
   Excel.Application.CalculateFull
Else
   Excel.Application.CalculateFullRebuild
End If
End Sub
