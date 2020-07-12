Attribute VB_Name = "ADD_INS_CONTROL"
'Error Trapping with Visual Basic for Applications
'http://support.microsoft.com/kb/146864
'For solver use: CVErr(xlErrNA) / CVErr(xlErrNum)
'http://www.cpearson.com/excel/vbe.aspx

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.
Option Private Module

Private PUB_ROUTINES_ARR As Variant

'Permission for free direct or derivative use is granted subject to
'compliance with any conditions that the originators of the algorithm
'place on its exploitation.

'This routine is run every time the application is opened.
'It handles initialization of the application.

Public Sub Auto_Open()

Dim i As Long
Dim T_OBJ As Office.CommandBarPopup
Dim A_OBJ As Office.CommandBarButton

On Error Resume Next

Call RESET_WEB_DATA_SYSTEM_FUNC
If IsArray(PUB_ROUTINES_ARR) = False Then: Call LOAD_ROUTINES_FUNC
'Adds a menu item to the Tools menu
Set T_OBJ = Excel.Application.CommandBars(1).FindControl(, 30007)
For i = LBound(PUB_ROUTINES_ARR, 1) To UBound(PUB_ROUTINES_ARR, 1)
    Set A_OBJ = T_OBJ.Controls.Add(msoControlButton)
    A_OBJ.Caption = CStr(PUB_ROUTINES_ARR(i, 2))
    A_OBJ.OnAction = CStr(PUB_ROUTINES_ARR(i, 1))
Next i
    
End Sub


'This routine is run every time the application is closed.
'It handles the orderly shutdown of the application.

Public Sub Auto_Close()

Dim i As Long
Dim T_OBJ As Office.CommandBarPopup

On Error Resume Next

'Excel.Application.OnTime Now, "SYSTEM_BACKUP_FUNC"

If IsArray(PUB_ROUTINES_ARR) = False Then: Call LOAD_ROUTINES_FUNC
''' Removes the Code Documentor menu items.
Set T_OBJ = Excel.Application.CommandBars(1).FindControl(, 30007)
For i = LBound(PUB_ROUTINES_ARR, 1) To UBound(PUB_ROUTINES_ARR, 1)
    T_OBJ.Controls(CStr(PUB_ROUTINES_ARR(i, 2))).Delete
Next i
Erase PUB_ROUTINES_ARR
    
End Sub


Private Function LOAD_ROUTINES_FUNC()

Dim i As Long
Dim j As Long
Dim k As Long

Dim HEADINGS_ARR As Variant

On Error Resume Next
'    "SHOW_YAHOO_INDEX_QUOTES_FORM_FUNC", "Index Quote(s)", _
'    "PRINT_ADVFN_STATISTICS", "ADVFN Financial Summary", _
'    "SHOW_FINVIZ_FORM_FUNC", "Finviz Screener", _

HEADINGS_ARR = Array( _
    "REMOVE_OLD_ADDINS_LINKS_FUNC", "Remove Old AddIns Links", _
    "EXCEL_CALCULATE_FULL_REBUILD_FUNC", "Recalculate Worksheet", _
    "WSHEETS_REMOVE_CURRENT_FUNC", "Remove Worksheets", _
    "RESET_WEB_DATA_SYSTEM_FUNC", "Reset System Data Cache", _
    "SHOW_YAHOO_QUOTES_FORM_FUNC", "Stock Quote(s)", _
    "PRINT_YAHOO_OPTION_QUOTES_FUNC", "Option Quote(s)", _
    "SHOW_YAHOO_FX_QUOTES_FORM_FUNC", "FX Quote(s)", _
    "SHOW_YAHOO_HISTORICAL_DATA_FORM_FUNC", "Historical Data", _
    "PRINT_YAHOO_KEY_STATISTICS", "Yahoo Key Stats", _
    "SHOW_FINANCIAL_CHARTS_FORM", "eFinancial Charts")

k = UBound(HEADINGS_ARR) - LBound(HEADINGS_ARR) + 1: k = k / 2
ReDim PUB_ROUTINES_ARR(1 To k, 1 To 2)
j = LBound(HEADINGS_ARR)
For i = 1 To k
    PUB_ROUTINES_ARR(i, 1) = HEADINGS_ARR(j + 0)
    PUB_ROUTINES_ARR(i, 2) = HEADINGS_ARR(j + 1)
    j = j + 2
Next i

'The web browser will appear. When you close it, you can reload
'the browser from the Tools menu."
    
End Function

