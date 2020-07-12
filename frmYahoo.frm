VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmYahoo 
   Caption         =   "Yahoo Historical Data"
   ClientHeight    =   4725
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5925
   OleObjectBlob   =   "frmYahoo.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmYahoo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.

Private PUB_YAHOO_TICKERS_RNG As Excel.Range

Private Sub btnClose_Click()
    Unload Me
    End
End Sub

Private Sub btnSymbols_Click()
On Error Resume Next
Call EXCEL_TURN_ON_EVENTS_FUNC
Call EXCEL_DISPLAY_ALERTS_FUNC(True)
Set PUB_YAHOO_TICKERS_RNG = Nothing
Me.Hide ' (False)
Set PUB_YAHOO_TICKERS_RNG = Excel.Application.InputBox("Symbols", "Yahoo Finance", , , , , , 8)
Me.show
End Sub


Private Sub btnDownload_Click()

Dim START_DATE As Date
Dim END_DATE As Date

Dim PERIOD_STR As String
Dim ELEMENT_STR As String

Dim HEADER_FLAG As Boolean
Dim ADJUST_FLAG As Boolean
Dim RESORT_FLAG As Boolean
Dim VALIDATE_FLAG As Boolean

Dim TICKERS_VECTOR As Variant
Dim DST_WSHEET As Excel.Worksheet
Dim TEMP_FLAG As Boolean

On Error GoTo ERROR_LABEL

'calStartDate.Refresh
'calEndDate.Refresh
    
If PUB_YAHOO_TICKERS_RNG Is Nothing Then: GoTo ERROR_LABEL
If PUB_YAHOO_TICKERS_RNG.Cells.COUNT = 1 Then
    ReDim TICKERS_VECTOR(1 To 1, 1 To 1)
    TICKERS_VECTOR(1, 1) = PUB_YAHOO_TICKERS_RNG
Else
    TICKERS_VECTOR = PUB_YAHOO_TICKERS_RNG
    If UBound(TICKERS_VECTOR, 1) = 1 Then
        TICKERS_VECTOR = MATRIX_TRANSPOSE_FUNC(TICKERS_VECTOR)
    End If
End If

' Capture the dates.
START_DATE = CDate(calStartDate.value)
END_DATE = CDate(calEndDate.value)

' Make sure the beginning date isn't after the ending date.
If START_DATE >= END_DATE Then
    MsgBox "The starting date should be before the ending date.", _
        vbInformation, "Invalid dates"
    calStartDate.SetFocus
    Exit Sub
'    ElseIf START_DATE < DateSerial(1970, 1, 1) Then
'        MsgBox "The starting date shouldn't be before 1970.", _
'                vbInformation, "Start date too early"
'        calStartDate.SetFocus
'        Exit Sub
'ElseIf END_DATE > Date Then
 '   MsgBox "The ending date shouldn't be after today's date.", _
            vbInformation, "End date too late"
  '  calEndDate.SetFocus
   ' Exit Sub
End If

HEADER_FLAG = chkHeader
ADJUST_FLAG = chkAdjust
RESORT_FLAG = chkResort
VALIDATE_FLAG = chkValidate

Select Case True
    Case optMonth: PERIOD_STR = "Monthly"
    Case optWeek: PERIOD_STR = "Weekly"
    Case optDay: PERIOD_STR = "Daily"
    Case OptDiv: PERIOD_STR = "Dividends"
End Select

Select Case True
    Case optOpen: ELEMENT_STR = "Open"
    Case optHigh: ELEMENT_STR = "High"
    Case optLow: ELEMENT_STR = "Low"
    Case optClose: ELEMENT_STR = "Close"
    Case optVolume: ELEMENT_STR = "Volume"
    Case optAdj: ELEMENT_STR = "Adj. Close"
End Select

Me.Hide
Call EXCEL_TURN_OFF_EVENTS_FUNC
Call EXCEL_DISPLAY_ALERTS_FUNC(False)
    Set DST_WSHEET = WSHEET_ADD_FUNC(PARSE_CURRENT_TIME_FUNC("_"), ActiveSheet.Parent)
'    Debug.Print START_DATE, END_DATE
    TEMP_FLAG = PRINT_YAHOO_HISTORICAL_DATA_SERIES_FUNC( _
                DST_WSHEET.Cells(6, 2), TICKERS_VECTOR, START_DATE, _
                END_DATE, PERIOD_STR, ELEMENT_STR, _
                HEADER_FLAG, ADJUST_FLAG, _
                RESORT_FLAG, VALIDATE_FLAG)
    If TEMP_FLAG = False Then: GoTo ERROR_LABEL
Call EXCEL_TURN_ON_EVENTS_FUNC
Call EXCEL_DISPLAY_ALERTS_FUNC(True)
Me.show

Exit Sub
ERROR_LABEL:
On Error Resume Next
Call EXCEL_TURN_ON_EVENTS_FUNC
Call EXCEL_DISPLAY_ALERTS_FUNC(True)
Me.show
'ADD MSG HERE; Err.Description
End Sub


Private Sub MultiPage1_Change()

End Sub

Private Sub UserForm_Initialize()
    Set PUB_YAHOO_TICKERS_RNG = Nothing
'    MsgBox "Choose the time interval and the starting and ending dates for the data you want to retrieve. The starting date shouldn't be before 1970, and the ending date shouldn't be after today's date.", vbInformation, "Yahoo Historical Data"
    calStartDate.value = Format(DateSerial(Year(Date) - 5, Month(Date), Day(Date)), "dd/mm/yyyy")
    calEndDate.value = Format(Date, "dd/mm/yyyy")
    optMonth.value = True
    optAdj.value = True
    chkHeader.value = True
    chkResort.value = True
End Sub
