Attribute VB_Name = "EXCEL_DO_EVENTS_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.

Public PUB_STOP_DO_EVENTS_FLAG As Boolean

'************************************************************************************
'************************************************************************************
'FUNCTION      : PROCEDURE_DO_EVENTS_FUNC
'DESCRIPTION   : The function is useful for allowing the user to test a time
'consuming process. It allows the user to interrupt execution or exit
'the application.
'LIBRARY       : TIME
'GROUP         : DO EVENTS
'ID            : 001
'LAST UPDATE   : 12/02/2008
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'************************************************************************************
'************************************************************************************

Public Function PROCEDURE_DO_EVENTS_FUNC( _
Optional ByVal FUNC_NAME_STR As String = "", _
Optional ByVal nLOOPS As Double = 10 ^ 5, _
Optional ByVal STOP_TIME As Double = 0)

'TO TEST:
'Dim STOP_TIME As Double
'STOP_TIME = Now + TimeSerial(0, 0, 5)
'Debug.Print PROCEDURE_DO_EVENTS_FUNC(, , STOP_TIME)

Dim i As Double
Dim START_TIMER As Double

On Error GoTo ERROR_LABEL

PUB_STOP_DO_EVENTS_FLAG = False

START_TIMER = Timer
For i = 1 To nLOOPS
     If DO_EVENTS_FUNC() = False Then Exit For
     If STOP_TIME <> 0 Then If Now() = STOP_TIME Then Exit For
     If FUNC_NAME_STR = "" Then
        Excel.Application.CalculateFullRebuild
     Else
        Call Excel.Application.Run(FUNC_NAME_STR)
     End If
     Debug.Print "Tick", (i & ": " & Now)
Next i

PROCEDURE_DO_EVENTS_FUNC = nLOOPS / (Timer - START_TIMER) / 1000 / 32

Exit Function
ERROR_LABEL:
PROCEDURE_DO_EVENTS_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : PRINT_PROCEDURE_FUNC
'DESCRIPTION   : Print Procedure Results
'LIBRARY       : TIME
'GROUP         : DO EVENTS
'ID            : 002
'LAST UPDATE   : 12/02/2008
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'************************************************************************************
'************************************************************************************

Function PRINT_PROCEDURE_FUNC(ByVal FUNC_NAME_STR As String, _
Optional ByVal FILE_STR_NAME As String = _
"C:\Documents and Settings\HOME\Desktop\NICO.txt", _
Optional ByVal nLOOPS As Long = 1000)

Dim i As Long
Dim j As Integer
Dim TEMP_STR As String

On Error GoTo ERROR_LABEL

PRINT_PROCEDURE_FUNC = False

j = FreeFile()
Open FILE_STR_NAME For Output As #j    'open the output file

Print #j, Trim(nLOOPS); " outputs of " & FUNC_NAME_STR
For i = 1 To nLOOPS
    TEMP_STR = Excel.Application.Run(FUNC_NAME_STR)
    Print #j, "Tick " & CStr(i) & ": " & TEMP_STR
Next i

Close #j

PRINT_PROCEDURE_FUNC = True

Exit Function
ERROR_LABEL:
PRINT_PROCEDURE_FUNC = False
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : DO_EVENTS_BREAK_FUNC
'DESCRIPTION   : Yields execution so that the operating system can
'process other events for X seconds.
'LIBRARY       : TIME
'GROUP         : DO EVENTS
'ID            : 003
'LAST UPDATE   : 12/02/2008
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'************************************************************************************
'************************************************************************************

Public Function DO_EVENTS_BREAK_FUNC(ByVal SECONDS_INTERVAL As Long)
Dim START_TIMER As Double
On Error GoTo ERROR_LABEL
DO_EVENTS_BREAK_FUNC = False
'In Seconds
START_TIMER = Timer
Do While Timer - START_TIMER < SECONDS_INTERVAL: DoEvents: Loop
DO_EVENTS_BREAK_FUNC = True
Exit Function
ERROR_LABEL:
DO_EVENTS_BREAK_FUNC = False
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : PROCEDURES_SWITCHER_FUNC
'DESCRIPTION   : This flag interrupt execution or exit the procedure.
'LIBRARY       : TIME
'GROUP         : DO EVENTS
'ID            : 004
'LAST UPDATE   : 12/02/2008
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'************************************************************************************
'************************************************************************************

Public Sub PROCEDURES_SWITCHER_FUNC()
If PUB_STOP_DO_EVENTS_FLAG = False Then
   PUB_STOP_DO_EVENTS_FLAG = True
Else
   PUB_STOP_DO_EVENTS_FLAG = False
End If
End Sub

'************************************************************************************
'************************************************************************************
'FUNCTION      : DO_EVENTS_FUNC
'DESCRIPTION   : Calling this function returns control to the system, so that
'it can process any other events: key presses, timers etc. It returns when all
'the events have been processed.
'LIBRARY       : TIME
'GROUP         : DO EVENTS
'ID            : 005
'LAST UPDATE   : 12/02/2008
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'************************************************************************************
'************************************************************************************

Public Function DO_EVENTS_FUNC() As Boolean
On Error GoTo ERROR_LABEL
Select Case PUB_STOP_DO_EVENTS_FLAG
Case False
    DoEvents 'yields execution so that the operating system can
    'process other events. If the return value is False, your
    'application can continue to execute normally. If the return
    'value is True, this indicates that the user is switching to
    'another application; you should immediately interrupt any
    'processing taking place.
    DO_EVENTS_FUNC = True
Case True
    DO_EVENTS_FUNC = False
End Select
    
'-----------------------------------------------------------------------------
'The following example uses the DoEvents function to cause execution
'to yield to the operating system once every 1000 iterations of the
'loop. DoEvents returns the number of open Visual Basic forms, but
'only when the host application is Visual Basic.

'Create a variable to hold number of Visual Basic forms loaded and
'visible.
'Dim i, OpenForms
'For i = 1 To 150000    ' Start loop.
'    If i Mod 1000 = 0 Then     ' If loop has repeated 1000 times.
'        OpenForms = DoEvents    ' Yield to operating system.
'    End If
'Next i    ' Increment loop counter.
'-----------------------------------------------------------------------------

Exit Function
ERROR_LABEL:
DO_EVENTS_FUNC = False
End Function
