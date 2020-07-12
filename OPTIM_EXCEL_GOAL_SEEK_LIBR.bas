Attribute VB_Name = "OPTIM_EXCEL_GOAL_SEEK_LIBR"
'// PERFECT

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.


'************************************************************************************
'************************************************************************************
'FUNCTION      : GOAL_SEEK_CALL_FUNC
'DESCRIPTION   : Calculates the values necessary to achieve a specific goal
'through Excel
'LIBRARY       : OPTIMIZATION
'GROUP         : EXCEL_GOAL_SEEK
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/08/2008
'************************************************************************************
'************************************************************************************

Function GOAL_SEEK_CALL_FUNC(ByVal TARGET_VAL As Double, _
ByRef TARGET_CELL As Excel.Range, _
ByRef CHANGING_CELL As Excel.Range)

On Error GoTo ERROR_LABEL
    
GOAL_SEEK_CALL_FUNC = TARGET_CELL.GoalSeek(TARGET_VAL, CHANGING_CELL)

Exit Function
ERROR_LABEL:
GOAL_SEEK_CALL_FUNC = False
End Function

