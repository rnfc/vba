Attribute VB_Name = "OPTIM_EXCEL_SOLVER_LIBR"

'-------------------------------------------------------------------------------------
Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.
'-------------------------------------------------------------------------------------

'atpvbaen.xls! --> Analysis ToolPak
'Private PUB_SOLVER_MATRIX As Variant

'************************************************************************************
'************************************************************************************
'FUNCTION      : CALL_EXCEL_SOLVER_FUNC

'DESCRIPTION   : Check if argument is a vector array

'The following SOLVER Function will only run if your Excel
'is set up as follows:

'1) You must have SOLVER installed with your Excel.
'2) Go to Tools Menu and see whether item Solver appears there.
'3) If it does not, go to Tools - Add-ins and tick "Solver Add-in".

'This 1st step will allow you to use SOLVER from Excel but because SOLVER is
'also called by a VBA macro, you will also need to establish a reference to
'the Solver add-in in the VBA editor:

'With a Visual Basic module active, click References on the Tools menu, and
'then select the Solver.xlam check box under Available References. If Solver.xlam
'doesn't appear under Available References, click Browse and open Solver.xlam
'in the \Office\Library subfolder.
'----------------------------------------------------------------------------------
'Before using any of these functions, you must establish a reference to
'the Solver add-in. With a Visual Basic module active, click References
'on the Tools menu, and then select the Solver.xlam check box under Available
'References. If Solver.xlam doesn't appear under Available References, click
'Browse and open Solver.xlam in the \Office\Library subfolder.
'----------------------------------------------------------------------------------

'LIBRARY       : OPTIMIZATION
'GROUP         : EXCEL_SOLVER
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/08/2008
'************************************************************************************
'************************************************************************************

Function CALL_EXCEL_SOLVER_FUNC(ByRef TARGET_RNG As Excel.Range, _
ByRef CHG_CELLS_RNG As Excel.Range, _
ByRef REFER_CELLS_RNG() As Excel.Range, _
ByRef CONST_CELLS_RNG() As Excel.Range, _
ByRef RELATION_ARR() As Integer, _
Optional ByVal SOLVER_MODE As Integer = 2, _
Optional ByVal MATCH_VALUE As Double = 0, _
Optional ByVal MAX_TIME As Double = 100, _
Optional ByVal ITERATIONS As Double = 100, _
Optional ByVal PRECISION As Double = 0.000001, _
Optional ByVal LINEAR_FLAG As Boolean = False, _
Optional ByVal STEPS_FLAG As Boolean = False, _
Optional ByVal EST_VAL As Variant = 1, _
Optional ByVal DERIVAT_VAL As Variant = 1, _
Optional ByVal SEARCH_VAL As Variant = 1, _
Optional ByVal TOLERANCE_VAL As Double = 5, _
Optional ByVal SCALE_FLAG As Boolean = False, _
Optional ByVal CONVERGENCE As Double = 0.0001, _
Optional ByVal NON_NEG_FLAG As Boolean = False)

'If RELATION_ARR(y) = 1 then constraint is <=
'If RELATION_ARR(y) = 2 then constraint is =
'If RELATION_ARR(y) = 3 then constraint is >=

'If SOLVER_MODE = 1 Then solver will maximize
'If SOLVER_MODE = 2 then solver will minimize
'If SOLVER_MODE = 3 then solver will match specified value (MATCH_VALUE)

Dim i As Long
Dim NCOLUMNS As Long 'CONSTRAINTS

On Error GoTo ERROR_LABEL

CALL_EXCEL_SOLVER_FUNC = False

If UBound(REFER_CELLS_RNG()) <> UBound(RELATION_ARR()) Then: GoTo ERROR_LABEL

NCOLUMNS = UBound(REFER_CELLS_RNG())
'ReDim PUB_SOLVER_MATRIX(1 To NCOLUMNS + 5, 1 To 1)

'PUB_SOLVER_MATRIX(1, 1) = Excel.Application.Run("Solver.xlam!SolverReset")
Excel.Application.Run "Solver.xlam!SolverReset"
For i = 1 To NCOLUMNS 'CONSTRAINTS
'    PUB_SOLVER_MATRIX(i + 1, 1) = Excel.Application.Run("Solver.xlam!SolverAdd", _
        REFER_CELLS_RNG(i).Address, RELATION_ARR(i), _
        CONST_CELLS_RNG(i).Address)
    Excel.Application.Run "Solver.xlam!SolverAdd", _
        REFER_CELLS_RNG(i).Address, RELATION_ARR(i), _
        CONST_CELLS_RNG(i).Address
'---------------------------------------------------------------------------------------
    'SolverAdd CellRef:="", Relation:=, FormulaText:=""
'---------------------------------------------------------------------------------------
Next i

'PUB_SOLVER_MATRIX(i, 1) = Excel.Application.Run("Solver.xlam!SolverOk", _
TARGET_RNG.Address, SOLVER_MODE, MATCH_VALUE, CHG_CELLS_RNG.Address)

Excel.Application.Run "Solver.xlam!SolverOk", _
    TARGET_RNG.Address, SOLVER_MODE, MATCH_VALUE, _
    CHG_CELLS_RNG.Address

'PUB_SOLVER_MATRIX(i + 1, 1) = Excel.Application.Run("Solver.xlam!SolverOptions", _
MAX_TIME, ITERATIONS, PRECISION, LINEAR_FLAG, STEPS_FLAG, EST_VAL, _
DERIVAT_VAL, SEARCH_VAL, TOLERANCE_VAL, _
SCALE_FLAG, CONVERGENCE, NON_NEG_FLAG)

Excel.Application.Run "Solver.xlam!SolverOptions", _
    MAX_TIME, ITERATIONS, PRECISION, LINEAR_FLAG, STEPS_FLAG, EST_VAL, _
    DERIVAT_VAL, SEARCH_VAL, TOLERANCE_VAL, _
    SCALE_FLAG, CONVERGENCE, NON_NEG_FLAG


'PUB_SOLVER_MATRIX(i + 2, 1) = Excel.Application.Run("Solver.xlam!SolverSolve", True)
Excel.Application.Run "Solver.xlam!SolverSolve", True 'SolverSolve UserFinish:=True

'Debug.Print PUB_SOLVER_MATRIX(i + 2, 1)
' report on success of analysis
'0  Solver found a solution. All constraints and optimality conditions are satisfied.
'1  Solver has converged to the current solution. All constraints are satisfied.
'2  Solver cannot improve the current solution. All constraints are satisfied.
'3  Stop chosen when the maximum iteration limit was reached.
'4  The Set Cell values do not converge.
'5  Solver could not find a feasible solution.
'6  Solver stopped at user's request.
'7  The conditions for Assume Linear Model are not satisfied.
'8  The problem is too large for Solver to handle.
'9  Solver encountered an error value in a target or constraint cell.
'10  Stop chosen when maximum time limit was reached.
'11  There is not enough memory available to solve the problem.
'12  Another Excel instance is using SOLVER.DLL. Try again later.
'13  Error in model. Please verify that all cells and constraints are valid.

CALL_EXCEL_SOLVER_FUNC = True

Exit Function
ERROR_LABEL:
CALL_EXCEL_SOLVER_FUNC = False
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : CHECK_EXCEL_SOLVER_FUNC
'DESCRIPTION   : Establish a reference to the Solver add-in.
'LIBRARY       : OPTIMIZATION
'GROUP         : EXCEL_SOLVER
'ID            : 002
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/08/2008
'************************************************************************************
'************************************************************************************

Public Function CHECK_EXCEL_SOLVER_FUNC() As Boolean

'Adjusted for Excel.Application.Run() to avoid Reference problems with Solver
  
  Dim INSTALLED_FLAG As Boolean

  '' Assume true unless otherwise
  CHECK_EXCEL_SOLVER_FUNC = True

  On Error Resume Next
  ' check whether Solver is installed
  INSTALLED_FLAG = Excel.Application.AddIns("Solver Add-In").Installed
  Err.Clear

  If INSTALLED_FLAG Then
    ' uninstall temporarily
    Excel.Application.AddIns("Solver Add-In").Installed = False
    ' check whether Solver is installed (should be false)
    INSTALLED_FLAG = Excel.Application.AddIns("Solver Add-In").Installed
  End If

  If Not INSTALLED_FLAG Then
    ' (re)install Solver
    Excel.Application.AddIns("Solver Add-In").Installed = True
    ' check whether Solver is installed (should be true)
    INSTALLED_FLAG = Excel.Application.AddIns("Solver Add-In").Installed
  End If

  If Not INSTALLED_FLAG Then
    'MsgBox "Solver not found. This workbook will not work.", vbCritical
    CHECK_EXCEL_SOLVER_FUNC = False
  End If

  If CHECK_EXCEL_SOLVER_FUNC Then
    ' initialize Solver
    If Val(Excel.Application.VERSION) >= 12 Then
        Excel.Application.Run "solver.xlam!SOLVER.Solver2.Auto_open"
    Else
        Excel.Application.Run "solver.xla!SOLVER.Solver2.Auto_open"
    End If
  End If

  On Error GoTo 0

End Function

