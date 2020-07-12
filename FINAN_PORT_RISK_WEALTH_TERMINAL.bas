Attribute VB_Name = "FINAN_PORT_RISK_WEALTH_TERMINAL"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.
   

'************************************************************************************
'************************************************************************************
'FUNCTION      : RNG_PORT_TERMINAL_WEALTH_FUNC

'DESCRIPTION   : Terminal Wealth & VaR Analysis: Various comparative calculations
'related to terminal wealth and Value-At-Risk.

'LIBRARY       : PORT_RISK
'GROUP         : WEALTH_TERMINAL
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 21/01/2009
'************************************************************************************
'************************************************************************************

Function RNG_PORT_TERMINAL_WEALTH_FUNC(ByRef DST_RNG As Excel.Range, _
Optional ByRef NO_SCENARIOS As Long = 10)

Dim i As Long
Dim k As Long

Dim TEMP_RNG As Excel.Range

On Error GoTo ERROR_LABEL

RNG_PORT_TERMINAL_WEALTH_FUNC = False
k = -4165632
DST_RNG.Cells(1, 1) = "INTITIAL WEALTH"
DST_RNG.Cells(1, 2) = "TIME HORIZON"
DST_RNG.Cells(1, 3) = "CONFIDENCE"
DST_RNG.Cells(1, 4) = "ONE-SIDED CONFIDENCE PARAMETER"
DST_RNG.Cells(1, 5) = "TWO-SIDED CONFIDENCE PARAMETER"
DST_RNG.Cells(1, 6) = "DISCRETE RETURN"
DST_RNG.Cells(1, 7) = "DISCRETE VOLATILITY"
DST_RNG.Cells(1, 8) = "CONTINUOUS RETURN"
DST_RNG.Cells(1, 9) = "CONTINUOUS VOLATILITY"
DST_RNG.Cells(1, 10) = "VAR"
DST_RNG.Cells(1, 11) = "EXPECTED TERMINAL WEALTH"
DST_RNG.Cells(1, 12) = "LOWER EXPECTED BOUNDARY"
DST_RNG.Cells(1, 13) = "UPPER EXPECETED BOUNDARY"

Set TEMP_RNG = Range(DST_RNG.Cells(1, 1), DST_RNG.Cells(1, 13))
GoSub FORMAT_LINE

'-----------------------------------------------------------------------------
'Terminal Wealth as Function of Confidence/Time/Volatility/Expected Return
'-----------------------------------------------------------------------------
Set DST_RNG = DST_RNG.Offset(2, 0)
For i = 1 To NO_SCENARIOS
    DST_RNG.Cells(i, 1).formula = _
            "=1000"
    DST_RNG.Cells(i, 2).formula = _
            "=1"
    DST_RNG.Cells(i, 3).formula = _
            "=50%*(1+" & i / 100 & ")"
    DST_RNG.Cells(i, 4).formula = _
            "=NORMSINV(" & DST_RNG.Cells(i, 3).Address & ")"
    DST_RNG.Cells(i, 5).formula = _
            "=-NORMSINV((1-" & DST_RNG.Cells(i, 3).Address & ")/2)"
    DST_RNG.Cells(i, 6).formula = "=8%"
    DST_RNG.Cells(i, 7).formula = "=16%"
    DST_RNG.Cells(i, 8).formula = _
            "=MAX(LN(1+" & DST_RNG.Cells(i, 6).Address & _
            ")-0.5*LN(1+(" & DST_RNG.Cells(i, 7).Address & _
            "/(1+" & DST_RNG.Cells(i, 6).Address & "))^2),0)"
    DST_RNG.Cells(i, 9).formula = _
            "=SQRT(LN(1+(" & DST_RNG.Cells(i, 7).Address & _
            "/(1+" & DST_RNG.Cells(i, 6).Address & "))^2))"
    DST_RNG.Cells(i, 10).formula = _
            "=EXP(" & DST_RNG.Cells(i, 2).Address & _
            "*" & DST_RNG.Cells(i, 8).Address & "-" & _
            DST_RNG.Cells(i, 4).Address & "*" & _
            DST_RNG.Cells(i, 9).Address & "*SQRT(" & _
            DST_RNG.Cells(i, 2).Address & "))-1"
    DST_RNG.Cells(i, 11).formula = _
            "=EXP(" & DST_RNG.Cells(i, 8).Address & "*" & _
            DST_RNG.Cells(i, 2).Address & ")*" & _
            DST_RNG.Cells(i, 1).Address & ""
    DST_RNG.Cells(i, 12).formula = _
            "=" & DST_RNG.Cells(i, 1).Address & "*EXP(" & _
            DST_RNG.Cells(i, 8).Address & "*" & _
            DST_RNG.Cells(i, 2).Address & _
            "-" & DST_RNG.Cells(i, 5).Address & _
            "*SQRT(" & DST_RNG.Cells(i, 2).Address & ")*" & _
            DST_RNG.Cells(i, 9).Address & ")"
    DST_RNG.Cells(i, 13).formula = _
            "=" & DST_RNG.Cells(i, 1).Address & "*EXP(" & _
            DST_RNG.Cells(i, 8).Address & "*" & _
            DST_RNG.Cells(i, 2).Address & _
            "+" & DST_RNG.Cells(i, 5).Address & _
            "*SQRT(" & DST_RNG.Cells(i, 2).Address & ")*" & _
            DST_RNG.Cells(i, 9).Address & ")"
Next i
Set TEMP_RNG = Range(DST_RNG.Cells(1, 1), DST_RNG.Cells(NO_SCENARIOS, 13))
TEMP_RNG.Columns(1).Font.Color = k
TEMP_RNG.Columns(2).Font.Color = k
TEMP_RNG.Columns(3).Font.Color = k
TEMP_RNG.Columns(6).Font.Color = k
TEMP_RNG.Columns(7).Font.Color = k

RNG_PORT_TERMINAL_WEALTH_FUNC = True

Exit Function
'-----------------------------------------------------------------------------
FORMAT_LINE:
'-----------------------------------------------------------------------------
With TEMP_RNG
    .Font.Bold = True
    .HorizontalAlignment = xlCenter
    .VerticalAlignment = xlCenter
    .WrapText = True
    .Orientation = 0
    .AddIndent = False
    .IndentLevel = 0
    .ShrinkToFit = False
    .ReadingOrder = xlContext
    .MergeCells = False
    .Borders(xlDiagonalDown).LineStyle = xlNone
    .Borders(xlDiagonalUp).LineStyle = xlNone
    With .Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .WEIGHT = xlThin
    End With
    With .Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .WEIGHT = xlThin
    End With
    With .Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .WEIGHT = xlThin
    End With
    With .Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .WEIGHT = xlThin
    End With
    With .Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 65535
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
End With
'-----------------------------------------------------------------------------
Return
'-----------------------------------------------------------------------------
ERROR_LABEL:
RNG_PORT_TERMINAL_WEALTH_FUNC = False
End Function

