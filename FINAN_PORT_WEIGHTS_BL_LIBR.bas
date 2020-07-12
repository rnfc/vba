Attribute VB_Name = "FINAN_PORT_WEIGHTS_BL_LIBR"


Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.


'************************************************************************************
'************************************************************************************
'FUNCTION      : RNG_PORT_BLACK_LITTERMAN_WEIGHTS_FUNC

'DESCRIPTION   : This routine calculates the Black-Litterman implied
'expected returns (both excess and total) for a n asset class portfolio.
'The user must provide the following inputs: (i) global risk premium and risk-free
'rate, (ii) asset class market capitalizations, (iii) asset class covariance,
'and (iv) factor and sigma.

'The program then calculates the asset class expected returns consistent with
'the notion that the global market is in equilibrium, as implied by the current
'level of investment.

'LIBRARY       : PORTFOLIO
'GROUP         : FRONTIER_BL
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 21/01/2009
'************************************************************************************
'************************************************************************************

Function RNG_PORT_BLACK_LITTERMAN_WEIGHTS_FUNC(ByRef DST_RNG As Excel.Range, _
Optional ByVal NASSETS As Long = 8)

Dim i As Long
Dim j As Long

Dim MODEL_RNG As Excel.Range
Dim COVAR_RNG As Excel.Range
Dim DUMMY_RNG As Excel.Range
Dim FACTOR_RNG As Excel.Range

'-------------------------------------------------------------------------
'-------------------------------------------------------------------------
On Error GoTo ERROR_LABEL
'-------------------------------------------------------------------------
'-------------------------------------------------------------------------

RNG_PORT_BLACK_LITTERMAN_WEIGHTS_FUNC = False

DST_RNG.Cells(1, 1) = "RISK_AVERSION"
DST_RNG.Cells(2, 1) = "RISK_FREE"
DST_RNG.Cells(3, 1) = "SHARPE"
DST_RNG.Cells(4, 1) = "FACTOR"
Range(DST_RNG.Cells(1, 1), DST_RNG.Cells(4, 1)).Font.Bold = True

DST_RNG.Cells(1, 2) = 5#
DST_RNG.Cells(1, 2).Font.ColorIndex = 5
DST_RNG.Cells(2, 2) = 0.03
DST_RNG.Cells(2, 2).Font.ColorIndex = 5
DST_RNG.Cells(3, 2).formula = _
                    "=(" & DST_RNG.Offset(6, 0).Cells(NASSETS + 2, 3).Address _
                    & "-" & DST_RNG.Cells(2, 2).Address _
                    & ") / " & DST_RNG.Offset(6, 0).Cells(NASSETS + 3, 2).Address
DST_RNG.Cells(3, 2).Font.ColorIndex = 0

DST_RNG.Cells(4, 2) = 0.05
DST_RNG.Cells(4, 2).Font.ColorIndex = 5

DST_RNG.Cells(1, 3) = "SIGMA"
DST_RNG.Cells(2, 3) = "TRACKING_ERROR"
DST_RNG.Cells(3, 3) = "95%_CONFIDENCE"
DST_RNG.Cells(4, 3) = "INFO_RATIO"
Range(DST_RNG.Cells(1, 3), DST_RNG.Cells(4, 3)).Font.Bold = True

DST_RNG.Cells(1, 4) = 0.5
DST_RNG.Cells(1, 4).Font.ColorIndex = 5

DST_RNG.Cells(2, 4).formula = _
                    "=ABS(" & DST_RNG.Offset(6, 0).Cells(NASSETS + 3, 2).Address _
                    & "-" & DST_RNG.Offset(6, 0).Cells(NASSETS + 3, 9).Address _
                    & ")"
DST_RNG.Cells(2, 4).Font.ColorIndex = 0

DST_RNG.Cells(3, 4).formula = _
                    "=(" & DST_RNG.Cells(2, 4).Address _
                    & "*-1*NORMSINV(0.05/2))"
DST_RNG.Cells(3, 4).Font.ColorIndex = 0

DST_RNG.Cells(4, 4).formula = _
                    "=(" & DST_RNG.Offset(6, 0).Cells(NASSETS + 2, 9).Address _
                    & "-" & DST_RNG.Offset(6, 0).Cells(NASSETS + 2, 4).Address _
                    & ") / " & DST_RNG.Cells(2, 4).Address
DST_RNG.Cells(4, 4).Font.ColorIndex = 0

'-------------------------------------------------------------------------
DST_RNG.Cells(6, 1) = ""
DST_RNG.Cells(6, 2) = "WEIGHT"
DST_RNG.Cells(6, 3) = "EQUIL.CAPM RET" 'neutral returns
DST_RNG.Cells(6, 4) = "EST. RETURNS"
DST_RNG.Cells(6, 5) = "BL E[R]" 'Black Litterman
DST_RNG.Cells(6, 6) = "BL BET"
DST_RNG.Cells(6, 7) = "DUMMY"
DST_RNG.Cells(6, 8) = "RISK FREE"
DST_RNG.Cells(6, 9) = "BL NEW WEIGHT" 'Black Litterman
Range(DST_RNG.Cells(6, 1), DST_RNG.Cells(6, 9)).Font.Bold = True
'-------------------------------------------------------------------------
Set MODEL_RNG = DST_RNG.Offset(6, 0)

For i = 1 To NASSETS
    With MODEL_RNG
        .Cells(i, 1) = "ASSET - " & i
        .Cells(i, 1).Font.Bold = True
        .Cells(i, 1).Font.ColorIndex = 3
        
        .Cells(i, 2) = 0
        .Cells(i, 2).Font.ColorIndex = 5
    
        .Cells(i, 4) = 0
        .Cells(i, 4).Font.ColorIndex = 5
        
        .Cells(i, 7) = 1
        .Cells(i, 7).Font.ColorIndex = 5
        
        .Cells(i, 8).formula = "=" & DST_RNG.Cells(2, 2).Address
        .Cells(i, 8).Font.ColorIndex = 0
        
        .Cells(i, 9).formula = "=" & MODEL_RNG.Cells(i, 6).Address & "+" & _
        MODEL_RNG.Cells(i, 2).Address
        .Cells(i, 9).Font.ColorIndex = 0
    
    End With
Next i

'-------------------------------------------------------------------------
Set COVAR_RNG = MODEL_RNG.Offset(0, 10)
COVAR_RNG.Cells(0, 1) = "COVARIANCES_MATRIX"
COVAR_RNG.Cells(0, 1).Font.Bold = True
For i = 1 To NASSETS
    With COVAR_RNG
        .Cells(i, 1).formula = "=" & MODEL_RNG.Cells(i, 1).Address
        .Cells(i, 1).Font.Bold = True
    
        .Cells(0, i + 1).formula = "=" & .Cells(i, 1).Address
        .Cells(0, i + 1).Font.Bold = True
        For j = 1 To NASSETS
            .Cells(i, j + 1) = 0
            .Cells(i, j + 1).Font.ColorIndex = 5
        Next j
    End With
Next i

'-------------------------------------------------------------------------
Set DUMMY_RNG = COVAR_RNG.Offset(0, NASSETS + 2)
DUMMY_RNG.Cells(0, 1) = "DUMMY_MAT"
DUMMY_RNG.Cells(0, 1).Font.Bold = True
For i = 1 To NASSETS
    With DUMMY_RNG
        .Cells(i, 1).formula = "=" & MODEL_RNG.Cells(i, 1).Address
        .Cells(i, 1).Font.Bold = True
    
        .Cells(0, i + 1).formula = "=" & .Cells(i, 1).Address
        .Cells(0, i + 1).Font.Bold = True
        For j = 1 To NASSETS
            If j = i Then
                  .Cells(i, j + 1).formula = "=" & MODEL_RNG.Cells(i, 7).Address
                  .Cells(i, i + 1).Font.ColorIndex = 0
            Else: .Cells(i, j + 1) = 0
                  .Cells(i, j + 1).Font.ColorIndex = 5
            End If
        Next j
    End With
Next i
'-------------------------------------------------------------------------
Set FACTOR_RNG = DUMMY_RNG.Offset(0, NASSETS + 2)
FACTOR_RNG.Cells(0, 1) = "FACTOR_MAT"
FACTOR_RNG.Cells(0, 1).Font.Bold = True
For i = 1 To NASSETS
    With FACTOR_RNG
        .Cells(i, 1).formula = "=" & MODEL_RNG.Cells(i, 1).Address
        .Cells(i, 1).Font.Bold = True
    
        .Cells(0, i + 1).formula = "=" & .Cells(i, 1).Address
        .Cells(0, i + 1).Font.Bold = True
        For j = 1 To NASSETS
            If j = i Then
                  .Cells(i, j + 1).formula = "=" & DST_RNG.Cells(4, 2).Address
                  .Cells(i, i + 1).Font.ColorIndex = 0
            Else: .Cells(i, j + 1) = 0
                  .Cells(i, j + 1).Font.ColorIndex = 5
            End If
        Next j
    End With
Next i

'-------------------------------------------------------------------------
'-------------------------------------------------------------------------
'-------------------------------------------------------------------------
Range(MODEL_RNG.Cells(1, 3), MODEL_RNG.Cells(NASSETS, 3)).FormulaArray = _
    "=" & DST_RNG.Cells(1, 2).Address & "*MMULT(" & _
    Range(COVAR_RNG.Cells(1, 2), COVAR_RNG.Cells(NASSETS, NASSETS + 1)).Address & _
    "," & _
    Range(MODEL_RNG.Cells(1, 2), MODEL_RNG.Cells(NASSETS, 2)).Address & _
    ")+" & DST_RNG.Cells(2, 2).Address
    '
Range(MODEL_RNG.Cells(1, 5), MODEL_RNG.Cells(NASSETS, 5)).FormulaArray = _
    "=MMULT(MINVERSE(MINVERSE(" & DST_RNG.Cells(1, 4).Address & "*" & _
    Range(COVAR_RNG.Cells(1, 2), COVAR_RNG.Cells(NASSETS, NASSETS + 1)).Address & _
    ")+MMULT(MMULT(TRANSPOSE(" & _
    Range(DUMMY_RNG.Cells(1, 2), DUMMY_RNG.Cells(NASSETS, NASSETS + 1)).Address & _
    "),MINVERSE(" & _
    Range(FACTOR_RNG.Cells(1, 2), FACTOR_RNG.Cells(NASSETS, NASSETS + 1)).Address _
    & "))," & _
    Range(DUMMY_RNG.Cells(1, 2), DUMMY_RNG.Cells(NASSETS, NASSETS + 1)).Address & _
    ")),MMULT(MINVERSE(" & DST_RNG.Cells(1, 4).Address & "*" & _
    Range(COVAR_RNG.Cells(1, 2), COVAR_RNG.Cells(NASSETS, NASSETS + 1)).Address & _
    ")," & Range(MODEL_RNG.Cells(1, 3), MODEL_RNG.Cells(NASSETS, 3)).Address & _
    ")+MMULT(MMULT(TRANSPOSE(" & _
    Range(DUMMY_RNG.Cells(1, 2), DUMMY_RNG.Cells(NASSETS, NASSETS + 1)).Address & _
    "),MINVERSE(" & _
    Range(FACTOR_RNG.Cells(1, 2), FACTOR_RNG.Cells(NASSETS, NASSETS + 1)).Address & _
    "))," & Range(MODEL_RNG.Cells(1, 4), MODEL_RNG.Cells(NASSETS, 4)).Address & "))"

Range(MODEL_RNG.Cells(1, 6), MODEL_RNG.Cells(NASSETS, 6)).FormulaArray = _
    "=MMULT(MINVERSE(" & Range(COVAR_RNG.Cells(1, 2), _
    COVAR_RNG.Cells(NASSETS, NASSETS + 1)).Address & _
    ")/" & DST_RNG.Cells(1, 2).Address & _
    ",(" & Range(MODEL_RNG.Cells(1, 5), MODEL_RNG.Cells(NASSETS, 5)).Address & _
    "-" & Range(MODEL_RNG.Cells(1, 3), MODEL_RNG.Cells(NASSETS, 3)).Address & _
    ")-" & Range(MODEL_RNG.Cells(1, 7), MODEL_RNG.Cells(NASSETS, 7)).Address & _
    "*MMULT(MMULT(TRANSPOSE(" & Range(MODEL_RNG.Cells(1, 7), _
    MODEL_RNG.Cells(NASSETS, 7)).Address & "),MINVERSE(" & _
    Range(COVAR_RNG.Cells(1, 2), COVAR_RNG.Cells(NASSETS, NASSETS + 1)).Address & _
    ")),(" & Range(MODEL_RNG.Cells(1, 5), MODEL_RNG.Cells(NASSETS, 5)).Address & _
    "-" & Range(MODEL_RNG.Cells(1, 3), MODEL_RNG.Cells(NASSETS, 3)).Address & _
    "))/MMULT(MMULT(TRANSPOSE(" & Range(MODEL_RNG.Cells(1, 7), _
    MODEL_RNG.Cells(NASSETS, 7)).Address & "),MINVERSE(" & _
    Range(COVAR_RNG.Cells(1, 2), COVAR_RNG.Cells(NASSETS, NASSETS + 1)).Address & _
    "))," & Range(MODEL_RNG.Cells(1, 7), MODEL_RNG.Cells(NASSETS, 7)).Address & "))"
    

MODEL_RNG.Cells(NASSETS + 2, 1) = "RETURN"
MODEL_RNG.Cells(NASSETS + 2, 1).Font.Bold = True

MODEL_RNG.Cells(NASSETS + 2, 3).formula = "=SUMPRODUCT(" & _
    Range(MODEL_RNG.Cells(1, 2), MODEL_RNG.Cells(NASSETS, 2)).Address & "," & _
    Range(MODEL_RNG.Cells(1, 3), MODEL_RNG.Cells(NASSETS, 3)).Address & ")"

MODEL_RNG.Cells(NASSETS + 2, 4).formula = "=SUMPRODUCT(" & _
    Range(MODEL_RNG.Cells(1, 2), MODEL_RNG.Cells(NASSETS, 2)).Address & "," & _
    Range(MODEL_RNG.Cells(1, 4), MODEL_RNG.Cells(NASSETS, 4)).Address & ")"
    
MODEL_RNG.Cells(NASSETS + 2, 5).formula = "=SUMPRODUCT(" & _
    Range(MODEL_RNG.Cells(1, 2), MODEL_RNG.Cells(NASSETS, 2)).Address & "," & _
    Range(MODEL_RNG.Cells(1, 5), MODEL_RNG.Cells(NASSETS, 5)).Address & ")"
    
MODEL_RNG.Cells(NASSETS + 2, 9).formula = "=SUMPRODUCT(" & _
    Range(MODEL_RNG.Cells(1, 9), MODEL_RNG.Cells(NASSETS, 9)).Address & "," & _
    Range(MODEL_RNG.Cells(1, 5), MODEL_RNG.Cells(NASSETS, 5)).Address & ")"


MODEL_RNG.Cells(NASSETS + 3, 1) = "RISK"

MODEL_RNG.Cells(NASSETS + 3, 2).FormulaArray = _
    "=SQRT(MMULT(TRANSPOSE(" & Range(MODEL_RNG.Cells(1, 2), _
    MODEL_RNG.Cells(NASSETS, 2)).Address & "),MMULT(" & _
    Range(COVAR_RNG.Cells(1, 2), COVAR_RNG.Cells(NASSETS, NASSETS + 1)).Address & _
    "," & Range(MODEL_RNG.Cells(1, 2), MODEL_RNG.Cells(NASSETS, 2)).Address & ")))"

MODEL_RNG.Cells(NASSETS + 3, 9).FormulaArray = _
    "=SQRT(MMULT(TRANSPOSE(" & Range(MODEL_RNG.Cells(1, 9), _
    MODEL_RNG.Cells(NASSETS, 9)).Address & "),MMULT(" & _
    Range(COVAR_RNG.Cells(1, 2), COVAR_RNG.Cells(NASSETS, NASSETS + 1)).Address & _
    "," & Range(MODEL_RNG.Cells(1, 9), MODEL_RNG.Cells(NASSETS, 9)).Address & ")))"

MODEL_RNG.Cells(NASSETS + 3, 1).Font.Bold = True

RNG_PORT_BLACK_LITTERMAN_WEIGHTS_FUNC = True

Exit Function
ERROR_LABEL:
RNG_PORT_BLACK_LITTERMAN_WEIGHTS_FUNC = False
End Function
