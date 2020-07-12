Attribute VB_Name = "FINAN_PORT_WEIGHTS_CONTRIB_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.


'************************************************************************************
'************************************************************************************
'FUNCTION      : RNG_PORT_RISK_CONTRIBUTION_OPTIMISATION_FUNC
'DESCRIPTION   : Calculation of contributions to absolute and relative return,
'risk and risk-adjusted performance.

'LIBRARY       : PORTFOLIO
'GROUP         : RISK_CONTRIBUTION
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/08/2008
'************************************************************************************
'************************************************************************************
'RNG_PORT_RISK_CONTRIBUTION_FUNC

Function RNG_PORT_RISK_CONTRIBUTION_OPTIMISATION_FUNC(ByRef DST_RNG As Excel.Range, _
ByVal NASSETS As Long, _
Optional ByVal FACTOR As Variant = 100, _
Optional ByVal ADD_RNG_NAME As Boolean = False)

Dim i As Long
Dim j As Long

Dim k As Variant
Dim m As Long

Dim CORREL_POS_RNG As Excel.Range
Dim VAR_COVAR_POS_RNG As Excel.Range
Dim CHOL_POS_RNG As Excel.Range
Dim WEIGHT_POS_RNG As Excel.Range
Dim COVAR_POS_RNG As Excel.Range
Dim BETA_POS_RNG As Excel.Range
Dim ABS_REL_POS_RNG As Excel.Range
Dim MARG_POS_RNG As Excel.Range
Dim CONT_POS_RNG As Excel.Range
Dim PER_POS_RNG As Excel.Range
Dim MIN_POS_RNG As Excel.Range

Dim CORREL_RNG As Excel.Range
Dim VOL_RNG As Excel.Range
Dim VAR_COVAR_RNG As Excel.Range
Dim CHOL_RNG As Excel.Range
Dim WEIGHTS_RNG As Excel.Range
Dim COVAR_RNG As Excel.Range
Dim BETA_RNG As Excel.Range
Dim ABS_REL_RNG As Excel.Range
Dim MARG_RNG As Excel.Range
Dim CONT_RNG As Excel.Range
Dim PER_RNG As Excel.Range
Dim MIN_RNG As Excel.Range

On Error GoTo ERROR_LABEL

RNG_PORT_RISK_CONTRIBUTION_OPTIMISATION_FUNC = False

k = 2
m = 6

If NASSETS < 2 Then: GoTo ERROR_LABEL

'---------------------------------------------------------------------------
'----------------------FIRST PASS: ASSET_CLASS_RISK-------------------------
'---------------------------------------------------------------------------

Set CORREL_POS_RNG = DST_RNG

With CORREL_POS_RNG
   Set VOL_RNG = Range(.Offset(NASSETS, 0), .Offset(1, 0))
   Set CORREL_RNG = Range(.Offset(NASSETS, 1), .Offset(1, NASSETS))
   If ADD_RNG_NAME = True Then: CORREL_RNG.name = "CORREL_MAT"
    
    For i = 1 To NASSETS
      With .Offset(0, i)
         .value = "Asset " & CStr(i)
         .Font.ColorIndex = 3
      End With
      With .Offset(i, -1)
          .formula = "=offset(" & CORREL_POS_RNG.Address & ",0," & CStr(i) & ")"
      End With
      With .Offset(i, 0)
          .value = 0
          .Font.ColorIndex = 5
      End With
    
    Next i
    
    With .Offset(-3, -1)
        .value = "1. ASSET CLASS RISKS"
        .Font.Bold = True
    End With
    With .Offset(-1, -1)
        .value = "SIGMA-CORRELATION MATRIX"
        .Font.Bold = True
    End With
    With .Offset(0, 0)
        .value = "Volatility"
    End With
End With

With CORREL_RNG
    For i = 1 To NASSETS 'NAssets - 1
         For j = i To NASSETS  'For j = i + 1
             With .Cells(i, j)
                 .value = 0
                 .Font.ColorIndex = 5
             End With
         Next j
         .Cells(i, i) = 1
    Next i
       
    For i = 2 To NASSETS
         For j = 1 To i - 1
             With .Cells(i, j)
                 .formula = "=offset(" & CORREL_POS_RNG.Address & _
                 "," & CStr(j) & "," & CStr(i) & ")"
             End With
         Next j
    Next i
End With

Set VAR_COVAR_POS_RNG = CORREL_POS_RNG.Offset(0, NASSETS + k)

With VAR_COVAR_POS_RNG
     Set VAR_COVAR_RNG = Range(.Offset(NASSETS, 1), .Offset(1, NASSETS))
     If ADD_RNG_NAME = True Then: VAR_COVAR_RNG.name = "VARCOV_MAT"
      
      For i = 1 To NASSETS
            With .Offset(0, i)
                .formula = "=offset(" & CORREL_POS_RNG.Address & ",0," & _
                CStr(i) & ")"
            End With
            
            With .Offset(i, 0)
                .formula = "=offset(" & CORREL_POS_RNG.Address & ",0," & _
                CStr(i) & ")"
            End With
      Next i

    With .Offset(-1, 0)
        .value = "VAR-COV MATRIX"
        .Font.Bold = True
    End With
End With

VAR_COVAR_RNG.FormulaArray = "=" & CORREL_RNG.Address & "*(TRANSPOSE(" & _
VOL_RNG.Address & ")*" & VOL_RNG.Address & ")"


Set CHOL_POS_RNG = VAR_COVAR_POS_RNG.Offset(0, NASSETS + k)

With CHOL_POS_RNG
     Set CHOL_RNG = Range(.Offset(NASSETS, 1), .Offset(1, NASSETS))
     If ADD_RNG_NAME = True Then: CHOL_RNG.name = "CHOL_MAT"
      
      For i = 1 To NASSETS
            With .Offset(0, i)
                .formula = "=offset(" & CORREL_POS_RNG.Address & ",0," & _
                CStr(i) & ")"
            End With
            
            With .Offset(i, 0)
                .formula = "=offset(" & CORREL_POS_RNG.Address & ",0," & _
                CStr(i) & ")"
            End With
      Next i

    With .Offset(-1, 0)
        .value = "CHOLESKI MATRIX"
        .Font.Bold = True
    End With
End With

CHOL_RNG.FormulaArray = "=MATRIX_CHOLESKY_FUNC(" & CORREL_RNG.Address & ")"

'---------------------------------------------------------------------------
'----------------------SECOND PASS: Portfolio and Benchmark Weights---------
'---------------------------------------------------------------------------

Set WEIGHT_POS_RNG = CORREL_POS_RNG.Offset(NASSETS + m, 0)

With WEIGHT_POS_RNG
     Set WEIGHTS_RNG = Range(.Offset(NASSETS, 0), .Offset(1, 2))
     If ADD_RNG_NAME = True Then: WEIGHTS_RNG.name = "WEIGHT_VEC"
      
      For i = 1 To NASSETS
            .Offset(i, -1).formula = "=offset(" & CORREL_POS_RNG.Address & _
                ",0," & CStr(i) & ")"
            
            .Offset(i, 0).value = 0
            .Offset(i, 0).Font.ColorIndex = 5
            
            .Offset(i, 1).value = 0
            .Offset(i, 1).Font.ColorIndex = 5
            
            .Offset(i, 2).formula = "=" & .Offset(i, 1).Address & "-" & _
            .Offset(i, 0).Address
      Next i

    With .Offset(-3, -1)
        .value = ("2. PORTFOLIO AND BENCHMARK WEIGHTS")
        .Font.Bold = True
    End With

    With .Offset(-1, -1)
        .value = ("WEIGHTS [%]")
        .Font.Bold = True
    End With
    With .Offset(0, 0)
        .value = "Benchmark"
    End With
    With .Offset(0, 1)
        .value = "Portfolio"
    End With
    With .Offset(0, 2)
        .value = "Active"
    End With
    With .Offset(NASSETS + 1, -1)
        .value = "Total"
        .Font.Bold = True
    End With
    With .Offset(NASSETS + 1, 0)
        .formula = "=SUM(" & WEIGHTS_RNG.Columns(1).Address & ")"
    End With
    With .Offset(NASSETS + 1, 1)
        .formula = "=SUM(" & WEIGHTS_RNG.Columns(2).Address & ")"
    End With
    With .Offset(NASSETS + 1, 2)
        .formula = "=SUM(" & WEIGHTS_RNG.Columns(3).Address & ")"
    End With
End With

Set COVAR_POS_RNG = WEIGHT_POS_RNG.Offset(0, 3 + k)

With COVAR_POS_RNG
     Set COVAR_RNG = Range(.Offset(NASSETS, 0), .Offset(1, 2))
     If ADD_RNG_NAME = True Then: COVAR_RNG.name = "COV_VECT"
      
      For i = 1 To NASSETS
            .Offset(i, -1).formula = "=offset(" & CORREL_POS_RNG.Address & _
                ",0," & CStr(i) & ")"
      Next i

    With .Offset(-1, -1)
        .value = ("COVARIANCES*")
        .Font.Bold = True
    End With
    With .Offset(0, 0)
        .value = "Benchmark"
    End With
    With .Offset(0, 1)
        .value = "Portfolio"
    End With
    With .Offset(0, 2)
        .value = "Active"
    End With
End With

COVAR_RNG.Columns(1).FormulaArray = "=MMULT(" & VAR_COVAR_RNG.Address & _
"," & WEIGHTS_RNG.Columns(1).Address & "/" & FACTOR & ")"
COVAR_RNG.Columns(2).FormulaArray = "=MMULT(" & VAR_COVAR_RNG.Address & _
"," & WEIGHTS_RNG.Columns(2).Address & "/" & FACTOR & ")"
COVAR_RNG.Columns(3).FormulaArray = "=MMULT(" & VAR_COVAR_RNG.Address & _
"," & WEIGHTS_RNG.Columns(3).Address & "/" & FACTOR & ")"

'---------------------------------------------------------------------------
'----------------------THIRD PASS: Absolute And Relative Risk---------
'---------------------------------------------------------------------------

Set ABS_REL_POS_RNG = WEIGHT_POS_RNG.Offset(NASSETS + m + 1, 0)

With ABS_REL_POS_RNG
     Set ABS_REL_RNG = Range(.Offset(0, 0), .Offset(0, 2))
     If ADD_RNG_NAME = True Then: ABS_REL_RNG.name = "ABS_REL_VAR"
      
     ABS_REL_RNG.Cells(0).value = "Volatility"
     
     ABS_REL_RNG.Cells(1).FormulaArray = "=SQRT(MMULT(MMULT(TRANSPOSE(" & _
     WEIGHTS_RNG.Columns(1).Address & "/" & FACTOR & ")," & _
     VAR_COVAR_RNG.Address & ")," & WEIGHTS_RNG.Columns(1).Address & "/" & _
     FACTOR & "))"
     
     ABS_REL_RNG.Cells(2).FormulaArray = "=SQRT(MMULT(MMULT(TRANSPOSE(" & _
     WEIGHTS_RNG.Columns(2).Address & "/" & FACTOR & ")," & _
     VAR_COVAR_RNG.Address & ")," & WEIGHTS_RNG.Columns(2).Address & "/" & _
     FACTOR & "))"
     
     ABS_REL_RNG.Cells(3).FormulaArray = "=SQRT(MMULT(MMULT(TRANSPOSE(" & _
     WEIGHTS_RNG.Columns(3).Address & "/" & FACTOR & ")," & _
     VAR_COVAR_RNG.Address & ")," & WEIGHTS_RNG.Columns(3).Address & "/" & _
     FACTOR & "))"
    
    With .Offset(-3, -1)
        .value = ("3. ABSOLUTE AND RELATIVE RISK")
        .Font.Bold = True
    End With

    With .Offset(-1, 0)
        .value = "Benchmark"
    End With
    With .Offset(-1, 1)
        .value = "Portfolio"
    End With
    With .Offset(-1, 2)
        .value = "Active"
    End With
End With

Set BETA_POS_RNG = COVAR_POS_RNG.Offset(0, 3 + k)

With BETA_POS_RNG
     Set BETA_RNG = Range(.Offset(NASSETS, 0), .Offset(1, 2))
     If ADD_RNG_NAME = True Then: COVAR_RNG.name = "BET_VECT"
      
      For i = 1 To NASSETS
            .Offset(i, -1).formula = "=offset(" & CORREL_POS_RNG.Address & _
                ",0," & CStr(i) & ")"
            .Offset(i, 0).formula = "=" & COVAR_RNG.Columns(1).Cells(i).Address & _
                "/" & ABS_REL_RNG.Cells(1).Address & "^2"
            .Offset(i, 1).formula = "=" & COVAR_RNG.Columns(2).Cells(i).Address & _
                "/" & ABS_REL_RNG.Cells(2).Address & "^2"
            .Offset(i, 2).formula = "=" & COVAR_RNG.Columns(3).Cells(i).Address & _
                "/" & ABS_REL_RNG.Cells(3).Address & "^2"
      Next i

    With .Offset(-1, -1)
        .value = ("BETAS*")
        .Font.Bold = True
    End With
    With .Offset(0, 0)
        .value = "Benchmark"
    End With
    With .Offset(0, 1)
        .value = "Portfolio"
    End With
    With .Offset(0, 2)
        .value = "Active"
    End With
End With

'---------------------------------------------------------------------------
'----------------------FORTH PASS: Marginal Contribution To Risk---------
'---------------------------------------------------------------------------

Set MARG_POS_RNG = ABS_REL_POS_RNG.Offset(-1, 3 + k)

With MARG_POS_RNG
     Set MARG_RNG = Range(.Offset(NASSETS, 0), .Offset(1, 2))
     If ADD_RNG_NAME = True Then: MARG_RNG.name = "MAR_VECT"
      
      For i = 1 To NASSETS
            .Offset(i, -1).formula = "=offset(" & CORREL_POS_RNG.Address & _
                ",0," & CStr(i) & ")"
      Next i

    With .Offset(-2, -1)
        .value = ("4. MARGINAL CONTRIBUTION TO RISK")
        .Font.Bold = True
    End With
    With .Offset(0, 0)
        .value = "Benchmark"
    End With
    With .Offset(0, 1)
        .value = "Portfolio"
    End With
    With .Offset(0, 2)
        .value = "Active"
    End With
End With

MARG_RNG.Columns(1).FormulaArray = "=MMULT(" & VAR_COVAR_RNG.Address & _
"," & WEIGHTS_RNG.Columns(1).Address & "/" & FACTOR & ")/" & _
ABS_REL_RNG.Cells(1).Address & ""

MARG_RNG.Columns(2).FormulaArray = "=MMULT(" & VAR_COVAR_RNG.Address & _
"," & WEIGHTS_RNG.Columns(2).Address & "/" & FACTOR & ")/" & _
ABS_REL_RNG.Cells(2).Address & ""

MARG_RNG.Columns(3).FormulaArray = "=MMULT(" & VAR_COVAR_RNG.Address & _
"," & WEIGHTS_RNG.Columns(3).Address & "/" & FACTOR & ")/" & _
ABS_REL_RNG.Cells(3).Address & ""

'---------------------------------------------------------------------------
'----------------------FIFTH PASS: Contribution To Risk---------
'---------------------------------------------------------------------------

Set CONT_POS_RNG = MARG_POS_RNG.Offset(0, 3 + k)

With CONT_POS_RNG
     Set CONT_RNG = Range(.Offset(NASSETS, 0), .Offset(1, 2))
     If ADD_RNG_NAME = True Then: CONT_RNG.name = "CONT_VECT"
      
      For i = 1 To NASSETS
            .Offset(i, -1).formula = "=offset(" & CORREL_POS_RNG.Address & _
                ",0," & CStr(i) & ")"
            .Offset(i, 0).formula = "=" & MARG_RNG.Columns(1).Cells(i).Address & _
                "*" & WEIGHTS_RNG.Columns(1).Cells(i).Address & "/" & FACTOR
            .Offset(i, 1).formula = "=" & MARG_RNG.Columns(2).Cells(i).Address & _
                "*" & WEIGHTS_RNG.Columns(2).Cells(i).Address & "/" & FACTOR
            .Offset(i, 2).formula = "=" & MARG_RNG.Columns(3).Cells(i).Address & _
                "*" & WEIGHTS_RNG.Columns(3).Cells(i).Address & "/" & FACTOR
      Next i

    With .Offset(-2, -1)
        .value = ("5. CONTRIBUTION TO RISK")
        .Font.Bold = True
    End With
    With .Offset(0, 0)
        .value = "Benchmark"
    End With
    With .Offset(0, 1)
        .value = "Portfolio"
    End With
    With .Offset(0, 2)
        .value = "Active"
    End With
    With .Offset(NASSETS + 1, -1)
        .value = "Total"
        .Font.Bold = True
    End With
    With .Offset(NASSETS + 1, 0)
        .formula = "=SUM(" & CONT_RNG.Columns(1).Address & ")"
    End With
    With .Offset(NASSETS + 1, 1)
        .formula = "=SUM(" & CONT_RNG.Columns(2).Address & ")"
    End With
    With .Offset(NASSETS + 1, 2)
        .formula = "=SUM(" & CONT_RNG.Columns(3).Address & ")"
    End With

End With

'---------------------------------------------------------------------------
'----------------------SIXTH PASS: Percent Contribution To Risk---------
'---------------------------------------------------------------------------

Set PER_POS_RNG = CONT_POS_RNG.Offset(0, 3 + k)

With PER_POS_RNG
     Set PER_RNG = Range(.Offset(NASSETS, 0), .Offset(1, 2))
     If ADD_RNG_NAME = True Then: PER_RNG.name = "PER_VECT"
      
      For i = 1 To NASSETS
            .Offset(i, -1).formula = "=offset(" & CORREL_POS_RNG.Address & _
                ",0," & CStr(i) & ")"
            .Offset(i, 0).formula = "=" & CONT_RNG.Columns(1).Cells(i).Address & _
                "/" & CONT_RNG.Columns(1).Cells(NASSETS).Offset(1, 0).Address & "*" & FACTOR
            .Offset(i, 1).formula = "=" & CONT_RNG.Columns(2).Cells(i).Address & _
                "/" & CONT_RNG.Columns(2).Cells(NASSETS).Offset(1, 0).Address & "*" & FACTOR
            .Offset(i, 2).formula = "=" & CONT_RNG.Columns(3).Cells(i).Address & _
                "/" & CONT_RNG.Columns(3).Cells(NASSETS).Offset(1, 0).Address & "*" & FACTOR
      Next i

    With .Offset(-2, -1)
        .value = ("6. PERCENT CONTRIBUTION TO RISK")
        .Font.Bold = True
    End With
    With .Offset(0, 0)
        .value = "Benchmark"
    End With
    With .Offset(0, 1)
        .value = "Portfolio"
    End With
    With .Offset(0, 2)
        .value = "Active"
    End With
    With .Offset(NASSETS + 1, -1)
        .value = "Total"
        .Font.Bold = True
    End With
    With .Offset(NASSETS + 1, 0)
        .formula = "=SUM(" & PER_RNG.Columns(1).Address & ")"
    End With
    With .Offset(NASSETS + 1, 1)
        .formula = "=SUM(" & PER_RNG.Columns(2).Address & ")"
    End With
    With .Offset(NASSETS + 1, 2)
        .formula = "=SUM(" & PER_RNG.Columns(3).Address & ")"
    End With
End With

'---------------------------------------------------------------------------
'----------------------LAST PASS: Minimum Variance Portfolio Weights--------
'---------------------------------------------------------------------------

Set MIN_POS_RNG = PER_POS_RNG.Offset(0, 3 + k)

With MIN_POS_RNG
     Set MIN_RNG = Range(.Offset(NASSETS, 0), .Offset(1, 0))
     If ADD_RNG_NAME = True Then: MIN_RNG.name = "MIN_VECT"
      
      For i = 1 To NASSETS
            .Offset(i, -1).formula = "=offset(" & CORREL_POS_RNG.Address & _
                ",0," & CStr(i) & ")"
            .Offset(i, 1).formula = "=1"
            .Offset(i, 1).Font.ColorIndex = 5
      Next i

    With .Offset(NASSETS + 3, -1)
        .value = "Minimum Variance"
        .Font.Bold = True
    End With
    With .Offset(NASSETS + 3, 0)
        .FormulaArray = "=1/MMULT(MMULT(TRANSPOSE(" & _
        MIN_RNG.Offset(0, 1).Address & _
        "),MINVERSE(" & VAR_COVAR_RNG.Address & "))," & _
        MIN_RNG.Offset(0, 1).Address & ")"
    End With
    
    MIN_RNG.FormulaArray = "=" & FACTOR & "*MMULT(" & _
    MIN_POS_RNG.Offset(NASSETS + 3, 0).Address & _
    "*MINVERSE(" & VAR_COVAR_RNG.Address & _
    ")," & MIN_RNG.Offset(0, 1).Address & ")"
    
    With .Offset(NASSETS + 4, -1)
        .value = "Minimum Volatility"
        .Font.Bold = True
    End With
    With .Offset(NASSETS + 4, 0)
        .formula = "=" & MIN_POS_RNG.Offset(NASSETS + 3, 0).Address & "^0.5"
    End With

    With .Offset(-2, -1)
        .value = ("7. MINIMUM VARIANCE PORTFOLIO WEIGHTS")
        .Font.Bold = True
    End With
    With .Offset(0, 0)
        .value = "Portfolio"
    End With
    With .Offset(0, 1)
        .value = "Factors"
    End With
    With .Offset(NASSETS + 1, -1)
        .value = "Total"
        .Font.Bold = True
    End With
    With .Offset(NASSETS + 1, 0)
        .formula = "=SUM(" & MIN_RNG.Columns(1).Address & ")"
    End With

End With

RNG_PORT_RISK_CONTRIBUTION_OPTIMISATION_FUNC = True


Exit Function
ERROR_LABEL:
RNG_PORT_RISK_CONTRIBUTION_OPTIMISATION_FUNC = False
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : RNG_PORT_TRIMABILITY_OPTIMISATION_FUNC

'DESCRIPTION   : Long-short Risk Calculations: Risk, return and marginal risk,
'risk contributions for a long-short portfolio.

'REFERENCE: Bruce I. Jacobs, Kenneth N. Levy, Harry M. Markowitz:
'"Trimability and Fast Optimisation of Long-Short Portfolios",
'Financial Analyst Journal, Vol 62, No 2, 2006

'LIBRARY       : PORTFOLIO
'GROUP         : RISK_CONTRIBUTION
'ID            : 002
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/08/2008
'************************************************************************************
'************************************************************************************
'RNG_PORT_ASSET_RISK_CONTRIBUTION_FUNC

Function RNG_PORT_TRIMABILITY_OPTIMISATION_FUNC(ByRef DST_RNG As Excel.Range, _
ByVal NASSETS As Long, _
Optional ByVal ADD_RNG_NAME As Boolean = False)

'---------------------------------------------------------------------
'NAssets = NAssets + Lending Rate (=return on Cash Collateral) --> Cash
'---------------------------------------------------------------------

Dim i As Long
Dim j As Long
Dim m As Long

Dim LONG_POS_RNG As Excel.Range
Dim LONG_RNG As Excel.Range

Dim LONG_RETURN As Excel.Range
Dim LONG_BETA As Excel.Range
Dim LONG_VAR As Excel.Range
Dim LONG_COVAR As Excel.Range

Dim SHORT_POS_RNG As Excel.Range
Dim SHORT_RNG As Excel.Range

Dim SHORT_RETURN As Excel.Range
Dim SHORT_BETA As Excel.Range
Dim SHORT_VAR As Excel.Range
Dim SHORT_COVAR As Excel.Range

Dim BENCH_POS_RNG As Excel.Range
Dim BENCH_RNG As Excel.Range

Dim BENCH_LONG_WEIGHT As Excel.Range
Dim BENCH_SHORT_WEIGHT As Excel.Range
Dim BENCH_CASH_WEIGHT As Excel.Range

Dim BENCH_LONG_PORT As Excel.Range
Dim BENCH_SHORT_PORT As Excel.Range
Dim BENCH_CASH_PORT As Excel.Range

Dim BENCH_LONG_RET As Excel.Range
Dim BENCH_SHORT_RET As Excel.Range
Dim BENCH_CASH_RET As Excel.Range

Dim BENCH_COVAR As Excel.Range

Dim RISK_POS_RNG As Excel.Range
Dim RISK_RNG As Excel.Range

Dim BENCH_MAR_POS_RNG As Excel.Range
Dim BENCH_MAR_RNG As Excel.Range

Dim PORT_MAR_POS_RNG As Excel.Range
Dim PORT_MAR_RNG As Excel.Range

Dim LONG_STR As String
Dim SHORT_STR As String

On Error GoTo ERROR_LABEL

RNG_PORT_TRIMABILITY_OPTIMISATION_FUNC = False

m = 6
LONG_STR = "/LONG"
SHORT_STR = "/SHORT"

If NASSETS < 2 Then: GoTo ERROR_LABEL

'---------------------------------------------------------------------------
'----------------------FIRST PASS: LONG POSITION STATISTICS-----------------
'---------------------------------------------------------------------------

Set LONG_POS_RNG = DST_RNG
LONG_POS_RNG.Offset(-1, 0).value = "LONG POSITION STATISTICS"
LONG_POS_RNG.Offset(-1, 0).Font.Bold = True

With LONG_POS_RNG
   Set LONG_RNG = Range(.Offset(NASSETS, 1), .Offset(1, NASSETS + m))
   If ADD_RNG_NAME = True Then: LONG_RNG.name = "LONG_STAT"
    
    .Offset(0, 1).value = "Expected Return"
    .Offset(0, 1).Font.Bold = True
    
    .Offset(0, 2).value = "Beta"
    .Offset(0, 2).Font.Bold = True
    
    .Offset(0, 3).value = "Idiosyncratic Variance"
    .Offset(0, 3).Font.Bold = True

    For i = 1 To NASSETS
      With .Offset(i, 0)
         .value = "Asset " & CStr(i)
         .Font.ColorIndex = 3
      End With
      With .Offset(i, 1)
          .value = 0
          .Font.ColorIndex = 5
      End With
      With .Offset(i, 2)
          .value = 0
          .Font.ColorIndex = 5
      End With
      With .Offset(i, 3)
          .value = 0
          .Font.ColorIndex = 5
      End With
    Next i
    .Offset(NASSETS + 1, 0).value = "Lending Rate"
    .Offset(NASSETS + 1, 0).Font.Bold = True
    .Offset(NASSETS + 1, 0).AddComment
    .Offset(NASSETS + 1, 0).Comment.Text Text:= _
    "Nicholas Fermin:" & Chr(10) & " Return on Cash Collateral"
    
    .Offset(NASSETS + 1, 1).value = 0.03
    .Offset(NASSETS + 1, 1).Font.ColorIndex = 3
    
    .Offset(NASSETS + 2, 0).value = "Factor Var"
    .Offset(NASSETS + 2, 0).Font.Bold = True

    .Offset(NASSETS + 2, 1).value = 0.04
    .Offset(NASSETS + 2, 1).Font.ColorIndex = 3
End With

LONG_POS_RNG.Cells(0, 1 + m).value = ("COVARIANCE MATRIX LONGS")
LONG_POS_RNG.Cells(0, 1 + m).Font.Bold = True

With LONG_RNG
        For i = 2 To NASSETS
            For j = 1 To i - 1
                With .Cells(i, j + m)
                    .formula = "=offset(" & LONG_POS_RNG.Address & "," & CStr(j) & _
                    ",2)*offset(" & LONG_POS_RNG.Address & "," & CStr(i) & _
                    ",2)*" & LONG_POS_RNG.Offset(NASSETS + 2, 1).Address
                End With
            Next j
        Next i
        
        For i = 1 To NASSETS
                With .Cells(i, i + m)
                    .formula = "=OFFSET(" & LONG_POS_RNG.Address & "," & CStr(i) & _
                    ",3)+Offset(" & LONG_POS_RNG.Address & ", " & CStr(i) & _
                    ",2)^2*" & LONG_POS_RNG.Offset(NASSETS + 2, 1).Address
                End With
        Next i

        For i = 1 To NASSETS - 1
            For j = i + 1 To NASSETS
                With .Cells(i, j + m)
                    .formula = "=offset(" & LONG_POS_RNG.Address & "," & CStr(i) & _
                    ",2)*offset(" & LONG_POS_RNG.Address & "," & CStr(j) & _
                    ",2)*" & LONG_POS_RNG.Offset(NASSETS + 2, 1).Address
                End With
            Next j
        Next i
    For i = 1 To NASSETS
          .Cells(i, m).formula = _
          "=offset(" & LONG_POS_RNG.Address & "," & CStr(i) & ",0)"
          .Cells(0, m + i).formula = _
          "=offset(" & LONG_POS_RNG.Address & "," & CStr(i) & ",0)"
    Next i
End With

Set LONG_RETURN = LONG_RNG.Columns(1)
Set LONG_BETA = LONG_RNG.Columns(2)
Set LONG_VAR = LONG_RNG.Columns(3)
Set LONG_COVAR = Range(LONG_RNG.Columns(m + 1), LONG_RNG.Columns(NASSETS + m))


'---------------------------------------------------------------------------
'----------------------SECOND PASS: SHORT POSITION STATISTICS-----------------
'---------------------------------------------------------------------------

Set SHORT_POS_RNG = LONG_POS_RNG.Offset(NASSETS + m, 0)
SHORT_POS_RNG.Offset(-1, 0).value = "SHORT POSITION STATISTICS"
SHORT_POS_RNG.Offset(-1, 0).Font.Bold = True

With SHORT_POS_RNG
   
   Set SHORT_RNG = Range(.Offset(NASSETS, 1), .Offset(1, NASSETS + m))
   If ADD_RNG_NAME = True Then: SHORT_RNG.name = "SHORT_STAT"
    
    .Offset(0, 1).value = "Expected Return"
    .Offset(0, 1).Font.Bold = True
    
    .Offset(0, 2).value = "Beta"
    .Offset(0, 2).Font.Bold = True
    
    .Offset(0, 3).value = "Idiosyncratic Variance"
    .Offset(0, 3).Font.Bold = True

    .Offset(NASSETS + 1, 0).value = "Borrowing Rate"
    .Offset(NASSETS + 1, 0).Font.Bold = True
    .Offset(NASSETS + 1, 1).value = -0.05
    .Offset(NASSETS + 1, 1).Font.ColorIndex = 3
End With

SHORT_POS_RNG.Cells(0, 1 + m).value = ("COVARIANCE MATRIX SHORTS")
SHORT_POS_RNG.Cells(0, 1 + m).Font.Bold = True

Set SHORT_RETURN = SHORT_RNG.Columns(1)
SHORT_RETURN.FormulaArray = "=-" & LONG_RETURN.Address

Set SHORT_BETA = SHORT_RNG.Columns(2)
SHORT_BETA.FormulaArray = "=-" & LONG_BETA.Address

Set SHORT_VAR = SHORT_RNG.Columns(3)
SHORT_VAR.FormulaArray = "=" & LONG_VAR.Address

Set SHORT_COVAR = Range(SHORT_RNG.Columns(m + 1), SHORT_RNG.Columns(NASSETS + m))
SHORT_COVAR.FormulaArray = "=-" & LONG_COVAR.Address

    For i = 1 To NASSETS
          SHORT_RNG.Cells(i, m).formula = _
          "=offset(" & LONG_POS_RNG.Address & "," & CStr(i) & ",0)"
          
          SHORT_POS_RNG.Cells(i + 1).formula = _
          "=offset(" & LONG_POS_RNG.Address & "," & CStr(i) & ",0)"
          
          SHORT_RNG.Cells(0, m + i).formula = _
          "=offset(" & LONG_POS_RNG.Address & "," & CStr(i) & ",0)"
    Next i

'---------------------------------------------------------------------------
'----------------------THIRD PASS: BENCHMARK STATISTICS-----------------
'---------------------------------------------------------------------------
m = m + 1
Set BENCH_POS_RNG = SHORT_POS_RNG.Offset(NASSETS + m - 1, 0)
BENCH_POS_RNG.Offset(-1, 0).value = "BENCHMARK STATISTICS"
BENCH_POS_RNG.Offset(-1, 0).Font.Bold = True

BENCH_POS_RNG.Offset(-1, m - 1).value = "COVARIANCES LONGS & SHORTS"
BENCH_POS_RNG.Offset(-1, m - 1).Font.Bold = True

With BENCH_POS_RNG
   Set BENCH_RNG = Range(.Offset(NASSETS * 2 + 2, 1), _
   .Offset(1, NASSETS * 2 + m + 1))
   
   If ADD_RNG_NAME = True Then: BENCH_RNG.name = "BENCH_STAT"

    .Offset(0, 1).value = "Benchmark Weights"
    .Offset(0, 1).Font.Bold = True
    
    .Offset(0, 2).value = "Portfolio Weights"
    .Offset(0, 2).Font.Bold = True
    
    .Offset(0, 3).value = "Returns"
    .Offset(0, 3).Font.Bold = True

    For i = 1 To NASSETS
      With .Offset(i, 0)
         .formula = "=offset(" & LONG_POS_RNG.Address & "," & CStr(i) & ",0)" & _
         "&"" " & LONG_STR & """"
      End With
      
      With .Offset(i + NASSETS, 0)
         .formula = "=offset(" & LONG_POS_RNG.Address & "," & CStr(i) & ",0)" & _
         "&"" " & SHORT_STR & """"
      End With
      
      With .Offset(i, 1)
          .value = 0
          .Font.ColorIndex = 5
      End With
      With .Offset(i + NASSETS, 1)
          .value = 0
          .Font.ColorIndex = 5
      End With
      
      With .Offset(i, 2)
          .value = 0
          .Font.ColorIndex = 5
      End With
      With .Offset(i + NASSETS, 2)
          .value = 0
          .Font.ColorIndex = 5
      End With
      
    Next i

    .Offset(2 * NASSETS + 1, 0) = "Cash " & LONG_STR
    .Offset(2 * NASSETS + 1, 3).formula = "=" & _
        LONG_POS_RNG.Offset(NASSETS + 1, 1).Address
      
      With .Offset(2 * NASSETS + 1, 1)
          .value = 0
          .Font.ColorIndex = 5
      End With
      With .Offset(2 * NASSETS + 1, 2)
          .value = 0
          .Font.ColorIndex = 5
      End With
    
     .Offset(2 * NASSETS + 2, 0) = "Cash " & SHORT_STR
     .Offset(2 * NASSETS + 2, 3).formula = "=" & _
        SHORT_POS_RNG.Offset(NASSETS + 1, 1).Address
      
      With .Offset(2 * NASSETS + 2, 1)
          .value = 0
          .Font.ColorIndex = 5
      End With
      With .Offset(2 * NASSETS + 2, 2)
          .value = 0
          .Font.ColorIndex = 5
      End With

'UPPER QUADRANTS ---------------------------------------------------------
      Range(.Offset(1, 3), .Offset(NASSETS, 3)).FormulaArray = "=" & _
        LONG_RETURN.Address
      
      Range(.Offset(1, m), .Offset(NASSETS, m + NASSETS - _
        1)).FormulaArray = "=" & LONG_COVAR.Address
      
      Range(.Offset(1, m + NASSETS), .Offset(NASSETS, m + NASSETS * 2 _
        - 1)).FormulaArray = "=" & SHORT_COVAR.Address

      Range(.Offset(1, m + NASSETS * 2), .Offset(NASSETS * 2, m + _
      NASSETS * 2 + 1)).value = 0
      
      Range(.Offset(1, m + NASSETS * 2), .Offset(NASSETS * 2, m + _
      NASSETS * 2 + 1)).Font.ColorIndex = 5


'LOWER QUADRANTS ----------------------------------------------------------
      Range(.Offset(NASSETS + 1, 3), .Offset(NASSETS * 2, _
        3)).FormulaArray = "=" & SHORT_RETURN.Address

      Range(.Offset(NASSETS + 1, m), .Offset(NASSETS * 2, m + _
        NASSETS - 1)).FormulaArray = "=" & SHORT_COVAR.Address
      
      Range(.Offset(NASSETS + 1, m + NASSETS), _
      .Offset(NASSETS * 2, m + NASSETS * 2 - 1)).FormulaArray = "=" & _
        LONG_COVAR.Address

      Range(.Offset(NASSETS * 2 + 1, m), .Offset(NASSETS * 2 + 2, _
      m + NASSETS * 2 + 1)).value = 0
      
      Range(.Offset(NASSETS * 2 + 1, m), .Offset(NASSETS * 2 + 2, _
      m + NASSETS * 2 + 1)).Font.ColorIndex = 5

   For i = 1 To NASSETS * 2 + 2
      With .Offset(i, m - 1)
         .formula = "=offset(" & BENCH_POS_RNG.Address & "," & CStr(i) & ",0)"
      End With
      With .Offset(0, m - 1 + i)
         .formula = "=offset(" & BENCH_POS_RNG.Address & "," & CStr(i) & ",0)"
      End With
   Next i
End With

With BENCH_RNG
    With .Columns(1)
        Set BENCH_LONG_WEIGHT = Range(.Cells(1, 1), _
            .Cells(NASSETS, 1))
        Set BENCH_SHORT_WEIGHT = Range(.Cells(NASSETS + 1, 1), _
            .Cells(NASSETS * 2, 1))
        Set BENCH_CASH_WEIGHT = Range(.Cells(NASSETS * 2 + 1, 1), _
            .Cells(NASSETS * 2 + 2, 1))
    End With
        
    With .Columns(2)
        Set BENCH_LONG_PORT = Range(.Cells(1, 1), _
            .Cells(NASSETS, 1))
        Set BENCH_SHORT_PORT = Range(.Cells(NASSETS + 1, 1), _
            .Cells(NASSETS * 2, 1))
        Set BENCH_CASH_PORT = Range(.Cells(NASSETS * 2 + 1, 1), _
            .Cells(NASSETS * 2 + 2, 1))
    End With
    
    With .Columns(3)
        Set BENCH_LONG_RET = Range(.Cells(1, 1), _
            .Cells(NASSETS, 1))
        Set BENCH_SHORT_RET = Range(.Cells(NASSETS + 1, 1), _
            .Cells(NASSETS * 2, 1))
        Set BENCH_CASH_RET = Range(.Cells(NASSETS * 2 + 1, 1), _
            .Cells(NASSETS * 2 + 2, 1))
    End With
        Set BENCH_COVAR = Range(.Columns(m), .Columns(m + NASSETS * 2 + 1))
End With


'---------------------------------------------------------------------------
'----------------------FORTH PASS: RISK_CONTRIBUTION -----------------
'---------------------------------------------------------------------------
Set RISK_POS_RNG = BENCH_POS_RNG.Offset(NASSETS * 2 + m, 0)
RISK_POS_RNG.Offset(-1, 0).value = "RISK CONTRIBUTION"
RISK_POS_RNG.Offset(-1, 0).Font.Bold = True

With RISK_POS_RNG
   Set RISK_RNG = Range(.Offset(2, 1), .Offset(12, 2))
   If ADD_RNG_NAME = True Then: RISK_RNG.name = "RISK_CONT_SUMM"

   .Offset(1, 1).value = "Benchmark"
   .Offset(1, 1).Font.Bold = True
   .Offset(1, 2).value = "Portfolio"
   .Offset(1, 2).Font.Bold = True
   
   .Offset(2, 0).value = ("EXPOSURE (GROSS)")
   
   .Offset(2, 1).formula = "=SUM(" & BENCH_LONG_WEIGHT.Address & ")+SUM(" & _
    BENCH_SHORT_WEIGHT.Address & ")+SUM(" & BENCH_CASH_WEIGHT.Address & ")"
   
   .Offset(2, 2).formula = "=SUM(" & BENCH_LONG_PORT.Address & ")+SUM(" & _
    BENCH_SHORT_PORT.Address & ")+SUM(" & BENCH_CASH_PORT.Address & ")"
   
   .Offset(3, 0).value = ("EXPOSURE (NET)")
   .Offset(3, 1).formula = "=SUM(" & BENCH_LONG_WEIGHT.Address & ")-SUM(" & _
    BENCH_SHORT_WEIGHT.Address & ")+" & BENCH_CASH_WEIGHT.Cells(1).Address & "-" & _
    BENCH_CASH_WEIGHT.Cells(2).Address
   
   .Offset(3, 2).formula = "=SUM(" & BENCH_LONG_PORT.Address & ")-SUM(" & _
    BENCH_SHORT_PORT.Address & ")+" & BENCH_CASH_PORT.Cells(1).Address & "-" & _
    BENCH_CASH_PORT.Cells(2).Address
    
   .Offset(4, 0).value = ("RETURN")
   .Offset(4, 1).FormulaArray = "=MMULT(TRANSPOSE(" & _
        BENCH_RNG.Columns(1).Address & ")," & BENCH_RNG.Columns(3).Address & ")"
   .Offset(4, 2).FormulaArray = _
        "=MMULT(TRANSPOSE(" & BENCH_RNG.Columns(2).Address & _
        ")," & BENCH_RNG.Columns(3).Address & ")"

   .Offset(5, 0).value = ("VOLATILITY")
   .Offset(5, 1).FormulaArray = _
        "=SQRT(MMULT(MMULT(TRANSPOSE(" & BENCH_RNG.Columns(1).Address & ")," & _
        BENCH_COVAR.Address & ")," & BENCH_RNG.Columns(1).Address & "))"
   .Offset(5, 2).FormulaArray = _
        "=SQRT(MMULT(MMULT(TRANSPOSE(" & BENCH_RNG.Columns(2).Address & ")," & _
        BENCH_COVAR.Address & ")," & BENCH_RNG.Columns(2).Address & "))"
   
   .Offset(6, 0).value = ("COVARIANCE")
   .Offset(6, 1).FormulaArray = "=MMULT(" & BENCH_COVAR.Address & _
        "," & BENCH_RNG.Columns(1).Address & ")"
   .Offset(6, 2).FormulaArray = "=MMULT(" & BENCH_COVAR.Address & _
        "," & BENCH_RNG.Columns(2).Address & ")"
   
   .Offset(7, 0).value = ("BETA")
   .Offset(7, 1).formula = "=" & .Offset(6, 1).Address & "/" & _
        .Offset(5, 1).Address & "^2"
   .Offset(7, 2).formula = "=" & .Offset(6, 2).Address & "/" & _
        .Offset(5, 2).Address & "^2"

'------------------------------------------------------------------------------
   .Offset(9, 0).value = ("VAR CONFIDENCE LEVEL")
   .Offset(9, 1).value = 0.99
   .Offset(9, 1).Font.ColorIndex = 3
   
   .Offset(10, 0).value = ("VAR HORIZON [DAYS PER BUSINESS YEAR]")
   .Offset(10, 1).value = 30
   .Offset(10, 1).Font.ColorIndex = 3
   
   .Offset(11, 0).value = ("ABSOLUTE VAR") '--> *absolute VaR, i.e. the _
   (negative) deviation from the mean return.
   .Offset(11, 1).formula = "=" & .Offset(4, 1).Address & "-" & _
        .Offset(5, 1).Address & "*NORMSINV(" & _
        .Offset(9, 1).Address & ")*SQRT(" & .Offset(10, 1).Address & "/252)"
   .Offset(11, 2).formula = "=" & .Offset(4, 2).Address & "-" & _
        .Offset(5, 2).Address & "*NORMSINV(" & _
        .Offset(9, 1).Address & ")*SQRT(" & .Offset(10, 1).Address & "/252)"
   .Offset(12, 0).value = ("PROBABILITY OF DEFAULT")
   '**probability of losses higher than the capital contributed,
   'i.e. returns < 100%
   .Offset(12, 1).formula = "=NORMDIST(-1," & .Offset(4, 1).Address & "," & _
        .Offset(5, 1).Address & ",TRUE)"
   .Offset(12, 2).formula = "=NORMDIST(-1," & .Offset(4, 2).Address & "," & _
        .Offset(5, 2).Address & ",TRUE)"
End With


'---------------------------------------------------------------------------
'---------------------FIFTH PASS: RISK CONT. SUMMARY------------------------
'---------------------------------------------------------------------------

'------------------------------------------------------------------------
Set BENCH_MAR_POS_RNG = BENCH_POS_RNG.Offset(NASSETS * 2 + m, m - 1)
BENCH_MAR_POS_RNG.Offset(-1, 0).value = "RISK SUMMARY"
BENCH_MAR_POS_RNG.Offset(-1, 0).Font.Bold = True

   Set BENCH_MAR_RNG = Range(BENCH_MAR_POS_RNG.Offset(2, 1), _
   BENCH_MAR_POS_RNG.Offset(NASSETS * 2 + 2 + 1, 3))
   If ADD_RNG_NAME = True Then: BENCH_MAR_RNG.name = "BENCH_RISK_CONT"

BENCH_MAR_POS_RNG.Offset(1, 0).value = ("BENCHMARK")
BENCH_MAR_POS_RNG.Offset(1, 0).Font.Bold = True
   For i = 1 To NASSETS * 2 + 2
      With BENCH_MAR_POS_RNG.Offset(i + 1, 0)
         .formula = "=offset(" & BENCH_POS_RNG.Address & "," & CStr(i) & ",0)"
      End With
   Next i

BENCH_MAR_POS_RNG.Offset(1, 1).value = "Marginal Contr To Risk"
BENCH_MAR_POS_RNG.Offset(1, 1).Font.Bold = True

Range(BENCH_MAR_POS_RNG.Offset(2, 1), _
BENCH_MAR_POS_RNG.Offset(NASSETS * 2 + 2 + 1, 1)).FormulaArray = _
    "=MMULT(" & BENCH_COVAR.Address & _
        "," & BENCH_RNG.Columns(1).Address & ")/" & _
        RISK_POS_RNG.Offset(5, 1).Address

BENCH_MAR_POS_RNG.Offset(1, 2).value = "Contr To Risk"
BENCH_MAR_POS_RNG.Offset(1, 2).Font.Bold = True

Range(BENCH_MAR_POS_RNG.Offset(2, 2), _
BENCH_MAR_POS_RNG.Offset(NASSETS * 2 + 2 + 1, 2)).FormulaArray = _
    "=" & BENCH_RNG.Columns(1).Address & _
        "*" & Range(BENCH_MAR_POS_RNG.Offset(2, 1), _
        BENCH_MAR_POS_RNG.Offset(NASSETS * 2 + 2 + 1, 1)).Address

BENCH_MAR_POS_RNG.Offset(1, 3).value = "Percent Contr To Risk"
BENCH_MAR_POS_RNG.Offset(1, 3).Font.Bold = True

Range(BENCH_MAR_POS_RNG.Offset(2, 3), _
BENCH_MAR_POS_RNG.Offset(NASSETS * 2 + 2 + 1, 3)).FormulaArray = _
    "=" & Range(BENCH_MAR_POS_RNG.Offset(2, 2), _
        BENCH_MAR_POS_RNG.Offset(NASSETS * 2 + 2 + 1, 2)).Address & _
        "/SUM(" & Range(BENCH_MAR_POS_RNG.Offset(2, 2), _
        BENCH_MAR_POS_RNG.Offset(NASSETS * 2 + 2 + 1, 2)).Address & ")"

'------------------------------------------------------------------------
Set PORT_MAR_POS_RNG = BENCH_MAR_POS_RNG.Offset(NASSETS * 2 + m - 1, 0)
   Set PORT_MAR_RNG = Range(PORT_MAR_POS_RNG.Offset(2, 1), _
   PORT_MAR_POS_RNG.Offset(NASSETS * 2 + 2 + 1, 3))
   If ADD_RNG_NAME = True Then: PORT_MAR_RNG.name = "PORT_RISK_CONT"

PORT_MAR_POS_RNG.Offset(1, 0).value = ("PORTFOLIO")
PORT_MAR_POS_RNG.Offset(1, 0).Font.Bold = True
   For i = 1 To NASSETS * 2 + 2
      With PORT_MAR_POS_RNG.Offset(i + 1, 0)
         .formula = "=offset(" & BENCH_POS_RNG.Address & "," & CStr(i) & ",0)"
      End With
   Next i

PORT_MAR_POS_RNG.Offset(1, 1).value = "Marginal Contr To Risk"
PORT_MAR_POS_RNG.Offset(1, 1).Font.Bold = True

Range(PORT_MAR_POS_RNG.Offset(2, 1), _
PORT_MAR_POS_RNG.Offset(NASSETS * 2 + 2 + 1, 1)).FormulaArray = _
    "=MMULT(" & BENCH_COVAR.Address & _
        "," & BENCH_RNG.Columns(2).Address & ")/" & _
        RISK_POS_RNG.Offset(5, 2).Address

PORT_MAR_POS_RNG.Offset(1, 2).value = "Contr To Risk"
PORT_MAR_POS_RNG.Offset(1, 2).Font.Bold = True

Range(PORT_MAR_POS_RNG.Offset(2, 2), _
PORT_MAR_POS_RNG.Offset(NASSETS * 2 + 2 + 1, 2)).FormulaArray = _
    "=" & BENCH_RNG.Columns(2).Address & _
        "*" & Range(PORT_MAR_POS_RNG.Offset(2, 1), _
        PORT_MAR_POS_RNG.Offset(NASSETS * 2 + 2 + 1, 1)).Address

PORT_MAR_POS_RNG.Offset(1, 3).value = "Percent Contr To Risk"
PORT_MAR_POS_RNG.Offset(1, 3).Font.Bold = True

Range(PORT_MAR_POS_RNG.Offset(2, 3), _
PORT_MAR_POS_RNG.Offset(NASSETS * 2 + 2 + 1, 3)).FormulaArray = _
    "=" & Range(PORT_MAR_POS_RNG.Offset(2, 2), _
        PORT_MAR_POS_RNG.Offset(NASSETS * 2 + 2 + 1, 2)).Address & _
        "/SUM(" & Range(PORT_MAR_POS_RNG.Offset(2, 2), _
        PORT_MAR_POS_RNG.Offset(NASSETS * 2 + 2 + 1, 2)).Address & ")"
        

RNG_PORT_TRIMABILITY_OPTIMISATION_FUNC = True

Exit Function
ERROR_LABEL:
RNG_PORT_TRIMABILITY_OPTIMISATION_FUNC = False
End Function


