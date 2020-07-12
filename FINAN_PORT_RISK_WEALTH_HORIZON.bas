Attribute VB_Name = "FINAN_PORT_RISK_WEALTH_HORIZON"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.

'Assuming a portfolio return that is normally distributed, various
'calculations related to the evolution and variation of the investment
'horizon are performed.

Function WEALTH_TIME_INVESTMENT_HORIZON_FUNC( _
ByRef EXPECTED_RETURN_RNG As Variant, _
ByRef VOLATILITY_RNG As Variant, _
ByRef INITIAL_WEALTH_RNG As Variant, _
ByRef INVESTMENT_HORIZON_RNG As Variant, _
Optional ByRef TERMINAL_LOSS_RNG As Variant = 1, _
Optional ByRef SHORTFALL_RETURN_RNG As Variant = 0.975, _
Optional ByVal CONFIDENCE_VAL As Double = 0.975, _
Optional ByVal DELTA_VAL As Double = 0.1, _
Optional ByVal NO_PERIODS As Long = 100, _
Optional ByVal OUTPUT As Integer = 0)

'EXPECTED_RETURN_RNG: Expected continuous return
'VOLATILITY_RNG: Volatiliy
'INITIAL_WEALTH_RNG: Initial Wealth
'INVESTMENT_HORIZON_RNG: Investment Horizon

'DELTA_VAL --> Prob...

Dim i As Long
Dim j As Long

Dim NROWS As Long
Dim NCOLUMNS As Long

Dim TEMP_VAL As Double
Dim CUMUL_VAL As Double
Dim PROB_VAL As Double

Dim INITIAL_WEALTH_VECTOR As Variant
Dim EXPECTED_RETURN_VECTOR As Variant
Dim INVESTMENT_HORIZON_VECTOR As Variant
Dim VOLATILITY_VECTOR As Variant
Dim TERMINAL_LOSS_VECTOR As Variant 'Expressed in % of orignal Investment Horizon
Dim SHORTFALL_RETURN_VECTOR As Variant

Dim epsilon As Double

'Normal distributed asset returns defined by the following parameters...
'...result in lognormal distributed terminal wealth.

On Error GoTo ERROR_LABEL

epsilon = 10 ^ -15

'-------------------------------------------------------------------------------------
Select Case OUTPUT
'-------------------------------------------------------------------------------------
Case 0
'-------------------------------------------------------------------------------------
'Distributional Characteristics of Terminal Wealth
    If IsArray(EXPECTED_RETURN_RNG) = True Then
        EXPECTED_RETURN_VECTOR = EXPECTED_RETURN_RNG
        If UBound(EXPECTED_RETURN_VECTOR, 1) = 1 Then
            EXPECTED_RETURN_VECTOR = MATRIX_TRANSPOSE_FUNC(EXPECTED_RETURN_VECTOR)
        End If
    Else
        ReDim EXPECTED_RETURN_VECTOR(1 To 1, 1 To 1)
        EXPECTED_RETURN_VECTOR(1, 1) = EXPECTED_RETURN_RNG
    End If
    NROWS = UBound(EXPECTED_RETURN_VECTOR, 1)
    
    If IsArray(VOLATILITY_RNG) = True Then
        VOLATILITY_VECTOR = VOLATILITY_RNG
        If UBound(VOLATILITY_VECTOR, 1) = 1 Then
            VOLATILITY_VECTOR = MATRIX_TRANSPOSE_FUNC(VOLATILITY_VECTOR)
        End If
    Else
        ReDim VOLATILITY_VECTOR(1 To 1, 1 To 1)
        VOLATILITY_VECTOR(1, 1) = VOLATILITY_RNG
    End If
    If NROWS <> UBound(VOLATILITY_VECTOR, 1) Then: GoTo ERROR_LABEL
    
    If IsArray(INITIAL_WEALTH_RNG) = True Then
        INITIAL_WEALTH_VECTOR = INITIAL_WEALTH_RNG
        If UBound(INITIAL_WEALTH_VECTOR, 1) = 1 Then
            INITIAL_WEALTH_VECTOR = MATRIX_TRANSPOSE_FUNC(INITIAL_WEALTH_VECTOR)
        End If
        If NROWS <> UBound(INITIAL_WEALTH_VECTOR, 1) Then: GoTo ERROR_LABEL
    Else
        ReDim INITIAL_WEALTH_VECTOR(1 To NROWS, 1 To 1)
        For i = 1 To NROWS
            INITIAL_WEALTH_VECTOR(i, 1) = INITIAL_WEALTH_RNG
        Next i
    End If
    
    If IsArray(INVESTMENT_HORIZON_RNG) = True Then
        INVESTMENT_HORIZON_VECTOR = INVESTMENT_HORIZON_RNG
        If UBound(INVESTMENT_HORIZON_VECTOR, 1) = 1 Then
            INVESTMENT_HORIZON_VECTOR = MATRIX_TRANSPOSE_FUNC(INVESTMENT_HORIZON_VECTOR)
        End If
        If NROWS <> UBound(INVESTMENT_HORIZON_VECTOR, 1) Then: GoTo ERROR_LABEL
    Else
        ReDim INVESTMENT_HORIZON_VECTOR(1 To NROWS, 1 To 1)
        For i = 1 To NROWS
            INVESTMENT_HORIZON_VECTOR(i, 1) = INVESTMENT_HORIZON_RNG
        Next i
    End If
    
    ReDim TEMP_MATRIX(0 To NROWS, 1 To 10)
    TEMP_MATRIX(0, 1) = "INVESTMENT HORIZON" 'Periods
    TEMP_MATRIX(0, 2) = "INITIAL WEALTH"
    TEMP_MATRIX(0, 3) = "EXPECTED CONTINUOUS RETURN"
    TEMP_MATRIX(0, 4) = "VOLATILITY"
    TEMP_MATRIX(0, 5) = "MEAN WEALTH"
    TEMP_MATRIX(0, 6) = "VARIANCE"
    TEMP_MATRIX(0, 7) = "VOLATILITY"
    TEMP_MATRIX(0, 8) = "MEDIAN WEALTH"
    TEMP_MATRIX(0, 9) = "MODE WEALTH"
    TEMP_MATRIX(0, 10) = "SKEWNESS"
    For i = 1 To NROWS
        TEMP_MATRIX(i, 1) = INVESTMENT_HORIZON_VECTOR(i, 1)
        TEMP_MATRIX(i, 2) = INITIAL_WEALTH_VECTOR(i, 1)
        TEMP_MATRIX(i, 3) = EXPECTED_RETURN_VECTOR(i, 1)
        TEMP_MATRIX(i, 4) = VOLATILITY_VECTOR(i, 1)
        
        TEMP_MATRIX(i, 5) = TEMP_MATRIX(i, 2) * Exp((TEMP_MATRIX(i, 3) + _
                            0.5 * TEMP_MATRIX(i, 4) ^ 2) * _
                            INVESTMENT_HORIZON_VECTOR(i, 1))
        
        TEMP_MATRIX(i, 6) = TEMP_MATRIX(i, 2) ^ 2 * Exp((2 * TEMP_MATRIX(i, 3) + _
                            TEMP_MATRIX(i, 4) ^ 2) * INVESTMENT_HORIZON_VECTOR(i, 1)) _
                            * (Exp((TEMP_MATRIX(i, 4) ^ 2) * _
                            INVESTMENT_HORIZON_VECTOR(i, 1)) - 1)
                            
        TEMP_MATRIX(i, 7) = Sqr(TEMP_MATRIX(i, 6))
        
        TEMP_MATRIX(i, 8) = TEMP_MATRIX(i, 2) * Exp(TEMP_MATRIX(i, 3) * _
                            INVESTMENT_HORIZON_VECTOR(i, 1))
        
        TEMP_MATRIX(i, 9) = TEMP_MATRIX(i, 2) * Exp((TEMP_MATRIX(i, 3) - _
                            TEMP_MATRIX(i, 4) ^ 2) * INVESTMENT_HORIZON_VECTOR(i, 1))
        
        TEMP_MATRIX(i, 10) = (Exp(INVESTMENT_HORIZON_VECTOR(i, 1) * _
                            TEMP_MATRIX(i, 4) ^ 2) + 2) * _
                            Sqr(Exp(INVESTMENT_HORIZON_VECTOR(i, 1) * _
                            TEMP_MATRIX(i, 4) ^ 2) - 1)
    Next i
'-------------------------------------------------------------------------------------
Case 1
'-------------------------------------------------------------------------------------
    If IsArray(TERMINAL_LOSS_RNG) = True Then
        TERMINAL_LOSS_VECTOR = TERMINAL_LOSS_RNG
        If UBound(TERMINAL_LOSS_VECTOR, 1) = 1 Then
            TERMINAL_LOSS_VECTOR = MATRIX_TRANSPOSE_FUNC(TERMINAL_LOSS_VECTOR)
        End If
    Else
        ReDim TERMINAL_LOSS_VECTOR(1 To 1, 1 To 1)
        TERMINAL_LOSS_VECTOR(1, 1) = TERMINAL_LOSS_RNG
    End If
    NCOLUMNS = UBound(TERMINAL_LOSS_VECTOR, 1)
        
    DELTA_VAL = DELTA_VAL / 100
    NROWS = Int((1 - DELTA_VAL) / DELTA_VAL) + 1
    ReDim TEMP_MATRIX(0 To NROWS, 1 To 3 + NCOLUMNS + 2)
    TEMP_MATRIX(0, 1) = "CUM PROB"
    TEMP_MATRIX(0, 2) = "WEALTH"
    TEMP_MATRIX(0, 3) = "CUM PROB*"
    For j = 1 To NCOLUMNS
        TEMP_MATRIX(0, 3 + j) = "TL: " & Format(TERMINAL_LOSS_VECTOR(j, 1), "0.00%")
    Next j
    TEMP_MATRIX(0, 3 + NCOLUMNS + 1) = "CUM PROB"
    TEMP_MATRIX(0, 3 + NCOLUMNS + 2) = "'100%"
    
    CUMUL_VAL = DELTA_VAL
    For i = 1 To NROWS
        If CUMUL_VAL >= 1 Then: CUMUL_VAL = 1 - epsilon
        TEMP_MATRIX(i, 1) = CUMUL_VAL
        
        TEMP_MATRIX(i, 2) = INITIAL_WEALTH_RNG * _
                            Exp(NORMSINV_FUNC(TEMP_MATRIX(i, 1), _
                            (EXPECTED_RETURN_RNG - 1 * VOLATILITY_RNG ^ 2) * _
                            INVESTMENT_HORIZON_RNG, VOLATILITY_RNG * _
                            Sqr(INVESTMENT_HORIZON_RNG), 0))

        TEMP_MATRIX(i, 3) = NORMDIST_FUNC(Log(TEMP_MATRIX(i, 2) / _
                            INITIAL_WEALTH_RNG), (EXPECTED_RETURN_RNG - 1 * _
                            VOLATILITY_RNG ^ 2) * INVESTMENT_HORIZON_RNG, _
                            VOLATILITY_RNG * Sqr(INVESTMENT_HORIZON_RNG), 0)
        CUMUL_VAL = CUMUL_VAL + DELTA_VAL
        
        If (TEMP_MATRIX(i, 2) < INITIAL_WEALTH_RNG) Then
            TEMP_MATRIX(i, 3 + NCOLUMNS + 1) = TEMP_MATRIX(i, 1)
            TEMP_MATRIX(i, 3 + NCOLUMNS + 2) = TEMP_MATRIX(i, 2)
            For j = 1 To NCOLUMNS
                TEMP_MATRIX(i, 3 + j) = INITIAL_WEALTH_RNG * _
                                Exp(NORMSINV_FUNC(TEMP_MATRIX(i, 1), _
                                (EXPECTED_RETURN_RNG - 1 * VOLATILITY_RNG ^ 2) * _
                                INVESTMENT_HORIZON_RNG * TERMINAL_LOSS_VECTOR(j, 1), VOLATILITY_RNG * _
                                Sqr(INVESTMENT_HORIZON_RNG * TERMINAL_LOSS_VECTOR(j, 1)), 0))
            Next j
        Else
            TEMP_MATRIX(i, 3 + NCOLUMNS + 1) = CVErr(xlErrNA)
            TEMP_MATRIX(i, 3 + NCOLUMNS + 2) = CVErr(xlErrNA)
            For j = 1 To NCOLUMNS: TEMP_MATRIX(i, 3 + j) = CVErr(xlErrNA): Next j
        End If
    Next i
'-------------------------------------------------------------------------------------
Case Else
'-------------------------------------------------------------------------------------
    NROWS = NO_PERIODS + 1
    If IsArray(SHORTFALL_RETURN_RNG) = True Then
        SHORTFALL_RETURN_VECTOR = SHORTFALL_RETURN_RNG
        If UBound(SHORTFALL_RETURN_VECTOR, 1) = 1 Then
            SHORTFALL_RETURN_VECTOR = MATRIX_TRANSPOSE_FUNC(SHORTFALL_RETURN_VECTOR)
        End If
    Else
        ReDim SHORTFALL_RETURN_VECTOR(1 To 1, 1 To 1)
        SHORTFALL_RETURN_VECTOR(1, 1) = SHORTFALL_RETURN_RNG
    End If
    NCOLUMNS = UBound(SHORTFALL_RETURN_VECTOR, 1)
    
    ReDim TEMP_MATRIX(0 To NROWS, 1 To 15 + NCOLUMNS)

    TEMP_MATRIX(0, 1) = "i"
    TEMP_MATRIX(0, 2) = "dT"
    TEMP_MATRIX(0, 3) = "$ Wealth: Expected Wealth"
    TEMP_MATRIX(0, 4) = "$ Wealth: Lower " & Format(CONFIDENCE_VAL, "0.00%") & " Confidence"
    TEMP_MATRIX(0, 5) = "$ Wealth: Upper " & Format(CONFIDENCE_VAL, "0.00%") & " Confidence"
    TEMP_MATRIX(0, 6) = "$ Wealth: Modus"
    TEMP_MATRIX(0, 7) = "$ Wealth: Median"
    TEMP_MATRIX(0, 8) = "Period Returns (returns in dT): Expected Period Return"
    TEMP_MATRIX(0, 9) = "Period Returns (returns in dT): Lower " & Format(CONFIDENCE_VAL, "0.00%") & " Confidence"
    TEMP_MATRIX(0, 10) = "Period Returns (returns in dT): Upper " & Format(CONFIDENCE_VAL, "0.00%") & " Confidence"
    TEMP_MATRIX(0, 11) = "Period Returns (returns in dT): Modus"
    TEMP_MATRIX(0, 12) = "Period Returns (returns in dT): Median"
    TEMP_MATRIX(0, 13) = "Average Return: Average Return"
    TEMP_MATRIX(0, 14) = "Average Return: Lower " & Format(CONFIDENCE_VAL, "0.00%") & " Confidence"
    TEMP_MATRIX(0, 15) = "Average Return: Upper " & Format(CONFIDENCE_VAL, "0.00%") & " Confidence"
    For j = 1 To NCOLUMNS
        TEMP_MATRIX(0, 15 + j) = "Shortfall Risk: Minimum Return " & _
                                 Format(SHORTFALL_RETURN_VECTOR(j, 1), "0.00%")
    Next j
    
    For i = 1 To NROWS
        PROB_VAL = NORMSINV_FUNC(CONFIDENCE_VAL, 0, 1, 0)
        If i > 1 Then
            TEMP_MATRIX(i, 1) = TEMP_MATRIX(i - 1, 1) + 1
        Else
            TEMP_MATRIX(i, 1) = 0
        End If
        TEMP_MATRIX(i, 2) = (i - 1) * INVESTMENT_HORIZON_RNG / NO_PERIODS
        TEMP_MATRIX(i, 3) = INITIAL_WEALTH_RNG * Exp((EXPECTED_RETURN_RNG + 0.5 * VOLATILITY_RNG ^ 2) * TEMP_MATRIX(i, 2))
        TEMP_MATRIX(i, 4) = INITIAL_WEALTH_RNG * Exp(TEMP_MATRIX(i, 2) * EXPECTED_RETURN_RNG - PROB_VAL * VOLATILITY_RNG * Sqr(TEMP_MATRIX(i, 2)))
        TEMP_MATRIX(i, 5) = INITIAL_WEALTH_RNG * Exp(TEMP_MATRIX(i, 2) * EXPECTED_RETURN_RNG + PROB_VAL * VOLATILITY_RNG * Sqr(TEMP_MATRIX(i, 2)))
        TEMP_MATRIX(i, 6) = INITIAL_WEALTH_RNG * Exp((EXPECTED_RETURN_RNG - VOLATILITY_RNG ^ 2) * TEMP_MATRIX(i, 2))
        TEMP_MATRIX(i, 7) = INITIAL_WEALTH_RNG * Exp(EXPECTED_RETURN_RNG * TEMP_MATRIX(i, 2))
        TEMP_MATRIX(i, 8) = (EXPECTED_RETURN_RNG + 0.5 * VOLATILITY_RNG ^ 2) * TEMP_MATRIX(i, 2)
        TEMP_MATRIX(i, 9) = TEMP_MATRIX(i, 2) * EXPECTED_RETURN_RNG - PROB_VAL * VOLATILITY_RNG * Sqr(TEMP_MATRIX(i, 2))
        TEMP_MATRIX(i, 10) = TEMP_MATRIX(i, 2) * EXPECTED_RETURN_RNG + PROB_VAL * VOLATILITY_RNG * Sqr(TEMP_MATRIX(i, 2))
        TEMP_MATRIX(i, 11) = (EXPECTED_RETURN_RNG - VOLATILITY_RNG ^ 2) * TEMP_MATRIX(i, 2)
        TEMP_MATRIX(i, 12) = EXPECTED_RETURN_RNG * TEMP_MATRIX(i, 2)
        If i > 1 Then
            TEMP_MATRIX(i, 13) = EXPECTED_RETURN_RNG
            TEMP_MATRIX(i, 14) = EXPECTED_RETURN_RNG - PROB_VAL * VOLATILITY_RNG / Sqr(TEMP_MATRIX(i, 2))
            TEMP_MATRIX(i, 15) = EXPECTED_RETURN_RNG + PROB_VAL * VOLATILITY_RNG / Sqr(TEMP_MATRIX(i, 2))
            For j = 1 To NCOLUMNS
                TEMP_VAL = SHORTFALL_RETURN_VECTOR(j, 1)
                TEMP_MATRIX(i, 15 + j) = NORMSDIST_FUNC((TEMP_VAL - EXPECTED_RETURN_RNG) / _
                                        (VOLATILITY_RNG / Sqr(TEMP_MATRIX(i, 2))), 0, 1, 0)
            Next j
        Else
            For j = 13 To 15 + NCOLUMNS: TEMP_MATRIX(i, j) = CVErr(xlErrNA): Next j
        End If
    Next i
'-------------------------------------------------------------------------------------
End Select
'-------------------------------------------------------------------------------------

WEALTH_TIME_INVESTMENT_HORIZON_FUNC = TEMP_MATRIX

Exit Function
ERROR_LABEL:
WEALTH_TIME_INVESTMENT_HORIZON_FUNC = Err.number
End Function

'// PERFECT

Function WEALTH_TIME_INVESTMENT_HORIZON_SIMULATION_FUNC( _
ByVal EXPECTED_EXPECTED_RETURN_VAL As Double, _
ByVal VOLATILITY_VAL As Double, _
Optional ByVal INITIAL_WEALTH_VAL As Double = 100, _
Optional ByVal INVESTMENT_HORIZON_VAL As Long = 10, _
Optional ByVal NO_PERIODS As Long = 100, _
Optional ByVal nLOOPS As Long = 1000, _
Optional ByVal OUTPUT As Integer = 2)

'Expected continuous return
'Volatiliy
'Initial Wealth
'Investment Horizon (in years)

Dim i As Long
Dim j As Long
Dim TEMP_MATRIX As Variant

On Error GoTo ERROR_LABEL

ReDim TEMP_MATRIX(1 To nLOOPS + 2, 1 To NO_PERIODS + 2)
TEMP_MATRIX(1, 1) = "j"
TEMP_MATRIX(2, 1) = "dT"
For j = 1 To NO_PERIODS + 1
    TEMP_MATRIX(1, j + 1) = (j - 1)
    TEMP_MATRIX(2, j + 1) = TEMP_MATRIX(1, j + 1) * INVESTMENT_HORIZON_VAL / NO_PERIODS
Next j
j = 1
For i = 3 To nLOOPS + 2
    TEMP_MATRIX(i, 1) = "i" & i - 2
    TEMP_MATRIX(i, j + 1) = INITIAL_WEALTH_VAL * Exp(( _
    EXPECTED_EXPECTED_RETURN_VAL + 0.5 * _
    VOLATILITY_VAL ^ 2) * TEMP_MATRIX(2, j + 1) + VOLATILITY_VAL * _
    Sqr(TEMP_MATRIX(2, j + 1)) * NORMSINV_FUNC(Rnd(), 0, 1, 0))
Next i

For j = 2 To NO_PERIODS + 1
    For i = 3 To nLOOPS + 2
        TEMP_MATRIX(i, j + 1) = TEMP_MATRIX(i, j) * _
        Exp((EXPECTED_EXPECTED_RETURN_VAL + 0.5 * VOLATILITY_VAL ^ 2) * _
        (INVESTMENT_HORIZON_VAL / NO_PERIODS) + VOLATILITY_VAL * _
        Sqr(INVESTMENT_HORIZON_VAL / NO_PERIODS) * NORMSINV_FUNC(Rnd(), 0, 1, 0))
    Next i
Next j

'--------------------------------------------------------------------------------------------------------------
Select Case OUTPUT
'--------------------------------------------------------------------------------------------------------------
Case 0
'--------------------------------------------------------------------------------------------------------------
    WEALTH_TIME_INVESTMENT_HORIZON_SIMULATION_FUNC = TEMP_MATRIX
'--------------------------------------------------------------------------------------------------------------
Case Else
'--------------------------------------------------------------------------------------------------------------
    TEMP_MATRIX = MATRIX_GET_SUB_MATRIX_FUNC(TEMP_MATRIX, 3, nLOOPS + 2, 2, NO_PERIODS + 2)
    If OUTPUT = 1 Then
        WEALTH_TIME_INVESTMENT_HORIZON_SIMULATION_FUNC = TEMP_MATRIX
    Else
        TEMP_MATRIX = DATA_BASIC_MOMENTS_FUNC(TEMP_MATRIX, 0, 0, 0.05, 1)
        TEMP_MATRIX = MATRIX_ADD_COLUMNS_FUNC(TEMP_MATRIX, 1, 1)
        TEMP_MATRIX(1, 1) = "DT"
        For j = 1 To NO_PERIODS + 1
            TEMP_MATRIX(j + 1, 1) = (j - 1) * INVESTMENT_HORIZON_VAL / NO_PERIODS
        Next j
        WEALTH_TIME_INVESTMENT_HORIZON_SIMULATION_FUNC = TEMP_MATRIX
    End If
'--------------------------------------------------------------------------------------------------------------
End Select
'--------------------------------------------------------------------------------------------------------------

Exit Function
ERROR_LABEL:
WEALTH_TIME_INVESTMENT_HORIZON_SIMULATION_FUNC = Err.number
End Function
