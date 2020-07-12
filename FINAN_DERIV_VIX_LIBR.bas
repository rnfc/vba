Attribute VB_Name = "FINAN_DERIV_VIX_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.

'************************************************************************************
'************************************************************************************
'FUNCTION      : VIX_OPTION_FUNC

'DESCRIPTION   : The VIX has become a popular volatility index that is based on a
'weighted average of S&P 500 options that straddle a 30-day maturity. This manner
'of calculating the VIX emerged in September of 2003 and is documented with an
'example by the CBOE. In this routine, the calculation of the VIX is reproduced in a
'template to automate and to some degree simplify the calculation. Further,
'one can also apply other option series to calculate a VIX-type analysis for the
'underlying security which is of great benefit because the calculation is independent
'of option pricing model biases.

'LIBRARY       : DERIVATIVES
'GROUP         : VIX
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 06/24/2010
'REFERENCE     :
'http://papers.ssrn.com/sol3/papers.cfm?abstract_id=1103971
'http://www.optiontradingtips.com/options101/vix-sp500.xls
'http://www.cboe.com/micro/vix/vixwhite.pdf
'************************************************************************************
'************************************************************************************

Function VIX_OPTION_FUNC(ByRef NT_STRIKE_RNG As Variant, _
ByRef NT_CALL_RNG As Variant, _
ByRef NT_PUT_RNG As Variant, _
ByVal NT_EXPIRATION_DAYS As Double, _
ByRef FT_STRIKE_RNG As Variant, _
ByRef FT_CALL_RNG As Variant, _
ByRef FT_PUT_RNG As Variant, _
ByVal FT_EXPIRATION_DAYS As Double, _
ByVal RISK_FREE_RATE As Double, _
Optional ByVal CURRENT_TIME As Date = 0, _
Optional ByVal DAYS_PER_MONTH As Double = 30, _
Optional ByVal DAYS_PER_YEAR As Double = 365, _
Optional ByVal tolerance As Double = 0.01, _
Optional ByVal OUTPUT As Integer = 1)

'NT --> Near Term
'FT --> Far Term
'Risk-free Rate --> Annual

Const h As Long = 7
Dim i As Long
Dim j As Long
Dim k As Long
Dim l As Long

Dim NROWS As Long
Dim NCOLUMNS As Long

'-----------------------------------------------------------------
Dim NT_NROWS As Long
Dim FT_NROWS As Long

Dim NT_YEARS As Double
Dim NT_DAYS As Double

Dim FT_YEARS As Double
Dim FT_DAYS As Double

'-----------------------------------------------------------------
Dim NT_SUM_VAL As Double
Dim FT_SUM_VAL As Double

Dim NT_VIX_VAL As Double
Dim FT_VIX_VAL As Double

Dim VIX_CALC_VAL As Double
'-----------------------------------------------------------------
Dim NT_TEMP0_VAL As Double
Dim NT_TEMP1_VAL As Double
Dim NT_DELTA_VAL As Double

Dim NT_MIN_VAL As Double
Dim NT_MAX_VAL As Double
Dim NT_STRIKE0_VAL As Double

Dim NT_CALL_VAL As Double
Dim NT_PUT_VAL As Double
Dim NT_LEVEL_VAL As Double
Dim NT_STRIKE1_VAL As Double

Dim NT_CP_VAL As Double
Dim NT_VAR_VAL As Double 'Variance
Dim NT_TERM_VAL As Double 'Term1
'-----------------------------------------------------------------
Dim FT_TEMP0_VAL As Double
Dim FT_TEMP1_VAL As Double
Dim FT_DELTA_VAL As Double

Dim FT_MIN_VAL As Double
Dim FT_MAX_VAL As Double
Dim FT_STRIKE0_VAL As Double

Dim FT_CALL_VAL As Double
Dim FT_PUT_VAL As Double
Dim FT_LEVEL_VAL As Double
Dim FT_STRIKE1_VAL As Double

Dim FT_CP_VAL As Double
Dim FT_VAR_VAL As Double 'Variance
Dim FT_TERM_VAL As Double 'Term1
'-----------------------------------------------------------------

Dim NT_STRIKE_VECTOR As Variant
Dim NT_CALL_VECTOR As Variant
Dim NT_PUT_VECTOR As Variant

Dim FT_STRIKE_VECTOR As Variant
Dim FT_CALL_VECTOR As Variant
Dim FT_PUT_VECTOR As Variant
'-----------------------------------------------------------------

Dim TEMP_MATRIX As Variant
Dim HEADINGS_STR As String

On Error GoTo ERROR_LABEL

'--------------------------------------------------------------------------------------------------------
NT_STRIKE_VECTOR = NT_STRIKE_RNG
If UBound(NT_STRIKE_VECTOR, 1) = 1 Then
    NT_STRIKE_VECTOR = MATRIX_TRANSPOSE_FUNC(NT_STRIKE_VECTOR)
End If
NT_NROWS = UBound(NT_STRIKE_VECTOR, 1)

NT_CALL_VECTOR = NT_CALL_RNG
If UBound(NT_CALL_VECTOR, 1) = 1 Then
    NT_CALL_VECTOR = MATRIX_TRANSPOSE_FUNC(NT_CALL_VECTOR)
End If
If UBound(NT_CALL_VECTOR, 1) <> NT_NROWS Then: GoTo ERROR_LABEL

NT_PUT_VECTOR = NT_PUT_RNG
If UBound(NT_PUT_VECTOR, 1) = 1 Then
    NT_PUT_VECTOR = MATRIX_TRANSPOSE_FUNC(NT_PUT_VECTOR)
End If
If UBound(NT_PUT_VECTOR, 1) <> NT_NROWS Then: GoTo ERROR_LABEL
'--------------------------------------------------------------------------------------------------------
FT_STRIKE_VECTOR = FT_STRIKE_RNG
If UBound(FT_STRIKE_VECTOR, 1) = 1 Then
    FT_STRIKE_VECTOR = MATRIX_TRANSPOSE_FUNC(FT_STRIKE_VECTOR)
End If
FT_NROWS = UBound(FT_STRIKE_VECTOR, 1)

FT_CALL_VECTOR = FT_CALL_RNG
If UBound(FT_CALL_VECTOR, 1) = 1 Then
    FT_CALL_VECTOR = MATRIX_TRANSPOSE_FUNC(FT_CALL_VECTOR)
End If
If UBound(FT_CALL_VECTOR, 1) <> FT_NROWS Then: GoTo ERROR_LABEL

FT_PUT_VECTOR = FT_PUT_RNG
If UBound(FT_PUT_VECTOR, 1) = 1 Then
    FT_PUT_VECTOR = MATRIX_TRANSPOSE_FUNC(FT_PUT_VECTOR)
End If
If UBound(FT_PUT_VECTOR, 1) <> FT_NROWS Then: GoTo ERROR_LABEL
'--------------------------------------------------------------------------------------------------------
GoSub TIME_LINE
'------------------------------------------------------------------------------------------------------------

If OUTPUT > 1 Then
    GoSub REDIM_LINE
End If
i = 0
GoSub NT_STRIKE0_LINE
For i = 1 To NT_NROWS: GoSub NT_STRIKE0_LINE: Next i
i = 0
GoSub FT_STRIKE0_LINE
For i = 1 To FT_NROWS: GoSub FT_STRIKE0_LINE: Next i

If OUTPUT > 1 Then
    For i = 1 To NT_NROWS
        TEMP_MATRIX(i, 1) = Abs(NT_CALL_VECTOR(i, 1) - NT_PUT_VECTOR(i, 1))
        TEMP_MATRIX(i, 2) = NT_STRIKE_VECTOR(i, 1)
        TEMP_MATRIX(i, 3) = NT_CALL_VECTOR(i, 1)
        TEMP_MATRIX(i, 4) = NT_PUT_VECTOR(i, 1)
    Next i
    For i = 1 To FT_NROWS
        TEMP_MATRIX(i, 1 + h) = Abs(FT_CALL_VECTOR(i, 1) - FT_PUT_VECTOR(i, 1))
        TEMP_MATRIX(i, 2 + h) = FT_STRIKE_VECTOR(i, 1)
        TEMP_MATRIX(i, 3 + h) = FT_CALL_VECTOR(i, 1)
        TEMP_MATRIX(i, 4 + h) = FT_PUT_VECTOR(i, 1)
    Next i
End If
GoSub NT_STRIKE1_LINE
NT_SUM_VAL = 0
For i = 1 To NT_NROWS
    GoSub NT_CP_LINE
    If OUTPUT > 1 Then
        TEMP_MATRIX(i, 5) = NT_DELTA_VAL
        TEMP_MATRIX(i, 6) = NT_CP_VAL
        TEMP_MATRIX(i, 7) = NT_VIX_VAL
    End If
Next i
GoSub FT_STRIKE1_LINE
FT_SUM_VAL = 0
For i = 1 To FT_NROWS
    GoSub FT_CP_LINE
    If OUTPUT > 1 Then
        TEMP_MATRIX(i, 5 + h) = FT_DELTA_VAL
        TEMP_MATRIX(i, 6 + h) = FT_CP_VAL
        TEMP_MATRIX(i, 7 + h) = FT_VIX_VAL
    End If
Next i
GoSub VIX_CALC_LINE

'------------------------------------------------------------------------------------------------------------
Select Case OUTPUT
'------------------------------------------------------------------------------------------------------------
Case 0
    VIX_OPTION_FUNC = VIX_CALC_VAL
Case 1
    VIX_OPTION_FUNC = Array(NT_STRIKE0_VAL, NT_CALL_VAL, NT_PUT_VAL, NT_LEVEL_VAL, NT_STRIKE1_VAL, NT_VAR_VAL, NT_TERM_VAL, _
                            FT_STRIKE0_VAL, FT_CALL_VAL, FT_PUT_VAL, FT_LEVEL_VAL, FT_STRIKE1_VAL, FT_VAR_VAL, FT_TERM_VAL, _
                            VIX_CALC_VAL)
Case Else
    TEMP_MATRIX(NROWS + 2, 1) = NT_STRIKE0_VAL
    TEMP_MATRIX(NROWS + 2, 2) = NT_CALL_VAL
    TEMP_MATRIX(NROWS + 2, 3) = NT_PUT_VAL
    TEMP_MATRIX(NROWS + 2, 4) = NT_LEVEL_VAL
    TEMP_MATRIX(NROWS + 2, 5) = NT_STRIKE1_VAL
    TEMP_MATRIX(NROWS + 2, 6) = NT_VAR_VAL
    TEMP_MATRIX(NROWS + 2, 7) = NT_TERM_VAL
    
    TEMP_MATRIX(NROWS + 2, 8) = FT_STRIKE0_VAL
    TEMP_MATRIX(NROWS + 2, 9) = FT_CALL_VAL
    TEMP_MATRIX(NROWS + 2, 10) = FT_PUT_VAL
    TEMP_MATRIX(NROWS + 2, 11) = FT_LEVEL_VAL
    TEMP_MATRIX(NROWS + 2, 12) = FT_STRIKE1_VAL
    TEMP_MATRIX(NROWS + 2, 13) = FT_VAR_VAL
    TEMP_MATRIX(NROWS + 2, 14) = FT_TERM_VAL
    
    TEMP_MATRIX(NROWS + 3, 2) = VIX_CALC_VAL
    VIX_OPTION_FUNC = TEMP_MATRIX
'------------------------------------------------------------------------------------------------------------
End Select
'------------------------------------------------------------------------------------------------------------

'------------------------------------------------------------------------------------------------------------
Exit Function
'------------------------------------------------------------------------------------------------------------
TIME_LINE:
'------------------------------------------------------------------------------------------------------------
    If CURRENT_TIME = 0 Then: CURRENT_TIME = Now
    NT_YEARS = (1440 - (Hour(CURRENT_TIME) * 60 + Minute(CURRENT_TIME) + Second(CURRENT_TIME) / 60) + 510) / (1440 * DAYS_PER_YEAR) + (NT_EXPIRATION_DAYS - 2) / DAYS_PER_YEAR
    NT_DAYS = NT_YEARS * DAYS_PER_YEAR
    
    FT_YEARS = (1440 - (Hour(CURRENT_TIME) * 60 + Minute(CURRENT_TIME) + Second(CURRENT_TIME) / 60) + 510) / (1440 * DAYS_PER_YEAR) + (FT_EXPIRATION_DAYS - 2) / DAYS_PER_YEAR
    FT_DAYS = FT_YEARS * DAYS_PER_YEAR
'------------------------------------------------------------------------------------------------------------
Return
'------------------------------------------------------------------------------------------------------------
REDIM_LINE:
'------------------------------------------------------------------------------------------------------------
    l = Len(HEADINGS_STR)
    NCOLUMNS = 14
    NROWS = IIf(NT_NROWS > FT_NROWS, NT_NROWS, FT_NROWS)
    ReDim TEMP_MATRIX(0 To NROWS + 3, 1 To NCOLUMNS)
    For i = 1 To NROWS + 3: For j = 1 To NCOLUMNS: TEMP_MATRIX(i, j) = "": Next j: Next i
    HEADINGS_STR = "Near-Term: Difference,Near-Term: Strike,Near-Term: Call,Near-Term: Put,Near-Term: Delta-Strike,Near-Term: C/P-Near,Near-Term: VIX-Near,Far-term: Difference,Far-term: Strike,Far-term: Call,Far-term: Put,Far-term: Delta-Strike,Far-term: C/P-Far,Far-term: VIX-Far,"
    i = 1
    For k = 1 To NCOLUMNS
        j = InStr(i, HEADINGS_STR, ",")
        TEMP_MATRIX(0, k) = Mid(HEADINGS_STR, i, j - i)
        i = j + 1
    Next k
    HEADINGS_STR = "Near-Term: Strike,Near-Term: Call,Near-Term: Put,Near-Term: Level,Near-Term: Ref-Strike,Near-Term: Variance,Near-Term: Term1,Far-Term: Strike,Far-Term: Call,Far-Term: Put,Far-Term: Level,Far-Term: Ref-Strike,Far-Term: Variance,Far-Term: Term2,"
    i = 1
    For k = 1 To NCOLUMNS
        j = InStr(i, HEADINGS_STR, ",")
        TEMP_MATRIX(NROWS + 1, k) = Mid(HEADINGS_STR, i, j - i)
        i = j + 1
    Next k
    TEMP_MATRIX(NROWS + 3, 1) = "VIX Calculation"
'------------------------------------------------------------------------------------------------------------
Return
'------------------------------------------------------------------------------------------------------------
NT_STRIKE0_LINE:
'------------------------------------------------------------------------------------------------------------
    If i > 0 Then
        If NT_STRIKE_VECTOR(i, 1) < NT_MIN_VAL Then: NT_MIN_VAL = NT_STRIKE_VECTOR(i, 1)
        If NT_STRIKE_VECTOR(i, 1) > NT_MAX_VAL Then: NT_MAX_VAL = NT_STRIKE_VECTOR(i, 1)
        
        NT_TEMP0_VAL = Abs(NT_CALL_VECTOR(i, 1) - NT_PUT_VECTOR(i, 1))
        If NT_TEMP0_VAL < NT_TEMP1_VAL Then
            NT_TEMP1_VAL = NT_TEMP0_VAL
            NT_STRIKE0_VAL = NT_STRIKE_VECTOR(i, 1)
            NT_CALL_VAL = NT_CALL_VECTOR(i, 1)
            NT_PUT_VAL = NT_PUT_VECTOR(i, 1)
        End If
    Else
        NT_MIN_VAL = 2 ^ 52
        NT_MAX_VAL = -2 ^ 52
        NT_TEMP1_VAL = 2 ^ 52
    End If
'------------------------------------------------------------------------------------------------------------
Return
'------------------------------------------------------------------------------------------------------------
FT_STRIKE0_LINE:
'------------------------------------------------------------------------------------------------------------
    If i > 0 Then
        If FT_STRIKE_VECTOR(i, 1) < FT_MIN_VAL Then: FT_MIN_VAL = FT_STRIKE_VECTOR(i, 1)
        If FT_STRIKE_VECTOR(i, 1) > FT_MAX_VAL Then: FT_MAX_VAL = FT_STRIKE_VECTOR(i, 1)
        
        FT_TEMP0_VAL = Abs(FT_CALL_VECTOR(i, 1) - FT_PUT_VECTOR(i, 1))
        If FT_TEMP0_VAL < FT_TEMP1_VAL Then
            FT_TEMP1_VAL = FT_TEMP0_VAL
            FT_STRIKE0_VAL = FT_STRIKE_VECTOR(i, 1)
            FT_CALL_VAL = FT_CALL_VECTOR(i, 1)
            FT_PUT_VAL = FT_PUT_VECTOR(i, 1)
        End If
    Else
        FT_MIN_VAL = 2 ^ 52
        FT_MAX_VAL = -2 ^ 52
        FT_TEMP1_VAL = 2 ^ 52
    End If
'------------------------------------------------------------------------------------------------------------
Return
'------------------------------------------------------------------------------------------------------------
NT_STRIKE1_LINE:
'------------------------------------------------------------------------------------------------------------
    NT_LEVEL_VAL = NT_STRIKE0_VAL + Exp(RISK_FREE_RATE * NT_YEARS) * (NT_CALL_VAL - NT_PUT_VAL)
    If NT_CALL_VAL > NT_PUT_VAL Then
        NT_STRIKE1_VAL = NT_STRIKE0_VAL
    Else
        NT_TEMP0_VAL = NT_STRIKE0_VAL - tolerance
        NT_STRIKE1_VAL = NT_TEMP0_VAL
        For i = 2 To NT_NROWS
            If NT_STRIKE_VECTOR(i - 1, 1) < NT_TEMP0_VAL And NT_STRIKE_VECTOR(i, 1) >= NT_TEMP0_VAL Then
                NT_STRIKE1_VAL = NT_STRIKE_VECTOR(i, 1)
            End If
        Next i
    End If
'------------------------------------------------------------------------------------------------------------
Return
'------------------------------------------------------------------------------------------------------------
FT_STRIKE1_LINE:
'------------------------------------------------------------------------------------------------------------
    FT_LEVEL_VAL = FT_STRIKE0_VAL + Exp(RISK_FREE_RATE * FT_YEARS) * (FT_CALL_VAL - FT_PUT_VAL)
    If FT_CALL_VAL > FT_PUT_VAL Then
        FT_STRIKE1_VAL = FT_STRIKE0_VAL
    Else
        FT_TEMP0_VAL = FT_STRIKE0_VAL - tolerance
        FT_STRIKE1_VAL = FT_TEMP0_VAL
        For i = 2 To FT_NROWS
            If FT_STRIKE_VECTOR(i - 1, 1) < FT_TEMP0_VAL And FT_STRIKE_VECTOR(i, 1) >= FT_TEMP0_VAL Then
                FT_STRIKE1_VAL = FT_STRIKE_VECTOR(i, 1)
            End If
        Next i
    End If
'------------------------------------------------------------------------------------------------------------
Return
'------------------------------------------------------------------------------------------------------------
NT_CP_LINE:
'------------------------------------------------------------------------------------------------------------
    If NT_STRIKE_VECTOR(i, 1) = NT_STRIKE1_VAL Then
        NT_CP_VAL = (NT_CALL_VECTOR(i, 1) + NT_PUT_VECTOR(i, 1)) / 2
    Else
        If NT_STRIKE_VECTOR(i, 1) > NT_STRIKE1_VAL Then
            NT_CP_VAL = NT_CALL_VECTOR(i, 1)
        Else
            NT_CP_VAL = NT_PUT_VECTOR(i, 1)
        End If
    End If
    If NT_STRIKE_VECTOR(i, 1) = NT_MIN_VAL Then
        If i < NT_NROWS Then
            NT_DELTA_VAL = NT_STRIKE_VECTOR(i + 1, 1) - NT_STRIKE_VECTOR(i, 1)
        Else
            NT_DELTA_VAL = 0 '0 - NT_STRIKE_VECTOR(i, 1)
        End If
    Else
        If NT_STRIKE_VECTOR(i, 1) = NT_MAX_VAL Then
            If i > 1 Then
                NT_DELTA_VAL = NT_STRIKE_VECTOR(i, 1) - NT_STRIKE_VECTOR(i - 1, 1)
            Else
                NT_DELTA_VAL = 0 'NT_STRIKE_VECTOR(i, 1) - 0
            End If
        Else
            If i < NT_NROWS Then
                NT_DELTA_VAL = (NT_STRIKE_VECTOR(i + 1, 1) - NT_STRIKE_VECTOR(i - 1, 1)) / 2
            Else
                NT_DELTA_VAL = (NT_STRIKE_VECTOR(i, 1) - NT_STRIKE_VECTOR(i - 1, 1)) / 2
            End If
        End If
    End If
    NT_VIX_VAL = NT_DELTA_VAL / NT_STRIKE_VECTOR(i, 1) ^ 2 * Exp(RISK_FREE_RATE * NT_YEARS) * NT_CP_VAL
    NT_SUM_VAL = NT_SUM_VAL + NT_VIX_VAL
    NT_VAR_VAL = (2 / NT_YEARS) * NT_SUM_VAL - ((NT_LEVEL_VAL / NT_STRIKE1_VAL - 1) ^ 2) / NT_YEARS
    NT_TERM_VAL = NT_YEARS * NT_VAR_VAL * ((FT_DAYS - DAYS_PER_MONTH) / (FT_DAYS - NT_DAYS))
'------------------------------------------------------------------------------------------------------------
Return
'------------------------------------------------------------------------------------------------------------
FT_CP_LINE:
'------------------------------------------------------------------------------------------------------------
    If FT_STRIKE_VECTOR(i, 1) = FT_STRIKE1_VAL Then
        FT_CP_VAL = (FT_CALL_VECTOR(i, 1) + FT_PUT_VECTOR(i, 1)) / 2
    Else
        If FT_STRIKE_VECTOR(i, 1) > FT_STRIKE1_VAL Then
            FT_CP_VAL = FT_CALL_VECTOR(i, 1)
        Else
            FT_CP_VAL = FT_PUT_VECTOR(i, 1)
        End If
    End If
    If FT_STRIKE_VECTOR(i, 1) = FT_MIN_VAL Then
        If i < FT_NROWS Then
            FT_DELTA_VAL = FT_STRIKE_VECTOR(i + 1, 1) - FT_STRIKE_VECTOR(i, 1)
        Else
            FT_DELTA_VAL = 0 '0 - FT_STRIKE_VECTOR(i, 1)
        End If
    Else
        If FT_STRIKE_VECTOR(i, 1) = FT_MAX_VAL Then
            If i > 1 Then
                FT_DELTA_VAL = FT_STRIKE_VECTOR(i, 1) - FT_STRIKE_VECTOR(i - 1, 1)
            Else
                FT_DELTA_VAL = 0 'FT_STRIKE_VECTOR(i, 1) - 0
            End If
        Else
            If i < FT_NROWS Then
                FT_DELTA_VAL = (FT_STRIKE_VECTOR(i + 1, 1) - FT_STRIKE_VECTOR(i - 1, 1)) / 2
            Else
                FT_DELTA_VAL = (FT_STRIKE_VECTOR(i, 1) - FT_STRIKE_VECTOR(i - 1, 1)) / 2
            End If
        End If
    End If
    FT_VIX_VAL = FT_DELTA_VAL / FT_STRIKE_VECTOR(i, 1) ^ 2 * Exp(RISK_FREE_RATE * FT_YEARS) * FT_CP_VAL
    FT_SUM_VAL = FT_SUM_VAL + FT_VIX_VAL
    FT_VAR_VAL = (2 / FT_YEARS) * FT_SUM_VAL - ((FT_LEVEL_VAL / FT_STRIKE1_VAL - 1) ^ 2) / FT_YEARS
    FT_TERM_VAL = FT_YEARS * FT_VAR_VAL * ((DAYS_PER_MONTH - NT_DAYS) / (FT_DAYS - NT_DAYS))
'------------------------------------------------------------------------------------------------------------
Return
'------------------------------------------------------------------------------------------------------------
VIX_CALC_LINE:
'------------------------------------------------------------------------------------------------------------
    VIX_CALC_VAL = ((NT_TERM_VAL + FT_TERM_VAL) * DAYS_PER_YEAR / DAYS_PER_MONTH) ^ 0.5
    VIX_CALC_VAL = VIX_CALC_VAL * 100
'------------------------------------------------------------------------------------------------------------
Return
'------------------------------------------------------------------------------------------------------------
ERROR_LABEL:
VIX_OPTION_FUNC = Err.number
End Function
