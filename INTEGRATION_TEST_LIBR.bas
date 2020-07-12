Attribute VB_Name = "INTEGRATION_TEST_LIBR"

'-------------------------------------------------------------------------------------
'-------------------------------------------------------------------------------------

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.
'-------------------------------------------------------------------------------------
'-------------------------------------------------------------------------------------

Private PUB_INGRATION_FUNC_STR As String

'WORK IN PROGRESS!!!

Function TEST_INTEGRATION_FUNC(Optional ByVal VERSION As Integer = 2, _
Optional ByVal tolerance As Double = 10 ^ -14, _
Optional ByVal MAX_LEVEL As Long = 14)

Dim i As Long
Dim j As Long
Dim k As Long

Dim ii As Long
Dim jj As Long

Dim NROWS As Long
Dim NCOLUMNS As Long

Dim TEMP_ARR As Variant
Dim TEMP_STR As String
Dim FORMULA_STR As String

Dim TEMP_MATRIX As Variant

On Error GoTo ERROR_LABEL
FORMULA_STR = _
"1/Sqrt(x)|0|1|2|,Sqrt(4-x^2)|0|2|3.14159265358979|,LN(x)|0|1|-1|,x*LN(x)|0|1|-0.25|,LN(x)/Sqrt(x)|0|1|-4|,4/(1+x^2)|0|1|3.14159265358979|," & _
"(SIN(x)^4)*(COS(x)^2)|0|1.5707963267949|0.098174770424681|,COS(x)|0|3.14159265358979|0|,COS(LN(x))|0|1|0.5|,Sqrt(4*x-x^2)|0|2|3.14159265358979|," & _
"5*x^2|0|10|1666.66666666667|,x^0.125|0|1|0.888888888888889|,1/x|1|10|2.30258509299405|,LN(x)/(1-x)|0.5|1|-0.582240526465013|," & _
"EXP(-1/COS(x))|0|1.0471975511966|0.307694394903451|,(x*(x+88)*(x-88)*(x+47)*(x-47)*(x+117)*(x-117))^2|-128|128|1.31026895522267E+28|," & _
"EXP(-(x^2))|0|100|0.886226925452758|,2*x^2/(x+1)/(x-1)-x/LN(x)|0|1|0.0364899739785765|,x*LN(1+x)|0|1|0.25|,x^2*ATAN(x)|0|1|0.210657251225807|," & _
"EXP(x)*COS(x)|0|1.5707963267949|1.90523869048268|,ATAN(Sqrt(x^2+2))/(1+x^2)/Sqrt(x^2+2)|0|1|0.514041895890071|,LN(x)*Sqrt(x)|0|1|-0.444444444444444|," & _
"Sqrt(1-x^2)|0|1|0.785398163397448|,Sqrt(x)/Sqrt(1-x^2)|0|1|1.19814023473559|,(LN(x))^2|0|1|2|,LN(COS(x))|0|1.5707963267949|-1.0887930451518|," & _
"Sqrt(TAN(x))|0|1.5707963267949|2.22144146907918|,LN(x^2)|0|1|-2|,x*SIN(x)/(1+(COS(x))^2)|0|3.14159265358979|2.46740110027234|," & _
"1/(x-2)/((1-x)^0.25)/((1+x)^0.75)|0|1|-0.691183688767295|,COS(PI()*x)/Sqrt(1-x)|-1|1|-0.690494588746605|,x^2*LN(x)/(x^2-1)/(x^4+1)|0|1|0.180671262590655|," & _
"1/(1-2*x+2*x^2)|0|1|1.5707963267949|,EXP(1-1/x)/Sqrt(x^3-x^4)|0|1|1.77245385090552|,EXP(1-1/x)*COS(1/x-1)/x^2|1|2|0.618664059898892|," & _
"x/Sqrt(1-x^2)|0|1|1|,(1-x)^4*x^4/(1+x^2)|0|1|0.00126448926734968|,x^4*(1-x)^4|0|1|0.00158730158730159|,ATAN(Sqrt(x^2+1))/(x^2+1)^(3/2)|0|1|0.590489270886385|," & _
"(COS(3*x))^2/(5-4*COS(2*x))|0|6.28318530717959|1.17809724509617|,1/(1+x^2+x^4+x^6)|-1|1|1.40862340353768|,(PI()/4-x*TAN(x))*TAN(x)|0|0.785398163397448|0.141798825704517|," & _
"x^2/(SIN(x))^2|0|1.5707963267949|2.1775860903036|,(LN(COS(x)))^2|0|1.5707963267949|2.04662202447274|,(LN(x))^2/(x^2+x+1)|0|1|1.76804762350016|," & _
"(LN(1+x^2))/x^2|0|1|0.877649146234951|,ATAN(x)/(x*Sqrt(1-x^2))|0|1|1.38445839302434|,x^2/(1+x^4)/Sqrt(1-x^4)|0|1|0.392699081698724|," & _
"EXP(-(x^2))*LN(x)|17|42|2.5657285005611E-127|,1/Sqrt(1-x^2)|-1|1|3.14159265358979|,(1+x)^2*SIN(2*PI()/(1+x))|0|1|-1.2577347112237|," & _
"x*(1-x)^0.1|0|1|0.432900432900432|,(LN(COS(x)))^2|0|1.5707963267949|2.04662202447274|,LN(SIN(x)^3)*COS(x)|0|1|-2.96013608748744|"
k = Len(FORMULA_STR)
i = 1: NCOLUMNS = 0
Do While Mid(FORMULA_STR, i, 1) <> ","
    If Mid(FORMULA_STR, i, 1) = "|" Then: NCOLUMNS = NCOLUMNS + 1
    i = i + 1
Loop
i = 1: NROWS = 0
Do While i <= k
    If Mid(FORMULA_STR, i, 1) = "," Then: NROWS = NROWS + 1
    i = i + 1
Loop
ReDim TEMP_MATRIX(0 To NROWS, 1 To NCOLUMNS + 5)
'Integral    Est. Error  Level   Time (secs)

TEMP_MATRIX(0, 1) = "f(x)"
TEMP_MATRIX(0, 2) = "a"
TEMP_MATRIX(0, 3) = "b"
TEMP_MATRIX(0, 4) = "True Integral" 'To 15 significant digits
Select Case VERSION
Case 0
    TEMP_MATRIX(0, 5) = "Tanh Sinh"
Case 1
    TEMP_MATRIX(0, 5) = "Romberg"
Case Else
    TEMP_MATRIX(0, 5) = "Gauss Kronrod"
End Select
TEMP_MATRIX(0, 6) = "Est. Error"
TEMP_MATRIX(0, 7) = "Level"
TEMP_MATRIX(0, 8) = "Time (secs)"
TEMP_MATRIX(0, 9) = "True Error"

i = 1
For ii = 1 To NROWS
    jj = 1: GoSub PARSE_LINE
    For jj = 2 To NCOLUMNS
        GoSub PARSE_LINE
    Next jj
    GoSub INTEGRATE_LINE
    i = i + 1
Next ii

TEST_INTEGRATION_FUNC = TEMP_MATRIX

Exit Function
'-----------------------------------------------------------------------------------------------------------------
PARSE_LINE:
'-----------------------------------------------------------------------------------------------------------------
    j = InStr(i, FORMULA_STR, "|")
    TEMP_STR = Mid(FORMULA_STR, i, j - i)
    If IsNumeric(TEMP_STR) Then
        TEMP_MATRIX(ii, jj) = Val(TEMP_STR)
    Else
        TEMP_MATRIX(ii, jj) = TEMP_STR
    End If
    i = j + 1
'-----------------------------------------------------------------------------------------------------------------
Return
'-----------------------------------------------------------------------------------------------------------------
INTEGRATE_LINE:
'-----------------------------------------------------------------------------------------------------------------
    PUB_INGRATION_FUNC_STR = TEMP_MATRIX(ii, 1)
    Select Case VERSION
    Case 0 'None of the functions reached Max. Level (10).
        'TEMP_ARR = TANH_SINH_FUNC("TEST_INTEGRATION_OBJ_FUNC", TEMP_MATRIX(ii, 2), TEMP_MATRIX(ii, 3), tolerance, MAX_LEVEL)
    Case 1 'Almost half the functions reached Max. Loops (15).
        'TEMP_ARR = ROMBERG_FUNC("TEST_INTEGRATION_OBJ_FUNC", TEMP_MATRIX(ii, 2), TEMP_MATRIX(ii, 3), tolerance, MAX_LEVEL)
    Case Else 'Three functions reached Max. Panels (400).
        'TEMP_ARR = GAUSS_KRONROD_FUNC("TEST_INTEGRATION_OBJ_FUNC", TEMP_MATRIX(ii, 2), TEMP_MATRIX(ii, 3), tolerance, MAX_LEVEL)
    End Select
    If IsArray(TEMP_ARR) = True Then
        TEMP_MATRIX(ii, 5) = TEMP_ARR(LBound(TEMP_ARR) + 0)
        TEMP_MATRIX(ii, 6) = TEMP_ARR(LBound(TEMP_ARR) + 1)
        TEMP_MATRIX(ii, 7) = TEMP_ARR(LBound(TEMP_ARR) + 2)
        TEMP_MATRIX(ii, 8) = TEMP_ARR(LBound(TEMP_ARR) + 3)
        If ii <> 16 Then
            TEMP_MATRIX(ii, 9) = Abs(TEMP_MATRIX(ii, 4) - TEMP_MATRIX(ii, 5))
        Else '(x*(x+88)*(x-88)*(x+47)*(x-47)*(x+117)*(x-117))^2
            'ROMBERG: Unable to calculate result
            TEMP_MATRIX(ii, 9) = Abs(TEMP_MATRIX(ii, 4) / 1E+28 - TEMP_MATRIX(ii, 5) / 1E+28)
        End If
    Else
        TEMP_MATRIX(ii, 5) = False
        TEMP_MATRIX(ii, 6) = False
        TEMP_MATRIX(ii, 7) = False
        TEMP_MATRIX(ii, 8) = False
        TEMP_MATRIX(ii, 9) = False
    End If
'-----------------------------------------------------------------------------------------------------------------
Return
'-----------------------------------------------------------------------------------------------------------------
ERROR_LABEL:
TEST_INTEGRATION_FUNC = Err.number
End Function

Function TEST_INTEGRATION_OBJ_FUNC(ByVal X_VAL As Double)
TEST_INTEGRATION_OBJ_FUNC = Excel.Application.Evaluate(Replace(PUB_INGRATION_FUNC_STR, "x", X_VAL))
End Function
