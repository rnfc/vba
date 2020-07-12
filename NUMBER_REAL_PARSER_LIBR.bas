Attribute VB_Name = "NUMBER_REAL_PARSER_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.
                    
'---------------------------------------------------------------------------------------------------------------
'---------------------------------------------------------------------------------------------------------------
'We note that the function evaluation – the focal point of this task - is performed in 5 principal statements:
'1) Function declaration, storage, and parsing
'2) Syntax error check
'3) Load variable value
'4) Evaluate (eval)
'5) Activate error trap for checking domain errors
'Just clean and straightforward. No other statement is necessary. This takes an overall advantage in complicated
'math routines, when the main focus must be concentrated on the mathematical algorithm, without any other
'technical distraction and extra subroutines calls. Note also that declaration is needed only once at a time.
'Of course that computation speed is also important and must be put in evidence.
'---------------------------------------------------------------------------------------------------------------
'SOURCE:
'http://digilander.libero.it/foxes/mathparser/MathExpressionsParser.htm
'http://digilander.libero.it/foxes/integr/Newton_Cotes_variab.htm
'---------------------------------------------------------------------------------------------------------------
'---------------------------------------------------------------------------------------------------------------

'//PERFECT

'Excel user function for evaluating math expression.

Function MATH_PARSER_EVAL0_FUNC(ByVal FORMULA_STR As String, _
ParamArray PARAMETERS_RNG() As Variant)

Dim i As Long
Dim j As Long

Dim NSIZE As Long
Dim OK_FLAG As Boolean
Dim ERROR_STR As String
Dim PARSER_OBJ As New clsMathParser

On Error GoTo ERROR_LABEL

'----------------------------------
ERROR_STR = "--"
OK_FLAG = PARSER_OBJ.StoreExpression(FORMULA_STR)
'----------------------------------
If Not OK_FLAG Then
    ERROR_STR = PARSER_OBJ.ErrorDescription 'Err.Raise 1001
    GoTo ERROR_LABEL
End If
NSIZE = UBound(PARAMETERS_RNG) - LBound(PARAMETERS_RNG) + 1 'number of parameters
If NSIZE < PARSER_OBJ.VarTop Then
    ERROR_STR = "missing parameter" 'Err.Raise 1002
    GoTo ERROR_LABEL
End If
j = 1
For i = LBound(PARAMETERS_RNG) To UBound(PARAMETERS_RNG) 'load parameters values
    PARSER_OBJ.variable(j) = PARAMETERS_RNG(i)
    j = j + 1
Next i
MATH_PARSER_EVAL0_FUNC = PARSER_OBJ.Eval()
If Err.number <> 0 Then GoTo ERROR_LABEL

Exit Function
ERROR_LABEL:
MATH_PARSER_EVAL0_FUNC = ERROR_STR
End Function


'//PERFECT

Function MATH_PARSER_EVAL1_FUNC(ByVal FORMULA_STR As String, _
ByRef PARAMETERS_RNG As Variant, _
Optional ByVal ANGLE_DEG_FLAG As Boolean = False, _
Optional ByVal EXPLICIT_VAR_FLAG As Boolean = False, _
Optional ByRef ERROR_STR As String = "--", _
Optional ByRef PARSER_OBJ As clsMathParser)

Dim i As Long
Dim k As Long
Dim CHR_STR As String
Dim OK_FLAG As Boolean
Dim PARAMETERS_VECTOR As Variant

On Error GoTo ERROR_LABEL

If PARSER_OBJ Is Nothing Then: Set PARSER_OBJ = New clsMathParser
If IsArray(PARAMETERS_RNG) = True Then
    PARAMETERS_VECTOR = PARAMETERS_RNG
    If UBound(PARAMETERS_VECTOR, 1) = 1 Then
        PARAMETERS_VECTOR = MATRIX_TRANSPOSE_FUNC(PARAMETERS_VECTOR)
    End If
    GoSub PARSER_LINE
    If Not OK_FLAG Then
        ERROR_STR = PARSER_OBJ.ErrorDescription
        GoTo ERROR_LABEL 'Err.Raise 1001, , PARSER_OBJ.ErrorDescription
    End If
    k = UBound(PARAMETERS_VECTOR) - LBound(PARAMETERS_VECTOR) + 1 'UBound(PARAMETERS_VECTOR) + 1 'number of parameters
    If k < PARSER_OBJ.VarTop Then GoTo ERROR_LABEL 'Err.Raise 1002, , "missing parameter"
    If UBound(PARAMETERS_VECTOR, 2) = 1 Then
        For i = 1 To k
            PARSER_OBJ.variable(i) = PARAMETERS_VECTOR(i, 1) 'i - 1)
        Next i
    Else
        For i = 1 To k
            CHR_STR = PARAMETERS_VECTOR(i, 1)
            PARSER_OBJ.variable(CHR_STR) = PARAMETERS_VECTOR(i, 2) 'i - 1)
        Next i
    End If
    MATH_PARSER_EVAL1_FUNC = PARSER_OBJ.Eval() 'f(x,y)
Else 'Univariate
    GoSub PARSER_LINE
    If Not OK_FLAG Then
        ERROR_STR = PARSER_OBJ.ErrorDescription
        GoSub ERROR_LABEL
    Else
        MATH_PARSER_EVAL1_FUNC = PARSER_OBJ.Eval1(CDbl(PARAMETERS_RNG)) 'f(x): evaluate function value
        'Debug.Print "x=" + CStr(X_VAL); Tab(25); "f(x)=" + CStr(Y_VAL)
    End If
End If

Exit Function
'----------------------------------------------------------------------------------------------------------------
PARSER_LINE:
'----------------------------------------------------------------------------------------------------------------
    If PARSER_OBJ.Expression <> FORMULA_STR Then
        OK_FLAG = PARSER_OBJ.StoreExpression(FORMULA_STR) 'parse function
        'the unit of measure for angle computing (RAD (default), DEG or GRAD)
        If ANGLE_DEG_FLAG = True Then: PARSER_OBJ.AngleUnit = "DEG"
        PARSER_OBJ.OpAssignExplicit = EXPLICIT_VAR_FLAG 'If True all variables must be assigned
    Else
        OK_FLAG = True 'Already Parsed
'        Debug.Print "NICO"
    End If
'----------------------------------------------------------------------------------------------------------------
Return
'----------------------------------------------------------------------------------------------------------------
ERROR_LABEL:
MATH_PARSER_EVAL1_FUNC = ERROR_STR
End Function

'//PERFECT

Function MATH_PARSER_EVAL2_FUNC(ByVal FORMULA_STR As String, _
ByRef VARIABLES_RNG As Variant, _
ByRef PARAMETERS_RNG As Variant, _
ByRef XDATA_ARR() As Double, _
Optional ByVal XVAR_STR As String = "x", _
Optional ByVal ANGLE_DEG_FLAG As Boolean = False, _
Optional ByVal EXPLICIT_VAR_FLAG As Boolean = False, _
Optional ByRef ERROR_STR As String = "--", _
Optional ByRef PARSER_OBJ As clsMathParser)

Dim i As Long
Dim NROWS As Long
Dim OK_FLAG As Boolean
Dim PARAMETERS_VECTOR As Variant
Dim VARIABLES_VECTOR As Variant

On Error GoTo ERROR_LABEL

'-----------------------------------------------------------------------------------------
If PARSER_OBJ Is Nothing Then: Set PARSER_OBJ = New clsMathParser
'-----------------------------------------------------------------------------------------

If IsArray(VARIABLES_RNG) = True Then
    VARIABLES_VECTOR = VARIABLES_RNG
    If UBound(VARIABLES_VECTOR, 1) = 1 Then
        VARIABLES_VECTOR = MATRIX_TRANSPOSE_FUNC(VARIABLES_VECTOR)
    End If
Else
    ReDim VARIABLES_VECTOR(1 To 1, 1 To 1)
    VARIABLES_VECTOR(1, 1) = VARIABLES_RNG
End If
NROWS = UBound(VARIABLES_VECTOR, 1)
If IsArray(PARAMETERS_RNG) = True Then
    PARAMETERS_VECTOR = PARAMETERS_RNG
    If UBound(PARAMETERS_VECTOR, 1) = 1 Then
        PARAMETERS_VECTOR = MATRIX_TRANSPOSE_FUNC(PARAMETERS_VECTOR)
    End If
Else
    ReDim PARAMETERS_VECTOR(1 To 1, 1 To 1)
    PARAMETERS_VECTOR(1, 1) = PARAMETERS_RNG
End If
If NROWS <> UBound(PARAMETERS_VECTOR, 1) Then: GoTo ERROR_LABEL

GoSub PARSER_LINE
If Not OK_FLAG Then
    ERROR_STR = PARSER_OBJ.ErrorDescription
    GoSub ERROR_LABEL
Else
    For i = 1 To NROWS: PARSER_OBJ.variable(CStr(VARIABLES_VECTOR(i, 1))) = CDbl(PARAMETERS_VECTOR(i, 1)): Next i
    MATH_PARSER_EVAL2_FUNC = PARSER_OBJ.EvalMulti(XDATA_ARR, XVAR_STR) 'f(x): evaluate function value
    'Debug.Print "x=" + CStr(X_VAL); Tab(25); "f(x)=" + CStr(Y_VAL)
End If

Exit Function
'----------------------------------------------------------------------------------------------------------------
PARSER_LINE:
'----------------------------------------------------------------------------------------------------------------
    If PARSER_OBJ.Expression <> FORMULA_STR Then
        OK_FLAG = PARSER_OBJ.StoreExpression(FORMULA_STR) 'parse function
        'the unit of measure for angle computing (RAD (default), DEG or GRAD)
        If ANGLE_DEG_FLAG = True Then: PARSER_OBJ.AngleUnit = "DEG"
        PARSER_OBJ.OpAssignExplicit = EXPLICIT_VAR_FLAG 'If True all variables must be assigned
    Else
        OK_FLAG = True 'Already Parsed
'        Debug.Print "NICO"
    End If
'----------------------------------------------------------------------------------------------------------------
Return
'----------------------------------------------------------------------------------------------------------------
ERROR_LABEL:
MATH_PARSER_EVAL2_FUNC = ERROR_STR
End Function

Function MATH_PARSER_EVAL3_FUNC(ByRef FUNCTION_RNG As Variant, _
ByRef DOMAIN_RNG As Variant, _
ByRef XDATA_ARR As Variant, _
Optional ByRef ERROR_STR As String = "--")

Dim i As Long
Dim j As Long

Dim NROWS As Long
Dim NCOLUMNS As Long

Dim X_VAL As Double
Dim D_VAL As Double
Dim F_VAL As Double

Dim OK_FLAG As Boolean

Dim DOMAIN_VECTOR As Variant
Dim FUNCTION_VECTOR As Variant

On Error GoTo ERROR_LABEL

If IsArray(FUNCTION_RNG) = True Then
    FUNCTION_VECTOR = FUNCTION_RNG
    If UBound(FUNCTION_VECTOR, 1) = 1 Then
        FUNCTION_VECTOR = MATRIX_TRANSPOSE_FUNC(FUNCTION_VECTOR)
    End If
Else
    ReDim FUNCTION_VECTOR(1 To 1, 1 To 1)
    FUNCTION_VECTOR(1, 1) = FUNCTION_RNG
End If
NCOLUMNS = UBound(FUNCTION_VECTOR, 1)

If IsArray(DOMAIN_RNG) = True Then
    DOMAIN_VECTOR = DOMAIN_RNG
    If UBound(DOMAIN_VECTOR, 1) = 1 Then
        DOMAIN_VECTOR = MATRIX_TRANSPOSE_FUNC(DOMAIN_VECTOR)
    End If
Else
    ReDim DOMAIN_VECTOR(1 To 1, 1 To 1)
    DOMAIN_VECTOR(1, 1) = DOMAIN_RNG
End If
If UBound(FUNCTION_VECTOR, 1) <> UBound(DOMAIN_VECTOR, 1) Then: GoTo ERROR_LABEL

ReDim PARSER0_OBJ(1 To NCOLUMNS) As New clsMathParser
ReDim PARSER1_OBJ(1 To NCOLUMNS) As New clsMathParser

For j = 1 To NCOLUMNS
    OK_FLAG = PARSER0_OBJ(j).StoreExpression(CStr(FUNCTION_VECTOR(j, 1)))
    If Not OK_FLAG Then
        ERROR_STR = "Function" + CStr(j) & ", " & PARSER0_OBJ(j).ErrorDescription 'Err.Raise 513,
    End If
    OK_FLAG = PARSER1_OBJ(j).StoreExpression(CStr(DOMAIN_VECTOR(j, 1)))
    If Not OK_FLAG Then
        ERROR_STR = "Domain" + CStr(j) & ", " & PARSER1_OBJ(j).ErrorDescription   'Err.Raise 514
    End If
Next j

NROWS = UBound(XDATA_ARR) - LBound(XDATA_ARR) + 1
ReDim TEMP_VECTOR(1 To NROWS, 1 To 2)
'TEMP_VECTOR(0, 1) = "x": TEMP_VECTOR(0, 2) = "f(x)"
For i = LBound(XDATA_ARR) To UBound(XDATA_ARR)
    For j = 1 To NCOLUMNS
        X_VAL = XDATA_ARR(i)
        D_VAL = PARSER1_OBJ(j).Eval1(X_VAL)
        If Err.number <> 0 Then GoTo ERROR_LABEL
        If D_VAL = 1 Then
            F_VAL = PARSER0_OBJ(j).Eval1(X_VAL)
            If Err.number <> 0 Then GoTo ERROR_LABEL
            Exit For
        End If
    Next j
    'Debug.Print "x=" + Str(X_VAL); Tab(25); "f(x)=" + Str(F_VAL)
    TEMP_VECTOR(i, 1) = X_VAL: TEMP_VECTOR(i, 2) = F_VAL
Next i
MATH_PARSER_EVAL3_FUNC = TEMP_VECTOR

Exit Function
ERROR_LABEL:
MATH_PARSER_EVAL3_FUNC = ERROR_STR ' Err.source, Err.Description
End Function

'//PERFECT

Function MATH_PARSER_EVAL4_FUNC(ByVal FORMULA_STR As String, _
ByRef VARIABLES_RNG As Variant, _
ByRef PARAMETERS_RNG As Variant, _
ByRef XDATA_RNG As Variant, _
Optional ByVal XVAR_STR As String = "x", _
Optional ByVal ANGLE_DEG_FLAG As Boolean = False, _
Optional ByVal EXPLICIT_VAR_FLAG As Boolean = False, _
Optional ByRef ERROR_STR As String = "--", _
Optional ByRef PARSER_OBJ As clsMathParser)

Dim i As Long
Dim j As Long

Dim NROWS As Long
Dim NCOLUMNS As Long
Dim NSIZE As Long

Dim CHR_STR As String
Dim OK_FLAG As Boolean

Dim TEMP_VECTOR As Variant
Dim XDATA_MATRIX As Variant
Dim PARAMETERS_VECTOR As Variant
Dim VARIABLES_VECTOR As Variant

On Error GoTo ERROR_LABEL

If PARSER_OBJ Is Nothing Then: Set PARSER_OBJ = New clsMathParser

XDATA_MATRIX = XDATA_RNG
NROWS = UBound(XDATA_MATRIX, 1)
NCOLUMNS = UBound(XDATA_MATRIX, 2)

If IsArray(VARIABLES_RNG) = True Then
    VARIABLES_VECTOR = VARIABLES_RNG
    If UBound(VARIABLES_VECTOR, 1) = 1 Then
        VARIABLES_VECTOR = MATRIX_TRANSPOSE_FUNC(VARIABLES_VECTOR)
    End If
Else
    ReDim VARIABLES_VECTOR(1 To 1, 1 To 1)
    VARIABLES_VECTOR(1, 1) = VARIABLES_RNG
End If
NSIZE = UBound(VARIABLES_VECTOR, 1)
If IsArray(PARAMETERS_RNG) = True Then
    PARAMETERS_VECTOR = PARAMETERS_RNG
    If UBound(PARAMETERS_VECTOR, 1) = 1 Then
        PARAMETERS_VECTOR = MATRIX_TRANSPOSE_FUNC(PARAMETERS_VECTOR)
    End If
Else
    ReDim PARAMETERS_VECTOR(1 To 1, 1 To 1)
    PARAMETERS_VECTOR(1, 1) = PARAMETERS_RNG
End If
If NSIZE <> UBound(PARAMETERS_VECTOR, 1) Then: GoTo ERROR_LABEL

GoSub PARSER_LINE
If Not OK_FLAG Then
    ERROR_STR = PARSER_OBJ.ErrorDescription
    GoTo ERROR_LABEL 'Err.Raise 1001, , PARSER_OBJ.ErrorDescription
End If
If PARSER_OBJ.VarTop <> (NCOLUMNS + NSIZE) Then: GoTo ERROR_LABEL

For i = 1 To NSIZE
    CHR_STR = VARIABLES_VECTOR(i, 1)
    PARSER_OBJ.variable(CHR_STR) = PARAMETERS_VECTOR(i, 1)
Next i

ReDim TEMP_VECTOR(1 To NROWS, 1 To 1)
For i = 1 To NROWS
    For j = 1 To NCOLUMNS
        CHR_STR = CStr(XVAR_STR & j)
        PARSER_OBJ.variable(CHR_STR) = XDATA_MATRIX(i, j)
    Next j
    TEMP_VECTOR(i, 1) = PARSER_OBJ.Eval() 'f(x,y)
Next i

MATH_PARSER_EVAL4_FUNC = TEMP_VECTOR

'----------------------------------------------------------------------------------------------------------------
Exit Function
'----------------------------------------------------------------------------------------------------------------
PARSER_LINE:
'----------------------------------------------------------------------------------------------------------------
    If PARSER_OBJ.Expression <> FORMULA_STR Then
        OK_FLAG = PARSER_OBJ.StoreExpression(FORMULA_STR) 'parse function
        'the unit of measure for angle computing (RAD (default), DEG or GRAD)
        If ANGLE_DEG_FLAG = True Then: PARSER_OBJ.AngleUnit = "DEG"
        PARSER_OBJ.OpAssignExplicit = EXPLICIT_VAR_FLAG 'If True all variables must be assigned
    Else
        OK_FLAG = True 'Already Parsed
'        Debug.Print "NICO"
    End If
'----------------------------------------------------------------------------------------------------------------
Return
'----------------------------------------------------------------------------------------------------------------
ERROR_LABEL:
MATH_PARSER_EVAL4_FUNC = ERROR_STR
End Function

'//PERFECT

Function MATH_PARSER_VARIABLES_FUNC(ByVal FORMULA_STR As String, _
Optional ByRef PARSER_OBJ As clsMathParser, _
Optional ByRef ERROR_STR As String = "--")

'TEMP_VECTOR = MATH_PARSER_VARIABLES_FUNC("(a-b)*(a-b)+30*time/y")
'TEMP_VECTOR = MATH_PARSER_VARIABLES_FUNC("(x^2+1)/(y^2-y+2)")
'TEMP_VECTOR = MATH_PARSER_VARIABLES_FUNC("(x<=0)* x^2 + (0<x<=1)* Log(x+1) + (x>1)* Sqr(x-Log(2))")
Dim i As Long
Dim j As Long
Dim OK_FLAG As Boolean
Dim VARIABLES_VECTOR() As String
On Error GoTo ERROR_LABEL
'----------------------------------------------------------------------
If PARSER_OBJ Is Nothing Then: Set PARSER_OBJ = New clsMathParser
'Define expression, perform syntax check and detect all variables
OK_FLAG = PARSER_OBJ.StoreExpression(FORMULA_STR)
If Not OK_FLAG Then
    ERROR_STR = PARSER_OBJ.ErrorDescription
    GoTo ERROR_LABEL
End If
'----------------------------------------------------------------------
j = PARSER_OBJ.VarTop
ReDim VARIABLES_VECTOR(1 To j, 1 To 1)
For i = 1 To j
    VARIABLES_VECTOR(i, 1) = PARSER_OBJ.VarName(i) '& ": " & CStr(i) + "° variable"
    'Debug.Print PARSER_OBJ.VarName(i); " = "; PARSER_OBJ.variable(i) --> Coefficients
Next i

MATH_PARSER_VARIABLES_FUNC = VARIABLES_VECTOR

Exit Function
ERROR_LABEL:
MATH_PARSER_VARIABLES_FUNC = ERROR_STR
End Function

'//PERFECT

Private Function MATH_PARSER_CHECK_FUNC(ByRef DATA_ARR() As Variant, _
ByRef PARSER_OBJ As clsMathParser)

Dim i As Long
Dim j As Long

On Error GoTo ERROR_LABEL

Debug.Print "Formula:= "; PARSER_OBJ.Expression
Debug.Print "Result:= "; PARSER_OBJ.Eval
PARSER_OBJ.ET_Dump DATA_ARR '<<< array table returned
For i = LBound(DATA_ARR, 1) To UBound(DATA_ARR, 1)
    If i > 0 Then Debug.Print i, Else Debug.Print "Id",
    For j = LBound(DATA_ARR, 2) To UBound(DATA_ARR, 2)
        Debug.Print DATA_ARR(i, j),
    Next j
    Debug.Print ""
Next i
MATH_PARSER_CHECK_FUNC = True

Exit Function
ERROR_LABEL:
MATH_PARSER_CHECK_FUNC = False
End Function

'// PERFECT

Function MATH_PARSER_MCLAURIN_SERIE_EXAMPLE_FUNC( _
Optional ByVal X0_VAL As Double = 0.5, _
Optional ByVal NMAX_VAL As Long = 16, _
Optional ByVal ERROR_STR As String = "--", _
Optional ByRef PARSER_OBJ As clsMathParser)

'This example computes Mc Lauren's series up to 16° order for exp(x) with x=0.5
'1.64872127070013

'X0_VAL --> set value of Taylor's series expansion
'NMAX_VAL --> set max series expansion

Dim i As Long
Dim Y_VAL As Double
Dim OK_FLAG As Boolean
Dim FORMULA_STR As String

On Error GoTo ERROR_LABEL

'----------------------------------------------------------------------
'Define expression, perform syntax check and get its handle
If PARSER_OBJ Is Nothing Then: Set PARSER_OBJ = New clsMathParser
FORMULA_STR = "x^n / n!" 'Expression to evaluate. Has two variable
OK_FLAG = PARSER_OBJ.StoreExpression(FORMULA_STR)
If Not OK_FLAG Then
    ERROR_STR = PARSER_OBJ.ErrorDescription
    GoTo ERROR_LABEL
End If
'begin formula evaluation -------------------------
PARSER_OBJ.variable(1) = X0_VAL 'load value x
For i = 0 To NMAX_VAL
    PARSER_OBJ.variable(2) = i 'increments the i variables
    If Err Then GoTo ERROR_LABEL
    Y_VAL = Y_VAL + PARSER_OBJ.Eval 'accumulates partial i-term
Next i
MATH_PARSER_MCLAURIN_SERIE_EXAMPLE_FUNC = Y_VAL

Exit Function
ERROR_LABEL:
MATH_PARSER_MCLAURIN_SERIE_EXAMPLE_FUNC = ERROR_STR
End Function

'//PERFECT

Private Sub TEST_MATH_PARSER_CHECK_FUNC()

Dim TEMP_ARR As Variant

Static PARSER_OBJ As clsMathParser
'If PARSER_OBJ Is Nothing Then: Set PARSER_OBJ = New clsMathParser
Set PARSER_OBJ = New clsMathParser

Debug.Print "---------------------------------------------------------------------------------------------------------"
Debug.Print "Formula1: 3y1+2x-y3"
Debug.Print "---------------------------------------------------------------------------------------------------------"
TEMP_ARR = ROMBERG_INTEGRATION_FUNC("sin(2*pi*x)+cos(2*pi*x)", 0, 0.5)
Debug.Print TEMP_ARR(1) ' 0.318309886183791
Debug.Print TEMP_ARR(2) '64
Debug.Print TEMP_ARR(3) ' 5.55111512312578E-17
Debug.Print "---------------------------------------------------------------------------------------------------------"
Debug.Print "Formula2: x^n / n!"
Debug.Print "---------------------------------------------------------------------------------------------------------"
Debug.Print MATH_PARSER_MCLAURIN_SERIE_EXAMPLE_FUNC(0.5, 16, "--", PARSER_OBJ)
Debug.Print "---------------------------------------------------------------------------------------------------------"
Debug.Print "Formula3: x ^ 2 + y" 'f(x, y)
'x = -3
'y = 0
'f(x, y) = 9
Debug.Print "---------------------------------------------------------------------------------------------------------"
Debug.Print MATH_PARSER_EVAL0_FUNC("x ^ 2 + y", -3, 0)
Debug.Print "---------------------------------------------------------------------------------------------------------"
ReDim TEMP_ARR(1 To 4, 1 To 1)
TEMP_ARR(1, 1) = 3: TEMP_ARR(2, 1) = 4
TEMP_ARR(3, 1) = 5: TEMP_ARR(4, 1) = 6
Debug.Print "Formula4: 3y1+2x-y3"
Debug.Print "---------------------------------------------------------------------------------------------------------"
Debug.Print MATH_PARSER_EVAL1_FUNC("3y1+2x-y3", TEMP_ARR, False, False, , PARSER_OBJ)
Debug.Print "---------------------------------------------------------------------------------------------------------"
Debug.Print "Formula5: (a^2*exp(y/T)-x)"
Debug.Print "---------------------------------------------------------------------------------------------------------"
Debug.Print MATH_PARSER_EVAL1_FUNC("(a^2*exp(y/T)-x)", TEMP_ARR, False, False, , PARSER_OBJ)
Debug.Print "---------------------------------------------------------------------------------------------------------"
Debug.Print "Formula6: (a^2*exp(y/T)-x)"
Debug.Print "---------------------------------------------------------------------------------------------------------"
Dim DATA_ARR() As Variant
ReDim DATA_ARR(1 To 1)
Debug.Print MATH_PARSER_CHECK_FUNC(DATA_ARR, PARSER_OBJ)
Debug.Print "---------------------------------------------------------------------------------------------------------"
Debug.Print "Formula7: x^3-x^2+3*x+6"
Debug.Print "---------------------------------------------------------------------------------------------------------"
Debug.Print MATH_PARSER_EVAL1_FUNC("x^3-x^2+3*x+6", 4, False, False, , PARSER_OBJ)
'For the example above, with a Pentium III° at 1.2 GHz, we have got 1000 points of the cubic polynomial in less
'than 15 ms (1.4E-2), which is very good performance for this kind of parser (70.000 points / sec)
'Note also that the variable name it is not important; the example also works fine with other strings, such as:
'“t^3-t^2+3*t+6” , “a^3-a^2+3*a+6”, etc.
'The parser, simply substitutes the first variables found with the passed value.
Debug.Print "---------------------------------------------------------------------------------------------------------"
Debug.Print "Formula8: (x^2+1)/(x^2-1)"
Debug.Print "---------------------------------------------------------------------------------------------------------"
Debug.Print MATH_PARSER_EVAL1_FUNC("(x^2+1)/(x^2-1)", 3, False, False, , PARSER_OBJ)
Debug.Print "---------------------------------------------------------------------------------------------------------"
'Piecewise function definition
' x^2 for x<=0
' log(x+1) for 0<x<=1
' sqr(x-log2) for x>1
'"(x<=0)* x^2 + (0<x<=1)* Log(x+1) + (x>1)* Sqr(x-Log(2))"
Debug.Print "Formula9: (x<=0)* x^2 + (0<x<=1)* Log(x+1) + (x>1)* Sqr(x-Log(2))"
'Thanks to the Conditioned Branch algorithm – it is also possible to evaluate a piecewise
'expression directly and in a very short way. Look how compact the code is in this case.
Debug.Print "---------------------------------------------------------------------------------------------------------"
Debug.Print MATH_PARSER_EVAL1_FUNC("(x<=0)* x^2 + (0<x<=1)* Log(x+1) + (x>1)* Sqr(x-Log(2))", 3, False, False, , PARSER_OBJ)
Debug.Print "---------------------------------------------------------------------------------------------------------"
Dim i, nLOOPS As Long
Dim H_VAL As Double
Dim X0_VAL As Double
Dim X1_VAL As Double

Dim PARAMETERS_VECTOR(1 To 3, 1 To 1)
Dim VARIABLES_VECTOR(1 To 3, 1 To 1)

PARAMETERS_VECTOR(1, 1) = 0.123
PARAMETERS_VECTOR(2, 1) = 0.4
PARAMETERS_VECTOR(3, 1) = 100

VARIABLES_VECTOR(1, 1) = "y"
VARIABLES_VECTOR(2, 1) = "t"
VARIABLES_VECTOR(3, 1) = "a"
nLOOPS = 10000: X0_VAL = 0: X1_VAL = 1
H_VAL = (X1_VAL - X0_VAL) / nLOOPS
Dim XDATA_ARR() As Double
ReDim XDATA_ARR(1 To nLOOPS) 'load x-samples
For i = 1 To nLOOPS
    XDATA_ARR(i) = i * H_VAL + X0_VAL
Next i
Debug.Print "Formula10: a^2*exp(x/T)-y"
Debug.Print "---------------------------------------------------------------------------------------------------------"
TEMP_ARR = MATH_PARSER_EVAL2_FUNC("a^2*exp(x/T)-y", VARIABLES_VECTOR, PARAMETERS_VECTOR, XDATA_ARR, "x", False, False, , PARSER_OBJ)
For i = LBound(TEMP_ARR) To 5 'UBound(TEMP_ARR)
    Debug.Print TEMP_ARR(i)
Next i

Debug.Print "---------------------------------------------------------------------------------------------------------"
Debug.Print "Formula11: (x<=0)* x^2 + (0<x<=1)* Log(x+1) + (x>1)* Sqr(x-Log(2))"
Debug.Print "---------------------------------------------------------------------------------------------------------"
'Piecewise function definition ---------------------------------------------------------------------------------------
' x^2 for x<=0
' log(x+1) for 0<x<=1
' sqr(x-log2) for x>1
'---------------------------------------------------------------------------------------------------------------------
X0_VAL = -2: X1_VAL = 2: nLOOPS = 10
ReDim FUNCTION_VECTOR(1 To 3, 1 To 1)
ReDim DOMAIN_VECTOR(1 To 3, 1 To 1)
FUNCTION_VECTOR(1, 1) = "x^2"
DOMAIN_VECTOR(1, 1) = "x<=0"
FUNCTION_VECTOR(2, 1) = "log(x+1)"
DOMAIN_VECTOR(2, 1) = "(0<x)*(x<=1)"
FUNCTION_VECTOR(3, 1) = "sqr(x-log(2))"
DOMAIN_VECTOR(3, 1) = "sqr(x-log(2))"
H_VAL = (X1_VAL - X0_VAL) / (nLOOPS - 1)
ReDim TEMP_ARR(1 To nLOOPS)
For i = 1 To nLOOPS
    TEMP_ARR(i) = X0_VAL + (i - 1) * H_VAL 'value x
Next i
TEMP_ARR = MATH_PARSER_EVAL3_FUNC(FUNCTION_VECTOR, DOMAIN_VECTOR, TEMP_ARR, "x")
For i = LBound(TEMP_ARR) To UBound(TEMP_ARR)
    Debug.Print "x=" + Str(TEMP_ARR(i, 1)); Tab(25); "f(x)=" + Str(TEMP_ARR(i, 2))
Next i
Debug.Print "---------------------------------------------------------------------------------------------------------"
End Sub
