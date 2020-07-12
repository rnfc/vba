Attribute VB_Name = "NUMBER_REAL_LOGARITHM_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.

'************************************************************************************
'************************************************************************************
'FUNCTION      : LN_FUNC
'DESCRIPTION   : Returns the natural logarithm of a number. Natural
'logarithms are based on the constant e (2.71828182845904).
'LIBRARY       : NUMBER_REAL
'GROUP         : LOGARITHM
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/08/2008
'************************************************************************************
'************************************************************************************

Function LN_FUNC(ByVal X_VAL As Variant)

Dim i As Long
Dim j As Long

Dim ETEMP_VAL As Variant
Dim PTEMP_VAL As Variant
Dim TTEMP_VAL As Variant
Dim UTEMP_VAL As Variant
Dim YTEMP_VAL As Variant

Dim LN_FACTOR As Variant
Dim MULT_FACTOR As Variant

On Error GoTo ERROR_LABEL

LN_FACTOR = CDec("0.69314718055994530941723212145")

If X_VAL = 1 Then LN_FUNC = 0: Exit Function
If X_VAL <= 0 Then GoTo ERROR_LABEL

i = Int(Log(X_VAL) / Log(2))
MULT_FACTOR = 2 ^ Abs(i) 'potenza in precisione estesa.

If i < 0 Then
      TTEMP_VAL = CDec(X_VAL * MULT_FACTOR)
Else: TTEMP_VAL = CDec(X_VAL / MULT_FACTOR)
End If

UTEMP_VAL = (TTEMP_VAL - 1) / (TTEMP_VAL + 1)
PTEMP_VAL = UTEMP_VAL
YTEMP_VAL = UTEMP_VAL
UTEMP_VAL = UTEMP_VAL * UTEMP_VAL

For j = 3 To 1000 Step 2
    PTEMP_VAL = PTEMP_VAL * UTEMP_VAL
    ETEMP_VAL = PTEMP_VAL / j
    YTEMP_VAL = YTEMP_VAL + ETEMP_VAL
    If Abs(ETEMP_VAL) <= 10 ^ -30 Then Exit For
Next j
YTEMP_VAL = 2 * YTEMP_VAL + i * LN_FACTOR
LN_FUNC = YTEMP_VAL

Exit Function
ERROR_LABEL:
LN_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : LOG_FUNC
'DESCRIPTION   : Returns the logarithm of a number to the base you specify
'LIBRARY       : NUMBER_REAL
'GROUP         : LOGARITHM
'ID            : 002
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/08/2008
'************************************************************************************
'************************************************************************************

Function LOG_FUNC(ByVal X_VAL As Variant, _
Optional ByVal BASE_VAL As Variant = 10)

Dim ATEMP_VAL As Variant
Dim BTEMP_VAL As Variant

On Error GoTo ERROR_LABEL

If BASE_VAL = 0 Then: Exit Function
If X_VAL = 1 Then LOG_FUNC = 0: Exit Function
If X_VAL <= 0 Then GoTo ERROR_LABEL

Select Case BASE_VAL
    Case 10: BASE_VAL = CDec("2.30258509299404568401799145467")
    Case 2:  BASE_VAL = CDec("0.69314718055994530941723212145")
    Case Else: BASE_VAL = LN_FUNC(BASE_VAL)
End Select

ATEMP_VAL = LN_FUNC(X_VAL)
BTEMP_VAL = (ATEMP_VAL / BASE_VAL)

LOG_FUNC = BTEMP_VAL

Exit Function
ERROR_LABEL:
LOG_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : EXP_FUNC
'DESCRIPTION   : Returns e raised to the power of number. The constant e
'equals 2.71828182845904, the base of the natural logarithm
'LIBRARY       : NUMBER_REAL
'GROUP         : LOGARITHM
'ID            : 003
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/08/2008
'************************************************************************************
'************************************************************************************

Function EXP_FUNC(ByVal X_VAL As Variant)

Dim h As Long
Dim i As Long
Dim j As Long
Dim k As Long

Dim ATEMP_VAL As Variant
Dim BTEMP_VAL As Variant
Dim CTEMP_VAL As Variant
Dim DTEMP_VAL As Variant
Dim ETEMP_VAL As Variant
Dim FTEMP_VAL As Variant

Dim LN_VALUE As Variant

Dim tolerance As Variant

On Error GoTo ERROR_LABEL

tolerance = 10 ^ -30

If X_VAL = 0 Then EXP_FUNC = 1: Exit Function
If X_VAL < 0 Then ATEMP_VAL = -X_VAL Else ATEMP_VAL = X_VAL

LN_VALUE = CDec("0.69314718055994530941723212145")
h = Int(ATEMP_VAL / LN_VALUE)

BTEMP_VAL = ATEMP_VAL - h * LN_VALUE

j = 0
Do Until Abs(BTEMP_VAL) < 0.1
    BTEMP_VAL = BTEMP_VAL / 2
    j = j + 1
Loop

''start Taylor's sum
CTEMP_VAL = 1 + BTEMP_VAL
DTEMP_VAL = CTEMP_VAL
ETEMP_VAL = BTEMP_VAL

For k = 2 To 1000
    ETEMP_VAL = ETEMP_VAL * BTEMP_VAL / k
    CTEMP_VAL = CTEMP_VAL + ETEMP_VAL
    If ETEMP_VAL <= tolerance Then Exit For
    DTEMP_VAL = CTEMP_VAL
Next k
'
For i = 1 To j
    CTEMP_VAL = CTEMP_VAL * CTEMP_VAL
Next i

If h > 0 Then
    FTEMP_VAL = 2 ^ h
    CTEMP_VAL = CTEMP_VAL * FTEMP_VAL
End If
'
If X_VAL < 0 Then CTEMP_VAL = 1 / CTEMP_VAL
EXP_FUNC = CTEMP_VAL

Exit Function
ERROR_LABEL:
EXP_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : POWER_FUNC
'DESCRIPTION   : Returns the result of a number raised to a power
'LIBRARY       : NUMBER_REAL
'GROUP         : LOGARITHM
'ID            : 004
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/08/2008
'************************************************************************************
'************************************************************************************

Function POWER_FUNC(ByVal BASE_VAL As Variant, _
ByVal EXPONENT_VAL As Variant)

'BASE_VAL: is the base number. It can be any real number.
'EXPONENT_VAL: is the exponent to which the base number is raised.

On Error GoTo ERROR_LABEL

POWER_FUNC = EXP_FUNC(EXPONENT_VAL * LN_FUNC(BASE_VAL))

Exit Function
ERROR_LABEL:
POWER_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : EXPONENTIAL_ASYMPTOTIC_EXPANSION_FUNC
'DESCRIPTION   : Calculates exponential for n=1 using asymptotic expansion
'LIBRARY       : NUMBER_REAL
'GROUP         : LOGARITHM
'ID            : 005
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/08/2008
'************************************************************************************
'************************************************************************************

Function EXPONENTIAL_ASYMPTOTIC_EXPANSION_FUNC(ByVal X_VAL As Double, _
Optional ByVal epsilon As Double = 2 ^ (-52))
  
  Dim i As Long
  Dim j As Long
  
  Dim TEMP_SUM As Double
  Dim TEMP_MULT As Double
  Dim TEMP_DIFF As Double
  
  Dim FACT_VAL As Double
  Dim DELTA_VAL As Double
  
  On Error GoTo ERROR_LABEL
  
  If X_VAL <= 0 Then: GoTo ERROR_LABEL 'not defined for x<0"
  
  j = 100
  FACT_VAL = (-Log(X_VAL) - 0.5772156649)
  
  TEMP_SUM = 0
  TEMP_MULT = 1
  
  TEMP_DIFF = -1
  DELTA_VAL = 0
  
  For i = 1 To j
    TEMP_SUM = TEMP_SUM + ((-X_VAL) ^ i) / (i * TEMP_MULT)
    TEMP_MULT = TEMP_MULT * (i + 1)
    TEMP_DIFF = FACT_VAL - TEMP_SUM
    If (Abs(DELTA_VAL - TEMP_DIFF) < epsilon) Then
      EXPONENTIAL_ASYMPTOTIC_EXPANSION_FUNC = TEMP_DIFF
      Exit Function
    End If
    DELTA_VAL = TEMP_DIFF
  Next i

Exit Function
ERROR_LABEL:
EXPONENTIAL_ASYMPTOTIC_EXPANSION_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : LOG_DIFFERENCE_FUNC
'DESCRIPTION   : Log Difference
'LIBRARY       : NUMBER_REAL
'GROUP         : LOGARITHM
'ID            : 006
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/08/2008
'************************************************************************************
'************************************************************************************

Function LOG_DIFFERENCE_FUNC(ByVal X_VAL As Double, _
Optional ByVal Y_VAL As Double = 1, _
Optional ByVal THRESHOLD As Double = 0.5)

If Abs((X_VAL - Y_VAL) / Y_VAL) >= THRESHOLD Then
    LOG_DIFFERENCE_FUNC = Log(X_VAL / Y_VAL)
Else
    LOG_DIFFERENCE_FUNC = LOG_PLUS_FUNC(X_VAL)
End If

Exit Function
ERROR_LABEL:
LOG_DIFFERENCE_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : LOG_PLUS_FUNC
'DESCRIPTION   : Accurate calculation of log(1+x), particularly for small x
'LIBRARY       : NUMBER_REAL
'GROUP         : LOGARITHM
'ID            : 007
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/08/2008
'************************************************************************************
'************************************************************************************

Function LOG_PLUS_FUNC(ByVal X_VAL As Double)
   
   Dim Z_VAL As Double
   
   On Error GoTo ERROR_LABEL
   
   If (Abs(X_VAL) > 0.5) Then
      LOG_PLUS_FUNC = Log(1# + X_VAL)
   Else
     Z_VAL = X_VAL / (2# + X_VAL)
     
     LOG_PLUS_FUNC = 2# * Z_VAL * LOG_SERIE_FUNC(Z_VAL * _
                Z_VAL, 1#, 2#)
   End If

Exit Function
ERROR_LABEL:
LOG_PLUS_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : LOG_PLUS_MINUS_FUNC
'DESCRIPTION   : Accurate calculation of log(1+x)-x, particularly for small x
'LIBRARY       : NUMBER_REAL
'GROUP         : LOGARITHM
'ID            : 008
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/08/2008
'************************************************************************************
'************************************************************************************

Function LOG_PLUS_MINUS_FUNC(ByVal X_VAL As Double)
   
   Dim Z_VAL As Double
   Dim Y_VAL As Double
   
   If (Abs(X_VAL) < 0.01) Then
      Z_VAL = X_VAL / (2# + X_VAL)
      Y_VAL = Z_VAL * Z_VAL
      LOG_PLUS_MINUS_FUNC = Z_VAL * ((((2# / 9# * Y_VAL + 2# / 7#) * _
      Y_VAL + 0.4) * Y_VAL + 2# / 3#) * Y_VAL - X_VAL)
   ElseIf (X_VAL < -0.79149064 Or X_VAL > 1#) Then
      LOG_PLUS_MINUS_FUNC = Log(1# + X_VAL) - X_VAL
   Else
      Z_VAL = X_VAL / (2# + X_VAL)
      Y_VAL = Z_VAL * Z_VAL
      LOG_PLUS_MINUS_FUNC = Z_VAL * (2# * Y_VAL * _
        LOG_SERIE_FUNC(Y_VAL, 3#, 2#) - X_VAL)
   End If
   
Exit Function
ERROR_LABEL:
LOG_PLUS_MINUS_FUNC = Err.number
End Function



'************************************************************************************
'************************************************************************************
'FUNCTION      : EXP_MINUS_FUNC
'DESCRIPTION   : Accurate calculation of exp(x)-1, particularly for small x
'LIBRARY       : NUMBER_REAL
'GROUP         : LOGARITHM
'ID            : 009
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/08/2008
'************************************************************************************
'************************************************************************************

Function EXP_MINUS_FUNC(ByVal X_VAL As Double)

     Dim ATEMP_VAL As Double
     Dim BTEMP_VAL As Double
     
     Dim CTEMP_VAL As Double
     Dim DTEMP_VAL As Double
     
     Dim ETEMP_VAL As Double
     Dim FTEMP_VAL As Double

     Dim tolerance As Double
     
     On Error GoTo ERROR_LABEL
  
    tolerance = 0.00000000000001
'---------------------------------------------------------------------------------------------
'----------used for rescaling calcs w/o impacting accuracy, to avoid over/underflow-----------
'                    8.6361685550944446253863518628004e-78 = 2^-256
'                    1.1579208923731619542357098500869e+77 = 2^256
'---------------------------------------------------------------------------------------------
  
  If (Abs(X_VAL) < 2) Then
     
     ATEMP_VAL = 24#
     CTEMP_VAL = 2# * (12# - X_VAL * (6# - X_VAL))
     FTEMP_VAL = X_VAL * X_VAL * 0.25
     BTEMP_VAL = 8# * (15# + FTEMP_VAL)
     DTEMP_VAL = 120# - X_VAL * (60# - X_VAL * (12# - X_VAL))
     ETEMP_VAL = 7#

     Do While ((Abs(BTEMP_VAL * CTEMP_VAL - ATEMP_VAL * DTEMP_VAL) > Abs(tolerance * CTEMP_VAL * BTEMP_VAL)))

       ATEMP_VAL = ETEMP_VAL * BTEMP_VAL + FTEMP_VAL * ATEMP_VAL
       CTEMP_VAL = ETEMP_VAL * DTEMP_VAL + FTEMP_VAL * CTEMP_VAL
       ETEMP_VAL = ETEMP_VAL + 2#

       BTEMP_VAL = ETEMP_VAL * ATEMP_VAL + FTEMP_VAL * BTEMP_VAL
       DTEMP_VAL = ETEMP_VAL * CTEMP_VAL + FTEMP_VAL * DTEMP_VAL
       ETEMP_VAL = ETEMP_VAL + 2#
       
       If (DTEMP_VAL > 1.15792089237316E+77) Then
             ATEMP_VAL = ATEMP_VAL * 8.63616855509444E-78
             CTEMP_VAL = CTEMP_VAL * 8.63616855509444E-78
             BTEMP_VAL = BTEMP_VAL * 8.63616855509444E-78
             DTEMP_VAL = DTEMP_VAL * 8.63616855509444E-78
       End If
       
     Loop

     EXP_MINUS_FUNC = X_VAL * BTEMP_VAL / DTEMP_VAL
  Else
     EXP_MINUS_FUNC = Exp(X_VAL) - 1#
  End If
  
Exit Function
ERROR_LABEL:
EXP_MINUS_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : LOG_SERIE_FUNC
'DESCRIPTION   : Continued fraction for calculation of 1/i + x/(i+d) + x*x/(i+2*d) +
'x*x*x/(i+3d) + ...
'LIBRARY       : NUMBER_REAL
'GROUP         : LOGARITHM
'ID            : 010
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/08/2008
'************************************************************************************
'************************************************************************************

Function LOG_SERIE_FUNC(ByVal X_VAL As Double, _
ByVal I_VAL As Double, _
ByVal D_VAL As Double)

Dim ATEMP_VAL As Double
Dim BTEMP_VAL As Double

Dim CTEMP_VAL As Double
Dim DTEMP_VAL As Double

Dim ETEMP_VAL As Double
Dim FTEMP_VAL As Double
Dim GTEMP_VAL As Double
Dim HTEMP_VAL As Double

Dim tolerance As Double
     
On Error GoTo ERROR_LABEL

tolerance = 0.000000000000001

ETEMP_VAL = 2# * D_VAL
FTEMP_VAL = I_VAL + D_VAL
HTEMP_VAL = FTEMP_VAL + D_VAL
ATEMP_VAL = FTEMP_VAL
CTEMP_VAL = I_VAL * (FTEMP_VAL - I_VAL * X_VAL)
DTEMP_VAL = D_VAL * D_VAL * X_VAL
BTEMP_VAL = HTEMP_VAL * FTEMP_VAL - DTEMP_VAL
DTEMP_VAL = HTEMP_VAL * CTEMP_VAL - I_VAL * DTEMP_VAL

Do While ((Abs(BTEMP_VAL * CTEMP_VAL - ATEMP_VAL * DTEMP_VAL) > _
           Abs(tolerance * CTEMP_VAL * BTEMP_VAL)))

    GTEMP_VAL = FTEMP_VAL * FTEMP_VAL * X_VAL
    FTEMP_VAL = FTEMP_VAL + D_VAL
    HTEMP_VAL = HTEMP_VAL + D_VAL
    ATEMP_VAL = HTEMP_VAL * BTEMP_VAL - GTEMP_VAL * ATEMP_VAL
    CTEMP_VAL = HTEMP_VAL * DTEMP_VAL - GTEMP_VAL * CTEMP_VAL
    
    GTEMP_VAL = ETEMP_VAL * ETEMP_VAL * X_VAL
    ETEMP_VAL = ETEMP_VAL + D_VAL
    HTEMP_VAL = HTEMP_VAL + D_VAL
    BTEMP_VAL = HTEMP_VAL * ATEMP_VAL - GTEMP_VAL * BTEMP_VAL
    DTEMP_VAL = HTEMP_VAL * CTEMP_VAL - GTEMP_VAL * DTEMP_VAL
           
    If (DTEMP_VAL > 1.15792089237316E+77) Then
        ATEMP_VAL = ATEMP_VAL * 8.63616855509444E-78
        CTEMP_VAL = CTEMP_VAL * 8.63616855509444E-78
        BTEMP_VAL = BTEMP_VAL * 8.63616855509444E-78
        DTEMP_VAL = DTEMP_VAL * 8.63616855509444E-78
    ElseIf (DTEMP_VAL < 8.63616855509444E-78) Then
        ATEMP_VAL = ATEMP_VAL * 1.15792089237316E+77
        CTEMP_VAL = CTEMP_VAL * 1.15792089237316E+77
        BTEMP_VAL = BTEMP_VAL * 1.15792089237316E+77
        DTEMP_VAL = DTEMP_VAL * 1.15792089237316E+77
    End If
     
Loop
     
LOG_SERIE_FUNC = BTEMP_VAL / DTEMP_VAL

Exit Function
ERROR_LABEL:
LOG_SERIE_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : GAMMA_FUNC
'DESCRIPTION   : Gamma Function
'LIBRARY       : NUMBER_REAL
'GROUP         : LOGARITHM
'ID            : 011
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/08/2008
'************************************************************************************
'************************************************************************************

Function GAMMA_FUNC(ByVal X_VAL As Double)

    Dim i As Long

    Dim ATEMP_VAL As Double
    Dim BTEMP_VAL As Double
    Dim CTEMP_VAL As Double

    Dim PTEMP_VAL As Double
    Dim QTEMP_VAL As Double
    Dim VTEMP_VAL As Double
    Dim WTEMP_VAL As Double
    Dim YTEMP_VAL As Double
    Dim ZTEMP_VAL As Double
    
    Dim SGN_VAL As Double
    Dim RESULT_VAL As Double
    Dim FACT_VAL As Double
    
    Dim PI_VAL As Double
    
    On Error GoTo ERROR_LABEL

'GAMMA_FUNC function
'Input parameters: X_VAL   -   argument
'Domain:
'    0 < X_VAL < 171.6
'    -170 < X_VAL < 0, X_VAL is not an integer.
'Relative error:
' arithmetic   domain     # trials      peak         rms
'    IEEE    -170,-33      20000       2.3e-15     3.3e-16
'    IEEE     -33,  33     20000       9.4e-16     2.2e-16
'    IEEE      33, 171.6   20000       2.3e-15     3.2e-16
'
'REFERENCES:
'Cephes Math Library Release 2.8:  June, 2000

    PI_VAL = 3.14159265358979
    SGN_VAL = 1#
    QTEMP_VAL = Abs(X_VAL)
    
'   Arguments |x| <= 34 are reduced by recurrence and the function
'   approximated by a rational function of degree 6/7 in the interval
'   (2,3). Large arguments are handled by Stirling's formula. Large
'   negative arguments are made positive using a reflection formula.
    
    If QTEMP_VAL > 33# Then
        If X_VAL < 0# Then
            PTEMP_VAL = Int(QTEMP_VAL)
            i = Round(PTEMP_VAL)
            If i Mod 2# = 0# Then
                SGN_VAL = -1#
            End If
            ZTEMP_VAL = QTEMP_VAL - PTEMP_VAL
            If ZTEMP_VAL > 0.5 Then
                PTEMP_VAL = PTEMP_VAL + 1#
                ZTEMP_VAL = QTEMP_VAL - PTEMP_VAL
            End If
            
            ZTEMP_VAL = QTEMP_VAL * Sin(PI_VAL * ZTEMP_VAL)
            ZTEMP_VAL = Abs(ZTEMP_VAL)

            WTEMP_VAL = 1# / QTEMP_VAL
            FACT_VAL = 7.87311395793094E-04
            FACT_VAL = -2.29549961613378E-04 + WTEMP_VAL * FACT_VAL
            FACT_VAL = -2.68132617805781E-03 + WTEMP_VAL * FACT_VAL
            FACT_VAL = 3.47222221605459E-03 + WTEMP_VAL * FACT_VAL
            FACT_VAL = 8.33333333333482E-02 + WTEMP_VAL * FACT_VAL
            WTEMP_VAL = 1# + WTEMP_VAL * FACT_VAL
            YTEMP_VAL = Exp(QTEMP_VAL)
            If QTEMP_VAL > 143.01608 Then
                VTEMP_VAL = QTEMP_VAL ^ (0.5 * QTEMP_VAL - 0.25)
                YTEMP_VAL = VTEMP_VAL * (VTEMP_VAL / YTEMP_VAL)
            Else: YTEMP_VAL = QTEMP_VAL ^ (QTEMP_VAL - 0.5) / YTEMP_VAL
            End If
            CTEMP_VAL = 2.506628274631 * YTEMP_VAL * WTEMP_VAL
            ZTEMP_VAL = PI_VAL / (ZTEMP_VAL * CTEMP_VAL)
        Else
            
            WTEMP_VAL = 1# / X_VAL
            FACT_VAL = 7.87311395793094E-04
            FACT_VAL = -2.29549961613378E-04 + WTEMP_VAL * FACT_VAL
            FACT_VAL = -2.68132617805781E-03 + WTEMP_VAL * FACT_VAL
            FACT_VAL = 3.47222221605459E-03 + WTEMP_VAL * FACT_VAL
            FACT_VAL = 8.33333333333482E-02 + WTEMP_VAL * FACT_VAL
            WTEMP_VAL = 1# + WTEMP_VAL * FACT_VAL
            YTEMP_VAL = Exp(X_VAL)
            If X_VAL > 143.01608 Then
                VTEMP_VAL = X_VAL ^ (0.5 * X_VAL - 0.25)
                YTEMP_VAL = VTEMP_VAL * (VTEMP_VAL / YTEMP_VAL)
            Else: YTEMP_VAL = X_VAL ^ (X_VAL - 0.5) / YTEMP_VAL
            End If
            CTEMP_VAL = 2.506628274631 * YTEMP_VAL * WTEMP_VAL
            ZTEMP_VAL = CTEMP_VAL
        End If
        RESULT_VAL = SGN_VAL * ZTEMP_VAL
        GAMMA_FUNC = RESULT_VAL
        Exit Function
    End If
    ZTEMP_VAL = 1#
    Do While X_VAL >= 3#
        X_VAL = X_VAL - 1#
        ZTEMP_VAL = ZTEMP_VAL * X_VAL
    Loop
    Do While X_VAL < 0#
        If X_VAL > -0.000000001 Then
            RESULT_VAL = ZTEMP_VAL / ((1# + 0.577215664901533 * _
                X_VAL) * X_VAL)
            GAMMA_FUNC = RESULT_VAL
            Exit Function
        End If
        ZTEMP_VAL = ZTEMP_VAL / X_VAL
        X_VAL = X_VAL + 1#
    Loop
    Do While X_VAL < 2#
        If X_VAL < 0.000000001 Then
            RESULT_VAL = ZTEMP_VAL / ((1# + 0.577215664901533 * _
                X_VAL) * X_VAL)
            GAMMA_FUNC = RESULT_VAL
            Exit Function
        End If
        ZTEMP_VAL = ZTEMP_VAL / X_VAL
        X_VAL = X_VAL + 1#
    Loop
    If X_VAL = 2# Then
        RESULT_VAL = ZTEMP_VAL
        GAMMA_FUNC = RESULT_VAL
        Exit Function
    End If
    X_VAL = X_VAL - 2#
    ATEMP_VAL = 1.60119522476752E-04
    ATEMP_VAL = 1.19135147006586E-03 + X_VAL * ATEMP_VAL
    ATEMP_VAL = 1.04213797561762E-02 + X_VAL * ATEMP_VAL
    ATEMP_VAL = 4.76367800457137E-02 + X_VAL * ATEMP_VAL
    ATEMP_VAL = 0.207448227648436 + X_VAL * ATEMP_VAL
    ATEMP_VAL = 0.494214826801497 + X_VAL * ATEMP_VAL
    ATEMP_VAL = 1# + X_VAL * ATEMP_VAL
    BTEMP_VAL = -2.3158187332412E-05
    BTEMP_VAL = 5.39605580493303E-04 + X_VAL * BTEMP_VAL
    BTEMP_VAL = -4.45641913851797E-03 + X_VAL * BTEMP_VAL
    BTEMP_VAL = 0.011813978522206 + X_VAL * BTEMP_VAL
    BTEMP_VAL = 3.58236398605499E-02 + X_VAL * BTEMP_VAL
    BTEMP_VAL = -0.234591795718243 + X_VAL * BTEMP_VAL
    BTEMP_VAL = 7.14304917030273E-02 + X_VAL * BTEMP_VAL
    BTEMP_VAL = 1# + X_VAL * BTEMP_VAL
    RESULT_VAL = ZTEMP_VAL * ATEMP_VAL / BTEMP_VAL
    

    GAMMA_FUNC = RESULT_VAL

Exit Function
ERROR_LABEL:
GAMMA_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : GAMMA_BETA_FUNC
'DESCRIPTION   : Beta function
'LIBRARY       : NUMBER_REAL
'GROUP         : LOGARITHM
'ID            : 012
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/08/2008
'************************************************************************************
'************************************************************************************

Function GAMMA_BETA_FUNC(ByVal Z_VAL As Double, _
ByVal W_VAL As Double)

On Error GoTo ERROR_LABEL

    GAMMA_BETA_FUNC = Exp(GAMMA_LN_FUNC(Z_VAL) + _
            GAMMA_LN_FUNC(W_VAL) - GAMMA_LN_FUNC(Z_VAL + _
            W_VAL))

Exit Function
ERROR_LABEL:
GAMMA_BETA_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : GAMMA_LN_FUNC
'DESCRIPTION   : Natural logarithm of gamma function
'LIBRARY       : NUMBER_REAL
'GROUP         : LOGARITHM
'ID            : 013
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/08/2008
'************************************************************************************
'************************************************************************************

Function GAMMA_LN_FUNC(ByVal X_VAL As Double)
    
    Dim i As Long
    
    Dim ATEMP_VAL As Double
    Dim BTEMP_VAL As Double
    Dim CTEMP_VAL As Double
    
    Dim PTEMP_VAL As Double
    Dim QTEMP_VAL As Double
    Dim UTEMP_VAL As Double
    Dim WTEMP_VAL As Double
    Dim ZTEMP_VAL As Double
    
    Dim PI_VAL As Double
    Dim LOG_PI_VAL As Double
    Dim LS2_PI_VAL As Double
    
    Dim SGN_VAL As Double
    Dim RESULT_VAL As Double

    On Error GoTo ERROR_LABEL
    
'Input parameters: X_VAL       -   argument
'RESULT_VAL: logarithm of the absolute value of the Gamma(X_VAL).
'Output parameters: SGN_VAL  -   sign(Gamma(X_VAL))
'
'Domain:
'    0 < X_VAL < 2.55e305
'    -2.55e305 < X_VAL < 0, X_VAL is not an integer.
'
'ACCURACY:
'arithmetic      domain        # trials     peak         rms
'   IEEE    0, 3                 28000     5.4e-16     1.1e-16
'   IEEE    2.718, 2.556e305     40000     3.5e-16     8.3e-17
'The error criterion was relative when the function magnitude
'was greater than one but absolute when it was less than one.
'
'The following test used the relative error criterion, though
'at certain points the relative error could be much higher than
'indicated.
'   IEEE    -200, -4             10000     4.8e-16     1.3e-16
'
'Cephes Math Library Release 2.8:  June, 2000
'Translated to AlgoPascal by Bochkanov Sergey (2005, 2006, 2007)
  
    SGN_VAL = 1#
    PI_VAL = 3.14159265358979
    LOG_PI_VAL = 1.1447298858494
    LS2_PI_VAL = 0.918938533204673
    
'   For arguments greater than 13, the logarithm of the gamma function
'   is approximated by the logarithmic version of Stirling's formula
'   using ATEMP_VAL polynomial approximation of degree 4. Arguments between -33
'   and +33 are reduced by recurrence to the interval [2,3] of ATEMP_VAL rational
'   approximation. The cosecant reflection formula is employed for arguments
'   less than -33.
    
    If X_VAL < -34# Then
        QTEMP_VAL = -X_VAL
        WTEMP_VAL = GAMMA_LN_FUNC(QTEMP_VAL)
        PTEMP_VAL = Int(QTEMP_VAL)
        i = Round(PTEMP_VAL)
        If i Mod 2# = 0# Then
              SGN_VAL = -1#
        Else: SGN_VAL = 1#
        End If
        ZTEMP_VAL = QTEMP_VAL - PTEMP_VAL
        If ZTEMP_VAL > 0.5 Then
            PTEMP_VAL = PTEMP_VAL + 1#
            ZTEMP_VAL = PTEMP_VAL - QTEMP_VAL
        End If
        ZTEMP_VAL = QTEMP_VAL * Sin(PI_VAL * ZTEMP_VAL)
        RESULT_VAL = LOG_PI_VAL - Log(ZTEMP_VAL) - WTEMP_VAL
        GAMMA_LN_FUNC = RESULT_VAL
        Exit Function
    End If
    If X_VAL < 13# Then
        ZTEMP_VAL = 1#
        PTEMP_VAL = 0#
        UTEMP_VAL = X_VAL
        Do While UTEMP_VAL >= 3#
            PTEMP_VAL = PTEMP_VAL - 1#
            UTEMP_VAL = X_VAL + PTEMP_VAL
            ZTEMP_VAL = ZTEMP_VAL * UTEMP_VAL
        Loop
        Do While UTEMP_VAL < 2#
            ZTEMP_VAL = ZTEMP_VAL / UTEMP_VAL
            PTEMP_VAL = PTEMP_VAL + 1#
            UTEMP_VAL = X_VAL + PTEMP_VAL
        Loop
        If ZTEMP_VAL < 0# Then
            SGN_VAL = -1#
            ZTEMP_VAL = -ZTEMP_VAL
        Else
            SGN_VAL = 1#
        End If
        If UTEMP_VAL = 2# Then
            RESULT_VAL = Log(ZTEMP_VAL)
            GAMMA_LN_FUNC = RESULT_VAL
            Exit Function
        End If
        PTEMP_VAL = PTEMP_VAL - 2#
        X_VAL = X_VAL + PTEMP_VAL
        BTEMP_VAL = -1378.25152569121
        BTEMP_VAL = -38801.6315134638 + X_VAL * BTEMP_VAL
        BTEMP_VAL = -331612.992738871 + X_VAL * BTEMP_VAL
        BTEMP_VAL = -1162370.97492762 + X_VAL * BTEMP_VAL
        BTEMP_VAL = -1721737.0082084 + X_VAL * BTEMP_VAL
        BTEMP_VAL = -853555.664245765 + X_VAL * BTEMP_VAL
        CTEMP_VAL = 1#
        CTEMP_VAL = -351.815701436523 + X_VAL * CTEMP_VAL
        CTEMP_VAL = -17064.2106651881 + X_VAL * CTEMP_VAL
        CTEMP_VAL = -220528.590553854 + X_VAL * CTEMP_VAL
        CTEMP_VAL = -1139334.44367983 + X_VAL * CTEMP_VAL
        CTEMP_VAL = -2532523.07177583 + X_VAL * CTEMP_VAL
        CTEMP_VAL = -2018891.41433533 + X_VAL * CTEMP_VAL
        PTEMP_VAL = X_VAL * BTEMP_VAL / CTEMP_VAL
        RESULT_VAL = Log(ZTEMP_VAL) + PTEMP_VAL
        
        GAMMA_LN_FUNC = RESULT_VAL
        Exit Function
    End If
    QTEMP_VAL = (X_VAL - 0.5) * Log(X_VAL) - X_VAL + LS2_PI_VAL
    If X_VAL > 100000000# Then
        RESULT_VAL = QTEMP_VAL
        
        GAMMA_LN_FUNC = RESULT_VAL
        Exit Function
    End If
    PTEMP_VAL = 1# / (X_VAL * X_VAL)
    If X_VAL >= 1000# Then
        QTEMP_VAL = QTEMP_VAL + ((7.93650793650794 * 0.0001 * PTEMP_VAL - _
            2.77777777777778 * 0.001) * PTEMP_VAL + _
            8.33333333333333E-02) / X_VAL
    Else
        ATEMP_VAL = 8.11614167470508 * 0.0001
        ATEMP_VAL = -(5.95061904284301 * 0.0001) + PTEMP_VAL * ATEMP_VAL
        ATEMP_VAL = 7.93650340457717 * 0.0001 + PTEMP_VAL * ATEMP_VAL
        ATEMP_VAL = -(2.777777777301 * 0.001) + PTEMP_VAL * ATEMP_VAL
        ATEMP_VAL = 8.33333333333332 * 0.01 + PTEMP_VAL * ATEMP_VAL
        QTEMP_VAL = QTEMP_VAL + ATEMP_VAL / X_VAL
    End If
    RESULT_VAL = QTEMP_VAL

    GAMMA_LN_FUNC = RESULT_VAL
    
Exit Function
ERROR_LABEL:
GAMMA_LN_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : GAMMA_SPLIT_FUNC
'DESCRIPTION   : Approximation algorithm for gamma function
'LIBRARY       : NUMBER_REAL
'GROUP         : LOGARITHM
'ID            : 014
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/08/2008
'************************************************************************************
'************************************************************************************

Function GAMMA_SPLIT_FUNC(ByVal X_VAL As Double)

'    TEMP_ARR = GAMMA_SPLIT_FUNC(X_VAL)
'    GAMMA_LN_FUNC = Log(TEMP_ARR(1)) + _
    (TEMP_ARR(2) * Log(10)) logarithm gamma function

Dim i As Long

Dim ATEMP_VAL As Double
Dim BTEMP_VAL As Double
Dim CTEMP_VAL As Double
Dim DTEMP_VAL As Double
Dim ETEMP_VAL As Double
Dim FTEMP_VAL As Double

Dim AFACTOR_VAL As Double
Dim BFACTOR_VAL As Double

Dim TEMP_ARR As Variant

On Error GoTo ERROR_LABEL

AFACTOR_VAL = 6.28318530717959
BFACTOR_VAL = 4.7421875  '607/128

CTEMP_VAL = X_VAL - 1
    
ReDim TEMP_ARR(0 To 14)
    
TEMP_ARR(0) = 0.999999999999997
TEMP_ARR(1) = 57.1562356658629
TEMP_ARR(2) = -59.5979603554755
TEMP_ARR(3) = 14.1360979747417
TEMP_ARR(4) = -0.49191381609762
TEMP_ARR(5) = 3.39946499848119E-05
TEMP_ARR(6) = 4.65236289270486E-05
TEMP_ARR(7) = -9.83744753048796E-05
TEMP_ARR(8) = 1.58088703224912E-04
TEMP_ARR(9) = -2.10264441724105E-04
TEMP_ARR(10) = 2.17439618115213E-04
TEMP_ARR(11) = -1.64318106536764E-04
TEMP_ARR(12) = 8.44182239838528E-05
TEMP_ARR(13) = -2.61908384015814E-05
TEMP_ARR(14) = 3.68991826595316E-06
    
DTEMP_VAL = Exp(BFACTOR_VAL) / (AFACTOR_VAL) ^ 0.5
ETEMP_VAL = TEMP_ARR(0)
For i = 1 To 14
    ETEMP_VAL = ETEMP_VAL + TEMP_ARR(i) / (CTEMP_VAL + i)
Next i
    
ETEMP_VAL = ETEMP_VAL / DTEMP_VAL
FTEMP_VAL = Log((CTEMP_VAL + BFACTOR_VAL + 0.5) / _
Exp(1)) * (CTEMP_VAL + 0.5) / Log(10) 'split in ATEMP_VAL and
    'exponent to avoid overflow
    
BTEMP_VAL = Int(FTEMP_VAL)
FTEMP_VAL = FTEMP_VAL - Int(FTEMP_VAL)
    
ATEMP_VAL = 10 ^ FTEMP_VAL * ETEMP_VAL 'rescaling
FTEMP_VAL = Int(Log(ATEMP_VAL) / Log(10))
    
ATEMP_VAL = ATEMP_VAL * 10 ^ -FTEMP_VAL
BTEMP_VAL = BTEMP_VAL + FTEMP_VAL
    
ReDim TEMP_ARR(1 To 2)
TEMP_ARR(1) = ATEMP_VAL 'Mantissa
TEMP_ARR(2) = BTEMP_VAL 'Expo
    
GAMMA_SPLIT_FUNC = TEMP_ARR
    
Exit Function
ERROR_LABEL:
GAMMA_SPLIT_FUNC = Err.number
End Function
