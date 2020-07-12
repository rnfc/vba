Attribute VB_Name = "STAT_HISTOGRAM_BIN_LIBR"


Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.

'************************************************************************************
'************************************************************************************
'FUNCTION      : HISTOGRAM_BIN_PRECISION_FUNC
'DESCRIPTION   : Scale the the reference interval
'LIBRARY       : STATISTICS
'GROUP         : HISTOGRAM_BIN
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/22/2009
'************************************************************************************
'************************************************************************************

Function HISTOGRAM_BIN_PRECISION_FUNC(ByVal BIN_VALUE As Double)

Dim TEMP_VAL As Double

On Error GoTo ERROR_LABEL

If BIN_VALUE < 0 Then: GoTo ERROR_LABEL
If IsNumeric(BIN_VALUE) = False Then: GoTo ERROR_LABEL

TEMP_VAL = 1
Do While TEMP_VAL > BIN_VALUE
    TEMP_VAL = TEMP_VAL / 10
Loop

HISTOGRAM_BIN_PRECISION_FUNC = -(Log(TEMP_VAL) / Log(10#)) + 1

Exit Function
ERROR_LABEL:
HISTOGRAM_BIN_PRECISION_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : HISTOGRAM_BIN_MIN_PRECISION_FUNC
'DESCRIPTION   : Bin Min Precision Tune Up
'LIBRARY       : STATISTICS
'GROUP         : HISTOGRAM_BIN
'ID            : 002
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/22/2009
'************************************************************************************
'************************************************************************************

Function HISTOGRAM_BIN_MIN_PRECISION_FUNC(ByVal BIN_MIN As Double, _
ByVal BIN_WIDTH As Double)

'REMEMBER THAT BIN_WIDTH = HISTOGRAM_BIN_WIDTH_PRECISION_FUNC(BIN_WIDTH)


Dim A_VAL As Double
Dim B_VAL As Double

Dim NEG_VAL As Double
Dim TEMP_VAL As Double

On Error GoTo ERROR_LABEL

TEMP_VAL = BIN_MIN

If TEMP_VAL < 0 Then
    TEMP_VAL = -TEMP_VAL
    NEG_VAL = -1
Else
    NEG_VAL = 1
End If

A_VAL = BIN_WIDTH
B_VAL = A_VAL * Int(TEMP_VAL / A_VAL)
B_VAL = B_VAL * NEG_VAL

If NEG_VAL = -1 Then
    HISTOGRAM_BIN_MIN_PRECISION_FUNC = B_VAL - A_VAL
Else
    HISTOGRAM_BIN_MIN_PRECISION_FUNC = B_VAL
End If

Exit Function
ERROR_LABEL:
HISTOGRAM_BIN_MIN_PRECISION_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : HISTOGRAM_BIN_WIDTH_PRECISION_FUNC
'DESCRIPTION   : Bin Width Precision Function
'LIBRARY       : STATISTICS
'GROUP         : HISTOGRAM_BIN
'ID            : 003
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/22/2009
'************************************************************************************
'************************************************************************************

Function HISTOGRAM_BIN_WIDTH_PRECISION_FUNC(ByVal BIN_WIDTH As Double)

Dim A_VAL As Double
Dim B_VAL As Double
Dim C_VAL As Double
Dim D_VAL As Double
Dim E_VAL As Double
Dim F_VAL As Double

Dim TEMP_VAL As Double

On Error GoTo ERROR_LABEL

TEMP_VAL = BIN_WIDTH

If TEMP_VAL <= 0 Then
        HISTOGRAM_BIN_WIDTH_PRECISION_FUNC = 0
    Exit Function
End If

F_VAL = Int(Log(TEMP_VAL) / Log(10#))
E_VAL = (Log(TEMP_VAL) / Log(10#)) - F_VAL

A_VAL = (Log(1.75) / Log(10#))
B_VAL = (Log(2.25) / Log(10#))
C_VAL = (Log(4.5) / Log(10#))
D_VAL = (Log(8.75) / Log(10#))

TEMP_VAL = 10 ^ F_VAL

If A_VAL < E_VAL Then
    If B_VAL < E_VAL Then
        If C_VAL < E_VAL Then
            If D_VAL < E_VAL Then
                TEMP_VAL = TEMP_VAL * 10
            Else
                TEMP_VAL = TEMP_VAL * 5
            End If
        Else
            TEMP_VAL = TEMP_VAL * 2.5
        End If
    Else
        TEMP_VAL = TEMP_VAL * 2
    End If
End If

HISTOGRAM_BIN_WIDTH_PRECISION_FUNC = TEMP_VAL

Exit Function
ERROR_LABEL:
HISTOGRAM_BIN_WIDTH_PRECISION_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : HISTOGRAM_BIN_LIMITS_FUNC
'DESCRIPTION   : Histogram Limits Maker
'LIBRARY       : STATISTICS
'GROUP         : HISTOGRAM_BIN
'ID            : 004
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/22/2009
'************************************************************************************
'************************************************************************************

Function HISTOGRAM_BIN_LIMITS_FUNC(ByVal MIN_VAL As Double, _
ByVal MAX_VAL As Double, _
ByVal NROWS As Double, _
Optional ByVal OUTPUT As Integer = 0)

Dim NBINS As Long

Dim OLD_TEMP As Double
Dim MULT_TEMP As Double
Dim ADJ_TEMP As Double
Dim FACT_TEMP As Double

Dim PREC_TEMP As Double
Dim BIN_MIN As Double
Dim BIN_WIDTH As Double

Dim ABS_TEMP As Double
Dim REMAIN_TEMP As Double

On Error GoTo ERROR_LABEL

If MIN_VAL >= MAX_VAL Then
    HISTOGRAM_BIN_LIMITS_FUNC = 0
    Exit Function
End If

If NROWS < 1000 Then ' BIN_WIDTH:  this is binWIDTH_TEMP
     BIN_WIDTH = (MAX_VAL - MIN_VAL) / Sqr(NROWS)
Else
     BIN_WIDTH = (MAX_VAL - MIN_VAL) / (10 * (Log(NROWS) / Log(10#)))
End If

ABS_TEMP = Abs(MIN_VAL)  ' absolute value of minimum
BIN_MIN = 100
PREC_TEMP = 1

If ABS_TEMP > 0 Then
    If BIN_MIN < ABS_TEMP Then
    ' Bring BIN_MIN back above ABS_TEMP
        Do While ABS_TEMP > BIN_MIN
            BIN_MIN = BIN_MIN * 10
        Loop
        PREC_TEMP = BIN_MIN / 10
        'BIN_MIN = 9 * PREC_TEMP
        FACT_TEMP = 10
        Do While ABS_TEMP < BIN_MIN
            FACT_TEMP = FACT_TEMP - 1
            BIN_MIN = FACT_TEMP * PREC_TEMP
        Loop
    ' We're done:  BIN_MIN is just below PREC_TEMP
    ElseIf ABS_TEMP < BIN_MIN Then
        Do While ABS_TEMP < BIN_MIN
            BIN_MIN = BIN_MIN / 10
        Loop
        PREC_TEMP = BIN_MIN
        FACT_TEMP = 1
        If ABS_TEMP <> BIN_MIN Then
            Do While ABS_TEMP > BIN_MIN
                FACT_TEMP = FACT_TEMP + 1
                BIN_MIN = FACT_TEMP * PREC_TEMP
            Loop
            BIN_MIN = (FACT_TEMP - 1) * PREC_TEMP
      ' We're done:  BIN_MIN is just below PREC_TEMP
        End If
    
    End If
Else
    BIN_MIN = 0
    PREC_TEMP = 1
End If
REMAIN_TEMP = ABS_TEMP - BIN_MIN  'This is the difference
' Next step it to adjust BIN_WIDTH so that it is even
ADJ_TEMP = 1  ' the default;
Do While REMAIN_TEMP > BIN_WIDTH
    ADJ_TEMP = 10  'we're going to have to lower BIN_MIN by 10*PREC_TEMP
    'if MIN_VAL < 0
    Do While BIN_MIN < ABS_TEMP
        BIN_MIN = BIN_MIN + PREC_TEMP
    Loop
    BIN_MIN = BIN_MIN - PREC_TEMP
    REMAIN_TEMP = ABS_TEMP - BIN_MIN
    PREC_TEMP = PREC_TEMP / 10
Loop
If MIN_VAL < 0 Then BIN_MIN = -BIN_MIN - ADJ_TEMP * PREC_TEMP
' Next step is to get the BIN_WIDTH to be reasonable
' MIN_VAL with original BIN_WIDTH
MULT_TEMP = 0
If PREC_TEMP > BIN_WIDTH Then
    Do While BIN_WIDTH < PREC_TEMP
        OLD_TEMP = PREC_TEMP
        If MULT_TEMP = 0 Then
            PREC_TEMP = PREC_TEMP / 2
            MULT_TEMP = 1
        ElseIf MULT_TEMP = 1 Then
            PREC_TEMP = PREC_TEMP / 2.5
            MULT_TEMP = 2
        Else
            PREC_TEMP = PREC_TEMP / 2
            MULT_TEMP = 0
        End If
    Loop
    BIN_WIDTH = PREC_TEMP
Else
    Do While PREC_TEMP < BIN_WIDTH
        OLD_TEMP = PREC_TEMP
        If MULT_TEMP = 0 Then
    
            PREC_TEMP = PREC_TEMP * 2
            MULT_TEMP = 1
            
        ElseIf MULT_TEMP = 1 Then
            PREC_TEMP = PREC_TEMP * 2.5
            MULT_TEMP = 2
        Else
            PREC_TEMP = PREC_TEMP * 2
            MULT_TEMP = 0
        End If
    Loop
    BIN_WIDTH = OLD_TEMP
End If

If BIN_WIDTH > Abs(BIN_MIN) Then
    If BIN_MIN >= 0 Then
        BIN_MIN = 0
    Else
        BIN_MIN = -BIN_WIDTH
    End If
End If

NBINS = Int((MAX_VAL - BIN_MIN) / BIN_WIDTH) + 1

Select Case OUTPUT
    Case 0 'BIN_WIDTH
        HISTOGRAM_BIN_LIMITS_FUNC = BIN_WIDTH
    Case 1 'BIN_MIN
        HISTOGRAM_BIN_LIMITS_FUNC = BIN_MIN
    Case 2 'Number of Bins
        HISTOGRAM_BIN_LIMITS_FUNC = NBINS
    Case Else
        HISTOGRAM_BIN_LIMITS_FUNC = Array(BIN_WIDTH, BIN_MIN, NBINS)
End Select

Exit Function
ERROR_LABEL:
HISTOGRAM_BIN_LIMITS_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : HISTOGRAM_BIN_SLIMITS_FUNC
'DESCRIPTION   : Histogram Scaled Limits
'LIBRARY       : STATISTICS
'GROUP         : HISTOGRAM_BIN
'ID            : 005
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/22/2009
'************************************************************************************
'************************************************************************************

Function HISTOGRAM_BIN_SLIMITS_FUNC(ByVal APPROX_SE As Double, _
ByVal BIN_CENTER As Double, _
Optional ByVal WIDTH_LIMIT As Double = 0.00001, _
Optional ByVal SD_MULT As Double = 4, _
Optional ByVal OUTPUT As Integer = 0)

' LOGIC:
' Adjust the scaling of the histogram
'   (1) Find the SE of the sample averages for
'       the average with greater spread
'            TEMP_VAL = 2* this SE
'   (2) BIN_WIDTH keeps on getting bigger in
'       increments like this: 0.1, 0.2, 0.5,
'                               1, 2, 5,....
'       until it is less than TEMP_VAL.
'   (3) lower end of scale of histogram scale is set at
'       true BIN_CENTER less 4 times resulting WIDTH_LIMIT
'       upper end is true BIN_CENTER plus 4 times WIDTH_LIMIT

Dim i As Long

Dim BIN_MAX As Double
Dim BIN_MIN As Double
Dim TEMP_MULT As Double
Dim TEMP_VAL As Double

On Error GoTo ERROR_LABEL

i = 0

TEMP_MULT = 0
TEMP_VAL = SD_MULT * APPROX_SE

Do While WIDTH_LIMIT < TEMP_VAL
    If TEMP_MULT = 0 Then
        WIDTH_LIMIT = WIDTH_LIMIT * 2
        TEMP_MULT = 1
        
    ElseIf TEMP_MULT = 1 Then
        WIDTH_LIMIT = WIDTH_LIMIT * 2.5
        TEMP_MULT = 2
    Else
        WIDTH_LIMIT = WIDTH_LIMIT * 2
        TEMP_MULT = 0
    End If
    i = i + 1
    If i > 100 Then: GoTo ERROR_LABEL
Loop
    
BIN_MIN = BIN_CENTER - 2 * WIDTH_LIMIT
BIN_MAX = BIN_CENTER + 2 * WIDTH_LIMIT

Select Case OUTPUT
Case 0
    HISTOGRAM_BIN_SLIMITS_FUNC = BIN_MIN
Case 1
    HISTOGRAM_BIN_SLIMITS_FUNC = BIN_MAX
Case 2
    HISTOGRAM_BIN_SLIMITS_FUNC = WIDTH_LIMIT ' BIN_WIDTH
Case Else
    HISTOGRAM_BIN_SLIMITS_FUNC = Array(BIN_MIN, BIN_MAX, WIDTH_LIMIT)
End Select

Exit Function
ERROR_LABEL:
HISTOGRAM_BIN_SLIMITS_FUNC = Err.number
End Function
