Attribute VB_Name = "FINAN_PORT_WEIGHTS_TURN_LIBR"


'-----------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------
Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.
'-----------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------


'************************************************************************************
'************************************************************************************
'FUNCTION      : PORT_TURNOVER_FUNC

'DESCRIPTION   : Indicator of trading activity in a portfolio
' The rate of trading activity in a fund's portfolio of investments,
' equal to the lesser of purchases or sales divided by average total
' exposure.

' Increase of Total Exposure from X to Y

'LIBRARY       : PORTFOLIO
'GROUP         : TURNOVER
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 21/01/2009
'************************************************************************************
'************************************************************************************

Function PORT_TURNOVER_FUNC(ByRef NEW_ALLOC_RNG As Variant, _
ByRef OLD_ALLOC_RNG As Variant, _
Optional ByVal OUTPUT As Integer = 0)

Dim i As Long
Dim NROWS As Long

Dim TEMP1_SUM As Double
Dim TEMP2_SUM As Double
Dim TEMP3_SUM As Double
Dim TEMP4_SUM As Double

Dim OLD_VECTOR As Variant
Dim NEW_VECTOR As Variant

Dim TEMP_MATRIX As Variant

On Error GoTo ERROR_LABEL

NEW_VECTOR = NEW_ALLOC_RNG
If UBound(NEW_VECTOR, 1) = 1 Then
    NEW_VECTOR = MATRIX_TRANSPOSE_FUNC(NEW_VECTOR)
End If

OLD_VECTOR = OLD_ALLOC_RNG
If UBound(OLD_VECTOR, 1) = 1 Then
    OLD_VECTOR = MATRIX_TRANSPOSE_FUNC(OLD_VECTOR)
End If

If UBound(NEW_VECTOR, 1) <> UBound(OLD_VECTOR, 1) Then: GoTo ERROR_LABEL

NROWS = UBound(NEW_VECTOR, 1)


'----------------------------------------------------------------------------------
If OUTPUT = 0 Then
'----------------------------------------------------------------------------------
    TEMP1_SUM = 0
    TEMP2_SUM = 0
    TEMP3_SUM = 0
    TEMP4_SUM = 0

    For i = 1 To NROWS
        If NEW_VECTOR(i, 1) < OLD_VECTOR(i, 1) Then
            TEMP2_SUM = TEMP2_SUM + OLD_VECTOR(i, 1) - NEW_VECTOR(i, 1)
        Else
            TEMP1_SUM = TEMP1_SUM + NEW_VECTOR(i, 1) - OLD_VECTOR(i, 1)
        End If
        TEMP3_SUM = TEMP3_SUM + OLD_VECTOR(i, 1)
        TEMP4_SUM = TEMP4_SUM + NEW_VECTOR(i, 1)
    Next i

    If TEMP2_SUM < TEMP1_SUM Then: TEMP1_SUM = TEMP2_SUM
    
    PORT_TURNOVER_FUNC = TEMP1_SUM / ((TEMP3_SUM + TEMP4_SUM) * 0.5)
    'TurnOver Function
'----------------------------------------------------------------------------------
Else
'----------------------------------------------------------------------------------
    ReDim TEMP_MATRIX(0 To NROWS + 2, 1 To 4)

    TEMP_MATRIX(0, 1) = ("OLD ALLOCATION")
    TEMP_MATRIX(0, 2) = ("NEW ALLOCATION")
    TEMP_MATRIX(0, 3) = ("PURCHASES")
    TEMP_MATRIX(0, 4) = ("SALES")
    
    For i = 1 To NROWS
        TEMP_MATRIX(i, 1) = OLD_VECTOR(i, 1)
        TEMP1_SUM = TEMP1_SUM + TEMP_MATRIX(i, 1)
        
        TEMP_MATRIX(i, 2) = NEW_VECTOR(i, 1)
        TEMP2_SUM = TEMP2_SUM + TEMP_MATRIX(i, 2)
        
        TEMP_MATRIX(i, 3) = IIf((NEW_VECTOR(i, 1) - OLD_VECTOR(i, 1)) > 0, (NEW_VECTOR(i, 1) - OLD_VECTOR(i, 1)), 0)
        TEMP3_SUM = TEMP3_SUM + TEMP_MATRIX(i, 3)
        
        TEMP_MATRIX(i, 4) = IIf((NEW_VECTOR(i, 1) - OLD_VECTOR(i, 1)) < 0, (NEW_VECTOR(i, 1) - OLD_VECTOR(i, 1)), 0) * -1
        TEMP4_SUM = TEMP4_SUM + TEMP_MATRIX(i, 4)
    
    Next i
    
    'Total Exposures
    TEMP_MATRIX(NROWS + 1, 1) = ""
    TEMP_MATRIX(NROWS + 1, 2) = ""
    TEMP_MATRIX(NROWS + 1, 3) = ""
    TEMP_MATRIX(NROWS + 1, 4) = ""

    TEMP_MATRIX(NROWS + 2, 1) = TEMP1_SUM
    TEMP_MATRIX(NROWS + 2, 2) = TEMP2_SUM
    TEMP_MATRIX(NROWS + 2, 3) = TEMP3_SUM
    TEMP_MATRIX(NROWS + 2, 4) = TEMP4_SUM

    PORT_TURNOVER_FUNC = TEMP_MATRIX

'----------------------------------------------------------------------------------
End If
'----------------------------------------------------------------------------------
Exit Function
ERROR_LABEL:
PORT_TURNOVER_FUNC = Err.number
End Function
