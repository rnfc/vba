Attribute VB_Name = "STAT_MOMENTS_STOCHASTIC_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.


'************************************************************************************
'************************************************************************************
'FUNCTION      : MATRIX_STOCHASTIC_DOMINANCE_FUNC

'DESCRIPTION   : Comparison of multiple investment alternatives in
'terms of FSD, SSD and TSD.

'---------------------------------------------------------------------------
'Drawback of the concept od Stochastic Dominance: "It is well known that
'one of the disadvantages of SD analysis in comparison to MV analysis
'is that in the SD framework we do not have yet an algorithm to find
'the SD efficient diversification strategies."

'Literature: http://www.casact.org/pubs/forum/01sforum/01sf095.pdf
'---------------------------------------------------------------------------

'LIBRARY       : STATISTICS
'GROUP         : STOCHASTIC_DOMINANCE
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/22/2009
'************************************************************************************
'************************************************************************************

Function MATRIX_STOCHASTIC_DOMINANCE_FUNC(ByRef DATA_RNG As Variant, _
Optional ByVal DATA_TYPE As Integer = 0, _
Optional ByVal LOG_SCALE As Integer = 0, _
Optional ByVal NBINS As Long = 50, _
Optional ByVal OUTPUT As Integer = 0)


Dim h As Long
Dim i As Long '
Dim j As Long '
Dim k As Long '
Dim l As Long '

Dim NROWS As Long
Dim NCOLUMNS As Long

Dim BIN_MIN As Double
Dim BIN_MAX As Double
Dim BIN_DELTA As Double

Dim ATEMP_FLAG As Boolean
Dim BTEMP_FLAG As Boolean
Dim CTEMP_FLAG As Boolean

Dim BIN_VECTOR As Variant
Dim FREQ_VECTOR As Variant

Dim BTEMP_MATRIX As Variant
Dim ATEMP_MATRIX As Variant
Dim DATA_MATRIX As Variant

On Error GoTo ERROR_LABEL

DATA_MATRIX = DATA_RNG
If UBound(DATA_MATRIX, 1) = 1 Then
    DATA_MATRIX = MATRIX_TRANSPOSE_FUNC(DATA_MATRIX)
End If
If DATA_TYPE <> 0 Then: DATA_MATRIX = MATRIX_PERCENT_FUNC(DATA_MATRIX, LOG_SCALE)

l = 8

NROWS = UBound(DATA_MATRIX, 1)
NCOLUMNS = UBound(DATA_MATRIX, 2)

BIN_MIN = MATRIX_ELEMENTS_MIN_FUNC(DATA_MATRIX, 0)
BIN_MAX = MATRIX_ELEMENTS_MAX_FUNC(DATA_MATRIX, 0)
BIN_DELTA = (Abs(BIN_MIN) + Abs(BIN_MAX)) / NBINS

ReDim ATEMP_MATRIX(1 To 3, 1 To NCOLUMNS + 1)
ReDim BTEMP_MATRIX(0 To NBINS, 1 To (l * NCOLUMNS))

ATEMP_MATRIX(1, 1) = "FSD" 'First-Order Stochastic Dominance
ATEMP_MATRIX(2, 1) = "SSD" 'Second-Order Stochastic Dominance
ATEMP_MATRIX(3, 1) = "TSD" 'Third-Order Stoastic Dominance

For i = 1 To 3: For j = 1 To NCOLUMNS: ATEMP_MATRIX(i, j + 1) = True: Next j: Next i
ReDim BIN_VECTOR(1 To NBINS, 1 To 1)

BIN_VECTOR(1, 1) = BIN_MIN
For i = 2 To NBINS: BIN_VECTOR(i, 1) = BIN_VECTOR(i - 1, 1) + BIN_DELTA: Next i
j = 0: k = 1
Do Until k > (NCOLUMNS)
    
    FREQ_VECTOR = HISTOGRAM_FREQUENCY_FUNC(MATRIX_GET_COLUMN_FUNC(DATA_MATRIX, k, 1), NBINS, _
    BIN_MIN, BIN_DELTA, 1)
    
    BTEMP_MATRIX(0, 1 + j) = "LOWER_LIMIT_" & k
    BTEMP_MATRIX(0, 2 + j) = "DISTR_" & k
    BTEMP_MATRIX(0, 3 + j) = "CUM_DISTR_" & k
    BTEMP_MATRIX(0, 4 + j) = "CUM_CUM_DISTR_" & k
    BTEMP_MATRIX(0, 5 + j) = "CUM_CUM_CUM_DISTR" & k
    BTEMP_MATRIX(0, 6 + j) = "FSD_" & k
    BTEMP_MATRIX(0, 7 + j) = "SSD_" & k
    BTEMP_MATRIX(0, 8 + j) = "TSD_" & k
    
    For i = 1 To NBINS
        BTEMP_MATRIX(i, 1 + j) = BIN_VECTOR(i, 1)
        BTEMP_MATRIX(i, 2 + j) = FREQ_VECTOR(i, 2) / NROWS
        If i = 1 Then
            BTEMP_MATRIX(i, 3 + j) = BTEMP_MATRIX(i, 2 + j)
            BTEMP_MATRIX(i, 4 + j) = BTEMP_MATRIX(i, 3 + j)
            BTEMP_MATRIX(i, 5 + j) = BTEMP_MATRIX(i, 4 + j)
        Else
            BTEMP_MATRIX(i, 3 + j) = BTEMP_MATRIX(i - 1, 3 + j) + BTEMP_MATRIX(i, 2 + j)
            BTEMP_MATRIX(i, 4 + j) = BTEMP_MATRIX(i - 1, 4 + j) + BTEMP_MATRIX(i, 3 + j)
            BTEMP_MATRIX(i, 5 + j) = BTEMP_MATRIX(i - 1, 5 + j) + BTEMP_MATRIX(i, 4 + j)
        End If
    Next i
    j = j + l
    k = k + 1
Loop

j = 0
k = 1
Do Until k > (NCOLUMNS)
    For i = 1 To NBINS
        
        h = (NCOLUMNS - 1) * l
        Do
            If (j = 0) And (h = 0) Then: GoTo 1983
        
            If BTEMP_MATRIX(i, 3 + j) >= BTEMP_MATRIX(i, 3 + h) Then
                ATEMP_MATRIX(1, k + 1) = False
                ATEMP_FLAG = False
            Else
                ATEMP_FLAG = True
            End If
        
            If BTEMP_MATRIX(i, 4 + j) >= BTEMP_MATRIX(i, 4 + h) Then
                ATEMP_MATRIX(2, k + 1) = False
                BTEMP_FLAG = False
            Else
                BTEMP_FLAG = True
            End If
        
            If BTEMP_MATRIX(i, 5 + j) >= BTEMP_MATRIX(i, 5 + h) Then
                ATEMP_MATRIX(3, k + 1) = False
                CTEMP_FLAG = False
            Else
                CTEMP_FLAG = True
            End If
1983:       h = h - l
        Loop Until h < 0
        
        BTEMP_MATRIX(i, 6 + j) = ATEMP_FLAG
        BTEMP_MATRIX(i, 7 + j) = BTEMP_FLAG
        BTEMP_MATRIX(i, 8 + j) = CTEMP_FLAG
    Next i
    j = j + l
    k = k + 1
Loop

Select Case OUTPUT
Case 0
    MATRIX_STOCHASTIC_DOMINANCE_FUNC = ATEMP_MATRIX
Case 1
    MATRIX_STOCHASTIC_DOMINANCE_FUNC = BTEMP_MATRIX
Case Else
    MATRIX_STOCHASTIC_DOMINANCE_FUNC = Array(ATEMP_MATRIX, BTEMP_MATRIX)
End Select

Exit Function
ERROR_LABEL:
MATRIX_STOCHASTIC_DOMINANCE_FUNC = Err.number
End Function
