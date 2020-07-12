Attribute VB_Name = "FINAN_ASSET_MOMENTS_STOCH_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.

'************************************************************************************
'************************************************************************************
'FUNCTION      : ASSET_STOCHASTIC_DOMINANCE_FUNC

'DESCRIPTION   : Comparison of two assets in terms of FSD, SSD and TSD.
'---------------------------------------------------------------------------
'Drawback of the concept od Stochastic Dominance: "It is well known that
'one of the disadvantages of SD analysis in comparison to MV analysis
'is that in the SD framework we do not have yet an algorithm to find
'the SD efficient diversification strategies."

'Literature: http://www.casact.org/pubs/forum/01sforum/01sf095.pdf
'---------------------------------------------------------------------------

'LIBRARY       : FINAN_ASSET
'GROUP         : MOMENTS
'ID            : 010
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA

'************************************************************************************
'************************************************************************************

Function ASSET_STOCHASTIC_DOMINANCE_FUNC(ByRef DATE_RNG As Variant, _
ByRef DATA1_RNG As Variant, _
ByRef DATA2_RNG As Variant, _
Optional ByVal DATA_TYPE As Variant = 0, _
Optional ByVal LOG_SCALE As Variant = 0, _
Optional ByVal OUTPUT As Variant = 0, _
Optional ByVal VERSION As Variant = 1, _
Optional ByVal PERCENTILE_VAL As Variant = 90, _
Optional ByVal FACTOR As Variant = 0.5)

'VERSION = 0 --> Percentile Approach
'VERSION = 1 --> Sorting Approach

Dim h As Long
Dim i As Long
Dim j As Long
Dim k As Long
Dim l As Long
Dim m As Long

Dim NSIZE As Long

Dim FLAG1_VAL As Boolean
Dim FLAG2_VAL As Boolean
Dim FLAG3_VAL As Boolean

Dim TEMP1_VECTOR As Variant
Dim TEMP2_VECTOR As Variant

Dim TEMP_MATRIX As Variant
Dim SUMMARY_MATRIX As Variant

Dim DATE_VECTOR As Variant
Dim DATA1_VECTOR As Variant
Dim DATA2_VECTOR As Variant

On Error GoTo ERROR_LABEL

DATE_VECTOR = DATE_RNG
If UBound(DATE_VECTOR, 1) = 1 Then
    DATE_VECTOR = MATRIX_TRANSPOSE_FUNC(DATE_VECTOR)
End If

DATA1_VECTOR = DATA1_RNG
If UBound(DATA1_VECTOR, 1) = 1 Then
    DATA1_VECTOR = MATRIX_TRANSPOSE_FUNC(DATA1_VECTOR)
End If

If UBound(DATE_VECTOR, 1) <> UBound(DATA1_VECTOR, 1) Then: GoTo ERROR_LABEL

DATA2_VECTOR = DATA2_RNG
If UBound(DATA2_VECTOR, 1) = 1 Then
    DATA2_VECTOR = MATRIX_TRANSPOSE_FUNC(DATA2_VECTOR)
End If

If UBound(DATE_VECTOR, 1) <> UBound(DATA2_VECTOR, 1) Then: GoTo ERROR_LABEL
h = 0
If DATA_TYPE <> 0 Then
    h = 1
    DATA1_VECTOR = MATRIX_PERCENT_FUNC(DATA1_VECTOR, LOG_SCALE)
    DATA2_VECTOR = MATRIX_PERCENT_FUNC(DATA2_VECTOR, LOG_SCALE)
End If

NSIZE = UBound(DATA1_VECTOR, 1)
ReDim SUMMARY_MATRIX(0 To 3, 1 To 3)
SUMMARY_MATRIX(0, 1) = "-"
SUMMARY_MATRIX(0, 2) = "AB SD"
SUMMARY_MATRIX(0, 3) = "DEGREE"

SUMMARY_MATRIX(1, 1) = "FSD"
SUMMARY_MATRIX(2, 1) = "SSD"
SUMMARY_MATRIX(3, 1) = "TSD"

FLAG1_VAL = True
FLAG2_VAL = True
FLAG3_VAL = True

j = 0
k = 0
l = 0

TEMP1_VECTOR = MATRIX_QUICK_SORT_FUNC(DATA1_VECTOR, 1, 1)
TEMP2_VECTOR = MATRIX_QUICK_SORT_FUNC(DATA2_VECTOR, 1, 1)

'--------------------------------------------------------------------------------
If VERSION = 0 Then 'Percentile Approach
'--------------------------------------------------------------------------------
    ReDim TEMP_MATRIX(0 To NSIZE, 1 To 19)
    
    TEMP_MATRIX(0, 1) = "DATE"
    TEMP_MATRIX(0, 2) = "ASSET A"
    TEMP_MATRIX(0, 3) = "ASSET B"
    TEMP_MATRIX(0, 4) = "SORTED ASSET A"
    TEMP_MATRIX(0, 5) = "SORTED ASSET B"
    TEMP_MATRIX(0, 6) = "PERCENTILE"
    TEMP_MATRIX(0, 7) = "P(A)"
    TEMP_MATRIX(0, 8) = "P(B)"
    TEMP_MATRIX(0, 9) = "P(A) >= P(B)"
    TEMP_MATRIX(0, 10) = "Cum P(A)"
    TEMP_MATRIX(0, 11) = "Cum P(B)"
    TEMP_MATRIX(0, 12) = "Cum P(A) >= Cum P(B)"
    TEMP_MATRIX(0, 13) = "E(A)"
    TEMP_MATRIX(0, 14) = "E(B)"
    TEMP_MATRIX(0, 15) = "I: E(A)>=E(B)"
    TEMP_MATRIX(0, 16) = "Cum Cum P(A)"
    TEMP_MATRIX(0, 17) = "Cum Cum P(B)"
    TEMP_MATRIX(0, 18) = "II: Cum Cum P(A) >= Cum Cum P(B)"
    TEMP_MATRIX(0, 19) = "I and II"
    
    For i = 1 To NSIZE

        TEMP_MATRIX(i, 1) = DATE_VECTOR(i + h, 1)
        TEMP_MATRIX(i, 2) = DATA1_VECTOR(i, 1)
        TEMP_MATRIX(i, 3) = DATA2_VECTOR(i, 1)
        TEMP_MATRIX(i, 4) = TEMP1_VECTOR(i, 1)
        TEMP_MATRIX(i, 5) = TEMP2_VECTOR(i, 1)
        
        If i <= PERCENTILE_VAL Then
        
            TEMP_MATRIX(i, 6) = (i - FACTOR) / PERCENTILE_VAL
            TEMP_MATRIX(i, 7) = HISTOGRAM_PERCENTILE_FUNC(TEMP1_VECTOR, TEMP_MATRIX(i, 6), 0)
            TEMP_MATRIX(i, 8) = HISTOGRAM_PERCENTILE_FUNC(TEMP2_VECTOR, TEMP_MATRIX(i, 6), 0)
            TEMP_MATRIX(i, 9) = IIf(TEMP_MATRIX(i, 7) >= TEMP_MATRIX(i, 8), True, False)
                                
            If TEMP_MATRIX(i, 9) = False Then
                FLAG1_VAL = False
                j = j + 1
            End If
    
            If i = 1 Then
                TEMP_MATRIX(i, 10) = TEMP_MATRIX(i, 7)
                TEMP_MATRIX(i, 11) = TEMP_MATRIX(i, 8)
            Else
                TEMP_MATRIX(i, 10) = TEMP_MATRIX(i - 1, 10) + TEMP_MATRIX(i, 7)
                TEMP_MATRIX(i, 11) = TEMP_MATRIX(i - 1, 11) + TEMP_MATRIX(i, 8)
            End If
    
            TEMP_MATRIX(i, 12) = IIf(TEMP_MATRIX(i, 10) >= _
                                TEMP_MATRIX(i, 11), True, False)
                                
            If TEMP_MATRIX(i, 12) = False Then
                FLAG2_VAL = False
                k = k + 1
            End If
    
            TEMP_MATRIX(i, 13) = TEMP_MATRIX(i, 6) * TEMP_MATRIX(i, 7)
            TEMP_MATRIX(i, 14) = TEMP_MATRIX(i, 6) * TEMP_MATRIX(i, 8)
    
            TEMP_MATRIX(i, 15) = IIf(TEMP_MATRIX(i, 13) >= _
                                TEMP_MATRIX(i, 14), True, False)
    
            If i = 1 Then
                TEMP_MATRIX(i, 16) = TEMP_MATRIX(i, 10)
                TEMP_MATRIX(i, 17) = TEMP_MATRIX(i, 11)
            Else
                TEMP_MATRIX(i, 16) = TEMP_MATRIX(i - 1, 16) + TEMP_MATRIX(i, 10)
                TEMP_MATRIX(i, 17) = TEMP_MATRIX(i - 1, 17) + TEMP_MATRIX(i, 11)
            End If
    
            TEMP_MATRIX(i, 18) = IIf(TEMP_MATRIX(i, 16) >= _
                                TEMP_MATRIX(i, 17), True, False)
                                
            TEMP_MATRIX(i, 19) = IIf(TEMP_MATRIX(i, 15) = _
                                TEMP_MATRIX(i, 18), True, False)
                                
            If TEMP_MATRIX(i, 19) = False Then
                FLAG3_VAL = False
                l = l + 1
            End If
        
        Else
            For m = 6 To 19: TEMP_MATRIX(i, m) = "": Next m
        End If

    Next i
    
    SUMMARY_MATRIX(1, 2) = FLAG1_VAL
    SUMMARY_MATRIX(2, 2) = FLAG2_VAL
    SUMMARY_MATRIX(3, 2) = FLAG3_VAL

    SUMMARY_MATRIX(1, 3) = 1 - j / PERCENTILE_VAL
    SUMMARY_MATRIX(2, 3) = 1 - k / PERCENTILE_VAL
    SUMMARY_MATRIX(3, 3) = 1 - l / PERCENTILE_VAL
    

'--------------------------------------------------------------------------------
Else 'Sorting Approach
'--------------------------------------------------------------------------------
    
    ReDim TEMP_MATRIX(0 To NSIZE, 1 To 12)
    TEMP_MATRIX(0, 1) = "DATE"
    TEMP_MATRIX(0, 2) = "A ASSET"
    TEMP_MATRIX(0, 3) = "B ASSET"
    TEMP_MATRIX(0, 4) = "SORTED A ASSET"
    TEMP_MATRIX(0, 5) = "SORTED B ASSET"
    TEMP_MATRIX(0, 6) = "CUM SORTED A ASSET"
    TEMP_MATRIX(0, 7) = "CUM SORTED B ASSET"
    TEMP_MATRIX(0, 8) = "CUM CUM SORTED A ASSET"
    TEMP_MATRIX(0, 9) = "CUM CUM SORTED B ASSET"
    TEMP_MATRIX(0, 10) = "FSD"
    TEMP_MATRIX(0, 11) = "SSD"
    TEMP_MATRIX(0, 12) = "TSD"
        
    For i = 1 To NSIZE
        TEMP_MATRIX(i, 1) = DATE_VECTOR(i + h, 1)
        TEMP_MATRIX(i, 2) = DATA1_VECTOR(i, 1)
        TEMP_MATRIX(i, 3) = DATA2_VECTOR(i, 1)
        TEMP_MATRIX(i, 4) = TEMP1_VECTOR(i, 1)
        TEMP_MATRIX(i, 5) = TEMP2_VECTOR(i, 1)
        If i > 1 Then
            TEMP_MATRIX(i, 6) = TEMP_MATRIX(i - 1, 6) + TEMP_MATRIX(i, 4)
            TEMP_MATRIX(i, 7) = TEMP_MATRIX(i - 1, 7) + TEMP_MATRIX(i, 5)
            TEMP_MATRIX(i, 8) = TEMP_MATRIX(i - 1, 8) + TEMP_MATRIX(i, 6)
            TEMP_MATRIX(i, 9) = TEMP_MATRIX(i - 1, 9) + TEMP_MATRIX(i, 7)
        Else
            TEMP_MATRIX(i, 6) = TEMP_MATRIX(i, 4)
            TEMP_MATRIX(i, 7) = TEMP_MATRIX(i, 5)
            TEMP_MATRIX(i, 8) = TEMP_MATRIX(i, 6)
            TEMP_MATRIX(i, 9) = TEMP_MATRIX(i, 7)
        End If
        TEMP_MATRIX(i, 10) = IIf(TEMP_MATRIX(i, 4) >= TEMP_MATRIX(i, 5), True, False)
        If TEMP_MATRIX(i, 10) = False Then
            FLAG1_VAL = False
            j = j + 1
        End If
        
        TEMP_MATRIX(i, 11) = IIf(TEMP_MATRIX(i, 6) >= TEMP_MATRIX(i, 7), True, False)
        If TEMP_MATRIX(i, 11) = False Then
            FLAG2_VAL = False
            k = k + 1
        End If
        
        TEMP_MATRIX(i, 12) = IIf(TEMP_MATRIX(i, 8) >= TEMP_MATRIX(i, 9), True, False)
        If TEMP_MATRIX(i, 12) = False Then
            FLAG2_VAL = False
            l = l + 1
        End If
    Next i
    
    SUMMARY_MATRIX(1, 2) = FLAG1_VAL
    SUMMARY_MATRIX(2, 2) = FLAG2_VAL
    SUMMARY_MATRIX(3, 2) = FLAG3_VAL

    SUMMARY_MATRIX(1, 3) = 1 - j / NSIZE
    SUMMARY_MATRIX(2, 3) = 1 - k / NSIZE
    SUMMARY_MATRIX(3, 3) = 1 - l / NSIZE

'--------------------------------------------------------------------------------
End If
'--------------------------------------------------------------------------------

Select Case OUTPUT
    Case 0
        ASSET_STOCHASTIC_DOMINANCE_FUNC = SUMMARY_MATRIX
    Case Else
        ASSET_STOCHASTIC_DOMINANCE_FUNC = TEMP_MATRIX
End Select

Exit Function
ERROR_LABEL:
ASSET_STOCHASTIC_DOMINANCE_FUNC = Err.number
End Function
