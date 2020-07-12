Attribute VB_Name = "STAT_MOMENTS_QUATERLY_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.

'************************************************************************************
'************************************************************************************
'FUNCTION      : MONTHLY_QUATERLY_DATA_FUNC
'DESCRIPTION   : Converting Monthly Data to Quarterly Data
'LIBRARY       : STATISTICS
'GROUP         :
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 14/02/2008
'************************************************************************************
'************************************************************************************

Function MONTHLY_QUATERLY_DATA_FUNC(ByRef DATE_RNG As Variant, _
ByRef DATA_RNG As Variant, _
Optional ByVal OUTPUT As Integer = 0)

Dim i As Long
Dim j As Long
Dim k As Long
Dim l() As Long
Dim NROWS As Long

Dim DATA_VECTOR As Variant
Dim DATE_VECTOR As Variant

Dim TEMP_MATRIX As Variant

On Error GoTo ERROR_LABEL

DATE_VECTOR = DATE_RNG 'Monthly
If UBound(DATE_VECTOR, 1) = 1 Then
    DATE_VECTOR = MATRIX_TRANSPOSE_FUNC(DATE_VECTOR)
End If
NROWS = UBound(DATE_VECTOR, 1)
DATA_VECTOR = DATA_RNG 'Monthly
If UBound(DATA_VECTOR, 1) = 1 Then
    DATA_VECTOR = MATRIX_TRANSPOSE_FUNC(DATA_VECTOR)
End If
If NROWS <> UBound(DATA_VECTOR, 1) Then: GoTo ERROR_LABEL

'-----------------------------------------------------------------------------------------
Select Case OUTPUT
'-----------------------------------------------------------------------------------------
Case 0
'-----------------------------------------------------------------------------------------
    k = 1: ReDim l(1 To k): l(k) = 2
    j = 0
    For i = 2 To NROWS
        If j = 3 Then
            j = 0
            k = k + 1
            ReDim Preserve l(1 To k)
            l(k) = i
        End If
        j = j + 1
    Next i
    ReDim TEMP_MATRIX(0 To k, 1 To 3)
    TEMP_MATRIX(0, 1) = "MONTH"
    TEMP_MATRIX(0, 2) = "DATE"
    TEMP_MATRIX(0, 3) = "QTRLY AVE"
    For i = 1 To k
        j = l(i)
        TEMP_MATRIX(i, 1) = j - 1
        TEMP_MATRIX(i, 2) = DATE_VECTOR(j, 1)
        If j < NROWS Then
            TEMP_MATRIX(i, 3) = (DATA_VECTOR(j - 1, 1) + DATA_VECTOR(j, 1) + DATA_VECTOR(j + 1, 1)) * (1 / 3)
        Else
            TEMP_MATRIX(i, 3) = (DATA_VECTOR(j - 1, 1) + DATA_VECTOR(j, 1) + 0) * (1 / 2)
        End If
    Next i
'-----------------------------------------------------------------------------------------
Case Else
'-----------------------------------------------------------------------------------------
    ReDim TEMP_MATRIX(0 To NROWS, 1 To 3)
    i = 0
    TEMP_MATRIX(i, 1) = "DATE"
    TEMP_MATRIX(i, 2) = "DATA"
    TEMP_MATRIX(i, 3) = "3M-AVE"
    i = 1
    TEMP_MATRIX(i, 1) = DATE_VECTOR(i, 1)
    TEMP_MATRIX(i, 2) = DATA_VECTOR(i, 1)
    TEMP_MATRIX(i, 3) = 0
    For i = 2 To NROWS
        TEMP_MATRIX(i, 1) = DATE_VECTOR(i, 1)
        TEMP_MATRIX(i, 2) = DATA_VECTOR(i, 1)
        If i < NROWS Then 'The same method could be applied to obtain quarterly sums, using the SUM function
        'instead of the AVERAGE function.
            TEMP_MATRIX(i, 3) = (DATA_VECTOR(i - 1, 1) + DATA_VECTOR(i, 1) + DATA_VECTOR(i + 1, 1)) * (1 / 3)
        Else
            TEMP_MATRIX(i, 3) = (DATA_VECTOR(i - 1, 1) + DATA_VECTOR(i, 1) + 0) * (1 / 2)
        End If
    Next i
'-----------------------------------------------------------------------------------------
End Select
'-----------------------------------------------------------------------------------------

MONTHLY_QUATERLY_DATA_FUNC = TEMP_MATRIX

Exit Function
ERROR_LABEL:
MONTHLY_QUATERLY_DATA_FUNC = Err.number
End Function


