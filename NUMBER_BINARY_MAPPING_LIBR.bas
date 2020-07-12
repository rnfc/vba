Attribute VB_Name = "NUMBER_BINARY_MAPPING_LIBR"


Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.

'************************************************************************************
'************************************************************************************
'FUNCTION      : BINARY_MAPPING_FUNC

'DESCRIPTION   : This function returns the binary linear function
'This function is useful to calculate several digital functions
'for every combination of independent variables (laso called input variables)

'LIBRARY       : NUMBER_BINARY
'GROUP         : MAPPING
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 14/02/2008
'************************************************************************************
'************************************************************************************


Function BINARY_MAPPING_FUNC(ByRef DATA_RNG As Variant, _
Optional ByVal VERSION As Integer = 0, _
Optional ByVal LOGIC_FLAG As Boolean = False, _
Optional ByVal AND_FLAG As Boolean = False, _
Optional ByVal SWITCH_FLAG As Boolean = False, _
Optional ByVal OUTPUT As Integer = 0)

'DATA_RNG: following digital network
'Each element of the matrix must be “1” if the corresponding
'variable x is mapped as is; “-1” if the corresponding variable
'is complemented; “0” if the variable is no mapped at all.

'VERSION: If you like, you can indicate if
'the table will be the least significant digit column
'to the right or to the left.

'The function can map 3 different kind
'of logic functions:

'1) Logical XOR
'2) Logical AND
'3) Logical OR

Dim i As Long
Dim j As Long
Dim k As Long

Dim NROWS As Long
Dim NCOLUMNS As Long

Dim XTEMP_MATRIX As Variant
Dim YTEMP_MATRIX As Variant

Dim DATA_MATRIX As Variant

On Error GoTo ERROR_LABEL

DATA_MATRIX = DATA_RNG
'Each variable can be complemeted or not depending its corresponding
'coefficients of the mappin matrix M. The mapping matrix matrix M must
'have dimension (k x n), where “n” must be the number of x variable,
'“k” is number of y digits, or the number of implicants.


XTEMP_MATRIX = BINARY_GENERATOR_FUNC(UBound(DATA_MATRIX, 2), VERSION)

If LOGIC_FLAG = True Then
    YTEMP_MATRIX = BINARY_MATRIX_MULT_FUNC(DATA_MATRIX, XTEMP_MATRIX)
    
    'This routine returns a digit variable Y for each row of the mapping
    'matrix. It performes the linear combination of n binary variables
Else
    k = 1
    If AND_FLAG = True Then: k = 0
    YTEMP_MATRIX = BINARY_LOGIC_FUNC(DATA_MATRIX, XTEMP_MATRIX, k)
     'This routine easily map the logic function for each input
    'configuration.
End If

'----------------------------------------------------------------------
Select Case OUTPUT
'----------------------------------------------------------------------
Case 0 ' 1-0 logic symbol
'----------------------------------------------------------------------
    'fast fill
    NROWS = UBound(XTEMP_MATRIX, 1)
    NCOLUMNS = UBound(XTEMP_MATRIX, 2)
    If SWITCH_FLAG = True Then
        For i = 1 To NROWS
            For j = 1 To NCOLUMNS
                If XTEMP_MATRIX(i, j) = 0 Then XTEMP_MATRIX(i, j) = False
                If XTEMP_MATRIX(i, j) = 1 Then XTEMP_MATRIX(i, j) = True
            Next j
        Next i
    End If
    BINARY_MAPPING_FUNC = XTEMP_MATRIX
'----------------------------------------------------------------------
Case Else
'----------------------------------------------------------------------
    NROWS = UBound(YTEMP_MATRIX, 1)
    NCOLUMNS = UBound(YTEMP_MATRIX, 2)
    If SWITCH_FLAG = True Then
        For i = 1 To NROWS
            For j = 1 To NCOLUMNS
                If YTEMP_MATRIX(i, j) = 0 Then YTEMP_MATRIX(i, j) = False
                If YTEMP_MATRIX(i, j) = 1 Then YTEMP_MATRIX(i, j) = True
            Next j
        Next i
    End If
'----------------------------------------------------------------------
BINARY_MAPPING_FUNC = YTEMP_MATRIX
'----------------------------------------------------------------------
'----------------------------------------------------------------------------
'complete mapping table, where, for each
'number x, there is the corresponding value
'y, given by the mapping matrix
'----------------------------------------------------------------------------
End Select


Exit Function
ERROR_LABEL:
BINARY_MAPPING_FUNC = Err.number
End Function

