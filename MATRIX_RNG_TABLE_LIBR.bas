Attribute VB_Name = "MATRIX_RNG_TABLE_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.

'************************************************************************************
'************************************************************************************
'FUNCTION      : RNG_MATRIX_DATA_TABLE_FUNC
'DESCRIPTION   : Creates a data table based on input values and formulas that you
'define on a worksheet
'LIBRARY       : RANGE
'GROUP         : RNG_TABLE
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/19/2009
'************************************************************************************
'************************************************************************************

Function RNG_MATRIX_DATA_TABLE_FUNC(ByVal VERSION As Long, _
ByRef DST_RNG As Excel.Range, _
ByVal RMIN_VAL As Double, _
ByVal RMAX_VAL As Double, _
ByVal RDELTA_VAL As Double, _
ByRef FUNCTION_RNG As Excel.Range, _
Optional ByRef ROW_RNG As Excel.Range, _
Optional ByRef COLUMN_RNG As Excel.Range, _
Optional ByVal CMIN_VAL As Double = 0, _
Optional ByVal CMAX_VAL As Double = 0, _
Optional ByVal CDELTA_VAL As Double = 0)

Dim i As Long
Dim j As Long
Dim NROWS As Long
Dim NCOLUMNS As Long

Dim DATA_RNG As Excel.Range

On Error GoTo ERROR_LABEL

RNG_MATRIX_DATA_TABLE_FUNC = False
Select Case VERSION
'-------------------------------------------------------------------------------------
    Case 0 'Using ROW_RNG
'-------------------------------------------------------------------------------------
        NCOLUMNS = (CMAX_VAL - CMIN_VAL) / CDELTA_VAL
        DST_RNG.Offset(1, 0) = "=" & FUNCTION_RNG.Address
        DST_RNG.Offset(0, 1) = CMIN_VAL
            For i = 1 To NCOLUMNS
                DST_RNG.Offset(0, i + 1).formula = "=" & DST_RNG.Offset(0, i).Address _
                & "+" & RDELTA_VAL
            Next i
        Set DATA_RNG = Range(DST_RNG, DST_RNG.Offset(1, NCOLUMNS + 1))
        DATA_RNG.Table ROW_RNG
'-------------------------------------------------------------------------------------
    Case 1 'Using COLUMN_RNG
'-------------------------------------------------------------------------------------
        NROWS = (RMAX_VAL - RMIN_VAL) / RDELTA_VAL
        DST_RNG.Offset(0, 1) = "=" & FUNCTION_RNG.Address
        DST_RNG.Offset(1, 0) = RMIN_VAL
            For i = 1 To NROWS
                DST_RNG.Offset(i + 1, 0).formula = "=" & DST_RNG.Offset(i, 0).Address _
                & "+" & RDELTA_VAL
            Next i
        Set DATA_RNG = Range(DST_RNG, DST_RNG.Offset(NROWS + 1, 1))
        DATA_RNG.Table , COLUMN_RNG
'-------------------------------------------------------------------------------------
    Case Else 'Using ROW_RNG & COLUMN_RNG
'-------------------------------------------------------------------------------------
        NCOLUMNS = (CMAX_VAL - CMIN_VAL) / CDELTA_VAL
        NROWS = (RMAX_VAL - RMIN_VAL) / RDELTA_VAL
        
        DST_RNG = "=" & FUNCTION_RNG.Address
        
        DST_RNG.Offset(1, 0) = RMIN_VAL
        DST_RNG.Offset(0, 1) = CMIN_VAL
        
            For i = 1 To NROWS
                DST_RNG.Offset(i + 1, 0).formula = "=" & _
                DST_RNG.Offset(i, 0).Address & "+" & CDELTA_VAL
            Next i
            
            For j = 1 To NCOLUMNS
                DST_RNG.Offset(0, 1 + j).formula = "=" & _
                DST_RNG.Offset(0, j).Address & "+" & RDELTA_VAL
            Next j
        
        Set DATA_RNG = Range(DST_RNG, DST_RNG.Offset(NROWS + 1, NCOLUMNS + 1))
        DATA_RNG.Table ROW_RNG, COLUMN_RNG
'-------------------------------------------------------------------------------------
End Select
'-------------------------------------------------------------------------------------

RNG_MATRIX_DATA_TABLE_FUNC = True

Exit Function
ERROR_LABEL:
RNG_MATRIX_DATA_TABLE_FUNC = False
End Function

