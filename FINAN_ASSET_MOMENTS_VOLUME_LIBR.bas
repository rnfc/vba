Attribute VB_Name = "FINAN_ASSET_MOMENTS_VOLUME_LIBR"


Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.
                    
'daily volume of stock trades and (as usual) compared it to the average daily volume
'and shares outstanding

Function ASSETS_VOLUME_SHARES_OUTSTANDING_FUNC(ByRef TICKERS_RNG As Variant, _
Optional ByVal REFRESH_CALLER As Variant, _
Optional ByVal SERVER_STR As String = "UNITED STATES")

Dim i As Long
Dim NROWS As Long
Const FACTOR_VAL As Double = 1000000 'millions
Dim DATA_MATRIX As Variant
Dim TICKERS_VECTOR As Variant

On Error GoTo ERROR_LABEL

If IsArray(TICKERS_RNG) = True Then
    TICKERS_VECTOR = TICKERS_RNG
    If UBound(TICKERS_VECTOR, 1) = 1 Then
        TICKERS_VECTOR = MATRIX_TRANSPOSE_FUNC(TICKERS_VECTOR)
    End If
Else
    ReDim TICKERS_VECTOR(1 To 1, 1 To 1)
    TICKERS_VECTOR(1, 1) = TICKERS_RNG
End If

ReDim DATA_MATRIX(1 To 1, 1 To 4)
DATA_MATRIX(1, 1) = "name"
DATA_MATRIX(1, 2) = "last trade"
DATA_MATRIX(1, 3) = "average daily volume"
DATA_MATRIX(1, 4) = "market capitalization"

DATA_MATRIX = YAHOO_QUOTES_FUNC(TICKERS_VECTOR, DATA_MATRIX, REFRESH_CALLER, True, SERVER_STR)
NROWS = UBound(DATA_MATRIX, 1)

ReDim Preserve DATA_MATRIX(LBound(DATA_MATRIX) To UBound(DATA_MATRIX), 1 To 7)
DATA_MATRIX(0, 6) = "shares outstanding"
DATA_MATRIX(0, 7) = "avg.vol / outstanding"
For i = 1 To NROWS
    DATA_MATRIX(i, 5) = DATA_MATRIX(i, 5) / FACTOR_VAL 'in billions
    If DATA_MATRIX(i, 3) <> 0 Then
        DATA_MATRIX(i, 6) = DATA_MATRIX(i, 5) / DATA_MATRIX(i, 3)
    End If
    If DATA_MATRIX(i, 6) <> 0 Then
        DATA_MATRIX(i, 7) = DATA_MATRIX(i, 4) / DATA_MATRIX(i, 6) / (FACTOR_VAL * 1000)
    End If
Next i
ASSETS_VOLUME_SHARES_OUTSTANDING_FUNC = DATA_MATRIX

Exit Function
ERROR_LABEL:
ASSETS_VOLUME_SHARES_OUTSTANDING_FUNC = Err.number
End Function



