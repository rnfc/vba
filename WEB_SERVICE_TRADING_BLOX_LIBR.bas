Attribute VB_Name = "WEB_SERVICE_TRADING_BLOX_LIBR"


Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.

Function TRADING_BLOCKS_DATABASE_FUNC(ByRef TICKERS_RNG As Variant, _
Optional ByVal START_DATE As Date = 0, _
Optional ByVal END_DATE As Date = 0, _
Optional ByVal FOLDER_NAME_STR As String = "")

'Call TRADING_BLOCKS_DATABASE_FUNC(Range("NICO"), , , "C:\Users\nfermincota\Desktop\DATABASE\SP")

Dim i As Long
Dim j As Long
Dim k As Long
Dim ii As Long
Dim jj As Long
Dim NROWS As Long
Dim DELIM_STR As String
Dim TICKER_STR As String
Dim PATH_SEP_STR As String
Dim FILE_PATH_STR As String
Dim EXTENSION_STR As String

Dim TEMP_MATRIX As Variant
Dim TICKERS_VECTOR As Variant

On Error GoTo ERROR_LABEL

DELIM_STR = ","
EXTENSION_STR = ".txt"
PATH_SEP_STR = Excel.Application.PathSeparator
If FOLDER_NAME_STR = "" Then: FOLDER_NAME_STR = Excel.Application.Path
If END_DATE = 0 Then: END_DATE = Now()
If START_DATE = 0 Then: START_DATE = DateSerial(Year(END_DATE) - 10, Month(END_DATE), Day(END_DATE))

If IsArray(TICKERS_RNG) = True Then
    TICKERS_VECTOR = TICKERS_RNG
    If UBound(TICKERS_VECTOR, 1) = 1 Then
        TICKERS_VECTOR = MATRIX_TRANSPOSE_FUNC(TICKERS_VECTOR)
    End If
Else
    ReDim TICKERS_VECTOR(1 To 1, 1 To 1)
    TICKERS_VECTOR(1, 1) = TICKERS_RNG
End If
NROWS = UBound(TICKERS_VECTOR, 1)

For k = 1 To NROWS
    TICKER_STR = TICKERS_VECTOR(k, 1)
    TEMP_MATRIX = YAHOO_HISTORICAL_DATA_SERIE_FUNC(TICKER_STR, _
                  START_DATE, END_DATE, "DAILY", "DOHLCV", _
                  False, True, True)
    If IsArray(TEMP_MATRIX) = False Then: GoTo 1983
    ii = UBound(TEMP_MATRIX, 1)
    jj = UBound(TEMP_MATRIX, 2)
    For i = 1 To ii
        For j = 1 To jj
            If j <> 1 Then
                If j <> 6 Then
                    TEMP_MATRIX(i, j) = Format(TEMP_MATRIX(i, j), "0.00")
                Else
                    TEMP_MATRIX(i, j) = TEMP_MATRIX(i, j) / 100
                    TEMP_MATRIX(i, j) = Format(TEMP_MATRIX(i, j), "0")
                End If
            Else
                TEMP_MATRIX(i, j) = Format(TEMP_MATRIX(i, j), "yyyymmdd")
            End If
        Next j
    Next i
    FILE_PATH_STR = FOLDER_NAME_STR & PATH_SEP_STR & TICKER_STR & EXTENSION_STR
    Call CONVERT_MATRIX_TEXT_FILE_FUNC(FILE_PATH_STR, TEMP_MATRIX, DELIM_STR, 0)
1983:
Next k

TRADING_BLOCKS_DATABASE_FUNC = True

Exit Function
ERROR_LABEL:
TRADING_BLOCKS_DATABASE_FUNC = False
End Function

