Attribute VB_Name = "WEB_SERVICE_OANDA_LIBR"

'--------------------------------------------------------------------------------------------------------------
'--------------------------------------------------------------------------------------------------------------
Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.
'--------------------------------------------------------------------------------------------------------------
'--------------------------------------------------------------------------------------------------------------
Private PUB_OANDA_HASH_OBJ As clsTypeHash
'--------------------------------------------------------------------------------------------------------------
'--------------------------------------------------------------------------------------------------------------

'************************************************************************************
'************************************************************************************
'FUNCTION      : OANDA_HISTORICAL_FX_RATES_FUNC

'DESCRIPTION   : Tool to obtain historical exchange rate for any currency pair,
'select the range of dates and the currencies you would like to obtain exchange
'rates for. You can obtain the historical exchange rates with the desired rate
'(cash, interbank, credit card) by adding a % point. Adjust this for higher periods
'499 points

'The download operation is efficient: only one request is made even if
'several symbols are requested at once. The return value is an array,
'with the following elements: Bid, Ask, and last traded price.

'ID            : 001
'LAST UPDATE   : 12/05/2011
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'************************************************************************************
'************************************************************************************

Function OANDA_HISTORICAL_FX_RATES_FUNC(ByVal START_DATE As Date, _
ByVal END_DATE As Date, _
Optional ByVal FROM_FX_STR As String = "USD", _
Optional ByVal TO_FX_STR As String = "AUD", _
Optional ByVal PERCENT As Double = -0.05)

'Interbank rate +0%
'Typical credit card rate: +2%
'Typical cash rate: +4%
'Interbank rate: +1%
'Interbank rate: +3%
'Interbank rate: +5%
'Interbank rate: +6%
'Interbank rate: +7%
'Interbank rate: +8%
'Interbank rate: +9%
'Interbank rate: +10%
 
Dim i As Long
Dim j As Long
Dim k As Long

Dim ii As Long
Dim jj As Long
Dim kk As Long

Dim SROW As Long
Dim NROWS As Long

Dim SDATE_VAL As Date
Dim EDATE_VAL As Date

Dim CHR_STR As String
Dim TEMP_STR As String
Dim DATA_STR As String
Dim SERVER_STR As String
Dim DELIM_STR As String

Dim SRC_URL_STR As String
Dim TEMP_VECTOR As Variant

On Error GoTo ERROR_LABEL

If START_DATE >= END_DATE Then: GoTo ERROR_LABEL
k = 499 'periods limit
DELIM_STR = ","
SERVER_STR = "http://www.oanda.com/convert/fxhistory?date_fmt=us&date="
DATA_STR = ""
    
If DateDiff("d", START_DATE, END_DATE) > k Then
    SDATE_VAL = START_DATE
    EDATE_VAL = SDATE_VAL + k
    Do
        GoSub RETRIEVE_LINE
        DATA_STR = DATA_STR & TEMP_STR
        SDATE_VAL = EDATE_VAL + 1
        EDATE_VAL = SDATE_VAL + k
    Loop Until EDATE_VAL >= END_DATE
    EDATE_VAL = END_DATE
    GoSub RETRIEVE_LINE
    DATA_STR = DATA_STR & TEMP_STR
Else
    SDATE_VAL = START_DATE
    EDATE_VAL = END_DATE
    GoSub RETRIEVE_LINE
    DATA_STR = DATA_STR & TEMP_STR
End If

'DATA_STR = Split(DATA_STR, Chr(10))
'SROW = LBound(DATA_STR)
'NROWS = UBound(DATA_STR) '- SROW + 1
    
CHR_STR = Chr(10)
kk = Len(CHR_STR)
NROWS = COUNT_CHARACTERS_FUNC(DATA_STR, CHR_STR)
ReDim TEMP_VECTOR(0 To NROWS, 1 To 2)
TEMP_VECTOR(0, 1) = "DATE"
TEMP_VECTOR(0, 2) = FROM_FX_STR & " to " & TO_FX_STR

ii = 1
For k = 1 To NROWS
    jj = InStr(ii, DATA_STR, CHR_STR)
    TEMP_STR = Mid(DATA_STR, ii, jj - ii)
    ii = jj + kk
    
    i = 1: j = InStr(i, TEMP_STR, DELIM_STR)
    TEMP_VECTOR(k, 1) = OANDA_PARSE_DATE_VAL(Trim(Mid(TEMP_STR, i, j - i)))
    i = j + 1: j = Len(TEMP_STR)
    TEMP_VECTOR(k, 2) = CDbl(Trim(Mid(TEMP_STR, i, j - i))) * (1 + PERCENT)
    SROW = SROW + 1
Next k

OANDA_HISTORICAL_FX_RATES_FUNC = TEMP_VECTOR

Exit Function
'------------------------------------------------------------------------------------------------
RETRIEVE_LINE:
'------------------------------------------------------------------------------------------------
    SRC_URL_STR = SERVER_STR & Format(EDATE_VAL, "mm/dd/yy") & "&date1=" & _
        Format(SDATE_VAL, "mm/dd/yy") & "&exch=" & FROM_FX_STR & "&expr=" & _
        TO_FX_STR & "&lang=en&margin_fixed=0&format=CSV&redirected=1"
    
    TEMP_STR = SAVE_WEB_DATA_PAGE_FUNC(SRC_URL_STR, 0, False, 0, False)
    CHR_STR = "Conversion Table:"
    i = InStr(1, TEMP_STR, CHR_STR)
    If i = 0 Then: GoTo ERROR_LABEL
    i = i + Len(CHR_STR)
    
    CHR_STR = "<PRE>"
    i = InStr(1, TEMP_STR, CHR_STR)
    If i = 0 Then: GoTo ERROR_LABEL
    i = i + Len(CHR_STR)
    
    CHR_STR = "</PRE>"
    j = InStr(i, TEMP_STR, CHR_STR)
    If j = 0 Then: GoTo ERROR_LABEL
    
    TEMP_STR = Mid(TEMP_STR, i, j - i)
'------------------------------------------------------------------------------------------------
Return
'------------------------------------------------------------------------------------------------
ERROR_LABEL:
OANDA_HISTORICAL_FX_RATES_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : OANDA_HISTORICAL_DATA_FUNC

'DESCRIPTION   : Tool to obtain historical exchange rate for any currency pair,
'select the range of dates and the currencies you would like to obtain exchange
'rates for. The download operation is efficient: uses a hash table to save each
'data set on the memory ram and avoid making several requests. The return value
'is an array.

'LIBRARY       : OANDA
'GROUP         : WEB SERVICE
'ID            : 002
'LAST UPDATE   : 12/05/2011
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'************************************************************************************
'************************************************************************************

Function OANDA_HISTORICAL_DATA_FUNC(ByVal START_DATE As Date, _
ByVal END_DATE As Date, _
ByRef FROM_FX_RNG As Variant, _
ByRef TO_FX_RNG As Variant, _
Optional ByVal HEADERS_FLAG As Boolean = True, _
Optional ByVal RESORT_FLAG As Boolean = True)

'IF RESORT_FLAG = True Then: Ascending Order else Descending
'AUDCAD,AUDJPY,AUDUSD,EURAUD,EURCAD,EURCHF,EURGBP,EURJPY,EURUSD,GBPJPY,GBPUSD,USDJPY,USDCHF,USDCAD,GBPCHF

Dim h As Long
Dim i As Long
Dim j As Long
Dim k As Long
Dim l As Long
Dim m As Long

Dim hh As Long
Dim ii As Long
Dim jj As Long
Dim kk As Long

Dim NROWS As Long
Dim NCOLUMNS As Long

Dim KEY_STR As String
Dim CHR_STR As String
Dim DELIM_STR As String

Dim SDATE_VAL As Date
Dim EDATE_VAL As Date

Dim DATE_VAL As String
Dim DATA_VAL As String
Dim DATA_STR As String
Dim FROM_FX_STR As String
Dim TO_FX_STR As String
Dim SRC_URL_STR As String

Dim TEMP_STR As String
Dim LEFT_STR As String
Dim RIGHT_STR As String
Dim LINE_STR As String
Dim SERVER_STR As String

Dim DATE_VECTOR As Variant
Dim TEMP_MATRIX As Variant
Dim FROM_FX_VECTOR As Variant
Dim TO_FX_VECTOR As Variant

Dim DATA_OBJ As clsTypeHash

On Error GoTo ERROR_LABEL

If PUB_OANDA_HASH_OBJ Is Nothing Then
    Set PUB_OANDA_HASH_OBJ = New clsTypeHash
    PUB_OANDA_HASH_OBJ.SetSize 10000
    PUB_OANDA_HASH_OBJ.IgnoreCase = False
End If

m = 499: DELIM_STR = ","
SERVER_STR = "http://www.oanda.com/convert/fxhistory?date_fmt=us&date="
   
If IsArray(FROM_FX_RNG) = True Then
    FROM_FX_VECTOR = FROM_FX_RNG
    If UBound(FROM_FX_VECTOR, 1) = 1 Then
        FROM_FX_VECTOR = MATRIX_TRANSPOSE_FUNC(FROM_FX_VECTOR)
    End If
Else
    ReDim FROM_FX_VECTOR(1 To 1, 1 To 1)
    FROM_FX_VECTOR(1, 1) = FROM_FX_RNG
End If
NCOLUMNS = UBound(FROM_FX_VECTOR, 1)

If IsArray(TO_FX_RNG) = True Then
    TO_FX_VECTOR = TO_FX_RNG
    If UBound(TO_FX_VECTOR, 1) = 1 Then
        TO_FX_VECTOR = MATRIX_TRANSPOSE_FUNC(TO_FX_VECTOR)
    End If
Else
    ReDim TO_FX_VECTOR(1 To 1, 1 To 1)
    TO_FX_VECTOR(1, 1) = TO_FX_RNG
End If

If UBound(FROM_FX_VECTOR, 1) <> UBound(TO_FX_VECTOR, 1) Then: GoTo ERROR_LABEL

LINE_STR = DELIM_STR
KEY_STR = Year(START_DATE) & Month(START_DATE) & Day(START_DATE) & "|" & Year(END_DATE) & Month(END_DATE) & Day(END_DATE)
For jj = 1 To NCOLUMNS
    LINE_STR = LINE_STR & "0" & DELIM_STR
    KEY_STR = KEY_STR & "|" & FROM_FX_VECTOR(jj, 1) & "|" & TO_FX_VECTOR(jj, 1)
Next jj
KEY_STR = KEY_STR & "|" & CStr(HEADERS_FLAG) & "|" & CStr(RESORT_FLAG) & "|"

If PUB_OANDA_HASH_OBJ.Exists(KEY_STR) = True Then
    TEMP_MATRIX = PUB_OANDA_HASH_OBJ(KEY_STR)
    OANDA_HISTORICAL_DATA_FUNC = CONVERT_STRING_NUMBER_FUNC(TEMP_MATRIX)
    Exit Function
End If

Set DATA_OBJ = New clsTypeHash
DATA_OBJ.SetSize 10000
DATA_OBJ.IgnoreCase = False

hh = 0
For jj = 1 To NCOLUMNS
    FROM_FX_STR = FROM_FX_VECTOR(jj, 1)
    TO_FX_STR = TO_FX_VECTOR(jj, 1)
    GoSub DATA_LINE
    h = 1
    For ii = 1 To NROWS
        j = InStr(h, DATA_STR, DELIM_STR)
        If j = 0 Then Exit For
        DATE_VAL = Trim(Mid(DATA_STR, h, j - h))
        DATE_VAL = OANDA_PARSE_DATE_VAL(DATE_VAL)
                
        If DATA_OBJ.Exists(DATE_VAL) = False Then
            TEMP_STR = DATE_VAL & LINE_STR
            Call DATA_OBJ.Add(DATE_VAL, TEMP_STR)
            hh = hh + 1
        End If
        
        i = j + 1
        DATA_VAL = "0"
        j = InStr(i, DATA_STR, DELIM_STR)
        If j = 0 Then GoTo 1982
        DATA_VAL = Trim(Mid(DATA_STR, i, j - i))
        i = j + 1

        If DATA_VAL <> "0" Then
            TEMP_STR = DATA_OBJ.Item(DATE_VAL)
            
            i = 1
            j = InStr(i, TEMP_STR, DELIM_STR)
            If j = 0 Then GoTo 1982
            i = j + 1
            For kk = 1 To jj
                LEFT_STR = Mid(TEMP_STR, 1, i - 1)
                j = InStr(i, TEMP_STR, DELIM_STR)
                If j = 0 Then GoTo 1982
                RIGHT_STR = Mid(TEMP_STR, j, Len(TEMP_STR) - j + 1)
                i = j + 1
            Next kk
            
            DATA_OBJ.Item(DATE_VAL) = LEFT_STR & DATA_VAL & RIGHT_STR
        End If
1982:
        h = InStr(h, DATA_STR, CHR_STR)
        If h = 0 Then: Exit For
        h = h + 1
    Next ii

1983:
Next jj
If hh = 0 Then: GoTo ERROR_LABEL
GoSub DATES_LINE
GoSub HEADERS_LINE
GoSub HASH_LINE

Set DATA_OBJ = Nothing
Call PUB_OANDA_HASH_OBJ.Add(KEY_STR, TEMP_MATRIX)
OANDA_HISTORICAL_DATA_FUNC = TEMP_MATRIX

Exit Function
'------------------------------------------------------------------------------------------------
RETRIEVE_LINE:
'------------------------------------------------------------------------------------------------
    SRC_URL_STR = SERVER_STR & Format(EDATE_VAL, "mm/dd/yy") & "&date1=" & _
        Format(SDATE_VAL, "mm/dd/yy") & "&exch=" & FROM_FX_STR & "&expr=" & _
        TO_FX_STR & "&lang=en&margin_fixed=0&format=CSV&redirected=1"
    
    TEMP_STR = SAVE_WEB_DATA_PAGE_FUNC(SRC_URL_STR, 0, False, 0, False)
    
    CHR_STR = "Conversion Table:"
    i = InStr(1, TEMP_STR, CHR_STR)
    If i = 0 Then
        TEMP_STR = ""
        Return
    End If
    i = i + Len(CHR_STR)
    
    CHR_STR = "<PRE>"
    i = InStr(1, TEMP_STR, CHR_STR)
    If i = 0 Then: GoTo ERROR_LABEL
    i = i + Len(CHR_STR)
    
    CHR_STR = "</PRE>"
    j = InStr(i, TEMP_STR, CHR_STR)
    If j = 0 Then: GoTo ERROR_LABEL
    
    TEMP_STR = Trim(Mid(TEMP_STR, i, j - i))
'------------------------------------------------------------------------------------------------
Return
'------------------------------------------------------------------------------------------------
DATA_LINE:
'------------------------------------------------------------------------------------------------
    DATA_STR = ""
    If DateDiff("d", START_DATE, END_DATE) > m Then
        SDATE_VAL = START_DATE
        EDATE_VAL = SDATE_VAL + m
        Do
            GoSub RETRIEVE_LINE
            DATA_STR = DATA_STR & TEMP_STR
            SDATE_VAL = EDATE_VAL + 1
            EDATE_VAL = SDATE_VAL + m
        Loop Until EDATE_VAL >= END_DATE
        EDATE_VAL = END_DATE
        GoSub RETRIEVE_LINE
        DATA_STR = DATA_STR & TEMP_STR
    Else
        SDATE_VAL = START_DATE
        EDATE_VAL = END_DATE
        GoSub RETRIEVE_LINE
        DATA_STR = DATA_STR & TEMP_STR
    End If
    CHR_STR = Chr(10)
    TEMP_STR = ""
    DATA_STR = Replace(DATA_STR, CHR_STR, DELIM_STR & CHR_STR)
    'Debug.Print DATA_STR
    NROWS = COUNT_CHARACTERS_FUNC(DATA_STR, CHR_STR)
'------------------------------------------------------------------------------------------------
Return
'------------------------------------------------------------------------------------------------
DATES_LINE:
'------------------------------------------------------------------------------------------------
    ReDim DATE_VECTOR(1 To hh, 1 To 1)
    For ii = 1 To hh
        h = ii - 1
        DATE_VAL = DATA_OBJ.GetKey(h)
        DATE_VECTOR(ii, 1) = CDate(DATE_VAL)
    Next ii
    DATE_VECTOR = MATRIX_QUICK_SORT_FUNC(DATE_VECTOR, 1, IIf(RESORT_FLAG = True, 1, 0))
'------------------------------------------------------------------------------------------------
Return
'------------------------------------------------------------------------------------------------
HEADERS_LINE:
'------------------------------------------------------------------------------------------------
    If HEADERS_FLAG = True Then
        l = 0
        ReDim TEMP_MATRIX(l To hh, 1 To NCOLUMNS + 1)
        TEMP_MATRIX(l, 1) = "DATES"
        For jj = 1 To NCOLUMNS
            TEMP_MATRIX(l, jj + 1) = FROM_FX_VECTOR(jj, 1) & " to " & TO_FX_VECTOR(jj, 1)
        Next jj
    Else
        l = 1
        ReDim TEMP_MATRIX(l To hh, 1 To NCOLUMNS + 1)
    End If
'-------------------------------------------------------------------------------------------------
Return
'-------------------------------------------------------------------------------------------------
HASH_LINE:
'-------------------------------------------------------------------------------------------------
    For ii = 1 To hh
        TEMP_MATRIX(ii, 1) = DATE_VECTOR(ii, 1)
        DATE_VAL = DATE_VECTOR(ii, 1)
        DATA_VAL = DATA_OBJ.Item(DATE_VAL)
        k = Len(DATA_VAL)
        i = InStr(1, DATA_VAL, DELIM_STR)
        If i = 0 Then: GoTo 1984
        i = i + 1
        jj = 1
        Do
            j = InStr(i, DATA_VAL, DELIM_STR)
            If j = 0 Then: Exit Do
            TEMP_MATRIX(ii, jj + 1) = CDec(Mid(DATA_VAL, i, j - i))
            i = j + 1
            jj = jj + 1
        Loop Until jj > NCOLUMNS
1984:
    Next ii
'-------------------------------------------------------------------------------------------------
Return
'-------------------------------------------------------------------------------------------------
ERROR_LABEL:
OANDA_HISTORICAL_DATA_FUNC = Err.number
End Function

Function OANDA_PARSE_DATE_VAL(ByVal DATE_VAL As String) As Date

Dim i As Integer
Dim j As Integer
Const DELIM_CHR As String = "/"
Dim DAY_INT As Integer
Dim MONTH_INT As Integer
Dim YEAR_INT As Integer

On Error GoTo ERROR_LABEL

i = 1: j = InStr(i, DATE_VAL, DELIM_CHR)
MONTH_INT = Val(Mid(DATE_VAL, i, j - i))
i = j + 1: j = InStr(i, DATE_VAL, DELIM_CHR)
DAY_INT = Val(Mid(DATE_VAL, i, j - i))
i = j + 1: j = Len(DATE_VAL)
YEAR_INT = Val(Mid(DATE_VAL, i, j - i + 1))

OANDA_PARSE_DATE_VAL = DateSerial(YEAR_INT, MONTH_INT, DAY_INT)

Exit Function
ERROR_LABEL:
OANDA_PARSE_DATE_VAL = Err.number
End Function
