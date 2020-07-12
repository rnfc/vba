Attribute VB_Name = "WEB_SERVICE_MORNING_CHART_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.

Function MORNINGSTAR_FUNDAMENTAL_CHARTS_FUNC( _
ByRef TICKERS_RNG As Variant, _
ByRef NAMES_RNG As Variant, _
Optional ByRef SUFFIX_STR As String = _
"PB,PC,PE,PS,RG,OIG,EPSG,EQG,CFO,EPS,ROEG10,ROAG10,PROA,ROEA,TOTR,CR,DE,DTC", _
Optional ByVal DELIM_CHR As String = ",")

Dim i As Long
Dim j As Long
Dim k As Long
Dim l As Long

Dim ii As Long
Dim jj As Long
Dim kk As Long

Dim NROWS As Long
Dim NSIZE As Long

Dim BODY_STR As String
Dim DST_FILE_NAME As String

Dim WIDTH_VAL As Double
Dim HEIGHT_VAL As Double

Dim TEMP1_VECTOR As Variant
Dim TEMP2_VECTOR As Variant
Dim NAMES_VECTOR As Variant
Dim TICKERS_VECTOR As Variant

On Error GoTo ERROR_LABEL

SUFFIX_STR = Trim(SUFFIX_STR)
MORNINGSTAR_FUNDAMENTAL_CHARTS_FUNC = False

If IsArray(TICKERS_RNG) = True Then
    TICKERS_VECTOR = TICKERS_RNG
    If UBound(TICKERS_VECTOR, 1) = 1 Then: _
        TICKERS_VECTOR = MATRIX_TRANSPOSE_FUNC(TICKERS_VECTOR)
    
    NAMES_VECTOR = NAMES_RNG
    If UBound(NAMES_VECTOR, 1) = 1 Then: _
        NAMES_VECTOR = MATRIX_TRANSPOSE_FUNC(NAMES_VECTOR)
Else
    ReDim TICKERS_VECTOR(1 To 1, 1 To 1)
    TICKERS_VECTOR(1, 1) = TICKERS_RNG
    
    ReDim NAMES_VECTOR(1 To 1, 1 To 1)
    NAMES_VECTOR(1, 1) = NAMES_RNG
End If

If UBound(TICKERS_VECTOR, 1) <> UBound(NAMES_VECTOR, 1) Then: _
GoTo ERROR_LABEL

NROWS = UBound(TICKERS_VECTOR, 1)
NSIZE = 0
i = InStr(1, SUFFIX_STR, DELIM_CHR)
Do While i > 0 'Count Charts
  i = i + 1
  NSIZE = NSIZE + 1
  i = InStr(i, SUFFIX_STR, DELIM_CHR)
Loop
NSIZE = NSIZE + 1


l = NROWS * NSIZE
ReDim TEMP1_VECTOR(1 To l, 1 To 1)
ReDim TEMP2_VECTOR(1 To l, 1 To 1)

kk = 1
For ii = 1 To NROWS
    j = 0
    For jj = 1 To NSIZE
        If jj <> NSIZE Then
            i = j + 1
            j = InStr(i, SUFFIX_STR, DELIM_CHR)
        Else
            i = j + 1
            j = Len(SUFFIX_STR) + 1
        End If
        TEMP1_VECTOR(kk, 1) = TICKERS_VECTOR(ii, 1) & DELIM_CHR & Mid(SUFFIX_STR, i, j - i)
        TEMP2_VECTOR(kk, 1) = NAMES_VECTOR(ii, 1)
        kk = kk + 1
    Next jj
Next ii

TICKERS_VECTOR = TEMP1_VECTOR
NAMES_VECTOR = TEMP2_VECTOR

ReDim TEMP1_VECTOR(1 To l, 1 To 4) 'Reference/Source/Width/Height
HEIGHT_VAL = 360: WIDTH_VAL = 350
For k = 1 To l
    If TICKERS_VECTOR(k, 1) = "" Then: GoTo 1983
'-----------------------------------------------------------------------------------------
'PB,PC,PE,PS,RG,OIG,EPSG,EQG,CFO,EPS,ROEG10,ROAG10,PROA,ROEA,TOTR,CR,DE,DTC
'-----------------------------------------------------------------------------------------
    TEMP2_VECTOR = Split(TICKERS_VECTOR(k, 1), ",", -1, vbBinaryCompare)
    If IsArray(TEMP2_VECTOR) = False Then: GoTo 1983
    i = LBound(TEMP2_VECTOR): j = UBound(TEMP2_VECTOR)
    TEMP1_VECTOR(k, 1) = "http://quote.morningstar.com/Quote/Quote.aspx?pgid=hetopquote&ticker=" & TEMP2_VECTOR(i)
    TEMP1_VECTOR(k, 2) = "http://tools.morningstar.com/charts/MStarCharts.aspx?" & "Security=" & TEMP2_VECTOR(i) & "&bSize=460&Fundamental=" & TEMP2_VECTOR(j) & "&Options=F&Stock=&FPrime=" & TEMP2_VECTOR(i)
    TEMP1_VECTOR(k, 3) = WIDTH_VAL
    TEMP1_VECTOR(k, 4) = HEIGHT_VAL
1983:
Next k
TICKERS_VECTOR = TEMP1_VECTOR

BODY_STR = CREATE_HTML_IMAGES_STR_FUNC(TICKERS_VECTOR, NAMES_VECTOR, NSIZE)
DST_FILE_NAME = WEB_BROWSER_TEMP_DIR_FUNC() & "financials_" & Format(Now, "yymmddhhmmss") & ".html"
i = FreeFile 'Returns an Integer representing the next file number
   'available for use by the Open statement.
Open DST_FILE_NAME For Output As #i 'Output Text --> Clean the entire file
Print #i, BODY_STR;
Close #i

Call OPEN_WEB_BROWSER_FUNC(DST_FILE_NAME)

MORNINGSTAR_FUNDAMENTAL_CHARTS_FUNC = True 'return result

Exit Function
ERROR_LABEL:
On Error Resume Next
Close #i
MORNINGSTAR_FUNDAMENTAL_CHARTS_FUNC = False
End Function