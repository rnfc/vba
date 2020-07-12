Attribute VB_Name = "WEB_SERVICE_CHARTS_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.
                    
'Function to create a char object and insert image/text into the Range
'Creates a comment box for the cell that can contain text and/or an image (e.g.
'a stock chart).

Function RNG_FINANCIAL_CHART_COMMENT_FUNC( _
Optional ByVal TICKER_STR As String = "", _
Optional ByVal VERSION As Integer = 1, _
Optional ByVal WIDTH_VAL As Integer = 0, _
Optional ByVal HEIGHT_VAL As Integer = 0, _
Optional ByVal VISIBLE_FLAG As Integer = 0, _
Optional ByVal TOP_VAL As Integer = 1, _
Optional ByVal LEFT_VAL As Integer = 1, _
Optional ByVal SCALE_VAL As Single = 1#, _
Optional ByVal TEXT_STR As String = "", _
Optional ByVal RETURN_STR As String = "Chart")

Dim i As Long
Dim NSIZE As Long
Dim MIN_VAL As Variant
Dim MAX_VAL As Variant
Dim DATA_VAL As Variant
Dim DATA_ARR As Variant
Dim SRC_URL_STR As String
Dim SRC_RNG As Excel.Range

'Width = Width (in pixels) of comment box. Optional except when Choice=99. Default
'value s vary depending on the value of "Choice".

'Height = Height (in pixels) of comment box. Optional except when Choice=99.
'Default values vary depending on the value of "Choice".

'Visible = A binary value to indicate whether the comment should be visible
'or hidden; opt ional; defaults to 0 (hidden).
'0 = Keep comment hidden
'1 = Make comment visible

'Top = Top position (in pixels) of comment relative to the top of the cell
'containing the formula. Defaults to 1.

'Left = Leftmost position (in pixels) of comment relative to the left of the
'cell containing the formula. Defaults to 1.

'Scale = Scaling amount to apply to Choice-based height and width sizes. Defaults
'to 100 %.

'Text = Text data to place into comment box. In most cases, you would NOT use
'this if you are loading an image. Defaults to "".

Set SRC_RNG = Cells(Application.Caller.Cells.row, Application.Caller.Cells.Column)
If ActiveWorkbook.name <> Application.Caller.Parent.Parent.name Then Exit Function
If ActiveSheet.name <> Application.Caller.Worksheet.name Then Exit Function

On Error Resume Next
SRC_RNG.Comment.Delete
On Error GoTo 0

Select Case True
Case UCase(TICKER_STR) = "NONE": RNG_FINANCIAL_CHART_COMMENT_FUNC = "None": Exit Function
Case VERSION = 0
    SRC_URL_STR = ""
    If WIDTH_VAL = 0 Then WIDTH_VAL = 300
    If HEIGHT_VAL = 0 Then HEIGHT_VAL = 200
Case TICKER_STR = "": GoTo ERROR_LABEL
Case SCALE_VAL <= 0: GoTo ERROR_LABEL
Case WIDTH_VAL < 0: GoTo ERROR_LABEL
Case HEIGHT_VAL < 0: GoTo ERROR_LABEL
Case VERSION = 1      ' Daily Chart of Gallery View from StockCharts
SRC_URL_STR = "http://stockcharts.com/c-sc/sc?chart=" & TICKER_STR & ",uu[h,a]daclyyay[pb50!b200!f][vc60][iue12,26,9!lc20]"
    If WIDTH_VAL = 0 Then WIDTH_VAL = 350 * SCALE_VAL
    If HEIGHT_VAL = 0 Then HEIGHT_VAL = 390 * SCALE_VAL
Case VERSION = 2      ' P&F Chart from StockCharts
    SRC_URL_STR = "http://stockcharts.com/def/servlet/SharpChartv05.ServletDriver?chart=" & TICKER_STR & ",pltad[pa][da][f!3!!]&pnf=y"
    If WIDTH_VAL = 0 Then WIDTH_VAL = 390 * SCALE_VAL
    If HEIGHT_VAL = 0 Then HEIGHT_VAL = 314 * SCALE_VAL
Case VERSION = 3      ' 6-month Candleglance Chart from StockCharts
    SRC_URL_STR = "http://stockcharts.com/c-sc/sc?chart=" & TICKER_STR & ",uu[305,a]dacayaci[pb20!b50][dc]"
    If WIDTH_VAL = 0 Then WIDTH_VAL = 229 * SCALE_VAL
    If HEIGHT_VAL = 0 Then HEIGHT_VAL = 132 * SCALE_VAL
Case VERSION = 4      ' 6-month chart from Business Week Online
    SRC_URL_STR = "http://stockcharts.com/c-sc/sc?s=" & TICKER_STR & "&p=D&yr=0&mn=6&dy=0&i=t94339682869&r=4806"
    If WIDTH_VAL = 0 Then WIDTH_VAL = 638 * SCALE_VAL
    If HEIGHT_VAL = 0 Then HEIGHT_VAL = 501 * SCALE_VAL
Case VERSION = 5      ' 6-month chart Rule #1 Technicals from StockCharts
    SRC_URL_STR = "http://stockcharts.com/c-sc/sc?s=" & TICKER_STR & "&p=D&yr=0&mn=6&dy=0&i=t39628903145&r=9933"
    If WIDTH_VAL = 0 Then WIDTH_VAL = 350 * SCALE_VAL
    If HEIGHT_VAL = 0 Then HEIGHT_VAL = 360 * SCALE_VAL
Case VERSION = 97
    SRC_URL_STR = "http://www.advfn.com/p.php?pid=financialgraphs" '&a0=13&a1=13&a2=10&a3=8&a4=8&a5=10"
    DATA_ARR = Split(TICKER_STR, ",")
    NSIZE = UBound(DATA_ARR, 1)
    If NSIZE <> 4 Then GoTo ERROR_LABEL
    For i = 0 To NSIZE
    SRC_URL_STR = SRC_URL_STR & "&a" & i & "=" & DATA_ARR(i)
    Next i
    If WIDTH_VAL = 0 Then WIDTH_VAL = 263 * SCALE_VAL
    If HEIGHT_VAL = 0 Then HEIGHT_VAL = 169 * SCALE_VAL
Case VERSION = 98
    SRC_URL_STR = "http://ogres-crypt.com/php/chart.php?d="
    DATA_ARR = Split(TICKER_STR, ",")
    NSIZE = UBound(DATA_ARR, 1)
    MAX_VAL = 0.01
    MIN_VAL = 999999999
    For i = 0 To NSIZE
        DATA_VAL = 0
        On Error Resume Next
        DATA_VAL = CDec(DATA_ARR(i))
        On Error GoTo 0
        If (DATA_VAL > MAX_VAL) Then MAX_VAL = DATA_VAL
        If (DATA_VAL < MIN_VAL And DATA_VAL > 0) Then MIN_VAL = DATA_VAL
    Next i
    For i = 0 To NSIZE
        DATA_VAL = 0
        On Error Resume Next
        DATA_VAL = CDec(DATA_ARR(i))
        On Error GoTo 0
        DATA_VAL = IIf(DATA_VAL > 0, 1 + 97 * (DATA_VAL - MIN_VAL) / (MAX_VAL - MIN_VAL), 0)
        SRC_URL_STR = SRC_URL_STR & CInt(DATA_VAL) & IIf(i = NSIZE, "", ",")
    Next i
    If WIDTH_VAL = 0 Then WIDTH_VAL = 36 * NSIZE * SCALE_VAL
    If HEIGHT_VAL = 0 Then HEIGHT_VAL = 90 * SCALE_VAL
Case VERSION = 99
    SRC_URL_STR = TICKER_STR
    If WIDTH_VAL = 0 Then WIDTH_VAL = 400
    If HEIGHT_VAL = 0 Then HEIGHT_VAL = 300
    Case Else: GoTo ERROR_LABEL
End Select

SRC_RNG.AddComment ("")
If SRC_URL_STR <> "" Then SRC_RNG.Comment.Shape.Fill.UserPicture SRC_URL_STR
SRC_RNG.Comment.Text Text:=IIf(TEXT_STR = "", Chr(32), TEXT_STR)
SRC_RNG.Comment.Shape.Width = WIDTH_VAL
SRC_RNG.Comment.Shape.Height = HEIGHT_VAL
SRC_RNG.Comment.Shape.Top = TOP_VAL + SRC_RNG.Top
SRC_RNG.Comment.Shape.Left = LEFT_VAL + SRC_RNG.Left
SRC_RNG.Comment.Shape.Line.Visible = False             ' Doesn't work
SRC_RNG.Comment.Shape.Line.ForeColor.SchemeColor = 9   ' Set line color to background color
SRC_RNG.Comment.Shape.Shadow.Visible = False
SRC_RNG.Comment.Visible = IIf(VISIBLE_FLAG = 1, True, False)
RNG_FINANCIAL_CHART_COMMENT_FUNC = RETURN_STR

Exit Function
ERROR_LABEL:
RNG_FINANCIAL_CHART_COMMENT_FUNC = "Error"
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      :
'DESCRIPTION   :
'LIBRARY       : HTML
'GROUP         :
'ID            : 001
'LAST UPDATE   : 21/02/2008
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'************************************************************************************
'************************************************************************************
                    
Function PRINT_WEB_FINANCIAL_CHARTS_FUNC( _
ByRef TICKERS_RNG As Variant, _
ByRef HEADERS_RNG As Variant, _
Optional ByVal VERSION As Integer = 6, _
Optional ByVal NSIZE As Integer = 1, _
Optional ByVal DST_FILE_NAME As String = "")

Dim i As Integer
Dim BODY_STR As String
Dim TICKERS_VECTOR As Variant
Dim HEADERS_VECTOR As Variant

On Error GoTo ERROR_LABEL

PRINT_WEB_FINANCIAL_CHARTS_FUNC = False

If IsArray(TICKERS_RNG) = True Then
    TICKERS_VECTOR = TICKERS_RNG
    If UBound(TICKERS_VECTOR, 1) = 1 Then
        TICKERS_VECTOR = MATRIX_TRANSPOSE_FUNC(TICKERS_VECTOR)
    End If
Else
    ReDim TICKERS_VECTOR(1 To 1, 1 To 1)
    TICKERS_VECTOR(1, 1) = TICKERS_RNG
End If

If IsArray(HEADERS_RNG) = True Then
    HEADERS_VECTOR = HEADERS_RNG
    If UBound(HEADERS_VECTOR, 1) = 1 Then
        HEADERS_VECTOR = MATRIX_TRANSPOSE_FUNC(HEADERS_VECTOR)
    End If
Else
    ReDim HEADERS_VECTOR(1 To 1, 1 To 1)
    HEADERS_VECTOR(1, 1) = HEADERS_RNG
End If

TICKERS_VECTOR = CREATE_WEB_FINANCIAL_CHARTS_FUNC(TICKERS_VECTOR, 1, VERSION)
BODY_STR = CREATE_HTML_IMAGES_STR_FUNC(TICKERS_VECTOR, HEADERS_VECTOR, NSIZE)

If DST_FILE_NAME = "" Then
    DST_FILE_NAME = "financial_" & Format(Now, "yymmddhhmmss") & ".html"
    If WRITE_TEMP_HTML_TEXT_FILE_FUNC(DST_FILE_NAME, "", BODY_STR) = True Then
        PRINT_WEB_FINANCIAL_CHARTS_FUNC = True
    End If
Else
   i = FreeFile
      Open DST_FILE_NAME For Output As #i
         Print #i, BODY_STR;
   Close #i
   PRINT_WEB_FINANCIAL_CHARTS_FUNC = True
End If

Exit Function
ERROR_LABEL:
PRINT_WEB_FINANCIAL_CHARTS_FUNC = False
End Function

Sub SHOW_FINANCIAL_CHARTS_FORM()
    frmCharts.show
End Sub


Private Function CREATE_WEB_FINANCIAL_CHARTS_FUNC(ByRef TICKERS_RNG As Variant, _
Optional ByRef SCALE_VAL As Double = 1.5, _
Optional ByVal VERSION As Integer = 10)

Dim i As Integer
Dim j As Integer
Dim k As Integer
Dim NROWS As Integer

Dim TICKER_STR As String

Dim REF_URL_STR As String
Dim SRC_URL_STR As String

Dim WIDTH_VAL As Double
Dim HEIGHT_VAL As Double

Dim MIN_VAL As Double
Dim MAX_VAL As Double
Dim TEMP_VAL As Double

Dim TEMP_ARR As Variant
Dim TEMP_MATRIX As Variant
Dim TICKERS_VECTOR As Variant

On Error GoTo ERROR_LABEL

If SCALE_VAL <= 0 Then: SCALE_VAL = 1#

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

ReDim TEMP_MATRIX(1 To NROWS, 1 To 4) 'Reference/Source/Width/Height

For k = 1 To NROWS
    TICKER_STR = TICKERS_VECTOR(k, 1)
    If TICKER_STR = "" Then: GoTo 1983

'-----------------------------------------------------------------------------------------
Select Case VERSION
'-----------------------------------------------------------------------------------------
Case 1 'Yahoo Finance 1-Day Range
'-----------------------------------------------------------------------------------------
     REF_URL_STR = "http://finance.yahoo.com/q/bc?s=" & _
                    TICKER_STR
     
     SRC_URL_STR = "http://ichart.finance.yahoo.com/b?s=" & _
                    TICKER_STR
                    'Yahoo Generic http://ichart.yahoo.com/v?s=
     WIDTH_VAL = 638 * SCALE_VAL
     HEIGHT_VAL = 501 * SCALE_VAL
'-----------------------------------------------------------------------------------------
Case 2 'Yahoo Finance 1-Week Range
'-----------------------------------------------------------------------------------------
     REF_URL_STR = "http://finance.yahoo.com/q/bc?s=" & _
                    TICKER_STR
     
     SRC_URL_STR = "http://ichart.finance.yahoo.com/w?s=" & _
                    TICKER_STR
     WIDTH_VAL = 638 * SCALE_VAL
     HEIGHT_VAL = 501 * SCALE_VAL
'-----------------------------------------------------------------------------------------
Case 3 'Implied Volatility
'-----------------------------------------------------------------------------------------
     REF_URL_STR = "http://www.ivolatility.com/options.j?ticker=" & _
                    TICKER_STR
     SRC_URL_STR = _
     "http://www.ivolatility.com/nchart.j?charts=volatility,options_volume&1=ticker*" _
     & TICKER_STR & CStr(",R*1,period*12,all*4,schema*options_big&2=ticker*") _
     & TICKER_STR & CStr(",R*1,period*12,schema*options_big_narrow&add=x:1")
     
     WIDTH_VAL = 638 * SCALE_VAL
     HEIGHT_VAL = 501 * SCALE_VAL

'-----------------------------------------------------------------------------------------
 Case 4 'Finviz daily technical chart
'-----------------------------------------------------------------------------------------
     
     REF_URL_STR = "http://elite.finviz.com/quote.ashx?t=" & _
                    TICKER_STR & "&ta=1&p=d"
     
     SRC_URL_STR = "http://elite.finviz.com/chart.ashx?t=" & _
                    TICKER_STR & "&ta=1&p=d&s=l"

     WIDTH_VAL = 638 * SCALE_VAL
     HEIGHT_VAL = 501 * SCALE_VAL

'-----------------------------------------------------------------------------------------
 Case 5 'Finviz intraday basic
'-----------------------------------------------------------------------------------------
     
     REF_URL_STR = "http://elite.finviz.com/chart.ashx?t=" & _
                    TICKER_STR & "&ta=0&p=i"
     
     SRC_URL_STR = "http://elite.finviz.com/chart.ashx?t=" & _
                    TICKER_STR & "&ta=0&p=i&s=l"

     WIDTH_VAL = 638 * SCALE_VAL
     HEIGHT_VAL = 501 * SCALE_VAL
'-----------------------------------------------------------------------------------------
 Case 6 'Fred Historical Charts
'-----------------------------------------------------------------------------------------
     REF_URL_STR = FRED_HISTORICAL_DATA_CHART_FUNC(TICKER_STR, 0)
     SRC_URL_STR = FRED_HISTORICAL_DATA_CHART_FUNC(TICKER_STR, 1)

     WIDTH_VAL = 638 * SCALE_VAL
     HEIGHT_VAL = 501 * SCALE_VAL

'-----------------------------------------------------------------------------------------
Case 7 'Draw X/Y Bar Chart using ADVFN
'-----------------------------------------------------------------------------------------
     REF_URL_STR = "http://www.advfn.com/"
     SRC_URL_STR = "http://www.advfn.com/p.php?pid=financialgraphs"
     '&a0=13&a1=13&a2=10&a3=8&a4=8 'a4-last fiscal year
     TEMP_ARR = Split(TICKER_STR, ",", -1, vbBinaryCompare)
     For i = LBound(TEMP_ARR, 1) To UBound(TEMP_ARR, 1)
         SRC_URL_STR = SRC_URL_STR & "&a" & i & "=" & TEMP_ARR(i)
     Next i
     WIDTH_VAL = 263 * SCALE_VAL
     HEIGHT_VAL = 169 * SCALE_VAL
'-----------------------------------------------------------------------------------------
Case Else 'Draw X/Y Bar Chart using Ogres
'-----------------------------------------------------------------------------------------
     REF_URL_STR = "http://ogres-crypt.com/php/"
     SRC_URL_STR = "http://ogres-crypt.com/php/chart.php?d="
     TEMP_ARR = Split(TICKER_STR, ",", -1, vbBinaryCompare)
     MAX_VAL = 0.01
     MIN_VAL = 999999999
     For i = LBound(TEMP_ARR, 1) To UBound(TEMP_ARR, 1)
         TEMP_VAL = 0
         On Error Resume Next
         TEMP_VAL = CDec(TEMP_ARR(i))
         On Error GoTo ERROR_LABEL
         If (TEMP_VAL > MAX_VAL) Then MAX_VAL = TEMP_VAL
         If (TEMP_VAL < MIN_VAL And TEMP_VAL > 0) Then MIN_VAL = TEMP_VAL
     Next i
     For i = LBound(TEMP_ARR, 1) To UBound(TEMP_ARR, 1)
         TEMP_VAL = 0
         On Error Resume Next
         TEMP_VAL = CDec(TEMP_ARR(i))
         On Error GoTo ERROR_LABEL
         TEMP_VAL = IIf(TEMP_VAL > 0, 1 + 97 * (TEMP_VAL - MIN_VAL) / (MAX_VAL - MIN_VAL), 0)
         SRC_URL_STR = SRC_URL_STR & CInt(TEMP_VAL) & IIf(i = j, "", ",")
     Next i
     If WIDTH_VAL = 0 Then WIDTH_VAL = 36 * j * SCALE_VAL
     If HEIGHT_VAL = 0 Then HEIGHT_VAL = 90 * SCALE_VAL
     
'-----------------------------------------------------------------------------------------
End Select
'-----------------------------------------------------------------------------------------
TEMP_MATRIX(k, 1) = Trim(REF_URL_STR)
TEMP_MATRIX(k, 2) = Trim(SRC_URL_STR)
TEMP_MATRIX(k, 3) = WIDTH_VAL
TEMP_MATRIX(k, 4) = HEIGHT_VAL
1983:
Next k
'-----------------------------------------------------------------------------------------

CREATE_WEB_FINANCIAL_CHARTS_FUNC = TEMP_MATRIX

Exit Function
ERROR_LABEL:
CREATE_WEB_FINANCIAL_CHARTS_FUNC = Err.number
End Function


Function STOCKCHARTS_URL_FUNC(ByVal TICKER_STR As String, _
Optional ByVal CHART_TYPE As String = "D", _
Optional ByVal nYEARS As Integer = 5, _
Optional ByVal nMONTHS As Integer = 10, _
Optional ByVal nDAYS As Integer = 26)

Dim TEMP_STR As String
Dim SRC_URL_STR As String

On Error GoTo ERROR_LABEL

TEMP_STR = CStr("&p=") & CHART_TYPE & CStr("&yr=") & nYEARS & CStr("&mn=") _
           & nMONTHS & CStr("&dy=") & nDAYS & _
           "&id=p20409734855"

SRC_URL_STR = "http://stockcharts.com/h-sc/ui?s="
SRC_URL_STR = SRC_URL_STR & TICKER_STR
SRC_URL_STR = SRC_URL_STR & TEMP_STR


STOCKCHARTS_URL_FUNC = SRC_URL_STR
Exit Function
ERROR_LABEL:
STOCKCHARTS_URL_FUNC = Err.number
End Function


Function OPTIONISTICS_CHART_URL(ByVal TICKER_STR As String, _
Optional ByVal PERIODS As Integer = 22, _
Optional ByVal EXPIRATION As String = "2007-11-17", _
Optional ByVal OUTPUT As Integer = 0)

'1 Week --> PERIODS = 5
'1 Month --> PERIODS = 22
'2 Months --> PERIODS = 45
'3 Months --> PERIODS = 66

On Error GoTo ERROR_LABEL

Select Case OUTPUT
Case 0
    OPTIONISTICS_CHART_URL = _
        "http://www.optionistics.com/f/inset.pl?vol=0&stk=1&isopt=0&symbol=" & _
        TICKER_STR & CStr("&pc=0&numdays=") & PERIODS
Case 1
    OPTIONISTICS_CHART_URL = _
        "http://www.optionistics.com/f/inset.pl?symbol=" & _
        TICKER_STR & CStr("&expiry=") & (EXPIRATION) & "&v=daily"
Case 2
    OPTIONISTICS_CHART_URL = _
        "http://www.optionistics.com/f/inset.pl?vol=1&stk=1&isopt=0&symbol=" & _
        TICKER_STR & CStr("&pc=0&numdays=") & PERIODS
Case Else
    OPTIONISTICS_CHART_URL = _
        "http://www.optionistics.com/f/inset.pl?vol=1&symbol=" & _
        TICKER_STR & CStr("&expiry=") & (EXPIRATION) & "&v=daily"
End Select

Exit Function
ERROR_LABEL:
OPTIONISTICS_CHART_URL = Err.number
End Function

Function iVOLATILITY_CHART_FUNC( _
Optional ByVal TICKER_STR As String = "MSFT:NASDAQ", _
Optional ByVal months As Integer = 3, _
Optional ByVal VERSION As Integer = 1, _
Optional ByVal OUTPUT As Integer = 1)

Dim SRC_URL_STR As String

On Error GoTo ERROR_LABEL

SRC_URL_STR = "http://www.ivolatility.com"

Select Case OUTPUT
Case 0
    SRC_URL_STR = SRC_URL_STR & _
    "/nchart.j?charts=volatility,options_volume&amp;1=ticker*" & TICKER_STR & _
    ",R*1,period*" & months & ",all*" & VERSION & _
    ",schema*options_big&amp;2=ticker*" & TICKER_STR & _
    ",R*1,period*" & months & ",schema*options_big_narrow&amp;add=x:1"
Case Else
    SRC_URL_STR = SRC_URL_STR & _
    "/nchart.j?charts=options_volume,volatility&amp;1=ticker*" & TICKER_STR & _
    ",R*1,period*" & months & ",schema*options_big_narrow&amp;2=ticker*" & _
    TICKER_STR & ",R*1,period*" & months & ",all*" & VERSION & _
    ",schema*options_big&amp;add=x:1"
End Select

iVOLATILITY_CHART_FUNC = SRC_URL_STR

Exit Function
ERROR_LABEL:
iVOLATILITY_CHART_FUNC = Err.number
End Function
