Attribute VB_Name = "WEB_SERVICE_YAHOO_NEWS_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.

'************************************************************************************
'************************************************************************************
'FUNCTION      : YAHOO_NEWS_FUNC
'DESCRIPTION   : Yahoo News Wrapper
'LIBRARY       : HTML
'GROUP         : eNews
'ID            : 001
'LAST UPDATE   : 29/02/2008
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'REFERENCE:
'http://www.wilmott.com/messageview.cfm?catid=38&threadid=73237
'************************************************************************************
'************************************************************************************

Sub YAHOO_NEWS_FUNC()

Dim h As Long
Dim i As Long
Dim j As Long
Dim k As Long

Dim DATE_STR As String
Dim DATA_STR As String

Dim ATEMP_STR As String
Dim BTEMP_STR As String
Dim CTEMP_STR As String
Dim DTEMP_STR As String

Dim TICKER_STR As Variant

Dim TEMP_GROUP() As String
Dim SRC_URL_STR As String
Dim TEMP_MATRIX() As String
Dim START_DATE As Date
On Error GoTo ERROR_LABEL

TICKER_STR = Excel.Application.InputBox( _
    prompt:="Stock Symbol", Type:=1 + 2)
If TICKER_STR = False Then: Exit Sub

START_DATE = 0

If START_DATE = 0 Then 'Today
    DATE_STR = ""
Else
    DATE_STR = "&t=" & Year(START_DATE) & "-" & _
                Month(START_DATE) & "-" & Day(START_DATE)
End If

SRC_URL_STR = _
"http://finance.yahoo.com/q/h?s=" & TICKER_STR & DATE_STR
DATA_STR = SAVE_WEB_DATA_PAGE_FUNC(SRC_URL_STR, 0, False, 0, False)

ReDim TEMP_GROUP(1 To 1)

k = 1
j = InStr(1, DATA_STR, "#149;")
If j = 0 Then: GoTo ERROR_LABEL

Do
'----------------------------------------------------------------------------
    i = InStr(j, DATA_STR, "<a href=") + Len("<a href=") 'Start URL
    If i = 0 Then: GoTo ERROR_LABEL
    
    j = InStr(i, DATA_STR, ">") 'End URL
    DTEMP_STR = Replace(Mid(DATA_STR, i, j - i), """", "")
'----------------------------------------------------------------------------
    i = InStr(j, DATA_STR, ">") + 1 'Start Headline
    j = InStr(i, DATA_STR, "<") 'End Headline
    BTEMP_STR = Mid(DATA_STR, i, j - i)
'----------------------------------------------------------------------------
    i = InStr(j, DATA_STR, "<b>") + 3 'Start Source
    j = InStr(i, DATA_STR, "</b>") 'End Source
    CTEMP_STR = Mid(DATA_STR, i, j - i)
'----------------------------------------------------------------------------
    i = InStr(j, DATA_STR, "&nbsp;(") + 7 'Start Date
    j = InStr(i, DATA_STR, ")") 'End Source
    ATEMP_STR = Mid(DATA_STR, i, j - i)
'----------------------------------------------------------------------------
    ReDim Preserve TEMP_GROUP(1 To k)
    TEMP_GROUP(k) = ATEMP_STR & "|" & BTEMP_STR & "|" & CTEMP_STR & "|" & _
                    DTEMP_STR & "|"
'----------------------------------------------------------------------------
    k = k + 1
    h = InStr(j, DATA_STR, "</small>")
'----------------------------------------------------------------------------
Loop Until Mid(DATA_STR, h, Len("</small></td></tr></table>")) = _
                                "</small></td></tr></table>"
'----------------------------------------------------------------------------

ReDim TEMP_MATRIX(0 To UBound(TEMP_GROUP), 1 To 4)
TEMP_MATRIX(0, 1) = "DATE"
TEMP_MATRIX(0, 2) = "HEADLINE"
TEMP_MATRIX(0, 3) = "SOURCE"
TEMP_MATRIX(0, 4) = "URL"

For k = 1 To UBound(TEMP_GROUP)
    ATEMP_STR = TEMP_GROUP(k)
    If ATEMP_STR = "" Then: GoTo 1983
    i = 1
    j = InStr(i, ATEMP_STR, "|")
    TEMP_MATRIX(k, 1) = Trim(Mid(ATEMP_STR, i, j - i))
    
    i = j + 1
    j = InStr(i, ATEMP_STR, "|")
    TEMP_MATRIX(k, 2) = Trim(Mid(ATEMP_STR, i, j - i))
    
    i = j + 1
    j = InStr(i, ATEMP_STR, "|")
    TEMP_MATRIX(k, 3) = Trim(Mid(ATEMP_STR, i, j - i))

    i = j + 1
    j = InStr(i, ATEMP_STR, "|")
    TEMP_MATRIX(k, 4) = Trim(Mid(ATEMP_STR, i, j - i))

1983:
Next k

Call SHOW_RESIZER_FORM_FUNC(TEMP_MATRIX, _
     "Breaking News: " & SRC_URL_STR, _
     "Yahoo Finance", True)


'YAHOO_NEWS_FUNC = TEMP_MATRIX

Exit Sub
ERROR_LABEL:
'YAHOO_NEWS_FUNC = Err.NUMBER
End Sub


' Shows DATA_RNG on the userform frmShowTable, with a
' maximum width and height.

Function SHOW_RESIZER_FORM_FUNC(DATA_RNG As Variant, _
Optional ByVal TABLE_TITLE_STR_NAME As String = "", _
Optional ByVal CAPTION_STR_NAME As String = "Show Table On Form", _
Optional ByVal AUTO_COL_WIDTHS_FLAG As Boolean = True)

    Dim FORM_OBJ As New frmShowTable
    
    On Error GoTo ERROR_LABEL
    
    SHOW_RESIZER_FORM_FUNC = False
    With FORM_OBJ
        .Table = DATA_RNG
        .Title = TABLE_TITLE_STR_NAME
        .Caption = CAPTION_STR_NAME
        .AutoColWidths = AUTO_COL_WIDTHS_FLAG
        .Initialise
        .show
    End With
1983:
    On Error GoTo 0
    SHOW_RESIZER_FORM_FUNC = True
Exit Function
ERROR_LABEL:
SHOW_RESIZER_FORM_FUNC = False
End Function


