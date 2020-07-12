Attribute VB_Name = "WEB_SERVICE_YAHOO_WEB_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.

Private PUB_YAHOO_WEB_PAGE_LOGIN As Boolean

Function YAHOO_WEB_PAGE_LOGIN_FUNC()

On Error GoTo ERROR_LABEL

If PUB_YAHOO_WEB_PAGE_LOGIN = False Then
    Call XML_WEB_PAGE_POST_FUNC("https://login.yahoo.com/", "login=XXXX&passwd=YYYY")
    PUB_YAHOO_WEB_PAGE_LOGIN = True
End If

YAHOO_WEB_PAGE_LOGIN_FUNC = PUB_YAHOO_WEB_PAGE_LOGIN

Exit Function
ERROR_LABEL:
PUB_YAHOO_WEB_PAGE_LOGIN = False
YAHOO_WEB_PAGE_LOGIN_FUNC = PUB_YAHOO_WEB_PAGE_LOGIN
End Function

Function PARSE_YAHOO_ASSET_WEB_PAGE_FUNC(ByRef TICKER_STR As String)

Dim i As Long
Dim j As Long
Dim URL_STR As String
Dim DATA_STR As String

On Error GoTo ERROR_LABEL

URL_STR = "http://finance.yahoo.com/q/pr?s=" & TICKER_STR
DATA_STR = SAVE_WEB_DATA_PAGE_FUNC(URL_STR, 0, False, 0, False)

i = InStr(1, DATA_STR, "Website:")
If i = 0 Then: GoTo ERROR_LABEL
i = InStr(i, DATA_STR, ">"): i = i + 1
j = InStr(i, DATA_STR, "<")

PARSE_YAHOO_ASSET_WEB_PAGE_FUNC = Mid(DATA_STR, i, j - i)

Exit Function
ERROR_LABEL:
PARSE_YAHOO_ASSET_WEB_PAGE_FUNC = Err.number
End Function
