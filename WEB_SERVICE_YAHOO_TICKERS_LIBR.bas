Attribute VB_Name = "WEB_SERVICE_YAHOO_TICKERS_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.

'User defined function to convert a Yahoo ticker symbol into another data provider's ticker symbol
'2013.10.07

Function CONVERT_YAHOO_TICKER_FUNC(ByVal TICKER0_STR As String, _
Optional ByVal VERSION As Variant = 0)
                         
Dim TICKER1_STR As String
TICKER1_STR = Trim(UCase(TICKER0_STR))

On Error GoTo ERROR_LABEL

Const YAHOO_NDX_STR = "~^DJI  "
Const GOOGLE_NDX_STR = "~.DJI  "
Const MSN_NDX_STR = "~$INDU "
                            
Select Case True
'   Case VERSION = 0 Or Left(UCase(VERSION), 5) = "ADVFN": GoTo ADVFN_LINE
   Case VERSION = 1 Or Left(UCase(VERSION), 8) = "BARCHART": GoTo BARCHART_LINE
   Case VERSION = 2 Or Left(UCase(VERSION), 8) = "EARNINGS": GoTo EARNINGS_LINE
   Case VERSION = 3 Or Left(UCase(VERSION), 3) = "MSN": GoTo MSN_LINE
   Case VERSION = 4 Or Left(UCase(VERSION), 6) = "GOOGLE": GoTo GOOGLE_LINE
   Case VERSION = 5 Or Left(UCase(VERSION), 11) = "MORNINGSTAR": GoTo MORNINGSTAR_LINE
   Case VERSION = 6 Or Left(UCase(VERSION), 7) = "REUTERS": GoTo REUTERS_LINE
   Case VERSION = 7 Or Left(UCase(VERSION), 11) = "STOCKCHARTS": GoTo STOCKCHARTS_LINE
   Case VERSION = 8 Or Left(UCase(VERSION), 10) = "STOCKHOUSE": GoTo STOCKHOUSE_LINE
   Case VERSION = 9 Or Left(UCase(VERSION), 5) = "ZACKS": GoTo ZACKS_LINE
   End Select
GoTo 1983

'ADVFN_LINE:
'Select Case True
'   Case Right(TICKER1_STR, 3) = ".OB": TICKER1_STR = "USBB:" & Replace(TICKER1_STR, ".OB", "")
'   Case Right(TICKER1_STR, 3) = ".TO": TICKER1_STR = "TSX:" & Replace(TICKER1_STR, ".TO", "")
'   Case Right(TICKER1_STR, 2) = ".V": TICKER1_STR = "TSX:" & Replace(TICKER1_STR, ".V", "")
'End Select
'If InStr(1, TICKER1_STR, "-") > 0 Then: TICKER1_STR = "." & Replace(TICKER1_STR, "-", ".")
'GoTo 1983
   
BARCHART_LINE:
Select Case True
   Case Right(TICKER1_STR, 3) = ".OB": TICKER1_STR = Replace(TICKER1_STR, ".OB", "")
   'Case Right(TICKER1_STR, 3) = ".TO"
   End Select
GoTo 1983
   
EARNINGS_LINE:
Select Case True
   Case Right(TICKER1_STR, 3) = ".OB": TICKER1_STR = Replace(TICKER1_STR, ".OB", "")
   'Case Right(TICKER1_STR, 3) = ".TO"
   End Select
GoTo 1983
  
GOOGLE_LINE:
Select Case True
   Case Right(TICKER1_STR, 3) = ".OB": TICKER1_STR = "OTC:" & Replace(TICKER1_STR, ".OB", "")
   Case Right(TICKER1_STR, 3) = ".TO": TICKER1_STR = "TSE:" & Replace(TICKER1_STR, ".TO", "")
   Case Right(TICKER1_STR, 2) = ".V": TICKER1_STR = "TSE:" & Replace(TICKER1_STR, ".V", "")
   Case InStr(YAHOO_NDX_STR, "~" & TICKER1_STR) > 0: TICKER1_STR = Trim(Mid(GOOGLE_NDX_STR, InStr(YAHOO_NDX_STR, "~" & TICKER1_STR) + 1, 5))
   End Select
GoTo 1983

MORNINGSTAR_LINE:
Select Case True
   Case Right(TICKER1_STR, 3) = ".OB": TICKER1_STR = Replace(TICKER1_STR, ".OB", "")
   Case Right(TICKER1_STR, 3) = ".TO": TICKER1_STR = "XTSE:" & Replace(TICKER1_STR, ".TO", "")
   End Select
GoTo 1983
   
MSN_LINE:
Select Case True
   Case Right(TICKER1_STR, 3) = ".OB": TICKER1_STR = Replace(TICKER1_STR, ".OB", "")
   Case Right(TICKER1_STR, 3) = ".TO": TICKER1_STR = "CA:" & Replace(TICKER1_STR, ".TO", "")
   Case Right(TICKER1_STR, 2) = ".V": TICKER1_STR = "CA:" & Replace(TICKER1_STR, ".V", "")
   Case Right(TICKER1_STR, 2) = ".X": TICKER1_STR = "." & Replace(TICKER1_STR, ".X", "")
   Case InStr(TICKER1_STR, "-P") > 0: TICKER1_STR = Replace(TICKER1_STR, "-P", "-")
   Case InStr(TICKER1_STR, "-") > 0: TICKER1_STR = Replace(TICKER1_STR, "-", "/")
   Case InStr(YAHOO_NDX_STR, "~" & TICKER1_STR) > 0: TICKER1_STR = Trim(Mid(MSN_NDX_STR, InStr(YAHOO_NDX_STR, "~" & TICKER1_STR) + 1, 5))
   End Select
GoTo 1983

REUTERS_LINE:
Select Case True
   'Case Right(TICKER1_STR, 3) = ".OB"
   'Case Right(TICKER1_STR, 3) = ".TO"
   End Select
GoTo 1983
   
STOCKCHARTS_LINE:
Select Case True
   Case Right(TICKER1_STR, 3) = ".OB": TICKER1_STR = Replace(TICKER1_STR, ".OB", "")
   'Case Right(TICKER1_STR, 3) = ".TO"
   End Select
GoTo 1983

STOCKHOUSE_LINE:
Select Case True
   Case Right(TICKER1_STR, 3) = ".OB": TICKER1_STR = Replace(TICKER1_STR, ".OB", "")
   Case Right(TICKER1_STR, 3) = ".TO": TICKER1_STR = "T." & Replace(TICKER1_STR, ".TO", "")
   End Select
GoTo 1983

ZACKS_LINE:
Select Case True
   Case Right(TICKER1_STR, 3) = ".OB": TICKER1_STR = Replace(TICKER1_STR, ".OB", "")
   Case Right(TICKER1_STR, 3) = ".TO": TICKER1_STR = "T." & Replace(TICKER1_STR, ".TO", "")
   End Select
GoTo 1983
               
1983:
CONVERT_YAHOO_TICKER_FUNC = TICKER1_STR

Exit Function
ERROR_LABEL:
CONVERT_YAHOO_TICKER_FUNC = TICKER1_STR
End Function


Function CONVERT_MSN_YAHOO_TICKER_FUNC(ByVal SYMBOL_STR As String)

Dim OUT_STR As String

On Error GoTo ERROR_LABEL

If InStr(1, SYMBOL_STR, "-", vbTextCompare) Then
    SYMBOL_STR = Replace(SYMBOL_STR, "-", "-P")
    OUT_STR = SYMBOL_STR
Else
    OUT_STR = SYMBOL_STR
End If

If InStr(1, SYMBOL_STR, "/", vbTextCompare) Then
    SYMBOL_STR = Replace(SYMBOL_STR, "/", "-")
    OUT_STR = SYMBOL_STR
Else
    OUT_STR = SYMBOL_STR
End If

CONVERT_MSN_YAHOO_TICKER_FUNC = OUT_STR

Exit Function
ERROR_LABEL:
CONVERT_MSN_YAHOO_TICKER_FUNC = Err.number
End Function

Function CONVERT_YAHOO_MSN_TICKER_FUNC(ByVal SYMBOL_STR As String)

Dim OUT_STR As String

On Error GoTo ERROR_LABEL

If InStr(1, SYMBOL_STR, "-P", vbTextCompare) Or _
   InStr(1, SYMBOL_STR, "-p", vbTextCompare) Then
    SYMBOL_STR = Replace(SYMBOL_STR, "-P", "-")
    SYMBOL_STR = Replace(SYMBOL_STR, "-p", "-")
    OUT_STR = SYMBOL_STR
    Exit Function
Else
    OUT_STR = SYMBOL_STR
End If

If InStr(1, SYMBOL_STR, "-", vbTextCompare) Then
    SYMBOL_STR = Replace(SYMBOL_STR, "-", "/")
    OUT_STR = SYMBOL_STR
Else
    OUT_STR = SYMBOL_STR
End If

CONVERT_YAHOO_MSN_TICKER_FUNC = OUT_STR

Exit Function
ERROR_LABEL:
CONVERT_YAHOO_MSN_TICKER_FUNC = Err.number
End Function
