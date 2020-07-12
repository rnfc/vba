Attribute VB_Name = "WEB_SERVICE_YAHOO_BONDS_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.

Function YAHOO_US_BONDS_SNAPSHOT_FUNC(Optional ByVal VERSION As Integer = 0)

'http://www.treas.gov/tic/mfh.txt

Dim h, i, j, k, l, m As Long
Dim TEMP_STR As String
Dim DATA_STR As String
Dim REFER_STR As String

Dim SRC_URL_STR As String

Dim TEMP_GROUP() As String
Dim TEMP_MATRIX() As Variant

On Error GoTo ERROR_LABEL

Select Case VERSION
Case 0
    REFER_STR = "US Treasury Bonds"
    h = 8
Case 1
    REFER_STR = "Municipal Bonds"
    h = 13
Case Else
    REFER_STR = "Corporate Bonds"
    h = 12
End Select

SRC_URL_STR = "http://finance.yahoo.com/bonds/composite_bond_rates"
DATA_STR = SAVE_WEB_DATA_PAGE_FUNC(SRC_URL_STR, 0, False, 0, False)

i = InStr(1, DATA_STR, REFER_STR)
If i = 0 Then: GoTo ERROR_LABEL

REFER_STR = "</table></div>"
j = InStr(i, DATA_STR, REFER_STR)
If j = 0 Then: GoTo ERROR_LABEL

DATA_STR = Mid(DATA_STR, i, j - i + Len(REFER_STR))

ReDim TEMP_GROUP(1 To 1)
j = 1
k = 0
Do
    i = InStr(j, DATA_STR, "yfifrstchld")
    If i = 0 Then: GoTo 1982
    i = i + Len("yfifrstchld") + 2
    j = InStr(i, DATA_STR, "<")
    If j = 0 Then: GoTo 1982
    TEMP_STR = Mid(DATA_STR, i, j - i) & ","
    
    For l = 1 To 4
        i = InStr(j, DATA_STR, IIf(k <> 0, "<td", "<th"))
        i = InStr(i, DATA_STR, ">") + 1
        j = InStr(i, DATA_STR, "<")
        If j = 0 Then: GoTo 1982
        TEMP_STR = TEMP_STR & Mid(DATA_STR, i, j - i) & _
        IIf(l <> 4, ",", "")
    Next l

    k = k + 1
    ReDim Preserve TEMP_GROUP(1 To k)
    TEMP_GROUP(k) = TEMP_STR
Loop Until k = h
1982:

ReDim TEMP_MATRIX(1 To UBound(TEMP_GROUP), 1 To 5)

For k = 1 To UBound(TEMP_GROUP)
    TEMP_STR = TEMP_GROUP(k)
    If TEMP_STR = "" Then: GoTo 1983
    i = 1: j = InStr(i, TEMP_STR, ",")
    TEMP_MATRIX(k, 1) = Mid(TEMP_STR, i, j - i)
    For m = 1 To 3
        i = j + 1
        j = InStr(i, TEMP_STR, ",")
        TEMP_MATRIX(k, 1 + m) = Mid(TEMP_STR, i, j - i)
    Next m
    i = j + 1: j = Len(TEMP_STR) + 1
    TEMP_MATRIX(k, 5) = Mid(TEMP_STR, i, j - i)
1983:
Next k

YAHOO_US_BONDS_SNAPSHOT_FUNC = CONVERT_STRING_NUMBER_FUNC(TEMP_MATRIX)

Exit Function
ERROR_LABEL:
YAHOO_US_BONDS_SNAPSHOT_FUNC = Err.number
End Function
