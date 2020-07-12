Attribute VB_Name = "WEB_SERVICE_YAHOO_INDUSTRY_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.


Private PUB_YAHOO_INDUSTRY_ELEMENTS As Variant

'************************************************************************************
'************************************************************************************
'FUNCTION      : YAHOO_INDUSTRY_QUOTES_FUNC
'DESCRIPTION   : Yahoo Industries Key Statistics Wrapper
'LIBRARY       : YAHOO
'GROUP         : INDUSTRY
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'************************************************************************************
'************************************************************************************

Function YAHOO_INDUSTRY_QUOTES_FUNC(ByVal INDUSTRY_NAME As Variant, _
Optional ByVal CALLER_REFRESH As Variant = 0)

Dim h As Integer
Dim i As Integer
Dim j As Integer
Dim k As Integer

Dim NSIZE As Integer
Dim NROWS As Integer
Dim NCOLUMNS As Integer

Dim DATA_STR As Variant
Dim DATA_ARR() As String
Dim SRC_URL_STR As String
Dim TEMP_MATRIX As Variant

On Error GoTo ERROR_LABEL

If IsArray(PUB_YAHOO_INDUSTRY_ELEMENTS) = False Then
    PUB_YAHOO_INDUSTRY_ELEMENTS = YAHOO_INDUSTRY_QUOTES_ELEMENTS_FUNC()
End If
NSIZE = UBound(PUB_YAHOO_INDUSTRY_ELEMENTS, 1)
        
Select Case INDUSTRY_NAME
Case 0 'U.S. Sectors Summary
    SRC_URL_STR = "http://biz.yahoo.com/p/csv/s_conameu.csv"
Case 1 'U.S. Industries Summary
    SRC_URL_STR = "http://biz.yahoo.com/p/csv/sum_conameu.csv"
Case Else 'Sectors
    k = 1
    Do While k <= NSIZE
        If INDUSTRY_NAME = PUB_YAHOO_INDUSTRY_ELEMENTS(k, 2) Then
            j = PUB_YAHOO_INDUSTRY_ELEMENTS(k, 1)
            Exit Do
        End If
        k = k + 1
    Loop
    If k > NSIZE Then: GoTo ERROR_LABEL
    SRC_URL_STR = "http://biz.yahoo.com/p/csv/" & CStr(j) & "conameu.csv"
End Select
        
'--------------------------------------------------------------------------------------
DATA_STR = SAVE_WEB_DATA_PAGE_FUNC(SRC_URL_STR, 0, False, 0, False)
DATA_STR = Replace(DATA_STR, Chr(13), "", 1, -1, vbBinaryCompare)
DATA_STR = Replace(DATA_STR, """", "", 1, -1, vbBinaryCompare)
'--------------------------------------------------------------------------------------
DATA_ARR = Split(DATA_STR, Chr(10), -1, vbBinaryCompare)
        
NROWS = UBound(DATA_ARR) - LBound(DATA_ARR)
NCOLUMNS = 10
ReDim TEMP_MATRIX(1 To NROWS, 1 To NCOLUMNS)
        
i = 1
For h = LBound(DATA_ARR) To UBound(DATA_ARR) - 1
    DATA_STR = Split(DATA_ARR(h), ",", -1, vbBinaryCompare)
    If IsArray(DATA_STR) = False Then: GoTo 1984
            
    j = NCOLUMNS
    For k = UBound(DATA_STR) To LBound(DATA_STR) Step -1
        TEMP_MATRIX(i, j) = Trim(DATA_STR(k))
        j = j - 1
        If j = 0 Then: Exit For
1983:
    Next k
    TEMP_MATRIX(i, 1) = DATA_STR(LBound(DATA_STR))
1984:
    i = i + 1
Next h
'--------------------------------------------------------------------------------------
YAHOO_INDUSTRY_QUOTES_FUNC = CONVERT_STRING_NUMBER_FUNC(TEMP_MATRIX)
'--------------------------------------------------------------------------------------

Exit Function
ERROR_LABEL:
YAHOO_INDUSTRY_QUOTES_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : YAHOO_INDUSTRY_QUOTES_ELEMENTS_FUNC
'DESCRIPTION   : Yahoo Industries Fields Wrapper
'LIBRARY       : YAHOO
'GROUP         : INDUSTRY
'ID            : 002
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 21/02/2008
'************************************************************************************
'************************************************************************************
  
Function YAHOO_INDUSTRY_QUOTES_ELEMENTS_FUNC()

Dim i As Long
Dim j As Long
Dim k As Long

Dim TEMP_STR As String
Dim DATA_STR As String
Dim REFER_STR As String
Dim SRC_URL_STR As String

Dim TEMP_GROUP() As String
Dim TEMP_MATRIX() As String

On Error GoTo ERROR_LABEL

REFER_STR = "Alphabetical"
SRC_URL_STR = "http://biz.yahoo.com/ic/ind_index.html"
DATA_STR = SAVE_WEB_DATA_PAGE_FUNC(SRC_URL_STR, 0, False, 0, False)
DATA_STR = Replace(DATA_STR, Chr(10), " ")
DATA_STR = Replace(DATA_STR, "&amp;", "&")

i = InStr(1, DATA_STR, REFER_STR) + Len(REFER_STR)
REFER_STR = _
"</a></font></td></tr></table></td></tr></table>" & _
"</td><td>&nbsp&nbsp&nbsp</td><td"
j = InStr(i, DATA_STR, REFER_STR)
DATA_STR = Mid(DATA_STR, i, j - i + Len(REFER_STR))

ReDim TEMP_GROUP(1 To 1)
j = 1
k = 0
Do
    i = InStr(j, DATA_STR, "/ic/")
    If i = 0 Then: Exit Do
    i = i + Len("/ic/")
    
    j = InStr(i, DATA_STR, ".html")
    If j = 0 Then: Exit Do
    
    TEMP_STR = Mid(DATA_STR, i, j - i) & ","
    
    i = InStr(j, DATA_STR, ">")
    If i = 0 Then: Exit Do
    i = i + Len(">")
    
    j = InStr(i, DATA_STR, "<")
    If j = 0 Then: Exit Do
    
    TEMP_STR = TEMP_STR & Mid(DATA_STR, i, j - i)
    k = k + 1
    ReDim Preserve TEMP_GROUP(1 To k)
    TEMP_GROUP(k) = TEMP_STR
    If k > 215 Then: Exit Do 'As of January 31, 2008
Loop Until Mid(DATA_STR, j, Len(REFER_STR)) = REFER_STR
 
ReDim TEMP_MATRIX(1 To UBound(TEMP_GROUP), 1 To 2)
For i = 1 To UBound(TEMP_GROUP)
    REFER_STR = TEMP_GROUP(i)
    If REFER_STR = "" Then: GoTo 1983
    j = InStr(1, REFER_STR, ",")
    TEMP_MATRIX(i, 1) = Mid(REFER_STR, 1, j - 1)
    TEMP_MATRIX(i, 2) = Mid(REFER_STR, j + 1, Len(REFER_STR) - j)
1983:
Next i
            
YAHOO_INDUSTRY_QUOTES_ELEMENTS_FUNC = TEMP_MATRIX

Exit Function
ERROR_LABEL:
YAHOO_INDUSTRY_QUOTES_ELEMENTS_FUNC = Err.number
End Function
