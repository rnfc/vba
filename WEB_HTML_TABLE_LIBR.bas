Attribute VB_Name = "WEB_HTML_TABLE_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.

Function CREATE_HTML_DATA_TABLE_FUNC(ByRef DATA_MATRIX As Variant, _
Optional ByVal PAGE_NAME As String = " ", _
Optional ByVal PAGE_TITLE As String = " ", _
Optional ByVal DST_FILE_NAME As String = "")

Dim i As Long
Dim j As Long
Dim k As Long

Dim NROWS As Long
Dim NCOLUMNS As Long

Dim CR_STR As String
Dim DCR_STR As String
Dim QUOT_STR As String
Dim TOP_STR As String
Dim BODY_STR As String
Dim LDGSPC_STR As String

On Error GoTo ERROR_LABEL

CREATE_HTML_DATA_TABLE_FUNC = False

'//////////////////////////////////////////////////////////////////////////////////
' These color values can be either hexadecimal -- like "#ffffff" for white --
' or they can be color names like "blue" or "green"
'//////////////////////////////////////////////////////////////////////////////////

Const TITLE_STR = " & "" & "
Const TITLE_COLOR_STR = "Red" ' page title
Const HEADING_COLOR_STR = "Blue" ' individual headings
Const COMMENT_COLOR_STR = "Black" ' the comments to the right of links
Const TEXT_COLOR_STR = "Black" ' default color for text not covered by those above
Const BACKGROUND_COLOR_STR = "#ffffff" ' use this if there is no background graphic"

'------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------
CR_STR = Chr$(13) & Chr$(10) 'ADJUST STRINGS
DCR_STR = CR_STR & CR_STR 'INSERT TWO ROWS
QUOT_STR = Chr$(34)
LDGSPC_STR = "&nbsp;&nbsp;&nbsp;&nbsp;"
k = 0
'------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------

TOP_STR = ""
TOP_STR = TOP_STR & "<HTML><Head><Title>" & TITLE_STR & _
            "</Title></Head>" & CR_STR & CR_STR
TOP_STR = TOP_STR & "<Body bgcolor=" & BACKGROUND_COLOR_STR & ">" & CR_STR

TOP_STR = TOP_STR & "<a name=" & QUOT_STR & "top" & QUOT_STR & ">"
            ' mark the top of the page

TOP_STR = TOP_STR & "<CENTER><font size =+2 color=" & TITLE_COLOR_STR & ">" & _
PAGE_TITLE & "</font><br><br></A>" & DCR_STR

'------------------------------------------------------------------------------------

BODY_STR = "<font color=" & TEXT_COLOR_STR & ">"
BODY_STR = BODY_STR & "<table width=95% border=0>" & DCR_STR
    
If k = 0 Then: TOP_STR = TOP_STR & "<center><table><tr><td><ul>" & CR_STR

TOP_STR = TOP_STR & "<li><b><a href=" & QUOT_STR & "#" & _
           Chr$(k & 65) & QUOT_STR & ">" & PAGE_NAME & _
           "</a></i></b>" & CR_STR ' mark this location with an anchor
    
BODY_STR = BODY_STR & "<tr><td colspan=3><a name=" & _
            QUOT_STR & Chr$(k + 65) & QUOT_STR & ">" & CR_STR
BODY_STR = BODY_STR & "<font color=" & TITLE_COLOR_STR & _
            " size=+1><b>" & PAGE_NAME & "</b></i></font><br></A>" & CR_STR
BODY_STR = BODY_STR & "</td></tr>" & CR_STR
        
NROWS = UBound(DATA_MATRIX, 1)
NCOLUMNS = UBound(DATA_MATRIX, 2)
    
For i = 1 To NROWS
    BODY_STR = BODY_STR & "<tr>"
        For j = 1 To NCOLUMNS
            If i = 1 Then
                BODY_STR = BODY_STR & "<td><font color=" & _
                                HEADING_COLOR_STR & "><b><u>" & _
                                DATA_MATRIX(i, j) & "</font></td>"
            Else
                BODY_STR = BODY_STR & "<td><font color=" & _
                                COMMENT_COLOR_STR & ">" & _
                                CStr(DATA_MATRIX(i, j)) & _
                                "</font></td>"
            End If
        Next j
    BODY_STR = BODY_STR & "</tr>" & CR_STR
    LDGSPC_STR = "" ' you only need to pad in these spaces _
    once -- now save some file size
Next i
    
BODY_STR = BODY_STR & "<tr><td colspan=3 align=center><i><b><a href=" & _
            QUOT_STR & "#top" & QUOT_STR & ">Back to top</a></b></i></td></tr>" _
            & CR_STR

BODY_STR = BODY_STR & "<tr><td colspan=3>&nbsp;</td></tr>" & DCR_STR
k = k + 1

BODY_STR = BODY_STR & "</table>" & CR_STR
BODY_STR = BODY_STR & "</font>" & CR_STR
BODY_STR = BODY_STR & "<H3><A HREF=" & QUOT_STR & "" & _
           QUOT_STR & ">Main Page</A></H3>" & CR_STR

BODY_STR = BODY_STR & "</center>" & CR_STR
BODY_STR = BODY_STR & "</Body>" & CR_STR & "</HTML>"

TOP_STR = TOP_STR & "</ul></td></tr></table></center><br><br>" & DCR_STR
BODY_STR = TOP_STR & BODY_STR


If DST_FILE_NAME = "" Then
    DST_FILE_NAME = "webpage_" & Format(Now, "yymmddhhmmss") & ".html"
    If WRITE_TEMP_HTML_TEXT_FILE_FUNC(DST_FILE_NAME, "", BODY_STR) = True Then
        CREATE_HTML_DATA_TABLE_FUNC = True
    End If
Else
   i = FreeFile
      Open DST_FILE_NAME For Output As #i
         Print #i, BODY_STR;
   Close #i
   CREATE_HTML_DATA_TABLE_FUNC = True
End If

Exit Function
ERROR_LABEL:
CREATE_HTML_DATA_TABLE_FUNC = False
End Function
