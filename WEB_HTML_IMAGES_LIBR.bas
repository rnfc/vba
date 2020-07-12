Attribute VB_Name = "WEB_HTML_IMAGES_LIBR"


Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.

Function PRINT_HTML_IMAGES_DOC_FUNC( _
ByRef URL_VECTOR As Variant, _
ByRef HEADER_VECTOR As Variant, _
Optional ByVal NSIZE As Integer = 2, _
Optional ByVal DST_FILE_NAME As String = "")

Dim i As Integer
Dim BODY_STR As String

On Error GoTo ERROR_LABEL

PRINT_HTML_IMAGES_DOC_FUNC = False

BODY_STR = CREATE_HTML_IMAGES_STR_FUNC(URL_VECTOR, HEADER_VECTOR, NSIZE)

If DST_FILE_NAME = "" Then
    '"forecasting_" &
    DST_FILE_NAME = Format(Now, "yymmddhhmmss") & ".html"
    If WRITE_TEMP_HTML_TEXT_FILE_FUNC(DST_FILE_NAME, "", BODY_STR) = True Then
        PRINT_HTML_IMAGES_DOC_FUNC = True
    End If
Else
   i = FreeFile
      Open DST_FILE_NAME For Output As #i
         Print #i, BODY_STR;
   Close #i
   PRINT_HTML_IMAGES_DOC_FUNC = True
End If

Exit Function
ERROR_LABEL:
PRINT_HTML_IMAGES_DOC_FUNC = False
End Function

'// PERFECT

Function CREATE_HTML_IMAGES_STR_FUNC(ByRef URL_VECTOR As Variant, _
ByRef HEADER_VECTOR As Variant, _
Optional ByVal NSIZE As Integer = 3)

Dim i As Integer
Dim j As Integer
Dim k As Integer

Dim NROWS As Integer

Dim CR_STR As String
Dim DCT_STR As String
Dim BODY_STR As String

Dim TEMP_STR As String
Dim REF_URL_STR As String
Dim SRC_URL_STR As String


On Error GoTo ERROR_LABEL

If UBound(HEADER_VECTOR, 1) <> UBound(URL_VECTOR, 1) Then: GoTo ERROR_LABEL
NROWS = UBound(URL_VECTOR, 1)

BODY_STR = ""
CR_STR = Chr$(13) & Chr$(10)
DCT_STR = "<TR>"

    
BODY_STR = BODY_STR & _
"<CENTER>&nbsp;<TABLE cellSpacing=3 cellPadding=10 width=" & 100 _
& " border=5>" & CR_STR
    
BODY_STR = BODY_STR & "  <TBODY>" & CR_STR
BODY_STR = BODY_STR & DCT_STR

k = 0

For j = 1 To NROWS
    For i = 1 To NSIZE
        If (k + i) <= NROWS Then
           TEMP_STR = "    <TD>" & HEADER_VECTOR(i + k, 1) & "<BR><A"
           BODY_STR = BODY_STR & TEMP_STR & CR_STR
           
           REF_URL_STR = URL_VECTOR(i + k, 1) 'Reference
           If UBound(URL_VECTOR, 2) > 1 Then
               SRC_URL_STR = URL_VECTOR(i + k, 2) 'Source of Image
           Else
               SRC_URL_STR = URL_VECTOR(i + k, 1) 'Source of Image
           End If
           TEMP_STR = "      href=" & """" & REF_URL_STR & """" & "><IMG"
           BODY_STR = BODY_STR & TEMP_STR & CR_STR
           
           TEMP_STR = "      src=" & """" & SRC_URL_STR & """" & "></A></TD>"
           BODY_STR = BODY_STR & TEMP_STR & CR_STR
        End If
    Next i
    k = k + i - 1
    BODY_STR = BODY_STR & DCT_STR
Next j


CREATE_HTML_IMAGES_STR_FUNC = BODY_STR

Exit Function
ERROR_LABEL:
CREATE_HTML_IMAGES_STR_FUNC = Err.number
End Function


'User defined function to create a table of images/references

Function CREATE_HTML_IMAGE_STR_FUNC(ByRef URL_VECTOR As Variant, _
Optional ByVal BREAKS_VAL As Integer = -1)
    
Dim i As Integer
Dim NSIZE As Integer

Dim TEMP_STR As String
Dim BODY_STR As String

Dim PREFIX_1_STR As String
Dim SUFFIX_1_STR As String

Dim PREFIX_2_STR As String
Dim SUFFIX_2_STR As String

     
On Error GoTo ERROR_LABEL

NSIZE = UBound(URL_VECTOR, 1)

'-----------------------------------------------------------------------------------
If BREAKS_VAL >= 0 Then
'-----------------------------------------------------------------------------------
   PREFIX_1_STR = ""
   SUFFIX_1_STR = ""
   PREFIX_2_STR = ""
   SUFFIX_2_STR = Replace(String(BREAKS_VAL, "!"), "!", "<br>", 1, -1, vbBinaryCompare)
'-----------------------------------------------------------------------------------
Else
'-----------------------------------------------------------------------------------
   PREFIX_1_STR = "<table>"
   SUFFIX_1_STR = "</table>"
   
   PREFIX_2_STR = "<tr><td>"
   SUFFIX_2_STR = "</td></tr>"
'-----------------------------------------------------------------------------------
End If
'-----------------------------------------------------------------------------------

BODY_STR = PREFIX_1_STR
For i = 1 To NSIZE
    TEMP_STR = "<img src=""" & URL_VECTOR(i, 1) & """>"
    TEMP_STR = PREFIX_2_STR & TEMP_STR & SUFFIX_2_STR
    BODY_STR = BODY_STR & TEMP_STR
Next i

BODY_STR = BODY_STR & SUFFIX_1_STR


CREATE_HTML_IMAGE_STR_FUNC = BODY_STR

Exit Function
ERROR_LABEL:
CREATE_HTML_IMAGE_STR_FUNC = Err.number
End Function


