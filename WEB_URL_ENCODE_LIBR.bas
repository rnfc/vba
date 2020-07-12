Attribute VB_Name = "WEB_URL_ENCODE_LIBR"


Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.


Private Const MAX_PATH As Long = 260
Private Const ERROR_SUCCESS As Long = 0

'Treat entire URL param as one URL segment
'Replace only spaces with escape
'sequences. This flag takes precedence
'over URL_ESCAPE_UNSAFE, but does not
'apply to opaque URLs.

Private Const URL_ESCAPE_SPACES_ONLY As Long = &H4000000
'URL_ESCAPE_SPACES_ONLY: Only escape space characters. This flag
'cannot be combined with URL_ESCAPE_PERCENT or URL_ESCAPE_SEGMENT_ONLY.
'UrlEscape


Private Const URL_ESCAPE_SEGMENT_ONLY    As Long = &H2000
'URL_ESCAPE_SEGMENT_ONLY: Escape the sections following the server
'component, but not the extra information sections following a
'# or ? character

Private Const URL_ESCAPE_PERCENT         As Long = &H1000
'URL_ESCAPE_PERCENT: Escape the % character. By default, this character
'is not escaped


'Replace unsafe values with their
'escape sequences. This flag applies
'to all URLs, including opaque URLs.
Private Const URL_ESCAPE_UNSAFE       As Long = &H20000000
'URL_ESCAPE_UNSAFE: Replace unsafe values with their escape sequences.
'This flag applies to all URLs, including opaque URLs.


'Unescape any escape sequences that
'the URLs contain, with two exceptions.
'The escape sequences for '?' and '#'
'will not be unescaped. If one of the
'URL_ESCAPE_XXX flags is also set, the
'two URLs will unescaped, then combined,
'then escaped.
Private Const URL_UNESCAPE            As Long = &H10000000
'URL_UNESCAPE: Unescape any escape sequences that the URLs contain,
'with two exceptions. The escape sequences for '?' and '#' will not
'be unescaped. If one of the URL_ESCAPE_XXX flags is also set, the
'two URLs will unescaped, then combined, then escaped.

Private Const URL_UNESCAPE_INPLACE       As Long = &H100000
'URL_UNESCAPE_INPLACE: Use pszURL to return the converted string
'instead of pszUnescaped

'escape #'s in paths
Private Const URL_INTERNAL_PATH          As Long = &H800000
Private Const URL_DONT_ESCAPE_EXTRA_INFO As Long = &H2000000
'URL_DONT_ESCAPE_EXTRA_INFO: Don't convert the # or ? character, or
'any characters following them in the string.

Private Const URL_DONT_SIMPLIFY          As Long = &H8000000
'URL_DONT_SIMPLIFY: '/./' and '/../' in a URL string as literal
'characters, not as shorthand for navigation.

'Combine URLs with client-defined
'pluggable protocols, according to
'the W3C specification. This flag
'does not apply to standard protocols
'such as ftp, http, gopher, and so on.
'If this flag is set, UrlCombine will
'not simplify URLs, so there is no need
'to also set URL_DONT_SIMPLIFY.
Private Const URL_PLUGGABLE_PROTOCOL  As Long = &H40000000
'URL_PLUGGABLE_PROTOCOL: Combine URLs with client-defined pluggable
'protocols, according to the W3C specification. This flag does not
'apply to standard protocols such as ftp, http, gopher, and so on.
'If this flag is set, UrlCombine will not simplify URLs, so there is
'no need to also set URL_DONT_SIMPLIFY.

'Converts unsafe characters, such as spaces, into their
'corresponding escape sequences.
Private Declare PtrSafe Function UrlEscape Lib "shlwapi" _
   Alias "UrlEscapeA" _
  (ByVal pszUrl As String, _
   ByVal pszEscaped As String, _
   pcchEscaped As Long, _
   ByVal dwFlags As Long) As Long

'Converts escape sequences back into ordinary characters.
Private Declare PtrSafe Function UrlUnescape Lib "shlwapi" _
   Alias "UrlUnescapeA" _
  (ByVal pszUrl As String, _
   ByVal pszUnescaped As String, _
   pcchUnescaped As Long, _
   ByVal dwFlags As Long) As Long

Private Declare PtrSafe Function UrlCanonicalize Lib "shlwapi" _
   Alias "UrlCanonicalizeA" _
  (ByVal pszUrl As String, _
   ByVal pszCanonicalized As String, _
   pcchCanonicalized As Long, _
   ByVal dwFlags As Long) As Long

Function ENCODE_URL_FUNC(ByVal SRC_URL_STR As String)

Dim ii As Long
Dim jj As Long

Dim BUFF_STR As String

On Error GoTo ERROR_LABEL

If Len(SRC_URL_STR) > 0 Then
   
   BUFF_STR = Space$(MAX_PATH)
   ii = Len(BUFF_STR)
   jj = URL_DONT_SIMPLIFY
   
   If UrlEscape(SRC_URL_STR, _
                BUFF_STR, _
                ii, _
                jj) = ERROR_SUCCESS Then
                
      ENCODE_URL_FUNC = Left$(BUFF_STR, ii)
   
   End If  'UrlEscape
End If  'Len(SRC_URL_STR)

Exit Function
ERROR_LABEL:
ENCODE_URL_FUNC = Err.number
End Function

Function DECODE_URL_FUNC(ByVal SRC_URL_STR As String)

Dim ii As Long
Dim jj As Long
Dim BUFF_STR As String

On Error GoTo ERROR_LABEL

If Len(SRC_URL_STR) > 0 Then
   
   BUFF_STR = Space$(MAX_PATH)
   ii = Len(BUFF_STR)
   jj = &H20000000
   
   If UrlUnescape(SRC_URL_STR, _
                BUFF_STR, _
                ii, _
                jj) = ERROR_SUCCESS Then
                
      DECODE_URL_FUNC = Left$(BUFF_STR, ii)
   
   End If  'UrlUnescape
End If  'Len(SRC_URL_STR)

Exit Function
ERROR_LABEL:
DECODE_URL_FUNC = Err.number
End Function

Function CANON_ENCODE_URL_FUNC(ByVal SRC_URL_STR As String, _
Optional ByVal VERSION As Long = 0)
   
   
Dim ii As Long
Dim ESC_STR As String
Dim INDEX_FLAG As Long
  
On Error GoTo ERROR_LABEL

If VERSION = 0 Then
   INDEX_FLAG = URL_ESCAPE_UNSAFE
Else
   INDEX_FLAG = URL_UNESCAPE
End If

If Len(SRC_URL_STR) > 0 Then
   
   ESC_STR = Space$(MAX_PATH)
   ii = Len(ESC_STR)
   
   
   If UrlCanonicalize(SRC_URL_STR, _
                      ESC_STR, _
                      ii, _
                      INDEX_FLAG) = ERROR_SUCCESS Then
                
      CANON_ENCODE_URL_FUNC = Left$(ESC_STR, ii)
   
   End If  'If UrlCanonicalize
Else
   GoTo ERROR_LABEL
End If  'If Len(SRC_URL_STR) > 0

Exit Function
ERROR_LABEL:
CANON_ENCODE_URL_FUNC = Err.number
End Function

'This function converts non-alpha or numeric chars to ASCII equivalent
'so the webserver can read them

Function URL_SAFE_STRING_FUNC(DATA_STR As String) As String
    
Dim i As Long
Dim j As Long

Dim BUFFER_STR As String
Dim TEMP_STR As String

On Error GoTo ERROR_LABEL

j = Len(DATA_STR)
TEMP_STR = DATA_STR
For i = j To 1 Step -1
    BUFFER_STR = Mid(DATA_STR, i, 1)

    If Not BUFFER_STR Like "[a-z,A-Z,0-9]" Then
        TEMP_STR = Left$(TEMP_STR, i - 1) & "%" & _
                    Right$("00" & Hex(Asc(BUFFER_STR)), 2) & _
                    Mid$(TEMP_STR, i + 1)
    End If
Next i
URL_SAFE_STRING_FUNC = TEMP_STR
    
Exit Function
ERROR_LABEL:
URL_SAFE_STRING_FUNC = Err.number
End Function

Function ESCAPE_URL_FUNC(ByVal SRC_URL_STR As String)

Dim i As Long
Dim j As Long
Dim TEMP_CHR As String

On Error GoTo ERROR_LABEL

'ATEMP_ARR = Array(" ", "#", "$", "%", "&", "/", ":", ";", _
                 "<", "=", ">", "?", "@", "[", "\", "]", _
                 "^", "`", "{", "|", "}", "~")

'BTEMP_ARR = Array("%20", "%23", "%24", "%25", "%26", _
                  "%2F", "%3A", "%3B", "%3C", "%3D", "%3E", _
                  "%3F", "%40", "%5B", "%5C", "%5D", "%5E", _
                  "%60", "%7B", "%7C", "%7D", "%7E")

TEMP_CHR = "<>%=&!@#$^()+{[}]|\;:'"",/?"
j = Len(TEMP_CHR)
For i = 1 To j
    SRC_URL_STR = Replace(SRC_URL_STR, Mid(TEMP_CHR, i, 1), "%" & _
                          Hex(Asc(Mid(TEMP_CHR, i, 1))))
Next i

SRC_URL_STR = Replace(SRC_URL_STR, " ", "+")

ESCAPE_URL_FUNC = SRC_URL_STR

Exit Function
ERROR_LABEL:
ESCAPE_URL_FUNC = Err.number
End Function
