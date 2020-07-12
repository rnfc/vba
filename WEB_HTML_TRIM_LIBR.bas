Attribute VB_Name = "WEB_HTML_TRIM_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.

Function TRIM_HTML_TEXT_STRING_FUNC(ByVal HTML_DATA_STR As String, _
Optional ByVal IMAGE_TEXT_STR As String = "") As String
  
'Adapted from http://www.vbarchiv.de/

Dim i As Integer
Dim j As Integer

Dim CHR_STR As String
Dim HTML_STR As String
Dim HREF_STR As String

Dim PARAM_STR As String
Dim IMG_ALT_STR As String
Dim IMG_FILE_STR As String

Dim BLOCK_FLAG As Boolean

On Error Resume Next

' Prüfen, ob <BODY>-Tag vorhanden
' wenn ja befindet sich der relevante HTML-Teil
' zwischen <BODY> und </BODY>
If InStr(LCase$(HTML_DATA_STR), "<body") > 0 Then
  HTML_DATA_STR = Mid$(HTML_DATA_STR, InStr(LCase$(HTML_DATA_STR), "<body"))
  HTML_DATA_STR = Mid$(HTML_DATA_STR, InStr(HTML_DATA_STR, ">") + 1)
  If InStr(LCase$(HTML_DATA_STR), "</body>") > 0 Then _
    HTML_DATA_STR = Left$(HTML_DATA_STR, InStr(LCase$(HTML_DATA_STR), _
    "</body>") - 1)
End If

' HTML-Text zeichenweise "scannen"
Do While Len(HTML_DATA_STR)
  CHR_STR = Left$(HTML_DATA_STR, 1)
  HTML_DATA_STR = Mid$(HTML_DATA_STR, 2)
  Select Case CHR_STR
    Case " "
      HTML_DATA_STR = LTrim$(HTML_DATA_STR)
    Case vbCr, vbLf
      If Right$(HTML_STR, 1) <> " " And _
        Right$(HTML_STR, 2) <> vbCrLf Then _
        CHR_STR = " " Else CHR_STR = ""
      If Left$(HTML_DATA_STR, 1) = vbLf Then _
        HTML_DATA_STR = Mid$(HTML_DATA_STR, 2)
      HTML_DATA_STR = LTrim$(HTML_DATA_STR)

    ' HTML-Steuerzeichen
    Case "<"
      CHR_STR = ""
      If InStr(HTML_DATA_STR, ">") > 0 Then
        CHR_STR = Left$(HTML_DATA_STR, InStr(HTML_DATA_STR, ">") - 1)
        HTML_DATA_STR = Mid$(HTML_DATA_STR, InStr(HTML_DATA_STR, ">") + 1)

        If InStr(CHR_STR, " ") > 0 Then
          PARAM_STR = REPLACE_CHR_FUNC(Trim$(Mid$(CHR_STR, _
            InStr(CHR_STR, " ") + 1)), vbCrLf, "")
          CHR_STR = Left$(CHR_STR, InStr(CHR_STR, " ") - 1)
        Else
          PARAM_STR = ""
        End If
        Dim ParamColl As Collection

        Select Case LCase$(CHR_STR)
          Case "p"
            If Right$(HTML_STR, 4) <> vbCrLf & vbCrLf Then _
              CHR_STR = vbCrLf & vbCrLf Else CHR_STR = ""
          Case "/div"
            If Right$(HTML_STR, 2) <> vbCrLf Then _
              CHR_STR = vbCrLf Else CHR_STR = ""
          Case "br"
            CHR_STR = vbCrLf
          Case "ul", "/ul", "ol", "/ol"
            CHR_STR = vbCrLf
          Case "li"
            CHR_STR = vbCrLf & "   - "
          Case "BLOCK_FLAG"
            BLOCK_FLAG = True
            CHR_STR = ""
          Case "/BLOCK_FLAG"
            BLOCK_FLAG = False
            CHR_STR = ""
          Case "img"
            If Len(IMAGE_TEXT_STR) > 0 Then
              Set ParamColl = STRIP_CHR_QUOTES_OBJ_FUNC(PARAM_STR, " ")
              IMG_ALT_STR = ""
              IMG_FILE_STR = ""
              For j = 1 To ParamColl.COUNT
                If LCase$(Left$(ParamColl(j), 4)) = "src=" Then
                  IMG_FILE_STR = STRIP_QUOTES_FUNC(Mid$(ParamColl(j), 5))
                ElseIf LCase$(Left$(ParamColl(j), 4)) = "alt=" Then
                  IMG_ALT_STR = STRIP_QUOTES_FUNC(Mid$(ParamColl(j), 5))
                End If
                If Len(IMG_ALT_STR) > 0 And Len(IMG_FILE_STR) > 0 Then Exit For
              Next
              If Len(Trim$(IMG_ALT_STR)) = 0 Then _
                IMG_ALT_STR = "Image " & IMG_FILE_STR
              IMAGE_TEXT_STR = REPLACE_CHR_FUNC(IMAGE_TEXT_STR, "%IMG_ALT_STR", IMG_ALT_STR)
              IMAGE_TEXT_STR = REPLACE_CHR_FUNC(IMAGE_TEXT_STR, "%imgsrc", IMG_FILE_STR)

              CHR_STR = " " & IMAGE_TEXT_STR & " "
            Else
              CHR_STR = ""
            End If
          Case "a"
            Set ParamColl = STRIP_CHR_QUOTES_OBJ_FUNC(PARAM_STR, " ")
            HREF_STR = ""
            For j = 1 To ParamColl.COUNT
              If LCase$(Left$(ParamColl(j), 5)) = "href=" Then
                HREF_STR = STRIP_QUOTES_FUNC(Mid$(ParamColl(j), 6))
                Exit For
              End If
            Next
            CHR_STR = ""
          Case "/a"
            CHR_STR = ""
            If Len(Trim$(HREF_STR)) > 0 Then
              If LCase$(Left$(HREF_STR, 7)) = "mailto:" Then _
                HREF_STR = Mid$(HREF_STR, 8)
              If LCase$(Right$(Trim$(HTML_STR), _
                Len(HREF_STR))) <> LCase$(HREF_STR) Then
                If Not (LCase$(Left$(HREF_STR, 7)) = "http://" And _
                 LCase$(Right$(Trim$(HTML_STR), _
                 Len(HREF_STR) - 7)) = LCase$(Mid$(HREF_STR, 8))) Then
                  CHR_STR = " [" & HREF_STR & "]"
                  HREF_STR = ""
                End If
              End If
            End If
          Case "hr"
            CHR_STR = vbCrLf
            For i = 1 To 70
              CHR_STR = CHR_STR & "-"
            Next
            CHR_STR = CHR_STR & vbCrLf
          Case "sigboundary"
            CHR_STR = vbCrLf & "-- " & vbCrLf
          Case "script"
            CHR_STR = ""
            If InStr(LCase$(HTML_DATA_STR), "</script>") > 0 Then _
              HTML_DATA_STR = Mid$(HTML_DATA_STR, InStr(LCase$(HTML_DATA_STR), _
                "</script>"))
          Case "pre"
            CHR_STR = vbCrLf
            If InStr(LCase$(HTML_DATA_STR), "</pre>") > 0 Then
              CHR_STR = CHR_STR & Left$(HTML_DATA_STR, _
                InStr(LCase$(HTML_DATA_STR), "</pre>") - 1)
              HTML_DATA_STR = Mid$(HTML_DATA_STR, InStr(LCase$(HTML_DATA_STR), _
                "</pre>"))
            End If
            CHR_STR = CHR_STR & vbCrLf
          Case Else
            CHR_STR = ""
        End Select
      End If
    Case "&"
      If InStr(HTML_DATA_STR, ";") > 0 And (InStr(HTML_DATA_STR, ";") < _
        InStr(HTML_DATA_STR, " ") Or InStr(HTML_DATA_STR, " ") = 0) Then
        CHR_STR = Left$(HTML_DATA_STR, InStr(HTML_DATA_STR, ";") - 1)
        HTML_DATA_STR = Mid$(HTML_DATA_STR, InStr(HTML_DATA_STR, ";") + 1)

        Select Case CHR_STR
          Case "amp"
            CHR_STR = "&"
          Case "quot"
            CHR_STR = """"
          Case "lt"
            CHR_STR = "<"
          Case "gt"
            CHR_STR = ">"
          Case "nbsp"
            CHR_STR = " "
          Case "Auml"
            CHR_STR = "Ä"
          Case "auml"
            CHR_STR = "ä"
          Case "iexcl"
            CHR_STR = "¡"
          Case "cent"
            CHR_STR = "¢"
          Case "pound"
            CHR_STR = "£"
          Case "curren"
            CHR_STR = "¤"
          Case "yen"
            CHR_STR = "¥"
          Case "brvbar"
            CHR_STR = "|"
          Case "sect"
            CHR_STR = "§"
          Case "uml"
            CHR_STR = "¨"
          Case "copy"
            CHR_STR = "©"
          Case "ordf"
            CHR_STR = "ª"
          Case "laquo"
            CHR_STR = "«"
          Case "not"
            CHR_STR = "¬"
          Case "reg"
            CHR_STR = "®"
          Case "macr"
            CHR_STR = "¯"
          Case "deg"
            CHR_STR = "°"
          Case "plusm"
            CHR_STR = "±"
          Case "sup2"
            CHR_STR = "²"
          Case "sup3"
            CHR_STR = "³"
          Case "acute"
            CHR_STR = "´"
          Case "micro"
            CHR_STR = "µ"
          Case "para"
            CHR_STR = "¶"
          Case "middot"
            CHR_STR = "·"
          Case "cedil"
            CHR_STR = "¸"
          Case "sup1"
            CHR_STR = "¹"
          Case "ordm"
            CHR_STR = "º"
          Case "raquo"
            CHR_STR = "»"
          Case "frac14"
            CHR_STR = "¼"
          Case "frac12"
            CHR_STR = "½"
          Case "frac34"
            CHR_STR = "¾"
          Case "iquest"
            CHR_STR = "¿"
          Case "Agrave"
            CHR_STR = "À"
          Case "Aacute"
            CHR_STR = "Á"
          Case "Acirc"
            CHR_STR = "Â"
          Case "Atilde"
            CHR_STR = "Ã"
          Case "Aring"
            CHR_STR = "Å"
          Case "AElig"
            CHR_STR = "Æ"
          Case "Ccedil"
            CHR_STR = "Ç"
          Case "Egrave"
            CHR_STR = "È"
          Case "Eacute"
            CHR_STR = "É"
          Case "Ecirc"
            CHR_STR = "Ê"
          Case "Euml"
            CHR_STR = "Ë"
          Case "Igrave"
            CHR_STR = "Ì"
          Case "Iacute"
            CHR_STR = "Í"
          Case "Icirc"
            CHR_STR = "Î"
          Case "Iuml"
            CHR_STR = "Ï"
          Case "ETH"
            CHR_STR = "Ð"
          Case "Ntilde"
            CHR_STR = "Ñ"
          Case "Ograve"
            CHR_STR = "Ò"
          Case "Oacute"
            CHR_STR = "Ó"
          Case "Ocirc"
            CHR_STR = "Ô"
          Case "Otilde"
            CHR_STR = "Õ"
          Case "Ouml"
            CHR_STR = "Ö"
          Case "times"
            CHR_STR = "×"
          Case "Oslash"
            CHR_STR = "Ø"
          Case "Ugrave"
            CHR_STR = "Ù"
          Case "Uacute"
            CHR_STR = "Ú"
          Case "Ucirc"
            CHR_STR = "Û"
          Case "Uuml"
            CHR_STR = "Ü"
          Case "Yacute"
            CHR_STR = "Ý"
          Case "THORN"
            CHR_STR = "Þ"
          Case "szlig"
            CHR_STR = "ß"
          Case "agrave"
            CHR_STR = "à"
          Case "aacute"
            CHR_STR = "á"
          Case "acirc"
            CHR_STR = "â"
          Case "atilde"
            CHR_STR = "ã"
          Case "aring"
            CHR_STR = "å"
          Case "aelig"
            CHR_STR = "æ"
          Case "ccedil"
            CHR_STR = "ç"
          Case "egrave"
            CHR_STR = "è"
          Case "eacute"
            CHR_STR = "é"
          Case "ecirc"
            CHR_STR = "ê"
          Case "euml"
            CHR_STR = "ë"
          Case "igrave"
            CHR_STR = "ì"
          Case "iacute"
            CHR_STR = "í"
          Case "icirc"
            CHR_STR = "î"
          Case "iuml"
            CHR_STR = "ï"
          Case "eth"
            CHR_STR = "ð"
          Case "ntilde"
            CHR_STR = "ñ"
          Case "ograve"
            CHR_STR = "ò"
          Case "oacute"
            CHR_STR = "ó"
          Case "ocirc"
            CHR_STR = "ô"
          Case "otilde"
            CHR_STR = "õ"
          Case "ouml"
            CHR_STR = "ö"
          Case "divide"
            CHR_STR = "÷"
          Case "oslash"
            CHR_STR = "ø"
          Case "ugrave"
            CHR_STR = "ù"
          Case "uacute"
            CHR_STR = "ú"
          Case "ucirc"
            CHR_STR = "û"
          Case "uuml"
            CHR_STR = "ü"
          Case "yacute"
            CHR_STR = "ý"
          Case "thorn"
            CHR_STR = "þ"
          Case "yuml"
            CHR_STR = "ÿ"
          Case Else
            CHR_STR = "&" & CHR_STR & ";"
        End Select
      End If
  End Select
  If Right$(CHR_STR, 2) = vbCrLf And BLOCK_FLAG Then _
    CHR_STR = CHR_STR & "> "
  HTML_STR = HTML_STR & CHR_STR
Loop

HTML_STR = Trim$(HTML_STR)

Do While Left$(HTML_STR, 2) = vbCrLf
  HTML_STR = Trim$(Mid$(HTML_STR, 3))
Loop
Do While Right$(HTML_STR, 2) = vbCrLf
  HTML_STR = Trim$(Left$(HTML_STR, Len(HTML_STR) - 2))
Loop

TRIM_HTML_TEXT_STRING_FUNC = HTML_STR
End Function

Private Function REPLACE_CHR_FUNC(ByVal DATA_STR As String, _
ByVal REPLACE_STR As String, _
ByVal INSERT_STR As String) As String
 
Dim i As Long
Dim j As Long

i = 0
Do While InStr(i + 1, LCase$(DATA_STR), LCase$(REPLACE_STR)) <> 0
  j = InStr(i + 1, LCase$(DATA_STR), LCase$(REPLACE_STR))
  DATA_STR = Left$(DATA_STR, InStr(i + 1, LCase$(DATA_STR), _
    LCase$(REPLACE_STR)) - 1) & INSERT_STR & Mid$(DATA_STR, _
    InStr(i + 1, LCase$(DATA_STR), LCase$(REPLACE_STR)) + _
    Len(REPLACE_STR))
  i& = j + Len(INSERT_STR) - 1
Loop

REPLACE_CHR_FUNC = DATA_STR

Exit Function
ERROR_LABEL:
REPLACE_CHR_FUNC = Err.number
End Function

Private Function STRIP_CHR_QUOTES_OBJ_FUNC(ByVal DATA_STR As String, _
ByVal STRIP_CHR_STR As String) As Collection

Dim i As Long
Dim j As Long
Dim TEMP_OBJ As Collection

On Error GoTo ERROR_LABEL

Set TEMP_OBJ = New Collection
Do
  i = InStr(DATA_STR, STRIP_CHR_STR)
  j = InStr(DATA_STR, Chr$(34))

  If j > 0 And j < i Then
    j = InStr(j + 1, DATA_STR, Chr$(34))
    Do While j > i And j > 0 And i > 0
      i = InStr(i + 1, DATA_STR, STRIP_CHR_STR)
    Loop
  End If

  If i > 0 Then
    TEMP_OBJ.Add Left$(DATA_STR, i - 1)
    DATA_STR = Mid$(DATA_STR, i + 1)
  Else
    TEMP_OBJ.Add DATA_STR
    DATA_STR = ""
  End If
Loop While Len(DATA_STR) > 0

Set STRIP_CHR_QUOTES_OBJ_FUNC = TEMP_OBJ
Set TEMP_OBJ = Nothing
  
Exit Function
ERROR_LABEL:
Set STRIP_CHR_QUOTES_OBJ_FUNC = Nothing
End Function

Private Function STRIP_QUOTES_FUNC(ByVal TEXT_STR As String) As String

On Error GoTo ERROR_LABEL

If Left$(TEXT_STR, 1) = Chr$(34) Then TEXT_STR = Mid$(TEXT_STR, 2)
If Right$(TEXT_STR, 1) = Chr$(34) Then TEXT_STR = Left$(TEXT_STR, Len(TEXT_STR) - 1)

Exit Function
ERROR_LABEL:
STRIP_QUOTES_FUNC = Err.number
End Function


'HTML Characters
'  &#00;-&#08; Unused
'  &#09; \t Horizontal tab
'  &#10; \n Line feed
'  &#11;-&#12; Unused
'  &#13; \r Carriage Return
'  &#14;-&#31; Unused
'  &#32; \s Space
'  &#33; ! Exclamation mark
'  &#34; " Quotation mark
'  &#35; # Number sign
'  &#36; $ Dollar sign
'  &#37; % Percent sign
'  &#38;(&amp;) & Ampersand
'  &#39; ' Apostrophe
'  &#40; ( Left parenthesis
'  &#41; ) Right parenthesis
'  &#42; * Asterisk
'  &#43; + Plus sign
'  &#44; , Comma
'  &#45; - Hyphen
'  &#46; . Period (fullstop)
'  &#47; / Solidus (slash)
'  &#48;-&#57; 0-9 Digits 0-9
'  &#58; : Colon
'  &#59; ; Semi-colon
'  &#60;(&lt;) < Less than
'  &#61; = Equals sign
'  &#62;(&gt;) > Greater than
'  &#63; ? Question mark
'  &#64; @ Commercial at
'  &#65;-&#90; A-Z Letters A-Z
'  &#91; [ Left square bracket
'  &#92; \ Reverse solidus (backslash)
'  &#93; ] Right square bracket
'  &#94; ^ Caret
'  &#95; _ Horizontal bar (underscore)
'  &#96; ` Acute accent
'  &#97;-&#122; a-z Letters a-z
'  &#123; { Left curly brace
'  &#124; | Vertical bar
'  &#125; } Right curly brace
'  &#126; ~ Tilde
'  &#127;-&#159; Unused
'  &#160;(&nbsp;) \s(nb) Non-breaking Space
'  &#161; ¡ Inverted exclamation
'  &#162; ¢ Cent sign
'  &#163; £ Pound sterling
'  &#164; ¤ General currency sign
'  &#165; ¥ Yen sign
'  &#166; ¦ Broken vertical bar
'  &#167; § Section sign
'  &#168; ¨ Umlaut (dieresis)
'  &#169;(&copy;) © Copyright
'  &#170; ª Feminine ordinal
'  &#171; « Left angle quote, guillemotleft
'  &#172; ¬ Not sign
'  &#173; ­ Soft hyphen
'  &#174; ® Registered trademark
'  &#175; ¯ Macron accent
'  &#176; ° Degree sign
'  &#177; ± Plus or minus
'  &#178; ² Superscript two
'  &#179; ³ Superscript three
'  &#180; ´ Acute accent
'  &#181; µ Micro sign
'  &#182; ¶ Paragraph sign
'  &#183; · Middle dot
'  &#184; ¸ Cedilla
'  &#185; ¹ Superscript one
'  &#186; º Masculine ordinal
'  &#187; » Right angle quote, guillemotright
'  &#188;(&frac14;) ¼ Fraction one-fourth
'  &#189;(&frac12;) ½ Fraction one-half
'  &#190;(&frac34;) ¾ Fraction three-fourths
'  &#191; ¿ Inverted question mark
'  &#192; À Capital A, grave accent
'  &#193; Á Capital A, acute accent
'  &#194; Â Capital A, circumflex accent
'  &#195; Ã Capital A, tilde
'  &#196; Ä Capital A, dieresis or umlaut mark
'  &#197; Å Capital A, ring
'  &#198; Æ Capital AE dipthong (ligature)
'  &#199; Ç Capital C, cedilla
'  &#200; È Capital E, grave accent
'  &#201; É Capital E, acute accent
'  &#202; Ê Capital E, circumflex accent
'  &#203; Ë Capital E, dieresis or umlaut mark
'  &#204; Ì Capital I, grave accent
'  &#205; Í Capital I, acute accent
'  &#206; Î Capital I, circumflex accent
'  &#207; Ï Capital I, dieresis or umlaut mark
'  &#208; Ð Capital Eth, Icelandic
'  &#209; Ñ Capital N, tilde
'  &#210; Ò Capital O, grave accent
'  &#211; Ó Capital O, acute accent
'  &#212; Ô Capital O, circumflex accent
'  &#213; Õ Capital O, tilde
'  &#214; Ö Capital O, dieresis or umlaut mark
'  &#215; × Multiply sign
'  &#216; Ø Capital O, slash
'  &#217; Ù Capital U, grave accent
'  &#218; Ú Capital U, acute accent
'  &#219; Û Capital U, circumflex accent
'  &#220; Ü Capital U, dieresis or umlaut mark
'  &#221; Ý Capital Y, acute accent
'  &#222; Þ Capital THORN, Icelandic
'  &#223; ß Small sharp s, German (sz ligature)
'  &#224; à Small a, grave accent
'  &#225; á Small a, acute accent
'  &#226; â Small a, circumflex accent
'  &#227; ã Small a, tilde
'  &#228; ä Small a, dieresis or umlaut mark
'  &#229; å Small a, ring
'  &#230; æ Small ae dipthong (ligature)
'  &#231; ç Small c, cedilla
'  &#232; è Small e, grave accent
'  &#233; é Small e, acute accent
'  &#234; ê Small e, circumflex accent
'  &#235; ë Small e, dieresis or umlaut mark
'  &#236; ì Small i, grave accent
'  &#237; í Small i, acute accent
'  &#238; î Small i, circumflex accent
'  &#239; ï Small i, dieresis or umlaut mark
'  &#240; ð Small eth, Icelandic
'  &#241; ñ Small n, tilde
'  &#242; ò Small o, grave accent
'  &#243; ó Small o, acute accent
'  &#244; ô Small o, circumflex accent
'  &#245; õ Small o, tilde
'  &#246; ö Small o, dieresis or umlaut mark
'  &#247; ÷ Division sign
'  &#248; ø Small o, slash
'  &#249; ù Small u, grave accent
'  &#250; ú Small u, acute accent
'  &#251; û Small u, circumflex accent
'  &#252; ü Small u, dieresis or umlaut mark
'  &#253; ý Small y, acute accent
'  &#254; þ Small thorn, Icelandic
'  &#255; ÿ Small y, dieresis or umlaut mark
