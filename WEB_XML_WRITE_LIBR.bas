Attribute VB_Name = "WEB_XML_WRITE_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.

Function XML_WRITE_MAP_FUNC(ByRef DATA_MATRIX As Variant, _
Optional ByVal DST_PATH_NAME As String = "C:\Users\NICO\Desktop\nico.xml")

Dim i As Long
Dim j As Long

Dim SROW As Long
Dim NROWS As Long

Dim SCOLUMN As Long
Dim NCOLUMNS As Long

Dim ATEMP_STR As String 'item
Dim BTEMP_STR As String 'field

'Dim DATA_MATRIX As Variant

Dim FIELD_OBJ As Object
Dim DOC_OBJ As New MSXML2.DOMDocument
Dim ROOT_OBJ As MSXML2.IXMLDOMElement
Dim RECORD_OBJ As MSXML2.IXMLDOMElement
    
On Error GoTo ERROR_LABEL

XML_WRITE_MAP_FUNC = False

'DATA_MATRIX = DATA_RNG

SROW = LBound(DATA_MATRIX, 1)
NROWS = UBound(DATA_MATRIX, 1)

SCOLUMN = LBound(DATA_MATRIX, 2)
NCOLUMNS = UBound(DATA_MATRIX, 2)

Set ROOT_OBJ = DOC_OBJ.createElement("Root")
DOC_OBJ.appendChild ROOT_OBJ

For i = (SROW + 1) To NROWS
    Set RECORD_OBJ = DOC_OBJ.createElement("Record")
    ROOT_OBJ.appendChild RECORD_OBJ
    For j = 1 To NCOLUMNS
        BTEMP_STR = DATA_MATRIX(SROW, j)
        GoSub CLEAN_LINE
        ATEMP_STR = DATA_MATRIX(i, j)
        Set FIELD_OBJ = DOC_OBJ.createElement(BTEMP_STR)
        FIELD_OBJ.Text = ATEMP_STR
        RECORD_OBJ.appendChild FIELD_OBJ
    Next j
Next i
    
DOC_OBJ.Save DST_PATH_NAME

XML_WRITE_MAP_FUNC = True

'-----------------------------------------------------------------------------------------
Exit Function
'-----------------------------------------------------------------------------------------
CLEAN_LINE:
'-----------------------------------------------------------------------------------------
    BTEMP_STR = REMOVE_EXTRA_SPACES_FUNC(BTEMP_STR)
    BTEMP_STR = REMOVE_CHARACTERS_FUNC(BTEMP_STR, " /-:;!@#$%^&*()+=,<>")
    BTEMP_STR = XML_CONVERT_ACCENT_FUNC(BTEMP_STR)
'-----------------------------------------------------------------------------------------
Return
'-----------------------------------------------------------------------------------------
ERROR_LABEL:
'-----------------------------------------------------------------------------------------
XML_WRITE_MAP_FUNC = False
End Function

Function XML_CONVERT_ACCENT_FUNC(ByVal INPUT_STR As String) As String
' http://www.vbforums.com/archive/index.php/t-483965.html

Dim i As Long
Dim j As Long
Dim k As Long
Dim DATA_STR As String
Dim CHR_STR As String
Dim MATCH_FLAG As Boolean

On Error GoTo ERROR_LABEL

Const ACC_STR As String = _
"äéöûü¿¡¬√ƒ≈«»… ÀÃÕŒœ–—“”‘’÷Ÿ⁄€‹›‡·‚„‰ÂÁËÈÍÎÏÌÓÔÒÚÛÙıˆ˘˙˚¸˝ˇ"
Const REG_STR As String = _
"SZszYAAAAAACEEEEIIIIDNOOOOOUUUUYaaaaaaceeeeiiiidnooooouuuuyy"

DATA_STR = INPUT_STR

' loop through the shorter string
Select Case True
Case Len(ACC_STR) <= Len(INPUT_STR)
    ' accent character list is shorter (or same)
    ' loop through accent character string
    j = Len(ACC_STR)
    For i = 1 To j
        ' get next accent character
        CHR_STR = Mid(ACC_STR, i, 1)
        ' replace with corresponding character in "regular" array
        If InStr(DATA_STR, CHR_STR) > 0 Then
            DATA_STR = Replace(DATA_STR, CHR_STR, Mid(REG_STR, i, 1))
        End If
    Next i
Case Len(ACC_STR) > Len(INPUT_STR)
    ' input string is shorter
    ' loop through input string
    j = Len(INPUT_STR)
    For i = 1 To j
        ' grab current character from input string and
        ' determine if it is a special char
        CHR_STR = Mid(INPUT_STR, i, 1)
        MATCH_FLAG = (InStr(ACC_STR, CHR_STR) > 0)
        If MATCH_FLAG Then
            ' find position of special character in special array
            k = InStr(ACC_STR, CHR_STR)
            ' replace with corresponding character in "regular" array
            DATA_STR = Replace(DATA_STR, CHR_STR, Mid(REG_STR, k, 1))
        End If
    Next i
End Select

XML_CONVERT_ACCENT_FUNC = DATA_STR

Exit Function
ERROR_LABEL:
XML_CONVERT_ACCENT_FUNC = Err.number
End Function


