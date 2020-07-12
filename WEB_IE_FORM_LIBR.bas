Attribute VB_Name = "WEB_IE_FORM_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.
                    
Private Type HTML_FORMS_OBJ
    Form As String
    Type As String
    name As String
End Type

Public Function IE_SUBMIT_FORM_FUNC(ByVal SRC_URL_STR As String, _
ParamArray PARAMS_RNG() As Variant)

'Debug.Print IE_SUBMIT_FORM_FUNC("http://www.google.com", _
        "q", _
        "tushar mehta", _
        "f")

    Dim HTML_DOC_OBJ As HTMLDocument
    Dim CONNECT_OBJ As InternetExplorer

'----------------------------------------------------------------------------------
    On Error GoTo ERROR_LABEL
'----------------------------------------------------------------------------------

    Set CONNECT_OBJ = New InternetExplorer
    With CONNECT_OBJ
        .navigate SRC_URL_STR
        .Visible = True
        Do While .Busy: DoEvents: Loop
        Do While .readyState <> 4: DoEvents: Loop
    End With
    
    Set HTML_DOC_OBJ = CONNECT_OBJ.document
    Call WEB_HTML_SUBMIT_FORM_FUNC(HTML_DOC_OBJ, PARAMS_RNG)

    With CONNECT_OBJ
        Do While .Busy: DoEvents: Loop
        Do While .readyState <> 4: DoEvents: Loop
    End With

    If Not CONNECT_OBJ Is Nothing Then CONNECT_OBJ.Quit
    Set CONNECT_OBJ = Nothing

    IE_SUBMIT_FORM_FUNC = HTML_DOC_OBJ.DocumentElement.outerHTML

Exit Function
ERROR_LABEL:
On Error Resume Next
If Not CONNECT_OBJ Is Nothing Then CONNECT_OBJ.Quit
Set CONNECT_OBJ = Nothing
IE_SUBMIT_FORM_FUNC = Err.number
End Function

Public Function IE_FORM_LIST_FUNC(ByVal SRC_URL_STR As String)

Dim i As Integer
Dim j As Integer
Dim l As Integer

Dim TEMP_ARR() As HTML_FORMS_OBJ
Dim TEMP_MATRIX() As String

Dim ITEM_OBJ As HTMLFormElement
Dim HTML_DOC_OBJ As HTMLDocument

Dim CONNECT_OBJ As InternetExplorer

On Error GoTo ERROR_LABEL

Set CONNECT_OBJ = New InternetExplorer

With CONNECT_OBJ
    .navigate SRC_URL_STR
    .Visible = False
    Do While .Busy: DoEvents: Loop
    Do While .readyState <> 4: DoEvents: Loop
End With

Set HTML_DOC_OBJ = CONNECT_OBJ.document
'CONNECT_OBJ.document.all.Item("Username")

ReDim TEMP_ARR(1 To 1)
l = 1

With HTML_DOC_OBJ.forms
    For i = 0 To .length - 1
        Set ITEM_OBJ = .Item(i)
        With ITEM_OBJ
            For j = 0 To .length - 1
                ReDim Preserve TEMP_ARR(1 To l)
                TEMP_ARR(l).Form = .name
                TEMP_ARR(l).Type = .Item(j).Type
                TEMP_ARR(l).name = .Item(j).name
                l = l + 1
            Next j
        End With
    Next i
End With

If Not CONNECT_OBJ Is Nothing Then CONNECT_OBJ.Quit
Set CONNECT_OBJ = Nothing

j = UBound(TEMP_ARR, 1)

ReDim TEMP_MATRIX(0 To j, 1 To 3)
TEMP_MATRIX(0, 1) = "Form"
TEMP_MATRIX(0, 2) = "Type"
TEMP_MATRIX(0, 3) = "Name"

For i = 1 To j
    TEMP_MATRIX(i, 1) = TEMP_ARR(i).Form
    TEMP_MATRIX(i, 2) = TEMP_ARR(i).Type
    TEMP_MATRIX(i, 3) = TEMP_ARR(i).name
Next i

IE_FORM_LIST_FUNC = TEMP_MATRIX

Exit Function
ERROR_LABEL:
On Error Resume Next
If Not CONNECT_OBJ Is Nothing Then CONNECT_OBJ.Quit
Set CONNECT_OBJ = Nothing
IE_FORM_LIST_FUNC = Err.number
End Function
