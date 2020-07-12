Attribute VB_Name = "WEB_HTML_DOC_LIBR"
        

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.

'The subroutine below navigates to a specified URL, uses the values in the
'ParamArray to load the indicated fields, and sends the form to the server.
'It expects the field values to be in pairs with the first element in the
'pair indicating the field ID and the 2nd the value. If the ParamArray
'contains an even number of elements then the last pair contains the ID
'of a field and the name of the event to be fired.  If, on the other hand,
'the ParamArray contains an odd number of elements then the last element
'is the name of the form that we want to submit.

Function WEB_HTML_SUBMIT_FORM_FUNC(ByRef HTML_DOC_OBJ As HTMLDocument, _
ByVal PARAMS_RNG As Variant)

'Set HTML_DOC_OBJ = New HTMLDocument
'HTML_DOC_OBJ.body.innerHTML = ResponseText

    Dim i As Integer
    
    On Error GoTo ERROR_LABEL

'----------------------------------------------------------------------------------
    WEB_HTML_SUBMIT_FORM_FUNC = False
'----------------------------------------------------------------------------------
    If UBound(PARAMS_RNG) < LBound(PARAMS_RNG) Then GoTo ERROR_LABEL
    For i = LBound(PARAMS_RNG) To UBound(PARAMS_RNG) - 2 Step 2
        HTML_DOC_OBJ.getElementById(PARAMS_RNG(i)).value = PARAMS_RNG(i + 1)
    Next i
'----------------------------------------------------------------------------------
    If UBound(PARAMS_RNG) Mod 2 = 0 Then 'PARAMS_RNG contains an odd number _
    The last is the name of the form to submit
'----------------------------------------------------------------------------------
        HTML_DOC_OBJ.getElementById(PARAMS_RNG(UBound(PARAMS_RNG))).submit
'----------------------------------------------------------------------------------
    Else    'Even number of elements in PARAMS_RNG.  The last pair _
             is the name of an element and its event to fire
'----------------------------------------------------------------------------------
        HTML_DOC_OBJ.getElementById(PARAMS_RNG(UBound(PARAMS_RNG) - 1)) _
            .FireEvent PARAMS_RNG(UBound(PARAMS_RNG))
'If the form has no submit button.  As soon as we select an option, the
'change event occurs and that results in the go() function being called.
'The go function loads the new page.  So, unlike the Google
'algo, we don’t have a form’s submit method to call.  The first attempt at
'implementing this lookup programmatically would be to simply change the
'value of the control and check if the change event fires. We already know
'enough so that we can easily change the value of the drop-down control.
'Once we establish a reference to the document in the InternetExplorer
'window through the HTMLDoc variable, we can use
'HTMLDoc.getElementById("XYZ").Value = {new value}
'If we use the above, we will discover nothing happens. Changing the value
'programmatically doesn’t trigger the change event.  It turns out we have
'to do so ourselves with the FireEvent method.  The statement below invokes
'the FireEvent method and provides it with the name of the event to fire.
'HTMLDoc.getElementById("XYZ").FireEvent "change"
'So, we need this generic subroutine that loads any form on any
'page and then sends it to the server, we need to (a) set an arbitrary
'number of fields to the desired values, and (b) either use the form’s submit
'method or fire an appropriate event for a particular control.
'----------------------------------------------------------------------------------
    End If
'----------------------------------------------------------------------------------

    WEB_HTML_SUBMIT_FORM_FUNC = True

Exit Function
ERROR_LABEL:
WEB_HTML_SUBMIT_FORM_FUNC = False
End Function


Function WEB_HTML_OUTER_PAGE_FUNC(ByVal SRC_URL_STR As String)
Dim CONNECT_OBJ As New HTMLDocument
Dim HTML_DOC_OBJ As New HTMLDocument

On Error GoTo ERROR_LABEL

Set CONNECT_OBJ = HTML_DOC_OBJ.createDocumentFromUrl(SRC_URL_STR, vbNullString)
Do: DoEvents: Loop Until CONNECT_OBJ.readyState = "complete"

WEB_HTML_OUTER_PAGE_FUNC = CONNECT_OBJ.DocumentElement.outerHTML
Exit Function
ERROR_LABEL:
WEB_HTML_OUTER_PAGE_FUNC = Err.number
End Function
