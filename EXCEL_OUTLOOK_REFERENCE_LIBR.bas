Attribute VB_Name = "EXCEL_OUTLOOK_REFERENCE_LIBR"
'-------------------------------------------------------------------------------------
Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.
'-------------------------------------------------------------------------------------

'All this code only works if Excel has a reference to
'the Outlook library the sub below ensures we do

Function CHECK_OUTLOOK_REFERENCES_FUNC( _
Optional ByVal SRC_PATH_STR As String = _
"C:\Program Files\Microsoft Office\Office12\msoutl.olb") 'msoutl9.olb

Dim TEMP_REF As Object
Static FIND_FLAG As Boolean

On Error GoTo ERROR_LABEL

If FIND_FLAG = True Then Exit Function

'do we have a reference already? if so, exit
For Each TEMP_REF In Application.VBE.ActiveVBProject.References
  If TEMP_REF.name = "Outlook" Then
    FIND_FLAG = True
    Exit Function
  End If
Next TEMP_REF

'no reference? add it
If FIND_FLAG = False Then
  'if you get an error below, the Outlook library is somewhere
  'else. Search for msoutl9.olb on your C drive, and
  'when you find it, change the pathname on the next line and try again
  Application.VBE.ActiveVBProject.References.AddFromFile SRC_PATH_STR
End If

Exit Function
ERROR_LABEL:
CHECK_OUTLOOK_REFERENCES_FUNC = Err.number
End Function





