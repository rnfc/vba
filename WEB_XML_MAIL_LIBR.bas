Attribute VB_Name = "WEB_XML_MAIL_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.

Function XML_MAIL_DATA_FUNC( _
ByRef DATA_RNG As Variant, _
Optional ByVal DST_ADDRESS_STR As String = "XYZ", _
Optional ByVal REPORT_NAME_STR As String = "PE-Analysis", _
Optional ByVal SUBJECT_STR As String = "Nico", _
Optional ByVal BODY_TEXT_STR As String = "The file")

Dim i As Long
Dim j As Long

Dim SROW As Long
Dim SCOLUMN As Long

Dim NROWS As Long
Dim NCOLUMNS As Long

Dim ITEM_STR As String
Dim FIELD_STR As String

Dim FIELD_OBJ As Object

Dim TEMP_DATE As Date

Dim DAY_VAL As Integer
Dim MONTH_VAL As Integer
Dim YEAR_VAL As Integer

Dim FILE_PATH_STR As String
Dim FILE_NAME_STR As String

Dim DATA_MATRIX As Variant

Dim XML_DOC_OBJ As New MSXML2.DOMDocument
Dim XML_ROOT_OBJ As MSXML2.IXMLDOMElement
Dim XML_RECORD_OBJ As MSXML2.IXMLDOMElement

'Ref Microsoft Outlook 11.0 Object Library
Dim OUTLOOK_OBJ As Outlook.Application
Dim OUTLOOK_MSG_OBJ As Outlook.MailItem
   
On Error GoTo ERROR_LABEL

XML_MAIL_DATA_FUNC = False

DATA_MATRIX = DATA_RNG
SROW = LBound(DATA_MATRIX, 1)
SCOLUMN = LBound(DATA_MATRIX, 2)

NROWS = UBound(DATA_MATRIX, 1)
NCOLUMNS = UBound(DATA_MATRIX, 2)

TEMP_DATE = Date
DAY_VAL = Day(TEMP_DATE)
MONTH_VAL = Month(TEMP_DATE)
YEAR_VAL = Year(TEMP_DATE)

Set XML_ROOT_OBJ = XML_DOC_OBJ.createElement("Root")
XML_DOC_OBJ.appendChild XML_ROOT_OBJ
    
For i = SROW + 1 To NROWS
    Set XML_RECORD_OBJ = XML_DOC_OBJ.createElement("Record")
    XML_ROOT_OBJ.appendChild XML_RECORD_OBJ
    For j = SCOLUMN To NCOLUMNS
        FIELD_STR = DATA_MATRIX(SROW, j)
        ITEM_STR = DATA_MATRIX(i, j)
        Set FIELD_OBJ = XML_DOC_OBJ.createElement(FIELD_STR)
        FIELD_OBJ.Text = ITEM_STR
        XML_RECORD_OBJ.appendChild FIELD_OBJ
    Next j
Next i

' MsgBox XML_DOC_OBJ.FirstChild.XML

FILE_NAME_STR = REPORT_NAME_STR & "-" & YEAR_VAL & "-" & _
                MONTH_VAL & "-" & DAY_VAL

FILE_PATH_STR = Environ("temp") & "\" & FILE_NAME_STR & ".xml"
XML_DOC_OBJ.Save (FILE_PATH_STR)

Set OUTLOOK_OBJ = CreateObject("Outlook.Application")
Set OUTLOOK_MSG_OBJ = OUTLOOK_OBJ.CreateItem(olMailItem)

With OUTLOOK_MSG_OBJ
    .To = DST_ADDRESS_STR
    .Subject = SUBJECT_STR
    .body = BODY_TEXT_STR
    .Attachments.Add (FILE_PATH_STR)
    .ReadReceiptRequested = True
    .display
    .send
End With

Set OUTLOOK_MSG_OBJ = Nothing
Set OUTLOOK_OBJ = Nothing

Set OUTLOOK_MSG_OBJ = Nothing
Set OUTLOOK_OBJ = Nothing

Kill FILE_PATH_STR

XML_MAIL_DATA_FUNC = True 'Report has been sent to DST_ADDRESS_STR

Exit Function
ERROR_LABEL:
XML_MAIL_DATA_FUNC = False
End Function
