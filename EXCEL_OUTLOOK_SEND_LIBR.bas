Attribute VB_Name = "EXCEL_OUTLOOK_SEND_LIBR"

'-----------------------------------------------------------------------------------
Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.
'-----------------------------------------------------------------------------------
Private Declare PtrSafe Function GetDesktopWindow Lib "user32" () As Long

Private Declare PtrSafe Function ShellExecute Lib "shell32.dll" _
   Alias "ShellExecuteA" _
  (ByVal hwnd As Long, _
   ByVal lpOperation As String, _
   ByVal lpFile As String, _
   ByVal lpParameters As String, _
   ByVal lpDirectory As String, _
   ByVal nShowCmd As Long) As Long
    
Private Const SW_SHOWNORMAL As Long = 1
Private Const SW_SHOWMAXIMIZED As Long = 3
Private Const SW_SHOWDEFAULT As Long = 10

Private Type EmailInfo
   sAddrTo As String
   sAddrCC As String
   sAddrBCC As String
   sAddrFrom As String
   sSubject As String
   sMessage As String
   sPriority As Long
End Type

'************************************************************************************
'************************************************************************************
'FUNCTION      : OUTLOOK_SEND_ATTACHMENT_FUNC
'DESCRIPTION   : Send Large Emails in Microsoft Outlook
'LIBRARY       : OUTLOOK
'GROUP         : SEND
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'************************************************************************************
'************************************************************************************

Function OUTLOOK_SEND_ATTACHMENT_FUNC(ByVal TO_STR_NAME As String, _
ByVal SUBJECT_STR_NAME As String, _
ByVal TEXT_MSG As String, _
ByVal FILE_PATH_STR As String, _
Optional ByVal CC_STR_NAME As String = "", _
Optional ByVal BCC_STR_NAME As String = "")

'TEMP_MSG = ""
'TEMP_MSG = TEMP_MSG & "Nicotico "
'TEMP_MSG = TEMP_MSG & "-" & vbCrLf & vbCrLf
   
'Debug.Print OUTLOOK_SEND_EXPRESS_MAIL_FUNC("rnfermincota@gmail.com", _
"Re:", TEMP_MSG)

Dim OUTLOOK_OBJ As Outlook.Application
Dim OUTLOOK_MSG_OBJ As Outlook.MailItem
'Dim OUTLOOK_RECIPIENT_OBJ As Outlook.Recipient
'Dim OUTLOOK_ATTACHMENT_OBJ As Outlook.Attachment

On Error GoTo ERROR_LABEL

OUTLOOK_SEND_ATTACHMENT_FUNC = False

Set OUTLOOK_OBJ = CreateObject("Outlook.Application")
Set OUTLOOK_MSG_OBJ = OUTLOOK_OBJ.CreateItem(olMailItem)

If OUTLOOK_VALIDATE_EMAIL_ADDRESS_FUNC(TO_STR_NAME) = False Then: GoTo ERROR_LABEL

With OUTLOOK_MSG_OBJ
    .To = TO_STR_NAME
    .CC = CC_STR_NAME
    .BCC = BCC_STR_NAME
    .Subject = SUBJECT_STR_NAME
    .body = TEXT_MSG
    .Attachments.Add (FILE_PATH_STR)
    .ReadReceiptRequested = True
    .display
    .send
End With

Set OUTLOOK_OBJ = Nothing
Set OUTLOOK_MSG_OBJ = Nothing


OUTLOOK_SEND_ATTACHMENT_FUNC = True

Exit Function
ERROR_LABEL:
On Error Resume Next
Set OUTLOOK_MSG_OBJ = Nothing
Set OUTLOOK_OBJ = Nothing
OUTLOOK_SEND_ATTACHMENT_FUNC = False
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : OUTLOOK_SEND_EXPRESS_MAIL_FUNC
'DESCRIPTION   : Send Large Emails in Microsoft Outlook
'LIBRARY       : OUTLOOK
'GROUP         : EMAIL
'ID            : 002

'AUTHOR        : RAFAEL NICOLAS FERMIN COTA

'************************************************************************************
'************************************************************************************

Function OUTLOOK_SEND_EXPRESS_MAIL_FUNC(ByVal FROM_STR_NAME As String, _
ByVal TO_STR_NAME As String, _
ByVal SUBJECT_STR_NAME As String, _
ByVal TEXT_MSG As String, _
Optional ByVal CC_STR_NAME As String = "", _
Optional ByVal BCC_STR_NAME As String = "", _
Optional ByVal PRIOR_INDEX As Long = 2, _
Optional ByRef SRC_PATH_NAME As Variant = "")

'PRIOR_INDEX: 2 --> High; 3 --> Mid; 4 --> Low Priority

'   TEMP_MSG = ""
'   TEMP_MSG = TEMP_MSG & "Nicotico "
'   TEMP_MSG = TEMP_MSG & "-" & vbCrLf & vbCrLf
   
'   Debug.Print OUTLOOK_SEND_EXPRESS_MAIL_FUNC("rnfermincota@gmail.com", _
    "rafael_nicolas@hotmail.com", _
    "Re:", TEMP_MSG)
      
   Dim ii As Long
   Dim jj As Long
   Dim kk As Long
   Dim EMAIL_OBJ As EmailInfo
   
   On Error GoTo ERROR_LABEL
   

   OUTLOOK_SEND_EXPRESS_MAIL_FUNC = False
   
   If OUTLOOK_VALIDATE_EMAIL_ADDRESS_FUNC(FROM_STR_NAME) = False Then: GoTo ERROR_LABEL
   If OUTLOOK_VALIDATE_EMAIL_ADDRESS_FUNC(TO_STR_NAME) = False Then: GoTo ERROR_LABEL
   
  'the temp email file
   If SRC_PATH_NAME = "" Then
        SRC_PATH_NAME = Excel.Application.Path & _
        Excel.Application.PathSeparator & "temp.eml"
   End If
  
  'complete the fields to be used
   With EMAIL_OBJ
      .sAddrFrom = FROM_STR_NAME
      .sAddrTo = TO_STR_NAME
      .sAddrCC = CC_STR_NAME
      .sAddrBCC = BCC_STR_NAME
      .sSubject = SUBJECT_STR_NAME
      .sMessage = TEXT_MSG
      .sPriority = PRIOR_INDEX
   End With
   
  'create the temp file
   ii = FreeFile
   Open SRC_PATH_NAME For Output As #ii
   
  'if successful,
   If ii <> 0 Then
   
     'write out the data and
     'send the email
      If OUTLOOK_WRITE_BODY_FUNC(ii, EMAIL_OBJ) Then
      
      'the desktop will be the
      'default for error messages
      'execute the passed operation
      
          'the desktop will be the default for error messages
           kk = GetDesktopWindow()
          'execute the passed operation
           jj = ShellExecute(kk, "Open", _
                    SRC_PATH_NAME, vbNullString, _
                    vbNullString, vbNormalFocus)
        
          'This is optional. Uncomment the three lines
          'below to have the "Open With.." dialog appear
          'when the ShellExecute API call fails
        '  If success < 32 Then
        '     Call Shell("rundll32.exe shell32.dll,OpenAs_RunDLL " & _
                         SRC_PATH_NAME, vbNormalFocus)
        '  End If
      
      End If
   
   End If

OUTLOOK_SEND_EXPRESS_MAIL_FUNC = True

Exit Function
ERROR_LABEL:
OUTLOOK_SEND_EXPRESS_MAIL_FUNC = False
End Function

'This routine shows how to send Outlook mail from Excel.
'This could be useful for error reporting, for feedback,
'for any form of interaction with users. Try it out by
'specifying who to send the message to, and the message

Public Function OUTLOOK_SEND_MAIL_FUNC(ByVal RECIPIENT_ADDRESS_STR As String, _
ByVal SUBJECT_STR As String, _
ByVal BODY_STR As String, _
Optional ByVal SHOW_FLAG As Boolean = False)

Dim MAIL_OBJ As Outlook.MailItem
Dim OUTLOOK_OBJ As New Outlook.Application

On Error GoTo ERROR_LABEL

OUTLOOK_SEND_MAIL_FUNC = False
'in case you need current user name..
'Dim SPACE_OBJ As Outlook.NameSpace
'Set SPACE_OBJ = OUTLOOK_OBJ.Session
'Debug.Print SPACE_OBJ.CurrentUser.Name

Set MAIL_OBJ = OUTLOOK_OBJ.CreateItem(olMailItem)
With MAIL_OBJ.Recipients.Add(RECIPIENT_ADDRESS_STR)
  .Type = olTo
  .Resolve 'check address is valid
  If Not .Resolved Then: GoTo ERROR_LABEL 'This email address does not
  'appear to be valid
End With

MAIL_OBJ.Subject = SUBJECT_STR
MAIL_OBJ.body = BODY_STR
If SHOW_FLAG = True Then
  MAIL_OBJ.display
Else
  MAIL_OBJ.send
End If

Set MAIL_OBJ = Nothing
Set OUTLOOK_OBJ = Nothing

OUTLOOK_SEND_MAIL_FUNC = True

Exit Function
ERROR_LABEL:
On Error Resume Next
Set MAIL_OBJ = Nothing
Set OUTLOOK_OBJ = Nothing
OUTLOOK_SEND_MAIL_FUNC = False
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : OUTLOOK_WRITE_BODY_FUNC
'DESCRIPTION   : Script for writing emails
'LIBRARY       : OUTLOOK
'GROUP         : BODY
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA

'************************************************************************************
'************************************************************************************

Function OUTLOOK_WRITE_BODY_FUNC(ByVal ii As Long, _
ByRef EMAIL_OBJ As EmailInfo) As Boolean

  'write the EMAIL_OBJ fields to the file
   
   On Error GoTo ERROR_LABEL
   
   Print #ii, "From: <"; EMAIL_OBJ.sAddrFrom; ">"
   
   If Len(EMAIL_OBJ.sAddrTo) Then
      Print #ii, "To: "; Chr$(34); EMAIL_OBJ.sAddrTo; Chr$(34)
   Else
     'no to address, so bail
      OUTLOOK_WRITE_BODY_FUNC = False
      Exit Function
   End If
   
   If Len(EMAIL_OBJ.sAddrCC) Then
      Print #ii, "CC: "; Chr$(34); EMAIL_OBJ.sAddrCC; Chr$(34)
   End If
      
   If Len(EMAIL_OBJ.sAddrBCC) Then
      Print #ii, "BCC: "; Chr$(34); EMAIL_OBJ.sAddrBCC; Chr$(34)
   End If
   
   Print #ii, "Subject: "; EMAIL_OBJ.sSubject
   Print #ii, "X-Priority:"; EMAIL_OBJ.sPriority   '1=high,3=normal,5=low
   
  'this is the last header line - everything
  'after this appears in the message.
   Print #ii, "X-Unsent: 1"
   Print #ii, ""    'Blank line
   Print #ii, EMAIL_OBJ.sMessage
   
   Close #ii
   
   OUTLOOK_WRITE_BODY_FUNC = True

Exit Function
ERROR_LABEL:
OUTLOOK_WRITE_BODY_FUNC = False
End Function


