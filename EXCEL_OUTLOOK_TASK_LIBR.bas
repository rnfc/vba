Attribute VB_Name = "EXCEL_OUTLOOK_TASK_LIBR"


'-------------------------------------------------------------------------------------
Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.
'-------------------------------------------------------------------------------------


'set up new tasks in Outlook from Excel

Function CREATE_OUTLOOK_TASK_FUNC(ByVal SUBJECT_STR As String, _
ByVal BODY_STR As String, _
ByVal REMINDER_DATE As Long)
        
Dim OUTLOOK_OBJ As New Outlook.Application

Dim SPACE_OBJ As Outlook.Namespace
Dim TASK_FOLDER_OBJ As Outlook.MAPIFolder
Dim TASK_ITEMS_OBJ As Outlook.Items
Dim TASK_OBJ As Outlook.TaskItem

On Error GoTo ERROR_LABEL
  
CREATE_OUTLOOK_TASK_FUNC = False
  
Set SPACE_OBJ = OUTLOOK_OBJ.Session
Set TASK_FOLDER_OBJ = SPACE_OBJ.GetDefaultFolder(olFolderTasks)
Set TASK_ITEMS_OBJ = TASK_FOLDER_OBJ.Items
Set TASK_OBJ = TASK_ITEMS_OBJ.Add

TASK_OBJ.DueDate = REMINDER_DATE
TASK_OBJ.ReminderTime = REMINDER_DATE + 0.3333 '8am
TASK_OBJ.body = BODY_STR
TASK_OBJ.Subject = SUBJECT_STR
TASK_OBJ.Importance = olImportanceLow
TASK_OBJ.Save

Set TASK_OBJ = Nothing
Set TASK_ITEMS_OBJ = Nothing
Set TASK_FOLDER_OBJ = Nothing
Set SPACE_OBJ = Nothing
Set OUTLOOK_OBJ = Nothing
    
CREATE_OUTLOOK_TASK_FUNC = True
    
Exit Function
ERROR_LABEL:
On Error Resume Next
Set TASK_OBJ = Nothing
Set TASK_ITEMS_OBJ = Nothing
Set TASK_FOLDER_OBJ = Nothing
Set SPACE_OBJ = Nothing
Set OUTLOOK_OBJ = Nothing
CREATE_OUTLOOK_TASK_FUNC = False
End Function
