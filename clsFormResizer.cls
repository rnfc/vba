VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsFormResizer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'Makes a userform resizeable and handles the sizing and positioning
'of all controls, using resizing information specified in each
'control's Tag property.

Option Explicit

' **************************************************************
' Windows API Declarations and Constants Follow
' **************************************************************

'Find the userform's window handle
Private Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, _
ByVal lpWindowName As String) As Long

'Get the userform's window style
Private Declare PtrSafe Function GetWindowLong Lib "user32" Alias _
"GetWindowLongA" (ByVal hwnd As Long, _
ByVal nIndex As Long) As Long

'Set the userform's window style
Private Declare PtrSafe Function SetWindowLong Lib "user32" Alias _
"SetWindowLongA" (ByVal hwnd As Long, _
ByVal nIndex As Long, _
ByVal dwNewLong As Long) As Long

'The offset of a window's style
Private Const GWL_STYLE As Long = (-16)

'Style to add a sizable frame
Private Const WS_THICKFRAME As Long = &H40000
Private Const WS_MINIMIZEBOX As Long = &H20000
'Style to add a Minimize box on the title bar
Private Const WS_MAXIMIZEBOX As Long = &H10000
'Style to add a Maximize box to the title bar
''''''''''''''''''''''''''''''''''''''''''''''''''
' Module-level Declarations Follow
''''''''''''''''''''''''''''''''''''''''''''''''''
Dim moForm As Object
'The userform we're handling
Dim mhWndForm As Long
'The userform's window handle
Dim mdWidth As Double
'The previous width of the form
Dim mdHeight As Double
'The previous height of the form
Dim msRegKey As String
'The registry key for storing the form's size and position

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Comments: Default for the registry key to store the dimensions
Private Sub Class_Initialize()
    On Error GoTo ERROR_LABEL
    msRegKey = "ShowTableOnForm"
1983:
    Exit Sub
ERROR_LABEL:
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Comments: Properties to identify where in the registry to store
'           the userform position information
Public Property Let RegistryKey(sNew As String)
    msRegKey = sNew
End Property

Public Property Get RegistryKey() As String
    RegistryKey = msRegKey
End Property

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Comments: We're told which form to handle the resizing for,
'           set in the UserForm_Initialize event. Make the form
'           resizable and set its size and position
'
' Arguments:    oForm       The userform to handle
' --------------------------------------------------------------
Public Property Set Form(oNew As Object)

    Dim sSizes As String
    Dim iStyle As Long

    'Remember the form for later
    On Error GoTo ERROR_LABEL
    Set moForm = oNew

    'Get the userform's window handle
    If Val(Excel.Application.VERSION) < 9 Then
        mhWndForm = FindWindow("ThunderXFrame", moForm.Caption)    'XL97
    Else
        mhWndForm = FindWindow("ThunderDFrame", moForm.Caption)    'XL2000
    End If


    'Make the form resizable
    iStyle = GetWindowLong(mhWndForm, GWL_STYLE)
    iStyle = iStyle Or WS_THICKFRAME
    '    iStyle = iStyle Or WS_MINIMIZEBOX
    '    iStyle = iStyle Or WS_MAXIMIZEBOX
    SetWindowLong mhWndForm, GWL_STYLE, iStyle

    'Read its dimensions from the registry (if there)
    'The string has the form of "<Top>;<Left>;<Height>;<Width>"
    sSizes = GetSetting(msRegKey, "Forms", moForm.name, "")

    'Remember the current size for use in the Resize routine
    mdWidth = moForm.FormWidth
    mdHeight = moForm.FormHeight

1983:
    On Error GoTo 0
    Exit Property
ERROR_LABEL:
End Property


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Comments: Called from the User_Form resize event. Also triggered
'           when we change the size ourself
' --------------------------------------------------------------
Public Sub FormResize()

    Dim dWidthAdj As Double, dHeightAdj As Double
    Dim bSomeWidthChange As Boolean
    Dim bSomeHeightChange As Boolean
    Dim sTag As String
    Dim oCtl As MSForms.Control

    Static bResizing As Boolean

    On Error GoTo ERROR_LABEL

    'Resizing can be triggered from within this routine,
    'so use a flag to prevent recursion
    If bResizing Then Exit Sub
    If moForm.InsideHeight = 0 Then Exit Sub
    bResizing = True

    'Calculate the change in height and width
    dHeightAdj = moForm.FormHeight - mdHeight
    dWidthAdj = moForm.FormWidth - mdWidth

    'Check if we can perform the adjustment
    '(i.e. widths and heights can't be negative)
    For Each oCtl In moForm.Controls

        'Read the control's Tag property, which contains the resizing info
        sTag = UCase(oCtl.Tag)

        'If we're changing the Top, check that it won't move off the top
        'of the form
        If InStr(1, sTag, "T", vbBinaryCompare) Then
            If oCtl.Top + dHeightAdj * ResizeFactor(sTag, "T") <= 0 Then
                moForm.FormHeight = mdHeight
            End If

            bSomeHeightChange = True
        End If

        'If we're changing the Left, check that it won't move off the
        'left of the form
        If InStr(1, sTag, "L", vbBinaryCompare) Then
            If oCtl.Left + dWidthAdj * ResizeFactor(sTag, "L") <= 0 Then
                moForm.Width = mdWidth
            End If

            bSomeWidthChange = True
        End If

        'If we're changing the Height, check that it won't go negative
        If InStr(1, sTag, "H", vbBinaryCompare) Then
            If oCtl.Height + dHeightAdj * ResizeFactor(sTag, "H") <= 0 Then
                moForm.Height = mdHeight
            End If

            bSomeHeightChange = True
        End If

        'If we're changing the Width, check that it won't go negative
        If InStr(1, sTag, "W", vbBinaryCompare) Then
            If oCtl.Width + dWidthAdj * ResizeFactor(sTag, "W") <= 0 Then
                moForm.Width = mdWidth
            End If

            bSomeWidthChange = True
        End If
    Next    'Control

    'If none of the controls move or size,
    'don't allow the form to resize in that direction
    If Not bSomeHeightChange Then moForm.FormHeight = mdHeight
    If Not bSomeWidthChange Then moForm.Width = mdWidth

    'Recalculate the height and width changes,
    'in case the previous checks reset them
    dHeightAdj = moForm.FormHeight - mdHeight
    dWidthAdj = moForm.Width - mdWidth

    'Loop through all the controls on the form,
    'adjusting their position and size
    For Each oCtl In moForm.Controls
        With oCtl
            sTag = UCase(.Tag)

            'Changing the Top
            If InStr(1, sTag, "T", vbBinaryCompare) Then
                .Top = .Top + dHeightAdj * ResizeFactor(sTag, "T")
            End If

            'Changing the Left
            If InStr(1, sTag, "L", vbBinaryCompare) Then
                .Left = .Left + dWidthAdj * ResizeFactor(sTag, "L")
            End If

            'Changing the Height
            If InStr(1, sTag, "H", vbBinaryCompare) Then
                .Height = .Height + dHeightAdj * ResizeFactor(sTag, "H")
            End If

            'Changing the Width
            If InStr(1, sTag, "W", vbBinaryCompare) Then
                .Width = .Width + dWidthAdj * ResizeFactor(sTag, "W")
            End If
        End With
    Next    'Control

    'Remember the new dimensions of the form for next time
    mdWidth = moForm.Width
    mdHeight = moForm.FormHeight
    'Store the size and position in the registry
    With moForm
        SaveSetting msRegKey, "Forms", .name, Str(.Top) & ";" & _
                                              Str(.Left) & ";" & _
                                              Str(.Height) & ";" & Str(.Width)
    End With

    'Reset the recursion flag, now that we're done
    bResizing = False

    On Error GoTo 0
1983:
    Exit Sub
ERROR_LABEL:
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Comments: Function to locate a property letter (T, L, H or W)
'           in the Tag string and return the resizing factor for it
' Arguments:    sTag        The full text of the control's Tag string
'               sChange     The resize letter to look for (T, L, H or W)
' Returns:      The percentage change factor
' --------------------------------------------------------------
Private Function ResizeFactor(sTag As String, sChange As String)

    Dim ipos As Long, dfactor As Double

    'Locate the property letter in the tag string
    On Error GoTo ERROR_LABEL
    ipos = InStr(1, sTag, sChange, vbBinaryCompare)

    'If we found it...
    If ipos > 0 Then

        '... read the number following it
        dfactor = Val(Mid$(sTag, ipos + 1))

        'If there was no number, use a factor of 100%
        If dfactor = 0 Then dfactor = 1
    End If

    'Return the factor
    ResizeFactor = dfactor
1983:
    On Error GoTo 0
    Exit Function
ERROR_LABEL:
End Function


