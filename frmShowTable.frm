VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmShowTable 
   Caption         =   "Data Table"
   ClientHeight    =   6225
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7650
   OleObjectBlob   =   "frmShowTable.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmShowTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'-------------------------------------------------------------------------
' Purpose   : Code to handle showing of table on this userform
'-------------------------------------------------------------------------
Option Explicit

Private mvTable As Variant
Private mbAutoColWidths As Boolean

Private mdFormWidth As Double
Private mdFormHeight As Double

'Declare an object for the clsFormResizer class to handle
'resizing for this form
Dim mclsResizer As clsFormResizer

'----------------------EVENT CODE ----------------------

Private Sub cmbClose_Click()
    Me.Hide
End Sub

Private Sub lblTableTitle_Click()

End Sub

Private Sub lbxTable_Click()

End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Comments: When the form is resized, the UserForm_Resize event
'           is raised, which we just pass on to the Resizer class
Private Sub UserForm_Resize()
    If mclsResizer Is Nothing Then Exit Sub
    mclsResizer.FormResize
End Sub
'----------------------METHODS----------------------

Public Sub Initialise()
'-------------------------------------------------------------------------
'Initialises the form and makes sure the listbox resizes
'according to the data
'-------------------------------------------------------------------------
    Dim lRowCt As Long
    Dim lColCt As Long
    Dim lLengths() As Long

    On Error GoTo ERROR_LABEL
    On Error GoTo ERROR_LABEL
    ReDim lLengths(UBound(mvTable, 2))
    With lbxTable
        .Clear
        .ColumnCount = UBound(mvTable, 2) + 1
        For lRowCt = LBound(mvTable, 1) To UBound(mvTable, 1)
            For lColCt = LBound(mvTable, 2) To UBound(mvTable, 2)
                'Store the largest string length of each column of the array
                lLengths(lColCt) = _
                Excel.Application.max(lLengths(lColCt), Len(mvTable(lRowCt, lColCt)))
                If lColCt = LBound(mvTable, 2) Then
                    'first item has to be added through additem
                    .AddItem mvTable(lRowCt, lColCt)
                Else
                    .List(.ListCount - 1, lColCt - 1) = _
                    CStr(mvTable(lRowCt, lColCt))
                End If
            Next
        Next
    End With
    If AutoColWidths Then
        'Now autosize the ColumnWidths
        SetWidths lLengths()
    End If
    
    'Create the instance of the form resizer class
    Set mclsResizer = New clsFormResizer
    'Tell it where to store the form dimensions
    mclsResizer.RegistryKey = "ShowTableOnForm"
    'Tell it which form it's handling
    Set mclsResizer.Form = Me
    
    'Temporarily disable adjusting lbxtable, it has been sized already
    lbxTable.Tag = ""
    
    'Adjust dimensions of form using new dimensions of the listbox
    'The form_resize event handles the positioning of the other
    'controls on the form
    Me.Width = lbxTable.Left + lbxTable.Width + 12
    Me.Height = lbxTable.Top + lbxTable.Height + 30 + cmbClose.Height
    
    'Enable size of listbox again
    lbxTable.Tag = "WH"
1983:
    On Error GoTo 0
    Exit Sub
ERROR_LABEL:
End Sub

Private Function SetWidths(lLengths() As Long)
'-------------------------------------------------------------------------
' Purpose   : Sets the column widths of the listbox according to an
'array of max text lengths
'-------------------------------------------------------------------------
    Dim lCt As Long
    Dim sWidths As String
    Dim dTotWidth As Double
    On Error GoTo ERROR_LABEL
    For lCt = 1 To UBound(lLengths)
        With lblHidden
            'Using repeating letter m to determine width because that
            'is an average size letter.
            'To ensure text always fits, use capital M instead
            .Caption = String(lLengths(lCt), "m")
        End With
        dTotWidth = dTotWidth + lblHidden.Width
        If Len(sWidths) = 0 Then
            sWidths = CStr(Int(lblHidden.Width) + 1)
        Else
            sWidths = sWidths & ";" & CStr(Int(lblHidden.Width) + 1)
        End If
    Next
    
    'Now set the widths of the columns
    lbxTable.ColumnWidths = sWidths
    
    'Adjust the dimensions of the listbox itself. You may want to adjust
    'the constants
    'I hard coded here.
    
    'Listbox will always be at least 200 wide
    lbxTable.Width = Excel.Application.Min(Excel.Application.max(200, dTotWidth + 12), _
    lbxTable.Width)
    
    'Listbox will always be at least 48 high.
    lbxTable.Height = _
    Excel.Application.Min(Excel.Application.max((lbxTable.ListCount + 1) * 12, 48), _
    lbxTable.Height)
1983:
    On Error GoTo 0
    Exit Function
ERROR_LABEL:
End Function

'----------------------PROPERTIES----------------------
Public Property Get Table() As Variant
    Table = mvTable
End Property

Public Property Let Table(ByVal vTable As Variant)
    mvTable = vTable
End Property

Public Property Let Title(ByVal sTitle As String)
    lblTableTitle.Caption = sTitle
End Property

Public Property Get AutoColWidths() As Boolean
    AutoColWidths = mbAutoColWidths
End Property

Public Property Let AutoColWidths(ByVal bAutoColWidths As Boolean)
    mbAutoColWidths = bAutoColWidths
End Property

Public Property Get FormWidth() As Double
    FormWidth = Me.Width
End Property

Public Property Let FormWidth(ByVal dFormWidth As Double)
    Me.Width = dFormWidth
End Property

Public Property Get FormHeight() As Double
    FormHeight = Me.Height
End Property

Public Property Let FormHeight(ByVal dFormHeight As Double)
    Me.Height = dFormHeight
End Property
