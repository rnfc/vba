VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmMorn 
   Caption         =   "Statements Charts"
   ClientHeight    =   3105
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   5790
   OleObjectBlob   =   "frmMorn.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmMorn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Base 1

Private PUB_NAMES_STR As String
Private PUB_SYMBOLS_STR As String
Private PUB_ELEMENT_STR As String

Private Sub CommandButton1_Click()
    On Error Resume Next
    Call MORNINGSTAR_FUNDAMENTAL_CHARTS_FUNC(Range(PUB_SYMBOLS_STR), Range(PUB_NAMES_STR), _
    PUB_ELEMENT_STR, ",")
End Sub

Private Sub CommandButton2_Click()
    frmMorn.Hide
End Sub
Private Sub RefEdit1_Change()
    On Error Resume Next
    PUB_NAMES_STR = RefEdit1.value
End Sub
Private Sub RefEdit2_Change()
    On Error Resume Next
    PUB_SYMBOLS_STR = RefEdit2.value
End Sub
Private Sub TextBox2_Change()
    PUB_ELEMENT_STR = frmMorn.TextBox2.value
End Sub

Public Sub UserForm_Initialize()
On Error Resume Next
frmMorn.TextBox2.value = _
"PB,PC,PE,PS,RG,OIG,EPSG,EQG,CFO,EPS,ROEG10,ROAG10,PROA,ROEA,TOTR,CR,DE,DTC "
End Sub
