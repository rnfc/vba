VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmCharts 
   Caption         =   "eFinancial Charts"
   ClientHeight    =   3465
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   6330
   OleObjectBlob   =   "frmCharts.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmCharts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Base 1

Private PUB_NAMES_STR As String
Private PUB_SYMBOLS_STR As String
Private PUB_NO_CHARTS As Long

Private Sub CommandButton1_Click()
    Dim k As Long
    On Error Resume Next
    k = frmCharts.Preset.ListIndex + 1
    Call PRINT_WEB_FINANCIAL_CHARTS_FUNC(Range(PUB_SYMBOLS_STR), Range(PUB_NAMES_STR), k, PUB_NO_CHARTS, "")
End Sub

Private Sub CommandButton2_Click()
    frmCharts.Hide
End Sub

Private Sub RefEdit1_Change()
    On Error Resume Next
    PUB_NAMES_STR = RefEdit1.value
End Sub
Private Sub RefEdit2_Change()
    On Error Resume Next
    PUB_SYMBOLS_STR = RefEdit2.value
End Sub

Private Sub ScrollBar1_Change()
PUB_NO_CHARTS = frmCharts.ScrollBar1.value
frmCharts.Frame1.Caption = CStr(PUB_NO_CHARTS) & " Charts per Row"
End Sub

Public Sub UserForm_Initialize()
    
On Error Resume Next

frmCharts.Preset.Clear
frmCharts.Preset.AddItem "Yahoo Finance 1-Day Range"
frmCharts.Preset.AddItem "Yahoo Finance 1-Week Range"
frmCharts.Preset.AddItem "Implied Volatility"
frmCharts.Preset.AddItem "Finviz daily technical chart"
frmCharts.Preset.AddItem "Finviz intraday basic"
frmCharts.Preset.AddItem "Fred Historical Charts"

frmCharts.ScrollBar1.SetFocus
frmCharts.ScrollBar1.value = 1
Call ScrollBar1_Change
End Sub