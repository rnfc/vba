VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmFinviz 
   Caption         =   "Finviz Stock Screener"
   ClientHeight    =   6975
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   4710
   OleObjectBlob   =   "frmFinviz.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmFinviz"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.

Private Sub CommandButton1_Click()
Call PRINT_FINVIZ_SCREENER_FUNC
End Sub

Private Sub CommandButton2_Click()
frmFinviz.Hide
End Sub

Private Sub CommandButton3_Click()
Call EXCEL_TURN_OFF_EVENTS_FUNC
Workbooks.Open (PUB_FINVIZ_URL_STR & PUB_FINVIZ_DATA_SUFFIX_STR)
Call EXCEL_TURN_ON_EVENTS_FUNC
End Sub

Private Sub CommandButton4_Click()
Call UserForm_Initialize
End Sub

Private Sub Label16_Click()

End Sub

Private Sub Preset_Change()
'Debt/Equity <30% --> dont require tons of debt to operate
'Current Ratio >2 --> Have adequate short term liquidity
'Quick Ratio > 2 --> Have adequate short term liquidity without their inventories
'P/B < 2 --> Trading near its liquidation value
'Trailing P/E < 15 --> Producing significant earnings relative to its marcap cap (both past and future)
'FORWARD PE < 15
'PEG < 1
'PE-PEG -are subject to downward revisions as the
'macro picture deteriorates more and more.

    If frmFinviz.Preset = "Custom Screen" Then
        frmFinviz.PEG.value = "Any"
        frmFinviz.Return_on_assets.value = "Any"
        frmFinviz.Sales_Growth.value = "Any"
        frmFinviz.PCash.value = "Any"
        frmFinviz.ROI.value = "Any"
        frmFinviz.InsiderT.value = "Any"
        frmFinviz.InsiderO.value = "Any"
        frmFinviz.Debt_Equity.value = "Any"
        frmFinviz.EPS_Growth.value = "Any"
        frmFinviz.PBook.value = "Any"
        frmFinviz.PEarnings.value = "Any"
        frmFinviz.Market_Cap.value = "Any"
    ElseIf frmFinviz.Preset = "Value Screen" Then
        frmFinviz.PEG.value = "Any"
        frmFinviz.Return_on_assets.value = "Over 5%"
        frmFinviz.Sales_Growth.value = "Over 5%"
        frmFinviz.PCash.value = "Under 5"
        frmFinviz.ROI.value = "Over 10%"
        frmFinviz.InsiderT.value = "Any"
        frmFinviz.InsiderO.value = "Over 10%"
        frmFinviz.Debt_Equity.value = "Under 0.3"
        frmFinviz.EPS_Growth.value = "Over 10%"
        frmFinviz.PBook.value = "Under 2"
        frmFinviz.PEarnings.value = "Under 10"
        frmFinviz.Market_Cap.value = "Under 10bn"
    ElseIf frmFinviz.Preset = "Growth at a Reasonable Price" Then
        frmFinviz.PEG.value = "Under 1"
        frmFinviz.Div.value = "Over 5%"
        frmFinviz.Sales_Growth.value = "Over 10%"
        frmFinviz.EPS_Growth.value = "Over 20%"
    End If
    
End Sub

Public Sub PRINT_FINVIZ_SCREENER_FUNC()

Dim SROW As Long
Dim NROWS As Long

Dim SCOLUMN As Long
Dim NCOLUMNS As Long

Dim DST_RNG As Excel.Range

Dim STR_SIZE As String
Dim STR_DE As String
Dim STR_EPSG As String
Dim STR_PB As String
Dim STR_PCASH As String
Dim STR_PE As String
Dim STR_PEG As String
Dim STR_ROA As String
Dim STR_ROI As String
Dim STR_SGROWTH As String
Dim STR_INSIDERO As String
Dim STR_INSIDERT As String
Dim STR_DIVYIELD As String
Dim STR_SUFFIX As String

Dim TEMP_MATRIX As Variant

On Error GoTo ERROR_LABEL

Call EXCEL_TURN_OFF_EVENTS_FUNC

If frmFinviz.Market_Cap.value <> "Any" Then
    If frmFinviz.Market_Cap.value = "Under 200bn" Then
        STR_SIZE = "cap_largeunder,"
    ElseIf frmFinviz.Market_Cap.value = "Under 10bn" Then
        STR_SIZE = "cap_midunder,"
    ElseIf frmFinviz.Market_Cap.value = "Under 2bn" Then
        STR_SIZE = "cap_smallunder,"
    End If
Else
    STR_SIZE = ""
End If

If frmFinviz.Debt_Equity.value <> "Any" Then
    If frmFinviz.Debt_Equity.value = "Under 0.2" Then
        STR_DE = "fa_debteq_u0.2"
    ElseIf frmFinviz.Debt_Equity.value = "Under 0.3" Then
        STR_DE = "fa_debteq_u0.3,"
    ElseIf frmFinviz.Debt_Equity.value = "Under 0.4" Then
        STR_DE = "fa_debteq_u0.4,"
    ElseIf frmFinviz.Debt_Equity.value = "Under 0.5" Then
        STR_DE = "fa_debteq_u0.5,"
    ElseIf frmFinviz.Debt_Equity.value = "Under 0.6" Then
        STR_DE = "fa_debteq_u0.6,"
    ElseIf frmFinviz.Debt_Equity.value = "Under 0.7" Then
        STR_DE = "fa_debteq_u0.7,"
    ElseIf frmFinviz.Debt_Equity.value = "Under 0.8" Then
        STR_DE = "fa_debteq_u0.8,"
    ElseIf frmFinviz.Debt_Equity.value = "Under 0.9" Then
        STR_DE = "fa_debteq_u0.9,"
    End If
Else
    STR_DE = ""
End If

If frmFinviz.EPS_Growth.value <> "Any" Then
    If frmFinviz.EPS_Growth.value = "Over 10%" Then
        STR_EPSG = "fa_eps5years_o10,"
    ElseIf frmFinviz.EPS_Growth.value = "Over 20%" Then
        STR_EPSG = "fa_eps5years_o20,"
    Else
        STR_EPSG = "fa_eps5years_o30,"
    End If
Else
STR_EPSG = ""
End If

If frmFinviz.PBook.value <> "Any" Then
    If frmFinviz.PBook.value = "Low <1" Then
        STR_PB = "fa_pb_low,"
    ElseIf frmFinviz.PBook.value = "High >5" Then
        STR_PB = "fa_pb_high,"
    ElseIf frmFinviz.PBook.value = "Under 2" Then
        STR_PB = "fa_pb_u2,"
    Else
        STR_PB = "fa_pb_u5,"
    End If
Else
STR_PB = ""
End If

If frmFinviz.PCash.value <> "Any" Then
    If frmFinviz.PCash.value = "Low <3" Then
        STR_PCASH = "fa_pc_low,"
    ElseIf frmFinviz.PCash.value = "High >50" Then
        STR_PCASH = "fa_pc_high,"
    ElseIf frmFinviz.PCash.value = "Under 4" Then
        STR_PCASH = "fa_pc_u4,"
    Else
        STR_PCASH = "fa_pc_u5,"
    End If
Else
STR_PCASH = ""
End If

If frmFinviz.PEarnings.value <> "Any" Then
    If frmFinviz.PEarnings.value = "Under 5" Then
        STR_PE = "fa_pe_u5,"
    ElseIf frmFinviz.PEarnings.value = "Under 10" Then
        STR_PE = "fa_pe_u10,"
    ElseIf frmFinviz.PEarnings.value = "Under 20" Then
        STR_PE = "fa_pe_u20,"
    Else
        STR_PE = "fa_pe_u40,"
    End If
Else
STR_PE = ""
End If

If frmFinviz.PEG.value <> "Any" Then
    If frmFinviz.PEG.value = "Under 1" Then
        STR_PEG = "fa_peg_u1,"
    ElseIf frmFinviz.PEG.value = "Under 2" Then
        STR_PEG = "fa_peg_u2,"
    Else
        STR_PEG = "fa_peg_u3,"
    End If
Else
STR_PEG = ""
End If

If frmFinviz.Return_on_assets.value <> "Any" Then
    If frmFinviz.Return_on_assets.value = "Positive >0%" Then
        STR_ROA = "fa_roa_pos,"
    ElseIf frmFinviz.Return_on_assets.value = "Over 5%" Then
        STR_ROA = "fa_roa_o5,"
    ElseIf frmFinviz.Return_on_assets.value = "Over 10%" Then
        STR_ROA = "fa_roa_o10,"
    Else
        STR_ROA = "fa_roa_o15,"
    End If
Else
STR_ROA = ""
End If

If frmFinviz.ROI.value <> "Any" Then
    If frmFinviz.ROI.value = "Positive >0%" Then
        STR_ROI = "fa_roi_pos,"
    ElseIf frmFinviz.ROI.value = "Over 5%" Then
        STR_ROI = "fa_roi_o5,"
    ElseIf frmFinviz.ROI.value = "Over 10%" Then
        STR_ROI = "fa_roi_o10,"
    Else
        STR_ROI = "fa_roi_o15,"
    End If
Else
STR_ROI = ""
End If

If frmFinviz.Sales_Growth.value <> "Any" Then
    If frmFinviz.Sales_Growth.value = "Over 5%" Then
        STR_SGROWTH = "fa_sales5years_o5,"
    ElseIf frmFinviz.Sales_Growth.value = "Over 10%" Then
        STR_SGROWTH = "fa_sales5years_o10,"
    ElseIf frmFinviz.Sales_Growth.value = "Over 15%" Then
        STR_SGROWTH = "fa_sales5years_o15,"
    Else
        STR_SGROWTH = "fa_sales5years_o20,"
    End If
Else
STR_SGROWTH = ""
End If

If frmFinviz.InsiderO.value <> "Any" Then
    If frmFinviz.InsiderO.value = "Over 10%" Then
        STR_INSIDERO = "sh_insiderown_o10,"
    ElseIf frmFinviz.InsiderO.value = "Over 20%" Then
        STR_INSIDERO = "sh_insiderown_o20,"
    ElseIf frmFinviz.InsiderO.value = "Over 30%" Then
        STR_INSIDERO = "sh_insiderown_o30,"
    ElseIf frmFinviz.InsiderO.value = "Over 40%" Then
        STR_INSIDERO = "sh_insiderown_o40,"
    Else
        STR_INSIDERO = "sh_insiderown_o50,"
    End If
Else
STR_INSIDERO = ""
End If

If frmFinviz.InsiderT.value <> "Any" Then
    If frmFinviz.InsiderT.value = "Very Negative <20%" Then
        STR_INSIDERT = "sh_insidertrans_veryneg,"
    ElseIf frmFinviz.InsiderT.value = "Negative <0%" Then
        STR_INSIDERT = "sh_insidertrans_neg,"
    ElseIf frmFinviz.InsiderT.value = "Positive >0%" Then
        STR_INSIDERT = "sh_insidertrans_pos,"
    Else
        STR_INSIDERT = "sh_insidertrans_verypos,"
    End If
Else
STR_INSIDERT = ""
End If

If frmFinviz.Div.value <> "Any" Then
    If frmFinviz.Div.value = "Over 2%" Then
        STR_DIVYIELD = "fa_div_o2,"
    ElseIf frmFinviz.Div.value = "Over 5%" Then
        STR_DIVYIELD = "fa_div_o5,"
    ElseIf frmFinviz.Div.value = "Over 10%" Then
        STR_DIVYIELD = "fa_div_o10,"
    End If
Else
STR_DIVYIELD = ""
End If

STR_SUFFIX = "v=152&f=" & STR_SIZE & STR_DE & STR_DIVYIELD & STR_EPSG & STR_PB & STR_PCASH & _
              STR_PE & STR_PEG & STR_ROA & STR_ROI & STR_SGROWTH & STR_INSIDERO & STR_INSIDERT
STR_SUFFIX = Mid(STR_SUFFIX, 1, Len(STR_SUFFIX) - 1) & PUB_FINVIZ_DATA_SUFFIX_STR

TEMP_MATRIX = FINVIZ_SCREENER_FUNC(STR_SUFFIX)
If IsArray(TEMP_MATRIX) = False Then: GoTo 1983
        
SROW = LBound(TEMP_MATRIX, 1)
NROWS = UBound(TEMP_MATRIX, 1)
            
SCOLUMN = LBound(TEMP_MATRIX, 2)
NCOLUMNS = UBound(TEMP_MATRIX, 2)
                        
Set DST_RNG = WSHEET_ADD_FUNC(PARSE_CURRENT_TIME_FUNC("_"), ActiveWorkbook).Cells(3, 3)
Set DST_RNG = Range(DST_RNG.Cells(SROW, SCOLUMN), DST_RNG.Cells(NROWS, NCOLUMNS))
DST_RNG.value = TEMP_MATRIX
GoSub FORMAT_LINE

1983:
Call EXCEL_TURN_ON_EVENTS_FUNC

Exit Sub
'-----------------------------------------------------------------------------
FORMAT_LINE:
'-----------------------------------------------------------------------------
    With DST_RNG
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
        .Borders(xlDiagonalDown).LineStyle = xlNone
        .Borders(xlDiagonalUp).LineStyle = xlNone
        With .Borders(xlEdgeLeft)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .WEIGHT = xlThin
        End With
        With .Borders(xlEdgeTop)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .WEIGHT = xlThin
        End With
        With .Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .WEIGHT = xlThin
        End With
        With .Borders(xlEdgeRight)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .WEIGHT = xlThin
        End With
        With .Borders(xlInsideVertical)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .WEIGHT = xlThin
        End With
        With .Borders(xlInsideHorizontal)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .WEIGHT = xlThin
        End With
        With .Rows(1)
            .Font.Bold = True
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .WrapText = True
            .Orientation = 0
            .AddIndent = False
            .IndentLevel = 0
            .ShrinkToFit = False
            .ReadingOrder = xlContext
            .MergeCells = False
            With .Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .Color = 65535
                .TintAndShade = 0
                .PatternTintAndShade = 0
            End With
        End With
        .ColumnWidth = 15
        .RowHeight = 15
    End With
    Return
'-----------------------------------------------------------------------------
'-----------------------------------------------------------------------------
ERROR_LABEL:
Call EXCEL_TURN_ON_EVENTS_FUNC
End Sub

Public Sub UserForm_Initialize()
    
    frmFinviz.Market_Cap.Clear
    frmFinviz.PEarnings.Clear
    frmFinviz.PBook.Clear
    frmFinviz.EPS_Growth.Clear
    frmFinviz.Debt_Equity.Clear
    frmFinviz.InsiderO.Clear
    frmFinviz.InsiderT.Clear
    frmFinviz.ROI.Clear
    frmFinviz.PCash.Clear
    frmFinviz.Sales_Growth.Clear
    frmFinviz.Return_on_assets.Clear
    frmFinviz.PEG.Clear
    frmFinviz.Preset.Clear
    
    frmFinviz.Market_Cap.AddItem "Any"
    frmFinviz.Market_Cap.AddItem "Under 200bn"
    frmFinviz.Market_Cap.AddItem "Under 10bn"
    frmFinviz.Market_Cap.AddItem "Under 2bn"
    
    frmFinviz.PEarnings.AddItem "Any"
    frmFinviz.PEarnings.AddItem "Under 5"
    frmFinviz.PEarnings.AddItem "Under 10"
    frmFinviz.PEarnings.AddItem "Under 20"
    frmFinviz.PEarnings.AddItem "Under 40"
    
    frmFinviz.PBook.AddItem "Any"
    frmFinviz.PBook.AddItem "Low <1"
    frmFinviz.PBook.AddItem "High >5"
    frmFinviz.PBook.AddItem "Under 2"
    frmFinviz.PBook.AddItem "Under 5"
    
    frmFinviz.EPS_Growth.AddItem "Any"
    frmFinviz.EPS_Growth.AddItem "Over 10%"
    frmFinviz.EPS_Growth.AddItem "Over 20%"
    frmFinviz.EPS_Growth.AddItem "Over 30%"
    
    frmFinviz.Debt_Equity.AddItem "Any"
    frmFinviz.Debt_Equity.AddItem "Under 0.2"
    frmFinviz.Debt_Equity.AddItem "Under 0.3"
    frmFinviz.Debt_Equity.AddItem "Under 0.4"
    frmFinviz.Debt_Equity.AddItem "Under 0.5"
    frmFinviz.Debt_Equity.AddItem "Under 0.6"
    frmFinviz.Debt_Equity.AddItem "Under 0.7"
    frmFinviz.Debt_Equity.AddItem "Under 0.8"
    frmFinviz.Debt_Equity.AddItem "Under 0.9"
    
    frmFinviz.InsiderO.AddItem "Any"
    frmFinviz.InsiderO.AddItem "Over 10%"
    frmFinviz.InsiderO.AddItem "Over 20%"
    frmFinviz.InsiderO.AddItem "Over 30%"
    frmFinviz.InsiderO.AddItem "Over 40%"
    frmFinviz.InsiderO.AddItem "Over 50%"
    
    frmFinviz.InsiderT.AddItem "Any"
    frmFinviz.InsiderT.AddItem "Very Negative <20%"
    frmFinviz.InsiderT.AddItem "Negative <0%"
    frmFinviz.InsiderT.AddItem "Positive >0%"
    frmFinviz.InsiderT.AddItem "Very Positive >20%"
    
    frmFinviz.ROI.AddItem "Any"
    frmFinviz.ROI.AddItem "Positive >0%"
    frmFinviz.ROI.AddItem "Over 5%"
    frmFinviz.ROI.AddItem "Over 10%"
    
    frmFinviz.PCash.AddItem "Any"
    frmFinviz.PCash.AddItem "Low <3"
    frmFinviz.PCash.AddItem "High >50"
    frmFinviz.PCash.AddItem "Under 4"
    frmFinviz.PCash.AddItem "Under 5"
    
    frmFinviz.Sales_Growth.AddItem "Any"
    frmFinviz.Sales_Growth.AddItem "Over 5%"
    frmFinviz.Sales_Growth.AddItem "Over 10%"
    frmFinviz.Sales_Growth.AddItem "Over 15%"
    frmFinviz.Sales_Growth.AddItem "Over 20%"
    
    frmFinviz.Return_on_assets.AddItem "Any"
    frmFinviz.Return_on_assets.AddItem "Positive >0%"
    frmFinviz.Return_on_assets.AddItem "Over 5%"
    frmFinviz.Return_on_assets.AddItem "Over 10%"
    frmFinviz.Return_on_assets.AddItem "Over 15%"
    
    frmFinviz.PEG.AddItem "Any"
    frmFinviz.PEG.AddItem "Under 1"
    frmFinviz.PEG.AddItem "Under 2"
    frmFinviz.PEG.AddItem "Under 3"
    
    frmFinviz.Div.AddItem "Any"
    frmFinviz.Div.AddItem "Over 2%"
    frmFinviz.Div.AddItem "Over 5%"
    frmFinviz.Div.AddItem "Over 10%"
    
    frmFinviz.Preset.AddItem "Value Screen"
    frmFinviz.Preset.AddItem "Growth at a Reasonable Price"
    frmFinviz.Preset.AddItem "Custom Screen"

    frmFinviz.PEG.value = "Any"
    frmFinviz.Return_on_assets.value = "Any"
    frmFinviz.Sales_Growth.value = "Any"
    frmFinviz.PCash.value = "Any"
    frmFinviz.ROI.value = "Any"
    frmFinviz.InsiderT.value = "Any"
    frmFinviz.InsiderO.value = "Any"
    frmFinviz.Debt_Equity.value = "Any"
    frmFinviz.EPS_Growth.value = "Any"
    frmFinviz.PBook.value = "Any"
    frmFinviz.PEarnings.value = "Any"
    frmFinviz.Market_Cap.value = "Any"
    frmFinviz.Div.value = "Any"
    
End Sub
