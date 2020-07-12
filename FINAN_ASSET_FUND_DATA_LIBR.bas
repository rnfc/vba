Attribute VB_Name = "FINAN_ASSET_FUND_DATA_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.

'Written by Nico - Ben Graham Centre for Value Investing

Function MUTUAL_FUNDS_DATABASE_FUNC(ByRef TICKERS_RNG As Variant)

Dim j As Long
Dim i As Long
Dim k As Long

Dim ii As Long
Dim jj As Long

Dim NROWS As Long
Dim NCOLUMNS As Long

Dim TEMP_VAL As Variant
Dim ELEMENT_STR As String
Dim TEMP_MATRIX As Variant
Dim TICKERS_VECTOR As Variant

Const ERROR_STR As String = "Error"
On Error GoTo ERROR_LABEL

GoSub ELEMENTS_LINE
If IsArray(TICKERS_RNG) Then
    TICKERS_VECTOR = TICKERS_RNG
    If UBound(TICKERS_VECTOR, 1) = 1 Then
        TICKERS_VECTOR = MATRIX_TRANSPOSE_FUNC(TICKERS_VECTOR)
    End If
Else
    ReDim TICKERS_VECTOR(1 To 1, 1 To 1)
    TICKERS_VECTOR(1, 1) = TICKERS_RNG
End If
NROWS = UBound(TICKERS_VECTOR, 1)
ReDim TEMP_MATRIX(0 To NROWS, 1 To NCOLUMNS + 1)
TEMP_MATRIX(0, 1) = "TICKERS"
For i = 1 To NROWS: TEMP_MATRIX(i, 1) = Trim(TICKERS_VECTOR(i, 1)): Next i
ii = 1
For j = 1 To NCOLUMNS
    jj = InStr(ii, ELEMENT_STR, ",")
    k = Val(Mid(ELEMENT_STR, ii, jj - ii))
    If k = 0 Then: GoTo 1983
    TEMP_MATRIX(0, j + 1) = RETRIEVE_WEB_DATA_ELEMENT_FUNC("ELEMENT", k, ERROR_STR)
    For i = 1 To NROWS
        TEMP_VAL = RETRIEVE_WEB_DATA_ELEMENT_FUNC(TEMP_MATRIX(i, 1), k, ERROR_STR)
        If TEMP_VAL <> ERROR_STR Then
            TEMP_MATRIX(i, j + 1) = Left(TEMP_VAL, 255)
        Else
            TEMP_MATRIX(i, j + 1) = "--"
        End If
    Next i
    'TEMP_MATRIX(1, j + 1) = k
1983:
    ii = jj + 1
Next j
MUTUAL_FUNDS_DATABASE_FUNC = CONVERT_STRING_NUMBER_FUNC(TEMP_MATRIX)

Exit Function
'------------------------------------------------------------------------------------------------------------------------------------
ELEMENTS_LINE:
'------------------------------------------------------------------------------------------------------------------------------------
ELEMENT_STR = _
"4669,4670,4671,4672,4673,4674,4675,4676,4677,4678,4679,4680,4681,4682,4683,4684,4685,4686,4687,4688,4689,4690,4691,4692,4693," & _
"4694,4695,4696,4697,4698,4699,4700,4701,4702,4703,4704,4705,4706,4707,4708,4709,4710,4711,4712,4713,4714,4715,4716,4717,4718," & _
"4719,4720,4721,4722,4723,4724,4725,4726,4727,4728,4729,4730,4731,4732,4733,4734,4735,4736,4737,4738,4739,4740,4741,4742,4743," & _
"4744,4745,4746,4747,4748,4749,4750,4751,4752,4753,4754,4755,4756,4757,4758,4759,4760,4761,4762,4763,4764,4765,4766,4767,4768," & _
"4769,4770,4771,4772,4773,4774,4775,4776,4777,4778,4779,4780,4781,4782,4783,4784,4785,4786,4787,4788,4789,4790,4791,4792,4793," & _
"4794,4795,4796,4797,4798,4799,4800,4801,4802,4803,4804,4805,4806,4807,4808,4809,4810,4811,4812,4813,4814,4815,4816,4817,4818," & _
"4819,4820,4821,4822,4823,4824,4825,4826,4827,4828,4829,4830,4832,4833,4834,4835,4836,4837,4838,4840,4841,4842,4843,4844,4845," & _
"4846,4848,4849,4850,4851,4852,4853,4854,4856,4857,4858,4859,4860,4861,4862,4863,4864,4865,4866,4867,4868,4869,4870,4871,4872," & _
"4873,4874,4875,4876,4877,4878,4879,4880,4881,4882,4883,4884,4885,4886,4887,4888,4889,4890,4891,4892,4893,4894,4895,4896,4897," & _
"4898,4899,4900,4901,4902,4903,4904,4905,4905,4906,4907,4908,4909,4910,4911,4912,4913,4915,4916,4917,4918,4919,4920,4921,4922," & _
"4923,4930,4931,4932,4933,4934,4935,4936,4937,4938,4939,4940,4941,4942,4943,4944,4945,4946,4947,4948,4949,4950,4951,4952,4953," & _
"4954,4955,4956,4957,4958,4959,4960,4961,4962,4963,4964,4965,4966,4967,4968,4969,4970,4971,4972,4973,4974,4975,4976,4977,4978," & _
"4979,4980,4981,4982,4983,4984,4985,4986,4987,4988,4989,4990,4991,4992,4993,4994,4995,4996,4997,4998,4999,5000,5001,5002,5003," & _
"5004,5005,5006,5007,5008,5009,5010,5011,5012,5013,5014,5015,5016,5017,5018,5019,5020,5021,5022,5023,5024,5025,5026,5027,5028," & _
"5029,5030,5031,5032,5033,5034,5035,5036,5037,5038,5039,5040,5041,5042,5043,5044,5045,5046,5047,5048,5049,5050,5051,5052,5053," & _
"5054,5055,5056,5057,5058,5059,5060,5061,5062,5063,5064,5065,5066,5067,5068,5069,5070,5071,5072,5073,5074,5075,5076,5077,5078," & _
"5079,5080,5081,5082,5083,5084,5085,5086,5087,5088,5089,5090,5091,5092,5093,5094,5095,5096,5097,5098,5099,5100,5101,5102,5103," & _
"5104,5105,5106,5107,5108,5109,5110,5111,5112,5113,5114,5115,5116,5117,5118,5119,5120,5121,5122,5123,5124,5125,5126,5127,5128," & _
"5129,5130,5131,5132,5133,5134,5135,5136,5137,5138,5139,5140,5141,5142,5143,5144,5145,5146,5147,5148,5149,5150,5151,5152,5153," & _
"5154,5155,5156,5157,5158,5159,5160,5161,5162,5163,5164,5165,5166,5167,5168,5169,5170,5171,5172,5173,5174,5175,5176,5177,5178," & _
"5179,5180,5181,5182,5183,5184,5185,5186,5187,5188,5189,5190,5191,5192,5193,5194,5195," '517
NCOLUMNS = 0
i = Len(ELEMENT_STR)
For j = 1 To i
    If Mid(ELEMENT_STR, j, 1) = "," Then: NCOLUMNS = NCOLUMNS + 1
Next j
'------------------------------------------------------------------------------------------------------------------------------------
Return
'------------------------------------------------------------------------------------------------------------------------------------
ERROR_LABEL:
MUTUAL_FUNDS_DATABASE_FUNC = ERROR_STR
End Function

Sub PRINT_MUTUAL_FUNDS_DATABASE_FUNC()

Dim SROW As Long
Dim SCOLUMN As Long

Dim NROWS As Long
Dim NCOLUMNS As Long

Dim DATA_RNG As Excel.Range
Dim DST_RNG As Excel.Range
Dim TEMP_RNG As Excel.Range
Dim TEMP_MATRIX As Variant

On Error GoTo ERROR_LABEL

Set DATA_RNG = Excel.Application.InputBox("Symbols", "Mutual Funds Fundamentals", , , , , , 8)
If DATA_RNG Is Nothing Then: Exit Sub

Call EXCEL_TURN_OFF_EVENTS_FUNC

Set DST_RNG = _
WSHEET_ADD_FUNC(PARSE_CURRENT_TIME_FUNC("_"), ActiveWorkbook).Cells(3, 3)

TEMP_MATRIX = MUTUAL_FUNDS_DATABASE_FUNC(DATA_RNG)
If IsArray(TEMP_MATRIX) = False Then: GoTo 1983
        
SROW = LBound(TEMP_MATRIX, 1)
NROWS = UBound(TEMP_MATRIX, 1)
            
SCOLUMN = LBound(TEMP_MATRIX, 2)
NCOLUMNS = UBound(TEMP_MATRIX, 2)

Set TEMP_RNG = Range(DST_RNG.Cells(SROW, SCOLUMN), DST_RNG.Cells(NROWS, NCOLUMNS))
TEMP_RNG.value = TEMP_MATRIX
GoSub FORMAT_LINE

1983:
Call EXCEL_TURN_ON_EVENTS_FUNC

Exit Sub
'-----------------------------------------------------------------------------
FORMAT_LINE:
'-----------------------------------------------------------------------------
    With TEMP_RNG
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
        With Union(.Columns(1), .Rows(1))
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
