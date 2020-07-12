Attribute VB_Name = "FINAN_ASSET_TA_ADX_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.
                    

Function ASSET_TA_VDX_ADX_FUNC(ByVal TICKER_STR As Variant, _
Optional ByVal START_DATE As Date, _
Optional ByVal END_DATE As Date, _
Optional ByVal MA1_PERIOD As Long = 15, _
Optional ByVal VOLUME_WEIGHTED_FLAG As Boolean = True)

Dim i As Long
Dim j As Long

Dim NROWS As Long
Dim NCOLUMNS As Long

Dim ALPHA_VAL As Double
Dim VOLUME_VAL As Double

Dim DATA_MATRIX As Variant
Dim TEMP_MATRIX As Variant

On Error GoTo ERROR_LABEL

If IsArray(TICKER_STR) = False Then
    DATA_MATRIX = YAHOO_HISTORICAL_DATA_SERIE_FUNC(TICKER_STR, START_DATE, END_DATE, _
                  "d", "DOHLCVA", False, True, True)
Else
    DATA_MATRIX = TICKER_STR
End If

NROWS = UBound(DATA_MATRIX, 1)
NCOLUMNS = UBound(DATA_MATRIX, 2)

ALPHA_VAL = 1 - 2 / (MA1_PERIOD + 1)
'-----------------------------------------------------------------------------
ReDim TEMP_MATRIX(0 To NROWS, 1 To 16)
'-----------------------------------------------------------------------------
TEMP_MATRIX(0, 1) = "DATE"
TEMP_MATRIX(0, 2) = "OPEN"
TEMP_MATRIX(0, 3) = "HIGH"
TEMP_MATRIX(0, 4) = "LOW"
TEMP_MATRIX(0, 5) = "CLOSE"
TEMP_MATRIX(0, 6) = "VOLUME"
TEMP_MATRIX(0, 7) = "HI_DIFF"
TEMP_MATRIX(0, 8) = "LOW_DIFF"
TEMP_MATRIX(0, 9) = "BULL_PTS"
TEMP_MATRIX(0, 10) = "BEAR_PTS"
TEMP_MATRIX(0, 11) = "NUMBER_BULLS"
TEMP_MATRIX(0, 12) = "DEN"
TEMP_MATRIX(0, 13) = "NUMBER_BEAR"
TEMP_MATRIX(0, 14) = "DMI+"
TEMP_MATRIX(0, 15) = "DMI-"
TEMP_MATRIX(0, 16) = IIf(VOLUME_WEIGHTED_FLAG = True, _
                     "VDX = {(VDI+)-(VDI-)}/{(VDI+)+(VDI-)}", _
                     "ADX = {(DMI+)-(DMI-)}/{(DMI+)+(DMI-)}")

i = 1
For j = 1 To NCOLUMNS: TEMP_MATRIX(i, j) = DATA_MATRIX(i, j): Next j
TEMP_MATRIX(i, 7) = ""
TEMP_MATRIX(i, 8) = ""
TEMP_MATRIX(i, 9) = 0
TEMP_MATRIX(i, 10) = 0

If VOLUME_WEIGHTED_FLAG = True Then
    TEMP_MATRIX(i, 12) = (1 - ALPHA_VAL) * TEMP_MATRIX(i, 6)
    VOLUME_VAL = TEMP_MATRIX(i, 6)
Else
    TEMP_MATRIX(i, 12) = 1
    VOLUME_VAL = 1
End If

TEMP_MATRIX(i, 11) = (1 - ALPHA_VAL) * TEMP_MATRIX(i, 9) * VOLUME_VAL
TEMP_MATRIX(i, 13) = (1 - ALPHA_VAL) * TEMP_MATRIX(i, 10) * VOLUME_VAL

TEMP_MATRIX(i, 14) = TEMP_MATRIX(i, 11) / TEMP_MATRIX(i, 12)
TEMP_MATRIX(i, 15) = TEMP_MATRIX(i, 13) / TEMP_MATRIX(i, 12)
TEMP_MATRIX(i, 16) = TEMP_MATRIX(i, 14) - TEMP_MATRIX(i, 15)

For i = 2 To NROWS
    For j = 1 To NCOLUMNS: TEMP_MATRIX(i, j) = DATA_MATRIX(i, j): Next j
    TEMP_MATRIX(i, 7) = TEMP_MATRIX(i, 3) - TEMP_MATRIX(i - 1, 3)
    TEMP_MATRIX(i, 8) = TEMP_MATRIX(i - 1, 4) - TEMP_MATRIX(i, 4)
    
    If TEMP_MATRIX(i, 7) >= TEMP_MATRIX(i, 8) Then
        TEMP_MATRIX(i, 9) = TEMP_MATRIX(i, 7)
    Else
        TEMP_MATRIX(i, 9) = 0
    End If
    If TEMP_MATRIX(i, 9) < 0 Then: TEMP_MATRIX(i, 9) = 0
    
    If TEMP_MATRIX(i, 8) > TEMP_MATRIX(i, 7) Then
        TEMP_MATRIX(i, 10) = TEMP_MATRIX(i, 8)
    Else
        TEMP_MATRIX(i, 10) = 0
    End If
    If TEMP_MATRIX(i, 10) < 0 Then: TEMP_MATRIX(i, 10) = 0
        
    If VOLUME_WEIGHTED_FLAG = True Then
        TEMP_MATRIX(i, 12) = ALPHA_VAL * TEMP_MATRIX(i - 1, 12) + (1 - ALPHA_VAL) * TEMP_MATRIX(i, 6)
        VOLUME_VAL = TEMP_MATRIX(i - 1, 6)
    Else
        TEMP_MATRIX(i, 12) = 1
        VOLUME_VAL = 1
    End If
    TEMP_MATRIX(i, 11) = ALPHA_VAL * TEMP_MATRIX(i - 1, 11) + (1 - ALPHA_VAL) * TEMP_MATRIX(i, 9) * VOLUME_VAL
    TEMP_MATRIX(i, 13) = ALPHA_VAL * TEMP_MATRIX(i - 1, 13) + (1 - ALPHA_VAL) * TEMP_MATRIX(i, 10) * VOLUME_VAL
    TEMP_MATRIX(i, 14) = TEMP_MATRIX(i, 11) / TEMP_MATRIX(i, 12)
    TEMP_MATRIX(i, 15) = TEMP_MATRIX(i, 13) / TEMP_MATRIX(i, 12)
    TEMP_MATRIX(i, 16) = (TEMP_MATRIX(i, 14) - TEMP_MATRIX(i, 15)) / (TEMP_MATRIX(i, 14) + TEMP_MATRIX(i, 15))
                        '--> Trigger Buy > .30 ; Sell < -.3 else Hold
Next i


ASSET_TA_VDX_ADX_FUNC = TEMP_MATRIX

Exit Function
ERROR_LABEL:
ASSET_TA_VDX_ADX_FUNC = Err.number
End Function


