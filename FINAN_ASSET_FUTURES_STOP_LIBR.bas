Attribute VB_Name = "FINAN_ASSET_FUTURES_STOP_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.

'Evolution of a Futures with a Long /Short strategy

Function FUTURES_STOP_LOSS_RETURNS_FUNC(ByRef DATA_RNG As Variant, _
Optional ByVal STOP_LOSS_PERCENT As Double = 0.01, _
Optional ByVal SHORT_FLAG As Boolean = False)

'Decide whether you want to use a daily Long strategy … or Short.
'Pick a Stop - Your STOP

'Hi Op: High/Open - 1 --> Buy at Open, Sell at High
'Lo Op: Low/Open - 1 --> Buy at Open, Sell at Low
'Op Cl: Open/Close-1 --> Buy at Close, Sell at Open
'Cl Op: Close/Open-1 --> Buy at Open, Sell at Close

Dim i As Long
Dim NROWS As Long
Dim NCOLUMNS As Long
Dim DATA_MATRIX As Variant
Dim TEMP_VECTOR As Variant

On Error GoTo ERROR_LABEL

DATA_MATRIX = DATA_RNG
NROWS = UBound(DATA_MATRIX, 1)
NCOLUMNS = UBound(DATA_MATRIX, 2) 'Date/Open/High/Low/Close
'--------------------------------------------------------------------------------
ReDim TEMP_VECTOR(1 To NROWS, 1 To 1)
'--------------------------------------------------------------------------------
If SHORT_FLAG = False Then 'For a Long strategy, you BUY at the Open and SELL during the day.
'If the stock drops by more than 1% (that is: Low/Open - 1 < -1%), you Sell with just a 1% loss.
'If the stock drops by less than 1% (that is: Low/Open - 1 > -1%), then you Sell at the Close, with a return equal to Close/Open - 1.
'There is a limit of 1% on your losses but no limit on gains … especially if the stock price goes UP!
'--------------------------------------------------------------------------------
    For i = 1 To NROWS
        If (DATA_MATRIX(i, 4) / DATA_MATRIX(i, 2) - 1) < -STOP_LOSS_PERCENT Then
            TEMP_VECTOR(i, 1) = -STOP_LOSS_PERCENT
        Else
            TEMP_VECTOR(i, 1) = (DATA_MATRIX(i, 5) / DATA_MATRIX(i, 2) - 1)
        End If
    Next i
'--------------------------------------------------------------------------------
Else 'For a Short strategy, you SELL at the Open and BUY during the day.
'If the stock increases by more than 1% (that is: High/Open - 1 > 1%), you Sell with just a 1% loss.
'If the stock increases by less than 1% (that is: High/Open - 1 < 1%), you Sell at the Close, with a return equal to 1 - Close/Open.
'There is a limit of 1% on your losses but no limit on gains … especially if the stock price goes DOWN!
'--------------------------------------------------------------------------------
    For i = 1 To NROWS
        If (DATA_MATRIX(i, 3) / DATA_MATRIX(i, 2) - 1) > STOP_LOSS_PERCENT Then
            TEMP_VECTOR(i, 1) = -STOP_LOSS_PERCENT
        Else
            TEMP_VECTOR(i, 1) = (DATA_MATRIX(i, 2) / DATA_MATRIX(i, 5) - 1)
        End If
    Next i
'--------------------------------------------------------------------------------
End If
'--------------------------------------------------------------------------------

FUTURES_STOP_LOSS_RETURNS_FUNC = TEMP_VECTOR

Exit Function
ERROR_LABEL:
FUTURES_STOP_LOSS_RETURNS_FUNC = Err.number
End Function


