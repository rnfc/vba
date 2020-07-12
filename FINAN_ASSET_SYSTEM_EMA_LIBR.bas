Attribute VB_Name = "FINAN_ASSET_SYSTEM_EMA_LIBR"

'---------------------------------------------------------------------------------
'---------------------------------------------------------------------------------
Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.
'---------------------------------------------------------------------------------
'---------------------------------------------------------------------------------

'Scroll thrugh past stock prices and see how two moving
'averages behave, whether their intersections mean something, whether
'crossing from above or below signifies buy and/or sell signals...
'EMA(i+1) = a * EMA(i) + (1-a)*P(i+1)

Function ASSET_SINGLE_EMA_SIGNAL_FUNC(ByRef TICKER_STR As Variant, _
Optional ByVal START_DATE As Date, _
Optional ByVal END_DATE As Date, _
Optional ByVal REFERENCE_DATE As Date, _
Optional ByVal EMA1_PERIODS As Long = 5, _
Optional ByVal EMA2_PERIODS As Long = 50, _
Optional ByVal MA_PERIODS As Long = 50, _
Optional ByVal INITIAL_CASH As Double = 1000, _
Optional ByVal OUTPUT As Integer = 1)

Dim i As Long
Dim j As Long
Dim k As Long

Dim o As Long
Dim p As Long

Dim NROWS As Long
Dim NCOLUMNS As Long

Dim EMA1_FACTOR As Double
Dim EMA2_FACTOR As Double

Dim TEMP_SUM As Double
Dim MEAN_VAL As Double
Dim SIGMA_VAL As Double

Dim DATA_MATRIX As Variant
Dim TEMP_MATRIX As Variant

On Error GoTo ERROR_LABEL

If IsArray(TICKER_STR) = True Then
    DATA_MATRIX = TICKER_STR
Else
    DATA_MATRIX = _
    YAHOO_HISTORICAL_DATA_SERIE_FUNC(TICKER_STR, START_DATE, END_DATE, _
    "d", "DOHLCVA", False, True, True)
End If
NROWS = UBound(DATA_MATRIX, 1)


If REFERENCE_DATE = 0 Or REFERENCE_DATE <= DATA_MATRIX(1, 1) Then
    REFERENCE_DATE = DATA_MATRIX(2, 1)
End If

NCOLUMNS = 15
ReDim TEMP_MATRIX(0 To NROWS, 1 To NCOLUMNS)
For j = 1 To NCOLUMNS: For i = 1 To NROWS: TEMP_MATRIX(i, j) = "": Next i: Next j

TEMP_MATRIX(0, 1) = "DATE"
TEMP_MATRIX(0, 2) = "OPEN"
TEMP_MATRIX(0, 3) = "HIGH"
TEMP_MATRIX(0, 4) = "LOW"
TEMP_MATRIX(0, 5) = "CLOSE"
TEMP_MATRIX(0, 6) = "VOLUME"
TEMP_MATRIX(0, 7) = "ADJ CLOSE"

TEMP_MATRIX(0, 8) = "EMA1 " & EMA1_PERIODS & " DAY"
TEMP_MATRIX(0, 9) = "EMA2 " & EMA2_PERIODS & " DAY"
TEMP_MATRIX(0, 10) = "MA " & MA_PERIODS & " DAY"

TEMP_MATRIX(0, 11) = "Sell when Price moves below " & Format(EMA1_PERIODS, "0") & "-day EMA."
TEMP_MATRIX(0, 12) = "Buy when Price moves above " & Format(EMA1_PERIODS, "0") & "-day EMA and the " & _
                      Format(EMA1_PERIODS, "0") & "-day EMA slopes up."

TEMP_MATRIX(0, 13) = "EQUITY"
TEMP_MATRIX(0, 14) = "CASH"
TEMP_MATRIX(0, 15) = "SYSTEM BALANCE"


'-----------------------------------------------------------------------------------
EMA1_FACTOR = 1 - 2 / (EMA1_PERIODS + 1)
EMA2_FACTOR = 1 - 2 / (EMA2_PERIODS + 1)
'-----------------------------------------------------------------------------------
o = 0: p = 0
'-----------------------------------------------------------------------------------
For i = 1 To NROWS
'-----------------------------------------------------------------------------------
    For j = 1 To 7
        TEMP_MATRIX(i, j) = DATA_MATRIX(i, j)
    Next j
    TEMP_MATRIX(i, 6) = TEMP_MATRIX(i, 6) / 1000 'Volume
    
    If TEMP_MATRIX(i, 1) >= REFERENCE_DATE Then
        
        If TEMP_MATRIX(i, 1) > REFERENCE_DATE Then
            TEMP_MATRIX(i, 8) = EMA1_FACTOR * TEMP_MATRIX(i - 1, 8) + _
                                (1 - EMA1_FACTOR) * TEMP_MATRIX(i, 5) '7)
            
            TEMP_MATRIX(i, 9) = EMA2_FACTOR * TEMP_MATRIX(i - 1, 9) + _
                                (1 - EMA2_FACTOR) * TEMP_MATRIX(i, 5) '7)

            k = i - (MA_PERIODS + 1)
            TEMP_SUM = TEMP_SUM - TEMP_MATRIX(k, 5) '7)
            TEMP_SUM = TEMP_SUM + TEMP_MATRIX(i, 5) '7)
            TEMP_MATRIX(i, 10) = TEMP_SUM / (MA_PERIODS + 1)
            
            TEMP_MATRIX(i, 11) = IIf((TEMP_MATRIX(i - 1, 5) > TEMP_MATRIX(i - 1, 8) And _
                                      TEMP_MATRIX(i, 5) < TEMP_MATRIX(i - 1, 8)), TEMP_MATRIX(i, 5), 0)
            
            TEMP_MATRIX(i, 12) = IIf((TEMP_MATRIX(i - 1, 5) < TEMP_MATRIX(i - 1, 8) And _
                                      TEMP_MATRIX(i, 5) > TEMP_MATRIX(i - 1, 8) And _
                                      TEMP_MATRIX(i, 8) > TEMP_MATRIX(i - 1, 8)), TEMP_MATRIX(i, 5), 0)

            If TEMP_MATRIX(i, 11) > 0 Then
                TEMP_MATRIX(i, 13) = 0
            Else
                If TEMP_MATRIX(i, 12) > 0 And TEMP_MATRIX(i - 1, 14) > 0 Then
                    TEMP_MATRIX(i, 13) = TEMP_MATRIX(i - 1, 14)
                Else
                    TEMP_MATRIX(i, 13) = TEMP_MATRIX(i - 1, 13) * TEMP_MATRIX(i, 5) / TEMP_MATRIX(i - 1, 5)
                End If
            End If
            
            If TEMP_MATRIX(i, 11) > 0 And TEMP_MATRIX(i - 1, 13) > 0 Then
                TEMP_MATRIX(i, 14) = TEMP_MATRIX(i - 1, 13) * TEMP_MATRIX(i, 5) / TEMP_MATRIX(i - 1, 5)
            Else
                If TEMP_MATRIX(i, 12) > 0 Then
                    TEMP_MATRIX(i, 14) = 0
                Else
                    TEMP_MATRIX(i, 14) = TEMP_MATRIX(i - 1, 14)
                End If
            End If
        
            TEMP_MATRIX(i, 15) = TEMP_MATRIX(i, 13) + TEMP_MATRIX(i, 14)
            MEAN_VAL = MEAN_VAL + (TEMP_MATRIX(i, 15) / TEMP_MATRIX(i - 1, 15) - 1)
            p = p + 1
        
        ElseIf TEMP_MATRIX(i, 1) = REFERENCE_DATE Then
        
            k = i - EMA1_PERIODS
            TEMP_MATRIX(i, 8) = 0
            For j = i To k Step -1
                TEMP_MATRIX(i, 8) = TEMP_MATRIX(i, 8) + TEMP_MATRIX(j, 5) '7)
            Next j
            TEMP_MATRIX(i, 8) = TEMP_MATRIX(i, 8) / (EMA1_PERIODS + 1)
            
            k = i - EMA2_PERIODS
            TEMP_MATRIX(i, 9) = 0
            For j = i To k Step -1
                TEMP_MATRIX(i, 9) = TEMP_MATRIX(i, 9) + TEMP_MATRIX(j, 5) '7)
            Next j
            TEMP_MATRIX(i, 9) = TEMP_MATRIX(i, 9) / (EMA2_PERIODS + 1)
            
            k = i - MA_PERIODS
            TEMP_MATRIX(i, 10) = 0
            For j = i To k Step -1
                TEMP_MATRIX(i, 10) = TEMP_MATRIX(i, 10) + TEMP_MATRIX(j, 5) '7)
            Next j
            TEMP_SUM = TEMP_MATRIX(i, 10)
            TEMP_MATRIX(i, 10) = TEMP_MATRIX(i, 10) / (MA_PERIODS + 1)
            
            TEMP_MATRIX(i, 11) = 0
            TEMP_MATRIX(i, 12) = 0
            TEMP_MATRIX(i, 13) = 0
            TEMP_MATRIX(i, 14) = INITIAL_CASH
            TEMP_MATRIX(i, 15) = TEMP_MATRIX(i, 13) + TEMP_MATRIX(i, 14)
            o = i + 1
        End If
                
    End If
    
'-----------------------------------------------------------------------------------
Next i
'-----------------------------------------------------------------------------------

'-----------------------------------------------------------------------------------
Select Case OUTPUT
'-----------------------------------------------------------------------------------
Case 0
'-----------------------------------------------------------------------------------
    ASSET_SINGLE_EMA_SIGNAL_FUNC = TEMP_MATRIX
'-----------------------------------------------------------------------------------
Case Else
'-----------------------------------------------------------------------------------
    If p = 0 Then: GoTo ERROR_LABEL
    MEAN_VAL = MEAN_VAL / p
    SIGMA_VAL = 0
    For i = o To NROWS
        SIGMA_VAL = SIGMA_VAL + ((TEMP_MATRIX(i, 15) / TEMP_MATRIX(i - 1, 15) - 1) - MEAN_VAL) ^ 2
    Next i
    SIGMA_VAL = (SIGMA_VAL / p) ^ 0.5
    If OUTPUT = 1 Then
        ASSET_SINGLE_EMA_SIGNAL_FUNC = MEAN_VAL / SIGMA_VAL
    Else
        ASSET_SINGLE_EMA_SIGNAL_FUNC = Array(MEAN_VAL / SIGMA_VAL, MEAN_VAL, SIGMA_VAL)
    End If
'-----------------------------------------------------------------------------------
End Select
'-----------------------------------------------------------------------------------

Exit Function
ERROR_LABEL:
ASSET_SINGLE_EMA_SIGNAL_FUNC = "--"
End Function

'--------------------------------------------------------------------------------------------------------------------------------------------------
'Scroll thrugh past stock prices and see how two moving averages behave, whether their intersections mean something, whether
'crossing from above or below signifies buy and/or sell signals:
'--------------------------------------------------------------------------------------------------------------------------------------------------
'EMA(i+1) = a * EMA(i) + (1-a)*P(i+1)
'--------------------------------------------------------------------------------------------------------------------------------------------------
'First the basics:
'--------------------------------------------------------------------------------------------------------------------------------------------------
'http://www.gummy-stuff.org/moving-CAGR.htm
'http://www.gummy-stuff.org/MA.htm

'http://www.gummy-stuff.org/VMA.htm
'http://www.gummy-stuff.org/EMA.htm
'http://www.gummy-stuff.org/ema-formula.htm

'http://www.gummy-stuff.org/Bollinger.htm#VMA
'http://www.gummy-stuff.org/Bollinger.htm#EMA

'In class you will learn how to replicate this function directly from Excel:
'http://www.gummy-stuff.org/Excel/moving-average.xls
'--------------------------------------------------------------------------------------------------------------------------------------------------

Function ASSET_MA_EMA_VMA_SIGNAL_FUNC(ByRef TICKER_STR As Variant, _
Optional ByVal START_DATE As Date, _
Optional ByVal END_DATE As Date, _
Optional ByVal MA1_PERIODS As Long = 40, _
Optional ByVal MA2_PERIODS As Long = 5, _
Optional ByVal INITIAL_CASH As Double = 1000, _
Optional ByVal INITIAL_EQUITY As Double = 1000, _
Optional ByVal EMA_FLAG As Boolean = True, _
Optional ByVal VMA_FLAG As Boolean = True, _
Optional ByVal TRIGGER_INT As Integer = 1, _
Optional ByVal OUTPUT As Integer = 0)

'The moving averages are calculated at yesterday's Close … and any Buy/Sell is at today's Open (!!)
'Note: TRIGGER_INT is a number (1 or -1).
'If you choose 1 the buy (sell) signals occur when the #1 average crosses #2 from below (above).
'If you choose -1 the buy (sell) signals occur when the #1 average crosses #2 from above (below).

Dim i As Long
Dim j As Long
Dim k As Long

Dim NROWS As Long
Dim NCOLUMNS As Long

Dim MA1_SUM As Double
Dim MA2_SUM As Double

Dim MA1_VAL As Double
Dim MA2_VAL As Double

Dim VMA1V_SUM As Double
Dim VMA2V_SUM As Double

Dim VMA1P_SUM As Double
Dim VMA2P_SUM As Double

Dim VMA1_VAL As Double
Dim VMA2_VAL As Double

Dim EMA1_VAL As Double
Dim EMA2_VAL As Double

Dim EMA1_FACTOR As Double
Dim EMA2_FACTOR As Double

Dim TEMP_VAL As Double
Dim MEAN_VAL As Double
Dim SIGMA_VAL As Double

Dim DATA_MATRIX As Variant
Dim TEMP_MATRIX As Variant

On Error GoTo ERROR_LABEL

If IsArray(TICKER_STR) = True Then
    DATA_MATRIX = TICKER_STR
Else
    DATA_MATRIX = _
    YAHOO_HISTORICAL_DATA_SERIE_FUNC(TICKER_STR, START_DATE, END_DATE, _
    "d", "DOHLCVA", False, True, True)
End If
NROWS = UBound(DATA_MATRIX, 1)

NCOLUMNS = 15
ReDim TEMP_MATRIX(0 To NROWS, 1 To NCOLUMNS)
For j = 1 To NCOLUMNS: For i = 1 To NROWS: TEMP_MATRIX(i, j) = "": Next i: Next j

TEMP_MATRIX(0, 1) = "DATE"
TEMP_MATRIX(0, 2) = "OPEN"
TEMP_MATRIX(0, 3) = "HIGH"
TEMP_MATRIX(0, 4) = "LOW"
TEMP_MATRIX(0, 5) = "CLOSE"
TEMP_MATRIX(0, 6) = "VOLUME"
TEMP_MATRIX(0, 7) = "ADJ CLOSE"

If EMA_FLAG = True Then
    EMA1_FACTOR = 2 / (MA1_PERIODS + 1)
    EMA2_FACTOR = 2 / (MA2_PERIODS + 1)
    TEMP_MATRIX(0, 8) = "EMA1: " & MA1_PERIODS & " DAYS"
    TEMP_MATRIX(0, 9) = "EMA2: " & MA2_PERIODS & " DAYS"
Else
    If VMA_FLAG = True Then
        TEMP_MATRIX(0, 8) = "VMA1: " & MA1_PERIODS & " DAYS"
        TEMP_MATRIX(0, 9) = "VMA2: " & MA2_PERIODS & " DAYS"
    Else
        TEMP_MATRIX(0, 8) = "MA1: " & MA1_PERIODS & " DAYS"
        TEMP_MATRIX(0, 9) = "MA2: " & MA2_PERIODS & " DAYS"
    End If
End If

TEMP_MATRIX(0, 10) = "BUY SIGNAL"
TEMP_MATRIX(0, 11) = "SELL SIGNAL"
TEMP_MATRIX(0, 12) = "SHARES HELD"
TEMP_MATRIX(0, 13) = "CASH"
TEMP_MATRIX(0, 14) = "EQUITY"
TEMP_MATRIX(0, 15) = "SYSTEM BALANCE"
'-----------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------

i = 1
For j = 1 To 7
    TEMP_MATRIX(i, j) = DATA_MATRIX(i, j)
Next j
TEMP_MATRIX(i, 6) = TEMP_MATRIX(i, 6) / 1000 'Volume
TEMP_MATRIX(i, 8) = TEMP_MATRIX(i, 7)
TEMP_MATRIX(i, 9) = TEMP_MATRIX(i, 7)

If EMA_FLAG = True Then
    EMA1_VAL = TEMP_MATRIX(i, 7)
    EMA2_VAL = TEMP_MATRIX(i, 7)
Else
    If VMA_FLAG = True Then
        VMA1P_SUM = TEMP_MATRIX(i, 6) * TEMP_MATRIX(i, 7)
        VMA1V_SUM = TEMP_MATRIX(i, 6)
        VMA1_VAL = (VMA1P_SUM / VMA1V_SUM)
        
        VMA2P_SUM = TEMP_MATRIX(i, 6) * TEMP_MATRIX(i, 7)
        VMA2V_SUM = TEMP_MATRIX(i, 6)
        VMA2_VAL = (VMA2P_SUM / VMA2V_SUM)
    Else
        MA1_SUM = TEMP_MATRIX(i, 7)
        MA2_SUM = TEMP_MATRIX(i, 7)
        
        MA1_VAL = MA1_SUM / i
        MA2_VAL = MA2_SUM / i
    End If
End If

TEMP_MATRIX(i, 10) = 0
TEMP_MATRIX(i, 11) = 0
TEMP_MATRIX(i, 12) = INITIAL_EQUITY / TEMP_MATRIX(i, 7)

TEMP_MATRIX(i, 13) = INITIAL_CASH
TEMP_MATRIX(i, 14) = INITIAL_EQUITY
TEMP_MATRIX(i, 15) = TEMP_MATRIX(i, 13) + TEMP_MATRIX(i, 14)

MEAN_VAL = 0
If TRIGGER_INT <> 1 Then: TRIGGER_INT = -1
'-----------------------------------------------------------------------------------
For i = 2 To NROWS
'-----------------------------------------------------------------------------------
    For j = 1 To 7
        TEMP_MATRIX(i, j) = DATA_MATRIX(i, j)
    Next j
    TEMP_MATRIX(i, 6) = TEMP_MATRIX(i, 6) / 1000 'Volume
    If EMA_FLAG = True Then
        TEMP_MATRIX(i, 8) = EMA1_VAL
        TEMP_MATRIX(i, 9) = EMA2_VAL
        GoSub EMA_LINE
    Else
        If VMA_FLAG = True Then
            TEMP_MATRIX(i, 8) = VMA1_VAL
            TEMP_MATRIX(i, 9) = VMA2_VAL
            GoSub VMA_LINE
        Else
            TEMP_MATRIX(i, 8) = MA1_VAL
            TEMP_MATRIX(i, 9) = MA2_VAL
            GoSub MA_LINE
        End If
    End If
    
    TEMP_VAL = TEMP_MATRIX(i, 7) / TEMP_MATRIX(i, 5) * TEMP_MATRIX(i, 2) '(Adj/Close*Open)
    
    If (TRIGGER_INT * TEMP_MATRIX(i - 1, 8) < TRIGGER_INT * TEMP_MATRIX(i - 1, 9) And _
        TRIGGER_INT * TEMP_MATRIX(i, 8) > TRIGGER_INT * TEMP_MATRIX(i, 9)) Then
        TEMP_MATRIX(i, 10) = TEMP_VAL
    Else
        TEMP_MATRIX(i, 10) = 0
    End If
    
    If (TRIGGER_INT * TEMP_MATRIX(i - 1, 8) > TRIGGER_INT * TEMP_MATRIX(i - 1, 9) And _
        TRIGGER_INT * TEMP_MATRIX(i, 8) < TRIGGER_INT * TEMP_MATRIX(i, 9)) Then
        TEMP_MATRIX(i, 11) = TEMP_VAL
    Else
        TEMP_MATRIX(i, 11) = 0
    End If
    
    If TEMP_MATRIX(i, 10) > 0 Then
        TEMP_MATRIX(i, 12) = TEMP_MATRIX(i - 1, 12) + TEMP_MATRIX(i - 1, 13) / TEMP_MATRIX(i, 10)
    Else
        If TEMP_MATRIX(i, 11) > 0 Then
            TEMP_MATRIX(i, 12) = 0
        Else
            TEMP_MATRIX(i, 12) = TEMP_MATRIX(i - 1, 12)
        End If
    End If
    
    If TEMP_MATRIX(i, 10) > 0 Then
        TEMP_MATRIX(i, 13) = 0
    Else
        If TEMP_MATRIX(i, 11) > 0 Then
            TEMP_MATRIX(i, 13) = TEMP_MATRIX(i - 1, 13) + TEMP_MATRIX(i - 1, 12) * TEMP_MATRIX(i, 11)
        Else
            TEMP_MATRIX(i, 13) = TEMP_MATRIX(i - 1, 13)
        End If
    End If
    
    TEMP_MATRIX(i, 14) = TEMP_MATRIX(i, 12) * TEMP_MATRIX(i, 7)
    TEMP_MATRIX(i, 15) = TEMP_MATRIX(i, 13) + TEMP_MATRIX(i, 14)
    MEAN_VAL = MEAN_VAL + (TEMP_MATRIX(i, 15) / TEMP_MATRIX(i - 1, 15) - 1)

'-----------------------------------------------------------------------------------
Next i
'-----------------------------------------------------------------------------------

'-----------------------------------------------------------------------------------
Select Case OUTPUT
'-----------------------------------------------------------------------------------
Case 0
'-----------------------------------------------------------------------------------
    ASSET_MA_EMA_VMA_SIGNAL_FUNC = TEMP_MATRIX
'-----------------------------------------------------------------------------------
Case Else
'-----------------------------------------------------------------------------------
    MEAN_VAL = MEAN_VAL / (NROWS - 1)
    SIGMA_VAL = 0
    For i = 2 To NROWS
        SIGMA_VAL = SIGMA_VAL + ((TEMP_MATRIX(i, 15) / TEMP_MATRIX(i - 1, 15) - 1) - MEAN_VAL) ^ 2
    Next i
    SIGMA_VAL = (SIGMA_VAL / (NROWS - 1)) ^ 0.5
    If OUTPUT = 1 Then
        ASSET_MA_EMA_VMA_SIGNAL_FUNC = MEAN_VAL / SIGMA_VAL 'Objective function to MAXIMIZE!!!!
    Else
        ASSET_MA_EMA_VMA_SIGNAL_FUNC = Array(MEAN_VAL / SIGMA_VAL, MEAN_VAL, SIGMA_VAL)
    End If
'-----------------------------------------------------------------------------------
End Select
'-----------------------------------------------------------------------------------

Exit Function
'-----------------------------------------------------------------------------------
EMA_LINE:
    EMA1_VAL = (1 - EMA1_FACTOR) * EMA1_VAL + EMA1_FACTOR * TEMP_MATRIX(i, 7)
    EMA2_VAL = (1 - EMA2_FACTOR) * EMA2_VAL + EMA2_FACTOR * TEMP_MATRIX(i, 7)
Return
'-----------------------------------------------------------------------------------
MA_LINE:
    MA1_SUM = MA1_SUM + TEMP_MATRIX(i, 7)
    If i <= MA1_PERIODS Then
        MA1_VAL = MA1_SUM / i
    Else
        MA1_VAL = MA1_SUM / (MA1_PERIODS + 1)
        k = i - MA1_PERIODS
        MA1_SUM = MA1_SUM - TEMP_MATRIX(k, 7)
    End If
    
    MA2_SUM = MA2_SUM + TEMP_MATRIX(i, 7)
    If i <= MA2_PERIODS Then
        MA2_VAL = MA2_SUM / i
    Else
        MA2_VAL = MA2_SUM / (MA2_PERIODS + 1)
        k = i - MA2_PERIODS
        MA2_SUM = MA2_SUM - TEMP_MATRIX(k, 7)
    End If
Return
'-----------------------------------------------------------------------------------
VMA_LINE:

    VMA1P_SUM = VMA1P_SUM + TEMP_MATRIX(i, 6) * TEMP_MATRIX(i, 7)
    VMA1V_SUM = VMA1V_SUM + TEMP_MATRIX(i, 6)
        
    If i <= MA1_PERIODS Then
        VMA1_VAL = (VMA1P_SUM / VMA1V_SUM)
    Else
        k = i - MA1_PERIODS
        VMA1P_SUM = VMA1P_SUM - TEMP_MATRIX(k, 6) * TEMP_MATRIX(k, 7)
        VMA1V_SUM = VMA1V_SUM - TEMP_MATRIX(k, 6)
        
        VMA1_VAL = (VMA1P_SUM / VMA1V_SUM)
    End If
    
    VMA2P_SUM = VMA2P_SUM + TEMP_MATRIX(i, 6) * TEMP_MATRIX(i, 7)
    VMA2V_SUM = VMA2V_SUM + TEMP_MATRIX(i, 6)
        
    If i <= MA2_PERIODS Then
        VMA2_VAL = (VMA2P_SUM / VMA2V_SUM)
    Else
        k = i - MA2_PERIODS
        VMA2P_SUM = VMA2P_SUM - TEMP_MATRIX(k, 6) * TEMP_MATRIX(k, 7)
        VMA2V_SUM = VMA2V_SUM - TEMP_MATRIX(k, 6)
        
        VMA2_VAL = (VMA2P_SUM / VMA2V_SUM)
    End If
Return
'-----------------------------------------------------------------------------------
ERROR_LABEL:
ASSET_MA_EMA_VMA_SIGNAL_FUNC = "--"
End Function
