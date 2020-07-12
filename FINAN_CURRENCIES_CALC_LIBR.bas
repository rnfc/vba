Attribute VB_Name = "FINAN_CURRENCIES_CALC_LIBR"

'-------------------------------------------------------------------------------------
Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.
'-------------------------------------------------------------------------------------


'************************************************************************************
'************************************************************************************
'FUNCTION      : CURRENCIES_RELATIVE_CALC_FUNC

'DESCRIPTION   : FX Relative Calculator
'Trades are expressed in exchange rate pairs, where the first
'currency in the pair is the object of the trade, for example,
'1 million dollar/mark at 1.5825 means US Dollar 1
'million exchanged for Deutsche Marks 1,585,000.

'To understand triangular arbitrage, let us take a hypothetical
'example where the exchange rate between the dollar and the British
'Pound is denoted by S$/£ (“S” for Spot, and the subscripts denoting
'the currencies involved), and the exchange rate between the
'dollar and the German Mark is denoted by S$/DM .

'For example, if the exchange rate between the dollar and the
'British Pound is $1.50/£1 and the exchange rate between the
'dollar and the German Mark is $.75/DM1, the exchange rate between
'the German Mark and the British Pound is
'SDM/£ =($1.50/£1) / ($0.75/DM1) = DM2/£1 .

'If the exchange rate between the German Mark and the British Pound
'were either greater or less than DM2/£1, then a triangular arbitrage
'opportunity will be available.

'For example, suppose that the Mark/Pound exchange rate were
'DM2.1/£1. Then a trader with two German Marks would (1) exchange
'them for $1.50 (2 x $.75/DM1).

'The $1.5 would then (2) be exchanged for one British Pound, which
'would then (3) be used to purchase DM 2.1, which is greater than
'the number of German Marks that the trader started with.
'------------------------------------------------------------------------------

'LIBRARY       : CURRENCIES
'GROUP         : CALC
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/05/2011
'************************************************************************************
'************************************************************************************

Function CURRENCIES_RELATIVE_CALC_FUNC(ByRef DATA_RNG As Variant, _
ByVal X_VAL As Double, _
Optional ByVal BASE_RATE As Long = 1)

Dim i As Long
Dim NROWS As Long
Dim DATA_MATRIX As Variant
Dim TEMP_MATRIX As Variant

'-----------------------------------------------------------------------------------------------------
On Error GoTo ERROR_LABEL
'-----------------------------------------------------------------------------------------------------
DATA_MATRIX = DATA_RNG
If UBound(DATA_MATRIX, 1) = 1 Then
    DATA_MATRIX = MATRIX_TRANSPOSE_FUNC(DATA_MATRIX)
End If
NROWS = UBound(DATA_MATRIX, 1)
'-----------------------------------------------------------------------------------------------------
ReDim TEMP_MATRIX(0 To NROWS, 1 To 5)
'-----------------------------------------------------------------------------------------------------
TEMP_MATRIX(0, 1) = CStr("BASE EXCHANGE RATE: <" & DATA_MATRIX(BASE_RATE, 1) & ">")
TEMP_MATRIX(0, 2) = ("TIME")
TEMP_MATRIX(0, 3) = ("BID RATES")
TEMP_MATRIX(0, 4) = ("ASK RATES")
TEMP_MATRIX(0, 5) = ("AVG. RATES")
'-----------------------------------------------------------------------------------------------------
For i = 1 To NROWS
'-----------------------------------------------------------------------------------------------------
    If IsNumeric(DATA_MATRIX(i, 5)) = True And IsNumeric(DATA_MATRIX(BASE_RATE, 5)) = True Then
        If DATA_MATRIX(BASE_RATE, 5) <> 0 Then
            TEMP_MATRIX(i, 5) = ((X_VAL / DATA_MATRIX(BASE_RATE, 5)) * DATA_MATRIX(i, 5))
        Else
            TEMP_MATRIX(i, 5) = ""
        End If
    Else
        TEMP_MATRIX(i, 5) = ""
    End If
    
    If IsNumeric(DATA_MATRIX(i, 4)) = True And IsNumeric(DATA_MATRIX(BASE_RATE, 4)) = True Then
        If DATA_MATRIX(BASE_RATE, 4) <> 0 Then
            TEMP_MATRIX(i, 4) = ((X_VAL / DATA_MATRIX(BASE_RATE, 4)) * DATA_MATRIX(i, 4))
        Else
            TEMP_MATRIX(i, 4) = ""
        End If
    Else
        TEMP_MATRIX(i, 4) = ""
    End If

    If IsNumeric(DATA_MATRIX(i, 3)) = True And IsNumeric(DATA_MATRIX(BASE_RATE, 3)) = True Then
        If DATA_MATRIX(BASE_RATE, 3) <> 0 Then
            TEMP_MATRIX(i, 3) = ((X_VAL / DATA_MATRIX(BASE_RATE, 3)) * DATA_MATRIX(i, 3))
        Else
            TEMP_MATRIX(i, 3) = ""
        End If
    Else
        TEMP_MATRIX(i, 3) = ""
    End If
    TEMP_MATRIX(i, 2) = DATA_MATRIX(i, 2)
    TEMP_MATRIX(i, 1) = DATA_MATRIX(i, 1) 'Name
'-----------------------------------------------------------------------------------------------------
Next i
'-----------------------------------------------------------------------------------------------------
CURRENCIES_RELATIVE_CALC_FUNC = TEMP_MATRIX
'-----------------------------------------------------------------------------------------------------

Exit Function
ERROR_LABEL:
CURRENCIES_RELATIVE_CALC_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : EXCHANGE_RATE_TABLE_FUNC

'DESCRIPTION   : Country, Currency, Exchange Rate, Starting Value and Increment
'LIBRARY       : CURRENCIES
'GROUP         : CALC
'ID            : 002
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/05/2011
'************************************************************************************
'************************************************************************************

Function EXCHANGE_RATE_TABLE_FUNC(ByVal EXCHANGE_RATE As Double, _
Optional ByVal QUOTE_CURRENCY_STR As String = "CDN", _
Optional ByVal BASE_CURRENCY_STR As String = "USD", _
Optional ByVal MIN_VAL As Double = 10, _
Optional ByVal DELTA_VAL As Double = 10, _
Optional ByVal NBINS As Long = 10)

Dim i As Long
Dim TEMP_SUM As Double
Dim TEMP_VECTOR As Variant

On Error GoTo ERROR_LABEL

ReDim TEMP_VECTOR(0 To NBINS, 1 To 2)
TEMP_VECTOR(0, 1) = Format(EXCHANGE_RATE, "0.00") & " " & QUOTE_CURRENCY_STR & " = "
TEMP_VECTOR(0, 2) = 1 & " " & BASE_CURRENCY_STR

TEMP_SUM = MIN_VAL
For i = 1 To NBINS
    TEMP_VECTOR(i, 1) = TEMP_SUM
    TEMP_VECTOR(i, 2) = TEMP_VECTOR(i, 1) / EXCHANGE_RATE
    TEMP_SUM = TEMP_SUM + DELTA_VAL
Next i

EXCHANGE_RATE_TABLE_FUNC = TEMP_VECTOR

Exit Function
ERROR_LABEL:
EXCHANGE_RATE_TABLE_FUNC = Err.number
End Function
