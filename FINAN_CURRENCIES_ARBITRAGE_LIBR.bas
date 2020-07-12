Attribute VB_Name = "FINAN_CURRENCIES_ARBITRAGE_LIBR"


Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.

'************************************************************************************
'************************************************************************************
'FUNCTION      : CURRENCIES_ROUND_TRIP_FUNC

'DESCRIPTION   : The notion of arbitrage in this function is simple. It means
'making money by exploiting price differences in the value of an asset in
'different markets.

'This function looks at the possibilities of making money in the foreign exchange
'markets by buying and selling foreign currencies in the spot market using different
'cross currency rates.

'Spot trading in the foreign currency markets is for immediate delivery,
'though in reality it often means T+2, i.e. trades are settled 2 days
'after the transaction is done. The majority of foreign currency transactions
'involve the US Dollar. Spot rates between two currencies are normally
'quoted in the format Bid/Offer. For example, a quotation on dollar /
'mark of 1.5825-30 means that the person (or bank) making the quote is
'willing to sell marks for dollars at the rate of 1 dollar = 1.5825 marks,
'and is willing to buy marks for dollars at the rate of 1 dollar = 1.5830.


'LIBRARY       : CURRENCIES
'GROUP         : ARBITRAGE
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 03/16/2009
'************************************************************************************
'************************************************************************************

Function CURRENCIES_ROUND_TRIP_FUNC(ByVal BID_ASK_RNG As Variant, _
Optional ByVal NSIZE As Integer = 6, _
Optional ByVal INITIAL_INVESTMENT As Double = 1000000, _
Optional ByVal TARGET_PROFIT As Double = 0, _
Optional ByVal UNIT_TRANSACTION_COST As Double = 20, _
Optional ByVal BASE_FX_STR As String = "USD", _
Optional ByVal SANITIZED_FLAG As Boolean = True, _
Optional ByVal OUTPUT As Integer = 1)

Dim DATA_MATRIX As Variant

On Error GoTo ERROR_LABEL

DATA_MATRIX = BID_ASK_RNG
If UBound(DATA_MATRIX, 1) <> UBound(DATA_MATRIX, 2) Then: GoTo ERROR_LABEL

If NSIZE < 2 Then: GoTo ERROR_LABEL
If NSIZE > 8 Then: NSIZE = 8
If NSIZE > (UBound(DATA_MATRIX, 1) - 1) Then GoTo ERROR_LABEL

If SANITIZED_FLAG = True Then
    DATA_MATRIX = CURRENCIES_SANITIZED_CROSS_MATRIX_FUNC(DATA_MATRIX)
End If

Select Case NSIZE
Case 2
    CURRENCIES_ROUND_TRIP_FUNC = CURRENCIES_TRIANGLE_TWO_FUNC(DATA_MATRIX, INITIAL_INVESTMENT, _
    TARGET_PROFIT, UNIT_TRANSACTION_COST, BASE_FX_STR, OUTPUT)
Case 3
    CURRENCIES_ROUND_TRIP_FUNC = CURRENCIES_TRIANGLE_THREE_FUNC(DATA_MATRIX, INITIAL_INVESTMENT, _
    TARGET_PROFIT, UNIT_TRANSACTION_COST, BASE_FX_STR, OUTPUT)
Case 4
    CURRENCIES_ROUND_TRIP_FUNC = CURRENCIES_TRIANGLE_FOUR_FUNC(DATA_MATRIX, INITIAL_INVESTMENT, _
    TARGET_PROFIT, UNIT_TRANSACTION_COST, BASE_FX_STR, OUTPUT)
Case 5
    CURRENCIES_ROUND_TRIP_FUNC = CURRENCIES_TRIANGLE_FIVE_FUNC(DATA_MATRIX, INITIAL_INVESTMENT, _
    TARGET_PROFIT, UNIT_TRANSACTION_COST, BASE_FX_STR, OUTPUT)
Case 6
    CURRENCIES_ROUND_TRIP_FUNC = CURRENCIES_TRIANGLE_SIX_FUNC(DATA_MATRIX, INITIAL_INVESTMENT, _
    TARGET_PROFIT, UNIT_TRANSACTION_COST, BASE_FX_STR, OUTPUT)
Case 7
    CURRENCIES_ROUND_TRIP_FUNC = CURRENCIES_TRIANGLE_SEVEN_FUNC(DATA_MATRIX, INITIAL_INVESTMENT, _
    TARGET_PROFIT, UNIT_TRANSACTION_COST, BASE_FX_STR, OUTPUT)
Case 8
    CURRENCIES_ROUND_TRIP_FUNC = CURRENCIES_TRIANGLE_EIGHT_FUNC(DATA_MATRIX, INITIAL_INVESTMENT, _
    TARGET_PROFIT, UNIT_TRANSACTION_COST, BASE_FX_STR, OUTPUT)
End Select

Exit Function
ERROR_LABEL:
CURRENCIES_ROUND_TRIP_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************



'************************************************************************************
'************************************************************************************
'FUNCTION      : CURRENCIES_SANITIZED_CROSS_MATRIX_FUNC
'DESCRIPTION   : FX Sanitized Cross Matrix Function
'LIBRARY       : CURRENCIES
'GROUP         : ARBITRAGE
'ID            : 002
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 03/16/2009
'************************************************************************************
'************************************************************************************

Private Function CURRENCIES_SANITIZED_CROSS_MATRIX_FUNC( _
ByVal DATA_MATRIX As Variant)

Dim i As Long
Dim j As Long
Dim NSIZE As Long

Dim TEMP_MATRIX As Variant
'Dim DATA_MATRIX As Variant

On Error GoTo ERROR_LABEL
'DATA_MATRIX = BID_ASK_RNG
NSIZE = UBound(DATA_MATRIX, 1) - 1

ReDim TEMP_MATRIX(1 To NSIZE + 1, 1 To NSIZE + 1)

For i = 1 To NSIZE
    TEMP_MATRIX(1, i + 1) = DATA_MATRIX(1, i + 1)
    TEMP_MATRIX(i + 1, 1) = DATA_MATRIX(1, i + 1)
Next i

TEMP_MATRIX(1, 1) = "--"

For j = 2 To NSIZE + 1
    For i = j To NSIZE + 1
        If i = j Then
            TEMP_MATRIX(i, j) = 1
        Else
            TEMP_MATRIX(i, j) = DATA_MATRIX(NSIZE + 3 - i, j)
        End If
    Next i
    For i = 2 To j
        If i = j Then
            TEMP_MATRIX(i, j) = 1
        Else
            TEMP_MATRIX(i, j) = DATA_MATRIX(NSIZE + 3 - i, j)
        End If
    Next i
Next j

CURRENCIES_SANITIZED_CROSS_MATRIX_FUNC = TEMP_MATRIX

Exit Function
ERROR_LABEL:
CURRENCIES_SANITIZED_CROSS_MATRIX_FUNC = Err.number
End Function



'************************************************************************************
'************************************************************************************



'************************************************************************************
'************************************************************************************

'FUNCTION      : CURRENCIES_TRIANGLE_TWO_FUNC; CURRENCIES_TRIANGLE_THREE_FUNC; CURRENCIES_TRIANGLE_FOUR_FUNC;
'CURRENCIES_TRIANGLE_FIVE_FUNC; CURRENCIES_TRIANGLE_SIX_FUNC; CURRENCIES_TRIANGLE_SEVEN_FUNC; CURRENCIES_TRIANGLE_EIGHT_FUNC

'DESCRIPTION   : As different currency pairs are traded in different markets and
'the exchange rates may be momentarily out of sync, this may allow the arbitrageur
'to make a small profit by doing such a “round-trip”, say for example by converting
'USD..GBP..EUR..USD. But there is no reason why a “round-trip” should be limited
'to two intervening currencies. It is possible to do much longer round-trips,
'for example: USD..CHF..EUR..JPY..AUD..USD.

'This gives us complete flexibility in examining the entire range of possibilities
'that are available with the given set of currencies.

'However, as the number of possible “round-trips” increases very rapidly as we increase
'the number of currencies, the model examines only “round-trips” with six currencies
'otherwise the computational time required goes up significantly without any additional
'contribution to our understanding of market efficiency.

'The model considers all possible combinations of currencies in a 2 to 8-currency
'“roundtrip” (surrounded by USD on each side), and identifies the profit maximising
'“roundtrip” for an investor who “invests” USD 1 million into the round-trip.

'It is important to note that under spot transactions this “investment” is recouped
'instantaneously, and is therefore not really an investment but just a convenient
'number.

'The amount of USD 1 million is entirely arbitrary, we can use any amount, even $1,
'only that the profits will be proportionately less. The amount of $X million is
'chosen to see round numbers that are easier to understand than a large number of
'decimals.

'Considering bid-offer spreads: Bid offer spreads (that are really a form of
'transaction costs) are considered as part of the exchange rates table mentioned
'earlier.

'Considering transaction costs: Large institutional investors do not normally pay
'pertransaction costs of trading in the foreign exchange markets. However,
'individuals or smaller investors normally incur a fixed fee per transaction.
'This fixed fee is user specified as an input parameter when the model is run.
'It can be specified as “0” (zero) if that be the case.

'LIBRARY       : CURRENCIES
'GROUP         : ARBITRAGE
'ID            : 003-009
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 03/16/2009
'************************************************************************************
'************************************************************************************

Private Function CURRENCIES_TRIANGLE_TWO_FUNC(ByRef DATA_RNG As Variant, _
Optional ByVal INITIAL_INVESTMENT As Double = 15000000, _
Optional ByVal TARGET_PROFIT As Double = 0, _
Optional ByVal UNIT_TRANSACTION_COST As Double = 20, _
Optional ByVal BASE_FX_STR As String = "USD", _
Optional ByVal OUTPUT As Integer = 1)

'DATA_RNG = Sanitised cross currency exchange rates
'where, lower triangle = bid rates, upper triangle = offer rates
'TARGET_PROFIT: initial value for profit

'UNIT_TRANSACTION_COST: transaction costs: enter transaction cost per
'transaction in base currency.

'TARGET_PROFIT: TARGET_PROFIT_PROFIT

'--------------------------------Cross rates-------------------------------
'Most currencies are expressed against dollars – which means there is
'always a buy-sell spread that would normally make it difficult to make
'money by routing transactions through dollars. However, there are some
'cross rates that trade directly. A cross rate is a rate between two
'non US Dollar currencies. For example, there may be trading
'between British Pounds and the Euro. Therefore, we have a rate for the
'Dollar-Euro, we have a rate for the Dollar-Pound, and we also have a rate
'for Pound-Euro.
'--------------------------------------------------------------------------

Dim h As Long
Dim i As Long
Dim j As Long
Dim k As Long

Dim C0 As Long
Dim C1 As Long
Dim C2 As Long
Dim c3 As Long

Dim NSIZE As Long

Dim DST_ZERO As Long
Dim DST_ONE As Long
Dim DST_TWO As Long
Dim DST_THREE As Long

Dim ZERO_STR As String
Dim ONE_STR As String
Dim TWO_STR As String
Dim THREE_STR As String

Dim TEMP_STR As String
Dim DATA_MATRIX As Variant
Dim TEMP_VECTOR As Variant

Dim TEMP_SUM As Double

On Error GoTo ERROR_LABEL

DATA_MATRIX = MATRIX_REMOVE_ROWS_FUNC(MATRIX_REMOVE_COLUMNS_FUNC(DATA_RNG, 1, 1), 1, 1)
TEMP_VECTOR = MATRIX_REMOVE_COLUMNS_FUNC(MATRIX_GET_ROW_FUNC(DATA_RNG, 1, 1), 1, 1)
NSIZE = 2

If UBound(DATA_MATRIX, 1) < NSIZE Then GoTo ERROR_LABEL
If UBound(DATA_MATRIX, 2) < NSIZE Then GoTo ERROR_LABEL

'For any given number of currencies ‘N’, the number of cross-rates
'theoretically possible is N x (N-1)/2. The number of variations
'possible with N currencies to construct a “round-trip” would be NN,
'which means that the number of possibilities increases exponentially
'as we increase the number of currencies being considered.

'Note that each of the intervening currencies CC1 to 6 can be any of
'the 6 currencies, there is no limitation that a currency trade cannot
'be done twice, therefore the number of possible combinations is 6 x 6
'x 6 x 6 x 6 x 6, or 66. This is important because it allows us to look
'at profitable combinations where the number of currencies is less than
'six.

If IsNumeric(UNIT_TRANSACTION_COST) = False Then UNIT_TRANSACTION_COST = 0

'Use six nested “For…Next” loops to cycle through each possible combination
'of currencies. Each time, calculate the profit that would result from
'executing the given combination of currency transactions, and if this
'profit is greater than the profit from a previous combination, then store
'it away in a variable, otherwise discard it. (The initial value of profit is 0)
'If a profitable combination is not found – which for instance will be the case
'where transaction costs are very high, the macro says so. If a profitable
'combination is found, the same is listed together with all combinations.

For C1 = 1 To NSIZE
For C2 = 1 To NSIZE

'Update the status bar to show the combination being checked.
'      Excel.Application.StatusBar = _
 '     TEMP_VECTOR(1, c1) & "-" & _
      'TEMP_VECTOR(1, c2)

'Calculate true number of transactions. When the same _
'currency is repeated, then it is not a transaction, for _
'example USD->GBP->GBP->JPY->USD = USD->GBP->JPY->USD, ie _
'4 transactions and not 5, as GBP->GBP is costless. This is
'needed to calculate true transaction costs.

j = NSIZE + 1

If 1 = C1 Then j = j - 1
If C1 = C2 Then j = j - 1
If C2 = 1 Then j = j - 1

'For the currency combination in the loop, compare _
'TARGET_PROFIT to previous highest profits.

If TARGET_PROFIT < ((INITIAL_INVESTMENT * _
    DATA_MATRIX(C1, 1) * _
    DATA_MATRIX(C2, C1) * _
    DATA_MATRIX(1, C2) - _
    INITIAL_INVESTMENT) - (j * UNIT_TRANSACTION_COST)) _
Then
    
    DST_ZERO = 1
    DST_ONE = C1
    DST_TWO = C2
    DST_THREE = 1
    
    k = 1
    
    TARGET_PROFIT = ((INITIAL_INVESTMENT * _
    DATA_MATRIX(C1, 1) * _
    DATA_MATRIX(C2, C1) * _
    DATA_MATRIX(1, C2) - _
    INITIAL_INVESTMENT) - (j * UNIT_TRANSACTION_COST))
    
    h = j
End If

'ENHANCING THE MODEL:

'The model can be extended beyond spot transactions to look at forward
'transactions, interest rates, and identify any opportunities that may
'arise from an imbalance between the spot and forward rates given the
'interest rates in the two currencies. Such arbitrage would be based
'on violations of interest rate parity.

'In economies where foreign exchange markets are regulated, there often
'exists a black market exchange rate that is different from the official
'rate. Arbitrage opportunities between these rates are always possible
'and there is some money to be made. However, this may involve some risk
'taking in the form of possible violations of foreign exchange regulations.
'The cost of these violations with the probability of being discovered
'will need to be factored into the model.

TEMP_SUM = TEMP_SUM + 1   'merely a count of how many combinations checked.

Next C2
Next C1

'By now, the most profitable opportunity has been identified. If there _
'is no profitable opportunity, then the "GoSub finito" would not have _
'been visited and k would be zero)

If k = 0 Then
    CURRENCIES_TRIANGLE_TWO_FUNC = "No profitable possibilities!"
    Exit Function
Else
    
    ZERO_STR = TEMP_VECTOR(1, 1)
    'Replace currency numbers by their respective English
    ONE_STR = TEMP_VECTOR(1, DST_ONE)       'codes, eg 1 means USD etc
    TWO_STR = TEMP_VECTOR(1, DST_TWO)
    THREE_STR = TEMP_VECTOR(1, 1)
    
    
    For i = 1 To NSIZE + 1  'Remove duplicates, ie "GBP --> GBP" is _
    merely "GBP" in the transaction.
        If ZERO_STR = ONE_STR Then ZERO_STR = ""
        If ONE_STR = TWO_STR Then ONE_STR = ""
        If TWO_STR = THREE_STR Then TWO_STR = ""
    Next i
End If

Select Case OUTPUT
Case 0
    
    TEMP_STR = "Transact currencies as follows: " & _
        ZERO_STR & "-" & _
        ONE_STR & "-" & _
        TWO_STR & "-" & _
        THREE_STR & _
        " Profit on a " & _
        Format(INITIAL_INVESTMENT, "#,0.00") & " = " & _
        Format(TARGET_PROFIT, "#,0.00") & _
        "  (" & TEMP_SUM & " possibilities checked )"
        
        CURRENCIES_TRIANGLE_TWO_FUNC = TEMP_STR
        
Case Else

    ReDim TEMP_MATRIX(0 To NSIZE + 4, 1 To 6)
    
    TEMP_MATRIX(0, 1) = "Transaction"
    TEMP_MATRIX(0, 2) = "Details"
    TEMP_MATRIX(0, 3) = "Amount sold"
    TEMP_MATRIX(0, 4) = "Exchange rate"
    TEMP_MATRIX(0, 5) = "Details"
    TEMP_MATRIX(0, 6) = "Amount bought"

    If ZERO_STR = "" Then ZERO_STR = BASE_FX_STR
    If ONE_STR = "" Then ONE_STR = BASE_FX_STR
    If TWO_STR = "" Then TWO_STR = BASE_FX_STR
    If THREE_STR = "" Then THREE_STR = BASE_FX_STR
    
    'Write the transaction
    TEMP_MATRIX(1, 1) = "Sell " & ZERO_STR & ", Buy " & ONE_STR
    TEMP_MATRIX(2, 1) = "Sell " & ONE_STR & ", Buy " & TWO_STR
    TEMP_MATRIX(3, 1) = "Sell " & TWO_STR & ", Buy " & THREE_STR
    
    'Write the transaction details for currency sold
    TEMP_MATRIX(1, 2) = ZERO_STR & " sold:"
    TEMP_MATRIX(2, 2) = ONE_STR & " sold:"
    TEMP_MATRIX(3, 2) = TWO_STR & " sold:"
    
    'Write the transaction details for currency bought
    TEMP_MATRIX(1, 5) = ONE_STR & " bought:"
    TEMP_MATRIX(2, 5) = TWO_STR & " bought:"
    TEMP_MATRIX(3, 5) = THREE_STR & " bought:"
    
    'MATRIX_FIND_ELEMENT_FUNC
    
    C0 = MATRIX_FIND_ELEMENT_FUNC(TEMP_VECTOR, ZERO_STR, 1, 1, 0)
    C1 = MATRIX_FIND_ELEMENT_FUNC(TEMP_VECTOR, ONE_STR, 1, 1, 0)
    C2 = MATRIX_FIND_ELEMENT_FUNC(TEMP_VECTOR, TWO_STR, 1, 1, 0)
    c3 = MATRIX_FIND_ELEMENT_FUNC(TEMP_VECTOR, THREE_STR, 1, 1, 0)
    
    TEMP_MATRIX(1, 4) = DATA_MATRIX(C1, C0)
    
    TEMP_MATRIX(2, 4) = DATA_MATRIX(C2, C1)
    
    TEMP_MATRIX(3, 4) = DATA_MATRIX(c3, C2)
    
        
    TEMP_MATRIX(1, 3) = INITIAL_INVESTMENT
    
    TEMP_MATRIX(2, 3) = TEMP_MATRIX(1, 3) * TEMP_MATRIX(1, 4)
    TEMP_MATRIX(3, 3) = TEMP_MATRIX(2, 3) * TEMP_MATRIX(2, 4)
    
    TEMP_MATRIX(1, 6) = TEMP_MATRIX(1, 3) * TEMP_MATRIX(1, 4)
    TEMP_MATRIX(2, 6) = TEMP_MATRIX(2, 3) * TEMP_MATRIX(2, 4)
    TEMP_MATRIX(3, 6) = TEMP_MATRIX(3, 3) * TEMP_MATRIX(3, 4)

    TEMP_MATRIX(NSIZE + 2, 1) = "Gross Profit:"
    TEMP_MATRIX(NSIZE + 2, 2) = "Net: " & BASE_FX_STR
    TEMP_MATRIX(NSIZE + 2, 3) = TEMP_MATRIX(NSIZE + 1, 6) - INITIAL_INVESTMENT
    
    TEMP_MATRIX(NSIZE + 2, 4) = ""
    TEMP_MATRIX(NSIZE + 2, 5) = ""
    TEMP_MATRIX(NSIZE + 2, 6) = ""
    
    TEMP_MATRIX(NSIZE + 3, 1) = "Less: Transaction Cost (" & _
    Format(h, "0.00") & _
    " * $" & (Format(UNIT_TRANSACTION_COST, "0.00")) & ")"
    TEMP_MATRIX(NSIZE + 3, 2) = ""

    TEMP_MATRIX(NSIZE + 3, 3) = -UNIT_TRANSACTION_COST * h
    
    TEMP_MATRIX(NSIZE + 3, 4) = ""
    TEMP_MATRIX(NSIZE + 3, 5) = ""
    TEMP_MATRIX(NSIZE + 3, 6) = ""
    
    TEMP_MATRIX(NSIZE + 4, 1) = "Net Profit"
    TEMP_MATRIX(NSIZE + 4, 2) = ""
    
    TEMP_MATRIX(NSIZE + 4, 3) = TEMP_MATRIX(NSIZE + 3, 3) + _
    TEMP_MATRIX(NSIZE + 2, 3)

    TEMP_MATRIX(NSIZE + 4, 4) = ""
    TEMP_MATRIX(NSIZE + 4, 5) = ""
    TEMP_MATRIX(NSIZE + 4, 6) = ""

    CURRENCIES_TRIANGLE_TWO_FUNC = TEMP_MATRIX
End Select

Exit Function
ERROR_LABEL:
CURRENCIES_TRIANGLE_TWO_FUNC = Err.number
End Function
Private Function CURRENCIES_TRIANGLE_THREE_FUNC(ByRef DATA_RNG As Variant, _
Optional ByVal INITIAL_INVESTMENT As Double = 15000000, _
Optional ByVal TARGET_PROFIT As Double = 0, _
Optional ByVal UNIT_TRANSACTION_COST As Double = 20, _
Optional ByVal BASE_FX_STR As String = "USD", _
Optional ByVal OUTPUT As Integer = 1)

'DATA_RNG = Sanitised cross currency exchange rates
'where, lower triangle = bid rates, upper triangle = offer rates
'TARGET_PROFIT: initial value for profit

'UNIT_TRANSACTION_COST: transaction costs: enter transaction cost per
'transaction in base currency.

'TARGET_PROFIT: TARGET_PROFIT_PROFIT

'--------------------------------Cross rates-------------------------------
'Most currencies are expressed against dollars – which means there is
'always a buy-sell spread that would normally make it difficult to make
'money by routing transactions through dollars. However, there are some
'cross rates that trade directly. A cross rate is a rate between two
'non US Dollar currencies. For example, there may be trading
'between British Pounds and the Euro. Therefore, we have a rate for the
'Dollar-Euro, we have a rate for the Dollar-Pound, and we also have a rate
'for Pound-Euro.
'--------------------------------------------------------------------------

Dim h As Long
Dim i As Long
Dim j As Long
Dim k As Long

Dim C0 As Long
Dim C1 As Long
Dim C2 As Long
Dim c3 As Long
Dim c4 As Long

Dim NSIZE As Long

Dim DST_ZERO As Long
Dim DST_ONE As Long
Dim DST_TWO As Long
Dim DST_THREE As Long
Dim DST_FOUR As Long

Dim ZERO_STR As String
Dim ONE_STR As String
Dim TWO_STR As String
Dim THREE_STR As String
Dim FOUR_STR As String

Dim TEMP_STR As String
Dim DATA_MATRIX As Variant
Dim TEMP_VECTOR As Variant

Dim TEMP_SUM As Double

On Error GoTo ERROR_LABEL

DATA_MATRIX = MATRIX_REMOVE_ROWS_FUNC(MATRIX_REMOVE_COLUMNS_FUNC(DATA_RNG, 1, 1), 1, 1)
TEMP_VECTOR = MATRIX_REMOVE_COLUMNS_FUNC(MATRIX_GET_ROW_FUNC(DATA_RNG, 1, 1), 1, 1)
NSIZE = 3

If UBound(DATA_MATRIX, 1) < NSIZE Then GoTo ERROR_LABEL
If UBound(DATA_MATRIX, 2) < NSIZE Then GoTo ERROR_LABEL

'For any given number of currencies ‘N’, the number of cross-rates
'theoretically possible is N x (N-1)/2. The number of variations
'possible with N currencies to construct a “round-trip” would be NN,
'which means that the number of possibilities increases exponentially
'as we increase the number of currencies being considered.

'Note that each of the intervening currencies CC1 to 6 can be any of
'the 6 currencies, there is no limitation that a currency trade cannot
'be done twice, therefore the number of possible combinations is 6 x 6
'x 6 x 6 x 6 x 6, or 66. This is important because it allows us to look
'at profitable combinations where the number of currencies is less than
'six.

If IsNumeric(UNIT_TRANSACTION_COST) = False Then UNIT_TRANSACTION_COST = 0

'Use six nested “For…Next” loops to cycle through each possible combination
'of currencies. Each time, calculate the profit that would result from
'executing the given combination of currency transactions, and if this
'profit is greater than the profit from a previous combination, then store
'it away in a variable, otherwise discard it. (The initial value of profit is 0)
'If a profitable combination is not found – which for instance will be the case
'where transaction costs are very high, the macro says so. If a profitable
'combination is found, the same is listed together with all combinations.

For C1 = 1 To NSIZE
For C2 = 1 To NSIZE
For c3 = 1 To NSIZE

'Update the status bar to show the combination being checked.
'      Excel.Application.StatusBar = _
 '     TEMP_VECTOR(1, c1) & "-" & _
  '    TEMP_VECTOR(1, c2) & "-" & _
      'TEMP_VECTOR(1, c3)

'Calculate true number of transactions. When the same _
'currency is repeated, then it is not a transaction, for _
'example USD->GBP->GBP->JPY->USD = USD->GBP->JPY->USD, ie _
'4 transactions and not 5, as GBP->GBP is costless. This is
'needed to calculate true transaction costs.

j = NSIZE + 1

If 1 = C1 Then j = j - 1
If C1 = C2 Then j = j - 1
If C2 = c3 Then j = j - 1
If c3 = 1 Then j = j - 1

'For the currency combination in the loop, compare _
'TARGET_PROFIT to previous highest profits.

If TARGET_PROFIT < ((INITIAL_INVESTMENT * _
    DATA_MATRIX(C1, 1) * _
    DATA_MATRIX(C2, C1) * _
    DATA_MATRIX(c3, C2) * _
    DATA_MATRIX(1, c3) - _
    INITIAL_INVESTMENT) - (j * UNIT_TRANSACTION_COST)) _
Then
    
    DST_ZERO = 1
    DST_ONE = C1
    DST_TWO = C2
    DST_THREE = c3
    DST_FOUR = 1
    
    k = 1
    
    TARGET_PROFIT = ((INITIAL_INVESTMENT * _
    DATA_MATRIX(C1, 1) * _
    DATA_MATRIX(C2, C1) * _
    DATA_MATRIX(c3, C2) * _
    DATA_MATRIX(1, c3) - _
    INITIAL_INVESTMENT) - (j * UNIT_TRANSACTION_COST))
    
    h = j
End If

'ENHANCING THE MODEL:

'The model can be extended beyond spot transactions to look at forward
'transactions, interest rates, and identify any opportunities that may
'arise from an imbalance between the spot and forward rates given the
'interest rates in the two currencies. Such arbitrage would be based
'on violations of interest rate parity.

'In economies where foreign exchange markets are regulated, there often
'exists a black market exchange rate that is different from the official
'rate. Arbitrage opportunities between these rates are always possible
'and there is some money to be made. However, this may involve some risk
'taking in the form of possible violations of foreign exchange regulations.
'The cost of these violations with the probability of being discovered
'will need to be factored into the model.

TEMP_SUM = TEMP_SUM + 1   'merely a count of how many combinations checked.

Next c3
Next C2
Next C1

'By now, the most profitable opportunity has been identified. If there _
'is no profitable opportunity, then the "GoSub finito" would not have _
'been visited and k would be zero)

If k = 0 Then
    CURRENCIES_TRIANGLE_THREE_FUNC = "No profitable possibilities!"
    Exit Function
Else
    
    ZERO_STR = TEMP_VECTOR(1, 1)
    'Replace currency numbers by their respective English
    ONE_STR = TEMP_VECTOR(1, DST_ONE)       'codes, eg 1 means USD etc
    TWO_STR = TEMP_VECTOR(1, DST_TWO)
    THREE_STR = TEMP_VECTOR(1, DST_THREE)
    FOUR_STR = TEMP_VECTOR(1, 1)
    
    
    For i = 1 To NSIZE + 1  'Remove duplicates, ie "GBP --> GBP" is _
    merely "GBP" in the transaction.
        If ZERO_STR = ONE_STR Then ZERO_STR = ""
        If ONE_STR = TWO_STR Then ONE_STR = ""
        If TWO_STR = THREE_STR Then TWO_STR = ""
        If THREE_STR = FOUR_STR Then THREE_STR = ""
    Next i
End If

Select Case OUTPUT
Case 0
    
    TEMP_STR = "Transact currencies as follows: " & _
        ZERO_STR & "-" & _
        ONE_STR & "-" & _
        TWO_STR & "-" & _
        THREE_STR & "-" & _
        FOUR_STR & _
        " Profit on a " & _
        Format(INITIAL_INVESTMENT, "#,0.00") & " = " & _
        Format(TARGET_PROFIT, "#,0.00") & _
        "  (" & TEMP_SUM & " possibilities checked )"
        
        CURRENCIES_TRIANGLE_THREE_FUNC = TEMP_STR
        
Case Else

    ReDim TEMP_MATRIX(0 To NSIZE + 4, 1 To 6)
    
    TEMP_MATRIX(0, 1) = "Transaction"
    TEMP_MATRIX(0, 2) = "Details"
    TEMP_MATRIX(0, 3) = "Amount sold"
    TEMP_MATRIX(0, 4) = "Exchange rate"
    TEMP_MATRIX(0, 5) = "Details"
    TEMP_MATRIX(0, 6) = "Amount bought"

    If ZERO_STR = "" Then ZERO_STR = BASE_FX_STR
    If ONE_STR = "" Then ONE_STR = BASE_FX_STR
    If TWO_STR = "" Then TWO_STR = BASE_FX_STR
    If THREE_STR = "" Then THREE_STR = BASE_FX_STR
    If FOUR_STR = "" Then FOUR_STR = BASE_FX_STR
    
    'Write the transaction
    TEMP_MATRIX(1, 1) = "Sell " & ZERO_STR & ", Buy " & ONE_STR
    TEMP_MATRIX(2, 1) = "Sell " & ONE_STR & ", Buy " & TWO_STR
    TEMP_MATRIX(3, 1) = "Sell " & TWO_STR & ", Buy " & THREE_STR
    TEMP_MATRIX(4, 1) = "Sell " & THREE_STR & ", Buy " & FOUR_STR
    
    'Write the transaction details for currency sold
    TEMP_MATRIX(1, 2) = ZERO_STR & " sold:"
    TEMP_MATRIX(2, 2) = ONE_STR & " sold:"
    TEMP_MATRIX(3, 2) = TWO_STR & " sold:"
    TEMP_MATRIX(4, 2) = THREE_STR & " sold:"
    
    'Write the transaction details for currency bought
    TEMP_MATRIX(1, 5) = ONE_STR & " bought:"
    TEMP_MATRIX(2, 5) = TWO_STR & " bought:"
    TEMP_MATRIX(3, 5) = THREE_STR & " bought:"
    TEMP_MATRIX(4, 5) = FOUR_STR & " bought:"
    
    'MATRIX_FIND_ELEMENT_FUNC
    
    C0 = MATRIX_FIND_ELEMENT_FUNC(TEMP_VECTOR, ZERO_STR, 1, 1, 0)
    C1 = MATRIX_FIND_ELEMENT_FUNC(TEMP_VECTOR, ONE_STR, 1, 1, 0)
    C2 = MATRIX_FIND_ELEMENT_FUNC(TEMP_VECTOR, TWO_STR, 1, 1, 0)
    c3 = MATRIX_FIND_ELEMENT_FUNC(TEMP_VECTOR, THREE_STR, 1, 1, 0)
    c4 = MATRIX_FIND_ELEMENT_FUNC(TEMP_VECTOR, FOUR_STR, 1, 1, 0)
    
    TEMP_MATRIX(1, 4) = DATA_MATRIX(C1, C0)
    
    TEMP_MATRIX(2, 4) = DATA_MATRIX(C2, C1)
    
    TEMP_MATRIX(3, 4) = DATA_MATRIX(c3, C2)
    
    TEMP_MATRIX(4, 4) = DATA_MATRIX(c4, c3)
    
        
    TEMP_MATRIX(1, 3) = INITIAL_INVESTMENT
    
    TEMP_MATRIX(2, 3) = TEMP_MATRIX(1, 3) * TEMP_MATRIX(1, 4)
    TEMP_MATRIX(3, 3) = TEMP_MATRIX(2, 3) * TEMP_MATRIX(2, 4)
    TEMP_MATRIX(4, 3) = TEMP_MATRIX(3, 3) * TEMP_MATRIX(3, 4)
    
    TEMP_MATRIX(1, 6) = TEMP_MATRIX(1, 3) * TEMP_MATRIX(1, 4)
    TEMP_MATRIX(2, 6) = TEMP_MATRIX(2, 3) * TEMP_MATRIX(2, 4)
    TEMP_MATRIX(3, 6) = TEMP_MATRIX(3, 3) * TEMP_MATRIX(3, 4)
    TEMP_MATRIX(4, 6) = TEMP_MATRIX(4, 3) * TEMP_MATRIX(4, 4)

    TEMP_MATRIX(NSIZE + 2, 1) = "Gross Profit:"
    TEMP_MATRIX(NSIZE + 2, 2) = "Net: " & BASE_FX_STR
    TEMP_MATRIX(NSIZE + 2, 3) = TEMP_MATRIX(NSIZE + 1, 6) - INITIAL_INVESTMENT
    
    TEMP_MATRIX(NSIZE + 2, 4) = ""
    TEMP_MATRIX(NSIZE + 2, 5) = ""
    TEMP_MATRIX(NSIZE + 2, 6) = ""
    
    TEMP_MATRIX(NSIZE + 3, 1) = "Less: Transaction Cost (" & _
    Format(h, "0.00") & _
    " * $" & (Format(UNIT_TRANSACTION_COST, "0.00")) & ")"
    TEMP_MATRIX(NSIZE + 3, 2) = ""

    TEMP_MATRIX(NSIZE + 3, 3) = -UNIT_TRANSACTION_COST * h
    
    TEMP_MATRIX(NSIZE + 3, 4) = ""
    TEMP_MATRIX(NSIZE + 3, 5) = ""
    TEMP_MATRIX(NSIZE + 3, 6) = ""
    
    TEMP_MATRIX(NSIZE + 4, 1) = "Net Profit"
    TEMP_MATRIX(NSIZE + 4, 2) = ""
    
    TEMP_MATRIX(NSIZE + 4, 3) = TEMP_MATRIX(NSIZE + 3, 3) + _
    TEMP_MATRIX(NSIZE + 2, 3)

    TEMP_MATRIX(NSIZE + 4, 4) = ""
    TEMP_MATRIX(NSIZE + 4, 5) = ""
    TEMP_MATRIX(NSIZE + 4, 6) = ""

    CURRENCIES_TRIANGLE_THREE_FUNC = TEMP_MATRIX
End Select

Exit Function
ERROR_LABEL:
CURRENCIES_TRIANGLE_THREE_FUNC = Err.number
End Function


Private Function CURRENCIES_TRIANGLE_FOUR_FUNC(ByRef DATA_RNG As Variant, _
Optional ByVal INITIAL_INVESTMENT As Double = 15000000, _
Optional ByVal TARGET_PROFIT As Double = 0, _
Optional ByVal UNIT_TRANSACTION_COST As Double = 20, _
Optional ByVal BASE_FX_STR As String = "USD", _
Optional ByVal OUTPUT As Integer = 1)

'DATA_RNG = Sanitised cross currency exchange rates
'where, lower triangle = bid rates, upper triangle = offer rates
'TARGET_PROFIT: initial value for profit

'UNIT_TRANSACTION_COST: transaction costs: enter transaction cost per
'transaction in base currency.

'TARGET_PROFIT: TARGET_PROFIT_PROFIT

'--------------------------------Cross rates-------------------------------
'Most currencies are expressed against dollars – which means there is
'always a buy-sell spread that would normally make it difficult to make
'money by routing transactions through dollars. However, there are some
'cross rates that trade directly. A cross rate is a rate between two
'non US Dollar currencies. For example, there may be trading
'between British Pounds and the Euro. Therefore, we have a rate for the
'Dollar-Euro, we have a rate for the Dollar-Pound, and we also have a rate
'for Pound-Euro.
'--------------------------------------------------------------------------

Dim h As Long
Dim i As Long
Dim j As Long
Dim k As Long

Dim C0 As Long
Dim C1 As Long
Dim C2 As Long
Dim c3 As Long
Dim c4 As Long
Dim c5 As Long

Dim NSIZE As Long

Dim DST_ZERO As Long
Dim DST_ONE As Long
Dim DST_TWO As Long
Dim DST_THREE As Long
Dim DST_FOUR As Long
Dim DST_FIVE As Long

Dim ZERO_STR As String
Dim ONE_STR As String
Dim TWO_STR As String
Dim THREE_STR As String
Dim FOUR_STR As String
Dim FIVE_STR As String

Dim TEMP_STR As String
Dim DATA_MATRIX As Variant
Dim TEMP_VECTOR As Variant
Dim TEMP_SUM As Double

On Error GoTo ERROR_LABEL

DATA_MATRIX = MATRIX_REMOVE_ROWS_FUNC(MATRIX_REMOVE_COLUMNS_FUNC(DATA_RNG, 1, 1), 1, 1)
TEMP_VECTOR = MATRIX_REMOVE_COLUMNS_FUNC(MATRIX_GET_ROW_FUNC(DATA_RNG, 1, 1), 1, 1)
NSIZE = 4

If UBound(DATA_MATRIX, 1) < NSIZE Then GoTo ERROR_LABEL
If UBound(DATA_MATRIX, 2) < NSIZE Then GoTo ERROR_LABEL

'For any given number of currencies ‘N’, the number of cross-rates
'theoretically possible is N x (N-1)/2. The number of variations
'possible with N currencies to construct a “round-trip” would be NN,
'which means that the number of possibilities increases exponentially
'as we increase the number of currencies being considered.

'Note that each of the intervening currencies CC1 to 6 can be any of
'the 6 currencies, there is no limitation that a currency trade cannot
'be done twice, therefore the number of possible combinations is 6 x 6
'x 6 x 6 x 6 x 6, or 66. This is important because it allows us to look
'at profitable combinations where the number of currencies is less than
'six.

If IsNumeric(UNIT_TRANSACTION_COST) = False Then UNIT_TRANSACTION_COST = 0

'Use six nested “For…Next” loops to cycle through each possible combination
'of currencies. Each time, calculate the profit that would result from
'executing the given combination of currency transactions, and if this
'profit is greater than the profit from a previous combination, then store
'it away in a variable, otherwise discard it. (The initial value of profit is 0)
'If a profitable combination is not found – which for instance will be the case
'where transaction costs are very high, the macro says so. If a profitable
'combination is found, the same is listed together with all combinations.

For C1 = 1 To NSIZE
For C2 = 1 To NSIZE
For c3 = 1 To NSIZE
For c4 = 1 To NSIZE

'Update the status bar to show the combination being checked.
'      Excel.Application.StatusBar = _
 '     TEMP_VECTOR(1, c1) & "-" & _
  '    TEMP_VECTOR(1, c2) & "-" & _
   '   TEMP_VECTOR(1, c3) & "-" & _
      'TEMP_VECTOR(1, c4)

'Calculate true number of transactions. When the same _
'currency is repeated, then it is not a transaction, for _
'example USD->GBP->GBP->JPY->USD = USD->GBP->JPY->USD, ie _
'4 transactions and not 5, as GBP->GBP is costless. This is
'needed to calculate true transaction costs.

j = NSIZE + 1

If 1 = C1 Then j = j - 1
If C1 = C2 Then j = j - 1
If C2 = c3 Then j = j - 1
If c3 = c4 Then j = j - 1
If c4 = 1 Then j = j - 1

'For the currency combination in the loop, compare _
'TARGET_PROFIT to previous highest profits.

If TARGET_PROFIT < ((INITIAL_INVESTMENT * _
    DATA_MATRIX(C1, 1) * _
    DATA_MATRIX(C2, C1) * _
    DATA_MATRIX(c3, C2) * _
    DATA_MATRIX(c4, c3) * _
    DATA_MATRIX(1, c4) - _
    INITIAL_INVESTMENT) - (j * UNIT_TRANSACTION_COST)) _
Then
    
    DST_ZERO = 1
    DST_ONE = C1
    DST_TWO = C2
    DST_THREE = c3
    DST_FOUR = c4
    DST_FIVE = 1
    
    k = 1
    
    TARGET_PROFIT = ((INITIAL_INVESTMENT * _
    DATA_MATRIX(C1, 1) * _
    DATA_MATRIX(C2, C1) * _
    DATA_MATRIX(c3, C2) * _
    DATA_MATRIX(c4, c3) * _
    DATA_MATRIX(1, c4) - _
    INITIAL_INVESTMENT) - (j * UNIT_TRANSACTION_COST))
    
    h = j
End If

'ENHANCING THE MODEL:

'The model can be extended beyond spot transactions to look at forward
'transactions, interest rates, and identify any opportunities that may
'arise from an imbalance between the spot and forward rates given the
'interest rates in the two currencies. Such arbitrage would be based
'on violations of interest rate parity.

'In economies where foreign exchange markets are regulated, there often
'exists a black market exchange rate that is different from the official
'rate. Arbitrage opportunities between these rates are always possible
'and there is some money to be made. However, this may involve some risk
'taking in the form of possible violations of foreign exchange regulations.
'The cost of these violations with the probability of being discovered
'will need to be factored into the model.

TEMP_SUM = TEMP_SUM + 1   'merely a count of how many combinations checked.

Next c4
Next c3
Next C2
Next C1

'By now, the most profitable opportunity has been identified. If there _
'is no profitable opportunity, then the "GoSub finito" would not have _
'been visited and k would be zero)

If k = 0 Then
    CURRENCIES_TRIANGLE_FOUR_FUNC = "No profitable possibilities!"
    Exit Function
Else
    
    ZERO_STR = TEMP_VECTOR(1, 1)
    'Replace currency numbers by their respective English
    ONE_STR = TEMP_VECTOR(1, DST_ONE)       'codes, eg 1 means USD etc
    TWO_STR = TEMP_VECTOR(1, DST_TWO)
    THREE_STR = TEMP_VECTOR(1, DST_THREE)
    FOUR_STR = TEMP_VECTOR(1, DST_FOUR)
    FIVE_STR = TEMP_VECTOR(1, 1)
    
    
    For i = 1 To NSIZE + 1  'Remove duplicates, ie "GBP --> GBP" is _
    merely "GBP" in the transaction.
        If ZERO_STR = ONE_STR Then ZERO_STR = ""
        If ONE_STR = TWO_STR Then ONE_STR = ""
        If TWO_STR = THREE_STR Then TWO_STR = ""
        If THREE_STR = FOUR_STR Then THREE_STR = ""
        If FOUR_STR = FIVE_STR Then FOUR_STR = ""
    Next i
End If

Select Case OUTPUT
Case 0
    
    TEMP_STR = "Transact currencies as follows: " & _
        ZERO_STR & "-" & _
        ONE_STR & "-" & _
        TWO_STR & "-" & _
        THREE_STR & "-" & _
        FOUR_STR & "-" & _
        FIVE_STR & _
        " Profit on a " & _
        Format(INITIAL_INVESTMENT, "#,0.00") & " = " & _
        Format(TARGET_PROFIT, "#,0.00") & _
        "  (" & TEMP_SUM & " possibilities checked )"
        
        CURRENCIES_TRIANGLE_FOUR_FUNC = TEMP_STR
        
Case Else

    ReDim TEMP_MATRIX(0 To NSIZE + 4, 1 To 6)
    
    TEMP_MATRIX(0, 1) = "Transaction"
    TEMP_MATRIX(0, 2) = "Details"
    TEMP_MATRIX(0, 3) = "Amount sold"
    TEMP_MATRIX(0, 4) = "Exchange rate"
    TEMP_MATRIX(0, 5) = "Details"
    TEMP_MATRIX(0, 6) = "Amount bought"

    If ZERO_STR = "" Then ZERO_STR = BASE_FX_STR
    If ONE_STR = "" Then ONE_STR = BASE_FX_STR
    If TWO_STR = "" Then TWO_STR = BASE_FX_STR
    If THREE_STR = "" Then THREE_STR = BASE_FX_STR
    If FOUR_STR = "" Then FOUR_STR = BASE_FX_STR
    If FIVE_STR = "" Then FIVE_STR = BASE_FX_STR
    
    'Write the transaction
    TEMP_MATRIX(1, 1) = "Sell " & ZERO_STR & ", Buy " & ONE_STR
    TEMP_MATRIX(2, 1) = "Sell " & ONE_STR & ", Buy " & TWO_STR
    TEMP_MATRIX(3, 1) = "Sell " & TWO_STR & ", Buy " & THREE_STR
    TEMP_MATRIX(4, 1) = "Sell " & THREE_STR & ", Buy " & FOUR_STR
    TEMP_MATRIX(5, 1) = "Sell " & FOUR_STR & ", Buy " & FIVE_STR
    
    'Write the transaction details for currency sold
    TEMP_MATRIX(1, 2) = ZERO_STR & " sold:"
    TEMP_MATRIX(2, 2) = ONE_STR & " sold:"
    TEMP_MATRIX(3, 2) = TWO_STR & " sold:"
    TEMP_MATRIX(4, 2) = THREE_STR & " sold:"
    TEMP_MATRIX(5, 2) = FOUR_STR & " sold:"
    
    'Write the transaction details for currency bought
    TEMP_MATRIX(1, 5) = ONE_STR & " bought:"
    TEMP_MATRIX(2, 5) = TWO_STR & " bought:"
    TEMP_MATRIX(3, 5) = THREE_STR & " bought:"
    TEMP_MATRIX(4, 5) = FOUR_STR & " bought:"
    TEMP_MATRIX(5, 5) = FIVE_STR & " bought:"
    
    'MATRIX_FIND_ELEMENT_FUNC
    
    C0 = MATRIX_FIND_ELEMENT_FUNC(TEMP_VECTOR, ZERO_STR, 1, 1, 0)
    C1 = MATRIX_FIND_ELEMENT_FUNC(TEMP_VECTOR, ONE_STR, 1, 1, 0)
    C2 = MATRIX_FIND_ELEMENT_FUNC(TEMP_VECTOR, TWO_STR, 1, 1, 0)
    c3 = MATRIX_FIND_ELEMENT_FUNC(TEMP_VECTOR, THREE_STR, 1, 1, 0)
    c4 = MATRIX_FIND_ELEMENT_FUNC(TEMP_VECTOR, FOUR_STR, 1, 1, 0)
    c5 = MATRIX_FIND_ELEMENT_FUNC(TEMP_VECTOR, FIVE_STR, 1, 1, 0)
    
    TEMP_MATRIX(1, 4) = DATA_MATRIX(C1, C0)
    
    TEMP_MATRIX(2, 4) = DATA_MATRIX(C2, C1)
    
    TEMP_MATRIX(3, 4) = DATA_MATRIX(c3, C2)
    
    TEMP_MATRIX(4, 4) = DATA_MATRIX(c4, c3)
    
    TEMP_MATRIX(5, 4) = DATA_MATRIX(c5, c4)
    
        
    TEMP_MATRIX(1, 3) = INITIAL_INVESTMENT
    
    TEMP_MATRIX(2, 3) = TEMP_MATRIX(1, 3) * TEMP_MATRIX(1, 4)
    TEMP_MATRIX(3, 3) = TEMP_MATRIX(2, 3) * TEMP_MATRIX(2, 4)
    TEMP_MATRIX(4, 3) = TEMP_MATRIX(3, 3) * TEMP_MATRIX(3, 4)
    TEMP_MATRIX(5, 3) = TEMP_MATRIX(4, 3) * TEMP_MATRIX(4, 4)
    
    TEMP_MATRIX(1, 6) = TEMP_MATRIX(1, 3) * TEMP_MATRIX(1, 4)
    TEMP_MATRIX(2, 6) = TEMP_MATRIX(2, 3) * TEMP_MATRIX(2, 4)
    TEMP_MATRIX(3, 6) = TEMP_MATRIX(3, 3) * TEMP_MATRIX(3, 4)
    TEMP_MATRIX(4, 6) = TEMP_MATRIX(4, 3) * TEMP_MATRIX(4, 4)
    TEMP_MATRIX(5, 6) = TEMP_MATRIX(5, 3) * TEMP_MATRIX(5, 4)

    TEMP_MATRIX(NSIZE + 2, 1) = "Gross Profit:"
    TEMP_MATRIX(NSIZE + 2, 2) = "Net: " & BASE_FX_STR
    TEMP_MATRIX(NSIZE + 2, 3) = TEMP_MATRIX(NSIZE + 1, 6) - INITIAL_INVESTMENT
    
    TEMP_MATRIX(NSIZE + 2, 4) = ""
    TEMP_MATRIX(NSIZE + 2, 5) = ""
    TEMP_MATRIX(NSIZE + 2, 6) = ""
    
    TEMP_MATRIX(NSIZE + 3, 1) = "Less: Transaction Cost (" & _
    Format(h, "0.00") & _
    " * $" & (Format(UNIT_TRANSACTION_COST, "0.00")) & ")"
    TEMP_MATRIX(NSIZE + 3, 2) = ""

    TEMP_MATRIX(NSIZE + 3, 3) = -UNIT_TRANSACTION_COST * h
    
    TEMP_MATRIX(NSIZE + 3, 4) = ""
    TEMP_MATRIX(NSIZE + 3, 5) = ""
    TEMP_MATRIX(NSIZE + 3, 6) = ""
    
    TEMP_MATRIX(NSIZE + 4, 1) = "Net Profit"
    TEMP_MATRIX(NSIZE + 4, 2) = ""
    
    TEMP_MATRIX(NSIZE + 4, 3) = TEMP_MATRIX(NSIZE + 3, 3) + _
    TEMP_MATRIX(NSIZE + 2, 3)

    TEMP_MATRIX(NSIZE + 4, 4) = ""
    TEMP_MATRIX(NSIZE + 4, 5) = ""
    TEMP_MATRIX(NSIZE + 4, 6) = ""

    CURRENCIES_TRIANGLE_FOUR_FUNC = TEMP_MATRIX
End Select

Exit Function
ERROR_LABEL:
CURRENCIES_TRIANGLE_FOUR_FUNC = Err.number
End Function


Private Function CURRENCIES_TRIANGLE_FIVE_FUNC(ByRef DATA_RNG As Variant, _
Optional ByVal INITIAL_INVESTMENT As Double = 15000000, _
Optional ByVal TARGET_PROFIT As Double = 0, _
Optional ByVal UNIT_TRANSACTION_COST As Double = 20, _
Optional ByVal BASE_FX_STR As String = "USD", _
Optional ByVal OUTPUT As Integer = 1)

'DATA_RNG = Sanitised cross currency exchange rates
'where, lower triangle = bid rates, upper triangle = offer rates
'TARGET_PROFIT: initial value for profit

'UNIT_TRANSACTION_COST: transaction costs: enter transaction cost per
'transaction in base currency.

'TARGET_PROFIT: TARGET_PROFIT_PROFIT

'--------------------------------Cross rates-------------------------------
'Most currencies are expressed against dollars – which means there is
'always a buy-sell spread that would normally make it difficult to make
'money by routing transactions through dollars. However, there are some
'cross rates that trade directly. A cross rate is a rate between two
'non US Dollar currencies. For example, there may be trading
'between British Pounds and the Euro. Therefore, we have a rate for the
'Dollar-Euro, we have a rate for the Dollar-Pound, and we also have a rate
'for Pound-Euro.
'--------------------------------------------------------------------------

Dim h As Long
Dim i As Long
Dim j As Long
Dim k As Long

Dim C0 As Long
Dim C1 As Long
Dim C2 As Long
Dim c3 As Long
Dim c4 As Long
Dim c5 As Long
Dim c6 As Long

Dim NSIZE As Long

Dim DST_ZERO As Long
Dim DST_ONE As Long
Dim DST_TWO As Long
Dim DST_THREE As Long
Dim DST_FOUR As Long
Dim DST_FIVE As Long
Dim DST_SIX As Long

Dim ZERO_STR As String
Dim ONE_STR As String
Dim TWO_STR As String
Dim THREE_STR As String
Dim FOUR_STR As String
Dim FIVE_STR As String
Dim SIX_STR As String

Dim TEMP_STR As String
Dim DATA_MATRIX As Variant
Dim TEMP_VECTOR As Variant

Dim TEMP_SUM As Double

On Error GoTo ERROR_LABEL

DATA_MATRIX = MATRIX_REMOVE_ROWS_FUNC(MATRIX_REMOVE_COLUMNS_FUNC(DATA_RNG, 1, 1), 1, 1)
TEMP_VECTOR = MATRIX_REMOVE_COLUMNS_FUNC(MATRIX_GET_ROW_FUNC(DATA_RNG, 1, 1), 1, 1)
NSIZE = 5

If UBound(DATA_MATRIX, 1) < NSIZE Then GoTo ERROR_LABEL
If UBound(DATA_MATRIX, 2) < NSIZE Then GoTo ERROR_LABEL

'For any given number of currencies ‘N’, the number of cross-rates
'theoretically possible is N x (N-1)/2. The number of variations
'possible with N currencies to construct a “round-trip” would be NN,
'which means that the number of possibilities increases exponentially
'as we increase the number of currencies being considered.

'Note that each of the intervening currencies CC1 to 6 can be any of
'the 6 currencies, there is no limitation that a currency trade cannot
'be done twice, therefore the number of possible combinations is 6 x 6
'x 6 x 6 x 6 x 6, or 66. This is important because it allows us to look
'at profitable combinations where the number of currencies is less than
'six.

If IsNumeric(UNIT_TRANSACTION_COST) = False Then UNIT_TRANSACTION_COST = 0

'Use six nested “For…Next” loops to cycle through each possible combination
'of currencies. Each time, calculate the profit that would result from
'executing the given combination of currency transactions, and if this
'profit is greater than the profit from a previous combination, then store
'it away in a variable, otherwise discard it. (The initial value of profit is 0)
'If a profitable combination is not found – which for instance will be the case
'where transaction costs are very high, the macro says so. If a profitable
'combination is found, the same is listed together with all combinations.

For C1 = 1 To NSIZE
For C2 = 1 To NSIZE
For c3 = 1 To NSIZE
For c4 = 1 To NSIZE
For c5 = 1 To NSIZE

'Update the status bar to show the combination being checked.
'      Excel.Application.StatusBar = _
 '     TEMP_VECTOR(1, c1) & "-" & _
  '    TEMP_VECTOR(1, c2) & "-" & _
   '   TEMP_VECTOR(1, c3) & "-" & _
    '  TEMP_VECTOR(1, c4) & "-" & _
      'TEMP_VECTOR(1, c5)

'Calculate true number of transactions. When the same _
'currency is repeated, then it is not a transaction, for _
'example USD->GBP->GBP->JPY->USD = USD->GBP->JPY->USD, ie _
'4 transactions and not 5, as GBP->GBP is costless. This is
'needed to calculate true transaction costs.

j = NSIZE + 1

If 1 = C1 Then j = j - 1
If C1 = C2 Then j = j - 1
If C2 = c3 Then j = j - 1
If c3 = c4 Then j = j - 1
If c4 = c5 Then j = j - 1
If c5 = 1 Then j = j - 1

'For the currency combination in the loop, compare _
'TARGET_PROFIT to previous highest profits.

If TARGET_PROFIT < ((INITIAL_INVESTMENT * _
    DATA_MATRIX(C1, 1) * _
    DATA_MATRIX(C2, C1) * _
    DATA_MATRIX(c3, C2) * _
    DATA_MATRIX(c4, c3) * _
    DATA_MATRIX(c5, c4) * _
    DATA_MATRIX(1, c5) - _
    INITIAL_INVESTMENT) - (j * UNIT_TRANSACTION_COST)) _
Then
    
    DST_ZERO = 1
    DST_ONE = C1
    DST_TWO = C2
    DST_THREE = c3
    DST_FOUR = c4
    DST_FIVE = c5
    DST_SIX = 1
    
    k = 1
    
    TARGET_PROFIT = ((INITIAL_INVESTMENT * _
    DATA_MATRIX(C1, 1) * _
    DATA_MATRIX(C2, C1) * _
    DATA_MATRIX(c3, C2) * _
    DATA_MATRIX(c4, c3) * _
    DATA_MATRIX(c5, c4) * _
    DATA_MATRIX(1, c5) - _
    INITIAL_INVESTMENT) - (j * UNIT_TRANSACTION_COST))
    
    h = j
End If

'ENHANCING THE MODEL:

'The model can be extended beyond spot transactions to look at forward
'transactions, interest rates, and identify any opportunities that may
'arise from an imbalance between the spot and forward rates given the
'interest rates in the two currencies. Such arbitrage would be based
'on violations of interest rate parity.

'In economies where foreign exchange markets are regulated, there often
'exists a black market exchange rate that is different from the official
'rate. Arbitrage opportunities between these rates are always possible
'and there is some money to be made. However, this may involve some risk
'taking in the form of possible violations of foreign exchange regulations.
'The cost of these violations with the probability of being discovered
'will need to be factored into the model.

TEMP_SUM = TEMP_SUM + 1   'merely a count of how many combinations checked.

Next c5
Next c4
Next c3
Next C2
Next C1

'By now, the most profitable opportunity has been identified. If there _
'is no profitable opportunity, then the "GoSub finito" would not have _
'been visited and k would be zero)

If k = 0 Then
    CURRENCIES_TRIANGLE_FIVE_FUNC = "No profitable possibilities!"
    Exit Function
Else
    
    ZERO_STR = TEMP_VECTOR(1, 1)
    'Replace currency numbers by their respective English
    ONE_STR = TEMP_VECTOR(1, DST_ONE)       'codes, eg 1 means USD etc
    TWO_STR = TEMP_VECTOR(1, DST_TWO)
    THREE_STR = TEMP_VECTOR(1, DST_THREE)
    FOUR_STR = TEMP_VECTOR(1, DST_FOUR)
    FIVE_STR = TEMP_VECTOR(1, DST_FIVE)
    SIX_STR = TEMP_VECTOR(1, 1)
    
    
    For i = 1 To NSIZE + 1  'Remove duplicates, ie "GBP --> GBP" is _
    merely "GBP" in the transaction.
        If ZERO_STR = ONE_STR Then ZERO_STR = ""
        If ONE_STR = TWO_STR Then ONE_STR = ""
        If TWO_STR = THREE_STR Then TWO_STR = ""
        If THREE_STR = FOUR_STR Then THREE_STR = ""
        If FOUR_STR = FIVE_STR Then FOUR_STR = ""
        If FIVE_STR = SIX_STR Then FIVE_STR = ""
    Next i
End If

Select Case OUTPUT
Case 0
    
    TEMP_STR = "Transact currencies as follows: " & _
        ZERO_STR & "-" & _
        ONE_STR & "-" & _
        TWO_STR & "-" & _
        THREE_STR & "-" & _
        FOUR_STR & "-" & _
        FIVE_STR & "-" & _
        SIX_STR & _
        " Profit on a " & _
        Format(INITIAL_INVESTMENT, "#,0.00") & " = " & _
        Format(TARGET_PROFIT, "#,0.00") & _
        "  (" & TEMP_SUM & " possibilities checked )"
        
        CURRENCIES_TRIANGLE_FIVE_FUNC = TEMP_STR
        
Case Else

    ReDim TEMP_MATRIX(0 To NSIZE + 4, 1 To 6)
    
    TEMP_MATRIX(0, 1) = "Transaction"
    TEMP_MATRIX(0, 2) = "Details"
    TEMP_MATRIX(0, 3) = "Amount sold"
    TEMP_MATRIX(0, 4) = "Exchange rate"
    TEMP_MATRIX(0, 5) = "Details"
    TEMP_MATRIX(0, 6) = "Amount bought"

    If ZERO_STR = "" Then ZERO_STR = BASE_FX_STR
    If ONE_STR = "" Then ONE_STR = BASE_FX_STR
    If TWO_STR = "" Then TWO_STR = BASE_FX_STR
    If THREE_STR = "" Then THREE_STR = BASE_FX_STR
    If FOUR_STR = "" Then FOUR_STR = BASE_FX_STR
    If FIVE_STR = "" Then FIVE_STR = BASE_FX_STR
    If SIX_STR = "" Then SIX_STR = BASE_FX_STR
    
    'Write the transaction
    TEMP_MATRIX(1, 1) = "Sell " & ZERO_STR & ", Buy " & ONE_STR
    TEMP_MATRIX(2, 1) = "Sell " & ONE_STR & ", Buy " & TWO_STR
    TEMP_MATRIX(3, 1) = "Sell " & TWO_STR & ", Buy " & THREE_STR
    TEMP_MATRIX(4, 1) = "Sell " & THREE_STR & ", Buy " & FOUR_STR
    TEMP_MATRIX(5, 1) = "Sell " & FOUR_STR & ", Buy " & FIVE_STR
    TEMP_MATRIX(6, 1) = "Sell " & FIVE_STR & ", Buy " & SIX_STR
    
    'Write the transaction details for currency sold
    TEMP_MATRIX(1, 2) = ZERO_STR & " sold:"
    TEMP_MATRIX(2, 2) = ONE_STR & " sold:"
    TEMP_MATRIX(3, 2) = TWO_STR & " sold:"
    TEMP_MATRIX(4, 2) = THREE_STR & " sold:"
    TEMP_MATRIX(5, 2) = FOUR_STR & " sold:"
    TEMP_MATRIX(6, 2) = FIVE_STR & " sold:"
    
    'Write the transaction details for currency bought
    TEMP_MATRIX(1, 5) = ONE_STR & " bought:"
    TEMP_MATRIX(2, 5) = TWO_STR & " bought:"
    TEMP_MATRIX(3, 5) = THREE_STR & " bought:"
    TEMP_MATRIX(4, 5) = FOUR_STR & " bought:"
    TEMP_MATRIX(5, 5) = FIVE_STR & " bought:"
    TEMP_MATRIX(6, 5) = SIX_STR & " bought:"
    
    'MATRIX_FIND_ELEMENT_FUNC
    
    C0 = MATRIX_FIND_ELEMENT_FUNC(TEMP_VECTOR, ZERO_STR, 1, 1, 0)
    C1 = MATRIX_FIND_ELEMENT_FUNC(TEMP_VECTOR, ONE_STR, 1, 1, 0)
    C2 = MATRIX_FIND_ELEMENT_FUNC(TEMP_VECTOR, TWO_STR, 1, 1, 0)
    c3 = MATRIX_FIND_ELEMENT_FUNC(TEMP_VECTOR, THREE_STR, 1, 1, 0)
    c4 = MATRIX_FIND_ELEMENT_FUNC(TEMP_VECTOR, FOUR_STR, 1, 1, 0)
    c5 = MATRIX_FIND_ELEMENT_FUNC(TEMP_VECTOR, FIVE_STR, 1, 1, 0)
    c6 = MATRIX_FIND_ELEMENT_FUNC(TEMP_VECTOR, SIX_STR, 1, 1, 0)
    
    TEMP_MATRIX(1, 4) = DATA_MATRIX(C1, C0)
    
    TEMP_MATRIX(2, 4) = DATA_MATRIX(C2, C1)
    
    TEMP_MATRIX(3, 4) = DATA_MATRIX(c3, C2)
    
    TEMP_MATRIX(4, 4) = DATA_MATRIX(c4, c3)
    
    TEMP_MATRIX(5, 4) = DATA_MATRIX(c5, c4)
    
    TEMP_MATRIX(6, 4) = DATA_MATRIX(c6, c5)
        
    TEMP_MATRIX(1, 3) = INITIAL_INVESTMENT
    
    TEMP_MATRIX(2, 3) = TEMP_MATRIX(1, 3) * TEMP_MATRIX(1, 4)
    TEMP_MATRIX(3, 3) = TEMP_MATRIX(2, 3) * TEMP_MATRIX(2, 4)
    TEMP_MATRIX(4, 3) = TEMP_MATRIX(3, 3) * TEMP_MATRIX(3, 4)
    TEMP_MATRIX(5, 3) = TEMP_MATRIX(4, 3) * TEMP_MATRIX(4, 4)
    TEMP_MATRIX(6, 3) = TEMP_MATRIX(5, 3) * TEMP_MATRIX(5, 4)
    
    TEMP_MATRIX(1, 6) = TEMP_MATRIX(1, 3) * TEMP_MATRIX(1, 4)
    TEMP_MATRIX(2, 6) = TEMP_MATRIX(2, 3) * TEMP_MATRIX(2, 4)
    TEMP_MATRIX(3, 6) = TEMP_MATRIX(3, 3) * TEMP_MATRIX(3, 4)
    TEMP_MATRIX(4, 6) = TEMP_MATRIX(4, 3) * TEMP_MATRIX(4, 4)
    TEMP_MATRIX(5, 6) = TEMP_MATRIX(5, 3) * TEMP_MATRIX(5, 4)
    TEMP_MATRIX(6, 6) = TEMP_MATRIX(6, 3) * TEMP_MATRIX(6, 4)

    TEMP_MATRIX(NSIZE + 2, 1) = "Gross Profit:"
    TEMP_MATRIX(NSIZE + 2, 2) = "Net: " & BASE_FX_STR
    TEMP_MATRIX(NSIZE + 2, 3) = TEMP_MATRIX(NSIZE + 1, 6) - INITIAL_INVESTMENT
    
    TEMP_MATRIX(NSIZE + 2, 4) = ""
    TEMP_MATRIX(NSIZE + 2, 5) = ""
    TEMP_MATRIX(NSIZE + 2, 6) = ""
    
    TEMP_MATRIX(NSIZE + 3, 1) = "Less: Transaction Cost (" & _
    Format(h, "0.00") & _
    " * $" & (Format(UNIT_TRANSACTION_COST, "0.00")) & ")"
    TEMP_MATRIX(NSIZE + 3, 2) = ""

    TEMP_MATRIX(NSIZE + 3, 3) = -UNIT_TRANSACTION_COST * h
    
    TEMP_MATRIX(NSIZE + 3, 4) = ""
    TEMP_MATRIX(NSIZE + 3, 5) = ""
    TEMP_MATRIX(NSIZE + 3, 6) = ""
    
    TEMP_MATRIX(NSIZE + 4, 1) = "Net Profit"
    TEMP_MATRIX(NSIZE + 4, 2) = ""
    
    TEMP_MATRIX(NSIZE + 4, 3) = TEMP_MATRIX(NSIZE + 3, 3) + _
    TEMP_MATRIX(NSIZE + 2, 3)

    TEMP_MATRIX(NSIZE + 4, 4) = ""
    TEMP_MATRIX(NSIZE + 4, 5) = ""
    TEMP_MATRIX(NSIZE + 4, 6) = ""

    CURRENCIES_TRIANGLE_FIVE_FUNC = TEMP_MATRIX
End Select

Exit Function
ERROR_LABEL:
CURRENCIES_TRIANGLE_FIVE_FUNC = Err.number
End Function


Private Function CURRENCIES_TRIANGLE_SIX_FUNC(ByRef DATA_RNG As Variant, _
Optional ByVal INITIAL_INVESTMENT As Double = 1000000, _
Optional ByVal TARGET_PROFIT As Double = 0, _
Optional ByVal UNIT_TRANSACTION_COST As Double = 20, _
Optional ByVal BASE_FX_STR As String = "USD", _
Optional ByVal OUTPUT As Integer = 1)

'DATA_RNG = Sanitised cross currency exchange rates
'where, lower triangle = bid rates, upper triangle = offer rates
'TARGET_PROFIT: initial value for profit

'UNIT_TRANSACTION_COST: transaction costs: enter transaction cost per
'transaction in base currency.

'TARGET_PROFIT: TARGET_PROFIT_PROFIT

'--------------------------------Cross rates-------------------------------
'Most currencies are expressed against dollars – which means there is
'always a buy-sell spread that would normally make it difficult to make
'money by routing transactions through dollars. However, there are some
'cross rates that trade directly. A cross rate is a rate between two
'non US Dollar currencies. For example, there may be trading
'between British Pounds and the Euro. Therefore, we have a rate for the
'Dollar-Euro, we have a rate for the Dollar-Pound, and we also have a rate
'for Pound-Euro.
'--------------------------------------------------------------------------

Dim h As Long
Dim i As Long
Dim j As Long
Dim k As Long

Dim C0 As Long
Dim C1 As Long
Dim C2 As Long
Dim c3 As Long
Dim c4 As Long
Dim c5 As Long
Dim c6 As Long
Dim c7 As Long

Dim NSIZE As Long

Dim DST_ZERO As Long
Dim DST_ONE As Long
Dim DST_TWO As Long
Dim DST_THREE As Long
Dim DST_FOUR As Long
Dim DST_FIVE As Long
Dim DST_SIX As Long
Dim DST_SEVEN As Long

Dim ZERO_STR As String
Dim ONE_STR As String
Dim TWO_STR As String
Dim THREE_STR As String
Dim FOUR_STR As String
Dim FIVE_STR As String
Dim SIX_STR As String
Dim SEVEN_STR As String

Dim TEMP_STR As String
Dim DATA_MATRIX As Variant
Dim TEMP_VECTOR As Variant
Dim TEMP_SUM As Double

On Error GoTo ERROR_LABEL

DATA_MATRIX = MATRIX_REMOVE_ROWS_FUNC(MATRIX_REMOVE_COLUMNS_FUNC(DATA_RNG, 1, 1), 1, 1)
TEMP_VECTOR = MATRIX_REMOVE_COLUMNS_FUNC(MATRIX_GET_ROW_FUNC(DATA_RNG, 1, 1), 1, 1)
NSIZE = 6

If UBound(DATA_MATRIX, 1) < NSIZE Then GoTo ERROR_LABEL
If UBound(DATA_MATRIX, 2) < NSIZE Then GoTo ERROR_LABEL

'For any given number of currencies ‘N’, the number of cross-rates
'theoretically possible is N x (N-1)/2. The number of variations
'possible with N currencies to construct a “round-trip” would be NN,
'which means that the number of possibilities increases exponentially
'as we increase the number of currencies being considered.

'Note that each of the intervening currencies CC1 to 6 can be any of
'the 6 currencies, there is no limitation that a currency trade cannot
'be done twice, therefore the number of possible combinations is 6 x 6
'x 6 x 6 x 6 x 6, or 66. This is important because it allows us to look
'at profitable combinations where the number of currencies is less than
'six.

If IsNumeric(UNIT_TRANSACTION_COST) = False Then UNIT_TRANSACTION_COST = 0

'Use six nested “For…Next” loops to cycle through each possible combination
'of currencies. Each time, calculate the profit that would result from
'executing the given combination of currency transactions, and if this
'profit is greater than the profit from a previous combination, then store
'it away in a variable, otherwise discard it. (The initial value of profit is 0)
'If a profitable combination is not found – which for instance will be the case
'where transaction costs are very high, the macro says so. If a profitable
'combination is found, the same is listed together with all combinations.

For C1 = 1 To NSIZE
For C2 = 1 To NSIZE
For c3 = 1 To NSIZE
For c4 = 1 To NSIZE
For c5 = 1 To NSIZE
For c6 = 1 To NSIZE

'Update the status bar to show the combination being checked.
'      Excel.Application.StatusBar = _
 '     TEMP_VECTOR(1, c1) & "-" & _
  '    TEMP_VECTOR(1, c2) & "-" & _
   '   TEMP_VECTOR(1, c3) & "-" & _
    '  TEMP_VECTOR(1, c4) & "-" & _
     ' TEMP_VECTOR(1, c5) & "-" & _
      'TEMP_VECTOR(1, c6)

'Calculate true number of transactions. When the same _
'currency is repeated, then it is not a transaction, for _
'example USD->GBP->GBP->JPY->USD = USD->GBP->JPY->USD, ie _
'4 transactions and not 5, as GBP->GBP is costless. This is
'needed to calculate true transaction costs.

j = NSIZE + 1

If 1 = C1 Then j = j - 1
If C1 = C2 Then j = j - 1
If C2 = c3 Then j = j - 1
If c3 = c4 Then j = j - 1
If c4 = c5 Then j = j - 1
If c5 = c6 Then j = j - 1
If c6 = 1 Then j = j - 1

'For the currency combination in the loop, compare _
'TARGET_PROFIT to previous highest profits.

If TARGET_PROFIT < ((INITIAL_INVESTMENT * _
    DATA_MATRIX(C1, 1) * _
    DATA_MATRIX(C2, C1) * _
    DATA_MATRIX(c3, C2) * _
    DATA_MATRIX(c4, c3) * _
    DATA_MATRIX(c5, c4) * _
    DATA_MATRIX(c6, c5) * _
    DATA_MATRIX(1, c6) - _
    INITIAL_INVESTMENT) - (j * UNIT_TRANSACTION_COST)) _
Then
    
    DST_ZERO = 1
    DST_ONE = C1
    DST_TWO = C2
    DST_THREE = c3
    DST_FOUR = c4
    DST_FIVE = c5
    DST_SIX = c6
    DST_SEVEN = 1
    
    k = 1
    
    TARGET_PROFIT = ((INITIAL_INVESTMENT * _
    DATA_MATRIX(C1, 1) * _
    DATA_MATRIX(C2, C1) * _
    DATA_MATRIX(c3, C2) * _
    DATA_MATRIX(c4, c3) * _
    DATA_MATRIX(c5, c4) * _
    DATA_MATRIX(c6, c5) * _
    DATA_MATRIX(1, c6) - _
    INITIAL_INVESTMENT) - (j * UNIT_TRANSACTION_COST))
    
    h = j
End If

'ENHANCING THE MODEL:

'The model can be extended beyond spot transactions to look at forward
'transactions, interest rates, and identify any opportunities that may
'arise from an imbalance between the spot and forward rates given the
'interest rates in the two currencies. Such arbitrage would be based
'on violations of interest rate parity.

'In economies where foreign exchange markets are regulated, there often
'exists a black market exchange rate that is different from the official
'rate. Arbitrage opportunities between these rates are always possible
'and there is some money to be made. However, this may involve some risk
'taking in the form of possible violations of foreign exchange regulations.
'The cost of these violations with the probability of being discovered
'will need to be factored into the model.

TEMP_SUM = TEMP_SUM + 1   'merely a count of how many combinations checked.

Next c6
Next c5
Next c4
Next c3
Next C2
Next C1

'By now, the most profitable opportunity has been identified. If there _
'is no profitable opportunity, then the "GoSub finito" would not have _
'been visited and k would be zero)

If k = 0 Then
    CURRENCIES_TRIANGLE_SIX_FUNC = "No profitable possibilities!"
    Exit Function
Else
    
    ZERO_STR = TEMP_VECTOR(1, 1)
    'Replace currency numbers by their respective English
    ONE_STR = TEMP_VECTOR(1, DST_ONE)       'codes, eg 1 means USD etc
    TWO_STR = TEMP_VECTOR(1, DST_TWO)
    THREE_STR = TEMP_VECTOR(1, DST_THREE)
    FOUR_STR = TEMP_VECTOR(1, DST_FOUR)
    FIVE_STR = TEMP_VECTOR(1, DST_FIVE)
    SIX_STR = TEMP_VECTOR(1, DST_SIX)
    SEVEN_STR = TEMP_VECTOR(1, 1)
    
    
    For i = 1 To NSIZE + 1  'Remove duplicates, ie "GBP --> GBP" is _
    merely "GBP" in the transaction.
        If ZERO_STR = ONE_STR Then ZERO_STR = ""
        If ONE_STR = TWO_STR Then ONE_STR = ""
        If TWO_STR = THREE_STR Then TWO_STR = ""
        If THREE_STR = FOUR_STR Then THREE_STR = ""
        If FOUR_STR = FIVE_STR Then FOUR_STR = ""
        If FIVE_STR = SIX_STR Then FIVE_STR = ""
        If SIX_STR = SEVEN_STR Then SIX_STR = ""
    Next i
End If

Select Case OUTPUT
Case 0
    
    TEMP_STR = "Transact currencies as follows: " & _
        ZERO_STR & "-" & _
        ONE_STR & "-" & _
        TWO_STR & "-" & _
        THREE_STR & "-" & _
        FOUR_STR & "-" & _
        FIVE_STR & "-" & _
        SIX_STR & "-" & _
        SEVEN_STR & _
        " Profit on a " & _
        Format(INITIAL_INVESTMENT, "#,0.00") & " = " & _
        Format(TARGET_PROFIT, "#,0.00") & _
        "  (" & TEMP_SUM & " possibilities checked )"
        
        CURRENCIES_TRIANGLE_SIX_FUNC = TEMP_STR
        
Case Else

    ReDim TEMP_MATRIX(0 To NSIZE + 4, 1 To 6)
    
    TEMP_MATRIX(0, 1) = "Transaction"
    TEMP_MATRIX(0, 2) = "Details"
    TEMP_MATRIX(0, 3) = "Amount sold"
    TEMP_MATRIX(0, 4) = "Exchange rate"
    TEMP_MATRIX(0, 5) = "Details"
    TEMP_MATRIX(0, 6) = "Amount bought"

    If ZERO_STR = "" Then ZERO_STR = BASE_FX_STR
    If ONE_STR = "" Then ONE_STR = BASE_FX_STR
    If TWO_STR = "" Then TWO_STR = BASE_FX_STR
    If THREE_STR = "" Then THREE_STR = BASE_FX_STR
    If FOUR_STR = "" Then FOUR_STR = BASE_FX_STR
    If FIVE_STR = "" Then FIVE_STR = BASE_FX_STR
    If SIX_STR = "" Then SIX_STR = BASE_FX_STR
    If SEVEN_STR = "" Then SEVEN_STR = BASE_FX_STR
    
    'Write the transaction
    TEMP_MATRIX(1, 1) = "Sell " & ZERO_STR & ", Buy " & ONE_STR
    TEMP_MATRIX(2, 1) = "Sell " & ONE_STR & ", Buy " & TWO_STR
    TEMP_MATRIX(3, 1) = "Sell " & TWO_STR & ", Buy " & THREE_STR
    TEMP_MATRIX(4, 1) = "Sell " & THREE_STR & ", Buy " & FOUR_STR
    TEMP_MATRIX(5, 1) = "Sell " & FOUR_STR & ", Buy " & FIVE_STR
    TEMP_MATRIX(6, 1) = "Sell " & FIVE_STR & ", Buy " & SIX_STR
    TEMP_MATRIX(7, 1) = "Sell " & SIX_STR & ", Buy " & SEVEN_STR
    
    'Write the transaction details for currency sold
    TEMP_MATRIX(1, 2) = ZERO_STR & " sold:"
    TEMP_MATRIX(2, 2) = ONE_STR & " sold:"
    TEMP_MATRIX(3, 2) = TWO_STR & " sold:"
    TEMP_MATRIX(4, 2) = THREE_STR & " sold:"
    TEMP_MATRIX(5, 2) = FOUR_STR & " sold:"
    TEMP_MATRIX(6, 2) = FIVE_STR & " sold:"
    TEMP_MATRIX(7, 2) = SIX_STR & " sold:"
    
    'Write the transaction details for currency bought
    TEMP_MATRIX(1, 5) = ONE_STR & " bought:"
    TEMP_MATRIX(2, 5) = TWO_STR & " bought:"
    TEMP_MATRIX(3, 5) = THREE_STR & " bought:"
    TEMP_MATRIX(4, 5) = FOUR_STR & " bought:"
    TEMP_MATRIX(5, 5) = FIVE_STR & " bought:"
    TEMP_MATRIX(6, 5) = SIX_STR & " bought:"
    TEMP_MATRIX(7, 5) = SEVEN_STR & " bought:"
    
    'MATRIX_FIND_ELEMENT_FUNC
    
    C0 = MATRIX_FIND_ELEMENT_FUNC(TEMP_VECTOR, ZERO_STR, 1, 1, 0)
    C1 = MATRIX_FIND_ELEMENT_FUNC(TEMP_VECTOR, ONE_STR, 1, 1, 0)
    C2 = MATRIX_FIND_ELEMENT_FUNC(TEMP_VECTOR, TWO_STR, 1, 1, 0)
    c3 = MATRIX_FIND_ELEMENT_FUNC(TEMP_VECTOR, THREE_STR, 1, 1, 0)
    c4 = MATRIX_FIND_ELEMENT_FUNC(TEMP_VECTOR, FOUR_STR, 1, 1, 0)
    c5 = MATRIX_FIND_ELEMENT_FUNC(TEMP_VECTOR, FIVE_STR, 1, 1, 0)
    c6 = MATRIX_FIND_ELEMENT_FUNC(TEMP_VECTOR, SIX_STR, 1, 1, 0)
    c7 = MATRIX_FIND_ELEMENT_FUNC(TEMP_VECTOR, SEVEN_STR, 1, 1, 0)
    
    TEMP_MATRIX(1, 4) = DATA_MATRIX(C1, C0)
    
    TEMP_MATRIX(2, 4) = DATA_MATRIX(C2, C1)
    
    TEMP_MATRIX(3, 4) = DATA_MATRIX(c3, C2)
    
    TEMP_MATRIX(4, 4) = DATA_MATRIX(c4, c3)
    
    TEMP_MATRIX(5, 4) = DATA_MATRIX(c5, c4)
    
    TEMP_MATRIX(6, 4) = DATA_MATRIX(c6, c5)
    
    TEMP_MATRIX(7, 4) = DATA_MATRIX(c7, c6)
        
    TEMP_MATRIX(1, 3) = INITIAL_INVESTMENT
    
    TEMP_MATRIX(2, 3) = TEMP_MATRIX(1, 3) * TEMP_MATRIX(1, 4)
    TEMP_MATRIX(3, 3) = TEMP_MATRIX(2, 3) * TEMP_MATRIX(2, 4)
    TEMP_MATRIX(4, 3) = TEMP_MATRIX(3, 3) * TEMP_MATRIX(3, 4)
    TEMP_MATRIX(5, 3) = TEMP_MATRIX(4, 3) * TEMP_MATRIX(4, 4)
    TEMP_MATRIX(6, 3) = TEMP_MATRIX(5, 3) * TEMP_MATRIX(5, 4)
    TEMP_MATRIX(7, 3) = TEMP_MATRIX(6, 3) * TEMP_MATRIX(6, 4)
    
    TEMP_MATRIX(1, 6) = TEMP_MATRIX(1, 3) * TEMP_MATRIX(1, 4)
    TEMP_MATRIX(2, 6) = TEMP_MATRIX(2, 3) * TEMP_MATRIX(2, 4)
    TEMP_MATRIX(3, 6) = TEMP_MATRIX(3, 3) * TEMP_MATRIX(3, 4)
    TEMP_MATRIX(4, 6) = TEMP_MATRIX(4, 3) * TEMP_MATRIX(4, 4)
    TEMP_MATRIX(5, 6) = TEMP_MATRIX(5, 3) * TEMP_MATRIX(5, 4)
    TEMP_MATRIX(6, 6) = TEMP_MATRIX(6, 3) * TEMP_MATRIX(6, 4)
    TEMP_MATRIX(7, 6) = TEMP_MATRIX(7, 3) * TEMP_MATRIX(7, 4)

    TEMP_MATRIX(NSIZE + 2, 1) = "Gross Profit:"
    TEMP_MATRIX(NSIZE + 2, 2) = "Net: " & BASE_FX_STR
    TEMP_MATRIX(NSIZE + 2, 3) = TEMP_MATRIX(NSIZE + 1, 6) - INITIAL_INVESTMENT
    
    TEMP_MATRIX(NSIZE + 2, 4) = ""
    TEMP_MATRIX(NSIZE + 2, 5) = ""
    TEMP_MATRIX(NSIZE + 2, 6) = ""
    
    TEMP_MATRIX(NSIZE + 3, 1) = "Less: Transaction Cost (" & _
    Format(h, "0.00") & _
    " * $" & (Format(UNIT_TRANSACTION_COST, "0.00")) & ")"
    TEMP_MATRIX(NSIZE + 3, 2) = ""

    TEMP_MATRIX(NSIZE + 3, 3) = -UNIT_TRANSACTION_COST * h
    
    TEMP_MATRIX(NSIZE + 3, 4) = ""
    TEMP_MATRIX(NSIZE + 3, 5) = ""
    TEMP_MATRIX(NSIZE + 3, 6) = ""
    
    TEMP_MATRIX(NSIZE + 4, 1) = "Net Profit"
    TEMP_MATRIX(NSIZE + 4, 2) = ""
    
    TEMP_MATRIX(NSIZE + 4, 3) = TEMP_MATRIX(NSIZE + 3, 3) + _
    TEMP_MATRIX(NSIZE + 2, 3)

    TEMP_MATRIX(NSIZE + 4, 4) = ""
    TEMP_MATRIX(NSIZE + 4, 5) = ""
    TEMP_MATRIX(NSIZE + 4, 6) = ""

    CURRENCIES_TRIANGLE_SIX_FUNC = TEMP_MATRIX
End Select

Exit Function
ERROR_LABEL:
CURRENCIES_TRIANGLE_SIX_FUNC = Err.number
End Function


Private Function CURRENCIES_TRIANGLE_SEVEN_FUNC(ByRef DATA_RNG As Variant, _
Optional ByVal INITIAL_INVESTMENT As Double = 15000000, _
Optional ByVal TARGET_PROFIT As Double = 0, _
Optional ByVal UNIT_TRANSACTION_COST As Double = 20, _
Optional ByVal BASE_FX_STR As String = "USD", _
Optional ByVal OUTPUT As Integer = 1)

'DATA_RNG = Sanitised cross currency exchange rates
'where, lower triangle = bid rates, upper triangle = offer rates
'TARGET_PROFIT: initial value for profit

'UNIT_TRANSACTION_COST: transaction costs: enter transaction cost per
'transaction in base currency.

'TARGET_PROFIT: TARGET_PROFIT_PROFIT

'--------------------------------Cross rates-------------------------------
'Most currencies are expressed against dollars – which means there is
'always a buy-sell spread that would normally make it difficult to make
'money by routing transactions through dollars. However, there are some
'cross rates that trade directly. A cross rate is a rate between two
'non US Dollar currencies. For example, there may be trading
'between British Pounds and the Euro. Therefore, we have a rate for the
'Dollar-Euro, we have a rate for the Dollar-Pound, and we also have a rate
'for Pound-Euro.
'--------------------------------------------------------------------------

Dim h As Long
Dim i As Long
Dim j As Long
Dim k As Long

Dim C0 As Long
Dim C1 As Long
Dim C2 As Long
Dim c3 As Long
Dim c4 As Long
Dim c5 As Long
Dim c6 As Long
Dim c7 As Long
Dim c8 As Long

Dim NSIZE As Long

Dim DST_ZERO As Long
Dim DST_ONE As Long
Dim DST_TWO As Long
Dim DST_THREE As Long
Dim DST_FOUR As Long
Dim DST_FIVE As Long
Dim DST_SIX As Long
Dim DST_SEVEN As Long
Dim DST_EIGHT As Long

Dim ZERO_STR As String
Dim ONE_STR As String
Dim TWO_STR As String
Dim THREE_STR As String
Dim FOUR_STR As String
Dim FIVE_STR As String
Dim SIX_STR As String
Dim SEVEN_STR As String
Dim EIGHT_STR As String

Dim TEMP_STR As String
Dim DATA_MATRIX As Variant
Dim TEMP_VECTOR As Variant
Dim TEMP_SUM As Double

On Error GoTo ERROR_LABEL

DATA_MATRIX = MATRIX_REMOVE_ROWS_FUNC(MATRIX_REMOVE_COLUMNS_FUNC(DATA_RNG, 1, 1), 1, 1)
TEMP_VECTOR = MATRIX_REMOVE_COLUMNS_FUNC(MATRIX_GET_ROW_FUNC(DATA_RNG, 1, 1), 1, 1)
NSIZE = 7

If UBound(DATA_MATRIX, 1) < NSIZE Then GoTo ERROR_LABEL
If UBound(DATA_MATRIX, 2) < NSIZE Then GoTo ERROR_LABEL

'For any given number of currencies ‘N’, the number of cross-rates
'theoretically possible is N x (N-1)/2. The number of variations
'possible with N currencies to construct a “round-trip” would be NN,
'which means that the number of possibilities increases exponentially
'as we increase the number of currencies being considered.

'Note that each of the intervening currencies CC1 to 6 can be any of
'the 6 currencies, there is no limitation that a currency trade cannot
'be done twice, therefore the number of possible combinations is 6 x 6
'x 6 x 6 x 6 x 6, or 66. This is important because it allows us to look
'at profitable combinations where the number of currencies is less than
'six.

If IsNumeric(UNIT_TRANSACTION_COST) = False Then UNIT_TRANSACTION_COST = 0

'Use six nested “For…Next” loops to cycle through each possible combination
'of currencies. Each time, calculate the profit that would result from
'executing the given combination of currency transactions, and if this
'profit is greater than the profit from a previous combination, then store
'it away in a variable, otherwise discard it. (The initial value of profit is 0)
'If a profitable combination is not found – which for instance will be the case
'where transaction costs are very high, the macro says so. If a profitable
'combination is found, the same is listed together with all combinations.

For C1 = 1 To NSIZE
For C2 = 1 To NSIZE
For c3 = 1 To NSIZE
For c4 = 1 To NSIZE
For c5 = 1 To NSIZE
For c6 = 1 To NSIZE
For c7 = 1 To NSIZE

'Update the status bar to show the combination being checked.
'      Excel.Application.StatusBar = _
 '     TEMP_VECTOR(1, c1) & "-" & _
  '    TEMP_VECTOR(1, c2) & "-" & _
   '   TEMP_VECTOR(1, c3) & "-" & _
    '  TEMP_VECTOR(1, c4) & "-" & _
     ' TEMP_VECTOR(1, c5) & "-" & _
     ' TEMP_VECTOR(1, c6) & "-" & _
      'TEMP_VECTOR(1, c7)

'Calculate true number of transactions. When the same _
'currency is repeated, then it is not a transaction, for _
'example USD->GBP->GBP->JPY->USD = USD->GBP->JPY->USD, ie _
'4 transactions and not 5, as GBP->GBP is costless. This is
'needed to calculate true transaction costs.

j = NSIZE + 1

If 1 = C1 Then j = j - 1
If C1 = C2 Then j = j - 1
If C2 = c3 Then j = j - 1
If c3 = c4 Then j = j - 1
If c4 = c5 Then j = j - 1
If c5 = c6 Then j = j - 1
If c6 = c7 Then j = j - 1
If c7 = 1 Then j = j - 1

'For the currency combination in the loop, compare _
'TARGET_PROFIT to previous highest profits.

If TARGET_PROFIT < ((INITIAL_INVESTMENT * _
    DATA_MATRIX(C1, 1) * _
    DATA_MATRIX(C2, C1) * _
    DATA_MATRIX(c3, C2) * _
    DATA_MATRIX(c4, c3) * _
    DATA_MATRIX(c5, c4) * _
    DATA_MATRIX(c6, c5) * _
    DATA_MATRIX(c7, c6) * _
    DATA_MATRIX(1, c7) - _
    INITIAL_INVESTMENT) - (j * UNIT_TRANSACTION_COST)) _
Then
    
    DST_ZERO = 1
    DST_ONE = C1
    DST_TWO = C2
    DST_THREE = c3
    DST_FOUR = c4
    DST_FIVE = c5
    DST_SIX = c6
    DST_SEVEN = c7
    DST_EIGHT = 1
    
    k = 1
    
    TARGET_PROFIT = ((INITIAL_INVESTMENT * _
    DATA_MATRIX(C1, 1) * _
    DATA_MATRIX(C2, C1) * _
    DATA_MATRIX(c3, C2) * _
    DATA_MATRIX(c4, c3) * _
    DATA_MATRIX(c5, c4) * _
    DATA_MATRIX(c6, c5) * _
    DATA_MATRIX(c7, c6) * _
    DATA_MATRIX(1, c7) - _
    INITIAL_INVESTMENT) - (j * UNIT_TRANSACTION_COST))
    
    h = j
End If

'ENHANCING THE MODEL:

'The model can be extended beyond spot transactions to look at forward
'transactions, interest rates, and identify any opportunities that may
'arise from an imbalance between the spot and forward rates given the
'interest rates in the two currencies. Such arbitrage would be based
'on violations of interest rate parity.

'In economies where foreign exchange markets are regulated, there often
'exists a black market exchange rate that is different from the official
'rate. Arbitrage opportunities between these rates are always possible
'and there is some money to be made. However, this may involve some risk
'taking in the form of possible violations of foreign exchange regulations.
'The cost of these violations with the probability of being discovered
'will need to be factored into the model.

TEMP_SUM = TEMP_SUM + 1   'merely a count of how many combinations checked.

Next c7
Next c6
Next c5
Next c4
Next c3
Next C2
Next C1

'By now, the most profitable opportunity has been identified. If there _
'is no profitable opportunity, then the "GoSub finito" would not have _
'been visited and k would be zero)

If k = 0 Then
    CURRENCIES_TRIANGLE_SEVEN_FUNC = "No profitable possibilities!"
    Exit Function
Else
    
    ZERO_STR = TEMP_VECTOR(1, 1)
    'Replace currency numbers by their respective English
    ONE_STR = TEMP_VECTOR(1, DST_ONE)       'codes, eg 1 means USD etc
    TWO_STR = TEMP_VECTOR(1, DST_TWO)
    THREE_STR = TEMP_VECTOR(1, DST_THREE)
    FOUR_STR = TEMP_VECTOR(1, DST_FOUR)
    FIVE_STR = TEMP_VECTOR(1, DST_FIVE)
    SIX_STR = TEMP_VECTOR(1, DST_SIX)
    SEVEN_STR = TEMP_VECTOR(1, DST_SEVEN)
    EIGHT_STR = TEMP_VECTOR(1, 1)
    
    
    For i = 1 To NSIZE + 1  'Remove duplicates, ie "GBP --> GBP" is _
    merely "GBP" in the transaction.
        If ZERO_STR = ONE_STR Then ZERO_STR = ""
        If ONE_STR = TWO_STR Then ONE_STR = ""
        If TWO_STR = THREE_STR Then TWO_STR = ""
        If THREE_STR = FOUR_STR Then THREE_STR = ""
        If FOUR_STR = FIVE_STR Then FOUR_STR = ""
        If FIVE_STR = SIX_STR Then FIVE_STR = ""
        If SIX_STR = SEVEN_STR Then SIX_STR = ""
        If SEVEN_STR = EIGHT_STR Then SEVEN_STR = ""
    Next i
End If

Select Case OUTPUT
Case 0
    
    TEMP_STR = "Transact currencies as follows: " & _
        ZERO_STR & "-" & _
        ONE_STR & "-" & _
        TWO_STR & "-" & _
        THREE_STR & "-" & _
        FOUR_STR & "-" & _
        FIVE_STR & "-" & _
        SIX_STR & "-" & _
        SEVEN_STR & "-" & _
        EIGHT_STR & _
        " Profit on a " & _
        Format(INITIAL_INVESTMENT, "#,0.00") & " = " & _
        Format(TARGET_PROFIT, "#,0.00") & _
        "  (" & TEMP_SUM & " possibilities checked )"
        
        CURRENCIES_TRIANGLE_SEVEN_FUNC = TEMP_STR
        
Case Else

    ReDim TEMP_MATRIX(0 To NSIZE + 4, 1 To 6)
    
    TEMP_MATRIX(0, 1) = "Transaction"
    TEMP_MATRIX(0, 2) = "Details"
    TEMP_MATRIX(0, 3) = "Amount sold"
    TEMP_MATRIX(0, 4) = "Exchange rate"
    TEMP_MATRIX(0, 5) = "Details"
    TEMP_MATRIX(0, 6) = "Amount bought"

    If ZERO_STR = "" Then ZERO_STR = BASE_FX_STR
    If ONE_STR = "" Then ONE_STR = BASE_FX_STR
    If TWO_STR = "" Then TWO_STR = BASE_FX_STR
    If THREE_STR = "" Then THREE_STR = BASE_FX_STR
    If FOUR_STR = "" Then FOUR_STR = BASE_FX_STR
    If FIVE_STR = "" Then FIVE_STR = BASE_FX_STR
    If SIX_STR = "" Then SIX_STR = BASE_FX_STR
    If SEVEN_STR = "" Then SEVEN_STR = BASE_FX_STR
    If EIGHT_STR = "" Then EIGHT_STR = BASE_FX_STR
    
    'Write the transaction
    TEMP_MATRIX(1, 1) = "Sell " & ZERO_STR & ", Buy " & ONE_STR
    TEMP_MATRIX(2, 1) = "Sell " & ONE_STR & ", Buy " & TWO_STR
    TEMP_MATRIX(3, 1) = "Sell " & TWO_STR & ", Buy " & THREE_STR
    TEMP_MATRIX(4, 1) = "Sell " & THREE_STR & ", Buy " & FOUR_STR
    TEMP_MATRIX(5, 1) = "Sell " & FOUR_STR & ", Buy " & FIVE_STR
    TEMP_MATRIX(6, 1) = "Sell " & FIVE_STR & ", Buy " & SIX_STR
    TEMP_MATRIX(7, 1) = "Sell " & SIX_STR & ", Buy " & SEVEN_STR
    TEMP_MATRIX(8, 1) = "Sell " & SEVEN_STR & ", Buy " & EIGHT_STR
    
    'Write the transaction details for currency sold
    TEMP_MATRIX(1, 2) = ZERO_STR & " sold:"
    TEMP_MATRIX(2, 2) = ONE_STR & " sold:"
    TEMP_MATRIX(3, 2) = TWO_STR & " sold:"
    TEMP_MATRIX(4, 2) = THREE_STR & " sold:"
    TEMP_MATRIX(5, 2) = FOUR_STR & " sold:"
    TEMP_MATRIX(6, 2) = FIVE_STR & " sold:"
    TEMP_MATRIX(7, 2) = SIX_STR & " sold:"
    TEMP_MATRIX(8, 2) = SEVEN_STR & " sold:"
    
    'Write the transaction details for currency bought
    TEMP_MATRIX(1, 5) = ONE_STR & " bought:"
    TEMP_MATRIX(2, 5) = TWO_STR & " bought:"
    TEMP_MATRIX(3, 5) = THREE_STR & " bought:"
    TEMP_MATRIX(4, 5) = FOUR_STR & " bought:"
    TEMP_MATRIX(5, 5) = FIVE_STR & " bought:"
    TEMP_MATRIX(6, 5) = SIX_STR & " bought:"
    TEMP_MATRIX(7, 5) = SEVEN_STR & " bought:"
    TEMP_MATRIX(8, 5) = EIGHT_STR & " bought:"
    
    'MATRIX_FIND_ELEMENT_FUNC
    
    C0 = MATRIX_FIND_ELEMENT_FUNC(TEMP_VECTOR, ZERO_STR, 1, 1, 0)
    C1 = MATRIX_FIND_ELEMENT_FUNC(TEMP_VECTOR, ONE_STR, 1, 1, 0)
    C2 = MATRIX_FIND_ELEMENT_FUNC(TEMP_VECTOR, TWO_STR, 1, 1, 0)
    c3 = MATRIX_FIND_ELEMENT_FUNC(TEMP_VECTOR, THREE_STR, 1, 1, 0)
    c4 = MATRIX_FIND_ELEMENT_FUNC(TEMP_VECTOR, FOUR_STR, 1, 1, 0)
    c5 = MATRIX_FIND_ELEMENT_FUNC(TEMP_VECTOR, FIVE_STR, 1, 1, 0)
    c6 = MATRIX_FIND_ELEMENT_FUNC(TEMP_VECTOR, SIX_STR, 1, 1, 0)
    c7 = MATRIX_FIND_ELEMENT_FUNC(TEMP_VECTOR, SEVEN_STR, 1, 1, 0)
    c8 = MATRIX_FIND_ELEMENT_FUNC(TEMP_VECTOR, EIGHT_STR, 1, 1, 0)
    
    TEMP_MATRIX(1, 4) = DATA_MATRIX(C1, C0)
    
    TEMP_MATRIX(2, 4) = DATA_MATRIX(C2, C1)
    
    TEMP_MATRIX(3, 4) = DATA_MATRIX(c3, C2)
    
    TEMP_MATRIX(4, 4) = DATA_MATRIX(c4, c3)
    
    TEMP_MATRIX(5, 4) = DATA_MATRIX(c5, c4)
    
    TEMP_MATRIX(6, 4) = DATA_MATRIX(c6, c5)
    
    TEMP_MATRIX(7, 4) = DATA_MATRIX(c7, c6)
    
    TEMP_MATRIX(8, 4) = DATA_MATRIX(c8, c7)
        
    TEMP_MATRIX(1, 3) = INITIAL_INVESTMENT
    
    TEMP_MATRIX(2, 3) = TEMP_MATRIX(1, 3) * TEMP_MATRIX(1, 4)
    TEMP_MATRIX(3, 3) = TEMP_MATRIX(2, 3) * TEMP_MATRIX(2, 4)
    TEMP_MATRIX(4, 3) = TEMP_MATRIX(3, 3) * TEMP_MATRIX(3, 4)
    TEMP_MATRIX(5, 3) = TEMP_MATRIX(4, 3) * TEMP_MATRIX(4, 4)
    TEMP_MATRIX(6, 3) = TEMP_MATRIX(5, 3) * TEMP_MATRIX(5, 4)
    TEMP_MATRIX(7, 3) = TEMP_MATRIX(6, 3) * TEMP_MATRIX(6, 4)
    TEMP_MATRIX(8, 3) = TEMP_MATRIX(7, 3) * TEMP_MATRIX(7, 4)
    
    TEMP_MATRIX(1, 6) = TEMP_MATRIX(1, 3) * TEMP_MATRIX(1, 4)
    TEMP_MATRIX(2, 6) = TEMP_MATRIX(2, 3) * TEMP_MATRIX(2, 4)
    TEMP_MATRIX(3, 6) = TEMP_MATRIX(3, 3) * TEMP_MATRIX(3, 4)
    TEMP_MATRIX(4, 6) = TEMP_MATRIX(4, 3) * TEMP_MATRIX(4, 4)
    TEMP_MATRIX(5, 6) = TEMP_MATRIX(5, 3) * TEMP_MATRIX(5, 4)
    TEMP_MATRIX(6, 6) = TEMP_MATRIX(6, 3) * TEMP_MATRIX(6, 4)
    TEMP_MATRIX(7, 6) = TEMP_MATRIX(7, 3) * TEMP_MATRIX(7, 4)
    TEMP_MATRIX(8, 6) = TEMP_MATRIX(8, 3) * TEMP_MATRIX(8, 4)

    TEMP_MATRIX(NSIZE + 2, 1) = "Gross Profit:"
    TEMP_MATRIX(NSIZE + 2, 2) = "Net: " & BASE_FX_STR
    TEMP_MATRIX(NSIZE + 2, 3) = TEMP_MATRIX(NSIZE + 1, 6) - INITIAL_INVESTMENT
    
    TEMP_MATRIX(NSIZE + 2, 4) = ""
    TEMP_MATRIX(NSIZE + 2, 5) = ""
    TEMP_MATRIX(NSIZE + 2, 6) = ""
    
    TEMP_MATRIX(NSIZE + 3, 1) = "Less: Transaction Cost (" & _
    Format(h, "0.00") & _
    " * $" & (Format(UNIT_TRANSACTION_COST, "0.00")) & ")"
    TEMP_MATRIX(NSIZE + 3, 2) = ""

    TEMP_MATRIX(NSIZE + 3, 3) = -UNIT_TRANSACTION_COST * h
    
    TEMP_MATRIX(NSIZE + 3, 4) = ""
    TEMP_MATRIX(NSIZE + 3, 5) = ""
    TEMP_MATRIX(NSIZE + 3, 6) = ""
    
    TEMP_MATRIX(NSIZE + 4, 1) = "Net Profit"
    TEMP_MATRIX(NSIZE + 4, 2) = ""
    
    TEMP_MATRIX(NSIZE + 4, 3) = TEMP_MATRIX(NSIZE + 3, 3) + _
    TEMP_MATRIX(NSIZE + 2, 3)

    TEMP_MATRIX(NSIZE + 4, 4) = ""
    TEMP_MATRIX(NSIZE + 4, 5) = ""
    TEMP_MATRIX(NSIZE + 4, 6) = ""

    CURRENCIES_TRIANGLE_SEVEN_FUNC = TEMP_MATRIX
End Select

Exit Function
ERROR_LABEL:
CURRENCIES_TRIANGLE_SEVEN_FUNC = Err.number
End Function


Private Function CURRENCIES_TRIANGLE_EIGHT_FUNC(ByRef DATA_RNG As Variant, _
Optional ByVal INITIAL_INVESTMENT As Double = 15000000, _
Optional ByVal TARGET_PROFIT As Double = 0, _
Optional ByVal UNIT_TRANSACTION_COST As Double = 20, _
Optional ByVal BASE_FX_STR As String = "USD", _
Optional ByVal OUTPUT As Integer = 0)

'DATA_RNG = Sanitised cross currency exchange rates
'where, lower triangle = bid rates, upper triangle = offer rates
'TARGET_PROFIT: initial value for profit

'UNIT_TRANSACTION_COST: transaction costs: enter transaction cost per
'transaction in base currency.

'TARGET_PROFIT: TARGET_PROFIT_PROFIT

'--------------------------------Cross rates-------------------------------
'Most currencies are expressed against dollars – which means there is
'always a buy-sell spread that would normally make it difficult to make
'money by routing transactions through dollars. However, there are some
'cross rates that trade directly. A cross rate is a rate between two
'non US Dollar currencies. For example, there may be trading
'between British Pounds and the Euro. Therefore, we have a rate for the
'Dollar-Euro, we have a rate for the Dollar-Pound, and we also have a rate
'for Pound-Euro.
'--------------------------------------------------------------------------

Dim h As Long
Dim i As Long
Dim j As Long
Dim k As Long

Dim C0 As Long
Dim C1 As Long
Dim C2 As Long
Dim c3 As Long
Dim c4 As Long
Dim c5 As Long
Dim c6 As Long
Dim c7 As Long
Dim c8 As Long
Dim c9 As Long

Dim NSIZE As Long

Dim DST_ZERO As Long
Dim DST_ONE As Long
Dim DST_TWO As Long
Dim DST_THREE As Long
Dim DST_FOUR As Long
Dim DST_FIVE As Long
Dim DST_SIX As Long
Dim DST_SEVEN As Long
Dim DST_EIGHT As Long
Dim DST_NINE As Long

Dim ZERO_STR As String
Dim ONE_STR As String
Dim TWO_STR As String
Dim THREE_STR As String
Dim FOUR_STR As String
Dim FIVE_STR As String
Dim SIX_STR As String
Dim SEVEN_STR As String
Dim EIGHT_STR As String
Dim NINE_STR As String

Dim TEMP_STR As String
Dim DATA_MATRIX As Variant
Dim TEMP_VECTOR As Variant
Dim TEMP_SUM As Double

On Error GoTo ERROR_LABEL

DATA_MATRIX = MATRIX_REMOVE_ROWS_FUNC(MATRIX_REMOVE_COLUMNS_FUNC(DATA_RNG, 1), 1, 1, 1)
TEMP_VECTOR = MATRIX_REMOVE_COLUMNS_FUNC(MATRIX_GET_ROW_FUNC(DATA_RNG, 1, 1), 1, 1)
NSIZE = 8

If UBound(DATA_MATRIX, 1) < NSIZE Then GoTo ERROR_LABEL
If UBound(DATA_MATRIX, 2) < NSIZE Then GoTo ERROR_LABEL

'For any given number of currencies ‘N’, the number of cross-rates
'theoretically possible is N x (N-1)/2. The number of variations
'possible with N currencies to construct a “round-trip” would be NN,
'which means that the number of possibilities increases exponentially
'as we increase the number of currencies being considered.

'Note that each of the intervening currencies CC1 to 6 can be any of
'the 6 currencies, there is no limitation that a currency trade cannot
'be done twice, therefore the number of possible combinations is 6 x 6
'x 6 x 6 x 6 x 6, or 66. This is important because it allows us to look
'at profitable combinations where the number of currencies is less than
'six.

If IsNumeric(UNIT_TRANSACTION_COST) = False Then UNIT_TRANSACTION_COST = 0

'Use six nested “For…Next” loops to cycle through each possible combination
'of currencies. Each time, calculate the profit that would result from
'executing the given combination of currency transactions, and if this
'profit is greater than the profit from a previous combination, then store
'it away in a variable, otherwise discard it. (The initial value of profit is 0)
'If a profitable combination is not found – which for instance will be the case
'where transaction costs are very high, the macro says so. If a profitable
'combination is found, the same is listed together with all combinations.

For C1 = 1 To NSIZE
For C2 = 1 To NSIZE
For c3 = 1 To NSIZE
For c4 = 1 To NSIZE
For c5 = 1 To NSIZE
For c6 = 1 To NSIZE
For c7 = 1 To NSIZE
For c8 = 1 To NSIZE

'Update the status bar to show the combination being checked.
'      Excel.Application.StatusBar = _
 '     TEMP_VECTOR(1, c1) & "-" & _
  '    TEMP_VECTOR(1, c2) & "-" & _
   '   TEMP_VECTOR(1, c3) & "-" & _
    '  TEMP_VECTOR(1, c4) & "-" & _
     ' TEMP_VECTOR(1, c5) & "-" & _
     ' TEMP_VECTOR(1, c6) & "-" & _
     ' TEMP_VECTOR(1, c7) & "-" & _
      'TEMP_VECTOR(1, c8)

'Calculate true number of transactions. When the same _
'currency is repeated, then it is not a transaction, for _
'example USD->GBP->GBP->JPY->USD = USD->GBP->JPY->USD, ie _
'4 transactions and not 5, as GBP->GBP is costless. This is
'needed to calculate true transaction costs.

j = NSIZE + 1

If 1 = C1 Then j = j - 1
If C1 = C2 Then j = j - 1
If C2 = c3 Then j = j - 1
If c3 = c4 Then j = j - 1
If c4 = c5 Then j = j - 1
If c5 = c6 Then j = j - 1
If c6 = c7 Then j = j - 1
If c7 = c8 Then j = j - 1
If c8 = 1 Then j = j - 1

'For the currency combination in the loop, compare _
'TARGET_PROFIT to previous highest profits.

If TARGET_PROFIT < ((INITIAL_INVESTMENT * _
    DATA_MATRIX(C1, 1) * _
    DATA_MATRIX(C2, C1) * _
    DATA_MATRIX(c3, C2) * _
    DATA_MATRIX(c4, c3) * _
    DATA_MATRIX(c5, c4) * _
    DATA_MATRIX(c6, c5) * _
    DATA_MATRIX(c7, c6) * _
    DATA_MATRIX(c8, c7) * _
    DATA_MATRIX(1, c8) - _
    INITIAL_INVESTMENT) - (j * UNIT_TRANSACTION_COST)) _
Then
    
    DST_ZERO = 1
    DST_ONE = C1
    DST_TWO = C2
    DST_THREE = c3
    DST_FOUR = c4
    DST_FIVE = c5
    DST_SIX = c6
    DST_SEVEN = c7
    DST_EIGHT = c8
    DST_NINE = 1
    
    k = 1
    
    TARGET_PROFIT = ((INITIAL_INVESTMENT * _
    DATA_MATRIX(C1, 1) * _
    DATA_MATRIX(C2, C1) * _
    DATA_MATRIX(c3, C2) * _
    DATA_MATRIX(c4, c3) * _
    DATA_MATRIX(c5, c4) * _
    DATA_MATRIX(c6, c5) * _
    DATA_MATRIX(c7, c6) * _
    DATA_MATRIX(c8, c7) * _
    DATA_MATRIX(1, c8) - _
    INITIAL_INVESTMENT) - (j * UNIT_TRANSACTION_COST))
    
    h = j
End If

'ENHANCING THE MODEL:

'The model can be extended beyond spot transactions to look at forward
'transactions, interest rates, and identify any opportunities that may
'arise from an imbalance between the spot and forward rates given the
'interest rates in the two currencies. Such arbitrage would be based
'on violations of interest rate parity.

'In economies where foreign exchange markets are regulated, there often
'exists a black market exchange rate that is different from the official
'rate. Arbitrage opportunities between these rates are always possible
'and there is some money to be made. However, this may involve some risk
'taking in the form of possible violations of foreign exchange regulations.
'The cost of these violations with the probability of being discovered
'will need to be factored into the model.

TEMP_SUM = TEMP_SUM + 1   'merely a count of how many combinations checked.

Next c8
Next c7
Next c6
Next c5
Next c4
Next c3
Next C2
Next C1

'By now, the most profitable opportunity has been identified. If there _
'is no profitable opportunity, then the "GoSub finito" would not have _
'been visited and k would be zero)

If k = 0 Then
    CURRENCIES_TRIANGLE_EIGHT_FUNC = "No profitable possibilities!"
    Exit Function
Else
    
    ZERO_STR = TEMP_VECTOR(1, 1)
    'Replace currency numbers by their respective English
    ONE_STR = TEMP_VECTOR(1, DST_ONE)       'codes, eg 1 means USD etc
    TWO_STR = TEMP_VECTOR(1, DST_TWO)
    THREE_STR = TEMP_VECTOR(1, DST_THREE)
    FOUR_STR = TEMP_VECTOR(1, DST_FOUR)
    FIVE_STR = TEMP_VECTOR(1, DST_FIVE)
    SIX_STR = TEMP_VECTOR(1, DST_SIX)
    SEVEN_STR = TEMP_VECTOR(1, DST_SEVEN)
    EIGHT_STR = TEMP_VECTOR(1, DST_EIGHT)
    NINE_STR = TEMP_VECTOR(1, 1)
    
    
    For i = 1 To NSIZE + 1  'Remove duplicates, ie "GBP --> GBP" is _
    merely "GBP" in the transaction.
        If ZERO_STR = ONE_STR Then ZERO_STR = ""
        If ONE_STR = TWO_STR Then ONE_STR = ""
        If TWO_STR = THREE_STR Then TWO_STR = ""
        If THREE_STR = FOUR_STR Then THREE_STR = ""
        If FOUR_STR = FIVE_STR Then FOUR_STR = ""
        If FIVE_STR = SIX_STR Then FIVE_STR = ""
        If SIX_STR = SEVEN_STR Then SIX_STR = ""
        If SEVEN_STR = EIGHT_STR Then SEVEN_STR = ""
        If EIGHT_STR = NINE_STR Then EIGHT_STR = ""
    Next i
End If

Select Case OUTPUT
Case 0
    
    TEMP_STR = "Transact currencies as follows: " & _
        ZERO_STR & "-" & _
        ONE_STR & "-" & _
        TWO_STR & "-" & _
        THREE_STR & "-" & _
        FOUR_STR & "-" & _
        FIVE_STR & "-" & _
        SIX_STR & "-" & _
        SEVEN_STR & "-" & _
        EIGHT_STR & "-" & _
        NINE_STR & _
        " Profit on a " & _
        Format(INITIAL_INVESTMENT, "#,0.00") & " = " & _
        Format(TARGET_PROFIT, "#,0.00") & _
        "  (" & TEMP_SUM & " possibilities checked )"
        
        CURRENCIES_TRIANGLE_EIGHT_FUNC = TEMP_STR
        
Case Else

    ReDim TEMP_MATRIX(0 To NSIZE + 4, 1 To 6)
    
    TEMP_MATRIX(0, 1) = "Transaction"
    TEMP_MATRIX(0, 2) = "Details"
    TEMP_MATRIX(0, 3) = "Amount sold"
    TEMP_MATRIX(0, 4) = "Exchange rate"
    TEMP_MATRIX(0, 5) = "Details"
    TEMP_MATRIX(0, 6) = "Amount bought"

    If ZERO_STR = "" Then ZERO_STR = BASE_FX_STR
    If ONE_STR = "" Then ONE_STR = BASE_FX_STR
    If TWO_STR = "" Then TWO_STR = BASE_FX_STR
    If THREE_STR = "" Then THREE_STR = BASE_FX_STR
    If FOUR_STR = "" Then FOUR_STR = BASE_FX_STR
    If FIVE_STR = "" Then FIVE_STR = BASE_FX_STR
    If SIX_STR = "" Then SIX_STR = BASE_FX_STR
    If SEVEN_STR = "" Then SEVEN_STR = BASE_FX_STR
    If EIGHT_STR = "" Then EIGHT_STR = BASE_FX_STR
    If NINE_STR = "" Then NINE_STR = BASE_FX_STR
    
    'Write the transaction
    TEMP_MATRIX(1, 1) = "Sell " & ZERO_STR & ", Buy " & ONE_STR
    TEMP_MATRIX(2, 1) = "Sell " & ONE_STR & ", Buy " & TWO_STR
    TEMP_MATRIX(3, 1) = "Sell " & TWO_STR & ", Buy " & THREE_STR
    TEMP_MATRIX(4, 1) = "Sell " & THREE_STR & ", Buy " & FOUR_STR
    TEMP_MATRIX(5, 1) = "Sell " & FOUR_STR & ", Buy " & FIVE_STR
    TEMP_MATRIX(6, 1) = "Sell " & FIVE_STR & ", Buy " & SIX_STR
    TEMP_MATRIX(7, 1) = "Sell " & SIX_STR & ", Buy " & SEVEN_STR
    TEMP_MATRIX(8, 1) = "Sell " & SEVEN_STR & ", Buy " & EIGHT_STR
    TEMP_MATRIX(9, 1) = "Sell " & EIGHT_STR & ", Buy " & NINE_STR
    
    'Write the transaction details for currency sold
    TEMP_MATRIX(1, 2) = ZERO_STR & " sold:"
    TEMP_MATRIX(2, 2) = ONE_STR & " sold:"
    TEMP_MATRIX(3, 2) = TWO_STR & " sold:"
    TEMP_MATRIX(4, 2) = THREE_STR & " sold:"
    TEMP_MATRIX(5, 2) = FOUR_STR & " sold:"
    TEMP_MATRIX(6, 2) = FIVE_STR & " sold:"
    TEMP_MATRIX(7, 2) = SIX_STR & " sold:"
    TEMP_MATRIX(8, 2) = SEVEN_STR & " sold:"
    TEMP_MATRIX(9, 2) = EIGHT_STR & " sold:"
    
    'Write the transaction details for currency bought
    TEMP_MATRIX(1, 5) = ONE_STR & " bought:"
    TEMP_MATRIX(2, 5) = TWO_STR & " bought:"
    TEMP_MATRIX(3, 5) = THREE_STR & " bought:"
    TEMP_MATRIX(4, 5) = FOUR_STR & " bought:"
    TEMP_MATRIX(5, 5) = FIVE_STR & " bought:"
    TEMP_MATRIX(6, 5) = SIX_STR & " bought:"
    TEMP_MATRIX(7, 5) = SEVEN_STR & " bought:"
    TEMP_MATRIX(8, 5) = EIGHT_STR & " bought:"
    TEMP_MATRIX(9, 5) = NINE_STR & " bought:"
    
    'MATRIX_FIND_ELEMENT_FUNC
    
    C0 = MATRIX_FIND_ELEMENT_FUNC(TEMP_VECTOR, ZERO_STR, 1, 1, 0)
    C1 = MATRIX_FIND_ELEMENT_FUNC(TEMP_VECTOR, ONE_STR, 1, 1, 0)
    C2 = MATRIX_FIND_ELEMENT_FUNC(TEMP_VECTOR, TWO_STR, 1, 1, 0)
    c3 = MATRIX_FIND_ELEMENT_FUNC(TEMP_VECTOR, THREE_STR, 1, 1, 0)
    c4 = MATRIX_FIND_ELEMENT_FUNC(TEMP_VECTOR, FOUR_STR, 1, 1, 0)
    c5 = MATRIX_FIND_ELEMENT_FUNC(TEMP_VECTOR, FIVE_STR, 1, 1, 0)
    c6 = MATRIX_FIND_ELEMENT_FUNC(TEMP_VECTOR, SIX_STR, 1, 1, 0)
    c7 = MATRIX_FIND_ELEMENT_FUNC(TEMP_VECTOR, SEVEN_STR, 1, 1, 0)
    c8 = MATRIX_FIND_ELEMENT_FUNC(TEMP_VECTOR, EIGHT_STR, 1, 1, 0)
    c9 = MATRIX_FIND_ELEMENT_FUNC(TEMP_VECTOR, NINE_STR, 1, 1, 0)
    
    TEMP_MATRIX(1, 4) = DATA_MATRIX(C1, C0)
    
    TEMP_MATRIX(2, 4) = DATA_MATRIX(C2, C1)
    
    TEMP_MATRIX(3, 4) = DATA_MATRIX(c3, C2)
    
    TEMP_MATRIX(4, 4) = DATA_MATRIX(c4, c3)
    
    TEMP_MATRIX(5, 4) = DATA_MATRIX(c5, c4)
    
    TEMP_MATRIX(6, 4) = DATA_MATRIX(c6, c5)
    
    TEMP_MATRIX(7, 4) = DATA_MATRIX(c7, c6)
    
    TEMP_MATRIX(8, 4) = DATA_MATRIX(c8, c7)
    
    TEMP_MATRIX(9, 4) = DATA_MATRIX(c9, c8)
        
    TEMP_MATRIX(1, 3) = INITIAL_INVESTMENT
    
    TEMP_MATRIX(2, 3) = TEMP_MATRIX(1, 3) * TEMP_MATRIX(1, 4)
    TEMP_MATRIX(3, 3) = TEMP_MATRIX(2, 3) * TEMP_MATRIX(2, 4)
    TEMP_MATRIX(4, 3) = TEMP_MATRIX(3, 3) * TEMP_MATRIX(3, 4)
    TEMP_MATRIX(5, 3) = TEMP_MATRIX(4, 3) * TEMP_MATRIX(4, 4)
    TEMP_MATRIX(6, 3) = TEMP_MATRIX(5, 3) * TEMP_MATRIX(5, 4)
    TEMP_MATRIX(7, 3) = TEMP_MATRIX(6, 3) * TEMP_MATRIX(6, 4)
    TEMP_MATRIX(8, 3) = TEMP_MATRIX(7, 3) * TEMP_MATRIX(7, 4)
    TEMP_MATRIX(9, 3) = TEMP_MATRIX(8, 3) * TEMP_MATRIX(8, 4)
    
    TEMP_MATRIX(1, 6) = TEMP_MATRIX(1, 3) * TEMP_MATRIX(1, 4)
    TEMP_MATRIX(2, 6) = TEMP_MATRIX(2, 3) * TEMP_MATRIX(2, 4)
    TEMP_MATRIX(3, 6) = TEMP_MATRIX(3, 3) * TEMP_MATRIX(3, 4)
    TEMP_MATRIX(4, 6) = TEMP_MATRIX(4, 3) * TEMP_MATRIX(4, 4)
    TEMP_MATRIX(5, 6) = TEMP_MATRIX(5, 3) * TEMP_MATRIX(5, 4)
    TEMP_MATRIX(6, 6) = TEMP_MATRIX(6, 3) * TEMP_MATRIX(6, 4)
    TEMP_MATRIX(7, 6) = TEMP_MATRIX(7, 3) * TEMP_MATRIX(7, 4)
    TEMP_MATRIX(8, 6) = TEMP_MATRIX(8, 3) * TEMP_MATRIX(8, 4)
    TEMP_MATRIX(9, 6) = TEMP_MATRIX(9, 3) * TEMP_MATRIX(9, 4)

    TEMP_MATRIX(NSIZE + 2, 1) = "Gross Profit:"
    TEMP_MATRIX(NSIZE + 2, 2) = "Net: " & BASE_FX_STR
    TEMP_MATRIX(NSIZE + 2, 3) = TEMP_MATRIX(NSIZE + 1, 6) - INITIAL_INVESTMENT
    
    TEMP_MATRIX(NSIZE + 2, 4) = ""
    TEMP_MATRIX(NSIZE + 2, 5) = ""
    TEMP_MATRIX(NSIZE + 2, 6) = ""
    
    TEMP_MATRIX(NSIZE + 3, 1) = "Less: Transaction Cost (" & _
    Format(h, "0.00") & _
    " * $" & (Format(UNIT_TRANSACTION_COST, "0.00")) & ")"
    TEMP_MATRIX(NSIZE + 3, 2) = ""

    TEMP_MATRIX(NSIZE + 3, 3) = -UNIT_TRANSACTION_COST * h
    
    TEMP_MATRIX(NSIZE + 3, 4) = ""
    TEMP_MATRIX(NSIZE + 3, 5) = ""
    TEMP_MATRIX(NSIZE + 3, 6) = ""
    
    TEMP_MATRIX(NSIZE + 4, 1) = "Net Profit"
    TEMP_MATRIX(NSIZE + 4, 2) = ""
    
    TEMP_MATRIX(NSIZE + 4, 3) = TEMP_MATRIX(NSIZE + 3, 3) + _
    TEMP_MATRIX(NSIZE + 2, 3)

    TEMP_MATRIX(NSIZE + 4, 4) = ""
    TEMP_MATRIX(NSIZE + 4, 5) = ""
    TEMP_MATRIX(NSIZE + 4, 6) = ""

    CURRENCIES_TRIANGLE_EIGHT_FUNC = TEMP_MATRIX
End Select

Exit Function
ERROR_LABEL:
CURRENCIES_TRIANGLE_EIGHT_FUNC = Err.number
End Function
