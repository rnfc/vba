Attribute VB_Name = "FINAN_DERIV_BS_CAGR_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.



'************************************************************************************
'************************************************************************************
'FUNCTION      : OPTION_REQUIRED_CAGR_REPORT_FUNC
'DESCRIPTION   :

'You buy a call option for C = $4.06 with strike price K = $25.00.
'It expires in T = 143 days.

'That gives you the right to buy the stock at $25 any time over the
'next 143 days.

'If you do excercise the option (buying at $25), then you've paid
'K + C = $25.00 + $4.06 = $29.06 for each share.

'What are the chances that the stock price will be greater than
'$29.06 in 143 days?"

'How rapidly must the stock price increase over the next 143 days in
'order to reach that magic number: $29.06?

'If we find that the stock must increase at the rate of, say, 50% per
'year, we might look for other options.

'For example, if the stock price (when you buy the option) is $28.07,
'then it's got 143 days to get to $$29.06.

'That 's an increase of (29.06/28.07)(365/143) -1 = 0.0925 or 9.25% on
'an annualized basis.

'But that's not 50%.

'Yes, but if the current stock price were $15.45 then we'd require an
'annual rate of (29.06/15.45)(365/143) -1 = 0.50 or 50%.
'Does that happen? I mean, a required annual rate of 50%?

'Reference: http://www.gummy-stuff.org/Option-things.htm

'LIBRARY       : DERIVATIVES
'GROUP         : BS_RETURN
'ID            : 003
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/29/2009
'************************************************************************************
'************************************************************************************

Function OPTION_REQUIRED_CAGR_REPORT_FUNC(ByVal STRIKE_PRICE As Double, _
ByVal STOCK_PRICE As Double, _
ByVal EXPIRATION As Double, _
ByVal VOLATILITY As Double, _
ByVal RISK_FREE_RATE As Double, _
Optional ByVal OPTION_FLAG As Integer = 0, _
Optional ByVal CND_TYPE As Integer = 0)

'If you don't know the Stock Volatility or Risk-free Rate then
'stick in some values so that the Black-Scholes estimate agrees with KNOWN
'option premiums for the stock. (Example: If the Option Premium, today, is $8.72
'instead of $8.52, then increase Volatility and/or Risk-free Rate
'so that the Black-Scholes estimate agrees with Today's option premium.
'(Increasing either will increase the Black-Scholes estimate.)
                                                
Dim TEMP_STR As String
Dim CAGR_VAL As Double
Dim PREMIUM_VAL As Double
Dim TEMP_VECTOR As Variant

On Error GoTo ERROR_LABEL

If OPTION_FLAG = 1 Then
    TEMP_STR = "Call"
Else
    TEMP_STR = "Put"
End If

PREMIUM_VAL = BLACK_SCHOLES_OPTION_FUNC(STOCK_PRICE, STRIKE_PRICE, _
              EXPIRATION, RISK_FREE_RATE, VOLATILITY, OPTION_FLAG, CND_TYPE)

CAGR_VAL = OPTION_REQUIRED_CAGR_FUNC(STRIKE_PRICE, STOCK_PRICE, EXPIRATION, _
           PREMIUM_VAL)
    
ReDim TEMP_VECTOR(1 To 7, 1 To 1)

TEMP_VECTOR(1, 1) = "Suppose the " & TEMP_STR & " option expires in " & _
                    Format(EXPIRATION, "0.00") & " years."

TEMP_VECTOR(2, 1) = "Suppose, too, that we're interested in an " & _
                    "option with Strike Price of " & _
                    Format(STRIKE_PRICE, "Currency")

TEMP_VECTOR(3, 1) = "Suppose, too, that we assume a Volatility of " & _
                    Format(VOLATILITY, "0.00%") & _
                    " and a Risk-free Rate of " & _
                    Format(RISK_FREE_RATE, "0.00%")

TEMP_VECTOR(4, 1) = "We require the Black-Scholes estimate, assuming a " & _
                    "current Stock Price of " & _
                    Format(STOCK_PRICE, "Currency") & "."

TEMP_VECTOR(5, 1) = "According to Black-Scholes, the option should be " & _
                    "worth " & Format(PREMIUM_VAL, "Currency") & " … today."

If OPTION_FLAG = 1 Then 'Call
    TEMP_VECTOR(6, 1) = "To make money, the Stock must be worth " & _
                        Format(STRIKE_PRICE, "Currency") & "+" & _
                        Format(PREMIUM_VAL, "Currency") & "=" & _
                        Format(STRIKE_PRICE + PREMIUM_VAL, "Currency") & _
                        " in " & Format(EXPIRATION, "0.00") & " years… plus commissions!"
Else 'Put
    TEMP_VECTOR(6, 1) = "To make money, the Stock must be worth " & _
                        Format(STRIKE_PRICE, "Currency") & "-" & _
                        Format(PREMIUM_VAL, "Currency") & "=" & _
                        Format(STRIKE_PRICE - PREMIUM_VAL, "Currency") & _
                        " in " & Format(EXPIRATION, "0.00") & " years… plus commissions!"

End If
TEMP_VECTOR(7, 1) = "Required CAGR " & Format(CAGR_VAL, "0.00%")

OPTION_REQUIRED_CAGR_REPORT_FUNC = TEMP_VECTOR


Exit Function
ERROR_LABEL:
OPTION_REQUIRED_CAGR_REPORT_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : OPTION_REQUIRED_CAGR_FUNC
'DESCRIPTION   : At what annual rate must the stock price change in order to make
'money buying a Call/Put Option ??

'LIBRARY       : DERIVATIVES
'GROUP         : BS_RETURN
'ID            : 002
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/29/2009
'************************************************************************************
'************************************************************************************

Function OPTION_REQUIRED_CAGR_FUNC(ByVal STRIKE_PRICE As Double, _
ByVal STOCK_PRICE As Double, _
ByVal EXPIRATION As Double, _
ByVal ACTUAL_PREMIUM As Double, _
Optional ByVal OPTION_FLAG As Integer = 1)

On Error GoTo ERROR_LABEL

If OPTION_FLAG = 1 Then 'Stock Price > Strike + Call Premium
    OPTION_REQUIRED_CAGR_FUNC = ((ACTUAL_PREMIUM + STRIKE_PRICE) / STOCK_PRICE) ^ (1 / EXPIRATION) - 1
Else 'Stock Price < Strike - Call Premium
    OPTION_REQUIRED_CAGR_FUNC = ((-ACTUAL_PREMIUM + STRIKE_PRICE) / STOCK_PRICE) ^ (1 / EXPIRATION) - 1
End If
    
Exit Function
ERROR_LABEL:
OPTION_REQUIRED_CAGR_FUNC = Err.number
End Function

