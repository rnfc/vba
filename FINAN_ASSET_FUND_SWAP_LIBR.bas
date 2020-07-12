Attribute VB_Name = "FINAN_ASSET_FUND_SWAP_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.

'-------------------------------------------------------------------------------------------
'-------------------------------------------------------------------------------------------
'Re-balancing and Hedging Demands: Daily Hedging Demands
'-------------------------------------------------------------------------------------------
'-------------------------------------------------------------------------------------------
'The re-balancing of inverse and leveraged funds implies certain hedging demands.
'Since extant funds promise a multiple of the day's return, it makes sense to
'focus on end-of-day hedging demands. One benefit of modeling returns in continuous
'time, however, is that our analysis generalizes to any arbitrary re-balancing
'interval. Indeed, there has been recent discusssion of new leveraged funds that
'track an index on a monthly versus daily basis.
'-------------------------------------------------------------------------------------------
'-------------------------------------------------------------------------------------------

Function TOTAL_RETURN_SWAPS_EXPOSURE_FUNC(ByRef RETURNS_RNG As Variant, _
Optional ByVal INITIAL_NAV_VAL As Double = 100, _
Optional ByVal LEVERAGE_VAL As Double = 2, _
Optional ByVal OUTPUT As Integer = 0)
'-------------------------------------------------------------------------------------------
'http://en.wikipedia.org/wiki/Total_return_swap
'ftp://174.34.168.194/Specs_%26_Docs/ETFs/Dynamics_Leveraged_&_Inverse_ETFs.pdf
'-------------------------------------------------------------------------------------------

Dim i As Long
Dim NROWS As Long

Dim S_VAL As Double
Dim E_VAL As Double
Dim A_VAL As Double
Dim L_VAL As Double
Dim D_VAL As Double
Dim R_VAL As Double
Dim H_VAL As Double

Dim TEMP_MATRIX As Variant
Dim RETURNS_VECTOR As Variant

On Error GoTo ERROR_LABEL

RETURNS_VECTOR = RETURNS_RNG
If UBound(RETURNS_VECTOR, 1) = 1 Then
    RETURNS_VECTOR = MATRIX_TRANSPOSE_FUNC(RETURNS_VECTOR)
End If
NROWS = UBound(RETURNS_VECTOR, 1)
'-------------------------------------------------------------------------------------------
If OUTPUT = 0 Then
    ReDim TEMP_MATRIX(0 To NROWS, 1 To 7)
    TEMP_MATRIX(0, 1) = "PERIOD(T)"
    TEMP_MATRIX(0, 2) = "INDEX VALUE(S)"
    TEMP_MATRIX(0, 3) = "FUND NAV(A)"
    TEMP_MATRIX(0, 4) = "SWAP NOTIONAL(L)" 'REQUIRED NOTIONAL AMOUNT
    TEMP_MATRIX(0, 5) = "SWAP EXPOSURE(E)"
    TEMP_MATRIX(0, 6) = "SWAP RE-HEDGED(D)"
    TEMP_MATRIX(0, 7) = "HEDGING TERM(H)"
End If
'-------------------------------------------------------------------------------------------
'Let At(n) represent a leveraged or inverse ETF's NAV at the close of day n or at time
't(n). Corresponding to At(n), let Lt(n) represent the notional amount of the total
'return swaps exposure that is required before the market opens ont eh next day to
'replicate the intended leveraged return of the index for fund from calendar time t(n)
'to time t(n+1). With the fund's NAV at At(n) at time t(n), the notional amount of the
'total return swaps required is given by: Lt(n) = x At(n)
A_VAL = INITIAL_NAV_VAL
L_VAL = A_VAL * LEVERAGE_VAL
S_VAL = 100
H_VAL = (LEVERAGE_VAL ^ 2 - LEVERAGE_VAL)
'-------------------------------------------------------------------------------------------
i = 1
If OUTPUT = 0 Then
    TEMP_MATRIX(i, 1) = i
    TEMP_MATRIX(i, 2) = S_VAL
    TEMP_MATRIX(i, 3) = A_VAL
    TEMP_MATRIX(i, 4) = L_VAL
    TEMP_MATRIX(i, 5) = ""
    TEMP_MATRIX(i, 6) = ""
    TEMP_MATRIX(i, 7) = ""
End If
'-------------------------------------------------------------------------------------------
For i = 2 To NROWS
'-------------------------------------------------------------------------------------------
    R_VAL = RETURNS_VECTOR(i, 1) / RETURNS_VECTOR(i - 1, 1) - 1
    'Index Value
    S_VAL = S_VAL * (1 + R_VAL)
'-------------------------------------------------------------------------------------------
    'On day n + 1, the underlying index generates a return of Rt(n),t(n+1) and the
    'exposure of the total return swaps, denoted by Et(n+1), becomes:
    E_VAL = L_VAL * (1 + R_VAL)
'-------------------------------------------------------------------------------------------
    'D_VAL = A_VAL * (LEVERAGE_VAL ^ 2 - LEVERAGE_VAL) * R_VAL
'-------------------------------------------------------------------------------------------
    'At the same time, reflecting the gain or loss that is x times the index's performance
    'between t(n) and t(n+1), the leveraged fund's NAV at the close of day n + 1 becomes:
    A_VAL = A_VAL * (1 + LEVERAGE_VAL * R_VAL)
'-------------------------------------------------------------------------------------------
    'which suggests that the notional amount of the total return swaps this is
    'required before the market opens next day to maintain cosntant exposure is:
    L_VAL = LEVERAGE_VAL * A_VAL
'-------------------------------------------------------------------------------------------
    'The different between E_VAL and L_VAL; denoted by D_VAL, is the amount
    'by which the exposure of the total return swaps that need to be adjusted or re-hedged at
    'time tn+1, as given by:
    D_VAL = L_VAL - E_VAL
'-------------------------------------------------------------------------------------------
    If OUTPUT = 0 Then
        TEMP_MATRIX(i, 1) = i
        TEMP_MATRIX(i, 2) = S_VAL
        TEMP_MATRIX(i, 3) = A_VAL
        TEMP_MATRIX(i, 4) = L_VAL
        TEMP_MATRIX(i, 5) = E_VAL
        TEMP_MATRIX(i, 6) = D_VAL
        TEMP_MATRIX(i, 7) = H_VAL
    End If
'-------------------------------------------------------------------------------------------
Next i
'-------------------------------------------------------------------------------------------


Select Case OUTPUT 'The dynamics of Stn, Atn, Ltn, Etn and Dtn
Case 0
    TOTAL_RETURN_SWAPS_EXPOSURE_FUNC = TEMP_MATRIX
Case Else
    TOTAL_RETURN_SWAPS_EXPOSURE_FUNC = Array(S_VAL, A_VAL, L_VAL, E_VAL, D_VAL, H_VAL)
End Select

'We can illustrate the above using an example:

'Example 1: Dynamics of a double-leveraged ETF (x = 2, rt0; t1 = ..10% and rt1; t2 = 10%)
'With an initial NAV of $100 on day 0 for the double-leveraged ETF, the
'required notional amount of the total return swaps is $200 (or 2 times $100). As the index
'falls from 100 to 90 on day 1, the fund's NAV drops to $80 whereas the exposure of the
'total return swaps falls to $180, reflecting a 10% drop of its value. Meanwhile the required
'notional amount for the total return swaps for day 2 is $160 (or 2 times $80), which means
'the fund will need to reduce its exposure of total return swaps by $20 (or $180 minus $160)
'at the end of day 1. And note 100 x (2^2 - 2) x 10% = 20.

'Example 2 and 3 provide examples for an inverse ETF and a double-inverse ETF, re-
'spectively, with the same assumptions of the index's performance over two days. These
'examples highlight the critical role of the hedging term (x^2 - x). This term is
'non-linear and asymmetric. For example, it takes the value 6 for triple-leveraged (x = 3)
'and double-inverse (x = -2) ETFs. As (x2 - x) is always positive (except for when x = 1
'when the funds are not leveraged or inverse), the reset or re-balance flows are always
'in the same direction as the underlying index's performance. So, when the underlying index is
'up, additional total return swap exposure must be added, but when the underlying index is
'down, the exposure of total return swaps needs to be reduced. This is intuitively clear for
'a leveraged long fund, is always true whether the ETFs are leveraged, inverse or leveraged
'inverse. Why is the effect the same for funds that are short the index? Intuitively, an
'inverse or leveraged inverse fund's NAV will increase if the index falls, which requires it to
'increase short exposure still further, generating selling pressure. In other words, there is no
'offset or pairing off of leveraged long and short ETFs on the same index, which is why the
're -balance flows are in the same direction between Example 1 for a double-leveraged ETF and
'Example 2 and 3 for an inverse and a double-inverse ETF on day 1 and day 2, respectively.

'Note the need for daily re-hedging is unique to leveraged and inverse ETFs due to
'their product design. Traditional ETFs that are not leveraged or inverse, whether they are
'holding physicals, total return swaps or other derivatives, have no need to re-balance daily.

Exit Function
ERROR_LABEL:
TOTAL_RETURN_SWAPS_EXPOSURE_FUNC = Err.number
End Function


