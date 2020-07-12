Attribute VB_Name = "STAT_DIST_MIXTURES_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.

'************************************************************************************
'************************************************************************************
'FUNCTION      : GAMMA_POISSON_DIST_FUNC_FUNC
'DESCRIPTION   : Returns the Mixtures/Compound Gamma-Poisson Distibution.
'LIBRARY       : STATISTICS
'GROUP         : DIST_MIXTURES
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/22/2009
'************************************************************************************
'************************************************************************************

Function GAMMA_POISSON_DIST_FUNC_FUNC(ByVal NO_EVENTS As Double, _
ByVal ALPHA As Double, _
ByVal beta As Double, _
Optional ByVal CUMUL_FLAG As Boolean = True, _
Optional ByVal COMP_FLAG As Boolean = True)

'NO_EVENTS: Number of Events

'ALPHA: is a shape parameter to the distribution.

'BETA: is a scale parameter to the distribution.

'CUMUL_FLAG: is a logical value that determines the form of the
'function. If cumulative is TRUE, then it returns the
'cumulative distribution function; if FALSE, it returns the
'probability mass function.

'COMP_FLAG: 1-cumulative distribution function

On Error GoTo ERROR_LABEL

Select Case CUMUL_FLAG
    Case True
        If COMP_FLAG = True Then
            GAMMA_POISSON_DIST_FUNC_FUNC = cdf_GammaPoisson((NO_EVENTS), _
            (ALPHA), (beta))
        ElseIf COMP_FLAG = False Then
            GAMMA_POISSON_DIST_FUNC_FUNC = comp_cdf_GammaPoisson((NO_EVENTS), _
            (ALPHA), (beta))
        End If
    Case False 'probability density function
        GAMMA_POISSON_DIST_FUNC_FUNC = pmf_GammaPoisson((NO_EVENTS), _
        (ALPHA), (beta))
End Select

Exit Function
ERROR_LABEL:
GAMMA_POISSON_DIST_FUNC_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : CRIT_GAMMA_POISSON_DIST_FUNC_FUNC
'DESCRIPTION   : Returns the smallest value for which the cumulative
'Gamma-Poisson distribution is greater than or equal to a criterion value.
'LIBRARY       : STATISTICS
'GROUP         : DIST_MIXTURES
'ID            : 002
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/22/2009
'************************************************************************************
'************************************************************************************

Function CRIT_GAMMA_POISSON_DIST_FUNC_FUNC(ByVal PROBABILITY_VAL As Double, _
ByVal ALPHA As Double, _
ByVal beta As Double, _
Optional ByVal COMP_FLAG As Boolean = True)

'PROBABILITY_VAL: Crit Prob
'ALPHA: is a shape parameter to the distribution.
'BETA: is a scale parameter to the distribution.

On Error GoTo ERROR_LABEL

Select Case COMP_FLAG
    Case True
        CRIT_GAMMA_POISSON_DIST_FUNC_FUNC = crit_GammaPoisson((ALPHA), _
            (beta), (PROBABILITY_VAL))
    Case False
        CRIT_GAMMA_POISSON_DIST_FUNC_FUNC = comp_crit_GammaPoisson((ALPHA), _
            (beta), (PROBABILITY_VAL))
End Select

Exit Function
ERROR_LABEL:
CRIT_GAMMA_POISSON_DIST_FUNC_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : BETA_BINOMDIST_FUNC
'DESCRIPTION   : Returns the Mixtures/Compound Beta Binomial Distibution
'LIBRARY       : STATISTICS
'GROUP         : DIST_MIXTURES
'ID            : 003
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/22/2009
'************************************************************************************
'************************************************************************************
'***************************************************************************

Function BETA_BINOMDIST_FUNC(ByVal NO_EVENTS As Double, _
ByVal SAMPLE_SIZE As Double, _
ByVal FIRST_ALPHA As Double, _
ByVal SECOND_ALPHA As Double, _
Optional ByVal CUMUL_FLAG As Boolean = True, _
Optional ByVal COMP_FLAG As Boolean = True)

'NO_EVENTS: Number of Events

'FIRST_ALPHA: is a shape parameter to the distribution.

'SECOND_ALPHA: is a scale parameter to the distribution.

'CUMUL_FLAG: is a logical value that determines the form of the
'function. If cumulative is TRUE, then it returns the
'cumulative distribution function; if FALSE, it returns the
'probability mass function.

'COMP_FLAG: 1-cumulative distribution function

On Error GoTo ERROR_LABEL

Select Case CUMUL_FLAG
    Case True
        If COMP_FLAG = True Then
            BETA_BINOMDIST_FUNC = cdf_BetaBinomial((NO_EVENTS), _
            (SAMPLE_SIZE), (FIRST_ALPHA), (SECOND_ALPHA))
        ElseIf COMP_FLAG = False Then
            BETA_BINOMDIST_FUNC = comp_cdf_BetaBinomial((NO_EVENTS), _
            (SAMPLE_SIZE), (FIRST_ALPHA), (SECOND_ALPHA))
        End If
    Case False 'probability density function
        BETA_BINOMDIST_FUNC = pmf_BetaBinomial((NO_EVENTS), _
        (SAMPLE_SIZE), (FIRST_ALPHA), (SECOND_ALPHA))
End Select

Exit Function
ERROR_LABEL:
BETA_BINOMDIST_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : CRIT_BETA_BINOMDIST_FUNC
'DESCRIPTION   : Returns the smallest value for which the cumulative
'Beta-Binomial distribution is greater than or equal to a criterion value.
'LIBRARY       : STATISTICS
'GROUP         : DIST_MIXTURES
'ID            : 004
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/22/2009
'************************************************************************************
'************************************************************************************

Function CRIT_BETA_BINOMDIST_FUNC(ByVal PROBABILITY_VAL As Double, _
ByVal SAMPLE_SIZE As Double, _
ByVal FIRST_ALPHA As Double, _
ByVal SECOND_ALPHA As Double, _
Optional ByVal COMP_FLAG As Boolean = True)

'PROBABILITY_VAL: Crit Prob
'FIRST_ALPHA: is a shape parameter to the distribution.
'SECOND_ALPHA: is a scale parameter to the distribution.

On Error GoTo ERROR_LABEL

Select Case COMP_FLAG
    Case True
        CRIT_BETA_BINOMDIST_FUNC = crit_BetaBinomial((SAMPLE_SIZE), _
        (FIRST_ALPHA), (SECOND_ALPHA), (PROBABILITY_VAL))
    Case False
        CRIT_BETA_BINOMDIST_FUNC = comp_crit_BetaBinomial((SAMPLE_SIZE), _
        (FIRST_ALPHA), (SECOND_ALPHA), (PROBABILITY_VAL))
End Select

Exit Function
ERROR_LABEL:
CRIT_BETA_BINOMDIST_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : BETA_NEG_BINOMDIST_FUNC
'DESCRIPTION   : Returns the Mixtures/Compound Beta NegBinomial Distibution.
'LIBRARY       : STATISTICS
'GROUP         : DIST_MIXTURES
'ID            : 005
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/22/2009
'************************************************************************************
'************************************************************************************

Function BETA_NEG_BINOMDIST_FUNC(ByVal NO_EVENTS As Double, _
ByVal RADIUS As Double, _
ByVal FIRST_ALPHA As Double, _
ByVal SECOND_ALPHA As Double, _
Optional ByVal CUMUL_FLAG As Boolean = True, _
Optional ByVal COMP_FLAG As Boolean = True)

'NO_EVENTS: Number of Events

'FIRST_ALPHA: is a shape parameter to the distribution.

'SECOND_ALPHA: is a scale parameter to the distribution.

'CUMUL_FLAG: is a logical value that determines the form of the
'function. If cumulative is TRUE, then it returns the
'cumulative distribution function; if FALSE, it returns the
'probability mass function.

'COMP_FLAG: 1-cumulative distribution function

On Error GoTo ERROR_LABEL

Select Case CUMUL_FLAG
    Case True
        If COMP_FLAG = True Then
            BETA_NEG_BINOMDIST_FUNC = cdf_BetaNegativeBinomial((NO_EVENTS), _
            (RADIUS), (FIRST_ALPHA), (SECOND_ALPHA))
        ElseIf COMP_FLAG = False Then
            BETA_NEG_BINOMDIST_FUNC = comp_cdf_BetaNegativeBinomial((NO_EVENTS), _
            (RADIUS), (FIRST_ALPHA), (SECOND_ALPHA))
        End If
    Case False 'probability density function
        BETA_NEG_BINOMDIST_FUNC = pmf_BetaNegativeBinomial((NO_EVENTS), _
        (RADIUS), (FIRST_ALPHA), (SECOND_ALPHA))
End Select

Exit Function
ERROR_LABEL:
BETA_NEG_BINOMDIST_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : CRIT_BETA_NEG_BINOMDIST_FUNC
'DESCRIPTION   : Returns the smallest value for which the cumulative
'Beta-NegBinomial distribution is greater than or equal to a criterion value.
'LIBRARY       : STATISTICS
'GROUP         : DIST_MIXTURES
'ID            : 006
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/22/2009
'************************************************************************************
'************************************************************************************

Function CRIT_BETA_NEG_BINOMDIST_FUNC(ByVal PROBABILITY_VAL As Double, _
ByVal RADIUS As Double, _
ByVal FIRST_ALPHA As Double, _
ByVal SECOND_ALPHA As Double, _
Optional ByVal COMP_FLAG As Boolean = True)

'PROBABILITY_VAL: Crit Prob
'FIRST_ALPHA: is a shape parameter to the distribution.
'SECOND_ALPHA: is a scale parameter to the distribution.

On Error GoTo ERROR_LABEL

Select Case COMP_FLAG
    Case True
        CRIT_BETA_NEG_BINOMDIST_FUNC = crit_BetaNegativeBinomial((RADIUS), _
        (FIRST_ALPHA), (SECOND_ALPHA), (PROBABILITY_VAL))
    Case False
        CRIT_BETA_NEG_BINOMDIST_FUNC = comp_crit_BetaNegativeBinomial((RADIUS), _
        (FIRST_ALPHA), (SECOND_ALPHA), (PROBABILITY_VAL))
End Select

Exit Function
ERROR_LABEL:
CRIT_BETA_NEG_BINOMDIST_FUNC = Err.number
End Function

