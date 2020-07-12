Attribute VB_Name = "STAT_DIST_DISCRETE_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.

'************************************************************************************
'************************************************************************************
'FUNCTION      : BINOMDIST_FUNC
'DESCRIPTION   : Returns the individual term binomial distribution probability.
'Use this function in problems with a fixed number of tests or trials,
'when the outcomes of any trial are only success or failure, when trials
'are independent, and when the probability of success is constant
'throughout the experiment. For example, it can calculate the probability
'that two of the next three babies born are male.
'LIBRARY       : STATISTICS
'GROUP         : DIST_DISCRETE
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/22/2009
'************************************************************************************
'************************************************************************************

Function BINOMDIST_FUNC(ByVal NO_SUCCESS As Double, _
ByVal SAMPLE_SIZE As Double, _
ByVal PROB_SUCCESS As Double, _
Optional ByVal CUMUL_FLAG As Boolean = True, _
Optional ByVal COMP_FLAG As Boolean = True)

'NO_SUCCESS: is the number of successes in trials.
'SAMPLE_SIZE: is the number of independent trials.
'PROB_SUCESS: is the probability of success on each trial.

'CUMUL_FLAG: is a logical value that determines the form of the
'function. If cumulative is TRUE, then BINOMDIST returns the
'cumulative distribution function, which is the probability that
'there are at most number_s successes; if FALSE, it returns the
'probability mass function, which is the probability that there
'are number_s successes.

'COMP_FLAG: 1-cumulative distribution function

On Error GoTo ERROR_LABEL

Select Case CUMUL_FLAG
    Case True
        If COMP_FLAG = True Then
            BINOMDIST_FUNC = cdf_binomial(SAMPLE_SIZE, NO_SUCCESS, PROB_SUCCESS)
            'SAME as BINOMDIST(NO_SUCCESS, SAMPLE_SIZE, PROB_SUCCESS,TRUE)
        ElseIf COMP_FLAG = False Then
            BINOMDIST_FUNC = comp_cdf_binomial(SAMPLE_SIZE, NO_SUCCESS, PROB_SUCCESS)
            'SAME as 1 - BINOMDIST(NO_SUCCESS, SAMPLE_SIZE, PROB_SUCCESS,TRUE)
        End If
    Case False 'probability density function
        BINOMDIST_FUNC = pmf_binomial(SAMPLE_SIZE, NO_SUCCESS, PROB_SUCCESS)
        'SAME as BINOMDIST(NO_SUCCESS, SAMPLE_SIZE, PROB_SUCCESS,FALSE)
End Select

Exit Function
ERROR_LABEL:
BINOMDIST_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : CRIT_BINOMDIST_FUNC
'DESCRIPTION   : Returns the smallest value for which the cumulative binomial
'distribution is greater than or equal to a criterion value.
'Use this function for quality assurance applications. For
'example, use this function to determine the greatest number of
'defective parts that are allowed to come off an assembly line
'run without rejecting the entire lot.
'LIBRARY       : STATISTICS
'GROUP         : DIST_DISCRETE
'ID            : 002
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/22/2009
'************************************************************************************
'************************************************************************************

Function CRIT_BINOMDIST_FUNC(ByVal SAMPLE_SIZE As Double, _
ByVal PROB_SUCCESS As Double, _
ByVal ALPHA As Double, _
Optional ByVal COMP_FLAG As Boolean = True)

'SAMPLE_SIZE: is the number of independent trials.
'PROB_SUCESS: is the probability of success on each trial.
'Alpha: is the criterion value.

On Error GoTo ERROR_LABEL

Select Case COMP_FLAG
    Case True
        CRIT_BINOMDIST_FUNC = crit_binomial(SAMPLE_SIZE, PROB_SUCCESS, ALPHA)
    Case False
        CRIT_BINOMDIST_FUNC = comp_crit_binomial(SAMPLE_SIZE, PROB_SUCCESS, ALPHA)
End Select

Exit Function
ERROR_LABEL:
CRIT_BINOMDIST_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : CONF_BINOMDIST_FUNC
'DESCRIPTION   : Lower / Upper Binomial CONFIDENCE Bound
'LIBRARY       : STATISTICS
'GROUP         : DIST_DISCRETE
'ID            : 003
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/22/2009
'************************************************************************************
'************************************************************************************

Function CONF_BINOMDIST_FUNC(ByVal NO_SUCCESS As Double, _
ByVal SAMPLE_SIZE As Double, _
ByVal PROBABILITY_VAL As Double, _
Optional ByVal LOWER_FLAG As Boolean = True)

'NO_SUCCESS: is the number of successes in trials.
'SAMPLE_SIZE: is the number of independent trials.
'PROBABILITY_VAL: prob {conf limits are 100(1-prob)%}

On Error GoTo ERROR_LABEL

Select Case LOWER_FLAG
    Case True
        CONF_BINOMDIST_FUNC = lcb_binomial(SAMPLE_SIZE, NO_SUCCESS, PROBABILITY_VAL)
    Case False
        CONF_BINOMDIST_FUNC = ucb_binomial(SAMPLE_SIZE, NO_SUCCESS, PROBABILITY_VAL)
End Select

Exit Function
ERROR_LABEL:
CONF_BINOMDIST_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : NEGBINOMDIST_FUNC
'DESCRIPTION   : Returns the negative BINOMIAL distribution. NEGBINOMDIST returns
'the probability that there will be number_f failures before the
'number_s-th success, when the constant probability of a success
'is probability_s. This function is similar to the BINOMIAL distribution,
'except that the number of successes is fixed, and the number of NO_FAILURES
'is variable. Like the BINOMIAL, NO_FAILURES are assumed to be independent.
'For example, you need to find 10 people with excellent reflexes, and you
'know the probability that a candidate has these qualifications is 0.3.
'NEGBINOMDIST calculates the probability that you will interview a certain
'number of unqualified candidates before finding all 10 qualified candidates
'LIBRARY       : STATISTICS
'GROUP         : DIST_DISCRETE
'ID            : 004
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/22/2009
'************************************************************************************
'************************************************************************************


Function NEGBINOMDIST_FUNC(ByVal THRESHOLD As Double, _
ByVal NO_FAILURES As Double, _
ByVal PROB_SUCCESS As Double, _
Optional ByVal CUMUL_FLAG As Boolean = True, _
Optional ByVal COMP_FLAG As Boolean = True)

'THRESHOLD: is the threshold number of successes.
'NO_FAILURES: is the number of failures.
'PROB_SUCESS: is the probability of success on each trial.

'COMP_FLAG: 1-cumulative distribution function

On Error GoTo ERROR_LABEL

Select Case CUMUL_FLAG
    Case True
        If COMP_FLAG = True Then
            NEGBINOMDIST_FUNC = cdf_negbinomial(NO_FAILURES, _
            PROB_SUCCESS, THRESHOLD)
        ElseIf COMP_FLAG = False Then
            NEGBINOMDIST_FUNC = comp_cdf_negbinomial(NO_FAILURES, _
            PROB_SUCCESS, THRESHOLD)
        End If
    Case False 'probability density function
        NEGBINOMDIST_FUNC = pmf_negbinomial(NO_FAILURES, _
            PROB_SUCCESS, THRESHOLD)
End Select

Exit Function
ERROR_LABEL:
NEGBINOMDIST_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : CRIT_NEGBINOMDIST_FUNC
'DESCRIPTION   : Returns the smallest value for which the cumulative NEGBINOMIAL
'distribution is greater than or equal to a criterion value.
'LIBRARY       : STATISTICS
'GROUP         : DIST_DISCRETE
'ID            : 005
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/22/2009
'************************************************************************************
'************************************************************************************

Function CRIT_NEGBINOMDIST_FUNC(ByVal THRESHOLD As Double, _
ByVal PROB_SUCCESS As Double, _
ByVal ALPHA As Double, _
Optional ByVal COMP_FLAG As Boolean = True)

'ALPHA: Crit Prob Value

On Error GoTo ERROR_LABEL

Select Case COMP_FLAG
    Case True
        CRIT_NEGBINOMDIST_FUNC = crit_negbinomial(PROB_SUCCESS, _
            THRESHOLD, ALPHA)
    Case False
        CRIT_NEGBINOMDIST_FUNC = comp_crit_negbinomial(PROB_SUCCESS, _
            THRESHOLD, ALPHA)
End Select

Exit Function
ERROR_LABEL:
CRIT_NEGBINOMDIST_FUNC = Err.number
End Function



'************************************************************************************
'************************************************************************************
'FUNCTION      : CONF_NEGBINOMDIST_FUNC
'DESCRIPTION   : Lower / Upper NEGBINOMIAL CONFIDENCE Bound
'LIBRARY       : STATISTICS
'GROUP         : DIST_DISCRETE
'ID            : 006
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/22/2009
'************************************************************************************
'************************************************************************************

Function CONF_NEGBINOMDIST_FUNC(ByVal THRESHOLD As Double, _
ByVal NO_FAILURES As Double, _
ByVal PROBABILITY_VAL As Double, _
Optional ByVal LOWER_FLAG As Boolean = True)

'PROBABILITY_VAL: prob {conf limits are 100(1-prob)%}

On Error GoTo ERROR_LABEL

Select Case LOWER_FLAG
    Case True
        CONF_NEGBINOMDIST_FUNC = lcb_negbinomial(NO_FAILURES, _
            THRESHOLD, PROBABILITY_VAL)
    Case False
        CONF_NEGBINOMDIST_FUNC = ucb_negbinomial(NO_FAILURES, _
            THRESHOLD, PROBABILITY_VAL)
End Select

Exit Function
ERROR_LABEL:
CONF_NEGBINOMDIST_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : GEOMETRIC_DIST_FUNC
'DESCRIPTION   : Typically, a Geometric random variable is the number
'of trials required to obtain the first failure, for example, the
'number of tosses of a coin untill the first 'tail' is obtained,
'or a process where components from a production line are tested,
'in turn, until the first defective item is found.
'LIBRARY       : STATISTICS
'GROUP         : DIST_DISCRETE
'ID            : 007
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/22/2009
'************************************************************************************
'************************************************************************************

Function GEOMETRIC_DIST_FUNC(ByVal NO_FAILURES As Double, _
ByVal PROB_SUCCESS As Double, _
Optional ByVal CUMUL_FLAG As Boolean = True, _
Optional ByVal COMP_FLAG As Boolean = True)

'NO_FAILURES: is the number of failures in trials.
'PROB_SUCESS: is the probability of success on each trial.

'CUMUL_FLAG: is a logical value that determines the form of the
'function. If cumulative is TRUE, then it returns the
'cumulative distribution function; if FALSE, it returns the
'probability mass function.

'COMP_FLAG: 1-cumulative distribution function

On Error GoTo ERROR_LABEL

Select Case CUMUL_FLAG
    Case True
        If COMP_FLAG = True Then
            GEOMETRIC_DIST_FUNC = cdf_geometric(NO_FAILURES, PROB_SUCCESS)
        ElseIf COMP_FLAG = False Then
            GEOMETRIC_DIST_FUNC = comp_cdf_geometric(NO_FAILURES, PROB_SUCCESS)
        End If
    Case False 'probability density function
        GEOMETRIC_DIST_FUNC = pmf_geometric(NO_FAILURES, PROB_SUCCESS)
End Select

Exit Function
ERROR_LABEL:
GEOMETRIC_DIST_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : CRIT_GEOMETRIC_DIST_FUNC
'DESCRIPTION   : Returns the smallest value for which the cumulative GEOMETRIC
'distribution is greater than or equal to a criterion value.
'LIBRARY       : STATISTICS
'GROUP         : DIST_DISCRETE
'ID            : 008
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/22/2009
'************************************************************************************
'************************************************************************************

Function CRIT_GEOMETRIC_DIST_FUNC(ByVal PROB_SUCCESS As Double, _
ByVal ALPHA As Double, _
Optional ByVal COMP_FLAG As Boolean = True)

'PROB_SUCESS: is the probability of success on each trial.
'Alpha: is the criterion value.

On Error GoTo ERROR_LABEL

Select Case COMP_FLAG
    Case True
        CRIT_GEOMETRIC_DIST_FUNC = crit_geometric(PROB_SUCCESS, ALPHA)
    Case False
        CRIT_GEOMETRIC_DIST_FUNC = comp_crit_geometric(PROB_SUCCESS, ALPHA)
End Select

Exit Function
ERROR_LABEL:
CRIT_GEOMETRIC_DIST_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : CONF_GEOMETRIC_DIST_FUNC
'DESCRIPTION   : Lower / Upper GEOMETRIC CONFIDENCE Bound
'LIBRARY       : STATISTICS
'GROUP         : DIST_DISCRETE
'ID            : 009
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/22/2009
'************************************************************************************
'************************************************************************************

Function CONF_GEOMETRIC_DIST_FUNC(ByVal NO_FAILURES As Double, _
ByVal PROBABILITY_VAL As Double, _
Optional ByVal LOWER_FLAG As Boolean = True)

'NO_FAILURES: is the number of failures in trials.
'PROBABILITY_VAL: prob {conf limits are 100(1-prob)%}

On Error GoTo ERROR_LABEL

Select Case LOWER_FLAG
    Case True
        CONF_GEOMETRIC_DIST_FUNC = lcb_geometric(NO_FAILURES, PROBABILITY_VAL)
    Case False
        CONF_GEOMETRIC_DIST_FUNC = ucb_geometric(NO_FAILURES, PROBABILITY_VAL)
End Select

Exit Function
ERROR_LABEL:
CONF_GEOMETRIC_DIST_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : POISSON_DIST_FUNC
'DESCRIPTION   : Poisson random variable is a count of the number of events
'that occur in a certain time interval or spatial area. For example,
'the number of cars passing a fixed point in a 5 minute interval, or
'the number of calls received by a switchboard during a given period
'of time.
'LIBRARY       : STATISTICS
'GROUP         : DIST_DISCRETE
'ID            : 010
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/22/2009
'************************************************************************************
'************************************************************************************

Function POISSON_DIST_FUNC(ByVal MEAN_VAL As Double, _
ByVal NO_EVENTS As Double, _
Optional ByVal CUMUL_FLAG As Boolean = True, _
Optional ByVal COMP_FLAG As Boolean = True)

'MEAN_VAL: is the expected numeric value.
'NO_EVENTS: Number of Events

'CUMUL_FLAG: is a logical value that determines the form of the
'function. If cumulative is TRUE, then it returns the
'cumulative distribution function; if FALSE, it returns the
'probability mass function.

'COMP_FLAG: 1-cumulative distribution function

On Error GoTo ERROR_LABEL

Select Case CUMUL_FLAG
    Case True
        If COMP_FLAG = True Then
            POISSON_DIST_FUNC = cdf_poisson(MEAN_VAL, NO_EVENTS)
        ElseIf COMP_FLAG = False Then
            POISSON_DIST_FUNC = comp_cdf_poisson(MEAN_VAL, NO_EVENTS)
        End If
    Case False 'probability density function
        POISSON_DIST_FUNC = pmf_poisson(MEAN_VAL, NO_EVENTS)
End Select

Exit Function
ERROR_LABEL:
POISSON_DIST_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : CRIT_POISSON_DIST_FUNC
'DESCRIPTION   : Returns the smallest value for which the cumulative POISSON
'distribution is greater than or equal to a criterion value.
'LIBRARY       : STATISTICS
'GROUP         : DIST_DISCRETE
'ID            : 011
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/22/2009
'************************************************************************************
'************************************************************************************

Function CRIT_POISSON_DIST_FUNC(ByVal NO_EVENTS As Double, _
ByVal ALPHA As Double, _
Optional ByVal COMP_FLAG As Boolean = True)

'NO_EVENTS: Number of Events
'Alpha: is the criterion value.

On Error GoTo ERROR_LABEL

Select Case COMP_FLAG
    Case True
        CRIT_POISSON_DIST_FUNC = crit_poisson(NO_EVENTS, ALPHA)
    Case False
        CRIT_POISSON_DIST_FUNC = comp_crit_poisson(NO_EVENTS, ALPHA)
End Select

Exit Function
ERROR_LABEL:
CRIT_POISSON_DIST_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : CONF_POISSON_DIST_FUNC
'DESCRIPTION   : Lower / Upper POISSON CONFIDENCE Bound
'LIBRARY       : STATISTICS
'GROUP         : DIST_DISCRETE
'ID            : 012
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/22/2009
'************************************************************************************
'************************************************************************************

Function CONF_POISSON_DIST_FUNC(ByVal NO_EVENTS As Double, _
ByVal PROBABILITY_VAL As Double, _
Optional ByVal LOWER_FLAG As Boolean = True)

'NO_EVENTS: is the number of events.
'PROBABILITY_VAL: prob {conf limits are 100(1-prob)%}

On Error GoTo ERROR_LABEL

Select Case LOWER_FLAG
    Case True
        CONF_POISSON_DIST_FUNC = lcb_poisson(NO_EVENTS, PROBABILITY_VAL)
    Case False
        CONF_POISSON_DIST_FUNC = ucb_poisson(NO_EVENTS, PROBABILITY_VAL)
End Select

Exit Function
ERROR_LABEL:
CONF_POISSON_DIST_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : HYPERGEOM_DIST_FUNC
'DESCRIPTION   : Returns the hypergeometric distribution. It returns the
'probability of a given number of sample successes, given the sample
'size, population successes, and population size. Use this for
'problems with a finite population, where each observation is either
'a success or a failure, and where each subset of a given size is
'chosen with equal likelihood.
'LIBRARY       : STATISTICS
'GROUP         : DIST_DISCRETE
'ID            : 013
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/22/2009
'************************************************************************************
'************************************************************************************

Function HYPERGEOM_DIST_FUNC(ByVal NO_SUCCESS As Double, _
ByVal SAMPLE_SIZE As Double, _
ByVal POPULATION_SUCCESS As Double, _
ByVal POPULATION_SIZE As Double, _
Optional ByVal CUMUL_FLAG As Boolean = True, _
Optional ByVal COMP_FLAG As Boolean = True)

'NO_SUCCESS: is the number of successes in the sample.
'SAMPLE_SIZE: is the size of the sample.
'POPULATION_SUCCESS: is the number of successes in the population.
'POPULATION_SIZE: is the population size.


'CUMUL_FLAG: is a logical value that determines the form of the
'function. If cumulative is TRUE, then it returns the
'cumulative distribution function, which is the probability that
'there are at most number_s successes; if FALSE, it returns the
'probability mass function, which is the probability that there
'are number_s successes.

'COMP_FLAG: 1-cumulative distribution function

On Error GoTo ERROR_LABEL

Select Case CUMUL_FLAG
    Case True
        If COMP_FLAG = True Then
            HYPERGEOM_DIST_FUNC = cdf_hypergeometric(NO_SUCCESS, SAMPLE_SIZE, _
                POPULATION_SUCCESS, POPULATION_SIZE)
        ElseIf COMP_FLAG = False Then
            HYPERGEOM_DIST_FUNC = comp_cdf_hypergeometric(NO_SUCCESS, _
                SAMPLE_SIZE, POPULATION_SUCCESS, POPULATION_SIZE)
        End If
    Case False 'probability density function
        HYPERGEOM_DIST_FUNC = pmf_hypergeometric(NO_SUCCESS, SAMPLE_SIZE, _
            POPULATION_SUCCESS, POPULATION_SIZE)
End Select

Exit Function
ERROR_LABEL:
HYPERGEOM_DIST_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : CRIT_HYPERGEOM_DIST_FUNC
'DESCRIPTION   : Returns the smallest value for which the cumulative HYPERGEOM
'distribution is greater than or equal to a criterion value.
'LIBRARY       : STATISTICS
'GROUP         : DIST_DISCRETE
'ID            : 014
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/22/2009
'************************************************************************************
'************************************************************************************

Function CRIT_HYPERGEOM_DIST_FUNC(ByVal SAMPLE_SIZE As Double, _
ByVal POPULATION_SUCCESS As Double, _
ByVal POPULATION_SIZE As Double, _
ByVal ALPHA As Double, _
Optional ByVal COMP_FLAG As Boolean = True)

'SAMPLE_SIZE: is the size of the sample.
'POPULATION_SUCCESS: is the number of successes in the population.
'POPULATION_SIZE: is the population size.
'Alpha: is the criterion value.

On Error GoTo ERROR_LABEL

Select Case COMP_FLAG
    Case True
        CRIT_HYPERGEOM_DIST_FUNC = crit_hypergeometric(SAMPLE_SIZE, _
            POPULATION_SUCCESS, POPULATION_SIZE, ALPHA)
    Case False
        CRIT_HYPERGEOM_DIST_FUNC = comp_crit_hypergeometric(SAMPLE_SIZE, _
            POPULATION_SUCCESS, POPULATION_SIZE, ALPHA)
End Select

Exit Function
ERROR_LABEL:
CRIT_HYPERGEOM_DIST_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : CONF_HYPERGEOM_DIST_FUNC
'DESCRIPTION   : Lower / Upper HYPERGEOM CONFIDENCE Bound
'LIBRARY       : STATISTICS
'GROUP         : DIST_DISCRETE
'ID            : 015
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/22/2009
'************************************************************************************
'************************************************************************************

Function CONF_HYPERGEOM_DIST_FUNC(ByVal NO_SUCCESS As Double, _
ByVal SAMPLE_SIZE As Double, _
ByVal POPULATION_SIZE As Double, _
ByVal PROBABILITY_VAL As Double, _
Optional ByVal LOWER_FLAG As Boolean = True)

'NO_SUCCESS: is the number of successes in the sample.
'SAMPLE_SIZE: is the size of the sample.
'POPULATION_SIZE: is the population size.
'PROBABILITY_VAL: prob {conf limits are 100(1-prob)%}

On Error GoTo ERROR_LABEL

Select Case LOWER_FLAG
    Case True
        CONF_HYPERGEOM_DIST_FUNC = lcb_hypergeometric(NO_SUCCESS, _
            SAMPLE_SIZE, POPULATION_SIZE, PROBABILITY_VAL)
    Case False
        CONF_HYPERGEOM_DIST_FUNC = ucb_hypergeometric(NO_SUCCESS, SAMPLE_SIZE, _
            POPULATION_SIZE, PROBABILITY_VAL)
End Select

Exit Function
ERROR_LABEL:
CONF_HYPERGEOM_DIST_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : NEGHYPERGEOM_DIST_FUNC
'DESCRIPTION   : Returns the neghypergeometric distribution.
'LIBRARY       : STATISTICS
'GROUP         : DIST_DISCRETE
'ID            : 016
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/22/2009
'************************************************************************************
'************************************************************************************

Function NEGHYPERGEOM_DIST_FUNC(ByVal NO_SUCCESS As Double, _
ByVal SAMPLE_SIZE As Double, _
ByVal POPULATION_SUCCESS As Double, _
ByVal POPULATION_SIZE As Double, _
Optional ByVal CUMUL_FLAG As Boolean = True, _
Optional ByVal COMP_FLAG As Boolean = True)

'NO_SUCCESS: is the number of successes in the sample.
'SAMPLE_SIZE: is the size of the sample.
'POPULATION_SUCCESS: is the number of successes in the population.
'POPULATION_SIZE: is the population size.

'CUMUL_FLAG: is a logical value that determines the form of the
'function. If cumulative is TRUE, then it returns the
'cumulative distribution function, which is the probability that
'there are at most number_s successes; if FALSE, it returns the
'probability mass function, which is the probability that there
'are number_s successes.

'COMP_FLAG: 1-cumulative distribution function

On Error GoTo ERROR_LABEL

Select Case CUMUL_FLAG
    Case True
        If COMP_FLAG = True Then
            NEGHYPERGEOM_DIST_FUNC = cdf_neghypergeometric(NO_SUCCESS, SAMPLE_SIZE, _
                POPULATION_SUCCESS, POPULATION_SIZE)
        ElseIf COMP_FLAG = False Then
            NEGHYPERGEOM_DIST_FUNC = comp_cdf_neghypergeometric(NO_SUCCESS, _
                SAMPLE_SIZE, POPULATION_SUCCESS, POPULATION_SIZE)
        End If
    Case False 'probability density function
        NEGHYPERGEOM_DIST_FUNC = pmf_neghypergeometric(NO_SUCCESS, SAMPLE_SIZE, _
            POPULATION_SUCCESS, POPULATION_SIZE)
End Select

Exit Function
ERROR_LABEL:
NEGHYPERGEOM_DIST_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : CRIT_NEGHYPERGEOM_DIST_FUNC
'DESCRIPTION   : Returns the smallest value for which the cumulative NEGHYPERGEOM
'distribution is greater than or equal to a criterion value.
'LIBRARY       : STATISTICS
'GROUP         : DIST_DISCRETE
'ID            : 017
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/22/2009
'************************************************************************************
'************************************************************************************

Function CRIT_NEGHYPERGEOM_DIST_FUNC(ByVal SAMPLE_SIZE As Double, _
ByVal POPULATION_SUCCESS As Double, _
ByVal POPULATION_SIZE As Double, _
ByVal ALPHA As Double, _
Optional ByVal COMP_FLAG As Boolean = True)

'SAMPLE_SIZE: is the size of the sample.
'POPULATION_SUCCESS: is the number of successes in the population.
'POPULATION_SIZE: is the population size.
'Alpha: is the criterion value.

On Error GoTo ERROR_LABEL

Select Case COMP_FLAG
    Case True
        CRIT_NEGHYPERGEOM_DIST_FUNC = crit_neghypergeometric(SAMPLE_SIZE, _
            POPULATION_SUCCESS, POPULATION_SIZE, ALPHA)
    Case False
        CRIT_NEGHYPERGEOM_DIST_FUNC = comp_crit_neghypergeometric(SAMPLE_SIZE, _
            POPULATION_SUCCESS, POPULATION_SIZE, ALPHA)
End Select

Exit Function
ERROR_LABEL:
CRIT_NEGHYPERGEOM_DIST_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : CONF_NEGHYPERGEOM_DIST_FUNC
'DESCRIPTION   : Lower / Upper NEGHYPERGEOM CONFIDENCE Bound
'LIBRARY       : STATISTICS
'GROUP         : DIST_DISCRETE
'ID            : 018
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/22/2009
'************************************************************************************
'************************************************************************************

Function CONF_NEGHYPERGEOM_DIST_FUNC(ByVal NO_SUCCESS As Double, _
ByVal SAMPLE_SIZE As Double, _
ByVal POPULATION_SIZE As Double, _
ByVal PROBABILITY_VAL As Double, _
Optional ByVal LOWER_FLAG As Boolean = True)

'NO_SUCCESS: is the number of successes in the sample.
'SAMPLE_SIZE: is the size of the sample.
'POPULATION_SIZE: is the population size.
'PROBABILITY_VAL: prob {conf limits are 100(1-prob)%}

On Error GoTo ERROR_LABEL

Select Case LOWER_FLAG
    Case True
        CONF_NEGHYPERGEOM_DIST_FUNC = lcb_neghypergeometric(NO_SUCCESS, _
            SAMPLE_SIZE, POPULATION_SIZE, PROBABILITY_VAL)
    Case False
        CONF_NEGHYPERGEOM_DIST_FUNC = ucb_neghypergeometric(NO_SUCCESS, _
            SAMPLE_SIZE, POPULATION_SIZE, PROBABILITY_VAL)
End Select

Exit Function
ERROR_LABEL:
CONF_NEGHYPERGEOM_DIST_FUNC = Err.number
End Function
