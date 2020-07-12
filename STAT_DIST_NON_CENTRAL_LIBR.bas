Attribute VB_Name = "STAT_DIST_NON_CENTRAL_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.

                            
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.


'************************************************************************************
'************************************************************************************
'FUNCTION      : NON_CENTRAL_GAMMA_DIST_FUNC
'DESCRIPTION   : Returns the non-central GAMMA_NC distribution.
'LIBRARY       : STATISTICS
'GROUP         : DIST_NON_CENTRAL
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/22/2009
'************************************************************************************
'************************************************************************************

Function NON_CENTRAL_GAMMA_DIST_FUNC(ByVal X_VAL As Double, _
ByVal ALPHA As Double, _
ByVal BETA_NC As Double, _
Optional ByVal CUMUL_FLAG As Boolean = True, _
Optional ByVal COMP_FLAG As Boolean = True)

'X_VAL: is the value at which you want to evaluate the distribution.

'ALPHA: is a shape parameter to the distribution.

'BETA_NC: is a noncentrality parameter to the distribution.

'CUMUL_FLAG: is a logical value that determines the form of the function. If
'cumulative is TRUE, NON_CENTRAL_GAMMA_DIST_FUNC returns the cumulative distribution function;
'if FALSE, it returns the probability density function.

'COMP_FLAG: 1-cumulative distribution function

On Error GoTo ERROR_LABEL

    Select Case CUMUL_FLAG
        
        Case True
        
            If COMP_FLAG = True Then
                NON_CENTRAL_GAMMA_DIST_FUNC = cdf_gamma_nc(X_VAL, ALPHA, BETA_NC)
            ElseIf COMP_FLAG = False Then
                NON_CENTRAL_GAMMA_DIST_FUNC = comp_cdf_gamma_nc(X_VAL, ALPHA, BETA_NC)
            End If

        Case False 'probability density function
        
        NON_CENTRAL_GAMMA_DIST_FUNC = pdf_gamma_nc(X_VAL, ALPHA, BETA_NC)
    
    End Select

Exit Function
ERROR_LABEL:
NON_CENTRAL_GAMMA_DIST_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : INVERSE_NON_CENTRAL_GAMMA_DIST_FUNC
'DESCRIPTION   : Returns the inverse of the GAMMA_NC cumulative distribution.
'LIBRARY       : STATISTICS
'GROUP         : DIST_NON_CENTRAL
'ID            : 002
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/22/2009
'************************************************************************************
'************************************************************************************

Function INVERSE_NON_CENTRAL_GAMMA_DIST_FUNC(ByVal PROBABILITY_VAL As Double, _
ByVal ALPHA As Double, _
ByVal BETA_NC As Double, _
Optional ByVal COMP_FLAG As Boolean = True)

'PROBABILITY_VAL: is the probability associated with the prob {used by inv fns}

'ALPHA: is a shape parameter to the distribution.

'BETA_NC: is a noncentrality parameter to the distribution.

'COMP_FLAG: 1-cumulative distribution function

On Error GoTo ERROR_LABEL

    Select Case COMP_FLAG
        Case True
            INVERSE_NON_CENTRAL_GAMMA_DIST_FUNC = inv_gamma_nc(PROBABILITY_VAL, ALPHA, BETA_NC)
        Case False
            INVERSE_NON_CENTRAL_GAMMA_DIST_FUNC = comp_inv_gamma_nc(PROBABILITY_VAL, ALPHA, BETA_NC)
    End Select

Exit Function
ERROR_LABEL:
INVERSE_NON_CENTRAL_GAMMA_DIST_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : CDF_NON_CENTRAL_GAMMA_DIST_FUNC
'DESCRIPTION   : Find non-centrality parameter with corresponding cdf value on a
'gamma distribution
'LIBRARY       : STATISTICS
'GROUP         : DIST_NON_CENTRAL
'ID            : 003
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/22/2009
'************************************************************************************
'************************************************************************************

Function CDF_NON_CENTRAL_GAMMA_DIST_FUNC(ByVal PROBABILITY_VAL As Double, _
ByVal X_VAL As Double, _
ByVal ALPHA As Double, _
Optional ByVal COMP_FLAG As Boolean = True)

'PROBABILITY_VAL: is the probability associated with the prob {used by ncp fns}

'X_VAL: is the value at which you want to evaluate the distribution.

'ALPHA: is a shape parameter to the distribution.

'COMP_FLAG: 1-cumulative distribution function

On Error GoTo ERROR_LABEL

    Select Case COMP_FLAG
        Case True 'Find non-centrality parameter with corresponding cdf value
            CDF_NON_CENTRAL_GAMMA_DIST_FUNC = ncp_gamma_nc(PROBABILITY_VAL, X_VAL, ALPHA)
        Case False 'Find non-centrality parameter with corresponding comp_cdf value
            CDF_NON_CENTRAL_GAMMA_DIST_FUNC = comp_ncp_gamma_nc(PROBABILITY_VAL, X_VAL, ALPHA)
    End Select

Exit Function
ERROR_LABEL:
CDF_NON_CENTRAL_GAMMA_DIST_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : NON_CENTRAL_CHI_SQUARED_DIST_FUNC
'DESCRIPTION   : Returns the non-central CHI_SQ distribution
'LIBRARY       : STATISTICS
'GROUP         : DIST_NON_CENTRAL
'ID            : 004
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/22/2009
'************************************************************************************
'************************************************************************************

Function NON_CENTRAL_CHI_SQUARED_DIST_FUNC(ByVal X_VAL As Double, _
ByVal DEG_FREEDOM As Double, _
ByVal BETA_NC As Double, _
Optional ByVal CUMUL_FLAG As Boolean = True, _
Optional ByVal COMP_FLAG As Boolean = True)

'X_VAL: is the value at which you want to evaluate the distribution.

'DEG_FREEDOM: is the number of degrees of freedom.

'BETA_NC: is a noncentrality parameter to the distribution.

'CUMUL_FLAG: is a logical value that determines the form of the function. If
'cumulative is TRUE, NON_CENTRAL_CHI_SQUARED_DIST_FUNC returns the cumulative distribution function;
'if FALSE, it returns the probability density function.

'COMP_FLAG: 1-cumulative distribution function

On Error GoTo ERROR_LABEL

    Select Case CUMUL_FLAG
        
        Case True
        
            If COMP_FLAG = True Then
                NON_CENTRAL_CHI_SQUARED_DIST_FUNC = cdf_Chi2_nc(X_VAL, DEG_FREEDOM, BETA_NC)
            ElseIf COMP_FLAG = False Then
                NON_CENTRAL_CHI_SQUARED_DIST_FUNC = comp_cdf_Chi2_nc(X_VAL, DEG_FREEDOM, BETA_NC)
            End If

        Case False 'probability density function
        
        NON_CENTRAL_CHI_SQUARED_DIST_FUNC = pdf_Chi2_nc(X_VAL, DEG_FREEDOM, BETA_NC)
    
    End Select

Exit Function
ERROR_LABEL:
NON_CENTRAL_CHI_SQUARED_DIST_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : INVERSE_NON_CENTRAL_CHI_SQUARED_DIST_FUNC
'DESCRIPTION   : Returns the inverse of the CHI_SQ_NC cumulative distribution
'LIBRARY       : STATISTICS
'GROUP         : DIST_NON_CENTRAL
'ID            : 005
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/22/2009
'************************************************************************************
'************************************************************************************

Function INVERSE_NON_CENTRAL_CHI_SQUARED_DIST_FUNC(ByVal PROBABILITY_VAL As Double, _
ByVal DEG_FREEDOM As Double, _
ByVal BETA_NC As Double, _
Optional ByVal COMP_FLAG As Boolean = True)

'PROBABILITY_VAL: is the probability associated with the prob {used by inv fns}

'DEG_FREEDOM: is the number of degrees of freedom.

'BETA_NC: is a noncentrality parameter to the distribution.

'COMP_FLAG: 1-cumulative distribution function

On Error GoTo ERROR_LABEL

    Select Case COMP_FLAG
        Case True
            INVERSE_NON_CENTRAL_CHI_SQUARED_DIST_FUNC = inv_Chi2_nc(PROBABILITY_VAL, DEG_FREEDOM, BETA_NC)
        Case False
            INVERSE_NON_CENTRAL_CHI_SQUARED_DIST_FUNC = comp_inv_Chi2_nc(PROBABILITY_VAL, DEG_FREEDOM, BETA_NC)
    End Select

Exit Function
ERROR_LABEL:
INVERSE_NON_CENTRAL_CHI_SQUARED_DIST_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : CDF_NON_CENTRAL_CHI_SQUARED_DIST_FUNC
'DESCRIPTION   : Find non-centrality parameter with corresponding cdf value on a
'chi-squared distribution function
'LIBRARY       : STATISTICS
'GROUP         : DIST_NON_CENTRAL
'ID            : 006
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/22/2009
'************************************************************************************
'************************************************************************************

Function CDF_NON_CENTRAL_CHI_SQUARED_DIST_FUNC(ByVal PROBABILITY_VAL As Double, _
ByVal X_VAL As Double, _
ByVal DEG_FREEDOM As Double, _
Optional ByVal COMP_FLAG As Boolean = True)

'PROBABILITY_VAL: is the probability associated with the prob {used by ncp fns}

'X_VAL: is the value at which you want to evaluate the distribution.

'DEG_FREEDOM: is the number of degrees of freedom.

'COMP_FLAG: 1-cumulative distribution function

On Error GoTo ERROR_LABEL

    Select Case COMP_FLAG
        Case True 'Find non-centrality parameter with corresponding cdf value
            CDF_NON_CENTRAL_CHI_SQUARED_DIST_FUNC = ncp_Chi2_nc(PROBABILITY_VAL, X_VAL, DEG_FREEDOM)
        Case False 'Find non-centrality parameter with corresponding comp_cdf value
            CDF_NON_CENTRAL_CHI_SQUARED_DIST_FUNC = comp_ncp_Chi2_nc(PROBABILITY_VAL, X_VAL, DEG_FREEDOM)
    End Select

Exit Function
ERROR_LABEL:
CDF_NON_CENTRAL_CHI_SQUARED_DIST_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : NON_CENTRAL_BETA_DIST_FUNC
'DESCRIPTION   : Returns the non-central BETA_NC distribution
'LIBRARY       : STATISTICS
'GROUP         : DIST_NON_CENTRAL
'ID            : 007
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/22/2009
'************************************************************************************
'************************************************************************************

Function NON_CENTRAL_BETA_DIST_FUNC(ByVal X_VAL As Double, _
ByVal FIRST_ALPHA As Double, _
ByVal SECOND_ALPHA As Double, _
ByVal BETA_NC As Double, _
Optional ByVal CUMUL_FLAG As Boolean = True, _
Optional ByVal COMP_FLAG As Boolean = True)

'X_VAL: is the value at which you want to evaluate the distribution.

'FIRST_ALPHA: is a shape parameter A to the distribution.

'SECOND_ALPHA: is a shape parameter B to the distribution.

'BETA_NC: is a noncentrality parameter to the distribution.

'CUMUL_FLAG: is a logical value that determines the form of the function. If
'cumulative is TRUE, NON_CENTRAL_BETA_DIST_FUNC returns the cumulative distribution function;
'if FALSE, it returns the probability density function.

'COMP_FLAG: 1-cumulative distribution function

On Error GoTo ERROR_LABEL

    Select Case CUMUL_FLAG
        
        Case True
        
            If COMP_FLAG = True Then
                NON_CENTRAL_BETA_DIST_FUNC = cdf_BETA_nc(X_VAL, FIRST_ALPHA, _
                SECOND_ALPHA, BETA_NC)
            ElseIf COMP_FLAG = False Then
                NON_CENTRAL_BETA_DIST_FUNC = comp_cdf_BETA_nc(X_VAL, FIRST_ALPHA, _
                SECOND_ALPHA, BETA_NC)
            End If

        Case False 'probability density function
        
        NON_CENTRAL_BETA_DIST_FUNC = pdf_BETA_nc(X_VAL, FIRST_ALPHA, _
        SECOND_ALPHA, BETA_NC)
    
    End Select

Exit Function
ERROR_LABEL:
NON_CENTRAL_BETA_DIST_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : INVERSE_NON_CENTRAL_BETA_DIST_FUNC
'DESCRIPTION   : Returns the inverse of the BETA_NC cumulative distribution
'LIBRARY       : STATISTICS
'GROUP         : DIST_NON_CENTRAL
'ID            : 008
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/22/2009
'************************************************************************************
'************************************************************************************

Function INVERSE_NON_CENTRAL_BETA_DIST_FUNC(ByVal PROBABILITY_VAL As Double, _
ByVal FIRST_ALPHA As Double, _
ByVal SECOND_ALPHA As Double, _
ByVal BETA_NC As Double, _
Optional ByVal COMP_FLAG As Boolean = True)

'PROBABILITY_VAL: is the probability associated with the prob {used by inv fns}

'FIRST_ALPHA: is a shape parameter A to the distribution.

'SECOND_ALPHA: is a shape parameter B to the distribution.

'BETA_NC: is a noncentrality parameter to the distribution.

'COMP_FLAG: 1-cumulative distribution function

On Error GoTo ERROR_LABEL

    Select Case COMP_FLAG
        Case True
            INVERSE_NON_CENTRAL_BETA_DIST_FUNC = inv_BETA_nc(PROBABILITY_VAL, FIRST_ALPHA, _
            SECOND_ALPHA, BETA_NC)
        Case False
            INVERSE_NON_CENTRAL_BETA_DIST_FUNC = comp_inv_BETA_nc(PROBABILITY_VAL, FIRST_ALPHA, _
            SECOND_ALPHA, BETA_NC)
    End Select

Exit Function
ERROR_LABEL:
INVERSE_NON_CENTRAL_BETA_DIST_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : CDF_NON_CENTRAL_BETA_DIST_FUNC
'DESCRIPTION   : Find non-centrality parameter with corresponding cdf value on a BETA
'Distribution
'LIBRARY       : STATISTICS
'GROUP         : DIST_NON_CENTRAL
'ID            : 009
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/22/2009
'************************************************************************************
'************************************************************************************

Function CDF_NON_CENTRAL_BETA_DIST_FUNC(ByVal PROBABILITY_VAL As Double, _
ByVal X_VAL As Double, _
ByVal FIRST_ALPHA As Double, _
ByVal SECOND_ALPHA As Double, _
Optional ByVal COMP_FLAG As Boolean = True)

'PROBABILITY_VAL: is the probability associated with the prob {used by ncp fns}

'X_VAL: is the value at which you want to evaluate the distribution.

'FIRST_ALPHA: is a shape parameter A to the distribution.

'SECOND_ALPHA: is a shape parameter B to the distribution.

'COMP_FLAG: 1-cumulative distribution function

On Error GoTo ERROR_LABEL


    Select Case COMP_FLAG
        Case True 'Find non-centrality parameter with corresponding cdf value
            CDF_NON_CENTRAL_BETA_DIST_FUNC = ncp_BETA_nc(PROBABILITY_VAL, X_VAL, _
            FIRST_ALPHA, SECOND_ALPHA)
        Case False 'Find non-centrality parameter with corresponding comp_cdf value
            CDF_NON_CENTRAL_BETA_DIST_FUNC = comp_ncp_BETA_nc(PROBABILITY_VAL, X_VAL, _
            FIRST_ALPHA, SECOND_ALPHA)
    End Select

Exit Function
ERROR_LABEL:
CDF_NON_CENTRAL_BETA_DIST_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : NON_CENTRAL_FDIST_FUNC
'DESCRIPTION   : Returns the non-central F_NC distribution
'LIBRARY       : STATISTICS
'GROUP         : DIST_NON_CENTRAL
'ID            : 010
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/22/2009
'************************************************************************************
'************************************************************************************

Function NON_CENTRAL_FDIST_FUNC(ByVal X_VAL As Double, _
ByVal DEG_FREEDOM_1 As Double, _
ByVal DEG_FREEDOM_2 As Double, _
ByVal BETA_NC As Double, _
Optional ByVal CUMUL_FLAG As Boolean = True, _
Optional ByVal COMP_FLAG As Boolean = True)

'X_VAL: is the value at which you want to evaluate the distribution.

'DEG_FREEDOM_1: is a shape parameter A to the distribution.

'DEG_FREEDOM_2: is a shape parameter B to the distribution.

'F_NC: is a noncentrality parameter to the distribution.

'CUMUL_FLAG: is a logical value that determines the form of the function. If
'cumulative is TRUE, NON_CENTRAL_FDIST_FUNC returns the cumulative distribution function;
'if FALSE, it returns the probability density function.

'COMP_FLAG: 1-cumulative distribution function

On Error GoTo ERROR_LABEL


    Select Case CUMUL_FLAG
        
        Case True
        
            If COMP_FLAG = True Then
                NON_CENTRAL_FDIST_FUNC = cdf_fdist_nc(X_VAL, DEG_FREEDOM_1, _
                DEG_FREEDOM_2, BETA_NC)
            ElseIf COMP_FLAG = False Then
                NON_CENTRAL_FDIST_FUNC = comp_cdf_fdist_nc(X_VAL, DEG_FREEDOM_1, _
                DEG_FREEDOM_2, BETA_NC)
            End If

        Case False 'probability density function
        
        NON_CENTRAL_FDIST_FUNC = pdf_fdist_nc(X_VAL, DEG_FREEDOM_1, _
        DEG_FREEDOM_2, BETA_NC)
    
    End Select

Exit Function
ERROR_LABEL:
NON_CENTRAL_FDIST_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : INVERSE_NON_CENTRAL_FDIST_FUNC
'DESCRIPTION   : Returns the inverse of the F_NC cumulative distribution
'LIBRARY       : STATISTICS
'GROUP         : DIST_NON_CENTRAL
'ID            : 011
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/22/2009
'************************************************************************************
'************************************************************************************

Function INVERSE_NON_CENTRAL_FDIST_FUNC(ByVal PROBABILITY_VAL As Double, _
ByVal DEG_FREEDOM_1 As Double, _
ByVal DEG_FREEDOM_2 As Double, _
ByVal BETA_NC As Double, _
Optional ByVal COMP_FLAG As Boolean = True)

'PROBABILITY_VAL: is the probability associated with the prob {used by inv fns}

'DEG_FREEDOM_1: First degree of freedom

'DEG_FREEDOM_2: Second degree of freedom

'BETA_NC: is a noncentrality parameter to the distribution.

'COMP_FLAG: 1-cumulative distribution function

On Error GoTo ERROR_LABEL

    Select Case COMP_FLAG
        Case True
            INVERSE_NON_CENTRAL_FDIST_FUNC = inv_fdist_nc(PROBABILITY_VAL, DEG_FREEDOM_1, _
            DEG_FREEDOM_2, BETA_NC)
        Case False
            INVERSE_NON_CENTRAL_FDIST_FUNC = comp_inv_fdist_nc(PROBABILITY_VAL, DEG_FREEDOM_1, _
            DEG_FREEDOM_2, BETA_NC)
    End Select

Exit Function
ERROR_LABEL:
INVERSE_NON_CENTRAL_FDIST_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : CDF_NON_CENTRAL_FDIST_FUNC
'DESCRIPTION   : Find non-centrality parameter with corresponding cdf value on a
'F DIST
'LIBRARY       : STATISTICS
'GROUP         : DIST_NON_CENTRAL
'ID            : 012
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/22/2009
'************************************************************************************
'************************************************************************************

Function CDF_NON_CENTRAL_FDIST_FUNC(ByVal PROBABILITY_VAL As Double, _
ByVal X_VAL As Double, _
ByVal DEG_FREEDOM_1 As Double, _
ByVal DEG_FREEDOM_2 As Double, _
Optional ByVal COMP_FLAG As Boolean = True)

'PROBABILITY_VAL: is the probability associated with the prob {used by ncp fns}

'X_VAL: is the value at which you want to evaluate the distribution.

'DEG_FREEDOM_1: First degree of freedom

'DEG_FREEDOM_2: Second degree of freedom

'COMP_FLAG: 1-cumulative distribution function

On Error GoTo ERROR_LABEL

    Select Case COMP_FLAG
        Case True 'Find non-centrality parameter with corresponding cdf value
            CDF_NON_CENTRAL_FDIST_FUNC = ncp_fdist_nc(PROBABILITY_VAL, X_VAL, _
            DEG_FREEDOM_1, DEG_FREEDOM_2)
        Case False 'Find non-centrality parameter with corresponding comp_cdf value
            CDF_NON_CENTRAL_FDIST_FUNC = comp_ncp_fdist_nc(PROBABILITY_VAL, X_VAL, _
            DEG_FREEDOM_1, DEG_FREEDOM_2)
    End Select

Exit Function
ERROR_LABEL:
CDF_NON_CENTRAL_FDIST_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : NON_CENTRAL_TDIST_FUNC
'DESCRIPTION   : Returns the non-central T distribution
'LIBRARY       : STATISTICS
'GROUP         : DIST_NON_CENTRAL
'ID            : 013
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/22/2009
'************************************************************************************
'************************************************************************************

Function NON_CENTRAL_TDIST_FUNC(ByVal X_VAL As Double, _
ByVal DEG_FREEDOM As Double, _
ByVal BETA_NC As Double, _
Optional ByVal CUMUL_FLAG As Boolean = True, _
Optional ByVal COMP_FLAG As Boolean = True)

'X_VAL: is the value at which you want to evaluate the distribution.

'DEG_FREEDOM: is the number of degrees of freedom.

'BETA_NC: is a noncentrality parameter to the distribution.

'CUMUL_FLAG: is a logical value that determines the form of the function. If
'cumulative is TRUE, NON_CENTRAL_TDIST_FUNC returns the cumulative distribution function;
'if FALSE, it returns the probability density function.

'COMP_FLAG: 1-cumulative distribution function


'------------------------------------------------------------------------------
'------------------------------------------------------------------------------
'------------------------------------------------------------------------------
' The noncentral t distribution does not work when the X_VAL
' and the noncentrality parameter have different signs. This is
' because the algorithm used would be prone to potentially huge
' cancellation errors.

' For the same reason the inverse function will not work when the
' X_VAL is required to have the opposite sign to the
' noncentrality parameter.

' The noncentrality parameter is limited to Sqrt(2*nc_limit)
'------------------------------------------------------------------------------
'------------------------------------------------------------------------------
'------------------------------------------------------------------------------

On Error GoTo ERROR_LABEL

    Select Case CUMUL_FLAG
        
        Case True
        
            If COMP_FLAG = True Then
                NON_CENTRAL_TDIST_FUNC = cdf_t_nc(X_VAL, DEG_FREEDOM, BETA_NC)
            ElseIf COMP_FLAG = False Then
                NON_CENTRAL_TDIST_FUNC = comp_cdf_t_nc(X_VAL, DEG_FREEDOM, BETA_NC)
            End If

        Case False 'probability density function
        
        NON_CENTRAL_TDIST_FUNC = pdf_t_nc(X_VAL, DEG_FREEDOM, BETA_NC)
    
    End Select

Exit Function
ERROR_LABEL:
NON_CENTRAL_TDIST_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : INVERSE_NON_CENTRAL_TDIST_FUNC
'DESCRIPTION   : Returns the inverse of the T_NC cumulative distribution
'LIBRARY       : STATISTICS
'GROUP         : DIST_NON_CENTRAL
'ID            : 014
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/22/2009
'************************************************************************************
'************************************************************************************


Function INVERSE_NON_CENTRAL_TDIST_FUNC(ByVal PROBABILITY_VAL As Double, _
ByVal DEG_FREEDOM As Double, _
ByVal BETA_NC As Double, _
Optional ByVal COMP_FLAG As Boolean = True)

'PROBABILITY_VAL: is the probability associated with the prob {used by inv fns}

'DEG_FREEDOM: is the number of degrees of freedom.

'BETA_NC: is a noncentrality parameter to the distribution.

'COMP_FLAG: 1-cumulative distribution function

On Error GoTo ERROR_LABEL

    Select Case COMP_FLAG
        Case True
            INVERSE_NON_CENTRAL_TDIST_FUNC = inv_t_nc(PROBABILITY_VAL, DEG_FREEDOM, BETA_NC)
        Case False
            INVERSE_NON_CENTRAL_TDIST_FUNC = comp_inv_t_nc(PROBABILITY_VAL, DEG_FREEDOM, BETA_NC)
    End Select

Exit Function
ERROR_LABEL:
INVERSE_NON_CENTRAL_TDIST_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : CDF_NON_CENTRAL_TDIST_FUNC
'DESCRIPTION   : Find non-centrality parameter with corresponding cdf value on a
'T DIST
'LIBRARY       : STATISTICS
'GROUP         : DIST_NON_CENTRAL
'ID            : 015
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/22/2009
'************************************************************************************
'************************************************************************************

Function CDF_NON_CENTRAL_TDIST_FUNC(ByVal PROBABILITY_VAL As Double, _
ByVal X_VAL As Double, _
ByVal DEG_FREEDOM As Double, _
Optional ByVal COMP_FLAG As Boolean = True)

'PROBABILITY_VAL: is the probability associated with the prob {used by ncp fns}

'X_VAL: is the value at which you want to evaluate the distribution.

'DEG_FREEDOM: is the number of degrees of freedom.

'COMP_FLAG: 1-cumulative distribution function

On Error GoTo ERROR_LABEL

    Select Case COMP_FLAG
        Case True 'Find non-centrality parameter with corresponding cdf value
            CDF_NON_CENTRAL_TDIST_FUNC = ncp_t_nc(PROBABILITY_VAL, X_VAL, DEG_FREEDOM)
        Case False 'Find non-centrality parameter with corresponding comp_cdf value
            CDF_NON_CENTRAL_TDIST_FUNC = comp_ncp_t_nc(PROBABILITY_VAL, X_VAL, DEG_FREEDOM)
    End Select

Exit Function
ERROR_LABEL:
CDF_NON_CENTRAL_TDIST_FUNC = Err.number
End Function
