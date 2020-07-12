Attribute VB_Name = "STAT_DIST_SINGLE_ORDER_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.

'************************************************************************************
'************************************************************************************
'FUNCTION      : SINGLE_ORDER_PDF_NORMAL_FUNC

'DESCRIPTION   : Single order probability density function
' pdf for rth smallest (r>0) or -rth largest (r<0) of the N_VAL order
' statistics from a sample of iid N_VAL(0,1) variables
' based on formula 2.1.5 "Order Statistics" by H.A. David (any edition)

'LIBRARY       : STATISTICS
'GROUP         : DIST_SINGLE_ORDER
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/22/2009
'************************************************************************************
'************************************************************************************

Function SINGLE_ORDER_PDF_NORMAL_FUNC(ByVal X_VAL As Double, _
ByVal N_VAL As Double, _
Optional ByVal R_VAL As Double = -1)

On Error GoTo ERROR_LABEL

    SINGLE_ORDER_PDF_NORMAL_FUNC = pdf_normal_os(X_VAL, N_VAL, R_VAL)

Exit Function
ERROR_LABEL:
SINGLE_ORDER_PDF_NORMAL_FUNC = Err.number
End Function
 

'************************************************************************************
'************************************************************************************
'FUNCTION      : SINGLE_ORDER_CDF_NORMAL_FUNC

'DESCRIPTION   : Single order cumulative distribution function
' cdf for rth smallest (r>0) or -rth largest (r<0) of the N_VAL order
' statistics from a sample of iid N_VAL(0,1) variables
' based on formula 2.1.5 "Order Statistics" by H.A. David (any edition)

'LIBRARY       : STATISTICS
'GROUP         : DIST_SINGLE_ORDER
'ID            : 002
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/22/2009
'************************************************************************************
'************************************************************************************

Function SINGLE_ORDER_CDF_NORMAL_FUNC(ByVal X_VAL As Double, _
ByVal N_VAL As Double, _
Optional ByVal R_VAL As Double = -1)
    
On Error GoTo ERROR_LABEL

    SINGLE_ORDER_CDF_NORMAL_FUNC = cdf_normal_os(X_VAL, N_VAL, R_VAL)
 
Exit Function
ERROR_LABEL:
SINGLE_ORDER_CDF_NORMAL_FUNC = Err.number
End Function
 

'************************************************************************************
'************************************************************************************
'FUNCTION      : SINGLE_ORDER_COMP_CDF_NORMAL_FUNC

'DESCRIPTION   : Single order 1-cumulative distribution function
'1-cdf for rth smallest (r>0) or -rth largest (r<0) of the N_VAL order statistics
'from a sample of iid N_VAL(0,1) variables
'based on formula 2.1.5 "Order Statistics" by H.A. David (any edition)

'LIBRARY       : STATISTICS
'GROUP         : DIST_SINGLE_ORDER
'ID            : 003
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/22/2009
'************************************************************************************
'************************************************************************************

Function SINGLE_ORDER_COMP_CDF_NORMAL_FUNC(ByVal X_VAL As Double, _
ByVal N_VAL As Double, _
Optional ByVal R_VAL As Double = -1)
    
On Error GoTo ERROR_LABEL
    
    SINGLE_ORDER_COMP_CDF_NORMAL_FUNC = comp_cdf_normal_os(X_VAL, N_VAL, R_VAL)

Exit Function
ERROR_LABEL:
SINGLE_ORDER_COMP_CDF_NORMAL_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : SINGLE_ORDER_INVERSE_NORMAL_FUNC

'DESCRIPTION   : Single order inverse of cdf for continuous density functions
'based on formula 2.1.5 "Order Statistics" by H.A. David (any edition)
'accuracy for median of extreme order statistic is limited by accuracy of
'IEEE double precision representation of N_VAL >> 10^15, not by this routine

'LIBRARY       : STATISTICS
'GROUP         : DIST_SINGLE_ORDER
'ID            : 004
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/22/2009
'************************************************************************************
'************************************************************************************

Function SINGLE_ORDER_INVERSE_NORMAL_FUNC(ByVal P_VAL As Double, _
ByVal N_VAL As Double, _
Optional ByVal R_VAL As Double = -1)
    
On Error GoTo ERROR_LABEL
    
SINGLE_ORDER_INVERSE_NORMAL_FUNC = inv_normal_os(P_VAL, N_VAL, R_VAL)
 
Exit Function
ERROR_LABEL:
SINGLE_ORDER_INVERSE_NORMAL_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : SINGLE_ORDER_COMP_INVERSE_NORMAL_FUNC

'DESCRIPTION   : Single order inverse of comp_cdf for continuous density functions
'based on formula 2.1.5 "Order Statistics" by H.A. David (any edition)

'LIBRARY       : STATISTICS
'GROUP         : DIST_SINGLE_ORDER
'ID            : 005
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/22/2009
'************************************************************************************
'************************************************************************************

Function SINGLE_ORDER_COMP_INVERSE_NORMAL_FUNC(ByVal P_VAL As Double, _
ByVal N_VAL As Double, _
Optional ByVal R_VAL As Double = -1)

On Error GoTo ERROR_LABEL
    
    SINGLE_ORDER_COMP_INVERSE_NORMAL_FUNC = comp_inv_normal_os(P_VAL, N_VAL, R_VAL)

Exit Function
ERROR_LABEL:
SINGLE_ORDER_COMP_INVERSE_NORMAL_FUNC = Err.number
End Function
