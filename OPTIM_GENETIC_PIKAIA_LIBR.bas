Attribute VB_Name = "OPTIM_GENETIC_PIKAIA_LIBR"


Option Explicit     'Requires that all variables to be declared explicitly.

                            
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.

'----------------------------------------------------------------------------------
'----------------------------------------------------------------------------------
'----------------------------------------------------------------------------------

Private XDATA_ARR() As Double
'Array  x(1:n)  is the "fittest" (optimal) solution found,
'i.e., the solution which maximizes fitness function ff

Private PUB_Y_VAL As Double
'is the value of the fitness function at x

Private XTEMP_ARR() As Double
'temporary scratch array for x() to build and pass to ff

Private PUB_CONVERG_VAL As Integer
'is an indicator of the success or failure
'of the call to pikaia (0=success; non-zero=failure)

'----------------------------------------------------------------------------------
'----------------------------------------------------------------------------------
Private Const NO_PARAM As Integer = 256 'is the maximum number of
'adjustable parameters

Private Const MAX_POP As Integer = 512
'maximum population (CTRL(1) <= MAX_POP

Private Const MAX_GENES As Integer = 6
'maximum number of Genes (digits) per Chromosome
'segement(Parameter)(CTRL(3) <= DMax)
'----------------------------------------------------------------------------------
'----------------------------------------------------------------------------------

'for sub PIKAIA_REPORT_FUNC
Private PUB_BEST_FIT_VAL As Double
Private PUB_PMUTPV_VAL As Double
Private PUB_SEED_VAL As Long 'for random number generator
'----------------------------------------------------------------------------------
'----------------------------------------------------------------------------------

Private PARAM_ARR() As Double
Private LOWER_ARR() As Double
Private UPPER_ARR() As Double

Private PUB_REPORT_GROUP As Variant
Private PUB_FUNC_NAME_STR As String
'----------------------------------------------------------------------------------
'----------------------------------------------------------------------------------


'************************************************************************************
'************************************************************************************
'FUNCTION      : PIKAIA_OPTIMIZATION_FUNC

'DESCRIPTION   : fully self-contained, general purpose optimization subroutine.
'The routine maximizes a user-supplied FORTRAN function, the name of which
'is passed in as an argument.

'PIKAIA (pronounced ``pee-kah-yah'') is a general purpose function optimization
'FORTRAN-77 subroutine based on a genetic algorithm. The subroutine is particularly
'useful (and robust) in treating multimodal optimization problems. The development
'of genetic algorithm-based inversion methods is but one aspect of the research
'in helioseismology carried out in the Solar Interior Section of the High Altitude
'Observatory, a scientific division of the National Center for Atmospheric Research
'in Boulder, Colorado.


'     PIKAIA is an optimization (maximization) of user-supplied "fitness" function
'     over n-dimensional parameter space  x  using a basic genetic algorithm method.
'
'     Genetic algorithms are heuristic search techniques that
'     incorporate in a computational setting, the biological notion
'     of evolution by means of natural selection.  This subroutine
'     implements the three basic operations of selection, crossover,
'     and mutation, operating on "genotypes" encoded as strings.
'
'     This version of the PIKAIA algorithm includes (1) two-point crossover,
'     (2) creep mutation, and (3) dynamical adjustment of the mutation rate
'     based on metric distance in parameter space.

'
'      o Integer  n  is the parameter space dimension, i.e., the number
'        of adjustable parameters.
'
'      o Function  ff  is a user-supplied scalar function of n vari-
'        ables, which must have the calling sequence f = ff(n,x), where
'        x is a real parameter array of length n.  This function must
'        be written so as to bound all parameters to the interval [0,1];
'        that is, the user must determine a priori bounds for the para-
'        meter space, and ff must use these bounds to perform the appro-
'        priate scalings to recover true parameter values in the
'        a priori ranges.
'
'        By convention, ff should return higher values for more optimal
'        parameter values (i.e., individuals which are more "fit").
'        For example, in fitting a function through data points, ff
'        could return the inverse of chi**2.
'
'        In most cases initialization code will have to be written
'        (either in a driver or in a separate subroutine) which loads
'        in data values and communicates with ff via one or more labeled
'        common blocks.  An example exercise driver and fitness function
'        are provided in the accompanying file, xpkaia.f.
'
'
'      Input/Output:
'       real CTRL(12)
'
'      o Array  CTRL  is an array of control flags and parameters, to
'        control the genetic behavior of the algorithm, and also printed
'        output.  A default value will be used for any control variable
'        which is supplied with a value less than zero.  On exit, CTRL
'        contains the actual values used as control variables.  The
'        elements of CTRL and their defaults are:
'
'           CTRL( 1) - number of individuals in a population (default
'                      is 100)
'           CTRL( 2) - number of generations over which solution is
'                      to evolve (default is 500)
'           CTRL( 3) - number of significant digits (i.e., number of
'                      genes) retained in chromosomal encoding (default
'                      is 6)  (Note: This number is limited by the
'                      machine floating point precision.  Most 32-bit
'                      floating point representations have only 6 full
'                      digits of precision.  To achieve greater preci-
'                      sion this routine could be converted to double
'                      precision, but note that this would also require
'                      a double precision random number generator, which
'                      likely would not have more than 9 digits of
'                      precision if it used 4-byte integers internally.)
'           CTRL( 4) - crossover probability; must be  <= 1.0 (default
'                      is 0.85). If crossover takes place, either one
'                      or two splicing points are used, with equal
'                      probabilities
'           CTRL( 5) - mutation mode; 1/2/3/4/5 (default is 2)
'                      1=one-point mutation, fixed rate
'                      2=one-point, adjustable rate based on fitness
'                      3=one-point, adjustable rate based on distance
'                      4=one-point+creep, fixed rate
'                      5=one-point+creep, adjustable rate based on fitness
'                      6=one-point+creep, adjustable rate based on distance
'           CTRL( 6) - initial mutation rate; should be small (default
'                      is 0.005) (Note: the mutation rate is the proba-
'                      bility that any one gene locus will mutate in
'                      any one generation.)
'           CTRL( 7) - minimum mutation rate; must be >= 0.0 (default
'                      is 0.0005)
'           CTRL( 8) - maximum mutation rate; must be <= 1.0 (default
'                      is 0.25)
'           CTRL( 9) - relative fitness differential; range from 0
'                      (none) to 1 (maximum).  (default is 1.)
'           CTRL(10) - reproduction plan; 1/2/3=Full generational
'                      replacement/Steady-state-Replace-random/Steady-
'                      State - Replace - worst(Default Is 3)
'           CTRL(11) - elitism flag; 0/1=off/on (default is 0)
'                      (Applies only to reproduction plans 1 and 2)
'           CTRL(12) - printed output 0/1/2=None/Minimal/Verbose
'                      (Default Is 0)
'
'
' Output:
'      real x(n), f
'      integer   status
'
'      o Array  x(1:n)  is the "fittest" (optimal) solution found,
'         i.e., the solution which maximizes fitness function ff
'
'      o Scalar  f  is the value of the fitness function at x
'
'      o Integer  status  is an indicator of the success or failure
'         of the call to pikaia (0=success; non-zero=failure)

' References:
'        Charbonneau, Paul. "An introduction to gemetic algorithms for
'           numerical optimization", NCAR Technical Note TN-450+IA
'           (April 2002)
'        Charbonneau, Paul. "Release Notes for PIKAIA 1.2",
'           NCAR Technical Note TN-451+STR (April 2002)
'        Charbonneau, Paul, and Knapp, Barry. "A User's Guide
'           to PIKAIA 1.0" NCAR Technical Note TN-418+IA
'           (December 1995)
'        Goldberg, David E.  Genetic Algorithms in Search, Optimization,
'           & Machine Learning.  Addison-Wesley, 1989.
'        Davis, Lawrence, ed.  Handbook of Genetic Algorithms.
'           Van Nostrand Reinhold, 1991.

'LIBRARY       : OPTIMIZATION
'GROUP         : GENETIC_PIKAIA
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 02/13/2009
'************************************************************************************
'************************************************************************************

Function PIKAIA_OPTIMIZATION_FUNC(ByVal FUNC_NAME_STR As String, _
ByVal CONST_RNG As Variant, _
Optional ByVal TRACE_FLAG As Boolean = False, _
Optional ByRef ERROR_STR As String = "", _
Optional ByVal RND_NUMB_SEED As Long = 123456, _
Optional ByVal NUMB_INDIV_POPUP As Integer = 100, _
Optional ByVal NUMB_GENER_EVOL As Integer = 50, _
Optional ByVal NUMB_DIGITS_ENCODE As Integer = 5, _
Optional ByVal CROSS_PROBAB As Double = 0.85, _
Optional ByVal MUTAT_MODE As Integer = 2, _
Optional ByVal INIT_MUTAT_RATE As Double = 0.005, _
Optional ByVal MIN_MUTAT_RATE As Double = 0.0005, _
Optional ByVal MAX_MUTAT_RATE As Double = 0.25, _
Optional ByVal RELAT_FITNESS As Integer = 1, _
Optional ByVal REPROD_PLAN As Integer = 1, _
Optional ByVal ELITISM As Integer = 1, _
Optional ByVal OUTPUT As Integer = 0)

PUB_FUNC_NAME_STR = FUNC_NAME_STR

Dim i As Long
Dim j As Integer
Dim NSIZE As Integer
'the number of parameters in the fitness function
Dim PARAM_VECTOR As Variant
Dim TRACE_MATRIX As Variant
Dim CONST_BOX As Variant

On Error GoTo ERROR_LABEL

Dim CTRL_ARR(1 To 12) As Variant
'is an input/output array of control flags and parameters, to
'control the genetic behavior of the algorithm, and also printed
'output.  A default value will be used for any control variable
'which is supplied with a value less than zero.  On exit, CTRL
'contains the actual values used as control variables.

ERROR_STR = ""
CTRL_ARR(1) = NUMB_INDIV_POPUP
'number of individuals in a population (default is 100)
CTRL_ARR(2) = NUMB_GENER_EVOL
'number of generations over which solution is to evolve (default is 500)
CTRL_ARR(3) = NUMB_DIGITS_ENCODE
'number of significant digits (i.e., number of genes) retained in
'chromosomal encoding (default is 6)  (Note: This number is limited by the
'machine floating point precision.  Most 32-bit floating point representations
'have only 6 full digits of precision.  To achieve greater preci-
'sion this routine could be converted to double precision, but note that
'this would also require a double precision random number generator, which
'likely would not have more than 9 digits of precision if it used 4-byte
'integers internally.)
CTRL_ARR(4) = CROSS_PROBAB
'crossover probability; must be  <= 1.0 (default is 0.85). If crossover takes
'place, either one or two splicing points are used, with equal probabilities
CTRL_ARR(5) = MUTAT_MODE
'mutation mode; 1/2/3/4/5 (default is 2)
'1=one-point mutation, fixed rate
'2=one-point, adjustable rate based on fitness
'3=one-point, adjustable rate based on distance
'4=one-point+creep, fixed rate
'5=one-point+creep, adjustable rate based on fitness
'6=one-point+creep, adjustable rate based on distance
CTRL_ARR(6) = INIT_MUTAT_RATE
'initial mutation rate; should be small (default
'is 0.005) (Note: the mutation rate is the proba-
'bility that any one gene locus will mutate in
'any one generation.)
CTRL_ARR(7) = MIN_MUTAT_RATE
'minimum mutation rate; must be >= 0.0 (default is 0.0005)
CTRL_ARR(8) = MAX_MUTAT_RATE
'maximum mutation rate; must be <= 1.0 (default is 0.25)
CTRL_ARR(9) = RELAT_FITNESS
'relative fitness differential; range from 0
'(none) to 1 (maximum).  (default is 1.)
CTRL_ARR(10) = REPROD_PLAN
'reproduction plan; 1/2/3=Full generational
'replacement/Steady-state-Replace-random/Steady-
'State - Replace - worst(Default Is 3)
CTRL_ARR(11) = ELITISM
'elitism flag; 0/1=off/on (default is 0)
'(Applies only to reproduction plans 1 and 2)
CTRL_ARR(12) = OUTPUT
'printed output 0/1/2=None/Minimal/Verbose

For i = 1 To 11
    Select Case i
    Case 1
        If CTRL_ARR(i) > MAX_POP Or CTRL_ARR(i) < 2 Then
            ERROR_STR = "This population size must be between 2 and " & MAX_POP
            GoTo ERROR_LABEL
        End If
    Case 2
        If CTRL_ARR(i) < 1 Then
            ERROR_STR = "There must be at least one generation."
            GoTo ERROR_LABEL
        End If
    Case 3
        If CTRL_ARR(i) > MAX_GENES Or CTRL_ARR(i) < 1 Then
            ERROR_STR = "This number of digits for encoding must be between 1 and " & _
            MAX_GENES
            GoTo ERROR_LABEL
        End If
    Case 4
        If CTRL_ARR(i) < 0 Or CTRL_ARR(i) > 1 Then
            ERROR_STR = "This crossover probability must be between 0 and 1."
            GoTo ERROR_LABEL
        End If
    Case 5
        If CTRL_ARR(i) <> 1 And CTRL_ARR(i) <> 2 And CTRL_ARR(i) <> 3 And _
            CTRL_ARR(i) <> 4 And CTRL_ARR(i) <> 5 And CTRL_ARR(i) <> 6 Then
            ERROR_STR = "This mutation mode must be 1, 2, 3, 4, 5, or 6."
            GoTo ERROR_LABEL
        End If
    Case 6
    If CTRL_ARR(i) < 0 Or CTRL_ARR(i) > 1 Then
        ERROR_STR = "This initial mutation rate must be between 0 and 1."
        GoTo ERROR_LABEL
    End If
    Case 7
        If CTRL_ARR(i) < 0 Or CTRL_ARR(i) > 1 Then
        ERROR_STR = "This minimum mutation rate must be between 0 and 1."
        GoTo ERROR_LABEL
    End If
    Case 8
        If CTRL_ARR(i) < 0 Or CTRL_ARR(i) > 1 Then
        ERROR_STR = "This maximum mutation rate must be between 0 and 1."
        GoTo ERROR_LABEL
    End If
    Case 9
        If CTRL_ARR(i) < 0 Or CTRL_ARR(i) > 1 Then
        ERROR_STR = "This relative fitness differential must be between 0 and 1."
        GoTo ERROR_LABEL
    End If
    Case 10
        If CTRL_ARR(i) <> 1 And CTRL_ARR(i) <> 2 And CTRL_ARR(i) <> 3 Then
        ERROR_STR = "This reproduction plan must be 1, 2, or 3."
        GoTo ERROR_LABEL
    End If
    Case 11
        If CTRL_ARR(i) <> 0 And CTRL_ARR(i) <> 1 Then
        ERROR_STR = "Elitism must be 0 or 1."
        GoTo ERROR_LABEL
    End If
    End Select
Next i

CONST_BOX = CONST_RNG
NSIZE = UBound(CONST_BOX, 2)

If TRACE_FLAG = True Then
    ReDim TRACE_MATRIX(0 To CTRL_ARR(1) * (CTRL_ARR(2) + 1), 1 To NSIZE + 6)
End If

ReDim XDATA_ARR(1 To NSIZE) As Double
ReDim XTEMP_ARR(1 To NSIZE) As Double
ReDim LOWER_ARR(1 To NSIZE) As Double
ReDim UPPER_ARR(1 To NSIZE) As Double
ReDim PARAM_ARR(1 To NSIZE) As Double


For i = 1 To NSIZE
    LOWER_ARR(i) = CONST_BOX(1, i)
    UPPER_ARR(i) = CONST_BOX(2, i)
Next i

Call INIT_RND_GENER_FUNC(RND_NUMB_SEED)
ERROR_STR = ""
PUB_CONVERG_VAL = GENETIC_PIKAIA_OPTIM_FUNC(NSIZE, CTRL_ARR, _
                  TRACE_MATRIX, ERROR_STR)

If ERROR_STR <> "" Then: GoTo ERROR_LABEL

'------------------------------------------------------------------------------------
If OUTPUT <> 0 Then
    PIKAIA_OPTIMIZATION_FUNC = PUB_REPORT_GROUP
    Exit Function
End If

'------------------------------------------------------------------------------------
If TRACE_FLAG = True Then
'------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------
'These entries shows the value of 1-fitness during the evolution. This is a
'measure of the error (smaller values of 1-ff means less error in this example).
'For a different problem this chart can be changed to show either the fitness
'value or 1/fitness of any other measure of fitness.
'------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------
    TRACE_MATRIX(0, NSIZE + 5) = "1/Fitness"
    TRACE_MATRIX(0, NSIZE + 6) = "1-Fitness"
    For i = 1 To CTRL_ARR(1) * (CTRL_ARR(2) + 1) 'optimum value
        TRACE_MATRIX(i, NSIZE + 5) = 1 / TRACE_MATRIX(i, NSIZE + 4) '1 / Fitness
        TRACE_MATRIX(i, NSIZE + 6) = 1 - TRACE_MATRIX(i, NSIZE + 4) '1 - Fitness
    Next i
    PIKAIA_OPTIMIZATION_FUNC = TRACE_MATRIX
'------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------
Else
'------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------
    ReDim PARAM_VECTOR(1 To NSIZE, 1 To 1)
    'PARAM_VECTOR(0, 1) = PUB_Y_VAL
    For j = 1 To NSIZE
        PARAM_VECTOR(j, 1) = PIKAI_PARAM_SCALE_FUNC(j, XDATA_ARR(j))
    Next j
    PIKAIA_OPTIMIZATION_FUNC = PARAM_VECTOR
'------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------
End If
'------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------

Exit Function
ERROR_LABEL:
PIKAIA_OPTIMIZATION_FUNC = PUB_CONVERG_VAL
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : GENETIC_PIKAIA_OPTIM_FUNC

'DESCRIPTION   :

'PIKAIA incorporates only the two basic genetic operators: uniform one-point
'crossover, and uniform one-point mutation. Unlike many GA packages available
'commercially or in the public domain, the encoding within PIKAIA is based on
'a decimal alphabet made of the 10 simple integers (0 through 9); this is because
'binary operations are usually carried out through platform-dependent functions
'in FORTRAN. Three reproduction plans are available: Full generational replacement,
'Steady-State-Delete-Random, and Steady-State-Delete-Worst. Elitism
'is available and is a default option. The mutation rate can be dynamically
'controlled by monitoring the difference in fitness between the current best and
'median in the population (also a default option). Selection is rank-based and
'stochastic, making use of the Roulette

'Wheel Algorithm.
'PIKAIA is supplied with a ranking subroutine based on the Quicksort algorithm,
'and a random number generator based on the minimal standard Lehmer multiplicative
'linear congruential generator of Park and Miller (1988, Communications of the
'ACM, 31, 1192).

'LIBRARY       : OPTIMIZATION
'GROUP         : GENETIC_PIKAIA
'ID            : 002
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 02/13/2009
'************************************************************************************
'************************************************************************************

Private Function GENETIC_PIKAIA_OPTIM_FUNC(ByVal NSIZE As Integer, _
ByRef CTRL_ARR As Variant, _
Optional ByRef TRACE_MATRIX As Variant, _
Optional ByRef ERROR_STR As String = "")

Dim i As Long
Dim j As Integer
Dim k As Integer

Dim ND_INT As Integer
Dim NPOP_INT As Integer
Dim NGEN_INT As Integer

Dim IMUT_INT As Integer
Dim IREP_INT As Integer
Dim IELITE_INT As Integer

Dim IP1_INT As Integer
Dim IP2_INT As Integer

Dim IVRB_INT As Integer
Dim IPOP_INT As Integer
Dim IGEN_INT As Integer

Dim NEWW_INT As Integer
Dim NEWTOT_INT As Integer

Dim FDIF_VAL As Double
Dim PMUT_VAL As Double
Dim PMUTMN_VAL As Double
Dim PMUTMX_VAL As Double
Dim PCROSS_VAL As Double

On Error GoTo ERROR_LABEL

Dim IFIT_ARR(1 To MAX_POP) As Integer
Dim JFIT_ARR(1 To MAX_POP) As Integer

Dim FITNS_ARR(1 To MAX_POP) As Double
Dim PH_MAT(1 To NO_PARAM, 1 To 2) As Double

Dim OLDPH_MAT(1 To NO_PARAM, 1 To MAX_POP) As Double
Dim NEWPH_MAT(1 To NO_PARAM, 1 To MAX_POP) As Double

Dim GN1_ARR(1 To NO_PARAM * MAX_GENES) As Integer
Dim GN2_ARR(1 To NO_PARAM * MAX_GENES) As Integer

'Set control variables from input and defaults

PUB_CONVERG_VAL = PIKAIA_CTRL_FUNC(CTRL_ARR, NSIZE, NPOP_INT, NGEN_INT, _
            ND_INT, PCROSS_VAL, PMUTMN_VAL, PMUTMX_VAL, _
            PMUT_VAL, IMUT_INT, FDIF_VAL, IREP_INT, IELITE_INT, _
            IVRB_INT, ERROR_STR)

If (PUB_CONVERG_VAL <> 0) Then 'Program stopped because control vector (CTRL)
'argument(s) invalid
    GoTo ERROR_LABEL
End If

'Make sure locally-dimensioned arrays are big enough

If (NSIZE > NO_PARAM Or NPOP_INT > MAX_POP Or ND_INT > MAX_GENES) Then
    ERROR_STR = "Program stopped because number of parameters, " & _
    "population, or genes too large"
    PUB_CONVERG_VAL = -1
    GoTo ERROR_LABEL
End If


If IsArray(TRACE_MATRIX) = True Then
'gp set the header row and pointer for output of the populations
'during each generation
    TRACE_MATRIX(0, 1) = "Generation"
    TRACE_MATRIX(0, 2) = "Rank"
    TRACE_MATRIX(0, 3) = "Quantile"
    For k = 1 To NSIZE
        TRACE_MATRIX(0, 3 + k) = "X(" & k & ")"
    Next k
    TRACE_MATRIX(0, 3 + NSIZE + 1) = "Fitness"
End If

'Compute initial (random but bounded) phenotypes

For IPOP_INT = 1 To NPOP_INT
    For k = 1 To NSIZE
        OLDPH_MAT(k, IPOP_INT) = NEXT_PSEUDO_RND_FUNC()
        XTEMP_ARR(k) = OLDPH_MAT(k, IPOP_INT)
    Next k
    FITNS_ARR(IPOP_INT) = PIKAIA_OBJ_FUNC(NSIZE, XTEMP_ARR())
Next IPOP_INT


'Rank initial population by fitness order
Call PIKAIA_RANK_FUNC(NPOP_INT, FITNS_ARR, IFIT_ARR, JFIT_ARR)

If IsArray(TRACE_MATRIX) = True Then
    'output of ranked populations for each generation
    i = 1
    For j = NPOP_INT To 1 Step -1
        TRACE_MATRIX(i, 1) = 0 'generation number 0 for initial population
        TRACE_MATRIX(i, 2) = j 'rank
        TRACE_MATRIX(i, 3) = j / (NPOP_INT + 1) 'quantile
        
        For k = 1 To NSIZE
            TRACE_MATRIX(i, 3 + k) = OLDPH_MAT(k, IFIT_ARR(j))
            'ranked phenomes of this individual
        Next k
        TRACE_MATRIX(i, 3 + NSIZE + 1) = FITNS_ARR(IFIT_ARR(j))
        'fitness of that individual
        i = i + 1
    Next j
End If

'main generation loop
For IGEN_INT = 1 To NGEN_INT
    NEWTOT_INT = 0
    For IPOP_INT = 1 To NPOP_INT / 2
        'pick two parents
        Call PIKAIA_PARENT_FUNC(NPOP_INT, JFIT_ARR, FDIF_VAL, IP1_INT) 'pick dad
21:
        Call PIKAIA_PARENT_FUNC(NPOP_INT, JFIT_ARR, FDIF_VAL, IP2_INT) 'pick mom
        If (IP1_INT = IP2_INT) Then GoTo 21 'no breeding with yourself!
        
        'encode parent phenotypes
        For k = 1 To NSIZE
            XTEMP_ARR(k) = OLDPH_MAT(k, IP1_INT)
        Next k
        
        Call PIKAIA_ENCODE_FUNC(NSIZE, ND_INT, XTEMP_ARR, GN1_ARR)
        For k = 1 To NSIZE
            XTEMP_ARR(k) = OLDPH_MAT(k, IP2_INT)
        Next k
        
        Call PIKAIA_ENCODE_FUNC(NSIZE, ND_INT, XTEMP_ARR, GN2_ARR)
        
        'breed
        Call PIKAIA_CROSS_FUNC(NSIZE, ND_INT, PCROSS_VAL, GN1_ARR, GN2_ARR)
        Call PIKAIA_MUTATE_FUNC(NSIZE, ND_INT, PMUT_VAL, GN1_ARR, IMUT_INT)
        Call PIKAIA_MUTATE_FUNC(NSIZE, ND_INT, PMUT_VAL, GN2_ARR, IMUT_INT)
        Call PIKAIA_DECODE_FUNC(NSIZE, ND_INT, GN1_ARR, XTEMP_ARR)

        For k = 1 To NSIZE
            PH_MAT(k, 1) = XTEMP_ARR(k)
        Next k

        'decode offspring genotypes
        Call PIKAIA_DECODE_FUNC(NSIZE, ND_INT, GN2_ARR, XTEMP_ARR)
        For k = 1 To NSIZE
            PH_MAT(k, 2) = XTEMP_ARR(k)
        Next k
        
        'insert into population
        If (IREP_INT = 1) Then
            Call PIKAIA_GENREP_FUNC(NO_PARAM, NSIZE, NPOP_INT, _
                 IPOP_INT, PH_MAT, NEWPH_MAT)
        Else
            Call PIKAIA_STDREP_FUNC(NO_PARAM, NSIZE, NPOP_INT, _
                IREP_INT, IELITE_INT, PH_MAT, _
                OLDPH_MAT, FITNS_ARR, IFIT_ARR, JFIT_ARR, NEWW_INT)
                NEWTOT_INT = NEWTOT_INT + NEWW_INT
        End If
    Next IPOP_INT

    'if running full generational replacement: swap populations
    If (IREP_INT = 1) Then
        Call PIKAIA_NEWPOP_FUNC(IELITE_INT, NO_PARAM, NSIZE, _
            NPOP_INT, OLDPH_MAT, NEWPH_MAT, _
            IFIT_ARR, JFIT_ARR, FITNS_ARR, NEWTOT_INT)
    End If
    
    If (IMUT_INT = 2 Or IMUT_INT = 3 Or IMUT_INT = 5 Or IMUT_INT = 6) Then
    'adjust mutation rate?
        PUB_CONVERG_VAL = VARIAB_MUTAT_RATE_FUNC(NO_PARAM, NSIZE, _
                      NPOP_INT, OLDPH_MAT, FITNS_ARR, _
                      IFIT_ARR, PMUTMN_VAL, PMUTMX_VAL, _
                      PMUT_VAL, IMUT_INT, ERROR_STR)
        If PUB_CONVERG_VAL <> 0 Then: GoTo ERROR_LABEL
    End If
    
    
    If (IVRB_INT > 0) Then
        If IGEN_INT = 1 Then: ReDim PUB_REPORT_GROUP(1 To NGEN_INT)
        PUB_REPORT_GROUP(IGEN_INT) = PIKAIA_REPORT_FUNC(IVRB_INT, NO_PARAM, _
             NSIZE, NPOP_INT, ND_INT, OLDPH_MAT, FITNS_ARR, _
             IFIT_ARR, PMUT_VAL, IGEN_INT, NEWTOT_INT)
    End If
    
    If IsArray(TRACE_MATRIX) = True Then
  'output of ranked populations for each generation
        For j = NPOP_INT To 1 Step -1
            TRACE_MATRIX(i, 1) = IGEN_INT 'generation number
            TRACE_MATRIX(i, 2) = j 'rank
            TRACE_MATRIX(i, 3) = j / (NPOP_INT + 1) 'quantile
            For k = 1 To NSIZE
                TRACE_MATRIX(i, 3 + k) = OLDPH_MAT(k, IFIT_ARR(j))
                'ranked phenomes of this individual
            Next k
            TRACE_MATRIX(i, 3 + NSIZE + 1) = FITNS_ARR(IFIT_ARR(j))
            'fitness of that individual
            i = i + 1
        Next j
    End If
'End of Main Generation Loop
Next IGEN_INT

'Return best phenotype and its fitness
For k = 1 To NSIZE
    XDATA_ARR(k) = OLDPH_MAT(k, IFIT_ARR(NPOP_INT))
Next k
PUB_Y_VAL = FITNS_ARR(IFIT_ARR(NPOP_INT))

GENETIC_PIKAIA_OPTIM_FUNC = PUB_CONVERG_VAL

Exit Function
ERROR_LABEL:
GENETIC_PIKAIA_OPTIM_FUNC = PUB_CONVERG_VAL
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : PIKAIA_CTRL_FUNC
'DESCRIPTION   : Set control variables and flags from input and defaults
'LIBRARY       : OPTIMIZATION
'GROUP         : GENETIC_PIKAIA
'ID            : 003
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 02/13/2009
'************************************************************************************
'************************************************************************************

Private Function PIKAIA_CTRL_FUNC(ByRef CTRL_ARR As Variant, _
ByRef NSIZE As Integer, _
ByRef NPOP_INT As Integer, _
ByRef NGEN_INT As Integer, _
ByRef ND_INT As Integer, _
ByRef PCROSS_VAL As Double, _
ByRef PMUTMN_VAL As Double, _
ByRef PMUTMX_VAL As Double, _
ByRef PMUT_VAL As Double, _
ByRef IMUT_INT As Integer, _
ByRef FDIF_VAL As Double, _
ByRef IREP_INT As Integer, _
ByRef IELITE_INT As Integer, _
ByRef IVRB_INT As Integer, _
ByRef ERROR_STR As String)

Dim i As Integer

On Error GoTo ERROR_LABEL

Dim DEFAULT_ARR(1 To 12) As Variant
'for sub setctl to set defaults for CTRL array

ERROR_STR = ""
DEFAULT_ARR(1) = 100
DEFAULT_ARR(2) = 500
DEFAULT_ARR(3) = 5
DEFAULT_ARR(4) = 0.85
DEFAULT_ARR(5) = 2
DEFAULT_ARR(6) = 0.005
DEFAULT_ARR(7) = 0.0005
DEFAULT_ARR(8) = 0.25
DEFAULT_ARR(9) = 1
DEFAULT_ARR(10) = 1
DEFAULT_ARR(11) = 1
DEFAULT_ARR(12) = 0
'IVRB_INT for PIKAIA Genetic Algorithm Report: 0=off, 1=on (normal default is
'0 for off)

For i = 1 To 12
    If (CTRL_ARR(i) < 0) Then CTRL_ARR(i) = DEFAULT_ARR(i)
Next i

NPOP_INT = CTRL_ARR(1)
NGEN_INT = CTRL_ARR(2)
ND_INT = CTRL_ARR(3)
PCROSS_VAL = CTRL_ARR(4)
IMUT_INT = CTRL_ARR(5)
PMUT_VAL = CTRL_ARR(6)
PMUTMN_VAL = CTRL_ARR(7)
PMUTMX_VAL = CTRL_ARR(8)
FDIF_VAL = CTRL_ARR(9)
IREP_INT = CTRL_ARR(10)
IELITE_INT = CTRL_ARR(11)
IVRB_INT = CTRL_ARR(12)
PUB_CONVERG_VAL = 0

'Check some control values

If (IMUT_INT <> 1 And IMUT_INT <> 2 And IMUT_INT <> 3 And IMUT_INT <> 4 And _
    IMUT_INT <> 5 And IMUT_INT <> 6) Then
    ERROR_STR = "ERROR: illegal value for IMUT_INT (CTRL(5))"
    PUB_CONVERG_VAL = 5
End If

If (FDIF_VAL > 1) Then
    ERROR_STR = "ERROR: illegal value for FDIF_VAL (CTRL(9))"
    PUB_CONVERG_VAL = 9
End If

If (IREP_INT <> 1 And IREP_INT <> 2 And IREP_INT <> 3) Then
    ERROR_STR = "ERROR: illegal value for IREP_INT (CTRL(10))"
    PUB_CONVERG_VAL = 10
End If

If (PCROSS_VAL > 1# Or PCROSS_VAL < 0) Then
    ERROR_STR = "ERROR: illegal value for PCROSS_VAL (CTRL(4))"
    PUB_CONVERG_VAL = 4
End If

If (IELITE_INT <> 0 And IELITE_INT <> 1) Then
    ERROR_STR = "ERROR: illegal value for IELITE_INT (CTRL(11))"
    PUB_CONVERG_VAL = 11
End If

If (IREP_INT = 1 And IMUT_INT = 1 And PMUT_VAL > 0.5 And IELITE_INT = 0) Then
    ERROR_STR = _
    "WARNING: dangerously high value for PMUT_VAL (CTRL(6)). " & _
    "Should enforce elitism with CTRL(11)=1"
End If

If (IREP_INT = 1 And IMUT_INT = 2 And PMUTMX_VAL > 0.5 And IELITE_INT = 0) Then
    ERROR_STR = _
    "WARNING: dangerously high value for PMUTMX_VAL (CTRL(8)). " & _
    "Should enforce elitism with CTRL(11)=1"
End If

If (FDIF_VAL < 0.33 And IREP_INT <> 3) Then
    ERROR_STR = _
    "WARNING: dangerously low value of FDIF_VAL (CTRL(9))"
End If

If NPOP_INT Mod 2 > 0 Then
    NPOP_INT = NPOP_INT - 1
    ERROR_STR = _
    "WARNING: decreasing population size (CTRL(1)) to npop=" & NPOP_INT
End If

PIKAIA_CTRL_FUNC = PUB_CONVERG_VAL

Exit Function
ERROR_LABEL:
PIKAIA_CTRL_FUNC = PUB_CONVERG_VAL
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : VARIAB_MUTAT_RATE_FUNC
'DESCRIPTION   : Implements variable mutation rate
'LIBRARY       : OPTIMIZATION
'GROUP         : GENETIC_PIKAIA
'ID            : 004
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 02/13/2009
'************************************************************************************
'************************************************************************************

Private Function VARIAB_MUTAT_RATE_FUNC(ByRef NO_PARAM As Integer, _
ByRef NSIZE As Integer, _
ByRef NPOP_INT As Integer, _
ByRef OLDPH_MAT As Variant, _
ByRef FITNS_ARR As Variant, _
ByRef IFIT_ARR As Variant, _
ByRef PMUTMN_VAL As Double, _
ByRef PMUTMX_VAL As Double, _
ByRef PMUT_VAL As Double, _
ByRef IMUT_INT As Integer, _
ByRef ERROR_STR As String)

'dynamical adjustment of mutation rate;
'IMUT_INT=2 or IMUT_INT=5 : adjustment based on fitness differential
'between best and median individuals
'IMUT_INT=3 or IMUT_INT=6 : adjustment based on metric distance
'between best and median individuals

Dim i As Integer
Dim RDIF_VAL As Double

On Error GoTo ERROR_LABEL

Const RDIFLO_VAL As Double = 0.05
Const RDIFHI_VAL As Double = 0.25
Const DELTA_VAL As Double = 1.5

ERROR_STR = ""
PUB_CONVERG_VAL = 0
' Adjustment based on fitness differential
If (IMUT_INT = 2 Or IMUT_INT = 5) Then
    If FITNS_ARR(IFIT_ARR(NPOP_INT)) + FITNS_ARR(IFIT_ARR(NPOP_INT / 2)) = 0 Then
        ERROR_STR = "Invalid fitness function"
        PUB_CONVERG_VAL = -1
        GoTo ERROR_LABEL
    End If
    RDIF_VAL = Abs(FITNS_ARR(IFIT_ARR(NPOP_INT)) - _
            FITNS_ARR(IFIT_ARR(NPOP_INT / 2))) / _
          (FITNS_ARR(IFIT_ARR(NPOP_INT)) + FITNS_ARR(IFIT_ARR(NPOP_INT / 2)))
ElseIf (IMUT_INT = 3 Or IMUT_INT = 6) Then
    RDIF_VAL = 0
    For i = 1 To NSIZE
        RDIF_VAL = RDIF_VAL + (OLDPH_MAT(i, IFIT_ARR(NPOP_INT)) - _
               OLDPH_MAT(i, IFIT_ARR(NPOP_INT / 2))) ^ 2
    Next i
    RDIF_VAL = Sqr(RDIF_VAL) / NSIZE
End If

'Adjustment based on normalized metric distance
If (RDIF_VAL <= RDIFLO_VAL) Then
    PMUT_VAL = MINIMUM_FUNC(PMUTMX_VAL, PMUT_VAL * DELTA_VAL)
ElseIf (RDIF_VAL >= RDIFHI_VAL) Then
    PMUT_VAL = MAXIMUM_FUNC(PMUTMN_VAL, PMUT_VAL / DELTA_VAL)
End If

VARIAB_MUTAT_RATE_FUNC = PUB_CONVERG_VAL

Exit Function
ERROR_LABEL:
VARIAB_MUTAT_RATE_FUNC = PUB_CONVERG_VAL
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : PIKAIA_REPORT_FUNC
'DESCRIPTION   : Write generation report to standard output
'LIBRARY       : OPTIMIZATION
'GROUP         : GENETIC_PIKAIA
'ID            : 005
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 02/13/2009
'************************************************************************************
'************************************************************************************

Private Function PIKAIA_REPORT_FUNC(ByRef IVRB_INT As Integer, _
ByRef NO_PARAM As Integer, _
ByRef NSIZE As Integer, _
ByRef NPOP_INT As Integer, _
ByRef ND_INT As Integer, _
ByRef OLDPH_MAT As Variant, _
ByRef FITNS_ARR As Variant, _
ByRef IFIT_ARR As Variant, _
ByRef PMUT_VAL As Double, _
ByRef IGEN_INT As Integer, _
ByRef NNEW_INT As Integer)

Dim k As Long
Dim NDPWR_INT As Long
Dim RPT_FLAG As Boolean

On Error GoTo ERROR_LABEL

PUB_BEST_FIT_VAL = 0
PUB_PMUTPV_VAL = 0
RPT_FLAG = False

If (PMUT_VAL <> PUB_PMUTPV_VAL) Then
    PUB_PMUTPV_VAL = PMUT_VAL
    RPT_FLAG = True
End If

If (FITNS_ARR(IFIT_ARR(NPOP_INT)) <> PUB_BEST_FIT_VAL) Then
    PUB_BEST_FIT_VAL = FITNS_ARR(IFIT_ARR(NPOP_INT))
    RPT_FLAG = True
End If

If (RPT_FLAG Or IVRB_INT >= 2) Then 'Power of 10 to make integer
'genotypes for display
  NDPWR_INT = Round(10 ^ ND_INT)
  ReDim GENO_MAT(1 To NSIZE + 1, 1 To 3)
  GENO_MAT(1, 1) = FITNS_ARR(IFIT_ARR(NPOP_INT)) 'igen
  GENO_MAT(1, 2) = FITNS_ARR(IFIT_ARR(NPOP_INT - 1)) 'nnew
  GENO_MAT(1, 3) = FITNS_ARR(IFIT_ARR(NPOP_INT / 2)) 'pmut
  For k = 1 To NSIZE
    GENO_MAT(k + 1, 1) = Round(NDPWR_INT * OLDPH_MAT(k, IFIT_ARR(NPOP_INT)))
    GENO_MAT(k + 1, 2) = Round(NDPWR_INT * OLDPH_MAT(k, IFIT_ARR(NPOP_INT)))
    GENO_MAT(k + 1, 3) = Round(NDPWR_INT * OLDPH_MAT(k, IFIT_ARR(NPOP_INT)))
  Next k
End If

PIKAIA_REPORT_FUNC = GENO_MAT

Exit Function
ERROR_LABEL:
PIKAIA_REPORT_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : PIKAIA_ENCODE_FUNC
'DESCRIPTION   : Encodes phenotype into genotype
'LIBRARY       : OPTIMIZATION
'GROUP         : GENETIC_PIKAIA
'ID            : 006
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 02/13/2009
'************************************************************************************
'************************************************************************************

Private Function PIKAIA_ENCODE_FUNC(ByRef NSIZE As Integer, _
ByRef ND_INT As Integer, _
ByRef PH_MAT As Variant, _
ByRef GN_ARR As Variant)

Dim i As Long
Dim j As Long

Dim ii As Long
Dim jj As Long

Dim Z_VAL As Double

On Error GoTo ERROR_LABEL

PIKAIA_ENCODE_FUNC = False

'encode phenotype parameters into integer genotype
'PH_MAT(k) are x,y coordinates [ 0 < x,y < 1 ]

Z_VAL = 10 ^ ND_INT
ii = 0
For i = 1 To NSIZE
    jj = Int(PH_MAT(i) * Z_VAL)
    For j = ND_INT To 1 Step -1
        GN_ARR(ii + j) = jj Mod 10
        jj = Int(jj / 10)
            'gp debug add Int to force VBA not to round to nearest value

    Next j
    ii = ii + ND_INT
Next i

PIKAIA_ENCODE_FUNC = True

Exit Function
ERROR_LABEL:
PIKAIA_ENCODE_FUNC = False
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : PIKAIA_DECODE_FUNC
'DESCRIPTION   : decodes genotype into phenotype
'LIBRARY       : OPTIMIZATION
'GROUP         : GENETIC_PIKAIA
'ID            : 007
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 02/13/2009
'************************************************************************************
'************************************************************************************

Private Function PIKAIA_DECODE_FUNC(ByRef NSIZE As Integer, _
ByRef ND_INT As Integer, _
ByRef GN_ARR As Variant, _
ByRef PH_MAT As Variant)

'decode genotype into phenotype parameters
'PH_MAT(k) are x,y coordinates [ 0 < x,y < 1 ]

Dim i As Long
Dim j As Long
Dim ii As Long
Dim jj As Long
Dim Z_VAL As Double

On Error GoTo ERROR_LABEL

PIKAIA_DECODE_FUNC = False

Z_VAL = 10 ^ (-ND_INT)
ii = 0
For i = 1 To NSIZE
    jj = 0
    For j = 1 To ND_INT
        jj = Int(10 * jj + GN_ARR(ii + j))
            'gp add Int for force VBA not to round

    Next j
    PH_MAT(i) = jj * Z_VAL
    ii = ii + ND_INT
Next i

PIKAIA_DECODE_FUNC = True

Exit Function
ERROR_LABEL:
PIKAIA_DECODE_FUNC = False
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : PIKAIA_CROSS_FUNC
'DESCRIPTION   : Breeds two offspring from two parents
'LIBRARY       : OPTIMIZATION
'GROUP         : GENETIC_PIKAIA
'ID            : 008
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 02/13/2009
'************************************************************************************
'************************************************************************************

Private Function PIKAIA_CROSS_FUNC(ByRef NSIZE As Integer, _
ByRef ND_INT As Integer, _
ByRef PCROSS_VAL As Double, _
ByRef GN1_ARR As Variant, _
ByRef GN2_ARR As Variant)

'breeds two parent chromosomes into two offspring chromosomes
'breeding occurs through crossover. If the crossover probability
'test yields true (crossover taking place), either one-point or
'two-point crossover is used, with equal probabilities.
'Compatibility with version 1.0: To enforce 100% use of one-point
'crossover, un-comment appropriate line in source code below

Dim h As Integer
Dim i As Integer
Dim j As Integer '
Dim k As Integer
Dim l As Integer

On Error GoTo ERROR_LABEL

PIKAIA_CROSS_FUNC = False

'Use crossover probability to decide whether a crossover occurs
If (NEXT_PSEUDO_RND_FUNC() < PCROSS_VAL) Then
'Compute first crossover point
    j = Int(NEXT_PSEUDO_RND_FUNC() * NSIZE * ND_INT) + 1
    ' Now choose between one-point and two-point crossover
    If (NEXT_PSEUDO_RND_FUNC() < 0.5) Then
        k = NSIZE * ND_INT
    Else
        k = Int(NEXT_PSEUDO_RND_FUNC() * NSIZE * ND_INT) + 1
        'Un-comment following line to enforce one-point crossover
            If (k < j) Then
                l = k
                k = j
                j = l
            End If
    End If
    ' Swap genes from j to k
    For i = j To k
        h = GN2_ARR(i)
        GN2_ARR(i) = GN1_ARR(i)
        GN1_ARR(i) = h
    Next i
End If

PIKAIA_CROSS_FUNC = True

Exit Function
ERROR_LABEL:
PIKAIA_CROSS_FUNC = False
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : PIKAIA_MUTATE_FUNC
'DESCRIPTION   : Introduces random mutation in a genotype
'LIBRARY       : OPTIMIZATION
'GROUP         : GENETIC_PIKAIA
'ID            : 009
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 02/13/2009
'************************************************************************************
'************************************************************************************

Private Function PIKAIA_MUTATE_FUNC(ByRef NSIZE As Integer, _
ByRef ND_INT As Integer, _
ByRef PMUT_VAL As Double, _
ByRef GN_ARR As Variant, _
ByRef IMUT_INT As Integer)

'Mutations occur at rate PMUT_VAL at all gene loci
'IMUT_INT=1    Uniform mutation, constant rate
'IMUT_INT=2    Uniform mutation, variable rate based on fitness
'IMUT_INT=3    Uniform mutation, variable rate based on distance
'IMUT_INT=4    Uniform or creep mutation, constant rate
'IMUT_INT=5    Uniform or creep mutation, variable rate based on
'fitness
'IMUT_INT=6    Uniform or creep mutation, variable rate based on
'distance

Dim i As Integer
Dim j As Integer
Dim k As Integer
Dim l As Integer

Dim ii As Integer
Dim jj As Integer
Dim kk As Integer

On Error GoTo ERROR_LABEL

PIKAIA_MUTATE_FUNC = False

'Decide which type of mutation is to occur


If (IMUT_INT >= 4 And NEXT_PSEUDO_RND_FUNC() <= 0.5) Then
  'CREEP MUTATION OPERATOR
  'Subject each locus to random +/- 1 increment at the rate PMUT_VAL

    For i = 1 To NSIZE
        For j = 1 To ND_INT
        'Construct integer
            If (NEXT_PSEUDO_RND_FUNC() < PMUT_VAL) Then
                kk = (i - 1) * ND_INT + j
                jj = Round(NEXT_PSEUDO_RND_FUNC()) * 2 - 1
                ii = (i - 1) * ND_INT + 1
                GN_ARR(kk) = GN_ARR(kk) + jj
                If (jj < 0 And GN_ARR(kk) < 0) Then
        'This is where we carry over the one (up to two digits)
        'first take care of decrement below 0 case
                    If (j = 1) Then
                        GN_ARR(kk) = 0
                    Else
                        For k = kk To ii + 1 Step -1
                            GN_ARR(k) = 9
                            GN_ARR(k - 1) = GN_ARR(k - 1) - 1
                            If (GN_ARR(k - 1) >= 0) Then GoTo 4
                        Next k
                        If (GN_ARR(ii) < 0) Then
                            For l = ii To kk
                                GN_ARR(l) = 0
                            Next l
                        End If
4:
                    End If
                End If
                'we popped under 0.00000 lower bound; fix it up
                If (jj > 0 And GN_ARR(kk) > 9) Then
                    If (j = 1) Then
                        GN_ARR(kk) = 9
                    Else
                        For k = kk To ii + 1 Step -1
                            GN_ARR(k) = 0
                            GN_ARR(k - 1) = GN_ARR(k - 1) + 1
                            If (GN_ARR(k - 1) <= 9) Then GoTo 7
                        Next k
                        'we popped over 9.99999 upper bound; fix it up
                        If (GN_ARR(ii) > 9) Then
                            For l = ii To kk
                                GN_ARR(l) = 9
                            Next l
                        End If
7:
                    End If
                End If
            End If
        Next j
    Next i
Else
  'UNIFORM MUTATION OPERATOR
  'Subject each locus to random mutation at the rate PMUT_VAL
    
    For i = 1 To NSIZE * ND_INT
        If (NEXT_PSEUDO_RND_FUNC() < PMUT_VAL) Then
            GN_ARR(i) = Int(NEXT_PSEUDO_RND_FUNC() * 10)
        End If
    Next i
End If

PIKAIA_MUTATE_FUNC = True

Exit Function
ERROR_LABEL:
PIKAIA_MUTATE_FUNC = False
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : PIKAIA_PARENT_FUNC
'DESCRIPTION   :
'Selects a parent from the population, using roulette wheel
'algorithm with the relative fitnesses of the phenotypes as
'the "hit" probabilities [see Davis 1991, chap. 1].

'LIBRARY       : OPTIMIZATION
'GROUP         : GENETIC_PIKAIA
'ID            : 010
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 02/13/2009
'************************************************************************************
'************************************************************************************

Private Function PIKAIA_PARENT_FUNC(ByRef NPOP_INT As Integer, _
ByRef JFIT_ARR As Variant, _
ByRef FDIF_VAL As Double, _
ByRef IP_INT As Integer)

Dim i As Integer
Dim j As Integer

Dim DICE_VAL As Double
Dim FIT_VAL As Double

On Error GoTo ERROR_LABEL

PIKAIA_PARENT_FUNC = False

j = NPOP_INT + 1
DICE_VAL = NEXT_PSEUDO_RND_FUNC() * NPOP_INT * j
FIT_VAL = 0

For i = 1 To NPOP_INT
    FIT_VAL = FIT_VAL + j + FDIF_VAL * (j - 2 * JFIT_ARR(i))
    If (FIT_VAL >= DICE_VAL) Then
        IP_INT = i
        GoTo 2
    End If
Next i
2:
'Assert: loop will never exit by falling through

PIKAIA_PARENT_FUNC = True

Exit Function
ERROR_LABEL:
PIKAIA_PARENT_FUNC = False
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : PIKAIA_RANK_FUNC
'DESCRIPTION   : Ranks initial population
'LIBRARY       : OPTIMIZATION
'GROUP         : GENETIC_PIKAIA
'ID            : 011
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 02/13/2009
'************************************************************************************
'************************************************************************************

Private Function PIKAIA_RANK_FUNC(ByRef NPOP_INT As Integer, _
ByRef FITNS_ARR As Variant, _
ByRef IFIT_ARR As Variant, _
ByRef JFIT_ARR As Variant)

'Calls external sort routine to produce key index and rank order
'of input array FITNS_ARR (which is not altered).

Dim i As Integer

On Error GoTo ERROR_LABEL

PIKAIA_RANK_FUNC = False

Call PIKAIA_SORT_FUNC(NPOP_INT, FITNS_ARR, IFIT_ARR)
'External sort subroutine external PIKAIA_SORT_FUNC
'Compute the key index
'...and the rank order
For i = 1 To NPOP_INT
    JFIT_ARR(IFIT_ARR(i)) = NPOP_INT - i + 1
Next i

PIKAIA_RANK_FUNC = True

Exit Function
ERROR_LABEL:
PIKAIA_RANK_FUNC = False
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : PIKAIA_GENREP_FUNC
'DESCRIPTION   : Full generational replacement: accumulate offspring into new
'population array. Inserts offspring into population, for full generational replacement

'LIBRARY       : OPTIMIZATION
'GROUP         : GENETIC_PIKAIA
'ID            : 012
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 02/13/2009
'************************************************************************************
'************************************************************************************

Private Function PIKAIA_GENREP_FUNC(ByRef NO_PARAM As Integer, _
ByRef NSIZE As Integer, _
ByRef NPOP_INT As Integer, _
ByRef IPOP_INT As Integer, _
ByRef PH_MAT As Variant, _
ByRef NEWPH_MAT As Variant)

Dim i As Integer
Dim j As Integer
Dim k As Integer

On Error GoTo ERROR_LABEL

PIKAIA_GENREP_FUNC = False

'Insert one offspring pair into new population
i = 2 * IPOP_INT - 1
j = i + 1
For k = 1 To NSIZE
    NEWPH_MAT(k, i) = PH_MAT(k, 1)
    NEWPH_MAT(k, j) = PH_MAT(k, 2)
Next k

PIKAIA_GENREP_FUNC = True

Exit Function
ERROR_LABEL:
PIKAIA_GENREP_FUNC = False
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : PIKAIA_STDREP_FUNC
'DESCRIPTION   : Inserts offspring into population, for steady-state reproduction
'LIBRARY       : OPTIMIZATION
'GROUP         : GENETIC_PIKAIA
'ID            : 013
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 02/13/2009
'************************************************************************************
'************************************************************************************

Private Function PIKAIA_STDREP_FUNC(ByRef NO_PARAM As Integer, _
ByRef NSIZE As Integer, _
ByRef NPOP_INT As Integer, _
ByRef IREP_INT As Integer, _
ByRef IELITE_INT As Integer, _
ByRef PH_MAT As Variant, _
ByRef OLDPH_MAT As Variant, _
ByRef FITNS_ARR As Variant, _
ByRef IFIT_ARR As Variant, _
ByRef JFIT_ARR As Variant, _
ByRef NNEW_INT As Integer)

'steady-state reproduction: insert offspring pair into population
'only if they are fit enough (Replace-random if irep=2 or
'Replace-worst if irep=3).


Dim h As Integer
Dim i As Integer
Dim j As Integer
Dim k As Integer
Dim l As Integer

Dim FIT_VAL As Double

On Error GoTo ERROR_LABEL

PIKAIA_STDREP_FUNC = False

NNEW_INT = 0
'compute offspring fitness (with caller's fitness function)
For j = 1 To 2
'if fit enough, insert in population
    For k = 1 To NSIZE
        XTEMP_ARR(k) = PH_MAT(k, j)
    Next k
    FIT_VAL = PIKAIA_OBJ_FUNC(NSIZE, XTEMP_ARR())
    For i = NPOP_INT To 1 Step -1
        If (FIT_VAL > FITNS_ARR(IFIT_ARR(i))) Then
        'make sure the phenotype is not already in the population
            
            If (i < NPOP_INT) Then
                For k = 1 To NSIZE
                    If (OLDPH_MAT(k, IFIT_ARR(i + 1)) <> PH_MAT(k, j)) Then GoTo 6
                Next k
                GoTo 1
6:
            End If
            
'offspring is fit enough for insertion, and is unique
'(i) insert phenotype at appropriate place in population
            
            If (IREP_INT = 3) Then
                h = 1
            ElseIf (IELITE_INT = 0 Or i = NPOP_INT) Then
                h = Int(NEXT_PSEUDO_RND_FUNC() * NPOP_INT) + 1
            Else
                h = Int(NEXT_PSEUDO_RND_FUNC() * (NPOP_INT - 1)) + 1
            End If
            
            l = IFIT_ARR(h)
            FITNS_ARR(l) = FIT_VAL
            For k = 1 To NSIZE
                OLDPH_MAT(k, l) = PH_MAT(k, j)
            Next k
            'shift and update ranking arrays
            If (i < h) Then
                JFIT_ARR(l) = NPOP_INT - i
                For k = h - 1 To i + 1 Step -1
                    JFIT_ARR(IFIT_ARR(k)) = JFIT_ARR(IFIT_ARR(k)) - 1
                    IFIT_ARR(k + 1) = IFIT_ARR(k)
                Next k
                IFIT_ARR(i + 1) = l
            Else
                'shift down
                JFIT_ARR(l) = NPOP_INT - i + 1
                For k = h + 1 To i
                    JFIT_ARR(IFIT_ARR(k)) = JFIT_ARR(IFIT_ARR(k)) + 1
                    IFIT_ARR(k - 1) = IFIT_ARR(k)
                Next k
                IFIT_ARR(i) = l
            End If
            NNEW_INT = NNEW_INT + 1
GoTo 1
        End If
    Next i
1:
Next j

PIKAIA_STDREP_FUNC = True

Exit Function
ERROR_LABEL:
PIKAIA_STDREP_FUNC = False
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : PIKAIA_NEWPOP_FUNC
'DESCRIPTION   : Replaces old population by new; recomputes fitnesses & ranks
'LIBRARY       : OPTIMIZATION
'GROUP         : GENETIC_PIKAIA
'ID            : 014
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 02/13/2009
'************************************************************************************
'************************************************************************************

Private Function PIKAIA_NEWPOP_FUNC(ByRef IELITE_INT As Integer, _
ByRef NO_PARAM As Integer, _
ByRef NSIZE As Integer, _
ByRef NPOP_INT As Integer, _
ByRef OLDPH_MAT As Variant, _
ByRef NEWPH_MAT As Variant, _
ByRef IFIT_ARR As Variant, _
ByRef JFIT_ARR As Variant, _
ByRef FITNS_ARR As Variant, _
ByRef NNEW_INT As Integer)

Dim i As Integer
Dim k As Integer

On Error GoTo ERROR_LABEL

PIKAIA_NEWPOP_FUNC = False

'if using elitism, introduce in new population fittest of old
'population (if greater than fitness of the individual it is
'to Replace)

NNEW_INT = NPOP_INT
For k = 1 To NSIZE
    XTEMP_ARR(k) = NEWPH_MAT(k, 1)
Next k

If (IELITE_INT = 1 And _
    PIKAIA_OBJ_FUNC(NSIZE, XTEMP_ARR()) < FITNS_ARR(IFIT_ARR(NPOP_INT))) Then
    For k = 1 To NSIZE
        NEWPH_MAT(k, 1) = OLDPH_MAT(k, IFIT_ARR(NPOP_INT))
    Next k
    NNEW_INT = NNEW_INT - 1
End If

'Replace population
For i = 1 To NPOP_INT
    For k = 1 To NSIZE
        OLDPH_MAT(k, i) = NEWPH_MAT(k, i)
    Next k
    For k = 1 To NSIZE
        XTEMP_ARR(k) = OLDPH_MAT(k, i)
    Next k
    'get fitness using caller's fitness function
    FITNS_ARR(i) = PIKAIA_OBJ_FUNC(NSIZE, XTEMP_ARR())
Next i
'compute new population fitness rank order
Call PIKAIA_RANK_FUNC(NPOP_INT, FITNS_ARR, IFIT_ARR, JFIT_ARR)

PIKAIA_NEWPOP_FUNC = True

Exit Function
ERROR_LABEL:
PIKAIA_NEWPOP_FUNC = False
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : PIKAIA_SORT_FUNC
'DESCRIPTION   : Return integer array p which indexes array a in increasing order.
'Array A is not disturbed.  The Quicksort algorithm is used.
'B.G.Knapp, 86 / 12 / 23
'Reference: N. Wirth, Algorithms and Data Structures,
'Prentice - Hall, 1986

'LIBRARY       : OPTIMIZATION
'GROUP         : GENETIC_PIKAIA
'ID            : 015
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 02/13/2009
'************************************************************************************
'************************************************************************************

Private Function PIKAIA_SORT_FUNC(ByRef NROWS As Integer, _
ByRef FITNS_ARR As Variant, _
ByRef IFIT_ARR As Variant)

Const Q_VAL As Integer = 11
' Q_VAL = smallest subfile to use quicksort on
Const LGN_VAL As Integer = 32
' LGN_VAL = log base 2 of maximum n;

Dim i As Integer
Dim j As Integer
Dim l As Integer
Dim m As Integer
Dim r As Integer
Dim s As Integer
Dim t As Integer

Dim XDATA_VAL As Double

On Error GoTo ERROR_LABEL

Dim LTEMP_ARR(1 To LGN_VAL) As Integer
Dim RTEMP_ARR(1 To LGN_VAL) As Integer

PIKAIA_SORT_FUNC = False

'Initialize the stack
LTEMP_ARR(1) = 1
RTEMP_ARR(1) = NROWS
s = 1
'Initialize the pointer array
For i = 1 To NROWS
    IFIT_ARR(i) = i
Next i
2:

If s > 0 Then
    l = LTEMP_ARR(s)
    r = RTEMP_ARR(s)
    s = s - 1
3:
    If (r - l) < Q_VAL Then 'Use straight insertion
        For i = l + 1 To r
            t = IFIT_ARR(i)
            XDATA_VAL = FITNS_ARR(t)
            For j = i - 1 To l Step -1
                If FITNS_ARR(IFIT_ARR(j)) <= XDATA_VAL Then GoTo 5
                IFIT_ARR(j + 1) = IFIT_ARR(j)
            Next j
            j = l - 1
5:            IFIT_ARR(j + 1) = t
        Next i
    Else
        'Use quicksort, with pivot as median of FITNS_ARR(l), FITNS_ARR(m), FITNS_ARR(r)
        m = (l + r) / 2
        t = IFIT_ARR(m)
        If FITNS_ARR(t) < FITNS_ARR(IFIT_ARR(l)) Then
            IFIT_ARR(m) = IFIT_ARR(l)
            IFIT_ARR(l) = t
            t = IFIT_ARR(m)
        End If
        If FITNS_ARR(t) > FITNS_ARR(IFIT_ARR(r)) Then
            IFIT_ARR(m) = IFIT_ARR(r)
            IFIT_ARR(r) = t
            t = IFIT_ARR(m)
            If FITNS_ARR(t) < FITNS_ARR(IFIT_ARR(l)) Then
                IFIT_ARR(m) = IFIT_ARR(l)
                IFIT_ARR(l) = t
                t = IFIT_ARR(m)
            End If
        End If
        
        'Partition
        XDATA_VAL = FITNS_ARR(t)
        i = l + 1
        j = r - 1

7:      If i <= j Then
8:        If FITNS_ARR(IFIT_ARR(i)) < XDATA_VAL Then
            i = i + 1
            GoTo 8
          End If
9:        If XDATA_VAL < FITNS_ARR(IFIT_ARR(j)) Then
            j = j - 1
            GoTo 9
          End If
          
          If i <= j Then
            t = IFIT_ARR(i)
            IFIT_ARR(i) = IFIT_ARR(j)
            IFIT_ARR(j) = t
            i = i + 1
            j = j - 1
          End If
          GoTo 7
        End If
    
'Stack the larger subfile
        s = s + 1
        If (j - l) > (r - i) Then
            LTEMP_ARR(s) = l
            RTEMP_ARR(s) = j
            l = i
        Else
            LTEMP_ARR(s) = i
            RTEMP_ARR(s) = r
            r = j
        End If
        GoTo 3
    End If
    GoTo 2
End If

PIKAIA_SORT_FUNC = True

Exit Function
ERROR_LABEL:
PIKAIA_SORT_FUNC = False
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : NEXT_PSEUDO_RND_FUNC

'DESCRIPTION   : This routine does not take any arguments.  If the user wishes
'to be able to initialize NEXT_PSEUDO_RND_FUNC, so that the same sequence of
'random numbers can be repeated, this capability could be imple-
'mented with a separate subroutine, and called from the user's
'driver program.  An example NEXT_PSEUDO_RND_FUNC function (and initialization
'subroutine) which uses the function PSEUDO_RND_FUNC (the "minimal standard"
'random number generator of Park and Miller [Comm. ACM 31, 1192-
'1201, Oct 1988; Comm. ACM 36 No. 7, 105-110, July 1993]) is
'provided.

'LIBRARY       : OPTIMIZATION
'GROUP         : GENETIC_PIKAIA
'ID            : 016
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 02/13/2009
'************************************************************************************
'************************************************************************************

Private Function NEXT_PSEUDO_RND_FUNC()

'Return the next pseudo-random deviate from a sequence which is
'uniformly distributed in the interval [0,1]
'Uses the function PSEUDO_RND_FUNC, the "minimal standard" random number
'generator of Park and Miller (Comm. ACM 31, 1192-1201, Oct 1988;
'Comm. ACM 36 No. 7, 105-110, July 1993).

'Common block to make PUB_SEED_VAL visible to INIT_RND_GENER_FUNC (and to save
'it between calls) common /rnseed/ PUB_SEED_VAL

On Error GoTo ERROR_LABEL
If PUB_SEED_VAL <= 0 Then PUB_SEED_VAL = 123456
NEXT_PSEUDO_RND_FUNC = PSEUDO_RND_FUNC(PUB_SEED_VAL)
Exit Function
ERROR_LABEL:
NEXT_PSEUDO_RND_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : INIT_RND_GENER_FUNC
'DESCRIPTION   : Initialize random number generator
'LIBRARY       : OPTIMIZATION
'GROUP         : GENETIC_PIKAIA
'ID            : 017
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 02/13/2009
'************************************************************************************
'************************************************************************************

Private Function INIT_RND_GENER_FUNC(ByRef SEED_VAL As Long)

On Error GoTo ERROR_LABEL
PUB_SEED_VAL = SEED_VAL
If PUB_SEED_VAL <= 0 Then PUB_SEED_VAL = 123456
Exit Function
ERROR_LABEL:
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : PSEUDO_RND_FUNC
'DESCRIPTION   :

'"Minimal standard" pseudo-random number generator of Park and
'Miller.  Returns a uniform random deviate r s.t. 0 < r < 1.0.
'Set SEED_VAL to any non-zero integer value to initialize a sequence,
'then do not change SEED_VAL between calls for successive deviates
'in the sequence.

'References:
'Park, S. and Miller, K., "Random Number Generators: Good Ones
' are Hard to Find", Comm. ACM 31, 1192-1201 (Oct. 1988)
'Park, S. and Miller, K., in "Remarks on Choosing and Imple-
' menting Random Number Generators", Comm. ACM 36 No. 7,
' 105-110 (July 1993)

'LIBRARY       : OPTIMIZATION
'GROUP         : GENETIC_PIKAIA
'ID            : 018
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 02/13/2009
'************************************************************************************
'************************************************************************************

Private Function PSEUDO_RND_FUNC(ByRef SEED_VAL As Long)

Const A_VAL As Long = 48271
Const M_VAL As Long = 2147483647
Const Q_VAL As Long = 44488
Const R_VAL As Long = 3399

Dim j As Long

Dim SCALE_VAL As Double
Dim epsilon As Double
Dim tolerance As Double

On Error GoTo ERROR_LABEL

SCALE_VAL = 1# / M_VAL
epsilon = 0.00000012
tolerance = 1# - epsilon

'Executable section
j = SEED_VAL / Q_VAL
SEED_VAL = A_VAL * (SEED_VAL - j * Q_VAL) - R_VAL * j
If SEED_VAL < 0 Then SEED_VAL = SEED_VAL + M_VAL
PSEUDO_RND_FUNC = MINIMUM_FUNC(SEED_VAL * SCALE_VAL, tolerance)

Exit Function
ERROR_LABEL:
PSEUDO_RND_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : PIKAIA_OBJ_FUNC
'DESCRIPTION   : This is the fitness function that is called from sub pikaia
'LIBRARY       : OPTIMIZATION
'GROUP         : GENETIC_PIKAIA
'ID            : 019
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 02/13/2009
'************************************************************************************
'************************************************************************************

Private Function PIKAIA_OBJ_FUNC(ByVal NROWS As Integer, _
ByRef XDATA_ARR As Variant)

Dim i As Integer
Dim YTEMP_VAL As Double
Dim PARAM_VECTOR() As Variant

ReDim PARAM_VECTOR(1 To NROWS, 1 To 1)

On Error GoTo ERROR_LABEL
For i = 1 To NROWS
    PARAM_ARR(i) = PIKAI_PARAM_SCALE_FUNC(i, XDATA_ARR(i))
    PARAM_VECTOR(i, 1) = PARAM_ARR(i)
Next i
YTEMP_VAL = Excel.Application.Run(PUB_FUNC_NAME_STR, PARAM_VECTOR)
PIKAIA_OBJ_FUNC = YTEMP_VAL

Exit Function
ERROR_LABEL:
PIKAIA_OBJ_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : PIKAI_PARAM_SCALE_FUNC
'DESCRIPTION   : This function scales the parameter values from the 0-1 fraction (frac)
'LIBRARY       : OPTIMIZATION
'GROUP         : GENETIC_PIKAIA
'ID            : 020
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 02/13/2009
'************************************************************************************
'************************************************************************************

Private Function PIKAI_PARAM_SCALE_FUNC(ByRef i As Integer, _
ByVal SCALE_FACTOR As Double)

'paramID = parameter ID number (1-256)
'frac = scaling fraction for the rate (0-1)
'paramScale = scaled paramter value in the original units for the parameter

On Error GoTo ERROR_LABEL

PIKAI_PARAM_SCALE_FUNC = LOWER_ARR(i) + SCALE_FACTOR * _
                        (UPPER_ARR(i) - LOWER_ARR(i))

Exit Function
ERROR_LABEL:
PIKAI_PARAM_SCALE_FUNC = Err.number
End Function
