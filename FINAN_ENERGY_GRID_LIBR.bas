Attribute VB_Name = "FINAN_ENERGY_GRID_LIBR"


Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.


'************************************************************************************
'************************************************************************************
'FUNCTION      : ENERGY_REAL_OPTION_LEASE_FUNC

'DESCRIPTION   : Many firms, faced with projects that do not meet their financial
'benchmarks, use the argument that these projects should be taken anyway because of
'"strategic considerations". In other words, it is argued that these projects
'will accomplish other goals for the firm or allow the firm to enter into other
'markets. There are cases where these strategic considerations are really
'referring to options embedded in projects - options to produce new products or
'expand into new markets.

'The differences between using option pricing and the "strategic considerations"
'argument are the following:

'1. Option pricing assigns value to only some of the "strategic considerations" that
'firms may have. It considers cases where the initial investment is necessary for the
'strategic option (to expand, for instance), and values those investments as options.
'However, strategic considerations that are not clearly defined or include generic
'terms such as "corporate image" or "growth potential" may not have any value from
'an option pricing standpoint.

'2. Option pricing attempts to put a dollar value on the "strategic consideration"
'being valued. As a consequence, the existence of strategic considerations does
'not guarantee that the project will be taken.

'LIBRARY       : ENERGY
'GROUP         : GRID
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/29/2009
'************************************************************************************
'************************************************************************************

Function ENERGY_REAL_OPTION_LEASE_FUNC(ByVal SPOT As Double, _
ByVal EXTRACTION_COST As Double, _
ByVal MAX_EXTRACTION As Double, _
ByVal ENH_EXTRACTION_COST As Double, _
ByVal ENH_MAX_EXTRACTION As Double, _
ByVal ENH_FIXED_COST As Double, _
Optional ByVal STEPS As Long = 10, _
Optional ByVal UP_STEP_SIZE As Double = 1.2, _
Optional ByVal DOWN_STEP_SIZE As Double = 0.9, _
Optional ByVal GROWTH_FACTOR As Double = 1.1, _
Optional ByVal OUTPUT As Integer = 0)

'SPOT --> 400#
'EXTRACT --> 200#
'MAX_EXTRACT --> 10000
'EXTRACT --> 240#
'MAX EXTRAC  --> 12500
'ENHACEMENT --> 4,000,000

Dim i As Long
Dim j As Long

Dim PROB_UP_MOVE As Double
Dim PROB_DOWN_MOVE As Double

Dim SPOT_MATRIX As Variant
Dim ENHANCE_MATRIX As Variant
Dim OPTION_MATRIX As Variant

On Error GoTo ERROR_LABEL

ReDim SPOT_MATRIX(1 To STEPS + 2, 1 To STEPS + 2)
ReDim ENHANCE_MATRIX(1 To STEPS + 2, 1 To STEPS + 1)
ReDim OPTION_MATRIX(1 To STEPS + 2, 1 To STEPS + 1)
 
'-------------------------FIRST PASS: SETTING UP THE GRIDS---------
SPOT_MATRIX(1, 1) = ("ASSET LATTICE")

For i = 1 To STEPS + 1
    SPOT_MATRIX(1, i + 1) = i - 1
    SPOT_MATRIX(i + 1, 1) = STEPS - i + 1
    
    If i <> (STEPS + 1) Then
        ENHANCE_MATRIX(1, i + 1) = i - 1
        ENHANCE_MATRIX(i + 1, 1) = STEPS - i + 1
        
        If ENH_FIXED_COST <> 0 Then
            OPTION_MATRIX(1, i + 1) = i - 1
            OPTION_MATRIX(i + 1, 1) = STEPS - i + 1
        End If
    End If
Next i

PROB_UP_MOVE = (GROWTH_FACTOR - DOWN_STEP_SIZE) / (UP_STEP_SIZE - DOWN_STEP_SIZE)
PROB_DOWN_MOVE = 1 - PROB_UP_MOVE

SPOT_MATRIX(STEPS + 2, 2) = SPOT


'-------------------------SECOND PASS: FILLING THE GRIDS-----------

SPOT_MATRIX(1, 1) = ("ASSET_GRID")
For i = (STEPS + 2) To 2 Step -1
    For j = 3 To (STEPS + 2)
        If SPOT_MATRIX(i, 1) < SPOT_MATRIX(1, j) Then
             SPOT_MATRIX(i, j) = SPOT_MATRIX(i, j - 1) * DOWN_STEP_SIZE
        ElseIf SPOT_MATRIX(i, 1) = SPOT_MATRIX(1, j) Then
             SPOT_MATRIX(i, j) = SPOT_MATRIX(i + 1, j - 1) * UP_STEP_SIZE
        End If
    Next j
Next i

ENHANCE_MATRIX(1, 1) = ("LEASE VALUE ASSUMING ENHANCEMENT IN PLACE")

For i = 2 To (STEPS + 2)
    For j = (STEPS + 1) To 2 Step -1
        If j <> (STEPS + 1) Then
            If ENHANCE_MATRIX(i, 1) <= ENHANCE_MATRIX(1, j) Then
                ENHANCE_MATRIX(i, j) = ((PROB_UP_MOVE * ENHANCE_MATRIX(i - 1, j + 1) + _
                PROB_DOWN_MOVE * ENHANCE_MATRIX(i, j + 1)) / GROWTH_FACTOR) + _
               (MAXIMUM_FUNC(0, SPOT_MATRIX(i, j) - ENH_EXTRACTION_COST) * _
               (ENH_MAX_EXTRACTION / GROWTH_FACTOR))
            End If
        Else
                ENHANCE_MATRIX(i, j) = MAXIMUM_FUNC(0, SPOT_MATRIX(i, j) - _
                    ENH_EXTRACTION_COST) * _
                    ENH_MAX_EXTRACTION / GROWTH_FACTOR
        End If
    Next j
Next i

OPTION_MATRIX(1, 1) = ("LEASE WITH OPTION FOR ENHANCEMENT")

For i = 2 To (STEPS + 2)
    For j = (STEPS + 1) To 2 Step -1
        If j <> (STEPS + 1) Then
            If OPTION_MATRIX(i, 1) <= OPTION_MATRIX(1, j) Then
                OPTION_MATRIX(i, j) = MAXIMUM_FUNC(((PROB_UP_MOVE * OPTION_MATRIX(i - 1, j + 1) + _
                PROB_DOWN_MOVE * OPTION_MATRIX(i, j + 1)) / GROWTH_FACTOR) + _
               (MAXIMUM_FUNC(0, SPOT_MATRIX(i, j) - EXTRACTION_COST) * _
               (MAX_EXTRACTION / GROWTH_FACTOR)), ENHANCE_MATRIX(i, j) - ENH_FIXED_COST)
            End If
        Else
                OPTION_MATRIX(i, j) = MAXIMUM_FUNC(MAXIMUM_FUNC(0, SPOT_MATRIX(i, j) - _
                    EXTRACTION_COST) * MAX_EXTRACTION / GROWTH_FACTOR, _
                    ENHANCE_MATRIX(i, j) - ENH_FIXED_COST)
        End If
    Next j
Next i

'----------------------------------------------------------------------------------
Select Case OUTPUT
'----------------------------------------------------------------------------------
Case 0
'-----------------------------------SOME HOUSE-KEEPING-----------------------------
    For j = 1 To UBound(OPTION_MATRIX, 2)
        For i = 1 To UBound(OPTION_MATRIX, 1)
            If IsEmpty(OPTION_MATRIX(i, j)) = True Then: OPTION_MATRIX(i, j) = ""
        Next i
    Next j
'----------------------------------------------------------------------------------
    ENERGY_REAL_OPTION_LEASE_FUNC = OPTION_MATRIX
'----------------------------------------------------------------------------------
Case 1
'-----------------------------------SOME HOUSE-KEEPING-----------------------------
    For j = 1 To UBound(ENHANCE_MATRIX, 2)
        For i = 1 To UBound(ENHANCE_MATRIX, 1)
            If IsEmpty(ENHANCE_MATRIX(i, j)) = True Then: ENHANCE_MATRIX(i, j) = ""
        Next i
    Next j
'----------------------------------------------------------------------------------
    ENERGY_REAL_OPTION_LEASE_FUNC = ENHANCE_MATRIX
'----------------------------------------------------------------------------------
Case Else
'-----------------------------------SOME HOUSE-KEEPING-----------------------------
    For j = 1 To UBound(SPOT_MATRIX, 2)
        For i = 1 To UBound(SPOT_MATRIX, 1)
            If IsEmpty(SPOT_MATRIX(i, j)) = True Then: SPOT_MATRIX(i, j) = ""
        Next i
    Next j
'----------------------------------------------------------------------------------
    ENERGY_REAL_OPTION_LEASE_FUNC = SPOT_MATRIX
End Select

Exit Function
ERROR_LABEL:
ENERGY_REAL_OPTION_LEASE_FUNC = Err.number
End Function
