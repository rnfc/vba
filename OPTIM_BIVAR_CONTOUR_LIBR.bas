Attribute VB_Name = "OPTIM_BIVAR_CONTOUR_LIBR"

'-------------------------------------------------------------------------------------
Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.
'-------------------------------------------------------------------------------------

Private PUB_X_MID_VAL As Double
Private PUB_Y_MID_VAL As Double

Private PUB_X_SCALE_VAL As Double
Private PUB_Y_SCALE_VAL As Double

Private PUB_OBJ_FUNC_STR As String


'************************************************************************************
'************************************************************************************
'FUNCTION      : BIVAR_ZERO_PATH_FUNC
'DESCRIPTION   : This algorithm solves the implicit equation f(x, y) = 0
'returning a set of points (xi, yi) that satisfy the given equation

'LIBRARY       : OPTIMIZATION
'GROUP         : BIVAR_OPTIM
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 02/13/2009
'************************************************************************************
'************************************************************************************

Function BIVAR_ZERO_PATH_FUNC(ByVal FUNC_NAME_STR As String, _
ByRef CONST_RNG As Variant, _
Optional ByVal NTRIALS As Long = 16, _
Optional ByVal nSTEPS As Long = 25, _
Optional ByVal nLOOPS As Long = 2000)

PUB_OBJ_FUNC_STR = FUNC_NAME_STR

Dim i As Long
Dim j As Long
Dim k As Long

Dim ii As Long
Dim jj As Long

Dim NSIZE As Long
Dim COUNTER As Long

Dim OK_FLAG As Boolean
Dim ERROR_STR As String

Dim INIT_ARR(1 To 2) As Double
Dim DATA_MATRIX() As Double

Dim HI_VAL As Double
Dim HJ_VAL As Double
Dim DELTA_VAL As Double

Dim XMAX_VAL As Double
Dim XMIN_VAL As Double
Dim YMAX_VAL As Double
Dim YMIN_VAL As Double

Dim CONST_BOX As Variant

On Error GoTo ERROR_LABEL

CONST_BOX = CONST_RNG
NSIZE = UBound(CONST_BOX, 2)

If NSIZE <> 2 Then: GoTo ERROR_LABEL ' Only 2 variables here

XMIN_VAL = CONST_BOX(1, 1)
XMAX_VAL = CONST_BOX(2, 1)

YMIN_VAL = CONST_BOX(1, 2)
YMAX_VAL = CONST_BOX(2, 2)
'--------------------------
'set the global scale factors
PUB_X_SCALE_VAL = (XMAX_VAL - XMIN_VAL) / 2
PUB_Y_SCALE_VAL = (YMAX_VAL - YMIN_VAL) / 2
PUB_X_MID_VAL = (XMAX_VAL + XMIN_VAL) / 2
PUB_Y_MID_VAL = (YMAX_VAL + YMIN_VAL) / 2
'setting the step
DELTA_VAL = 1 / nSTEPS '1/50  1/100

'set the starting mesh
jj = Sqr(NTRIALS)
ii = (NTRIALS / jj)

HJ_VAL = 2 / jj
HI_VAL = 2 / ii
'algorithm zero finder begins
i = 0
j = 0
k = 0
Do
    INIT_ARR(1) = i * HI_VAL - 1 + HI_VAL * (Rnd - 0.5) / 5
    INIT_ARR(2) = j * HJ_VAL - 1 + HJ_VAL * (Rnd - 0.5) / 5
    Call BIVAR_START_POINT_FUNC(1, -1, 1, -1, INIT_ARR, DELTA_VAL, ERROR_STR)
    
    If ERROR_STR = "" Then
        'check if the point is never taken
        OK_FLAG = Not BIVAR_PROX_POINT_CONTOUR_FUNC(INIT_ARR(1), INIT_ARR(2), _
                            DATA_MATRIX, 3 * DELTA_VAL)
        
        If OK_FLAG Then
            'find a new contour line
            Call BIVAR_CONTOUR_ZERO_FUNC(DELTA_VAL, 1, -1, 1, -1, INIT_ARR, _
                DATA_MATRIX, ERROR_STR, nLOOPS)
            
            If ERROR_STR <> "" Then
                If ERROR_STR = "points discharged" Then
                    'MsgBox ERROR_STR
                Else
                    GoTo ERROR_LABEL
                End If
            Else
                COUNTER = COUNTER + 1
            End If
        End If
    End If
    k = k + 1
    i = i + 1
    If i > ii Then i = 0: j = j + 1
Loop Until j > jj

ReDim ATEMP_MATRIX(1 To UBound(DATA_MATRIX, 1) + 1, 1 To 3)

'Output result
j = 0
If COUNTER > 0 Then
    For i = 1 To UBound(DATA_MATRIX)
       If i > 1 Then
            If DATA_MATRIX(i, 3) > DATA_MATRIX(i - 1, 3) Then j = j + 1
       End If
       ATEMP_MATRIX(1 + j, 1) = PUB_X_SCALE_VAL * DATA_MATRIX(i, 1) + PUB_X_MID_VAL
       ATEMP_MATRIX(1 + j, 2) = PUB_Y_SCALE_VAL * DATA_MATRIX(i, 2) + PUB_Y_MID_VAL
       ATEMP_MATRIX(1 + j, 3) = BIVAR_ZERO_OBJ_FUNC(DATA_MATRIX(i, 1), DATA_MATRIX(i, 2))
       j = j + 1
    Next i
Else
    GoTo ERROR_LABEL 'Path not found
End If

BIVAR_ZERO_PATH_FUNC = ATEMP_MATRIX

Exit Function
ERROR_LABEL:
BIVAR_ZERO_PATH_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : BIVAR_INTERSECTION_FUNC
'DESCRIPTION   : Return an array with the intersection points
'LIBRARY       : OPTIMIZATION
'GROUP         : BIVAR_OPTIM
'ID            : 002
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 02/13/2009
'************************************************************************************
'************************************************************************************

Public Function BIVAR_INTERSECTION_FUNC(ByRef FDATA_RNG As Variant, _
ByRef GDATA_RNG As Variant)

Dim i As Long
Dim j As Long
Dim k As Long

Dim X_VAL As Double
Dim Y_VAL As Double

Dim XTEMP_ARR() As Double
Dim YTEMP_ARR() As Double

Dim TEMP_MATRIX() As Variant
    
Dim PDATA_VECTOR As Variant
Dim QDATA_VECTOR As Variant

Dim CONVERG_VAL As Integer

On Error GoTo ERROR_LABEL

PDATA_VECTOR = FDATA_RNG
If UBound(PDATA_VECTOR, 1) = 1 Then: PDATA_VECTOR = MATRIX_TRANSPOSE_FUNC(PDATA_VECTOR)
QDATA_VECTOR = GDATA_RNG
If UBound(QDATA_VECTOR, 1) = 1 Then: QDATA_VECTOR = MATRIX_TRANSPOSE_FUNC(QDATA_VECTOR)

k = 0

For i = 1 To UBound(PDATA_VECTOR) - 1
    For j = 1 To UBound(QDATA_VECTOR) - 1
        If PDATA_VECTOR(i, 1) <> "" And _
           PDATA_VECTOR(i + 1, 1) <> "" And _
           QDATA_VECTOR(j, 1) <> "" And _
           QDATA_VECTOR(j + 1, 1) <> "" Then
            
            Call BIVAR_SEGM_INTERS_FUNC(CDbl(PDATA_VECTOR(i, 1)), _
               CDbl(PDATA_VECTOR(i, 2)), _
               CDbl(PDATA_VECTOR(i + 1, 1)), _
               CDbl(PDATA_VECTOR(i + 1, 2)), _
               CDbl(QDATA_VECTOR(j, 1)), _
               CDbl(QDATA_VECTOR(j, 2)), _
               CDbl(QDATA_VECTOR(j + 1, 1)), _
               CDbl(QDATA_VECTOR(j + 1, 2)), _
               X_VAL, Y_VAL, CONVERG_VAL)
            
            If CONVERG_VAL > 0 Then
                k = k + 1
                ReDim Preserve XTEMP_ARR(1 To k)
                ReDim Preserve YTEMP_ARR(1 To k)
                XTEMP_ARR(k) = X_VAL
                YTEMP_ARR(k) = Y_VAL
            End If
        End If
    Next j
Next i
If k > 0 Then
'--------------------------------------------------------------------------------------
'--------------------------------------------------------------------------------------
'Each value satisfies the given system with an accuracy of about 1E-3 (0.1%),
'sufficient for graphical representation or as starting point for other, more power,
'rootfinding algorithms (Newton, Broyden, Brown, etc.)
'--------------------------------------------------------------------------------------
'--------------------------------------------------------------------------------------

    ReDim TEMP_MATRIX(0 To k, 1 To 2)
    For i = 1 To k
        TEMP_MATRIX(i, 1) = XTEMP_ARR(i)
        TEMP_MATRIX(i, 2) = YTEMP_ARR(i)
    Next i
    TEMP_MATRIX(0, 1) = "X_VAR"
    TEMP_MATRIX(0, 2) = "Y_VAR"
Else
    GoTo ERROR_LABEL
End If

BIVAR_INTERSECTION_FUNC = TEMP_MATRIX

Exit Function
ERROR_LABEL:
BIVAR_INTERSECTION_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : BIVAR_SEGM_INTERS_FUNC

'DESCRIPTION   :
' In the 2D plane, computes the cross point x of two segments S12 and S23
' S12: (x1,y1)-(x2,y2)
' S34: (x3,y3)-(x4,y4)
' CONVERG_VAL =  1 cross point x is internal
' CONVERG_VAL = -1 cross point x is external
' CONVERG_VAL =  0 no interesction (parallel segments)


'LIBRARY       : OPTIMIZATION
'GROUP         : BIVAR_OPTIM
'ID            : 003
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 02/13/2009
'************************************************************************************
'************************************************************************************


Private Function BIVAR_SEGM_INTERS_FUNC(ByRef X1_VAL As Double, _
ByRef Y1_VAL As Double, _
ByRef X2_VAL As Double, _
ByRef Y2_VAL As Double, _
ByRef X3_VAL As Double, _
ByRef Y3_VAL As Double, _
ByRef X4_VAL As Double, _
ByRef Y4_VAL As Double, _
ByRef X_VAL As Double, _
ByRef Y_VAL As Double, _
ByRef CONVERG_VAL As Integer)

Dim T_VAL As Double
Dim S_VAL As Double
Dim D_VAL As Double

Dim TEMP_MATRIX() As Double

Dim tolerance As Double

On Error GoTo ERROR_LABEL

ReDim TEMP_MATRIX(1 To 2, 1 To 3)

tolerance = 2 * 10 ^ -16

TEMP_MATRIX(1, 1) = X2_VAL - X1_VAL
TEMP_MATRIX(2, 1) = Y2_VAL - Y1_VAL
TEMP_MATRIX(1, 2) = X3_VAL - X4_VAL
TEMP_MATRIX(2, 2) = Y3_VAL - Y4_VAL
TEMP_MATRIX(1, 3) = X3_VAL - X1_VAL
TEMP_MATRIX(2, 3) = Y3_VAL - Y1_VAL

D_VAL = TEMP_MATRIX(1, 1) * TEMP_MATRIX(2, 2) - _
        TEMP_MATRIX(2, 1) * TEMP_MATRIX(1, 2)

If Abs(D_VAL) < tolerance Then CONVERG_VAL = 0: Exit Function

T_VAL = (TEMP_MATRIX(1, 3) * TEMP_MATRIX(2, 2) - _
         TEMP_MATRIX(2, 3) * TEMP_MATRIX(1, 2)) / D_VAL

S_VAL = (TEMP_MATRIX(1, 1) * TEMP_MATRIX(2, 3) - _
         TEMP_MATRIX(2, 1) * TEMP_MATRIX(1, 3)) / D_VAL
X_VAL = TEMP_MATRIX(1, 1) * T_VAL + X1_VAL
Y_VAL = TEMP_MATRIX(2, 1) * T_VAL + Y1_VAL

If T_VAL >= 0 And T_VAL <= 1 And S_VAL >= 0 And _
   S_VAL <= 1 Then CONVERG_VAL = 1 Else CONVERG_VAL = -1

Exit Function
ERROR_LABEL:
BIVAR_SEGM_INTERS_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : BIVAR_CONTOUR_ZERO_FUNC
'DESCRIPTION   : draw a contour-line of zeros for a bivariate function f(x,y)
'the contour-line can be open or close

'LIBRARY       : OPTIMIZATION
'GROUP         : BIVAR_OPTIM
'ID            : 004
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 02/13/2009
'************************************************************************************
'************************************************************************************


Private Function BIVAR_CONTOUR_ZERO_FUNC(ByRef FACTOR_VAL As Double, _
ByRef XMAX_VAL As Double, _
ByRef XMIN_VAL As Double, _
ByRef YMAX_VAL As Double, _
ByRef YMIN_VAL As Double, _
ByRef INIT_ARR() As Double, _
ByRef DATA_MATRIX() As Double, _
ByRef ERROR_STR As String, _
Optional nLOOPS As Long = 1000)

Dim i As Long
Dim j As Long
Dim k As Long
Dim l As Long

Dim ii As Long
Dim jj As Long
Dim kk As Long
Dim ll As Long

Dim HT_VAL As Double
Dim HN_VAL As Double

Dim FA_VAL As Double
Dim FB_VAL As Double
Dim FC_VAL As Double
Dim FD_VAL As Double

Dim TL_VAL As Double

Dim ATEMP_MATRIX() As Double
Dim BTEMP_MATRIX() As Double

Dim ATEMP_ARR(1 To 2) As Double
Dim BTEMP_ARR(1 To 2) As Double
Dim CTEMP_ARR(1 To 2) As Double
Dim DTEMP_ARR(1 To 2) As Double
Dim ETEMP_ARR(1 To 2) As Double
Dim FTEMP_ARR(1 To 2) As Double
Dim GTEMP_ARR(1 To 2) As Double
Dim HTEMP_ARR(1 To 2) As Double
Dim ITEMP_ARR(1 To 2) As Double
Dim JTEMP_ARR(1 To 2) As Double

Dim OK_FLAG As Boolean
Dim END_FLAG As Boolean

Dim CONVERG_VAL As Integer

On Error GoTo ERROR_LABEL

ERROR_STR = ""
CONVERG_VAL = 0
'save the starting position. It must be on the zero line
JTEMP_ARR(1) = INIT_ARR(1)
JTEMP_ARR(2) = INIT_ARR(2)
kk = 1
'choose a kk
BIVAR_GRADIENT_FUNC FACTOR_VAL, JTEMP_ARR, FTEMP_ARR
If BIVAR_VALID_PARAM_FUNC(FTEMP_ARR) Then
    'take the normal kk
    CTEMP_ARR(1) = FTEMP_ARR(1)
    CTEMP_ARR(2) = FTEMP_ARR(2)
Else
    'take a random kk
    CTEMP_ARR(1) = Rnd - 0.5
    CTEMP_ARR(2) = Rnd - 0.5
    BIVAR_VALID_PARAM_FUNC CTEMP_ARR
End If
'point p0 and kk T
HTEMP_ARR(1) = CTEMP_ARR(2)
HTEMP_ARR(2) = -CTEMP_ARR(1)
'save the starting kk
ITEMP_ARR(1) = HTEMP_ARR(1)
ITEMP_ARR(2) = HTEMP_ARR(2)

ReDim BTEMP_MATRIX(1 To nLOOPS, 1 To 2)

i = 1
BTEMP_MATRIX(i, 1) = JTEMP_ARR(1)
BTEMP_MATRIX(i, 2) = JTEMP_ARR(2)
Do
    'prediction values
    OK_FLAG = False
    HT_VAL = FACTOR_VAL
    HN_VAL = FACTOR_VAL / 2
    j = 0
    
    Do
        Call BIVAR_PREDICTOR_FUNC(ATEMP_ARR, INIT_ARR, HTEMP_ARR, CTEMP_ARR, HT_VAL, HN_VAL)
        Call BIVAR_PREDICTOR_FUNC(BTEMP_ARR, INIT_ARR, HTEMP_ARR, CTEMP_ARR, HT_VAL, -HN_VAL)
        
        'check the function sign at the points A and B
        
        FA_VAL = BIVAR_ZERO_OBJ_FUNC(ATEMP_ARR(1), ATEMP_ARR(2))
        FB_VAL = BIVAR_ZERO_OBJ_FUNC(BTEMP_ARR(1), BTEMP_ARR(2))
        
        If FA_VAL * FB_VAL <= 0 Then
            'find the zero performing Pegasus
            Call BIVAR_PEGASUS_ZERO_FUNC(ATEMP_ARR, BTEMP_ARR, _
                    GTEMP_ARR, 100, CONVERG_VAL)
            If CONVERG_VAL < 0 Then
                ERROR_STR = "unable to find zero = " & _
                            Chr(13) & ATEMP_ARR(1) & "," & ATEMP_ARR(2) & _
                            " " & BTEMP_ARR(1) & "," & BTEMP_ARR(2)
                GoTo ERROR_LABEL
            End If
            OK_FLAG = True
        Else
            DTEMP_ARR(1) = ATEMP_ARR(1)
            DTEMP_ARR(2) = ATEMP_ARR(2)
            FC_VAL = FA_VAL
            ETEMP_ARR(1) = BTEMP_ARR(1)
            ETEMP_ARR(2) = BTEMP_ARR(2)
            FD_VAL = FB_VAL
            HT_VAL = HT_VAL + FACTOR_VAL
            HN_VAL = FACTOR_VAL / 2
            j = j + 1
        End If
    Loop Until j > 1 Or OK_FLAG = True
    
    If Not OK_FLAG Then
        'try to left and right alternatively
        HT_VAL = -FACTOR_VAL / 2
        HN_VAL = FACTOR_VAL / 2
        j = 0
        Do
            Call BIVAR_PREDICTOR_FUNC(ATEMP_ARR, INIT_ARR, HTEMP_ARR, _
                    CTEMP_ARR, HT_VAL, HN_VAL)
            FA_VAL = BIVAR_ZERO_OBJ_FUNC(ATEMP_ARR(1), ATEMP_ARR(2))
            If FA_VAL * FC_VAL < 0 Then
                'find the zero performing Pegasus
                Call BIVAR_PEGASUS_ZERO_FUNC(ATEMP_ARR, DTEMP_ARR, _
                        GTEMP_ARR, 100, CONVERG_VAL)
                If CONVERG_VAL < 0 Then
                    ERROR_STR = "unable to find zero = " & _
                                Chr(13) & ATEMP_ARR(1) & "," & ATEMP_ARR(2) & _
                                " " & DTEMP_ARR(1) & "," & DTEMP_ARR(2)
                    GoTo ERROR_LABEL
                End If
                OK_FLAG = True
                Exit Do
            End If
            Call BIVAR_PREDICTOR_FUNC(BTEMP_ARR, INIT_ARR, HTEMP_ARR, _
                                 CTEMP_ARR, HT_VAL, -HN_VAL)
            
            FB_VAL = BIVAR_ZERO_OBJ_FUNC(BTEMP_ARR(1), BTEMP_ARR(2))
            If FB_VAL * FD_VAL < 0 Then
                'find the zero performing Pegasus
                Call BIVAR_PEGASUS_ZERO_FUNC(BTEMP_ARR, ETEMP_ARR, _
                                GTEMP_ARR, 100, CONVERG_VAL)
                If CONVERG_VAL < 0 Then
                    ERROR_STR = "unable to find zero = " & _
                                Chr(13) & BTEMP_ARR(1) & "," & BTEMP_ARR(2) & " " & _
                                ETEMP_ARR(1) & "," & ETEMP_ARR(2)
                    GoTo ERROR_LABEL
                End If
                OK_FLAG = True
                Exit Do
            End If
            HN_VAL = HN_VAL / 2
            'save old values
            DTEMP_ARR(1) = ATEMP_ARR(1)
            DTEMP_ARR(2) = ATEMP_ARR(2)
            FC_VAL = FA_VAL
            ETEMP_ARR(1) = BTEMP_ARR(1)
            ETEMP_ARR(2) = BTEMP_ARR(2)
            FD_VAL = FB_VAL
            j = j + 1
        Loop Until j > 10
    End If
    'save point
    i = i + 1
    BTEMP_MATRIX(i, 1) = GTEMP_ARR(1)
    BTEMP_MATRIX(i, 2) = GTEMP_ARR(2)
    'new kk
    HTEMP_ARR(1) = GTEMP_ARR(1) - INIT_ARR(1)
    HTEMP_ARR(2) = GTEMP_ARR(2) - INIT_ARR(2)
    TL_VAL = BIVAR_NORM_FUNC(HTEMP_ARR)
    HTEMP_ARR(1) = HTEMP_ARR(1) / TL_VAL
    HTEMP_ARR(2) = HTEMP_ARR(2) / TL_VAL
    'check if the current point has come-back to the starting point
    If i - ll > 2 Then _
        END_FLAG = _
        BIVAR_PROX_POINT_SEGMENT_FUNC(JTEMP_ARR(1), JTEMP_ARR(2), _
                    GTEMP_ARR(1), GTEMP_ARR(2), INIT_ARR(1), _
                    INIT_ARR(2), FACTOR_VAL / 3)
    If END_FLAG Then
        'closing point
        i = i + 1
        BTEMP_MATRIX(i, 1) = JTEMP_ARR(1)
        BTEMP_MATRIX(i, 2) = JTEMP_ARR(2)
    Else
        'check if the point is out of the drawing-box
         If (GTEMP_ARR(1) < 1.5 * XMIN_VAL Or GTEMP_ARR(1) > 1.5 * XMAX_VAL) Or _
            (GTEMP_ARR(2) < 1.5 * YMIN_VAL Or GTEMP_ARR(2) > 1.5 * YMAX_VAL) Then
            If kk = 1 Then
                'go to the starting point and go in back kk
                kk = -1
                HTEMP_ARR(1) = -ITEMP_ARR(1)
                HTEMP_ARR(2) = -ITEMP_ARR(2)
                GTEMP_ARR(1) = JTEMP_ARR(1)
                GTEMP_ARR(2) = JTEMP_ARR(2)
                ll = i
            Else
                END_FLAG = True
            End If
         End If
    End If
    CTEMP_ARR(1) = -HTEMP_ARR(2)
    CTEMP_ARR(2) = HTEMP_ARR(1)
    INIT_ARR(1) = GTEMP_ARR(1)
    INIT_ARR(2) = GTEMP_ARR(2)
    'check corner
    
    If i > 4 Then
        If BIVAR_CHECK_CORNER1_FUNC(ATEMP_ARR, BTEMP_MATRIX, i) Then
            Call BIVAR_STORE_CORNER_FUNC(ATEMP_ARR, BTEMP_MATRIX, i)
        End If
    End If
   
   'check point proximity to another contour line
   If i Mod 10 = 0 Or l > 0 Then
        If BIVAR_PROX_POINT_CONTOUR_FUNC(INIT_ARR(1), INIT_ARR(2), _
                DATA_MATRIX, FACTOR_VAL) Then
            l = l + 1
            If l > 3 Then
                ERROR_STR = "points discharged"
                GoTo ERROR_LABEL
            End If
        Else
            l = 0   'reset
        End If
   End If
'   If i Mod 4 = 0 Then
        'Debug.Print "Running...  (path: " & (COUNTER + 1) & _
            "  points: " & i & ")"
'   End If
   'If Flag_Stop Then Exit Do  'user stop
Loop Until i >= nLOOPS Or END_FLAG

If i >= nLOOPS Then
    ERROR_STR = "too many points:" & i
    GoTo ERROR_LABEL
End If
k = i
'save data into the Contours List
On Error Resume Next
ii = UBound(DATA_MATRIX)
If Err = 0 Then
    On Error GoTo 0
    'contours already exist
    ReDim ATEMP_MATRIX(1 To ii, 1 To 3)
    For i = 1 To ii
        For j = 1 To 3
            ATEMP_MATRIX(i, j) = DATA_MATRIX(i, j)
        Next j
    Next i
    ReDim DATA_MATRIX(1 To ii + k, 1 To 3)
    For i = 1 To ii
        For j = 1 To 3
            DATA_MATRIX(i, j) = ATEMP_MATRIX(i, j)
        Next j
    Next i
    jj = DATA_MATRIX(ii, 3) + 1
    Erase ATEMP_MATRIX
Else
    On Error GoTo 0
    'first contour
    ii = 0
    ReDim DATA_MATRIX(1 To k, 1 To 3)
    jj = 1
End If
'add the current contour to the list
j = 0
For i = ll To 1 Step -1
    j = j + 1
    DATA_MATRIX(j + ii, 1) = BTEMP_MATRIX(i, 1)
    DATA_MATRIX(j + ii, 2) = BTEMP_MATRIX(i, 2)
    DATA_MATRIX(j + ii, 3) = jj
Next i
For i = ll + 1 To k
    j = j + 1
    DATA_MATRIX(j + ii, 1) = BTEMP_MATRIX(i, 1)
    DATA_MATRIX(j + ii, 2) = BTEMP_MATRIX(i, 2)
    DATA_MATRIX(j + ii, 3) = jj
Next i
Erase BTEMP_MATRIX
    '
Exit Function
ERROR_LABEL:
BIVAR_CONTOUR_ZERO_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : BIVAR_START_POINT_FUNC
'DESCRIPTION   : Find a single zero of the f(x,y) random point
'LIBRARY       : OPTIMIZATION
'GROUP         : BIVAR_OPTIM
'ID            : 005
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 02/13/2009
'************************************************************************************
'************************************************************************************

Private Function BIVAR_START_POINT_FUNC(ByRef XMAX_VAL As Double, _
ByRef XMIN_VAL As Double, _
ByRef YMAX_VAL As Double, _
ByRef YMIN_VAL As Double, _
ByRef PARAM_ARR() As Double, _
ByRef FACTOR_VAL As Double, _
ByRef ERROR_STR As String)

Dim i As Integer
Dim FA_VAL As Double
Dim FB_VAL As Double

Dim ATEMP_ARR(1 To 2) As Double
Dim BTEMP_ARR(1 To 2) As Double
Dim CTEMP_ARR(1 To 2) As Double
Dim DTEMP_ARR(1 To 2) As Double

Dim CONVERG_VAL As Integer

On Error GoTo ERROR_LABEL

CONVERG_VAL = 0
ERROR_STR = ""
ATEMP_ARR(1) = PARAM_ARR(1)
ATEMP_ARR(2) = PARAM_ARR(2)
'find the nearest zero
FA_VAL = BIVAR_ZERO_OBJ_FUNC(ATEMP_ARR(1), ATEMP_ARR(2))
If FA_VAL = 0 Then
    PARAM_ARR(1) = ATEMP_ARR(1)
    PARAM_ARR(2) = ATEMP_ARR(2)
    CONVERG_VAL = 0
    Exit Function
End If
BIVAR_GRADIENT_FUNC FACTOR_VAL, ATEMP_ARR, CTEMP_ARR
If BIVAR_VALID_PARAM_FUNC(CTEMP_ARR) Then
    'take the normal direction
    DTEMP_ARR(1) = -Sgn(FA_VAL) * CTEMP_ARR(1)
    DTEMP_ARR(2) = -Sgn(FA_VAL) * CTEMP_ARR(2)
Else
    'take BTEMP_ARR random direction
    DTEMP_ARR(1) = Rnd - 0.5
    DTEMP_ARR(2) = Rnd - 0.5
    BIVAR_VALID_PARAM_FUNC DTEMP_ARR
End If
i = 0
Do
    BTEMP_ARR(1) = ATEMP_ARR(1) + FACTOR_VAL * DTEMP_ARR(1)
    BTEMP_ARR(2) = ATEMP_ARR(2) + FACTOR_VAL * DTEMP_ARR(2)
    FB_VAL = BIVAR_ZERO_OBJ_FUNC(BTEMP_ARR(1), BTEMP_ARR(2))
    If FB_VAL * FA_VAL < 0 Then Exit Do
    ATEMP_ARR(1) = BTEMP_ARR(1)
    ATEMP_ARR(2) = BTEMP_ARR(2)
    FA_VAL = FB_VAL
Loop While XMIN_VAL <= BTEMP_ARR(1) And _
           BTEMP_ARR(1) <= XMAX_VAL And _
           YMIN_VAL <= BTEMP_ARR(2) And _
           BTEMP_ARR(2) <= YMAX_VAL

If FB_VAL * FA_VAL < 0 Then
    'find the zero performing Pegasus
    Call BIVAR_PEGASUS_ZERO_FUNC(BTEMP_ARR, ATEMP_ARR, _
                                 PARAM_ARR, 100, CONVERG_VAL)
    If CONVERG_VAL < 0 Then
        ERROR_STR = _
        "unable to find starting point = " & _
        Chr(13) & BTEMP_ARR(1) & "," & BTEMP_ARR(2) & " " & _
        ATEMP_ARR(1) & "," & ATEMP_ARR(2)
        GoTo ERROR_LABEL
    End If
Else
    ERROR_STR = "unable to find starting point"
    GoTo ERROR_LABEL
End If

Exit Function
ERROR_LABEL:
BIVAR_START_POINT_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : BIVAR_VALID_PARAM_FUNC
'DESCRIPTION   :
'LIBRARY       : OPTIMIZATION
'GROUP         : BIVAR_OPTIM
'ID            : 006
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 02/13/2009
'************************************************************************************
'************************************************************************************

Private Function BIVAR_VALID_PARAM_FUNC(ByRef PARAM_ARR() As Double) As Boolean
Dim F_VAL As Double
On Error GoTo ERROR_LABEL

F_VAL = BIVAR_NORM_FUNC(PARAM_ARR)
If F_VAL > 0 Then
    PARAM_ARR(1) = PARAM_ARR(1) / F_VAL
    PARAM_ARR(2) = PARAM_ARR(2) / F_VAL
    BIVAR_VALID_PARAM_FUNC = True
Else
    BIVAR_VALID_PARAM_FUNC = False
End If

Exit Function
ERROR_LABEL:
BIVAR_VALID_PARAM_FUNC = False
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : BIVAR_PREDICTOR_FUNC
'DESCRIPTION   :
'LIBRARY       : OPTIMIZATION
'GROUP         : BIVAR_OPTIM
'ID            : 007
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 02/13/2009
'************************************************************************************
'************************************************************************************

Private Function BIVAR_PREDICTOR_FUNC(ByRef PARAM_ARR() As Double, _
ByRef INIT_ARR() As Double, _
ByRef MULT_ARR() As Double, _
ByRef VAR_ARR() As Double, _
ByRef A_DELTA As Double, _
ByRef B_DELTA As Double)

On Error GoTo ERROR_LABEL

PARAM_ARR(1) = INIT_ARR(1) + A_DELTA * MULT_ARR(1) + B_DELTA * VAR_ARR(1)
PARAM_ARR(2) = INIT_ARR(2) + A_DELTA * MULT_ARR(2) + B_DELTA * VAR_ARR(2)

Exit Function
ERROR_LABEL:
BIVAR_PREDICTOR_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : BIVAR_GRADIENT_FUNC
'DESCRIPTION   :
'LIBRARY       : OPTIMIZATION
'GROUP         : BIVAR_OPTIM
'ID            : 008
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 02/13/2009
'************************************************************************************
'************************************************************************************

Private Function BIVAR_GRADIENT_FUNC(ByRef FACTOR_VAL As Double, _
ByRef PARAM_ARR() As Double, _
ByRef RESULT_ARR() As Double)

Dim F1_VAL As Double
Dim F2_VAL As Double
Dim F3_VAL As Double
Dim F4_VAL As Double

On Error GoTo ERROR_LABEL

F1_VAL = BIVAR_ZERO_OBJ_FUNC(PARAM_ARR(1) - FACTOR_VAL, PARAM_ARR(2))
F2_VAL = BIVAR_ZERO_OBJ_FUNC(PARAM_ARR(1) + FACTOR_VAL, PARAM_ARR(2))

F3_VAL = BIVAR_ZERO_OBJ_FUNC(PARAM_ARR(1), PARAM_ARR(2) - FACTOR_VAL)
F4_VAL = BIVAR_ZERO_OBJ_FUNC(PARAM_ARR(1), PARAM_ARR(2) + FACTOR_VAL)

RESULT_ARR(1) = (F2_VAL - F1_VAL) / FACTOR_VAL
RESULT_ARR(2) = (F4_VAL - F3_VAL) / FACTOR_VAL

Exit Function
ERROR_LABEL:
BIVAR_GRADIENT_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : BIVAR_NORM_FUNC
'DESCRIPTION   :
'LIBRARY       : OPTIMIZATION
'GROUP         : BIVAR_OPTIM
'ID            : 009
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 02/13/2009
'************************************************************************************
'************************************************************************************

Private Function BIVAR_NORM_FUNC(ByRef PARAM_ARR() As Double)
On Error GoTo ERROR_LABEL

BIVAR_NORM_FUNC = Sqr(PARAM_ARR(1) ^ 2 + PARAM_ARR(2) ^ 2)

Exit Function
ERROR_LABEL:
BIVAR_NORM_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : BIVAR_PEGASUS_ZERO_FUNC
'DESCRIPTION   :
'LIBRARY       : OPTIMIZATION
'GROUP         : BIVAR_OPTIM
'ID            : 010
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 02/13/2009
'************************************************************************************
'************************************************************************************

Private Function BIVAR_PEGASUS_ZERO_FUNC(ByRef LOWER_ARR() As Double, _
ByRef UPPER_ARR() As Double, _
ByRef PARAM_ARR() As Double, _
ByRef nLOOPS As Long, _
ByRef CONVERG_VAL As Integer)

Dim i As Long

Dim F1_VAL As Double
Dim F2_VAL As Double
Dim F3_VAL As Double

Dim T1_VAL As Double
Dim T2_VAL As Double
Dim T3_VAL As Double

Dim DELTA_VAL As Double

Dim ATEMP_ARR(1 To 2) As Double
Dim BTEMP_ARR(1 To 2) As Double

Dim epsilon As Double

On Error GoTo ERROR_LABEL

'-------------------------------------------------------------------
'  CONVERG_VAL:
'          =-2, the maximum number of iteration steps was
'               reached without meeting the break-off criterion.
'          =-1, the condition FCT(LOWER_ARR)*FCT(UPPER_ARR) < 0.0 is not met.
'          = 0, LOWER_ARR or UPPER_ARR already are LOWER_ARR zero of FCT.
'          = 1, PARAM_ARR is LOWER_ARR zero with |F(PARAM_ARR)| < 4* epsilon
'               constant.
'          = 2, break-off criterion has been met.
'               The absolute error of the computed zero is small
'               but the value |F(PARAM_ARR)| is not small.
'               It usually happens if F is discontinue
'-------------------------------------------------------------------

    ATEMP_ARR(1) = LOWER_ARR(1)
    ATEMP_ARR(2) = LOWER_ARR(2)
    T1_VAL = 0
    BTEMP_ARR(1) = UPPER_ARR(1)
    BTEMP_ARR(2) = UPPER_ARR(2)
    T2_VAL = 1
    
    epsilon = 2 * 10 ^ -16
'
      F1_VAL = BIVAR_ZERO_OBJ_FUNC(ATEMP_ARR(1), ATEMP_ARR(2))
      F2_VAL = BIVAR_ZERO_OBJ_FUNC(BTEMP_ARR(1), BTEMP_ARR(2))
'
      If (F1_VAL * F2_VAL > 0) Then
         CONVERG_VAL = -1
         Exit Function
      ElseIf (F1_VAL * F2_VAL = 0) Then
        If F1_VAL = 0 Then PARAM_ARR(1) = ATEMP_ARR(1)
            PARAM_ARR(2) = ATEMP_ARR(2)
        If F2_VAL = 0 Then PARAM_ARR(1) = BTEMP_ARR(1)
            PARAM_ARR(2) = BTEMP_ARR(2)
         CONVERG_VAL = 0
         Exit Function
      End If
      
'  executing the Pegasus-method.
'
      For i = 1 To nLOOPS
        If (Abs(F2_VAL) < 4 * epsilon) Then
            PARAM_ARR(1) = LOWER_ARR(1) + (UPPER_ARR(1) - LOWER_ARR(1)) * T2_VAL
            PARAM_ARR(2) = LOWER_ARR(2) + (UPPER_ARR(2) - LOWER_ARR(2)) * T2_VAL
            CONVERG_VAL = 1
            Exit Function
        End If
'     testing for the break-off criterion.
         If (Abs(T2_VAL - T1_VAL) <= Abs(T2_VAL) * 4 * epsilon) Then
            If (Abs(F1_VAL) < Abs(F2_VAL)) Then
                PARAM_ARR(1) = LOWER_ARR(1) + (UPPER_ARR(1) - LOWER_ARR(1)) * T1_VAL
                PARAM_ARR(2) = LOWER_ARR(2) + (UPPER_ARR(2) - LOWER_ARR(2)) * T1_VAL
            Else
                PARAM_ARR(1) = LOWER_ARR(1) + (UPPER_ARR(1) - LOWER_ARR(1)) * T2_VAL
                PARAM_ARR(2) = LOWER_ARR(2) + (UPPER_ARR(2) - LOWER_ARR(2)) * T2_VAL
            End If
            CONVERG_VAL = 2
            Exit Function
         Else
'     calculating the secant slope.
            DELTA_VAL = (F2_VAL - F1_VAL) / (T2_VAL - T1_VAL)
'     calculating the secant intercept T3_VAL.
            T3_VAL = T2_VAL - F2_VAL / DELTA_VAL
'     calculating LOWER_ARR new functional value at T3_VAL.
            PARAM_ARR(1) = LOWER_ARR(1) + (UPPER_ARR(1) - LOWER_ARR(1)) * T3_VAL
            PARAM_ARR(2) = LOWER_ARR(2) + (UPPER_ARR(2) - LOWER_ARR(2)) * T3_VAL
            F3_VAL = BIVAR_ZERO_OBJ_FUNC(PARAM_ARR(1), PARAM_ARR(2))
            If (F2_VAL * F3_VAL <= 0) Then
               T1_VAL = T2_VAL
               F1_VAL = F2_VAL
            Else
               F1_VAL = F1_VAL * F2_VAL / (F2_VAL + F3_VAL)
            End If
            T2_VAL = T3_VAL
            F2_VAL = F3_VAL
         End If
      Next i
      CONVERG_VAL = -2

Exit Function
ERROR_LABEL:
BIVAR_PEGASUS_ZERO_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : BIVAR_PROX_POINT_SEGMENT_FUNC
'DESCRIPTION   : Check if a point is proxime to the segment
'LIBRARY       : OPTIMIZATION
'GROUP         : BIVAR_OPTIM
'ID            : 011
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 02/13/2009
'************************************************************************************
'************************************************************************************

Private Function BIVAR_PROX_POINT_SEGMENT_FUNC(ByRef X_VAL As Double, _
ByRef Y_VAL As Double, _
ByRef X1_VAL As Double, _
ByRef Y1_VAL As Double, _
ByRef X2_VAL As Double, _
ByRef Y2_VAL As Double, _
ByRef FACTOR_VAL As Double)

Dim A_VAL As Double

Dim D1_VAL As Double
Dim D2_VAL As Double

Dim VX_VAL As Double
Dim VY_VAL As Double
Dim PX_VAL As Double
Dim PY_VAL As Double
Dim NX_VAL As Double
Dim NY_VAL As Double
Dim TX_VAL As Double
Dim TY_VAL As Double
Dim VL_VAL As Double
Dim TL_VAL As Double
Dim NL_VAL As Double

Dim TEMP_FLAG As Boolean

On Error GoTo ERROR_LABEL

VX_VAL = X2_VAL - X1_VAL
VY_VAL = Y2_VAL - Y1_VAL
VL_VAL = Sqr(VX_VAL ^ 2 + VY_VAL ^ 2)
PX_VAL = X_VAL - X1_VAL
PY_VAL = Y_VAL - Y1_VAL
A_VAL = (PX_VAL * VX_VAL + PY_VAL * VY_VAL) / VL_VAL
TX_VAL = A_VAL * VX_VAL / VL_VAL
TY_VAL = A_VAL * VY_VAL / VL_VAL
TL_VAL = Sqr(TX_VAL ^ 2 + TY_VAL ^ 2)
NX_VAL = PX_VAL - TX_VAL
NY_VAL = PY_VAL - TY_VAL
NL_VAL = Sqr(NX_VAL ^ 2 + NY_VAL ^ 2)

If A_VAL >= 0 And TL_VAL <= VL_VAL And NL_VAL <= FACTOR_VAL Then
    TEMP_FLAG = True
Else
    D1_VAL = Sqr((X_VAL - X1_VAL) ^ 2 + (Y_VAL - Y1_VAL) ^ 2)
    D2_VAL = Sqr((X_VAL - X2_VAL) ^ 2 + (Y_VAL - Y2_VAL) ^ 2)
    If D1_VAL <= FACTOR_VAL Or D2_VAL <= FACTOR_VAL Then TEMP_FLAG = True
End If
BIVAR_PROX_POINT_SEGMENT_FUNC = TEMP_FLAG

Exit Function
ERROR_LABEL:
BIVAR_PROX_POINT_SEGMENT_FUNC = False
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : BIVAR_PROX_POINT_CONTOUR_FUNC
'DESCRIPTION   :
'LIBRARY       : OPTIMIZATION
'GROUP         : BIVAR_OPTIM
'ID            : 012
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 02/13/2009
'************************************************************************************
'************************************************************************************

Private Function BIVAR_PROX_POINT_CONTOUR_FUNC(ByRef X_VAL As Double, _
ByRef Y_VAL As Double, _
ByRef DATA_MATRIX() As Double, _
ByRef FACTOR_VAL As Double)

Dim i As Long
Dim NSIZE As Long
Dim CHECK_FLAG As Boolean
Dim TEMP_DELTA As Double

On Error GoTo ERROR_LABEL

TEMP_DELTA = FACTOR_VAL / 3
On Error GoTo 1983
NSIZE = UBound(DATA_MATRIX)
On Error GoTo 0
For i = 1 To NSIZE - 1
    If DATA_MATRIX(i, 3) = DATA_MATRIX(i + 1, 3) Then
        If (DATA_MATRIX(i, 1) - FACTOR_VAL < X_VAL And _
            X_VAL < DATA_MATRIX(i + 1, 1) + FACTOR_VAL) And _
            DATA_MATRIX(i, 2) - FACTOR_VAL < Y_VAL And _
            Y_VAL < DATA_MATRIX(i + 1, 2) + FACTOR_VAL Then
                CHECK_FLAG = BIVAR_PROX_POINT_SEGMENT_FUNC(X_VAL, Y_VAL, _
                    DATA_MATRIX(i, 1), _
                    DATA_MATRIX(i, 2), _
                    DATA_MATRIX(i + 1, 1), _
                    DATA_MATRIX(i + 1, 2), _
                    TEMP_DELTA)
            If CHECK_FLAG = True Then Exit For
        End If
    End If
Next i
1983:
BIVAR_PROX_POINT_CONTOUR_FUNC = CHECK_FLAG

Exit Function
ERROR_LABEL:
BIVAR_PROX_POINT_CONTOUR_FUNC = False
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : BIVAR_CHECK_CORNER2_FUNC
'DESCRIPTION   : Compute the intersection point
'LIBRARY       : OPTIMIZATION
'GROUP         : BIVAR_OPTIM
'ID            : 013
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 02/13/2009
'************************************************************************************
'************************************************************************************

Private Function BIVAR_CHECK_CORNER2_FUNC(ByRef PARAM_ARR() As Double, _
ByRef DATA_MATRIX() As Double, _
ByRef COUNTER As Long)

Dim i As Long
Dim j As Long

Dim W_VAL As Double
Dim FA_VAL As Double
Dim FB_VAL As Double
Dim FM_VAL As Double

Dim CONVERG_VAL As Integer
Dim DIFF_VAL As Double

Dim INIT_ARR(1 To 2) As Double
Dim P1_ARR(1 To 2) As Double
Dim P2_ARR(1 To 2) As Double
Dim P3_ARR(1 To 2) As Double

Dim PM_ARR(1 To 2) As Double

Dim PA_ARR(1 To 2) As Double
Dim PB_ARR(1 To 2) As Double

Dim A_ARR(1 To 2) As Double
Dim B_ARR(1 To 2) As Double

Dim epsilon As Double
Dim CHECK_FLAG As Boolean

On Error GoTo ERROR_LABEL

CHECK_FLAG = False
epsilon = 5 * 10 ^ -16
For j = 1 To 2
    INIT_ARR(j) = DATA_MATRIX(COUNTER, j)
    P1_ARR(j) = DATA_MATRIX(COUNTER - 1, j)
    P2_ARR(j) = DATA_MATRIX(COUNTER - 2, j)
    P3_ARR(j) = DATA_MATRIX(COUNTER - 3, j)
Next j

A_ARR(1) = INIT_ARR(1) - P1_ARR(1)
A_ARR(2) = INIT_ARR(2) - P1_ARR(2)
B_ARR(1) = P2_ARR(1) - P3_ARR(1)
B_ARR(2) = P2_ARR(2) - P3_ARR(2)

W_VAL = (A_ARR(1) * B_ARR(1) + A_ARR(2) * B_ARR(2)) / _
        (BIVAR_NORM_FUNC(A_ARR) * BIVAR_NORM_FUNC(B_ARR))
If W_VAL < 0.86 Then
    'direction changed
    j = 0
    Do
        PM_ARR(1) = (P1_ARR(1) + P2_ARR(1)) / 2
        PM_ARR(2) = (P1_ARR(2) + P2_ARR(2)) / 2
        FM_VAL = BIVAR_ZERO_OBJ_FUNC(PM_ARR(1), PM_ARR(2))
        If Abs(FM_VAL) <= epsilon Then GoTo 1983
        A_ARR(1) = PM_ARR(1)
        A_ARR(2) = PM_ARR(2)
        i = 0
        Do
            A_ARR(1) = A_ARR(1) + P2_ARR(1) - P3_ARR(1)
            A_ARR(2) = A_ARR(2) + P2_ARR(2) - P3_ARR(2)
            FA_VAL = BIVAR_ZERO_OBJ_FUNC(A_ARR(1), A_ARR(2))
            If FA_VAL * FM_VAL < 0 Then
                Call BIVAR_PEGASUS_ZERO_FUNC(A_ARR, PM_ARR, PA_ARR, 100, CONVERG_VAL)
                Exit Do
            End If
            i = i + 1
        Loop Until i > 8
        If i > 8 Or CONVERG_VAL < 0 Then GoTo 1983
        B_ARR(1) = PM_ARR(1)
        B_ARR(2) = PM_ARR(2)
        i = 0
        Do
            B_ARR(1) = B_ARR(1) + P1_ARR(1) - INIT_ARR(1)
            B_ARR(2) = B_ARR(2) + P1_ARR(2) - INIT_ARR(2)
            FB_VAL = BIVAR_ZERO_OBJ_FUNC(B_ARR(1), B_ARR(2))
            If FB_VAL * FM_VAL < 0 Then
                Call BIVAR_PEGASUS_ZERO_FUNC(B_ARR, PM_ARR, _
                    PB_ARR, 100, CONVERG_VAL)
                Exit Do
            End If
            i = i + 1
        Loop Until i > 8
        If i > 8 Or CONVERG_VAL < 0 Then GoTo 1983
        
        DIFF_VAL = Abs(PA_ARR(1) - PB_ARR(1)) + Abs(PA_ARR(2) - PB_ARR(2))
        If DIFF_VAL <= 10000 * epsilon Then
            CHECK_FLAG = True
            PARAM_ARR(1) = PA_ARR(1)
            PARAM_ARR(2) = PB_ARR(2)
            Exit Do
        End If
        INIT_ARR(1) = P1_ARR(1)
        INIT_ARR(2) = P1_ARR(2)
        P1_ARR(1) = PA_ARR(1)
        P1_ARR(2) = PA_ARR(2)
        P3_ARR(1) = P2_ARR(1)
        P3_ARR(2) = P2_ARR(2)
        P2_ARR(1) = PB_ARR(1)
        P2_ARR(2) = PB_ARR(2)
        j = j + 1
    Loop Until j > 100
End If
1983:
BIVAR_CHECK_CORNER2_FUNC = CHECK_FLAG

Exit Function
ERROR_LABEL:
BIVAR_CHECK_CORNER2_FUNC = False
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : BIVAR_CHECK_CORNER1_FUNC
'DESCRIPTION   : Compute the intersection point
'LIBRARY       : OPTIMIZATION
'GROUP         : BIVAR_OPTIM
'ID            : 014
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 02/13/2009
'************************************************************************************
'************************************************************************************

Private Function BIVAR_CHECK_CORNER1_FUNC(ByRef PARAM_ARR() As Double, _
ByRef DATA_MATRIX() As Double, _
ByRef COUNTER As Long)

Dim j As Long

Dim T1_VAL As Double
Dim T2_VAL As Double

Dim DET_VAL As Double
Dim DET1_VAL As Double
Dim DET2_VAL As Double

Dim W_VAL As Double
Dim F_VAL As Double

Dim DL1_VAL As Double
Dim DL3_VAL As Double

Dim D1_ARR(1 To 2) As Double
Dim D2_ARR(1 To 2) As Double
Dim D3_ARR(1 To 2) As Double

Dim CHECK_FLAG As Boolean

Dim epsilon As Double

On Error GoTo ERROR_LABEL

epsilon = 5 * 10 ^ -16
For j = 1 To 2
    D1_ARR(j) = DATA_MATRIX(COUNTER, j) - DATA_MATRIX(COUNTER - 1, j)
    D2_ARR(j) = DATA_MATRIX(COUNTER - 1, j) - DATA_MATRIX(COUNTER - 2, j)
    D3_ARR(j) = DATA_MATRIX(COUNTER - 2, j) - DATA_MATRIX(COUNTER - 3, j)
Next j

DL1_VAL = BIVAR_NORM_FUNC(D1_ARR)
DL3_VAL = BIVAR_NORM_FUNC(D3_ARR)
W_VAL = (D1_ARR(1) * D3_ARR(1) + D1_ARR(2) * D3_ARR(2)) / (DL1_VAL * DL3_VAL)
If W_VAL < 0.86 Then
    'direction changed
    DET_VAL = D1_ARR(1) * D3_ARR(2) - D1_ARR(2) * D3_ARR(1)
    If DET_VAL <> 0 Then
        DET1_VAL = D2_ARR(1) * D3_ARR(2) - D2_ARR(2) * D3_ARR(1)
        DET2_VAL = D1_ARR(1) * D2_ARR(2) - D1_ARR(2) * D2_ARR(1)
        T1_VAL = DET1_VAL / DET_VAL
        T2_VAL = DET2_VAL / DET_VAL
        If T1_VAL > 0.2 And T2_VAL > 0.2 Then
            PARAM_ARR(1) = DATA_MATRIX(COUNTER - 1, 1) - D1_ARR(1) * T1_VAL
            PARAM_ARR(2) = DATA_MATRIX(COUNTER - 1, 2) - D1_ARR(2) * T1_VAL
            F_VAL = BIVAR_ZERO_OBJ_FUNC(PARAM_ARR(1), PARAM_ARR(2))
            If Abs(F_VAL) <= 1000 * epsilon Then
                CHECK_FLAG = True
            Else
                'try check_corner1
                CHECK_FLAG = BIVAR_CHECK_CORNER2_FUNC(PARAM_ARR, _
                             DATA_MATRIX, COUNTER)
            End If
        End If
    End If
End If
BIVAR_CHECK_CORNER1_FUNC = CHECK_FLAG

Exit Function
ERROR_LABEL:
BIVAR_CHECK_CORNER1_FUNC = False
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : BIVAR_STORE_CORNER_FUNC
'DESCRIPTION   : Store the corner in the contour-list at position COUNTER
'LIBRARY       : OPTIMIZATION
'GROUP         : BIVAR_OPTIM
'ID            : 015
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 02/13/2009
'************************************************************************************
'************************************************************************************

Private Function BIVAR_STORE_CORNER_FUNC(ByRef PARAM_ARR() As Double, _
ByRef DATA_MATRIX() As Double, _
ByRef COUNTER As Long)

Dim j As Long

On Error GoTo ERROR_LABEL

COUNTER = COUNTER + 1
For j = 1 To 2
    DATA_MATRIX(COUNTER, j) = DATA_MATRIX(COUNTER - 1, j)
    DATA_MATRIX(COUNTER - 1, j) = DATA_MATRIX(COUNTER - 2, j)
    DATA_MATRIX(COUNTER - 2, j) = PARAM_ARR(j)
Next j

Exit Function
ERROR_LABEL:
BIVAR_STORE_CORNER_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : BIVAR_ZERO_OBJ_FUNC
'DESCRIPTION   :
'LIBRARY       : OPTIMIZATION
'GROUP         : BIVAR_OPTIM
'ID            : 016
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 02/13/2009
'************************************************************************************
'************************************************************************************

Private Function BIVAR_ZERO_OBJ_FUNC(ByVal XS_VAL As Double, _
ByVal YS_VAL As Double)

Dim X_VAL As Double
Dim Y_VAL As Double
Dim PARAM_VECTOR(1 To 2, 1 To 1) As Double

On Error GoTo ERROR_LABEL

'coordinates rescaling
X_VAL = PUB_X_SCALE_VAL * XS_VAL + PUB_X_MID_VAL
Y_VAL = PUB_Y_SCALE_VAL * YS_VAL + PUB_Y_MID_VAL
PARAM_VECTOR(1, 1) = X_VAL
PARAM_VECTOR(2, 1) = Y_VAL
BIVAR_ZERO_OBJ_FUNC = Excel.Application.Run(PUB_OBJ_FUNC_STR, PARAM_VECTOR)

Exit Function
ERROR_LABEL:
BIVAR_ZERO_OBJ_FUNC = Err.number
End Function
