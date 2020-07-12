VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsStaticMemory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'--------------------------------------------------------------
'A flip-flop is a circuit that can store a binary value as
'long as power is supplied to the circuit. A flip-flop can
'store the binary value 1 and 0. When a value has been set,
'the flip-flop remains in this state (ie a 1 or 0) until it
'is told to change states.
'For example, memory cells are implemented using flip-flops.
'A 32 bit register would contain 32 flip-flops to store each
'of the binary digits
'--------------------------------------------------------------

Option Explicit

Dim PUB_STACK_ARR(0 To 100, 0 To 5)
Dim PUB_LAST_ITEM  As Long
Dim PUB_STATIC_COUNTER As Long

Function Store(ByRef INDEX_STR As Variant, _
ByRef X1_VAL As Variant, _
Optional ByRef X2_VAL As Variant, _
Optional ByRef X3_VAL As Variant)
    Store = STATIC_STACK_FUNC("w", INDEX_STR, X1_VAL, X2_VAL, X3_VAL)
End Function

Property Get COUNTER()
COUNTER = PUB_STATIC_COUNTER
End Property

Function Recall(ByRef INDEX_STR As Variant, _
ByRef X1_VAL As Variant, _
Optional ByRef X2_VAL As Variant, _
Optional ByRef X3_VAL As Variant) As Boolean
    Recall = STATIC_STACK_FUNC("r", INDEX_STR, X1_VAL, X2_VAL, X3_VAL)
End Function

Private Function STATIC_STACK_FUNC(ByRef OPERATOR_STR As Variant, _
ByRef INDEX_STR As Variant, _
Optional ByRef X1_VAL As Variant, _
Optional ByRef X2_VAL As Variant, _
Optional ByRef X3_VAL As Variant) As Boolean

Dim j As Long
Dim k As Long

Dim NROWS As Long
Dim NCOLUMNS As Long
Dim FIND_FLAG As Boolean
Dim RESULT_FLAG As Boolean

RESULT_FLAG = False

If INDEX_STR = "" Then Exit Function

NROWS = UBound(PUB_STACK_ARR, 1)
NCOLUMNS = UBound(PUB_STACK_ARR, 2)

'----------------------------------------------------------------------
Select Case OPERATOR_STR
'----------------------------------------------------------------------
    Case "w"
'----------------------------------------------------------------------
        GoSub 1983
        If Not FIND_FLAG Then PUB_LAST_ITEM = j
        PUB_STACK_ARR(j, 0) = INDEX_STR
        PUB_STACK_ARR(j, 1) = X1_VAL
        PUB_STACK_ARR(j, 2) = X2_VAL
        PUB_STACK_ARR(j, 3) = X3_VAL
        RESULT_FLAG = True
        '..........
'----------------------------------------------------------------------
    Case "r"
'----------------------------------------------------------------------
        If INDEX_STR <> "" Then
            GoSub 1983
            If FIND_FLAG Then
                X1_VAL = PUB_STACK_ARR(j, 1)
                X2_VAL = PUB_STACK_ARR(j, 2)
                X3_VAL = PUB_STACK_ARR(j, 3)
                RESULT_FLAG = True
                '..........
            End If
        End If
'----------------------------------------------------------------------
    Case "c"
'----------------------------------------------------------------------
    
'----------------------------------------------------------------------
End Select
'----------------------------------------------------------------------

STATIC_STACK_FUNC = RESULT_FLAG

Exit Function
1983: 'Search_Record
    j = PUB_LAST_ITEM
    k = IndexBound(PUB_LAST_ITEM + 1, LBound(PUB_STACK_ARR, 1), _
             UBound(PUB_STACK_ARR, 1))
    FIND_FLAG = False
    Do
'----------------------------------------------------------------------
        If IsEmpty(PUB_STACK_ARR(j, 0)) Then
'----------------------------------------------------------------------
            j = k
            PUB_STATIC_COUNTER = PUB_STATIC_COUNTER + 1
            Exit Do
'----------------------------------------------------------------------
        ElseIf (j = k) Then
'----------------------------------------------------------------------
            Exit Do
'----------------------------------------------------------------------
        ElseIf PUB_STACK_ARR(j, 0) = INDEX_STR Then
'----------------------------------------------------------------------
            FIND_FLAG = True
            Exit Do
'----------------------------------------------------------------------
        Else
'----------------------------------------------------------------------
            j = j - 1
            If j < LBound(PUB_STACK_ARR, 1) Then j = UBound(PUB_STACK_ARR, 1)
'----------------------------------------------------------------------
        End If
'----------------------------------------------------------------------
    Loop
Return
End Function

Private Function IndexBound(ByRef i As Long, _
ByRef j As Long, _
ByRef k As Long)
If i < j Then i = k
If i > k Then i = j
IndexBound = i
End Function
