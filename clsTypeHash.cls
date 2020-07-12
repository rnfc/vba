VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsTypeHash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'--------------------------------------------------------------------------------
'--------------------------------------------------------------------------------
Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 0       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.
'-------------------------------------------------------------------------------
'-------------------------------------------------------------------------------
' HASHTABLE class module
'-------------------------------------------------------------------------------
' This class implements a hashtable, a structure that offers many
' of the features of a collectior or dictionary, and is often
' even faster than the built-in collection.
'-------------------------------------------------------------------------------
'-------------------------------------------------------------------------------

Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias _
"RtlMoveMemory" (Dest As Any, _
source As Any, _
ByVal bytes As Long)

'-------------------------------------------------------------------------------
'-------------------------------------------------------------------------------

' default values
Const DEFAULT_HASH_SIZE_VAL = 1024
Const DEFAULT_LIST_SIZE_VAL = 2048
Const DEFAULT_CHUNK_SIZE_VAL = 1024

'-------------------------------------------------------------------------------
'-------------------------------------------------------------------------------

Private Type PUB_SLOT_TYPE_OBJ
    Key As String
    Data As Variant
    Next As Long      ' 0 if last item
End Type

'-----------------------------------------------------------------------------
'-----------------------------------------------------------------------------

' for each hash code this array holds the first element
' in PUB_SLOT_TABLE() with the corresponding hash code
Dim PUB_HASH_TABLE() As Long
' the array that holds the data
Dim PUB_SLOT_TABLE() As PUB_SLOT_TYPE_OBJ

' pointer to first free slot
Dim PUB_FREE_INDEX As Long

' size of hash table
Dim PUB_HASH_SIZE_VAL As Long
' size of slot table
Dim PUB_LIST_SIZE_VAL As Long
' chunk size
Dim PUB_CHUNK_SIZE_VAL As Long
' items in the slot table
Dim PUB_COUNTER_VAL As Long

' member variable for IgnoreCase property
Private PUB_IGNORE_CASE_FLAG As Boolean

'-----------------------------------------------------------------------------
'-----------------------------------------------------------------------------


' initialize the hash table

Sub SetSize(ByVal HASH_SIZE_VAL As Long, _
Optional ByVal LIST_SIZE_VAL As Long, _
Optional ByVal CHUNK_SIZE_VAL As Long)
    ' provide defaults
    If LIST_SIZE_VAL <= 0 Then LIST_SIZE_VAL = PUB_LIST_SIZE_VAL
    If CHUNK_SIZE_VAL <= 0 Then CHUNK_SIZE_VAL = PUB_CHUNK_SIZE_VAL
    ' save size values
    PUB_HASH_SIZE_VAL = HASH_SIZE_VAL
    PUB_LIST_SIZE_VAL = LIST_SIZE_VAL
    PUB_CHUNK_SIZE_VAL = CHUNK_SIZE_VAL
    PUB_COUNTER_VAL = 0
    ' rebuild tables
    PUB_FREE_INDEX = 0
    ReDim PUB_HASH_TABLE(0 To HASH_SIZE_VAL - 1) As Long
    ReDim PUB_SLOT_TABLE(0 To 0) As PUB_SLOT_TYPE_OBJ '0 to 0
    ExpandSlotTable PUB_LIST_SIZE_VAL
End Sub


Property Let IgnoreCase(ByVal NEW_VAL_FLAG As Boolean)
    If PUB_COUNTER_VAL Then
        Err.Raise 1001, , "The Hash Table isn't empty"
    End If
    PUB_IGNORE_CASE_FLAG = NEW_VAL_FLAG
End Property


' True if keys are searched in case-unsensitive mode
' this can be assigned to only when the hash table is empty

Property Get IgnoreCase() As Boolean
    IgnoreCase = PUB_IGNORE_CASE_FLAG
End Property


' add a new element to the hash table

Sub Add(KEY_STR As String, _
DATA_VAL As Variant)
    Dim k As Long
    Dim CREATE_FLAG As Boolean
    
    ' get the index to the slot where the value is
    ' (allocate a new slot if necessary)
    CREATE_FLAG = True
    k = GetSlotIndex(KEY_STR, CREATE_FLAG)
    
    If CREATE_FLAG Then
        ' the item was actually added
        If IsObject(DATA_VAL) Then
            Set PUB_SLOT_TABLE(k).Data = DATA_VAL
        Else
            PUB_SLOT_TABLE(k).Data = DATA_VAL
        End If
    Else
        ' raise error _
        "This key is already associated with an item of this collection"
        Err.Raise 457
    End If
End Sub

Property Get GetKey(INDEX_VAL As Long) As String
    GetKey = PUB_SLOT_TABLE(INDEX_VAL + 1).Key
End Property

' the value associated to a key (empty if not found)

Property Get Item(KEY_STR As String) As Variant
Attribute Item.VB_UserMemId = 0
    Dim k As Long
    ' get the index to the slot where the value is
    k = GetSlotIndex(KEY_STR)
    If k = 0 Then
        ' return Empty if not found
    ElseIf IsObject(PUB_SLOT_TABLE(k).Data) Then
        Set Item = PUB_SLOT_TABLE(k).Data
    Else
        Item = PUB_SLOT_TABLE(k).Data
    End If
End Property

Property Let Item(KEY_STR As String, _
DATA_VAL As Variant)
    Dim k As Long
    ' get the index to the slot where the value is
    ' (allocate a new slot if necessary)
    k = GetSlotIndex(KEY_STR, True)
    ' store the value
    PUB_SLOT_TABLE(k).Data = DATA_VAL
End Property

Property Set Item(KEY_STR As String, _
DATA_OBJ As Object)
    Dim k As Long
    ' get the index to the slot where the value is
    ' (allocate a new slot if necessary)
    k = GetSlotIndex(KEY_STR, True)
    ' store the value
    Set PUB_SLOT_TABLE(k).Data = DATA_OBJ
End Property


' check whether an item is in the hash table

Function Exists(KEY_STR As String) As Boolean
    Exists = GetSlotIndex(KEY_STR) <> 0
End Function

' remove an item from the hash table

Sub Remove(KEY_STR As String)
    Dim k As Long
    Dim CODE_VAL As Long
    Dim LAST_INDEX_VAL As Long
    k = GetSlotIndex(KEY_STR, False, CODE_VAL, LAST_INDEX_VAL)
    ' raise error if no such element
    If k = 0 Then Err.Raise 5
    
    If LAST_INDEX_VAL Then
        ' this isn't the first item in the PUB_SLOT_TABLE() array
        PUB_SLOT_TABLE(LAST_INDEX_VAL).Next = _
        PUB_SLOT_TABLE(k).Next
    ElseIf PUB_SLOT_TABLE(k).Next Then
        ' this is the first item in the PUB_SLOT_TABLE() array
        ' and is followed by one or more items
        PUB_HASH_TABLE(CODE_VAL) = PUB_SLOT_TABLE(k).Next
    Else
        ' this is the only item in the PUB_SLOT_TABLE() array
        ' for this hash code
        PUB_HASH_TABLE(CODE_VAL) = 0
    End If
    
    ' put the element back in the free list
    PUB_SLOT_TABLE(k).Next = PUB_FREE_INDEX
    PUB_FREE_INDEX = k
    ' we have deleted an item
    PUB_COUNTER_VAL = PUB_COUNTER_VAL - 1
    
End Sub

' remove all items from the hash table

Sub RemoveAll()
    SetSize PUB_HASH_SIZE_VAL, PUB_LIST_SIZE_VAL, PUB_CHUNK_SIZE_VAL
End Sub

' the number of items in the hash table

Property Get COUNT() As Long
    COUNT = PUB_COUNTER_VAL
End Property

' the array of all keys
' (VB5 users: convert return type to Variant)

Property Get Keys() 'As Variant()
    Dim i As Long
    Dim j As Long
    Dim k As Long
    
    Dim DATA_ARR As Variant 'Dim DATA_ARR() As Variant
    
    ReDim DATA_ARR(0 To PUB_COUNTER_VAL - 1)
    
    For i = 0 To PUB_HASH_SIZE_VAL - 1 ' take the pointer from the hash table
        k = PUB_HASH_TABLE(i) ' walk the PUB_SLOT_TABLE() array
        Do While k
            DATA_ARR(j) = PUB_SLOT_TABLE(k).Key
            j = j + 1
            k = PUB_SLOT_TABLE(k).Next
        Loop
    Next i
        
    ' assign to the result
    Keys = DATA_ARR 'DATA_ARR()
End Property

' the array of all values
' (VB5 users: convert return type to Variant)

Property Get Values() 'As Variant()
    Dim i As Long
    Dim j As Long
    Dim k As Long
    Dim DATA_ARR As Variant 'Dim DATA_ARR() As Variant
    
    ReDim DATA_ARR(0 To PUB_COUNTER_VAL - 1)
    
    For i = 0 To PUB_HASH_SIZE_VAL - 1 ' take the pointer from the hash table
        k = PUB_HASH_TABLE(i) ' walk the PUB_SLOT_TABLE() array
        Do While k
            DATA_ARR(j) = PUB_SLOT_TABLE(k).Data
            j = j + 1
            k = PUB_SLOT_TABLE(k).Next
        Loop
    Next i
        
    ' assign to the result
    Values = DATA_ARR 'DATA_ARR()
End Property

'-----------------------------------------
' Private procedures
'-----------------------------------------

Private Sub Class_Initialize()
    ' initialize the tables at default size
    SetSize DEFAULT_HASH_SIZE_VAL, DEFAULT_LIST_SIZE_VAL, _
            DEFAULT_CHUNK_SIZE_VAL
End Sub

' expand the PUB_SLOT_TABLE() array

Private Sub ExpandSlotTable(ByVal k As Long)
    Dim i As Long
    Dim j As Long
    j = UBound(PUB_SLOT_TABLE) + 1
    
    ReDim Preserve PUB_SLOT_TABLE(0 To UBound(PUB_SLOT_TABLE) + k) _
    As PUB_SLOT_TYPE_OBJ
    ' CREATE_FLAG the linked list of free items
    For i = j To UBound(PUB_SLOT_TABLE)
        PUB_SLOT_TABLE(i).Next = i + 1
    Next i
    ' overwrite the last (wrong) value
    PUB_SLOT_TABLE(UBound(PUB_SLOT_TABLE)).Next = PUB_FREE_INDEX
    ' we now know where to pick the first free item
    PUB_FREE_INDEX = j
End Sub

' return the hash code of a string

Private Function HashCode(KEY_STR As String) As Long
    Dim i As Long
    Dim j As Long 'lastEl
    
    ' copy ansi CODES_ARR into an array of long
    j = (Len(KEY_STR) - 1) \ 3
    ReDim CODES_ARR(0 To j) As Long
    
    For i = 1 To Len(KEY_STR)
        CODES_ARR((i - 1) \ 3) = CLng(CODES_ARR((i - 1) \ 3)) * 256 + _
        Asc(Mid(KEY_STR, i, 1))
    Next
    ' this also converts from Unicode to ANSI
'    CopyMemory CODES_ARR(0), ByVal KEY_STR, Len(KEY_STR)
    
    ' XOR the ANSI CODES_ARR of all characters
    For i = 0 To j
        HashCode = HashCode Xor CODES_ARR(i)
    Next
    
End Function

' get the index where an item is stored or 0 if not found
' if CREATE_FLAG = True the item is created
'
' on exit Create=True only if a slot has been actually created

Private Function GetSlotIndex(ByVal KEY_STR As String, _
Optional CREATE_FLAG As Boolean, _
Optional CODE_VAL As Long, _
Optional LAST_INDEX_VAL As Long) As Long
    Dim k As Long
    
    ' raise error if invalid key
    If Len(KEY_STR) = 0 Then Err.Raise 1001, , "Invalid key"
    
    ' keep case-unsensitiveness into account
    If PUB_IGNORE_CASE_FLAG Then KEY_STR = UCase$(KEY_STR)
    ' get the index in the PUB_HASH_TABLE() array
    CODE_VAL = HashCode(KEY_STR) Mod PUB_HASH_SIZE_VAL
    ' get the pointer to the PUB_SLOT_TABLE() array
    k = PUB_HASH_TABLE(CODE_VAL)
    
    ' exit if there is no item with that hash code
    Do While k
        ' compare key with actual value
        If PUB_SLOT_TABLE(k).Key = KEY_STR Then Exit Do
        ' remember last pointer
        LAST_INDEX_VAL = k
        ' check the next item
        k = PUB_SLOT_TABLE(k).Next
    Loop
    
    ' CREATE_FLAG a new item if not there
    If k = 0 And CREATE_FLAG Then
        k = GetFreeSlot()
        PrepareSlot k, KEY_STR, CODE_VAL, LAST_INDEX_VAL
    Else
        ' signal that no item has been created
        CREATE_FLAG = False
    End If
    ' this is the return value
    GetSlotIndex = k

End Function

' return the first free slot

Private Function GetFreeSlot() As Long
    ' allocate new memory if necessary
    If PUB_FREE_INDEX = 0 Then ExpandSlotTable PUB_CHUNK_SIZE_VAL
    ' use the first slot
    GetFreeSlot = PUB_FREE_INDEX
    ' update the pointer to the first slot
    PUB_FREE_INDEX = PUB_SLOT_TABLE(GetFreeSlot).Next
    ' signal this as the end of the linked list
    PUB_SLOT_TABLE(GetFreeSlot).Next = 0
    ' we have one more item
    PUB_COUNTER_VAL = PUB_COUNTER_VAL + 1
End Function

' assign a key and value to a given slot

Private Sub PrepareSlot(ByVal INDEX_VAL As Long, _
ByVal KEY_STR As String, _
ByVal CODE_VAL As Long, _
ByVal LAST_INDEX_VAL As Long)
    ' assign the key
    ' keep case-sensitiveness into account
    If PUB_IGNORE_CASE_FLAG Then KEY_STR = UCase$(KEY_STR)
    PUB_SLOT_TABLE(INDEX_VAL).Key = KEY_STR
    
    If LAST_INDEX_VAL Then
        ' this is the successor of another slot
        PUB_SLOT_TABLE(LAST_INDEX_VAL).Next = INDEX_VAL
    Else
        ' this is the first slot for a given hash code
        PUB_HASH_TABLE(CODE_VAL) = INDEX_VAL
    End If
End Sub
