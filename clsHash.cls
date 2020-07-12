VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsHash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'--------------------------------------------------------------------------------
'--------------------------------------------------------------------------------
Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 0       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.
'--------------------------------------------------------------------------------
'--------------------------------------------------------------------------------

Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDst As Any, pSrc As Any, ByVal ByteLen As Long)

Dim PUB_HASH_NODE_ARR() As Long
'this table is used to store the first and last indexes of each
'hash node
Dim PUB_HASH_INDEX_ARR() As Long
'this table stores the position of the next item in this node
Dim PUB_HASH_KEY_ARR() As String
'this table stores the lookup keys
Dim PUB_HASH_TABLE_SIZE As Long
Dim PUB_HASH_COUNT As Long
'the number of items in the hash table
Dim PUB_HASH_ARR_SIZE As Long
'the dimension set for the hash table

Dim PUB_HASH_CONST_ARR() As Long

Dim PUB_ITEMS_COUNT_ARR() As Long
Dim PUB_MAX_ITEMS As Long


Function KeyFromIndex(i As Long) As String
Key = PUB_HASH_KEY_ARR(i)
End Function

Sub InitialiseHash(i As Long)

Erase PUB_ITEMS_COUNT_ARR
PUB_MAX_ITEMS = 0

PUB_HASH_TABLE_SIZE = i
ReDim PUB_HASH_NODE_ARR(0 To PUB_HASH_TABLE_SIZE, 0 To 2)
Erase PUB_HASH_KEY_ARR, PUB_HASH_INDEX_ARR
PUB_HASH_COUNT = 0
PUB_HASH_ARR_SIZE = 0

Dim k As Long
Dim TEMP_ARR As Variant

TEMP_ARR = Array(7, 37, 91, 67, 37, 11, 53, 19, 43, _
         73, 17, 31, 59, 13, 41, 83, 43, 71, _
         87, 47, 3)

ReDim PUB_HASH_CONST_ARR(0 To 20)
For k = 1 To 20
  PUB_HASH_CONST_ARR(k) = TEMP_ARR(k)
Next k

End Sub

'This function does all the work.
'It can both insert and look up an item in the hashtable
'The function returns the index of the item in the hashtable
'Parameter descriptions:

'KEY_STR is the text we use to look up the item, the "key" text
'ADD_STR_FLAG tells the function whether or not to insert this item if
'it isn't in the table already

Function GetHash(ByVal KEY_STR As String, _
Optional ADD_STR_FLAG As Boolean) As Long

Dim h As Long
Dim x As Long
Dim Y As Long
Dim W As Long

'calculate hash value of string by adding the ASCII value
'of each letter this needs to produce an evenly spread set
'of values, if they are clustered you defeat the
'object, so you may need to change how this is calculated
'If your KEY_STR is a set of numbers you might simply use
'KeyNumber mod X, where X is the number of nodes in the hashtable

'convert string to byte array for speed
h = Len(KEY_STR)
ReDim BYTES_ARR(0 To h) As Byte
CopyMemory BYTES_ARR(1), ByVal KEY_STR, h
'add up ASCII values
x = 0
For W = 1 To h
  Y = Y + 1: If Y > 20 Then Y = Y - 20
  x = x + PUB_HASH_CONST_ARR(Y) * BYTES_ARR(W)
Next W
'take the mod of the answer
'this produces an answer between 0 and PUB_HASH_TABLE_SIZE-1
'this gives us the node we will use for this item
x = x Mod PUB_HASH_TABLE_SIZE

'look up hash in table starting with the first item stored in this node
W = PUB_HASH_NODE_ARR(x, 1) 'this is the index of the first item
If W > 0 Then 'if there is at least one item already there, try to find a match
  If PUB_HASH_KEY_ARR(W) = KEY_STR Then 'the first item is a match
    'do nothing else if we found it first time
  Else 'if we didn't find it, loop through hash table & look for match
    Do While W > 0
      W = PUB_HASH_INDEX_ARR(W)
      'loop until we find a match or run out of items (w=0)
      If W = 0 Or PUB_HASH_KEY_ARR(W) = KEY_STR Then Exit Do
    Loop
  End If
End If

If W = 0 Then 'no match found
  If ADD_STR_FLAG Then 'we have been asked to add the item to the hashtable
    'increment node counter
    PUB_HASH_COUNT = PUB_HASH_COUNT + 1
    'increase array size if necessary
    If PUB_HASH_COUNT > PUB_HASH_ARR_SIZE Then
      PUB_HASH_ARR_SIZE = PUB_HASH_ARR_SIZE + 10
      ReDim Preserve PUB_HASH_INDEX_ARR(0 To PUB_HASH_ARR_SIZE)
      ReDim Preserve PUB_HASH_KEY_ARR(0 To PUB_HASH_ARR_SIZE)
    End If
    
    'store the hash text
    PUB_HASH_KEY_ARR(PUB_HASH_COUNT) = KEY_STR
    
    'if this is the first item added to this node, put it in both
    'elements of the Hash array
    If PUB_HASH_NODE_ARR(x, 1) = 0 Then
      PUB_HASH_NODE_ARR(x, 1) = PUB_HASH_COUNT
      'this holds the first item stored in each node
      PUB_HASH_NODE_ARR(x, 2) = PUB_HASH_COUNT
      'this holds the last item stored in each node
    Else 'otherwise put it in the second element, also set an element
    'in the PUB_HASH_INDEX_ARR array
      PUB_HASH_INDEX_ARR(PUB_HASH_NODE_ARR(x, 2)) = PUB_HASH_COUNT
      'this holds the last item stored in each node
      PUB_HASH_NODE_ARR(x, 2) = PUB_HASH_COUNT
    End If
    If PUB_HASH_COUNT > PUB_MAX_ITEMS Then
        PUB_MAX_ITEMS = PUB_MAX_ITEMS + 50
        ReDim Preserve PUB_ITEMS_COUNT_ARR(0 To PUB_MAX_ITEMS)
    End If
    
    PUB_ITEMS_COUNT_ARR(PUB_HASH_COUNT) = PUB_ITEMS_COUNT_ARR(PUB_HASH_COUNT) + 1
    GetHash = PUB_HASH_COUNT
  End If
Else 'a match was found
  If W > PUB_MAX_ITEMS Then
    PUB_MAX_ITEMS = PUB_MAX_ITEMS + 50
    ReDim Preserve PUB_ITEMS_COUNT_ARR(0 To PUB_MAX_ITEMS)
  End If
  
  PUB_ITEMS_COUNT_ARR(W) = PUB_ITEMS_COUNT_ARR(W) + 1
  GetHash = W
End If
End Function

Function HashCountSizeTables()

Dim i As Long
Dim j As Long
Dim k As Long

Dim HASH_SIZE_TABLE() As String
Dim HASH_COUNT_TABLE() As String

ReDim HASH_SIZE_TABLE(0 To PUB_HASH_TABLE_SIZE, 1 To 4)
HASH_SIZE_TABLE(0, 1) = "Hash Table"
HASH_SIZE_TABLE(0, 2) = "First item #"
HASH_SIZE_TABLE(0, 3) = "Last item #"
HASH_SIZE_TABLE(0, 4) = "No of items"

For j = 1 To PUB_HASH_TABLE_SIZE 'AddName
    HASH_SIZE_TABLE(j, 1) = j
    HASH_SIZE_TABLE(j, 2) = PUB_HASH_NODE_ARR(j - 1, 1)
    HASH_SIZE_TABLE(j, 3) = PUB_HASH_NODE_ARR(j - 1, 2)
    i = 0
    If PUB_HASH_NODE_ARR(j - 1, 1) = 0 Then
      HASH_SIZE_TABLE(j, 4) = 0
    Else
      k = PUB_HASH_NODE_ARR(j - 1, 1)
      i = 1
      Do While PUB_HASH_INDEX_ARR(k) > 0
        k = PUB_HASH_INDEX_ARR(k)
        i = i + 1
      Loop
      HASH_SIZE_TABLE(j, 4) = i
    End If
Next j

ReDim HASH_COUNT_TABLE(0 To PUB_HASH_COUNT, 1 To 4)
HASH_COUNT_TABLE(0, 1) = "Data array"
HASH_COUNT_TABLE(0, 2) = "Name"
HASH_COUNT_TABLE(0, 3) = "Count"
HASH_COUNT_TABLE(0, 4) = "Hash pointer"

For j = 1 To PUB_HASH_COUNT
    HASH_COUNT_TABLE(j, 1) = j
    HASH_COUNT_TABLE(j, 2) = PUB_HASH_KEY_ARR(j)
    HASH_COUNT_TABLE(j, 3) = PUB_ITEMS_COUNT_ARR(j)
    HASH_COUNT_TABLE(j, 4) = PUB_HASH_INDEX_ARR(j)
Next j

HashCountSizeTables = Array(HASH_SIZE_TABLE, HASH_COUNT_TABLE)
End Function
