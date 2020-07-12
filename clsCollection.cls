VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsCollection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'this class allows you to store a number of items under a "key" and
'retrieve them later you can retrieve the keys themselves, which an
'ordinary collection can't do the class can also sort the items on the keyname

Dim PUB_KEYS_OBJ As New Collection 'stores list of initial keys
Dim PUB_ITEMS_OBJ() As New Collection

Dim PUB_KEYS_ARR() As String
Dim PUB_SORT_ARR() As Long
Dim PUB_SORT_FLAG As Boolean
Dim PUB_SIZE_VAL As Long

'runs when class is first used
Private Sub Class_Initialize()

'start this array off
ReDim PUB_KEYS_ARR(0 To 10)

End Sub

'this sets the number of data items to be stored for each key
'if you are going to store data, set this before you do any storage!
Property Let NumberOfDataItems(ByVal tValue As Long)

PUB_SIZE_VAL = tValue
ReDim PUB_ITEMS_OBJ(0 To PUB_SIZE_VAL - 1)

End Property

'this adds the data for a given key
'Function Add(KEY_STR, ParamArray DATA_ARR()) As Long
Function Add(KEY_STR, _
DATA_ARR()) As Long

'check inputs are valid
If KEY_STR = "" Or UBound(DATA_ARR) = 0 Then
  Err.Raise Number:=1004, Description:="Missing key or data"
  Exit Function
End If

Dim k As Long
Dim INDEX_STR As String
'errortrapping is necessary since if you add a key which is already
'present, you'll get an error
On Error Resume Next

'add the item to the collection
'what we actually store is the sequence number in the collection
'this is because a collection can only store one item, and we want to store more
'also, a collection cannot return the keyname, even though it uses it to
'return data
PUB_KEYS_OBJ.Add PUB_KEYS_OBJ.COUNT + 1, CStr(KEY_STR)
'now look up the sequence number using our key
'if this was a new item, this number will be the same as the number of items
'in the collection, whereas if it was an existing item, the number could be
'1, 3, 5 whatever
INDEX_STR = PUB_KEYS_OBJ(CStr(KEY_STR)) 'get it in string form
k = PUB_KEYS_OBJ(CStr(KEY_STR)) 'and as a number

'store the key itself in an array so we can recover it (and sort on it) - the
'collection doesn't permit this so we have to do it external to the collection
'check our key array is large enough, if not, extend it
'only store the key if this is a new key, ie the key index = collection count
If k = PUB_KEYS_OBJ.COUNT Then
  If UBound(PUB_KEYS_ARR) < k Then _
  ReDim Preserve PUB_KEYS_ARR(0 To UBound(PUB_KEYS_ARR) + 50)
  PUB_KEYS_ARR(k) = KEY_STR
End If

'now add the data, as many items as were provided
'we store them in collections too because then they can be absolutely anything
For k = 0 To UBound(DATA_ARR)
  PUB_ITEMS_OBJ(k).Add DATA_ARR(k), INDEX_STR
Next k

'return the key index number in case the user wants to use it
Add = INDEX_STR

'reset the sort variable - ie force a resort if we add data
PUB_SORT_FLAG = False

On Error GoTo 0

End Function

'returns a specific data item for a given key
Function GetData(KEY_STR, _
j As Long)

Dim k As Long

'use error trapping as if the key doesn't exist in the collection,
'there will be an error
On Error Resume Next
k = PUB_KEYS_OBJ(CStr(KEY_STR))
On Error GoTo 0
If k = 0 Then
  Exit Function
End If

GetData = PUB_ITEMS_OBJ(j - 1).Item(k)

End Function

'return PUB_SORT_FLAG data
'the inputs are the sequence number in the collection, and the
'index of the data item required
Function GetSortedData(i As Long, _
j As Long)

Dim k As Long

'sort if necessary
If Not PUB_SORT_FLAG Then Sort

'get the index number corresponding to the PUB_SORT_FLAG index number
On Error Resume Next
k = PUB_SORT_ARR(i)
On Error GoTo 0
If k = 0 Then Exit Function

GetSortedData = PUB_ITEMS_OBJ(j - 1).Item(CStr(k))

End Function

'returns size of collection
Function COUNT()

COUNT = PUB_KEYS_OBJ.COUNT

End Function

'returns key for a given index number
Function Key(ByVal i As Long) As String

If i > 0 And i <= UBound(PUB_KEYS_ARR) Then
  Key = PUB_KEYS_ARR(i)
Else
  Err.Raise Number:=1004, Description:="Key index number of " & i _
  & " is outside array bounds"
End If

End Function

'returns index of key provided
Function index(ByVal KEY_STR As String) As Long

index = 0
On Error Resume Next
index = PUB_KEYS_OBJ(KEY_STR)

End Function

'returns key for a given sorted index number

Function SortedKey(ByVal i As Long) As String

Dim k As Long

'sort if necessary
If Not PUB_SORT_FLAG Then Sort

On Error Resume Next
k = PUB_SORT_ARR(i)
On Error GoTo 0
If k = 0 Then Exit Function

If k > 0 And k <= UBound(PUB_KEYS_ARR) Then
  SortedKey = PUB_KEYS_ARR(k)
Else
  Err.Raise Number:=1004, Description:="Sorted index number of " & _
  i & " refers to item number " & k & " which is outside array bounds"
End If

End Function

'returns original sequence for a given sorted index number

Function SortedIndex(ByVal i As Long)

Dim k As Long

'sort if necessary
If Not PUB_SORT_FLAG Then Sort

On Error Resume Next
k = PUB_SORT_ARR(i)
On Error GoTo 0
If k = 0 Then Exit Function

If k > 0 And k <= UBound(PUB_KEYS_ARR) Then
  SortedIndex = k
Else
  Err.Raise Number:=1004, Description:="PUB_SORT_FLAG index number of " & _
  i & " refers to item number " & k & " which is outside array bounds"
End If

End Function

'uses the PUB_SORT_ARR array to sort the PUB_KEYS_ARR array. Note this
'approach means the items in the PUB_KEYS_ARR array are not reordered,
'but the items in the PUB_SORT_ARR array are reordered
'This approach, sorting a list of sequences rather than the original
'array, is called a linked list
Private Sub Sort()

Dim h  As Long
Dim i As Long
Dim j As Long
Dim k As Long
Dim l As Long

Dim TEMP_ARR As Variant

If PUB_SORT_FLAG = False Then
  ReDim Preserve PUB_KEYS_ARR(0 To PUB_KEYS_OBJ.COUNT)
  ReDim Preserve PUB_SORT_ARR(0 To PUB_KEYS_OBJ.COUNT)
End If

For h = 1 To UBound(PUB_SORT_ARR)
  PUB_SORT_ARR(h) = h
Next h

j = 1
k = UBound(PUB_KEYS_ARR)
i = k - j + 1
Do
  i = Int((i / 13) * 10)
  If i < 1 Then i = 1
  If i = 9 Or i = 10 Then i = 11
  l = 0
  For h = j To k - i
    If PUB_KEYS_ARR(PUB_SORT_ARR(h)) > _
       PUB_KEYS_ARR(PUB_SORT_ARR(h + i)) Then
      TEMP_ARR = PUB_SORT_ARR(h)
      PUB_SORT_ARR(h) = PUB_SORT_ARR(h + i)
      PUB_SORT_ARR(h + i) = TEMP_ARR
      l = 1
    End If
  Next h
Loop Until l + i = 1

PUB_SORT_FLAG = True

End Sub
