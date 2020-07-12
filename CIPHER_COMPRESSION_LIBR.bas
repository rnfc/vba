Attribute VB_Name = "CIPHER_COMPRESSION_LIBR"

'/////////////////////////////////////////////////////////////////////////

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.

Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
    pDst As Any, _
    pSrc As Any, _
    ByVal ByteLen As Long)

Private Declare PtrSafe Function GetTickCount Lib "kernel32" () As Long

Private PUB_SORT_LIST_ARR() As Long
Private PUB_GROUP_LIST_ARR() As Long
Private PUB_GROUP_SIZE_ARR() As Long
Private PUB_NEW_GROUP_LIST_ARR() As Long
Private PUB_NEW_GROUP_SIZE_ARR() As Long
Private PUB_ORIGINAL_SORT_ARR() As Long
Private PUB_GROUP_ORDER_ARR() As Long
Private PUB_N_IN_BYTES_VAL As Long
Private PUB_NGROUPS_VAL As Long
Private PUB_N_NEW_GROUPS_VAL As Long

Private PUB_SLIST_ARR() As Long
Private PUB_NBYTES_VAL As Long

'these settings are for arithmetic coding
'adaptive coding starts after this number of chars
Private Const PUB_ADJUST_FACTOR_START_VAL = 1 '5
'rescale prob table after this number of chars
Private Const PUB_ADJUST_FACTOR_MAX_VAL = 1000
'rescaling factor
Private Const PUB_ADJUST_FACTOR_VAL = 0.5
'weight to add to prob table for each char
Private Const PUB_PROB_ADD_VAL = 30
'initial weight for each char
Private Const PUB_START_P_VAL = 4
Private Const PUB_BWT_CHUNK_VAL = 5000000

'these define the resolution of the coder, as big as possible
Private Const PUB_B32_VAL = 4294967296#
Private Const PUB_B24_VAL = 16777216
'Private Const PUB_B16_VAL = 65536
'-----------------------------------------------------------------------------
'Data Compression
'It uses well known and modern compression algorithms
'notably the Burrows-Wheeler transform, assisted by Bring to Front,
'Run Length Encoding, and Arithmetic coding.
'-----------------------------------------------------------------------------

Public Function WSHEET_EMBED_DATA_FUNC(ByRef DST_RNG As Excel.Range, _
ByVal DATA_FILE_STR As String, _
Optional ByVal COMPRESS_FLAG As Boolean = True, _
Optional ByVal VALIDATE_FLAG As Boolean = True)

Dim i As Long
Dim j As Long
Dim k As Long 'PUB_NBYTES_VAL

Dim A_DATA_TEXT As String
Dim B_DATA_TEXT As String

Dim A_BYTES_ARR() As Byte
Dim B_BYTES_ARR() As Byte

On Error GoTo ERROR_LABEL

'get datafile & compress it if required
If COMPRESS_FLAG Then
  k = READ_FILE_FUNC(DATA_FILE_STR, A_BYTES_ARR())
  'compress it if required
  RLE_FUNC A_BYTES_ARR(), B_BYTES_ARR()
  BWT_FUNC B_BYTES_ARR(), A_BYTES_ARR()
  BTF_FUNC A_BYTES_ARR(), B_BYTES_ARR()
  RLE_FUNC B_BYTES_ARR(), A_BYTES_ARR()
  ARI_FUNC A_BYTES_ARR(), B_BYTES_ARR()
Else
  k = READ_FILE_FUNC(DATA_FILE_STR, B_BYTES_ARR())
End If

'REMAP_FUNC characters that won't go into cell comments
'this adds about 12% to the text length
REMAP_FUNC B_BYTES_ARR(), A_BYTES_ARR()
A_DATA_TEXT = StrConv(A_BYTES_ARR(), vbUnicode)
Erase A_BYTES_ARR()
Erase B_BYTES_ARR()

'embed data in chunks of 32767 characters
j = 1
i = 0
Do
  If j > Len(A_DATA_TEXT) Then Exit Do
  i = i + 1
  DST_RNG.Cells(i, 1).AddComment Mid$(A_DATA_TEXT, j, 32767)
  j = j + 32767
Loop

WSHEET_EMBED_DATA_FUNC = False

'check it worked, if requested
If VALIDATE_FLAG Then
  B_DATA_TEXT = WSHEET_RECOVER_DATA_FUNC(DST_RNG.Parent, COMPRESS_FLAG)
  'return success flag
  k = READ_FILE_FUNC(DATA_FILE_STR, A_BYTES_ARR())
  A_DATA_TEXT = StrConv(A_BYTES_ARR(), vbUnicode)
  WSHEET_EMBED_DATA_FUNC = (A_DATA_TEXT = B_DATA_TEXT)
'if not testing, just return success flag
Else
  WSHEET_EMBED_DATA_FUNC = True
End If

Exit Function
ERROR_LABEL:
WSHEET_EMBED_DATA_FUNC = False
End Function

'It shows how to embed binary data efficiently in an Excel file
'It may be used freely, with attribution

Public Function WSHEET_RECOVER_DATA_FUNC( _
ByRef SRC_WSHEET As Excel.Worksheet, _
Optional COMPRESS_FLAG As Boolean = True) As String

Dim DATA_TEXT As String

Dim A_BYTES_ARR() As Byte
Dim B_BYTES_ARR() As Byte

Dim TMP_COMMENT As Excel.Comment

On Error GoTo ERROR_LABEL

For Each TMP_COMMENT In SRC_WSHEET.comments
  DATA_TEXT = DATA_TEXT & TMP_COMMENT.Text
Next TMP_COMMENT

'convert to byte array and REMAP_FUNC
ReDim A_BYTES_ARR(Len(DATA_TEXT))
CopyMemory A_BYTES_ARR(1), ByVal DATA_TEXT, Len(DATA_TEXT)
DEMAP_FUNC A_BYTES_ARR(), B_BYTES_ARR()

'decompress if necessary
If COMPRESS_FLAG = True Then 'Decompress
  ARI_REVERSE_FUNC B_BYTES_ARR(), A_BYTES_ARR()
  RLE_REVERSE_FUNC A_BYTES_ARR(), B_BYTES_ARR()
  BTF_REVERSE_FUNC B_BYTES_ARR(), A_BYTES_ARR()
  BWT_REVERSE_FUNC A_BYTES_ARR(), B_BYTES_ARR()
  RLE_REVERSE_FUNC B_BYTES_ARR(), A_BYTES_ARR()
  WSHEET_RECOVER_DATA_FUNC = StrConv(A_BYTES_ARR(), vbUnicode)
Else
  WSHEET_RECOVER_DATA_FUNC = StrConv(B_BYTES_ARR(), vbUnicode)
End If
Erase A_BYTES_ARR(), B_BYTES_ARR()

Exit Function
ERROR_LABEL:
WSHEET_RECOVER_DATA_FUNC = Err.number
End Function


'this sub compresses data
Public Function COMPRESS_FILE_FUNC( _
ByVal IN_TEXT_FILE_STR As String, _
ByVal OUT_TEXT_FILE_STR As String, _
Optional ByVal VERIFY_FLAG As Boolean = True)

Dim i As Long
Dim NSIZE As Long

Dim A_BYTES_ARR() As Byte
Dim B_BYTES_ARR() As Byte

On Error GoTo ERROR_LABEL

'read data into byte array
READ_FILE_FUNC IN_TEXT_FILE_STR, A_BYTES_ARR()

'compress in 5 steps
RLE_FUNC A_BYTES_ARR(), B_BYTES_ARR()
BWT_FUNC B_BYTES_ARR(), A_BYTES_ARR()
BTF_FUNC A_BYTES_ARR(), B_BYTES_ARR()
RLE_FUNC B_BYTES_ARR(), A_BYTES_ARR()
ARI_FUNC A_BYTES_ARR(), B_BYTES_ARR()

'write compressed output
WRITE_FILE_FUNC OUT_TEXT_FILE_STR, B_BYTES_ARR()

'VERIFY_FLAG result if asked
If VERIFY_FLAG Then
  
  'read in file just written
  Erase B_BYTES_ARR
  READ_FILE_FUNC OUT_TEXT_FILE_STR, B_BYTES_ARR()
  
  'decompress in the same order
  ARI_REVERSE_FUNC B_BYTES_ARR(), A_BYTES_ARR()
  RLE_REVERSE_FUNC A_BYTES_ARR(), B_BYTES_ARR()
  BTF_REVERSE_FUNC B_BYTES_ARR(), A_BYTES_ARR()
  BWT_REVERSE_FUNC A_BYTES_ARR(), B_BYTES_ARR()
  RLE_REVERSE_FUNC B_BYTES_ARR(), A_BYTES_ARR()
  
  'read in the original file again
  READ_FILE_FUNC IN_TEXT_FILE_STR, B_BYTES_ARR()
  
  'compare byte by byte
  'exit loop if faulty
  NSIZE = UBound(B_BYTES_ARR)
  For i = 1 To NSIZE
    If A_BYTES_ARR(i) <> B_BYTES_ARR(i) Then Exit For
  Next i
  
  'return True if OK
  If i > NSIZE Then COMPRESS_FILE_FUNC = True

Else 'otherwise return True if compressed without code errors
  COMPRESS_FILE_FUNC = True
End If


Exit Function
ERROR_LABEL:
COMPRESS_FILE_FUNC = False
End Function


'this sub compresses data
Public Function DECOMPRESS_FILE_FUNC( _
ByVal IN_TEXT_FILE_STR As String, _
ByVal OUT_TEXT_FILE_STR As String)

Dim A_BYTES_ARR() As Byte
Dim B_BYTES_ARR() As Byte

On Error GoTo ERROR_LABEL

'read data into byte array
READ_FILE_FUNC IN_TEXT_FILE_STR, B_BYTES_ARR()
  
'decompress
ARI_REVERSE_FUNC B_BYTES_ARR(), A_BYTES_ARR()
RLE_REVERSE_FUNC A_BYTES_ARR(), B_BYTES_ARR()
BTF_REVERSE_FUNC B_BYTES_ARR(), A_BYTES_ARR()
BWT_REVERSE_FUNC A_BYTES_ARR(), B_BYTES_ARR()
RLE_REVERSE_FUNC B_BYTES_ARR(), A_BYTES_ARR()

WRITE_FILE_FUNC OUT_TEXT_FILE_STR, A_BYTES_ARR()

DECOMPRESS_FILE_FUNC = True

Exit Function
ERROR_LABEL:
DECOMPRESS_FILE_FUNC = False
End Function


'Zero-order Arithmetic coder

Private Function ARI_FUNC(ByRef IN_BYTES_ARR() As Byte, _
ByRef OUT_BYTES_ARR() As Byte)

Dim cumStep As Long
Dim cumDiv As Long

Dim i As Long
Dim j As Long


Dim bNo As Long
Dim l As Double
Dim r As Double
Dim RR As Double
Dim W As Long
Dim U As Long
Dim V As Single

Dim NSIZE As Long
Dim nonFF As Byte
Dim nFF As Long
Dim LastProbUpdate As Long
Dim BEEN_FLAG As Boolean
Dim TEMP_BYTE As Byte

Dim P_ARR(0 To 255) As Long
Dim C_ARR(-1 To 255) As Long

NSIZE = UBound(IN_BYTES_ARR)
ReDim OUT_BYTES_ARR(NSIZE * 2 + 3)

P_ARR(0) = PUB_START_P_VAL
C_ARR(0) = P_ARR(0)
For i = 1 To 255
  P_ARR(i) = PUB_START_P_VAL
  C_ARR(i) = C_ARR(i - 1) + P_ARR(i)
Next i

cumStep = PUB_ADJUST_FACTOR_START_VAL
cumDiv = PUB_ADJUST_FACTOR_MAX_VAL
LastProbUpdate = PUB_ADJUST_FACTOR_START_VAL

W = 3 'number of characters output, leave the first 3 to count the total bytes
l = 0 'left hand value
r = PUB_B32_VAL

For i = 1 To NSIZE
  
  bNo = IN_BYTES_ARR(i)
  RR = Int(r / C_ARR(255))
  l = l + RR * C_ARR(bNo - 1)
  
  r = RR * (C_ARR(bNo) - C_ARR(bNo - 1))
    
  If l >= PUB_B32_VAL Then
    l = l - PUB_B32_VAL
    nonFF = nonFF + 1
    For U = 1 To nFF
      W = W + 1
      OUT_BYTES_ARR(W) = nonFF
      nonFF = 0
    Next U
    nFF = 0
  End If
  
  Do While r <= PUB_B24_VAL
    TEMP_BYTE = Int(l / PUB_B24_VAL)
  
    If Not BEEN_FLAG Then
      nonFF = TEMP_BYTE
      nFF = 0
      BEEN_FLAG = True
    ElseIf TEMP_BYTE = 255 Then
      nFF = nFF + 1
    Else
      W = W + 1
      OUT_BYTES_ARR(W) = nonFF
      For U = 1 To nFF
        W = W + 1
        OUT_BYTES_ARR(W) = 255
      Next U
      nFF = 0
      nonFF = TEMP_BYTE
    End If
    
    l = (l - CDbl(TEMP_BYTE) * PUB_B24_VAL) * 256
    r = r * 256
    
  Loop

  'update frequency count
  P_ARR(bNo) = P_ARR(bNo) + PUB_PROB_ADD_VAL
  'update cumulative stats if we reach the next step
  If i = cumStep Then

    'every 1000 or so steps, divide all the values by 2 to "age" them
    'and give more weight to subsequent items
    
    If cumStep > cumDiv Then 'age the stats
      V = 1 / PUB_ADJUST_FACTOR_VAL
        If P_ARR(0) < V Then P_ARR(0) = 1 Else P_ARR(0) = P_ARR(0) / V
        C_ARR(0) = P_ARR(0)
      For j = 1 To 255
        If P_ARR(j) < V Then P_ARR(j) = 1 Else P_ARR(j) = P_ARR(j) / V
        C_ARR(j) = C_ARR(j - 1) + P_ARR(j)
      Next j
      cumDiv = cumDiv + PUB_ADJUST_FACTOR_MAX_VAL
    Else 'don't age the stats
      C_ARR(0) = P_ARR(0)
      For j = 1 To 255
        C_ARR(j) = C_ARR(j - 1) + P_ARR(j)
      Next j
    End If
    
    'increase the size of the next step just a little
    cumStep = cumStep + LastProbUpdate
    If LastProbUpdate < 1000 Then LastProbUpdate = LastProbUpdate + 5
    
  End If

Next i

Do While l > 0
  TEMP_BYTE = Int(l / PUB_B24_VAL)
  If TEMP_BYTE = 255 Then
    nFF = nFF + 1
  Else
    W = W + 1
    OUT_BYTES_ARR(W) = nonFF
    For U = 1 To nFF
      W = W + 1
      OUT_BYTES_ARR(W) = 255
    Next U
    nFF = 0
    nonFF = TEMP_BYTE
  End If
  l = (l - CDbl(TEMP_BYTE) * PUB_B24_VAL) * 256
Loop

If nonFF > 0 Then
  W = W + 1
  OUT_BYTES_ARR(W) = nonFF
End If
For U = 1 To nFF
  W = W + 1
  OUT_BYTES_ARR(W) = 255
Next U

If W < 6 Then W = 6
ReDim Preserve OUT_BYTES_ARR(W)

U = Int(NSIZE / 256)
OUT_BYTES_ARR(3) = NSIZE - U * 256
NSIZE = U
U = Int(NSIZE / 256)
OUT_BYTES_ARR(2) = NSIZE - U * 256
OUT_BYTES_ARR(1) = U

End Function

Private Function ARI_REVERSE_FUNC(ByRef IN_BYTES_ARR() As Byte, _
ByRef OUT_BYTES_ARR() As Byte)

Dim i As Long
Dim j As Long

Dim bNo As Long
Dim r As Double
Dim W As Long
Dim U As Long
Dim V As Single
Dim nW As Long
Dim NSIZE As Long
Dim D As Double
Dim RR As Double
Dim LastProbUpdate As Long
Dim PUB_NBYTES_VAL As Long

Dim cumDiv As Long
Dim cumStep As Long

Dim P_ARR(0 To 255) As Long
Dim C_ARR(-1 To 255) As Long

LastProbUpdate = PUB_ADJUST_FACTOR_START_VAL

PUB_NBYTES_VAL = (CLng(IN_BYTES_ARR(1)) * 256 + _
          IN_BYTES_ARR(2)) * 256 + IN_BYTES_ARR(3)

NSIZE = UBound(IN_BYTES_ARR)
nW = NSIZE * 2
ReDim OUT_BYTES_ARR(PUB_NBYTES_VAL)

'initialise cumulative probability array
P_ARR(0) = PUB_START_P_VAL
C_ARR(0) = P_ARR(0)
For i = 1 To 255
  P_ARR(i) = PUB_START_P_VAL
  C_ARR(i) = C_ARR(i - 1) + P_ARR(i)
Next i
cumStep = PUB_ADJUST_FACTOR_START_VAL
cumDiv = PUB_ADJUST_FACTOR_MAX_VAL

'read in the first 3 bytes
D = ((CDbl(IN_BYTES_ARR(4)) * 256 + IN_BYTES_ARR(5)) * 256 + _
           IN_BYTES_ARR(6)) * 256 + IN_BYTES_ARR(7)
i = 7
r = PUB_B32_VAL

For W = 1 To PUB_NBYTES_VAL
  
  RR = Int(r / C_ARR(255))
  U = Int(D / RR)
  
  For bNo = 0 To 255
    If U < C_ARR(bNo) Then Exit For
  Next bNo
  If bNo > 255 Then If U = C_ARR(255) Then bNo = 255 Else bNo = 255: Stop
    
  D = D - Int(RR * C_ARR(bNo - 1))
  
  r = RR * (C_ARR(bNo) - C_ARR(bNo - 1))
    
  Do While r <= PUB_B24_VAL
    r = r * 256
    i = i + 1
    If i <= NSIZE Then
      D = D * 256 + IN_BYTES_ARR(i)
    Else
      D = D * 256
    End If
  Loop
  
  OUT_BYTES_ARR(W) = bNo
  
  'update frequency count
  P_ARR(bNo) = P_ARR(bNo) + PUB_PROB_ADD_VAL
  
  'update cumulative stats if we reach the next step
  If W = cumStep Then

    'every 1000 or so steps, divide all the values by 2 to "age"
    'them and give more weight to subsequent items
    
    If cumStep > cumDiv Then 'age the stats
      V = 1 / PUB_ADJUST_FACTOR_VAL
        If P_ARR(0) < V Then P_ARR(0) = 1 Else P_ARR(0) = P_ARR(0) / V
        C_ARR(0) = P_ARR(0)
      For j = 1 To 255
        If P_ARR(j) < V Then P_ARR(j) = 1 Else P_ARR(j) = P_ARR(j) / V
        C_ARR(j) = C_ARR(j - 1) + P_ARR(j)
      Next j
      cumDiv = cumDiv + PUB_ADJUST_FACTOR_MAX_VAL
    Else 'don't age the stats
      C_ARR(0) = P_ARR(0)
      For j = 1 To 255
        C_ARR(j) = C_ARR(j - 1) + P_ARR(j)
      Next j
    End If
    
    'increase the size of the next step just a little
    cumStep = cumStep + LastProbUpdate
    If LastProbUpdate < 1000 Then LastProbUpdate = LastProbUpdate + 5
    
  End If
  
Next W

ReDim Preserve OUT_BYTES_ARR(PUB_NBYTES_VAL)

End Function


'Bring to front
Private Function BTF_FUNC(ByRef IN_BYTES_ARR() As Byte, _
ByRef OUT_BYTES_ARR() As Byte)

Dim i As Long
Dim j As Long
Dim tPosString As String '* 256
Dim tChr As String * 1
Dim tPos(0 To 255) As Long
Dim NSIZE As Long

NSIZE = UBound(IN_BYTES_ARR)
ReDim OUT_BYTES_ARR(NSIZE)

For i = 0 To 255
  tPosString = tPosString & Chr$(i)
Next i

For i = 1 To NSIZE
  tChr = Chr$(IN_BYTES_ARR(i))
  OUT_BYTES_ARR(i) = InStr(tPosString, tChr) - 1
  CopyMemory ByVal StrPtr(tPosString) + 2, _
             ByVal StrPtr(tPosString), OUT_BYTES_ARR(i) * 2
  CopyMemory ByVal StrPtr(tPosString), ByVal StrPtr(tChr), 2
Next i

Erase tPos
For i = 1 To NSIZE
  j = OUT_BYTES_ARR(i)
  tPos(j) = tPos(j) + 1
Next i

End Function

Private Function BTF_REVERSE_FUNC(ByRef IN_BYTES_ARR() As Byte, _
ByRef OUT_BYTES_ARR() As Byte)

Dim i As Long
Dim NSIZE As Long
Dim tChr As String * 1
'Dim tPos(0 To 255) As Long
Dim tPosString As String '* 256

NSIZE = UBound(IN_BYTES_ARR)
ReDim OUT_BYTES_ARR(NSIZE)

For i = 0 To 255
  tPosString = tPosString & Chr$(i)
Next i

For i = 1 To NSIZE
  tChr = Mid$(tPosString, IN_BYTES_ARR(i) + 1, 1)
  OUT_BYTES_ARR(i) = Asc(tChr)
  CopyMemory ByVal StrPtr(tPosString) + 2, _
             ByVal StrPtr(tPosString), IN_BYTES_ARR(i) * 2
  CopyMemory ByVal StrPtr(tPosString), ByVal StrPtr(tChr), 2
Next i

End Function

Private Function BWT_FUNC(ByRef IN_BYTES_ARR() As Byte, _
ByRef OUT_BYTES_ARR() As Byte)

Dim U1_VAL As Long
Dim U2_VAL As Long
Dim U3_VAL As Long
Dim U4_VAL As Long

Dim NSIZE As Long

Dim A_BYTES_ARR() As Byte
Dim B_BYTES_ARR() As Byte

On Error GoTo ERROR_LABEL

BWT_FUNC = False
NSIZE = UBound(IN_BYTES_ARR)

If NSIZE > PUB_BWT_CHUNK_VAL - 4 Then
  Do While U2_VAL < NSIZE
    U2_VAL = U1_VAL + PUB_BWT_CHUNK_VAL - 4
    If U2_VAL > NSIZE Then U2_VAL = NSIZE
    ReDim A_BYTES_ARR(U2_VAL - U1_VAL)
    CopyMemory ByVal VarPtr(A_BYTES_ARR(1)), _
               ByVal VarPtr(IN_BYTES_ARR(U1_VAL + 1)), (U2_VAL - U1_VAL)
    SORT_BYTE_ARR_FUNC A_BYTES_ARR(), B_BYTES_ARR()
    U4_VAL = UBound(B_BYTES_ARR)
    If U3_VAL > 0 Then
      ReDim Preserve OUT_BYTES_ARR(UBound(OUT_BYTES_ARR) + U4_VAL)
    Else
      ReDim OUT_BYTES_ARR(U4_VAL)
    End If
    CopyMemory OUT_BYTES_ARR(U3_VAL + 1), B_BYTES_ARR(1), U4_VAL
    U1_VAL = U2_VAL
    U3_VAL = U3_VAL + U4_VAL
  Loop
Else
  SORT_BYTE_ARR_FUNC IN_BYTES_ARR(), OUT_BYTES_ARR()
End If

BWT_FUNC = True

Exit Function
ERROR_LABEL:
BWT_FUNC = False
End Function


Private Function BWT_REVERSE_FUNC(ByRef IN_BYTES_ARR() As Byte, _
ByRef OUT_BYTES_ARR() As Byte)

Dim U1_VAL As Long
Dim U2_VAL As Long
Dim U3_VAL As Long
Dim U4_VAL As Long
Dim NSIZE As Long

On Error GoTo ERROR_LABEL

BWT_REVERSE_FUNC = False

NSIZE = UBound(IN_BYTES_ARR)

If NSIZE > PUB_BWT_CHUNK_VAL Then
  Dim A_BYTES_ARR() As Byte, B_BYTES_ARR() As Byte
  Do While U2_VAL < NSIZE
    U2_VAL = U1_VAL + PUB_BWT_CHUNK_VAL
    If U2_VAL > NSIZE Then U2_VAL = NSIZE
    ReDim A_BYTES_ARR(U2_VAL - U1_VAL)
    CopyMemory ByVal VarPtr(A_BYTES_ARR(1)), _
                     ByVal VarPtr(IN_BYTES_ARR(U1_VAL + 1)), (U2_VAL - U1_VAL)
    BWT_DECODE_FUNC A_BYTES_ARR(), B_BYTES_ARR()
    U4_VAL = UBound(B_BYTES_ARR)
    If U3_VAL > 0 Then
      ReDim Preserve OUT_BYTES_ARR(UBound(OUT_BYTES_ARR) + U4_VAL)
    Else
      ReDim OUT_BYTES_ARR(U4_VAL)
    End If
    CopyMemory OUT_BYTES_ARR(U3_VAL + 1), B_BYTES_ARR(1), U4_VAL
    U1_VAL = U2_VAL
    U3_VAL = U3_VAL + U4_VAL
  Loop
Else
  BWT_DECODE_FUNC IN_BYTES_ARR(), OUT_BYTES_ARR()
End If

BWT_REVERSE_FUNC = True

Exit Function
ERROR_LABEL:
BWT_REVERSE_FUNC = False
End Function


Private Function BWT_DECODE_FUNC(ByRef IN_BYTES_ARR() As Byte, _
ByRef OUT_BYTES_ARR() As Byte)

Dim i As Long
Dim j As Long
Dim k As Long

Dim NSIZE As Long
Dim TEMP_ARR(0 To 255) As Long

On Error GoTo ERROR_LABEL

BWT_DECODE_FUNC = False

NSIZE = UBound(IN_BYTES_ARR)

ReDim OUT_BYTES_ARR(NSIZE - 4)
ReDim PUB_SLIST_ARR(NSIZE)

'get starting item, stored in first 4 bytes
k = 0
j = 1
For i = 1 To 3
  k = k + IN_BYTES_ARR(i) * j
  j = j * 256
Next i
k = k + IN_BYTES_ARR(i) * j

For i = 5 To NSIZE
  j = IN_BYTES_ARR(i)
  TEMP_ARR(j) = TEMP_ARR(j) + 1
Next i

For i = 1 To 255
  TEMP_ARR(i) = TEMP_ARR(i) + TEMP_ARR(i - 1)
Next i

For i = NSIZE To 5 Step -1
  j = IN_BYTES_ARR(i)
  PUB_SLIST_ARR(TEMP_ARR(j)) = i - 4
  TEMP_ARR(j) = TEMP_ARR(j) - 1
Next i

j = k
For i = 1 To NSIZE - 4
  OUT_BYTES_ARR(i) = IN_BYTES_ARR(PUB_SLIST_ARR(j) + 4)
  j = PUB_SLIST_ARR(j)
Next i

BWT_DECODE_FUNC = True

Exit Function
ERROR_LABEL:
BWT_DECODE_FUNC = False
End Function

Private Function READ_FILE_FUNC(ByRef TEXT_FILE_STR As String, _
ByRef OUT_BYTES_ARR() As Byte) As Long
    
Dim k As Long
            
On Error GoTo ERROR_LABEL
    
k = FreeFile 'Open file
    
Open TEXT_FILE_STR For Binary Access Read As #k
'Size the array to hold the file contents
ReDim OUT_BYTES_ARR(1 To LOF(k))
    
Get #k, , OUT_BYTES_ARR
Close #k
       
READ_FILE_FUNC = UBound(OUT_BYTES_ARR)
       
Exit Function
ERROR_LABEL:
READ_FILE_FUNC = Err.number
End Function

Private Function WRITE_FILE_FUNC(ByRef TEXT_FILE_STR As String, _
ByRef IN_BYTES_ARR() As Byte) As Long
    
Dim k As Long
    
On Error GoTo ERROR_LABEL
    
k = FreeFile
    
Open TEXT_FILE_STR For Output As #k
Close
    
Open TEXT_FILE_STR For Binary Access Write As #k
    'Size the array to hold the file contents
ReDim OUT_BYTES_ARR(1 To UBound(IN_BYTES_ARR))
    
Put #k, , IN_BYTES_ARR
Close #k
       
WRITE_FILE_FUNC = UBound(IN_BYTES_ARR)

Exit Function
ERROR_LABEL:
WRITE_FILE_FUNC = Err.number
End Function

Private Function RLE_FUNC(ByRef IN_BYTES_ARR() As Byte, _
ByRef OUT_BYTES_ARR() As Byte) As Long

Dim i As Long
Dim j As Long
Dim m As Long
Dim c As Long
Dim NSIZE As Long

NSIZE = UBound(IN_BYTES_ARR)

ReDim OUT_BYTES_ARR(NSIZE * 1.5) As Byte

j = 0

For i = 1 To NSIZE

  j = j + 1
  OUT_BYTES_ARR(j) = IN_BYTES_ARR(i)
  
  If IN_BYTES_ARR(i) = c Then
    
    m = -1
    
    Do
      m = m + 1
      If i = NSIZE Then
        Exit Do
      End If
      i = i + 1
      If m = 255 Then
        Exit Do
      End If
    Loop While IN_BYTES_ARR(i) = c
    
    If m >= 0 Then
      j = j + 1
      OUT_BYTES_ARR(j) = m
    End If
    
    If IN_BYTES_ARR(i) <> c Or m = 255 Then
      j = j + 1
      OUT_BYTES_ARR(j) = IN_BYTES_ARR(i)
    End If
    
  End If
  
  c = IN_BYTES_ARR(i)

Next i

RLE_FUNC = j
ReDim Preserve OUT_BYTES_ARR(j)

'For i = 1 To j
'  OUT_BYTES_ARR(i) = tmp(i)
'Next i

End Function

Private Function RLE_REVERSE_FUNC(ByRef IN_BYTES_ARR() As Byte, _
ByRef OUT_BYTES_ARR() As Byte) As Long

Dim i As Long
Dim j As Long
Dim m As Long
Dim k As Long
Dim c As Long
Dim U As Long
Dim NSIZE As Long

NSIZE = UBound(IN_BYTES_ARR)
k = NSIZE * 3
ReDim tmp(k) As Byte

j = 0

For i = 1 To NSIZE

  j = j + 1
  If j > k Then
    k = k + NSIZE
    ReDim Preserve tmp(k)
  End If
  tmp(j) = IN_BYTES_ARR(i)
  
  If IN_BYTES_ARR(i) = c Then
    'If i > 55 Then Stop
    i = i + 1
    If i > NSIZE Then Exit For
    U = IN_BYTES_ARR(i)
    For m = 1 To U
      j = j + 1
      If j > k Then
        k = k + NSIZE
        ReDim Preserve tmp(k)
      End If
      tmp(j) = c
    Next m
    c = -1
  Else
    c = IN_BYTES_ARR(i)
  End If

Next i

RLE_REVERSE_FUNC = j
ReDim OUT_BYTES_ARR(j)

For i = 1 To j
  OUT_BYTES_ARR(i) = tmp(i)
Next i

End Function



Private Function SORT_BYTE_ARR_FUNC(ByRef IN_BYTES_ARR() As Byte, _
ByRef OUT_BYTES_ARR() As Byte)

Dim i As Long
Dim j As Long
Dim k As Long
Dim m As Long
Dim U As Long
Dim V As Long
Dim W As Long

Dim aSize As Long
Dim StartByte As Long
Dim StartGroup As Long
Dim nCurrGroups As Long
Dim TEMP_MATRIX(0 To 255, 0 To 255) As Long

On Error GoTo ERROR_LABEL

SORT_BYTE_ARR_FUNC = False

PUB_N_IN_BYTES_VAL = UBound(IN_BYTES_ARR)

ReDim PUB_ORIGINAL_SORT_ARR(0 To PUB_N_IN_BYTES_VAL)
ReDim PUB_GROUP_ORDER_ARR(PUB_N_IN_BYTES_VAL) As Long
ReDim PUB_SORT_LIST_ARR(PUB_N_IN_BYTES_VAL)
U = PUB_N_IN_BYTES_VAL '* 5 + 10000
ReDim PUB_GROUP_LIST_ARR(0 To U), PUB_GROUP_SIZE_ARR(0 To U)

Erase TEMP_MATRIX
For i = 1 To PUB_N_IN_BYTES_VAL - 1
  U = IN_BYTES_ARR(i)
  V = IN_BYTES_ARR(i + 1)
  TEMP_MATRIX(U, V) = TEMP_MATRIX(U, V) + 1
Next i
TEMP_MATRIX(IN_BYTES_ARR(PUB_N_IN_BYTES_VAL), IN_BYTES_ARR(1)) = _
    TEMP_MATRIX(IN_BYTES_ARR(PUB_N_IN_BYTES_VAL), IN_BYTES_ARR(1)) + 1

m = 1
PUB_NGROUPS_VAL = 0
PUB_GROUP_LIST_ARR(0) = 0

For i = 0 To 255
  For j = 0 To 255
    If TEMP_MATRIX(i, j) > 0 Then
      PUB_NGROUPS_VAL = PUB_NGROUPS_VAL + 1
      U = m + TEMP_MATRIX(i, j) - 1
      For k = m To U
        PUB_GROUP_ORDER_ARR(k) = m
      Next k
      PUB_GROUP_LIST_ARR(PUB_NGROUPS_VAL) = m
      PUB_GROUP_SIZE_ARR(PUB_NGROUPS_VAL - 1) = _
        m - PUB_GROUP_LIST_ARR(PUB_NGROUPS_VAL - 1)
      m = m + TEMP_MATRIX(i, j)
      TEMP_MATRIX(i, j) = m - 1
    End If
  Next j
Next i
If m > PUB_GROUP_LIST_ARR(PUB_NGROUPS_VAL) Then
  PUB_GROUP_SIZE_ARR(PUB_NGROUPS_VAL) = _
    m - PUB_GROUP_LIST_ARR(PUB_NGROUPS_VAL)
End If

PUB_SORT_LIST_ARR(TEMP_MATRIX(IN_BYTES_ARR(PUB_N_IN_BYTES_VAL), _
IN_BYTES_ARR(1))) = PUB_N_IN_BYTES_VAL

PUB_ORIGINAL_SORT_ARR(PUB_N_IN_BYTES_VAL) = _
    PUB_GROUP_ORDER_ARR(TEMP_MATRIX(IN_BYTES_ARR(PUB_N_IN_BYTES_VAL), _
    IN_BYTES_ARR(1)))

TEMP_MATRIX(IN_BYTES_ARR(PUB_N_IN_BYTES_VAL), IN_BYTES_ARR(1)) = _
    TEMP_MATRIX(IN_BYTES_ARR(PUB_N_IN_BYTES_VAL), IN_BYTES_ARR(1)) - 1

For i = PUB_N_IN_BYTES_VAL - 1 To 1 Step -1
  U = IN_BYTES_ARR(i)
  V = IN_BYTES_ARR(i + 1)
  W = TEMP_MATRIX(U, V)
  PUB_SORT_LIST_ARR(W) = i
  PUB_ORIGINAL_SORT_ARR(i) = PUB_GROUP_ORDER_ARR(W)
  W = W - 1
  TEMP_MATRIX(U, V) = W
Next i

'make copy of original sort group list so we can update it
aSize = UBound(PUB_ORIGINAL_SORT_ARR) - LBound(PUB_ORIGINAL_SORT_ARR) + 1
ReDim PUB_GROUP_ORDER_ARR(0 To PUB_N_IN_BYTES_VAL)
CopyMemory ByVal VarPtr(PUB_GROUP_ORDER_ARR(0)), _
           ByVal VarPtr(PUB_ORIGINAL_SORT_ARR(0)), _
           LenB(PUB_ORIGINAL_SORT_ARR(0)) * aSize

W = 1
Do
  W = W * 2
   
  PUB_N_NEW_GROUPS_VAL = 0
  ReDim PUB_NEW_GROUP_LIST_ARR(0 To PUB_N_IN_BYTES_VAL), _
        PUB_NEW_GROUP_SIZE_ARR(0 To PUB_N_IN_BYTES_VAL)

  StartGroup = 1 'nCurrGroups + 1
  nCurrGroups = PUB_N_NEW_GROUPS_VAL
  
  For i = 1 To PUB_NGROUPS_VAL
    If PUB_GROUP_SIZE_ARR(i) > 1 Then
      SORT_GROUP_FUNC i, W
    End If
  Next i
  
  'replace sort group list with new version
  CopyMemory ByVal VarPtr(PUB_ORIGINAL_SORT_ARR(0)), _
             ByVal VarPtr(PUB_GROUP_ORDER_ARR(0)), _
             LenB(PUB_GROUP_ORDER_ARR(0)) * aSize
  
  If PUB_N_NEW_GROUPS_VAL = 0 Then Exit Do
  
  CopyMemory ByVal VarPtr(PUB_GROUP_LIST_ARR(0)), _
             ByVal VarPtr(PUB_NEW_GROUP_LIST_ARR(0)), _
             LenB(PUB_NEW_GROUP_LIST_ARR(0)) * UBound(PUB_NEW_GROUP_LIST_ARR)

  CopyMemory ByVal VarPtr(PUB_GROUP_SIZE_ARR(0)), _
             ByVal VarPtr(PUB_NEW_GROUP_SIZE_ARR(0)), _
             LenB(PUB_NEW_GROUP_SIZE_ARR(0)) * UBound(PUB_NEW_GROUP_SIZE_ARR)
  PUB_NGROUPS_VAL = PUB_N_NEW_GROUPS_VAL
  
Loop

ReDim OUT_BYTES_ARR(PUB_N_IN_BYTES_VAL + 4)

For i = 1 To PUB_N_IN_BYTES_VAL
  j = PUB_SORT_LIST_ARR(i) - 1
  If j = 0 Then
    j = PUB_N_IN_BYTES_VAL
    StartByte = i
  End If
  OUT_BYTES_ARR(i + 4) = IN_BYTES_ARR(j)
Next i

For i = 1 To 3
  j = Int(StartByte / 256)
  OUT_BYTES_ARR(i) = StartByte - j * 256
  StartByte = j
Next i
OUT_BYTES_ARR(4) = StartByte

SORT_BYTE_ARR_FUNC = True

Exit Function
ERROR_LABEL:
SORT_BYTE_ARR_FUNC = False
End Function

Private Function SORT_GROUP_FUNC(ByRef GROUP_NO_VAL As Long, _
ByRef DEPTH_VAL As Long) As Long

Dim g As Long
Dim i As Long
Dim j As Long

Dim U As Long
Dim V As Long
Dim W As Long
Dim z As Long

Dim index As Long
Dim Index2 As Long
Dim FirstItem As Long
Dim Distance As Long
Dim value As Long
Dim NumEls As Long

On Error GoTo ERROR_LABEL

SORT_GROUP_FUNC = False

FirstItem = PUB_GROUP_LIST_ARR(GROUP_NO_VAL)
NumEls = PUB_GROUP_SIZE_ARR(GROUP_NO_VAL) - 1
ReDim tGroup(NumEls)

Distance = 0
Do
  Distance = Distance * 3 + 1
Loop Until Distance > NumEls + 1

Do
  Distance = Distance \ 3
  For index = FirstItem + Distance To FirstItem + NumEls
    value = PUB_SORT_LIST_ARR(index)
    U = value + DEPTH_VAL
    Do While U > PUB_N_IN_BYTES_VAL
        U = U - PUB_N_IN_BYTES_VAL
    Loop
    z = PUB_ORIGINAL_SORT_ARR(U)
    Index2 = index
    Do
      W = PUB_SORT_LIST_ARR(Index2 - Distance) + DEPTH_VAL
      Do While W > PUB_N_IN_BYTES_VAL
        W = W - PUB_N_IN_BYTES_VAL
      Loop
      If PUB_ORIGINAL_SORT_ARR(W) <= z Then Exit Do
      PUB_SORT_LIST_ARR(Index2) = PUB_SORT_LIST_ARR(Index2 - Distance)
      Index2 = Index2 - Distance
      If Index2 < FirstItem + Distance Then Exit Do
    Loop
    PUB_SORT_LIST_ARR(Index2) = value
  Next
Loop Until Distance <= 1

W = PUB_GROUP_LIST_ARR(GROUP_NO_VAL) + PUB_GROUP_SIZE_ARR(GROUP_NO_VAL) - 1
U = PUB_SORT_LIST_ARR(PUB_GROUP_LIST_ARR(GROUP_NO_VAL)) + DEPTH_VAL
Do While U > PUB_N_IN_BYTES_VAL
    U = U - PUB_N_IN_BYTES_VAL
Loop
U = PUB_ORIGINAL_SORT_ARR(U)
z = 1
For i = PUB_GROUP_LIST_ARR(GROUP_NO_VAL) + 1 To W
  V = PUB_SORT_LIST_ARR(i) + DEPTH_VAL
  Do While V > PUB_N_IN_BYTES_VAL
    V = V - PUB_N_IN_BYTES_VAL
  Loop
  V = PUB_ORIGINAL_SORT_ARR(V)
  If U = V Then
    z = z + 1
  Else
    If z > 1 Then
      PUB_N_NEW_GROUPS_VAL = PUB_N_NEW_GROUPS_VAL + 1
      g = i - z
      PUB_NEW_GROUP_LIST_ARR(PUB_N_NEW_GROUPS_VAL) = g
      PUB_NEW_GROUP_SIZE_ARR(PUB_N_NEW_GROUPS_VAL) = z
      For j = g To i - 1
        PUB_GROUP_ORDER_ARR(PUB_SORT_LIST_ARR(j)) = g
        If g = 0 Then Stop
      Next j
    Else
      PUB_GROUP_ORDER_ARR(PUB_SORT_LIST_ARR(i - 1)) = i - 1
      If i - 1 = 0 Then Stop
    End If
    z = 1
  End If
  U = V
Next i
If z > 1 Then
  PUB_N_NEW_GROUPS_VAL = PUB_N_NEW_GROUPS_VAL + 1
  g = i - z
  PUB_NEW_GROUP_LIST_ARR(PUB_N_NEW_GROUPS_VAL) = g
  PUB_NEW_GROUP_SIZE_ARR(PUB_N_NEW_GROUPS_VAL) = z
  For j = g To i - 1
    PUB_GROUP_ORDER_ARR(PUB_SORT_LIST_ARR(j)) = g
    If g = 0 Then Stop
  Next j
Else
  PUB_GROUP_ORDER_ARR(PUB_SORT_LIST_ARR(i - 1)) = i - 1
  If i - 1 = 0 Then Stop
End If

PUB_GROUP_SIZE_ARR(GROUP_NO_VAL) = 0

SORT_GROUP_FUNC = True

Exit Function
ERROR_LABEL:
SORT_GROUP_FUNC = False
End Function


Private Function REMAP_FUNC(ByRef IN_BYTES_ARR() As Byte, _
ByRef OUT_BYTES_ARR() As Byte) 'PERFECT

Dim i As Long
Dim j As Long
Dim NSIZE As Long

Dim TEMP_ARR(0 To 255) As Long

On Error GoTo ERROR_LABEL

REMAP_FUNC = False

TEMP_ARR(0) = 1
TEMP_ARR(1) = 1
TEMP_ARR(128) = 1
TEMP_ARR(130) = 1
TEMP_ARR(131) = 1
TEMP_ARR(132) = 1
TEMP_ARR(133) = 1
TEMP_ARR(134) = 1
TEMP_ARR(135) = 1
TEMP_ARR(136) = 1
TEMP_ARR(137) = 1
TEMP_ARR(138) = 1
TEMP_ARR(139) = 1
TEMP_ARR(140) = 1
TEMP_ARR(142) = 1
TEMP_ARR(145) = 1
TEMP_ARR(146) = 1
TEMP_ARR(147) = 1
TEMP_ARR(148) = 1
TEMP_ARR(149) = 1
TEMP_ARR(150) = 1
TEMP_ARR(151) = 1
TEMP_ARR(152) = 1
TEMP_ARR(153) = 1
TEMP_ARR(154) = 1
TEMP_ARR(155) = 1
TEMP_ARR(156) = 1
TEMP_ARR(157) = 1
TEMP_ARR(158) = 1
TEMP_ARR(159) = 1

j = 0
NSIZE = UBound(IN_BYTES_ARR)
ReDim OUT_BYTES_ARR(NSIZE * 1.2) As Byte
For i = 1 To NSIZE
  j = j + 1
  If TEMP_ARR(IN_BYTES_ARR(i)) > 0 Then
    OUT_BYTES_ARR(j) = 1
    j = j + 1
    OUT_BYTES_ARR(j) = IN_BYTES_ARR(i) + 40
  Else
    OUT_BYTES_ARR(j) = IN_BYTES_ARR(i)
  End If
Next i
ReDim Preserve OUT_BYTES_ARR(j)

REMAP_FUNC = True

Exit Function
ERROR_LABEL:
REMAP_FUNC = False
End Function


Private Function DEMAP_FUNC(ByRef IN_BYTES_ARR() As Byte, _
ByRef OUT_BYTES_ARR() As Byte) 'PERFECT

Dim i As Long
Dim j As Long
Dim NSIZE As Long

On Error GoTo ERROR_LABEL

DEMAP_FUNC = False

j = 0
NSIZE = UBound(IN_BYTES_ARR)
ReDim OUT_BYTES_ARR(NSIZE) As Byte
For i = 1 To NSIZE
  j = j + 1
  If IN_BYTES_ARR(i) = 1 Then
    i = i + 1
    OUT_BYTES_ARR(j) = IN_BYTES_ARR(i) - 40
  Else
    OUT_BYTES_ARR(j) = IN_BYTES_ARR(i)
  End If
Next i
ReDim Preserve OUT_BYTES_ARR(j)

DEMAP_FUNC = True

Exit Function
ERROR_LABEL:
DEMAP_FUNC = False
End Function
