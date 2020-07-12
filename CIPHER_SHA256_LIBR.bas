Attribute VB_Name = "CIPHER_SHA256_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.

'http://csrc.nist.gov/publications/fips/fips180-2/fips180-2.pdf

Private PUB_M_lO_ARR(0 To 30) As Long 'm_l2Power
Private PUB_M_l2_ARR(0 To 30) As Long 'm_lOnBits
Private PUB_K_ARR(0 To 63) As Long

Private Const BITS_TO_A_BYTE  As Long = 8
Private Const BYTES_TO_A_WORD As Long = 4
Private Const BITS_TO_A_WORD  As Long = BYTES_TO_A_WORD * BITS_TO_A_BYTE
Private Const MODULUS_BITS As Long = 512
Private Const CONGRUENT_BITS As Long = 448

'SHA256 is a well known algorithm for calculating a short string
'based on the contents of a file, which is effectively unique.
'It would be extremely difficult to alter the file and obtain the
'same hash.

Function SHA256_ENCRYPTION_FUNC(DATA_STR As String) As String
    
    Dim HASH(0 To 7) As Long
    Dim m() As Long
    Dim W(0 To 63) As Long
    Dim A As Long
    Dim B As Long
    Dim c As Long
    Dim D As Long
    Dim E As Long
    Dim f As Long
    Dim g As Long
    Dim h As Long
    Dim i As Long
    Dim j As Long
    Dim T1 As Long
    Dim T2 As Long
    
    On Error GoTo ERROR_LABEL
    
    If SHA256_ARRAY_FUNC() = False Then: GoTo ERROR_LABEL
    
    HASH(0) = &H6A09E667
    HASH(1) = &HBB67AE85
    HASH(2) = &H3C6EF372
    HASH(3) = &HA54FF53A
    HASH(4) = &H510E527F
    HASH(5) = &H9B05688C
    HASH(6) = &H1F83D9AB
    HASH(7) = &H5BE0CD19
    
    SHA256_CONVERT_FUNC DATA_STR, m()
    
    For i = 0 To UBound(m) Step 16
        A = HASH(0)
        B = HASH(1)
        c = HASH(2)
        D = HASH(3)
        E = HASH(4)
        f = HASH(5)
        g = HASH(6)
        h = HASH(7)
        
        For j = 0 To 63
            If j < 16 Then W(j) = m(j + i) Else _
            W(j) = SHA256_ADD_UNSIGNED_FUNC(SHA256_ADD_UNSIGNED_FUNC( _
            SHA256_ADD_UNSIGNED_FUNC(SHA256_GAMMA1_FUNC(W(j - 2)), _
            W(j - 7)), SHA256_GAMMA0_FUNC(W(j - 15))), W(j - 16))
            
            T1 = SHA256_ADD_UNSIGNED_FUNC(SHA256_ADD_UNSIGNED_FUNC( _
            SHA256_ADD_UNSIGNED_FUNC(SHA256_ADD_UNSIGNED_FUNC(h, _
            SHA256_SIGMA1_FUNC(E)), SHA256_CH_FUNC(E, f, g)), PUB_K_ARR(j)), W(j))
            
            T2 = SHA256_ADD_UNSIGNED_FUNC(SHA256_SIGMA0_FUNC(A), _
            SHA256_MAJ_FUNC(A, B, c))
            
            h = g
            g = f
            f = E
            E = SHA256_ADD_UNSIGNED_FUNC(D, T1)
            D = c
            c = B
            B = A
            A = SHA256_ADD_UNSIGNED_FUNC(T1, T2)
        Next
        
        HASH(0) = SHA256_ADD_UNSIGNED_FUNC(A, HASH(0))
        HASH(1) = SHA256_ADD_UNSIGNED_FUNC(B, HASH(1))
        HASH(2) = SHA256_ADD_UNSIGNED_FUNC(c, HASH(2))
        HASH(3) = SHA256_ADD_UNSIGNED_FUNC(D, HASH(3))
        HASH(4) = SHA256_ADD_UNSIGNED_FUNC(E, HASH(4))
        HASH(5) = SHA256_ADD_UNSIGNED_FUNC(f, HASH(5))
        HASH(6) = SHA256_ADD_UNSIGNED_FUNC(g, HASH(6))
        HASH(7) = SHA256_ADD_UNSIGNED_FUNC(h, HASH(7))
    Next
    SHA256_ENCRYPTION_FUNC = LCase$(Right$("00000000" & _
                            Hex(HASH(0)), 8) & Right$("00000000" & _
                            Hex(HASH(1)), 8) & Right$("00000000" & _
                            Hex(HASH(2)), 8) & Right$("00000000" & _
                            Hex(HASH(3)), 8) & Right$("00000000" & _
                            Hex(HASH(4)), 8) & Right$("00000000" & _
                            Hex(HASH(5)), 8) & Right$("00000000" & _
                            Hex(HASH(6)), 8) & Right$("00000000" & _
                            Hex(HASH(7)), 8))

Exit Function
ERROR_LABEL:
SHA256_ENCRYPTION_FUNC = Err.number
End Function

Function TEST_SHA256_ENCRYPTION_FUNC() As Boolean

'the following is a test string provided in the SHA-256 specification
'it should produce the following string

Dim DATA_STR As String
Dim RESULT_STR As String

On Error GoTo ERROR_LABEL

DATA_STR = "abcdbcdecdefdefgefghfghighijhijkijkljklmklmnlmnomnopnopq"
RESULT_STR = "248d6a61d20638b8e5c026930c3e6039a33ce45964ff2167f6ecedd419db06c1"

If SHA256_ENCRYPTION_FUNC(DATA_STR) = RESULT_STR Then
  TEST_SHA256_ENCRYPTION_FUNC = True
Else
  TEST_SHA256_ENCRYPTION_FUNC = False
End If

Exit Function
ERROR_LABEL:
TEST_SHA256_ENCRYPTION_FUNC = False
End Function
Private Function SHA256_ARRAY_FUNC()
        
    On Error GoTo ERROR_LABEL
    
    SHA256_ARRAY_FUNC = False
    
    PUB_M_lO_ARR(0) = 1            ' 00000000000000000000000000000001
    PUB_M_lO_ARR(1) = 3            ' 00000000000000000000000000000011
    PUB_M_lO_ARR(2) = 7            ' 00000000000000000000000000000111
    PUB_M_lO_ARR(3) = 15           ' 00000000000000000000000000001111
    PUB_M_lO_ARR(4) = 31           ' 00000000000000000000000000011111
    PUB_M_lO_ARR(5) = 63           ' 00000000000000000000000000111111
    PUB_M_lO_ARR(6) = 127          ' 00000000000000000000000001111111
    PUB_M_lO_ARR(7) = 255          ' 00000000000000000000000011111111
    PUB_M_lO_ARR(8) = 511          ' 00000000000000000000000111111111
    PUB_M_lO_ARR(9) = 1023         ' 00000000000000000000001111111111
    PUB_M_lO_ARR(10) = 2047        ' 00000000000000000000011111111111
    PUB_M_lO_ARR(11) = 4095        ' 00000000000000000000111111111111
    PUB_M_lO_ARR(12) = 8191        ' 00000000000000000001111111111111
    PUB_M_lO_ARR(13) = 16383       ' 00000000000000000011111111111111
    PUB_M_lO_ARR(14) = 32767       ' 00000000000000000111111111111111
    PUB_M_lO_ARR(15) = 65535       ' 00000000000000001111111111111111
    PUB_M_lO_ARR(16) = 131071      ' 00000000000000011111111111111111
    PUB_M_lO_ARR(17) = 262143      ' 00000000000000111111111111111111
    PUB_M_lO_ARR(18) = 524287      ' 00000000000001111111111111111111
    PUB_M_lO_ARR(19) = 1048575     ' 00000000000011111111111111111111
    PUB_M_lO_ARR(20) = 2097151     ' 00000000000111111111111111111111
    PUB_M_lO_ARR(21) = 4194303     ' 00000000001111111111111111111111
    PUB_M_lO_ARR(22) = 8388607     ' 00000000011111111111111111111111
    PUB_M_lO_ARR(23) = 16777215    ' 00000000111111111111111111111111
    PUB_M_lO_ARR(24) = 33554431    ' 00000001111111111111111111111111
    PUB_M_lO_ARR(25) = 67108863    ' 00000011111111111111111111111111
    PUB_M_lO_ARR(26) = 134217727   ' 00000111111111111111111111111111
    PUB_M_lO_ARR(27) = 268435455   ' 00001111111111111111111111111111
    PUB_M_lO_ARR(28) = 536870911   ' 00011111111111111111111111111111
    PUB_M_lO_ARR(29) = 1073741823  ' 00111111111111111111111111111111
    PUB_M_lO_ARR(30) = 2147483647  ' 01111111111111111111111111111111
    PUB_M_l2_ARR(0) = 1            ' 00000000000000000000000000000001
    PUB_M_l2_ARR(1) = 2            ' 00000000000000000000000000000010
    PUB_M_l2_ARR(2) = 4            ' 00000000000000000000000000000100
    PUB_M_l2_ARR(3) = 8            ' 00000000000000000000000000001000
    PUB_M_l2_ARR(4) = 16           ' 00000000000000000000000000010000
    PUB_M_l2_ARR(5) = 32           ' 00000000000000000000000000100000
    PUB_M_l2_ARR(6) = 64           ' 00000000000000000000000001000000
    PUB_M_l2_ARR(7) = 128          ' 00000000000000000000000010000000
    PUB_M_l2_ARR(8) = 256          ' 00000000000000000000000100000000
    PUB_M_l2_ARR(9) = 512          ' 00000000000000000000001000000000
    PUB_M_l2_ARR(10) = 1024        ' 00000000000000000000010000000000
    PUB_M_l2_ARR(11) = 2048        ' 00000000000000000000100000000000
    PUB_M_l2_ARR(12) = 4096        ' 00000000000000000001000000000000
    PUB_M_l2_ARR(13) = 8192        ' 00000000000000000010000000000000
    PUB_M_l2_ARR(14) = 16384       ' 00000000000000000100000000000000
    PUB_M_l2_ARR(15) = 32768       ' 00000000000000001000000000000000
    PUB_M_l2_ARR(16) = 65536       ' 00000000000000010000000000000000
    PUB_M_l2_ARR(17) = 131072      ' 00000000000000100000000000000000
    PUB_M_l2_ARR(18) = 262144      ' 00000000000001000000000000000000
    PUB_M_l2_ARR(19) = 524288      ' 00000000000010000000000000000000
    PUB_M_l2_ARR(20) = 1048576     ' 00000000000100000000000000000000
    PUB_M_l2_ARR(21) = 2097152     ' 00000000001000000000000000000000
    PUB_M_l2_ARR(22) = 4194304     ' 00000000010000000000000000000000
    PUB_M_l2_ARR(23) = 8388608     ' 00000000100000000000000000000000
    PUB_M_l2_ARR(24) = 16777216    ' 00000001000000000000000000000000
    PUB_M_l2_ARR(25) = 33554432    ' 00000010000000000000000000000000
    PUB_M_l2_ARR(26) = 67108864    ' 00000100000000000000000000000000
    PUB_M_l2_ARR(27) = 134217728   ' 00001000000000000000000000000000
    PUB_M_l2_ARR(28) = 268435456   ' 00010000000000000000000000000000
    PUB_M_l2_ARR(29) = 536870912   ' 00100000000000000000000000000000
    PUB_M_l2_ARR(30) = 1073741824  ' 01000000000000000000000000000000
    
    PUB_K_ARR(0) = &H428A2F98
    PUB_K_ARR(1) = &H71374491
    PUB_K_ARR(2) = &HB5C0FBCF
    PUB_K_ARR(3) = &HE9B5DBA5
    PUB_K_ARR(4) = &H3956C25B
    PUB_K_ARR(5) = &H59F111F1
    PUB_K_ARR(6) = &H923F82A4
    PUB_K_ARR(7) = &HAB1C5ED5
    PUB_K_ARR(8) = &HD807AA98
    PUB_K_ARR(9) = &H12835B01
    PUB_K_ARR(10) = &H243185BE
    PUB_K_ARR(11) = &H550C7DC3
    PUB_K_ARR(12) = &H72BE5D74
    PUB_K_ARR(13) = &H80DEB1FE
    PUB_K_ARR(14) = &H9BDC06A7
    PUB_K_ARR(15) = &HC19BF174
    PUB_K_ARR(16) = &HE49B69C1
    PUB_K_ARR(17) = &HEFBE4786
    PUB_K_ARR(18) = &HFC19DC6
    PUB_K_ARR(19) = &H240CA1CC
    PUB_K_ARR(20) = &H2DE92C6F
    PUB_K_ARR(21) = &H4A7484AA
    PUB_K_ARR(22) = &H5CB0A9DC
    PUB_K_ARR(23) = &H76F988DA
    PUB_K_ARR(24) = &H983E5152
    PUB_K_ARR(25) = &HA831C66D
    PUB_K_ARR(26) = &HB00327C8
    PUB_K_ARR(27) = &HBF597FC7
    PUB_K_ARR(28) = &HC6E00BF3
    PUB_K_ARR(29) = &HD5A79147
    PUB_K_ARR(30) = &H6CA6351
    PUB_K_ARR(31) = &H14292967
    PUB_K_ARR(32) = &H27B70A85
    PUB_K_ARR(33) = &H2E1B2138
    PUB_K_ARR(34) = &H4D2C6DFC
    PUB_K_ARR(35) = &H53380D13
    PUB_K_ARR(36) = &H650A7354
    PUB_K_ARR(37) = &H766A0ABB
    PUB_K_ARR(38) = &H81C2C92E
    PUB_K_ARR(39) = &H92722C85
    PUB_K_ARR(40) = &HA2BFE8A1
    PUB_K_ARR(41) = &HA81A664B
    PUB_K_ARR(42) = &HC24B8B70
    PUB_K_ARR(43) = &HC76C51A3
    PUB_K_ARR(44) = &HD192E819
    PUB_K_ARR(45) = &HD6990624
    PUB_K_ARR(46) = &HF40E3585
    PUB_K_ARR(47) = &H106AA070
    PUB_K_ARR(48) = &H19A4C116
    PUB_K_ARR(49) = &H1E376C08
    PUB_K_ARR(50) = &H2748774C
    PUB_K_ARR(51) = &H34B0BCB5
    PUB_K_ARR(52) = &H391C0CB3
    PUB_K_ARR(53) = &H4ED8AA4A
    PUB_K_ARR(54) = &H5B9CCA4F
    PUB_K_ARR(55) = &H682E6FF3
    PUB_K_ARR(56) = &H748F82EE
    PUB_K_ARR(57) = &H78A5636F
    PUB_K_ARR(58) = &H84C87814
    PUB_K_ARR(59) = &H8CC70208
    PUB_K_ARR(60) = &H90BEFFFA
    PUB_K_ARR(61) = &HA4506CEB
    PUB_K_ARR(62) = &HBEF9A3F7
    PUB_K_ARR(63) = &HC67178F2
    
    SHA256_ARRAY_FUNC = True

Exit Function
ERROR_LABEL:
SHA256_ARRAY_FUNC = False
End Function


Private Function SHA256_CONVERT_FUNC(DATA_STR As String, _
lWordArray() As Long)
    
    Dim lMessageLength As Long
    Dim lNumberOfWords As Long
    Dim lBytePosition As Long
    Dim lByteCount As Long
    Dim lWordCount As Long
    Dim lByte As Long
    
    lMessageLength = Len(DATA_STR)
    lNumberOfWords = (((lMessageLength + ((MODULUS_BITS - _
                    CONGRUENT_BITS) \ BITS_TO_A_BYTE)) \ _
                    (MODULUS_BITS \ BITS_TO_A_BYTE)) + 1) * _
                    (MODULUS_BITS \ BITS_TO_A_WORD)
    ReDim lWordArray(0 To lNumberOfWords - 1)
    Do Until lByteCount >= lMessageLength
        lWordCount = lByteCount \ BYTES_TO_A_WORD
        lBytePosition = (3 - (lByteCount Mod BYTES_TO_A_WORD)) * BITS_TO_A_BYTE
        lByte = AscB(Mid$(DATA_STR, lByteCount + 1, 1))
        lWordArray(lWordCount) = lWordArray(lWordCount) Or _
        SHA256_LSHIFT_FUNC(lByte, lBytePosition)
        lByteCount = lByteCount + 1
    Loop
    lWordCount = lByteCount \ BYTES_TO_A_WORD
    lBytePosition = (3 - (lByteCount Mod BYTES_TO_A_WORD)) * BITS_TO_A_BYTE
    lWordArray(lWordCount) = lWordArray(lWordCount) Or _
                            SHA256_LSHIFT_FUNC(&H80, lBytePosition)
    lWordArray(lNumberOfWords - 1) = SHA256_LSHIFT_FUNC(lMessageLength, 3)
    lWordArray(lNumberOfWords - 2) = SHA256_RSHIFT_FUNC(lMessageLength, 29)
    'SHA256_CONVERT_FUNC = lWordArray
End Function


Private Function SHA256_LSHIFT_FUNC(ByVal lValue As Long, _
ByVal iShiftBits As Integer) As Long
    If iShiftBits = 0 Then
        SHA256_LSHIFT_FUNC = lValue
        Exit Function
    ElseIf iShiftBits = 31 Then
        If lValue And 1 Then SHA256_LSHIFT_FUNC = &H80000000 Else _
        SHA256_LSHIFT_FUNC = 0
        Exit Function
    ElseIf iShiftBits < 0 Or iShiftBits > 31 Then
        Err.Raise 6
    End If
    If (lValue And PUB_M_l2_ARR(31 - iShiftBits)) Then _
    SHA256_LSHIFT_FUNC = ((lValue And PUB_M_lO_ARR(31 - _
    (iShiftBits + 1))) * PUB_M_l2_ARR(iShiftBits)) Or _
    &H80000000 Else SHA256_LSHIFT_FUNC = ((lValue And _
    PUB_M_lO_ARR(31 - iShiftBits)) * PUB_M_l2_ARR(iShiftBits))
End Function
Private Function SHA256_RSHIFT_FUNC(ByVal lValue As Long, _
ByVal iShiftBits As Integer) As Long
    If iShiftBits = 0 Then
        SHA256_RSHIFT_FUNC = lValue
        Exit Function
    ElseIf iShiftBits = 31 Then
        If lValue And &H80000000 Then SHA256_RSHIFT_FUNC = 1 Else _
        SHA256_RSHIFT_FUNC = 0
        Exit Function
    ElseIf iShiftBits < 0 Or iShiftBits > 31 Then
        Err.Raise 6
    End If
    SHA256_RSHIFT_FUNC = (lValue And &H7FFFFFFE) \ PUB_M_l2_ARR(iShiftBits)
    If (lValue And &H80000000) Then SHA256_RSHIFT_FUNC = (SHA256_RSHIFT_FUNC _
    Or (&H40000000 \ PUB_M_l2_ARR(iShiftBits - 1)))
End Function

Private Function SHA256_ADD_UNSIGNED_FUNC(ByVal lX As Long, _
ByVal lY As Long) As Long
    Dim lX4 As Long
    Dim lY4 As Long
    Dim lX8 As Long
    Dim lY8 As Long
    Dim lResult As Long
    lX8 = lX And &H80000000
    lY8 = lY And &H80000000
    lX4 = lX And &H40000000
    lY4 = lY And &H40000000
    lResult = (lX And &H3FFFFFFF) + (lY And &H3FFFFFFF)
    If lX4 And lY4 Then
        lResult = lResult Xor &H80000000 Xor lX8 Xor lY8
    ElseIf lX4 Or lY4 Then
        If lResult And &H40000000 Then lResult = lResult Xor _
        &HC0000000 Xor lX8 Xor lY8 Else lResult = lResult Xor _
        &H40000000 Xor lX8 Xor lY8
    Else
        lResult = lResult Xor lX8 Xor lY8
    End If
    SHA256_ADD_UNSIGNED_FUNC = lResult
End Function
Private Function SHA256_CH_FUNC(ByVal x As Long, _
ByVal Y As Long, _
ByVal z As Long) As Long
    SHA256_CH_FUNC = ((x And Y) Xor ((Not x) And z))
End Function
Private Function SHA256_MAJ_FUNC(ByVal x As Long, _
ByVal Y As Long, _
ByVal z As Long) As Long
    SHA256_MAJ_FUNC = ((x And Y) Xor (x And z) Xor (Y And z))
End Function

Private Function SHA256_S_FUNC(ByVal x As Long, _
ByVal n As Long) As Long
    SHA256_S_FUNC = (SHA256_RSHIFT_FUNC(x, (n And PUB_M_lO_ARR(4))) Or _
    SHA256_LSHIFT_FUNC(x, (32 - (n And PUB_M_lO_ARR(4)))))
End Function
Private Function SHA256_R_FUNC(ByVal x As Long, _
ByVal n As Long) As Long
    SHA256_R_FUNC = SHA256_RSHIFT_FUNC(x, CInt(n And PUB_M_lO_ARR(4)))
End Function
Private Function SHA256_SIGMA0_FUNC(ByVal x As Long) As Long
    SHA256_SIGMA0_FUNC = (SHA256_S_FUNC(x, 2) Xor _
    SHA256_S_FUNC(x, 13) Xor SHA256_S_FUNC(x, 22))
End Function

Private Function SHA256_SIGMA1_FUNC(ByVal x As Long) As Long
    SHA256_SIGMA1_FUNC = (SHA256_S_FUNC(x, 6) Xor _
    SHA256_S_FUNC(x, 11) Xor SHA256_S_FUNC(x, 25))
End Function

Private Function SHA256_GAMMA0_FUNC(ByVal x As Long) As Long
    SHA256_GAMMA0_FUNC = (SHA256_S_FUNC(x, 7) Xor _
    SHA256_S_FUNC(x, 18) Xor SHA256_R_FUNC(x, 3))
End Function

Private Function SHA256_GAMMA1_FUNC(ByVal x As Long) As Long
    SHA256_GAMMA1_FUNC = (SHA256_S_FUNC(x, 17) Xor _
    SHA256_S_FUNC(x, 19) Xor SHA256_R_FUNC(x, 10))
End Function
