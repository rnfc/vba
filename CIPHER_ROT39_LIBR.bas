Attribute VB_Name = "CIPHER_ROT39_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.
                    
Function ROT_39_ENCRYPTION_FUNC(ByVal DATA_STR As String, _
Optional ByVal LOWER_BOUND As Long = 48, _
Optional ByVal UPPER_BOUND As Long = 125, _
Optional ByVal CHARMAP As Long = 39)

'Possibly the most famous (or is that infamous) encryption
'technique is the 'Caesar Cipher' - also known as ROT-13,
'purportedly used by Julius Caesar to keep his messages a
'secret should they fall into Brutus' hand. Some web pages
'I've read in researching this material have also implied
'that ROT-13 is the encryption mechanism used in Netscape
'and Eudora. With the ROT-13 cipher, the alphabet is shifted
'13 characters and wrapped around to form a one-to-one
'correspondence between a real letter and a fake
'letter (the 'code'):

'ROT-13 Legend
'A   B   C   D   E   F   G   H   I   J   K   L   M
'O   P   Q   R   S   T   U   V   W   X   Y   Z   A
'N   O   P   Q   R   S   T   U   V   W   X   Y   Z
'B   C   D   E   F   G   H   I   J   K   L   M   N

'Thus, to encode my name using ROT-13, one finds each letter in
'my name in the black letters in the table, and substitutes the
'letter in blue .. my 'encrypted' name becomes FOBRK PWFQV,
'which actually sounds like a European hockey player in the
'NHL. But I digress.

'While ROT-13 was successful in fooling Brutus, ROT-13 can't
'stop a child from decrypting material. You can read more on
'the history and mechanics behind ROT-13 from
'http://www.msen.com/fievel/mmill/book_html/doc004.html.

'ROT-39 follows the same theme, but introduces additional
'characters (the full ASCII set from chr 48  to chr 125.
'Because it encompasses more characters, a ROT-39 message
'appears 'more cryptic', i.e. my ROT-39-encrypted name is "y:G=R iBK<A".

'ASCII Legend used by ROT-39
'48  49  50  51  52  53  54  55  56  57  58  59  60  61  62  63
'0   1   2   3   4   5   6   7   8   9   :   ;   <   =   >   ?
'64  65  66  67  68  69  70  71  72  73  74  75  76  77  78  79
'@   A   B   C   D   E   F   G   H   I   J   K   L   M   N   O
'80  81  82  83  84  85  86  87  88  89  90  91  92  93  94  95
'P   Q   R   S   T   U   V   W   X   Y   Z   [   \   ]   ^ _
'96  97  98  99  100 101 102 103 104 105 106 107 108 109 110 111
'`   a   b   c   d   e   f   g   h   i   j   k   l   m   n   o
'112 113 114 115 116 117 118 119 120 121 122 123 124 125
'p   q   r   s   t   u   v   w   x   y   z   {   |   }


   Dim i As Long
   Dim j As Long
   
   Dim TEMP_ARR() As Byte
   Dim TEMP_STR As String
   
   On Error GoTo ERROR_LABEL
   
  'initialize the byte array to the
  'size of the string passed.
   ReDim TEMP_ARR(0 To Len(DATA_STR)) As Byte
    
  'cast string into the byte array
   TEMP_ARR = StrConv(DATA_STR, vbFromUnicode)
    
   For i = 0 To UBound(TEMP_ARR)
    
     'with the ASCII value of the character
      j = TEMP_ARR(i)
        
     'assure the ASCII value is between
     'the lower and upper limits
      If ((j >= LOWER_BOUND) And (j <= UPPER_BOUND)) Then
         
        'shift the ASCII value by the
        'CHARMAP const value
         j = j + CHARMAP
         
        'perform a check against the upper
        'limit. If the new value exceeds the
        'upper limit, rotate the value to offset
        'from the beginning of the character set.
         If j > UPPER_BOUND Then
            j = j - UPPER_BOUND + LOWER_BOUND - 1
         End If
      End If
        
     'reassign the new shifted value to
     'the current byte
      TEMP_ARR(i) = j
        
   Next i
    
  'convert the byte array back
  'to a string and exit
   TEMP_STR = StrConv(TEMP_ARR, vbUnicode)

ROT_39_ENCRYPTION_FUNC = TEMP_STR

Exit Function
ERROR_LABEL:
ROT_39_ENCRYPTION_FUNC = Err.number
End Function
