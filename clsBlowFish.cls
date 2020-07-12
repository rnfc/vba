VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsBlowFish"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'-------------------------------------------------------------------------------
'-------------------------------------------------------------------------------
Option Explicit
'-------------------------------------------------------------------------------
'-------------------------------------------------------------------------------

'This class allows users to encode member data using secure Blowfish
'encryption. The encoded data can then be copied to an add-in, along with
'the decryption routines, to be accessed

'You can find out about Blowfish --> http://www.schneier.com/blowfish.html.
'It is free. This version is written in VBA and is not intended for heavy data
'processing, because it would be too slow. It works fine for providing secure
'data which can be accessed by logging in, because it only has to decrypt one
'line of data.

'You put your data in columns in the data sheet, anything you like, as much
'as you want. The first column is for the PIN number, and the first row
'contains headings.

'On the main sheet, you select the length of the PIN number and also the
'characters which can be used to make it up. The more the better because
'it increases the number of combinations.

'The workbook then creates a unique PIN number for each record and
'concatenates the data as a comma-delimited string, all in the Plaintext
'sheet.

'It then encrypts the PIN number and also the data string, and concatenates
'them and puts them in the EncodedData sheet.

'All of the above is done with buttons on the Main sheet, which are clearly
'labelled.

'How to implement it in your user workbook

'You then copy this data range to the workbook you plan to give to users,
'along with the Login and clsBlowFish code modules. To see how to handle
'logins, go to the Login sheet, which shows you how to decrypt the data
'for a given user.

'-------------------------------------------------------------------------------
'-------------------------------------------------------------------------------
'-------------------------------------------------------------------------------
'-------------------------------------------------------------------------------

Private Const EXPORT As Boolean = False
'TRUE means limit key length to 40 bits for
'legal export (otherwise key limited to 448 bits)

Const ERROR_NONE As Long = 0
Const ERROR_BUG As Long = 1
Const ERROR_NO_PASSWORD As Long = 2
Const ERROR_NOT_EIGHT_BYTES As Long = 3

Private Type typWord32
    lngByte1 As Long
    lngByte2 As Long
    lngByte3 As Long
    lngByte4 As Long
End Type

Private Type typWord64
    udtLow As typWord32
    udtHigh As typWord32
End Type

Private udtPBox(17) As typWord32
Private udtSBox(3, 255) As typWord32

Private intError As Integer

'=================================================
'P U B L I C     M E M B E R     F U N C T I O N S
'=================================================
Public Function error() As Integer
    error = intError
End Function

Public Function test() As Boolean
    On Error GoTo ErrorHandler
    intError = ERROR_NONE
    
    test = BF_TestVectors()
    
    Exit Function

ErrorHandler:
    intError = ERROR_BUG
End Function

Public Sub password(strPassword As String)
    On Error GoTo ErrorHandler
    intError = ERROR_NONE
    
    If Len(strPassword) = 0 Then
        intError = ERROR_NO_PASSWORD
        Exit Sub
    End If
    
    Call BF_Password(strPassword)
    Exit Sub
    
ErrorHandler:
    intError = ERROR_BUG
End Sub

Public Function EncryptString(ByVal tString As String, _
Optional ConvertToHEX As Boolean) As String

Dim ReturnString As String, PartialString As String * 8

Do While tString <> ""
  PartialString = Left$(tString & "        ", 8)
  ReturnString = ReturnString & Encrypt(PartialString)
  tString = Mid$(tString, 9)
Loop

If ConvertToHEX = True Then
  EncryptString = ToHEX(ReturnString)
    Else
  EncryptString = ReturnString
End If

End Function

Public Function DecryptString(ByVal tString As String, _
Optional ConvertFromHEX As Boolean) As String

Dim ReturnString As String, PartialString As String * 8

If ConvertFromHEX = True Then
  tString = HexToString(tString)
End If

Do While tString <> ""
  PartialString = Left$(tString, 8)
  ReturnString = ReturnString & Decrypt(PartialString)
  tString = Mid$(tString, 9)
Loop

DecryptString = ReturnString

End Function

Public Function Encrypt(strInput As String) As String
    Dim udtWord64 As typWord64
    Dim bytString() As Byte
    Dim lngError As Long
    
    On Error GoTo ErrorHandler
    lngError = ERROR_NONE
        
    bytString = StrConv(strInput, vbFromUnicode)
    
    If UBound(bytString) <> 7 Then
        lngError = ERROR_NOT_EIGHT_BYTES
        Exit Function
    End If

    udtWord64.udtLow.lngByte1 = bytString(0)
    udtWord64.udtLow.lngByte2 = bytString(1)
    udtWord64.udtLow.lngByte3 = bytString(2)
    udtWord64.udtLow.lngByte4 = bytString(3)
    udtWord64.udtHigh.lngByte1 = bytString(4)
    udtWord64.udtHigh.lngByte2 = bytString(5)
    udtWord64.udtHigh.lngByte3 = bytString(6)
    udtWord64.udtHigh.lngByte4 = bytString(7)

    udtWord64 = BF_Encrypt(udtWord64)
    
    bytString(0) = udtWord64.udtLow.lngByte1
    bytString(1) = udtWord64.udtLow.lngByte2
    bytString(2) = udtWord64.udtLow.lngByte3
    bytString(3) = udtWord64.udtLow.lngByte4
    bytString(4) = udtWord64.udtHigh.lngByte1
    bytString(5) = udtWord64.udtHigh.lngByte2
    bytString(6) = udtWord64.udtHigh.lngByte3
    bytString(7) = udtWord64.udtHigh.lngByte4
        
    Encrypt = StrConv(bytString, vbUnicode)
    Exit Function
    
ErrorHandler:
    lngError = ERROR_BUG
End Function

Public Function Decrypt(strInput As String) As String
    Dim udtWord64 As typWord64
    Dim bytString() As Byte
    Dim lngError As Long
    
    On Error GoTo ErrorHandler
    lngError = ERROR_NONE
    
    bytString = StrConv(strInput, vbFromUnicode)
    
    If UBound(bytString) <> 7 Then
        lngError = ERROR_NOT_EIGHT_BYTES
        Exit Function
    End If

    udtWord64.udtLow.lngByte1 = bytString(0)
    udtWord64.udtLow.lngByte2 = bytString(1)
    udtWord64.udtLow.lngByte3 = bytString(2)
    udtWord64.udtLow.lngByte4 = bytString(3)
    udtWord64.udtHigh.lngByte1 = bytString(4)
    udtWord64.udtHigh.lngByte2 = bytString(5)
    udtWord64.udtHigh.lngByte3 = bytString(6)
    udtWord64.udtHigh.lngByte4 = bytString(7)

    udtWord64 = BF_Decrypt(udtWord64)
    
    bytString(0) = udtWord64.udtLow.lngByte1
    bytString(1) = udtWord64.udtLow.lngByte2
    bytString(2) = udtWord64.udtLow.lngByte3
    bytString(3) = udtWord64.udtLow.lngByte4
    bytString(4) = udtWord64.udtHigh.lngByte1
    bytString(5) = udtWord64.udtHigh.lngByte2
    bytString(6) = udtWord64.udtHigh.lngByte3
    bytString(7) = udtWord64.udtHigh.lngByte4
        
    Decrypt = StrConv(bytString, vbUnicode)
    Exit Function
    
ErrorHandler:
    lngError = ERROR_BUG
End Function

'=========================
'P R I V A T E     C O D E
'=========================
Private Function Xor32(udtWord1 As typWord32, _
udtWord2 As typWord32) As typWord32
    Xor32.lngByte1 = udtWord1.lngByte1 Xor udtWord2.lngByte1
    Xor32.lngByte2 = udtWord1.lngByte2 Xor udtWord2.lngByte2
    Xor32.lngByte3 = udtWord1.lngByte3 Xor udtWord2.lngByte3
    Xor32.lngByte4 = udtWord1.lngByte4 Xor udtWord2.lngByte4
End Function

Private Function Add32(udtWord1 As typWord32, _
udtWord2 As typWord32) As typWord32
    Dim lngSum As Long
    
    lngSum = udtWord1.lngByte4 + udtWord2.lngByte4
    Add32.lngByte4 = lngSum And 255
        
    lngSum = udtWord1.lngByte3 + udtWord2.lngByte3 + lngSum \ 256
    Add32.lngByte3 = lngSum And 255
        
    lngSum = udtWord1.lngByte2 + udtWord2.lngByte2 + lngSum \ 256
    Add32.lngByte2 = lngSum And 255
        
    lngSum = udtWord1.lngByte1 + udtWord2.lngByte1 + lngSum \ 256
    Add32.lngByte1 = lngSum And 255
End Function

Private Sub BF_Initialise()

    Dim strPi As String
    Dim lngCount1 As Long
    Dim lngCount2 As Long
    Dim lngCount3 As Long

    strPi = _
    "243f6a8885a308d313198a2e03707344a4093822299f31d0082efa98ec4e6c" & _
    "89452821e638d01377be5466cf34e90c6cc0ac29b7c97c50dd3f84d5b5b54709179216" & _
    "d5d98979fb1b"
    
    lngCount3 = 0
    For lngCount1 = 0 To 17
        udtPBox(lngCount1).lngByte1 = "&H" & Mid(strPi, lngCount3 + 1, 2)
        udtPBox(lngCount1).lngByte2 = "&H" & Mid(strPi, lngCount3 + 3, 2)
        udtPBox(lngCount1).lngByte3 = "&H" & Mid(strPi, lngCount3 + 5, 2)
        udtPBox(lngCount1).lngByte4 = "&H" & Mid(strPi, lngCount3 + 7, 2)
        
        lngCount3 = lngCount3 + 8
    Next
    
    strPi = "d1310ba698dfb5ac2ffd72dbd01adfb7b8e1afed6a267e96ba7c9045f12c7" & _
    "f9924a19947b3916cf70801f2e2858efc16636920d871574e69a458fea3f4933d7e0" & _
    "d95748f728eb658718bcd5882154aee7b54a41dc25a59b59c30d5392af26013"
    
    strPi = strPi & "c5d1b023286085f0ca417918b8db38ef8e79dcb0603a180e6c9e0e8bb" & _
    "01e8a3ed71577c1bd314b2778af2fda55605c60e65525f3aa55ab945748986263e8144055" & _
    "ca396a2aab10b6b4cc5c341141e8cea15486af7c72e993b3ee1411636fbc2a2ba9c55d" & _
    "741831f6ce5c3e169b87931eafd6ba336c24cf5c7a325381289586773b8f48986b4bb9af" & _
    "c4bfe81b6628219361d809ccfb21a991487cac605dec8032ef845d5de98575b1dc262302" & _
    "eb651b8823893e81d396acc50f6d6ff383f442392e0b4482a484200469c8f04a9e1f9b5e" & _
    "21c66842f6e96c9a670c9c61abd388f06a51a0d2d8542f68960fa728ab5133a36eef0b6c" & _
    "137a3be4ba3bf0507efb2a98a1f1651d39af017666ca593e82430e888cee8619456f9fb47d" & _
    "84a5c33b8b5ebe"
    
    strPi = strPi & "e06f75d885c12073401a449f56c16aa64ed3aa62363f77061bfedf72429b" & _
    "023d37d0d724d00a1248db0fead349f1c09b075372c980991b7b25d479d8f6e8def7e3f" & _
    "e501ab6794c3b976ce0bd04c006bac1a94fb6409f60c45e5c9ec2196a246368fb6faf3e6" & _
    "c53b51339b2eb3b52ec6f6dfc511f9b30952ccc814544af5ebd09bee3d004de334afd660f" & _
    "2807192e4bb3c0cba85745c8740fd20b5f39b9d3fbdb5579c0bd1a60320ad6a100c6402c" & _
    "7279679f25fefb1fa3cc8ea5e9f8db3222f83c7516dffd616b152f501ec8ad0552ab323db" & _
    "5fafd23876053317b483e00df829e5c57bbca6f8ca01a87562edf1769dbd542a8f6287eff" & _
    "c3ac6732c68c4f5573695b27b0bbca58c8e1ffa35db8f011a010fa3d98fd2183b84afcb56" & _
    "c2dd1d35b"
    
    strPi = strPi & "9a53e479b6f84565d28e49bc4bfb9790e1ddf2daa4cb7e3362fb1341c" & _
    "ee4c6e8ef20cada36774c01d07e9efe2bf11fb495dbda4dae909198eaad8e716b93d5a0d" & _
    "08ed1d0afc725e08e3c5b2f8e7594b78ff6e2fbf2122b648888b812900df01c4fad5ea0688" & _
    "fc31cd1cff191b3a8c1ad2f2f2218be0e1777ea752dfe8b021fa1e5a0cc0fb56f74e818ac" & _
    "f3d6ce89e299b4a84fe0fd13e0b77cc43b81d2ada8d9165fa2668095770593cc7314211a" & _
    "1477e6ad206577b5fa86c75442f5fb9d35cfebcdaf0c7b3e89a0d6411bd3ae1e7e4900250e" & _
    "2d2071b35e226800bb57b8e0af2464369bf009b91e5563911d59dfa6aa78c14389d95a537f" & _
    "207d5ba202e5b9c5832603766295cfa911c819684e734a41b3472dca7b14a94a1b5100529" & _
    "a532915"
    
    strPi = strPi & "d60f573fbc9bc6e42b60a47681e6740008ba6fb5571be91ff296ec6b2" & _
    "a0dd915b6636521e7b9f9b6ff34052ec585566453b02d5da99f8fa108ba47996e85076a4b" & _
    "7a70e9b5b32944db75092ec4192623ad6ea6b049a7df7d9cee60b88fedb266ecaa8c71699" & _
    "a17ff5664526cc2b19ee1193602a575094c29a0591340e4183a3e3f54989a5b429d656b8" & _
    "fe4d699f73fd6a1d29c07efe830f54d2d38e6f0255dc14cdd20868470eb266382e9c6021" & _
    "ecc5e09686b3f3ebaefc93c9718146b6a70a1687f358452a0e286b79c5305aa5007373e" & _
    "07841c7fdeae5c8e7d44ec5716f2b8b03ada37f0500c0df01c1f040200b3ffae0cf51a3" & _
    "cb574b225837a58dc0921bdd19113f97ca92ff69432477322f547013ae5e58137c2dadcc" & _
    "8b576349af3dda7"
    
    strPi = strPi & "a94461460fd0030eecc8c73ea4751e41e238cd993bea0e2f3280bba" & _
    "1183eb3314e548b384f6db9086f420d03f60a04bf2cb8129024977c795679b072bcaf89" & _
    "afde9a771fd9930810b38bae12dccf3f2e5512721f2e6b7124501adde69f84cd877a58" & _
    "47187408da17bc9f9abce94b7d8cec7aec3adb851dfa63094366c464c3d2ef1c18473215" & _
    "d908dd433b3724c2ba1612a14d432a65c45150940002133ae4dd71dff89e10314e5581a" & _
    "c77d65f11199b043556f1d7a3c76b3c11183b5924a509f28fe6ed97f1fbfa9ebabf2c1e" & _
    "153c6e86e34570eae96fb1860e5e0a5a3e2ab3771fe71c4e3d06fa2965dcb999e71d0f8" & _
    "03e89d65266c8252e4cc9789c10b36ac6150eba94e2ea78a5fc3c531e0a2df4f2f74ea7" & _
    "361d2b3d1939260f19c27960"
    
    strPi = strPi & "5223a708f71312b6ebadfe6eeac31f66e3bc4595a67bc883b17f37d101" & _
    "8cff28c332ddefbe6c5aa56558218568ab9802eecea50fdb2f953b2aef7dad5b6e2f841521" & _
    "b62829076170ecdd4775619f151013cca830eb61bd960334fe1eaa0363cfb5735c904c70a2" & _
    "39d59e9e0bcbaade14eecc86bc60622ca79cab5cabb2f3846e648b1eaf19bdf0caa02369b9" & _
    "655abb5040685a323c2ab4b3319ee9d5c021b8f79b540b19875fa09995f7997e623d7da8f8" & _
    "37889a97e32d7711ed935f166812810e358829c7e61fd696dedfa17858ba9957f584a51b22" & _
    "72639b83c3ff1ac24696cdb30aeb532e30548fd948e46dbc312858ebf2ef34c6ffeafe28ed" & _
    "61ee7c3c735d4a14d9e864b7e342105d14203e13e045eee2b6a3aaabeadb6c4f15facb4fd0"
    
    strPi = strPi & "c742f442ef6abbb5654f3b1d41cd2105d81e799e86854dc7e44b476a3d" & _
    "816250cf62a1f25b8d2646fc8883a0c1c7b6a37f1524c369cb749247848a0b5692b285095b" & _
    "bf00ad19489d1462b17423820e0058428d2a0c55f5ea1dadf43e233f70613372f0928d937e" & _
    "41d65fecf16c223bdb7cde3759cbee74604085f2a7ce77326ea607808419f8509ee8efd855" & _
    "61d99735a969a7aac50c06c25a04abfc800bcadc9e447a2ec3453484fdd567050e1e9ec9db7" & _
    "3dbd3105588cd675fda79e3674340c5c43465713e38d83d28f89ef16dff20153e21e78fb03d" & _
    "4ae6e39f2bdb83adf7e93d5a68948140f7f64c261c94692934411520f77602d4f7bcf46b2ed" & _
    "4a20068d40824713320f46a43b7d4b7500061af1e39f62e9724454614214f74bf8b8840"
    
    strPi = strPi & "4d95fc1d96b591af70f4ddd366a02f45bfbc09ec03bd97857fac6dd031cb" & _
    "850496eb27b355fd3941da2547e6abca0a9a28507825530429f40a2c86dae9b66dfb68dc1462" & _
    "d7486900680ec0a427a18dee4f3ffea2e887ad8cb58ce0067af4d6b6aace1e7cd3375fecce78" & _
    "a399406b2a4220fe9e35d9f385b9ee39d7ab3b124e8b1dc9faf74b6d185626a36631eae397b2" & _
    "3a6efa74dd5b43326841e7f7ca7820fbfb0af54ed8feb397454056acba48952755533a3a2083" & _
    "8d87fe6ba9b7d096954b55a867bca1159a58cca9296399e1db33a62a4a563f3125f95ef47e1c" & _
    "9029317cfdf8e80204272f7080bb155c05282ce395c11548e4c66d2248c1133fc70f86dc07f9" & _
    "c9ee41041f0f404779a45d886e17325f51ebd59bc0d1f2bcc18f41113564"
    
    strPi = strPi & "257b7834602a9c60dff8e8a31f636c1b0e12b4c202e1329eaf664fd1cad18" & _
    "1156b2395e0333e92e13b240b62eebeb92285b2a20ee6ba0d99de720c8c2da2f728d01278459" & _
    "5b794fd647d0862e7ccf5f05449a36f877d48fac39dfd27f33e8d1e0a476341992eff743a6f6" & _
    "eabf4f8fd37a812dc60a1ebddf8991be14cdb6e6b0dc67b55106d672c372765d43bdcd0e804f" & _
    "1290dc7cc00ffa3b5390f92690fed0b667b9ffbcedb7d9ca091cf0bd9155ea3bb132f88515bad" & _
    "247b9479bf763bd6eb37392eb3cc1159798026e297f42e312d6842ada7c66a2b3b12754ccc782" & _
    "ef11c6a124237b79251e706a1bbe64bfb63501a6b101811caedfa3d25bdd8e2e1c3c944421659" & _
    "0a121386d90cec6ed5abea2a64af674eda86a85fbebfe98864e4c3fe"
    
    strPi = strPi & "9dbc8057f0f7c08660787bf86003604dd1fd8346f6381fb07745ae04d736f" & _
    "ccc83426b33f01eab71b08041873c005e5f77a057bebde8ae2455464299bf582e614e58f48ff2" & _
    "ddfda2f474ef388789bdc25366f9c3c8b38e74b475f25546fcd9b97aeb26618b1ddf84846a0e79" & _
    "915f95e2466e598e20b457708cd55591c902de4cb90bace1bb8205d011a862487574a99eb77f1" & _
    "9b6e0a9dc09662d09a1c4324633e85a1f0209f0be8c4a99a0251d6efe101ab93d1d0ba5a4dfa1" & _
    "86f20f2868f169dcb7da83573906fea1e2ce9b4fcd7f5250115e01a70683faa002b5c40de6d02" & _
    "79af88c27773f8641c3604c0661a806b5f0177a28c0f586e0006058aa30dc7d6211e69ed72338e" & _
    "a6353c2dd94c2c21634bbcbee5690bcb6deebfc7da1ce591d76"
    
    strPi = strPi & "6f05e4094b7c018839720a3d7c927c2486e3725f724d9db91ac15bb4d39eb" & _
    "8fced54557808fca5b5d83d7cd34dad0fc41e50ef5eb161e6f8a28514d96c51133c6fd5c7e75" & _
    "6e14ec4362abfceddc6c837d79a323492638212670efa8e406000e03a39ce37d3faf5cfabc27" & _
    "7375ac52d1b5cb0679e4fa33742d382274099bc9bbed5118e9dbf0f7315d62d1c7ec700c47bb" & _
    "78c1b6b21a19045b26eb1be6a366eb45748ab2fbc946e79c6a376d26549c2c8530ff8ee468dd" & _
    "e7dd5730a1d4cd04dc62939bbdba9ba4650ac9526e8be5ee304a1fad5f06a2d519a63ef8ce29" & _
    "a86ee22c089c2b843242ef6a51e03aa9cf2d0a483c061ba9be96a4d8fe51550ba645bd62826a" & _
    "2f9a73a3ae14ba99586ef5562e9c72fefd3f752f7da3f046f6977fa0a59"
    
    strPi = strPi & "80e4a91587b086019b09e6ad3b3ee593e990fd5a9e34d7972cf0b7d9022b8b" & _
    "5196d5ac3a017da67dd1cf3ed67c7d2d281f9f25cfadf2b89b5ad6b4725a88f54ce029ac71e019" & _
    "a5e647b0acfded93fa9be8d3c48d283b57ccf8d5662979132e28785f0191ed756055f7960e44e3" & _
    "d35e8c15056dd488f46dba03a161250564f0bdc3eb9e153c9057a297271aeca93a072a1b3f6d9b" & _
    "1e6321f5f59c66fb26dcf3197533d928b155fdf5035634828aba3cbb28517711c20ad9f8abcc51" & _
    "67ccad925f4de817513830dc8e379d58629320f991ea7a90c2fb3e7bce5121ce64774fbe32a8b6" & _
    "e37ec3293d4648de53696413e680a2ae0810dd6db22469852dfd09072166b39a460a6445c0dd58" & _
    "6cdecf1c20c8ae5bbef7dd1b588d40ccd2017f6bb4e3bb"
    
    strPi = strPi & "dda26a7e3a59ff453e350a44bcb4cdd572eacea8fa6484bb8d6612aebf3c6f4" & _
    "7d29be463542f5d9eaec2771bf64e6370740e0d8de75b1357f8721671af537d5d4040cb084eb4e2" & _
    "cc34d2466a0115af84e1b0042895983a1d06b89fb4ce6ea0486f3f3b823520ab82011a1d4b27722" & _
    "7f8611560b1e7933fdcbb3a792b344525bda08839e151ce794b2f32c9b7a01fbac9e01cc87ebcc7" & _
    "d1f6cf0111c3a1e8aac71a908749d44fbd9ad0dadecbd50ada380339c32ac69136678df9317ce0b" & _
    "12b4ff79e59b743f5bb3af2d519ff27d9459cbf97222c15e6fc2a0f91fc719b941525fae59361ce" & _
    "b69cebc2a8645912baa8d1b6c1075ee3056a0c10d25065cb03a442e0ec6e0e1698db3b4c98a0be3" & _
    "278e9649f1f9532e0d392dfd3a0342b8971f21e"
    
    strPi = strPi & "1b0a74414ba3348cc5be7120c37632d8df359f8d9b992f2ee60b6f470fe3f1" & _
    "1de54cda541edad891ce6279cfcd3e7e6f1618b166fd2c1d05848fd2c5f6fb2299f523f357a632" & _
    "762393a8353156cccd02acf081625a75ebb56e16369788d273ccde96629281b949d04c50901b71" & _
    "c65614e6c6c7bd327a140a45e1d006c3f27b9ac9aa53fd62a80f00bb25bfe235bdd2f671126905b" & _
    "2040222b6cbcf7ccd769c2b53113ec01640e3d338abbd602547adf0ba38209cf746ce7677afa1c5" & _
    "2075606085cbfe4e8ae88dd87aaaf9b04cf9aa7e1948c25c02fb8a8c01c36ae4d6ebe1f990d4f86" & _
    "9a65cdea03f09252dc208e69fb74e6132ce77e25b578fdfe33ac372e6"
    
    lngCount3 = 0

    For lngCount2 = 0 To 3
        For lngCount1 = 0 To 255
            udtSBox(lngCount2, lngCount1).lngByte1 = _
                "&H" & Mid(strPi, lngCount3 + 1, 2)
            udtSBox(lngCount2, lngCount1).lngByte2 = _
                "&H" & Mid(strPi, lngCount3 + 3, 2)
            udtSBox(lngCount2, lngCount1).lngByte3 = _
                "&H" & Mid(strPi, lngCount3 + 5, 2)
            udtSBox(lngCount2, lngCount1).lngByte4 = _
                "&H" & Mid(strPi, lngCount3 + 7, 2)
            
            lngCount3 = lngCount3 + 8
        Next
    Next
End Sub

Private Sub BF_Password(strPassword As String)
    Dim lngPasswordLength As Long
    Dim lngCount1 As Long
    Dim lngCount2 As Long
    Dim udtTemp32 As typWord32
    Dim udtTemp64 As typWord64
    
    Call BF_Initialise
    
    lngPasswordLength = Len(strPassword)
    
    If lngPasswordLength > 56 Then lngPasswordLength = 56
    If EXPORT And (lngPasswordLength > 5) Then lngPasswordLength = 5
    
    For lngCount1 = 0 To 17
        udtTemp32.lngByte1 = Asc(Mid(strPassword, (lngCount2 Mod _
            lngPasswordLength) + 1, 1))
        udtTemp32.lngByte2 = Asc(Mid(strPassword, ((lngCount2 + 1) Mod _
            lngPasswordLength) + 1, 1))
        udtTemp32.lngByte3 = Asc(Mid(strPassword, ((lngCount2 + 2) Mod _
            lngPasswordLength) + 1, 1))
        udtTemp32.lngByte4 = Asc(Mid(strPassword, ((lngCount2 + 3) Mod _
            lngPasswordLength) + 1, 1))
    
        udtPBox(lngCount1) = Xor32(udtPBox(lngCount1), udtTemp32)
        
        lngCount2 = (lngCount2 + 4) Mod lngPasswordLength
    Next
        
    For lngCount1 = 0 To 16 Step 2
        udtTemp64 = BF_Encrypt(udtTemp64)
        udtPBox(lngCount1) = udtTemp64.udtLow
        udtPBox(lngCount1 + 1) = udtTemp64.udtHigh
    Next
    
    For lngCount1 = 0 To 3
        For lngCount2 = 0 To 255 Step 2
            udtTemp64 = BF_Encrypt(udtTemp64)
            udtSBox(lngCount1, lngCount2) = udtTemp64.udtLow
            udtSBox(lngCount1, lngCount2 + 1) = udtTemp64.udtHigh
        Next
    Next
End Sub

Private Sub bf_R(ByRef udtA As typWord32, _
udtB As typWord32, _
lngPBox As Long)
    udtA = Xor32(udtA, Xor32(Add32(Xor32(Add32(udtSBox(0, udtB.lngByte1), _
            udtSBox(1, udtB.lngByte2)), udtSBox(2, udtB.lngByte3)), _
            udtSBox(3, udtB.lngByte4)), udtPBox(lngPBox)))
End Sub

Private Function BF_Encrypt(udtInput As typWord64) As typWord64
    Dim udtLeft As typWord32
    Dim udtRight As typWord32

    udtLeft = udtInput.udtLow
    udtRight = udtInput.udtHigh

    udtLeft = Xor32(udtLeft, udtPBox(0))
    
    Call bf_R(udtRight, udtLeft, 1)
    Call bf_R(udtLeft, udtRight, 2)
    Call bf_R(udtRight, udtLeft, 3)
    Call bf_R(udtLeft, udtRight, 4)
    Call bf_R(udtRight, udtLeft, 5)
    Call bf_R(udtLeft, udtRight, 6)
    Call bf_R(udtRight, udtLeft, 7)
    Call bf_R(udtLeft, udtRight, 8)
    Call bf_R(udtRight, udtLeft, 9)
    Call bf_R(udtLeft, udtRight, 10)
    Call bf_R(udtRight, udtLeft, 11)
    Call bf_R(udtLeft, udtRight, 12)
    Call bf_R(udtRight, udtLeft, 13)
    Call bf_R(udtLeft, udtRight, 14)
    Call bf_R(udtRight, udtLeft, 15)
    Call bf_R(udtLeft, udtRight, 16)
    
    udtRight = Xor32(udtRight, udtPBox(17))
    
    BF_Encrypt.udtLow = udtRight
    BF_Encrypt.udtHigh = udtLeft
End Function

Private Function BF_Decrypt(udtInput As typWord64) As typWord64
    Dim udtLeft As typWord32
    Dim udtRight As typWord32

    udtLeft = udtInput.udtLow
    udtRight = udtInput.udtHigh

    udtLeft = Xor32(udtLeft, udtPBox(17))
    
    Call bf_R(udtRight, udtLeft, 16)
    Call bf_R(udtLeft, udtRight, 15)
    Call bf_R(udtRight, udtLeft, 14)
    Call bf_R(udtLeft, udtRight, 13)
    Call bf_R(udtRight, udtLeft, 12)
    Call bf_R(udtLeft, udtRight, 11)
    Call bf_R(udtRight, udtLeft, 10)
    Call bf_R(udtLeft, udtRight, 9)
    Call bf_R(udtRight, udtLeft, 8)
    Call bf_R(udtLeft, udtRight, 7)
    Call bf_R(udtRight, udtLeft, 6)
    Call bf_R(udtLeft, udtRight, 5)
    Call bf_R(udtRight, udtLeft, 4)
    Call bf_R(udtLeft, udtRight, 3)
    Call bf_R(udtRight, udtLeft, 2)
    Call bf_R(udtLeft, udtRight, 1)
    
    udtRight = Xor32(udtRight, udtPBox(0))
        
    BF_Decrypt.udtLow = udtRight
    BF_Decrypt.udtHigh = udtLeft
End Function

Private Function BF_TestVectors()
    Dim strV As String
    Dim lngCount1 As Long
    Dim strPassword As String
    Dim udtPlalngext As typWord64
    Dim udtEncryptedText As typWord64
    Dim udtDecryptedText As typWord64
    
    '8 bytes each of key, plain text, and scrambled text
    strV = "000000000000000000000000000000004EF997456198DD78"
    strV = strV & "FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF51866FD5B85ECB8A"
    strV = strV & "300000000000000010000000000000017D856F9A613063F2"
    strV = strV & "111111111111111111111111111111112466DD878B963C9D"
    strV = strV & "0123456789ABCDEF111111111111111161F9C3802281B096"
    strV = strV & "11111111111111110123456789ABCDEF7D0CC630AFDA1EC7"
    strV = strV & "000000000000000000000000000000004EF997456198DD78"
    strV = strV & "FEDCBA98765432100123456789ABCDEF0ACEAB0FC6A0A28D"
    strV = strV & "7CA110454A1A6E5701A1D6D03977674259C68245EB05282B"
    strV = strV & "0131D9619DC1376E5CD54CA83DEF57DAB1B8CC0B250F09A0"
    strV = strV & "07A1133E4A0B26860248D43806F671721730E5778BEA1DA4"
    strV = strV & "3849674C2602319E51454B582DDF440AA25E7856CF2651EB"
    strV = strV & "04B915BA43FEB5B642FD443059577FA2353882B109CE8F1A"
    strV = strV & "0113B970FD34F2CE059B5E0851CF143A48F4D0884C379918"
    strV = strV & "0170F175468FB5E60756D8E0774761D2432193B78951FC98"
    strV = strV & "43297FAD38E373FE762514B829BF486A13F04154D69D1AE5"
    strV = strV & "07A7137045DA2A163BDD1190493728022EEDDA93FFD39C79"
    strV = strV & "04689104C2FD3B2F26955F6835AF609AD887E0393C2DA6E3"
    strV = strV & "37D06BB516CB7546164D5E404F2752325F99D04F5B163969"
    strV = strV & "1F08260D1AC2465E6B056E18759F5CCA4A057A3B24D3977B"
    strV = strV & "584023641ABA6176004BD6EF09176062452031C1E4FADA8E"
    strV = strV & "025816164629B007480D39006EE762F27555AE39F59B87BD"
    strV = strV & "49793EBC79B3258F437540C8698F3CFA53C55F9CB49FC019"
    strV = strV & "4FB05E1515AB73A7072D43A0770752927A8E7BFA937E89A3"
    strV = strV & "49E95D6D4CA229BF02FE55778117F12ACF9C5D7A4986ADB5"
    strV = strV & "018310DC409B26D61D9D5C5018F728C2D1ABB290658BC778"
    strV = strV & "1C587F1C13924FEF305532286D6F295A55CB3774D13EF201"
    strV = strV & "01010101010101010123456789ABCDEFFA34EC4847B268B2"
    strV = strV & "1F1F1F1F0E0E0E0E0123456789ABCDEFA790795108EA3CAE"
    strV = strV & "E0FEE0FEF1FEF1FE0123456789ABCDEFC39E072D9FAC631D"
    strV = strV & "0000000000000000FFFFFFFFFFFFFFFF014933E0CDAFF6E4"
    strV = strV & "FFFFFFFFFFFFFFFF0000000000000000F21E9A77B71C49BC"
    strV = strV & "0123456789ABCDEF0000000000000000245946885754369A"
    strV = strV & "FEDCBA9876543210FFFFFFFFFFFFFFFF6B5C5A9C5D9E0A5A"
    
    For lngCount1 = 1 To Len(strV) Step 48
        strPassword = HexToString(Mid(strV, lngCount1, 16))
        
        udtPlalngext = HexToWord64(Mid(strV, lngCount1 + 16, 16))
        
        Call BF_Password(strPassword)
        
        udtEncryptedText = BF_Encrypt(udtPlalngext)
                       
        If Word64ToHex(udtEncryptedText) <> Mid(strV, lngCount1 + 32, 16) Then
            BF_TestVectors = False
            Exit Function
        End If
                       
        udtDecryptedText = BF_Decrypt(udtEncryptedText)
        
        If Word64ToHex(udtDecryptedText) <> Word64ToHex(udtPlalngext) Then
            BF_TestVectors = False
            Exit Function
        End If
    Next

    BF_TestVectors = True
End Function

Private Function Word64ToHex(udtWord64 As typWord64) As String
    Word64ToHex = Right("00" & Hex(udtWord64.udtLow.lngByte1), 2)
    Word64ToHex = Word64ToHex & Right("00" & Hex(udtWord64.udtLow.lngByte2), 2)
    Word64ToHex = Word64ToHex & Right("00" & Hex(udtWord64.udtLow.lngByte3), 2)
    Word64ToHex = Word64ToHex & Right("00" & Hex(udtWord64.udtLow.lngByte4), 2)
    Word64ToHex = Word64ToHex & Right("00" & Hex(udtWord64.udtHigh.lngByte1), 2)
    Word64ToHex = Word64ToHex & Right("00" & Hex(udtWord64.udtHigh.lngByte2), 2)
    Word64ToHex = Word64ToHex & Right("00" & Hex(udtWord64.udtHigh.lngByte3), 2)
    Word64ToHex = Word64ToHex & Right("00" & Hex(udtWord64.udtHigh.lngByte4), 2)
End Function

Private Function HexToWord64(strHex As String) As typWord64
    HexToWord64.udtLow.lngByte1 = Asc(Chr("&h" & Mid(strHex, 1, 2)))
    HexToWord64.udtLow.lngByte2 = Asc(Chr("&h" & Mid(strHex, 3, 2)))
    HexToWord64.udtLow.lngByte3 = Asc(Chr("&h" & Mid(strHex, 5, 2)))
    HexToWord64.udtLow.lngByte4 = Asc(Chr("&h" & Mid(strHex, 7, 2)))
    HexToWord64.udtHigh.lngByte1 = Asc(Chr("&h" & Mid(strHex, 9, 2)))
    HexToWord64.udtHigh.lngByte2 = Asc(Chr("&h" & Mid(strHex, 11, 2)))
    HexToWord64.udtHigh.lngByte3 = Asc(Chr("&h" & Mid(strHex, 13, 2)))
    HexToWord64.udtHigh.lngByte4 = Asc(Chr("&h" & Mid(strHex, 15, 2)))
End Function

Private Function HexToString(strHex As String) As String
    Dim lngCount As Long
    
    For lngCount = 1 To Len(strHex) Step 2
        HexToString = HexToString & Chr("&h" & Mid(strHex, lngCount, 2))
    Next
End Function

Private Function ToHEX(tString As String) As String

Dim i As Integer, s As String
s = ""
For i = 1 To Len(tString)
  s = s & Right$("00" & Hex(Asc(Mid$(tString, i, 1))), 2)
Next i

ToHEX = s

End Function
