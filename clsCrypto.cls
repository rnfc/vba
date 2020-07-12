VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsCrypto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit

'--------------------------------------------------------------------------------
'--------------------------------------------------------------------------------
'Microsoft provides an encryption library called the CryptoAPI which is
'built into most Windows installations. Here's how to use it. It's not the
'strongest or fastest algorithm around but will serve most purposes.
'--------------------------------------------------------------------------------
'--------------------------------------------------------------------------------

Private sPassword As String
Private sInputBuffer As String
Private sOutputBuffer As String
Private vbNullChar As String

Private Declare PtrSafe Function CryptAcquireContext Lib "advapi32.dll" Alias _
    "CryptAcquireContextA" ( _
        ByRef hCryptProv As Long, _
        ByVal pszContainer As String, _
        ByVal pszProvider As String, _
        ByVal dwProvType As Long, _
        ByVal dwFlags As Long) As Long

Private Declare PtrSafe Function CryptCreateHash Lib "advapi32.dll" ( _
    ByVal hProv As Long, _
    ByVal Algid As Long, _
    ByVal hKey As Long, _
    ByVal dwFlags As Long, _
    phHash As Long) As Long

Private Declare PtrSafe Function CryptHashData Lib "advapi32.dll" ( _
    ByVal hHash As Long, _
    ByVal pbData As String, _
    ByVal dwDataLen As Long, _
    ByVal dwFlags As Long) As Long

Private Declare PtrSafe Function CryptDeriveKey Lib "advapi32.dll" ( _
    ByVal hProv As Long, _
    ByVal Algid As Long, _
    ByVal hBaseData As Long, _
    ByVal dwFlags As Long, _
    phKey As Long) As Long

Private Declare PtrSafe Function CryptDestroyHash Lib "advapi32.dll" ( _
    ByVal hHash As Long) As Long

Private Declare PtrSafe Function CryptDestroyKey Lib "advapi32.dll" ( _
    ByVal hKey As Long) As Long

Private Declare PtrSafe Function CryptEncrypt Lib "advapi32.dll" ( _
    ByVal hKey As Long, _
    ByVal hHash As Long, _
    ByVal Final As Long, _
    ByVal dwFlags As Long, _
    ByVal pbData As String, _
    pdwDataLen As Long, _
    ByVal dwBufLen As Long) As Long

Private Declare PtrSafe Function CryptDecrypt Lib "advapi32.dll" ( _
    ByVal hKey As Long, _
    ByVal hHash As Long, _
    ByVal Final As Long, _
    ByVal dwFlags As Long, _
    ByVal pbData As String, _
    pdwDataLen As Long) As Long

Private Declare PtrSafe Function CryptReleaseContext Lib "advapi32.dll" ( _
    ByVal hProv As Long, _
    ByVal dwFlags As Long) As Long

Private Declare PtrSafe Function GetLastError Lib "kernel32" () As Long
    
'constants for Cryptography API functions
Private Const CRYPT_NEWKEYSET = &H8
Private Const MS_DEF_PROV = "Microsoft Base Cryptographic Provider v1.0"
Private Const PROV_RSA_FULL = 1
Private Const ALG_CLASS_DATA_ENCRYPT = 24576
Private Const ALG_CLASS_HASH = 32768

Private Const ALG_TYPE_ANY = 0
Private Const ALG_TYPE_BLOCK = 1536
Private Const ALG_TYPE_STREAM = 2048

Private Const ALG_SID_RC2 = 2

Private Const ALG_SID_RC4 = 1
Private Const ALG_SID_MD5 = 3
Private Const CALG_MD5 = _
    ((ALG_CLASS_HASH Or ALG_TYPE_ANY) Or ALG_SID_MD5)
Private Const CALG_RC2 = _
    ((ALG_CLASS_DATA_ENCRYPT Or ALG_TYPE_BLOCK) Or ALG_SID_RC2)
Private Const CALG_RC4 = _
    ((ALG_CLASS_DATA_ENCRYPT Or ALG_TYPE_STREAM) Or ALG_SID_RC4)

Private Const ENCRYPT_ALGORITHM = CALG_RC4
Private Const ENCRYPT_BLOCK_SIZE = 1

Private Const CRYPT_EXPORTABLE = 1


Public Property Get password() As String
password = sPassword
End Property

Public Property Let password(ByVal vNewValue As String)
sPassword = vNewValue
End Property

Public Property Get InBuffer() As String
InBuffer = sInputBuffer
End Property

Public Property Let InBuffer(ByVal vNewValue As String)
sInputBuffer = vNewValue
End Property

Public Property Get OutBuffer() As String
OutBuffer = sOutputBuffer
End Property

Public Property Let OutBuffer(ByVal vNewValue As String)
sOutputBuffer = vNewValue
End Property

Public Sub Encrypt()
Call CryptoEncrypt
End Sub

Public Sub Decrypt()
Call CryptoDecrypt
End Sub

Private Sub CryptoEncrypt()

Dim sProvider As String
Dim lHHash As Long
Dim lHkey As Long
Dim lResult As Long
Dim lHExchgKey As Long
Dim lHCryptprov As Long

Dim lCryptLength As Long
Dim lCryptBufLen As Long
Dim sCryptBuffer As String

On Error GoTo EncryptError

vbNullChar = Chr$(0)

'Get handle to the default CSP
sProvider = MS_DEF_PROV & Chr$(0)
If Not CBool(CryptAcquireContext(lHCryptprov, "Excel", sProvider, PROV_RSA_FULL, 0)) Then
    'If there is no default key container then create one using Flags field
    If GetLastError = 0 Then
        If Not CBool(CryptAcquireContext(lHCryptprov, "Excel", sProvider, PROV_RSA_FULL, CRYPT_NEWKEYSET)) Then
            Debug.Print ("Error " & CStr(GetLastError) & " during CryptAcquireContext!")
            GoTo Finished
        End If
    End If
End If

'Create a hash object
If Not CBool(CryptCreateHash(lHCryptprov, CALG_MD5, 0, 0, lHHash)) Then
    Debug.Print ("Error " & CStr(GetLastError) & " during CryptCreateHash!")
    GoTo Finished
End If
 
'Hash in the password text
If Not CBool(CryptHashData(lHHash, sPassword, Len(sPassword), 0)) Then
    Debug.Print ("Error " & CStr(GetLastError) & " during CryptHashData!")
    GoTo Finished
End If
 
'Create a session key from the hash object.
If Not CBool(CryptDeriveKey(lHCryptprov, ENCRYPT_ALGORITHM, lHHash, 0, lHkey)) Then
    Debug.Print ("Error " & CStr(GetLastError) & " during CryptDeriveKey!")
    GoTo Finished
End If

'Destroy the hash object.
CryptDestroyHash (lHHash)
lHHash = 0

'Create a buffer for the CryptEncrypt function
lCryptLength = Len(sInputBuffer)
lCryptBufLen = lCryptLength * 2
sCryptBuffer = String(lCryptBufLen, vbNullChar)
LSet sCryptBuffer = sInputBuffer

'Encrypt the text data
If Not CBool(CryptEncrypt(lHkey, 0, 1, 0, sCryptBuffer, lCryptLength, lCryptBufLen)) Then
    Debug.Print ("bytes required:" & CStr(lCryptLength))
    Debug.Print ("Error " & CStr(GetLastError) & " during CryptEncrypt!")
End If

sOutputBuffer = Mid$(sCryptBuffer, 1, lCryptLength)
    
Finished:

'Destroy session key.
If (lHkey) Then lResult = CryptDestroyKey(lHkey)

'Destroy key exchange key handle
If lHExchgKey Then CryptDestroyKey (lHExchgKey)

'Destroy hash object
If lHHash Then CryptDestroyHash (lHHash)

'Release Context provider handle
If lHCryptprov Then lResult = CryptReleaseContext(lHCryptprov, 0)

Exit Sub

EncryptError:

Debug.Print ("Encrypt Error: " & error$())

GoTo Finished

End Sub

Private Sub CryptoDecrypt()
Dim lHExchgKey As Long
Dim lHCryptprov As Long
Dim lHHash As Long
Dim lHkey As Long
Dim lResult As Long

Dim sProvider As String

Dim sCryptBuffer As String
Dim lCryptBufLen As Long
    
On Error GoTo DecryptError

vbNullChar = Chr$(0)

'Clear the Output buffer
sOutputBuffer = ""

'Get handle to the default CSP.
sProvider = vbNullChar
sProvider = MS_DEF_PROV & vbNullChar
If Not CBool(CryptAcquireContext(lHCryptprov, "Excel", sProvider, PROV_RSA_FULL, 0)) Then
    Debug.Print ("Error " & CStr(GetLastError) & " during CryptAcquireContext!")
    GoTo Finished
End If

'Create a hash object
If Not CBool(CryptCreateHash(lHCryptprov, CALG_MD5, 0, 0, lHHash)) Then
    Debug.Print ("Error " & CStr(GetLastError) & " during CryptCreateHash!")
    GoTo Finished
End If

'Hash in the password text
If Not CBool(CryptHashData(lHHash, sPassword, Len(sPassword), 0)) Then
    Debug.Print ("Error " & CStr(GetLastError) & " during CryptHashData!")
    GoTo Finished
End If
    
'Create a session key from the hash object
If Not CBool(CryptDeriveKey(lHCryptprov, ENCRYPT_ALGORITHM, lHHash, 0, lHkey)) Then
    Debug.Print ("Error " & CStr(GetLastError) & " during CryptDeriveKey!")
    GoTo Finished
End If

'Destroy the hash object.
CryptDestroyHash (lHHash)
lHHash = 0

'Prepare sCryptBuffer for CryptDecrypt
lCryptBufLen = Len(sInputBuffer) * 2
sCryptBuffer = String(lCryptBufLen, vbNullChar)
LSet sCryptBuffer = sInputBuffer

'Decrypt data
If Not CBool(CryptDecrypt(lHkey, 0, 1, 0, sCryptBuffer, lCryptBufLen)) Then
    Debug.Print ("bytes required:" & CStr(lCryptBufLen))
    Debug.Print ("Error " & CStr(GetLastError) & " during CryptDecrypt!")
    GoTo Finished
End If

'Setup output buffer with just decrypted data
sOutputBuffer = Mid$(sCryptBuffer, 1, lCryptBufLen / 2)

Finished:

'Destroy session key
If (lHkey) Then lResult = CryptDestroyKey(lHkey)

'Destroy key exchange key handle
If lHExchgKey Then CryptDestroyKey (lHExchgKey)

'Destroy hash object
If lHHash Then CryptDestroyHash (lHHash)

'Release Context provider handle
If lHCryptprov Then lResult = CryptReleaseContext(lHCryptprov, 0)

Exit Sub

DecryptError:
Debug.Print ("Decrypt Error: " & error$())
GoTo Finished

End Sub
