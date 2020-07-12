Attribute VB_Name = "WEB_CONNECTION_CACHE_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.

Private Const ERROR_CACHE_FIND_FAIL As Long = 0
Private Const ERROR_CACHE_FIND_SUCCESS As Long = 1
Private Const ERROR_FILE_NOT_FOUND As Long = 2
Private Const ERROR_ACCESS_DENIED As Long = 5
Private Const ERROR_INSUFFICIENT_BUFFER As Long = 122

Private Const MAX_PATH As Long = 260
Private Const MAX_CACHE_ENTRY_INFO_SIZE As Long = 4096

Private Const LMEM_FIXED As Long = &H0
Private Const LMEM_ZEROINIT As Long = &H40
Private Const LPTR As Long = (LMEM_FIXED Or LMEM_ZEROINIT)

Private Const NORMAL_CACHE_ENTRY As Long = &H1
Private Const EDITED_CACHE_ENTRY As Long = &H8
Private Const TRACK_OFFLINE_CACHE_ENTRY As Long = &H10
Private Const TRACK_ONLINE_CACHE_ENTRY As Long = &H20
Private Const STICKY_CACHE_ENTRY As Long = &H40
Private Const SPARSE_CACHE_ENTRY As Long = &H10000
Private Const COOKIE_CACHE_ENTRY As Long = &H100000
Private Const URLHISTORY_CACHE_ENTRY As Long = &H200000
Private Const URLCACHE_FIND_DEFAULT_FILTER As Long = NORMAL_CACHE_ENTRY Or _
                                                    COOKIE_CACHE_ENTRY Or _
                                                    URLHISTORY_CACHE_ENTRY Or _
                                                    TRACK_OFFLINE_CACHE_ENTRY Or _
                                                    TRACK_ONLINE_CACHE_ENTRY Or _
                                                    STICKY_CACHE_ENTRY
Private Type FILETIME
   dwLowDateTime As Long
   dwHighDateTime As Long
End Type

Private Type INTERNET_CACHE_ENTRY_INFO
   dwStructSize As Long
   lpszSourceUrlName As Long
   lpszLocalFileName As Long
   CacheEntryType  As Long
   dwUseCount As Long
   dwHitRate As Long
   dwSizeLow As Long
   dwSizeHigh As Long
   LastModifiedTime As FILETIME
   ExpireTime As FILETIME
   LastAccessTime As FILETIME
   LastSyncTime As FILETIME
   lpHeaderInfo As Long
   dwHeaderInfoSize As Long
   lpszFileExtension As Long
   dwExemptDelta  As Long
End Type

'------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------

'To use the Find URL APIs, a call is made passing a type of 0 length. The call
'fails with ERROR_INSUFFICIENT_BUFFER, and the APIs dwBuffer member contains the
'size of the structure needed to retrieve the data. Memory is allocated to this
'buffer using LocalAlloc, and the pointer to this memory is passed in place of
'the structure (a non-typical use of UDT-based APIs to VB programmers). One the
'call succeeds, CopyMemory is used again to move the memory buffer into a VB
'user-defined type, and then the string information is extracted using lstrlenA
'and lstrcpyA. Finally the memory pointer is freed with LocalFree.   This process
'is repeated for each cache entry while the handle to the Find method remains
'valid. Once the last entry has been retrieved, FindNextUrlCacheEntry returns 0,
'and the loop terminates.
   
Private Declare PtrSafe Function FindFirstUrlCacheEntry Lib "wininet" _
   Alias "FindFirstUrlCacheEntryA" _
  (ByVal lpszUrlSearchPattern As String, _
   lpFirstCacheEntryInfo As Any, _
   lpdwFirstCacheEntryInfoBufferSize As Long) As Long

Private Declare PtrSafe Function FindNextUrlCacheEntry Lib "wininet" _
   Alias "FindNextUrlCacheEntryA" _
  (ByVal hEnumHandle As Long, _
   lpNextCacheEntryInfo As Any, _
   lpdwNextCacheEntryInfoBufferSize As Long) As Long

Private Declare PtrSafe Function FindCloseUrlCache Lib "wininet" _
   (ByVal hEnumHandle As Long) As Long

Private Declare PtrSafe Function DeleteUrlCacheEntry Lib "wininet" _
   Alias "DeleteUrlCacheEntryA" _
  (ByVal lpszUrlName As String) As Long
   
   
'------------------------------------------------------------------------------
'------------------------------------------------------------------------------
'------------------------------------------------------------------------------

'While, as VB developers, we are familiar with passing C structures as VB
'user-defined types to API methods, these structures usually contained all
'the data returned, and so were defined with strings of fixed-lengths. The
'URL cache APIs however are variable-size structures - each call will result
'in data types differing in size - so the application must allocate sufficient
'data space for the successful call prior to calling. This involves using
'LocalAlloc, LocalFree, as well as CopyMemory to set up pointers to the allocated
'memory space.

Private Declare PtrSafe Sub CopyMemory Lib "kernel32" _
   Alias "RtlMoveMemory" _
   (pDest As Any, _
    pSource As Any, _
    ByVal dwLength As Long)

Private Declare PtrSafe Function lstrcpyA Lib "kernel32" _
  (ByVal retval As String, ByVal Ptr As Long) As Long
                        
Private Declare PtrSafe Function lstrlenA Lib "kernel32" _
  (ByVal Ptr As Any) As Long
  
Private Declare PtrSafe Function LocalAlloc Lib "kernel32" _
   (ByVal uFlags As Long, _
    ByVal uBytes As Long) As Long
    
Private Declare PtrSafe Function LocalFree Lib "kernel32" _
   (ByVal hMem As Long) As Long


Function COUNT_WEB_CACHE_FUNC(Optional ByRef VERSION As Long = 0) As Long
    
   Dim ii As Long
   Dim jj As Long 'hFile
   Dim kk As Long 'dwBuffer
   Dim ll As Long 'pntrICE
   
   Dim CACHE_TYPE As Long
   Dim CACHE_ARR As Variant
   Dim CACHE_FILE_STR As String
   Dim ICEI_OBJ As INTERNET_CACHE_ENTRY_INFO
   
   On Error GoTo ERROR_LABEL
   
   Select Case VERSION
        Case 0 '"Normal Entry"
      CACHE_TYPE = &H1
        Case 1 '"Edited Entry (IE5)"
      CACHE_TYPE = &H8
        Case 2 '"Offline Entry"
      CACHE_TYPE = &H10
        Case 3 '"Online Entry"
      CACHE_TYPE = &H20
        Case 4 '"Stick Entry"
      CACHE_TYPE = &H40
        Case 5 '"Sparse Entry (n/a)"
      CACHE_TYPE = &H10000
        Case 6 '"Cookies"
      CACHE_TYPE = &H100000
        Case 7 '"Visited History"
      CACHE_TYPE = &H200000
        Case Else '"Default Filter"
      CACHE_TYPE = URLCACHE_FIND_DEFAULT_FILTER
    End Select
    
   ii = 1
   ReDim CACHE_ARR(1 To ii)
  'Like other APIs, calling FindFirstUrlCacheEntry or
  'FindNextUrlCacheEntry with an insufficient buffer will
  'cause the API to fail, and its kk points to the
  'correct size required for a successful call.
   kk = 0
   
  'Call to determine the required buffer size
   jj = FindFirstUrlCacheEntry(vbNullString, ByVal 0, kk)
   
  'both conditions should be met by the first call
   If (jj = ERROR_CACHE_FIND_FAIL) And _
      (Err.LastDllError = ERROR_INSUFFICIENT_BUFFER) Then
   
     'The INTERNET_CACHE_ENTRY_INFO data type is a
     'variable-length type. It is necessary to allocate
     'memory for the result of the call and pass the
     'pointer to this memory location to the API.
      ll = LocalAlloc(LMEM_FIXED, kk)
        
     'allocation successful
      If ll Then
         
        'set a Long pointer to the memory location
         CopyMemory ByVal ll, kk, 4
         
        'and call the first find API again passing the
        'pointer to the allocated memory
         jj = FindFirstUrlCacheEntry(vbNullString, ByVal ll, kk)
       
        'jj should = 1 (success)
         If jj <> ERROR_CACHE_FIND_FAIL Then
         
           'now just loop through the cache
            Do
            
              'the pointer has been filled, so move the
              'data back into a ICEI_OBJ structure
               CopyMemory ICEI_OBJ, ByVal ll, Len(ICEI_OBJ)
            
              'CacheEntryType is a long representing the type of
              'entry returned, and should match our passed param.
               If (ICEI_OBJ.CacheEntryType And CACHE_TYPE) Then
               
                  'extract the string from the memory location
                  'pointed to by the lpszSourceUrlName member
                  'and add to a list
                   CACHE_FILE_STR = String$(lstrlenA(ByVal _
                                    ICEI_OBJ.lpszSourceUrlName), 0)
                   Call lstrcpyA(ByVal CACHE_FILE_STR, _
                        ByVal ICEI_OBJ.lpszSourceUrlName)

                   ReDim Preserve CACHE_ARR(1 To ii)
                   CACHE_ARR(ii) = CACHE_FILE_STR
                   ii = ii + 1
               End If
               
              'free the pointer and memory associated
              'with the last-retrieved file
               Call LocalFree(ll)
               
              'and again repeat the procedure, this time calling
              'FindNextUrlCacheEntry with a buffer size set to 0.
              'This will cause the call to once again fail,
              'returning the required size as kk
               kk = 0
               Call FindNextUrlCacheEntry(jj, ByVal 0, kk)
               
              'allocate and assign the memory to the pointer
               ll = LocalAlloc(LMEM_FIXED, kk)
               CopyMemory ByVal ll, kk, 4
               
           'and call again with the valid parameters.
           'If the call fails (no more data), the loop exits.
           'If the call is successful, the Do portion of the
           'loop is executed again, extracting the data from
           'the returned type
            Loop While FindNextUrlCacheEntry(jj, ByVal ll, kk)
  
         End If 'jj
         
      End If 'll
   
   End If 'jj
   

  'clean up by closing the find handle, as
  'well as calling LocalFree again to be safe
   Call LocalFree(ll)
   Call FindCloseUrlCache(jj)
   
   COUNT_WEB_CACHE_FUNC = ii
   
Exit Function
ERROR_LABEL:
COUNT_WEB_CACHE_FUNC = Err.number
End Function

'UrlDownloadToFile hauls the file down to the IE cache, then
'copies the file to the local or remote file system destination you specify.
'Because it uses the cache, subsequent calls to this API will pull the cached
'copy, not the copy off the site. Where assurance is needed that the calls
'always pulling from the site, you need to delete the cached copy of the
'file first.

Function LOOK_WEB_CACHE_FUNC(ByVal URL_STR_NAME As String)

Dim i As Long
Dim TEMP_ARR As Variant
Dim TEMP_FLAG As Boolean

On Error GoTo ERROR_LABEL

LOOK_WEB_CACHE_FUNC = False

TEMP_FLAG = LOAD_WEB_CACHE_ARR_FUNC(TEMP_ARR)

If TEMP_FLAG = False Then: GoTo ERROR_LABEL
For i = LBound(TEMP_ARR) To UBound(TEMP_ARR)
    If TEMP_ARR(i) = URL_STR_NAME Then
        LOOK_WEB_CACHE_FUNC = True
        Exit Function
    End If
Next i

Exit Function
ERROR_LABEL:
LOOK_WEB_CACHE_FUNC = False
End Function

Function REMOVE_WEB_PAGE_CACHE_FUNC(ByVal SRC_URL_STR As String)

On Error GoTo ERROR_LABEL

'Attempt to delete any cached version of the file. Since we're only interested in
'nuking the file, the routine is called as a sub. If the return value is requires
'(calling as a function). Note that the remote URL is passed as this is the
'name the cached file is known by. This does NOT delete the file from the remote
'server.
  
    If DeleteUrlCacheEntry(SRC_URL_STR) = 1 Then
       REMOVE_WEB_PAGE_CACHE_FUNC = True 'cached file found and deleted
    Else
       REMOVE_WEB_PAGE_CACHE_FUNC = False 'no cached file for SRC_URL_STR
    End If

Exit Function
ERROR_LABEL:
REMOVE_WEB_PAGE_CACHE_FUNC = False
End Function

Sub REMOVE_CACHE_HISTORY_FUNC()

Dim i As Long
Dim TEMP_ARR As Variant
Dim TEMP_FLAG As Boolean

On Error GoTo ERROR_LABEL

'REMOVE_CACHE_HISTORY_FUNC = False
TEMP_FLAG = LOAD_WEB_CACHE_ARR_FUNC(TEMP_ARR)
If TEMP_FLAG = False Then: GoTo ERROR_LABEL

For i = LBound(TEMP_ARR) To UBound(TEMP_ARR)
    Call REMOVE_WEB_PAGE_CACHE_FUNC(TEMP_ARR(i))
Next i

'REMOVE_CACHE_HISTORY_FUNC = True

Exit Sub
ERROR_LABEL:
'REMOVE_CACHE_HISTORY_FUNC = False
End Sub


Function LOAD_WEB_CACHE_ARR_FUNC(ByRef CACHE_ARR As Variant)
    
'-----------------------------------------------------------------------------
'TEST:
'-----------------------------------------------------------------------------
'   Dim TEMP_ARR As Variant
'   Call REMOVE_CACHE_HISTORY_FUNC
'   Debug.Print LOAD_WEB_CACHE_ARR_FUNC(TEMP_ARR)
'   Debug.Print TEMP_ARR(1)
'-----------------------------------------------------------------------------
    
   Dim ii As Long
   Dim jj As Long 'hFile
   Dim kk As Long 'dwBuffer
   Dim ll As Long 'pntrICE
   
   Dim CACHE_FILE_STR As String
   Dim ICEI_OBJ As INTERNET_CACHE_ENTRY_INFO
   
   On Error GoTo ERROR_LABEL
   
   ii = 1
   ReDim CACHE_ARR(1 To ii)
  'Like other APIs, calling FindFirstUrlCacheEntry or
  'FindNextUrlCacheEntry with an buffer of insufficient
  'size will cause the API to fail. Call first to
  'determine the required buffer size.
   jj = FindFirstUrlCacheEntry(0&, ByVal 0, kk)
   
  'both conditions should be met by the first call
   If (jj = ERROR_CACHE_FIND_FAIL) And _
      (Err.LastDllError = ERROR_INSUFFICIENT_BUFFER) Then
   
     'The INTERNET_CACHE_ENTRY_INFO data type
     'is a variable-length UDT. It is therefore
     'necessary to allocate memory for the result
     'of the call and to pass a pointer to this
     'memory location to the API.
      ll = LocalAlloc(LMEM_FIXED, kk)
        
     'allocation successful
      If ll <> 0 Then
         
        'set a Long pointer to the memory location
         CopyMemory ByVal ll, kk, 4
         
        'call FindFirstUrlCacheEntry again
        'now passing the pointer to the
        'allocated memory
         jj = FindFirstUrlCacheEntry(vbNullString, _
                                        ByVal ll, _
                                        kk)
       
        'jj should = 1 (success)
         If jj <> ERROR_CACHE_FIND_FAIL Then
         
           'loop through the cache
            Do
            
              'the pointer has been filled, so move the
              'data back into a ICEI_OBJ structure
               CopyMemory ICEI_OBJ, ByVal ll, Len(ICEI_OBJ)
            
              'CacheEntryType is a long representing
              'the type of entry returned
               If (ICEI_OBJ.CacheEntryType And _
                   NORMAL_CACHE_ENTRY) = NORMAL_CACHE_ENTRY Then
               
                 'extract the string from the memory location
                 'pointed to by the lpszSourceUrlName member
                 'and add to a list
                  CACHE_FILE_STR = _
                  String$(lstrlenA(ByVal ICEI_OBJ.lpszSourceUrlName), 0)
                  Call lstrcpyA(ByVal CACHE_FILE_STR, _
                  ByVal ICEI_OBJ.lpszSourceUrlName)
                  ReDim Preserve CACHE_ARR(1 To ii)
                  CACHE_ARR(ii) = CACHE_FILE_STR
                  ii = ii + 1
               End If
               
              'free the pointer and memory associated
              'with the last-retrieved file
               Call LocalFree(ll)
               
              'and again by repeating the procedure but
              'now calling FindNextUrlCacheEntry. Again,
              'the buffer size set to 0 causing the call
              'to fail and return the required size as kk
               kk = 0
               Call FindNextUrlCacheEntry(jj, ByVal 0, kk)
               
              'allocate and assign the memory to the pointer
               ll = LocalAlloc(LMEM_FIXED, kk)
               CopyMemory ByVal ll, kk, 4
               
           'and call again with the valid parameters.
           'If the call fails (no more data), the loop exits.
           'If the call is successful, the Do portion of the
           'loop is executed again, extracting the data from
           'the returned type
            Loop While FindNextUrlCacheEntry(jj, ByVal ll, kk)
  
         End If 'jj
         
      End If 'll
   
   End If 'jj
   
  'clean up by closing the find handle,
  'as well as calling LocalFree again
  'to be safe
   Call LocalFree(ll)
   Call FindCloseUrlCache(jj)
   
   LOAD_WEB_CACHE_ARR_FUNC = IIf((IsEmpty(CACHE_ARR(1)) = False), True, False)
   
Exit Function
ERROR_LABEL:
LOAD_WEB_CACHE_ARR_FUNC = False
End Function
