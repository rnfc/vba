VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsPrngKci"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 0       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.

' ***************************************************************************
' Module:        PrngKci
'
' Description:   This random number generator originally appeared in "Toward
'                a Universal Random Number Generator" by George Marsaglia and
'                Arif Zaman.
'                Florida State University Report:  FSU-SCRI-87-50 (1987)
'
'                It was later modified by F. James and published in "A Review
'                of Pseudo-Random Number Generators"
'
'                It passes ALL of the tests for random number generators and
'                has a period of 2^144, is completely portable (gives bit
'                identical results on all machines with at least 24-bit
'                mantissas in the floating point representation).
'
'                The algorithm is a combination of a Fibonacci sequence (with
'                randomly selected lags ([17,5] [33,13] [39,14] [52,24] [63,31]
'                [73,25] [97,33] [607,273]) and operation "subtraction plus one,
'                modulo one" and an "arithmetic sequence" (using subtraction)
'
'                The C language version was written by Jim Butler and was based
'                on a FORTRAN program posted by David LaSalle of Florida State
'                University.
'
'                NOTE:  The seed values are:
'
'                       lngSeed1 = 1 to 31328
'                       lngSeed2 = 1 to 30081
'
'                The random number sequences created by these two seeds are of
'                sufficient length to complete an entire calculation with. For
'                example, if several different groups are working on different
'                parts of the same calculation, each group could be assigned
'                its own lngSeed1 seed. This would leave each group with 30,000
'                choices for the second seed. That is to say, this random number
'                generator can create 900 million different subsequences--with
'                each subsequence having a length of approximately 10^30.
'
'                Converted from C to VB by Kenneth Ives kenaso@home.com
'                and made into a DLL.
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 10-May-2001  Kenneth Ives  kenaso@home.com
'              Module created
' ***************************************************************************

' ---------------------------------------------------------------------------
' Define module level variables
' ---------------------------------------------------------------------------
  Private m_blnTestMode     As Boolean
  Private m_blnRandomize    As Boolean
  Private m_lngIndex1       As Long
  Private m_lngIndex2       As Long
  Private m_lngLag1         As Long
  Private m_lngLag2         As Long
  Private m_strErrorMsg     As String
  Private m_C1              As Double
  Private m_C2              As Double
  Private m_C3              As Double
  Private m_dblMedianValue  As Double
  'Private m_dblRnd          As Double
  Private m_adblHold()      As Double
  
  Private Const TEST_SEED1  As Long = 1802
  Private Const TEST_SEED2  As Long = 9373
  'Private Const A_LONG      As Long = 2147483647
  

' ---------------------------------------------------------------------------
' Declares
' ---------------------------------------------------------------------------
  ' The GetTickCount() API will capture the time in milliseconds.
  ' The counter overflows after 1192.8 hours (49.7 days) from the
  ' last reboot.
  Private Declare PtrSafe Function GetTickCount Lib "kernel32" () As Long

' ---------------------------------------------------------------------------
' CopyMemory moves the contents of a portion of memory from one location
' to another. The two locations are identified by pointers to the memory
' addresses. After the copy, the original contents in the source are set
' to zeros.
'
' Useful whenever you want to move a block of bytes between two memory
' locations.  When the source or the destination is an array of numbers
' (or of UDTs that contains only numeric and fixed-length strings), you
' must pass the first element of the array by reference.  Example below
' depicts zero based arrays.
'
' Copy the first 1000 elements of array a() to b().  Both arrays must be
' of the same type, and cannot be objects or variable-length strings.
'
'    CopyMemory b(0), a(0), 1000 * Len(a(0))
'
' ---------------------------------------------------------------------------
'Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
          (Dest As Any, source As Any, ByVal bytes As Long)

'Public Property Let Randomize(ByVal blnData As Boolean)

' ---------------------------------------------------------------------------
' used when assigning a value to the property, on the left side of an
' assignment.     Syntax: X.Randomize = TRUE
' ---------------------------------------------------------------------------
'  m_blnRandomize = blnData

'End Property

Public Property Get ErrorMsg() As String

' ---------------------------------------------------------------------------
' used when retrieving value of a property, on the right side of an
' assignment.     Syntax: Debug.Print X.ErrorMsg
' ---------------------------------------------------------------------------
  ErrorMsg = m_strErrorMsg

End Property
                  
Public Property Get MedianValue() As Double

' ---------------------------------------------------------------------------
' Define local variables
' ---------------------------------------------------------------------------
  Dim adblNumber()  As Double
  Dim dblCount      As Double
  Dim lngIndex      As Long
  
' ---------------------------------------------------------------------------
' Create 100,000 random numbers.
' ---------------------------------------------------------------------------
  dblCount = 0
  m_blnRandomize = True
  RndArray adblNumber(), 100000
  
' ---------------------------------------------------------------------------
' Loop thru and total all the numbers
' ---------------------------------------------------------------------------
  For lngIndex = 1 To 100000
      dblCount = dblCount + adblNumber(lngIndex)
  Next
  
' ---------------------------------------------------------------------------
' Get the median value
' ---------------------------------------------------------------------------
  MedianValue = dblCount / 100000

End Property

Public Property Let TestMode(ByVal blnData As Boolean)

' ---------------------------------------------------------------------------
' used when assigning a value to the property, on the left side of an
' assignment.    Syntax: X.TestMode = True
' ---------------------------------------------------------------------------
  m_blnTestMode = blnData

End Property

Public Property Get TestResults() As String

' ---------------------------------------------------------------------------
' Define local variables
' ---------------------------------------------------------------------------
  Dim lngIndex      As Long
  Dim strTestData   As String
  Dim strOutput     As String
  Dim strTemp       As String
  Dim adblData()    As Double
  
  Const dblValue    As Double = 4096#
    
' ---------------------------------------------------------------------------
' Intialize variables
' ---------------------------------------------------------------------------
  strTemp = ""
  strOutput = "Testing calculation output using predetermined seeds." & vbCrLf
  strOutput = strOutput & "If the random number generator is working properly, " & vbCrLf
  strOutput = strOutput & "the next six generated numbers should be:" & vbCrLf & vbCrLf
  strOutput = strOutput & "11124796.0   1453581.0   2764514.0" & vbCrLf
  strOutput = strOutput & " 9822910.0  11242638.0  11257225.0" & vbCrLf & vbCrLf
  strOutput = strOutput & "Actual output is:  " & vbCrLf & vbCrLf
  
' ---------------------------------------------------------------------------
' Generate first the 20,000 initializing numbers and then just 6 more
' ---------------------------------------------------------------------------
  If m_blnTestMode Then
      If PreloadRnd Then
          FillArray adblData(), 6           ' generate 6 random numbers
  
          For lngIndex = 1 To 3
              strTestData = Format$((dblValue * dblValue * adblData(lngIndex)), "0.0")
              strTemp = strTemp & strTestData & "   "
          Next
  
          strOutput = strOutput & strTemp & vbCrLf
          strTemp = " "
  
          For lngIndex = 4 To 6
              strTestData = Format$((dblValue * dblValue * adblData(lngIndex)), "0.0")
              strTemp = strTemp & strTestData & "  "
          Next
  
          strOutput = strOutput & strTemp
      Else
          strOutput = ""
      End If
  Else
      strOutput = ""
      m_strErrorMsg = "PrngKci.TestResults" & vbCrLf
      m_strErrorMsg = m_strErrorMsg & "TestMode switch must be set to TRUE."
  End If
  
' ---------------------------------------------------------------------------
' Return the test results in string format
' ---------------------------------------------------------------------------
  TestResults = strOutput
  
End Property

Public Function Rnd(ByVal sngLow As Single, _
                    Optional ByVal sngHigh As Single = 0) As Double

' ***************************************************************************
' Routine:       Rnd
'
' Description:   A random value will be generated based on the number passed
'
' Parameters:    snglow  - low Value to be randomized
'                sngHigh - High value to be the max range if two values are
'                          passed.
'
' Returns:       Double precision number
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 13-May-2001  Kenneth Ives  kenaso@home.com
'              Routine created
' 19-May-2001  Kenneth Ives  kenaso@home.com
'              Fixed logic flaw for negative testing
' ***************************************************************************

' ---------------------------------------------------------------------------
' Define local variables
' ---------------------------------------------------------------------------
  Dim adblData()  As Double
  Dim blnSingle   As Boolean
  Dim sngTemp     As Single
  
' ---------------------------------------------------------------------------
' Initialize variables
' ---------------------------------------------------------------------------
  m_strErrorMsg = ""
  
' ---------------------------------------------------------------------------
' if negative then leave
' ---------------------------------------------------------------------------
  If sngLow < 0 Or sngHigh < 0 Then
      m_strErrorMsg = "PrngKci.Rnd" & vbCrLf
      m_strErrorMsg = m_strErrorMsg & "Value(s) requested must be a positive number."
      Rnd = 0
      Exit Function
  End If
  
' ---------------------------------------------------------------------------
' if all zeros then leave
' ---------------------------------------------------------------------------
  If sngLow <= 0 And sngHigh <= 0 Then
      m_strErrorMsg = "PrngKci.Rnd" & vbCrLf
      m_strErrorMsg = m_strErrorMsg & "Value(s) requested must be a positive number."
      Rnd = 0
      Exit Function
  End If
  
' ---------------------------------------------------------------------------
' Evaluate parameters
' ---------------------------------------------------------------------------
  If (sngLow >= 1 And sngHigh < 1) Then
      blnSingle = True
  
  ElseIf (sngLow < 1 And sngHigh >= 1) Then
      ' swap places
      sngLow = sngHigh
      sngHigh = 0
      blnSingle = True
      
  ElseIf (sngLow >= 1 And sngHigh >= 1) Then
      If sngLow > sngHigh Then
          ' Swap places
          sngTemp = sngLow
          sngLow = sngHigh
          sngHigh = sngTemp
      End If
      blnSingle = False
  End If
  
' ---------------------------------------------------------------------------
' generate a single random value
' ---------------------------------------------------------------------------
  FillArray adblData(), 1
 
' ---------------------------------------------------------------------------
' was one number passed or two?
' ---------------------------------------------------------------------------
  If blnSingle Then
      ' if only one number was passed then generate a number
      ' from 0 to n, inclusive
      Rnd = (adblData(1) * sngLow) '+ 1
  Else
      ' generate a value that falls between two numbers, inclusive
      Rnd = (((sngHigh + 1) - sngLow) * adblData(1)) + sngLow
  End If
  
End Function
                     
Public Function RndArray(ByVal lngNbrOfElements As Long, adblData() As Double)

' ***************************************************************************
' Routine:       RndArray
'
' Description:   If one number is passed "RndArray(10)", a array with 10
'                elements of random generated data will be returned.
'
' Parameters:    lngNbrOfElements - Number of elements in array
'
' Returns:       an array of double precision values
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 13-May-2001  Kenneth Ives  kenaso@home.com
'              Routine created
' ***************************************************************************
  
' ---------------------------------------------------------------------------
' Initialize variables
' ---------------------------------------------------------------------------
  m_strErrorMsg = ""
  
' ---------------------------------------------------------------------------
' If less than 1 elements, then return nothing.
' ---------------------------------------------------------------------------
  If lngNbrOfElements < 1 Then
      m_strErrorMsg = "PrngKci.RndArray" & vbCrLf
      m_strErrorMsg = m_strErrorMsg & "Number of elements requested must exceed zero."
      
      ' return a negative value to indicate something went wrong
      ReDim adblData(1)
      adblData(1) = -1
      Exit Function
  End If
  
' ---------------------------------------------------------------------------
' Prepare the generator and load the array
' ---------------------------------------------------------------------------
  If PreloadRnd Then
      FillArray lngNbrOfElements, adblData()
  Else
      ' return a negative value to indicate something went wrong
      ReDim adblData(1)
      adblData(1) = -1
  End If

End Function
                     
Private Function CalcSeedValue(lngMaxValue As Long) As Long

' ***************************************************************************
' Routine:       CalcSeedValue
'
' Description:   Routine will detemine the seed values needed to perform
'                the random number calculations.  Use the current system
'                millisecond count to perform random selection.
'
' Returns:       A long integer
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 20-May-2001  Kenneth Ives  kenaso@home.com
'              Routine created
' 05-JUN-2001  Michael Meelix  michael@meelix.com
'              Suggested that I convert the captured time to binary and then
'              back to a long integer.
' ***************************************************************************

' ---------------------------------------------------------------------------
' Define local variables
' ---------------------------------------------------------------------------
  Dim lngTemp  As Long
  Dim dblTemp  As Double
  Dim strTemp  As String
      
' ---------------------------------------------------------------------------
' Initialize variables
' ---------------------------------------------------------------------------
  DoEvents        ' allow other system processes to perform
  dblTemp = 0#    ' reset hiolding area to zero
      
' ---------------------------------------------------------------------------
' capture the last 5 digits and convert to a double precision number and
' make sure the values exceeds zero
' ---------------------------------------------------------------------------
  While dblTemp < 1
      lngTemp = GetTickCount() + lngMaxValue    ' capture time in milliseconds
      strTemp = CStr(lngTemp)                   ' convert to string data
      dblTemp = CDbl(Val(Right$(strTemp, 5)))   ' capture last 5 digits and
  Wend                                          '    convert to a double
      
' ---------------------------------------------------------------------------
' loop until the captured data has been manipulated to fall between one and
' the maximum value passed
' ---------------------------------------------------------------------------
  Do
      lngTemp = 0    ' reset holding area
      
      ' If within range then save the data
      ' and jump out of the loop
      If dblTemp <= lngMaxValue Then
          lngTemp = CLng(dblTemp)
          Exit Do
      Else
          ' if the range exceeds the limits, convert to binary
          ' then back to a long integer
          strTemp = Right$(ConvertNumberToBinary(dblTemp), 15)
          If StrComp(strTemp, "-E-", vbTextCompare) = 0 Then
              ' if an error, then default to one.
              ' Safety net, should never happen.
              lngTemp = 1
          Else
              ' convert binary to long integer
              lngTemp = ConvertBinaryToLong(strTemp)
          End If
      End If
      
      ' See if the new seed value exceeds the range
      If lngTemp > lngMaxValue Then
          ' if it does, then reduce its size
          ' and try again
          dblTemp = dblTemp - 10000
      End If
              
  Loop Until lngTemp > 0 And lngTemp <= lngMaxValue
      
' ---------------------------------------------------------------------------
' Return the new seed value
' ---------------------------------------------------------------------------
  CalcSeedValue = lngTemp

End Function

Private Sub CalcLags()

' ***************************************************************************
' Routine:       CalcLags
'
' Description:   Routine will detemine sset of lag values to use.
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 20-May-2001  Kenneth Ives  kenaso@home.com
'              Routine created
' ***************************************************************************

' ---------------------------------------------------------------------------
' Define local variables
' ---------------------------------------------------------------------------
  Dim lngChoice  As Long
  Dim strTemp    As String
        
' ---------------------------------------------------------------------------
' Initialize variables
' ---------------------------------------------------------------------------
  m_lngLag1 = 0
  m_lngLag2 = 0

' ---------------------------------------------------------------------------
' randomly determine which lag values to use by capturing the last digit
' from the system current millisecond count
' ---------------------------------------------------------------------------
  strTemp = CStr(GetTickCount())
  lngChoice = Val(Right$(strTemp, 1))
  
' ---------------------------------------------------------------------------
' Evaluate the created value.   Do not modify any of these Lag values because
' you WILL NOT get your desired randomness.  It may look like random data,
' but will it pass all the tests?
' ---------------------------------------------------------------------------
  Select Case lngChoice
         Case 0
              m_lngLag1 = 17
              m_lngLag2 = 5
              
         Case 1
              m_lngLag1 = 33
              m_lngLag2 = 13
         
         Case 2
              m_lngLag1 = 39
              m_lngLag2 = 14
         
         Case 3
              m_lngLag1 = 52
              m_lngLag2 = 24
         
         Case 4
              m_lngLag1 = 63
              m_lngLag2 = 31
         
         Case 5
              m_lngLag1 = 73
              m_lngLag2 = 25
         
         Case 6
              m_lngLag1 = 97
              m_lngLag2 = 33
         
         Case 7
              m_lngLag1 = 607
              m_lngLag2 = 273
         
         Case Else
              m_lngLag1 = 97
              m_lngLag2 = 33
  End Select

End Sub

Private Function PreloadRnd() As Boolean

' ***************************************************************************
' Routine:       PreloadRnd
'
' Description:   This is the routine that starts the random number generating
'                process.  This routine will build 20,000 values as a primer.
'
' Returns:       TRUE/FALSE based on completion
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 13-May-2001  Kenneth Ives  kenaso@home.com
'              Routine converted from C to VB. Documented and modified.
' ***************************************************************************

' ---------------------------------------------------------------------------
' Define local variables
' ---------------------------------------------------------------------------
  Dim lngSeed1  As Long
  Dim lngSeed2  As Long

' ---------------------------------------------------------------------------
' Initialize variables
' ---------------------------------------------------------------------------
  lngSeed1 = 0
  lngSeed2 = 0

' ---------------------------------------------------------------------------
' Create two new seed values
' ---------------------------------------------------------------------------
  If m_blnTestMode Then
      ' use predetermined seed values
      lngSeed1 = TEST_SEED1
      lngSeed2 = TEST_SEED2
  Else
      ' not in test mode.  Calculate
      ' the two seed values
      While lngSeed1 = lngSeed2
          lngSeed1 = CalcSeedValue(31328)  ' 1 to 31328
          lngSeed2 = CalcSeedValue(30081)  ' 1 to 30081
      Wend
  End If
  
' ---------------------------------------------------------------------------
' determine what set of lag values to use
' ---------------------------------------------------------------------------
  If m_blnTestMode Then
      ' if in test mode, use predetermined values
      m_lngLag1 = 97
      m_lngLag2 = 33
  Else
      ' not in Test mode
      CalcLags
  End If
  
' ---------------------------------------------------------------------------
' Initialize the module variables in order to perform the calculations
' ---------------------------------------------------------------------------
  If Not InitializeRnd(lngSeed1, lngSeed2) Then
      PreloadRnd = False
      Exit Function
  End If

' ---------------------------------------------------------------------------
' Pre-Generate 20,000 random numbers.
' ---------------------------------------------------------------------------
  Dim tR() As Double
  FillArray tR(), 20000
  
' ---------------------------------------------------------------------------
' Finished Okay
' ---------------------------------------------------------------------------
  PreloadRnd = True
  
End Function

Private Function FillArray(adblData() As Double, Optional ByVal lngMaxReturn As Long = 1)
  
' ***************************************************************************
' Routine:       FillArray
'
' Description:   This is the routine that creates the random numbers loads
'                them into an array.
'
' Parameters:    adblRndNbrs() - array to hold the random generated values.
'                lngMaxReturn - [Optional] number of values to be returned.
'                               Defualt value of 1.
'
' Returns:       An array of random generated numbers.  TRUE/FALSE based
'                on completion.
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 13-May-2001  Kenneth Ives  kenaso@home.com
'              Routine converted from C to VB. Documented and modified.
' ***************************************************************************

' ---------------------------------------------------------------------------
' Define local variables
' ---------------------------------------------------------------------------
  Dim lngIndex       As Long
  Dim dblNewNbr      As Double
  Dim adblRndNbrs()  As Double
  
' ---------------------------------------------------------------------------
' Initialize variables
' ---------------------------------------------------------------------------
  m_strErrorMsg = ""
  
' ---------------------------------------------------------------------------
' Make sure request is for at least 1 number
' ---------------------------------------------------------------------------
  If lngMaxReturn < 1 Then
      m_strErrorMsg = "PrngKci.FillArray" & vbCrLf
      m_strErrorMsg = m_strErrorMsg & "Number of elements requested must exceed zero."
      Exit Function
  End If
  
' ---------------------------------------------------------------------------
' Initialize output array
' ---------------------------------------------------------------------------
  ReDim adblRndNbrs(1 To lngMaxReturn)
  ReDim adblData(1 To lngMaxReturn)
  m_dblMedianValue = 0
  
' ---------------------------------------------------------------------------
' If cRandomize = TRUE then reinitialize the random number generator
' ---------------------------------------------------------------------------
  If m_blnRandomize Then
      m_blnRandomize = False    ' reset flag
      ' reseed the random number array
      If Not PreloadRnd Then
          ' if there was an error then return a zero
          ReDim adblData(1 To 1)
          adblData(1) = 0
          Exit Function
      End If
  End If
  
' ---------------------------------------------------------------------------
' FillArray the new numbers
' ---------------------------------------------------------------------------
  For lngIndex = 1 To lngMaxReturn
      
      ' Create new number
      dblNewNbr = m_adblHold(m_lngIndex1) - m_adblHold(m_lngIndex2)
      
      If dblNewNbr < 0# Then                ' if negative then make positive
          dblNewNbr = dblNewNbr + 1#
      End If
      
      m_adblHold(m_lngIndex1) = dblNewNbr   ' Add to the holding array
      m_lngIndex1 = m_lngIndex1 - 1         ' Decrement first index by 1
      
      If m_lngIndex1 = 0 Then               ' if equal zero, reset to m_lngLag1
          m_lngIndex1 = m_lngLag1
      End If
      
      m_lngIndex2 = m_lngIndex2 - 1         ' Decrement second index by 1
      
      If m_lngIndex2 = 0 Then               ' if equal zero then reset to m_lngLag2
          m_lngIndex2 = m_lngLag2
      End If
      
      m_C1 = m_C1 - m_C2                    ' Decrement m_C1 by value of m_C2
      
      If m_C1 < 0# Then                     ' if m_C1 less than zero, add m_C3
          m_C1 = m_C1 + m_C3
      End If
      
      dblNewNbr = dblNewNbr - m_C1          ' decrement the new number
      
      If dblNewNbr < 0# Then                ' if negative then make positive
          dblNewNbr = dblNewNbr + 1#
      End If
    
      dblNewNbr = CDec(dblNewNbr)
      ' insert new number into output array
      CopyMemory adblRndNbrs(lngIndex), dblNewNbr, 8
  Next
  
' ---------------------------------------------------------------------------
' We finished successfully.  (1 double = 8 bytes)
' ---------------------------------------------------------------------------
  CopyMemory adblData(1), adblRndNbrs(1), (lngMaxReturn * 8)
  
End Function

Private Function InitializeRnd(lngSeed1 As Long, _
                               lngSeed2 As Long) As Boolean

' ***************************************************************************
' Routine:       InitializeRnd
'
' Description:   This is the initialization routine for preparing the
'                variables and arrays for random number generation.
'
' Parameters:    lngSeed1 - randomly generated seed value
'                lngSeed2 - randomly generated seed value
'
' Returns:       TRUE/FALSE based on completion
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 13-May-2001  Kenneth Ives  kenaso@home.com
'              Routine converted from C to VB. Documented and modified.
' ***************************************************************************

' ---------------------------------------------------------------------------
' Define local variables
' ---------------------------------------------------------------------------
  Dim lng_I           As Long
  Dim lng_J           As Long
  Dim lng_K           As Long
  Dim lng_L           As Long
  Dim lng_M           As Long
  Dim lngIndex1       As Long
  Dim lngIndex2       As Long
  Dim lngMax          As Long
  Dim dblGenNbr       As Double
  Dim dblIncrementor  As Double
  Dim adblTemp()      As Double
  
' ---------------------------------------------------------------------------
' Initialize variables
' ---------------------------------------------------------------------------
  m_strErrorMsg = ""
  
' ---------------------------------------------------------------------------
' If the input seeds do not meet the criteria then leave
'         lngSeed1 = 0 to 31328
'         lngSeed2 = 0 to 30081
' ---------------------------------------------------------------------------
  If (lngSeed1 < 0 Or lngSeed1 > 31328) Or _
     (lngSeed2 < 0 Or lngSeed2 > 30081) Then
           m_strErrorMsg = "PrngKci.InitializeRnd" & vbCrLf
           m_strErrorMsg = m_strErrorMsg & "Seed values invalid." & vbCrLf
           m_strErrorMsg = m_strErrorMsg & "Seed 1 = " & CStr(lngSeed1) & vbCrLf
           m_strErrorMsg = m_strErrorMsg & "Seed 2 = " & CStr(lngSeed2) & vbCrLf & vbCrLf
           m_strErrorMsg = m_strErrorMsg & "Seed 1 must be 0 to 31328" & vbCrLf
           m_strErrorMsg = m_strErrorMsg & "Seed 2 must be 0 to 30081" & vbCrLf
           InitializeRnd = False
           Exit Function
  End If
  
' ---------------------------------------------------------------------------
' Empty the temporary holding array and resize according to m_lngLag1 plus 1.
' The plus 1 is my buffer.
' ---------------------------------------------------------------------------
  lngMax = m_lngLag1 + 1
  Erase m_adblHold()
  ReDim m_adblHold(1 To lngMax)
  ReDim adblTemp(1 To lngMax)
  
' ---------------------------------------------------------------------------
' Initialize the starting long values.
' ---------------------------------------------------------------------------
  lng_I = (lngSeed1 / 185) Mod 185 + 2
  lng_J = lngSeed1 Mod 185 + 2
  lng_K = (lngSeed2 / 177) Mod 186 + 1
  lng_L = lngSeed2 Mod 177
  
' ---------------------------------------------------------------------------
' make m_lngLag1 iterations
' ---------------------------------------------------------------------------
  For lngIndex1 = 1 To m_lngLag1
  
      ' re-initialize
      dblGenNbr = 0#
      dblIncrementor = 0.5
    
      ' make 24 iterations
      For lngIndex2 = 1 To 24
          
          ' Calculate new base value
          lng_M = (((lng_I * lng_J) Mod 179) * lng_K) Mod 179
          
          ' swap the values
          lng_I = lng_J
          lng_J = lng_K
          lng_K = lng_M
          
          lng_L = (53 * lng_L + 1) Mod 169
      
          ' if the two calculated values divided by 64 have
          ' a remainder greater than 31 then increment the
          ' generated number by the incrementor value.
          If (lng_L * lng_M) Mod 64 >= 32 Then
              dblGenNbr = dblGenNbr + dblIncrementor
          End If
      
          ' adjust the incrementor
          dblIncrementor = dblIncrementor * 0.5
      Next
    
     dblGenNbr = CDec(dblGenNbr)
     CopyMemory adblTemp(lngIndex1), dblGenNbr, 8
  Next
  
' ---------------------------------------------------------------------------
' Copy data from memory.  (1 double = 8 bytes)
' ---------------------------------------------------------------------------
  CopyMemory m_adblHold(1), adblTemp(1), (lngMax * 8)
  
' ---------------------------------------------------------------------------
' Use "/" to divide and return a floating point result.
' ---------------------------------------------------------------------------
  m_C1 = 362436# / 16777216#
  m_C2 = 7654321# / 16777216#
  m_C3 = 16777213# / 16777216#
  
' ---------------------------------------------------------------------------
' Initialize module level variables
' ---------------------------------------------------------------------------
  m_lngIndex1 = m_lngLag1
  m_lngIndex2 = m_lngLag2

' ---------------------------------------------------------------------------
' Set the flags depicting we were successful.
' ---------------------------------------------------------------------------
  InitializeRnd = True
  
End Function

Public Function BuildRndString(lngDataLength As Long, _
                               intLower As Integer, _
                               intUpper As Integer, _
                      Optional bUse_Hex_Chars As Boolean = False, _
                      Optional bConvertAllData As Boolean = False) As String

' ***************************************************************************
' Routine:       BuildRndString
'
' Description:   Build a random data string.  Keep it within reason so as
'                not to clog your memory processes.  The maximum value I
'                have ever used was 32768 (1024 * 32).
'
'
' Parameters:    lngDataLength = Length of the data string
'                intLower - Lower value limit (ASCII 0-255) in a range
'                intUpper - Upper value limit (ASCII 0-255) in a range
'                           ex: lower = 33  upper=126  (Keyboard chars)
'
'                bUse_Hex_Chars - (Optional) (Default) FALSE - do not
'                      convert to hex
'                                 TRUE  - convert each byte to its 2-char hex
'                      representation before returning (Ex:  15  becomes  0F)
'
'                bConvertAllData - (Optional) Used only if the data is to be
'                      converted to HEX prior to its return.  (Default) FALSE-
'                      Return the data in the same length as the variable
'                      lngDataLength after the conversion.
'                                 TRUE - Return all the data after the
'                      conversion.  Data length will double due to the change.
'
'
' Returns:       String of random generated data
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 24-DEC-1999  Kenneth Ives  kenaso@home.com
'              Created routine
' 12-NOV-2000  Kenneth Ives  kenaso@home.com
'              Modified to use byte arrays and changed the logic flow.  Added
'              a new switch for the amount to data to be returned if it is
'              converted to Hex format.
' 01-APR-2001  Plex          PlexusR@netscape.net
'              Redimmed the byte array (bad) instead of the string array prior
'              to returning the values when creating ASCII printable characters.
' 12-MAY-2001  Plex  PlexusR@netscape.net
'              Added intLower and intUpper to represent the lower and upper
'              Ascii decimal value limits, inclusive.
' ***************************************************************************

' ---------------------------------------------------------------------------
' Define local variables
' ---------------------------------------------------------------------------
  Dim blnBadData  As Boolean
  Dim lngIndex    As Long
  Dim lngMax      As Long
  Dim strTemp     As String
  Dim astrData()  As String
  Dim abytData()  As Byte
  
' ---------------------------------------------------------------------------
' Define local variables
' ---------------------------------------------------------------------------
  strTemp = ""
  m_strErrorMsg = ""
  blnBadData = False
  
' ---------------------------------------------------------------------------
' Test desired length
' ---------------------------------------------------------------------------
  If lngDataLength < 1 Or lngDataLength > 32768 Then
      m_strErrorMsg = "PrngKci.BuildRndString" & vbCrLf
      m_strErrorMsg = m_strErrorMsg & "Length of the return data is invalid" & vbCrLf
      m_strErrorMsg = m_strErrorMsg & "Length limits are 1-32768."
      BuildRndString = ""
      Exit Function
  End If

' ---------------------------------------------------------------------------
' See if the limits are within range.  If not, then leave.
' ---------------------------------------------------------------------------
  If intLower > intUpper Then
      m_strErrorMsg = "PrngKci.BuildRndString" & vbCrLf
      m_strErrorMsg = m_strErrorMsg & "Invalid ASCII range limits." & vbCrLf
      m_strErrorMsg = m_strErrorMsg & "Low limit (0-255) cannot exceed the Upper limit (0-255)."
      BuildRndString = ""
      Exit Function
  End If

  If (intLower < 0 Or intLower > 255) Or _
     (intUpper < 0 Or intUpper > 255) Then
          m_strErrorMsg = "PrngKci.BuildRndString" & vbCrLf
          m_strErrorMsg = m_strErrorMsg & "Invalid ASCII range limits." & vbCrLf
          m_strErrorMsg = m_strErrorMsg & "Low limit (0-255)   Upper limit (0-255)."
          BuildRndString = ""
          Exit Function
  End If

' ---------------------------------------------------------------------------
' Build a string of randomly generated characters
' ---------------------------------------------------------------------------
  abytData = CreateCharData(intLower, intUpper, blnBadData, lngDataLength)
  If blnBadData Then
      BuildRndString = ""
      Exit Function
  End If
  
' ---------------------------------------------------------------------------
' Build a string of randomly generated characters
' ---------------------------------------------------------------------------
  abytData = RemixBytes(abytData(), blnBadData)
  If blnBadData Then
      BuildRndString = ""
      Exit Function
  End If
  
' ---------------------------------------------------------------------------
' See if the data is to be converted to hex format
' ---------------------------------------------------------------------------
  If bUse_Hex_Chars Then
      
      lngMax = lngDataLength * 2  ' double the size due to conversion
      ReDim astrData(lngMax)      ' size the array before loading
      
      For lngIndex = 0 To lngDataLength - 1
          ' convert byte data to 2-char hex
          astrData(lngIndex) = Right$("00" & Hex(abytData(lngIndex)), 2)
      Next
  Else
      ' Redim the astrData() array to hold the returning values.
      ' Thanks to Plex for finding this oversight. (KCI 1-Apr-2001)
      ReDim astrData(lngDataLength)
      
      ' convert byte data to string data
      For lngIndex = 0 To lngDataLength - 1
          astrData(lngIndex) = Chr$(abytData(lngIndex))
      Next
  End If
  
' ---------------------------------------------------------------------------
' Build an output string before leaving
' ---------------------------------------------------------------------------
  strTemp = BuildOutputString(astrData())
  
' ---------------------------------------------------------------------------
' See if we have to truncate the data string before leaving
' ---------------------------------------------------------------------------
  If bUse_Hex_Chars Then
      If bConvertAllData Then
          BuildRndString = strTemp
      Else
          ' truncate the data string before leaving
          BuildRndString = Left$(strTemp, lngDataLength)
      End If
  Else
      BuildRndString = strTemp
  End If
  
End Function

Private Function BuildOutputString(astrData() As String) As String

' ***************************************************************************
' Routine:       BuildOutputString
'
' Description:   This routine will build the final output string
'
' Parameters:    astrData() - input string array
'
' Returnss:      String of data
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 12-NOV-2000  Kenneth Ives  kenaso@home.com
'              Created routine
' 15-MAY-2001  VB2-The-Max   http://www.vb2themax.com/
'              From an article titled "10 Hot Tips from VB-2-the-Max"
'              by Francesco Balena
'              This is tip no.9 on faster string concatenation with a little
'              modification.
' ***************************************************************************

' ---------------------------------------------------------------------------
' Define variables
' ---------------------------------------------------------------------------
  Dim lngLoop        As Long
  Dim lngIndex       As Long
  Dim lngMax         As Long
  Dim lngLength      As Long
  Dim lngPaddingLen  As Long
  Dim strTemp        As String
  Dim strOutput      As String
  
  Const ADD_SPACES   As Long = 10000
  
' ---------------------------------------------------------------------------
' Initialize variables
' ---------------------------------------------------------------------------
  strTemp = ""
  m_strErrorMsg = ""
  lngIndex = 1                        ' Output string pointer
  lngMax = UBound(astrData)           ' max number of elements in array
  lngPaddingLen = (ADD_SPACES * 9)    ' 90000 blank spaces
  strOutput = Space$(lngPaddingLen)   ' preload output string
  
' ---------------------------------------------------------------------------
' If no data then leave
' ---------------------------------------------------------------------------
  If lngMax < 0 Then
      m_strErrorMsg = "PrngKci.BuildOutputString" & vbCrLf
      m_strErrorMsg = m_strErrorMsg & "No data to build output string."
      BuildOutputString = ""
      Exit Function
  End If
  
' ---------------------------------------------------------------------------
' build the output string
' ---------------------------------------------------------------------------
  For lngLoop = 0 To lngMax - 1
      
      ' Test the length on the data to be added to the output string
      lngLength = Len(astrData(lngLoop))
      
      ' see if some more padding has to be added to the output string
      If (lngIndex + lngLength) >= lngPaddingLen Then
          lngPaddingLen = lngPaddingLen + ADD_SPACES   ' boost blank space counter
          strOutput = strOutput & Space$(ADD_SPACES)   ' append some blank spaces
      End If
      
      ' insert data into output string
      Mid$(strOutput, lngIndex, lngLength) = astrData(lngLoop)
      lngIndex = lngIndex + lngLength                 ' increment output string pointer
  Next
  
  strOutput = RTrim$(strOutput)     ' remove trailing blanks
  BuildOutputString = strOutput     ' Return the completed string.

End Function

Private Function CreateCharData(ByVal intLower As Integer, _
                                ByVal intUpper As Integer, _
                                blnBadData As Boolean, _
                                Optional lngDataLength As Long = 50) As Variant

' ***************************************************************************
' Routine:       CreateCharData
'
' Description:   This routine will build a string of random generated data
'                based on type of data desired and the length of the string.
'
' Parameters:    intLower/intUpper - return Ascii values between these two
'                       numbers, inclusive
'                blnBadData - Switch is set to TRUE only if there are problems
'                       during the generation of the numbers.
'                lngDataLength - amount of data to create
'
' Returns:       Byte array of ASCII values
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 24-DEC-1999  Kenneth Ives  kenaso@home.com
'              Created routine
' 12-NOV-2000  Kenneth Ives  kenaso@home.com
'              Modified to use byte arrays and changed the logic flow.
' 12-MAY-2001  Plex  PlexusR@netscape.net
'              1.  Added Lower/Upper parameters and removed the boolean switch
'              2.  Updated flower box with corrected documentation
'              3.  Added two new variables to accomodate the random value
'                  limits
' 17-MAY-2001  Kenneth Ives kenaso@home.com
'              Added a boolean flag to identify when random values were not
'              generated.
' ***************************************************************************

' ---------------------------------------------------------------------------
' Define local variables
' ---------------------------------------------------------------------------
  Dim lngIndex    As Long
  Dim lngCount    As Long
  Dim intChar     As Integer
  Dim abytData()  As Byte
  Dim adblData()  As Double

' ---------------------------------------------------------------------------
' Initialize local variables
' ---------------------------------------------------------------------------
  Erase abytData()
  lngCount = 0
  ReDim abytData(lngDataLength)
              
' ---------------------------------------------------------------------------
' Create an array of random numbers
' ---------------------------------------------------------------------------
  If PreloadRnd Then
      FillArray adblData(), lngDataLength + 1 ' Plus 1 = buffer zone
  Else
      blnBadData = True
      CreateCharData = abytData()
      Exit Function
  End If
  
' ---------------------------------------------------------------------------
' Create the data
' ---------------------------------------------------------------------------
  For lngIndex = 0 To lngDataLength - 1
      intChar = CInt(Int(((intUpper + 1) - intLower) * adblData(lngIndex + 1)) + intLower)
      abytData(lngIndex) = intChar
  Next
  
' ---------------------------------------------------------------------------
' Return the array of data
' ---------------------------------------------------------------------------
  CreateCharData = abytData()
  
End Function

Public Function NonRepeatingNbrs(alngMixed() As Long, ByVal lngReturnQty As Long, _
                                 ByVal lngMinValue As Long, _
                                 ByVal lngMaxValue As Long, _
                              Optional bSortData As Boolean = True)

' ***************************************************************************
' Routine:       NonRepeatingNbrs
'
' Description:   With this routine you can generate a series of
'                non-repeating numbers.  An array will be loaded starting
'                with the base number (lngMinValue) requested up to the
'                maximum value requested (lngMaxValue).  This array is
'                then passed to another routine (MixTheNumbers) to be
'                throughly rearranged.  When it is returned, the requested
'                number of elements (lngReturnQty) from the mixed array
'                are transferred sequentially to the return array
'                (alngMixed()).
'
' Parameters:    lngReturnQty - Total count of numbers to be returned
'                lngMinValue - This is the lowest possible value that could
'                            be returned.  If missing, defaults to zero.
'                lngMaxValue - Choose the return values from values up
'                            to this maximum value
'                bSortData - TRUE/FALSE to sort the returned data.  Default
'                            is TRUE.
'
' Syntax:        x = NonRepeatingNbrs(10, 0, 9)
'                Return 10 numbers, lowest = 0, highest = 9
'
' Returns:       An array of random renerated numbers.
'
' NOTE:          To see any error messages, you must request property
'                ErrMessage
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 20-MAY-2001  Kenneth Ives kenaso@home.com
'              Created routine
' 22-MAY-2001  Kenneth Ives kenaso@home.com
'              Corrected logic in creation of output array.
' 01-JUN-2001  Kenneth Ives kenaso@home.com
'              Corrected logic in creation of input array.
' ***************************************************************************
    
' ---------------------------------------------------------------------------
' define local variables
' ---------------------------------------------------------------------------
  Dim lngIndex         As Long
  Dim lngTmpNbr        As Long
  Dim lngMaxSize       As Long
  Dim alngSelection()  As Long
  Dim blnBadData       As Boolean
  
' ---------------------------------------------------------------------------
' Initialize variables.
' ---------------------------------------------------------------------------
  Erase alngMixed()       ' Verify arrays are empty
  Erase alngSelection()
  lngTmpNbr = 0
  m_strErrorMsg = ""
  blnBadData = False

' ---------------------------------------------------------------------------
' Test data input.  Leave if nothing to work with.
' ---------------------------------------------------------------------------
  If lngReturnQty <= 0 Or lngMaxValue <= 0 Then
      m_strErrorMsg = "PrngKci.NonRepeatingNbrs" & vbCrLf
      m_strErrorMsg = m_strErrorMsg & "Value requested must be a positive "
      m_strErrorMsg = m_strErrorMsg & "number above zero."
      ReDim alngMixed(1)
      alngMixed(0) = -1
      Exit Function
  End If

' ---------------------------------------------------------------------------
' Test the lowest possible number
' ---------------------------------------------------------------------------
  If lngMinValue >= lngMaxValue Then
      m_strErrorMsg = "PrngKci.NonRepeatingNbrs" & vbCrLf
      m_strErrorMsg = m_strErrorMsg & "The lowest possible number is greater "
      m_strErrorMsg = m_strErrorMsg & "than or equal to the maximum value."
      ReDim alngMixed(1)
      alngMixed(0) = -1
      Exit Function
  End If
  
' ---------------------------------------------------------------------------
' Test data input.  Leave if quantity requested is greater than the amount
' to work with.
' ---------------------------------------------------------------------------
  If lngReturnQty > (lngMaxValue + 1) Then
      m_strErrorMsg = "PrngKci.NonRepeatingNbrs" & vbCrLf
      m_strErrorMsg = m_strErrorMsg & "Quantity requested is greater than the "
      m_strErrorMsg = m_strErrorMsg & "value to choose from."
      ReDim alngMixed(1)
      alngMixed(0) = -1
      Exit Function
  End If

' ---------------------------------------------------------------------------
' Test data input.  Leave if quantity requested is greater than the amount
' to work with.
' ---------------------------------------------------------------------------
  If lngReturnQty > (lngMaxValue - lngMinValue + 1) Then
      m_strErrorMsg = "PrngKci.NonRepeatingNbrs" & vbCrLf
      m_strErrorMsg = m_strErrorMsg & "Quantity requested is greater than the "
      m_strErrorMsg = m_strErrorMsg & "value to choose from."
      ReDim alngMixed(1)
      alngMixed(0) = -1
      Exit Function
  End If
  
' ---------------------------------------------------------------------------
' Resize arrays based on values passed to this routine.
' ---------------------------------------------------------------------------
  lngMaxSize = (lngMaxValue - lngMinValue) + 3  ' 3 Buffer elements
  ReDim alngSelection(lngMaxSize)               ' number of elements to pick from
  ReDim alngMixed(1 To lngReturnQty)            ' number of elements to be returned
  
' ---------------------------------------------------------------------------
' Preload the selection array.  These are the values to choose from.
' ---------------------------------------------------------------------------
  For lngIndex = 0 To (lngMaxSize - 1)
      alngSelection(lngIndex) = lngMinValue + lngIndex
  Next
  
' ---------------------------------------------------------------------------
' Mix the numbers
' ---------------------------------------------------------------------------
  alngSelection = MixTheNumbers(alngSelection(), blnBadData)
  If blnBadData Then
     ReDim alngMixed(1)
     alngMixed(1) = -1
     Exit Function
  End If
  
' ---------------------------------------------------------------------------
' Transfer the data from the array that was just created back to the original
' array.
' ---------------------------------------------------------------------------
  For lngIndex = 0 To lngMaxSize - 1
      If alngSelection(lngIndex) >= lngMinValue Then
          If alngSelection(lngIndex) <= lngMaxValue Then
              ' if the value that was mixed is greater than or equal
              ' to the base value requested then add it to the output
              ' array.  Should always be TRUE.
              lngTmpNbr = lngTmpNbr + 1
              alngMixed(lngTmpNbr) = alngSelection(lngIndex)
          
              ' see if we have the return quantity requested
              If lngTmpNbr = lngReturnQty Then
                  Exit For
              End If
          End If
      End If
  Next
  
' ---------------------------------------------------------------------------
' If more than one number is requested and the user wants to sort the data,
' we will use a Sort algorithym to sort in ascending order.
' ---------------------------------------------------------------------------
  If bSortData Then
  
      Select Case lngReturnQty
      
             ' do a bubble sort
             Case 2 To 25
                  For lngIndex = 1 To (lngReturnQty - 1)
                      If alngMixed(lngIndex) > alngMixed(lngIndex + 1) Then
                          SwapData alngMixed(lngIndex), alngMixed(lngIndex + 1)
                          lngIndex = 0   ' Reset the index and start over
                      End If
                  Next
           
             ' do a QSortLong
             Case Is > 25
                  QSortLong alngMixed(), LBound(alngMixed), UBound(alngMixed)
                
             Case Else
                  ' fall thru and do nothing to the array
      End Select
  End If
  
' ---------------------------------------------------------------------------
' Successful finish
' ---------------------------------------------------------------------------
  
End Function

Private Function RemixBytes(abytData() As Byte, _
                            blnBadData As Boolean) As Variant

' ***************************************************************************
' Routine:       RemixBytes
'
' Description:   An array of numbers is passed to this routine.  Now loop
'                thru this array and randomly pick an index.  Move the
'                element from the random position to the current position.
'                Do the same process two to ten times.
'
' Parameters:    abytData() - return the mixed values in this array.
'                blnBadData - return flag to indicate if the operation was
'                        successful
'
' Returns:       An array of numbers that have truely been reshuffled.
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 24-DEC-1999  Kenneth Ives  kenaso@home.com
'              Created routine. Got this idea from VB2TheMax web site at
'              http://www.vb2themax.com/
' ***************************************************************************
    
' ---------------------------------------------------------------------------
' define local variables
' ---------------------------------------------------------------------------
  Dim lngIndex1      As Long
  Dim lngIndex2      As Long
  Dim lngNewIndex    As Long
  Dim lngMin         As Long
  Dim lngMax         As Long
  Dim lngMaxSize     As Long
  Dim lngCount       As Long
  Dim lngOuterLoop   As Long
  Dim strTemp        As String
  Dim adblData()     As Double
  
' ---------------------------------------------------------------------------
' Determine the low and high end of the array
' ---------------------------------------------------------------------------
  ' Get minimum number of elements in array
  lngMin = LBound(abytData)
  
  ' Get maximum number of elements in array
  If lngMin = 0 Then
      lngMax = UBound(abytData) - 1
  Else
      lngMax = UBound(abytData)
  End If
  
  lngMaxSize = lngMax * 2
  lngOuterLoop = 0
  
' ---------------------------------------------------------------------------
' See if anything was passed that has to be mixed
' ---------------------------------------------------------------------------
  If lngMin >= lngMax Then
      blnBadData = True
      RemixBytes = abytData()
      Exit Function
  End If

' ---------------------------------------------------------------------------
' Determine outer loop size
' ---------------------------------------------------------------------------
  While lngOuterLoop < 3 Or lngOuterLoop > 7
      DoEvents
      strTemp = CStr(GetTickCount())
      lngOuterLoop = Val(Right$(strTemp, 1))
  Wend
    
' ---------------------------------------------------------------------------
' The array will now undergo multiple mixing operations
' ---------------------------------------------------------------------------
  For lngIndex1 = 1 To lngOuterLoop
        
      FillArray adblData(), lngMaxSize   ' fill arraay with random numbers
      lngCount = 0                      ' reset the counter
      
      ' go thru the input array and randomly
      ' pick an element to move to
      ' another position within the array.
      For lngIndex2 = lngMin To lngMax
                
          ' generate a new index that does mot
          ' match the current index
          Do
              ' see if we need some more numbers
              If (lngCount + 1) >= lngMaxSize Then
                  FillArray adblData(), lngMaxSize    ' fill arraay with random numbers
                  lngCount = 0                      ' reset the counter
              End If
              
              lngCount = lngCount + 1
              lngNewIndex = CLng((((lngMax + 1) - lngMin) * adblData(lngCount)) + lngMin)
          Loop Until lngNewIndex <> lngIndex2
          
          ' swap the the data
          SwapData abytData(lngIndex2), abytData(lngNewIndex)
      Next
  Next

' ---------------------------------------------------------------------------
' Return the rearranged data
' ---------------------------------------------------------------------------
  blnBadData = False
  RemixBytes = abytData()
  
End Function

Private Function MixTheNumbers(alngSelection() As Long, _
                               blnBadData As Boolean) As Variant

' ***************************************************************************
' Routine:       MixTheNumbers
'
' Description:   An array of numbers is passed to this routine.  Now loop
'                thru this array and randomly pick an index.  Move the
'                element from the random position to the current position.
'                Do the same process two to ten times.
'
' Parameters:    alngSelection() - return the mixed values in this array.
'                blnBadData - return flag to indicate if the operation was
'                        successful
'
' Returns:       An array of numbers that have truely been reshuffled.
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 24-DEC-1999  Kenneth Ives  kenaso@home.com
'              Created routine. Got this idea from VB2TheMax web site. Please
'              visit http://www.vb2themax.com/
' ***************************************************************************
    
' ---------------------------------------------------------------------------
' define local variables
' ---------------------------------------------------------------------------
  Dim lngIndex1      As Long
  Dim lngIndex2      As Long
  Dim lngNewIndex    As Long
  Dim lngMin         As Long
  Dim lngMax         As Long
  Dim lngMaxSize     As Long
  Dim lngCount       As Long
  Dim lngOuterLoop   As Long
  Dim strTemp        As String
  Dim adblData()     As Double
  
' ---------------------------------------------------------------------------
' Determine the low and high end of the array
' ---------------------------------------------------------------------------
  ' Get minimum number of elements in array
  lngMin = LBound(alngSelection)
  
  ' Get maximum number of elements in array
  If lngMin = 0 Then
      lngMax = UBound(alngSelection) - 1
  Else
      lngMax = UBound(alngSelection)
  End If

  lngMaxSize = lngMax * 2
  lngOuterLoop = 0
  
' ---------------------------------------------------------------------------
' See if anything was passed that has to be mixed
' ---------------------------------------------------------------------------
  If lngMin >= lngMax Then
      blnBadData = True
      MixTheNumbers = alngSelection()
      Exit Function
  End If

' ---------------------------------------------------------------------------
' Determine outer loop size
' ---------------------------------------------------------------------------
  While lngOuterLoop < 3 Or lngOuterLoop > 7
      DoEvents
      strTemp = CStr(GetTickCount())
      lngOuterLoop = Val(Right$(strTemp, 1))
  Wend
  
' ---------------------------------------------------------------------------
' The array will now undergo multiple mixing operations
' ---------------------------------------------------------------------------
  For lngIndex1 = 1 To lngOuterLoop
        
      FillArray adblData(), lngMaxSize   ' fill arraay with random numbers
      lngCount = 0                      ' reset the counter
      
      ' go thru the input array and randomly
      ' pick an element to move to
      ' another position within the array.
      For lngIndex2 = lngMin To lngMax
                
          ' generate a new index that does mot
          ' match the current index
          Do
              ' see if we need some more numbers
              If (lngCount + 1) >= lngMaxSize Then
                  FillArray adblData(), lngMaxSize    ' fill arraay with random numbers
                  lngCount = 0                      ' reset the counter
              End If
              
              lngCount = lngCount + 1
              lngNewIndex = CLng((((lngMax + 1) - lngMin) * adblData(lngCount)) + lngMin)
          Loop Until (lngNewIndex <> lngIndex2) And _
                     (lngNewIndex <= lngMax)
          
          ' swap the the data
          SwapData alngSelection(lngIndex2), alngSelection(lngNewIndex)
      Next
  Next

' ---------------------------------------------------------------------------
' Return the rearranged data
' ---------------------------------------------------------------------------
  blnBadData = False
  MixTheNumbers = alngSelection()
  
End Function

Private Sub QSortLong(alngData() As Long, _
                      lngLow As Long, _
                      lngHigh As Long)

' ***************************************************************************
' Routine:       QSortLong
'
' Description:   This routine will accept and sort in ascending order a
'                numeric array of data.  This routine is used when the data
'                to be sorted is known to be ALL numeric, non-decimal.
'
' Parameters:    alngData() - Array to be sorted
'                lngLow      - Minimum number of elements in the array
'                lngHigh     - Maximum numer of elements in the array
'
' ===========================================================================
'    DATE      NAME             DESCRIPTION
' -----------  ---------------  ---------------------------------------------
' 30-APR-2000  Kenneth Ives     Routine created kenaso@home.com
' ***************************************************************************

' ----------------------------------------------------------------------
' Define local variables
' ----------------------------------------------------------------------
  Dim lngMidPoint  As Long      ' midpoint of the array to be sorted
  Dim lngTmpLow    As Long      ' Index pointer
  Dim lngTmpHigh   As Long      ' Index pointer
   
' ----------------------------------------------------------------------
' See if this is an empty array by checking to see if there is data in
' the first element.  If not, then leave.
' ----------------------------------------------------------------------
  If Len(Trim(alngData(1))) = 0 Then
      Exit Sub
  End If
  
' ----------------------------------------------------------------------
' Leave if there is nothing to sort
' ----------------------------------------------------------------------
  If lngLow >= lngHigh Then
      Exit Sub
  End If

' ----------------------------------------------------------------------
' Save the count of the minimum and maximum number of elements in the
' array to be sorted.
' ----------------------------------------------------------------------
  lngTmpLow = lngLow
  lngTmpHigh = lngHigh
   
' ----------------------------------------------------------------------
' Calculate the midpoint of the array
' ----------------------------------------------------------------------
  lngMidPoint = alngData((lngLow + lngHigh) / 2)

' ----------------------------------------------------------------------
' Start the sorting process
' ----------------------------------------------------------------------
  While (lngTmpLow <= lngTmpHigh)
       
      ' Always process the low end first.  Loop as long the array data
      ' element is LESS than the data in the temporary holding area
      ' and the temporary low value is LESS than the maximum number of
      ' array elements.
      While (alngData(lngTmpLow) < lngMidPoint And lngTmpLow < lngHigh)
          lngTmpLow = lngTmpLow + 1     ' Increment the temp low counter
      Wend
   
      ' Now, we will process the high end.  Loop as long the data in the
      ' temporary holding area is LESS than the array data element
      ' and the temporary high value is GREATER than the minimum number
      ' of array elements.
      While (lngMidPoint < alngData(lngTmpHigh) And lngTmpHigh > lngLow)
          lngTmpHigh = lngTmpHigh - 1   ' Decrement the temp high counter
      Wend

      ' if the temp low end is LESS than or equal to the temp high end,
      ' then swap places
      If (lngTmpLow <= lngTmpHigh) Then
          SwapData alngData(lngTmpLow), alngData(lngTmpHigh)
          lngTmpLow = lngTmpLow + 1     ' Increment the temp low counter
          lngTmpHigh = lngTmpHigh - 1   ' Decrement the temp high counter
      End If
 Wend
    
' ----------------------------------------------------------------------
' If the minimum number of elements in the array is LESS than the temp
' high end, then make a recursive call to this routine.  Always sort
' the low end of the array first.  This gives you a solid base.
' ----------------------------------------------------------------------
  If (lngLow < lngTmpHigh) Then
      QSortLong alngData(), lngLow, lngTmpHigh
  End If
   
' ----------------------------------------------------------------------
' If the temp low end is LESS than the maximum number of elements in
' the array, then make a recursive call to this routine.  The high end
' is always sorted last.
' ----------------------------------------------------------------------
  If (lngTmpLow < lngHigh) Then
      QSortLong alngData(), lngTmpLow, lngHigh
  End If

End Sub

Private Function ConvertBinaryToLong(strBin As String) As Long
  
' ***************************************************************************
' Routine:       ConvertBinaryToLong
'
' Description:   Converts string of Binary bits to Long value
'
' Parameters:    strBin - data to be converted
'
' Return:        Long integer
'
' ===========================================================================
'    DATE      NAME             DESCRIPTION
' -----------  ---------------  ---------------------------------------------
' Unknown      Serge Baronovsky VBCity.com    http://www.vbcity.com/
' 05-JUN-2001  Kenneth Ives     kenaso@home.com
'                               modified and documented
' ***************************************************************************
    
' ---------------------------------------------------------------------------
' Define local variables
' ---------------------------------------------------------------------------
  Dim lngNumber  As Long
  Dim lngIndex  As Long
  
' ---------------------------------------------------------------------------
' Initialize variables
' ---------------------------------------------------------------------------
  lngNumber = 0
  
' ---------------------------------------------------------------------------
' Parse binary data and convert to numeric data
' ---------------------------------------------------------------------------
  For lngIndex = Len(strBin) To 1 Step -1
      If Mid$(strBin, lngIndex, 1) = "1" Then
          lngNumber = lngNumber + 2 ^ (Len(strBin) - lngIndex)
      End If
  Next 'lngIndex
  
' ---------------------------------------------------------------------------
' Return the converted data
' ---------------------------------------------------------------------------
  ConvertBinaryToLong = lngNumber
  
End Function

Private Function ConvertNumberToBinary(ByVal dblNumber As Double) As String
  
' ***************************************************************************
' Routine:       ConvertNumberToBinary
'
' Description:   Converts numbers to string of Binary bits
'
' Parameters:    dblNumber - value to be converted to binary string
'
' Return:        Data string of binary "1" and "0"
'
' ===========================================================================
'    DATE      NAME             DESCRIPTION
' -----------  ---------------  ---------------------------------------------
' Unknown      Serge Baronovsky VBCity.com    http://www.vbcity.com/
' 05-JUN-2001  Kenneth Ives     kenaso@home.com
'                               modified and documented
' ***************************************************************************
    
' ---------------------------------------------------------------------------
' Define local variables
' ---------------------------------------------------------------------------
  Dim strBin    As String
  Dim lngIndex  As Long

' ---------------------------------------------------------------------------
' Initialize variables
' ---------------------------------------------------------------------------
  strBin = ""
  
  On Error GoTo ConvertNumberToBinary_Error
' ---------------------------------------------------------------------------
' Start processing input value
' ---------------------------------------------------------------------------
  For lngIndex = 64 To 0 Step -1

      If Int(dblNumber / (2 ^ lngIndex)) = 1 Then
          strBin = strBin & "1"
          dblNumber = dblNumber - (2 ^ lngIndex)
      Else
          If Len(Trim$(strBin)) > 0 Then
              strBin = strBin & "0"
          End If
      End If
  Next 'lngIndex

' ---------------------------------------------------------------------------
' Return converted data
' ---------------------------------------------------------------------------
  ConvertNumberToBinary = strBin
  Exit Function
  

ConvertNumberToBinary_Error:
' ---------------------------------------------------------------------------
' If an error occurs then return some arbitrary data to test
' ---------------------------------------------------------------------------
  If Err <> 0 Or Len(Trim$(strBin)) = 0 Then
      ConvertNumberToBinary = "-E-"
  End If

End Function

Private Function SwapData(varValue1 As Variant, varValue2 As Variant)

' ***************************************************************************
' Routine:       SwapData
'
' Description:   Swap data with each other.  I wrote this function since
'                BASIC stopped having its own SWAP function.
'
' Parameters:    varValue1 - long integer to be swapped with varValue2
'                varValue2 - long integer to be swapped with varValue1
'
' Return Values: Swapped data
'
' ===========================================================================
'    DATE      NAME             DESCRIPTION
' -----------  ---------------  ---------------------------------------------
' 15-JUN-2000  Kenneth Ives     Routine created by kenaso@home.com
' ***************************************************************************
    
' ---------------------------------------------------------------------------
' Define local variables
' ---------------------------------------------------------------------------
  Dim varTmp As Variant
  
' ---------------------------------------------------------------------------
' Swap the values with each other
' ---------------------------------------------------------------------------
  varTmp = varValue1      ' copy value1 to temp hold area
  varValue1 = varValue2   ' copy value2 to value1
  varValue2 = varTmp      ' copy temp hold area to value2
  
End Function

Private Sub Class_Initialize()

' ---------------------------------------------------------------------------
' Empty the module level variables
' ---------------------------------------------------------------------------
  m_C1 = 0
  m_C2 = 0
  m_C3 = 0
  m_dblMedianValue = 0
  m_strErrorMsg = ""
  m_blnTestMode = False
  m_blnRandomize = False

' ---------------------------------------------------------------------------
' Initialize the random number generator when this class is initialized.
' ---------------------------------------------------------------------------
  If Not PreloadRnd Then
      MsgBox "Failed to initialize PrngKCI.dll" & vbCrLf & _
             "This file is corrupted.", vbCritical + vbOKOnly, "Bad PrngKCI.DLL"
      Class_Terminate
  End If
  
End Sub

Private Sub Class_Terminate()

' ---------------------------------------------------------------------------
' Empty the module level variables
' ---------------------------------------------------------------------------
  m_C1 = 0
  m_C2 = 0
  m_C3 = 0
  m_blnTestMode = False
  m_dblMedianValue = 0
  m_strErrorMsg = ""
  Erase m_adblHold()
  
End Sub

