Attribute VB_Name = "STAT_RANDOM_PRNG_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.


'************************************************************************************
'************************************************************************************
'FUNCTION      : RANDOM_SAMPLING_SIMULATION_FUNC

'DESCRIPTION   : Reliable Pseudo-Random Number Generator (PRNG)

'Excel creates random numbers using a pretty simple formula which
'only repeats about every 2 billion numbers. That's plenty, I hear
'you cry. The problem is that once you use more than 16,000 numbers
'at a time, the numbers thereafter may become very non-random especially
'if you are interested in numbers in a very small range, eg < 0.01.

'This is slower than the VBA Rnd function, but it produces numbers
'that should remain reliable no matter how many numbers you use. It
'repeats only every 2^144 and 2^121 numbers respectively.

'---------------------------------------------------------------------------------
'PRNGcki
'---------------------------------------------------------------------------------
'It comes from the well known George Marsiglia - or write the original
'Fortran code, or convert that code to C, or convert the C code to VB

'---------------------------------------------------------------------------------
'L 'Ecuyer
'---------------------------------------------------------------------------------
'L 'Ecuyer is another well known expert in generating random numbers. This
'is a very solid algorithm which passes all the Diehard tests.
'---------------------------------------------------------------------------------

'LIBRARY       : STATISTICS
'GROUP         : RANDOM_PRNG
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 14/02/2008
'************************************************************************************
'************************************************************************************

Function PRNG_TABLE_FUNC(Optional ByVal NROWS As Long = 1000, _
Optional ByVal nLOOPS As Long = 100000)

'----------------------------------------------------------------------------------
'This function calculates 100,000 numbers and collects them in 1000 "buckets",
'by size, ie the first bucket is for the range 0-0.001.

'They should be relatively uniform.

'This gives you some idea of whether the metohds are working (there are much
'better tests of course), and it shows you their relative speeds.
'----------------------------------------------------------------------------------

Dim i As Long
Dim j As Long

Dim X_VAL As Single
Dim T_VAL As Single

Dim MIN_VAL As Long
Dim MAX_VAL As Long

Dim TEMP_MATRIX() As Variant

Dim ECUYER_OBJ As New clsEcuyer
Dim PRNGKCI_OBJ As New clsPrngKci

On Error GoTo ERROR_LABEL

ReDim TEMP_MATRIX(0 To NROWS + 3, 0 To 3)

'------------------------------------------------------------------------------------
TEMP_MATRIX(0, 0) = "---": TEMP_MATRIX(0, 1) = "Rnd"
TEMP_MATRIX(0, 2) = "PRNG": TEMP_MATRIX(0, 3) = "L'Ecuyer"

TEMP_MATRIX(1, 0) = "Time Taken": TEMP_MATRIX(2, 0) = "Min"
TEMP_MATRIX(3, 0) = "Max"
'------------------------------------------------------------------------------------
'Debug.Print "Time taken to generate " & nLOOPS & " random numbers"

T_VAL = Timer
For i = 1 To nLOOPS
  j = Int(Rnd * 1000 + 1) + 3
  TEMP_MATRIX(j, 1) = TEMP_MATRIX(j, 1) + 1
Next i
TEMP_MATRIX(1, 1) = Timer - T_VAL
MIN_VAL = nLOOPS: MAX_VAL = 0
For i = 4 To NROWS + 3
  TEMP_MATRIX(i, 0) = i - 3
  If TEMP_MATRIX(i, 1) < MIN_VAL Then MIN_VAL = TEMP_MATRIX(i, 1)
  If TEMP_MATRIX(i, 1) > MAX_VAL Then MAX_VAL = TEMP_MATRIX(i, 1)
Next i
TEMP_MATRIX(2, 1) = MIN_VAL
TEMP_MATRIX(3, 1) = MAX_VAL

'Debug.Print "Rnd function = " & Format(Timer - T_VAL, "0.0000") & " secs"

'------------------------------------------------------------------------------------

'Dim PRNGKCI_OBJ As New PrngKci
'MsgBox Format(PRNGKCI_OBJ.Rnd(1), "0.000000000000000")

T_VAL = Timer
For i = 1 To nLOOPS
  j = Int(PRNGKCI_OBJ.Rnd(1) * 1000 + 1) + 3
  TEMP_MATRIX(j, 2) = TEMP_MATRIX(j, 2) + 1
Next i
TEMP_MATRIX(1, 2) = Timer - T_VAL

MIN_VAL = nLOOPS
MAX_VAL = 0
For i = 4 To NROWS + 3
  If TEMP_MATRIX(i, 2) < MIN_VAL Then MIN_VAL = TEMP_MATRIX(i, 2)
  If TEMP_MATRIX(i, 2) > MAX_VAL Then MAX_VAL = TEMP_MATRIX(i, 2)
Next i
TEMP_MATRIX(2, 2) = MIN_VAL
TEMP_MATRIX(3, 2) = MAX_VAL

'Debug.Print "PRNG = " & Format(Timer - T_VAL, "0.0000") & " secs"
'------------------------------------------------------------------------------------

T_VAL = Timer

'always start with this to initialise the RNG
Call ECUYER_OBJ.Initialize(31, 41)

'call ECUYER_OBJ.Rnd(1) as often as you like
'you can have up to 100 different RNGs going at once
'you specify which you want with the parameter you pass through here
'if you only want one set, just pass 1 every time

For i = 1 To nLOOPS
  X_VAL = ECUYER_OBJ.Rnd(1)
  j = Int(ECUYER_OBJ.Rnd(1) * 1000 + 1) + 3
  TEMP_MATRIX(j, 3) = TEMP_MATRIX(j, 3) + 1
Next i
TEMP_MATRIX(1, 3) = Timer - T_VAL

MIN_VAL = nLOOPS
MAX_VAL = 0
For i = 4 To NROWS + 3
  If TEMP_MATRIX(i, 3) < MIN_VAL Then MIN_VAL = TEMP_MATRIX(i, 3)
  If TEMP_MATRIX(i, 3) > MAX_VAL Then MAX_VAL = TEMP_MATRIX(i, 3)
Next i
TEMP_MATRIX(2, 3) = MIN_VAL
TEMP_MATRIX(3, 3) = MAX_VAL

'Debug.Print "L'Ecuyer = " & Format(Timer - T_VAL, "0.0000") & " secs"
'------------------------------------------------------------------------------------

PRNG_TABLE_FUNC = TEMP_MATRIX

Exit Function
ERROR_LABEL:
PRNG_TABLE_FUNC = Err.number
End Function
