Attribute VB_Name = "STAT_PROCESS_FOURIER_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.


'************************************************************************************
'************************************************************************************
'FUNCTION      : FOURIER_TRANSFORM_FUNC
'DESCRIPTION   : Detecting periodicities in a time series, by transforming
'data from the time domain into the frequency domain.
'LIBRARY       : STATISTICS
'GROUP         : FOURIER
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/08/2008
'************************************************************************************
'************************************************************************************

Function FOURIER_TRANSFORM_FUNC(ByRef REAL_RNG As Variant, _
ByRef IMAG_RNG As Variant, _
ByVal EXPONENT As Long, _
Optional ByVal VERSION As Integer = 1, _
Optional ByRef INDEX_RNG As Variant)

Dim g As Double '
Dim i As Double
Dim j As Double
Dim k As Double
Dim l As Double

Dim E As Double
Dim f As Double

Dim o As Double
Dim p As Double
Dim q As Double

Dim NROWS As Double
Dim NSIZE As Double

Dim PI_VAL As Double

Dim REAL_SUM As Double
Dim IMAG_SUM As Double
Dim TEMP_FACTOR As Double

Dim STEMP_REAL As Double
Dim STEMP_IMAG As Double

Dim TTEMP_REAL As Double
Dim TTEMP_IMAG As Double

Dim UTEMP_REAL As Double
Dim UTEMP_IMAG As Double

Dim VTEMP_REAL As Double
Dim VTEMP_IMAG As Double

Dim REAL_VECTOR As Variant
Dim IMAG_VECTOR As Variant
Dim INDEX_VECTOR As Variant

Dim TEMP_MATRIX As Variant

On Error GoTo ERROR_LABEL

PI_VAL = 3.14159265358979

REAL_VECTOR = REAL_RNG
    If UBound(REAL_VECTOR, 1) = 1 Then: _
    REAL_VECTOR = MATRIX_TRANSPOSE_FUNC(REAL_VECTOR)
IMAG_VECTOR = IMAG_RNG
    If UBound(IMAG_VECTOR, 1) = 1 Then: _
        IMAG_VECTOR = MATRIX_TRANSPOSE_FUNC(IMAG_VECTOR)
    If UBound(IMAG_VECTOR, 1) <> UBound(REAL_VECTOR) Then: GoTo ERROR_LABEL

NSIZE = 2 ^ EXPONENT
ReDim TEMP_MATRIX(1 To NSIZE, 1 To 2) ' array of NSIZE rows, 2 columns _
for real and imag

'------------------------------------------------------------------------
'Filtering. Once the fourier series of a data set has been determined,
'synthetic data that exclude or filter particular frequencies can be
'generated to allow further analyis. For example the low frequency
'annual and monthly variations can be removed from data to reveal the
'weekly and daily variations.

'Smoothing data. This is a special case of filtering where high frequency
'data is filtered. It can for example be used to smooth out the daily and
'weekly variation in a data set but leave the longer term monthly and
'annual variations.

'Developing mathematical models of data sets.

'This Fourier Transforms function is not the fastest, but reasonable and
'for serious applications with large datasets anybody would choose a more
'appropriate environment. This here works up to 2^14 data as input.

'To step in: just view at FFT and inverse FFT as a black box technique to
'compute certain sums which come from discretization Fourier transforms.
'So consider them as discrete Fourier transforms DFT and inverse DFT:

'z = x + y*I
'------------------------------------------------------------------------
Select Case VERSION
'--------------------------------------------------------------------------
Case 0 'Fourier Transform
'--------------------------------------------------------------------------
                
    For i = 1 To NSIZE
      TEMP_MATRIX(i, 1) = REAL_VECTOR(i, 1) ' x(i) = REAL_RNG(i)
      TEMP_MATRIX(i, 2) = IMAG_VECTOR(i, 1) ' y(i) = IMAG_RNG(i)
    Next i
                
    f = NSIZE
    For g = 1 To EXPONENT
        E = f
        f = E / 2
        UTEMP_REAL = 1
        UTEMP_IMAG = 0
        VTEMP_REAL = (Cos(PI_VAL / f))
        VTEMP_IMAG = (-Sin(PI_VAL / f))
        For j = 1 To f
            For i = j To NSIZE Step E
                q = i + f
                TTEMP_REAL = TEMP_MATRIX(i, 1) + TEMP_MATRIX(q, 1)
                TTEMP_IMAG = TEMP_MATRIX(i, 2) + TEMP_MATRIX(q, 2)
      
                STEMP_REAL = TEMP_MATRIX(i, 1) - TEMP_MATRIX(q, 1)
                STEMP_IMAG = TEMP_MATRIX(i, 2) - TEMP_MATRIX(q, 2)
      
                TEMP_MATRIX(q, 1) = STEMP_REAL * UTEMP_REAL - _
                STEMP_IMAG * UTEMP_IMAG
                
                TEMP_MATRIX(q, 2) = STEMP_REAL * UTEMP_IMAG + _
                STEMP_IMAG * UTEMP_REAL
                            
                TEMP_MATRIX(i, 1) = TTEMP_REAL
                TEMP_MATRIX(i, 2) = TTEMP_IMAG
            Next i
            TTEMP_REAL = UTEMP_REAL * VTEMP_REAL - _
            UTEMP_IMAG * VTEMP_IMAG
                        
            UTEMP_IMAG = UTEMP_REAL * VTEMP_IMAG + _
            UTEMP_IMAG * VTEMP_REAL
                        
            UTEMP_REAL = TTEMP_REAL
        Next j
    Next g

    o = NSIZE / 2
    p = NSIZE - 1
    j = 1
    For i = 1 To p
        If i < j Then
            TTEMP_REAL = TEMP_MATRIX(j, 1)
            TTEMP_IMAG = TEMP_MATRIX(j, 2)
                        
            TEMP_MATRIX(j, 1) = TEMP_MATRIX(i, 1)
            TEMP_MATRIX(j, 2) = TEMP_MATRIX(i, 2)
    
            TEMP_MATRIX(i, 1) = TTEMP_REAL
            TEMP_MATRIX(i, 2) = TTEMP_IMAG
        End If
        k = o
        Do While k < j
            j = j - k
            k = k / 2
        Loop
        j = j + k
    Next i
'--------------------------------------------------------------------------
Case 1 ' Inverse
'--------------------------------------------------------------------------
                
    For i = 1 To NSIZE 'do scaling by 1/NSIZE
      TEMP_MATRIX(i, 1) = REAL_VECTOR(i, 1) / NSIZE ' x(i) = REAL_RNG(i)
      TEMP_MATRIX(i, 2) = IMAG_VECTOR(i, 1) / NSIZE ' y(i) = IMAG_RNG(i)
    Next i

    f = NSIZE
    For g = 1 To EXPONENT
        E = f
        f = E / 2
        UTEMP_REAL = 1
        UTEMP_IMAG = 0
  
        VTEMP_REAL = (Cos(PI_VAL / f))
        VTEMP_IMAG = (Sin(PI_VAL / f))
  
        For j = 1 To f
            For i = j To NSIZE Step E
                q = i + f
                TTEMP_REAL = TEMP_MATRIX(i, 1) + TEMP_MATRIX(q, 1)
                TTEMP_IMAG = TEMP_MATRIX(i, 2) + TEMP_MATRIX(q, 2)
      
                STEMP_REAL = TEMP_MATRIX(i, 1) - TEMP_MATRIX(q, 1)
                STEMP_IMAG = TEMP_MATRIX(i, 2) - TEMP_MATRIX(q, 2)
      
                TEMP_MATRIX(q, 1) = STEMP_REAL * UTEMP_REAL - _
                STEMP_IMAG * UTEMP_IMAG
                
                TEMP_MATRIX(q, 2) = STEMP_REAL * UTEMP_IMAG + _
                STEMP_IMAG * UTEMP_REAL
      
                TEMP_MATRIX(i, 1) = TTEMP_REAL
                TEMP_MATRIX(i, 2) = TTEMP_IMAG
            Next i
            TTEMP_REAL = UTEMP_REAL * VTEMP_REAL - UTEMP_IMAG * VTEMP_IMAG
            UTEMP_IMAG = UTEMP_REAL * VTEMP_IMAG + UTEMP_IMAG * VTEMP_REAL
            UTEMP_REAL = TTEMP_REAL
        Next j
    Next g

    o = NSIZE / 2
    p = NSIZE - 1
    j = 1
                
    For i = 1 To p
        If i < j Then
            TTEMP_REAL = TEMP_MATRIX(j, 1)
            TTEMP_IMAG = TEMP_MATRIX(j, 2)
    
            TEMP_MATRIX(j, 1) = TEMP_MATRIX(i, 1)
            TEMP_MATRIX(j, 2) = TEMP_MATRIX(i, 2)
    
            TEMP_MATRIX(i, 1) = TTEMP_REAL
            TEMP_MATRIX(i, 2) = TTEMP_IMAG
  
        End If
        k = o
        Do While k < j
            j = j - k
            k = k / 2
        Loop
        j = j + k
    Next i
                
    'For i = 1 To NSIZE
        '  TEMP_MATRIX(i, 1) = TEMP_MATRIX(i, 1) / NSIZE _
            'x(i) = x(i) / NSIZE
        '  TEMP_MATRIX(i, 2) = TEMP_MATRIX(i, 2) / NSIZE _
            'y(i) = y(i) / NSIZE
    'Next i
                                
'--------------------------------------------------------------------------
Case 2 'Check FFT
'--------------------------------------------------------------------------

    If IsArray(INDEX_RNG) = True Then
        INDEX_VECTOR = INDEX_RNG
        If UBound(INDEX_VECTOR, 1) = 1 Then: _
        INDEX_VECTOR = MATRIX_TRANSPOSE_FUNC(INDEX_VECTOR)
        NROWS = UBound(INDEX_VECTOR, 1)
    Else
        NROWS = UBound(REAL_VECTOR, 1)
        ReDim INDEX_VECTOR(1 To NROWS, 1 To 1)
        For l = 1 To NROWS
            INDEX_VECTOR(l, 1) = l
        Next l
    End If
    
    ReDim TEMP_MATRIX(1 To NROWS, 1 To 2)
    
    For l = 1 To NROWS
        REAL_SUM = 0
        IMAG_SUM = 0
        For j = 1 To NSIZE
              TEMP_FACTOR = 2 * PI_VAL * (j - 1) * (INDEX_VECTOR(l, 1) - 1) / NSIZE
              REAL_SUM = REAL_VECTOR(j, 1) * Cos(TEMP_FACTOR) + _
                    IMAG_VECTOR(j, 1) * Sin(TEMP_FACTOR) + REAL_SUM
              IMAG_SUM = IMAG_VECTOR(j, 1) * Cos(TEMP_FACTOR) - _
                    REAL_VECTOR(j, 1) * Sin(TEMP_FACTOR) + IMAG_SUM
        Next j
        TEMP_MATRIX(l, 1) = REAL_SUM
        TEMP_MATRIX(l, 2) = IMAG_SUM
    Next l
    
'--------------------------------------------------------------------------
Case Else 'Check Inverse FFT
'--------------------------------------------------------------------------

    If IsArray(INDEX_RNG) = True Then
        INDEX_VECTOR = INDEX_RNG
        If UBound(INDEX_VECTOR, 1) = 1 Then: _
        INDEX_VECTOR = MATRIX_TRANSPOSE_FUNC(INDEX_VECTOR)
        NROWS = UBound(INDEX_VECTOR, 1)
    Else
        NROWS = UBound(REAL_VECTOR, 1)
        ReDim INDEX_VECTOR(1 To NROWS, 1 To 1)
        For l = 1 To NROWS
            INDEX_VECTOR(l, 1) = l
        Next l
    End If
    
    ReDim TEMP_MATRIX(1 To NROWS, 1 To 2)
    
    For l = 1 To NROWS
        REAL_SUM = 0
        IMAG_SUM = 0
        
        For j = 1 To NSIZE
              TEMP_FACTOR = 2 * PI_VAL * (j - 1) * (INDEX_VECTOR(l, 1) - 1) / NSIZE
              REAL_SUM = REAL_VECTOR(j, 1) * Cos(TEMP_FACTOR) - _
                    IMAG_VECTOR(j, 1) * Sin(TEMP_FACTOR) + REAL_SUM
              IMAG_SUM = IMAG_VECTOR(j, 1) * Cos(TEMP_FACTOR) + _
                    REAL_VECTOR(j, 1) * Sin(TEMP_FACTOR) + IMAG_SUM
        Next j
        
        TEMP_MATRIX(l, 1) = REAL_SUM / NSIZE
        TEMP_MATRIX(l, 2) = IMAG_SUM / NSIZE
    Next l

End Select

FOURIER_TRANSFORM_FUNC = TEMP_MATRIX
'First Column Real; Second Column Imaginary

Exit Function
ERROR_LABEL:
FOURIER_TRANSFORM_FUNC = Err.number
End Function
