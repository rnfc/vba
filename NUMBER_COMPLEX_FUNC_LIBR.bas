Attribute VB_Name = "NUMBER_COMPLEX_FUNC_LIBR"
'********************************************************************************
'* Complex Math module                                 by Arnaud de Grammont    *
'* v. 1.0, 13-01-2003                                                           *
'********************************************************************************
Option Explicit
Option Private Module

Function ComplexFunctionNVar(x() As Complexe, V() As String, name As Long) As Complexe
    Dim Y As Complexe
    Select Case name
        Case symInteg
            Y = Romberg(V(1), V(2), x(3), x(4))
        Case symSerie
            Y = Serie(V(1), V(2), x(3), x(4), V(5), x(6))
        Case Else
            setErrorMsg "Function <" & name & "> missing?"
    End Select
    ComplexFunctionNVar = Y
End Function

Function Romberg(chaine As String, variable As String, A As Complexe, B As Complexe) As Complexe
    Dim Ris As Complexe
    Dim Funct As New clsMathParserC
    Dim ok As Boolean
    Dim index As Integer
    Dim NombreVariables As Long
    Dim VarName As String
    Dim VarValue() As Complexe
    Dim indexvar As Integer
    Const Rank = 16
    Const ErrMax = 10 ^ -15
    Dim r() As Complexe
    Dim Y() As Complexe
    Dim ErrLoop
    Dim i&, Nodes&, n%
    Dim deux As Complexe
    Dim h As Complexe
    Dim s As Complexe
    Dim y1 As Complexe
    Dim y2 As Complexe
    Dim denom As Complexe
    Dim j As Integer

On Error GoTo ErrorMsg

    chaine = Mid(chaine, 2, Len(chaine) - 2)
    
    If Len(variable) >= 2 Then variable = Mid(variable, 2, Len(variable) - 2)
    ok = Funct.StoreExpression(chaine)
    If Not ok Then
        setErrorMsg Funct.ErrorDescription
        Exit Function
    End If
    
    indexvar = 0
    NombreVariables = Funct.VarTop
    If NombreVariables > 0 Then
        ReDim VarValue(1 To NombreVariables)
        For index = 1 To NombreVariables
            VarName = Funct.VarName(index)
            If LCase(VarName) = LCase(variable) Then
                indexvar = index
            Else
                If Left(VarName, 1) <> ElementChaine Then
                        setErrorMsg "One Variable Only"
                        Exit Function
                End If
            End If
        Next
    End If
    
    
    If indexvar = 0 Then
        Ris = Funct.Eval
        If Funct.ErrorDescription <> "" Then
            setErrorMsg Funct.ErrorDescription
            Exit Function
        End If
        Romberg = fois(moins(B, A), Ris)
    Else
        n = 0
        Nodes = 1
        ReDim r(Rank, Rank), Y(Nodes)
        
        VarValue(indexvar) = A
        Ris = Funct.EvalComplexe(VarValue)
        If Funct.ErrorDescription <> "" Then
            setErrorMsg Funct.ErrorDescription
            Exit Function
        End If
        Y(0) = Ris
        VarValue(indexvar) = B
        Ris = Funct.EvalComplexe(VarValue)
        If Funct.ErrorDescription <> "" Then
            setErrorMsg Funct.ErrorDescription
            Exit Function
        End If
        Y(1) = Ris
        deux.reel = 2
        deux.imag = 0
        h = moins(B, A)
        r(n, n) = fois(h, divis(plus(Y(0), Y(1)), deux))
        Do
            n = n + 1
            Nodes = 2 * Nodes
            h = divis(h, deux)
            ReDim Preserve Y(Nodes)
            For i = Nodes To 1 Step -1
                If i Mod 2 = 0 Then
                    Y(i) = Y(i / 2)
                Else
                    VarValue(indexvar).reel = A.reel + i * h.reel
                    VarValue(indexvar).imag = A.imag + i * h.imag
                    Ris = Funct.EvalComplexe(VarValue)
                    If Funct.ErrorDescription <> "" Then
                        setErrorMsg Funct.ErrorDescription
                        Exit Function
                    End If
                    Y(i) = Ris
                End If
            Next i
            s.reel = 0
            s.imag = 0
            For i = 1 To Nodes
                s = plus(plus(s, Y(i)), Y(i - 1))
            Next
            r(n, 0) = divis(fois(h, s), deux)
            For j = 1 To n
                y1 = r(n - 1, j - 1)
                y2 = r(n, j - 1)
                denom.reel = 4 ^ j - 1
                denom.imag = 0
                r(n, j) = plus(y2, divis(moins(y2, y1), denom))
            Next j
            ErrLoop = absol(moins(r(n, n), r(n, n - 1))).reel
            If absol(r(n, n)).reel > 10 Then
                ErrLoop = ErrLoop / absol(r(n, n)).reel
            End If
        Loop Until ErrLoop < ErrMax Or n >= Rank
        Romberg = r(n, n)
    End If
    Exit Function
    
ErrorMsg:
    If getErrorMsg <> "" Then setErrorMsg "Syntax Error"
End Function

Function Serie(chaine As String, variableN As String, N1 As Complexe, N2 As Complexe, variableX As String, x As Complexe) As Complexe
    Dim Ris As Complexe
    Dim Funct As New clsMathParserC
    Dim ok As Boolean
    Dim index As Integer
    Dim NombreVariables As Long
    Dim VarName As String
    Dim VarValue() As Complexe
    Dim indexvarX As Integer
    Dim indexvarN As Integer
    Dim n_min As Integer
    Dim N_MAX As Integer
    Dim k As Integer

On Error GoTo ErrorMsg

        chaine = Mid(chaine, 2, Len(chaine) - 2)
    
    If Len(variableN) >= 2 Then variableN = Mid(variableN, 2, Len(variableN) - 2)
    If Len(variableX) >= 2 Then variableX = Mid(variableX, 2, Len(variableX) - 2)
    
    ok = Funct.StoreExpression(chaine)
    If Not ok Then
        setErrorMsg Funct.ErrorDescription
        Exit Function
    End If
    
    indexvarX = 0
    indexvarN = 0
    NombreVariables = Funct.VarTop
    If NombreVariables > 0 Then
        ReDim VarValue(1 To NombreVariables)
        For index = 1 To NombreVariables
            VarName = Funct.VarName(index)
            If LCase(VarName) = LCase(variableX) Then
                indexvarX = index
                VarValue(index) = x
            ElseIf LCase(VarName) = LCase(variableN) Then
                indexvarN = index
            Else
                If Left(VarName, 1) <> ElementChaine Then
                        setErrorMsg "One Variable Only"
                        Exit Function
                End If
            End If
        Next
    End If
    
    Serie.reel = 0
    Serie.imag = 0
    n_min = N1.reel
    N_MAX = N2.reel
    VarValue(indexvarN).imag = 0
    For k = n_min To N_MAX
        If indexvarN <> 0 Then VarValue(indexvarN).reel = k
        Ris = Funct.EvalComplexe(VarValue)
        If Funct.ErrorDescription <> "" Then
            setErrorMsg Funct.ErrorDescription
            Exit Function
        End If
        Serie = plus(Serie, Ris)
    Next
    Exit Function

ErrorMsg:
    If getErrorMsg <> "" Then setErrorMsg "Syntax Error"
End Function

