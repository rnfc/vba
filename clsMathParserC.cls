VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsMathParserC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'********************************************************************************
'* CLASS: clsMathParserC  Complex version                               v.3.4.3 *
'*                                                        Leonardo Volpi        *
'*                                                        Michael Ruder         *
'*                                                        Thomas Zeutschler     *
'*                                                        Lieven Doscshe        *
'*                                                        Arnaud de Grammont    *
'*                                                                              *
'*  Math-Physical Complex Expression Evaluator                                  *
'*  It uses the modules fctComplexes 1.5 and fctMathComplex 1.0 by A.d.G.       *
'********************************************************************************
Option Explicit
#Const DEBUG_MODE = 0
'-------------------------------------------------------------------------------
' CONSTANTS
'-------------------------------------------------------------------------------
Const HiOPER     As Long = 100
Const HiVT       As Long = 100
Const HiET       As Long = 100
Const HiARG      As Long = 6
Const PileFunk   As Long = 100

Const ERROR_RAISE = 0           'Arnaud d. G. 6-12-02

'-------------------------------------------------------------------------------
' TYPE DECLARATIONS
'-------------------------------------------------------------------------------
Private Type T_VTREC           'Variable Table Record
  idx              As Long
  name             As String
  value            As Complexe
  sign             As Integer
End Type
'
Private Type T_ETREC           'Expression Table record
  Fun              As String
  FunTok           As Long
  arg(1 To HiARG)  As T_VTREC
  ArgTop           As Long
  ArgOf            As Long
  ArgIdx           As Long
  value            As Complexe
  sign             As Integer
  PosInExpr        As Long
  PriLvl           As Long
  PriIdx           As Long
  NVarId           As Long
  Cond             As Long
End Type
'
Private Type T_Funk
  FunName           As String
  NbArg             As Long
  NbFoundArg        As Long
  idFunc            As Long
  lastparenthesis   As Boolean
End Type

'-------------------------------------------------------------------------------
' LOCALS
'-------------------------------------------------------------------------------
Dim Expr            As String
Dim ExprOK          As Boolean  'expression OK
Dim ExprNoOK        As Boolean  'expr OK, no variables
Dim Expr1OK         As Boolean  'expr OK, exact 1 var
Dim VT()            As T_VTREC
Dim ET()            As T_ETREC
Dim PF()            As T_Funk
Dim VTtop           As Long
Dim ETtop           As Long
Dim ErrMsg          As String   'error message  11-11-02 VL
Dim funk(1 To HiARG) As String
Dim VarsTbl         As Collection   'additional object (bb 6-1-04)
'-------------------------------------------------------------------------------
'functions lists
Private Sub initFunkString()
    funk(1) = "Abs Atn Cos Exp Fix Int Ln Log Rnd Sgn Sin Sqr Tan Acos Asin " & _
            "Cosh Sinh Tanh Acosh Asinh Atanh Fact Not Erf Gamma Gammaln Digamma Zeta Ei Erfc " & _
            "Re Im Conj Neg Sq Alog Inv Arg"
    funk(2) = "xroot Comb Max Min Mcm Mcd Mod And Or Xor Beta DPoisson CPoisson" + " " + SeparateurComplexe
    funk(3) = ""
    funk(4) = "Integral"
    funk(5) = ""
    funk(6) = "Serie"
End Sub

'-------------------------------------------------------------------------------
' PUBLIC FUCTIONS
'-------------------------------------------------------------------------------

' store expression as array of records; check syntax
Public Function StoreExpression(ByVal strExpr As String) As Boolean
  Expr = Trim(strExpr)
  ExprOK = Parse(Expr)
  ExprNoOK = ExprOK And (VTtop = 0)
  Expr1OK = ExprOK And (VTtop = 1)
  StoreExpression = ExprOK
End Function
'-------------------------------------------------------------------------------
' get the expression
Public Property Get Expression() As String
  Expression = Expr
End Property
'-------------------------------------------------------------------------------
' get the top of the var array (=N-1 bacause starts on 0)
Public Property Get VarTop() As Long
  VarTop = VTtop
End Property
'-------------------------------------------------------------------------------
' get name of a variable by index. VL
Public Property Get VarName(ByVal index As Long) As String
  If index <= VTtop Then
    VarName = VT(index).name
  End If
End Property
'-------------------------------------------------------------------------------
' get value assigned to a variable by index
Public Property Get VarValue(ByVal index As Long) As Complexe
  If index <= VTtop Then
    VarValue = VT(index).value
  End If
End Property
'-------------------------------------------------------------------------------
' assign a value to a certain variable by index
Public Property Let VarValue(ByVal index As Long, VarVal As Complexe)
  If index <= VTtop Then
    VT(index).value = VarVal
  End If
End Property
'-------------------------------------------------------------------------------
' get the error message  11-11-02 VL
Public Property Get ErrorDescription() As String
  ErrorDescription = ErrMsg
End Property
'-------------------------------------------------------------------------------
' evaluate expression
Public Function Eval() As Complexe
Dim ExprVal As Complexe
ErrMsg = ""         'Arnaud d. G. 6-12-02
If Not ExprOK Then GoTo Error_Handler
SubstVars
If Not Eval_(ExprVal) Then GoTo Error_Handler
Eval = ExprVal
Exit Function
'
Error_Handler:
    If ERROR_RAISE = 1 Then Err.Raise 1001, "MathParser", ErrMsg
End Function
'-------------------------------------------------------------------------------
' assign vars Complex vars and evaluate :
Public Function EvalComplexe(VarValues() As Complexe) As Complexe ' AdG
Dim ExprVal As Complexe
ErrMsg = ""         'Arnaud d. G. 6-12-02
If Not ExprOK Then GoTo Error_Handler
SubstVarsComplexe VarValues
If Not Eval_(ExprVal) Then GoTo Error_Handler
EvalComplexe = ExprVal
Exit Function
'
Error_Handler:
    If ERROR_RAISE = 1 Then Err.Raise 1001, "MathParser", ErrMsg
End Function
'-------------------------------------------------------------------------------
' get a variable's value by name (bb 6-1-04)
Public Property Get VarSymb(ByVal name As String) As Complexe
    VarSymb.reel = 0#
    VarSymb.imag = 0#
    On Error Resume Next
    VarSymb = VT(VarsTbl(name)).value
End Property
'-------------------------------------------------------------------------------
' assign a value to a certain variable by name (bb 6-1-04)
Public Property Let VarSymb(ByVal name As String, VarVal As Complexe)
    On Error GoTo Error_Handler
    VT(VarsTbl(name)).value = VarVal
    Exit Property
Error_Handler:
End Property
'-------------------------------------------------------------------------------
' get current setting for angle computing (RAD (default), DEG or GRAD)
Public Property Get AngleUnit() As String
    AngleUnit = getAngleOption
End Property
'-------------------------------------------------------------------------------
' set the unit of measure for angle computing (RAD (default), DEG or GRAD)
Public Property Let AngleUnit(ByVal AngleUM As String)
    setAngleOption AngleUM
End Property
'-------------------------------------------------------------------------------
'class end routine
Private Sub Class_Terminate()
  Set VarsTbl = Nothing           'xxz6
End Sub

'-------------------------------------------------------------------------------
' Math Parser Routine
'         rev 30-08-02 Leonardo Volpi;  rev 20-10-02 L.Dos
Private Function Parse(ByVal strExpr As String) As Boolean
  Dim lExpr          As String
  Dim char           As String * 1
  Dim SubExpr        As String
  Dim lenExpr        As Long
  Dim FunName        As String
  Dim GetNextArg     As Boolean
  Dim SaveArg        As String
  Dim Npart          As Long
  Dim Nabs           As Long
  Dim retval         As Double
  Dim arrPriLvl()    As Long
  Dim srtLo          As Long      'vars for sort algoritme
  Dim srtHi          As Long
  Dim tmp            As Long
  Dim flag_exchanged As Boolean
  Dim i              As Long
  Dim j              As Long
  Dim LogicSymb      As String
  Dim sign           As Integer
  Dim Lastarg        As Integer
  Dim Node_Cond()    As Long        'store the condition-nodes
  Dim Node_Switch()  As Long        'store the switch-node
  Dim Node_max       As Long
  Dim ParseAngle     As Boolean
  '*********************************************************************************
  Dim NbArg As Integer
  Dim FctNV As Boolean
  Dim idFun As Long
  Dim indexfun As Long
  Dim chaine As Boolean
  Dim funknbVar As Integer
  idFun = 0
  indexfun = 0
  chaine = False
  ReDim PF(PileFunk)
  initFunkString
  '*********************************************************************************
  setAngleOption ("Rad")  'Set the angle unit default
  ParseAngle = False

  ReDim ET(HiET)
  ReDim VT(HiVT)
  Set VarsTbl = New Collection   '(bb 6-1-04)
  ETtop = 0
  VTtop = 0
  ErrMsg = vbNullString   'VL
  lExpr = Trim(strExpr)
  '***** abs |.| function counter
  i = NabsCount(lExpr)
  Nabs = i / 2
  If (2 * Nabs <> i) Then
    ErrMsg = "abs symbols |.| mismatch"  'VL
    GoTo ErrHandler
  End If
  '***** begin parse process
  lenExpr = Len(lExpr)
  For i = 1 To lenExpr
   Lastarg = False
   j = 1
   char = Mid$(lExpr, i, 1)
   '*********************************************************************************
   '*                                                                               *
   If chaine = True Then
        If char = ElementChaine Then chaine = False
        SubExpr = SubExpr + Mid$(lExpr, i, 1)
   Else
   '*                                                                               *
   '*********************************************************************************
    Select Case char
      '*********************************************************************************
      '*                                                                               *
      Case ElementChaine
        If ParseAngle = False Then
            chaine = True
        End If
        SubExpr = SubExpr + Mid$(lExpr, i, 1)
      '*                                                                               *
      '*********************************************************************************
      Case "°"
        ParseAngle = True
        SubExpr = SubExpr + Mid$(lExpr, i, 1)
      Case " "                                    '***** skip spaces
      Case "(", "[", "{"                          '***** open parentheses
        Npart = Npart + 1                         'inc # open parentheses
        If LenB(SubExpr) <> 0 Then                     'eval preceding text
          Catch_Sign SubExpr, sign                'catch the function sign (if any) 7-1-04
          If SubExpr = "" And sign = -1 Then
            SubExpr = "Neg": sign = 1             'insert change-sign function 7-1-04
          End If
          If InList(SubExpr, funk(1)) Then          'monovariable function
            ETtop = ETtop + 1                     '   store in ET
            With ET(ETtop)
              .PosInExpr = i                      'position in expr
              .Fun = SubExpr                      'function name
              .FunTok = GetFunTok(SubExpr)        'function Token (enum)
              .PriLvl = Npart * 10                'priority level=open parenth*10
              .ArgTop = 1                         'ntal Args=1
              .sign = sign
            End With
        '*********************************************************************************
        '*                                                                               *
            NbArg = 1
          Else
            NbArg = 0
            For funknbVar = 2 To HiARG
                If InList(SubExpr, funk(funknbVar)) Then
                    NbArg = funknbVar
                End If
            Next
            If NbArg = 0 Then
              ErrMsg = "Function <" + SubExpr + "> unknown at pos " + Str$(i)
              GoTo ErrHandler
            End If
          End If
          
          If NbArg > 1 Then
            FunName = SubExpr
          End If
          indexfun = indexfun + 1
          idFun = idFun + 1
          PF(indexfun).FunName = SubExpr
          PF(indexfun).NbArg = NbArg
          PF(indexfun).NbFoundArg = NbArg
          PF(indexfun).idFunc = idFun
          If NbArg > 2 Then
            FctNV = True
          Else
            FctNV = False
          End If
          SubExpr = vbNullString                           'start parsing for new subexpr
          PF(indexfun).lastparenthesis = False
        Else
          indexfun = indexfun + 1
          PF(indexfun).lastparenthesis = True
        End If
        '*                                                                               *
        '*********************************************************************************
      Case ")", "]", "}"                          '***** open parentheses
        Npart = Npart - 1                         'dec # open parentheses
        If Npart < 0 Then                         'want to close to many brackets
          ErrMsg = "Too many closing brackets at pos " + Str$(i)
          GoTo ErrHandler
        End If
        '*********************************************************************************
        '*
        If (NbArg > 1) Then
            ErrMsg = "Missing argument"
            GoTo ErrHandler
        End If
        If PF(indexfun).lastparenthesis = False Then
            If PF(indexfun).NbArg > 2 Then
                FctNV = True
            Else
                FctNV = False
            End If
            FunName = PF(indexfun).FunName
            NbArg = PF(indexfun).NbFoundArg
        End If
        indexfun = indexfun - 1
        '*                                                                               *
        '*********************************************************************************
      Case "+", "-"                               '*****
        'check the exponential sign (preceding was eg 1.23E of 1.23E-2)
        If CheckExpo(SubExpr) Or SubExpr = vbNullString Then                'fix bug 18-1-03  thanks to Michael Ruder
          SubExpr = SubExpr + char
        Else
          ETtop = ETtop + 1                       'store in ET
          With ET(ETtop)                          'thanks to Javier M. Montalban 14.2.03
            .PosInExpr = i
            .Fun = char
            .FunTok = GetFunTok(char)
            .PriLvl = 2 + Npart * 10
            .ArgTop = 2                           'two arguments
          End With
          GetNextArg = True                       'get second argument
          GoSub Catch_Argument
        End If
      Case "*", "/", "\"                     '*****
        ETtop = ETtop + 1
        With ET(ETtop)
          .PosInExpr = i
          .Fun = char
          .FunTok = GetFunTok(char)
          .PriLvl = 3 + Npart * 10
          .ArgTop = 2                             'two arguments
        End With
        GetNextArg = True
        GoSub Catch_Argument
      Case "^"
        ETtop = ETtop + 1
        With ET(ETtop)
          .PosInExpr = i
          .Fun = "^"
          .FunTok = GetFunTok(char)
          .PriLvl = 4 + Npart * 10
          .ArgTop = 2                             'two arguments
        End With
        GetNextArg = True
        GoSub Catch_Argument
      Case "!"
        ETtop = ETtop + 1
        With ET(ETtop)
          .PosInExpr = i
          .Fun = "!"
          .FunTok = GetFunTok(char)
          .PriLvl = 9 + Npart * 10
          .ArgTop = 1                             'one argument
        End With
        GetNextArg = True
        SaveArg = SubExpr
        GoSub Catch_Argument
        SubExpr = SaveArg
      Case "%"  'percentage
        ETtop = ETtop + 1
        With ET(ETtop)
          .PosInExpr = i
          .Fun = "%"
          .FunTok = GetFunTok(char)
          .PriLvl = 9 + Npart * 10
          .ArgTop = 1                             'one argument
        End With
        GetNextArg = True
        SaveArg = SubExpr
        GoSub Catch_Argument
        SubExpr = SaveArg
      Case SeparateurComplexe
        ETtop = ETtop + 1
        With ET(ETtop)
          .PosInExpr = i
          .Fun = SeparateurComplexe
          .FunTok = GetFunTok(char)
          .PriLvl = 1 + Npart * 10 'change. VL 5-12-01
          .ArgTop = 2                             'two arguments
        End With
        GetNextArg = True
        GoSub Catch_Argument
        SaveArg = SubExpr
      Case ";"                                    'comes from bivariate function f(x;y)
        If LenB(FunName) = 0 Then
          ErrMsg = "syntax error at pos:" + Str$(i)
          GoTo ErrHandler
        End If
        ETtop = ETtop + 1
        With ET(ETtop)
          .PosInExpr = i
          .Fun = FunName                          'previous stored
          .FunTok = GetFunTok(FunName)
          .PriLvl = Npart * 10
          .ArgTop = 2                             'two arguments
        End With
        GetNextArg = True
        GoSub Catch_Argument
        '*********************************************************************************
        '*                                                                               *
        If NbArg = 2 Then
            NbArg = NbArg - 1
            FunName = vbNullString                             'reset function
            If FctNV Then
                ET(ETtop).NVarId = PF(indexfun).idFunc
            Else
                ET(ETtop).NVarId = 0
            End If
        Else
            NbArg = NbArg - 1
            PF(indexfun).NbFoundArg = NbArg
            ET(ETtop).NVarId = PF(indexfun).idFunc
        End If
        '*                                                                               *
        '*********************************************************************************
      Case "|"                                    '***** absolute symbol |.|
        If SubExpr = "" Or SubExpr = "-" Then
          Npart = Npart + 1                       'increment brackets PriLvl
          ETtop = ETtop + 1
          With ET(ETtop)
            .PosInExpr = i
            .Fun = "abs"                          'symbols |.| is similar to  abs(.)
            .FunTok = GetFunTok("abs")
            .PriLvl = Npart * 10
            .ArgTop = 1                           'one argument
            If SubExpr = "-" Then                 'fix sign bug 1.3.04. Thanks to Rodrigo Farinha
                .sign = -1
                SubExpr = ""
            End If
          End With
        Else
          Npart = Npart - 1
          If Npart < 0 Then                       'too many closing brackets
            ErrMsg = "Too many closing brackets at pos " + Str$(i) ' MR 16-01-03
            GoTo ErrHandler
          End If
        End If
      Case "=", "<", ">"                          'Logical operators
        If LenB(LogicSymb) = 0 Then
            If ETtop > 0 Then
                'detect the Interval:=(a < x < b)
                If InStr(1, " > < = <= >= <> =< =>", ET(ETtop).Fun) > 0 Then
                    'transform the Interval into (a<x)*(x<b) form
                    ETtop = ETtop + 1
                    With ET(ETtop)
                      .PosInExpr = i
                      .Fun = "*"                 'insert the hidden multiplication
                      .FunTok = GetFunTok("*")
                      .PriLvl = 3 + (Npart - 1) * 10
                      .ArgTop = 2
                    End With
                    SaveArg = SubExpr
                    GoSub Catch_Argument
                    SubExpr = SaveArg           'pass the argument also to the logic symbol
                End If
            End If
            ETtop = ETtop + 1
            GetNextArg = True
            GoSub Catch_Argument
        End If
        LogicSymb = LogicSymb + char
        With ET(ETtop)
          .PosInExpr = i
          .Fun = LogicSymb                        'logic symbol
          .FunTok = GetFunTok(LogicSymb)
          .PriLvl = 1 + Npart * 10
          .ArgTop = 2                             'two argument
        End With
      Case "x", "y", "z", "X", "Y", "Z"           ''monomial coeff.
        If IsNumeric_(SubExpr) Then               'fix 2.3.2003 thanks to Michael Ruder
            ETtop = ETtop + 1                     'Ex: 7x  is converted into product 7*x
            With ET(ETtop)
              .PosInExpr = i
              .Fun = "*"
              .FunTok = GetFunTok("*")
              .PriLvl = 3 + Npart * 10
              .ArgTop = 2                             'two argument
            End With
            GetNextArg = True
            GoSub Catch_Argument
            i = i - 1  'one step back
        Else
            SubExpr = SubExpr + char
        End If
      Case Else                                   '***** continue parsing
        SubExpr = SubExpr & char
    End Select
   End If
    
    If GetNextArg Then
        GetNextArg = False
    Else
        LogicSymb = vbNullString
    End If
      
  Next
  ' End of the Main LOOP
  
  If Npart > 0 Then                               'parentheses
    ErrMsg = "Not enough closing brackets"
    GoTo ErrHandler
  End If
  If ETtop < 1 Then                               'no operation detected
    ETtop = 1
    With ET(ETtop)
      .PosInExpr = 1
      .Fun = "+"
      .FunTok = GetFunTok("+")
      .PriLvl = 1
      .ArgTop = 2
    End With
  End If
  For i = 1 To ETtop                              'init 2e argument
    ET(i).arg(ET(i).ArgTop) = ET(i + 1).arg(1)
  Next

  If LenB(SubExpr) <> 0 Then
    j = ET(ETtop).ArgTop
    Lastarg = True
    GoSub Catch_Argument
  Else
    ErrMsg = "missing argument"
    GoTo ErrHandler
  End If
  
  If ETtop > 0 Then
    ReDim Preserve ET(0 To ETtop)
  Else
    ReDim Preserve ET(0 To 0)
  End If
  If VTtop > 0 Then
    ReDim Preserve VT(0 To VTtop)
  Else
    ReDim Preserve VT(0 To 0)
  End If

#If DEBUG_MODE = 1 Then  '<<<<<<<<<<<<<<<<<<<<<<<<<<<<
    Parse = True
    Exit Function
#End If                  '<<<<<<<<<<<<<<<<<<<<<<<<<<<<

  GoSub Sort_table
  
  GoSub Build_Relations

#If DEBUG_MODE = 2 Then  '<<<<<<<<<<<<<<<<<<<<<<<<<<<<
    Parse = True
    Exit Function
#End If                  '<<<<<<<<<<<<<<<<<<<<<<<<<<<<

  For i = 1 To ETtop                              'eliminate dependent arguments
    j = ET(i).ArgOf
    If j > 0 Then
      With ET(j).arg(ET(i).ArgIdx)
       .idx = 0
       .name = vbNullString
      End With
    End If
  Next

'* Gestion des fonctions à plus de 2 variables
Dim fctId As Integer
Dim k As Integer
Dim varid As Integer
Dim varfound As Boolean

 For i = 1 To ETtop
    fctId = ET(i).NVarId
    If fctId > 0 Then
        varid = 3 'traitement de la 3ème variable
        ' recherche dans lignes suivantes de la même fonction
        For j = i + 1 To ETtop
            If ET(j).NVarId = fctId Then
                varfound = False
                For k = 1 To ETtop
                    If ET(k).ArgOf = j And ET(k).NVarId <> fctId And ET(k).NVarId <> -1 Then
                        'la variable est le résultat d'une autre ligne
                        varfound = True
                        ET(k).ArgOf = i
                        ET(k).ArgIdx = varid
                    End If
                Next
                If varfound = False Then
                    ' la variable est une constante ou une variable
                    ET(i).arg(varid).idx = ET(j).arg(2).idx
                    ET(i).arg(varid).name = ET(j).arg(2).name
                    ET(i).arg(varid).value = ET(j).arg(2).value
                    ET(i).arg(varid).sign = ET(j).arg(2).sign   'fix sign bug 3.6.2004 VL
                End If
                ET(j).NVarId = -1
                ET(i).ArgOf = ET(j).ArgOf
                ET(i).ArgIdx = ET(j).ArgIdx
                ET(i).ArgTop = varid
                varid = varid + 1 'traitement de la (i+1)ème variable
            End If
        Next
    End If
 Next
 
#If DEBUG_MODE = 3 Then  '<<<<<<<<<<<<<<<<<<<<<<<<<<<<
    Parse = True
    Exit Function
#End If                  '<<<<<<<<<<<<<<<<<<<<<<<<<<<<

 ' supression des lignes inutiles
 j = 1
 For i = 1 To ETtop
    ET(j) = ET(i)
    For k = 1 To ETtop
        If ET(k).ArgOf = i Then
            ET(k).ArgOf = j
        End If
    Next
    If ET(i).NVarId <> -1 Then j = j + 1
 Next
 ETtop = j - 1
 
 ' On trie à nouveau
 GoSub Sort_table
 
 '-------------------------------------------------------------------------------
 'Search for the condition nodes in the graph appling the logic condition rules
    Control_Nodes Node_Cond, Node_Switch, Node_max
    If Node_max > 0 Then GoSub Sort_table '***** start sort algorithm
 '------------------------------------------------------------------------------
 
 Parse = True
 Exit Function
  
'internal routines -----------------------
Catch_Argument:
    If SubExpr = "" Then                        'no next argument found
        ErrMsg = "missing argument"
        GoTo ErrHandler
    End If
    Catch_Sign SubExpr, sign                    'breack the string into name and its sign +/-(if any)
    If SubExpr = "" Then                        'mod 5.4.2004 VL
        ErrMsg = "missing argument"             'fix bug for ++ or -- string. Thanks to PJ Weng
        GoTo ErrHandler
    ElseIf IsPureImmaginary(SubExpr, retval) Then       'check if argument is pure immaginary
        ET(ETtop).arg(j).value.reel = 0
        ET(ETtop).arg(j).value.imag = sign * retval
    ElseIf convEGU(SubExpr, retval) Then                'check if argument is Eng Units
        ET(ETtop).arg(j).value.reel = sign * retval
        ET(ETtop).arg(j).value.imag = 0
    ElseIf convSymbConst(SubExpr, retval) Then          'check if argument is a symbolic constant #
        If ErrMsg <> "" Then GoTo ErrHandler
        ET(ETtop).arg(j).value.reel = sign * retval
        ET(ETtop).arg(j).value.imag = 0
    ElseIf IsNumeric_(SubExpr) Then                     'check if argument is number
        ET(ETtop).arg(j).value.reel = sign * Val(SubExpr)
        ET(ETtop).arg(j).value.imag = 0
    ElseIf cvDegree(SubExpr, retval) Then               'check if argument is ddmmss format degree
        ET(ETtop).arg(j).value.reel = sign * retval
        ET(ETtop).arg(j).value.imag = 0
    Else
      If Not IsLetter(Left(SubExpr, 1)) Then
        ErrMsg = "variable name must start with a letter: " + SubExpr
        GoTo ErrHandler
      End If
      StoreVar SubExpr, Lastarg, sign
      If VTtop > HiVT Then
        ErrMsg = "too many Vars"
        GoTo ErrHandler
      End If
    End If
    SubExpr = ""  'reset the substring
Return

Sort_table:
'sort table with exchanges algorithm
  ReDim arrPriLvl(0 To ETtop)                     'create array with priority levels
  For i = 1 To ETtop                              'and copy then from main array
    arrPriLvl(i) = ET(i).PriLvl
  Next
  For i = 1 To ETtop                              'fill sort order default 0 to ETtop
    ET(i).PriIdx = i
  Next
  srtLo = 1                                       '***** start sort algorithm
  srtHi = ETtop - 1
  Do
    flag_exchanged = False
    For i = srtLo To srtHi Step 2
      j = i + 1
      If arrPriLvl(i) < arrPriLvl(j) Then
        tmp = arrPriLvl(j)
        arrPriLvl(j) = arrPriLvl(i)
        arrPriLvl(i) = tmp
        tmp = ET(j).PriIdx
        ET(j).PriIdx = ET(i).PriIdx
        ET(i).PriIdx = tmp
        flag_exchanged = True
      End If
    Next
    If srtLo = 1 Then
      srtLo = 2
    Else
      srtLo = 1
    End If
  Loop Until (srtLo = 1) And Not flag_exchanged
Return

Build_Relations:
  For i = 1 To ETtop                              'build relations
    j = ET(i).PriIdx
    srtLo = j - 1
    Do While srtLo >= 0
      If ET(srtLo).ArgOf = 0 Then
        Exit Do
      End If
      srtLo = srtLo - 1
    Loop
    srtHi = j + 1
    Do While srtHi <= ETtop
      If ET(srtHi).ArgOf = 0 Then
        Exit Do
      End If
      srtHi = srtHi + 1
    Loop
    If (srtLo < 1) And (srtHi <= ETtop) Then            '
      ET(j).ArgOf = srtHi
      ET(j).ArgIdx = 1
    ElseIf (srtLo > 0) And (srtHi > ETtop) Then        '
      ET(j).ArgOf = srtLo
      ET(j).ArgIdx = ET(srtLo).ArgTop
    ElseIf (srtLo > 0) And (srtHi <= ETtop) Then       '
      If (ET(srtLo).PriLvl) >= (ET(srtHi).PriLvl) Then  'take that one with the upper PriLvl
        ET(j).ArgOf = srtLo
        ET(j).ArgIdx = ET(srtLo).ArgTop
      Else                                        '
        ET(j).ArgOf = srtHi
        ET(j).ArgIdx = 1
      End If
    Else
      Exit For
    End If
  Next

Return
  
  
ErrHandler:
  ETtop = ETtop
  Parse = False
End Function

'-------------------------------------------------------------------------------
Private Function CheckExpo(ByVal SubExpr As String) As Boolean
Dim s_1 As String, s_2 As String, ls As Long
'detect if SubExpr is the mantissa of an expo format number 1.2E+01 , 4E-12, 1.0E-6
    CheckExpo = False
    ls = Len(SubExpr)
    If ls < 2 Then Exit Function
    s_1 = Right(SubExpr, 1)
    s_2 = Left(SubExpr, ls - 1)
    If (UCase(s_1) = "E") And IsNumeric(s_2) Then CheckExpo = True
End Function
'-------------------------------------------------------------------------------
'[modified 10/02 by Thomas Zeutschler]
Private Function Eval_(EvalValue As Complexe) As Boolean
  Dim A(1 To HiARG) As Complexe
  Dim s(1 To HiARG) As String
  Dim Ris   As Complexe
  Dim j     As Long
  Dim k     As Long
  Dim Pos   As Long
  Dim l     As Integer
  Dim m     As Long
  Dim n     As Long
  Dim idx   As Integer
  
  On Error GoTo ErrHandler  '<<< comment for debug  VL 30-8-02
'  setAngleOption ("Rad")
  For j = 1 To ETtop    'Evaluation procedure begins
    k = ET(j).PriIdx
    With ET(k)
      For l = 1 To HiARG
        A(l) = .arg(l).value
        s(l) = .arg(l).name
      Next
      initErrorMsg
      
      Select Case .FunTok
            'Operation
            Case OffsetOper To OffsetOper + 99
                For idx = 1 To 2
                    If s(idx) <> "" Then
                        If Left(s(idx), 1) = ElementChaine Then
                            ErrMsg = "Invalid Arg " & CStr(idx) & " in function <" & CStr(.FunTok) & ">"
                            GoTo ErrHandler
                        End If
                    End If
                Next
                Ris = ComplexOperation(A(1), A(2), .FunTok)
            'Functions
            Case OffsetFunc To OffsetFunc + 99
                If s(1) <> "" Then
                    If Left(s(1), 1) = ElementChaine Then
                        ErrMsg = "Invalid Arg 1 in function <" & CStr(.FunTok) & ">"
                        GoTo ErrHandler
                    End If
                End If
                Ris = ComplexFunction(A(1), .FunTok)
            'Functions N variables
            Case OffsetNVar To OffsetNVar + 99
                Ris = ComplexFunctionNVar(A(), s(), .FunTok)
            Case Else
                ErrMsg = "Function <" & CStr(.FunTok) & "> missing?"  'VL
                GoTo ErrHandler
      End Select
      If getErrorMsg <> "" Then
        ErrMsg = getErrorMsg
        GoTo ErrHandler
      End If
      .value = Ris
      m = .ArgOf
      n = .ArgIdx
      If m = 0 Or n = 0 Then Exit For
      ET(m).arg(n).value = Ris
    End With
  Next
  If getErrorMsg <> "" Then
     ErrMsg = getErrorMsg
     GoTo ErrHandler                        'Arnaud d. G. 6-12-02
  End If
  EvalValue = ET(k).value
  Eval_ = True
  Exit Function
ErrHandler:
  'ErrMsg = "Evaluation error"  'VL         'Arnaud d. G. 6-12-02
  EvalValue.reel = 0
  EvalValue.imag = 0
  Eval_ = False
End Function
'-------------------------------------------------------------------------------
' Assignes a value to symbolic Vars
Private Sub SubstVars()
  Dim i  As Long, j  As Long

  For i = 1 To ETtop
    For j = 1 To HiARG
    With ET(i).arg(j)
        If .idx <> 0 Then
            .value.reel = .sign * VT(.idx).value.reel
            .value.imag = .sign * VT(.idx).value.imag
        End If
    End With
    Next
  Next
End Sub

Private Sub SubstVarsComplexe(VarValues() As Complexe)
  Dim i  As Long, j As Long, Id As Long
  
  For i = 1 To ETtop
    For j = 1 To HiARG
      Id = ET(i).arg(j).idx
      If Id <> 0 Then
        ET(i).arg(j).value = VarValues(Id)
      End If
    Next
  Next
End Sub

Private Sub SubstVarsReel(VarValue() As Double)
  Dim i  As Long, j As Long, Id As Long
  
  For i = 1 To ETtop
    For j = 1 To HiARG
      Id = ET(i).arg(j).idx
      If Id <> 0 Then
        ET(i).arg(j).value.reel = VarValue(Id)
      End If
    Next
  Next
End Sub

Private Sub SubstVarsImag(VarValue() As Double)
  Dim i  As Long, j As Long, Id As Long
    
  For i = 1 To ETtop
    For j = 1 To HiARG
      Id = ET(i).arg(j).idx
      If Id <> 0 Then
        ET(i).arg(j).value.imag = VarValue(Id)
      End If
    Next
  Next
End Sub

' search if var already exists in table, if not add it
Private Sub StoreVar(ByVal SubExpr As String, ByVal Lastarg As Boolean, ByVal sign As Integer)
  Dim VTIdx  As Long
  Dim ArgIdx As Long
  Dim Found  As Boolean
    
  Found = False
  For VTIdx = 1 To VTtop
    If VT(VTIdx).name = SubExpr Then
      Found = True
      Exit For
    End If
  Next
  If Not Found Then
    VTtop = VTtop + 1     'new variable
    If VTtop > HiVT Then  'to many Vars
      Exit Sub
    End If
    VT(VTtop).name = SubExpr
    'add a new variable to the object collection (bb 6-1-04)
    VarsTbl.Add VTtop, SubExpr
  End If
  If Lastarg Then
    ArgIdx = ET(ETtop).ArgTop
  Else
    ArgIdx = 1
  End If
  With ET(ETtop).arg(ArgIdx)
    .name = SubExpr
    .idx = VTIdx
    .sign = sign
  End With
End Sub

'-------------------------------------------------------------------------------
' translate egu to multiplication factor
'  accepts a string contains a measure like "2ms" ,"234.12Mhz", "0.1uF" , 12Km , etc
'  [relaxed parsing: allow space between number and unit and allow numbers without units]
Private Function convEGU(ByVal strSource As String, retval As Double) As Boolean
  Dim EguStr   As String
  Dim EguStart As Long
  Dim EguLen   As Long
  Dim EguMult  As String
  Dim EguCoeff As Double
  Dim EguFact  As Long
  Dim EguSym   As String
  Dim EguBase  As Double
    
  EguStr = strSource      'trim niet nodig; alle spaties zijn weg
  EguLen = Len(EguStr)
  For EguStart = 1 To EguLen
    If IsLetter(Mid(EguStr, EguStart, 1)) Then
      Exit For
    End If
  Next
  If EguStart = 1 Then
'    Debug.Print "missing coefficient"
    convEGU = False
    Exit Function
  ElseIf EguStart > EguLen Then
'    Debug.Print "missing unit of measure"
    convEGU = False
    Exit Function
  End If
  EguCoeff = Left(EguStr, EguStart - 1)   'get number
  EguStr = Mid(EguStr, EguStart)          'extract literal substring
  EguLen = Len(EguStr)
  If EguLen > 1 Then                      'extract multiply factor
    EguMult = Left(EguStr, 1)
    Select Case EguMult
      Case "p":  EguFact = -12
      Case "n":  EguFact = -9
      Case "u":  EguFact = -6
      Case "m":  EguFact = -3
      Case "c":  EguFact = -2
      Case "d":  EguFact = -1
      Case "D":  EguFact = 1
      Case "H":  EguFact = 2
      Case "K":  EguFact = 3
      Case "M":  EguFact = 6
      Case "G":  EguFact = 9
      Case "T":  EguFact = 12
      Case Else: EguFact = 0
    End Select
  Else
    EguFact = 0
  End If
  If EguFact <> 0 Then       'extract um symbol
    EguSym = Mid(EguStr, 2)
  End If
  Select Case EguSym         'check if um exists and compute numeric value
    Case "s":   EguBase = 1                 'second
    Case "Hz":  EguBase = 1                 'frequency
    Case "m":   EguBase = 1                 'meter
    Case "g":   EguBase = 0.001             'gramme
    Case "rad", "Rad", "RAD": EguBase = 1   'radiant  '18-10-02 VL
    Case "S":   EguBase = 1                 'siemens
    Case "V":   EguBase = 1                 'volt
    Case "A":   EguBase = 1                 'ampere
    Case "W":   EguBase = 1                 'watt
    Case "F":   EguBase = 1                 'farad
    Case "bar": EguBase = 1                 'bar
    Case "Pa":  EguBase = 1                 'pascal
    Case "Nm":  EguBase = 1                 'newtonmeter
    Case "Ohm", "ohm": EguBase = 1          'ohm     '18-10-02 VL
    Case Else
      ErrMsg = "unknown unit of measure: " + EguSym
      convEGU = False
      Exit Function
  End Select
  retval = EguCoeff * EguBase * 10 ^ EguFact   'finally compute the numeric value
  convEGU = True
End Function
'-------------------------------------------------------------------------------
'check if it is a number
Private Function IsNumeric(ByVal A) As Boolean
  Dim x As Double
  
  On Error Resume Next
  x = A 'assign "a" to a double
  IsNumeric = Err.number <= 0
End Function
'-------------------------------------------------------------------------------
'check if it is a letter
Private Function IsLetter(ByVal char As String) As Boolean
  Dim code As Long
    
  code = Asc(char)
  'IsLetter = (65 <= code And code <= 90) Or (97 <= code And code <= 122)
  '*********************************************************************************
  IsLetter = (65 <= code And code <= 90) Or (97 <= code And code <= 122) Or char = ElementChaine
  '*********************************************************************************
End Function
'-------------------------------------------------------------------------------
'check for an expression to occur in a list
Private Function InList(ByVal strElem As String, ByVal strList As String) As Boolean
  Dim lstrElem As String
  Dim lstrList As String
    
  lstrList = " " & strList & " "
  lstrElem = " " & strElem & " "
  InList = InStr(1, lstrList, lstrElem, vbTextCompare) > 0
End Function
'-------------------------------------------------------------------------------
' count number of abs sybol sets in formula
Private Function NabsCount(ByVal s As String) As Long
  Dim n As Long
  Dim p As Long

  n = 0
  p = InStr(1, s, "|")
  Do While p > 0
    p = p + 1
    n = n + 1
    p = InStr(p, s, "|")
  Loop
  NabsCount = n
End Function
'-------------------------------------------------------------------------------
' check if a string may be a pure immaginary number.
Private Function IsPureImmaginary(ByVal strSource As String, retval As Double) As Boolean
Dim temp As String
    IsPureImmaginary = False
    If Right(strSource, 1) <> "i" And Right(strSource, 1) <> "j" Then Exit Function
    temp = (Left(strSource, Len(strSource) - 1))
    If temp = "-" Then temp = "-1"
    If temp = "+" Or temp = "" Then temp = "+1"
    If IsNumeric(temp) Then     'may be a complex-immaginary number Es  2i or 1.33j
        retval = Val(temp)
        IsPureImmaginary = True
    End If
End Function

'break the variable string into the name and its sign (if any). Es -x
Private Sub Catch_Sign(Str As String, sign As Integer)
Dim s As String, VarName As String
sign = 1
s = Left(Str, 1)
If s = "-" Or s = "+" Then
    Str = Right(Str, Len(Str) - 1)
    If s = "-" Then sign = -sign
End If
End Sub

Private Function IsNumeric_(ByVal x As String) As Boolean
  'independent check function from international system setting
  'x must have always the decimal point "123.756", ".0056", "1.3455E-12"
  'by LV, 27.2.2003. Thanks to Michael Ruder
  Dim p As Integer
  IsNumeric_ = False
  If InStr(1, x, ",") Then Exit Function
  IsNumeric_ = IsNumeric(x)
  If Not IsNumeric_ Then
  'try with decimal comma
  p = InStr(1, x, ".")
  If p > 0 Then
  Mid$(x, p, 1) = ","
  IsNumeric_ = IsNumeric(x)
  End If
  End If
End Function

'-------------------------------------------------------------------------------
'Search for the condition nodes in the graph appling the logic condition rules
Private Sub Control_Nodes(Node_Cond() As Long, Node_Switch() As Long, Node_max As Long)
Dim n&, i&, j&, k&, p&, count_iter&, Node_dup As Boolean
Dim Node_aux() As Long, Node_aux1() As Long
Dim Ns&, nc&, Level_max&

n = UBound(ET)
ReDim Node_Cond(1 To n), Node_Switch(1 To n), Node_aux(1 To n)
j = 0
For i = 1 To n
    If InStr(1, " > < = <= >= <> =< =>", ET(i).Fun) > 0 Then
        Node_aux(i) = 2
        'search for duplicate nodes
        Node_dup = False
        For p = 1 To j
            If Node_Switch(p) = ET(i).ArgOf Then
                Node_dup = True
                Exit For
            End If
        Next
        'apply rule
        If Node_dup Then
            Node_Cond(p) = Node_Switch(p)
            Node_Switch(p) = ET(Node_Cond(p)).ArgOf
            Node_aux(Node_Cond(p)) = 2
        Else
            j = j + 1
            Node_Cond(j) = i
            Node_Switch(j) = ET(i).ArgOf
        End If
    End If
    If Level_max < ET(i).PriLvl Then Level_max = ET(i).PriLvl
Next i
Node_max = j
If Node_max = 0 Then Exit Sub 'no logic function detected
ReDim Preserve Node_Cond(1 To Node_max), Node_Switch(1 To Node_max)

For p = 1 To Node_max
    'load control node information
    nc = Node_Cond(p)
    Ns = Node_Switch(p)
    If ET(Ns).Fun = "*" Then
        ReDim Node_aux1(1 To n)
        For i = 1 To n: Node_aux1(i) = Node_aux(i): Next i
        'apply the graph-condition-rules to the table
        count_iter = 0
        For k = 1 To n
            i = k: j = 0
            Do
                If Node_aux1(i) = -1 Or ET(i).ArgOf = 0 Then
                    'node k independent
                    Node_aux1(k) = -1
                    Exit Do
                End If
                If Node_aux1(i) = 2 Then
                    'assign higher priority to the condition node and its childs
                    ET(k).PriLvl = ET(k).PriLvl + 100
                    Exit Do
                End If
                If Node_aux1(i) = 1 Or ET(i).ArgOf = Ns Then
                    'node k dependent
                    Node_aux1(k) = 1
                    ET(k).Cond = nc
                    Exit Do
                End If
                i = ET(i).ArgOf
                j = j + 1
            Loop Until j > n
            count_iter = count_iter + j
        Next k
    End If
Next p

End Sub

'-------------------------------------------------------------------------------
' translate a symbolic Constant to its double value
Private Function convSymbConst(ByVal strSource As String, retval As Double) As Boolean
  Dim CostToken   As String
  Dim SymbConst   As String
  CostToken = "#"
  convSymbConst = False:  ErrMsg = ""
  'check if string is "pi" only for compatibility with previous release.
  If LCase(strSource) = "pi" Then strSource = strSource + CostToken
  If Right(strSource, 1) <> CostToken Then Exit Function
  retval = 0
  SymbConst = Left(strSource, Len(strSource) - 1)
  Select Case SymbConst
    Case "pi", "PI":  retval = ValPi                'pi-greek
    Case "pi2": retval = ValPi / 2                  'pi-greek/2
    Case "pi3": retval = ValPi / 3                  'pi-greek/3
    Case "pi4": retval = ValPi / 4                  'pi-greek/4
    Case "e":   retval = 2.71828182845905           'Euler-Napier
    Case "eu":  retval = 0.577215664901533          'Euler-Mascheroni
    Case "phi": retval = 1.61803398874989           'golden ratio
    Case "g":   retval = 9.80665                    'Acceleration due to gravity
    Case "G":   retval = 6.672 * 10 ^ -11           'Gravitational constant
    Case "R":   retval = 8.31451                    'Gas constant
    Case "eps": retval = 8.854187817 * 10 ^ -12     'Permittivity of vacuum
    Case "mu":  retval = 12.566370614 * 10 ^ -7     'Permeability of vacuum
    Case "c":   retval = 2.99792458 * 10 ^ 8        'Speed of light
    Case "q":   retval = 1.60217733 * 10 ^ -19      'Elementary charge
    Case "me":  retval = 9.1093897 * 10 ^ -31       'Electron rest mass
    Case "mp":  retval = 1.6726231 * 10 ^ -27       'Proton rest mass
    Case "K":   retval = 1.380658 * 10 ^ -23        'Boltzmann constant
    Case "h":   retval = 6.6260755 * 10 ^ -34       'Planck constant
    Case Else
        ErrMsg = "Constant unknown: " + SymbConst
  End Select
  convSymbConst = True
End Function

Private Function cvDegree(ByVal DMS As String, ByRef Angle As Double) As Boolean
'converts a string dd°mm'ss" (degrees, minutes, seconds) into a decimal-degree angle
Dim p1 As Long, p2 As Long, P3 As Long
Dim DD As Double, mm As Double, ss As Double
    cvDegree = False
    If Left$(DMS, 1) = ElementChaine Then Exit Function 'AdG
    p1 = InStr(1, DMS, "°") ' degrees °
    p2 = InStr(1, DMS, "'")  ' minutes '
    P3 = InStr(1, DMS, """")  ' seconds "
    If p1 = 0 And p2 = 0 And P3 = 0 Then Exit Function
    On Error Resume Next
    DD = CDbl(Mid(DMS, 1, p1 - 1))
    mm = CDbl(Mid(DMS, p1 + 1, p2 - p1 - 1))
    ss = CDbl(Mid(DMS, p2 + 1, P3 - p2 - 1))
    On Error GoTo 0
    If mm > 60 Or ss > 60 Then Exit Function
    Angle = DD + (mm + ss / 60) / 60
    cvDegree = True

End Function

'use only for debug
Sub ET_Dump(ByRef ETable)
  ReDim ETable(ETtop, 30)
  Dim i As Long, j As Long
  
  j = j + 1: ETable(0, j) = "Fun"
  j = j + 1: ETable(0, j) = "ArgTop"
  j = j + 1: ETable(0, j) = "A1 Idx"
  j = j + 1: ETable(0, j) = "Arg1 Name"
  j = j + 1: ETable(0, j) = "Arg1 Value"
  j = j + 1: ETable(0, j) = "A2 Idx"
  j = j + 1: ETable(0, j) = "Arg2 Name"
  j = j + 1: ETable(0, j) = "Arg2 Value"
  j = j + 1: ETable(0, j) = "ArgOf"
  j = j + 1: ETable(0, j) = "ArgIdx"
  j = j + 1: ETable(0, j) = "Value"
  j = j + 1: ETable(0, j) = "PriLvl"
  j = j + 1: ETable(0, j) = "PosInExpr"
  j = j + 1: ETable(0, j) = "PriIdx"
  j = j + 1: ETable(0, j) = "Cond"
  j = j + 1: ETable(0, j) = "NVarId"
  
  ReDim Preserve ETable(ETtop, j)
  For i = 1 To UBound(ET)
    j = 0
    With ET(i)
        j = j + 1: ETable(i, j) = .Fun
        j = j + 1: ETable(i, j) = .ArgTop
        j = j + 1: ETable(i, j) = .arg(1).idx
        j = j + 1: ETable(i, j) = .arg(1).name
        j = j + 1: ETable(i, j) = .arg(1).value.reel
        j = j + 1: ETable(i, j) = .arg(2).idx
        j = j + 1: ETable(i, j) = .arg(2).name
        j = j + 1: ETable(i, j) = .arg(2).value.reel
        j = j + 1: ETable(i, j) = .ArgOf
        j = j + 1: ETable(i, j) = .ArgIdx
        j = j + 1: ETable(i, j) = .value.reel
        j = j + 1: ETable(i, j) = .PriLvl
        j = j + 1: ETable(i, j) = .PosInExpr
        j = j + 1: ETable(i, j) = .PriIdx
        j = j + 1: ETable(i, j) = .Cond
        j = j + 1: ETable(i, j) = .NVarId
    End With
  Next
End Sub


