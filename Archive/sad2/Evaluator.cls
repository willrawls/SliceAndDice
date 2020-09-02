VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Evaluator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Copyright© 1999, Tretyakov Konstantin
'_____________________________________________________
'This is the 'Evaluator' class: it inputs a string
'like "2+2" or "2+4*sin(3.4)^2-8*arccos(0.55)", etc
'_____________________________________________________
'You may use the code for free, if you give me credit.
'If you modify it or make your own program with it,
'I would VERY APPRECIATE, if you mail me it (or better-
'a link to it)
'On the whole - just do not stamp your name on what you haven't
'done quite alone.
'This code was written totally by me, and 'it took me about
'2 days to code it (and about a year
'-that is,from the very moment I got interested in programming-
'I spent dreaming of having such a thing)

'(BTW this code seems to be quite unique-
'I searched all over the Internet for such, but NOONE
'is giving the source for such things)
'______________________________________________________
'Yours Sincerely, Konstantin Tretyakov (kt_ee@yahoo.com)

'********************Here we go...********************
'Well, at the very beginning (when I had only + and -)
'These constants didplay a role:
'e.g. I could change the PLUS_SIGN to "plus"
'and the MINUS_SIGN to "minus", so that I could
'write an expression like "1 plus 2 minus 3"
'But now it will not go. :(
Const PLUS_SIGN = "+"
Const MINUS_SIGN = "-"
Const MULTIPLY_SIGN = "*"
Const DIVIDE_SIGN = "/"
Const POWER_SIGN = "^"
Const POINT_SIGN = ","
Const BRACKET_LEFT = "("
Const BRACKET_RIGHT = ")"

'This is the part to be improved - I mean this error-handling
Public Enum EvalError
    ERR_NONE = 0
    ERR_DBL_POINT = 1
    ERR_WRONG_SYNTAX = 2
    ERR_WRONG_SIGN = 4
    ERR_WRONG_BRACKETS = 8
    ERR_WRONG_FUNCTION = 16
End Enum

'This entry was needed for my other project - Function Analyzer
'(look for it at the same place, where you found this one)
Private m_Assigned As Boolean
'I hope you get, what these do
Private m_Expression As String
Private m_Result As Double
Private m_Error As EvalError

Public Property Let Expression(ByVal NewExpr As String)
1        m_Expression = ReplaceText(UCase(RemoveSpaces(NewExpr)), ".", POINT_SIGN)
End Property
Public Property Get Expression() As String
1        Expression = m_Expression
End Property
Public Property Get Error() As EvalError
1        Error = m_Error
End Property

Public Property Get Result() As Double
'Reset the Error
1        m_Error = ERR_NONE
    'Calculate
2        m_Result = Eval(m_Expression)
3        m_Assigned = (m_Error = ERR_NONE)
    'Return
4        Result = m_Result
End Property
Public Property Get Assigned() As Boolean
1        Assigned = m_Assigned
End Property
Public Function Evaluate(ByVal Expressn As String, Optional ByVal Silent As Boolean = False) As Double
'That's the wrapper for the main procedure
'You may use this class in 2 ways:
'1) Set the 'Expression' property and then read the 'Result' property
'2) Call this sub. If you set Silent to False, then the sub will generate a message automatically
1        Dim Res As Double
2        Expression = Expressn
3        Res = Result
4        If Not Silent Then
5            If m_Error <> ERR_NONE Then
            Select Case m_Error
                Case ERR_DBL_POINT: MsgBox "Error: Wrong decimal separator placement!", vbCritical, "Eval Error"
6                    Case ERR_WRONG_BRACKETS: MsgBox "Error: Wrong bracket placement!", vbCritical, "Eval Error"
7                    Case ERR_WRONG_SIGN: MsgBox "Error: Wrong sign or bracket placement!", vbCritical, "Eval Error"
8                    Case ERR_WRONG_SYNTAX: MsgBox "Error: Wrong syntax!", vbCritical, "Eval Error"
9                End Select
10           Else
11               MsgBox "Result: " & Res, vbExclamation, "Eval Result"
12           End If
13       End If
14       Evaluate = m_Result
End Function

'***********************************************************
' 2 helper functions, well they are too 'universal' for this class
' (Here we use them only to remove spaces and replace the '.' to ','

Private Function RemoveSpaces(S$) As String
1        RemoveSpaces = ReplaceText(S$)
End Function

Public Function ReplaceText(ByVal SourceText$, Optional ByVal StrToReplace$ = " ", Optional ByVal StrToInsert$ = vbNullString) As String
1        Dim RetS$, I%
2        If StrToReplace = StrToInsert Or StrToReplace = vbNullString Then Exit Function
3        RetS = SourceText$
4        I = InStr(RetS, StrToReplace)
5        Do While I <> 0
6            RetS = IIf(I = 1, vbNullString, Left$(RetS, I - 1)) & StrToInsert$ & IIf(I = Len(RetS) - Len(StrToReplace) + 1, vbNullString, Right$(RetS, Len(RetS) - I - Len(StrToReplace) + 1))
7            I = InStr(RetS, StrToReplace)
8        Loop
9        ReplaceText = RetS
End Function
'***********************************************************

'The HEART of the class.
'What it does? - it just splits the expression to monomials
'(that is: 2*3+3^(3-2)-(2+3) has 3 monomials:
'      +2*3,  +3^(3-2)  -(2+3)
'Then it calls the CalcMonomial for each and sums the result

Private Function Eval(ByVal Expr As String) As Double
1        Dim sEval$, I&, MonomArray As Variant, dResult As Double
2        sEval = Expr
3        MonomArray = SplitToMonomials(sEval)
4        For I = LBound(MonomArray) To UBound(MonomArray)
5            dResult = dResult + CalcMonomial(MonomArray(I))
6        Next
7        Eval = dResult
End Function



Private Function SplitToMonomials(ByVal EvalStr As String, Optional ByVal Sign1 As String = PLUS_SIGN, Optional ByVal Sign2 As String = MINUS_SIGN) As Variant
'Divides the given string in parts using the given sign (Sign1 and Sign2) parameter

'Returns an array where each item is a string
'For example SplitToMonomials("2+3*8-4","+","-") returns [2, +3*8, -4]
'        and SplitToMonomials("3*2/23","*","/") returns [3, *2, /23]

'The function also doesn't split brackets so that
'      SplitToMonominals("(3+2)*2-3","+","-") will return [(3+2)*2, -3]

1        Dim MonomArray As Variant, I&, Count&
2        Dim CurMonom As String, sEval As String
3        ReDim MonomArray(0)
4        sEval = EvalStr
    'Find the first PLUS or MINUS (MUL or DIV) that are not in Bracket
    '(GetSplitPos is Just an Improved Instr, that considers brackets)
5        I = GetSplitPos(EvalStr, Sign1, Sign2)
6        Do While I > 0
        'NOT DONE:
        'Check for expressions of a kind: "2-3*4+6*-5"
        'because we must not split between 6 and 5
7            CurMonom = Left$(sEval, I - 1)
        'Populate the Array
8            ReDim Preserve MonomArray(Count)
9            MonomArray(Count) = CurMonom
10           Count = Count + 1
11           sEval = Mid$(sEval, I)
12           I = GetSplitPos(sEval, Sign1, Sign2)
13       Loop

14       CurMonom = sEval
15       ReDim Preserve MonomArray(Count)
16       MonomArray(Count) = CurMonom
17       SplitToMonomials = MonomArray
End Function

'Calculates a monomial (expression without PLUSes and MINUSes inside)
'The work is in fact like of the Eval function:
'We split it to smaller parts (the ones, that may contain only the ^ sign)
'and then Calculate each part separately
Private Function CalcMonomial(ByVal Monomial As String) As Double
1        On Error GoTo ErrCalcMember
2        If m_Error <> ERR_NONE Then Exit Function
3        Dim MemberArray As Variant, Sign As String
4        Dim I&, dResult As Double, TempRes As Double
    'Split again, but now by * and /
5        MemberArray = SplitToMonomials(Monomial, MULTIPLY_SIGN, DIVIDE_SIGN)
6        For I = LBound(MemberArray) To UBound(MemberArray)
7            TempRes = CalcMember(MemberArray(I), Sign)
        Select Case Sign
                'Remember - we may have the Plus_sign left in a monomial
                '(like a monomial may be "+2^2*3")
            Case PLUS_SIGN: dResult = dResult + TempRes
8                Case MULTIPLY_SIGN: dResult = dResult * TempRes
9                Case DIVIDE_SIGN: dResult = dResult / TempRes
10           End Select
11       Next
12       CalcMonomial = dResult
13       Exit Function
14 ErrCalcMember:
15       m_Error = ERR_WRONG_FUNCTION
End Function

'Calculates an expression, that contains only the operands
'higher in proirity than * and /

'TODO: It raises an error on X^Y^Z and calculates only X^Y,
'That is, for correct calculation you must specify either (X^Y)^Z
'or X^(Y^Z) (btw which is right ???)
Private Function CalcMember(ByVal Member As String, ByRef Sign As String) As Double
1        Dim sSign As String, sEval As String, HaveMinus As Boolean, GotNum1 As Boolean
2        Dim Num1 As Double, Num2 As Double, Op As String, dResult As Double
3        Dim Func As String, FuncExpr As String
4        If m_Error <> ERR_NONE Then Exit Function
    'Here we calculate the results of operations
    'whose priority is higher than * and /
    'The sample given string may be: "+5^2", "*4^2", "/6", "6^2,3"
    'or +(expr)^2, or (expr)^(expr)
5        Sign = PLUS_SIGN
6        sEval = Member
7        sSign = Left$(sEval, 1)
    'Determine the Sign (or find the Bracket or a function)
8        If Not IsNumeric(sSign) Then
        Select Case sSign
            Case MINUS_SIGN
9                    HaveMinus = True
10                   sEval = Mid$(sEval, 2)
11                   If Left$(sEval, 1) = BRACKET_LEFT Then GoTo LBrack
12                   If IsNumeric(Left$(sEval, 1)) = False Then GoTo HaveFunc
13               Case PLUS_SIGN, MULTIPLY_SIGN, DIVIDE_SIGN
14                   Sign = sSign
15                   sEval = Mid$(sEval, 2)
16                   If Left$(sEval, 1) = BRACKET_LEFT Then GoTo LBrack
17                   If IsNumeric(Left$(sEval, 1)) = False Then GoTo HaveFunc
18               Case BRACKET_LEFT
19 LBrack:
                'That's easy - when we find a bracket - we just 'Eval' the expression in the brackets
20                   Num1 = Eval(ExtractBrackets(sEval))
21                   GotNum1 = True
22               Case Else
                'Here Must make some checks for Functions (like when it's SIN(expr))
23 HaveFunc:
24                   Func = ExtractFunction(sEval, FuncExpr)
25                   Num1 = CalcFunction(Func, FuncExpr)
26                   GotNum1 = True
27           End Select
28       End If
    'Now Do the Calculation
29       If Not GotNum1 Then Num1 = ExtractNumber(sEval)
30       If Len(sEval) <> 0 Then
31           Op = Left$(sEval, 1)
32           sEval = Mid$(sEval, 2)
        'Check if the second number is a bracketed expression
33           If Left$(sEval, 1) = BRACKET_LEFT Then
34               Num2 = Eval(ExtractBrackets(sEval))
35           Else
36               If IsNumeric(Left$(sEval, 1)) = False Then
37                   Func = ExtractFunction(sEval, FuncExpr)
38                   Num2 = CalcFunction(Func, FuncExpr)
39               Else
40                   Num2 = ExtractNumber(sEval)
41               End If
42           End If
        Select Case Op
            Case POWER_SIGN
43                   On Error GoTo ErrCalcMember
44                   dResult = Num1 ^ Num2
45               Case Else
46                   m_Error = ERR_WRONG_SIGN
47           End Select
48       Else
49           dResult = Num1
50       End If
51       If Len(sEval) <> 0 Then m_Error = ERR_WRONG_SYNTAX
52       CalcMember = IIf(HaveMinus, -dResult, dResult)
53       Exit Function
54 ErrCalcMember:
55       m_Error = ERR_WRONG_FUNCTION
End Function

'***********************************************************
'This is nearly an equivalent of VAL,
'only here we may know if there was an error
'and it also modifies the string by removing the "Extracted" number

'TODO: It doesn't support the "2.34E+2" notation
Private Function ExtractNumber(ByRef EvalExpr$) As Double
1        Dim HavePoint As Boolean, I As Integer, NewNum As String
2        Dim TempChar As String, TempSign As String, HaveMinus As Boolean
3        Dim sEval As String
    'Determine whether there is a sign in front of the string
4        TempSign = Left$(EvalExpr, 1)
5        If TempSign = POINT_SIGN Then
6            sEval = "0" & EvalExpr
7        Else
8            If Not IsNumeric(TempSign) Then
9                sEval = Mid$(EvalExpr, 2)
10               HaveMinus = (TempSign = MINUS_SIGN)
11           Else: sEval = EvalExpr
12           End If
13       End If

14       For I = 1 To Len(sEval)
15           TempChar = Mid$(sEval, I, 1)
16           If IsNumeric(TempChar) Then
17               NewNum = NewNum & TempChar
18           Else
19               If TempChar = POINT_SIGN Then
20                   If HavePoint Then
                    'We have already a point, that's an error
21                       m_Error = ERR_DBL_POINT
22                       Exit For
23                   Else
24                       HavePoint = True
25                       NewNum = NewNum + "."             'We shall use val in the end
26                   End If
27               Else
28                   Exit For
29               End If
30           End If
31       Next
32       If Len(NewNum) = 0 Then
33           m_Error = ERR_WRONG_SYNTAX
34       Else                                              'Cut out the number from the string
35           EvalExpr = Mid$(sEval, Len(NewNum) + 1)
36       End If
37       ExtractNumber = IIf(HaveMinus, -Val(NewNum), Val(NewNum))
End Function


'***********************************************************
'This is a Helper-func to SplitToMonomials
'it returns the position in a string of a Sign(1 or 2)
'it doesn't return the signs that are in brackets and the sign on the 1st place
Private Function GetSplitPos(ByVal EvalStr$, ByVal Sign1$, ByVal Sign2$, Optional StartPos As Integer = 1)
1        Dim I%, InBracket%, TempChar$

2        For I = StartPos To Len(EvalStr$)
3            TempChar = Mid$(EvalStr, I, 1)
        Select Case TempChar
            Case Sign1, Sign2
4                    If InBracket = 0 And I > 1 Then
5                        GetSplitPos = I
6                        Exit Function
7                    End If
8                Case BRACKET_LEFT
9                    InBracket = InBracket + 1
10               Case BRACKET_RIGHT
11                   InBracket = InBracket - 1
12                   If InBracket < 0 Then
13                       m_Error = ERR_WRONG_BRACKETS
14                       Exit Function
15                   End If
16           End Select
17       Next
End Function

'Gets a String, beginning with a Left Bracket and
'returns the expression in this bracket
'deletes this expression(with both brackets) from the string
Private Function ExtractBrackets(ByRef EvalExpr As String) As String
1        Dim InBracket%, I&, TempChar$, RetStr$
    'We Suppose that the first sign in the string is BRACKET_LEFT
2        InBracket = 1
3        For I = 2 To Len(EvalExpr)
4            TempChar = Mid$(EvalExpr, I, 1)
        Select Case TempChar
            Case BRACKET_LEFT
5                    InBracket = InBracket + 1
6                Case BRACKET_RIGHT
7                    InBracket = InBracket - 1
8            End Select
9            If InBracket = 0 Then
10               RetStr = Mid$(EvalExpr, 2, I - 2)
11               EvalExpr = Mid$(EvalExpr, I + 1)
12               ExtractBrackets = RetStr
13               Exit Function
14           End If
15       Next
16       m_Error = ERR_WRONG_BRACKETS
End Function

'Process the expression "FUNC(expr)"
'Returns "FUNC"
Private Function ExtractFunction(ByRef EvalExpr As String, ByRef FuncExpr As String)
1        Dim FuncID As String, I&
2        I = InStr(EvalExpr, BRACKET_LEFT)
3        If I = 0 Then
4            m_Error = ERR_WRONG_SYNTAX
5            Exit Function
6        Else
7            ExtractFunction = Left$(EvalExpr, I - 1)
8            EvalExpr = Mid$(EvalExpr, I)
9            FuncExpr = ExtractBrackets(EvalExpr)
10       End If
End Function

'You give it a function name and an expression in the brackets after it
'as 2 separate strings, and it calculates
'ADD ANY of the Functions you like
'(E.G. it's interesting to add some 'acting' functions, like, say, MsgBox :)
'Then there are only several steps towards your own Script-Language
Private Function CalcFunction(ByVal FunctionID As String, ByVal FuncExpr As String) As Double
1        On Error GoTo ErrCalc
2        If m_Error <> ERR_NONE Then Exit Function
3        Dim Arg As Double
4        Arg = Eval(FuncExpr)
    Select Case FunctionID
        Case "ABS"
5                CalcFunction = Abs(Arg)
6            Case "ATN"
7                CalcFunction = Atn(Arg)
8            Case "COS"
9                CalcFunction = Cos(Arg)
10           Case "EXP"
11               CalcFunction = Exp(Arg)
12           Case "FIX"
13               CalcFunction = Fix(Arg)
14           Case "INT"
15               CalcFunction = Int(Arg)
16           Case "LOG"
17               CalcFunction = Log(Arg)
18           Case "RND"
19               CalcFunction = Rnd(Arg)
20           Case "SGN"
21               CalcFunction = Sgn(Arg)
22           Case "SIN"
23               CalcFunction = Sin(Arg)
24           Case "SQR"
25               CalcFunction = Sqr(Arg)
26           Case "TAN"
27               CalcFunction = Tan(Arg)
            'Derived
28           Case "SEC"
29               CalcFunction = 1 / Cos(Arg)
30           Case "COSEC"
31               CalcFunction = 1 / Sin(Arg)
32           Case "COTAN"
33               CalcFunction = 1 / Tan(Arg)
34           Case "ARCSIN"
35               CalcFunction = Atn(Arg / Sqr(-Arg * Arg + 1))
36           Case "ARCCOS"
37               CalcFunction = Atn(-Arg / Sqr(-Arg * Arg + 1)) + 2 * Atn(1)
38           Case "ARCSEC"
39               CalcFunction = Atn(Arg / Sqr(Arg * Arg - 1)) + Sgn(Arg - 1) * (2 * Atn(1))
40           Case "ARCCOSEC"
41               CalcFunction = Atn(Arg / Sqr(Arg * Arg - 1)) + (Sgn(Arg) - 1) * (2 * Atn(1))
42           Case "ARCCOTAN"
43               CalcFunction = Atn(Arg) + 2 * Atn(1)
44           Case "HSIN"
45               CalcFunction = (Exp(Arg) - Exp(-Arg)) / 2
46           Case "HCOS"
47               CalcFunction = (Exp(Arg) + Exp(-Arg)) / 2
48           Case "HTAN"
49               CalcFunction = (Exp(Arg) - Exp(-Arg)) / (Exp(Arg) + Exp(-Arg))
50           Case "HSEC"
51               CalcFunction = 2 / (Exp(Arg) + Exp(-Arg))
52           Case "HCOSEC"
53               CalcFunction = 2 / (Exp(Arg) - Exp(-Arg))
54           Case "HCOTAN"
55               CalcFunction = (Exp(Arg) + Exp(-Arg)) / (Exp(Arg) - Exp(-Arg))
56           Case "HARCSIN"
57               CalcFunction = Log(Arg + Sqr(Arg * Arg + 1))
58           Case "HARCCOS"
59               CalcFunction = Log(Arg + Sqr(Arg * Arg - 1))
60           Case "HARCTAN"
61               CalcFunction = Log((1 + Arg) / (1 - Arg)) / 2
62           Case "HARCSEC"
63               CalcFunction = Log((Sqr(-Arg * Arg + 1) + 1) / Arg)
64           Case "HARCCOSEC"
65               CalcFunction = Log((Sgn(Arg) * Sqr(Arg * Arg + 1) + 1) / Arg)
66           Case "HARCCOTAN"
67               CalcFunction = Log((Arg + 1) / (Arg - 1)) / 2
            'Not Math functions, but also useful
68           Case "TIMER"
69               CalcFunction = Timer
70           Case "YEAR"
71               CalcFunction = Year(Now)
72           Case "MONTH"
73               CalcFunction = Month(Now)
74           Case "DAY"
75               CalcFunction = Day(Now)
76           Case "WEEKDAY"
77               CalcFunction = Weekday(Now)
78           Case "HOUR"
79               CalcFunction = Hour(Time)
80           Case "MINUTE"
81               CalcFunction = Minute(Time)
82           Case "SECOND"
83               CalcFunction = Second(Time)
            'These should be constants, but here you must use them as functions
            '(i.e. with an argument, no matter what)
84           Case "PI"
85               CalcFunction = 3.14159265358979
86           Case "E"
87               CalcFunction = 2.71828182845905
88           Case "ZERO"
89               CalcFunction = 0
90           Case Else
91               m_Error = ERR_WRONG_SYNTAX
92       End Select
93       Exit Function
94 ErrCalc:
95       m_Error = ERR_WRONG_FUNCTION
End Function

