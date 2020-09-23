VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BasicLanguage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements SCLangEngine
Private m_colKeyWords As Collection
Private m_colOperators As Collection
Private r_gslLanguage As SCLanguage

Private Function SCLangEngine_StandardParse(ByVal Expression As String) As SemanticResult
    SCLangEngine_StandardParse = SemanticParse(LexicalParse(Expression))
End Function

Private Function SCLangEngine_SemanticParse(LexicalInput As LexicalResult) As SemanticResult
    SCLangEngine_SemanticParse = SemanticParse(LexicalInput)
End Function

Private Function SCLangEngine_LexicalParse(ByVal Expression As String) As LexicalResult
    SCLangEngine_LexicalParse = LexicalParse(Expression)
End Function

Private Property Get SCLangEngine_Language() As SCLanguage
    Set SCLangEngine_Language = r_gslLanguage
End Property

Private Property Set SCLangEngine_Language(ByVal v_gslLanguage As SCLanguage)
    Set r_gslLanguage = v_gslLanguage
    If Not v_gslLanguage Is Nothing Then
        With v_gslLanguage
            .AddKeyWord m_gl_CST_strKwdAlias, m_gl_CST_strKwdAlias
            .AddKeyWord "Open", "Open"
            .AddKeyWord "Close", "Close"
            .AddKeyWord "Exit", "Exit"
            .AddKeyWord "Output", "Output"
            .AddKeyWord "Debug", "Debug"
            .AddKeyWord "Print", "Print"
            .AddKeyWord "Binary", "Binary"
            .AddKeyWord "Random", "Random"
            .AddKeyWord "Access", "Access"
            .AddKeyWord "Implements", "Implements"
            .AddKeyWord "If", "If"
            .AddKeyWord "ElseIf", "ElseIf"
            .AddKeyWord "#End", "#End"
            .AddKeyWord "#ElseIf", "#ElseIf"
            .AddKeyWord "#Else", "#Else"
            .AddKeyWord "Else", "Else"
            .AddKeyWord "Then", "Then"
            .AddKeyWord "Not", "Not"
            .AddKeyWord "Preserve", "Preserve"
            .AddKeyWord "Redim", "Redim"
            .AddKeyWord "Goto", "Goto"
            .AddKeyWord "Resume", "Resume"
            .AddKeyWord "Error", "Error"
            .AddKeyWord "New", "New"
            .AddKeyWord "On", "On"
            .AddKeyWord "Is", "Is"
            .AddKeyWord "Nothing", "Nothing"
            .AddKeyWord "True", "True"
            .AddKeyWord "False", "False"
            .AddKeyWord "Define Directive", "#define"
            .AddKeyWord "#if", "#if"
            .AddKeyWord "type", "type"
            .AddKeyWord "Const", "Const"
            .AddKeyWord "Enum", "Enum"
            .AddKeyWord "Sub", "Sub"
            .AddKeyWord "Let", "Let"
            .AddKeyWord "Optional", "Optional"
            .AddKeyWord "Set", "Set"
            .AddKeyWord "Get", "Get"
            .AddKeyWord "Or", "Or"
            .AddKeyWord "And", "And"
            .AddKeyWord "Function", "Function"
            .AddKeyWord "Select", "Select"
            .AddKeyWord "Case", "Case"
            .AddKeyWord "LBound", "LBound"
            .AddKeyWord "UBound", "UBound"
            .AddKeyWord "Boolean", "Boolean"
            .AddKeyWord "ByVal", "ByVal"
            .AddKeyWord "ByRef", "ByRef"
            .AddKeyWord "End", "End"
            .AddKeyWord "DConst", "#const"
            .AddKeyWord "Variable Const", "VConst"
            .AddKeyWord "Do", "Do"
            .AddKeyWord "Dimention Variable", "Dim"
            .AddKeyWord "With Block", "With"
            .AddKeyWord "While", "While"
            .AddKeyWord "Until", "Until"
            .AddKeyWord "Loop", "Loop"
            .AddKeyWord "WEnd", "WEnd"
            .AddKeyWord "Declare", "Declare"
            .AddKeyWord "To", "To"
            .AddKeyWord "For", "For"
            .AddKeyWord "Each", "Each"
            .AddKeyWord "In", "In"
            .AddKeyWord "Next", "Next"
            .AddKeyWord "Step", "Step"
            .AddKeyWord "Option", "Option"
            .AddKeyWord "Explicit", "Explicit"
            .AddKeyWord "Remainder", "Mod"
            .AddKeyWord "Method Call", "Call"
            .AddKeyWord "Convert To Integer", "CInt"
            .AddKeyWord "Convert To Single", "CSng"
            .AddKeyWord "Convert To Long", "CLng"
            .AddKeyWord "Convert To Byte", "CByt"
            .AddKeyWord "Convert To Double", "CDbl"
            .AddKeyWord "Private Declarator", "Private"
            .AddKeyWord "Public Declarator", "Public"
            .AddKeyWord "WithEvents", "WithEvents"
            .AddKeyWord "Property", "Property"
            .AddKeyWord "Long data type", "Long"
            .AddKeyWord "Integer data type", "Integer"
            .AddKeyWord "Byte data type", "Byte"
            .AddKeyWord "String data type", "String"
            .AddKeyWord "Double data type", "Double"
            .AddKeyWord "Library", "Lib"
            .AddKeyWord "Define 'As' Type", "As"
            .AddOperator "Group Selective Bracket - Left", "(", 0
            .AddOperator "Group Selective Bracket - Right", ")", 0
            .AddOperator "Word Selective Bracket - Left", "[", 0
            .AddOperator "Word Selective Bracket - Right", "]", 0
            .AddOperator "Colon", ":"
            .AddOperator "Plus", "+", 3
            .AddOperator "Minus", "-", 3
            .AddOperator "Percent", "%", 3
            .AddOperator "Minus", "*", 2
            .AddOperator "Division", "/", 2
            .AddOperator "Integer Division", "\", 2
            .AddOperator "Also", "&", 4
            .AddOperator "Object.Property Seperator", "."
            .AddOperator "Item List (Seperator, Item)", ","
            .AddOperator "Object!Property Get/Let/Set Seperator", "!"
            .AddOperator "Greater Than", ">", 5
            .AddOperator "Less Than", "<", 5
            .AddOperator "Equals", "=", 5
            .AddOperator ">=", ">=", 5
            .AddOperator "<=", "<=", 5
            .AddOperator "=>", "=>", 5
            .AddOperator "=<", "=<", 5
            .AddOperator "Power Of", "^", 1
            .AddOperator "Number", "#"
            .AddOperator "_", "_"
        End With
    End If
    FixLanguage
End Property

Public Property Get Language() As SCLanguage
Attribute Language.VB_Description = "Reference to the language"
    Set Language = r_gslLanguage
End Property

Public Property Set Language(ByVal v_gslLanguage As SCLanguage)
    Set r_gslLanguage = v_gslLanguage
End Property

Friend Function LexicalParse(ByVal Expression As String) As LexicalResult
Attribute LexicalParse.VB_Description = "Parses text into a token result"
    Dim m_lngPosition As Long
    Dim m_glpProcess As LexicalProcess
    Dim m_lprResult As LexicalProcResult
    Dim m_glrResult As LexicalResult
    m_glpProcess.Expression = Expression
    m_glpProcess.ExpressionLength = Len(Expression)
    m_lngPosition = 1
    m_glpProcess.CharIndex = 1
    Do
        m_glpProcess.Position = m_lngPosition
        With m_lprResult
            m_lprResult = LexicalKeyWord(m_glpProcess)
            If Not .Success Then _
                m_lprResult = LexicalConstant(m_glpProcess)
            If Not .Success Then _
                m_lprResult = LexicalIdentifier(m_glpProcess)
            If Not .Success Then _
                m_lprResult = LexicalOperator(m_glpProcess)
            If Not .Success Then _
                m_lprResult = LexicalString(m_glpProcess)
            If Not .Success Then _
                m_lprResult = LexicalComment(m_glpProcess)
            If Not .Success Then _
                m_lprResult = LexicalWhiteSpace(m_glpProcess)
            If Not .Success Then
                m_lngPosition = m_lngPosition + 1
                    '//This should NEVER happen.
            ElseIf .Success Then
                AddToken m_lprResult.Token, m_glrResult.Tokens
                m_lngPosition = .NewPosition
            End If
        End With
        'Form1.ProgressUpdate m_glpProcess.Position, m_glpProcess.ExpressionLength
    Loop Until m_lngPosition > Len(Expression)
    'Form1.ProgressUpdate m_glpProcess.ExpressionLength, m_glpProcess.ExpressionLength
    LexicalParse = m_glrResult
End Function

Friend Function SemanticParse(LexicalInput As LexicalResult) As SemanticResult
    
End Function

Friend Function StandardParse(ByVal Expression As String) As SemanticResult
    StandardParse = SemanticParse(LexicalParse(Expression))
End Function

Private Function LexicalString(Process As LexicalProcess) As LexicalProcResult
Declarations:
    Dim m_lngPosition As Long
    Dim m_strChar As String * 1
    Dim m_lprResult As LexicalProcResult

    Dim m_booFoundQuote As Boolean
    Dim m_chiInfo As CharInfo
Try:
    On Error GoTo Catch
    m_lngPosition = Process.Position
    Do
        m_strChar = Mid$(Process.Expression, m_lngPosition, 1)
        m_chiInfo = GetCharInfo(m_strChar)
        If Process.CharIndex = 1 Then
            If Not m_strChar = m_gl_CST_strOprQuote Then
                GoTo Failure
            End If
        Else
            If m_booFoundQuote Then
                If IsOperator(m_strChar, m_colOperators) Or m_strChar = vbCr Or m_strChar = vbLf Or m_strChar = m_gl_CST_strOprSpace Then
                    Process.CharIndex = Process.CharIndex - 1
                    GoTo Success
                ElseIf m_strChar = m_gl_CST_strOprQuote Then
                    m_booFoundQuote = False
                Else
                    GoTo Failure
                End If
            ElseIf m_strChar = m_gl_CST_strOprQuote Then
                m_booFoundQuote = Not m_booFoundQuote
                If m_lngPosition = Process.ExpressionLength Then
                    Process.CharIndex = Process.CharIndex + 1
                    GoTo Success
                End If
            End If
        End If
        Process.CharIndex = Process.CharIndex + 1
        m_lngPosition = m_lngPosition + 1
    Loop Until m_lngPosition > Process.ExpressionLength
Finally:
    GoTo EndTry
Success:
    With m_lprResult
        .Success = True
        With .Token
            .TokenType = G_LTT_String
            .Value = Mid$(Process.Expression, Process.Position, Process.CharIndex)
            .Length = Len(.Value)
        End With
        .NewPosition = Process.Position + Process.CharIndex
    End With
    LexicalString = m_lprResult
    GoTo EndTry
Failure:
    m_lprResult.Success = False
    GoTo EndTry
Catch:
    m_lprResult.Success = False
    GoTo EndTry
EndTry:
    Process.CharIndex = 1
    Exit Function
End Function

Private Function LexicalComment(Process As LexicalProcess) As LexicalProcResult
Declarations:
    Dim m_lngPosition As Long
    Dim m_strChar As String * 1
    Dim m_lprResult As LexicalProcResult

    Dim m_chiInfo As CharInfo
    Dim m_lngNextItem As Long
    Dim m_lngNextUnderItem As Long
    Dim m_lngCRPos As Long
Try:
    On Error GoTo Catch
    m_lngPosition = Process.Position
    m_strChar = Mid$(Process.Expression, m_lngPosition, 1)
    m_chiInfo = GetCharInfo(m_strChar)
    If Process.CharIndex = 1 Then
        If m_chiInfo.IsAlpha Or m_chiInfo.IsNumeric Then
            GoTo Failure
        ElseIf m_strChar = m_gl_CST_strOprApostrophe Then
            GoTo Success
        End If
    End If
    Process.CharIndex = Process.CharIndex + 1
Finally:
    GoTo EndTry
Success:
    With m_lprResult
        m_lngNextItem = Process.Position
        m_lngNextUnderItem = m_lngNextItem
        m_lngCRPos = m_lngNextItem
        Do
            m_lngNextItem = NextCarrageReturn(Process.Expression, m_lngCRPos)
            m_lngNextUnderItem = InStr(m_lngCRPos, Process.Expression, " _" & vbCr)
            m_lngCRPos = m_lngNextItem + 1
        Loop Until m_lngNextItem = 0 Or (m_lngNextUnderItem > m_lngNextItem) Or m_lngNextUnderItem = 0
        If m_lngNextItem = 0 Then
            Process.CharIndex = (Process.ExpressionLength - (Process.Position - 1))
        Else
            Process.CharIndex = ((m_lngNextItem - 1) - (Process.Position - 1))
        End If
        .Success = True
        .NewPosition = Process.Position + Process.CharIndex
        With .Token
            .Value = Mid$(Process.Expression, Process.Position, Process.CharIndex)
            .TokenType = G_LTT_Comment
            .Length = Len(.Value)
        End With
    End With
    LexicalComment = m_lprResult
    GoTo EndTry
Failure:
    m_lprResult.Success = False
    GoTo EndTry
Catch:
    m_lprResult.Success = False
    GoTo EndTry
EndTry:
    Process.CharIndex = 1
    Exit Function
End Function

Private Function LexicalConstant(Process As LexicalProcess) As LexicalProcResult
Declarations:
    Dim m_lngPosition As Long
    Dim m_strChar As String * 1
    Dim m_lprResult As LexicalProcResult
    Dim m_booAndLast As Boolean
    Dim m_chiInfo As CharInfo
    Dim m_booIsNegative As Boolean
    Dim m_booIsHex As Boolean
Try:
    On Error GoTo Catch
    m_lngPosition = Process.Position
    Do
        m_strChar = Mid$(Process.Expression, m_lngPosition, 1)
        m_chiInfo = GetCharInfo(m_strChar)
        If Process.CharIndex = 1 Then
            If (m_strChar = m_gl_CST_strOprMinus Or m_strChar = m_gl_CST_strOprAnd) And Not Process.Position = Process.ExpressionLength Then
                If m_strChar = m_gl_CST_strOprMinus Then
                    m_booIsNegative = True
                ElseIf m_strChar = m_gl_CST_strOprAnd Then
                    m_booIsHex = True
                End If
            ElseIf Not m_chiInfo.IsNumeric Then
                GoTo Failure
            ElseIf m_lngPosition = Process.ExpressionLength Then
                Process.CharIndex = Process.CharIndex + 1
                GoTo Success
            End If
        Else
            If Not m_booIsHex Then
                If m_booIsNegative Then
                    If Not m_chiInfo.IsNumeric Then
                        If Process.CharIndex = 2 Then
                            m_booIsNegative = False
                            GoTo Failure
                        Else
                            If m_strChar = vbCr Or m_strChar = vbLf Or m_strChar = m_gl_CST_strOprSpace Or IsOperator(m_strChar, m_colOperators) Or (m_lngPosition >= Process.ExpressionLength) Then
                                If m_strChar = m_gl_CST_strOprAnd Or m_strChar = m_gl_CST_strOprPercent Then
                                    Process.CharIndex = Process.CharIndex + 1
                                    m_booAndLast = True
                                End If
                                GoTo Success
                            Else
                                GoTo Failure
                            End If
                        End If
                    Else
                        If m_lngPosition = Process.ExpressionLength Then
                            Process.CharIndex = Process.CharIndex + 1
                            GoTo Success
                        End If
                    End If
                Else
                    If Not m_chiInfo.IsNumeric Then
                        If m_strChar = vbCr Or m_strChar = vbLf Or m_strChar = m_gl_CST_strOprSpace Or IsOperator(m_strChar, m_colOperators) Then
                            If m_strChar = m_gl_CST_strOprAnd Or m_strChar = m_gl_CST_strOprPercent Then
                                Process.CharIndex = Process.CharIndex + 1
                                m_booAndLast = True
                            End If
                            GoTo Success
                        Else
                            GoTo Failure
                        End If
                    End If
                End If
            Else
                If Process.CharIndex = 2 Then
                    If Not LCase$(m_strChar) = "h" Then
                        GoTo Failure
                    End If
                ElseIf Process.CharIndex = 3 Then
                    Select Case LCase$(m_strChar)
                        Case "0" To "9", "a" To "f"
                        Case Else
                            GoTo Failure
                    End Select
                Else
                    Select Case LCase$(m_strChar)
                        Case "0" To "9", "a" To "f"
                            If m_lngPosition = Process.ExpressionLength Then
                                Process.CharIndex = Process.CharIndex + 1
                                GoTo Success
                            End If
                        Case m_gl_CST_strOprAnd
                            m_booAndLast = True
                            Process.CharIndex = Process.CharIndex + 1
                            GoTo Success
                        Case Else
                            If m_strChar = vbCr Or m_strChar = vbLf Or m_strChar = m_gl_CST_strOprSpace Or IsOperator(m_strChar, m_colOperators) Then
                                GoTo Success
                            Else
                                GoTo Failure
                            End If
                    End Select
                End If
            End If
        End If
        Process.CharIndex = Process.CharIndex + 1
        m_lngPosition = m_lngPosition + 1
    Loop Until m_lngPosition > Process.ExpressionLength
Finally:
    GoTo Success
Success:
    With m_lprResult
        .Success = True
        .NewPosition = Process.Position + Process.CharIndex - 1
        With .Token
            .Length = Process.CharIndex - 1
            .Value = CLng(Mid$(Process.Expression, Process.Position, Process.CharIndex - 1 - IIf((m_booAndLast), 1, 0)))
            .TokenType = G_LTT_Constant
            .Position = Process.Position
        End With
    End With
    LexicalConstant = m_lprResult
    GoTo EndTry
Failure:
    m_lprResult.Success = False
    GoTo EndTry
Catch:
    m_lprResult.Success = False
    GoTo EndTry
EndTry:
    Process.CharIndex = 1
    Exit Function
End Function

Private Function LexicalWhiteSpace(Process As LexicalProcess) As LexicalProcResult
Declarations:
    Dim m_lngPosition As Long
    Dim m_strChar As String * 1
    Dim m_lprResult As LexicalProcResult

    Dim m_chiInfo As CharInfo
Try:
    On Error GoTo Catch
    m_lngPosition = Process.Position
    Do
        m_strChar = Mid$(Process.Expression, m_lngPosition, 1)
        m_chiInfo = GetCharInfo(m_strChar)
        Select Case m_strChar
            Case m_gl_CST_strOprSpace, vbTab, vbCr, vbLf
                If m_lngPosition = Process.ExpressionLength Then
                    Process.CharIndex = Process.CharIndex + 1
                    GoTo Success
                End If
            Case Else
                If Process.CharIndex > 1 Then
                    Process.CharIndex = Process.CharIndex - 1
                    GoTo Success
                Else
                    GoTo Failure
                End If
        End Select
        Process.CharIndex = Process.CharIndex + 1
        m_lngPosition = m_lngPosition + 1
    Loop Until m_lngPosition > Process.ExpressionLength
Finally:
    GoTo Success
Success:
    With m_lprResult
        .Success = True
        .NewPosition = Process.Position + Process.CharIndex
        With .Token
            .Value = Mid$(Process.Expression, Process.Position, Process.CharIndex)
            .TokenType = G_LTT_WhiteSpace
            .Length = Process.CharIndex
            .Position = Process.Position
        End With
    End With
    LexicalWhiteSpace = m_lprResult
    GoTo EndTry
Failure:
    m_lprResult.Success = False
    GoTo EndTry
Catch:
    m_lprResult.Success = False
    GoTo EndTry
EndTry:
    Process.CharIndex = 1
    Exit Function
End Function

Private Function LexicalKeyWord(Process As LexicalProcess) As LexicalProcResult
Declarations:
    Dim m_lngPosition As Long
    Dim m_strChar As String * 1
    Dim m_lprResult As LexicalProcResult
    Dim m_chiInfo As CharInfo
Try:
    On Error GoTo Catch
    m_lngPosition = Process.Position
    Do
        m_strChar = Mid$(Process.Expression, m_lngPosition, 1)
        m_chiInfo = GetCharInfo(m_strChar)
        If Process.CharIndex = 1 Then
            If (m_chiInfo.IsNumeric Or Not m_chiInfo.IsAlpha) And Not IsOperator(m_strChar, m_colOperators) Then
                GoTo Failure
            End If
        Else
            Select Case True
                Case m_strChar = m_gl_CST_strOprUnderscore
                    If m_lngPosition = Process.ExpressionLength Then
                        Process.CharIndex = Process.CharIndex + 1
                        GoTo Success
                    End If
                Case IsOperator(m_strChar, m_colOperators), m_strChar = m_gl_CST_strOprSpace, m_strChar = vbCr, m_strChar = vbLf, m_strChar = vbTab
                    GoTo Success
                Case Else
                    If m_lngPosition = Process.ExpressionLength Then
                        Process.CharIndex = Process.CharIndex + 1
                        GoTo Success
                    End If
            End Select
        End If
        Process.CharIndex = Process.CharIndex + 1
        m_lngPosition = m_lngPosition + 1
    Loop Until m_lngPosition > Process.ExpressionLength
Finally:
    GoTo EndTry
Success:
    With m_lprResult
        .Success = True
        .NewPosition = Process.Position + Process.CharIndex - 1
        With .Token
            .Value = Mid$(Process.Expression, Process.Position, Process.CharIndex - 1)
            If Not IsKeyword(.Value, m_colKeyWords) Then
                GoTo Failure
            End If
            .Length = Len(.Value)
            .TokenType = G_LTT_Keyword
            .Position = Process.Position
        End With
    End With
    LexicalKeyWord = m_lprResult
    GoTo EndTry
Failure:
    m_lprResult.Success = False
    GoTo EndTry
Catch:
    m_lprResult.Success = False
    GoTo EndTry
EndTry:
    Process.CharIndex = 1
    Exit Function
End Function

Private Function LexicalOperator(Process As LexicalProcess) As LexicalProcResult
Declarations:
    Dim m_lngPosition As Long
    Dim m_strChar As String * 1
    Dim m_lprResult As LexicalProcResult
    Dim m_strExp As String
    Dim m_chiInfo As CharInfo
    Dim m_lngSuccessChar As Long
Try:
    On Error GoTo Catch
    m_lngPosition = Process.Position
    Do
        m_strChar = Mid$(Process.Expression, m_lngPosition, 1)
        m_chiInfo = GetCharInfo(m_strChar)
        m_strExp = m_strExp & m_strChar
        If Process.CharIndex = 1 Then
            If IsOperator(m_strChar, m_colOperators) Then
                m_lngSuccessChar = Process.CharIndex
                If Process.Position = Process.ExpressionLength Then
                    Process.CharIndex = Process.CharIndex + 1
                    GoTo Success
                End If
            End If
        Else
            If Not IsOperator(m_strExp, m_colOperators) Then
                GoTo Success
            Else
                m_lngSuccessChar = Process.CharIndex + 1
                If m_lngPosition = Process.ExpressionLength Then
                    Process.CharIndex = Process.CharIndex + 1
                    GoTo Success
                End If
            End If
        End If
        Process.CharIndex = Process.CharIndex + 1
        m_lngPosition = m_lngPosition + 1
    Loop Until m_lngPosition > Process.ExpressionLength
Finally:
    GoTo EndTry
Success:
    With m_lprResult
        .Success = True
        .NewPosition = Process.Position + Process.CharIndex - 1
        With .Token
            .Length = Process.CharIndex - 1
            .Position = Process.Position
            .TokenType = G_LTT_Operator
            .Value = Mid$(Process.Expression, Process.Position, Process.CharIndex - 1)
            If Not IsOperator(.Value, m_colOperators) Then
                If m_lngSuccessChar = 0 Then
                    GoTo Failure
                Else
                    Process.CharIndex = m_lngSuccessChar
                    GoTo Success
                        '//update
                End If
            End If
        End With
    End With
    LexicalOperator = m_lprResult
    GoTo EndTry
Failure:
    m_lprResult.Success = False
    GoTo EndTry
Catch:
    m_lprResult.Success = False
    GoTo EndTry
EndTry:
    Process.CharIndex = 1
    Exit Function
End Function

Private Function LexicalIdentifier(Process As LexicalProcess) As LexicalProcResult
Declarations:
    Dim m_lngPosition As Long
    Dim m_strChar As String * 1
    Dim m_lprResult As LexicalProcResult
    Dim m_chiInfo As CharInfo
Try:
    On Error GoTo Catch
    m_lngPosition = Process.Position
    Do
        m_strChar = Mid$(Process.Expression, m_lngPosition, 1)
        m_chiInfo = GetCharInfo(m_strChar)
        If Process.CharIndex = 1 Then
            If m_chiInfo.IsNumeric Or Not m_chiInfo.IsAlpha Then
                GoTo Failure
            ElseIf m_lngPosition = Process.ExpressionLength Then
                Process.CharIndex = Process.CharIndex + 1
                GoTo Success
            End If
        Else
            Select Case True
                Case m_strChar = m_gl_CST_strOprUnderscore
                    If m_lngPosition = Process.ExpressionLength Then
                        Process.CharIndex = Process.CharIndex + 1
                        GoTo Success
                    End If
                Case IsOperator(m_strChar, m_colOperators), m_strChar = m_gl_CST_strOprSpace, m_strChar = vbCr, m_strChar = vbLf, m_strChar = vbTab
                    GoTo Success
                Case Else
                    If m_lngPosition = Process.ExpressionLength Then
                        Process.CharIndex = Process.CharIndex + 1
                        GoTo Success
                    End If
            End Select
        End If
        Process.CharIndex = Process.CharIndex + 1
        m_lngPosition = m_lngPosition + 1
    Loop Until m_lngPosition > Process.ExpressionLength
Finally:
    GoTo EndTry
Success:
    With m_lprResult
        .Success = True
        .NewPosition = Process.Position + Process.CharIndex - 1
        With .Token
            .Value = Mid$(Process.Expression, Process.Position, Process.CharIndex - 1)
            .Length = Len(.Value)
            .TokenType = G_LTT_Identifier
            .Position = Process.Position
        End With
    End With
    LexicalIdentifier = m_lprResult
    GoTo EndTry
Failure:
    m_lprResult.Success = False
    GoTo EndTry
Catch:
    m_lprResult.Success = False
    GoTo EndTry
EndTry:
    Process.CharIndex = 1
    Exit Function
End Function

Private Sub AddToken(Token As LexicalToken, Tokens As LexicalTokens)
    With Tokens
        If .Count = 0 Then
            ReDim .Tokens(1 To .Count + 1)
        Else
            ReDim Preserve .Tokens(1 To .Count + 1)
        End If
        .Count = .Count + 1
        With .Tokens(.Count)
            .CustID = Token.CustID
            .Length = Token.Length
            .Position = Token.Position
            .TokenType = Token.TokenType
            .Value = Token.Value
        End With
    End With
End Sub

Public Sub FixLanguage()
Attribute FixLanguage.VB_Description = "Stores Operators and Keywrods into collections"
    Dim m_lngLoop As Long
    Set m_colKeyWords = New Collection
    Set m_colOperators = New Collection
    For m_lngLoop = 1 To Language.KeywordList.Count
        m_colKeyWords.Add Language.KeywordList.Keywords(m_lngLoop).StringValue, "k" & LCase$(Language.KeywordList.Keywords(m_lngLoop).StringValue)
    Next
    For m_lngLoop = 1 To Language.OperatorList.Count
        m_colOperators.Add Language.OperatorList.Operators(m_lngLoop).Value, "o" & LCase$(Language.OperatorList.Operators(m_lngLoop).Value)
    Next
    'Debug.Print m_colKeyWords.Count; m_colOperators.Count
End Sub

Public Property Get OperatorCol() As Collection
    Set OperatorCol = m_colOperators
End Property
