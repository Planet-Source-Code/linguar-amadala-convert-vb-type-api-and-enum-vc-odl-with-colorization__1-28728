VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ODLanguage"
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
            .AddKeyWord "", "appobject"
            .AddKeyWord "", "control"
            .AddKeyWord "", "defaultvalue"
            .AddKeyWord "", "dllname"
            .AddKeyWord "", "entry"
            .AddKeyWord "", "helpstring"
            .AddKeyWord "", "hidden"
            .AddKeyWord "", "immediatebind"
            .AddKeyWord "", "lcid"
            .AddKeyWord "", "nonextensible"
            .AddKeyWord "", "oleautomation"
            .AddKeyWord "", "propget"
            .AddKeyWord "", "propputref"
            .AddKeyWord "", "readonly"
            .AddKeyWord "", "restricted"
            .AddKeyWord "", "source"
            .AddKeyWord "", "vararg"
            .AddKeyWord "", "out"
            .AddKeyWord "", "in"
            .AddKeyWord "", "bindable"
            .AddKeyWord "", "default"
            .AddKeyWord "", "displaybind"
            .AddKeyWord "", "dual"
            .AddKeyWord "", "helpcontext"
            .AddKeyWord "", "helpfile"
            .AddKeyWord "", "id"
            .AddKeyWord "", "licensed"
            .AddKeyWord "", "optional"
            .AddKeyWord "", "propput"
            .AddKeyWord "", "public"
            .AddKeyWord "", "retval"
            .AddKeyWord "", "requestedit"
            .AddKeyWord "", "uuid"
            .AddKeyWord "", "version"
            .AddKeyWord "", "unsigned"
            .AddKeyWord "", "long"
            .AddKeyWord "", "int"
            .AddKeyWord "", "BSTR"
            .AddKeyWord "", "UINT"
            .AddKeyWord "", "typedef"
            .AddKeyWord "", "LPSTR"
            .AddKeyWord "", "double"
            .AddKeyWord "", "float"
            .AddKeyWord "", "char"
            .AddKeyWord "", "struct"
            .AddKeyWord "", "enum"
            .AddKeyWord "", "union"
            .AddKeyWord "", "module"
            .AddKeyWord "", "boolean"
            .AddKeyWord "", "coclass"
            .AddKeyWord "", "importlib"
            .AddOperator "{", "{"
            .AddOperator "}", "}"
            .AddOperator "\", "\"
            .AddOperator "@", "@"
            .AddOperator "+", "+"
            .AddOperator "-", "-"
            .AddOperator "->*", "->*"
            .AddOperator ".*", ".*"
            .AddOperator ".", "."
            .AddOperator "->", "->"
            .AddOperator "==", "=="
            .AddOperator "=", "="
            .AddOperator ">", ">"
            .AddOperator "<", "<"
            .AddOperator "/", "/"
            .AddOperator "*", "*"
            .AddOperator "(", "("
            .AddOperator ")", ")"
            .AddOperator "[", "["
            .AddOperator "]", "]"
            .AddOperator ",", ","
            .AddOperator ";", ";"
            .AddOperator ":", ":"
            .AddOperator "&", "&"
            .AddOperator "|", "|"
            .AddOperator "#", "#"
            .AddOperator "!", "!"
            .AddOperator "^", "^"
        End With
    End If
    FixLanguage
End Property

Public Property Get Language() As SCLanguage
    Set Language = r_gslLanguage
End Property

Public Property Set Language(ByVal v_gslLanguage As SCLanguage)
    Set r_gslLanguage = v_gslLanguage
End Property

Friend Function LexicalParse(ByVal Expression As String) As LexicalResult
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
                m_lprResult = LexicalComment(m_glpProcess)
            If Not .Success Then _
                m_lprResult = LexicalIdentifier(m_glpProcess)
            If Not .Success Then _
                m_lprResult = LexicalOperator(m_glpProcess)
            If Not .Success Then _
                m_lprResult = LexicalString(m_glpProcess)
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
    Loop Until m_lngPosition > Len(Expression)
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
    Dim m_booFoundSlash As Boolean
    Dim m_chiInfo As CharInfo
Try:
    On Error GoTo Catch
    m_lngPosition = Process.Position
    Do
        m_strChar = Mid(Process.Expression, m_lngPosition, 1)
        m_chiInfo = GetCharInfo(m_strChar)
        If Process.CharIndex = 1 Then
            If Not m_strChar = m_gl_CST_strOprQuote Then
                GoTo Failure
            End If
        Else
            If m_booFoundSlash Then
                m_booFoundSlash = False
            ElseIf m_strChar = m_gl_CST_strOprBackSlash Then
                m_booFoundSlash = True
                If m_lngPosition = Process.ExpressionLength Then
                    GoTo Failure
                End If
            ElseIf m_strChar = m_gl_CST_strOprQuote Then
                Process.CharIndex = Process.CharIndex
                GoTo Success
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
            .Value = Mid(Process.Expression, Process.Position, Process.CharIndex)
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
    Dim m_lngCommentType As Long
    Dim m_chiInfo As CharInfo
    Dim m_lngNextItem As Long
    Dim m_lngNextUnderItem As Long
    Dim m_lngCRPos As Long
Try:
    On Error GoTo Catch
    m_lngPosition = Process.Position
    m_strChar = Mid(Process.Expression, m_lngPosition, 1)
    m_chiInfo = GetCharInfo(m_strChar)
    If Process.CharIndex = 1 Then
        If m_chiInfo.IsAlpha Or m_chiInfo.IsNumeric Then
            GoTo Failure
        ElseIf m_strChar = m_gl_CST_strOprForwardSlash Then
            If Mid(Process.Expression, m_lngPosition + 1, 1) = m_gl_CST_strOprForwardSlash Then
                m_lngCommentType = 1
                GoTo Success
            ElseIf Mid(Process.Expression, m_lngPosition + 1, 1) = m_gl_CST_strOprTimes Then
                m_lngCommentType = 2
                GoTo Success
            End If
        End If
    End If
    Process.CharIndex = Process.CharIndex + 1
Finally:
    GoTo EndTry
Success:
    With m_lprResult
        If m_lngCommentType = 1 Then
            m_lngNextItem = Process.Position
            m_lngNextUnderItem = m_lngNextItem
            m_lngCRPos = m_lngNextItem
            Do
                m_lngNextItem = NextCarrageReturn(Process.Expression, m_lngCRPos)
                m_lngNextUnderItem = InStr(m_lngCRPos, Process.Expression, " \" & vbCr)
                m_lngCRPos = m_lngNextItem + 1
            Loop Until m_lngNextItem = 0 Or (m_lngNextUnderItem > m_lngNextItem) Or m_lngNextUnderItem = 0
            If m_lngNextItem = 0 Then
                Process.CharIndex = (Process.ExpressionLength - (Process.Position - 1))
            Else
                Process.CharIndex = ((m_lngNextItem - 1) - (Process.Position - 1))
            End If
        Else
            Process.CharIndex = InStr(Process.Position, Process.Expression, "*/") - (Process.Position - 1) + 1
            If Process.CharIndex <= 0 Then
                Process.CharIndex = Process.ExpressionLength - (Process.Position - 1)
            End If
        End If
        .Success = True
        .NewPosition = Process.Position + Process.CharIndex
        With .Token
            .Value = Mid(Process.Expression, Process.Position, Process.CharIndex)
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
    Dim m_booIsHex As Boolean
    Dim m_chiInfo As CharInfo
    Dim m_booIsNegative As Boolean
Try:
    On Error GoTo Catch
    m_lngPosition = Process.Position
    Do
        m_strChar = Mid(Process.Expression, m_lngPosition, 1)
        m_chiInfo = GetCharInfo(m_strChar)
        If Process.CharIndex = 1 Then
            If m_strChar = m_gl_CST_strOprMinus And Not Process.Position = Process.ExpressionLength Then
                m_booIsNegative = True
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
                        If Not Process.CharIndex = 2 And m_strChar = "x" Then
                            If m_strChar = vbCr Or m_strChar = vbLf Or m_strChar = m_gl_CST_strOprSpace Or IsOperator(m_strChar, m_colOperators) Then
                                GoTo Success
                            Else
                                GoTo Failure
                            End If
                        Else
                            m_booIsHex = True
                        End If
                    End If
                End If
            Else
                Select Case m_strChar
                    Case "0" To "9", "a" To "f", "A" To "F"
                        
                    Case Else
                        If Process.CharIndex = 3 Then
                            '//If it's suspected of being a hexadecimal value,
                            '//and the first item after the 'x' isn't a hexa-
                            '//decimal value or a decimal value, thne it's
                            '//clearly not valid.
                            GoTo Failure
                        End If
                        If m_strChar = vbCr Or m_strChar = vbLf Or m_strChar = m_gl_CST_strOprSpace Or IsOperator(m_strChar, m_colOperators) Then
                            GoTo Success
                        Else
                            GoTo Failure
                        End If
                End Select
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
        If Not m_booIsHex Then
            With .Token
                .Length = Process.CharIndex - 1
                .Value = CLng(Mid(Process.Expression, Process.Position, Process.CharIndex - 1))
                .TokenType = G_LTT_Constant
                .Position = Process.Position
            End With
        Else
            With .Token
                .Length = Process.CharIndex - 1
                .Value = (Mid(Process.Expression, Process.Position, Process.CharIndex - 1))
                .Value = CLng("&H" & Mid(.Value, 3))
                .TokenType = G_LTT_Constant
                .Position = Process.Position
                .CustID = True
                    '//Internal... signifies that it was a number in 0x0N
                    '//form
            End With
        End If
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
        m_strChar = Mid(Process.Expression, m_lngPosition, 1)
        m_chiInfo = GetCharInfo(m_strChar)
        Select Case m_strChar
            Case m_gl_CST_strOprSpace, vbTab, vbCr, vbLf
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
            .Value = Mid(Process.Expression, Process.Position, Process.CharIndex)
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
        m_strChar = Mid(Process.Expression, m_lngPosition, 1)
        m_chiInfo = GetCharInfo(m_strChar)
        If Process.CharIndex = 1 Then
            If (m_chiInfo.IsNumeric) Then 'And Not IsOperator(m_strChar, m_colOperators) Then
                GoTo Failure
            End If
        Else
            Select Case True
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
            .Value = Mid(Process.Expression, Process.Position, Process.CharIndex - 1)
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
    Dim m_lngSuccessChar As Long
    Dim m_chiInfo As CharInfo
    Dim m_strExp As String
Try:
    On Error GoTo Catch
    m_lngPosition = Process.Position
    Do
        m_strChar = Mid(Process.Expression, m_lngPosition, 1)
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
            .Value = Mid(Process.Expression, Process.Position, Process.CharIndex - 1)
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
        m_strChar = Mid(Process.Expression, m_lngPosition, 1)
        m_chiInfo = GetCharInfo(m_strChar)
        If Process.CharIndex = 1 Then
            If m_chiInfo.IsNumeric Or Not (m_chiInfo.IsAlpha Or m_chiInfo.IsUnderscore) Then
                GoTo Failure
            ElseIf Process.Position = Process.ExpressionLength Then
                Process.CharIndex = Process.CharIndex + 1
                GoTo Success
            End If
        Else
            Select Case True
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
            .Value = Mid(Process.Expression, Process.Position, Process.CharIndex - 1)
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
    Dim m_lngLoop As Long
    Set m_colKeyWords = New Collection
    Set m_colOperators = New Collection
    For m_lngLoop = 1 To Language.KeywordList.Count
        m_colKeyWords.Add Language.KeywordList.Keywords(m_lngLoop).StringValue, "k" & LCase(Language.KeywordList.Keywords(m_lngLoop).StringValue)
    Next
    For m_lngLoop = 1 To Language.OperatorList.Count
        m_colOperators.Add Language.OperatorList.Operators(m_lngLoop).Value, "o" & LCase(Language.OperatorList.Operators(m_lngLoop).Value)
    Next
    'Debug.Print m_colKeyWords.Count; m_colOperators.Count
End Sub

Public Property Get OperatorCol() As Collection
    Set OperatorCol = m_colOperators
End Property
