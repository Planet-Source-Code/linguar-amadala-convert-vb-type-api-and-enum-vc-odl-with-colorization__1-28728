Attribute VB_Name = "modLexProcs"
Option Explicit
Public Type CharInfo
    IsNumeric As Boolean
    IsAlpha As Boolean
    IsUnderscore As Boolean
End Type

Public Function IsOperator(ByVal Expr As String, Col As Collection) As Boolean
Try:
    On Error GoTo Catch
    Call Col.Item("o" & LCase$(Expr))
Finally:
    IsOperator = True
    GoTo EndTry
Catch:
    IsOperator = False
EndTry:
    Exit Function
End Function

Public Function IsKeyword(ByVal Expr As String, Col As Collection) As Boolean
Try:
    On Error GoTo Catch
    Call Col.Item("k" & LCase$(Expr))
Finally:
    IsKeyword = True
    GoTo EndTry
Catch:
    IsKeyword = False
EndTry:
    Exit Function
End Function

Public Function IsAlpha(Char As String)
    Select Case Char
        Case "a" To "z", "A" To "Z"
            IsAlpha = True
    End Select
End Function

Public Function GetCharInfo(Char As String) As CharInfo
    Dim m_chiInfo As CharInfo
    With m_chiInfo
        .IsNumeric = IsNumeric(Char)
        .IsAlpha = IsAlpha(Char)
        .IsUnderscore = CBool(Char = m_gl_CST_strOprUnderscore)
    End With
    GetCharInfo = m_chiInfo
End Function

Public Function NextCarrageReturn(Expression As String, Position As Long) As Long
    NextCarrageReturn = InStr(Position, Expression, vbCr)
End Function

Public Function ToPostfix(Tokens As LexicalTokens, Operators As Operators, ScopeStart As Long, ScopeEnd As Long) As String
    Dim m_strResult As String
    Dim m_lngScopeLevel As Long
    AppendPostfix m_strResult, Operators, Tokens, ScopeStart, ScopeEnd, m_lngScopeLevel
    ToPostfix = m_strResult
End Function

Public Function AppendPostfix(CurrentExpression As String, Operators As Operators, Tokens As LexicalTokens, ScopeStart As Long, ScopeEnd As Long, ByRef ScopeLevel As Long) As Long
    Dim m_lngPosition As Long
    Dim m_gltToken As LexicalToken
    ScopeLevel = ScopeLevel + 1
    m_lngPosition = ScopeStart
    Do
        m_gltToken = Tokens.Tokens(m_lngPosition)
        If m_gltToken.TokenType = G_LTT_Operator Then
            If m_gltToken.Value = m_gl_CST_strOprRightBracket Then
                m_lngPosition = m_lngPosition + 1
                ScopeLevel = ScopeLevel - 1
                Exit Do
            End If
        End If
        m_lngPosition = m_lngPosition + 1
    Loop Until m_lngPosition > ScopeEnd
    AppendPostfix = m_lngPosition
End Function

Public Function ProcessOperatorLevel(Operators As Collection, Level As Long)
    Dim m_lngLoop As Long
    
End Function

Public Function OperatorIsIn(Expression As String, Operators As Collection)
Try:
    On Error GoTo Catch
    Call Operators.Item(Expression)
        '//If it works, it exists...
Finally:
    OperatorIsIn = True
    GoTo EndTry
Catch:
    OperatorIsIn = False
    GoTo EndTry
EndTry:
    Exit Function
End Function

Public Function GetOperatorLists(Operators As Operators) As Collection
    Dim m_lngCount As Long
    Dim m_lngLoop As Long
    Dim m_colCollections As Collection
    Dim m_colCurCol As Collection
    Set m_colCollections = New Collection
    Dim m_gopOperator As Operator
    m_lngCount = GetHighestPriority(Operators)
    For m_lngLoop = 1 To m_lngCount
        Set m_colCurCol = New Collection
        m_colCollections.Add m_colCurCol
    Next
    For m_lngLoop = 1 To Operators.Count
        m_gopOperator = Operators.Operators(m_lngLoop)
        If m_gopOperator.Priority > 0 Then
            Set m_colCurCol = m_colCollections(m_gopOperator.Priority)
            m_colCurCol.Add m_gopOperator.Value, "o" & m_gopOperator.Value
            Set m_colCurCol = Nothing
        End If
    Next
    Set GetOperatorLists = m_colCollections
End Function

Public Function GetHighestPriority(Operators As Operators) As Long
    Dim m_lngLoop As Long
    Dim m_gopOperator As Operator
    Dim m_lngHighest As Long
    For m_lngLoop = 1 To Operators.Count
        m_gopOperator = Operators.Operators(m_lngLoop)
        If m_gopOperator.Priority > m_lngHighest Then
            m_lngHighest = m_gopOperator.Priority
        End If
    Next
    GetHighestPriority = m_lngHighest
End Function

Public Function DeflateString(Expression As String, DeflateBy As Long) As String
    Dim m_lngDeflate As Long
    Dim m_strDeflate As String
    m_lngDeflate = DeflateBy \ 2
    m_strDeflate = Right(Expression, Len(Expression) - m_lngDeflate)
    DeflateString = Left(m_strDeflate, Len(m_strDeflate) - (DeflateBy - m_lngDeflate))
End Function
