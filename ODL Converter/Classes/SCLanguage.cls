VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SCLanguage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private r_strName As String
Private r_sleEngine As SCLangEngine
Private r_gosOperatorList As Operators
Private r_gksKeywordList As Keywords
Private r_clsCollection As SCLanguages

Public Property Get Collection() As SCLanguages
    Set Collection = r_clsCollection
End Property

Public Property Set Collection(ByVal v_clsCollection As SCLanguages)
    Set r_clsCollection = v_clsCollection
End Property

Friend Property Get OperatorList() As Operators
    OperatorList = r_gosOperatorList
End Property

Friend Property Let OperatorList(v_gosOperatorList As Operators)
    r_gosOperatorList = v_gosOperatorList
End Property

Public Property Get Name() As String
    Name = r_strName
End Property

Public Property Let Name(ByVal v_strName As String)
    r_strName = v_strName
End Property

Public Property Get Engine() As SCLangEngine
    Set Engine = r_sleEngine
End Property

Public Property Set Engine(ByVal v_sleSCLangEngine As SCLangEngine)
    If r_sleEngine Is v_sleSCLangEngine Then _
        Exit Property
    If Not r_sleEngine Is Nothing Then
        If r_sleEngine.Language Is Me Then
            Set r_sleEngine = Nothing
        End If
    End If
    Set r_sleEngine = v_sleSCLangEngine
    If Not v_sleSCLangEngine Is Nothing Then
        Set v_sleSCLangEngine.Language = Me
    End If
End Property

Private Sub Class_Initialize()
    Set r_sleEngine = New SCLangEngine
    Set r_sleEngine.Language = Me
End Sub

Friend Property Get KeywordList() As Keywords
    KeywordList = r_gksKeywordList
End Property

Friend Property Let KeywordList(v_gksKeywordList As Keywords)
    r_gksKeywordList = v_gksKeywordList
End Property

Public Sub AddKeyWord(ByVal Description As String, ByVal TokenValue As String, Optional ByVal Flags As KeywordFlags = G_KF_Flexable)
    With r_gksKeywordList
        If .Count = 0 Then
            ReDim .Keywords(1 To .Count + 1)
        Else
            ReDim Preserve .Keywords(1 To .Count + 1)
        End If
        .Count = .Count + 1
        With .Keywords(.Count)
            .Description = Description
            .StringValue = TokenValue
            .Flags = Flags
        End With
    End With
End Sub

Public Sub AddOperator(ByVal Description As String, ByVal TokenValue As String, Optional Priority As Long = 1)
    With r_gosOperatorList
        If .Count = 0 Then
            ReDim .Operators(1 To .Count + 1)
        Else
            ReDim Preserve .Operators(1 To .Count + 1)
        End If
        .Count = .Count + 1
        With .Operators(.Count)
            .Value = TokenValue
            .Priority = Priority
            .Description = Description
        End With
    End With
End Sub

Public Sub RemoveKeyword(ByVal Index As Long)
    With r_gksKeywordList
        Do Until Index >= .Count - 1
            .Keywords(Index) = .Keywords(Index + 1)
            Index = Index + 1
        Loop
        ReDim .Keywords(1 To .Count)
    End With
End Sub

Friend Sub AddToken(Token As LexicalToken, Tokens As LexicalTokens)
    With Tokens
        If .Count = 0 Then
            ReDim .Tokens(1 To .Count + 1)
        Else
            ReDim Preserve .Tokens(1 To .Count + 1)
        End If
        .Count = .Count + 1
        .Tokens(.Count) = Token
    End With
End Sub
