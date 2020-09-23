Attribute VB_Name = "modSemProcess"
Option Explicit
Option Compare Text
Public Enum IfStatementConditionType
    ISCT_Operation
    ISCT_Variable
    ISCT_Constant
End Enum
Public Enum IfStatementOperatorType
    ISOT_Add
    ISOT_And
    ISOT_Divide
    ISOT_Equals
    ISOT_GreaterThan '//>
    ISOT_GreaterThanOrEqualTo '//>=
    ISOT_LessThan '//<
    ISOT_LessThanOrEqualTo '//<=
    ISOT_Minus
    ISOT_Or
    ISOT_ToThePower
End Enum
Public Enum SectType
    ST_Declare
    ST_DeclareArgument
    ST_Enum
    ST_EnumItem
    ST_Struct
    ST_StructItem
    ST_NumberConstant
    ST_OtherConstant
    ST_Sub
    ST_Function
End Enum
Public Enum NameSpaceItemScope
    NSIS_Public = 0
    NSIS_Private = 1
    NSIS_Friend = 2
    NSIS_Static = 4
End Enum
Public Type NameSpaceItem
    Name As String
    Scope As Long
End Type
Public Type NameSpace
    Count As Long
    Names() As NameSpaceItem
End Type
Public Type CodeSection
    SectType As Long
    ptrData As Long
    Position As Long
End Type
Public Type CodeSections
    Count As Long
    Sections() As CodeSection
End Type
Public Type ResultItem
    Name As String
    Type As String
    bData As Boolean
    lData As Long
    Section As Long
    ByVal As Boolean
End Type
Public Type ResultItems
    Count As Long
    Items() As ResultItem
End Type
Public Type DeclareStatement
    Header As ResultItem
    Alias As String
    Library As String
    Arguments As ResultItems
    Section As Long
End Type
Public Type DeclareStatements
    Count As Long
    Statements() As DeclareStatement
End Type
Public Type StructStatement
    Name As String
    Members As ResultItems
    Section As Long
End Type
Public Type ConstantStatement
    Name As String
    Value As Variant
    Section As Long
End Type
Public Type EnumStatementItem
    Name As String
    Value As Long
    Section As Long
End Type
Public Type EnumStatementItems
    Count As Long
    Item() As EnumStatementItem
End Type
Public Type EnumStatement
    Name As String
    Section As Long
    Members As EnumStatementItems
End Type
Public Type IfStatementCondition
    ConditionType As Long
    Data1 As String
        '//String Value
    Data2 As Long
        '//Constant Value/Operator Value
End Type
Public Type IfStatementConditions
    Count As Long
    Conditions() As IfStatementCondition
End Type
Public Type IfStatement
    Conditions As IfStatementConditions
    NextSection As Long
End Type
Public Type IfStatements
    Count As Long
    Ifs() As IfStatement
End Type
Public Type FunctionStatement
    Header As ResultItem
    IfStatements As IfStatements
End Type
Public Type EnumStatements
    Count As Long
    Enums() As EnumStatement
End Type
Public Type StructStatements
    Count As Long
    Structs() As StructStatement
End Type
Public Type CodeResult
    Sections As CodeSections
    Declares As DeclareStatements
    Structs As StructStatements
    Enums As EnumStatements
    GlobalNamespace As NameSpace
End Type

Public Function SemanticParse(Code As String, Engine As BasicLanguage) As CodeResult
    Dim m_glrRes As LexicalResult
    Dim m_gltToken As LexicalToken
    Dim m_lngLoop As Long
    Dim m_booIgnoreLine As Boolean
    Dim m_corRes As CodeResult
    Dim m_lngItem As Long
    Dim m_strScope As String
    Dim m_lngScope As Long
    m_glrRes = Engine.LexicalParse(Code)
    For m_lngLoop = 1 To m_glrRes.Tokens.Count
        m_gltToken = m_glrRes.Tokens.Tokens(m_lngLoop)
        If Not m_booIgnoreLine Then
            Select Case m_gltToken.TokenType
                Case G_LTT_Keyword
                    Select Case LCase$(m_gltToken.Value)
                        Case m_gl_CST_strKwdDeclare
                            GetDeclare m_glrRes, m_corRes, m_lngLoop, m_lngScope
                            m_lngScope = 0
                        Case m_gl_CST_strKwdPublic
                            m_lngScope = m_lngLoop
                        Case m_gl_CST_strKwdPrivate
                            m_lngScope = m_lngLoop
                        Case m_gl_CST_strKwdType
                            GetType m_glrRes, m_corRes, m_lngLoop, m_lngScope
                            m_lngScope = 0
                        Case m_gl_CST_strKwdEnum
                            GetEnum m_glrRes, m_corRes, m_lngLoop, m_lngScope
                            m_lngScope = 0
                        Case m_gl_CST_strKwdConst
                            GetConst m_glrRes, m_corRes, m_lngLoop, m_lngScope
                            m_lngScope = 0
                        Case Else
                            m_booIgnoreLine = True
                    End Select
                Case G_LTT_WhiteSpace
                    
                Case G_LTT_Constant
                    
                Case Else
                    m_booIgnoreLine = True
            End Select
        Else
            Select Case m_gltToken.TokenType
                Case G_LTT_WhiteSpace
                    If InStr(1, m_gltToken.Value, vbCr, vbTextCompare) <> 0 Then
                        m_booIgnoreLine = False
                        m_lngScope = 0
                    End If
            End Select
        End If
    Next
    SemanticParse = m_corRes
End Function

Public Sub AddSection(Sections As CodeSections, SectType As SectType, Pointer As Long, Position As Long)
    With Sections
        If .Count = 0 Then
            ReDim .Sections(1 To .Count + 1)
        Else
            ReDim Preserve .Sections(1 To .Count + 1)
        End If
        .Count = .Count + 1
        With .Sections(.Count)
            .SectType = SectType
            .ptrData = Pointer
            .Position = Position
        End With
    End With
End Sub

Public Sub GetConst(LexResult As LexicalResult, CodeResult As CodeResult, Index As Long, ScopeTok As Long)
    Dim m_lngLoop As Long
        '//The loop variable
    Dim m_gltTok As LexicalToken
        '//The token structure
    Dim m_booNameEnc As Boolean
        '//If the name has been encountered...
    Dim m_booAsEnc As Boolean
        '//If the 'as' has been encountered...
    Dim m_booTypeEnc As Boolean
        '//If the type has been encountered...
    Dim m_booEqualsEnc As Boolean
        '//If we've hit the equals sign...
    Dim m_booValueEnc As Boolean
        '//The flag of encounter
    Dim m_varValue As Variant
        '//The constant value
    Dim m_strAsType As String
        '//The type of the constant
    Dim m_lngPosition As Long
        '//
    Dim m_cosConst As ConstantStatement
        '//Constant variable.
    If ScopeTok = 0 Then
        m_lngPosition = LexResult.Tokens.Tokens(Index).Position
    Else
        m_lngPosition = LexResult.Tokens.Tokens(ScopeTok).Position
    End If
    With LexResult.Tokens
        If .Tokens(Index + 1).TokenType = G_LTT_WhiteSpace And .Tokens(Index + 2).TokenType = G_LTT_Identifier Then
            m_cosConst.Name = .Tokens(Index + 2).Value
                '//Store the name
            m_booNameEnc = True
        Else
            Exit Sub
        End If
    End With
    For m_lngLoop = Index + 3 To LexResult.Tokens.Count
        '//Loop through the tokens, skip the whitespace after the 'const' statement
        m_gltTok = LexResult.Tokens.Tokens(m_lngLoop)
            '//Return the current token
        Select Case m_gltTok.TokenType
            '//Process based upon the token's type
            Case G_LTT_Comment
                '//If it's a comment, then...
                If m_booValueEnc Then
                    Exit For
                End If
            Case G_LTT_Constant
                '//If it's a constant, then...
                If m_booEqualsEnc Then
                    '//If we've encountered the equals sign, then...
                    If Not m_booValueEnc Then
                        '//If the value has not been encountered, then...
                        m_booValueEnc = True
                            '//We've encountered it.
                        m_varValue = m_gltTok.Value
                            '//Set the value
                    Else
                        '//... otherwise...
                        Exit For
                            '//We should NOT be encountering another value!
                            '//Exit the procedure
                    End If
                Else
                    Exit For
                        '//Fail
                End If
            Case G_LTT_String
                '//If it's a string, then...
                If m_booEqualsEnc Then
                    '//If we've encountered the equals, then...
                    If Not m_booValueEnc Then
                        '//If the value has not been encountered, then...
                        m_booValueEnc = True
                            '//We've encountered it.
                        m_varValue = m_gltTok.Value
                            '//Set the value
                    Else
                        '//... otherwise...
                        Exit For
                            '//We should NOT be encountering another value!
                            '//Exit the procedure.
                    End If
                Else
                    '//... otherwise...
                    Exit For
                        '//Fail
                End If
            Case G_LTT_Keyword
                '//If it's a keyword, then...
                If m_booNameEnc Then
                    If AsType(m_gltTok) = m_gltTok.Value Then
                        
                    Else
                        Exit Sub
                            '//This checks one line, therefore an error is fatal.
                    End If
                End If
            Case G_LTT_Identifier
                '//If it's an identifier, then...
                
            Case G_LTT_Operator
                '//If it's an operator, then...
                If Not m_gltTok.Value = m_gl_CST_strOprUnderscore Then
                    If m_booNameEnc Then
                        If m_booAsEnc Then
                            If m_booTypeEnc Then
                                If m_gltTok.Value = m_gl_CST_strOprEquals Then
                                    
                                End If
                            Else
                                Exit Sub
                            End If
                        Else
                            If m_gltTok.Value = m_gl_CST_strOprEquals Then
                                
                            End If
                        End If
                    End If
                Else
                    
                End If
            Case G_LTT_WhiteSpace
                '//If it's whitespace characters, then...
                If InStr(1, m_gltTok.Value, vbCr) <> 0 Then
                    '//If there is a carrage return in the text specified, then...
                    If m_booValueEnc Then
                        '//If the value has been encountered, then...
                        Exit For
                            '//We're done.
                    Else
                        '//... otherwise...
                    End If
                End If
        End Select
    Next '[m_gltTok (m_lngLoop)]
End Sub
    
Public Sub GetType(LexResult As LexicalResult, CodeResult As CodeResult, Index As Long, ScopeTok As Long)
    Dim m_lngLoop As Long
        '//Loop variable
    Dim m_gltTok As LexicalToken
        '//Token Variable
    Dim m_stsStatement As StructStatement
        '//Struct Statement variable
    Dim m_booEndEnc As Boolean
        '//End encountered flag
    Dim m_lngPosition As Long
        '//Struct starting position
    Dim m_booString As Boolean
        '//If it's a string
    Dim m_booIDEnc As Boolean
        '//ID Encountered flag
    Dim m_booAsEnc As Boolean
        '//As encoutnered flag
    Dim m_booConstLengthString As Boolean
        '//Weather or not it is a Constant Length String
    Dim m_booIsArray As Boolean
        '//Weather or not it is an array
    Dim m_booIgnoreLine As Boolean
        '//In case a type def has an invalid member... (or a comment is introduced)
    Dim m_rimItem As ResultItem
        '//Resultant Item Variable for the Structure's members.
    Dim m_lngItemPosition As Long
    If ScopeTok = 0 Then
        '//If the scope hasn't been identified, then...
        m_lngPosition = LexResult.Tokens.Tokens(Index).Position
            '//Set the position to that of the given index
    Else
        '//... otherwise...
        m_lngPosition = LexResult.Tokens.Tokens(ScopeTok).Position
            '//Set the position to that of the Scope token index.
    End If
    For m_lngLoop = Index + 2 To LexResult.Tokens.Count
        '//Loop through the tokens, excluding the tokens that make up the 'type' and whitespace following
        m_gltTok = LexResult.Tokens.Tokens(m_lngLoop)
            '//Get the current token.
        If m_lngLoop = Index + 2 Then
            '//If the index is where the ID Should be, then...
            If Not m_gltTok.TokenType = G_LTT_Identifier Then
                '//If the token type is not that of an identifier, then...
                Exit Sub
                    '//We've failed
            Else
                '//... otherwise...
                m_stsStatement.Name = m_gltTok.Value
                    '//Store the name
            End If
        Else
            '//... otherwise...
            If Not m_booIgnoreLine Then
                '//If we're not ignoring the line, then...
                Select Case m_gltTok.TokenType
                    '//Select the token's type for processing
                    Case G_LTT_Keyword
                        '//If it's a keyword, then...
                        Select Case LCase$(m_gltTok.Value)
                            '//Select the token's value for processing
                            Case "end"
                                '//If it's the 'End' keyword, then...
                                m_booEndEnc = True
                                    '//Set the 'end encountered' flag
                            Case "type"
                                '//If it's the 'Type' keyword, then...
                                If m_booEndEnc Then
                                    '//If we've encountered the 'End' Keyword, then...
                                    AddSection CodeResult.Sections, ST_Struct, CodeResult.Structs.Count + 1, m_lngPosition
                                        '//Add the structure's section.
                                    m_stsStatement.Section = CodeResult.Sections.Count
                                        '//Set the Structure variable's section pointer.
                                    AddStruct CodeResult.Structs, m_stsStatement, CodeResult.Sections
                                        '//Add the struct to the others.
                                    Index = m_lngLoop
                                    Exit Sub
                                        '//Exit
                                Else
                                    '//... otherwise...
                                    m_booIDEnc = True
                                        '//We've encountered the identifier of the current
                                        '//item (it is allowed to have 'type' for a member's
                                        '//name)
                                    m_rimItem.Name = m_gltTok.Value
                                        '//Set the name
                                    m_lngItemPosition = m_gltTok.Position
                                End If
                            Case "as"
                                '//If it's the 'As' keyword, then...
                                If m_booIDEnc Then
                                    '//If we've encountered the identifier, then...
                                    m_booIDEnc = False
                                        '//Un-set the flag.
                                End If
                                m_booAsEnc = True
                                    '//We've encountered the 'As' keyword.
                            Case AsType(m_gltTok)
                                '//For keyword, or base data types, select them if
                                '//they are valid keyword types
                                If m_booAsEnc Then
                                    '//If we've encounterd the 'As' keyword, then...
                                    m_rimItem.Type = m_gltTok.Value
                                        '//Store the type of the member
                                    If Not LCase(m_gltTok.Value) = "string" Then
                                        '//If it is not a string type, then...
                                        AddStructItem m_stsStatement.Members, m_rimItem, m_lngItemPosition
                                            '//Add it to the array
                                        m_booIDEnc = False
                                        m_booAsEnc = False
                                        m_booIgnoreLine = False
                                        m_booIsArray = False
                                        m_booEndEnc = False
                                            '//Reset flags
                                        m_rimItem = BlankResultItem
                                            '//Clear the variable
                                    Else
                                        '//... otherwise...
                                        m_booString = True
                                            '//Set the string flag
                                    End If
                                End If
                            Case Else
                                '//... otherwise...
                                If m_booEndEnc Then
                                    'If the end has been encountered, then...
                                    m_booIgnoreLine = True
                                        '//Ignore this line, it's invalid syntax.
                                    m_booIDEnc = False
                                    m_booAsEnc = False
                                    m_booIsArray = False
                                    m_booEndEnc = False
                                        '//Reset the flags
                                    m_rimItem = BlankResultItem
                                        '//Clear the variable
                                Else
                                    m_rimItem.Name = m_gltTok.Value
                                    m_lngItemPosition = m_gltTok.Position
                                    m_booIDEnc = True
                                End If
                        End Select
                    Case G_LTT_Operator
                        '//If it is an operator, then...
                        If m_booString Then
                            '//If it's a string, then...
                            If m_gltTok.Value = m_gl_CST_strOprTimes Then
                                '//If it's a multiplication operator, then...
                                m_booConstLengthString = True
                                    '//It's a constant length string
                            End If
                        ElseIf m_booIDEnc And Not m_booAsEnc Then
                            '//... otherwise, if it's id has been encountered, and it's
                            '//not defined as a type, then...
                            If Not m_booIsArray Then
                                '//If it's not an array, then...
                                If m_gltTok.Value = m_gl_CST_strOprLeftBracket Then
                                    '//If the token is a left bracket, then...
                                    If m_booIDEnc Then
                                        '//If the ID has been encountered, then...
                                        m_booIsArray = True
                                            '//It is an array (so far at least).
                                    Else
                                        '//... otherwise...
                                        m_booIgnoreLine = True
                                            '//It's invalid syntax
                                        m_booIDEnc = False
                                        m_booAsEnc = False
                                        m_booIsArray = False
                                        m_booEndEnc = False
                                            '//Reset the flags
                                        m_rimItem = BlankResultItem
                                            '//Clear the variable.
                                    End If
                                Else
                                    m_booIgnoreLine = True
                                        '//It's invalid syntax.
                                    m_booIDEnc = False
                                    m_booAsEnc = False
                                    m_booIsArray = False
                                    m_booEndEnc = False
                                        '//Reset the flags
                                    m_rimItem = BlankResultItem
                                        '//Clear the variable
                                End If
                            Else
                                '//... otherwise (if it is an array so far)...
                                If m_gltTok.Value = "." Then
                                    
                                Else
                                    If m_gltTok.Value = m_gl_CST_strOprRightBracket Then
                                        '//If the token is a right bracket, then...
                                        If m_booIsArray And m_booIDEnc And Not m_booAsEnc Then
                                            '//If it is an array, the id has been encountered, and the type
                                            '//hasn't been defined, then...
                                            m_rimItem.bData = True
                                                '//Set the 'isarray' flag
                                        Else
                                            '//... otherwise...
                                            m_rimItem = BlankResultItem
                                                '//Clear the variable
                                            m_booIgnoreLine = True
                                                '//It is invalid syntax.
                                            m_booIDEnc = False
                                            m_booAsEnc = False
                                            m_booEndEnc = False
                                            m_booIsArray = False
                                                '//Reset the flags
                                        End If
                                    ElseIf m_gltTok.Value = m_gl_CST_strOprComma Then
                                        '//Not yet implemented.
                                    End If
                                End If
                            End If
                        ElseIf m_booEndEnc Then
                            '//If the end has been encountered, then...
                            If Not m_gltTok.Value = m_gl_CST_strOprUnderscore Then
                                '//If the token is anything but an underscore, then...
                                m_booIgnoreLine = True
                                m_booIsArray = False
                                m_booIDEnc = False
                                m_booAsEnc = False
                                m_rimItem = BlankResultItem
                                m_booEndEnc = False
                            ElseIf m_booIDEnc And m_booAsEnc Then
                                m_booIDEnc = False
                                m_booEndEnc = False
                                m_booAsEnc = False
                                m_booIgnoreLine = True
                                m_booIsArray = False
                                m_rimItem = BlankResultItem
                            End If
                        End If
                    Case G_LTT_WhiteSpace
                        If InStr(1, m_gltTok.Value, vbCr) <> 0 Then
                            If m_booString Then
                                m_booString = False
                                m_booConstLengthString = False
                                m_booIDEnc = False
                                m_booEndEnc = False
                                m_booAsEnc = False
                                m_booIsArray = False
                                AddStructItem m_stsStatement.Members, m_rimItem, m_lngItemPosition
                            End If
                            If m_booIDEnc Then
                                If m_booAsEnc Then
                                    m_booIDEnc = False
                                    m_booEndEnc = False
                                    m_booAsEnc = False
                                    m_booIsArray = False
                                    m_rimItem = BlankResultItem
                                Else
                                    m_rimItem.Type = "Variant"
                                    AddStructItem m_stsStatement.Members, m_rimItem, m_lngItemPosition
                                    m_rimItem = BlankResultItem
                                    m_booIDEnc = False
                                    m_booEndEnc = False
                                    m_booAsEnc = False
                                    m_booIsArray = False
                                End If
                            End If
                        End If
                    Case G_LTT_Identifier
                        If m_booAsEnc Then
                            m_rimItem.Type = m_gltTok.Value
                            AddStructItem m_stsStatement.Members, m_rimItem, m_lngItemPosition
                            m_rimItem = BlankResultItem
                            m_booAsEnc = False
                            m_booIDEnc = False
                            m_booString = False
                            m_booConstLengthString = False
                        Else
                            If m_booIDEnc Then
                                m_booIDEnc = False
                                m_booEndEnc = False
                                m_booAsEnc = False
                                m_booIgnoreLine = True
                                m_booIsArray = False
                                m_rimItem = BlankResultItem
                            Else
                                m_rimItem.Name = m_gltTok.Value
                                m_lngItemPosition = m_gltTok.Position
                                m_booIDEnc = True
                            End If
                        End If
                    Case G_LTT_Constant
                        If m_booConstLengthString And m_booString Then
                            If m_gltTok.Value <= 0 Then
                                m_booIDEnc = False
                                m_booEndEnc = False
                                m_booAsEnc = False
                                m_booIgnoreLine = True
                                m_booIsArray = False
                                m_rimItem = BlankResultItem
                            Else
                                m_rimItem.lData = m_gltTok.Value
                                AddStructItem m_stsStatement.Members, m_rimItem, m_lngItemPosition
                                m_booConstLengthString = False
                                m_booString = False
                                m_booIDEnc = False
                                m_booEndEnc = False
                                m_booAsEnc = False
                                m_booIsArray = False
                            End If
                        End If
                    Case G_LTT_Comment
                        m_booIgnoreLine = True
                        If m_booIDEnc Then
                            m_rimItem.Type = "Variant"
                            AddStructItem m_stsStatement.Members, m_rimItem, m_lngItemPosition
                            m_rimItem = BlankResultItem
                            m_booIDEnc = False
                            m_booEndEnc = False
                            m_booAsEnc = False
                            m_booIsArray = False
                        Else
                            If m_booAsEnc Then
                                If m_booString Then
                                    m_booString = False
                                    m_booConstLengthString = False
                                    m_booIDEnc = False
                                    m_booEndEnc = False
                                    m_booAsEnc = False
                                    m_booIsArray = False
                                    AddStructItem m_stsStatement.Members, m_rimItem, m_lngItemPosition
                                ElseIf Len(m_rimItem.Type) = 0 Then
                                    m_booString = False
                                    m_booConstLengthString = False
                                    m_booIDEnc = False
                                    m_booEndEnc = False
                                    m_booAsEnc = False
                                    m_booIsArray = False
                                    m_booIgnoreLine = True
                                End If
                            End If
                        End If
                    Case Else
                        Exit Sub
                End Select
            Else
                If m_gltTok.TokenType = G_LTT_WhiteSpace Then
                    If InStr(1, m_gltTok.Value, vbCr) <> 0 Then
                        m_booIgnoreLine = False
                    End If
                End If
            End If
        End If
    Next
End Sub

Public Sub GetEnum(LexResult As LexicalResult, CodeResult As CodeResult, Index As Long, ScopeTok As Long)
    Dim m_lngLoop As Long
    Dim m_gltTok As LexicalToken
    Dim m_lngLastValue As Long
    Dim m_booIDEnc As Boolean
    Dim m_booValueEnc As Boolean
    Dim m_booEndEnc As Boolean
    Dim m_booEnumEnc As Boolean
    Dim m_booIgnoreLine As Boolean
    Dim m_booEqualsEnc As Boolean
    Dim m_steEnum As EnumStatement
    Dim m_esiItem As EnumStatementItem
    Dim m_lngItemPosition As Long
    Dim m_lngPosition As Long
    m_lngLastValue = -1
    If LexResult.Tokens.Tokens(Index + 1).TokenType = G_LTT_WhiteSpace And LexResult.Tokens.Tokens(Index + 2).TokenType = G_LTT_Identifier Then
        m_steEnum.Name = LexResult.Tokens.Tokens(Index + 2).Value
        If Not ScopeTok = 0 Then
            m_lngPosition = LexResult.Tokens.Tokens(ScopeTok).Position
        Else
            m_lngPosition = LexResult.Tokens.Tokens(Index).Position
        End If
    Else
        Exit Sub
    End If
    For m_lngLoop = Index + 3 To LexResult.Tokens.Count
        m_gltTok = LexResult.Tokens.Tokens(m_lngLoop)
        If Not m_booIgnoreLine Then
            Select Case m_gltTok.TokenType
                Case G_LTT_Comment
                    If m_booIDEnc And Not m_booEqualsEnc Then
                        m_lngLastValue = m_lngLastValue + 1
                        m_esiItem.Value = m_lngLastValue
                        AddEnumItem m_steEnum.Members, m_esiItem, m_lngItemPosition
                        m_esiItem = BlankEnumItem
                        m_booEqualsEnc = False
                        m_booIDEnc = False
                        m_booValueEnc = False
                    ElseIf m_booIDEnc And m_booEqualsEnc And Not m_booValueEnc Then
                        m_booIgnoreLine = True
                    ElseIf m_booIDEnc And m_booEqualsEnc And m_booValueEnc Then
                        '//Do nothing
                    Else
                        m_booIgnoreLine = True
                    End If
                Case G_LTT_Identifier
                    If Not m_booIDEnc Then
                        m_booIDEnc = True
                        m_esiItem.Name = m_gltTok.Value
                        m_lngItemPosition = m_gltTok.Position
                    Else
                        m_booIgnoreLine = True
                    End If
                Case G_LTT_Operator
                    If Not m_booEndEnc Then
                        If m_booIDEnc Then
                            If Not m_booEqualsEnc Then
                                If Not m_booValueEnc Then
                                    If m_gltTok.Value = "=" Then
                                        m_booEqualsEnc = True
                                    End If
                                End If
                            Else
                                m_booEqualsEnc = False
                                m_booValueEnc = False
                                m_booIDEnc = False
                                m_booIgnoreLine = True
                            End If
                        Else
                            m_booIgnoreLine = True
                        End If
                    Else
                        m_booEndEnc = False
                        m_booValueEnc = False
                        m_booIgnoreLine = True
                        m_booEqualsEnc = False
                    End If
                Case G_LTT_Constant
                    If m_booEndEnc Then
                        m_booIDEnc = False
                        m_booValueEnc = False
                        m_booEqualsEnc = False
                        m_booIgnoreLine = True
                        m_booEndEnc = False
                    Else
                        If m_booValueEnc Then
                            m_booIDEnc = False
                            m_booEqualsEnc = False
                            m_booEndEnc = False
                            m_booValueEnc = False
                            m_booIgnoreLine = True
                        Else
                            If m_booIDEnc And m_booEqualsEnc Then
                                m_booValueEnc = True
                                m_esiItem.Value = m_gltTok.Value
                                m_lngLastValue = m_esiItem.Value
                            Else
                                m_booValueEnc = False
                                m_booIDEnc = False
                                m_booIgnoreLine = True
                            End If
                        End If
                    End If
                Case G_LTT_Keyword
                    Select Case m_gltTok.Value
                        Case m_gl_CST_strKwdEnd
                            If Not m_booEndEnc Then
                                If Not m_booIDEnc Then
                                    m_booEndEnc = True
                                Else
                                    m_booIgnoreLine = True
                                    m_booIDEnc = False
                                    m_booEqualsEnc = False
                                    m_booValueEnc = False
                                End If
                            Else
                                m_booIgnoreLine = True
                                m_booEndEnc = False
                                m_booEqualsEnc = False
                                m_booValueEnc = False
                            End If
                        Case m_gl_CST_strKwdEnum
                            If m_booEndEnc Then
                                m_booEnumEnc = True
                                If m_lngLoop = LexResult.Tokens.Count Then
                                    AddEnum CodeResult.Enums, m_steEnum, CodeResult.Sections, m_lngPosition
                                    Index = m_lngLoop
                                    Exit Sub
                                End If
                            Else
                                m_booIgnoreLine = True
                            End If
                    End Select
                Case G_LTT_WhiteSpace
                    If Not m_booEnumEnc Then
                        If InStr(1, m_gltTok.Value, vbCr) > 0 Then
                            If m_booIDEnc And Not m_booEqualsEnc Then
                                m_lngLastValue = m_lngLastValue + 1
                                m_esiItem.Value = m_lngLastValue
                                AddEnumItem m_steEnum.Members, m_esiItem, m_lngItemPosition
                            ElseIf m_booIDEnc And m_booEqualsEnc And Not m_booValueEnc Then
                                '//Do nothing
                            ElseIf m_booIDEnc And m_booEqualsEnc And m_booValueEnc Then
                                AddEnumItem m_steEnum.Members, m_esiItem, m_lngItemPosition
                            End If
                            m_esiItem = BlankEnumItem
                            m_booEqualsEnc = False
                            m_booIDEnc = False
                            m_booValueEnc = False
                        Else
                            '//Do nothing
                        End If
                    Else
                        If m_booEndEnc Then
                            AddEnum CodeResult.Enums, m_steEnum, CodeResult.Sections, m_lngPosition
                            Index = m_lngLoop
                            Exit Sub
                        End If
                    End If
                Case Else
                    m_booIgnoreLine = True
                    m_booValueEnc = False
                    m_booIDEnc = False
                    m_booEndEnc = False
                    m_booEnumEnc = False
            End Select
        Else
            Select Case m_gltTok.TokenType
                Case G_LTT_WhiteSpace
                    If m_booIgnoreLine Then
                        If InStr(1, m_gltTok.Value, vbCr) Then
                            m_booIgnoreLine = False
                        End If
                    End If
            End Select
        End If
    Next
End Sub

Public Sub GetDeclare(LexResult As LexicalResult, CodeResult As CodeResult, Index As Long, ScopeTok As Long)
    Dim m_lngLoop As Long
    Dim m_gltTok As LexicalToken
    Dim m_desDeclare As DeclareStatement
    Dim m_booFunction As Boolean
    Dim m_booNameEnc As Boolean
    Dim m_booAliasEnc As Boolean
    Dim m_booArgsStarted As Boolean
    Dim m_booArgsEnded As Boolean
    Dim m_booArray As Boolean
    Dim m_booAliasFound As Boolean
    Dim m_booUnderscoreFound As Boolean
    Dim m_booLibEnc As Boolean
    Dim m_booAsEnc As Boolean
    Dim m_booAsArgEnc As Boolean
    Dim m_booInArray As Boolean
    Dim m_lngPosition As Long
    Dim m_lngArgPosition As Long
    Dim m_booArgIDEnc As Boolean
    Dim m_booArgTypeEnc As Boolean
    Dim m_booByWhatEnc As Boolean
    Dim m_rimArg As ResultItem
    Dim m_booVoid As Boolean
    With LexResult.Tokens
        If .Tokens(Index + 1).TokenType = G_LTT_WhiteSpace And .Tokens(Index + 2).TokenType = G_LTT_Keyword Then
            If ((.Tokens(Index + 2).Value = m_gl_CST_strKwdFunction Or .Tokens(Index + 2).Value = m_gl_CST_strKwdSub) And .Tokens(Index + 3).TokenType = G_LTT_WhiteSpace And .Tokens(Index + 4).TokenType = G_LTT_Identifier) Then
                m_booVoid = .Tokens(Index + 2).Value = m_gl_CST_strKwdSub
                m_desDeclare.Header.Name = .Tokens(Index + 4).Value
                If ScopeTok = 0 Then
                    With .Tokens(Index)
                        m_lngPosition = .Position
                    End With
                Else
                    With .Tokens(ScopeTok)
                        m_lngPosition = .Position
                    End With
                End If
            Else
                Exit Sub
            End If
        Else
            Exit Sub
        End If
    End With
    m_booNameEnc = True
    For m_lngLoop = Index + 5 To LexResult.Tokens.Count
        m_gltTok = LexResult.Tokens.Tokens(m_lngLoop)
        Select Case m_gltTok.TokenType
            Case G_LTT_Comment
                If m_booArgsStarted And Not m_booArgsEnded Then
                    Exit Sub
                ElseIf m_booArgsStarted And m_booArgsEnded Then
                    If m_booAsEnc Then
                        Exit Sub
                    End If
                    '//Do nothing
                End If
            Case G_LTT_String
                If Not m_booArgsStarted And Not m_booArgsEnded Then
                    If m_booNameEnc And m_booLibEnc And Not m_booAliasEnc Then
                        m_desDeclare.Library = DeflateString(CStr(m_gltTok.Value), 2)
                    ElseIf m_booNameEnc And m_booLibEnc And m_booAliasEnc Then
                        m_desDeclare.Alias = DeflateString(CStr(m_gltTok.Value), 2)
                    End If
                Else
                    Exit Sub
                End If
            Case G_LTT_Keyword
                Select Case m_gltTok.Value
                    Case m_gl_CST_strKwdByRef
                        If Not m_booByWhatEnc Then
                            m_rimArg.ByVal = False
                            m_booByWhatEnc = True
                            m_lngArgPosition = m_gltTok.Position
                        Else
                            Exit Sub
                        End If
                    Case m_gl_CST_strKwdByVal
                        If m_booArgsStarted Then
                            If Not (m_booArgIDEnc Or m_booAsArgEnc Or m_booArgTypeEnc) And Not m_booByWhatEnc Then
                                m_rimArg.ByVal = True
                                m_booByWhatEnc = True
                                m_lngArgPosition = m_gltTok.Position
                            Else
                                Exit Sub
                            End If
                        End If
                    Case m_gl_CST_strKwdLibrary
                        If m_booNameEnc Then
                            m_booLibEnc = True
                        End If
                    Case m_gl_CST_strKwdAlias
                        If m_booAliasEnc Then
                            Exit Sub
                        Else
                            If m_booNameEnc And m_booLibEnc Then
                                m_booAliasEnc = True
                            Else
                                Exit Sub
                            End If
                        End If
                    Case m_gl_CST_strKwdAs
                        If m_booArgsStarted Then
                            If m_booAsArgEnc Then
                                Exit Sub
                            Else
                                If m_booArgIDEnc Then
                                    m_booAsArgEnc = True
                                End If
                            End If
                        ElseIf m_booArgsEnded Then
                            If m_booAsEnc Then
                                Exit Sub
                            Else
                                m_booAsEnc = True
                            End If
                        End If
                    Case AsType(m_gltTok)
                        If m_booAsArgEnc Or m_booAsEnc Then
                            If m_booAsEnc Then
                                If m_booArgsEnded Then
                                    m_desDeclare.Header.Type = m_gltTok.Value
                                    AddSection CodeResult.Sections, ST_Declare, CodeResult.Declares.Count + 1, m_lngPosition
                                    m_desDeclare.Section = CodeResult.Sections.Count
                                    AddDeclare CodeResult.Declares, m_desDeclare, CodeResult.Sections
                                    Exit Sub
                                End If
                            Else
                                m_rimArg.Type = m_gltTok.Value
                                m_booArgTypeEnc = True
                            End If
                        Else
                            Exit Sub
                        End If
                End Select
            Case G_LTT_Identifier
                If m_booArgsStarted Then
                    If Not m_booAsArgEnc Then
                        If Not m_booArgIDEnc Then
                            m_booArgIDEnc = True
                            m_rimArg.Name = m_gltTok.Value
                            If Not m_booByWhatEnc Then
                                m_lngArgPosition = m_gltTok.Position
                            End If
                        End If
                    Else
                        m_rimArg.Type = m_gltTok.Value
                        m_booArgTypeEnc = True
                    End If
                ElseIf m_booArgsEnded Then
                    If m_booAsEnc Then
                        m_desDeclare.Header.Type = m_gltTok.Value
                        AddSection CodeResult.Sections, ST_Declare, CodeResult.Declares.Count + 1, m_lngPosition
                        m_desDeclare.Section = CodeResult.Sections.Count
                        AddDeclare CodeResult.Declares, m_desDeclare, CodeResult.Sections
                        Exit Sub
                    Else
                        Exit Sub
                    End If
                Else
                
                End If
            Case G_LTT_String
                
            Case G_LTT_WhiteSpace
                If m_booUnderscoreFound Then
                    If (InStr(1, m_gltTok.Value, vbCr) = 0) Then
                        Exit Sub
                    Else
                        m_booUnderscoreFound = False
                    End If
                Else
                    If InStr(1, m_gltTok.Value, vbCr) Then
                        Exit Sub
                    End If
                End If
            Case G_LTT_Operator
                Select Case m_gltTok.Value
                    Case m_gl_CST_strOprUnderscore
                        If Not m_booUnderscoreFound Then
                            m_booUnderscoreFound = True
                        Else
                            Exit Sub
                        End If
                    Case m_gl_CST_strOprLeftBracket
                        If Not m_booArgsStarted Then
                            m_booArgsStarted = True
                        Else
                            
                        End If
                    Case m_gl_CST_strOprRightBracket
                        If m_booArgsStarted Then
                            If Not m_booInArray Then
                                m_booArgsStarted = False
                                m_booArgsEnded = True
                                If m_booArgIDEnc And m_booAsArgEnc And m_booArgTypeEnc Then
                                    AddStructItem m_desDeclare.Arguments, m_rimArg, m_lngArgPosition
                                    m_rimArg = BlankResultItem
                                    m_booArgIDEnc = False
                                    m_booArgTypeEnc = False
                                    m_booAsArgEnc = False
                                    m_booArray = False
                                    m_booByWhatEnc = False
                                ElseIf m_booArgIDEnc And Not m_booAsArgEnc And Not m_booArgTypeEnc Then
                                    AddStructItem m_desDeclare.Arguments, m_rimArg, m_lngArgPosition
                                    m_rimArg = BlankResultItem
                                    m_booArgIDEnc = False
                                    m_booArgTypeEnc = False
                                    m_booAsArgEnc = False
                                    m_booArray = False
                                    m_booByWhatEnc = False
                                End If
                                If m_booVoid Then
                                    m_desDeclare.Header.Type = "void"
                                    AddSection CodeResult.Sections, ST_Declare, CodeResult.Declares.Count + 1, m_lngPosition
                                    m_desDeclare.Section = CodeResult.Sections.Count
                                    AddDeclare CodeResult.Declares, m_desDeclare, CodeResult.Sections
                                    Exit Sub
                                End If
                            Else
                                m_booInArray = False
                            End If
                        End If
                    Case m_gl_CST_strOprComma
                        If m_booArgsStarted Then
                            If m_booArgIDEnc And m_booAsArgEnc And m_booArgTypeEnc Then
                                AddStructItem m_desDeclare.Arguments, m_rimArg, m_lngArgPosition
                                m_rimArg = BlankResultItem
                                m_booArgIDEnc = False
                                m_booArgTypeEnc = False
                                m_booAsArgEnc = False
                                m_booArray = False
                                m_booByWhatEnc = False
                            ElseIf m_booArgIDEnc And Not m_booAsArgEnc And Not m_booArgTypeEnc Then
                                AddStructItem m_desDeclare.Arguments, m_rimArg, m_lngArgPosition
                                m_rimArg = BlankResultItem
                                m_booArgIDEnc = False
                                m_booArgTypeEnc = False
                                m_booAsArgEnc = False
                                m_booArray = False
                                m_booByWhatEnc = False
                            ElseIf Not m_booArgIDEnc Then
                                
                            End If
                        Else
                            Exit Sub
                        End If
                End Select
        End Select
    Next
End Sub

Public Sub AddStruct(Structs As StructStatements, Struct As StructStatement, Sections As CodeSections)
    Dim m_lngLoop As Long
    With Structs
        If .Count = 0 Then
            ReDim .Structs(1 To .Count + 1)
        Else
            ReDim Preserve .Structs(1 To .Count + 1)
        End If
        .Count = .Count + 1
        .Structs(.Count) = Struct
        For m_lngLoop = 1 To .Structs(.Count).Members.Count
            With .Structs(.Count).Members.Items(m_lngLoop)
                AddSection Sections, ST_StructItem, m_lngLoop, .Section
                .Section = Sections.Count
            End With
        Next
    End With
End Sub

Private Function AsType(Token As LexicalToken) As String
    Select Case LCase(Token.Value)
        Case "string", "long", "double", "date", "integer", "byte", "single", "boolean", "variant"
            AsType = (Token.Value)
    End Select
End Function

Private Function BlankResultItem() As ResultItem
    '//
End Function

Public Sub AddStructItem(Members As ResultItems, Item As ResultItem, Position As Long)
    With Members
        If .Count = 0 Then
            ReDim .Items(1 To .Count + 1)
        Else
            ReDim Preserve .Items(1 To .Count + 1)
        End If
        .Count = .Count + 1
        .Items(.Count) = Item
        With .Items(.Count)
            .Section = Position
        End With
        Position = 0
    End With
End Sub

Public Sub AppendArrayStr(ArrayStr() As String, ByVal AppendText As String)
    ArrayStr(UBound(ArrayStr)) = ArrayStr(UBound(ArrayStr)) & AppendText
End Sub

Public Sub AddEnum(Enums As EnumStatements, EnumVar As EnumStatement, Sections As CodeSections, Position As Long)
    Dim m_lngLoop As Long
    With Enums
        If .Count = 0 Then
            ReDim .Enums(1 To .Count + 1)
        Else
            ReDim Preserve .Enums(1 To .Count + 1)
        End If
        .Count = .Count + 1
        .Enums(.Count) = EnumVar
        With .Enums(.Count)
            AddSection Sections, ST_Enum, Enums.Count, Position
            .Section = Sections.Count
        End With
        For m_lngLoop = 1 To .Enums(.Count).Members.Count
            With .Enums(.Count).Members.Item(m_lngLoop)
                AddSection Sections, ST_EnumItem, m_lngLoop, .Section
                .Section = Sections.Count
            End With
        Next
    End With
End Sub

Public Sub AddEnumItem(EnumMembers As EnumStatementItems, Item As EnumStatementItem, Position As Long)
    With EnumMembers
        If .Count = 0 Then
            ReDim .Item(1 To .Count + 1)
        Else
            ReDim Preserve .Item(1 To .Count + 1)
        End If
        'Debug.Print Item.Name & " = " & "0x" & Hex(Item.Value) & ";"
        .Count = .Count + 1
        .Item(.Count) = Item
        With .Item(.Count)
            .Section = Position
        End With
    End With
End Sub

Public Function BlankEnumItem() As EnumStatementItem
    
End Function

Public Sub AddDeclare(Declares As DeclareStatements, DeclareStatement As DeclareStatement, Sections As CodeSections)
    Dim m_lngSubItem As Long
    With Declares
        If .Count = 0 Then
            ReDim .Statements(1 To .Count + 1)
        Else
            ReDim Preserve .Statements(1 To .Count + 1)
        End If
        .Count = .Count + 1
        For m_lngSubItem = 1 To DeclareStatement.Arguments.Count
            With DeclareStatement.Arguments.Items(m_lngSubItem)
                AddSection Sections, ST_DeclareArgument, m_lngSubItem, .Section
                .Section = Sections.Count
            End With
        Next
        .Statements(.Count) = DeclareStatement
    End With
End Sub

Public Function ItemExists(Col As Collection, Key As String) As Boolean
    Dim m_booResult As Boolean
Try:
    On Error GoTo Catch
    Call Col(Key)
    m_booResult = True
    GoTo Finally
Catch:
    m_booResult = False
Finally:
    ItemExists = m_booResult
EndTry:
    Exit Function
End Function
