Attribute VB_Name = "modLangTypes"
Option Explicit
Public Enum SemanticLineItemType
    G_SLIT_Keyword = 0
    G_SLIT_String
    G_SLIT_Constant
    G_SLIT_Operator
    G_SLIT_BlockStart
    G_SLIT_BlockEnd
End Enum
Public Enum LexicalTokenType
    G_LTT_Constant = 0
    G_LTT_String
    G_LTT_Identifier
    G_LTT_WhiteSpace
    G_LTT_Keyword
    G_LTT_Operator
    G_LTT_Comment
    G_LTT_Custom
End Enum
Public Type LexicalToken
    TokenType As LexicalTokenType
    Value As Variant
    Length As Long
    Position As Long
    CustID As Long
End Type
Public Type LexicalTokens
    Count As Long
    Tokens() As LexicalToken
End Type
Public Type LexicalResult
    Tokens As LexicalTokens
    Data As Variant
End Type
Public Type SemanticLineItem
    ItemType As SemanticLineItemType
    Token As LexicalToken
End Type
Public Type SemanticLineItems
    Count As Long
    Lines() As SemanticLineItem
End Type
Public Type SemanticLine
    LineTypeID As Long
    Items As SemanticLineItems
End Type
Public Type SemanticSection
    Count As Long
    Lines() As SemanticLine
End Type
Public Type SemanticResult
    Count As Long
    Sections() As SemanticSection
End Type
Public Type Operator
    Value As String
    Description As String
    Priority As Long
End Type
Public Type Operators
    Count As Long
    Operators() As Operator
End Type
Public Enum KeywordFlags
    G_KF_Flexable = 0&
    G_KF_Strict = 1&
    G_KF_Special = 2
End Enum
Public Type Keyword
    StringValue As String
    Description As String
    Flags As KeywordFlags
End Type
Public Type Keywords
    Count As Long
    Keywords() As Keyword
End Type
Public Type LexicalProcResult
    Success As Boolean
        '//Function was successful
    Token As LexicalToken
        '//Token to be inserted into the loop
    NewPosition As Long
        '//When a lexical process ends, the new position _
           is used to ensure that proper processing occurs.
End Type
Public Type LexicalProcess
    Expression As String
        '//Code in...
    ExpressionLength As Long
        '//Length of the expression
    SelExpression As String
        '//Selected expression
    CharIsAlpha As Boolean
        '//If the current character is an alpha-numeric character
    CharIsOperator As Boolean
        '//If the current character is an operator
    CharIsNumeric As Boolean
        '//If the current character is a number
    ExprIsKeyWord As Boolean
        '//If the current selected expression is a keyword
    CharIndex As Long
        '//Current Character index in this sequence
    Char As String
        '//Current Character
    Position As Long
        '//Position in the expression
    Result As LexicalProcResult
End Type

