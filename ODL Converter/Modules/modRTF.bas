Attribute VB_Name = "modRTF"
Option Explicit

Public Function MakeCRTF(Text As String) As String
    Dim m_slnLanguage As SCLanguage
    Set m_slnLanguage = New SCLanguage
    Set m_slnLanguage.Engine = New ODLanguage
    MakeCRTF = MakeRTF(m_slnLanguage.Engine, Text)
End Function

Public Function GenerateFontTBL(ParamArray FontNames() As Variant)
    Dim m_strResult As String
    Dim m_strItem As String
    Dim m_varItem As Variant
    Dim m_lngIndex As Long
    m_strResult = "{\fonttbl"
    For Each m_varItem In FontNames
        m_strItem = "{\f" & m_lngIndex & "\fcharset0 " & m_varItem & ";}"
        m_strResult = m_strResult & m_strItem
        m_lngIndex = m_lngIndex + 1
    Next
    m_strResult = m_strResult & "}"
    GenerateFontTBL = m_strResult
End Function

Public Function GenerateColorTBL(ParamArray Colors() As Variant)
    Dim m_strResult As String
    Dim m_strItem As String
    Dim m_varItem As Variant
    Dim m_lngRed As Long, _
        m_lngGreen As Long, _
        m_lngBlue As Long
    m_strResult = "{\colortbl ;"
    For Each m_varItem In Colors
        m_lngRed = m_varItem Mod 256
        m_lngGreen = m_varItem \ 256 Mod 256
        m_lngBlue = m_varItem \ 256 \ 256 Mod 256
        m_strItem = "\red" & m_lngRed & "\green" & m_lngGreen & "\blue" & m_lngBlue & ";"
        m_strResult = m_strResult & m_strItem
    Next
    m_strResult = m_strResult & "}"
    GenerateColorTBL = m_strResult
End Function

Public Function MakeRTF(Engine As SCLangEngine, Text As String) As String
    Const BufferSize As Long = 2 ^ 11 '//2048, powered to allow ease of
        '//increase or decrease on a half or double basis
    Dim m_clnEngine As ODLanguage
    Dim m_blnEngine As BasicLanguage
    Dim m_lreResult As LexicalResult
    Dim m_lngToken As Long
    Dim m_lngBufferIndex As Long
    Dim m_staBuffer() As String
    Dim m_ltoToken As LexicalToken
    Dim m_lngLastType As Long
    Dim m_strPass As String
        '//Text for this token's pass...
    ReDim m_staBuffer(0)
    If TypeOf Engine Is ODLanguage Then
        Set m_clnEngine = Engine
        m_lreResult = m_clnEngine.LexicalParse(Text)
    ElseIf TypeOf Engine Is BasicLanguage Then
        Set m_blnEngine = Engine
        m_lreResult = m_blnEngine.LexicalParse(Text)
    End If
    m_staBuffer(0) = "{\rtf1\ansi\deff0" & GenerateFontTBL("Courier New") & vbCrLf
    m_staBuffer(0) = m_staBuffer(0) & GenerateColorTBL(RGB(0, 128, 128), vbRed, vbBlue, RGB(128, 0, 128), RGB(0, 128, 0)) & vbCrLf & "\fs20 "
    m_lngLastType = -1
    For m_lngToken = 1 To m_lreResult.Tokens.Count
        m_ltoToken = m_lreResult.Tokens.Tokens(m_lngToken)
        m_strPass = vbNullString
        If Not m_ltoToken.TokenType = G_LTT_WhiteSpace Then
            m_ltoToken.Value = Replace(m_ltoToken.Value, "\", "\\")
            m_ltoToken.Value = Replace(m_ltoToken.Value, "{", "\{")
            m_ltoToken.Value = Replace(m_ltoToken.Value, "}", "\}")
        End If
        Select Case m_ltoToken.TokenType
            Case LexicalTokenType.G_LTT_Comment
                If m_lngLastType <> m_ltoToken.TokenType Then
                    m_strPass = "\cf5 "
                End If
                m_strPass = m_strPass & m_ltoToken.Value
            Case LexicalTokenType.G_LTT_Constant
                If m_lngLastType <> m_ltoToken.TokenType Then
                    m_strPass = "\cf2 "
                End If
                If m_ltoToken.CustID Then
                    '//Checking this is just good coding...
                    m_strPass = m_strPass & "0x" & Hex(m_ltoToken.Value)
                Else
                    m_strPass = m_strPass & m_ltoToken.Value
                End If
            Case LexicalTokenType.G_LTT_Identifier
                If m_lngLastType <> m_ltoToken.TokenType Then
                    m_strPass = "\cf1 "
                End If
                m_strPass = m_strPass & m_ltoToken.Value
            Case LexicalTokenType.G_LTT_Keyword
                If m_lngLastType <> m_ltoToken.TokenType Then
                    m_strPass = "\cf3 "
                End If
                m_strPass = m_strPass & Trim(m_ltoToken.Value)
            Case LexicalTokenType.G_LTT_Operator
                If m_lngLastType <> m_ltoToken.TokenType Then
                    m_strPass = "\cf4 "
                End If
                m_strPass = m_strPass & m_ltoToken.Value
            Case LexicalTokenType.G_LTT_String
                If m_lngLastType <> m_ltoToken.TokenType Then
                    m_strPass = "\cf5 "
                End If
                m_strPass = m_strPass & m_ltoToken.Value
            Case LexicalTokenType.G_LTT_WhiteSpace
                m_strPass = m_ltoToken.Value
                m_strPass = Replace(m_strPass, vbCrLf, "\par" & vbCrLf)
                    '//RTF style breaks
        End Select
        m_staBuffer(m_lngBufferIndex) = m_staBuffer(m_lngBufferIndex) & m_strPass
        If Len(m_staBuffer(m_lngBufferIndex)) >= BufferSize Then
            m_lngBufferIndex = m_lngBufferIndex + 1
            ReDim Preserve m_staBuffer(m_lngBufferIndex)
        End If
        If Not m_ltoToken.TokenType = G_LTT_WhiteSpace Then
            m_lngLastType = m_ltoToken.TokenType
        End If
    Next
    m_staBuffer(m_lngBufferIndex) = m_staBuffer(m_lngBufferIndex) & "}" & vbCrLf & m_gl_CST_strOprSpace
    MakeRTF = Join(m_staBuffer, vbNullString)
End Function
