Attribute VB_Name = "modConversion"
Option Explicit

Public Function GetTypeConversion(ByVal AsType As String)
    Select Case LCase$(AsType)
        Case "string"
            GetTypeConversion = "BSTR"
        Case "boolean"
            GetTypeConversion = "boolean"
        Case "long"
            GetTypeConversion = "long"
        Case "byte"
            GetTypeConversion = "unsigned char"
        Case "integer"
            GetTypeConversion = "int"
        Case "single"
            GetTypeConversion = "float"
        Case "double"
            GetTypeConversion = "double"
        Case "date"
            GetTypeConversion = "?"
        Case Else
            GetTypeConversion = AsType
    End Select
End Function

Public Function GetDeclConversion(ByVal AsType As String, IsEnum)
    Select Case LCase$(AsType)
        Case "string"
            GetDeclConversion = "LPSTR"
        Case "boolean"
            GetDeclConversion = "boolean"
        Case "long"
            GetDeclConversion = "long"
        Case "byte"
            GetDeclConversion = "unsigned char"
        Case "integer"
            GetDeclConversion = "int"
        Case "single"
            GetDeclConversion = "float"
        Case "double"
            GetDeclConversion = "double"
        Case "date"
            GetDeclConversion = "?"
        Case "any"
            GetDeclConversion = "void*"
        Case "void"
            GetDeclConversion = "void"
        Case Else
            If IsEnum Then
                GetDeclConversion = AsType
            Else
                GetDeclConversion = AsType
            End If
    End Select
End Function
