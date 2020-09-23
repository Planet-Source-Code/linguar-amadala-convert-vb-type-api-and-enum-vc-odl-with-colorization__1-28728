VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SCLangEngine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private r_gslLanguage As SCLanguage

Public Property Get Language() As SCLanguage
    Set Language = r_gslLanguage
End Property

Public Property Set Language(ByVal v_gslLanguage As SCLanguage)
    Set r_gslLanguage = v_gslLanguage
End Property

Friend Function LexicalParse(ByVal Expression As String) As LexicalResult
    
End Function

Friend Function SemanticParse(LexicalInput As LexicalResult) As SemanticResult
    
End Function

Friend Function StandardParse(ByVal Expression As String) As SemanticResult
    StandardParse = SemanticParse(LexicalParse(Expression))
End Function
