VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SCLanguages"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private m_colLangauges As Collection

Public Property Get Count() As Long
    Count = m_colLangauges.Count
End Property

Public Property Get Item(Index As Variant) As SCLanguage
Attribute Item.VB_UserMemId = 0
    Set Item = m_colLangauges(Index)
End Property

Public Property Get NewEnum() As stdole.IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "440"
    Set NewEnum = m_colLangauges.[_NewEnum]
End Property

Public Function Add(Name As String, Optional Engine As SCLangEngine) As SCLanguage
Declarations:
    Dim m_sclLanguage As SCLanguage
Try:
    On Error GoTo Catch
    Set m_sclLanguage = New SCLanguage
    With m_sclLanguage
        '//Select the Language namespace
        Set .Collection = Me
            '//Set the parent collection
        .Name = Name
            '//Set the name.
        If Not Engine Is Nothing Then _
            Set .Engine = Engine
            '//If the engine is present, then...
                '//Set the reference.
        m_colLangauges.Add m_sclLanguage, Name
    End With
Finally:
    '//Cleanup
    Set Add = m_sclLanguage
        '//Return the language reference.
    Set m_sclLanguage = Nothing
        '//Destroy the reference
    GoTo EndTry
        '//Exit
Catch:
    Set m_sclLanguage = Nothing
        '//Destroy current reference.
    Set Add = Nothing
    '//Make sure it's nothing
EndTry:
    Exit Function
        '//Exit the procedure safely
End Function

Public Sub Remove(Index As Variant)
    m_colLangauges.Remove Index
End Sub

Public Sub Clear()
    Do Until Count = 0
        m_colLangauges.Remove 1
    Loop
End Sub

Private Sub Class_Initialize()
    Set m_colLangauges = New Collection
End Sub

Private Sub Class_Terminate()
    Set m_colLangauges = Nothing
End Sub
