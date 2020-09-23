VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form frmMain 
   Caption         =   "ODL Converter"
   ClientHeight    =   5925
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   6750
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   395
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   450
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog cdFile 
      Left            =   3135
      Top             =   2715
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      Filter          =   "Basic Files (*.bas)|*.bas|All Files|*.*"
   End
   Begin RichTextLib.RichTextBox rtbODL 
      Height          =   4455
      Left            =   240
      TabIndex        =   3
      Top             =   600
      Visible         =   0   'False
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   7858
      _Version        =   393217
      BorderStyle     =   0
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   3
      RightMargin     =   90000
      TextRTF         =   $"frmMain.frx":000C
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton cmdTest 
      Caption         =   "Process"
      Height          =   495
      Left            =   2760
      TabIndex        =   2
      Top             =   5280
      Width           =   1215
   End
   Begin RichTextLib.RichTextBox rtbCode 
      Height          =   4455
      Left            =   240
      TabIndex        =   1
      Top             =   600
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   7858
      _Version        =   393217
      BorderStyle     =   0
      Enabled         =   -1  'True
      HideSelection   =   0   'False
      ReadOnly        =   -1  'True
      ScrollBars      =   3
      RightMargin     =   9e6
      AutoVerbMenu    =   -1  'True
      TextRTF         =   $"frmMain.frx":0087
   End
   Begin MSComctlLib.TreeView tvMembers 
      Height          =   4455
      Left            =   240
      TabIndex        =   4
      Top             =   600
      Visible         =   0   'False
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   7858
      _Version        =   393217
      Indentation     =   0
      LineStyle       =   1
      Style           =   7
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.TabStrip tsCode 
      Height          =   5055
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   8916
      HotTracking     =   -1  'True
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   3
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Code"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Members"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "ODL Result"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuFileOpenModule 
         Caption         =   "Open Module..."
      End
      Begin VB.Menu mnuFileImportModule 
         Caption         =   "Import Module..."
      End
      Begin VB.Menu mnuFileExportODL 
         Caption         =   "Export ODL Result"
      End
      Begin VB.Menu mnuFileS1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'//[entry("alias"), helpstring("alias")] type Name({[in, out] argtype argname});
'//void* = any
'//type* = struct
Private m_corResult As CodeResult
Private m_clsLangs As SCLanguages

Private Sub cmdTest_Click()
    m_corResult = SemanticParse(rtbCode.Text, m_clsLangs(1).Engine)
    UpdateTreeView
    UpdateODLView
End Sub

Private Sub Form_Load()
    Set m_clsLangs = New SCLanguages
    m_clsLangs.Add "Basic", New BasicLanguage
End Sub

Private Sub Form_Resize()
    Dim m_objMove As Object
    On Error Resume Next
    Select Case tsCode.SelectedItem.Index
        Case 1
            Set m_objMove = rtbCode
        Case 2
            Set m_objMove = tvMembers
        Case 3
            Set m_objMove = rtbODL
    End Select
    cmdTest.Move (ScaleWidth - cmdTest.Width) / 2, ScaleHeight - 3 - cmdTest.Height
    tsCode.Move 3, 3, ScaleWidth - 6, cmdTest.Top - 9
    If m_objMove Is Nothing Then _
        Exit Sub
    m_objMove.Move tsCode.ClientLeft + 3, tsCode.ClientTop + 3, tsCode.ClientWidth - 6, tsCode.ClientHeight - 6
End Sub

Private Sub mnuFileExit_Click()
    Unload Me
End Sub

Private Sub mnuFileExportODL_Click()
    Dim m_strFile As String
    Dim m_lngFile As Long
    On Error GoTo Catch
    cdFile.FileName = vbNullString
    cdFile.Filter = "Text Files (*.txt)|*.txt|All Files|*.*"
    cdFile.ShowSave
    m_strFile = cdFile.FileName
    m_lngFile = FreeFile
    Open m_strFile For Output Lock Write As #m_lngFile
        Print #m_lngFile, rtbODL.Text
    Close #m_lngFile
    Exit Sub
Catch:
End Sub

Private Sub mnuFileImportModule_Click()
    Const BufferSize = 2 ^ 11
    Dim m_strFile As String
    Dim m_strLine As String
    Dim m_lngFile As Long
    Dim m_lngBufferIndex As Long
    Dim m_staBuffer() As String
    On Error GoTo Catch
    ReDim m_staBuffer(0)
    cdFile.FileName = vbNullString
    cdFile.Filter = "Basic Files (*.bas)|*.bas|All Files|*.*"
    cdFile.ShowOpen
    m_strFile = cdFile.FileName
    m_lngFile = FreeFile
    Open m_strFile For Input Lock Read As #m_lngFile
        Do Until EOF(m_lngFile)
            Line Input #m_lngFile, m_strLine
            If Not LCase(m_strLine) Like "attribute vb_*=*""*""" Then
                m_staBuffer(m_lngBufferIndex) = m_staBuffer(m_lngBufferIndex) & m_strLine & vbCrLf
                If Len(m_staBuffer(m_lngBufferIndex)) > BufferSize Then
                    m_lngBufferIndex = m_lngBufferIndex + 1
                    ReDim Preserve m_staBuffer(m_lngBufferIndex)
                End If
            End If
        Loop
    Close #m_lngFile
    GoTo Finally
Catch:
    
Finally:
    rtbCode.TextRTF = MakeRTF(m_clsLangs(1).Engine, rtbCode.Text & vbCrLf & Join(m_staBuffer, vbNullString))
End Sub

Private Sub mnuFileOpenModule_Click()
    Const BufferSize = 2 ^ 11
    Dim m_strFile As String
    Dim m_strLine As String
    Dim m_lngFile As Long
    Dim m_lngBufferIndex As Long
    Dim m_staBuffer() As String
    On Error GoTo Catch
    ReDim m_staBuffer(0)
    cdFile.FileName = vbNullString
    cdFile.Filter = "Basic Files (*.bas)|*.bas|All Files|*.*"
    cdFile.ShowOpen
    m_strFile = cdFile.FileName
    m_lngFile = FreeFile
    Open m_strFile For Input Lock Read As #m_lngFile
        Do Until EOF(m_lngFile)
            Line Input #m_lngFile, m_strLine
            If Not LCase(m_strLine) Like "attribute vb_*=*""*""" Then
                m_staBuffer(m_lngBufferIndex) = m_staBuffer(m_lngBufferIndex) & m_strLine & vbCrLf
                If Len(m_staBuffer(m_lngBufferIndex)) > BufferSize Then
                    m_lngBufferIndex = m_lngBufferIndex + 1
                    ReDim Preserve m_staBuffer(m_lngBufferIndex)
                End If
            End If
        Loop
    Close #m_lngFile
    GoTo Finally
Catch:
    
Finally:
    rtbCode.TextRTF = MakeRTF(m_clsLangs(1).Engine, Join(m_staBuffer, vbNullString))
End Sub

Private Sub tsCode_Click()
    Dim m_objVis As Object
    Select Case tsCode.SelectedItem.Index
        Case 1
            Set m_objVis = rtbCode
        Case 2
            Set m_objVis = tvMembers
        Case 3
            Set m_objVis = rtbODL
    End Select
    rtbCode.Visible = False
    rtbODL.Visible = False
    tvMembers.Visible = False
    If m_objVis Is Nothing Then _
        Exit Sub
    Form_Resize
    m_objVis.Visible = True
    m_objVis.SetFocus
End Sub

Private Sub UpdateTreeView()
    Dim m_tvnNodeMain As Node
    Dim m_tvnSub As Node
    Dim m_lngLoop As Long
    Dim m_lngSubItem As Long
    Dim m_stsStatement As StructStatement
    Dim m_desStatement As DeclareStatement
    Dim m_rimMember As ResultItem
    Dim m_ensStatement As EnumStatement
    Dim m_esiItem As EnumStatementItem
    tvMembers.Nodes.Clear
    For m_lngLoop = 1 To m_corResult.Enums.Count
        m_ensStatement = m_corResult.Enums.Enums(m_lngLoop)
        Set m_tvnNodeMain = tvMembers.Nodes.Add(, , "tag" & m_ensStatement.Name, m_ensStatement.Name)
        m_tvnNodeMain.Tag = "ens_" & m_lngLoop
        For m_lngSubItem = 1 To m_ensStatement.Members.Count
            m_esiItem = m_ensStatement.Members.Item(m_lngSubItem)
            Set m_tvnSub = tvMembers.Nodes.Add(m_tvnNodeMain, tvwChild, , m_esiItem.Name & " = " & m_esiItem.Value & " " & "(&H" & Hex(m_esiItem.Value) & ")")
            m_tvnSub.Tag = "esi_" & m_lngSubItem
        Next
    Next
    For m_lngLoop = 1 To m_corResult.Structs.Count
        m_stsStatement = m_corResult.Structs.Structs(m_lngLoop)
        Set m_tvnNodeMain = tvMembers.Nodes.Add(, , , m_stsStatement.Name)
        m_tvnNodeMain.Tag = "tag_" & m_lngLoop
        With m_stsStatement
            For m_lngSubItem = 1 To .Members.Count
                m_rimMember = .Members.Items(m_lngSubItem)
                Set m_tvnSub = tvMembers.Nodes.Add(m_tvnNodeMain, tvwChild, , m_rimMember.Name & " As " & m_rimMember.Type)
                m_tvnSub.Tag = "tsb_" & m_lngSubItem
            Next
        End With
    Next
    For m_lngLoop = 1 To m_corResult.Declares.Count
        m_desStatement = m_corResult.Declares.Statements(m_lngLoop)
        Set m_tvnNodeMain = tvMembers.Nodes.Add(, , , m_desStatement.Header.Name & " As " & m_desStatement.Header.Type)
        m_tvnNodeMain.Tag = "dec_" & m_lngLoop
        For m_lngSubItem = 1 To m_desStatement.Arguments.Count
            m_rimMember = m_desStatement.Arguments.Items(m_lngSubItem)
            Set m_tvnSub = tvMembers.Nodes.Add(m_tvnNodeMain, tvwChild, , m_rimMember.Name & " As " & m_rimMember.Type)
            m_tvnSub.Tag = "dea_" & m_lngSubItem
        Next
    Next
End Sub

Private Sub tvMembers_DblClick()
    Dim m_nodNode As Node
    Dim m_lngLoop As Long
    Dim m_lngItemPtr As Long
    Dim m_lngSubItemPtr As Long
    Dim m_stsStruct As StructStatement
    Dim m_ensEnum As EnumStatement
    Dim m_esiItem As EnumStatementItem
    Dim m_rimMember As ResultItem
    Dim m_desDeclare As DeclareStatement
    Set m_nodNode = tvMembers.SelectedItem
    On Error Resume Next
    If Not m_nodNode Is Nothing Then
        If Not m_nodNode.Tag = vbNullString Then
            Select Case VBA.Left$(m_nodNode.Tag, 3)
                Case "tag"
                    m_lngItemPtr = Mid(m_nodNode.Tag, 5)
                    m_stsStruct = m_corResult.Structs.Structs(m_lngItemPtr)
                    rtbCode.SelStart = m_corResult.Sections.Sections(m_stsStruct.Section).Position - 1
                    rtbCode.Span vbCr, True, True
                    tsCode.Tabs(1).Selected = True
                Case "tsb"
                    m_lngItemPtr = Mid(m_nodNode.Parent.Tag, 5)
                    m_lngSubItemPtr = Mid(m_nodNode.Tag, 5)
                    m_stsStruct = m_corResult.Structs.Structs(m_lngItemPtr)
                    m_rimMember = m_stsStruct.Members.Items(m_lngSubItemPtr)
                    rtbCode.SelStart = m_corResult.Sections.Sections(m_rimMember.Section).Position - 1
                    rtbCode.Span vbCr, True, True
                    tsCode.Tabs(1).Selected = True
                Case "ens"
                    m_lngItemPtr = Mid(m_nodNode.Tag, 5)
                    m_ensEnum = m_corResult.Enums.Enums(m_lngItemPtr)
                    rtbCode.SelStart = m_corResult.Sections.Sections(m_ensEnum.Section).Position - 1
                    rtbCode.Span vbCr, True, True
                    tsCode.Tabs(1).Selected = True
                Case "esi"
                    m_lngItemPtr = Mid(m_nodNode.Parent.Tag, 5)
                    m_lngSubItemPtr = Mid(m_nodNode.Tag, 5)
                    m_ensEnum = m_corResult.Enums.Enums(m_lngItemPtr)
                    m_esiItem = m_ensEnum.Members.Item(m_lngSubItemPtr)
                    rtbCode.SelStart = m_corResult.Sections.Sections(m_esiItem.Section).Position - 1
                    rtbCode.Span vbCr, True, True
                    tsCode.Tabs(1).Selected = True
                Case "dec"
                    m_lngItemPtr = Mid(m_nodNode.Tag, 5)
                    m_desDeclare = m_corResult.Declares.Statements(m_lngItemPtr)
                    rtbCode.SelStart = m_corResult.Sections.Sections(m_desDeclare.Section).Position - 1
                    rtbCode.Span vbCr, True, True
                    tsCode.Tabs(1).Selected = True
                Case "dea"
                    m_lngItemPtr = Mid(m_nodNode.Parent.Tag, 5)
                    m_lngSubItemPtr = Mid(m_nodNode.Tag, 5)
                    m_desDeclare = m_corResult.Declares.Statements(m_lngItemPtr)
                    m_rimMember = m_desDeclare.Arguments.Items(m_lngSubItemPtr)
                    rtbCode.SelStart = m_corResult.Sections.Sections(m_rimMember.Section).Position - 1
                    rtbCode.Span ",)", True, True
                    tsCode.Tabs(1).Selected = True
            End Select
        End If
    End If
    m_nodNode.Expanded = True
    rtbCode.SetFocus
End Sub

Private Sub UpdateODLView()
    Dim m_lngLoop As Long
    Dim m_lngSubItem As Long
    Dim m_stsStatement As StructStatement
    Dim m_rimMember As ResultItem
    Dim m_lngBufferIndex As Long
    Dim m_staBuffer() As String
    Dim m_strAppend As String
    Dim m_lngArg As Long
    Dim m_ensEnum As EnumStatement
    Dim m_esiItem As EnumStatementItem
    Dim m_colModules As Collection
    Dim m_colModule As Collection
    Dim m_colEnums As Collection
    Dim m_desDeclares As DeclareStatement
    Dim m_strName As String
    Set m_colModules = New Collection
    Set m_colEnums = New Collection
    Dim m_lngIndex As Long
    ReDim m_staBuffer(1 To 1)
    m_lngBufferIndex = 1
    For m_lngLoop = 1 To m_corResult.Enums.Count
        Set m_colEnums = New Collection
        m_ensEnum = m_corResult.Enums.Enums(m_lngLoop)
        m_colEnums.Add m_ensEnum.Name, m_ensEnum.Name
        m_strAppend = "typedef enum tag" & m_ensEnum.Name & vbCrLf & "{" & vbCrLf
        For m_lngSubItem = 1 To m_ensEnum.Members.Count
            m_esiItem = m_ensEnum.Members.Item(m_lngSubItem)
            m_strAppend = m_strAppend & Space(4) & m_esiItem.Name & " = 0x" & Hex(m_esiItem.Value)
            If Not m_lngSubItem = m_ensEnum.Members.Count Then
                m_strAppend = m_strAppend & ", "
            End If
            m_strAppend = m_strAppend & vbCrLf
        Next '//[m_lngSubItem]
        m_strAppend = m_strAppend & "}" & m_ensEnum.Name & ";" & vbCrLf & vbCrLf
        m_staBuffer(m_lngBufferIndex) = m_staBuffer(m_lngBufferIndex) & m_strAppend
        If Len(m_staBuffer(m_lngBufferIndex)) > 5000 Then
            m_lngBufferIndex = m_lngBufferIndex + 1
            ReDim Preserve m_staBuffer(1 To m_lngBufferIndex)
        End If
    Next
    For m_lngLoop = 1 To m_corResult.Declares.Count
        m_desDeclares = m_corResult.Declares.Statements(m_lngLoop)
        If Not ItemExists(m_colModules, m_desDeclares.Library) Then
            Set m_colModule = New Collection
            m_colModule.Add m_desDeclares.Library, "Index"
            m_colModules.Add m_colModule, m_desDeclares.Library
        End If
    Next
    For m_lngLoop = 1 To m_corResult.Declares.Count
        m_desDeclares = m_corResult.Declares.Statements(m_lngLoop)
        Set m_colModule = m_colModules(m_desDeclares.Library)
        If Not m_colModule Is Nothing Then
            On Error Resume Next
            m_colModule.Add m_lngLoop, m_desDeclares.Header.Name
        End If
    Next
    For m_lngLoop = 1 To m_corResult.Structs.Count
        m_stsStatement = m_corResult.Structs.Structs(m_lngLoop)
        m_strAppend = "typedef struct tag" & m_stsStatement.Name & vbCrLf & "{" & vbCrLf
        For m_lngSubItem = 1 To m_stsStatement.Members.Count
            m_rimMember = m_stsStatement.Members.Items(m_lngSubItem)
            If m_rimMember.bData Then
                m_strAppend = m_strAppend & Space(4) & "SAFEARRAY" & m_gl_CST_strOprLeftBracket & GetTypeConversion(m_rimMember.Type) & m_gl_CST_strOprRightBracket & m_gl_CST_strOprSpace & m_rimMember.Name & ";" & vbCrLf
            Else
                m_strAppend = m_strAppend & Space(4) & GetTypeConversion(m_rimMember.Type) & m_gl_CST_strOprSpace & m_rimMember.Name & ";" & vbCrLf
            End If
        Next
        m_strAppend = m_strAppend & "}" & m_stsStatement.Name & ";" & vbCrLf & vbCrLf
        m_staBuffer(m_lngBufferIndex) = m_staBuffer(m_lngBufferIndex) & m_strAppend
        If Len(m_staBuffer(m_lngBufferIndex)) > 5000 Then
            m_lngBufferIndex = m_lngBufferIndex + 1
            ReDim Preserve m_staBuffer(1 To m_lngBufferIndex)
        End If
    Next
    For m_lngLoop = 1 To m_colModules.Count
        Set m_colModule = m_colModules(m_lngLoop)
        m_strName = m_colModule.Item("Index")
        m_colModule.Remove "Index"
        m_strAppend = "[dllname(""" & m_strName & """)]" & vbCrLf
        m_strAppend = m_strAppend & "module " & m_strName & vbCrLf
        m_strAppend = m_strAppend & "{" & vbCrLf
        For m_lngSubItem = 1 To m_colModule.Count
            m_lngIndex = m_colModule.Item(m_lngSubItem)
            m_desDeclares = m_corResult.Declares.Statements(m_lngIndex)
            If m_desDeclares.Alias = vbNullString Then
                m_desDeclares.Alias = m_desDeclares.Header.Name
            End If
            m_strAppend = m_strAppend & "[entry(""" & m_desDeclares.Alias & """), helpstring(""" & m_desDeclares.Alias & """)]" & " " & GetDeclConversion(m_desDeclares.Header.Type, ItemExists(m_colEnums, m_desDeclares.Header.Type)) & " " & _
                m_desDeclares.Header.Name & "("
            For m_lngArg = 1 To m_desDeclares.Arguments.Count
                m_rimMember = m_desDeclares.Arguments.Items(m_lngArg)
                If m_rimMember.ByVal Then
                    m_strAppend = m_strAppend & "[in] "
                    m_strAppend = m_strAppend & GetDeclConversion(m_rimMember.Type, ItemExists(m_colEnums, m_rimMember.Type)) & " " & m_rimMember.Name
                Else
                    m_strAppend = m_strAppend & "[in, out] "
                    m_strAppend = m_strAppend & GetDeclConversion(m_rimMember.Type, ItemExists(m_colEnums, m_rimMember.Type)) & "* " & m_rimMember.Name
                End If
                If Not m_lngArg = m_desDeclares.Arguments.Count Then
                    m_strAppend = m_strAppend & ", "
                End If
            Next
                m_strAppend = m_strAppend & ");" & vbCrLf
        Next
        m_strAppend = m_strAppend & "};" & vbCrLf & vbCrLf
        m_staBuffer(m_lngBufferIndex) = m_staBuffer(m_lngBufferIndex) & m_strAppend
        If Len(m_staBuffer(m_lngBufferIndex)) > 5000 Then
            m_lngBufferIndex = m_lngBufferIndex + 1
            ReDim Preserve m_staBuffer(1 To m_lngBufferIndex)
        End If
    Next
    rtbODL.Font.Name = "Courier New"
    rtbODL.TextRTF = MakeCRTF(Join(m_staBuffer, vbNullString))
End Sub
