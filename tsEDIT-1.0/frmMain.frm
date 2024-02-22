VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   Caption         =   "New Notepad"
   ClientHeight    =   4725
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   8880
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4725
   ScaleWidth      =   8880
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog CD1 
      Left            =   1920
      Top             =   3840
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin tsEDIT.Gtxt Gtxt1 
      Height          =   3495
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   6165
      BackColor       =   16777215
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Text            =   ""
      MouseIcon       =   "frmMain.frx":014A
      Title           =   ""
      NumBackColor    =   14737632
      NumForeColor    =   8388608
      TitleForeColor  =   8388608
      NumbarVisible   =   0   'False
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileNew 
         Caption         =   "&New"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuFileOpen 
         Caption         =   "&Open"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuFileSave 
         Caption         =   "&Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuFileSaveAs 
         Caption         =   "Save &As"
         Shortcut        =   ^{F12}
      End
      Begin VB.Menu Sep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFilePageSetup 
         Caption         =   "Page Se&tup"
      End
      Begin VB.Menu mnuFilePrint 
         Caption         =   "&Print"
         Shortcut        =   ^P
      End
      Begin VB.Menu Sep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuA 
         Caption         =   "Associations"
      End
      Begin VB.Menu sep6 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuEditUndo 
         Caption         =   "&Undo"
         Enabled         =   0   'False
      End
      Begin VB.Menu Sep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditCut 
         Caption         =   "Cu&t"
      End
      Begin VB.Menu mnuEditCopy 
         Caption         =   "&Copy"
      End
      Begin VB.Menu mnuEditPaste 
         Caption         =   "&Paste"
      End
      Begin VB.Menu mnuEditDelete 
         Caption         =   "De&lete"
      End
      Begin VB.Menu Sep4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditSelectAll 
         Caption         =   "Select &All"
      End
      Begin VB.Menu mnuEditTimeDate 
         Caption         =   "Time/&Date"
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnuEditWordWrap 
         Caption         =   "&Word Wrap"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuEditSetFont 
         Caption         =   "Set &Font"
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      Begin VB.Menu mnuViewNumBar 
         Caption         =   "View &Number Bar"
         Checked         =   -1  'True
         Shortcut        =   ^B
      End
   End
   Begin VB.Menu mnuSearch 
      Caption         =   "&Search"
      Begin VB.Menu mnuSearchFind 
         Caption         =   "&Find"
      End
      Begin VB.Menu mnuSearchFindNext 
         Caption         =   "Find &Next"
         Shortcut        =   {F3}
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Visible         =   0   'False
      Begin VB.Menu mnuHelpHelpTopics 
         Caption         =   "&Help Topics"
         Enabled         =   0   'False
      End
      Begin VB.Menu Sep5 
         Caption         =   "-"
      End
   End
   Begin VB.Menu mnuHelpAbout 
      Caption         =   "&About"
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim SaveFileName As String
Dim Saved As Boolean
Dim Loading As Boolean

Private Sub Form_Load()
    Gtxt1.Title = "Untitled"
    Gtxt1.Numbar_Visible = True
    Load frmFind
    modSysMenu.LoadSysMenu Me.Hwnd
    Dim Cmd As String
    Cmd = Command$
    If Not Cmd = "" Then
        If Dir(Cmd) <> "" Then
            OpenDocument Cmd
        End If
    End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If HasSaved Then
        AppEnd = True
        Unload frmFind
    Else
        Cancel = True
    End If
End Sub

Private Sub Form_Resize()
    Gtxt1.Move 0, 0, ScaleWidth, ScaleHeight
End Sub

Private Sub Form_Unload(Cancel As Integer)
    modSysMenu.UnloadSysMenu Me.Hwnd
End Sub

Private Sub Gtxt1_BeforeChangeTitle(Cancel As Boolean)
    If Saved Then Cancel = True
End Sub

Private Sub Gtxt1_Change()
    If Not Loading Then Saved = False
End Sub

Private Sub Gtxt1_ChangeTitle(NewTitle As String)
    Me.Caption = NewTitle & " - tsEDIT"
End Sub

Private Sub mnuA_Click()
    frmA.Show 1
End Sub

Private Sub mnuEdit_Click()
    If Gtxt1.SelText = "" Then
        Me.mnuEditCut.Enabled = False
        Me.mnuEditCopy.Enabled = False
        Me.mnuEditDelete.Enabled = False
    Else
        Me.mnuEditCut.Enabled = True
        Me.mnuEditCopy.Enabled = True
        Me.mnuEditDelete.Enabled = True
    End If
    If Clipboard.GetText() = "" Then
        Me.mnuEditPaste.Enabled = False
    Else
        Me.mnuEditPaste.Enabled = True
    End If
    If Gtxt1.Text = "" Then
        Me.mnuEditSelectAll.Enabled = False
    Else
        Me.mnuEditSelectAll.Enabled = True
    End If
End Sub

Private Sub mnuEditCopy_Click()
    Gtxt1.txtCopy
End Sub

Private Sub mnuEditCut_Click()
    Gtxt1.txtCut
End Sub

Private Sub mnuEditDelete_Click()
    Gtxt1.txtDelete
End Sub

Private Sub mnuEditPaste_Click()
    Gtxt1.txtPaste
End Sub

Private Sub mnuEditSelectAll_Click()
    Gtxt1.txtSelect
End Sub

Private Sub mnuEditSetFont_Click()
    On Error GoTo ErrorHandler
    CD1.Flags = cdlCFScreenFonts + cdlCFEffects
    With Gtxt1
        CD1.FontName = .Font.Name
        CD1.FontSize = .Font.Size
        CD1.FontBold = .Font.Bold
        CD1.FontItalic = .Font.Italic
        CD1.FontUnderline = .Font.Underline
        CD1.FontStrikethru = .Font.Strikethrough
        CD1.Color = 0
    End With
    CD1.ShowFont
    With Gtxt1
        .Font.Name = CD1.FontName
        .Font.Size = CD1.FontSize
        .Font.Bold = CD1.FontBold
        .Font.Italic = CD1.FontItalic
        .Font.Underline = CD1.FontUnderline
        .Font.Strikethrough = CD1.FontStrikethru
    End With
ErrorHandler:
End Sub

Private Function HasSaved() As Boolean
    Dim Result As VbMsgBoxResult
    If Not Saved And Gtxt1.Text <> "" Then
        Result = MsgBox("Do you want to save this document before closing it?", vbQuestion + vbYesNoCancel, "Save?")
        If Result = vbYes Then
            SaveDocument
            HasSaved = True
        ElseIf Result = vbNo Then
            HasSaved = True
        Else
            HasSaved = False
        End If
    Else
        HasSaved = True
    End If
End Function

Private Sub SaveDocument()
    If Not Saved And SaveFileName <> "" Then
        Close #1
        Open SaveFileName For Output As #1
            Print #1, , Gtxt1.Text
        Close #1
        Saved = True
    Else
        SaveDocumentAs
    End If
End Sub

Private Sub SaveDocumentAs()
    Dim EXT As String
    On Error GoTo ErrorHandler
    CD1.Filter = "Text Document (*.txt)|*.txt;|All Files (*.*)|*.*;"
    CD1.Flags = cdlOFNExplorer + cdlOFNOverwritePrompt ' + cdlOFNNoReadOnlyReturn
    CD1.FileName = Gtxt1.Title
    CD1.ShowSave
    'If CD1.FilterIndex = 1 And LCase(Right(CD1.FileTitle, 4)) <> ".txt" Then EXT = ".txt"
    SaveFileName = CD1.FileName
    Gtxt1.Title = CD1.FileTitle
    Saved = False
    SaveDocument
ErrorHandler:
End Sub

Private Sub LoadNewDocument()
    If HasSaved Then
        Loading = True
        Gtxt1.Text = ""
        Gtxt1.Title = "Untitled"
        Saved = False
        SaveFileName = ""
        Loading = False
    End If
End Sub

Private Sub OpenDocument(Optional ByVal FileName As String)
    If HasSaved Then
        Dim FileTitle As String
        On Error GoTo ErrorHandler
        Loading = True
        If FileName = "" Then
            CD1.Filter = "All Supported File Types|*.txt;*.htm;*.html;*.frm;*.js;*.nfo|Text Document (*.txt)|*.txt|All Files (*.*)|*.*"
            CD1.Flags = cdlOFNExplorer + cdlOFNFileMustExist
            CD1.ShowOpen
            FileName = CD1.FileName
            FileTitle = CD1.FileTitle
        Else
            FileTitle = getfiletitle(FileName)
        End If
        SaveFileName = FileName
        Saved = True
        Gtxt1.Title = FileTitle
        Close #1
        Dim stg As String, lne As String
        Open SaveFileName For Input As #1
            Do Until EOF(1)
                Line Input #1, lne
                stg = stg & lne
                If Not EOF(1) Then stg = stg & vbNewLine
                DoEvents
            Loop
        Close #1
        Gtxt1.Text = stg
    End If
ErrorHandler:
    Loading = False
End Sub

Private Function getfiletitle(ByVal FileName As String) As String
    Dim stg As String, Pos As Integer, Posb As Integer
    Pos = InStr(1, FileName, "\")
    Do Until Pos = 0
        Posb = Pos
        Pos = InStr(Pos + 1, FileName, "\")
    Loop
    getfiletitle = Mid(FileName, Posb + 1)
End Function

Private Sub mnuEditTimeDate_Click()
    Gtxt1.SelText = Time & " " & Date
End Sub

Private Sub mnuEditWordWrap_Click()
    If mnuEditWordWrap.Checked = True Then
        mnuEditWordWrap.Checked = False
        Gtxt1.WordWrap True
    Else
        mnuEditWordWrap.Checked = True
        Gtxt1.WordWrap False
    End If
End Sub

Private Sub mnuFileExit_Click()
    Unload Me
End Sub

Private Sub mnuFileNew_Click()
    LoadNewDocument
End Sub

Private Sub mnuFileOpen_Click()
    OpenDocument
End Sub

Private Sub mnuFilePageSetup_Click()
    On Error GoTo ErrorHandler
    CD1.Flags = cdlPDPrintSetup
    CD1.ShowPrinter
ErrorHandler:
End Sub

Private Sub mnuFileSave_Click()
    SaveDocument
End Sub

Private Sub mnuFileSaveAs_Click()
    SaveDocumentAs
End Sub

Private Sub mnuHelpAbout_Click()
    MsgBox "tsEDIT 1.1", vbInformation, "About"
End Sub

Private Sub mnuSearchFind_Click()
    frmFind.Show , Me
    frmFind.Text1.Text = Gtxt1.SelText
    frmFind.Text1.SelStart = 0
    frmFind.Text1.SelLength = Len(Gtxt1.SelText)
End Sub

Private Sub mnuSearchFindNext_Click()
    If frmFind.Text1.Text = "" And Gtxt1.SelText = "" Then
        mnuSearchFind_Click
    ElseIf frmFind.Text1.Text = "" Then
        frmFind.Text1.Text = Gtxt1.SelText
        frmFind.Find
    Else
        frmFind.Find
    End If
End Sub

Private Sub mnuViewNumBar_Click()
    Gtxt1.Numbar_Visible = Not Gtxt1.Numbar_Visible
    mnuViewNumBar.Checked = Not mnuViewNumBar.Checked
End Sub
