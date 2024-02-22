VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.UserControl Gtxt 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   ToolboxBitmap   =   "GTxtBox.ctx":0000
   Begin VB.PictureBox picSep 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   1095
      Left            =   600
      MouseIcon       =   "GTxtBox.ctx":0312
      ScaleHeight     =   1095
      ScaleWidth      =   135
      TabIndex        =   5
      Top             =   1800
      Width           =   135
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Left            =   1440
      Top             =   3000
   End
   Begin VB.PictureBox TTLBack 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      FillColor       =   &H00808080&
      Height          =   495
      Left            =   600
      ScaleHeight     =   495
      ScaleWidth      =   3015
      TabIndex        =   3
      Top             =   120
      Width           =   3015
      Begin VB.TextBox txtTitle 
         BorderStyle     =   0  'None
         Height          =   195
         Left            =   240
         MaxLength       =   255
         TabIndex        =   4
         Text            =   "Text2"
         Top             =   120
         Visible         =   0   'False
         Width           =   255
      End
   End
   Begin VB.CommandButton Command1 
      Enabled         =   0   'False
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   4095
   End
   Begin VB.PictureBox NumBar 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   2655
      Left            =   0
      ScaleHeight     =   2655
      ScaleWidth      =   495
      TabIndex        =   2
      Top             =   240
      Width           =   495
   End
   Begin RichTextLib.RichTextBox Text1 
      Height          =   2175
      Left            =   1080
      TabIndex        =   0
      Top             =   360
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   3836
      _Version        =   393217
      BorderStyle     =   0
      HideSelection   =   0   'False
      ScrollBars      =   3
      Appearance      =   0
      TextRTF         =   $"GTxtBox.ctx":061C
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Menu AVM 
      Caption         =   "AVM"
      Begin VB.Menu AVMundo 
         Caption         =   "Undo"
      End
      Begin VB.Menu Sep1 
         Caption         =   "-"
      End
      Begin VB.Menu AVMcut 
         Caption         =   "Cut"
      End
      Begin VB.Menu AVMcopy 
         Caption         =   "Copy"
      End
      Begin VB.Menu AVMpaste 
         Caption         =   "Paste"
      End
      Begin VB.Menu AVMdelete 
         Caption         =   "Delete"
      End
      Begin VB.Menu sep2 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu AVMfont 
         Caption         =   "Font"
         Visible         =   0   'False
      End
      Begin VB.Menu sep3 
         Caption         =   "-"
      End
      Begin VB.Menu AVMselectall 
         Caption         =   "Select All"
      End
      Begin VB.Menu sep4 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuEasyRead 
         Caption         =   "Easy Read"
         Visible         =   0   'False
      End
   End
End
Attribute VB_Name = "Gtxt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Declare Function apiLockWindowUpdate Lib "user32" Alias "LockWindowUpdate" (ByVal hWndLock As Long) As Long
Private Declare Function apiSendMessage Lib "user32" Alias "SendMessageA" (ByVal Hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Private Const WM_USER = &H400
Private Const EM_GETTEXTRANGE = (WM_USER + 75)
Private Const EM_GETLINECOUNT = &HBA
Private Const EM_GETFIRSTVISIBLELINE = &HCE
Private Const EM_GETSEL = &HB0
Private Const EM_LINEFROMCHAR = &HC9
Private Const EM_LINEINDEX = &HBB
Private Const EM_LINELENGTH = &HC1
Private Const EM_UNDO = &HC7
Private Declare Function GetCaretPos Lib "user32" (lpPoint As PointAPI) As Long
Private Declare Function SetCaretPos Lib "user32" (ByVal X As Long, ByVal Y As Long) As Long

Private Type PointAPI
    X As Long
    Y As Long
End Type

Public Enum BDRstyle
    GTB_NoBorder = 0
    GTB_FixedSingle = 1
End Enum

Private Declare Function ReleaseCapture Lib "user32" () As Long

Private MyTitle As String
Private AVMenabled As Boolean
Private TBchanged As Boolean
Private MyFileName As String
Private SStart As Long
Private OldLine As Long

Dim tbMouseX As Single
Dim tbMouseY As Single

Event Click() 'MappingInfo=UserControl,UserControl,-1,Click
Event DblClick() 'MappingInfo=UserControl,UserControl,-1,DblClick
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseDown
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseMove
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseUp
Event Change()
Event KeyDown(KeyCode As Integer, Shift As Integer)
Event KeyUp(KeyCode As Integer, Shift As Integer)
Event KeyPress(KeyAscii As Integer)
Event ChangeTitle(NewTitle As String)
Event BeforeChangeTitle(Cancel As Boolean)

Const Title_Height = 255

Dim PrintedTopNumber As Long

Public SyntaxColoring As Boolean

'Public Sub GetSyntax()
'doGetScriptKeywords
'End Sub

Private Sub ResizeControls()
Dim EdgeSize As Integer, NBWid As Long

If UserControl.Width < 1000 Then Exit Sub
If UserControl.Height < 1000 Then Exit Sub

EdgeSize = (Width - ScaleWidth) / 2
UserControl.Font = Text1.Font
TTLBack.Font = Text1.Font
Command1.Move 0, 0, ScaleWidth, UserControl.TextHeight("ABC") + (EdgeSize * 4)
If NumBar.Visible Then
    NumBar.Move 0, Title_Height, NumBar.Width, ScaleHeight - Title_Height
    NBWid = NumBar.Width
Else
    NBWid = 0
End If
picSep.Move NBWid, Title_Height, 30, ScaleHeight - Title_Height
Text1.Move NBWid + picSep.Width, Title_Height, ScaleWidth - (NBWid + picSep.Width), ScaleHeight - Title_Height
TTLBack.Move 0, 0, ScaleWidth, Title_Height
'PrintTitle
PrintLineCount
If NBWid <> 0 Then PrintNums
End Sub

'Sub colorize()

'apiLockWindowUpdate Text1.Text

'when finished

'apiLockWindowUpdate 0

'End Sub

Private Sub AVMcopy_Click()
    Clipboard.Clear
    Clipboard.SetText Text1.SelText
    Text1.SetFocus
End Sub

Private Sub AVMcut_Click()
    Clipboard.Clear
    Clipboard.SetText Text1.SelText
    Text1.SelText = ""
    Text1.SetFocus
End Sub

Private Sub AVMdelete_Click()
    Text1.SelText = ""
    Text1.SetFocus
End Sub

Private Sub AVMfont_Click()
On Error Resume Next
MyFont.FontBold = IIf(Text1.SelBold, True, False)
MyFont.FontColor = Text1.SelColor
MyFont.FontItalic = IIf(Text1.SelItalic, True, False)
MyFont.FontName = Text1.SelFontName
MyFont.FontSize = Text1.SelFontSize
MyFont.FontStrikethru = IIf(Text1.SelStrikeThru, True, False)
MyFont.FontUnderline = IIf(Text1.SelUnderline, True, False)
frmFont.Show 1
If frmFont.Canceled = False Then
    Text1.SelBold = MyFont.FontBold
    Text1.SelColor = MyFont.FontColor
    Text1.SelItalic = MyFont.FontItalic
    Text1.SelFontName = MyFont.FontName
    Text1.SelFontSize = MyFont.FontSize
    Text1.SelStrikeThru = MyFont.FontStrikethru
    Text1.SelUnderline = MyFont.FontUnderline
End If
End Sub

Private Sub AVMpaste_Click()
    If Clipboard.GetFormat(vbCFRTF) Then
        Text1.SelText = Clipboard.GetText(vbCFRTF)
    ElseIf Clipboard.GetFormat(vbCFText) Then
        Text1.SelText = Clipboard.GetText(vbCFRTF)
    End If
End Sub

Private Sub AVMselectall_Click()
    Text1.SelStart = 0
    Text1.SelLength = Len(Text1.Text)
    Text1.SetFocus
End Sub

Private Sub AVMundo_Click()
    Text1.SetFocus
    DoEvents
    SendKeys "^z"
End Sub

Private Sub mnuEasyRead_Click()
Dim i As Long, j As Long
If mnuEasyRead.Checked = False Then
    Text1.BackColor = &H800000
    picSep.BackColor = &H800000
    i = Text1.SelStart
    j = Text1.SelLength
    Text1.SelStart = 0
    Text1.SelLength = Len(Text1.Text)
    Text1.SelColor = vbWhite
    Text1.SelStart = i
    Text1.SelLength = j
    mnuEasyRead.Checked = True
Else
    Text1.BackColor = vbWhite
    picSep.BackColor = vbWhite
    i = Text1.SelStart
    j = Text1.SelLength
    Text1.SelStart = 0
    Text1.SelLength = Len(Text1.Text)
    Text1.SelColor = vbBlack
    Text1.SelStart = i
    Text1.SelLength = j
    mnuEasyRead.Checked = False
End If
End Sub

Private Sub NumBar_Click()
Text1.SetFocus
End Sub

Private Sub NumBar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Text1.SetFocus
End Sub

Private Sub NumBar_Resize()
ResizeControls
End Sub

Private Sub picSep_Click()
    Text1.SetFocus
End Sub

Private Sub Text1_Change()
Dim S1 As Single, S2 As Single
TBchanged = True
If Not Text1.SelFontSize = Text1.Font.Size Or Not Text1.SelFontName = Text1.Font.Name Then
'    LockWindowUpdate Text1.Hwnd
    S1 = Text1.SelStart
    S2 = Text1.SelLength
    Text1.SelStart = 0
    Text1.SelLength = Len(Text1.Text)
    Text1.SelFontSize = Text1.Font.Size
    Text1.SelFontName = Text1.Font.Name
    Text1.SelStart = S1
    Text1.SelLength = S2
'    LockWindowUpdate 0
End If
If GetTopLineNumber <> PrintedTopNumber Then PrintNums
PrintLineCount
RaiseEvent Change
End Sub

Private Sub Text1_Click()
'PrintLineCount
RaiseEvent Click
'Text1.SelLength = 0
End Sub

Private Sub Text1_DblClick()
Text1.SelLength = 0
RaiseEvent DblClick
End Sub

Private Sub Text1_GotFocus()
    PrintLineCount
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
'PrintLineCount
If KeyCode = 9 And Shift = 0 And Text1.SelText = "" Then
    Text1.SelText = vbTab
    KeyCode = 0
End If
If KeyCode = 67 And Shift = 2 Then
    Exit Sub
End If
If GetLineNumber <> OldLine And OldLine <> 0 Then PrintNums
OldLine = GetLineNumber
RaiseEvent KeyDown(KeyCode, Shift)
'If KeyCode = 46 Or KeyCode = 8 Then
'    PrintNums
'End If
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
'Dim stg As String
'PrintLineCount
RaiseEvent KeyPress(KeyAscii)
'If KeyAscii = 13 Then
'    stg = GetLineOfText
'    If LCase(Left(stg, 11)) = "private sub" Then
'        If InStr(11, stg, "()") = 0 Then
'
'        End If
'    ElseIf LCase(Left(stg, 16)) = "private function" Then
'
'    End If
'End If
End Sub

Private Sub Text1_KeyUp(KeyCode As Integer, Shift As Integer)
'    Dim CommentColor As Long
'    Dim StringColor As Long
'    Dim KeysColor As Long
'
'If SyntaxColoring = True Then
'    LockWindowUpdate UserControl.hwnd
'
'    Select Case KeyCode
'
'    Case 13, 32 ', 38, 40, 37, 39
'        'CommentColor = RGB(0, 128, 0)       '// DARK GREEN  //
'        'StringColor = RGB(0, 0, 0)          '// BLACK       //
'        'KeysColor = RGB(0, 0, 128)          '// DARK BLUE   //
'        KeysColor = &H800000
'        StringColor = vbBlack
'        CommentColor = &H8000&
'
'        Colorize Text1, CommentColor, StringColor, KeysColor, KeyCode
'
'            Text1.SelColor = StringColor
'
'    End Select
'
'    LockWindowUpdate 0&
'End If
'###########
'###########
'###########
PrintLineCount
'If GetLineNumber <> OldLine And OldLine <> 0 Then PrintNums
'OldLine = GetLineNumber
RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub Text1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
SStart = Text1.SelStart
RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub Text1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub Text1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Text1.SelStart <> SStart Then PrintLineCount: PrintNums

If AVMenabled = True And Button = 2 Then
    If Text1.SelText = "" Then
        AVMcut.Enabled = False
        AVMcopy.Enabled = False
        AVMdelete.Enabled = False
    Else
        AVMcut.Enabled = True
        AVMcopy.Enabled = True
        AVMdelete.Enabled = True
    End If
    If Clipboard.GetText = "" Then
        AVMpaste.Enabled = False
    Else
        AVMpaste.Enabled = True
    End If
    If Text1.Text = "" Then
        AVMselectall.Enabled = False
    Else
        AVMselectall.Enabled = True
    End If
    If TBchanged = False Then
        AVMundo.Enabled = False
    Else
        AVMundo.Enabled = True
    End If
    PopupMenu AVM
End If
RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

Private Sub Text1_SelChange()
    If GetTopLineNumber <> PrintedTopNumber Then PrintNums
End Sub

'Private Sub Text1_OLEDragDrop(Data As RichTextLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
''    If Effect = vbDropEffectMove Then
'        If Data.GetFormat(vbCFFiles) = True Then
'
'        ElseIf Data.GetFormat(vbCFText) = True Then
'            SetCaretPos x, y
'            Text1.SelText = Data.GetData(vbCFText)
'        End If
''    End If
'End Sub
'
'Private Sub Text1_OLEDragOver(Data As RichTextLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single, State As Integer)
'    If Data.GetFormat(vbCFFiles) = True Then
'        Effect = vbDropEffectMove
'    ElseIf Data.GetFormat(vbCFText) = True Then
'        Effect = vbDropEffectMove
'        SetCaretPos x / Screen.TwipsPerPixelX, y / Screen.TwipsPerPixelY
'    Else
'        Effect = vbDropEffectNone
'    End If
'End Sub

Private Sub Timer1_Timer()
    If NumBar.Visible Then If GetTopLineNumber <> PrintedTopNumber Then PrintNums
End Sub

Private Sub TTLBack_Click()
Text1.SetFocus
End Sub

Private Sub TTLBack_DblClick()
Dim SI As Boolean, TTLw As Long
SI = False
TTLw = TTLBack.TextWidth(MyTitle)
If TTLw < TTLBack.TextWidth("  ") Then TTLw = TTLBack.TextWidth("  ")
Text1.SetFocus
If tbMouseX > 50 And tbMouseX < 50 + TTLw Then
    If tbMouseY > 0 And tbMouseY < TTLBack.ScaleHeight Then
        RaiseEvent BeforeChangeTitle(SI)
        If SI = True Then Exit Sub
        Dim Txt_1 As String
        Txt_1 = MyTitle
        txtTitle.Font = TTLBack.Font
        txtTitle.FontSize = TTLBack.FontSize
        txtTitle.Text = MyTitle
        txtTitle.Height = TTLBack.ScaleHeight - 15
        wID = TTLBack.TextWidth(MyTitle & "  ")
        If wID > 3465 Then
            txtTitle.Width = 3465
        Else
            txtTitle.Width = wID
        End If
        txtTitle.Move 50, 0 '(TTLBack.ScaleHeight - txtTitle.Height) / 2
        txtTitle.BackColor = TTLBack.BackColor
        txtTitle.ForeColor = TTLBack.ForeColor
        MyTitle = ""
        PrintLineCount
        MyTitle = Txt_1
        txtTitle.Visible = True
        txtTitle.SetFocus
        txtTitle.SelStart = 0
        txtTitle.SelLength = Len(txtTitle.Text)
    End If
End If
End Sub

Private Sub TTLBack_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Dim SI As Boolean, TTLw As Long
'SI = False
'
'TTLw = TTLBack.TextWidth(MyTitle)
'
'If TTLw < TTLBack.TextWidth("  ") Then TTLw = TTLBack.TextWidth("  ")
'
'Text1.SetFocus
'If x > 50 And x < 50 + TTLw Then
'    If y > 0 And y < TTLBack.ScaleHeight Then
'        If Timer1.Enabled = False Then
'            Timer1.Interval = 10
'            Timer1.Enabled = True
'        Else
'            RaiseEvent BeforeChangeTitle(SI)
'            If SI = True Then Exit Sub
'            txtTitle.Font = TTLBack.Font
'            txtTitle.FontSize = TTLBack.FontSize
'            txtTitle.Text = MyTitle
'            txtTitle.Height = 50
'            wid = TTLBack.TextWidth(MyTitle & "  ")
'            If wid > 3465 Then
'                txtTitle.Width = 3465
'            Else
'                txtTitle.Width = wid
'            End If
'            txtTitle.Move 50, (TTLBack.ScaleHeight - txtTitle.Height) / 2
'            txtTitle.Visible = True
'            txtTitle.SetFocus
'            txtTitle.SelStart = 0
'            txtTitle.SelLength = Len(txtTitle.Text)
'        End If
'    End If
'End If
End Sub

Private Sub TTLBack_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    tbMouseX = X: tbMouseY = Y
End Sub

Private Sub txtTitle_Change()
Dim posa As Integer
Dim Posb As Integer, wID As Integer
'posa = txtTitle.SelStart
'Posb = txtTitle.SelLength
'txtTitle.SelStart = 0
'txtTitle.SelStart = posa
'txtTitle.SelLength = pob
wID = TTLBack.TextWidth(txtTitle.Text & "   ")
If wID < 3465 Then
    txtTitle.Width = wID
End If
End Sub

Private Sub txtTitle_KeyDown(KeyCode As Integer, Shift As Integer)
'    txtTitle_Change
End Sub

Private Sub txtTitle_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtTitle.Visible = False
    MyTitle = txtTitle.Text
    PrintLineCount
    RaiseEvent ChangeTitle(MyTitle)
ElseIf KeyAscii = 27 Then
    txtTitle.Visible = False
    PrintLineCount
End If
End Sub

Private Sub txtTitle_LostFocus()
txtTitle.Visible = False
MyTitle = txtTitle.Text
PrintLineCount
RaiseEvent ChangeTitle(MyTitle)
End Sub


Private Sub UserControl_Initialize()
Text1.Text = ""
TBchanged = False
End Sub

Private Sub UserControl_Resize()
ResizeControls
PrintLineCount
PrintNums
End Sub

Private Sub DrawTitleBorder()
    TTLBack.Cls
    TTLBack.Line (0, TTLBack.ScaleHeight - 15)-(TTLBack.ScaleWidth, TTLBack.ScaleHeight - 15), &H808080
End Sub

Private Sub PrintTitle()
    Dim wID As Integer
    TTLBack.CurrentX = 50
    TTLBack.CurrentY = (TTLBack.ScaleHeight - TTLBack.TextHeight("ABC")) / 2
    wID = TTLBack.TextWidth(MyTitle & "  ")
    If wID > 3465 Then
        TTLBack.Print Left(MyTitle, 30) & "..."
    Else
        TTLBack.Print MyTitle
    End If
End Sub

Private Sub PrintLineCount()
  Dim LineCount As Long
  Dim LineNumber As Long

    LineCount = apiSendMessage(Text1.Hwnd, EM_GETLINECOUNT, 0&, 0&)
    
    DrawTitleBorder
    PrintTitle
    TTLBack.CurrentX = TTLBack.ScaleWidth - (TTLBack.TextWidth("Line count: " & CStr(LineCount)) + 100)
    TTLBack.CurrentY = (TTLBack.ScaleHeight - TTLBack.TextHeight("ABC")) / 2
    TTLBack.Print "Line count: " & LineCount

End Sub

Private Function GetLineNumber() As Long
Dim CaretPos As Long
Dim Txt As String, posa As Long, Posb As Long

    DoEvents
    If InStr(1, Text1.Text, vbCr) = 0 Then GetLineNumber = 1: Exit Function
    CaretPos = Text1.SelStart
    DoEvents
    Txt = Left(Text1.Text, CaretPos)
    If InStr(1, Txt, vbCr) = 0 Then GetLineNumber = 1: Exit Function
    
    posa = 1
    Posb = 1
    
    Do While posa <> 0 And posa <> Len(Txt)
        If posa = 1 Then
            posa = InStr(posa, Txt, vbCr)
            If posa = 1 Then posa = posa + 1
        Else
            posa = InStr(posa + 1, Txt, vbCr)
        End If
        If posa <> 0 Then
            Posb = Posb + 1
        End If
    Loop
    
    GetLineNumber = Posb
End Function

Private Sub PrintNumsold()
Dim TopNumber As Long
Dim Numbers As Integer, CNum As Long

NumBar.Font = Text1.Font

NumBar.FontSize = Text1.Font.Size

TopNumber = apiSendMessage(Text1.Hwnd, EM_GETFIRSTVISIBLELINE, 0&, 0&)
TopNumber = TopNumber + 1

Numbers = NumBar.ScaleHeight / (NumBar.TextHeight("ABC") + 0.545)

NumBar.Cls

NumBar.Width = 200 + NumBar.TextWidth(CStr(TopNumber + Numbers))

CNum = 0

For i = 1 To Numbers
    NumBar.CurrentX = 40
    NumBar.CurrentY = CNum
    NumBar.Print TopNumber
    TopNumber = TopNumber + 1
    CNum = CNum + NumBar.ScaleHeight / Numbers
Next i

End Sub

Private Sub PrintNums(Optional Override As Boolean)
Dim TopNumber As Long
Dim Numbers As Integer, CNum As Long
Dim oldsize As Currency
Dim txtHgt As Integer

    NumBar.Font = Text1.Font
    NumBar.FontSize = Text1.Font.Size
    
    TopNumber = GetTopLineNumber
    
    PrintedTopNumber = TopNumber
    
    If NumBar.Visible Then
        Numbers = Text1.Height / NumBar.TextHeight("ABC")
        NumBar.Line (0, 0)-(NumBar.ScaleWidth, NumBar.ScaleHeight), NumBar.BackColor, BF
        NumBar.Width = 210 + (NumBar.TextWidth(CStr(TopNumber + Numbers)) * 1.1)
        
        CNum = 0
        txtHgt = NumBar.TextHeight("ABC")
        
        For i = 1 To Numbers
            NumBar.CurrentX = 0
            NumBar.CurrentY = CNum
            NumBar.Print TopNumber
            TopNumber = TopNumber + 1
            CNum = CNum + txtHgt
        Next i
        NumBar.Refresh
    End If

End Sub

Private Function GetTopLineNumber() As Long
    Dim TopNumber As Long
    TopNumber = apiSendMessage(Text1.Hwnd, EM_GETFIRSTVISIBLELINE, 0&, 0&)
    GetTopLineNumber = TopNumber + 1
End Function

'Sub SelTextColor(Color As OLE_COLOR)
'Text1.SelColor = Color
'End Sub

Sub LoadFile(File As String)
Call Text1.LoadFile(File, rtfText)
End Sub

Sub SaveFile(File As String, Optional UseTitleAsFilename As Boolean = True)
Dim fFile As String
fFile = File
If UseTitleAsFilename = True Then
    If Right(fFile, 1) <> "\" Then fFile = fFile & "\"
    fFile = fFile & Title
    If Right(LCase(fFile), 4) <> ".txt" Then fFile = fFile & ".txt"
    Call Text1.SaveFile(fFile)
Else
    Call Text1.SaveFile(File, rtfText)
End If
End Sub




'--- Properties -------------------------------------------------------------

Public Property Let SelTextColour(ByVal Colour As OLE_COLOR)
    Text1.SelColor = Colour
    'PropertyChanged "SelTextColour"
End Property

Public Property Let SelTextBold(ByVal New_Bold As Boolean)
    Text1.SelBold = New_Bold
End Property

Public Property Let SelTextItalic(ByVal New_Italic As Boolean)
    Text1.SelItalic = New_Italic
End Property

Public Property Let SelTextUnderline(ByVal New_Under As Boolean)
    Text1.SelUnderline = New_Under
End Property

Public Property Let SelTextStrikeThru(ByVal New_Strike As Boolean)
    Text1.SelStrikeThru = New_Strike
End Property

Public Property Get SelText() As String
    SelText = Text1.SelText
End Property

Public Property Let SelText(ByVal New_Txt As String)
    Text1.SelText = New_Txt
    'PropertyChanged "SelText"
End Property

Public Property Get SelStart() As Long
    SelStart = Text1.SelStart
End Property

Public Property Let SelStart(ByVal New_strt As Long)
    Text1.SelStart = New_strt
    'PropertyChanged "SelText"
End Property

Public Property Get SelLength() As Long
    SelLength = Text1.SelLength
End Property

Public Property Let SelLength(ByVal New_strt As Long)
    Text1.SelLength = New_strt
    'PropertyChanged "SelText"
End Property







Public Property Get Font() As Font
    Set Font = Text1.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Dim S1 As Single, S2 As Single
    Set Text1.Font = New_Font
    Set NumBar.Font = New_Font
    
'    LockWindowUpdate Text1.hwnd
'    S1 = Text1.SelStart
'    S2 = Text1.SelLength
'    Text1.SelStart = 0
'    Text1.SelLength = Len(Text1.Text)
'    Text1.SelFontSize = Text1.Font.Size
'    Text1.SelFontName = Text1.Font.Name
'    Text1.SelBold = Text1.Font.Bold
'    Text1.SelItalic = Text1.Font.Italic
'    Text1.SelStrikeThru = Text1.Font.Strikethrough
'    Text1.SelUnderline = Text1.Font.Underline
'    Text1.SelStart = S1
'    Text1.SelLength = S2
'    LockWindowUpdate 0
    
    PrintNums
    PropertyChanged "Font"
End Property

Public Property Get Enabled() As Boolean
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled() = New_Enabled
    PropertyChanged "Enabled"
End Property

Public Property Get Text() As String
    Text = Text1.Text
End Property

Public Property Let Text(ByVal New_Text As String)
    Text1.Text() = New_Text
    PropertyChanged "Text"
End Property

Public Property Get MouseIcon() As Picture
    Set MouseIcon = Text1.MouseIcon
End Property

Public Property Set MouseIcon(ByVal New_MouseIcon As Picture)
    Set Text1.MouseIcon = New_MouseIcon
    PropertyChanged "MouseIcon"
End Property

Public Property Get MousePointer() As MousePointerConstants
    MousePointer = Text1.MousePointer
End Property

Public Property Let MousePointer(ByVal New_MousePointer As MousePointerConstants)
    ' Validation is supplied by UserControl.
    Text1.MousePointer() = New_MousePointer
    PropertyChanged "MousePointer"
End Property

Public Property Get BackColor() As OLE_COLOR
    BackColor = Text1.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    Text1.BackColor() = New_BackColor
    picSep.BackColor() = New_BackColor
    PropertyChanged "BackColor"
End Property

'Public Property Get SelTextColor() As OLE_COLOR
'    SelTextColor = Text1.SelColor
'End Property

'Public Property Let SelTextColor(ByVal New_SelTextColor As OLE_COLOR)
'    Text1.SelColor() = New_SelTextColor
'    PropertyChanged "SelTextColor"
'End Property

Public Property Get Title() As String
    Title = MyTitle
End Property

Public Property Let Title(ByVal New_Title As String)
    MyTitle = New_Title
    RaiseEvent ChangeTitle(MyTitle)
    PrintLineCount
    PropertyChanged "Title"
End Property

Public Property Get Locked() As Boolean
    Locked = Text1.Locked
End Property

Public Property Let Locked(ByVal New_Locked As Boolean)
    Text1.Locked = New_Locked
    PropertyChanged "Locked"
End Property

Public Property Get BorderStyle() As BDRstyle
    BorderStyle = UserControl.BorderStyle
End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As BDRstyle)
    UserControl.BorderStyle = New_BorderStyle
    PropertyChanged "BorderStyle"
End Property

Public Property Get AutoVerbMenu() As Boolean
    AutoVerbMenu = AVMenabled
End Property

Public Property Let AutoVerbMenu(ByVal New_Menu As Boolean)
    'Text1.AutoVerbMenu = New_Menu
    AVMenabled = New_Menu
    PropertyChanged "AutoVerbMenu"
End Property

Public Property Get NumBackColor() As OLE_COLOR
    NumBackColor = NumBar.BackColor
End Property

Public Property Let NumBackColor(ByVal New_NumBackColor As OLE_COLOR)
    NumBar.BackColor = New_NumBackColor
    PrintNums
    PropertyChanged "NumBackColor"
End Property
Public Property Get NumForeColor() As OLE_COLOR
    NumForeColor = NumBar.ForeColor
End Property

Public Property Let NumForeColor(ByVal New_NumForeColor As OLE_COLOR)
    NumBar.ForeColor = New_NumForeColor
    PrintNums
    PropertyChanged "NumForeColor"
End Property

Public Property Get TitleForeColor() As OLE_COLOR
    TitleForeColor = TTLBack.ForeColor
End Property

Public Property Let TitleForeColor(ByVal New_TitleForeColor As OLE_COLOR)
    TTLBack.ForeColor = New_TitleForeColor
    PrintLineCount
    PropertyChanged "TitleForeColor"
End Property

Public Property Get TitleBackColor() As OLE_COLOR
    TitleBackColor = TTLBack.BackColor
End Property

Public Property Let TitleBackColor(ByVal New_TitleBackColor As OLE_COLOR)
    TTLBack.BackColor = New_TitleBackColor
    PrintLineCount
    PropertyChanged "TitleBackColor"
End Property

Public Property Get FileName() As String
    FileName = MyFileName
End Property

Public Property Let FileName(ByVal New_FileName As String)
    MyFileName = New_FileName
    PropertyChanged "FileName"
End Property

Public Property Get Numbar_Visible() As Boolean
    Numbar_Visible = NumBar.Visible
End Property

Public Property Let Numbar_Visible(ByVal NewVal As Boolean)
    NumBar.Visible = NewVal
    ResizeControls
    PropertyChanged "NumbarVisible"
End Property

'Public Property Get SyntaxColoring() As Boolean
'    SyntaxColoring = ColorText
'End Property

'Public Property Let SyntaxColoring(New_SC As Boolean)
'    ColorText = New_SC
'    If ColorText = True Then
'        doGetScriptKeywords
'    End If
'    PropertyChanged "ColorText"
'End Property


'--------Read and write properties ---------------->

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    Text1.BackColor = PropBag.ReadProperty("BackColor", &H80000005)
    picSep.BackColor = PropBag.ReadProperty("BackColor", &H80000005)
    'Text1.SelColor = PropBag.ReadProperty("SelTextColour", &H80000012)
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
    Set Font = PropBag.ReadProperty("Font", Ambient.Font)
    UserControl.BorderStyle = PropBag.ReadProperty("BorderStyle", 1)
    'm_Alignment = PropBag.ReadProperty("Alignment", 2)
    Text1.Text = PropBag.ReadProperty("Text", "Text")
    Set MouseIcon = PropBag.ReadProperty("MouseIcon", Nothing)
    Text1.MousePointer = PropBag.ReadProperty("MousePointer", 0)
    MyTitle = PropBag.ReadProperty("Title", "Untitled")
    Text1.Locked = PropBag.ReadProperty("Locked", False)
    'Text1.AutoVerbMenu = PropBag.ReadProperty("AutoVerbMenu", True)
    AVMenabled = PropBag.ReadProperty("AutoVerbMenu", True)
    NumBar.BackColor = PropBag.ReadProperty("NumBackColor", &H808080)
    NumBar.ForeColor = PropBag.ReadProperty("NumForeColor", &HFFFFFF)
    TTLBack.ForeColor = PropBag.ReadProperty("TitleForeColor", vbBlack)
    TTLBack.BackColor = PropBag.ReadProperty("TitleBackColor", vbButtonFace)
    MyFileName = PropBag.ReadProperty("FileName", "")
    'ColorText = PropBag.ReadProperty("ColorText", False)
    NumBar.Visible = PropBag.ReadProperty("NumbarVisible", True)
End Sub

Private Sub UserControl_Show()
    If Ambient.UserMode Then
        Timer1.Interval = 10
        Timer1.Enabled = True
    Else
        Timer1.Enabled = False
    End If
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("BackColor", Text1.BackColor, &H80000005)
    Call PropBag.WriteProperty("BackColor", picSep.BackColor, &H80000005)
    'Call PropBag.WriteProperty("SelTextColour", Text1.SelColor, &H80000012)
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
    Call PropBag.WriteProperty("Font", Font, Ambient.Font)
    Call PropBag.WriteProperty("BorderStyle", UserControl.BorderStyle, 1)
    'Call PropBag.WriteProperty("Alignment", m_Alignment, 2)
    Call PropBag.WriteProperty("Text", Text1.Text, "Text")
    Call PropBag.WriteProperty("MouseIcon", Text1.MouseIcon, Nothing)
    Call PropBag.WriteProperty("MousePointer", Text1.MousePointer, 0)
    Call PropBag.WriteProperty("Title", MyTitle, "Untitled")
    Call PropBag.WriteProperty("Locked", Text1.Locked, False)
    Call PropBag.WriteProperty("AutoverbMenu", AVMenabled, True)
    Call PropBag.WriteProperty("NumBackColor", NumBar.BackColor, &H808080)
    Call PropBag.WriteProperty("NumForeColor", NumBar.ForeColor, &HFFFFFF)
    Call PropBag.WriteProperty("TitleForeColor", TTLBack.ForeColor, vbBlack)
    Call PropBag.WriteProperty("TitleBackColor", TTLBack.BackColor, vbButtonFace)
    Call PropBag.WriteProperty("FileName", MyFileName, "")
    'Call PropBag.WriteProperty("ColorText", ColorText, False)
    Call PropBag.WriteProperty("NumbarVisible", NumBar.Visible, True)
End Sub

Function GetLineOfText() As String
Dim SP As Long, EP As Long
SP = Text1.SelStart
On Local Error Resume Next
If SP = 0 Then Exit Function
'If SP = Len(Text1.Text) Then Exit Function
Do Until Mid(Text1.Text, SP - 1, 1) = vbCrLf
    'Debug.Print """" & Mid(Text1.Text, SP - 1, 1) & """"
    If Mid(Text1.Text, SP - 1, 1) = vbCr Then Exit Do
    If Mid(Text1.Text, SP - 1, 1) = vbLf Then Exit Do
    If Mid(Text1.Text, SP - 1, 1) = vbCrLf Then Exit Do
    SP = SP - 1
    If SP = 0 Or SP = 1 Then Exit Do
Loop
EP = SP
Do Until EP = Len(Text1.Text) 'Or Mid(Text1.Text, EP, 1) = vbCr
    EP = EP + 1
Loop
'EP = EP + 2
    If Mid(Text1.Text, SP, EP - (SP - 1)) = vbCr Then Exit Function
    If Mid(Text1.Text, SP, EP - (SP - 1)) = vbLf Then Exit Function
    If Mid(Text1.Text, SP, EP - (SP - 1)) = vbCrLf Then Exit Function

GetLineOfText = """" & Mid(Text1.Text, SP, EP - (SP - 1)) & """"
End Function

Public Sub txtCut()
    AVMcut_Click
End Sub
Public Sub txtCopy()
    AVMcopy_Click
End Sub
Public Sub txtPaste()
    AVMpaste_Click
End Sub
Public Sub txtDelete()
    AVMdelete_Click
End Sub
Public Sub txtSelect()
    AVMselectall_Click
End Sub

Public Sub WordWrap(ww As Boolean)
    If ww = True Then
        Text1.RightMargin = 5000000
'        Text1.ScrollBars = rtfBoth
    Else
        Text1.RightMargin = 0
'        Text1.ScrollBars = rtfVertical
    End If
    PrintLineCount
End Sub

'Private Sub ColorizeCurText()
'    Dim SelPos As Long
'    SelPos = apiSendMessage(Text1.hwnd, EM_GETSEL + EM_GETFIRSTVISIBLELINE, 0&, 0&)
'    Debug.Print SelPos
'End Sub
