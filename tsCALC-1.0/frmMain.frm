VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00404040&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "tsCALC"
   ClientHeight    =   4215
   ClientLeft      =   150
   ClientTop       =   840
   ClientWidth     =   2775
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4215
   ScaleWidth      =   2775
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox TxtResult 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   420
      Left            =   240
      TabIndex        =   20
      Top             =   120
      Width           =   2295
   End
   Begin VB.CommandButton Command2 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      Caption         =   "Sqr"
      Height          =   495
      Left            =   1440
      MaskColor       =   &H00808080&
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   2880
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      Caption         =   "%"
      Height          =   495
      Left            =   840
      MaskColor       =   &H00808080&
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   3480
      Width           =   495
   End
   Begin VB.CommandButton CmdNegative 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      Caption         =   "-/+"
      Height          =   495
      Left            =   240
      MaskColor       =   &H00808080&
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   3480
      Width           =   495
   End
   Begin VB.TextBox TxtInputNum 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   420
      Left            =   240
      MaxLength       =   17
      TabIndex        =   16
      Top             =   600
      Width           =   2295
   End
   Begin VB.CommandButton Signs 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      Caption         =   "/"
      Height          =   495
      Index           =   3
      Left            =   2040
      MaskColor       =   &H00808080&
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   2880
      Width           =   495
   End
   Begin VB.CommandButton Signs 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      Caption         =   "*"
      Height          =   495
      Index           =   2
      Left            =   2040
      MaskColor       =   &H00808080&
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   2280
      Width           =   495
   End
   Begin VB.CommandButton Signs 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      Caption         =   "-"
      Height          =   495
      Index           =   1
      Left            =   2040
      MaskColor       =   &H00808080&
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   1680
      Width           =   495
   End
   Begin VB.CommandButton Signs 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      Caption         =   "+"
      Height          =   495
      Index           =   0
      Left            =   2040
      MaskColor       =   &H00808080&
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   1080
      Width           =   495
   End
   Begin VB.CommandButton cmdCalculate 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      Caption         =   "="
      Height          =   495
      Left            =   1440
      MaskColor       =   &H00808080&
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   3480
      Width           =   1095
   End
   Begin VB.CommandButton cmdDecimal 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      Caption         =   "."
      Height          =   495
      Left            =   840
      MaskColor       =   &H00808080&
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   2880
      Width           =   495
   End
   Begin VB.CommandButton Number 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      Caption         =   "9"
      Height          =   495
      Index           =   9
      Left            =   1440
      MaskColor       =   &H00808080&
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   2280
      Width           =   495
   End
   Begin VB.CommandButton Number 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      Caption         =   "8"
      Height          =   495
      Index           =   8
      Left            =   840
      MaskColor       =   &H00808080&
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   2280
      Width           =   495
   End
   Begin VB.CommandButton Number 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      Caption         =   "7"
      Height          =   495
      Index           =   7
      Left            =   240
      MaskColor       =   &H00808080&
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   2280
      Width           =   495
   End
   Begin VB.CommandButton Number 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      Caption         =   "6"
      Height          =   495
      Index           =   6
      Left            =   1440
      MaskColor       =   &H00808080&
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   1680
      Width           =   495
   End
   Begin VB.CommandButton Number 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      Caption         =   "5"
      Height          =   495
      Index           =   5
      Left            =   840
      MaskColor       =   &H00808080&
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1680
      Width           =   495
   End
   Begin VB.CommandButton Number 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      Caption         =   "4"
      Height          =   495
      Index           =   4
      Left            =   240
      MaskColor       =   &H00808080&
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1680
      Width           =   495
   End
   Begin VB.CommandButton Number 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      Caption         =   "3"
      Height          =   495
      Index           =   3
      Left            =   1440
      MaskColor       =   &H00808080&
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1080
      Width           =   495
   End
   Begin VB.CommandButton Number 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      Caption         =   "2"
      Height          =   495
      Index           =   2
      Left            =   840
      MaskColor       =   &H00808080&
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1080
      Width           =   495
   End
   Begin VB.CommandButton Number 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      Caption         =   "1"
      Height          =   495
      Index           =   1
      Left            =   240
      MaskColor       =   &H00808080&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1080
      Width           =   495
   End
   Begin VB.CommandButton Number 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      Caption         =   "0"
      Height          =   495
      Index           =   0
      Left            =   240
      MaskColor       =   &H00808080&
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2880
      Width           =   495
   End
   Begin VB.Menu mnuClear 
      Caption         =   "Clear"
   End
   Begin VB.Menu mnuAbout 
      Caption         =   "About"
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'I've just seen heaps of calculators on this site with heaps
'and heaps of unnecessary code so I decided to make my own
'which is very simple.
'If you have any comments like - something that I could have
'done better or more efficiently. Then please tell me because I am
'still learning to programme and need feedback. I have only been coding for a
'month so go easy if ya decide to comment.
Option Explicit
    Dim DecFlag     As Boolean 'Decimal Flag
    Dim Calculate   As Boolean 'Calculate flag
    Dim Opt1        As Variant 'First Number
    Dim Opt2        As Variant 'Second Number
    Dim StrSign     As String * 1 'A string that can only accept 1 char
Private Sub cmdCalculate_Click() 'If the equals button was pressed
     If StrSign = "+" Then
        TxtInputNum.Text = Val(Opt1) + Val(TxtInputNum.Text)
     ElseIf StrSign = "-" Then
        TxtInputNum.Text = Val(Opt1) - Val(TxtInputNum.Text)
     ElseIf StrSign = "*" Then
        TxtInputNum.Text = Val(Opt1) * Val(TxtInputNum.Text)
     ElseIf StrSign = "/" Then
        TxtInputNum.Text = Val(Opt1) / Val(TxtInputNum.Text)
     End If
     
     If Calculate = True Then Call mnuClear_Click
     Calculate = True
End Sub
Private Sub cmdDecimal_Click()
    If DecFlag = True Then 'Check if a decimal has already being pressed
        Exit Sub
    Else
        TxtInputNum.Text = TxtInputNum.Text & "."
        DecFlag = True
    End If
End Sub
Private Sub CmdNegative_Click()
   TxtInputNum.Text = Val(TxtInputNum) * -1 'Nice trick eh
End Sub

Private Sub Command1_Click()
    TxtInputNum.Text = Val(TxtInputNum.Text) / 100
End Sub

Private Sub Command2_Click()
    TxtInputNum.Text = Sqr(TxtInputNum.Text)
End Sub

Private Sub mnuAbout_Click()
    MsgBox "tsCALC 1.0.0.1", vbInformation, "About"
End Sub

Private Sub mnuClear_Click() 'Clear everything
    DecFlag = False
    Opt1 = ""
    Opt2 = ""
    StrSign = ""
    TxtInputNum.Text = ""
    TxtResult = ""
End Sub

Private Sub Number_Click(Index As Integer) 'Put a corresponding number in the text box
    If TxtInputNum.Text = "" And Number(0) Then Exit Sub
    TxtInputNum.Text = TxtInputNum.Text & Number(Index).Caption
End Sub

Private Sub Signs_Click(Index As Integer) 'Put and calculate for the sign
    Calculate = False
    Select Case Signs(Index)
        Case Signs(0) 'Addition
                StrSign = "+"
                If Opt1 = "" Then
                    Opt1 = TxtInputNum.Text
                Else
                    Opt2 = Val(Opt1) + Val(TxtInputNum.Text)
                    Opt1 = Opt2
                End If
                TxtInputNum.Text = ""
         Case Signs(1) 'Subtraction
                StrSign = "-"
                If Opt1 = "" Then
                    Opt1 = TxtInputNum.Text
                Else
                    Opt2 = Val(Opt1) - Val(TxtInputNum.Text)
                    Opt1 = Opt2
                End If
                TxtInputNum.Text = ""
        Case Signs(2) 'Multiplication
                StrSign = "*"
                If Opt1 = "" Then
                    Opt1 = TxtInputNum.Text
                Else
                    Opt2 = Val(Opt1) * Val(TxtInputNum.Text)
                    Opt1 = Opt2
                End If
                TxtInputNum.Text = ""
        Case Signs(3) 'Division
                StrSign = "/"
                If Opt1 = "" Then
                    Opt1 = TxtInputNum.Text
                Else
                    Opt2 = Val(Opt1) / Val(TxtInputNum.Text)
                    Opt1 = Opt2
                End If
                TxtInputNum.Text = ""
    End Select
    TxtResult.Text = Opt1
End Sub
