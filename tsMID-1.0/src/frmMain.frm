VERSION 5.00
Object = "{C1A8AF28-1257-101B-8FB0-0020AF039CA3}#1.1#0"; "mci32.ocx"
Begin VB.Form frmMain 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "tsMID"
   ClientHeight    =   1995
   ClientLeft      =   1125
   ClientTop       =   1455
   ClientWidth     =   4950
   BeginProperty Font 
      Name            =   "Terminal"
      Size            =   6
      Charset         =   255
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   Picture         =   "frmMain.frx":058A
   ScaleHeight     =   1995
   ScaleWidth      =   4950
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   375
      Left            =   3360
      TabIndex        =   5
      Top             =   1320
      Width           =   1215
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Stop"
      Height          =   375
      Left            =   1920
      TabIndex        =   4
      Top             =   1320
      Width           =   1215
   End
   Begin VB.TextBox txtFilename 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   960
      TabIndex        =   3
      Text            =   "C:\WinNT\Media\passport.mid"
      Top             =   2040
      Width           =   3375
   End
   Begin VB.CommandButton cmdPlay 
      Caption         =   "Play Me"
      Height          =   375
      Left            =   480
      TabIndex        =   0
      Top             =   1320
      Width           =   1215
   End
   Begin MCI.MMControl MMControl1 
      Height          =   495
      Left            =   720
      TabIndex        =   2
      Top             =   2400
      Visible         =   0   'False
      Width           =   3540
      _ExtentX        =   6244
      _ExtentY        =   873
      _Version        =   393216
      DeviceType      =   ""
      FileName        =   ""
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "A Simple Midi-Player"
      ForeColor       =   &H00FFFF00&
      Height          =   135
      Left            =   1560
      TabIndex        =   7
      Top             =   840
      Width           =   1815
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "tsMID"
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   13.5
         Charset         =   255
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   255
      Left            =   2040
      TabIndex        =   6
      Top             =   360
      Width           =   855
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "WAV File"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   240
      TabIndex        =   1
      Top             =   2040
      Width           =   660
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClose_Click()
    cmdPlay.Enabled = True
    MMControl1.Command = "Close"
End Sub

Private Sub cmdExit_Click()
    cmdPlay.Enabled = False
    MMControl1.Command = "Close"
    End
End Sub

' Open the device and play the sound.
Private Sub cmdPlay_Click()
    cmdPlay.Enabled = False

    ' Set the file name.
    MMControl1.FileName = txtFilename.Text

    ' Open the MCI device.
    MMControl1.Wait = True
    MMControl1.Command = "Open"

    ' Play the sound without waiting.
    MMControl1.Command = "Play"
End Sub

Private Sub Form_Load()
    txtFilename.Text = App.Path
    If Right$(txtFilename.Text, 1) <> "\" _
        Then txtFilename.Text = txtFilename.Text & "\"
    txtFilename.Text = txtFilename.Text & "tsMID.mid"

    ' Prepare the MCI control for WaveAudio.
    MMControl1.Notify = False
    MMControl1.Wait = True
    MMControl1.Shareable = False
    MMControl1.DeviceType = "Sequencer"
End Sub


' Close the device.
Private Sub MMControl1_Done(NotifyCode As Integer)
    MMControl1.Command = "Close"
    cmdPlay.Enabled = True
End Sub
