VERSION 5.00
Begin VB.Form frmMain 
   Appearance      =   0  'Flat
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "TSID-KEYGEN"
   ClientHeight    =   1710
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5310
   BeginProperty Font 
      Name            =   "Terminal"
      Size            =   9
      Charset         =   255
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1710
   ScaleWidth      =   5310
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Quit 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      Caption         =   "Quit"
      Height          =   255
      Index           =   2
      Left            =   3600
      MaskColor       =   &H00FFFF00&
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   1200
      Width           =   1455
   End
   Begin VB.CommandButton About 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      Caption         =   "About"
      Height          =   255
      Index           =   1
      Left            =   1920
      MaskColor       =   &H00FFFF00&
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1200
      Width           =   1455
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      ForeColor       =   &H00FFFF00&
      Height          =   255
      Left            =   720
      TabIndex        =   2
      Top             =   720
      Width           =   4335
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      ForeColor       =   &H00FFFF80&
      Height          =   255
      Left            =   720
      TabIndex        =   0
      Top             =   240
      Width           =   4335
   End
   Begin VB.CommandButton Cmd 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      Caption         =   "Generate"
      Height          =   255
      Index           =   0
      Left            =   240
      MaskColor       =   &H00FFFF00&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1200
      Width           =   1455
   End
   Begin VB.Label Lbl 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "TSID:"
      ForeColor       =   &H00FFFF00&
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   4
      Top             =   720
      Width           =   495
   End
   Begin VB.Label Lbl 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "TEXT:"
      ForeColor       =   &H00FFFF00&
      Height          =   180
      Index           =   0
      Left            =   120
      TabIndex        =   3
      Top             =   240
      Width           =   495
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private TSID As clsTSID

Private Sub About_Click(Index As Integer)
MsgBox ("TSID-KEYGEN V1.0.5")
End Sub

Private Sub Cmd_Click(Index As Integer)
    Set TSID = New clsTSID
    Text2 = TSID.CalculateTSID(Text1)
End Sub

Private Sub Quit_Click(Index As Integer)
End
End Sub
