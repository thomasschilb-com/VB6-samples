VERSION 5.00
Begin VB.Form frmMain 
   Appearance      =   0  'Flat
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "TSID-KEYGEN"
   ClientHeight    =   2025
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5295
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
   Picture         =   "frmMain.frx":058A
   ScaleHeight     =   2025
   ScaleWidth      =   5295
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Quit 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   3600
      MaskColor       =   &H00FFFF00&
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   1560
      Width           =   1455
   End
   Begin VB.CommandButton About 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      Caption         =   "About"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   1920
      MaskColor       =   &H00FFFF00&
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1560
      Width           =   1455
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   1080
      Width           =   4815
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF80&
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   480
      Width           =   4815
   End
   Begin VB.CommandButton Cmd 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      Caption         =   "Generate"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   240
      MaskColor       =   &H00FFFF00&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1560
      Width           =   1455
   End
   Begin VB.Label Lbl 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "TSID:"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   4
      Top             =   840
      Width           =   1335
   End
   Begin VB.Label Lbl 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "INPUT:"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   225
      Index           =   0
      Left            =   240
      TabIndex        =   3
      Top             =   240
      Width           =   1335
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
MsgBox ("TSID-KEYGEN V1.0.7")
End Sub

Private Sub Cmd_Click(Index As Integer)
    Set TSID = New clsTSID
    Text2 = TSID.CalculateTSID(Text1)
End Sub

Private Sub Quit_Click(Index As Integer)
End
End Sub
