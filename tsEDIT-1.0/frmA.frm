VERSION 5.00
Begin VB.Form frmA 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Asociations"
   ClientHeight    =   3675
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4620
   Icon            =   "frmA.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   245
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   308
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command3 
      Caption         =   "&Default"
      Height          =   495
      Left            =   3240
      TabIndex        =   20
      Top             =   3000
      Width           =   1215
   End
   Begin VB.PictureBox Back 
      BackColor       =   &H00FFFFFF&
      Height          =   2175
      Left            =   120
      ScaleHeight     =   141
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   189
      TabIndex        =   3
      Top             =   1320
      Width           =   2895
      Begin VB.CheckBox chkO 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   4
         Left            =   0
         TabIndex        =   18
         Top             =   960
         Value           =   1  'Checked
         Width           =   195
      End
      Begin VB.CheckBox chkE 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   4
         Left            =   240
         TabIndex        =   17
         Top             =   960
         Width           =   195
      End
      Begin VB.CheckBox chkO 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   3
         Left            =   0
         TabIndex        =   15
         Top             =   720
         Width           =   195
      End
      Begin VB.CheckBox chkE 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   14
         Top             =   720
         Value           =   1  'Checked
         Width           =   195
      End
      Begin VB.CheckBox chkO 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   2
         Left            =   0
         TabIndex        =   12
         Top             =   480
         Width           =   195
      End
      Begin VB.CheckBox chkE 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   11
         Top             =   480
         Value           =   1  'Checked
         Width           =   195
      End
      Begin VB.CheckBox chkO 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   1
         Left            =   0
         TabIndex        =   9
         Top             =   240
         Width           =   195
      End
      Begin VB.CheckBox chkE 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   8
         Top             =   240
         Value           =   1  'Checked
         Width           =   195
      End
      Begin VB.CheckBox chkE 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   4
         Top             =   0
         Width           =   195
      End
      Begin VB.CheckBox chkO 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   0
         Left            =   0
         TabIndex        =   5
         Top             =   0
         Value           =   1  'Checked
         Width           =   195
      End
      Begin VB.Label Label7 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Information Files (*.nfo;*.diz)"
         Height          =   225
         Left            =   480
         TabIndex        =   19
         Top             =   990
         Width           =   2415
      End
      Begin VB.Label Label6 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Configuration Files (*.ini;*.inf)"
         Height          =   225
         Left            =   480
         TabIndex        =   16
         Top             =   750
         Width           =   2415
      End
      Begin VB.Label Label5 
         BackColor       =   &H00FFFFFF&
         Caption         =   "VB Files (*.vbp;*.frm;*.bas;*.cls)"
         Height          =   225
         Left            =   480
         TabIndex        =   13
         Top             =   510
         Width           =   2415
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FFFFFF&
         Caption         =   "HTML Files (*.htm;*.html;*.js)"
         Height          =   225
         Left            =   480
         TabIndex        =   10
         Top             =   270
         Width           =   2415
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Text Files (*.txt)"
         Height          =   225
         Left            =   480
         TabIndex        =   6
         Top             =   30
         Width           =   2415
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Cancel"
      Height          =   495
      Left            =   3240
      TabIndex        =   2
      Top             =   1920
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&OK"
      Height          =   495
      Left            =   3240
      TabIndex        =   1
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   1
      X1              =   8
      X2              =   296
      Y1              =   41
      Y2              =   41
   End
   Begin VB.Label Label3 
      Caption         =   "Check the left checkbox for Open or Check the Right checkbox for Edit"
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   720
      Width           =   4335
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000010&
      Index           =   0
      X1              =   8
      X2              =   296
      Y1              =   40
      Y2              =   40
   End
   Begin VB.Label Label1 
      Caption         =   "Select the file extensions that you want New Notepad to asociate with:"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4335
   End
End
Attribute VB_Name = "frmA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chkE_Click(Index As Integer)
    chkO(Index).Value = IIf(chkE(Index).Value = 0, 1, 0)
End Sub

Private Sub chkO_Click(Index As Integer)
    chkE(Index).Value = IIf(chkO(Index).Value = 0, 1, 0)
End Sub

Private Sub Command1_Click()
    Associate App.Path, "rtf", "e"
    Unload Me
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Command3_Click()
    chkO(0).Value = 1
    chkO(1).Value = 0
    chkO(2).Value = 0
    chkO(3).Value = 0
    chkO(4).Value = 1
    
    chkE(0).Value = 0
    chkE(1).Value = 1
    chkE(2).Value = 1
    chkE(3).Value = 1
    chkE(4).Value = 0
End Sub
