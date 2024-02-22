VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FormConfig 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "   Options"
   ClientHeight    =   2535
   ClientLeft      =   3540
   ClientTop       =   2835
   ClientWidth     =   5055
   ControlBox      =   0   'False
   Icon            =   "Config.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2535
   ScaleWidth      =   5055
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   1575
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   4815
      Begin VB.CommandButton CmdOK 
         Caption         =   "OK"
         Default         =   -1  'True
         Height          =   375
         Left            =   3360
         TabIndex        =   3
         Top             =   480
         Width           =   1095
      End
      Begin VB.CommandButton CmdCancel 
         Cancel          =   -1  'True
         Caption         =   "Cancel"
         Height          =   375
         Left            =   3360
         TabIndex        =   2
         Top             =   960
         Width           =   1095
      End
      Begin MSComctlLib.Slider Slider1 
         Height          =   375
         Left            =   240
         TabIndex        =   1
         Top             =   840
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   661
         _Version        =   393216
         Min             =   5
         Max             =   30
         SelStart        =   5
         Value           =   5
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "Lines"
         Height          =   255
         Left            =   360
         TabIndex        =   4
         Top             =   480
         Width           =   2295
      End
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "Sine Line"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   240
      Width           =   4815
   End
End
Attribute VB_Name = "FormConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Sine Function Line Screen Saver

' Created by:
' Rod Stephens (Core functionality)
' Neil Fraser (Integration & Win 98 multi-monitors)
' Elliot Spencer (Win 9x locking & registry lookup)
' Don Bradner & Jim Deutch (Password handling)
' Lucian Wischik & Alex Millman (NT information)
' Ken Slater - 0x34 - (Grafix and Motion Code)

Private Sub Form_Load()
  LoadConfig
  Slider1 = Density
  Label2 = Slider1.Value
End Sub


Private Sub CmdOK_Click()
  On Error Resume Next
  Density = Slider1.Value
  On Error GoTo 0
  SaveConfig
  CmdCancel_Click
End Sub

Private Sub CmdCancel_Click()
  Unload Me
End Sub

Private Sub Slider1_Change()
    Label2 = "Shadow Density = " & Slider1.Value
End Sub

Private Sub Slider1_Click()
    Label2 = "Shadow Density = " & Slider1.Value
End Sub

Private Sub Slider1_Scroll()
    Label2 = "Shadow Density = " & Slider1.Value
End Sub
