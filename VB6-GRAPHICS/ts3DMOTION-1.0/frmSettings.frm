VERSION 5.00
Begin VB.Form frmSettings 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "3D Motion"
   ClientHeight    =   3735
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4575
   Icon            =   "frmSettings.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3735
   ScaleWidth      =   4575
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Caption         =   "Effects : "
      ForeColor       =   &H00FF0000&
      Height          =   855
      Left            =   240
      TabIndex        =   9
      Top             =   1560
      Width           =   4095
      Begin VB.CheckBox ChkFillColor 
         Caption         =   "Fill circles with color ."
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   480
         Value           =   1  'Checked
         Width           =   3855
      End
      Begin VB.CheckBox ChkFading 
         Caption         =   "Fading Effect ( Not recommended for slow PCs )"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Value           =   1  'Checked
         Width           =   3735
      End
   End
   Begin VB.HScrollBar HSCNumPoints 
      Height          =   255
      LargeChange     =   5
      Left            =   240
      Max             =   200
      Min             =   1
      TabIndex        =   7
      Top             =   2760
      Value           =   25
      Width           =   3615
   End
   Begin VB.CommandButton CmdCancle 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   2400
      TabIndex        =   5
      Top             =   3120
      Width           =   1935
   End
   Begin VB.CommandButton CmdOk 
      Caption         =   "&Ok"
      Default         =   -1  'True
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   3120
      Width           =   2055
   End
   Begin VB.Frame Frame1 
      Caption         =   "Style of Rotation :"
      ForeColor       =   &H00FF0000&
      Height          =   1335
      Left            =   240
      TabIndex        =   0
      Top             =   150
      Width           =   4095
      Begin VB.OptionButton OptLines2 
         Caption         =   "Lines rotate in 3D motion but move in 2D motion."
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   120
         TabIndex        =   12
         Top             =   960
         Width           =   3855
      End
      Begin VB.OptionButton OptLines1 
         Caption         =   "Lines rotate and move  in 3D motion."
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   120
         TabIndex        =   3
         Top             =   720
         Width           =   3855
      End
      Begin VB.OptionButton OptCircles2 
         Caption         =   "Circles rotate in 3D motion but move in 2D motion."
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   120
         TabIndex        =   2
         Top             =   480
         Width           =   3855
      End
      Begin VB.OptionButton OptCircles1 
         Caption         =   "Circles rotate and move in 3D motion ."
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Value           =   -1  'True
         Width           =   3015
      End
   End
   Begin VB.Label LblNumPoints 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "25"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   3960
      TabIndex        =   8
      Top             =   2760
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Number of points:"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   2520
      Width           =   1455
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H8000000F&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00808080&
      FillColor       =   &H80000012&
      Height          =   3495
      Left            =   120
      Top             =   120
      Width           =   4335
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C0C0C0&
      FillColor       =   &H80000012&
      Height          =   3525
      Left            =   150
      Top             =   150
      Width           =   4365
   End
End
Attribute VB_Name = "FrmSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub ChkFading_Click()
    'Fadign effect
    FadingEffect = ChkFading
End Sub

Private Sub ChkFillColor_Click()
    'Fill circles with color
    CirclesFillColor = ChkFillColor
End Sub

Private Sub CmdCancle_Click()
    Unload Me
End Sub

Private Sub CmdOk_Click()
    'Show the main from
    FrmMain.Show
    
    'Unload Settings form
    Unload Me
End Sub

Private Sub Form_Load()
    'Default Values
    ObjectStyle = Circles1
    FadingEffect = True
    CirclesFillColor = True
    'Number of points
    HSCNumPoints_Change
End Sub

Private Sub HSCNumPoints_Change()
    
    Dim MyValue As Integer
    
    'Get the value of the scroll bar
    MyValue = HSCNumPoints.Value
    
    'If the style is lines, then you need 4 points at least
    If ObjectStyle = Lines1 Or ObjectStyle = Lines2 Then
        MyValue = CheckNumberOfPoints()
    End If
    
    'Change the number of Points
    LblNumPoints = MyValue
    NumOfPoints = MyValue
End Sub

Private Sub OptCircles1_Click()
    'Circles rotate and move in 3D motion .
    ObjectStyle = Circles1
End Sub

Private Sub OptCircles2_Click()
    'Circles rotate in 3D but move in 2D.
    ObjectStyle = Circles2
End Sub

Private Sub OptLines1_Click()
    'Lines rotate and move in 3D.
    ObjectStyle = Lines1
    
    'Check if you have less than 4 points for this style
    Call CheckNumberOfPoints
End Sub

Private Sub OptLines2_Click()
    'Lines rotate in 3D motion but move in 2D motion.
    ObjectStyle = Lines2
    
    'Check if you have less than 4 points for this style
    Call CheckNumberOfPoints
End Sub

Private Function CheckNumberOfPoints() As Integer
    'If the style is lines, then you need 4 points at least
    If HSCNumPoints.Value < 4 Then
        'Show error message
        MsgBox "You need at least 4 points when you use the 'Lines' style", vbExclamation, "Style error"
        HSCNumPoints.Value = 4
    End If
    'Return value
    CheckNumberOfPoints = HSCNumPoints.Value
End Function

