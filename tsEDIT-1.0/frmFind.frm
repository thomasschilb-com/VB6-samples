VERSION 5.00
Begin VB.Form frmFind 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Find"
   ClientHeight    =   1965
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5520
   Icon            =   "frmFind.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1965
   ScaleWidth      =   5520
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox Text2 
      Height          =   525
      Left            =   1320
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   1200
      Width           =   2655
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Ca&se Sensetive"
      Height          =   255
      Left            =   1320
      TabIndex        =   2
      Top             =   840
      Width           =   2655
   End
   Begin VB.CommandButton ReplaceButton 
      Caption         =   "&Replace"
      Height          =   375
      Left            =   4200
      TabIndex        =   5
      Top             =   720
      Width           =   1215
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   4200
      TabIndex        =   6
      Top             =   1200
      Width           =   1215
   End
   Begin VB.CommandButton FindNextButton 
      Caption         =   "Find &Next"
      Height          =   375
      Left            =   4200
      TabIndex        =   4
      Top             =   240
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   525
      Left            =   1320
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   240
      Width           =   2655
   End
   Begin VB.Label Label2 
      Caption         =   "Replace &With:"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   1200
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Fi&nd What:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   1335
   End
End
Attribute VB_Name = "frmFind"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public SearchString As String
Public CurrentPosition As Long
Public TotalShutdown As Boolean

Public Sub Find(Optional mReplace As Boolean)
    Dim posa As Long, triedGoingBack As Boolean
    
    CurrentPosition = frmMain.Gtxt1.SelStart + 2
    If CurrentPosition = 0 Then CurrentPosition = 1
    
TryAgain:
    
    If Check1.Value = 0 Then
        posa = InStr(CurrentPosition, frmMain.Gtxt1.Text, Text1.Text, vbTextCompare)
    Else
        posa = InStr(CurrentPosition, frmMain.Gtxt1.Text, Text1.Text)
    End If
    
    On Error GoTo EXT
    
    If posa <> 0 Then
        frmMain.Gtxt1.SelStart = posa - 1
        frmMain.Gtxt1.SelLength = Len(Text1.Text)
        If mReplace Then
            frmMain.Gtxt1.SelText = Text2.Text
        End If
    Else
        If Not triedGoingBack Then triedGoingBack = True: CurrentPosition = 1: GoTo TryAgain
        MsgBox "Can't find any more of """ & Text1.Text & """", vbInformation, "Find"
    End If
    
EXT:
End Sub

Private Sub CancelButton_Click()
    Unload Me
End Sub

Private Sub FindNextButton_Click()
    Find
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not AppEnd Then Cancel = True: Me.Hide
End Sub

Private Sub ReplaceButton_Click()
    If LCase(frmMain.Gtxt1.SelText) <> LCase(Text1.Text) Then
        Find
    Else
        frmMain.Gtxt1.SelText = Text2.Text
        Find
    End If
End Sub
