VERSION 5.00
Begin VB.Form FormDisplay 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   1845
   ClientLeft      =   2010
   ClientTop       =   2430
   ClientWidth     =   1845
   ControlBox      =   0   'False
   FillStyle       =   0  'Solid
   Icon            =   "Display.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   123
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   123
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   120
      Top             =   120
   End
End
Attribute VB_Name = "FormDisplay"
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

Private BufferTime As Single
Private PwMode As Boolean
Const pi = 3.14159265358979
Dim A As Integer
Dim B As Integer
Dim C As Integer
Dim D As Integer
Dim BP1 As Integer
Dim BP2 As Integer
Dim BP3 As Integer
Dim BP4 As Integer
Dim CoX1 As Long
Dim CoY1 As Long
Dim CoX2 As Long
Dim CoY2 As Long
Dim iColor As ColorConstants
Dim Bucket(30, 3) As Long
Dim MCR As Boolean
Dim MCG As Boolean
Dim MCB As Boolean
Dim MCRd As Boolean
Dim MCGd As Boolean
Dim MCBd As Boolean
Dim Rm As Integer
Dim Gm As Integer
Dim Bm As Integer

' Try to terminate.  This code would never be called in NT.
Private Sub HumanEvent()
  If PreviewMode Then
    ' human events shouldn't close the preview mode
  ElseIf BufferTime <= Timer And Timer < BufferTime + 0.5 Then
    ' give the person a half second to stop moving the mouse after cancelling a password box
  ElseIf PwMode Then
    ' if passwords are turned on check the password
    ShowCursor True ' need cursor for the dialog box
    If VerifyScreenSavePwd(Me.hwnd) Then Unload Me
    ShowCursor False
    BufferTime = Timer
  Else
    Unload Me
  End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  HumanEvent
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
  HumanEvent
End Sub

Private Sub Form_Click()
  HumanEvent
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
  HumanEvent
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Static x0 As Single
Static y0 As Single
  ' unload on large mouse movements
  If ((x0 = 0) And (y0 = 0)) Or _
    ((Abs(x0 - x) <= 1) And (Abs(y0 - y) <= 1)) Then
      ' it's just a small movement
      x0 = x
      y0 = y
  Else
    HumanEvent
  End If
End Sub

Private Sub Form_Load()
  ' load configuration information
  LoadConfig
    If Density < 5 Or Density > 30 Then 'For first time runs
        Density = 15
        SaveConfig
    End If
  
    Rm = 100
    Gm = 50
    Bm = 255
    A = 5
    B = 3
    C = 1
    D = 2
    BP1 = 50
    BP2 = 100
    BP3 = 75
    BP4 = 125
    Timer1.Interval = 10
    Timer1.Enabled = True
  
  If Not PreviewMode Then
    ' record the time, ignore human actions for the next half second
    BufferTime = Timer
    
    ' Find out if the user has enabled passwords for this screen saver
    PwMode = CBool(Val(ReadRegistry(HKEY_CURRENT_USER, "Control Panel\Desktop", "ScreenSaveUsePassword")))
    ' Stop Ctrl+Alt+Del (among other things)
    SystemParametersInfo SPI_SCREENSAVERRUNNING, 1&, 0&, 0&
    
    ' hide the cursor
    ShowCursor False
    
  End If
  
  Randomize
End Sub

Private Sub Form_Unload(Cancel As Integer)
  If Not PreviewMode Then
    ' Redisplay the cursor
    ShowCursor True
    ' Reenable Ctrl+Alt+Del (among other things)
    SystemParametersInfo SPI_SCREENSAVERRUNNING, 0&, 0&, 0&
  End If
End Sub

' Draw lines
Private Sub Timer1_Timer()
    'Caluculate Two X,Y Coordinates
    CoY1 = (((Me.ScaleHeight) * Sin(pi * BP1 / 100) / 2) + (Me.ScaleHeight / 2))
    CoX1 = (((Me.ScaleWidth) * Sin(pi * BP2 / 100) / 2) + (Me.ScaleWidth / 2))
    CoY2 = (((Me.ScaleHeight) * Sin(pi * BP3 / 100) / 2) + (Me.ScaleHeight / 2))
    CoX2 = (((Me.ScaleWidth) * Sin(pi * BP4 / 100) / 2) + (Me.ScaleWidth / 2))
    'Draw a line on the Form between the two Coordinates
    Line (CoX1, CoY1)-(CoX2, CoY2), GetColor
    'Store these Coordinates in the bucket and clear the last line
    HndlBucket
    'Move the coordinates based on random values
    BP1 = BP1 + D
    BP2 = BP2 + C
    BP3 = BP3 + A
    BP4 = BP4 + B
    'If we've cycled through 360 degrees of motion on one or both coords, select new move value
    If BP1 > 200 Then
        BP1 = 0
        D = GetSpeed(D)
    End If
    If BP2 > 200 Then
        BP2 = 0
        C = GetSpeed(C)
    End If
    If BP3 > 200 Then
        BP3 = 0
        A = GetSpeed(A)
    End If
    If BP4 > 200 Then
        BP4 = 0
        B = GetSpeed(B)
    End If
End Sub

'This sub handles line removal - FILO Bucket Bergade
Private Sub HndlBucket()
Dim T As Integer
Dim U As Integer
    'Shift Bucket based on user selected density
    For T = Density To 1 Step -1
        For U = 0 To 3
            Bucket(T, U) = Bucket((T - 1), U)
        Next
    Next
    'Place newest line in the bucket
    Bucket(0, 0) = CoX1
    Bucket(0, 1) = CoY1
    Bucket(0, 2) = CoX2
    Bucket(0, 3) = CoY2
    'Erase (Clear) the last line in the bucket
    Line (Bucket(Density, 0), Bucket(Density, 1))-(Bucket(Density, 2), Bucket(Density, 3)), Me.BackColor
End Sub

' Random speed select for both axis
Private Function GetSpeed(ByVal M As Integer) As Integer
Dim W As Integer
Dim K As Integer
    Randomize
    K = Int(Rnd * 6)
        If K = 2 And M < 6 Then
            M = M + 1
        ElseIf K = 5 Then
            If M < 1 Then
                M = M + 1
            Else
                M = M - 1
            End If
        End If
        If M = 0 Then M = 1
    GetSpeed = M
End Function

'Create Molly Color Morphing
Public Function GetColor() As ColorConstants    'Create Molly Color Morphing
Dim Clrdr As Integer
    'Select a color to change
    If MCR = False And MCB = False And MCG = False Then
        Randomize
        Clrdr = (75 * Rnd + 1)
        If Clrdr > 50 Then MCR = True
        If Clrdr < 25 Then
            MCG = True
        Else
            MCB = True
        End If
    End If
    'Change currently selected color (RGB)
    If MCR Then 'RED
        If MCRd Then
            Rm = Rm + 1
            If (Rm + 1) > 254 Then
                MCRd = False
                MCR = False
            End If
        Else
            Rm = Rm - 1
            If (Rm - 1) < 50 Then
                MCRd = True
                MCR = False
            End If
        End If
    End If
    If MCG Then 'GREEN
        If MCGd Then
            Gm = Gm + 1
            If (Gm + 1) > 254 Then
                MCGd = False
                MCG = False
            End If
        Else
            Gm = Gm - 1
            If (Gm - 1) < 50 Then
                MCGd = True
                MCG = False
            End If
        End If
    End If
    If MCB Then 'BLUE
        If MCBd Then
            Bm = Bm + 1
            If (Bm + 1) > 254 Then
                MCBd = False
                MCB = False
            End If
        Else
            Bm = Bm - 1
            If (Bm - 1) < 50 Then
                MCBd = True
                MCB = False
            End If
        End If
    End If
    'Send new value as a ColorConstant
    GetColor = "&h" & CStr(Hex(Bm)) & CStr(Hex(Gm)) & CStr(Hex(Rm))
End Function


