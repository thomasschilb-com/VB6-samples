VERSION 5.00
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "3D Motion"
   ClientHeight    =   4815
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6600
   FillStyle       =   0  'Solid
   ForeColor       =   &H8000000F&
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4815
   ScaleWidth      =   6600
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox PicFade 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00000000&
      Height          =   3015
      Left            =   720
      ScaleHeight     =   3015
      ScaleWidth      =   4815
      TabIndex        =   0
      Top             =   960
      Visible         =   0   'False
      Width           =   4815
   End
   Begin VB.Timer TmrRnd 
      Interval        =   3000
      Left            =   840
      Top             =   360
   End
   Begin VB.Timer TmrDraw 
      Interval        =   1
      Left            =   240
      Top             =   360
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'For the fading effect
Private Declare Function AlphaBlend Lib "msimg32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal widthSrc As Long, ByVal heightSrc As Long, ByVal blendFunct As Long) As Boolean
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSrc As Any, ByVal ByteLen As Long)
Private Type BLENDFUNCTION
    BlendOp As Byte
    BlendFlags As Byte
    SourceConstantAlpha As Byte
    AlphaFormat As Byte
End Type

'Global variables in this form
Dim SizeW As Single, SizeH As Single        'The size of the object

Dim X As Long, Y As Long                    'For position
Dim MovVelX As Single, MovVelY As Single    'The movement velocity
Dim MovXAcc As Single, MovYAcc As Single    'The movement Acceleration

Dim Start As Single                         'The starting angle for spining
Dim SpinVel As Single                       'Spining Velocity
Dim SpinAcc As Single                       'Spining Acceleration

Dim RotX As Single, RotY As Single          'For rotation
Dim RotXVel As Single, RotYVel As Single    'Rotations Velocities
Dim RotXAcc As Single, RotYAcc As Single    'Rotations Accelerations

Dim ColorR As Integer, ColorG As Integer, ColorB As Integer   'Colors variables
Dim IncrColorR As Integer, IncrColorG As Integer, IncrColorB As Integer  'Increment colors


Private Sub Form_Click()
    Unload Me
End Sub
    
Private Sub Form_Load()
    
    Randomize
    
    'Size of the big circle
    SizeW = Screen.Width / 8
    SizeH = Screen.Height / 8
    
    'For the rotations
    RotX = 0
    RotY = 0
    
    'For the spining ( start with no velocity )
    Start = 0
    
    'For moving Velocity
    MovVelX = 50
    MovVelY = 20
    
    'Initial colors
    ColorR = 100
    ColorG = 200
    ColorB = 50
    
    'Randome numbers
    TmrRnd_Timer
    
    'Full Screen
    Me.Width = Screen.Width
    Me.Height = Screen.Height
    
    'The coordinates of the screan's center
    X = Me.ScaleWidth / 2
    Y = Me.ScaleHeight / 2

End Sub

Private Sub Form_Resize()
    'For the fading effect
    If FadingEffect Then PicFade.Move 0, 0, ScaleWidth, ScaleHeight
End Sub

Private Sub PicFade_Click()
    Unload Me
End Sub

Private Sub TmrDraw_Timer()
    'Clear screen or apply fading effect
    Call ClearScreen
    
    'Set new values for positions variables resposible for movement
    Call SetNewPosition
    
    'Set new values for velocites variables responsible for rotations
    Call SetNewVelocities
    
    'Draw the object
    Call DrawObject
End Sub

Private Sub TmrRnd_Timer()
    
    'Randome values
    
    'The movement acceleration on x and y ranged from -1 to 1
    MovXAcc = Rnd * 2 - 1
    MovYAcc = Rnd * 2 - 1
    
    'The spining acceleration ranged from -.0005 to .0005
    SpinAcc = Rnd * 0.001 - 0.0005
    
    'The rotation acceleration on x and y axis ranged from -.0005 to .0005
    'This also affect the acceleration on z direction (Prepindecular to the screen)
    RotXAcc = Rnd * 0.001 - 0.0005
    RotYAcc = Rnd * 0.001 - 0.0005
    
    'For the color rotation ranged from 0 to 2
    IncrColorR = Rnd * 2
    IncrColorG = Rnd * 2
    IncrColorB = Rnd * 2
    
End Sub

Private Sub ClearScreen()
    
    If FadingEffect Then
        'Varialbes
        Dim Blend As BLENDFUNCTION
        Dim BlendPtr As Long
        
        'The fading effect is affected with the brightness of the colors
        Blend.SourceConstantAlpha = (255 * 3 - (ColorR + ColorG + ColorB)) / 3 '20
        
        'Fix fading
        'Blend.SourceConstantAlpha = 120
        
        
        'The fading effect is affected with the spinning Velocity
        'Blend.SourceConstantAlpha = 255 - Abs(SpinVel * (250 / 0.05))
        
        CopyMemory BlendPtr, Blend, 4
        AlphaBlend hDC, 0, 0, ScaleWidth / 15, ScaleHeight / 15, PicFade.hDC, 0, 0, ScaleWidth / 15, ScaleHeight / 15, BlendPtr
    Else
        'Clear Screen
        Me.Cls
    End If
End Sub


Private Sub SetNewPosition()
    Static IncreaseX As Boolean, IncreaseY As Boolean
    Dim Tolerance As Integer
    
    MovVelX = MovVelX + MovXAcc
    MovVelY = MovVelY + MovYAcc
    
    'Check the Vel range -50 to 50
    If Abs(MovVelX) > 50 Then
        MovXAcc = -MovXAcc
        MovVelX = MovVelX + MovXAcc
    End If
    
    If Abs(MovVelY) > 50 Then
        MovYAcc = -MovYAcc
        MovVelY = MovVelY + MovYAcc
    End If
    
    'The least distance possible between the object and the walls
    Tolerance = 200
    
    'When hit the wall on the x-axis, bounce back
    If X >= Me.ScaleWidth - Int(SizeW) - Tolerance Then
        IncreaseX = False
        X = Me.ScaleWidth - Int(SizeW) - Tolerance
    ElseIf X <= 0 + Int(SizeW) + Tolerance Then
        IncreaseX = True
        X = Int(SizeW) + Tolerance
    End If
    
    If IncreaseX Then
        X = X + MovVelX
    Else
        X = X - MovVelX
    End If
    
    'When hit the wall on the y-axis, bounce back
    If Y >= Me.ScaleHeight - Int(SizeH) - Tolerance Then
        IncreaseY = False
        Y = Me.ScaleHeight - Int(SizeH) - Tolerance
    ElseIf Y <= 0 + Int(SizeH) + Tolerance Then
        IncreaseY = True
        Y = Int(SizeH) + Tolerance
    End If
    
    If IncreaseY Then
        Y = Y + MovVelY
    Else
        Y = Y - MovVelY
    End If
End Sub

Private Sub SetNewVelocities()
    
    'Velocity of the spinning
    SpinVel = SpinAcc + SpinAcc
    'Limit the Velocity from -.1 to .1
    If Abs(SpinVel) > 0.1 Then
        SpinAcc = -SpinAcc
        SpinVel = SpinVel + SpinAcc
    End If
    
    Start = Start + SpinVel
    
    
    'Velocity of the the second rotation on the x-axis
    RotXVel = RotXVel + RotXAcc
    'Limit the Velocity from -.05 to .05
    If Abs(RotXVel) > 0.05 Then
        RotXAcc = -RotXAcc
        RotXVel = RotXVel + RotXAcc
    End If
    RotX = RotX + RotXVel
    
    
    'Velocity of the the rotation on the y-asix
    RotYVel = RotYVel + RotYAcc
    'Limit the Velocity from -.05 to .05
    If Abs(RotYVel) > 0.05 Then
        RotYAcc = -RotYAcc
        RotYVel = RotYVel + RotYAcc
    End If
    RotY = RotY + RotYVel
    
End Sub
Private Sub DrawObject()
    
    Dim I As Single
    Dim Stp As Single
    Dim pColor As Long
    Dim DPi As Single
    
    '2 x Pi
    DPi = 6.28318530717959
        
    'Create randome color
    pColor = GetNewColor()
    
    'Set the fill color of the form
    If CirclesFillColor Then Me.FillColor = pColor
    
    'Draw circles
    Stp = DPi / NumOfPoints
    
    For I = Start To Start + DPi Step Stp
    
        Select Case ObjectStyle
            Case Circles1
                'Draw circles
                'Rotate and move in 3D motion
                Circle (X + SizeW * RotXVel * 20 * Cos(I + RotX), Y + RotYVel * 20 * SizeH * Sin(I + RotY)), 10 + 800 * (RotXVel ^ 2 + RotYVel ^ 2) ^ 0.5, pColor
            Case Circles2
                'Draw circles
                'Rotate in 3D motion, but move in 2D
                Circle (X + SizeW * Cos(I + RotX), Y + SizeH * Sin(I + RotY)), 50 + 30 * (Sin(I) + Cos(I)), pColor
            Case Lines1
                'Draw Lines
                'Rotate and move in 3D motion
                Line (X + SizeW * RotXVel * 20 * Cos(I + RotX), Y + SizeH * Sin(I + RotY))-(X + SizeW * RotXVel * 20 * Cos(I + Stp * Int(NumOfPoints / 4) + RotX), Y + SizeH * Sin(I + Stp * Int(NumOfPoints / 4) + RotY)), pColor

            Case Lines2
                'Draw Lines
                'Rotate in 3D motion, but move in 2D
                Line (X + SizeW * Cos(I + RotX), Y + RotYVel * 20 * SizeH * Sin(I + RotY))-(X + SizeW * Cos(I + Stp * Int(NumOfPoints / 4) + RotX), Y + RotYVel * 20 * SizeH * Sin(I + Stp * Int(NumOfPoints / 4) + RotY)), pColor
            Case Else
                'Interesting patterns
                Circle (X + SizeW * Cos(I * RotX), Y + SizeH * Sin(I * RotY)), 50 + 30 * (Sin(I) + Cos(I)), pColor
        End Select
    Next I
End Sub

Private Function GetNewColor() As Long
    'Color Rotations
    ColorR = ColorR + IncrColorR
    ColorG = ColorG + IncrColorG
    ColorB = ColorB + IncrColorB
    
    'Check color limits 0-255
    If ColorR > 255 Or ColorR < 0 Then
        IncrColorR = -IncrColorR
        ColorR = ColorR + IncrColorR
    End If
    
    If ColorG > 255 Or ColorG < 0 Then
        IncrColorG = -IncrColorG
        ColorG = ColorG + IncrColorG
    End If
    
    If ColorB > 255 Or ColorB < 0 Then
        IncrColorB = -IncrColorB
        ColorB = ColorB + IncrColorB
    End If
    
    'Retrun the value of the new color
    GetNewColor = RGB(ColorR, ColorG, ColorB)
End Function
