Attribute VB_Name = "Module1"
Option Explicit

'The rotation style
Public Enum Style
    Circles1 = 0  'Circles rotate and move in 3D motion
    Circles2 = 1  'Circles rotate in 3D motion but move in 2D motion.
    Lines1 = 2    'Lines rotate and move in 3D motion
    Lines2 = 3    'Lines rotate in 3D motion but move in 2D motion.
End Enum

'Global Variables
Public ObjectStyle As Style
Public NumOfPoints As Integer
Public FadingEffect As Boolean
Public CirclesFillColor As Boolean

