Attribute VB_Name = "Module1"
Public Declare Function LockWindowUpdate Lib "user32" (ByVal hWndLock As Long) As Long

Public Type FntSettings
    FontName As String
    FontSize As Currency
    FontColor As OLE_COLOR
    FontBold As Boolean
    FontItalic As Boolean
    FontStrikethru As Boolean
    FontUnderline As Boolean
End Type

Public MyFont As FntSettings


Public AppEnd As Boolean

    Declare Function DrawEdge Lib "user32" (ByVal hDC As Long, qrc As RECT, ByVal edge As Long, ByVal grfFlags As Long) As Long
    Declare Function CopyRect Lib "user32" (lpDestRect As RECT, lpSourceRect As RECT) As Long
    Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
    Declare Function Rectangle Lib "gdi32" (ByVal hDC As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Public Const BDR_INNER = &HC
Public Const BDR_OUTER = &H3
Public Const BDR_RAISED = &H5
Public Const BDR_RAISEDINNER = &H4
Public Const BDR_RAISEDOUTER = &H1
Public Const BDR_SUNKEN = &HA
Public Const BDR_SUNKENINNER = &H8
Public Const BDR_SUNKENOUTER = &H2
Public Const BF_ADJUST = &H2000
Public Const BF_BOTTOM = &H8
Public Const BF_DIAGONAL = &H10
Public Const BF_FLAT = &H4000
Public Const BF_LEFT = &H1
Public Const BF_MIDDLE = &H800
Public Const BF_MONO = &H8000
Public Const BF_RIGHT = &H4
Public Const BF_SOFT = &H1000
Public Const BF_TOP = &H2
Public Const BF_TOPLEFT = (BF_TOP Or BF_LEFT)
Public Const BF_TOPRIGHT = (BF_TOP Or BF_RIGHT)
Public Const BF_RECT = (BF_LEFT Or BF_TOP Or BF_RIGHT Or BF_BOTTOM)
Public Const BF_BOTTOMLEFT = (BF_BOTTOM Or BF_LEFT)
Public Const BF_BOTTOMRIGHT = (BF_BOTTOM Or BF_RIGHT)
Public Const BF_DIAGONAL_ENDBOTTOMLEFT = (BF_DIAGONAL Or BF_BOTTOM Or BF_LEFT)
Public Const BF_DIAGONAL_ENDBOTTOMRIGHT = (BF_DIAGONAL Or BF_BOTTOM Or BF_RIGHT)
Public Const BF_DIAGONAL_ENDTOPLEFT = (BF_DIAGONAL Or BF_TOP Or BF_LEFT)
Public Const BF_DIAGONAL_ENDTOPRIGHT = (BF_DIAGONAL Or BF_TOP Or BF_RIGHT)

