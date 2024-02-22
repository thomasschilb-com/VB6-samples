Attribute VB_Name = "SubMain"
Option Explicit

' Sine Function Line Screen Saver

' Created by:
' Rod Stephens (Core functionality)
' Neil Fraser (Integration & Win 98+ multi-monitors)
' Elliot Spencer (Win 9x locking & registry lookup)
' Don Bradner & Jim Deutch (Password handling)
' Lucian Wischik & Alex Millman (NT information)
' Ken Slater - 0x34 - (Grafix and Motion Code)

Type RECT
  Left As Long
  Top As Long
  Right As Long
  Bottom As Long
End Type

Public Const SWP_NOACTIVATE = &H10
Public Const SWP_NOZORDER = &H4
Public Const SWP_SHOWWINDOW = &H40

Public Const HWND_TOP = 0

Public Const WS_CHILD = &H40000000
Public Const GWL_HWNDPARENT = (-8)
Public Const GWL_STYLE = (-16)

Public Const SPI_SCREENSAVERRUNNING = 97&
Public Const HKEY_CURRENT_USER = &H80000001
Public Const REG_DWORD As Long = 4

'Virtual Desktop sizes
Public Const SM_XVIRTUALSCREEN = 76    'Virtual Left
Public Const SM_YVIRTUALSCREEN = 77    'Virtual Top
Public Const SM_CXVIRTUALSCREEN = 78   'Virtual Width
Public Const SM_CYVIRTUALSCREEN = 79   'Virtual Height
Public Const SM_CMONITORS = 80         'Get number of monitors

Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long
Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Declare Function PwdChangePassword Lib "mpr" Alias "PwdChangePasswordA" (ByVal lpcRegkeyname As String, ByVal hwnd As Long, ByVal uiReserved1 As Long, ByVal uiReserved2 As Long) As Long
Declare Function VerifyScreenSavePwd Lib "password.cpl" (ByVal hwnd As Long) As Boolean
Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, lpvParam As Any, ByVal fuWinIni As Long) As Long

Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByVal lpData As String, lpcbData As Long) As Long
Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long

' Global variables.
Public PreviewMode As Boolean
Public Density As Integer

' Private variables.
Private Const APP_NAME_RUNNING = "Running Screen Saver"
Private Const APP_NAME_PREVIEW = "Preview Screen Saver"

' Load configuration information from the registry.
Public Sub LoadConfig()
  Density = CInt(GetSetting(App.EXEName + ".scr", "Settings", "Density", "50"))
End Sub

' Save configuration information to the registry.
Public Sub SaveConfig()
  SaveSetting App.EXEName + ".scr", "Settings", "Density", Format$(Density)
End Sub

' See if another instance of the program is
' running in screen saver mode.
Private Sub CheckShouldRun()
  ' If no instance is running, we're safe.
  If Not App.PrevInstance Then Exit Sub

  ' See if there is a screen saver mode instance.
  If FindWindow(vbNullString, APP_NAME_RUNNING) Then End
End Sub

' Get the hWnd for the preview window from the
' command line arguments.
Private Function GetHwndFromCommand(ByVal args As String) As Long
Dim argslen As Integer
Dim i As Integer
Dim ch As String

  ' Take the rightmost numeric characters.
  args = Trim$(args)
  argslen = Len(args)
  For i = argslen To 1 Step -1
    ch = Mid$(args, i, 1)
    If ch < "0" Or ch > "9" Then Exit For
  Next i

  GetHwndFromCommand = CLng(Mid$(args, i + 1))
End Function

' Start the program.
Sub Main()
Dim args As String
Dim preview_hwnd As Long
Dim preview_rect As RECT
Dim window_style As Long

  args = UCase$(Trim$(Command$))

  ' examine the second character from the commandline arguments
  Select Case Mid$(args, 2, 1)
  
  ' the command-line argument /c is used to launch the screen saver's config panel
  ' the context menu on a .scr file launches configuration mode with no arguments
  Case "C", ""
    FormConfig.Show
  
  ' the command-line argument /s is used to launch the screen saver in normal operating mode
  Case "S"
    PreviewMode = False
    ' Make sure there isn't another one running.
    CheckShouldRun

    ' Display the cover form.
    Load FormDisplay
    ' Set the caption for Windows 95.
    FormDisplay.Caption = APP_NAME_RUNNING
    
    FormDisplay.Show
    
    ' if there are multi-monitors, span across them
    If GetSystemMetrics(SM_CMONITORS) <> 0 Then
      FormDisplay.WindowState = vbNormal
      FormDisplay.Top = GetSystemMetrics(SM_YVIRTUALSCREEN) * Screen.TwipsPerPixelY
      FormDisplay.Left = GetSystemMetrics(SM_XVIRTUALSCREEN) * Screen.TwipsPerPixelX
      FormDisplay.Width = GetSystemMetrics(SM_CXVIRTUALSCREEN) * Screen.TwipsPerPixelX
      FormDisplay.Height = GetSystemMetrics(SM_CYVIRTUALSCREEN) * Screen.TwipsPerPixelY
    End If

    ' make this form topmost
    Const SWP_NOMOVE = 2
    Const SWP_NOSIZE = 1
    Const HWND_TOPMOST = -1
    SetWindowPos FormDisplay.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
    
  ' the command-line argument /p is used to launch the preview box
  Case "P", "L"
    PreviewMode = True
    ' Get the preview area hWnd.
    preview_hwnd = GetHwndFromCommand(args)

    ' Get the dimensions of the preview area.
    GetClientRect preview_hwnd, preview_rect

    Load FormDisplay

    ' Set the caption for Windows 95.
    FormDisplay.Caption = APP_NAME_PREVIEW

    ' Get the current window style.
    window_style = GetWindowLong(FormDisplay.hwnd, GWL_STYLE)

    ' Add WS_CHILD to make this a child window.
    window_style = (window_style Or WS_CHILD)

    ' Set the window's new style.
    SetWindowLong FormDisplay.hwnd, GWL_STYLE, window_style

    ' Set the window's parent so it appears
    ' inside the preview area.
    SetParent FormDisplay.hwnd, preview_hwnd

    ' Save the preview area's hWnd in
    ' the form's window structure.
    SetWindowLong FormDisplay.hwnd, GWL_HWNDPARENT, preview_hwnd

    ' Show the preview.
    SetWindowPos FormDisplay.hwnd, HWND_TOP, 0&, 0&, _
      preview_rect.Right, preview_rect.Bottom, _
      SWP_NOZORDER Or SWP_NOACTIVATE Or SWP_SHOWWINDOW
  
  ' the command-line argument /a is used to launch the set password box
  Case "A"
    ' get the preview area hWnd
    preview_hwnd = GetHwndFromCommand(args)
    ' tell Windows to open the set password window
    PwdChangePassword "SCRSAVE", preview_hwnd, 0, 0
  
  Case Else ' this shouldn't happen
    MsgBox "Unknown command-line arguments: [" + Command$ + "]", vbCritical
  
  End Select
End Sub

' Generic registry access
Public Function ReadRegistry(ByVal Group As Long, ByVal Section As String, ByVal Key As String) As String
  Dim lResult As Long, lKeyValue As Long, lDataTypeValue As Long, lValueLength As Long, sValue As String, td As Double
  On Error GoTo ErrorHandler
  lResult = RegOpenKey(Group, Section, lKeyValue)
  sValue = Space$(2048)
  lValueLength = Len(sValue)
  lResult = RegQueryValueEx(lKeyValue, Key, 0&, lDataTypeValue, sValue, lValueLength)
  If (lResult = 0) And (Err.Number = 0) Then
    If lDataTypeValue = REG_DWORD Then
      td = Asc(Mid$(sValue, 1, 1)) + &H100& * Asc(Mid$(sValue, 2, 1)) + &H10000 * Asc(Mid$(sValue, 3, 1)) + &H1000000 * CDbl(Asc(Mid$(sValue, 4, 1)))
      sValue = Format$(td, "000")
    End If
    sValue = Left$(sValue, lValueLength - 1)
  Else
    sValue = "Not Found"
  End If
  lResult = RegCloseKey(lKeyValue)
  ReadRegistry = sValue
  On Error GoTo 0
Exit Function

ErrorHandler:
  ' Don't know/care what happened.  Maybe perms problem in NT?
  ReadRegistry = "Not Found"
End Function
