Attribute VB_Name = "modSysMenu"
Option Explicit

Public Declare Function GetSystemMenu Lib "user32.dll" (ByVal Hwnd As Long, ByVal bRevert _
    As Long) As Long
Public Declare Function GetMenuItemCount Lib "user32.dll" (ByVal hMenu As Long) As Long
Public Type MENUITEMINFO
    cbSize As Long
    fMask As Long
    fType As Long
    fState As Long
    wID As Long
    hSubMenu As Long
    hbmpChecked As Long
    hbmpUnchecked As Long
    dwItemData As Long
    dwTypeData As String
    cch As Long
End Type
Public Const MIIM_STATE = &H1
Public Const MIIM_ID = &H2
Public Const MIIM_TYPE = &H10
Public Const MFT_SEPARATOR = &H800
Public Const MFT_STRING = &H0
Public Const MFS_ENABLED = &H0
Public Const MFS_CHECKED = &H8
Public Declare Function InsertMenuItem Lib "user32.dll" Alias "InsertMenuItemA" (ByVal _
    hMenu As Long, ByVal uItem As Long, ByVal fByPosition As Long, lpmii _
    As MENUITEMINFO) As Long
Public Declare Function SetMenuItemInfo Lib "user32.dll" Alias "SetMenuItemInfoA" (ByVal _
    hMenu As Long, ByVal uItem As Long, ByVal fByPosition As Long, lpmii _
    As MENUITEMINFO) As Long

Public Declare Function SetWindowPos Lib "user32.dll" (ByVal Hwnd As Long, ByVal _
    hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal _
    cy As Long, ByVal wFlags As Long) As Long
Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOSIZE = &H1
Public Declare Function SetWindowLong Lib "user32.dll" Alias "SetWindowLongA" (ByVal Hwnd _
    As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Const GWL_WNDPROC = -4
Public Declare Function CallWindowProc Lib "user32.dll" Alias "CallWindowProcA" (ByVal _
    lpPrevWndFunc As Long, ByVal Hwnd As Long, ByVal Msg As Long, ByVal wParam _
    As Long, ByVal lParam As Long) As Long
Public Const WM_SYSCOMMAND = &H112
Public Const WM_INITMENU = &H116

Public pOldProc As Long
Public ontop As Boolean

Public Function WindowProc(ByVal Hwnd As Long, ByVal uMsg As Long, ByVal wParam _
        As Long, ByVal lParam As Long) As Long
    Dim hSysMenu As Long
    Dim mii As MENUITEMINFO
    Dim retval As Long
    
    Select Case uMsg
    Case WM_INITMENU
        hSysMenu = GetSystemMenu(Hwnd, 0)
        With mii
            .cbSize = Len(mii)
            .fMask = MIIM_STATE
            .fState = MFS_ENABLED Or IIf(ontop, MFS_CHECKED, 0)
        End With
        retval = SetMenuItemInfo(hSysMenu, 1, 0, mii)
        WindowProc = 0
    Case WM_SYSCOMMAND
        If wParam = 1 Then
            ontop = Not ontop
            If ontop = True Then
                SetWindowPos Hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
            Else
                SetWindowPos Hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
            End If
            WindowProc = 0
        Else
            WindowProc = CallWindowProc(pOldProc, Hwnd, uMsg, wParam, lParam)
        End If
    Case Else
        WindowProc = CallWindowProc(pOldProc, Hwnd, uMsg, wParam, lParam)
    End Select
End Function



Public Function LoadSysMenu(Hwnd As Long)
    Dim hSysMenu As Long
    Dim count As Long
    Dim mii As MENUITEMINFO
    Dim retval As Long
    
    hSysMenu = GetSystemMenu(Hwnd, 0)
    count = GetMenuItemCount(hSysMenu)
    
    With mii
        .cbSize = Len(mii)
        .fMask = MIIM_ID Or MIIM_TYPE
        .fType = MFT_SEPARATOR
        .wID = 0
    End With
    retval = InsertMenuItem(hSysMenu, count, 1, mii)
    
    With mii
        .fMask = MIIM_STATE Or MIIM_ID Or MIIM_TYPE
        .fType = MFT_STRING
        .fState = MFS_ENABLED
        .wID = 1
        .dwTypeData = "Window on &Top"
        .cch = Len(.dwTypeData)
    End With
    retval = InsertMenuItem(hSysMenu, count + 1, 1, mii)
    
    ontop = False
    pOldProc = SetWindowLong(Hwnd, GWL_WNDPROC, AddressOf WindowProc)
End Function

Public Sub UnloadSysMenu(Hwnd As Long)
    Dim retval As Long
    
    retval = SetWindowLong(Hwnd, GWL_WNDPROC, pOldProc)
    retval = GetSystemMenu(Hwnd, 1)

End Sub

