Attribute VB_Name = "modWinFunctions"
Option Explicit

'This file was created 1/10/00
'by Shannon Little
'http://go.to/neotrix
'This file contains functions relating to the shell
'Such as window position, colors and systemmetrics

'**** Window Animation W98/2000 only
Private Declare Function AnimateWindow Lib "user32" ( _
            ByVal hWnd As Long, _
            ByVal dwTime As Long, _
            ByVal dwFlags As Long) As Long

Public Enum WindowTransition
    LeftToRight_ = &O1
    RightToLeft_ = &H2
    TopToBottom_ = &H4
    BottomToTop_ = &H8
    Hide_ = &H10000
    Activate_ = &H20000
    Blend_ = &H40000
    Slide_ = &H40000
    Center_ = &H10
End Enum
'****

'**** System Colors
Public Enum SysColorItems
    SCROLLBAR = 0
    BACKGROUND = 1
    ACTIVECAPTION = 2
    INACTIVECAPTION = 3
    Menu = 4
    WINDOW = 5
    WINDOWFRAME = 6
    MENUTEXT = 7
    WINDOWTEXT = 8
    CAPTIONTEXT = 9
    ACTIVEBORDER = 10
    INACTIVEBORDER = 11
    APPWORKSPACE = 12
    HIGHLIGHT = 13
    HIGHLIGHTTEXT = 14
    BTNFACE = 15
    BTNSHADOW = 16
    GRAYTEXT = 17
    BTNTEXT = 18
    INACTIVECAPTIONTEXT = 19
    BTNHIGHLIGHT = 20
End Enum

Private Declare Function GetSysColor Lib "user32" ( _
            ByVal nIndex As Long) As Long
'Private Declare Function SetSysColors Lib "user32" ( _
            'ByVal nChanges As Long, _
            'lpSysColor As Long, _
            'lpColorValues As Long) As Long
      
'****

Private Declare Function CreateRoundRectRgn Lib "gdi32" ( _
            ByVal x1 As Long, _
            ByVal Y1 As Long, _
            ByVal X2 As Long, _
            ByVal Y2 As Long, _
            ByVal X3 As Long, _
            ByVal Y3 As Long) As Long
        
Private Declare Function SetWindowRgn Lib "user32" ( _
            ByVal hWnd As Long, _
            ByVal hRgn As Long, _
            ByVal bRedraw As Boolean) As Long
        
Public Declare Function FlashWindow Lib "user32" ( _
            ByVal hWnd As Long, _
            ByVal bInvert As Long) As Long
            
Private Declare Function SetWindowPos Lib "user32" _
            (ByVal hWnd As Long, _
            ByVal hWndInsertAfter As Long, _
            ByVal x As Long, _
            ByVal y As Long, _
            ByVal cx As Long, _
            ByVal cy As Long, _
            ByVal wFlags As Long) As Long
'*****
Private Const LB_FINDSTRING = &H18F
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" ( _
            ByVal hWnd As Long, _
            ByVal wMsg As Long, _
            ByVal wParam As Long, _
            lParam As Any) As Long
'*****
'***** System Menu ****

'SetMenuItemInfo fState constants.
Private Const MFS_GRAYED     As Long = &H3&
Private Const MFS_CHECKED    As Long = &H8&

'SendMessage constants.
Private Const WM_NCACTIVATE  As Long = &H86

'User-defined Types.
Private Type MENUITEMINFO
    cbSize        As Long
    fMask         As Long
    fType         As Long
    fState        As Long
    wID           As Long
    hSubMenu      As Long
    hbmpChecked   As Long
    hbmpUnchecked As Long
    dwItemData    As Long
    dwTypeData    As String
    cch           As Long
End Type

'Application-specific constants and variables.
Private Const xSC_CLOSE  As Long = -10
Private Const DisableID     As Long = 1
Private Const EnableID     As Long = 2
Private Const ResetID    As Long = 3
Private MII    As MENUITEMINFO

'Menu item constants
Const SC_SIZE         As Long = &HF000&
Const SC_SEPARATOR    As Long = &HF00F&
Const SC_MOVE         As Long = &HF010&
Const SC_MINIMIZE     As Long = &HF020&
Const SC_MAXIMIZE     As Long = &HF030&
Const SC_CLOSE        As Long = &HF060&
Const SC_RESTORE      As Long = &HF120&

'SetMenuItemInfo fMask Constants
Const MIIM_STATE      As Long = &H1&
Const MIIM_ID         As Long = &H2&
Const MIIM_SUBMENU    As Long = &H4&
Const MIIM_CHECKMARKS As Long = &H8&
Const MIIM_TYPE       As Long = &H10&
Const MIIM_DATA       As Long = &H20&
       
       
Private Declare Function GetSystemMenu Lib "user32" ( _
            ByVal hWnd As Long, _
            ByVal bRevert As Long) As Long

Private Declare Function GetMenuItemInfo Lib "user32" Alias "GetMenuItemInfoA" ( _
            ByVal hMenu As Long, _
            ByVal un As Long, _
            ByVal b As Boolean, _
            lpMenuItemInfo As MENUITEMINFO) As Long

Private Declare Function SetMenuItemInfo Lib "user32" Alias "SetMenuItemInfoA" ( _
            ByVal hMenu As Long, _
            ByVal un As Long, _
            ByVal bool As Boolean, _
            lpcMenuItemInfo As MENUITEMINFO) As Long

'***** End System Menu *****
      
Private hWnd As Long
Private hMenu As Long

'Use: Locates nearest match for text in a list box
'Sub Text1_Change()
'       List1.ListIndex = SendMessage(List1.hwnd, LB_FINDSTRING, -1, _
'       ByVal CStr(Text1.Text))
'    End Sub
'*****

'Creates 1/10/00
'Copied
Public Function SetTopWindow(hWnd As Long, blnTopOrNormal As Boolean) As Long
    Dim SWP_NOMOVE
    Dim SWP_NOSIZE
    Dim FLAGS
    Dim HWND_TOPMOST
    Dim HWND_NOTOPMOST
    
    SWP_NOMOVE = 2
    SWP_NOSIZE = 1
    FLAGS = SWP_NOMOVE Or SWP_NOSIZE
    HWND_TOPMOST = -1
    HWND_NOTOPMOST = -2
    
    If blnTopOrNormal = True Then 'Make the window the topmost
        SetTopWindow = SetWindowPos(hWnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS)
    Else    'Make it normal
        SetTopWindow = SetWindowPos(hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, FLAGS)
        SetTopWindow = False
    End If
End Function

'Created 1/10/00
'Copied
Public Sub Initialize(myForm As Form)
    'All initialization for all function is done here
    hWnd = myForm.hWnd
    hMenu = GetSystemMenu(myForm.hWnd, 0)
    Dim Ret As Long

    '**** System Menu *****
    MII.cbSize = Len(MII)
    MII.dwTypeData = String(80, 0)
    MII.cch = Len(MII.dwTypeData)
    MII.fMask = MIIM_STATE
    MII.wID = SC_CLOSE
    Ret = GetMenuItemInfo(hMenu, MII.wID, False, MII)
    '**** End System Menu ****
End Sub

'Created 1/10/00
'Copied
'***** All part of changing the system menu *****
Public Sub DiableCloseMenu()
    Dim Ret As Long

    Ret = SetId(DisableID)
    If Ret <> 0 Then
        'If its is not already disabled then disable it
        If MII.fState <> MFS_GRAYED Then
            MII.fState = MFS_GRAYED
            MII.fMask = MIIM_STATE
            Ret = SetMenuItemInfo(hMenu, MII.wID, False, MII)
            If Ret = 0 Then
                Ret = SetId(ResetID)
            End If
    
            Ret = SendMessage(hWnd, WM_NCACTIVATE, True, 0)
        End If
    End If
End Sub

'Created 1/10/00
'Copied
Public Sub EnableCloseMenu()
    Dim Ret As Long

    Ret = SetId(EnableID)
    If Ret <> 0 Then
        If MII.fState = MFS_GRAYED Then     'Its already disabled so enable it
            MII.fState = MII.fState - MFS_GRAYED    'Enable

            MII.fMask = MIIM_STATE
            Ret = SetMenuItemInfo(hMenu, MII.wID, False, MII)
            If Ret = 0 Then
                Ret = SetId(ResetID)
            End If
            'Send message that windows need to repaint the non-client area (sys menu)
            Ret = SendMessage(hWnd, WM_NCACTIVATE, True, 0)
        End If
    End If
End Sub

'Created 1/10/00
'Copied
Private Function SetId(Action As Long) As Long
    Dim MenuID As Long
    Dim Ret As Long

    MenuID = MII.wID
    If MII.fState = (MII.fState Or MFS_GRAYED) Then 'If its disabled
        If Action = EnableID Then                  'And the action is to enabled, enabled
            MII.wID = xSC_CLOSE
        End If
        'If the action was to disabled, then do nothing
    Else
        If Action = DisableID Then
            MII.wID = xSC_CLOSE
        End If
    End If

    MII.fMask = MIIM_ID
    Ret = SetMenuItemInfo(hMenu, MenuID, False, MII)
    If Ret = 0 Then
        MII.wID = MenuID
    End If
    SetId = Ret
    
'   'User-defined Types.
'Private Type MENUITEMINFO
    'cbSize        As Long
    'fMask         As Long
    'fType         As Long
    'fState        As Long
    'wID           As Long
    'hSubMenu      As Long
    'hbmpChecked   As Long
    'hbmpUnchecked As Long
    'dwItemData    As Long
    'dwTypeData    As String
    'cch           As Long
'End Type
End Function

'**** End system menu change *****

'Created 2/22/00
'Adds rounded edges to a window
Public Sub MakeWindowEdgesRound(pForm As Form, lngValue As Long)
    Dim lngRet As Long
    Dim lng As Long
    Dim lngWidth As Long
    Dim lngHeight As Long
            
    'Get Form size in pixels
    lngWidth = pForm.Width / Screen.TwipsPerPixelX
    lngHeight = pForm.Height / Screen.TwipsPerPixelY
    
    'Create Form with Rounded Corners
    lngRet = CreateRoundRectRgn(0, 0, lngWidth, lngHeight, lngValue, lngValue)
                              
    lng = SetWindowRgn(pForm.hWnd, lngRet, True)
End Sub


'Created 2/22/00
'Returns the color of a system item
Public Function GetSystemColor(ColorItem As SysColorItems)
    GetSystemColor = GetSysColor(ColorItem)
End Function

'Sets the color of a system item
'Created 2/22/00
'Public Function SetSystemColor(SysColorItems)
    'The first parameter indicates the total number of system colors you are attempting to change.
    'The second parameter is an array of the numeric values for the display aspects you want to change.
    'The third parameter is also an array whose elements are the new colors for the display aspects defined by the first array
    
'End Function

'Created 2/22/00
'Animates the window
Public Sub AnimateWindowOpening(pForm As Form, Trans As WindowTransition, Optional Speed As Long)
    Dim lngSpeed As Long
    
    If IsMissing(Speed) Then
        lngSpeed = 1000
    Else
        lngSpeed = Speed
    End If
    
    AnimateWindow pForm.hWnd, lngSpeed, Trans
    pForm.Refresh
End Sub

