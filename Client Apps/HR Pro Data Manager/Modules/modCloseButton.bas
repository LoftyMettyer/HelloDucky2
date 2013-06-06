Attribute VB_Name = "modCloseButton"
Option Explicit

Private Const SC_CLOSE As Long = &HF060&
Private Const SC_MAXIMIZE As Long = &HF030&
Private Const SC_MINIMIZE As Long = &HF020&

Private Const xSC_CLOSE As Long = -10&
Private Const xSC_MAXIMIZE As Long = -11&
Private Const xSC_MINIMIZE As Long = -12&

Private Const GWL_STYLE = (-16)
Private Const WS_MAXIMIZEBOX = &H10000
Private Const WS_MINIMIZEBOX = &H20000

Private Const hWnd_NOTOPMOST = -2
Private Const SWP_NOZORDER = &H4
Private Const SWP_NOSIZE = &H1
Private Const SWP_NOMOVE = &H2
Private Const SWP_FRAMECHANGED = &H20

Private Const MIIM_STATE As Long = &H1&
Private Const MIIM_ID As Long = &H2&
Private Const MFS_GRAYED As Long = &H3&
Private Const WM_NCACTIVATE As Long = &H86

Private Type MENUITEMINFO
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

Private Declare Function GetSystemMenu Lib "user32" ( _
    ByVal hWnd As Long, ByVal bRevert As Long) As Long

Private Declare Function GetMenuItemInfo Lib "user32" Alias _
    "GetMenuItemInfoA" (ByVal hMenu As Long, ByVal un As Long, _
    ByVal b As Boolean, lpMenuItemInfo As MENUITEMINFO) As Long

Private Declare Function SetMenuItemInfo Lib "user32" Alias _
    "SetMenuItemInfoA" (ByVal hMenu As Long, ByVal un As Long, _
    ByVal bool As Boolean, lpcMenuItemInfo As MENUITEMINFO) As Long

Private Declare Function SendMessage Lib "user32" Alias _
    "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, _
    ByVal wParam As Long, lParam As Any) As Long

Private Declare Function IsWindow Lib "user32" _
    (ByVal hWnd As Long) As Long

Private Declare Function GetWindowLong Lib "user32" Alias _
    "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
       
Private Declare Function GetParent Lib "user32" (ByVal hWnd As Long) _
    As Long
    
Private Declare Function SetParent Lib "user32" ( _
    ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
    
Private Declare Function SetWindowPos Lib "user32" ( _
    ByVal hWnd As Long, ByVal hWndInsertAfter As Long, _
    ByVal x As Long, ByVal y As Long, ByVal cx As Long, _
    ByVal cy As Long, ByVal wFlags As Long) As Long

'*******************************************************************************
' Enables / Disables the close button on the titlebar and in the system menu
' of the form window passed.
'-------------------------------------------------------------------------------
' Return Values:
'
'    0  Close button state changed succesfully / nothing to do.
'   -1  Invalid Window Handle (hWnd argument) Passed to the function
'   -2  Failed to switch command ID of Close menu item in system menu
'   -3  Failed to switch enabled state of Close menu item in system menu
'
'-------------------------------------------------------------------------------
' Parameters:
'
'   hWnd    The window handle of the form whose close button is to be enabled/
'           disabled / greyed out.
'
'   Enable  True if the close button is to be enabled, or False if it is to
'           be disabled / greyed out.
'
'-------------------------------------------------------------------------------
' Example:
'
' Add a form window to your project, and place a button on the form. Add the
' following in the form's code window:
'
'    Option Explicit
'
'    Private m_blnCloseEnabled As Boolean
'
'    Private Sub Form_Load()
'        m_blnCloseEnabled = True
'        Command1.Caption = "Disable"
'    End Sub
'
'    Private Sub Command1_Click()
'        m_blnCloseEnabled = Not m_blnCloseEnabled
'        EnableCloseButton Me.hWnd, m_blnCloseEnabled
'
'        If m_blnCloseEnabled Then
'            Command1.Caption = "Disable"
'        Else
'            Command1.Caption = "Enable"
'        End If
'    End Sub
'
'-------------------------------------------------------------------------------

Public Function EnableCloseButton(ByVal hWnd As Long, Enable As Boolean) _
                                                                As Integer

    EnableSystemMenuItem hWnd, SC_CLOSE, xSC_CLOSE, Enable, "EnableCloseButton"
End Function

'*******************************************************************************
' Enable / Disable Minimise Button
'-------------------------------------------------------------------------------

Public Sub EnableMinButton(ByVal hWnd As Long, Enable As Boolean)

    ' Enable / Disable System Menu Item

    EnableSystemMenuItem hWnd, SC_MINIMIZE, xSC_MINIMIZE, Enable, _
                                                    "EnableMinButton"
                                                    
    ' Enable / Disable TitleBar button
    
    Dim lngFormStyle As Long
    lngFormStyle = GetWindowLong(hWnd, GWL_STYLE)
    If Enable Then
        lngFormStyle = lngFormStyle Or WS_MINIMIZEBOX
    Else
        lngFormStyle = lngFormStyle And Not WS_MINIMIZEBOX
    End If
    SetWindowLong hWnd, GWL_STYLE, lngFormStyle
        
    ' Dirty, slimy, devious hack to ensure that the changes to the
    ' window's style take immediate effect before the form is shown
    
    SetParent hWnd, GetParent(hWnd)
    SetWindowPos hWnd, hWnd_NOTOPMOST, 0, 0, 0, 0, _
            SWP_NOZORDER Or SWP_NOSIZE Or SWP_NOMOVE Or SWP_FRAMECHANGED
End Sub

'*******************************************************************************
' Enable / Disable Maximise Button
'-------------------------------------------------------------------------------

Public Sub EnableMaxButton(ByVal hWnd As Long, Enable As Boolean)

    ' Enable / Disable System Menu Item

    EnableSystemMenuItem hWnd, SC_MAXIMIZE, xSC_MAXIMIZE, Enable, _
                                                    "EnableMaxButton"
                                                    
    ' Enable / Disable TitleBar button
    
    Dim lngFormStyle As Long
    lngFormStyle = GetWindowLong(hWnd, GWL_STYLE)
    If Enable Then
        lngFormStyle = lngFormStyle Or WS_MAXIMIZEBOX
    Else
        lngFormStyle = lngFormStyle And Not WS_MAXIMIZEBOX
    End If
    SetWindowLong hWnd, GWL_STYLE, lngFormStyle
        
    ' Dirty, slimy, devious hack to ensure that the changes to the
    ' window's style take immediate effect before the form is shown
    
    SetParent hWnd, GetParent(hWnd)
    SetWindowPos hWnd, hWnd_NOTOPMOST, 0, 0, 0, 0, _
            SWP_NOZORDER Or SWP_NOSIZE Or SWP_NOMOVE Or SWP_FRAMECHANGED
End Sub

'*******************************************************************************
'
'-------------------------------------------------------------------------------

Private Sub EnableSystemMenuItem(hWnd As Long, Item As Long, _
                    Dummy As Long, Enable As Boolean, FuncName As String)
    
    If IsWindow(hWnd) = 0 Then
        Err.Raise vbObjectError, "modCloseBtn::" & FuncName, _
            "modCloseBtn::" & FuncName & "() - Invalid Window Handle"
        Exit Sub
    End If
    
    ' Retrieve a handle to the window's system menu
    
    Dim hMenu As Long
    hMenu = GetSystemMenu(hWnd, 0)
    
    ' Retrieve the menu item information for the Max menu item/button
    
    Dim MII As MENUITEMINFO
    MII.cbSize = Len(MII)
    MII.dwTypeData = String$(80, 0)
    MII.cch = Len(MII.dwTypeData)
    MII.fMask = MIIM_STATE
    
    If Enable Then
        MII.wID = Dummy
    Else
        MII.wID = Item
    End If
    
    If GetMenuItemInfo(hMenu, MII.wID, False, MII) = 0 Then
        Err.Raise vbObjectError, "modCloseBtn::" & FuncName, _
            "modCloseBtn::" & FuncName & "() - Menu Item Not Found"
        Exit Sub
    End If
    
    ' Switch the ID of the menu item so that VB can not undo the action itself
    
    Dim lngMenuID As Long
    lngMenuID = MII.wID
    
    If Enable Then
        MII.wID = Item
    Else
        MII.wID = Dummy
    End If
    
    MII.fMask = MIIM_ID
    If SetMenuItemInfo(hMenu, lngMenuID, False, MII) = 0 Then
        Err.Raise vbObjectError, "modCloseBtn::" & FuncName, _
            "modCloseBtn::" & FuncName & "() - Error encountered " & _
            "changing ID"
        Exit Sub
    End If
    
    ' Set the enabled / disabled state of the menu item
    
    If Enable Then
        MII.fState = MII.fState And Not MFS_GRAYED
    Else
        MII.fState = MII.fState Or MFS_GRAYED
    End If
    
    MII.fMask = MIIM_STATE
    If SetMenuItemInfo(hMenu, MII.wID, False, MII) = 0 Then
         Err.Raise vbObjectError, "modCloseBtn::" & FuncName, _
            "modCloseBtn::" & FuncName & "() - Error encountered " & _
            "changing state"
        Exit Sub
    End If
    
    ' Activate the non-client area of the window to update the titlebar, and
    ' draw the Max button in its new state.
    
    SendMessage hWnd, WM_NCACTIVATE, True, 0
    
End Sub

'*******************************************************************************
'
'-------------------------------------------------------------------------------



