VERSION 5.00
Begin VB.Form frmDiaryAlert 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "OpenHR Alarmed Diary Events"
   ClientHeight    =   1425
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6390
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmDiaryAlert.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1425
   ScaleWidth      =   6390
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdNo 
      Cancel          =   -1  'True
      Caption         =   "&No"
      Height          =   350
      Left            =   3330
      TabIndex        =   2
      Top             =   880
      Width           =   1120
   End
   Begin VB.CommandButton cmdYes 
      Caption         =   "&Yes"
      Default         =   -1  'True
      Height          =   350
      Left            =   2085
      TabIndex        =   1
      Top             =   880
      Width           =   1120
   End
   Begin VB.Label lblMessage 
      Caption         =   "There are alarmed events prior to the current date and time. Would you like to view these now?"
      Height          =   540
      Left            =   960
      TabIndex        =   0
      Top             =   240
      Width           =   5310
   End
   Begin VB.Image imgIcon 
      Height          =   480
      Index           =   1
      Left            =   200
      Picture         =   "frmDiaryAlert.frx":000C
      Top             =   200
      Width           =   480
   End
End
Attribute VB_Name = "frmDiaryAlert"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mlngResponse As Long
Private mlngErrorNumber As Long

Public Function ShowAlert(blnCurrent As Boolean) As Long

  On Local Error GoTo LocalErr

  ShowAlert = 0
  mlngErrorNumber = 0

  If blnCurrent Then
    Me.lblMessage.Caption = _
        "There are alarmed events for the current date and time." & vbCr & _
        "Would you like to view these now?"
  End If

  EnableCloseButton Me.hWnd, False

  'Always on top!
  Me.Show
  frmMain.Enabled = False
  Call SetWindowPos(Me.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE)

  mlngResponse = 0
  Do While mlngResponse = 0
    DoEvents
  Loop
  
  If mlngResponse = vbYes Then
    SetForegroundWindow frmMain.hWnd
  End If
  ShowAlert = mlngResponse
  
  frmMain.Enabled = True
  Me.Hide

Exit Function

LocalErr:
  mlngErrorNumber = Err.Number

End Function

Public Property Get ErrorNumber() As Long
  ErrorNumber = mlngErrorNumber
End Property


Private Sub cmdNo_Click()
  mlngResponse = vbNo
End Sub

Private Sub cmdYes_Click()
  mlngResponse = vbYes
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
  Case vbKeyF1
    If ShowAirHelp(Me.HelpContextID) Then
      KeyCode = 0
    End If
End Select
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)

  Select Case KeyCode
  Case vbKeyY
    mlngResponse = vbYes
  Case vbKeyN
    mlngResponse = vbNo
  End Select

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

  'THIS HANDLES THE DOUBLE CLICK OF THE ICON IN THE SYSTEM TRAY

  'Static rec As Boolean, msg As Long, oldmsg As Long
  Dim msg
  'oldmsg = msg
  msg = x / Screen.TwipsPerPixelX

  'If rec = False Then
  '  rec = True
    If msg = WM_LBUTTONDBLCLK Then
      mlngResponse = vbYes
    End If
  'End If

End Sub



Private Function EnableCloseButton(ByVal hWnd As Long, Enable As Boolean) As Integer
    EnableSystemMenuItem hWnd, SC_CLOSE, xSC_CLOSE, Enable, "EnableCloseButton"
End Function


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

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If mlngResponse = 0 Then
    mlngResponse = vbNo
  End If
End Sub
