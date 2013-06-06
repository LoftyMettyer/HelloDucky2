Attribute VB_Name = "modSysTray"
Option Explicit

'Declare a user-defined variable to pass to the Shell_NotifyIcon
'Changes
      'function.
      Private Type NOTIFYICONDATA
         cbSize As Long
         hWnd As Long
         uId As Long
         uFlags As Long
         uCallBackMessage As Long
         hIcon As Long
         szTip As String * 64
      End Type

      'Declare the constants for the API function. These constants can be
      'found in the header file Shellapi.h.

      'The following constants are the messages sent to the
      'Shell_NotifyIcon function to add, modify, or delete an icon from the
      'taskbar status area.
      Private Const NIM_ADD = &H0
      Private Const NIM_MODIFY = &H1
      Private Const NIM_DELETE = &H2

      'The following constant is the message sent when a mouse event occurs
      'within the rectangular boundaries of the icon in the taskbar status
      'area.
      Private Const WM_MOUSEMOVE = &H200

      'The following constants are the flags that indicate the valid
      'members of the NOTIFYICONDATA data type.
      Private Const NIF_MESSAGE = &H1
      Private Const NIF_ICON = &H2
      Private Const NIF_TIP = &H4

      'Declare the API function call.
      Private Declare Function Shell_NotifyIcon Lib "shell32" _
         Alias "Shell_NotifyIconA" _
         (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean

      'Dimension a variable as the user-defined data type.
      Dim nid As NOTIFYICONDATA

      Public Sub AddSysTrayIcon(strToolTip As String)
         'Click this button to add an icon to the taskbar status area.

        'Set frmMain.Icon = LoadResPicture("BELL", 1)
        'Set frmMain.Icon = LoadResPicture("ASR", 1)
         
         'Set the individual values of the NOTIFYICONDATA data type.
         nid.cbSize = Len(nid)
         nid.hWnd = frmDiaryAlert.hWnd
         nid.uId = vbNull
         nid.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
         nid.uCallBackMessage = WM_MOUSEMOVE
         nid.hIcon = frmMain.Icon ' LoadResPicture("BELL", 1)  'frmMain.Icon
         nid.szTip = strToolTip & vbNullChar

         'Call the Shell_NotifyIcon function to add the icon to the taskbar
         'status area.
         Shell_NotifyIcon NIM_ADD, nid
         
         'frmMain.Icon = LoadResPicture("!ASRSMALL", 1)

      End Sub
      Public Sub RemoveSysTrayIcon()
         'Click this button to delete the added icon from the taskbar
         'status area by calling the Shell_NotifyIcon function.
         Shell_NotifyIcon NIM_DELETE, nid
      End Sub
