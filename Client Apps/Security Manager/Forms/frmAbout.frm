VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About OpenHR - Security Manager"
   ClientHeight    =   6825
   ClientLeft      =   345
   ClientTop       =   4815
   ClientWidth     =   6165
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   8001
   Icon            =   "frmAbout.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6825
   ScaleWidth      =   6165
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdTech 
      Caption         =   "&Support..."
      Height          =   400
      Left            =   1440
      TabIndex        =   1
      Top             =   6240
      Width           =   1425
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   400
      Left            =   4560
      TabIndex        =   0
      Top             =   6240
      Width           =   1455
   End
   Begin VB.CommandButton cmdSysInfo 
      Caption         =   "System &Info..."
      Height          =   400
      Left            =   3000
      TabIndex        =   2
      Top             =   6240
      Width           =   1425
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00808080&
      X1              =   120
      X2              =   6015
      Y1              =   6000
      Y2              =   6000
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      X1              =   120
      X2              =   6015
      Y1              =   3360
      Y2              =   3360
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "About OpenHR - Security Manager"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   120
      TabIndex        =   11
      Top             =   3000
      Width           =   3705
   End
   Begin VB.Image Image1 
      Height          =   2820
      Left            =   0
      Picture         =   "frmAbout.frx":000C
      Top             =   0
      Width           =   6150
   End
   Begin VB.Label lblAdvancedConnectURL 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Visit Advanced Connect for the latest OpenHR news and events"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   150
      MouseIcon       =   "frmAbout.frx":25DE
      MousePointer    =   99  'Custom
      TabIndex        =   10
      Top             =   5520
      Width           =   5430
   End
   Begin VB.Label lblURL 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "https://www.oneadvanced.com"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   150
      MouseIcon       =   "frmAbout.frx":2730
      MousePointer    =   99  'Custom
      TabIndex        =   9
      Top             =   5115
      Width           =   2670
   End
   Begin VB.Label lblDatabase 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Database : "
      Height          =   195
      Left            =   120
      TabIndex        =   8
      Top             =   3735
      Width           =   1110
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "OpenHR Security Manager - version"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   120
      TabIndex        =   7
      Top             =   3525
      Width           =   2970
   End
   Begin VB.Label lblSql 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "SQL Server Version : "
      Height          =   195
      Left            =   120
      TabIndex        =   6
      Top             =   4365
      Width           =   3240
   End
   Begin VB.Label lblSecurity 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "User Group : "
      Height          =   195
      Left            =   120
      TabIndex        =   5
      Top             =   4155
      Width           =   3285
   End
   Begin VB.Label lblUser 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "User : "
      Height          =   180
      Left            =   120
      TabIndex        =   4
      Top             =   3945
      Width           =   705
   End
   Begin VB.Label lblCopyRight 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Copyright © Advanced"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   150
      TabIndex        =   3
      Top             =   4785
      Width           =   1965
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Reg Key Security Options...
Const READ_CONTROL = &H20000
Const KEY_QUERY_VALUE = &H1
Const KEY_SET_VALUE = &H2
Const KEY_CREATE_SUB_KEY = &H4
Const KEY_ENUMERATE_SUB_KEYS = &H8
Const KEY_NOTIFY = &H10
Const KEY_CREATE_LINK = &H20
Const KEY_ALL_ACCESS = KEY_QUERY_VALUE + KEY_SET_VALUE + _
                       KEY_CREATE_SUB_KEY + KEY_ENUMERATE_SUB_KEYS + _
                       KEY_NOTIFY + KEY_CREATE_LINK + READ_CONTROL
                     
' Reg Key ROOT Types...
Const HKEY_LOCAL_MACHINE = &H80000002
Const ERROR_SUCCESS = 0
Const REG_SZ = 1                         ' Unicode nul terminated string
Const REG_DWORD = 4                      ' 32-bit number

Const gREGKEYSYSINFOLOC = "SOFTWARE\Microsoft\Shared Tools Location"
Const gREGVALSYSINFOLOC = "MSINFO"
Const gREGKEYSYSINFO = "SOFTWARE\Microsoft\Shared Tools\MSINFO"
Const gREGVALSYSINFO = "PATH"

Private Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, ByRef phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByVal lpData As String, ByRef lpcbData As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long




Private Sub cmdSysInfo_Click()
  Call StartSysInfo
  
End Sub

Private Sub cmdOK_Click()
  Unload Me
  
End Sub

Private Sub cmdTech_Click()
  frmTechSupport.Show vbModal

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
  Case vbKeyF1
    If ShowAirHelp(Me.HelpContextID) Then
      KeyCode = 0
    End If
End Select
End Sub

Private Sub Form_Load()
  On Error GoTo ErrorTrap
    
  'Dim sAniPath As String
  Dim sSQL As String
  Dim rsUser As Recordset
  Dim sngMaxX As Single
    
  'Start animation
  'sAniPath = App.Path & "\Videos\about.avi"
  'aniLogo.Open sAniPath
    
  ' lblTitle.Caption = "OpenHR Security Manager - v" & App.Major & "." & App.Minor & "." & App.Revision
  lblTitle.Caption = "Version : " & App.Major & "." & App.Minor & "." & App.Revision
  lblDatabase.Caption = "Database : " & gsDatabaseName
  lblUser.Caption = "Current User : " & Trim(gsUserName)
  lblSecurity.Caption = "User Group : " & gsUserGroup
  lblSql.Caption = GetSqlVersion
  lblCopyRight.Caption = "Copyright © Advanced"
  
  sngMaxX = lblTitle.Left + lblTitle.Width
  sngMaxX = IIf(lblDatabase.Left + lblDatabase.Width > sngMaxX, lblDatabase.Left + lblDatabase.Width, sngMaxX)
  sngMaxX = IIf(lblUser.Left + lblUser.Width > sngMaxX, lblUser.Left + lblUser.Width, sngMaxX)
  sngMaxX = IIf(lblSecurity.Left + lblSecurity.Width > sngMaxX, lblSecurity.Left + lblSecurity.Width, sngMaxX)
  sngMaxX = IIf(lblSql.Left + lblSql.Width > sngMaxX, lblSql.Left + lblSql.Width, sngMaxX)
  sngMaxX = IIf(lblCopyRight.Left + lblCopyRight.Width > sngMaxX, lblCopyRight.Left + lblCopyRight.Width, sngMaxX)
  
  'cmdOK.Left = sngMaxX + 250
  'cmdSysInfo.Left = cmdOK.Left
  'cmdTech.Left = cmdOK.Left
  
  'Me.Width = cmdOK.Left + cmdOK.Width + 200
  
  Exit Sub
    
ErrorTrap:
  'If Err.Number = 53 Then
  '  aniLogo.Visible = False
  '  imgASR.Visible = True
  '  Resume Next
  'End If

End Sub

Private Function GetSqlVersion() As String
  Dim sResult As String
  Dim sSQL As String
  Dim rsResult As New ADODB.Recordset

  sSQL = "exec sp_server_info 2"
  rsResult.Open sSQL, gADOCon, adOpenForwardOnly, adLockReadOnly
  sResult = rsResult(2)
  rsResult.Close
  Set rsResult = Nothing

  ' Trim the result.
  If InStr(sResult, vbLf) > 0 Then
    sResult = Left(sResult, InStr(sResult, vbLf) - 1)
  End If
  
  GetSqlVersion = sResult

End Function


'''Private Function GetUserDetails() As String
'''  ' Return the current user's user group.
'''  On Error GoTo ErrorTrap
'''
'''  Dim sSQL As String
'''  Dim sUserGroup As String
'''
'''  'JPD 20050812 Fault 10166
'''  sSQL = "exec sp_helpuser '" & Replace(gsUserName, "'", "''") & "'"
'''
'''  Set rsUser = rdoCon.OpenResultset(sSQL, _
'''    rdOpenStatic, rdConcurReadOnly, rdExecDirect)
'''
'''  'MH20031107 Fault 5627
'''  'If rsUser!GroupName = "db_owner" Then
'''  '  rsUser.MoveNext
'''  'End If
'''  Do While rsUser!GroupName = "db_owner" _
'''        Or LCase(Left(rsUser!GroupName, 6)) = "asrsys"
'''    rsUser.MoveNext
'''  Loop
'''
'''  sUserGroup = rsUser!GroupName
'''  rsUser.Close
'''
'''TidyUpAndExit:
'''  Set rsUser = Nothing
'''  GetUserDetails = sUserGroup
'''  Exit Function
'''
'''ErrorTrap:
'''  sUserGroup = "<none>"
'''  Resume TidyUpAndExit
'''
'''End Function

Public Sub StartSysInfo()
  On Error GoTo SysInfoErr
  
  Dim rc As Long
  Dim SysInfoPath As String
    
  ' Try To Get System Info Program Path\Name From Registry...
  If GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFO, gREGVALSYSINFO, SysInfoPath) Then
    ' Try To Get System Info Program Path Only From Registry...
  ElseIf GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFOLOC, gREGVALSYSINFOLOC, SysInfoPath) Then
    ' Validate Existance Of Known 32 Bit File Version
    If (Dir(SysInfoPath & "\MSINFO32.EXE") <> "") Then
      SysInfoPath = SysInfoPath & "\MSINFO32.EXE"
            
      ' Error - File Can Not Be Found...
    Else
      GoTo SysInfoErr
    End If
    ' Error - Registry Entry Can Not Be Found...
  Else
    GoTo SysInfoErr
  End If
    
  Call Shell(SysInfoPath, vbNormalFocus)
    
  Exit Sub

SysInfoErr:
  MsgBox "System information is unavailable at this time", vbOKOnly
  
End Sub

Public Function GetKeyValue(KeyRoot As Long, KeyName As String, SubKeyRef As String, ByRef KeyVal As String) As Boolean
    Dim i As Long                                           ' Loop Counter
    Dim rc As Long                                          ' Return Code
    Dim hKey As Long                                        ' Handle To An Open Registry Key
    Dim hDepth As Long                                      '
    Dim KeyValType As Long                                  ' Data Type Of A Registry Key
    Dim tmpVal As String                                    ' Tempory Storage For A Registry Key Value
    Dim KeyValSize As Long                                  ' Size Of Registry Key Variable
    '------------------------------------------------------------
    ' Open RegKey Under KeyRoot {HKEY_LOCAL_MACHINE...}
    '------------------------------------------------------------
    rc = RegOpenKeyEx(KeyRoot, KeyName, 0, KEY_ALL_ACCESS, hKey) ' Open Registry Key
    
    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' Handle Error...
    
    tmpVal = String$(1024, 0)                             ' Allocate Variable Space
    KeyValSize = 1024                                       ' Mark Variable Size
    
    '------------------------------------------------------------
    ' Retrieve Registry Key Value...
    '------------------------------------------------------------
    rc = RegQueryValueEx(hKey, SubKeyRef, 0, _
                         KeyValType, tmpVal, KeyValSize)    ' Get/Create Key Value
                        
    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' Handle Errors
    
    If (Asc(Mid(tmpVal, KeyValSize, 1)) = 0) Then           ' Win95 Adds Null Terminated String...
        tmpVal = Left(tmpVal, KeyValSize - 1)               ' Null Found, Extract From String
    Else                                                    ' WinNT Does NOT Null Terminate String...
        tmpVal = Left(tmpVal, KeyValSize)                   ' Null Not Found, Extract String Only
    End If
    '------------------------------------------------------------
    ' Determine Key Value Type For Conversion...
    '------------------------------------------------------------
    Select Case KeyValType                                  ' Search Data Types...
    Case REG_SZ                                             ' String Registry Key Data Type
        KeyVal = tmpVal                                     ' Copy String Value
    Case REG_DWORD                                          ' Double Word Registry Key Data Type
        For i = Len(tmpVal) To 1 Step -1                    ' Convert Each Bit
            KeyVal = KeyVal + Hex(Asc(Mid(tmpVal, i, 1)))   ' Build Value Char. By Char.
        Next
        KeyVal = Format$("&h" + KeyVal)                     ' Convert Double Word To String
    End Select
    
    GetKeyValue = True                                      ' Return Success
    rc = RegCloseKey(hKey)                                  ' Close Registry Key
    Exit Function                                           ' Exit
    
GetKeyError:      ' Cleanup After An Error Has Occured...
    KeyVal = ""                                             ' Set Return Val To Empty String
    GetKeyValue = False                                     ' Return Failure
    rc = RegCloseKey(hKey)                                  ' Close Registry Key
End Function

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  ' Redo link colour
  lblURL.ForeColor = &HFF0000
  lblAdvancedConnectURL.ForeColor = &HFF0000
  DoEvents

End Sub

Private Sub Form_Resize()
  'JPD 20030908 Fault 5756
  DisplayApplication

End Sub


Private Sub lblURL_Click()
  On Error GoTo ErrTrap

  Dim plngID As Integer
  
  'Show that the 'hyperlink' has been clicked on
  'lblURL.ForeColor = &H800080
  DoEvents
  
  ' Replaced the following line in the hope of making ShellExecute work on all PCs.
  ' Dont think it worked !
  plngID = ShellExecute(0&, vbNullString, Trim(lblURL.Caption), vbNullString, vbNullString, vbMaximizedFocus)
  
  If plngID = 0 Then
    ' Uh oh...the browser wasnt initiated...tell the user
    MsgBox "OpenHR cannot automatically open your default web browser." & vbCrLf & vbCrLf & "Please open your web browser manually and navigate to the " & vbCrLf & "web address which has been placed in your clipboard." & IIf(Err.Description = "", "", vbCrLf & vbCrLf & "(" & Err.Number & " - " & Err.Description & ")"), vbInformation + vbOKOnly, "Technical Support"
  End If
  
  Exit Sub
  
ErrTrap:
    MsgBox "OpenHR cannot automatically open your default web browser." & vbCrLf & vbCrLf & "Please open your web browser manually and navigate to the " & vbCrLf & "web address which has been placed in your clipboard." & IIf(Err.Description = "", "", vbCrLf & vbCrLf & "(" & Err.Number & " - " & Err.Description & ")"), vbInformation + vbOKOnly, "Technical Support"

End Sub


Private Sub lblURL_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  ' Highlight the link
  lblURL.ForeColor = vbRed
  DoEvents

End Sub

Private Sub lblAdvancedConnectURL_Click()
  On Error GoTo ErrTrap

  Dim plngID As Integer
  Dim URLTarget As String
  
  URLTarget = "http://www.advancedconnect.co.uk/"
  
  'Show that the 'hyperlink' has been clicked on
  'lblURL.ForeColor = &H800080
  DoEvents
  
  ' Replaced the following line in the hope of making ShellExecute work on all PCs.
  ' Dont think it worked !
  plngID = ShellExecute(0&, vbNullString, URLTarget, vbNullString, vbNullString, vbMaximizedFocus)
  
  If plngID = 0 Then
    ' Uh oh...the browser wasnt initiated...tell the user
    MsgBox "OpenHR cannot automatically open your default web browser." & vbCrLf & vbCrLf & "Please open your web browser manually and navigate to the " & vbCrLf & "web address which has been placed in your clipboard." & IIf(Err.Description = "", "", vbCrLf & vbCrLf & "(" & Err.Number & " - " & Err.Description & ")"), vbInformation + vbOKOnly, "Technical Support"
  End If
  
  Exit Sub
  
ErrTrap:
    MsgBox "OpenHR cannot automatically open your default web browser." & vbCrLf & vbCrLf & "Please open your web browser manually and navigate to the " & vbCrLf & "web address which has been placed in your clipboard." & IIf(Err.Description = "", "", vbCrLf & vbCrLf & "(" & Err.Number & " - " & Err.Description & ")"), vbInformation + vbOKOnly, "Technical Support"

End Sub


Private Sub lblAdvancedConnectURL_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  ' Highlight the link
  lblAdvancedConnectURL.ForeColor = vbRed
  DoEvents

End Sub

