VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About HR Pro Data Manager"
   ClientHeight    =   1845
   ClientLeft      =   345
   ClientTop       =   4815
   ClientWidth     =   6570
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
   HelpContextID   =   1001
   Icon            =   "frmAbout.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1845
   ScaleWidth      =   6570
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      ForeColor       =   &H8000000F&
      Height          =   285
      Left            =   4050
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   5000
      Width           =   1395
   End
   Begin VB.CommandButton cmdTech 
      Caption         =   "&Support..."
      Height          =   400
      Left            =   5000
      TabIndex        =   2
      Top             =   1300
      Width           =   1425
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1050
      Left            =   200
      Picture         =   "frmAbout.frx":000C
      ScaleHeight     =   1020
      ScaleWidth      =   1125
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   150
      Width           =   1155
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   400
      Left            =   5000
      TabIndex        =   0
      Top             =   150
      Width           =   1425
   End
   Begin VB.CommandButton cmdSysInfo 
      Caption         =   "System &Info..."
      Height          =   400
      Left            =   5000
      TabIndex        =   1
      Top             =   725
      Width           =   1425
   End
   Begin VB.Label lblURL 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "http://www.advancedcomputersoftware.com"
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
      MouseIcon       =   "frmAbout.frx":3C14
      MousePointer    =   99  'Custom
      TabIndex        =   11
      Top             =   1515
      Width           =   3810
   End
   Begin VB.Label lblDatabase 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Database : "
      Height          =   195
      Left            =   1500
      TabIndex        =   10
      Top             =   370
      Width           =   2280
   End
   Begin VB.Label lblUser 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "User : "
      Height          =   195
      Left            =   1500
      TabIndex        =   8
      Top             =   580
      Width           =   2100
   End
   Begin VB.Label lblSecurity 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "User Group : "
      Height          =   195
      Left            =   1500
      TabIndex        =   7
      Top             =   795
      Width           =   1170
   End
   Begin VB.Label lblSql 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "SQL Server Version : "
      Height          =   195
      Left            =   1500
      TabIndex        =   6
      Top             =   1000
      Width           =   2295
   End
   Begin VB.Label lblCopyRight 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Copyright © Advanced Computer Solutions"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   150
      TabIndex        =   4
      Top             =   1305
      Width           =   3720
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "HR Pro Data Manager - version"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   1500
      TabIndex        =   3
      Top             =   160
      Width           =   2670
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

'Private Sub aniAbout_DblClick()
'Text1.Text = ""
'Text1.SetFocus
'End Sub

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
    
  On Error GoTo Err_Trap
  
  'Dim sAniPath As String
  Dim sSQL As String
  Dim rsUser As Recordset
  Dim datGeneral As New HRProDataMgr.clsGeneral
  Dim sngMaxX As Single
     
  'sAniPath = App.Path & "\Videos\about.avi"
  'aniAbout.Open sAniPath
    
  lblTitle.Caption = "HR Pro Data Manager - v" & App.Major & "." & App.Minor & "." & App.Revision
  lblDatabase.Caption = "Database : " & gsDatabaseName
  lblUser.Caption = "Current User : " & gsUserName
  
  If LCase(gsSQLUserName) <> LCase(gsUserName) Then
    lblUser.Caption = lblUser.Caption & " (" & gsSQLUserName & ")"
  End If
  If ASRDEVELOPMENT Then
    lblUser.Caption = lblUser.Caption & " (SPID = " & datGeneral.GetSqlProcessID & ")"
  End If
  
  lblSecurity = "User Group : " & gsUserGroup
  lblSql = datGeneral.GetSqlVersion
  lblCopyRight.Caption = "Copyright © COA Solutions Limited 1997-" & Format(Date, "yyyy")
  
  Set datGeneral = Nothing
  Screen.MousePointer = vbDefault
  
  sngMaxX = lblTitle.Left + lblTitle.Width
  sngMaxX = IIf(lblDatabase.Left + lblDatabase.Width > sngMaxX, lblDatabase.Left + lblDatabase.Width, sngMaxX)
  sngMaxX = IIf(lblUser.Left + lblUser.Width > sngMaxX, lblUser.Left + lblUser.Width, sngMaxX)
  sngMaxX = IIf(lblSecurity.Left + lblSecurity.Width > sngMaxX, lblSecurity.Left + lblSecurity.Width, sngMaxX)
  sngMaxX = IIf(lblSql.Left + lblSql.Width > sngMaxX, lblSql.Left + lblSql.Width, sngMaxX)
  sngMaxX = IIf(lblCopyRight.Left + lblCopyRight.Width > sngMaxX, lblCopyRight.Left + lblCopyRight.Width, sngMaxX)
  
  cmdOK.Left = sngMaxX + 250
  cmdSysInfo.Left = cmdOK.Left
  cmdTech.Left = cmdOK.Left
  
  Me.Width = cmdOK.Left + cmdOK.Width + 200
  
  Exit Sub
    
Err_Trap:
  'If Err.Number = 53 Then
  '    aniAbout.Visible = False
  '    imgASR.Visible = True
  '    Resume Next
  'End If
  Screen.MousePointer = vbDefault
    
End Sub

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
    COAMsgBox "System information is unavailable.", vbOKOnly
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

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
  ' Redo link colour
  lblURL.ForeColor = &HFF0000
  DoEvents

End Sub

Private Sub Form_Resize()
  'JPD 20030908 Fault 5756
  DisplayApplication
  
End Sub

Private Sub imgASR_Click()
  Text1.Text = ""
  Text1.SetFocus
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
    COAMsgBox "HR Pro cannot automatically open your default web browser." & vbCrLf & vbCrLf & "Please open your web browser manually and navigate to the " & vbCrLf & "web address which has been placed in your clipboard." & IIf(Err.Description = "", "", vbCrLf & vbCrLf & "(" & Err.Number & " - " & Err.Description & ")"), vbInformation + vbOKOnly, "Technical Support"
  End If
  
  Exit Sub
  
ErrTrap:
    COAMsgBox "HR Pro cannot automatically open your default web browser." & vbCrLf & vbCrLf & "Please open your web browser manually and navigate to the " & vbCrLf & "web address which has been placed in your clipboard." & IIf(Err.Description = "", "", vbCrLf & vbCrLf & "(" & Err.Number & " - " & Err.Description & ")"), vbInformation + vbOKOnly, "Technical Support"

End Sub

Private Sub lblURL_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
  ' Highlight the link
  lblURL.ForeColor = vbRed
  DoEvents

End Sub


Private Sub Picture1_Click()
  Text1.Text = ""
  Text1.SetFocus
End Sub

Private Sub Text1_Change()
  If UCase(Text1.Text) = "LIVERPOOL" Then
    If Not frmAboutEgg Is Nothing Then
      frmAboutEgg.Show vbModal
      Text1.Text = ""
    End If
  End If
End Sub


