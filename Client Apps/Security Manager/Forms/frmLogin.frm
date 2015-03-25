VERSION 5.00
Object = "{BD0C1912-66C3-49CC-8B12-7B347BF6C846}#13.1#0"; "Codejock.SkinFramework.v13.1.0.ocx"
Begin VB.Form frmLogin 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "OpenHR Security Manager - Login"
   ClientHeight    =   3675
   ClientLeft      =   1530
   ClientTop       =   3285
   ClientWidth     =   6405
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   8003
   Icon            =   "frmLogin.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3675
   ScaleWidth      =   6405
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkUseWindowsAuthentication 
      Caption         =   "&Use Windows Authentication"
      Height          =   210
      Left            =   240
      TabIndex        =   3
      Top             =   2520
      Value           =   1  'Checked
      Width           =   3240
   End
   Begin VB.CommandButton cmdDetails 
      Caption         =   "&Details  >>"
      Height          =   400
      Left            =   4950
      TabIndex        =   8
      Top             =   2835
      Width           =   1200
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   400
      Left            =   4950
      TabIndex        =   6
      Top             =   1680
      Width           =   1200
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   400
      Left            =   4950
      TabIndex        =   7
      Top             =   2235
      Width           =   1200
   End
   Begin VB.TextBox txtUID 
      Height          =   315
      Left            =   1515
      TabIndex        =   1
      Top             =   1680
      Width           =   3280
   End
   Begin VB.TextBox txtPWD 
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   1515
      MaxLength       =   128
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   2085
      Width           =   3280
   End
   Begin VB.TextBox txtDatabase 
      Height          =   315
      Left            =   1515
      TabIndex        =   4
      Top             =   2835
      Visible         =   0   'False
      Width           =   3280
   End
   Begin VB.TextBox txtServer 
      Height          =   315
      Left            =   1515
      TabIndex        =   5
      Top             =   3240
      Visible         =   0   'False
      Width           =   3280
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   65535
      Left            =   5055
      Top             =   2985
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000005&
      X1              =   240
      X2              =   6120
      Y1              =   1515
      Y2              =   1515
   End
   Begin XtremeSkinFramework.SkinFramework SkinFramework1 
      Left            =   5625
      Top             =   3105
      _Version        =   851969
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
   Begin VB.Label lblDevelopmentMode 
      AutoSize        =   -1  'True
      Caption         =   "(Dev mode)"
      Height          =   195
      Left            =   5145
      TabIndex        =   13
      Top             =   1215
      Width           =   1035
   End
   Begin VB.Label lblUser 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Username :"
      Height          =   195
      Left            =   240
      TabIndex        =   12
      Top             =   1740
      Width           =   1005
   End
   Begin VB.Label lblPassword 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Password :"
      Height          =   195
      Left            =   240
      TabIndex        =   11
      Top             =   2145
      Width           =   930
   End
   Begin VB.Label lblDatabase 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Database :"
      Height          =   195
      Left            =   240
      TabIndex        =   10
      Top             =   2895
      Visible         =   0   'False
      Width           =   930
   End
   Begin VB.Label lblServer 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Server :"
      Height          =   195
      Left            =   240
      TabIndex        =   9
      Top             =   3300
      Visible         =   0   'False
      Width           =   810
   End
   Begin VB.Label lblVersion 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "lblVersion"
      Height          =   195
      Left            =   240
      TabIndex        =   0
      Top             =   1215
      Width           =   840
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000010&
      BorderWidth     =   2
      X1              =   240
      X2              =   6100
      Y1              =   1515
      Y2              =   1515
   End
   Begin VB.Image imgLogo 
      Height          =   1050
      Left            =   120
      Picture         =   "frmLogin.frx":000C
      Top             =   120
      Width           =   3135
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' API functions.
Private Declare Function SQLAllocEnv% Lib "odbc32.dll" (env&)
Private Declare Function SQLDrivers Lib "odbc32.dll" (ByVal henv&, ByVal fDirection%, ByVal szDriverDesc$, ByVal cbDriverDescMax%, pcbDriverDescBuff%, ByVal szDriverAttrib$, ByVal cbDriverAttribMax%, pcbDriverAttribBuff%) As Integer

' API constants.
Const SQL_SUCCESS As Long = 0
Const SQL_FETCH_NEXT As Long = 1

' Constants.
Const ODBCDRIVER As String = "SQL Server"

' Private classes.
Private UI As New UI
Private Net As New Net

' Globals.
Private gDatStartDate As Date
Private mblnForceReadOnly As Boolean
Private mlngTimeOut As Long
Public OK As Boolean
Private mbForceTotalRelogin As Boolean

Private Sub chkUseWindowsAuthentication_Click()
  
  Dim strTrustedUser As String

  strTrustedUser = gstrWindowsCurrentDomain & "\" & gstrWindowsCurrentUser

  txtUID.Text = IIf(chkUseWindowsAuthentication.Value = vbChecked, strTrustedUser, txtUID.Text)
  txtUID.Enabled = IIf(chkUseWindowsAuthentication.Value = vbChecked, False, True)
  txtPWD.Enabled = IIf(chkUseWindowsAuthentication.Value = vbChecked, False, True)

  ' Grey out controls
  txtUID.BackColor = IIf(txtUID.Enabled, vbWindowBackground, vbButtonFace)
  txtPWD.BackColor = IIf(txtPWD.Enabled, vbWindowBackground, vbButtonFace)

End Sub

Private Sub txtUID_KeyPress(KeyAscii As Integer)

  If Len(Me.txtUID.Text) >= 50 Then
    KeyAscii = 0
  End If
  
End Sub

Private Sub txtUID_Change()

  If Len(Me.txtUID.Text) > 50 Then
    Me.txtUID.Text = Left(Me.txtUID.Text, 50)
  End If
  
End Sub

Private Sub cmdDetails_Click()
  ' Display/hide the Database and Server controls.
  Dim fDisplayControls As Boolean
  
  fDisplayControls = Not lblDatabase.Visible
  
  FormatControls fDisplayControls

  If txtDatabase.Visible Then
    txtDatabase.SetFocus
  End If
  
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
  gDatStartDate = Now
  
End Sub

Private Sub Form_Load()
  On Error GoTo ErrorTrap
  
  Dim fDatabaseInfoMissing  As Boolean
  Dim sAnimationPath As String
  Dim sUserName As String
  Dim sDatabaseName As String
  Dim sServerName As String
  Dim bUseWindowsAuthentication As Boolean
  
  ' Load the CodeJock Styles
  Call LoadSkin(Me, Me.SkinFramework1)

  'ShowInTaskbar
  SetWindowLong Me.hWnd, GWL_EXSTYLE, (GetWindowLong(hWnd, GWL_EXSTYLE) Or WS_EX_APPWINDOW)


  ' Position the login screen in the centre of the screen.
  UI.frmAtCenter Me
  lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision

  If ASRDEVELOPMENT Then
    lblDevelopmentMode.Caption = "(Dev mode)"
    lblDevelopmentMode.Visible = True
    lblDevelopmentMode.Enabled = True
  Else
    lblDevelopmentMode.Caption = ""
    lblDevelopmentMode.Visible = False
    lblDevelopmentMode.Enabled = False
  End If
  
  ' Display the animated logo.
  'sAnimationPath = App.Path & "\videos\asr.avi"
  'On Error GoTo ErrorNoAnimation
  'aniLogo.Open sAnimationPath
  'On Error GoTo ErrorTrap
  
  ' Check that the SQL Server driver is installed.
  If Not CheckSQLDriver Then
    MsgBox "The required ODBC driver '" & ODBCDRIVER & "' is not installed" & vbCrLf & _
      "Install the driver before running OpenHR", _
      vbExclamation + vbOKOnly, App.ProductName
    Unload frmLogin
    Exit Sub
  End If
  
  ' Get the details of the last connection from the registry.
  sUserName = GetPCSetting("Login", "SecMgr_UserName", Net.UserName)
  sDatabaseName = GetPCSetting("Login", "SecMgr_Database", vbNullString)
  sServerName = GetPCSetting("Login", "SecMgr_Server", vbNullString)
  bUseWindowsAuthentication = GetPCSetting("Login", "SecMgr_AuthenticationMode", 0)
      
  ' If the database/server information is missing from the registry then display the
  ' controls.
  fDatabaseInfoMissing = (sDatabaseName = vbNullString) Or _
    (sServerName = vbNullString)
  
  FormatControls fDatabaseInfoMissing
   
  txtUID.Text = sUserName
  txtDatabase.Text = sDatabaseName
  txtServer.Text = sServerName
  chkUseWindowsAuthentication.Value = IIf(bUseWindowsAuthentication = True, vbChecked, vbUnchecked)
  
  CheckCommandLine
  
  gDatStartDate = Now
  Timer1.Enabled = True

TidyUpAndExit:
  Exit Sub

ErrorTrap:
  Resume TidyUpAndExit
  
ErrorNoAnimation:
  'If Err.Number = 53 Then
  '  aniLogo.Visible = False
  '  imgLogo.Visible = True
  '  Resume Next
  'End If
  Resume TidyUpAndExit
  
End Sub

Private Sub Form_Activate()
  
  'If ASRDEVELOPMENT Then
  '  FormatControls True
  'End If

  If Len(txtUID.Text) > 0 Then
    txtPWD.SetFocus
  End If
  
  ' Normal/Windows authentication
  chkUseWindowsAuthentication_Click
  
End Sub

Private Sub DisplayOtherUsers(ByVal pasUsers As Variant)
  ' Format the from controls to display the array of users passed in.
  Dim fOneUser As Boolean
  Dim iLoop As Integer
  Dim sDisplay As String
  
  fOneUser = (UBound(pasUsers, 2) = 1)
  
  ' First warning row.
  sDisplay = "The following user" & IIf(fOneUser, " is", "s are") & " currently logged into the OpenHR databases :" & vbCrLf
  
  ' User rows.
  For iLoop = 1 To UBound(pasUsers, 2)
    sDisplay = sDisplay & vbCrLf & "        User '" & Trim(pasUsers(1, iLoop)) & "' on machine '" & Trim(pasUsers(2, iLoop)) & "'."
  Next iLoop

  ' Second warning row.
  sDisplay = sDisplay & vbCrLf & vbCrLf & "All users must be logged off before the OpenHR Security Manager can be used."
  
  ' Display the warning.
  MsgBox sDisplay, vbOKOnly, App.ProductName

End Sub

Private Sub cmdOK_Click()
  
  ' A Little Easter Egg!
  If txtUID.Text = "MattLucas" And txtPWD.Text = "LittleBritain" Then
    MsgBox "Computer Says No!!!", vbExclamation
    Exit Sub
  End If
  
  Login
  
  ' SQL 2005 may have forced us to change the password, so we need to re-attempt the logon process in full
  If mbForceTotalRelogin Then
    txtPWD.Text = gsPassword
    Login
  End If
  
End Sub

Private Sub cmdCancel_Click()
  OK = False
  Me.Hide
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloaSMode As Integer)
  If UnloaSMode <> vbFormCode Then
    OK = False
  End If
End Sub

Private Function CheckSQLDriver() As Boolean
  ' Check that the SQL ODBC Driver is installed.
  On Error GoTo ErrorTrap
    
  Dim fOK As Boolean
  Dim iReturnCode As Integer
  Dim iDriverDescLen As Integer
  Dim iDriverAttribLen As Integer
  Dim lngEnvironmentHandle As Long
  Dim sDriver As String
  Dim sDriverDesc As String * 1024
  Dim sDriverAttrib As String * 1024

  fOK = False
  
  ' Get the DSN and driver information.
  If SQLAllocEnv(lngEnvironmentHandle) <> -1 Then
  
    Do Until (iReturnCode <> SQL_SUCCESS) Or fOK
      
      sDriverDesc = Space(1024)
      
      iReturnCode = SQLDrivers(lngEnvironmentHandle, SQL_FETCH_NEXT, _
        sDriverDesc, 1024, iDriverDescLen, _
        sDriverAttrib, 1024, iDriverAttribLen)

      sDriver = UCase(Trim(Left(sDriverDesc, iDriverDescLen)))
                
      If Len(sDriver) > 0 Then
        If sDriver = UCase(ODBCDRIVER) Then
          fOK = True
        End If
      End If
    Loop
  End If
  
TidyUpAndExit:
  CheckSQLDriver = fOK
  Exit Function
  
ErrorTrap:
  fOK = False
  Resume TidyUpAndExit
  
End Function

Private Sub FormatControls(pfDisplayDatabaseControls As Boolean)
  ' Format the screen controls as required.
  Const GAP As Long = 500
  
  lblDatabase.Visible = pfDisplayDatabaseControls
  txtDatabase.Visible = pfDisplayDatabaseControls
  lblServer.Visible = pfDisplayDatabaseControls
  txtServer.Visible = pfDisplayDatabaseControls
  cmdDetails.Caption = IIf(pfDisplayDatabaseControls, "&Details  <<", "&Details  >>")

  'MH20020807
  'Don't change the height of the form - looks bad on Windows XP
  'frmLogin.Height = IIf(pfDisplayDatabaseControls, _
    txtServer.Top + txtServer.Height + GAP, _
    cmdDetails.Top + cmdDetails.Height + GAP)
  
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Set Net = Nothing
End Sub

Private Sub Timer1_Timer()
  If DateDiff("n", Now, gDatStartDate) >= 5 Then
    Timer1.Enabled = False
    cmdCancel_Click
  End If
  
End Sub

Private Sub txtDatabase_GotFocus()
  UI.txtSelText
End Sub

Private Sub txtPWD_GotFocus()
  UI.txtSelText
End Sub

Private Sub txtServer_GotFocus()
  UI.txtSelText
End Sub

Private Sub txtUID_GotFocus()
  UI.txtSelText
  
End Sub

Private Sub CheckPassword()

  On Error GoTo Check_ERROR
  
  Dim rsInfo As New ADODB.Recordset ' Recordset used to retrieve data
  Dim lMinimumLength As Long        ' The minimum length for passwords
  Dim lChangeFrequency As Long      ' How often passwords must be changed
  Dim sChangePeriod As String       ' How often passwords must be changed
  Dim dLastChanged As Date          ' Date user last changed password
  Dim fForceChange As Boolean       ' Must the user change password
  Dim iForceChange As PasswordChangeReason       ' 0 = Does not have to change
                                                 ' 1 = must change - length violation
                                                 ' 2 = must change - date violation
                                                 ' 3 = must change - forced to change
                                                 ' 4 = must change - not found in ASRSyspasswords
  Dim sSQL As String
  Dim iUsers As Integer
  
  '' First store the config info in local variables
  'Set rsInfo = rdoCon.OpenResultset("Select * From ASRSysConfig")
  'lMinimumLength = rsInfo!MinimumPasswordLength
  'lChangeFrequency = rsInfo!changepasswordfrequency
  'sChangePeriod = IIf(IsNull(rsInfo!changepasswordperiod), vbNu llString, rsInfo!changepasswordperiod)
  lMinimumLength = GetSystemSetting("Password", "Minimum Length", 0)
  lChangeFrequency = GetSystemSetting("Password", "Change Frequency", 0)
  sChangePeriod = GetSystemSetting("Password", "Change Period", "")

 
  If sChangePeriod = "W" Then sChangePeriod = "WW"
  If sChangePeriod = "Y" Then sChangePeriod = "YYYY"
  
  
  ' Get the users specific Info From ASRSysPasswords
  'JPD 20050812 Fault 10166
  rsInfo.Open "Select * From ASRSysPasswords WHERE Username = '" & LCase(Replace(gsUserName, "'", "''")) & "'", gADOCon, adOpenForwardOnly, adLockReadOnly
  
  If rsInfo.BOF And rsInfo.EOF Then
    ' User isnt in the table, so force a change
    'iForceChange = 4
  
    'JPD 20041117 Fault 9484
    ' RH 19/09/00 - BUG 961. If user is not in Asrsyspasswords, put them in with
    '                        the current date.
    'JPD 20050812 Fault 10166
    sSQL = "INSERT INTO AsrSysPasswords (Username, LastChanged, ForceChange) " & _
       "VALUES ('" & LCase(Replace(gsUserName, "'", "''")) & "','" & Replace(Format(Now, "mm/dd/yyyy"), UI.GetSystemDateSeparator, "/") & "',0)"
    gADOCon.Execute sSQL, , adExecuteNoRecords
    dLastChanged = Format(Now, "dd/mm/yyyy")
    fForceChange = False
  Else
    dLastChanged = rsInfo!LastChanged
    fForceChange = rsInfo!ForceChange
  End If
  
  rsInfo.Close
  Set rsInfo = Nothing
  
  ' If nothing has been configured, then might as well exit here
  If (lMinimumLength <> 0) Or (lChangeFrequency <> 0) Then
  
    ' Check for minimum length
    If iForceChange = 0 Then
      If lMinimumLength > Len(gsPassword) Then iForceChange = 1
    End If
    
    ' Check for Date last changed
    If iForceChange = 0 Then
      If lChangeFrequency > 0 Then
        If DateAdd(sChangePeriod, -lChangeFrequency, Now) >= dLastChanged Then iForceChange = 2
      End If
    End If
  
  End If
   
  ' Check for forced to change
  If iForceChange = 0 Then
    If fForceChange <> 0 Then iForceChange = 3
  End If
  
  ' If we are here and iforcechange = 0 then we dont have to change, so exit
  If iForceChange = 0 Then Exit Sub
  
  'MH20061017 Fault 11376
  ''' JPD 20020218 Fault 3527
  'iUsers = UserSessions(gsUserName)
  iUsers = GetCurrentUsersCountOnServer(gsUserName)
  If iUsers < 2 Then
    If frmChangePassword.Initialise(iForceChange, lMinimumLength) Then
      
      frmChangePassword.Show vbModal
      
      If frmChangePassword.Exiting Then
        Unload frmChangePassword
        gADOCon.Close
        End
      Else
        Unload frmChangePassword
        Set frmChangePassword = Nothing
        UpdateConfig
      End If
      
    End If
  End If
  
  Exit Sub
  
Check_ERROR:
  
  MsgBox "Error checking passwords." & vbCrLf & vbCrLf & "(" & Err.Number & " - " & Err.Description & ")", vbExclamation + vbOKOnly, App.Title
  
End Sub

Private Sub UpdateConfig()

  On Error GoTo Update_ERROR
  
  Dim rsInfo As New ADODB.Recordset

  ' Get the users specific Info From ASRSysPasswords
  'JPD 20050812 Fault 10166
  rsInfo.Open "Select * From ASRSysPasswords WHERE Username = '" & LCase(Replace(gsUserName, "'", "''")) & "'", gADOCon, adOpenDynamic, adLockOptimistic

  If rsInfo.BOF And rsInfo.EOF Then
    rsInfo.AddNew
  Else
    rsInfo.Update
  End If
  
  rsInfo!UserName = LCase(gsUserName)
  'JPD 20041117 Fault 9484
  rsInfo!LastChanged = Replace(Format(Now, "mm/dd/yyyy"), UI.GetSystemDateSeparator, "/")
  rsInfo!ForceChange = 0
  
  rsInfo.Update

  Set rsInfo = Nothing
  Exit Sub
  
Update_ERROR:
  
  MsgBox "Error updating AsrSysPasswords." & vbCrLf & vbCrLf & _
         "(" & Err.Number & " - " & Err.Description & ")", vbExclamation + vbOKOnly, App.Title
         
  Set rsInfo = Nothing
         
End Sub

Private Sub Login()

  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim sSQL As String
  Dim sConnect As String
  Dim rsUserInfo As New ADODB.Recordset
  Dim rsSQLInfo As New ADODB.Recordset
  Dim fSystemPermission As Boolean
  Dim sTempErrDescription As String
  Dim rsUser As New ADODB.Recordset
  Dim bIsSQLSystemAdmin As Boolean
  Dim bIsSecurityAdmin As Boolean
  Dim sUserName As String
  Dim sPassword As String
  
  fOK = True
  
  Screen.MousePointer = vbHourglass
  
  ' Clear any existing database connection.
  
  '14/08/2001 MH Fault 2447
  'This is only used in connection string, viewing users
  'and display so should be okay to leave case alone!
  'gsDatabaseName = LCase(txtDatabase.Text)
  sConnect = vbNullString
  gbUseWindowsAuthentication = chkUseWindowsAuthentication.Value
  gsDatabaseName = txtDatabase.Text
  gsServerName = txtServer.Text

  ' Build the ODBC connection string.
  sConnect = sConnect & "Driver={SQL Server};"
  If Len(Trim(txtServer.Text)) > 0 Then
    sConnect = sConnect & "Server=" & txtServer.Text & ";"
    gsSQLServerName = txtServer.Text
  Else
    Screen.MousePointer = vbNormal
    gobjProgress.CloseProgress
    MsgBox "Please enter the name of the server on which the OpenHR database is located", _
      vbExclamation + vbOKOnly, App.ProductName
    Exit Sub
  End If
      
  sUserName = Replace(txtUID.Text, ";", "")
  sPassword = Replace(txtPWD.Text, ";", "")
      
  If Len(sUserName) > 0 Then
    sConnect = sConnect & "UID=" & sUserName & ";PWD=" & sPassword & ";"
  Else
    Screen.MousePointer = vbNormal
    gobjProgress.CloseProgress
    MsgBox "Please enter a user name", _
      vbExclamation + vbOKOnly, App.ProductName
    Exit Sub
  End If
    
  If Len(txtDatabase.Text) > 0 Then
    sConnect = sConnect & "Database=" & txtDatabase.Text & ";"
  Else
    Screen.MousePointer = vbNormal
    gobjProgress.CloseProgress
    MsgBox "Please enter the name of the OpenHR database", _
      vbExclamation + vbOKOnly, App.ProductName
    Exit Sub
  End If
  
  'sConnect = sConnect & "APP=" & App.ProductName & ";"
  sConnect = sConnect & "Application Name=" & App.ProductName & ";"

  'Windows Authentication
  If gbUseWindowsAuthentication Then
    sConnect = sConnect & ";Integrated Security=SSPI;"
  End If

  ' Clear any existing database connection.
  Set gADOCon = Nothing
  
  ' Establish the database connection.
  Set gADOCon = New ADODB.Connection
  With gADOCon
    .ConnectionString = sConnect
    .Provider = "SQLOLEDB"
    .CommandTimeout = 0               ' Some commands on saving changes will take a huge time.
    .ConnectionTimeout = 5
    .CursorLocation = adUseServer
    .Mode = adModeReadWrite
    .Properties("Packet Size") = 32767
    .Open
  End With


  ' Get the SQL Server version number.
  glngSQLVersion = 0
  gstrSQLFullVersion = ""
  sSQL = "master..xp_msver ProductVersion"
  rsSQLInfo.Open sSQL, gADOCon, adOpenForwardOnly, adLockReadOnly
  With rsSQLInfo
    If Not (.BOF And .EOF) Then
      glngSQLVersion = Val(.Fields("character_value").Value)
      gstrSQLFullVersion = .Fields("character_value").Value
    End If
    .Close
  End With
  Set rsSQLInfo = Nothing
  
  ' The version of SQL Server is not 2008 so tell the user.
  If glngSQLVersion < 9 Then
    ' The version of SQL Server is neither 6.5 nor 7.0 so tell the user.
    Screen.MousePointer = vbNormal
    gobjProgress.CloseProgress
    MsgBox "You are running an unsupported version of SQL Server." & vbCrLf & vbCrLf & _
      "OpenHR requires SQL Server version 2005 or above.", _
      vbOKOnly, App.ProductName
    Exit Sub
  End If
  
  ' Is Windows authentication disabled on this server (have bypass on development)
  gbCanUseWindowsAuthentication = IIf(glngSQLVersion < 8, False, True)
  
  If Not gbCanUseWindowsAuthentication And gbUseWindowsAuthentication Then
    Screen.MousePointer = vbNormal
    gobjProgress.CloseProgress
    MsgBox "Windows authenticated users are not supported on your SQL Server" & vbCrLf & vbCrLf & _
      "OpenHR requires SQL Server version 2008 or above.", _
      vbExclamation, App.ProductName
    Exit Sub
  End If
  
  ' Is the server configured only for Windows logins. Sometimes SQL7 automatically logs you in using Windows authentication
  '   if the server is confiigured for Windows Only (i.e. it ignores the username and password in the connect string - ANNOYING!!!!)
  giSQLServerAuthenticationType = iWINDOWSONLY
  sSQL = "master..xp_loginconfig 'login mode'"
  rsSQLInfo.Open sSQL, gADOCon, adOpenForwardOnly, adLockReadOnly
  With rsSQLInfo
    If Not (.BOF And .EOF) Then
      Select Case !config_value
        Case "Mixed"
          giSQLServerAuthenticationType = iMIXEDMODE
        Case "Windows NT Authentication"
          giSQLServerAuthenticationType = iWINDOWSONLY
      End Select
    End If
    .Close
  End With
  Set rsSQLInfo = Nothing

  ' Only allow sql login if server is configured this way (see comments above)
  If giSQLServerAuthenticationType = iWINDOWSONLY And Not gbUseWindowsAuthentication Then
    Screen.MousePointer = vbNormal
    MsgBox "Your server is configured for Windows Only security." & vbCrLf & "Please see your system administrator." _
    , vbInformation, App.ProductName
    Exit Sub
  End If
  
  gsConnectString = sConnect
  gsUserName = sUserName
  gsPassword = sPassword
  
  ' Is the user a system administrator on the server or is logged in as 'sa'
  sSQL = "SELECT IS_SRVROLEMEMBER('sysadmin') AS Permission"
  rsUser.Open sSQL, gADOCon, adOpenForwardOnly, adLockReadOnly
  bIsSQLSystemAdmin = IIf(rsUser!Permission = 1, True, False)
  rsUser.Close
  
  ' Is the user a security administrator
  sSQL = "SELECT IS_SRVROLEMEMBER('securityadmin') AS Permission"
  rsUser.Open sSQL, gADOCon, adOpenForwardOnly, adLockReadOnly
  bIsSecurityAdmin = IIf(rsUser!Permission = 1, True, bIsSQLSystemAdmin)
  rsUser.Close
    
  If bIsSQLSystemAdmin Then
    gfCurrentUserIsSysSecMgr = True
    gsUserGroup = "<None>"
    gbUserCanManageLogins = True
  Else
    gfCurrentUserIsSysSecMgr = CurrentUserIsSysSecMgr
    gsUserGroup = CurrentUserGroup
    gbUserCanManageLogins = bIsSecurityAdmin
  End If

  ' Populate licence key
  gobjLicence.LicenceKey = GetSystemSetting("Licence", "Key", vbNullString)

  ' Misc security settings
  gbDeleteOrphanWindowsLogins = GetSystemSetting("Misc", "CFG_DELETEORPHANLOGINS", False)
  gbDeleteOrphanUsers = GetSystemSetting("Misc", "CFG_DELETEORPHANUSERS", False)
  
  ' Sometimes the ODBC API reports the username being empty so check it
  If IsEmpty(gsUserName) Then
    Err.Raise vbObjectError + 512, "Login", "Login Failed Due To ODBC Driver Error."
  End If

  ' Check the database version is the right one for the application version.
  If Not CheckVersion Then
    If Not gADOCon Is Nothing Then
      gADOCon.Close
      Set gADOCon = Nothing
    End If

    Exit Sub
  End If
  
  ' If we're logged in under a windows authenticated group
  If gbUseWindowsAuthentication Then
    gsActualSQLLogin = GetActualLoginName
  Else
    gsActualSQLLogin = gsUserName
  End If
  
  'MH20010501 Needs to be done after gsUserName is set.
  Call CheckApplicationAccess
  If Application.AccessMode = accNone Then
    If Not gADOCon Is Nothing Then
      gADOCon.Close
      Set gADOCon = Nothing
    End If
    Screen.MousePointer = vbDefault
    Exit Sub
  End If

  ' Check licence details
  If Not CheckLicence() Then
    Exit Sub
  End If

  ' Save the connection details in the registry..
  SavePCSetting "Login", "SecMgr_UserName", txtUID.Text
  SavePCSetting "Login", "SecMgr_Database", txtDatabase.Text
  SavePCSetting "Login", "SecMgr_Server", txtServer.Text
  SavePCSetting "Login", "SecMgr_AuthenticationMode", chkUseWindowsAuthentication.Value

  ' Check the password
  If Not gbUseWindowsAuthentication And glngSQLVersion < 9 Then
    If LCase(gsUserName) <> "sa" Then
      Screen.MousePointer = vbDefault
      CheckPassword
      Screen.MousePointer = vbHourglass
    End If
  End If

  ' AE20080303 Fault #12766 If the users default database is not 'master' then make it so.
  sSQL = "IF EXISTS(SELECT 1 FROM master..syslogins WHERE loginname = SUSER_NAME() AND dbname <> 'master')" & vbNewLine & _
                             "  EXEC sp_defaultdb [" & gsUserName & "], master"
  gADOCon.Execute sSQL, , adCmdText

  LC_SaveSettingsToRegistry
  
TidyUpAndExit:
  Set rsUser = Nothing
  
  gobjProgress.CloseProgress
  OK = True
  Me.Hide
  Exit Sub

ErrorTrap:
'NHRD - 17042003 - Fault 5058 Added this errorfile writing function.
  Dim iErrCount As Integer
  Dim strRDOError As String
  Dim sMsg As String
  Dim iForceChangeReason As PasswordChangeReason
  
  sTempErrDescription = "Cannot login - unknown error." & vbNewLine & "Please see your system administrator."
  
  For iErrCount = 0 To gADOCon.Errors.Count - 1
     
      Select Case gADOCon.Errors(iErrCount).NativeError
        ' 14 - Invalid Connection String
        ' 17 -No such server
        ' 4060 - No such database
        ' 18456 - Inavlid username/password
        Case 14, 17, 4060, 18456
         sTempErrDescription = "The system could not log you on. Make sure your details are correct, then retype your password."
      
       ' .NET error, SQL process login details are incorrect
        Case 6522
          sTempErrDescription = "The SQL process account has not been defined or is invalid." & vbNewLine & _
            "Please contact your system administrator."
            
        Case 18452
          If gbUseWindowsAuthentication Then
            sTempErrDescription = "The system could not log you on. Make sure your details are correct, then retype your password."
          Else
            sTempErrDescription = "Your server is configured for Windows Only security." & vbNewLine & "Please see your system administrator."
          End If
                 
        Case 18486
          sTempErrDescription = gADOCon.Errors(iErrCount).Description
        
        ' Password expired / must be changed (SQL 2005)
        Case 18463, 18465, 18466, 18467
          iForceChangeReason = giPasswordChange_ComplexitySettings
        
        ' Password too short (SQL 2005)
        Case 18464
          iForceChangeReason = giPasswordChange_MinLength
        
        ' Account locked out (SQL 2005)
        Case 15113
          sMsg = gADOCon.Errors(iErrCount).Description
        
        ' Account password expired (SQL 2005)
        Case 18487
          iForceChangeReason = giPasswordChange_Expired
        
        ' Force password change (SQL 2005)
        Case 18488
          iForceChangeReason = giPasswordChange_AdminRequested
                
        Case Else
          'sTempErrDescription = gADOCon.Errors(iErrCount).Description
          sTempErrDescription = "The system could not log you on. Make sure your details are correct, then retype your password."
      
      End Select
      
  Next iErrCount
        
  On Local Error Resume Next
  
  If iForceChangeReason <> 0 Then
    
    ' Change the password
    gsUserName = txtUID.Text
    gsPassword = txtPWD.Text
    If frmChangePassword.Initialise(iForceChangeReason, 0) Then
      frmChangePassword.Show vbModal
      sMsg = vbNullString
      mbForceTotalRelogin = Not frmChangePassword.Cancelled
      Unload frmChangePassword
    End If
  
  Else
  
    strRDOError = CreateLoginErrFile(sTempErrDescription)
    
    If Len(strRDOError) > 0 Then
      sMsg = strRDOError
    ElseIf Len(Err.Description) > 0 And InStr(sMsg, Err.Description) = 0 Then
      sMsg = IIf(Len(sMsg) > 0, _
             sMsg & vbCrLf & "(" & Err.Description & ")", _
             Err.Description)
    End If
  
    If Not gADOCon Is Nothing Then
      If gADOCon.State <> adStateClosed Then
        gADOCon.Close
      End If
      Set gADOCon = Nothing
    End If
  
    gobjProgress.CloseProgress
    Screen.MousePointer = vbDefault
    If Len(sTempErrDescription) > 0 Then
      MsgBox sTempErrDescription, vbExclamation + vbOKOnly, Me.Caption & " Error"
    End If
    
    If txtPWD.Enabled Then
      txtPWD.SetFocus
    End If
    
    'TM05012004 - Failed login therefore increment the bad attempts count.
    LC_IncrementBadAttempt
  
    Err.Clear
  End If
End Sub

Public Sub CheckCommandLine()

  Dim strCommandLine() As String
  Dim strParameters() As String
  Dim strOption As String
  Dim strValue As String
  Dim lngCount As Long
  Dim blnPassword As Boolean
  
  Dim strUserName As String
  Dim strPassword As String
  Dim strDatabaseName As String
  Dim strServerName As String


  blnPassword = False
  mlngTimeOut = 30

  If Command$ <> vbNullString Then
    'Read command line into array
    strCommandLine = Split(Command$, "/")

    For lngCount = 1 To UBound(strCommandLine)
      
      strParameters = Split(strCommandLine(lngCount), "=")
      strOption = Trim(strParameters(0))
      If UBound(strParameters) > 0 Then
        strValue = Trim(strParameters(1))
      Else
        strValue = vbNullString
      End If

      Select Case LCase(strOption)
      Case "user", "username"
        txtUID.Text = strValue

      Case "pass", "password"
        txtPWD.Text = strValue
        blnPassword = True

      Case "database"
        txtDatabase.Text = strValue

      Case "server"
        txtServer.Text = strValue

      Case "details"
        FormatControls (LCase(strValue) = "true")

      Case "readonly"
        mblnForceReadOnly = (LCase(strValue) = "true")

      Case "timeout"
        mlngTimeOut = CLng(strValue)
      
      End Select

    Next

    If blnPassword Then
      'gobjProgress.AviFile = "" 'App.Path & "\videos\about.avi"
      gobjProgress.AVI = dbLogin
      gobjProgress.NumberOfBars = 0
      gobjProgress.MainCaption = "Login"
      gobjProgress.Caption = "Attempting log on to OpenHR..."
      gobjProgress.OpenProgress
      Login
    End If

  End If

End Sub


'Private Sub CheckApplicationAccess()
'
'  Dim blnCurrentlyLocked As Boolean
'  Dim blnReadWriteLock As Boolean
'  Dim strLockDetails As String
'  Dim strLockUser As String
'  Dim strLockType As String
'  Dim strMBText As String
'  Dim intMBResponse As Integer
'
'
'  rdoCon.BeginTrans
'  GetLockDetails strLockUser, strLockType
'  blnCurrentlyLocked = (strLockUser <> vbNullString)
'  blnReadWriteLock = (strLockType = "Lock Read Write")
'
'
'  'Check if we can get the system with Write Access...
'  If Not blnCurrentlyLocked Then
'    'Now have write access .. lock database so nobody else can...
'    SaveSystemSetting "Lock Read Write", "User", gsUserName
'    SaveSystemSetting "Lock Read Write", "DateTime", Format(Now, DateFormat & " hh:nn")
'    SaveSystemSetting "Lock Read Write", "Machine", UI.GetHostName
'    rdoCon.CommitTrans
'
'    Application.AccessMode = accFull
'
'  Else
'    'If not locked by current app then can we get read only access...
'    strLockDetails = "User :  " & strLockUser & vbCrLf & _
'                     "Date/Time :  " & GetSystemSetting(strLockType, "DateTime", "") & vbCrLf & _
'                     "Machine :  " & GetSystemSetting(strLockType, "Machine", "") & vbCrLf & _
'                     "Type :  " & strLockType
'    rdoCon.RollbackTrans
'
'    'Ask if user wants read only if somebody else has system read/write
'
'    Screen.MousePointer = vbNormal
'    If blnReadWriteLock Or LCase(gsUserName) = "sa" Then
'      strMBText = "The database is currently locked as follows:" & _
'                  vbCrLf & vbCrLf & strLockDetails & vbCrLf & vbCrLf & _
'                  "Would you like to proceed with read only access?"
'      intMBResponse = MsgBox(strMBText, vbYesNo + vbExclamation, Application.Name)
'
'      Application.AccessMode = IIf(intMBResponse = vbYes, accSystemReadOnly, accNone)
'    Else
'      'Database is totally locked (saving or manual lock)
'      MsgBox "Unable to login to " & gsDatabaseName & " as the database has been locked." & _
'             vbCrLf & vbCrLf & strLockDetails, vbExclamation, Application.Name
'      Application.AccessMode = accNone
'    End If
'
'  End If
'
'End Sub


Private Sub CheckApplicationAccess()

  Dim rsTemp As New ADODB.Recordset
  Dim blnCurrentlyLocked As Boolean
  Dim blnReadWriteLock As Boolean
  Dim strLockDetails As String
  Dim strLockUser As String
  Dim strLockType As String
  Dim strMBText As String
  Dim intMBResponse As Integer
  
  
  'Check security permissions
  ' AE20090325 Fault #13628
  'If (SystemPermission("MODULEACCESS", "SECURITYMANAGER") = True Or LCase(gsUserName) = "sa" Or gfCurrentUserIsSysSecMgr) And Not mblnForceReadOnly Then
  If (SystemPermission("MODULEACCESS", "SECURITYMANAGER") = True Or LCase(gsUserName) = "sa") And gfCurrentUserIsSysSecMgr And (Not mblnForceReadOnly) Then
    Application.AccessMode = accFull
  ElseIf SystemPermission("MODULEACCESS", "SECURITYMANAGERRO") = True Then
    Application.AccessMode = accSystemReadOnly
  Else
    Application.AccessMode = accNone
    Screen.MousePointer = vbNormal
    MsgBox "You do not have permission to run the Security Manager." & vbCrLf & vbCrLf & _
           "Please contact your OpenHR security administrator.", vbOKOnly + vbExclamation, App.ProductName
    Exit Sub
  End If


  ' Begin transaction
  'gADOCon.BeginTrans
  
  rsTemp.Open "sp_ASRLockCheck", gADOCon, adOpenForwardOnly, adLockReadOnly
  blnCurrentlyLocked = (Not rsTemp.BOF And Not rsTemp.EOF)
  
  If blnCurrentlyLocked Then
    'Ignore users own manual lock
    If LCase(gsUserName) = LCase(rsTemp!UserName) And rsTemp!Priority = lckManual Then
      rsTemp.MoveNext
    End If
    blnCurrentlyLocked = (Not rsTemp.BOF And Not rsTemp.EOF)

    If blnCurrentlyLocked Then
  
      'If not locked by current app then can we get read only access...
      strLockDetails = "User :  " & rsTemp!UserName & vbCrLf & _
                       "Date/Time :  " & rsTemp!Lock_Time & vbCrLf & _
                       "Machine :  " & rsTemp!HostName & vbCrLf & _
                       "Type :  " & rsTemp!Description
  
      Screen.MousePointer = vbNormal
      
      If rsTemp!Priority = lckReadWrite Or _
        (rsTemp!Priority = lckManual And LCase(gsUserName) = "sa") Then
        'Ask if user wants read only if somebody else has system read/write
  
        If Application.AccessMode = accFull Then
          strMBText = "The database is currently locked as follows:" & _
                      vbCrLf & vbCrLf & strLockDetails & vbCrLf & vbCrLf & _
                      "Would you like to proceed with read only access?"
          intMBResponse = MsgBox(strMBText, vbYesNo + vbExclamation, Application.Name)
          Application.AccessMode = IIf(intMBResponse = vbYes, accSystemReadOnly, accNone)
        End If
  
      Else
        'Database is totally locked (saving or manual lock)
        MsgBox "Unable to login to '" & gsDatabaseName & "' as the database has been locked." & _
               vbCrLf & vbCrLf & strLockDetails, vbExclamation, Application.Name
        Application.AccessMode = accNone
  
      End If

      Screen.MousePointer = vbHourglass

    End If
  
  End If
  rsTemp.Close
  
  If Application.AccessMode = accFull Then
    'Now have write access .. lock database so nobody else can...
    'SaveSystemSetting "Lock Read Write", "User", gsUserName
    'SaveSystemSetting "Lock Read Write", "DateTime", Format(Now, DateFormat & " hh:nn")
    'SaveSystemSetting "Lock Read Write", "Machine", UI.GetHostName
    gADOCon.Execute "sp_ASRLockWrite 3", , adExecuteNoRecords
    'gADOCon.CommitTrans
  Else
    'gADOCon.RollbackTrans
  End If
  ' End Transaction

  Set rsTemp = Nothing

End Sub
Public Function CreateLoginErrFile(sMsg As String) As String
'NHRD - 17042003 - Fault 5058 Added this errorfile writing function.
'Nicked from Data manager frmLogin

  'This sub will return the last ADO error.
  CreateLoginErrFile = vbNullString

  Const lngFileNum As Integer = 99
  Dim lngCount As Long
  On Local Error Resume Next

  Open App.Path & "\LoginErr.txt" For Output As #lngFileNum
  Print #lngFileNum, "Server    : " & txtServer.Text
  Print #lngFileNum, "Database  : " & txtDatabase.Text
  Print #lngFileNum, "Username  : " & txtUID.Text
  Print #lngFileNum, "Version   : " & CStr(App.Major) & "." & CStr(App.Minor) & "." & CStr(App.Revision)
  Print #lngFileNum, ""

  Print #lngFileNum, "Date/Time : " & CStr(Now)
  If Not gADOCon Is Nothing Then
    ' JPD20030211 Fault 5044
    Print #lngFileNum, "Connection: " & IIf(gADOCon.StillExecuting = True, "Still Executing", "Finished Executing")
  Else
    Print #lngFileNum, "Connection: Failed"
  End If
  Print #lngFileNum, ""

  Print #lngFileNum, sMsg
  Print #lngFileNum, Err.Description
  Print #lngFileNum, ""

  If Not gADOCon Is Nothing Then
    If gADOCon.Errors.Count > 0 Then
      Print #lngFileNum, "ADO Connection Errors"
      Print #lngFileNum, "=====================" & vbCrLf
      
      For lngCount = 0 To gADOCon.Errors.Count - 1
        CreateLoginErrFile = gADOCon.Errors(lngCount).Description
        Print #lngFileNum, CreateLoginErrFile
      Next
    End If
  Else
    CreateLoginErrFile = "Connection failed"
  End If
  
  Close #lngFileNum

  CreateLoginErrFile = Mid(CreateLoginErrFile, InStrRev(CreateLoginErrFile, "]") + 1)

  
End Function

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
  Case vbKeyF1
    If ShowAirHelp(Me.HelpContextID) Then
      KeyCode = 0
    End If
  Case ((Shift And vbShiftMask) > 0)
    gbShiftSave = True
End Select
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
  gbShiftSave = False
End Sub

Private Function GetActualLoginName() As String

  Dim cmdDetail As New ADODB.Command
  Dim pmADO As ADODB.Parameter

  With cmdDetail
    .CommandText = "spASRGetActualUserDetails"
    .CommandType = adCmdStoredProc
    .CommandTimeout = 0
    Set .ActiveConnection = gADOCon

    Set pmADO = .CreateParameter("UserName", adVarChar, adParamOutput, 255)
    .Parameters.Append pmADO

    Set pmADO = .CreateParameter("Group", adVarChar, adParamOutput, 255)
    .Parameters.Append pmADO

    Set pmADO = .CreateParameter("GroupID", adInteger, adParamOutput, 50)
    .Parameters.Append pmADO

    ' AE20090311 Fault #13598
    Set pmADO = .CreateParameter("ModuleKey", adVarChar, adParamInput, 20, "SECURITYMANAGER")
    .Parameters.Append pmADO
    
    .Execute

    GetActualLoginName = IIf(IsNull(.Parameters(0).Value), "", .Parameters(0).Value)
  End With
  
  Set cmdDetail = Nothing

End Function

Public Function CheckLicence() As Boolean

  Dim sMsg As String
  Dim lActualCount As Long
  Dim lngCurrentHeadcount As Long
  Dim dToday As Date
  
  On Error GoTo Err_Trap
  
  CheckLicence = False
  dToday = DateValue(Now)
    
  ' Expiry date checks
  If gobjLicence.HasExpiryDate Then
    If (dToday > gobjLicence.ExpiryDate) Then
      sMsg = "Your licence to use this product has expired." & vbNewLine & _
            "Please contact OpenHR Customer Services on 08451 609 999 as soon as possible."
      GoTo Exit_Fail:
      
    End If
            
    If (dToday > DateAdd("d", -7, gobjLicence.ExpiryDate)) Then
      sMsg = "Your licence to use this product will expire on " & gobjLicence.ExpiryDate & "." & vbNewLine & vbNewLine & _
            "Please contact OpenHR Customer Services on 08451 609 999 as soon as possible."
      MsgBox sMsg, vbInformation
    End If
  End If
    
     
  ' Headcount checks
  If gobjLicence.Headcount > 0 Then
    Select Case gobjLicence.LicenceType
      Case LicenceType.Headcount, LicenceType.DMIConcurrencyAndHeadcount
        lngCurrentHeadcount = GetSystemSetting("Headcount", "current", 0)
        
      Case LicenceType.P14Headcount, LicenceType.DMIConcurrencyAndP14
       lngCurrentHeadcount = GetSystemSetting("Headcount", "P14", 0)

    End Select
   
    If lngCurrentHeadcount >= gobjLicence.Headcount Then
      sMsg = "You have reached or exceeded the headcount limit set within the terms of your licence agreement." & vbNewLine & vbNewLine & _
                            "You are no longer able to add new employee records, but you may access the system for other purposes." & vbNewLine & vbNewLine & _
                            "Please contact OpenHR Customer Services on 08451 609 999 as soon as possible to increase the licence headcount number."
      MsgBox sMsg, vbCritical
    
    ElseIf lngCurrentHeadcount >= gobjLicence.Headcount * 0.95 Then
      
      If DisplayWarningToUser(gsUserName, Headcount95Percent, 7) Then
        sMsg = "You are currently within 95% (" & lngCurrentHeadcount & " of " & gobjLicence.Headcount & " employees) of reaching the headcount limit set within the terms of your licence agreement." & vbNewLine & vbNewLine & _
                              "Once this limit is reached, you will no longer be able to add new employee records to the system." & vbNewLine & vbNewLine & _
                              "If you wish to increase the headcount number, please contact OpenHR Customer Services on 08451 609 999 as soon as possible."
        MsgBox sMsg, vbInformation
      End If
    
    End If
  End If
    
Exit_Ok:
  CheckLicence = True
  Exit Function
    
Exit_Fail:
  Screen.MousePointer = vbDefault
  MsgBox sMsg, vbCritical
  CheckLicence = False
  Exit Function

Err_Trap:
  CheckLicence = False

End Function

Private Function DisplayWarningToUser(UserName As String, WarningType As WarningType, warningRefreshRate As Integer) As Boolean

  Dim cmADO As ADODB.Command
  Dim pmADO As ADODB.Parameter
  Dim bResult As Boolean

  On Error GoTo ErrorTrap
  
  ' Run the stored procedure to see if the given record has changed
  Set cmADO = New ADODB.Command
  With cmADO
    .CommandText = "spASRUpdateWarningLog"
    .CommandType = adCmdStoredProc
    .CommandTimeout = 0
    Set .ActiveConnection = gADOCon
                      
    Set pmADO = .CreateParameter("Username", adVarChar, adParamInput, 255)
    .Parameters.Append pmADO
    pmADO.Value = UserName
    
    Set pmADO = .CreateParameter("WarningType", adInteger, adParamInput)
    .Parameters.Append pmADO
    pmADO.Value = WarningType
          
    Set pmADO = .CreateParameter("WarningRefreshRate", adInteger, adParamInput)
    .Parameters.Append pmADO
    pmADO.Value = warningRefreshRate
          
    Set pmADO = .CreateParameter("WarnUser", adBoolean, adParamOutput)
    .Parameters.Append pmADO
          
    Set pmADO = Nothing

    cmADO.Execute
    bResult = CBool(.Parameters(3).Value)
    
  End With
               
TidyUpAndExit:
  DisplayWarningToUser = bResult
  Exit Function

ErrorTrap:
  bResult = False
  GoTo TidyUpAndExit

End Function



