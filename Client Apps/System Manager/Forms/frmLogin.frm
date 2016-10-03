VERSION 5.00
Object = "{BD0C1912-66C3-49CC-8B12-7B347BF6C846}#13.1#0"; "Codejock.SkinFramework.v13.1.0.ocx"
Begin VB.Form frmLogin 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "OpenHR System Manager - Login"
   ClientHeight    =   3675
   ClientLeft      =   1530
   ClientTop       =   2595
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
   HelpContextID   =   5003
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
   Begin VB.CheckBox chkRecover 
      Caption         =   "Recover?"
      Height          =   195
      Left            =   3915
      TabIndex        =   13
      Top             =   1215
      Width           =   1215
   End
   Begin VB.CheckBox chkUseWindowsAuthentication 
      Caption         =   "&Use Windows Authentication"
      Height          =   210
      Left            =   240
      TabIndex        =   2
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
      TabIndex        =   5
      Top             =   1680
      Width           =   1200
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   400
      Left            =   4950
      TabIndex        =   6
      Top             =   2235
      Width           =   1200
   End
   Begin VB.TextBox txtUID 
      Height          =   315
      Left            =   1515
      TabIndex        =   0
      Top             =   1680
      Width           =   3285
   End
   Begin VB.TextBox txtPWD 
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   1515
      MaxLength       =   128
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   2085
      Width           =   3285
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   65535
      Left            =   5010
      Top             =   3000
   End
   Begin VB.TextBox txtDatabase 
      Height          =   315
      Left            =   1515
      TabIndex        =   3
      Top             =   2835
      Visible         =   0   'False
      Width           =   3285
   End
   Begin VB.TextBox txtServer 
      Height          =   315
      Left            =   1515
      TabIndex        =   4
      Top             =   3240
      Visible         =   0   'False
      Width           =   3285
   End
   Begin VB.Image imgASRLogo 
      Height          =   1065
      Left            =   240
      Picture         =   "frmLogin.frx":000C
      Top             =   120
      Width           =   4575
   End
   Begin VB.Label lblDevelopmentMode 
      AutoSize        =   -1  'True
      Caption         =   "(Dev mode)"
      Height          =   195
      Left            =   5145
      TabIndex        =   14
      Top             =   1215
      Width           =   1035
   End
   Begin XtremeSkinFramework.SkinFramework SkinFramework1 
      Left            =   5655
      Top             =   3075
      _Version        =   851969
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
   Begin VB.Line lnTopWhiteLine 
      BorderColor     =   &H80000005&
      X1              =   240
      X2              =   6120
      Y1              =   1515
      Y2              =   1515
   End
   Begin VB.Label lblVersion 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "lblVersion"
      Height          =   195
      Left            =   240
      TabIndex        =   12
      Top             =   1215
      Width           =   840
   End
   Begin VB.Line lnTopGreyLine 
      BorderColor     =   &H80000010&
      BorderWidth     =   2
      X1              =   240
      X2              =   6100
      Y1              =   1515
      Y2              =   1515
   End
   Begin VB.Label lblUser 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Username :"
      Height          =   195
      Left            =   240
      TabIndex        =   11
      Top             =   1740
      Width           =   1005
   End
   Begin VB.Label lblPassword 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Password :"
      Height          =   195
      Left            =   240
      TabIndex        =   10
      Top             =   2145
      Width           =   945
   End
   Begin VB.Label lblDatabase 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Database :"
      Height          =   195
      Left            =   240
      TabIndex        =   9
      Top             =   2895
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.Label lblServer 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Server :"
      Height          =   195
      Left            =   240
      TabIndex        =   7
      Top             =   3300
      Visible         =   0   'False
      Width           =   810
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Implements IObjectSafetyTLB.IObjectSafety

' API functions.
Private Declare Function SQLAllocEnv% Lib "odbc32.dll" (env&)
Private Declare Function SQLDrivers Lib "odbc32.dll" (ByVal henv&, ByVal fDirection%, ByVal szDriverDesc$, ByVal cbDriverDescMax%, pcbDriverDescBuff%, ByVal szDriverAttrib$, ByVal cbDriverAttribMax%, pcbDriverAttribBuff%) As Integer

' API constants.
Const SQL_SUCCESS As Long = 0
Const SQL_FETCH_NEXT As Long = 1

' ODBC constants.
Const ODBCDRIVER As String = "SQL Server"

' Private classes.
Private Net As New Net

' Globals.
Private gDatStartDate As Date
Private mblnForceReadOnly As Boolean
Private mfReRunScript As Boolean
Private mlngTimeOut As Long
Public OK As Boolean
Private mbForceTotalRelogin As Boolean

Private Sub IObjectSafety_GetInterfaceSafetyOptions(ByVal riid As Long, _
                                                    pdwSupportedOptions As Long, _
                                                    pdwEnabledOptions As Long)

    Dim rc      As Long
    Dim rClsId  As udtGUID
    Dim IID     As String
    Dim bIID()  As Byte

    pdwSupportedOptions = INTERFACESAFE_FOR_UNTRUSTED_CALLER Or _
                          INTERFACESAFE_FOR_UNTRUSTED_DATA

    If (riid <> 0) Then
        CopyMemory rClsId, ByVal riid, Len(rClsId)

        bIID = String$(MAX_GUIDLEN, 0)
        rc = StringFromGUID2(rClsId, VarPtr(bIID(0)), MAX_GUIDLEN)
        rc = InStr(1, bIID, vbNullChar) - 1
        IID = Left$(UCase(bIID), rc)

        Select Case IID
            Case IID_IDispatch
                pdwEnabledOptions = IIf(m_fSafeForScripting, _
              INTERFACESAFE_FOR_UNTRUSTED_CALLER, 0)
                Exit Sub
            Case IID_IPersistStorage, IID_IPersistStream, _
               IID_IPersistPropertyBag
                pdwEnabledOptions = IIf(m_fSafeForInitializing, _
              INTERFACESAFE_FOR_UNTRUSTED_DATA, 0)
                Exit Sub
            Case Else
                Err.Raise E_NOINTERFACE
                Exit Sub
        End Select
    End If
    
End Sub

Private Sub IObjectSafety_SetInterfaceSafetyOptions(ByVal riid As Long, _
                                                    ByVal dwOptionsSetMask As Long, _
                                                    ByVal dwEnabledOptions As Long)
    Dim rc          As Long
    Dim rClsId      As udtGUID
    Dim IID         As String
    Dim bIID()      As Byte

    If (riid <> 0) Then
        CopyMemory rClsId, ByVal riid, Len(rClsId)

        bIID = String$(MAX_GUIDLEN, 0)
        rc = StringFromGUID2(rClsId, VarPtr(bIID(0)), MAX_GUIDLEN)
        rc = InStr(1, bIID, vbNullChar) - 1
        IID = Left$(UCase(bIID), rc)

        Select Case IID
            Case IID_IDispatch
                If ((dwEnabledOptions And dwOptionsSetMask) <> _
             INTERFACESAFE_FOR_UNTRUSTED_CALLER) Then
                    Err.Raise E_FAIL
                    Exit Sub
                Else
                    If Not m_fSafeForScripting Then
                        Err.Raise E_FAIL
                    End If
                    Exit Sub
                End If

            Case IID_IPersistStorage, IID_IPersistStream, _
          IID_IPersistPropertyBag
                If ((dwEnabledOptions And dwOptionsSetMask) <> _
              INTERFACESAFE_FOR_UNTRUSTED_DATA) Then
                    Err.Raise E_FAIL
                    Exit Sub
                Else
                    If Not m_fSafeForInitializing Then
                        Err.Raise E_FAIL
                    End If
                    Exit Sub
                End If

            Case Else
                Err.Raise E_NOINTERFACE
                Exit Sub
        End Select
    End If
    
End Sub

Private Sub chkRecover_Click()
  gbAttemptRecovery = IIf(chkRecover.value = vbChecked, True, False)
End Sub

Private Sub chkUseWindowsAuthentication_Click()

  Dim strTrustedUser As String

  strTrustedUser = gstrWindowsCurrentDomain & "\" & gstrWindowsCurrentUser

  txtUID.Text = IIf(chkUseWindowsAuthentication.value = vbChecked, strTrustedUser, txtUID.Text)
  txtUID.Enabled = IIf(chkUseWindowsAuthentication.value = vbChecked, False, True)
  txtPWD.Enabled = IIf(chkUseWindowsAuthentication.value = vbChecked, False, True)

  ' Grey out controls
  txtUID.BackColor = IIf(txtUID.Enabled, vbWindowBackground, vbButtonFace)
  txtPWD.BackColor = IIf(txtPWD.Enabled, vbWindowBackground, vbButtonFace)

End Sub

'Private Function GetRightChars(strInput As String, strSearch As String) As String
'  GetRightChars = Mid(strInput, InStrRev(strInput, strSearch) + 1)
'End Function

Private Sub cmdDetails_Click()
  ' Display/hide the Database and Server controls.
  Dim fDisplayControls As Boolean
  
  fDisplayControls = Not lblDatabase.Visible
  
  FormatControls fDisplayControls
  
  With txtDatabase
    If .Visible Then
      .SetFocus
    End If
  End With

End Sub


Private Sub Form_KeyPress(KeyAscii As Integer)
  ' Record the time the user last did anything on this screen.
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

  lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision

  If ASRDEVELOPMENT Then
    lblDevelopmentMode.Caption = "(Dev mode)"
    lblDevelopmentMode.Visible = True
    lblDevelopmentMode.Enabled = True
    chkRecover.Visible = True
  Else
    lblDevelopmentMode.Caption = ""
    lblDevelopmentMode.Visible = False
    lblDevelopmentMode.Enabled = False
    chkRecover.Visible = False
  End If
  
  gblnAutomaticLogon = False
  
  ' Check that the SQL Server driver is installed.
  If Not CheckSQLDriver Then
    MsgBox "The required ODBC Driver '" & ODBCDRIVER & "' is not installed." & vbNewLine & _
      "Install the driver before running System Manager.", _
      vbExclamation + vbOKOnly, App.ProductName
    UnLoad frmLogin
    Exit Sub
  End If

  ' Get the details of the last connection from the registry.
  sUserName = GetPCSetting("Login", "SysMgr_UserName", Net.userName)
  sDatabaseName = GetPCSetting("Login", "SysMgr_Database", vbNullString)
  sServerName = GetPCSetting("Login", "SysMgr_Server", vbNullString)
  bUseWindowsAuthentication = GetPCSetting("Login", "SysMgr_AuthenticationMode", 0)

  ' If the database/server information is missing from the registry then display the
  ' controls.
  fDatabaseInfoMissing = (sDatabaseName = vbNullString) Or _
    (sServerName = vbNullString)
  
  FormatControls fDatabaseInfoMissing
  
  txtUID.Text = sUserName
  txtDatabase.Text = sDatabaseName
  txtServer.Text = sServerName
  chkUseWindowsAuthentication.value = IIf(bUseWindowsAuthentication = True, vbChecked, vbUnchecked)

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
  '  imgASRLogo.Visible = True
  '  Resume Next
  'End If
  Resume TidyUpAndExit
  
End Sub

Private Sub Form_Activate()
  
  ' If we already have a UID, then set the focus on the password text control.
  If Me.Visible Then
    If Len(txtUID.Text) > 0 Then
      If txtPWD.Enabled Then
        txtPWD.SetFocus
      End If
    End If
  End If

  ' Normal/Windows authentication
  chkUseWindowsAuthentication_Click

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

Select Case KeyCode
  Case vbKeyF1
    If ShowAirHelp(Me.HelpContextID) Then
      KeyCode = 0
    End If
End Select

  If ((Shift And vbShiftMask) > 0) Then
    mfReRunScript = True
  End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
  mfReRunScript = False
End Sub

Private Sub cmdOK_Click()
  Login
  
  ' SQL 2005 may have forced us to change the password, so we need to re-attempt the logon process in full
  If mbForceTotalRelogin Then
    txtPWD.Text = gsPassword
    Login
  End If
  
  mfReRunScript = False
  
End Sub

Public Sub Login()

  Dim fOk As Boolean
  Dim iLoop As Integer
  Dim sSQL As String
  Dim sConnect As String
  Dim rsUserInfo As New ADODB.Recordset
  Dim rsSQLInfo As New ADODB.Recordset
  Dim sTempErrDescription As String
  Dim strRDOError As String
  Dim objNET As New SystemMgr.Net
  Dim sMsg As String

  fOk = True

  On Error GoTo ErrorTrap
  
  Screen.MousePointer = vbHourglass
  
  ' Get temporary file path - need to remove \ to make consistent with App.Path function
  gsLogDirectory = Space(1024)
  Call GetTempPath(1024, gsLogDirectory)
  gsLogDirectory = Mid(gsLogDirectory, 1, Len(Trim(gsLogDirectory)) - 2)

  glngSQLVersion = 0
  gbUseWindowsAuthentication = chkUseWindowsAuthentication.value
  gsDatabaseName = Replace(txtDatabase.Text, ";", "")
  gsUserName = Replace(txtUID.Text, ";", "")
  gsActualSQLLogin = gsUserName
  gsPassword = Replace(txtPWD.Text, ";", "")
  
  'check if the database name has appostrophes in it!
  If InStr(1, gsDatabaseName, "'") > 0 Then
    gobjProgress.CloseProgress
    Screen.MousePointer = vbDefault
    MsgBox "Error logging in." & vbNewLine & vbNewLine & _
      "The database name contains an apostrophe.", _
      vbOKOnly + vbExclamation, Application.Name
    txtDatabase.SetFocus
    Exit Sub
  End If

  ' Build the ODBC connection string.
  ' Specify the Driver in the connection string.
  sConnect = sConnect & "Driver=SQL Server;"
  
  ' Specify the Server in the connection string.
  If Len(Trim(txtServer.Text)) > 0 Then
    sConnect = sConnect & "Server={" & txtServer.Text & "};"
    gsServerName = txtServer.Text
  Else
    Screen.MousePointer = vbDefault
    MsgBox "Please enter the name of the server on which the OpenHR database is located.", _
      vbExclamation + vbOKOnly, App.ProductName
    Exit Sub
  End If
    
  ' Specify the User Login in the connection string.
  If Len(gsUserName) > 0 Then
    sConnect = sConnect & "UID=" & gsUserName
  Else
    Screen.MousePointer = vbDefault
    MsgBox "Please enter a user name.", _
      vbExclamation + vbOKOnly, App.ProductName
    Exit Sub
  End If
    
  ' Specify the Password in the connection string.
  sConnect = sConnect & ";PWD=" & gsPassword & ";"
    
  ' Specify the Database name in the connection string.
  If Len(txtDatabase.Text) > 0 Then
    sConnect = sConnect & "Database=" & txtDatabase.Text & ";"
  Else
    Screen.MousePointer = vbDefault
    MsgBox "Please enter the name of the OpenHR database.", _
      vbExclamation + vbOKOnly, App.ProductName
    Exit Sub
  End If
    
  'Windows Authentication
  If gbUseWindowsAuthentication Then
    If Not objNET.MakeDSN(txtServer.Text, txtDatabase.Text) Then
      Screen.MousePointer = vbDefault
      gobjProgress.CloseProgress
      MsgBox "Error creating System DSN entry for HRPro." & vbNewLine & vbNewLine _
             & "Please contact your system administrator.", _
              vbOKOnly & vbExclamation, App.ProductName
      Set objNET = Nothing
      Exit Sub
    End If

    sConnect = "DSN=" & gstrWindowsAuthentication_DNSName & ";UID=;PWD=;Database=" & txtDatabase.Text & ";Server=" & txtServer.Text & ";Integrated Security=SSPI;"
  End If

  'sConnect = sConnect & "APP=" & App.ProductName & ";"
  sConnect = sConnect & "Application Name=" & App.ProductName & ";"

  ' Clear any existing database connection.
  On Error GoTo NoFramework
  Set gADOCon = Nothing
  Set gobjHRProEngine = New SystemFramework.SysMgr
  gobjHRProEngine.Initialise
  On Error GoTo ErrorTrap
  
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
'  sSQL = "master..xp_msver ProductVersion"
'  rsSQLInfo.Open sSQL, gADOCon, adOpenForwardOnly, adLockReadOnly
'  With rsSQLInfo
'    If Not (.BOF And .EOF) Then
'      glngSQLVersion = Val(.Fields("character_value").Value)
'      gstrSQLFullVersion = .Fields("character_value").Value
'    End If
'    .Close
'  End With
'  Set rsSQLInfo = Nothing
  sSQL = "SELECT SERVERPROPERTY('ProductVersion')"
  
  Set rsSQLInfo = New ADODB.Recordset
  rsSQLInfo.Open sSQL, gADOCon, adOpenForwardOnly, adLockReadOnly
  With rsSQLInfo
    If Not (.BOF And .EOF) Then
      glngSQLVersion = val(.Fields(0).value)
      gstrSQLFullVersion = CStr(.Fields(0).value)
    End If
    .Close
  End With
  Set rsSQLInfo = Nothing

'  If Not IsVersion10 Then
'    ' The version of SQL Server is below 2008
'    Screen.MousePointer = vbDefault
'    MsgBox "You are running an unsupported version of SQL Server." & vbNewLine & vbNewLine & _
'      "OpenHR requires SQL Server version 2008 or above.", _
'      vbOKOnly, App.ProductName
'    Exit Sub
'  End If

  ' Is Windows authentication disabled on this server (have bypass on development)
  gbCanUseWindowsAuthentication = IIf(glngSQLVersion < 9, False, True)
  
  If Not gbCanUseWindowsAuthentication And gbUseWindowsAuthentication Then
    Screen.MousePointer = vbDefault
    gobjProgress.CloseProgress
    MsgBox "Windows authenticated users are not supported on your SQL Server" & vbNewLine & vbNewLine & _
      "OpenHR requires SQL Server version 2005 or above.", _
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
    Screen.MousePointer = vbDefault
    MsgBox "Your server is configured for Windows Only security." & vbNewLine & "Please see your system administrator." _
    , vbInformation, App.ProductName
    Exit Sub
  End If

TryUsingGroupSecurity:

  If Not gblnAutomaticLogon Then
    ' Save the connection details in the registry..
    SavePCSetting "Login", "SysMgr_UserName", txtUID.Text
    SavePCSetting "Login", "SysMgr_Database", txtDatabase.Text
    SavePCSetting "Login", "SysMgr_Server", txtServer.Text
    SavePCSetting "Login", "SysMgr_AuthenticationMode", chkUseWindowsAuthentication.value
  End If

  gsServerName = Trim(txtServer.Text)
  
  ' Is the user a system administrator on the server or is logged in as 'sa'
  sSQL = "SELECT IS_SRVROLEMEMBER('sysadmin') AS Permission"
  Set rsUserInfo = New ADODB.Recordset
  rsUserInfo.Open sSQL, gADOCon, adOpenForwardOnly, adLockReadOnly
  gbIsUserSystemAdmin = IIf(rsUserInfo!Permission = 1, True, False)

  Set rsUserInfo = Nothing

  If gbIsUserSystemAdmin Then
    gbCurrentUserIsSysSecMgr = True
    gsSecurityGroup = "<None>"
  Else
    gbCurrentUserIsSysSecMgr = CurrentUserIsSysSecMgr
    gsSecurityGroup = CurrentUserGroup
  End If

  ' Populate licence key
  gobjLicence.LicenceKey = GetSystemSetting("Licence", "Key", vbNullString)

  ' Check the database version is the right one for the application version.
  If Not CheckVersion(sConnect, mfReRunScript, gbIsUserSystemAdmin) Then
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
  
  If Not (LCase(gsUserName) = "sa" Or gbUseWindowsAuthentication) And glngSQLVersion < 9 Then
  
    'JDM - 04/12/01 - Fault 3257 - Nice pretty hourglass is now off
    Screen.MousePointer = vbDefault
    CheckPassword
    Screen.MousePointer = vbHourglass
'  Else
'    'If 'sa' then make sure that we have permission on these SPs...
'    sSQL = "USE master" & vbnewline & _
'           "GRANT ALL ON sp_OACreate TO public" & vbnewline & _
'           "GRANT ALL ON sp_OADestroy TO public" & vbnewline & _
'           "GRANT ALL ON sp_OAGetErrorInfo TO public" & vbnewline & _
'           "GRANT ALL ON sp_OAGetProperty TO public" & vbnewline & _
'           "GRANT ALL ON sp_OAMethod TO public" & vbnewline & _
'           "GRANT ALL ON sp_OASetProperty TO public" & vbnewline & _
'           "GRANT ALL ON sp_OAStop TO public" & vbnewline & _
'           "GRANT ALL ON xp_StartMail TO public" & vbnewline & _
'           "GRANT ALL ON xp_SendMail TO public" & vbnewline & _
'           "USE " & gsDatabaseName
'    rdoCon.Execute sSQL
  End If

  ' AE20080303 Fault #12766 If the users default database is not 'master' then make it so.
  sSQL = "IF EXISTS(SELECT 1 FROM master..syslogins WHERE loginname = SUSER_NAME() AND dbname <> 'master')" & vbNewLine & _
                             "  EXEC sp_defaultdb [" & gsUserName & "], master"
  gADOCon.Execute sSQL, , adCmdText

  LC_SaveSettingsToRegistry

TidyUpAndExit:
  gobjProgress.CloseProgress
  OK = True
  Me.Hide
  Exit Sub

ErrorTrap:
  Dim iErrCount As Integer
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
    
    ' Login exists, but no user name attached - force retry using windows groups security
    Case 15198
      GoTo TryUsingGroupSecurity
    
    Case Else
      sTempErrDescription = gADOCon.Errors(iErrCount).Description
  
  End Select
  
  Next iErrCount
  
  On Local Error Resume Next

  If iForceChangeReason <> 0 Then
    
    'MH20061025 Fault 11625
    'gsUserName = txtUID.Text
    'gsPassword = txtPWD.Text
    If frmChangePassword.Initialise(iForceChangeReason, 0) Then
      frmChangePassword.Show vbModal
      sMsg = vbNullString
      mbForceTotalRelogin = Not frmChangePassword.Cancelled
      Set gADOCon = Nothing
      UnLoad frmChangePassword
    End If
  
  Else
  
    strRDOError = CreateLoginErrFile(sTempErrDescription)
    
    If LenB(strRDOError) <> 0 Then
      sMsg = strRDOError
    ElseIf Len(Err.Description) > 0 And InStr(sMsg, Err.Description) = 0 Then
      sMsg = IIf(LenB(sMsg) <> 0, _
             sMsg & vbNewLine & "(" & Err.Description & ")", _
             Err.Description)
    End If
  
    If Not gADOCon Is Nothing Then
      If gADOCon.State = adStateOpen Then
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
  
  Exit Sub

NoFramework:
  MsgBox "The System Framework is not installed." & vbNewLine & vbNewLine & _
    "Contact your System Administrator to install the latest System Framework" & vbNewLine & vbNewLine _
    , vbExclamation + vbOKOnly, Application.Name
  Screen.MousePointer = vbDefault

End Sub


Private Sub cmdCancel_Click()
  ' Quit the system.
  OK = False
  Me.Hide
  
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloaSMode As Integer)
  If UnloaSMode <> vbFormCode Then
    OK = False
  End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Set UI = Nothing
  Set Net = Nothing
End Sub

Private Sub Timer1_Timer()
  ' If the user has done nothing for 5 minutes then cancel the login.
  If DateDiff("n", Now, gDatStartDate) >= 5 Then
    Timer1.Enabled = False
    cmdCancel_Click
  End If
  
End Sub

Private Sub txtDatabase_GotFocus()
  ' Select all of the text in the textbox.
  UI.txtSelText

End Sub

Private Sub txtPWD_GotFocus()
  ' Select all of the text in the textbox.
  UI.txtSelText
  
End Sub

Private Sub txtServer_GotFocus()
  ' Select all of the text in the textbox.
  UI.txtSelText

End Sub

Private Sub txtUID_GotFocus()
  ' Select all of the text in the textbox.
  UI.txtSelText
  
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

Private Function CheckSQLDriver() As Boolean
  ' Check that the SQL ODBC Driver is installed.
  On Error GoTo ErrorTrap
    
  Dim fOk As Boolean
  Dim iReturnCode As Integer
  Dim iDriverDescLen As Integer
  Dim iDriverAttribLen As Integer
  Dim lngEnvironmentHandle As Long
  Dim sDriver As String
  Dim sDriverDesc As String * 1024
  Dim sDriverAttrib As String * 1024

  fOk = False
  
  ' Get the DSN and driver information.
  If SQLAllocEnv(lngEnvironmentHandle) <> -1 Then
  
    Do Until (iReturnCode <> SQL_SUCCESS) Or fOk
      
      sDriverDesc = Space(1024)
      
      iReturnCode = SQLDrivers(lngEnvironmentHandle, SQL_FETCH_NEXT, _
        sDriverDesc, 1024, iDriverDescLen, _
        sDriverAttrib, 1024, iDriverAttribLen)

      sDriver = UCase(Trim(Left(sDriverDesc, iDriverDescLen)))
                
      If Len(sDriver) > 0 Then
        If sDriver = UCase(ODBCDRIVER) Then
          fOk = True
        End If
      End If
    Loop
  End If
  
TidyUpAndExit:
  CheckSQLDriver = fOk
  Exit Function
  
ErrorTrap:
  fOk = False
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

Private Sub DisplayOtherUsers(ByVal pasUsers As Variant)
  ' Format the from controls to display the array of users passed in.
  Dim iLoop As Integer
  Dim sDisplay As String
  
  ' First warning row.
  sDisplay = "The following user(s) are currently logged into the OpenHR databases :" & vbNewLine
  
  ' User rows.
  For iLoop = 1 To UBound(pasUsers, 2)
    sDisplay = sDisplay & vbNewLine & "        '" & Trim(pasUsers(1, iLoop)) & "' on '" & Trim(pasUsers(2, iLoop)) & "'"
  Next iLoop

  ' Second warning row.
  sDisplay = sDisplay & vbNewLine & vbNewLine & "All user must be logged off before the System Manager can be used."
  
  ' Display the warning.
  MsgBox sDisplay, vbOKOnly, App.ProductName

End Sub

'Private Function CheckForLock(sUserName As String, sLock As String) As Boolean
'
'    Dim sSQL As String
'    Dim rsTemp As rdoResultset
'
'    sSQL = "exec sp_ASRGetLockInfo"
'    Set rsTemp = rdoCon.OpenResultset(sSQL)
'
'    If rsTemp(0) <> "No Lock" Then
'        If rsTemp(0) <> sUserName Then
'            CheckForLock = True
'        End If
'        sLock = rsTemp(0)
'    End If
'
'    rsTemp.Close
'    Set rsTemp = Nothing
'
'End Function
'
'
'Private Function CheckUsers() As Boolean
'
'  Dim rsUsers As rdoResultset
'  Dim sDisplay As String
'  Dim sSQL As String
'  Dim sDatabase As String
'  Dim sComputerName As String
'
'  Dim sSystemName As String
'  Dim sSecurityName As String
'  Dim sUserModuleName As String
'  Dim sIntranetName As String
'
'  Dim sProgName As String
'  Dim sHostName As String
'  Dim sLoginName As String
'
'  sComputerName = UI.GetHostName
'  sDatabase = txtDatabase.Text
'
'  'Now we're connected, check for number of users logged on. First check if anyone is using
'  'System Manager or Security Manager
'  sSQL = "Select * From ASRSysConfig"
'  Set rsUsers = rdoCon.OpenResultset(sSQL, rdOpenForwardOnly, rdConcurReadOnly, rdExecDirect)
'
'  If rsUsers.BOF And rsUsers.EOF Then
'    MsgBox "No configuration data setup, please contact COA Solutions.", vbExclamation
'    CheckUsers = False
'    Exit Function
'  End If
'
'  sSystemName = IIf(IsNull(rsUsers!SystemManagerAppName), "", rsUsers!SystemManagerAppName)
'  sSecurityName = IIf(IsNull(rsUsers!SecurityManagerAppName), "", rsUsers!SecurityManagerAppName)
'  sUserModuleName = IIf(IsNull(rsUsers!UserModuleAppName), "", rsUsers!UserModuleAppName)
'  sIntranetName = IIf(IsNull(rsUsers!IntranetModuleAppName), "", rsUsers!IntranetModuleAppName)
'  rsUsers.Close
'
'  ' RH 29/09/00
'  '
'  ' LOGIN CHECKS...Noticed that when retrieving the recordset via an SQL
'  ' stored procedure, it sometime 'duplicates' one entry, which could be
'  ' the reason why Lloyds doesnt work. Avoid this by using the stored
'  ' procedure code direct in VB.
'  '
'  ' In this case, if the stored procedure exists, use it, otherwise, use the
'  ' VB Code
'  '
'
'  sSQL = "select count(*) as cnt from sysobjects where name = 'sp_asrgetusersandapp'"
'  Set rsUsers = rdoCon.OpenResultset(sSQL, rdOpenForwardOnly, rdConcurReadOnly, rdExecDirect)
'
'  If rsUsers!cnt = 0 Then
'
'      sSQL = "SELECT DISTINCT hostname, loginame, program_name, hostprocess " & _
'         "FROM master..sysprocesses " & _
'         "WHERE dbid in (" & _
'                         "SELECT dbid " & _
'                         "FROM master..sysdatabases " & _
'                         "WHERE name = '" & gsDatabaseName & "') " & _
'         "and spid <> @@spid " & _
'         "ORDER BY loginame"
'
'  Else
'
'    sSQL = "exec sp_ASRGetUsersAndApp '" & sDatabase & "'"
'
'  End If
'
'
'  Set rsUsers = rdoCon.OpenResultset(sSQL, rdOpenForwardOnly, rdConcurReadOnly, rdExecDirect)
'
'
'  Do While Not rsUsers.EOF
'
'    sProgName = Trim(rsUsers!Program_name)
'    sHostName = Trim(rsUsers!HostName)
'    sLoginName = Trim(rsUsers!loginame)
'
'
'    'MH20010129 Fault 1167
'    'If its this computer using this app but there is a previous existance
'    'of this app then this is an old connection and should be ignored!)
'    If sHostName = sComputerName And _
'       sProgName = sSystemName And _
'       App.PrevInstance = False Then
'            sProgName = vbNullString
'    End If
'
'
'    If LCase(Trim(sProgName)) = LCase(Trim(sSystemName)) Or _
'       LCase(Trim(sProgName)) = LCase(Trim(sSecurityName)) Or _
'       LCase(Trim(sProgName)) = LCase(Trim(sUserModuleName)) Or _
'       LCase(Trim(sProgName)) = LCase(Trim(sIntranetName)) Then
'        'Found a OpenHR module running so report it
'        sDisplay = sDisplay & _
'                   "User '" & Trim(sLoginName) & "'" & _
'                   " logged onto Machine '" & Trim(sHostName) & "'" & _
'                   " is currently using " & Trim(sProgName) & "." & vbnewline
'
'    End If
'
'    rsUsers.MoveNext
'  Loop
'
'  rsUsers.Close
'
'  CheckUsers = (sDisplay = vbNullString)
'  If Not CheckUsers Then
'    If ASRDEVELOPMENT Then
'      CheckUsers = True
'    Else
'      Set rdoCon = Nothing
'    End If
'
'    sDisplay = "The following user(s) are currently logged into the OpenHR databases :" & vbnewline & vbnewline & _
'               sDisplay & vbnewline & vbnewline & _
'               "All user must be logged off before the " & App.ProductName & " can be used." & _
'               IIf(ASRDEVELOPMENT, vbnewline & "(COA Solutions Development bypass!)", "")
'    MsgBox sDisplay, vbOKOnly, App.ProductName
'  End If
'
'End Function


Private Sub CheckPassword()

  Dim rsInfo As New ADODB.Recordset      ' Recordset used to retrieve data
  Dim lMinimumLength As Long      ' The minimum length for passwords
  Dim lChangeFrequency As Long    ' How often passwords must be changed
  Dim sChangePeriod As String     ' How often passwords must be changed
  Dim dLastChanged As Date        ' Date user last changed password
  Dim fForceChange As Boolean     ' Must the user change password
  Dim iForceChange As PasswordChangeReason     ' 0 = Does not have to change
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
        UnLoad frmChangePassword
        gADOCon.Close
        End
      Else
        UnLoad frmChangePassword
        Set frmChangePassword = Nothing
        UpdateConfig
      End If
    End If
  End If
End Sub

Private Sub UpdateConfig()

  On Error GoTo Update_ERROR
  
  Dim rsInfo As New ADODB.Recordset

  ' Get the users specific Info From ASRSysPasswords
  'JPD 20050812 Fault 10166
  rsInfo.Open "Select * From ASRSysPasswords WHERE Username = '" & LCase(Replace(gsUserName, "'", "''")) & "'", gADOCon, adOpenDynamic, adLockOptimistic
  
  If rsInfo.BOF And rsInfo.EOF Then
    rsInfo.AddNew
  End If
  
  rsInfo!userName = LCase(gsUserName)
  rsInfo!LastChanged = Replace(Format(Now, "mm/dd/yyyy"), UI.GetSystemDateSeparator, "/")
  rsInfo!ForceChange = 0
  
  rsInfo.Update

  Set rsInfo = Nothing
  Exit Sub
  
Update_ERROR:
  
  MsgBox "Error updating AsrSysPasswords." & vbNewLine & vbNewLine & _
         "(" & Err.Number & " - " & Err.Description & ")", vbExclamation + vbOKOnly, App.Title
         
  Set rsInfo = Nothing
         
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
        gblnAutomaticLogon = True
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

      Case "trusted"
        gblnAutomaticLogon = True
        blnPassword = True
        chkUseWindowsAuthentication.value = vbChecked

      Case "script"
        gblnAutomaticScript = (LCase(strValue) = "true")
      
      Case "save"
        gblnAutomaticSave = (LCase(strValue) = "true")

      End Select

    Next

    If blnPassword Then
      gobjProgress.AVI = dbLogin
      gobjProgress.MainCaption = "Login"
      gobjProgress.NumberOfBars = 0
      gobjProgress.Caption = "Attempting log on to OpenHR..."
      gobjProgress.OpenProgress
      Login
    End If

  End If

End Sub


Private Sub CheckApplicationAccess()

  On Error GoTo ErrorTrap

  Dim rsTemp As ADODB.Recordset
  Dim blnCurrentlyLocked As Boolean
  Dim blnReadWriteLock As Boolean
  Dim strLockDetails As String
  Dim strLockUser As String
  Dim strLockType As String
  Dim strMBText As String
  Dim intMBResponse As Integer
  
  
  'Check security permissions
  ' AE20090325 Fault #13628
  'If (SystemPermission("MODULEACCESS", "SYSTEMMANAGER") = True Or LCase(gsUserName) = "sa" Or gfCurrentUserIsSysSecMgr) And Not mblnForceReadOnly Then
  If (SystemPermission("MODULEACCESS", "SYSTEMMANAGER") = True Or LCase(gsUserName) = "sa") And gbCurrentUserIsSysSecMgr And (Not mblnForceReadOnly) Then
    If IsModuleEnabled(modFullSysMgr) = True Then
      Application.AccessMode = accFull
    Else
      Application.AccessMode = accLimited
    End If
  ElseIf SystemPermission("MODULEACCESS", "SYSTEMMANAGERRO") = True Then
    Application.AccessMode = accSystemReadOnly
  Else
    Application.AccessMode = accNone
    Screen.MousePointer = vbDefault
    MsgBox "You do not have permission to run the System Manager." & vbNewLine & vbNewLine & _
           "Please contact your security administrator.", vbOKOnly + vbExclamation, App.ProductName
    Exit Sub
  End If

  Set rsTemp = New ADODB.Recordset
  rsTemp.Open "sp_ASRLockCheck", gADOCon, adOpenDynamic, adLockReadOnly, adCmdStoredProc
  blnCurrentlyLocked = (Not rsTemp.BOF And Not rsTemp.EOF)
  
  If blnCurrentlyLocked Then
    'Ignore users own manual lock
    If LCase(gsUserName) = LCase(rsTemp!userName) And rsTemp!Priority = lckManual Then
      rsTemp.MoveNext
    End If
    blnCurrentlyLocked = (Not rsTemp.BOF And Not rsTemp.EOF)

    If blnCurrentlyLocked Then
  
      'If not locked by current app then can we get read only access...
      strLockDetails = "User :  " & rsTemp!userName & vbNewLine & _
                       "Date/Time :  " & rsTemp!Lock_Time & vbNewLine & _
                       "Machine :  " & rsTemp!HostName & vbNewLine & _
                       "Type :  " & rsTemp!Description
  
      Screen.MousePointer = vbDefault
      
      'If rsTemp!Priority = lckReadWrite Or LCase(gsUserName) = "sa" Then
      If rsTemp!Priority = lckReadWrite Or _
        (rsTemp!Priority = lckManual And LCase(gsUserName) = "sa") Then
        'Ask if user wants read only if somebody else has system read/write
  
        If Application.AccessMode <> accSystemReadOnly Then
          strMBText = "The database is currently locked as follows:" & _
                      vbNewLine & vbNewLine & strLockDetails & vbNewLine & vbNewLine & _
                      "Would you like to proceed with read only access?"
          intMBResponse = MsgBox(strMBText, vbYesNo + vbExclamation, Application.Name)
          Application.AccessMode = IIf(intMBResponse = vbYes, accSystemReadOnly, accNone)
        End If
  
      Else
        'Database is totally locked (saving or manual lock)
        MsgBox "Unable to login to '" & gsDatabaseName & "' as the database has been locked." & _
               vbNewLine & vbNewLine & strLockDetails, vbExclamation, Application.Name
        Application.AccessMode = accNone
  
      End If

      Screen.MousePointer = vbHourglass

    End If
  
  End If
    
  rsTemp.Close
    
  If Application.AccessMode = accFull Then
    'Now have write access .. lock database so nobody else can...
    gADOCon.Execute "sp_ASRLockWrite 3, 1, ''"
  End If
'END TRANS

  Set rsTemp = Nothing
  Exit Sub
  
ErrorTrap:




End Sub


Private Function SystemPermission(strCategory As String, strItem As String) As Boolean

  Dim cmdSystemPermission As New ADODB.Command
  Dim pmADO As ADODB.Parameter

  With cmdSystemPermission
    .CommandText = "sp_ASRSystemPermission"
    .CommandType = adCmdStoredProc
    .CommandTimeout = 0
    Set .ActiveConnection = gADOCon

    Set pmADO = .CreateParameter("Result", adBoolean, adParamOutput)
    .Parameters.Append pmADO

    Set pmADO = .CreateParameter("Category", adVarChar, adParamInput, 50)
    .Parameters.Append pmADO
    pmADO.value = strCategory

    Set pmADO = .CreateParameter("Item", adVarChar, adParamInput, 50)
    .Parameters.Append pmADO
    pmADO.value = strItem
    
    Set pmADO = .CreateParameter("UserName", adVarChar, adParamInput, 200)
    .Parameters.Append pmADO
    pmADO.value = gsActualSQLLogin

    .Execute

    SystemPermission = IIf(IsNull(.Parameters(0).value), "", .Parameters(0).value)
  End With
  
  Set cmdSystemPermission = Nothing

End Function
Public Function CreateLoginErrFile(sMsg As String) As String
'NHRD - 17042003 - Fault 5058 Added this errorfile writing function.
'Nicked from Data manager frmLogin

  'This sub will return the last ADO error.
  CreateLoginErrFile = vbNullString

  Const lngFileNum As Integer = 99
  Dim lngCount As Long
  On Local Error Resume Next

  Open gsLogDirectory & "\LoginErr.txt" For Output As #lngFileNum
  Print #lngFileNum, "Server    : " & txtServer.Text
  Print #lngFileNum, "Database  : " & txtDatabase.Text
  Print #lngFileNum, "Username  : " & txtUID.Text
  Print #lngFileNum, "Version   : " & CStr(App.Major) & "." & CStr(App.Minor) & "." & CStr(App.Revision)
  Print #lngFileNum, ""

  Print #lngFileNum, "Date/Time : " & CStr(Now)
  If Not gADOCon Is Nothing Then
    ' JPD20030211 Fault 5044
    Print #lngFileNum, "Connection: " & IIf(gADOCon.State = 4, "Still Executing", "Finished Executing")
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
      Print #lngFileNum, "=====================" & vbNewLine
      
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
    Set pmADO = .CreateParameter("ModuleKey", adVarChar, adParamInput, 20, "SYSTEMMANAGER")
    .Parameters.Append pmADO
      
    .Execute

    GetActualLoginName = IIf(IsNull(.Parameters(0).value), "", .Parameters(0).value)
  End With
  
  Set cmdDetail = Nothing

End Function
