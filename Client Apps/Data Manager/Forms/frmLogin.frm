VERSION 5.00
Object = "{BD0C1912-66C3-49CC-8B12-7B347BF6C846}#13.1#0"; "Codejock.SkinFramework.v13.1.0.ocx"
Begin VB.Form frmLogin 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "OpenHR Data Manager - Login"
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
   HelpContextID   =   1003
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
   Begin VB.TextBox txtServer 
      Height          =   315
      Left            =   1515
      TabIndex        =   5
      Top             =   3240
      Visible         =   0   'False
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
   Begin VB.TextBox txtUID 
      Height          =   315
      Left            =   1515
      TabIndex        =   1
      Top             =   1680
      Width           =   3280
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
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   400
      Left            =   4950
      TabIndex        =   6
      Top             =   1680
      Width           =   1200
   End
   Begin VB.CommandButton cmdDetails 
      Caption         =   "&Details  >>"
      Height          =   400
      Left            =   4950
      TabIndex        =   8
      Top             =   2835
      Width           =   1200
   End
   Begin VB.CheckBox chkUseWindowsAuthentication 
      Caption         =   "&Use Windows Authentication"
      Height          =   210
      Left            =   240
      TabIndex        =   3
      Top             =   2520
      Value           =   1  'Checked
      Width           =   3240
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   65535
      Left            =   5085
      Top             =   3030
   End
   Begin XtremeSkinFramework.SkinFramework SkinFramework1 
      Left            =   5625
      Top             =   3135
      _Version        =   851969
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
   Begin VB.Label lblDevelopmentMode 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "(Dev mode)"
      Height          =   195
      Left            =   5025
      TabIndex        =   13
      Top             =   1215
      Width           =   1035
   End
   Begin VB.Label lblServer 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Server :"
      Height          =   195
      Left            =   240
      TabIndex        =   12
      Top             =   3300
      Visible         =   0   'False
      Width           =   810
   End
   Begin VB.Label lblDatabase 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Database :"
      Height          =   195
      Left            =   240
      TabIndex        =   11
      Top             =   2895
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label lblPassword 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Password :"
      Height          =   195
      Left            =   240
      TabIndex        =   10
      Top             =   2145
      Width           =   975
   End
   Begin VB.Label lblUser 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Username :"
      Height          =   195
      Left            =   240
      TabIndex        =   9
      Top             =   1740
      Width           =   1005
   End
   Begin VB.Line Line4 
      BorderColor     =   &H80000005&
      X1              =   120
      X2              =   6000
      Y1              =   1515
      Y2              =   1515
   End
   Begin VB.Image imgASR 
      Height          =   1065
      Left            =   120
      Picture         =   "frmLogin.frx":000C
      Top             =   120
      Width           =   4575
   End
   Begin VB.Label lblVersion 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "lblVersion"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   1215
      Width           =   840
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000010&
      BorderWidth     =   2
      X1              =   120
      X2              =   5980
      Y1              =   1515
      Y2              =   1515
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function SQLDataSources Lib "odbc32.dll" (ByVal henv&, ByVal fDirection%, ByVal szDSN$, ByVal cbDSNMax%, pcbDSN%, ByVal szDescription$, ByVal cbDescriptionMax%, pcbDescription%) As Integer
Private Declare Function SQLAllocEnv% Lib "odbc32.dll" (env&)
Private Declare Function SQLDrivers Lib "odbc32.dll" (ByVal henv&, ByVal fDirection%, ByVal szDriverDesc$, ByVal cbDriverDescMax%, pcbDriverDescBuff%, ByVal szDriverAttrib$, ByVal cbDriverAttribMax%, pcbDriverAttribBuff%) As Integer

Const SQL_SUCCESS As Long = 0
Const SQL_FETCH_NEXT As Long = 1

Public OK As Boolean

'Instantiate internal classes
Private UI As New UI
Private Net As New Net

Private gDatStartDate As Date
Private mlngTimeOut As Long
Private mbForceTotalRelogin As Boolean

Private Sub chkUseWindowsAuthentication_Click()

  Dim strTrustedUser As String

  strTrustedUser = gstrWindowsCurrentDomain & "\" & gstrWindowsCurrentUser

  txtUID.Text = IIf(chkUseWindowsAuthentication.Value = vbChecked, strTrustedUser, txtUID.Text)
  txtUID.Enabled = IIf(chkUseWindowsAuthentication.Value = vbChecked, False, True)
  
  txtPWD.Text = IIf(chkUseWindowsAuthentication.Value = vbChecked, "", "")
  txtPWD.Enabled = IIf(chkUseWindowsAuthentication.Value = vbChecked, False, True)

  ' Grey out controls
  txtUID.BackColor = IIf(txtUID.Enabled, vbWindowBackground, vbButtonFace)
  txtPWD.BackColor = IIf(txtPWD.Enabled, vbWindowBackground, vbButtonFace)

End Sub

Private Sub cmdDetails_Click()
  ' Display/hide the database options.
  FormDisplay (cmdDetails.Caption = "&Details  >>")
  If txtDatabase.Visible Then
    txtDatabase.SetFocus
  End If
  
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyF1
      If ShowAirHelp(Me.HelpContextID) Then
        KeyCode = 0
      End If
    Case KeyCode = 192
        KeyCode = 0
  End Select
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    
  ' Record the time the user last did anything on this screen.
  gDatStartDate = Now
  
End Sub

Private Sub Form_Load()

  On Error GoTo ErrorTrap
  
  Dim fDatabaseInfoMissing  As Boolean
  Dim sAnimationPath As String
  Dim sUsername As String
  Dim sDatabaseName As String
  Dim sServerName As String
  Dim sPassword As String
  Dim bUseWindowsAuthentication As Boolean
  Dim bBypassLoginScreen As Boolean

  ASRDEVELOPMENT = Not vbCompiled
  gblnAutomaticLogon = False
  
  ' Load the CodeJock Styles
  Call LoadSkin(Me, Me.SkinFramework1)
  
  'ShowInTaskbar
  SetWindowLong Me.hWnd, GWL_EXSTYLE, (GetWindowLong(hWnd, GWL_EXSTYLE) Or WS_EX_APPWINDOW)

  
  ' Position the login screen in the centre of the screen.
  UI.frmAtCenter Me

  'lblVersion.Caption = "OpenHR Data Manager - Version " & App.Major & "." & App.Minor & "." & App.Revision
  lblVersion.Caption = "Version " & app.Major & "." & app.Minor & "." & app.Revision
  
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
  'sAnimationPath = GetPCSetting( "DataPaths", "VideoPath") & "\asr.avi"
  'sAnimationPath = App.Path & "\Videos\asr.avi"
  'On Error GoTo ErrorNoAnimation
  'aniLogon.Open sAnimationPath
  'On Error GoTo ErrorTrap

  If Not CheckSQLDriver Then
    COAMsgBox "The required ODBC Driver '" & ODBCDRIVER & "' is not installed." & vbNewLine & _
      "Install the driver before running OpenHR.", _
      vbExclamation + vbOKOnly, app.ProductName
    Unload frmLogin
    Exit Sub
  End If
  
  ' Get the details of the last connection from the registry.
  sUsername = GetPCSetting("Login", "DataMgr_UserName", Net.userName)
  sDatabaseName = GetPCSetting("Login", "DataMgr_Database", vbNullString)
  sServerName = GetPCSetting("Login", "DataMgr_Server", vbNullString)
  bUseWindowsAuthentication = GetPCSetting("Login", "DataMgr_AuthenticationMode", 0)
  bBypassLoginScreen = GetPCSetting("Login", "DataMgr_Bypass", False)
      
  ' If the database/server information is missing from the registry then display the
  ' controls.
  fDatabaseInfoMissing = (sDatabaseName = vbNullString) Or _
    (sServerName = vbNullString)
  FormDisplay fDatabaseInfoMissing

  txtUID.Text = sUsername
  txtPWD.Text = vbNullString
  txtDatabase.Text = sDatabaseName
  txtServer.Text = sServerName
  chkUseWindowsAuthentication.Value = IIf(bUseWindowsAuthentication = True, vbChecked, vbUnchecked)
  chkUseWindowsAuthentication_Click

  CheckCommandLine

  ' Bypass login screen
  If bBypassLoginScreen And Not gbForceLogonScreen Then
    'gobjProgress.AviFile = App.Path & "\videos\about.avi"
    gobjProgress.AVI = dbLogin
    gobjProgress.MainCaption = "Login"
    gobjProgress.NumberOfBars = 0
    gobjProgress.Caption = "Attempting login..."
    gobjProgress.OpenProgress
    Login
  End If

  gDatStartDate = Now
  Timer1.Enabled = True

TidyUpAndExit:
  Exit Sub

ErrorTrap:
  Resume TidyUpAndExit
  
ErrorNoAnimation:
  'If Err.Number = 53 Then
  '  aniLogon.Visible = False
  '  imgASR.Visible = True
  '  Resume Next
  'End If
  Resume TidyUpAndExit

End Sub

Private Sub Form_Activate()

  ' Set focus on the password if the username is already entered.
  If Len(Trim(txtUID.Text)) > 0 Then
    If txtPWD.Enabled Then
      txtPWD.SetFocus
    End If
  End If
  
  mbForceTotalRelogin = False
  
End Sub

Private Sub cmdOK_Click()
  Login
  
  ' SQL 2005 may have forced us to change the password, so we need to re-attempt the logon process in full
  If mbForceTotalRelogin Then
    txtPWD.Text = gsPassword
    Login
  End If
End Sub

Private Sub cmdCancel_Click()
  
  ' Quit the system.
  OK = False
  Me.Hide
  
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

  If UnloadMode <> vbFormCode Then
    OK = False
  End If
  
End Sub

Private Sub Form_Unload(Cancel As Integer)

  Set UI = Nothing
  Set Net = Nothing
End Sub


Private Sub Timer1_Timer()

  ' If the user has done nothing for 5 minutes then
  ' quit the system.
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

Private Sub txtUID_Change()

  If Len(Me.txtUID.Text) > 50 Then
    Me.txtUID.Text = Left(Me.txtUID.Text, 50)
  End If
  
End Sub

Private Sub txtUID_GotFocus()

  ' Select all of the text in the textbox.
  UI.txtSelText
  
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

Private Sub FormDisplay(bDetails As Boolean)

    If bDetails Then
        cmdDetails.Caption = "&Details  <<"
        'Me.Height = 4110
        txtDatabase.Visible = True
        lblDatabase.Visible = True
        txtServer.Visible = True
        lblServer.Visible = True
    Else
        cmdDetails.Caption = "&Details  >>"
        'Me.Height = 3000
        txtDatabase.Visible = False
        lblDatabase.Visible = False
        txtServer.Visible = False
        lblServer.Visible = False
    End If

End Sub

Private Sub CheckRegistrySettings()
  ' Check the required registry settings are defined.
  ' If not, then get the user to initialise them.
  On Error GoTo ErrorTrap
  
  Dim fPathExists As Boolean
  
  Dim sCrystalPath As String
  Dim sDocumentsPath As String
  Dim sLocalOLEPath As String
  Dim sOLEPath As String
  Dim sPhotoPath As String
  
  Dim sBrowserPath As String
  Dim sCurrentPath As String
  Dim frmSelectPath As frmPathSel
  Dim fContinue As Boolean

  Dim bHasOleColumn As Boolean
  Dim bHasPhotoColumn As Boolean
  
  ' Remember the current directory.
  sCurrentPath = CurDir
  
  ' Retrieve all the paths
  gsPhotoPath = GetPCSetting("DataPaths", "PhotoPath_" & gsDatabaseName, vbNullString)
  gsOLEPath = GetPCSetting("DataPaths", "OLEPath_" & gsDatabaseName, vbNullString)
  gsCrystalPath = GetPCSetting("DataPaths", "crystalpath_" & gsDatabaseName, vbNullString)
  gsDocumentsPath = GetPCSetting("DataPaths", "documentspath_" & gsDatabaseName, vbNullString)
  gsLocalOLEPath = GetPCSetting("DataPaths", "localolePath_" & gsDatabaseName, vbNullString)
 
  gbPrinterPrompt = GetPCSetting("Printer", "Prompt", True)
  gbPrinterConfirm = GetPCSetting("Printer", "Confirm", False)
 
 
  bHasOleColumn = DBContains_DataType(sqlOle)
  bHasPhotoColumn = DBContains_DataType(sqlVarBinary)
 
  ' Set the continue flag to start with
  fContinue = False

  'JDM - 09/10/01 - Fault 2932 - Only allow if user has access to PC configuration
'  If datGeneral.SystemPermission("CONFIGURATION", "PC") Then
    If Not gblnBatchJobsOnly And Not ASRDEVELOPMENT Then
      'TM20011008 Fault 2261
      'Only show the message if the database has columns of these particular datatypes and
      ' a path has not yet been defined.
      If (gsPhotoPath = vbNullString And bHasPhotoColumn) _
         Or (gsOLEPath = vbNullString And bHasOleColumn) _
         Or (gsLocalOLEPath = vbNullString And bHasOleColumn) _
         Or (gsDocumentsPath = vbNullString) Then
        
        fContinue = COAMsgBox("One or more data paths have not yet been defined for this database." & vbNewLine & _
                           "OpenHR may not function correctly without these paths defined." & vbNewLine & _
                           "Would you like to define these now?", vbYesNo + vbQuestion + vbDefaultButton2, app.Title) = vbYes
      End If
      
    End If
'  End If
  
  If fContinue Then
  
' NPG20100824 Fault 1096 - allow UNC's as folder paths.
    If frmConfiguration.Initialise(False) Then
      frmConfiguration.Show vbModal
    End If
    Set frmConfiguration = Nothing
    frmMain.EnableMenu Me
  
'    ' Check that the Documents path is defined.
'    If sDocumentsPath = vbNullString Then
'      Set frmSelectPath = New frmPathSel
'      With frmSelectPath
'        .SelectionType = 8
'        .QuietMode = False
'        .Show vbModal
'      End With
'      Set frmSelectPath = Nothing
'    Else
'      ' Change directory to the OLE path to check that it exists.
'      fPathExists = True
'      ChDir sDocumentsPath
'      'If the directory doesn't exist, ask the user to select it.
'      If Not fPathExists Then
'        Set frmSelectPath = New frmPathSel
'        With frmSelectPath
'          .SelectionType = 8
'          .QuietMode = False
'          .Show vbModal
'        End With
'        Set frmSelectPath = Nothing
'      End If
'    End If
'
'    If bHasOleColumn Then
'      ' Check that the OLE path is defined.
'      If sOLEPath = vbNullString Then
'        Set frmSelectPath = New frmPathSel
'        With frmSelectPath
'          .SelectionType = 2
'          .QuietMode = False
'          .Show vbModal
'        End With
'        Set frmSelectPath = Nothing
'      Else
'        ' Change directory to the OLE path to check that it exists.
'        fPathExists = True
'        ChDir sOLEPath
'        'If the directory doesn't exist, ask the user to select it.
'        If Not fPathExists Then
'          Set frmSelectPath = New frmPathSel
'          With frmSelectPath
'            .SelectionType = 2
'            .QuietMode = False
'            .Show vbModal
'          End With
'          Set frmSelectPath = Nothing
'        End If
'      End If
'
'      ' Check that the Local OLE path is defined.
'      If sLocalOLEPath = vbNullString Then
'        Set frmSelectPath = New frmPathSel
'        With frmSelectPath
'          .SelectionType = 16
'          .QuietMode = False
'          .Show vbModal
'        End With
'        Set frmSelectPath = Nothing
'      Else
'        ' Change directory to the OLE path to check that it exists.
'        fPathExists = True
'        ChDir sLocalOLEPath
'        'If the directory doesn't exist, ask the user to select it.
'        If Not fPathExists Then
'          Set frmSelectPath = New frmPathSel
'          With frmSelectPath
'            .SelectionType = 16
'            .QuietMode = False
'            .Show vbModal
'          End With
'          Set frmSelectPath = Nothing
'        End If
'      End If
'    End If
'
'    If bHasPhotoColumn Then
'      ' Check that the Photo path is defined.
'      If sPhotoPath = vbNullString Then
'        Set frmSelectPath = New frmPathSel
'        With frmSelectPath
'          .SelectionType = 1
'          .QuietMode = False
'          .Show vbModal
'        End With
'        Set frmSelectPath = Nothing
'      Else
'        ' Change directory to the photo path to check that it exists.
'        fPathExists = True
'        ChDir sPhotoPath
'        'If the directory doesn't exist, ask the user to select it.
'        If Not fPathExists Then
'          Set frmSelectPath = New frmPathSel
'          With frmSelectPath
'            .SelectionType = 1
'            .QuietMode = False
'            .Show vbModal
'          End With
'          Set frmSelectPath = Nothing
'        End If
'      End If
'    End If
  End If
 
  ' Let windows get rid of the screen residue.
  DoEvents
  
  ' Return the original current directory.
  ChDir sCurrentPath
  
  'Load them all into global strings
  gsPhotoPath = GetPCSetting("DataPaths", "PhotoPath_" & gsDatabaseName, vbNullString)
  gsOLEPath = GetPCSetting("DataPaths", "OLEPath_" & gsDatabaseName, vbNullString)
  gsCrystalPath = GetPCSetting("DataPaths", "crystalpath_" & gsDatabaseName, vbNullString)
  gsDocumentsPath = GetPCSetting("DataPaths", "documentspath_" & gsDatabaseName, vbNullString)
  gsLocalOLEPath = GetPCSetting("DataPaths", "localolePath_" & gsDatabaseName, vbNullString)
  
  ' What start day of the week do we display
  giWeekdayStart = GetSystemSetting("General", "WeekdayStart", vbSunday)
  
  ' Load the default display for RecEdit
  gcPrimary = GetUserSetting("RecordEditing", "Primary", disFindWindow)
  gcHistory = GetUserSetting("RecordEditing", "History", disFindWindow)
  gcLookUp = GetUserSetting("RecordEditing", "LookUp", disFindWindow)
  gcQuickAccess = GetUserSetting("RecordEditing", "QuickAccess", disRecEdit_New)
  
  ' Load the default settings for DefSel
  gbCloseDefSelAfterRun = CBool(GetUserSetting("DefSel", "CloseAfterRun", False))
  gbRecentDisplayDefSel = CBool(GetUserSetting("DefSel", "RecentDisplayDefSel", False))
  gbRememberDefSelID = CBool(GetUserSetting("DefSel", "RememberLastID", True))
   
  Exit Sub
  
ErrorTrap:
  ' Catch the error if the photo, OLE, documents or local OLE directories do not exist.
  If Err = 76 Then
    fPathExists = False
    Resume Next
  End If
  
End Sub


Private Sub CheckPassword()

  On Error GoTo Check_ERROR
  
  Dim rsInfo As Recordset         ' Recordset used to retrieve data
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
  Dim objData As clsDataAccess
  Dim sSQL As String
  Dim iUsers As Integer
  
  Set objData = New clsDataAccess
  
  '' First store the config info in local variables
  'Set rsInfo = datGeneral.GetReadOnlyRecords("Select * From ASRSysConfig")
  'lMinimumLength = rsInfo!MinimumPasswordLength
  'lChangeFrequency = rsInfo!changepasswordfrequency
  'sChangePeriod = IIf(IsNull(rsInfo!changepasswordperiod), vbNullString, rsInfo!changepasswordperiod)
  lMinimumLength = GetSystemSetting("Password", "Minimum Length", 0)
  lChangeFrequency = GetSystemSetting("Password", "Change Frequency", 0)
  sChangePeriod = GetSystemSetting("Password", "Change Period", "")


  If sChangePeriod = "W" Then sChangePeriod = "WW"
  If sChangePeriod = "Y" Then sChangePeriod = "YYYY"
  
  ' Get the users specific Info From ASRSysPasswords
  Set rsInfo = datGeneral.GetReadOnlyRecords("Select * From ASRSysPasswords WHERE Username = '" & LCase(datGeneral.UserNameForSQL) & "'")
  
  If rsInfo.BOF And rsInfo.EOF Then
    ' User isnt in the table, so force a change
    'iForceChange = 4
    
    ' RH 19/09/00 - BUG 961 - If not in the table, put them in
    sSQL = "INSERT INTO AsrSysPasswords (Username, LastChanged, ForceChange) " & _
       "VALUES ('" & LCase(datGeneral.UserNameForSQL) & "','" & Replace(Format(Now, "mm/dd/yyyy"), UI.GetSystemDateSeparator, "/") & "',0)"
    objData.ExecuteSql (sSQL)
    
    'JPD 20040213 Fault 8091
    'dLastChanged = Format(Now, "dd/mm/yyyy")
    dLastChanged = Date
    
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
      
        'JPD 20040213 Fault 8091
        'If DateAdd(sChangePeriod, -lChangeFrequency, Format(Now, "dd/mm/yyyy")) >= dLastChanged Then iForceChange = 2
        If DateAdd(sChangePeriod, -lChangeFrequency, Date) >= dLastChanged Then iForceChange = 2
      End If
    End If
  
  End If
   
  ' Check for forced to change
  If iForceChange = 0 Then
    If fForceChange <> 0 Then iForceChange = 3
  End If
  
  ' If we are here and iforcechange = 0 then we dont have to change, so exit
  If iForceChange = 0 Then
    Set objData = Nothing
    Exit Sub
  End If
  
  'MH20061017 Fault 11376
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
  
  Set objData = Nothing
  Exit Sub
  
Check_ERROR:
  
  COAMsgBox "Error checking passwords." & vbNewLine & vbNewLine & "(" & Err.Number & " - " & Err.Description & ")", vbExclamation + vbOKOnly, app.Title
  
End Sub

Private Sub UpdateConfig()

  On Error GoTo Update_ERROR

  Dim rsInfo As Recordset
  Dim sSQL As String
  
  ' Get the users specific Info From ASRSysPasswords
  Set rsInfo = datGeneral.GetMainRecordset("Select * From ASRSysPasswords WHERE Username = '" & LCase(datGeneral.UserNameForSQL) & "'")

  If rsInfo.BOF And rsInfo.EOF Then
    sSQL = "INSERT INTO AsrSysPasswords (Username, LastChanged, ForceChange) " & _
           "VALUES ('" & LCase(datGeneral.UserNameForSQL) & "','" & Replace(Format(Now, "mm/dd/yyyy"), UI.GetSystemDateSeparator, "/") & "',0)"
  Else
    sSQL = "UPDATE AsrSysPasswords SET LastChanged = '" & Replace(Format(Now, "mm/dd/yyyy"), UI.GetSystemDateSeparator, "/") & "', " & _
           "ForceChange = 0 WHERE Username = '" & LCase(datGeneral.UserNameForSQL) & "'"
  End If

  gADOCon.Execute sSQL
  
  Set rsInfo = Nothing
  Exit Sub

Update_ERROR:

  COAMsgBox "Error updating AsrSysPasswords." & vbNewLine & vbNewLine & _
         "(" & Err.Number & " - " & Err.Description & ")", vbExclamation + vbOKOnly, app.Title

  Set rsInfo = Nothing
         
End Sub


Public Sub Login()

  Dim fInvalidUser As Boolean
  'Dim sConnect As String
  Dim sDSN As String
  Dim sDataSrc As String
  Dim sMsg As String
  Dim sSQL As String
  Dim rsUser As Recordset
  'Dim strBatchEmail As String
  Dim strADOError As String
  Dim strUserName As String
  Dim iCount As Integer
  Dim bLoginFound As Boolean
  Dim bHasDBAccess As Boolean
  Dim bDBLocked As Boolean
  Dim sPassword As String
  
  
  Screen.MousePointer = vbHourglass
   
  '14/08/2001 MH Fault 2447
  'This is only used in connection string, viewing users
  'and display so should be okay to leave case alone!
  'gsDatabaseName = LCase(txtDatabase.Text)
  gbUseWindowsAuthentication = chkUseWindowsAuthentication.Value
  gsDatabaseName = Replace(txtDatabase.Text, ";", "")
  gsServerName = Replace(txtServer.Text, ";", "")
  
  DebugOutput "frmLogin.Login", "ClearConnection"
  
  Set gcolSystemPermissions = New Collection

  datGeneral.ClearConnection
  
  'check if the database name has appostrophes in it!
  If InStr(1, gsDatabaseName, "'") > 0 Then
    gobjProgress.CloseProgress
    Screen.MousePointer = vbDefault
    COAMsgBox "Error logging in." & vbNewLine & vbNewLine & _
      "The database name contains an apostrophe.", _
      vbOKOnly + vbExclamation, Application.Name
    txtDatabase.SetFocus
    Exit Sub
  End If
    
  DebugOutput "frmLogin.Login", "GetSQLProviderString"
  
  ' Build the ODBC connection string.
  gsConnectionString = GetSQLProviderString
  If Len(gsConnectionString) = 0 Then
    Screen.MousePointer = vbDefault
    COAMsgBox "Microsoft SQL Native Client has not been correctly installed on this machine.", _
        vbExclamation + vbOKOnly, app.ProductName
    Exit Sub
  End If
  
  If Len(Trim(txtServer.Text)) > 0 Then
    'AE20071005 Fault #12135
    'gsConnectionString = gsConnectionString & "Server=" & txtServer.Text & ";"
    gsConnectionString = gsConnectionString & "Data Source=" & gsServerName & ";"
  Else
    gobjProgress.CloseProgress
    Screen.MousePointer = vbDefault
    COAMsgBox "Please enter the name of the server on which the OpenHR database is located.", _
        vbExclamation + vbOKOnly, app.ProductName
    txtServer.SetFocus
    Exit Sub
  End If
  
  strUserName = Replace(txtUID.Text, ";", "")
 
  If LenB(strUserName) <> 0 Then
    gsConnectionString = gsConnectionString & "User ID=" & strUserName
  Else
    
    If Not gblnBatchJobsOnly Then
      gobjProgress.CloseProgress
      Screen.MousePointer = vbDefault
      COAMsgBox "Please enter a user name.", vbExclamation + vbOKOnly, app.ProductName
      txtUID.SetFocus
      Exit Sub
    End If
  End If
  
  gsConnectionString = gsConnectionString & ";Password='" & Replace(Replace(txtPWD.Text, ";", ""), "'", "''") & "';"
  
  If LenB(gsDatabaseName) <> 0 Then
    gsConnectionString = gsConnectionString & "Initial Catalog=" & gsDatabaseName & ";"
  Else
    gobjProgress.CloseProgress
    Screen.MousePointer = vbDefault
    COAMsgBox "Please enter the name of the OpenHR database.", _
        vbExclamation + vbOKOnly, app.ProductName
    txtDatabase.SetFocus
    Exit Sub
  End If
  
  'gsConnectionString = gsConnectionString & "APP=" & App.ProductName & ";"
  gsConnectionString = gsConnectionString & "Application Name=" & app.ProductName & ";"
  
  ' Different connection string depending if use are using Windows Authentication
  If gbUseWindowsAuthentication Then
    gsConnectionString = gsConnectionString & ";Integrated Security=SSPI;"
  End If
  
  gsConnectionString = gsConnectionString & ";DataTypeCompatibility = 80;"
  
  'Set error trap
  On Error GoTo LoginError
  
  DebugOutput "frmLogin.Login", "GeneralConnect"
  
  'Establish database connection
  If Not datGeneral.Connect(gsConnectionString, sMsg, strUserName, mlngTimeOut, bDBLocked) Then
    OK = False
    GoTo LoginError
  End If
    
  ' Populate licence key
  gobjLicence.LicenceKey = GetSystemSetting("Licence", "Key", vbNullString)
  gsCustomerName = GetSystemSetting("Licence", "Customer Name", "<Unknown>")
  
  ' If using trusted connection try and find any security groups that this user is a member of
TryUsingGroupSecurity:

  DebugOutput "frmLogin.Login", "CheckVersion"
  
  ' Check the database version is the right one for the application version.
  If Not CheckVersion(Trim(txtServer.Text)) Then
    If Not gADOCon Is Nothing Then
      gADOCon.Close
      Set gADOCon = Nothing
    End If

    Screen.MousePointer = vbDefault
    Exit Sub
  End If
  
  DebugOutput "frmLogin.Login", "GetActualUserDetails"
  
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
    Set pmADO = .CreateParameter("ModuleKey", adVarChar, adParamInput, 20, "DATAMANAGER")
    .Parameters.Append pmADO

    .Execute
    
    gsSQLUserName = IIf(IsNull(.Parameters(0).Value), "", .Parameters(0).Value)
    gsUserGroup = IIf(IsNull(.Parameters(1).Value), "", .Parameters(1).Value)

  End With
  
  Set cmdDetail = Nothing

  DebugOutput "frmLogin.Login", "ModuleAccess DataManager"
  
  ' AE20090623 Fault #13674
  If datGeneral.SystemPermission("MODULEACCESS", "DATAMANAGER") = False Then
    Screen.MousePointer = vbDefault

    sMsg = "You do not have permission to run the Data Manager" & vbNewLine & vbNewLine & _
           "Contact your OpenHR security administrator"

    gobjProgress.CloseProgress
    Screen.MousePointer = vbDefault
    If ASRDEVELOPMENT Then
      COAMsgBox sMsg & vbNewLine & "(ASR Development bypass)", vbExclamation, app.ProductName
    Else
      If Not gADOCon Is Nothing Then
        gADOCon.Close
        Set gADOCon = Nothing
      End If
      COAMsgBox sMsg, vbExclamation, app.ProductName
      Exit Sub
    End If
    
  End If
  
  fInvalidUser = (Trim$(gsUserGroup) = vbNullString) _
    Or (Trim$(gsSQLUserName) = vbNullString)
  
  If fInvalidUser Then
    Screen.MousePointer = vbDefault
    If Not gADOCon Is Nothing Then
      gADOCon.Close
      Set gADOCon = Nothing
    End If
  
    gobjProgress.CloseProgress
    Screen.MousePointer = vbDefault
    
    COAMsgBox "Error logging in." & vbNewLine & vbNewLine & _
      "The user is not a member of any OpenHR user group.", _
      vbOKOnly + vbExclamation, Application.Name
    
    Exit Sub
  End If
  
  ' Is user a system or security user?
  gfCurrentUserIsSysSecMgr = CurrentUserIsSysSecMgr

  DebugOutput "frmLogin.Login", "CheckLicence"
  
  'MH20050309 Fault 9872
  ' Check the database version is the right one for the application version.
  If Not datGeneral.CheckLicence Then
    If Not gADOCon Is Nothing Then
      gADOCon.Close
      Set gADOCon = Nothing
    End If

    Screen.MousePointer = vbDefault
    Exit Sub
  End If

Bypass:
  'JPD 20040219 - replaced the gsUserName line as it reads the users
  'group AND whether or not they are SysSecMgr users.
  gsUserName = StrConv(strUserName, vbProperCase)
  gsPassword = Me.txtPWD.Text

  DebugOutput "frmLogin.Login", "SavePCSettings"
  
  If Not gblnBatchJobsOnly And Not gblnAutomaticLogon Then
    SavePCSetting "Login", "DataMgr_UserName", strUserName
    SavePCSetting "Login", "DataMgr_Database", txtDatabase.Text
    SavePCSetting "Login", "DataMgr_Server", txtServer.Text
    SavePCSetting "Login", "DataMgr_AuthenticationMode", chkUseWindowsAuthentication.Value
  End If

  'TM20011128 Fault 3224 - Switched around the checking of paths and version checking.

  'RH 24/6/99
  CheckRegistrySettings
  
  DebugOutput "frmLogin.Login", "CheckPasswordExpiry"
  
  'MH20020508 Don't check password expiry for batch user...
  If Not gblnBatchJobsOnly And Not gbUseWindowsAuthentication And glngSQLVersion < 9 Then
    CheckPassword
  End If
  
  ' AE20080303 Fault #12766 If the users default database is not 'master' then make it so.
  If Not gblnBatchJobsOnly Then
    sSQL = "IF EXISTS(SELECT 1 FROM master..syslogins WHERE loginname = SUSER_NAME() AND dbname <> 'master')" & vbNewLine & _
                               "  EXEC sp_defaultdb [" & gsUserName & "], master"
    gADOCon.Execute sSQL, , adCmdText
  End If
  
  LC_ResetLock
  LC_SaveSettingsToRegistry

  DebugOutput "frmLogin.Login", "CloseProgress"
  
  OK = True
  Me.Hide
  gobjProgress.CloseProgress
  Exit Sub

LoginError:
  Dim iErrCount As Integer
  Dim iForceChangeReason As PasswordChangeReason
  
  DebugOutput "frmLogin.Login", "LoginError"
  
  iForceChangeReason = giPasswordChange_None
  
  For iErrCount = 0 To gADOCon.Errors.Count - 1
     
    DebugOutput "frmLogin.Login", "LoginError " & gADOCon.Errors(iErrCount).NativeError
    
    Select Case gADOCon.Errors(iErrCount).NativeError
       
      ' 14 - Invalid Connection String
      ' 17 -No such server
      ' 4060 - No such database
      ' 18456 - Inavlid username/password
      Case 14, 17, 4060, 18456
        sMsg = "The system could not log you on. Make sure your details are correct, then retype your password."
      
      ' .NET error, SQL process login details are incorrect
        Case 6522
          sMsg = "The SQL process account has not been defined or is invalid." & vbNewLine & _
            "Please contact your system administrator."
            
      ' Framework or Assembly error
      Case 10314
        sMsg = "Unable to login to the Data Manager." & vbCrLf & _
          "Please ask the System Administrator to update the database in the System Manager."
    
      ' ?
      Case 18468
        sMsg = gADOCon.Errors(iErrCount).Description
    
      ' Trusted connecttion problems
      Case 18452
        If gbUseWindowsAuthentication Then
          sMsg = "The system could not log you on. Make sure your details are correct, then retype your password."
        Else
          sMsg = "Your server is configured for Windows Only security." & vbNewLine & "Please see your system administrator."
        End If
        
      ' Generic login error
      
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
        fInvalidUser = True
        GoTo TryUsingGroupSecurity
      
      Case Else
        sMsg = gADOCon.Errors(iErrCount).Description
    
    End Select
      
  Next iErrCount
  
  gobjProgress.CloseProgress
  Screen.MousePointer = vbDefault
   
  If iForceChangeReason <> 0 Then
    
    DebugOutput "frmLogin.Login", "LoginError ChangePassword"
    
    ' Change the password
    gsUserName = txtUID.Text
    gsPassword = txtPWD.Text
    If frmChangePassword.Initialise(iForceChangeReason, 0) Then
      frmChangePassword.Show vbModal
      
      ' JDM - Fault 11710 - Close application if they elect not to change password.
      If frmChangePassword.Cancelled = True Then
        End
      End If
      
      sMsg = vbNullString
      mbForceTotalRelogin = Not frmChangePassword.Cancelled
      Unload frmChangePassword
    End If
  
  Else

    DebugOutput "frmLogin.Login", "LoginError Not ChangePassword"
    
    ' Report the error
    CreateLoginErrFile sMsg
  
    If ASRDEVELOPMENT And gADOCon.State <> adStateClosed Then
      If LenB(sMsg) <> 0 Then
        COAMsgBox sMsg & vbNewLine & "(ASR Development bypass!)", vbExclamation, Me.Caption
      End If
      gsSQLUserName = txtUID.Text
      GoTo Bypass
    Else
      On Local Error Resume Next
      If gADOCon.State <> adStateClosed Then
        gADOCon.Close
      End If
      Set gADOCon = Nothing
    
      If LenB(sMsg) <> 0 Then
        
        If gblnBatchJobsOnly Then
            'frmEmailSel.SendEmail _
            '  strBatchEmail, "OpenHR Batch Logon Failure", sMsg, False
            'Unload frmEmailSel
            'Set frmEmailSel = Nothing
            SendBatchLogonFailure sMsg
          'End If
        Else
          COAMsgBox sMsg, vbExclamation, Me.Caption
        End If
      End If
    
      Err = False
      Screen.MousePointer = vbDefault
      txtPWD.SetFocus
      
      'Also need to call this in case the control already has focus !
      Call txtPWD_GotFocus
    End If
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
  mlngTimeOut = 600

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

      'If Left(strOption, 1) = "/" And Len(strOption) > 1 Then
      '  strOption = Mid(strOption, 2)
      'End If

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
        FormDisplay (LCase(strValue) = "true")
      
      Case "batch"

        gblnBatchJobsOnly = (LCase(strValue) = "true")
        If gblnBatchJobsOnly Then

'      Open "c:\mike.txt" For Output As #1
'      Print #1, "Before"
'      Close #1

        If GetPCSetting("BatchLogon", "Enabled", False) = True Then
          GetBatchLogon strUserName, strPassword, strDatabaseName, strServerName
          If strUserName = vbNullString And GetPCSetting("BatchLogon", "TrustedConnection", False) = True Then
            strUserName = gstrWindowsCurrentDomain & "\" & gstrWindowsCurrentUser
          End If
          txtUID.Text = strUserName
          txtPWD.Text = strPassword
          chkUseWindowsAuthentication.Value = IIf(GetPCSetting("BatchLogon", "TrustedConnection", False), vbChecked, vbUnchecked)
          txtDatabase.Text = strDatabaseName
          txtServer.Text = strServerName
          Exit For
        End If

'      Open "c:\mike.txt" For Output As #1
'      Print #1, "After"
'      Close #1

        End If
      
      Case "timeout"
        mlngTimeOut = CLng(strValue)

      Case "trusted"
        gblnAutomaticLogon = True
        blnPassword = True
        chkUseWindowsAuthentication.Value = vbChecked

      'Case "debug"
      '  If LCase(strValue) <> "false" Then
      '    gstrDebugOutputFile = IIf(LCase(strValue) <> "true", strValue, "debug.txt")
      '    If Dir(gstrDebugOutputFile) Then
      '      Kill gstrDebugOutputFile
      '    End If
      '  End If

      End Select

    Next

    If blnPassword Or gblnBatchJobsOnly Then
      'gobjProgress.AviFile = App.Path & "\videos\about.avi"
      gobjProgress.AVI = dbLogin
      gobjProgress.MainCaption = "Login"
      gobjProgress.NumberOfBars = 0
      gobjProgress.Caption = "Attempting login..."
      gobjProgress.OpenProgress
      Login
    End If

  End If

End Sub


Public Function CreateLoginErrFile(sMsg As String) As String

  'This sub will return the last ADO error.
  CreateLoginErrFile = vbNullString

  Const lngFileNum As Integer = 99
  Dim lngCount As Long
  On Local Error Resume Next

  Open app.Path & "\LoginErr.txt" For Output As #lngFileNum
  Print #lngFileNum, "Server    : " & txtServer.Text
  Print #lngFileNum, "Database  : " & txtDatabase.Text
  Print #lngFileNum, "Username  : " & txtUID.Text
  Print #lngFileNum, "Version   : " & CStr(app.Major) & "." & CStr(app.Minor) & "." & CStr(app.Revision)
  Print #lngFileNum, ""

  Print #lngFileNum, "Date/Time : " & CStr(Now)
  If Not gADOCon Is Nothing Then
    ' JPD20030211 Fault 5044
    Print #lngFileNum, "Connection: " & IIf(gADOCon.State <> adStateClosed, "Successful", "Failed")
  Else
    Print #lngFileNum, "Connection: Failed"
  End If
  Print #lngFileNum, ""

  Print #lngFileNum, sMsg
  Print #lngFileNum, Err.Description
  Print #lngFileNum, ""

  If Not gADOCon Is Nothing Then
    If Not gADOCon.Errors Is Nothing Then
      For lngCount = 0 To gADOCon.Errors.Count - 1
        CreateLoginErrFile = gADOCon.Errors(lngCount).Description
        Print #lngFileNum, CreateLoginErrFile
      Next
    End If
  End If
  
  Close #lngFileNum

  CreateLoginErrFile = Mid(CreateLoginErrFile, InStrRev(CreateLoginErrFile, "]") + 1)

End Function

Private Sub txtUID_KeyPress(KeyAscii As Integer)

  If Len(Me.txtUID.Text) >= 50 Then
    KeyAscii = 0
  End If
  
End Sub


Private Sub SendBatchLogonFailure(strMsgText As String)

  Dim objOutputEmail As clsOutputEMail
  Dim strTo As String

  strTo = GetPCSetting("BatchLogon", "Email", vbNullString)
  If strTo <> vbNullString Then
    Set objOutputEmail = New clsOutputEMail
    objOutputEmail.SendEmailFromClient strTo, "", "", "OpenHR Batch Logon Failure", strMsgText, "", False
    Set objOutputEmail = Nothing
  End If

End Sub
