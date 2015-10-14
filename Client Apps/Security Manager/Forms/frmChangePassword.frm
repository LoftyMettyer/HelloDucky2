VERSION 5.00
Begin VB.Form frmChangePassword 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Change Password"
   ClientHeight    =   3480
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5130
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   8012
   Icon            =   "frmChangePassword.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3480
   ScaleWidth      =   5130
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   400
      Left            =   3810
      TabIndex        =   11
      Top             =   2925
      Width           =   1200
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   400
      Left            =   2460
      TabIndex        =   10
      Top             =   2925
      Width           =   1200
   End
   Begin VB.Frame fraPassword 
      Height          =   2745
      Left            =   90
      TabIndex        =   3
      Top             =   45
      Width           =   4935
      Begin VB.TextBox txtUserName 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   2280
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   1005
         Width           =   2500
      End
      Begin VB.TextBox txtNewPWRetype 
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   2280
         PasswordChar    =   "*"
         TabIndex        =   2
         Top             =   2220
         Width           =   2500
      End
      Begin VB.TextBox txtNewPW 
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   2280
         PasswordChar    =   "*"
         TabIndex        =   1
         Top             =   1815
         Width           =   2500
      End
      Begin VB.TextBox txtOldPW 
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   2280
         PasswordChar    =   "*"
         TabIndex        =   0
         Top             =   1410
         Width           =   2500
      End
      Begin VB.Label lblReason 
         BackStyle       =   0  'Transparent
         Caption         =   "<<< Password change reason goes here ! >>>"
         Height          =   660
         Left            =   150
         TabIndex        =   9
         Top             =   255
         Width           =   4500
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "User Name :"
         Height          =   195
         Index           =   0
         Left            =   150
         TabIndex        =   8
         Top             =   1065
         Width           =   885
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Re-type New Password :"
         Height          =   195
         Index           =   3
         Left            =   150
         TabIndex        =   7
         Top             =   2280
         Width           =   2100
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "New Password :"
         Height          =   195
         Index           =   2
         Left            =   150
         TabIndex        =   6
         Top             =   1875
         Width           =   1155
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Old Password :"
         Height          =   195
         Index           =   1
         Left            =   150
         TabIndex        =   5
         Top             =   1470
         Width           =   1080
      End
   End
End
Attribute VB_Name = "frmChangePassword"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private fOK As Boolean
Private mstrErrorMessage As String
Private mblnForceChange  As Boolean
Private mlngMinimumLength As Long
Private mblnExiting As Boolean
Private mbCancelled As Boolean

Public Property Get Exiting() As Boolean

  Exiting = mblnExiting
  
End Property

Public Property Get ForceChange() As Boolean

  ForceChange = mblnForceChange
  
End Property

Public Property Get Cancelled() As Boolean
  Cancelled = mbCancelled
End Property

Public Function Initialise(iForceReason As PasswordChangeReason, lMinimumLength As Long) As Boolean

  On Error GoTo Init_ERROR

  txtUserName.Text = gsUserName
  mlngMinimumLength = lMinimumLength

  If iForceReason > 0 Then
    ForceChange = True
    Reason = iForceReason
    Screen.MousePointer = vbNormal
  Else
    ForceChange = False
  End If
 
  Initialise = True
  Exit Function
  
Init_ERROR:

  MsgBox "Error initialising the Change Password form." & vbNewLine & vbNewLine & _
         "(" & Err.Number & " - " & Err.Description & ")", vbExclamation + vbOKOnly, App.Title
  
  Initialise = False

End Function

Public Property Let ForceChange(ByVal bNewValue As Boolean)

  mblnForceChange = bNewValue
  
  If mblnForceChange Then
  
    ' Show the extra guff
  
  Else
  
    ' Hide the extra guff
    Me.lblReason.Visible = False
    Me.txtUserName.Top = 300
    Me.Label1(0).Top = 360
    Me.txtOldPW.Top = 705
    Me.Label1(1).Top = 765
    Me.txtNewPW.Top = 1110
    Me.Label1(2).Top = 1170
    Me.txtNewPWRetype.Top = 1515
    Me.Label1(3).Top = 1575
    Me.cmdCancel.Top = 2235
    Me.cmdOK.Top = 2235
    Me.fraPassword.Height = 2030
    Me.Height = 3150
  End If
  
End Property

Public Property Let Reason(ByVal iReasonCode As PasswordChangeReason)

  Select Case iReasonCode
    Case giPasswordChange_MinLength
      lblReason.Caption = "Your existing password is less than the minimum password length set by your system administrator." & vbNewLine & _
            "Your new password must be a minimum of " & mlngMinimumLength & " characters long."
    
    Case giPasswordChange_Expired
      lblReason.Caption = "Your existing password has expired." & vbNewLine & "You must change your password."
    
    Case giPasswordChange_AdminRequested
      lblReason.Caption = "Your system administrator has requested you change your password."
    
    Case giPasswordChange_LastChangeUnknown
      lblReason.Caption = "The system cannot determine when you last changed your password." & vbNewLine & "You must change your password."
    
    Case giPasswordChange_ComplexitySettings
      lblReason.Caption = "Your existing password does not meet your domain complexity requirements." & vbNewLine & "You must change your password."
    
    Case Else
      lblReason.Caption = "You must change your password."
  
  End Select
  
End Property

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
  Case vbKeyF1
    If ShowAirHelp(Me.HelpContextID) Then
      KeyCode = 0
    End If
End Select
End Sub

Private Sub Form_Load()
  txtUserName = gsUserName
  mbCancelled = True
End Sub

Private Sub cmdOK_Click()
  Dim iUsers As Integer
  
  'JPD 20020218 Fault 3343
  'If glngSQLVersion > 0 Then
  'AE20071130 Fault #12660
  If glngSQLVersion > 0 And gADOCon.State = adStateOpen Then
    'iUsers = UserSessions(gsUserName)
    iUsers = GetCurrentUsersCountOnServer(gsUserName)
  End If
  
  If iUsers < 2 Or glngSQLVersion = 0 Then
    If PasswordChange = True Then
      'Unload Me
      Me.Hide
    End If
  Else
    MsgBox "Cannot change password. This account is currently being used " & _
            "by " & IIf(iUsers > 2, iUsers & " users", "another user") & " in the system.", vbExclamation + vbOKOnly, App.Title
    If mblnForceChange Then
      mblnExiting = True
    End If
    Me.Hide
  End If

  mbCancelled = False

End Sub

Private Sub cmdCancel_Click()
  
  If mblnForceChange Then
    
    If MsgBox("You must change your password before you can log in to OpenHR." & vbNewLine & _
              "Are you sure you wish to exit ?", vbYesNo + vbQuestion, App.Title) = vbYes Then
      mblnExiting = True
      Me.Hide
    Else
      Exit Sub
    End If
    
  Else
  
    Me.Hide
    
  End If
  
  mbCancelled = True
  
End Sub


Private Function PasswordChange() As Boolean

  Dim strSQL As String
  Dim intResponse As Integer
  Dim strOldPassword As String
  Dim strNewPassword As String
  Dim iCount As Integer
  Dim strNewConnection As String
  
  On Error GoTo ErrorTrap
  
  fOK = True

   ' old password entered wrongly
  If txtOldPW.Text <> gsPassword Then
    mstrErrorMessage = "Old password has not been entered correctly."
    fOK = False
  End If
  
  ' new password the same as old !
  If fOK = True Then
    If LCase(txtOldPW.Text) = LCase(txtNewPW.Text) Then
      mstrErrorMessage = "New password cannot be the same as your old password."
      fOK = False
    End If
  End If
  
  'new does not match retyped
  If fOK = True Then
    If txtNewPW.Text <> txtNewPWRetype.Text Then
      mstrErrorMessage = "New password retyped incorrectly."
      fOK = False
    End If
  End If

  strOldPassword = IIf(Len(txtOldPW.Text) = 0, "null", "'" & Replace(txtOldPW.Text, "'", "''") & "'")
  strNewPassword = IIf(Len(txtNewPW.Text) = 0, "null", "'" & Replace(txtNewPW.Text, "'", "''") & "'")

  ' If we don't know the SQL version then the server is forcing us to change password before we can login.
  ' You won't believe what a right royal pain in the ass this is!
  If fOK = True And (glngSQLVersion = 0 Or gADOCon.State = adStateClosed) Then

     ' AE20090601 Fault #13685
    ' No open ADO connection. We need to establish a slightly modified ado connection provider...
'    gADOCon.ConnectionString = "Provider=SQLNCLI;DataTypeCompatibility=80;Server=" & gsServerName _
'      & ";UID=" & gsUserName & ";Database=" & gsDatabaseName & ";APP=OpenHR Security Manager" _
'      & ";Old Password=" & strOldPassword & ";Password=" & strNewPassword

    Dim sConn As String
    Select Case GetSQLNCLIVersion
    Case 9 ' SQL Native Client 2005
        ' AE20090624 Fault #13689
'      sConn = "Provider=SQLNCLI;Persist Security Info=True;DataTypeCompatibility=80;APP=OpenHR Security Manager;" & _
'              "User ID=" & gsUserName & ";" & _
'              "Initial Catalog=" & gsDatabaseName & ";" & _
'              "Data Source=" & gsServerName & ";" & _
'              "Old Password=" & strOldPassword & ";" & _
'              "Password=" & strNewPassword & ";"

      sConn = "Provider=SQLNCLI;Persist Security Info=False;DataTypeCompatibility=80;" & _
              "Application Name=OpenHR Security Manager;" & _
              "User ID=" & gsUserName & ";" & _
              "Initial Catalog='';" & _
              "Data Source=" & gsServerName & ";" & _
              "Old Password=" & strOldPassword & ";" & _
              "Password=" & strNewPassword & ";"
              
    Case 10 ' SQL Native Client 2008
    
      ' AE20090624 Fault #13689
'      sConn = "Provider=SQLNCLI10;Persist Security Info=True;DataTypeCompatibility=80;APP=OpenHR Security Manager;" & _
'              "User ID=" & gsUserName & ";" & _
'              "Initial Catalog=" & gsDatabaseName & ";" & _
'              "Data Source=" & gsServerName & ";" & _
'              "Old Password=" & strOldPassword & ";" & _
'              "Password=" & strNewPassword & ";"

      sConn = "Provider=SQLNCLI10;Persist Security Info=False;DataTypeCompatibility=80;" & _
              "Application Name=OpenHR Security Manager;" & _
              "User ID=" & gsUserName & ";" & _
              "Initial Catalog='';" & _
              "Data Source=" & gsServerName & ";" & _
              "Old Password=" & strOldPassword & ";" & _
              "Password=" & strNewPassword & ";"
                  Case 10 ' SQL Native Client 2008
    
    Case 11 ' SQL Native CLient 2012


      sConn = "Provider=SQLNCLI11;Persist Security Info=False;DataTypeCompatibility=80;" & _
              "Application Name=OpenHR Security Manager;" & _
              "User ID=" & gsUserName & ";" & _
              "Initial Catalog='';" & _
              "Data Source=" & gsServerName & ";" & _
              "Old Password=" & strOldPassword & ";" & _
              "Password=" & strNewPassword & ";"
              
    End Select
    
    gADOCon.ConnectionString = sConn
    gADOCon.Open
    gADOCon.Close
    
  ElseIf glngSQLVersion < 9 Then
    If fOK = True Then
      If (Len(txtNewPW.Text) < mlngMinimumLength) And mlngMinimumLength <> 0 Then
        mstrErrorMessage = "New password must be a minimum of " & mlngMinimumLength & " characters long."
        fOK = False
      End If
    End If
    
    'Attempt to change the password
    'but the old password could still be wrong! (refer localerr:)
    If fOK = True Then
      strSQL = "sp_password " & strOldPassword & ", " & strNewPassword
      gADOCon.Execute strSQL, , adExecuteNoRecords
    End If
  
  Else
    If fOK Then
      gADOCon.Execute "ALTER LOGIN [" & gsUserName & "] WITH PASSWORD = " & strNewPassword & " OLD_PASSWORD=" & strOldPassword, , adExecuteNoRecords
    End If
  
  End If

  If fOK = True Then
    Call Relogin
  End If
  
  If fOK = True Then
    Call UpdateConfig
  End If
  
TidyAndExit:

  If fOK = True Then
    MsgBox "Password successfully changed", vbInformation
  
  Else
    'Check if there is an error then reset all of the text boxes
    MsgBox mstrErrorMessage, vbExclamation
    txtOldPW.Text = vbNullString
    txtNewPW.Text = vbNullString
    txtNewPWRetype.Text = vbNullString
    txtOldPW.SetFocus
    mstrErrorMessage = vbNullString
  
  End If

  PasswordChange = fOK
  
  On Error GoTo 0

Exit Function


ErrorTrap:

  If gADOCon.Errors.Count > 0 Then
    For iCount = 0 To gADOCon.Errors.Count - 1
  
      Select Case gADOCon.Errors(iCount).NativeError
        
        ' Password does not meet domain policy
        Case 15114, 15115, 15116, 15117, 15118
          mstrErrorMessage = gADOCon.Errors(iCount).Description
        
        ' Been used in the past
        Case 18463
          mstrErrorMessage = gADOCon.Errors(iCount).Description
        
        ' Someone else has changed the password
        Case 15151
          mstrErrorMessage = "Old password incorrect."
        
        Case Else
          mstrErrorMessage = gADOCon.Errors(iCount).Description
        
      End Select
  
    Next iCount
  
  End If
  
  fOK = False
  Resume TidyAndExit

End Function


Private Sub Relogin()

  Dim sConnect As String
  'Dim sDatabaseName As String
  Dim sServerName As String
  
  On Local Error GoTo LocalErr

  Screen.MousePointer = vbHourglass
    
  fOK = True
  'sDatabaseName = GetPCSetting("Login", "DataMgr_Database", vbNullString)
  If Trim$(gsDatabaseName) = vbNullString Then
    mstrErrorMessage = "Error during re-login <database name not found>"
    fOK = False
  End If
  
  If Trim$(gsServerName) = vbNullString Then
    mstrErrorMessage = "Error during re-login <server name not found>"
    fOK = False
  End If
  
  sConnect = "Driver=SQL Server;" & _
             "Server=" & gsServerName & ";" & _
             "UID=" & gsUserName & ";" & _
             "PWD=" & txtNewPW.Text & ";" & _
             "Database=" & gsDatabaseName & ";"

  'Re-establish database connection
  Set gADOCon = New ADODB.Connection
  With gADOCon
    .ConnectionString = sConnect
    .Provider = "SQLOLEDB"
    .CommandTimeout = 0
    .ConnectionTimeout = 0
    .CursorLocation = adUseServer
    .Mode = adModeReadWrite
    .Properties("Packet Size") = 32767
    .Open
  End With
  
  If fOK Then gsPassword = txtNewPW.Text
  
LocalErr_Handler:
  Screen.MousePointer = vbDefault
  
  Exit Sub

LocalErr:
  mstrErrorMessage = "Error during re-login (" & Err.Description & ")" & vbNewLine & vbNewLine & _
                     "Please quit HR-Pro and attempt to relogin." & vbNewLine & _
                     "NOTE: Your password may have been changed."
  fOK = False

  Resume LocalErr_Handler
End Sub

Private Sub UpdateConfig()

  On Error GoTo Update_ERROR

  Dim rsInfo As New ADODB.Recordset
  Dim sSQL As String
  
  ' Get the users specific Info From ASRSysPasswords
  rsInfo.Open "Select * From ASRSysPasswords WHERE Username = '" & LCase(gsUserName) & "'", gADOCon, adOpenForwardOnly, adLockOptimistic

  If rsInfo.BOF And rsInfo.EOF Then
    sSQL = "INSERT INTO AsrSysPasswords (Username, LastChanged, ForceChange) " & _
           "VALUES ('" & LCase(gsUserName) & "','" & Replace(Format(Now, "mm/dd/yyyy"), UI.GetSystemDateSeparator, "/") & "',0)"
  Else
    sSQL = "UPDATE AsrSysPasswords SET LastChanged = '" & Replace(Format(Now, "mm/dd/yyyy"), UI.GetSystemDateSeparator, "/") & "', " & _
           "ForceChange = 0 WHERE Username = '" & LCase(gsUserName) & "'"
  End If

  gADOCon.Execute sSQL
  
  Set rsInfo = Nothing
  Exit Sub

Update_ERROR:

  MsgBox "Error updating AsrSysPasswords." & vbNewLine & vbNewLine & _
         "(" & Err.Number & " - " & Err.Description & ")", vbExclamation + vbOKOnly, App.Title
  fOK = False
  Set rsInfo = Nothing
         
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

  If Not mblnForceChange Then Exit Sub

  If UnloadMode = vbFormControlMenu Then
    If MsgBox("You must change your password before you can log in to OpenHR." & vbNewLine & _
              "Are you sure you wish to exit ?", vbYesNo + vbQuestion, App.Title) = vbYes Then
      mblnExiting = True
      Me.Hide
    End If
    Cancel = True
'  Else
'    If mblnForceChange Then Screen.MousePointer = vbHourglass
  End If

End Sub


Private Sub Form_Resize()
  'JPD 20030908 Fault 5756
  DisplayApplication

End Sub

Private Sub txtNewPW_GotFocus()

  UI.txtSelText

End Sub

Private Sub txtNewPWRetype_GotFocus()

  UI.txtSelText

End Sub

Private Sub txtOldPW_GotFocus()
  
  UI.txtSelText

End Sub




