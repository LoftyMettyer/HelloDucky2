Attribute VB_Name = "modSettings"
Option Explicit

Public Const VARCHAR_MAX_Size = 2147483646 'Yup one below the actual max, needs to be otherwise things go so awfully wrong, you don't believe me, well go on then, change it, see if I care!!!)

' Generic API for doing "interesting" stuff (mainly icon handling)
Public Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" ( _
      ByVal hWnd As Long, ByVal wMsg As Long, _
      ByVal wParam As Long, ByVal lParam As Long _
   ) As Long

Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" _
    (ByVal hWnd As Long, ByVal nIndex As Long) As Long

Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" _
    (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long


Private Declare Function GetSystemMetrics Lib "user32" ( _
      ByVal nIndex As Long _
   ) As Long

Private Const SM_CXICON = 11
Private Const SM_CYICON = 12

Private Const SM_CXSMICON = 49
Private Const SM_CYSMICON = 50
   
Private Declare Function LoadImageAsString Lib "user32" Alias "LoadImageA" ( _
      ByVal hInst As Long, _
      ByVal lpsz As String, _
      ByVal uType As Long, _
      ByVal cxDesired As Long, _
      ByVal cyDesired As Long, _
      ByVal fuLoad As Long _
   ) As Long
   
Private Const LR_DEFAULTCOLOR = &H0
Private Const LR_MONOCHROME = &H1
Private Const LR_COLOR = &H2
Private Const LR_COPYRETURNORG = &H4
Private Const LR_COPYDELETEORG = &H8
Private Const LR_LOADFROMFILE = &H10
Private Const LR_LOADTRANSPARENT = &H20
Private Const LR_DEFAULTSIZE = &H40
Private Const LR_VGACOLOR = &H80
Private Const LR_LOADMAP3DCOLORS = &H1000
Private Const LR_CREATEDIBSECTION = &H2000
Private Const LR_COPYFROMRESOURCE = &H4000
Private Const LR_SHARED = &H8000&

Private Const IMAGE_ICON = 1

Private Const WM_SETICON = &H80

Private Const ICON_SMALL = 0
Private Const ICON_BIG = 1

Private Declare Function GetWindow Lib "user32" ( _
   ByVal hWnd As Long, ByVal wCmd As Long) As Long
Private Const GW_OWNER = 4



Private Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, ByRef phkResult As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long

Private Const HKEY_LOCAL_MACHINE = &H80000002
Private Const KEY_READ = &H20019

Public Const WS_THICKFRAME As Long = &H40000
Public Const WS_MAXIMIZE As Long = &H1000000
Public Const WS_MAXIMIZEBOX As Long = &H10000
Public Const WS_MINIMIZE As Long = &H20000000
Public Const WS_MINIMIZEBOX As Long = &H20000
Public Const WS_EX_WINDOWEDGE As Long = &H100
Public Const WS_EX_APPWINDOW As Long = &H40000
Public Const WS_EX_DLGMODALFRAME As Long = &H1
Public Const GWL_EXSTYLE As Long = (-20)
Public Const GWL_STYLE As Long = (-16)

'Background type constants
Public Enum BackgroundLocationTypes
  giLOCATION_TOPLEFT = 0
  giLOCATION_TOPRIGHT = 1
  giLOCATION_CENTRE = 2
  giLOCATION_LEFTTILE = 3
  giLOCATION_RIGHTTILE = 4
  giLOCATION_TOPTILE = 5
  giLOCATION_BOTTOMTILE = 6
  giLOCATION_TILE = 7
End Enum

Public glngDesktopBitmapID As Long
Public glngDesktopBitmapLocation As BackgroundLocationTypes
Public glngDeskTopColour As Long

Public gbEnableUDFFunctions As Boolean
Public gbDeleteOrphanWindowsLogins As Boolean
Public gbDeleteOrphanUsers As Boolean

Public gbLoginMaintAutoAdd As Boolean
Public gstrLoginMaintAutoAddGroup As String
Public gbLoginMaintDisableOnLeave As Boolean
Public gbLoginMaintSendEmail As Boolean

'Security Settings - Login Checks i.e.Bad Attempts
Private gblnCFG_PCL As Boolean
Private gintCFG_BA As Integer
Private glngCFG_RT As Integer
Private glngCFG_LD As Long
Private gintPC_BA As Integer
Private gdtPC_LA As Date
Private gdtPC_LKD As Date

Public gblnDomainPCLockout As Boolean
Public gintDomainAttempts As Integer
Public glngDomainResetTime As Long
Public glngDomainLockoutDuration As Long
Public glngDomainMinimumLength As Long      ' The minimum length for passwords
Public glngDomainMinPasswordAge As Long
Public glngDomainChangeFrequency As Long    ' How often passwords must be changed
Public gstrDomainChangePeriod As String     ' How often passwords must be changed
Public giDomainComplexity As Integer
Public giDomainPasswordsRemembered As Integer

Public Enum PasswordChangeReason
  giPasswordChange_None = 0
  giPasswordChange_MinLength = 1
  giPasswordChange_Expired = 2
  giPasswordChange_AdminRequested = 3
  giPasswordChange_LastChangeUnknown = 4
  giPasswordChange_ComplexitySettings = 5
End Enum

Public Enum LicenceType
  Concurrency = 0
  P14Headcount = 1
  Headcount = 2
  DMIConcurrencyAndP14 = 3
  DMIConcurrencyAndHeadcount = 4
End Enum

Public Enum WarningType
  Headcount95Percent = 0
  Licence5DayExpiry = 1
End Enum

Public Function LoadDomainSecurityPolicy() As Boolean

  On Error GoTo ErrorTrap
  
  Dim cmdPolicy As New ADODB.Command
  Dim pmADO As ADODB.Parameter
  Dim iResult As Integer
  Dim piAssemblyRetryCount As Integer
  
  ' Get the policy settings from the server
  If glngSQLVersion > 8 Then
    With cmdPolicy
      .CommandText = "dbo.[spASRGetDomainPolicy]"
      .CommandType = adCmdStoredProc
      .CommandTimeout = 0
      Set .ActiveConnection = gADOCon
    
      Set pmADO = .CreateParameter("LockoutDuration", adInteger, adParamOutput)
      .Parameters.Append pmADO
    
      Set pmADO = .CreateParameter("lockoutThreshold", adInteger, adParamOutput)
      .Parameters.Append pmADO
    
      Set pmADO = .CreateParameter("lockoutObservationWindow", adInteger, adParamOutput)
      .Parameters.Append pmADO
    
      Set pmADO = .CreateParameter("maxPwdAge", adInteger, adParamOutput)
      .Parameters.Append pmADO
    
      Set pmADO = .CreateParameter("minPwdAge", adInteger, adParamOutput)
      .Parameters.Append pmADO
    
      Set pmADO = .CreateParameter("minPwdLength", adInteger, adParamOutput)
      .Parameters.Append pmADO
    
      Set pmADO = .CreateParameter("pwdHistoryLength", adInteger, adParamOutput)
      .Parameters.Append pmADO
    
      Set pmADO = .CreateParameter("pwdProperties", adInteger, adParamOutput)
      .Parameters.Append pmADO
    
ExecuteSP:
      piAssemblyRetryCount = piAssemblyRetryCount + 1
      .Execute
    
      glngDomainMinimumLength = IIf(IsNull(.Parameters(5).Value), 0, .Parameters(5).Value)
      glngDomainChangeFrequency = IIf(IsNull(.Parameters(3).Value), 0, .Parameters(3).Value)
      gstrDomainChangePeriod = "D"
      giDomainComplexity = IIf(IsNull(.Parameters(7).Value), 0, .Parameters(7).Value)
      glngDomainMinPasswordAge = IIf(IsNull(.Parameters(4).Value), 0, .Parameters(4).Value)
      gintDomainAttempts = IIf(IsNull(.Parameters(1).Value), 0, .Parameters(1).Value)
      gblnDomainPCLockout = (gintDomainAttempts > 0)
      glngDomainResetTime = IIf(IsNull(.Parameters(0).Value), 0, .Parameters(0).Value)
      glngDomainLockoutDuration = IIf(IsNull(.Parameters(2).Value), 0, .Parameters(2).Value)
      giDomainPasswordsRemembered = IIf(IsNull(.Parameters(6).Value), 0, .Parameters(6).Value)

    End With
    Set cmdPolicy = Nothing
  
  Else
  
    ' Password options
    glngDomainMinimumLength = GetSystemSetting("Password", "Minimum Length", 0)
    glngDomainChangeFrequency = GetSystemSetting("Password", "Change Frequency", 0)
    gstrDomainChangePeriod = GetSystemSetting("Password", "Change Period", "")
    giDomainComplexity = GetSystemSetting("Password", "Use Complexity", 0)
    
     ' Lockout settings
    gintDomainAttempts = GetSystemSetting("Misc", "CFG_BA", 3)
    gblnDomainPCLockout = GetSystemSetting("Misc", "CFG_PCL", True)
    glngDomainResetTime = GetSystemSetting("Misc", "CFG_RT", 3600)
    glngDomainLockoutDuration = GetSystemSetting("Misc", "CFG_LD", 300)
    giDomainPasswordsRemembered = 0
    
  End If
  
  LoadDomainSecurityPolicy = True
  
TidyUpAndExit:
  Exit Function

ErrorTrap:
  If gADOCon.Errors.Count > 0 Then
    Select Case gADOCon.Errors(0).NativeError
    ' .NET error, SQL process login details are incorrect
    Case 6522
      If piAssemblyRetryCount > 1 Then
        MsgBox "The OpenHR Server Assembly is out of date." & vbNewLine & _
          "Please ask the System Administrator to update the database in the System Manager.", vbExclamation + vbOKOnly, App.Title
      Else
        Call ReRegisterAssembly
        Resume ExecuteSP
      End If
      
    Case Else
      MsgBox "Error initialising the Domain Security Settings." & vbNewLine & _
              IIf(glngSQLVersion = 8, "Ensure the OpenHR Server DLL is registered." & vbNewLine, "") & _
             "(" & gADOCon.Errors(0).NativeError & " - " & gADOCon.Errors(0).Description & ")", vbExclamation + vbOKOnly, App.Title
    End Select
  Else
    MsgBox "Error initialising the Domain Security Settings." & vbNewLine & _
          IIf(glngSQLVersion = 8, "Ensure the OpenHR Server DLL is registered." & vbNewLine, "") & _
         "(" & Err.Number & " - " & Err.Description & ")", vbExclamation + vbOKOnly, App.Title
  End If
  
  LoadDomainSecurityPolicy = False
  GoTo TidyUpAndExit
  

End Function

Private Function ReRegisterAssembly()
  'NPG20090323 Fault 13618
  'Re-register the assembly
  Dim sProcSQL As String
  On Error Resume Next
  sProcSQL = "DECLARE @sSPCode_0 nvarchar(4000)" & vbNewLine & _
      "DECLARE @FrameworkPath nvarchar(4000)" & vbNewLine & _
      "DECLARE @bIs64Bit bit" & vbNewLine & vbNewLine & _
      "-- 64 bit ?" & vbNewLine & _
      "SELECT @bIs64Bit = CASE PATINDEX ('%X64)%' , @@version) WHEN 0 THEN 0 ELSE 1 END" & vbNewLine & _
      "IF @bIs64Bit = 1" & vbNewLine & _
      "  SET @FrameworkPath = 'C:\Windows\Microsoft.NET\Framework64\v2.0.50727\System.DirectoryServices.DLL'" & vbNewLine & _
      "ELSE" & vbNewLine & _
      "  SET @FrameworkPath = 'C:\Windows\Microsoft.NET\Framework\v2.0.50727\System.DirectoryServices.DLL'" & vbNewLine & _
      "ALTER ASSEMBLY [System.DirectoryServices]" & vbNewLine & _
      "FROM @FrameworkPath;"
  
  gADOCon.Execute sProcSQL, , adExecuteNoRecords
End Function
Private Function LC_FormatDateTimeMessage(plngSeconds As Long) As String

  Dim strTemp As String
  Dim intTemp As Integer
  Dim intFix As Integer
  
  intTemp = plngSeconds
  strTemp = vbNullString
  
  'Calculate days
  If intTemp > 86400 Then
    intFix = Fix(intTemp / 86400)
    intTemp = (intTemp Mod 86400)
    strTemp = strTemp & CStr(intFix) & IIf((intFix > 1), " days", " day")
  End If

  'Calculate months
  If intTemp > 3600 Then
    intFix = Fix(intTemp / 3600)
    intTemp = (intTemp Mod 3600)
    strTemp = strTemp & IIf((Len(strTemp) > 0), " ", "")
    strTemp = strTemp & CStr(intFix) & IIf((intFix > 1), " hours", " hour")
  End If
 
  'Calculate minutes
  If intTemp > 60 Then
    intFix = Fix(intTemp / 60)
    intTemp = (intTemp Mod 60)
    strTemp = strTemp & IIf((Len(strTemp) > 0), " ", "")
    strTemp = strTemp & CStr(intFix) & IIf((intFix > 1), " minutes", " minute")
  End If

  'Calculate seconds
  If intTemp > 0 Then
    strTemp = strTemp & IIf((Len(strTemp) > 0), " ", "")
    strTemp = strTemp & CStr(intTemp) & IIf((intTemp > 1), " seconds", " second")
  Else
    If (Len(strTemp) = 0) Then
      strTemp = strTemp & "0 seconds"
    End If
  End If

  LC_FormatDateTimeMessage = strTemp
  
End Function

Public Function LC_Initialise() As Boolean
 
  Dim strCFG_PCL As String
  Dim strCFG_BA As String
  Dim strCFG_RT As String
  Dim strCFG_LD As String
 
  Dim strInput As String
  Dim strEKey As String
  Dim strLens As String
  Dim lngStart As Long
  Dim lngFinish As Long
  
  Dim strOutput As String
  Dim lngCount As Long

  On Error GoTo ErrorTrap
  
  strInput = GetPCSetting("Misc", "ASRSysParam1", vbNullString)
  If strInput = vbNullString Then
    strCFG_PCL = "1"
    strCFG_BA = "3"
    strCFG_LD = "300"
  Else
    lngStart = Len(strInput) - 13
    strEKey = Mid(strInput, lngStart + 1, 10)
    strLens = Right(strInput, 3)
    strInput = XOREncript(Left(strInput, lngStart), strEKey)
  
    lngStart = 1
    lngFinish = Asc(Mid(strLens, 1, 1)) - 127
    strCFG_PCL = Mid(strInput, lngStart, lngFinish)
    
    lngStart = lngStart + lngFinish
    lngFinish = Asc(Mid(strLens, 2, 1)) - 127
    strCFG_BA = Mid(strInput, lngStart, lngFinish)
    
    lngStart = lngStart + lngFinish
    lngFinish = Asc(Mid(strLens, 3, 1)) - 127
    strCFG_LD = Mid(strInput, lngStart, lngFinish)
  End If

  gblnCFG_PCL = CBool(strCFG_PCL)
  gintCFG_BA = CInt(strCFG_BA)
  glngCFG_LD = CLng(strCFG_LD)
  
  
  strInput = GetPCSetting("Misc", "ASRSysParam3", vbNullString)
  If strInput = vbNullString Then
    strCFG_RT = "3600"
  Else
    lngStart = Len(strInput) - 11
    strEKey = Mid(strInput, lngStart + 1, 10)
    strLens = Right(strInput, 1)
    strInput = XOREncript(Left(strInput, lngStart), strEKey)
  
    lngStart = 1
    lngFinish = Asc(Mid(strLens, 1, 1)) - 127
    strCFG_RT = Mid(strInput, lngStart, lngFinish)
  End If
  
  glngCFG_RT = CLng(strCFG_RT)
  
  
  LC_Initialise = True
  
TidyUpAndExit:
  Exit Function

ErrorTrap:
  LC_Initialise = False
  MsgBox "Error Reading Security Settings" & vbCrLf & vbCrLf & Err.Number & " " & Err.Description, vbOKOnly + vbExclamation, App.Title
  GoTo TidyUpAndExit
 
End Function


Private Function LC_ReadCurrentLockStatus() As Boolean

  Dim strPC_BA As String
  Dim strPC_LA As String
  Dim strPC_LKD As String

  Dim strInput As String
  Dim strEKey As String
  Dim strLens As String
  Dim lngStart As Long
  Dim lngFinish As Long
  
  On Error GoTo ErrorTrap
  
  strInput = GetPCSetting("Misc", "ASRSysParam2", vbNullString)
  If strInput = vbNullString Then
    strPC_BA = "0"
    strPC_LKD = "0"
  
  Else
    lngStart = Len(strInput) - 12
    strEKey = Mid(strInput, lngStart + 1, 10)
    strLens = Right(strInput, 2)
    strInput = XOREncript(Left(strInput, lngStart), strEKey)
  
    lngStart = 1
    lngFinish = Asc(Mid(strLens, 1, 1)) - 127
    strPC_BA = Mid(strInput, lngStart, lngFinish)
    
    lngStart = lngStart + lngFinish
    lngFinish = Asc(Mid(strLens, 2, 1)) - 127
    strPC_LKD = Mid(strInput, lngStart, lngFinish)
    
  End If
  
  gintPC_BA = strPC_BA
  gdtPC_LKD = strPC_LKD


  strInput = GetPCSetting("Misc", "ASRSysParam4", vbNullString)
  If strInput = vbNullString Then
    strPC_LA = "0"
  
  Else
    lngStart = Len(strInput) - 11
    strEKey = Mid(strInput, lngStart + 1, 10)
    strLens = Right(strInput, 1)
    strInput = XOREncript(Left(strInput, lngStart), strEKey)
  
    lngStart = 1
    lngFinish = Asc(Mid(strLens, 1, 1)) - 127
    strPC_LA = Mid(strInput, lngStart, lngFinish)
  
  End If
  
  gdtPC_LA = strPC_LA

  LC_ReadCurrentLockStatus = True
  
TidyUpAndExit:
  Exit Function

ErrorTrap:
  LC_ReadCurrentLockStatus = False
  MsgBox "Error Reading PC Security Settings" & vbCrLf & vbCrLf & Err.Number & " " & Err.Description, vbOKOnly + vbExclamation, App.Title
  GoTo TidyUpAndExit
    
End Function

Public Function LC_ResetLock() As Boolean
  
  'Exit the Login Check functions if the PC Lockout functionality has been disabled.
  If Not gblnCFG_PCL Then
    LC_ResetLock = True
    Exit Function
  End If
  
  LC_SaveCurrentLockStatus "0", "0", "0"
  
End Function

Public Function LC_SaveSettingsToRegistry() As Boolean
  
  Dim strCFG_PCL As String
  Dim strCFG_BA As String
  Dim strCFG_RT As String
  Dim strCFG_LD As String

  Dim strEKey As String
  Dim strLens As String
  Dim strOutput As String
  Dim lngCount As Long

  On Error GoTo ErrorTrap
  
  'Reset the PC Lock as there must have been a successful login!
  LC_ResetLock
  
  strCFG_PCL = GetSystemSetting("Misc", "CFG_PCL", "1")
  strCFG_BA = GetSystemSetting("Misc", "CFG_BA", "3")
  strCFG_LD = GetSystemSetting("Misc", "CFG_LD", "300")
  strOutput = strCFG_PCL & strCFG_BA & strCFG_LD
  strLens = Chr(Len(strCFG_PCL) + 127) & Chr(Len(strCFG_BA) + 127) & _
            Chr(Len(strCFG_LD) + 127)
  strEKey = vbNullString
  For lngCount = 1 To 10
    strEKey = strEKey & Chr(Int(Rnd * 255) + 1)
  Next
  strOutput = XOREncript(strOutput, strEKey) & strEKey & strLens
  
  SavePCSetting "Misc", "ASRSysParam1", strOutput

  
  strCFG_RT = GetSystemSetting("Misc", "CFG_RT", "3600")
  strOutput = strCFG_RT
  strLens = Chr(Len(strCFG_RT) + 127)
  strEKey = vbNullString
  For lngCount = 1 To 10
    strEKey = strEKey & Chr(Int(Rnd * 255) + 1)
  Next
  strOutput = XOREncript(strOutput, strEKey) & strEKey & strLens
  
  SavePCSetting "Misc", "ASRSysParam3", strOutput

TidyUpAndExit:
  Exit Function

ErrorTrap:
  MsgBox "Error Saving Security Settings" & vbCrLf & vbCrLf & Err.Number & " " & Err.Description, vbOKOnly + vbExclamation, App.Title
  GoTo TidyUpAndExit
  
End Function


Public Function LC_IncrementBadAttempt() As Boolean

  Dim intBadAttempts As Integer
  
  'Exit the Login Check functions if the PC Lockout functionality has been disabled.
  If Not gblnCFG_PCL Then
    LC_IncrementBadAttempt = True
    Exit Function
  End If
  
  LC_ReadCurrentLockStatus
  
  intBadAttempts = gintPC_BA + 1
  
  If (intBadAttempts >= gintCFG_BA) Then
    LC_SaveCurrentLockStatus CStr(intBadAttempts), CStr(Format(Now(), "yyyy/mm/dd hh:mm:ss")), CStr(Format(Now(), "yyyy/mm/dd hh:mm:ss"))
    LC_PCLocked
  Else
    LC_SaveCurrentLockStatus CStr(intBadAttempts), CStr(Format(Now(), "yyyy/mm/dd hh:mm:ss")), "0"
  End If
 
  'Re-initialise PC variables as these are always refreshed by "LC_ReadCurrentLockStatus"
  gintPC_BA = 0
  gdtPC_LA = 0
  gdtPC_LKD = 0
 
End Function

Public Function LC_PCLocked() As Boolean
  
  Dim blnPCLocked As Boolean
  Dim strMessage As String
  Dim strTimeRemaining As String
 
  strMessage = vbNullString
  strTimeRemaining = vbNullString
 
  'Exit the Login Check functions if the PC Lockout functionality has been disabled.
  If Not gblnCFG_PCL Then
    LC_PCLocked = False
    Exit Function
  End If

  LC_ReadCurrentLockStatus
  
  If (Now() >= DateAdd("s", glngCFG_RT, gdtPC_LA)) Then
    blnPCLocked = False
    LC_ResetLock
  Else
    If (gintPC_BA >= gintCFG_BA) Then
      If (DateAdd("s", glngCFG_LD, gdtPC_LKD) >= Now()) Then
        blnPCLocked = True
      Else
        blnPCLocked = False
        LC_ResetLock
      End If
    Else
      blnPCLocked = False
    End If
  End If
  
  If blnPCLocked Then
    strTimeRemaining = LC_FormatDateTimeMessage(DateDiff("s", Now(), DateAdd("s", glngCFG_LD, gdtPC_LKD)))
    strMessage = "This PC has been temporarily locked from using OpenHR." & vbCrLf & vbCrLf & "The lock will be removed in " & strTimeRemaining & "."
    MsgBox strMessage, vbOKOnly + vbExclamation, App.Title
  End If
  
  'Re-initialise PC variables as these are always refreshed by "LC_ReadCurrentLockStatus"
  gintPC_BA = 0
  gdtPC_LA = 0
  gdtPC_LKD = 0
  
  LC_PCLocked = blnPCLocked
  
End Function


Private Function LC_SaveCurrentLockStatus(pstrBadAttempts As String, pstrLastAttempt As String, pstrLocked As String) As Boolean

  Dim strEKey As String
  Dim strLens As String
  Dim strOutput As String
  Dim lngCount As Long

  On Error GoTo ErrorTrap
  
  strOutput = pstrBadAttempts & pstrLocked
  strLens = Chr(Len(pstrBadAttempts) + 127) & Chr(Len(pstrLocked) + 127)
  strEKey = vbNullString
  For lngCount = 1 To 10
    strEKey = strEKey & Chr(Int(Rnd * 255) + 1)
  Next
  strOutput = XOREncript(strOutput, strEKey) & strEKey & strLens
  
  SavePCSetting "Misc", "ASRSysParam2", strOutput


  strOutput = pstrLastAttempt
  strLens = Chr(Len(pstrLastAttempt) + 127)
  strEKey = vbNullString
  For lngCount = 1 To 10
    strEKey = strEKey & Chr(Int(Rnd * 255) + 1)
  Next
  strOutput = XOREncript(strOutput, strEKey) & strEKey & strLens
  
  SavePCSetting "Misc", "ASRSysParam4", strOutput
  
  
  LC_SaveCurrentLockStatus = True
  
TidyUpAndExit:
  Exit Function

ErrorTrap:
  LC_SaveCurrentLockStatus = False
  MsgBox "Error Saving PC Security Settings" & vbCrLf & vbCrLf & Err.Number & " " & Err.Description, vbOKOnly + vbExclamation, App.Title
  GoTo TidyUpAndExit
  
End Function

Public Function SavePCSetting(strSection As String, strKey As String, varSetting As Variant) As Boolean
  'Trap error in case user doesn't have permission to write to the registry
  On Local Error Resume Next
  SaveSetting "HR Pro", strSection, strKey, varSetting
End Function

Private Function XOREncript(strInput, strKey) As String

  Dim lngCount As Long
  Dim strOutput As String
  Dim strChar As String
  
  For lngCount = 1 To Len(strInput)
    strChar = Mid(strKey, lngCount Mod Len(strKey) + 1, 1)
    strOutput = strOutput & Chr(Asc(strChar) Xor Asc(Mid(strInput, lngCount, 1)))
  Next

  XOREncript = strOutput

End Function

Public Function GetPCSetting(strSection As String, strKey As String, varDefault As Variant) As String
  GetPCSetting = GetSetting("HR Pro", strSection, strKey, varDefault)
End Function

Public Function SaveSystemSetting(strSection As String, strKey As String, varSetting As Variant) As Boolean

  Dim strSQL As String

  DeleteSystemSetting strSection, strKey

  strSQL = "INSERT ASRSysSystemSettings " & _
           "(Section, SettingKey, SettingValue) " & _
           "VALUES " & _
           "('" & LCase(strSection) & "'," & _
           " '" & LCase(strKey) & "'," & _
           " '" & CStr(varSetting) & "')"
  gADOCon.Execute strSQL, , adExecuteNoRecords

End Function


Public Function DeleteSystemSetting(strSection As String, strKey As String) As Boolean

  Dim strSQL As String
  
  strSQL = "DELETE FROM ASRSysSystemSettings " & _
           " WHERE Section = '" & LCase(strSection) & "'" & _
           " AND SettingKey = '" & LCase(strKey) & "'"
  gADOCon.Execute strSQL, , adExecuteNoRecords

End Function


Public Function GetSystemSetting(strSection As String, strKey As String, varDefault As Variant) As Variant
  
  Dim rsTemp As New ADODB.Recordset
  Dim strSQL As String

  On Local Error GoTo LocalErr

  strSQL = "SELECT SettingValue FROM ASRSysSystemSettings " & _
           " WHERE Section = '" & LCase(strSection) & "'" & _
           " AND SettingKey = '" & LCase(strKey) & "'"
  rsTemp.Open strSQL, gADOCon, adOpenForwardOnly, adLockReadOnly
  
  With rsTemp
    If Not .BOF And Not .EOF Then
      GetSystemSetting = rsTemp.Fields("SettingValue").Value
      If IsNull(GetSystemSetting) Then GetSystemSetting = vbNullString
    Else
      GetSystemSetting = varDefault
    End If
  End With

  rsTemp.Close
  Set rsTemp = Nothing

Exit Function

LocalErr:
  GetSystemSetting = vbNullString

End Function


Public Function LockDatabase(intLockType As LockTypes) As Boolean

  Dim rsTemp As New ADODB.Recordset
  Dim objGroup As SecurityGroup
  Dim sGroupsToLogout As String

  gADOCon.BeginTrans

  'Check that no other user has the database locked...
  rsTemp.Open "sp_ASRLockCheck", gADOCon, adOpenDynamic, adLockReadOnly

  LockDatabase = True
  If Not rsTemp.BOF And Not rsTemp.EOF Then
    If LCase(rsTemp!UserName) <> LCase(gsUserName) And _
       Trim(rsTemp!UserName) <> vbNullString Then
            LockDatabase = False
    End If
  End If
  rsTemp.Close
  
  ' What groups to lock out
  For Each objGroup In gObjGroups
    If objGroup.RequireLogout Then
      sGroupsToLogout = sGroupsToLogout & IIf(Len(sGroupsToLogout) > 0, ", ", "") & objGroup.Name
    End If
  Next objGroup
  
  If LockDatabase Then
    gADOCon.Execute "sp_ASRLockWrite " & CStr(intLockType) & ", 2, '" & sGroupsToLogout & "'", , adExecuteNoRecords
    gADOCon.CommitTrans
  Else
    gADOCon.RollbackTrans
  End If

  Set rsTemp = Nothing

End Function


Public Function UnlockDatabase(intLockType As LockTypes, Optional blnForceUnlock As Boolean) As Boolean

  Dim rsTemp As New ADODB.Recordset

  gADOCon.BeginTrans

  UnlockDatabase = True
  If Not blnForceUnlock Then

    'Check that no other user has the database locked...
    rsTemp.Open "sp_ASRLockCheck", gADOCon, adOpenForwardOnly, adLockReadOnly

    If Not rsTemp.BOF And Not rsTemp.EOF Then
      If LCase(rsTemp!UserName) <> LCase(gsUserName) And _
         Trim(rsTemp!UserName) <> vbNullString Then
              UnlockDatabase = False
      End If
    End If

    rsTemp.Close
    Set rsTemp = Nothing

  End If

  If UnlockDatabase Then
    gADOCon.Execute "sp_ASRLockDelete " & CStr(intLockType) & ", 2", , adExecuteNoRecords
    gADOCon.CommitTrans
  Else
    gADOCon.RollbackTrans
  End If

  Exit Function

End Function


Public Function GetTmpFName() As String

  Dim strTmpPath As String, strTmpName As String
  
  strTmpPath = Space(1024)
  strTmpName = Space(1024)

  Call GetTempPath(1024, strTmpPath)
  Call GetTempFileName(strTmpPath, "_T", 0, strTmpName)
  
  strTmpName = Trim(strTmpName)
  If Len(strTmpName) > 0 Then
    strTmpName = Left(strTmpName, Len(strTmpName) - 1)
  Else
    strTmpName = vbNullString
  End If
    
  GetTmpFName = Trim(strTmpName)

End Function

Public Function GetUniqueID(strSetting As String, strTable As String, strColumn As String) As Long

  Dim lngNewMethodID As Long
  Dim lngOldMethodID As Long

  lngOldMethodID = UniqueColumnValue(strTable, strColumn)
  lngNewMethodID = GetSystemSetting("AutoID", strSetting, 0) + 1

  GetUniqueID = IIf(lngOldMethodID > lngNewMethodID, lngOldMethodID, lngNewMethodID)
  SaveSystemSetting "AutoID", strSetting, GetUniqueID

End Function


Public Function SystemPermission(strCategory As String, strItem As String) As Boolean

  Dim cmdSystemPermission As New ADODB.Command
  Dim pmADO As ADODB.Parameter

  With cmdSystemPermission
     .CommandText = "dbo.sp_ASRSystemPermission"
     .CommandType = adCmdStoredProc
     .CommandTimeout = 0
     Set .ActiveConnection = gADOCon
 
     Set pmADO = .CreateParameter("PermissionGranted", adBoolean, adParamOutput)
     .Parameters.Append pmADO
 
     Set pmADO = .CreateParameter("CategoryKey", adVarChar, adParamInput, 50)
     .Parameters.Append pmADO
     pmADO.Value = strCategory

     Set pmADO = .CreateParameter("Item", adVarChar, adParamInput, 50)
     .Parameters.Append pmADO
     pmADO.Value = strItem

     Set pmADO = .CreateParameter("User", adVarChar, adParamInput, 200)
     .Parameters.Append pmADO
     pmADO.Value = gsActualSQLLogin

     .Execute
 
     SystemPermission = .Parameters(0).Value
   End With
   Set cmdSystemPermission = Nothing

End Function

Public Function GetUserSetting(strSection As String, strKey As String, varDefault As Variant) As Variant
  
  Dim rsTemp As New ADODB.Recordset
  Dim strSQL As String
  
  'JPD 20050812 Fault 10166
  strSQL = "SELECT SettingValue FROM ASRSysUserSettings " & _
           " WHERE UserName = '" & LCase(Replace(gsUserName, "'", "''")) & "'" & _
           " AND Section = '" & LCase(strSection) & "'" & _
           " AND SettingKey = '" & LCase(strKey) & "'"
  rsTemp.Open strSQL, gADOCon, adOpenForwardOnly, adLockReadOnly
  With rsTemp
    If Not .BOF And Not .EOF Then
      GetUserSetting = rsTemp!SettingValue
    Else
      GetUserSetting = varDefault
    End If
  End With

  rsTemp.Close

  Set rsTemp = Nothing

End Function

Public Function GetSQLNCLIVersion() As Integer
On Error GoTo SQLNCLI_Err

  Dim rc As Long                                          ' Return Code
  Dim hKey As Long                                        ' Handle To An Open Registry Key
  Dim tmpKey As Integer
  tmpKey = 0
  
  ' Paths to the SQL Native Client registry keys
  Const sREGKEYSQLNCLI = "SOFTWARE\Microsoft\Microsoft SQL Native Client\CurrentVersion"
  Const sREGKEYSQLNCLI10 = "SOFTWARE\Microsoft\Microsoft SQL Server Native Client 10.0\CurrentVersion"
  Const sREGKEYSQLNCLI11 = "SOFTWARE\Microsoft\Microsoft SQL Server Native Client 11.0\CurrentVersion"

  rc = RegOpenKeyEx(HKEY_LOCAL_MACHINE, sREGKEYSQLNCLI, 0, KEY_READ, hKey) ' Open Registry Key
  If (rc = 0) Then
    tmpKey = 9
    rc = RegCloseKey(hKey) ' Close Registry Key
  End If

  rc = RegOpenKeyEx(HKEY_LOCAL_MACHINE, sREGKEYSQLNCLI10, 0, KEY_READ, hKey) ' Open Registry Key
  If (rc = 0) Then
    tmpKey = 10
    rc = RegCloseKey(hKey) ' Close Registry Key
  End If

  rc = RegOpenKeyEx(HKEY_LOCAL_MACHINE, sREGKEYSQLNCLI11, 0, KEY_READ, hKey) ' Open Registry Key
  If (rc = 0) Then
    tmpKey = 11
    rc = RegCloseKey(hKey) ' Close Registry Key
  End If



SQLNCLI_Err_Handler:
  GetSQLNCLIVersion = tmpKey
  Exit Function
    
SQLNCLI_Err:
  rc = RegCloseKey(hKey) ' Close Registry Key
  tmpKey = 0
  Resume SQLNCLI_Err_Handler
End Function


Public Sub SetIcon( _
      ByVal hWnd As Long, _
      ByVal sIconResName As String, _
      Optional ByVal bSetAsAppIcon As Boolean = True _
   )
Dim lhWndTop As Long
Dim lhWnd As Long
Dim cx As Long
Dim cy As Long
Dim hIconLarge As Long
Dim hIconSmall As Long
      
   If (bSetAsAppIcon) Then
      ' Find VB's hidden parent window:
      lhWnd = hWnd
      lhWndTop = lhWnd
      Do While Not (lhWnd = 0)
         lhWnd = GetWindow(lhWnd, GW_OWNER)
         If Not (lhWnd = 0) Then
            lhWndTop = lhWnd
         End If
      Loop
   End If
   
   cx = GetSystemMetrics(SM_CXICON)
   cy = GetSystemMetrics(SM_CYICON)
   hIconLarge = LoadImageAsString( _
         App.hInstance, sIconResName, _
         IMAGE_ICON, _
         cx, cy, _
         LR_SHARED)
   If (bSetAsAppIcon) Then
      SendMessageLong lhWndTop, WM_SETICON, ICON_BIG, hIconLarge
   End If
   SendMessageLong hWnd, WM_SETICON, ICON_BIG, hIconLarge
   
   cx = GetSystemMetrics(SM_CXSMICON)
   cy = GetSystemMetrics(SM_CYSMICON)
   hIconSmall = LoadImageAsString( _
         App.hInstance, sIconResName, _
         IMAGE_ICON, _
         cx, cy, _
         LR_SHARED)
   If (bSetAsAppIcon) Then
      SendMessageLong lhWndTop, WM_SETICON, ICON_SMALL, hIconSmall
   End If
   SendMessageLong hWnd, WM_SETICON, ICON_SMALL, hIconSmall
   
End Sub


