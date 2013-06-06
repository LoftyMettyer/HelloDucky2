Attribute VB_Name = "modSettings"
Option Explicit

Public Const VARCHAR_MAX_Size = 2147483646 'Yup one below the actual max, needs to be otherwise things go so awfully wrong, you don't believe me, well go on then, change it, see if I care!!!)

Private Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, ByRef phkResult As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long

Private Const HKEY_LOCAL_MACHINE = &H80000002
Private Const KEY_READ = &H20019


' --------------------------
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

Private Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" ( _
      ByVal hWnd As Long, ByVal wMsg As Long, _
      ByVal wParam As Long, ByVal lParam As Long _
   ) As Long

Private Const WM_SETICON = &H80

Private Const ICON_SMALL = 0
Private Const ICON_BIG = 1

Private Declare Function GetWindow Lib "user32" ( _
   ByVal hWnd As Long, ByVal wCmd As Long) As Long
Private Const GW_OWNER = 4
' -------------------------------


'Security Settings - Login Checks i.e.Bad Attempts
Private gblnCFG_PCL As Boolean
Private gintCFG_BA As Integer
Private glngCFG_RT As Integer
Private glngCFG_LD As Long
Private gintPC_BA As Integer
Private gdtPC_LA As Date
Private gdtPC_LKD As Date

Public Enum PasswordChangeReason
  giPasswordChange_None = 0
  giPasswordChange_MinLength = 1
  giPasswordChange_Expired = 2
  giPasswordChange_AdminRequested = 3
  giPasswordChange_LastChangeUnknown = 4
  giPasswordChange_ComplexitySettings = 5
End Enum

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
  COAMsgBox "Error Reading Security Settings" & vbCrLf & vbCrLf & Err.Number & " " & Err.Description, vbOKOnly + vbExclamation, App.Title
  GoTo TidyUpAndExit
 
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
    strMessage = "This PC has been temporarily locked from using " & Application.Name & "." & vbCrLf & vbCrLf & "The lock will be removed in " & strTimeRemaining & "."
    COAMsgBox strMessage, vbOKOnly + vbExclamation, App.Title
  End If
  
  'Re-initialise PC variables as these are always refreshed by "LC_ReadCurrentLockStatus"
  gintPC_BA = 0
  gdtPC_LA = 0
  gdtPC_LKD = 0
  
  LC_PCLocked = blnPCLocked
  
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
  COAMsgBox "Error Reading PC Security Settings" & vbCrLf & vbCrLf & Err.Number & " " & Err.Description, vbOKOnly + vbExclamation, App.Title
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
  COAMsgBox "Error Saving PC Security Settings" & vbCrLf & vbCrLf & Err.Number & " " & Err.Description, vbOKOnly + vbExclamation, App.Title
  GoTo TidyUpAndExit
  
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
  COAMsgBox "Error Saving Security Settings" & vbCrLf & vbCrLf & Err.Number & " " & Err.Description, vbOKOnly + vbExclamation, App.Title
  GoTo TidyUpAndExit
  
End Function

Public Function SavePCSetting(strSection As String, strKey As String, varSetting As Variant) As Boolean
  'Trap error in case user doesn't have permission to write to the registry
  On Local Error Resume Next
  SaveSetting "HR Pro", strSection, strKey, varSetting
End Function

Public Function GetPCSetting(strSection As String, strKey As String, varDefault As Variant) As String
  GetPCSetting = GetSetting("HR Pro", strSection, strKey, varDefault)
End Function


Public Function SaveUserSetting(strSection As String, strKey As String, varSetting As Variant) As Boolean

  Dim cmADO As ADODB.Command
  Dim pmADO As ADODB.Parameter
  Dim strSQL As String

  'JPD 20040209 Fault 8065
  If UCase(Left(strSection, 9)) = "FINDORDER" Then
    varSetting = CLng(varSetting)
  End If

  Set cmADO = New ADODB.Command
  With cmADO
    .CommandText = "dbo.[spASRSaveUserSetting]"
    .CommandType = adCmdStoredProc
    .CommandTimeout = 0
    Set .ActiveConnection = gADOCon
            
    Set pmADO = .CreateParameter("Section", adVarChar, adParamInput, 50)
    .Parameters.Append pmADO
    pmADO.Value = strSection
  
    'NPG20080418 Fault 13100
    ' Set pmADO = .CreateParameter("SettingKey", adVarChar, adParamInput, 50)
    Set pmADO = .CreateParameter("SettingKey", adVarChar, adParamInput, 255)
    .Parameters.Append pmADO
    pmADO.Value = strKey

    Set pmADO = .CreateParameter("SettingValue", adVarChar, adParamInput, 255)
    .Parameters.Append pmADO
    pmADO.Value = varSetting
    
    .Execute
  End With

End Function

Public Function SetUserSettingDefaults() As Boolean
  
'********************************************************************************
' SetUserSettingDefaults -  Removes all user settings from the                  *
'                           ASRSysUserSettings table.                           *
'Added as part of suggestion TM20010726 Fault 1607.                             *
'********************************************************************************

  On Error GoTo ErrorTrap

  Dim datData As clsDataAccess
  Dim strSQL As String
  
  Set datData = New clsDataAccess
 
  'Delete all user configurations from the User Settings table.
  strSQL = "DELETE FROM ASRSysUserSettings " & _
           "WHERE UserName = '" & LCase(datGeneral.UserNameForSQL) & "' "

  datData.ExecuteSql strSQL

  Set datData = Nothing
  
  SetUserSettingDefaults = True
  
  Exit Function
  
ErrorTrap:
  SetUserSettingDefaults = False
  
End Function

Public Function SaveSystemSetting(strSection As String, strKey As String, varSetting As Variant) As Boolean

  Dim datData As clsDataAccess
  Dim strSQL As String
  
  Set datData = New clsDataAccess
  
  strSQL = "DELETE FROM ASRSysSystemSettings " & _
           " WHERE Section = '" & Replace(LCase(strSection), "'", "''") & "'" & _
           " AND SettingKey = '" & Replace(LCase(strKey), "'", "''") & "'"
  datData.ExecuteSql strSQL

  strSQL = "INSERT ASRSysSystemSettings " & _
           "(Section, SettingKey, SettingValue) " & _
           "VALUES " & _
           "('" & Replace(LCase(strSection), "'", "''") & "'," & _
           " '" & Replace(LCase(strKey), "'", "''") & "'," & _
           " '" & Replace(CStr(varSetting), "'", "''") & "')"
  datData.ExecuteSql strSQL

  Set datData = Nothing

End Function

Public Function GetSystemSetting(strSection As String, strKey As String, varDefault As Variant) As Variant
  
  Dim datData As clsDataAccess
  Dim rsTemp As Recordset
  Dim strSQL As String
  
  On Local Error GoTo LocalErr

  Set datData = New clsDataAccess

  strSQL = "SELECT SettingValue FROM ASRSysSystemSettings " & _
           " WHERE Section = '" & Replace(LCase(strSection), "'", "''") & "'" & _
           " AND SettingKey = '" & Replace(LCase(strKey), "'", "''") & "'"
  Set rsTemp = datData.OpenRecordset(strSQL, adOpenForwardOnly, adLockReadOnly)

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
  Set datData = Nothing

Exit Function

LocalErr:
  GetSystemSetting = vbNullString

End Function

Public Function GetUserOrSystemSetting(strSection As String, strKey As String, varDefault As Variant) As String
  GetUserOrSystemSetting = GetUserSetting(strSection, strKey, GetSystemSetting(strSection, strKey, varDefault))
End Function

Public Function SaveBatchLogon(strUserName As String, strPassword As String, strDatabase As String, strServer As String) As Boolean

  Dim strOutput As String
  Dim strEKey As String
  Dim strLens As String
  Dim lngCount As Long

  strOutput = strUserName & strPassword & strDatabase & strServer
  strLens = Chr(Len(strUserName) + 127) & Chr(Len(strPassword) + 127) & _
            Chr(Len(strDatabase) + 127) & Chr(Len(strServer) + 127)
  
  strEKey = vbNullString
  For lngCount = 1 To 10
    strEKey = strEKey & Chr(Int(Rnd * 255) + 1)
  Next

  strOutput = XOREncript(strOutput, strEKey) & strEKey & strLens
  SavePCSetting "BatchLogon", "Data", strOutput

End Function

'Public Function GetBatchLogon(ByVal strUserName As String, ByVal strPassword As String, ByVal strDatabase As String, ByVal strServer As String) As Boolean
Public Function GetBatchLogon(strUserName As String, strPassword As String, strDatabase As String, strServer As String) As Boolean

  Dim strInput As String
  Dim strEKey As String
  Dim strLens As String
  Dim lngStart As Long
  Dim lngFinish As Long

  strInput = GetPCSetting("BatchLogon", "Data", vbNullString)
  If strInput = vbNullString Then
    Exit Function
  End If

  lngStart = Len(strInput) - 14
  strEKey = Mid(strInput, lngStart + 1, 10)
  strLens = Right(strInput, 4)
  strInput = XOREncript(Left(strInput, lngStart), strEKey)

  lngStart = 1
  lngFinish = Asc(Mid(strLens, 1, 1)) - 127
  strUserName = Mid(strInput, lngStart, lngFinish)

  lngStart = lngStart + lngFinish
  lngFinish = Asc(Mid(strLens, 2, 1)) - 127
  strPassword = Mid(strInput, lngStart, lngFinish)

  lngStart = lngStart + lngFinish
  lngFinish = Asc(Mid(strLens, 3, 1)) - 127
  strDatabase = Mid(strInput, lngStart, lngFinish)

  lngStart = lngStart + lngFinish
  lngFinish = Asc(Mid(strLens, 4, 1)) - 127
  strServer = Mid(strInput, lngStart, lngFinish)

End Function

Public Function XOREncript(strInput, strKey) As String

  Dim lngCount As Long
  Dim strOutput As String
  Dim strChar As String
  
  For lngCount = 1 To Len(strInput)
    strChar = Mid(strKey, lngCount Mod Len(strKey) + 1, 1)
    strOutput = strOutput & Chr(Asc(strChar) Xor Asc(Mid(strInput, lngCount, 1)))
  Next

  XOREncript = strOutput

End Function

'MH20010712
Public Function GetUniqueID(strSetting As String, strTable As String, strColumn As String) As Long

  Dim lngNewMethodID As Long    'From ASRSysSettings
  Dim lngOldMethodID As Long    'SELECT MAX ID

  lngOldMethodID = UniqueColumnValue(strTable, strColumn)
  lngNewMethodID = GetSystemSetting("AutoID", strSetting, 0) + 1

  GetUniqueID = IIf(lngOldMethodID > lngNewMethodID, lngOldMethodID, lngNewMethodID)
  SaveSystemSetting "AutoID", strSetting, GetUniqueID

End Function

' Alters the passed in activebar according to whats specified by user configuration
Public Sub OrganiseToolbarControls(pabActiveBar As ActiveBarLibraryCtl.ActiveBar)
  Dim strBandName As String
  Dim strKey As String
  Dim iBandCount As Integer
  Dim iToolCount As Integer
  Dim iToolbarPosition As ToolbarPositions
  Dim iOriginalToolCount As Integer
  Dim iTemp As Integer
  
  ' Get toolbar position
  iToolbarPosition = GetUserSetting("Toolbar", "Position", giTOOLBAR_TOP)

  'JPD 20030827 Fault 6036 & Fault 6598
  ' Original implementation of this subroutine copied the menu tools to a temp menu,
  ' cleared out the original menu, and then put the copied tools back into the original
  ' menu in the required order. However this caused problems when using the SetPicture method
  ' of the tool object to copy the tool image. If the top-left pixel of the image was NOT
  ' transparent then the SetPicture method would lose the image's transparency and the image would
  ' appear screwed up when the tool was disabled.
  ' To avoid this I've avoided using the SetPicture method. Instead of copying the tools
  ' to a temporary menu, I simply copy them into the original menu in the required order,
  ' and then remove the tools that are not required.
  For iBandCount = 0 To pabActiveBar.Bands.Count - 1
    strBandName = pabActiveBar.Tag & "%%" & pabActiveBar.Bands(iBandCount).Name
    pabActiveBar.Bands(iBandCount).Visible = pabActiveBar.Bands(iBandCount).Visible And (iToolbarPosition <> giTOOLBAR_NONE)
    pabActiveBar.Bands(iBandCount).DockingArea = iToolbarPosition
    pabActiveBar.Bands(iBandCount).DisplayHandles = False

    iOriginalToolCount = pabActiveBar.Bands(iBandCount).Tools.Count
    For iToolCount = 0 To iOriginalToolCount - 1
      strKey = GetUserSetting("toolbar_order", strBandName & "%%" & Trim(Str(iToolCount)), "notcustomised")

      If (strKey <> "notcustomised") And (Not gbReadToolbarDefaults) Then
        pabActiveBar.Bands(iBandCount).Tools.Insert iToolCount, pabActiveBar.Bands(iBandCount).Tools(strKey)
        
        'JPD 20040220
        'pabActiveBar.Bands(iBandCount).Tools(iToolCount).Visible = GetUserSetting("toolbar_showtool", strBandName & "%%" & Trim(Str(pabActiveBar.Bands(iBandCount).Tools(strKey).ToolID)), True) Or gbReadToolbarDefaults
        pabActiveBar.Bands(iBandCount).Tools(iToolCount).Visible = CBool(GetUserSetting("toolbar_showtool", strBandName & "%%" & Trim(Str(pabActiveBar.Bands(iBandCount).Tools(strKey).ToolID)), True)) Or gbReadToolbarDefaults
      
        'JPD 20030908 Fault 6917
        For iTemp = iToolCount + 1 To pabActiveBar.Bands(iBandCount).Tools.Count
          If pabActiveBar.Bands(iBandCount).Tools(iTemp).ToolID = strKey Then
            pabActiveBar.Bands(iBandCount).Tools.Remove iTemp
            Exit For
          End If
        Next iTemp
      End If
    Next iToolCount

    'JPD 20030908 Fault 6917
    'Do While iOriginalToolCount < pabActiveBar.Bands(iBandCount).Tools.Count
    '  pabActiveBar.Bands(iBandCount).Tools.Remove pabActiveBar.Bands(iBandCount).Tools.Count - 1
    'Loop
  Next iBandCount

End Sub

Public Function GetUserSetting(strSection As String, strKey As String, varDefault As Variant) As Variant
   
  On Error GoTo ErrorTrap
   
  Dim cmADO As ADODB.Command
  Dim pmADO As ADODB.Parameter
  
  Dim sTemp1 As String
  Dim sTemp2 As String
  Dim iLoop As Integer
  
TryAgain:
  
  Set cmADO = New ADODB.Command
  With cmADO
    .CommandText = "dbo.spASRGetUserSetting"
    .CommandType = adCmdStoredProc
    .CommandTimeout = 0
    Set .ActiveConnection = gADOCon

    Set pmADO = .CreateParameter("Section", adVarChar, adParamInput, 50)
    .Parameters.Append pmADO
    pmADO.Value = strSection

    Set pmADO = .CreateParameter("SectionKey", adVarChar, adParamInput, 50)
    .Parameters.Append pmADO
    pmADO.Value = strKey

    Set pmADO = .CreateParameter("SettingValue", adVarChar, adParamOutput, 255)
    .Parameters.Append pmADO

    .Execute
    
    GetUserSetting = .Parameters(2).Value
  
    If IsNull(GetUserSetting) Then
      GetUserSetting = varDefault
    End If
  
  End With
    
  Set pmADO = Nothing

  'JPD 20040209 Fault 8065
  ' Drop any decimal places from the Find Order column widths
  ' to avoid French localisation errors.
  If UCase$(Left$(strSection, 9)) = "FINDORDER" Then
    sTemp2 = ""

    For iLoop = 1 To Len(GetUserSetting)
      sTemp1 = Mid(GetUserSetting, iLoop, 1)

      If (sTemp1 = "0") Or _
        (sTemp1 = "1") Or _
        (sTemp1 = "2") Or _
        (sTemp1 = "3") Or _
        (sTemp1 = "4") Or _
        (sTemp1 = "5") Or _
        (sTemp1 = "6") Or _
        (sTemp1 = "7") Or _
        (sTemp1 = "8") Or _
        (sTemp1 = "9") Then

        sTemp2 = sTemp2 & sTemp1
      Else
        Exit For
      End If
    Next iLoop

    GetUserSetting = sTemp2
  End If
  
  Exit Function

ErrorTrap:

  If InStr(1, Err.Description, CONNECTIONBROKEN_MESSAGE, vbTextCompare) Then
    datGeneral.ReEstablishADOConnection
    Resume TryAgain
  End If

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

SQLNCLI_Err_Handler:
  GetSQLNCLIVersion = tmpKey
  Exit Function
    
SQLNCLI_Err:
  rc = RegCloseKey(hKey) ' Close Registry Key
  tmpKey = 0
  Resume SQLNCLI_Err_Handler
End Function

' Get native provider string
Public Function GetSQLProviderString() As String

  If GetSQLNCLIVersion = 9 Then
    GetSQLProviderString = "Provider=SQLNCLI;"
  ElseIf GetSQLNCLIVersion = 10 Then
    GetSQLProviderString = "Provider=SQLNCLI10;"
  Else
    GetSQLProviderString = vbNullString
  End If

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


