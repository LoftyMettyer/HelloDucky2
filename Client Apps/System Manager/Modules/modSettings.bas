Attribute VB_Name = "modSettings"
Option Explicit

Private Const HKEY_LOCAL_MACHINE = &H80000002
Private Const KEY_READ = &H20019

' Setting icons
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
Private Const SM_CXSMICON = 49
Private Const SM_CYSMICON = 50
Private Const GW_OWNER = 4
                                              
' --------------------------------------------------------------------------------------------

'Security Settings - Login Checks i.e.Bad Attempts
Private gblnCFG_PCL As Boolean
Private gintCFG_BA As Integer
Private glngCFG_RT As Integer
Private glngCFG_LD As Long
Private gintPC_BA As Integer
Private gdtPC_LA As Date
Private gdtPC_LKD As Date

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
  MsgBox "Error Reading Security Settings" & vbCrLf & vbCrLf & Err.Number & " " & Err.Description, vbOKOnly + vbExclamation, App.Title
  GoTo TidyUpAndExit
 
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

Public Function SavePCSetting(strSection As String, strKey As String, varSetting As Variant) As Boolean
  'Trap error in case user doesn't have permission to write to the registry
  On Local Error Resume Next
  SaveSetting "HR Pro", strSection, strKey, varSetting
End Function

Public Function GetPCSetting(strSection As String, strKey As String, varDefault As Variant) As String
  GetPCSetting = GetSetting("HR Pro", strSection, strKey, varDefault)
End Function

Public Function SaveSystemSetting(strSection As String, strKey As String, varSetting As Variant) As Boolean

  Dim cmADO As ADODB.Command
  Dim pmADO As ADODB.Parameter

  On Error GoTo LocalErr
   
  Set cmADO = New ADODB.Command
  With cmADO
    .CommandText = "spsys_setsystemsetting"
    .CommandType = adCmdStoredProc
    .CommandTimeout = 0
    Set .ActiveConnection = gADOCon

    Set pmADO = .CreateParameter("@section", adVarChar, adParamInput, 50)
    .Parameters.Append pmADO
    pmADO.value = strSection
    
    Set pmADO = .CreateParameter("@settingkey", adVarChar, adParamInput, 50)
    .Parameters.Append pmADO
    pmADO.value = strKey

    Set pmADO = .CreateParameter("@settingvalue", adVarChar, adParamInput, VARCHAR_MAX_Size)
    .Parameters.Append pmADO
    pmADO.value = varSetting

    cmADO.Execute

  End With
  Set pmADO = Nothing
  Set cmADO = Nothing

  SaveSystemSetting = True

Exit Function

LocalErr:
  SaveSystemSetting = False

End Function


'Public Function DeleteSystemSetting(strSection As String, strKey As String) As Boolean
'
'  Dim strSQL As String
'
'  strSQL = "DELETE FROM ASRSysSystemSettings " & _
'           " WHERE Section = '" & LCase(strSection) & "'" & _
'           " AND SettingKey = '" & LCase(strKey) & "'"
'  gADOCon.Execute strSQL, , adExecuteNoRecords
'
'End Function


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
      GetSystemSetting = rsTemp.Fields("SettingValue").value
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

  Dim rsTemp As ADODB.Recordset

  gADOCon.BeginTrans
  Set rsTemp = New ADODB.Recordset
  
  'Check that no other user has the database locked...
  rsTemp.Open "dbo.sp_ASRLockCheck", gADOCon, adOpenForwardOnly, adLockOptimistic

  LockDatabase = True
  If Not rsTemp.BOF And Not rsTemp.EOF Then
    If LCase(rsTemp.Fields("UserName").value) <> LCase(gsUserName) And _
       Trim(rsTemp.Fields("UserName").value) <> vbNullString Then
            LockDatabase = False
    End If
  End If
  
  rsTemp.Close
  
  If LockDatabase Then
    gADOCon.Execute "sp_ASRLockWrite " & CStr(intLockType)
    gADOCon.CommitTrans
  Else
    gADOCon.RollbackTrans
  End If

  
  Set rsTemp = Nothing

End Function


Public Function UnlockDatabase(intLockType As LockTypes, Optional blnForceUnlock As Boolean) As Boolean
  On Error GoTo ErrorTrap
  
  Dim bOK As Boolean
  Dim rsTemp As ADODB.Recordset
  
  gADOCon.BeginTrans
  Set rsTemp = New ADODB.Recordset
  
  bOK = True
  If Not blnForceUnlock Then

    'Check that no other user has the database locked...
    rsTemp.Open "dbo.sp_ASRLockCheck", gADOCon, adOpenForwardOnly, adLockOptimistic

    If Not rsTemp.BOF And Not rsTemp.EOF Then
      If LCase(rsTemp!userName) <> LCase(gsUserName) And _
         Trim(rsTemp!userName) <> vbNullString Then
              UnlockDatabase = False
      End If
    End If

    rsTemp.Close
    Set rsTemp = Nothing
  
  End If

TidyUpAndExit:
  If bOK Then
    gADOCon.Execute "sp_ASRLockDelete " & CStr(intLockType)
    gADOCon.CommitTrans
  Else
    gADOCon.RollbackTrans
  End If
  
  Exit Function
  
ErrorTrap:
  bOK = False
  Resume TidyUpAndExit
End Function


'MH20010712
Public Function GetUniqueID(strSetting As String, strTable As String, strColumn As String) As Long

  Dim lngNewMethodID As Long    'From ASRSysSettings
  Dim lngOldMethodID As Long    'SELECT MAX ID

'BEGIN TRANS
  gADOCon.BeginTrans

  lngOldMethodID = UniqueColumnValue(strTable, strColumn)
  lngNewMethodID = GetSystemSetting("AutoID", strSetting, 0) + 1

  GetUniqueID = IIf(lngOldMethodID > lngNewMethodID, lngOldMethodID, lngNewMethodID)
  SaveSystemSetting "AutoID", strSetting, GetUniqueID

  gADOCon.CommitTrans
'COMMIT TRANS

End Function

Public Function EncryptLogonDetails(ByVal strUserName As String, ByVal strPassword As String _
  , ByVal strDatabase As String, ByVal strServer As String) As String

  Dim strOutput As String
  Dim strEKey As String
  Dim strLens As String
  Dim lngCount As Long
  Dim iChar As Integer
  Dim strEncryptionDetails As String
  
  strOutput = strUserName & strPassword & strDatabase & strServer
  strLens = Chr(Len(strUserName) + 127) & Chr(Len(strPassword) + 127) & _
            Chr(Len(strDatabase) + 127) & Chr(Len(strServer) + 127)
  
  ' AE20080228 Fault #12939, #12959
'  strEKey = vbNullString
'  For lngCount = 1 To 10
'    iChar = 0
'    Do While iChar = 144 Or iChar = 0
'      iChar = Int(Rnd * 255) + 1
'    Loop
'    strEKey = strEKey & Chr(iChar)
'  Next


  ' NPG20081111 Fault 13373
  ' Do While EncryptLogonDetails = vbNullString _
  '   Or (CBool(InStr(EncryptLogonDetails, Chr(0))) Or CBool(InStr(EncryptLogonDetails, Chr(144))))
  'Do While strEncryptionDetails = vbNullString
    strEncryptionDetails = vbNullString
    strEKey = vbNullString
    
    For lngCount = 1 To 10
      iChar = 0
      iChar = Int(Rnd * 255) + 1
      strEKey = strEKey & Chr(iChar)
    Next
  
    strEncryptionDetails = XOREncript(strOutput, strEKey) & strEKey & strLens
  'Loop

  EncryptLogonDetails = ProcessEncryptString(strEncryptionDetails)

End Function

Public Function DecryptLogonDetails(ByVal strInput As String, ByRef strUserName As String, ByRef strPassword As String _
  , ByRef strDatabase As String, ByRef strServer As String) As String

  Dim strEKey As String
  Dim strLens As String
  Dim lngStart As Long
  Dim lngFinish As Long
  Dim lngDBVersion As Long
  Dim strVersion As String
  Dim intMajor As Integer
  Dim intMinor As Integer

  If strInput = vbNullString Then
    Exit Function
  End If

  ' NPG20081111 Fault 13373
  ' Use the new decryption routine if 3.7 or later...
  strVersion = GetDBVersion
  
  intMajor = CInt(Split(strVersion, ".")(0))
  intMinor = CInt(Split(strVersion, ".")(1))
  If (intMajor = 3 And intMinor >= 7) Or intMajor >= 4 Then
    strInput = ProcessDecryptString(strInput)
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


Public Sub SaveModuleSetting(pstrModuleKey As String, _
  pstrParameterKey As String, _
  pstrParameterType As String, _
  pvarValue As Variant)
  
  With recModuleSetup
    .Index = "idxModuleParameter"
    .Seek "=", pstrModuleKey, pstrParameterKey
    
    If .NoMatch Then
      .AddNew
      !moduleKey = pstrModuleKey
      !parameterkey = pstrParameterKey
    Else
      .Edit
    End If
    
    !ParameterType = pstrParameterType
    !parametervalue = pvarValue
    .Update
  End With

End Sub

Public Function GetModuleSetting(pstrModuleKey As String, _
  pstrParameterKey As String, _
  pvarDefault As Variant) As Variant
  
  Dim fCheckPersThenWorkflow As Boolean
  Dim varResult As Variant
  Dim fFound As Boolean
  
  fFound = False
  varResult = pvarDefault
  
  ' If getting some Workflow settings, read them from Personnel setup
  ' if possible first.
  fCheckPersThenWorkflow = Application.PersonnelModule _
    And ((pstrModuleKey = gsMODULEKEY_PERSONNEL) _
      Or (pstrModuleKey = gsMODULEKEY_WORKFLOW)) _
    And ((pstrParameterKey = gsPARAMETERKEY_PERSONNELTABLE) _
      Or (pstrParameterKey = gsPARAMETERKEY_LOGINNAME) _
      Or (pstrParameterKey = gsPARAMETERKEY_SECONDLOGINNAME))
      
  If fCheckPersThenWorkflow _
    And (pstrModuleKey = gsMODULEKEY_WORKFLOW) Then
    
    pstrModuleKey = gsMODULEKEY_PERSONNEL
  End If
  
  With recModuleSetup
    .Index = "idxModuleParameter"
    .Seek "=", pstrModuleKey, pstrParameterKey
    
    If Not .NoMatch Then
      If Not IsNull(!parametervalue) And Len(!parametervalue) > 0 Then
        varResult = !parametervalue
        fFound = True
      End If
    End If
  End With

  If fCheckPersThenWorkflow _
    And (Not fFound) Then
  
    pstrModuleKey = gsMODULEKEY_WORKFLOW
    
    With recModuleSetup
      .Index = "idxModuleParameter"
      .Seek "=", pstrModuleKey, pstrParameterKey
      
      If Not .NoMatch Then
        If Not IsNull(!parametervalue) And Len(!parametervalue) > 0 Then
          varResult = !parametervalue
        End If
      End If
    End With
  End If
  
  GetModuleSetting = varResult

End Function


Private Function ProcessEncryptString(psString) As String
' NPG20081111 Fault 13373
  On Error GoTo ErrorTrap
  
  Dim sOutput As String
  Dim lngLoop As Long
  Dim iAscCode As Integer
  Dim sChar As String
  Dim sOutputPreProcess  As String
  Dim sSubTemp As String
  
  Const MARKERCHAR_1 = "J"
  Const MARKERCHAR_2 = "P"
  Const MARKERCHAR_3 = "D"
  Const DODGYCHARACTER_INCREMENT_1 = 174
  Const DODGYCHARACTER_INCREMENT_2 = 83
  Const DODGYCHARACTER_INCREMENT_3 = 1
  
  sOutputPreProcess = Replace(psString, MARKERCHAR_1, MARKERCHAR_1 & MARKERCHAR_1)
  sOutputPreProcess = Replace(sOutputPreProcess, MARKERCHAR_2, MARKERCHAR_2 & MARKERCHAR_2)
  sOutputPreProcess = Replace(sOutputPreProcess, MARKERCHAR_3, MARKERCHAR_3 & MARKERCHAR_3)

  sOutput = ""
  lngLoop = 1

  ' Loop through the output replacing dodgy characters with a MARKERCHAR and a safe character offset from the dodgy character.
  ' This is to avoid the dodgy characters messing up the querystring when used in a link to the Workflow website.
  Do While lngLoop <= Len(sOutputPreProcess)
    ' Process the next character.
    sChar = Mid(sOutputPreProcess, lngLoop, 1)
    iAscCode = Asc(sChar)
    
    If (iAscCode <= 32) _
      Or (iAscCode = 34) _
      Or (iAscCode = 35) _
      Or (iAscCode = 37) _
      Or (iAscCode = 60) _
      Or (iAscCode = 62) Then
      
      ' Dodgy character. Must replace with the MARKERCHAR_1 and a different character that we know is OK.
      ' Adding DODGYCHARACTER_INCREMENT_1 on the dodgy character's ASC value causes non-dodgy characters to be used.
      sOutput = sOutput & MARKERCHAR_1 & Chr(iAscCode + DODGYCHARACTER_INCREMENT_1)
      
    ElseIf (iAscCode = 91) _
      Or (iAscCode = 93) _
      Or (iAscCode = 94) _
      Or (iAscCode = 95) _
      Or (iAscCode = 96) _
      Or (iAscCode = 123) _
      Or (iAscCode = 125) _
      Or (iAscCode = 127) _
      Or (iAscCode = 129) _
      Or (iAscCode = 141) _
      Or (iAscCode = 143) _
      Or (iAscCode = 144) _
      Or (iAscCode = 157) _
      Or (iAscCode = 160) Then
      
      ' Dodgy character. Must replace with the MARKERCHAR_2 and a different character that we know is OK.
      ' Adding DODGYCHARACTER_INCREMENT_2 on the dodgy character's ASC value causes non-dodgy characters to be used.
      sOutput = sOutput & MARKERCHAR_2 & Chr(iAscCode + DODGYCHARACTER_INCREMENT_2)
      
    ElseIf (iAscCode = 173) Then
      
      ' Dodgy character. Must replace with the MARKERCHAR_3 and a different character that we know is OK.
      ' Adding DODGYCHARACTER_INCREMENT_3 on the dodgy character's ASC value causes non-dodgy characters to be used.
      sOutput = sOutput & MARKERCHAR_3 & Chr(iAscCode + DODGYCHARACTER_INCREMENT_3)
      
    Else
      ' NOT a dodgy character. Put it straight in the output string with out reprocessing.
      sOutput = sOutput & sChar
    End If

    lngLoop = lngLoop + 1
  Loop
  
  ' Always end with a decent character to avoid training code characters from being chopped.
  ' Use a random character between 65 and 90 (all safe)
  'Randomize
  'sChar = Chr(Int((90 - 65 + 1) * Rnd + 65))
  'sOutput = sOutput & sChar
  
TidyUpAndExit:
  ProcessEncryptString = sOutput
  Exit Function
  
ErrorTrap:
  Resume TidyUpAndExit
  
End Function

Public Function ProcessDecryptString(ByVal psString) As String
  On Error GoTo ErrorTrap
  
  Dim sOutput As String
  Dim lngLoop As Integer
  Dim sOutputPreProcess As String
  Dim sSubTemp As String
  Dim sChar As String
  Dim sNextChar As String
  Dim iJumpChars As Integer
  Dim iAscCode As Integer
  Const MARKERCHAR_1 As String = "J"
  Const MARKERCHAR_2 As String = "P"
  Const MARKERCHAR_3 As String = "D"
  Const DODGYCHARACTER_INCREMENT_1 As Integer = 174
  Const DODGYCHARACTER_INCREMENT_2 As Integer = 83
  Const DODGYCHARACTER_INCREMENT_3 As Integer = 1

  ' NPG20081120 Fault 13420
  If Len(psString) = 14 Then
    ' The encrypted string has no parameters, so exit out before we remove any MARKERCHAR's that
    ' are actually part of the original encrypted lens or ekey.
    ProcessDecryptString = psString
    Exit Function
  End If

  sOutput = vbNullString
  lngLoop = 1
  sOutputPreProcess = psString
  sSubTemp = vbNullString

  ' Loop through the output replacing dodgy characters with a MARKERCHAR and a safe character offset from the dodgy character.
  ' This is to avoid the dodgy characters messing up the querystring when used in a link to the Workflow website.
  Do While lngLoop <= Len(sOutputPreProcess) - 1
    ' Process the next character.

    sChar = Mid(sOutputPreProcess, lngLoop, 1)

    sNextChar = Mid(sOutputPreProcess, lngLoop + 1, 1)

    iJumpChars = 1

    If (sChar = MARKERCHAR_1 Or sChar = MARKERCHAR_2 Or sChar = MARKERCHAR_3) Then
      If sChar <> sNextChar Then
        iAscCode = Asc(sNextChar)

        If sChar = MARKERCHAR_1 Then
          ' Dodgy character marker. Must remove the MARKERCHAR_1 and substract
          ' DODGYCHARACTER_INCREMENT_1 from the dodgy character's ASC value.
          sOutput = sOutput & Chr(iAscCode - DODGYCHARACTER_INCREMENT_1)
        ElseIf sChar = MARKERCHAR_2 Then
          ' Dodgy character marker. Must remove the MARKERCHAR_2 and substract
          ' DODGYCHARACTER_INCREMENT_2 from the dodgy character's ASC value.
          sOutput = sOutput & Chr(iAscCode - DODGYCHARACTER_INCREMENT_2)
        ElseIf sChar = MARKERCHAR_3 Then
          ' Dodgy character marker. Must remove the MARKERCHAR_3 and substract
          ' DODGYCHARACTER_INCREMENT_3 from the dodgy character's ASC value.
          sOutput = sOutput & Chr(iAscCode - DODGYCHARACTER_INCREMENT_3)
        End If
        iJumpChars = 2
      Else
        ' NOT a dodgy character. Put it straight in the output string with out reprocessing.
        sOutput = sOutput & sChar
        iJumpChars = 2
      End If
    Else
      sOutput = sOutput & sChar
      iJumpChars = 1
    End If

    lngLoop = lngLoop + iJumpChars
  Loop

  ' process the last character now
  If lngLoop <= Len(sOutputPreProcess) Then
    sOutput = sOutput + Mid(sOutputPreProcess, lngLoop, 1)
  End If

TidyUpAndExit:
  ProcessDecryptString = sOutput
  Exit Function
  
ErrorTrap:
  Resume TidyUpAndExit

End Function

Private Function GetDBVersion() As String

  Dim rsInfo As New ADODB.Recordset
  
  GetDBVersion = GetSystemSetting("Database", "Version", vbNullString)

  If GetDBVersion = vbNullString Then
    rsInfo.Open "SELECT SystemManagerVersion FROM ASRSysConfig", gADOCon, adOpenForwardOnly, adLockReadOnly
  
    If Not rsInfo.BOF And Not rsInfo.EOF Then
      GetDBVersion = rsInfo.Fields(0).value
    End If
  
    rsInfo.Close
    Set rsInfo = Nothing
  
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

Public Function GetFontDescription(pObjFont As StdFont) As String
  ' Return the description of the given font for display.
  On Error GoTo ErrorTrap
  
  Dim sFontDescription As String
  
  If Not pObjFont Is Nothing Then
    With pObjFont
      sFontDescription = .Name
          
      If .Bold Then
        If .Italic Then
          sFontDescription = sFontDescription & ", Bold Italic"
        Else
          sFontDescription = sFontDescription & ", Bold"
        End If
      Else
        If .Italic Then
          sFontDescription = sFontDescription & ", Italic"
        Else
          sFontDescription = sFontDescription & ", Regular"
        End If
      End If
      
      sFontDescription = sFontDescription & IIf(.Strikethrough, ", Strikethrough", "")
      sFontDescription = sFontDescription & IIf(.Underline, ", Underline", "")
    End With
  Else
    sFontDescription = ""
  End If
  
TidyUpAndExit:
  GetFontDescription = sFontDescription
  Exit Function
  
ErrorTrap:
  sFontDescription = "<unknown>"
  Resume TidyUpAndExit
  
End Function


