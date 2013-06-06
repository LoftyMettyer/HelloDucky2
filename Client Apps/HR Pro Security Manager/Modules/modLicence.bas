Attribute VB_Name = "modLicence"
Option Explicit

Public Enum Module
  modPersonnel = 1
  modRecruitment = 2
  modAbsence = 4
  modTraining = 8
  modIntranet = 16
  modAFD = 32
  modFullSysMgr = 64
  modCMG = 128
  modQAddress = 256
  modAccord = 512
  modWorkflow = 1024
End Enum


Public Function GetLicenceKey()

  'MH 26/07/2001
  'This function will try to get the new licence key
  'If it can't then it will convert the old keys into
  'a single new key and save it away so that next time
  'it can just get the new key from asrsyssystemsettings


  Dim objLicence As COALicence.clsOldLicence
  Dim rsTemp As New ADODB.Recordset
  'Dim rsTemp As Recordset
  Dim strOutput As String


  strOutput = GetSystemSetting("Licence", "Key", vbNullString)

  If Not (strOutput Like "?????-?????-?????-?????") Then

    If Not (strOutput Like "????-????-????-????-????") And _
       Not (strOutput Like "????-????-????-????") Then
      rsTemp.Open "SELECT CustNo, CustName, AuthCode, ModuleCode from ASRSysConfig", gADOCon, adOpenForwardOnly, adLockReadOnly
      If Not rsTemp.BOF And Not rsTemp.EOF Then
  
        'Get Old Licence and convert to new
        Set objLicence = New COALicence.clsOldLicence
        strOutput = objLicence.ConvertOldLicenceToNew2(rsTemp!CustNo, rsTemp!AuthCode, rsTemp!ModuleCode)
        Set objLicence = Nothing
  
        SaveSystemSetting "Licence", "Key", strOutput
        SaveSystemSetting "Licence", "Customer No", rsTemp!CustNo
        SaveSystemSetting "Licence", "Customer Name", rsTemp!CustName

      End If
      rsTemp.Close
      Set rsTemp = Nothing

    Else
      Set objLicence = New COALicence.clsOldLicence
      strOutput = objLicence.ConvertOldLicenceToNew2(GetSystemSetting("Licence", "Customer No", 0), strOutput, "")
      Set objLicence = Nothing
      If Not ASRDEVELOPMENT Then
        SaveSystemSetting "Licence", "Key", strOutput
      End If

    End If

  End If

  GetLicenceKey = strOutput

End Function


Public Function IsModuleEnabled(lngModuleCode As Module) As Boolean

  Dim objLicence As COALicence.clsLicence2

  Set objLicence = New COALicence.clsLicence2
  objLicence.LicenceKey = GetLicenceKey
  IsModuleEnabled = (objLicence.Modules And lngModuleCode)
  Set objLicence = Nothing

End Function


Public Function GetLicencedUsers() As Long

  Dim objLicence As COALicence.clsLicence2

  Set objLicence = New COALicence.clsLicence2
  objLicence.LicenceKey = GetLicenceKey
  GetLicencedUsers = objLicence.DATUsers
  Set objLicence = Nothing

End Function

