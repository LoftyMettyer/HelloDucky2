Attribute VB_Name = "modLicence"
Option Explicit

Public Function IsModuleEnabled(lngModuleCode As enum_Module) As Boolean
  IsModuleEnabled = (gobjLicence.Modules And lngModuleCode)
End Function

Public Function GetLicencedUsers() As Long
  GetLicencedUsers = gobjLicence.DATUsers
End Function

