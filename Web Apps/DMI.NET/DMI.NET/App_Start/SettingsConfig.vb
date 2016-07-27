Option Strict On
Option Explicit On

Imports DMI.NET.Classes
Imports HR.Intranet.Server

' Register the module parameters
Public Class SettingsConfig

	Public Shared Personnel_EmpTableID As Integer
   Public Shared Post_TableID As Integer
   Public Shared Hierarchy_TableID As Integer

   Public Shared Sub Register()

    Try

      Dim objDataAccess As New clsDataAccess(ConfigurationManager.ConnectionStrings("OpenHR").ConnectionString)

         If Licence.IsModuleLicenced(SoftwareModule.Personnel) Then
            Personnel_EmpTableID = CInt(objDataAccess.GetModuleSetting("MODULE_PERSONNEL", "Param_TablePersonnel", "PType_TableID"))
         End If

         Post_TableID = CInt(objDataAccess.GetModuleSetting("MODULE_POST", "Param_PostTable", "PType_TableID"))
         Hierarchy_TableID = CInt(objDataAccess.GetModuleSetting("MODULE_HIERARCHY", "Param_TableHierarchy", "PType_TableID"))

      Catch ex As Exception
      Throw

    End Try

	End Sub  

End Class
