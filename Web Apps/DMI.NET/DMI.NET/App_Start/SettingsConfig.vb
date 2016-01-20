Option Strict On
Option Explicit On

Imports DMI.NET.Classes
Imports HR.Intranet.Server

' Register the module parameters
Public Class SettingsConfig

	Public Shared Personnel_EmpTableID As Integer
  public Shared Post_TableID as Integer
	
	Public Shared Sub Register()

    Try

      Dim objDataAccess As New clsDataAccess(ConfigurationManager.ConnectionStrings("OpenHR").ConnectionString)

 		  If Licence.IsModuleLicenced(SoftwareModule.Personnel) Then
        Personnel_EmpTableID = CInt(objDataAccess.GetModuleSetting("MODULE_PERSONNEL", "Param_TablePersonnel", "PType_TableID"))
		  End If

      Post_TableID = CInt(objDataAccess.GetModuleSetting("MODULE_POST", "Param_PostTable", "PType_TableID"))

    Catch ex As Exception
      Throw

    End Try

	End Sub  

End Class
