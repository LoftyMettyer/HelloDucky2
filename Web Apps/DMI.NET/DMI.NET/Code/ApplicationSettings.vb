Option Strict On
Option Explicit On

Namespace Code
   Public Class ApplicationSettings
      Public Shared Property LoginPage_Database As String
      Public Shared Property LoginPage_Server As String

      Public Shared ReadOnly Property UI_Admin_Theme As String
         Get
            Return ConfigurationManager.AppSettings("UI_Admin_Theme")
         End Get
      End Property
      Public Shared ReadOnly Property UI_Tiles_Theme As String
         Get
            Return ConfigurationManager.AppSettings("UI_Tiles_Theme")
         End Get
      End Property
      Public Shared ReadOnly Property UI_Wireframe_Theme As String
         Get
            Return ConfigurationManager.AppSettings("UI_Wireframe_Theme")
         End Get
      End Property
      Public Shared ReadOnly Property UI_Winkit_Theme As String
         Get
            Return ConfigurationManager.AppSettings("UI_Winkit_Theme")
         End Get
      End Property
      Public Shared ReadOnly Property UI_Banner_Colour As String
         Get
            Return ConfigurationManager.AppSettings("UI_Banner_Colour")
         End Get
      End Property
      Public Shared ReadOnly Property UI_Banner_Justification As String
         Get
            Return ConfigurationManager.AppSettings("UI_Banner_Justification")
         End Get
      End Property
      Public Shared ReadOnly Property UI_Self_Service_Layout As String
         Get
            If ConfigurationManager.AppSettings("UI_Self_Service_Layout") Is Nothing Then
               Return "winkit"
            ElseIf (ConfigurationManager.AppSettings("UI_Self_Service_Layout").ToUpper() = "WINKIT" Or ConfigurationManager.AppSettings("UI_Self_Service_Layout").ToUpper() = "WIREFRAME" Or ConfigurationManager.AppSettings("UI_Self_Service_Layout").ToUpper() = "TILES".ToUpper()) Then
               Return ConfigurationManager.AppSettings("UI_Self_Service_Layout").ToLower()
            Else
               Return "winkit"
            End If
         End Get
      End Property
      Public Shared ReadOnly Property UI_Layout_Selectable As String 'Strictly speaking this property should return a Boolean, but since everything else is boolean, this is too.
         Get
            Return ConfigurationManager.AppSettings("UI_Layout_Selectable")
         End Get
      End Property
      Public Shared ReadOnly Property ValidFileExtensions As String
         Get
            Return ConfigurationManager.AppSettings("ValidFileExtensions")
         End Get
      End Property

      Public Shared ReadOnly Property EnableViewCurrentUsers As Boolean
         Get
            Dim setting = ConfigurationManager.AppSettings("EnableViewCurrentUsers")

            If setting IsNot Nothing Then
               Return CBool(setting)
            End If

            Return False
         End Get
      End Property

      Public Shared ReadOnly Property OpenAmAuthenticateUri As String = ConfigurationManager.AppSettings("OpenAmAuthenticateUri")
      Public Shared ReadOnly Property OpenAmGetIdFromSessionUri As String = ConfigurationManager.AppSettings("OpenAmGetIdFromSessionUri")

   End Class
End Namespace