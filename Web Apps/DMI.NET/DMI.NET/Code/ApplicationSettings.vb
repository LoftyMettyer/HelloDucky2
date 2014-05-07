﻿Option Strict On
Option Explicit On

Namespace Code
	Public Class ApplicationSettings
		Public Shared ReadOnly Property LoginPage_Database As String
			Get
				Return ConfigurationManager.AppSettings("LoginPage_Database")
			End Get
		End Property
		Public Shared ReadOnly Property LoginPage_Server As String
			Get
				Return ConfigurationManager.AppSettings("LoginPage_Server")
			End Get
		End Property
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
		Public Shared ReadOnly Property AdminRequiresIE As String	'Strictly speaking this property should return a Boolean, but it has been used as a string throughout the application, so String it is
			Get
				Return ConfigurationManager.AppSettings("AdminRequiresIE")
			End Get
		End Property
		Public Shared ReadOnly Property SessionTimeOutInMinutes As String
			Get
				Return ConfigurationManager.AppSettings("SessionTimeOutInMinutes")
			End Get
		End Property
	End Class
End Namespace