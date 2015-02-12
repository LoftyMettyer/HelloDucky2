Imports System.Web.Mvc
Imports DMI.NET.ViewModels.Account
Imports DMI.NET.Classes
Imports DMI.NET.Code.Hubs

Namespace Controllers
	Public Class ErrorController
		Inherits Controller

		Public Function PageNotFound() As ActionResult

			Dim objErrors = New ConfigurationErrorsModel

			objErrors.Errors.Add(New ConfigurationError With {.Code = "0001",
																												.Message = "Page not found",
																												.Detail = "This resource has not been found"})

			Return View("PageNotFound", objErrors)

		End Function

		Public Function InternalServer() As ActionResult
			Return View("InternalServer")
		End Function

		Public Function Configuration() As ActionResult

			Dim objErrors = New ConfigurationErrorsModel

			If Not DatabaseHub.HeartbeatOK Then
				objErrors.Errors.Add(New ConfigurationError With {.Code = "0001",
																													.Message = "Database connectivity failure",
																													.Detail = "The IIS server was unable to establish a heartbeat to the OpenHR database."})
			Else

				If Not DatabaseHub.ServiceBrokerOK Then
					objErrors.Errors.Add(New ConfigurationError With {.Code = "0002",
																															.Message = "SQL Service Broker not running",
																															.Detail = "The SQL server is unable to start the service broker."})
				End If

			End If

			Return View("Configuration", objErrors)

		End Function

	End Class
End Namespace