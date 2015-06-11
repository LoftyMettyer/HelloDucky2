Imports System.Web.Mvc

Namespace Controllers
	Public Class ApiController
		Inherits Controller

		' GET: Api
		Function Index() As ActionResult

			If Session("sessionCurrentUser") Is Nothing Then Return New HttpStatusCodeResult(401)

			Return Json("AUTHORIZED", JsonRequestBehavior.AllowGet)

		End Function
	End Class
End Namespace