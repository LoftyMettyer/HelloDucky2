Namespace DMI.NET

	<SessionState(SessionStateBehavior.ReadOnly)>
 Public Class SessionLessController
		Inherits System.Web.Mvc.Controller

		'
		' GET: /SessionLess

		Function Test(index As Integer) As JsonResult
			Dim con = Session("databaseConnection")
			System.Threading.Thread.Sleep(1000)
			Return Json(index, JsonRequestBehavior.AllowGet)
		End Function

	End Class
End Namespace
