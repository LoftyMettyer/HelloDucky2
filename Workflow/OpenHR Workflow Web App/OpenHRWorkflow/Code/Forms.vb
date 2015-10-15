Imports System.Data.SqlClient

Public Class Forms
	Public Shared Sub RedirectIfNotLicensed()

		Dim db As New Database(App.Config.ConnectionString)
		If Not db.IsMobileModuleLicensed() Then
			HttpContext.Current.Session("message") = "You are not licensed for the OpenHR Mobile module. Please contact your Advanced Business Solutions Account Manager for details"
			HttpContext.Current.Server.Transfer("~/Message.aspx")
		End If

	End Sub

	Public Shared Sub RedirectIfDbLocked()

		Dim db As New Database(App.Config.ConnectionString)
		If db.IsSystemLocked() Then
			HttpContext.Current.Session("message") = "The system is currently being modified. Please retry again shortly."
			HttpContext.Current.Server.Transfer("~/Message.aspx")
		End If

	End Sub


    Public Shared Sub RedirectToNotConfigured()

		Dim errors As New List(Of String)

		If App.Config.WorkflowUrl.IsNullOrEmpty Then errors.Add("The Workflow Url has not been configured.")

		Dim db As New Database(App.Config.ConnectionString)

		If Not db.CanConnect() Then
			errors.Add("Unable to connect to the OpenHR database<BR><BR>Please contact your system administrator. (Error Code: CE001).")
		Else
            If db.IsUserProhibited() Then
                errors.Add("Unable to connect to the OpenHR database<BR><BR>Please contact your system administrator. (Error Code: CE002).")
            ElseIf Not db.ServiceLoginIsValid() Then
                errors.Add("Unable to connect to the OpenHR database<BR><BR>Please contact your system administrator. (Error Code: CE004).")
            ElseIf Not db.IsMobileModuleLicensed() Then
                errors.Add("You are not licensed for the OpenHR Mobile module. Please contact your Advanced Business Solutions Account Manager for details")
			ElseIf Not db.IsIntranetFunctionInstalled() Then
				errors.Add("The database is out of date, re-run the latest intranet update script.")
			End If
		End If

		If errors.Count > 0 Then
			HttpContext.Current.Session("message") = "The system is not configured correctly for the following reasons:<BR><BR>" & String.Join("<BR>", errors)
			HttpContext.Current.Server.Transfer("~/Message.aspx")
		End If
	End Sub

	Public Shared Sub RedirectToMobileModuleNotInstalled()

		Dim errors As New List(Of String)

		Dim db As New Database(App.Config.ConnectionString)

		If Not db.IsMobileModuleInstalled() Then
			errors.Add("The mobile module is not configured correctly in system manager.")
		End If

		If errors.Count > 0 Then
			HttpContext.Current.Session("message") = "The system is not configured correctly for the following reasons:<BR><BR>" & String.Join("<BR>", errors)
			HttpContext.Current.Server.Transfer("~/Message.aspx")
		End If
	End Sub

	Public Shared Sub RedirectToHomeIfAuthentcated()

		'Go to the home page if already logged in
		If HttpContext.Current.Request.IsAuthenticated Then
			HttpContext.Current.Response.Redirect("~/Home.aspx")
		End If

	End Sub

	Public Shared Sub LoadControlData(page As Page, formId As Integer)

		Using conn As New SqlConnection(App.Config.ConnectionString)

			conn.Open()

			Dim cmd As New SqlCommand("SELECT * FROM tbsys_mobileformelements WHERE form = " & formId, conn)
			Dim dr As SqlDataReader = cmd.ExecuteReader()

			While dr.Read()

				Dim control As Control = page.Master.FindControl("mainCPH").FindControl(CStr(dr("Name")))
				If control Is Nothing Then control = page.Master.FindControl("footerCPH").FindControl(CStr(dr("Name")))

				Select Case CInt(dr("Type"))

					Case 0 ' Button

						CType(control.Controls(0), Image).ImageUrl = "~/Image.ashx?id=" & NullSafeInteger(dr("PictureID"))
						CType(control.Controls(1), Label).Text = NullSafeString(dr("caption"))

					Case 2 ' Label

						With CType(control, Label)
							.Text = NullSafeString(dr("caption"))
							.Font.Name = NullSafeString(dr("FontName"))
							.Font.Size = New FontUnit(NullSafeSingle(dr("FontSize")))
							.Font.Bold = NullSafeBoolean(dr("FontBold"))
							.Font.Italic = NullSafeBoolean(dr("FontItalic"))
							.Font.Underline = NullSafeBoolean(dr("FontUnderline"))
							.Font.Strikeout = NullSafeBoolean(dr("FontStrikeout"))

							.Style.Add("color", General.GetHtmlColour(NullSafeInteger(dr("ForeColor"))))
							.Style("word-wrap") = "break-word"
						End With

					Case 3 ' Input value - character

						With CType(control, TextBox)
							.Font.Name = NullSafeString(dr("FontName"))
							.Font.Size = New FontUnit(NullSafeSingle(dr("FontSize")))
							.Font.Bold = NullSafeBoolean(dr("FontBold"))
							.Font.Italic = NullSafeBoolean(dr("FontItalic"))
							.Font.Underline = NullSafeBoolean(dr("FontUnderline"))
							.Font.Strikeout = NullSafeBoolean(dr("FontStrikeout"))

							.Style.Add("color", General.GetHtmlColour(NullSafeInteger(dr("ForeColor"))))
						End With

				End Select
			End While

		End Using

	End Sub

End Class

Public Class FontSetting
	Public Name As String
	Public Size As Single
	Public Bold As Boolean
	Public Italic As Boolean
	Public Underline As Boolean
	Public Strikeout As Boolean
End Class