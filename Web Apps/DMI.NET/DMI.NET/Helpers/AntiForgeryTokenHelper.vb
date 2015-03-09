Option Strict On
Option Explicit On

Imports System.Runtime.CompilerServices

Namespace Helpers

	Public Module AntiForgeryTokenHelper

		<Extension> _
		Public Function AntiForgeryTokenForAjaxPost(helper As HtmlHelper) As MvcHtmlString
			Dim antiForgeryInputTag = helper.AntiForgeryToken().ToString()
			' Above gets the following: <input name="__RequestVerificationToken" type="hidden" value="PnQE7R0MIBBAzC7SqtVvwrJpGbRvPgzWHo5dSyoSaZoabRjf9pCyzjujYBU_qKDJmwIOiPRDwBV1TNVdXFVgzAvN9_l2yt9-nf4Owif0qIDz7WRAmydVPIm6_pmJAI--wvvFQO7g0VvoFArFtAR2v6Ch1wmXCZ89v0-lNOGZLZc1" />
			Dim removedStart = antiForgeryInputTag.Replace("<input name=""__RequestVerificationToken"" type=""hidden"" value=""", "")
			Dim tokenValue = removedStart.Replace(""" />", "")
			If antiForgeryInputTag = removedStart OrElse removedStart = tokenValue Then
				Throw New InvalidOperationException("Oops! The Html.AntiForgeryToken() method seems to return something I did not expect.")
			End If
			Return New MvcHtmlString(String.Format("{0}:""{1}""", "__RequestVerificationToken", tokenValue))
		End Function

	End Module

End Namespace
