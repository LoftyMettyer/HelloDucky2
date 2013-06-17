
Partial Class Messages
  Inherits System.Web.UI.Page

  Protected Sub Page_Load(sender As Object, e As System.EventArgs) Handles Me.Load

    MessageLabel.Text = CStr(Session("messages"))

  End Sub

End Class
