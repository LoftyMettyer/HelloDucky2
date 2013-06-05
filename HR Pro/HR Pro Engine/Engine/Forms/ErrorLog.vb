Namespace Forms

  Public Class ErrorLog

    Private Sub ErrorLog_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

      grdErrors.DataSource = Globals.ErrorLog


    End Sub
  End Class

End Namespace
