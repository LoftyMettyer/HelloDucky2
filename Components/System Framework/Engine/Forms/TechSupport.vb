Public Class TechSupport
  Implements IForm

#Region "COMInterfaces.iForm"

  Public Sub Show1() Implements IForm.Show
    Show()
  End Sub

  Public Sub ShowDialog1() Implements IForm.ShowDialog
    ShowDialog()
  End Sub

#End Region

  Private Sub TechSupport_Load(ByVal sender As System.Object, ByVal e As EventArgs) Handles MyBase.Load

    Try
      lblTelephone.Text = "Telephone : " & SystemSettings.Setting("support", "telephone no").Value
      linkEmail.Text = SystemSettings.Setting("support", "email").Value
      LinkWeb.Text = SystemSettings.Setting("support", "webpage").Value

      linkEmail.Links.Add(0, linkEmail.Text.Length, linkEmail.Text)
      LinkWeb.Links.Add(0, LinkWeb.Text.Length, LinkWeb.Text)

    Catch ex As Exception

    End Try

  End Sub

  Private Sub LinkWeb_LinkClicked(ByVal sender As System.Object, ByVal e As Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkWeb.LinkClicked
    Process.Start(e.Link.LinkData.ToString)
  End Sub

  Private Sub linkEmail_LinkClicked(ByVal sender As System.Object, ByVal e As Windows.Forms.LinkLabelLinkClickedEventArgs) Handles linkEmail.LinkClicked
    Process.Start("mailto:" & e.Link.LinkData.ToString)
  End Sub
End Class