Public Class TechSupport
  Implements COMInterfaces.IForm

#Region "COMInterfaces.iForm"

  Public Sub Show1() Implements COMInterfaces.IForm.Show
    Me.Show()
  End Sub

  Public Sub ShowDialog1() Implements COMInterfaces.IForm.ShowDialog
    Me.ShowDialog()
  End Sub

#End Region

  Private Sub TechSupport_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

    Try
      lblTelephone.Text = "Telephone : " & Globals.SystemSettings.Setting("support", "telephone no").Value
      linkEmail.Text = Globals.SystemSettings.Setting("support", "email").Value
      LinkWeb.Text = Globals.SystemSettings.Setting("support", "webpage").Value

      linkEmail.Links.Add(0, linkEmail.Text.Length, linkEmail.Text)
      LinkWeb.Links.Add(0, LinkWeb.Text.Length, LinkWeb.Text)

    Catch ex As Exception

    End Try

  End Sub

  Private Sub LinkWeb_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkWeb.LinkClicked
    System.Diagnostics.Process.Start(e.Link.LinkData.ToString)
  End Sub

  Private Sub linkEmail_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles linkEmail.LinkClicked
    System.Diagnostics.Process.Start("mailto:" & e.Link.LinkData.ToString)
  End Sub
End Class