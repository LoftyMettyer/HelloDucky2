Imports System.Windows.Forms
Namespace Forms

  Public Class ErrorLog

    Public Abort As Boolean = False
    Private mlngInitialHeight As Long = 166

    Private Sub ErrorLog_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

      '      Dim ErrorBindingSource As New BindingSource
      '     ErrorBindingSource.DataSource = Globals.ErrorLog

      lblTelephone.Text = "Telephone : " & Globals.SystemSettings.Setting("support", "telephone no").Value
      linkEmail.Text = Globals.SystemSettings.Setting("support", "email").Value
      LinkWeb.Text = Globals.SystemSettings.Setting("support", "webpage").Value

      linkEmail.Links.Add(0, linkEmail.Text.Length, linkEmail.Text)
      LinkWeb.Links.Add(0, LinkWeb.Text.Length, LinkWeb.Text)

      txtDetails.Text = Globals.ErrorLog.QuickReport()
      txtDetails.Text = txtDetails.Text & Globals.ErrorLog.DetailedReport

      Me.Height = mlngInitialHeight

    End Sub

    Private Sub butDetails_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles butDetails.Click
      txtDetails.Visible = Not txtDetails.Visible

      If txtDetails.Visible Then
        butDetails.Text = "Details <<<"
        Me.Height = mlngInitialHeight + txtDetails.Height + 10 + cmdCopy.Height
      Else
        butDetails.Text = "Details >>>"
        Me.Height = mlngInitialHeight
      End If

    End Sub

    Private Sub butContinue_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles butContinue.Click
      If MsgBox("System integrity is compromised. Are you sure you want to continue?", MsgBoxStyle.YesNo + MsgBoxStyle.Question, "System Framework") = MsgBoxResult.Yes Then
        Me.Close()
      End If
    End Sub

    Private Sub cmdAbort_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdAbort.Click
      Abort = True
      Me.Close()
    End Sub

    Private Sub cmdCopy_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCopy.Click
      Clipboard.SetText(txtDetails.Text)
    End Sub

    Private Sub LinkWeb_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkWeb.LinkClicked
      System.Diagnostics.Process.Start(e.Link.LinkData.ToString)
    End Sub

    Private Sub linkEmail_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles linkEmail.LinkClicked

      Dim sMessage As String

      Try

        sMessage = txtDetails.Text.Replace(vbNewLine, "%0d")
        sMessage = sMessage.Replace("""", "'")

        sMessage = String.Format("{0}HR Pro System Framework version : {1}" & _
                        "%0d%0d%0d%0dDetails%0d{2}", vbLf, "XXX.XXX", sMessage)

        System.Diagnostics.Process.Start("mailto:" & e.Link.LinkData.ToString & "?Subject=Urgent Support Assistance Required" & "&Body=" & sMessage)

      Catch ex As Exception

      End Try

    End Sub
  End Class

End Namespace
