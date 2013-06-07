Imports System.Windows.Forms
Imports SystemFramework.Enums.Errors

Namespace Forms

  Public Class ErrorLog

    Public Abort As Boolean
    Private Const MlngInitialHeight As Integer = 160
    Private _errorSeverity As Severity = Enums.Errors.Severity.Warning

    Private Sub ErrorLog_Load(ByVal sender As System.Object, ByVal e As EventArgs) Handles MyBase.Load

      lblTelephone.Text = "Telephone : " & SystemSettings.Setting("support", "telephone no").Value
      linkEmail.Text = SystemSettings.Setting("support", "email").Value
      LinkWeb.Text = SystemSettings.Setting("support", "webpage").Value

      linkEmail.Links.Add(0, linkEmail.Text.Length, linkEmail.Text)
      LinkWeb.Links.Add(0, LinkWeb.Text.Length, LinkWeb.Text)

      For Each objError As Structures.Error In Globals.ErrorLog
        Dim objListViewItem = lvwErrors.Items.Add(objError.Message)
        objListViewItem.ImageIndex = objError.Severity
        objListViewItem.SubItems.Add(objError.Detail)

        If objError.Severity = Enums.Errors.Severity.Error Then
          _errorSeverity = Enums.Errors.Severity.Error
        End If
      Next

      Select Case _errorSeverity
        Case Enums.Errors.Severity.Warning
          PictureBox1.Image = imagelist48.Images("warning")
          TextBox2.Text = "Warnings were encountered by the .NET framework. The system can continue saving, but some calculations may not function correctly. Please contact support for assistance."

        Case Else
          PictureBox1.Image = imagelist48.Images("error")
          TextBox2.Text = "Critical errors were encountered by the .NET framework. The system can continue saving, but some calculations may not function correctly and data may be lost. Please contact support for assistance."

      End Select

      lvwErrors.Select()

      Height = MlngInitialHeight

    End Sub

    Private Sub butDetails_Click(ByVal sender As System.Object, ByVal e As EventArgs) Handles butDetails.Click
      txtDetails.Visible = Not txtDetails.Visible

      If txtDetails.Visible Then
        butDetails.Text = "Details <<"
        Height = 515
      Else
        butDetails.Text = "Details >>"
        Height = MlngInitialHeight
      End If

    End Sub

    Private Sub butContinue_Click(ByVal sender As System.Object, ByVal e As EventArgs) Handles butContinue.Click
      If _errorSeverity = Enums.Errors.Severity.Error Then
        If MsgBox("System integrity is compromised. Are you sure you want to continue?", MsgBoxStyle.YesNo Or MsgBoxStyle.Question, "System Framework") = MsgBoxResult.Yes Then
          Close()
        End If
      Else
        Close()
      End If
    End Sub

    Private Sub cmdAbort_Click(ByVal sender As System.Object, ByVal e As EventArgs) Handles cmdAbort.Click
      Abort = True
      Close()
    End Sub

    Private Sub cmdCopy_Click(ByVal sender As System.Object, ByVal e As EventArgs) Handles cmdCopy.Click
      Clipboard.SetText(Globals.ErrorLog.DetailedReport)
    End Sub

    Private Sub LinkWeb_LinkClicked(ByVal sender As System.Object, ByVal e As LinkLabelLinkClickedEventArgs) Handles LinkWeb.LinkClicked
      Process.Start(e.Link.LinkData.ToString)
    End Sub

    Private Sub linkEmail_LinkClicked(ByVal sender As System.Object, ByVal e As LinkLabelLinkClickedEventArgs) Handles linkEmail.LinkClicked

      Dim sMessage As String

      Try

        sMessage = Globals.ErrorLog.DetailedReport.Replace(vbNewLine, "%0d")
        sMessage = sMessage.Replace("""", "'")

        sMessage = String.Format("{0}OpenHR System Framework version : {1}" & _
                        "%0d%0d%0d%0dDetails%0d{2}", vbLf, Version.Major & "." & Version.Minor & "." & Version.Build & "." & Version.Revision, sMessage)

        Process.Start("mailto:" & e.Link.LinkData.ToString & "?Subject=Urgent Support Assistance Required" & "&Body=" & sMessage)

      Catch ex As Exception

      End Try

    End Sub

    Private Sub lvwErrors_SelectedIndexChanged(sender As System.Object, e As EventArgs) Handles lvwErrors.SelectedIndexChanged
      If lvwErrors.SelectedItems.Count > 0 Then
        txtDetails.Text = lvwErrors.SelectedItems(0).SubItems(1).Text
      End If
    End Sub
	End Class

End Namespace
