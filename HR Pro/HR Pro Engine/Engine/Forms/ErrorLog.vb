Imports System.Windows.Forms
Namespace Forms

  Public Class ErrorLog

    Public Abort As Boolean
    Private mlngInitialHeight As Integer = 160
    Private ErrorSeverity As ErrorHandler.Severity = ErrorHandler.Severity.Warning

    Private Sub ErrorLog_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

      Dim objListViewItem As ListViewItem

      lblTelephone.Text = "Telephone : " & Globals.SystemSettings.Setting("support", "telephone no").Value
      linkEmail.Text = Globals.SystemSettings.Setting("support", "email").Value
      LinkWeb.Text = Globals.SystemSettings.Setting("support", "webpage").Value

      linkEmail.Links.Add(0, linkEmail.Text.Length, linkEmail.Text)
      LinkWeb.Links.Add(0, LinkWeb.Text.Length, LinkWeb.Text)

      'txtDetails.Text = Globals.ErrorLog.QuickReport()
      '      txtDetails.Text = txtDetails.Text & Globals.ErrorLog.DetailedReport

      For Each objError As ErrorHandler.Error In Globals.ErrorLog
				objListViewItem = lvwErrors.Items.Add("")
        objListViewItem.ImageIndex = objError.Severity
        objListViewItem.SubItems.Add(objError.Message)
        objListViewItem.SubItems.Add(objError.Detail)

        If objError.Severity = ErrorHandler.Severity.Error Then
          ErrorSeverity = ErrorHandler.Severity.Error
        End If
      Next

      Select Case ErrorSeverity
        Case ErrorHandler.Severity.Warning
          PictureBox1.Image = imagelist48.Images("warning")
          TextBox2.Text = "Warnings were encountered by the .NET framework. The system can continue saving, but some calculations may not function correctly. Please contact support for assistance."

        Case Else
          PictureBox1.Image = imagelist48.Images("error")
          TextBox2.Text = "Critical errors were encountered by the .NET framework. The system can continue saving, but some calculations may not function correctly and data may be lost. Please contact support for assistance."

      End Select

      lvwErrors.Select()

      Me.Height = mlngInitialHeight

    End Sub

    Private Sub butDetails_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles butDetails.Click
      txtDetails.Visible = Not txtDetails.Visible

      If txtDetails.Visible Then
        butDetails.Text = "Details <<<"
        Me.Height = mlngInitialHeight + txtDetails.Height + 20 + cmdCopy.Height
      Else
        butDetails.Text = "Details >>>"
        Me.Height = mlngInitialHeight
      End If

    End Sub

    Private Sub butContinue_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles butContinue.Click
      If ErrorSeverity = ErrorHandler.Severity.Error Then
        If MsgBox("System integrity is compromised. Are you sure you want to continue?", MsgBoxStyle.YesNo Or MsgBoxStyle.Question, "System Framework") = MsgBoxResult.Yes Then
          Me.Close()
        End If
      Else
        Me.Close()
      End If
    End Sub

    Private Sub cmdAbort_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdAbort.Click
      Abort = True
      Me.Close()
    End Sub

    Private Sub cmdCopy_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCopy.Click
      Clipboard.SetText(Globals.ErrorLog.DetailedReport)
    End Sub

    Private Sub LinkWeb_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkWeb.LinkClicked
      System.Diagnostics.Process.Start(e.Link.LinkData.ToString)
    End Sub

    Private Sub linkEmail_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles linkEmail.LinkClicked

      Dim sMessage As String

      Try

        sMessage = Globals.ErrorLog.DetailedReport.Replace(vbNewLine, "%0d")
        sMessage = sMessage.Replace("""", "'")

        sMessage = String.Format("{0}OpenHR System Framework version : {1}" & _
                        "%0d%0d%0d%0dDetails%0d{2}", vbLf, Version.Major & "." & Version.Minor & "." & Version.Build & "." & Version.Revision, sMessage)

        System.Diagnostics.Process.Start("mailto:" & e.Link.LinkData.ToString & "?Subject=Urgent Support Assistance Required" & "&Body=" & sMessage)

      Catch ex As Exception

      End Try

    End Sub

    Private Sub lvwErrors_Click(sender As Object, e As System.EventArgs) Handles lvwErrors.Click
      txtDetails.Text = lvwErrors.SelectedItems(0).SubItems(2).Text
    End Sub

  End Class

End Namespace
