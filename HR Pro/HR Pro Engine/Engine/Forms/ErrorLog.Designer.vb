Namespace Forms

  <Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
  Partial Class ErrorLog
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
      Try
        If disposing AndAlso components IsNot Nothing Then
          components.Dispose()
        End If
      Finally
        MyBase.Dispose(disposing)
      End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
      Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(ErrorLog))
      Me.txtDetails = New System.Windows.Forms.TextBox()
      Me.TextBox2 = New System.Windows.Forms.TextBox()
      Me.butDetails = New System.Windows.Forms.Button()
      Me.butContinue = New System.Windows.Forms.Button()
      Me.lblTelephone = New System.Windows.Forms.Label()
      Me.lblEmail = New System.Windows.Forms.Label()
      Me.lblWeb = New System.Windows.Forms.Label()
      Me.LinkWeb = New System.Windows.Forms.LinkLabel()
      Me.linkEmail = New System.Windows.Forms.LinkLabel()
      Me.cmdAbort = New System.Windows.Forms.Button()
      Me.PictureBox1 = New System.Windows.Forms.PictureBox()
      Me.cmdCopy = New System.Windows.Forms.Button()
      CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).BeginInit()
      Me.SuspendLayout()
      '
      'txtDetails
      '
      Me.txtDetails.Location = New System.Drawing.Point(16, 130)
      Me.txtDetails.Multiline = True
      Me.txtDetails.Name = "txtDetails"
      Me.txtDetails.ReadOnly = True
      Me.txtDetails.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
      Me.txtDetails.Size = New System.Drawing.Size(707, 351)
      Me.txtDetails.TabIndex = 0
      Me.txtDetails.Visible = False
      '
      'TextBox2
      '
      Me.TextBox2.BackColor = System.Drawing.SystemColors.Control
      Me.TextBox2.BorderStyle = System.Windows.Forms.BorderStyle.None
      Me.TextBox2.Cursor = System.Windows.Forms.Cursors.Default
      Me.TextBox2.Location = New System.Drawing.Point(75, 12)
      Me.TextBox2.Multiline = True
      Me.TextBox2.Name = "TextBox2"
      Me.TextBox2.ReadOnly = True
      Me.TextBox2.Size = New System.Drawing.Size(648, 58)
      Me.TextBox2.TabIndex = 2
      Me.TextBox2.TabStop = False
      Me.TextBox2.Text = resources.GetString("TextBox2.Text")
      '
      'butDetails
      '
      Me.butDetails.Location = New System.Drawing.Point(444, 95)
      Me.butDetails.Name = "butDetails"
      Me.butDetails.Size = New System.Drawing.Size(89, 23)
      Me.butDetails.TabIndex = 3
      Me.butDetails.TabStop = False
      Me.butDetails.Text = "Details >>>"
      Me.butDetails.UseVisualStyleBackColor = True
      '
      'butContinue
      '
      Me.butContinue.Location = New System.Drawing.Point(634, 95)
      Me.butContinue.Name = "butContinue"
      Me.butContinue.Size = New System.Drawing.Size(89, 23)
      Me.butContinue.TabIndex = 2
      Me.butContinue.Text = "Continue"
      Me.butContinue.UseVisualStyleBackColor = True
      '
      'lblTelephone
      '
      Me.lblTelephone.AutoSize = True
      Me.lblTelephone.Location = New System.Drawing.Point(75, 47)
      Me.lblTelephone.Name = "lblTelephone"
      Me.lblTelephone.Size = New System.Drawing.Size(79, 13)
      Me.lblTelephone.TabIndex = 6
      Me.lblTelephone.Text = "Telephone : "
      '
      'lblEmail
      '
      Me.lblEmail.AutoSize = True
      Me.lblEmail.Location = New System.Drawing.Point(75, 71)
      Me.lblEmail.Name = "lblEmail"
      Me.lblEmail.Size = New System.Drawing.Size(47, 13)
      Me.lblEmail.TabIndex = 7
      Me.lblEmail.Text = "Email :"
      '
      'lblWeb
      '
      Me.lblWeb.AutoSize = True
      Me.lblWeb.Location = New System.Drawing.Point(75, 95)
      Me.lblWeb.Name = "lblWeb"
      Me.lblWeb.Size = New System.Drawing.Size(41, 13)
      Me.lblWeb.TabIndex = 8
      Me.lblWeb.Text = "Web :"
      '
      'LinkWeb
      '
      Me.LinkWeb.AutoSize = True
      Me.LinkWeb.Location = New System.Drawing.Point(146, 95)
      Me.LinkWeb.Name = "LinkWeb"
      Me.LinkWeb.Size = New System.Drawing.Size(50, 13)
      Me.LinkWeb.TabIndex = 9
      Me.LinkWeb.Text = "www...."
      '
      'linkEmail
      '
      Me.linkEmail.AutoSize = True
      Me.linkEmail.Location = New System.Drawing.Point(146, 71)
      Me.linkEmail.Name = "linkEmail"
      Me.linkEmail.Size = New System.Drawing.Size(59, 13)
      Me.linkEmail.TabIndex = 10
      Me.linkEmail.Text = "mailto:..."
      '
      'cmdAbort
      '
      Me.cmdAbort.Location = New System.Drawing.Point(539, 95)
      Me.cmdAbort.Name = "cmdAbort"
      Me.cmdAbort.Size = New System.Drawing.Size(89, 23)
      Me.cmdAbort.TabIndex = 1
      Me.cmdAbort.Text = "Abort"
      Me.cmdAbort.UseVisualStyleBackColor = True
      '
      'PictureBox1
      '
      Me.PictureBox1.Image = Global.HRProEngine.My.Resources.Resources.Cancel48
      Me.PictureBox1.Location = New System.Drawing.Point(14, 12)
      Me.PictureBox1.Name = "PictureBox1"
      Me.PictureBox1.Size = New System.Drawing.Size(48, 48)
      Me.PictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.AutoSize
      Me.PictureBox1.TabIndex = 3
      Me.PictureBox1.TabStop = False
      '
      'cmdCopy
      '
      Me.cmdCopy.Location = New System.Drawing.Point(16, 487)
      Me.cmdCopy.Name = "cmdCopy"
      Me.cmdCopy.Size = New System.Drawing.Size(128, 23)
      Me.cmdCopy.TabIndex = 11
      Me.cmdCopy.TabStop = False
      Me.cmdCopy.Text = "Copy To Clipboard"
      Me.cmdCopy.UseVisualStyleBackColor = True
      '
      'ErrorLog
      '
      Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 13.0!)
      Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
      Me.ClientSize = New System.Drawing.Size(732, 513)
      Me.ControlBox = False
      Me.Controls.Add(Me.cmdCopy)
      Me.Controls.Add(Me.cmdAbort)
      Me.Controls.Add(Me.linkEmail)
      Me.Controls.Add(Me.LinkWeb)
      Me.Controls.Add(Me.lblWeb)
      Me.Controls.Add(Me.lblEmail)
      Me.Controls.Add(Me.lblTelephone)
      Me.Controls.Add(Me.butContinue)
      Me.Controls.Add(Me.butDetails)
      Me.Controls.Add(Me.PictureBox1)
      Me.Controls.Add(Me.TextBox2)
      Me.Controls.Add(Me.txtDetails)
      Me.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
      Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
      Me.MaximizeBox = False
      Me.MinimizeBox = False
      Me.Name = "ErrorLog"
      Me.ShowIcon = False
      Me.ShowInTaskbar = False
      Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
      Me.Text = ".NET System Framework"
      Me.TopMost = True
      CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).EndInit()
      Me.ResumeLayout(False)
      Me.PerformLayout()

    End Sub
    Friend WithEvents txtDetails As System.Windows.Forms.TextBox
    Friend WithEvents TextBox2 As System.Windows.Forms.TextBox
    Friend WithEvents PictureBox1 As System.Windows.Forms.PictureBox
    Friend WithEvents butDetails As System.Windows.Forms.Button
    Friend WithEvents butContinue As System.Windows.Forms.Button
    Friend WithEvents lblTelephone As System.Windows.Forms.Label
    Friend WithEvents lblEmail As System.Windows.Forms.Label
    Friend WithEvents lblWeb As System.Windows.Forms.Label
    Friend WithEvents LinkWeb As System.Windows.Forms.LinkLabel
    Friend WithEvents linkEmail As System.Windows.Forms.LinkLabel
    Friend WithEvents cmdAbort As System.Windows.Forms.Button
    Friend WithEvents cmdCopy As System.Windows.Forms.Button
  End Class
End Namespace
