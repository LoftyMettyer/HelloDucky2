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
      Me.components = New System.ComponentModel.Container()
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
      Me.imagelist48 = New System.Windows.Forms.ImageList(Me.components)
      Me.imagelist16 = New System.Windows.Forms.ImageList(Me.components)
      Me.lvwErrors = New System.Windows.Forms.ListView()
      Me.Severity = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
      Me.Message = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
      Me.ID = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
      CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).BeginInit()
      Me.SuspendLayout()
      '
      'txtDetails
      '
      Me.txtDetails.Location = New System.Drawing.Point(428, 132)
      Me.txtDetails.Multiline = True
      Me.txtDetails.Name = "txtDetails"
      Me.txtDetails.ReadOnly = True
      Me.txtDetails.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
      Me.txtDetails.Size = New System.Drawing.Size(279, 311)
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
      Me.TextBox2.Text = "Critical errors were encountered by the .NET framework. The system can continue s" & _
      "aving, but some calculations may not function correctly and data may be lost. Pl" & _
      "ease contact support for assistance."
      '
      'butDetails
      '
      Me.butDetails.Location = New System.Drawing.Point(428, 95)
      Me.butDetails.Name = "butDetails"
      Me.butDetails.Size = New System.Drawing.Size(89, 23)
      Me.butDetails.TabIndex = 3
      Me.butDetails.TabStop = False
      Me.butDetails.Text = "Details >>>"
      Me.butDetails.UseVisualStyleBackColor = True
      '
      'butContinue
      '
      Me.butContinue.Location = New System.Drawing.Point(618, 95)
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
      Me.LinkWeb.TabStop = True
      Me.LinkWeb.Text = "www...."
      '
      'linkEmail
      '
      Me.linkEmail.AutoSize = True
      Me.linkEmail.Location = New System.Drawing.Point(146, 71)
      Me.linkEmail.Name = "linkEmail"
      Me.linkEmail.Size = New System.Drawing.Size(59, 13)
      Me.linkEmail.TabIndex = 10
      Me.linkEmail.TabStop = True
      Me.linkEmail.Text = "mailto:..."
      '
      'cmdAbort
      '
      Me.cmdAbort.Location = New System.Drawing.Point(523, 95)
      Me.cmdAbort.Name = "cmdAbort"
      Me.cmdAbort.Size = New System.Drawing.Size(89, 23)
      Me.cmdAbort.TabIndex = 1
      Me.cmdAbort.Text = "Abort"
      Me.cmdAbort.UseVisualStyleBackColor = True
      '
      'PictureBox1
      '
      Me.PictureBox1.Image = Global.SystemFramework.My.Resources.Resources.Cancel48
      Me.PictureBox1.Location = New System.Drawing.Point(14, 12)
      Me.PictureBox1.Name = "PictureBox1"
      Me.PictureBox1.Size = New System.Drawing.Size(48, 48)
      Me.PictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.AutoSize
      Me.PictureBox1.TabIndex = 3
      Me.PictureBox1.TabStop = False
      '
      'cmdCopy
      '
      Me.cmdCopy.Location = New System.Drawing.Point(16, 462)
      Me.cmdCopy.Name = "cmdCopy"
      Me.cmdCopy.Size = New System.Drawing.Size(128, 23)
      Me.cmdCopy.TabIndex = 11
      Me.cmdCopy.TabStop = False
      Me.cmdCopy.Text = "Copy To Clipboard"
      Me.cmdCopy.UseVisualStyleBackColor = True
      '
      'imagelist48
      '
      Me.imagelist48.ImageStream = CType(resources.GetObject("imagelist48.ImageStream"), System.Windows.Forms.ImageListStreamer)
      Me.imagelist48.TransparentColor = System.Drawing.Color.Transparent
      Me.imagelist48.Images.SetKeyName(0, "Error")
      Me.imagelist48.Images.SetKeyName(1, "Warning")
      '
      'imagelist16
      '
      Me.imagelist16.ImageStream = CType(resources.GetObject("imagelist16.ImageStream"), System.Windows.Forms.ImageListStreamer)
      Me.imagelist16.TransparentColor = System.Drawing.Color.Transparent
      Me.imagelist16.Images.SetKeyName(0, "Error")
      Me.imagelist16.Images.SetKeyName(1, "Warning")
      '
      'lvwErrors
      '
      Me.lvwErrors.AutoArrange = False
      Me.lvwErrors.Columns.AddRange(New System.Windows.Forms.ColumnHeader() {Me.Severity, Me.Message, Me.ID})
      Me.lvwErrors.FullRowSelect = True
      Me.lvwErrors.Location = New System.Drawing.Point(16, 132)
      Me.lvwErrors.Name = "lvwErrors"
      Me.lvwErrors.Size = New System.Drawing.Size(397, 311)
      Me.lvwErrors.SmallImageList = Me.imagelist16
      Me.lvwErrors.TabIndex = 13
      Me.lvwErrors.UseCompatibleStateImageBehavior = False
      Me.lvwErrors.View = System.Windows.Forms.View.Details
      '
      'Severity
      '
      Me.Severity.Text = ""
      Me.Severity.Width = 20
      '
      'Message
      '
      Me.Message.Text = "Message"
      Me.Message.Width = 373
      '
      'ID
      '
      Me.ID.Width = 0
      '
      'ErrorLog
      '
      Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 13.0!)
      Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
      Me.ClientSize = New System.Drawing.Size(716, 497)
      Me.ControlBox = False
      Me.Controls.Add(Me.lvwErrors)
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
    Friend WithEvents imagelist48 As System.Windows.Forms.ImageList
    Friend WithEvents imagelist16 As System.Windows.Forms.ImageList
    Friend WithEvents lvwErrors As System.Windows.Forms.ListView
    Friend WithEvents Severity As System.Windows.Forms.ColumnHeader
    Friend WithEvents Message As System.Windows.Forms.ColumnHeader
    Friend WithEvents ID As System.Windows.Forms.ColumnHeader
  End Class
End Namespace
