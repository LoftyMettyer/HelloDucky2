<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class TechSupport
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
    Me.butOK = New System.Windows.Forms.Button()
    Me.linkEmail = New System.Windows.Forms.LinkLabel()
    Me.LinkWeb = New System.Windows.Forms.LinkLabel()
    Me.lblWeb = New System.Windows.Forms.Label()
    Me.lblEmail = New System.Windows.Forms.Label()
    Me.lblTelephone = New System.Windows.Forms.Label()
    Me.Panel1 = New System.Windows.Forms.Panel()
    Me.Panel1.SuspendLayout()
    Me.SuspendLayout()
    '
    'butOK
    '
    Me.butOK.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
    Me.butOK.Location = New System.Drawing.Point(357, 150)
    Me.butOK.Name = "butOK"
    Me.butOK.Size = New System.Drawing.Size(90, 28)
    Me.butOK.TabIndex = 0
    Me.butOK.Text = "OK"
    Me.butOK.UseVisualStyleBackColor = True
    '
    'linkEmail
    '
    Me.linkEmail.AutoSize = True
    Me.linkEmail.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
    Me.linkEmail.Location = New System.Drawing.Point(86, 40)
    Me.linkEmail.Name = "linkEmail"
    Me.linkEmail.Size = New System.Drawing.Size(59, 13)
    Me.linkEmail.TabIndex = 15
    Me.linkEmail.Text = "mailto:..."
    '
    'LinkWeb
    '
    Me.LinkWeb.AutoSize = True
    Me.LinkWeb.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
    Me.LinkWeb.Location = New System.Drawing.Point(86, 64)
    Me.LinkWeb.Name = "LinkWeb"
    Me.LinkWeb.Size = New System.Drawing.Size(50, 13)
    Me.LinkWeb.TabIndex = 14
    Me.LinkWeb.Text = "www...."
    '
    'lblWeb
    '
    Me.lblWeb.AutoSize = True
    Me.lblWeb.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
    Me.lblWeb.Location = New System.Drawing.Point(15, 64)
    Me.lblWeb.Name = "lblWeb"
    Me.lblWeb.Size = New System.Drawing.Size(41, 13)
    Me.lblWeb.TabIndex = 13
    Me.lblWeb.Text = "Web :"
    '
    'lblEmail
    '
    Me.lblEmail.AutoSize = True
    Me.lblEmail.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
    Me.lblEmail.Location = New System.Drawing.Point(15, 40)
    Me.lblEmail.Name = "lblEmail"
    Me.lblEmail.Size = New System.Drawing.Size(47, 13)
    Me.lblEmail.TabIndex = 12
    Me.lblEmail.Text = "Email :"
    '
    'lblTelephone
    '
    Me.lblTelephone.AutoSize = True
    Me.lblTelephone.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
    Me.lblTelephone.Location = New System.Drawing.Point(15, 16)
    Me.lblTelephone.Name = "lblTelephone"
    Me.lblTelephone.Size = New System.Drawing.Size(79, 13)
    Me.lblTelephone.TabIndex = 11
    Me.lblTelephone.Text = "Telephone : "
    '
    'Panel1
    '
    Me.Panel1.Controls.Add(Me.lblTelephone)
    Me.Panel1.Controls.Add(Me.linkEmail)
    Me.Panel1.Controls.Add(Me.lblEmail)
    Me.Panel1.Controls.Add(Me.LinkWeb)
    Me.Panel1.Controls.Add(Me.lblWeb)
    Me.Panel1.Location = New System.Drawing.Point(12, 12)
    Me.Panel1.Name = "Panel1"
    Me.Panel1.Size = New System.Drawing.Size(435, 132)
    Me.Panel1.TabIndex = 16
    '
    'TechSupport
    '
    Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
    Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
    Me.ClientSize = New System.Drawing.Size(459, 190)
    Me.Controls.Add(Me.Panel1)
    Me.Controls.Add(Me.butOK)
    Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
    Me.MaximizeBox = False
    Me.MinimizeBox = False
    Me.Name = "TechSupport"
    Me.ShowIcon = False
    Me.ShowInTaskbar = False
    Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
    Me.Text = "Support Contact Details"
    Me.Panel1.ResumeLayout(False)
    Me.Panel1.PerformLayout()
    Me.ResumeLayout(False)

  End Sub
  Friend WithEvents butOK As System.Windows.Forms.Button
  Friend WithEvents linkEmail As System.Windows.Forms.LinkLabel
  Friend WithEvents LinkWeb As System.Windows.Forms.LinkLabel
  Friend WithEvents lblWeb As System.Windows.Forms.Label
  Friend WithEvents lblEmail As System.Windows.Forms.Label
  Friend WithEvents lblTelephone As System.Windows.Forms.Label
  Friend WithEvents Panel1 As System.Windows.Forms.Panel
End Class
