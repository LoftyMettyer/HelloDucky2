<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class DataSource
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
    Me.cboProvider = New System.Windows.Forms.ComboBox()
    Me.TextBox1 = New System.Windows.Forms.TextBox()
    Me.Label1 = New System.Windows.Forms.Label()
    Me.lblProvider = New System.Windows.Forms.Label()
    Me.Label2 = New System.Windows.Forms.Label()
    Me.txtProviderString = New System.Windows.Forms.TextBox()
    Me.lblDatabase = New System.Windows.Forms.Label()
    Me.lblLoginType = New System.Windows.Forms.Label()
    Me.cboLoginType = New System.Windows.Forms.ComboBox()
    Me.txtDatabase = New System.Windows.Forms.TextBox()
    Me.lblLoginName = New System.Windows.Forms.Label()
    Me.lblLoginPassword = New System.Windows.Forms.Label()
    Me.txtLoginName = New System.Windows.Forms.TextBox()
    Me.txtLoginPassword = New System.Windows.Forms.TextBox()
    Me.SuspendLayout()
    '
    'cboProvider
    '
    Me.cboProvider.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
    Me.cboProvider.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
    Me.cboProvider.FormattingEnabled = True
    Me.cboProvider.Items.AddRange(New Object() {"SQL", "Oracle", "Progress", "Access", "FoxPro", "Excel"})
    Me.cboProvider.Location = New System.Drawing.Point(121, 47)
    Me.cboProvider.Name = "cboProvider"
    Me.cboProvider.Size = New System.Drawing.Size(218, 21)
    Me.cboProvider.TabIndex = 0
    '
    'TextBox1
    '
    Me.TextBox1.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
    Me.TextBox1.Location = New System.Drawing.Point(121, 20)
    Me.TextBox1.Name = "TextBox1"
    Me.TextBox1.Size = New System.Drawing.Size(218, 21)
    Me.TextBox1.TabIndex = 1
    '
    'Label1
    '
    Me.Label1.AutoSize = True
    Me.Label1.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
    Me.Label1.Location = New System.Drawing.Point(13, 28)
    Me.Label1.Name = "Label1"
    Me.Label1.Size = New System.Drawing.Size(49, 13)
    Me.Label1.TabIndex = 2
    Me.Label1.Text = "Name :"
    '
    'lblProvider
    '
    Me.lblProvider.AutoSize = True
    Me.lblProvider.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
    Me.lblProvider.Location = New System.Drawing.Point(13, 56)
    Me.lblProvider.Name = "lblProvider"
    Me.lblProvider.Size = New System.Drawing.Size(64, 13)
    Me.lblProvider.TabIndex = 3
    Me.lblProvider.Text = "Provider :"
    '
    'Label2
    '
    Me.Label2.AutoSize = True
    Me.Label2.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
    Me.Label2.Location = New System.Drawing.Point(16, 86)
    Me.Label2.Name = "Label2"
    Me.Label2.Size = New System.Drawing.Size(106, 13)
    Me.Label2.TabIndex = 4
    Me.Label2.Text = "Provider String : "
    '
    'txtProviderString
    '
    Me.txtProviderString.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
    Me.txtProviderString.Location = New System.Drawing.Point(146, 79)
    Me.txtProviderString.Name = "txtProviderString"
    Me.txtProviderString.Size = New System.Drawing.Size(218, 21)
    Me.txtProviderString.TabIndex = 5
    '
    'lblDatabase
    '
    Me.lblDatabase.AutoSize = True
    Me.lblDatabase.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
    Me.lblDatabase.Location = New System.Drawing.Point(16, 118)
    Me.lblDatabase.Name = "lblDatabase"
    Me.lblDatabase.Size = New System.Drawing.Size(70, 13)
    Me.lblDatabase.TabIndex = 6
    Me.lblDatabase.Text = "Database :"
    '
    'lblLoginType
    '
    Me.lblLoginType.AutoSize = True
    Me.lblLoginType.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
    Me.lblLoginType.Location = New System.Drawing.Point(19, 164)
    Me.lblLoginType.Name = "lblLoginType"
    Me.lblLoginType.Size = New System.Drawing.Size(69, 13)
    Me.lblLoginType.TabIndex = 7
    Me.lblLoginType.Text = "Login Type"
    '
    'cboLoginType
    '
    Me.cboLoginType.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
    Me.cboLoginType.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
    Me.cboLoginType.FormattingEnabled = True
    Me.cboLoginType.Items.AddRange(New Object() {"Pass Through", "Static"})
    Me.cboLoginType.Location = New System.Drawing.Point(137, 155)
    Me.cboLoginType.Name = "cboLoginType"
    Me.cboLoginType.Size = New System.Drawing.Size(215, 21)
    Me.cboLoginType.TabIndex = 8
    '
    'txtDatabase
    '
    Me.txtDatabase.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
    Me.txtDatabase.Location = New System.Drawing.Point(143, 113)
    Me.txtDatabase.Name = "txtDatabase"
    Me.txtDatabase.Size = New System.Drawing.Size(220, 21)
    Me.txtDatabase.TabIndex = 9
    '
    'lblLoginName
    '
    Me.lblLoginName.AutoSize = True
    Me.lblLoginName.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
    Me.lblLoginName.Location = New System.Drawing.Point(39, 212)
    Me.lblLoginName.Name = "lblLoginName"
    Me.lblLoginName.Size = New System.Drawing.Size(83, 13)
    Me.lblLoginName.TabIndex = 10
    Me.lblLoginName.Text = "Login Name :"
    '
    'lblLoginPassword
    '
    Me.lblLoginPassword.AutoSize = True
    Me.lblLoginPassword.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
    Me.lblLoginPassword.Location = New System.Drawing.Point(42, 240)
    Me.lblLoginPassword.Name = "lblLoginPassword"
    Me.lblLoginPassword.Size = New System.Drawing.Size(104, 13)
    Me.lblLoginPassword.TabIndex = 11
    Me.lblLoginPassword.Text = "Login Password :"
    '
    'txtLoginName
    '
    Me.txtLoginName.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
    Me.txtLoginName.Location = New System.Drawing.Point(172, 209)
    Me.txtLoginName.Name = "txtLoginName"
    Me.txtLoginName.Size = New System.Drawing.Size(167, 21)
    Me.txtLoginName.TabIndex = 12
    '
    'txtLoginPassword
    '
    Me.txtLoginPassword.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
    Me.txtLoginPassword.Location = New System.Drawing.Point(175, 236)
    Me.txtLoginPassword.Name = "txtLoginPassword"
    Me.txtLoginPassword.Size = New System.Drawing.Size(167, 21)
    Me.txtLoginPassword.TabIndex = 13
    '
    'DataSource
    '
    Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
    Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
    Me.ClientSize = New System.Drawing.Size(768, 493)
    Me.Controls.Add(Me.txtLoginPassword)
    Me.Controls.Add(Me.txtLoginName)
    Me.Controls.Add(Me.lblLoginPassword)
    Me.Controls.Add(Me.lblLoginName)
    Me.Controls.Add(Me.txtDatabase)
    Me.Controls.Add(Me.cboLoginType)
    Me.Controls.Add(Me.lblLoginType)
    Me.Controls.Add(Me.lblDatabase)
    Me.Controls.Add(Me.txtProviderString)
    Me.Controls.Add(Me.Label2)
    Me.Controls.Add(Me.lblProvider)
    Me.Controls.Add(Me.Label1)
    Me.Controls.Add(Me.TextBox1)
    Me.Controls.Add(Me.cboProvider)
    Me.Name = "DataSource"
    Me.Text = "Data Source"
    Me.ResumeLayout(False)
    Me.PerformLayout()

  End Sub
  Friend WithEvents cboProvider As System.Windows.Forms.ComboBox
  Friend WithEvents TextBox1 As System.Windows.Forms.TextBox
  Friend WithEvents Label1 As System.Windows.Forms.Label
  Friend WithEvents lblProvider As System.Windows.Forms.Label
  Friend WithEvents Label2 As System.Windows.Forms.Label
  Friend WithEvents txtProviderString As System.Windows.Forms.TextBox
  Friend WithEvents lblDatabase As System.Windows.Forms.Label
  Friend WithEvents lblLoginType As System.Windows.Forms.Label
  Friend WithEvents cboLoginType As System.Windows.Forms.ComboBox
  Friend WithEvents txtDatabase As System.Windows.Forms.TextBox
  Friend WithEvents lblLoginName As System.Windows.Forms.Label
  Friend WithEvents lblLoginPassword As System.Windows.Forms.Label
  Friend WithEvents txtLoginName As System.Windows.Forms.TextBox
  Friend WithEvents txtLoginPassword As System.Windows.Forms.TextBox
End Class
