<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Form1
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
    Me.Button1 = New System.Windows.Forms.Button()
    Me.ProgressBar1 = New System.Windows.Forms.ProgressBar()
    Me.ProgressBar2 = New System.Windows.Forms.ProgressBar()
    Me.Label1 = New System.Windows.Forms.Label()
    Me.CurrentPhase = New System.Windows.Forms.Label()
    Me.Button2 = New System.Windows.Forms.Button()
    Me.Button3 = New System.Windows.Forms.Button()
    Me.txtUpdateScript = New System.Windows.Forms.TextBox()
    Me.Button4 = New System.Windows.Forms.Button()
    Me.Button5 = New System.Windows.Forms.Button()
    Me.Button6 = New System.Windows.Forms.Button()
    Me.Button7 = New System.Windows.Forms.Button()
    Me.cmdDatasource = New System.Windows.Forms.Button()
    Me.txtErrors = New System.Windows.Forms.TextBox()
    Me.txtDatabase = New System.Windows.Forms.TextBox()
    Me.Button8 = New System.Windows.Forms.Button()
    Me.txtServer = New System.Windows.Forms.TextBox()
    Me.chkDebugMode = New System.Windows.Forms.CheckBox()
    Me.Label2 = New System.Windows.Forms.Label()
    Me.Label3 = New System.Windows.Forms.Label()
    Me.SuspendLayout()
    '
    'Button1
    '
    Me.Button1.Enabled = False
    Me.Button1.Location = New System.Drawing.Point(24, 31)
    Me.Button1.Name = "Button1"
    Me.Button1.Size = New System.Drawing.Size(166, 38)
    Me.Button1.TabIndex = 0
    Me.Button1.Text = "Calcs & Triggas"
    Me.Button1.UseVisualStyleBackColor = True
    '
    'ProgressBar1
    '
    Me.ProgressBar1.Location = New System.Drawing.Point(32, 118)
    Me.ProgressBar1.Name = "ProgressBar1"
    Me.ProgressBar1.Size = New System.Drawing.Size(530, 23)
    Me.ProgressBar1.TabIndex = 1
    '
    'ProgressBar2
    '
    Me.ProgressBar2.Location = New System.Drawing.Point(32, 148)
    Me.ProgressBar2.Name = "ProgressBar2"
    Me.ProgressBar2.Size = New System.Drawing.Size(530, 23)
    Me.ProgressBar2.TabIndex = 2
    '
    'Label1
    '
    Me.Label1.AutoSize = True
    Me.Label1.Location = New System.Drawing.Point(218, 14)
    Me.Label1.Name = "Label1"
    Me.Label1.Size = New System.Drawing.Size(39, 13)
    Me.Label1.TabIndex = 3
    Me.Label1.Text = "Label1"
    '
    'CurrentPhase
    '
    Me.CurrentPhase.AutoSize = True
    Me.CurrentPhase.Font = New System.Drawing.Font("Microsoft Sans Serif", 24.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
    Me.CurrentPhase.Location = New System.Drawing.Point(568, 118)
    Me.CurrentPhase.Name = "CurrentPhase"
    Me.CurrentPhase.Size = New System.Drawing.Size(134, 37)
    Me.CurrentPhase.TabIndex = 4
    Me.CurrentPhase.Text = "Phase..."
    '
    'Button2
    '
    Me.Button2.Enabled = False
    Me.Button2.Location = New System.Drawing.Point(458, 18)
    Me.Button2.Name = "Button2"
    Me.Button2.Size = New System.Drawing.Size(104, 22)
    Me.Button2.TabIndex = 5
    Me.Button2.Text = "RemoteView On"
    Me.Button2.UseVisualStyleBackColor = True
    '
    'Button3
    '
    Me.Button3.Enabled = False
    Me.Button3.Location = New System.Drawing.Point(221, 46)
    Me.Button3.Name = "Button3"
    Me.Button3.Size = New System.Drawing.Size(121, 23)
    Me.Button3.TabIndex = 6
    Me.Button3.Text = "Read And Spit"
    Me.Button3.UseVisualStyleBackColor = True
    '
    'txtUpdateScript
    '
    Me.txtUpdateScript.Location = New System.Drawing.Point(221, 76)
    Me.txtUpdateScript.Name = "txtUpdateScript"
    Me.txtUpdateScript.Size = New System.Drawing.Size(257, 20)
    Me.txtUpdateScript.TabIndex = 7
    Me.txtUpdateScript.Text = "c:\dev\updatescript\tables.sql"
    '
    'Button4
    '
    Me.Button4.Enabled = False
    Me.Button4.Location = New System.Drawing.Point(458, 46)
    Me.Button4.Name = "Button4"
    Me.Button4.Size = New System.Drawing.Size(104, 23)
    Me.Button4.TabIndex = 8
    Me.Button4.Text = "RemoteView Off"
    Me.Button4.UseVisualStyleBackColor = True
    '
    'Button5
    '
    Me.Button5.Enabled = False
    Me.Button5.Location = New System.Drawing.Point(501, 78)
    Me.Button5.Name = "Button5"
    Me.Button5.Size = New System.Drawing.Size(142, 23)
    Me.Button5.TabIndex = 9
    Me.Button5.Text = "DAO Populate"
    Me.Button5.UseVisualStyleBackColor = True
    '
    'Button6
    '
    Me.Button6.Enabled = False
    Me.Button6.Location = New System.Drawing.Point(694, 71)
    Me.Button6.Name = "Button6"
    Me.Button6.Size = New System.Drawing.Size(87, 29)
    Me.Button6.TabIndex = 10
    Me.Button6.Text = "Button6"
    Me.Button6.UseVisualStyleBackColor = True
    '
    'Button7
    '
    Me.Button7.Location = New System.Drawing.Point(454, 335)
    Me.Button7.Name = "Button7"
    Me.Button7.Size = New System.Drawing.Size(104, 45)
    Me.Button7.TabIndex = 11
    Me.Button7.Text = "Go Script."
    Me.Button7.UseVisualStyleBackColor = True
    '
    'cmdDatasource
    '
    Me.cmdDatasource.Enabled = False
    Me.cmdDatasource.Location = New System.Drawing.Point(684, 18)
    Me.cmdDatasource.Name = "cmdDatasource"
    Me.cmdDatasource.Size = New System.Drawing.Size(75, 23)
    Me.cmdDatasource.TabIndex = 13
    Me.cmdDatasource.Text = "Datasource"
    Me.cmdDatasource.UseVisualStyleBackColor = True
    '
    'txtErrors
    '
    Me.txtErrors.Location = New System.Drawing.Point(12, 177)
    Me.txtErrors.Multiline = True
    Me.txtErrors.Name = "txtErrors"
    Me.txtErrors.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
    Me.txtErrors.Size = New System.Drawing.Size(359, 179)
    Me.txtErrors.TabIndex = 14
    '
    'txtDatabase
    '
    Me.txtDatabase.Location = New System.Drawing.Point(458, 300)
    Me.txtDatabase.Name = "txtDatabase"
    Me.txtDatabase.Size = New System.Drawing.Size(100, 20)
    Me.txtDatabase.TabIndex = 15
    Me.txtDatabase.Text = "recur43"
    '
    'Button8
    '
    Me.Button8.Location = New System.Drawing.Point(728, 251)
    Me.Button8.Name = "Button8"
    Me.Button8.Size = New System.Drawing.Size(75, 23)
    Me.Button8.TabIndex = 16
    Me.Button8.Text = "Button8"
    Me.Button8.UseVisualStyleBackColor = True
    '
    'txtServer
    '
    Me.txtServer.Location = New System.Drawing.Point(458, 274)
    Me.txtServer.Name = "txtServer"
    Me.txtServer.Size = New System.Drawing.Size(100, 20)
    Me.txtServer.TabIndex = 17
    Me.txtServer.Text = "harpdev01"
    '
    'chkDebugMode
    '
    Me.chkDebugMode.AutoSize = True
    Me.chkDebugMode.Checked = True
    Me.chkDebugMode.CheckState = System.Windows.Forms.CheckState.Checked
    Me.chkDebugMode.Location = New System.Drawing.Point(575, 274)
    Me.chkDebugMode.Name = "chkDebugMode"
    Me.chkDebugMode.Size = New System.Drawing.Size(88, 17)
    Me.chkDebugMode.TabIndex = 18
    Me.chkDebugMode.Text = "Debug Mode"
    Me.chkDebugMode.UseVisualStyleBackColor = True
    '
    'Label2
    '
    Me.Label2.AutoSize = True
    Me.Label2.Location = New System.Drawing.Point(399, 277)
    Me.Label2.Name = "Label2"
    Me.Label2.Size = New System.Drawing.Size(38, 13)
    Me.Label2.TabIndex = 19
    Me.Label2.Text = "Server"
    '
    'Label3
    '
    Me.Label3.AutoSize = True
    Me.Label3.Location = New System.Drawing.Point(399, 303)
    Me.Label3.Name = "Label3"
    Me.Label3.Size = New System.Drawing.Size(53, 13)
    Me.Label3.TabIndex = 20
    Me.Label3.Text = "Database"
    '
    'Form1
    '
    Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
    Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
    Me.ClientSize = New System.Drawing.Size(806, 426)
    Me.Controls.Add(Me.Label3)
    Me.Controls.Add(Me.Label2)
    Me.Controls.Add(Me.chkDebugMode)
    Me.Controls.Add(Me.txtServer)
    Me.Controls.Add(Me.Button8)
    Me.Controls.Add(Me.txtDatabase)
    Me.Controls.Add(Me.txtErrors)
    Me.Controls.Add(Me.cmdDatasource)
    Me.Controls.Add(Me.Button7)
    Me.Controls.Add(Me.Button6)
    Me.Controls.Add(Me.Button5)
    Me.Controls.Add(Me.Button4)
    Me.Controls.Add(Me.txtUpdateScript)
    Me.Controls.Add(Me.Button3)
    Me.Controls.Add(Me.Button2)
    Me.Controls.Add(Me.CurrentPhase)
    Me.Controls.Add(Me.Label1)
    Me.Controls.Add(Me.ProgressBar2)
    Me.Controls.Add(Me.ProgressBar1)
    Me.Controls.Add(Me.Button1)
    Me.Name = "Form1"
    Me.Text = "DB"
    Me.ResumeLayout(False)
    Me.PerformLayout()

  End Sub
  Friend WithEvents Button1 As System.Windows.Forms.Button
  Friend WithEvents ProgressBar1 As System.Windows.Forms.ProgressBar
  Friend WithEvents ProgressBar2 As System.Windows.Forms.ProgressBar
  Friend WithEvents Label1 As System.Windows.Forms.Label
  Friend WithEvents CurrentPhase As System.Windows.Forms.Label
  Friend WithEvents Button2 As System.Windows.Forms.Button
  Friend WithEvents Button3 As System.Windows.Forms.Button
  Friend WithEvents txtUpdateScript As System.Windows.Forms.TextBox
  Friend WithEvents Button4 As System.Windows.Forms.Button
  Friend WithEvents Button5 As System.Windows.Forms.Button
  Friend WithEvents Button6 As System.Windows.Forms.Button
  Friend WithEvents Button7 As System.Windows.Forms.Button
  Friend WithEvents cmdDatasource As System.Windows.Forms.Button
  Friend WithEvents txtErrors As System.Windows.Forms.TextBox
  Friend WithEvents txtDatabase As System.Windows.Forms.TextBox
  Friend WithEvents Button8 As System.Windows.Forms.Button
  Friend WithEvents txtServer As System.Windows.Forms.TextBox
  Friend WithEvents chkDebugMode As System.Windows.Forms.CheckBox
  Friend WithEvents Label2 As System.Windows.Forms.Label
  Friend WithEvents Label3 As System.Windows.Forms.Label

End Class
