<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class ImportObjects
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
    Me.pnlMappings = New System.Windows.Forms.Panel()
    Me.Button1 = New System.Windows.Forms.Button()
    Me.SuspendLayout()
    '
    'pnlMappings
    '
    Me.pnlMappings.AutoScroll = True
    Me.pnlMappings.Location = New System.Drawing.Point(12, 12)
    Me.pnlMappings.Name = "pnlMappings"
    Me.pnlMappings.Size = New System.Drawing.Size(471, 396)
    Me.pnlMappings.TabIndex = 13
    '
    'Button1
    '
    Me.Button1.Location = New System.Drawing.Point(500, 12)
    Me.Button1.Name = "Button1"
    Me.Button1.Size = New System.Drawing.Size(111, 23)
    Me.Button1.TabIndex = 14
    Me.Button1.Text = "Load From File"
    Me.Button1.UseVisualStyleBackColor = True
    '
    'ImportObjects
    '
    Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
    Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
    Me.ClientSize = New System.Drawing.Size(716, 431)
    Me.Controls.Add(Me.Button1)
    Me.Controls.Add(Me.pnlMappings)
    Me.Name = "ImportObjects"
    Me.Text = "ImportObjects"
    Me.ResumeLayout(False)

  End Sub
  Friend WithEvents pnlMappings As System.Windows.Forms.Panel
  Friend WithEvents Button1 As System.Windows.Forms.Button
End Class
