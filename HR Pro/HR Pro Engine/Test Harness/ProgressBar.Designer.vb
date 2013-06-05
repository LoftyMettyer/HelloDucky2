<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class ProgressBar
    Inherits System.Windows.Forms.UserControl

    'UserControl overrides dispose to clean up the component list.
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
    Me.txtDescription = New System.Windows.Forms.Label()
    Me.ProgressBar2 = New System.Windows.Forms.ProgressBar()
    Me.ProgressBar1 = New System.Windows.Forms.ProgressBar()
    Me.SuspendLayout()
    '
    'txtDescription
    '
    Me.txtDescription.AutoSize = True
    Me.txtDescription.Font = New System.Drawing.Font("Microsoft Sans Serif", 24.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
    Me.txtDescription.Location = New System.Drawing.Point(3, 0)
    Me.txtDescription.Name = "txtDescription"
    Me.txtDescription.Size = New System.Drawing.Size(177, 37)
    Me.txtDescription.TabIndex = 7
    Me.txtDescription.Text = "Description"
    '
    'ProgressBar2
    '
    Me.ProgressBar2.Location = New System.Drawing.Point(3, 69)
    Me.ProgressBar2.Name = "ProgressBar2"
    Me.ProgressBar2.Size = New System.Drawing.Size(442, 23)
    Me.ProgressBar2.TabIndex = 6
    '
    'ProgressBar1
    '
    Me.ProgressBar1.Location = New System.Drawing.Point(3, 39)
    Me.ProgressBar1.Name = "ProgressBar1"
    Me.ProgressBar1.Size = New System.Drawing.Size(442, 23)
    Me.ProgressBar1.TabIndex = 5
    '
    'ProgressBar
    '
    Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
    Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
    Me.Controls.Add(Me.txtDescription)
    Me.Controls.Add(Me.ProgressBar2)
    Me.Controls.Add(Me.ProgressBar1)
    Me.Name = "ProgressBar"
    Me.Size = New System.Drawing.Size(450, 97)
    Me.ResumeLayout(False)
    Me.PerformLayout()

  End Sub
  Friend WithEvents txtDescription As System.Windows.Forms.Label
  Friend WithEvents ProgressBar2 As System.Windows.Forms.ProgressBar
  Friend WithEvents ProgressBar1 As System.Windows.Forms.ProgressBar

End Class
