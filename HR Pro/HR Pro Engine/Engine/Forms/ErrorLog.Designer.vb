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
      Me.grdErrors = New System.Windows.Forms.DataGridView()
      CType(Me.grdErrors, System.ComponentModel.ISupportInitialize).BeginInit()
      Me.SuspendLayout()
      '
      'grdErrors
      '
      Me.grdErrors.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
      Me.grdErrors.Location = New System.Drawing.Point(46, 25)
      Me.grdErrors.Name = "grdErrors"
      Me.grdErrors.Size = New System.Drawing.Size(533, 331)
      Me.grdErrors.TabIndex = 0
      '
      'ErrorLog
      '
      Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
      Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
      Me.ClientSize = New System.Drawing.Size(640, 425)
      Me.Controls.Add(Me.grdErrors)
      Me.Name = "ErrorLog"
      Me.Text = "ErrorLog"
      CType(Me.grdErrors, System.ComponentModel.ISupportInitialize).EndInit()
      Me.ResumeLayout(False)

    End Sub
    Friend WithEvents grdErrors As System.Windows.Forms.DataGridView
  End Class
End Namespace
