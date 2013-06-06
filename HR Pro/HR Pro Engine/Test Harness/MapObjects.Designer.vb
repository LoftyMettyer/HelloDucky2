<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class MapObjects
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
    Me.butGetObjectSelection = New System.Windows.Forms.Button()
    Me.txtUpdateScript = New System.Windows.Forms.TextBox()
    Me.CurrentPhase = New System.Windows.Forms.Label()
    Me.ProgressBar2 = New System.Windows.Forms.ProgressBar()
    Me.ProgressBar1 = New System.Windows.Forms.ProgressBar()
    Me.butGetMappings = New System.Windows.Forms.Button()
    Me.grdObjectSelection = New Test_Harness.HCMObjectPane()
    Me.SuspendLayout()
    '
    'butGetObjectSelection
    '
    Me.butGetObjectSelection.Location = New System.Drawing.Point(12, 12)
    Me.butGetObjectSelection.Name = "butGetObjectSelection"
    Me.butGetObjectSelection.Size = New System.Drawing.Size(106, 23)
    Me.butGetObjectSelection.TabIndex = 0
    Me.butGetObjectSelection.Text = "Populate Objects"
    Me.butGetObjectSelection.UseVisualStyleBackColor = True
    '
    'txtUpdateScript
    '
    Me.txtUpdateScript.Location = New System.Drawing.Point(390, 43)
    Me.txtUpdateScript.Name = "txtUpdateScript"
    Me.txtUpdateScript.Size = New System.Drawing.Size(182, 20)
    Me.txtUpdateScript.TabIndex = 8
    Me.txtUpdateScript.Text = "c:\dev\updatescript\exportdata.xml"
    '
    'CurrentPhase
    '
    Me.CurrentPhase.AutoSize = True
    Me.CurrentPhase.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
    Me.CurrentPhase.Location = New System.Drawing.Point(8, 477)
    Me.CurrentPhase.Name = "CurrentPhase"
    Me.CurrentPhase.Size = New System.Drawing.Size(78, 24)
    Me.CurrentPhase.TabIndex = 11
    Me.CurrentPhase.Text = "Phase..."
    '
    'ProgressBar2
    '
    Me.ProgressBar2.Location = New System.Drawing.Point(2, 534)
    Me.ProgressBar2.Name = "ProgressBar2"
    Me.ProgressBar2.Size = New System.Drawing.Size(328, 23)
    Me.ProgressBar2.TabIndex = 10
    '
    'ProgressBar1
    '
    Me.ProgressBar1.Location = New System.Drawing.Point(2, 504)
    Me.ProgressBar1.Name = "ProgressBar1"
    Me.ProgressBar1.Size = New System.Drawing.Size(328, 23)
    Me.ProgressBar1.TabIndex = 9
    '
    'butGetMappings
    '
    Me.butGetMappings.Location = New System.Drawing.Point(390, 69)
    Me.butGetMappings.Name = "butGetMappings"
    Me.butGetMappings.Size = New System.Drawing.Size(155, 38)
    Me.butGetMappings.TabIndex = 16
    Me.butGetMappings.Text = "Export Selected objects to file"
    Me.butGetMappings.UseVisualStyleBackColor = True
    '
    'grdObjectSelection
    '
    Me.grdObjectSelection.Location = New System.Drawing.Point(13, 42)
    Me.grdObjectSelection.Name = "grdObjectSelection"
    Me.grdObjectSelection.Size = New System.Drawing.Size(326, 393)
    Me.grdObjectSelection.TabIndex = 17
    '
    'MapObjects
    '
    Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
    Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
    Me.ClientSize = New System.Drawing.Size(853, 568)
    Me.Controls.Add(Me.grdObjectSelection)
    Me.Controls.Add(Me.butGetMappings)
    Me.Controls.Add(Me.CurrentPhase)
    Me.Controls.Add(Me.ProgressBar2)
    Me.Controls.Add(Me.ProgressBar1)
    Me.Controls.Add(Me.txtUpdateScript)
    Me.Controls.Add(Me.butGetObjectSelection)
    Me.Name = "MapObjects"
    Me.Text = "MapObjects"
    Me.ResumeLayout(False)
    Me.PerformLayout()

  End Sub
  Friend WithEvents butGetObjectSelection As System.Windows.Forms.Button
  Friend WithEvents txtUpdateScript As System.Windows.Forms.TextBox
  Friend WithEvents CurrentPhase As System.Windows.Forms.Label
  Friend WithEvents ProgressBar2 As System.Windows.Forms.ProgressBar
  Friend WithEvents ProgressBar1 As System.Windows.Forms.ProgressBar
  Friend WithEvents butGetMappings As System.Windows.Forms.Button
  Friend WithEvents grdObjectSelection As Test_Harness.HCMObjectPane
End Class
