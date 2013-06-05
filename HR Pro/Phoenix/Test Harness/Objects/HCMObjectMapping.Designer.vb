<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class HCMObjectMapping
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
    Me.cboFrom = New Infragistics.Win.UltraWinEditors.UltraComboEditor()
    Me.cboTo = New Infragistics.Win.UltraWinEditors.UltraComboEditor()
    Me.Label1 = New System.Windows.Forms.Label()
    Me.chkNewObject = New Infragistics.Win.UltraWinEditors.UltraCheckEditor()
    Me.lblType = New System.Windows.Forms.Label()
    CType(Me.cboFrom, System.ComponentModel.ISupportInitialize).BeginInit()
    CType(Me.cboTo, System.ComponentModel.ISupportInitialize).BeginInit()
    Me.SuspendLayout()
    '
    'cboFrom
    '
    Me.cboFrom.Enabled = False
    Me.cboFrom.Font = New System.Drawing.Font("Calibri", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
    Me.cboFrom.Location = New System.Drawing.Point(51, 3)
    Me.cboFrom.Name = "cboFrom"
    Me.cboFrom.Size = New System.Drawing.Size(180, 25)
    Me.cboFrom.TabIndex = 0
    '
    'cboTo
    '
    Me.cboTo.Font = New System.Drawing.Font("Calibri", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
    Me.cboTo.Location = New System.Drawing.Point(379, 3)
    Me.cboTo.Name = "cboTo"
    Me.cboTo.Size = New System.Drawing.Size(166, 25)
    Me.cboTo.TabIndex = 1
    '
    'Label1
    '
    Me.Label1.AutoSize = True
    Me.Label1.Font = New System.Drawing.Font("Calibri", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
    Me.Label1.Location = New System.Drawing.Point(246, 7)
    Me.Label1.Name = "Label1"
    Me.Label1.Size = New System.Drawing.Size(51, 15)
    Me.Label1.TabIndex = 2
    Me.Label1.Text = "maps to"
    '
    'chkNewObject
    '
    Me.chkNewObject.Location = New System.Drawing.Point(312, 5)
    Me.chkNewObject.Name = "chkNewObject"
    Me.chkNewObject.Size = New System.Drawing.Size(48, 20)
    Me.chkNewObject.TabIndex = 3
    Me.chkNewObject.Text = "New"
    '
    'lblType
    '
    Me.lblType.AutoSize = True
    Me.lblType.Location = New System.Drawing.Point(4, 11)
    Me.lblType.Name = "lblType"
    Me.lblType.Size = New System.Drawing.Size(31, 13)
    Me.lblType.TabIndex = 4
    Me.lblType.Text = "Type"
    '
    'HCMObjectMapping
    '
    Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
    Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
    Me.Controls.Add(Me.lblType)
    Me.Controls.Add(Me.chkNewObject)
    Me.Controls.Add(Me.Label1)
    Me.Controls.Add(Me.cboTo)
    Me.Controls.Add(Me.cboFrom)
    Me.Name = "HCMObjectMapping"
    Me.Size = New System.Drawing.Size(610, 36)
    CType(Me.cboFrom, System.ComponentModel.ISupportInitialize).EndInit()
    CType(Me.cboTo, System.ComponentModel.ISupportInitialize).EndInit()
    Me.ResumeLayout(False)
    Me.PerformLayout()

  End Sub
  Friend WithEvents cboFrom As Infragistics.Win.UltraWinEditors.UltraComboEditor
  Friend WithEvents cboTo As Infragistics.Win.UltraWinEditors.UltraComboEditor
  Friend WithEvents Label1 As System.Windows.Forms.Label
  Friend WithEvents chkNewObject As Infragistics.Win.UltraWinEditors.UltraCheckEditor
  Friend WithEvents lblType As System.Windows.Forms.Label

End Class
