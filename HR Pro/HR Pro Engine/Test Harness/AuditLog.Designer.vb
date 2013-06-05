<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class AuditLog
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
    Dim Appearance2 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
    Dim Appearance3 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance("Highlight", 74988782)
    Dim Appearance4 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
    Me.grdAudit = New Infragistics.Win.UltraWinGrid.UltraGrid()
    Me.butOutput = New System.Windows.Forms.Button()
    Me.UltraGridExcelExporter1 = New Infragistics.Win.UltraWinGrid.ExcelExport.UltraGridExcelExporter(Me.components)
    Me.txtFilePath = New System.Windows.Forms.TextBox()
    CType(Me.grdAudit, System.ComponentModel.ISupportInitialize).BeginInit()
    Me.SuspendLayout()
    '
    'grdAudit
    '
    Me.grdAudit.Cursor = System.Windows.Forms.Cursors.Hand
    Appearance2.BackColor = System.Drawing.SystemColors.Window
    Appearance2.BorderColor = System.Drawing.SystemColors.Window
    Me.grdAudit.DisplayLayout.Appearance = Appearance2
    Appearance3.BackColor = System.Drawing.SystemColors.Highlight
    Appearance3.ForeColor = System.Drawing.SystemColors.HighlightText
    Me.grdAudit.DisplayLayout.Appearances.Add(Appearance3)
    Me.grdAudit.DisplayLayout.InterBandSpacing = 0
    Me.grdAudit.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No
    Me.grdAudit.DisplayLayout.Override.AllowColMoving = Infragistics.Win.UltraWinGrid.AllowColMoving.NotAllowed
    Me.grdAudit.DisplayLayout.Override.AllowColSwapping = Infragistics.Win.UltraWinGrid.AllowColSwapping.NotAllowed
    Me.grdAudit.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.[False]
    Me.grdAudit.DisplayLayout.Override.AllowRowFiltering = Infragistics.Win.DefaultableBoolean.[True]
    Me.grdAudit.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.[False]
    Me.grdAudit.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.RowSelect
    Me.grdAudit.DisplayLayout.Override.FilterUIType = Infragistics.Win.UltraWinGrid.FilterUIType.HeaderIcons
    Me.grdAudit.DisplayLayout.Override.FixedCellSeparatorColor = System.Drawing.SystemColors.Window
    Me.grdAudit.DisplayLayout.Override.FixedHeaderIndicator = Infragistics.Win.UltraWinGrid.FixedHeaderIndicator.None
    Appearance4.BackColor = System.Drawing.SystemColors.InactiveCaptionText
    Me.grdAudit.DisplayLayout.Override.HeaderAppearance = Appearance4
    Me.grdAudit.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti
    Me.grdAudit.DisplayLayout.Override.RowSelectors = Infragistics.Win.DefaultableBoolean.[False]
    Me.grdAudit.DisplayLayout.Override.RowSizing = Infragistics.Win.UltraWinGrid.RowSizing.Fixed
    Me.grdAudit.DisplayLayout.Override.SelectedRowAppearance = New Infragistics.Win.LinkedAppearance(74988782)
    Me.grdAudit.DisplayLayout.Override.SelectTypeRow = Infragistics.Win.UltraWinGrid.SelectType.Extended
    Me.grdAudit.DisplayLayout.Override.TipStyleCell = Infragistics.Win.UltraWinGrid.TipStyle.Hide
    Me.grdAudit.DisplayLayout.ScrollBounds = Infragistics.Win.UltraWinGrid.ScrollBounds.ScrollToFill
    Me.grdAudit.DisplayLayout.ScrollStyle = Infragistics.Win.UltraWinGrid.ScrollStyle.Immediate
    Me.grdAudit.DisplayLayout.UseFixedHeaders = True
    Me.grdAudit.Dock = System.Windows.Forms.DockStyle.Top
    Me.grdAudit.Font = New System.Drawing.Font("Calibri", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Pixel)
    Me.grdAudit.Location = New System.Drawing.Point(0, 0)
    Me.grdAudit.Margin = New System.Windows.Forms.Padding(6, 5, 6, 5)
    Me.grdAudit.Name = "grdAudit"
    Me.grdAudit.Size = New System.Drawing.Size(1077, 469)
    Me.grdAudit.TabIndex = 18
    Me.grdAudit.TextRenderingMode = Infragistics.Win.TextRenderingMode.GDI
    '
    'butOutput
    '
    Me.butOutput.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
    Me.butOutput.Location = New System.Drawing.Point(307, 484)
    Me.butOutput.Name = "butOutput"
    Me.butOutput.Size = New System.Drawing.Size(104, 23)
    Me.butOutput.TabIndex = 19
    Me.butOutput.Text = "Output..."
    Me.butOutput.UseVisualStyleBackColor = True
    '
    'txtFilePath
    '
    Me.txtFilePath.Location = New System.Drawing.Point(13, 484)
    Me.txtFilePath.Name = "txtFilePath"
    Me.txtFilePath.Size = New System.Drawing.Size(273, 21)
    Me.txtFilePath.TabIndex = 20
    Me.txtFilePath.Text = "c:\dev\output.xls"
    '
    'AuditLog
    '
    Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 13.0!)
    Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
    Me.ClientSize = New System.Drawing.Size(1077, 519)
    Me.Controls.Add(Me.txtFilePath)
    Me.Controls.Add(Me.butOutput)
    Me.Controls.Add(Me.grdAudit)
    Me.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
    Me.Name = "AuditLog"
    Me.ShowIcon = False
    Me.Text = "Audit Log"
    CType(Me.grdAudit, System.ComponentModel.ISupportInitialize).EndInit()
    Me.ResumeLayout(False)
    Me.PerformLayout()

  End Sub
  Friend WithEvents grdAudit As Infragistics.Win.UltraWinGrid.UltraGrid
  Friend WithEvents butOutput As System.Windows.Forms.Button
  Friend WithEvents UltraGridExcelExporter1 As Infragistics.Win.UltraWinGrid.ExcelExport.UltraGridExcelExporter
  Friend WithEvents txtFilePath As System.Windows.Forms.TextBox
End Class
