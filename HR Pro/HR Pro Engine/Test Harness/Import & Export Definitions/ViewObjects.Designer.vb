<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class ViewObjects
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
    Dim Appearance2 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
    Dim Appearance3 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance("Highlight", 74988782)
    Dim Appearance4 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
    Me.Button1 = New System.Windows.Forms.Button()
    Me.grdThings = New Infragistics.Win.UltraWinGrid.UltraGrid()
    Me.butExport = New System.Windows.Forms.Button()
    Me.butImport = New System.Windows.Forms.Button()
    CType(Me.grdThings, System.ComponentModel.ISupportInitialize).BeginInit()
    Me.SuspendLayout()
    '
    'Button1
    '
    Me.Button1.Location = New System.Drawing.Point(600, 423)
    Me.Button1.Name = "Button1"
    Me.Button1.Size = New System.Drawing.Size(75, 23)
    Me.Button1.TabIndex = 16
    Me.Button1.Text = "Cancel"
    Me.Button1.UseVisualStyleBackColor = True
    '
    'grdThings
    '
    Me.grdThings.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                Or System.Windows.Forms.AnchorStyles.Left) _
                Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.grdThings.Cursor = System.Windows.Forms.Cursors.Hand
    Appearance2.BackColor = System.Drawing.SystemColors.Window
    Appearance2.BorderColor = System.Drawing.SystemColors.Window
    Me.grdThings.DisplayLayout.Appearance = Appearance2
    Appearance3.BackColor = System.Drawing.SystemColors.Highlight
    Appearance3.ForeColor = System.Drawing.SystemColors.HighlightText
    Me.grdThings.DisplayLayout.Appearances.Add(Appearance3)
    Me.grdThings.DisplayLayout.InterBandSpacing = 0
    Me.grdThings.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No
    Me.grdThings.DisplayLayout.Override.AllowColMoving = Infragistics.Win.UltraWinGrid.AllowColMoving.NotAllowed
    Me.grdThings.DisplayLayout.Override.AllowColSwapping = Infragistics.Win.UltraWinGrid.AllowColSwapping.NotAllowed
    Me.grdThings.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.[False]
    Me.grdThings.DisplayLayout.Override.AllowRowFiltering = Infragistics.Win.DefaultableBoolean.[False]
    Me.grdThings.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.[False]
    Me.grdThings.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.RowSelect
    Me.grdThings.DisplayLayout.Override.FixedCellSeparatorColor = System.Drawing.SystemColors.Window
    Me.grdThings.DisplayLayout.Override.FixedHeaderIndicator = Infragistics.Win.UltraWinGrid.FixedHeaderIndicator.None
    Appearance4.BackColor = System.Drawing.SystemColors.InactiveCaptionText
    Me.grdThings.DisplayLayout.Override.HeaderAppearance = Appearance4
    Me.grdThings.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti
    Me.grdThings.DisplayLayout.Override.RowSelectors = Infragistics.Win.DefaultableBoolean.[False]
    Me.grdThings.DisplayLayout.Override.RowSizing = Infragistics.Win.UltraWinGrid.RowSizing.Fixed
    Me.grdThings.DisplayLayout.Override.SelectedRowAppearance = New Infragistics.Win.LinkedAppearance(74988782)
    Me.grdThings.DisplayLayout.Override.SelectTypeRow = Infragistics.Win.UltraWinGrid.SelectType.Extended
    Me.grdThings.DisplayLayout.Override.TipStyleCell = Infragistics.Win.UltraWinGrid.TipStyle.Hide
    Me.grdThings.DisplayLayout.ScrollBounds = Infragistics.Win.UltraWinGrid.ScrollBounds.ScrollToFill
    Me.grdThings.DisplayLayout.ScrollStyle = Infragistics.Win.UltraWinGrid.ScrollStyle.Immediate
    Me.grdThings.DisplayLayout.UseFixedHeaders = True
    Me.grdThings.Font = New System.Drawing.Font("Calibri", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Pixel)
    Me.grdThings.Location = New System.Drawing.Point(14, 14)
    Me.grdThings.Margin = New System.Windows.Forms.Padding(5)
    Me.grdThings.Name = "grdThings"
    Me.grdThings.Size = New System.Drawing.Size(567, 432)
    Me.grdThings.TabIndex = 17
    Me.grdThings.TextRenderingMode = Infragistics.Win.TextRenderingMode.GDI
    '
    'butExport
    '
    Me.butExport.Location = New System.Drawing.Point(600, 32)
    Me.butExport.Name = "butExport"
    Me.butExport.Size = New System.Drawing.Size(75, 23)
    Me.butExport.TabIndex = 18
    Me.butExport.Text = "Export This..."
    Me.butExport.UseVisualStyleBackColor = True
    '
    'butImport
    '
    Me.butImport.Font = New System.Drawing.Font("Calibri", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
    Me.butImport.Location = New System.Drawing.Point(600, 198)
    Me.butImport.Name = "butImport"
    Me.butImport.Size = New System.Drawing.Size(75, 47)
    Me.butImport.TabIndex = 19
    Me.butImport.Text = "Import File"
    Me.butImport.UseVisualStyleBackColor = True
    '
    'ViewObjects
    '
    Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
    Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
    Me.ClientSize = New System.Drawing.Size(703, 461)
    Me.Controls.Add(Me.butImport)
    Me.Controls.Add(Me.butExport)
    Me.Controls.Add(Me.grdThings)
    Me.Controls.Add(Me.Button1)
    Me.Name = "ViewObjects"
    Me.Text = "View Objects"
    CType(Me.grdThings, System.ComponentModel.ISupportInitialize).EndInit()
    Me.ResumeLayout(False)

  End Sub
  Friend WithEvents Button1 As System.Windows.Forms.Button
  Friend WithEvents grdThings As Infragistics.Win.UltraWinGrid.UltraGrid
  Friend WithEvents butExport As System.Windows.Forms.Button
  Friend WithEvents butImport As System.Windows.Forms.Button
End Class
