<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class HCMObjectPane
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
    Dim Appearance1 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
    Dim Appearance2 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
    Dim Appearance4 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
    Dim Appearance3 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
    Dim Appearance9 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
    Dim Appearance5 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
    Dim Appearance12 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
    Dim Appearance8 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
    Dim Appearance6 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
    Dim Appearance7 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
    Dim Appearance10 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
    Dim Appearance11 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
    Me.UltraGrid1 = New Infragistics.Win.UltraWinGrid.UltraGrid()
    CType(Me.UltraGrid1, System.ComponentModel.ISupportInitialize).BeginInit()
    Me.SuspendLayout()
    '
    'UltraGrid1
    '
    Appearance1.BackColor = System.Drawing.SystemColors.Window
    Appearance1.BorderColor = System.Drawing.SystemColors.InactiveCaption
    Me.UltraGrid1.DisplayLayout.Appearance = Appearance1
    Me.UltraGrid1.DisplayLayout.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid
    Me.UltraGrid1.DisplayLayout.CaptionVisible = Infragistics.Win.DefaultableBoolean.[False]
    Appearance2.BackColor = System.Drawing.SystemColors.ActiveBorder
    Appearance2.BackColor2 = System.Drawing.SystemColors.ControlDark
    Appearance2.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical
    Appearance2.BorderColor = System.Drawing.SystemColors.Window
    Me.UltraGrid1.DisplayLayout.GroupByBox.Appearance = Appearance2
    Appearance4.ForeColor = System.Drawing.SystemColors.GrayText
    Me.UltraGrid1.DisplayLayout.GroupByBox.BandLabelAppearance = Appearance4
    Me.UltraGrid1.DisplayLayout.GroupByBox.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid
    Appearance3.BackColor = System.Drawing.SystemColors.ControlLightLight
    Appearance3.BackColor2 = System.Drawing.SystemColors.Control
    Appearance3.BackGradientStyle = Infragistics.Win.GradientStyle.Horizontal
    Appearance3.ForeColor = System.Drawing.SystemColors.GrayText
    Me.UltraGrid1.DisplayLayout.GroupByBox.PromptAppearance = Appearance3
    Me.UltraGrid1.DisplayLayout.MaxColScrollRegions = 1
    Me.UltraGrid1.DisplayLayout.MaxRowScrollRegions = 1
    Appearance9.BackColor = System.Drawing.SystemColors.Window
    Appearance9.ForeColor = System.Drawing.SystemColors.ControlText
    Me.UltraGrid1.DisplayLayout.Override.ActiveCellAppearance = Appearance9
    Appearance5.BackColor = System.Drawing.SystemColors.Highlight
    Appearance5.ForeColor = System.Drawing.SystemColors.HighlightText
    Me.UltraGrid1.DisplayLayout.Override.ActiveRowAppearance = Appearance5
    Me.UltraGrid1.DisplayLayout.Override.BorderStyleCell = Infragistics.Win.UIElementBorderStyle.Dotted
    Me.UltraGrid1.DisplayLayout.Override.BorderStyleRow = Infragistics.Win.UIElementBorderStyle.Dotted
    Appearance12.BackColor = System.Drawing.SystemColors.Window
    Me.UltraGrid1.DisplayLayout.Override.CardAreaAppearance = Appearance12
    Appearance8.BorderColor = System.Drawing.Color.Silver
    Appearance8.TextTrimming = Infragistics.Win.TextTrimming.EllipsisCharacter
    Me.UltraGrid1.DisplayLayout.Override.CellAppearance = Appearance8
    Me.UltraGrid1.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.EditAndSelectText
    Me.UltraGrid1.DisplayLayout.Override.CellPadding = 0
    Appearance6.BackColor = System.Drawing.SystemColors.Control
    Appearance6.BackColor2 = System.Drawing.SystemColors.ControlDark
    Appearance6.BackGradientAlignment = Infragistics.Win.GradientAlignment.Element
    Appearance6.BackGradientStyle = Infragistics.Win.GradientStyle.Horizontal
    Appearance6.BorderColor = System.Drawing.SystemColors.Window
    Me.UltraGrid1.DisplayLayout.Override.GroupByRowAppearance = Appearance6
    Appearance7.TextHAlignAsString = "Left"
    Me.UltraGrid1.DisplayLayout.Override.HeaderAppearance = Appearance7
    Me.UltraGrid1.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti
    Me.UltraGrid1.DisplayLayout.Override.HeaderStyle = Infragistics.Win.HeaderStyle.WindowsXPCommand
    Appearance10.BackColor = System.Drawing.SystemColors.Window
    Appearance10.BorderColor = System.Drawing.Color.Silver
    Me.UltraGrid1.DisplayLayout.Override.RowAppearance = Appearance10
    Me.UltraGrid1.DisplayLayout.Override.RowSelectors = Infragistics.Win.DefaultableBoolean.[False]
    Appearance11.BackColor = System.Drawing.SystemColors.ControlLight
    Me.UltraGrid1.DisplayLayout.Override.TemplateAddRowAppearance = Appearance11
    Me.UltraGrid1.DisplayLayout.ScrollBounds = Infragistics.Win.UltraWinGrid.ScrollBounds.ScrollToFill
    Me.UltraGrid1.DisplayLayout.ScrollStyle = Infragistics.Win.UltraWinGrid.ScrollStyle.Immediate
    Me.UltraGrid1.DisplayLayout.ViewStyleBand = Infragistics.Win.UltraWinGrid.ViewStyleBand.OutlookGroupBy
    Me.UltraGrid1.Dock = System.Windows.Forms.DockStyle.Fill
    Me.UltraGrid1.Location = New System.Drawing.Point(0, 0)
    Me.UltraGrid1.Name = "UltraGrid1"
    Me.UltraGrid1.Size = New System.Drawing.Size(572, 247)
    Me.UltraGrid1.TabIndex = 0
    Me.UltraGrid1.Text = "UltraGrid1"
    '
    'HCMObjectPane
    '
    Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
    Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
    Me.Controls.Add(Me.UltraGrid1)
    Me.Name = "HCMObjectPane"
    Me.Size = New System.Drawing.Size(572, 247)
    CType(Me.UltraGrid1, System.ComponentModel.ISupportInitialize).EndInit()
    Me.ResumeLayout(False)

  End Sub
  Friend WithEvents UltraGrid1 As Infragistics.Win.UltraWinGrid.UltraGrid

End Class
