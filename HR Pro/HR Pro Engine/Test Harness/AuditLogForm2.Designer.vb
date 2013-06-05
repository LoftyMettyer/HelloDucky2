<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class AuditLogForm2
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
        Dim DateButton1 As Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton = New Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton()
        Dim DateButton2 As Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton = New Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton()
        Me.auditLogsGrid = New Infragistics.Win.UltraWinGrid.UltraGrid()
        Me.datePanel = New System.Windows.Forms.Panel()
        Me.dateLabel = New Infragistics.Win.Misc.UltraLabel()
        Me.showButton = New Infragistics.Win.Misc.UltraButton()
        Me.dateFromEditor = New Infragistics.Win.UltraWinSchedule.UltraCalendarCombo()
        Me.dateToEditor = New Infragistics.Win.UltraWinSchedule.UltraCalendarCombo()
        Me.periodEditor = New Infragistics.Win.UltraWinEditors.UltraComboEditor()
        CType(Me.auditLogsGrid, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.datePanel.SuspendLayout()
        CType(Me.dateFromEditor, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.dateToEditor, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.periodEditor, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'auditLogsGrid
        '
        Me.auditLogsGrid.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.auditLogsGrid.Cursor = System.Windows.Forms.Cursors.Hand
        Appearance2.BackColor = System.Drawing.SystemColors.Window
        Appearance2.BorderColor = System.Drawing.SystemColors.Window
        Me.auditLogsGrid.DisplayLayout.Appearance = Appearance2
        Appearance3.BackColor = System.Drawing.SystemColors.Highlight
        Appearance3.ForeColor = System.Drawing.SystemColors.HighlightText
        Me.auditLogsGrid.DisplayLayout.Appearances.Add(Appearance3)
        Me.auditLogsGrid.DisplayLayout.CaptionVisible = Infragistics.Win.DefaultableBoolean.[False]
        Me.auditLogsGrid.DisplayLayout.InterBandSpacing = 0
        Me.auditLogsGrid.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No
        Me.auditLogsGrid.DisplayLayout.Override.AllowColSwapping = Infragistics.Win.UltraWinGrid.AllowColSwapping.NotAllowed
        Me.auditLogsGrid.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.[False]
        Me.auditLogsGrid.DisplayLayout.Override.AllowGroupBy = Infragistics.Win.DefaultableBoolean.[True]
        Me.auditLogsGrid.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.[False]
        Me.auditLogsGrid.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.RowSelect
        Me.auditLogsGrid.DisplayLayout.Override.FixedCellSeparatorColor = System.Drawing.SystemColors.Window
        Appearance4.BackColor = System.Drawing.SystemColors.InactiveCaptionText
        Me.auditLogsGrid.DisplayLayout.Override.HeaderAppearance = Appearance4
        Me.auditLogsGrid.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti
        Me.auditLogsGrid.DisplayLayout.Override.RowSelectors = Infragistics.Win.DefaultableBoolean.[False]
        Me.auditLogsGrid.DisplayLayout.Override.SelectedRowAppearance = New Infragistics.Win.LinkedAppearance(74988782)
        Me.auditLogsGrid.DisplayLayout.Override.SelectTypeCell = Infragistics.Win.UltraWinGrid.SelectType.None
        Me.auditLogsGrid.DisplayLayout.Override.SelectTypeCol = Infragistics.Win.UltraWinGrid.SelectType.None
        Me.auditLogsGrid.DisplayLayout.Override.SelectTypeGroupByRow = Infragistics.Win.UltraWinGrid.SelectType.None
        Me.auditLogsGrid.DisplayLayout.Override.SelectTypeRow = Infragistics.Win.UltraWinGrid.SelectType.Extended
        Me.auditLogsGrid.DisplayLayout.ScrollStyle = Infragistics.Win.UltraWinGrid.ScrollStyle.Immediate
        Me.auditLogsGrid.DisplayLayout.TabNavigation = Infragistics.Win.UltraWinGrid.TabNavigation.NextControlOnLastCell
        Me.auditLogsGrid.DisplayLayout.ViewStyleBand = Infragistics.Win.UltraWinGrid.ViewStyleBand.OutlookGroupBy
        Me.auditLogsGrid.Font = New System.Drawing.Font("Calibri", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Pixel)
        Me.auditLogsGrid.Location = New System.Drawing.Point(0, 41)
        Me.auditLogsGrid.Margin = New System.Windows.Forms.Padding(6, 5, 6, 5)
        Me.auditLogsGrid.Name = "auditLogsGrid"
        Me.auditLogsGrid.Size = New System.Drawing.Size(678, 333)
        Me.auditLogsGrid.TabIndex = 19
        Me.auditLogsGrid.TextRenderingMode = Infragistics.Win.TextRenderingMode.GDI
        '
        'datePanel
        '
        Me.datePanel.Controls.Add(Me.dateLabel)
        Me.datePanel.Controls.Add(Me.showButton)
        Me.datePanel.Controls.Add(Me.dateFromEditor)
        Me.datePanel.Controls.Add(Me.dateToEditor)
        Me.datePanel.Location = New System.Drawing.Point(152, 10)
        Me.datePanel.Name = "datePanel"
        Me.datePanel.Size = New System.Drawing.Size(356, 23)
        Me.datePanel.TabIndex = 23
        Me.datePanel.Visible = False
        '
        'dateLabel
        '
        Me.dateLabel.AutoSize = True
        Me.dateLabel.Location = New System.Drawing.Point(108, 3)
        Me.dateLabel.Name = "dateLabel"
        Me.dateLabel.Size = New System.Drawing.Size(23, 14)
        Me.dateLabel.TabIndex = 31
        Me.dateLabel.Text = "and"
        '
        'showButton
        '
        Me.showButton.Location = New System.Drawing.Point(247, -1)
        Me.showButton.Name = "showButton"
        Me.showButton.Size = New System.Drawing.Size(58, 23)
        Me.showButton.TabIndex = 30
        Me.showButton.Text = "Show"
        '
        'dateFromEditor
        '
        Me.dateFromEditor.BackColor = System.Drawing.SystemColors.Window
        Me.dateFromEditor.DateButtons.Add(DateButton1)
        Me.dateFromEditor.Location = New System.Drawing.Point(0, 0)
        Me.dateFromEditor.Name = "dateFromEditor"
        Me.dateFromEditor.NonAutoSizeHeight = 21
        Me.dateFromEditor.NullDateLabel = ""
        Me.dateFromEditor.Size = New System.Drawing.Size(102, 21)
        Me.dateFromEditor.TabIndex = 29
        Me.dateFromEditor.Value = ""
        '
        'dateToEditor
        '
        Me.dateToEditor.BackColor = System.Drawing.SystemColors.Window
        Me.dateToEditor.DateButtons.Add(DateButton2)
        Me.dateToEditor.Location = New System.Drawing.Point(139, 0)
        Me.dateToEditor.Name = "dateToEditor"
        Me.dateToEditor.NonAutoSizeHeight = 21
        Me.dateToEditor.NullDateLabel = ""
        Me.dateToEditor.Size = New System.Drawing.Size(102, 21)
        Me.dateToEditor.TabIndex = 28
        Me.dateToEditor.Value = ""
        '
        'periodEditor
        '
        Me.periodEditor.DropDownStyle = Infragistics.Win.DropDownStyle.DropDownList
        Me.periodEditor.LimitToList = True
        Me.periodEditor.Location = New System.Drawing.Point(13, 10)
        Me.periodEditor.Name = "periodEditor"
        Me.periodEditor.Size = New System.Drawing.Size(123, 21)
        Me.periodEditor.TabIndex = 24
        Me.periodEditor.ValueMember = ""
        '
        'AuditLogForm2
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(678, 414)
        Me.Controls.Add(Me.periodEditor)
        Me.Controls.Add(Me.datePanel)
        Me.Controls.Add(Me.auditLogsGrid)
        Me.Name = "AuditLogForm2"
        Me.Text = "Audit Log"
        CType(Me.auditLogsGrid, System.ComponentModel.ISupportInitialize).EndInit()
        Me.datePanel.ResumeLayout(False)
        Me.datePanel.PerformLayout()
        CType(Me.dateFromEditor, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.dateToEditor, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.periodEditor, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents auditLogsGrid As Infragistics.Win.UltraWinGrid.UltraGrid
    Friend WithEvents datePanel As System.Windows.Forms.Panel
    Friend WithEvents dateLabel As Infragistics.Win.Misc.UltraLabel
    Friend WithEvents showButton As Infragistics.Win.Misc.UltraButton
    Friend WithEvents dateFromEditor As Infragistics.Win.UltraWinSchedule.UltraCalendarCombo
    Friend WithEvents dateToEditor As Infragistics.Win.UltraWinSchedule.UltraCalendarCombo
    Friend WithEvents periodEditor As Infragistics.Win.UltraWinEditors.UltraComboEditor
End Class
