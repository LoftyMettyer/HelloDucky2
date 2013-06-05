<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class AuditLogForm
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
        Dim Appearance3 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance("Highlight", 74988782)
        Dim ValueListItem2 As Infragistics.Win.ValueListItem = New Infragistics.Win.ValueListItem()
        Dim ValueListItem3 As Infragistics.Win.ValueListItem = New Infragistics.Win.ValueListItem()
        Dim ValueListItem4 As Infragistics.Win.ValueListItem = New Infragistics.Win.ValueListItem()
        Dim DateButton1 As Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton = New Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton()
        Dim DateButton2 As Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton = New Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton()
        Me.grdAudit = New Infragistics.Win.UltraWinGrid.UltraGrid()
        Me.butOutput = New System.Windows.Forms.Button()
        Me.UltraGridExcelExporter1 = New Infragistics.Win.UltraWinGrid.ExcelExport.UltraGridExcelExporter(Me.components)
        Me.txtFilePath = New System.Windows.Forms.TextBox()
        Me.periodEditor = New Infragistics.Win.UltraWinEditors.UltraComboEditor()
        Me.datePanel = New System.Windows.Forms.Panel()
        Me.dateLabel = New Infragistics.Win.Misc.UltraLabel()
        Me.showButton = New Infragistics.Win.Misc.UltraButton()
        Me.dateFromEditor = New Infragistics.Win.UltraWinSchedule.UltraCalendarCombo()
        Me.dateToEditor = New Infragistics.Win.UltraWinSchedule.UltraCalendarCombo()
        CType(Me.grdAudit, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.periodEditor, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.datePanel.SuspendLayout()
        CType(Me.dateFromEditor, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.dateToEditor, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'grdAudit
        '
        Me.grdAudit.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Appearance3.BackColor = System.Drawing.SystemColors.Highlight
        Appearance3.ForeColor = System.Drawing.SystemColors.HighlightText
        Me.grdAudit.DisplayLayout.Appearances.Add(Appearance3)
        Me.grdAudit.DisplayLayout.AutoFitStyle = Infragistics.Win.UltraWinGrid.AutoFitStyle.ResizeAllColumns
        Me.grdAudit.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No
        Me.grdAudit.DisplayLayout.Override.AllowColSwapping = Infragistics.Win.UltraWinGrid.AllowColSwapping.NotAllowed
        Me.grdAudit.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.[False]
        Me.grdAudit.DisplayLayout.Override.AllowGroupBy = Infragistics.Win.DefaultableBoolean.[True]
        Me.grdAudit.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.[False]
        Me.grdAudit.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.RowSelect
        Me.grdAudit.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti
        Me.grdAudit.DisplayLayout.Override.RowSelectors = Infragistics.Win.DefaultableBoolean.[False]
        Me.grdAudit.DisplayLayout.Override.SelectTypeCell = Infragistics.Win.UltraWinGrid.SelectType.None
        Me.grdAudit.DisplayLayout.Override.SelectTypeCol = Infragistics.Win.UltraWinGrid.SelectType.None
        Me.grdAudit.DisplayLayout.Override.SelectTypeGroupByRow = Infragistics.Win.UltraWinGrid.SelectType.None
        Me.grdAudit.DisplayLayout.Override.SelectTypeRow = Infragistics.Win.UltraWinGrid.SelectType.[Single]
        Me.grdAudit.DisplayLayout.ScrollStyle = Infragistics.Win.UltraWinGrid.ScrollStyle.Immediate
        Me.grdAudit.DisplayLayout.TabNavigation = Infragistics.Win.UltraWinGrid.TabNavigation.NextControlOnLastCell
        Me.grdAudit.Location = New System.Drawing.Point(0, 41)
        Me.grdAudit.Margin = New System.Windows.Forms.Padding(6, 5, 6, 5)
        Me.grdAudit.Name = "grdAudit"
        Me.grdAudit.Size = New System.Drawing.Size(833, 336)
        Me.grdAudit.TabIndex = 18
        Me.grdAudit.TextRenderingMode = Infragistics.Win.TextRenderingMode.GDI
        Me.grdAudit.UseOsThemes = Infragistics.Win.DefaultableBoolean.[True]
        '
        'butOutput
        '
        Me.butOutput.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.butOutput.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.butOutput.Location = New System.Drawing.Point(291, 385)
        Me.butOutput.Name = "butOutput"
        Me.butOutput.Size = New System.Drawing.Size(66, 23)
        Me.butOutput.TabIndex = 19
        Me.butOutput.Text = "Output..."
        Me.butOutput.UseVisualStyleBackColor = True
        '
        'txtFilePath
        '
        Me.txtFilePath.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.txtFilePath.Location = New System.Drawing.Point(13, 385)
        Me.txtFilePath.Name = "txtFilePath"
        Me.txtFilePath.Size = New System.Drawing.Size(273, 21)
        Me.txtFilePath.TabIndex = 20
        Me.txtFilePath.Text = "c:\dev\output.xls"
        '
        'periodEditor
        '
        Me.periodEditor.DropDownStyle = Infragistics.Win.DropDownStyle.DropDownList
        ValueListItem2.DataValue = 1
        ValueListItem2.DisplayText = "This Month"
        ValueListItem3.DataValue = 2
        ValueListItem3.DisplayText = "Last Month"
        ValueListItem4.DataValue = 3
        ValueListItem4.DisplayText = "Between"
        Me.periodEditor.Items.AddRange(New Infragistics.Win.ValueListItem() {ValueListItem2, ValueListItem3, ValueListItem4})
        Me.periodEditor.LimitToList = True
        Me.periodEditor.Location = New System.Drawing.Point(13, 10)
        Me.periodEditor.Name = "periodEditor"
        Me.periodEditor.Size = New System.Drawing.Size(123, 22)
        Me.periodEditor.TabIndex = 21
        Me.periodEditor.ValueMember = ""
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
        Me.datePanel.TabIndex = 22
        '
        'dateLabel
        '
        Me.dateLabel.AutoSize = True
        Me.dateLabel.Location = New System.Drawing.Point(108, 3)
        Me.dateLabel.Name = "dateLabel"
        Me.dateLabel.Size = New System.Drawing.Size(25, 15)
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
        'AuditLogForm
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(833, 413)
        Me.Controls.Add(Me.datePanel)
        Me.Controls.Add(Me.periodEditor)
        Me.Controls.Add(Me.txtFilePath)
        Me.Controls.Add(Me.butOutput)
        Me.Controls.Add(Me.grdAudit)
        Me.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Name = "AuditLogForm"
        Me.ShowIcon = False
        Me.Text = "Audit Log"
        CType(Me.grdAudit, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.periodEditor, System.ComponentModel.ISupportInitialize).EndInit()
        Me.datePanel.ResumeLayout(False)
        Me.datePanel.PerformLayout()
        CType(Me.dateFromEditor, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.dateToEditor, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents grdAudit As Infragistics.Win.UltraWinGrid.UltraGrid
    Friend WithEvents butOutput As System.Windows.Forms.Button
    Friend WithEvents UltraGridExcelExporter1 As Infragistics.Win.UltraWinGrid.ExcelExport.UltraGridExcelExporter
    Friend WithEvents txtFilePath As System.Windows.Forms.TextBox
    Friend WithEvents periodEditor As Infragistics.Win.UltraWinEditors.UltraComboEditor
    Friend WithEvents datePanel As System.Windows.Forms.Panel
    Friend WithEvents dateLabel As Infragistics.Win.Misc.UltraLabel
    Friend WithEvents showButton As Infragistics.Win.Misc.UltraButton
    Friend WithEvents dateFromEditor As Infragistics.Win.UltraWinSchedule.UltraCalendarCombo
    Friend WithEvents dateToEditor As Infragistics.Win.UltraWinSchedule.UltraCalendarCombo
End Class
