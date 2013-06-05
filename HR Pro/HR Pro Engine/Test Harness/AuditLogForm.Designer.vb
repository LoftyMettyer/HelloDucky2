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
        Dim DateButton1 As Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton = New Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton()
        Dim DateButton2 As Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton = New Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton()
        Dim Appearance14 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance3 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance("Highlight", 74988782)
        Me.datePanel = New System.Windows.Forms.Panel()
        Me.dateLabel = New Infragistics.Win.Misc.UltraLabel()
        Me.dateFromEditor = New Infragistics.Win.UltraWinSchedule.UltraCalendarCombo()
        Me.dateToEditor = New Infragistics.Win.UltraWinSchedule.UltraCalendarCombo()
        Me.findButton = New Infragistics.Win.Misc.UltraButton()
        Me.periodEditor = New Infragistics.Win.UltraWinEditors.UltraComboEditor()
        Me.userEditor = New Infragistics.Win.UltraWinEditors.UltraComboEditor()
        Me.txtFilePath = New System.Windows.Forms.TextBox()
        Me.butOutput = New System.Windows.Forms.Button()
        Me.auditLogsGrid = New Infragistics.Win.UltraWinGrid.UltraGrid()
        Me.gridExporter = New Infragistics.Win.UltraWinGrid.ExcelExport.UltraGridExcelExporter(Me.components)
        Me.UltraLabel1 = New Infragistics.Win.Misc.UltraLabel()
        Me.UltraLabel2 = New Infragistics.Win.Misc.UltraLabel()
        Me.datePanel.SuspendLayout()
        CType(Me.dateFromEditor, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.dateToEditor, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.periodEditor, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.userEditor, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.auditLogsGrid, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'datePanel
        '
        Me.datePanel.Controls.Add(Me.dateLabel)
        Me.datePanel.Controls.Add(Me.dateFromEditor)
        Me.datePanel.Controls.Add(Me.dateToEditor)
        Me.datePanel.Location = New System.Drawing.Point(188, 10)
        Me.datePanel.Name = "datePanel"
        Me.datePanel.Size = New System.Drawing.Size(261, 23)
        Me.datePanel.TabIndex = 1
        Me.datePanel.Visible = False
        '
        'dateLabel
        '
        Me.dateLabel.AutoSize = True
        Me.dateLabel.Location = New System.Drawing.Point(108, 3)
        Me.dateLabel.Name = "dateLabel"
        Me.dateLabel.Size = New System.Drawing.Size(23, 14)
        Me.dateLabel.TabIndex = 1
        Me.dateLabel.Text = "and"
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
        Me.dateFromEditor.TabIndex = 0
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
        Me.dateToEditor.TabIndex = 2
        Me.dateToEditor.Value = ""
        '
        'findButton
        '
        Me.findButton.Location = New System.Drawing.Point(188, 37)
        Me.findButton.Name = "findButton"
        Me.findButton.Size = New System.Drawing.Size(58, 23)
        Me.findButton.TabIndex = 3
        Me.findButton.Text = "Show"
        '
        'periodEditor
        '
        Me.periodEditor.DropDownStyle = Infragistics.Win.DropDownStyle.DropDownList
        Me.periodEditor.LimitToList = True
        Me.periodEditor.Location = New System.Drawing.Point(49, 10)
        Me.periodEditor.Name = "periodEditor"
        Me.periodEditor.Size = New System.Drawing.Size(123, 21)
        Me.periodEditor.TabIndex = 0
        Me.periodEditor.ValueMember = ""
        '
        'userEditor
        '
        Me.userEditor.DropDownStyle = Infragistics.Win.DropDownStyle.DropDownList
        Me.userEditor.LimitToList = True
        Me.userEditor.Location = New System.Drawing.Point(49, 37)
        Me.userEditor.Name = "userEditor"
        Me.userEditor.Size = New System.Drawing.Size(123, 21)
        Me.userEditor.TabIndex = 2
        Me.userEditor.ValueMember = ""
        '
        'txtFilePath
        '
        Me.txtFilePath.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.txtFilePath.Location = New System.Drawing.Point(13, 380)
        Me.txtFilePath.Name = "txtFilePath"
        Me.txtFilePath.Size = New System.Drawing.Size(273, 20)
        Me.txtFilePath.TabIndex = 5
        '
        'butOutput
        '
        Me.butOutput.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.butOutput.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.butOutput.Location = New System.Drawing.Point(291, 378)
        Me.butOutput.Name = "butOutput"
        Me.butOutput.Size = New System.Drawing.Size(66, 23)
        Me.butOutput.TabIndex = 6
        Me.butOutput.Text = "Export"
        Me.butOutput.UseVisualStyleBackColor = True
        '
        'auditLogsGrid
        '
        Me.auditLogsGrid.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Appearance14.BackColor = System.Drawing.SystemColors.Window
        Me.auditLogsGrid.DisplayLayout.Appearance = Appearance14
        Appearance3.BackColor = System.Drawing.SystemColors.Highlight
        Appearance3.ForeColor = System.Drawing.SystemColors.HighlightText
        Me.auditLogsGrid.DisplayLayout.Appearances.Add(Appearance3)
        Me.auditLogsGrid.DisplayLayout.AutoFitStyle = Infragistics.Win.UltraWinGrid.AutoFitStyle.ResizeAllColumns
        Me.auditLogsGrid.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No
        Me.auditLogsGrid.DisplayLayout.Override.AllowColSwapping = Infragistics.Win.UltraWinGrid.AllowColSwapping.NotAllowed
        Me.auditLogsGrid.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.[False]
        Me.auditLogsGrid.DisplayLayout.Override.AllowGroupBy = Infragistics.Win.DefaultableBoolean.[True]
        Me.auditLogsGrid.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.[False]
        Me.auditLogsGrid.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.RowSelect
        Me.auditLogsGrid.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti
        Me.auditLogsGrid.DisplayLayout.Override.RowSelectors = Infragistics.Win.DefaultableBoolean.[False]
        Me.auditLogsGrid.DisplayLayout.Override.SelectTypeCell = Infragistics.Win.UltraWinGrid.SelectType.None
        Me.auditLogsGrid.DisplayLayout.Override.SelectTypeCol = Infragistics.Win.UltraWinGrid.SelectType.None
        Me.auditLogsGrid.DisplayLayout.Override.SelectTypeGroupByRow = Infragistics.Win.UltraWinGrid.SelectType.None
        Me.auditLogsGrid.DisplayLayout.Override.SelectTypeRow = Infragistics.Win.UltraWinGrid.SelectType.[Single]
        Me.auditLogsGrid.DisplayLayout.ScrollStyle = Infragistics.Win.UltraWinGrid.ScrollStyle.Immediate
        Me.auditLogsGrid.DisplayLayout.TabNavigation = Infragistics.Win.UltraWinGrid.TabNavigation.NextControlOnLastCell
        Me.auditLogsGrid.Location = New System.Drawing.Point(1, 66)
        Me.auditLogsGrid.Margin = New System.Windows.Forms.Padding(6, 5, 6, 5)
        Me.auditLogsGrid.Name = "auditLogsGrid"
        Me.auditLogsGrid.Size = New System.Drawing.Size(809, 306)
        Me.auditLogsGrid.TabIndex = 4
        '
        'UltraLabel1
        '
        Me.UltraLabel1.AutoSize = True
        Me.UltraLabel1.Location = New System.Drawing.Point(13, 41)
        Me.UltraLabel1.Name = "UltraLabel1"
        Me.UltraLabel1.Size = New System.Drawing.Size(31, 14)
        Me.UltraLabel1.TabIndex = 7
        Me.UltraLabel1.Text = "User:"
        '
        'UltraLabel2
        '
        Me.UltraLabel2.AutoSize = True
        Me.UltraLabel2.Location = New System.Drawing.Point(12, 14)
        Me.UltraLabel2.Name = "UltraLabel2"
        Me.UltraLabel2.Size = New System.Drawing.Size(31, 14)
        Me.UltraLabel2.TabIndex = 8
        Me.UltraLabel2.Text = "Date:"
        '
        'AuditLogForm
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(809, 411)
        Me.Controls.Add(Me.UltraLabel2)
        Me.Controls.Add(Me.UltraLabel1)
        Me.Controls.Add(Me.txtFilePath)
        Me.Controls.Add(Me.butOutput)
        Me.Controls.Add(Me.auditLogsGrid)
        Me.Controls.Add(Me.userEditor)
        Me.Controls.Add(Me.periodEditor)
        Me.Controls.Add(Me.findButton)
        Me.Controls.Add(Me.datePanel)
        Me.Name = "AuditLogForm"
        Me.Text = "Audit Log"
        Me.datePanel.ResumeLayout(False)
        Me.datePanel.PerformLayout()
        CType(Me.dateFromEditor, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.dateToEditor, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.periodEditor, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.userEditor, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.auditLogsGrid, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents datePanel As System.Windows.Forms.Panel
    Friend WithEvents dateLabel As Infragistics.Win.Misc.UltraLabel
    Friend WithEvents findButton As Infragistics.Win.Misc.UltraButton
    Friend WithEvents dateFromEditor As Infragistics.Win.UltraWinSchedule.UltraCalendarCombo
    Friend WithEvents dateToEditor As Infragistics.Win.UltraWinSchedule.UltraCalendarCombo
    Friend WithEvents periodEditor As Infragistics.Win.UltraWinEditors.UltraComboEditor
    Friend WithEvents userEditor As Infragistics.Win.UltraWinEditors.UltraComboEditor
    Friend WithEvents txtFilePath As System.Windows.Forms.TextBox
    Friend WithEvents butOutput As System.Windows.Forms.Button
    Friend WithEvents auditLogsGrid As Infragistics.Win.UltraWinGrid.UltraGrid
    Friend WithEvents gridExporter As Infragistics.Win.UltraWinGrid.ExcelExport.UltraGridExcelExporter
    Friend WithEvents UltraLabel1 As Infragistics.Win.Misc.UltraLabel
    Friend WithEvents UltraLabel2 As Infragistics.Win.Misc.UltraLabel
End Class
