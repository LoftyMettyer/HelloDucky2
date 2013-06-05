Public Class AuditLog

  Public Database As SystemFramework.HCM

  Private Sub AuditLog_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

    Dim obj As New DataSet

    grdAudit.DataSource = Database.GetAuditLogDataSource
    '    grdAudit.DataSource = Database.GetAuditLogDescriptions

  End Sub

  Private Sub grdAudit_InitializeLayout1(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs) Handles grdAudit.InitializeLayout
    e.Layout.Override.FilterUIType = Infragistics.Win.UltraWinGrid.FilterUIType.HeaderIcons

  End Sub

  Private Sub butOutput_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles butOutput.Click

    UltraGridExcelExporter1.Export(grdAudit, txtFilePath.Text)

  End Sub
End Class