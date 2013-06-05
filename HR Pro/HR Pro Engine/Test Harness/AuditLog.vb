Public Class AuditLog

  Public Database As SystemFramework.HCM

  Private Sub AuditLog_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Dim db As New DBContextDataContext

        Dim data = From a In db.ASRSysAuditTrails
                   Select a

        Dim sw As New System.Diagnostics.Stopwatch

        sw.Start()

        grdAudit.DataSource = data.ToList()

        Console.WriteLine(sw.ElapsedMilliseconds)

  End Sub

    Private Sub grdAudit_InitializeLayout1(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs) Handles grdAudit.InitializeLayout

        e.Layout.Override.FilterUIType = Infragistics.Win.UltraWinGrid.FilterUIType.HeaderIcons

    End Sub

  Private Sub butOutput_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles butOutput.Click

        UltraGridExcelExporter1.Export(grdAudit, txtFilePath.Text)

  End Sub
End Class