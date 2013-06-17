Public Class WebForm1
    Inherits System.Web.UI.Page

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

      If Not IsPostBack Then

         'Dim db As New Database(App.Config.ConnectionString)

         'Dim result = db.GetWorkflowItemValues(46799, 1718)

         'GridView1.DataSource = result.Data
         'GridView1.DataBind()
      End If
      SqlDataSource1.ConnectionString = App.Config.ConnectionString
      SqlDataSource1.SelectCommand = "SELECT * FROM tbuser_Personnel_Records"

    End Sub

End Class