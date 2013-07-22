Imports System.Net
Imports System.IO

Public Class SalarySummary
		Inherits System.Web.UI.Page

		Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
			Dim request As WebRequest = WebRequest.Create("http://10.7.33.19:6680/openpeopleapi/payslips/108/summary") 'Hardcoded
			request.Credentials = New NetworkCredential("fake", "fake") 'Any credentials will do

			Dim wResponse As WebResponse = request.GetResponse() 'Get the response
			Dim dataStream As Stream = wResponse.GetResponseStream()
			Dim reader As New StreamReader(dataStream)
			Dim responseFromServer As String = "<?xml version=""1.0"" encoding=""utf-8"" ?>" & reader.ReadToEnd()
			Dim stringReader As New StringReader(responseFromServer)
			Dim ds As New DataSet
			ds.ReadXml(stringReader) 'Read the XML into a dataset

			reader.Close()
			wResponse.Close()

			'Add rows to the HTML table; crude, I know but it's only a prototype!
			Dim tr As TableRow
			Dim tc As TableCell
			Dim r As DataRow = ds.Tables(0).Rows(0)
			tr = New TableRow()
			tr.CssClass = "alt"

			tc = New TableCell()
			tc.Text = "Gross Pay"
			tr.Cells.Add(tc)

			tc = New TableCell()
			tc.Text = "£" & r(0)
			tr.Cells.Add(tc)

			SalarySummaryTable.Rows.Add(tr)

			tr = New TableRow()

			tc = New TableCell()
			tc.Text = "Net Pay"
			tr.Cells.Add(tc)

			tc = New TableCell()
			tc.Text = "£" & r(1)
			tr.Cells.Add(tc)

			SalarySummaryTable.Rows.Add(tr)

			tr = New TableRow()
			tr.CssClass = "alt"

			tc = New TableCell()
			tc.Text = "Tax"
			tr.Cells.Add(tc)

			tc = New TableCell()
			tc.Text = "£" & r(2)
			tr.Cells.Add(tc)

			SalarySummaryTable.Rows.Add(tr)

		End Sub
End Class