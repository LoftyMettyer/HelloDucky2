<%@ Control Language="VB" Inherits="System.Web.Mvc.ViewUserControl" %>
<%@ Import Namespace="DMI.NET" %>

		<script type="text/javascript">
				function reportdata_window_onload() {

				<%
				
				If Session("CR_Mode") = "" Then
						Response.Write("  customreport_loadAddRecords();" & vbCrLf & vbCrLf)
				Else
						Response.Write("  ExportData('OUTPUTREPORT');" & vbCrLf)
				End If
				%>
				}
		</script>

		<%	
				If CStr(Session("EmailGroupID")) = "" Then
						Session("EmailGroupID") = 0
				End If
		
			Dim cmdReportsCols As ADODB.Command
			Dim prmEmailGroupID As ADODB.Parameter
			Dim rstReportColumns As ADODB.Recordset
			Dim sErrorDescription As String = ""
			Dim iLoop As Integer
		
		If Session("EmailGroupID") > 0 Then
				cmdReportsCols = New ADODB.Command()
				cmdReportsCols.CommandText = "spASRIntGetEmailGroupAddresses"
				cmdReportsCols.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
				cmdReportsCols.ActiveConnection = Session("databaseConnection")

				prmEmailGroupID = cmdReportsCols.CreateParameter("EmailGroupID", 3, 1) ' 3=integer, 1=input
				cmdReportsCols.Parameters.Append(prmEmailGroupID)
				prmEmailGroupID.value = cleanNumeric(Session("EmailGroupID"))

				Err.Clear()
				rstReportColumns = cmdReportsCols.Execute

				If (Err.Number <> 0) Then
						sErrorDescription = "Error getting the email addresses for group." & vbCrLf & formatError(Err.Description)
				End If

				If Len(sErrorDescription) = 0 Then
						iLoop = 1
						Response.Write("<INPUT id=txtEmailGroupAddr name=txtEmailGroupAddr class=""text"" value=""")
						Do While Not rstReportColumns.EOF
								If iLoop > 1 Then
										Response.Write(";")
								End If
								Response.Write(Replace(rstReportColumns.Fields("Fixed").Value, """", "&quot;"))
								rstReportColumns.MoveNext()
								iLoop = iLoop + 1
						Loop
						Response.Write(""">" & vbCrLf)

						' Release the ADO recordset object.
						rstReportColumns.close()
				End If
				
				rstReportColumns = Nothing
				cmdReportsCols = Nothing

		Else
				Response.Write("<INPUT id=txtEmailGroupAddr name=txtEmailGroupAddr class=""text"" value="""">" & vbCrLf)
		End If
%>
	

<form action="util_run_customreportsDataSubmit" method="post" id="frmGetReportData" name="frmGetReportData">
		<input id="txtMode" name="txtMode" class="text" value="<%=Session("CR_Mode")%>">
		<input id="txtEmailGroupID" name="txtEmailGroupID" class="text" value="<%=Session("EmailGroupID")%>">
</form>

<script type="text/javascript">
	reportdata_window_onload();
	$(".popup").dialog('option', 'title', $("#txtDefn_Name").val());
	$("#PageDivTitle").html($("#txtDefn_Name").val());
</script>
