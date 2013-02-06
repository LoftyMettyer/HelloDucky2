<%@ Control Language="VB" Inherits="System.Web.Mvc.ViewUserControl" %>
<%@ Import Namespace="DMI.NET" %>

    <script type="text/javascript">
        function reportdata_window_onload() {
            $("#workframe").attr("data-framesource", "UTIL_RUN_CUSTOMREPORTS_DATA");

        <%
        If Session("CR_Mode") = "" Then
            Response.Write("  loadAddRecords();" & vbCrLf & vbCrLf)
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
		
    Dim cmdReportsCols
    Dim prmEmailGroupID
    Dim rstReportColumns
    Dim sErrorDescription As String
    Dim iLoop As Integer
    
    If Session("EmailGroupID") > 0 Then
        cmdReportsCols = CreateObject("ADODB.Command")
        cmdReportsCols.CommandText = "spASRIntGetEmailGroupAddresses"
        cmdReportsCols.CommandType = 4 ' Stored procedure
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
    
    <script type="text/javascript">
        function ExportData(strMode) {
            ExportData("OUTPUTREPORT");
            return;
        }
    </script>
    

    <FORM action="util_run_customreportsDataSubmit" method=post id=frmGetData name=frmGetData>
		<INPUT id=txtMode name=txtMode class="text" value="<%=Session("CR_Mode")%>">
		<INPUT id=txtEmailGroupID name=txtEmailGroupID class="text" value="<%=Session("EmailGroupID")%>">
	</FORM>
    
    <script type="text/javascript">
        reportdata_window_onload();
    </script>
