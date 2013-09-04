<%@ Control Language="VB" Inherits="System.Web.Mvc.ViewUserControl" %>
<%@ Import Namespace="DMI.NET" %>

<link href="<%= Url.Content("~/Scripts/jquery/jOrgChart/css/bootstrap.min.css")%>" rel="stylesheet" />
<link href="<%= Url.Content("~/Scripts/jquery/jOrgChart/css/jquery.jOrgChart.css")%>" rel="stylesheet" />
<link href="<%= Url.Content("~/Scripts/jquery/jOrgChart/css/custom.css")%>" rel="stylesheet" />
<link href="<%= Url.Content("~/Scripts/jquery/jOrgChart/css/prettify.css")%>" rel="stylesheet" />

<script>
	$(document).ready(function () {

		//process the results into unordered list.

		$("#hiddenTags").find(":hidden").not("script").each(function () {
			var props = $(this).val().split("\t");
			//props[5] is the hierarchyLevel.
			if (props[5] == "1") {
				$('#root').append('<li>' + props[4] + '<p>' + props[0] + " " + props[1] + '</p><ul id="' + props[2] + '"></ul></li>');
			}

			if (props[5] >= "2") {
				$('#' + props[3]).append('<li>' + props[4] + '<p>' + props[0] + ' ' + props[1] + '</p><ul id="' + props[2] + '"></li>');
			}
		});

		$('#workframe').attr('overflow', 'auto');
		$("#org").jOrgChart({
			chartElement: '#chart',
			dragAndDrop: true
		});
	});
</script>

<ul id='org' style="display: none;">
	<li>Me
		<ul id="root"></ul>
	</li>
</ul>

<div id="hiddenTags">
	<%
		Const adStateOpen = 1
	
		Dim cmdThousandFindColumns = CreateObject("ADODB.Command")
		cmdThousandFindColumns.CommandText = "spASRIntOrgChart"
		cmdThousandFindColumns.CommandType = 4 ' Stored Procedure
		cmdThousandFindColumns.ActiveConnection = Session("databaseConnection")
		cmdThousandFindColumns.CommandTimeout = 180
		
		Dim prmRootID = cmdThousandFindColumns.CreateParameter("RootID", 3, 1)
		cmdThousandFindColumns.Parameters.Append(prmRootID)
		prmRootID.value = CleanNumeric(Session("TopLevelRecID"))	'"00000101"
	
		Err.Clear()
		Dim rstHierarchyRecords = cmdThousandFindColumns.Execute
		Dim sErrorDescription = ""
	
	
		If (Err.Number <> 0) Then
			sErrorDescription = "Error reading the link find records." & vbCrLf & Err.Description
		End If
	
		If Len(sErrorDescription) = 0 Then
			If rstHierarchyRecords.state = adStateOpen Then
				Dim inputString As String
				Do While Not rstHierarchyRecords.EOF
				
					inputString = CType((rstHierarchyRecords.fields(0).value & vbTab &
															 rstHierarchyRecords.fields(1).value & vbTab &
															 rstHierarchyRecords.fields(2).value & vbTab &
															 rstHierarchyRecords.fields(3).value & vbTab &
															 rstHierarchyRecords.fields(4).value & vbTab &
															 rstHierarchyRecords.fields(5).value & vbTab), String)
				
					Response.Write("<input type='hidden' value='" & inputString & "'/>" & vbCrLf)
				
					rstHierarchyRecords.moveNext()
				Loop
			End If
		End If
		
	%>
</div>

<div id="chart" class="orgChart"></div>


