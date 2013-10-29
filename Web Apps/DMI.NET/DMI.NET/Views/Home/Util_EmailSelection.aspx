<%@ Page Language="VB" Inherits="System.Web.Mvc.ViewPage" %>
<%@ Import Namespace="DMI.NET" %>

<script runat="server">
	Private Function GetEmailSelection() As String
		Dim emailSelectionHtmlTable As New StringBuilder 'Used to construct the (temporary) HTML table that will be transformed into a jQuery grid table
        
		'Get the records.
		Dim cmdDefSelRecords = CreateObject("ADODB.Command")
		cmdDefSelRecords.CommandText = "spASRIntGetEmailGroups"
		cmdDefSelRecords.CommandType = 4 'Stored Procedure
		cmdDefSelRecords.ActiveConnection = Session("databaseConnection")
		Err.Clear()
		Dim rstDefSelRecords = cmdDefSelRecords.Execute

		'Create an HTML table
		With emailSelectionHtmlTable
			.Append("<table id=""EmailSelectionTable"">")
			.Append("<tr>")
			.Append("<th id=""EmailGroupIDHeader"">EmailGroupID</th>")
			.Append("<th id=""NameHeader"">Name</th>")
			.Append("</tr>")
		End With
        
		'Populate the table
		Dim i As Integer = 1
		Do While Not rstDefSelRecords.EOF
			With emailSelectionHtmlTable
				.Append("<tr>")
				.Append("<td id='Row" & i & "'>" & rstDefSelRecords.Fields("emailGroupID").Value & "</td>")
				.Append("<td>" & rstDefSelRecords.Fields("name").Value.ToString.Replace("_", " ").Replace("""", "&quot;") & "</td>")
				.Append("</tr>")
				i += 1
			End With
			rstDefSelRecords.MoveNext()
		Loop
        
		emailSelectionHtmlTable.Append("</table>")
        
		rstDefSelRecords.close()
		rstDefSelRecords = Nothing
        
		' Release the ADO command object.
		cmdDefSelRecords = Nothing
        
		Return emailSelectionHtmlTable.ToString
	End Function
</script>


<!DOCTYPE html>
<html>
<head>
	<title>OpenHR Intranet</title>
	<script src="<%: Url.Content("~/bundles/jQuery")%>" type="text/javascript"></script>
	<script src="<%: Url.Content("~/bundles/jQueryUI7")%>" type="text/javascript"></script>
	<script src="<%: Url.Content("~/bundles/OpenHR_General")%>" type="text/javascript"></script>
	<script id="officebarscript" src="<%: Url.Content("~/Scripts/officebar/jquery.officebar.js") %>" type="text/javascript"></script>
	<link href="<%: Url.Content("~/Content/OpenHR.css") %>" rel="stylesheet" type="text/css" />
	<link href="<%: Url.LatestContent("~/Content/Site.css")%>" rel="stylesheet" type="text/css" />
	<link href="<%: Url.Content("~/Content/Site.css?v=8.0.8.0")%>" rel="stylesheet" type="text/css" />
	<link href="<%: Url.Content("~/Content/themes/Redmond/jquery-ui.min.css?v=8.0.8.0") %>" rel="stylesheet" type="text/css" />
	<link href="<%: Url.Content("~/Content/ui.jqgrid.css")%>" rel="stylesheet" type="text/css" />
	<link href="<%: Url.LatestContent("~/Content/OpenHR.css")%>" rel="stylesheet" type="text/css" />
	<link id="DMIthemeLink" href="<%: Url.LatestContent("~/Content/themes/" & Session("ui-theme").ToString() & "/jquery-ui.min.css")%>" rel="stylesheet" type="text/css" />
	<link href="<%= Url.LatestContent("~/Content/general_enclosed_foundicons.css")%>" rel="stylesheet" type="text/css" />
	<link href="<%= Url.LatestContent("~/Content/font-awesome.css")%>" rel="stylesheet" type="text/css" />
	<link href="<%= Url.LatestContent("~/Content/fonts/SSI80v194934/style.css")%>" rel="stylesheet" />

	<script type="text/javascript">
		window.onload = function () {
			var iResizeBy, iNewWidth, iNewHeight;

			// Resize the popup.
			iResizeBy = (bdyMain.scrollWidth - bdyMain.clientWidth);
			if (bdyMain.offsetWidth + iResizeBy > screen.width) {
				window.dialogWidth = new String(screen.width) + "px";
			} else {
				iNewWidth = new Number(window.dialogWidth.substr(0, window.dialogWidth.length - 2));
				iNewWidth = iNewWidth + iResizeBy;
				window.dialogWidth = new String(iNewWidth) + "px";
			}

			iResizeBy = bdyMain.scrollHeight - bdyMain.clientHeight;
			if (bdyMain.offsetHeight + iResizeBy > screen.height) {
				window.dialogHeight = new String(screen.height) + "px";
			} else {
				iNewHeight = new Number(window.dialogHeight.substr(0, window.dialogHeight.length - 2));
				iNewHeight = iNewHeight + iResizeBy;
				window.dialogHeight = new String(iNewHeight) + "px";
			}
		};

		function setForm() {
			var frmPopup = document.getElementById("frmPopup");
			window.dialogArguments.document.getElementById('txtEmailGroup').value = frmPopup.txtSelectedName.value;
			window.dialogArguments.document.getElementById('txtEmailGroupID').value = frmPopup.txtSelectedID.value;

			self.close();
			return false;
		}
	</script>
</head>

<body id="bdyMain" name="bdyMain" <%=session("BodyColour")%> leftmargin="20" topmargin="20" bottommargin="20" rightmargin="20">
	<form id="frmPopup" name="frmPopup" onsubmit="return setForm();" style="visibility: hidden; display: none">
		<input type="hidden" id="txtSelectedID" name="txtSelectedID">
		<input type="hidden" id="txtSelectedName" name="txtSelectedName">
		<input type="hidden" id="txtSelectedAccess" name="txtSelectedAccess">
		<input type="hidden" id="txtSelectedUserName" name="txtSelectedUserName">
	</form>

	<table align="center" class="outline" cellpadding="5" cellspacing="0" width="100%" height="100%">
		<tr>
			<td>
				<table width="100%" height="100%" class="invisible" cellspacing="0" cellpadding="0">
					<tr height="10">
						<td colspan="3" align="center" height="10">
							<h3>Email Groups</h3>
						</td>
					</tr>
					<tr>
						<td width="20"></td>
						<td>
							<%=GetEmailSelection()%>
						</td>
						<td width="20"></td>
					</tr>
					<tr height="10">
						<td height="10" colspan="3">&nbsp;</td>
					</tr>
					<tr height="10">
						<td width="20"></td>
						<td height="10">
							<table width="100%" class="invisible" cellspacing="0" cellpadding="0">
								<tr>
									<td>&nbsp;</td>
									<td width="10">
										<input id="cmdok" type="button" value="OK" name="cmdok" 
											style="width: 80px"
											class="button ui-button ui-widget ui-state-default ui-widget-content ui-corner-tl ui-corner-br" 
										/>
									</td>
									<td width="10">&nbsp;</td>
									<td width="10">
										<input id="cmdnone" type="button" value="None" name="cmdnone"
											style="width: 80px"
											class="button ui-button ui-widget ui-state-default ui-widget-content ui-corner-tl ui-corner-br" 
											onclick="frmPopup.txtSelectedID.value = 0; frmPopup.txtSelectedName.value = ''; frmPopup.txtSelectedAccess.value = ''; frmPopup.txtSelectedUserName.value = ''; setForm();" />
									</td>
									<td width="10">&nbsp;</td>
									<td width="10">
										<input id="cmdcancel" type="button" value="Cancel" name="cmdcancel" 
											style="width: 80px"
											class="button ui-button ui-widget ui-state-default ui-widget-content ui-corner-tl ui-corner-br" 
											onclick="self.close();" />
									</td>
								</tr>
							</table>
						</td>
					</tr>
				</table>
			</td>
		</tr>
	</table>

	<form id="frmFromOpener" name="frmFromOpener" style="visibility: hidden; display: none">
		<input type="hidden" id="calcEmailCurrentID" name="calcEmailCurrentID" value='<%= Request("emailSelCurrentID") %>'>
	</form>

	<input type="hidden" id="txtTicker" name="txtTicker" value="0">
	<input type="hidden" id="txtLastKeyFind" name="txtLastKeyFind" value="">
</body>

<script type="text/javascript">
	// Table to jQuery grid
	tableToGrid("#EmailSelectionTable", {
		onSelectRow: function (rowID) {
		},
		ondblClickRow: function (rowID) {
		},
		rowNum: 1000
	});

	//Hide the EmailGroup table header and its column
	$('.ui-jqgrid-htable tr th:nth-child(1)').hide();
	$('#EmailSelectionTable tr td:nth-child(1)').hide();

	//On clicking "Ok", set the values
	$("#cmdok").click(function () {
		//Get the selected value (ID) and name and assign them to the input tags that will be used by setForm() to set the values in the parent window
		var emailSelectedId = $("#EmailSelectionTable .ui-state-highlight [aria-describedby=EmailSelectionTable_EmailGroupIDHeader]").html();
		var emailSelectedName = $("#EmailSelectionTable .ui-state-highlight [aria-describedby=EmailSelectionTable_NameHeader]").html();
		//debugger;
		$("#txtSelectedID").val(emailSelectedId);
		$("#txtSelectedName").val(emailSelectedName);

		setForm();
	});
</script>

</html>
