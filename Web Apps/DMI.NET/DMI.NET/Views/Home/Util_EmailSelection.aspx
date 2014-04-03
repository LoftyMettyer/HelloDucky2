<%@ Page Language="VB" Inherits="System.Web.Mvc.ViewPage" %>
<%@ Import Namespace="DMI.NET" %>
<%@ Import Namespace="HR.Intranet.Server" %>
<%@ Import Namespace="System.Data.SqlClient" %>
<%@ Import Namespace="System.Data" %>

<script runat="server">
	Private Function GetEmailSelection() As String
		Dim emailSelectionHtmlTable As New StringBuilder 'Used to construct the (temporary) HTML table that will be transformed into a jQuery grid table
		Dim objSession As SessionInfo = CType(Session("SessionContext"), SessionInfo)	'Set session info
		Dim objDataAccess As New clsDataAccess(objSession.LoginInfo) 'Instantiate DataAccess class
		
		'Get the records.
		Dim rstDefSelRecords As DataTable = objDataAccess.GetDataTable("spASRIntGetEmailGroups", CommandType.StoredProcedure)

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
		For Each r As DataRow In rstDefSelRecords.Rows
			With emailSelectionHtmlTable
				.Append("<tr>")
				.Append("<td id='Row" & i & "'>" & r("emailGroupID").ToString & "</td>")
				.Append("<td>" & r("name").ToString.Replace("_", " ").Replace("""", "&quot;") & "</td>")
				.Append("</tr>")
				i += 1
			End With
		Next

		emailSelectionHtmlTable.Append("</table>")

		Return emailSelectionHtmlTable.ToString
	End Function
</script>


<!DOCTYPE html>
<html>
<head>
	<title>OpenHR</title>
	<script src="<%: Url.LatestContent("~/bundles/jQuery")%>" type="text/javascript"></script>
	<script src="<%: Url.LatestContent("~/bundles/jQueryUI7")%>" type="text/javascript"></script>
	<script src="<%: Url.LatestContent("~/bundles/Microsoft")%>" type="text/javascript"></script>
	<script src="<%: Url.LatestContent("~/bundles/OpenHR_General")%>" type="text/javascript"></script>

	<link id="DMIthemeLink" href="<%: Url.LatestContent("~/Content/themes/" & Session("ui-admin-theme").ToString() & "/jquery-ui.min.css")%>" rel="stylesheet" type="text/css" />
	<script id="officebarscript" src="<%: Url.LatestContent("~/Scripts/officebar/jquery.officebar.js")%>" type="text/javascript"></script>

	<script src="<%: Url.LatestContent("~/Scripts/ctl_SetStyles.js")%>" type="text/javascript"></script>
	<link href="<%: Url.LatestContent("~/Content/OpenHR.css")%>" rel="stylesheet" type="text/css" />
	<link href="<%: Url.LatestContent("~/Content/Site.css")%>" rel="stylesheet" type="text/css" />
	<link href="<%: Url.LatestContent("~/Content/OpenHR.css")%>" rel="stylesheet" type="text/css" />
	<link href="<%= Url.LatestContent("~/Content/general_enclosed_foundicons.css")%>" rel="stylesheet" type="text/css" />
	<link href="<%= Url.LatestContent("~/Content/font-awesome.css")%>" rel="stylesheet" type="text/css" />

	<%--jQuery Grid Stylesheet--%>
	<link href="<%: Url.LatestContent("~/Content/ui.jqgrid.css")%>" rel="stylesheet" type="text/css" />

	<script type="text/javascript">
		window.onload = function () {
			//Get some cookies that we need to determine the CSS to apply
			var SSIMode = OpenHR.getCookie("SSIMode");
			var currentLayout = OpenHR.getCookie("currentLayout");
			var currentTheme = OpenHR.getCookie("currentTheme");
			var cookiewireframeTheme = OpenHR.getCookie("cookiewireframeTheme");
			var cookieapplyWireframeTheme = OpenHR.getCookie("cookieapplyWireframeTheme");

			if (($("#fixedlinksframe").length > 0) && (currentLayout != "winkit"))
				$("link[id=DMIthemeLink]").attr({ href: "" });

			//The logic below is taken from Site.Master, it should be abstracted somewhere else, but no time to do that now
			if (SSIMode != "True") {
				$("link[id=layoutLink]").attr({ href: "<%:Url.LatestContent("~/Content/DashboardStyles/layouts/winkit.css")%>" });
				$("link[id=themeLink]").attr({ href: "<%:Url.LatestContent("~/Content/DashboardStyles/themes/white.css")%>" });
				$('body').addClass('DMI');
			} else {
				switch (OpenHR.getCookie("Intranet_Layout")) {
				case "winkit":
					$("link[id=layoutLink]").attr({ href: "<%:Url.LatestContent("~/Content/DashboardStyles/layouts/winkit.css")%>" });
					$("link[id=SSIthemeLink]").attr({ href: "<%:Url.LatestContent("~/Content/themes/redmond/jquery-ui.min.css")%>" });
					$("link[id=DMIthemeLink]").attr({ href: "<%:Url.LatestContent("~/Content/themes/redmond/jquery-ui.min.css")%>" });
					break;
				case "wireframe":
					if (cookieapplyWireframeTheme == "true") $("link[id=WireframethemeLink]").attr({ href: "../Content/DashboardStyles/themes/upgraded.css" });

					$("link[id=layoutLink]").attr({ href: "<%:Url.LatestContent("~/Content/DashboardStyles/layouts/wireframe.css")%>" });
					$("link[id=SSIthemeLink]").attr({ href: "../Content/themes/" + cookiewireframeTheme + "/jquery-ui.min.css" });
					$("link[id=DMIthemeLink]").attr({ href: "../Content/themes/" + cookiewireframeTheme + "/jquery-ui.min.css" });
					break;
				case "tiles":
					$("link[id=layoutLink]").attr({ href: "<%:Url.LatestContent("~/Content/DashboardStyles/layouts/tiles.css")%>" });
					$("link[id=SSIthemeLink]").attr({ href: "<%:Url.LatestContent("~/Content/themes/jMetro/jquery-ui.min.css")%>" });
					$("link[id=DMIthemeLink]").attr({ href: "<%:Url.LatestContent("~/Content/themes/jMetro/jquery-ui.min.css")%>" });
					break;
				}

				switch (currentTheme) {
				case "red":
					$("link[id=themeLink]").attr({ href: "<%: Url.LatestContent("~/Content/DashboardStyles/themes/Red.css")%>" });
					break;
				case "blue":
					$("link[id=themeLink]").attr({ href: "<%: Url.LatestContent("~/Content/DashboardStyles/themes/Blue.css")%>" });
					break;
				case "white":
					$("link[id=themeLink]").attr({ href: "<%: Url.LatestContent("~/Content/DashboardStyles/themes/White.css")%>" });
				default:
					break;
				}
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

	<div style="text-align: center">
		<h3>Email Groups</h3>
 	</div>
	<div style="margin-left: 15px;">
		<%=GetEmailSelection()%>
	</div>
	<div style="margin-top: 10px; margin-right: 10px; float: right;">
			<input id="cmdok" type="button" value="OK" name="cmdok" 
					style="width: 80px"
					class="button ui-button ui-widget ui-state-default ui-widget-content ui-corner-tl ui-corner-br" />
			<input id="cmdnone" type="button" value="None" name="cmdnone"
					style="width: 80px"
					class="button ui-button ui-widget ui-state-default ui-widget-content ui-corner-tl ui-corner-br" 
					onclick="frmPopup.txtSelectedID.value = 0; frmPopup.txtSelectedName.value = ''; frmPopup.txtSelectedAccess.value = ''; frmPopup.txtSelectedUserName.value = ''; setForm();" />
			<input id="cmdcancel" type="button" value="Cancel" name="cmdcancel" 
					style="width: 80px"
					class="button ui-button ui-widget ui-state-default ui-widget-content ui-corner-tl ui-corner-br" 
					onclick="self.close();" />
	</div>

	<form id="frmFromOpener" name="frmFromOpener" style="visibility: hidden; display: none">
		<input type="hidden" id="calcEmailCurrentID" name="calcEmailCurrentID" value='<%= Request("emailSelCurrentID") %>'>
	</form>

	<input type="hidden" id="txtTicker" name="txtTicker" value="0">
	<input type="hidden" id="txtLastKeyFind" name="txtLastKeyFind" value="">
</body>

<script type="text/javascript">
	// Table to jQuery grid
	tableToGrid("#EmailSelectionTable", {
		colNames: ['EmailGroupIDHeader', 'Name'],
		colModel: [
			{ name: 'EmailGroupIDHeader', hidden: true },
			{ name: 'NameHeader', sortable: true }
		],
		rowNum: 1000,   //TODO set this to blocksize...
		height: 320,
		width: (screen.width) / 3 + 5,
		scrollerbar: true
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
