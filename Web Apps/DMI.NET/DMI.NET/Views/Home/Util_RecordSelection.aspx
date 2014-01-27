<%@ Page Language="VB" Inherits="System.Web.Mvc.ViewPage" %>

<%@ Import Namespace="DMI.NET" %>
<%@ Import Namespace="ADODB" %>

<script runat="server">
	Private _RecordSelectionHTMLTable As New StringBuilder	'Used to construct the (temporary) HTML table that will be transformed into a jQuey grid table
	
	Private Sub Page_Load(sender As Object, e As EventArgs) Handles Me.Load
		Dim cmdDefSelRecords As New Command
		
		cmdDefSelRecords.CommandText = "spASRIntGetRecordSelection"
		cmdDefSelRecords.CommandType = CommandTypeEnum.adCmdStoredProc
		cmdDefSelRecords.ActiveConnection = Session("databaseConnection")

		Dim prmType = cmdDefSelRecords.CreateParameter("type", DataTypeEnum.adVarChar, ParameterDirectionEnum.adParamInput, 8000)
		cmdDefSelRecords.Parameters.Append(prmType)
		prmType.Value = Request("recseltype")

		Dim prmTableID = cmdDefSelRecords.CreateParameter("tableID", DataTypeEnum.adInteger, ParameterDirectionEnum.adParamInput)
		cmdDefSelRecords.Parameters.Append(prmTableID)
		prmTableID.Value = CleanNumeric(Request("recseltableID"))

		Err.Clear()
		Dim rstDefSelRecords As Recordset = cmdDefSelRecords.Execute
		
		Dim IDFieldName As String = "" 'The name of the ID field varies depending on the recseltype; check spASRIntGetRecordSelection for a possible list of values
		Select Case UCase(Request("recseltype"))
			Case "PICKLIST"
				IDFieldName = "picklistid"
			Case "FILTER", "CALC"
				IDFieldName = "exprid"
			Case Else
				Throw New Exception("Exception in Util_RecordSelection.aspx: Unable to determine name of ID field for record selection; the 'IF UPPER(@psType) = ...' in spASRIntGetRecordSelection must match the Select Case")
		End Select
		
		With _RecordSelectionHTMLTable
			.Append("<table id='RecordSelectionHTMLTable'>")
			.Append("<tr>")
			.Append("<th id='ID'>ID</th>")
			.Append("<th id='Name'>Name</th>")
			.Append("<th id='UserName'>UserName</th>")
			.Append("<th id='Access'>Access</th>")
			.Append("</tr>")
		End With
		'Loop over the records
		Do Until rstDefSelRecords.EOF
			With _RecordSelectionHTMLTable
				.Append("<tr>")
				.Append("<td>" & rstDefSelRecords.Fields(IDFieldName).Value & "</td>")
				.Append("<td>" & rstDefSelRecords.Fields("name").Value & "</td>")
				.Append("<td>" & rstDefSelRecords.Fields("username").Value & "</td>")
				.Append("<td>" & rstDefSelRecords.Fields("access").Value & "</td>")
				.Append("</tr>")
			End With
			rstDefSelRecords.MoveNext()
		Loop
            
		_RecordSelectionHTMLTable.Append("</table>")
     
		rstDefSelRecords.Close()
		rstDefSelRecords = Nothing
	End Sub
</script>


<!DOCTYPE html>
<html>
<head>
	<title>OpenHR Intranet</title>
	<%--External script resources--%>
	<script src="<%: Url.Content("~/bundles/jQuery")%>" type="text/javascript"></script>
	<script src="<%: Url.Content("~/bundles/jQueryUI7")%>" type="text/javascript"></script>
	<script src="<%: Url.Content("~/bundles/Microsoft")%>" type="text/javascript"></script>
	<script src="<%: Url.Content("~/bundles/OpenHR_General")%>" type="text/javascript"></script>

	<link id="DMIthemeLink" href="<%: Url.LatestContent("~/Content/themes/" & Session("ui-theme").ToString() & "/jquery-ui.min.css")%>" rel="stylesheet" type="text/css" />
	<script id="officebarscript" src="<%: Url.Content("~/Scripts/officebar/jquery.officebar.js") %>" type="text/javascript"></script>

	<script src="<%: Url.Content("~/Scripts/ctl_SetStyles.js") %>" type="text/javascript"></script>
	<link href="<%: Url.Content("~/Content/OpenHR.css") %>" rel="stylesheet" type="text/css" />
	<link href="<%: Url.LatestContent("~/Content/Site.css")%>" rel="stylesheet" type="text/css" />
	<link href="<%: Url.LatestContent("~/Content/OpenHR.css")%>" rel="stylesheet" type="text/css" />
	<link href="<%= Url.LatestContent("~/Content/general_enclosed_foundicons.css")%>" rel="stylesheet" type="text/css" />
	<link href="<%= Url.LatestContent("~/Content/font-awesome.css")%>" rel="stylesheet" type="text/css" />

	<%--jQuery Grid Stylesheet--%>
	<link href="<%: Url.LatestContent("~/Content/ui.jqgrid.css")%>" rel="stylesheet" type="text/css" />
</head>

<body id="bdyMain" bgcolor='<%=session("ConvertedDesktopColour")%>' leftmargin="20" topmargin="20" bottommargin="20" rightmargin="5">
	
<div style="margin-left: 10px; margin-right: 10px;">
	<form id="frmPopup" name="frmPopup" onsubmit="return setForm();" style="visibility: hidden; display: none">
		<input type="hidden" id="txtSelectedID" name="txtSelectedID">
		<input type="hidden" id="txtSelectedName" name="txtSelectedName">
		<input type="hidden" id="txtSelectedAccess" name="txtSelectedAccess">
		<input type="hidden" id="txtSelectedUserName" name="txtSelectedUserName">
	</form>

	<h3 style="text-align: center;">
		<% 
			If Request("recseltype") = "picklist" Then
				Response.Write("Picklists")
			Else
				Response.Write("Filters")
			End If
		%>
	</h3>
	<p>
		<%Response.Write(_RecordSelectionHTMLTable.ToString())%>
	</p>
	<table style="width: 100%" class="invisible">
		<tr>
			<td>&nbsp;</td>
			<td width="10">
				<input id="cmdok" type="button" value="OK" name="cmdok" 
					style="width: 80px"
					class="button ui-button ui-widget ui-state-default ui-widget-content ui-corner-tl ui-corner-br" 
					onclick="setForm();" />
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

</div>

	<form id="frmFromOpener" name="frmFromOpener" style="visibility: hidden; display: none">
		<input type="hidden" id="txtSelType" name="txtSelType" value='<% =Request("recSelType") %>'>
		<input type="hidden" id="txtSelTableid" name="txtSelTableid" value='<% =Request("recSelTableID") %>'>
		<input type="hidden" id="txtSelCurrentID" name="txtSelCurrentID" value='<% =Request("recSelCurrentID") %>'>
		<input type="hidden" id="txtSelTable" name="txtSelTable" value='<% =Request("recSelTable") %>'>
		<input type="hidden" id="txtSelDefOwner" name="txtSelDefOwner" value='<% =Request("recSelDefOwner") %>'>
		<input type="hidden" id="txtSelDefType" name="txtSelDefType" value="<% =Request("recSelDefType") %>">
	</form>

	<input type='hidden' id="txtTicker" name="txtTicker" value="0">
	<input type='hidden' id="txtLastKeyFind" name="txtLastKeyFind" value="">

	<script type="text/javascript">
		var frmPopup = document.getElementById("frmPopup");
		var frmFromOpener = document.getElementById("frmFromOpener");

		window.onload = util_recordselection_window_onload;

		function setForm() {
			if (frmFromOpener.txtSelDefOwner.value == "0" && frmPopup.txtSelectedAccess.value == "HD") {

				var sMessage = "Unable to select this " + frmFromOpener.txtSelType.value +
					" as it is a hidden " + frmFromOpener.txtSelType.value +
						" and you are not the owner of this definition";

				OpenHR.messageBox(sMessage, 48, frmFromOpener.txtSelDefType.value);

				self.close();
				return false;
			}

			//we are doing this for the base table
			if (frmFromOpener.txtSelTable.value == 'base') {
				//we are doing this for the picklist
				if (frmFromOpener.txtSelType.value == 'picklist') {
					window.dialogArguments.document.getElementById('txtBasePicklist').value = frmPopup.txtSelectedName.value;
					window.dialogArguments.document.getElementById('txtBasePicklistID').value = frmPopup.txtSelectedID.value;

					//if its hidden, set the relevant textbox value
					if (frmPopup.txtSelectedAccess.value == "HD") {
						window.dialogArguments.document.getElementById('baseHidden').value = 'Y';
					}
					else {
						window.dialogArguments.document.getElementById('baseHidden').value = '';
					}

					try {
						window.dialogArguments.document.getElementById('cmdBasePicklist').focus();
					}
					catch (e) {
					}
				}

				//we are doing this for the filter
				if (frmFromOpener.txtSelType.value == 'filter') {
					//we are doing this for the picklist
					window.dialogArguments.document.getElementById('txtBaseFilter').value = frmPopup.txtSelectedName.value;
					window.dialogArguments.document.getElementById('txtBaseFilterID').value = frmPopup.txtSelectedID.value;

					//if its hidden, set the relevant textbox value
					if (frmPopup.txtSelectedAccess.value == "HD") {
						window.dialogArguments.document.getElementById('baseHidden').value = 'Y';
					}
					else {
						window.dialogArguments.document.getElementById('baseHidden').value = '';
					}

					try {
						window.dialogArguments.document.getElementById('cmdBaseFilter').focus();
					}
					catch (e) {
					}
				}
			}

			//we are doing this for the parent 1 table
			if (frmFromOpener.txtSelTable.value == 'p1') {
				//we are doing this for the picklist
				if (frmFromOpener.txtSelType.value == 'picklist') {
					window.dialogArguments.document.getElementById('txtParent1Picklist').value = frmPopup.txtSelectedName.value;
					window.dialogArguments.document.getElementById('txtParent1PicklistID').value = frmPopup.txtSelectedID.value;

					//if its hidden, set the relevant textbox value
					if (frmPopup.txtSelectedAccess.value == "HD") {
						window.dialogArguments.document.getElementById('p1Hidden').value = 'Y';
					}
					else {
						window.dialogArguments.document.getElementById('p1Hidden').value = '';
					}

					try {
						window.dialogArguments.document.getElementById('cmdParent1Picklist').focus();
					}
					catch (e) {
					}
				}

				//we are doing this for the filter
				if (frmFromOpener.txtSelType.value == 'filter') {
					//we are doing this for the picklist
					window.dialogArguments.document.getElementById('txtParent1Filter').value = frmPopup.txtSelectedName.value;
					window.dialogArguments.document.getElementById('txtParent1FilterID').value = frmPopup.txtSelectedID.value;

					//if its hidden, set the relevant textbox value
					if (frmPopup.txtSelectedAccess.value == "HD") {
						window.dialogArguments.document.getElementById('p1Hidden').value = 'Y';
					}
					else {
						window.dialogArguments.document.getElementById('p1Hidden').value = '';
					}

					try {
						window.dialogArguments.document.getElementById('cmdParent1Filter').focus();
					}
					catch (e) {
					}
				}
			}

			//we are doing this for the parent 2 table
			if (frmFromOpener.txtSelTable.value == 'p2') {
				//we are doing this for the picklist
				if (frmFromOpener.txtSelType.value == 'picklist') {
					window.dialogArguments.document.getElementById('txtParent2Picklist').value = frmPopup.txtSelectedName.value;
					window.dialogArguments.document.getElementById('txtParent2PicklistID').value = frmPopup.txtSelectedID.value;

					//if its hidden, set the relevant textbox value
					if (frmPopup.txtSelectedAccess.value == "HD") {
						window.dialogArguments.document.getElementById('p2Hidden').value = 'Y';
					}
					else {
						window.dialogArguments.document.getElementById('p2Hidden').value = '';
					}

					try {
						window.dialogArguments.document.getElementById('cmdParent2Picklist').focus();
					}
					catch (e) {
					}
				}

				//we are doing this for the filter
				if (frmFromOpener.txtSelType.value == 'filter') {
					//we are doing this for the picklist
					window.dialogArguments.document.getElementById('txtParent2Filter').value = frmPopup.txtSelectedName.value;
					window.dialogArguments.document.getElementById('txtParent2FilterID').value = frmPopup.txtSelectedID.value;

					//if its hidden, set the relevant textbox value
					if (frmPopup.txtSelectedAccess.value == "HD") {
						window.dialogArguments.document.getElementById('p2Hidden').value = 'Y';
					}
					else {
						window.dialogArguments.document.getElementById('p2Hidden').value = '';
					}

					try {
						window.dialogArguments.document.getElementById('cmdParent2Filter').focus();
					}
					catch (e) {
					}
				}
			}

			//we are doing this for the child table
			if (frmFromOpener.txtSelTable.value == 'child') {
				//we are doing this for the filter
				if (frmFromOpener.txtSelType.value == 'filter') {
					//we are doing this for the filter
					window.dialogArguments.document.getElementById('txtChildFilter').value = frmPopup.txtSelectedName.value;
					window.dialogArguments.document.getElementById('txtChildFilterID').value = frmPopup.txtSelectedID.value;

					//if its hidden, set the relevant textbox value
					if (frmPopup.txtSelectedAccess.value == "HD") {
						window.dialogArguments.document.getElementById('childHidden').value = 'Y';
					}
					else {
						window.dialogArguments.document.getElementById('childHidden').value = '';
					}

					try {
						window.dialogArguments.document.getElementById('cmdChildFilter').focus();
					}
					catch (e) {
					}
				}
			}

			// Are we are doing this for a standard report
			if (frmFromOpener.txtSelTable.value == 'standardreport') {
				//we are doing this for the filter
				if (frmFromOpener.txtSelType.value == 'filter') {
					//we are doing this for the filter
					window.dialogArguments.document.getElementById('txtFilter').value = frmPopup.txtSelectedName.value;
					window.dialogArguments.document.getElementById('txtFilterID').value = frmPopup.txtSelectedID.value;

					try {
						window.dialogArguments.document.getElementById('cmdFilter').focus();
					}
					catch (e) {
					}
				}

				if (frmFromOpener.txtSelType.value == 'picklist') {
					window.dialogArguments.document.getElementById('txtPicklist').value = frmPopup.txtSelectedName.value;
					window.dialogArguments.document.getElementById('txtPicklistID').value = frmPopup.txtSelectedID.value;

					try {
						window.dialogArguments.document.getElementById('cmdPicklist').focus();
					}
					catch (e) {
					}
				}
			}

			// Are we are doing this for a calendar report
			if (frmFromOpener.txtSelTable.value == 'event') {
				//we are doing this for the filter
				if (frmFromOpener.txtSelType.value == 'filter') {
					//we are doing this for the filter
					window.dialogArguments.document.getElementById('txtEventFilter').value = frmPopup.txtSelectedName.value;
					window.dialogArguments.document.getElementById('txtEventFilterID').value = frmPopup.txtSelectedID.value;

					//if its hidden, set the relevant textbox value
					if (frmPopup.txtSelectedAccess.value == "HD") {
						window.dialogArguments.document.getElementById('baseHidden').value = 'Y';
					}
					else {
						window.dialogArguments.document.getElementById('baseHidden').value = '';
					}

					try {
						window.dialogArguments.document.getElementById('cmdEventFilter').focus();
					}
					catch (e) {
					}
				}
			}

			self.close();
			return false;
		}

		function util_recordselection_window_onload() {

			var iResizeBy,
					iNewWidth,
					iNewHeight,
					bdyMain = document.getElementById("bdyMain");

			//jQuery styling
			$(function () {
				$("input[type=submit], input[type=button], button")
					.button();
				$("input").addClass("ui-widget ui-widget-content ui-corner-all");
				$("input").removeClass("text");

				$("textarea").addClass("ui-widget ui-widget-content ui-corner-tl ui-corner-bl");
				$("textarea").removeClass("text");

				$("select").addClass("ui-widget ui-widget-content ui-corner-tl ui-corner-bl");
				$("select").removeClass("text");
				$("input[type=submit], input[type=button], button").removeClass("ui-corner-all");
				$("input[type=submit], input[type=button], button").addClass("ui-corner-tl ui-corner-br");

			});

			tableToGrid("#RecordSelectionHTMLTable", {
				colNames: ['ID', 'Name', 'UserName', 'Access'],
				colModel: [
					{ name: 'ID', hidden: true },
					{ name: 'Name', sortable: false },
					{ name: 'UserName', hidden: true },
					{ name: 'Access', hidden: true }
				],
				onSelectRow: function (rowID){
					//Get the values selected by the user...
					var ID = $("#RecordSelectionHTMLTable").jqGrid('getGridParam').data[$("#RecordSelectionHTMLTable").jqGrid('getGridParam', 'selrow') - 1].ID;
					var Name = $("#RecordSelectionHTMLTable").jqGrid('getGridParam').data[$("#RecordSelectionHTMLTable").jqGrid('getGridParam', 'selrow') - 1].Name;
					var UserName = $("#RecordSelectionHTMLTable").jqGrid('getGridParam').data[$("#RecordSelectionHTMLTable").jqGrid('getGridParam', 'selrow') - 1].UserName;
					var Access = $("#RecordSelectionHTMLTable").jqGrid('getGridParam').data[$("#RecordSelectionHTMLTable").jqGrid('getGridParam', 'selrow') - 1].Access;

					// ... and set the form values accordingly
					frmPopup.txtSelectedID.value = ID;
					frmPopup.txtSelectedUserName.value = UserName;
					frmPopup.txtSelectedAccess.value = Access;
					frmPopup.txtSelectedName.value = Name;
				},
				ondblClickRow: function (rowID)
				{
					//Get the values selected by the user...
					var ID = $("#RecordSelectionHTMLTable").jqGrid('getGridParam').data[$("#RecordSelectionHTMLTable").jqGrid('getGridParam', 'selrow') - 1].ID;
					var Name = $("#RecordSelectionHTMLTable").jqGrid('getGridParam').data[$("#RecordSelectionHTMLTable").jqGrid('getGridParam', 'selrow') - 1].Name;
					var UserName = $("#RecordSelectionHTMLTable").jqGrid('getGridParam').data[$("#RecordSelectionHTMLTable").jqGrid('getGridParam', 'selrow') - 1].UserName;
					var Access = $("#RecordSelectionHTMLTable").jqGrid('getGridParam').data[$("#RecordSelectionHTMLTable").jqGrid('getGridParam', 'selrow') - 1].Access;

					// ... and set the form values accordingly
					frmPopup.txtSelectedID.value = ID;
					frmPopup.txtSelectedUserName.value = UserName;
					frmPopup.txtSelectedAccess.value = Access;
					frmPopup.txtSelectedName.value = Name;
					
					//Set the form and close the dialog
					setForm();
				},
				rowNum: 1000,   //TODO set this to blocksize...
				height: 320,
				width: (screen.width) / 3 + 5,
				scrollerbar: true
			});
			
			//Select the first row
			$("#RecordSelectionHTMLTable").jqGrid('setSelection', 1);

			// Resize the popup.
			iResizeBy = bdyMain.scrollWidth - bdyMain.clientWidth;
			if (bdyMain.offsetWidth + iResizeBy > screen.width) {
				window.dialogWidth = new String(screen.width) + "px";
			}
			else {
				iNewWidth = new Number(window.dialogWidth.substr(0, window.dialogWidth.length - 2));
				iNewWidth = iNewWidth + iResizeBy;
				window.dialogWidth = new String(iNewWidth) + "px";
			}

			iResizeBy = bdyMain.scrollHeight - bdyMain.clientHeight;
			if (bdyMain.offsetHeight + iResizeBy > screen.height) {
				window.dialogHeight = new String(screen.height) + "px";
			}
			else {
				iNewHeight = new Number(window.dialogHeight.substr(0, window.dialogHeight.length - 2));
				iNewHeight = iNewHeight + iResizeBy;
				window.dialogHeight = new String(iNewHeight) + "px";
			}
			
		}
	</script>
	
	<style>
		#gbox_RecordSelectionHTMLTable {margin-left: 5px; }
	</style>
</body>
</html>
