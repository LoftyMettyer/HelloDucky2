<%@ Control Language="VB" Inherits="System.Web.Mvc.ViewUserControl" %>
<%@ Import Namespace="DMI.NET" %>

<%-- For other devs: Do not remove below line. --%>
<%="" %>
<%-- For other devs: Do not remove above line. --%>

<script type="text/javascript">
	function workpendingsteps_window_onload() {
		// Table to jQuery grid
		tableToGrid("#PendingStepsTable", {
			colNames: ['Name', 'URL'],
			colModel: [
				{ name: 'Name' },
				{ name: 'URL', hidden: true}
			],
			multiselect: true,
			onSelectRow: function (rowID) {
				refreshControls();
			},
			onSelectAll: function(rowID) {
				refreshControls();
			},
			ondblClickRow: function(rowID) {
			},
			rowNum: 1000   //TODO set this to blocksize...
		});
		
		//Select the first row
		//JIRA 3356 - don't set top row, and disable run button.
		//$("#PendingStepsTable").jqGrid('setSelection', 1);		
		refreshControls();

		//On clicking "Refresh",
		$("#mnutoolRefreshWFPendingStepsFind").click(function() {
			setrefresh();	
		});

		//On clicking "Run", open window with the selected item's URL
		$("#mnutoolRunWFPendingStepsFind").unbind().click(function () {
			var selectedRows = $("#PendingStepsTable [aria-selected='true']"); //Get the selected row			
			for (var i = 0; i < selectedRows.length; i++) {
				try {
					var url = selectedRows[i].children[2].innerHTML;					
					spawnWindow(url, '_blank', screen.availWidth, screen.availHeight, 'yes');
				}
				catch(e) {}
			}		
		});
		
		//On clicking "Close" generic closeclck in general.js
		$("#mnutoolCloseWFPendingStepsFind").hide(function () {
			// We're hiding this for now but I'm leaving it's 
			//click code below in case we need it in the future
		});
			
		$("#mnutoolCloseWFPendingStepsFind").click(function() {
			closeclick();
		});
		
		<%If _StepCount = 0 And Session("fromMenu") = 0 Then%>
		//no pending steps, and this is called from login, so show default page
		menu_loadPage('_default');
		<%Else%>
		<%If _StepCount = 0 Then%>
		//disable run button if no steps pending
		menu_toolbarEnableItem("mnutoolRunWFPendingStepsFind", false);
		<%End If%>

		$('#tblMessage').removeClass('hidden');
		showDefaultRibbon();
		$("#toolbarWFPendingStepsFind").parent().show();
		setTimeout('$("#toolbarWFPendingStepsFind").click()', 50);

		var newGridHeight = $("#findGridRow").height() - 50;
		$("#PendingStepsTable").jqGrid('setGridHeight', newGridHeight, true);
		<%End If%>
		

		$("#optionframe").hide();
		$("#workframe").show();

	}
</script>

<script runat="server">
		Private _PendingWorkflowStepsHTMLTable As New StringBuilder 'Used to construct the (temporary) HTML table that will be transformed into a jQuey grid table
		Private _StepCount As Integer = 0
		Private _WorkflowGood As Boolean = True
		
		Private Sub GetPendingWorkflowSteps
				'Get the pendings workflow steps from the database
				Dim _cmdDefSelRecords = CreateObject("ADODB.Command")
				
				_cmdDefSelRecords.CommandText = "spASRIntCheckPendingWorkflowSteps"
				_cmdDefSelRecords.CommandType = 4 ' Stored Procedure
				_cmdDefSelRecords.ActiveConnection = Session("databaseConnection")

				Err.Clear()
				Dim _rstDefSelRecords = _cmdDefSelRecords.Execute

				If Err.Number <> 0 Then
						' Workflow not licensed or configured. Go to default page.
						_WorkflowGood = False
				Else
			With _PendingWorkflowStepsHTMLTable
				.Append("<table id=""PendingStepsTable"">")
				.Append("<tr>")
				.Append("<th id=""Name"">Name</th>")
				.Append("<th id=""URL"">URL</th>")
				.Append("</tr>")
			End With
						'Loop over the records
						Do Until _rstDefSelRecords.eof
								_StepCount += 1
								With _PendingWorkflowStepsHTMLTable
										.Append("<tr>")
										.Append("<td>" & _rstDefSelRecords.Fields("description").Value & "</td>")
										.Append("<td>" & _rstDefSelRecords.Fields("url").Value & "</td>")
										.Append("</tr>")
								End With
								_rstDefSelRecords.movenext()
						Loop
						
						_PendingWorkflowStepsHTMLTable.Append("</table>")
						
						_rstDefSelRecords.close()
						_rstDefSelRecords = Nothing
				End If
				
				' Release the ADO command object.
			_cmdDefSelRecords = Nothing
	End Sub
		
		Private Sub Page_Load(sender As Object, e As System.EventArgs) Handles Me.Load
				GetPendingWorkflowSteps()
		End Sub

</script>

<form id="frmSteps" name="frmSteps" style="visibility: hidden; display: none">
	<%On Error Resume Next

		Response.Expires = -1
	
		If (Session("fromMenu") = 0) And (Session("reset") = 1) Then
			' Reset the Workflow OutOfOffice flag.
			Dim cmdOutOfOffice = CreateObject("ADODB.Command")
			cmdOutOfOffice.CommandText = "spASRWorkflowOutOfOfficeSet"
			cmdOutOfOffice.CommandType = 4 ' Stored Procedure
			cmdOutOfOffice.ActiveConnection = Session("databaseConnection")

			Dim prmValue = cmdOutOfOffice.CreateParameter("value", 11, 1)		' 11=bit, 1=input
			cmdOutOfOffice.Parameters.Append(prmValue)
			prmValue.value = 0

			Err.Clear()
			cmdOutOfOffice.Execute()

			cmdOutOfOffice = Nothing

			Session("reset") = 0
		End If
	%>
	<input type='hidden' id="txtFromMenu" name="txtFromMenu" value="<%=Session("fromMenu")%>">
</form>

<script type="text/javascript">
	function refreshControls() {
		var sSelectionList = jQuery("#PendingStepsTable").jqGrid('getGridParam', 'selarrrow');
		sSelectionList = (sSelectionList == null ? '' : sSelectionList);
		menu_toolbarEnableItem('mnutoolRunWFPendingStepsFind', (sSelectionList.length > 0));
	}

	function setrefresh() {
		OpenHR.submitForm("frmRefresh");
		<%If Session("fromMenu") = 0 Then%>
		menu_autoLoadPage("workflowPendingSteps", true);
		<%Else%>
		menu_autoLoadPage("workflowPendingSteps", false);
		<%End If%>
	}
	
	function pausecomp(millis) {
		var date = new Date();
		var curDate = null;

		do {
			curDate = new Date();
		}
		while (curDate - date < millis);
	}

	function spawnWindow(mypage, myname, w, h, scroll) {
		var newWin;
		var winl = (screen.availWidth - w) / 2;
		var wint = (screen.availHeight - h) / 2;
		var winprops = 'height=' + h + ',width=' + w + ',top=' + wint + ',left=' + winl + ',scrollbars=' + scroll + ',resizable';

		newWin = window.open(mypage, myname, winprops);

		try {
			if (parseInt(navigator.appVersion) >= 4) {
				pausecomp(300);
				newWin.window.focus();
			}
		}
		catch (e) { }
	}
</script>

<div <%=session("BodyTag")%>>

	<form name="frmDefSel" method="post" id="frmDefSel">
		<%If (_WorkflowGood = True) Or (Session("fromMenu") = 1) Then%>
		<% If _StepCount > 0 Then%>
		<div class="absolutefull">
			<div id="row1" style="margin-left: 20px; margin-right: 20px">
				<div class="pageTitleDiv">
					<a href='javascript:loadPartialView("linksMain", "Home", "workframe", null);' title='Home'>
						<i class='pageTitleIcon icon-circle-arrow-left'></i>
					</a>
					<span style="margin-left: 40px; margin-right: 20px" class="pageTitle">Pending Workflow Steps</span>
				</div>
			</div>
			<div id="findGridRow" style="height: 85%; margin-right: 20px; margin-left: 20px;">
				<%Response.Write(_PendingWorkflowStepsHTMLTable.ToString())%>
				<table class='outline' style='width: 100%;' cellspacing="0" cellpadding="0">
					<tr>
						<td width="20">&nbsp;&nbsp;&nbsp;&nbsp;</td>

						<td width="20">&nbsp;&nbsp;&nbsp;&nbsp;</td>

						<td style="visibility: hidden">
							<table style="height: 100%" class="invisible" cellspacing="0" cellpadding="0">
								<tr>
									<td>
										<input type="button"
											name="cmdRefresh"
											value="Refresh"
											id="cmdRefresh"
											class="btn"
											onclick="setrefresh();" />
									</td>
								</tr>
								<tr height="3px">
									<td></td>
								</tr>
								<tr>
									<td>
										<input type="button" name="cmdRun" value="Run" id="cmdRun" class="btn" />
									</td>
								</tr>
								<tr>
									<td></td>
								</tr>
								<tr>
									<td></td>
								</tr>
							</table>
						</td>
					</tr>
				</table>
			</div>
		</div>
		<%							
		Else
			Dim sMessage As String
			If _WorkflowGood = True Then
				' Display message saying no pending steps.
				sMessage = "No pending workflow steps"
			Else
				' Display error message.
				sMessage = "Error getting the pending workflow steps"
			End If
		%>
		<table align="center" class="outline hidden" cellpadding="5" cellspacing="0" id="tblMessage">
			<tr>
				<td width="20"></td>
				<td>
					<table class="invisible" cellspacing="0" cellpadding="0">
						<tr>
							<td height="10"></td>
						</tr>

						<tr>
							<td align="center">
								<h3>Pending Workflow Steps</h3>
							</td>
						</tr>

						<tr>
							<td align="center"><%=sMessage%></td>
						</tr>

						<tr>
							<td height="20"></td>
						</tr>

						<tr>
							<td height="10" align="center"></td>
						</tr>

						<tr>
							<td height="10"></td>
						</tr>
					</table>
				</td>
				<td width="20"></td>
			</tr>
		</table>
		<%			
	End If
End If
		%>
	</form>

	<form action="default_Submit" method="post" id="frmGoto" name="frmGoto" style="visibility: hidden; display: none">
		<%Html.RenderPartial("~/Views/Shared/gotoWork.ascx")%>
	</form>
</div>

<script type="text/javascript">
	workpendingsteps_window_onload();
</script>

