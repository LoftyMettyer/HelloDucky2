<%@ Control Language="VB" Inherits="System.Web.Mvc.ViewUserControl" %>
<%@ Import Namespace="DMI.NET.Classes" %>
<%@ Import Namespace="DMI.NET" %>
<%@ Import Namespace="HR.Intranet.Server" %>
<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="System.Data.SqlClient" %>

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
				{ name: 'URL', hidden: true }
			],
			multiselect: true,
			onSelectRow: function (rowID) {
				refreshControls();
			},
			onSelectAll: function (rowID) {
				refreshControls();
			},
			ondblClickRow: function (rowID) {
			},
			cmTemplate: { sortable: false },
			rowNum: 1000   //TODO set this to blocksize...			
		});
		//Select the first row
		refreshControls();
		
		if (window.currentLayout == "tiles") {
			//Hide the Home and Find buttons on this screen. JIRA-4028
			$('#toolbarHome').css('display', 'none');
			$('#toolbarWFPendingStepsFind').css('display', 'none');
		};

		//On clicking "Refresh",
		$("#mnutoolRefreshWFPendingStepsFind").click(function () {
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
				catch (e) { }
			}
		});

		//On clicking "Close" generic closeclck in general.js
		$("#mnutoolCloseWFPendingStepsFind").hide(function () {
			// We're hiding this for now but I'm leaving it's 
			//click code below in case we need it in the future
		});

		$("#mnutoolCloseWFPendingStepsFind").click(function () {
			closeclick();
		});

		<%If _StepCount = 0 Then%>
			//disable run button if no steps pending
			menu_toolbarEnableItem("mnutoolRunWFPendingStepsFind", false);
		<%End If%>
		
		 <%If _StepCount = 0 And Session("fromMenu") = 0 Then%>
			if (!menu_isSSIMode()) {
				//Maybe some code here that could fade out the workflow pending steps back to a blank home screen
			}
		<%End If%>
		
		$('#tblMessage').removeClass('hidden');
		showDefaultRibbon();
		
		// If current layout is tiles then hide searchBox for pending workflow steps screen.
		if (window.currentLayout == "tiles") {
			$('#searchBox').hide();
		}
		$("#toolbarWFPendingStepsFind").parent().show();
		setTimeout('$("#toolbarWFPendingStepsFind").click()', 50);

		var newGridHeight = $("#workframeset").height() - 150;
		$("#PendingStepsTable").jqGrid('setGridHeight', newGridHeight, true);

		var newGridWidth = $('#workframeset').width() - 75;
		$("#PendingStepsTable").jqGrid('setGridWidth', newGridWidth, true);
		
		$("#optionframe").hide();
		$("#workframe").show();
		$("#ToolsFrame").html('').hide();
	}
</script>

<script runat="server">
	Private ReadOnly _PendingWorkflowStepsHTMLTable As New StringBuilder 'Used to construct the (temporary) HTML table that will be transformed into a jQuery grid table
	Private _StepCount As Integer = 0
	Private _WorkflowEnabled As Boolean
		
	'Get the pendings workflow steps from the database
	Private Sub GetPendingWorkflowSteps()
		
		Dim objSession As SessionInfo = CType(Session("SessionContext"), SessionInfo)
		Dim objDataAccess As New clsDataAccess(objSession.LoginInfo)

		Try

			If Licence.IsModuleLicenced(SoftwareModule.Workflow) Then
				Dim _rstDefSelRecords = objDataAccess.GetDataTable("spASRIntCheckPendingWorkflowSteps", CommandType.StoredProcedure)

				With _PendingWorkflowStepsHTMLTable
					.Append("<table id=""PendingStepsTable"">")
					.Append("<tr>")
					.Append("<th id=""Name"">Name</th>")
					.Append("<th id=""URL"">URL</th>")
					.Append("</tr>")
				End With
				'Loop over the records
				For Each objRow As DataRow In _rstDefSelRecords.Rows
				
					_StepCount += 1
					With _PendingWorkflowStepsHTMLTable
						.Append("<tr>")
						.Append("<td>" & objRow("description").ToString() & "</td>")
						.Append("<td>" & objRow("url").ToString() & "</td>")
						.Append("</tr>")
					End With
				Next
						
				_PendingWorkflowStepsHTMLTable.Append("</table>")
				_WorkflowEnabled = True
				
			End If
			
		Catch ex As Exception
			_WorkflowEnabled = False
			
		End Try
		

	End Sub
		
		Private Sub Page_Load(sender As Object, e As System.EventArgs) Handles Me.Load
				GetPendingWorkflowSteps()
		End Sub

</script>

<script type="text/javascript">
	function refreshControls() {
		var sSelectionList = jQuery("#PendingStepsTable").jqGrid('getGridParam', 'selarrrow');
		sSelectionList = (sSelectionList == null ? '' : sSelectionList);
		menu_toolbarEnableItem('mnutoolRunWFPendingStepsFind', (sSelectionList.length > 0));
	}

	function setrefresh() {
		OpenHR.submitForm("frmRefresh");
		<%If Session("fromMenu") = 0 Then%>
		menu_loadWorkflowPendingSteps(true);
		<%Else%>
		menu_loadWorkflowPendingSteps(false);
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
		<%If (_WorkflowEnabled = True) Or (Session("fromMenu") = 1) Then%>
		<% If _StepCount > 0 Then%>
		<div class="absolutefull">
			<div id="row1" style="margin-left: 20px; margin-right: 20px">
				<div class="pageTitleDiv">
					<a href='javascript:loadPartialView("linksMain", "Home", "workframe", null);' title='Back'>
						<i class='pageTitleIcon icon-circle-arrow-left'></i>
					</a>
					<span class="pageTitle">Pending Workflow Steps</span>
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
			If _WorkflowEnabled = True Then
				sMessage = "No Pending Workflow Steps"
			Else
				sMessage = "Error getting the pending workflow steps"
			End If
		%>
		<div class="absolutefull">
			<div id="row1" style="margin-left: 20px; margin-right: 20px">
				<div class="pageTitleDiv">
					<a href='javascript:loadPartialView("linksMain", "Home", "workframe", null);' title='Back'>
						<i class='pageTitleIcon icon-circle-arrow-left'></i>
					</a>
					<span class="pageTitle"><%=sMessage %></span>
				</div>
			</div>
		</div>
		<%			
		End If
	End If
		%>
		<%=Html.AntiForgeryToken()%>
	</form>

</div>

<script type="text/javascript">
	workpendingsteps_window_onload();
</script>

