<%@ Control Language="VB" Inherits="System.Web.Mvc.ViewUserControl(of DMI.NET.Models.ObjectRequests.WorkflowRunModel)" %>
<%@ Import Namespace="HR.Intranet.Server" %>
<%@ Import Namespace="System.Data.SqlClient" %>
<%@ Import Namespace="System.Data" %>

<%-- For other devs: Do not remove below line. --%>
<%="" %>
<%-- For other devs: Do not remove above line. --%>

<% 
	Response.Expires = 0

	Dim objDatabase As Database = CType(Session("DatabaseFunctions"), Database)
	Dim objDataAccess As clsDataAccess = CType(Session("DatabaseAccess"), clsDataAccess)

	Dim sMessage = ""
	Dim sFormElements = ""
	Dim sURL As String = objDatabase.GetModuleParameter("MODULE_WORKFLOW", "Param_URL")
	Dim sInstanceID = ""
	Dim iInstanceID As Integer
	
	If Len(sURL) > 0 Then
		
		Try
			Dim prmInstanceID = New SqlParameter("piInstanceID", SqlDbType.Int) With {.Direction = ParameterDirection.Output}
			Dim prmFormElements = New SqlParameter("psFormElements", SqlDbType.VarChar, -1) With {.Direction = ParameterDirection.Output}
			Dim prmMessage = New SqlParameter("psMessage", SqlDbType.VarChar, -1) With {.Direction = ParameterDirection.Output}

			objDataAccess.ExecuteSP("spASRInstantiateWorkflow", _
						New SqlParameter("piWorkflowID", SqlDbType.Int) With {.Value = Model.ID}, _
						prmInstanceID, prmFormElements, prmMessage)

			sInstanceID = prmInstanceID.Value.ToString()
			sFormElements = prmFormElements.Value.ToString()
			sMessage = prmMessage.Value.ToString()

		Catch ex As Exception
			Throw

		End Try
		
	End If
%>

<script type="text/JavaScript">

	var sMessage = $("#statusmessage").val();
	OpenHR.modalPrompt(sMessage, 0, "Workflow");

	var dataCollection = frmPopup.elements;
	if (dataCollection != null) {
		for (i = 0; i < dataCollection.length; i++) {
			sControlName = dataCollection.item(i).name;
			sControlName = sControlName.substr(0, 9);

			if (sControlName === "utilform_") {
				sForm = dataCollection.item(i).value;
				spawnWindow(sForm, '_blank', screen.availWidth, screen.availHeight, 'yes');
			}
		}
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
		winprops = 'height=' + h + ',width=' + w + ',top=' + wint + ',left=' + winl + ',scrollbars=' + scroll + ',resizable';

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

<div>
	<form name="frmPopup" id="frmPopup">
		<%
			Dim iFormCount = 0
			
			If Len(sMessage) = 0 Then
			
				Do While InStr(sFormElements, vbTab) > 0
					Dim sTemp = ""
					iFormCount = iFormCount + 1
					Dim iIndex = InStr(sFormElements, vbTab)

					Dim sStep = Left(sFormElements, iIndex - 1)
					
					Try

						Dim prmQueryString = New SqlParameter("psQueryString", SqlDbType.VarChar, -1) With {.Direction = ParameterDirection.Output}
						objDataAccess.ExecuteSP("spASRGetWorkflowQueryString", _
									New SqlParameter("piInstanceID", SqlDbType.Int) With {.Value = CInt(sInstanceID)}, _
									New SqlParameter("piElementID", SqlDbType.Int) With {.Value = CInt(sStep)}, _
									prmQueryString)

						sTemp = prmQueryString.Value.ToString()
						
					Catch ex As Exception
						Throw
						
					End Try
					

		%>
		<input type="hidden" id="utilform_<%=iFormCount%>" name="utilform_<%:iFormCount%>" value="<%=sTemp%>">
		<%
			sFormElements = Mid(sFormElements, iIndex + 1)
		Loop
	End If
		%>
		<input type="hidden" id="utilformcount" name="utilformcount" value="<%:iFormCount%>">
		<input type="hidden" id="utilinstance" name="utilinstance" value="<%:iInstanceID%>">
		<input type="hidden" id="utilid" name="utilid" value='<%:Model.ID%>'>
		<input type="hidden" id="utilname" name="utilname" value="<%:Model.Name%>">
		<input type="hidden" id="action" name="action" value="RUN">

		<%
			Dim sStatus = Model.Name
			If Len(sURL) = 0 Then
				sStatus &= " failed to initiate successfully.<br>No Workflow URL has been configured.<br>Contact your system administrator."
			Else
				If Len(sMessage) = 0 Then
					sStatus &= " initiated successfully."
				Else
					sStatus &= " failed to initiate successfully.<br><br>" & sMessage
				End If
			End If
		%>						
								
		<input type="hidden" id="statusmessage" name="statusmessage" value="<%:sStatus%>">


	</form>

</div>

