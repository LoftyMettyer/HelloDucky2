<%@ Control Language="VB" Inherits="System.Web.Mvc.ViewUserControl" %>
<%@ Import Namespace="DMI.NET" %>

<%-- For other devs: Do not remove below line. --%>
<%="" %>
<%-- For other devs: Do not remove above line. --%>

<% 
	Response.Expires = 0

	Dim sMessage = ""
	Dim sFormElements = ""
	Dim sURL = ""
	Dim sInstanceID = ""
	Dim iInstanceID As Integer
	
	Session("utiltype") = Request.Form("utiltype")
	Session("utilid") = Request.Form("utilid")
	Session("action") = "RUN"
	Session("utilname") = Request.Form("utilname")
	
	Dim cmdURL = CreateObject("ADODB.Command")
	cmdURL.CommandText = "sp_ASRIntGetModuleParameter"
	cmdURL.CommandType = 4 ' Stored Procedure
	cmdURL.ActiveConnection = Session("databaseConnection")

	Dim prmModuleKey = cmdURL.CreateParameter("ModuleKey", 200, 1, 8000)
	cmdURL.Parameters.Append(prmModuleKey)
	prmModuleKey.value = "MODULE_WORKFLOW"

	Dim prmParameterKey = cmdURL.CreateParameter("ParameterKey", 200, 1, 8000)
	cmdURL.Parameters.Append(prmParameterKey)
	prmParameterKey.value = "Param_URL"

	Dim prmUrl = cmdURL.CreateParameter("url", 200, 2, 8000) ' 200=adVarChar, 2=output, 8000=size
	cmdURL.Parameters.Append(prmUrl)

	Err.Clear()
	cmdURL.Execute()
		
	If Err.Number = 0 Then
		sURL = CType(cmdURL.Parameters("url").Value, String)
	End If
	
	' Release the ADO command object.
	cmdURL = Nothing
	
	If Len(sURL) > 0 Then
		Dim cmdInitiate = CreateObject("ADODB.Command")
		cmdInitiate.CommandText = "spASRInstantiateWorkflow"
		cmdInitiate.CommandType = 4	' Stored Procedure
		cmdInitiate.ActiveConnection = Session("databaseConnection")

		Dim prmUtilID = cmdInitiate.CreateParameter("WorkflowID", 3, 1)
		cmdInitiate.Parameters.Append(prmUtilID)
		prmUtilID.value = CleanNumeric(CType(Session("utilid"), String))

		Dim prmInstanceID = cmdInitiate.CreateParameter("instanceID", 3, 2)	' 3=integer, 2=output
		cmdInitiate.Parameters.Append(prmInstanceID)

		Dim prmFormElements = cmdInitiate.CreateParameter("formElements", 200, 2, 8000)	' 200=adVarChar, 2=output, 8000=size
		cmdInitiate.Parameters.Append(prmFormElements)

		Dim prmMessage = cmdInitiate.CreateParameter("message", 200, 2, 8000)	' 200=adVarChar, 2=output, 8000=size
		cmdInitiate.Parameters.Append(prmMessage)

		Err.Clear()
		cmdInitiate.Execute()
			
		If (Err.Number = 0) Then
			sInstanceID = CType(cmdInitiate.Parameters("instanceID").Value, String)
			sFormElements = CType(cmdInitiate.Parameters("formElements").Value, String)
			sMessage = CType(cmdInitiate.Parameters("message").Value, String)
		End If
	
		' Release the ADO command object.
		cmdInitiate = Nothing
	End If
%>
<script type="text/JavaScript">

	// Resize the popup.
	iResizeByHeight = frmPopup.offsetParent.scrollHeight - frmPopup.offsetParent.clientHeight;
	if (frmPopup.offsetParent.offsetHeight + iResizeByHeight > screen.availHeight) {
		try {
			window.parent.moveTo((screen.width - frmPopup.offsetParent.offsetWidth) / 2, 0);
			window.parent.resizeTo(frmPopup.offsetParent.offsetWidth, screen.availHeight);
		}
		catch (e) { }
	}
	else {
		try {
			window.parent.moveTo((screen.width - frmPopup.offsetParent.offsetWidth) / 2, (screen.availHeight - (frmPopup.offsetParent.offsetHeight + iResizeByHeight)) / 2);
			window.parent.resizeBy(0, iResizeByHeight);
		}
		catch (e) { }
	}

	iResizeByWidth = frmPopup.offsetParent.scrollWidth - frmPopup.offsetParent.clientWidth;
	if (frmPopup.offsetParent.offsetWidth + iResizeByWidth > screen.width) {
		try {
			window.parent.moveTo(0, (screen.availHeight - frmPopup.offsetParent.offsetHeight) / 2);
			window.parent.resizeTo(screen.width, frmPopup.offsetParent.offsetHeight);
		}
		catch (e) { }
	}
	else {
		try {
			window.parent.moveTo((screen.width - (frmPopup.offsetParent.offsetWidth + iResizeByWidth)) / 2, (screen.availHeight - frmPopup.offsetParent.offsetHeight) / 2);
			window.parent.resizeBy(iResizeByWidth, 0);
		}
		catch (e) { }
	}

	var dataCollection = frmPopup.elements;
	if (dataCollection != null) {
		for (i = 0; i < dataCollection.length; i++) {
			sControlName = dataCollection.item(i).name;
			sControlName = sControlName.substr(0, 9);
			if (sControlName == "utilform_") {
				sForm = dataCollection.item(i).value;
				spawnWindow(sForm, '_blank', screen.availWidth, screen.availHeight, 'yes');
			}
		}
	}

		<%
	If (Len(sMessage) = 0) _
		And (Len(sFormElements) > 0) _
		And (Len(sURL) > 0) Then
%>
	try {
		self.close();
	}
	catch (e) { }
		<%
End If
%>

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

					Dim cmdQs = CreateObject("ADODB.Command")
					cmdQs.CommandText = "spASRGetWorkflowQueryString"
					cmdQs.CommandType = 4	' Stored Procedure
					cmdQs.ActiveConnection = Session("databaseConnection")

					Dim prmInstance = cmdQs.CreateParameter("instance", 3, 1)
					cmdQs.Parameters.Append(prmInstance)
					prmInstance.value = CLng(sInstanceID)

					Dim prmElement = cmdQs.CreateParameter("element", 200, 1, 8000)
					cmdQs.Parameters.Append(prmElement)
					prmElement.value = CLng(sStep)

					Dim prmQs = cmdQs.CreateParameter("qs", 200, 2, 8000)	' 200=adVarChar, 2=output, 8000=size
					cmdQs.Parameters.Append(prmQs)

					Err.Clear()
					cmdQs.Execute()
			
					If Err.Number = 0 Then
						sTemp = CType(cmdQs.Parameters("qs").Value, String)
					End If
	
					' Release the ADO command object.
					cmdQs = Nothing
		%>
		<input type="hidden" id="utilform_<%=iFormCount%>" name="utilform_<%=iFormCount%>" value="<%=sTemp%>">
		<%
			sFormElements = Mid(sFormElements, iIndex + 1)
		Loop
	End If
		%>
		<input type="hidden" id="utilformcount" name="utilformcount" value="<%=iFormCount%>">
		<input type="hidden" id="utilinstance" name="utilinstance" value="<%=iInstanceID%>">
		<input type="hidden" id="utiltype" name="utiltype" value="<%=Session("utiltype")%>">
		<input type="hidden" id="utilid" name="utilid" value='<%=Session("utilid")%>'>
		<input type="hidden" id="utilname" name="utilname" value="<%=replace(CType(Session("utilname"), String), """", "&quot;")%>">
		<input type="hidden" id="action" name="action" value='<%=Session("action")%>'>

		<table align="center" class="outline" cellpadding="5" cellspacing="0">
			<tr>
				<td>
					<table class="invisible" cellspacing="0" cellpadding="0">
						<tr>
							<td colspan="3" height="10"></td>
						</tr>
						<tr>
							<td width="20" height="10"></td>
							<td align="center">Workflow '<%=replace(CType(session("utilname"), String), """", "&quot;")%>'
								<%
									If Len(sURL) = 0 Then
								%>
							failed to initiate successfully.<br>
								No Workflow URL has been configured.<br>
								Contact your system administrator.
								<%
								Else
									If Len(sMessage) = 0 Then
								%>
							initiated successfully.
								<%
								Else
								%>
							failed to initiate successfully.<br>
								<%=sMessage%>
								<%
								End If
							End If
								%>							</td>
							<td width="20"></td>
						</tr>
						<tr>
							<td colspan="3" height="20"></td>
						</tr>
						<tr>
							<td colspan="3" height="10" align="center">
								<input type="button" value="OK" name="cmdClose" class="btn" style="WIDTH: 80px" width="80" id="cmdClose"
									onclick="window.parent.parent.self.close();"
									onmouseover="try{button_onMouseOver(this);}catch(e){}"
									onmouseout="try{button_onMouseOut(this);}catch(e){}"
									onfocus="try{button_onFocus(this);}catch(e){}"
									onblur="try{button_onBlur(this);}catch(e){}" />
							</td>
						</tr>
						<tr>
							<td colspan="3" height="10"></td>
						</tr>
					</table>
				</td>
			</tr>
		</table>
	</form>

</div>

