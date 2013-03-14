<%@ Control Language="VB" Inherits="System.Web.Mvc.ViewUserControl" %>


<%@ Language=VBScript %>
<!--#INCLUDE FILE="include/svrCleanup.asp" -->
<% 
	Response.Expires = 0 

	sMessage = ""
	
	session("utiltype") = Request.Form("utiltype")
	session("utilid") = Request.Form("utilid")
	session("action") = "RUN"
	session("utilname") = Request.Form("utilname")
	
	Set cmdURL = Server.CreateObject("ADODB.Command")
	cmdURL.CommandText = "sp_ASRIntGetModuleParameter"
	cmdURL.CommandType = 4 ' Stored Procedure
	Set cmdURL.ActiveConnection = session("databaseConnection")

	Set prmModuleKey = cmdURL.CreateParameter("ModuleKey",200,1,8000)
	cmdURL.Parameters.Append prmModuleKey
	prmModuleKey.value = "MODULE_WORKFLOW"

	Set prmParameterKey = cmdURL.CreateParameter("ParameterKey",200,1,8000)
	cmdURL.Parameters.Append prmParameterKey
	prmParameterKey.value = "Param_URL"

	Set prmURL = cmdURL.CreateParameter("url",200,2, 8000) ' 200=adVarChar, 2=output, 8000=size
	cmdURL.Parameters.Append prmURL

	err = 0
	cmdURL.Execute
		
	if (err = 0) then
		sURL = cmdURL.Parameters("url").Value
	end if
	
	' Release the ADO command object.
	Set cmdURL = nothing
	
	if len(sURL) > 0 then
		Set cmdInitiate = Server.CreateObject("ADODB.Command")
		cmdInitiate.CommandText = "spASRInstantiateWorkflow"
		cmdInitiate.CommandType = 4 ' Stored Procedure
		Set cmdInitiate.ActiveConnection = session("databaseConnection")

		Set prmUtilID = cmdInitiate.CreateParameter("WorkflowID",3,1)
		cmdInitiate.Parameters.Append prmUtilID
		prmUtilID.value = cleanNumeric(clng(session("utilid")))

		Set prmInstanceID = cmdInitiate.CreateParameter("instanceID",3,2) ' 3=integer, 2=output
		cmdInitiate.Parameters.Append prmInstanceID

		Set prmFormElements = cmdInitiate.CreateParameter("formElements",200,2, 8000) ' 200=adVarChar, 2=output, 8000=size
		cmdInitiate.Parameters.Append prmFormElements

		Set prmMessage = cmdInitiate.CreateParameter("message",200,2, 8000) ' 200=adVarChar, 2=output, 8000=size
		cmdInitiate.Parameters.Append prmMessage

		err = 0
		cmdInitiate.Execute
			
		if (err = 0) then
			sInstanceID = cmdInitiate.Parameters("instanceID").Value
			sFormElements = cmdInitiate.Parameters("formElements").Value
			sMessage = cmdInitiate.Parameters("message").Value
		end if
	
		' Release the ADO command object.
		Set cmdInitiate = nothing
	end if	
%>
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<LINK href="OpenHR.css" rel=stylesheet type=text/css >
<TITLE>OpenHR Intranet</TITLE>
<meta http-equiv="X-UA-Compatible" content="IE=5">
<SCRIPT FOR=window EVENT=onload LANGUAGE=JavaScript>
	// Resize the popup.
	iResizeByHeight = frmPopup.offsetParent.scrollHeight - frmPopup.offsetParent.clientHeight;
	if (frmPopup.offsetParent.offsetHeight + iResizeByHeight > screen.availHeight) {
		try
		{
			window.parent.moveTo((screen.width - frmPopup.offsetParent.offsetWidth) / 2, 0);
			window.parent.resizeTo(frmPopup.offsetParent.offsetWidth, screen.availHeight);
		}
		catch(e) {}
	}
	else {
		try
		{
			window.parent.moveTo((screen.width - frmPopup.offsetParent.offsetWidth) / 2, (screen.availHeight - (frmPopup.offsetParent.offsetHeight + iResizeByHeight)) / 2);
			window.parent.resizeBy(0, iResizeByHeight);
		}
		catch(e) {}
	}

	iResizeByWidth = frmPopup.offsetParent.scrollWidth - frmPopup.offsetParent.clientWidth;
	if (frmPopup.offsetParent.offsetWidth + iResizeByWidth > screen.width) {
		try
		{
			window.parent.moveTo(0, (screen.availHeight - frmPopup.offsetParent.offsetHeight) / 2);
			window.parent.resizeTo(screen.width, frmPopup.offsetParent.offsetHeight);
		}
		catch(e) {}
	}
	else {
		try
		{
			window.parent.moveTo((screen.width - (frmPopup.offsetParent.offsetWidth + iResizeByWidth)) / 2, (screen.availHeight - frmPopup.offsetParent.offsetHeight) / 2);
			window.parent.resizeBy(iResizeByWidth, 0);
		}
		catch(e) {}
	}

	var dataCollection = frmPopup.elements;
	if (dataCollection!=null) 
	{
		for (i=0; i<dataCollection.length; i++)  
		{
			sControlName = dataCollection.item(i).name;
			sControlName = sControlName.substr(0, 9);
			if (sControlName=="utilform_") 
			{
				sForm = dataCollection.item(i).value;
				spawnWindow(sForm, '_blank', screen.availWidth, screen.availHeight,'yes');
			}
		}
	}	

<%
	if (len(sMessage) = 0) _
		and (len(sFormElements) > 0)  _
		and (len(sURL) > 0) then
%>		
	try 
	{
		self.close();
	}
	catch(e) {}		
	<%
	end if 
%>		
</script>

<script LANGUAGE="JavaScript">
<!--
	function pausecomp(millis) 
	{
		var date = new Date();
		var curDate = null;

		do 
		{ 
			curDate = new Date(); 
		} 
		while(curDate-date < millis);
	} 

	function spawnWindow(mypage, myname, w, h, scroll) 
	{
		var newWin;
		var winl = (screen.availWidth - w) / 2;
		var wint = (screen.availHeight - h) / 2;
		winprops = 'height='+h+',width='+w+',top='+wint+',left='+winl+',scrollbars='+scroll+',resizable';

		newWin = window.open(mypage, myname, winprops);

		try
		{
			if (parseInt(navigator.appVersion) >= 4) 
			{ 
				pausecomp(300);
				newWin.window.focus(); 
			}
		}
		catch(e) {}
	}
	-->
</script>
<!--#INCLUDE FILE="include/ctl_SetStyles.txt" -->
</HEAD>

<BODY <%=session("BodyColour")%>>
<FORM name=frmPopup id=frmPopup>
<%
	if len(sMessage) = 0 then
		iFormCount = 0
		Do While InStr(sFormElements, vbTab) > 0
			sTemp = ""
			iFormCount = iFormCount + 1
			iIndex = InStr(sFormElements, vbTab)

			sStep = Left(sFormElements, iIndex - 1)

			Set cmdQS = Server.CreateObject("ADODB.Command")
			cmdQS.CommandText = "spASRGetWorkflowQueryString"
			cmdQS.CommandType = 4 ' Stored Procedure
			Set cmdQS.ActiveConnection = session("databaseConnection")

			Set prmInstance = cmdQS.CreateParameter("instance",3,1)
			cmdQS.Parameters.Append prmInstance
			prmInstance.value = clng(sInstanceID)

			Set prmElement = cmdQS.CreateParameter("element",200,1,8000)
			cmdQS.Parameters.Append prmElement
			prmElement.value = clng(sStep)

			Set prmQS = cmdQS.CreateParameter("qs",200,2, 8000) ' 200=adVarChar, 2=output, 8000=size
			cmdQS.Parameters.Append prmQS

			err = 0
			cmdQS.Execute
			
			if (err = 0) then
				sTemp = cmdQS.Parameters("qs").Value
			end if
	
			' Release the ADO command object.
			Set cmdQS = nothing
%>	                  
	<input type="hidden" id="utilform_<%=iFormCount%>" name="utilform_<%=iFormCount%>" value="<%=sTemp%>">
<%
		  sFormElements = Mid(sFormElements, iIndex + 1)
		Loop
	end if
%>
	<input type="hidden" id="utilformcount" name="utilformcount" value="<%=iFormCount%>">
	<input type="hidden" id="utilinstance" name="utilinstance" value="<%=iInstanceID%>">
	<input type="hidden" id="utiltype" name="utiltype" value="<%=Session("utiltype")%>">
	<input type="hidden" id="utilid" name="utilid" value=<%=Session("utilid")%>>
	<input type="hidden" id="utilname" name="utilname" value="<%=replace(Session("utilname"), """", "&quot;")%>">
	<input type="hidden" id="action" name="action" value=<%=Session("action")%>>

	<table align=center class="outline" cellPadding=5 cellSpacing=0>
		<TR>
			<TD>
				<table class="invisible" cellspacing=0 cellpadding=0>
					<tr>
						<td colspan=3 height=10></td>
					</tr>
					<tr> 
						<td width=20 height=10></td> 
						<td align=center> 
							Workflow '<%=replace(session("utilname"), """", "&quot;")%>'
<%
	if len(sURL) = 0 then
%>
							failed to initiate successfully.<BR>No Workflow URL has been configured.<BR>Contact your system administrator.
<%
	else
		if len(sMessage) = 0 then
%>
							initiated successfully.
<%
		else
%>
							failed to initiate successfully.<BR><%=sMessage%>
<%
		end if
	end if
%>							</td>
						<td width=20></td> 
					</tr>
					<tr>
						<td colspan=3 height=20></td>
					</tr>
					<tr> 
						<td colspan=3 height=10 align=center> 
							<INPUT TYPE=button VALUE="OK" NAME=cmdClose class="btn" style="WIDTH: 80px" width=80 id=cmdClose
							    OnClick=window.parent.parent.self.close(); 
                                onmouseover="try{button_onMouseOver(this);}catch(e){}" 
                                onmouseout="try{button_onMouseOut(this);}catch(e){}"
                                onfocus="try{button_onFocus(this);}catch(e){}"
                                onblur="try{button_onBlur(this);}catch(e){}" />
						</td>
					</tr>
					<tr> 
						<td colspan=3 height=10></td>
					</tr>
				</table>
			</td>
		</tr>
	</table>
</FORM>

</BODY>
</HTML>
