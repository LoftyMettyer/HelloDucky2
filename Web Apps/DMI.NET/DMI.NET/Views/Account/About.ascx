<%@ Import Namespace="DMI.NET.Code" %>

	<div class="pageTitleDiv" style="margin-bottom: 15px">
		<span class="pageTitle" id="PopupReportDefinition_PageTitle">About OpenHR</span>
	</div>
	
	<div class="formField">
		<label>OpenHR :</label>
		<label>Version <%:session("Version")%></label>
		<%If Len(Session("Username")) > 0 Then%>
			<br />
			<label>Current user :</label>
			<label><%:session("Username")%></label>
			<br />
			<label>User Group :</label>
			<label><%:session("Usergroup")%></label>
		<%End If%>		
		<br/>
		<label>User Locale :</label>
		<span id="spnAbout_LocaleCultureName"></span>		

		<h4>Copyright © Advanced</h4>
		<a target="Advanced Website" href="http://www.oneadvanced.com" class="hypertext">http://www.oneadvanced.com</a>
		<h4>Contacts for Customer Services : </h4>
		<label>Telephone :</label>
		<%If Session("SupportTelNo") = "" Then%>
		<label>08451 609 999</label>
		<%Else%>
		<label><%:Session("SupportTelNo")%></label>
		<%End If%>
		<br/>
		<label>Email :</label>
		<%If Session("SupportEmail") = "" Then%>
		<a href="mailto://ohrsupport@oneadvanced.com?subject=OpenHR Support Query - Web Login" class="hypertext">ohrsupport@oneadvanced.com</a>
		<%Else%>
		<a href="mailto://<%:session("SupportEmail") %>?subject=OpenHR Support Query - Web Login" class="hypertext"><%:session("SupportEmail") %></a>
		<%End If%>
		<br/>
		<label>Web site :</label>
		<%If Session("SupportWebpage") = "" Then%>
		<a target="AdvancedSupportWebsite" href="https://customers.oneadvanced.com/" class="hypertext">https://customers.oneadvanced.com/</a>
		<%Else%>
		<a target="AdvancedSupportWebsite" href="<%:session("SupportWebpage") %>" class="hypertext"><%:session("SupportWebpage") %></a>
		<%End If%>
		<br/>
		<br/>
		<a target="AdvancedConnectWebsite" href="http://www.advancedconnect.co.uk/" class="hypertext">Visit Advanced Connect for the latest OpenHR news and events</a>		
	</div>

<script type="text/javascript">
	$("#spnAbout_LocaleCultureName")[0].innerHTML = window.UserLocale;
</script>


