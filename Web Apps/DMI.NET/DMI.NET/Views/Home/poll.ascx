<%@ Control Language="VB" Inherits="System.Web.Mvc.ViewUserControl" %>    
<%@Import namespace="DMI.NET" %>
<%@ Import Namespace="HR.Intranet.Server" %>
<%@ Import Namespace="System.Data" %>

<link href="<%: Url.LatestContent("~/Content/OpenHR.css")%>" rel="stylesheet" type="text/css">

		<script type="text/javascript">
				function poll_window_onload() {
					
						$("#pollframe").attr("data-framesource", "POLL");
						var frmMessages = document.getElementById("frmMessages");
						var sMessage = new String("");
					try {
						var controlCollection = frmMessages.elements;
						if (controlCollection != null) {
							for (var i = 0; i < controlCollection.length; i++) {
								if (sMessage.length > 0) {
									sMessage = sMessage + "\n\n";
								}
								sMessage = sMessage + controlCollection.item(i).value;
							}
							if (sMessage.length > 0) {
								var frmPollMsg = OpenHR.getForm("pollmessageframe", "frmSetMessage");
								frmPollMsg.txtMessage.value = sMessage;
								pollmessage_refreshMessage();
							}
						}
					}
					catch (e) {
						//alert("poll failed");
					}
				}
		</script>

		<form action="poll" method="post" id="frmHit" name="frmHit">
				<input type="hidden" id="txtDummy" name="txtDummy" value="0">
		</form>

		<form id="frmMessages" name="frmMessages">
				<%
					
					Dim objDataAccess As clsDataAccess = CType(Session("DatabaseAccess"), clsDataAccess)
					
					Try
						Dim rstMessages = objDataAccess.GetDataTable("sp_ASRIntPoll")

						Dim iloop = 1
						For Each objRow As DataRow In rstMessages.Rows
							%>		
								<input type='hidden' 
									id=txtMessage_<%=iLoop%> 
									name=txtMessage_<%=iLoop%> 
									value="<%=Replace(objRow(0).ToString(), """", "&quot;")%>">
						<%    
							iloop += 1
						Next
						
					Catch ex As Exception

					End Try
											
				%>
		</form>
		
		<script type="text/javascript">poll_window_onload();</script>
