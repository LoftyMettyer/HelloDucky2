<%@ Page Title="" Language="VB" MasterPageFile="~/Views/Shared/Site.Master" Inherits="System.Web.Mvc.ViewPage" %>

<asp:Content ID="Content1" ContentPlaceHolderID="TitleContent" runat="server">
AboutHRPro
</asp:Content>

<asp:Content ID="Content2" ContentPlaceHolderID="MainContent" runat="server">

<%
	Const sReferringPage = ""
	Dim iNumRows = 0
%>


    <script type="text/javascript">
	    function AboutHRPro_window_onload() {		   
			<%if instr(sReferringPage, "LOGIN") <= 0 and instr(sReferringPage, ".ASP") > 0 then %>
				window.parent.document.all.item("workframeset").cols = "*, 0";	

				// Get menu.asp to refresh the menu.			 
				menu_refreshMenu();
			<%end if %>		    
	    }
    </script>

    <!--Client script to handle the screen events.-->
	<script type="text/javascript">
		<!--
		/* Return to the default page. */
		function cancelClick() {
			var sRPage = "<%=sReferringPage %>";
			//if 	("<%=sReferringPage %>"=="")
			if ((sRPage.indexOf("LOGIN") >= 0) || (sRPage.indexOf(".ASP") == -1)) {
				window.location.href = "login";
			}
			else {
				window.location.href = "default";
			}
		}
    -->
	</script>

   <form method="post" id="frmAboutForm" name="frmAboutForm">
        <br>
        <table align="center" class="outline" cellpadding="5" cellspacing="0">
            <tr>
                <td>
                    <table align="center" class="invisible" cellpadding="0" cellspacing="0">
                        <tr>
                            <td colspan="6" height="10">
                            </td>
                        </tr>
                        <tr>
                            <td width="40">
                            </td>
                            <td colspan="4">
                                <h3 align="center">
                                    About OpenHR</h3>
                            </td>
                            <td width="40">
                            </td>
                        </tr>
                        <%if sReferringPage = "" then 
                            iNumRows = 12
                          else
                            iNumRows = 16
                          end if
                        %>                        
                        <tr>
                            <td width="40" rowspan="<%=iNumRows %>">
                            </td>
<!--                            <td rowspan="<%=iNumRows %>" valign="top" align="LEFT" nowrap>
                                <img src="images/Logo.gif" width="100" height="87" alt="" style="border=1px solid #F9F7FB" />
                            </td>-->
                            <td width="20" rowspan="<%=iNumRows %>">
                            </td>
                            <td valign="top" align="LEFT" nowrap style="padding-right: 10;">
                                OpenHR Data Manager Intranet :
                            </td>
                            <td valign="top" align="left" nowrap>
                                Version
                                <%=session("Version")%>
                            </td>
                            <td width="40" rowspan="<%=iNumRows %>">
                            </td>
                        </tr>
                      <%if sReferringPage <> "" then %>
                        <tr>
                            <td valign="top" align="LEFT" nowrap style="padding-right: 10;">
                                Server :
                            </td>
                            <td valign="top" align="LEFT" nowrap>
                              <%=session("Server")%>
                            </td>
                        </tr>
                        <tr>
                            <td valign="top" align="LEFT" nowrap style="padding-right: 10;">
                                Database :
                            </td>
                            <td valign="top" align="LEFT" nowrap>
                              <%=session("Database")%>
                            </td>
                        </tr>
                        <tr>
                            <td valign="top" align="LEFT" nowrap style="padding-right: 10;">
                                Current user :
                            </td>
                            <td valign="top" align="LEFT" nowrap>
                                <%=session("Username")%>
                            </td>
                        </tr>
                        <tr>
                            <td valign="top" align="LEFT" nowrap style="padding-right: 10;">
                                User Group :
                            </td>
                            <td valign="top" align="LEFT" nowrap>
                                <%=session("Usergroup")%>
                            </td>
                        </tr>
                      <%end if %>
                        <tr>
                            <td colspan=2 valign="top" align="LEFT" nowrap>
                                <br />Copyright © Advanced Business Software and Solutions Ltd 2012
                            </td>
                        </tr>
                        <tr>
                            <td colspan=2 valign="top" align="LEFT" nowrap>
                                <a TARGET="Advanced Website" href="http://www.advancedcomputersoftware.com/abs" class="hypertext"
					                onmouseover="try{hypertextARef_onMouseOver(this);}catch(e){}" 
					                onmouseout="try{hypertextARef_onMouseOut(this);}catch(e){}"
		                            onfocus="try{hypertextARef_onFocus(this);}catch(e){}"
		                            onblur="try{hypertextARef_onBlur(this);}catch(e){}">
                                    http://www.advancedcomputersoftware.com/abs
                                </a>
                            </td>
                        </tr>
                        <tr>
                            <td colspan=2 valign="top" align="LEFT" nowrap>
                                &nbsp;
                            </td>
                        </tr>
                        <tr>
                            <td colspan=2 valign="top" align="LEFT" nowrap>
                                Contacts for Support :
                            </td>
                        </tr>
                        <tr>
                            <td valign="top" align="LEFT" nowrap style="padding-right: 10;">
                                Telephone :
                            </td>
                            <td valign="top" align="LEFT" nowrap>
                            <%if session("SupportTelNo") = "" then %>
                              08451 609 999
                            <%else
                              response.write(session("SupportTelNo"))
                            end if%>
                            </td>
                        </tr>
<!--                        <tr>
                            <td valign="top" align="LEFT" nowrap style="padding-right: 10;">
                                Fax :
                            </td>
                            <td valign="top" align="LEFT" nowrap>
                                <%=session("SupportFax") %>
                            </td>
                        </tr>
-->                        
                        <tr>
                            <td valign="top" align="LEFT" nowrap style="padding-right: 10;">
                                Email :
                            </td>
                            <td valign="top" align="LEFT" nowrap>
                                <%if session("SupportEmail") = "" then %>
                                  <a href="mailto://service.delivery@advancedcomputersoftware.com?subject=OpenHR Support Query - Data Manager Intranet" class="hypertext"
					                          onmouseover="try{hypertextARef_onMouseOver(this);}catch(e){}" 
					                          onmouseout="try{hypertextARef_onMouseOut(this);}catch(e){}"
                                      onfocus="try{hypertextARef_onFocus(this);}catch(e){}"
                                      onblur="try{hypertextARef_onBlur(this);}catch(e){}">
                                    service.delivery@advancedcomputersoftware.com</a>
                                    <%else%>                                
                                  <a href="mailto://<%=session("SupportEmail") %>?subject=OpenHR Support Query - Data Manager Intranet" class="hypertext"
					                          onmouseover="try{hypertextARef_onMouseOver(this);}catch(e){}" 
					                          onmouseout="try{hypertextARef_onMouseOut(this);}catch(e){}"
                                      onfocus="try{hypertextARef_onFocus(this);}catch(e){}"
                                      onblur="try{hypertextARef_onBlur(this);}catch(e){}">
                                    <%=session("SupportEmail") %></a>
                                <%end if %>
                            </td>
                        </tr>
                        <tr>
                            <td valign="top" align="LEFT" nowrap style="padding-right: 10;">
                                Web site :
                            </td>
                            <td valign="top" align="LEFT" nowrap>
                            <%if session("SupportWebpage") = "" then %>
                              <a TARGET="AdvancedSupportWebsite" href="http://webfirst.advancedcomputersoftware.com" class="hypertext"
					                      onmouseover="try{hypertextARef_onMouseOver(this);}catch(e){}" 
					                      onmouseout="try{hypertextARef_onMouseOut(this);}catch(e){}"
		                            onfocus="try{hypertextARef_onFocus(this);}catch(e){}"
		                            onblur="try{hypertextARef_onBlur(this);}catch(e){}">
                                http://webfirst.advancedcomputersoftware.com</a>
                             <%else %>                                
                              <a TARGET="AdvancedSupportWebsite" href="<%=session("SupportWebpage") %>" class="hypertext"
					                      onmouseover="try{hypertextARef_onMouseOver(this);}catch(e){}" 
					                      onmouseout="try{hypertextARef_onMouseOut(this);}catch(e){}"
		                            onfocus="try{hypertextARef_onFocus(this);}catch(e){}"
		                            onblur="try{hypertextARef_onBlur(this);}catch(e){}">
		                            <%=session("SupportWebpage") %></a>
		                        <%end if%>
                            </td>
                        </tr>
                        <tr>
                            <td colspan="2" valign="top" align="LEFT" nowrap>
                                &nbsp;
                            </td>
                        </tr>                       
                        <tr>
                            <td colspan="2" valign="top" align="LEFT" nowrap>
                                <font face='<%=session("Config-logintext-font")%>' color='<%=session("Config-logintext-colour")%>'
                                    style="font-size: <%=session("Config-logintext-size")%>pt; <%=session("Config-logintext-italics")%> <%=session("Config-logintext-bold")%>">
                                <a href="http://www.advancedconnect.co.uk/" target="_blank">Visit Advanced Connect for the latest OpenHR news and events</a>    
                                    
                                </font>
                            </td>
                        </tr>                                              
                        <tr>
                            <td colspan=6 valign="top" align="LEFT" nowrap>
                                &nbsp;
                            </td>
                        </tr>
                        <tr>
                            <td colspan="6" align="center">
                                <input id="btnCancel" name="btnCancel" type="button" class="btn" value="OK" style="width: 75px" width="75" 
                                    onclick="cancelClick()"
		                            onmouseover="try{button_onMouseOver(this);}catch(e){}" 
                                    onmouseout="try{button_onMouseOut(this);}catch(e){}"
                                    onfocus="try{button_onFocus(this);}catch(e){}"
                                    onblur="try{button_onBlur(this);}catch(e){}" />
                            </td>
                        </tr>
								<tr>
									 <td colspan="7" height="10">
									 </td>
								</tr>
				        </table>
						</td>
					</tr>
				</table>
    </form>
	 
	 <script type="text/javascript"> AboutHRPro_window_onload();</script>

    <form action="default_Submit" method="post" id="frmGoto" name="frmGoto">
        <%Html.RenderPartial("~/Views/Shared/gotoWork.ascx")%>
    </form>
</asp:Content>
