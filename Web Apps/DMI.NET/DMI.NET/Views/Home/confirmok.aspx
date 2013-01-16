<%@ Page Title="" Language="VB" MasterPageFile="~/Views/Shared/Site.Master" Inherits="System.Web.Mvc.ViewPage" %>

<asp:Content ID="Content1" ContentPlaceHolderID="TitleContent" runat="server">
</asp:Content>

<asp:Content ID="Content2" ContentPlaceHolderID="MainContent" runat="server">

<h2>confirmok</h2>

<script ID="clientEventHandlersJS" type="text/javascript">
<!--
    function window_onload()
    {
        // remove the popup if its there
        OpenHR.Closepopup();
    }

    function okClick() 
    {
        if (txtReloadMenu.value == 1)
        {
            window.parent.location.href = "main.asp";
            return;
        }
	
        sAction = txtReaction.value;

        if (sAction == "LOGOFF") {
            window.parent.location.href = window.parent.frames("menuframe").document.forms("frmMenuInfo").txtDefaultStartPage.value;
            return;	
        }

        if (sAction == "EXIT") {
            window.parent.close();
        }

        //if (txtUtilType.value > 0) {
        //	window.parent.frames("menuframe").loadDefSelPage(txtUtilType.value, txtUtilID.value, false);
        //}

        if (sAction == "CROSSTABS") {            
            menu_loadDefSelPage(txtUtilType.value, txtUtilID.value, 0, false);
        }

        if (sAction == "CUSTOMREPORTS") {
            menu_loadDefSelPage(2, txtUtilID.value, 0, false);
        }
	
        if (sAction == "CALENDARREPORTS") {
            menu_loadDefSelPage(17, txtUtilID.value, 0, false);
        }
	
        if (sAction == "MAILMERGE") {
            menu_loadDefSelPage(9, txtUtilID.value, 0, false);
        }

        if (sAction == "WORKFLOW") {
            menu_loadDefSelPage(25, txtUtilID.value, 0, false);
        }

        if (sAction == "WORKFLOWPENDINGSTEPS") {
            menu_autoLoadPage("workflowPendingSteps", false);
        }

        if (sAction == "WORKFLOWOUTOFOFFICE") {
            menu_WorkflowOutOfOffice();
        }

        if (sAction == "PICKLISTS") {
            menu_loadDefSelPage(10, txtUtilID.value, txtUtilTableID.value, false);
        }

        if (sAction == "FILTERS") {
            menu_loadDefSelPage(11, txtUtilID.value, txtUtilTableID.value, false);
        }

        if (sAction == "CALCULATIONS") {
            menu_loadDefSelPage(12, txtUtilID.value, txtUtilTableID.value, false);
        }

        if (sAction == "DEFAULT") {
            window.location.href = "default.asp";
        }
				
        if (sAction.substring(0, 7) == "mnutool") {
					
            menu_loadPage(sAction.substring(7, sAction.length));
            return;
        }
        else {

            var frmData = OpenHR.getForm("dataframe","frmData");
            var frmMenuInfo = OpenHR.getForm("menuframe","frmMenuInfo");

            if ((sAction.substring(0, 3) == "PT_") ||
                (sAction.substring(0, 3) == "PV_")) {
                // PT_ = primary table
                // PV_ = primary table view
                if (frmMenuInfo.txtPrimaryStartMode.value == 3) {
                    frmData.txtRecordDescription.value = ""
                    menu_loadFindPageFirst(sAction);
                }
                else {
                    menu_loadRecordEditPage(sAction);
                }
                return;
            }
		
            if (sAction.substring(0, 3) == "TS_") {
                // TS_ = Table screen
                if (frmMenuInfo.txtLookupStartMode.value == 3) {
                    frmData.txtRecordDescription.value = ""
                    menu_loadFindPageFirst(sAction);
                }
                else {
                    menu_loadRecordEditPage(sAction);
                }
                return;
            }

            if (sAction.substring(0, 3) == "QE_") {
                // QE_ = quick entry screen
                if (frmMenuInfo.txtQuickAccessStartMode.value == 3) {
                    frmData.txtRecordDescription.value = ""
                    menu_loadFindPageFirst(sAction);
                }
                else {
                    menu_loadRecordEditPage(sAction);
                }
                return;
            }

            if (sAction.substring(0, 3) == "HT_") {
                // HT_ = history table
                if (frmMenuInfo.txtHistoryStartMode.value == 3) {
                    frmData.txtRecordDescription.value = ""
                    menu_loadFindPageFirst(sAction);
                }
                else {
                    menu_loadRecordEditPage(sAction);
                }
                return;
            }
        }
    }
    -->
</script>

<table align=center class="outline" cellPadding=5 cellSpacing=0>
	<TR>
		<TD>
			<table class="invisible" cellspacing="0" cellpadding="0">
			    <tr> 
			        <td colspan=3 height=10></td>
			    </tr>

			    <tr> 
			        <td colspan=3 align=center> 
			            <H3><%= session("confirmtitle")%></H3>
			        </td>
			    </tr>
			  
			    <tr> 
			        <td width=20 height=10></td> 
			        <td> 
						<%=session("confirmtext")%>
			        </td>
			        <td width=20></td> 
			    </tr>
			  
			    <tr> 
			        <td colspan=3 height=20></td>
			    </tr>

			    <tr> 
			        <td colspan=3 height=10 align=center> 
		                <input id="cmdOK" name="cmdOK" type=button class="btn" value="OK" style="WIDTH: 75px" width="75" 
		                    onclick="okClick()"
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
<%
    Response.Write("<INPUT type='hidden' id=txtFollowPage name=txtFollowPage value=" & Session("followpage") & ">")
    Response.Write("<INPUT type='hidden' id=txtReaction name=txtReaction value=""" & Session("reaction") & """>")
    Response.Write("<INPUT type='hidden' id=txtUtilID name=txtUtilID value=" & Session("utilid") & ">")
    Response.Write("<INPUT type='hidden' id=txtUtilType name=txtUtilType value=" & Session("utiltype") & ">")
    Response.Write("<INPUT type='hidden' id=txtUtilTableID name=txtUtilTableID value=" & Session("utilTableID") & ">")
    Response.Write("<INPUT type='hidden' id=txtReloadMenu name=txtReloadMenu value=" & Session("reloadMenu") & ">")
    Session("confirmtitle") = Nothing
    Session("confirmtext") = Nothing
    Session("followpage") = Nothing
    Session("reaction") = Nothing
	session("reloadMenu") = 0
%>

<FORM action="default_Submit" method=post id=FORM1 name=frmGoto style="visibility:hidden;display:none">
    <%Html.RenderPartial("~/Views/Shared/gotoWork.ascx")%>
</FORM>

</asp:Content>
