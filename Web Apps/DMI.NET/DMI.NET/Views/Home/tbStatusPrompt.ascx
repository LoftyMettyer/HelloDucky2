<%@ Control Language="VB" Inherits="System.Web.Mvc.ViewUserControl" %>

<SCRIPT type="text/javascript">
	function tbStatusPrompt_onload() {

		var frmForm = document.getElementById("frmForm");
		// Expand the option frame and hide the work frame.
		//TODO: window.parent.document.all.item("workframeset").cols = "0, *";


		// Set focus onto one of the form controls. 
		frmForm.optStatus_Booked.focus();

		// Get menu.asp to refresh the menu.
		menu_refreshMenu();
	}
</SCRIPT>

<script type="text/javascript">
	function Select() {
		
		var frmForm = document.getElementById("frmForm");
		var frmGotoOption = document.getElementById("frmGotoOption");

		if (frmForm.txtOptionAction.value == "SELECTADDFROMWAITINGLIST_1") {
			frmGotoOption.txtGotoOptionAction.value = "SELECTADDFROMWAITINGLIST_2";
			frmGotoOption.action = "tbAddFromWaitingListFind_Submit";
		}
		else {
			frmGotoOption.txtGotoOptionAction.value = "SELECTBOOKCOURSE_2";
			frmGotoOption.action = "tbBookCourseFind_Submit";
		}
	
		if (frmForm.optStatus_Provisional.checked) {
			frmGotoOption.txtGotoOptionLookupValue.value = "P";
		}
		else {
			frmGotoOption.txtGotoOptionLookupValue.value = "B";
		}

		//TODO: window.parent.frames("workframe").document.forms("frmFindForm").ssOleDBGridFindRecords.style.visibility = "visible";

		frmGotoOption.txtGotoOptionRecordID.value = frmForm.txtOptionRecordID.value;
		frmGotoOption.txtGotoOptionLinkRecordID.value = frmForm.txtOptionLinkRecordID.value;
		frmGotoOption.txtGotoOptionPage.value = "emptyoption";
		OpenHR.submitForm(frmGotoOption);
	}

	function Cancel()
	{  
		//TODO: window.parent.frames("workframe").document.forms("frmFindForm").ssOleDBGridFindRecords.style.visibility = "visible";
		var frmGotoOption = document.getElementById("frmGotoOption");
		
		frmGotoOption.txtGotoOptionAction.value = "CANCEL";
		frmGotoOption.txtGotoOptionRecordID.value = 0;
		frmGotoOption.txtGotoOptionLinkRecordID.value = 0;
		frmGotoOption.txtGotoOptionPage.value = "emptyoption";
		OpenHR.submitForm(frmGotoOption);
	}
	
</script>
<script src="<%: Url.Content("~/Scripts/ctl_SetStyles.js") %>" type="text/javascript"></script>

<div <%=session("BodyTag")%>>
<FORM action="" method="POST" id="frmForm" name="frmForm">

<table align=center class="outline" cellPadding=5 cellSpacing=0> 
    <tr>
	    <td>
		    <table align=center class="invisible" cellPadding=0 cellSpacing=0> 
			    <tr>
			        <td colSpan=4 height=10></td>
			    </tr>
			    <tr>
			        <td colSpan=4>
			            <H3 align=center>Book Course</H3>
			        </td>
			    </tr>

				<tr>
					<td width=20></td>
				    <TD align="center" colSpan=2>
				        Select the required booking status :
				    </TD>
					<td width=20></td>
				</tr>
	
				<TR>
				    <TD colSpan=4 height=20></TD>
				</TR>

				<TR>
					<td width=20></td>
				    <TD colspan=2 nowrap>
				        <INPUT type="radio" id=optStatus_Booked name=optStatus CHECKED
                            onmouseover="try{radio_onMouseOver(this);}catch(e){}" 
                            onmouseout="try{radio_onMouseOut(this);}catch(e){}"
                            onfocus="try{radio_onFocus(this);}catch(e){}"
                            onblur="try{radio_onBlur(this);}catch(e){}"/>
                        <label 
                            tabindex="-1"
                            for="optStatus_Booked"
                            class="radio"
                            onmouseover="try{radioLabel_onMouseOver(this);}catch(e){}" 
                            onmouseout="try{radioLabel_onMouseOut(this);}catch(e){}"
                        >
        				    Booked
	    			    </label>
    				</TD>
					<td width=20></td>
				</TR>

				<TR>
					<td width=20></td>
				    <TD colspan=2>
				        <INPUT type="radio" id=optStatus_Provisional name=optStatus
                            onmouseover="try{radio_onMouseOver(this);}catch(e){}" 
                            onmouseout="try{radio_onMouseOut(this);}catch(e){}"
                            onfocus="try{radio_onFocus(this);}catch(e){}"
                            onblur="try{radio_onBlur(this);}catch(e){}"/>
                       <label 
                            tabindex="-1"
                            for="optStatus_Provisional"
                            class="radio"
                            onmouseover="try{radioLabel_onMouseOver(this);}catch(e){}" 
                            onmouseout="try{radioLabel_onMouseOut(this);}catch(e){}"
                        >
    				        Provisional
				        </label>
                    </TD>
					<td width=20></td>
				</TR>

				<TR>
				    <TD colSpan=4 height=20></TD>
				</TR>

				<TR>
					<td width=20></td>
				    <TD colSpan=2>
				        <TABLE CLASS="invisible" CELLSPACING="0" CELLPADDING="0" align="center">
				            <TR>
					            <TD align="center">
									<input id="cmdSelect" name="cmdSelect" type="button" value="OK" class="btn" style="WIDTH: 75px" width="75" 
									    onclick="Select()"
                                        onmouseover="try{button_onMouseOver(this);}catch(e){}" 
                                        onmouseout="try{button_onMouseOut(this);}catch(e){}"
                                        onfocus="try{button_onFocus(this);}catch(e){}"
                                        onblur="try{button_onBlur(this);}catch(e){}" />
					            </TD>
								<TD width=20></TD>
				    		    <TD align="center">
				    			    <input id="cmdCancel" name="cmdCancel" type="button" value="Cancel" class="btn" style="WIDTH: 75px" width="75" 
				    			        onclick="Cancel()"
                                        onmouseover="try{button_onMouseOver(this);}catch(e){}" 
                                        onmouseout="try{button_onMouseOut(this);}catch(e){}"
                                        onfocus="try{button_onFocus(this);}catch(e){}"
                                        onblur="try{button_onBlur(this);}catch(e){}" />
				    		    </TD>
							</TR>
				        </TABLE>
				    </TD>
					<td width=20></td>
				</TR>

				<TR>
				    <TD colSpan=4 height=10></TD>
				</TR>
			</TABLE>
		</td>
	</TR>
</TABLE>

<INPUT type='hidden' id="txtOptionLinkRecordID" name="txtOptionLinkRecordID" value=<%=session("optionLinkRecordID")%>>
<INPUT type='hidden' id="txtOptionRecordID" name="txtOptionRecordID" value=<%=session("optionRecordID")%>>
<INPUT type='hidden' id="txtOptionAction" name="txtOptionAction" value=<%=session("optionAction")%>>
</FORM>

<FORM action="tbBookCourseFind_Submit" method=post id=frmGotoOption name=frmGotoOption style="visibility:hidden;display:none">
<%Html.RenderPartial("~/Views/Shared/gotoOption.ascx")%>
</FORM>

</div>
