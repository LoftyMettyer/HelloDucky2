<%@ Page Language="VB" Inherits="System.Web.Mvc.ViewPage" %>
<%@ Import Namespace="DMI.NET" %>


<!DOCTYPE html>
<html>
<head>
    <title>Event Log Selection - OpenHR Intranet</title>   
    <script src="<%: Url.Content("~/bundles/jQuery")%>" type="text/javascript"></script>
    <script src="<%: Url.Content("~/bundles/OpenHR_General")%>" type="text/javascript"></script>           
    <link href="<%: Url.Content("~/Content/OpenHR.css") %>" rel="stylesheet" type="text/css"/>
</head>
    
    <body>
<script type="text/javascript">

    function eventlogselection_window_onload() {

        self.focus();

        // Resize the grid to show all prompted values.
        iResizeBy = frmEventSelection.offsetParent.scrollWidth - frmEventSelection.offsetParent.clientWidth;
        if (frmEventSelection.offsetParent.offsetWidth + iResizeBy > screen.width) {
            window.dialogWidth = new String(screen.width) + "px";
        } else {
            iNewWidth = new Number(window.dialogWidth.substr(0, window.dialogWidth.length - 2));
            iNewWidth = iNewWidth + iResizeBy;
            window.dialogWidth = new String(iNewWidth) + "px";
        }

        iResizeBy = frmEventSelection.offsetParent.scrollHeight - frmEventSelection.offsetParent.clientHeight;
        if (frmEventSelection.offsetParent.offsetHeight + iResizeBy > screen.height) {
            window.dialogHeight = new String(screen.height) + "px";
        } else {
            iNewHeight = new Number(window.dialogHeight.substr(0, window.dialogHeight.length - 2));
            iNewHeight = iNewHeight + iResizeBy;
            window.dialogHeight = new String(iNewHeight) + "px";
        }
    }

</script>

<script type="text/javascript" id="scptGeneralFunctions">
<!--

    function cancelClick()
    {
        self.close();
    }

    function deleteClick()
    {
        var sEventIDs;
	
        var frmOpenerDelete =  window.dialogArguments.OpenHR.getForm("workframe","frmDelete");
        var frmOpenerLog =  window.dialogArguments.OpenHR.getForm("workframe","frmLog");
	
        sEventIDs = '';
	
        if (frmEventSelection.optSelection1.checked == true)
        {
            frmOpenerDelete.txtDeleteSel.value = 0;
		
            frmOpenerLog.ssOleDBGridEventLog.Redraw = false;
		
            for (var i=0; i<frmOpenerLog.ssOleDBGridEventLog.selbookmarks.count; i++)
            {
                sEventIDs = sEventIDs + frmOpenerLog.ssOleDBGridEventLog.Columns("ID").cellvalue(frmOpenerLog.ssOleDBGridEventLog.selbookmarks(i)) + ',';
            }
			
            sEventIDs = sEventIDs.substr(0, sEventIDs.length - 1);
		
            frmOpenerLog.ssOleDBGridEventLog.Redraw = true;
        }
		
		
        else if (frmEventSelection.optSelection2.checked == true)
        {
            frmOpenerDelete.txtDeleteSel.value = 1;
		
            frmOpenerLog.ssOleDBGridEventLog.Redraw = false;

            frmOpenerLog.ssOleDBGridEventLog.MoveFirst();
		
            for (var i=0; i<frmOpenerLog.ssOleDBGridEventLog.Rows; i++)
            {
                sEventIDs = sEventIDs + frmOpenerLog.ssOleDBGridEventLog.Columns("ID").cellvalue(frmOpenerLog.ssOleDBGridEventLog.AddItemBookmark(i)) + ',';
            }
			
            sEventIDs = sEventIDs.substr(0, sEventIDs.length - 1);
		
            frmOpenerLog.ssOleDBGridEventLog.Redraw = true;
        }
		
		
        else if (frmEventSelection.optSelection3.checked == true)
        {
            frmOpenerDelete.txtDeleteSel.value = 2;
        }
	
        frmOpenerDelete.txtSelectedIDs.value = sEventIDs;
	
        frmOpenerDelete.txtCurrentUsername.value = frmOpenerLog.cboUsername.options[frmOpenerLog.cboUsername.selectedIndex].value;
        frmOpenerDelete.txtCurrentType.value = frmOpenerLog.cboType.options[frmOpenerLog.cboType.selectedIndex].value;
        frmOpenerDelete.txtCurrentMode.value = frmOpenerLog.cboMode.options[frmOpenerLog.cboMode.selectedIndex].value;
        frmOpenerDelete.txtCurrentStatus.value = frmOpenerLog.cboStatus.options[frmOpenerLog.cboStatus.selectedIndex].value;
	
        frmOpenerDelete.txtViewAllPerm.value = frmOpenerLog.txtELViewAllPermission.value;

        window.dialogArguments.OpenHR.submitForm(frmOpenerDelete);
        self.close();
    }

    -->
</script>


<form id="frmEventSelection" name="frmEventSelection">
<table align="center" class="outline" cellPadding="5" cellSpacing="0" width="100%" height="100%">
	<tr>
		<td>
			<table WIDTH="100%" height="100%" class="invisible" cellspacing="0" cellpadding="0">
				<tr> 
					<td>
						<table HEIGHT="100%" WIDTH="100%" class="invisible" CELLSPACING="0" CELLPADDING="4">
							<tr height="30">
								<td>
									<img src="images/Question.gif" WIDTH="38" HEIGHT="39">
								</td>
								<td colspan="3">
									You have opted to delete entries from the Event Log.
									<br>
									Please make a selection from the options below : 
								</td>
							</tr> 
							<tr height="15">
								<td>
								</td>
								<td width="8">
								</td>
								<td>
									<input id="optSelection1" name="optSelection" type="radio" checked 		                                                        
									    onmouseover="try{radio_onMouseOver(this);}catch(e){}" 
                                        onmouseout="try{radio_onMouseOut(this);}catch(e){}"
                                        onfocus="try{radio_onFocus(this);}catch(e){}"
                                        onblur="try{radio_onBlur(this);}catch(e){}"/>
								</td>
								<td>
                                    <label 
                                        tabindex="-1"
                                        for="optSelection1"
                                        class="radio"
                                        onmouseover="try{radioLabel_onMouseOver(this);}catch(e){}" 
                                        onmouseout="try{radioLabel_onMouseOut(this);}catch(e){}"
                                    />
    									Only the currently highlighted row(s)
								</td>
							</tr> 
							<tr height="15">
								<td>
								</td>
								<td width="8">
								</td>
								<td>
									<input id="optSelection2" name="optSelection" type="radio" 
									    onmouseover="try{radio_onMouseOver(this);}catch(e){}" 
                                        onmouseout="try{radio_onMouseOut(this);}catch(e){}"
                                        onfocus="try{radio_onFocus(this);}catch(e){}"
                                        onblur="try{radio_onBlur(this);}catch(e){}"/>
								</td>
								<td>
                                    <label 
                                        tabindex="-1"
                                        for="optSelection2"
                                        class="radio"
                                        onmouseover="try{radioLabel_onMouseOver(this);}catch(e){}" 
                                        onmouseout="try{radioLabel_onMouseOut(this);}catch(e){}"
                                    />
    									All entries currently displayed
								</td>
							</tr> 
							<tr height="15">
								<td>
								</td>
								<td width="8">
								</td>
								<td>
									<input id="optSelection3" name="optSelection" type="radio"
									    onmouseover="try{radio_onMouseOver(this);}catch(e){}" 
                                        onmouseout="try{radio_onMouseOut(this);}catch(e){}"
                                        onfocus="try{radio_onFocus(this);}catch(e){}"
                                        onblur="try{radio_onBlur(this);}catch(e){}"/>
								</td>
								<td nowrap>
                                    <label 
                                        tabindex="-1"
                                        for="optSelection3"
                                        class="radio"
                                        onmouseover="try{radioLabel_onMouseOver(this);}catch(e){}" 
                                        onmouseout="try{radioLabel_onMouseOut(this);}catch(e){}"
                                    />
    									All entries (that the current user has permission to see)
								</td>
							</tr> 
							<tr height="5">
								<td colspan="4">
								</td>
							</tr>
							<tr>
								<td width="100%" colspan="4">
									<table HEIGHT="100%" WIDTH="100%" class="invisible" CELLSPACING="0" CELLPADDING="4">
										<tr>
											<td>
											</td>
											<td width="5">
												<input id="cmdDelete" type="button" value="Delete" name="cmdDelete" style="WIDTH: 80px" width="80" class="btn"
												    onclick="deleteClick();" 
                                                    onmouseover="try{button_onMouseOver(this);}catch(e){}" 
                                                    onmouseout="try{button_onMouseOut(this);}catch(e){}"
                                                    onfocus="try{button_onFocus(this);}catch(e){}"
                                                    onblur="try{button_onBlur(this);}catch(e){}" />
											</td>
											<td width="5">
												<input id="cmdCancel" type="button" value="Cancel" name="cmdCancel" style="WIDTH: 80px" width="80" class="btn"
												    onclick="cancelClick();" 
                                                    onmouseover="try{button_onMouseOver(this);}catch(e){}" 
                                                    onmouseout="try{button_onMouseOut(this);}catch(e){}"
                                                    onfocus="try{button_onFocus(this);}catch(e){}"
                                                    onblur="try{button_onBlur(this);}catch(e){}" />
											</td>
											<td>
											</td>
										</tr>
									</table>
								</td>
							</tr> 
						</table>
					</td>
					<td width="5"></td>
				</tr> 
			</table>
		</td>	
	</tr> 
</table>
</form>
    
    <script type="text/javascript">
        eventlogselection_window_onload();
</script>

</body>
</html>
