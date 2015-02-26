<%@ Control Language="VB" Inherits="System.Web.Mvc.ViewUserControl" %>


<script type="text/javascript">
		// Resize the popup.
		function progress_window_onload() {

				//debugger;

	 //     iResizeBy = frmPopup.offsetParent.scrollHeight - frmPopup.offsetParent.clientHeight;
 //       if (frmPopup.offsetParent.offsetHeight + iResizeBy > screen.height) {
//try {
 //             window.parent.moveTo((screen.width - frmPopup.offsetParent.offsetWidth) / 2, 0);
	//              window.parent.resizeTo(frmPopup.offsetParent.offsetWidth, screen.height);
		//        }
	 //         catch (e) { }
	//   }
	 //    else {
	 //         try {
		 //           window.parent.moveTo((screen.width - frmPopup.offsetParent.offsetWidth) / 2, (screen.height - (frmPopup.offsetParent.offsetHeight + iResizeBy)) / 2);
	//              window.parent.resizeBy(0, iResizeBy);
	 //         }
		//        catch (e) { }
	 //     }
		}
</script>

<FORM name=frmPopup id=frmPopup>

<table align=center class="outline" cellPadding=5 cellSpacing=0> 
	<tr>
		<td>
			<table align=center class="invisible" cellPadding=0 cellSpacing=0> 
				<tr>
					<td colSpan=3 height=20></td>
				</tr>
				<tr>
					<td width=20></td>
					<td align=center>Running
<%
	If Session("utiltype") = 1 Then
		Response.Write("Cross Tab.&nbsp; ")
		Response.Write("<INPUT value=""Cross Tab '" & Replace(Session("utilname"), """", "&quot;") & "'"" id=txtUtilTypeDesc name=txtUtilTypeDesc type=hidden>")

	ElseIf Session("utiltype") = 2 Then
		Response.Write("Custom Report.&nbsp; ")
		Response.Write("<INPUT value=""Custom Report '" & Replace(Session("utilname"), """", "&quot;") & "'"" id=txtUtilTypeDesc name=txtUtilTypeDesc type=hidden>")
	ElseIf Session("utiltype") = 9 Then
		Response.Write("Mail Merge.&nbsp; ")
		Response.Write("<INPUT value=""Mail Merge '" & Replace(Session("utilname"), """", "&quot;") & "'"" id=txtUtilTypeDesc name=txtUtilTypeDesc type=hidden>")
	ElseIf Session("utiltype") = 15 Then
		Response.Write("Absence Breakdown.&nbsp; ")
		Response.Write("<INPUT value=""Absence Breakdown"" id=txtUtilTypeDesc name=txtUtilTypeDesc type=hidden>")
	ElseIf Session("utiltype") = 16 Then
		Response.Write("Bradford Factor.&nbsp; ")
		Response.Write("<INPUT value=""Bradford Factor"" id=txtUtilTypeDesc name=txtUtilTypeDesc type=hidden>")
	ElseIf Session("utiltype") = 17 Then
		Response.Write("Calendar Report.&nbsp; ")
		Response.Write("<INPUT value=""Calendar Report '" & Replace(Session("utilname"), """", "&quot;") & "'"" id=txtUtilTypeDesc name=txtUtilTypeDesc type=hidden>")
	ElseIf Session("utiltype") = 35 Then
		Response.Write("9-Box Grid Report.&nbsp; ")
		Response.Write("<INPUT value=""9-Box Grid Report '" & Replace(Session("utilname"), """", "&quot;") & "'"" id=txtUtilTypeDesc name=txtUtilTypeDesc type=hidden>")

	Else
		Response.Write("<INPUT value=""Unknown"" id=txtUtilTypeDesc name=txtUtilTypeDesc type=hidden>")
	End If

	'						<INPUT onclick=window.parent.self.close(); id=Cancel style="WIDTH: 80px" type=button width=80 value=Cancel name=Cancel> 
%>
						Please wait...
								</td>
									<td width="20"></td>
							</tr>

								<tr>
										<td colspan="3" height="5"></td>
								</tr>

								<tr>
										<td colspan="3" align="center">
												<img src="images/ProgressBar.gif" width="220" height="19">
										</td>
								</tr>

								<tr>
										<td colspan="3" height="10"></td>
								</tr>

								<tr>
										<td width="20"></td>

										<td align="center">
												<input id="Cancel" style="WIDTH: 80px" type="button" width="80" value="Cancel" name="Cancel" class="btn"
														onclick="window.parent.raiseError('', true, true);"
														onmouseover="try{button_onMouseOver(this);}catch(e){}"
														onmouseout="try{button_onMouseOut(this);}catch(e){}"
														onfocus="try{button_onFocus(this);}catch(e){}"
														onblur="try{button_onBlur(this);}catch(e){}" />
										</td>
										<td width="20"></td>
								</tr>
								<tr>
										<td colspan="5" height="10"></td>
								</tr>
						</table>
			</td>
	</tr>
		</table>

</form>

<script type="text/javascript">
		//TODO This whole form to be replaced by cleaner progress control?
		progress_window_onload();
</script>
