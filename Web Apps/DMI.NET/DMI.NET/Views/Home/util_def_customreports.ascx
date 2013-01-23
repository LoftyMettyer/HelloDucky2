<%@ Control Language="VB" Inherits="System.Web.Mvc.ViewUserControl" %>
<%@ Import Namespace="DMI.NET" %>

<script src="<%: Url.Content("~/Scripts/FormScripts/Util_Def_CustomReports.js") %>" type="text/javascript"></script>


<%Html.RenderPartial("Util_Def_CustomReports/dialog")%>

<div <%=session("BodyTag")%>>
<form id=frmDefinition name=frmDefinition>

<table align=center class="outline" cellPadding=5 cellSpacing=0 width=100% height=100%>
	<TR>
		<TD>
			<TABLE WIDTH="100%" height="100%" class="invisible" cellspacing=0 cellpadding=0>
				<tr height=5> 
					<td colspan=3></td>
				</tr> 

				<tr height=10>
					<TD width=10></td>
					<td>
						<INPUT type="button" value="Definition" id=btnTab1 name=btnTab1 disabled="disabled" class="btn btndisabled"
						    onclick="displayPage(1)"
                            onmouseover="try{button_onMouseOver(this);}catch(e){}" 
                            onmouseout="try{button_onMouseOut(this);}catch(e){}"
                            onfocus="try{button_onFocus(this);}catch(e){}"
                            onblur="try{button_onBlur(this);}catch(e){}" />
						<INPUT type="button" value="Related Tables" id=btnTab2 name=btnTab2  class="btn" 
						    onclick="displayPage(2)"
                            onmouseover="try{button_onMouseOver(this);}catch(e){}" 
                            onmouseout="try{button_onMouseOut(this);}catch(e){}"
                            onfocus="try{button_onFocus(this);}catch(e){}"
                            onblur="try{button_onBlur(this);}catch(e){}" />
						<INPUT type="button" value="Columns" id=btnTab3 name=btnTab3  class="btn" 
						    onclick="displayPage(3)"
                            onmouseover="try{button_onMouseOver(this);}catch(e){}" 
                            onmouseout="try{button_onMouseOut(this);}catch(e){}"
                            onfocus="try{button_onFocus(this);}catch(e){}"
                            onblur="try{button_onBlur(this);}catch(e){}" />
						<INPUT type="button" value="Sort Order" id=btnTab4 name=btnTab4  class="btn" 
						    onclick="displayPage(4)"
                            onmouseover="try{button_onMouseOver(this);}catch(e){}" 
                            onmouseout="try{button_onMouseOut(this);}catch(e){}"
                            onfocus="try{button_onFocus(this);}catch(e){}"
                            onblur="try{button_onBlur(this);}catch(e){}" />
						<INPUT type="button" value="Output" id=btnTab5 name=btnTab5  class="btn" 
						    onclick="displayPage(5)"
                            onmouseover="try{button_onMouseOver(this);}catch(e){}" 
                            onmouseout="try{button_onMouseOut(this);}catch(e){}"
                            onfocus="try{button_onFocus(this);}catch(e){}"
                            onblur="try{button_onBlur(this);}catch(e){}" />
					</td>
					<TD width=10></td>
				</tr> 
				
				<tr height=10> 
					<td colspan=3></td>
				</tr> 

				<tr> 
					<TD width=10></td>
					<td>
						<!-- First tab -->
						<DIV id=div1>
							<TABLE WIDTH="100%" height="100%" class="outline" cellspacing=0 cellpadding=5>
								<tr valign=top> 
									<td>
										<TABLE WIDTH="100%" class="invisible" CELLSPACING=0 CELLPADDING=0>
											<TR>
												<TD colspan=10 height=5></TD>
											</TR>

											<TR height=10>
												<TD width=5>&nbsp;</TD>
												<TD width=10>Name :</TD>
												<TD width=5>&nbsp;</TD>
												<TD colspan=2>
													<INPUT id=txtName name=txtName class="text" maxlength="50" style="WIDTH: 100%" onkeyup="changeTab1Control()">
												</TD>
												<TD width=20>&nbsp;</TD>
												<TD width=10>Owner :</TD>
												<TD width=5>&nbsp;</TD>
												<TD width="40%">
													<INPUT id=txtOwner name=txtOwner class="text textdisabled" style="WIDTH: 100%" disabled="disabled">
												</TD>
												<TD width=5>&nbsp;</TD>
											</TR>
											
											<TR>
												<TD colspan=10 height=5></TD>
											</TR>
											
											<TR height=60>
												<TD width=5>&nbsp;</TD>
												<TD width=10 nowrap valign=top>Description :</TD>
												<TD width=5>&nbsp;</TD>
												<TD width="40%" rowspan="3" colspan=2>
													<TEXTAREA id=txtDescription name=txtDescription class="textarea" style="HEIGHT: 99%; WIDTH: 100%" wrap=VIRTUAL height="0" maxlength="255" 
													    onkeyup="changeTab1Control()" 
													    onpaste="var selectedLength = document.selection.createRange().text.length;var pasteData = window.clipboardData.getData('Text');if ((this.value.length + pasteData.length - selectedLength) > parseInt(this.maxlength)) {return(false);}else {return(true);}" 
													    onkeypress="var selectedLength = document.selection.createRange().text.length;if ((this.value.length + 1 - selectedLength) > parseInt(this.maxlength)) {return(false);}else {return(true);}">
													</TEXTAREA>
												</TD>
												<TD width=20 nowrap>&nbsp;</TD>
												<TD width=10 valign=top>Access :</TD>
												<TD width=5>&nbsp;</TD>
												<TD width="40%" rowspan="3" valign=top>
													<%Html.RenderPartial("Util_Def_CustomReports/grdaccess")%>         
												</TD>
												<TD width=5>&nbsp;</TD>
											</TR>
											
											<TR height=10>
												<TD colspan=8>&nbsp;</TD>
											</TR>

											<TR height=10>
												<TD colspan=8>&nbsp;</TD>
											</TR>
										
											<TR height=40>
												<TD width=5>&nbsp;</TD>
												<TD colspan=8><hr></TD>
												<TD width=5>&nbsp;</TD>
											</TR>
											
											<TR height=10>
												<TD width=5>&nbsp;</TD>
												<TD width=85 nowrap vAlign=top>Base Table :</TD>
												<TD width=5>&nbsp;</TD>
												<TD vAlign=top colspan=2>
													<select id=cboBaseTable name=cboBaseTable style="WIDTH: 100%" class="combo combodisabled"
													    onchange="changeBaseTable()" disabled="disabled"> 
													</select>
												</TD>
												<TD width=20 nowrap>&nbsp;</TD>
												<TD width=10 vAlign=top>Records :</TD>
												<TD width=5>&nbsp;</TD>
												<TD width="40%"> 
													<TABLE class="invisible" CELLSPACING=0 CELLPADDING=0 width="100%">
														<TR>
															<TD width=5>
																<input CHECKED id=optRecordSelection1 name=optRecordSelection type=radio 
																    onclick="changeBaseTableRecordOptions()"
		                                                            onmouseover="try{radio_onMouseOver(this);}catch(e){}" 
		                                                            onmouseout="try{radio_onMouseOut(this);}catch(e){}"
                                                                    onfocus="try{radio_onFocus(this);}catch(e){}"
                                                                    onblur="try{radio_onBlur(this);}catch(e){}"/>
															</TD>
															<TD width=5>&nbsp;</TD>
															<TD>
                                                                <label tabindex="-1"
	                                                                for="optRecordSelection1"
	                                                                class="radio"
		                                                            onmouseover="try{radioLabel_onMouseOver(this);}catch(e){}" 
		                                                            onmouseout="try{radioLabel_onMouseOut(this);}catch(e){}">All</label>
    														</TD>
															<TD colspan=3>&nbsp;</TD>
														</TR>
													</Table>
												</TD>
											</TR>
											<TR height=10>
												<TD width=5>&nbsp;</TD>
												<TD width=85 nowrap vAlign=top></TD>
												<TD width=5>&nbsp;</TD>
												<TD vAlign=top colspan=2>
												</TD>
												<TD width=20 nowrap>&nbsp;</TD>
												<TD width=10 vAlign=top></TD>
												<TD width=5>&nbsp;</TD>
												<TD width="40%"> 
													<TABLE class="invisible" CELLSPACING=0 CELLPADDING=0 width="100%">
														<TR>
															<TD width=5>
																<input id=optRecordSelection2 name=optRecordSelection type=radio 
																    onclick="changeBaseTableRecordOptions()"
		                                                            onmouseover="try{radio_onMouseOver(this);}catch(e){}" 
		                                                            onmouseout="try{radio_onMouseOut(this);}catch(e){}"
                                                                    onfocus="try{radio_onFocus(this);}catch(e){}"
                                                                    onblur="try{radio_onBlur(this);}catch(e){}"/>
															</TD>
															<TD width=5>&nbsp;</TD>
															<TD width=50 nowrap>
                                                                <label 
                                                                    tabindex="-1"
	                                                                for="optRecordSelection2"
	                                                                class="radio"
		                                                            onmouseover="try{radioLabel_onMouseOver(this);}catch(e){}" 
		                                                            onmouseout="try{radioLabel_onMouseOut(this);}catch(e){}">Picklist</label>
															</TD>
															<TD width=5>&nbsp;</TD>
															<TD>
																<INPUT id=txtBasePicklist name=txtBasePicklist disabled="disabled" class="text textdisabled" style="WIDTH: 100%"> 
															</TD>
															<TD width=30 nowrap>
																<INPUT id=cmdBasePicklist name=cmdBasePicklist style="WIDTH: 100%" type=button disabled="disabled" class="btn btndisabled" value="..." 
																    onclick="selectRecordOption('base', 'picklist')"
                                                                    onmouseover="try{button_onMouseOver(this);}catch(e){}" 
                                                                    onmouseout="try{button_onMouseOut(this);}catch(e){}"
                                                                    onfocus="try{button_onFocus(this);}catch(e){}"
                                                                    onblur="try{button_onBlur(this);}catch(e){}" />
															</TD>
														</TR>
													</table>
												</td>
											</tr>	
											<TR height=10>
												<TD width=5>&nbsp;</TD>
												<TD width=85 nowrap vAlign=top></TD>
												<TD width=5>&nbsp;</TD>
												<TD vAlign=top colspan=2>
												</TD>
												<TD width=20 nowrap>&nbsp;</TD>
												<TD width=10 vAlign=top></TD>
												<TD width=5>&nbsp;</TD>
												<TD width="40%"> 
													<TABLE class="invisible" CELLSPACING=0 CELLPADDING=0 width="100%">
														<TR>
															<td width=5>
																<input id=optRecordSelection3 name=optRecordSelection type=radio
																    onclick=changeBaseTableRecordOptions() 
		                                                            onmouseover="try{radio_onMouseOver(this);}catch(e){}" 
		                                                            onmouseout="try{radio_onMouseOut(this);}catch(e){}"
                                                                    onfocus="try{radio_onFocus(this);}catch(e){}"
                                                                    onblur="try{radio_onBlur(this);}catch(e){}"/>
															</TD>
															<TD width=5>&nbsp;</TD>
															<TD width=50 nowrap>
                                                                <label 
                                                                    tabindex="-1"
	                                                                for="optRecordSelection3"
	                                                                class="radio"
		                                                            onmouseover="try{radioLabel_onMouseOver(this);}catch(e){}" 
		                                                            onmouseout="try{radioLabel_onMouseOut(this);}catch(e){}">Filter</label>
															</TD>
															<TD width=5>&nbsp;</TD>
															<TD>
																<INPUT id=txtBaseFilter name=txtBaseFilter class="text textdisabled" disabled="disabled" style="WIDTH: 100%"> 
															</TD>
															<TD width=30 nowrap>
																<INPUT id=cmdBaseFilter name=cmdBaseFilter style="WIDTH: 100%" type=button disabled="disabled" value="..." class="btn btndisabled" 
																    onclick="selectRecordOption('base', 'filter')" 
                                                                    onmouseover="try{button_onMouseOver(this);}catch(e){}" 
                                                                    onmouseout="try{button_onMouseOut(this);}catch(e){}"
                                                                    onfocus="try{button_onFocus(this);}catch(e){}"
                                                                    onblur="try{button_onBlur(this);}catch(e){}" />
															</TD>
														</TR>
													</TABLE>
												</TD>
											</TR>
											<TR height=10>
												<TD width=5>&nbsp;</TD>
												<TD width=85 nowrap vAlign=top></TD>
												<TD width=5>&nbsp;</TD>
												<TD vAlign=top colspan=2>
												</TD>
												<TD width=20 nowrap>&nbsp;</TD>
												<TD width=10 vAlign=top></TD>
												<TD width=5>&nbsp;</TD>
												<TD width="40%"> 
													<TABLE class="invisible" CELLSPACING=0 CELLPADDING=0 width="100%">
														<TR>
															<TD colspan=6 height=5></TD>
														</TR>
														<TR>
															<TD colspan=6 nowrap>
																<input name=chkPrintFilter id=chkPrintFilter type=checkbox disabled="disabled" tabindex="-1" 
	                                                                onmouseover="try{checkbox_onMouseOver(this);}catch(e){}" 
	                                                                onmouseout="try{checkbox_onMouseOut(this);}catch(e){}" 
																    onClick="changeTab1Control();"/>
                                                                <label 
                                                                    id="lblPrintFilter"
                                                                    name="lblPrintFilter"
				                                                    for="chkPrintFilter"
				                                                    class="checkbox checkboxdisabled"
				                                                    tabindex=0 
				                                                    onkeypress="try{checkboxLabel_onKeyPress(this);}catch(e){}"
					                                                onmouseover="try{checkboxLabel_onMouseOver(this);}catch(e){}" 
					                                                onmouseout="try{checkboxLabel_onMouseOut(this);}catch(e){}"
		                                                            onfocus="try{checkboxLabel_onFocus(this);}catch(e){}"
		                                                            onblur="try{checkboxLabel_onBlur(this);}catch(e){}">Display filter or picklist title in the report header</label>
															</TD>
														</TR>
														<TR>	
															<TD colspan=6 height=5></TD>
														</TR>
													</TABLE>
												</TD>
												<TD width=5>&nbsp;</TD>
											</TR>
											<TR>
												<TD colspan=10 height=30></TD>
											</TR>
										
										</TABLE>
									</td>
								</tr>
							</TABLE>
						</DIV>

						<!-- Second tab -->
						<DIV id=div2 style="visibility:hidden;display:none">
							<TABLE WIDTH="100%" height="100%" class="outline" cellspacing=0 cellpadding=5>
								<tr valign=top> 
									<td>
										<TABLE WIDTH="100%" class="invisible" CELLSPACING=0 CELLPADDING=0>
											<TR>
												<TD colspan=9 height=5></TD>
											</TR>
											<TR height=10>
												<TD width=5>&nbsp;</TD>
												<TD width=110 nowrap vAlign=top>Parent Table 1 :</TD>
												<TD width=5>&nbsp;</TD>
												<TD width="40%" vAlign=top>
													<INPUT id=txtParent1 name=txtParent1 class="text textdisabled" style="WIDTH: 100%" disabled="disabled">
												</TD>
												<TD width=20 nowrap>&nbsp;</TD>
												<TD width=10 vAlign=top>Records :</TD>
												<TD width=5>&nbsp;</TD>
												<TD width="40%">
													<TABLE class="invisible" CELLSPACING=0 CELLPADDING=0 width="100%">
														<TR>
															<TD width=5>
																<input CHECKED id=optParent1RecordSelection1 name=optParent1RecordSelection type=radio 
																    onclick="changeParent1TableRecordOptions()"
		                                                            onmouseover="try{radio_onMouseOver(this);}catch(e){}" 
		                                                            onmouseout="try{radio_onMouseOut(this);}catch(e){}"
                                                                    onfocus="try{radio_onFocus(this);}catch(e){}"
                                                                    onblur="try{radio_onBlur(this);}catch(e){}"/>
															</TD>
															<TD width=5>&nbsp;</TD>
															<TD width=30>
                                                                <label 
                                                                    tabindex="-1"
	                                                                for="optParent1RecordSelection1"
	                                                                class="radio"
		                                                            onmouseover="try{radioLabel_onMouseOver(this);}catch(e){}" 
		                                                            onmouseout="try{radioLabel_onMouseOut(this);}catch(e){}">All</label>
															</TD>
															<TD>&nbsp;</TD>
														</TR>
														<TR>
															<TD colspan=3 height=5></TD>
														</TR>						
														<TR>
															<TD width=5>
																<input id=optParent1RecordSelection2 name=optParent1RecordSelection type=radio 
																    onclick="changeParent1TableRecordOptions()"
		                                                            onmouseover="try{radio_onMouseOver(this);}catch(e){}" 
		                                                            onmouseout="try{radio_onMouseOut(this);}catch(e){}"
                                                                    onfocus="try{radio_onFocus(this);}catch(e){}"
                                                                    onblur="try{radio_onBlur(this);}catch(e){}"/>
															</TD>
															<TD width=5>&nbsp;</TD>
															<TD width=20>
                                                                <label 
                                                                    tabindex="-1"
	                                                                for="optParent1RecordSelection2"
	                                                                class="radio"
		                                                            onmouseover="try{radioLabel_onMouseOver(this);}catch(e){}" 
		                                                            onmouseout="try{radioLabel_onMouseOut(this);}catch(e){}">Picklist</label>
															</TD>
															<TD width=5>&nbsp;</TD>
															<TD>
																<INPUT id=txtParent1Picklist name=txtParent1Picklist disabled="disabled" style="WIDTH: 100%" class="text textdisabled"> 
															</TD>
															<TD width=30>
																<INPUT id=cmdParent1Picklist name=cmdParent1Picklist style="WIDTH: 100%" type=button class="btn btndisabled" value="..." disabled="disabled" 
																    onclick="selectRecordOption('p1', 'picklist')"
																    onmouseover="try{button_onMouseOver(this);}catch(e){}" 
                                                                    onmouseout="try{button_onMouseOut(this);}catch(e){}"
                                                                    onfocus="try{button_onFocus(this);}catch(e){}"
                                                                    onblur="try{button_onBlur(this);}catch(e){}" />
															</TD>
														</TR>
														<TR>
															<TD colspan=3 height=5></TD>
														</TR>						
														<TR>
															<TD width=5>
																<input id=optParent1RecordSelection3 name=optParent1RecordSelection type=radio 
																    onclick=changeParent1TableRecordOptions() 
		                                                            onmouseover="try{radio_onMouseOver(this);}catch(e){}" 
		                                                            onmouseout="try{radio_onMouseOut(this);}catch(e){}"
                                                                    onfocus="try{radio_onFocus(this);}catch(e){}"
                                                                    onblur="try{radio_onBlur(this);}catch(e){}"/>
															</TD>
															<TD width=5>&nbsp;</TD>
															<TD width=20>
                                                                <label 
                                                                    tabindex="-1"
	                                                                for="optParent1RecordSelection3"
	                                                                class="radio"
		                                                            onmouseover="try{radioLabel_onMouseOver(this);}catch(e){}" 
		                                                            onmouseout="try{radioLabel_onMouseOut(this);}catch(e){}">Filter</label>
															</TD>
															<TD width=5>&nbsp;</TD>
															<TD>
																<INPUT id=txtParent1Filter name=txtParent1Filter class="text textdisabled" disabled="disabled" style="WIDTH: 100%"> 
															</TD>
															<TD width=30>
																<INPUT id=cmdParent1Filter name=cmdParent1Filter style="WIDTH: 100%" type=button value="..." disabled="disabled" class="btn btndisabled"
																    onclick="selectRecordOption('p1', 'filter')" 
                                                                    onmouseover="try{button_onMouseOver(this);}catch(e){}" 
                                                                    onmouseout="try{button_onMouseOut(this);}catch(e){}"
                                                                    onfocus="try{button_onFocus(this);}catch(e){}"
                                                                    onblur="try{button_onBlur(this);}catch(e){}" />
															</TD>
														</TR>
													</TABLE>
												</TD>
												<TD width=5>&nbsp;</TD>
											</TR>

											<TR height=15>
												<TD colspan=9>
													<hr>
												</TD>
											</TR>

											<TR height=10>
												<TD width=5>&nbsp;</TD>
												<TD width=110 nowrap vAlign=top>Parent Table 2 :</TD>
												<TD width=5>&nbsp;</TD>
												<TD width="40%" vAlign=top>
													<INPUT id=txtParent2 name=txtParent2 class="text textdisabled" style="WIDTH: 100%" disabled="disabled">
												</TD>
												<TD width=20 nowrap>&nbsp;</TD>
												<TD width=10 vAlign=top>Records :</TD>
												<TD width=5>&nbsp;</TD>
												<TD width="40%">
													<TABLE class="invisible" CELLSPACING=0 CELLPADDING=0 width="100%">
														<TR>
															<TD width=5>
																<input CHECKED id=optParent2RecordSelection1 name=optParent2RecordSelection type=radio 
																    onclick="changeParent2TableRecordOptions()"
		                                                            onmouseover="try{radio_onMouseOver(this);}catch(e){}" 
		                                                            onmouseout="try{radio_onMouseOut(this);}catch(e){}"
                                                                    onfocus="try{radio_onFocus(this);}catch(e){}"
                                                                    onblur="try{radio_onBlur(this);}catch(e){}"/>
															</TD>
															<TD width=5>&nbsp;</TD>
															<TD width=30>
                                                                <label 
                                                                    tabindex="-1"
	                                                                for="optParent2RecordSelection1"
	                                                                class="radio"
		                                                            onmouseover="try{radioLabel_onMouseOver(this);}catch(e){}" 
		                                                            onmouseout="try{radioLabel_onMouseOut(this);}catch(e){}">All</label>
															</TD>
															<TD>&nbsp;</TD>
														</TR>
														<TR>
															<TD colspan=3 height=5></TD>
														</TR>						
														<TR>
															<TD width=5>
																<input id=optParent2RecordSelection2 name=optParent2RecordSelection type=radio 
																    onclick="changeParent2TableRecordOptions()"
		                                                            onmouseover="try{radio_onMouseOver(this);}catch(e){}" 
		                                                            onmouseout="try{radio_onMouseOut(this);}catch(e){}"
                                                                    onfocus="try{radio_onFocus(this);}catch(e){}"
                                                                    onblur="try{radio_onBlur(this);}catch(e){}"/>
															</TD>
															<TD width=5>&nbsp;</TD>
															<TD width=20>
                                                                <label 
                                                                    tabindex="-1"
	                                                                for="optParent2RecordSelection2"
	                                                                class="radio"
		                                                            onmouseover="try{radioLabel_onMouseOver(this);}catch(e){}" 
		                                                            onmouseout="try{radioLabel_onMouseOut(this);}catch(e){}">Picklist</label>
															</TD>
															<TD width=5>&nbsp;</TD>
															<TD>
																<INPUT id=txtParent2Picklist name=txtParent2Picklist disabled="disabled" style="WIDTH: 100%" class="text textdisabled"> 
															</TD>
															<TD width=30>
																<INPUT id=cmdParent2Picklist name=cmdParent2Picklist style="WIDTH: 100%" type=button class="btn btndisabled" value="..." disabled="disabled" 
																    onclick="selectRecordOption('p2', 'picklist')" 
																    onmouseover="try{button_onMouseOver(this);}catch(e){}" 
                                                                    onmouseout="try{button_onMouseOut(this);}catch(e){}"
                                                                    onfocus="try{button_onFocus(this);}catch(e){}"
                                                                    onblur="try{button_onBlur(this);}catch(e){}" />
															</TD>
														</TR>
														<TR>
															<TD colspan=3 height=5></TD>
														</TR>						
														<TR>
															<TD width=5>
																<input id=optParent2RecordSelection3 name=optParent2RecordSelection type=radio
																    onclick=changeParent2TableRecordOptions() 
		                                                            onmouseover="try{radio_onMouseOver(this);}catch(e){}" 
		                                                            onmouseout="try{radio_onMouseOut(this);}catch(e){}"
                                                                    onfocus="try{radio_onFocus(this);}catch(e){}"
                                                                    onblur="try{radio_onBlur(this);}catch(e){}"/>
															</TD>
															<TD width=5>&nbsp;</TD>
															<TD width=20>
                                                                <label 
                                                                    tabindex="-1"
	                                                                for="optParent2RecordSelection3"
	                                                                class="radio"
		                                                            onmouseover="try{radioLabel_onMouseOver(this);}catch(e){}" 
		                                                            onmouseout="try{radioLabel_onMouseOut(this);}catch(e){}">Filter</label>
															</TD>
															<TD width=5>&nbsp;</TD>
															<TD>
																<INPUT id=txtParent2Filter name=txtParent2Filter class="text textdisabled" disabled="disabled" style="WIDTH: 100%"> 
															</TD>
															<TD width=30>
																<INPUT id=cmdParent2Filter name=cmdParent2Filter style="WIDTH: 100%" type=button value="..." disabled="disabled" class="btn btndisabled"
																    onclick="selectRecordOption('p2', 'filter')" 
                                                                    onmouseover="try{button_onMouseOver(this);}catch(e){}" 
                                                                    onmouseout="try{button_onMouseOut(this);}catch(e){}"
                                                                    onfocus="try{button_onFocus(this);}catch(e){}"
                                                                    onblur="try{button_onBlur(this);}catch(e){}" />
															</TD>
														</TR>
													</TABLE>
												</TD>
												<TD width=5>&nbsp;</TD>
											</TR>

<!---------------------------------------------------------------------------------------------->
											
											<TR height=15>
												<TD colspan=9>
													<hr>
												</TD>
											</TR>
											
											<TR height=5>
												<TD width=5></TD>
												<TD width=90 nowrap colspan=7>Child Tables :</TD>
												<TD width=5></TD>
											</TR>
											<TR>
												<TD width=5>&nbsp;</TD>
												<TD colspan=7>
													<TABLE WIDTH="100%"  height="100%" class="invisible" CELLSPACING=0 CELLPADDING=0>
														<TR>
															<TD colspan=3 height=5></TD>
														</TR>
														<TR height=5>
															<TD rowspan=7>
																<%Html.RenderPartial("Util_Def_CustomReports/ssOleDBGridChildren")%>
															</TD>

															<TD width=10>&nbsp;</TD>
															<TD width=90>
																<input type="button" id=cmdAddChild name=cmdAddChild value="Add..." style="WIDTH: 100%" class="btn"
																    onclick="childAdd()"
								                                    onmouseover="try{button_onMouseOver(this);}catch(e){}" 
                                                                    onmouseout="try{button_onMouseOut(this);}catch(e){}"
                                                                    onfocus="try{button_onFocus(this);}catch(e){}"
                                                                    onblur="try{button_onBlur(this);}catch(e){}" />
			                                                </TD>
															<TD width=5>&nbsp;</TD>
														</TR>

														<TR height=5>
															<TD colspan=3></TD>
														</TR>
																	
														<TR height=5>
															
															<TD width=5>&nbsp;</TD>
															<TD width=90>
																<input type="button" id=cmdEditChild name=cmdChildEdit value="Edit..." style="WIDTH: 100%" class="btn" 
																    onclick="childEdit()"
								                                    onmouseover="try{button_onMouseOver(this);}catch(e){}" 
                                                                    onmouseout="try{button_onMouseOut(this);}catch(e){}"
                                                                    onfocus="try{button_onFocus(this);}catch(e){}"
                                                                    onblur="try{button_onBlur(this);}catch(e){}" />
															</TD>
															<TD width=5>&nbsp;</TD>
														</TR>
																	
														<TR height=11>
															<TD colspan=3></TD>
														</TR>

														<TR height=5>
															
															<TD width=5>&nbsp;</TD>
															<TD width=90>
																<input type="button" id=cmdRemoveChild name=cmdRemoveChild value="Remove" style="WIDTH: 100%" class="btn" 
																    onclick="childRemove()"
								                                    onmouseover="try{button_onMouseOver(this);}catch(e){}" 
                                                                    onmouseout="try{button_onMouseOut(this);}catch(e){}"
                                                                    onfocus="try{button_onFocus(this);}catch(e){}"
                                                                    onblur="try{button_onBlur(this);}catch(e){}" />
															</TD>
															<TD width=5>&nbsp;</TD>
														</TR>
																	
														<TR height=5>
															<TD colspan=3></TD>
														</TR>
																	
														<TR height=5>
															
															<TD width=5>&nbsp;</TD>
															<TD width=90>
																<input type="button" id=cmdRemoveAllChilds name=cmdRemoveAllChilds value="Remove All" style="WIDTH: 100%" class="btn" 
																    onclick="childRemoveAll()"
								                                    onmouseover="try{button_onMouseOver(this);}catch(e){}" 
                                                                    onmouseout="try{button_onMouseOut(this);}catch(e){}"
                                                                    onfocus="try{button_onFocus(this);}catch(e){}"
                                                                    onblur="try{button_onBlur(this);}catch(e){}" />
															</TD>
															<TD width=5>&nbsp;</TD>
														</TR>
													</TABLE>
												</TD>
												<TD width=5>&nbsp;</TD>
											</TR>
<!---------------------------------------------------------------------------------------------->

											<TR>
												<TD colspan=9 height=5></TD>
											</TR>
										</TABLE>
									</td>
								</tr>
							</TABLE>
						</DIV>

						<!-- Third tab -->
						<DIV id=div3 style="visibility:hidden;display:none">
							<TABLE WIDTH="100%"  height="100%" CLASS="outline" cellspacing=0 cellpadding=5>
								<tr valign=top> 
									<td>
										<TABLE WIDTH="100%" height="100%" class="invisible" CELLSPACING=0 CELLPADDING=0>
											<TR height=5>
												<TD colspan=7 height=5></TD>
											</TR>
											<TR height=5>
												<TD width=5 height=5></TD>
												<TD valign=top height=5>
													<TABLE WIDTH="100%"  HEIGHT=100% class="invisible" cellspacing=0 cellpadding=0>
														<TR height=5>
															<TD Height=5 colspan=7 width=100%>													
																<select id=cboTblAvailable name=cboTblAvailable style="WIDTH: 100%;" disabled="disabled" class="combo combodisabled"
																    onchange="refreshAvailableColumns();" > 
																</select>
															</td>
														</tr>
														<tr height=10>
															<td height=10 colspan=7 width=100%></td>
														</tr>
														<TR height=5>
															<TD height=5></TD>
															<TD height=5>
																<INPUT id=optColumns name=optAvailType type=radio CHECKED disabled="disabled" 
																    onclick="refreshAvailableColumns();"
		                                                            onmouseover="try{radio_onMouseOver(this);}catch(e){}" 
		                                                            onmouseout="try{radio_onMouseOut(this);}catch(e){}"
                                                                    onfocus="try{radio_onFocus(this);}catch(e){}"
                                                                    onblur="try{radio_onBlur(this);}catch(e){}"/>
															</TD>
															<TD height=5 width=5>
                                                                <label 
                                                                    tabindex="-1"
	                                                                for="optColumns"
	                                                                class="radio radiodisabled"
		                                                            onmouseover="try{radioLabel_onMouseOver(this);}catch(e){}" 
		                                                            onmouseout="try{radioLabel_onMouseOut(this);}catch(e){}">Columns</label>
															</TD>
															<TD width=5 height=5></TD>
															<TD height=5>
																<INPUT id=optCalc name=optAvailType type=radio disabled="disabled" 
																    onclick="refreshAvailableColumns();"
		                                                            onmouseover="try{radio_onMouseOver(this);}catch(e){}" 
		                                                            onmouseout="try{radio_onMouseOut(this);}catch(e){}"
                                                                    onfocus="try{radio_onFocus(this);}catch(e){}"
                                                                    onblur="try{radio_onBlur(this);}catch(e){}"/>
															</TD>
															<TD width=5 height=5>
																<label 
																	tabindex="-1"
																	for="optCalc"
																	class="radio radiodisabled"
																	onmouseover="try{radioLabel_onMouseOver(this);}catch(e){}" 
																	onmouseout="try{radioLabel_onMouseOut(this);}catch(e){}">Calculations</label>    
															</TD>
															<TD height=5></TD>
														<tr height=10>
															<td height=10 colspan=7 width=100%></td>
														</tr>
														</TR>
													</TABLE>
												</TD>
												<TD width=10></TD>
												<TD width=5 nowrap></TD>
												<TD width=10></TD>
												<TD rowspan="3" width=40% height=100%>
													<TABLE WIDTH="100%" Height=100% class="invisible" cellspacing=0 cellpadding=0>
														<TR>
															<TD WIDTH=100% HEIGHT=100%>													
																<%Html.RenderPartial("Util_Def_CustomReports/ssOleDBGridSelectedColumns")%>
															</TD>
														</TR>
													</TABLE>
												</TD>
												<TD width=5></TD>
											</TR>
											
											<TR Height=5>
												<TD height=5 colspan=7></TD>
											</TR>
											
											<TR>
												<TD width=5></TD>
												<TD rowspan="5" width=40% height=100%>
													<%Html.RenderPartial("Util_Def_CustomReports/ssOleDBGridAvailableColumns")%>
												</TD>
												<TD width=10 nowrap></TD>
												<TD height=5 valign=top align=center>
													<TABLE WIDTH=100% class="invisible" CELLSPACING=0 CELLPADDING=0>
														<TR height=25>
															<td>&nbsp</TD>
															<TD width=90 nowrap align=center>
																<input type="button" name=cmdColumnAdd id=cmdColumnAdd value="Add..." style="WIDTH: 100%; HEIGHT: 100%" class="btn"
																    onClick="columnSwap(true)"
                                                                    onmouseover="try{button_onMouseOver(this);}catch(e){}" 
                                                                    onmouseout="try{button_onMouseOut(this);}catch(e){}"
                                                                    onfocus="try{button_onFocus(this);}catch(e){}"
                                                                    onblur="try{button_onBlur(this);}catch(e){}" />
															</TD>
															<td>&nbsp</TD>
														</TR>
														<tr height=5><td></td></tr>
														<TR height=25>
															<td></TD>
															<TD width=90 nowrap align=center>
																<input type="button" name=cmdColumnAddAll id=cmdColumnAddAll value="Add All" style="WIDTH: 100%; HEIGHT: 100%" class="btn"
																    onClick="columnSwapAll(true)"
                                                                    onmouseover="try{button_onMouseOver(this);}catch(e){}" 
                                                                    onmouseout="try{button_onMouseOut(this);}catch(e){}"
                                                                    onfocus="try{button_onFocus(this);}catch(e){}"
                                                                    onblur="try{button_onBlur(this);}catch(e){}" />
															</TD>
															<td></TD>
														</TR>
														<tr height=15><td></td></tr>
														<TR height=25>
															<td></TD>
															<TD width=90 nowrap align=center>
																<input type="button" name=cmdColumnRemove id=cmdColumnRemove value="Remove" style="WIDTH: 100%; HEIGHT: 100%" class="btn"
																    onClick="columnSwap(false)"
                                                                    onmouseover="try{button_onMouseOver(this);}catch(e){}" 
                                                                    onmouseout="try{button_onMouseOut(this);}catch(e){}"
                                                                    onfocus="try{button_onFocus(this);}catch(e){}"
                                                                    onblur="try{button_onBlur(this);}catch(e){}" />
															</TD>
															<td></TD>
														</TR>
														<tr height=5><td></td></tr>
														<TR height=25>
															<td></TD>
															<TD width=90 nowrap align=center>
																<input type="button" name=cmdColumnRemoveAll id=cmdColumnRemoveAll value="Remove All" style="WIDTH: 100%; HEIGHT: 100%" class="btn" 
																    onClick="columnSwapAll(false)"
                                                                    onmouseover="try{button_onMouseOver(this);}catch(e){}" 
                                                                    onmouseout="try{button_onMouseOut(this);}catch(e){}"
                                                                    onfocus="try{button_onFocus(this);}catch(e){}"
                                                                    onblur="try{button_onBlur(this);}catch(e){}" />
															</TD>
															<td></TD>
														</TR>
														<tr height=15><td></td></tr>
														<TR height=25>
															<td></TD>
															<TD width=90 nowrap align=center>
																<input type="button" name=cmdColumnMoveUp id=cmdColumnMoveUp value="Up" style="WIDTH: 100%; HEIGHT: 100%" class="btn" 
																    onClick="columnMove(true)"
                                                                    onmouseover="try{button_onMouseOver(this);}catch(e){}" 
                                                                    onmouseout="try{button_onMouseOut(this);}catch(e){}"
                                                                    onfocus="try{button_onFocus(this);}catch(e){}"
                                                                    onblur="try{button_onBlur(this);}catch(e){}" />
															</TD>
															<td></TD>
														</TR>
														<tr height=5><td></td></tr>
														<TR height=25>
															<td></TD>
															<TD width=90 nowrap align=center>
																<input type="button" name=cmdColumnMoveDown id=cmdColumnMoveDown value="Down" style="WIDTH: 100%; HEIGHT: 100%" class="btn" 
																    onClick="columnMove(false)"
                                                                    onmouseover="try{button_onMouseOver(this);}catch(e){}" 
                                                                    onmouseout="try{button_onMouseOut(this);}catch(e){}"
                                                                    onfocus="try{button_onFocus(this);}catch(e){}"
                                                                    onblur="try{button_onBlur(this);}catch(e){}" />
															</TD>
															<td></TD>
														</TR>
													</TABLE>
												</TD>
												<TD width=10 nowrap></TD>
												<TD width=5></TD>
											</TR>

											<TR Height=5>
												<TD colspan=7 height=5></TD>
											</TR>

											<TR height=5>
												<TD width=5></TD>
												<TD width=10></TD>
												<TD width=80></TD>
												<TD width=10></TD>
												<TD valign=top>
													<TABLE WIDTH="100%" class="invisible" CELLSPACING=0 CELLPADDING=0>
														<TR>
															<TD width=125>Heading :</TD>
															<TD width=5></TD>
															<TD>
																<INPUT id=txtColHeading name=txtColHeading maxlength="50" class="text" style="WIDTH: 100%" 
																    onchange="validateColHeading()" 
																    onkeyup="validateColHeading();" 
																    onblur="trimColHeading();">
															</TD>
														</TR>
														<TR>
															<TD width=125>Size :</TD>
															<TD width=5></TD>
															<TD>
																<INPUT id=txtSize name=txtSize maxlength="50" class="text" style="WIDTH: 100%" 
																    onchange="validateColSize();" 
																    onkeyup="validateColSize();">
															</TD>
														</TR>
														<TR>
															<TD width=125>Decimals :</TD>
															<TD width=5></TD>
															<TD>
																<INPUT id=txtDecPlaces name=txtDecPlaces maxlength="50" class="text" style="WIDTH: 100%" 
																    onchange="validateColDecimals();" 
																    onkeyup="validateColDecimals();">						 									
															</TD>
														</TR>
													</TABLE>
												</TD>
												<TD width=5></TD>
											</TR>
											
											<TR Height=5>
												<TD colspan=7 height=5></TD>
											</TR>

											<TR height=5>
												<TD width=5></TD>

												<TD width=10></TD>
												<TD width=80></TD>
												<TD width=10></TD>
												<TD valign=top>
													<TABLE WIDTH="100%" class="invisible" CELLSPACING=0 CELLPADDING=0>
														<TR>
															<TD width="33%" align=left nowrap>
																<input type=checkbox name=chkColAverage id=chkColAverage tabindex="-1"
																    onclick="setAggregate(0)"
		                                                            onmouseover="try{checkbox_onMouseOver(this);}catch(e){}" 
		                                                            onmouseout="try{checkbox_onMouseOut(this);}catch(e){}" />
                                                                <label 
	                                                                for="chkColAverage"
	                                                                class="checkbox"
	                                                                tabindex="0"
	                                                                onkeypress="try{checkboxLabel_onKeyPress(this);}catch(e){}"
		                                                            onmouseover="try{checkboxLabel_onMouseOver(this);}catch(e){}" 
		                                                            onmouseout="try{checkboxLabel_onMouseOut(this);}catch(e){}"
                                                                    onfocus="try{checkboxLabel_onFocus(this);}catch(e){}"
                                                                    onblur="try{checkboxLabel_onBlur(this);}catch(e){}">
																    Average
                                        	    		        </label>
															</TD>
															<TD width="33%" align=left nowrap>
																<input type=checkbox name=chkColCount id=chkColCount tabindex="-1"
																    onclick="setAggregate(1)"
		                                                            onmouseover="try{checkbox_onMouseOver(this);}catch(e){}" 
		                                                            onmouseout="try{checkbox_onMouseOut(this);}catch(e){}" />
                                                                <label 
	                                                                for="chkColCount"
	                                                                class="checkbox"
	                                                                tabindex="0"
	                                                                onkeypress="try{checkboxLabel_onKeyPress(this);}catch(e){}"
		                                                            onmouseover="try{checkboxLabel_onMouseOver(this);}catch(e){}" 
		                                                            onmouseout="try{checkboxLabel_onMouseOut(this);}catch(e){}"
                                                                    onfocus="try{checkboxLabel_onFocus(this);}catch(e){}"
                                                                    onblur="try{checkboxLabel_onBlur(this);}catch(e){}">
																    Count 
                                        	    		        </label>
															</TD>
															<TD width="33%" align=left nowrap>
																<input type=checkbox name=chkColTotal id=chkColTotal tabindex="-1"
																    onclick="setAggregate(2)"
		                                                            onmouseover="try{checkbox_onMouseOver(this);}catch(e){}" 
		                                                            onmouseout="try{checkbox_onMouseOut(this);}catch(e){}" />
                                                                <label 
	                                                                for="chkColTotal"
	                                                                class="checkbox"
	                                                                tabindex="0"
	                                                                onkeypress="try{checkboxLabel_onKeyPress(this);}catch(e){}"
		                                                            onmouseover="try{checkboxLabel_onMouseOver(this);}catch(e){}" 
		                                                            onmouseout="try{checkboxLabel_onMouseOut(this);}catch(e){}"
                                                                    onfocus="try{checkboxLabel_onFocus(this);}catch(e){}"
                                                                    onblur="try{checkboxLabel_onBlur(this);}catch(e){}">
																    Total 
                                        	    		        </label>
															</TD>
														</TR>
														<TR Height=5>
															<TD colspan=3 height=5></TD>
														</TR>
														<TR>
															<TD width="33%" align=left nowrap>
																<input type=checkbox name=chkColHidden id=chkColHidden tabindex="-1"
																    onclick="setAggregate(3);"
		                                                            onmouseover="try{checkbox_onMouseOver(this);}catch(e){}" 
		                                                            onmouseout="try{checkbox_onMouseOut(this);}catch(e){}" />
                                                                <label 
	                                                                for="chkColHidden"
	                                                                class="checkbox"
	                                                                tabindex="0"
	                                                                onkeypress="try{checkboxLabel_onKeyPress(this);}catch(e){}"
		                                                            onmouseover="try{checkboxLabel_onMouseOver(this);}catch(e){}" 
		                                                            onmouseout="try{checkboxLabel_onMouseOut(this);}catch(e){}"
                                                                    onfocus="try{checkboxLabel_onFocus(this);}catch(e){}"
                                                                    onblur="try{checkboxLabel_onBlur(this);}catch(e){}">
																    Hidden
                                        	    		        </label>
															</TD>
															<TD colspan=2 align=left nowrap>
																<input type=checkbox name=chkColGroup id=chkColGroup tabindex="-1"
																    onclick="setAggregate(4)"
		                                                            onmouseover="try{checkbox_onMouseOver(this);}catch(e){}" 
		                                                            onmouseout="try{checkbox_onMouseOut(this);}catch(e){}" />
                                                                <label 
	                                                                for="chkColGroup"
	                                                                class="checkbox"
	                                                                tabindex="0"
	                                                                onkeypress="try{checkboxLabel_onKeyPress(this);}catch(e){}"
		                                                            onmouseover="try{checkboxLabel_onMouseOver(this);}catch(e){}" 
		                                                            onmouseout="try{checkboxLabel_onMouseOut(this);}catch(e){}"
                                                                    onfocus="try{checkboxLabel_onFocus(this);}catch(e){}"
                                                                    onblur="try{checkboxLabel_onBlur(this);}catch(e){}">
																    Group With Next 
                                        	    		        </label>
															</TD>
														</TR>
													</TABLE>
												</TD>
												<TD width=5></TD>
											</TR>
											
											<TR Height=5>
												<TD colspan=7 height=5></TD>
											</TR>
										</TABLE>
									</td>
								</tr>
							</TABLE>
						</DIV>

						<!-- Fourth tab -->
						<DIV id=div4 style="visibility:hidden;display:none">
							<TABLE WIDTH="100%" height="100%" class="outline" cellspacing=0 cellpadding=5>
								<tr valign=top> 
									<td>
										<TABLE WIDTH="100%" class="invisible" CELLSPACING=0 CELLPADDING=0>
											<TR>
												<TD colspan=5 height=5></TD>
											</TR>
											
											<TR height=20>
												<TD width=5>&nbsp;</TD>
												<TD colspan=3>Sort Order :</TD>
												<TD width=5>&nbsp;</TD>
											</TR>

											<TR height=5>
												<TD width=5>&nbsp;</TD>
												<TD rowspan=11>
													<%Html.RenderPartial("Util_Def_CustomReports/ssOleDBGridSortOrder")%>
												</TD>

												<TD width=10>&nbsp;</TD>
												<TD width=100>
													<input type="button" id=cmdSortAdd name=cmdSortAdd class="btn" value="Add..." style="WIDTH: 100%" 
													    onclick="sortAdd()"
                                                        onmouseover="try{button_onMouseOver(this);}catch(e){}" 
                                                        onmouseout="try{button_onMouseOut(this);}catch(e){}"
                                                        onfocus="try{button_onFocus(this);}catch(e){}"
                                                        onblur="try{button_onBlur(this);}catch(e){}" />
												</TD>
												<TD width=5>&nbsp;</TD>
											</TR>

											<TR height=5>
												<TD colspan=4>&nbsp;</TD>
											</TR>
														
											<TR height=5>
												<TD width=5>&nbsp;</TD>
												<TD width=5>&nbsp;</TD>
												<TD width=100>
													<input type="button" id=cmdSortEdit name=cmdSortEdit class="btn" value="Edit..." style="WIDTH: 100%" 
													    onclick="sortEdit()"
                                                        onmouseover="try{button_onMouseOver(this);}catch(e){}" 
                                                        onmouseout="try{button_onMouseOut(this);}catch(e){}"
                                                        onfocus="try{button_onFocus(this);}catch(e){}"
                                                        onblur="try{button_onBlur(this);}catch(e){}" />
												</TD>
												<TD width=5>&nbsp;</TD>
											</TR>
														
											<TR height=5>
												<TD colspan=4>&nbsp;</TD>
											</TR>

											<TR height=5>
												<TD width=5>&nbsp;</TD>

												<TD width=5>&nbsp;</TD>
												<TD width=100>
													<input type="button" id=cmdSortRemove name=cmdSortRemove class="btn" value="Remove" style="WIDTH: 100%" 
													    onclick="sortRemove()"
                                                        onmouseover="try{button_onMouseOver(this);}catch(e){}" 
                                                        onmouseout="try{button_onMouseOut(this);}catch(e){}"
                                                        onfocus="try{button_onFocus(this);}catch(e){}"
                                                        onblur="try{button_onBlur(this);}catch(e){}" />
												</TD>
												<TD width=5>&nbsp;</TD>
											</TR>
														
											<TR height=5>
												<TD colspan=4>&nbsp;</TD>
											</TR>
														
											<TR height=5>
												<TD width=5>&nbsp;</TD>

												<TD width=5>&nbsp;</TD>
												<TD width=100>
													<input type="button" id=cmdSortRemoveAll name=cmdSortRemoveAll class="btn" value="Remove All" style="WIDTH: 100%" 
													    onclick="sortRemoveAll()"
                                                        onmouseover="try{button_onMouseOver(this);}catch(e){}" 
                                                        onmouseout="try{button_onMouseOut(this);}catch(e){}"
                                                        onfocus="try{button_onFocus(this);}catch(e){}"
                                                        onblur="try{button_onBlur(this);}catch(e){}" />
												</TD>
												<TD width=5>&nbsp;</TD>
											</TR>
																									
											<TR height=5>
												<TD colspan=4>&nbsp;</TD>
											</TR>
														
											<TR height=5>
												<TD width=5>&nbsp;</TD>
												<TD width=5>&nbsp;</TD>
												<TD width=100>
													<input type="button" id=cmdSortMoveUp name=cmdSortMoveUp class="btn" value="Move Up" style="WIDTH: 100%" 
													    onclick="sortMove(true)"
                                                        onmouseover="try{button_onMouseOver(this);}catch(e){}" 
                                                        onmouseout="try{button_onMouseOut(this);}catch(e){}"
                                                        onfocus="try{button_onFocus(this);}catch(e){}"
                                                        onblur="try{button_onBlur(this);}catch(e){}" />
												</TD>
												<TD width=5>&nbsp;</TD>
											</TR>
														
											<TR height=5>
												<TD colspan=4>&nbsp;</TD>
											</TR>
														
											<TR height=5>
												<TD width=5>&nbsp;</TD>
												<TD width=5>&nbsp;</TD>
												<TD width=100>
													<input type="button" id=cmdSortMoveDown name=cmdSortMoveDown class="btn" value="Move Down" style="WIDTH: 100%" 
													    onclick="sortMove(false)"
                                                        onmouseover="try{button_onMouseOver(this);}catch(e){}" 
                                                        onmouseout="try{button_onMouseOut(this);}catch(e){}"
                                                        onfocus="try{button_onFocus(this);}catch(e){}"
                                                        onblur="try{button_onBlur(this);}catch(e){}" />
												</TD>

												<TD width=5>&nbsp;</TD>
											</TR>
											
											<TR height=5>
												<TD colspan=5></TD>
											</TR>
											
											<TR height=20>
												<TD colspan=5 valign=center>
													<hr>
												</TD>
											</TR>

											<TR height=20>
												<TD width=5>&nbsp;</TD>
												<TD colspan=3>Repetition :</TD>
												<TD width=5>&nbsp;</TD>
											</TR>

											<TR height=5>
												<TD width=5>&nbsp;</TD>
												<TD colspan=3 rowspan=9>
													<%Html.RenderPartial("Util_Def_CustomReports/ssOleDBGridRepetition")%>
												</TD>

												<TD width=5>&nbsp;</TD>
											</TR>
											
											<TR height=5>
												<TD colspan=5>&nbsp;</TD>
											</TR>
														
											<TR height=5>
												<TD colspan=5>&nbsp;</TD>
											</TR>
														
											<TR height=5>
												<TD colspan=5>&nbsp;</TD>
											</TR>

											<TR height=5>
												<TD colspan=5>&nbsp;</TD>
											</TR>
														
											<TR height=5>
												<TD colspan=5>&nbsp;</TD>
											</TR>
														
											<TR height=5>
												<TD colspan=5>&nbsp;</TD>
											</TR>
														
											<TR height=5>
												<TD colspan=5>&nbsp;</TD>
											</TR>
														
											<TR height=5>
												<TD colspan=5>&nbsp;</TD>
											</TR>

										</TABLE>
									</td>
								</tr>

							</TABLE>
						</DIV>

						<!-- Fifth tab -->
						<DIV id="div5" style="visibility:hidden;display:none">
							<TABLE WIDTH="100%" height="100%" class="outline" cellspacing=0 cellpadding=5>
								<tr valign=top> 
									<td>
										<TABLE WIDTH="100%" class="invisible" CELLSPACING=10 CELLPADDING=0>
											<TR>
												<TD valign=top colspan=2 width=100% height=65>
													<table class="outline" cellspacing="0" cellpadding="4" width=100% height=100%>
														<tr height=10> 
															<td height=10 align=left valign=top>
																Report Options : <BR><BR>
																<TABLE class="invisible" cellspacing="0" cellpadding="0" width=100%>
																	<tr height=20>
																		<td width=5>&nbsp</td>
																		<td align=left>
																			<INPUT type="checkbox" id=chkSummary name=chkSummary tabindex="-1"
																			    onclick="changeTab5Control()"
																    		    onmouseover="try{checkbox_onMouseOver(this);}catch(e){}" 
		                                                                        onmouseout="try{checkbox_onMouseOut(this);}catch(e){}" />
                                                                            <label 
	                                                                            for="chkSummary"
	                                                                            class="checkbox"
	                                                                            tabindex=0 
	                                                                            onkeypress="try{checkboxLabel_onKeyPress(this);}catch(e){}"
		                                                                        onmouseover="try{checkboxLabel_onMouseOver(this);}catch(e){}" 
		                                                                        onmouseout="try{checkboxLabel_onMouseOut(this);}catch(e){}"
                                                                                onfocus="try{checkboxLabel_onFocus(this);}catch(e){}"
                                                                                onblur="try{checkboxLabel_onBlur(this);}catch(e){}">Summary report</label>
																		</td>
																		<td width=5>&nbsp</td>
																	</tr>
																	<tr height=10> 
																		<td colspan=3></td>
																	</tr>
																	<tr height=20>
																		<td width=5>&nbsp</td>
																		<td align=left>
																			<INPUT type="checkbox" id=chkIgnoreZeros name=chkIgnoreZeros tabindex="-1"
																			    onclick="changeTab5Control()"
																    		    onmouseover="try{checkbox_onMouseOver(this);}catch(e){}" 
		                                                                        onmouseout="try{checkbox_onMouseOut(this);}catch(e){}" />
                                                                            <label 
	                                                                            for="chkIgnoreZeros"
	                                                                            class="checkbox"
	                                                                            tabindex=0 
	                                                                            onkeypress="try{checkboxLabel_onKeyPress(this);}catch(e){}"
		                                                                        onmouseover="try{checkboxLabel_onMouseOver(this);}catch(e){}" 
		                                                                        onmouseout="try{checkboxLabel_onMouseOut(this);}catch(e){}"
                                                                                onfocus="try{checkboxLabel_onFocus(this);}catch(e){}"
                                                                                onblur="try{checkboxLabel_onBlur(this);}catch(e){}">Ignore zeros when calculating aggregates</label>
																	</td>
																		<td width=5>&nbsp</td>
																	</tr>
																	<tr height=10> 
																		<td colspan=3></td>
																	</tr>
																</TABLE>
															</TD>
														</tr>
													</table>
												</TD>
											</TR>
											<tr>						
												<td valign=top rowspan=2 width=25% height="100%">
													<table class="outline" cellspacing="0" cellpadding="4" width=100% height=100%>
														<tr height=10> 
															<td height=10 align=left valign=top>
																Output Format : <BR><BR>
																<TABLE class="invisible" cellspacing="0" cellpadding="0" width=100%>
																	<tr height=20>
																		<td width=5>&nbsp</td>
																		<td align=left width=15>
																			<INPUT type=radio width=20 style="WIDTH: 20px" name=optOutputFormat id=optOutputFormat0 value=0
																			    onClick="formatClick(0);" 
                                                                                onmouseover="try{radio_onMouseOver(this);}catch(e){}" 
                                                                                onmouseout="try{radio_onMouseOut(this);}catch(e){}"
                                                                                onfocus="try{radio_onFocus(this);}catch(e){}"
                                                                                onblur="try{radio_onBlur(this);}catch(e){}"/>
																		</td>
																		<td align=left nowrap>
                                                                            <label 
                                                                                tabindex=-1
                                                                                for="optOutputFormat0"
                                                                                class="radio"
                                                                                onmouseover="try{radioLabel_onMouseOver(this);}catch(e){}" 
                                                                                onmouseout="try{radioLabel_onMouseOut(this);}catch(e){}">Data Only</label>
																		</td>
																		<td width=5>&nbsp</td>
																	</tr>
																	<tr height=10> 
																		<td colspan=4></td>
																	</tr>
																	<tr height=20>
																		<td width=5>&nbsp</td>
																		<td align=left width=15>
																			<INPUT type=radio width=20 style="WIDTH: 20px" name=optOutputFormat id=optOutputFormat1 value=1
																			    onClick="formatClick(1);" 
                                                                                onmouseover="try{radio_onMouseOver(this);}catch(e){}" 
                                                                                onmouseout="try{radio_onMouseOut(this);}catch(e){}"
                                                                                onfocus="try{radio_onFocus(this);}catch(e){}"
                                                                                onblur="try{radio_onBlur(this);}catch(e){}"/>
																		</td>
																		<td align=left nowrap>
                                                                            <label 
                                                                                tabindex=-1
                                                                                for="optOutputFormat1"
                                                                                class="radio"
                                                                                onmouseover="try{radioLabel_onMouseOver(this);}catch(e){}" 
                                                                                onmouseout="try{radioLabel_onMouseOut(this);}catch(e){}">CSV File</label>
																		</td>
																		<td width=5>&nbsp</td>
																	</tr>
																	<tr height=10> 
																		<td colspan=4></td>
																	</tr>
																	<tr height=20>
																		<td width=5>&nbsp</td>
																		<td align=left width=15>
																			<INPUT type=radio width=20 style="WIDTH: 20px" name=optOutputFormat id=optOutputFormat2 value=2
																			    onClick="formatClick(2);" 
                                                                                onmouseover="try{radio_onMouseOver(this);}catch(e){}" 
                                                                                onmouseout="try{radio_onMouseOut(this);}catch(e){}"
                                                                                onfocus="try{radio_onFocus(this);}catch(e){}"
                                                                                onblur="try{radio_onBlur(this);}catch(e){}"/>
																		</td>
																		<td align=left nowrap>
                                                                            <label 
                                                                                tabindex=-1
                                                                                for="optOutputFormat2"
                                                                                class="radio"
                                                                                onmouseover="try{radioLabel_onMouseOver(this);}catch(e){}" 
                                                                                onmouseout="try{radioLabel_onMouseOut(this);}catch(e){}">HTML Document</label>
																		</td>
																		<td width=5>&nbsp</td>
																	</tr>
																	<tr height=10> 
																		<td colspan=4></td>
																	</tr>
																	<tr height=20>
																		<td width=5>&nbsp</td>
																		<td align=left width=15>
																			<INPUT type=radio width=20 style="WIDTH: 20px" name=optOutputFormat id=optOutputFormat3 value=3
																			    onClick="formatClick(3);" 
                                                                                onmouseover="try{radio_onMouseOver(this);}catch(e){}" 
                                                                                onmouseout="try{radio_onMouseOut(this);}catch(e){}"
                                                                                onfocus="try{radio_onFocus(this);}catch(e){}"
                                                                                onblur="try{radio_onBlur(this);}catch(e){}"/>
																		</td>
																		<td align=left nowrap>
                                                                            <label 
                                                                                tabindex=-1
                                                                                for="optOutputFormat3"
                                                                                class="radio"
                                                                                onmouseover="try{radioLabel_onMouseOver(this);}catch(e){}" 
                                                                                onmouseout="try{radioLabel_onMouseOut(this);}catch(e){}">Word Document</label>
																		</td>
																		<td width=5>&nbsp</td>
																	</tr>
																	<tr height=10> 
																		<td colspan=4></td>
																	</tr>
																	<tr height=20>
																		<td width=5>&nbsp</td>
																		<td align=left width=15>
																			<INPUT type=radio width=20 style="WIDTH: 20px" name=optOutputFormat id=optOutputFormat4 value=4
																			    onClick="formatClick(4);" 
                                                                                onmouseover="try{radio_onMouseOver(this);}catch(e){}" 
                                                                                onmouseout="try{radio_onMouseOut(this);}catch(e){}"
                                                                                onfocus="try{radio_onFocus(this);}catch(e){}"
                                                                                onblur="try{radio_onBlur(this);}catch(e){}"/>
																		</td>
																		<td align=left nowrap>
                                                                            <label 
                                                                                tabindex=-1
                                                                                for="optOutputFormat4"
                                                                                class="radio"
                                                                                onmouseover="try{radioLabel_onMouseOver(this);}catch(e){}" 
                                                                                onmouseout="try{radioLabel_onMouseOut(this);}catch(e){}">Excel Worksheet</label>
																		</td>
																		<td width=5>&nbsp</td>
																	</tr>
																	<tr height=10> 
																		<td colspan=4></td>
																	</tr>
																	<tr height=5>
																		<td width=5>&nbsp</td>
																		<td align=left width=15>
																			<INPUT type=radio width=20 style="WIDTH: 20px" name=optOutputFormat id=optOutputFormat5 value=5
																			    onClick="formatClick(5);" 
                                                                                onmouseover="try{radio_onMouseOver(this);}catch(e){}" 
                                                                                onmouseout="try{radio_onMouseOut(this);}catch(e){}"
                                                                                onfocus="try{radio_onFocus(this);}catch(e){}"
                                                                                onblur="try{radio_onBlur(this);}catch(e){}"/>
																		</td>
																		<td>
                                                                            <label 
                                                                                tabindex=-1
                                                                                for="optOutputFormat5"
                                                                                class="radio"
                                                                                onmouseover="try{radioLabel_onMouseOver(this);}catch(e){}" 
                                                                                onmouseout="try{radioLabel_onMouseOut(this);}catch(e){}">Excel Chart</label>
																		</td>
																		<td width=5>&nbsp</td>
																	</tr>
																	<tr height=10> 
																		<td colspan=4></td>
																	</tr>
																	<tr height=5>
																		<td width=5>&nbsp</td>
																		<td align=left width=15>
																			<INPUT type=radio width=20 style="WIDTH: 20px" name=optOutputFormat id=optOutputFormat6 value=6
																			    onClick="formatClick(6);" 
                                                                                onmouseover="try{radio_onMouseOver(this);}catch(e){}" 
                                                                                onmouseout="try{radio_onMouseOut(this);}catch(e){}"
                                                                                onfocus="try{radio_onFocus(this);}catch(e){}"
                                                                                onblur="try{radio_onBlur(this);}catch(e){}"/>
																		</td>
																		<td nowrap>
                                                                            <label 
                                                                                tabindex=-1
                                                                                for="optOutputFormat6"
                                                                                class="radio"
                                                                                onmouseover="try{radioLabel_onMouseOver(this);}catch(e){}" 
                                                                                onmouseout="try{radioLabel_onMouseOut(this);}catch(e){}">Excel Pivot Table</label>
																		</td>
																		<td width=5>&nbsp</td>
																	</tr>
																	<tr height=5> 
																		<td colspan=4></td>
																	</tr>
																</TABLE>
															</td>
														</tr>
													</table>
												</td>
												<td valign=top width=75%>
													<table class="outline" cellspacing="0" cellpadding="4" width=100%  height=100%>
														<tr height=10> 
															<td height=10 align=left valign=top>
																Output Destination(s) : <BR><BR>
																
																<TABLE class="invisible" cellspacing="0" cellpadding="0" width=100%>
																	<tr height=20>
																		<td width=5>&nbsp</td>
																		<td align=left colspan=6 nowrap>
																			<input name=chkPreview id=chkPreview type=checkbox disabled="disabled" tabindex="-1"
																			    onClick="changeTab5Control();"
		                                                                        onmouseover="try{checkbox_onMouseOver(this);}catch(e){}" 
		                                                                        onmouseout="try{checkbox_onMouseOut(this);}catch(e){}" />
                                                                            <label 
	                                                                            for="chkPreview"
	                                                                            class="checkbox checkboxdisabled"
	                                                                            tabindex=0 
	                                                                            onkeypress="try{checkboxLabel_onKeyPress(this);}catch(e){}"
		                                                                        onmouseover="try{checkboxLabel_onMouseOver(this);}catch(e){}" 
		                                                                        onmouseout="try{checkboxLabel_onMouseOut(this);}catch(e){}"
                                                                                onfocus="try{checkboxLabel_onFocus(this);}catch(e){}"
                                                                                onblur="try{checkboxLabel_onBlur(this);}catch(e){}">Preview on screen</label>
																		</td>
																		<td width=5>&nbsp</td>
																	</tr>
																	
																	<tr height=10> 
																		<td colspan=8></td>
																	</tr>
																	
																	<tr height=20>
																		<td></td>
																		<td align=left colspan=6 nowrap>
																			<input name=chkDestination0 id=chkDestination0 type=checkbox disabled="disabled" tabindex="-1"
																			    onClick="changeTab5Control();"
		                                                                        onmouseover="try{checkbox_onMouseOver(this);}catch(e){}" 
		                                                                        onmouseout="try{checkbox_onMouseOut(this);}catch(e){}" />
                                                                            <label 
	                                                                            for="chkDestination0"
	                                                                            class="checkbox checkboxdisabled"
	                                                                            tabindex=0 
	                                                                            onkeypress="try{checkboxLabel_onKeyPress(this);}catch(e){}"
		                                                                        onmouseover="try{checkboxLabel_onMouseOver(this);}catch(e){}" 
		                                                                        onmouseout="try{checkboxLabel_onMouseOut(this);}catch(e){}"
                                                                                onfocus="try{checkboxLabel_onFocus(this);}catch(e){}"
                                                                                onblur="try{checkboxLabel_onBlur(this);}catch(e){}">Display output on screen</label>
																		</td>
																		<td></td>
																	</tr>
																	
																	<tr height=10> 
																		<td colspan=8></td>
																	</tr>
																	
																	<tr height=20>
																		<td></td>
																		<td align=left nowrap>
																			<input name=chkDestination1 id=chkDestination1 type=checkbox disabled="disabled" tabindex="-1"
																			    onClick="changeTab5Control();"
		                                                                        onmouseover="try{checkbox_onMouseOver(this);}catch(e){}" 
		                                                                        onmouseout="try{checkbox_onMouseOut(this);}catch(e){}" />
                                                                            <label 
	                                                                            for="chkDestination1"
	                                                                            class="checkbox checkboxdisabled"
	                                                                            tabindex=0 
	                                                                            onkeypress="try{checkboxLabel_onKeyPress(this);}catch(e){}"
		                                                                        onmouseover="try{checkboxLabel_onMouseOver(this);}catch(e){}" 
		                                                                        onmouseout="try{checkboxLabel_onMouseOut(this);}catch(e){}"
                                                                                onfocus="try{checkboxLabel_onFocus(this);}catch(e){}"
                                                                                onblur="try{checkboxLabel_onBlur(this);}catch(e){}">Send to printer</label>
																		</td>
																		<td width=30 nowrap>&nbsp</td>
																		<td align=left nowrap>
																			Printer location : 
																		</td>
																		<td width=15>&nbsp</td>
																		<td colspan=2>
																			<select id=cboPrinterName name=cboPrinterName class="combo" style="WIDTH: 400px" 
																			    onchange="changeTab5Control()">
																			</select>								
																		</td>
																		<td></td>
																	</tr>
																	
																	<tr height=10> 
																		<td colspan=8></td>
																	</tr>
																	
																	<tr height=20>
																		<td></td>
																		<td align=left nowrap>
																			<input name=chkDestination2 id=chkDestination2 type=checkbox disabled="disabled" tabindex="-1"
																			    onClick="changeTab5Control();"
		                                                                        onmouseover="try{checkbox_onMouseOver(this);}catch(e){}" 
		                                                                        onmouseout="try{checkbox_onMouseOut(this);}catch(e){}" />
                                                                            <label 
	                                                                            for="chkDestination2"
	                                                                            class="checkbox checkboxdisabled"
	                                                                            tabindex=0 
	                                                                            onkeypress="try{checkboxLabel_onKeyPress(this);}catch(e){}"
		                                                                        onmouseover="try{checkboxLabel_onMouseOver(this);}catch(e){}" 
		                                                                        onmouseout="try{checkboxLabel_onMouseOut(this);}catch(e){}"
                                                                                onfocus="try{checkboxLabel_onFocus(this);}catch(e){}"
                                                                                onblur="try{checkboxLabel_onBlur(this);}catch(e){}">
																			    Save to file
                                                        	    		    </label>
																		</td>
																		<td nowrap></td>
																		<td align=left nowrap>
																			File name :   
																		</td>
																		<td nowrap></td>
																		<td colspan=2>
																			<TABLE class="invisible" CELLSPACING=0 CELLPADDING=0 style=" WIDTH: 400px">
																				<TR>
																					<TD>
																						<INPUT id=txtFilename name=txtFilename class="text textdisabled" disabled="disabled" tabindex="-1" style="WIDTH: 375px">
																					</TD>
																					<TD width=25>
																						<INPUT id=cmdFilename name=cmdFilename class="btn" style="WIDTH: 100%" type=button value="..."
																						    onClick="saveFile();changeTab5Control();" 
	                                                                                        onmouseover="try{button_onMouseOver(this);}catch(e){}" 
	                                                                                        onmouseout="try{button_onMouseOut(this);}catch(e){}"
	                                                                                        onfocus="try{button_onFocus(this);}catch(e){}"
	                                                                                        onblur="try{button_onBlur(this);}catch(e){}" />
																					</TD>
																				</TR>
																			</TABLE>
																		</td>
																		<td width=5>&nbsp</td>
																	</tr>

																	<tr height=10> 
																		<td colspan=8></td>
																	</tr>
																	
																	<tr height=20>
																		<td colspan=3></td>
																		<td align=left nowrap>
																			If existing file :
																		</td>
																		<td></td>
																		<td colspan=2 width=100% nowrap>
																			<select id=cboSaveExisting name=cboSaveExisting style="WIDTH: 400px" class="combo" onchange="changeTab5Control()"></select>								
																		</td>
																		<td></td>
																	</tr>

																	<tr height=10> 
																		<td colspan=8></td>
																	</tr>
																	
																	<tr height=20>
																		<td></td>
																		<td align=left nowrap>
																			<input name=chkDestination3 id=chkDestination3 type=checkbox disabled="disabled" tabindex="-1"
																			    onClick="changeTab5Control();"
		                                                                        onmouseover="try{checkbox_onMouseOver(this);}catch(e){}" 
		                                                                        onmouseout="try{checkbox_onMouseOut(this);}catch(e){}" />
                                                                            <label 
	                                                                            for="chkDestination3"
	                                                                            class="checkbox checkboxdisabled"
	                                                                            tabindex=0 
	                                                                            onkeypress="try{checkboxLabel_onKeyPress(this);}catch(e){}"
		                                                                        onmouseover="try{checkboxLabel_onMouseOver(this);}catch(e){}" 
		                                                                        onmouseout="try{checkboxLabel_onMouseOut(this);}catch(e){}"
                                                                                onfocus="try{checkboxLabel_onFocus(this);}catch(e){}"
                                                                                onblur="try{checkboxLabel_onBlur(this);}catch(e){}">Send as email</label>
																		</td>
																		<td></td>
																		<td align=left nowrap>
																			Email group :   
																		</td>
																		<td></td>
																		<td colspan=2>
																			<TABLE class="invisible" CELLSPACING=0 CELLPADDING=0 style="WIDTH: 400px">
																				<TR>
																					<TD>
																						<INPUT id=txtEmailGroup name=txtEmailGroup class="text textdisabled" disabled="disabled" tabindex="-1" style="WIDTH: 100%">
																						<INPUT id=txtEmailGroupID name=txtEmailGroupID type=hidden>
																					</TD>
																					<TD width=25>
																						<INPUT id=cmdEmailGroup name=cmdEmailGroup style="WIDTH: 100%" type=button value="..." class="btn"
																						    onClick="selectEmailGroup();changeTab5Control();" 
	                                                                                        onmouseover="try{button_onMouseOver(this);}catch(e){}" 
	                                                                                        onmouseout="try{button_onMouseOut(this);}catch(e){}"
	                                                                                        onfocus="try{button_onFocus(this);}catch(e){}"
	                                                                                        onblur="try{button_onBlur(this);}catch(e){}" />
																					</TD>
																				</TR>
																			</TABLE>
																		</td>
																		<td></td>
																	</tr>
																	
																	<tr height=10> 
																		<td colspan=8></td>
																	</tr>
																	
																	<tr height=20>
																		<td colspan=3></td>
																		<td align=left nowrap>
																			Email subject :   
																		</td>
																		<td></td>
																		<TD colspan=2 width=100% nowrap>
																			<INPUT id=txtEmailSubject class="text textdisabled" disabled="disabled" maxlength=255 name=txtEmailSubject style=" WIDTH: 400px" 
																			    onchange="frmUseful.txtChanged.value = 1;" 
																			    onkeydown="frmUseful.txtChanged.value = 1;">
																		</TD>
																		<td width=5>&nbsp</td>
																	</tr>
																	
																	<tr height=10> 
																		<td colspan=8></td>
																	</tr>
																	
																	<tr height=20>
																		<td colspan=3></td>
																		<td align=left nowrap>
																			Attach as :   
																		</td>
																		<td></td>
																		<TD colspan=2 width=100% nowrap>
																			<INPUT id=txtEmailAttachAs class="text textdisabled" disabled="disabled" maxlength=255 name=txtEmailAttachAs style=" WIDTH: 400px" 
																			    onchange="frmUseful.txtChanged.value = 1;" 
																			    onkeydown="frmUseful.txtChanged.value = 1;">
																		</TD>
																		<td></td>
																	</tr>
																	
																	<tr height=10> 
																		<td colspan=8></td>
																	</tr>
																</TABLE>
															</td>
														</tr>
													</table>
												</td>
											</tr>
										</TABLE>
									</td>
								</tr>
							</TABLE>
						</DIV>
						
					</td>
					<TD width=10></td>
				</tr> 

				<tr height=10> 
					<td colspan=3></td>
				</tr> 

				<TR height=10>
					<TD width=10></td>
					<TD>
						<TABLE WIDTH="100%" class="invisible" CELLSPACING=0 CELLPADDING=0>
							<TR>
								<TD>&nbsp;</TD>
								<TD width=80>
									<input type=button id=cmdOK name=cmdOK value=OK style="WIDTH: 100%" class="btn"
									    onclick="okClick()"
                                        onmouseover="try{button_onMouseOver(this);}catch(e){}" 
                                        onmouseout="try{button_onMouseOut(this);}catch(e){}"
                                        onfocus="try{button_onFocus(this);}catch(e){}"
                                        onblur="try{button_onBlur(this);}catch(e){}" />
								</TD>
								<TD width=10></TD>
								<TD width=80>
									<input type=button id=cmdCancel name=cmdCancel value=Cancel style="WIDTH: 100%"  class="btn" 
									    onclick="cancelClick()"
                                        onmouseover="try{button_onMouseOver(this);}catch(e){}" 
                                        onmouseout="try{button_onMouseOut(this);}catch(e){}"
                                        onfocus="try{button_onFocus(this);}catch(e){}"
                                        onblur="try{button_onBlur(this);}catch(e){}" />
								</TD>
							</TR>
						</TABLE>
					</td>
					<TD width=10></td>
				</tr> 

				<tr height=5> 
					<td colspan=3></td>
				</tr> 
			</TABLE>
		</td>
	</tr> 
</table>

<INPUT type='hidden' id=txtBasePicklistID name=txtBasePicklistID>
<INPUT type='hidden' id=txtBaseFilterID name=txtBaseFilterID>

<INPUT type='hidden' id=txtParent1ID name=txtParent1ID>
<INPUT type='hidden' id=txtParent2ID name=txtParent2ID>
<INPUT type='hidden' id=txtParent1FilterID name=txtParent1FilterID>
<INPUT type='hidden' id=txtParent1PicklistID name=txtParent1PicklistID>
<INPUT type='hidden' id=txtParent2FilterID name=txtParent2FilterID>
<INPUT type='hidden' id=txtParent2PicklistID name=txtParent2PicklistID>

<INPUT type='hidden' id=txtBaseTableChildCount name=txtBaseTableChildCount>
<INPUT type='hidden' id=txtDatabase name=txtDatabase value="<%=session("Database")%>">

<INPUT type='hidden' id=txtWordVer name=txtWordVer value="<%=Session("WordVer")%>">
<INPUT type='hidden' id=txtExcelVer name=txtExcelVer value="<%=Session("ExcelVer")%>">
<INPUT type='hidden' id=txtWordFormats name=txtWordFormats value="<%=Session("WordFormats")%>">
<INPUT type='hidden' id=txtExcelFormats name=txtExcelFormats value="<%=Session("ExcelFormats")%>">
<INPUT type='hidden' id=txtWordFormatDefaultIndex name=txtWordFormatDefaultIndex value="<%=Session("WordFormatDefaultIndex")%>">
<INPUT type='hidden' id=txtExcelFormatDefaultIndex name=txtExcelFormatDefaultIndex value="<%=Session("ExcelFormatDefaultIndex")%>">

<input type='hidden' id="txtChangeCancelled" name="txtChangeCancelled" value="0">
<input type='hidden' id="txtCheckingSuppressOptions" name="txtCheckingSuppressOptions" value="0">
</form>

<form id=frmTables style="visibility:hidden;display:none">
<%
	Dim sErrorDescription = ""
	
	' Get the table records.
	Dim cmdTables = CreateObject("ADODB.Command")
	cmdTables.CommandText = "sp_ASRIntGetTablesInfo"
	cmdTables.CommandType = 4 ' Stored Procedure
	cmdTables.ActiveConnection = Session("databaseConnection")
	
	Response.Write("<B>Set Connection</B>")
	
	Err.Clear()
	Dim rstTablesInfo = cmdTables.Execute
	
	Response.Write("<B>Executed SP</B>")
	
	If (Err.Number <> 0) Then
		sErrorDescription = "The tables information could not be retrieved." & vbCrLf & formatError(Err.Description)
	End If

	if len(sErrorDescription) = 0 then
		Dim iCount = 0
		do while not rstTablesInfo.EOF
			Response.Write("<INPUT type='hidden' id=txtTableName_" & rstTablesInfo.fields("tableID").value & " name=txtTableName_" & rstTablesInfo.fields("tableID").value & " value=""" & rstTablesInfo.fields("tableName").value & """>" & vbCrLf)
			Response.Write("<INPUT type='hidden' id=txtTableType_" & rstTablesInfo.fields("tableID").value & " name=txtTableType_" & rstTablesInfo.fields("tableID").value & " value=" & rstTablesInfo.fields("tableType").value & ">" & vbCrLf)
			Response.Write("<INPUT type='hidden' id=txtTableChildren_" & rstTablesInfo.fields("tableID").value & " name=txtTableChildren_" & rstTablesInfo.fields("tableID").value & " value=""" & rstTablesInfo.fields("childrenString").value & """>" & vbCrLf)
			Response.Write("<INPUT type='hidden' id=txtTableChildrenNames_" & rstTablesInfo.fields("tableID").value & " name=txtTableChildrenNames_" & rstTablesInfo.fields("tableID").value & " value=""" & rstTablesInfo.fields("childrenNames").value & """>" & vbCrLf)
			Response.Write("<INPUT type='hidden' id=txtTableParents_" & rstTablesInfo.fields("tableID").value & " name=txtTableParents_" & rstTablesInfo.fields("tableID").value & " value=""" & rstTablesInfo.fields("parentsString").value & """>" & vbCrLf)

			rstTablesInfo.MoveNext
		loop

		' Release the ADO recordset object.
		rstTablesInfo.close
		rstTablesInfo = Nothing
	end if
	
	' Release the ADO command object.
	cmdTables = Nothing
%>
</form>

<FORM action="default_Submit" method=post id=frmGoto name=frmGoto style="visibility:hidden;display:none">
	<%Html.RenderPartial("~/Views/Shared/gotoWork.ascx")%>
</FORM>

<form id=frmOriginalDefinition name=frmOriginalDefinition style="visibility:hidden;display:none">
<%
	Dim sErrMsg = ""

	if session("action") <> "new"	then
		Dim cmdDefn = CreateObject("ADODB.Command")
		cmdDefn.CommandText = "sp_ASRIntGetReportDefinition"
		cmdDefn.CommandType = 4 ' Stored Procedure
		cmdDefn.ActiveConnection = Session("databaseConnection")
		
		Dim prmUtilID = cmdDefn.CreateParameter("utilID", 3, 1) ' 3=integer, 1=input
		cmdDefn.Parameters.Append(prmUtilID)
		prmUtilID.value = cleanNumeric(session("utilid"))

		Dim prmUser = cmdDefn.CreateParameter("user", 200, 1, 8000)	' 200=varchar, 1=input, 8000=size
		cmdDefn.Parameters.Append(prmUser)
		prmUser.value = session("username")

		Dim prmAction = cmdDefn.CreateParameter("action", 200, 1, 8000)	' 200=varchar, 1=input, 8000=size
		cmdDefn.Parameters.Append(prmAction)
		prmAction.value = session("action")

		Dim prmErrMsg = cmdDefn.CreateParameter("errMsg", 200, 2, 8000)	'200=varchar, 2=output, 8000=size
		cmdDefn.Parameters.Append(prmErrMsg)

		Dim prmName = cmdDefn.CreateParameter("name", 200, 2, 8000)	'200=varchar, 2=output, 8000=size
		cmdDefn.Parameters.Append(prmName)

		Dim prmOwner = cmdDefn.CreateParameter("owner", 200, 2, 8000) '200=varchar, 2=output, 8000=size
		cmdDefn.Parameters.Append(prmOwner)

		Dim prmDescription = cmdDefn.CreateParameter("description", 200, 2, 8000) '200=varchar, 2=output, 8000=size
		cmdDefn.Parameters.Append(prmDescription)

		Dim prmBaseTableID = cmdDefn.CreateParameter("baseTableID", 3, 2)	'3=integer, 2=output
		cmdDefn.Parameters.Append(prmBaseTableID)

		Dim prmAllRecords = cmdDefn.CreateParameter("allRecords", 11, 2) '11=bit, 2=output
		cmdDefn.Parameters.Append(prmAllRecords)

		Dim prmPicklistID = cmdDefn.CreateParameter("picklistID", 3, 2)	'3=integer, 2=output
		cmdDefn.Parameters.Append(prmPicklistID)

		Dim prmPicklistName = cmdDefn.CreateParameter("picklistName", 200, 2, 8000)	'200=varchar, 2=output, 8000=size
		cmdDefn.Parameters.Append(prmPicklistName)

		Dim prmPicklistHidden = cmdDefn.CreateParameter("picklistHidden", 11, 2)	'11=bit, 2=output
		cmdDefn.Parameters.Append(prmPicklistHidden)

		Dim prmFilterID = cmdDefn.CreateParameter("filterID", 3, 2)	'3=integer, 2=output
		cmdDefn.Parameters.Append(prmFilterID)

		Dim prmFilterName = cmdDefn.CreateParameter("filterName", 200, 2, 8000)	'200=varchar, 2=output, 8000=size
		cmdDefn.Parameters.Append(prmFilterName)

		Dim prmFilterHidden = cmdDefn.CreateParameter("filterHidden", 11, 2)	'11=bit, 2=output
		cmdDefn.Parameters.Append(prmFilterHidden)

		Dim prmParent1TableID = cmdDefn.CreateParameter("parent1TableID", 3, 2)	'3=integer, 2=output
		cmdDefn.Parameters.Append(prmParent1TableID)

		Dim prmParent1TableName = cmdDefn.CreateParameter("parent1TableName", 200, 2, 8000)	'200=varchar, 2=output, 8000=size
		cmdDefn.Parameters.Append(prmParent1TableName)

		Dim prmParent1FilterID = cmdDefn.CreateParameter("parent1FilterID", 3, 2) '3=integer, 2=output
		cmdDefn.Parameters.Append(prmParent1FilterID)

		Dim prmParent1FilterName = cmdDefn.CreateParameter("parent1FilterName", 200, 2, 8000) '200=varchar, 2=output, 8000=size
		cmdDefn.Parameters.Append(prmParent1FilterName)

		Dim prmParent1FilterHidden = cmdDefn.CreateParameter("parent1FilterHidden", 11, 2) '11=bit, 2=output
		cmdDefn.Parameters.Append(prmParent1FilterHidden)

		Dim prmParent2TableID = cmdDefn.CreateParameter("parent2TableID", 3, 2)	'3=integer, 2=output
		cmdDefn.Parameters.Append(prmParent2TableID)

		Dim prmParent2TableName = cmdDefn.CreateParameter("parent2TableName", 200, 2, 8000)	'200=varchar, 2=output, 8000=size
		cmdDefn.Parameters.Append(prmParent2TableName)

		Dim prmParent2FilterID = cmdDefn.CreateParameter("parent2FilterID", 3, 2) '3=integer, 2=output
		cmdDefn.Parameters.Append(prmParent2FilterID)

		Dim prmParent2FilterName = cmdDefn.CreateParameter("parent2FilterName", 200, 2, 8000) '200=varchar, 2=output, 8000=size
		cmdDefn.Parameters.Append(prmParent2FilterName)
		
		Dim prmParent2FilterHidden = cmdDefn.CreateParameter("parent2FilterHidden", 11, 2) '11=bit, 2=output
		cmdDefn.Parameters.Append(prmParent2FilterHidden)

'********************************************************************************

		Dim cmdReportChilds = CreateObject("ADODB.Command")
		cmdReportChilds.CommandText = "sp_ASRIntGetReportChilds"
		cmdReportChilds.CommandType = 4	'Stored Procedure
		cmdReportChilds.ActiveConnection = Session("databaseConnection")
		
		Dim prmUtilID2 = cmdReportChilds.CreateParameter("utilID2", 3, 1)	' 3=integer, 1=input
		cmdReportChilds.Parameters.Append(prmUtilID2)
		prmUtilID2.value = cleanNumeric(session("utilid"))
		
		Err.Clear()
		Dim rstChilds = cmdReportChilds.Execute
		Dim iHiddenChildFilterCount = 0
		Dim iCount = 0
		Dim sChildInfo = ""
		
		If (Err.Number <> 0) Then
			sErrMsg = "'" & Session("utilname") & "' report definition could not be read." & vbCrLf & formatError(Err.Description)
		Else
			If rstChilds.state <> 0 Then
				' Read recordset values.
				
				Do While Not rstChilds.EOF
					iCount = iCount + 1
					Response.Write("<INPUT type='hidden' id=txtReportDefnChildTableID_" & iCount & " name=txtReportDefnChildTableID_" & iCount & " value=""" & rstChilds.fields("TableID").value & """>" & vbCrLf)
					Response.Write("<INPUT type='hidden' id=txtReportDefnChildTable_" & iCount & " name=txtReportDefnChildTable_" & iCount & " value=""" & rstChilds.fields("Table").value & """>" & vbCrLf)
					Response.Write("<INPUT type='hidden' id=txtReportDefnChildFilterID_" & iCount & " name=txtReportDefnChildFilterID_" & iCount & " value=""" & rstChilds.fields("FilterID").value & """>" & vbCrLf)
					
					Dim sTemp As String
					If IsDBNull(rstChilds.fields("Filter").value) Then
						sTemp = ""
					Else
						sTemp = Replace(rstChilds.fields("Filter").value, """", "&quot;")
					End If
					Response.Write("<INPUT type='hidden' id=txtReportDefnChildFilter_" & iCount & " name=txtReportDefnChildFilter_" & iCount & " value=""" & sTemp & """>" & vbCrLf)
					
					Response.Write("<INPUT type='hidden' id=txtReportDefnChildOrderID_" & iCount & " name=txtReportDefnChildOrderID_" & iCount & " value=""" & rstChilds.fields("OrderID").value & """>" & vbCrLf)
					Response.Write("<INPUT type='hidden' id=txtReportDefnChildOrder_" & iCount & " name=txtReportDefnChildOrder_" & iCount & " value=""" & rstChilds.fields("Order").value & """>" & vbCrLf)
					Response.Write("<INPUT type='hidden' id=txtReportDefnChildRecords_" & iCount & " name=txtReportDefnChildRecords_" & iCount & " value=""" & rstChilds.fields("Records").value & """>" & vbCrLf)
					Response.Write("<INPUT type='hidden' id=txtReportDefnChildGridString_" & iCount & " name=txtReportDefnChildGridString_" & iCount & " value=""" & Replace(rstChilds.fields("gridstring").value, """", "&quot;") & vbTab & rstChilds.fields("FilterHidden").value & """>" & vbCrLf)
					Response.Write("<INPUT type='hidden' id=txtReportDefnChildFilterHidden_" & iCount & " name=txtReportDefnChildFilterHidden_" & iCount & " value=""" & rstChilds.fields("FilterHidden").value & """>" & vbCrLf)

					' Check if the child table filter is a hidden calc.
					If rstChilds.fields("FilterHidden").value = "Y" Then
						iHiddenChildFilterCount = iHiddenChildFilterCount + 1
					End If

					If rstChilds.fields("OrderDeleted").value = "Y" Then
						If Len(sChildInfo) > 0 Then
							sChildInfo = sChildInfo & vbCrLf
						End If
						sChildInfo = sChildInfo & "The '" & rstChilds.fields("Table").value & "' table order will be removed from this definition as it has been deleted by another user."
					End If

					If rstChilds.fields("FilterDeleted").value = "Y" Then
						If Len(sChildInfo) > 0 Then
							sChildInfo = sChildInfo & vbCrLf
						End If
						sChildInfo = sChildInfo & "The '" & rstChilds.fields("Table").value & "' table filter will be removed from this definition as it has been deleted by another user."
					End If

					If rstChilds.fields("FilterHiddenByOther").value = "Y" Then
						If Len(sChildInfo) > 0 Then
							sChildInfo = sChildInfo & vbCrLf
						End If
						sChildInfo = sChildInfo & "The '" & rstChilds.fields("Table").value & "' table filter will be removed from this definition as it has been made hidden by another user."
					End If

					rstChilds.MoveNext()
				Loop
				' Release the ADO recordset object.
				rstChilds.close()
			End If
			rstChilds = Nothing
		End If
		cmdReportChilds = Nothing

		session("childcount") = iCount
		session("hiddenfiltercount") = iHiddenChildFilterCount
		
'********************************************************************************
		
		Dim prmSummary = cmdDefn.CreateParameter("summary", 11, 2) '11=bit, 2=output
		cmdDefn.Parameters.Append(prmSummary)

		Dim prmPrintFilterHeader = cmdDefn.CreateParameter("printFilterHeader", 11, 2)	'11=bit, 2=output
		cmdDefn.Parameters.Append(prmPrintFilterHeader)

		'-----------------------------------------
		Dim prmOutputPreview = cmdDefn.CreateParameter("outputPreview", 11, 2) '11=bit, 2=output
		cmdDefn.Parameters.Append(prmOutputPreview)
		
		Dim prmOutputFormat = cmdDefn.CreateParameter("outputFormat", 3, 2) '3=integer, 2=output
		cmdDefn.Parameters.Append(prmOutputFormat)
		
		Dim prmOutputScreen = cmdDefn.CreateParameter("outputScreen", 11, 2)	'11=bit, 2=output
		cmdDefn.Parameters.Append(prmOutputScreen)
		
		Dim prmOutputPrinter = cmdDefn.CreateParameter("outputPrinter", 11, 2) '11=bit, 2=output
		cmdDefn.Parameters.Append(prmOutputPrinter)
		
		Dim prmOutputPrinterName = cmdDefn.CreateParameter("outputPrinterName", 200, 2, 8000) '200=varchar, 2=output, 8000=size
		cmdDefn.Parameters.Append(prmOutputPrinterName)
		
		Dim prmOutputSave = cmdDefn.CreateParameter("outputSave", 11, 2) '11=bit, 2=output
		cmdDefn.Parameters.Append(prmOutputSave)
		
		Dim prmOutputSaveExisting = cmdDefn.CreateParameter("outputSaveExisting", 3, 2) '3=integer, 2=output
		cmdDefn.Parameters.Append(prmOutputSaveExisting)
		
		Dim prmOutputEmail = cmdDefn.CreateParameter("outputEmail", 11, 2)	'11=bit, 2=output
		cmdDefn.Parameters.Append(prmOutputEmail)
		
		Dim prmOutputEmailAddr = cmdDefn.CreateParameter("outputEmailAddr", 3, 2) '3=integer, 2=output
		cmdDefn.Parameters.Append(prmOutputEmailAddr)
		
		Dim prmOutputEmailAddrName = cmdDefn.CreateParameter("outputEmailAddrName", 200, 2, 8000)	'200=varchar, 2=output, 8000=size
		cmdDefn.Parameters.Append(prmOutputEmailAddrName)
		
		Dim prmOutputEmailSubject = cmdDefn.CreateParameter("outputEmailSubject", 200, 2, 8000) '200=varchar, 2=output, 8000=size
		cmdDefn.Parameters.Append(prmOutputEmailSubject)

		Dim prmOutputEmailAttachAs = cmdDefn.CreateParameter("outputEmailAttachAs", 200, 2, 8000)	'200=varchar, 2=output, 8000=size
		cmdDefn.Parameters.Append(prmOutputEmailAttachAs)

		Dim prmOutputFilename = cmdDefn.CreateParameter("outputFilename", 200, 2, 8000) '200=varchar, 2=output, 8000=size
		cmdDefn.Parameters.Append(prmOutputFilename)
		'-----------------------------------------
		
		Dim prmTimestamp = cmdDefn.CreateParameter("timestamp", 3, 2) ' 3=integer, 2=output
		cmdDefn.Parameters.Append(prmTimestamp)

		Dim prmParent1AllRecords = cmdDefn.CreateParameter("parent1AllRecords", 11, 2) '11=bit, 2=output
		cmdDefn.Parameters.Append(prmParent1AllRecords)

		Dim prmParent1PicklistID = cmdDefn.CreateParameter("parent1PicklistID", 3, 2)	'3=integer, 2=output
		cmdDefn.Parameters.Append(prmParent1PicklistID)

		Dim prmParent1PicklistName = cmdDefn.CreateParameter("parent1PicklistName", 200, 2, 8000)	'200=varchar, 2=output, 8000=size
		cmdDefn.Parameters.Append(prmParent1PicklistName)

		Dim prmParent1PicklistHidden = cmdDefn.CreateParameter("parent1PicklistHidden", 11, 2)	'11=bit, 2=output
		cmdDefn.Parameters.Append(prmParent1PicklistHidden)

		Dim prmParent2AllRecords = cmdDefn.CreateParameter("parent2AllRecords", 11, 2) '11=bit, 2=output
		cmdDefn.Parameters.Append(prmParent2AllRecords)

		Dim prmParent2PicklistID = cmdDefn.CreateParameter("parent2PicklistID", 3, 2)	'3=integer, 2=output
		cmdDefn.Parameters.Append(prmParent2PicklistID)

		Dim prmParent2PicklistName = cmdDefn.CreateParameter("parent2PicklistName", 200, 2, 8000)	'200=varchar, 2=output, 8000=size
		cmdDefn.Parameters.Append(prmParent2PicklistName)

		Dim prmParent2PicklistHidden = cmdDefn.CreateParameter("parent2PicklistHidden", 11, 2)	'11=bit, 2=output
		cmdDefn.Parameters.Append(prmParent2PicklistHidden)

		Dim prmInfo = cmdDefn.CreateParameter("info", 200, 2, 8000)	'200=varchar, 2=output, 8000=size
		cmdDefn.Parameters.Append(prmInfo)

		Dim prmIgnoreZeros = cmdDefn.CreateParameter("ignoreZeros", 11, 2) '11=bit, 2=output
		cmdDefn.Parameters.Append(prmIgnoreZeros)
		
		Err.Clear()
		Dim rstDefinition = cmdDefn.Execute
		
		Dim iHiddenCalcCount = 0
		If (Err.Number <> 0) Then
			sErrMsg = "'" & Session("utilname") & "' report definition could not be read." & vbCrLf & formatError(Err.Description)
		Else
			If rstDefinition.state <> 0 Then
				' Read recordset values.
				iCount = 0
				Do While Not rstDefinition.EOF
					iCount = iCount + 1
					If rstDefinition.fields("definitionType").value = "ORDER" Then
						Response.Write("<INPUT type='hidden' id=txtReportDefnOrder_" & iCount & " name=txtReportDefnOrder_" & iCount & " value=""" & rstDefinition.fields("definitionString").value & """>" & vbCrLf)
					ElseIf rstDefinition.fields("definitionType").value = "REPETITION" Then
						Response.Write("<INPUT type='hidden' id=txtReportDefnRepetition_" & iCount & " name=txtReportDefnRepetition_" & iCount & " value=""" & rstDefinition.fields("definitionString").value & """>" & vbCrLf)
					Else
						Response.Write("<INPUT type='hidden' id=txtReportDefnColumn_" & iCount & " name=txtReportDefnColumn_" & iCount & " value=""" & Replace(rstDefinition.fields("definitionString").value, """", "&quot;") & """>" & vbCrLf)
	
						' Check if the report column is a hidden calc.
						If rstDefinition.fields("hidden").value = "Y" Then
							iHiddenCalcCount = iHiddenCalcCount + 1
						End If
					End If
					rstDefinition.MoveNext()
				Loop

				' Release the ADO recordset object.
				rstDefinition.close()
			End If
			rstDefinition = Nothing
			
			' NB. IMPORTANT ADO NOTE.
			' When calling a stored procedure which returns a recordset AND has output parameters
			' you need to close the recordset and set it to nothing before using the output parameters. 
			If Len(cmdDefn.Parameters("errMsg").value) > 0 Then
				sErrMsg = "'" & Session("utilname") & "' " & cmdDefn.Parameters("errMsg").value
			End If

			Response.Write("<INPUT type='hidden' id=txtDefn_Name name=txtDefn_Name value=""" & Replace(cmdDefn.Parameters("name").value, """", "&quot;") & """>" & vbCrLf)
			Response.Write("<INPUT type='hidden' id=txtDefn_Owner name=txtDefn_Owner value=""" & Replace(cmdDefn.Parameters("owner").value, """", "&quot;") & """>" & vbCrLf)
			Response.Write("<INPUT type='hidden' id=txtDefn_Description name=txtDefn_Description value=""" & Replace(cmdDefn.Parameters("description").value, """", "&quot;") & """>" & vbCrLf)
			Response.Write("<INPUT type='hidden' id=txtDefn_BaseTableID name=txtDefn_BaseTableID value=" & cmdDefn.Parameters("baseTableID").value & ">" & vbCrLf)
			Response.Write("<INPUT type='hidden' id=txtDefn_AllRecords name=txtDefn_AllRecords value=" & cmdDefn.Parameters("allRecords").value & ">" & vbCrLf)
			Response.Write("<INPUT type='hidden' id=txtDefn_PicklistID name=txtDefn_PicklistID value=" & cmdDefn.Parameters("picklistID").value & ">" & vbCrLf)
			Response.Write("<INPUT type='hidden' id=txtDefn_PicklistName name=txtDefn_PicklistName value=""" & Replace(cmdDefn.Parameters("picklistName").value, """", "&quot;") & """>" & vbCrLf)
			Response.Write("<INPUT type='hidden' id=txtDefn_PicklistHidden name=txtDefn_PicklistHidden value=" & cmdDefn.Parameters("picklistHidden").value & ">" & vbCrLf)
			Response.Write("<INPUT type='hidden' id=txtDefn_FilterID name=txtDefn_FilterID value=" & cmdDefn.Parameters("filterID").value & ">" & vbCrLf)
			Response.Write("<INPUT type='hidden' id=txtDefn_FilterName name=txtDefn_FilterName value=""" & Replace(cmdDefn.Parameters("filterName").value, """", "&quot;") & """>" & vbCrLf)
			Response.Write("<INPUT type='hidden' id=txtDefn_FilterHidden name=txtDefn_FilterHidden value=" & cmdDefn.Parameters("filterHidden").value & ">" & vbCrLf)
			Response.Write("<INPUT type='hidden' id=txtDefn_Parent1TableID name=txtDefn_Parent1TableID value=" & cmdDefn.Parameters("parent1TableID").value & ">" & vbCrLf)
			Response.Write("<INPUT type='hidden' id=txtDefn_Parent1TableName name=txtDefn_Parent1TableName value=""" & cmdDefn.Parameters("parent1TableName").value & """>" & vbCrLf)
			Response.Write("<INPUT type='hidden' id=txtDefn_Parent1FilterID name=txtDefn_Parent1FilterID value=" & cmdDefn.Parameters("parent1FilterID").value & ">" & vbCrLf)
			Response.Write("<INPUT type='hidden' id=txtDefn_Parent1FilterName name=txtDefn_Parent1FilterName value=""" & Replace(cmdDefn.Parameters("parent1FilterName").value, """", "&quot;") & """>" & vbCrLf)
			Response.Write("<INPUT type='hidden' id=txtDefn_Parent1FilterHidden name=txtDefn_Parent1FilterHidden value=" & cmdDefn.Parameters("parent1FilterHidden").value & ">" & vbCrLf)
			Response.Write("<INPUT type='hidden' id=txtDefn_Parent2TableID name=txtDefn_Parent2TableID value=" & cmdDefn.Parameters("parent2TableID").value & ">" & vbCrLf)
			Response.Write("<INPUT type='hidden' id=txtDefn_Parent2TableName name=txtDefn_Parent2TableName value=""" & cmdDefn.Parameters("parent2TableName").value & """>" & vbCrLf)
			Response.Write("<INPUT type='hidden' id=txtDefn_Parent2FilterID name=txtDefn_Parent2FilterID value=" & cmdDefn.Parameters("parent2FilterID").value & ">" & vbCrLf)
			Response.Write("<INPUT type='hidden' id=txtDefn_Parent2FilterName name=txtDefn_Parent2FilterName value=""" & Replace(cmdDefn.Parameters("parent2FilterName").value, """", "&quot;") & """>" & vbCrLf)
			Response.Write("<INPUT type='hidden' id=txtDefn_Parent2FilterHidden name=txtDefn_Parent2FilterHidden value=" & cmdDefn.Parameters("parent2FilterHidden").value & ">" & vbCrLf)

			Response.Write("<INPUT type='hidden' id=txtDefn_Summary name=txtDefn_Summary value=" & cmdDefn.Parameters("summary").value & ">" & vbCrLf)
			Response.Write("<INPUT type='hidden' id=txtDefn_PrintFilterHeader name=txtDefn_PrintFilterHeader value=" & cmdDefn.Parameters("printFilterHeader").value & ">" & vbCrLf)

			Response.Write("<INPUT type='hidden' id=txtDefn_OutputPreview name=txtDefn_OutputPreview value=" & cmdDefn.Parameters("OutputPreview").value & ">" & vbCrLf)
			Response.Write("<INPUT type='hidden' id=txtDefn_OutputFormat name=txtDefn_OutputFormat value=" & cmdDefn.Parameters("OutputFormat").value & ">" & vbCrLf)
			Response.Write("<INPUT type='hidden' id=txtDefn_OutputScreen name=txtDefn_OutputScreen value=" & cmdDefn.Parameters("OutputScreen").value & ">" & vbCrLf)
			Response.Write("<INPUT type='hidden' id=txtDefn_OutputPrinter name=txtDefn_OutputPrinter value=" & cmdDefn.Parameters("OutputPrinter").value & ">" & vbCrLf)
			Response.Write("<INPUT type='hidden' id=txtDefn_OutputPrinterName name=txtDefn_OutputPrinterName value=""" & cmdDefn.Parameters("OutputPrinterName").value & """>" & vbCrLf)
			Response.Write("<INPUT type='hidden' id=txtDefn_OutputSave name=txtDefn_OutputSave value=" & cmdDefn.Parameters("OutputSave").value & ">" & vbCrLf)
			Response.Write("<INPUT type='hidden' id=txtDefn_OutputSaveExisting name=txtDefn_OutputSaveExisting value=" & cmdDefn.Parameters("OutputSaveExisting").value & ">" & vbCrLf)
			Response.Write("<INPUT type='hidden' id=txtDefn_OutputEmail name=txtDefn_OutputEmail value=" & cmdDefn.Parameters("OutputEmail").value & ">" & vbCrLf)
			Response.Write("<INPUT type='hidden' id=txtDefn_OutputEmailAddr name=txtDefn_OutputEmailAddr value=" & cmdDefn.Parameters("OutputEmailAddr").value & ">" & vbCrLf)
			Response.Write("<INPUT type='hidden' id=txtDefn_OutputEmailAddrName name=txtDefn_OutputEmailAddrName value=""" & Replace(cmdDefn.Parameters("OutputEmailAddrName").value, """", "&quot;") & """>" & vbCrLf)
			Response.Write("<INPUT type='hidden' id=txtDefn_OutputEmailSubject name=txtDefn_OutputEmailSubject value=""" & Replace(cmdDefn.Parameters("OutputEmailSubject").value, """", "&quot;") & """>" & vbCrLf)
			Response.Write("<INPUT type='hidden' id=txtDefn_OutputEmailAttachAs name=txtDefn_OutputEmailAttachAs value=""" & Replace(cmdDefn.Parameters("OutputEmailAttachAs").value, """", "&quot;") & """>" & vbCrLf)
			Response.Write("<INPUT type='hidden' id=txtDefn_OutputFilename name=txtDefn_OutputFilename value=""" & cmdDefn.Parameters("OutputFilename").value & """>" & vbCrLf)

			Response.Write("<INPUT type='hidden' id=txtDefn_Timestamp name=txtDefn_Timestamp value=" & cmdDefn.Parameters("timestamp").value & ">" & vbCrLf)
			Response.Write("<INPUT type='hidden' id=txtDefn_HiddenCalcCount name=txtDefn_HiddenCalcCount value=" & iHiddenCalcCount & ">" & vbCrLf)
			Response.Write("<INPUT type='hidden' id=txtDefn_Parent1AllRecords name=txtDefn_Parent1AllRecords value=" & cmdDefn.Parameters("parent1AllRecords").value & ">" & vbCrLf)
			Response.Write("<INPUT type='hidden' id=txtDefn_Parent1PicklistID name=txtDefn_Parent1PicklistID value=" & cmdDefn.Parameters("parent1PicklistID").value & ">" & vbCrLf)
			Response.Write("<INPUT type='hidden' id=txtDefn_Parent1PicklistName name=txtDefn_Parent1PicklistName value=""" & Replace(cmdDefn.Parameters("parent1PicklistName").value, """", "&quot;") & """>" & vbCrLf)
			Response.Write("<INPUT type='hidden' id=txtDefn_Parent1PicklistHidden name=txtDefn_Parent1PicklistHidden value=" & cmdDefn.Parameters("parent1PicklistHidden").value & ">" & vbCrLf)
			Response.Write("<INPUT type='hidden' id=txtDefn_Parent2AllRecords name=txtDefn_Parent2AllRecords value=" & cmdDefn.Parameters("parent2AllRecords").value & ">" & vbCrLf)
			Response.Write("<INPUT type='hidden' id=txtDefn_Parent2PicklistID name=txtDefn_Parent2PicklistID value=" & cmdDefn.Parameters("parent2PicklistID").value & ">" & vbCrLf)
			Response.Write("<INPUT type='hidden' id=txtDefn_Parent2PicklistName name=txtDefn_Parent2PicklistName value=""" & Replace(cmdDefn.Parameters("parent2PicklistName").value, """", "&quot;") & """>" & vbCrLf)
			Response.Write("<INPUT type='hidden' id=txtDefn_Parent2PicklistHidden name=txtDefn_Parent2PicklistHidden value=" & cmdDefn.Parameters("parent2PicklistHidden").value & ">" & vbCrLf)
			Response.Write("<INPUT type='hidden' id=txtDefn_IgnoreZeros name=txtDefn_IgnoreZeros value=" & cmdDefn.Parameters("ignoreZeros").value & ">" & vbCrLf)
			
			Dim sInfo = cmdDefn.Parameters("info").value
			If Len(sChildInfo) > 0 Then
				If Len(sInfo) > 0 Then
					sInfo = sInfo & vbCrLf
				End If
				
				sInfo = sInfo & sChildInfo
			End If
			Response.Write("<INPUT type='hidden' id=txtDefn_Info name=txtDefn_Info value=""" & Replace(sInfo, """", "&quot;") & """>" & vbCrLf)
		End If

		' Release the ADO command object.
		cmdDefn = Nothing

		if len(sErrMsg) > 0 then
			session("confirmtext") = sErrMsg
			session("confirmtitle") = "OpenHR Intranet"
			Session("followpage") = "defsel"
			Session("reaction") = "CUSTOMREPORTS"
			Response.Clear
			Response.Redirect("confirmok")
		end if

	else
		session("childcount") = 0
		session("hiddenfiltercount") =  0
	end if
%>
</form>

<form id=frmAccess>
<%
	sErrorDescription = ""
	
	' Get the table records.
	Dim cmdAccess = CreateObject("ADODB.Command")
	cmdAccess.CommandText = "spASRIntGetUtilityAccessRecords"
	cmdAccess.CommandType = 4 ' Stored Procedure
	cmdAccess.ActiveConnection = Session("databaseConnection")

	Dim prmUtilType = cmdAccess.CreateParameter("utilType", 3, 1) ' 3=integer, 1=input
	cmdAccess.Parameters.Append(prmUtilType)
	prmUtilType.value = 2 ' 2 = custom report

	Dim prmUtilID3 = cmdAccess.CreateParameter("utilID", 3, 1) ' 3=integer, 1=input
	cmdAccess.Parameters.Append(prmUtilID3)
	if ucase(session("action")) = "NEW"	then
		prmUtilID3.value = 0
	else
		prmUtilID3.value = cleanNumeric(Session("utilid"))
	end if

	Dim prmFromCopy = cmdAccess.CreateParameter("fromCopy", 3, 1) ' 3=integer, 1=input
	cmdAccess.Parameters.Append(prmFromCopy)
	if ucase(session("action")) = "COPY" then
		prmFromCopy.value = 1 
	else
		prmFromCopy.value = 0
	end if

	Err.Clear()
	Dim rstAccessInfo = cmdAccess.Execute
	If (Err.Number <> 0) Then
		sErrorDescription = "The access information could not be retrieved." & vbCrLf & formatError(Err.Description)
	End If

	if len(sErrorDescription) = 0 then
		Dim iCount = 0
		do while not rstAccessInfo.EOF
			Response.Write("<INPUT type='hidden' id=txtAccess_" & iCount & " name=txtAccess_" & iCount & " value=""" & rstAccessInfo.fields("accessDefinition").value & """>" & vbCrLf)

			iCount = iCount + 1
			rstAccessInfo.MoveNext
		loop

		' Release the ADO recordset object.
		rstAccessInfo.Close()
		rstAccessInfo = Nothing
	end if
	
	' Release the ADO command object.
	cmdAccess = Nothing
%>
</form>

<FORM id=frmUseful name=frmUseful style="visibility:hidden;display:none">
	<INPUT type="hidden" id=txtUserName name=txtUserName value="<%=session("username")%>">
	<INPUT type="hidden" id=txtLoading name=txtLoading value="Y">
	<INPUT type="hidden" id=txtCurrentBaseTableID name=txtCurrentBaseTableID>
	<INPUT type="hidden" id=txtCurrentChildTableID name=txtCurrentChildTableID value=0>
	<INPUT type="hidden" id=txtTablesChanged name=txtTablesChanged>
	<INPUT type="hidden" id=txtSelectedColumnsLoaded name=txtSelectedColumnsLoaded value=0>
	<INPUT type="hidden" id=txtSortLoaded name=txtSortLoaded value=0>
	<INPUT type="hidden" id=txtRepetitionLoaded name=txtRepetitionLoaded value=0>
	<INPUT type="hidden" id=txtChildsLoaded name=txtChildsLoaded value=0>
	<INPUT type="hidden" id=txtChanged name=txtChanged value=0>
	<INPUT type="hidden" id=txtUtilID name=txtUtilID value=<%=session("utilid")%>>
	<INPUT type="hidden" id=txtChildCount name=txtChildCount value=<%=session("childcount")%>>
	<INPUT type="hidden" id=txtHiddenChildFilterCount name=txtHiddenChildFilterCount value=<%=session("hiddenfiltercount")%>>
	<INPUT type="hidden" id=txtLockGridEvents name=txtLockGridEvents value=0>
	<INPUT type="hidden" id=txtChildColumnSelected name=txtChildColumnSelected value=0>
	<INPUT type="hidden" id=txtGridActionCancelled name=txtGridActionCancelled value=0>
	<INPUT type="hidden" id=txtGridChangeRecursive name=txtGridChangeRecursive value=0>

<%
	Dim cmdDefinition = CreateObject("ADODB.Command")
	cmdDefinition.CommandText = "sp_ASRIntGetModuleParameter"
	cmdDefinition.CommandType = 4 ' Stored procedure.
	cmdDefinition.ActiveConnection = Session("databaseConnection")

	Dim prmModuleKey = cmdDefinition.CreateParameter("moduleKey", 200, 1, 8000) ' 200=varchar, 1=input, 8000=size
	cmdDefinition.Parameters.Append(prmModuleKey)
	prmModuleKey.value = "MODULE_PERSONNEL"

	Dim prmParameterKey = cmdDefinition.CreateParameter("paramKey", 200, 1, 8000)	' 200=varchar, 1=input, 8000=size
	cmdDefinition.Parameters.Append(prmParameterKey)
	prmParameterKey.value = "Param_TablePersonnel"

	Dim prmParameterValue = cmdDefinition.CreateParameter("paramValue", 200, 2, 8000) '200=varchar, 2=output, 8000=size
	cmdDefinition.Parameters.Append(prmParameterValue)

	Err.Clear()
	cmdDefinition.Execute

	Response.Write("<INPUT type='hidden' id=txtPersonnelTableID name=txtPersonnelTableID value=" & cmdDefinition.Parameters("paramValue").value & ">" & vbCrLf)
	
	cmdDefinition = Nothing

	Response.Write("<INPUT type='hidden' id=txtErrorDescription name=txtErrorDescription value=""" & sErrorDescription & """>" & vbCrLf)
	Response.Write("<INPUT type='hidden' id=txtAction name=txtAction value=" & Session("action") & ">" & vbCrLf)
%>
</FORM>

<FORM id=frmValidate name=frmValidate target=validate method=post action=util_validate_customreports style="visibility:hidden;display:none">
	<INPUT type=hidden id="validateBaseFilter" name=validateBaseFilter value=0>
	<INPUT type=hidden id="validateBasePicklist" name=validateBasePicklist value=0>
	<INPUT type=hidden id="validateEmailGroup" name=validateEmailGroup value=0>
	<INPUT type=hidden id="validateP1Filter" name=validateP1Filter value=0>
	<INPUT type=hidden id="validateP1Picklist" name=validateP1Picklist value=0>
	<INPUT type=hidden id="validateP2Filter" name=validateP2Filter value=0>
	<INPUT type=hidden id="validateP2Picklist" name=validateP2Picklist value=0>
<!-- need the array of child filters -->
	<INPUT type=hidden id="validateChildFilter" name=validateChildFilter value=0>
	<INPUT type=hidden id="validateChildOrders" name=validateChildOrders value=0>
	
	<INPUT type=hidden id="validateCalcs" name=validateCalcs value = ''>
	<INPUT type=hidden id="validateHiddenGroups" name=validateHiddenGroups value = ''>
	<INPUT type=hidden id="validateName" name=validateName value=''>
	<INPUT type=hidden id="validateTimestamp" name=validateTimestamp value=''>
	<INPUT type=hidden id="validateUtilID" name=validateUtilID value=''>
</FORM>

<FORM id=frmSend name=frmSend method=post action=util_def_customreports_Submit style="visibility:hidden;display:none">
	<INPUT type="hidden" id=txtSend_ID name=txtSend_ID>	
	<INPUT type="hidden" id=txtSend_name name=txtSend_name>
	<INPUT type="hidden" id=txtSend_description name=txtSend_description>
	<INPUT type="hidden" id=txtSend_baseTable name=txtSend_baseTable>
	<INPUT type="hidden" id=txtSend_allRecords name=txtSend_allRecords>
	<INPUT type="hidden" id=txtSend_picklist name=txtSend_picklist>
	<INPUT type="hidden" id=txtSend_filter name=txtSend_filter>
	<INPUT type="hidden" id=txtSend_parent1Table name=txtSend_parent1Table>
	<INPUT type="hidden" id=txtSend_parent1AllRecords name=txtSend_parent1AllRecords>
	<INPUT type="hidden" id=txtSend_parent1Filter name=txtSend_parent1Filter>
	<INPUT type="hidden" id=txtSend_parent1Picklist name=txtSend_parent1Picklist>
	<INPUT type="hidden" id=txtSend_parent2Table name=txtSend_parent2Table>
	<INPUT type="hidden" id=txtSend_parent2AllRecords name=txtSend_parent2AllRecords>
	<INPUT type="hidden" id=txtSend_parent2Filter name=txtSend_parent2Filter>
	<INPUT type="hidden" id=txtSend_parent2Picklist name=txtSend_parent2Picklist>

<!-- need the array of child info to send -->
	<INPUT type="hidden" id=txtSend_childTable name=txtSend_childTable>
	<INPUT type="hidden" id=txtSend_summary name=txtSend_summary>
	<INPUT type="hidden" id=txtSend_IgnoreZeros name=txtSend_IgnoreZeros>
	<INPUT type="hidden" id=txtSend_printFilterHeader name=txtSend_printFilterHeader>
	<INPUT type="hidden" id=txtSend_access name=txtSend_access>
	<INPUT type="hidden" id=txtSend_userName name=txtSend_userName>

	<INPUT type="hidden" id=txtSend_OutputPreview name=txtSend_OutputPreview>
	<INPUT type="hidden" id=txtSend_OutputFormat name=txtSend_OutputFormat>
	<INPUT type="hidden" id=txtSend_OutputScreen name=txtSend_OutputScreen>
	<INPUT type="hidden" id=txtSend_OutputPrinter name=txtSend_OutputPrinter>
	<INPUT type="hidden" id=txtSend_OutputPrinterName name=txtSend_OutputPrinterName>
	<INPUT type="hidden" id=txtSend_OutputSave name=txtSend_OutputSave>
	<INPUT type="hidden" id=txtSend_OutputSaveExisting name=txtSend_OutputSaveExisting>
	<INPUT type="hidden" id=txtSend_OutputEmail name=txtSend_OutputEmail>
	<INPUT type="hidden" id=txtSend_OutputEmailAddr name=txtSend_OutputEmailAddr>
	<INPUT type="hidden" id=txtSend_OutputEmailSubject name=txtSend_OutputEmailSubject>
	<INPUT type="hidden" id=txtSend_OutputEmailAttachAs name=txtSend_OutputEmailAttachAs>
	<INPUT type="hidden" id=txtSend_OutputFilename name=txtSend_OutputFilename>

	<INPUT type="hidden" id=txtSend_columns name=txtSend_columns>
	<INPUT type="hidden" id=txtSend_columns2 name=txtSend_columns2>

	<INPUT type="hidden" id=txtSend_reaction name=txtSend_reaction>

	<INPUT type="hidden" id=txtSend_jobsToHide name=txtSend_jobsToHide>
	<INPUT type="hidden" id=txtSend_jobsToHideGroups name=txtSend_jobsToHideGroups>
</FORM>

<FORM id=frmCustomReportChilds name=frmCustomReportChilds target="childselection" action="util_customreportchilds" method=post style="visibility:hidden;display:none">
	<INPUT type="hidden" id=childTableID name=childTableID>
	<INPUT type="hidden" id=childTable name=childTable>
	<INPUT type="hidden" id=childFilterID name=childFilterID>
	<INPUT type="hidden" id=childFilter name=childFilter>
	<INPUT type="hidden" id=childOrderID name=childOrderID>
	<INPUT type="hidden" id=childOrder name=childOrder>
	<INPUT type="hidden" id=childRecords name=childRecords>
	<INPUT type="hidden" id=childrenString name=childrenString>
	<INPUT type="hidden" id=childrenNames name=childrenNames>
	<INPUT type="hidden" id=selectedChildString name=selectedChildString>
	<INPUT type="hidden" id=childAction name=childAction value="NEW">
	<INPUT type="hidden" id=childMax name=childMax value=5>
</FORM>

<FORM id=frmRecordSelection name=frmRecordSelection target="recordSelection" action="util_recordSelection" method=post style="visibility:hidden;display:none">
	<INPUT type="hidden" id=recSelType name=recSelType>
	<INPUT type="hidden" id=recSelTableID name=recSelTableID>
	<INPUT type="hidden" id=recSelCurrentID name=recSelCurrentID>
	<INPUT type="hidden" id=recSelTable name=recSelTable>
	<INPUT type="hidden" id=recSelDefOwner name=recSelDefOwner>
	<INPUT type="hidden" id=recSelDefType name=recSelDefType>
</FORM>

<FORM id=frmEmailSelection name=frmEmailSelection target="emailSelection" action="util_emailSelection" method=post style="visibility:hidden;display:none">
	<INPUT type="hidden" id=EmailSelCurrentID name=EmailSelCurrentID>
</FORM>

<FORM id=frmSortOrder name=frmSortOrder action="util_sortorderselection" target="sortorderselection" method=post style="visibility:hidden;display:none">
	<INPUT type=hidden id=txtSortInclude name=txtSortInclude>
	<INPUT type=hidden id=txtSortExclude name=txtSortExclude>
	<INPUT type=hidden id=txtSortEditing name=txtSortEditing>
	<INPUT type=hidden id=txtSortColumnID name=txtSortColumnID>
	<INPUT type=hidden id=txtSortColumnName name=txtSortColumnName>
	<INPUT type=hidden id=txtSortOrder name=txtSortOrder>	
	<INPUT type=hidden id=txtSortBOC name=txtSortBOC>
	<INPUT type=hidden id=txtSortPOC name=txtSortPOC>
	<INPUT type=hidden id=txtSortVOC name=txtSortVOC>
	<INPUT type=hidden id=txtSortSRV name=txtSortSRV>
</FORM>

<FORM id=frmSelectionAccess name=frmSelectionAccess style="visibility:hidden;display:none">
	<INPUT type="hidden" id=forcedHidden name=forcedHidden value="N">
	<INPUT type="hidden" id=baseHidden name=baseHidden value="N">
	<INPUT type="hidden" id=p1Hidden name=p1Hidden value="N">
	<INPUT type="hidden" id=p2Hidden name=p2Hidden value="N">

<!-- need the count of hidden child filter access info -->
	<INPUT type="hidden" id=childHidden name=childHidden value=0>
	<INPUT type="hidden" id=calcsHiddenCount name=calcsHiddenCount value=0>
</FORM>

<INPUT type='hidden' id=txtTicker name=txtTicker value=0>
<INPUT type='hidden' id=txtLastKeyFind name=txtLastKeyFind value="">

<%Html.RenderPartial("Util_Def_CustomReports/grdColProps")%>

</div>

<script type="text/javascript">

	function updateCurrentColProp(psProp, pbValue) {
		with (grdColProps) {
			Columns(psProp).Value = pbValue;
		}
		return;
	}

	function getCurrentColProp(psProp) {
		with (grdColProps) {
			if (Columns(psProp).Value == "-1") {
				return true;
			}
			else {
				return false;
			}
		}
	}

</script>


<script type="text/javascript">
	util_def_customreports_onload();
	util_def_customreports_addhandlers();
</script>



