<%@ Control Language="VB" Inherits="System.Web.Mvc.ViewUserControl" %>
<%@ Import Namespace="DMI.NET" %>
<%@ Import Namespace="HR.Intranet.Server" %>

<object classid="clsid:F9043C85-F6F2-101A-A3C9-08002B2F49FB"
    id="dialog"
    codebase="cabs/comdlg32.cab#Version=1,0,0,0"
    style="LEFT: 0px; TOP: 0px"
    viewastext>
    <param name="_ExtentX" value="847">
    <param name="_ExtentY" value="847">
    <param name="_Version" value="393216">
    <param name="CancelError" value="0">
    <param name="Color" value="0">
    <param name="Copies" value="1">
    <param name="DefaultExt" value="">
    <param name="DialogTitle" value="">
    <param name="FileName" value="">
    <param name="Filter" value="">
    <param name="FilterIndex" value="0">
    <param name="Flags" value="0">
    <param name="FontBold" value="0">
    <param name="FontItalic" value="0">
    <param name="FontName" value="">
    <param name="FontSize" value="8">
    <param name="FontStrikeThru" value="0">
    <param name="FontUnderLine" value="0">
    <param name="FromPage" value="0">
    <param name="HelpCommand" value="0">
    <param name="HelpContext" value="0">
    <param name="HelpFile" value="">
    <param name="HelpKey" value="">
    <param name="InitDir" value="">
    <param name="Max" value="0">
    <param name="Min" value="0">
    <param name="MaxFileSize" value="260">
    <param name="PrinterDefault" value="1">
    <param name="ToPage" value="0">
    <param name="Orientation" value="1">
</object>

<%
    Dim objCrossTab As HR.Intranet.Server.CrossTab
    
    objCrossTab = CType(Session("objCrossTab" & Session("UtilID")), CrossTab)
    If (objCrossTab.ErrorString = "") Then

        Response.Write("<script type=""text/javascript"">" & vbCrLf)
        Response.Write("   function util_run_crosstabs_window_onload() {" & vbCrLf)
%>
    setGridFont(ssOutputGrid);
    setGridFont(ssHiddenGrid);
<%
    Response.Write("	crosstab_loadAddRecords();" & vbCrLf)
    Response.Write("    frmError.txtEventLogID.value = """ & CleanStringForJavaScript(objCrossTab.EventLogID) & """;" & vbCrLf)
    Response.Write("  }" & vbCrLf)
    Response.Write("</script>" & vbCrLf)


    Response.Write("<script type=""text/javascript"">" & vbCrLf)
	Response.Write("  function ssOutputGrid_DblClick() {" & vbCrLf)
    
    If objCrossTab.RecordDescExprID = 0 Then
        Response.Write("    OpenHR.messageBox(""Unable to show cell breakdown details as no record description has been set up for the '" & CleanStringForJavaScript(objCrossTab.BaseTableName) & "' table."",64,""Cross Tab Breakdown"");" & vbCrLf)
    Else
        Response.Write("	if (ssOutputGrid.Col > 0) {" & vbCrLf)
		Response.Write("      frmData = OpenHR.getFrame(""reportdataframe"");" & vbCrLf)
        Response.Write("      lngPage = 0;" & vbCrLf)
        Response.Write("      if (cboPage.selectedIndex != -1) {" & vbCrLf)
        Response.Write("        lngPage = cboPage.options[cboPage.selectedIndex].Value;" & vbCrLf)
        Response.Write("      }" & vbCrLf)       
        Response.Write("      getBreakdown(ssOutputGrid.Col - 1, ssOutputGrid.AddItemRowIndex(ssOutputGrid.Bookmark), lngPage, cboIntersectionType.options[cboIntersectionType.selectedIndex].Value, ssOutputGrid.ActiveCell.Value);" & vbCrLf)
        Response.Write("    }" & vbCrLf)
        Response.Write("  }" & vbCrLf)
    End If

    Response.Write("</script>" & vbCrLf)

    objCrossTab.EventLogChangeHeaderStatus(3)  'Successful

	else 
%>

<%
    Response.Write("<FORM Name=frmPopup ID=frmPopup>" & vbCrLf)
    Response.Write("<table align=center class=""outline="" cellPadding=5 cellSpacing=0>" & vbCrLf)
    Response.Write("	<TR>" & vbCrLf)
    Response.Write("		<TD>" & vbCrLf)
    Response.Write("			<table class=""invisible"" cellspacing=0 cellpadding=0>" & vbCrLf)
    Response.Write("			  <tr>" & vbCrLf)
    Response.Write("			    <td colspan=3 height=10></td>" & vbCrLf)
    Response.Write("			  </tr>" & vbCrLf)
    Response.Write("			  <tr> " & vbCrLf)
    Response.Write("			    <td width=20 height=10></td> " & vbCrLf)
    Response.Write("			    <td align=center> " & vbCrLf)

    If objCrossTab.NoRecords Then
        If objCrossTab.CrossTabType = 3 Then
            Response.Write("						<H4>Absence Breakdown Completed successfully.</H4>" & vbCrLf)
        Else
            Response.Write("						<H4>Cross Tab '" & Session("utilname") & "' Completed successfully.</H4>" & vbCrLf)
        End If
        objCrossTab.EventLogChangeHeaderStatus(3)    'Successful
    Else
        If objCrossTab.CrossTabType = 3 Then
            Response.Write("						<H4>Absence Breakdown Failed." & vbCrLf)
        Else
            Response.Write("						<H4>Cross Tab '" & Session("utilname") & "' Failed." & vbCrLf)
        End If
        objCrossTab.EventLogChangeHeaderStatus(2)    'Failed
    End If

    Response.Write("			    </td>" & vbCrLf)
    Response.Write("			    <td width=20></td> " & vbCrLf)
    Response.Write("			  </tr>" & vbCrLf)
    Response.Write("			  <tr> " & vbCrLf)
    Response.Write("			    <td width=20 height=10></td> " & vbCrLf)
    Response.Write("			    <td align=center nowrap>" & objCrossTab.ErrorString & vbCrLf)
    Response.Write("			    </td>" & vbCrLf)
    Response.Write("			    <td width=20></td> " & vbCrLf)
    Response.Write("			  </tr>" & vbCrLf)
    Response.Write("			  <tr>" & vbCrLf)
    Response.Write("			    <td colspan=3 height=10>&nbsp;</td>" & vbCrLf)
    Response.Write("			  </tr>" & vbCrLf)
    Response.Write("			  <tr> " & vbCrLf)
    Response.Write("			    <td colspan=3 height=10 align=center> " & vbCrLf)
    Response.Write("						<input type=button id=cmdClose name=cmdClose value=Close style=""WIDTH: 80px"" width=80px class=""btn""" & vbCrLf)
    Response.Write("                      onclick=""closeclick();""" & vbCrLf)
    Response.Write("                      onmouseover=""try{button_onMouseOver(this);}catch(e){}""" & vbCrLf)
    Response.Write("                      onmouseout=""try{button_onMouseOut(this);}catch(e){}""" & vbCrLf)
    Response.Write("                      onfocus=""try{button_onFocus(this);}catch(e){}""" & vbCrLf)
    Response.Write("                      onblur=""try{button_onBlur(this);}catch(e){}"" />" & vbCrLf)
    Response.Write("			    </td>" & vbCrLf)
    Response.Write("			  </tr>" & vbCrLf)
    Response.Write("			  <tr> " & vbCrLf)
    Response.Write("			    <td colspan=3 height=10></td>" & vbCrLf)
    Response.Write("			  </tr>" & vbCrLf)
    Response.Write("			</table>" & vbCrLf)
    Response.Write("		</td>" & vbCrLf)
    Response.Write("	</tr>" & vbCrLf)
    Response.Write("</table>" & vbCrLf)
    Response.Write("</FORM>" & vbCrLf)
		
		if objCrossTab.ErrorString <> "" then
			objCrossTab.FailedMessage = objCrossTab.ErrorString
		end if

		Response.End
	end if
%>

<table align=center class="outline" cellPadding=5 cellSpacing=0 width=100% height=100%>
	<TR>
		<TD>
			<table HEIGHT="100%" WIDTH="100%" class="invisible" CELLSPACING=0 CELLPADDING=0>
				<TR>
					<TD COLSPAN=50>
						<OBJECT classid=clsid:4A4AA697-3E6F-11D2-822F-00104B9E07A1
							codebase="cabs/COAInt_Grid.cab#version=3,1,3,6" 
							id=ssOutputGrid name=ssOutputGrid
							style="HEIGHT: 400px; LEFT: 0px; TOP: 0px; WIDTH: 100%">
							<PARAM NAME="ScrollBars" VALUE="4">
							<PARAM NAME="_Version" VALUE="196617">
							<PARAM NAME="DataMode" VALUE="2">
							<PARAM NAME="Cols" VALUE="0">
							<PARAM NAME="Rows" VALUE="0">
							<PARAM NAME="BorderStyle" VALUE="1">
							<PARAM NAME="RecordSelectors" VALUE="0">
							<PARAM NAME="GroupHeaders" VALUE="1">
							<PARAM NAME="ColumnHeaders" VALUE="1">
							<PARAM NAME="GroupHeadLines" VALUE="1">
							<PARAM NAME="HeadLines" VALUE="1">
							<PARAM NAME="FieldDelimiter" VALUE="(None)">
							<PARAM NAME="FieldSeparator" VALUE="(Tab)">
							<PARAM NAME="Row.Count" VALUE="0">
							<PARAM NAME="Col.Count" VALUE="1">
							<PARAM NAME="stylesets.count" VALUE="1">
							<PARAM NAME="TagVariant" VALUE="EMPTY">
							<PARAM NAME="UseGroups" VALUE="0">
							<PARAM NAME="HeadFont3D" VALUE="0">
							<PARAM NAME="Font3D" VALUE="0">
							<PARAM NAME="DividerType" VALUE="3">
							<PARAM NAME="DividerStyle" VALUE="1">
							<PARAM NAME="DefColWidth" VALUE="3528">
							<PARAM NAME="BeveColorScheme" VALUE="2">
							<PARAM NAME="BevelColorFrame" VALUE="-2147483642">
							<PARAM NAME="BevelColorHighlight" VALUE="-2147483643">
							<PARAM NAME="BevelColorShadow" VALUE="-2147483632">
							<PARAM NAME="BevelColorFace" VALUE="-2147483633">
							<PARAM NAME="CheckBox3D" VALUE="1">
							<PARAM NAME="AllowAddNew" VALUE="0">
							<PARAM NAME="AllowDelete" VALUE="0">
							<PARAM NAME="AllowUpdate" VALUE="1">
							<PARAM NAME="MultiLine" VALUE="0">
							<PARAM NAME="ActiveCellStyleSet" VALUE="Highlight">
							<PARAM NAME="RowSelectionStyle" VALUE="0">
							<PARAM NAME="AllowRowSizing" VALUE="1">
							<PARAM NAME="AllowGroupSizing" VALUE="1">
							<PARAM NAME="AllowColumnSizing" VALUE="1">
							<PARAM NAME="AllowGroupMoving" VALUE="0">
							<PARAM NAME="AllowColumnMoving" VALUE="0">
							<PARAM NAME="AllowGroupSwapping" VALUE="0">
							<PARAM NAME="AllowColumnSwapping" VALUE="0">
							<PARAM NAME="AllowGroupShrinking" VALUE="1">
							<PARAM NAME="AllowColumnShrinking" VALUE="1">
							<PARAM NAME="AllowDragDrop" VALUE="0">
							<PARAM NAME="UseExactRowCount" VALUE="1">
							<PARAM NAME="SelectTypeCol" VALUE="0">
							<PARAM NAME="SelectTypeRow" VALUE="0">
							<PARAM NAME="SelectByCell" VALUE="1">
							<PARAM NAME="BalloonHelp" VALUE="0">
							<PARAM NAME="RowNavigation" VALUE="0">
							<PARAM NAME="CellNavigation" VALUE="0">
							<PARAM NAME="MaxSelectedRows" VALUE="1">
							<PARAM NAME="HeadStyleSet" VALUE="">
							<PARAM NAME="StyleSet" VALUE="">
							<PARAM NAME="ForeColorEven" VALUE="0">
							<PARAM NAME="ForeColorOdd" VALUE="0">
							<PARAM NAME="BackColorEven" VALUE="-2147483643">
							<PARAM NAME="BackColorOdd" VALUE="-2147483643">
							<PARAM NAME="Levels" VALUE="1">
							<PARAM NAME="RowHeight" VALUE="239">
							<PARAM NAME="ExtraHeight" VALUE="239">
							<PARAM NAME="ActiveRowStyleSet" VALUE="">
							<PARAM NAME="CaptionAlignment" VALUE="2">
							<PARAM NAME="SplitterPos" VALUE="0">
							<PARAM NAME="SplitterVisible" VALUE="0">
							<PARAM NAME="Columns.Count" VALUE="1">
							<PARAM NAME="Columns(0).Width" VALUE="3528">
							<PARAM NAME="Columns(0).Visible" VALUE="-1">
							<PARAM NAME="Columns(0).Columns.Count" VALUE="1">
							<PARAM NAME="Columns(0).Caption" VALUE="  ">
							<PARAM NAME="Columns(0).Name" VALUE="">
							<PARAM NAME="Columns(0).Alignment" VALUE="0">
							<PARAM NAME="Columns(0).CaptionAlignment" VALUE="3">
							<PARAM NAME="Columns(0).Bound" VALUE="0">
							<PARAM NAME="Columns(0).AllowSizing" VALUE="1">
							<PARAM NAME="Columns(0).DataField" VALUE="">
							<PARAM NAME="Columns(0).DataType" VALUE="8">
							<PARAM NAME="Columns(0).Level" VALUE="0">
							<PARAM NAME="Columns(0).NumberFormat" VALUE="">
							<PARAM NAME="Columns(0).Case" VALUE="0">
							<PARAM NAME="Columns(0).FieldLen" VALUE="4096">
							<PARAM NAME="Columns(0).VertScrollBar" VALUE="0">
							<PARAM NAME="Columns(0).Locked" VALUE="0">
							<PARAM NAME="Columns(0).Style" VALUE="0">
							<PARAM NAME="Columns(0).ButtonsAlways" VALUE="0">
							<PARAM NAME="Columns(0).RowCount" VALUE="0">
							<PARAM NAME="Columns(0).ColCount" VALUE="1">
							<PARAM NAME="Columns(0).HasHeadForeColor" VALUE="0">
							<PARAM NAME="Columns(0).HasHeadBackColor" VALUE="0">
							<PARAM NAME="Columns(0).HasForeColor" VALUE="0">
							<PARAM NAME="Columns(0).HasBackColor" VALUE="0">
							<PARAM NAME="Columns(0).HeadForeColor" VALUE="0">
							<PARAM NAME="Columns(0).HeadBackColor" VALUE="0">
							<PARAM NAME="Columns(0).ForeColor" VALUE="0">
							<PARAM NAME="Columns(0).BackColor" VALUE="0">
							<PARAM NAME="Columns(0).HeadStyleSet" VALUE="">
							<PARAM NAME="Columns(0).StyleSet" VALUE="">
							<PARAM NAME="Columns(0).Nullable" VALUE="1">
							<PARAM NAME="Columns(0).Mask" VALUE="">
							<PARAM NAME="Columns(0).PromptInclude" VALUE="0">
							<PARAM NAME="Columns(0).ClipMode" VALUE="0">
							<PARAM NAME="Columns(0).PromptChar" VALUE="95">
							<PARAM NAME="UseDefaults" VALUE="-1">
							<PARAM NAME="TabNavigation" VALUE="1">
							<PARAM NAME="BatchUpdate" VALUE="0">
							<PARAM NAME="_ExtentX" VALUE="2646">
							<PARAM NAME="_ExtentY" VALUE="1323">
							<PARAM NAME="_StockProps" VALUE="79">
							<PARAM NAME="Caption" VALUE="SSDBGrid1">
							<PARAM NAME="ForeColor" VALUE="0">
							<PARAM NAME="BackColor" VALUE="16777215">
							<PARAM NAME="Enabled" VALUE="-1">
							<PARAM NAME="DataMember" VALUE=""></OBJECT>
					</TD>
				</TR>

                <tr height="5">
                    <td colspan="50">
                        <table width="100%" class="outline" cellspacing="0" cellpadding="0">
                            <tr height="5">
                                <td></td>
                            </tr>

                            <TR>
								<TD>&nbsp;&nbsp;<U>Intersection</U>
									<table WIDTH="100%" class="invisible" CELLSPACING=0 CELLPADDING=0>
                                    <td>
                                        <table width="100%" class="invisible" cellspacing="0" cellpadding="0">
                                            <% If CLng(Session("utiltype")) = 15 Then%>
                                                <input type="HIDDEN" id="txtIntersectionColumn" name="txtIntersectionColumn" style="BACKGROUND-COLOR: threedface; WIDTH: 100%" 

readonly></table>
                                                </td>
                                            <% Else%>
                                                <td width="20"></td>
                                                <td width="100">Column :</td>
                                                <td width="5"></td>
                                                <td width="300">
                                                    <input id="Text1" name="txtIntersectionColumn" class="text textdisabled" style="WIDTH: 100%" 

disabled="disabled"></td>

                                                <tr height="5">
                                                    <td></td>
                                                </tr>
                                            <% End If%>

								            <td width="20"></td>
                                            <td width="100" valign="top">Type :</td>
                                            <td width="5"></td>
                                            <td width="300" valign="top">
                                                <select id="cboIntersectionType" name="cboIntersectionType" class="combo" style="WIDTH: 100%" onchange="UpdateGrid()"></select>
                                            </td>
                                            <td width="20"></td>

                                        </table>
                                    
										<TR height=5>
											<TD></TD>
										</TR>
                                        
			</table>
									</TD>
									<TD width="20%" valign=top nowrap>


				        <input type="checkbox" id="chkPercentType" name="chkPercentType" value="checkbox" 
				            onclick="chkPercentType_Click()" 
		                    onmouseover="try{checkbox_onMouseOver(this);}catch(e){}" 
		                    onmouseout="try{checkbox_onMouseOut(this);}catch(e){}" />
                        <label 
	                        for="chkPercentType"
	                        class="checkbox"
	                        tabindex=0 
	                        onkeypress="try{checkboxLabel_onKeyPress(this);}catch(e){}"
		                    onmouseover="try{checkboxLabel_onMouseOver(this);}catch(e){}" 
		                    onmouseout="try{checkboxLabel_onMouseOut(this);}catch(e){}"
                            onfocus="try{checkboxLabel_onFocus(this);}catch(e){}"
                            onblur="try{checkboxLabel_onBlur(this);}catch(e){}">
<%
	if objCrossTab.CrossTabType <> 3 then
		Response.write(" Percentage of Type")
	end if
%>
      	    		        </label>
										<BR>
										
				        <input type="checkbox" id="chkPercentPage" name="chkPercentPage" value="checkbox"
				            onclick="UpdateGrid();" 
		                    onmouseover="try{checkbox_onMouseOver(this);}catch(e){}" 
		                    onmouseout="try{checkbox_onMouseOut(this);}catch(e){}" />
                        <label 
	                        for="chkPercentPage"
	                        class="checkbox"
	                        tabindex=0 
	                        onkeypress="try{checkboxLabel_onKeyPress(this);}catch(e){}"
		                    onmouseover="try{checkboxLabel_onMouseOver(this);}catch(e){}" 
		                    onmouseout="try{checkboxLabel_onMouseOut(this);}catch(e){}"
                            onfocus="try{checkboxLabel_onFocus(this);}catch(e){}"
                            onblur="try{checkboxLabel_onBlur(this);}catch(e){}">									
<%
	if objCrossTab.CrossTabType <> 3 then
		Response.write(" Percentage of Page")
	end if
%>
                        </label>
                        
										<BR>

				        <input type="checkbox" id="chkSuppressZeros" name="chkSuppressZeros" value="checkbox" 
				            onclick="UpdateGrid()" 
		                    onmouseover="try{checkbox_onMouseOver(this);}catch(e){}" 
		                    onmouseout="try{checkbox_onMouseOut(this);}catch(e){}" />
                        <label 
	                        for="chkSuppressZeros"
	                        class="checkbox"
	                        tabindex=0 
	                        onkeypress="try{checkboxLabel_onKeyPress(this);}catch(e){}"
		                    onmouseover="try{checkboxLabel_onMouseOver(this);}catch(e){}" 
		                    onmouseout="try{checkboxLabel_onMouseOut(this);}catch(e){}"
                            onfocus="try{checkboxLabel_onFocus(this);}catch(e){}"
                            onblur="try{checkboxLabel_onBlur(this);}catch(e){}">									
                            Suppress Zeros<br>
                         </label>

				        <input type="checkbox" id="chkUse1000" name="chkUse1000" value="checkbox" 
				            onclick="UpdateGrid()" 
		                    onmouseover="try{checkbox_onMouseOver(this);}catch(e){}" 
		                    onmouseout="try{checkbox_onMouseOut(this);}catch(e){}" />
                        <label 
	                        for="chkUse1000"
	                        class="checkbox"
	                        tabindex=0 
	                        onkeypress="try{checkboxLabel_onKeyPress(this);}catch(e){}"
		                    onmouseover="try{checkboxLabel_onMouseOver(this);}catch(e){}" 
		                    onmouseout="try{checkboxLabel_onMouseOut(this);}catch(e){}"
                            onfocus="try{checkboxLabel_onFocus(this);}catch(e){}"
                            onblur="try{checkboxLabel_onBlur(this);}catch(e){}">									
<%
	if objCrossTab.CrossTabType <> 3 then
		Response.write(" Use 1000 Separators (,)")
	end if
%>
                         </label>

										<BR>
									</TD>
									<TD id=CrossTabPage name=CrossTabPage>&nbsp;&nbsp;<U>Page</U>
                                        <table width="100%" outline="invisible" cellspacing="0" cellpadding="0">
                                            <tr>
                                                <td width="20"></td>
                                                <td width="100">Column :</td>
                                                <td width="5"></td>
                                                <td width="300">
                                                    <input id="txtPageColumn" name="txtPageColumn" style="WIDTH: 100%" class="text textdisabled" disabled="disabled"></td>
                                                <tr height="5">
                                                    <td></td>
                                                </tr>
                                                <td width="20"></td>
                                                <td width="100">Value :</td>
                                                <td width="5"></td>
                                                <td width="300">
                                                    <select id="cboPage" name="cboPage" style="WIDTH: 100%" class="combo" onchange="UpdateGrid()">
                                                    </select>
                                                </td>
                                                <td width="20"></td>
                                            </tr>
                                        </table>
									</TD>
								</TR>

							  <TR height=1>
							    <TD width=40%></TD>
							    <TD width=150></TD>
							    <TD></TD>
							  </TR>
							</table>

    <table width="100%" class="invisible" cellspacing="0" cellpadding="0">
        <tr height="5">
            <td></td>
        </tr>
        <tr height="5">
            <td colspan="3">
                <table width="100%" class="invisible" cellspacing="0" cellpadding="0">
                    <td align="RIGHT">
                        <input type="button" id="cmdOutput" name="cmdOutput" value="Output" style="WIDTH: 80px"
                            onclick="ViewExportOptions();"
                            onmouseover="try{button_onMouseOver(this);}catch(e){}"
                            onmouseout="try{button_onMouseOut(this);}catch(e){}"
                            onfocus="try{button_onFocus(this);}catch(e){}"
                            onblur="try{button_onBlur(this);}catch(e){}" />
                    </td>
                    <td width="15"></td>
                    <td width="5" align="RIGHT">
                        <input type="button" id="cmdClose" name="cmdClose" value="Close" style="WIDTH: 80px" class="btn"
                            onclick="try { closeclick(); } catch (e) { }"
                            onmouseover="try{button_onMouseOver(this);}catch(e){}"
                            onmouseout="try{button_onMouseOut(this);}catch(e){}"
                            onfocus="try{button_onFocus(this);}catch(e){}"
                            onblur="try{button_onBlur(this);}catch(e){}" />
                    </td>
                </table>
            </td>
        </tr>
        <tr height="5">
            <td></td>
        </tr>
    </table>


</table>
</table>
</table>


<form id="frmOriginalDefinition">
    <input type="hidden" id="txtDefn_Name" name="txtDefn_Name" value="<%=session("utilname")%>">
    <input type="hidden" id="txtUserName" name="txtUserName" value="<%=session("username")%>">
    <input type="hidden" id="txtDateFormat" name="txtDateFormat" value="<%=session("LocaleDateFormat")%>">
    <input type="hidden" id="txtDatabase" name="txtDatabase" value="<%=session("database")%>">

    <input type="hidden" id="txtCurrentPrintPage" name="txtCurrentPrintPage">
    <input type="hidden" id="txtCancelPrint" name="txtCancelPrint">
    <input type="hidden" id="txtOptionsDone" name="txtOptionsDone">
    <input type="hidden" id="txtOptionsPortrait" name="txtOptionsPortrait">
    <input type="hidden" id="txtOptionsMarginLeft" name="txtOptionsMarginLeft">
    <input type="hidden" id="txtOptionsMarginRight" name="txtOptionsMarginRight">
    <input type="hidden" id="txtOptionsMarginTop" name="txtOptionsMarginTop">
    <input type="hidden" id="txtOptionsMarginBottom" name="txtOptionsMarginBottom">
    <input type="hidden" id="txtOptionsCopies" name="txtOptionsCopies">
</form>

<object classid="clsid:4A4AA697-3E6F-11D2-822F-00104B9E07A1"
    codebase="cabs/COAInt_Grid.cab#version=3,1,3,6"
    id="ssHiddenGrid" name="ssHiddenGrid"
    style="HEIGHT: 0px; LEFT: 0px; TOP: 0px; WIDTH: 0px; POSITION: absolute">
    <param name="ScrollBars" value="4">
    <param name="_Version" value="196617">
    <param name="DataMode" value="2">
    <param name="Cols" value="0">
    <param name="Rows" value="0">
    <param name="BorderStyle" value="1">
    <param name="RecordSelectors" value="0">
    <param name="GroupHeaders" value="1">
    <param name="ColumnHeaders" value="1">
    <param name="GroupHeadLines" value="1">
    <param name="HeadLines" value="1">
    <param name="FieldDelimiter" value="(None)">
    <param name="FieldSeparator" value="(Tab)">
    <param name="Row.Count" value="0">
    <param name="Col.Count" value="1">
    <param name="stylesets.count" value="1">
    <param name="TagVariant" value="EMPTY">
    <param name="UseGroups" value="0">
    <param name="HeadFont3D" value="0">
    <param name="Font3D" value="0">
    <param name="DividerType" value="3">
    <param name="DividerStyle" value="1">
    <param name="DefColWidth" value="3528">
    <param name="BeveColorScheme" value="2">
    <param name="BevelColorFrame" value="-2147483642">
    <param name="BevelColorHighlight" value="-2147483643">
    <param name="BevelColorShadow" value="-2147483632">
    <param name="BevelColorFace" value="-2147483633">
    <param name="CheckBox3D" value="1">
    <param name="AllowAddNew" value="0">
    <param name="AllowDelete" value="0">
    <param name="AllowUpdate" value="1">
    <param name="MultiLine" value="0">
    <param name="ActiveCellStyleSet" value="Highlight">
    <param name="RowSelectionStyle" value="0">
    <param name="AllowRowSizing" value="1">
    <param name="AllowGroupSizing" value="1">
    <param name="AllowColumnSizing" value="1">
    <param name="AllowGroupMoving" value="0">
    <param name="AllowColumnMoving" value="0">
    <param name="AllowGroupSwapping" value="0">
    <param name="AllowColumnSwapping" value="0">
    <param name="AllowGroupShrinking" value="1">
    <param name="AllowColumnShrinking" value="1">
    <param name="AllowDragDrop" value="0">
    <param name="UseExactRowCount" value="1">
    <param name="SelectTypeCol" value="0">
    <param name="SelectTypeRow" value="0">
    <param name="SelectByCell" value="1">
    <param name="BalloonHelp" value="0">
    <param name="RowNavigation" value="0">
    <param name="CellNavigation" value="0">
    <param name="MaxSelectedRows" value="1">
    <param name="HeadStyleSet" value="">
    <param name="StyleSet" value="">
    <param name="ForeColorEven" value="0">
    <param name="ForeColorOdd" value="0">
    <param name="BackColorEven" value="-2147483643">
    <param name="BackColorOdd" value="-2147483643">
    <param name="Levels" value="1">
    <param name="RowHeight" value="239">
    <param name="ExtraHeight" value="239">
    <param name="ActiveRowStyleSet" value="">
    <param name="CaptionAlignment" value="2">
    <param name="SplitterPos" value="0">
    <param name="SplitterVisible" value="0">
    <param name="Columns.Count" value="1">
    <param name="Columns(0).Width" value="3528">
    <param name="Columns(0).Visible" value="-1">
    <param name="Columns(0).Columns.Count" value="1">
    <param name="Columns(0).Caption" value="  ">
    <param name="Columns(0).Name" value="">
    <param name="Columns(0).Alignment" value="0">
    <param name="Columns(0).CaptionAlignment" value="3">
    <param name="Columns(0).Bound" value="0">
    <param name="Columns(0).AllowSizing" value="1">
    <param name="Columns(0).DataField" value="">
    <param name="Columns(0).DataType" value="8">
    <param name="Columns(0).Level" value="0">
    <param name="Columns(0).NumberFormat" value="">
    <param name="Columns(0).Case" value="0">
    <param name="Columns(0).FieldLen" value="4096">
    <param name="Columns(0).VertScrollBar" value="0">
    <param name="Columns(0).Locked" value="0">
    <param name="Columns(0).Style" value="0">
    <param name="Columns(0).ButtonsAlways" value="0">
    <param name="Columns(0).RowCount" value="0">
    <param name="Columns(0).ColCount" value="1">
    <param name="Columns(0).HasHeadForeColor" value="0">
    <param name="Columns(0).HasHeadBackColor" value="0">
    <param name="Columns(0).HasForeColor" value="0">
    <param name="Columns(0).HasBackColor" value="0">
    <param name="Columns(0).HeadForeColor" value="0">
    <param name="Columns(0).HeadBackColor" value="0">
    <param name="Columns(0).ForeColor" value="0">
    <param name="Columns(0).BackColor" value="0">
    <param name="Columns(0).HeadStyleSet" value="">
    <param name="Columns(0).StyleSet" value="">
    <param name="Columns(0).Nullable" value="1">
    <param name="Columns(0).Mask" value="">
    <param name="Columns(0).PromptInclude" value="0">
    <param name="Columns(0).ClipMode" value="0">
    <param name="Columns(0).PromptChar" value="95">
    <param name="UseDefaults" value="-1">
    <param name="TabNavigation" value="1">
    <param name="BatchUpdate" value="0">
    <param name="_ExtentX" value="2646">
    <param name="_ExtentY" value="1323">
    <param name="_StockProps" value="79">
    <param name="Caption" value="SSDBGrid1">
    <param name="ForeColor" value="0">
    <param name="BackColor" value="16777215">
    <param name="Enabled" value="-1">
    <param name="DataMember" value="">
</object>

<form target="Output" action="util_run_outputoptions" method="post" id="frmExportData" name="frmExportData">
    <input type="hidden" id="txtPreview" name="txtPreview" value="">
    <input type="hidden" id="txtFormat" name="txtFormat" value=0>
    <input type="hidden" id="txtScreen" name="txtScreen" value="">
    <input type="hidden" id="txtPrinter" name="txtPrinter" value="">
    <input type="hidden" id="txtPrinterName" name="txtPrinterName" value="">
    <input type="hidden" id="txtSave" name="txtSave" value="">
    <input type="hidden" id="txtSaveExisting" name="txtSaveExisting" value="">
    <input type="hidden" id="txtEmail" name="txtEmail" value="">
    <input type="hidden" id="txtEmailAddr" name="txtEmailAddr" value="">
    <input type="hidden" id="txtEmailAddrName" name="txtEmailAddrName" value="">
    <input type="hidden" id="txtEmailSubject" name="txtEmailSubject" value="">
    <input type="hidden" id="txtEmailAttachAs" name="txtEmailAttachAs" value="">
    <input type="hidden" id="txtEmailGroupAddr" name="txtEmailGroupAddr" value="">
    <input type="hidden" id="txtFileName" name="txtFileName" value="">
    <input type="hidden" id="txtUtilType" name="txtUtilType" value="<%=session("utilType")%>">
</form>

<select style="visibility: hidden; display: none" id="cboDummy" name="cboDummy">
</select>



