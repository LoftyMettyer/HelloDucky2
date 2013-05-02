<%@ Page Language="VB" Inherits="System.Web.Mvc.ViewPage" %>

<!DOCTYPE html>
<html>
<head>

    <title>OpenHR Intranet</title>
    <script src="<%: Url.Content("~/bundles/jQuery")%>" type="text/javascript"></script>
    <script src="<%: Url.Content("~/bundles/OpenHR_General")%>" type="text/javascript"></script>           

<script type="text/javascript">

	window.onload = function () {

		var ssOleDBGridDefSelRecords = document.getElementById("ssOleDBGridDefSelRecords");
		var bdyMain = document.getElementById("bdyMain");
		
		setGridFont(ssOleDBGridDefSelRecords);

		var iResizeBy, iNewWidth, iNewHeight;

		// Resize the popup.
		iResizeBy = bdyMain.scrollWidth - bdyMain.clientWidth;
		if (bdyMain.offsetWidth + iResizeBy > screen.width) {
			window.dialogWidth = new String(screen.width) + "px";
		}
		else {
			iNewWidth = new Number(window.dialogWidth.substr(0, window.dialogWidth.length - 2));
			iNewWidth = iNewWidth + iResizeBy;
			window.dialogWidth = new String(iNewWidth) + "px";
		}

		iResizeBy = bdyMain.scrollHeight - bdyMain.clientHeight;
		if (bdyMain.offsetHeight + iResizeBy > screen.height) {
			window.dialogHeight = new String(screen.height) + "px";
		}
		else {
			iNewHeight = new Number(window.dialogHeight.substr(0, window.dialogHeight.length - 2));
			iNewHeight = iNewHeight + iResizeBy;
			window.dialogHeight = new String(iNewHeight) + "px";
		}
	}

</script>

<script type="text/javascript">

    function setForm() {

		var frmPopup = document.getElementById("frmPopup");
		window.dialogArguments.document.getElementById('txtEmailGroup').value = frmPopup.txtSelectedName.value;
		window.dialogArguments.document.getElementById('txtEmailGroupID').value = frmPopup.txtSelectedID.value;

		self.close();
		return false;
	}

	function locateRecord(psSearchFor) {
		
		var fFound = false;
		var ssOleDBGridDefSelRecords = document.getElementById("ssOleDBGridDefSelRecords");
		ssOleDBGridDefSelRecords.Redraw = false;
		ssOleDBGridDefSelRecords.MoveLast();
		ssOleDBGridDefSelRecords.MoveFirst();

		for (var iIndex = 1; iIndex <= ssOleDBGridDefSelRecords.rows; iIndex++) {
			var sGridValue = new String(ssOleDBGridDefSelRecords.Columns("name").value);
			sGridValue = sGridValue.substr(0, psSearchFor.length).toUpperCase();
			if (sGridValue == psSearchFor.toUpperCase()) {
				ssOleDBGridDefSelRecords.SelBookmarks.Add(ssOleDBGridDefSelRecords.Bookmark);
				fFound = true;
				break;
			}

			if (iIndex < ssOleDBGridDefSelRecords.rows) {
				ssOleDBGridDefSelRecords.MoveNext();
			}
			else {
				break;
			}
		}

		if ((fFound == false) && (ssOleDBGridDefSelRecords.rows > 0)) {
			// Select the top row.
			ssOleDBGridDefSelRecords.MoveFirst();
			ssOleDBGridDefSelRecords.SelBookmarks.Add(ssOleDBGridDefSelRecords.Bookmark);
		}

		ssOleDBGridDefSelRecords.redraw = true;
	}

</script>

</head>

<body id=bdyMain name=bdyMain <%=session("BodyColour")%> leftmargin=20 topmargin=20 bottommargin=20 rightmargin=20>
	
	<form id="frmPopup" name="frmPopup" onsubmit="return setForm();" style="visibility: hidden;display: none">
	<input type="hidden" id="txtSelectedID" name="txtSelectedID">
	<input type="hidden" id="txtSelectedName" name="txtSelectedName">
	<input type="hidden" id="txtSelectedAccess" name="txtSelectedAccess">
	<input type="hidden" id="txtSelectedUserName" name="txtSelectedUserName">
	</form>

<table align=center class="outline" cellpadding=5 cellspacing=0 width=100% height=100%>
	<tr>
		<td>
			<table width="100%" height="100%" class="invisible" cellspacing="0" cellpadding="0" >
				<tr height=10>
					<td colspan=3 align=center height=10>
						<H3>Email Groups</H3>
					</td>
				</tr>
				<tr>
					<td width=20></td>
					<td>
<%
	' Get the order records.
	Dim cmdDefSelRecords = CreateObject("ADODB.Command")
	cmdDefSelRecords.CommandText = "spASRIntGetEmailGroups"
	cmdDefSelRecords.CommandType = 4 ' Stored Procedure
	cmdDefSelRecords.ActiveConnection = Session("databaseConnection")
	Err.Clear()
	Dim rstDefSelRecords = cmdDefSelRecords.Execute

	' Instantiate and initialise the grid. 
%>
	                    <OBJECT classid="clsid:4A4AA697-3E6F-11D2-822F-00104B9E07A1" id=ssOleDBGridDefSelRecords name=ssOleDBGridDefselRecords codebase="cabs/COAInt_Grid.cab#version=3,1,3,6" style="LEFT: 0px; TOP: 0px; WIDTH:100%; HEIGHT:300px">
					        <PARAM NAME="ScrollBars" VALUE="4">
					        <PARAM NAME="_Version" VALUE="196616">
					        <PARAM NAME="DataMode" VALUE="2">
					        <PARAM NAME="Cols" VALUE="0">
					        <PARAM NAME="Rows" VALUE="0">
					        <PARAM NAME="BorderStyle" VALUE="1">
					        <PARAM NAME="RecordSelectors" VALUE="0">
					        <PARAM NAME="GroupHeaders" VALUE="0">
					        <PARAM NAME="ColumnHeaders" VALUE="0">
					        <PARAM NAME="GroupHeadLines" VALUE="0">
					        <PARAM NAME="HeadLines" VALUE="0">
					        <PARAM NAME="FieldDelimiter" VALUE="(None)">
					        <PARAM NAME="FieldSeparator" VALUE="(Tab)">
					        <PARAM NAME="Col.Count" VALUE="<%=rstDefselRecords.fields.count%>">
					        <PARAM NAME="stylesets.count" VALUE="0">
					        <PARAM NAME="TagVariant" VALUE="EMPTY">
					        <PARAM NAME="UseGroups" VALUE="0">
					        <PARAM NAME="HeadFont3D" VALUE="0">
					        <PARAM NAME="Font3D" VALUE="0">
					        <PARAM NAME="DividerType" VALUE="3">
					        <PARAM NAME="DividerStyle" VALUE="1">
					        <PARAM NAME="DefColWidth" VALUE="0">
					        <PARAM NAME="BeveColorScheme" VALUE="2">
					        <PARAM NAME="BevelColorFrame" VALUE="-2147483642">
					        <PARAM NAME="BevelColorHighlight" VALUE="-2147483628">
					        <PARAM NAME="BevelColorShadow" VALUE="-2147483632">
					        <PARAM NAME="BevelColorFace" VALUE="-2147483633">
					        <PARAM NAME="CheckBox3D" VALUE="-1">
					        <PARAM NAME="AllowAddNew" VALUE="0">
					        <PARAM NAME="AllowDelete" VALUE="0">
					        <PARAM NAME="AllowUpdate" VALUE="0">
					        <PARAM NAME="MultiLine" VALUE="0">
					        <PARAM NAME="ActiveCellStyleSet" VALUE="">
					        <PARAM NAME="RowSelectionStyle" VALUE="0">
					        <PARAM NAME="AllowRowSizing" VALUE="0">
					        <PARAM NAME="AllowGroupSizing" VALUE="0">
					        <PARAM NAME="AllowColumnSizing" VALUE="0">
					        <PARAM NAME="AllowGroupMoving" VALUE="0">
					        <PARAM NAME="AllowColumnMoving" VALUE="0">
					        <PARAM NAME="AllowGroupSwapping" VALUE="0">
					        <PARAM NAME="AllowColumnSwapping" VALUE="0">
					        <PARAM NAME="AllowGroupShrinking" VALUE="0">
					        <PARAM NAME="AllowColumnShrinking" VALUE="0">
					        <PARAM NAME="AllowDragDrop" VALUE="0">
					        <PARAM NAME="UseExactRowCount" VALUE="-1">
					        <PARAM NAME="SelectTypeCol" VALUE="0">
					        <PARAM NAME="SelectTypeRow" VALUE="1">
					        <PARAM NAME="SelectByCell" VALUE="-1">
					        <PARAM NAME="BalloonHelp" VALUE="0">
					        <PARAM NAME="RowNavigation" VALUE="1">
					        <PARAM NAME="CellNavigation" VALUE="0">
					        <PARAM NAME="MaxSelectedRows" VALUE="1">
					        <PARAM NAME="HeadStyleSet" VALUE="">
					        <PARAM NAME="StyleSet" VALUE="">
					        <PARAM NAME="ForeColorEven" VALUE="0">
					        <PARAM NAME="ForeColorOdd" VALUE="0">
					        <PARAM NAME="BackColorEven" VALUE="16777215">
					        <PARAM NAME="BackColorOdd" VALUE="16777215">
					        <PARAM NAME="Levels" VALUE="1">
					        <PARAM NAME="RowHeight" VALUE="503">
					        <PARAM NAME="ExtraHeight" VALUE="0">
					        <PARAM NAME="ActiveRowStyleSet" VALUE="">
					        <PARAM NAME="CaptionAlignment" VALUE="2">
					        <PARAM NAME="SplitterPos" VALUE="0">
					        <PARAM NAME="SplitterVisible" VALUE="0">
					        <PARAM NAME="Columns.Count" VALUE="<%=rstDefSelRecords.fields.count%>">
<%
	for iLoop = 0 to (rstDefSelRecords.fields.count - 1)
		if lcase(rstDefSelRecords.fields(iLoop).name) = "name" then
%>
							<PARAM NAME="Columns(<%=iLoop%>).Width" VALUE="100000">
							<PARAM NAME="Columns(<%=iLoop%>).Visible" VALUE="-1">
<%
		else
%>
							<PARAM NAME="Columns(<%=iLoop%>).Width" VALUE="0">
							<PARAM NAME="Columns(<%=iLoop%>).Visible" VALUE="0">
<%
		end if
%>							
						    <PARAM NAME="Columns(<%=iLoop%>).Columns.Count" VALUE="1">
						    <PARAM NAME="Columns(<%=iLoop%>).Caption" VALUE="<%=replace(rstDefSelRecords.fields(iLoop).name, "_", " ")%>">
						    <PARAM NAME="Columns(<%=iLoop%>).Name" VALUE="<%=rstDefSelRecords.fields(iLoop).name%>">			
						    <PARAM NAME="Columns(<%=iLoop%>).Alignment" VALUE="0">
						    <PARAM NAME="Columns(<%=iLoop%>).CaptionAlignment" VALUE="3">
						    <PARAM NAME="Columns(<%=iLoop%>).Bound" VALUE="0">
						    <PARAM NAME="Columns(<%=iLoop%>).AllowSizing" VALUE="1">
						    <PARAM NAME="Columns(<%=iLoop%>).DataField" VALUE="Column <%=iLoop%>">
						    <PARAM NAME="Columns(<%=iLoop%>).DataType" VALUE="8">
						    <PARAM NAME="Columns(<%=iLoop%>).Level" VALUE="0">
						    <PARAM NAME="Columns(<%=iLoop%>).NumberFormat" VALUE="">			
						    <PARAM NAME="Columns(<%=iLoop%>).Case" VALUE="0">
						    <PARAM NAME="Columns(<%=iLoop%>).FieldLen" VALUE="4096">
						    <PARAM NAME="Columns(<%=iLoop%>).VertScrollBar" VALUE="0">
						    <PARAM NAME="Columns(<%=iLoop%>).Locked" VALUE="0">			
						    <PARAM NAME="Columns(<%=iLoop%>).Style" VALUE="0">
						    <PARAM NAME="Columns(<%=iLoop%>).ButtonsAlways" VALUE="0">
						    <PARAM NAME="Columns(<%=iLoop%>).RowCount" VALUE="0">
						    <PARAM NAME="Columns(<%=iLoop%>).ColCount" VALUE="1">
						    <PARAM NAME="Columns(<%=iLoop%>).HasHeadForeColor" VALUE="0">
						    <PARAM NAME="Columns(<%=iLoop%>).HasHeadBackColor" VALUE="0">
						    <PARAM NAME="Columns(<%=iLoop%>).HasForeColor" VALUE="0">
						    <PARAM NAME="Columns(<%=iLoop%>).HasBackColor" VALUE="0">
						    <PARAM NAME="Columns(<%=iLoop%>).HeadForeColor" VALUE="0">
						    <PARAM NAME="Columns(<%=iLoop%>).HeadBackColor" VALUE="0">
						    <PARAM NAME="Columns(<%=iLoop%>).ForeColor" VALUE="0">
						    <PARAM NAME="Columns(<%=iLoop%>).BackColor" VALUE="0">
						    <PARAM NAME="Columns(<%=iLoop%>).HeadStyleSet" VALUE="">
						    <PARAM NAME="Columns(<%=iLoop%>).StyleSet" VALUE="">
						    <PARAM NAME="Columns(<%=iLoop%>).Nullable" VALUE="1">
						    <PARAM NAME="Columns(<%=iLoop%>).Mask" VALUE="">
						    <PARAM NAME="Columns(<%=iLoop%>).PromptInclude" VALUE="0">
						    <PARAM NAME="Columns(<%=iLoop%>).ClipMode" VALUE="0">
						    <PARAM NAME="Columns(<%=iLoop%>).PromptChar" VALUE="95">
<%
	next 
%>
					        <PARAM NAME="UseDefaults" VALUE="-1">
					        <PARAM NAME="TabNavigation" VALUE="1">
					        <PARAM NAME="_ExtentX" VALUE="17330">
					        <PARAM NAME="_ExtentY" VALUE="1323">
					        <PARAM NAME="_StockProps" VALUE="79">
					        <PARAM NAME="Caption" VALUE="">
					        <PARAM NAME="ForeColor" VALUE="0">
					        <PARAM NAME="BackColor" VALUE="16777215">
					        <PARAM NAME="Enabled" VALUE="-1">
					        <PARAM NAME="DataMember" VALUE="">
<%							
	Dim lngRowCount = 0
	do while not rstDefSelRecords.EOF
		for iLoop = 0 to (rstDefSelRecords.fields.count - 1)	
%>								
			                <PARAM NAME="Row(<%=lngRowCount%>).Col(<%=iLoop%>)" VALUE="<%=replace(replace(rstDefSelRecords.Fields(iLoop).Value, "_", " "), """", "&quot;")%>">
<%
		next 				
		lngRowCount = lngRowCount + 1
		rstDefSelRecords.MoveNext
	loop
%>
	                        <PARAM NAME="Row.Count" VALUE="<%=lngRowCount%>">
	                    </OBJECT>
					</td>
					<td width=20></td>
				</tr>
				<tr height=10>
					<td height=10 colspan=3>&nbsp;</td>
				</tr>
				<tr height=10>
					<td width=20></td>
					<td height=10>
						<table WIDTH=100% class="invisible" CELLSPACING=0 CELLPADDING=0>
							<TR>
								<TD>&nbsp;</TD>
								<TD width=10>
									<input id=cmdok type=button class="btn" value=OK name=cmdok style="WIDTH: 80px" width="80" 
									    onclick="setForm();" 
                                        onmouseover="try{button_onMouseOver(this);}catch(e){}" 
                                        onmouseout="try{button_onMouseOut(this);}catch(e){}"
                                        onfocus="try{button_onFocus(this);}catch(e){}"
                                        onblur="try{button_onBlur(this);}catch(e){}" />
								</TD>
								<TD width=10>&nbsp;</TD>
								<TD width=10>
									<input id=cmdnone type=button class="btn" value=None name=cmdnone style="WIDTH: 80px" width="80" 
									    onclick="frmPopup.txtSelectedID.value=0;frmPopup.txtSelectedName.value='';frmPopup.txtSelectedAccess.value='';frmPopup.txtSelectedUserName.value='';setForm();" 
                                        onmouseover="try{button_onMouseOver(this);}catch(e){}" 
                                        onmouseout="try{button_onMouseOut(this);}catch(e){}"
                                        onfocus="try{button_onFocus(this);}catch(e){}"
                                        onblur="try{button_onBlur(this);}catch(e){}" />
								</TD>
								<TD width=10>&nbsp;</TD>
								<TD width=10>
									<input id=cmdcancel type=button class="btn" value=Cancel name=cmdcancel style="WIDTH: 80px" width="80" onclick="self.close();" 
                                        onmouseover="try{button_onMouseOver(this);}catch(e){}" 
                                        onmouseout="try{button_onMouseOut(this);}catch(e){}"
                                        onfocus="try{button_onFocus(this);}catch(e){}"
                                        onblur="try{button_onBlur(this);}catch(e){}" />
								</TD>
							</tr>
						</table>
					</td>
				</tr>
			</table>	
		</td>
	</tr>
</table>

	<form id="frmFromOpener" name="frmFromOpener" style="visibility: hidden; display: none">
		<input type="hidden" id="calcEmailCurrentID" name="calcEmailCurrentID" value='<%= Request("emailSelCurrentID") %>'>
	</form>

	<input type='hidden' id="txtTicker" name="txtTicker" value="0">
	<input type='hidden' id="txtLastKeyFind" name="txtLastKeyFind" value="">
	
<script type="text/javascript">
<!--
		OpenHR.addActiveXHandler("ssOleDBGridDefSelRecords", "rowcolchange", ssOleDBGridDefSelRecords_rowcolchange);

		function ssOleDBGridDefSelRecords_rowcolchange() {
			// Populate the textboxs with the selected rows details
			var frmPopup = document.getElementById("frmPopup");
			frmPopup.txtSelectedID.value = document.getElementById('ssOleDBGridDefSelRecords').Columns(0).Value;
			frmPopup.txtSelectedUserName.value = document.getElementById('ssOleDBGridDefSelRecords').Columns("username").Value;
			frmPopup.txtSelectedAccess.value = document.getElementById('ssOleDBGridDefSelRecords').Columns("access").Value;
			frmPopup.txtSelectedName.value = document.getElementById('ssOleDBGridDefSelRecords').Columns("name").Value;
		}

		OpenHR.addActiveXHandler("ssOleDBGridDefSelRecords", "dblClick", ssOleDBGridDefSelRecords_dblClick);

		function ssOleDBGridDefSelRecords_dblClick() {
			setForm();
		}

		OpenHR.addActiveXHandler("ssOleDBGridDefSelRecords", "KeyPress", ssOleDBGridDefSelRecords_KeyPress);

		function ssOleDBGridDefSelRecords_KeyPress(iKeyAscii) {

			var txtLastKeyFind = document.getElementById("txtLastKeyFind"),
		    txtTicker = document.getElementById("txtTicker"),
		    sFind,
		    iLastTick;

			if ((iKeyAscii >= 32) && (iKeyAscii <= 255)) {
				var dtTicker = new Date();
				var iThisTick = new Number(dtTicker.getTime());
				if (txtLastKeyFind.value.length > 0) {
					iLastTick = new Number(txtTicker.value);
				}
				else {
					iLastTick = new Number("0");
				}

				if (iThisTick > (iLastTick + 1500)) {
					sFind = String.fromCharCode(iKeyAscii);
				}
				else {
					sFind = txtLastKeyFind.value + String.fromCharCode(iKeyAscii);
				}

				txtTicker.value = iThisTick;
				txtLastKeyFind.value = sFind;

				locateRecord(sFind);
			}
		}
-->
</script>
</body>
</html>
