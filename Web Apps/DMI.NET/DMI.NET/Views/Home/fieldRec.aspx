<%@ Page Language="VB" Inherits="System.Web.Mvc.ViewPage" %>
<%@ Import Namespace="DMI.NET" %>

<%
    Session("selectionType") = Request("selectionType")
	session("selectionTableID") = Request("txtTableID")
	session("selectedID") = Request("selectedID")
%>

<!DOCTYPE html>

<html>
<head runat="server">
    <title>OpenHR Intranet</title>

    <link href="<%: Url.Content("~/Content/OpenHR.css") %>" rel="stylesheet" type="text/css"/>
    <script src="<%: Url.Content("~/Scripts/jquery-1.8.2.js") %>" type="text/javascript"></script>
    <script src="<%: Url.Content("~/Scripts/openhr.js") %>" type="text/javascript"></script>
    <script src="<%: Url.Content("~/Scripts/ctl_SetFont.js") %>" type="text/javascript"></script>
    <script src="<%: Url.Content("~/Scripts/ctl_SetStyles.js") %>" type="text/javascript"></script>
            
    <object classid="clsid:5220cb21-c88d-11cf-b347-00aa00a28331" id="Microsoft_Licensed_Class_Manager_1_0" VIEWASTEXT>
	    <param NAME="LPKPath" VALUE="lpks/main.lpk">
    </object>



    <script type="text/javascript">
<!--
        function fieldRec_window_onload() {
            
            fOK = true;

            cmdCancel.focus();
	
            // Set focus onto one of the form controls. 
            // NB. This needs to be done before making any reference to the grid
            ssOleDBGridSelRecords.focus();
            locateRecordID(frmUseful.txtSelectedID.value);

            setGridFont(ssOleDBGridSelRecords);
	
            refreshControls();

            // Resize the popup.
            iResizeBy = bdyMain.scrollWidth	- bdyMain.clientWidth;
            if (bdyMain.offsetWidth + iResizeBy > screen.width) {
                window.dialogWidth = new String(screen.width) + "px";
            }
            else {
                iNewWidth = new Number(window.dialogWidth.substr(0, window.dialogWidth.length-2));
                iNewWidth = iNewWidth + iResizeBy;
                window.dialogWidth = new String(iNewWidth) + "px";
            }

            iResizeBy = bdyMain.scrollHeight	- bdyMain.clientHeight;
            if (bdyMain.offsetHeight + iResizeBy > screen.height) {
                window.dialogHeight = new String(screen.height) + "px";
            }
            else {
                iNewHeight = new Number(window.dialogHeight.substr(0, window.dialogHeight.length-2));
                iNewHeight = iNewHeight + iResizeBy;
                window.dialogHeight = new String(iNewHeight) + "px";
            }
        }

        function refreshControls()
        {
            button_disable(cmdOK, (ssOleDBGridSelRecords.SelBookmarks.Count == 0));
        }

        function setForm()
        {
            //we are doing this for the order
            if (frmUseful.txtSelectionType.value == 'ORDER') {
                window.dialogArguments.document.getElementById('txtFieldRecOrder').value = frmPopup.txtSelectedName.value;
                window.dialogArguments.document.getElementById('txtChildFieldOrderID').value = frmPopup.txtSelectedID.value;

                try {
                    window.dialogArguments.document.getElementById('btnFieldRecOrder').focus();
                }
                catch(e) {
                }
            }
            else {
                //we are doing this for the filter
                window.dialogArguments.document.getElementById('txtFieldRecFilter').value = frmPopup.txtSelectedName.value;
                window.dialogArguments.document.getElementById('txtChildFieldFilterID').value =  frmPopup.txtSelectedID.value;
			
                //if its hidden, set the relevant textbox value
                if (frmPopup.txtSelectedAccess.value == "HD") {
                    window.dialogArguments.document.getElementById('txtChildFieldFilterHidden').value = 'Y';
                }
                else {
                    window.dialogArguments.document.getElementById('txtChildFieldFilterHidden').value = '';
                }
			
                try {
                    window.dialogArguments.document.getElementById('btnFieldRecFilter').focus();
                }
                catch(e) {
                }
            }

            self.close();
            return false;
        }

        function makeSelection()
        {
            frmPopup.txtSelectedID.value = ssOleDBGridSelRecords.Columns("id").Value; 	
            frmPopup.txtSelectedUserName.value = ssOleDBGridSelRecords.Columns("username").Value;
            frmPopup.txtSelectedAccess.value = ssOleDBGridSelRecords.Columns("access").Value;
            frmPopup.txtSelectedName.value = ssOleDBGridSelRecords.Columns("name").Value;
            setForm();
        }

        function clearSelection()
        {
            frmPopup.txtSelectedID.value=0;
            frmPopup.txtSelectedName.value='';
            frmPopup.txtSelectedAccess.value='';
            frmPopup.txtSelectedUserName.value='';
            setForm();
        }

        function locateRecord(psSearchFor)
        {  
            var fFound

            fFound = false;
	
            ssOleDBGridSelRecords.redraw = false;

            ssOleDBGridSelRecords.MoveLast();
            ssOleDBGridSelRecords.MoveFirst();

            for (iIndex = 1; iIndex <= ssOleDBGridSelRecords.rows; iIndex++) {	
                var sGridValue = new String(ssOleDBGridSelRecords.Columns("name").value);
                sGridValue = sGridValue.substr(0, psSearchFor.length).toUpperCase();
                if (sGridValue == psSearchFor.toUpperCase()) {
                    ssOleDBGridSelRecords.SelBookmarks.Add(ssOleDBGridSelRecords.Bookmark);
                    fFound = true;
                    break;
                }

                if (iIndex < ssOleDBGridSelRecords.rows) {
                    ssOleDBGridSelRecords.MoveNext();
                }
                else {
                    break;
                }
            }

            if ((fFound == false) && (ssOleDBGridSelRecords.rows > 0)) {
                // Select the top row.
                ssOleDBGridSelRecords.MoveFirst();
                ssOleDBGridSelRecords.SelBookmarks.Add(ssOleDBGridSelRecords.Bookmark);
            }

            ssOleDBGridSelRecords.redraw = true;
        }

        function locateRecordID(piRecordID)
        {  
            var fFound

            fFound = false;
	
            ssOleDBGridSelRecords.redraw = false;

            ssOleDBGridSelRecords.MoveLast();
            ssOleDBGridSelRecords.MoveFirst();

            if (frmUseful.txtSelectedID.value > 0) {
                for (iIndex = 1; iIndex <= ssOleDBGridSelRecords.rows; iIndex++) {	
                    var sGridValue = new String(ssOleDBGridSelRecords.Columns("id").value);
                    if (sGridValue == piRecordID) {
                        ssOleDBGridSelRecords.SelBookmarks.Add(ssOleDBGridSelRecords.Bookmark);
                        fFound = true;
                        break;
                    }

                    if (iIndex < ssOleDBGridSelRecords.rows) {
                        ssOleDBGridSelRecords.MoveNext();
                    }
                    else {
                        break;
                    }
                }
            }
	
            if ((fFound == false) && (ssOleDBGridSelRecords.rows > 0)) {
                // Select the top row.
                ssOleDBGridSelRecords.MoveFirst();
                ssOleDBGridSelRecords.SelBookmarks.Add(ssOleDBGridSelRecords.Bookmark);
            }

            ssOleDBGridSelRecords.redraw = true;
        }

    -->
    </script>

    <script type="text/javascript">
<!--
        function fieldrec_addhandlers() {        
            OpenHR.addActiveXHandler("ssOleDBGridSelRecords", "rowcolchange", ssOleDBGridSelRecords_rowcolchange);
            OpenHR.addActiveXHandler("ssOleDBGridSelRecords", "dblClick", ssOleDBGridSelRecords_dblClick);
            OpenHR.addActiveXHandler("ssOleDBGridSelRecords", "KeyPress", ssOleDBGridSelRecords_KeyPress);
        }

        function ssOleDBGridSelRecords_rowcolchange() {
            // Populate the textboxs with the selected rows details
            refreshControls();
        }

        function ssOleDBGridSelRecords_dblClick() {
            makeSelection();        
        }

        function ssOleDBGridSelRecords_KeyPress(iKeyAscii) {

            if ((iKeyAscii >= 32) && (iKeyAscii <= 255)) {	
                var dtTicker = new Date();
                var iThisTick = new Number(dtTicker.getTime());
                if (txtLastKeyFind.value.length > 0) {
                    var iLastTick = new Number(txtTicker.value);
                }
                else {
                    var iLastTick = new Number("0");
                }
		
                if (iThisTick > (iLastTick + 1500)) {
                    var sFind = String.fromCharCode(iKeyAscii);
                }
                else {
                    var sFind = txtLastKeyFind.value + String.fromCharCode(iKeyAscii);
                }
		
                txtTicker.value = iThisTick;
                txtLastKeyFind.value = sFind;

                locateRecord(sFind);
            }
        
        }
    -->
</script>
   
    

</head>
<body id=bdyMain >
    
    <table align=center class="outline" cellPadding=5 cellSpacing=0 width=100% height=100%>
	<tr>
		<td>
			<table align=center class="invisible" cellspacing=0 cellpadding=0 width=100% height=100%>
				<tr height=10>
					<td colspan=3 align=center height=10>
						<H3 align=center>
<% 
	if ucase(session("selectionType")) = ucase("order") then 
        Response.Write("Select Order")
    Else
        Response.Write("Select Filter")
    End If
%>
						</H3>
					</td>
				</tr>
				<tr>
					<td width=20></td>
                    <td>
<%
    Dim cmdSelRecords
    Dim prmTableID
    Dim prmUser
    Dim rstSelRecords
    Dim lngRowCount As Long
    
    cmdSelRecords = Server.CreateObject("ADODB.Command")
	cmdSelRecords.CommandType = 4
    cmdSelRecords.ActiveConnection = Session("databaseConnection")

	if ucase(session("selectionType")) = ucase("order") then
		cmdSelRecords.CommandText = "spASRIntGetAvailableOrdersInfo"
        cmdSelRecords.CommandType = 4
        
        prmTableID = cmdSelRecords.CreateParameter("tableID", 3, 1) ' 3 = integer, 1 = input
        cmdSelRecords.Parameters.Append(prmTableID)
		prmTableID.value = clng(cleanNumeric(session("selectionTableID")))
	else
        cmdSelRecords.CommandText = "spASRIntGetAvailableFiltersInfo"
        cmdSelRecords.CommandType = 4

        prmTableID = cmdSelRecords.CreateParameter("tableID", 3, 1) ' 3 = integer, 1 = input
        cmdSelRecords.Parameters.Append(prmTableID)
		prmTableID.value = clng(cleanNumeric(session("selectionTableID")))
			
        prmUser = cmdSelRecords.CreateParameter("user", 200, 1, 8000) ' 200 = varchar, 1 = input, 8000=size
        cmdSelRecords.Parameters.Append(prmUser)
        prmUser.value = CStr(Session("username"))
	end if

    Err.Clear()
    rstSelRecords = cmdSelRecords.Execute

	' Instantiate and initialise the grid. 
%>
					    <OBJECT classid="clsid:4A4AA697-3E6F-11D2-822F-00104B9E07A1" id=ssOleDBGridSelRecords name=ssOleDBGridSelRecords codebase="cabs/COAInt_Grid.cab#version=3,1,3,6" style="LEFT: 0px; TOP: 0px; WIDTH:100%; HEIGHT:300px">
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
							<PARAM NAME="Col.Count" VALUE="<%=rstSelRecords.fields.count%>">
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
							<PARAM NAME="Columns.Count" VALUE="<%=rstSelRecords.fields.count%>">
<%
	for iLoop = 0 to (rstSelRecords.fields.count - 1)
		if rstSelRecords.fields(iLoop).name <> "name" then
%>
                            <PARAM NAME="Columns(<%=iLoop%>).Width" VALUE="0">
			                <PARAM NAME="Columns(<%=iLoop%>).Visible" VALUE="0">
<%
		else
%>
			                <PARAM NAME="Columns(<%=iLoop%>).Width" VALUE="100000">
			                <PARAM NAME="Columns(<%=iLoop%>).Visible" VALUE="-1">
<%
		end if
%>								
		                    <PARAM NAME="Columns(<%=iLoop%>).Columns.Count" VALUE="1">
		                    <PARAM NAME="Columns(<%=iLoop%>).Caption" VALUE="<%=replace(rstSelRecords.fields(iLoop).name, "_", " ")%>">
		                    <PARAM NAME="Columns(<%=iLoop%>).Name" VALUE="<%=rstSelRecords.fields(iLoop).name%>">
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
    lngRowCount = 0
    Do While Not rstSelRecords.EOF
        For iLoop = 0 To (rstSelRecords.fields.count - 1)
%>		
			                <PARAM NAME="Row(<%=lngRowCount%>).Col(<%=iLoop%>)" VALUE="<%=replace(replace(rstSelRecords.Fields(iLoop).Value, "_", " "), "", "&quot;")%>">
<%
		next 				
		lngRowCount = lngRowCount + 1
		rstSelRecords.MoveNext
	loop
%>
	                        <PARAM NAME="Row.Count" VALUE="<%=lngRowCount%>">
	                    </OBJECT>
<%	
	rstSelRecords.close
    rstSelRecords = Nothing

	' Release the ADO command object.
    cmdSelRecords = Nothing
%>

					</td>
					<td width=20></td>
				</tr>
				<tr height=10>
					<td height=10 colspan=3>&nbsp;</td>
				</tr>
				<tr height=10>
					<td width=20></td>
					<td height=10>
						<TABLE WIDTH=100% class="invisible" CELLSPACING=0 CELLPADDING=0>
							<TR>
								<TD>&nbsp;</TD>
								<TD width=10>
									<INPUT id=cmdOK type=button value=OK name=cmdOK class="btn" style="WIDTH: 80px" width="80"
									    onclick="makeSelection()" 
                                        onmouseover="try{button_onMouseOver(this);}catch(e){}" 
                                        onmouseout="try{button_onMouseOut(this);}catch(e){}"
                                        onfocus="try{button_onFocus(this);}catch(e){}"
                                        onblur="try{button_onBlur(this);}catch(e){}" />
								</TD>
								<TD width=10>&nbsp;</TD>
								<TD width=10>
									<INPUT id=cmdnone type=button value=None name=cmdnone class="btn" style="WIDTH: 80px" width="80"
									    onclick="clearSelection()" 
                                        onmouseover="try{button_onMouseOver(this);}catch(e){}" 
                                        onmouseout="try{button_onMouseOut(this);}catch(e){}"
                                        onfocus="try{button_onFocus(this);}catch(e){}"
                                        onblur="try{button_onBlur(this);}catch(e){}" />
								</TD>
								<TD width=10>&nbsp;</TD>
								<TD width=10>
									<INPUT id=cmdCancel type=button value=Cancel name=cmdCancel class="btn" style="WIDTH: 80px" width="80"
									    onclick="self.close()" 
                                        onmouseover="try{button_onMouseOver(this);}catch(e){}" 
                                        onmouseout="try{button_onMouseOut(this);}catch(e){}"
                                        onfocus="try{button_onFocus(this);}catch(e){}"
                                        onblur="try{button_onBlur(this);}catch(e){}" />
								</TD>
							</TR>
						</TABLE>
					</td>
					<td width=20></td>
				</tr>
			</TABLE>
		</td>
	</tr>
</table>

<INPUT type='hidden' id=txtTicker name=txtTicker value=0>
<INPUT type='hidden' id=txtLastKeyFind name=txtLastKeyFind value="">

<FORM id=frmUseful name=frmUseful style="visibility:hidden;display:none">
	<INPUT type="hidden" id=txtIEVersion name=txtIEVersion value=<%=session("IEVersion")%>>
	<INPUT type='hidden' id=txtSelectionType name=txtSelectionType value=<%=Request("selectionType")%>>
	<INPUT type='hidden' id=txtTableID name=txtTableID value=<%=Request("txtTableID")%>>
	<INPUT type='hidden' id=txtSelectedID name=txtSelectedID value=<%=Request("selectedID")%>>
</FORM>

<FORM id=frmPopup name=frmPopup style="visibility:hidden;display:none">
	<INPUT type=hidden id=Hidden1 name=txtSelectedID>
	<INPUT type=hidden id=txtSelectedName name=txtSelectedName>
	<INPUT type=hidden id=txtSelectedAccess name=txtSelectedAccess>
	<INPUT type=hidden id=txtSelectedUserName name=txtSelectedUserName>
</FORM>

</body>
</html>

    <script type="text/javascript">
        fieldrec_addhandlers();
        fieldRec_window_onload();
    </script>
