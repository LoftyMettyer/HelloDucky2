<%@ Control Language="VB" Inherits="System.Web.Mvc.ViewUserControl" %>
<%@ Import Namespace="DMI.NET" %>

<object
    classid="clsid:5220cb21-c88d-11cf-b347-00aa00a28331"
    id="Microsoft_Licensed_Class_Manager_1_0">
    <param name="LPKPath" value="lpks/main.lpk">
</object>

<object
    classid="CLSID:41021C13-8D42-4364-8388-9506F0755AE3"
    codebase="cabs/COAInt_CalRepDates.cab#version=1,0,0,2"
    id="tempCalDates"
    name="TempCalDates"
    style="VISIBILITY: hidden; WIDTH: 0px"
    width="0">
</object>

<object
    classid="CLSID:252D73AF-D7C6-4833-8539-A2C0293950B1"
    codebase="cabs/COAInt_CalRepRecord.CAB#version=1,0,0,2"
    width="0"
    name="TempCalRec"
    id="TempCalRec"
    style="LEFT: 0px; VISIBILITY: hidden; WIDTH: 0px; TOP: 0px">
    <param name="_ExtentX" value="14737">
    <param name="_ExtentY" value="714">
</object>

<object
    classid="CLSID:8E2F1EF1-3812-4678-A084-16384DE3EA6D"
    codebase="cabs/COAInt_CalRepKey.cab#version=1,0,0,2"
    id="ctlKey"
    name="ctlKey"
    width="0"
    height="0"
    style="VISIBILITY: hidden; width: 0px; height: 0px">
</object>

									
<%
	dim icount
	dim definition
	dim fok
    Dim objCalendar As HR.Intranet.Server.clsCalendarReportsRUN
	dim fNotCancelled
	dim fBadUtilDef
	dim fNoRecords
    Dim blnShowCalendar As Boolean
    Dim CalRep_UtilID
    Dim aPrompts
	
	CalRep_UtilID = Session("UtilID")
    Session("firstload") = 0

    
	fBadUtilDef = (session("utiltype") = "") or _ 
	   (session("utilname") = "") or _ 
	   (session("utilid") = "") or _ 
	   (session("action") = "")
	
	fok = not fBadUtilDef
	fNotCancelled = true
	
    objCalendar = Nothing
    Session("objCalendar" & CalRep_UtilID) = Nothing
	Session("objCalendar" & CalRep_UtilID) = ""
	
	if fOK then	
		' Create the reference to the DLL (Report Class)
        objCalendar = New HR.Intranet.Server.clsCalendarReportsRUN
        
		' Pass required info to the DLL
		objCalendar.Username = session("username")
        CallByName(objCalendar, "Connection", CallType.Let, Session("databaseConnection"))
        objCalendar.CalendarReportID = Session("utilid")
		objCalendar.ClientDateFormat = session("LocaleDateFormat")
		objCalendar.LocalDecimalSeparator = session("LocaleDecimalSeparator")
		objCalendar.SingleRecordID = Session("singleRecordID")
		
		aPrompts =  Session("Prompts_" & session("utiltype") & "_" & CalRep_UtilID)
		if fok then 
			fok = objCalendar.SetPromptedValues(aPrompts)
			fNotCancelled = Response.IsClientConnected 
			if fok then fok = fNotCancelled
		end if

		if fok then 
			fok = objCalendar.GetCalendarReportDefinition
			fNotCancelled = Response.IsClientConnected 
			if fok then fok = fNotCancelled
		end if
		
		if fok then 
			fok = objCalendar.GetEventsCollection
			fNotCancelled = Response.IsClientConnected 
			if fok then fok = fNotCancelled
		end if

		if fok then 
			fok = objCalendar.GetOrderArray
			fNotCancelled = Response.IsClientConnected 
			if fok then fok = fNotCancelled
		end if

		if fok then 
			fok = objCalendar.GenerateSQL 
			fNotCancelled = Response.IsClientConnected 
			if fok then fok = fNotCancelled
		end if

		if fok then 
			fok = objCalendar.ExecuteSql  
			fNotCancelled = Response.IsClientConnected 
			if fok then fok = fNotCancelled
		end if

		if fok then
			fok = objCalendar.Initialise_WP_Region
			fNotCancelled = Response.IsClientConnected 
			if fok then fok = fNotCancelled
		end if
		
		objCalendar.SetLastRun()

		fNoRecords = objCalendar.NoRecords 
		
		if fok then
			if Response.IsClientConnected then
				objCalendar.Cancelled = false
			else
				objCalendar.Cancelled = true
			end if
		else
			if not fNoRecords then
				if fNotCancelled then
					objCalendar.FailedMessage = objCalendar.ErrorString
					objCalendar.Failed = True
				else
					objCalendar.Cancelled = True
				end if
			end if		
		end if
		
		blnShowCalendar = (objCalendar.OutputPreview Or (objCalendar.OutputFormat = 0 And objCalendar.OutputScreen))
		
        Session("objCalendar" & CalRep_UtilID) = objCalendar
	end if
%>

<script type="text/javascript">

    function util_run_calendarreport_main_window_onload() {

        return;

        if ((txtPreview.value == 0) && (txtOK.value == "True")) {
            setGridFont(frmOutput.grdCalendarOutput);
            setGridFont(frmOutput.ssHiddenGrid);

            outputReport();
            document.getElementById('tdDisplay').innerText = 'Calendar Report Output Complete.';
            document.getElementById('Cancel').value = 'OK';
        } else {
            window.parent.document.all.item("myframeset").document.title = txtTitle.value;

            if (txtOK.value == "True") {
                try {
                    var lngPreviewWidth = new Number(760);
                    var lngPreviewHeight = new Number(540);
                    window.parent.moveTo((screen.width - lngPreviewWidth) / 2, (screen.height - lngPreviewHeight) / 2);
                    window.parent.resizeTo(lngPreviewWidth, lngPreviewHeight);
                } catch(e) {
                }
            } else {
                // Resize the popup.
                iResizeByHeight = frmPopup.offsetParent.scrollHeight - frmPopup.offsetParent.offsetHeight;
                if (frmPopup.offsetParent.offsetHeight + iResizeByHeight > screen.height) {
                    try {
                        window.parent.moveTo((screen.width - frmPopup.offsetParent.offsetWidth) / 2, 0);
                        window.parent.resizeTo(frmPopup.offsetParent.offsetWidth, screen.height);
                    } catch(e) {
                    }
                } else {
                    try {
                        window.parent.moveTo((screen.width - frmPopup.offsetParent.offsetWidth) / 2, (screen.height - (frmPopup.offsetParent.offsetHeight + iResizeByHeight)) / 2);
                        window.parent.resizeBy(0, iResizeByHeight);
                    } catch(e) {
                    }
                }

                iResizeByWidth = frmPopup.offsetParent.scrollWidth - frmPopup.offsetParent.offsetWidth;
                if (frmPopup.offsetParent.offsetWidth + iResizeByWidth > screen.width) {
                    try {
                        window.parent.moveTo(0, (screen.height - frmPopup.offsetParent.offsetHeight) / 2);
                        window.parent.resizeTo(screen.width, frmPopup.offsetParent.offsetHeight);
                    } catch(e) {
                    }
                } else {
                    try {
                        window.parent.moveTo((screen.width - (frmPopup.offsetParent.offsetWidth + iResizeByWidth)) / 2, (screen.height - frmPopup.offsetParent.offsetHeight) / 2);
                        window.parent.resizeBy(iResizeByWidth, 0);
                    } catch(e) {
                    }
                }
            }
        }
    }
</script>

<script type="text/javascript">

    function refreshDefSelAndClose()
    {
        try
        {
            window.parent.opener.window.parent.frames("menuframe").refreshDefSel();
        }
        catch(e) {}
	
        window.parent.self.close();
    }

    function replace(sExpression, sFind, sReplace)
    {
        //gi (global search, ignore case)
        var re = new RegExp(sFind,"gi");
        sExpression = sExpression.replace(re, sReplace);
        return(sExpression);
    }

    function dataOnlyPrint()
    {
        var lngPageColumnCount = 3;
        var lngActualRow = new Number(0);
        var blnSettingsDone = false;
        var sColHeading = new String('');
        var iColDataType = new Number(12);
        var iColDecimals = new Number(0);
        var blnNewPage = true;
        var lngPageCount = new Number(0);

        frmOutput.grdCalendarOutput.focus();
        frmOutput.grdCalendarOutput.caption = replace(frmOutput.grdCalendarOutput.caption,'&','&&')
	
        if (frmOutput.ssHiddenGrid.Columns.Count > 0) 
        {
            frmOutput.ssHiddenGrid.Columns.RemoveAll();
        }
		
        if (frmOutput.ssHiddenGrid.Rows > 0)
        {	
            frmOutput.ssHiddenGrid.RemoveAll();
        }
		
        frmOutput.ssHiddenGrid.Font.Name = "Verdana";
        frmOutput.ssHiddenGrid.Font.Size = 8;
        frmOutput.ssHiddenGrid.Font.Bold = false;
        frmOutput.ssHiddenGrid.Font.Underline = false;
        frmOutput.ssHiddenGrid.focus();
	
        frmOriginalDefinition.txtOptionsDone.value = 0;
	
        // Need to loop through the grid, selecting rows until we find a '*' in
        // the first column ('PageBreak').  
        frmOriginalDefinition.txtCancelPrint.value = 0;
        frmOutput.grdCalendarOutput.MoveFirst();
        for (var lngRow=0; lngRow<frmOutput.grdCalendarOutput.Rows; lngRow++)
        {
            bm = frmOutput.grdCalendarOutput.AddItemBookmark(lngRow);
	
            if (lngRow == (frmOutput.grdCalendarOutput.Rows - 1))
            {
                sBreakValue = frmOutput.grdCalendarOutput.Columns(1).CellText(bm);
                frmOutput.ssHiddenGrid.Caption = txtTitle.value + ' - ' + sBreakValue;

                // PRINT DATA
                if (frmOriginalDefinition.txtOptionsDone.value == 0) 
                {
                    frmOutput.ssHiddenGrid.PrintData(23,false,true);	
                    frmOriginalDefinition.txtOptionsDone.value = 1;
                    if (frmOriginalDefinition.txtCancelPrint.value == 1) 
                    {
                        frmOutput.grdCalendarOutput.redraw = true;
                        return;
                    }
                }
                else 
                {
                    frmOutput.ssHiddenGrid.PrintData(23,false,false);	
                }
                frmOutput.ssHiddenGrid.RemoveAll();			
                blnBreakCheck = true;
                sBreakValue = '';
                lngActualRow = 0;
            }
            else if ((frmOutput.grdCalendarOutput.Columns(0).CellText(bm) == '*') 
                  && (!blnBreakCheck))
            {
                sBreakValue = frmOutput.grdCalendarOutput.Columns(1).CellText(bm);
                frmOutput.ssHiddenGrid.Caption = txtTitle.value + ' - ' + sBreakValue;

                // PRINT DATA
                if (frmOriginalDefinition.txtOptionsDone.value == 0) 
                {
                    frmOutput.ssHiddenGrid.PrintData(23,false,true);	
                    frmOriginalDefinition.txtOptionsDone.value  = 1;
                    if (frmOriginalDefinition.txtCancelPrint.value == 1) 
                    {
                        frmOutput.grdCalendarOutput.redraw = true;
                        return;
                    }
                }
                else 
                {
                    frmOutput.ssHiddenGrid.PrintData(23,false,false);	
                }
                frmOutput.ssHiddenGrid.RemoveAll();
                frmOutput.ssHiddenGrid.Columns.RemoveAll();
			
                lngPageColumnCount = 38;
                lngPageCount++;
                blnBreakCheck = true;
                sBreakValue = '';
                lngActualRow = 0;
                blnNewPage = true;
            } 
            else if (frmOutput.grdCalendarOutput.Columns(0).CellText(bm) != '*')
            {
                if (blnNewPage)
                {
                    for (var lngCol=0; lngCol<lngPageColumnCount; lngCol++)
                    {
                        // ADD COLUMN TO GRID
                        frmOutput.ssHiddenGrid.Columns.Add(lngCol);
                        frmOutput.ssHiddenGrid.Columns(lngCol).width = 15;
                    }
                }
                blnBreakCheck = false;
                blnNewPage = false;

                // CREATE TAB DELIMITED STRING AND ADD TO GRID
                sAddItem = new String("");
                for (var lngCol=0; lngCol<lngPageColumnCount; lngCol++)
                {
                    if (frmOutput.grdCalendarOutput.Columns(lngCol).visible)
                    {
                        if (lngCol > 0) 
                        {
                            sAddItem = sAddItem + "	";
                        }
                        var sValue = new String(trim(frmOutput.grdCalendarOutput.Columns(lngCol).CellText(bm)));
                        sAddItem = sAddItem + sValue;
                        if ((sValue.length * 8) > frmOutput.ssHiddenGrid.Columns(lngCol).width)
                        {
                            frmOutput.ssHiddenGrid.Columns(lngCol).width = (sValue.length * 8);
                        }
                    }
                }
                frmOutput.ssHiddenGrid.AddItem(sAddItem);
            }
		
            if (!blnNewPage) 
            {
                lngActualRow = lngActualRow + 1; 
            }
        }
    }

    function trim(strInput)
    {
        if (strInput.length < 1)
        {
            return "";
        }
		
        while (strInput.substr(strInput.length-1, 1) == " ") 
        {
            strInput = strInput.substr(0, strInput.length - 1);
        }
	
        while (strInput.substr(0, 1) == " ") 
        {
            strInput = strInput.substr(1, strInput.length);
        }
	
        return strInput;
    }
	
    function styleArgument(psDefnString, psParameter)
    {
        var iCharIndex;
        var sDefn;
	
        sDefn = new String(psDefnString);
        psParameter = psParameter.toUpperCase(); 
	
        iCharIndex = sDefn.indexOf("	");
        if (iCharIndex >= 0) 
        {
            if (psParameter == "TYPE") return sDefn.substr(0, iCharIndex);
            sDefn = sDefn.substr(iCharIndex + 1);
            iCharIndex = sDefn.indexOf("	");
            if (iCharIndex >= 0) 
            {
                if (psParameter == "STARTCOL") return sDefn.substr(0, iCharIndex);
                sDefn = sDefn.substr(iCharIndex + 1);
                iCharIndex = sDefn.indexOf("	");
                if (iCharIndex >= 0) 
                {
                    if (psParameter == "STARTROW") return sDefn.substr(0, iCharIndex);
                    sDefn = sDefn.substr(iCharIndex + 1);
                    iCharIndex = sDefn.indexOf("	");
                    if (iCharIndex >= 0) 
                    {
                        if (psParameter == "ENDCOL") return sDefn.substr(0, iCharIndex);
                        sDefn = sDefn.substr(iCharIndex + 1);
                        iCharIndex = sDefn.indexOf("	");
                        if (iCharIndex >= 0) 
                        {
                            if (psParameter == "ENDROW") return sDefn.substr(0, iCharIndex);
                            sDefn = sDefn.substr(iCharIndex + 1);
                            iCharIndex = sDefn.indexOf("	");
                            if (iCharIndex >= 0) 
                            {
                                if (psParameter == "BACKCOLOR") return sDefn.substr(0, iCharIndex);
                                sDefn = sDefn.substr(iCharIndex + 1);
                                iCharIndex = sDefn.indexOf("	");
                                if (iCharIndex >= 0) 
                                {
                                    if (psParameter == "FORECOLOR") return sDefn.substr(0, iCharIndex);
                                    sDefn = sDefn.substr(iCharIndex + 1);
                                    iCharIndex = sDefn.indexOf("	");
                                    if (iCharIndex >= 0) 
                                    {
                                        if (psParameter == "BOLD") return sDefn.substr(0, iCharIndex);
                                        sDefn = sDefn.substr(iCharIndex + 1);
                                        iCharIndex = sDefn.indexOf("	");
                                        if (iCharIndex >= 0) 
                                        {
                                            if (psParameter == "UNDERLINE") return sDefn.substr(0, iCharIndex);
                                            sDefn = sDefn.substr(iCharIndex + 1);

                                            if (psParameter == "GRIDLINES") return sDefn;
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }
        }
	
        return "";
    }

    function mergeArgument(psDefnString, psParameter)
    {
        var iCharIndex;
        var sDefn;
	
        sDefn = new String(psDefnString);
        psParameter = psParameter.toUpperCase(); 
	
        iCharIndex = sDefn.indexOf("	");
        if (iCharIndex >= 0) 
        {
            if (psParameter == "STARTCOL") return sDefn.substr(0, iCharIndex);
            sDefn = sDefn.substr(iCharIndex + 1);
            iCharIndex = sDefn.indexOf("	");
            if (iCharIndex >= 0) 
            {
                if (psParameter == "STARTROW") return sDefn.substr(0, iCharIndex);
                sDefn = sDefn.substr(iCharIndex + 1);
                iCharIndex = sDefn.indexOf("	");
                if (iCharIndex >= 0) 
                {
                    if (psParameter == "ENDCOL") return sDefn.substr(0, iCharIndex);
                    sDefn = sDefn.substr(iCharIndex + 1);

                    if (psParameter == "ENDROW") return sDefn;
                }
            }
        }
	
        return "";	
    }

    function getDBName()
    {
        return window.parent.opener.window.parent.frames("menuframe").document.forms("frmMenuInfo").txtDatabase.value;
    }
	
    function loadAddRecords(sFrom) {
        var iCount;
	
        iCount = new Number(txtLoadCount.value);

        txtLoadCount.value = iCount + 1;

        if (iCount > 1)	
        {	
            startCalendar();
        }

    }

</script>



<%	
	if fok then
%>	
<INPUT type='hidden' id=txtLoadCount name=txtLoadCount value=0>
<input type='hidden' id=txtOK name=txtOK value="True">
<%
    Dim objUser As New HR.Intranet.Server.clsSettings
    Dim cmdEmailAddr As Object
    
    Dim arrayDefinition
    Dim arrayColumnsDefinition
    Dim arrayDataDefinition
    Dim arrayStyles
    Dim arrayMerges
    Dim INPUT_VALUE
    Dim prmEmailGroupID As Object
    Dim rstEmailAddr As Object
    Dim sErrorDescription As String
    Dim iLoop As Integer
    Dim sEmailAddresses As String
    
    Session("CalRepUtilID") = Request.Form("utilid")
    
    If blnShowCalendar Then
        Response.Write("<input type='hidden' id=txtPreview name=txtPreview value=1>" & vbCrLf)
    Else
        Response.Write("<input type='hidden' id=txtPreview name=txtPreview value=0>" & vbCrLf)
    End If
		
		if blnShowCalendar then
%>		

<div id="calendarframeset">

    <div id="dataframe" data-framesource="util_run_calendarreport_data" style="display: block;">
         <%Html.RenderPartial("~/views/home/util_run_calendarreport_data.ascx")%>
    </div>

    <div id="navframeset">
        <div id="calendarworkframe" data-framesource="util_run_calendarreport_nav" style="display: block;">
             <%Html.RenderPartial("~/views/home/util_run_calendarreport_nav.ascx")%>
        </div>
        <div id="workframefiller" data-framesource="util_run_calendarreport_nav" style="display: block;">
             <%Html.RenderPartial("~/views/home/util_run_calendarreport_navfiller.ascx")%>           
        </div>
    </div>

    <div id="calendarframe_calendar" data-framesource="util_run_calendarreport_calendar" style="display: block;">
        <%Html.RenderPartial("~/views/home/util_run_calendarreport_calendar.ascx")%>                   
    </div>

    <div id="optionsframeset">
        <div id="calendarframe_key" data-framesource="util_run_calendarreport_key" style="display: block;">
             <%Html.RenderPartial("~/views/home/util_run_calendarreport_key.ascx")%>
        </div>
        <div id="calendarframe_options" data-framesource="util_run_calendarreport_options" style="display: block;">
             <%Html.RenderPartial("~/views/home/util_run_calendarreport_options.ascx")%>
        </div>
    </div>
    <div id="calendarreport_output" data-framesource="util_run_calendarreport_output" style="display: block;"></div>
</div>


<%
		else
			'*****************************************
			'DO THE OUTPUT WITHOUT RUNNING TO PREVIEW
			'*****************************************
			if fok then
				fok = objCalendar.OutputGridDefinition 
				fNotCancelled = Response.IsClientConnected 
				if fok then fok = fNotCancelled
			end if

			if fok then 
				fok = objCalendar.OutputGridColumns 
				fNotCancelled = Response.IsClientConnected 
				if fok then fok = fNotCancelled
			end if

			if fok then 
				fok = objCalendar.OutputReport(true) 
				fNotCancelled = Response.IsClientConnected 
				if fok then fok = fNotCancelled
			end if

			if fok then

			  arrayDefinition = objCalendar.OutputArray_Definition 
				arrayColumnsDefinition = objCalendar.OutputArray_Columns 
				arrayDataDefinition = objCalendar.OutputArray_Data 
			end if	
		%>


<form id="frmOutput" name="frmOutput">
		<table align=center class="outline" cellPadding=5 cellSpacing=0> 
		    <tr>
			    <td>
					<table align=center class="invisible" cellPadding=0 cellSpacing=0> 
					    <tr>
					        <td colSpan=3 height=10></td>
					    </tr>
					    <tr>
							<td width=20></td>
					        <td align=center ID=tdDisplay>
					            Outputting Calendar Report.&nbsp;Please Wait...
							</td>
							<td width=20></td>
					    </tr>
					    <tr>
					        <td colSpan=3 height=20></td>
					    </tr>
					    <tr>
							<td width=20></td>
					        <td align=center>
						        <INPUT id=Cancel style="WIDTH: 80px" type=button width=80 value=Cancel name=Cancel class="btn"
						            onclick=window.parent.self.close(); 
	                                onmouseover="try{button_onMouseOver(this);}catch(e){}" 
	                                onmouseout="try{button_onMouseOut(this);}catch(e){}"
	                                onfocus="try{button_onFocus(this);}catch(e){}"
	                                onblur="try{button_onBlur(this);}catch(e){}" />
							</td>
							<td width=20></td>
					    </tr>
					    <tr>
					        <td colSpan=5 height=10></td>
					    </tr>
					</table>
				</td>
		    </tr>	
	    </TABLE>

	    <OBJECT classid="clsid:4A4AA697-3E6F-11D2-822F-00104B9E07A1"
		    id=grdCalendarOutput 
		    name=grdCalendarOutput 
		    codebase="cabs/COAInt_Grid.cab#version=3,1,3,6" 
		    style="HEIGHT: 0px; VISIBILITY: visible; WIDTH: 0px; display: block"
		    height="0" 
		    width="0">
<%
    For icount = 1 To UBound(arrayDefinition)
        Response.Write(arrayDefinition(icount))
    Next

    For icount = 1 To UBound(arrayColumnsDefinition)
        Response.Write(arrayColumnsDefinition(icount))
    Next
				
    For icount = 1 To UBound(arrayDataDefinition)
        Response.Write(arrayDataDefinition(icount))
    Next

    %>

     </OBJECT>
                        
            <%
				
			if fok then
        arrayStyles = objCalendar.OutputArray_Styles
				arrayMerges = objCalendar.OutputArray_Merges
			end if	
				
    '************************* START OF HIDDEN GRID ******************************
    Response.Write("<OBJECT classid=""clsid:4A4AA697-3E6F-11D2-822F-00104B9E07A1""    codebase=""cabs/COAInt_Grid.cab#version=3,1,3,6"" id=ssHiddenGrid name=ssHiddenGrid style=""visibility: visible; display: block; HEIGHT: 0px; WIDTH: 0px"" WIDTH=0 HEIGHT=0>" & vbCrLf)
    Response.Write("	<PARAM NAME=""ScrollBars"" VALUE=""4"">" & vbCrLf)
    Response.Write("	<PARAM NAME=""_Version"" VALUE=""196617"">" & vbCrLf)
    Response.Write("	<PARAM NAME=""DataMode"" VALUE=""2"">" & vbCrLf)
    Response.Write("	<PARAM NAME=""Cols"" VALUE=""1"">" & vbCrLf)
    Response.Write("	<PARAM NAME=""Rows"" VALUE=""0"">" & vbCrLf)
    Response.Write("	<PARAM NAME=""BorderStyle"" VALUE=""1"">" & vbCrLf)
    Response.Write("	<PARAM NAME=""RecordSelectors"" VALUE=""0"">" & vbCrLf)
    Response.Write("	<PARAM NAME=""GroupHeaders"" VALUE=""0"">" & vbCrLf)
    Response.Write("	<PARAM NAME=""ColumnHeaders"" VALUE=""0"">" & vbCrLf)
    Response.Write("	<PARAM NAME=""GroupHeadLines"" VALUE=""0"">" & vbCrLf)
    Response.Write("	<PARAM NAME=""HeadLines"" VALUE=""0"">" & vbCrLf)
    Response.Write("	<PARAM NAME=""FieldDelimiter"" VALUE=""(None)"">" & vbCrLf)
    Response.Write("	<PARAM NAME=""FieldSeparator"" VALUE=""(Tab)"">" & vbCrLf)
    Response.Write("	<PARAM NAME=""Row.Count"" VALUE=""0"">" & vbCrLf)
    Response.Write("	<PARAM NAME=""Col.Count"" VALUE=""1"">" & vbCrLf)
    Response.Write("	<PARAM NAME=""stylesets.count"" VALUE=""0"">" & vbCrLf)
    Response.Write("	<PARAM NAME=""TagVariant"" VALUE=""EMPTY"">" & vbCrLf)
    Response.Write("	<PARAM NAME=""UseGroups"" VALUE=""0"">" & vbCrLf)
    Response.Write("	<PARAM NAME=""HeadFont3D"" VALUE=""0"">" & vbCrLf)
    Response.Write("	<PARAM NAME=""Font3D"" VALUE=""0"">" & vbCrLf)
    Response.Write("	<PARAM NAME=""DividerType"" VALUE=""0"">" & vbCrLf)
    Response.Write("	<PARAM NAME=""DividerStyle"" VALUE=""1"">" & vbCrLf)
    Response.Write("	<PARAM NAME=""DefColWidth"" VALUE=""0"">" & vbCrLf)
    Response.Write("	<PARAM NAME=""BeveColorScheme"" VALUE=""2"">" & vbCrLf)
    Response.Write("	<PARAM NAME=""BevelColorFrame"" VALUE=""-2147483642"">" & vbCrLf)
    Response.Write("	<PARAM NAME=""BevelColorHighlight"" VALUE=""-2147483628"">" & vbCrLf)
    Response.Write("	<PARAM NAME=""BevelColorShadow"" VALUE=""-2147483632"">" & vbCrLf)
    Response.Write("	<PARAM NAME=""BevelColorFace"" VALUE=""-2147483633"">" & vbCrLf)
    Response.Write("	<PARAM NAME=""CheckBox3D"" VALUE=""-1"">" & vbCrLf)
    Response.Write("	<PARAM NAME=""AllowAddNew"" VALUE=""0"">" & vbCrLf)
    Response.Write("	<PARAM NAME=""AllowDelete"" VALUE=""0"">" & vbCrLf)
    Response.Write("	<PARAM NAME=""AllowUpdate"" VALUE=""0"">" & vbCrLf)
    Response.Write("	<PARAM NAME=""MultiLine"" VALUE=""0"">" & vbCrLf)
    Response.Write("	<PARAM NAME=""ActiveCellStyleSet"" VALUE="""">" & vbCrLf)
    Response.Write("	<PARAM NAME=""RowSelectionStyle"" VALUE=""0"">" & vbCrLf)
    Response.Write("	<PARAM NAME=""AllowRowSizing"" VALUE=""0"">" & vbCrLf)
    Response.Write("	<PARAM NAME=""AllowGroupSizing"" VALUE=""0"">" & vbCrLf)
    Response.Write("	<PARAM NAME=""AllowColumnSizing"" VALUE=""-1"">" & vbCrLf)
    Response.Write("	<PARAM NAME=""AllowGroupMoving"" VALUE=""0"">" & vbCrLf)
    Response.Write("	<PARAM NAME=""AllowColumnMoving"" VALUE=""0"">" & vbCrLf)
    Response.Write("	<PARAM NAME=""AllowGroupSwapping"" VALUE=""0"">" & vbCrLf)
    Response.Write("	<PARAM NAME=""AllowColumnSwapping"" VALUE=""0"">" & vbCrLf)
    Response.Write("	<PARAM NAME=""AllowGroupShrinking"" VALUE=""0"">" & vbCrLf)
    Response.Write("	<PARAM NAME=""AllowColumnShrinking"" VALUE=""0"">" & vbCrLf)
    Response.Write("	<PARAM NAME=""AllowDragDrop"" VALUE=""0"">" & vbCrLf)
    Response.Write("	<PARAM NAME=""UseExactRowCount"" VALUE=""-1"">" & vbCrLf)
    Response.Write("	<PARAM NAME=""SelectTypeCol"" VALUE=""0"">" & vbCrLf)
    Response.Write("	<PARAM NAME=""SelectTypeRow"" VALUE=""2"">" & vbCrLf)
    Response.Write("	<PARAM NAME=""SelectByCell"" VALUE=""0"">" & vbCrLf)
    Response.Write("	<PARAM NAME=""BalloonHelp"" VALUE=""0"">" & vbCrLf)
    Response.Write("	<PARAM NAME=""RowNavigation"" VALUE=""1"">" & vbCrLf)
    Response.Write("	<PARAM NAME=""CellNavigation"" VALUE=""0"">" & vbCrLf)
    Response.Write("	<PARAM NAME=""MaxSelectedRows"" VALUE=""0"">" & vbCrLf)
    Response.Write("	<PARAM NAME=""HeadStyleSet"" VALUE="""">" & vbCrLf)
    Response.Write("	<PARAM NAME=""StyleSet"" VALUE="""">" & vbCrLf)
    Response.Write("	<PARAM NAME=""ForeColorEven"" VALUE=""0"">" & vbCrLf)
    Response.Write("	<PARAM NAME=""ForeColorOdd"" VALUE=""0"">" & vbCrLf)
    Response.Write("	<PARAM NAME=""BackColorEven"" VALUE=""16777215"">" & vbCrLf)
    Response.Write("	<PARAM NAME=""BackColorOdd"" VALUE=""16777215"">" & vbCrLf)
    Response.Write("	<PARAM NAME=""Levels"" VALUE=""1"">" & vbCrLf)
    Response.Write("	<PARAM NAME=""RowHeight"" VALUE=""238"">" & vbCrLf)
    Response.Write("	<PARAM NAME=""ExtraHeight"" VALUE=""0"">" & vbCrLf)
    Response.Write("	<PARAM NAME=""ActiveRowStyleSet"" VALUE="""">" & vbCrLf)
    Response.Write("	<PARAM NAME=""CaptionAlignment"" VALUE=""2"">" & vbCrLf)
    Response.Write("	<PARAM NAME=""SplitterPos"" VALUE=""0"">" & vbCrLf)
    Response.Write("	<PARAM NAME=""SplitterVisible"" VALUE=""0"">" & vbCrLf)
    Response.Write("	<PARAM NAME=""Columns.Count"" VALUE=""1"">" & vbCrLf)
    Response.Write("	<PARAM NAME=""Columns(0).Width"" VALUE=""1000"">" & vbCrLf)
    Response.Write("	<PARAM NAME=""Columns(0).Visible"" VALUE=""0"">" & vbCrLf)
    Response.Write("	<PARAM NAME=""Columns(0).Columns.Count"" VALUE=""1"">" & vbCrLf)
    Response.Write("	<PARAM NAME=""Columns(0).Caption"" VALUE=""PageBreak"">" & vbCrLf)
    Response.Write("	<PARAM NAME=""Columns(0).Name"" VALUE=""PageBreak"">" & vbCrLf)
    Response.Write("	<PARAM NAME=""Columns(0).Alignment"" VALUE=""0"">" & vbCrLf)
    Response.Write("	<PARAM NAME=""Columns(0).CaptionAlignment"" VALUE=""2"">" & vbCrLf)
    Response.Write("	<PARAM NAME=""Columns(0).Bound"" VALUE=""0"">" & vbCrLf)
    Response.Write("	<PARAM NAME=""Columns(0).AllowSizing"" VALUE=""1"">" & vbCrLf)
    Response.Write("	<PARAM NAME=""Columns(0).DataField"" VALUE=""Column 0"">" & vbCrLf)
    Response.Write("	<PARAM NAME=""Columns(0).DataType"" VALUE=""8"">" & vbCrLf)
    Response.Write("	<PARAM NAME=""Columns(0).Level"" VALUE=""0"">" & vbCrLf)
    Response.Write("	<PARAM NAME=""Columns(0).NumberFormat"" VALUE="""">" & vbCrLf)
    Response.Write("	<PARAM NAME=""Columns(0).Case"" VALUE=""0"">" & vbCrLf)
    Response.Write("	<PARAM NAME=""Columns(0).FieldLen"" VALUE=""4096"">" & vbCrLf)
    Response.Write("	<PARAM NAME=""Columns(0).VertScrollBar"" VALUE=""0"">" & vbCrLf)
    Response.Write("	<PARAM NAME=""Columns(0).Locked"" VALUE=""0"">" & vbCrLf)
    Response.Write("	<PARAM NAME=""Columns(0).Style"" VALUE=""0"">" & vbCrLf)
    Response.Write("	<PARAM NAME=""Columns(0).ButtonsAlways"" VALUE=""0"">" & vbCrLf)
    Response.Write("	<PARAM NAME=""Columns(0).RowCount"" VALUE=""0"">" & vbCrLf)
    Response.Write("	<PARAM NAME=""Columns(0).ColCount"" VALUE=""1"">" & vbCrLf)
    Response.Write("	<PARAM NAME=""Columns(0).HasHeadForeColor"" VALUE=""0"">" & vbCrLf)
    Response.Write("	<PARAM NAME=""Columns(0).HasHeadBackColor"" VALUE=""0"">" & vbCrLf)
    Response.Write("	<PARAM NAME=""Columns(0).HasForeColor"" VALUE=""0"">" & vbCrLf)
    Response.Write("	<PARAM NAME=""Columns(0).HasBackColor"" VALUE=""0"">" & vbCrLf)
    Response.Write("	<PARAM NAME=""Columns(0).HeadForeColor"" VALUE=""0"">" & vbCrLf)
    Response.Write("	<PARAM NAME=""Columns(0).HeadBackColor"" VALUE=""0"">" & vbCrLf)
    Response.Write("	<PARAM NAME=""Columns(0).ForeColor"" VALUE=""0"">" & vbCrLf)
    Response.Write("	<PARAM NAME=""Columns(0).BackColor"" VALUE=""0"">" & vbCrLf)
    Response.Write("	<PARAM NAME=""Columns(0).HeadStyleSet"" VALUE="""">" & vbCrLf)
    Response.Write("	<PARAM NAME=""Columns(0).StyleSet"" VALUE="""">" & vbCrLf)
    Response.Write("	<PARAM NAME=""Columns(0).Nullable"" VALUE=""1"">" & vbCrLf)
    Response.Write("	<PARAM NAME=""Columns(0).Mask"" VALUE="""">" & vbCrLf)
    Response.Write("	<PARAM NAME=""Columns(0).PromptInclude"" VALUE=""0"">" & vbCrLf)
    Response.Write("	<PARAM NAME=""Columns(0).ClipMode"" VALUE=""0"">" & vbCrLf)
    Response.Write("	<PARAM NAME=""Columns(0).PromptChar"" VALUE=""95"">" & vbCrLf)
    Response.Write("	<PARAM NAME=""UseDefaults"" VALUE=""-1"">" & vbCrLf)
    Response.Write("	<PARAM NAME=""TabNavigation"" VALUE=""1"">" & vbCrLf)
    Response.Write("	<PARAM NAME=""BatchUpdate"" VALUE=""0"">" & vbCrLf)
    Response.Write("	<PARAM NAME=""_ExtentX"" VALUE=""0"">" & vbCrLf)
    Response.Write("	<PARAM NAME=""_ExtentY"" VALUE=""0"">" & vbCrLf)
    Response.Write("	<PARAM NAME=""_StockProps"" VALUE=""79"">" & vbCrLf)
    Response.Write("	<PARAM NAME=""Caption"" VALUE="""">" & vbCrLf)
    Response.Write("	<PARAM NAME=""ForeColor"" VALUE=""0"">" & vbCrLf)
    Response.Write("	<PARAM NAME=""BackColor"" VALUE=""16777215"">" & vbCrLf)
    Response.Write("	<PARAM NAME=""Enabled"" VALUE=""-1"">" & vbCrLf)
    Response.Write("	<PARAM NAME=""DataMember"" VALUE="""">" & vbCrLf)

                Response.Write("</OBJECT>" & vbCrLf)

                '***************************** END OF HIDDEN GRID **************************************				
                Response.Write("<INPUT type='hidden' id=txtCalendarPageCount name=txtCalendarPageCount value=" & UBound(arrayMerges) & ">" & vbCrLf)
    
    %>

    </form>
            
    <%

        Session("firstload") = 1
			
			if fok then 
				on error resume next
				
				dim iPage
				dim iStyle
				dim iMerge
				dim arrayPageStyles
				dim arrayPageMerges
			
        For iPage = 0 To UBound(arrayMerges)
            arrayPageMerges = arrayMerges(iPage)
            Response.Write("<form id=frmCalendarMerge_" & iPage & " name=frmCalendarMerge_" & iPage & " style=""visibility:hidden;display:none"">" & vbCrLf)
            For iMerge = 0 To UBound(arrayPageMerges)
                INPUT_VALUE = arrayPageMerges(iMerge)
                Response.Write("	<INPUT type=hidden name=Merge_" & iPage & "_" & iMerge & " ID=Merge_" & iPage & "_" & iMerge & " VALUE=""" & INPUT_VALUE & """>" & vbCrLf)
            Next
            Response.Write("</form>" & vbCrLf)
        Next

        For iPage = 0 To UBound(arrayStyles)
            arrayPageStyles = arrayStyles(iPage)
            Response.Write("<form id=frmCalendarStyle_" & iPage & " name=frmCalendarStyle_" & iPage & " style=""visibility:hidden;display:none"">" & vbCrLf)
            For iStyle = 0 To UBound(arrayPageStyles)
                INPUT_VALUE = arrayPageStyles(iStyle)
                Response.Write("	<INPUT type=hidden name=Style_" & iPage & "_" & iStyle & " ID=Style_" & iPage & "_" & iStyle & " VALUE=""" & INPUT_VALUE & """>" & vbCrLf)
            Next
            Response.Write("</form>" & vbCrLf)
        Next
			end if
			
			if fok then
				objCalendar.OutputArray_Clear
			end if

    'Write the function that Outputs the report to the Output Classes in the Client DLL.
    Response.Write("<script type=""text/javascript"">" & vbCrLf)

    Response.Write("function outputReport() " & vbCrLf)
    Response.Write("	{" & vbCrLf & vbCrLf)
	
    Response.Write("	var lngPageColumnCount = 3;" & vbCrLf)
    Response.Write("  var lngActualRow = new Number(0);" & vbCrLf)
    Response.Write("  var blnSettingsDone = false;" & vbCrLf)
    Response.Write("	var sColHeading = new String(''); " & vbCrLf)
    Response.Write("	var iColDataType = new Number(12); " & vbCrLf)
    Response.Write("	var iColDecimals = new Number(0); " & vbCrLf)
    Response.Write("  var blnNewPage = false;" & vbCrLf)
    Response.Write("  var lngPageCount = new Number(0);" & vbCrLf)

    Response.Write("  var strType = new String('');" & vbCrLf)
    Response.Write("  var lngStartCol = new Number(0);" & vbCrLf)
    Response.Write("  var lngStartRow = new Number(0);" & vbCrLf)
    Response.Write("  var lngEndCol = new Number(0);" & vbCrLf)
    Response.Write("  var lngEndRow = new Number(0);" & vbCrLf)
    Response.Write("  var lngBackCol = new Number(0);" & vbCrLf)
    Response.Write("  var lngForeCol = new Number(0);" & vbCrLf)
    Response.Write("  var blnBold = false;" & vbCrLf)
    Response.Write("  var blnUnderline = false;" & vbCrLf)
    Response.Write("  var blnGridlines = false;" & vbCrLf)
	
        objUser.Connection = Session("databaseConnection")
			
    Response.Write("  window.parent.ASRIntranetOutput.UserName = """ & CleanStringForJavaScript(Session("Username")) & """;" & vbCrLf)
    Response.Write("  window.parent.ASRIntranetOutput.SaveAsValues = """ & CleanStringForJavaScript(Session("OfficeSaveAsValues")) & """;" & vbCrLf)

    Response.Write("  frmMenuFrame = window.parent.parent.opener.window.parent.frames(""menuframe"");" & vbCrLf)
		
    Response.Write("	window.parent.ASRIntranetOutput.SettingOptions(")
    Response.Write("""" & CleanStringForJavaScript(objUser.GetUserSetting("Output", "WordTemplate", "")) & """, ")
    Response.Write("""" & CleanStringForJavaScript(objUser.GetUserSetting("Output", "ExcelTemplate", "")) & """, ")

    If (objUser.GetUserSetting("Output", "ExcelGridlines", "0") = "1") Then
        Response.Write("true, ")
    Else
        Response.Write("false, ")
    End If

    If (objUser.GetUserSetting("Output", "ExcelHeaders", "0") = "1") Then
        Response.Write("true, ")
    Else
        Response.Write("false, ")
    End If

    If (objUser.GetUserSetting("Output", "AutoFitCols", "1") = "1") Then
        Response.Write("true, ")
    Else
        Response.Write("false, ")
    End If

    If (objUser.GetUserSetting("Output", "Landscape", "1") = "1") Then
        Response.Write("true, " & vbCrLf)
    Else
        Response.Write("false, " & vbCrLf)
    End If
			   
    Response.Write("frmMenuFrame.document.all.item(""txtSysPerm_EMAILGROUPS_VIEW"").value);" & vbCrLf)

    Response.Write("  window.parent.ASRIntranetOutput.SettingLocations(")
    Response.Write(CleanStringForJavaScript(objUser.GetUserSetting("Output", "TitleCol", "3")) & ", ")
    Response.Write(CleanStringForJavaScript(objUser.GetUserSetting("Output", "TitleRow", "2")) & ", ")
    Response.Write(CleanStringForJavaScript(objUser.GetUserSetting("Output", "DataCol", "2")) & ", ")
    Response.Write(CleanStringForJavaScript(objUser.GetUserSetting("Output", "DataRow", "4")) & ");" & vbCrLf)

    Response.Write("  window.parent.ASRIntranetOutput.SettingTitle(")
    If (objUser.GetUserSetting("Output", "TitleGridLines", "0") = "1") Then
        Response.Write("true, ")
    Else
        Response.Write("false, ")
    End If

    If (objUser.GetUserSetting("Output", "TitleBold", "1") = "1") Then
        Response.Write("true, ")
    Else
        Response.Write("false, ")
    End If

    If (objUser.GetUserSetting("Output", "TitleUnderline", "0") = "1") Then
        Response.Write("true, ")
    Else
        Response.Write("false, ")
    End If

    Response.Write(CleanStringForJavaScript(objUser.GetUserSetting("Output", "TitleBackcolour", "16777215")) & ", ")
    Response.Write(CleanStringForJavaScript(objUser.GetUserSetting("Output", "TitleForecolour", "6697779")) & ", ")
    Response.Write(CleanStringForJavaScript(objUser.GetWordColourIndex(objUser.GetUserSetting("Output", "TitleBackcolour", "16777215"))) & ", ")
    Response.Write(CleanStringForJavaScript(objUser.GetWordColourIndex(objUser.GetUserSetting("Output", "TitleForecolour", "6697779"))) & ");" & vbCrLf)

    Response.Write("window.parent.ASRIntranetOutput.SettingHeading(")
    Response.Write(CleanStringForJavaScript(objUser.GetUserSetting("Output", "HeadingGridLines", "1")) & ", ")
    Response.Write(CleanStringForJavaScript(objUser.GetUserSetting("Output", "HeadingBold", "1")) & ", ")
    Response.Write(CleanStringForJavaScript(objUser.GetUserSetting("Output", "HeadingUnderline", "0")) & ", ")
    Response.Write(CleanStringForJavaScript(objUser.GetUserSetting("Output", "HeadingBackcolour", "16248553")) & ", ")
    Response.Write(CleanStringForJavaScript(objUser.GetUserSetting("Output", "HeadingForecolour", "6697779")) & ", ")
    Response.Write(CleanStringForJavaScript(objUser.GetWordColourIndex(objUser.GetUserSetting("Output", "HeadingBackcolour", "16248553"))) & ", ")
    Response.Write(CleanStringForJavaScript(objUser.GetWordColourIndex(objUser.GetUserSetting("Output", "HeadingForecolour", "6697779"))) & ");" & vbCrLf)

    Response.Write("window.parent.ASRIntranetOutput.SettingData(")
    Response.Write(CleanStringForJavaScript(objUser.GetUserSetting("Output", "DataGridLines", "1")) & ", ")
    Response.Write(CleanStringForJavaScript(objUser.GetUserSetting("Output", "DataBold", "0")) & ", ")
    Response.Write(CleanStringForJavaScript(objUser.GetUserSetting("Output", "DataUnderline", "0")) & ", ")
    Response.Write(CleanStringForJavaScript(objUser.GetUserSetting("Output", "DataBackcolour", "15988214")) & ", ")
    Response.Write(CleanStringForJavaScript(objUser.GetUserSetting("Output", "DataForecolour", "6697779")) & ", ")
    Response.Write(CleanStringForJavaScript(objUser.GetWordColourIndex(objUser.GetUserSetting("Output", "DataBackcolour", "15988214"))) & ", ")
    Response.Write(CleanStringForJavaScript(objUser.GetWordColourIndex(objUser.GetUserSetting("Output", "DataForecolour", "6697779"))) & ");" & vbCrLf)
			
    Dim lngFormat
    Dim blnScreen
    Dim blnPrinter
    Dim strPrinterName
    Dim blnSave
    Dim lngSaveExisting
    Dim blnEmail
    Dim lngEmailGroupID
    Dim strEmailSubject
    Dim strEmailAttachAs
    Dim strFileName As String
    Dim sCloseFunction As String
			
			lngFormat = cleanStringForJavaScript(objCalendar.OutputFormat)
			blnScreen = cleanStringForJavaScript(LCase(objCalendar.OutputScreen))
			blnPrinter = cleanStringForJavaScript(LCase(objCalendar.OutputPrinter))
			strPrinterName = cleanStringForJavaScript(objCalendar.OutputPrinterName) 
			blnSave = cleanStringForJavaScript(LCase(objCalendar.OutputSave))
			lngSaveExisting = cleanStringForJavaScript(objCalendar.OutputSaveExisting)
			blnEmail = cleanStringForJavaScript(LCase(objCalendar.OutputEmail))
			lngEmailGroupID = CLng(objCalendar.OutputEmailID)
			strEmailSubject = cleanStringForJavaScript(objCalendar.OutputEmailSubject)
			strEmailAttachAs = cleanStringForJavaScript(objCalendar.OutputEmailAttachAs)
			strFileName = cleanStringForJavaScript(objCalendar.OutputFilename)

			if (blnEmail) and (lngEmailGroupID > 0) then
			
        cmdEmailAddr = CreateObject("ADODB.Command")
        cmdEmailAddr.CommandText = "spASRIntGetEmailGroupAddresses"
        cmdEmailAddr.CommandType = 4 ' Stored procedure
        cmdEmailAddr.ActiveConnection = Session("databaseConnection")

        prmEmailGroupID = cmdEmailAddr.CreateParameter("EmailGroupID", 3, 1) ' 3=integer, 1=input
        cmdEmailAddr.Parameters.Append(prmEmailGroupID)
        prmEmailGroupID.value = CleanNumeric(lngEmailGroupID)

        Err.Clear()
        rstEmailAddr = cmdEmailAddr.Execute

        If (Err.Number <> 0) Then
            sErrorDescription = "Error getting the email addresses for group." & vbCrLf & FormatError(Err.Description)
        End If

        If Len(sErrorDescription) = 0 Then
            iLoop = 1
            Do While Not rstEmailAddr.EOF
                If iLoop > 1 Then
                    sEmailAddresses = sEmailAddresses & ";"
                End If
                sEmailAddresses = sEmailAddresses & rstEmailAddr.Fields("Fixed").Value
                rstEmailAddr.MoveNext()
                iLoop = iLoop + 1
            Loop
					
            ' Release the ADO recordset object.
            rstEmailAddr.close()
        End If
						
        rstEmailAddr = Nothing
        cmdEmailAddr = Nothing
    End If
			
    Response.Write("fok = window.parent.ASRIntranetOutput.SetOptions(false, " & _
                                            lngFormat & "," & blnScreen & ", " & _
                                            blnPrinter & ",""" & strPrinterName & """, " & _
                                            blnSave & "," & lngSaveExisting & ", " & _
                                            blnEmail & ", """ & CleanStringForJavaScript(sEmailAddresses) & """, """ & _
                                            strEmailSubject & """,""" & strEmailAttachAs & """,""" & strFileName & """);" & vbCrLf)
			
    Response.Write("if (fok == true) {" & vbCrLf)
    If (objCalendar.OutputFormat = 0) And (objCalendar.OutputPrinter) Then
        Response.Write("	window.parent.ASRIntranetOutput.SetPrinter();" & vbCrLf)
        Response.Write("  dataOnlyPrint();" & vbCrLf)
        Response.Write("	window.parent.ASRIntranetOutput.ResetDefaultPrinter();" & vbCrLf)
    Else
        Response.Write("if (window.parent.ASRIntranetOutput.GetFile() == true) " & vbCrLf)
        Response.Write("	{" & vbCrLf)
        Response.Write("	window.parent.ASRIntranetOutput.InitialiseStyles();" & vbCrLf)
        Response.Write("	window.parent.ASRIntranetOutput.ResetStyles();" & vbCrLf)
        Response.Write("	window.parent.ASRIntranetOutput.ResetColumns();" & vbCrLf)
        Response.Write("	window.parent.ASRIntranetOutput.ResetMerges();" & vbCrLf)

        Response.Write("	window.parent.ASRIntranetOutput.HeaderRows = 1;" & vbCrLf)
        Response.Write("	window.parent.ASRIntranetOutput.HeaderCols = 0;" & vbCrLf)
        Response.Write("	window.parent.ASRIntranetOutput.SizeColumnsIndependently = true;" & vbCrLf)
	
        Response.Write("	window.parent.ASRIntranetOutput.ArrayDim((lngPageColumnCount-1), 0);" & vbCrLf & vbCrLf)
        Response.Write("  frmOutput.grdCalendarOutput.focus();")

        Response.Write("  frmOutput.grdCalendarOutput.MoveFirst();" & vbCrLf)
        Response.Write("  for (var lngRow=0; lngRow<frmOutput.grdCalendarOutput.Rows; lngRow++)" & vbCrLf)
        Response.Write("		{" & vbCrLf)
        Response.Write("		bm = frmOutput.grdCalendarOutput.AddItemBookmark(lngRow);" & vbCrLf)
	
	
        Response.Write("		if (lngRow == (frmOutput.grdCalendarOutput.Rows - 1))" & vbCrLf)
        Response.Write("			{" & vbCrLf)
        Response.Write("			sBreakValue = frmOutput.grdCalendarOutput.Columns(1).CellText(bm);" & vbCrLf)
        Response.Write("			if ((sBreakValue == 'Key') && (" & lngFormat & " != 4)) " & vbCrLf)
        Response.Write("				{ " & vbCrLf)
        Response.Write("				window.parent.ASRIntranetOutput.AddPage(replace(frmOutput.grdCalendarOutput.Caption,'&&','&') ,sBreakValue);" & vbCrLf)
        Response.Write("				} " & vbCrLf)
        Response.Write("			else " & vbCrLf)
        Response.Write("				{ " & vbCrLf)
        Response.Write("				window.parent.ASRIntranetOutput.AddPage(replace(frmOutput.grdCalendarOutput.Caption,'&&','&') + ' - ' + sBreakValue,sBreakValue);" & vbCrLf)
        Response.Write("				} " & vbCrLf)

        Response.Write("			var frmMerge = document.forms('frmCalendarMerge_'+lngPageCount);" & vbCrLf)
	
        Response.Write("			var dataCollection = frmMerge.elements;" & vbCrLf)
        Response.Write("			if (dataCollection!=null) " & vbCrLf)
        Response.Write("				{" & vbCrLf)
        Response.Write("				for (i=0; i<dataCollection.length; i++)  " & vbCrLf)
        Response.Write("					{" & vbCrLf)
        Response.Write("					strMergeString = dataCollection.item(i).value;" & vbCrLf)
        Response.Write("					if (strMergeString != '')" & vbCrLf)
        Response.Write("						{" & vbCrLf)
        Response.Write("						lngStartCol = Number(mergeArgument(strMergeString,'STARTCOL'));" & vbCrLf)
        Response.Write("						lngStartRow = Number(mergeArgument(strMergeString,'STARTROW'));" & vbCrLf)
        Response.Write("						lngEndCol = Number(mergeArgument(strMergeString,'ENDCOL'));" & vbCrLf)
        Response.Write("						lngEndRow = Number(mergeArgument(strMergeString,'ENDROW'));" & vbCrLf)
        Response.Write("						window.parent.ASRIntranetOutput.AddMerge(lngStartCol,lngStartRow,lngEndCol,lngEndRow);" & vbCrLf)
        Response.Write("						}" & vbCrLf)
        Response.Write("					}" & vbCrLf)
        Response.Write("				}" & vbCrLf)

        Response.Write("			var frmStyle = document.forms('frmCalendarStyle_'+lngPageCount);" & vbCrLf)
        Response.Write("			var dataCollection = frmStyle.elements;" & vbCrLf)
        Response.Write("			if (dataCollection!=null) " & vbCrLf)
        Response.Write("				{" & vbCrLf)
        Response.Write("				for (i=0; i<dataCollection.length; i++)  " & vbCrLf)
        Response.Write("					{" & vbCrLf)
        Response.Write("					strStyleString = dataCollection.item(i).value;" & vbCrLf)
        Response.Write("					if (strStyleString != '')" & vbCrLf)
        Response.Write("						{" & vbCrLf)
        Response.Write("						strType = styleArgument(strStyleString,'TYPE');" & vbCrLf)
        Response.Write("						lngStartCol = Number(styleArgument(strStyleString,'STARTCOL'));" & vbCrLf)
        Response.Write("						lngStartRow = Number(styleArgument(strStyleString,'STARTROW'));" & vbCrLf)
        Response.Write("						lngEndCol = Number(styleArgument(strStyleString,'ENDCOL'));" & vbCrLf)
        Response.Write("						lngEndRow = Number(styleArgument(strStyleString,'ENDROW'));" & vbCrLf)
        Response.Write("						lngBackCol = Number(styleArgument(strStyleString,'BACKCOLOR'));" & vbCrLf)
        Response.Write("						lngForeCol = Number(styleArgument(strStyleString,'FORECOLOR'));" & vbCrLf)
        Response.Write("						blnBold = styleArgument(strStyleString,'BOLD');" & vbCrLf)
        Response.Write("						blnUnderline = styleArgument(strStyleString,'UNDERLINE');" & vbCrLf)
        Response.Write("						blnGridlines = styleArgument(strStyleString,'GRIDLINES');" & vbCrLf)
        Response.Write("						window.parent.ASRIntranetOutput.AddStyle(strType,lngStartCol,lngStartRow,lngEndCol,lngEndRow,lngBackCol,lngForeCol,blnBold,blnUnderline,blnGridlines);" & vbCrLf)
        Response.Write("						}" & vbCrLf)
        Response.Write("					}" & vbCrLf)
        Response.Write("				}" & vbCrLf)
	
        Response.Write("			for (var lngCol=0; lngCol<lngPageColumnCount; lngCol++)" & vbCrLf)
        Response.Write("				{" & vbCrLf)
        Response.Write("				window.parent.ASRIntranetOutput.AddColumn(sColHeading, iColDataType, iColDecimals, false);" & vbCrLf)
        Response.Write("				}" & vbCrLf)
        Response.Write("			window.parent.ASRIntranetOutput.DataArray();" & vbCrLf)
        Response.Write("			blnBreakCheck = true;" & vbCrLf)
        Response.Write("			sBreakValue = '';" & vbCrLf)
        Response.Write("			lngActualRow = 0;" & vbCrLf)
        Response.Write("			}" & vbCrLf)
	
        Response.Write("    else if ((frmOutput.grdCalendarOutput.Columns(0).CellText(bm) == '*')" & vbCrLf)
        Response.Write("					&& (!blnBreakCheck))" & vbCrLf)
        Response.Write("			{" & vbCrLf)
        Response.Write("			sBreakValue = frmOutput.grdCalendarOutput.Columns(1).CellText(bm);" & vbCrLf)
        Response.Write("			window.parent.ASRIntranetOutput.AddPage(replace(frmOutput.grdCalendarOutput.Caption,'&&','&') + ' - ' + sBreakValue,sBreakValue);" & vbCrLf)
	
        Response.Write("			var frmMerge = document.forms('frmCalendarMerge_'+lngPageCount);" & vbCrLf)
        Response.Write("			var dataCollection = frmMerge.elements;" & vbCrLf)
        Response.Write("			if (dataCollection!=null) " & vbCrLf)
        Response.Write("				{" & vbCrLf)
        Response.Write("				for (i=0; i<dataCollection.length; i++)  " & vbCrLf)
        Response.Write("					{" & vbCrLf)
        Response.Write("					strMergeString = dataCollection.item(i).value;" & vbCrLf)
        Response.Write("					if (strMergeString != '')" & vbCrLf)
        Response.Write("						{" & vbCrLf)
        Response.Write("						lngStartCol = Number(mergeArgument(strMergeString,'STARTCOL'));" & vbCrLf)
        Response.Write("						lngStartRow = Number(mergeArgument(strMergeString,'STARTROW'));" & vbCrLf)
        Response.Write("						lngEndCol = Number(mergeArgument(strMergeString,'ENDCOL'));" & vbCrLf)
        Response.Write("						lngEndRow = Number(mergeArgument(strMergeString,'ENDROW'));" & vbCrLf)
        Response.Write("						window.parent.ASRIntranetOutput.AddMerge(lngStartCol,lngStartRow,lngEndCol,lngEndRow);" & vbCrLf)
        Response.Write("						}" & vbCrLf)
        Response.Write("					}" & vbCrLf)
        Response.Write("				}" & vbCrLf)

        Response.Write("			var frmStyle = document.forms('frmCalendarStyle_'+lngPageCount);" & vbCrLf)
        Response.Write("			var dataCollection = frmStyle.elements;" & vbCrLf)
        Response.Write("			if (dataCollection!=null) " & vbCrLf)
        Response.Write("				{" & vbCrLf)
        Response.Write("				for (i=0; i<dataCollection.length; i++)  " & vbCrLf)
        Response.Write("					{" & vbCrLf)
        Response.Write("					strStyleString = dataCollection.item(i).value;" & vbCrLf)
        Response.Write("					if (strStyleString != '')" & vbCrLf)
        Response.Write("						{" & vbCrLf)
        Response.Write("						strType = styleArgument(strStyleString,'TYPE');" & vbCrLf)
        Response.Write("						lngStartCol = Number(styleArgument(strStyleString,'STARTCOL'));" & vbCrLf)
        Response.Write("						lngStartRow = Number(styleArgument(strStyleString,'STARTROW'));" & vbCrLf)
        Response.Write("						lngEndCol = Number(styleArgument(strStyleString,'ENDCOL'));" & vbCrLf)
        Response.Write("						lngEndRow = Number(styleArgument(strStyleString,'ENDROW'));" & vbCrLf)
        Response.Write("						lngBackCol = Number(styleArgument(strStyleString,'BACKCOLOR'));" & vbCrLf)
        Response.Write("						lngForeCol = Number(styleArgument(strStyleString,'FORECOLOR'));" & vbCrLf)
        Response.Write("						blnBold = styleArgument(strStyleString,'BOLD');" & vbCrLf)
        Response.Write("						blnUnderline = styleArgument(strStyleString,'UNDERLINE');" & vbCrLf)
        Response.Write("						blnGridlines = styleArgument(strStyleString,'GRIDLINES');" & vbCrLf)
        Response.Write("						window.parent.ASRIntranetOutput.AddStyle(strType,lngStartCol,lngStartRow,lngEndCol,lngEndRow,lngBackCol,lngForeCol,blnBold,blnUnderline,blnGridlines);" & vbCrLf)
        Response.Write("						}" & vbCrLf)
        Response.Write("					}" & vbCrLf)
        Response.Write("				}" & vbCrLf)
	
        Response.Write("			for (var lngCol=0; lngCol<lngPageColumnCount; lngCol++)" & vbCrLf)
        Response.Write("				{" & vbCrLf)
        Response.Write("				window.parent.ASRIntranetOutput.AddColumn(sColHeading, iColDataType, iColDecimals, false);" & vbCrLf)
        Response.Write("				}" & vbCrLf)
        Response.Write("      window.parent.ASRIntranetOutput.DataArray();" & vbCrLf)
        Response.Write("			lngPageColumnCount = frmOutput.grdCalendarOutput.Columns.Count;" & vbCrLf)
        Response.Write("			if (!blnSettingsDone)" & vbCrLf)
        Response.Write("				{" & vbCrLf)
        Response.Write("				window.parent.ASRIntranetOutput.HeaderRows = 2;" & vbCrLf)
        Response.Write("				window.parent.ASRIntranetOutput.HeaderCols = 1;" & vbCrLf)
        Response.Write("				window.parent.ASRIntranetOutput.SizeColumnsIndependently = true;" & vbCrLf)
        Response.Write("				blnSettingsDone = true;" & vbCrLf)
        Response.Write("				}" & vbCrLf)
        Response.Write("			window.parent.ASRIntranetOutput.InitialiseStyles();" & vbCrLf)
        Response.Write("			window.parent.ASRIntranetOutput.ResetStyles();" & vbCrLf)
        Response.Write("			window.parent.ASRIntranetOutput.ResetColumns();" & vbCrLf)
        Response.Write("			window.parent.ASRIntranetOutput.ResetMerges();" & vbCrLf)
        Response.Write("			lngPageCount++;" & vbCrLf)
        Response.Write("			window.parent.ASRIntranetOutput.ArrayDim((lngPageColumnCount-1), 0);" & vbCrLf)
        Response.Write("			blnBreakCheck = true;" & vbCrLf)
        Response.Write("			sBreakValue = '';" & vbCrLf)
        Response.Write("			lngActualRow = 0;" & vbCrLf)
        Response.Write("			blnNewPage = true;" & vbCrLf)
        Response.Write("			}" & vbCrLf & vbCrLf)

        Response.Write("		else if (frmOutput.grdCalendarOutput.Columns(0).CellText(bm) != '*')" & vbCrLf)
        Response.Write("			{" & vbCrLf)
        Response.Write("			blnBreakCheck = false;" & vbCrLf)
        Response.Write("			blnNewPage = false;" & vbCrLf)
        Response.Write("			if (lngActualRow > 0)" & vbCrLf)
        Response.Write("				{" & vbCrLf)
        Response.Write("				window.parent.ASRIntranetOutput.ArrayReDim();" & vbCrLf)
        Response.Write("				}" & vbCrLf)
        Response.Write("			for (var lngCol=0; lngCol<lngPageColumnCount; lngCol++)" & vbCrLf)
        Response.Write("				{" & vbCrLf)
        Response.Write("				window.parent.ASRIntranetOutput.ArrayAddTo(lngCol, (lngActualRow), frmOutput.grdCalendarOutput.Columns(lngCol).CellText(bm));" & vbCrLf)
        Response.Write("				}" & vbCrLf)
        Response.Write("			}" & vbCrLf)
	
	
        Response.Write("		if (!blnNewPage) " & vbCrLf)
        Response.Write("			{" & vbCrLf)
        Response.Write("			lngActualRow = lngActualRow + 1; " & vbCrLf)
        Response.Write("			}" & vbCrLf)
        Response.Write("		}" & vbCrLf)
        Response.Write("    window.parent.ASRIntranetOutput.Complete();" & vbCrLf)
        Response.Write("    window.parent.parent.ShowDataFrame();" & vbCrLf)
        Response.Write("	}" & vbCrLf)
    End If

    Response.Write("}" & vbCrLf)

    If Not objCalendar.OutputPreview Then
        Response.Write("  window.parent.frmError.txtEventLogID.value = """ & CleanStringForJavaScript(objCalendar.EventLogID) & """;" & vbCrLf)
        Response.Write("  if (frmOriginalDefinition.txtCancelPrint.value == 1) {" & vbCrLf)
        Response.Write("    window.parent.parent.raiseError('',false,true);" & vbCrLf)
        Response.Write("  }" & vbCrLf)
        Response.Write("  else if (window.parent.ASRIntranetOutput.ErrorMessage != '') {" & vbCrLf)
        Response.Write("    window.parent.raiseError(window.parent.ASRIntranetOutput.ErrorMessage,false,false);" & vbCrLf)
        Response.Write("  }" & vbCrLf)
        Response.Write("  else {" & vbCrLf)
        Response.Write("    window.parent.raiseError('',true,false);" & vbCrLf)
        Response.Write("  }" & vbCrLf)
    Else
        Response.Write("  sUtilTypeDesc = window.parent.parent.parent.frames(""top"").frmPopup.txtUtilTypeDesc.value;" & vbCrLf)
        Response.Write("  if (window.parent.ASRIntranetOutput.ErrorMessage != """") {" & vbCrLf)
        Response.Write("    OpenHR.messageBox(sUtilTypeDesc+"" output failed.\n\n"" + window.parent.ASRIntranetOutput.ErrorMessage,48,""Calendar Report"");" & vbCrLf)
        Response.Write("  }" & vbCrLf)
        Response.Write("  else {" & vbCrLf)
        Response.Write("    OpenHR.messageBox(sUtilTypeDesc+"" output complete."",64,""Calendar Report"");" & vbCrLf)
        Response.Write("  }" & vbCrLf)
    End If
					
    Response.Write("	}" & vbCrLf)
    Response.Write("</script>" & vbCrLf & vbCrLf)
		end if
	else
		if fBadUtilDef then 
%>

<input type='hidden' id=Hidden1 name=txtOK value="False">
<table align=center class="outline" cellPadding=5 cellSpacing=0>
	<TR>
		<TD>
			<table class="invisible" cellspacing=0 cellpadding=0>
			    <tr>
			        <td colspan=3 height=10></td>
			    </tr>
			    <tr> 
			        <td colspan=3 align=center> 
						<H3>Error</H3>
			        </td>
			    </tr> 
			    <tr> 
			        <td width=20 height=10></td> 
			        <td> 
						<H4>Not all session variables found</H4>
			        </td>
			        <td width=20></td> 
			    </tr>
			    <tr> 
			        <td width=20 height=10></td> 
			        <td>
			            Type = <%Session("utiltype").ToString()%>
			        </td>
			        <td width=20></td> 
			    </tr>
			    <tr> 
			        <td width=20 height=10></td> 
			        <td>
			            Utility Name = <%Session("utilname").ToString()%>
			        </td>
			        <td width=20></td> 
			    </tr>
			    <tr> 
			        <td width=20 height=10></td> 
			        <td>
			            Utility ID = <%Session("utilid").ToString()%>
			        </td>
			        <td width=20></td> 
			    </tr>
			    <tr> 
			        <td width=20 height=10></td> 
			        <td>
			            Action = <%Session("action").ToString()%>
			        </td>
			        <td width=20></td> 
			    </tr>
			    <tr>
			        <td colspan=3 height=10>&nbsp;</td>
			    </tr>
			    <tr> 
			        <td colspan=3 height=10 align=center> 
						<INPUT TYPE=button VALUE=Close NAME=cmdClose style="WIDTH: 80px" width=80 id=cmdClose class="btn"
						    OnClick=window.parent.self.close(); 
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
<input type=hidden id=txtSuccessFlag name=txtSuccessFlag value=1>

<%
		else
%>

<input type='hidden' id=Hidden2 name=txtOK value="False">
<FORM ID=frmPopup Name=frmPopup>
<table align=center class="outline" cellPadding=5 cellSpacing=0>
	<TR>
		<TD>
			<table class="invisible" cellspacing=0 cellpadding=0>
			    <tr>
    			    <td colspan=3 height=10></td>
	            </tr>
<%
    Dim sCloseFunction As String
    
    Response.Write("			  <tr> " & vbCrLf)
    Response.Write("			    <td width=20 height=10></td> " & vbCrLf)
    Response.Write("			    <td align=center> " & vbCrLf)

    If fNoRecords Then
        Response.Write("						<H4>Calendar Report '" & Session("utilname") & "' Completed successfully.</H4>" & vbCrLf)
        sCloseFunction = "window.parent.self.close();"
    Else
        Response.Write("						<H4>Calendar Report '" & Session("utilname") & "' Failed." & vbCrLf)
        sCloseFunction = "refreshDefSelAndClose();"
    End If
    Response.Write("			    </td>" & vbCrLf)
    Response.Write("			    <td width=20></td> " & vbCrLf)
    Response.Write("			  </tr>" & vbCrLf)
%>
                <tr> 
			        <td width=20 height=10></td> 
			        <td align=center nowrap>
			            <%objCalendar.ErrorString.ToString()%>
			        </td>
			        <td width=20></td> 
			    </tr>
			    <tr>
			        <td colspan=3 height=10>&nbsp;</td>
			    </tr>
			    <tr> 
			        <td colspan=3 height=10 align=center> 
						<INPUT TYPE=button VALUE=Close NAME=cmdClose style="WIDTH: 80px" width=80 id=Button1 class="btn"
						    OnClick="<%sCloseFunction.ToString()%>" 
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
</FORM>
<input type=hidden id=Hidden3 name=txtSuccessFlag value=1>
<input type='hidden' id=txtPreview name=txtPreview value=0>
<%
		end if
	end if

Response.Write("<input type=hidden id=txtTitle name=txtTitle value=""" & Replace(objCalendar.CalendarReportName, """", "&quot;") & """>" & vbCrLf)
	
objCalendar = Nothing
%>

            <form id="frmOriginalDefinition" style="visibility: hidden; display: none">
                <%
                    Dim sErrMsg As String = ""
                    Response.Write("	<INPUT type='hidden' id=txtDefn_Name name=txtDefn_Name value=""" & Replace(Session("utilname").ToString(), """", "&quot;") & """>" & vbCrLf)
                    Response.Write("	<INPUT type='hidden' id=txtDefn_ErrMsg name=txtDefn_ErrMsg value=""" & sErrMsg & """>" & vbCrLf)
                %>
                <input type="hidden" id="txtUserName" name="txtUserName" value="<%Session("username").ToString()%>">
                <input type="hidden" id="txtDateFormat" name="txtDateFormat" value="<%Session("LocaleDateFormat").ToString()%>">
                <input type="hidden" id="txtCancelPrint" name="txtCancelPrint">
                <input type="hidden" id="txtOptionsDone" name="txtOptionsDone">
                <input type="hidden" id="txtOptionsPortrait" name="txtOptionsPortrait">
                <input type="hidden" id="txtOptionsMarginLeft" name="txtOptionsMarginLeft">
                <input type="hidden" id="txtOptionsMarginRight" name="txtOptionsMarginRight">
                <input type="hidden" id="txtOptionsMarginTop" name="txtOptionsMarginTop">
                <input type="hidden" id="txtOptionsMarginBottom" name="txtOptionsMarginBottom">
                <input type="hidden" id="txtOptionsCopies" name="txtOptionsCopies">
                <input type="hidden" id="txtCalRep_UtilID" name="txtCalRep_UtilID" value="<%CalRep_UtilID.ToString()%>">
            </form>

            
<script type="text/javascript">
   
    $("#workframe").hide();
    $("#reportframe").show();

    $("#top").hide();
    $("#calendarframeset").show();

    util_run_calendarreport_main_window_onload();

</script>