<%@ Page Language="VB" Inherits="System.Web.Mvc.ViewPage" %>
<%@ Import Namespace="DMI.NET" %>

<!DOCTYPE html>

<html>
<head runat="server">
    <title>Calendar Report Breakdown</title>
    <script src="<%: Url.Content("~/bundles/jQuery")%>" type="text/javascript"></script>
    <script src="<%: Url.Content("~/bundles/jQueryUI7")%>" type="text/javascript"></script>
    <script src="<%: Url.Content("~/bundles/OpenHR_General")%>" type="text/javascript"></script>
	<link href="<%: Url.LatestContent("~/Content/Site.css")%>" rel="stylesheet" type="text/css" />
	<link href="<%: Url.LatestContent("~/Content/OpenHR.css")%>" rel="stylesheet" type="text/css" />

    <script type="text/javascript">
        
        function calendar_report_breakdown_window_onload() {

            if (frmDetails.txtErrorDescription.value.length > 0) {
                OpenHR.messageBox(txtErrorDescription.value, 0, "OpenHR Intranet"); // 0 = vbOKonly
                self.close();
            }
            else 
            {
                setGridFont(frmDetails.grdDetails);
		
                if (abEventDetails == null) {
                    // The menu control was not loaded properly.
                    OpenHR.messageBox("Menu control not loaded.", 0, "OpenHR Intranet"); // 0 = vbOKOnly
                    self.close();
               }
                else 
                {
                    frmDetails.txtLoading.value = 1;
                    // Load the standard menu options into the menubar.
                    setMenuFont(abEventDetails);

                    abEventDetails.Attach();
                    abEventDetails.DataPath = "misc\\nav.htm";
                    abEventDetails.RecalcLayout();
                    abEventDetails.Refresh();

                    populateGrid();
			
                    frmDetails.txtLoading.value = 0;
                }
            }
        }
        
        function abEventDetails_PreCustomizeMenu(pfCancel) {
            // Do not let the user modify the layout.
            ASRIntranetFunctions.MessageBox("The menu cannot be customized. Errors will occur if you attempt to customize it. Click anywhere in your browser to remove the dummy customisation menu."); 
        }

        function abEventDetails_PreSysMenu(pBand) {

            if (pBand.Name == "SysCustomize") 
            {
                pBand.Tools.RemoveAll();
            }
        }
    
        function abEventDetails_Click(pTool) {
            // Perform the selected menu action.
            var sToolName;

            sToolName = pTool.name;

            with (frmDetails)
            {
                if (sToolName == "Previous") 
                {
                    grdDetails.MovePrevious();
                }
                else if (sToolName == "Next") 
                {
                    grdDetails.MoveNext();
                }
                else if (sToolName == "Last") 
                {
                    grdDetails.MoveLast();
                }
                else
                {
                    grdDetails.MoveFirst();
                }

                grdDetails.SelBookmarks.RemoveAll();
                grdDetails.SelBookmarks.Add(grdDetails.Bookmark);
            }
		
            updateLabels();
            updateRecordStatus();
        }

        function abEventDetails_DataReady() {

            var sKey;
            sKey = new String("tempmenufilepath_");

            try {
                sKey = sKey.concat(window.parent.window.dialogArguments.window.parent.getDBName());
            } catch(e) {
                ASRIntranetFunctions.MessageBox("The calendar window has been closed.", 0, "OpenHR Intranet"); // 0 = vbOKonly
                self.close();
                return;
            }

            sPath = ASRIntranetFunctions.GetRegistrySetting("HR Pro", "DataPaths", sKey);

            if (sPath == "") {
                sPath = "c:\\";
            }

            if (sPath == "<NONE>") {
                frmUseful.txtMenuSaved.value = 1;
                abEventDetails.RecalcLayout();
            } else {
                if (sPath.substr(sPath.length - 1, 1) != "\\") {
                    sPath = sPath.concat("\\");
                }

                sPath = sPath.concat("tempnav.asp");
                if ((abEventDetails.Bands.Count() > 0) && (frmUseful.txtMenuSaved.value == 0)) {
                    abEventDetails.save(sPath, "");
                    frmUseful.txtMenuSaved.value = 1;
                } else {
                    if ((abEventDetails.Bands.Count() == 0) && (frmUseful.txtMenuSaved.value == 1)) {
                        abEventDetails.DataPath = sPath;
                        abEventDetails.RecalcLayout();
                        return;
                    }
                }

                // Resize the grid to show all prompted values.
                iResizeBy = bdyMain.scrollWidth - bdyMain.clientWidth;
                if (bdyMain.offsetWidth + iResizeBy > screen.width) {
                    window.dialogWidth = new String(screen.width) + "px";
                } else {
                    iNewWidth = new Number(window.dialogWidth.substr(0, window.dialogWidth.length - 2));
                    iNewWidth = iNewWidth + iResizeBy;
                    window.dialogWidth = new String(iNewWidth) + "px";
                }

                iResizeBy = bdyMain.scrollHeight - bdyMain.clientHeight
                    + abEventDetails.Bands.Item("bndDetails").height;
                if (bdyMain.offsetHeight + iResizeBy > screen.height) {
                    window.dialogHeight = new String(screen.height) + "px";
                } else {
                    iNewHeight = new Number(window.dialogHeight.substr(0, window.dialogHeight.length - 2));
                    iNewHeight = iNewHeight + iResizeBy;
                    window.dialogHeight = new String(iNewHeight) + "px";
                }
            }
        }

        function okClick()
        {
            self.close();
            return;
        }
	
        function populateGrid() {

            var objBaseCTL;
            var vControlName;
            var frmCalDetails = window.dialogArguments.OpenHR.getForm("calendarframe_calendar", "frmEventDetails");
            var docCalendar = window.dialogArguments.document;

            var sCollectionKey;
            var strAddLine = new String("");
            var iCollectionCount;
	
            with (frmDetails)
            {
                sCollectionKey = frmCalDetails.txtBaseIndex.value + '_CALDATEINDEX_' + frmCalDetails.txtLabelIndex.value;
                vControlName = 'ctlCalRec_' + frmCalDetails.txtBaseIndex.value;
                objBaseCTL = docCalendar.getElementById(vControlName);
                iCollectionCount = objBaseCTL.InitialiseEventsCollection(sCollectionKey);

                while (grdDetails.Rows > 0) 
                {
                    grdDetails.RemoveAll();
                }
		
                grdDetails.visible = true;
                grdDetails.focus();
                grdDetails.Redraw = false;
		
                for (var i=1; i<=iCollectionCount; i++)
                {
                    strAddLine = new String(objBaseCTL.GetAddLine(i,frmCalDetails.txtLabelIndex.value));
                    grdDetails.AddItem(strAddLine);	
                }
	
                grdDetails.visible = true;
                grdDetails.Redraw = true;
		
                if (grdDetails.Rows > 0)
                {
                    grdDetails.Refresh();
                    grdDetails.MoveFirst();
                    grdDetails.SelBookmarks.RemoveAll();
                    grdDetails.SelBookmarks.Add(grdDetails.Bookmark);
                    updateRecordStatus();
                    updateLabels();
                }
            }	
            return;
        }
	
        function replace(sExpression, sFind, sReplace)
        {
            //gi (global search, ignore case)
            var re = new RegExp(sFind,"gi");
            sExpression = sExpression.replace(re, sReplace);
            return(sExpression);
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

        function updateLabels()
        {
            with (frmDetails)
            {
                document.getElementById('tdEventName').innerHTML = grdDetails.Columns('EventName').Text;
                document.getElementById('tdBaseDesc').innerHTML = grdDetails.Columns('BaseDescription').Text;
		
                document.getElementById('tdStartDate').innerHTML = grdDetails.Columns('StartDate').Text + ' ' + grdDetails.Columns('StartSession').Text;
                document.getElementById('tdEndDate').innerHTML = grdDetails.Columns('EndDate').Text + ' ' + grdDetails.Columns('EndSession').Text;

                document.getElementById('tdDuration').innerHTML = grdDetails.Columns('Duration').Text;
	
                if (trim(grdDetails.Columns('EventDescription1Column').Text) == '') 
                {
                    document.getElementById('tdEventDesc1Column').innerText = '';
                    document.getElementById('tdEventDesc1Value').innerHTML = '';
                }
                else
                {
                    document.getElementById('tdEventDesc1Column').innerText = replace(grdDetails.Columns('EventDescription1Column').Text,'_',' ') + ' : ';
                    document.getElementById('tdEventDesc1Value').innerHTML = grdDetails.Columns('EventDescription1Value').Text;
                }
  
                if (trim(grdDetails.Columns('EventDescription2Column').Text) == '')
                {
                    document.getElementById('tdEventDesc2Column').innerText = '';
                    document.getElementById('tdEventDesc2Value').innerHTML = '';
                }
                else
                {
                    document.getElementById('tdEventDesc2Column').innerText = replace(grdDetails.Columns('EventDescription2Column').Text,'_',' ') + ' : ';
                    document.getElementById('tdEventDesc2Value').innerHTML = grdDetails.Columns('EventDescription2Value').Text;
                }
		
                document.getElementById('tdCalendarCode').innerHTML = grdDetails.Columns('Legend').Text;
		
                if (txtShowRegion.value == 1)
                {
                    document.getElementById('tdRegion').innerHTML = grdDetails.Columns('Region').Text;
                }
                if (txtShowWorkingPattern.value == 1)
                {
                    populateWorkingPatternCTL(grdDetails.Columns("WorkingPattern").Text);
                }
            }
	
            return;
        }

        function updateRecordStatus()
        {
            try 
            {
                abEventDetails.Bands("bndDetails").Tools("First").Enabled = false;
            }
            catch(e)
            {
                window.setTimeout("updateRecordStatus()", 250);
                return;
            }
		
            with (abEventDetails)
            {
                RecalcLayout();
                Refresh();
		
                if (frmDetails.grdDetails.Rows == 1)
                {
                    Bands("bndDetails").Tools("First").Enabled = false;
                    Bands("bndDetails").Tools("Previous").Enabled = false;
                    Bands("bndDetails").Tools("Next").Enabled = false;
                    Bands("bndDetails").Tools("Last").Enabled = false;
                }
                else if (frmDetails.grdDetails.AddItemRowIndex(frmDetails.grdDetails.Bookmark) == 0)
                {
                    Bands("bndDetails").Tools("First").Enabled = false;
                    Bands("bndDetails").Tools("Previous").Enabled = false;
                    Bands("bndDetails").Tools("Next").Enabled = true;
                    Bands("bndDetails").Tools("Last").Enabled = true;
                }
                else if (frmDetails.grdDetails.AddItemRowIndex(frmDetails.grdDetails.Bookmark) == (frmDetails.grdDetails.Rows - 1)) 
                {
                    Bands("bndDetails").Tools("First").Enabled = true;
                    Bands("bndDetails").Tools("Previous").Enabled = true;
                    Bands("bndDetails").Tools("Next").Enabled = false;
                    Bands("bndDetails").Tools("Last").Enabled = false;
                }
                else
                {
                    Bands("bndDetails").Tools("First").Enabled = true;
                    Bands("bndDetails").Tools("Previous").Enabled = true;
                    Bands("bndDetails").Tools("Next").Enabled = true;
                    Bands("bndDetails").Tools("Last").Enabled = true;
                }
		
                Bands("bndDetails").Tools("Record").Caption = "Record " + (frmDetails.grdDetails.AddItemRowIndex(frmDetails.grdDetails.Bookmark) + 1) + " of " + frmDetails.grdDetails.Rows;
                RecalcLayout();
                Refresh();
            }
		
            return;
        }
	
        function populateWorkingPatternCTL(pstrWPValue)
        {
            var sControlName;
            var ctl;
            var strValue = new String(pstrWPValue);
	
            for (var i=0; i<14; i++)
            {
                sControlName = 'wp_' + (i+1);
		
                ctl = document.getElementById(sControlName);
		
                if (strValue.substring(i,i+1) != ' ')
                {
                    ctl.checked = true;
                }
                else
                {
                    ctl.checked = false;
                }
            }
        }

    </script>

</head>
<body>

    <object classid="clsid:6976CB54-C39B-4181-B1DC-1A829068E2E7"
        codebase="cabs/COAInt_Client.cab#Version=1,0,0,5"
        height="26" id="abEventDetails" name="abEventDetails" style="margin-top: -20px; LEFT: 0px; TOP: 0px" width="100%">
        <param name="_ExtentX" value="847">
        <param name="_ExtentY" value="847">
    </object>

    <div>


<FORM id=frmDetails name=frmDetails>

<table align=center width=100% height=95% class="invisible" cellPadding=2 cellSpacing=0>
	<tr>
		<td valign=top width=100% height=100%>
			<table class="outline" cellspacing="0" cellpadding="4" width=100% height=100%>
				<tr height=10> 
					<td height=10 colspan=5 align=left valign=top>
						Details : <BR><BR>
						<TABLE class="invisible" cellspacing="0" cellpadding="0" width=100%>
							<tr height=5>
								<td width=5></td>
								<td nowrap valign=top>Event Name :</td>
								<td width=5></td>
								<td ID=tdEventName NAME=tdEventName valign=top>
								</td>
								<td width=5></td>
							</tr>
							<tr height=5> 
								<td colspan=5></td>
							</tr>
							<tr height=5>
								<td width=5></td>
								<td nowrap valign=top>Description :</td>
								<td width=5></td>
								<td ID=tdBaseDesc NAME=tdBaseDesc valign=top>
								</td>
								<td width=5></td>
							</tr>
							<tr height=5> 
								<td colspan="5"><HR width=90%></td>
							</tr>
							<tr height=5>
								<td width=5></td>
								<td nowrap valign=top>Start Date :</td>
								<td width=5></td>
								<td ID=tdStartDate NAME=tdStartDate valign=top>
								</td>
								<td width=5></td>
							</tr>
							<tr height=5> 
								<td colspan="5"></td>
							</tr>
							<tr height=5>
								<td width=5></td>
								<td nowrap valign=top>End Date :</td>
								<td width=5></td>
								<td ID=tdEndDate NAME=tdEndDate valign=top>
								</td>
								<td width=5></td>
							</tr>
							<tr height=5> 
								<td colspan="5"></td>
							</tr>
							<tr height=5>
								<td width=5></td>
								<td nowrap valign=top>Duration (Actual) :</td>
								<td width=5></td>
								<td ID=tdDuration NAME=tdDuration valign=top>
									
								</td>
								<td width=5></td>
							</tr>
							<tr height=5> 
								<td colspan="5"><HR width=90%></td>
							</tr>	
							<tr height=5>
								<td width=5></td>
								<td nowrap ID=tdEventDesc1Column NAME=tdEventDesc1Column valign=top>Event Descripiton 1 :</td>
								<td width=5></td>
								<td ID=tdEventDesc1Value NAME=tdEventDesc1Value valign=top>
									
								</td>
								<td width=5></td>
							</tr>
							<tr height=5> 
								<td colspan="5"></td>
							</tr>
							<tr height=5>
								<td width=5></td>
								<td nowrap ID=tdEventDesc2Column NAME=tdEventDesc2Column valign=top>Event Descripiton 2 :</td>
								<td width=5></td>
								<td ID=tdEventDesc2Value NAME=tdEventDesc2Value valign=top>
									
								</td>
								<td width=5></td>
							</tr>
							<tr height=5> 
								<td colspan="5"><HR width=90%></td>
							</tr>
							<tr height=5>
								<td width=5></td>
								<td nowrap valign=top>Calendar Code :</td>
								<td width=5></td>
								<td ID=tdCalendarCode NAME=tdCalendarCode valign=top>
									
								</td>
								<td width=5></td>
							</tr>
<% 

    If (Request("txtShowRegion") = 1) Or (Request("txtShowWorkingPattern") = 1) Then
        Response.Write("			<tr height=5> " & vbCrLf)
        Response.Write("						<td colspan=5><HR width=90")
        Response.Write("%")
        Response.Write("						></td>" & vbCrLf)
        Response.Write("					</tr>" & vbCrLf)
    End If
	
    If Request("txtShowRegion") = "1" Then
        Response.Write("				<tr height=5>" & vbCrLf)
        Response.Write("					<td width=5></td>" & vbCrLf)
        Response.Write("					<td nowrap valign=top>Region :</td>" & vbCrLf)
        Response.Write("					<td width=5></td>" & vbCrLf)
        Response.Write("					<td ID=tdRegion NAME=tdRegion valign=top>" & vbCrLf)
        Response.Write("					</td>" & vbCrLf)
        Response.Write("					<td width=5></td>" & vbCrLf)
        Response.Write("				</tr>" & vbCrLf)
        Response.Write("				<tr height=5> " & vbCrLf)
        Response.Write("					<td colspan=5></td>" & vbCrLf)
        Response.Write("				</tr>		" & vbCrLf)
    End If
	
    If Request("txtShowWorkingPattern") = "1" Then
        Response.Write("				<tr height=5>" & vbCrLf)
        Response.Write("					<td width=5></td>" & vbCrLf)
        Response.Write("					<td nowrap valign=top>Working Pattern :</td>" & vbCrLf)
        Response.Write("					<td width=5></td>" & vbCrLf)
        Response.Write("					<td valign=top>" & vbCrLf)

        Response.Write("					<table class=""outline"" cellspacing=0 cellpadding=4 frame=0>" & vbCrLf)
        Response.Write("						<TR align=middle>" & vbCrLf)
        Response.Write("							<TD>" & vbCrLf)

        Response.Write("					<table class=""invisible"" cellspacing=0 cellpadding=1 frame=0>" & vbCrLf)
		
		
        Dim objCalendar As HR.Intranet.Server.CalendarReport
        objCalendar = Session("objCalendar" & Request("CalRepUtilID"))

        Response.Write(objCalendar.WorkingPatternTitle)
        Response.Write("						<TR>" & vbCrLf)
        Response.Write("							<TD ALIGN=center VALIGN=middle>" & vbCrLf)
        Response.Write("								AM" & vbCrLf)
        Response.Write("							</TD>" & vbCrLf)
        Response.Write("							<TD ALIGN=center VALIGN=middle>" & vbCrLf)
        Response.Write("								<INPUT id=wp_1 name=wp_1 type=checkbox style=""HEIGHT: 14px; WIDTH: 14px"" disabled=""disabled"">" & vbCrLf)
        Response.Write("							</TD>" & vbCrLf)
        Response.Write("							<TD ALIGN=center VALIGN=middle>" & vbCrLf)
        Response.Write("								<INPUT id=wp_3 name=wp_3 type=checkbox style=""HEIGHT: 14px; WIDTH: 14px"" disabled=""disabled"">" & vbCrLf)
        Response.Write("							</TD>" & vbCrLf)
        Response.Write("							<TD ALIGN=center VALIGN=middle>" & vbCrLf)
        Response.Write("								<INPUT id=wp_5 name=wp_5 type=checkbox style=""HEIGHT: 14px; WIDTH: 14px"" disabled=""disabled"">" & vbCrLf)
        Response.Write("							</TD>" & vbCrLf)
        Response.Write("							<TD ALIGN=center VALIGN=middle>" & vbCrLf)
        Response.Write("								<INPUT id=wp_7 name=wp_7 type=checkbox style=""HEIGHT: 14px; WIDTH: 14px"" disabled=""disabled"">" & vbCrLf)
        Response.Write("							</TD>" & vbCrLf)
        Response.Write("							<TD ALIGN=center VALIGN=middle>" & vbCrLf)
        Response.Write("								<INPUT id=wp_9 name=wp_9 type=checkbox style=""HEIGHT: 14px; WIDTH: 14px"" disabled=""disabled"">" & vbCrLf)
        Response.Write("							</TD>" & vbCrLf)
        Response.Write("							<TD ALIGN=center VALIGN=middle>" & vbCrLf)
        Response.Write("								<INPUT id=wp_11 name=wp_11 type=checkbox style=""HEIGHT: 14px; WIDTH: 14px"" disabled=""disabled"">" & vbCrLf)
        Response.Write("							</TD>" & vbCrLf)
        Response.Write("							<TD ALIGN=center VALIGN=middle>" & vbCrLf)
        Response.Write("								<INPUT id=wp_13 name=wp_13 type=checkbox style=""HEIGHT: 14px; WIDTH: 14px"" disabled=""disabled"">" & vbCrLf)
        Response.Write("							</TD>" & vbCrLf)
        Response.Write("						</TR>" & vbCrLf)
        Response.Write("						<TR>" & vbCrLf)
        Response.Write("							<TD ALIGN=center VALIGN=middle>" & vbCrLf)
        Response.Write("								PM" & vbCrLf)
        Response.Write("							</TD>" & vbCrLf)
        Response.Write("							<TD ALIGN=center VALIGN=middle>" & vbCrLf)
        Response.Write("								<INPUT id=wp_2 name=wp_2 type=checkbox style=""HEIGHT: 14px; WIDTH: 14px"" disabled>" & vbCrLf)
        Response.Write("							</TD>" & vbCrLf)
        Response.Write("							<TD ALIGN=center VALIGN=middle>" & vbCrLf)
        Response.Write("								<INPUT id=wp_4 name=wp_4 type=checkbox style=""HEIGHT: 14px; WIDTH: 14px"" disabled>" & vbCrLf)
        Response.Write("							</TD>" & vbCrLf)
        Response.Write("							<TD ALIGN=center VALIGN=middle>" & vbCrLf)
        Response.Write("								<INPUT id=wp_6 name=wp_6 type=checkbox style=""HEIGHT: 14px; WIDTH: 14px"" disabled>" & vbCrLf)
        Response.Write("							</TD>" & vbCrLf)
        Response.Write("							<TD ALIGN=center VALIGN=middle>" & vbCrLf)
        Response.Write("								<INPUT id=wp_8 name=wp_8 type=checkbox style=""HEIGHT: 14px; WIDTH: 14px"" disabled>" & vbCrLf)
        Response.Write("							</TD>" & vbCrLf)
        Response.Write("							<TD ALIGN=center VALIGN=middle>" & vbCrLf)
        Response.Write("								<INPUT id=wp_10 name=wp_10 type=checkbox style=""HEIGHT: 14px; WIDTH: 14px"" disabled>" & vbCrLf)
        Response.Write("							</TD>" & vbCrLf)
        Response.Write("							<TD ALIGN=center VALIGN=middle>" & vbCrLf)
        Response.Write("								<INPUT id=wp_12 name=wp_12 type=checkbox style=""HEIGHT: 14px; WIDTH: 14px"" disabled>" & vbCrLf)
        Response.Write("							</TD>" & vbCrLf)
        Response.Write("							<TD ALIGN=center VALIGN=middle>" & vbCrLf)
        Response.Write("								<INPUT id=wp_14 name=wp_14 type=checkbox style=""HEIGHT: 14px; WIDTH: 14px"" disabled>" & vbCrLf)
        Response.Write("							</TD>" & vbCrLf)
        Response.Write("						</TR>" & vbCrLf)
        Response.Write("					</TABLE>" & vbCrLf)

        Response.Write("							</TD>" & vbCrLf)
        Response.Write("						</TR>" & vbCrLf)
        Response.Write("					</TABLE>" & vbCrLf)
		
        Response.Write("					</td>" & vbCrLf)
        Response.Write("					<td width=5></td>" & vbCrLf)
        Response.Write("				</tr>" & vbCrLf)
        Response.Write("				<tr height=5> " & vbCrLf)
        Response.Write("					<td colspan=5></td>" & vbCrLf)
        Response.Write("				</tr>		" & vbCrLf)
    End If

    Response.Write("<INPUT type='hidden' id=txtErrorDescription name=txtErrorDescription value=''>" & vbCrLf)

%>
	
						</TABLE>
					</td>
				</tr>
			</table>
		</td>
		<tr height=5> 
			<td colspan=5 align=right>
			<input type=button id=cmdOK WIDTH=80 name=cmdOK value=OK style="WIDTH: 80px" class="btn"
			    onClick="okClick();" />
			</td>
		</tr>	
	</tr>
</table>

<%

	dim avColumnDef(13,4)
	
	avColumnDef(0,0) = "EventName"				'name
	avColumnDef(0,1) = "EventName"				'caption
	avColumnDef(0,2) = "2000"				'width
	avColumnDef(0,3) = "-1"					'visible
	
	avColumnDef(1,0) = "BaseDescription"			'name
	avColumnDef(1,1) = "BaseDescription"			'caption
	avColumnDef(1,2) = "1814"				'width
	avColumnDef(1,3) = "-1"					'visible

	avColumnDef(2,0) = "StartDate"				'name
	avColumnDef(2,1) = "StartDate"				'caption
	avColumnDef(2,2) = "2000"				'width
	avColumnDef(2,3) = "-1"					'visible

	avColumnDef(3,0) = "StartSession"			'name
	avColumnDef(3,1) = "StartSession"			'caption
	avColumnDef(3,2) = "1814"				'width
	avColumnDef(3,3) = "-1"					'visible

	avColumnDef(4,0) = "EndDate"				'name
	avColumnDef(4,1) = "EndDate"				'caption
	avColumnDef(4,2) = "2000"				'width
	avColumnDef(4,3) = "-1"					'visible

	avColumnDef(5,0) = "EndSession"		'name
	avColumnDef(5,1) = "EndSession"		'caption
	avColumnDef(5,2) = "1814"				'width
	avColumnDef(5,3) = "0"					'visible

	avColumnDef(6,0) = "Duration"			'name
	avColumnDef(6,1) = "Duration"			'caption
	avColumnDef(6,2) = "2000"				'width
	avColumnDef(6,3) = "-1"					'visible

	avColumnDef(7,0) = "EventDescription1Column"		'name
	avColumnDef(7,1) = "EventDescription1Column"		'caption
	avColumnDef(7,2) = "1814"				'width
	avColumnDef(7,3) = "-1"					'visible
	
	avColumnDef(8,0) = "EventDescription1Value"		'name
	avColumnDef(8,1) = "EventDescription1Value"		'caption
	avColumnDef(8,2) = "2250"				'width
	avColumnDef(8,3) = "-1"					'visible
	
	avColumnDef(9,0) = "EventDescription2Column"			'name
	avColumnDef(9,1) = "EventDescription2Column"			'caption
	avColumnDef(9,2) = "1814"				'width
	avColumnDef(9,3) = "-1"					'visible
	
	avColumnDef(10,0) = "EventDescription2Value"			'name
	avColumnDef(10,1) = "EventDescription2Value"			'caption
	avColumnDef(10,2) = "2000"				'width
	avColumnDef(10,3) = "-1"				'visible
	
	avColumnDef(11,0) = "Legend"		'name
	avColumnDef(11,1) = "Legend"		'caption
	avColumnDef(11,2) = "1814"				'width
	avColumnDef(11,3) = "-1"					'visible
	
	avColumnDef(12,0) = "WorkingPattern"		'name
	avColumnDef(12,1) = "WorkingPattern"		'caption
	avColumnDef(12,2) = "2250"				'width
	avColumnDef(12,3) = "-1"				'visible
	
	avColumnDef(13,0) = "Region"		'name
	avColumnDef(13,1) = "Region"		'caption
	avColumnDef(13,2) = "1814"				'width
	avColumnDef(13,3) = "-1"					'visible
	
    Response.Write("											<OBJECT classid=clsid:4A4AA697-3E6F-11D2-822F-00104B9E07A1" & vbCrLf)
    Response.Write("													 codebase=""cabs/COAInt_Grid.cab#version=3,1,3,6""" & vbCrLf)
    Response.Write("													id=grdDetails" & vbCrLf)
    Response.Write("													name=grdDetails" & vbCrLf)
    Response.Write("													style=""HEIGHT: 0px; LEFT: 0px; TOP: 0px; WIDTH: 0px; POSITION: absolute"">")
    Response.Write("													>" & vbCrLf)
    Response.Write("												<PARAM NAME=""ScrollBars"" VALUE=""3"">" & vbCrLf)
    Response.Write("												<PARAM NAME=""_Version"" VALUE=""196617"">" & vbCrLf)
    Response.Write("												<PARAM NAME=""DataMode"" VALUE=""2"">" & vbCrLf)
    Response.Write("												<PARAM NAME=""Cols"" VALUE=""0"">" & vbCrLf)
    Response.Write("												<PARAM NAME=""Rows"" VALUE=""0"">" & vbCrLf)
    Response.Write("												<PARAM NAME=""BorderStyle"" VALUE=""1"">" & vbCrLf)
    Response.Write("												<PARAM NAME=""RecordSelectors"" VALUE=""0"">" & vbCrLf)
    Response.Write("												<PARAM NAME=""GroupHeaders"" VALUE=""-1"">" & vbCrLf)
    Response.Write("												<PARAM NAME=""ColumnHeaders"" VALUE=""-1"">" & vbCrLf)
    Response.Write("												<PARAM NAME=""GroupHeadLines"" VALUE=""1"">" & vbCrLf)
    Response.Write("												<PARAM NAME=""HeadLines"" VALUE=""2"">" & vbCrLf)
    Response.Write("												<PARAM NAME=""FieldDelimiter"" VALUE=""(None)"">" & vbCrLf)
    Response.Write("												<PARAM NAME=""FieldSeparator"" VALUE=""(Tab)"">" & vbCrLf)
    Response.Write("												<PARAM NAME=""Row.Count"" VALUE=""0"">" & vbCrLf)
    Response.Write("												<PARAM NAME=""Col.Count"" VALUE=""1"">" & vbCrLf)
    Response.Write("												<PARAM NAME=""TagVariant"" VALUE=""EMPTY"">" & vbCrLf)
    Response.Write("												<PARAM NAME=""UseGroups"" VALUE=""0"">" & vbCrLf)
    Response.Write("												<PARAM NAME=""HeadFont3D"" VALUE=""0"">" & vbCrLf)
    Response.Write("												<PARAM NAME=""Font3D"" VALUE=""0"">" & vbCrLf)
    Response.Write("												<PARAM NAME=""DividerType"" VALUE=""3"">" & vbCrLf)
    Response.Write("												<PARAM NAME=""DividerStyle"" VALUE=""1"">" & vbCrLf)
    Response.Write("												<PARAM NAME=""DefColWidth"" VALUE=""0"">" & vbCrLf)
    Response.Write("												<PARAM NAME=""BeveColorScheme"" VALUE=""2"">" & vbCrLf)
    Response.Write("												<PARAM NAME=""BevelColorFrame"" VALUE=""-2147483642"">" & vbCrLf)
    Response.Write("												<PARAM NAME=""BevelColorHighlight"" VALUE=""-2147483628"">" & vbCrLf)
    Response.Write("												<PARAM NAME=""BevelColorShadow"" VALUE=""-2147483632"">" & vbCrLf)
    Response.Write("												<PARAM NAME=""BevelColorFace"" VALUE=""-2147483633"">" & vbCrLf)
    Response.Write("												<PARAM NAME=""CheckBox3D"" VALUE=""-1"">" & vbCrLf)
    Response.Write("												<PARAM NAME=""AllowAddNew"" VALUE=""0"">" & vbCrLf)
    Response.Write("												<PARAM NAME=""AllowDelete"" VALUE=""0"">" & vbCrLf)
    Response.Write("												<PARAM NAME=""AllowUpdate"" VALUE=""0"">" & vbCrLf)
    Response.Write("												<PARAM NAME=""MultiLine"" VALUE=""0"">" & vbCrLf)
    Response.Write("												<PARAM NAME=""ActiveCellStyleSet"" VALUE="""">" & vbCrLf)
    Response.Write("												<PARAM NAME=""RowSelectionStyle"" VALUE=""0"">" & vbCrLf)
    Response.Write("												<PARAM NAME=""AllowRowSizing"" VALUE=""0"">" & vbCrLf)
    Response.Write("												<PARAM NAME=""AllowGroupSizing"" VALUE=""0"">" & vbCrLf)
    Response.Write("												<PARAM NAME=""AllowColumnSizing"" VALUE=""-1"">" & vbCrLf)
    Response.Write("												<PARAM NAME=""AllowGroupMoving"" VALUE=""0"">" & vbCrLf)
    Response.Write("												<PARAM NAME=""AllowColumnMoving"" VALUE=""0"">" & vbCrLf)
    Response.Write("												<PARAM NAME=""AllowGroupSwapping"" VALUE=""0"">" & vbCrLf)
    Response.Write("												<PARAM NAME=""AllowColumnSwapping"" VALUE=""0"">" & vbCrLf)
    Response.Write("												<PARAM NAME=""AllowGroupShrinking"" VALUE=""0"">" & vbCrLf)
    Response.Write("												<PARAM NAME=""AllowColumnShrinking"" VALUE=""0"">" & vbCrLf)
    Response.Write("												<PARAM NAME=""AllowDragDrop"" VALUE=""0"">" & vbCrLf)
    Response.Write("												<PARAM NAME=""UseExactRowCount"" VALUE=""-1"">" & vbCrLf)
    Response.Write("												<PARAM NAME=""SelectTypeCol"" VALUE=""0"">" & vbCrLf)
    Response.Write("												<PARAM NAME=""SelectTypeRow"" VALUE=""3"">" & vbCrLf)
    Response.Write("												<PARAM NAME=""SelectByCell"" VALUE=""-1"">" & vbCrLf)
    Response.Write("												<PARAM NAME=""BalloonHelp"" VALUE=""0"">" & vbCrLf)
    Response.Write("												<PARAM NAME=""RowNavigation"" VALUE=""2"">" & vbCrLf)
    Response.Write("												<PARAM NAME=""CellNavigation"" VALUE=""0"">" & vbCrLf)
    Response.Write("												<PARAM NAME=""MaxSelectedRows"" VALUE=""0"">" & vbCrLf)
    Response.Write("												<PARAM NAME=""HeadStyleSet"" VALUE="""">" & vbCrLf)
    Response.Write("												<PARAM NAME=""StyleSet"" VALUE="""">" & vbCrLf)
    Response.Write("												<PARAM NAME=""ForeColorEven"" VALUE=""0"">" & vbCrLf)
    Response.Write("												<PARAM NAME=""ForeColorOdd"" VALUE=""0"">" & vbCrLf)
    Response.Write("												<PARAM NAME=""BackColorEven"" VALUE=""16777215"">" & vbCrLf)
    Response.Write("												<PARAM NAME=""BackColorOdd"" VALUE=""16777215"">" & vbCrLf)
    Response.Write("												<PARAM NAME=""Levels"" VALUE=""1"">" & vbCrLf)
    Response.Write("												<PARAM NAME=""RowHeight"" VALUE=""503"">" & vbCrLf)
    Response.Write("												<PARAM NAME=""ExtraHeight"" VALUE=""0"">" & vbCrLf)
    Response.Write("												<PARAM NAME=""ActiveRowStyleSet"" VALUE="""">" & vbCrLf)
    Response.Write("												<PARAM NAME=""CaptionAlignment"" VALUE=""2"">" & vbCrLf)
    Response.Write("												<PARAM NAME=""SplitterPos"" VALUE=""0"">" & vbCrLf)
    Response.Write("												<PARAM NAME=""SplitterVisible"" VALUE=""0"">" & vbCrLf)
    Response.Write("												<PARAM NAME=""Columns.Count"" VALUE=""" & (UBound(avColumnDef) + 1) & """>" & vbCrLf)
	
    For i = 0 To UBound(avColumnDef) Step 1
        Response.Write("												<!--" & avColumnDef(i, 0) & "-->  " & vbCrLf)
        Response.Write("												<PARAM NAME=""Columns(" & i & ").Width"" VALUE=""" & avColumnDef(i, 2) & """>" & vbCrLf)
        Response.Write("												<PARAM NAME=""Columns(" & i & ").Visible"" VALUE=""" & avColumnDef(i, 3) & """>" & vbCrLf)
        Response.Write("												<PARAM NAME=""Columns(" & i & ").Columns.Count"" VALUE=""1"">" & vbCrLf)
        Response.Write("												<PARAM NAME=""Columns(" & i & ").Caption"" VALUE=""" & avColumnDef(i, 1) & """>" & vbCrLf)
        Response.Write("												<PARAM NAME=""Columns(" & i & ").Name"" VALUE=""" & avColumnDef(i, 0) & """>" & vbCrLf)
        Response.Write("												<PARAM NAME=""Columns(" & i & ").Alignment"" VALUE=""0"">" & vbCrLf)
        Response.Write("												<PARAM NAME=""Columns(" & i & ").CaptionAlignment"" VALUE=""3"">" & vbCrLf)
        Response.Write("												<PARAM NAME=""Columns(" & i & ").Bound"" VALUE=""0"">" & vbCrLf)
        Response.Write("												<PARAM NAME=""Columns(" & i & ").AllowSizing"" VALUE=""1"">" & vbCrLf)
        Response.Write("												<PARAM NAME=""Columns(" & i & ").DataField"" VALUE=""Column " & i & """>" & vbCrLf)
        Response.Write("												<PARAM NAME=""Columns(" & i & ").DataType"" VALUE=""8"">" & vbCrLf)
        Response.Write("												<PARAM NAME=""Columns(" & i & ").Level"" VALUE=""0"">" & vbCrLf)
        Response.Write("												<PARAM NAME=""Columns(" & i & ").NumberFormat"" VALUE="""">" & vbCrLf)
        Response.Write("												<PARAM NAME=""Columns(" & i & ").Case"" VALUE=""0"">" & vbCrLf)
        Response.Write("												<PARAM NAME=""Columns(" & i & ").FieldLen"" VALUE=""256"">" & vbCrLf)
        Response.Write("												<PARAM NAME=""Columns(" & i & ").VertScrollBar"" VALUE=""0"">" & vbCrLf)
        Response.Write("												<PARAM NAME=""Columns(" & i & ").Locked"" VALUE=""0"">" & vbCrLf)
        Response.Write("												<PARAM NAME=""Columns(" & i & ").Style"" VALUE=""0"">" & vbCrLf)
        Response.Write("												<PARAM NAME=""Columns(" & i & ").ButtonsAlways"" VALUE=""0"">" & vbCrLf)
        Response.Write("												<PARAM NAME=""Columns(" & i & ").RowCount"" VALUE=""0"">" & vbCrLf)
        Response.Write("												<PARAM NAME=""Columns(" & i & ").ColCount"" VALUE=""1"">" & vbCrLf)
        Response.Write("												<PARAM NAME=""Columns(" & i & ").HasHeadForeColor"" VALUE=""0"">" & vbCrLf)
        Response.Write("												<PARAM NAME=""Columns(" & i & ").HasHeadBackColor"" VALUE=""0"">" & vbCrLf)
        Response.Write("												<PARAM NAME=""Columns(" & i & ").HasForeColor"" VALUE=""0"">" & vbCrLf)
        Response.Write("												<PARAM NAME=""Columns(" & i & ").HasBackColor"" VALUE=""0"">" & vbCrLf)
        Response.Write("												<PARAM NAME=""Columns(" & i & ").HeadForeColor"" VALUE=""0"">" & vbCrLf)
        Response.Write("												<PARAM NAME=""Columns(" & i & ").HeadBackColor"" VALUE=""0"">" & vbCrLf)
        Response.Write("												<PARAM NAME=""Columns(" & i & ").ForeColor"" VALUE=""0"">" & vbCrLf)
        Response.Write("												<PARAM NAME=""Columns(" & i & ").BackColor"" VALUE=""0"">" & vbCrLf)
        Response.Write("												<PARAM NAME=""Columns(" & i & ").HeadStyleSet"" VALUE="""">" & vbCrLf)
        Response.Write("												<PARAM NAME=""Columns(" & i & ").StyleSet"" VALUE="""">" & vbCrLf)
        Response.Write("												<PARAM NAME=""Columns(" & i & ").Nullable"" VALUE=""1"">" & vbCrLf)
        Response.Write("												<PARAM NAME=""Columns(" & i & ").Mask"" VALUE="""">" & vbCrLf)
        Response.Write("												<PARAM NAME=""Columns(" & i & ").PromptInclude"" VALUE=""0"">" & vbCrLf)
        Response.Write("												<PARAM NAME=""Columns(" & i & ").ClipMode"" VALUE=""0"">" & vbCrLf)
        Response.Write("												<PARAM NAME=""Columns(" & i & ").PromptChar"" VALUE=""95"">" & vbCrLf)
    Next
		
    Response.Write("												<PARAM NAME=""UseDefaults"" VALUE=""-1"">" & vbCrLf)
    Response.Write("												<PARAM NAME=""TabNavigation"" VALUE=""1"">" & vbCrLf)
    Response.Write("												<PARAM NAME=""BatchUpdate"" VALUE=""0"">" & vbCrLf)
    Response.Write("												<PARAM NAME=""_ExtentX"" VALUE=""11298"">" & vbCrLf)
    Response.Write("												<PARAM NAME=""_ExtentY"" VALUE=""3969"">" & vbCrLf)
    Response.Write("												<PARAM NAME=""_StockProps"" VALUE=""79"">" & vbCrLf)
    Response.Write("												<PARAM NAME=""Caption"" VALUE="""">" & vbCrLf)
    Response.Write("												<PARAM NAME=""ForeColor"" VALUE=""0"">" & vbCrLf)
    Response.Write("												<PARAM NAME=""BackColor"" VALUE=""16777215"">" & vbCrLf)
    Response.Write("												<PARAM NAME=""Enabled"" VALUE=""-1"">" & vbCrLf)
    Response.Write("												<PARAM NAME=""DataMember"" VALUE="""">" & vbCrLf)

    Response.Write("												<PARAM NAME=""Row.Count"" VALUE=""0"">" & vbCrLf)
    Response.Write("											</OBJECT>" & vbCrLf)
	
%>											

	<INPUT type='hidden' id=txtShowRegion name=txtShowRegion value='<%=Request("txtShowRegion")%>'>
	<INPUT type='hidden' id=txtShowWorkingPattern name=txtShowWorkingPattern value='<%=Request("txtShowWorkingPattern")%>'>
	<INPUT type='hidden' id=txtLoading name=txtLoading value=0>
</FORM>

        <form id="frmUseful" name="frmUseful" style="visibility: hidden; display: none">
            <input type="hidden" id="txtMenuSaved" name="txtMenuSaved" value="0">
        </form>


    </div>
</body>
</html>

<script type="text/javascript">
    calendar_report_breakdown_window_onload();
</script>
